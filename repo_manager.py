#!/usr/bin/env python3
"""Manage GitHub repositories from an Excel file.

Two commands:
  export  -  pull your repos into an editable .xlsx (dropdowns for New Visibility /
             Action / Set Description?). Add --descriptions to scan READMEs and
             suggest a description for repos that don't have one.
  apply   -  read the edited .xlsx and apply changes (preview by default; --yes to execute)

All GitHub operations go through the `gh` CLI, so authentication is whatever
`gh auth login` set up. No GitHub token handling lives in this script.
"""

import argparse
import concurrent.futures
import datetime
import json
import os
import re
import subprocess
import sys
import zipfile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Protection
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.worksheet.datavalidation import DataValidation

# --- configuration ----------------------------------------------------------

JSON_FIELDS = ("name,nameWithOwner,visibility,description,isArchived,isFork,"
               "createdAt,pushedAt,diskUsage,url")
MASS_DELETE_THRESHOLD = 3
DESCRIPTION_MAX = 250  # GitHub allows 350; keep suggestions concise

HEADERS = [
    "Owner", "Repo", "URL", "Current Visibility", "New Visibility", "Action",
    "Has README", "Current Description", "New Description", "Set Description?",
    "Archived", "Fork", "Created", "Last Push", "Size (KB)",
]
# Cells are read-only (locked) unless their header is in EDITABLE.
EDITABLE = {"New Visibility", "Action", "New Description", "Set Description?"}
COL = {h: i + 1 for i, h in enumerate(HEADERS)}  # 1-based column index by header

COL_WIDTHS = {
    "Owner": 16, "Repo": 28, "URL": 50, "Current Visibility": 16,
    "New Visibility": 15, "Action": 9, "Has README": 12,
    "Current Description": 40, "New Description": 45, "Set Description?": 15,
    "Archived": 9, "Fork": 7, "Created": 12, "Last Push": 12, "Size (KB)": 9,
}

SNAPSHOT_SHEET = "_snapshot"
DATA_SHEET = "repos"


# --- gh helpers -------------------------------------------------------------

class GhError(Exception):
    def __init__(self, cmd, code, stdout, stderr):
        self.cmd = cmd
        self.code = code
        self.stdout = stdout
        self.stderr = stderr
        super().__init__(f"gh exited {code}: {' '.join(cmd)}\n{(stderr or '').strip()}")


def run_gh(args, check=True, timeout=120):
    """Run `gh <args>` and return the CompletedProcess."""
    cmd = ["gh"] + args
    try:
        result = subprocess.run(cmd, capture_output=True, text=True,
                                encoding="utf-8", errors="replace", timeout=timeout)
    except FileNotFoundError:
        sys.exit("ERROR: GitHub CLI 'gh' not found on PATH.\n"
                 "Install it from https://cli.github.com/ and run 'gh auth login'.")
    except subprocess.TimeoutExpired:
        raise GhError(cmd, -1, "", f"command timed out after {timeout}s")
    if check and result.returncode != 0:
        raise GhError(cmd, result.returncode, result.stdout, result.stderr)
    return result


def ensure_auth():
    r = run_gh(["auth", "status"], check=False)
    if r.returncode != 0:
        sys.exit("ERROR: not authenticated with GitHub.\nRun 'gh auth login' first.\n" + (r.stderr or ""))


def get_login():
    return run_gh(["api", "user", "-q", ".login"]).stdout.strip()


def list_repos(owner):
    """Return a list of repo dicts for `owner`, visibility normalised to lowercase."""
    r = run_gh(["repo", "list", owner, "--limit", "4000", "--json", JSON_FIELDS])
    try:
        data = json.loads(r.stdout or "[]")
    except json.JSONDecodeError:
        sys.exit("ERROR: could not parse the repo list returned by gh.\n"
                 "Try 'gh repo list --limit 5' to check your gh setup.")
    for repo in data:
        repo["visibility"] = (repo.get("visibility") or "").lower()
    return data


# --- README scan + heuristic summary ----------------------------------------

_BADGE_RE = re.compile(r"^\s*\[?!\[")            # image / linked image (badges) at line start
_IMG_RE = re.compile(r"!\[[^\]]*\]\([^)]*\)")    # ![alt](url)
_LINK_RE = re.compile(r"\[([^\]]+)\]\([^)]*\)")  # [text](url) -> text
_HTML_TAG_RE = re.compile(r"<[^>]+>")
_EMPHASIS_RE = re.compile(r"[*_`~]+")


def fetch_readme(name_with_owner):
    """Return the README text for a repo, or None if it has no README.

    The raw bytes are decoded with a UTF-8 -> Windows-1252 -> Latin-1 fallback so
    READMEs written in legacy European encodings keep their accents.
    """
    cmd = ["gh", "api", f"repos/{name_with_owner}/readme",
           "-H", "Accept: application/vnd.github.raw"]
    try:
        result = subprocess.run(cmd, capture_output=True, timeout=30)
    except (FileNotFoundError, OSError, subprocess.TimeoutExpired):
        return None
    if result.returncode != 0:
        return None
    for enc in ("utf-8", "cp1252"):
        try:
            return result.stdout.decode(enc)
        except UnicodeDecodeError:
            continue
    return result.stdout.decode("latin-1")


def summarize_readme(text, max_len=DESCRIPTION_MAX):
    """Heuristically pull a one-line description from README markdown."""
    if not text:
        return ""
    paragraphs, current, in_code = [], [], False
    for raw in text.splitlines():
        line = raw.strip()
        if line.startswith("```") or line.startswith("~~~"):
            in_code = not in_code
            continue
        if in_code:
            continue
        if not line:
            if current:
                paragraphs.append(" ".join(current))
                current = []
            continue
        if line.startswith("#") or line.startswith(">") or line.startswith("<!--"):
            if current:
                paragraphs.append(" ".join(current))
                current = []
            continue
        if _BADGE_RE.match(line):
            continue
        line = re.sub(r"^([-*+]|\d+[.)])\s+", "", line)  # strip leading list marker
        current.append(line)
    if current:
        paragraphs.append(" ".join(current))

    for para in paragraphs:
        p = _IMG_RE.sub("", para)
        p = _LINK_RE.sub(r"\1", p)
        p = _HTML_TAG_RE.sub("", p)
        p = _EMPHASIS_RE.sub("", p)
        p = re.sub(r"\s+", " ", p).strip()
        if len(p) >= 20:
            if len(p) > max_len:
                p = p[:max_len].rsplit(" ", 1)[0].rstrip(",.;:") + "..."
            return p
    return ""


def scan_readmes(repos, max_workers=8):
    """Return {nameWithOwner: (has_readme: bool, suggestion: str)} for all repos."""
    def work(repo):
        nwo = repo["nameWithOwner"]
        try:
            text = fetch_readme(nwo)
        except Exception:
            return nwo, (False, "")
        if text is None:
            return nwo, (False, "")
        return nwo, (True, summarize_readme(text))

    result = {}
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as ex:
        for nwo, val in ex.map(work, repos):
            result[nwo] = val
    return result


# --- export -----------------------------------------------------------------

HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="305496")
LOCKED_FILL = PatternFill("solid", fgColor="F2F2F2")
PUBLIC_FILL = PatternFill("solid", fgColor="E2EFDA")
PRIVATE_FILL = PatternFill("solid", fgColor="FCE4D6")


def _add_list_validation(ws, header, last_row, items, error, prompt):
    dv = DataValidation(type="list", formula1=f'"{items}"', allow_blank=False)
    dv.error = error
    dv.prompt = prompt
    ws.add_data_validation(dv)
    letter = get_column_letter(COL[header])
    dv.add(f"{letter}2:{letter}{last_row}")


def _write_cell(ws, row, col, value):
    """Write a cell, storing strings that begin with '=' as literal text, not a formula."""
    c = ws.cell(row=row, column=col, value=value)
    if isinstance(value, str) and value.startswith("="):
        c.data_type = "s"
    return c


def cmd_export(args):
    ensure_auth()
    owner = args.owner or get_login()
    print(f"Fetching repos for '{owner}' via gh ...")
    repos = list_repos(owner)
    repos.sort(key=lambda r: r["nameWithOwner"].lower())
    print(f"  {len(repos)} repos found.")

    readme_map = {}
    if args.descriptions:
        print("Scanning READMEs (one request per repo) ...")
        readme_map = scan_readmes(repos)
        have = sum(1 for v in readme_map.values() if v[0])
        print(f"  {have}/{len(repos)} repos have a README.")

    wb = Workbook()
    ws = wb.active
    ws.title = DATA_SHEET

    for head in HEADERS:
        c = ws.cell(row=1, column=COL[head], value=head)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(vertical="center")

    for i, repo in enumerate(repos):
        row = i + 2
        vis = repo["visibility"]
        vis_fill = PUBLIC_FILL if vis == "public" else PRIVATE_FILL
        nwo = repo["nameWithOwner"]
        current_desc = repo.get("description") or ""

        if args.descriptions:
            has_readme, suggestion = readme_map.get(nwo, (False, ""))
            has_readme_val = "TRUE" if has_readme else "FALSE"
        else:
            has_readme, suggestion = False, ""
            has_readme_val = "not scanned"

        if args.descriptions and has_readme and suggestion and not current_desc:
            new_desc = suggestion
        else:
            new_desc = current_desc

        values = {
            "Owner": nwo.split("/")[0],
            "Repo": repo["name"],
            "URL": repo["url"],
            "Current Visibility": vis,
            "New Visibility": vis,
            "Action": "keep",
            "Has README": has_readme_val,
            "Current Description": current_desc,
            "New Description": new_desc,
            "Set Description?": "keep",
            "Archived": "TRUE" if repo.get("isArchived") else "FALSE",
            "Fork": "TRUE" if repo.get("isFork") else "FALSE",
            "Created": (repo.get("createdAt") or "")[:10],
            "Last Push": (repo.get("pushedAt") or "")[:10],
            "Size (KB)": repo.get("diskUsage") or 0,
        }
        for head in HEADERS:
            c = _write_cell(ws, row, COL[head], values[head])
            if head in EDITABLE:
                c.protection = Protection(locked=False)
            else:
                c.protection = Protection(locked=True)
                c.fill = LOCKED_FILL
        ws.cell(row=row, column=COL["Current Visibility"]).fill = vis_fill
        ws.cell(row=row, column=COL["New Visibility"]).fill = vis_fill

    last_row = max(len(repos) + 1, 2)

    _add_list_validation(ws, "New Visibility", last_row, "public,private",
                         "Choose 'public' or 'private'.", "Target visibility for this repo.")
    _add_list_validation(ws, "Action", last_row, "keep,delete",
                         "Choose 'keep' or 'delete'.", "Set to 'delete' to remove the repo on apply.")
    _add_list_validation(ws, "Set Description?", last_row, "keep,set",
                         "Choose 'keep' or 'set'.", "Set to 'set' to update the description on apply.")

    for head in HEADERS:
        ws.column_dimensions[get_column_letter(COL[head])].width = COL_WIDTHS[head]
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}{last_row}"

    ws.protection.sheet = True
    for allow in ("autoFilter", "sort", "formatCells", "formatColumns",
                  "formatRows", "selectLockedCells", "selectUnlockedCells"):
        setattr(ws.protection, allow, False)

    snap = wb.create_sheet(SNAPSHOT_SHEET)
    snap.append(["nameWithOwner", "visibility", "description"])
    for s_row, repo in enumerate(repos, start=2):
        _write_cell(snap, s_row, 1, repo["nameWithOwner"])
        _write_cell(snap, s_row, 2, repo["visibility"])
        _write_cell(snap, s_row, 3, repo.get("description") or "")
    snap.sheet_state = "hidden"

    try:
        wb.save(args.file)
    except PermissionError:
        sys.exit(f"ERROR: could not write {args.file}. "
                 f"If it's open in Excel, close it and run export again.")
    print(f"Wrote {args.file}")
    print("Next: edit the editable columns (New Visibility / Action / New Description /"
          " Set Description?), then run:")
    print(f"  python repo_manager.py apply {args.file}            # preview")
    print(f"  python repo_manager.py apply {args.file} --yes      # execute")
    if not args.descriptions:
        print("Tip: add --descriptions to scan READMEs and suggest descriptions for repos missing one.")


# --- apply ------------------------------------------------------------------

def _cell(value):
    return value if value is not None else ""


def read_rows(path):
    if not os.path.exists(path):
        sys.exit(f"ERROR: file not found: {path}\nRun 'export' first.")
    try:
        wb = load_workbook(path)
    except (InvalidFileException, zipfile.BadZipFile, PermissionError) as e:
        sys.exit(f"ERROR: could not open {path} ({e}).\n"
                 f"Make sure it's an .xlsx created by 'export' and not open in Excel.")
    if DATA_SHEET not in wb.sheetnames:
        sys.exit(f"ERROR: sheet '{DATA_SHEET}' not found in {path}. Was this made by 'export'?")
    ws = wb[DATA_SHEET]

    headers = [(_cell(c.value)).strip() if isinstance(c.value, str) else _cell(c.value) for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    for req in ("Owner", "Repo", "Current Visibility", "New Visibility", "Action"):
        if req not in idx:
            sys.exit(f"ERROR: column '{req}' missing in {path}. Re-export the file.")

    def get(row, name):
        i = idx.get(name)
        return row[i] if i is not None and i < len(row) else None

    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        owner = get(r, "Owner")
        repo = get(r, "Repo")
        if not owner and not repo:
            continue
        rows.append({
            "owner": str(owner).strip(),
            "repo": str(repo).strip(),
            "nameWithOwner": f"{str(owner).strip()}/{str(repo).strip()}",
            "current_visibility": str(_cell(get(r, "Current Visibility"))).strip().lower(),
            "new_visibility": str(_cell(get(r, "New Visibility"))).strip().lower(),
            "action": (str(_cell(get(r, "Action"))).strip().lower() or "keep"),
            "archived": str(_cell(get(r, "Archived"))).strip().upper() == "TRUE",
            "current_description": str(_cell(get(r, "Current Description"))).strip(),
            "new_description": str(_cell(get(r, "New Description"))).strip(),
            "set_description": (str(_cell(get(r, "Set Description?"))).strip().lower() or "keep"),
        })

    snapshot = {}
    if SNAPSHOT_SHEET in wb.sheetnames:
        for s in wb[SNAPSHOT_SHEET].iter_rows(min_row=2, values_only=True):
            if s and s[0]:
                snapshot[s[0]] = {
                    "visibility": str(_cell(s[1])).strip().lower() if len(s) > 1 else "",
                    "description": str(_cell(s[2])).strip() if len(s) > 2 else "",
                }
    return rows, snapshot


def compute_changes(rows):
    vis_changes, deletes, skipped_archived, desc_changes = [], [], [], []
    for row in rows:
        if row["action"] == "delete":
            deletes.append(row)
            continue

        nv = row["new_visibility"]
        if nv and nv != row["current_visibility"] and nv in ("public", "private"):
            if row["archived"]:
                skipped_archived.append(row)
            else:
                vis_changes.append(row)

        if row["set_description"] == "set":
            nd = row["new_description"]
            if nd and nd != row["current_description"]:
                if row["archived"]:
                    if row not in skipped_archived:
                        skipped_archived.append(row)
                else:
                    desc_changes.append(row)
    return vis_changes, deletes, skipped_archived, desc_changes


def fetch_live_map(owners):
    live = {}
    for owner in owners:
        for repo in list_repos(owner):
            live[repo["nameWithOwner"]] = {
                "visibility": repo["visibility"],
                "description": repo.get("description") or "",
            }
    return live


def print_summary(vis_changes, deletes, skipped_archived, desc_changes, warnings):
    print("\n=== Planned changes ===")
    to_priv = [r for r in vis_changes if r["new_visibility"] == "private"]
    to_pub = [r for r in vis_changes if r["new_visibility"] == "public"]
    print(f"Visibility: {len(to_priv)} public->private, {len(to_pub)} private->public")
    for r in vis_changes:
        print(f"  ~ {r['nameWithOwner']}: {r['current_visibility']} -> {r['new_visibility']}")
    print(f"Descriptions to set: {len(desc_changes)}")
    for r in desc_changes:
        preview = r["new_description"]
        if len(preview) > 60:
            preview = preview[:60] + "..."
        print(f"  = {r['nameWithOwner']}: \"{preview}\"")
    print(f"Delete: {len(deletes)}")
    for r in deletes:
        print(f"  - {r['nameWithOwner']}")
    if skipped_archived:
        print(f"Skipped (archived, can't edit): {len(skipped_archived)}")
        for r in skipped_archived:
            print(f"  . {r['nameWithOwner']}")
    if warnings:
        print("Warnings:")
        for w in warnings:
            print(f"  ! {w}")


def write_log(lines, ok, fail):
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    path = f"repo_manager_{ts}.log"
    with open(path, "w", encoding="utf-8") as f:
        f.write(f"repo_manager apply run at {ts}\n\n")
        f.write("\n".join(lines))
        f.write(f"\n\nSummary: {ok} succeeded, {fail} failed.\n")
    print(f"Log written to {path}")


def cmd_apply(args):
    ensure_auth()
    rows, _snapshot = read_rows(args.file)
    vis_changes, deletes, skipped_archived, desc_changes = compute_changes(rows)

    owners = sorted({r["owner"] for r in rows if r["owner"]})
    print(f"Checking live state on GitHub for: {', '.join(owners) or '(none)'} ...")
    live = fetch_live_map(owners)

    warnings = []
    for r in rows:
        if r["set_description"] == "set" and not r["new_description"]:
            warnings.append(f"{r['nameWithOwner']}: 'Set Description?' is 'set' but "
                            f"New Description is empty - skipped")

    for r in vis_changes + deletes + desc_changes:
        nwo = r["nameWithOwner"]
        if nwo not in live:
            warnings.append(f"{nwo}: not found on GitHub (already deleted/renamed) - will skip")
            r["_missing"] = True
        elif r["current_visibility"] and live[nwo]["visibility"] != r["current_visibility"]:
            warnings.append(f"{nwo}: file says '{r['current_visibility']}' but GitHub says "
                            f"'{live[nwo]['visibility']}' (spreadsheet may be stale)")
    vis_changes = [r for r in vis_changes if not r.get("_missing")]
    deletes = [r for r in deletes if not r.get("_missing")]
    desc_changes = [r for r in desc_changes if not r.get("_missing")]

    print_summary(vis_changes, deletes, skipped_archived, desc_changes, warnings)

    if not vis_changes and not deletes and not desc_changes:
        print("\nNothing to do.")
        return

    if not args.yes:
        print("\nDRY RUN - no changes made. Re-run with --yes to apply.")
        return

    if len(deletes) > MASS_DELETE_THRESHOLD and not args.allow_mass_delete:
        sys.exit(f"\nERROR: {len(deletes)} deletions requested (> {MASS_DELETE_THRESHOLD}).\n"
                 f"Re-run with --allow-mass-delete if you really mean it.")

    if deletes and not args.force:
        print("\nThe following repos will be PERMANENTLY DELETED:")
        for r in deletes:
            print(f"  - {r['nameWithOwner']}")
        try:
            answer = input("Type DELETE to confirm (anything else skips deletions): ").strip()
        except EOFError:
            answer = ""
        if answer != "DELETE":
            print("Deletions cancelled. Other changes (if any) will still proceed.")
            deletes = []

    log_lines, ok, fail = [], 0, 0

    for r in vis_changes:
        nwo, target = r["nameWithOwner"], r["new_visibility"]
        try:
            run_gh(["repo", "edit", nwo, "--visibility", target,
                    "--accept-visibility-change-consequences"])
            line = f"OK    visibility {r['current_visibility']}->{target}  {nwo}"
            ok += 1
        except GhError as e:
            line = f"FAIL  visibility {nwo}: {(e.stderr or '').strip()}"
            fail += 1
        print(line)
        log_lines.append(line)

    for r in desc_changes:
        nwo = r["nameWithOwner"]
        try:
            run_gh(["repo", "edit", nwo, "--description", r["new_description"]])
            line = f"OK    description set  {nwo}"
            ok += 1
        except GhError as e:
            line = f"FAIL  description {nwo}: {(e.stderr or '').strip()}"
            fail += 1
        print(line)
        log_lines.append(line)

    for r in deletes:
        nwo = r["nameWithOwner"]
        try:
            run_gh(["repo", "delete", nwo, "--yes"])
            line = f"OK    deleted  {nwo}"
            ok += 1
        except GhError as e:
            line = f"FAIL  delete {nwo}: {(e.stderr or '').strip()}"
            fail += 1
        print(line)
        log_lines.append(line)

    if log_lines:
        write_log(log_lines, ok, fail)
    print(f"\nDone: {ok} succeeded, {fail} failed.")
    print(f"Tip: re-run 'python repo_manager.py export {args.file}' to refresh the spreadsheet.")


# --- cli --------------------------------------------------------------------

def main():
    # Never crash just because a description has an emoji/CJK char the console
    # can't encode; degrade to a replacement char instead.
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass

    parser = argparse.ArgumentParser(
        description="Manage GitHub repos from an Excel file (export / apply).")
    sub = parser.add_subparsers(dest="command", required=True)

    pe = sub.add_parser("export", help="Export your repos to an editable Excel file.")
    pe.add_argument("file", nargs="?", default="repos.xlsx", help="output .xlsx (default: repos.xlsx)")
    pe.add_argument("--owner", help="GitHub owner (default: your authenticated login).")
    pe.add_argument("--descriptions", action="store_true",
                    help="Scan READMEs to fill 'Has README' and suggest descriptions "
                         "for repos that don't have one.")
    pe.set_defaults(func=cmd_export)

    pa = sub.add_parser("apply", help="Apply changes from the Excel file (preview unless --yes).")
    pa.add_argument("file", nargs="?", default="repos.xlsx", help="input .xlsx (default: repos.xlsx)")
    pa.add_argument("--yes", action="store_true", help="Actually perform changes (default: preview).")
    pa.add_argument("--allow-mass-delete", action="store_true",
                    help=f"Permit more than {MASS_DELETE_THRESHOLD} deletions in one run.")
    pa.add_argument("--force", action="store_true", help="Skip the interactive DELETE confirmation.")
    pa.set_defaults(func=cmd_apply)

    args = parser.parse_args()
    try:
        args.func(args)
    except GhError as e:
        sys.exit(f"ERROR: {e}")
    except KeyboardInterrupt:
        sys.exit("\nAborted.")


if __name__ == "__main__":
    main()
