#!/usr/bin/env python3
"""Manage GitHub repositories from an Excel file.

Two commands:
  export  -  pull your repos into an editable .xlsx (dropdowns for New Visibility / Action)
  apply   -  read the edited .xlsx and apply changes (preview by default; --yes to execute)

All GitHub operations go through the `gh` CLI, so authentication is whatever
`gh auth login` set up. No GitHub token handling lives in this script.
"""

import argparse
import datetime
import json
import os
import subprocess
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# --- configuration ----------------------------------------------------------

JSON_FIELDS = "name,nameWithOwner,visibility,description,isArchived,isFork,pushedAt,diskUsage,url"
MASS_DELETE_THRESHOLD = 3

IDENTITY_HEADERS = [
    "Owner", "Repo", "URL", "Current Visibility",
    "Archived", "Fork", "Last Push", "Size (KB)",
]
EDIT_HEADERS = ["New Visibility", "Action"]
HEADERS = IDENTITY_HEADERS + EDIT_HEADERS

# column indexes (1-based) for the two editable columns
COL_NEW_VIS = IDENTITY_HEADERS.index("Current Visibility") + 1  # not editable; kept for color
COL_NEW_VISIBILITY = len(IDENTITY_HEADERS) + 1  # "New Visibility"
COL_ACTION = len(IDENTITY_HEADERS) + 2          # "Action"

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


def run_gh(args, check=True):
    """Run `gh <args>` and return the CompletedProcess."""
    cmd = ["gh"] + args
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
    except FileNotFoundError:
        sys.exit("ERROR: GitHub CLI 'gh' not found on PATH.\n"
                 "Install it from https://cli.github.com/ and run 'gh auth login'.")
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
    data = json.loads(r.stdout or "[]")
    for repo in data:
        repo["visibility"] = (repo.get("visibility") or "").lower()
    return data


# --- export -----------------------------------------------------------------

HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="305496")
LOCKED_FILL = PatternFill("solid", fgColor="F2F2F2")
PUBLIC_FILL = PatternFill("solid", fgColor="E2EFDA")
PRIVATE_FILL = PatternFill("solid", fgColor="FCE4D6")
COL_WIDTHS = [16, 28, 52, 16, 9, 7, 12, 9, 15, 9]


def cmd_export(args):
    ensure_auth()
    owner = args.owner or get_login()
    print(f"Fetching repos for '{owner}' via gh ...")
    repos = list_repos(owner)
    repos.sort(key=lambda r: r["nameWithOwner"].lower())
    print(f"  {len(repos)} repos found.")

    wb = Workbook()
    ws = wb.active
    ws.title = DATA_SHEET

    for col, head in enumerate(HEADERS, start=1):
        c = ws.cell(row=1, column=col, value=head)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(vertical="center")

    for i, repo in enumerate(repos):
        row = i + 2
        vis = repo["visibility"]
        vis_fill = PUBLIC_FILL if vis == "public" else PRIVATE_FILL
        identity = [
            repo["nameWithOwner"].split("/")[0],
            repo["name"],
            repo["url"],
            vis,
            "TRUE" if repo.get("isArchived") else "FALSE",
            "TRUE" if repo.get("isFork") else "FALSE",
            (repo.get("pushedAt") or "")[:10],
            repo.get("diskUsage") or 0,
        ]
        for col, val in enumerate(identity, start=1):
            c = ws.cell(row=row, column=col, value=val)
            c.fill = LOCKED_FILL
            c.protection = Protection(locked=True)
        ws.cell(row=row, column=COL_NEW_VIS).fill = vis_fill  # color the current-visibility cell

        nv = ws.cell(row=row, column=COL_NEW_VISIBILITY, value=vis)
        nv.protection = Protection(locked=False)
        nv.fill = vis_fill
        act = ws.cell(row=row, column=COL_ACTION, value="keep")
        act.protection = Protection(locked=False)

    last_row = max(len(repos) + 1, 2)

    dv_vis = DataValidation(type="list", formula1='"public,private"', allow_blank=False)
    dv_vis.error = "Choose 'public' or 'private'."
    dv_vis.prompt = "Target visibility for this repo."
    ws.add_data_validation(dv_vis)
    dv_vis.add(f"{get_column_letter(COL_NEW_VISIBILITY)}2:{get_column_letter(COL_NEW_VISIBILITY)}{last_row}")

    dv_act = DataValidation(type="list", formula1='"keep,delete"', allow_blank=False)
    dv_act.error = "Choose 'keep' or 'delete'."
    dv_act.prompt = "Set to 'delete' to remove the repo on apply."
    ws.add_data_validation(dv_act)
    dv_act.add(f"{get_column_letter(COL_ACTION)}2:{get_column_letter(COL_ACTION)}{last_row}")

    for col, width in enumerate(COL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}{last_row}"

    # Protect identity columns (locked cells) but keep the sheet pleasant to use.
    ws.protection.sheet = True
    for allow in ("autoFilter", "sort", "formatCells", "formatColumns",
                  "formatRows", "selectLockedCells", "selectUnlockedCells"):
        setattr(ws.protection, allow, False)

    snap = wb.create_sheet(SNAPSHOT_SHEET)
    snap.append(["nameWithOwner", "visibility"])
    for repo in repos:
        snap.append([repo["nameWithOwner"], repo["visibility"]])
    snap.sheet_state = "hidden"

    wb.save(args.file)
    print(f"Wrote {args.file}")
    print("Next: edit the 'New Visibility' / 'Action' columns, then run:")
    print(f"  python repo_manager.py apply {args.file}            # preview")
    print(f"  python repo_manager.py apply {args.file} --yes      # execute")


# --- apply ------------------------------------------------------------------

def _cell(value):
    return value if value is not None else ""


def read_rows(path):
    if not os.path.exists(path):
        sys.exit(f"ERROR: file not found: {path}\nRun 'export' first.")
    wb = load_workbook(path)
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
        })

    snapshot = {}
    if SNAPSHOT_SHEET in wb.sheetnames:
        for s in wb[SNAPSHOT_SHEET].iter_rows(min_row=2, values_only=True):
            if s and s[0]:
                snapshot[s[0]] = str(_cell(s[1])).strip().lower()
    return rows, snapshot


def compute_changes(rows):
    vis_changes, deletes, skipped_archived = [], [], []
    for row in rows:
        if row["action"] == "delete":
            deletes.append(row)
            continue
        nv = row["new_visibility"]
        if nv and nv != row["current_visibility"]:
            if nv not in ("public", "private"):
                continue
            if row["archived"]:
                skipped_archived.append(row)
                continue
            vis_changes.append(row)
    return vis_changes, deletes, skipped_archived


def fetch_live_map(owners):
    live = {}
    for owner in owners:
        for repo in list_repos(owner):
            live[repo["nameWithOwner"]] = repo["visibility"]
    return live


def print_summary(vis_changes, deletes, skipped_archived, warnings):
    print("\n=== Planned changes ===")
    to_priv = [r for r in vis_changes if r["new_visibility"] == "private"]
    to_pub = [r for r in vis_changes if r["new_visibility"] == "public"]
    print(f"Visibility: {len(to_priv)} public->private, {len(to_pub)} private->public")
    for r in vis_changes:
        print(f"  ~ {r['nameWithOwner']}: {r['current_visibility']} -> {r['new_visibility']}")
    print(f"Delete: {len(deletes)}")
    for r in deletes:
        print(f"  - {r['nameWithOwner']}")
    if skipped_archived:
        print(f"Skipped (archived, visibility change ignored): {len(skipped_archived)}")
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
    vis_changes, deletes, skipped_archived = compute_changes(rows)

    owners = sorted({r["owner"] for r in rows if r["owner"]})
    print(f"Checking live state on GitHub for: {', '.join(owners) or '(none)'} ...")
    live = fetch_live_map(owners)

    warnings = []
    for r in vis_changes + deletes:
        nwo = r["nameWithOwner"]
        if nwo not in live:
            warnings.append(f"{nwo}: not found on GitHub (already deleted/renamed) - will skip")
            r["_missing"] = True
        elif r["current_visibility"] and live[nwo] != r["current_visibility"]:
            warnings.append(f"{nwo}: file says '{r['current_visibility']}' but GitHub says "
                            f"'{live[nwo]}' (spreadsheet may be stale)")
    vis_changes = [r for r in vis_changes if not r.get("_missing")]
    deletes = [r for r in deletes if not r.get("_missing")]

    print_summary(vis_changes, deletes, skipped_archived, warnings)

    if not vis_changes and not deletes:
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
        answer = input("Type DELETE to confirm (anything else skips deletions): ").strip()
        if answer != "DELETE":
            print("Deletions cancelled. Visibility changes (if any) will still proceed.")
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

    write_log(log_lines, ok, fail)
    print(f"\nDone: {ok} succeeded, {fail} failed.")
    print(f"Tip: re-run 'python repo_manager.py export {args.file}' to refresh the spreadsheet.")


# --- cli --------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Manage GitHub repos from an Excel file (export / apply).")
    sub = parser.add_subparsers(dest="command", required=True)

    pe = sub.add_parser("export", help="Export your repos to an editable Excel file.")
    pe.add_argument("file", nargs="?", default="repos.xlsx", help="output .xlsx (default: repos.xlsx)")
    pe.add_argument("--owner", help="GitHub owner (default: your authenticated login).")
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
