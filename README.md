# GitHub Repo Manager

A small command-line tool to manage your GitHub repositories from an Excel
spreadsheet. Export every repo to an `.xlsx`, edit it (change visibility, set
descriptions, mark repos for deletion), then apply the changes back to GitHub.

All GitHub actions run through the [`gh` CLI](https://cli.github.com/), so there
are no API tokens to manage in this tool — it uses whatever `gh auth login` set up.

## Features

- **Export** all your repos to a formatted Excel file with dropdowns and
  read-only/identity columns protected.
- **Change visibility** (public ↔ private) in bulk.
- **Set descriptions** — optionally scan READMEs and auto-suggest a description
  for repos that don't have one.
- **Delete** repos, with multiple safety guards.
- **Preview by default** — nothing changes on GitHub until you pass `--yes`.
- Drift detection, archived-repo handling, and a per-run log.

## Prerequisites

| Requirement | Why it's needed | Install / check |
|---|---|---|
| **Python 3.8+** | runs the script | `python --version` |
| **`openpyxl`** | reads/writes the Excel file | `pip install -r requirements.txt` |
| **GitHub CLI (`gh`)** | performs every GitHub operation | [cli.github.com](https://cli.github.com/), then `gh auth login` |
| **`repo` token scope** | list repos, change visibility, read READMEs, set descriptions | included by default with `gh auth login` |
| **`delete_repo` token scope** | **required for the `delete` action only** | not default — see below |

### One-time setup

```
gh auth login                      # sign in (gives the 'repo' scope by default)
pip install -r requirements.txt    # install openpyxl
```

### ⚠️ Deletion requires the `delete_repo` scope

`gh auth login` grants the `repo` scope, which is enough to **export**, **change
visibility**, and **set descriptions**. Deleting a repo needs the **extra**
`delete_repo` scope — without it, `apply --yes` fails on `delete` rows (other
changes still succeed).

Check your scopes, then add it if missing:
```
gh auth status                     # look at the 'Token scopes:' line
gh auth refresh -s delete_repo     # add the delete scope
```

> The README scan and description updates need **nothing extra** beyond the
> `repo` scope above — no additional scope and no additional Python package.

## Quick start

```
python repo_manager.py export repos.xlsx        # 1. create the spreadsheet
#                                                 2. edit it in Excel
python repo_manager.py apply repos.xlsx         # 3. preview the changes
python repo_manager.py apply repos.xlsx --yes   # 4. apply them
```

## Usage

### Export

```
python repo_manager.py export [repos.xlsx] [--owner NAME] [--descriptions]
```

Creates the spreadsheet with one row per repo. The grey **identity columns**
(Owner, Repo, URL, Current Visibility, Has README, Current Description, Archived,
Fork, Created, Last Push, Size) are read-only. You edit four columns:

| Column             | Choices              | Meaning                                             |
|--------------------|----------------------|-----------------------------------------------------|
| `New Visibility`   | `public` / `private` | Pre-filled with the current value; change it to flip visibility. |
| `Action`           | `keep` / `delete`    | Set to `delete` to remove the repo.                 |
| `New Description`  | free text            | The text applied when `Set Description?` is `set`.  |
| `Set Description?` | `keep` / `set`       | Set to `set` to update the repo's About field.      |

Options:
- `--owner NAME` — export a different account/org instead of your own.
- `--descriptions` — scan each repo's README to fill **Has README** and pre-fill
  **New Description** with a short summary for repos that have a README but no
  description yet. Makes one request per repo (slower); without it those columns
  show `not scanned`.

### Filling in missing descriptions

1. `python repo_manager.py export repos.xlsx --descriptions`
2. In Excel, find rows where **Current Description** is empty but **Has README**
   is `TRUE` — **New Description** already holds a suggested summary. Adjust the
   text if you like, then set **Set Description?** to `set`.
3. Preview, then apply (see below).

### Apply

```
python repo_manager.py apply [repos.xlsx] [--yes] [--allow-mass-delete] [--force]
```

Reads your edits, re-checks GitHub, and prints a summary of what will change.
**Without `--yes` it only previews** (a dry run). Add `--yes` to execute.

## Safety

- `apply` is a **dry run unless you pass `--yes`** — you always see a summary first.
- `apply` re-checks GitHub before acting and **warns if the spreadsheet is stale**
  or a repo no longer exists (those rows are skipped).
- **Deletes are permanent.** Even with `--yes` you must type `DELETE` to confirm
  (skip with `--force`). More than 3 deletions in one run requires
  `--allow-mass-delete`.
- Setting a description is non-destructive and reversible, so it needs no extra
  confirmation beyond `--yes`.
- Visibility/description changes on archived repos are skipped with a notice.
- Every `apply --yes` writes a timestamped `repo_manager_<date>.log`.

## Troubleshooting

| Symptom | Fix |
|---|---|
| `gh ... not found on PATH` | Install the GitHub CLI and reopen your terminal. |
| `not authenticated with GitHub` | Run `gh auth login`. |
| Delete fails with a 403 / scope error | Add the scope: `gh auth refresh -s delete_repo`. |
| `could not write repos.xlsx` | The file is open in Excel — close it and retry. |
| Accents look wrong **in the terminal** | Display-only; the `.xlsx` and GitHub get correct UTF-8. On Windows use Windows Terminal or run `chcp 65001`. |
| Visibility change refused | The repo is archived (unarchive it first) or blocked by an org policy. |
| A suggested description is poor | It's a heuristic (first real paragraph of the README) — just edit the cell before applying. |

## Limitations & notes

- Manages repos for **one owner per file** (default: you). Org repos need admin
  rights to change or delete.
- Visibility choices are **public/private** only (`internal`/org visibility isn't offered).
- Lists up to **4000 repos** per owner.
- README summaries are a **heuristic**, not AI — review them before applying.
- Deleting a repo on GitHub is **irreversible**; use the delete feature carefully.

## Command reference

```
export [file] [--owner NAME] [--descriptions]
apply  [file] [--yes] [--allow-mass-delete] [--force]
```
`file` defaults to `repos.xlsx`.
