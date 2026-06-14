# GitHub Repo Manager

Manage your GitHub repositories from an Excel spreadsheet: export them, edit
visibility / mark deletions in Excel, then apply the changes back to GitHub.

All GitHub actions run through the [`gh` CLI](https://cli.github.com/), so there
are no tokens to manage in this tool — it uses whatever `gh auth login` set up.

## Prerequisites

| Requirement | Why it's needed | Install / check |
|---|---|---|
| **Python 3.8+** | runs the script | `python --version` |
| **`openpyxl`** | reads/writes the Excel file | `pip install -r requirements.txt` |
| **GitHub CLI (`gh`)** | performs every GitHub operation | [cli.github.com](https://cli.github.com/), then `gh auth login` |
| **`repo` token scope** | list repos **and change visibility** | included by default with `gh auth login` |
| **`delete_repo` token scope** | **required for the `delete` action only** | not default — see below |

### One-time setup

```
gh auth login                      # sign in (gives the 'repo' scope by default)
pip install -r requirements.txt    # install openpyxl
```

### ⚠️ Deletion requires the `delete_repo` scope

`gh auth login` grants the `repo` scope, which is enough to **export** and to
**change visibility**. Deleting a repo needs the **extra** `delete_repo` scope —
without it, `apply --yes` will fail on any `delete` rows (visibility changes still
succeed).

Check whether you have it:
```
gh auth status
```
Look at the **`Token scopes:`** line for `delete_repo`. If it's missing, add it:
```
gh auth refresh -s delete_repo
```

## Workflow — two commands

### 1. Export
```
python repo_manager.py export repos.xlsx
```
Creates `repos.xlsx` with one row per repo. The grey **identity columns**
(Owner, Repo, URL, Current Visibility, Archived, Fork, Last Push, Size) are
read-only. Edit these two columns (they have dropdowns):

| Column          | Choices          | Meaning                                  |
|-----------------|------------------|------------------------------------------|
| `New Visibility`| `public` / `private` | Pre-filled with the current value. Change it to flip visibility. |
| `Action`        | `keep` / `delete`    | Set to `delete` to remove the repo.      |

Options: `--owner NAME` to export someone/somewhere other than your own account.

### 2. Apply
Preview first (default — makes **no** changes):
```
python repo_manager.py apply repos.xlsx
```
Then execute:
```
python repo_manager.py apply repos.xlsx --yes
```

## Safety

- `apply` is a **dry run unless you pass `--yes`** — you always see a summary first.
- Before export-vs-live drift could bite you, `apply` re-checks GitHub and warns
  if the spreadsheet is stale or a repo no longer exists (those rows are skipped).
- **Deletes are permanent.** Even with `--yes` you must type `DELETE` to confirm
  (skip with `--force`). More than 3 deletions in one run requires
  `--allow-mass-delete`.
- Visibility changes on archived repos are skipped with a notice.
- Every `apply --yes` writes a timestamped `repo_manager_<date>.log`.

## Flags reference

```
export [file] [--owner NAME]
apply  [file] [--yes] [--allow-mass-delete] [--force]
```
`file` defaults to `repos.xlsx`.
