# GitHub Repo Manager

Manage your GitHub repositories from an Excel file. **Export** every repo to an
`.xlsx`, edit it (visibility, descriptions, committed `.env` files, deletions),
then **apply** the changes back to GitHub.

All GitHub actions run through the [`gh` CLI](https://cli.github.com/) — no API
tokens are handled by this script; it uses whatever `gh auth login` set up.

## What it does

- **Export** all repos to a formatted sheet (newest first), with dropdowns and
  shaded reference (read-only) vs. editable columns.
- **Change visibility** (public ↔ private) in bulk.
- **Set descriptions** — optionally scan READMEs to suggest a description for
  repos that lack one (built-in heuristic, or an LLM via `--ai`).
- **Audit/manage committed `.env` files** — flag repos with a root `.env`, export
  its contents, and update or delete it.
- **Delete** repos, with safety guards.
- **Preview by default** — nothing changes on GitHub until you pass `--yes`.

## Prerequisites

| Requirement | Why | Install / check |
|---|---|---|
| **Python 3.8+** | runs the script | `python --version` |
| **`openpyxl`** | reads/writes the Excel file | `pip install -r requirements.txt` |
| **GitHub CLI (`gh`)** | every GitHub operation | [cli.github.com](https://cli.github.com/), then `gh auth login` |
| **`repo` scope** | list/visibility/READMEs/descriptions/`.env` | default with `gh auth login` |
| **`delete_repo` scope** | the `delete` action **only** | `gh auth refresh -s delete_repo` |
| **`anthropic` / `openai`** *(optional)* | `--ai` descriptions | in `requirements.txt` |
| **`ANTHROPIC_API_KEY` / `OPENAI_API_KEY`** *(optional)* | `--ai` only | env var or `.env` file |

```
gh auth login                      # sign in (grants the 'repo' scope)
pip install -r requirements.txt    # openpyxl + optional AI packages
```

Deleting a repo needs the extra `delete_repo` scope (`gh auth refresh -s
delete_repo`). Check with `gh auth status`.

## Workflow

```
python repo_manager.py export repos.xlsx        # 1. create the spreadsheet
#                                                 2. edit it in Excel
python repo_manager.py apply repos.xlsx         # 3. preview the changes
python repo_manager.py apply repos.xlsx --yes   # 4. apply them
```

## Editable columns

The grey columns are read-only references. Edit only these (all have dropdowns
except the free-text ones):

| Column | Choices | Meaning |
|---|---|---|
| `New Visibility` | `public` / `private` | Change to flip visibility. |
| `Action` | `keep` / `delete` | `delete` removes the repo. |
| `New Description` | free text | Applied when `Set Description?` is `set`. |
| `Set Description?` | `keep` / `set` | `set` updates the About field. |
| `New .env Content` | free text | Committed when `.env Action` is `update`. |
| `.env Action` | `keep` / `update` / `delete` | `update` commits New .env Content; `delete` removes the file. |

`New Description` / `New .env Content` are pre-filled (suggested summary, current
`.env`) when the relevant scan flag is used.

## AI descriptions (`--ai`)

`--ai` writes descriptions with an LLM instead of the heuristic. Choose the
provider with `--ai-provider` (default `claude`):

| Provider | Package | Key | Default model | Override env var |
|---|---|---|---|---|
| `claude` | `anthropic` | `ANTHROPIC_API_KEY` | `claude-haiku-4-5` | `REPO_MANAGER_CLAUDE_MODEL` |
| `openai` | `openai` | `OPENAI_API_KEY` | `gpt-4o-mini` | `REPO_MANAGER_OPENAI_MODEL` |

Keys are read from the environment or a local (gitignored) `.env`:

```
ANTHROPIC_API_KEY=sk-ant-...
OPENAI_API_KEY=sk-...
```

It calls the LLM only for repos that have a README but no description, and skips
READMEs with no real prose — so it stays cheap. The model replies `NONE` when a
README isn't descriptive enough (suggestion left blank), and on any API error it
falls back to the heuristic.

## Command reference

```
export [file] [--owner NAME] [--descriptions] [--ai] [--ai-provider claude|openai] [--check-env]
apply  [file] [--yes] [--allow-mass-delete] [--force]
```
`file` defaults to `repos.xlsx`.

### Export examples

```
python repo_manager.py export                                   # -> repos.xlsx (fast, no scans)
python repo_manager.py export mine.xlsx                         # custom output name
python repo_manager.py export repos.xlsx --owner some-org       # a different account/org
python repo_manager.py export repos.xlsx --descriptions         # heuristic README summaries
python repo_manager.py export repos.xlsx --ai                   # LLM summaries (Claude, default)
python repo_manager.py export repos.xlsx --ai --ai-provider claude
python repo_manager.py export repos.xlsx --ai --ai-provider openai
python repo_manager.py export repos.xlsx --check-env            # export committed root .env files
python repo_manager.py export repos.xlsx --ai --check-env       # combine scans
```

### Apply examples

```
python repo_manager.py apply                                    # preview repos.xlsx (dry run)
python repo_manager.py apply repos.xlsx                         # preview (dry run)
python repo_manager.py apply repos.xlsx --yes                   # execute changes
python repo_manager.py apply repos.xlsx --yes --allow-mass-delete  # allow >3 deletions
python repo_manager.py apply repos.xlsx --yes --force           # skip the typed DELETE prompt
```

### Flags

- `--owner NAME` — export another account/org (default: you).
- `--descriptions` — scan READMEs (one request per repo) to fill **Has README**
  and suggest descriptions (heuristic) for repos without one.
- `--ai` `[--ai-provider claude|openai]` — like `--descriptions` but LLM-written
  (higher quality; a few cents for a full account). Needs the provider package + key.
- `--check-env` — detect a committed root `.env` per repo and copy its contents
  into the sheet. ⚠️ contents are often secrets — see below.
- `--yes` — actually perform changes (otherwise apply is a dry run).
- `--allow-mass-delete` — required for more than 3 deletions in one run.
- `--force` — skip the interactive `DELETE` confirmation.

## Safety

- `apply` is a **dry run unless `--yes`** — you always see a summary first.
- It re-checks GitHub and warns if the sheet is stale or a repo is gone (skipped).
- Unrecognized values in editable columns are warned about and treated as no-ops.
- **Deletes are permanent**: even with `--yes` you must type `DELETE` (skip with
  `--force`); >3 deletions need `--allow-mass-delete`.
- Visibility/description/`.env` changes on archived repos are skipped with a notice.
- Each `apply --yes` writes a timestamped `repo_manager_<date>.log` (repo names
  and statuses only — never `.env` contents).
- ⚠️ **`.env` contents are secrets.** `--check-env` copies them into `repos.xlsx`
  (gitignored — keep it private, delete when done). `.env Action = update` commits
  them into git **history**, and `delete` does **not** purge history. Rotate any
  exposed credentials.

## Troubleshooting

| Symptom | Fix |
|---|---|
| `gh ... not found on PATH` | Install the GitHub CLI and reopen your terminal. |
| `not authenticated with GitHub` | Run `gh auth login`. |
| Delete fails with a 403 / scope error | `gh auth refresh -s delete_repo`. |
| `could not write repos.xlsx` | It's open in Excel — close it and retry. |
| Accents look wrong in the terminal | Display-only; the `.xlsx` and GitHub get correct UTF-8. |
| Visibility change refused | Repo is archived, or blocked by org policy. |

## Limitations

- One owner per file (default: you). Org repos need admin rights.
- Visibility is **public/private** only (no `internal`). Lists up to 4000 repos.
- `.env` features target the **root `.env`** only and commit to the **default branch**.
- Deleting a repo on GitHub is **irreversible**.
