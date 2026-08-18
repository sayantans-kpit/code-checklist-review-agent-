# Code Review Checklist Agent 📋

A VS Code Copilot Chat extension that automatically fills the **PSSM2.0 Source Code Review Checklist** Excel template by analysing a GitHub PR's code diff and review comments using AI.

---

## What it does

When you run `@checklist /generate`, the agent:

1. **Fetches the PR** from GitHub — code diff, review comments, assignees, reviewers, dates
2. **Analyses the code diff** (not just comments) — finds issues the comments may have missed
3. **Cross-references** — if a comment flagged an issue and the diff shows it was fixed, marks it as addressed
4. **Fills your exact Excel template** — the same PSSM2.0 file you share with stakeholders, with correct formatting
5. **Versions every run** — `v1`, `v2`, `v3`… so re-reviews are tracked automatically
6. **Adds a Re-review sheet** when Not Ok rows exist, listing only what needs attention

---

## Quick Start

### Step 1 — Install the extension

```bash
code --install-extension code-review-checklist-agent-0.2.1.vsix
```

Reload VS Code (`Ctrl+Shift+P` → **Developer: Reload Window**).

### Step 2 — Save your GitHub PAT (once only)

Open Copilot Chat (`Ctrl+Alt+I`) and run:

```
@checklist /token ghp_xxxxxxxxxxxxxxxxxxxx
```

Your token is stored in VS Code's encrypted OS keychain — never in any file.

Required GitHub scopes: **`repo`**
Generate one at: https://github.com/settings/tokens

### Step 3 — Generate a checklist

```
@checklist /generate https://github.com/your-org/your-repo/pull/123
```

The agent fetches everything automatically and saves the filled Excel to:
```
<workspace>/code-review/YYYY-MM-DD/CodeReview_PR-123_v1.xlsx
```

---

## Usage

### Basic — GitHub API (recommended)

```
@checklist /generate https://github.com/org/repo/pull/42
```

Auto-fetches code diff, PR comments, assignees, reviewers, and dates.
Uses your saved PAT automatically.

---

### With a Prompt: specification

Add extra context to guide the AI — business rules, sprint decisions, focus areas:

```
@checklist /generate https://github.com/org/repo/pull/42
Prompt: Backend-only Rails PR for case search with related ROs.
        Performance is top priority — we handle 10k+ cases in prod.
        i18n is intentionally skipped this sprint — mark those rows NA.
        No React components added — skip all frontend rows.
```

The spec is passed as a **high-priority instruction** to the AI, so it adjusts every row accordingly.

---

### Local git diff (no PAT needed)

```
@checklist /generate --branch feature/my-branch --base main
```

Reads the diff directly from your local git repository. Paste review comments below (optional).

---

### Manual comments only

```
@checklist /generate https://github.com/org/repo/pull/42
Comment 1: Missing error handling in UserController#update
Comment 2: N+1 query in app/views/users/index line 34
```

---

## What gets auto-filled in the Excel

| Cell | Field | Source |
|------|-------|--------|
| `C4` | Source Code | PR URL |
| `C5` | Version | `v1` / `v2` / `v3`… |
| `C6` | Author(s) | PR assignees (falls back to PR opener) |
| `C7` | Reviewer(s) | All PR reviewers (requested + submitted) |
| `C8` | Start Date | Date PR was raised (`pr.created_at`) |
| `C9` | End Date | Merged/closed date — blank if PR still open |
| `E11:E35` | Author status | `Yes` / `No` / `NA` per checklist row |
| `F11:F35` | Reviewer status | `Ok` / `Not Ok` / `NA` per checklist row |
| `G11:G35` | Description of Finding | AI-generated from diff + comments |
| `H11:H35` | Defect Type | Functional / Technical / Process/Compliance / etc. |
| `I11:I35` | Remarks by Author | AI-suggested action — you can edit |
| `H3:H9` | Defect counts | Computed directly (no recalculation needed) |
| `J4:J9` | Summary counts | Yes / No / NA / Ok / Not Ok counts |

### Extra sheets added automatically

| Sheet | When added | Contents |
|-------|-----------|----------|
| **Re-review Required** | When any row is `Not Ok` | Only the Not Ok rows with findings and suggested actions |
| **Version History** | Always | All past versions of this PR's checklist with dates and counts |

---

## File structure

Every generated checklist is saved in a structured folder inside your workspace:

```
<workspace>/
  code-review/
    2026-08-18/
      CodeReview_PR-20715_v1.xlsx    ← first review
      CodeReview_PR-20715_v2.xlsx    ← re-review after fixes
    2026-08-19/
      CodeReview_PR-20716_v1.xlsx
    index.json                        ← fast lookup index (PR URL → versions)
```

---

## Versioning

Every time you run `@checklist /generate` for the same PR:
- If `v1` already exists → creates `v2` automatically
- Each version is registered in `code-review/index.json`
- View history: `@checklist /history https://github.com/org/repo/pull/42`

---

## Row colour coding

| Colour | Meaning |
|--------|---------|
| 🔴 Light red | `Not Ok` — reviewer found an issue |
| 🟢 Light green | `Ok` — reviewer confirmed this area is fine |
| 🟡 Light yellow | `NA` — not applicable to this PR |

---

## All commands

| Command | Description |
|---------|-------------|
| `@checklist /generate <PR_URL>` | Generate checklist from a GitHub PR |
| `@checklist /generate --branch <name> --base <base>` | Use local git diff |
| `@checklist /token <PAT>` | Save GitHub PAT to OS keychain |
| `@checklist /token status` | Check if a PAT is stored |
| `@checklist /token clear` | Remove stored PAT |
| `@checklist /history <PR_URL>` | Show all checklist versions for a PR |
| `@checklist /help` | Show usage instructions |

---

## Building from source

Requires Node.js 18+ (uses `.nvmrc` to pin to v22):

```bash
cd code-review-checklist-agent
./build-vsix.sh
```

This installs dependencies, compiles TypeScript, bundles with esbuild, and produces a `.vsix` file.

---

## Architecture

```
src/extension.ts
├── PRData interface          — PR metadata, assignees, reviewers, dates, diff, comments
├── IndexStore                — Reads/writes code-review/index.json for versioning
├── fetchFromGitHub()         — GitHub API: PR info + reviews + inline comments + files + reviewers
├── fetchLocalDiff()          — Git CLI: local branch diff
├── analyseWithAI()           — Sends diff + comments + user spec to Copilot LLM
├── generateExcel()           — Loads template, fills cells, adds Re-review + History sheets
└── activate()                — Registers @checklist chat participant + /generate /token /history /help
```

---

## Notes

- The PAT is stored in VS Code `SecretStorage` which maps to the OS keychain (Gnome Keyring / macOS Keychain / Windows Credential Manager)
- The Excel template (`templates/checklist-template.xlsx`) is bundled inside the VSIX — never modified in place
- All findings are AI-generated suggestions — always review before sharing with stakeholders
