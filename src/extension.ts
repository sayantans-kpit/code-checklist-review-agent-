import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import * as cp from 'child_process';
import * as util from 'util';
import ExcelJS from 'exceljs';

const execAsync = util.promisify(cp.exec);

// ─────────────────────────────────────────────────────────────────────────────
// File reader — extracts PR URLs and Jira keys from any attached file
// Supports: .xlsx, .csv, .txt, .md, .json
// ─────────────────────────────────────────────────────────────────────────────

interface FileReadResult {
  rawText:    string;    // full readable content for AI context
  prUrls:     string[];  // GitHub PR URLs found
  jiraKeys:   string[];  // Jira story keys found
  allItems:   string[];  // everything that looks actionable
}

async function readAttachedFile(filePath: string): Promise<FileReadResult> {
  const ext = path.extname(filePath).toLowerCase();
  const PR_URL  = /https?:\/\/github\.com\/[^\s"',]+\/pull\/\d+/g;
  const JIRA_K  = /\b([A-Z][A-Z0-9]+-\d+)\b/g;
  const JIRA_URL = /atlassian\.net\/browse\/([A-Z][A-Z0-9]+-\d+)/g;

  let rawText = '';

  if (ext === '.xlsx' || ext === '.xls') {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const lines: string[] = [];
    wb.eachSheet(ws => {
      ws.eachRow(row => {
        const cells = (row.values as any[]).slice(1).map(v => v?.toString?.() ?? '').filter(Boolean);
        if (cells.length) { lines.push(cells.join('\t')); }
      });
    });
    rawText = lines.join('\n');

  } else if (ext === '.csv') {
    rawText = fs.readFileSync(filePath, 'utf8');

  } else if (['.txt', '.md', '.json'].includes(ext)) {
    rawText = fs.readFileSync(filePath, 'utf8');
    if (ext === '.json') {
      try {
        const parsed = JSON.parse(rawText);
        rawText = JSON.stringify(parsed, null, 2);
      } catch { /* keep raw */ }
    }
  } else {
    // Try reading as text anyway
    try { rawText = fs.readFileSync(filePath, 'utf8'); } catch { rawText = ''; }
  }

  const prUrls   = [...new Set([...rawText.matchAll(PR_URL)].map(m => m[0]))];
  const jiraKeys = [...new Set([
    ...[...rawText.matchAll(JIRA_K)].map(m => m[1]),
    ...[...rawText.matchAll(JIRA_URL)].map(m => m[1]),
  ])];
  const allItems = [...prUrls, ...jiraKeys.filter(k => !prUrls.some(u => u.includes(k)))];

  return { rawText: rawText.slice(0, 8000), prUrls, jiraKeys, allItems };
}

// ─────────────────────────────────────────────────────────────────────────────
// Checklist rows — PSSM2.0 template (rows 10-34)
// ─────────────────────────────────────────────────────────────────────────────
interface ChecklistRow {
  id: number;
  reviewArea: string;
  checkCategory: string;
  description: string;
}

const CHECKLIST_ROWS: ChecklistRow[] = [
  { id: 1,  reviewArea: 'Branch / Dependencies',            checkCategory: 'Branch currency and dependency alignment',    description: 'Keep the branch rebased/merged with master/main and verify shared UI/component library versions are current before relying on new APIs.' },
  { id: 2,  reviewArea: 'I18n / Translations',              checkCategory: 'Internationalization and user-facing text',   description: 'Do not hardcode user-facing strings. Add missing translations and pass them through the established Rails/React I18n flow.' },
  { id: 3,  reviewArea: 'Feature Flags / Dependency Ordering', checkCategory: 'Feature readiness and safe rollout',       description: 'Do not reference new associations, attributes, APIs, or upcoming-PR functionality in live code until they exist on master/main, unless safely feature-flagged.' },
  { id: 4,  reviewArea: 'Performance / DB',                 checkCategory: 'Database query efficiency',                   description: 'Eager-load associations used inside loops, table rows, and partials to avoid N+1 queries.' },
  { id: 5,  reviewArea: 'Architecture / Services',          checkCategory: 'Controller/service responsibility separation', description: 'Avoid duplicate controller methods and extract reusable business logic into service classes, helpers, or named domain methods.' },
  { id: 6,  reviewArea: 'Error Handling',                   checkCategory: 'HTTP failure semantics and UI error behavior', description: 'Handle HTTP statuses intentionally so business validation responses are displayed correctly while true failures still trigger error paths.' },
  { id: 7,  reviewArea: 'UI State / Conditions',            checkCategory: 'Boolean logic correctness',                  description: 'Review boolean conditions for impossible states, incorrect &&/|| usage, duplicate-message guards, and unreachable UI behavior.' },
  { id: 8,  reviewArea: 'Testing / Service Specs',          checkCategory: 'Test realism and contract fidelity',         description: 'Ensure tests pass objects/doubles that match runtime contracts rather than relying on coincidental method compatibility.' },
  { id: 9,  reviewArea: 'PR Hygiene',                       checkCategory: 'PR readiness and reviewer workflow',         description: 'Before marking a PR ready for review, resolve automated review/Copilot comments, formatter issues, and obvious cleanup items.' },
  { id: 10, reviewArea: 'Imports',                          checkCategory: 'Import consistency and shared package usage', description: 'Use the organization-preferred import style for shared components and avoid deep/internal imports unless the codebase standard requires them.' },
  { id: 11, reviewArea: 'Component API',                    checkCategory: 'Use existing component capabilities',        description: 'Prefer built-in shared component options over custom workaround code when the component library already supports the behavior.' },
  { id: 12, reviewArea: 'Props / Defaults',                 checkCategory: 'Explicit component contracts',               description: 'Define propTypes/defaultProps consistently, mark required props explicitly, and avoid scattered JavaScript fallback values for component defaults.' },
  { id: 13, reviewArea: 'Styling',                          checkCategory: 'Styling separation and render performance',  description: 'Move inline styles, ad-hoc style objects, and declarative styling into styles.js, styled-components, or the approved CSS/styling file.' },
  { id: 14, reviewArea: 'Design Tokens',                    checkCategory: 'Design-system token usage',                  description: 'Use approved design tokens for spacing and convert fixed pixel-like values to rem where that is the codebase pattern.' },
  { id: 15, reviewArea: 'Dead Code / Cleanup',              checkCategory: 'Remove unused or redundant code',            description: 'Remove unused props, parameters, style variables, spreads, callbacks, and redundant conditional branches unless they are intentionally required.' },
  { id: 16, reviewArea: 'Validation / Business Rules',      checkCategory: 'Validation readability and testability',     description: 'Extract complex validation, eligibility, or business-rule checks into named validation methods/services.' },
  { id: 17, reviewArea: 'Logging / Observability',          checkCategory: 'Production-ready logging',                   description: 'Use the approved structured logging format and avoid committing temporary/local-debug logging changes.' },
  { id: 18, reviewArea: 'Testing / Jest',                   checkCategory: 'Frontend test conventions',                  description: 'Use shared test utilities/providers and follow the application convention for Jest descriptors such as describe + it.' },
  { id: 19, reviewArea: 'Views / Partials',                 checkCategory: 'Partial/component API contract',             description: 'Pass only parameters that are consumed by the target partial/component, or update the target to use the parameter intentionally.' },
  { id: 20, reviewArea: 'Formatting / Linting',             checkCategory: 'Automated formatting and style consistency', description: 'Run the configured formatter/linter and fix indentation, spacing, JSX formatting, import cleanup, trailing commas, and newline-at-EOF issues.' },
  { id: 21, reviewArea: 'React Handlers',                   checkCategory: 'Simplify React callbacks and handlers',      description: 'Avoid wrapper handlers or unnecessary useCallback when a prop callback can be passed directly.' },
  { id: 22, reviewArea: 'DOM IDs',                          checkCategory: 'DOM/test selector namespacing',              description: 'Namespace React/KDS DOM IDs using the approved prefix convention.' },
  { id: 23, reviewArea: 'Ruby / Safe Navigation',           checkCategory: 'Nil handling clarity',                       description: 'Use safe navigation only when nil is actually possible; remove redundant safe navigation after explicit presence checks.' },
  { id: 24, reviewArea: 'Boilerplate / Tooling',            checkCategory: 'Scaffolding and generated structure',        description: 'Use approved generators/scaffolding commands for repetitive component or test setup when available.' },
  { id: 25, reviewArea: 'Review Communication',             checkCategory: 'Reviewer tone and non-blocking guidance',    description: 'Classify comments as blocking or non-blocking, and phrase minor suggestions as optional when they do not affect correctness.' },
];

// ─────────────────────────────────────────────────────────────────────────────
// Data types
// ─────────────────────────────────────────────────────────────────────────────

interface FilePatch {
  filename: string;
  status: string;          // added | modified | removed
  additions: number;
  deletions: number;
  patch: string;           // the raw unified diff for this file (may be truncated)
}

interface PRData {
  url: string;
  title: string;
  body: string;            // PR description — used to extract Jira links
  author: string;
  assignees: string[];
  reviewers: string[];
  branch: string;
  baseBranch: string;
  createdAt: string | null;
  closedAt:  string | null;
  comments: string[];
  files: FilePatch[];
  diffSummary: string;
}

interface RowFinding {
  rowId: number;
  /** Col E — Author's self-assessment: "Yes" (done) | "No" (not done) | "NA" (not applicable) */
  authorStatus: 'Yes' | 'No' | 'NA';
  /** Col F — Reviewer verdict: "Ok" | "Not Ok" | "NA"  (must match COUNTIF formula values exactly) */
  reviewerStatus: 'Ok' | 'Not Ok' | 'NA';
  /** Col G — Description of Finding by reviewer */
  finding: string;
  /** Col H — Defect Type */
  defectType: 'Functional' | 'Technical' | 'Process/Compliance' | 'Documentation' | 'Others' | '';
  /** Col I — Remarks by Author (AI suggestion; human can override) */
  remarks: string;
}

// ─────────────────────────────────────────────────────────────────────────────
// GitHub API — fetch PR metadata, comments AND file diffs
// ─────────────────────────────────────────────────────────────────────────────

function parsePRUrl(url: string): { owner: string; repo: string; number: number } | null {
  const m = url.match(/github\.com\/([^/]+)\/([^/]+)\/pull\/(\d+)/);
  return m ? { owner: m[1], repo: m[2], number: parseInt(m[3], 10) } : null;
}

/** Cap a patch to ~120 lines so we don't blow the LLM context on huge diffs */
function truncatePatch(patch: string, maxLines = 120): string {
  const lines = patch.split('\n');
  if (lines.length <= maxLines) { return patch; }
  return lines.slice(0, maxLines).join('\n') + `\n… [truncated ${lines.length - maxLines} lines]`;
}

/** Format ISO date string → "dd-Mmm-yyyy" (template format). Returns null if input is null. */
function fmtDate(iso: string | null): string | null {
  if (!iso) { return null; }
  return new Date(iso).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
}

/** HTTPS GET using Node built-in module — more reliable in VS Code extension host.
 *  Gives actionable errors for proxy/VPN/cert issues. */
function httpsGet(url: string, headers: Record<string, string>): Promise<{ok: boolean; status: number; _linkHeader?: string; text: () => Promise<string>; json: () => Promise<any>}> {
  return new Promise((resolve, reject) => {
    const https  = require('https') as typeof import('https');
    const urlObj = new URL(url);
    const req    = https.request(
      {hostname: urlObj.hostname, path: urlObj.pathname + urlObj.search, method: 'GET', headers, timeout: 15000},
      (res) => {
        const chunks: Buffer[] = [];
        res.on('data', (c: Buffer) => chunks.push(c));
        res.on('end', () => {
          const body       = Buffer.concat(chunks).toString('utf8');
          const status     = res.statusCode ?? 0;
          const linkHeader = (res.headers['link'] as string | undefined);
          resolve({ok: status >= 200 && status < 300, status, _linkHeader: linkHeader, text: () => Promise.resolve(body), json: () => Promise.resolve(JSON.parse(body))});
        });
      }
    );
    req.on('timeout', () => { req.destroy(); reject(new Error('Request timed out (15s) — check VPN/network')); });
    req.on('error',  (err: NodeJS.ErrnoException) => {
      const hint =
        err.code === 'ECONNREFUSED' ? ' — connection refused (check VPN/proxy)' :
        err.code === 'ENOTFOUND'    ? ' — DNS failed (check internet connection)' :
        (err.code ?? '').startsWith('CERT') ? ' — SSL cert error (corporate proxy may intercept HTTPS)' : '';
      reject(new Error(`Network error: ${err.message}${hint}`));
    });
    req.end();
  });
}

/** Fetch all pages of a GitHub API list endpoint (handles pagination automatically).
 *  GitHub returns max 100 items per page — this fetches until no next page. */
async function fetchAllPages(baseUrl: string, headers: Record<string, string>): Promise<any[]> {
  const all: any[] = [];
  let url: string | null = baseUrl.includes('per_page')
    ? baseUrl
    : (baseUrl.includes('?') ? `${baseUrl}&per_page=100` : `${baseUrl}?per_page=100`);

  while (url !== null) {
    const currentUrl: string = url;
    const res = await httpsGet(currentUrl, headers);
    if (!res.ok) { break; }

    const page = await res.json() as any[];
    if (Array.isArray(page)) { all.push(...page); }

    // GitHub Link header: <https://...?page=2>; rel="next", <https://...?page=5>; rel="last"
    const linkHeader: string = res._linkHeader ?? '';
    const nextMatch: RegExpMatchArray | null = linkHeader.match(/<([^>]+)>;\s*rel="next"/);
    url = nextMatch ? nextMatch[1] : null;
  }

  return all;
}

async function fetchFromGitHub(prUrl: string, token: string): Promise<PRData> {
  const parsed = parsePRUrl(prUrl);
  if (!parsed) { throw new Error('Cannot parse GitHub PR URL: ' + prUrl); }

  const { owner, repo, number } = parsed;
  const base = `https://api.github.com/repos/${owner}/${repo}`;
  const headers: Record<string, string> = {
    Authorization: `Bearer ${token}`,
    Accept: 'application/vnd.github+json',
    'X-GitHub-Api-Version': '2022-11-28',
    'User-Agent': 'code-review-checklist-agent',
  };

  // Fetch PR info and static data in parallel, paginated comments sequentially after
  const [prRes, filesRes, reqReviewersRes] = await Promise.all([
    httpsGet(`${base}/pulls/${number}`,                          headers),
    httpsGet(`${base}/pulls/${number}/files?per_page=100`,       headers),
    httpsGet(`${base}/pulls/${number}/requested_reviewers`,      headers),
  ]);

  if (!prRes.ok) {
    const body = await prRes.text();
    const hint = prRes.status === 401 ? ' — PAT expired or invalid. Run @code-review /token to update.' :
                 prRes.status === 403 ? ' — PAT lacks "repo" scope or rate limited.' :
                 prRes.status === 404 ? ' — PR not found. Check URL and repo access.' : '';
    throw new Error(`GitHub API ${prRes.status}${hint}: ${body.slice(0, 200)}`);
  }

  // Paginated comment fetches — all pages, no 100-item cap
  const [reviews, inlines, issueComments] = await Promise.all([
    fetchAllPages(`${base}/pulls/${number}/reviews`,   headers),   // review summaries
    fetchAllPages(`${base}/pulls/${number}/comments`,  headers),   // inline code comments
    fetchAllPages(`${base}/issues/${number}/comments`, headers),   // full conversation thread
  ]);

  const pr           = await prRes.json() as any;
  const rawFiles     = filesRes.ok        ? (await filesRes.json() as any[]) : [];
  const reqReviewers = reqReviewersRes.ok ? (await reqReviewersRes.json() as any) : {};

  // Build reviewer list: people who submitted a review + people still requested
  const reviewerSet = new Set<string>();
  reviews.forEach((r: any) => { if (r.user?.login) { reviewerSet.add(r.user.login); } });
  (reqReviewers.users ?? []).forEach((u: any) => { if (u.login) { reviewerSet.add(u.login); } });
  (reqReviewers.teams ?? []).forEach((t: any) => { if (t.slug) { reviewerSet.add(t.slug); } });
  const reviewers = Array.from(reviewerSet);

  const comments: string[] = [
    // Review-level summary comments (Approved / Request changes + body)
    ...reviews
      .filter((r: any) => r.body?.trim())
      .map((r: any) => `[Review by ${r.user?.login}] ${r.body.trim()}`),
    // Inline code comments on specific diff lines
    ...inlines
      .filter((c: any) => c.body?.trim())
      .map((c: any) => `[${c.user?.login} on ${c.path}:${c.line ?? c.original_line}] ${c.body.trim()}`),
    // General PR conversation thread — Copilot bot, LGTM, general discussion
    ...issueComments
      .filter((c: any) => c.body?.trim())
      .map((c: any) => `[${c.user?.login} (conversation)] ${c.body.trim()}`),
  ];

  const files: FilePatch[] = rawFiles.map((f: any) => ({
    filename:  f.filename,
    status:    f.status,
    additions: f.additions,
    deletions: f.deletions,
    patch:     f.patch ? truncatePatch(f.patch) : '(binary or no diff)',
  }));

  const diffSummary = rawFiles
    .map((f: any) => `  ${f.status.padEnd(8)} ${f.filename}  (+${f.additions}/-${f.deletions})`)
    .join('\n');

  return {
    url: prUrl,
    title: pr.title ?? '',
    body:  pr.body  ?? '',
    author: pr.user?.login ?? '',
    assignees: (pr.assignees ?? []).map((a: any) => a.login).filter(Boolean),
    reviewers,
    branch: pr.head?.ref ?? '',
    baseBranch: pr.base?.ref ?? 'main',
    createdAt: pr.created_at ?? null,
    closedAt:  pr.merged_at ?? (pr.state === 'closed' ? pr.closed_at : null) ?? null,
    comments,
    files,
    diffSummary,
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// Local git diff — used when no GitHub token is supplied but we're in a repo
// ─────────────────────────────────────────────────────────────────────────────

async function fetchLocalDiff(repoDir: string, branch: string, baseBranch: string): Promise<FilePatch[]> {
  // Make sure the branch is available locally (fetch quietly)
  try { await execAsync(`git fetch origin ${branch} ${baseBranch} --quiet`, { cwd: repoDir }); } catch { /* ignore */ }

  const mergeBase = (await execAsync(
    `git merge-base origin/${baseBranch} origin/${branch}`,
    { cwd: repoDir }
  )).stdout.trim();

  // File-by-file stat
  const statOut = (await execAsync(
    `git diff --name-status ${mergeBase} origin/${branch}`,
    { cwd: repoDir }
  )).stdout.trim();

  const files: FilePatch[] = [];

  for (const line of statOut.split('\n').filter(Boolean)) {
    const [statusChar, ...rest] = line.split('\t');
    const filename = rest[rest.length - 1];
    const statusMap: Record<string, string> = { A: 'added', M: 'modified', D: 'removed', R: 'renamed' };
    const status = statusMap[statusChar[0]] ?? statusChar;

    let patch = '';
    try {
      const raw = (await execAsync(
        `git diff ${mergeBase} origin/${branch} -- "${filename}"`,
        { cwd: repoDir, maxBuffer: 1024 * 512 }
      )).stdout;
      patch = truncatePatch(raw);
    } catch { patch = '(diff unavailable)'; }

    // Count +/- lines
    const additions = (patch.match(/^\+(?!\+\+)/mg) ?? []).length;
    const deletions  = (patch.match(/^-(?!--)/mg) ?? []).length;

    files.push({ filename, status, additions, deletions, patch });
  }

  return files;
}

// ─────────────────────────────────────────────────────────────────────────────
// AI analysis — reviews BOTH comments AND actual code diff
// ─────────────────────────────────────────────────────────────────────────────

const SYSTEM_PROMPT = `You are a senior software engineer performing a formal code review using a structured checklist.

You will receive:
1. A list of 25 numbered checklist rows (ID, area, rule)
2. PR review comments from human reviewers and GitHub Copilot
3. The actual git diff (unified patch format) of the changed files

Evaluate EVERY checklist row against BOTH the comments AND the code diff.

## Column definitions (match the Excel template exactly):
- authorStatus (col E): Did the author address this? "Yes" | "No" | "NA"
  - "Yes"  = the code shows this was handled correctly / issue is fixed in the diff
  - "No"   = the code shows this was NOT handled / issue still exists
  - "NA"   = this category does not apply to the changed files
- reviewerStatus (col F): Reviewer verdict: "Ok" | "Not Ok" | "NA"
  - "Ok"     = reviewer confirms this area looks good
  - "Not Ok" = reviewer found a problem here
  - "NA"     = not applicable
- finding (col G): 1-2 sentence description of what was found. Empty string if NA/Ok with nothing to note.
- defectType (col H): "Functional" | "Technical" | "Process/Compliance" | "Documentation" | "Others" | ""
  - Use "" when reviewerStatus is "Ok" or "NA"
- remarks (col I): Brief suggested author remark or action needed. Empty string if not applicable.

## Fixed/addressed detection:
- Lines starting with "+" are additions, "-" are deletions
- If a comment raised an issue AND the "+" lines show the fix → authorStatus="Yes", reviewerStatus="Ok", finding="FIXED: ..."
- If an issue is detected in "-" lines but corrected in "+" lines → same as above
- If a problem still exists in "+" lines → authorStatus="No", reviewerStatus="Not Ok"

## Output format — CRITICAL:
- Output ONLY a raw JSON array. No markdown, no code fences, no prose.
- Start with [ and end with ]
- Include ALL 25 rows. Do not skip any.
- Each element must have exactly: { "rowId": number, "authorStatus": string, "reviewerStatus": string, "finding": string, "defectType": string, "remarks": string }`;

function buildDiffSection(files: FilePatch[]): string {
  if (!files.length) { return 'No code diff available.'; }
  return files.map(f =>
    `### File: ${f.filename} [${f.status}] +${f.additions}/-${f.deletions}\n\`\`\`diff\n${f.patch}\n\`\`\``
  ).join('\n\n');
}

/** Robustly extract a JSON array from the raw LLM response */
function extractJsonArray(raw: string): string {
  const cleaned = raw.replace(/```(?:json)?/gi, '').trim();
  const start = cleaned.indexOf('[');
  const end   = cleaned.lastIndexOf(']');
  if (start !== -1 && end !== -1 && end > start) {
    return cleaned.slice(start, end + 1);
  }
  return cleaned;
}

async function analyseWithAI(
  prData: PRData,
  model: vscode.LanguageModelChat,
  token: vscode.CancellationToken,
  stream: vscode.ChatResponseStream,
  userSpec: string,
  jiraStories: JiraStory[]
): Promise<RowFinding[]> {

  const checklistContext = CHECKLIST_ROWS
    .map(r => `Row ${r.id} [${r.reviewArea}] — ${r.checkCategory}: ${r.description}`)
    .join('\n');

  const commentsSection = prData.comments.length
    ? prData.comments.map((c, i) => `${i + 1}. ${c}`).join('\n')
    : 'No PR review comments provided.';

  const diffSection = buildDiffSection(prData.files);
  const diffTruncated = diffSection.length > 24000
    ? diffSection.slice(0, 24000) + '\n\n… [diff truncated — remaining files not shown]'
    : diffSection;

  const specSection = userSpec.trim()
    ? `\n## Reviewer Specification (HIGH PRIORITY — follow these instructions first)\n${userSpec.trim()}\n`
    : '';

  // Build Jira section — one block per story, all ACs combined
  const jiraSection = jiraStories.length > 0 ? [
    `\n## Jira User Stor${jiraStories.length > 1 ? 'ies' : 'y'} (${jiraStories.length})`,
    `Check whether the code satisfies the acceptance criteria of ALL stories below.`,
    `If any AC is not met, mark the relevant row as "Not Ok".\n`,
    ...jiraStories.map(s => [
      `### [${s.key}] ${s.summary}`,
      `- Type: ${s.type}  |  Status: ${s.status}  |  URL: ${s.url}`,
      s.labels.length ? `- Labels: ${s.labels.join(', ')}` : '',
      '',
      s.description ? `**Description:**\n${s.description}` : '',
      s.acceptanceCriteria ? `**Acceptance Criteria:**\n${s.acceptanceCriteria}` : '',
    ].filter(Boolean).join('\n')),
  ].join('\n') : '';

  const userPrompt = `## PR Information
- Title:   ${prData.title}
- Author:  ${prData.author}
- Branch:  ${prData.branch} → ${prData.baseBranch}
- URL:     ${prData.url}
${specSection}${jiraSection}
## 25-Row Checklist
${checklistContext}

## PR Review Comments (${prData.comments.length} total)
${commentsSection}

## Code Diff (${prData.files.length} file(s) changed)
${diffTruncated}

Evaluate all 25 checklist rows and return the JSON array now.`;

  const messages = [
    vscode.LanguageModelChatMessage.User(SYSTEM_PROMPT),
    vscode.LanguageModelChatMessage.User(userPrompt),
  ];

  stream.markdown('⏳ Waiting for AI response…\n');
  const response = await model.sendRequest(messages, {}, token);
  let raw = '';
  for await (const chunk of response.text) { raw += chunk; }
  stream.markdown(`📥 AI responded (${raw.length} chars). Parsing…\n`);

  const jsonStr = extractJsonArray(raw);
  try {
    const parsed = JSON.parse(jsonStr) as RowFinding[];
    if (!Array.isArray(parsed) || parsed.length === 0) {
      throw new Error(`Expected non-empty array, got ${typeof parsed}`);
    }
    return parsed;
  } catch (err: any) {
    stream.markdown([
      `\n⚠️ **AI returned unparseable output** — findings will be blank in Excel.`,
      `**Error:** ${err.message}`,
      `**Raw preview:**\n\`\`\`\n${raw.slice(0, 600)}\n\`\`\``,
      `\nTip: Try again — the model occasionally produces malformed output.`,
    ].join('\n'));
    console.error('[checklist] Parse error:', err.message, '\nRaw:\n', raw);
    return [];
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Review summary — second AI call, plain-English, uses findings as input
// ─────────────────────────────────────────────────────────────────────────────

async function generateReviewSummary(
  prData: PRData,
  findings: RowFinding[],
  userSpec: string,
  jiraStories: JiraStory[],
  model: vscode.LanguageModelChat,
  token: vscode.CancellationToken
): Promise<string> {

  const notOk   = findings.filter(f => f.reviewerStatus === 'Not Ok');
  const ok      = findings.filter(f => f.reviewerStatus === 'Ok');
  const na      = findings.filter(f => f.reviewerStatus === 'NA');

  const findingsSummary = findings
    .filter(f => f.reviewerStatus !== 'NA' && f.finding)
    .map(f => {
      const row = CHECKLIST_ROWS.find(r => r.id === f.rowId);
      return `[${f.reviewerStatus}] Row ${f.rowId} ${row?.reviewArea}: ${f.finding}`;
    })
    .join('\n');

  const specNote  = userSpec ? `\nReviewer specification applied:\n${userSpec}\n` : '';
  const jiraNote  = jiraStories.length
    ? `\nJira stories: ${jiraStories.map(s => `[${s.key}] ${s.summary}`).join('; ')}\n`
    : '';

  const prompt = `You are writing a brief review summary for a developer who just ran an automated code review checklist.

PR: ${prData.title} (${prData.url})
Branch: ${prData.branch} → ${prData.baseBranch}
Files changed: ${prData.files.length} file(s)
${specNote}${jiraNote}
Checklist results:
- Not Ok: ${notOk.length} rows
- Ok: ${ok.length} rows  
- NA: ${na.length} rows (not applicable to this PR)

Findings:
${findingsSummary || 'No specific findings.'}

Write a concise review summary (150-250 words) that:
1. Opens with 1 sentence describing what this PR is about based on the branch name${jiraStories.length ? ' and the Jira stories' : ''}
2. Explains what areas were reviewed and what the reviewer spec influenced (if any)
3. Lists the key issues found (Not Ok rows) with plain-English explanation — no jargon, no row numbers
4. Mentions what looked clean (Ok rows worth calling out)
5. Notes what was intentionally skipped and why (NA rows if spec explains it)
${jiraStories.length ? '6. Notes which Jira story acceptance criteria appear to be met or not met\n7.' : '6.'} Closes with a clear verdict: Ready to merge / Needs fixes before merge

Format as clean markdown with short paragraphs. No bullet lists — prose only.`;

  const messages = [vscode.LanguageModelChatMessage.User(prompt)];
  const response = await model.sendRequest(messages, {}, token);
  let summary = '';
  for await (const chunk of response.text) { summary += chunk; }
  return summary.trim();
}

// ─────────────────────────────────────────────────────────────────────────────
// Index store — code-review/index.json tracks all versions per PR URL
// ─────────────────────────────────────────────────────────────────────────────

interface IndexEntry {
  version:    number;
  file:       string;   // absolute path
  date:       string;   // YYYY-MM-DD
  notOkCount: number;
  okCount:    number;
  branch:     string;
}

interface IndexFile {
  // key = PR URL (or branch name for local-git mode)
  [prKey: string]: IndexEntry[];
}


// ─────────────────────────────────────────────────────────────────────────────
// ContextWriter — writes a live .context.md that Copilot Chat reads via stream.reference()
// ─────────────────────────────────────────────────────────────────────────────

class ContextWriter {
  private contextPath: string;
  private sections: string[] = [];

  constructor(codeReviewDir: string) {
    this.contextPath = path.join(codeReviewDir, '.context.md');
    fs.mkdirSync(codeReviewDir, { recursive: true });
    this.sections = [
      `# Code Review Checklist Agent — Session Context`,
      `_Generated: ${new Date().toISOString()}_`,
      `_Reference this file in Copilot Chat for debugging or follow-up questions._`,
      '',
    ];
    this.flush();
  }

  add(heading: string, content: string): void {
    this.sections.push(`## ${heading}`, '', content, '');
    this.flush();
  }

  addError(heading: string, err: any, ctx: Record<string, string> = {}): void {
    const ctxLines = Object.entries(ctx).map(([k, v]) => `- ${k}: ${v}`);
    const msg = err?.message ?? String(err);
    const hints = [
      msg.includes('401')       ? '- GitHub PAT expired → run `@code-review /token <new_PAT>`' : '',
      msg.includes('403')       ? '- PAT missing `repo` scope → regenerate at https://github.com/settings/tokens' : '',
      msg.includes('ENOTFOUND') ? '- DNS failed → check VPN / internet connection' : '',
      msg.includes('timed out') ? '- Request timed out → check network, try again' : '',
      msg.includes('CERT')      ? '- SSL cert error → corporate proxy may intercept HTTPS' : '',
      msg.includes('Jira')      ? '- Check Jira credentials → `@code-review /jira status`' : '',
    ].filter(Boolean);

    this.sections.push(
      `## ❌ ${heading}`,
      '',
      `**Error:** \`${msg}\``,
      '',
      ctxLines.length ? `**Context:**
${ctxLines.join('\n')}` : '',
      hints.length    ? `**What to check:**
${hints.join('\n')}` : '',
      '',
    );
    this.flush();
  }

  get uri(): vscode.Uri { return vscode.Uri.file(this.contextPath); }

  private flush(): void {
    try { fs.writeFileSync(this.contextPath, this.sections.join('\n'), 'utf8'); } catch { /* non-fatal */ }
  }
}

class IndexStore {
  private indexPath: string;
  private data: IndexFile = {};

  constructor(private readonly codeReviewDir: string) {
    this.indexPath = path.join(codeReviewDir, 'index.json');
  }

  async load(): Promise<void> {
    try {
      const raw = fs.readFileSync(this.indexPath, 'utf8');
      this.data = JSON.parse(raw) as IndexFile;
    } catch {
      this.data = {};   // first run — start fresh
    }
  }

  /** Returns all past versions for this PR, sorted oldest-first */
  getVersions(prKey: string): IndexEntry[] {
    return (this.data[prKey] ?? []).sort((a, b) => a.version - b.version);
  }

  /** Returns the next version number for this PR (1 if never seen before) */
  nextVersion(prKey: string): number {
    const versions = this.getVersions(prKey);
    return versions.length === 0 ? 1 : versions[versions.length - 1].version + 1;
  }

  /** Registers a new version and persists the index */
  register(prKey: string, entry: IndexEntry): void {
    if (!this.data[prKey]) { this.data[prKey] = []; }
    this.data[prKey].push(entry);
    fs.mkdirSync(path.dirname(this.indexPath), { recursive: true });
    fs.writeFileSync(this.indexPath, JSON.stringify(this.data, null, 2), 'utf8');
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Excel generator — fills the original template file directly
// ─────────────────────────────────────────────────────────────────────────────

async function generateExcel(
  prData: PRData,
  findings: RowFinding[],
  templatePath: string,
  outputPath: string,
  version: number,
  previousVersions: IndexEntry[],
  jiraStories: JiraStory[],
  detectedJiraKeys: string[]
): Promise<void> {

  // eslint-disable-next-line @typescript-eslint/no-require-imports
  const XLSX = require('xlsx');

  // Read template — cellStyles:true preserves all formatting, fills, borders, merges
  const wb = XLSX.readFile(templatePath, { cellStyles: true, cellFormula: true, bookVBA: false });
  const ws = wb.Sheets['ROR'];
  if (!ws) { throw new Error('Sheet "ROR" not found in template'); }

  const today = new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });

  /** Set a cell value, preserving its existing style from the template */
  function setVal(addr: string, val: string | number | null): void {
    if (val === null || val === undefined || val === '') { return; }
    const existing = ws[addr];
    if (existing) {
      // Update value only — keep style, formula references, etc.
      existing.v = val;
      existing.t = typeof val === 'number' ? 'n' : 's';
      delete existing.f;   // remove formula — we're writing a computed value
      delete existing.w;   // remove cached formatted text so Excel recalculates
    } else {
      ws[addr] = { t: typeof val === 'number' ? 'n' : 's', v: val };
    }
  }

  // ── Fill header cells ─────────────────────────────────────────────────────
  const keyBrackets = (jiraStories.length > 0
    ? jiraStories.map((s: any) => s.key)
    : detectedJiraKeys
  ).map((k: string) => `[${k}]`).join(' ');

  setVal('C4', keyBrackets ? `${prData.url} ${keyBrackets}` : prData.url);
  setVal('C5', `v${version}`);

  const authorVal = prData.assignees.length ? prData.assignees.join(', ') : prData.author;
  if (authorVal) { setVal('C6', authorVal); }
  if (prData.reviewers.length) { setVal('C7', `SME - ${prData.reviewers.join(', ')}`); }

  setVal('C8', fmtDate(prData.createdAt) ?? today);
  const endDate = fmtDate(prData.closedAt);
  if (endDate) { setVal('C9', endDate); }

  // ── Fill checklist rows 11-35 (cols E, F, G, H, I) ───────────────────────
  const findingMap = new Map<number, RowFinding>();
  findings.forEach(f => findingMap.set(f.rowId, f));

  for (let rowId = 1; rowId <= 25; rowId++) {
    const rowNum = rowId + 10;
    const f = findingMap.get(rowId);
    if (!f) { continue; }
    setVal(`E${rowNum}`, f.authorStatus);
    setVal(`F${rowNum}`, f.reviewerStatus);
    setVal(`G${rowNum}`, f.finding);
    setVal(`H${rowNum}`, f.defectType);
    setVal(`I${rowNum}`, f.remarks);
  }

  // ── Pre-compute summary counts (bypasses stale formula cache) ─────────────
  const countE = (v: string) => findings.filter(f => f.authorStatus   === v).length;
  const countF = (v: string) => findings.filter(f => f.reviewerStatus === v).length;
  const countH = (v: string) => findings.filter(f => f.defectType     === v).length;

  setVal('J4', countE('Yes'));     setVal('J5', countE('No'));    setVal('J6', countE('NA'));
  setVal('J8', countF('Ok'));      setVal('J9', countF('Not Ok'));
  setVal('H3', countH('Functional'));   setVal('H4', countH('Technical'));
  setVal('H5', countH('Process/Compliance'));
  setVal('H6', countH('Documentation')); setVal('H7', countH('Others'));
  setVal('H9', findings.filter(f => f.defectType !== '').length);

  // ── Re-review Required sheet ──────────────────────────────────────────────
  const notOkFindings = findings.filter(f => f.reviewerStatus === 'Not Ok');
  if (notOkFindings.length > 0) {
    const rrData: any[][] = [
      ['#', 'Review Area', 'Description of Finding', 'Defect Type', 'Suggested Action', 'Re-review Status'],
      ...notOkFindings.map(f => [
        f.rowId,
        CHECKLIST_ROWS.find(r => r.id === f.rowId)?.reviewArea ?? '',
        f.finding,
        f.defectType,
        f.remarks,
        '',
      ]),
    ];
    const rrSheet = XLSX.utils.aoa_to_sheet(rrData);
    rrSheet['!cols'] = [5, 28, 60, 22, 50, 18].map(w => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, rrSheet, 'Re-review Required');
  }

  // ── Version History sheet ─────────────────────────────────────────────────
  const histData: any[][] = [
    ['Version', 'Date', 'Not Ok', 'Ok', 'NA', 'Comments'],
    ...previousVersions.map((v: IndexEntry) => [
      `v${v.version}`, v.date, v.notOkCount, v.okCount, 0, '',
    ]),
    [
      `v${version} (current)`,
      new Date().toISOString().slice(0, 10),
      notOkFindings.length,
      countF('Ok'),
      findings.filter(f => f.reviewerStatus === 'NA').length,
      '',
    ],
  ];
  const histSheet = XLSX.utils.aoa_to_sheet(histData);
  histSheet['!cols'] = [16, 14, 10, 10, 10, 60].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, histSheet, 'Version History');

  // ── Set ROR as the active tab ─────────────────────────────────────────────
  if (!wb.Workbook) { wb.Workbook = {}; }
  if (!wb.Workbook.Views) { wb.Workbook.Views = [{}]; }
  wb.Workbook.Views[0].RTL    = false;
  wb.Workbook.Views[0].ActiveTab = 0;  // ROR is always index 0

  // ── Write with maximum Windows/Mac/SharePoint compatibility ───────────────
  XLSX.writeFile(wb, outputPath, {
    bookSST:     true,   // shared string table — required by Windows Excel
    compression: true,   // deflate compression
    bookType:    'xlsx',
  });
}


// ─────────────────────────────────────────────────────────────────────────────
// Jira credential store — email + API token, both in SecretStorage
// ─────────────────────────────────────────────────────────────────────────────

const JIRA_EMAIL_KEY = 'checklist.jira.email';
const JIRA_TOKEN_KEY = 'checklist.jira.token';
const JIRA_HOST      = 'https://decisiv.atlassian.net';

interface JiraCreds { email: string; token: string; }

class JiraStore {
  constructor(private readonly secrets: vscode.SecretStorage) {}

  async save(email: string, token: string): Promise<void> {
    await this.secrets.store(JIRA_EMAIL_KEY, email.trim());
    await this.secrets.store(JIRA_TOKEN_KEY, token.trim());
  }

  async get(): Promise<JiraCreds | null> {
    const email = await this.secrets.get(JIRA_EMAIL_KEY);
    const token = await this.secrets.get(JIRA_TOKEN_KEY);
    return email && token ? { email, token } : null;
  }

  async delete(): Promise<void> {
    await this.secrets.delete(JIRA_EMAIL_KEY);
    await this.secrets.delete(JIRA_TOKEN_KEY);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Jira story fetcher
// ─────────────────────────────────────────────────────────────────────────────

interface JiraStory {
  key:         string;
  url:         string;
  summary:     string;
  type:        string;
  status:      string;
  sprint:      string;        // e.g. "Sprint 42" — from customfield_10020
  labels:      string[];
  description: string;
  acceptanceCriteria: string;
}

/** Extract plain text from Atlassian Document Format (ADF) node recursively */
function adfToText(node: any): string {
  if (!node) { return ''; }
  if (node.type === 'text') { return node.text ?? ''; }
  if (Array.isArray(node.content)) {
    return node.content.map(adfToText).join(node.type === 'paragraph' ? '\n' : '');
  }
  return '';
}

/** Pull acceptance criteria out of the description text — looks for AC headings */
function extractAC(fullText: string): string {
  const acMatch = fullText.match(
    /acceptance criteria[\s\S]*?(?=\n##|\n[A-Z][^\n]{0,30}\n[-=]{3,}|$)/i
  );
  return acMatch ? acMatch[0].trim() : '';
}

/** Parse sprint name → folder-safe string e.g. "PSSM 2.0 Sprint 42" → "Sprint-42" */
function parseSprintName(name: string): string {
  if (!name) { return 'No-Sprint'; }
  // Extract the last number in the name: "Sprint 42", "PSSM 2.0 Sprint 42", "S42"
  const num = name.match(/(\d+)\s*$/)?.[1] ?? name.match(/(\d+)/)?.[1];
  return num ? `Sprint-${num}` : name.replace(/\s+/g, '-');
}

/** Extract sprint from all possible field locations in a Jira issue's fields object */
function extractSprintFromFields(fields: any): string {
  // Try every field value — the sprint field may be customfield_10020 or any other ID
  for (const val of Object.values(fields)) {
    if (Array.isArray(val) && val.length > 0 && val[0]?.state !== undefined && val[0]?.name) {
      // This looks like a sprint array (objects with .state and .name)
      const sprintArr = val as any[];
      const active = sprintArr.find((s: any) => s.state === 'active')
                  ?? sprintArr.find((s: any) => s.state === 'future')
                  ?? sprintArr.find((s: any) => s.state === 'closed')
                  ?? sprintArr[sprintArr.length - 1];
      if (active?.name) { return parseSprintName(active.name); }
    }
  }
  return 'No-Sprint';
}

async function fetchJiraStory(key: string, creds: JiraCreds): Promise<JiraStory> {
  const auth    = Buffer.from(`${creds.email}:${creds.token}`).toString('base64');
  // Request ALL fields — this ensures we get the sprint field regardless of its custom field ID
  const url     = `${JIRA_HOST}/rest/api/3/issue/${key}?expand=renderedFields`;
  const headers = { Authorization: `Basic ${auth}`, Accept: 'application/json' };

  const res = await httpsGet(url, headers);
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Jira API ${res.status} for ${key}: ${body.slice(0, 200)}`);
  }

  const data   = await res.json() as any;
  const fields = data.fields ?? {};

  const descriptionText = fields.description ? adfToText(fields.description) : '';
  const acText          = extractAC(descriptionText);

  // Dynamically find sprint from any custom field — no hardcoded field ID
  const sprint = extractSprintFromFields(fields);

  return {
    key,
    url:         `${JIRA_HOST}/browse/${key}`,
    summary:     fields.summary ?? '',
    type:        fields.issuetype?.name ?? 'Issue',
    status:      fields.status?.name ?? '',
    sprint,
    labels:      (fields.labels ?? []) as string[],
    description: descriptionText.slice(0, 1500),
    acceptanceCriteria: acText.slice(0, 1000),
  };
}

/** Detect ALL Jira ticket keys from branch name, PR title, and PR description body */
function detectJiraKeys(prData: PRData): string[] {
  const JIRA_PATTERN = /\b([A-Z][A-Z0-9]+-\d+)\b/g;
  // Also match full Jira URLs: decisiv.atlassian.net/browse/SRM2-2072
  const JIRA_URL_PATTERN = /atlassian\.net\/browse\/([A-Z][A-Z0-9]+-\d+)/g;
  const found = new Set<string>();

  for (const text of [prData.branch, prData.title, prData.body]) {
    if (!text) { continue; }
    for (const m of text.matchAll(JIRA_PATTERN))     { found.add(m[1]); }
    for (const m of text.matchAll(JIRA_URL_PATTERN)) { found.add(m[1]); }
  }
  return Array.from(found);
}

/** Fetch multiple Jira stories in parallel, skip ones that fail */
async function fetchJiraStories(
  keys: string[],
  creds: JiraCreds,
  stream: vscode.ChatResponseStream
): Promise<JiraStory[]> {
  const results = await Promise.allSettled(
    keys.map(key => fetchJiraStory(key, creds))
  );
  const stories: JiraStory[] = [];
  results.forEach((r, i) => {
    if (r.status === 'fulfilled') {
      stories.push(r.value);
    } else {
      stream.markdown(`⚠️ Could not fetch **${keys[i]}**: ${r.reason?.message ?? r.reason}\n`);
    }
  });
  return stories;
}

// ─────────────────────────────────────────────────────────────────────────────
// Jira Development Info — fetches linked GitHub PRs from a Jira issue
// ─────────────────────────────────────────────────────────────────────────────

interface LinkedPR {
  url:    string;
  title:  string;
  status: string;   // open | merged | declined
}

/** Fetch GitHub PRs linked to a Jira issue via the Jira dev-info API */
async function fetchLinkedPRs(key: string, creds: JiraCreds): Promise<LinkedPR[]> {
  const auth = Buffer.from(`${creds.email}:${creds.token}`).toString('base64');
  const headers: Record<string, string> = {
    Authorization: `Basic ${auth}`,
    Accept: 'application/json',
    'User-Agent': 'code-review-checklist-agent',
  };

  // Step 1 — get the numeric issue ID (required by dev-info API)
  const issueRes = await httpsGet(
    `${JIRA_HOST}/rest/api/3/issue/${key}?fields=id`,
    headers
  );
  if (!issueRes.ok) {
    throw new Error(`Could not fetch issue ID for ${key}: HTTP ${issueRes.status}`);
  }
  const issueData = await issueRes.json() as any;
  const issueId   = issueData.id as string;

  // Step 2 — fetch dev info
  const devRes = await httpsGet(
    `${JIRA_HOST}/rest/dev-info/0.10/issue/detail?issueId=${issueId}&applicationType=github&dataType=pullrequest`,
    headers
  );

  if (!devRes.ok) {
    // Dev-info API requires Jira Software (not Jira Core) — graceful fallback
    throw new Error(`Dev-info API returned HTTP ${devRes.status} for ${key}. Ensure your Jira plan includes development panel.`);
  }

  const devData = await devRes.json() as any;

  // Collect PRs from all detail entries
  const prs: LinkedPR[] = [];
  const details: any[] = devData.detail ?? [];
  for (const detail of details) {
    for (const pr of (detail.pullRequests ?? [])) {
      if (pr.url) {
        prs.push({
          url:    pr.url,
          title:  pr.name  ?? pr.url,
          status: pr.status ?? 'unknown',
        });
      }
    }
  }
  return prs;
}

/** For --list mode: resolve each item to a PR URL.
 *  Items can be Jira keys (auto-fetch PR) or direct PR URLs. */
async function resolveListItems(
  items: string[],
  creds: JiraCreds | null,
  stream: vscode.ChatResponseStream
): Promise<string[]> {

  const JIRA_KEY = /^[A-Z][A-Z0-9]+-\d+$/;
  const PR_URL   = /github\.com\/[^/]+\/[^/]+\/pull\/\d+/;
  const resolved: string[] = [];

  for (const item of items) {
    if (PR_URL.test(item)) {
      // Direct PR URL — use as-is
      resolved.push(item);
      stream.markdown(`✅ \`${item}\` — direct PR URL\n`);
      continue;
    }

    if (JIRA_KEY.test(item.toUpperCase())) {
      const key = item.toUpperCase();
      if (!creds) {
        stream.markdown(`⚠️ **${key}** — Jira credentials not set. Use \`@code-review /jira\` to save them.\n`);
        continue;
      }

      stream.markdown(`🔗 Fetching PRs linked to **${key}**…\n`);
      let prs: LinkedPR[] = [];
      try {
        prs = await fetchLinkedPRs(key, creds);
      } catch (err: any) {
        stream.markdown(`⚠️ **${key}** — ${err.message}\n`);
        continue;
      }

      if (prs.length === 0) {
        stream.markdown(`ℹ️ **${key}** — no linked PRs found in Jira development panel.\n`);
        continue;
      }

      if (prs.length === 1) {
        stream.markdown(`✅ **${key}** → \`${prs[0].url}\` _(${prs[0].status})_\n`);
        resolved.push(prs[0].url);
        continue;
      }

      // Multiple PRs — show quick pick
      stream.markdown(`🔀 **${key}** has **${prs.length}** linked PRs — showing picker…\n`);
      const picked = await vscode.window.showQuickPick(
        prs.map(p => ({
          label:       p.url.match(/\/pull\/(\d+)$/)?.[0] ?? p.url,
          description: p.status,
          detail:      p.title,
          url:         p.url,
        })),
        {
          title:        `${key} — Pick a PR to review`,
          placeHolder:  'Select the PR you want to generate a checklist for',
          canPickMany:  false,
        }
      );

      if (picked) {
        stream.markdown(`✅ **${key}** → selected \`${(picked as any).url}\`\n`);
        resolved.push((picked as any).url);
      } else {
        stream.markdown(`⏭️ **${key}** — skipped (no PR selected)\n`);
      }
      continue;
    }

    stream.markdown(`⚠️ \`${item}\` — not a valid PR URL or Jira key, skipping.\n`);
  }

  return resolved;
}

const SECRET_KEY = 'checklist.github.pat';

class TokenStore {
  constructor(private readonly secrets: vscode.SecretStorage) {}

  async save(pat: string): Promise<void> {
    await this.secrets.store(SECRET_KEY, pat.trim());
  }

  async get(): Promise<string | undefined> {
    return this.secrets.get(SECRET_KEY);
  }

  async delete(): Promise<void> {
    await this.secrets.delete(SECRET_KEY);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Free-form conversational handler
// Answers any question using checklist knowledge + session context + history
// ─────────────────────────────────────────────────────────────────────────────

async function handleConversation(
  question: string,
  attachedFileContent: string,
  workspaceRoot: string,
  model: vscode.LanguageModelChat,
  token: vscode.CancellationToken,
  stream: vscode.ChatResponseStream
): Promise<void> {

  // Load last run context
  const ctxPath = path.join(workspaceRoot, 'code-review', '.context.md');
  const lastContext = fs.existsSync(ctxPath)
    ? fs.readFileSync(ctxPath, 'utf8').slice(0, 6000)
    : '(No previous review context found. Run @code-review /generate first.)';

  // Load review history summary from index.json
  const indexPath = path.join(workspaceRoot, 'code-review', 'index.json');
  let historySummary = '(No review history yet.)';
  if (fs.existsSync(indexPath)) {
    try {
      const idx = JSON.parse(fs.readFileSync(indexPath, 'utf8')) as Record<string, any[]>;
      const entries = Object.entries(idx);
      if (entries.length > 0) {
        historySummary = entries.map(([prKey, versions]) => {
          const last = versions[versions.length - 1];
          return `- ${prKey}: ${versions.length} version(s), last: v${last.version} on ${last.date} (${last.notOkCount} Not Ok)`;
        }).join('\n');
      }
    } catch { /* ignore */ }
  }

  const checklistRules = CHECKLIST_ROWS
    .map(r => `Row ${r.id} [${r.reviewArea}] — ${r.checkCategory}: ${r.description}`)
    .join('\n');

  const systemMsg = `You are the Code Review Checklist Agent for the PSSM2.0 project at Decisiv.

You have deep knowledge of the 25-row PSSM2.0 source code review checklist. You help developers:
- Understand checklist rules and how to apply them
- Analyse code review findings and suggest fixes
- Navigate review history and track progress
- Answer questions about past reviews, sprint status, or defect patterns
- Process files (Excel, CSV, text) containing PR lists or review data

## The 25-Row PSSM2.0 Checklist
${checklistRules}

## Last Review Session Context
${lastContext}

## Review History
${historySummary}

Answer the user's question helpfully and concisely. 
If they're asking about code fixes, reference the specific diff from the context above.
If they give you a file, extract actionable items and explain what you found.`;

  const userMsg = attachedFileContent
    ? `${question}\n\n## Attached file content:\n${attachedFileContent}`
    : question;

  const messages = [
    vscode.LanguageModelChatMessage.User(systemMsg),
    vscode.LanguageModelChatMessage.User(userMsg),
  ];

  const response = await model.sendRequest(messages, {}, token);
  for await (const chunk of response.text) {
    stream.markdown(chunk);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Extension entry point
// ─────────────────────────────────────────────────────────────────────────────

export function activate(context: vscode.ExtensionContext) {

  const tokenStore = new TokenStore(context.secrets);
  const jiraStore  = new JiraStore(context.secrets);

  const participant = vscode.chat.createChatParticipant('decisiv-pssm.code-review', async (
    request: vscode.ChatRequest,
    _ctx: vscode.ChatContext,
    stream: vscode.ChatResponseStream,
    token: vscode.CancellationToken
  ) => {

    // ── /help ─────────────────────────────────────────────────────────────
    if (request.command === 'help' || !request.prompt.trim()) {
      stream.markdown(HELP_TEXT);
      return;
    }

    // ── /token — save, show status, or clear the stored PAT ───────────────
    if (request.command === 'token') {
      const arg = request.prompt.trim();

      if (arg === 'clear' || arg === 'delete' || arg === 'remove') {
        await tokenStore.delete();
        stream.markdown('🗑️ GitHub PAT removed from secure storage.');
        return;
      }

      if (arg === 'status' || arg === 'check') {
        const existing = await tokenStore.get();
        if (existing) {
          stream.markdown(`✅ A GitHub PAT is stored (ends in \`…${existing.slice(-4)}\`).\nUse \`@code-review /token clear\` to remove it.`);
        } else {
          stream.markdown('ℹ️ No GitHub PAT stored yet.\nUse `@code-review /token <YOUR_PAT>` to save one.');
        }
        return;
      }

      if (arg) {
        // Validate it looks like a GitHub PAT (ghp_ / github_pat_ / gho_ prefix)
        if (!/^(ghp_|github_pat_|gho_|ghs_)/.test(arg)) {
          stream.markdown([
            '⚠️ That does not look like a GitHub PAT.',
            'GitHub PATs start with `ghp_`, `github_pat_`, `gho_`, or `ghs_`.',
            '',
            'Generate one at: https://github.com/settings/tokens',
            'Required scopes: **repo** (to read PRs and code)',
          ].join('\n'));
          return;
        }
        await tokenStore.save(arg);
        stream.markdown([
          `✅ GitHub PAT saved securely (ends in \`…${arg.slice(-4)}\`).`,
          '',
          '🔒 Stored in VS Code\'s encrypted secret storage (OS keychain) — never written to disk or any file.',
          '',
          'From now on, just use:',
          '```',
          '@code-review /generate https://github.com/org/repo/pull/42',
          '```',
          'No `--token` flag needed anymore.',
        ].join('\n'));
        return;
      }

      // No argument — show instructions
      stream.markdown([
        '## Save your GitHub PAT\n',
        '```',
        '@code-review /token <YOUR_GITHUB_PAT>',
        '```',
        '',
        '**Other commands:**',
        '- `@code-review /token status` — check if a PAT is stored',
        '- `@code-review /token clear` — remove the stored PAT',
        '',
        '**Generate a PAT:** https://github.com/settings/tokens',
        'Required scopes: `repo`',
      ].join('\n'));
      return;
    }

    // ── /jira — save, check, or clear Jira credentials ───────────────────
    if (request.command === 'jira') {
      const arg = request.prompt.trim();

      if (arg === 'clear' || arg === 'delete' || arg === 'remove') {
        await jiraStore.delete();
        stream.markdown('🗑️ Jira credentials removed from secure storage.');
        return;
      }

      if (arg === 'status' || arg === 'check') {
        const creds = await jiraStore.get();
        if (creds) {
          stream.markdown(`✅ Jira credentials stored for **${creds.email}**.\nUse \`@code-review /jira clear\` to remove.`);
        } else {
          stream.markdown([
            'ℹ️ No Jira credentials stored yet.',
            'Usage: `@code-review /jira <your@email.com> <ATLASSIAN_API_TOKEN>`',
            'Generate a token at: https://id.atlassian.net/manage-profile/security/api-tokens',
          ].join('\n'));
        }
        return;
      }

      // Expect: <email> <token>
      const parts = arg.split(/\s+/);
      if (parts.length >= 2) {
        const [email, apiToken] = parts;
        if (!email.includes('@')) {
          stream.markdown('⚠️ First argument must be your Atlassian email address.');
          return;
        }
        await jiraStore.save(email, apiToken);
        stream.markdown([
          `✅ Jira credentials saved securely for **${email}**.`,
          `🔒 Stored in VS Code's OS keychain — never written to disk.`,
          '',
          'The agent will now auto-fetch Jira stories when a ticket key is detected in the branch name or PR title.',
          'Or override with `--jira SRM2-1234` in your generate command.',
        ].join('\n'));
        return;
      }

      stream.markdown([
        '## Save Jira credentials\n',
        '```',
        '@code-review /jira your@email.com ATLASSIAN_API_TOKEN',
        '```',
        '',
        'Generate an API token at: https://id.atlassian.net/manage-profile/security/api-tokens',
        '',
        '**Other commands:**',
        '- `@code-review /jira status` — check stored credentials',
        '- `@code-review /jira clear` — remove credentials',
      ].join('\n'));
      return;
    }

    // ── /history — show all versions for a PR ─────────────────────────────
    if (request.command === 'history') {
      const prKey = request.prompt.trim();
      if (!prKey) {
        stream.markdown('Usage: `@code-review /history <PR_URL_or_branch>`');
        return;
      }
      const workspaceRoot = vscode.workspace.workspaceFolders?.[0]?.uri.fsPath ?? process.env.HOME ?? '/tmp';
      const idx = new IndexStore(path.join(workspaceRoot, 'code-review'));
      await idx.load();
      const versions = idx.getVersions(prKey);
      if (!versions.length) {
        stream.markdown(`No checklist history found for:\n\`${prKey}\``);
        return;
      }
      const rows = versions.map(v =>
        `| v${v.version} | ${v.date} | ❌ ${v.notOkCount} Not Ok / ✅ ${v.okCount} Ok | \`${v.file}\` |`
      ).join('\n');
      stream.markdown([
        `## Version history for PR`,
        `\`${prKey}\`\n`,
        `| Version | Date | Status | File |`,
        `|---------|------|--------|------|`,
        rows,
      ].join('\n'));
      return;
    }

    // ── Process file attachments (drag-and-drop from VS Code) ─────────────
    // Collect all file URIs from request.references
    const attachedFiles: string[] = [];
    let   attachedFileContent = '';

    for (const ref of request.references) {
      if (ref.value instanceof vscode.Uri && ref.value.scheme === 'file') {
        attachedFiles.push(ref.value.fsPath);
      }
    }

    // Read all attached files and aggregate content
    let fileItems: string[] = [];
    if (attachedFiles.length > 0) {
      stream.markdown(`📎 Reading **${attachedFiles.length}** attached file(s)…\n`);
      for (const fp of attachedFiles) {
        try {
          const result = await readAttachedFile(fp);
          attachedFileContent += `\n### File: ${path.basename(fp)}\n${result.rawText}\n`;
          fileItems.push(...result.allItems);

          if (result.allItems.length > 0) {
            stream.markdown(
              `✅ \`${path.basename(fp)}\` → found **${result.prUrls.length}** PR URL(s) + **${result.jiraKeys.length}** Jira key(s)\n`
            );
          } else {
            stream.markdown(`📄 \`${path.basename(fp)}\` → added as context (no PR/Jira items found)\n`);
          }
        } catch (err: any) {
          stream.markdown(`⚠️ Could not read \`${path.basename(fp)}\`: ${err.message}\n`);
        }
      }
      fileItems = [...new Set(fileItems)];
    }

    // ── Route: if files have PR/Jira items → auto-trigger --list mode ──────
    // Inject file items into the prompt as --list items
    const promptWithFileItems = fileItems.length > 0
      ? (request.prompt + ` --list ${fileItems.join(' ')}`).trim()
      : request.prompt;

    // ── Conversational mode — any non-command message ─────────────────────
    if (request.command !== 'generate') {
      const workspaceRoot = vscode.workspace.workspaceFolders?.[0]?.uri.fsPath ?? process.env.HOME ?? '/tmp';

      // If files had actionable items and no explicit command → treat as /generate --list
      if (fileItems.length > 0) {
        stream.markdown(`🔄 Found **${fileItems.length}** item(s) in file(s). Running checklist generation…\n`);
        // Fall through to generate with modified prompt by not returning here
        // We'll handle this by pretending the user ran /generate --list <items>
      } else {
        // Pure conversational question — answer using checklist knowledge + context
        let model: vscode.LanguageModelChat | undefined;
        for (const family of ['claude-sonnet', 'gpt-4o', 'gpt-4', 'default']) {
          const candidates = await vscode.lm.selectChatModels({ family });
          if (candidates.length) { model = candidates[0]; break; }
        }
        if (!model) {
          const all = await vscode.lm.selectChatModels({});
          model = all[0];
        }
        if (!model) {
          stream.markdown('⚠️ No AI model available. Ensure GitHub Copilot Chat is active.');
          return;
        }

        const userQuestion = request.prompt.trim() || 'Tell me about the last code review.';
        await handleConversation(userQuestion, attachedFileContent, workspaceRoot, model, token, stream);
        return;
      }
    }
    // ── Parse arguments ────────────────────────────────────────────────────
    // Use promptWithFileItems when files were attached (injects --list items from files)
    const prompt = promptWithFileItems.trim();
    const tokenMatch = prompt.match(/--token\s+(\S+)/);
    // Use inline --token if given, otherwise load from secure storage
    const inlineToken = tokenMatch ? tokenMatch[1] : null;
    const storedToken = await tokenStore.get();
    const githubToken = inlineToken ?? storedToken ?? null;

    // Branch / base detection: --branch <name> --base <name>
    const branchMatch = prompt.match(/--branch\s+(\S+)/);
    const baseMatch   = prompt.match(/--base\s+(\S+)/);
    const manualBranch = branchMatch ? branchMatch[1] : null;
    const manualBase   = baseMatch   ? baseMatch[1]   : 'main';

    // PR URL (optional when using local git mode)
    const urlMatch = prompt.match(/https?:\/\/github\.com\/[^\s]+\/pull\/\d+/);
    let   prUrl    = urlMatch ? urlMatch[0] : null;  // let — reassigned in --list loop

    // --list: space/newline separated Jira keys and/or PR URLs
    const listMatch = prompt.match(/--list\s+([\s\S]+?)(?=\n--|$)/i);
    const listItems: string[] = listMatch
      ? listMatch[1].trim().split(/[\s,]+/).filter(Boolean)
      : [];

    // --jira: accept one or more space-separated ticket keys
    const jiraMatch    = prompt.match(/--jira\s+((?:[A-Z][A-Z0-9]+-\d+\s*)+)/i);
    const jiraOverrides: string[] = jiraMatch
      ? jiraMatch[1].trim().toUpperCase().split(/\s+/)
      : [];

    // --sprint <number or name> — manual sprint folder override
    const sprintMatch   = prompt.match(/--sprint\s+(\S+)/i);
    const manualSprint  = sprintMatch ? `Sprint-${sprintMatch[1].replace(/^Sprint-?/i, '')}` : null;

    // Strip all flags and URLs — whatever free-form text remains is the user spec.
    // No "Prompt:" keyword needed — just type extra context on a new line after the command.
    const userSpec = prompt
      .replace(/--token\s+\S+/g,  '')
      .replace(/--branch\s+\S+/g, '')
      .replace(/--base\s+\S+/g,   '')
      .replace(/--jira\s+(?:[A-Z][A-Z0-9]+-\d+\s*)+/gi, '')
      .replace(/--sprint\s+\S+/g, '')
      .replace(/--list\s+[\s\S]+?(?=\n--|$)/gi, '')
      .replace(/https?:\/\/github\.com\/\S+/g, '')
      .replace(/Prompt\s*:/gi, '')   // strip "Prompt:" prefix if someone still uses it
      .split('\n')
      .map(l => l.trim())
      .filter(Boolean)
      .join('\n')
      .trim();

    // Pasted comments = lines that look like review comments (not flags, not URLs)
    // Re-use the same cleaned text for backward compatibility
    const pastedComments = userSpec
      ? userSpec.split('\n').filter(l => !l.startsWith('--') && !l.match(/^https?:\/\//))
      : [];

    // If --jira keys given without a PR URL → treat them as --list items
    // so the agent resolves the linked PR automatically from Jira dev-info.
    const effectiveListItems = listItems.length > 0
      ? listItems
      : (!prUrl && !manualBranch && jiraOverrides.length > 0)
        ? jiraOverrides
        : listItems;

    if (!prUrl && !manualBranch && effectiveListItems.length === 0) {
      stream.markdown('⚠️ Please provide a PR URL, `--branch <name>`, `--jira <key>`, or `--list <items>`.\n\n' + HELP_TEXT);
      return;
    }

    if (userSpec) {
      stream.markdown(`📋 Using your specification:\n> ${userSpec.replace(/\n/g, '\n> ')}\n`);
    }

    // ── Resolve --list items to PR URLs ────────────────────────────────────
    let prUrlsToProcess: string[] = [];
    if (effectiveListItems.length > 0) {
      const isJiraFallback = listItems.length === 0 && effectiveListItems === jiraOverrides;
      const label = isJiraFallback
        ? `🔗 No PR URL given — resolving PR from Jira story **${effectiveListItems.join(', ')}**…\n`
        : `📋 **--list mode:** Resolving **${effectiveListItems.length}** item(s)…\n`;
      stream.markdown(label);
      const listJiraCreds = await jiraStore.get();
      const resolved = await resolveListItems(effectiveListItems, listJiraCreds, stream);

      // Deduplicate — multiple Jira stories can link to the same PR.
      // Normalise URLs (strip trailing slashes) before comparing.
      const seen = new Set<string>();
      const duplicates: string[] = [];
      prUrlsToProcess = resolved.filter(url => {
        const key = url.replace(/\/+$/, '').toLowerCase();
        if (seen.has(key)) { duplicates.push(url); return false; }
        seen.add(key);
        return true;
      });

      if (duplicates.length > 0) {
        stream.markdown(
          `🔀 **${duplicates.length}** duplicate PR(s) removed — multiple stories linked to the same PR:\n` +
          duplicates.map(u => `- \`${u}\``).join('\n') + '\n'
        );
      }

      if (prUrlsToProcess.length === 0) {
        stream.markdown('⚠️ No PR URLs could be resolved from the list. Nothing to generate.\n');
        return;
      }
      stream.markdown(`\n✅ **${prUrlsToProcess.length}** unique PR(s) to process. Generating checklists…\n`);
    } else {
      prUrlsToProcess = prUrl ? [prUrl] : [];
    }

    // ── Create shared ContextWriter ─────────────────────────────────────────
    const workspaceRootForCtx = vscode.workspace.workspaceFolders?.[0]?.uri.fsPath ?? process.env.HOME ?? '/tmp';
    const ctxWriter = new ContextWriter(path.join(workspaceRootForCtx, 'code-review'));

    // ── Loop: one checklist per resolved PR URL ─────────────────────────────
    for (let prIdx = 0; prIdx < prUrlsToProcess.length; prIdx++) {
      prUrl = prUrlsToProcess[prIdx];

      if (prUrlsToProcess.length > 1) {
        stream.markdown(`\n---\n### 🔄 PR ${prIdx + 1} of ${prUrlsToProcess.length}: \`${prUrl}\`\n`);
      }

      stream.markdown(`🔍 Starting checklist generation…\n`);

      ctxWriter.add(`Run ${prIdx + 1} Started`, [
        `- PR URL: ${prUrl ?? '(local)'}`,
        `- Branch: ${manualBranch ?? '(from PR)'}`,
        `- Timestamp: ${new Date().toISOString()}`,
        `- User Spec: ${userSpec || '(none)'}`,
        `- Jira overrides: ${jiraOverrides.join(', ') || '(auto-detect)'}`,
        `- Sprint override: ${manualSprint ?? '(auto-detect)'}`,
      ].join('\n'));

    // ── Gather PR data ─────────────────────────────────────────────────────
    let prData: PRData;

    if (prUrl && githubToken) {
      // ── Path A: Full GitHub API (comments + code diff) ──────────────────
      stream.markdown(`📡 Fetching PR data + code diff from GitHub API ${storedToken && !inlineToken ? '*(using saved PAT)*' : ''}…\n`);
      try {
        prData = await fetchFromGitHub(prUrl, githubToken);
        ctxWriter.add('PR Data', [
          `- Title: ${prData.title}`,
          `- Author: ${prData.author}`,
          `- Assignees: ${prData.assignees.join(', ') || 'none'}`,
          `- Reviewers: ${prData.reviewers.join(', ') || 'none'}`,
          `- Branch: ${prData.branch} → ${prData.baseBranch}`,
          `- Files changed: ${prData.files.length}`,
          `- Comments: ${prData.comments.length}`,
          `- Created: ${prData.createdAt ?? 'unknown'}`,
          `- Closed: ${prData.closedAt ?? 'open'}`,
          '',
          '**Changed files:**',
          prData.diffSummary || '(none)',
        ].join('\n'));
        stream.markdown(
          `✅ GitHub: **${prData.comments.length}** comment(s), **${prData.files.length}** changed file(s)\n` +
          `   _(includes review comments, inline code comments + PR conversation thread)_\n` +
          (prData.assignees.length ? `👤 Assignee(s): **${prData.assignees.join(', ')}**\n` : '') +
          (prData.reviewers.length ? `👥 Reviewer(s): **${prData.reviewers.join(', ')}**\n` : '') +
          `\n**Changed files:**\n\`\`\`\n${prData.diffSummary}\n\`\`\`\n`
        );
      } catch (err: any) {
        ctxWriter.addError('GitHub API Failed', err, {
          'PR URL': prUrl ?? '',
          'PAT set': githubToken ? 'yes' : 'no',
          'Hint': 'Run @code-review /token to update PAT, or check VPN/network',
        });
        stream.markdown(`❌ GitHub API failed: ${err.message}\n`);
        stream.reference(ctxWriter.uri);
        return;
      }
      if (pastedComments.length) { prData.comments.push(...pastedComments); }

    } else if (manualBranch) {
      // ── Path B: Local git repo diff ─────────────────────────────────────
      const repoDir = vscode.workspace.workspaceFolders?.[0]?.uri.fsPath;
      if (!repoDir) {
        stream.markdown('❌ No workspace folder open. Please open the repository in VS Code first.');
        return;
      }

      stream.markdown(`📂 Reading local git diff: \`${manualBranch}\` → \`${manualBase}\`…\n`);
      let localFiles: FilePatch[] = [];
      try {
        localFiles = await fetchLocalDiff(repoDir, manualBranch, manualBase);
        stream.markdown(`✅ Local diff: **${localFiles.length}** changed file(s)\n`);
      } catch (err: any) {
        stream.markdown(`⚠️ Git diff failed: ${err.message}. Continuing with comments only…\n`);
      }

      prData = {
        url:         prUrl ?? `(local) ${manualBranch}`,
        title:       `Branch review: ${manualBranch}`,
        body:        '',
        author:      '',
        assignees:   [],
        reviewers:   [],
        branch:      manualBranch,
        baseBranch:  manualBase,
        createdAt:   null,
        closedAt:    null,
        comments:    pastedComments,
        files:       localFiles,
        diffSummary: localFiles.map(f => `  ${f.filename}  (+${f.additions}/-${f.deletions})`).join('\n'),
      };

    } else {
      // ── Path C: PR URL only, manual comments pasted ─────────────────────
      if (!pastedComments.length) {
        stream.markdown([
          '📝 **Please paste PR review comments below the URL:**\n',
          '```',
          '@code-review /generate https://github.com/org/repo/pull/42',
          '1. Missing error handling in UserController#update',
          '2. N+1 query in app/views/users/index line 34',
          '```',
          '\n**For code analysis (recommended), add:**',
          '- `--token <GITHUB_PAT>` — auto-fetch code diff from GitHub',
          '- `--branch <name> --base main` — use local git diff',
        ].join('\n'));
        return;
      }

      prData = {
        url: prUrl!,
        title: `PR ${prUrl}`,
        body: '',
        author: '',
        assignees: [],
        reviewers: [],
        branch: '',
        baseBranch: 'main',
        createdAt: null,
        closedAt:  null,
        comments: pastedComments,
        files: [],
        diffSummary: '',
      };
    }

    // ── Jira key detection + stories fetch ────────────────────────────────
    let jiraStories: JiraStory[] = [];

    // Always detect keys — branch name, PR title, and PR description body.
    // Keys are shown in C4 even when no Jira credentials are configured.
    const detectedKeys: string[] = jiraOverrides.length > 0
      ? jiraOverrides
      : detectJiraKeys(prData);

    if (detectedKeys.length > 0) {
      stream.markdown(`🔍 Detected Jira keys: **${detectedKeys.join(', ')}**\n`);
    }

    const jiraCreds = await jiraStore.get();

    if (jiraCreds && detectedKeys.length > 0) {
      stream.markdown(`🎫 Fetching ${detectedKeys.length} Jira stor${detectedKeys.length > 1 ? 'ies' : 'y'}: **${detectedKeys.join(', ')}**…\n`);
      jiraStories = await fetchJiraStories(detectedKeys, jiraCreds, stream);
      jiraStories.forEach(s => {
        stream.markdown(
          `✅ **[${s.key}]** ${s.summary}  _(${s.type} · ${s.status} · 📅 ${s.sprint})_` +
          (s.acceptanceCriteria ? '  📋 AC found' : '') + '\n'
        );
      });
    } else if (!jiraCreds && detectedKeys.length > 0) {
      stream.markdown(`ℹ️ Jira keys detected but no credentials saved — keys will be listed in the Excel.\nRun \`@code-review /jira your@email.com TOKEN\` to enable full story fetch.\n`);
    } else if (detectedKeys.length === 0) {
      stream.markdown(`ℹ️ No Jira ticket detected in branch, title, or PR description.\nUse \`--jira SRM2-XXXX SRM2-YYYY\` to specify manually.\n`);
    }

    // ── AI analysis ────────────────────────────────────────────────────────
    stream.markdown(`\n🤖 Analysing code diff + comments against all 25 checklist rows…\n`);

    let findings: RowFinding[] = [];
    try {
      // Try preferred models in order; pick the first available one
      let model: vscode.LanguageModelChat | undefined;
      for (const family of ['claude-sonnet', 'gpt-4o', 'gpt-4', 'copilot-gpt-4o', 'default']) {
        const candidates = await vscode.lm.selectChatModels({ family });
        if (candidates.length) { model = candidates[0]; break; }
      }
      if (!model) {
        // Last resort: any available model
        const all = await vscode.lm.selectChatModels({});
        model = all[0];
      }
      if (!model) {
        throw new Error('No language model available. Ensure GitHub Copilot Chat is installed and signed in.');
      }

      stream.markdown(`🧠 Using model: \`${model.name ?? model.id}\`\n`);

      findings = await analyseWithAI(prData, model, token, stream, userSpec, jiraStories);

      // Retry once if we got empty findings (transient LLM issue)
      if (findings.length === 0) {
        stream.markdown(`🔄 Got empty findings, retrying once…\n`);
        findings = await analyseWithAI(prData, model, token, stream, userSpec, jiraStories);
      }

      const fixedCount = findings.filter(f => f.authorStatus === 'Yes' && f.reviewerStatus === 'Ok').length;
      const notOk      = findings.filter(f => f.reviewerStatus === 'Not Ok').length;
      const ok         = findings.filter(f => f.reviewerStatus === 'Ok').length;
      const na         = findings.filter(f => f.reviewerStatus === 'NA').length;

      stream.markdown(
        `\n✅ Analysis complete: ` +
        `❌ **${notOk}** Not Ok · ✅ **${ok}** Ok · **${fixedCount}** addressed · — **${na}** NA\n`
      );

      // ── Write rich code-fix context ──────────────────────────────────────
      // For each Not Ok finding, include the actual diff hunk from the relevant
      // file so Copilot Chat can suggest specific code fixes — not just describe issues.
      const notOkFindings = findings.filter(f => f.reviewerStatus === 'Not Ok');

      ctxWriter.add('AI Findings Summary', [
        `- Model: ${model.name ?? model.id}`,
        `- Not Ok: ${notOk}  |  Ok: ${ok}  |  NA: ${na}  |  Addressed: ${fixedCount}`,
      ].join('\n'));

      if (notOkFindings.length > 0) {
        // Build one section per Not Ok finding with its diff hunk embedded
        const fixSections = notOkFindings.map(f => {
          const row = CHECKLIST_ROWS.find(r => r.id === f.rowId);

          // Match relevant file from diff using the finding text (filename likely mentioned)
          const findingLower = (f.finding ?? '').toLowerCase();
          const relevantFile = prData.files.find(file =>
            findingLower.includes(file.filename.toLowerCase().split('/').pop()!) ||
            prData.files.length === 1   // if only one file changed, use it
          ) ?? prData.files[0];

          const diffBlock = relevantFile
            ? `**Relevant diff (\`${relevantFile.filename}\`):**\n\`\`\`diff\n${relevantFile.patch.split('\n').slice(0, 40).join('\n')}\n\`\`\``
            : '';

          return [
            `### ❌ Row ${f.rowId} — ${row?.reviewArea}: ${row?.checkCategory}`,
            '',
            `**Rule:** ${row?.description}`,
            '',
            `**Finding:** ${f.finding}`,
            `**Defect type:** ${f.defectType}`,
            `**Suggested action:** ${f.remarks || 'Fix per the rule above'}`,
            '',
            diffBlock,
          ].filter(Boolean).join('\n');
        });

        ctxWriter.add(
          `Code Fix Tasks (${notOkFindings.length} issue${notOkFindings.length > 1 ? 's' : ''} to fix)`,
          [
            `> These are the issues found in PR: ${prData.url}`,
            `> Ask Copilot: _"How do I fix issue #1?"_ or _"Show me the fix for Row ${notOkFindings[0]?.rowId}"_`,
            '',
            ...fixSections,
          ].join('\n')
        );
      }

      // ── Review summary ────────────────────────────────────────────────────
      if (findings.length > 0) {
        stream.markdown(`\n📝 Generating review summary…\n`);
        try {
          const summary = await generateReviewSummary(prData, findings, userSpec, jiraStories, model, token);
          ctxWriter.add('Review Summary', summary);
          stream.markdown(`\n---\n## 🔍 Review Summary\n\n${summary}\n\n---\n`);
        } catch {
          // Non-fatal
        }
      }
    } catch (err: any) {
      ctxWriter.addError('AI Analysis Failed', err, { 'Model': 'see above', 'Hint': 'Try again — the model occasionally produces malformed output' });
      stream.markdown(`\n⚠️ AI analysis failed: **${err.message}**\nGenerating Excel with blank findings — you can fill them manually.\n`);
    }

    // ── Generate Excel ─────────────────────────────────────────────────────
    stream.markdown(`\n📊 Filling Excel template…\n`);

    const workspaceRoot = vscode.workspace.workspaceFolders?.[0]?.uri.fsPath ?? process.env.HOME ?? '/tmp';
    const templatePath  = path.join(context.extensionUri.fsPath, 'templates', 'checklist-template.xlsx');

    // Folder structure: <workspace>/code-review/YYYY-MM-DD/Sprint-N/
    const dateStr       = new Date().toISOString().slice(0, 10);   // YYYY-MM-DD
    const codeReviewDir = path.join(workspaceRoot, 'code-review');

    // Determine sprint folder — majority vote across all fetched stories:
    // Each story votes with its sprint. Most common sprint wins.
    // This handles: 3 stories in Sprint-42, 1 in Sprint-43 → Sprint-42 folder.
    let sprintFolder = manualSprint ?? 'No-Sprint';

    if (!manualSprint && jiraStories.length > 0) {
      // Count votes per sprint (exclude No-Sprint from voting)
      const votes = new Map<string, number>();
      jiraStories.forEach(s => {
        if (s.sprint !== 'No-Sprint') {
          votes.set(s.sprint, (votes.get(s.sprint) ?? 0) + 1);
        }
      });

      if (votes.size > 0) {
        // Pick the sprint with the most votes; tie → first story's sprint wins
        sprintFolder = [...votes.entries()]
          .sort((a, b) => b[1] - a[1])[0][0];

        if (votes.size > 1) {
          const breakdown = [...votes.entries()]
            .map(([sp, n]) => `${sp}(${n})`)
            .join(', ');
          stream.markdown(`📊 Sprint vote: ${breakdown} → using **${sprintFolder}**\n`);
        }
      }
    }

    const dayDir = path.join(codeReviewDir, dateStr, sprintFolder);
    fs.mkdirSync(dayDir, { recursive: true });

    // Index — look up existing versions for this PR
    const indexStore = new IndexStore(codeReviewDir);
    await indexStore.load();
    const prKey          = prData.url || prData.branch;
    const prevVersions   = indexStore.getVersions(prKey);
    const version        = indexStore.nextVersion(prKey);

    if (prevVersions.length > 0) {
      const last = prevVersions[prevVersions.length - 1];
      stream.markdown(
        `📂 Found **${prevVersions.length}** previous version(s) for this PR.\n` +
        `   Last: **v${last.version}** on ${last.date} (${last.notOkCount} Not Ok)\n` +
        `   Generating **v${version}** → \`${sprintFolder}\`\n`
      );
    } else {
      stream.markdown(`📁 Saving to: \`code-review/${dateStr}/${sprintFolder}/\`\n`);
    }

    // File name: CodeReview_PR-<number>_v<N>.xlsx
    const parsed    = parsePRUrl(prData.url);
    const prLabel   = parsed ? `PR-${parsed.number}` : (prData.branch || 'review').replace(/[^a-zA-Z0-9-_]/g, '_');
    const filename  = `CodeReview_${prLabel}_v${version}.xlsx`;
    const outputPath = path.join(dayDir, filename);

    const notOkCount = findings.filter(f => f.reviewerStatus === 'Not Ok').length;
    const okCount    = findings.filter(f => f.reviewerStatus === 'Ok').length;
    const naCount    = findings.filter(f => f.reviewerStatus === 'NA').length;

    try {
      await generateExcel(prData, findings, templatePath, outputPath, version, prevVersions, jiraStories, detectedKeys);
    } catch (err: any) {
      stream.markdown(`❌ Excel write failed: ${err.message}`);
      return;
    }

    // Register in index
    indexStore.register(prKey, {
      version,
      file:       outputPath,
      date:       dateStr,
      notOkCount,
      okCount,
      branch:     prData.branch,
    });

    const fileUri    = vscode.Uri.file(outputPath);
    const folderUri  = vscode.Uri.file(dayDir);

    const rrNote = notOkCount > 0
      ? `\n⚠️ **${notOkCount} Not Ok** row(s) found — a **"Re-review Required"** sheet has been added listing them.`
      : `\n✅ No Not Ok rows — clean review!`;

    stream.markdown([
      `\n---`,
      `## ✅ Checklist Excel ready! (v${version})`,
      `📁 \`${outputPath}\``,
      rrNote,
      ``,
      `| Reviewer Status | Count |`,
      `|----------------|-------|`,
      `| ❌ Not Ok | ${notOkCount} |`,
      `| ✅ Ok     | ${okCount}    |`,
      `| — NA      | ${naCount}    |`,
      ``,
      `**What to do next:**`,
      `- Fill in **Reviewer(s)** name (C7) if not auto-filled`,
      `- Review AI-filled findings (col G) and adjust if needed`,
      `- After re-review, run \`@code-review /generate\` again — it will create **v${version + 1}**`,
      `- \`@code-review /history ${prData.url || prData.branch}\` to see all versions`,
    ].join('\n'));

    // Write final context and surface it in chat via stream.reference()
    // Copilot Chat reads referenced files automatically — any follow-up question
    // in this chat session will have full run context without the user doing anything.
    ctxWriter.add('Result', [
      `- Excel: ${outputPath}`,
      `- Sprint folder: ${sprintFolder}`,
      `- Version: v${version}`,
      `- Not Ok: ${findings.filter(f => f.reviewerStatus === 'Not Ok').length}`,
      `- Ok: ${findings.filter(f => f.reviewerStatus === 'Ok').length}`,
      `- NA: ${findings.filter(f => f.reviewerStatus === 'NA').length}`,
    ].join('\n'));

    stream.button({ command: 'vscode.open',          title: '📄 Open Excel',          arguments: [fileUri]   });
    stream.button({ command: 'revealFileInOS',        title: '📂 Show in Folder',      arguments: [folderUri] });

    } // ── end of per-PR loop ─────────────────────────────────────────────

    // If multiple PRs processed, show a final summary
    if (prUrlsToProcess.length > 1) {
      stream.markdown([
        `\n---`,
        `## ✅ All ${prUrlsToProcess.length} PRs processed!`,
        ...prUrlsToProcess.map((u, i) => `${i + 1}. \`${u}\``),
      ].join('\n'));
    }

    // Attach context file — Copilot Chat reads this in follow-up questions automatically
    stream.reference(ctxWriter.uri);
  });

  participant.iconPath = vscode.Uri.joinPath(context.extensionUri, 'media', 'icon.png');
  context.subscriptions.push(participant);
}

export function deactivate() {}

// ─────────────────────────────────────────────────────────────────────────────
// Help text
// ─────────────────────────────────────────────────────────────────────────────
const HELP_TEXT = `## Code Review Checklist Agent 📋

Analyses a GitHub PR's **code diff + review comments** using Copilot AI and fills your
PSSM2.0 Excel checklist template — ready to share with stakeholders in seconds.

---

### 🔑 First time only — save your GitHub PAT
\`\`\`
@code-review /token ghp_xxxxxxxxxxxxxxxxxxxx
\`\`\`
Stored in VS Code's encrypted OS keychain. Never typed again.
*(Generate at: https://github.com/settings/tokens — scope: **repo**)*

---

### /generate — 3 ways to use

**Option 1 · GitHub API (recommended)**
\`\`\`
@code-review /generate https://github.com/org/repo/pull/42
\`\`\`
Auto-fetches: code diff · PR comments · assignees · reviewers · dates

**Option 2 · Local git diff**
\`\`\`
@code-review /generate --branch feature/my-branch --base main
\`\`\`
Reads diff from your open repo. Paste review comments below (optional).

**Option 3 · Manual comments**
\`\`\`
@code-review /generate https://github.com/org/repo/pull/42
Missing error handling in UserController
N+1 query in users/index line 34
\`\`\`

---

### 💡 Extra context — just type it after the command
Add any extra instructions on a new line — no keyword needed:
\`\`\`
@code-review /generate https://github.com/org/repo/pull/42
Backend-only Rails PR — no React changes.
Performance is top priority, 10k+ cases in prod.
i18n skipped this sprint — mark those rows NA.
Team agreed to defer test coverage to next PR.
\`\`\`

---

### What gets auto-filled in the Excel

| Cell | Value |
|------|-------|
| C4 | PR URL |
| C5 | Version (v1, v2…) |
| C6 | PR assignee(s) |
| C7 | PR reviewer(s) |
| C8 | PR raised date |
| C9 | Merged/closed date (blank if open) |
| E–I rows 11–35 | AI findings for all 25 checklist rows |
| H3–H9, J4–J9 | Defect + summary counts (pre-computed) |

**Extra sheets added when needed:**
- **Re-review Required** — only the Not Ok rows, with suggested actions
- **Version History** — all runs for this PR with dates and counts

---

### Saved to
\`\`\`
<workspace>/code-review/YYYY-MM-DD/CodeReview_PR-<number>_v<N>.xlsx
\`\`\`

---

### All commands

| Command | Description |
|---------|-------------|
| \`/generate <PR_URL>\` | Generate filled checklist from GitHub PR |
| \`/generate --branch <name>\` | Use local git diff instead |
| \`/token <PAT>\` | Save GitHub PAT to OS keychain |
| \`/token status\` | Check if a PAT is stored |
| \`/token clear\` | Remove stored PAT |
| \`/history <PR_URL>\` | Show all checklist versions for a PR |
| \`/help\` | Show this message |

---

**Row colours:** 🔴 Not Ok &nbsp;·&nbsp; 🟢 Ok &nbsp;·&nbsp; 🟡 NA`;
