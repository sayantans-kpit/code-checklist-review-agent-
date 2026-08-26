/**
 * Pure utility functions extracted from extension.ts for testability.
 * No vscode imports — safe to test in any Node environment.
 */

/** Parse a GitHub PR URL → owner/repo/number */
export function parsePRUrl(url: string): { owner: string; repo: string; number: number } | null {
  const m = url.match(/github\.com\/([^/]+)\/([^/]+)\/pull\/(\d+)/);
  return m ? { owner: m[1], repo: m[2], number: parseInt(m[3], 10) } : null;
}

/** Format ISO date string → "dd-Mmm-yyyy". Returns null if input is null. */
export function fmtDate(iso: string | null): string | null {
  if (!iso) { return null; }
  return new Date(iso).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
}

/** Detect all Jira ticket keys from a set of text sources */
export function detectJiraKeysFromText(sources: string[]): string[] {
  const JIRA_PATTERN  = /\b([A-Z][A-Z0-9]+-\d+)\b/g;
  const JIRA_URL_PAT  = /atlassian\.net\/browse\/([A-Z][A-Z0-9]+-\d+)/g;
  const found = new Set<string>();
  for (const text of sources) {
    if (!text) { continue; }
    for (const m of text.matchAll(JIRA_PATTERN))  { found.add(m[1]); }
    for (const m of text.matchAll(JIRA_URL_PAT))  { found.add(m[1]); }
  }
  return Array.from(found);
}

/** Robustly extract a JSON array from raw LLM response text */
export function extractJsonArray(raw: string): string {
  const cleaned = raw.replace(/```(?:json)?/gi, '').trim();
  const start = cleaned.indexOf('[');
  const end   = cleaned.lastIndexOf(']');
  if (start !== -1 && end !== -1 && end > start) {
    return cleaned.slice(start, end + 1);
  }
  return cleaned;
}

/** Parse sprint name → folder-safe string e.g. "PSSM 2.0 Sprint 42" → "Sprint-42" */
export function parseSprintName(name: string): string {
  if (!name) { return 'No-Sprint'; }
  const num = name.match(/(\d+)\s*$/)?.[1] ?? name.match(/(\d+)/)?.[1];
  return num ? `Sprint-${num}` : name.replace(/\s+/g, '-');
}

/** Extract acceptance criteria from Jira description text */
export function extractAC(fullText: string): string {
  const acMatch = fullText.match(
    /acceptance criteria[\s\S]*?(?=\n##|\n[A-Z][^\n]{0,30}\n[-=]{3,}|$)/i
  );
  return acMatch ? acMatch[0].trim() : '';
}

/** Deduplicate PR URLs — strips trailing slashes, case-insensitive */
export function deduplicatePRUrls(urls: string[]): { unique: string[]; duplicates: string[] } {
  const seen = new Set<string>();
  const unique: string[]     = [];
  const duplicates: string[] = [];
  for (const url of urls) {
    const key = url.replace(/\/+$/, '').toLowerCase();
    if (seen.has(key)) { duplicates.push(url); }
    else { seen.add(key); unique.push(url); }
  }
  return { unique, duplicates };
}

/** Truncate a patch to maxLines — preserves diff header */
export function truncatePatch(patch: string, maxLines = 120): string {
  const lines = patch.split('\n');
  if (lines.length <= maxLines) { return patch; }
  return lines.slice(0, maxLines).join('\n') + `\n… [truncated ${lines.length - maxLines} lines]`;
}
