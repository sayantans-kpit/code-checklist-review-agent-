import {
  parsePRUrl,
  fmtDate,
  detectJiraKeysFromText,
  extractJsonArray,
  parseSprintName,
  extractAC,
  deduplicatePRUrls,
  truncatePatch,
} from '../utils';

// ─────────────────────────────────────────────────────────────────────────────
// parsePRUrl
// ─────────────────────────────────────────────────────────────────────────────
describe('parsePRUrl', () => {
  it('parses a standard GitHub PR URL', () => {
    const result = parsePRUrl('https://github.com/Decisiv/pricing/pull/20320');
    expect(result).toEqual({ owner: 'Decisiv', repo: 'pricing', number: 20320 });
  });

  it('parses a PR URL with trailing slash', () => {
    const result = parsePRUrl('https://github.com/org/repo/pull/42/');
    expect(result).toEqual({ owner: 'org', repo: 'repo', number: 42 });
  });

  it('returns null for a non-PR URL', () => {
    expect(parsePRUrl('https://github.com/org/repo')).toBeNull();
    expect(parsePRUrl('https://jira.atlassian.net/browse/SRM2-1')).toBeNull();
    expect(parsePRUrl('')).toBeNull();
  });

  it('handles hyphenated repo names', () => {
    const result = parsePRUrl('https://github.com/sayantans-kpit/code-review-agent/pull/5');
    expect(result).toEqual({ owner: 'sayantans-kpit', repo: 'code-review-agent', number: 5 });
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// fmtDate
// ─────────────────────────────────────────────────────────────────────────────
describe('fmtDate', () => {
  it('returns null for null input', () => {
    expect(fmtDate(null)).toBeNull();
  });

  it('formats an ISO date string to dd-Mmm-yyyy', () => {
    const result = fmtDate('2026-08-18T10:30:00Z');
    expect(result).toMatch(/\d{2} \w{3} \d{4}/);   // "18 Aug 2026"
    expect(result).toContain('2026');
  });

  it('handles ISO date-only string', () => {
    const result = fmtDate('2026-01-01');
    expect(result).toContain('2026');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// detectJiraKeysFromText
// ─────────────────────────────────────────────────────────────────────────────
describe('detectJiraKeysFromText', () => {
  it('detects a key from branch name', () => {
    const keys = detectJiraKeysFromText(['feature/SRM2-2072-extend-case-search']);
    expect(keys).toContain('SRM2-2072');
  });

  it('detects multiple keys from PR title', () => {
    const keys = detectJiraKeysFromText(['Fix SRM2-2072 and SRM2-2073: improve search']);
    expect(keys).toContain('SRM2-2072');
    expect(keys).toContain('SRM2-2073');
  });

  it('detects keys from Jira URL in PR body', () => {
    const body = 'See https://decisiv.atlassian.net/browse/SRM2-1623 for context.';
    const keys = detectJiraKeysFromText([body]);
    expect(keys).toContain('SRM2-1623');
  });

  it('deduplicates keys appearing in multiple sources', () => {
    const keys = detectJiraKeysFromText(['feature/SRM2-2072', 'Fixes SRM2-2072']);
    expect(keys.filter(k => k === 'SRM2-2072')).toHaveLength(1);
  });

  it('returns empty array when no keys found', () => {
    const keys = detectJiraKeysFromText(['fix: update button padding', '']);
    expect(keys).toHaveLength(0);
  });

  it('does not match version numbers as Jira keys', () => {
    // "v2.0" should not match — requires uppercase prefix
    const keys = detectJiraKeysFromText(['version 2.0 release']);
    expect(keys).toHaveLength(0);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// extractJsonArray
// ─────────────────────────────────────────────────────────────────────────────
describe('extractJsonArray', () => {
  it('extracts clean JSON array', () => {
    const raw = '[{"rowId":1,"finding":"ok"}]';
    expect(extractJsonArray(raw)).toBe(raw);
  });

  it('strips markdown code fence', () => {
    const raw = '```json\n[{"rowId":1}]\n```';
    const result = extractJsonArray(raw);
    expect(result).toBe('[{"rowId":1}]');
  });

  it('extracts array from prose-wrapped response', () => {
    const raw = 'Here is the analysis:\n[{"rowId":1,"status":"Ok"}]\nDone.';
    const result = extractJsonArray(raw);
    expect(result).toBe('[{"rowId":1,"status":"Ok"}]');
  });

  it('handles array with nested objects', () => {
    const raw = '```\n[{"a":{"b":1}},{"c":[1,2]}]\n```';
    const result = JSON.parse(extractJsonArray(raw));
    expect(result).toHaveLength(2);
  });

  it('returns original string when no array brackets found', () => {
    const raw = 'no json here';
    expect(extractJsonArray(raw)).toBe('no json here');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// parseSprintName
// ─────────────────────────────────────────────────────────────────────────────
describe('parseSprintName', () => {
  it('extracts number from "Sprint 42"', () => {
    expect(parseSprintName('Sprint 42')).toBe('Sprint-42');
  });

  it('extracts number from "PSSM 2.0 Sprint 42"', () => {
    expect(parseSprintName('PSSM 2.0 Sprint 42')).toBe('Sprint-42');
  });

  it('extracts number from "S42"', () => {
    expect(parseSprintName('S42')).toBe('Sprint-42');
  });

  it('returns No-Sprint for empty string', () => {
    expect(parseSprintName('')).toBe('No-Sprint');
  });

  it('replaces spaces with hyphens when no number found', () => {
    expect(parseSprintName('Alpha Sprint')).toBe('Alpha-Sprint');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// extractAC
// ─────────────────────────────────────────────────────────────────────────────
describe('extractAC', () => {
  it('extracts acceptance criteria section', () => {
    const text = `
## Description
This PR adds case search.

## Acceptance Criteria
1. User can search by RO number
2. Results show within 2 seconds

## Notes
Other notes here.
    `.trim();
    const ac = extractAC(text);
    expect(ac.toLowerCase()).toContain('acceptance criteria');
    expect(ac).toContain('RO number');
  });

  it('returns empty string when no AC section found', () => {
    expect(extractAC('Just a description with no criteria.')).toBe('');
  });

  it('is case-insensitive for the heading', () => {
    const text = 'acceptance criteria\n- Must work\n- Must be fast';
    expect(extractAC(text)).toContain('Must work');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// deduplicatePRUrls
// ─────────────────────────────────────────────────────────────────────────────
describe('deduplicatePRUrls', () => {
  it('returns unique URLs unchanged', () => {
    const { unique, duplicates } = deduplicatePRUrls([
      'https://github.com/org/repo/pull/1',
      'https://github.com/org/repo/pull/2',
    ]);
    expect(unique).toHaveLength(2);
    expect(duplicates).toHaveLength(0);
  });

  it('removes exact duplicate URLs', () => {
    const { unique, duplicates } = deduplicatePRUrls([
      'https://github.com/org/repo/pull/1',
      'https://github.com/org/repo/pull/1',
    ]);
    expect(unique).toHaveLength(1);
    expect(duplicates).toHaveLength(1);
  });

  it('removes URL duplicates with trailing slash', () => {
    const { unique, duplicates } = deduplicatePRUrls([
      'https://github.com/org/repo/pull/1',
      'https://github.com/org/repo/pull/1/',
    ]);
    expect(unique).toHaveLength(1);
    expect(duplicates).toHaveLength(1);
  });

  it('is case-insensitive', () => {
    const { unique } = deduplicatePRUrls([
      'https://github.com/Org/Repo/pull/1',
      'https://github.com/org/repo/pull/1',
    ]);
    expect(unique).toHaveLength(1);
  });

  it('handles multiple stories pointing to same PR (real-world scenario)', () => {
    // SRM2-A and SRM2-B both link to PR #20715
    const urls = [
      'https://github.com/Decisiv/pricing/pull/20715',   // from SRM2-A
      'https://github.com/Decisiv/pricing/pull/20715',   // from SRM2-B
      'https://github.com/Decisiv/pricing/pull/20716',   // from SRM2-C (different PR)
    ];
    const { unique, duplicates } = deduplicatePRUrls(urls);
    expect(unique).toHaveLength(2);
    expect(duplicates).toHaveLength(1);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// truncatePatch
// ─────────────────────────────────────────────────────────────────────────────
describe('truncatePatch', () => {
  it('returns patch unchanged when under limit', () => {
    const patch = Array(10).fill('+line').join('\n');
    expect(truncatePatch(patch, 120)).toBe(patch);
  });

  it('truncates and appends truncation message', () => {
    const lines = Array(150).fill('+line');
    const patch  = lines.join('\n');
    const result = truncatePatch(patch, 120);
    expect(result).toContain('truncated 30 lines');
    expect(result.split('\n').length).toBeLessThan(130);
  });

  it('preserves diff headers in truncated output', () => {
    const lines = ['@@ -1,3 +1,3 @@', ...Array(150).fill('+line')];
    const result = truncatePatch(lines.join('\n'), 120);
    expect(result).toContain('@@ -1,3 +1,3 @@');
  });
});
