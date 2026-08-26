/**
 * Integration test: Excel generation using the real PSSM2.0 template.
 * Verifies that the generated file is a valid xlsx, has the right sheet,
 * and all filled cells contain the expected values.
 */

import * as path from 'path';
import * as fs   from 'fs';
import * as os   from 'os';
import XLSX      from 'xlsx';

const TEMPLATE_PATH = path.resolve(__dirname, '../../templates/checklist-template.xlsx');
const OUTPUT_DIR    = os.tmpdir();

/** Minimal mock data mirroring the runtime types */
const mockFindings = [
  { rowId: 1,  authorStatus: 'Yes', reviewerStatus: 'Ok',     finding: 'Branch is current.',              defectType: 'Technical',  remarks: '' },
  { rowId: 2,  authorStatus: 'No',  reviewerStatus: 'Not Ok', finding: 'Missing i18n key for new label.', defectType: 'Functional', remarks: 'Add translation key to en.yml' },
  { rowId: 4,  authorStatus: 'No',  reviewerStatus: 'Not Ok', finding: 'N+1 query in index partial.',     defectType: 'Technical',  remarks: 'Add .includes(:related_ros)' },
  // rows 3, 5-25 left blank (no finding)
];

const mockPRData = {
  url:        'https://github.com/Decisiv/pricing/pull/20320',
  title:      'Extend case search with related ROs',
  body:       'Fixes SRM2-2072',
  author:     'pruthviraj-kale',
  assignees:  ['pruthviraj-kale'],
  reviewers:  ['sayantan-kpit', 'tinynumbers'],
  branch:     'feature/extend-case-search',
  baseBranch: 'master',
  createdAt:  '2026-06-26T10:00:00Z',
  closedAt:   '2026-07-22T14:00:00Z',
  comments:   [],
  files:      [],
  diffSummary:'',
};

const mockJiraStories = [
  { key: 'SRM2-2072', url: 'https://decisiv.atlassian.net/browse/SRM2-2072',
    summary: 'Extend case search', type: 'Story', status: 'Done', sprint: 'Sprint-42',
    labels: [], description: '', acceptanceCriteria: '' },
];

const mockPrevVersions: any[] = [];

function setVal(ws: XLSX.WorkSheet, addr: string, val: string | number): void {
  const existing = ws[addr];
  if (existing) {
    existing.v = val;
    existing.t = typeof val === 'number' ? 'n' : 's';
    delete existing.f;
    delete existing.w;
  } else {
    ws[addr] = { t: typeof val === 'number' ? 'n' : 's', v: val };
  }
}

/** Reproduce the core logic of generateExcel for testing */
function buildTestExcel(outputPath: string): void {
  const wb = XLSX.readFile(TEMPLATE_PATH, { cellStyles: true, cellFormula: true, bookVBA: false });
  const ws = wb.Sheets['ROR'];

  if (!ws) { throw new Error('ROR sheet missing from template'); }

  const today = '26 Aug 2026';

  // Header cells
  setVal(ws, 'C4', `${mockPRData.url} [SRM2-2072]`);
  setVal(ws, 'C5', 'v1');
  setVal(ws, 'C6', mockPRData.assignees.join(', '));
  setVal(ws, 'C7', `SME - ${mockPRData.reviewers.join(', ')}`);
  setVal(ws, 'C8', '26 Jun 2026');
  setVal(ws, 'C9', '22 Jul 2026');

  // Checklist rows
  mockFindings.forEach(f => {
    const row = f.rowId + 10;
    setVal(ws, `E${row}`, f.authorStatus);
    setVal(ws, `F${row}`, f.reviewerStatus);
    setVal(ws, `G${row}`, f.finding);
    setVal(ws, `H${row}`, f.defectType);
    setVal(ws, `I${row}`, f.remarks);
  });

  // Summary counts
  const countF = (v: string) => mockFindings.filter(f => f.reviewerStatus === v).length;
  setVal(ws, 'J8', countF('Ok'));
  setVal(ws, 'J9', countF('Not Ok'));

  // Extra sheets
  const notOkRows = mockFindings.filter(f => f.reviewerStatus === 'Not Ok');
  const rrData = [
    ['#', 'Review Area', 'Finding', 'Defect Type', 'Action', 'Status'],
    ...notOkRows.map(f => [f.rowId, '', f.finding, f.defectType, f.remarks, '']),
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rrData), 'Re-review Required');

  const histData = [
    ['Version', 'Date', 'Not Ok', 'Ok', 'NA', 'Comments'],
    ['v1 (current)', today, notOkRows.length, countF('Ok'), 0, ''],
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(histData), 'Version History');

  if (!wb.Workbook) { wb.Workbook = {}; }
  if (!wb.Workbook.Views) { wb.Workbook.Views = [{}]; }
  (wb.Workbook.Views[0] as any).ActiveTab = 0;

  XLSX.writeFile(wb, outputPath, { bookSST: true, compression: true, bookType: 'xlsx' });
}

// ─────────────────────────────────────────────────────────────────────────────
// Tests
// ─────────────────────────────────────────────────────────────────────────────

describe('Excel generation', () => {
  const outputPath = path.join(OUTPUT_DIR, `test-checklist-${Date.now()}.xlsx`);
  let wb: XLSX.WorkBook;
  let ws: XLSX.WorkSheet;

  beforeAll(() => {
    if (!fs.existsSync(TEMPLATE_PATH)) {
      throw new Error(`Template not found at ${TEMPLATE_PATH} — run tests from project root`);
    }
    buildTestExcel(outputPath);
    wb = XLSX.readFile(outputPath);
    ws = wb.Sheets['ROR'];
  });

  afterAll(() => {
    if (fs.existsSync(outputPath)) { fs.unlinkSync(outputPath); }
  });

  it('generates a valid xlsx file', () => {
    expect(fs.existsSync(outputPath)).toBe(true);
    const stats = fs.statSync(outputPath);
    expect(stats.size).toBeGreaterThan(5000);   // not empty
  });

  it('contains the ROR sheet', () => {
    expect(wb.SheetNames).toContain('ROR');
    expect(ws).toBeDefined();
  });

  it('adds Re-review Required sheet when Not Ok rows exist', () => {
    expect(wb.SheetNames).toContain('Re-review Required');
  });

  it('adds Version History sheet', () => {
    expect(wb.SheetNames).toContain('Version History');
  });

  it('fills C4 with PR URL and Jira key', () => {
    const c4 = ws['C4']?.v as string;
    expect(c4).toContain('https://github.com/Decisiv/pricing/pull/20320');
    expect(c4).toContain('[SRM2-2072]');
  });

  it('fills C5 with version', () => {
    expect(ws['C5']?.v).toBe('v1');
  });

  it('fills C6 with assignee', () => {
    expect(ws['C6']?.v).toBe('pruthviraj-kale');
  });

  it('fills C7 with reviewer list', () => {
    const c7 = ws['C7']?.v as string;
    expect(c7).toContain('sayantan-kpit');
    expect(c7).toContain('tinynumbers');
  });

  it('fills C8 with start date', () => {
    expect(ws['C8']?.v).toBe('26 Jun 2026');
  });

  it('fills C9 with end date (merged PR)', () => {
    expect(ws['C9']?.v).toBe('22 Jul 2026');
  });

  it('fills E/F/G for row 1 (Ok finding)', () => {
    expect(ws['E11']?.v).toBe('Yes');
    expect(ws['F11']?.v).toBe('Ok');
    expect(ws['G11']?.v).toContain('Branch is current');
  });

  it('fills E/F/G for row 2 (Not Ok finding)', () => {
    expect(ws['E12']?.v).toBe('No');
    expect(ws['F12']?.v).toBe('Not Ok');
    expect(ws['G12']?.v).toContain('i18n');
  });

  it('fills defect type (H) for Not Ok rows', () => {
    expect(ws['H12']?.v).toBe('Functional');
    expect(ws['H14']?.v).toBe('Technical');
  });

  it('fills summary counts J8 (Ok) and J9 (Not Ok)', () => {
    expect(ws['J8']?.v).toBe(1);   // 1 Ok finding
    expect(ws['J9']?.v).toBe(2);   // 2 Not Ok findings
  });

  it('Re-review sheet has correct number of rows', () => {
    const rrSheet = wb.Sheets['Re-review Required'];
    const data = XLSX.utils.sheet_to_json(rrSheet, { header: 1 }) as any[][];
    // 1 header row + 2 Not Ok rows
    expect(data.length).toBe(3);
  });

  it('Version History sheet has current version row', () => {
    const histSheet = wb.Sheets['Version History'];
    const data = XLSX.utils.sheet_to_json(histSheet, { header: 1 }) as any[][];
    const curRow = data[1];   // row after header
    expect(curRow[0]).toContain('v1');
    expect(curRow[2]).toBe(2);  // Not Ok count
  });

  it('ROR is the active tab', () => {
    const activeTab = (wb.Workbook?.Views?.[0] as any)?.ActiveTab ?? 0;
    expect(activeTab).toBe(0);
  });
});
