/**
 * HRMBMER template column mapping.
 *
 * Discovered from inspecting the real HRMBMER (1).xlsx template:
 *   A   = Member name
 *   B   = Tasks Entered (FORMULA — NEVER written)
 *   C..BJ (30) = Task columns — each task takes exactly 3 columns (T, Q, D)
 *     Task 1: C (T), D (Q), E (D)
 *     Task 2: F (T), G (Q), H (D)
 *     ...
 *     Task 10: BH (T), BI (Q), BJ (D)
 *   BO  = Meetings Entered (FORMULA — NEVER written)
 *   BP..CD (15) = Meeting scores (0/1) — 15 columns
 *   CF  = Field Visits Entered (FORMULA — NEVER written)
 *   CG..CU (15) = Field Visit scores (0/1) — 15 columns
 *   CV  = Interaction /10 (direct input)
 *   CW  = Respect Hierarchy /10 (direct input)
 *   CX  = Bonus /10 (direct input)
 *   BN  = Tasks Total (formula)
 *   CE  = Meetings Total (formula)
 *   CV  = Field Visits Total (formula)
 *   CW..CY = direct category inputs
 *   CZ  = Final Total (formula)
 *   DA  = Percentage (formula)
 *   DB  = Grade (formula)
 *   DE  = Tie-breaker (formula)
 *
 * Template layout:
 *   row 2 = dark blue section headers (static text)
 *   row 3 = LIGHT YELLOW header row — task names injected into T columns (C3, F3, ...)
 *            meeting dates in BP3:CD3, field visit dates in CG3:CU3
 *   row 4 = DARK YELLOW numbering row — NEVER overwritten
 *   row 5..304 = one member per row (300 registry rows)
 */
export const HR_TEMPLATE = {
  sheetName: 'Member Evaluation',
  filePath: '/exel need/HRMBMER (1).xlsx',
  downloadFileName: 'HRMBMER (1).xlsx',
  headerRow: 3,
  firstDataRow: 5,
  lastDataRow: 304,
  columns: {
    name: 'A',
    tasksCount: 'B',
    tasksStart: 'C',
    tasksEnd: 'BJ',
    meetingsCount: 'BO',
    meetingsStart: 'BP',
    meetingsEnd: 'CD',
    fieldVisitsCount: 'CF',
    fieldVisitsStart: 'CG',
    fieldVisitsEnd: 'CU',
    interaction: 'CV',
    respectHierarchy: 'CW',
    bonus: 'CX',
  },
} as const;

/** Maximum number of HR tasks (30 columns / 3 columns per task). */
export const MAX_HR_TASKS = 10;

/** Maximum field visits and meetings. */
export const MAX_FIELD_VISITS = 15;
export const MAX_MEETINGS = 15;

export const MAX_REGISTRY_ROWS =
  HR_TEMPLATE.lastDataRow - HR_TEMPLATE.firstDataRow + 1;

/** Inject every HR calculation formula for the complete template registry. */
export function injectHrCalculationFormulas(
  sheet: import('exceljs').Worksheet,
  firstRow: number = HR_TEMPLATE.firstDataRow,
  lastRow: number = HR_TEMPLATE.lastDataRow,
): void {
  for (let row = firstRow; row <= lastRow; row++) {
    sheet.getCell(`BN${row}`).value = { formula: `=ROUND(SUM(BK${row}:BM${row}), 2)` };
    sheet.getCell(`CE${row}`).value = { formula: `=IFERROR(ROUND(SUM(BP${row}:CD${row}) / BO${row} * 30, 0), 0)` };
    sheet.getCell(`CV${row}`).value = { formula: `=IFERROR(ROUND(SUM(CG${row}:CU${row}) / CF${row} * 20, 0), 0)` };
    sheet.getCell(`CZ${row}`).value = { formula: `=ROUND(SUM(BN${row},CE${row},CV${row},CW${row},CX${row},CY${row}), 2)` };
    sheet.getCell(`DA${row}`).value = { formula: `=IFERROR(ROUND(CZ${row}/110*100, 1), 0)` };
    sheet.getCell(`DB${row}`).value = { formula: `=IF(CZ${row}>=90,"A",IF(CZ${row}>=80,"B",IF(CZ${row}>=70,"C",IF(CZ${row}>=60,"D",IF(CZ${row}>0,"F"," ")))))` };
    sheet.getCell(`DE${row}`).value = { formula: `=IF(A${row}="","",DA${row}-ROW()/100000000)` };
  }
}

/**
 * Get the ExcelJS column index (1-based) for a specific metric of a task.
 * Task 0 → columns C(3), D(4), E(5)
 * Task 1 → columns F(6), G(7), H(8)
 * ...
 */
export function taskColumnIndex(taskIndex: number, metric: 't' | 'q' | 'd'): number {
  const base = columnLetterToIndex(HR_TEMPLATE.columns.tasksStart);
  const offset = taskIndex * 3;
  const metricOffset = metric === 't' ? 0 : metric === 'q' ? 1 : 2;
  return base + offset + metricOffset;
}

export function columnLetterToIndex(letter: string): number {
  let index = 0;
  for (const char of letter.toUpperCase()) {
    index = index * 26 + (char.charCodeAt(0) - 64);
  }
  return index;
}
