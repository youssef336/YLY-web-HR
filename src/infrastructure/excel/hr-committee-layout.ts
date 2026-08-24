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
 *   CY  = Total (template formula)
 *   CZ  = % (template formula)
 *   DA  = Grade (template formula)
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
