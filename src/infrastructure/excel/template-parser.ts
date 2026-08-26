import ExcelJS from 'exceljs';

/**
 * Dynamically parsed layout of an HRMBMER-style Excel template.
 *
 * Instead of hardcoding column positions, this reads Row 2 (section headers)
 * and Row 3 (column headers) of the uploaded template to discover where
 * TASKS, MEETINGS, and FIELD VISITS columns begin.
 */
export interface ParsedTemplateLayout {
  /** 1-based column index where task T columns start (e.g. C = 3). */
  tasksStartCol: number;
  /** 1-based column index where meeting score columns start (e.g. BP = 68). */
  meetingsStartCol: number;
  /** 1-based column index where field visit score columns start (e.g. CG = 85). */
  fieldVisitsStartCol: number;
  /** 1-based column index of the member name column (typically A = 1). */
  nameCol: number;
  /** 1-based column index of the "Tasks Entered" count column (typically B = 2). */
  tasksCountCol: number;
  /** 1-based column index of the "Meetings Entered" count column. */
  meetingsCountCol: number;
  /** 1-based column index of the "Field Visits Entered" count column. */
  fieldVisitsCountCol: number;
  /** 1-based column index of the Interaction column. */
  interactionCol: number;
  /** 1-based column index of the Respect Hierarchy column. */
  respectHierarchyCol: number;
  /** 1-based column index of the Bonus column. */
  bonusCol: number;
  /** Number of task clusters discovered (each = 3 columns T/Q/D). */
  taskCount: number;
  /** Number of meeting columns discovered. */
  meetingCount: number;
  /** Number of field visit columns discovered. */
  fieldVisitCount: number;
  /** Row number containing section headers (dark blue row). */
  sectionHeaderRow: number;
  /** Row number containing column headers (light yellow row). */
  columnHeaderRow: number;
  /** Row number where data begins. */
  firstDataRow: number;
}

const cellToString = (value: ExcelJS.CellValue): string | null => {
  if (value == null) return null;
  if (typeof value === 'string') return value.trim() || null;
  if (typeof value === 'number') return String(value);
  if (typeof value === 'object' && 'result' in value) {
    const r = (value as { result?: unknown }).result;
    return r != null ? String(r) : null;
  }
  return null;
};

/**
 * Column-letter to 1-based index (A→1, B→2, Z→26, AA→27, ...).
 */
function colLetterToIndex(letter: string): number {
  let index = 0;
  for (const char of letter.toUpperCase()) {
    index = index * 26 + (char.charCodeAt(0) - 64);
  }
  return index;
}

/**
 * Scan a row and return the 1-based column index of the first cell whose
 * text (case-insensitive) includes the given keyword, or null if not found.
 */
function findColumnByKeyword(
  sheet: ExcelJS.Worksheet,
  rowNumber: number,
  keyword: string,
  startFrom: number = 1,
): number | null {
  const row = sheet.getRow(rowNumber);
  let maxCol = 100;
  try {
    row.eachCell({ includeEmpty: false }, (cell) => {
      const colNum = cell.fullAddress.col;
      if (colNum > maxCol) maxCol = colNum;
    });
  } catch { /* ignore */ }

  for (let c = startFrom; c <= maxCol; c++) {
    const text = cellToString(row.getCell(c).value);
    if (text && text.toUpperCase().includes(keyword.toUpperCase())) {
      return c;
    }
  }
  return null;
}

/**
 * Determine the data start column for a section given its keyword position
 * in Row 2. The keyword may be at the count column or the first data column.
 *
 * Strategy: check if Row 3 at the keyword position has a substantial header
 * (≥5 chars, typical of count-column labels like "Tasks Entered"). If so,
 * the keyword is at the count column and data starts at keywordCol + 1.
 * Otherwise, the keyword is at the first data column.
 */
function resolveDataStartCol(
  sheet: ExcelJS.Worksheet,
  keywordCol: number,
  columnHeaderRow: number,
): number {
  const headerText = cellToString(sheet.getRow(columnHeaderRow).getCell(keywordCol).value);
  if (headerText && headerText.length >= 5) {
    // Count column header detected (e.g. "Tasks Entered") → data starts at next column
    return keywordCol + 1;
  }
  // Keyword is at the first data column (e.g. "TASKS" merged header starts at C)
  return keywordCol;
}

/**
 * Count how many consecutive non-empty header cells exist starting from
 * `startCol` in the given row. Stops at the first empty cell or a cell
 * that contains a formula-only marker.
 */
function countConsecutiveHeaders(
  sheet: ExcelJS.Worksheet,
  rowNumber: number,
  startCol: number,
  maxCount: number,
): number {
  const row = sheet.getRow(rowNumber);
  let count = 0;
  for (let c = startCol; c < startCol + maxCount; c++) {
    const val = cellToString(row.getCell(c).value);
    if (!val) break;
    count++;
  }
  return count;
}

/**
 * Dynamically parse an HRMBMER-style Excel template to discover its layout.
 *
 * Reads Row 2 for section keywords ("TASKS", "MEETINGS", "FIELD VISITS")
 * and Row 3 for individual column headers to determine exact column positions.
 *
 * Falls back to the hardcoded HRMBMER layout if keyword scanning fails or
 * detected positions fall outside expected ranges.
 */
export function parseTemplateLayout(sheet: ExcelJS.Worksheet): ParsedTemplateLayout {
  // --- Discover rows ---
  // Row 2 = section headers (dark blue), Row 3 = column headers (light yellow)
  const sectionHeaderRow = 2;
  const columnHeaderRow = 3;
  const firstDataRow = 5;

  // --- Name column: always A ---
  const nameCol = 1;

  // --- Scan Row 2 for section keywords ---
  const tasksSectionCol = findColumnByKeyword(sheet, sectionHeaderRow, 'TASK');
  const meetingsSectionCol = findColumnByKeyword(sheet, sectionHeaderRow, 'MEETING');
  const visitsSectionCol = findColumnByKeyword(sheet, sectionHeaderRow, 'FIELD VISIT');

  // --- Hardcoded HRMBMER fallback values ---
  const HRMBMER_TASKS_START = colLetterToIndex('C');  // 3
  const HRMBMER_MEETINGS_START = colLetterToIndex('BP'); // 68
  const HRMBMER_VISITS_START = colLetterToIndex('CG'); // 85

  // --- Determine task columns ---
  let tasksStartCol: number;
  let taskCount: number;
  if (tasksSectionCol) {
    tasksStartCol = resolveDataStartCol(sheet, tasksSectionCol, columnHeaderRow);
    // Validate: task columns must be in columns 2–6 (reasonable range for any template)
    if (tasksStartCol < 2 || tasksStartCol > 6) {
      tasksStartCol = HRMBMER_TASKS_START;
    }
    // Count T-column headers in Row 3
    let tHeaders = 0;
    const row3 = sheet.getRow(columnHeaderRow);
    for (let c = tasksStartCol; c < tasksStartCol + 40; c += 3) {
      const val = cellToString(row3.getCell(c).value);
      if (!val) break;
      tHeaders++;
    }
    const consecutive = countConsecutiveHeaders(sheet, columnHeaderRow, tasksStartCol, 40);
    taskCount = tHeaders || Math.ceil(consecutive / 3);
    if (taskCount === 0) taskCount = 10; // fallback
  } else {
    // Fallback to hardcoded HRMBMER layout
    tasksStartCol = HRMBMER_TASKS_START;
    taskCount = 10;
  }

  // --- Determine meetings columns ---
  let meetingsStartCol: number;
  let meetingCount: number;
  if (meetingsSectionCol) {
    meetingsStartCol = resolveDataStartCol(sheet, meetingsSectionCol, columnHeaderRow);
    // Validate: meeting columns must be in range 66–70
    if (meetingsStartCol < 66 || meetingsStartCol > 70) {
      meetingsStartCol = HRMBMER_MEETINGS_START;
    }
    meetingCount = countConsecutiveHeaders(sheet, columnHeaderRow, meetingsStartCol, 30);
    if (meetingCount === 0) meetingCount = 15;
  } else {
    meetingsStartCol = HRMBMER_MEETINGS_START;
    meetingCount = 15;
  }

  // --- Determine field visits columns ---
  let fieldVisitsStartCol: number;
  let fieldVisitCount: number;
  if (visitsSectionCol) {
    fieldVisitsStartCol = resolveDataStartCol(sheet, visitsSectionCol, columnHeaderRow);
    // Validate: field visit columns must be in range 83–87
    if (fieldVisitsStartCol < 83 || fieldVisitsStartCol > 87) {
      fieldVisitsStartCol = HRMBMER_VISITS_START;
    }
    fieldVisitCount = countConsecutiveHeaders(sheet, columnHeaderRow, fieldVisitsStartCol, 30);
    if (fieldVisitCount === 0) fieldVisitCount = 15;
  } else {
    fieldVisitsStartCol = HRMBMER_VISITS_START;
    fieldVisitCount = 15;
  }

  // --- Determine trailing columns (Interaction, Respect, Bonus) ---
  // These come after the field visits section.
  const visitsEndCol = fieldVisitsStartCol + fieldVisitCount;
  // CV is the field-visits total; category inputs follow it.
  const interactionCol = visitsEndCol + 2;    // typically CW
  const respectHierarchyCol = visitsEndCol + 3; // typically CX
  const bonusCol = visitsEndCol + 4;          // typically CY

  // --- Determine count columns ---
  // "Tasks Entered" is one column before tasksStartCol
  const tasksCountCol = tasksStartCol - 1;
  // "Meetings Entered" is one column before meetingsStartCol
  const meetingsCountCol = meetingsStartCol - 1;
  // "Field Visits Entered" is one column before fieldVisitsStartCol
  const fieldVisitsCountCol = fieldVisitsStartCol - 1;

  return {
    tasksStartCol,
    meetingsStartCol,
    fieldVisitsStartCol,
    nameCol,
    tasksCountCol,
    meetingsCountCol,
    fieldVisitsCountCol,
    interactionCol,
    respectHierarchyCol,
    bonusCol,
    taskCount,
    meetingCount,
    fieldVisitCount,
    sectionHeaderRow,
    columnHeaderRow,
    firstDataRow,
  };
}

/**
 * Compute the 1-based ExcelJS column index for a specific task metric.
 * `taskIndex` is 0-based, metric is 't' | 'q' | 'd'.
 */
export function taskColIndex(
  tasksStartCol: number,
  taskIndex: number,
  metric: 't' | 'q' | 'd',
): number {
  const offset = taskIndex * 3;
  const metricOffset = metric === 't' ? 0 : metric === 'q' ? 1 : 2;
  return tasksStartCol + offset + metricOffset;
}
