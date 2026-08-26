import ExcelJS from 'exceljs';
import {
  HR_TEMPLATE,
  MAX_FIELD_VISITS,
  MAX_MEETINGS,
  MAX_REGISTRY_ROWS,
  injectHrCalculationFormulas,
} from './hr-committee-layout';
import { parseTemplateLayout, taskColIndex, type ParsedTemplateLayout } from './template-parser';

const { headerRow, firstDataRow, lastDataRow } = HR_TEMPLATE;

export interface OfficialEvent {
  name?: string;
  date: string;
  shift?: 'Day' | 'Night';
}

export interface OfficialTask {
  name: string;
  date: string;
}

function cellToNumber(value: ExcelJS.CellValue): number {
  if (value == null) return 0;
  if (typeof value === 'number') return value;
  if (typeof value === 'string') {
    const n = Number(value);
    return Number.isFinite(n) ? n : 0;
  }
  if (typeof value === 'object' && 'result' in value) {
    const r = (value as { result?: unknown }).result;
    return typeof r === 'number' ? r : 0;
  }
  return 0;
}

function cellToString(value: ExcelJS.CellValue): string | null {
  if (value == null) return null;
  if (typeof value === 'string') return value.trim() || null;
  if (typeof value === 'number') return String(value);
  if (typeof value === 'object' && 'result' in value) {
    const r = (value as { result?: unknown }).result;
    return r != null ? String(r) : null;
  }
  return null;
}

/**
 * Extract the date+shift key from a header label.
 * Visit headers are "Name - YYYY-MM-DD (Day)" → "YYYY-MM-DD (Day)"
 * Meeting headers are "Name - YYYY-MM-DD" → "YYYY-MM-DD"
 */
function extractDateKeyFromHeader(header: string | null): string | null {
  if (!header) return null;
  const dashIndex = header.lastIndexOf(' - ');
  const tail = dashIndex >= 0 ? header.substring(dashIndex + 3).trim() : header.trim();
  const match = tail.match(/^(\d{4}-\d{2}-\d{2})(?:\s*\((Day|Night)\))?$/);
  if (!match) return null;
  const date = match[1];
  const shift = match[2];
  return shift ? `${date} (${shift})` : date;
}

/**
 * Extract the date key from a task header label.
 * Task headers are "TaskName - YYYY-MM-DD" → "YYYY-MM-DD"
 */
function extractTaskDateKey(header: string | null): string | null {
  if (!header) return null;
  const dashIndex = header.lastIndexOf(' - ');
  if (dashIndex < 0) return null;
  const tail = header.substring(dashIndex + 3).trim();
  const match = tail.match(/^(\d{4}-\d{2}-\d{2})/);
  return match ? match[1] : null;
}

/**
 * Read Row 3 of a workbook and return slot→date maps for visits, meetings,
 * and tasks, using the dynamic template parser.
 */
function readSourceHeaders(workbook: ExcelJS.Workbook): {
  visits: Map<number, string>;
  meetings: Map<number, string>;
  tasks: Map<number, string>;
  layout: ParsedTemplateLayout;
} {
  const sheet = workbook.getWorksheet(HR_TEMPLATE.sheetName);
  const visits = new Map<number, string>();
  const meetings = new Map<number, string>();
  const tasks = new Map<number, string>();
  if (!sheet) return { visits, meetings, tasks, layout: parseTemplateLayout(sheet!) };

  const layout = parseTemplateLayout(sheet);

  const row = sheet.getRow(headerRow);

  // Read task date headers from T columns
  for (let i = 0; i < layout.taskCount; i++) {
    const header = cellToString(row.getCell(taskColIndex(layout.tasksStartCol, i, 't')).value);
    const dateKey = extractTaskDateKey(header);
    if (dateKey) tasks.set(i, dateKey);
  }

  // Read meeting date headers
  for (let i = 0; i < layout.meetingCount; i++) {
    const key = extractDateKeyFromHeader(cellToString(row.getCell(layout.meetingsStartCol + i).value));
    if (key) meetings.set(i, key);
  }

  // Read field visit date headers
  for (let i = 0; i < layout.fieldVisitCount; i++) {
    const key = extractDateKeyFromHeader(cellToString(row.getCell(layout.fieldVisitsStartCol + i).value));
    if (key) visits.set(i, key);
  }

  return { visits, meetings, tasks, layout };
}

/**
 * Read all member data rows from a workbook using the dynamic layout.
 */
function extractRawRows(layout: ParsedTemplateLayout, workbook: ExcelJS.Workbook): {
  name: string;
  tasks: Array<{ t: number; q: number; d: number }>;
  visits: (number | null)[];
  meetings: (number | null)[];
  interaction: number;
  respectHierarchy: number;
  bonus: number;
}[] {
  const sheet = workbook.getWorksheet(HR_TEMPLATE.sheetName);
  if (!sheet) return [];

  const rows: {
    name: string;
    tasks: Array<{ t: number; q: number; d: number }>;
    visits: (number | null)[];
    meetings: (number | null)[];
    interaction: number;
    respectHierarchy: number;
    bonus: number;
  }[] = [];

  for (let r = firstDataRow; r <= lastDataRow; r++) {
    const row = sheet.getRow(r);
    const name = cellToString(row.getCell(layout.nameCol).value);
    if (!name) break;

    // Read tasks from dynamic task columns
    const tasks: Array<{ t: number; q: number; d: number }> = [];
    for (let i = 0; i < layout.taskCount; i++) {
      const t = cellToNumber(row.getCell(taskColIndex(layout.tasksStartCol, i, 't')).value);
      const q = cellToNumber(row.getCell(taskColIndex(layout.tasksStartCol, i, 'q')).value);
      const d = cellToNumber(row.getCell(taskColIndex(layout.tasksStartCol, i, 'd')).value);
      tasks.push({ t, q, d });
    }

    // Read field visit scores
    const visits: (number | null)[] = [];
    for (let i = 0; i < layout.fieldVisitCount; i++) {
      const v = row.getCell(layout.fieldVisitsStartCol + i).value;
      visits.push(v != null ? cellToNumber(v) : null);
    }

    // Read meeting scores
    const meetings: (number | null)[] = [];
    for (let i = 0; i < layout.meetingCount; i++) {
      const v = row.getCell(layout.meetingsStartCol + i).value;
      meetings.push(v != null ? cellToNumber(v) : null);
    }

    rows.push({
      name,
      tasks,
      visits,
      meetings,
      interaction: cellToNumber(row.getCell(layout.interactionCol).value),
      respectHierarchy: cellToNumber(row.getCell(layout.respectHierarchyCol).value),
      bonus: cellToNumber(row.getCell(layout.bonusCol).value),
    });
  }

  return rows;
}

/**
 * Merge multiple uploaded HRMBMER .xlsx files into a single master template.
 *
 * The leader defines "official" events (field visits + meetings) and tasks
 * which become the master Row 3 headers. Uploaded files are mapped to the
 * master via DATE matching for all categories (tasks, visits, meetings).
 *
 * Monthly targets are written to "Entered" count columns for all rows.
 */
export async function mergeExcelFiles(
  files: File[],
  officialVisits: OfficialEvent[],
  officialMeetings: OfficialEvent[],
  officialTasks?: OfficialTask[],
  targets?: { tasks: number; meetings: number; fieldVisits: number },
): Promise<Uint8Array> {
  // 1. Load pristine master HRMBMER template
  const masterUrl = `${new URL(HR_TEMPLATE.filePath, window.location.origin).href}?v=${Date.now()}`;
  const masterRes = await fetch(masterUrl, {
    cache: 'no-store',
    headers: {
      'Cache-Control': 'no-cache, no-store, must-revalidate',
      'Pragma': 'no-cache',
      'Expires': '0',
    },
  });
  if (!masterRes.ok) throw new Error(`Master template not found (HTTP ${masterRes.status})`);

  const masterWb = new ExcelJS.Workbook();
  await masterWb.xlsx.load(await masterRes.arrayBuffer());
  const masterSheet = masterWb.getWorksheet(HR_TEMPLATE.sheetName);
  if (!masterSheet) throw new Error(`Sheet "${HR_TEMPLATE.sheetName}" not found in template`);

  // Dynamically parse the master template layout
  const masterLayout = parseTemplateLayout(masterSheet);

  // 2. Write official headers to Row 3 of the master
  const masterHeaderRow = masterSheet.getRow(headerRow);
  const masterVisitKeyToSlot = new Map<string, number>();
  const masterMeetingKeyToSlot = new Map<string, number>();
  const masterTaskDateToSlot = new Map<string, number>();

  // Write official task headers: "TaskName - YYYY-MM-DD" in T columns
  if (officialTasks) {
    for (let i = 0; i < officialTasks.length && i < masterLayout.taskCount; i++) {
      const task = officialTasks[i];
      if (task.name) {
        const label = task.date ? `${task.name} - ${task.date}` : task.name;
        masterHeaderRow.getCell(taskColIndex(masterLayout.tasksStartCol, i, 't')).value = label;
        if (task.date) masterTaskDateToSlot.set(task.date, i);
      }
    }
  }

  // Write official field visit headers
  for (let i = 0; i < officialVisits.length && i < MAX_FIELD_VISITS; i++) {
    const ev = officialVisits[i];
    const shift = ev.shift ?? 'Day';
    const label = ev.name ? `${ev.name} - ${ev.date} (${shift})` : `${ev.date} (${shift})`;
    masterHeaderRow.getCell(masterLayout.fieldVisitsStartCol + i).value = label;
    masterVisitKeyToSlot.set(`${ev.date} (${shift})`, i);
  }

  // Write official meeting headers
  for (let i = 0; i < officialMeetings.length && i < MAX_MEETINGS; i++) {
    const ev = officialMeetings[i];
    const name = ev.name ?? 'Meeting';
    const label = `${name} - ${ev.date}`;
    masterHeaderRow.getCell(masterLayout.meetingsStartCol + i).value = label;
    masterMeetingKeyToSlot.set(ev.date, i);
  }

  // 3. Process uploaded files — date-based mapping for ALL categories
  const allRemappedRows: {
    name: string;
    tasks: Array<{ t: number; q: number; d: number }>;
    visits: (number | null)[];
    meetings: (number | null)[];
    interaction: number;
    respectHierarchy: number;
    bonus: number;
  }[] = [];

  for (const file of files) {
    const buffer = await file.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);

    const sourceHeaders = readSourceHeaders(wb);
    const sourceRows = extractRawRows(sourceHeaders.layout, wb);

    for (const src of sourceRows) {
      // Remap task scores: source slot → source date → master slot
      const remappedTasks: Array<{ t: number; q: number; d: number }> = Array.from(
        { length: masterLayout.taskCount },
        () => ({ t: 0, q: 0, d: 0 }),
      );
      for (const [srcSlot, dateKey] of sourceHeaders.tasks) {
        const masterSlot = masterTaskDateToSlot.get(dateKey);
        if (masterSlot != null && src.tasks[srcSlot]) {
          remappedTasks[masterSlot] = src.tasks[srcSlot];
        }
      }

      // Remap visit scores: source slot → source date+key → master slot
      const remappedVisits: (number | null)[] = Array(MAX_FIELD_VISITS).fill(null);
      for (const [srcSlot, key] of sourceHeaders.visits) {
        const masterSlot = masterVisitKeyToSlot.get(key);
        if (masterSlot != null && src.visits[srcSlot] != null) {
          remappedVisits[masterSlot] = src.visits[srcSlot];
        }
      }

      // Remap meeting scores: source slot → source date → master slot
      const remappedMeetings: (number | null)[] = Array(MAX_MEETINGS).fill(null);
      for (const [srcSlot, key] of sourceHeaders.meetings) {
        const masterSlot = masterMeetingKeyToSlot.get(key);
        if (masterSlot != null && src.meetings[srcSlot] != null) {
          remappedMeetings[masterSlot] = src.meetings[srcSlot];
        }
      }

      allRemappedRows.push({
        name: src.name,
        tasks: remappedTasks,
        visits: remappedVisits,
        meetings: remappedMeetings,
        interaction: src.interaction,
        respectHierarchy: src.respectHierarchy,
        bonus: src.bonus,
      });
    }
  }

  if (allRemappedRows.length === 0) {
    throw new Error('No member data found in the uploaded files.');
  }

  if (allRemappedRows.length > MAX_REGISTRY_ROWS) {
    throw new Error(
      `Too many members (${allRemappedRows.length}). The template supports ${MAX_REGISTRY_ROWS} rows maximum.`,
    );
  }

  // 4. Write remapped rows to master template
  for (let i = 0; i < MAX_REGISTRY_ROWS; i++) {
    const data = allRemappedRows[i];
    const row = masterSheet.getRow(firstDataRow + i);

    if (!data) continue;

    row.getCell(masterLayout.nameCol).value = data.name;

    // Write tasks (dynamic columns)
    for (let t = 0; t < masterLayout.taskCount; t++) {
      const task = data.tasks[t];
      if (task && (task.t !== 0 || task.q !== 0 || task.d !== 0)) {
        row.getCell(taskColIndex(masterLayout.tasksStartCol, t, 't')).value = task.t;
        row.getCell(taskColIndex(masterLayout.tasksStartCol, t, 'q')).value = task.q;
        row.getCell(taskColIndex(masterLayout.tasksStartCol, t, 'd')).value = task.d;
      }
    }

    // Write field visit scores
    for (let s = 0; s < MAX_FIELD_VISITS; s++) {
      if (data.visits[s] != null) row.getCell(masterLayout.fieldVisitsStartCol + s).value = data.visits[s];
    }

    // Write meeting scores
    for (let s = 0; s < MAX_MEETINGS; s++) {
      if (data.meetings[s] != null) row.getCell(masterLayout.meetingsStartCol + s).value = data.meetings[s];
    }

    row.getCell(masterLayout.interactionCol).value = data.interaction;
    row.getCell(masterLayout.respectHierarchyCol).value = data.respectHierarchy;
    row.getCell(masterLayout.bonusCol).value = data.bonus;

    // 5. Overwrite "Entered" count columns with target numbers
    if (targets) {
      row.getCell(masterLayout.tasksCountCol).value = targets.tasks;
      row.getCell(masterLayout.meetingsCountCol).value = targets.meetings;
      row.getCell(masterLayout.fieldVisitsCountCol).value = targets.fieldVisits;
    }
  }

  injectHrCalculationFormulas(masterSheet, masterLayout.firstDataRow, HR_TEMPLATE.lastDataRow);

  // 6. Strip cached formula results so Excel recalculates on open
  masterSheet.eachRow((row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      if (cell && typeof cell.value === 'object' && cell.value !== null && 'formula' in cell.value) {
        delete (cell.value as unknown as Record<string, unknown>).result;
      }
    });
  });

  const buffer = await masterWb.xlsx.writeBuffer();
  return new Uint8Array(buffer as ArrayBuffer);
}
