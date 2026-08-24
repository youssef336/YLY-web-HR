import ExcelJS from 'exceljs';
import type { ExcelGenerator } from '@/application/ports/excel-generator.port';
import type { MemberProfile } from '@/domain/entities/member-profile';
import type { GlobalFieldVisit } from '@/domain/entities/global-field-visit';
import type { GlobalMeeting } from '@/domain/entities/global-meeting';
import { fieldVisitHeaderLabel } from '@/domain/entities/global-field-visit';
import { meetingHeaderLabel } from '@/domain/entities/global-meeting';
import {
  HR_TEMPLATE,
  MAX_FIELD_VISITS,
  MAX_MEETINGS,
} from './hr-committee-layout';
import { parseTemplateLayout, taskColIndex, type ParsedTemplateLayout } from './template-parser';

/**
 * Injects local member data into the REAL HRMBMER template using exceljs,
 * entirely in the browser (offline-first — no server involved).
 *
 * Template fidelity rules:
 *   - Loads the ORIGINAL .xlsx from public/exel need/ via workbook.xlsx.load().
 *   - NEVER creates a new workbook or worksheet — modifies the existing one.
 *   - ONLY writes to cells that have actual data to inject.
 *   - NEVER touches formula columns (B, BO, CF, CY, CZ, DA).
 *   - Strips cached formula results so Excel recalculates on open.
 */
export class ExceljsInjector implements ExcelGenerator {
  async generateAll(input: {
    profiles: MemberProfile[];
    globalFieldVisits: GlobalFieldVisit[];
    globalMeetings: GlobalMeeting[];
  }): Promise<Uint8Array> {
    const { profiles, globalFieldVisits, globalMeetings } = input;

    // Fetch the pristine HRMBMER template
    const cacheBust = new Date().getTime();
    const templateUrl = `${new URL(HR_TEMPLATE.filePath, window.location.origin).href}?v=${cacheBust}`;
    let templateBytes: ArrayBuffer;
    try {
      const res = await fetch(templateUrl, {
        cache: 'no-store',
        headers: {
          'Cache-Control': 'no-cache, no-store, must-revalidate',
          'Pragma': 'no-cache',
          'Expires': '0',
        },
      });
      if (!res.ok) throw new Error(`template not found (HTTP ${res.status})`);
      templateBytes = await res.arrayBuffer();
    } catch (cause) {
      throw new Error(
        `HRMBMER template not bundled. Copy "HRMBMER (1).xlsx" into public/exel need/ and rebuild.`,
        { cause },
      );
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(templateBytes);
    const sheet = workbook.getWorksheet(HR_TEMPLATE.sheetName);
    if (!sheet) throw new Error(`Sheet "${HR_TEMPLATE.sheetName}" not found in template`);

    this.clearCachedFormulaResults(sheet);

    // Dynamically parse the template layout
    const layout = parseTemplateLayout(sheet);

    // Build slot→label maps for meetings and field visits
    const meetingHeaders = this.buildSlotLabels(profiles, 'meetings', globalMeetings);
    const visitHeaders = this.buildSlotLabels(profiles, 'fieldVisits', globalFieldVisits);

    // Write meeting headers
    this.writeHeaders(sheet, meetingHeaders, layout.meetingsStartCol, MAX_MEETINGS, layout.columnHeaderRow);
    // Write field visit headers
    this.writeHeaders(sheet, visitHeaders, layout.fieldVisitsStartCol, MAX_FIELD_VISITS, layout.columnHeaderRow);

    // Write task name + date headers into the header row
    this.writeTaskHeaders(sheet, profiles, layout);

    // Sort profiles by total descending
    const sorted = [...profiles].sort((a, b) => {
      if (b.summary.total !== a.summary.total) return b.summary.total - a.summary.total;
      return a.member.name.localeCompare(b.member.name);
    });

    sorted.forEach((profile, index) => {
      const row = sheet.getRow(layout.firstDataRow + index);

      // Name (column A)
      row.getCell(layout.nameCol).value = profile.member.name;

      // HR Tasks — each task: T, Q, D in their dynamic columns
      for (const task of profile.hrTasks) {
        if (task.taskIndex >= (layout.taskCount || 10)) continue;
        row.getCell(taskColIndex(layout.tasksStartCol, task.taskIndex, 't')).value = task.t;
        row.getCell(taskColIndex(layout.tasksStartCol, task.taskIndex, 'q')).value = task.q;
        row.getCell(taskColIndex(layout.tasksStartCol, task.taskIndex, 'd')).value = task.d;
      }

      // Meeting scores
      for (const meeting of profile.meetings) {
        if (meeting.slot >= MAX_MEETINGS) continue;
        row.getCell(layout.meetingsStartCol + meeting.slot).value = meeting.score;
      }

      // Field visit scores
      for (const visit of profile.fieldVisits) {
        if (visit.slot >= MAX_FIELD_VISITS) continue;
        row.getCell(layout.fieldVisitsStartCol + visit.slot).value = visit.score;
      }

      // Category scores
      row.getCell(layout.interactionCol).value = profile.scores?.interaction ?? 0;
      row.getCell(layout.respectHierarchyCol).value = profile.scores?.respectHierarchy ?? 0;
      row.getCell(layout.bonusCol).value = profile.scores?.bonus ?? 0;
    });

    const buffer = await workbook.xlsx.writeBuffer();
    return new Uint8Array(buffer as ArrayBuffer);
  }

  /**
   * Write task name + date headers into the header row.
   * For each task cluster: T column = "TaskName - YYYY-MM-DD", Q/D columns empty.
   * Uses the first member that has a task at each index.
   */
  private writeTaskHeaders(
    sheet: ExcelJS.Worksheet,
    profiles: MemberProfile[],
    layout: ParsedTemplateLayout,
  ): void {
    const row = sheet.getRow(layout.columnHeaderRow);
    const taskData = new Map<number, { name: string; date: string }>();
    for (const profile of profiles) {
      for (const task of profile.hrTasks) {
        if (!taskData.has(task.taskIndex) && task.name) {
          taskData.set(task.taskIndex, { name: task.name, date: task.date });
        }
      }
    }
    for (let i = 0; i < (layout.taskCount || 10); i++) {
      const data = taskData.get(i);
      if (data) {
        const label = data.date
          ? `${data.name} - ${data.date}`
          : data.name;
        row.getCell(taskColIndex(layout.tasksStartCol, i, 't')).value = label;
      }
    }
  }

  private buildSlotLabels(
    profiles: MemberProfile[],
    kind: 'fieldVisits' | 'meetings',
    globalEvents: GlobalFieldVisit[] | GlobalMeeting[],
  ): Map<number, string> {
    const bySlot = new Map<number, string>();
    const globalMap = new Map(globalEvents.map((e) => [e.id, e]));
    for (const profile of profiles) {
      for (const entry of profile[kind]) {
        const globalEvent = globalMap.get(entry.globalEventId);
        if (!globalEvent) continue;
        const label =
          kind === 'fieldVisits'
            ? fieldVisitHeaderLabel(globalEvent as GlobalFieldVisit)
            : meetingHeaderLabel(globalEvent as GlobalMeeting);
        bySlot.set(entry.slot, label);
      }
    }
    return bySlot;
  }

  private writeHeaders(
    sheet: ExcelJS.Worksheet,
    labelsBySlot: Map<number, string>,
    startCol: number,
    maxSlots: number,
    rowNumber: number,
  ): void {
    const headerRowCells = sheet.getRow(rowNumber);
    for (let slot = 0; slot < maxSlots; slot++) {
      const label = labelsBySlot.get(slot);
      if (!label) continue;
      headerRowCells.getCell(startCol + slot).value = label;
    }
  }

  private clearCachedFormulaResults(sheet: ExcelJS.Worksheet): void {
    sheet.eachRow((row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (cell && typeof cell.value === 'object' && cell.value !== null && 'formula' in cell.value) {
          delete (cell.value as unknown as Record<string, unknown>).result;
        }
      });
    });
  }
}
