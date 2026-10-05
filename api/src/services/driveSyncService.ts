import { config } from '../config.js';
import type { GraphDriveItem, GraphWorkbookRow } from './graphClient.js';
import { GraphClient } from './graphClient.js';
import { parseDateValue } from './medicalDataService.js';

export type DriveSyncFileResult = {
  fileName: string;
  empId?: string;
  status: 'moved' | 'skipped' | 'failed';
  folderPath?: string;
  message?: string;
};

export function folderName(value: unknown, fallback: string): string {
  const normalized = String(value ?? '').trim().replace(/["*:<>?/\\|]/g, '-').replace(/[. ]+$/g, '');
  return normalized || fallback;
}

export function monthFolder(value: unknown): string {
  const date = parseDateValue(value);
  if (!date) return 'New Employee';
  return new Intl.DateTimeFormat('en-US', { month: 'short', year: 'numeric', timeZone: 'UTC' }).format(date);
}

export function findEmployeeId(fileName: string, records: Map<string, GraphWorkbookRow>): string | undefined {
  const candidateIds = [...records.keys()].sort((left, right) => right.length - left.length);
  for (const empId of candidateIds) {
    const escaped = empId.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const pattern = new RegExp(`(^|[^A-Za-z0-9])${escaped}(?=$|[^A-Za-z0-9])`, 'i');
    if (pattern.test(fileName)) return empId;
  }
  return undefined;
}

function isPdf(item: GraphDriveItem): boolean {
  return Boolean(item.file) && item.name.toLowerCase().endsWith('.pdf');
}

export async function runMasterDriveSync(graph: GraphClient): Promise<{ moved: number; skipped: number; failed: number; files: DriveSyncFileResult[] }> {
  graph.assertDriveSettings();
  const session = await graph.createWorkbookSession(true);
  const results: DriveSyncFileResult[] = [];
  try {
    const [workbookRows, incomingItems] = await Promise.all([
      graph.getMasterRows(session.id),
      graph.listRawReports(),
    ]);
    const recordById = new Map<string, GraphWorkbookRow>();
    for (const row of workbookRows) {
      const values = row.values?.[0] ?? [];
      const empId = String(values[1] ?? '').trim();
      if (empId) recordById.set(empId.toLowerCase(), row);
    }

    for (const item of incomingItems.filter(isPdf)) {
      const empId = findEmployeeId(item.name, recordById);
      if (!empId) {
        results.push({ fileName: item.name, status: 'skipped', message: 'No matching EMP ID in MasterMedicalTable.' });
        continue;
      }
      const row = recordById.get(empId)!;
      const values = [...(row.values?.[0] ?? [])];
      try {
        const region = folderName(values[5], 'Unassigned Region');
        const branch = folderName(values[4], 'Unassigned Branch');
        const period = monthFolder(values[11]);
        const regionFolder = await graph.ensureChildFolder(config.organizedReportsFolderId, region);
        const branchFolder = await graph.ensureChildFolder(regionFolder.id, branch);
        const periodFolder = await graph.ensureChildFolder(branchFolder.id, period);
        const moved = await graph.moveDriveItem(item.id, periodFolder.id);
        if (!moved.webUrl) throw new Error('Graph did not return a webUrl for the moved report.');
        await graph.updateReportLink(row.index, values, moved.webUrl, session.id);
        results.push({ fileName: item.name, empId, status: 'moved', folderPath: `${region}/${branch}/${period}` });
      } catch (error) {
        results.push({
          fileName: item.name,
          empId,
          status: 'failed',
          message: error instanceof Error ? error.message : 'Graph operation failed.',
        });
      }
    }
  } finally {
    await graph.closeWorkbookSession(session.id);
  }
  return {
    moved: results.filter((result) => result.status === 'moved').length,
    skipped: results.filter((result) => result.status === 'skipped').length,
    failed: results.filter((result) => result.status === 'failed').length,
    files: results,
  };
}