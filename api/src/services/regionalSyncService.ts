import type { GraphClient, GraphWorkbookRow } from './graphClient.js';

export type RegionalSyncResult = {
  sourceFileId: string;
  scanned: number;
  inserted: number;
  updated: number;
  skipped: number;
};

export async function syncRegionalWorkbooks(graph: GraphClient, sourceFileIds: string[]): Promise<RegionalSyncResult[]> {
  graph.assertWorkbookSettings();
  const session = await graph.createWorkbookSession(true);
  try {
    const masterRows = await graph.getMasterRows(session.id);
    const masterByEmpId = new Map<string, GraphWorkbookRow>();
    for (const row of masterRows) {
      const empId = String(row.values?.[0]?.[1] ?? '').trim().toLowerCase();
      if (empId) masterByEmpId.set(empId, row);
    }

    const results: RegionalSyncResult[] = [];
    for (const sourceFileId of sourceFileIds) {
      if (sourceFileId === '') continue;
      const sourceRows = await graph.getMasterRows(undefined, sourceFileId);
      const toAppend: unknown[][] = [];
      let updated = 0;
      let skipped = 0;
      for (const sourceRow of sourceRows) {
        const values = Array.isArray(sourceRow.values?.[0]) ? [...sourceRow.values[0]] : [];
        const empId = String(values[1] ?? '').trim().toLowerCase();
        if (!empId || values.length < 17) {
          skipped++;
          continue;
        }
        const existing = masterByEmpId.get(empId);
        if (existing) {
          await graph.updateMasterRow(existing.index, values, session.id);
          existing.values = [values];
          updated++;
        } else {
          toAppend.push(values);
          masterByEmpId.set(empId, { index: masterRows.length + toAppend.length - 1, values: [values] });
        }
      }
      await graph.appendMasterRows(toAppend, session.id);
      results.push({ sourceFileId, scanned: sourceRows.length, inserted: toAppend.length, updated, skipped });
    }
    return results;
  } finally {
    await graph.closeWorkbookSession(session.id);
  }
}
