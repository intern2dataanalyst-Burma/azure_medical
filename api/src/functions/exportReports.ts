import { app } from '@azure/functions';
import type { Archiver, ArchiverOptions } from 'archiver';
import { createRequire } from 'node:module';
import { PassThrough } from 'node:stream';
import { z } from 'zod';

import { GraphClient } from '../services/graphClient.js';
import { mapMasterRow } from '../services/medicalDataService.js';
import { errorResponse, getPrincipal } from '../services/requestAuth.js';
import { listAccessUsers } from '../services/accessControlService.js';

const requestSchema = z.object({ fileUrls: z.array(z.string().url()).min(1).max(100) });
const require = createRequire(import.meta.url);
const createArchive = require('archiver') as (format: 'zip', options: ArchiverOptions) => Archiver;

app.http('exportReports', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'exportReports',
  handler: async (request) => {
    try {
      const principal = getPrincipal(request);
      const body = await request.json().catch(() => null);
      const parsed = requestSchema.safeParse(body);
      if (!parsed.success) {
        return { status: 400, jsonBody: { error: { code: 'INVALID_REQUEST', message: 'Provide between 1 and 100 valid report URLs.' } } };
      }

      const graph = new GraphClient();
      const [rawRows, access] = await Promise.all([
        graph.getMasterRows(),
        principal.roles.includes('admin') ? Promise.resolve(null) : listAccessUsers(graph),
      ]);
      const records = rawRows.map(mapMasterRow);
      const requested = new Set(parsed.data.fileUrls);
      const allowed = records.filter((record) => requested.has(record.fileUrl) && record.fileUrl);
      if (!principal.roles.includes('admin')) {
        const user = access?.users.find((candidate) => candidate.email.toLowerCase() === principal.userDetails.toLowerCase());
        const branches = new Set((user?.branches ?? '').split(',').map((branch) => branch.trim().toLowerCase()).filter(Boolean));
        const denied = allowed.filter((record) => !branches.has(record.branch.toLowerCase()));
        if (denied.length) return { status: 403, jsonBody: { error: { code: 'FORBIDDEN', message: 'You do not have access to one or more selected reports.' } } };
      }
      if (allowed.length !== requested.size) {
        return { status: 403, jsonBody: { error: { code: 'FORBIDDEN', message: 'One or more selected reports are not listed in MasterMedicalTable.' } } };
      }

      const archive = createArchive('zip', { zlib: { level: 6 } });
      const output = new PassThrough();
      const chunks: Buffer[] = [];
      output.on('data', (chunk: Buffer) => chunks.push(chunk));
      archive.pipe(output);
      let appended = 0;
      const usedNames = new Set<string>();
      for (const record of allowed) {
        const item = await graph.resolveSharedItem(record.fileUrl);
        if (!item.file || !item.name.toLowerCase().endsWith('.pdf')) continue;
        const content = await graph.downloadDriveItem(item.id);
        const baseName = item.name.replace(/[\\/:*?"<>|]/g, '_');
        let fileName = `${record.empId}_${baseName}`;
        let suffix = 1;
        while (usedNames.has(fileName.toLowerCase())) fileName = `${record.empId}_${suffix++}_${baseName}`;
        usedNames.add(fileName.toLowerCase());
        archive.append(content, { name: fileName });
        appended++;
      }
      const streamComplete = new Promise<void>((resolve, reject) => {
        output.on('end', resolve);
        output.on('error', reject);
        archive.on('error', reject);
      });
      await archive.finalize();
      await streamComplete;
      const zip = Buffer.concat(chunks);
      if (appended === 0) return { status: 404, jsonBody: { error: { code: 'NO_REPORTS', message: 'No PDF reports were found for the selected rows.' } } };
      return {
        status: 200,
        body: zip,
        headers: {
          'Content-Type': 'application/zip',
          'Content-Disposition': 'attachment; filename="medical-reports.zip"',
        },
      };
    } catch (error) {
      return errorResponse(error);
    }
  },
});