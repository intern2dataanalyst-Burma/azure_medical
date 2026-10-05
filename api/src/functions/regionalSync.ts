import { app } from '@azure/functions';
import { z } from 'zod';

import { config } from '../config.js';
import { GraphClient } from '../services/graphClient.js';
import { errorResponse, requireAdmin } from '../services/requestAuth.js';
import { syncRegionalWorkbooks } from '../services/regionalSyncService.js';

const requestSchema = z.object({
  sourceFileIds: z.array(z.string().min(1)).max(50).optional(),
});

app.http('syncRegionalData', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'syncRegionalData',
  handler: async (request) => {
    try {
      requireAdmin(request);
      const body = await request.json().catch(() => ({}));
      const parsed = requestSchema.safeParse(body);
      if (!parsed.success) {
        return { status: 400, jsonBody: { error: { code: 'INVALID_REQUEST', message: 'Regional sourceFileIds must be an array of workbook ids.' } } };
      }
      const sourceFileIds = [...new Set(parsed.data.sourceFileIds ?? config.regionalSourceFileIds)]
        .filter((id) => id !== config.excelFileId);
      if (sourceFileIds.length === 0) {
        return { status: 400, jsonBody: { error: { code: 'MISSING_SOURCES', message: 'Provide sourceFileIds or configure REGIONAL_SOURCE_FILE_IDS.' } } };
      }
      const results = await syncRegionalWorkbooks(new GraphClient(), sourceFileIds);
      return { status: 200, jsonBody: { status: 'completed', files: results } };
    } catch (error) {
      return errorResponse(error);
    }
  },
});