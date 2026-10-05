import { app } from '@azure/functions';

import { GraphClient } from '../services/graphClient.js';
import { errorResponse, requireAdmin } from '../services/requestAuth.js';
import { runMasterDriveSync } from '../services/driveSyncService.js';

app.http('runMasterDriveSync', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'runMasterDriveSync',
  handler: async (request, context) => {
    try {
      requireAdmin(request);
      const result = await runMasterDriveSync(new GraphClient());
      context.log(`Master report sync completed: ${result.moved} moved, ${result.skipped} skipped, ${result.failed} failed.`);
      return {
        status: result.failed > 0 ? 207 : 200,
        jsonBody: { status: result.failed > 0 ? 'completed-with-errors' : 'completed', ...result },
      };
    } catch (error) {
      return errorResponse(error);
    }
  },
});