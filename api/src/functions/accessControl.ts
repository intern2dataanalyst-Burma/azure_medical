import { app } from '@azure/functions';
import { z } from 'zod';

import { GraphClient } from '../services/graphClient.js';
import { errorResponse, getPrincipal, requireAdmin } from '../services/requestAuth.js';
import { listAccessUsers, updateBranchAccess } from '../services/accessControlService.js';

const updateSchema = z.object({
  email: z.string().email(),
  branches: z.array(z.string().min(1)).max(200),
});

app.http('accessControl', {
  methods: ['GET', 'POST', 'PATCH'],
  authLevel: 'anonymous',
  route: 'accessControl',
  handler: async (request) => {
    try {
      const graph = new GraphClient();
      if (request.method === 'GET') {
        requireAdmin(request);
        return { status: 200, jsonBody: await listAccessUsers(graph) };
      }
      if (request.method === 'POST') {
        requireAdmin(request);
        const principal = getPrincipal(request);
        return { status: 200, jsonBody: { isValid: true, user: principal.userDetails } };
      }
      if (request.method === 'PATCH') {
        requireAdmin(request);
        const body = await request.json().catch(() => null);
        const parsed = updateSchema.safeParse(body);
        if (!parsed.success) {
          return { status: 400, jsonBody: { error: { code: 'INVALID_REQUEST', message: 'Provide a valid email and branch list.' } } };
        }
        await updateBranchAccess(graph, parsed.data.email, parsed.data.branches);
        return { status: 200, jsonBody: { success: true, email: parsed.data.email, branches: parsed.data.branches } };
      }
      return { status: 405, jsonBody: { error: { code: 'METHOD_NOT_ALLOWED', message: 'Method not allowed.' } } };
    } catch (error) {
      return errorResponse(error);
    }
  },
});