import { app } from '@azure/functions';

import { GraphClient } from '../services/graphClient.js';
import { config } from '../config.js';
import { buildEmployeePayload, mapUsedRangeRows } from '../services/medicalDataService.js';
import { errorResponse, getPrincipal, HttpError } from '../services/requestAuth.js';
import { listAccessUsers } from '../services/accessControlService.js';

function graphErrorDetails(error: unknown): { status: number; message: string; code?: string; requestId?: string; upstreamStatusCode?: number } | undefined {
  if (!error || typeof error !== 'object') return undefined;
  const value = error as {
    statusCode?: unknown;
    status?: unknown;
    message?: unknown;
    code?: unknown;
    requestId?: unknown;
    body?: unknown;
    response?: { status?: unknown; body?: unknown };
  };
  const upstreamStatusCode = Number(value.statusCode ?? value.status ?? value.response?.status);
  const isIdentityAuthenticationError = value.code === 'AuthenticationRequiredError'
    || /invalid_client|AADSTS\d+/i.test(String(value.message ?? ''));
  if ((!Number.isInteger(upstreamStatusCode) || upstreamStatusCode < 400 || upstreamStatusCode > 599) && !isIdentityAuthenticationError) return undefined;

  let message = typeof value.message === 'string' ? value.message : 'Microsoft Graph request failed.';
  let bodyValue = value.body ?? value.response?.body;
  if (typeof bodyValue === 'string') {
    try {
      bodyValue = JSON.parse(bodyValue);
    } catch {
      // Keep the original SDK message when the upstream body is not JSON.
    }
  }
  if (bodyValue && typeof bodyValue === 'object') {
    const body = bodyValue as { error?: { message?: unknown } };
    if (typeof body.error?.message === 'string') message = body.error.message;
  }
  return {
    status: Number.isInteger(upstreamStatusCode) && upstreamStatusCode >= 400 && upstreamStatusCode <= 599
      ? upstreamStatusCode
      : 502,
    message,
    ...(typeof value.code === 'string' ? { code: value.code } : {}),
    ...(typeof value.requestId === 'string' ? { requestId: value.requestId } : {}),
    ...(Number.isInteger(upstreamStatusCode) && upstreamStatusCode >= 400 && upstreamStatusCode <= 599 ? { upstreamStatusCode } : {}),
  };
}

function graphConfigurationStatus(): string {
  const settings = {
    TENANT_ID: config.graphTenantId,
    CLIENT_ID: config.graphClientId,
    CLIENT_SECRET: config.graphClientSecret,
    SITE_ID: config.sharePointSiteId,
    DRIVE_ID: config.sharePointDriveId,
    EXCEL_FILE_ID: config.excelFileId,
  };
  return Object.entries(settings)
    .map(([name, value]) => `${name}=${value ? 'loaded' : 'missing'}`)
    .join(', ');
}

app.http('getEmployeeData', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'getEmployeeData',
  handler: async (request, context) => {
    context.log(`Graph configuration: ${graphConfigurationStatus()}`);
    context.log('Reading values from the first worksheet used range; no Excel Table is required.');
    try {
      const principal = getPrincipal(request);
      const graph = new GraphClient();
      const worksheet = await graph.getFirstWorksheetUsedRange();
      context.log(`Loaded used range from worksheet ${worksheet.worksheetId}.`);
      let rows = mapUsedRangeRows(worksheet.values);
      const isAdmin = principal.roles.some((role) => role.toLowerCase() === 'admin');
      if (!isAdmin) {
        const access = await listAccessUsers(graph);
        const user = access.users.find((candidate) => candidate.email.toLowerCase() === principal.userDetails.toLowerCase());
        if (!user) throw new HttpError(403, 'No branch access is assigned to this account.');
        const branches = new Set(user.branches.split(',').map((branch) => branch.trim().toLowerCase()).filter(Boolean));
        rows = rows.filter((row) => branches.has(String(row.values?.[0]?.[4] ?? '').trim().toLowerCase()));
      }
      context.log(`Returned ${rows.length} workbook rows to an authenticated dashboard request.`);
      return { status: 200, jsonBody: buildEmployeePayload(rows, principal.userDetails, isAdmin) };
    } catch (error) {
      console.error('getEmployeeData request failed:', error);
      const graphError = graphErrorDetails(error);
      if (graphError) {
        context.error(`Microsoft Graph request failed with HTTP ${graphError.status}: ${graphError.message}`);
        return {
          status: graphError.status,
          jsonBody: {
            error: {
              code: graphError.upstreamStatusCode === undefined ? 'MICROSOFT_GRAPH_AUTH_ERROR' : 'MICROSOFT_GRAPH_ERROR',
              message: graphError.message,
              ...(graphError.upstreamStatusCode !== undefined ? { upstreamStatusCode: graphError.upstreamStatusCode } : {}),
              ...(graphError.code ? { upstreamCode: graphError.code } : {}),
              ...(graphError.requestId ? { requestId: graphError.requestId } : {}),
            },
          },
        };
      }
      return errorResponse(error);
    }
  },
});