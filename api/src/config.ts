import dotenv from 'dotenv';

dotenv.config();

export const config = {
  nodeEnv: process.env.NODE_ENV ?? 'development',
  azureFunctionsEnvironment: process.env.AZURE_FUNCTIONS_ENVIRONMENT ?? 'Development',
  storageConnectionString: process.env.STORAGE_CONNECTION_STRING ?? 'UseDevelopmentStorage=true',
  graphTenantId: process.env.TENANT_ID ?? process.env.GRAPH_TENANT_ID ?? '',
  graphClientId: process.env.CLIENT_ID ?? process.env.GRAPH_CLIENT_ID ?? '',
  graphClientSecret: process.env.CLIENT_SECRET ?? process.env.GRAPH_CLIENT_SECRET ?? '',
  sharePointSiteId: process.env.SITE_ID ?? process.env.SHAREPOINT_SITE_ID ?? '',
  sharePointDriveId: process.env.DRIVE_ID ?? process.env.SHAREPOINT_DRIVE_ID ?? '',
  excelFileId: process.env.EXCEL_FILE_ID ?? '',
  rawReportsFolderId: process.env.RAW_REPORTS_FOLDER_ID ?? '',
  organizedReportsFolderId: process.env.ORGANIZED_REPORTS_FOLDER_ID ?? '',
  regionalSourceFileIds: (process.env.REGIONAL_SOURCE_FILE_IDS ?? '').split(',').map((id) => id.trim()).filter(Boolean),
};

export class ConfigurationError extends Error {}

export function validateConfig(required = ['STORAGE_CONNECTION_STRING']) {
  const missing = required.filter((key) => !process.env[key]);

  if (missing.length > 0 && config.nodeEnv !== 'development') {
    throw new Error(`Missing required env values for production: ${missing.join(', ')}`);
  }

  return { ok: true, missing };
}

export function requireGraphSettings(settings: string[]) {
  const missing = settings.filter((setting) => !config[setting as keyof typeof config]);
  if (missing.length > 0) {
    throw new ConfigurationError(`Missing required Graph settings: ${missing.join(', ')}`);
  }
}
