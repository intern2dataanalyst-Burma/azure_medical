import { ClientSecretCredential } from '@azure/identity';
import type { TokenCredential } from '@azure/core-auth';
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';

import { config, requireGraphSettings } from '../config.js';

export type GraphDriveItem = {
  id: string;
  name: string;
  webUrl?: string;
  file?: Record<string, unknown>;
  folder?: Record<string, unknown>;
  parentReference?: { driveId?: string; id?: string };
};

export type GraphWorkbookRow = {
  index: number;
  values: unknown[][];
};

type GraphPage<T> = {
  value?: T[];
  '@odata.nextLink'?: string;
};

type GraphWorksheet = { id: string; name?: string };
type GraphUsedRange = { values?: unknown[][] };

type WorkbookSession = { id: string };
export type GraphPermission = {
  id: string;
  roles?: string[];
  inheritedFrom?: Record<string, unknown>;
  grantedToV2?: { user?: { id?: string; displayName?: string; email?: string }; siteUser?: { email?: string; displayName?: string } };
  grantedToIdentitiesV2?: Array<{ user?: { id?: string; displayName?: string; email?: string } }>;
};

export class GraphClient {
  private readonly client: Client;
  private readonly credential: TokenCredential;

  constructor() {
    requireGraphSettings(['graphTenantId', 'graphClientId', 'graphClientSecret']);
    this.credential = new ClientSecretCredential(config.graphTenantId, config.graphClientId, config.graphClientSecret);
    const authProvider = new TokenCredentialAuthenticationProvider(this.credential, {
      scopes: ['https://graph.microsoft.com/.default'],
    });
    this.client = Client.initWithMiddleware({ authProvider });
  }

  public assertWorkbookSettings(): void {
    requireGraphSettings(['sharePointSiteId', 'sharePointDriveId', 'excelFileId']);
  }

  public assertDriveSettings(): void {
    requireGraphSettings([
      'sharePointDriveId',
      'rawReportsFolderId',
      'organizedReportsFolderId',
    ]);
  }

  public assertAccessSettings(): void {
    requireGraphSettings(['sharePointDriveId', 'excelFileId', 'organizedReportsFolderId']);
  }

  public async createWorkbookSession(persistChanges: boolean): Promise<WorkbookSession> {
    this.assertWorkbookSettings();
    return this.client
      .api(this.workbookPath('/createSession'))
      .post({ persistChanges }) as Promise<WorkbookSession>;
  }

  public async closeWorkbookSession(sessionId: string): Promise<void> {
    await this.client.api(this.workbookPath('/closeSession')).header('workbook-session-id', sessionId).post({});
  }

  public async getMasterRows(sessionId?: string, workbookId = config.excelFileId): Promise<GraphWorkbookRow[]> {
    this.assertWorkbookSettings();
    const rows: GraphWorkbookRow[] = [];
    let request = this.client.api(this.workbookPath('/tables/MasterMedicalTable/rows', workbookId)).top(500);
    if (sessionId) request = request.header('workbook-session-id', sessionId);
    let page = await request.get() as GraphPage<GraphWorkbookRow>;
    while (true) {
      rows.push(...(page.value ?? []));
      const nextLink = page['@odata.nextLink'];
      if (!nextLink) return rows;
      const nextRequest = this.client.api(nextLink);
      if (sessionId) nextRequest.header('workbook-session-id', sessionId);
      page = await nextRequest.get() as GraphPage<GraphWorkbookRow>;
    }
  }

  public async getFirstWorksheetUsedRange(): Promise<{ worksheetId: string; values: unknown[][] }> {
    this.assertWorkbookSettings();
    const workbookPath = `/drives/${encodeURIComponent(config.sharePointDriveId)}/items/${encodeURIComponent(config.excelFileId)}/workbook`;
    const worksheets = await this.client.api(`${workbookPath}/worksheets`).get() as GraphPage<GraphWorksheet>;
    const firstWorksheet = worksheets.value?.[0];
    if (!firstWorksheet?.id) throw new Error('Microsoft Graph returned no worksheets for the configured workbook.');

    const usedRange = await this.client
      .api(`${workbookPath}/worksheets/${encodeURIComponent(firstWorksheet.id)}/usedRange(valuesOnly=true)`)
      .get() as GraphUsedRange;
    if (!Array.isArray(usedRange.values)) throw new Error('Microsoft Graph returned no values for the first worksheet used range.');
    return { worksheetId: firstWorksheet.id, values: usedRange.values };
  }

  public async updateReportLink(rowIndex: number, values: unknown[], fileUrl: string, sessionId: string): Promise<void> {
    this.assertWorkbookSettings();
    values[16] = fileUrl;
    await this.client
      .api(this.workbookPath(`/tables/MasterMedicalTable/rows/itemAt(index=${rowIndex})`))
      .header('workbook-session-id', sessionId)
      .patch({ values: [values] });
  }

  public async updateMasterRow(rowIndex: number, values: unknown[], sessionId: string): Promise<void> {
    this.assertWorkbookSettings();
    await this.client
      .api(this.workbookPath(`/tables/MasterMedicalTable/rows/itemAt(index=${rowIndex})`))
      .header('workbook-session-id', sessionId)
      .patch({ values: [values] });
  }

  public async appendMasterRows(rows: unknown[][], sessionId: string): Promise<void> {
    if (rows.length === 0) return;
    await this.client
      .api(this.workbookPath('/tables/MasterMedicalTable/rows/add'))
      .header('workbook-session-id', sessionId)
      .post({ index: null, values: rows });
  }

  public async listChildren(folderId: string): Promise<GraphDriveItem[]> {
    const path = `/drives/${encodeURIComponent(config.sharePointDriveId)}/items/${encodeURIComponent(folderId)}/children`;
    const items: GraphDriveItem[] = [];
    let page = await this.client.api(path).top(200).get() as GraphPage<GraphDriveItem>;
    while (true) {
      items.push(...(page.value ?? []));
      const nextLink = page['@odata.nextLink'];
      if (!nextLink) return items;
      page = await this.client.api(nextLink).get() as GraphPage<GraphDriveItem>;
    }
  }

  public async getItemPermissions(itemId: string): Promise<GraphPermission[]> {
    const response = await this.client
      .api(`/drives/${encodeURIComponent(config.sharePointDriveId)}/items/${encodeURIComponent(itemId)}/permissions`)
      .get() as GraphPage<GraphPermission>;
    return response.value ?? [];
  }

  public async grantFolderAccess(folderId: string, email: string): Promise<void> {
    await this.client
      .api(`/drives/${encodeURIComponent(config.sharePointDriveId)}/items/${encodeURIComponent(folderId)}/invite`)
      .post({
        recipients: [{ email }],
        message: '',
        requireSignIn: true,
        sendInvitation: false,
        roles: ['read'],
      });
  }

  public async removeItemPermission(itemId: string, permissionId: string): Promise<void> {
    await this.client
      .api(`/drives/${encodeURIComponent(config.sharePointDriveId)}/items/${encodeURIComponent(itemId)}/permissions/${encodeURIComponent(permissionId)}`)
      .delete();
  }

  public async listRawReports(): Promise<GraphDriveItem[]> {
    this.assertDriveSettings();
    const items: GraphDriveItem[] = [];
    let page = await this.client
      .api(`/drives/${encodeURIComponent(config.sharePointDriveId)}/items/${encodeURIComponent(config.rawReportsFolderId)}/children`)
      .top(200)
      .get() as GraphPage<GraphDriveItem>;
    while (true) {
      items.push(...(page.value ?? []));
      const nextLink = page['@odata.nextLink'];
      if (!nextLink) return items;
      page = await this.client.api(nextLink).get() as GraphPage<GraphDriveItem>;
    }
  }

  public async ensureChildFolder(parentId: string, name: string): Promise<GraphDriveItem> {
    const existing = await this.listChildren(parentId);
    const existingFolder = existing.find((item) => item.folder && item.name.toLocaleLowerCase() === name.toLocaleLowerCase());
    if (existingFolder) return existingFolder;

    try {
      return await this.client.api(`/drives/${encodeURIComponent(config.sharePointDriveId)}/items/${encodeURIComponent(parentId)}/children`).post({
        name,
        folder: {},
        '@microsoft.graph.conflictBehavior': 'fail',
      }) as GraphDriveItem;
    } catch (error) {
      const retry = await this.listChildren(parentId);
      const createdByAnotherRequest = retry.find((item) => item.folder && item.name.toLocaleLowerCase() === name.toLocaleLowerCase());
      if (createdByAnotherRequest) return createdByAnotherRequest;
      throw error;
    }
  }

  public async moveDriveItem(itemId: string, destinationFolderId: string): Promise<GraphDriveItem> {
    return this.client
      .api(`/drives/${encodeURIComponent(config.sharePointDriveId)}/items/${encodeURIComponent(itemId)}`)
      .patch({ parentReference: { driveId: config.sharePointDriveId, id: destinationFolderId } }) as Promise<GraphDriveItem>;
  }

  public async resolveSharedItem(shareUrl: string): Promise<GraphDriveItem> {
    const encoded = Buffer.from(shareUrl, 'utf8').toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
    return this.client.api(`/shares/u!${encoded}/driveItem`).get() as Promise<GraphDriveItem>;
  }

  public async downloadDriveItem(itemId: string): Promise<Buffer> {
    const token = await this.credential.getToken('https://graph.microsoft.com/.default');
    if (!token) throw new Error('Microsoft Graph authentication did not return an access token.');
    const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(config.sharePointDriveId)}/items/${encodeURIComponent(itemId)}/content`;
    const response = await fetch(url, { headers: { Authorization: `Bearer ${token.token}` } });
    if (!response.ok) throw new Error(`Microsoft Graph report download failed with HTTP ${response.status}.`);
    return Buffer.from(await response.arrayBuffer());
  }

  private workbookPath(suffix: string, workbookId = config.excelFileId): string {
    return `/sites/${encodeURIComponent(config.sharePointSiteId)}/drives/${encodeURIComponent(config.sharePointDriveId)}/items/${encodeURIComponent(workbookId)}/workbook${suffix}`;
  }
}