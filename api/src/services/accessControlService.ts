import { config } from '../config.js';
import type { GraphClient, GraphDriveItem, GraphPermission } from './graphClient.js';
import { mapMasterRow } from './medicalDataService.js';

type BranchFolder = { branch: string; region: string; id: string };
type AccessUser = { empId: string; email: string; name: string; branches: string; row: number };

function permissionUsers(permission: GraphPermission): Array<{ email: string; name: string }> {
  const principals = [
    permission.grantedToV2?.user,
    permission.grantedToV2?.siteUser,
    ...(permission.grantedToIdentitiesV2 ?? []).map((identity) => identity.user),
  ];
  return principals.flatMap((principal) => {
    const email = principal?.email?.trim().toLowerCase();
    return email ? [{ email, name: principal?.displayName ?? email }] : [];
  });
}

async function getExistingBranchFolders(graph: GraphClient): Promise<BranchFolder[]> {
  const regionFolders = (await graph.listChildren(config.organizedReportsFolderId)).filter((item) => item.folder);
  const result: BranchFolder[] = [];
  for (const regionFolder of regionFolders) {
    const branches = (await graph.listChildren(regionFolder.id)).filter((item) => item.folder);
    result.push(...branches.map((branch) => ({ branch: branch.name, region: regionFolder.name, id: branch.id })));
  }
  return result;
}

async function getWorkbookBranches(graph: GraphClient): Promise<Array<{ branch: string; region: string }>> {
  const rows = await graph.getMasterRows();
  const unique = new Map<string, { branch: string; region: string }>();
  for (const raw of rows) {
    const row = mapMasterRow(raw);
    if (row.branch && row.region) unique.set(`${row.region.toLowerCase()}\0${row.branch.toLowerCase()}`, { branch: row.branch, region: row.region });
  }
  return [...unique.values()];
}

export async function listAccessUsers(graph: GraphClient): Promise<{ users: AccessUser[]; branches: string[] }> {
  graph.assertAccessSettings();
  const [folders, workbookBranches] = await Promise.all([getExistingBranchFolders(graph), getWorkbookBranches(graph)]);
  const granted = new Map<string, { name: string; branches: Set<string> }>();
  for (const folder of folders) {
    const permissions = await graph.getItemPermissions(folder.id);
    for (const permission of permissions.filter((entry) => !entry.inheritedFrom)) {
      for (const principal of permissionUsers(permission)) {
        const user = granted.get(principal.email) ?? { name: principal.name, branches: new Set<string>() };
        user.branches.add(folder.branch);
        granted.set(principal.email, user);
      }
    }
  }
  const users = [...granted.entries()].sort(([left], [right]) => left.localeCompare(right)).map(([email, user], index) => ({
    empId: '',
    email,
    name: user.name,
    branches: [...user.branches].sort().join(', '),
    row: index + 1,
  }));
  return {
    users,
    branches: [...new Set([...folders.map((folder) => folder.branch), ...workbookBranches.map((branch) => branch.branch)])].sort(),
  };
}

export async function updateBranchAccess(graph: GraphClient, email: string, selectedBranches: string[]): Promise<void> {
  graph.assertAccessSettings();
  const normalizedEmail = email.trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalizedEmail)) throw new Error('A valid user email is required.');
  const validBranches = await getWorkbookBranches(graph);
  const selected = new Set(selectedBranches.map((branch) => branch.trim().toLowerCase()));
  const unknown = [...selected].filter((branch) => !validBranches.some((entry) => entry.branch.toLowerCase() === branch));
  if (unknown.length) throw new Error(`Unknown branch assignment: ${unknown.join(', ')}`);

  const folders = await getExistingBranchFolders(graph);
  for (const folder of folders) {
    const permissions = await graph.getItemPermissions(folder.id);
    const userPermissions = permissions.filter((permission) =>
      !permission.inheritedFrom && permissionUsers(permission).some((principal) => principal.email === normalizedEmail));
    const keep = selected.has(folder.branch.toLowerCase());
    if (!keep) {
      for (const permission of userPermissions) await graph.removeItemPermission(folder.id, permission.id);
    } else if (userPermissions.length === 0) {
      await graph.grantFolderAccess(folder.id, normalizedEmail);
    }
  }

  for (const pair of validBranches.filter((entry) => selected.has(entry.branch.toLowerCase()))) {
    const hasFolder = folders.some((folder) => folder.region.toLowerCase() === pair.region.toLowerCase() && folder.branch.toLowerCase() === pair.branch.toLowerCase());
    if (hasFolder) continue;
    const regionFolder = await graph.ensureChildFolder(config.organizedReportsFolderId, pair.region);
    const branchFolder = await graph.ensureChildFolder(regionFolder.id, pair.branch);
    await graph.grantFolderAccess(branchFolder.id, normalizedEmail);
  }
}