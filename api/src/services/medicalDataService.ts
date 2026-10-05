import type { GraphWorkbookRow } from './graphClient.js';

export type EmployeeRecord = {
  timestamp: string;
  empId: string;
  firstName: string;
  surname: string;
  branch: string;
  region: string;
  doj: string;
  department: string;
  designation: string;
  yearsOfService: string;
  status: string;
  lastCheckup: string;
  upcomingCheckup: string;
  employeeStatus: string;
  fitUnfit: string;
  hrRemarks: string;
  fileUrl: string;
  lastCheckupPeriod: string;
  checkupYear: string;
  rowIndex: number;
  values: unknown[];
};

const text = (value: unknown): string => value === null || value === undefined ? '' : String(value).trim();

export function normalizeExcelDate(value: unknown): string {
  if (typeof value === 'number' && Number.isFinite(value)) {
    const date = new Date(Date.UTC(1899, 11, 30) + value * 86_400_000);
    return Number.isNaN(date.getTime()) ? '' : new Intl.DateTimeFormat('en-GB', {
      day: '2-digit', month: 'short', year: 'numeric', timeZone: 'UTC',
    }).format(date).replace(',', '');
  }
  return text(value);
}

export function parseDateValue(value: unknown): Date | null {
  if (typeof value === 'number' && Number.isFinite(value)) {
    const date = new Date(Date.UTC(1899, 11, 30) + value * 86_400_000);
    return Number.isNaN(date.getTime()) ? null : date;
  }
  const clean = text(value);
  if (!clean || ['-', 'missing date', 'no record'].includes(clean.toLowerCase())) return null;
  const direct = new Date(clean);
  if (!Number.isNaN(direct.getTime())) return direct;
  const match = clean.match(/^(\d{1,2})\s+([A-Za-z]{3,})\s+(\d{4})$/);
  if (match) {
    const month = new Date(`${match[2].slice(0, 3)} 1, 2000`).getMonth();
    const date = new Date(Number(match[3]), month, Number(match[1]));
    return Number.isNaN(date.getTime()) ? null : date;
  }
  return null;
}

export function mapMasterRow(row: GraphWorkbookRow): EmployeeRecord {
  const values = Array.isArray(row.values?.[0]) ? row.values[0] : [];
  const lastCheckup = normalizeExcelDate(values[11]);
  const upcomingCheckup = normalizeExcelDate(values[12]);
  const lastDate = parseDateValue(values[11]);
  const futureDate = parseDateValue(values[12]);
  return {
    timestamp: normalizeExcelDate(values[0]),
    empId: text(values[1]),
    firstName: text(values[2]),
    surname: text(values[3]),
    branch: text(values[4]),
    region: text(values[5]),
    doj: normalizeExcelDate(values[6]),
    department: text(values[7]),
    designation: text(values[8]),
    yearsOfService: text(values[9]),
    status: text(values[10]),
    lastCheckup,
    upcomingCheckup,
    employeeStatus: text(values[13]),
    fitUnfit: text(values[14]),
    hrRemarks: text(values[15]),
    fileUrl: text(values[16]),
    lastCheckupPeriod: lastDate ? new Intl.DateTimeFormat('en-US', { month: 'short', year: 'numeric', timeZone: 'UTC' }).format(lastDate) : '',
    checkupYear: futureDate ? String(futureDate.getUTCFullYear()) : '',
    rowIndex: row.index,
    values,
  };
}

export function mapUsedRangeRows(values: unknown[][]): GraphWorkbookRow[] {
  return values.slice(1).map((row, index) => ({ index, values: [row] }));
}

export function buildEmployeePayload(rows: GraphWorkbookRow[], user: string, showManageAccess: boolean) {
  const records = rows.map(mapMasterRow).filter((record) => record.empId);
  return {
    rows: records,
    regions: [...new Set(records.map((record) => record.region).filter(Boolean))].sort(),
    branches: [...new Set(records.map((record) => record.branch).filter(Boolean))].sort(),
    departments: [...new Set(records.map((record) => record.department).filter(Boolean))].sort(),
    statuses: [...new Set(records.map((record) => record.status).filter(Boolean))].sort(),
    checkupYears: [...new Set(records.map((record) => record.checkupYear).filter(Boolean))].sort(),
    lastCheckupPeriods: [...new Set(records.map((record) => record.lastCheckupPeriod).filter(Boolean))].sort(),
    user,
    allMasterBranches: [...new Set(records.map((record) => record.branch).filter(Boolean))].sort(),
    showManageAccess,
  };
}