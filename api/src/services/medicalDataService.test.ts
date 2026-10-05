import { describe, expect, it } from 'vitest';

import { findEmployeeId, folderName, monthFolder } from './driveSyncService.js';
import { buildEmployeePayload, mapMasterRow, mapUsedRangeRows, parseDateValue } from './medicalDataService.js';

describe('MasterMedicalTable mapping', () => {
  it('maps the supplied 17 columns into the dashboard row contract', () => {
    const values = [
      'submitted', 'EMP-102', 'Ada', 'Lovelace', 'North Outlet', 'North', '2020-01-04',
      'Clinical', 'Nurse', 6, 'Active', '15 Mar 2026', '2027-03-15', 'Current', 'Fit',
      'Reviewed', 'https://reports.example/report.pdf',
    ];
    const row = mapMasterRow({ index: 4, values: [values] });
    expect(row).toMatchObject({
      empId: 'EMP-102', firstName: 'Ada', surname: 'Lovelace', branch: 'North Outlet', region: 'North',
      doj: '2020-01-04', department: 'Clinical', designation: 'Nurse', yearsOfService: '6',
      status: 'Active', lastCheckup: '15 Mar 2026', upcomingCheckup: '2027-03-15',
      employeeStatus: 'Current', fitUnfit: 'Fit', hrRemarks: 'Reviewed',
      fileUrl: 'https://reports.example/report.pdf', lastCheckupPeriod: 'Mar 2026', checkupYear: '2027', rowIndex: 4,
    });
    expect(buildEmployeePayload([{ index: 4, values: [values] }], 'user@example.com', false)).toMatchObject({
      user: 'user@example.com', showManageAccess: false, regions: ['North'], branches: ['North Outlet'],
      departments: ['Clinical'], statuses: ['Active'], checkupYears: ['2027'],
    });
  });

  it('skips the used-range header row and maps each remaining worksheet row', () => {
    const headers = ['Timestamp', 'EMP ID', 'First Name'];
    const employeeValues = ['submitted', 'EMP-102', 'Ada', 'Lovelace', 'North Outlet', 'North', '2020-01-04', 'Clinical', 'Nurse', 6, 'Active', '15 Mar 2026', '2027-03-15', 'Current', 'Fit', 'Reviewed', 'https://reports.example/report.pdf'];
    const mappedRows = mapUsedRangeRows([headers, employeeValues]);

    expect(mappedRows).toEqual([{ index: 0, values: [employeeValues] }]);
    expect(mapMasterRow(mappedRows[0]).empId).toBe('EMP-102');
  });

  it('parses Excel serial dates and rejects missing checkup values', () => {
    expect(parseDateValue(46082)?.toISOString().slice(0, 10)).toBe('2026-03-01');
    expect(parseDateValue('No Record')).toBeNull();
  });
});

describe('report folder routing', () => {
  it('uses the last-checkup month and falls back to New Employee for invalid dates', () => {
    expect(monthFolder('15 Mar 2026')).toBe('Mar 2026');
    expect(monthFolder('Missing Date')).toBe('New Employee');
  });

  it('matches complete EMP IDs in filenames without matching partial IDs', () => {
    const records = new Map([
      ['emp-12', { index: 0, values: [['', 'EMP-12']] }],
      ['emp-123', { index: 1, values: [['', 'EMP-123']] }],
    ]);
    expect(findEmployeeId('medical_EMP-123_2026.pdf', records)).toBe('emp-123');
    expect(findEmployeeId('medical_XEMP-12.pdf', records)).toBeUndefined();
  });

  it('replaces folder path separators and invalid characters', () => {
    expect(folderName('North/West', 'Unknown')).toBe('North-West');
    expect(folderName('  ', 'Unassigned')).toBe('Unassigned');
  });
});