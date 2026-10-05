import { app } from '@azure/functions';

import './functions/employeeData.js';
import './functions/regionalSync.js';
import './functions/masterDriveSync.js';
import './functions/accessControl.js';
import './functions/exportReports.js';

export { app };
