// ============================================================
// SSPI MASTER DASHBOARD — Google Apps Script Backend
// File: SSPI_Code.gs
// Paste in: script.google.com → New Project (container-bound to SSPI_BACKEND sheet)
// Deploy: Web App → Execute as Me → Anyone can access
// Version 6 — + Approvals workflow, Sales Order/Dispatch pipeline,
//             email auto-draft capture, file upload to Drive
// ============================================================

const SHEET_NAME = 'SSPI_BACKEND';

// ── SECURITY: Shared Auth Token ──────────────────────────────
// Change this value ONCE after deploying. Update in dashboard + all bots.
// Keep it secret — treat like a password.
const AUTH_TOKEN = 'SSPI-VC-2026-SECURE';

// ── ALERT CONFIG ─────────────────────────────────────────────
// Email address(es) to notify on critical events (comma-separated)
const ALERT_EMAIL = 'ldpareek4@gmail.com';
// WhatsApp number for CallMeBot alerts (include country code, no +)
// Register free at: https://www.callmebot.com/blog/free-api-whatsapp-messages/
const WHATSAPP_NUMBER = '919XXXXXXXXX'; // Replace with actual number after CallMeBot registration
const CALLMEBOT_API_KEY = 'XXXXXXXX';   // Get from CallMeBot after registration

// ── Sheet tab names ──────────────────────────────────────────
const TABS = {
  jobCards:      'Job Cards',
  prodPlan:      'Production Plans',
  production:    'Production Log',
  quality:       'Quality NCR',
  accounts:      'Accounts',
  attendance:    'Attendance',
  stores:        'Stores',
  grn:           'Stores',   // GRN receipts → unified Stores movement log
  rm_grn:        'Stores',   // RM GRN receipts → same Stores tab
  maintenance:   'Maintenance',
  sales:         'Sales Orders',
  purchase:      'Purchase',
  debtors:       'Debtors Invoices',
  debtorsPay:    'Debtors Receipts',
  debtorsFU:     'Debtors Followup',
  billsSales:    'Bills Sales',
  billsPurchase: 'Bills Purchase',
  billsExpense:  'Bills Expense',
  salary:        'Salary Monthly',
  salaryRev:     'Salary Revision',
  employeeMaster:'Employee Master',
  esiEpf:        'ESI EPF Compliance',
  advance:       'Advance Ledger',
  gstPurchase:   'GST Purchase',
  gstSales:      'GST Sales',
  // Approvals workflow
  approvalRequest:    'Approvals',
  debtorsEscalation:  'Debtors Escalations',
  creditLimitChange:  'Credit Limit Changes',
  routeCard:          'Route Cards',
  substitution:        'RM Substitutions',
  // Sales Order → Production → Dispatch pipeline
  dispatch:      'Dispatch Register',
  // Pre-calculated payroll from Monthly Payroll Bot
  payroll:       'Payroll Register',
};

// ── Sheet alias map (for bots posting with {sheet:'name'}) ────
const SHEET_ALIAS = {
  'attendance_summary':  'attendance',
  'bills_sales':         'billsSales',
  'bills_purchase':      'billsPurchase',
  'bills_expense':       'billsExpense',
  'stores_grn':          'stores',
  'stores_issue':        'stores',
  'stores_adjustment':   'stores',
  'salary_monthly':      'salary',
  'salary_revision':     'salaryRev',
  'debtors_invoice':     'debtors',
  'debtors_receipt':     'debtorsPay',
  'debtors_followup':    'debtorsFU',
  'gst_purchase':        'gstPurchase',
  'gst_sales':           'gstSales',
};

// ── CORS handler ─────────────────────────────────────────────
function setCORS(output) {
  return output; // GAS handles CORS automatically on "Anyone can access" deployments
}
function doOptions() {
  return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.TEXT);
}

// ── Token validation ─────────────────────────────────────────
// Returns true if valid. Logs warning for invalid but does not hard-reject
// (set STRICT_AUTH = true to hard-reject after all clients are updated)
const STRICT_AUTH = true;
function validateToken(token) {
  if (!token) {
    Logger.log('AUTH WARNING: Request received without token');
    return !STRICT_AUTH; // if strict, fail; if not strict, allow
  }
  if (token !== AUTH_TOKEN) {
    Logger.log('AUTH REJECTED: Invalid token: ' + token);
    return false;
  }
  return true;
}

// ── Entry Point: POST (save data) ────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (!validateToken(data.authToken)) {
      return setCORS(ContentService
        .createTextOutput(JSON.stringify({ status: 'error', msg: 'Unauthorized' }))
        .setMimeType(ContentService.MimeType.JSON));
    }
    // ── File upload (PDF/XLSX/DOCX/image) — saves to Drive, returns URL ──
    if (data.type === 'uploadFile') {
      const url = uploadFileToDrive_(data.fileName, data.mimeType, data.base64Data, data.folderName);
      return setCORS(ContentService
        .createTextOutput(JSON.stringify({ status: 'ok', url: url }))
        .setMimeType(ContentService.MimeType.JSON));
    }
    // ── Approval decision (Approve/Reject) ──
    if (data.type === 'decideApproval') {
      const result = decideApproval_(data.requestId, data.decision, data.decidedBy, data.notes);
      return setCORS(ContentService
        .createTextOutput(JSON.stringify({ status: result.ok ? 'ok' : 'error', msg: result.msg || result.status }))
        .setMimeType(ContentService.MimeType.JSON));
    }
    const type = data.type || SHEET_ALIAS[data.sheet] || data.sheet;
    const result = saveData(type, data);
    // Duplicate detected — return warning, do NOT send alert
    if (String(result).startsWith('DUPLICATE:')) {
      return setCORS(ContentService
        .createTextOutput(JSON.stringify({ status: 'duplicate', msg: 'Entry already exists: ' + String(result).replace('DUPLICATE:','') }))
        .setMimeType(ContentService.MimeType.JSON));
    }
    // Send critical alert if needed (non-blocking — errors here don't fail the save)
    try { sendCriticalAlert(type, data); } catch(alertErr) { Logger.log('Alert error: ' + alertErr.message); }
    return setCORS(ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', id: result }))
      .setMimeType(ContentService.MimeType.JSON));
  } catch (err) {
    return setCORS(ContentService
      .createTextOutput(JSON.stringify({ status: 'error', msg: err.message }))
      .setMimeType(ContentService.MimeType.JSON));
  }
}

// ── Entry Point: GET (fetch data OR write via JSONP) ─────────
function doGet(e) {
  try {
    const callback = e.parameter.callback || '';
    const action   = e.parameter.action   || 'read';
    const token    = e.parameter.token || e.parameter.authToken || '';

    if (!validateToken(token)) {
      const errJson = JSON.stringify({ status: 'error', msg: 'Unauthorized' });
      if (callback) return ContentService.createTextOutput(callback + '(' + errJson + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
      return setCORS(ContentService.createTextOutput(errJson).setMimeType(ContentService.MimeType.JSON));
    }

    // ── JSONP Write path (action=write) — solves no-cors confirmation ──
    if (action === 'write') {
      const type = e.parameter.type || '';
      const rawD = e.parameter.d    || '{}';
      const data = JSON.parse(rawD);
      const result = saveData(type, data);
      // ── Duplicate check: block duplicate Job Cards, GRNs, SOs ──
      if (String(result).startsWith('DUPLICATE:')) {
        const dupKey = String(result).replace('DUPLICATE:','');
        const json = JSON.stringify({ status: 'duplicate', msg: 'Entry already exists: ' + dupKey, key: dupKey });
        if (callback) return ContentService.createTextOutput(callback + '(' + json + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
        return setCORS(ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON));
      }
      try { sendCriticalAlert(type, data); } catch(alertErr) {}
      const json = JSON.stringify({ status: 'ok', id: result, type: type });
      if (callback) return ContentService.createTextOutput(callback + '(' + json + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
      return setCORS(ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON));
    }

    // ── Approval decision via JSONP GET (no-cors safe for browser bots) ──
    if (action === 'decide') {
      const result = decideApproval_(e.parameter.requestId, e.parameter.decision, e.parameter.decidedBy, e.parameter.notes || '');
      const json = JSON.stringify({ status: result.ok ? 'ok' : 'error', msg: result.msg || result.status });
      if (callback) return ContentService.createTextOutput(callback + '(' + json + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
      return setCORS(ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON));
    }

    // ── Health check ping ──
    if (action === 'ping') {
      const json = JSON.stringify({ status: 'ok', ts: new Date().toISOString(), warm: true });
      if (callback) return ContentService.createTextOutput(callback + '(' + json + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
      return setCORS(ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON));
    }

    // ── Standard read ──
    const type  = e.parameter.type  || 'jobCards';
    const limit = parseInt(e.parameter.limit) || 100;
    const rows  = getData(type, limit);
    const json  = JSON.stringify({ status: 'ok', data: rows });
    if (callback) return ContentService.createTextOutput(callback + '(' + json + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
    return setCORS(ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON));

  } catch (err) {
    const errJson = JSON.stringify({ status: 'error', msg: err.message });
    const cb = (e.parameter && e.parameter.callback) || '';
    if (cb) return ContentService.createTextOutput(cb + '(' + errJson + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
    return setCORS(ContentService.createTextOutput(errJson).setMimeType(ContentService.MimeType.JSON));
  }
}

// ── Get or create tab ─────────────────────────────────────────
function getTab(tabName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    const headers = HEADERS[tabName];
    if (headers) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground('#1E3A5F').setFontColor('white').setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

// ── Headers per tab ───────────────────────────────────────────
const HEADERS = {
  'Job Cards':          ['JC No','Date','Customer','Product','Qty','Machine','Operator','Start','Due','Drawing','Priority','Status','Instructions','QC Points','Safety','Location','Submitted By','Timestamp','Linked SO No'],
  'Production Plans':   ['Plan No','Date','Product','Planned Qty','Period','Dept','Shifts','Status','Location','Submitted By','Timestamp','Linked SO No'],
  'Production Log':     ['Date','Shift','Target','Actual','Efficiency%','Machine','Operator','Downtime Mins','Remarks','Location','Submitted By','Timestamp'],
  'Quality NCR':        ['Date','NCR No','Product','Defect Type','Qty Affected','Root Cause','Action Taken','Status','Location','Submitted By','Timestamp'],
  'Accounts':           ['Date','Voucher Type','Party','Amount','Description','Payment Mode','Reference','Location','Submitted By','Timestamp'],
  'Attendance':         ['Date','Employee','Department','In Time','Out Time','Hours','Status','Location','Submitted By','Timestamp'],
  'Stores':             ['Date','GRN/Issue No','Item','Qty','Unit','Type','Supplier/Dept','Remarks','Location','Submitted By','Timestamp'],
  'Maintenance':        ['Date','Machine','Issue','Action Taken','Technician','Hours','Status','Next Due','Location','Submitted By','Timestamp'],
  'Sales Orders':       ['Date','SO No','Customer','Product','Qty','Rate','Amount','Delivery Date','Status','Location','Submitted By','Timestamp',
                          'Attachment URL','Source','Linked Job Card','Dispatched Qty','Balance Qty','Extracted Body Content'],
  'Purchase':           ['Date','PO No','Supplier','Item','Qty','Rate','Amount','Delivery Date','Status','Location','Submitted By','Timestamp'],
  'Debtors Invoices':   ['Invoice No','Date','Customer','Amount','Due Date','PO Ref','Status','Location','Submitted By','Timestamp'],
  'Debtors Receipts':   ['Date','Customer','Invoice No','Amount','TDS','Net','Mode','Ref No','Entered By','Status','Location','Submitted By','Timestamp'],
  'Debtors Followup':   ['Date','Customer','Outcome','Next Action','Location','Submitted By','Timestamp'],
  'Bills Sales':        ['Date','Bill No','Party','Total','GST Amt','Item Count','Location','Submitted By','Timestamp'],
  'Bills Purchase':     ['Date','Bill No','Party','Total','GST Amt','Item Count','Location','Submitted By','Timestamp'],
  'Bills Expense':      ['Date','Bill No','Party','Total','GST Amt','Item Count','Location','Submitted By','Timestamp'],
  // Salary — extended with auto-calc columns (Base Salary..Net Pay Final) ahead of Location
  'Salary Monthly':     ['Month','Emp Code','Employee','Working Days','Present Days','OT Hours','Advance','II Advance','Other Deduction','Remarks',
                          'Base Salary','Total Days','Gross Salary','Net Pay (Before PF/ESI — Accounts to deduct exact PF/ESI before payment)',
                          'PF Deduction','ESI Deduction','Net Pay (Final)','Location','Submitted By','Timestamp'],
  'Salary Revision':    ['Date','Emp Code','Employee','Old Rate','New Rate','Effective From','Remarks','Location','Submitted By','Timestamp'],
  'Employee Master':    ['Emp Code','Employee Name','Category','Base Salary','Salary Base Days','Status'],
  'ESI EPF Compliance': ['Month','Emp Code','Employee','ESI Employee','ESI Employer','EPF Employee','EPF Employer','Admin+EDLI','PF Wage','Location','Submitted By','Timestamp'],
  'Advance Ledger':     ['Date','Emp Code','Employee','Entry Type','Amount','Reason','Recovery Mode','Installments','Approved By','Balance After','Location','Submitted By','Timestamp'],
  'GST Purchase':       ['Date','Supplier GSTIN','Supplier Name','Invoice No','Taxable','IGST','CGST','SGST','Total','Location','Submitted By','Timestamp'],
  'GST Sales':          ['Date','Buyer GSTIN','Buyer Name','Invoice No','Taxable','IGST','CGST','SGST','Total','Location','Submitted By','Timestamp'],
  // ── Approvals workflow ──
  'Approvals':          ['Request ID','Type','Reference','Summary','Details JSON','Requested By','Status','Decided By','Decision Notes','Decision Timestamp','Location','Submitted By','Timestamp'],
  'Debtors Escalations':['Date','Customer','Amount','Days Overdue','Escalation Level','Action Taken','Next Step','Approval Request ID','Status','Location','Submitted By','Timestamp'],
  'Credit Limit Changes':['Date','Customer','Current Limit','Proposed Limit','Reason','Approval Request ID','Status','Location','Submitted By','Timestamp'],
  'Route Cards':        ['Product Code','Product Name','Department','Steps JSON','Approval Request ID','Status','Location','Submitted By','Timestamp'],
  'RM Substitutions':   ['Date','Job Card / Product','Original Item','Substitute Item','Reason','Approval Request ID','Status','Location','Submitted By','Timestamp'],
  // ── Sales Order → Dispatch ──
  'Dispatch Register':  ['Date','Dispatch No','SO No','Customer','Product','Qty Dispatched','Vehicle/Transporter','LR No','Driver','Destination','Location','Submitted By','Timestamp'],
};

// ── Employee Master — real SSPI headcount (21 active), used for ──
// ── live salary auto-calculation. Matches the client's original  ──
// ── SALARY_MASTER_SHEET.xlsx Employee Master tab. Salary Base     ──
// ── Days = 30 for all (Regular salary type, none on Fixed yet).   ──
const EMPLOYEE_MASTER_SEED = [
  [1,  'SATVEER SINGH',      'ESI', 30000, 30, 'Active'],
  [2,  'AMAR CHAND',         'ESI', 20000, 30, 'Active'],
  [3,  'NAND PRATAP',        'ESI', 20000, 30, 'Active'],
  [4,  'BHARAT/RAM ADHAR',   'ESI', 20000, 30, 'Active'],
  [5,  'RAJ KUMAR PASWAN',   'ESI', 19000, 30, 'Active'],
  [6,  'SACHINE S/O MANOJ',  '',    11000, 30, 'Active'],
  [7,  'SHYAM PASWAN',       'ESI', 13500, 30, 'Active'],
  [9,  'RAM VILAS',          'ESI', 13000, 30, 'Active'],
  [10, 'SATISH AGRA',        '',    13000, 30, 'Active'],
  [11, 'MANJEET',            'ESI', 19000, 30, 'Active'],
  [12, 'RAJEEV',             'ESI', 13500, 30, 'Active'],
  [13, 'PREM CHAND',         'ESI', 12000, 30, 'Active'],
  [14, 'MANOJ',              '',    13000, 30, 'Active'],
  [15, 'HARENDARA SINGH',    'ESI', 12000, 30, 'Active'],
  [16, 'SANJEEV',            '',    11000, 30, 'Active'],
  [17, 'VIKAS PRAJAPAT',     'ESI', 13000, 30, 'Active'],
  [18, 'JAIKISAN RAM',       '',    11000, 30, 'Active'],
  [19, 'SANJEET',            'ESI', 14000, 30, 'Active'],
  [20, 'HARI SHANKAR',       'ESI', 13000, 30, 'Active'],
  [21, 'VIRENDER',           'ESI', 11000, 30, 'Active'],
  [22, 'ANKIT YADAV',        '',    11500, 30, 'Active'],
  [24, 'SATYNARAN',          'ESI', 13500, 30, 'Active'],
];

// Creates the Employee Master tab (if missing) and seeds it once with
// the real 21-employee headcount, so the salary bot can auto-calculate
// Gross/Net Pay without anyone re-typing the master list into Sheets.
function ensureEmployeeMasterSeeded_() {
  const sheet = getTab('Employee Master');
  if (sheet.getLastRow() <= 1) {
    sheet.getRange(2, 1, EMPLOYEE_MASTER_SEED.length, 6).setValues(EMPLOYEE_MASTER_SEED);
  }
  return sheet;
}

// Looks up an employee's Base Salary + Salary Base Days by Emp Code.
// Returns {baseSalary, baseDays} — defaults to {0, 30} if not found.
function lookupEmployeeMaster_(empCode) {
  const sheet = ensureEmployeeMasterSeeded_();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { baseSalary: 0, baseDays: 30 };
  const rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === String(empCode)) {
      return { baseSalary: Number(rows[i][3]) || 0, baseDays: Number(rows[i][4]) || 30 };
    }
  }
  return { baseSalary: 0, baseDays: 30 };
}

// Replicates the client's original SALARY_MASTER_SHEET.xlsx formulas:
//   Total Days = Present Days + OT-hours-as-days
//   Gross      = (Base Salary / Salary Base Days) * Total Days
//   Net Pay    = Gross - Advance - II Advance - Other Deduction (minus PF/ESI)
// PF/ESI is intentionally excluded — in the real workbook it is a manual,
// compliance-driven figure (exact EPFO/ESI portal amount), not a formula.
function calcSalaryFields_(empCode, presentDays, otHours, advance, iiAdvance, otherDeduction) {
  const emp = lookupEmployeeMaster_(empCode);
  const present = Number(presentDays) || 0;
  const ot = Number(otHours) || 0;
  const otDays = ot === 0 ? 0 : (ot < 1 ? (ot * 24) / 8 : ot / 8);
  const totalDays = present + otDays;
  const perDayRate = emp.baseDays > 0 ? emp.baseSalary / emp.baseDays : 0;
  const gross = perDayRate * totalDays;
  const netPayBeforePF = gross - (Number(advance) || 0) - (Number(iiAdvance) || 0) - (Number(otherDeduction) || 0);
  return {
    baseSalary: emp.baseSalary,
    totalDays: Math.round(totalDays * 100) / 100,
    gross: Math.round(gross * 100) / 100,
    netPayBeforePF: Math.round(netPayBeforePF * 100) / 100,
  };
}

// Finds the most recent Salary Monthly row for a given Month + Emp Code
// and writes back PF Deduction / ESI Deduction / Net Pay (Final). Called
// from the ESIC & EPF Compliance Bot's "Push to Salary Sheet" button —
// this is what actually closes the loop left open by Net Pay (Before PF/ESI).
function patchSalaryRow_(month, empCode, pfDed, esiDed) {
  const sheet = getTab('Salary Monthly');
  const headers = HEADERS['Salary Monthly'];
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;

  const monthCol     = headers.indexOf('Month') + 1;
  const empCol       = headers.indexOf('Emp Code') + 1;
  const netBeforeCol = headers.indexOf('Net Pay (Before PF/ESI — Accounts to deduct exact PF/ESI before payment)') + 1;
  const pfCol        = headers.indexOf('PF Deduction') + 1;
  const esiCol       = headers.indexOf('ESI Deduction') + 1;
  const netFinalCol  = headers.indexOf('Net Pay (Final)') + 1;

  const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  for (let i = data.length - 1; i >= 0; i--) {  // search most-recent-first
    if (String(data[i][monthCol - 1]) === String(month) && String(data[i][empCol - 1]) === String(empCode)) {
      const rowIndex = i + 2; // actual sheet row number
      const netBefore = Number(data[i][netBeforeCol - 1]) || 0;
      const netFinal = Math.round((netBefore - (Number(pfDed) || 0) - (Number(esiDed) || 0)) * 100) / 100;
      sheet.getRange(rowIndex, pfCol).setValue(Number(pfDed) || 0);
      sheet.getRange(rowIndex, esiCol).setValue(Number(esiDed) || 0);
      sheet.getRange(rowIndex, netFinalCol).setValue(netFinal);
      return true;
    }
  }
  return false; // no matching Salary Monthly row yet — compliance row is still logged separately
}

// Computes an employee's current outstanding advance balance by summing
// every prior Advance Ledger entry (Disbursement adds, Recovery subtracts).
function computeAdvanceBalance_(empCode) {
  const sheet = getTab('Advance Ledger');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;
  const headers = HEADERS['Advance Ledger'];
  const empCol  = headers.indexOf('Emp Code');
  const typeCol = headers.indexOf('Entry Type');
  const amtCol  = headers.indexOf('Amount');
  const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  let balance = 0;
  data.forEach(r => {
    if (String(r[empCol]) === String(empCode)) {
      balance += (r[typeCol] === 'Recovery' ? -1 : 1) * (Number(r[amtCol]) || 0);
    }
  });
  return Math.round(balance * 100) / 100;
}

// ═══════════════════════════════════════════════════════════════
// APPROVALS WORKFLOW — request/decide pattern used by Advance,
// Debtors Escalation, Credit Limit Change, Route Card, Substitution
// ═══════════════════════════════════════════════════════════════

function genRequestId_() {
  return 'APR-' + Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyMMdd') + '-' + Math.floor(Math.random() * 9000 + 1000);
}

// Creates a Pending row in the Approvals tab. `data.detailsObj` (a plain
// object, NOT a string) holds everything needed to execute the underlying
// action once approved — it gets JSON-stringified into 'Details JSON'.
function createApprovalRequest_(type, reference, summary, detailsObj, requestedBy, location) {
  const sheet = getTab('Approvals');
  const reqId = genRequestId_();
  const ts = new Date().toLocaleString('en-IN');
  sheet.appendRow([reqId, type, reference, summary, JSON.stringify(detailsObj || {}), requestedBy || 'User',
                   'Pending', '', '', '', location || 'Bhiwadi', requestedBy || 'User', ts]);
  return reqId;
}

// Approves or rejects a pending request. On Approved, executes the
// underlying side-effect (advance ledger entry, escalation row, credit
// limit row, route card activation, substitution log, etc.) by replaying
// the stored Details JSON through saveData() with the request's original type.
function decideApproval_(requestId, decision, decidedBy, notes) {
  const sheet = getTab('Approvals');
  const headers = HEADERS['Approvals'];
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { ok: false, msg: 'No approval requests yet' };

  const idCol     = headers.indexOf('Request ID') + 1;
  const typeCol   = headers.indexOf('Type') + 1;
  const detailCol = headers.indexOf('Details JSON') + 1;
  const statusCol = headers.indexOf('Status') + 1;
  const decByCol  = headers.indexOf('Decided By') + 1;
  const notesCol  = headers.indexOf('Decision Notes') + 1;
  const decTsCol  = headers.indexOf('Decision Timestamp') + 1;

  const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][idCol - 1]) === String(requestId)) {
      const rowIndex = i + 2;
      const status = String(data[i][statusCol - 1] || '');
      if (status !== 'Pending') return { ok: false, msg: 'Request already ' + status };

      const decisionLabel = decision === 'Approved' ? 'Approved' : 'Rejected';
      const ts = new Date().toLocaleString('en-IN');
      sheet.getRange(rowIndex, statusCol).setValue(decisionLabel);
      sheet.getRange(rowIndex, decByCol).setValue(decidedBy || 'Approver');
      sheet.getRange(rowIndex, notesCol).setValue(notes || '');
      sheet.getRange(rowIndex, decTsCol).setValue(ts);

      if (decisionLabel === 'Approved') {
        const reqType = String(data[i][typeCol - 1]);
        let details = {};
        try { details = JSON.parse(data[i][detailCol - 1] || '{}'); } catch (e) {}
        details.approvalRequestId = requestId;
        try { saveData(reqType, details); } catch (execErr) {
          Logger.log('Approval execution error for ' + requestId + ': ' + execErr.message);
        }
      }
      return { ok: true, status: decisionLabel };
    }
  }
  return { ok: false, msg: 'Request ID not found' };
}

// ═══════════════════════════════════════════════════════════════
// SALES ORDER → PRODUCTION → DISPATCH LINKING
// ═══════════════════════════════════════════════════════════════

// Finds the most recent Sales Order row by SO No and patches Status /
// Linked Job Card / Dispatched Qty / Balance Qty. Called when a Job Card
// is created against an SO, or when a Dispatch entry is logged.
function updateSalesOrderStatus_(soNo, newStatus, extra) {
  if (!soNo) return false;
  const sheet = getTab('Sales Orders');
  const headers = HEADERS['Sales Orders'];
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;

  const soCol       = headers.indexOf('SO No') + 1;
  const qtyCol      = headers.indexOf('Qty') + 1;
  const statusCol   = headers.indexOf('Status') + 1;
  const jcCol       = headers.indexOf('Linked Job Card') + 1;
  const dispQtyCol  = headers.indexOf('Dispatched Qty') + 1;
  const balQtyCol   = headers.indexOf('Balance Qty') + 1;

  const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (String(data[i][soCol - 1]) === String(soNo)) {
      const rowIndex = i + 2;
      if (newStatus) sheet.getRange(rowIndex, statusCol).setValue(newStatus);
      if (extra && extra.linkedJobCard) sheet.getRange(rowIndex, jcCol).setValue(extra.linkedJobCard);
      if (extra && extra.dispatchedQtyDelta) {
        const orderedQty = Number(data[i][qtyCol - 1]) || 0;
        const priorDispatched = Number(data[i][dispQtyCol - 1]) || 0;
        const newDispatched = priorDispatched + Number(extra.dispatchedQtyDelta);
        const balance = Math.max(0, orderedQty - newDispatched);
        sheet.getRange(rowIndex, dispQtyCol).setValue(newDispatched);
        sheet.getRange(rowIndex, balQtyCol).setValue(balance);
        if (!newStatus) sheet.getRange(rowIndex, statusCol).setValue(balance <= 0 ? 'Dispatched' : 'Partially Dispatched');
      }
      return true;
    }
  }
  return false; // SO No not found — nothing to patch, not fatal
}

// ═══════════════════════════════════════════════════════════════
// FILE UPLOAD — saves a base64 file (PDF/XLSX/DOCX/image) from a bot
// into a shared Drive folder and returns its shareable URL.
// ═══════════════════════════════════════════════════════════════
function uploadFileToDrive_(fileName, mimeType, base64Data, folderName) {
  const folders = DriveApp.getFoldersByName(folderName || 'SSPI Sales Order Attachments');
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName || 'SSPI Sales Order Attachments');
  const bytes = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(bytes, mimeType || 'application/octet-stream', fileName || 'attachment');
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// ═══════════════════════════════════════════════════════════════
// EMAIL BODY CONTENT EXTRACTION — when an order/projection email has
// NO file attachment (the customer pasted a table straight into the
// email body instead of attaching a PDF/XLSX), this pulls that table
// out of the HTML body so the operator doesn't have to retype it from
// Gmail by hand. Falls back to a plain-text snippet if no <table> is
// found. Returns a JSON string: {type:'table',headers:[...],rows:[[...]]}
// or {type:'text', text:'...'} — stored in "Extracted Body Content".
// ═══════════════════════════════════════════════════════════════
function extractBodyContent_(msg) {
  try {
    let html = msg.getBody() || '';
    html = html.replace(/<script[\s\S]*?<\/script>/gi, '').replace(/<style[\s\S]*?<\/style>/gi, '');

    // Find every <table>...</table> block and parse each one — pick whichever
    // has the most data rows (layout-only wrapper tables usually have just 1,
    // so the real data table wins even if it's nested inside a layout table).
    const tableRegex = /<table[^>]*>([\s\S]*?)<\/table>/gi;
    let best = null;
    let tableBlockMatch;
    while ((tableBlockMatch = tableRegex.exec(html)) !== null) {
      const rowRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
      const rows = [];
      let rowMatch;
      while ((rowMatch = rowRegex.exec(tableBlockMatch[1])) !== null) {
        const cellRegex = /<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi;
        const cells = [];
        let cellMatch;
        while ((cellMatch = cellRegex.exec(rowMatch[1])) !== null) {
          const text = cellMatch[1]
            .replace(/<br\s*\/?>/gi, ' ')
            .replace(/<[^>]+>/g, '')
            .replace(/&nbsp;/gi, ' ').replace(/&amp;/gi, '&')
            .replace(/&lt;/gi, '<').replace(/&gt;/gi, '>').replace(/&#39;/gi, "'")
            .replace(/\s+/g, ' ').trim();
          cells.push(text);
        }
        if (cells.some(c => c !== '')) rows.push(cells);
      }
      if (rows.length > 1 && (!best || rows.length > best.length)) best = rows;
    }
    if (best) {
      return JSON.stringify({ type: 'table', headers: best[0], rows: best.slice(1) });
    }

    // No usable table — fall back to a plain-text snippet of the body
    let text = (msg.getPlainBody() || '').trim();
    if (!text) {
      // Some HTML-only forwards have no plain-text MIME part — strip tags ourselves
      text = html.replace(/<[^>]+>/g, ' ').replace(/&nbsp;/gi, ' ')
                 .replace(/&amp;/gi, '&').replace(/\s+/g, ' ').trim();
    }
    text = text.substring(0, 1500);
    if (text) return JSON.stringify({ type: 'text', text: text });
  } catch (e) {
    Logger.log('extractBodyContent_ failed: ' + e.message);
  }
  return '';
}

// ═══════════════════════════════════════════════════════════════
// EMAIL AUTO-DRAFT CAPTURE — scans a Gmail label for new order emails,
// saves the first attachment to Drive, and creates a Draft Sales Order
// row for a human to complete. Run on a 30-min trigger (setupAllTriggers).
// SETUP: create a Gmail label called "SSPI-Sales-Orders" and apply it
// (manually, or via a Gmail filter) to incoming PO / order emails.
// If an email has NO attachment, any HTML table pasted into the body
// (e.g. a projection sheet) is captured automatically — see
// extractBodyContent_() above — so the data isn't lost either way.
// ═══════════════════════════════════════════════════════════════
function scanSalesOrderEmails_() {
  const LABEL_NAME = 'SSPI-Sales-Orders';
  const PROCESSED_LABEL_NAME = 'SSPI-Sales-Orders-Processed';
  let label, processedLabel;
  try {
    label = GmailApp.getUserLabelByName(LABEL_NAME);
    if (!label) { Logger.log('Label "' + LABEL_NAME + '" not found — skipping email scan.'); return 'Label not found'; }
    processedLabel = GmailApp.getUserLabelByName(PROCESSED_LABEL_NAME) || GmailApp.createLabel(PROCESSED_LABEL_NAME);
  } catch (e) { Logger.log('Gmail access error: ' + e.message); return 'Gmail error'; }

  const threads = label.getThreads(0, 20);
  let draftsCreated = 0;
  threads.forEach(thread => {
    const messages = thread.getMessages();
    const msg = messages[messages.length - 1]; // latest message in thread
    const alreadyProcessed = thread.getLabels().some(l => l.getName() === PROCESSED_LABEL_NAME);
    if (alreadyProcessed) return;

    const sender = msg.getFrom();
    const subject = msg.getSubject();
    const dateReceived = msg.getDate();
    let attachmentUrl = '';
    try {
      const attachments = msg.getAttachments();
      if (attachments.length > 0) {
        attachmentUrl = uploadFileToDrive_(attachments[0].getName(), attachments[0].getContentType(),
                                            Utilities.base64Encode(attachments[0].getBytes()),
                                            'SSPI Sales Order Attachments');
      }
    } catch (e) { Logger.log('Attachment save failed: ' + e.message); }

    // Always extract body content — attachments may be logos/images, not order data.
    let bodyExtract = extractBodyContent_(msg);

    // Parse total qty from the last numeric column of the extracted table
    let totalQty = '';
    if (bodyExtract) {
      try {
        const parsed = JSON.parse(bodyExtract);
        if (parsed.type === 'table' && parsed.rows && parsed.rows.length > 0) {
          const lastCol = parsed.rows[0].length - 1;
          let sum = 0;
          parsed.rows.forEach(function(row) {
            const val = parseFloat(String(row[lastCol]).replace(/,/g, ''));
            if (!isNaN(val)) sum += val;
          });
          if (sum > 0) totalQty = sum;
        }
      } catch(e) { Logger.log('Qty parse error: ' + e.message); }
    }

    const draftSoNo = 'SO-' + Utilities.formatDate(dateReceived, 'Asia/Kolkata', 'yyMMdd-HHmmss');
    saveData('sales', {
      date: Utilities.formatDate(dateReceived, 'Asia/Kolkata', 'yyyy-MM-dd'),
      soNo: draftSoNo,
      customer: sender,
      product: subject || '(see attachment)',
      qty: totalQty, rate: '', amount: '',
      deliveryDate: '',
      status: 'Draft - Needs Review',
      attachmentUrl: attachmentUrl,
      bodyExtract: bodyExtract,
      source: 'Email-Draft',
      submittedBy: 'Email Auto-Capture'
    });
    thread.addLabel(processedLabel);
    draftsCreated++;
  });
  Logger.log('Email scan: ' + draftsCreated + ' draft Sales Order(s) created');
  return draftsCreated + ' draft(s) created';
}

// ── Save data to correct tab ──────────────────────────────────
function saveData(type, data) {
  const tabName = TABS[type];
  if (!tabName) throw new Error('Unknown type: ' + type);
  const sheet = getTab(tabName);
  const ts = new Date().toLocaleString('en-IN');
  const submittedBy = data.submittedBy || data.enteredBy || 'User';
  const location = data.location || 'Bhiwadi';
  let row = [];

  switch (type) {
    case 'jobCards':
      row = [data.no, data.date, data.customer, data.product, data.qty,
             data.machine, data.operator, data.start, data.due, data.drawing,
             data.priority, 'Open', data.instructions, data.qcPoints,
             data.safety, location, submittedBy, ts, data.linkedSO || ''];
      if (data.linkedSO) updateSalesOrderStatus_(data.linkedSO, 'In Production', { linkedJobCard: data.no });
      break;
    case 'prodPlan':
      row = [data.planNo, data.date, data.product, data.qty, data.period,
             data.dept, data.shifts, 'Active', location, submittedBy, ts, data.linkedSO || ''];
      break;
    case 'production':
      row = [data.date, data.shift, data.target, data.actual, data.efficiency,
             data.machine, data.operator, data.downtime, data.remarks, location, submittedBy, ts];
      break;
    case 'quality':
      row = [data.date, data.ncrNo, data.product, data.defectType, data.qtyAffected,
             data.rootCause, data.action, 'Open', location, submittedBy, ts];
      break;
    case 'accounts':
      row = [data.date, data.voucherType, data.party, data.amount,
             data.description, data.paymentMode, data.reference, location, submittedBy, ts];
      break;
    case 'attendance':
      row = [data.date, data.employee || data.empName, data.department || data.dept,
             data.inTime || '', data.outTime || '', data.hours || data.otHours || '',
             data.status || 'Present', location, submittedBy, ts];
      break;
    case 'stores':
      row = [data.date, data.refNo || data.docno, data.item,
             data.qty, data.unit, data.transType || data.type || data.sheet,
             data.party || data.from || data.to, data.remarks, location, submittedBy, ts];
      break;
    case 'grn':
      // GRN from Materials Bot → unified Stores tab
      // Stores cols: Date | GRN/Issue No | Item | Qty | Unit | Type | Supplier/Dept | Remarks | Location | Submitted By | Timestamp
      row = [data.date, data.grnNo || data.refNo, data.material || data.item,
             data.qty, data.unit || '', 'GRN',
             data.supplier || data.party, data.poRef || data.remarks || '', location, submittedBy, ts];
      break;
    case 'rm_grn':
      // RM GRN from Materials Bot → same Stores tab, Type = 'RM GRN'
      row = [data.date, data.grnNo || data.refNo, data.material || data.item,
             data.qty, data.unit || '', 'RM GRN',
             data.supplier || data.party, data.poRef || data.remarks || '', location, submittedBy, ts];
      break;
    case 'maintenance':
      row = [data.date, data.machine, data.issue, data.action, data.technician,
             data.hours, data.status, data.nextDue, location, submittedBy, ts];
      break;
    case 'sales':
      row = [data.date, data.soNo, data.customer, data.product, data.qty,
             data.rate, data.amount, data.deliveryDate, data.status || 'Order Received', location, submittedBy, ts,
             data.attachmentUrl || '', data.source || 'Manual', '', 0, Number(data.qty) || 0, data.bodyExtract || ''];
      break;
    // ── Dispatch — logs the shipment and patches the linked Sales Order ──
    case 'dispatch': {
      row = [data.date, data.dispatchNo, data.soNo, data.customer, data.product,
             data.qtyDispatched, data.vehicle || '', data.lrNo || '', data.driver || '',
             data.destination || '', location, submittedBy, ts];
      if (data.soNo) updateSalesOrderStatus_(data.soNo, '', { dispatchedQtyDelta: Number(data.qtyDispatched) || 0 });
      break;
    }
    // ── Approvals workflow: create request (Pending) ──
    case 'approvalRequest': {
      const reqId = createApprovalRequest_(data.requestType, data.reference, data.summary, data.detailsObj, submittedBy, location);
      return 'REQ:' + reqId; // returned as the "id" to the caller instead of a row number
    }
    // ── Executed automatically on Approved decision (not called directly by bots) ──
    case 'debtorsEscalation':
      row = [data.date || new Date().toLocaleDateString('en-IN'), data.customer, data.amount, data.daysOverdue || '',
             data.escalationLevel || '', data.actionTaken || '', data.nextStep || '', data.approvalRequestId || '',
             'Approved', location, submittedBy, ts];
      break;
    case 'creditLimitChange':
      row = [data.date || new Date().toLocaleDateString('en-IN'), data.customer, data.currentLimit || '', data.proposedLimit || '',
             data.reason || '', data.approvalRequestId || '', 'Approved', location, submittedBy, ts];
      break;
    case 'routeCard':
      row = [data.prodCode, data.prodName, data.dept, JSON.stringify(data.steps || []), data.approvalRequestId || '',
             'Active', location, submittedBy, ts];
      break;
    case 'substitution':
      row = [data.date || new Date().toLocaleDateString('en-IN'), data.jobOrProduct || '', data.origItem, data.subItem,
             data.reason || '', data.approvalRequestId || '', 'Approved', location, submittedBy, ts];
      break;
    case 'purchase':
      row = [data.date, data.poNo, data.supplier, data.item, data.qty,
             data.rate, data.amount, data.deliveryDate, 'Pending', location, submittedBy, ts];
      break;
    case 'debtors':
      row = [data.invoiceNo, data.date, data.customer, data.amount,
             data.dueDate, data.poRef || '', data.status || 'Open', location, submittedBy, ts];
      break;
    case 'debtorsPay':
      row = [data.date, data.customer, data.invoiceNo, data.amount, data.tds || 0,
             data.net || data.amount, data.mode, data.refNo || '', data.enteredBy || submittedBy,
             data.status || 'Pending Verify', location, submittedBy, ts];
      break;
    case 'debtorsFU':
      row = [data.date || new Date().toLocaleDateString('en-IN'), data.customer,
             data.outcome, data.next || '', location, submittedBy, ts];
      break;
    case 'billsSales':
    case 'billsPurchase':
    case 'billsExpense':
      row = [data.date, data.billNo || '', data.party, data.total,
             data.gstAmt || 0, data.itemCount || 0, location, submittedBy, ts];
      break;
    // ── Salary — auto-calculates Gross/Net Pay from Employee Master ──
    case 'salary': {
      const calc = calcSalaryFields_(data.empCode, data.presentDays, data.otHours,
                                      data.advance || 0, data.iiAdvance || 0, data.otherDeduction || 0);
      row = [data.month, data.empCode, data.employee || data.empName,
             data.workingDays, data.presentDays, data.otHours,
             data.advance || 0, data.iiAdvance || 0, data.otherDeduction || 0,
             data.remarks || '', calc.baseSalary, calc.totalDays, calc.gross, calc.netPayBeforePF,
             location, submittedBy, ts];
      break;
    }
    // ── Pre-calculated payroll from Monthly Payroll Bot ──
    case 'payroll':
      row = [data.period, data.empCode, data.empName, data.dept,
             data.workDays, data.gross, data.esic, data.epf,
             data.advance || 0, data.netPay, data.mode || '',
             location, submittedBy, ts];
      break;
    case 'salaryRev':
      row = [data.date, data.empCode, data.employee, data.oldRate, data.newRate,
             data.effectiveFrom, data.remarks || '', location, submittedBy, ts];
      break;
    // ── ESI / EPF Compliance — patches PF/ESI Deduction + Net Pay Final back onto Salary Monthly ──
    case 'esiEpf':
      row = [data.month, data.empCode, data.employee || '', data.esiEmployee || 0, data.esiEmployer || 0,
             data.epfEmployee || 0, data.epfEmployer || 0, data.adminEdli || 0, data.pfWage || 0,
             location, submittedBy, ts];
      patchSalaryRow_(data.month, data.empCode, data.epfEmployee || 0, data.esiEmployee || 0);
      break;
    // ── Advance Against Salary — tracks running balance per employee ──
    case 'advance': {
      const priorBalance = computeAdvanceBalance_(data.empCode);
      const amt = Number(data.amount) || 0;
      const delta = data.entryType === 'Recovery' ? -amt : amt;
      const balanceAfter = Math.round((priorBalance + delta) * 100) / 100;
      row = [data.date, data.empCode, data.employee || '', data.entryType || 'Disbursement', amt,
             data.reason || '', data.recoveryMode || '', data.installments || '', data.approvedBy || '',
             balanceAfter, location, submittedBy, ts];
      break;
    }
    case 'gstPurchase':
      row = [data.date, data.gstin || data.p_gstin, data.name || data.p_name,
             data.inv || data.p_inv, data.taxable || data.p_taxable,
             data.igst || 0, data.cgst || 0, data.sgst || 0,
             (parseFloat(data.taxable||0)+parseFloat(data.igst||0)+parseFloat(data.cgst||0)+parseFloat(data.sgst||0)),
             location, submittedBy, ts];
      break;
    case 'gstSales':
      row = [data.date, data.gstin || data.s_gstin, data.name || data.s_name,
             data.inv || data.s_inv, data.taxable || data.s_taxable,
             data.igst || 0, data.cgst || 0, data.sgst || 0,
             (parseFloat(data.taxable||0)+parseFloat(data.igst||0)+parseFloat(data.cgst||0)+parseFloat(data.sgst||0)),
             location, submittedBy, ts];
      break;
    default:
      row = [JSON.stringify(data), location, submittedBy, ts];
  }

  // ── Duplicate Prevention: check key column before appending ──
  const dupColMap = {
    jobCards:      { col: 1, key: data.no },                         // Job Card No
    grn:           { col: 2, key: data.grnNo || data.refNo },        // GRN No
    rm_grn:        { col: 2, key: data.grnNo || data.refNo },        // RM GRN No
    sales:         { col: 1, key: data.soNo || data.no },            // SO No
    salary:        { col: null, keys: [data.empCode, data.month] },  // Emp+Month combo
    attendance:    { col: null, keys: [data.empName || data.employee, data.date] }, // Emp+Date
    prodPlan:      { col: 1, key: data.planNo || data.no },          // Production Plan No
    quality:       { col: 2, key: data.ncrNo || data.no },           // NCR No
    purchase:      { col: 2, key: data.poNo || data.no },            // PO No
    routeCard:     { col: 1, key: data.routeCardId || data.no },     // Route Card ID
    fixedAssets:   { col: 1, key: data.assetId || data.no },         // Asset ID
    billsSales:    { col: 2, key: data.billNo || data.inv || data.no }, // Bill No
    billsPurchase: { col: 2, key: data.billNo || data.inv || data.no }, // Bill No
    billsExpense:  { col: 2, key: data.billNo || data.inv || data.no }, // Bill No
    gstPurchase:   { col: 4, key: data.inv || data.p_inv || data.invoiceNo }, // GST Invoice No
    gstSales:      { col: 4, key: data.inv || data.s_inv || data.invoiceNo }, // GST Invoice No
    debtors:       { col: 1, key: data.invoiceNo || data.invNo || data.no },  // Invoice No
    esiEpf:        { col: null, keys: [data.empCode, data.month] },  // Emp+Month combo
    advance:       { col: null, keys: [data.empCode || data.employee, data.date || data.month] }, // Emp+Date
  };
  const dupCheck = dupColMap[type];
  if (dupCheck && isDuplicate_(sheet, dupCheck)) {
    return 'DUPLICATE:' + (dupCheck.key || (dupCheck.keys||[]).join('|'));
  }

  sheet.appendRow(row);
  return sheet.getLastRow();
}

// ── Duplicate checker ─────────────────────────────────────────
function isDuplicate_(sheet, dupCheck) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return false;
    const numCols = dupCheck.col ? dupCheck.col : (dupCheck.keys ? 5 : 1);
    const data = sheet.getRange(2, 1, lastRow - 1, Math.max(numCols, dupCheck.keys ? 5 : dupCheck.col)).getValues();
    if (dupCheck.col && dupCheck.key) {
      const keyVal = String(dupCheck.key).trim().toLowerCase();
      return data.some(row => String(row[dupCheck.col - 1]).trim().toLowerCase() === keyVal);
    }
    if (dupCheck.keys && dupCheck.keys.length === 2) {
      const k1 = String(dupCheck.keys[0] || '').trim().toLowerCase();
      const k2 = String(dupCheck.keys[1] || '').trim().toLowerCase();
      if (!k1 || !k2) return false;
      return data.some(row =>
        String(row[0]).trim().toLowerCase() === k1 &&
        String(row[1]).trim().toLowerCase() === k2
      );
    }
    return false;
  } catch(e) {
    Logger.log('isDuplicate_ error: ' + e.message);
    return false; // fail open — allow write if check errors
  }
}

// ── Fetch last N rows from tab ────────────────────────────────
function getData(type, limit) {
  const tabName = TABS[type];
  if (!tabName) throw new Error('Unknown type: ' + type);
  const sheet = getTab(tabName);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const startRow = Math.max(2, lastRow - limit + 1);
  const numRows  = lastRow - startRow + 1;
  const numCols  = sheet.getLastColumn();
  const values   = sheet.getRange(startRow, 1, numRows, numCols).getValues();
  const headers  = HEADERS[tabName] || [];
  return values.reverse().map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] || ''; });
    return obj;
  });
}

// ═══════════════════════════════════════════════════════════════
// CRITICAL ALERTS — Email + WhatsApp via CallMeBot (FREE)
// ═══════════════════════════════════════════════════════════════

/**
 * Sends email + WhatsApp alert for critical events.
 * Called automatically from doPost and JSONP write path.
 *
 * HOW TO ENABLE WHATSAPP (FREE):
 * 1. Send "I allow callmebot to send me messages" to +34 644 59 22 99 on WhatsApp
 * 2. You'll receive your API key in a reply
 * 3. Update WHATSAPP_NUMBER and CALLMEBOT_API_KEY constants above
 */
function sendCriticalAlert(type, data) {
  const ts = new Date().toLocaleString('en-IN');
  const loc = data.location || 'Bhiwadi';
  let subject = '', body = '', whatsappMsg = '';

  if (type === 'maintenance' && data.issue) {
    subject = '🔴 SSPI Machine Alert: ' + (data.machine || 'Unknown Machine');
    body = 'Machine breakdown / maintenance entry recorded.\n\nMachine: ' + data.machine +
           '\nIssue: ' + data.issue + '\nTechnician: ' + (data.technician||'Not assigned') +
           '\nStatus: ' + (data.status||'Open') + '\nLocation: ' + loc +
           '\nEntered by: ' + (data.submittedBy||'User') + '\nTime: ' + ts;
    whatsappMsg = '🔴 SSPI Machine Alert [' + loc + ']: ' + (data.machine||'Machine') + ' - ' + (data.issue||'Issue reported') + '. Time: ' + ts;
  }
  else if (type === 'quality' && data.ncrNo) {
    subject = '⚠ SSPI Quality NCR: ' + (data.product || 'Product') + ' - ' + (data.defectType || 'Defect');
    body = 'New Quality NCR created.\n\nNCR No: ' + data.ncrNo + '\nProduct: ' + (data.product||'') +
           '\nDefect: ' + (data.defectType||'') + '\nQty Affected: ' + (data.qtyAffected||'') +
           '\nLocation: ' + loc + '\nEntered by: ' + (data.submittedBy||'User') + '\nTime: ' + ts;
    whatsappMsg = '⚠ SSPI Quality NCR [' + loc + ']: ' + (data.product||'Product') + ' | ' + (data.defectType||'Defect') + ' | Qty: ' + (data.qtyAffected||'?') + ' | NCR: ' + data.ncrNo;
  }
  else if (type === 'debtorsFU') {
    subject = '📞 SSPI Debtors Followup: ' + (data.customer || 'Customer');
    body = 'Debtors followup entry.\n\nCustomer: ' + data.customer +
           '\nOutcome: ' + (data.outcome||'') + '\nNext Action: ' + (data.next||'') +
           '\nEntered by: ' + (data.submittedBy||'User') + '\nTime: ' + ts;
    whatsappMsg = ''; // Not critical enough for WhatsApp
  }
  else if (type === 'accounts' && parseFloat(data.amount||0) >= 100000) {
    subject = '💰 SSPI High-Value Transaction: ₹' + Number(data.amount).toLocaleString('en-IN');
    body = 'High-value accounting entry (≥₹1 Lakh).\n\nParty: ' + (data.party||'') +
           '\nAmount: ₹' + Number(data.amount||0).toLocaleString('en-IN') +
           '\nType: ' + (data.voucherType||'') + '\nReference: ' + (data.reference||'') +
           '\nEntered by: ' + (data.submittedBy||'User') + '\nTime: ' + ts;
    whatsappMsg = '💰 SSPI High-Value Entry [' + loc + ']: ₹' + Number(data.amount||0).toLocaleString('en-IN') + ' | ' + (data.party||'Party') + ' | ' + (data.voucherType||'');
  }
  else if (type === 'advance' && data.entryType !== 'Recovery' && parseFloat(data.amount||0) > 0) {
    subject = '💸 SSPI Advance Disbursed: ' + (data.employee || data.empCode || 'Employee');
    body = 'Advance against salary disbursed.\n\nEmployee: ' + (data.employee||'') + ' (Code: ' + (data.empCode||'') + ')' +
           '\nAmount: ₹' + Number(data.amount||0).toLocaleString('en-IN') +
           '\nReason: ' + (data.reason||'') + '\nApproved By: ' + (data.approvedBy||'') +
           '\nLocation: ' + loc + '\nEntered by: ' + (data.submittedBy||'User') + '\nTime: ' + ts;
    whatsappMsg = '💸 SSPI Advance [' + loc + ']: ₹' + Number(data.amount||0).toLocaleString('en-IN') + ' to ' + (data.employee||'Employee') + ' | ' + (data.reason||'');
  }
  else {
    return; // Not a critical event — no alert needed
  }

  // Send email alert
  if (ALERT_EMAIL && subject) {
    try {
      MailApp.sendEmail({
        to: ALERT_EMAIL,
        subject: subject,
        body: body + '\n\n---\nSSPI Virtual Coordinator Auto-Alert\nDo not reply to this email.'
      });
    } catch(e) { Logger.log('Email alert failed: ' + e.message); }
  }

  // Send WhatsApp via CallMeBot (free — requires registration)
  if (WHATSAPP_NUMBER !== '919XXXXXXXXX' && CALLMEBOT_API_KEY !== 'XXXXXXXX' && whatsappMsg) {
    try {
      const encodedMsg = encodeURIComponent(whatsappMsg);
      const url = 'https://api.callmebot.com/whatsapp.php?phone=' + WHATSAPP_NUMBER +
                  '&text=' + encodedMsg + '&apikey=' + CALLMEBOT_API_KEY;
      UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    } catch(e) { Logger.log('WhatsApp alert failed: ' + e.message); }
  }
}

// ═══════════════════════════════════════════════════════════════
// KEEP-WARM — Prevents cold-start delay
// ═══════════════════════════════════════════════════════════════

/**
 * Lightweight function that keeps the GAS instance warm.
 * Set a 5-minute trigger on this function (see setupAllTriggers below).
 * Cost: 0. Effect: Eliminates the 8-15 second cold-start delay for users.
 */
function keepWarm() {
  Logger.log('Keep-warm ping: ' + new Date().toLocaleString('en-IN'));
  // Touch the spreadsheet to keep the connection active
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('Sheet active: ' + ss.getName());
}

// ═══════════════════════════════════════════════════════════════
// BACKUP SYSTEM — Weekly + Daily
// ═══════════════════════════════════════════════════════════════

/**
 * Full weekly backup — copies entire spreadsheet to "SSPI Backups" folder.
 * Sends email notification on completion.
 */
function createWeeklyBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');
  const backupName = 'SSPI_BACKEND_Weekly_' + today;
  let folder;
  const folders = DriveApp.getFoldersByName('SSPI Backups');
  folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('SSPI Backups');
  const copy = DriveApp.getFileById(ss.getId()).makeCopy(backupName, folder);
  Logger.log('Weekly backup: ' + backupName);
  try {
    MailApp.sendEmail({
      to: ALERT_EMAIL,
      subject: '✅ SSPI Weekly Backup Complete — ' + today,
      body: 'Weekly backup created: ' + backupName + '\nFolder: SSPI Backups (Google Drive)\nTime: ' + new Date().toLocaleString('en-IN') + '\n\nBackup file ID: ' + copy.getId()
    });
  } catch(e) {}
  return 'Weekly backup done: ' + backupName;
}

/**
 * Daily lightweight backup — exports each sheet as CSV to "SSPI Daily Backups" folder.
 * Much smaller than full copy; retains last 30 days automatically.
 */
function createDailyBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');
  let folder;
  const folders = DriveApp.getFoldersByName('SSPI Daily Backups');
  folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('SSPI Daily Backups');

  // Export each sheet as CSV blob
  const sheets = ss.getSheets();
  let savedCount = 0;
  sheets.forEach(function(sheet) {
    if (sheet.getLastRow() <= 1) return; // skip empty sheets
    const data = sheet.getDataRange().getValues();
    const csvRows = data.map(row => row.map(cell => '"' + String(cell).replace(/"/g, '""') + '"').join(','));
    const csvContent = csvRows.join('\n');
    const blob = Utilities.newBlob(csvContent, 'text/csv', sheet.getName() + '_' + today + '.csv');
    folder.createFile(blob);
    savedCount++;
  });

  // Delete files older than 30 days
  const thirtyDaysAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);
  const files = folder.getFiles();
  let deletedCount = 0;
  while (files.hasNext()) {
    const file = files.next();
    if (file.getDateCreated() < thirtyDaysAgo) {
      file.setTrashed(true);
      deletedCount++;
    }
  }

  Logger.log('Daily backup: ' + savedCount + ' sheets saved, ' + deletedCount + ' old files deleted');
  return 'Daily backup done: ' + savedCount + ' sheets';
}

// ═══════════════════════════════════════════════════════════════
// TRIGGER SETUP — Run once from GAS editor
// ═══════════════════════════════════════════════════════════════

/**
 * Sets up ALL required triggers in one call.
 * Run this ONCE from Apps Script editor → Run → setupAllTriggers
 * Then verify with listTriggers().
 *
 * Creates:
 * - keepWarm: every 5 minutes (eliminates cold start)
 * - createDailyBacku