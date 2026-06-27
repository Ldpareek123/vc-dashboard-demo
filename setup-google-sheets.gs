// ============================================================
// SSPI Dashboard V4 - Google Sheets Setup
// Paste into Google Apps Script (script.google.com)
// ============================================================

function setupSSPIBackend() {
  Logger.log("🚀 Starting SSPI Dashboard Setup...");

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const sheetConfig = {
      "HR & Payroll": [
        "Employee Master",
        "Attendance",
        "Salary Monthly",
        "Salary Revision",
        "Advance Ledger",
        "Leave Register",
        "ESI EPF Compliance"
      ],
      "Finance & Accounting": [
        "Accounts",
        "Bills Purchase",
        "Bills Sales",
        "Bills Expense",
        "Debtors Invoices",
        "Debtors Receipts",
        "Debtors Followup",
        "GST Purchase",
        "GST Sales"
      ],
      "Manufacturing": [
        "Job Cards",
        "Production Plans",
        "Production Log",
        "Quality NCR",
        "Maintenance",
        "Route Cards"
      ],
      "Inventory & Materials": [
        "Stores",
        "Purchase"
      ],
      "Sales & Orders": [
        "Sales Orders",
        "Dispatch Register"
      ],
      "Master & Compliance": [
        "Approvals",
        "Payroll Register"
      ]
    };

    let totalSheets = 0;
    for (const category in sheetConfig) {
      Logger.log(`📁 Creating ${category} sheets...`);

      for (const sheetName of sheetConfig[category]) {
        try {
          const sheet = ss.insertSheet(sheetName);

          const headers = getHeadersForSheet(sheetName);
          if (headers.length > 0) {
            sheet.appendRow(headers);
          }

          const headerRange = sheet.getRange(1, 1, 1, headers.length);
          headerRange.setBackground("#1F3864")
                    .setFontColor("#FFFFFF")
                    .setFontWeight("bold");

          sheet.setFrozenRows(1);

          Logger.log(`  ✓ Created: ${sheetName}`);
          totalSheets++;
        } catch (e) {
          Logger.log(`  ✗ Error: ${sheetName}: ${e.message}`);
        }
      }
    }

    Logger.log(`\n✅ Setup complete! Created ${totalSheets} sheets`);

  } catch (error) {
    Logger.log(`❌ Error: ${error.message}`);
  }
}

function getHeadersForSheet(sheetName) {
  const headers = {
    "Employee Master": ["emp_id", "name", "dept", "designation", "doj", "salary_base", "hra", "da", "status"],
    "Attendance": ["att_id", "emp_id", "date", "punch_in", "punch_out", "status", "hours"],
    "Salary Monthly": ["sal_id", "month", "emp_id", "basic", "hra", "da", "allowances", "gross", "net_salary", "status"],
    "Advance Ledger": ["adv_id", "emp_id", "amount", "date_applied", "date_approved", "reason", "status"],
    "Leave Register": ["leave_id", "emp_id", "type", "from_date", "to_date", "days", "reason", "status"],
    "Accounts": ["vouch_id", "date", "type", "payee", "amount", "description", "status"],
    "Bills Purchase": ["bill_id", "vendor", "bill_date", "amount", "gst", "total", "status"],
    "Debtors Invoices": ["inv_id", "customer", "invoice_date", "amount", "gst", "total", "status"],
    "Job Cards": ["jc_id", "po_id", "mach_id", "operator", "start_time", "end_time", "status"],
    "Production Log": ["po_id", "customer", "item", "quantity", "due_date", "status"],
    "Quality NCR": ["qc_id", "po_id", "date", "defects", "status"],
    "Stores": ["mat_id", "name", "quantity", "reorder_level", "status"],
    "Sales Orders": ["so_id", "customer", "amount", "gst", "total", "status"],
    "Dispatch Register": ["disp_id", "so_id", "date", "courier", "status"]
  };

  return headers[sheetName] || [];
}

// INSTRUCTIONS:
// 1. Go to https://sheets.google.com
// 2. Create new workbook → name it: SSPI_BACKEND
// 3. Go to https://script.google.com
// 4. Create new project
// 5. Paste this code
// 6. Save
// 7. Click Run
// 8. Authorize access
// 9. Check Logger for completion
// 10. Go back to sheets - 20+ sheets created!
