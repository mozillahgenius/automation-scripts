// License Management System for Google Sheets - Dynamic Person Management Version

// Initialize the spreadsheet
function initializeSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('License Management');
  
  // Create sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet('License Management');
  } else {
    sheet.clear();
  }
  
  // Set up basic headers (without people initially)
  const headers = [
    'License Name',
    'Provider',
    'Category',
    'Base Cost (USD)',
    'Billing Period',
    'Per User',
    'Total Users',
    'Total Cost',
    'Notes'
  ];
  
  // Apply headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format header row (シンプルなスタイル)
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // Set column widths
  sheet.setColumnWidth(1, 150); // License Name
  sheet.setColumnWidth(2, 120); // Provider
  sheet.setColumnWidth(3, 100); // Category
  sheet.setColumnWidth(4, 100); // Base Cost
  sheet.setColumnWidth(5, 100); // Billing Period
  sheet.setColumnWidth(6, 80);  // Per User
  sheet.setColumnWidth(7, 80);  // Total Users
  sheet.setColumnWidth(8, 100); // Total Cost
  sheet.setColumnWidth(9, 200); // Notes
  
  // Add sample data
  addSampleLicenses(sheet);
  
  // Add formulas and data validation
  setupBasicFormulasAndValidation(sheet);
  
  // Create summary section
  createSummarySection(sheet);
  
  // Show completion message
  SpreadsheetApp.getActiveSpreadsheet().toast('License Management System initialized! Use "Add Person" to add team members.', 'Success', 5);
}

// Add sample license data
function addSampleLicenses(sheet) {
  const sampleData = [
    ['Salesforce', 'Salesforce, Inc.', 'Productivity', '', 'Yearly', 'TRUE'],
    ['QuotaPath', 'QuotaPath, Inc.', 'Productivity', '', '', 'TRUE'],
    ['Google Workspace', 'Google LLC', 'Productivity', 16.80, 'Monthly', 'TRUE'],
    ['Microsoft365', 'Microsoft Corporation', 'Productivity', '', 'Monthly', 'TRUE'],
    ['Bakuraku', 'LayerX Inc.', 'Security', '', 'Yearly', 'FALSE'],
    ['Zendesk', 'Zendesk, Inc.', 'Productivity', '', '', 'TRUE'],
    ['Grafana', 'Grafana Labs', 'Productivity', '', '', 'FALSE'],
    ['Notion', 'Notion Labs', 'Productivity', '', 'Monthly', 'TRUE'],
    ['Adobe', 'Adobe Inc.', 'Design', '', 'Yearly', 'TRUE'],
    ['Hibob', 'Hi Bob Ltd.', 'Productivity', '', 'Yearly', 'FALSE'],
    ['Redash', 'Databricks, Inc.', 'Productivity', '', '', ''],
    ['Asana', 'Asana, Inc.', 'Productivity', '', 'Yearly', 'FALSE'],
    ['Slack', 'Slack Technologies, LLC(Salesforce)', 'Communication', 8.75, 'Monthly', 'TRUE'],
    ['', '', '', '', '', ''], // Empty rows for new entries
    ['', '', '', '', '', '']
  ];
  
  // Insert sample data starting from row 2
  if (sampleData.length > 0) {
    sheet.getRange(2, 1, sampleData.length, 6).setValues(sampleData);
  }
}

// Set up basic formulas and data validation
function setupBasicFormulasAndValidation(sheet) {
  const lastRow = 20; // Fixed number of rows for licenses
  
  // Set up data validation for "Per User" column (TRUE/FALSE)
  const perUserRange = sheet.getRange(2, 6, lastRow - 1, 1);
  const perUserRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE', ''], true)
    .setAllowInvalid(false)
    .build();
  perUserRange.setDataValidation(perUserRule);
  
  // Set up data validation for "Billing Period" column
  const billingRange = sheet.getRange(2, 5, lastRow - 1, 1);
  const billingRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Monthly', 'Yearly', ''], true)
    .setAllowInvalid(false)
    .build();
  billingRange.setDataValidation(billingRule);
  
  // Set up data validation for "Category" column
  const categoryRange = sheet.getRange(2, 3, lastRow - 1, 1);
  const categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Productivity', 'Design', 'Development', 'Communication', 
                        'Project Management', 'Security', 'Infrastructure', 'Storage', 'Other', ''], true)
    .setAllowInvalid(false)
    .build();
  categoryRange.setDataValidation(categoryRule);
  
  // Initially set Total Users and Total Cost columns
  updateTotalFormulas(sheet);
  
  // Format cost columns as currency
  sheet.getRange(2, 4, lastRow - 1, 1).setNumberFormat('$#,##0.00');
  
  // Find and format Total Cost column
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const totalCostCol = headers.indexOf('Total Cost') + 1;
  if (totalCostCol > 0) {
    sheet.getRange(2, totalCostCol, lastRow - 1, 1).setNumberFormat('$#,##0.00');
  }
  
  // Add alternating row colors for better readability
  const lastCol = sheet.getLastColumn();
  const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
  // Remove any existing banding first
  const bandings = sheet.getBandings();
  bandings.forEach(banding => banding.remove());
  // Apply new banding
  dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
}

// Add a new person to the spreadsheet
function addPerson() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Add New Person',
    'Enter the name of the person:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const personName = response.getResponseText().trim();
    
    if (personName === '') {
      ui.alert('Error', 'Please enter a valid name.', ui.ButtonSet.OK);
      return;
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('License Management');
    
    // Find where to insert new columns (before Total Users)
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let totalUsersCol = headers.indexOf('Total Users') + 1;
    
    if (totalUsersCol === 0) {
      ui.alert('Error', 'Could not find Total Users column.', ui.ButtonSet.OK);
      return;
    }
    
    // Insert two new columns for the person
    sheet.insertColumns(totalUsersCol, 2);
    
    // Set headers for new columns
    sheet.getRange(1, totalUsersCol).setValue(`${personName} - Need`);
    sheet.getRange(1, totalUsersCol + 1).setValue(`${personName} - Cost`);
    
    // Format new headers
    const newHeaderRange = sheet.getRange(1, totalUsersCol, 1, 2);
    newHeaderRange.setBackground('#4285F4');
    newHeaderRange.setFontColor('#FFFFFF');
    newHeaderRange.setFontWeight('bold');
    newHeaderRange.setHorizontalAlignment('center');
    
    // Set column widths
    sheet.setColumnWidth(totalUsersCol, 100);
    sheet.setColumnWidth(totalUsersCol + 1, 100);
    
    // Add data validation for Need column
    const needRange = sheet.getRange(2, totalUsersCol, 19, 1);
    const needRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Yes', 'No', ''], true)
      .setAllowInvalid(false)
      .build();
    needRange.setDataValidation(needRule);
    
    // Add formulas for Cost column
    for (let row = 2; row <= 20; row++) {
      const formula = generatePersonCostFormula(sheet, row, totalUsersCol);
      sheet.getRange(row, totalUsersCol + 1).setFormula(formula);
    }
    
    // Format cost column as currency
    sheet.getRange(2, totalUsersCol + 1, 19, 1).setNumberFormat('$#,##0.00');
    
    // Update Total Users and Total Cost formulas
    updateTotalFormulas(sheet);
    
    // Update summary section
    updateSummarySection(sheet);
    
    // Re-apply banding to include new columns
    const lastRow = 20;
    const lastCol = sheet.getLastColumn();
    const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    // Remove any existing banding first
    const bandings = sheet.getBandings();
    bandings.forEach(banding => banding.remove());
    // Apply new banding
    dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(`Added ${personName} to the system!`, 'Success', 3);
  }
}

// Generate cost formula for a person
function generatePersonCostFormula(sheet, row, needCol) {
  // Simple formula for per-user licenses and shared fixed costs
  return `=IF(AND(A${row}<>"", ${String.fromCharCode(64 + needCol)}${row}="Yes"), D${row}, 0)`;
}

// Update Total Users and Total Cost formulas
function updateTotalFormulas(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const totalUsersCol = headers.indexOf('Total Users') + 1;
  const totalCostCol = headers.indexOf('Total Cost') + 1;
  
  // Find all Need columns
  const needColumns = [];
  for (let i = 0; i < headers.length; i++) {
    if (headers[i].includes('- Need')) {
      needColumns.push(i + 1);
    }
  }
  
  if (needColumns.length === 0) {
    // No people added yet
    for (let row = 2; row <= 20; row++) {
      sheet.getRange(row, totalUsersCol).setFormula(`=IF(A${row}<>"", 0, "")`);
      sheet.getRange(row, totalCostCol).setFormula(`=IF(A${row}<>"", 0, "")`);
    }
  } else {
    // Update formulas with all Need columns
    for (let row = 2; row <= 20; row++) {
      // Total Users formula - count how many people need this license
      const needCells = needColumns.map(col => `${String.fromCharCode(64 + col)}${row}`).join(',');
      sheet.getRange(row, totalUsersCol).setFormula(
        `=IF(A${row}<>"", COUNTIF({${needCells}},"Yes"), "")`
      );
      
      // Total Cost formula
      // For Per User licenses: Base Cost × Number of Users
      // For Fixed licenses: Base Cost (regardless of number of users)
      // 列番号を直接確認してG列（Total Users）を参照
      sheet.getRange(row, totalCostCol).setFormula(
        `=IF(A${row}<>"", IF(AND(F${row}="TRUE", D${row}<>""), D${row}*G${row}, ` +
        `IF(AND(F${row}="FALSE", G${row}>0, D${row}<>""), D${row}, 0)), "")`
      );
    }
  }
  
  // Format cost columns
  sheet.getRange(2, totalCostCol, 19, 1).setNumberFormat('$#,##0.00');
}

// Create summary section
function createSummarySection(sheet) {
  const summaryStartRow = 23;
  
  // Clear any existing content in summary area
  sheet.getRange(summaryStartRow, 1, 10, 10).clear();
  
  // Summary header
  sheet.getRange(summaryStartRow, 1).setValue('SUMMARY');
  sheet.getRange(summaryStartRow, 1, 1, 5).merge();
  sheet.getRange(summaryStartRow, 1).setBackground('#34A853');
  sheet.getRange(summaryStartRow, 1).setFontColor('#FFFFFF');
  sheet.getRange(summaryStartRow, 1).setFontWeight('bold');
  sheet.getRange(summaryStartRow, 1).setHorizontalAlignment('center');
  sheet.getRange(summaryStartRow, 1).setFontSize(12);
  
  // Add note if no people added
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const peopleCount = headers.filter(h => h.includes('- Need')).length;
  
  if (peopleCount === 0) {
    sheet.getRange(summaryStartRow + 2, 1).setValue('No people added yet. Use "Add Person" from the menu to add team members.');
    sheet.getRange(summaryStartRow + 2, 1, 1, 5).merge();
    sheet.getRange(summaryStartRow + 2, 1).setFontStyle('italic');
  }
}

// Update summary section when people are added
function updateSummarySection(sheet) {
  const summaryStartRow = 23;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Get all people names and their cost columns
  const people = [];
  for (let i = 0; i < headers.length; i++) {
    if (headers[i].includes('- Cost')) {
      const personName = headers[i].replace(' - Cost', '');
      people.push({
        name: personName,
        costCol: i + 1
      });
    }
  }
  
  if (people.length === 0) return;
  
  // Clear and rebuild summary
  sheet.getRange(summaryStartRow, 1, 10, people.length + 2).clear();
  
  // Summary headers
  const summaryHeaders = [
    ['SUMMARY'],
    [''].concat(people.map(p => p.name)).concat(['Company Total']),
    ['Monthly Licenses'],
    ['Yearly Licenses'],
    ['Total Monthly Cost'],
    ['Total Yearly Cost']
  ];
  
  // Apply headers (adjust columns based on number of people)
  sheet.getRange(summaryStartRow, 1, 1, people.length + 2).merge();
  sheet.getRange(summaryStartRow, 1).setValue('SUMMARY');
  sheet.getRange(summaryStartRow, 1).setBackground('#34A853');
  sheet.getRange(summaryStartRow, 1).setFontColor('#FFFFFF');
  sheet.getRange(summaryStartRow, 1).setFontWeight('bold');
  sheet.getRange(summaryStartRow, 1).setHorizontalAlignment('center');
  sheet.getRange(summaryStartRow, 1).setFontSize(12);
  
  // Person names header row
  sheet.getRange(summaryStartRow + 1, 1).setValue('');
  for (let i = 0; i < people.length; i++) {
    sheet.getRange(summaryStartRow + 1, i + 2).setValue(people[i].name);
  }
  sheet.getRange(summaryStartRow + 1, people.length + 2).setValue('Company Total');
  
  const headerRow = sheet.getRange(summaryStartRow + 1, 1, 1, people.length + 2);
  headerRow.setBackground('#E8F0FE');
  headerRow.setFontWeight('bold');
  
  // Row labels
  const rowLabels = ['Monthly Licenses', 'Yearly Licenses', 'Total Monthly Cost', 'Total Yearly Cost'];
  for (let i = 0; i < rowLabels.length; i++) {
    sheet.getRange(summaryStartRow + 2 + i, 1).setValue(rowLabels[i]);
  }
  
  // Add formulas for each person
  const dataLastRow = 20;
  
  for (let i = 0; i < people.length; i++) {
    const col = i + 2;
    const costCol = people[i].costCol;
    const costColLetter = String.fromCharCode(64 + costCol);
    
    // Monthly licenses cost
    sheet.getRange(summaryStartRow + 2, col).setFormula(
      `=SUMIFS($${costColLetter}$2:$${costColLetter}$${dataLastRow}, $E$2:$E$${dataLastRow}, "Monthly", $A$2:$A$${dataLastRow}, "<>")`
    );
    
    // Yearly licenses cost (divided by 12 for monthly view)
    sheet.getRange(summaryStartRow + 3, col).setFormula(
      `=SUMIFS($${costColLetter}$2:$${costColLetter}$${dataLastRow}, $E$2:$E$${dataLastRow}, "Yearly", $A$2:$A$${dataLastRow}, "<>")/12`
    );
    
    // Total monthly cost
    sheet.getRange(summaryStartRow + 4, col).setFormula(
      `=${String.fromCharCode(64 + col)}${summaryStartRow + 2}+${String.fromCharCode(64 + col)}${summaryStartRow + 3}`
    );
    
    // Total yearly cost
    sheet.getRange(summaryStartRow + 5, col).setFormula(
      `=(${String.fromCharCode(64 + col)}${summaryStartRow + 2}*12)+(${String.fromCharCode(64 + col)}${summaryStartRow + 3}*12)`
    );
  }
  
  // Company totals
  const totalCostCol = headers.indexOf('Total Cost') + 1;
  const totalCostColLetter = String.fromCharCode(64 + totalCostCol);
  const companyCol = people.length + 2;
  
  sheet.getRange(summaryStartRow + 2, companyCol).setFormula(
    `=SUMIFS($${totalCostColLetter}$2:$${totalCostColLetter}$${dataLastRow}, $E$2:$E$${dataLastRow}, "Monthly", $A$2:$A$${dataLastRow}, "<>")`
  );
  
  sheet.getRange(summaryStartRow + 3, companyCol).setFormula(
    `=SUMIFS($${totalCostColLetter}$2:$${totalCostColLetter}$${dataLastRow}, $E$2:$E$${dataLastRow}, "Yearly", $A$2:$A$${dataLastRow}, "<>")/12`
  );
  
  sheet.getRange(summaryStartRow + 4, companyCol).setFormula(
    `=${String.fromCharCode(64 + companyCol)}${summaryStartRow + 2}+${String.fromCharCode(64 + companyCol)}${summaryStartRow + 3}`
  );
  
  sheet.getRange(summaryStartRow + 5, companyCol).setFormula(
    `=(${String.fromCharCode(64 + companyCol)}${summaryStartRow + 2}*12)+(${String.fromCharCode(64 + companyCol)}${summaryStartRow + 3}*12)`
  );
  
  // Format summary cost cells
  sheet.getRange(summaryStartRow + 2, 2, 4, people.length + 1).setNumberFormat('$#,##0.00');
  
  // Add borders to summary section
  sheet.getRange(summaryStartRow, 1, 6, people.length + 2).setBorder(true, true, true, true, true, true);
}

// Remove a person from the spreadsheet
function removePerson() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('License Management');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Get list of people
  const people = [];
  for (let i = 0; i < headers.length; i++) {
    if (headers[i].includes('- Need')) {
      people.push(headers[i].replace(' - Need', ''));
    }
  }
  
  if (people.length === 0) {
    ui.alert('No People', 'No people have been added yet.', ui.ButtonSet.OK);
    return;
  }
  
  // Create HTML for selection dialog
  const htmlContent = '<p>Select a person to remove:</p><select id="person" size="5" style="width: 200px;">' +
    people.map(p => `<option value="${p}">${p}</option>`).join('') +
    '</select><br><br>' +
    '<button onclick="google.script.run.withSuccessHandler(() => google.script.host.close()).removePersonConfirmed(document.getElementById(\'person\').value)">Remove</button> ' +
    '<button onclick="google.script.host.close()">Cancel</button>';
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(250)
    .setHeight(200);
  
  ui.showModalDialog(htmlOutput, 'Remove Person');
}

// Confirm and remove person
function removePersonConfirmed(personName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('License Management');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Find the person's columns
  let needCol = -1;
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] === `${personName} - Need`) {
      needCol = i + 1;
      break;
    }
  }
  
  if (needCol > 0) {
    // Delete both Need and Cost columns
    sheet.deleteColumns(needCol, 2);
    
    // Update formulas
    updateTotalFormulas(sheet);
    updateSummarySection(sheet);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(`Removed ${personName} from the system.`, 'Success', 3);
  }
}

// Clear all license data (keep structure)
function clearAllData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Clear All Data',
    'This will clear all license data but keep the structure and people. Are you sure?',
    ui.ButtonSet.YES_NO
  );
  
  if (response == ui.Button.YES) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('License Management');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Clear license data
    sheet.getRange(2, 1, 19, 6).clearContent();
    
    // Clear all Need columns
    for (let i = 0; i < headers.length; i++) {
      if (headers[i].includes('- Need')) {
        sheet.getRange(2, i + 1, 19, 1).setValue('');
      }
    }
    
    SpreadsheetApp.getActiveSpreadsheet().toast('All data cleared!', 'Success', 3);
  }
}

// Export summary to new sheet
function exportSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('License Management');
  const headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  
  // Check if there are people to export
  const peopleCount = headers.filter(h => h.includes('- Need')).length;
  if (peopleCount === 0) {
    SpreadsheetApp.getUi().alert('No Data', 'Please add people before exporting summary.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Create new sheet with timestamp
  const timestamp = new Date().toISOString().slice(0, 10);
  const newSheetName = `Summary_${timestamp}`;
  const newSheet = ss.insertSheet(newSheetName);
  
  // Copy summary section
  const summaryWidth = peopleCount + 2;
  const summaryRange = sourceSheet.getRange(23, 1, 6, summaryWidth);
  summaryRange.copyTo(newSheet.getRange(1, 1), {contentsOnly: false});
  
  // Add license details
  newSheet.getRange(8, 1).setValue('LICENSE DETAILS');
  newSheet.getRange(8, 1, 1, summaryWidth).merge();
  newSheet.getRange(8, 1).setBackground('#4285F4');
  newSheet.getRange(8, 1).setFontColor('#FFFFFF');
  newSheet.getRange(8, 1).setFontWeight('bold');
  newSheet.getRange(8, 1).setHorizontalAlignment('center');
  
  // Copy license data with people selections
  const dataRange = sourceSheet.getRange(1, 1, 20, sourceSheet.getLastColumn());
  dataRange.copyTo(newSheet.getRange(10, 1), {contentsOnly: true});
  
  // Add timestamp
  newSheet.getRange(31, 1).setValue('Generated on: ' + new Date().toLocaleString());
  newSheet.getRange(31, 1).setFontStyle('italic');
  
  SpreadsheetApp.getActiveSpreadsheet().toast(`Summary exported to sheet: ${newSheetName}`, 'Success', 5);
}

// Menu creation
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('License Management')
    .addItem('Initialize Spreadsheet', 'initializeSpreadsheet')
    .addSeparator()
    .addItem('Add Person', 'addPerson')
    .addItem('Remove Person', 'removePerson')
    .addSeparator()
    .addItem('Clear All Data', 'clearAllData')
    .addItem('Export Summary', 'exportSummary')
    .addSeparator()
    .addItem('Refresh Calculations', 'refreshCalculations')
    .addToUi();
}

// Refresh all calculations
function refreshCalculations() {
  SpreadsheetApp.flush();
  SpreadsheetApp.getActiveSpreadsheet().toast('Calculations refreshed!', 'Success', 3);
}