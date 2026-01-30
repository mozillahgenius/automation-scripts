# license-management-en
> Dynamic column-based license cost management system with per-person tracking in Google Sheets

## Overview
A Google Sheets-based license management system that dynamically adds person-specific columns to track which team members need which software licenses and the associated costs. The system supports per-user and fixed-cost license types, calculates per-person and company-wide totals across monthly and yearly billing periods, and provides a summary dashboard with export functionality. All operations are accessible through a custom Google Sheets menu.

## Key Features
- **Dynamic person management**: Add/remove team members via custom menu, each getting a "Need" and "Cost" column pair
- **Per-user vs. fixed cost tracking**: Distinguishes between per-user licenses (cost x number of users) and fixed licenses (flat cost regardless of users)
- **Automatic formula generation**: Total Users (COUNTIF) and Total Cost formulas are auto-generated and updated when people are added/removed
- **Data validation**: Dropdown lists for Category (9 options), Billing Period (Monthly/Yearly), and Per User (TRUE/FALSE)
- **Summary section**: Auto-generated summary with Monthly Licenses, Yearly Licenses, Total Monthly Cost, Total Yearly Cost per person and company total
- **Export functionality**: Export summary + full license data to a timestamped snapshot sheet
- **Sample data**: Pre-loaded with 13 common SaaS licenses (Salesforce, Google Workspace, Slack, Notion, etc.)
- **Custom menu**: "License Management" menu with Initialize, Add/Remove Person, Clear Data, Export, Refresh

## Architecture
```
[Google Sheets: "License Management" sheet]
  |
  +-- Static Columns (A-F): License Name, Provider, Category, Base Cost, Billing Period, Per User
  |
  +-- Dynamic Person Columns (inserted before Total Users):
  |     [Person Name - Need] [Person Name - Cost]  (per person, added dynamically)
  |
  +-- Calculated Columns: Total Users (G), Total Cost (H), Notes (I)
  |
  +-- Summary Section (Row 23+):
        Per-person and company-wide cost breakdown by billing period

[Custom Menu: "License Management"]
  +-- Initialize Spreadsheet
  +-- Add Person      -> Inserts 2 columns (Need + Cost) before Total Users
  +-- Remove Person   -> Shows HTML dialog to select and remove a person
  +-- Clear All Data  -> Clears license and assignment data, keeps structure
  +-- Export Summary  -> Creates timestamped snapshot sheet
  +-- Refresh Calculations
```

### Column Insertion Flow
When "Add Person" is executed:
1. User is prompted for a name
2. Two columns are inserted before "Total Users": `[Name - Need]` and `[Name - Cost]`
3. "Need" column gets Yes/No data validation
4. "Cost" column gets formula: `=IF(AND(A{row}<>"", {NeedCol}{row}="Yes"), D{row}, 0)`
5. Total Users formula is rebuilt: `=COUNTIF({all Need columns}, "Yes")`
6. Total Cost formula is rebuilt: Per-user = Base Cost x Total Users; Fixed = Base Cost (if users > 0)
7. Summary section is rebuilt with the new person included

## File Structure
| File | Description |
|---|---|
| `Code.js` | All logic (561 lines): initialization, dynamic column management, formulas, summary, export, custom menu |
| `appsscript.json` | GAS project settings (timezone: Asia/Singapore, V8 runtime) |

## Key Functions

### Initialization
| Function | Description |
|---|---|
| `initializeSpreadsheet()` | Creates "License Management" sheet with headers, sample data, validation, formulas, summary, and banding |
| `addSampleLicenses(sheet)` | Inserts 13 sample SaaS licenses (Salesforce, Google Workspace, Slack, Notion, Adobe, etc.) |
| `setupBasicFormulasAndValidation(sheet)` | Sets data validation for Category, Billing Period, Per User columns; applies currency format and banding |

### Person Management
| Function | Description |
|---|---|
| `addPerson()` | Prompts for name, inserts Need + Cost columns, sets validation, generates formulas, updates totals and summary |
| `removePerson()` | Shows HTML selection dialog listing all people, calls `removePersonConfirmed()` on selection |
| `removePersonConfirmed(personName)` | Deletes the person's Need + Cost columns, rebuilds total formulas and summary |

### Formula Management
| Function | Description |
|---|---|
| `generatePersonCostFormula(sheet, row, needCol)` | Generates `=IF(AND(A{row}<>"", {NeedCol}{row}="Yes"), D{row}, 0)` |
| `updateTotalFormulas(sheet)` | Rebuilds Total Users (COUNTIF across all Need columns) and Total Cost (per-user vs. fixed logic) formulas |

### Summary & Export
| Function | Description |
|---|---|
| `createSummarySection(sheet)` | Creates summary header at row 23; shows placeholder if no people added |
| `updateSummarySection(sheet)` | Rebuilds full summary: per-person Monthly/Yearly costs using SUMIFS, Company Total column, formatting |
| `exportSummary()` | Creates a new sheet named `Summary_YYYY-MM-DD` with summary + full license data snapshot |

### Utilities
| Function | Description |
|---|---|
| `onOpen()` | Creates custom "License Management" menu with all operations |
| `clearAllData()` | Clears license data (rows 2-20, columns A-F) and all Need column values, keeps structure |
| `refreshCalculations()` | Forces SpreadsheetApp.flush() to recalculate all formulas |

## Services & APIs Used
- **Google Sheets API** (SpreadsheetApp) - Sheet creation, data management, formatting, formulas, data validation
- **HTML Service** (HtmlService) - Modal dialog for person removal selection
- **UI Service** (SpreadsheetApp.getUi()) - Custom menu, prompts, alerts, toast notifications

## Setup Instructions
1. Create a new Google Spreadsheet
2. Open Apps Script editor (Extensions > Apps Script)
3. Copy the contents of `Code.js` into the script editor
4. Save the script
5. Reload the spreadsheet (the custom menu "License Management" will appear)
6. Click "License Management" > "Initialize Spreadsheet" to set up the sheet with sample data
7. Use "Add Person" to add team members
8. Fill in Base Cost values and set each person's "Need" column to "Yes" or "No"

## Script Properties
This project does not use script properties. All configuration is managed through the spreadsheet UI and custom menu.

## Sheet Structure

### License Management Sheet (Rows 1-20)
| Column | Header | Description |
|---|---|---|
| A | License Name | Name of the software/service |
| B | Provider | Vendor/company name |
| C | Category | Dropdown: Productivity, Design, Development, Communication, Project Management, Security, Infrastructure, Storage, Other |
| D | Base Cost (USD) | Cost per unit (per user for per-user licenses, flat for fixed licenses) |
| E | Billing Period | Dropdown: Monthly, Yearly |
| F | Per User | Dropdown: TRUE (per-user cost), FALSE (fixed cost) |
| ... | {Person} - Need | Dropdown: Yes / No (dynamically inserted per person) |
| ... | {Person} - Cost | Formula: Base Cost if Need=Yes, else 0 (dynamically inserted per person) |
| G* | Total Users | Formula: COUNTIF across all Need columns for "Yes" |
| H* | Total Cost | Formula: Per-user = Base Cost x Total Users; Fixed = Base Cost if users > 0 |
| I* | Notes | Free text |

*Column letters shift right as person columns are added.

### Summary Section (Row 23+)
| Row | Content | Description |
|---|---|---|
| 23 | SUMMARY | Section header (green background) |
| 24 | Person names + Company Total | Column headers |
| 25 | Monthly Licenses | SUMIFS of Cost column where Billing Period = "Monthly" |
| 26 | Yearly Licenses | SUMIFS of Cost column where Billing Period = "Yearly", divided by 12 |
| 27 | Total Monthly Cost | Monthly + Yearly (monthly equivalent) |
| 28 | Total Yearly Cost | (Monthly x 12) + (Yearly monthly equivalent x 12) |

## Sample Licenses (Pre-loaded)
| License | Provider | Category |
|---|---|---|
| Salesforce | Salesforce, Inc. | Productivity |
| QuotaPath | QuotaPath, Inc. | Productivity |
| Google Workspace | Google LLC | Productivity |
| Microsoft365 | Microsoft Corporation | Productivity |
| Bakuraku | LayerX Inc. | Security |
| Zendesk | Zendesk, Inc. | Productivity |
| Grafana | Grafana Labs | Productivity |
| Notion | Notion Labs | Productivity |
| Adobe | Adobe Inc. | Design |
| Hibob | Hi Bob Ltd. | Productivity |
| Redash | Databricks, Inc. | Productivity |
| Asana | Asana, Inc. | Productivity |
| Slack | Slack Technologies, LLC (Salesforce) | Communication |

## Custom Menu Items
| Menu Item | Function | Description |
|---|---|---|
| Initialize Spreadsheet | `initializeSpreadsheet()` | Set up fresh sheet with headers, sample data, and formulas |
| Add Person | `addPerson()` | Add a new team member with Need/Cost columns |
| Remove Person | `removePerson()` | Remove a team member and their columns |
| Clear All Data | `clearAllData()` | Clear all license/assignment data, keep structure |
| Export Summary | `exportSummary()` | Create timestamped snapshot sheet |
| Refresh Calculations | `refreshCalculations()` | Force recalculation of all formulas |
