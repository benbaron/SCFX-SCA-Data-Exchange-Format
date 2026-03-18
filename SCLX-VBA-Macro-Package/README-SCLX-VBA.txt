
SCLX VBA macro package for the SCA Exchequer Report workbook

Files
- SCLX_Ledger_IO.bas
- README-SCLX-VBA.txt

What this macro does
- Exports the current workbook to SCLX JSON
- Imports SCLX JSON back into the workbook
- Handles:
  - Ledger rows from the Ledger tab
  - Outstanding items
  - Assets & Inventory
  - Supplies
  - Lightweight organization/reportingPeriod/chartOfAccounts synthesis

Important installation step
This module uses JsonConverter from VBA-JSON.

Import these modules into the VBA project:
1. SCLX_Ledger_IO.bas
2. JsonConverter.bas (from VBA-JSON)

How to install
1. Open the workbook in Excel.
2. Press Alt+F11.
3. File > Import File...
4. Import SCLX_Ledger_IO.bas
5. Import JsonConverter.bas
6. Save the workbook as .xlsm

How to run
- Alt+F8
- Run ExportSCLX to write a .json file
- Run ImportSCLX to load a .json file

Notes
- The workbook's Budget tab is highly formula-driven. This module leaves the SCLX "budgets" collection empty by default.
- Workbook-specific fields are preserved in extensions.workbook blocks so round-tripping back into this form is practical.
- The chart of accounts is synthesized from encountered category/account names because this workbook does not expose a single explicit numbered COA table in a simple flat entry range.

Recommended next refinement
- Add a finalized budget region map if you want full budget round-trip support.
- Add bankStatementImports / bankingItems if you want OFX-preserving import/export from the workbook.
