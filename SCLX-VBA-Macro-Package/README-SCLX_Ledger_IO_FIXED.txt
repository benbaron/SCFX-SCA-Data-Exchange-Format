Corrected SCLX VBA module.

Changes:
- Preserved the exact working GetSaveAsFilename fragment supplied by the user.
- Fixed object assignment bugs that caused Run-time error 450 when assigning
  Dictionary/Collection objects into Scripting.Dictionary items.
- Corrected nested object insertion points such as:
  root -> organization/reportingPeriod/chartOfAccounts/funds/budgets/transactions...
  tx -> extensions/lines
  d -> extensions/currentGuardian/guardianshipDetails/removalDetails/guardian
  account -> SupplementalKinds

Install:
1. Open the VBA editor.
2. Remove the old SCLX_Ledger_IO module.
3. Import SCLX_Ledger_IO_FIXED.bas.
4. Ensure JsonConverter.bas is also imported.
