# SCLX 1.3 Integrator Manual (Ledger-Native Revision)

## 1. Purpose

SCLX is a JSON exchange format for ledger data plus selected nonprofit operational records. It is intended for ledger export/import, audit review, archival interchange, spreadsheet integration, and bank import preservation.

This 1.3 ledger-native revision explicitly supports two transaction styles:

- **canonical balanced accounting transactions**, and
- **worksheet-native ledger entries** exported directly from a source workbook even when the source workbook exposes only one posting side.

## 2. What belongs in SCLX

SCLX contains:
- core ledger master data
- ledger transactions with one or more posting lines
- optional bank reconciliation records
- optional outstanding-item and other-asset schedules
- optional supplemental schedule items
- optional asset and supplies registries
- optional budgets
- optional OFX statement import metadata

SCLX does not try to replace the full OFX transport protocol, spreadsheet formatting, or UI-specific workflow state.

## 3. Top-level collections

Required:
- format
- version
- exportedAt
- organization
- reportingPeriod
- chartOfAccounts
- funds
- budgets
- people
- events
- documents
- transactions

Optional:
- bankAccounts
- officeAssignments
- committeeMemberships
- bankingItems
- outstandingItems
- otherAssetItems
- supplementalItems
- assets
- supplies
- bankStatementImports
- features
- compatibility
- extensions

## 4. Chart of Accounts

Each account follows the de facto row model:
- Number
- Name
- Type
- Parent
- IncreaseSide
- OpeningBalance
- SupplementalKinds

Optional fields may also include:
- accountId
- code
- subtype
- active
- reportingTags

Number is the canonical chart identifier for standalone chart exchange. accountId is allowed for systems that prefer a separate immutable internal key.

## 5. Transactions

Transactions are the core ledger records. A transaction may represent either:
- a fully balanced accounting transaction, or
- a worksheet-native ledger entry preserved directly from a source ledger workbook.

Each transaction must:
- have at least one line
- contain only nonnegative debit and credit values on each line
- use exactly one nonzero side per line

When the source system exposes enough information to emit a balanced multi-line transaction, producers should do so. When the source system exposes only one posting side, producers may emit a single-sided ledger entry without inventing the balancing side.

Use transactionDate for the economic event date and postingDate for the posting date preserved by the exporting system.

## 6. Transaction lines

Each line references one account and may optionally reference:
- fundId
- budgetId
- eventId
- personId
- documentId

A line may also carry:
- tags
- restrictionTag
- reportSection
- supplementalRefs
- usedFor
- itemNumber
- quantity
- workbookLink
- extensions

## 7. Funds, budgets, people, governance records, events, documents

These are all first-class master data collections. Readers should load them before validating line references.

Budgets may be empty if the source system has no stable budget-entry region or does not track budgets structurally.



## 8. Bank accounts

bankAccounts represent bank-account master records, not statement imports and not bank-side settlement items. Typical uses include:
- institution name and contact details
- masked account number
- account type
- account holder name
- interest-bearing flag
- signature requirements
- authorized signers
- linkage to the chart-of-accounts bank/cash account

This object is intended to preserve workbook- or system-level banking administration data that would otherwise be lost if only bankingItems and bankStatementImports were exchanged.

## 9. Banking items

bankingItems represent bank-side settlement facts, not ledger truth. Typical uses:
- cleared checks
- deposits
- bank fees
- interest
- statement adjustments

For CHECK items, require:
- checkNumber
- payee

For DEPOSIT items, require:
- depositDate
- payer

Amounts are always positive strings in bankingItems.

## 10. OFX preservation

If a bank import originated from OFX, preserve the statement-level metadata in bankStatementImports and preserve transaction-level metadata in bankingItems[].ofx.

High-value OFX fields:
- fitId
- transactionType
- datePosted
- dateUser
- dateAvailable
- checkNumber
- referenceNumber
- name
- memo
- payeeId
- sic
- correctFitId
- correctAction

Recommended duplicate key:
(sourceFormat, bankId?, accountId, fitId)

## 11. Outstanding items

Outstanding items are supplemental schedule records that may optionally link back to a transaction line using ledgerLink. They may track:
- outstanding checks
- deposits in transit
- transfers
- incoming checks
- card items

A null ledgerLink is allowed when the source workbook contains a standalone register row without a resolvable canonical transaction-line link.

These records do not replace accounting entries.

## 12. Other asset items

otherAssetItems represent schedule-style items such as:
- cash advances
- site security deposits
- other recoverable amounts

Use typeCode values:
- C = CASH_ADVANCE
- S = SITE_SECURITY_DEPOSIT
- O = OTHER

## 13. Supplemental schedule items

supplementalItems preserve structured schedule rows that do not fit cleanly as ordinary transactions or registries. Typical uses include:
- receivables
- prepaid expenses
- deferred revenue
- payables
- other asset schedules
- other liability schedules

A supplemental item may include a ledger row reference, budget label, subtype code, remaining balance, workbook section name, and workbookLink.

## 14. Assets registry

assets records track durable property and custodianship. The `itemType` field is a workbook-native raw value. Producers may preserve source labels directly rather than mapping them to a controlled enum.

Assets may also carry appraisalDetails when the source workbook or source system captures:
- appraiser name
- appraisal date
- revised value

These are operational records, not accounting postings. Link them to relatedTransactionIds or relatedLineIds when the acquisition, sale, or write-off appears in the ledger.

## 15. Supplies registry

supplies records track lower-value or consumable items that may still have guardianship and removal history. Use removalDetails to track sold, lost, donated, destroyed, or returned items.

## 16. Extensions

Any object may include an extensions object. Extensions should be namespaced by producer, for example:
- scaledger
- workbook
- sca

Extensions must not change the meaning of core accounting fields, but they may preserve workbook-specific provenance and helper values.

## 17. Validation model

Implement validation in two layers:
1. JSON Schema validation
2. semantic validator rules

The semantic validator should check:
- account uniqueness
- parent-account existence and acyclic hierarchy
- referential integrity
- nonzero transaction lines
- supplemental record resolution
- budget and inventory sanity checks

Balanced multi-line transactions are encouraged when the source system provides enough information, but they are not required for workbook-native ledger exports.

## 18. Versioning

SCLX follows semantic versioning for the format:
- 1.x for backward-compatible additions and policy relaxations
- 2.0 for semantic or structural breaks

Unknown optional sections should be ignored by tolerant readers unless the producer declares a semantic feature that the consumer does not support.

## 19. Recommended implementation order

1. Parse JSON
2. Validate against the schema
3. Index master data by id
4. Validate semantic rules
5. Load transactions
6. Load optional subsidiary collections
7. Preserve unknown extensions when possible

## 20. File naming and encoding

Recommended:
- UTF-8 encoding
- .sclx.json extension

## 21. Security

Treat SCLX as untrusted input:
- cap file size
- limit parser depth
- validate dates and amounts strictly
- reject broken references before posting to a database

## 22. Practical note for spreadsheet integrations

Spreadsheet-backed systems may store row/column provenance in extensions.workbook or workbookLink. This is acceptable and often useful for round-tripping to forms, as long as the core ledger meaning remains in the canonical fields that the producer actually knows.
