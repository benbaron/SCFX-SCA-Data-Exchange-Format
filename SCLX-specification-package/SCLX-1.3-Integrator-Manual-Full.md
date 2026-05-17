# SCLX 1.3 Integrator Manual

## 1. Purpose

SCLX is a canonical JSON exchange format for double-entry bookkeeping data plus selected nonprofit operational records. It is intended for ledger export/import, audit review, archival interchange, spreadsheet integration, and bank import preservation.

## 2. What belongs in SCLX

SCLX contains:
- core ledger master data
- balanced journal transactions
- optional bank reconciliation records
- optional outstanding-item and other-asset schedules
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
- bankingItems
- outstandingItems
- otherAssetItems
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

Transactions are the core accounting records. Each transaction must:
- have at least two lines
- balance debits and credits
- contain only nonnegative debit and credit values
- use exactly one nonzero side per line

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
- extensions



## 6A. Transaction-level budgets and spreadsheet ledger headers

SCLX 1.3 adds an optional `transaction.budgetId` field.

Use `transaction.budgetId` when the source system stores a single budget category at the transaction header or spreadsheet row level rather than at the individual posting-line level. This is common in spreadsheet ledger forms where a row has one visible budget category but up to several accounting split lines.

Line-level `budgetId` remains valid and should still be used when a source system truly assigns different budget categories to different posting lines within the same transaction.

## 6B. Ledger split columns and populated-line rules

Spreadsheet-backed ledgers may store one transaction header plus multiple split posting regions on a single visible row.

In SCLX:
- one spreadsheet ledger row may map to one `transaction`
- each populated split region maps to one `transactionLine`

A split region is considered populated only when it contains accounting content, such as:
- an amount
- an income category
- an expense category

Descriptive helper fields by themselves do not create a posting line. For example, `usedFor`, `itemNumber`, or `quantity` without an amount/category should not cause an exporter to emit a zero/zero line.

Exporters must continue to satisfy the core accounting rules:
- at least two lines per transaction
- one nonzero side per line
- total debits equal total credits

## 6C. Structured split-support fields

SCLX 1.3 adds the following optional first-class properties to `transactionLine`:
- `usedFor`
- `itemNumber`
- `quantity`

These are intended for spreadsheet and form integrations that carry structured operational detail alongside posting lines. Exporters may still preserve richer or producer-specific detail under `extensions`.

## 6D. Workbook linking on transactions and lines

SCLX 1.3 allows `workbookLink` on:
- `transaction`
- `transactionLine`

This is in addition to existing use of `workbookLink` on supplemental schedule-style records such as `outstandingItems` and `otherAssetItems`.

Use `workbookLink` for stable sheet/row provenance. Use `extensions.workbook` for richer producer-specific spreadsheet metadata such as:
- visible row number
- split index
- source column letters
- workbook-specific timing or presentation fields


## 7. Funds, budgets, events, people, documents

These are all first-class master data collections. Readers should load them before validating line references.

Budgets may be empty if the source system has no stable budget-entry region or does not track budgets structurally.

## 8. Banking items

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

## 9. OFX preservation

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

## 10. Outstanding items

Outstanding items are supplemental schedule records linked back to a transaction line using ledgerLink. They may track:
- outstanding checks
- deposits in transit
- transfers
- incoming checks
- card items

These records do not replace accounting entries.

## 11. Other asset items

otherAssetItems represent schedule-style items such as:
- cash advances
- site security deposits
- other recoverable amounts

Use typeCode values:
- C = CASH_ADVANCE
- S = SITE_SECURITY_DEPOSIT
- O = OTHER

## 12. Assets registry

assets records track durable property and custodianship. Suggested uses:
- regalia
- furniture
- banners
- site equipment
- loaner gear

These are operational records, not accounting postings. Link them to relatedTransactionIds or relatedLineIds when the acquisition, sale, or write-off appears in the ledger.

## 13. Supplies registry

supplies records track lower-value or consumable items that may still have guardianship and removal history. Use removalDetails to track sold, lost, donated, destroyed, or returned items.

## 14. Extensions

Any object may include an extensions object. Extensions should be namespaced by producer, for example:
- scaledger
- workbook
- sca

Extensions must not change the meaning of core accounting fields.

## 15. Validation model

Implement validation in two layers:
1. JSON Schema validation
2. semantic validator rules

The semantic validator should check:
- account uniqueness
- parent-account existence and acyclic hierarchy
- referential integrity
- balanced transactions
- supplemental record resolution
- budget and inventory sanity checks

## 16. Versioning

SCLX follows semantic versioning for the format:
- 1.x for backward-compatible additions
- 2.0 for semantic or structural breaks

Unknown optional sections should be ignored by tolerant readers unless the producer declares a semantic feature that the consumer does not support.

SCLX 1.3 is a backward-compatible additive release. Readers that already support tolerant 1.x parsing should preserve unknown optional fields where possible and may ignore the new `transaction.budgetId`, `transaction.workbookLink`, `transactionLine.usedFor`, `transactionLine.itemNumber`, `transactionLine.quantity`, and `transactionLine.workbookLink` fields if they do not use them.

## 17. Recommended implementation order

1. Parse JSON
2. Validate against the schema
3. Index master data by id
4. Validate semantic rules
5. Load transactions
6. Load optional subsidiary collections
7. Preserve unknown extensions when possible

## 18. File naming and encoding

Recommended:
- UTF-8 encoding
- .sclx.json extension

## 19. Security

Treat SCLX as untrusted input:
- cap file size
- limit parser depth
- validate dates and amounts strictly
- reject broken references before posting to a database

## 20. Practical note for spreadsheet integrations

Spreadsheet-backed systems may store row/column provenance in extensions.workbook or workbookLink. In SCLX 1.3, `workbookLink` is also allowed directly on transactions and transaction lines. This is acceptable and often useful for round-tripping to forms, as long as the core accounting meaning remains in the canonical fields.
