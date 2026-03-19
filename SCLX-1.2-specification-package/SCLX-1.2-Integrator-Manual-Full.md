# SCLX 1.2 Integrator Manual

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

Spreadsheet-backed systems may store row/column provenance in extensions.workbook or workbookLink. This is acceptable and often useful for round-tripping to forms, as long as the core accounting meaning remains in the canonical fields.
