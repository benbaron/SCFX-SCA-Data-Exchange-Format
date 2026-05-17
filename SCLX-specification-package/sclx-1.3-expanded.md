# SCLX 1.3 (Ledger-Native Expanded Overview)

**SCALedger Canonical Ledger Exchange Format**

## 1. Purpose

SCLX is a **canonical interchange format** for bookkeeping data. It is designed to:

* represent true double-entry accounting
* preserve audit-trail and source-document context
* carry nonprofit and SCA dimensions such as fund, budget, event, approvals, and reporting tags
* map cleanly to Java/Jackson and SQL
* export outward to CSV, Beancount/Ledger text, and later XBRL-GL, while treating OFX as an import source rather than a full accounting target. 

## 2. Scope

One SCLX file may represent:

* a full organization export
* a date-bounded ledger slice
* an import batch

The top-level document contains both master data and transactions. 

---

# 3. File format rules

## 3.1 Encoding

* UTF-8
* JSON text
* line endings may be LF or CRLF
* file extension: `.sclx.json`

## 3.2 Top-level required fields

```json
{
  "format": "SCLX",
  "version": "1.3",
  "exportedAt": "2026-03-17T14:25:00-06:00",
  "organization": {},
  "reportingPeriod": {},
  "chartOfAccounts": [],
  "funds": [],
  "budgets": [],
  "people": [],
  "events": [],
  "documents": [],
  "transactions": []
}
```

This top-level model follows the proposed canonical structure. 

## 3.3 Required top-level semantics

* `format` must be exactly `SCLX`
* `version` must be a supported semantic version string
* `exportedAt` must be an ISO-8601 offset timestamp
* arrays may be empty, but must exist unless a future minor version marks them optional
* IDs referenced anywhere in the file must resolve to an object present in the same file or in an explicitly supported external registry

---

# 4. Data model

## 4.1 Organization object

Required fields:

* `organizationId`
* `name`
* `baseCurrency`
* `fiscalYearStart`
* `fiscalYearEnd`

Recommended fields:

* `parentOrganization`

Example:

```json
{
  "organizationId": "botm-fy2026",
  "name": "Barony of the Monkey",
  "parentOrganization": "Kingdom of the Stag",
  "baseCurrency": "USD",
  "fiscalYearStart": "2026-01-01",
  "fiscalYearEnd": "2026-12-31"
}
```

This aligns with your draft organization structure. 

## 4.2 Reporting period object

Required fields:

* `startDate`
* `endDate`

Recommended fields:

* `label`
* `fiscalYear`
* `periodType`

Example:

```json
{
  "startDate": "2026-01-01",
  "endDate": "2026-12-31",
  "label": "FY2026",
  "fiscalYear": 2026,
  "periodType": "FISCAL_YEAR"
}
```

## 4.3 Chart of accounts

Each account must have a stable internal ID and a preserved display name. Your draft explicitly recommends immutable `accountId` and preserving `name` exactly for reporting. 

Required fields:

* `accountId`
* `name`
* `type`
* `normalBalance`
* `active`

Recommended fields:

* `code`
* `subtype`
* `parentAccountId`
* `reportingTags`

Example:

```json
{
  "accountId": "1000",
  "code": "1000",
  "name": "Checking",
  "type": "ASSET",
  "subtype": "BANK",
  "normalBalance": "DEBIT",
  "active": true,
  "parentAccountId": null,
  "reportingTags": ["BALANCE_SHEET", "CASH"]
}
```

## 4.4 Funds

Funds are first-class tracking dimensions, not accounts. 

Required fields:

* `fundId`
* `name`
* `restricted`

Recommended fields:

* `description`

## 4.5 Budgets

Required fields:

* `budgetId`
* `name`
* `fiscalYear`
* `fundId`
* `active`

## 4.6 People / counterparties

Required fields:

* `personId`
* `displayName`
* `kind`



## 4.6A Bank accounts

SCLX may include a `bankAccounts` collection for bank-account master data such as:
- institution name and contact details
- masked account number
- account type
- account holder name
- interest-bearing flag
- signature requirements
- authorized signers
- linkage to the chart-of-accounts bank account

These are not the same thing as `bankingItems` or `bankStatementImports`.

## 4.6B Governance records

SCLX may include:
- `officeAssignments`
- `committeeMemberships`

These preserve workbook or system governance records such as officers and financial committee members. They reference `people` records rather than replacing them.

## 4.7 Events

Required fields:

* `eventId`
* `name`

Recommended fields:

* `startDate`
* `endDate`
* `hostingOrganizationId`

## 4.8 Documents

Documents are part of the audit trail. 

Required fields:

* `documentId`
* `documentType`

Recommended fields:

* `referenceNumber`
* `documentDate`
* `fileName`
* `notes`

---



## 4.9 Bank statement imports, banking items, assets, supplies, and supplemental items

SCLX may also include:
- `bankStatementImports`
- `bankingItems`
- `outstandingItems`
- `otherAssetItems`
- `supplementalItems`
- `assets`
- `supplies`

These collections preserve operational, reconciliation, schedule, inventory, and custodianship data around the ledger.

Assets may carry workbook-native `itemType` values and may also include `appraisalDetails` such as appraiser name, appraisal date, and revised value.

# 5. Transaction model

The transaction is the core accounting unit. It contains a header and balanced posting lines. 

## 5.1 Transaction header

Required fields:

* `transactionId`
* `transactionDate`
* `postingDate`
* `description`
* `status`
* `source`
* `lines`

Recommended fields:

* `reference`
* `documentIds`
* `eventId`
* `approval`
* `bankTiming`
* `budgetTiming`
* `extensions`

* `budgetId`
* `workbookLink`

SCLX 1.3 adds `transaction.budgetId` for spreadsheet and ledger systems that assign a single budget category at the transaction/header level. This avoids forcing exporters to duplicate one row-level budget label onto every posting line when the source workbook really stores it at the row level.

`transaction.workbookLink` may be used to preserve stable sheet/row provenance for round-tripping.


Example:

```json
{
  "transactionId": "txn-2026-000123",
  "transactionDate": "2026-03-15",
  "postingDate": "2026-03-15",
  "description": "Office supplies purchase",
  "reference": "CHK-1042",
  "status": "POSTED",
  "source": "MANUAL",
  "documentIds": ["doc-2026-0041"],
  "eventId": null,
  "approval": {
    "policyRequired": true,
    "committeeApprovalRef": "FIN-2026-03-02-A",
    "approvedBy": ["financial-committee"],
    "approvalDate": "2026-03-02"
  },
  "lines": []
}
```

## 5.2 Transaction lines

Each line is one posting only. Your draft requires exactly one account per line and forbids a line from carrying both debit and credit as nonzero. 

Required fields:

* `lineId`
* `accountId`
* `debit`
* `credit`

Recommended fields:

* `description`
* `fundId`
* `budgetId`
* `eventId`
* `personId`
* `documentId`
* `memo`
* `tags`
* `restrictionTag`
* `reportSection`

* `usedFor`
* `itemNumber`
* `quantity`
* `workbookLink`

SCLX 1.3 adds these optional structured line fields for spreadsheet split columns. They allow producer-neutral carriage of common workbook details without forcing everything into `extensions`.


Example:

```json
{
  "lineId": "txn-2026-000123-ln1",
  "accountId": "6100",
  "description": "Office supplies",
  "debit": "20.00",
  "credit": "0.00",
  "fundId": "general",
  "budgetId": "office-supplies",
  "personId": "p-office-depot",
  "documentId": "doc-2026-0041",
  "memo": "Printer paper and pens",
  "tags": ["SUPPLIES"]
}
```

---


## 5.3 Spreadsheet split-line mapping

Many spreadsheet ledgers store one visible transaction row plus multiple split posting regions. In SCLX, that should be modeled as one `transaction` containing multiple `lines`.

A split region should only produce a `transactionLine` when it contains accounting content such as an amount, income category, or expense category. Descriptive helper fields alone should not create a posting line.

For round-tripping, use:
- `transaction.workbookLink` for row-level provenance
- `transactionLine.workbookLink` for split-line provenance
- `extensions.workbook` for richer producer-specific detail such as visible row numbers, split indexes, or source column letters


# 6. Required accounting rules

These are mandatory validator rules, directly matching your draft. 

## 6.1 Per transaction

* must contain at least 2 lines
* total debits must equal total credits
* every line must reference exactly one account
* a line must not have both debit and credit nonzero
* all referenced IDs must exist
* `status` must be a valid enum value

## 6.2 Amount rules

* amounts must be strings, not JSON floating-point numbers
* amounts must use fixed decimal format
* recommended regex: `^-?[0-9]+\.[0-9]{2}$`
* canonical debit and credit values should normally be nonnegative strings
* negative signed values are disallowed in debit/credit fields; sign belongs to which side is populated

## 6.3 Date rules

* `transactionDate` and `postingDate` must be `YYYY-MM-DD`
* `exportedAt` must be ISO-8601 with offset
* if present, `approvalDate`, `documentDate`, `startDate`, `endDate` must also be `YYYY-MM-DD`

---

# 7. Enums

Your draft recommends explicit enums to avoid drift. 

## 7.1 Account type

```text
ASSET
LIABILITY
NET_ASSETS
REVENUE
EXPENSE
```

## 7.2 Normal balance

```text
DEBIT
CREDIT
```

## 7.3 Transaction status

```text
DRAFT
POSTED
VOID
REVERSED
```

## 7.4 Source

```text
MANUAL
BANK_IMPORT
OFX_IMPORT
CSV_IMPORT
OPENING_BALANCE
SYSTEM_GENERATED
ADJUSTMENT
```

## 7.5 Person kind

```text
VENDOR
CUSTOMER
MEMBER
OFFICER
BRANCH
OTHER
```

## 7.6 Document type

```text
RECEIPT
INVOICE
CHECK_IMAGE
DEPOSIT_SLIP
BANK_STATEMENT
CONTRACT
APPROVAL_RECORD
OTHER
```

## 7.7 Timing metadata

These were proposed to hold workbook logic without distorting the accounting entry itself. 

```text
NOW
PREVIOUSLY
LATER
NONE
```

Use for:

* `bankTiming`
* `budgetTiming`

---

# 8. Extensions policy

Your draft separates core from extensions and recommends using extension blocks for SCA-specific detail while keeping the ledger core portable. 

## 8.1 Core

Keep these in core:

* accounts
* transactions
* lines
* funds
* budgets
* people
* documents
* events
* approval metadata

## 8.2 Extensions

Put these in `extensions`:

* exchequer workbook page mappings
* operator-guide references
* import diagnostics
* UI-only hints
* JRXML/report metadata

Example:

```json
{
  "extensions": {
    "sca": {
      "kingdom": "Kingdom of the Stag",
      "reportFormYear": 2026
    }
  }
}
```

---

# 9. Naming and ID stylesheet

This is the naming standard I recommend for SCLX 1.0.

## 9.1 General naming

* JSON property names: lowerCamelCase
* enum values: UPPER_SNAKE_CASE except timing values, which may remain simple uppercase words
* IDs: lowercase kebab-case unless numeric codes already exist
* human labels: preserve exact user-facing strings

## 9.2 ID prefixes

Use stable prefixes to avoid collisions:

```text
org-
acct-
fund-
budget-
person-
event-
doc-
txn-
line-
```

Examples:

```text
org-botm-fy2026
acct-1000
fund-general
budget-meeting-fight-space
person-vendor-scholars-abode
event-crown-tournament-2026
doc-lease-q1
txn-2026-000001
line-txn-2026-000001-ln1
```

## 9.3 Account naming

For your environment, preserve canonical account names exactly in user-facing `name` fields. That matches your long-term requirement to preserve account strings exactly. Internal `accountId` may still be normalized.

Recommended split:

* `accountId`: stable machine key
* `code`: optional external or chart code
* `name`: exact reporting label

## 9.4 Money fields

* use strings
* 2 decimal places
* no thousands separators
* no currency symbols

Good:

```json
"debit": "300.00"
```

Bad:

```json
"debit": 300
"debit": "$300.00"
"debit": "300"
```

---

# 10. File stylesheet

This is the formatting standard for `.sclx.json` files.

## 10.1 Object ordering

Use this order at the top level:

1. `format`
2. `version`
3. `exportedAt`
4. `organization`
5. `reportingPeriod`
6. `chartOfAccounts`
7. `funds`
8. `budgets`
9. `people`
10. `events`
11. `documents`
12. `transactions`
13. `extensions`

## 10.2 Within arrays

Sort by stable primary key:

* accounts by `code`, then `accountId`
* funds by `fundId`
* budgets by `budgetId`
* people by `personId`
* events by `eventId`
* documents by `documentId`
* transactions by `postingDate`, then `transactionId`
* lines in source posting order

## 10.3 Indentation

* 2 spaces
* no tabs

## 10.4 Null handling

* omit optional null fields if not used
* use explicit `null` only when semantic distinction matters

## 10.5 Comments

* JSON comments are not allowed
* use `notes`, `memo`, or `extensions.importDiagnostics` instead

---

# 11. Producer instructions

## 11.1 When creating an export

1. Emit all master data referenced by the transactions.
2. Emit stable immutable IDs.
3. Preserve user-facing labels exactly.
4. Serialize money as decimal strings.
5. Validate structural, referential, and accounting rules before writing.
6. Include workbook timing logic as metadata, not as substitute accounting lines. This matches the recommendation that real accounting truth stay canonical while workbook presentation logic remains metadata. 

## 11.2 When importing OFX

* treat OFX as a transaction candidate source
* do not assume OFX is a complete ledger format
* enrich imported bank items into full canonical transactions after matching payee, account, document, fund, budget, and approvals where possible. Your draft explicitly treats OFX as an input source rather than a full target. 

## 11.3 When importing CSV

* map each row either to a whole transaction or to a transaction line in a grouped transaction
* reject rows that would produce an unbalanced transaction unless import mode is explicitly staged/draft

## 11.4 When exporting CSV

Flatten each posting line into a row with at least:

* transactionId
* transactionDate
* description
* reference
* accountId
* accountName
* debit
* credit
* fundId
* budgetId
* eventId
* personId
* memo

This follows the mapping guidance in your draft. 

---

# 12. Consumer instructions

## 12.1 Minimum compliant reader

A compliant reader must:

* parse the file as UTF-8 JSON
* verify `format` and supported `version`
* load master data
* resolve IDs
* verify each transaction balances
* ignore unknown optional fields
* preserve unknown extension blocks when round-tripping, if possible

## 12.2 Error handling

Errors should be classified as:

* structural
* referential
* accounting
* policy
* extension-specific

Recommended behavior:

* hard-fail on structural, referential, and accounting errors
* configurable fail or warn on policy/extension errors

---

# 13. Versioning instructions

Your draft recommends semantic versioning with minor versions for additive optional fields and major versions for semantic breaks. 

## 13.1 Rules

* `1.0` → initial release
* `1.1` / `1.2` → additive clarifications and optional structures
* `1.3` → additive transaction-level budget and workbook-link support, plus structured split-support line fields
* `2.0` → changed required semantics or incompatible structure

## 13.2 Compatibility

* readers of `1.x` must ignore unknown optional properties
* writers should not emit `2.x` unless the target explicitly supports it
* 1.3 readers should continue to accept 1.2 files without upgrade

---

# 14. Canonical example

```json
{
  "format": "SCLX",
  "version": "1.3",
  "exportedAt": "2026-03-17T14:25:00-06:00",
  "organization": {
    "organizationId": "org-botm-fy2026",
    "name": "Barony of the Monkey",
    "parentOrganization": "Kingdom of the Stag",
    "baseCurrency": "USD",
    "fiscalYearStart": "2026-01-01",
    "fiscalYearEnd": "2026-12-31"
  },
  "reportingPeriod": {
    "startDate": "2026-01-01",
    "endDate": "2026-12-31",
    "label": "FY2026",
    "fiscalYear": 2026,
    "periodType": "FISCAL_YEAR"
  },
  "chartOfAccounts": [
    {
      "accountId": "acct-1000",
      "code": "1000",
      "name": "Checking",
      "type": "ASSET",
      "subtype": "BANK",
      "normalBalance": "DEBIT",
      "active": true
    },
    {
      "accountId": "acct-7200",
      "code": "7200",
      "name": "Site/Storage Rental (Occupancy)",
      "type": "EXPENSE",
      "subtype": "OCCUPANCY",
      "normalBalance": "DEBIT",
      "active": true
    }
  ],
  "funds": [
    {
      "fundId": "fund-general",
      "name": "General Fund",
      "restricted": false
    }
  ],
  "budgets": [
    {
      "budgetId": "budget-meeting-fight-space",
      "name": "Meeting & Fight Space",
      "fiscalYear": 2026,
      "fundId": "fund-general",
      "active": true
    }
  ],
  "people": [
    {
      "personId": "person-vendor-scholars-abode",
      "displayName": "The Scholar's Abode",
      "kind": "VENDOR"
    }
  ],
  "events": [],
  "documents": [
    {
      "documentId": "doc-lease-q1",
      "documentType": "INVOICE",
      "referenceNumber": "Q1-2026-SPACE",
      "documentDate": "2025-12-20",
      "fileName": "scholars-abode-q1-2026.pdf"
    }
  ],
  "transactions": [
    {
      "transactionId": "txn-2026-000001",
      "transactionDate": "2026-01-01",
      "postingDate": "2026-01-01",
      "description": "Q1 meeting space payment",
      "reference": "CHK-1001",
      "status": "POSTED",
      "source": "MANUAL",
      "documentIds": ["doc-lease-q1"],
      "bankTiming": "PREVIOUSLY",
      "budgetTiming": "NOW",
      "approval": {
        "policyRequired": true,
        "committeeApprovalRef": "BOTM-FIN-2025-12-15-01",
        "approvedBy": ["financial-committee"],
        "approvalDate": "2025-12-15"
      },
      "lines": [
        {
          "lineId": "line-txn-2026-000001-ln1",
          "accountId": "acct-7200",
          "description": "Meeting space for Q1",
          "debit": "300.00",
          "credit": "0.00",
          "fundId": "fund-general",
          "budgetId": "budget-meeting-fight-space",
          "personId": "person-vendor-scholars-abode",
          "documentId": "doc-lease-q1",
          "memo": "Occupancy expense recognized in current budget period"
        },
        {
          "lineId": "line-txn-2026-000001-ln2",
          "accountId": "acct-1000",
          "description": "Checking offset",
          "debit": "0.00",
          "credit": "300.00",
          "fundId": "fund-general",
          "personId": "person-vendor-scholars-abode",
          "documentId": "doc-lease-q1",
          "memo": "Bank effect treated as prior-period activity by workbook logic"
        }
      ]
    }
  ],
  "extensions": {
    "sca": {
      "kingdom": "Kingdom of the Stag",
      "reportFormYear": 2026
    }
  }
}
```

---

# 15. Short implementation guidance for Java

Your draft already notes the natural Java mapping and recommends `BigDecimal` for money, `LocalDate` for transaction dates, and `OffsetDateTime` for export timestamps. 

Use:

* Jackson for JSON binding
* immutable records or Lombok builders
* a 3-layer validator:

  * structural
  * referential
  * accounting

---

# 16. One-sentence stylesheet

**SCLX files are UTF-8, 2-space-indented JSON documents using lowerCamelCase properties, stable prefixed IDs, exact preserved reporting names, decimal-string money fields, balanced transaction lines, and extension blocks for SCA/workbook-specific metadata.**

If you want, the next step is to turn this into a formal **JSON Schema draft** and matching **Java record classes**.
