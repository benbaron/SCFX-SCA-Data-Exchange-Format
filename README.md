# SCLX

**SCLX** is the **SCALedger Canonical Ledger Exchange Format**: a JSON-based interchange format for double-entry bookkeeping data, nonprofit operational records, and bank-import preservation.

It is designed to support:

- canonical ledger export/import
- spreadsheet and workbook round-tripping
- validation and archival interchange
- OFX-derived bank data preservation
- nonprofit/SCA bookkeeping workflows
- future adapters to CSV, SQL, Beancount/Ledger text, and related tooling

---

## Project Goals

SCLX provides a **single canonical representation** for accounting data that is richer than a bank download and more structured than a workbook alone.

The format aims to:

- represent **balanced double-entry accounting**
- preserve **chart of accounts structure**
- carry **fund, budget, event, person, and document references**
- support **supplemental schedule records**
- preserve **OFX statement and transaction metadata**
- remain **human-readable** and **machine-validatable**
- support **backward-compatible evolution**

---

## Core Concepts

### Canonical ledger model

SCLX treats the ledger as the accounting truth.

Each transaction contains:

- transaction identity
- transaction and posting dates
- description and reference
- status and source
- two or more posting lines
- optional dimensions and extensions

Each line contains:

- one account reference
- debit or credit amount
- optional fund/budget/event/person/document references
- optional tags and supplemental references

### Master data

SCLX carries these core master-data collections:

- `organization`
- `reportingPeriod`
- `chartOfAccounts`
- `funds`
- `budgets`
- `people`
- `events`
- `documents`

### Operational and subsidiary records

SCLX can also include optional collections for operational tracking:

- `bankingItems`
- `outstandingItems`
- `otherAssetItems`
- `assets`
- `supplies`
- `bankStatementImports`

These do **not** replace core accounting postings. They supplement them.

---

## Main Files in This Project

### `sclx-1.2-full.schema.json`

The complete JSON Schema for SCLX 1.2.

This schema defines:

- the full top-level document structure
- all built-in record types
- all enums/literals
- all `$defs`
- OFX-related preservation blocks
- asset/supply/budget subsidiary record types

### `sclx-1.2-validator-rules.json`

Semantic validation rules that go beyond JSON Schema.

These cover rules such as:

- transaction balancing
- chart-of-accounts uniqueness
- parent-account cycle checks
- referential integrity
- supplemental-link resolution
- selected inventory/budget sanity checks

### `SCLX-1.2-Integrator-Manual-Full.md`

A detailed implementation guide for integrators.

This manual explains:

- how to read and write SCLX
- what each top-level collection means
- how OFX metadata is preserved
- how to validate documents
- how to treat optional subsidiary records
- how to handle extensions and versioning

### `SCLX_Ledger_IO.bas`

A VBA module for exporting/importing SCLX to and from the Excel-based **SCA Exchequer Report** workbook.

This macro currently targets:

- `Ledger`
- `Outstanding`
- `Assets&Inventory`
- `Supplies`

and synthesizes a lightweight top-level export for:

- `organization`
- `reportingPeriod`
- `chartOfAccounts`
- `funds`

### `README-SCLX-VBA.txt`

Notes for installing and using the VBA macro package.

---

## SCLX 1.2 Data Model Overview

A typical SCLX document contains:

```text
SCLX document
├── format
├── version
├── exportedAt
├── features
├── compatibility
├── organization
├── reportingPeriod
├── chartOfAccounts
├── funds
├── budgets
├── people
├── events
├── documents
├── transactions
├── bankingItems
├── outstandingItems
├── otherAssetItems
├── assets
├── supplies
├── bankStatementImports
└── extensions
```

---

## Record Types Included

### Chart of accounts

Accounts use the de facto row model:

- `Number`
- `Name`
- `Type`
- `Parent`
- `IncreaseSide`
- `OpeningBalance`
- `SupplementalKinds`

with optional richer fields such as:

- `accountId`
- `code`
- `subtype`
- `active`
- `reportingTags`

### Transactions

Transactions are balanced double-entry journal entries. Each transaction must contain at least two lines and total debits must equal total credits.

### Banking items

Bank-side records such as:

- checks
- deposits
- bank fees
- interest
- adjustments

These are useful for reconciliation and bank import preservation.

### Outstanding items

Schedule-style records for:

- outstanding checks
- deposits in transit
- transfers
- incoming checks
- card-related items

### Other asset items

Schedule-style records for:

- cash advances
- site security deposits
- other recoverable amounts

### Assets

Registry-style records for durable items such as:

- regalia
- equipment
- furniture
- banners
- site equipment
- loaner gear

### Supplies

Registry-style records for lower-value or consumable items that may still require guardian tracking and removal history.

### Budgets

SCLX supports both budget headers and budget lines, including:

- event name
- budgeted amount
- revenue category
- expense category
- account linkage
- notes

### OFX statement imports

Statement-level metadata for imported OFX data, including:

- statement range
- bank account identity
- ledger balance
- available balance
- source version
- document linkage

---

## Validation Model

SCLX validation is intended to happen in **two layers**.

### 1. JSON Schema validation

Use the schema file to validate document structure, field shapes, and allowed literals.

### 2. Semantic validation

Use the validator rules file for checks such as:

- account number uniqueness
- valid parent references
- acyclic chart hierarchy
- line/account/fund/budget/person/document resolution
- balanced transactions
- supplemental record linkage
- banking and OFX sanity checks
- asset/supply sanity checks

---

## Versioning

SCLX uses semantic format versioning.

- **minor versions** add backward-compatible optional structure
- **major versions** introduce incompatible semantic or structural changes

SCLX 1.2 is intended to remain readable by tolerant consumers that ignore unknown optional sections and unknown extensions.

---

## OFX Preservation Strategy

SCLX is **not** a general replacement for OFX transport or session semantics.

Instead, it preserves the parts of OFX that are valuable for bookkeeping:

### Statement-level preservation

Stored in `bankStatementImports`:

- source format and version
- bank account identity
- statement start/end dates
- ledger balance snapshot
- available balance snapshot

### Transaction-level preservation

Stored in `bankingItems[].ofx`:

- `fitId`
- `transactionType`
- `datePosted`
- `dateUser`
- `dateAvailable`
- `checkNumber`
- `referenceNumber`
- `name`
- `memo`
- `payeeId`
- `sic`
- `correctFitId`
- `correctAction`

This lets an SCLX consumer preserve high-value bank facts without becoming a full OFX protocol engine.

---

## Excel / VBA Integration

This project also includes a VBA-based bridge for an Excel workbook workflow.

### Requirements

The VBA module requires:

- Excel VBA
- `JsonConverter.bas` from the **VBA-JSON** library

### Current workbook integration scope

The macro supports export/import for:

- Ledger rows
- Outstanding items
- Assets & Inventory
- Supplies

It leaves workbook-specific details in `extensions.workbook` blocks where useful for round-tripping.

### Notes

The workbook's budget tab is formula-heavy and may not yet map cleanly to a stable structural import/export region. That can be expanded later with a finalized range map.

---

## Suggested Project Layout

A practical repository layout could look like this:

```text
sclx/
├── README.md
├── schema/
│   └── sclx-1.2-full.schema.json
├── validator/
│   └── sclx-1.2-validator-rules.json
├── docs/
│   └── SCLX-1.2-Integrator-Manual-Full.md
├── excel/
│   ├── SCLX_Ledger_IO.bas
│   └── README-SCLX-VBA.txt
├── examples/
│   ├── minimal-ledger.sclx.json
│   ├── full-ledger-with-ofx.sclx.json
│   └── workbook-roundtrip.sclx.json
└── tests/
    ├── valid/
    └── invalid/
```

---

## Recommended Next Steps

To make SCLX easier to adopt in real tools, the next most useful additions would be:

- sample valid and invalid SCLX files
- a reference validator implementation in Java or Python
- a reference import/export library
- conformance tests
- a formal RFC-style written specification
- JSON examples for each top-level record type
- a workbook mapping document for each spreadsheet page

---

## Intended Use Cases

SCLX is suitable for:

- ledger archival
- workbook data interchange
- migration between bookkeeping systems
- nonprofit accounting tooling
- reconciliation workflows
- audit and compliance review
- canonical storage for custom accounting applications

---

## Design Philosophy

SCLX separates:

- **accounting truth**
- **operational schedules**
- **source-format preservation**

That means:

- transactions remain canonical
- bank imports are preserved without replacing the ledger
- assets/supplies/budgets remain linked but distinct
- extensions can carry local workflow metadata without polluting the core model

---

## License / Project Status

This repository currently contains a specification draft, supporting artifacts, and spreadsheet integration code. Add your preferred license and governance model here once you decide how you want SCLX distributed.

Suggested placeholders:

- license
- contribution guide
- change log
- version support policy

---

## Quick Start

### Validate an SCLX file

1. Load `sclx-1.2-full.schema.json`
2. Validate a candidate SCLX document structurally
3. Apply `sclx-1.2-validator-rules.json`
4. Reject or warn based on semantic errors

### Use with Excel workbook integration

1. Import `JsonConverter.bas` into the workbook VBA project
2. Import `SCLX_Ledger_IO.bas`
3. Save workbook as `.xlsm`
4. Run `ExportSCLX` or `ImportSCLX`

---

## Summary

SCLX 1.2 is a **canonical accounting interchange format** that combines:

- double-entry ledger fidelity
- master-data structure
- supplemental schedules
- inventory-style subsidiary records
- OFX-preserving bank metadata
- spreadsheet-friendly round-trip support

in a single JSON-based format that is easier to validate, extend, and archive than ad hoc workbook-only solutions.
