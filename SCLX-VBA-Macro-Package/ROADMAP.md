Below is a roadmap draft you can place into the SCLX archive as something like `ROADMAP.md`.

---

# SCLX Roadmap

## Purpose

This roadmap records the planned evolution of the SCLX format and the Excel/VBA workbook bridge used to export and import SCLX files from the SCA Exchequer workbook layout.

It covers two related tracks:

1. **Schema and documentation evolution**
2. **Macro and workbook-bridge evolution**

The goal is to move from the current practical workbook bridge toward a more complete, authoritative, and round-trip-safe interchange package.

---

## Current state

### Schema state

The archive currently contains:

* SCLX 1.2 source documents
* a conservative SCLX 1.3 extension set
* updated schema, validator, and integrator-manual files for the 1.3 workbook bridge

SCLX 1.3 currently adds or formalizes:

* `transaction.budgetId`
* `transaction.workbookLink`
* `transactionLine.usedFor`
* `transactionLine.itemNumber`
* `transactionLine.quantity`
* `transactionLine.workbookLink`

These additions were made to support the actual spreadsheet structure, especially the ledger split columns.

### Macro state

The current reviewed macro:

* exports and imports the core workbook tabs now mapped
* uses corrected ledger split mappings
* synthesizes funds, budgets, people, and chart-of-accounts data where necessary
* preserves workbook provenance in workbook-link / extension data
* supports SCLX 1.3 output
* remains intentionally partial in several areas

### What is still intentionally incomplete

The current bridge is still lightweight in these areas:

* direct export of a canonical `Accounts` sheet
* direct export/import of the real `Budget` sheet structure
* material `bankingItems` population
* material `bankStatementImports` population
* `events` and `documents` master-data generation
* direct handling of other workbook detail sheets such as liability/asset detail schedules, where applicable
* authoritative rather than synthesized chart of accounts

---

## Design principles for future work

Future work should follow these principles.

### 1. Preserve backward compatibility where possible

New schema versions should be additive unless there is a compelling reason to break compatibility.

### 2. Keep canonical accounting data distinct from workbook-specific metadata

Workbook row numbers, split indices, visible row numbers, and spreadsheet-specific mappings should remain in:

* `workbookLink`
* `extensions.workbook`

unless they clearly deserve promotion into the canonical SCLX model.

### 3. Prefer authoritative workbook sources over synthesis

When the workbook contains a true master table for accounts, budgets, people, or banking items, future exporters should use that authoritative source rather than guessing from observed ledger usage.

### 4. Favor round-trip safety over elegance

A bridge that reconstructs the workbook faithfully is more valuable than a theoretically cleaner model that loses spreadsheet detail.

### 5. Validate aggressively

Each schema or macro expansion should come with:

* example files
* validation rules
* round-trip tests
* clear migration notes

---

## Roadmap structure

This roadmap is divided into phases. The phases are cumulative. A later phase does not invalidate earlier work unless explicitly noted.

---

# Phase 1 — Stabilize the current SCLX 1.3 workbook bridge

## Goal

Lock down the current 1.3 bridge so it is reliable, documented, and reproducible.

## Schema and documentation work

### 1. Finalize the SCLX 1.3 package

Complete and freeze the 1.3 archive contents:

* `sclx-1.3-full.schema.json`
* `sclx-1.3-validator-rules.json`
* `SCLX-1.3-Integrator-Manual-Full.md`
* `sclx-1.3.md`

### 2. Add migration notes

Add a formal section describing:

* what changed from 1.2 to 1.3
* why the new fields were introduced
* what remains in `extensions.workbook`
* how 1.2 consumers should handle 1.3 files

### 3. Add workbook-bridge guidance

The integrator manual should explicitly document:

* mapped workbook sheets
* mapped row anchors
* ledger split block layout
* workbookLink usage
* what is synthesized versus authoritative

## Macro work

### 1. Freeze the reviewed baseline

Treat the current reviewed 1.3 VBA module as the baseline working version.

### 2. Add internal validation helpers

Add VBA routines that check workbook assumptions before import/export, such as:

* required sheet names exist
* expected key cells are present
* split columns are in the expected positions
* row anchors are valid

### 3. Add a diagnostics mode

Add optional debug logging or a diagnostic export sheet showing:

* discovered last rows
* found budgets
* found people
* found accounts
* skipped rows and reasons

### 4. Create smoke-test procedures

Add optional VBA test procedures such as:

* `TestExportSCLX`
* `TestImportSCLX`
* `TestValidateWorkbookLayout`

## Deliverables

* stable 1.3 schema package
* stable 1.3 VBA module
* workbook mapping notes
* known limitations section

---

# Phase 2 — Authoritative master-data export

## Goal

Reduce synthesis and increase fidelity by exporting real workbook master-data tables where present.

## Schema and documentation work

### 1. Clarify canonical versus synthesized sources

The manual should define the precedence order for export:

1. authoritative workbook master table
2. stable workbook schedule/detail tab
3. synthesized fallback from transactional usage

### 2. Extend examples

Add example files showing:

* canonical accounts export
* canonical budgets export
* canonical people export

## Macro work

### 1. Export Accounts from a real Accounts sheet

Replace synthesized chart-of-accounts export when a proper `Accounts` sheet exists.

Future macro behavior should:

* read all accounts, not only used accounts
* preserve parent relationships
* preserve canonical number and name
* preserve type and increase side from source
* preserve supplemental kinds where supplied

### 2. Export budgets from a real Budget sheet

Replace or supplement synthesized budget creation from ledger categories.

Future macro behavior should:

* read actual budget rows
* preserve fund linkage
* preserve event linkage if present
* preserve status, effective period, and line categories where available

### 3. Build people from authoritative lists when present

If workbook sheets such as `Exchequers` and `FinancialCommittee` exist, use them as the primary source for people and role assignments.

The fallback synthesis from transaction names may remain, but only when no authoritative people table exists.

### 4. Improve ID generation

Standardize all generated IDs for:

* accounts
* funds
* budgets
* people
* transactions
* lines

and document normalization rules.

## Deliverables

* macro support for authoritative Accounts and Budget sheets
* richer people export
* reduced reliance on synthesis
* improved referential consistency

---

# Phase 3 — Banking and reconciliation support

## Goal

Implement the banking side of SCLX more fully, especially check, transfer, OFX, and reconciliation detail.

## Schema and documentation work

### 1. Clarify banking collections

Document intended distinctions among:

* `transactions`
* `bankingItems`
* `outstandingItems`
* `bankStatementImports`

This should explain when the same real-world activity appears in one collection versus another.

### 2. Add OFX preservation guidance

The manual should explain how OFX or bank-import payloads are preserved, including:

* raw OFX retention
* normalized transaction extraction
* linkages between bank statement lines and SCLX transaction records

### 3. Consider future schema additions only if needed

Before adding new schema fields, test whether banking needs can be met by:

* existing top-level collections
* `workbookLink`
* `extensions`

Schema changes should come only after a concrete mapping pass.

## Macro work

### 1. Populate `bankingItems`

Future macro work should derive or export:

* check numbers
* transfer references
* issue dates
* cleared dates
* statuses
* linked transaction IDs

### 2. Populate `bankStatementImports`

Where the workbook or workflow includes imported bank statement detail, preserve it in SCLX.

### 3. Improve outstanding linkage

Link outstanding items to:

* check or transfer references
* related banking items
* related ledger transactions

### 4. Add reconciliation checks

Implement validations such as:

* check/reference uniqueness where appropriate
* outstanding-versus-cleared consistency
* bank-line-to-ledger matching diagnostics

## Deliverables

* first materially populated banking collections
* clearer reconciliation mapping
* OFX/bank import preservation path

---

# Phase 4 — Events, documents, and richer references

## Goal

Strengthen the referential model so exported transactions point to actual top-level master-data collections rather than raw text only.

## Schema and documentation work

### 1. Clarify when to materialize `events`

Document how events should be represented when the workbook has:

* named events
* event budgets
* event-specific income/expense categories
* deposits or prepayments tied to events

### 2. Clarify when to materialize `documents`

Document how to represent:

* attached supporting documents
* note references
* receipt numbers
* external storage links
* spreadsheet notes or comments that deserve promotion

### 3. Add reference-resolution expectations

Document whether validators should require referenced IDs to resolve, and at what strictness level.

## Macro work

### 1. Build `events` when identifiable

If the workbook has stable event names or event budgets, synthesize or export top-level event records.

### 2. Build `documents` when identifiable

Where stable document identifiers exist, populate top-level document records and line-level or transaction-level `documentId`.

### 3. Improve person resolution

Refine top-level people building so that repeated payees/merchants/guardians resolve to the same person record where appropriate.

### 4. Add reference consistency checks

Check that exported:

* `personId`
* `budgetId`
* `eventId`
* `documentId`

all resolve correctly when top-level collections are populated.

## Deliverables

* richer top-level master-data collections
* stronger cross-reference integrity
* clearer distinction between free text and resolved IDs

---

# Phase 5 — Detail schedules and non-ledger registers

## Goal

Determine whether additional workbook detail tabs should be represented more directly in SCLX.

## Schema and documentation work

### 1. Review schedule/detail tabs case by case

Potential candidates include:

* liability detail sheets
* non-inventory asset detail sheets
* inventory registers
* cash advance schedules
* deposit/recoverable-amount schedules

### 2. Decide whether existing collections are sufficient

Prefer using existing SCLX collections where possible, such as:

* `otherAssetItems`
* `outstandingItems`
* `assets`
* `supplies`

Add new schema only if the existing model is truly insufficient.

### 3. Document schedule semantics

The manual should distinguish between:

* ledger postings
* status/register rows
* schedule-only records
* derived or report-only views

## Macro work

### 1. Add direct exporters for detail tabs

Where a tab represents a real register rather than a purely derived report, export it.

### 2. Add import logic only when safe

Import into schedule tabs should only be implemented when there is a stable, unambiguous mapping back into workbook rows.

### 3. Preserve provenance

All exported schedule records should include:

* `workbookLink`
* `extensions.workbook`
* any row-level provenance needed for round-tripping

## Deliverables

* a documented decision on each detail tab
* direct export/import where justified
* no silent loss of important schedule data

---

# Phase 6 — Validation, packaging, and release engineering

## Goal

Turn SCLX plus the workbook bridge into a more disciplined release package.

## Schema and documentation work

### 1. Add normative examples

For every major collection, include:

* minimal valid example
* realistic full example
* workbook-oriented example

### 2. Add conformance levels

Define levels such as:

* core SCLX producer
* workbook bridge producer
* full banking producer
* authoritative master-data producer

### 3. Add release notes per version

Each release should include:

* new fields
* validator changes
* compatibility notes
* migration notes

## Macro work

### 1. Package the VBA cleanly

Archive each released macro with:

* versioned `.bas`
* install README
* change log
* known limitations
* workbook mapping notes

### 2. Add regression tests

Maintain a set of reference workbooks and expected exported JSON files.

### 3. Add round-trip test cases

Test these scenarios:

* export only
* import only
* export then import to a clean workbook copy
* import 1.2 into 1.3-capable macro
* import 1.3 into future macro versions

### 4. Evaluate cross-platform support

Document the status of:

* Windows desktop Excel
* Mac desktop Excel
* Excel web limitations

## Deliverables

* repeatable release process
* versioned macro archive
* regression suite
* better long-term maintainability

---

## Potential future schema changes after 1.3

The following are candidates for later consideration, but should not be adopted without a concrete workbook or integration need.

### Candidate A — stronger workbookLink model

Possible future expansion:

* visible row number
* split index
* source range
* workbook or sheet identifier
* source cell addresses

Current preference:

* keep most of this in `extensions.workbook` unless interoperability demands more standardization

### Candidate B — richer people roles

Possible additions:

* role history
* source authority
* organizational role scope
* contact classification

Current preference:

* only add if authoritative people sheets are being exported consistently

### Candidate C — stronger banking linkage

Possible additions:

* explicit relationship keys among `transactions`, `bankingItems`, and `bankStatementImports`
* reconciliation status fields
* normalized imported bank-line identifiers

Current preference:

* first implement with current collections and extensions, then formalize if needed

### Candidate D — event budgeting model

Possible additions:

* event-specific budget objects
* event-to-budget linkage rules
* event fund segregation rules

Current preference:

* wait until actual workbook/event use cases justify it

### Candidate E — stricter referential validation

Possible later validator rules:

* every non-null `personId` must resolve
* every non-null `budgetId` must resolve
* every non-null `eventId` must resolve
* every non-null `documentId` must resolve

Current preference:

* add as optional strict mode first, not required core mode

---

## Potential future macro changes after the current reviewed build

The following are recommended macro enhancements beyond the current fixed 1.3 bridge.

### 1. Better normalization helpers

Add helpers for:

* currency normalization
* person name normalization
* merchant normalization
* date parsing and validation
* stable ID generation

### 2. Stronger duplicate handling

Prevent duplicate synthesized entries for:

* people
* budgets
* funds
* accounts

### 3. More robust import fallback logic

If workbook extension data is absent, import should still place rows sensibly using:

* canonical fields
* lookup-based mapping
* default/fallback sheet logic

### 4. Safer append logic

Improve append-row detection so template rows, formula rows, and placeholder rows are never mistaken for live data.

### 5. Structured error reporting

Instead of a single message box, optionally produce an error log sheet or a structured error summary.

### 6. Version-aware import

Import logic should branch by version:

* 1.2
* 1.3
* later versions as added

### 7. Optional strict mode

Allow the user to run in:

* permissive workbook-bridge mode
* strict schema-aligned mode

### 8. Better installability

Package the macro with:

* install instructions
* dependency notes
* version header
* compile-time checklist
* sample workbook/test JSON

---

## Priority order

### Highest priority

1. Freeze and document the current 1.3 package
2. Add macro layout validation
3. Add authoritative `Accounts` export
4. Add authoritative `Budget` export
5. Improve diagnostics and tests

### Medium priority

6. Populate `bankingItems`
7. Populate `bankStatementImports`
8. Improve people and reference resolution
9. Add events/documents where real identifiers exist

### Lower priority

10. Evaluate detail-tab expansion
11. Consider richer schema additions after real use confirms the need
12. Expand conformance levels and release engineering

---

## Acceptance criteria by milestone

### 1.3 stabilization milestone is complete when:

* schema, validator, and manual agree
* macro exports valid 1.3
* macro imports its own exports without major data loss on covered tabs
* workbook mappings are documented
* known limitations are documented

### Master-data milestone is complete when:

* accounts export from a real Accounts sheet when present
* budgets export from a real Budget sheet when present
* people export from authoritative role sheets when present
* referential IDs are more stable and less synthesized

### Banking milestone is complete when:

* bankingItems are materially populated
* bankStatementImports can be preserved
* outstanding and banking items can be reconciled more directly

---

## Archive maintenance recommendations

The SCLX archive should keep, at minimum:

* all released schema files
* all released validator files
* all released manuals
* all released overview/specification notes
* all released VBA modules
* installation README
* change log
* roadmap
* sample valid files
* regression test files

Suggested archive structure:

```text
/schema
/manual
/validator
/vba
/examples
/tests
/archive
ROADMAP.md
CHANGELOG.md
README.md
```

---

## Summary

The near-term goal is not to redesign SCLX. It is to make the existing workbook bridge stable, documented, and trustworthy.

After that, the main development path is:

1. replace synthesis with authoritative workbook master-data export
2. implement fuller banking support
3. strengthen references among transactions and top-level collections
4. add disciplined testing and release practices

That sequence gives the best balance of practical value, fidelity, and maintainability.
