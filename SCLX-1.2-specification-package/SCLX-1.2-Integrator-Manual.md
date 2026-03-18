
# SCLX 1.2 Integrator Manual

SCLX (SCALedger Exchange Format) is a canonical JSON-based accounting ledger interchange format.
It represents full double-entry accounting data together with optional nonprofit operational records.

## Core Concepts

### Transaction Model
Transactions contain balanced posting lines.
Debits must equal credits.

### Chart of Accounts
Accounts define the ledger structure and may form hierarchies.

### Subsidiary Records
Optional collections may include:

- assets
- budgets
- supplies
- bankingItems
- outstandingItems

These records track operational activity separate from the accounting ledger.

### Banking Data
SCLX preserves bank reconciliation information including:
- cleared checks
- deposits
- OFX metadata
- statement imports

### Validation Rules

Implementations should validate:

1. Schema structure
2. Referential integrity
3. Balanced transactions

### File Naming

Typical filename:

ledger.sclx.json

### Encoding

UTF‑8 JSON

### Versioning

Version numbers follow:

major.minor

Example:

1.2

Backward-compatible additions increment the minor version.

### Extensions

Objects may include:

extensions

These contain vendor-specific metadata and must not alter core accounting meaning.

### Security

Treat SCLX files as untrusted input.
Always apply schema validation and size limits.

### Typical Workflow

Accounting system
→ Export SCLX
→ Audit or analytics tools
→ Archive storage
