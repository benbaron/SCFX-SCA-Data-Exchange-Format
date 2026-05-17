
#!/usr/bin/env python3
import json
import sys
from decimal import Decimal

def validate_balancing(txn):
    total_debit = Decimal("0")
    total_credit = Decimal("0")
    for line in txn.get("lines", []):
        debit = Decimal(line.get("debit") or "0")
        credit = Decimal(line.get("credit") or "0")
        if debit and credit:
            raise ValueError(f"Line {line.get('lineId')} has both debit and credit.")
        total_debit += debit
        total_credit += credit
    if total_debit != total_credit:
        raise ValueError(f"Transaction {txn.get('transactionId')} not balanced: {total_debit} != {total_credit}")

def validate_ledger(data):
    if data.get("format") != "SCLX":
        raise ValueError("Invalid format identifier")
    if not str(data.get("version")).startswith("1.2"):
        raise ValueError("Unsupported version")
    txns = data.get("transactions", [])
    for txn in txns:
        if len(txn.get("lines", [])) < 2:
            raise ValueError(f"Transaction {txn.get('transactionId')} has fewer than two lines")
        validate_balancing(txn)

def main():
    if len(sys.argv) < 2:
        print("Usage: sclx-validator.py file.json")
        sys.exit(1)
    with open(sys.argv[1]) as f:
        data = json.load(f)
    validate_ledger(data)
    print("SCLX validation passed.")

if __name__ == "__main__":
    main()
