# ACFC Tally Bank Import

Bank-statement Excel to Tally Prime Accounting Voucher XML.

## Voucher logic

### Journal
- Journal: Journal Ledger = DR, Ledger Name = CR
- Auto-created Payment: Ledger Name = DR, Journal Payment Bank = CR

### Payment
- Ledger Name = DR
- Payment bank = CR

### Contra
- Contra Debit Bank = DR
- Contra Credit bank = CR
- Ledger Name is ignored

### Receipt
- Receipt Bank = DR
- Ledger Name = CR

## Input columns

Common: Date, Narration, Withdrawal Amt., Deposit Amt., ledger type, Ledger Name.

Type-specific:
- Journal: Journal Ledger, Journal Payment Bank
- Payment: Payment bank
- Contra: Contra Credit bank, Contra Debit Bank
- Receipt: Receipt Bank

The Master data is built into the application; no separate Master Excel is required.
