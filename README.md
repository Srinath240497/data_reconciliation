# Diesta Implementation Engineer Reconciliation Exercise

## Scenario

You are working with an MGA that receives premium settlements from brokers.
You have been given:

- broker remittance files;
- bank transaction data showing cash received;
- premium BDX data exported from the MGA system.

The business wants to know how much can be reconciled, what remains unresolved, and why.

## Files

- `data/bank_transactions_may_2026.csv`
- `data/premium_bdx_april_2026.csv`
- `data/remittances/apex_risk_partners_remittance_apr_2026.csv`
- `data/remittances/pioneer_wholesale_statement_may_2026.csv`
- `data/remittances/northshore_broking_remittance_apr_2026.xlsx`
- `data/remittances/continental_coverholders_apr_2026.xlsx`
- `data/remittances/harbour_specialty_remittance_apr_2026.xlsx`
- `starter_reconciliation_analysis.ipynb`
- `requirements.txt`

## Suggested Setup

Use your normal Python environment. If needed:

```bash
python -m pip install -r requirements.txt
```

## End Goal

Produce a Jupyter notebook that:

1. Loads and profiles all supplied files.
2. Reconciles bank receipts to broker remittance statements.
3. Reconciles remittance lines to premium BDX records.
4. Classifies matched and unmatched items with reason codes.
5. Summarises financial variances and unresolved exceptions.
6. Explains assumptions, tolerances, and cases requiring manual review.

Do not force every item to match. A defensible exception is better than an unsafe match.
