# Contract Pricing Compliance & Data Quality Tracker

## Business problem

Private pricing contracts are manually billed according to negotiated discount terms. If the applied discount deviates from the contracted rate, or if invoices are issued against expired or non-existent contracts, it can lead to revenue leakage or customer overcharging. These discrepancies often go undetected without a systematic reconciliation of every invoice against the contract master.

![Dashboard](dashboard.png)

## What this workbook does

Runs seven automated validation checks across 1,264 invoices against a master of 34 private pricing contracts, classifies every exception by type and severity, quantifies the dollar exposure, and tracks each exception to resolution.

## Validation checks performed

1. Contract exists in master file
2. Applied discount matches contracted rate
3. Invoice date falls within contract term
4. Billed amount reconciles to list price and discount
5. Customer ID on invoice matches contract
6. Invoice ID is not duplicated
7. Required fields are populated

## Key findings

1. 216 of 1,264 invoices failed at least one check, leading to a 17.1% exception rate.
2. Gross financial exposure of $403,671, comprising $203,655 under-billed and $200,016 over-billed.
3. Net impact is only $3,639, which masks the true scale since the two error directions nearly cancel. Thus, a net-only view would have reported the process as materially clean.
4. Discount mismatches are the single largest category at 88 invoices, 41% of all exceptions.
5. 19 invoices reference contract IDs absent from the master file entirely, representing unverifiable revenue.
6. 14 invoices were billed twice, producing 28 duplicate line items in total.
7. Exception rates are broadly uniform across all six data-entry accounts (13.5%–19.3%), indicating a systemic control gap rather than an individual performance issue.

## Recommendations

1. Introduce a validation rule at point of entry that pulls the contracted discount rather than accepting a manually keyed value.
2. Block invoice submission against contract IDs not present in the master, and against contracts past their end date.
3. Report gross exposure rather than net in all compliance reporting.

## How to Refresh

Click the Run Audit button on the Dashboard tab. This reloads both source files, recalculates all validation checks, and timestamps the run (via VBA macro - Run audit). Alternatively use Data > Refresh All.

## Note on macros

This workbook contains one VBA macro (RunAudit) which refreshes the Power Query connections, recalculates the audit, and stamps the run time and number of records processed on the Dashboard.

## Techniques/Formulas used

Power Query, VLOOKUP, INDEX/MATCH, SUMIF, SUMIFS, COUNTIF, COUNTIFS, IFERROR, nested logical functions, PivotTables, slicers, conditional formatting, data validation, VBA.

## Data note

Source data is synthetic, generated for demonstration purposes. Contract terms, customer names, and billing records do not represent any real organisation.

## Formatting Note

Dates are formatted as DD-MM-YYYY, and amounts are formatted as $x,xx,xxx.xx instead of $xxx,xxx.xx due to the auto settings of my laptop.
