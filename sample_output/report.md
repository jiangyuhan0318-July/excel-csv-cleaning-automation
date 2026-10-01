# Cleaning summary

input_rows: 20
cleaned_rows: 10
duplicates_removed: 2
rejected_rows: 8
normalized_cells: 45
accepted_missing_email: 2
amount_total: 3150.99
amount_mean: 315.10

Amounts are USD; statistics use accepted records only.

## Rejection reasons
- missing_customer: 1
- missing_amount: 1
- invalid_date: 1
- unknown_category: 1
- negative_amount: 1
- invalid_email: 1
- conflicting_order_id: 2

## Category counts
- Office: 4
- Technology: 3
- Furniture: 3

## Rejected source rows (header is row 1)
- Row 14, O011: missing_customer
- Row 15, O012: missing_amount
- Row 16, O013: invalid_date
- Row 17, O014: unknown_category
- Row 18, O015: negative_amount
- Row 19, O016: invalid_email
- Row 20, O017: conflicting_order_id
- Row 21, O017: conflicting_order_id

See before_after.csv for every row, normalized value and disposition. Optional missing emails remain blank. No missing values are invented.
