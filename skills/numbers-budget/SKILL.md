---
name: numbers-budget
description: Build a practical, formatted budget in Numbers with iwork-mcp tools.
---

# Build a Numbers budget

Clarify the budget period, currency, income categories, expense categories, and whether the user wants planned-versus-actual tracking. Prefer `numbers_create_sheet_with_table` for the initial model. Use formulas through `numbers_set_formula` for totals, variance, and savings rather than calculating static values.

Use clear column headers and separate income, fixed expenses, variable expenses, and summary rows. Apply currency and percent formats with `numbers_format_cells` or `numbers_format_range`. Use restrained colors, emphasizing totals and negative variances. Add a chart with `numbers_add_chart` only when it improves comparison or trend reading.

Verify values with `numbers_read_table`, formulas with `numbers_read_cell`, and formatting with `numbers_get_cell_format`. Save with `numbers_save_document` and export with `numbers_export_document` only when requested.
