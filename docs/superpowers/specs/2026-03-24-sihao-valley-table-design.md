# 四號谷地_1 動態篩選表設計

## Goal

Transform `Endfield Blueprints (Asia).xlsx` sheet `四號谷地_1` into a separate, refreshable analysis table without modifying the source sheet.

## Confirmed Decisions

- Source workbook: `Endfield Blueprints (Asia).xlsx`
- Source sheet: `四號谷地_1`
- Source sheet remains untouched
- All source rows are preserved
- Output is a dynamic table on a new sheet
- Primary filters: `時代`, `類別`, `提供者`
- Keyword search targets: `藍圖名稱`, `備註`
- Traditional Chinese text must remain intact end-to-end

## Source Shape

The source sheet uses these active columns:

- `時代`
- `類別`
- `藍圖名稱`
- `藍圖代碼`
- `提供者`
- `備註`

Additional empty columns to the right are ignored.

The expected header row is row 1, and header matching should be exact against the six expected Traditional Chinese labels.

## Recommended Design

### Architecture

Use Excel desktop automation against the native workbook to create a Power Query that reads the used range from `四號谷地_1`, applies light normalization, and loads the result into a new worksheet as an Excel Table.

The base output worksheet name should be `四號谷地_1_表格`, subject to Excel sheet-name constraints and existing-name conflicts.

### Components

1. Source worksheet: `四號谷地_1`
2. Power Query query: reads and lightly normalizes the six source columns
3. Output worksheet: hosts the loaded table and slicers
4. Excel Table: refreshable structured table for filtering and downstream use
5. Slicers: attached to `時代`, `類別`, and `提供者`
6. Search helper column: `搜尋文字`, generated from `藍圖名稱` and `備註`

## Data Flow

1. Read data from `四號谷地_1`
2. Ignore fully blank non-header rows if any exist inside the used range, and exclude them from row-count validation
3. Retain the six business columns as text-oriented fields
4. Add a helper text column named `搜尋文字` by combining `藍圖名稱` and `備註`
5. Keep the six business columns in the same order as the source sheet, followed by `搜尋文字`
6. Lightly normalize text values:
   - trim leading and trailing whitespace where safe
   - preserve internal content, line breaks, and long notes
   - keep blanks as blanks
7. Load the result into a new worksheet as an Excel Table
8. Add slicers for `時代`, `類別`, and `提供者`
9. Use `搜尋文字` for practical keyword filtering across `藍圖名稱` and `備註`

## Traditional Chinese and Encoding Handling

The workflow must stay inside the native `.xlsx` environment to avoid encoding corruption.

- Do not export to CSV as an intermediate step
- Do not depend on terminal encoding for workbook content validation
- Keep query column names in Traditional Chinese
- Preserve mixed-language values in `提供者`
- Preserve long Traditional Chinese text in `備註`

Terminal output may not render Traditional Chinese reliably on this machine, so validation should prioritize Excel-side inspection and workbook object operations rather than console text dumps.

## Output Table Behavior

The output table should:

- include all rows from the source sheet
- expose native Excel filter dropdowns on all columns
- support slicer-based filtering on `時代`, `類別`, and `提供者`
- support keyword-style narrowing for `藍圖名稱` and `備註` through `搜尋文字`
- refresh when source data changes

## Search Design

Excel does not provide a true free-text slicer for arbitrary partial matching across two columns, so the search feature should be implemented through a query-generated helper column inside the refreshable output table.

This design intentionally avoids VBA or macros.

Preferred behavior:

- add a visible helper column named `搜尋文字` to the output table
- populate `搜尋文字` in Power Query by concatenating `藍圖名稱` and `備註`, treating blanks safely
- let the user apply Excel's normal text filter or filter search box on `搜尋文字` to find partial matches
- keep the search behavior case-insensitive where Excel text filtering allows
- treat an empty keyword filter as showing all rows

This keeps the experience practical without over-engineering the workbook.

## Error Handling and Safety

- If the target output sheet name already exists, create a deterministic safe variant by appending a numeric suffix rather than overwrite silently
- If the source sheet cannot be found, stop and report the exact missing sheet name
- If the six expected headers are not present, stop with a clear header-mismatch error instead of guessing a remap
- Do not mutate or reorder source-sheet cells
- Avoid destructive workbook operations

## Validation Plan

After implementation, verify:

1. The source sheet `四號谷地_1` is unchanged
2. Output row count matches the source data row count
3. Output headers display correctly in Traditional Chinese
4. Slicers work for `時代`, `類別`, and `提供者`
5. Keyword search can narrow rows using text from `藍圖名稱` or `備註` through `搜尋文字`
6. Long `備註` values and mixed-language provider names remain intact
7. Refresh updates the output when source rows are added or edited

## Out of Scope

- Excluding rows by fixed business rules
- Editing the source sheet in place
- Creating dashboards or PivotTables unless requested later
- Translating Traditional Chinese content

## Implementation Notes

The implementation should be carried out through Excel desktop automation so the workbook remains refreshable and Unicode-safe. Power Query should be used for the import/refresh layer, with Excel Table, slicers, and native header text filtering on `搜尋文字` providing the main interaction layer.
