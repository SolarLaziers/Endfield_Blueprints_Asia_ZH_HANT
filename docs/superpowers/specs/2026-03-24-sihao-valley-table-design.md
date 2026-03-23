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

## Recommended Design

### Architecture

Use Power Query to read the used range from `四號谷地_1`, apply light normalization, and load the result into a new worksheet as an Excel Table.

The output worksheet should be named something close to `四號谷地_1_表格`, subject to Excel sheet-name constraints and existing-name conflicts.

### Components

1. Source worksheet: `四號谷地_1`
2. Power Query query: reads and lightly normalizes the six source columns
3. Output worksheet: hosts the loaded table and user-facing filter area
4. Excel Table: refreshable structured table for filtering and downstream use
5. Slicers: attached to `時代`, `類別`, and `提供者`
6. Search helper: supports partial text matching against `藍圖名稱` and `備註`

## Data Flow

1. Read data from `四號谷地_1`
2. Keep all non-header rows
3. Retain the six business columns as text-oriented fields
4. Lightly normalize text values:
   - trim leading and trailing whitespace where safe
   - preserve internal content, line breaks, and long notes
   - keep blanks as blanks
5. Load the result into a new worksheet as an Excel Table
6. Add slicers for `時代`, `類別`, and `提供者`
7. Add a practical keyword-search mechanism for `藍圖名稱` and `備註`

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
- support keyword-style narrowing for `藍圖名稱` and `備註`
- refresh when source data changes

## Search Design

Excel does not provide a true free-text slicer for arbitrary partial matching across two columns, so the search feature should be implemented as a helper-driven experience.

Preferred behavior:

- provide a visible search cell or small control area on the output sheet
- use a helper column or equivalent logic to evaluate whether `藍圖名稱` or `備註` contains the entered keyword
- allow the user to filter the helper result to matching rows

This keeps the experience practical without over-engineering the workbook.

## Error Handling and Safety

- If the target output sheet name already exists, create a safe variant rather than overwrite silently
- If the source sheet cannot be found, stop and report the exact missing sheet name
- If headers differ from the expected six-column shape, inspect and adapt before loading
- Do not mutate or reorder source-sheet cells
- Avoid destructive workbook operations

## Validation Plan

After implementation, verify:

1. The source sheet `四號谷地_1` is unchanged
2. Output row count matches the source data row count
3. Output headers display correctly in Traditional Chinese
4. Slicers work for `時代`, `類別`, and `提供者`
5. Keyword search can narrow rows using text from `藍圖名稱` or `備註`
6. Long `備註` values and mixed-language provider names remain intact
7. Refresh updates the output when source rows are added or edited

## Out of Scope

- Excluding rows by fixed business rules
- Editing the source sheet in place
- Creating dashboards or PivotTables unless requested later
- Translating Traditional Chinese content

## Implementation Notes

The preferred implementation toolchain is Excel-native automation so the workbook remains refreshable and Unicode-safe. Power Query should be used for the import/refresh layer, with Excel Table and slicers providing the main interaction layer.
