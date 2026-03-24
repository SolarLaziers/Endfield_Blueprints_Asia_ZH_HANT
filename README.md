# Endfield_Blueprints_Asia_ZH_HANT

## Valley Four Excel automation

Run the workbook automation in write-enabled mode to rebuild the managed output table and slicers:

```powershell
pwsh -File .\scripts\apply_sihao_valley_table.ps1 -EnableWorkbookWrites
```

To rebuild a temporary workbook copy instead of the repo workbook, pass `-WorkbookPath`:

```powershell
pwsh -File .\scripts\apply_sihao_valley_table.ps1 -EnableWorkbookWrites -WorkbookPath ".\temp\Endfield Blueprints (Asia).xlsx"
```

Verify the workbook state afterward:

```powershell
python .\scripts\verify_sihao_valley_table.py
```

Verification checks the source/output headers, source nonblank row count versus managed output row count, and basic text integrity for long `備註` values plus mixed-language `提供者` values.

Optionally pass a workbook path and a baseline workbook path:

```powershell
python .\scripts\verify_sihao_valley_table.py ".\Endfield Blueprints (Asia).xlsx" ".\baseline.xlsx"
```
