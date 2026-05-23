# Duplicate Transaction Monitor

Streamlit dashboard and analysis script for comparing `TransactionList.xlsb` and `SmartCITReport.xlsm`.

The scripts expect the source workbooks in:

- `C:\duplicate checker\TransactionList.xlsb`
- `C:\duplicate checker\SmartCITReport.xlsm`

The generated report is written to:

- `C:\duplicate checker\duplicate_report_june_2025_to_date.xlsx`

Actual duplicates are identified from the transaction workbook using account, Track/TRACE/VODS id, date, time, amount, and description. Possible duplicates are identified from the transaction workbook using account, date, time, amount, and description when the Track/VODS id differs. SmartCIT agent rankings use direct VODS matches when available; if there are no direct matches, the SmartCIT ranking uses possible duplicate groups with the same client, date, and amount.

## Run analysis

```powershell
python C:\Users\camilla\duplicate_checker\analyze_duplicates.py
```

## Run dashboard

```powershell
streamlit run C:\Users\camilla\duplicate_checker\dashboard.py
```

## Public repo safety

The repository is configured to exclude Excel workbooks, generated CSVs, and generated reports. Keep source data local.
