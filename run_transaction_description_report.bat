@echo off
setlocal

python "%~dp0analyze_transaction_descriptions.py"

echo.
echo Report generated at:
echo C:\duplicate checker\transaction_description_duplicate_report_june_2025_to_date.xlsx
pause
