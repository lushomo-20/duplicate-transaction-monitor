@echo off
setlocal

streamlit run "%~dp0dashboard_transaction_descriptions.py" --server.port 8502

pause
