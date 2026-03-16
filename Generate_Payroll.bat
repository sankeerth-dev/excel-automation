@echo off
echo ==========================================
echo    PathAxiom Payroll Automation System
echo ==========================================
echo.
echo Running Payroll Generation...
call .venv\Scripts\python.exe payroll_automation.py
echo.
echo Payroll generated successfully! Check the newly created Monthly folder.
pause
