@echo off
setlocal
cd /d "%~dp0"
python src\devsecops_requirements_extractor.py
if errorlevel 1 (
  echo.
  echo Execution failed. Ensure Python is installed and dependencies are available.
  pause
)
