@echo off
echo Installing...
python -m venv .env
.env\Scripts\pip install -r requirements.txt
echo Completed
pause
