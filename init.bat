@echo off
echo Installing...
python -m venv .env
.env\Scripts\pip install -U pip
.env\Scripts\pip install -r requirements.txt
echo Completed
pause
