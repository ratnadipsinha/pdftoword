@echo off
echo ============================================
echo  PDF2Word Pro - Installing dependencies
echo ============================================
echo.

python -m pip install --upgrade pip
pip install -r requirements.txt

echo.
echo ============================================
echo  Installation complete!
echo  Run the app with:  python main.py
echo ============================================
pause
