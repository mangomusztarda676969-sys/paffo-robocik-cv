@echo off
cd /d "%~dp0"
echo Uruchamiam panel bota CV...
echo Panel otworzy sie automatycznie w przegladarce za chwile.
echo Nie zamykaj tego okna, dopoki chcesz korzystac z panelu.
echo.
if exist app.py (
    python app.py
) else (
    python app.pyc
)
pause
