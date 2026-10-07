@echo off
cd /d "%~dp0"
title DEKA BOM Karsilastirma

echo ====================================================
echo        DEKA ELEKTRONIK - BOM KARSILASTIRMA
echo ====================================================
echo.

if not exist ".venv\Scripts\python.exe" (
    echo [1/3] Sanal ortam olusturuluyor...
    python -m venv .venv 2>nul
    if not exist ".venv\Scripts\python.exe" (
        py -m venv .venv 2>nul
    )
    if not exist ".venv\Scripts\python.exe" (
        echo.
        echo [HATA] Python bulunamadi!
        echo Lutfen Python 3.10 veya uzerini yukleyin.
        echo Kurulum sirasinda "Add python.exe to PATH" secenegini isaretleyin.
        echo.
        pause
        exit /b 1
    )
    echo [2/3] Paketler yukleniyor, lutfen bekleyin...
    ".venv\Scripts\python.exe" -m pip install -e ".[dev]"
)

echo [3/3] Web sunucusu baslatiliyor...
echo.
echo Tarayicinizda aciliyor: http://127.0.0.1:8000
start http://127.0.0.1:8000
".venv\Scripts\python.exe" -m bom_enricher.web
pause
