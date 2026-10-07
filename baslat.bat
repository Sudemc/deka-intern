@echo off
chcp 65001 >nul
title DEKA BOM Karşılaştırma

echo ====================================================
echo        DEKA ELEKTRONİK - BOM KARŞILAŞTIRMA
echo ====================================================
echo.

if not exist ".venv" (
    echo [1/3] Sanal ortam (.venv) oluşturuluyor...
    py -m venv .venv 2>nul || python -m venv .venv 2>nul
    if errorlevel 1 (
        echo [HATA] Python bulunamadı! Lütfen Python 3.10 veya üzerini yükleyin ve PATH'e ekleyin.
        pause
        exit /b 1
    )
    echo [2/3] Bağımlılıklar yükleniyor, lütfen bekleyin...
    call .venv\Scripts\activate.bat
    pip install -e ".[dev]" --quiet
) else (
    call .venv\Scripts\activate.bat
)

echo [3/3] Sunucu başlatılıyor...
echo.
echo Tarayıcınız açılıyor: http://127.0.0.1:8000
start http://127.0.0.1:8000
python -m bom_enricher.web
pause
