@echo off
setlocal
REM === Small EXE builds (ONEDIR), with and without heavy preview libs ===
where python >nul 2>&1 || (echo Python non trovato nel PATH & pause & exit /b 1)
where pyinstaller >nul 2>&1 || (echo PyInstaller non trovato. Installo... & python -m pip install --upgrade pip && python -m pip install pyinstaller) || (echo Errore nell'installazione di PyInstaller & pause & exit /b 1)

echo Installo dipendenze minime per il CORE (senza anteprime multipagina)...
python -m pip install pydicom pynetdicom requests || (echo Errore installazione dipendenze core & pause & exit /b 1)

echo.
echo [CORE - piccolo] Build ONEDIR senza librerie di rendering (fitz/pillow)
pyinstaller --clean --onedir --noconsole "%~dp0modality_pdf_uploader.py" ^
  --name ModalityPDF_Core ^
  --exclude-module fitz --exclude-module PIL --exclude-module PIL.Image || (echo Build CORE fallita & pause & exit /b 1)

echo.
echo (Opzionale) Installo librerie per ANTEPRIME (fitz/pillow) per la build FULL...
python -m pip install pillow pymupdf || (echo ATTENZIONE: impossibile installare pillow/pymupdf; salto build FULL)

IF %ERRORLEVEL% NEQ 0 goto :skip_full

echo.
echo [FULL - con anteprime] Build ONEDIR con librerie di rendering
pyinstaller --clean --onedir --noconsole "%~dp0modality_pdf_uploader.py" ^
  --name ModalityPDF_Full || (echo Build FULL fallita & goto :end)

:skip_full
echo.
echo (Se hai UPX nel PATH, PyInstaller comprime automaticamente alcuni binari)
echo.
echo Fatto. Trovi le cartelle in .\dist\ModalityPDF_Core\ e (se fatta) .\dist\ModalityPDF_Full\
:end
pause
