@echo off
echo ==========================================
echo      Membangun PCM Summary Generator
echo ==========================================

:: 1. Bersihkan build lama (Optional, biar bersih)
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"

:: 2. Jalankan PyInstaller menggunakan file .spec
echo Sedang memproses...
pyinstaller pcm-summary-generator.spec --clean --noconfirm

:: 3. Konfirmasi selesai
if %ERRORLEVEL% EQU 0 (
    echo.
    echo [SUKSES] File .exe ada di folder 'dist'
    echo.
) else (
    echo.
    echo [GAGAL] Terjadi error saat build.
    echo.
)

pause