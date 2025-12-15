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

echo [2/3] Menyalin User Manual...
if exist "USER_MANUAL.txt" (
    copy "USER_MANUAL.txt" "dist\USER_MANUAL.txt" >nul
    echo     - Manual berhasil disalin.
) else (
    echo     [WARNING] File User Manual.txt tidak ditemukan di folder project!
)


:: 4. Konfirmasi selesai
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