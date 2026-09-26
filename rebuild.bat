@echo off
REM ============================================================
REM  rebuild.bat - rebuild project VB RiplayWord2Html
REM  Klik 2x file ini. Menutup app yang masih jalan, lalu build.
REM ============================================================
setlocal
cd /d "%~dp0"

echo [1/3] Menutup RiplayWord2Html.exe yang masih berjalan (jika ada)...
taskkill /IM RiplayWord2Html.exe /F >nul 2>&1

echo [2/3] dotnet build...
dotnet build -nologo -v q
if errorlevel 1 (
    echo.
    echo *** BUILD GAGAL *** lihat error di atas.
    pause
    exit /b 1
)

echo [3/3] Build sukses.
echo EXE: %CD%\bin\Debug\net8.0-windows\RiplayWord2Html.exe
echo.
pause
