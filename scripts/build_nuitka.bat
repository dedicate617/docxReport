@echo off
setlocal enabledelayedexpansion

echo ===================================================
echo [MindReport] Starting Nuitka Single-File Build (Windows x64)
echo ===================================================

:: 1. Force Temp and Cache to Z: drive to prevent C: disk full (C: space is tight)
set "REPO_ROOT=%~dp0.."
cd /d "%REPO_ROOT%"
set "TEMP=%REPO_ROOT%\.tmp"
set "TMP=%REPO_ROOT%\.tmp"
set "NUITKA_CACHE_DIR=%REPO_ROOT%\.tmp\nuitka_cache"

if not exist "%TEMP%" mkdir "%TEMP%"
if not exist "%NUITKA_CACHE_DIR%" mkdir "%NUITKA_CACHE_DIR%"
if not exist "dist" mkdir "dist"

:: 2. Initialize MSVC 64-bit compiler environment
set "VCVARS=C:\Program Files (x86)\Microsoft Visual Studio\2017\Community\VC\Auxiliary\Build\vcvars64.bat"
if exist "%VCVARS%" (
    echo [MindReport] Initializing MSVC x64 build tools...
    call "%VCVARS%"
) else (
    echo [MindReport Warning] vcvars64.bat not found at default location, relying on existing PATH.
)

:: 3. Run Nuitka compilation
echo [MindReport] Compiling MindReport with Nuitka into standalone onefile exe...
python -m nuitka ^
  --standalone ^
  --onefile ^
  --assume-yes-for-downloads ^
  --output-dir=dist ^
  --output-filename=MindReport.exe ^
  --windows-console-mode=force ^
  --include-package=src.MindReport ^
  --include-package=docx ^
  --include-package=docxtpl ^
  --include-package=jinja2 ^
  --include-package=openpyxl ^
  --include-package=mammoth ^
  --include-package-data=docx ^
  --include-package-data=docxtpl ^
  --include-package-data=mammoth ^
  --include-data-file=bin/win-x64/CrossReportRunner.exe=bin/win-x64/CrossReportRunner.exe ^
  MindReportCli.py

if %ERRORLEVEL% NEQ 0 (
    echo [MindReport Error] Nuitka compilation failed with code %ERRORLEVEL%!
    exit /b %ERRORLEVEL%
)

:: 4. Also copy CrossReportRunner.exe alongside dist\MindReport.exe for external flexibility
if exist "bin\win-x64\CrossReportRunner.exe" (
    echo [MindReport] Copying CrossReportRunner.exe alongside dist\MindReport.exe...
    copy /y "bin\win-x64\CrossReportRunner.exe" "dist\" >nul
)

echo ===================================================
echo [MindReport] Build complete!
echo Output executable: %REPO_ROOT%\dist\MindReport.exe
echo ===================================================
