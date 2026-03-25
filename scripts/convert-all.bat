@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

:: DocxToDoc Batch Converter Script
:: Converts all DOCX files in a directory to DOC format

set "SCRIPT_DIR=%~dp0"
set "PROJECT_DIR=%SCRIPT_DIR%.."
set "CLI_DLL=%PROJECT_DIR%\src\Nedev.FileConverters.DocxToDoc.Cli\bin\Release\net8.0\Nedev.FileConverters.DocxToDoc.Cli.dll"

:: Check if CLI is built
if not exist "%CLI_DLL%" (
    echo Building CLI project...
    dotnet build "%PROJECT_DIR%\Nedev.FileConverters.DocxToDoc.sln" -c Release
    if errorlevel 1 (
        echo Build failed!
        exit /b 1
    )
)

:: Parse arguments
set "INPUT_DIR="
set "OUTPUT_DIR="
set "RECURSIVE="
set "VERBOSE="

:parse_args
if "%~1"=="" goto :done_parse
if /i "%~1"=="-h" goto :show_help
if /i "%~1"=="--help" goto :show_help
if /i "%~1"=="-i" set "INPUT_DIR=%~2" & shift & shift & goto :parse_args
if /i "%~1"=="--input" set "INPUT_DIR=%~2" & shift & shift & goto :parse_args
if /i "%~1"=="-o" set "OUTPUT_DIR=%~2" & shift & shift & goto :parse_args
if /i "%~1"=="--output" set "OUTPUT_DIR=%~2" & shift & shift & goto :parse_args
if /i "%~1"=="-r" set "RECURSIVE=-r" & shift & goto :parse_args
if /i "%~1"=="--recursive" set "RECURSIVE=-r" & shift & goto :parse_args
if /i "%~1"=="-v" set "VERBOSE=-v" & shift & goto :parse_args
if /i "%~1"=="--verbose" set "VERBOSE=-v" & shift & goto :parse_args
shift
goto :parse_args

:done_parse

:: Validate input directory
if "!INPUT_DIR!"=="" (
    echo Error: Input directory not specified
    goto :show_help
)

if not exist "!INPUT_DIR!" (
    echo Error: Input directory does not exist: !INPUT_DIR!
    exit /b 1
)

:: Set default output directory if not specified
if "!OUTPUT_DIR!"=="" (
    set "OUTPUT_DIR=!INPUT_DIR!\converted"
)

:: Create output directory if it doesn't exist
if not exist "!OUTPUT_DIR!" (
    mkdir "!OUTPUT_DIR!"
)

echo ============================================
echo DocxToDoc Batch Converter
echo ============================================
echo Input:  !INPUT_DIR!
echo Output: !OUTPUT_DIR!
echo Recursive: !RECURSIVE!
echo Verbose: !VERBOSE!
echo ============================================
echo.

:: Run conversion
dotnet "!CLI_DLL!" -b "!INPUT_DIR!" -o "!OUTPUT_DIR!" !RECURSIVE! !VERBOSE!

set "EXIT_CODE=!ERRORLEVEL!"

if !EXIT_CODE! == 0 (
    echo.
    echo Conversion completed successfully!
) else if !EXIT_CODE! == 4 (
    echo.
    echo Conversion completed with some failures.
) else (
    echo.
    echo Conversion failed with error code: !EXIT_CODE!
)

exit /b !EXIT_CODE!

:show_help
echo Usage: convert-all.bat [options]
echo.
echo Options:
echo   -i, --input DIR      Input directory containing DOCX files
echo   -o, --output DIR     Output directory for converted DOC files
echo   -r, --recursive      Process subdirectories recursively
echo   -v, --verbose        Enable verbose output
echo   -h, --help           Show this help message
echo.
echo Examples:
echo   convert-all.bat -i C:\Documents -o C:\Converted
echo   convert-all.bat -i C:\Documents -r -v
echo   convert-all.bat --input C:\Docs --recursive --verbose
exit /b 0
