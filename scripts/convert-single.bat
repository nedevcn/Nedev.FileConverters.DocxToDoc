@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

:: DocxToDoc Single File Converter Script
:: Converts a single DOCX file to DOC format

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

:: Check arguments
if "%~1"=="" goto :show_help
if /i "%~1"=="-h" goto :show_help
if /i "%~1"=="--help" goto :show_help

set "INPUT_FILE=%~1"
set "OUTPUT_FILE=%~2"
set "VERBOSE="

:: Check for verbose flag
if /i "%~2"=="-v" set "VERBOSE=-v" & set "OUTPUT_FILE="
if /i "%~2"=="--verbose" set "VERBOSE=-v" & set "OUTPUT_FILE="
if /i "%~3"=="-v" set "VERBOSE=-v"
if /i "%~3"=="--verbose" set "VERBOSE=-v"

:: Validate input file
if not exist "!INPUT_FILE!" (
    echo Error: Input file does not exist: !INPUT_FILE!
    exit /b 1
)

:: Check if it's a DOCX file
if /i not "%~x1"==".docx" (
    echo Warning: Input file does not have .docx extension
)

echo ============================================
echo DocxToDoc Single File Converter
echo ============================================
echo Input:  !INPUT_FILE!
if not "!OUTPUT_FILE!"=="" (
    echo Output: !OUTPUT_FILE!
)
echo ============================================
echo.

:: Run conversion
if "!OUTPUT_FILE!"=="" (
    dotnet "!CLI_DLL!" "!INPUT_FILE!" !VERBOSE!
) else (
    dotnet "!CLI_DLL!" "!INPUT_FILE!" "!OUTPUT_FILE!" !VERBOSE!
)

set "EXIT_CODE=!ERRORLEVEL!"

if !EXIT_CODE! == 0 (
    echo.
    echo Conversion completed successfully!
) else (
    echo.
    echo Conversion failed with error code: !EXIT_CODE!
)

exit /b !EXIT_CODE!

:show_help
echo Usage: convert-single.bat ^<input.docx^> [output.doc] [options]
echo.
echo Arguments:
echo   input.docx           Input DOCX file to convert
echo   output.doc           Optional output DOC file path
echo.
echo Options:
echo   -v, --verbose        Enable verbose output
echo   -h, --help           Show this help message
echo.
echo Examples:
echo   convert-single.bat C:\document.docx
echo   convert-single.bat C:\document.docx C:\output.doc
echo   convert-single.bat C:\document.docx -v
exit /b 0
