@echo off
setlocal
pushd %~dp0

REM === Validate input ===
if "%~1"=="" (
    echo Usage: build.bat ^<YourFile.cs^>
    exit /b 1
)

REM === Set vars ===
set "SOURCE=%~1"
set "BASENAME=%~n1"
set "OUTPUT=bin\%BASENAME%.exe"
set "DLL_NAME=bin\Newtonsoft.Json.dll"
set "CSC_PATH=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"

REM === Check if source file exists ===
if not exist "%SOURCE%" (
    echo ERROR: Source file "%SOURCE%" not found.
    exit /b 1
)

REM === Check if csc exists ===
if not exist "%CSC_PATH%" (
    echo ERROR: csc.exe not found at %CSC_PATH%
    exit /b 1
)

REM === Compile ===
echo Compiling %SOURCE% to %OUTPUT%...
"%CSC_PATH%" /nologo /platform:x64 /r:%DLL_NAME% /out:%OUTPUT% "%SOURCE%" Util.cs
if errorlevel 1 (
    echo Compilation failed.
    exit /b 1
)
echo Done: %OUTPUT%

popd
endlocal
