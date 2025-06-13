@echo off
setlocal
pushd %~dp0

REM === Validate input ===
if "%~1"=="" (
    echo Usage: build.bat ^<YourFile.cs^>  or  build.bat -a
    exit /b 1
)

REM === Set vars ===
set "DLL_NAME=bin\Newtonsoft.Json.dll"
set "CSC_PATH=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"

REM === Check if csc exists ===
if not exist "%CSC_PATH%" (
    echo ERROR: csc.exe not found at %CSC_PATH%
    exit /b 1
)

REM === Build all .cs files if -a flag is used ===
if /i "%~1"=="-a" (
    for %%F in (*.cs) do (
        if /i not "%%F"=="util.cs" (
            set "SOURCE=%%F"
            set "BASENAME=%%~nF"
            set "OUTPUT=bin\%%~nF.exe"
            call :buildone
        )
    )
    goto :eof
)

REM === Build single file ===
set "SOURCE=%~1"
set "BASENAME=%~n1"
set "OUTPUT=bin\%BASENAME%.exe"

call :buildone

popd
endlocal
goto :eof

:buildone
REM === Check if source file exists ===
if not exist "%SOURCE%" (
    echo ERROR: Source file "%SOURCE%" not found.
    exit /b 1
)

REM === Compile ===
echo Compiling %SOURCE% to %OUTPUT%...
"%CSC_PATH%" /nologo /platform:x64 /r:%DLL_NAME% /out:%OUTPUT% "%SOURCE%" util.cs
if errorlevel 1 (
    echo Compilation failed for %SOURCE%.
    exit /b 1
)
echo Done: %OUTPUT%
exit /b 0