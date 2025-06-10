@echo off
echo [BUILD] Transpiling modern JS to IE11-safe ES5 with Babel...

:: Transpile all JS files into a temp folder
call npx babel app/hta/js --out-dir transpiled --extensions ".js"

if errorlevel 1 (
    echo [ERROR] Babel failed
    pause
    exit /b 1
)

echo [BUILD] Bundling transpiled files with esbuild...

call npx esbuild transpiled/main.js --bundle --outfile=app/hta/js/bundle.js --target=es5 --loader:.json=json

if errorlevel 1 (
    echo [ERROR] esbuild failed
    pause
    exit /b 1
)

echo [CLEANUP] Removing transpiled folder...
rmdir /s /q transpiled

echo [SUCCESS] Build complete: app/hta/js/bundle.js is ready
pause
