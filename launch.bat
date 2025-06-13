@echo off
pushd %~dp0
start "" /b pythonw ./app/web/main.py
popd