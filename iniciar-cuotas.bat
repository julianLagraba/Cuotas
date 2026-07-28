@echo off
cd /d "%~dp0"

if exist package.json (
    if not exist node_modules (
        echo Instalando dependencias...
        npm.cmd install
    )
)

echo Iniciando programa...
node server.js

pause