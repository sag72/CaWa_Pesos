@echo off
setlocal
cd /d "%~dp0"
title Pesos - Lojas nos XML CAWA

rem --- Encontrar o Python (python ou py launcher) ---
set "PY="
python --version >nul 2>&1 && set "PY=python"
if not defined PY py -3 --version >nul 2>&1 && set "PY=py -3"
if not defined PY goto sempython

rem --- Garantir que o openpyxl esta instalado ---
%PY% -c "import openpyxl" >nul 2>&1
if not errorlevel 1 goto menu
echo O modulo openpyxl nao esta instalado.
choice /c SN /n /m "Instalar agora (pip install --user openpyxl)? [S/N] "
if errorlevel 2 goto fim
%PY% -m pip install --user openpyxl
if errorlevel 1 (
    echo [ERRO] Nao foi possivel instalar o openpyxl.
    pause
    goto fim
)

:menu
cls
echo ===============================================
echo   Pesos - acrescentar lojas nos XML do CAWA
echo ===============================================
echo.
echo   1. Simular (dry-run, nao grava nada)
echo   2. Aplicar alteracoes nos XML
echo   3. Abrir versao web
echo   0. Sair
echo.
set "OP="
set /p "OP=Opcao: "
if "%OP%"=="1" goto simular
if "%OP%"=="2" goto aplicar
if "%OP%"=="3" goto web
if "%OP%"=="0" goto fim
goto menu

:simular
echo.
%PY% -W ignore adicionar_lojas_xml.py --dry-run
echo.
pause
goto menu

:aplicar
echo.
choice /c SN /n /m "Vai gravar os XML (com backup .bak). Continuar? [S/N] "
if errorlevel 2 goto menu
echo.
%PY% -W ignore adicionar_lojas_xml.py
echo.
pause
goto menu

:web
start "" "%~dp0pesos_web.html"
goto menu

:sempython
echo [ERRO] Python nao encontrado neste computador.
echo A abrir a versao web (nao precisa de Python)...
start "" "%~dp0pesos_web.html"
pause

:fim
endlocal
