@echo off
setlocal

echo ============================================
echo  EQS Automate Conversor - Build Portable
echo ============================================
echo.

REM Ir para a pasta deste script
cd /d "%~dp0"

REM
if exist ".venv\Scripts\activate.bat" (
    call ".venv\Scripts\activate.bat"
) else (
    echo [ERRO] Ambiente virtual .venv não encontrado.
    echo Crie com: python -m venv .venv
    pause
    exit /b 1
)

REM Instalar dependencias
echo.
echo Instalando dependencias do requirements.txt...
pip install -r requirements.txt

REM Definir nome e icone
set APP_NAME=EQS Automate Conversor
set ICON_PATH=assets\icons\app_e2doc_icon.png

REM Verificar se icone existe
if not exist "%ICON_PATH%" (
    echo [AVISO] Icone "%ICON_PATH%" nao encontrado.
    echo O build continuara, mas sem icone personalizado.
    set ICON_PARAM=
) else (
    set ICON_PARAM=--icon "%ICON_PATH%"
)

REM
echo.
echo Gerando executavel portable com flet pack...
flet pack main.py --name "%APP_NAME%" --product-name "%APP_NAME%" %ICON_PARAM% --output-folder dist

echo.
echo Build concluido!
echo Arquivos gerados em: dist\
echo.

pause
endlocal
