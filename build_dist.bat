@echo off
setlocal ENABLEDELAYEDEXPANSION

echo ============================================
echo  BUILD - EQS Automate Conversor (Flet)
echo ============================================

cd /d "%~dp0"

echo Ativando venv...
call ".venv\Scripts\activate.bat"

echo Instalando dependencias...
pip install -r requirements.txt

set "APP_NAME=EQS Automate Conversor"
set "ICON_PATH=assets\icons\app_e2doc_icon.ico"

echo Gerando executavel...
flet pack main.py --name "%APP_NAME%" --product-name "%APP_NAME%" --icon "%ICON_PATH%"

echo -------------------------------------------
echo Build finalizado com sucesso!
echo Output gerado em: dist\
pause
endlocal
