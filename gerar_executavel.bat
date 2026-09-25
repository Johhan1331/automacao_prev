@echo off
setlocal
cd /d "%~dp0"
python -m pip install -r requirements.txt
python -m pip install pyinstaller
pyinstaller --noconfirm --clean --onefile --windowed --name AutomacaoRelatorios executar_automacao.py
 echo.
echo Executavel criado em dist\AutomacaoRelatorios.exe
pause
endlocal
