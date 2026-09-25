@echo off
setlocal
cd /d "%~dp0"
echo Iniciando automacao...
set "PYTHON=%LocalAppData%\Programs\Python\Python314\python.exe"
if exist "%PYTHON%" (
	"%PYTHON%" "%~dp0executar_automacao.py"
	set "codigo=%errorlevel%"
	goto finalizar
)
where py >nul 2>&1
if not errorlevel 1 (
	py -3 "%~dp0executar_automacao.py"
	set "codigo=%errorlevel%"
	goto finalizar
)
where python >nul 2>&1
if not errorlevel 1 (
	python "%~dp0executar_automacao.py"
	set "codigo=%errorlevel%"
	goto finalizar
)
echo Python nao foi encontrado neste computador.
set "codigo=1"
goto finalizar

:finalizar
echo.
if "%codigo%"=="0" (
	echo Automacao encerrada.
) else (
	echo A automacao terminou com erro. Codigo: %codigo%
)
pause
endlocal
exit /b %codigo%
