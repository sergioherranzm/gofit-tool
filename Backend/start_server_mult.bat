@echo off
setlocal enableextensions enabledelayedexpansion

::echo Esperamos 10 segundos
::ping 127.0.0.1 -n 10 > nul

FOR /L %%V IN (40, -1, 1) DO (
	echo Quedan %%V segundos para iniciar el script
	ping 127.0.0.1 -n 2 > nul
)

::Iniciamos los scripts
echo INICIANDO SCRIPT SERGIO
::start "C:\Users\BTC_server\AppData\Local\Programs\Python\Python39\python.exe" "C:\Users\BTC_server\Desktop\s_Folder\GoFit_Tool\Sergio\GoFit_Tool.py"

FOR /L %%V IN (400, -1, 1) DO (
	echo Quedan %%V segundos para iniciar el script de bea
	ping 127.0.0.1 -n 2 > nul
)

echo INICIANDO SCRIPT BEA
::start "C:\Users\BTC_server\AppData\Local\Programs\Python\Python39\python.exe" "C:\Users\BTC_server\Desktop\s_Folder\GoFit_Tool\Bea\GoFit_Tool.py"

exit