@echo off
setlocal enableextensions enabledelayedexpansion
::Esperamos 300 segundos

FOR /L %%V IN (10, -1, 1) DO (
	echo Quedan %%V segundos para apagar el ordenador
	ping 127.0.0.1 -n 2 > nul
)

::Apagamos el pc
::powershell -c "$wshell = New-Object -ComObject wscript.shell; stop-computer -ComputerName localhost -Force
powershell -c "$wshell = New-Object -ComObject wscript.shell; stop-computer -ComputerName localhost