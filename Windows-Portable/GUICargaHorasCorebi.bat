@echo off
::-------------------------------------------------------------------------------------------------------------------
:: Lanzador de Python
::-------------------------------------------------------------------------------------------------------------------
:: Ejecuta el programa cargaHorasCorebi.py (en windows)
:: Creado en Nov de 2017
:: @kingSelta

:: Para que se ejecute minimizado
if not DEFINED IS_MINIMIZED set IS_MINIMIZED=1 && start "" /min "%~dpnx0" %* && exit

	echo Obteniendo el directorio actual
	SET workingDirectory=%~dp0
	echo %workingDirectory%

	echo ejecutando cargaHorasCorebi.py
	"%workingDirectory%Python\python.exe" "%workingDirectory%Lib\GUICargaHorasCorebi.py"


	echo Finalizando ejecucion automatica
	:: Si se desea cerrar la consola cmd comentar la siguiente linea
	::pause

:: Cerrar y finalizar ejecucion
exit