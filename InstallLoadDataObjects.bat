@echo off
REM *************************************************************************
REM Script: InstallLoadDataObjects.bat
REM Author: Craig Nobili
REM 
REM Purpose: Install Load Data database objects which support the generic
REM          LoadData.cs program.
REM
REM *************************************************************************

REM Check if a servername was passed in
if %1nullparm == nullparm goto NEED_SERVER

SET servername=%1

REM Check if a database was passed in
if %2nullparm == nullparm goto NEED_DATABASE

SET database=%2

echo Installing Database Objects

sqlcmd -E -S %servername% -d %database% -i LoadDataDef.sql    -o LoadDataDef.log
sqlcmd -E -S %servername% -d %database% -i LoadDataOut.sql    -o LoadDataOut.log

sqlcmd -E -S %servername% -d %database% -i LogBegLoadData.sql -o LogBegLoadData.log
sqlcmd -E -S %servername% -d %database% -i LogEndLoadData.sql -o LogEndLoadData.log

goto END

:NEED_SERVER
echo Usage error, pass in the SERVERNAME as parameter
goto END

:NEED_DATABASE
echo Usage error, pass in the DATABASE as parameter
goto END

:END
