@echo off

rem 32 bit
rem copy "H:\C#\SSIS_CustomComponentUppercase\SSIS_CustomComponentUppercase.dll" "C:\ProgramFiles(x86)%\Microsoft SQL Server\110\DTS\PipelineComponents" /Y

rem 64 bit
copy "H:\C#\SSIS_CustomComponentUppercase\SSIS_CustomComponentUppercase.dll" "C:\ProgramFiles%\Microsoft SQL Server\110\DTS\PipelineComponents" /Y
echo.

rem ****************************************************************************************************
rem Correct the paths to the gacutil utility to reflect its actual location in your environment:
rem ****************************************************************************************************

H:\C#\gacutil /u SSIS_CustomComponentUppercase
echo.

rem H:\C#\gacutil /i "C:\ProgramFiles(x86)%\Microsoft SQL Server\110\DTS\PipelineComponents\SSIS_CustomComponentUppercase.dll"
echo.

H:\C#\gacutil /i "C:\ProgramFiles%\Microsoft SQL Server\110\DTS\PipelineComponents\SSIS_CustomComponentUppercase.dll"
