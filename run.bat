@echo off
title File To Hex Text Files
color 06

echo ---Get Input File Path---
set /p "FpIn=Source File: "
echo FpIn=%FpIn%

echo ---Get Output Folder---
set /p "DirOut=Output Folder: "
echo DirOut=%DirOut%

echo ---Row Count---
set /p "RowCount=Row Count: "
echo RowCount=%RowCount%

echo ---Column Count---
set /p "ColumnCount=Column Count: "
echo ColumnCount=%ColumnCount%

echo ---Run---
cscript //nologo src\main.vbs %FpIn% %DirOut% %RowCount% %ColumnCount%
call :CheckError

:CheckError
echo.
if %errorlevel% NEQ 0 (
echo === WARNING ===
echo ERROR ErrorLevel=%errorlevel%
pause
exit /b 1
) else (
exit /b 0
)
