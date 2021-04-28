REM 空ディレクトリを探す

echo off
for /R %%d in ( . ) do call :sub "%%d"
exit /b

:sub
for /f "tokens=1-3" %%a in ('dir %1 ^| find "個のファイル"') do set fnum=%%a & set fsize=%%c
set fsize=%fsize:,=%
for /f "tokens=1" %%a in ('dir %1 ^| find "個のディレクトリ"') do set dir=%%a
if %fnum% EQU 0 if %fsize% EQU 0 if %dir% EQU 2 echo %1
goto :EOF