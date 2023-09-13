@echo off
set VB6EXE=C:\Program Files\Microsoft Visual Studio\VB98\VB6.exe
set MSBLDEXE=C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe
set BUILDLOG=build.log

echo VB6EXE=%VB6EXE%
echo MSBLDEXE=%MSBLDEXE%
echo BUILDLOG=%BUILDLOG%

REM �e�v���W�F�N�g���r���h
echo Start Build > %BUILDLOG%

IF EXIST "%VB6EXE%" (
  echo VB6 Build [D:\Zsrc_testA\DSizing\testA\testA.vbp] >> %BUILDLOG%
  "%VB6EXE%" /m "D:\Zsrc_testA\DSizing\testA\testA.vbp" /out %BUILDLOG%
)

IF EXIST "%MSBLDEXE%" (
  echo VB.NET Build [D:\Zsrc_testB\DSizing\testB\testB.vbproj] >> %BUILDLOG%
  "%MSBLDEXE%" "C:\Zsrc_testB\DSizing\testB\testB.vbproj" /t:clean;rebuild /p:Configuration=Release /fl
)

IF EXIST "%VB6EXE%" (
  echo VB6 Build [D:\Zsrc_�S�p�v���W�F�N�g\DSizing\testC\�S�p�v���W�F�N�g.vbp] >> %BUILDLOG%
  "%VB6EXE%" /m "D:\Zsrc_�S�p�v���W�F�N�g\DSizing\testC\�S�p�v���W�F�N�g.vbp" /out %BUILDLOG%
)


pause
