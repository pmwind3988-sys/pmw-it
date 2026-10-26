@echo off
rem Builds ..\PMW-Device-Scan.exe with the C# 5 compiler that ships with Windows (.NET Framework 4).
setlocal
set "CSC=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC%" set "CSC=%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
if not exist "%CSC%" (echo csc.exe not found under %SystemRoot%\Microsoft.NET & exit /b 1)
pushd "%~dp0"
"%CSC%" /nologo /noconfig /langversion:5 /target:winexe /platform:anycpu /optimize+ ^
  /win32manifest:app.manifest ^
  /r:System.dll /r:System.Core.dll /r:System.Management.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll ^
  /out:"%~dp0..\PMW-Device-Scan.exe" Program.cs MainForm.cs Controls.cs SettingsForm.cs ReportBuilder.cs ReportFiles.cs Settings.cs AssemblyInfo.cs
set "RC=%ERRORLEVEL%"
popd
if not "%RC%"=="0" (echo Build failed with code %RC%. & exit /b 1)
echo Built "%~dp0..\PMW-Device-Scan.exe"
endlocal
