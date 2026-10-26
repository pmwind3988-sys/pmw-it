<# : PMW device scan -- double-click to run. No administrator rights needed.
@echo off
setlocal EnableExtensions DisableDelayedExpansion
title PMW Device Scan
rem This file is a batch file and a PowerShell script at once. cmd runs the
rem lines above the closing comment marker; PowerShell reads the whole file and
rem skips them. The same file saved as PMW-Device-Scan.ps1 is the PowerShell route.
rem
rem   1. PowerShell scan (Windows 7 and later, PowerShell 2 and later)
rem   2. If PowerShell is missing or fails, a basic scan with cmd and the
rem      registry only, written below as :basic. Same labels, fewer details.
rem
rem Reports go to the Reports folder beside this file. A machine that already
rem has a report there keeps its location and department and is not asked
rem again; a new one is asked for those two things and nothing else.
set "PMW_SCAN_SELF=%~f0"

rem PMW-Device-Scan.ini beside this file is written by the program's Settings
rem screen. The folder in it may be blank (Reports beside this file), a folder
rem on this drive kept without its letter (\IT\Scans), or a network folder.
set "PMW_FOLDER="
set "PMW_AUTOCLOSE=yes"
if exist "%~dp0PMW-Device-Scan.ini" for /f "usebackq eol=; tokens=1,* delims==" %%a in ("%~dp0PMW-Device-Scan.ini") do (
    if /i "%%a"=="ReportsFolder" set "PMW_FOLDER=%%b"
    if /i "%%a"=="AutoClose" set "PMW_AUTOCLOSE=%%b"
)
set "PMW_BESIDE=%~dp0Reports"
set "PMW_REPORTS=%PMW_BESIDE%"
if not defined PMW_FOLDER goto folderSet
set "PMW_REPORTS=%~dp0%PMW_FOLDER%"
if "%PMW_FOLDER:~1,1%"==":" set "PMW_REPORTS=%PMW_FOLDER%"
if "%PMW_FOLDER:~0,1%"=="\" set "PMW_REPORTS=%~d0%PMW_FOLDER%"
if "%PMW_FOLDER:~0,2%"=="\\" set "PMW_REPORTS=%PMW_FOLDER%"
:folderSet
if not exist "%PMW_REPORTS%" mkdir "%PMW_REPORTS%" >nul 2>&1
rem A network folder that cannot be reached: save beside this file instead.
if not exist "%PMW_REPORTS%" (
    echo   %PMW_REPORTS% cannot be reached. Saving beside this file instead.
    set "PMW_REPORTS=%PMW_BESIDE%"
)
if not exist "%PMW_BESIDE%" mkdir "%PMW_BESIDE%" >nul 2>&1
set "PMW_OWNER="
set "PMW_AV="
set "PMW_REMARKS="
set "PMW_PLACE="
set "PMW_KNOWN="
for %%d in ("%PMW_REPORTS%" "%PMW_BESIDE%") do for %%f in ("%%~d\[*] %COMPUTERNAME%_.txt") do (
    set "PMW_KNOWN=1"
    for /f "tokens=1 delims=[]" %%p in ("%%~nf") do set "PMW_PLACE=%%p"
)
echo.
echo   PMW device scan -- %COMPUTERNAME%
echo.
if defined PMW_KNOWN (
    echo   Scanned before: %PMW_PLACE%. Scanning again.
    goto asked
)
echo   New machine. Places already on this drive:
set "PMW_LIST="
for %%d in ("%PMW_REPORTS%" "%PMW_BESIDE%") do for %%f in ("%%~d\[*]*.txt") do for /f "tokens=1 delims=[]" %%p in ("%%~nf") do call :listPlace "%%p"
if not defined PMW_LIST echo     (none yet)
echo.
set "PMW_LOC="
set "PMW_DEPT="
set /p "PMW_LOC=  Location, e.g. F1, F3, PML: "
set /p "PMW_DEPT=  Department, e.g. ENGINEERING: "
set "PMW_PLACE=%PMW_LOC% %PMW_DEPT%"
:asked
set "PMW_ASKED=1"
echo.

set "PMW_PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PMW_PS%" goto basic
rem The plain way first: run the .ps1 beside this file. Antivirus treats a
rem script that builds itself from text (the line after) as suspicious, so
rem that is kept only for when the .ps1 has gone missing from the drive.
if not exist "%~dp0PMW-Device-Scan.ps1" goto selfScript
"%PMW_PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0PMW-Device-Scan.ps1"
if errorlevel 1 goto basic
goto finish
:selfScript
"%PMW_PS%" -NoProfile -ExecutionPolicy Bypass -Command "& ([scriptblock]::Create([IO.File]::ReadAllText($env:PMW_SCAN_SELF)))"
if errorlevel 1 goto basic
goto finish

:basic
echo.
echo   PowerShell did not finish on this machine.
echo   Writing a basic report with cmd and the registry instead...
call :basicScan

:finish
echo.
if /i "%PMW_AUTOCLOSE%"=="no" (
    echo   Completed.
    pause
    exit /b
)
echo   Completed. This window closes in 10 seconds.
timeout /t 10 >nul 2>&1 || ping -n 11 127.0.0.1 >nul
exit /b

rem :listPlace <place>   prints each place once
:listPlace
echo ;%PMW_LIST%; | find /i ";%~1;" >nul && exit /b
set "PMW_LIST=%PMW_LIST%;%~1"
echo     %~1
exit /b

rem ===========================================================================
rem  Basic scan. Writes the same labels as the PowerShell scan so the portal
rem  reads it the same way; whatever cmd cannot see is left blank. wmic is used
rem  where Windows still has it and the registry where it does not.
rem ===========================================================================
:basicScan
setlocal EnableDelayedExpansion
set "PLACE=!PMW_PLACE!"
if defined PLACE set "PLACE=!PLACE:[=!"
if defined PLACE set "PLACE=!PLACE:]=!"
if defined PLACE set "PLACE=!PLACE:\=!"
if defined PLACE set "PLACE=!PLACE:/=!"
if defined PLACE set "PLACE=!PLACE::=!"
if defined PLACE for /f "tokens=*" %%p in ("!PLACE!") do set "PLACE=%%p"
if defined PLACE if "!PLACE:~-1!"==" " set "PLACE=!PLACE:~0,-1!"
if defined PLACE (set "NAME=[!PLACE!] %COMPUTERNAME%_.txt") else set "NAME=%COMPUTERNAME%_.txt"
set "F=%PMW_REPORTS%\!NAME!"
(echo.) > "!F!" 2>nul || set "F=%USERPROFILE%\Desktop\!NAME!"
set "T=%TEMP%\pmw-scan-%RANDOM%"
mkdir "!T!" >nul 2>&1
set "HAS_WMIC="
wmic os get Caption >nul 2>&1 && set "HAS_WMIC=1"

echo   Listing installed software...
call :swCollect

> "!F!" (
    call :b_head
    for %%s in (model board winver cpu gpu ram ramslots disks net av monitor creds folders office adobe mail serial maker product family boardserial uuid bios chassis modules diskserials monitors battery build winlic officelic officever accounts subscription join software info) do call :b_%%s
    echo ===============================
    echo END OF REPORT
    echo ===============================
)
rmdir /s /q "!T!" >nul 2>&1
echo.
echo   Report saved to: !F!
endlocal
exit /b

rem --- helpers ---------------------------------------------------------------

rem :reg <key> <value> <var>   reads one registry value, blank when absent
:reg
set "%~3="
for /f "tokens=2,*" %%a in ('reg query "%~1" /v "%~2" 2^>nul ^| findstr /c:"REG_"') do set "%~3=%%b"
exit /b

rem :wmi <wmic class or alias> <property> <var>   first value, blank without wmic
:wmi
set "%~3="
if not defined HAS_WMIC exit /b
for /f "tokens=1,* delims==" %%a in ('wmic %~1 get %~2 /value 2^>nul ^| findstr /c:"="') do (
    for /f "delims=" %%c in ("%%b") do if not defined %~3 set "%~3=%%c"
)
exit /b

rem :gb <bytes> <var>   bytes to whole GB, inside cmd's 32-bit arithmetic
:gb
set "%~2="
set "_b=%~1"
if not defined _b exit /b
set "_m=!_b:~0,-6!"
if not defined _m set "_m=0"
set /a "%~2=(_m+537)/1074"
exit /b

rem :block <label>   label, then the caller echoes values, then :end adds a blank
:label
echo(%~1:
exit /b

:value
if "%~1"=="" exit /b
echo(%~1
exit /b

rem --- the original report -----------------------------------------------------

:b_head
echo(Name: !PMW_OWNER!
echo Anydesk:
set "AD="
for %%c in ("%ProgramData%\AnyDesk\system.conf" "%APPDATA%\AnyDesk\system.conf") do (
    if not defined AD if exist "%%~c" for /f "tokens=1,* delims==" %%a in ('findstr /c:"ad.anynet.id" "%%~c"') do set "AD=%%b"
)
if defined AD (echo(!AD!) else echo None
echo(Antivirus status: !PMW_AV!
echo(Remarks: !PMW_REMARKS!
echo.
echo.
echo ===============================
echo COMPUTER INFORMATION
echo ===============================
echo.
echo Computer Name:
echo(%COMPUTERNAME%
echo.
exit /b

:b_model
echo Computer Model:
call :wmi computersystem Model V
if not defined V call :reg "HKLM\HARDWARE\DESCRIPTION\System\BIOS" SystemProductName V
if defined V echo(!V!
echo.
exit /b

:b_board
echo Motherboard:
call :wmi baseboard Manufacturer M
call :wmi baseboard Product P
if not defined M call :reg "HKLM\HARDWARE\DESCRIPTION\System\BIOS" BaseBoardManufacturer M
if not defined P call :reg "HKLM\HARDWARE\DESCRIPTION\System\BIOS" BaseBoardProduct P
echo(!M! ^| !P!
echo.
exit /b

:b_winver
echo Windows Version:
call :wmi os Caption V
if not defined V call :reg "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" ProductName V
if defined V echo(!V!
echo.
exit /b

:b_cpu
echo Processor:
call :reg "HKLM\HARDWARE\DESCRIPTION\System\CentralProcessor\0" ProcessorNameString V
if defined V for /f "tokens=*" %%a in ("!V!") do echo(%%a
echo.
exit /b

:b_gpu
echo GPU:
if defined HAS_WMIC (
    for /f "tokens=1,* delims==" %%a in ('wmic path Win32_VideoController get Name /value 2^>nul ^| findstr /c:"="') do for /f "delims=" %%c in ("%%b") do echo(%%c
) else (
    for /f "tokens=2,*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}" /s /v DriverDesc 2^>nul ^| findstr /c:"REG_SZ"') do echo(%%b
)
echo.
exit /b

:b_ram
echo Total RAM:
call :wmi computersystem TotalPhysicalMemory V
call :gb "!V!" G
if not defined G (
    rem No wmic: systeminfo says "Total Physical Memory: 7,839 MB".
    for /f "tokens=1,* delims=:" %%a in ('systeminfo 2^>nul') do if /i "%%a"=="Total Physical Memory" (
        set "MB=%%b"
        set "MB=!MB: =!"
        set "MB=!MB:,=!"
        set "MB=!MB:.=!"
        set "MB=!MB:MB=!"
        set /a "G=(MB+512)/1024"
    )
)
if defined G echo(!G! GB
echo.
exit /b

rem Memory: one query feeds both RAM Slot Info and Memory Modules.
:b_ramslots
echo RAM Slot Info:
if defined HAS_WMIC (
    set "N=0"
    for /f "tokens=1,* delims==" %%a in ('wmic memorychip get Capacity^,DeviceLocator^,Manufacturer^,MemoryType^,PartNumber^,SMBIOSMemoryType^,SerialNumber^,Speed /value 2^>nul ^| findstr /c:"="') do (
        set "v=" & for /f "delims=" %%c in ("%%b") do set "v=%%c"
        if /i "%%a"=="Capacity" set "cap=!v!"
        if /i "%%a"=="DeviceLocator" set "loc=!v!"
        if /i "%%a"=="Manufacturer" set "mk=!v!"
        if /i "%%a"=="MemoryType" set "mt=!v!"
        if /i "%%a"=="PartNumber" set "pn=!v!"
        if /i "%%a"=="SMBIOSMemoryType" set "st=!v!"
        if /i "%%a"=="SerialNumber" set "sn=!v!"
        if /i "%%a"=="Speed" (
            set "sp=!v!"
            call :ramLine
        )
    )
    for /f "tokens=1,* delims==" %%a in ('wmic memphysical get MemoryDevices /value 2^>nul ^| findstr /c:"="') do for /f "delims=" %%c in ("%%b") do set "SLOTS=%%c"
    echo(Total Slots: !SLOTS! ^| Used Slots: !N!
)
echo.
exit /b

:ramLine
set /a N+=1
call :gb "!cap!" g
set "ty=!st!"
if not defined ty set "ty=!mt!"
if "!ty!"=="0" set "ty=!mt!"
set "tn=Unknown"
if "!ty!"=="20" set "tn=DDR"
if "!ty!"=="21" set "tn=DDR2"
if "!ty!"=="24" set "tn=DDR3"
if "!ty!"=="26" set "tn=DDR4"
if "!ty!"=="34" set "tn=DDR5"
for /f "tokens=*" %%x in ("!mk! ") do set "mk=%%x"
for /f "tokens=*" %%x in ("!pn! ") do set "pn=%%x"
echo(!g! GB ^| !tn! ^| !sp! MHz ^| !mk! ^| !pn!
if not defined sn set "sn=None"
>> "!T!\modules.txt" echo(!loc! ^| !g! GB ^| !mk! ^| P/N !pn! ^| S/N !sn!
set "cap=" & set "loc=" & set "mk=" & set "mt=" & set "pn=" & set "st=" & set "sn="
exit /b

rem Drives: one query feeds Storage Drives and Drive Serials.
:b_disks
echo Storage Drives:
if defined HAS_WMIC for /f "tokens=1,* delims==" %%a in ('wmic diskdrive get FirmwareRevision^,InterfaceType^,Model^,SerialNumber^,Size /value 2^>nul ^| findstr /c:"="') do (
    set "v=" & for /f "delims=" %%c in ("%%b") do set "v=%%c"
    if /i "%%a"=="FirmwareRevision" set "fw=!v!"
    if /i "%%a"=="InterfaceType" set "bus=!v!"
    if /i "%%a"=="Model" set "md=!v!"
    if /i "%%a"=="SerialNumber" set "sn=!v!"
    if /i "%%a"=="Size" (
        call :gb "!v!" g
        echo(!md! ^| Unspecified ^| !g! GB
        for /f "tokens=*" %%x in ("!sn! ") do set "sn=%%x"
        if "!sn!"=="" set "sn=None"
        >> "!T!\disks.txt" echo(!md! ^| S/N !sn! ^| !bus! ^| Firmware !fw!
        set "fw=" & set "bus=" & set "md=" & set "sn="
    )
)
echo.
exit /b

:b_net
echo Network Information:
set "SSID="
for /f "tokens=1,2,*" %%a in ('netsh wlan show interfaces 2^>nul') do if /i "%%a"=="SSID" set "SSID=%%c"
if defined HAS_WMIC for /f "tokens=1,* delims==" %%a in ('wmic nicconfig get DHCPEnabled^,Description^,IPAddress /value 2^>nul ^| findstr /c:"="') do (
    set "v=" & for /f "delims=" %%c in ("%%b") do set "v=%%c"
    if /i "%%a"=="DHCPEnabled" set "dh=!v!"
    if /i "%%a"=="Description" set "ds=!v!"
    if /i "%%a"=="IPAddress" if defined v (
        set ip=!v:"=!
        for /f "tokens=1 delims={}," %%i in ("!ip!") do set "ip=%%i"
        if /i "!dh!"=="TRUE" (set "as=Dynamic") else set "as=Static"
        set "wifi="
        for %%w in (Wireless Wi-Fi WiFi WLAN 802.11) do if not "!ds:%%w=!"=="!ds!" set "wifi=1"
        if defined wifi (echo(Wi-Fi ^| SSID: !SSID! ^| IP: !ip! ^| !as!) else echo(LAN ^| Network Name: ^| IP: !ip! ^| !as!
    )
)
if not defined HAS_WMIC (
    rem No wmic: read ipconfig. An adapter's name is the unindented line.
    set "ad=" & set "dh="
    for /f "tokens=1,* delims=:" %%a in ('ipconfig /all') do (
        set "k=%%a"
        if not "!k:~0,1!"==" " set "ad=%%a" & set "dh="
        if not "!k:DHCP Enabled=!"=="!k!" for /f "tokens=*" %%v in ("%%b") do set "dh=%%v"
        if not "!k:IPv4 Address=!"=="!k!" for /f "tokens=1 delims=( " %%i in ("%%b") do (
            if /i "!dh!"=="Yes" (set "as=Dynamic") else set "as=Static"
            set "wifi="
            for %%w in (Wireless Wi-Fi WiFi WLAN) do if not "!ad:%%w=!"=="!ad!" set "wifi=1"
            if defined wifi (echo(Wi-Fi ^| SSID: !SSID! ^| IP: %%i ^| !as!) else echo(LAN ^| Network Name: ^| IP: %%i ^| !as!
        )
    )
)
echo.
exit /b

:b_av
echo Antivirus:
if defined HAS_WMIC for /f "tokens=1,* delims==" %%a in ('wmic /namespace:\\root\SecurityCenter2 path AntiVirusProduct get displayName^,productState /value 2^>nul ^| findstr /c:"="') do (
    set "v=" & for /f "delims=" %%c in ("%%b") do set "v=%%c"
    if /i "%%a"=="displayName" set "nm=!v!"
    if /i "%%a"=="productState" (
        set /a "on=!v! & 4096"
        if "!on!"=="0" (echo(!nm! ^| Disabled) else echo(!nm! ^| Enabled
    )
)
echo.
exit /b

:b_monitor
echo Monitor:
if defined HAS_WMIC for /f "tokens=1,* delims==" %%a in ('wmic desktopmonitor get Name /value 2^>nul ^| findstr /c:"="') do for /f "delims=" %%c in ("%%b") do echo(%%c
echo.
exit /b

:b_creds
echo PMW Server and credentials:
set "found="
set "tg="
for /f "tokens=1,*" %%a in ('cmdkey /list 2^>nul') do (
    if /i "%%a"=="Target:" (
        set "tg=%%b"
        set "u=!tg:*target=!"
        if not "!u!"=="!tg!" set "tg=!u:~1!"
    )
    if /i "%%a"=="User:" (
        if /i "!tg!"=="server" (echo(!tg! ^| %%b& set "found=1")
        if "!tg!"=="192.168.1.250" (echo(!tg! ^| %%b& set "found=1")
    )
)
if not defined found echo None
echo.
exit /b

:b_folders
echo Server folder:
for /f "delims=" %%k in ('reg query "HKCU\Network" 2^>nul ^| findstr /b "HKEY_"') do (
    call :reg "%%k" RemotePath RP
    echo(%%~nxk: ^| !RP!
)
echo.
exit /b

:b_office
echo Microsoft Office:
call :reg "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" ProductReleaseIds C2R
if defined C2R echo(!C2R!
echo.
exit /b

:b_adobe
echo Adobe:
set "found="
for /f "tokens=1,2 delims=|" %%a in ('findstr /i /c:"Adobe Acrobat" "!T!\software.txt" 2^>nul') do (echo(%%a^|%%b& set "found=1")
if not defined found echo None
echo.
exit /b

:b_mail
echo Email data files found Active or Inactive account:
set "found="
for %%d in ("%LOCALAPPDATA%\Microsoft\Outlook" "%USERPROFILE%\Documents\Outlook Files") do (
    for /f "delims=" %%f in ('dir /s /b /a-d "%%~d\*.ost" "%%~d\*.pst" 2^>nul') do (echo(%%~nxf ^| %%f& set "found=1")
)
if not defined found echo None
echo.
exit /b

rem --- new in version 2 --------------------------------------------------------

:b_serial
echo Serial Number:
call :wmi bios SerialNumber V
if defined V (echo(!V!) else echo None
echo.
exit /b

:b_maker
echo Manufacturer:
call :reg "HKLM\HARDWARE\DESCRIPTION\System\BIOS" SystemManufacturer V
if not defined V call :wmi computersystem Manufacturer V
if defined V echo(!V!
echo.
exit /b

:b_product
echo Product Number:
call :reg "HKLM\HARDWARE\DESCRIPTION\System\BIOS" SystemSKU V
if defined V (echo(!V!) else echo None
echo.
exit /b

:b_family
echo Model Family:
call :reg "HKLM\HARDWARE\DESCRIPTION\System\BIOS" SystemVersion V
call :reg "HKLM\HARDWARE\DESCRIPTION\System\BIOS" SystemProductName P
if defined V if not "!V!"=="!P!" echo(!V!
echo.
exit /b

:b_boardserial
echo Motherboard Serial:
call :wmi baseboard SerialNumber V
if defined V (echo(!V!) else echo None
echo.
exit /b

:b_uuid
echo Device UUID:
call :wmi csproduct UUID V
if defined V echo(!V!
echo.
exit /b

:b_bios
echo BIOS:
call :reg "HKLM\HARDWARE\DESCRIPTION\System\BIOS" BIOSVendor A
call :reg "HKLM\HARDWARE\DESCRIPTION\System\BIOS" BIOSVersion B
call :reg "HKLM\HARDWARE\DESCRIPTION\System\BIOS" BIOSReleaseDate C
rem The registry writes MM/DD/YYYY; the report writes YYYY-MM-DD.
if defined C if "!C:~2,1!!C:~5,1!"=="//" set "C=!C:~6,4!-!C:~0,2!-!C:~3,2!"
if defined A echo(!A! ^| !B! ^| !C!
echo.
exit /b

:b_chassis
echo Chassis:
call :wmi systemenclosure ChassisTypes V
call :wmi systemenclosure SMBIOSAssetTag TAG
if defined V (
    for /f "tokens=1 delims={}," %%t in ("!V!") do set "V=%%t"
    set "K=Unknown"
    for %%p in ("3=Desktop" "4=Low profile desktop" "6=Mini tower" "7=Tower" "8=Portable" "9=Laptop" "10=Notebook" "13=All in one" "14=Sub notebook" "30=Tablet" "31=Convertible" "32=Detachable" "35=Mini PC") do (
        for /f "tokens=1,* delims==" %%x in (%%p) do if "%%x"=="!V!" set "K=%%y"
    )
    if not defined TAG set "TAG=None"
    echo(!K! ^| Asset tag: !TAG!
)
echo.
exit /b

:b_modules
echo Memory Modules:
if exist "!T!\modules.txt" type "!T!\modules.txt"
echo.
exit /b

:b_diskserials
echo Drive Serials:
if exist "!T!\disks.txt" type "!T!\disks.txt"
echo.
exit /b

:b_monitors
echo Monitor Details:
echo.
exit /b

:b_battery
echo Battery:
if defined HAS_WMIC for /f "tokens=1,* delims==" %%a in ('wmic path Win32_Battery get DeviceID^,EstimatedChargeRemaining^,Name /value 2^>nul ^| findstr /c:"="') do (
    set "v=" & for /f "delims=" %%c in ("%%b") do set "v=%%c"
    if /i "%%a"=="DeviceID" set "di=!v!"
    if /i "%%a"=="EstimatedChargeRemaining" set "ch=!v!"
    if /i "%%a"=="Name" echo(!v! ^| !di! ^| Charge !ch!%%
)
echo.
exit /b

:b_build
echo Windows Build:
set "K=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
call :reg "%K%" DisplayVersion DV
if not defined DV call :reg "%K%" ReleaseId DV
if not defined DV call :reg "%K%" CSDVersion DV
call :reg "%K%" CurrentBuildNumber CB
call :reg "%K%" UBR UBR
if defined UBR (set /a "UBR=!UBR!" & set "CB=!CB!.!UBR!")
set "ARCH=32-bit"
if /i "%PROCESSOR_ARCHITECTURE%"=="AMD64" set "ARCH=64-bit"
if /i "%PROCESSOR_ARCHITEW6432%"=="AMD64" set "ARCH=64-bit"
if /i "%PROCESSOR_ARCHITECTURE%"=="ARM64" set "ARCH=ARM 64-bit"
echo(!DV! ^| Build !CB! ^| !ARCH! ^| Installed
echo.
exit /b

rem slmgr.vbs is Windows' own licence tool and answers a standard user.
:b_winlic
echo Windows License:
set "nm=" & set "ds=" & set "ky=" & set "stt="
for /f "tokens=1,* delims=:" %%a in ('cscript //nologo "%SystemRoot%\System32\slmgr.vbs" /dli 2^>nul') do (
    for /f "tokens=*" %%v in ("%%b") do (
        if /i "%%a"=="Name" set "nm=%%v"
        if /i "%%a"=="Description" set "ds=%%v"
        if /i "%%a"=="Partial Product Key" set "ky=%%v"
        if /i "%%a"=="License Status" set "stt=%%v"
    )
)
if defined nm (
    set "ch=!ds:*, =!"
    set "ch=!ch: channel=!"
    echo(!ch! ^| !stt! ^| Key ends !ky! ^| !nm!
) else echo None found
call :wmi "path SoftwareLicensingService" OA3xOriginalProductKeyDescription OEM
if defined OEM (echo(Firmware key ^| !OEM!) else echo Firmware key ^| None
echo.
exit /b

:b_officelic
echo Office License:
set "found="
if defined C2R for %%i in (!C2R!) do (
    set "found=1"
    set "id=%%i"
    set "kind=Unknown"
    if /i "!id:~-6!"=="Retail" set "kind=Perpetual, retail"
    if /i "!id:~-6!"=="Volume" set "kind=Perpetual, volume"
    if /i "!id:~0,4!"=="O365" set "kind=Subscription (Microsoft 365)"
    if /i "!id:~0,4!"=="M365" set "kind=Subscription (Microsoft 365)"
    echo(%%i ^| !kind! ^| Click-to-Run
)
for %%d in ("%ProgramFiles%\Microsoft Office" "%ProgramFiles(x86)%\Microsoft Office") do for %%o in (Office16 Office15 Office14) do (
    if exist "%%~d\%%o\OSPP.VBS" call :ospp "%%~d\%%o\OSPP.VBS"
)
for /f "tokens=1,2" %%a in ('reg query "HKCU\Software\Microsoft\Office\16.0\Common\Licensing\LicensingNext" 2^>nul ^| findstr /c:"REG_DWORD"') do (
    if /i not "%%a"=="InstalledGraceKey" if /i not "%%a"=="MigrationToV5Done" (echo(Signed-in licence ^| %%a& set "found=1")
)
if not defined found echo None found
echo.
exit /b

:ospp
set "nm=" & set "ds=" & set "stt="
for /f "tokens=1,* delims=:" %%a in ('cscript //nologo "%~1" /dstatus 2^>nul') do (
    for /f "tokens=*" %%v in ("%%b") do (
        if /i "%%a"=="LICENSE NAME" set "nm=%%v"
        if /i "%%a"=="LICENSE DESCRIPTION" set "ds=%%v"
        if /i "%%a"=="LICENSE STATUS" set "stt=%%v"
        if /i "%%a"=="Last 5 characters of installed product key" (
            set "ch=!ds:*, =!"
            set "ch=!ch: channel=!"
            set "stt=!stt:-=!"
            echo(!nm! ^| !ch! ^| !stt! ^| Key ends %%v
            set "found=1"
        )
    )
)
exit /b

:b_officever
echo Office Version:
set "K=HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
call :reg "%K%" VersionToReport VR
call :reg "%K%" Platform PF
call :reg "%K%" UpdateChannel U1
call :reg "%K%" CDNBaseUrl U2
if defined VR (
    set "url=!U1! !U2!"
    set "chn=Unknown channel"
    for %%p in ("492350f6=Current Channel" "64256afe=Current Channel (Preview)" "55336b82=Monthly Enterprise Channel" "7ffbc6bf=Semi-Annual Enterprise Channel" "b8f9b850=Semi-Annual Enterprise Channel (Preview)" "5440fd1f=Beta Channel" "f2e724c1=LTSC 2019" "5030841d=LTSC 2021" "7983bac0=LTSC 2024") do (
        for /f "tokens=1,* delims==" %%x in (%%p) do if not "!url:%%x=!"=="!url!" set "chn=%%y"
    )
    echo(!VR! ^| !PF! ^| !chn!
) else echo None
echo.
exit /b

:b_accounts
echo Office Accounts:
set "found="
set "K=HKCU\Software\Microsoft\Office\16.0\Common\Identity"
for /f "delims=" %%k in ('reg query "%K%\Identities" 2^>nul ^| findstr /b "HKEY_"') do (
    call :reg "%%k" EmailAddress EM
    call :reg "%%k" ProviderId PI
    if defined EM (
        set "kind=Unknown"
        if not "!PI:AD=!"=="!PI!" set "kind=Work or school"
        if not "!PI:OrgId=!"=="!PI!" set "kind=Work or school"
        if not "!PI:LiveId=!"=="!PI!" set "kind=Personal Microsoft account"
        echo(!EM! ^| !kind!
        set "found=1"
    )
)
set /a "work=0, home=0"
for /f "delims=" %%k in ('reg query "%K%\Profiles" 2^>nul ^| findstr /b "HKEY_"') do (
    set "pk=%%~nxk"
    if /i "!pk:~-5!"=="_ADAL" set /a work+=1
    if /i "!pk:~-7!"=="_LiveId" set /a home+=1
)
if !work! gtr 0 (echo(Work or school sign-in ^| !work! profile^(s^)& set "found=1")
if !home! gtr 0 (echo(Personal Microsoft account sign-in ^| !home! profile^(s^)& set "found=1")
if not defined found echo None
echo.
exit /b

rem cmd cannot unpack Office's licence files, only the accounts they belong to.
:b_subscription
echo Office Subscription:
set "found="
for /f "tokens=1,2,*" %%a in ('reg query "HKCU\Software\Microsoft\Office\16.0\Common\Licensing\LicensingNext\LicenseIdToEmailMapping" 2^>nul ^| findstr /c:"REG_SZ"') do (
    echo(Unknown ^| Unknown ^| %%c ^| Tenant Unknown ^| Valid until Unknown ^| Unknown
    set "found=1"
)
if not defined found echo None
echo.
exit /b

:b_join
echo Device Join:
set "aj=" & set "dj=" & set "wj=" & set "tn="
for /f "tokens=1,2,*" %%a in ('dsregcmd /status 2^>nul') do if "%%b"==":" (
    if /i "%%a"=="AzureAdJoined" set "aj=%%c"
    if /i "%%a"=="DomainJoined" set "dj=%%c"
    if /i "%%a"=="WorkplaceJoined" set "wj=%%c"
    if /i "%%a"=="TenantName" set "tn=%%c"
)
if defined aj (
    echo(Azure AD joined ^| !aj!
    echo(Domain joined ^| !dj!
    echo(Workplace joined ^| !wj!
    if defined tn echo(Tenant ^| !tn!
)
echo.
exit /b

:b_software
echo Installed Software:
if exist "!T!\software.txt" type "!T!\software.txt"
echo.
exit /b

:b_info
echo Scan Info:
set "adm=No"
net session >nul 2>&1 && set "adm=Yes"
echo(Version 2 basic ^| %DATE% %TIME:~0,5% ^| Run by %USERDOMAIN%\%USERNAME% ^| Admin !adm!
echo.
exit /b

rem --- installed software, from the same uninstall keys PowerShell reads -------
:swCollect
for %%r in ("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall") do (
    for /f "delims=" %%k in ('reg query "%%~r" 2^>nul ^| findstr /b "HKEY_"') do call :swEntry "%%k"
)
if exist "!T!\raw.txt" (
    sort "!T!\raw.txt" /o "!T!\sorted.txt"
    set "prev="
    for /f "usebackq delims=" %%l in ("!T!\sorted.txt") do (
        if not "%%l"=="!prev!" >> "!T!\software.txt" echo(%%l
        set "prev=%%l"
    )
)
exit /b

:swEntry
set "n=" & set "v=" & set "p=" & set "d=" & set "sys=" & set "parent=" & set "rt="
for /f "tokens=1,2,*" %%a in ('reg query "%~1" 2^>nul') do (
    if "%%a"=="DisplayName" set "n=%%c"
    if "%%a"=="DisplayVersion" set "v=%%c"
    if "%%a"=="Publisher" set "p=%%c"
    if "%%a"=="InstallDate" set "d=%%c"
    if "%%a"=="SystemComponent" set "sys=%%c"
    if "%%a"=="ParentKeyName" set "parent=%%c"
    if "%%a"=="ReleaseType" set "rt=%%c"
)
if not defined n exit /b
if "!sys!"=="0x1" exit /b
if defined parent exit /b
if defined rt (
    if not "!rt:Update=!"=="!rt!" exit /b
    if not "!rt:Hotfix=!"=="!rt!" exit /b
    if not "!rt:ServicePack=!"=="!rt!" exit /b
)
if not "!n:(KB=!"=="!n!" exit /b
if /i "!n:~0,10!"=="Update for" exit /b
if /i "!n:~0,15!"=="Security Update" exit /b
if /i "!n:~0,6!"=="Hotfix" exit /b
set "n=!n:|=/!"
set "n=!n::= -!"
if defined p set "p=!p:|=/!"
set "dd="
if defined d if "!d:~8!"=="" if not "!d:~7!"=="" set "dd=!d:~0,4!-!d:~4,2!-!d:~6,2!"
>> "!T!\raw.txt" echo(!n! ^| !v! ^| !p! ^| !dd!
exit /b
#>

# ============================================================================
#  PMW device scan, version 2 (2026-10-26)
#
#  Writes the same report as "computer information - PMW.bat" -- every label,
#  in the same order, so the portal reads old and new reports alike -- and then
#  adds what the device list now keeps: licences, serial and part numbers, and
#  the installed software.
#
#  Runs on Windows PowerShell 2.0 and later (Windows 7 onwards). Everything
#  here is readable by a standard user. Nothing is changed on the machine,
#  nothing is sent anywhere; the report is one .txt file.
#
#  PMW-Device-Scan.exe writes the same report. A label added here goes there
#  too, and into KNOWN_LABELS in src/features/devices/parse/labels.js.
# ============================================================================

$ErrorActionPreference = 'SilentlyContinue'
$ScriptVersion = '2'
$out = New-Object System.Collections.Generic.List[string]
$HasCim = [bool](Get-Command Get-CimInstance -ErrorAction SilentlyContinue)

function Say($text) { Write-Host "  $text" -ForegroundColor Cyan }

# Get-CimInstance from PowerShell 3, Get-WmiObject on PowerShell 2 (Windows 7).
function Wmi([string]$class, [string]$namespace = 'root\cimv2', [string]$filter = '') {
    $args2 = @{ Namespace = $namespace; ErrorAction = 'SilentlyContinue' }
    if ($filter) { $args2.Filter = $filter }
    if ($HasCim) { Get-CimInstance -ClassName $class @args2 } else { Get-WmiObject -Class $class @args2 }
}

# One block of the report: the label, its values, a blank line. A block that
# fails writes its label with no value rather than stopping the scan.
function Block([string]$label, [scriptblock]$body) {
    $out.Add("$label`:")
    try {
        foreach ($line in @(& $body)) {
            if ($null -ne $line -and "$line".Trim() -ne '') { $out.Add("$line".Trim()) }
        }
    } catch {}
    $out.Add('')
}

# BIOS filler that is no value at all. The portal applies the same rule, this
# only decides when to try the next place a serial can live.
function CleanValue($value) {
    if ($null -eq $value) { return $null }
    $v = "$value".Trim().TrimEnd('.')
    $junk = 'To Be Filled By O.E.M.', 'To be filled by O.E.M', 'Default string', 'Not Specified',
            'System Serial Number', 'System Product Name', 'System Version', 'Unknown', 'None',
            'N/A', 'NA', '0', '00000000', 'INVALID', 'Not Available', 'Not Supported',
            'OEM_Define0', 'OEM_Define1', 'OEM_Define2', 'Chassis Serial Number', 'Base Board Serial Number'
    if ($v -eq '' -or $junk -contains $v -or $v -match '^(.)\1*$') { return $null }
    return $v
}

function T($value) {
    if ($null -eq $value) { return '' }
    return "$value".Trim()
}

# CIM hands back a date; WMI on PowerShell 2 hands back 20240320000000.000000+000.
function Day($value) {
    if (-not $value) { return '' }
    try {
        if ("$value" -match '^\d{14}\.') { $value = [Management.ManagementDateTimeConverter]::ToDateTime("$value") }
        return ([datetime]$value).ToString('yyyy-MM-dd')
    } catch { return '' }
}

# Uninstall entries store InstallDate as yyyyMMdd text.
function RegDay($value) {
    if ("$value" -match '^(\d{4})(\d{2})(\d{2})$') { return "$($Matches[1])-$($Matches[2])-$($Matches[3])" }
    return ''
}

# Reports go to the Reports folder beside the scan file. The .bat says where
# that is; run as a .ps1 the script knows its own path.
$reports = $env:PMW_REPORTS
if (-not $reports -and $MyInvocation.MyCommand.Path) {
    # Run as the .ps1: read the program's settings file the way the .bat does.
    $here = Split-Path $MyInvocation.MyCommand.Path
    $beside = Join-Path $here 'Reports'
    $reports = $beside
    $ini = Join-Path $here 'PMW-Device-Scan.ini'
    $folder = ''
    if (Test-Path $ini) {
        foreach ($line in Get-Content $ini) {
            if ($line -match '^\s*ReportsFolder\s*=(.*)$') { $folder = $Matches[1].Trim() }
        }
    }
    if ($folder -like '\\*') { $reports = $folder }
    elseif ($folder -like '\*') { $reports = (Split-Path $here -Qualifier) + $folder }
    elseif ($folder -match '^[A-Za-z]:') { $reports = $folder }
    elseif ($folder) { $reports = Join-Path $here $folder }
    if (-not (Test-Path $reports)) { New-Item -ItemType Directory -Path $reports | Out-Null }
    if (-not (Test-Path $reports)) { Say "$reports cannot be reached. Saving beside this file instead."; $reports = $beside }
}
if (-not $reports) { $reports = Join-Path ([Environment]::GetFolderPath('Desktop')) 'PMW Reports' }
if (-not (Test-Path $reports)) { New-Item -ItemType Directory -Path $reports | Out-Null }

# Name, antivirus status and remarks are typed into the report by hand later,
# as before; the scan itself only ever asks where a NEW machine is.
$ownerName = "$env:PMW_OWNER"
$avAnswer  = "$env:PMW_AV"
$remarks   = "$env:PMW_REMARKS"
if ($env:PMW_ASKED) {
    $placeAnswer = "$env:PMW_PLACE"
} else {
    $previous = Get-ChildItem -LiteralPath $reports -Filter '*.txt' |
        Where-Object { $_.Name -like "* $($env:COMPUTERNAME)_.txt" -and $_.Name -match '^\[(.+?)\]' } | Select-Object -First 1
    Write-Host ''
    Say "PMW device scan -- $($env:COMPUTERNAME)"
    if ($previous) {
        $placeAnswer = ([regex]::Match($previous.Name, '^\[(.+?)\]')).Groups[1].Value
        Say "Scanned before: $placeAnswer. Scanning again."
    } else {
        $places = Get-ChildItem -LiteralPath $reports -Filter '*.txt' |
            ForEach-Object { ([regex]::Match($_.Name, '^\[(.+?)\]')).Groups[1].Value } | Where-Object { $_ } | Sort-Object -Unique
        Say 'New machine. Places already on this drive:'
        if ($places) { $places | ForEach-Object { Write-Host "    $_" } } else { Write-Host '    (none yet)' }
        $location   = Read-Host '  Location, e.g. F1, F3, PML'
        $department = Read-Host '  Department, e.g. ENGINEERING'
        $placeAnswer = "$location $department"
    }
}
Write-Host ''
Say 'Scanning, this takes about a minute...'

# The licensing service is the slowest thing asked, so it starts first and is
# collected at the end. One query covers Windows and Office (Office 2013 and later).
$licenceQuery = {
    $q = @{ Class = 'SoftwareLicensingProduct'; Filter = 'PartialProductKey IS NOT NULL' }
    $items = if (Get-Command Get-CimInstance -ErrorAction SilentlyContinue) {
        Get-CimInstance -ClassName $q.Class -Filter $q.Filter -ErrorAction SilentlyContinue
    } else { Get-WmiObject -Class $q.Class -Filter $q.Filter -ErrorAction SilentlyContinue }
    $items | Select-Object Name, Description, ApplicationID, LicenseStatus, PartialProductKey, ProductKeyChannel
}
$licenceJob = $null
try { $licenceJob = Start-Job -ScriptBlock $licenceQuery -ErrorAction Stop } catch {}

$cs      = Wmi Win32_ComputerSystem
$bios    = Wmi Win32_BIOS
$cspr    = Wmi Win32_ComputerSystemProduct
$board   = Wmi Win32_BaseBoard | Select-Object -First 1
$os      = Wmi Win32_OperatingSystem
$biosReg = Get-ItemProperty 'HKLM:\HARDWARE\DESCRIPTION\System\BIOS'
$ntReg   = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
$c2r     = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'

# ---------------------------------------------------------------------------
#  The original report, unchanged
# ---------------------------------------------------------------------------

$out.Add("Name: $ownerName".TrimEnd())
$out.Add('Anydesk:')
$anydesk = $null
foreach ($conf in "$env:ProgramData\AnyDesk\system.conf", "$env:APPDATA\AnyDesk\system.conf") {
    if (-not $anydesk -and (Test-Path $conf)) {
        $line = (Select-String -Path $conf -Pattern 'ad.anynet.id' | Select-Object -First 1).Line
        if ($line) { $anydesk = ($line -replace '.*=').Trim() }
    }
}
$out.Add($(if ($anydesk) { $anydesk } else { 'None' }))
$out.Add("Antivirus status: $avAnswer".TrimEnd())
$out.Add("Remarks: $remarks".TrimEnd())
$out.Add('')
$out.Add('')
$out.Add('===============================')
$out.Add('COMPUTER INFORMATION')
$out.Add('===============================')
$out.Add('')

Block 'Computer Name'  { $env:COMPUTERNAME }
Block 'Computer Model' { $cs.Model }
Block 'Motherboard'    { "$($board.Manufacturer) | $($board.Product)" }
Block 'Windows Version' { $os.Caption }
Block 'Processor'      { Wmi Win32_Processor | ForEach-Object { $_.Name } }
Block 'GPU'            { Wmi Win32_VideoController | ForEach-Object { $_.Name } }
Block 'Total RAM'      { "$([math]::Round($cs.TotalPhysicalMemory / 1GB, 0)) GB" }

# SMBIOSMemoryType arrived with Windows 10; Windows 7 only fills MemoryType.
$ramModules = @(Wmi Win32_PhysicalMemory)
Block 'RAM Slot Info' {
    foreach ($m in $ramModules) {
        $code = if ($m.SMBIOSMemoryType) { $m.SMBIOSMemoryType } else { $m.MemoryType }
        $type = switch ($code) { 20 {'DDR'} 21 {'DDR2'} 24 {'DDR3'} 26 {'DDR4'} 34 {'DDR5'} default {'Unknown'} }
        "$([math]::Round($m.Capacity / 1GB, 0)) GB | $type | $($m.Speed) MHz | $(T $m.Manufacturer) | $(T $m.PartNumber)"
    }
    "Total Slots: $((Wmi Win32_PhysicalMemoryArray | Measure-Object MemoryDevices -Sum).Sum) | Used Slots: $($ramModules.Count)"
}

# Get-PhysicalDisk arrived with Windows 8. Before that the drive list comes
# from Win32_DiskDrive, which cannot tell an SSD from a spinning disk.
$disks = @()
if (Get-Command Get-PhysicalDisk -ErrorAction SilentlyContinue) {
    $disks = @(Get-PhysicalDisk | ForEach-Object {
        New-Object PSObject -Property @{ Name = $_.FriendlyName; Media = "$($_.MediaType)"; Size = $_.Size
                                         Serial = $_.SerialNumber; Bus = "$($_.BusType)"; Firmware = $_.FirmwareVersion }
    })
}
if (-not $disks.Count) {
    $disks = @(Wmi Win32_DiskDrive | ForEach-Object {
        New-Object PSObject -Property @{ Name = $_.Model; Media = 'Unspecified'; Size = $_.Size
                                         Serial = $_.SerialNumber; Bus = $_.InterfaceType; Firmware = $_.FirmwareRevision }
    })
}
Block 'Storage Drives' {
    foreach ($d in $disks) { "$($d.Name) | $($d.Media) | $([math]::Round($d.Size / 1GB, 0)) GB" }
}

# Get-NetConnectionProfile arrived with Windows 8 too; before it, the adapter
# settings and netsh give the same three facts.
Block 'Network Information' {
    if (Get-Command Get-NetConnectionProfile -ErrorAction SilentlyContinue) {
        foreach ($p in Get-NetConnectionProfile | Where-Object { $_.IPv4Connectivity -ne 'Disconnected' }) {
            $alias = $p.InterfaceAlias
            $ip = (Get-NetIPAddress -InterfaceAlias $alias -AddressFamily IPv4 | Where-Object { $_.IPAddress } |
                   Select-Object -ExpandProperty IPAddress) -join ', '
            $dhcp = if ("$((Get-NetIPInterface -InterfaceAlias $alias -AddressFamily IPv4).Dhcp)" -eq 'Enabled') { 'Dynamic' } else { 'Static' }
            if ($alias -match 'Wi[- ]?Fi') { "Wi-Fi | SSID: $($p.Name) | IP: $ip | $dhcp" }
            else { "LAN | Network Name: $($p.Name) | IP: $ip | $dhcp" }
        }
    } else {
        $ssid = ''
        foreach ($line in netsh wlan show interfaces) { if ($line -match '^\s*SSID\s*:\s*(.+)$') { $ssid = $Matches[1].Trim() } }
        foreach ($nic in Wmi Win32_NetworkAdapterConfiguration -filter 'IPEnabled = TRUE') {
            $ip = ($nic.IPAddress | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' }) -join ', '
            $dhcp = if ($nic.DHCPEnabled) { 'Dynamic' } else { 'Static' }
            if ($nic.Description -match 'Wireless|Wi-?Fi|WLAN|802\.11') { "Wi-Fi | SSID: $ssid | IP: $ip | $dhcp" }
            else { "LAN | Network Name: | IP: $ip | $dhcp" }
        }
    }
}

Block 'Antivirus' {
    Wmi AntivirusProduct 'root\SecurityCenter2' | ForEach-Object {
        "$($_.displayName) | $(if (($_.productState -band 0x1000) -ne 0) { 'Enabled' } else { 'Disabled' })"
    }
}

Block 'Monitor' { Wmi Win32_DesktopMonitor | ForEach-Object { $_.Name } }

Block 'PMW Server and credentials' {
    $servers = '192.168.1.250', 'server'
    $found = $false
    $target = $null
    foreach ($line in cmdkey /list) {
        if ($line -match '^\s*Target:\s*(.+)$') { $target = $Matches[1] -replace '.*target=' }
        elseif ($line -match '^\s*User:\s*(.+)$' -and $servers -contains $target) {
            $found = $true
            "$target | $($Matches[1])"
        }
    }
    if (-not $found) { 'None' }
}

Block 'Server folder' {
    Get-ChildItem 'HKCU:\Network' | ForEach-Object { "$($_.PSChildName): | $((Get-ItemProperty $_.PSPath).RemotePath)" }
}

Block 'Microsoft Office' { if ($c2r.ProductReleaseIds) { $c2r.ProductReleaseIds } }

$uninstallKeys = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
                 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
                 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
$uninstall = @(Get-ItemProperty $uninstallKeys | Where-Object { $_.DisplayName })

Block 'Adobe' {
    $apps = $uninstall | Where-Object { $_.DisplayName -match 'Adobe Acrobat|Adobe Acrobat Reader' }
    if ($apps) { $apps | ForEach-Object { "$($_.DisplayName) | $($_.DisplayVersion)" } } else { 'None' }
}

Block 'Email data files found Active or Inactive account' {
    $files = Get-ChildItem (Join-Path $env:LOCALAPPDATA 'Microsoft\Outlook'), (Join-Path $env:USERPROFILE 'Documents\Outlook Files') `
        -Include *.ost, *.pst -Recurse | Where-Object { -not $_.PSIsContainer }
    if ($files) { $files | ForEach-Object { "$($_.Name) | $($_.FullName)" } } else { 'None' }
}

# ---------------------------------------------------------------------------
#  New in version 2: identity, part numbers, serials
# ---------------------------------------------------------------------------

Say 'Reading serial and part numbers...'

Block 'Serial Number' {
    $s = CleanValue $bios.SerialNumber
    if (-not $s) { $s = CleanValue $cspr.IdentifyingNumber }
    if (-not $s) { $s = CleanValue $biosReg.SystemSerialNumber }
    if ($s) { $s } else { 'None' }
}

Block 'Manufacturer' { CleanValue $cs.Manufacturer }

# The maker's order code for this exact configuration: HP prints it as the
# "Product number" (6F1R8PA#AKL), Lenovo as the machine type (20XXS0AB00),
# Dell as the system SKU. Each maker keeps it somewhere different.
Block 'Product Number' {
    $model = CleanValue $cs.Model
    $isLenovo = "$($cs.Manufacturer)" -match 'Lenovo'
    $candidates = @($cs.SystemSKUNumber, $biosReg.SystemSKU, $cspr.Name, $cspr.Version)
    if ($isLenovo) { $candidates = @($cs.Model, $cspr.Name, $cs.SystemSKUNumber) }
    $pick = $null
    foreach ($c in $candidates) {
        $v = CleanValue $c
        $looksLikeName = $v -match '^(HP|Dell|Lenovo|ASUS|Acer|MSI|Toshiba|Samsung|LG|Fujitsu|Microsoft|ThinkPad|ThinkCentre)\s' -or "$v".Length -gt 30
        if ($v -and -not $pick -and -not $looksLikeName -and ($isLenovo -or $v -ne $model)) { $pick = $v }
    }
    if ($pick) { $pick } else { 'None' }
}

Block 'Model Family' {
    # Lenovo puts the friendly name here ("ThinkPad E14 Gen 4"), others repeat the model.
    $v = CleanValue $cspr.Version
    if ($v -and $v -ne (CleanValue $cs.Model)) { $v }
}

Block 'Motherboard Serial' {
    $s = CleanValue $board.SerialNumber
    if ($s) { $s } else { 'None' }
}

Block 'Device UUID' { CleanValue $cspr.UUID }

Block 'BIOS' { "$(T $bios.Manufacturer) | $(T $bios.SMBIOSBIOSVersion) | $(Day $bios.ReleaseDate)" }

Block 'Chassis' {
    $types = @{ 3='Desktop'; 4='Low profile desktop'; 6='Mini tower'; 7='Tower'; 8='Portable'; 9='Laptop'; 10='Notebook';
                13='All in one'; 14='Sub notebook'; 30='Tablet'; 31='Convertible'; 32='Detachable'; 35='Mini PC'; 36='Stick PC' }
    Wmi Win32_SystemEnclosure | ForEach-Object {
        $kind = ($_.ChassisTypes | ForEach-Object { $types[[int]$_] } | Where-Object { $_ }) -join ', '
        $tag = CleanValue $_.SMBIOSAssetTag
        "$(if ($kind) { $kind } else { 'Unknown' }) | Asset tag: $(if ($tag) { $tag } else { 'None' })"
    }
}

Block 'Memory Modules' {
    foreach ($m in $ramModules) {
        $sn = CleanValue $m.SerialNumber
        "$(T $m.DeviceLocator) | $([math]::Round($m.Capacity / 1GB, 0)) GB | $(T $m.Manufacturer) | P/N $(T $m.PartNumber) | S/N $(if ($sn) { $sn } else { 'None' })"
    }
}

Block 'Drive Serials' {
    foreach ($d in $disks) {
        $sn = CleanValue $d.Serial
        "$($d.Name) | S/N $(if ($sn) { $sn } else { 'None' }) | $($d.Bus) | Firmware $(T $d.Firmware)"
    }
}

Block 'Monitor Details' {
    $decode = { param($codes) (($codes | Where-Object { $_ -gt 0 } | ForEach-Object { [char]$_ }) -join '').Trim() }
    Wmi WmiMonitorID 'root\wmi' | ForEach-Object {
        $sn = CleanValue (& $decode $_.SerialNumberID)
        "$(& $decode $_.ManufacturerName) | $(& $decode $_.UserFriendlyName) | S/N $(if ($sn) { $sn } else { 'None' }) | Made $($_.YearOfManufacture) week $($_.WeekOfManufacture)"
    }
}

Block 'Battery' {
    Wmi Win32_Battery | ForEach-Object { "$($_.Name) | $($_.DeviceID) | Charge $($_.EstimatedChargeRemaining)%" }
}

# ---------------------------------------------------------------------------
#  New in version 2: Windows and Office licences
# ---------------------------------------------------------------------------

Say 'Reading licences...'

$licences = @()
if ($licenceJob) {
    if (Wait-Job $licenceJob -Timeout 90) { $licences = @(Receive-Job $licenceJob) }
    Remove-Job $licenceJob -Force
} else {
    $licences = @(& $licenceQuery)
}
$oemKey = (Wmi SoftwareLicensingService).OA3xOriginalProductKeyDescription

$statusWords = @{ 0='Unlicensed'; 1='Licensed'; 2='Grace period (out of box)'; 3='Grace period (hardware changed)';
                  4='Grace period (not genuine)'; 5='Notification -- not activated'; 6='Grace period (extended)' }

# The channel is what was bought: Retail (a boxed or online key, or a digital
# licence from an upgrade), OEM (came with the machine), Volume:MAK / GVLK
# (company volume licensing). Windows 7 has no ProductKeyChannel, only the
# description ("..., OEM_SLP channel").
function Channel($lic) {
    if ($lic.ProductKeyChannel) { return $lic.ProductKeyChannel }
    if ($lic.Description -match ',\s*(\S+)\s+channel') { return $Matches[1] }
    return 'Unknown'
}

$windowsAppId = '55c92734-d682-4d71-983e-d6ec3f16059f'
$officeAppId  = '0ff1ce15-a989-479d-af46-f275c6370663'

Block 'Windows Build' {
    $display = $ntReg.DisplayVersion
    if (-not $display) { $display = $ntReg.ReleaseId }
    if (-not $display) { $display = $ntReg.CSDVersion }
    $build = if ($null -ne $ntReg.UBR) { "$($ntReg.CurrentBuildNumber).$($ntReg.UBR)" } else { "$($ntReg.CurrentBuildNumber)" }
    "$display | Build $build | $($os.OSArchitecture) | Installed $(Day $os.InstallDate)"
}

Block 'Windows License' {
    $win = @($licences | Where-Object { $_.ApplicationID -eq $windowsAppId })
    if (-not $win.Count) { 'None found' }
    foreach ($lic in $win) {
        "$(Channel $lic) | $($statusWords[[int]$lic.LicenseStatus]) | Key ends $($lic.PartialProductKey) | $($lic.Name)"
    }
    # A key the maker burned into the firmware. Its presence means the machine
    # came with Windows paid for; only its edition is written, never the key.
    if ($oemKey) { "Firmware key | $oemKey" } else { 'Firmware key | None' }
}

Block 'Office License' {
    $found = $false
    # Click-to-Run: what is installed and how it is licensed. O365* is a
    # Microsoft 365 subscription; 2016/2019/2021/2024 are bought once.
    foreach ($id in "$($c2r.ProductReleaseIds)" -split ',' | Where-Object { $_ }) {
        $found = $true
        $kind = if ($id -match '^O365|^M365') { 'Subscription (Microsoft 365)' }
                elseif ($id -match 'Volume$') { 'Perpetual, volume' }
                elseif ($id -match 'Retail$') { 'Perpetual, retail' }
                else { 'Unknown' }
        "$id | $kind | Click-to-Run"
    }
    # Keys held by the licensing service: volume and older retail Office.
    foreach ($lic in $licences | Where-Object { $_.ApplicationID -eq $officeAppId }) {
        $found = $true
        "$($lic.Name) | $(Channel $lic) | $($statusWords[[int]$lic.LicenseStatus]) | Key ends $($lic.PartialProductKey)"
    }
    # Office 2010 and older MSI installs use their own service.
    Wmi OfficeSoftwareProtectionProduct 'root\cimv2' 'PartialProductKey IS NOT NULL' | ForEach-Object {
        $found = $true
        "$($_.Name) | $(Channel $_) | $($statusWords[[int]$_.LicenseStatus]) | Key ends $($_.PartialProductKey)"
    }
    # The products a signed-in Microsoft 365 account has licensed for this user.
    $next = Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\Common\Licensing\LicensingNext'
    if ($next) {
        $next.PSObject.Properties |
            Where-Object { $_.Name -notmatch '^PS|^InstalledGraceKey$|^MigrationToV5Done$' } |
            ForEach-Object { $found = $true; "Signed-in licence | $($_.Name)" }
    }
    if (-not $found) { 'None found' }
}

Block 'Office Version' {
    if ($c2r.VersionToReport) {
        $channels = @{
            '492350f6' = 'Current Channel'
            '64256afe' = 'Current Channel (Preview)'
            '55336b82' = 'Monthly Enterprise Channel'
            '7ffbc6bf' = 'Semi-Annual Enterprise Channel'
            'b8f9b850' = 'Semi-Annual Enterprise Channel (Preview)'
            '5440fd1f' = 'Beta Channel'
            'f2e724c1' = 'LTSC 2019'
            '5030841d' = 'LTSC 2021'
            '7983bac0' = 'LTSC 2024'
        }
        $url = "$($c2r.UpdateChannel) $($c2r.CDNBaseUrl)"
        $channel = 'Unknown channel'
        foreach ($key in $channels.Keys) { if ($url -match $key) { $channel = $channels[$key] } }
        "$($c2r.VersionToReport) | $($c2r.Platform) | $channel"
    } else { 'None' }
}

# Who Office is signed in as. A personal Microsoft account here, rather than a
# company one, is how a home subscription ends up on a work machine. Newer
# Office keeps only the KIND of sign-in (a _ADAL or _LiveId profile), not the address.
Block 'Office Accounts' {
    $identity = 'HKCU:\Software\Microsoft\Office\16.0\Common\Identity'
    $lines = @()
    Get-ChildItem "$identity\Identities" | ForEach-Object {
        $p = Get-ItemProperty $_.PSPath
        $email = T $p.EmailAddress
        if ($email) {
            $kind = switch -Regex ("$($p.ProviderId)") { 'AD|OrgId' { 'Work or school' } 'LiveId|MSA' { 'Personal Microsoft account' } default { 'Unknown' } }
            $lines += "$email | $kind"
        }
    }
    $profiles = @(Get-ChildItem "$identity\Profiles" | ForEach-Object { $_.PSChildName })
    $work = @($profiles | Where-Object { $_ -match '_ADAL$' }).Count
    $personal = @($profiles | Where-Object { $_ -match '_LiveId$' }).Count
    if ($work) { $lines += "Work or school sign-in | $work profile(s)" }
    if ($personal) { $lines += "Personal Microsoft account sign-in | $personal profile(s)" }
    $lines = @($lines | Select-Object -Unique)
    if ($lines.Count) { $lines } else { 'None' }
}

# Which account each Microsoft 365 licence on this PC belongs to. Office keeps
# one file per licence (JSON, base64 inside a UTF-16 wrapper) and maps the
# account's id to its email in the registry. "Valid until" is when this PC must
# next check the licence in online, not when the company's subscription ends.
# The exact plan (Business Basic, Standard, Premium) is not on the PC at all.
Block 'Office Subscription' {
    $emails = @{}
    $map = Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\Common\Licensing\LicensingNext\LicenseIdToEmailMapping'
    if ($map) { $map.PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' } | ForEach-Object { $emails[$_.Name] = "$($_.Value)" } }
    $found = $false
    foreach ($file in @(Get-ChildItem "$env:LOCALAPPDATA\Microsoft\Office\Licenses" -Recurse | Where-Object { -not $_.PSIsContainer })) {
        try {
            $raw = [IO.File]::ReadAllText($file.FullName, [Text.Encoding]::Unicode)
            if ($raw -notmatch '"License"\s*:\s*"([^"]+)"') { continue }
            $bytes = [Convert]::FromBase64String($Matches[1])
            $json = [Text.Encoding]::UTF8.GetString($bytes)
            if ($json -notmatch '^\s*\{') { $json = [Text.Encoding]::Unicode.GetString($bytes) }
            $field = { param($name) if ($json -match ('"' + $name + '"\s*:\s*"?([^",}]*)"?')) { $Matches[1] } else { '' } }
            $product = & $field 'ProductReleaseId'
            $type    = & $field 'LicenseType'
            $status  = & $field 'Status'
            $tenant  = & $field 'TenantId'
            $userId  = & $field 'UserId'
            $until   = & $field 'NotAfter'
            $trial   = (& $field 'IsTrial') -eq 'true'
            $email   = if ($emails[$userId]) { $emails[$userId] } else { 'Unknown account' }
            $day     = if ($until -match '^(\d{4}-\d{2}-\d{2})') { $Matches[1] } else { 'Unknown' }
            $found = $true
            "$product | $type | $email | Tenant $(if ($tenant) { $tenant } else { 'Unknown' }) | Valid until $day | $status$(if ($trial) { ' | Trial' })"
        } catch {}
    }
    # Licence files unreadable, but Office still says which accounts it has licences for.
    if (-not $found) {
        foreach ($email in $emails.Values | Select-Object -Unique) {
            $found = $true
            "Unknown | Unknown | $email | Tenant Unknown | Valid until Unknown | Unknown"
        }
    }
    if (-not $found) { 'None' }
}

# Whether the machine belongs to the company's Microsoft tenant. dsregcmd
# exists from Windows 10; on older machines this block stays empty.
Block 'Device Join' {
    $state = @{}
    foreach ($line in dsregcmd /status) {
        if ($line -match '^\s*(AzureAdJoined|DomainJoined|WorkplaceJoined|TenantName)\s*:\s*(.*)$') { $state[$Matches[1]] = $Matches[2].Trim() }
    }
    if ($state.Count) {
        "Azure AD joined | $($state['AzureAdJoined'])"
        "Domain joined | $($state['DomainJoined'])"
        "Workplace joined | $($state['WorkplaceJoined'])"
        if ($state['TenantName']) { "Tenant | $($state['TenantName'])" }
    }
}

# ---------------------------------------------------------------------------
#  New in version 2: installed software
# ---------------------------------------------------------------------------

Say 'Listing installed software...'

# Everything in Programs and Features, for every user-visible program. Updates,
# hotfixes and the hidden parts of bigger products are left out; they would
# bury the list and are not things anybody installed on purpose.
Block 'Installed Software' {
    $seen = @{}
    $uninstall |
        Where-Object {
            $_.SystemComponent -ne 1 -and -not $_.ParentKeyName -and
            $_.ReleaseType -notmatch 'Update|Hotfix|ServicePack' -and
            $_.DisplayName -notmatch '^(Update for|Security Update|Hotfix)|\(KB\d+'
        } |
        Sort-Object DisplayName |
        ForEach-Object {
            # A name with a colon could read as a new field, so the colon is
            # spaced out; the pipe is the separator, so a pipe becomes a slash.
            $name = ((T $_.DisplayName) -replace '\s*\|\s*', ' / ' -replace ':', ' -')
            $key = "$name|$($_.DisplayVersion)"
            if (-not $seen[$key]) {
                $seen[$key] = $true
                "$name | $(T $_.DisplayVersion) | $((T $_.Publisher) -replace '\|', '/') | $(RegDay $_.InstallDate)"
            }
        }
}

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)
Block 'Scan Info' {
    "Version $ScriptVersion | $((Get-Date).ToString('yyyy-MM-dd HH:mm')) | Run by $env:USERDOMAIN\$env:USERNAME | Admin $(if ($isAdmin) { 'Yes' } else { 'No' })"
}

$out.Add('===============================')
$out.Add('END OF REPORT')
$out.Add('===============================')

# The file is named the way IT has always named them by hand:
# "[F1 ENGINEERING] PC-NAME_.txt", or "PC-NAME_.txt" when no place was given.
$place = (("$placeAnswer".ToUpper() -replace '[\[\]\\/:*?"<>|]', '') -replace '\s+', ' ').Trim()
$fileName = if ($place) { "[$place] $($env:COMPUTERNAME)_.txt" } else { "$($env:COMPUTERNAME)_.txt" }

$text = ($out -join "`r`n") + "`r`n"
$utf8 = New-Object System.Text.UTF8Encoding $false
$path = Join-Path $reports $fileName
try {
    [IO.File]::WriteAllText($path, $text, $utf8)
} catch {
    # The scan disk can be read-only or full; the desktop always takes it.
    $path = Join-Path ([Environment]::GetFolderPath('Desktop')) $fileName
    try { [IO.File]::WriteAllText($path, $text, $utf8) } catch {
        Write-Host "  Could not save the report: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

Write-Host ''
Write-Host "  Report saved to: $path" -ForegroundColor Green
exit 0
