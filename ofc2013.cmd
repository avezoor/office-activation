@echo off
cls
color 0a
echo    __   _  _  ____  ____  __    __  ____ 
echo   / _\ / )( \(  __)(__  )/  \  /  \(  _ \
echo  /    \\ \/ / ) _)  / _/(  O )(  O ))   /
echo  \_/\_/ \__/ (____)(____)\__/  \__/(__\_)
echo.
echo  OFFICE 2013 ACTIVATION  
echo  Dev By Izzar Suly Nashrudin
pause

title Activate Microsoft Office 2013 without Software - Avezoor
cls
color 0a
(net session >nul 2>&1)
if %errorLevel% GTR 0 (
    echo.
    echo Activation cannot proceed. Please run this script as Administrator...
    echo.
    goto halt
)

echo =============================================================
echo [*] Activating Microsoft Office 2013 (No Software Needed)
echo =============================================================
echo.
echo [+] Supported Editions:
echo     - Microsoft Office Standard 2013 Volume
echo     - Microsoft Office Professional Plus 2013 Volume
echo.
echo [+] Preparing environment...

:: Navigate to Office15 folder
(if exist "%ProgramFiles%\Microsoft Office\Office15\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office15")
(if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office15")

:: Uninstall previous keys
echo [+] Removing old product keys...
cscript //nologo ospp.vbs /unpkey:92CD4 >nul
cscript //nologo ospp.vbs /unpkey:GVGXT >nul

:: Install valid Office 2013 volume keys
echo [+] Installing volume license keys...
cscript //nologo ospp.vbs /inpkey:KBKQT-2NMXY-JJWGP-M62JB-92CD4 >nul
cscript //nologo ospp.vbs /inpkey:YC7DK-G2NP3-2QQC3-J6H88-GVGXT >nul

:: Try activation servers
set i=1

:server
if %i%==1 set KMS_Sev=kms7.MSGuides.com
if %i%==2 set KMS_Sev=kms8.MSGuides.com
if %i%==3 set KMS_Sev=kms9.MSGuides.com
if %i%==4 goto notsupported

echo.
echo [+] Attempt %i%: Connecting to %KMS_Sev%
cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul

echo [+] Activating Office...
cscript //nologo ospp.vbs /act | find /i "successful" && (
    echo.
    echo =============================================================
    echo [✓] Activation Successful!
    echo -------------------------------------------------------------
    echo Support the project by starring the repo:
    echo https://github.com/avezoor/office-activation
    echo =============================================================
    echo.
    timeout /t 7 >nul
    explorer "https://github.com/avezoor"
    goto halt
) || (
    echo [!] Activation failed with %KMS_Sev%
    echo     Retrying with another server...
    echo.
    set /a i+=1
    timeout /t 5 >nul
    goto server
)

:notsupported
echo.
echo [X] Sorry, your version is not supported.
echo.

:halt
cd %~dp0
del %0 >nul
pause >nul
