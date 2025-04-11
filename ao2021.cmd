@echo off
cls
color 0a
echo    __   _  _  ____  ____  __    __  ____ 
echo   / _\ / )( \(  __)(__  )/  \  /  \(  _ \
echo  /    \\ \/ / ) _)  / _/(  O )(  O ))   /
echo  \_/\_/ \__/ (____)(____)\__/  \__/(__\_)
echo.
echo  OFFICE ACTIVATION 
echo  Dev By Izzar Suly N.
pause

title Activate Microsoft Office 2021 without Software - Avezoor
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
echo [*] Activating Microsoft Office 2021 (No Software Needed)
echo =============================================================
echo.
echo [+] Supported Editions:
echo     - Microsoft Office Standard 2021
echo     - Microsoft Office Professional Plus 2021
echo.
echo [+] Preparing environment...

(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")
(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")

tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>NUL | find /I /N "WINWORD.EXE">NUL && taskkill /F /IM WINWORD.EXE >nul
tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>NUL | find /I /N "EXCEL.EXE">NUL && taskkill /F /IM EXCEL.EXE >nul
tasklist /FI "IMAGENAME eq POWERPNT.EXE" 2>NUL | find /I /N "POWERPNT.EXE">NUL && taskkill /F /IM POWERPNT.EXE >nul

REG DELETE "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\Licensing" /f >nul 2>NUL
cscript //nologo ospp.vbs /setprt:1688 >nul || goto wshdisabled

for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul

cscript //nologo ospp.vbs /unpkey:WFG99 >nul
cscript //nologo ospp.vbs /unpkey:6MWKP >nul
cscript //nologo ospp.vbs /unpkey:6F7TH >nul

:ucpkey
cscript //nologo ospp.vbs /dstatus | find /i "Last 5 characters of installed product key" > temp.txt || goto skunpkey
set /P bpkey= < temp.txt
set pkeyt=%bpkey:~-5%
echo.
cscript //nologo ospp.vbs /unpkey:%pkeyt% >nul
goto ucpkey

:skunpkey
echo.
echo =============================================================
echo [+] Installing Volume License Key...
echo =============================================================

cscript //nologo slmgr.vbs /ckms >nul
cscript //nologo ospp.vbs /setprt:1688 >nul
set i=1
cscript //nologo ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH >nul || goto notsupported

:skms
if %i% GTR 5 (goto busy) else if %i% LEQ 5 (set KMS=s8.mshost.pro)
cscript //nologo ospp.vbs /sethst:%KMS% >nul

:ato
echo.
echo =============================================================
echo [+] Attempt %i%: Activating via %KMS%
echo =============================================================
cscript //nologo ospp.vbs /act >nul
timeout /t 5 >nul
echo.

cscript //nologo ospp.vbs /act | find /i "successful" && (
    echo.
    echo =============================================================
    echo [✓] Activation Successful!
    echo -------------------------------------------------------------
    echo ⭐ Support the project by starring the repo:
    echo https://github.com/avezoor/office-act
    echo =============================================================
    echo.
    if errorlevel 2 exit
) || (
    echo Activation is taking longer than expected, please wait...
    echo.
    set /a i+=1
    timeout /t 10 >nul
    goto skms
)

timeout /t 7 >nul
explorer "https://github.com/avezoor"
echo.
echo If activation fails, feel free to open an issue on GitHub: https://github.com/avezoor
echo.
goto halt

:notsupported
echo.
echo =============================================================
echo [X] Sorry, your version is not supported.
echo.
goto halt

:wshdisabled
echo =============================================================
echo [X] Windows Script Host is disabled.
echo [!] Please enable WSH via this page:
echo https://github.com/avezoor/office-act
echo.
timeout /t 7 >nul
explorer "https://github.com/avezoor"
goto halt

:busy
echo =============================================================
echo [X] Activation failed due to unstable connection.
echo [!] Please try a different network or check your internet.
echo.
goto halt

:halt
cd %~dp0
del %0 >nul
pause >nul
