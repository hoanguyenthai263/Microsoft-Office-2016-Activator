# Microsoft Office 2016 Activator

This command-line script enables the activation of Microsoft Office 2016 (Standard & Professional Plus) through a KMS server. It is provided as a free project and operates without the need for supplementary software or tool installations.

![Microsoft Office 2016](https://camo.githubusercontent.com/da6cc631bcd399a0cdeb08d81eef0ddd82e6c0b5a57a695f5ada792de3fecc67/68747470733a2f2f342e62702e626c6f6773706f742e636f6d2f2d6f372d365064685f4443342f57794338436762397046492f4141414141414141454d632f70354a77766c4949476f7739464949683867696d596f4c34354f663157786d3951434c63424741732f733332302f3230303070782d4d6963726f736f66745f4f66666963655f323031335f6c6f676f5f616e645f776f72646d61726b2e7376672e706e67)


## Creating the Activator File

There are two methods to acquire the necessary activator file. The first involves manually constructing a batch file `(.cmd)` by opening a text/code editor (for instance, Notepad, WordPad, Visual Studio Code), writing the appropriate commands, and saving the file as a .cmd file. The second method is to directly download the `Activator.cmd` file provided within this repository.

```bash
@echo off
title Microsoft Office 2016 Activator - HoafHokB3os&cls&echo.&echo ****************************************************************************&echo Microsoft Office 2016 Activator for FREE without any software!&echo Nguyen Thai Hoa&echo HoafHokB3os(c)2025 &echo.&echo.****************************************************************************&echo.&echo #This project is using KMS server.&echo.&echo #Supported products:&echo - Microsoft Office Standard 2016&echo - Microsoft Office Professional Plus 2016&echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&echo.&echo ****************************************************************************&echo Activating your Microsoft Office...&echo.&cscript //nologo ospp.vbs /unpkey:WFG99 >nul&cscript //nologo ospp.vbs /unpkey:DRTFM >nul&cscript //nologo ospp.vbs /unpkey:BTDRB >nul&cscript //nologo ospp.vbs /unpkey:CPQVG >nul&cscript //nologo ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 >nul&set i=1
:server
if %i%==1 set KMS_Sev=kms7.MSGuides.com
if %i%==2 set KMS_Sev=kms8.MSGuides.com
if %i%==3 set KMS_Sev=kms9.MSGuides.com
if %i%==4 goto notsupported
cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul&echo ****************************************************************************&echo.&choice /n /c YN /m "Done. Wanna see my project on Github [y/n]?" & if errorlevel 2 exit) || (echo The connection to my KMS server failed! Trying to connect to another one... & echo Please wait... & echo. & echo. & set /a i+=1 & goto server)
explorer "https://github.com/hoanguyenthai263"&goto halt
:notsupported
echo.&echo ***************************************************************************=&echo Sorry! Your version is not supported.&echo Please try installing the latest version!
:halt
pause
```
## Executing the Activation File

Ensure a stable internet connection is established before proceeding. To execute the activation script, locate the `.cmd` file, right-click on it, and select "Run as administrator." The script will then automatically attempt to connect to a KMS server to activate your Microsoft Office 2016 installation. Allow sufficient time for the activation process to complete.

![Microsoft Office 2016](https://github.com/hoanguyenthai263/Microsoft-Office-2016-Activator/blob/main/Activator.png)
## Enjoy Your Activated Office
If all went according to plan, your Microsoft Office 2016 should be activated. You can now type `n` to close this window, or type `y` if you'd like to see my other projects on GitHub. Enjoy using your activated Office!

