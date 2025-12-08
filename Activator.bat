title Activator
echo off
cls

REM Check if running with admin rights
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo This script requires administrator privileges. Restarting with elevated permissions...
    powershell -Command "Start-Process -Verb RunAs -FilePath '%comspec%' -ArgumentList '/c %~dpnx0'"
    exit /b
)

cls

echo Selection

echo [1]Windows 10 
echo [2]Windows 11
echo [3]Office
echo [4]Visio
echo [5]Project
echo [6]Search for Corrupted Files (Safety and Troubleshoot)

set firstselection=0
set /p firstselection="Please Choose an option: "

if %firstselection%==1 goto Windows10
if %firstselection%==2 goto Windows11
if %firstselection%==3 goto Office
if %firstselection%==4 goto Visio
if %firstselection%==5 goto Project
if %firstselection%==6 goto Search
goto END

:Windows10
cls
echo.
echo Windows 10 Version Selection

echo [1]Home
echo [2]Home N
echo [3]Home Single Language
echo [4]Home Country Specific
echo [5]Professional
echo [6]Professional N
echo [7]Education
echo [8]Education N
echo [9]Enterprise
echo [0]Enterprise N

set tenselection=0
set /p tenselection="Please Choose an option: "

if %tenselection%==1 goto Home
if %tenselection%==2 goto HomeN
if %tenselection%==3 goto HomeSingleLanguage
if %tenselection%==4 goto HomeCountrySpecific
if %tenselection%==5 goto Professional
if %tenselection%==6 goto ProfessionalN
if %tenselection%==7 goto Education
if %tenselection%==8 goto EducationN
if %tenselection%==9 goto Enterprise
if %tenselection%==0 goto EnterpriseN
goto END

:Windows11
cls
echo.
echo Windows 11 Version Selection

echo [1]Home
echo [2]Home N
echo [3]Home Single Language
echo [4]Home Country Specific
echo [5]Professional
echo [6]Professional N
echo [7]Education
echo [8]Education N
echo [9]Enterprise
echo [0]Enterprise N

set selection=0
set /p selection="Please Choose an option: "

if %selection%==1 goto Home
if %selection%==2 goto HomeN
if %selection%==3 goto HomeSingleLanguage
if %selection%==4 goto HomeCountrySpecific
if %selection%==5 goto Professional
if %selection%==6 goto ProfessionalN
if %selection%==7 goto Education
if %selection%==8 goto EducationN
if %selection%==9 goto Enterprise
if %selection%==0 goto EnterpriseN
goto END

:Office
cls
echo.
echo MS Office Version Selection

echo [1]Office 2010 or 2013
echo [2]Office 2016
echo [3]Office 2019
echo [4]Office 2021
echo [5]Office 365

set OfficeSelection=0
set /p OfficeSelection="Please Choose an option: "

if %OfficeSelection%==1 goto Office2010or13
if %OfficeSelection%==2 goto Office2016
if %OfficeSelection%==3 goto Office2019
if %OfficeSelection%==4 goto Office2021
if %OfficeSelection%==5 goto Office365
goto END

:Visio
cls
echo.
echo MS Visio Version Selection

echo [1]Visio 2016
echo [2]Visio 2019
echo [3]Visio 2021

set VisioSelection=0
set /p VisioSelection="Please Choose an option: "

if %VisioSelection%==1 goto Visio2016
if %VisioSelection%==2 goto Visio2019
if %VisioSelection%==3 goto Visio2021
goto END

:Project
cls
echo.
echo MS Project Version Selection

echo [1]Project 2016
echo [2]Project 2019
echo [3]Project 2021

set ProjectSelection=0
set /p ProjectSelection="Please Choose an option: "

if %ProjectSelection%==1 goto Project2016
if %ProjectSelection%==2 goto Project2019
if %ProjectSelection%==3 goto Project2021
goto END

:Project2016
cls
cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_KMS_Client-ppd.xrm-ms"&cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_KMS_Client-ul-oob.xrm-ms"&cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_KMS_Client-ul.xrm-ms"&cscript ospp.vbs /inslic:"..\root\Licenses16\client-issuance-bridge-office.xrm-ms"&cscript ospp.vbs /inslic:"..\root\Licenses16\client-issuance-root.xrm-ms"&cscript ospp.vbs /inslic:"..\root\Licenses16\client-issuance-root-bridge-test.xrm-ms"&cscript ospp.vbs /inslic:"..\root\Licenses16\client-issuance-stil.xrm-ms"&cscript ospp.vbs /inslic:"..\root\Licenses16\client-issuance-ul.xrm-ms"&cscript ospp.vbs /inslic:"..\root\Licenses16\client-issuance-ul-oob.xrm-ms"&cscript ospp.vbs /inslic:"..\root\Licenses16\pkeyconfig-office.xrm-ms"
cscript ospp.vbs /inpkey:YG9NW-3K39V-2T3HJ-93F3Q-G83KT
cscript ospp.vbs /sethst:107.175.77.7
cscript ospp.vbs /setprt:1688
cscript ospp.vbs /act
goto END

:Project2019
cls
cscript ospp.vbs /inslic:"..\root\Licenses16\pkeyconfig-office.xrm-ms"&(for /f %x in ('dir /b ..\root\Licenses16\client-issuance*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x")&(for /f %x in ('dir /b ..\root\Licenses16\projectprovl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x")&(for /f %x in ('dir /b ..\root\Licenses16\projectpro2019vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x")
cscript ospp.vbs /inpkey:B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B
cscript ospp.vbs /sethst:107.175.77.7
cscript ospp.vbs /setprt:1688
cscript ospp.vbs /act
goto END

:Project2021
cls
cscript ospp.vbs /inslic:"..\root\Licenses16\pkeyconfig-office.xrm-ms"&(for /f %x in ('dir /b ..\root\Licenses16\client-issuance*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x")&(for /f %x in ('dir /b ..\root\Licenses16\projectprovl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x")&(for /f %x in ('dir /b ..\root\Licenses16\projectpro2021vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x")
cscript ospp.vbs /inpkey:FTNWT-C6WBT-8HMGF-K9PRX-QV9H8
cscript ospp.vbs /sethst:107.175.77.7
cscript ospp.vbs /setprt:1688
cscript ospp.vbs /act
goto END

:Visio2016
cls
cscript ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_KMS_Client-ppd.xrm-ms" &cscript ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_KMS_Client-ul-oob.xrm-ms"&cscript ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_KMS_Client-ul.xrm-ms"&cscript ospp.vbs /inslic:"..\root\Licenses16\client-issuance-bridge-office.xrm-ms"&cscript ospp.vbs /inslic:"..\root\Licenses16\client-issuance-root.xrm-ms"&cscript ospp.vbs /inslic:"..\root\Licenses16\client-issuance-root-bridge-test.xrm-ms"&cscript ospp.vbs /inslic:"..\root\Licenses16\client-issuance-stil.xrm-ms"&cscript ospp.vbs /inslic:"..\root\Licenses16\client-issuance-ul.xrm-ms"&cscript ospp.vbs /inslic:"..\root\Licenses16\client-issuance-ul-oob.xrm-ms"&cscript ospp.vbs /inslic:"..\root\Licenses16\pkeyconfig-office.xrm-ms"
cscript ospp.vbs /inpkey:KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4
cscript ospp.vbs /sethst:107.175.77.7
cscript ospp.vbs /setprt:1688
cscript ospp.vbs /act
goto END

:Visio2019
cls
cscript ospp.vbs /inslic:"..\root\Licenses16\pkeyconfig-office.xrm-ms"&(for /f %x in ('dir /b ..\root\Licenses16\client-issuance*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x")&(for /f %x in ('dir /b ..\root\Licenses16\visioprovl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x")&(for /f %x in ('dir /b ..\root\Licenses16\visiopro2019vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x")
cscript ospp.vbs /inpkey:9BGNQ-K37YR-RQHF2-38RQ3-7VCBB
cscript ospp.vbs /sethst:107.175.77.7
cscript ospp.vbs /setprt:1688
cscript ospp.vbs /act
goto END

:Visio2021
cls
cscript ospp.vbs /inslic:"..\root\Licenses16\pkeyconfig-office.xrm-ms"&(for /f %x in ('dir /b ..\root\Licenses16\client-issuance*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x")&(for /f %x in ('dir /b ..\root\Licenses16\visioprovl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x")&(for /f %x in ('dir /b ..\root\Licenses16\visiopro2021vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x")
cscript ospp.vbs /inpkey:KNH8D-FGHT4-T8RK3-CTDYJ-K2HT
cscript ospp.vbs /sethst:107.175.77.7
cscript ospp.vbs /setprt:1688
cscript ospp.vbs /act
goto END

:Office2010or13
cls
if exist "%ProgramFiles%\Microsoft Office\Office15\ospp.vbs" set folder="%ProgramFiles%\Microsoft Office\Office15"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs" set folder="%ProgramFiles(x86)%\Microsoft Office\Office15"
if exist "%ProgramFiles%\Microsoft Office\Office14\ospp.vbs" set folder="%ProgramFiles%\Microsoft Office\Office14"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office14\ospp.vbs" set folder="%ProgramFiles(x86)%\Microsoft Office\Office14"
cscript %folder%\ospp.vbs /inpkey:VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB
cscript %folder%\ospp.vbs /inpkey:YC7DK-G2NP3-2QQC3-J6H88-GVGXT
cscript %folder%\ospp.vbs /sethst:kms8.msguides.com
cscript %folder%\ospp.vbs /setprt:1688
cscript %folder%\ospp.vbs /act
goto END

:Office2016
cls
cd /d %ProgramFiles%\Microsoft Office\Office16
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16
for /f %x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"
cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
cscript ospp.vbs /unpkey:BTDRB >nul
cscript ospp.vbs /unpkey:KHGM9 >nul
cscript ospp.vbs /unpkey:CPQVG >nul
cscript ospp.vbs /sethst:107.175.77.7
cscript ospp.vbs /setprt:1688
cscript ospp.vbs /act
goto END

:Office2019
cls
cd /d %ProgramFiles%\Microsoft Office\Office16
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16
for /f %x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"
cscript ospp.vbs /setprt:1688
cscript ospp.vbs /unpkey:6MWKP >nul
cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
cscript ospp.vbs /sethst:107.175.77.7
cscript ospp.vbs /act
goto END

:Office2021
cls
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16
cd /d %ProgramFiles%\Microsoft Office\Office16
for /f %x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"
cscript ospp.vbs /setprt:1688
cscript ospp.vbs /unpkey:6F7TH >nul
cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
cscript ospp.vbs /sethst:107.175.77.7
cscript ospp.vbs /act
goto END

:Office365
cls
cd /d %ProgramFiles%\Microsoft Office\Office16
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16
for /f %x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"
cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
cscript ospp.vbs /unpkey:BTDRB >nul
cscript ospp.vbs /unpkey:KHGM9 >nul
cscript ospp.vbs /unpkey:CPQVG >nul
cscript ospp.vbs /sethst:107.175.77.7
cscript ospp.vbs /setprt:1688
cscript ospp.vbs /act


:Home
cls
slmgr /ipk TX9XD-98N7V-6WMQ6-BX7FG-H8Q99
slmgr /skms kms8.msguides.com
slmgr /ato
goto END

:HomeN
cls
slmgr /ipk 3KHY7-WNT83-DGQKR-F7HPR-844BM
slmgr /skms kms8.msguides.com
slmgr /ato
goto END

:HomeSingleLanguage
cls
slmgr /ipk 7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH
slmgr /skms kms8.msguides.com
slmgr /ato
goto END

:HomeCountrySpecific
cls
slmgr /ipk PVMJN-6DFY6-9CCP6-7BKTT-D3WVR
slmgr /skms kms8.msguides.com
slmgr /ato
goto END

:Professional
cls
slmgr /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
slmgr /skms kms8.msguides.com
slmgr /ato
goto END

:ProfessionalN
cls
slmgr /ipk MH37W-N47XK-V7XM9-C7227-GCQG9
slmgr /skms kms8.msguides.com
slmgr /ato
goto END

:Education
cls
slmgr /ipk NW6C2-QMPVW-D7KKK-3GKT6-VCFB2
slmgr /skms kms8.msguides.com
slmgr /ato
goto END

:EducationN
cls
slmgr /ipk 2WH4N-8QGBV-H22JP-CT43Q-MDWWJ
slmgr /skms kms8.msguides.com
slmgr /ato
goto END

:Enterprise
cls
slmgr /ipk NPPR9-FWDCX-D2C8J-H872K-2YT43
slmgr /skms kms8.msguides.com
slmgr /ato
goto END

:EnterpriseN
cls
slmgr /ipk DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4
slmgr /skms kms8.msguides.com
slmgr /ato
goto END

:Search
cls
echo Searching for Corrupted Files
echo.
sfc /scannow
echo.
DISM /Online /Cleanup-Image /CheckHealth
echo.
DISM /Online /Cleanup-Image /ScanHealth
echo.
DISM /Online /Cleanup-Image /RestoreHealth
echo.

:END
echo.
echo This is the end of the script you should now have a active version of the product you choose
echo if it didnt work try again you probably did something wrong
echo Also you can Look above at the logs it should tell you if the activation was sucessful
echo if you used the Search Option it will display above if it found corrupted files and if they were fixed
pause