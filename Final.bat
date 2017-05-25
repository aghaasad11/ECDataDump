:loop
@echo off
taskkill /f /im excel.exe
Set log="C:\XIMA\ScriptsProduction\ECLVL1nBILLING\Log.txt"
echo %DATE% %TIME%: Starting Process... >> %log%
pushd %~dp0
if exist X:\*.csv (
echo %DATE% %TIME%: CSV Report for Enercare L1 found in X Drive. >> %log%
) else (
echo %DATE% %TIME%: CSV Report for EnercareL1 was NOT found in X Drive. >> %log%
)
cscript Final_ECLVL1_A.vbs
if exist C:\XIMA\Output\*.txt (
echo %DATE% %TIME%: TXT Conversion for ECLVL1 Complete. >> %log%
) else (
echo %DATE% %TIME%: TXT Conversion for ECLVL1 Failed. >> %log%
)

if exist X:\Billing\*.csv (
echo %DATE% %TIME%: CSV Report for Enercare Billing found in X Drive. >> %log%
) else (
echo %DATE% %TIME%: CSV Report for Enercare Billing was NOT found in X Drive. Not running Final_BILLING1.vbs! >> %log%
)
cscript Final_BILLING1_A.vbs
if exist C:\XIMA\Billing\Output\*.txt (
echo %DATE% %TIME%: TXT Conversion for Billing Complete. >> %log%
) else (
echo %DATE% %TIME%: TXT Conversion for Billing Complete. >> %log%
)

if exist X:\ECMarket\*.csv (
echo %DATE% %TIME%: CSV Report for Enercare Market found in X Drive. >> %log%
) else (
echo %DATE% %TIME%: CSV Report for Enercare Market was NOT found in X Drive. >> %log%
)
cscript Final_ECMarket_A.vbs
if exist C:\XIMA\ECMarket\Output\*.txt (
echo %DATE% %TIME%: TXT Conversion for Enercare Market Complete. >> %log%
) else (
echo %DATE% %TIME%: TXT Conversion for Enercare Market Failed. >> %log%
)


cscript CombineAllFiles_A.vbs
echo %DATE% %TIME%: Combine Files Started. >> %log%

REM Wherever these are saved they will need to encrypted and sent through the winscp command

echo %DATE% %TIME%: Encryption Starting. >>%log%
gpg.exe --recipient EnerCare_Prod --encrypt "C:\XIMA\Output\*_NCR_EL1.txt"
echo %DATE% %TIME%: Encryption Complete. >> %log%

echo %DATE% %TIME%: Archiving NCR_EL1.txt to C:\XIMA\Output\Archive. >> %log%
move C:\XIMA\Output\*_NCR_EL1.txt C:\XIMA\Output\Archive
echo %DATE% %TIME%: Archiving NCR_Billing.txt to C:\XIMA\Billing\Output\Archive. >> %log%
move C:\XIMA\Billing\Output\*_NCR_Billing.txt C:\XIMA\Billing\Output\Archive
echo %DATE% %TIME%: Archiving NCR_ECMarket.txt to C:\XIMA\ECMarket\Output\Archive. >> %log%
move C:\XIMA\ECMarket\Output\*_NCR_ECMarket.txt C:\XIMA\ECMarket\Output\Archive

echo %DATE% %TIME%: Starting WinSCP. >> C:\XIMA\ScriptsProduction\ECLVL1nBILLING\Log.txt
"C:\Program Files\WinSCP\WinSCP.com" /command ^
"open ncrftp:ncrftp@ftpserver.enercare.ca -hostkey=""ssh-rsa 2048 32:ec:af:9f:0a:b6:a7:84:8b:2d:d9:8d:ed:23:bc:e3""" "cd /inbound" "put C:\XIMA\Output\*_NCR_EL1.txt.gpg" "bye"
echo %DATE% %TIME%: Files Uploaded to SFTP. >> %log%

echo %DATE% %TIME%: Archiving .gpg files. >> %log%
move C:\XIMA\Output\*_NCR_EL1.txt.gpg C:\XIMA\Output\Archive

echo %DATE% %TIME%: Run of Final.bat Complete. >> %log%
timeout /T 1800
Goto loop
