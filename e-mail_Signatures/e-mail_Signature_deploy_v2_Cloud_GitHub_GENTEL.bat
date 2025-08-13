@echo off
setlocal EnableDelayedExpansion

:: Gerekli dosyalar TEMP dizininde aranır
set "TMP_PS1=%TEMP%\e-mail_Signature_deploy_v2_Cloud_GitHub_GENTEL.ps1"
set "TMP_HTML=%TEMP%\e-mail_Signature_Template.html"

:: Cloud üzerindeki dosyalar TEMP dizinine indirilir (GitHub RAW Link)
powershell -Command "Invoke-WebRequest -Uri 'https://github.com/GentelCodes/projects/blob/main/e-mail_Signatures/e-mail_Signature_deploy_v2_Cloud_GitHub_GENTEL.ps1' -OutFile '%TMP_PS1%'"
powershell -Command "Invoke-WebRequest -Uri 'https://github.com/GentelCodes/projects/blob/main/e-mail_Signatures/e-mail_Signature_Template.html' -OutFile '%TMP_HTML%'"

:: Kullanıcıdan bilgiler alınır
set /p fullName=Ad SOYAD: 
set /p title=Unvan: 
set /p mobilePhone=Mobil Telefon: 
set /p email=E-posta: 

:: PowerShell Script çalıştırılır
powershell.exe -ExecutionPolicy Bypass -File "%TMP_PS1%" -DisplayName "%fullName%" -Title "%title%" -MobilePhone "%mobilePhone%" -Email "%email%"

pause