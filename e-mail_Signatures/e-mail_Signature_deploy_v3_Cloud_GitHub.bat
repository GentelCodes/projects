@echo off
setlocal

:: ================================
:: AYARLAR
:: ================================
set "GITHUB_HTML_RAW=https://github.com/GentelCodes/projects/blob/main/e-mail_Signatures/e-mail_Signature_Template.html"
set "GITHUB_PS1_RAW=https://github.com/GentelCodes/projects/blob/main/e-mail_Signatures/e-mail_Signature_deploy_v3_Cloud_GitHub.ps1"

set "TEMPLATE_DST=%TMP%\e-mail_Signature_Template.html"
set "PS1_DST=%TMP%\e-mail_Signature_deploy.ps1"

:: ================================
:: GitHub’dan dosyaları indir
:: ================================
echo HTML şablon indiriliyor...
curl -s -L "%GITHUB_HTML_RAW%" -o "%TEMPLATE_DST%"
if errorlevel 1 (
    echo HATA: HTML dosyası indirilemedi.
    pause
    exit /b 1
)

echo PS1 script indiriliyor...
curl -s -L "%GITHUB_PS1_RAW%" -o "%PS1_DST%"
if errorlevel 1 (
    echo HATA: PS1 dosyası indirilemedi.
    pause
    exit /b 1
)

:: ================================
:: Kullanıcı bilgilerini al
:: ================================
set /p DisplayName=Ad Soyad: 
set /p Title=Unvan: 
set /p MobilePhone=Cep Telefonu: 
set /p Email=E-posta: 

:: ================================
:: PS1 scriptini çalıştır
:: ================================
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS1_DST%" ^
  -DisplayName "%DisplayName%" -Title "%Title%" -MobilePhone "%MobilePhone%" -Email "%Email%"

if errorlevel 1 (
  echo Hata olustu. Lütfen yukaridaki mesaji Sistem Yöneticisine iletin.
) else (
  echo Outlook e-Posta imzanız başarılı bir şekilde olusturuldu. | GENTEL Bilişim Teknolojileri A.Ş.
)

pause
endlocal