#requires -version 3.0
param(
    [Parameter(Mandatory=$true)][string]$DisplayName,
    [Parameter(Mandatory=$true)][string]$Title,
    [Parameter(Mandatory=$true)][string]$MobilePhone,
    [Parameter(Mandatory=$true)][string]$Email
)

# Yürütmeyi profilden etkilenmesin diye öneri: powershell -NoProfile ile çağır
# Şablon dosyası TEMP'e kopyalanacak (BAT tarafında)
$templatePath = Join-Path $env:TEMP "e-mail_Signature_Template.html"

# Outlook imza klasörü/isimleri
$signatureName    = "GENTEL-Imza"
$signatureFolder  = Join-Path $env:APPDATA "Microsoft\Signatures"
$htmlSignature    = Join-Path $signatureFolder "$signatureName.htm"
$rtfSignature     = Join-Path $signatureFolder "$signatureName.rtf"
$txtSignature     = Join-Path $signatureFolder "$signatureName.txt"

# PS1 içine yanlışlıkla HTML sızmış mı? (koruyucu kontrol)
$psSelfContent = Get-Content -Raw -LiteralPath $PSCommandPath
if ($psSelfContent -match '<html|<style|--tab-size-preference|</body>|</html>') {
    Write-Error "HTML içerik PS1 dosyasına karışmış görünüyor. BAT akışını kontrol et (PS1'e ECHO/APPEND yapılmamalı)."
    exit 2
}

# Şablonu oku
if (!(Test-Path -LiteralPath $templatePath)) {
    Write-Error "HTML imza şablonu bulunamadı: $templatePath"
    exit 1
}
$templateContent = Get-Content -Raw -LiteralPath $templatePath

# Yer tutucuları değiştir (regex değil, düz Replace)
$templateContent = $templateContent.Replace('%%DisplayName%%', $DisplayName)
$templateContent = $templateContent.Replace('%%Title%%',       $Title)
$templateContent = $templateContent.Replace('%%MobilePhone%%', $MobilePhone)
$templateContent = $templateContent.Replace('%%Email%%',       $Email)

# İmza klasörü yoksa oluştur
if (!(Test-Path -LiteralPath $signatureFolder)) {
    New-Item -ItemType Directory -Path $signatureFolder | Out-Null
}

# HTML imzayı yaz (Outlook .htm kullanır)
Set-Content -LiteralPath $htmlSignature -Value $templateContent -Encoding UTF8

# Boş RTF/TXT (bazı Outlook sürümleri bunları listeler)
Set-Content -LiteralPath $rtfSignature -Value '' -Encoding Unicode
Set-Content -LiteralPath $txtSignature -Value '' -Encoding UTF8

# Outlook registry ayarları (New/Reply imzaları ata)
$officeVersions = @('16.0','15.0','14.0')  # 16=2016/2019/365, 15=2013, 14=2010
foreach ($v in $officeVersions) {
    $reg = "HKCU:\Software\Microsoft\Office\$v\Common\MailSettings"
    if (!(Test-Path $reg)) { New-Item -Path $reg -Force | Out-Null }
    Set-ItemProperty -Path $reg -Name NewSignature   -Value $signatureName -Force
    Set-ItemProperty -Path $reg -Name ReplySignature -Value $signatureName -Force
}

Write-Host "Outlook e-Posta imzanız başarılı bir şekilde oluşturuldu: $htmlSignature"
