param(
    [string]$DisplayName,
    [string]$Title,
    [string]$MobilePhone,
    [string]$Email
)

# TEMP klasöründeki HTML şablon dosya yolu
$templatePath = "$env:TEMP\e-mail_Signature_Template.html"

# Hedef Outlook imza klasörü
$signatureName = "GENTEL-Imza"
$signatureFolder = Join-Path $env:APPDATA "Microsoft\Signatures"
$htmlSignaturePath = Join-Path $signatureFolder "$signatureName.htm"

# Şablonu oku ve değişkenleri yerleştir
try {
    $templateContent = Get-Content -Path $templatePath -Raw
} catch {
    Write-Error "HTML imza şablonu bulunamadı: $templatePath"
    exit 1
}

$templateContent = $templateContent -replace '%%DisplayName%%', $DisplayName
$templateContent = $templateContent -replace '%%Title%%', $Title
$templateContent = $templateContent -replace '%%MobilePhone%%', $MobilePhone
$templateContent = $templateContent -replace '%%Email%%', $Email

# Outlook imza klasörü yoksa oluştur
if (!(Test-Path $signatureFolder)) {
    New-Item -Path $signatureFolder -ItemType Directory | Out-Null
}

# HTML imzayı oluştur
$templateContent | Set-Content -Path $htmlSignaturePath -Encoding UTF8

# Boş RTF ve TXT dosyaları oluştur (gereklidir)
"" | Set-Content -Path (Join-Path $signatureFolder "$signatureName.rtf")
"" | Set-Content -Path (Join-Path $signatureFolder "$signatureName.txt")

# Outlook sürümlerine göre registry ayarları
$officeVersions = @("16.0", "15.0", "14.0")  # 16=2016/2019/365, 15=2013, 14=2010

foreach ($version in $officeVersions) {
    $regPath = "HKCU:\Software\Microsoft\Office\$version\Common\MailSettings"
    
    if (!(Test-Path $regPath)) {
        try {
            New-Item -Path $regPath -Force | Out-Null
        } catch {
            continue
        }
    }

    try {
        Set-ItemProperty -Path $regPath -Name "NewSignature" -Value $signatureName -Force
        Set-ItemProperty -Path $regPath -Name "ReplySignature" -Value $signatureName -Force
    } catch {
        Write-Warning "Registry ayarları yapılamadı: $regPath"
    }
}