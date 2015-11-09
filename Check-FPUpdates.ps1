#requires -version 3.0
#https://www.microsoft.com/en-us/download/details.aspx?id=40855

<#
    Check for Adobe Flash Player updates
    
    By n01d | https://github.com/0-d/FPUpdate
#>

param(
    #SERVER
    [parameter(mandatory=$true)]
    [string]$FPIntServerRoot,
    [string]$FPDownloadRoot = 'fpdownload2.macromedia.com/pub/flashplayer/update/current/sau',
    [string]$XMLMajor = "currentmajor.xml",
    [switch]$ESR,

    #WEBREQUEST
    [string]$DownloadProxy,
    [string]$ProxyCreds,
    [String]$UserAgent,
    [switch]$Force,

    #MAIL
    [parameter(mandatory=$true)]
    [string[]]$MailTo,
    [parameter(mandatory=$true)]
    [string]$MailFrom,
    [string]$SmtpServer
)

$FPIntServerRoot += '/pub/flashplayer/update/current/sau'

#region WEBREQUEST PARAMS
$WebrequestParams = @{}

if ($DownloadProxy) {
    $WebrequestParams['Proxy']=$DownloadProxy
}

if ($ProxyCreds) {
    $cred = Get-Credential $ProxyCreds
    $WebrequestParams['ProxyCredential']=$cred
}
elseif ($DownloadProxy -and !$ProxyCreds) {
    $WebrequestParams['ProxyUseDefaultCredentials']=$true
}

if ($UserAgent) {
    $WebrequestParams['UserAgent']=$UserAgent
}

if ($Force) {
    ## ACCEPT ALL CERT
    Add-Type @"
        using System.Net;
        using System.Security.Cryptography.X509Certificates;
        public class TrustAllCertsPolicy : ICertificatePolicy {
            public bool CheckValidationResult(
                ServicePoint srvPoint, X509Certificate certificate,
                WebRequest request, int certificateProblem) {
                return true;
            }
        }
"@
    [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
}
#endregion

#region MAIL PARAMS
$MailParams = @{}
$MailParams['To'] = $MailTo
$MailParams['From'] = $MailFrom
if ($SmtpServer) {
    $MailParams['SmtpServer'] = $SmtpServer
}
#endregion

#region GET INTERNAL SERVER VERSION
Write-Host -f Gray "Getting INTERNAL server version..." -NoNewline
try{
    ## MAJOR
    $VersionMajor = $TempData = $null
    $TempData = Invoke-WebRequest -Uri "https://$FPIntServerRoot/$XMLMajor" -ErrorAction Stop
    $VersionMajor = ($TempData.Content).split('"')[1]

    ## FULL
    $XMLFull = "https://$FPIntServerRoot/$VersionMajor/xml/version.xml"
    $TempXML = [System.IO.Path]::GetTempFileName()
    Invoke-WebRequest -Uri $XMLFull -OutFile $TempXML -ErrorAction Stop
    [xml]$Full = Get-Content -Path $TempXML -ErrorAction Stop
    $FullAX = $Full.version.ActiveX
    $FPServerVersion = "$($FullAX.major).$($FullAX.minor).$($FullAX.buildMajor).$($FullAX.buildMinor)"

    Write-Host -f Green "$FPServerVersion"
    Remove-Item $TempXML -Force -ErrorAction SilentlyContinue
}
catch {
    Write-Host -f Red 'FAIL!'
    Write-Host -f Red "Can't verify INTERNAL server version! Error: [$($_.Exception.Message)]."
    break
}
#endregion

#region GET MACROMEDIA VERSION
Write-Host -f Gray "Getting MACROMEDIA server version..." -NoNewline

if ($ESR) {
    ## ESR VERSION CHECK
    try{
        $tempfile = New-TemporaryFile
        Invoke-WebRequest -Uri 'https://www.adobe.com/products/flashplayer/distribution3.html' -OutFile $tempfile @WebrequestParams
        $FullVersion = (Get-Content $tempfile | Select-String -SimpleMatch '<h1>Extended Support Release').ToString().Split(' <')[-2]
    
        Write-Host -f Green "$FullVersion"
        Remove-Item $tempfile -Force -ErrorAction SilentlyContinue
    }
    catch {
        Write-Host -f Red 'FAIL!'
        Write-Host -f Red "Can't verify MACROMEDIA server version! Error: [$($_.Exception.Message)]."
        break
    }
}
else {
    ## PUBLIC VERSION CHECK
    try {
        ## MAJOR
        $VersionMajor = $TempData = $null
        $TempData = Invoke-WebRequest -Uri "https://$FPDownloadRoot/$XMLMajor" -ErrorAction Stop @WebrequestParams
        $VersionMajor = ($TempData.Content).split('"')[1]

        ## FULL
        $XMLFull = "https://$FPDownloadRoot/$VersionMajor/xml/version.xml"
        $TempXML = [System.IO.Path]::GetTempFileName()
        Invoke-WebRequest -Uri $XMLFull -OutFile $TempXML -ErrorAction Stop @WebrequestParams
        [xml]$Full = Get-Content -Path $TempXML -ErrorAction Stop
        $FullAX = $Full.version.ActiveX
        $FullVersion = "$($FullAX.major).$($FullAX.minor).$($FullAX.buildMajor).$($FullAX.buildMinor)"

        Write-Host -f Green "$FullVersion"
        Remove-Item $TempXML -Force -ErrorAction SilentlyContinue
    }
    catch {
        Write-Host -f Red 'FAIL!'
        Write-Host -f Red "Can't verify MACROMEDIA server version! Error: [$($_.Exception.Message)]."
        break
    }
}
#endregion

#region COMPARE VERSIONS
## TRY PARSE SERVER VERSION
if ( !([version]::TryParse($FPServerVersion,[ref]$FPServerVersionParsed)) ) {

    $Message = "Can't parse INTERNAL server version: [$FPServerVersion]!"
    Write-Host -f Red $Message
    Send-MailMessage -Body $Message -Subject 'FLASH PLAYER UPDATE ERROR' -Encoding UTF8 -BodyAsHtml @MailParams
    break
}
## TRY PARSE ADOBE VERSION
if ( !([version]::TryParse($FullVersion,[ref]$FPAdobeVersionParsed)) ) {

    $Message = "Can't parse ADOBE server version: [$FullVersion]!"
    Write-Host -f Red $Message
    Send-MailMessage -Body $Message -Subject 'FLASH PLAYER UPDATE ERROR' -Encoding UTF8 -BodyAsHtml @MailParams
    break
}

if ($FPAdobeVersionParsed -gt $FPServerVersionParsed) {
    $Message = "Update is available!`n`tAdobe version:`t[$FullVersion]`n`tServer version:`t[$FPServerVersion]"
    $HTMLMessage = "<b>Update is available!</b><br /><br />
    Adobe version: <b>$FullVersion</b><br />
    Server version: <b>$FPServerVersion</b><br /><br />
    <a href='https://www.adobe.com/products/flashplayer/distribution3.html'>Download link</a>"
    
    Write-Host -f Green $Message
    Send-MailMessage -Body $HTMLMessage -Subject 'FLASH PLAYER UPDATE' -Encoding UTF8 -BodyAsHtml @MailParams
}
else {
    Write-Host -f Gray "Flash Player is Up-To-Date."
}

#endregion
