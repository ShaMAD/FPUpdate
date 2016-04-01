#requires -version 3.0
#https://www.microsoft.com/en-us/download/details.aspx?id=40855

<#
    Download updates from Adobe to your local Flash Player update server.
    
    By n01d | https://github.com/0-d/FPUpdate
    Version: 20160401

    Syntax:
        FPCheckUpdate.ps1 -FPIntServerRoot <string> [-FPDownloadRoot <string>] [-XMLMajor <string>] [-ESR <string>] [-Proxy <string>] [-ProxyCreds <string>] [-UserAgent <string>] [-Force] [-SmtpServer <string>] [-LogFile <string>] [<CommonParameters>]
        FPCheckUpdate.ps1 -FPIntServerRoot <string> -MailTo <mailaddress[]> -MailFrom <mailaddress> [-FPDownloadRoot <string>] [-XMLMajor <string>] [-ESR <string>] [-Proxy <string>] [-ProxyCreds <string>] [-UserAgent <string>] [-Force] [-SmtpServer <string>] [-LogFile <string>] [<CommonParameters>]

    Example:
        .\FPCheckUpdate.ps1 -FPIntServerRoot 'fp-update.domain.local' -ESR '123098abc-blah-blah-blah-098765432' -Proxy 'http://proxy.domain.local' -UserAgent InternetExplorer -Force -MailTo 'admin@company.com','support@company.com' -MailFrom 'FP@company.com' -SmtpServer 'smtp.company.com'
#>

[CmdletBinding(DefaultParametersetName="Default")]
param(
    #SERVER
    [parameter(ParameterSetName='WithMail',mandatory=$true)]
    [parameter(ParameterSetName='Default',mandatory=$true)]
    [string]$FPIntServerRoot,
    [string]$FPDownloadRoot = 'fpdownload2.macromedia.com/pub/flashplayer/update/current/sau',
    [string]$XMLMajor = "currentmajor.xml",
    [string]$ESR,

    #WEBREQUEST
    [string]$Proxy,
    [string]$ProxyCreds,
    [ValidateSet('InternetExplorer', 'FireFox', 'Chrome', 'Opera', 'Safari')]
    [String]$UserAgent,
    [switch]$Force,

    #MAIL
    [parameter(ParameterSetName='WithMail',mandatory=$true)]
    [mailaddress[]]$MailTo,
    [parameter(ParameterSetName='WithMail',mandatory=$true)]
    [mailaddress]$MailFrom,
    [string]$SmtpServer,

    #MISC
    [string]$LogFile = "$env:TEMP\FP.CheckUpdates.log"
)

$FPIntServerRoot += '/pub/flashplayer/update/current/sau'
if ($ESR) {
    $ESR = "https://www.adobe.com/ru/products/flashplayer/distribution4.html?auth=$ESR"
    Write-Debug -Message "ESR link: [$ESR]"
}

function Write-Log {
    param(
        [parameter(mandatory=$true)]
        [string]$Message,
        [ValidateSet('Error','Warning','Info')]
        [string]$Type = 'Info',
        [ValidateSet('Black', 'DarkBlue', 'DarkGreen', 'DarkCyan', 'DarkRed', 'DarkMagenta', 
        'DarkYellow', 'Gray', 'DarkGray', 'Blue', 'Green', 'Cyan', 'Red', 'Magenta', 'Yellow', 'White')]
        [string]$Color,
        [switch]$NoNewLine
    )

    $InLine = @{}
    if ($NoNewLine) {
        $InLine['NoNewLine'] = $true
    }

    if (!$Color) {
        switch ($Type) {
            'Error' {$Color = 'Red'; break}
            'Warning' {$Color = 'Yellow'; break}
            'Info' {$Color = 'Gray'; break}
        }
    }

    if ( !(Test-Path $LogFile) ) {
        New-Item -Path $LogFile -ItemType File -Force
    }
    elseif ((Get-Item $LogFile).CreationTime -lt [datetime]::Today) {
        # Если создан не сегодня - обнуляем!
        $null | Out-File -FilePath $LogFile -Force -Encoding utf8
    }

    Write-Host -ForegroundColor $color $Message @InLine
    Write-Output "[$(Get-Date -Format 'dd.MM.yy HH.mm.ss')][$Type]`t$Message" | Out-File -FilePath $LogFile -Encoding utf8 -Append
}

Write-Log ">--------- SCRIPT STARTED --------->"

#region WEBREQUEST PARAMS
$WebrequestParams = @{}

if ($Proxy) {
    $WebrequestParams['Proxy']=$Proxy
}

if ($ProxyCreds) {
    $cred = Get-Credential $ProxyCreds
    $WebrequestParams['ProxyCredential']=$cred
}
elseif ($Proxy -and !$ProxyCreds) {
    $WebrequestParams['ProxyUseDefaultCredentials']=$true
}

if ($UserAgent) {
    $WebrequestParams['UserAgent'] = [Microsoft.PowerShell.Commands.PSUserAgent]::$UserAgent
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
Write-Log "Getting INTERNAL server version...`t" -NoNewline
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

    Write-Log -Color Green "$FPServerVersion"
    Remove-Item $TempXML -Force -ErrorAction SilentlyContinue
}
catch {
    Write-Log -Type Error 'FAIL!'
    Write-Log -Type Error "Can't verify INTERNAL server version! Error: [$($_.Exception.Message)]."
    Write-Log "<--------- SCRIPT FINISHED ---------<"
    if ($MailFrom) {
        Send-MailMessage -Subject "Error checking FlashPlayer version!" -Body $("Error details:`n"+$_) -Encoding UTF8 -Attachments $LogFile @MailParams
    }
    break
}
#endregion

#region GET ADOBE VERSION
for ($i = 0; $i -le 10; $i++) {
    if ($ESR) {
        ## ESR VERSION CHECK
        Write-Log "Getting ADOBE server version (ESR)...`t" -NoNewline
        try{
            $tempfile = [system.io.path]::GetTempFileName()
            Invoke-WebRequest -Uri $ESR -OutFile $tempfile -ErrorAction Stop @WebrequestParams
        
            $versions = @()
            $rx = [System.Text.RegularExpressions.Regex]('(?:[^\d]|^)(?<v>\d+\.\d+\.\d+\.\d+)(?:[^\d]|$)')
            $m = $rx.Matches((Get-Content $tempfile | Select-String -SimpleMatch '<h1>Extended Support Release').ToString())

            #Write-Host "Found [$($m.Count)] matches"
            $m | ForEach-Object {
                try {
                    $local:ver = [System.Version] ($_.Groups['v'].Value)
                    $versions +=  @($local:ver)
                }
                catch {
                    Write-Host "Error [$($_.Exception.Message)]"
                }
            }

            $FullVersion = ($($versions | Measure-Object -Maximum ).Maximum).ToString()
            $ESRMajorVersion = $FullVersion.Split('.')[0]

            $ESRDownloadURI = (Get-Content $tempfile | Select-String -SimpleMatch "/fp_background_update_$ESRMajorVersion.cab").ToString().Split('"')[1]
    
            Write-Log -Color Green "$FullVersion"
            $i = 11 # Stop FOR
            Remove-Item $tempfile -Force -ErrorAction SilentlyContinue
        }
        catch {
            if ($_ -like '*The operation has timed out.*') {
                Write-Log -Type Error 'Timeout!'
                # Next iteration
            }
            else {
                Write-Log -Type Error 'FAIL!'
                Write-Log -Type Error "Can't verify MACROMEDIA server version (ESR)! Error: [$($_.Exception.Message)]."
                Write-Log "<--------- SCRIPT FINISHED ---------<"
                if ($MailFrom) {
                    Send-MailMessage -Subject "Error checking FlashPlayer version (ESR)!" -Body $("Error details:`n"+$_) -Encoding UTF8 -Attachments $LogFile @MailParams
                }
                Exit
            }
        }
    }
    else {
        ## PUBLIC VERSION CHECK
        Write-Log "Getting ADOBE server version...`t" -NoNewline
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

            Write-Log -Color Green "$FullVersion"
            $i = 11 # Stop FOR
            Remove-Item $TempXML -Force -ErrorAction SilentlyContinue
        }
        catch {
            if ($_ -like '*The operation has timed out.*') {
                Write-Log -Type Error 'Timeout!'
                # Next iteration
            }
            else {
                Write-Log -Type Error 'FAIL!'
                Write-Log -Type Error "Can't verify MACROMEDIA server version! Error: [$($_.Exception.Message)]."
                Write-Log "<--------- SCRIPT FINISHED ---------<"
                if ($MailFrom) {
                    Send-MailMessage -Subject "Error checking FlashPlayer version!" -Body $("Error details:`n"+$_) -Encoding UTF8 -Attachments $LogFile @MailParams
                }
                Exit
            }
        }
    }
}
#endregion

#region COMPARE VERSIONS
[version]$FPServerVersionParsed = $null
[version]$FPAdobeVersionParsed = $null
## TRY PARSE SERVER VERSION
if ( !([version]::TryParse($FPServerVersion,[ref]$FPServerVersionParsed)) ) {

    $Message = "Can't parse INTERNAL server version: [$FPServerVersion]!"
    Write-Log -Type Error $Message
    if ($MailFrom) {
        Send-MailMessage -Body $Message -Subject 'FLASH PLAYER UPDATE ERROR' -Encoding UTF8 -BodyAsHtml @MailParams
    }
    Write-Log "<--------- SCRIPT FINISHED ---------<"
    break
}
## TRY PARSE ADOBE VERSION
if ( !([version]::TryParse($FullVersion,[ref]$FPAdobeVersionParsed)) ) {

    $Message = "Can't parse ADOBE server version: [$FullVersion]!"
    Write-Log -Type Error $Message
    if ($MailFrom) {
        Send-MailMessage -Body $Message -Subject 'FLASH PLAYER UPDATE ERROR' -Encoding UTF8 -BodyAsHtml @MailParams
    }
    Write-Log "<--------- SCRIPT FINISHED ---------<"
    break
}

if ($FPAdobeVersionParsed -gt $FPServerVersionParsed) {

        if ($ESRDownloadURI) {
            $DownloadLink = $ESRDownloadURI
        }
        else {
            $DownloadLink = 'https://www.adobe.com/products/flashplayer/distribution3.html'
        }

        $Message = "Update is available!`n`tAdobe version:`t[$FullVersion]`n`tServer version:`t[$FPServerVersion]"
        $HTMLMessage = "<b>Update is available!</b><br /><br />
        Adobe version: <b>$FullVersion</b><br />
        Server version: <b>$FPServerVersion</b><br /><br />
        <a href='$DownloadLink'>Download</a><br />
        Link: $DownloadLink"
    
        Write-Log -Color Green $Message
        Write-Log "Sending mail... " -NoNewline
        try {
            if ($MailFrom) {
                Send-MailMessage -Body $HTMLMessage -Subject 'FLASH PLAYER UPDATE AVAILABLE' -Encoding UTF8 -BodyAsHtml -ErrorAction Stop @MailParams
            }
            Write-Log -Color Green 'OK'
        }
        catch {
            Write-Log -Type Error 'FAIL'
            Write-Log -Type Error "Mail not sent! Error was: [$($_.Exception.Message)]."
            Write-Log "<--------- SCRIPT FINISHED ---------<"
            break
        }
}
else {
    Write-Log "Flash Player is Up-To-Date."
}
#endregion

Write-Log "<--------- SCRIPT FINISHED ---------<"
