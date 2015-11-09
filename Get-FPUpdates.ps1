#requires -version 3.0
#https://www.microsoft.com/en-us/download/details.aspx?id=40855

<#
    Download updates from Adobe to your local Flash Player update server.
    
    By n01d | https://github.com/0-d/FPUpdate
#>

param(
    [parameter(mandatory=$true)]
    [string]$FPRoot,
    [string]$FPDownloadRoot = 'fpdownload2.macromedia.com/pub/flashplayer/update/current/sau',
    [string]$DownloadProxy,
    [string]$ProxyCreds,
    [String]$UserAgent,
    [switch]$Force
)

#region CHECK PARAMETERS AND FLAGS
$WebrequestParams=@{}

if ($DownloadProxy) {
    $WebrequestParams['Proxy']=$DownloadProxy
}

if ($ProxyCreds) {
    $cred = Get-Credential $ProxyCreds # this prompts for credentials
    $WebrequestParams['ProxyCredential']=$cred
}
elseif ($DownloadProxy -and !$ProxyCreds) {
    $WebrequestParams['ProxyUseDefaultCredentials']=$true
}

if ($UserAgent) {
    $WebrequestParams['UserAgent']=$UserAgent
}

if ($Force) {
    ## Make invoke-webrequest to trust all certs
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

#region CHECK ROOT PATH AND WRITE PERMISSION
## CHECK PATH
if ( !(Test-Path $FPRoot) ) {
    Write-Host -f Red "[$FPRoot] is unreachable!"
    break
}

## CHECK WRITE PERM AND CREATE DIR STRUCTURE (IF NOT)
$FPRoot += "\pub\flashplayer\update\current\sau"
try {
    Write-Host -f Gray "Checking write permissions... " -NoNewline
    $temp = "$FPRoot\WritePermTest.tmp"
    New-Item -ItemType File -Path $temp -ErrorAction Stop -Value 'Delete Me' -Force | Out-Null
    Remove-Item $temp -Force -ErrorAction SilentlyContinue

    Write-Host -f Green 'OK'
}
catch {
    Write-Host -f Red 'FAIL'
    Write-Host -f Red "Unable to write into [$FPRoot]! Error was: [$($_.Exception.Message)]"
    Write-Host -f Yellow 'Check if you have write permissions.'
    break
}
#endregion

#region DOWNLOAD UPDATES
try {
    Write-Host -f Gray "Downloading [currentmajor.xml]... " -NoNewline
    Invoke-WebRequest -Uri "https://$FPDownloadRoot/currentmajor.xml" -OutFile "$FPRoot\currentmajor.xml" -ErrorAction Stop @WebrequestParams
    Write-Host -f Green 'OK'
}
catch {
    Write-Host -f Red 'FAIL'
    Write-Host -f Red $_.Exception.Message
    if ($($_.Exception.Message) -like '*Could not establish trust relationship*') {
        Write-Host -f Yellow "Try to use '-force' flag."
    }
    break
}

11,15,16,17,18,19 | ForEach-Object {
    
    $DestXML = "$FPRoot\$_\xml"
    $DestInstall = "$FPRoot\$_\install"

    if (!(Test-Path $DestXML)) {
        New-Item -ItemType directory -Path $DestXML -Force | Out-Null
    }
    if (!(Test-Path $DestInstall)) {
        New-Item -ItemType directory -Path $DestInstall -Force | Out-Null
    }
    
    $sourceXML = "https://$FPDownloadRoot/$_/xml/version.xml"
    $sourceWinAX = "http://$FPDownloadRoot/$_/install/install_all_win_ax_sgn.z"
    $sourceWinPL = "http://$FPDownloadRoot/$_/install/install_all_win_pl_sgn.z"
    $sourceWin64AX = "http://$FPDownloadRoot/$_/install/install_all_win_64_ax_sgn.z"
    $sourceWin64PL = "http://$FPDownloadRoot/$_/install/install_all_win_64_pl_sgn.z"
    $SourceInstall = $sourceWinAX,$sourceWinPL,$sourceWin64AX,$sourceWin64PL
    
    Write-Host -f Gray "Downloading files for Flash Player version [$_]... " -NoNewline

    try {
        Invoke-WebRequest -Uri $sourceXML -OutFile "$DestXML\$($sourceXML.Split('/')[-1])" -ErrorAction Stop @WebrequestParams

        ForEach ($URI in $SourceInstall) {
            Invoke-WebRequest -Uri $URI -OutFile "$DestInstall\$($URI.split('/')[-1])" -ErrorAction Stop @WebrequestParams
        }
        Write-Host -f Green 'OK'
    }
    catch {
        Write-Host -f Red 'FAIL!'
        Write-Host -f Red $_.Exception.Message
        
        ## REMOVE EMPTY DIRS (if empty)
        if ( !(Get-ChildItem -Path $DestXML) ) {
            Remove-Item -Path $DestXML
        }
        if ( !(Get-ChildItem -Path $DestInstall) ) {
            Remove-Item -Path $DestXML
        }
        if ( !(Get-ChildItem -Path "$FPRoot\$_") ) {
            Remove-Item -Path "$FPRoot\$_"
        }
    }
}
#endregion

Write-Host -f Yellow "Finished!"
