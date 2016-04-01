# FPUpdate
PowerShell scripts to check and update content on an internal Flash Player update server.

    Syntax:
        FPCheckUpdate.ps1 -FPIntServerRoot <string> [-FPDownloadRoot <string>] [-XMLMajor <string>] [-ESR <string>] [-Proxy <string>] [-ProxyCreds <string>] [-UserAgent <string>] [-Force] [-SmtpServer <string>] [-LogFile <string>] [<CommonParameters>]
        FPCheckUpdate.ps1 -FPIntServerRoot <string> -MailTo <mailaddress[]> -MailFrom <mailaddress> [-FPDownloadRoot <string>] [-XMLMajor <string>] [-ESR <string>] [-Proxy <string>] [-ProxyCreds <string>] [-UserAgent <string>] [-Force] [-SmtpServer <string>] [-LogFile <string>] [<CommonParameters>]

    Parameter explanation:
        Server params:
            -FPIntServerRoot - internal update server name. I.e. 'flashplayer-update.domain.local'.
            -FPDownloadRoot - macromedia update server (for public release only).
            -XMLMajor - name of the xml containing FP major version (may be changed by Adobe).
            -ESR - check ESR version. Needs Adobe auth key you receive as a download link: https://www.adobe.com/ru/products/flashplayer/distribution4.html?auth=<auth-code>.

        Webrequest params:
            -Proxy - proxy server address if needed (see example below).
            -ProxyCreds - username for proxy (if needed). Script will ask for password at start. If proxy specified and no ProxyCreds will be used "ProxyUseDefaultCredentials" flag.
            -UserAgent - uses custom useragent for webrequest: 'InternetExplorer', 'FireFox', 'Chrome', 'Opera', 'Safari' agents are available. Uses built in PowerShell useragents.
            -Force - forces webrequest to ignore sertificate checks.

        Mail params:
            -MailTo - recipient(s) mail.
            -MailFrom - sender mail.
            -SmtpServer - smtp server address.

        Misc:
            -LogFile (string) - logfile path. Default value: "$env:TEMP\FP.CheckUpdates.log".


    Version History:
        20160401:
            - MailTo and MailFrom type changed from string to mailaddress.
            - ParameterSets added for e-mail notification usage control: both MailTo and MailFrom are mandatory if one of them used.
            - UserAgent parameter now uses [ValidateSet('InternetExplorer', 'FireFox', 'Chrome', 'Opera', 'Safari')] which sets preconfigured PowerShell useragents.
            - ESR parameter now requires auth code (https://www.adobe.com/ru/products/flashplayer/distribution4.html?auth=<AUTHCODE>)/
            - If webrequest fails by timeout it will retry 10 times.
            - Some breaks replaced by exits. Otherwise script may exit FOR cycle and not the script.
            - Added version history to script description.
