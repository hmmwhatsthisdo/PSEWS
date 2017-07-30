Function Connect-EWS {
[CmdletBinding(
    DefaultParameterSetName = "AutoDiscover"
)]
Param (
    # The profile to connect with.
    [Parameter(
        ValueFromPipeline = $true, 
        Position = 0
    )]
    [ValidateCount(1, ([Int32]::MaxValue))]
    [PSCustomObject[]]$Profile = $(Get-EWSProfile -Default),

    [Parameter(
        Mandatory = $true,
        Position = 1,
        ParameterSetName = "ExplicitURI"
    )]
    [Alias(
        "Uri"
    )]
    [URI]$EndpointURI,

    [ValidateSet(
        "None",
        "Office365",
        "All"
    )]
    [ValidateNotNullOrEmpty()]
    [String]$RedirectPolicy = "None"
)

Process {

    $Profile | ForEach-Object {

        $_Profile = $Script:EWSProfiles[$_.GUID.ToString()]

        $Account = $_Profile.Credential.UserName

        $ExchSvc = [Microsoft.Exchange.WebServices.Data.ExchangeService]::new()

        $ExchSvc.UseDefaultCredentials = $false

        $ExchSvc.Credentials = [Microsoft.Exchange.WebServices.Data.ExchangeCredentials]($_Profile.Credential.GetNetworkCredential())

        If ($PSCmdlet.ParameterSetName -eq "AutoDiscover") {
            try {
                $ExchSvc.AutodiscoverUrl($Account, {
                    Param(
                        [Parameter(
                            Position = 0
                        )]
                        [String]$URL
                    )

                    if ($RedirectPolicy -ne "All") {
                        Write-Verbose "Autodiscover for $Account redirected to $URL"
                    }
                    

                    switch ($RedirectPolicy) {
                        "None" {
                            if ($URL.ToString() -eq "https://autodiscover-s.outlook.com/autodiscover/autodiscover.xml") {
                                Write-Error "Autodiscover for $Account attempted redirect to Office 365, but redirection has been disabled.`nUse `"-RedirectPolicy Office365`" to allow redirection to Office 365." -RecommendedAction "Use `"-RedirectPolicy Office365`" to allow redirection to Office 365."
                            } Else {
                                Write-Error "Autodiscover for $Account attempted to redirect to $URL, but redirection has been disabled.`nUse the -RedirectPolicy parameter to specify a different redirection policy." -RecommendedAction "Use the -RedirectPolicy parameter to specify a different redirection policy."
                            }

                            return $false 
                        }
                        "Office365" {
                            return ($URL.ToString() -eq "https://autodiscover-s.outlook.com/autodiscover/autodiscover.xml")
                        }
                        "All" { 
                            Write-Warning "Autodiscover for $Account redirected to $URL"
                            return $true 
                        }
                    }
                })
            } catch [Microsoft.Exchange.WebServices.Data.AutodiscoverLocalException] {
                Write-Error "Failed to perform EWS autodiscovery on $Account.`nVerify credentials and redirection policy, then try again." -RecommendedAction "Verify credentials and redirection policy, then try again." 
            }
            

        } Else {
        
            $ExchSvc.Url = $EndpointURI

        }

        if ($ExchSvc.Url) {

            Write-Verbose "$Account connected to $($ExchSvc.Url)"
            $_Profile.Server = ([URI]$ExchSvc.Url).Host
            $_Profile.ExchangeService = $ExchSvc
            
        }


    }

}

}