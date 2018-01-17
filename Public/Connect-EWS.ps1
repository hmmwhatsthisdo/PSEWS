Function Connect-EWS {
[CmdletBinding(
    DefaultParameterSetName = "AutoDiscover"
)]
Param (
    # The profile to connect with.
    [Parameter(
        ValueFromPipeline = $true
    )]
    [ValidateCount(1, ([Int32]::MaxValue))]
    [PSCustomObject[]]$Profile = $(Get-EWSProfile -Default),

    [Parameter(
        Mandatory = $true,
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

        Write-Verbose "Processing $Account."

        $ExchSvc = [Microsoft.Exchange.WebServices.Data.ExchangeService]::new()

        $ExchSvc.UseDefaultCredentials = $false

        $ExchSvc.Credentials = [Microsoft.Exchange.WebServices.Data.ExchangeCredentials]($_Profile.Credential.GetNetworkCredential())

        If ($PSCmdlet.ParameterSetName -eq "AutoDiscover") {
            try {
                Write-Verbose "Invoking Autodiscover for $Account..."
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
        
            Write-Verbose "EWS URI $EndpointURI specified, skipping autodiscover"
            $ExchSvc.Url = $EndpointURI

        }

        if ($ExchSvc.Url) {

            try {

                Write-Verbose "Attempting to bind to root folder..."

                # Try to bind to the root folder of our account, just to make sure we were able to form a connection of some sort
                [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ExchSvc, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root) | Out-Null

                Write-Verbose "$Account connected to $($ExchSvc.Url)"
                $_Profile.Server = ([URI]$ExchSvc.Url).Host
                $_Profile.ExchangeService = $ExchSvc
                
            } catch {
                Write-Error "Failed to validate connection to Exchange. Verify the credentials are correct and try again."
            }
            
        }


    }

}

}