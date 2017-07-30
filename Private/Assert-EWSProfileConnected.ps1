function Assert-EWSProfileConnected {
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory = $true,
            Position = 0
        )]
        [ValidateNotNull()]
        [pscustomobject]$Profile
    )
      
    process {
        if ($Profile.ExchangeService -eq $Null) {
            throw [InvalidOperationException]("Profile $($Profile.Credential.UserName)$(If ($Profile.Server) {" on server $($Profile.Server)"} ) has not been connected to EWS. Use the Connect-EWS cmdlet to connect the account before continuing.")
        }
    }
    
}