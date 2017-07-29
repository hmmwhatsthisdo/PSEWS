function Add-EWSProfile {
[CmdletBinding(
    DefaultParameterSetName = "ByAddress"
)]
param (
    # The address and credential to use for logging onto Exchange.
    [Parameter(
        ParameterSetName = "ByAddress",
        Mandatory = $true,
        Position = 0,
        ValueFromPipelineByPropertyName = $true
    )]
    [Alias(
        "emailAddress",
        "email",
        "primarySmtpAddress"
    )]
    [ValidateScript({
        [system.net.mail.mailaddress]::new($_.Username) | Out-Null
    })]
    [Credential()]
    [PSCredential]$Credential,

    [Switch]$SetAsDefault,

    [ValidateSet(
        "Session",
        "User"
        # "Machine" # Prerequisite: learning how CMS works
    )]
    [String]$Scope = "User"

)
    
process {
    $GUID = New-Guid

    $Script:Profiles[$GUID.ToString()] = @{

        Credential = $Credential
        Server = $null
        Scope = $Scope
    }

    # Force default if there's only one profile loaded
    if ($Script:Profiles.Count -eq 1) {

        $SetAsDefault = $true

    }

    if ($SetAsDefault) {

        $Script:DefaultProfileGUID = $GUID

    }

    if ($Scope -eq "User") {

        Export-EWSProfile $Profiles[$GUID.ToString()] -Guid $GUID

        if ($SetAsDefault) {

            $GUID.ToString() | Export-Clixml $Env:APPDATA\PSEWS\DefaultProfileGUID.clixml
        
        }

    }

    return $Script:Profiles[$GUID.ToString()]
}
    

}