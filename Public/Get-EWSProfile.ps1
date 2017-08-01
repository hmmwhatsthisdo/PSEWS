function Get-EWSProfile {
# [OutputType("PSEWS.Profile[]")]
[CmdletBinding(
    DefaultParameterSetName = "All"
)]
Param(
    [Parameter(
        Position = 0,
        ParameterSetName = "ByProfileName",
        ValueFromPipeline = $true
    )]
    [Alias(
        "emailAddress",
        "email",
        "primarySmtpAddress"
    )]
    [String]$ProfileName,

    [Parameter(
        Position = 0,
        ParameterSetName = "ByServer"
    )]
    [String]$Server,

    # The GUID referencing a profile.
    [Parameter(
        Mandatory = $true,
        Position = 0,
        ParameterSetName = "ByGUID",
        ValueFromPipelineByPropertyName = $true
    )]
    [Alias(
        "GUID",
        "ID"
    )]
    [GUID]$ProfileGUID,

    # Get the default profile.
    [Parameter(
        Mandatory = $true,
        ParameterSetName = "DefaultOnly"
    )]
    [Switch]$Default
)

    if ($PSCmdlet.ParameterSetName -eq "DefaultOnly") {
        return Get-EWSProfile -ProfileGUID $Script:DefaultProfileGUID
    } 

    return $Script:EWSProfiles.GetEnumerator() | ForEach-Object Value | ForEach-Object {

        if ($PSCmdlet.ParameterSetName -eq "ByProfileName") {
            $_ | Where-Object {$_.Credential.Username -like $ProfileName}
        } elseif ($PSCmdlet.ParameterSetName -eq "ByServer") {
            $_ | Where-Object Server -like $Server
        } elseif ($PSCmdlet.ParameterSetName -eq "ByGUID") {
            $_ | Where-Object Guid -eq $ProfileGUID
        } else {
            $_
        } 
    } | ForEach-Object {
        $_.PSObject.Copy() # This may not be necessary... not sure
    }

    

}