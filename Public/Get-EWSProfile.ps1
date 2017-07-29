function Get-EWSProfile {
[OutputType("PSEWS.Profile[]")]
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
        Position = 0,
        ParameterSetName = "ByGUID",
        ValueFromPipelineByPropertyName = $true
    )]
    [Alias(
        "GUID",
        "ID"
    )]
    [GUID]$ProfileGUID
)

    $Profiles = $Script:EWSProfiles.GetEnumerator()

    if ($PSCmdlet.ParameterSetName -eq "ByProfileName") {
        return $Profiles | ForEach-Object Value | Where-Object Name -like $ProfileName
    } elseif ($PSCmdlet.ParameterSetName -eq "ByServer") {
        return $Profiles | ForEach-Object Value | Where-Object Server -like $Server
    } elseif ($PSCmdlet.ParameterSetName -eq "ByGUID") {
        return $Script:EWSProfiles[$ProfileGUID.ToString()]
    } else {
        return $Profiles
    }

}