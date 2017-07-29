function Get-EWSProfile {
[OutputType("PSEWS.Profile[]")]
[CmdletBinding(
    DefaultParameterSetName = "All"
)]
Param(
    [Parameter(
        Position = 0,
        ParameterSetName = "ByProfileName"
    )]
    [String]$ProfileName,

    [Parameter(
        Position = 0,
        ParameterSetName = "ByServer"
    )]
    [String]$Server

)

    $Profiles = $Script:EWSProfiles.GetEnumerator()

    if ($PSCmdlet.ParameterSetName -eq "ByProfileName") {
        return $Profiles | Where-Object Name -like $ProfileName
    } elseif ($PSCmdlet.ParameterSetName -eq "ByServer") {
        return $Profiles | Where-Object Server -like $Server
    } else {
        return $Profiles
    }

}