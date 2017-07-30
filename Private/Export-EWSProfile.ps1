function Export-EWSProfile {
Param(
    # The profile to export
    [Parameter(Mandatory = $true, Position = 0)]
    [PSCustomObject]
    $Profile,

    # The GUID of the profile, so it can be used as an invariant.
    [guid]$Guid

)
    # Create a directory to drop the profile CLIXMLs into, if one doesn't already exist
    If (-not (Test-Path $env:APPDATA\PSEWS\Profiles -PathType Container)) {

        New-Item -Path $env:APPDATA\PSEWS -Name Profiles -ItemType Directory | Out-Null

    }

    $Profile | Export-Clixml -Path $env:APPDATA\PSEWS\Profiles\$Guid.clixml
}