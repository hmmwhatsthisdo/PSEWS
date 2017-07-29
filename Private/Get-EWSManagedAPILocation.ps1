Function Get-EWSManagedAPILocation {
[OutputType("String")]
[CmdletBinding()]
Param(

    [ValidateSet(
        "Session",
        "User",
        "Machine"
    )]
    [String]$Scope

)

    Function GetSessionLocation {

        return $Script:APILocation

    }

    Function GetUserLocation {

        if (Test-Path (Join-Path $env:APPDATA "PSEWS\APILocation.clixml")) {
            return Import-Clixml (Join-Path $env:APPDATA "PSEWS\APILocation.clixml")
        }

    }

    Function GetMachineLocation {

        if (Test-Path (Join-Path $env:ProgramData "PSEWS\APILocation.clixml")) {
            return Import-Clixml (Join-Path $env:ProgramData "PSEWS\APILocation.clixml")
        }

    }

    if ($Scope) {
        switch ($Scope) {
            "Session" { return GetSessionLocation; break }
            "User" { return GetUserLocation; break }
            "Machine" { return GetMachineLocation; break }
            Default {}
        }
    }

    return (((GetSessionLocation),(GetUserLocation),(GetMachineLocation)) -ne $null)[0]


}