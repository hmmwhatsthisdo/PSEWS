function Set-EWSManagedAPILocation {
[CmdletBinding()]
Param (

    # Specifies a path to one or more locations. Unlike the Path parameter, the value of the LiteralPath parameter is
    # used exactly as it is typed. No characters are interpreted as wildcards. If the path includes escape characters,
    # enclose it in single quotation marks. Single quotation marks tell Windows PowerShell not to interpret any
    # characters as escape sequences.
    [Parameter(
        Mandatory=$true,
        Position=0,
        ValueFromPipelineByPropertyName=$true,
        HelpMessage="Literal path to the EWS Managed API DLLs. Specify the containing folder or direct path to Microsoft.Exchange.WebServices.dll."
    )]
    [Alias("PSPath")]
    [ValidateNotNullOrEmpty()]
    [string[]]
    $LiteralPath,

    [ValidateSet(
        "Session",
        "User",
        "Machine"
    )]
    [String]$Scope

)

    Function SetSessionLocation ($LiteralPath) {
        $Script:APILocation = $LiteralPath
    }

    Function SetUserLocation ($LiteralPath) {
        New-Item -Path $env:APPDATA -Name PSEWS -Force -ItemType Directory | Out-Null
        $LiteralPath | Export-Clixml $env:APPDATA\PSEWS\APILocation.clixml        
    }

    Function SetMachineLocation ($LiteralPath) {
        New-Item -Path $env:ProgramData -Name PSEWS -Force -ItemType Directory | Out-Null
        $LiteralPath | Export-Clixml $env:ProgramData\PSEWS\APILocation.clixml         
    }

    switch ($Scope) {
        "Session" { SetSessionLocation $LiteralPath; break }
        "User" { SetUserLocation $LiteralPath; break }
        "Machine" { SetMachineLocation $LiteralPath; break }
        Default {
            try {
                SetMachineLocation $LiteralPath
            }
            catch {
                SetUserLocation $LiteralPath
            }
        }
    }

    if ($Script:FallbackMode) {
        Restart-PSEWS        
    }

}