function Import-EWSProfile {
[CmdletBinding()]
param (
    
)
    
If (Test-Path $env:APPDATA\PSEWS\Profiles) {

    $Profiles = Get-ChildItem $env:APPDATA\PSEWS\Profiles\*.clixml

    $Profiles | ForEach-Object {

        $Script:Profiles[$_.BaseName] = Import-Clixml $_

    }

    $DiskDefaultGUID = Import-Clixml $env:APPDATA\PSEWS\DefaultProfileGUID.clixml

    if ($Script:Profiles[$DiskDefaultGUID.ToString()]) {

        $Script:DefaultProfileGUID = $DiskDefaultGUID

    }

}

}