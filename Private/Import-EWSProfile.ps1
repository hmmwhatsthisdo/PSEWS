function Import-EWSProfile {
[CmdletBinding()]
param (
    
)
    
If (Test-Path $env:APPDATA\PSEWS\Profiles) {

    $Profiles = Get-ChildItem $env:APPDATA\PSEWS\Profiles\*.clixml

    $Profiles | ForEach-Object {

        $Script:EWSProfiles[$_.BaseName] = Import-Clixml $_

    }

    if (Test-path $env:APPDATA\PSEWS\DefaultProfileGUID.clixml) {
        
        $DiskDefaultGUID = [GUID](Import-Clixml $env:APPDATA\PSEWS\DefaultProfileGUID.clixml)

        if ($Script:EWSProfiles[$DiskDefaultGUID.ToString()]) {

            $Script:DefaultProfileGUID = $DiskDefaultGUID

        }

    }


    

}

}