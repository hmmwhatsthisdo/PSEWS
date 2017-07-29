Function Import-EWSManagedAPI {
[CmdletBinding()]
[OutputType("void")]
Param(
)


    # Is the EWS API already available?
    If (Get-Module Microsoft.Exchange.WebServices) {

        $Module = Get-Module Microsoft.Exchange.WebServices

        Write-Warning "EWS Managed API version $($module.Version) already imported."

        Return

    }


    If (Get-Module Microsoft.Exchange.WebServices -ListAvailable) {
        Write-Verbose "Attempting automatic import of EWS Managed API..."
        Import-Module Microsoft.Exchange.WebServices -ErrorAction Stop
    } Elseif (Test-Path "$env:ProgramFiles\Microsoft\Exchange\Web Services\") {
        Write-Verbose "Searching for EWS Managed API Versions..."
        Get-ChildItem -Directory "$env:ProgramFiles\Microsoft\Exchange\Web Services\" | Sort-Object -Property @{Expression = {$_.BaseName -as [version]}; Descending = $true} | ForEach-Object {
            $EWSApiVersion = $_.BaseName
            Write-Verbose "Attempting to import EWS Managed API version $EWSApiVersion..."
            $SuccessfulImport = $false
            try {
                Import-Module "$($_.FullName)\Microsoft.Exchange.WebServices.dll" -ErrorAction Stop
                $SuccessfulImport = $true
                Write-Verbose "Imported EWS Managed API version $EWSApiVersion successfully."
                Return
            } catch {
                Write-Warning "Import of EWS Managed API version $_ failed."
            }

        }
    } Else {
        throw "EWS Managed API not found. Run `"Install-Package Microsoft.Exchange.WebServices`" as local administrator to install via NuGet."
    }

}