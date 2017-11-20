Function Install-EWSManagedAPI {
[CmdletBinding()]
param (
    [ValidateSet(
        "User",
        "Machine"
    )]
    [String]$Scope,

    [String]$Source = "https://nuget.org/api/v2/"
)

    

    switch ($Scope) {
        "User" { Install-Package -Name Microsoft.Exchange.WebServices -ProviderName NuGet -Source $Source -ForceBootstrap -Scope CurrentUser ; break }
        "Machine" {Install-Package -Name Microsoft.Exchange.WebServices -ProviderName NuGet -Source $Source -ForceBootstrap -Scope AllUsers; break }
        Default {
            
            try {
                Install-Package -Name Microsoft.Exchange.WebServices -ProviderName NuGet -Source $Source -ForceBootstrap -Scope AllUsers -ErrorAction Stop

            }
            catch {
                Write-Warning "Failed to install EWS Managed API for system-wide usage. Falling back to installing for current user only."
                Install-Package -Name Microsoft.Exchange.WebServices -ProviderName NuGet -Source $Source -ForceBootstrap -Scope CurrentUser
            }

        }
    }

    If ($Script:FallbackMode) {

        Restart-PSEWS -Debug:$Debug

    }

}