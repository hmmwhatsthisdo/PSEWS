# PSEWS
PowerShell module for interacting with Exchange Web Services

## Features
* Auto
* Delegate management
    * Get, Add, Set, Remove
* Folder interaction
    * Permission management
        * Get, Add, Set, Remove
    * Limited Property retrieval

## Prerequisites
* Microsoft Exchange Web Services (EWS) Managed API
    * Available via the following:
        * [GitHub: OfficeDev/ews-managed-api](https://github.com/OfficeDev/ews-managed-api)
            * Clone, build, and provide path via `Set-EWSManagedAPILocation` (if needed)            
        * [NuGet: Microsoft.Exchange.WebServices](https://www.nuget.org/packages/Microsoft.Exchange.WebServices/)
            * `Install-Package Microsoft.Exchange.WebServices -MinVersion 2.2.0` on PoSH 3.0 and above
    * PSEWS will import into fallback mode and provide options for installing the EWS Managed API via NuGet (if available). (Internet access required)
    * PowerShell 5.0 or above
        * 3.0/4.0 _may_ work, but this has not yet been tested.

## Installation/Usage
1. Clone this repository to a folder on-disk.
    * To allow for importing without specifying a path (i.e. `Import-Module PSEWS`), clone the module to a folder in `$Env:PSModulePath` (e.g. `$HOME\Documents\Windows\PowerShell\Modules\PSEWS`) 
2. Import the module.
3. If the EWS Managed API was not found on-disk, PSEWS will import into fallback mode and export two functions:
    * `Set-EWSManagedAPILocation`: Manually specify the path to the EWS Managed API DLLs
    * `Install-EWSManagedAPI`: Attempt to install the EWS Managed API via NuGet

    Once the EWS Managed API is found, PSEWS will attempt a re-import.
4. Once PSEWS has imported, use `Add-EWSProfile` to configure a profile for connecting to Exchange Web Services.
5. After importing the PSEWS module, use `Connect-EWS` to invoke Autodiscover (if needed) and create a connection to Exchange Web Services.

