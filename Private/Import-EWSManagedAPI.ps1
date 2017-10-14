Function Import-EWSManagedAPI {
[CmdletBinding()]
[OutputType("void")]
# Rule is bugged, see https://github.com/PowerShell/PSScriptAnalyzer/issues/636
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]
Param(
    # Specifies a path to one or more locations.
    [Parameter(Mandatory=$false,
               Position=0,
               ValueFromPipeline=$true,
               ValueFromPipelineByPropertyName=$true,
               HelpMessage="Path to one or more locations where the EWS Managed API can be found. Specify either a directory containing the EWS Managed API DLLs or a direct path to Microsoft.Exchange.WebServices.dll.")]
    [Alias(
        "PSPath",
        "Path"
    )]
    [ValidateNotNullOrEmpty()]
    [string[]]
    $ImportPath
)

    $SuccessfulImport = $false


    # Is the EWS API already available?
    If (Get-Module Microsoft.Exchange.WebServices) {

        $Module = Get-Module Microsoft.Exchange.WebServices

        Write-Warning "EWS Managed API with DLL version $($module.Version) already imported. Reloading CLR assemblies without restarting PowerShell is currently unsupported."

        Return

    }

    # Was a path specified?
    If ($ImportPath) {

        $ImportPath | ForEach-Object {

            # Were we handed a directory?
            If (Test-Path $_ -PathType Container) {
                
                $DLLPath = Join-Path $_ "Microsoft.Exchange.WebServices.dll"
                
            # Were we handed a file?
            } ElseIf (Test-Path $_ -PathType Leaf) {
                
               $DLLPath = $_

            }

            try {

                Write-Verbose "Attempting to import EWS Managed API from $DLLPath..."
                $Module = Import-Module $DLLPath -ErrorAction Stop -PassThru
                $SuccessfulImport = $true
                $Script:ReferencedAssemblies += $DLLPath
                Write-Verbose "Imported EWS Managed API (DLL Version $($Module.Version)) successfully."
                Return

            }
            catch {

                Write-Warning "Import of EWS Managed API from $DLLPath failed: $_"
            
            }

        }

    }

    # Bail out if we already imported.
    If ($SuccessfulImport) {Return}

    # Is it [somehow] in the module path?
    If (Get-Module Microsoft.Exchange.WebServices -ListAvailable) {

        try {

            Write-Verbose "Attempting automatic import of EWS Managed API..."
            $Module = Import-Module Microsoft.Exchange.WebServices -ErrorAction Stop -PassThru
            $SuccessfulImport = $true
            $Script:ReferencedAssemblies += $Module.Path
            Write-Verbose "Imported EWS Managed API (DLL Version $($Module.Version)) successfully."    

        }
        catch {

            Write-Warning "Automatic import of EWS Managed API failed: $_."
            
        }

        
        

    } 

    # Is it installed outside of NuGet?
    If ( -not $SuccessfulImport -and (Test-Path "$env:ProgramFiles\Microsoft\Exchange\Web Services\")) {

        Write-Verbose "Searching for EWS Managed API Versions..."

        Get-ChildItem -Directory "$env:ProgramFiles\Microsoft\Exchange\Web Services\" | Sort-Object -Property @{Expression = {$_.BaseName -as [version]}; Descending = $true} | ForEach-Object {

            $EWSApiVersion = $_.BaseName
            
            try {
                Write-Verbose "Attempting to import EWS Managed API version $EWSApiVersion from $($_.Fullname)..."
                $Module = Import-Module "$($_.FullName)\Microsoft.Exchange.WebServices.dll" -ErrorAction Stop -PassThru
                $SuccessfulImport = $true 
                $Script:ReferencedAssemblies += "$($_.FullName)\Microsoft.Exchange.WebServices.dll"
                Write-Verbose "Imported EWS Managed API version $EWSApiVersion (DLL Version $($Module.Version)) successfully."
                Return

            } catch {

                Write-Warning "Import of EWS Managed API version $EWSApiVersion failed: $_."
            
            }

        }
    
    } 

    # Has it been installed via NuGet?
    If (-not $SuccessfulImport -and (Get-Package Microsoft.Exchange.WebServices -ProviderName NuGet -ErrorAction SilentlyContinue)){

        Write-Verbose "Found NuGet package entry for EWS Managed API."

        $EWSPackage = Get-Package Microsoft.Exchange.WebServices -ProviderName NuGet

        $EWSPath = Split-path $EWSPackage.Source -Parent

        try {

            Write-Verbose "Attempting to import EWS Managed API from NuGet package..."
            $Module = Import-Module (Join-Path $EWSPath "lib\40\Microsoft.Exchange.WebServices.dll") -PassThru
            $SuccessfulImport = $true                
            $Script:ReferencedAssemblies += (Join-Path $EWSPath "lib\40\Microsoft.Exchange.WebServices.dll")
            Write-Verbose "Imported EWS Managed API version $($EWSPackage.Version) (DLL Version $($Module.Version)) successfully."
            
        } catch {

            Write-Warning "Import of EWS Managed API version $($EWSPackage.Version) via NuGet package failed: $_"
            
        }
        
    } 
    
    If (-not $SuccessfulImport) {
    
        throw [System.IO.FileNotFoundException]("EWS Managed API not found or failed to import. Run `"Install-Package Microsoft.Exchange.WebServices`" as local administrator to install via NuGet, or specify a location using the -ImportPath parameter.")
    
    }
    


}