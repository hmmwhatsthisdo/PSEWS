# Get all of the pub/priv scripts 
$Scripts = @{
	Public = Get-ChildItem "$PSScriptRoot\Public\*.ps1" -ErrorAction SilentlyContinue
	Private = Get-ChildItem "$PSScriptRoot\Private\*.ps1" -ErrorAction SilentlyContinue
}

# Import them
foreach ($Type in @("Public","Private")) {
	Write-Verbose "Importing $Type functions..."
	foreach ($ImportScript in $Scripts.$Type) {
		Write-Verbose "Importing function $($ImportScript.Basename)..."
		try {
			. $ImportScript.FullName
		} 
		catch {
			Write-Error "Failed to import $type function $($ImportScript.BaseName): $_"
		}
		
	}
}



# We don't have any libraries for the time being, so this isn't a concern
<#
Write-Verbose "Importing Libraries..."
$Assemblies = @{}
foreach ($Library in (Get-Childitem "$PSScriptRoot\Library\*.dll")) {
	Write-Verbose "Importing $($Library.Basename)..."
	try {
		$Assemblies[$Library.BaseName] = Add-Type -Path $Library.Fullname -PassThru
	} catch {
		Write-Error "Failed to import DLL $($Library.BaseName): $_"
	}
}
#>

try {

	$LocationParams = @{}
	$Script:ReferencedAssemblies = @()


	# Do we already have a stored location?
	If (Get-EWSManagedAPILocation) {

		$LocationParams.Path = Get-EWSManagedAPILocation
		Write-Verbose "Stored API path: $($LocationParams.Path)"
	}
	
	# Try importing the API
	Import-EWSManagedAPI @LocationParams

	# We must have survived

	# Import classes, now that the EWS mAPI exists

	$Classes = @{
	
		ps1 = Get-ChildItem "$PSScriptRoot\Class\*.class.ps1" -ErrorAction SilentlyContinue
		cs = Get-ChildItem "$PSScriptRoot\Class\*.cs" -ErrorAction SilentlyContinue
	
	}
	
	# Import classes defined via PoSH
	$Classes.ps1 | ForEach-Object {
	
		$Class = $_
	
		Write-Verbose $Class.FullName
		try {
			Write-Verbose "Importing PowerShell classes from $($Class.Name)..."
			. $Class.FullName		
		}
		catch {
			Write-Error "Failed to import PowerShell class from $($Class.Name): $_"
		}
	
	}

	# Import classes defined via C# (because PoSH classes don't support namespaces... yet?)
	if ($Classes.cs) {

		Write-Verbose "Importing C# classes..."	
		try {
			Add-Type -Path ($Classes.cs | Foreach-Object Fullname) -Verbose -ReferencedAssemblies $Script:ReferencedAssemblies
		}
		catch {
			Write-Error "Failed to import C# classes: $_"
		}

	}
		
	# Export our functions
	$Scripts.Public | ForEach-Object BaseName | Export-ModuleMember

	if (-not $Script:EWSProfiles) {
		$Script:EWSProfiles = @{}
	}

	Import-EWSProfile
	
}
catch [System.IO.IOException] {
	$Script:FallbackMode = $true
	Export-ModuleMember -Function "Install-EWSManagedAPI","Set-EWSManagedAPILocation"
	Write-Error (
		"Failed to import EWS Managed API. Loading PSEWS in fallback mode. Use Set-EWSManagedAPILocation to specify the location of the EWS Managed API, or use Install-EWSManagedAPI to install via NuGet. The module will reload automatically."
	)
}



