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

# Export our public functions, now that they exist
$Scripts.Public | ForEach-Object BaseName | Export-ModuleMember 

