Function Restart-PSEWS {
[CmdletBinding()]
Param(

)

    Write-Warning "Reloading module."

    Import-Module $ExecutionContext.SessionState.Module.ModuleBase -Force -Global

}