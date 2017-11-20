Function Restart-PSEWS {
[CmdletBinding()]
Param(

)

    Write-Warning "Reloading module."

    Import-Module (Join-Path $ExecutionContext.SessionState.Module.ModuleBase ($ExecutionContext.SessionState.Module.Name + ".psd1")) -Force -Global

}