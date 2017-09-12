function Remove-EWSFolderPermission {
    [CmdletBinding()]
    param (
        # The folder to remove permissions for.
        [Alias(
            "Id"
        )]
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Microsoft.Exchange.WebServices.Data.FolderId]$FolderId,

        # The specific UserIDs to remove permissions for
        [Parameter(
            Mandatory = $true
        )]
        [ValidateCount(1,[int32]::MaxValue)]
        [Microsoft.Exchange.WebServices.Data.UserID[]]$UserID,

        [ValidateNotNull()]
        [PSCustomObject]$Profile = $(Get-EWSProfile -Default)
    )
    
    begin {

        Assert-EWSProfileConnected -Profile $Profile

    }
    
    process {

        # Grab a copy of the folder we're working on.
        $Folder = Get-EWSMailboxFolder -FolderId $FolderId -Profile $Profile -Properties Permissions,DisplayName

        Write-Verbose "Processing permission removal for folder $(If ($Folder.DisplayName -ne $null) {'"' + $($Folder.DisplayName) + '" '})with ID $FolderId"

        [System.Collections.ArrayList]$PermissionsToRemove = @()

        # Iterate over the users that (might) need to be removed from the permission manifest.
        $UserID | ForEach-Object {

            $_UserID = $_

            Write-Verbose "Processing $($_UserId.ToString())."
            
            # Iterate over all permission entries that loosely match our set of userIDs.
            $Folder.Permissions | Where-Object {Test-EWSUserIdMatch $_UserID $_.UserId} | ForEach-Object {

                Write-Verbose "Removing permissions for $($_.UserId.ToString())"
                $PermissionsToRemove.Add($_) | Out-Null

            }

        }

        # Now that we've finished iterating, remove the permissions from the manifest.
        $PermissionsToRemove | ForEach-Object {

            $Folder.Permissions.Remove($_) | Out-Null

        }

        # Finally, write back to EWS.
        Write-Verbose "Sending changes to EWS..."

        $Folder.Update()

        Write-Verbose "Processing complete."

    }
    
    end {
    }
}