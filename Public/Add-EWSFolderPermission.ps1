function Add-EWSFolderPermission {
    [CmdletBinding()]
    param (
        # The folder to add permissions for.
        [Alias(
            "Id"
        )]
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Microsoft.Exchange.WebServices.Data.FolderId]$FolderId,

        # The specific UserIDs to add permissions for.
        [ValidateCount(1,[int32]::MaxValue)]
        [Parameter(
            Mandatory = $true
        )]
        [Microsoft.Exchange.WebServices.Data.UserID[]]$UserID,

        [Alias(
            "Permission",
            "Level",
            "FolderPermissionLevel"
        )]
        [ValidateScript({
            if ($_ -eq [Microsoft.Exchange.WebServices.Data.FolderPermissionLevel]::Custom) {

                throw [System.ArgumentException]::new("Cannot specify a base permission level of Custom. Specify a different base permission level (if needed), then use other parameters to customize the specific permission attributes.", "FolderPermissionLevel")

            } else {
                return $true
            }
        })]
        [Microsoft.Exchange.WebServices.Data.FolderPermissionLevel]$PermissionLevel = [Microsoft.Exchange.WebServices.Data.FolderPermissionLevel]::None,

        [Alias(
            "CreateItems"
        )]
        [Switch]$CanCreateItems,

        [Alias(
            "CreateSubFolders",
            "SubFolders"
        )]
        [Switch]$CanCreateSubFolders,

        [Alias(
            "FolderOwner",
            "Owner"
        )]
        [Switch]$IsFolderOwner,

        [Alias(
            "FolderContact",
            "Contact"
        )]
        [Switch]$IsFolderContact,

        [Alias(
            "Edit"
        )]
        [Microsoft.Exchange.WebServices.Data.PermissionScope]$EditItems,

        [Alias(
            "Delete"
        )]
        [Microsoft.Exchange.WebServices.Data.PermissionScope]$DeleteItems,

        [Alias(
            "Read"
        )]
        [Microsoft.Exchange.WebServices.Data.FolderPermissionReadAccess]$ReadItems,

        [Switch]$Force,

        [ValidateNotNull()]
        [PSCustomObject]$Profile = $(Get-EWSProfile -Default)
    )
    
    begin {

        Assert-EWSProfileConnected -Profile $Profile

    }
    
    process {

        # First, get the most up to date copy of the folder.
        Write-Verbose "Processing permissions for folder with ID $($FolderId.UniqueId)"
        [Microsoft.Exchange.WebServices.Data.Folder]$Folder = Get-EWSMailboxFolder -FolderId $FolderId -Profile $Profile -Properties Permissions

        $PreexistingUserIDs = $Folder.Permissions | ForEach-Object UserId
        Write-Verbose "$($PreexistingUserIDs | Measure-Object | ForEach-Object Count) permission entries already exist on folder."
        
        # Now, iterate over the users specified.
        $UserID | ForEach-Object {

            # Hold onto that value.
            $_UserId = $_

            Write-Verbose "Processing $($_UserID.ToString())."

            # Iterate over the userIDs that already exist to ensure we're not going to add duplicate permissions.
            $ShouldContinue = $PreexistingUserIDs | ForEach-Object -ErrorAction Stop {

                # This comparison might do well in an ETS-implemented method on [UserID]
                If (Test-EWSUserIdMatch $_UserId $_) {

                    If (-not $Force) {
                        
                        Write-Error "User $($_UserId.ToString()) already has permissions for this folder. Use Set-EWSFolderPermission or the -Force parameter to [re]configure permissions."

                        return $false
    
                    } Else {

                        Write-Verbose "-Force specified, removing preexisting permissions."
    
                        $Folder.Permissions.Remove(($Folder.Permissions | Where-Object UserID -EQ $_))

                        return $true
    
                    }

                }

            }
            
            # At this point, $ShouldContinue is either false, true, or null - we should only bail if it's false.

            if ($ShouldContinue -eq $false) {

                return

            }

            # Now, build the permission entry.

            $NewPermissionEntry = [Microsoft.Exchange.WebServices.Data.FolderPermission]::new($_UserId, $PermissionLevel)

            # Apply any additional parameters passed in via arguments.
            @(
                "CanCreateItems"
                "CanCreateSubFolders"
                "IsFolderOwner"
                "IsFolderContact"
                "EditItems"
                "DeleteItems"
                "ReadItems"
            ) | ForEach-Object {

                # Is the variable defined?
                if (Get-Variable $_ -ValueOnly -ErrorAction SilentlyContinue) {

                    $NewPermissionEntry.$_ = (Get-Variable $_ -ValueOnly)

                }

            }

            # Finally, add the freshly-minted permission entry to the permissions manifest for the folder.

            $Folder.Permissions.Add($NewPermissionEntry)
        }

        # Last, .Update() the folder to push the permission changes back to EWS.

        Write-Verbose "Saving changes to EWS..."
        
        $Folder.Update()

    }
    
    end {
    }
}