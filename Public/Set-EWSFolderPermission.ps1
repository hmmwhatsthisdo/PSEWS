function Set-EWSFolderPermission {
    [CmdletBinding()]
    param (
        # The folder to set permissions on.
        [Alias(
            "Id"
        )]
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Microsoft.Exchange.WebServices.Data.FolderId]$FolderId,

        # The specific UserIDs to set permissions for.
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
        [Microsoft.Exchange.WebServices.Data.FolderPermissionLevel]$PermissionLevel,

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

        Write-Verbose "Processing folder permissions for folder with ID $($FolderId.UniqueId)."

        # Get the folder object, with permissions attached.
        $Folder = Get-EWSMailboxFolder -FolderId $FolderId -Profile $Profile -Properties Permissions

        # Iterate over the users we need to set permissions for.
        $EntriesToProcess = $UserID | ForEach-Object {

            # Keep track of the user we're iterating over
            $_UserID = $_

            Write-Verbose "Processing $($_UserID.ToString())..."

            # Do we actually have a matching userID to edit?
            If (($Folder.Permissions | Where-Object {Test-EWSUserIDMatch $_UserID $_.UserId} | Tee-Object -Variable "MatchingEntries")) {

                Write-Verbose "Found matching item(s) for $($_UserID.ToString()) in permissions manifest."

                # Momentarily remove the permission entries from the manifest (I'm not sure if it's by reference, so I'd rather err on the side of caution)
                $MatchingEntries | ForEach-Object {

                    Write-Output $_
                    $Folder.Permissions.Remove($_) | Out-Null

                }

            # Were we told to do otherwise?
            } Elseif ($Force) {
                Write-Verbose "-Force specified, adding permission entry for $($_UserID.ToString())."

                $NewPermissionObject = [Microsoft.Exchange.WebServices.Data.FolderPermission]::new()

                $NewPermissionObject.UserId = $_UserID

                Write-Output $NewPermissionObject

            # We shouldn't be adding permissions for this userID, bail
            } Else {

                Write-Error "Failed to set permissions for $($_UserID.ToString()) as there are no preexisting permissions for the user. Use Add-EWSFolderPermission or the -Force parameter to resolve this."

            }



        }

        # Imbue the permission entries with the new permissions
        $EntriesToProcess | ForEach-Object {

            $_EntriesToProcess = $_

            Write-Verbose "Setting permissions for $($_EntriesToProcess.UserId.ToString())"

            @(
                "PermissionLevel"
                "CanCreateItems"
                "CanCreateSubFolders"
                "IsFolderOwner"
                "IsFolderContact"
                "EditItems"
                "DeleteItems"
                "ReadItems"
            ) | ForEach-Object {

                Write-Debug "Checking variable $_"

                If ((Get-Variable -Name $_ -ValueOnly -ErrorAction SilentlyContinue) -ne $null) {

                    Write-Debug "Setting property $_."
                    $_EntriesToProcess.$_ = (Get-Variable -Name $_ -ValueOnly)

                }

            }

        }

        # Now that changes have been made, add the (new and improved?) entries to the permissions manifest
        $EntriesToProcess | ForEach-Object {

            Write-Verbose "Reregistering permission entry for $($_.UserId.ToString()) in folder manifest."
            $Folder.Permissions.Add($_) | Out-Null

        }

        # Finally, push changes back to EWS.
        Write-Verbose "Propagating changes to EWS..."
        $Folder.Update()

        Write-Verbose "Finished processing permission changes."

    }
    
    end {
    }
}