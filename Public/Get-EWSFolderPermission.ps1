function Get-EWSFolderPermission {
    [OutputType("Microsoft.Exchange.WebServices.Data.FolderPermission[]")]
    [CmdletBinding()]
    param (
        # The folder to retrieve permissions for.
        [Alias(
            "Id"
        )]
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Microsoft.Exchange.WebServices.Data.FolderId]$FolderId,

        # The specific UserIDs to retrieve permissions for, if desired.
        [ValidateCount(1,[int32]::MaxValue)]
        [Microsoft.Exchange.WebServices.Data.UserID[]]$UserID,

        [ValidateNotNull()]
        [PSCustomObject]$Profile = $(Get-EWSProfile -Default)
    )
    
    begin {

        Assert-EWSProfileConnected -Profile $Profile

    }
    
    process {

        # Shim to allow intellisense to know what's going on until we implement a strongly-typed container class for profiles
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$ExchSvc = $Profile.ExchangeService 

        [Microsoft.Exchange.WebServices.Data.Folder]$Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ExchSvc, $FolderId, [Microsoft.Exchange.WebServices.Data.PropertySet]::new([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.FolderSchema]::Permissions))

        
        if ($UserID.Count -gt 0) {

            # UserID doesn't appear to offer a loose comparison, so we have to do it ourselves.
            return $Folder.Permissions | Where-Object {
                ($_.UserID.PrimarySmtpAddress -in ($UserID | Where-Object PrimarySmtpAddress -NE $null | ForEach-Object PrimarySmtpAddress)) -or
                ($_.UserId.SID -in ($UserID | Where-Object UserID -NE $null | ForEach-Object UserID)) -or
                ($_.UserId.StandardUser -in ($UserID | Where-Object StandardUser -ne $null | ForEach-Object StandardUser))
            }

        } else {

            return $Folder.Permissions

        }

    }
    
    end {



    }
}