function Find-EWSFolder {
    [OutputType("Microsoft.Exchange.WebServices.Data.Folder[]")]
    [CmdletBinding()]
    param (
        # The base folder to begin searching from.
        [Alias(
            "RootId",
            "Root",
            "Id",
            "FolderId"
        )]
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0
        )]
        [ValidateNotNull()]
        [Microsoft.Exchange.WebServices.Data.FolderId]
        $RootFolderId = [Microsoft.Exchange.WebServices.Data.FolderId]::new([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root),

        # Search child folders as well as the children themselves.
        [Switch]$Recurse,

        # The profile to use.
        [ValidateNotNull()]
        [PSCustomObject]
        $Profile = $(Get-EWSProfile -Default)
    )
    
    begin {

        Assert-EWSProfileConnected $Profile

        [Microsoft.Exchange.WebServices.Data.ExchangeService]$ExchSvc = $Profile.ExchangeService
    }
    
    process {

        $View = [Microsoft.Exchange.WebServices.Data.FolderView]::new([Int32]::MaxValue)

        If ($Recurse) {

            $View.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep

        } Else {

            $View.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow

        }

        Write-Output $ExchSvc.FindFolders($RootFolderId, $View)

    }
    
    end {
    }
}