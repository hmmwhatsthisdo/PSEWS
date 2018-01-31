function Get-EWSChildFolder {
    [OutputType("Microsoft.Exchange.WebServices.Data.FindFolderResults")]
    [CmdletBinding(
        SupportsPaging=$true
    )]
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

        $View.PropertySet = [Microsoft.Exchange.WebServices.Data.PropertySet]::new([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName)

        If ($PSCmdlet.PagingParameters) {

            $View.Offset = $(If ($PSCmdlet.PagingParameters.Skip -gt [Int32]::MaxValue) {[int32]::MaxValue} Else {$PSCmdlet.PagingParameters.Skip})
            $View.PageSize = $(If ($PSCmdlet.PagingParameters.First -gt [Int32]::MaxValue) {[int32]::MaxValue} Else {$PSCmdlet.PagingParameters.First})

        }
        If ($Recurse) {

            $View.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep

        } Else {

            $View.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow

        }        

        [Microsoft.Exchange.WebServices.Data.FindFoldersResults]$Results = $ExchSvc.FindFolders($RootFolderId, $View)

        If ($PSCmdlet.PagingParameters.IncludeTotalCount) {

            $PSCmdlet.PagingParameters.NewTotalCount($Results.TotalCount, 1.0)

        }

        Write-Output $Results

    }
    
    end {
    }
}