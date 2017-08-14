function Remove-EWSMailboxDelegate {
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias(
            "Address",
            "email",
            "emailAddress",
            "primarySmtpAddress",
            "smtpAddress"
        )]
        [Microsoft.Exchange.WebServices.Data.Mailbox]
        $Mailbox,

        # The specific UserIDs to remove delegate access for.
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [ValidateCount(1,[int32]::MaxValue)]
        [Microsoft.Exchange.WebServices.Data.UserID[]]
        $UserID,

        [ValidateNotNull()]
        [PSCustomObject]
        $Profile = $(Get-EWSProfile -Default)
            
    )
        
    begin {

        Assert-EWSProfileConnected $Profile

    }

    process {

        $ExchSvc = [Microsoft.Exchange.WebServices.Data.ExchangeService]($Profile.ExchangeService)

        $Response = $ExchSvc.RemoveDelegates($Mailbox, $UserID)

        $Response | ForEach-Object {

            $_Response = $_

            $_UserID = $UserID[$Response.IndexOf($_Response)]

            switch ($_Response.Result) {
                "Success" {
                    Write-Verbose "Removed delegate access to $($Mailbox.ToString()) for $($_UserId.PrimarySmtpAddress) successfully."
                 }
                "Error" {
                    Write-Error "Unable to add remove delegate access to $Mailbox for user $($_UserId.PrimarySmtpAddress): $($_Response.ErrorMessage)" -ErrorId $_Response.ErrorCode -TargetObject $_Response.DelegateUser
                }
            }


        }

    }

    end {
    }
}