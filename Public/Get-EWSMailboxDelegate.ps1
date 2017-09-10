function Get-EWSMailboxDelegate {
    [OutputType("Microsoft.Exchange.WebServices.Data.DelegateUser")]
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
        [Microsoft.Exchange.WebServices.Data.Mailbox]$Mailbox,

        # The specific UserIDs to retrieve permissions for, if desired.
        [ValidateCount(1,[int32]::MaxValue)]
        [Microsoft.Exchange.WebServices.Data.UserID[]]$UserID,

        [ValidateNotNull()]
        [PSCustomObject]$Profile = $(Get-EWSProfile -Default)
    )
    
    begin {
        Assert-EWSProfileConnected $Profile
    }
    
    process {

        # Shim to allow intellisense to know what's going on until we implement a strongly-typed container class for profiles
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$ExchSvc = $Profile.ExchangeService 

        if ($UserID) {
            $DelegateResponse = $ExchSvc.GetDelegates($Mailbox, $true, $UserID)
        } else {
            $DelegateResponse = $ExchSvc.GetDelegates($Mailbox, $true)
        }
        
        $DelegateResponse.DelegateUserResponses | ForEach-Object {

            $_DelegateUserResponse = $_

            switch ($_.Result) {
                "Error" { 

                    Write-Error "Failed to obtain delegate for mailbox $Mailbox`: $($_DelegateUserResponse.ErrorMessage)" -ErrorId $_DelegateUserResponse.ErrorCode

                }
                "Success" {

                    $_DelegateUserResponse.DelegateUser | Add-Member -MemberType NoteProperty -Name Mailbox -Value $Mailbox
                    
                    Write-Output $_DelegateUserResponse.DelegateUser

                }
            }

        }


    }
    
}