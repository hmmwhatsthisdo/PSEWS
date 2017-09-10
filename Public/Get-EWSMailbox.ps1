# Stub function until I can determine a better way to resolve mailboxes 
function Get-EWSMailbox {
[OutputType("Microsoft.Exchange.WebServices.Data.Mailbox")]
[CmdletBinding()]
param (
    [Parameter(
        Position = 0,
        ValueFromPipeline = $true
    )]
    [Alias(
        "Address",
        "email",
        "emailAddress"
    )]
    [ValidateScript({
        $_ | Foreach-Object {[system.net.mail.mailaddress]::new($_)}
    })]
    [String[]]$SMTPAddress
)
    
    process {

        $SMTPAddress | ForEach-Object {
            Write-Output ([Microsoft.Exchange.WebServices.Data.Mailbox]::new($_))
        }

    }
}