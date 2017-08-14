function Add-EWSMailboxDelegate {
    [CmdletBinding()]
    # Rule is bugged, see https://github.com/PowerShell/PSScriptAnalyzer/issues/636
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')] 
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

        # The specific UserIDs to delegate to.
        [Parameter(
            Mandatory = $true
        )]
        [ValidateCount(1,[int32]::MaxValue)]
        [Microsoft.Exchange.WebServices.Data.UserID[]]
        $UserID,

        # The permission level the delegate(s) should receive for the Calendar folder.
        [Parameter(
            ParameterSetName = "ByParameter",
            ValueFromPipelineByPropertyName = $true            
        )]
        [Alias(
            "CalendarPermissionLevel",
            "CalendarPermissions",
            "CalendarPermission",
            "Calendar"
        )]
        [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]
        $CalendarFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::Editor,

        # The permission level the delegate(s) should receive for the Contacts folder.
        [Parameter(
            ParameterSetName = "ByParameter",
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias(
            "ContactsPermissionLevel",
            "ContactsPermissions",
            "ContactsPermission",
            "Contacts"
        )]
        [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]
        $ContactsFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::None,

        # The permission level the delegate(s) should receive for the Inbox folder.
        [Parameter(
            ParameterSetName = "ByParameter",
            ValueFromPipelineByPropertyName = $true            
        )]
        [Alias(
            "InboxPermissionLevel",
            "InboxPermissions",
            "InboxPermission",
            "Inbox"
        )]
        [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]
        $InboxFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::None,

        # The permission level the delegate(s) should receive for the Journal folder.
        [Parameter(
            ParameterSetName = "ByParameter",
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias(
            "JournalPermissionLevel",
            "JournalPermissions",
            "JournalPermission",
            "Journal"
        )]
        [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]
        $JournalFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::None,

        # The permission level the delegate(s) should receive for the Notes folder.
        [Parameter(
            ParameterSetName = "ByParameter",
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias(
            "NotesPermissionLevel",
            "NotesPermissions",
            "NotesPermission",
            "Notes"
        )]
        [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]
        $NotesFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::None,

        # The permission level the delegate(s) should receive for the Tasks folder.
        [Parameter(
            ParameterSetName = "ByParameter",
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias(
            "TasksPermissionLevel",
            "TasksPermissions",
            "TasksPermission",
            "Tasks"
        )]
        [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]
        $TasksFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::None,

        # A predefined DelegatePermissions object to be applied instead of individual permissions.
        [Parameter(
            Mandatory = $true,
            ParameterSetName = "ByObject",
            ValueFromPipelineByPropertyName = $true
        )]
        [ValidateNotNull()]
        [Microsoft.Exchange.WebServices.Data.DelegatePermissions]
        $Permissions,

        [Switch]$ReceiveMeetingMessageCopies,

        [Switch]$ViewPrivateItems,

        [Switch]$PassThru,

        [ValidateNotNull()]
        [PSCustomObject]
        $Profile = $(Get-EWSProfile -Default)
    )
    
    begin {

        Assert-EWSProfileConnected $Profile

        $ExchSvc = [Microsoft.Exchange.WebServices.Data.ExchangeService]($Profile.ExchangeService)

        $DelegatesToAdd = @()

    }
    
    process {

        $UserID | ForEach-Object {
            $_UserID = $_
            
            $delegate = [Microsoft.Exchange.WebServices.Data.DelegateUser]::new()

            # Copy relevant properties onto the delegate's userID, as the userID object itself is read-only
            "DisplayName","PrimarySmtpAddress","SID","StandardUser" | ForEach-Object {

                $delegate.UserId.$_ = $_UserID.$_

            }

            if ($PSCmdlet.ParameterSetName -eq "ByParameter") {

                "Calendar","Contacts","Inbox","Journal","Notes","Tasks" | ForEach-Object {
                    $delegate.Permissions."$($_)FolderPermissionLevel" = (Get-Variable "$($_)FolderPermissionLevel").Value
                }

            } else {

                # Copy properties onto the delegate's permissions, as Permissions itself is read-only
                "Calendar","Contacts","Inbox","Journal","Notes","Tasks" | ForEach-Object {
                    $delegate.Permissions."$($_)FolderPermissionLevel" = $Permissions."$($_)FolderPermissionLevel"
                }
                
            }

            $delegate.ReceiveCopiesOfMeetingMessages = $ReceiveMeetingMessageCopies

            $delegate.ViewPrivateItems = $ViewPrivateItems

            $DelegatesToAdd += $delegate

        }

        $Response = $ExchSvc.AddDelegates($Mailbox,$null, $DelegatesToAdd)

        $Response | ForEach-Object {

            $_Response = $_

            switch ($_Response.Result) {
                "Success" { 
                    Write-Verbose "Added $($_Response.DelegateUser.UserId.PrimarySmtpAddress) as delegate to $($Mailbox.ToString()) successfully."
                    If ($PassThru) {
                        Write-Output $_Response.DelegateUser
                    }
                 }
                "Error" {
                    Write-Error "Unable to add user $($_Response.DelegateUser.UserId.PrimarySmtpAddress): $($_Response.ErrorMessage)" -ErrorId $_Response.ErrorCode -TargetObject $_Response.DelegateUser
                }
            }


        }

    }
    
    end {


    }
}