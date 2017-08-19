function Set-EWSMailboxDelegate {
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
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
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
        $CalendarFolderPermissionLevel,

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
        $ContactsFolderPermissionLevel,

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
        $InboxFolderPermissionLevel,

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
        $JournalFolderPermissionLevel,

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
        $NotesFolderPermissionLevel,

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
        $TasksFolderPermissionLevel,

        # A predefined DelegatePermissions object to be evaluated before individual permissions.
        [Parameter(
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

        $DelegatesToUpdate = @()

    }
    
    process {

        $UserID | ForEach-Object {

            $_UserID = $_

            try {

                $delegate = [Microsoft.Exchange.WebServices.Data.DelegateUser](Get-EWSMailboxDelegate -Mailbox $Mailbox -UserID $_UserID -ErrorAction Stop)
            
            }
            catch {

                Write-Error -Message "Unable to update delegate $_UserID on mailbox $Mailbox`: $($_)"

                return

            }

            if ($Permissions) {

                # Copy properties onto the delegate's permissions, as Permissions itself is read-only
                "Calendar","Contacts","Inbox","Journal","Notes","Tasks" | ForEach-Object {

                    $delegate.Permissions."$($_)FolderPermissionLevel" = $Permissions."$($_)FolderPermissionLevel"

                }

            }

            "Calendar","Contacts","Inbox","Journal","Notes","Tasks" | ForEach-Object {

                if ((Get-Variable "$($_)FolderPermissionLevel").Value) {

                    $delegate.Permissions."$($_)FolderPermissionLevel" = (Get-Variable "$($_)FolderPermissionLevel").Value
                   
                }

            }
            
            If ($PSBoundParameters.ContainsKey("ReceiveMeetingMessageCopies")) {

                $delegate.ReceiveCopiesOfMeetingMessages = $ReceiveMeetingMessageCopies

            }

            If ($PSBoundParameters.ContainsKey("ViewPrivateItems")) {

                $delegate.ViewPrivateItems = $ViewPrivateItems

            }

            $DelegatesToUpdate += $delegate

        }

        $Response = $ExchSvc.UpdateDelegates($Mailbox, $null, $DelegatesToUpdate)

        $Response | ForEach-Object {

            $_Response = $_

            switch ($_Response.Result) {

                "Success" { 

                    Write-Verbose "Updated delegate permissions for user $($_Response.DelegateUser.UserId.PrimarySmtpAddress) on mailbox $($Mailbox.ToString()) successfully."
                    If ($PassThru) {
                        Write-Output $_Response.DelegateUser
                    }
                }
                "Error" {

                    Write-Error "Unable to update delegate permissions for user $($_Response.DelegateUser.UserId.PrimarySmtpAddress): $($_Response.ErrorMessage)" -ErrorId $_Response.ErrorCode -TargetObject $_Response.DelegateUser

                }

            }

        }

    }
    
    end {

    }
}