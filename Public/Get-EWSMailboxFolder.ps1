function Get-EWSMailboxFolder {
    [OutputType("Microsoft.Exchange.WebServices.Data.Folder")]    
    [CmdletBinding(
        DefaultParameterSetName = "WellKnown"
    )]
    param (
        
        [Parameter(
            Mandatory = $true,
            ParameterSetName = "WellKnown",
            Position = 0
        )]
        [Microsoft.Exchange.WebServices.Data.Mailbox]
        $Mailbox,

        [Parameter(
            Mandatory = $true,
            ParameterSetName = "WellKnown",
            Position = 1
        )]
        [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]
        $WellKnownFolderName,

        [Alias(
            "Id"
        )]
        [Parameter(
            Mandatory = $true,
            ParameterSetName = "ById",
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0
        )]
        [Microsoft.Exchange.WebServices.Data.FolderId]
        $FolderId,

        [ValidateNotNull()]
        [PSCustomObject]
        $Profile = $(Get-EWSProfile -Default)
    )

    DynamicParam {

        $AvailableProperties = [String[]](([Microsoft.Exchange.WebServices.Data.FolderSchema] | Get-Member -Static -MemberType Property | ForEach-Object Name) + '*')

        $AliasAttrib = [Alias]::new("Property","Prop")
        $ValidateSetAttrib = [ValidateSet]::new($AvailableProperties)
        $ParamAttrib = [Parameter]::new()
        

        $PropParamAttribs = [System.Collections.ObjectModel.Collection[Attribute]]::new()
        $PropParamAttribs.Add($AliasAttrib)
        $PropParamAttribs.Add($ValidateSetAttrib)
        $PropParamAttribs.Add($ParamAttrib)

        $PropParam = [System.Management.Automation.RuntimeDefinedParameter]::new("Properties", [String[]], $PropParamAttribs)

        $Params = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()

        $Params.Add("Properties", $PropParam)

        return $Params

    }
    
    begin {
        
        Assert-EWSProfileConnected $Profile

    }
    
    process {

        Write-Verbose "Creating property set..."

        $PropSet = [Microsoft.Exchange.WebServices.Data.PropertySet]::new([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)        
        If ($PSBoundParameters["Properties"] -ne $null) {

            Write-Verbose "Properties specified."

            if ($PSBoundParameters["Properties"] -contains '*') {
                
                Write-Verbose "Adding all properties to property set."
    
                [Microsoft.Exchange.WebServices.Data.FolderSchema] | Get-Member -Static -MemberType Property | ForEach-Object Name | ForEach-Object {
    
                    Write-Verbose "Adding property $_."
    
                    $PropSet.Add([Microsoft.Exchange.WebServices.Data.FolderSchema]::$_)
    
                }
    
            } else {
    
                Write-Verbose "Adding selected properties to property set."
    
                ($PSBoundParameters["Properties"]) | ForEach-Object {
    
                    Write-Verbose "Adding property $_."
    
                    $PropSet.Add([Microsoft.Exchange.WebServices.Data.FolderSchema]::$_)
    
                }
    
            }    
        } else {

            Write-Verbose "No properties specified. Defaulting to first-class properties only."

        }
        

        If ($PSCmdlet.ParameterSetName -eq "WellKnown") {

            $FolderId = [Microsoft.Exchange.WebServices.Data.FolderID]::new($WellKnownFolderName, $Mailbox)

        }


        return [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Profile.ExchangeService, $FolderId, $PropSet)

    }
    
    end {
    }
}