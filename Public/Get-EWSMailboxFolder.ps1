function Get-EWSMailboxFolder {
    [OutputType("Microsoft.Exchange.WebServices.Data.Folder")]    
    [CmdletBinding(
        DefaultParameterSetName = "WellKnown"
    )]
    param (
        
        [Microsoft.Exchange.WebServices.Data.Mailbox]
        $Mailbox,

        [Parameter(
            Mandatory = $true,
            ParameterSetName = "WellKnown"
        )]
        [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]
        $WellKnownFolderName,

        # [Microsoft.Exchange.WebServices.Data.PropertyDefinition[]]$Properties,

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

        if ($Properties -contains '*') {

            $PropSet = [Microsoft.Exchange.WebServices.Data.PropertySet]::new([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties, ([Microsoft.Exchange.WebServices.Data.FolderSchema] | Get-Member -Static -MemberType Property | ForEach-Object Name))

        } else {

            $PropSet = [Microsoft.Exchange.WebServices.Data.PropertySet]::new([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties, $Properties)

        }

        


        return [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Profile.ExchangeService,[Microsoft.Exchange.WebServices.Data.FolderID]::new($WellKnownFolderName, $Mailbox), $PropSet)

    }
    
    end {
    }
}