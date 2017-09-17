# The UserID class doesn't appear to offer a loose comparison, so we have to do it ourselves.
function Test-EWSUserIdMatch {
    [OutputType("bool")]
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory = $true,
            Position = 0
        )]
        [Microsoft.Exchange.WebServices.Data.UserId]$Left,

        [Parameter(
            Mandatory = $true,
            Position = 1
        )]
        [Microsoft.Exchange.WebServices.Data.UserId]$Right
    )
    
    begin {
    }
    
    process {

        If ($Left -eq $null) {

            Write-Debug "Left UserID is null, cannot compare. Returning false."

            return $false

        }

        If ($Right -eq $null) {

            Write-Debug "Right UserID is null, cannot compare. Returning false."

            return $false

        }

        Write-Debug "Comparing UserID $($Left.ToString()) and $($Right.ToString())..."

       foreach ($property in  @("PrimarySmtpAddress", "SID", "StandardUser")) {

            Write-Debug "Comparing property $property."

            Write-Debug "L: $((@($Left.$Property, "<null>") -ne $null)[0])"
            Write-Debug "R: $((@($Right.$Property, "<null>") -ne $null)[0])"


            if (($Left.$property -eq $null) -or ($Right.$property -eq $null)) {

                Write-Debug "Property $property is null on one or both objects. Inconclusive."

                Continue

            } elseif ($Left.$property -ne $Right.$property) {

                Write-Debug "Property $property does not match"

                return $false

            } else {

                Write-Debug "Property $property matches"

                return $true

            }

        }
        
        Write-Debug "Unable to conclusively compare. Returning false by elimination."

        return $false
    }
    
    end {
    }
}