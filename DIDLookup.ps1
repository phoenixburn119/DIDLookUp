Import-module ImportExcel
Connect-MicrosoftTeams

# CHANGE THIS PATH. This is the path to the existing DID sheet we use. Standard columns for script 
# to work are "LineURI", "Display Name", and "SipAddress".
$DIDSheet = Import-Excel -Path ".\Master DID.xlsx" -WorksheetName "Teams Users"
# Delcares the array used by the ValueCollector when combing through DIDSheet.
[System.Collections.ArrayList]$DataCollector = @()

# Used to test connection to MS Teams. Requires a test account, can be anything just has to exist.
Function MicrosoftTeamsConnect {
    Write-Host "Checking connection to MicrosoftTeams..." -BackgroundColor DarkGray
    Try{
        # Change "test@test.com" to an email that exists in your Teams environment (must have teams license).
        Get-CsOnlineUser -Identity test@test.com | Out-null
        Write-Host "Thanks for connecting to MicrosoftTeams prior ;)"
    } Catch{
        Write-Warning "Connecting to MicrosoftTeams...Please follow the popup"
        Connect-MicrosoftTeams
    }
    Pause
}

# Basic function used to generate logs and deposited into the AuditLog directory.
function Write-Log {
    Param(
        $Message,$Path = ".\Audit-Log $($env:username) $($(get-date).ToString("MM-dd-yyyy")).txt"
    )

    function TS {Get-Date -Format 'hh:mm:ss'}
    "[$(TS)]$Message" | Tee-Object -FilePath $Path -Append | Write-Verbose
}

# Function that is called by main loop to first try to see if the number belongs to an account and 
# return that account info to main array DataCollector. Otherwise will write any errors in the log.
Function NumberLookup {
    [CmdletBinding(DefaultParameterSetName='Name')]
    param(
        [parameter(Mandatory)]
        [string] $Number
    )
    Try{
        $TeamsReturn = get-csonlineuser -Filter "LineURI -eq '$($Number)'"
        Write-Log "$($DIDSheet[$idx].LineURI) : $($TeamsReturn.DisplayName) : $($TeamsReturn.UserPrincipalName) : Sheet States- $($DIDSheet[$idx].DisplayName)"
        
        $ValueCollector = [pscustomobject]@{'DisplayName'=$($TeamsReturn.UserPrincipalName);`
        'LineURI'=$($DIDSheet[$idx].LineURI);'SheetStates'=$($DIDSheet[$idx].DisplayName);`
        'SheetSIP'=$(If($null -ne $DIDSheet[$idx].SipAddress){Get-SIP -SIP $DIDSheet[$idx].SipAddress}else{$null});`
        'Match'=$null}
        
        $DataCollector.add($ValueCollector) | Out-Null
        $ValueCollector = $null
    } Catch{
        Write-Error "$_ : There was an issue retrieving the users data associated with that number."
        Write-Log "$_ : There was an issue retrieving the users data associated with that number."
    }
}

# Used to remove "sip:" that is inside the label of "SipAddress" column. Will skip unless "sip:" exists in beginning of string.
Function Get-SIP {
    [CmdletBinding(DefaultParameterSetName='Name')]
    param(
        [parameter(Mandatory)]
        [string] $SIP
    )
    If($null -ne $SIP) {    
        If((($DIDSheet[$idx].SipAddress).substring(0,4)) -eq "sip:") {
                $SIPVal = ($SIP).substring(4)
                Write-Output "$SIPVal"
            } else {
                Write-Output "$SIPVal"
            }
    }
}

# After the checks are complete this function compares the values and returns True or Not Allocated 
# values to be added to the final export.
Function Get-Match {
    If($DataCollector[$idx].DisplayName -eq $DataCollector[$idx].SheetSIP){
        $DataCollector[$idx].Match = "True"
    }
    If(($null -eq $DataCollector[$idx].DisplayName) -and ($null -eq $DataCollector[$idx].SheetSIP)) {
        $DataCollector[$idx].Match = "Not Allocated"
    }
}

# Main loop for the program. It also adds a progress bar showing progress through DIDSheet. Depending on size it may take awhile.
MicrosoftTeamsConnect
$ProgressLength = 100 / $DIDSheet.count
for($idx = 0; $idx -lt $DIDSheet.count; $idx++ ) {
    $e = $e + $ProgressLength
    Write-Progress -Activity "DID lookup in progress. Currently on : $($DIDSheet[$idx])" -Status "$e% Complete:" -PercentComplete $e
    NumberLookup -Number "$($DIDSheet[$idx].LineURI)"
    Get-Match
}
$DataCollector | export-csv -Path .\csv.csv