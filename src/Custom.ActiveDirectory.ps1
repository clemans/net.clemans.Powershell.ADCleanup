<#
 .Synopsis
    Displays a visual representation of a calendar.

 .Description
    Finds employee user object in Active Directory based on Human Resources CSV data.

 .Parameter Row
    The first month to display.

 .Parameter FQDN
    The last month to display.

 .Example
   # TBA
    Get-ADUserFromCSV -Row $Row -FQDN $Fqdn
#>
function Get-ADUserFromCsv {
    param (
        [Parameter(Mandatory=$false)][System.Object]$Row,
        [Parameter(Mandatory=$false)][System.String]$FQDN,
        [Parameter(Mandatory=$false)][System.String]$SamAccountName
    )
    if ($Row) {
        $GivenName = $Row.'First Name'
        $SurName = $Row.'Last Name'
        $FirstInitial = $GivenName.ToLower()[0]
        $FullName = "$GivenName $SurName"
        $FuzzyGivenName = $GivenName[0..2] -join ""
        $StandardsAMAccountName = ($FirstInitial + $SurName).ToLower()
        $InternalEmailAddress = "$StandardsAMAccountName@$FQDN"
        $ExternalEmailAddress = "$GivenName.$SurName@$FQDN"
        $_Properties = "GivenName", "Surname", "Name", "DisplayName", "mail", "EmailAddress", "proxyAddresses", "msExchShadowProxyAddresses", "Enabled", "sAMAccountName"
        $_Filter = {
            (
                ((GivenName -like "${FuzzyGivenName}*") -and (Surname -eq $SurName)) -or
                ((GivenName -eq $GivenName) -and (sAMAccountName -eq $StandardsAMAccountName)) -or
                (GivenName -eq $GivenName -and Surname -eq $SurName) -or
                ((Name -eq $FullName) -or(DisplayName -eq $FullName)) -or
                ((mail -eq $InternalEmailAddress) -or (EmailAddress -eq $InternalEmailAddress)) -or
                ((mail -eq $ExternalEmailAddress) -or (EmailAddress -eq $ExternalEmailAddress)) -or
                ((proxyAddresses -like "*SMTP:$InternalEmailAddress*") -or (proxyAddresses -like "*smtp:$InternalEmailAddress*")) -or
                ((proxyAddresses -like "*SMTP:$ExternalEmailAddress*") -or (proxyAddresses -like "*smtp:$ExternalEmailAddress*")) -or
                ((msExchShadowProxyAddresses -like "*SMTP:$InternalEmailAddress*") -or (msExchShadowProxyAddresses -like "*smtp:$InternalEmailAddress*")) -or
                ((msExchShadowProxyAddresses -like "*SMTP:$ExternalEmailAddress*") -or (msExchShadowProxyAddresses -like "*smtp:$ExternalEmailAddress*"))
            )
        }
        $Parameters = @{
            Filter     = $_Filter
            Properties = $_Properties
            ErrorAction = 'Stop'
        }
        try {
            return Get-ADUser @Parameters
        } catch {
            Write-Debug "Failing on $Row"
        }

    }
}