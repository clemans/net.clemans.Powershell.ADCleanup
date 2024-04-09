Function Get-FullNameFromCsv() {
    Param([Parameter(Mandatory = $true)]$User)
    return "$($User."First Name") $($User."Last Name")"
}

Function Get-SamAccountNameFromCsv() {
    Param([Parameter(Mandatory = $false)][System.Object]$UserRow)
    return "$(($UserRow.'First Name')[0])$($UserRow.'Last Name')"
}

Function Get-ADAccountFromAttributes() {
    Param(
        [Parameter(Mandatory = $true)]$User,
        [Parameter(Mandatory = $false)][switch]$OfflineMode
    )
    $fullyQualifiedDomainName = "moorecenter.org"
    $givenName = $User.'First Name'
    $givenNameFirst3Letters = $givenName[0..2] -join ""
    $surName = $User.'Last Name'
    $sAMAccountName = Get-SamAccountNameFromCsv -User $User
    $email = "${givenName}.${surName}@${fullyQualifiedDomainName}"
    $fullName = Get-FullNameFromCsv -User $User
    $filter = {
        (
            ((GivenName -like "${givenNameFirst3Letters}*") -and (Surname -eq $surName)) -or
            ((GivenName -eq $givenName) -and (sAMAccountName -eq $sAMAccountName)) -or
            (GivenName -eq $givenName -and Surname -eq $surName) -or
            (Name -eq $fullName) -or
            (DisplayName -eq $fullName) -or
            (mail -eq $email) -or
            (EmailAddress -eq $email) -or
            ((proxyAddresses -like "*SMTP:$email*") -or (proxyAddresses -like "*smtp:$email*")) -or
            ((msExchShadowProxyAddresses -like "*SMTP:$email*") -or (msExchShadowProxyAddresses -like "*smtp:$email*"))
        )
    }
    $properties = "GivenName", "Surname", "Name", "DisplayName", "mail", "EmailAddress", "proxyAddresses", "msExchShadowProxyAddresses", "Enabled"
    $parameters = @{
        Filter     = $filter
        Properties = $properties
    }
        if (Get-ADUser @parameters) {
            return Get-ADUser @parameters
        }
        return @{
            SamAccountName = $sAMAccountName
            Name = $fullName
            Exists = $false
            Enabled = $false
        }
}

Function ConvertTo-DataTable {
    Param(
        [Parameter(Mandatory = $true)]
        [System.Collections.ArrayList]$ArrayList
    )

    $table = New-Object System.Data.DataTable

    # Create columns based on the properties of the first object in the ArrayList
    $ArrayList[0].PSObject.Properties | ForEach-Object {
        $column = New-Object System.Data.DataColumn($_.Name)
        $table.Columns.Add($column)
    }

    # Add rows to the DataTable
    foreach ($item in $ArrayList) {
        $row = $table.NewRow()
        foreach ($prop in $item.PSObject.Properties) {
            $row[$prop.Name] = $prop.Value
        }
        $table.Rows.Add($row)
    }
    return $table | Sort-Object -Property sAMAccountName
}

Function Export-DataTableFile() {
    Param(
        [Parameter(Mandatory=$true)]
        [System.String]$FileName,
        [Parameter(Mandatory=$true)]
        [System.Collections.ArrayList]$ArrayList
    )
    $timestamp = Get-Date -Format "yyyyMMdd.HHmm"
    $csvFilePath = "./data/output/${timestamp}_${FileName}.csv"
    $dataTable = ConvertTo-DataTable -ArrayList $ArrayList
    $dataTable | Export-Csv -Path $csvFilePath -NoTypeInformation
}
Function Main() {
    Param(
        [Parameter(Mandatory = $true)]
        [System.Object]$HRFile,
        [Parameter(Mandatory = $true)]
        [System.Object]$MSPFile,
        [Parameter(Mandatory = $false)]
        [System.Boolean]$ExportToFile = $false
    )
    $HRObject = New-Object System.Collections.ArrayList
    $samAccountNames = @{}
    $fullNames = @{}

    foreach ($hrAccount in $HRFile) {
        $samAccountNames[$hrAccount] = (Get-SamAccountNameFromCsv -User $hrAccount)
        $fullNames[$hrAccount] = (Get-FullNameFromCsv -User $hrAccount)
        Get-ADAccountFromAttributes -User $hrAccount | ForEach-Object {
            $sAMAccountName = $_ ? $_.SamAccountName : $samAccountNames[$hrAccount]
            $fullName = $_ ? $_.Name : $fullNames[$hrAccount]
            $exists = $_.ObjectClass -eq 'user'
            $enabled = $_.Enabled
            $HRObject.Add([PSCustomObject]@{
                sAMAccountName = $sAMAccountName
                FullName       = $fullName
                Exists         = $exists
                Enabled        = $enabled
            })
            Write-Debug "`nsAMAccountName: ${sAMAccountName}`nFullName: ${fullName}`nExists: ${exists}`nStatus: ${enabled}`n"
        }
    }

    if ($ExportToFile) {
        Export-DataTableFile -FileName "HR_All" -ArrayList $HRObject
    }
    return 0
}

# Inputs
$_args = @{
    HRFile       = $(Import-Csv ".\data\input\Employee Roster 3.26.24.csv")
    MSPFile      = $(Get-Content ".\data\input\userprofiles.txt")
    ExportToFile = $true
    Debug        = $true
}

# Entrypoint
Main @_args