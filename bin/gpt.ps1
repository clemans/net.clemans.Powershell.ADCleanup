# Function to convert full name to sAMAccountName
function Convert-ToSamAccountName {
    param (
        [string]$FullName
    )

    # Split the full name into parts
    $NameParts = $FullName -split ', '

    Write-Output $NameParts
    # Extract last name and first name
    $LastName = $NameParts[0]
    $FirstName = ($NameParts[1] -split ' ')[0]
    
    Write-Debug "First name: $FirstName"
    # Check if the first name contains a middle initial
    if ($FirstName -match '\s[A-Z]\.$') {
        # Extract first initial and middle initial
        $FirstInitial = $FirstName.Substring(0,1).ToLower()
        # $MiddleInitial = $FirstName -replace '.*\s([A-Z])\.$', '$1'

        # Construct the sAMAccountName
        $SamAccountName = $FirstInitial + $LastName.ToLower()
    }
    else {
        # Extract first name without middle initial
        $FirstInitial = $FirstName.Substring(0,1).ToLower()

        # Construct the sAMAccountName
        $SamAccountName = $FirstInitial + $LastName.ToLower()
    }

    return $SamAccountName
}

# Function to output an employee object
function Convert-ToEmployeeObject {
    param (
        [string]$FullName
    )
    # Check if $FullName is null or empty
    if ([string]::IsNullOrWhiteSpace($FullName)) {
        Write-Error "Input full name is null or empty."
        return
    }

    # Use Convert-ToSamAccountName to get sAMAccountName
    $SamAccountName = Convert-ToSamAccountName $FullName

    # Create and return the object
    $EmployeeObject = @{
        GivenName = $FirstName
        SurName = $LastName
        MI = $MiddleInitial
        sAMAccountName = $SamAccountName
    }
    return $EmployeeObject
}

# Function to check Active Directory for account conditions
function Get-ADAccountDeleteStatus {
    param (
        [string]$EmployeeObject
    )
    $fullyQualifiedDomainName = "moorecenter.org"
    $givenName = $EmployeeObject.GivenName
    $givenNameFirst3Letters = $givenName[0..2] -join ""
    $surName = $EmployeeObject.SurName
    $sAMAccountName = $EmployeeObject.sAMAccountName
    $email = "${givenName}.${surName}@${fullyQualifiedDomainName}"
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
        ErrorAction = 'Stop'
    }
    $ADUser = Get-ADUser @parameters
    if ($ADUser) {
        # Check if the account is enabled
        if (-not $ADUser.Enabled) {
            return $true  # Account exists and is disabled
        }
        else {
            return $false  # Account exists and is enabled
        }
    }
    else {
        return $true  # Account does not exist
    }
}

Function Main() {
    Param(
        [Parameter(Mandatory=$true)][System.String]$HRFilePath,
        [Parameter(Mandatory=$true)][System.String]$MSPFilePath,
        [Parameter(Mandatory=$false)][System.Boolean]$ExportToFile
    )
    # Read HR CSV file and extract terminated employees' sAMAccountNames
$HRFile = Import-Csv $HRFilePath
$TerminatedEmployees = $HRFile | ForEach-Object {
    Convert-ToEmployeeObject $_.Employee
}

# Read MSP report and filter terminated employees
$MSPFile = Get-Content $MSPFilePath
$TerminatedMSPAccounts = $MSPFile | Where-Object {
    $EmployeeObject = Convert-ToEmployeeObject $_
    Write-Debug -Message "Employee Object: $EmployeeObject"
    $IsTerminated = $TerminatedEmployees -contains $EmployeeObject.sAMAccountName
    $IsServiceAccount = Get-ADAccountDeleteStatus $EmployeeObject

    $IsTerminated -and $IsServiceAccount
}

# Write filtered sAMAccountNames to new output file
$TerminatedMSPAccounts | Out-File "./data/output/2_TerminatedMSPAccounts.txt"
}

# Inputs
$_args = @{
    HRFilePath   = "./data/input/Terminations 2019 to current.csv"
    MSPFilePath  = "./data/input/userprofiles.txt"
    ExportToFile = $false
    Debug        = $true
}

# Entrypoint
Main @_args
