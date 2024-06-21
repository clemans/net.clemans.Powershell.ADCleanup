# Imports
. "$PSScriptRoot\..\src\Custom.ActiveDirectory.ps1"

# Main Inputs
$_args = @{
    Header = "Employee ID","First Name","Last Name","Hire Date ","Rehire Date ","Type","Exempt","Job Title ","Department","Manager"
    Csv    = "$PSScriptRoot\..\data\input\Employee Roster_NO J-K Cost Centers_6.19.csv"
    List   = "$PSScriptRoot\..\data\input\userprofiles.txt"
    Export = $true
    Debug  = $false
}

# Main Function
Function Main() {
    Param(
        [Parameter(Mandatory=$false)][System.String[]]$Header,
        [Parameter(Mandatory=$false)][System.String]$Csv,
        [Parameter(Mandatory=$false)][System.String]$List,
        [Parameter(Mandatory=$false)][System.Boolean]$Export
    )
    $CsvFile = Import-Csv $Csv -Header $Header
    $ListFile = Get-Content $List
    $ExportData = @()

    # Process CSV data
    foreach ($row in $CsvFile) {
        $Users = Get-ADUserFromCsv -Row $row
        foreach ($User in $Users) {
            $IsDeletable = -not $User.Enabled
            if ($IsDeletable) {
                $ExportData += [PSCustomObject]@{ sAMAccountName = $User.sAMAccountName }
            }
        }
    }

    # Process user profiles list
    foreach ($sam in $ListFile) {
        try {
            $User = Get-ADUser -Identity $sam
            $IsDeletable = -not $User.Enabled
            if ($IsDeletable) {
                $ExportData += [PSCustomObject]@{ sAMAccountName = $User.sAMAccountName }
            }
        } catch {
            $ExportData += [PSCustomObject]@{ sAMAccountName = $sam }
        }
    }

    # Sort the export data by sAMAccountName
    $ExportData = $ExportData | Sort-Object sAMAccountName

    if ($Export) {
        $outputFileName = "{0:yyyyMMdd-HHmm}-deletable-users.csv" -f (Get-Date)
        $outputPath = Join-Path -Path "$PSScriptRoot\..\data\output" -ChildPath $outputFileName
        $ExportData | Export-Csv -Path $outputPath -NoTypeInformation
    }

    return $ExportData
}

# Main Entrypoint
Main @_args