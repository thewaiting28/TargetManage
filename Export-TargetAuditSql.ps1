[CmdletBinding()]
param ()

# Set variables
$CSVPath = "\\servername\TargetAudit"
$SqlInstance = "sqlinstance.domain.local"
$Database = "TargetAudit"
$Table = "TargetData"

# Get CSV files
$CsvFiles = Get-ChildItem -Path $CSVPath | ? {$_.Name -like "TargetAudit*.csv"}

Function Create-Object (
    $DateTime,
    $KeyString,
    $ComputerName,
    $Username,
    $DisplayName,
    $UserDomain,
    $Target,
    $Type,
    $FullPath,
    $ItemName,
    $RegAppVersion,
    $RegAppName,
    $RegFullValue
) {
    [PSCustomObject]@{
        "DateTime" = $Datetime
        "KeyString" = $KeyString
        "ComputerName" = $ComputerName
        "Username" = $Username
        "DisplayName" = $DisplayName
        "UserDomain" = $UserDomain
        "Target" = $Target
        "Type" = $Type
        "FullPath" = $FullPath
        "ItemName" = $ItemName
        "RegAppVersion" = $RegAppVersion
        "RegAppName" = $RegAppName
        "RegFullValue" = $RegFullValue
    }
}

Function Write-Sql (
    $DateTime,
    $KeyString,
    $ComputerName,
    $Username,
    $DisplayName,
    $UserDomain,
    $Target,
    $Type,
    $FullPath,
    $ItemName,
    $RegAppVersion,
    $RegAppName,
    $RegFullValue,
    $SqlInstance,
    $Database,
    $Table
) {
$Query = @"
INSERT INTO $Table (
    DateTime,
    KeyString,
    ComputerName,
    Username,
    DisplayName,
    UserDomain,
    Target,
    Type,
    FullPath,
    ItemName,
    RegAppVersion,
    RegAppName,
    RegFullValue
    )
VALUES (
    '$DateTime',
    '$KeyString',
    '$ComputerName',
    '$Username',
    '$DisplayName',
    '$UserDomain',
    '$Target',
    '$Type',
    '$FullPath',
    '$Itemname',
    '$RegAppVersion',
    '$RegAppName',
    '$RegFullValue'
    )
"@
    # Write to SQL
    Invoke-Sqlcmd -ServerInstance $SQLInstance -Database $Database -Query $Query
}

function Remove-KeyedItems (
    $KeyString,
    $SqlInstance,
    $Database,
    $Table
) {
$Query = @"
delete from $Table where (KeyString = '$KeyString')
"@
    # Write to SQL
    Invoke-Sqlcmd -ServerInstance $SQLInstance -Database $Database -Query $Query
}

# Process CSVs
$CsvFiles | % {

    # Set variables
    $FileName = $_.Name
    $FilePath = $_.FullName

    # Import CSV contents
    $CSV = Import-Csv -Path $FilePath

    # Get Computer name and username from first record
    $UName = ($CSV | Select -first 1).Username
    $CompName = ($CSV | Select -first 1).ComputerName
    $Key1 = "$Uname-$CompName"

    Remove-KeyedItems -KeyString $Key1 -SqlInstance $SQLInstance -Database $Database -Table $Table

    # Process
    $CSV | % {

        # Set variables
        $DateTime = $_.DateTime
        $ComputerName = $_.ComputerName
        $Username = $_.Username
        $DisplayName = $_.DisplayName
        $UserDomain = $_.UserDomain
        $Target = ($_.Target).Replace("'","")
        $Type = $_.Type
        $FullPath = ($_.FullPath).Replace("'","")
        $ItemName = ($_.ItemName).Replace("'","")
        $RegAppVersion = $_.RegAppVersion
        $RegAppName = $_.RegAppName
        $RegFullValue = ($_.RegFullValue).Replace("'","")
        $KeyString = "$Username-$ComputerName"

        Create-Object `
            -DateTime $DateTime `
            -KeyString $KeyString `
            -ComputerName $ComputerName `
            -Username $Username `
            -DisplayName $DisplayName `
            -UserDomain $UserDomain `
            -Target $Target `
            -Type $Type `
            -FullPath $FullPath `
            -ItemName $ItemName `
            -RegAppVersion $RegAppVersion `
            -RegAppName $RegAppName `
            -RegFullValue $RegFullValue

        # Write to Sql
        Write-Sql -DateTime $DateTime `
            -KeyString $KeyString `
            -ComputerName $ComputerName `
            -Username $Username `
            -DisplayName $DisplayName `
            -UserDomain $UserDomain `
            -Target $Target `
            -Type $Type `
            -FullPath $FullPath `
            -ItemName $ItemName `
            -RegAppVersion $RegAppVersion `
            -RegAppName $RegAppName `
            -RegFullValue $RegFullValue `
            -SqlInstance $SqlInstance `
            -Database $Database `
            -Table $Table
        }
    }
