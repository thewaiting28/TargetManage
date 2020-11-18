[CmdletBinding()]
param ()

# Toggle options
$ConfigItemMode = $true
$OutputToConsole = $false
$OutputToCSV = $true

# Set variables
    $Username = $env:Username
    $ComputerName = $env:ComputerName
    $UserDomain = $env:UserDomain
    $DateTime = Get-Date -UFormat "%Y-%m-%d %H:%M:%S"

    # Display name try/catch
    try {$DisplayName = (([ADSI]"WinNT://$UserDomain/$UserName,user") | Select-Object -ExpandProperty FullName -ErrorAction SilentlyContinue)}
    catch {}

    # Target paths
    $FileMRU = "HKCU:\Software\Microsoft\Office\*\*\User MRU\*\File MRU"
    $PlaceMRU = "HKCU:\Software\Microsoft\Office\*\*\User MRU\*\Place MRU"
    $ProfilePath = $env:UserProfile
    $RecentFilesPath = "$env:UserProfile\AppData\Roaming\Microsoft\Windows\Recent"

# CSV
$CsvPath = "$ProfilePath\AppData\Roaming\TargetAudit-$Username-$ComputerName.csv"
    # If CSV path exists, delete it
    if (Test-Path $CsvPath) {
        try {Remove-Item -Path $CsvPath -Force}
        catch {Throw "Unable to remove CsvPath. ($CsvPath)"}
        }
    # Create CSV
    New-Item -Path $CsvPath -ItemType File -Force | Out-Null

Function Create-TargetItem (
    $DateTime,
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
        "DateTime" = $DateTime
        "ComputerName" =  $ComputerName
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

# Process registry
    $FileItems = Get-ChildItem -Path $FileMRU
    $PlaceItems = Get-ChildItem -Path $PlaceMRU


# Registry: FileItems
$FileCollection = $FileItems | % {

    # Set variables
    $RegPath = ($_.Name).Replace("HKEY_CURRENT_USER\","HKCU:\")
    $RegSplitPath = $RegPath.Split("\")
    $AppVersion = $RegSplitPath[4]
    $AppName = $RegSplitPath[5]

    # Collect properties
    $Properties = (Get-ItemProperty -Path $RegPath | Get-Member -ErrorAction SilentlyContinue | ? {$_.Name -like "Item*"}).Name

    # Process individual item properties
    $Properties | % {
        # Set property name
        $Name = $_

        # Get property value
        $RegFullValue = (Get-ItemProperty -Path $RegPath -Name $Name).$Name
        
        # Try to split the property value into just the file name
        try {$Target = ($RegFullValue).Split("*")[1]}
        catch {}

        # If target is on the C:\ drive, return
        if ($Target -like "C:\*") {return}

        # Create an object
        Create-TargetItem `
            -DateTime $DateTime `
            -ComputerName $ComputerName `
            -Username $Username `
            -DisplayName $DisplayName `
            -UserDomain $UserDomain `
            -Target $Target `
            -Type "FileMRU" `
            -FullPath $RegPath `
            -ItemName $Name `
            -RegAppVersion $AppVersion `
            -RegAppName $AppName `
            -RegFullValue $RegFullValue
        }
    }


# Registry: FileItems
$PlaceCollection = $PlaceItems | % {

    # Set variables
    $RegPath = ($_.Name).Replace("HKEY_CURRENT_USER\","HKCU:\")
    $RegSplitPath = $RegPath.Split("\")
    $AppVersion = $RegSplitPath[4]
    $AppName = $RegSplitPath[5]

    # Collect properties
    $Properties = (Get-ItemProperty -Path $RegPath | Get-Member -ErrorAction SilentlyContinue | ? {$_.Name -like "Item*"}).Name

    # Process individual item properties
    $Properties | % {
        # Set property name
        $Name = $_

        # Get property value
        $RegFullValue = (Get-ItemProperty -Path $RegPath -Name $Name).$Name
        
        # Try to split the property value into just the file name
        try {$Target = ($RegFullValue).Split("*")[1]}
        catch {}

        # If target is on the C:\ drive, return
        if ($Target -like "C:\*") {return}

        # Create an object
        Create-TargetItem `
            -DateTime $DateTime `
            -ComputerName $ComputerName `
            -Username $Username `
            -DisplayName $DisplayName `
            -UserDomain $UserDomain `
            -Target $Target `
            -Type "PlaceMRU" `
            -FullPath $RegPath `
            -ItemName $Name `
            -RegAppVersion $AppVersion `
            -RegAppName $AppName `
            -RegFullValue $RegFullValue
        }
    }


# Process File System
$Shortcuts = Get-ChildItem -Path $ProfilePath -Recurse | ? {$_.Name -like "*.lnk"}
$Recent = Get-ChildItem -Path $RecentFilesPath | ? {$_.Name -like "*.lnk"}

# FileSystem: Profile
$ShortcutCollection = $Shortcuts | % {

    # Set variables
    $ShortcutFileName = $_.Name
    $ShortcutFilePath = $_.FullName

    # Create shortcut object
    $Shortcut = New-Object -ComObject WScript.Shell
    $Target = $Shortcut.CreateShortcut($ShortcutFilePath).TargetPath

    # Filter out targets that point back to the C:\ drive
    if ($Target -like "C:\*") {Return}
    if ($Target -like "") {return}

    # Create an object
    Create-TargetItem `
        -DateTime $DateTime `
        -ComputerName $ComputerName `
        -Username $Username `
        -DisplayName $DisplayName `
        -UserDomain $UserDomain `
        -Target $Target `
        -Type "Shortcut" `
        -FullPath $ShortcutFilePath `
        -ItemName $ShortcutFileName `
        -RegAppVersion $null `
        -RegAppName $null `
        -RegFullValue $null
    }

# FileSystem: Profile
$RecentCollection = $Recent | % {

    # Set variables
    $ShortcutFileName = $_.Name
    $ShortcutFilePath = $_.FullName

    # Create shortcut object
    $Shortcut = New-Object -ComObject WScript.Shell
    $Target = $Shortcut.CreateShortcut($ShortcutFilePath).TargetPath

    # Filter out targets that point back to the C:\ drive
    if ($Target -like "C:\*") {Return}
    if ($Target -like "") {return}

    # Create an object
    Create-TargetItem `
        -DateTime $DateTime `
        -ComputerName $ComputerName `
        -Username $Username `
        -DisplayName $DisplayName `
        -UserDomain $UserDomain `
        -Target $Target `
        -Type "Shortcut" `
        -FullPath $ShortcutFilePath `
        -ItemName $ShortcutFileName `
        -RegAppVersion $null `
        -RegAppName $null `
        -RegFullValue $null
    }

# Export results
if ($OutputToCSV) {

    # FileMRU
    $FileCollection | Export-Csv -Path $CsvPath -NoTypeInformation -Append

    # PlaceMRU
    $PlaceCollection | Export-Csv -Path $CsvPath -NoTypeInformation -Append

    # Shortcut
    $ShortcutCollection | Export-Csv -Path $CsvPath -NoTypeInformation -Append

    # Recent
    $RecentCollection | Export-Csv -Path $CsvPath -NoTypeInformation -Append

    }

# Report results on screen
if ($OutputToConsole) {
    
    # FileMRU
    $FileCollection
    
    # PlaceMRU
    $PlaceCollection
    
    # Shortcut
    $ShortcutCollection
    
    # Recent
    $RecentCollection
    }

# Config item mode will return true on completion
if ($ConfigItemMode) {$True}