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
    $CsvDate = Get-Date -UFormat "%Y%m%d-%H%M%S"
    $ProfilePath = $env:UserProfile
    $RecentFilesPath = "$env:UserProfile\AppData\Roaming\Microsoft\Windows\Recent"
    $LogFileParent = "$ProfilePath\AppData\Roaming\TargetRewrite"
    $LogFilePath = "$LogFileParent\RewriteLog.log"
    $CsvPath = "$ProfilePath\AppData\Roaming\TargetRewrite\$CsvDate-Results.csv"
    $AutomaticDestinations = "$ProfilePath\AppData\Roaming\Microsoft\Windows\Recent\AutomaticDestinations"
    $CustomDestinations = "$ProfilePath\AppData\Roaming\Microsoft\Windows\Recent\CustomDestinations"

# Look for and create user log file
if (!(Test-Path $LogFileParent)) {
    try {New-Item -ItemType Directory -Path $LogFileParent -ErrorAction Stop | Out-Null}
    catch {Throw "Uanble to create TargetRewrite folder. ($LogFileParent)"}
    }
if (!(Test-Path $LogFilePath)) {
    try {New-Item -ItemType File -Path $LogFilePath -ErrorAction Stop | Out-Null}
    catch {Throw "Unable to create LogFilePath. ($LogFilePath)"}
    }

# Create Csv File
if ($OutputToCSV) {New-Item -ItemType File -Path $CsvPath -Force | Out-Null}

Function Create-TargetItem (
    $DateTime,
    $ComputerName,
    $Username,
    $DisplayName,
    $UserDomain,
    $OriginalTarget,
    $NewTarget,
    $Type,
    $FullPath,
    $ItemName,
    $RegAppVersion,
    $RegAppName,
    $RegFullValue,
    $ChangeApplied
    ) {
    [PSCustomObject]@{
        "DateTime" = $DateTime
        "ComputerName" =  $ComputerName
        "Username" = $Username
        "DisplayName" = $DisplayName
        "UserDomain" = $UserDomain
        "OriginalTarget" = $OriginalTarget
        "NewTarget" = $NewTarget
        "Type" = $Type
        "FullPath" = $FullPath
        "ItemName" = $ItemName
        "RegAppVersion" = $RegAppVersion
        "RegAppName" = $RegAppName
        "RegFullValue" = $RegFullValue
        "ChangeApplied" = $ChangeApplied
        }
    }

Function Convert-TargetString (
    [string]$OriginalString
    ) {
    switch -Wildcard ($OriginalString) {

        # Share01
        "*\\srv01\share01\*" {
            $OriginalString -replace [regex]::Escape("\\svr01\share01"), "\\dfsname.local\new\share01"
            }
        "*\\dfsname.local\old\share01\*" {
            $OriginalString -replace [regex]::Escape("\\dfsname.local\old\share01"), "\\dfsname.local\new\share01"
            }

        # Share02
        "*\\srv01\share02\*" {
            $OriginalString -replace [regex]::Escape("\\svr01\share02"), "\\dfsname.local\new\share02"
            }
        "*\\dfsname.local\old\share02\*" {
            $OriginalString -replace [regex]::Escape("\\dfsname.local\old\share02"), "\\dfsname.local\new\share02"
            }
        
        Default {$OriginalString}
        }
    }

# Process registry
$FileItems = Get-ChildItem -Path $FileMRU -ErrorAction SilentlyContinue
$PlaceItems = Get-ChildItem -Path $PlaceMRU -ErrorAction SilentlyContinue

# Registry: FileItems
$FileCollection = $FileItems | % {

    # Set variables
        # Replaces strings for easier queries
        $RegPath = ($_.Name).Replace("HKEY_CURRENT_USER\","HKCU:\")
    
        # Splits path and sets variables to indicate office app and version
        $RegSplitPath = $RegPath.Split("\")
        $AppVersion = $RegSplitPath[4]
        $AppName = $RegSplitPath[5]

    # Collect properties
    $Properties = (Get-ItemProperty -Path $RegPath -ErrorAction SilentlyContinue | Get-Member -ErrorAction SilentlyContinue | ? {$_.Name -like "Item*"}).Name

    # Process individual item properties
    $Properties | % {
        # Set property name
        $Name = $_

        # Get property value
        $OldFullValue = (Get-ItemProperty -Path $RegPath -Name $Name -ErrorAction SilentlyContinue).$Name
        
        # Try to split the property value into just the file name
        try {$OldShortValue = ($OldFullValue).Split("*")[1]}
        catch {}

        # Convert target value
        $NewFullValue = Convert-TargetString -OriginalString $OldFullValue
        $NewShortvalue = Convert-TargetString -OriginalString $OldShortValue

        # Rewrite to registry
        if ($OldShortValue -ne $NewShortValue) {
            try {Set-ItemProperty -Path $RegPath -Name $Name -Value $NewFullValue -ErrorAction SilentlyContinue}
            catch {}
            $ChangeApplied = $true
            "$DateTime // Original: $OldShortValue  |  New: $NewShortValue" | Out-File -FilePath $LogFilePath -Append -Encoding ascii
            }
        else {$ChangeApplied = $false}

        # If target is on the C:\ drive, return
        if ($Target -like "C:\*") {return}

        # Create an object
        Create-TargetItem `
            -DateTime $DateTime `
            -ComputerName $ComputerName `
            -Username $Username `
            -DisplayName $DisplayName `
            -UserDomain $UserDomain `
            -OriginalTarget $OldShortValue `
            -NewTarget $NewShortValue `
            -Type "FileMRU" `
            -FullPath $RegPath `
            -ItemName $Name `
            -RegAppVersion $AppVersion `
            -RegAppName $AppName `
            -RegFullValue $NewFullValue `
            -ChangeApplied $ChangeApplied
        }
    }

# Registry: FileItems
$PlaceCollection = $PlaceItems | % {
    # Set variables
        # Replaces strings for easier queries
        $RegPath = ($_.Name).Replace("HKEY_CURRENT_USER\","HKCU:\")
    
        # Splits path and sets variables to indicate office app and version
        $RegSplitPath = $RegPath.Split("\")
        $AppVersion = $RegSplitPath[4]
        $AppName = $RegSplitPath[5]

    # Collect properties
    $Properties = (Get-ItemProperty -Path $RegPath -ErrorAction SilentlyContinue | Get-Member -ErrorAction SilentlyContinue | ? {$_.Name -like "Item*"}).Name

    # Process individual item properties
    $Properties | % {
        # Set property name
        $Name = $_

        # Get property value
        $OldFullValue = (Get-ItemProperty -Path $RegPath -Name $Name -ErrorAction SilentlyContinue).$Name
        
        # Try to split the property value into just the file name
        try {$OldShortValue = ($OldFullValue).Split("*")[1]}
        catch {}

        # Convert target value
        $NewFullValue = Convert-TargetString -OriginalString $OldFullValue
        $NewShortvalue = Convert-TargetString -OriginalString $OldShortValue

        # Rewrite to registry
        if ($OldShortValue -ne $NewShortValue) {
            try {Set-ItemProperty -Path $RegPath -Name $Name -Value $NewFullValue -ErrorAction SilentlyContinue}
            catch {}
            $ChangeApplied = $true
            "$DateTime // Original: $OldShortValue  |  New: $NewShortValue" | Out-File -FilePath $LogFilePath -Append -Encoding ascii
            }
        else {$ChangeApplied = $false}

        # If target is on the C:\ drive, return
        if ($Target -like "C:\*") {return}

        # Create an object
        Create-TargetItem `
            -DateTime $DateTime `
            -ComputerName $ComputerName `
            -Username $Username `
            -DisplayName $DisplayName `
            -UserDomain $UserDomain `
            -OriginalTarget $OldShortValue `
            -NewTarget $NewShortValue `
            -Type "FileMRU" `
            -FullPath $RegPath `
            -ItemName $Name `
            -RegAppVersion $AppVersion `
            -RegAppName $AppName `
            -RegFullValue $NewFullValue `
            -ChangeApplied $ChangeApplied
        }
    }

# Process File System
$Shortcuts = Get-ChildItem -Path $ProfilePath -Recurse -ErrorAction SilentlyContinue | ? {$_.Name -like "*.lnk"}
$Recent = Get-ChildItem -Path $RecentFilesPath -ErrorAction SilentlyContinue | ? {$_.Name -like "*.lnk"}

# FileSystem: Profile
$ShortcutCollection = $Shortcuts | % {

    # Set variables
    $ShortcutFileName = $_.Name
    $ShortcutFilePath = $_.FullName

    # Create shortcut object
    try {
        $Shell = New-Object -ComObject WScript.Shell
        $Shortcut = $Shell.CreateShortcut($ShortcutFilePath)
        $OriginalTarget = $Shortcut.TargetPath
        }
    catch {}

    # Filter out targets that point back to the C:\ drive
    if ($OriginalTarget -like "C:\*") {Return}
    if ($OriginalTarget -like "") {return}

    # Modify existing shortcut
    try {
        $NewTarget = Convert-TargetString -OriginalString $OriginalTarget
        $Shortcut.TargetPath = $NewTarget
        }
    catch {}

    # Save the modified shortcut
    if ($OriginalTarget -ne $NewTarget) {
        try {$Shortcut.Save()}
        catch {}
        "$DateTime // Original: $OriginalTarget  |  New: $NewTarget" | Out-File -FilePath $LogFilePath -Append -Encoding ascii
        $ChangeApplied = $true
    }
    else {$ChangeApplied = $false}

    # Create an object
    Create-TargetItem `
        -DateTime $DateTime `
        -ComputerName $ComputerName `
        -Username $Username `
        -DisplayName $DisplayName `
        -UserDomain $UserDomain `
        -OriginalTarget $OriginalTarget `
        -NewTarget $NewTarget `
        -Type "Shortcut" `
        -FullPath $ShortcutFilePath `
        -ItemName $ShortcutFileName `
        -RegAppVersion $null `
        -RegAppName $null `
        -RegFullValue $null `
        -ChangeApplied $ChangeApplied
        }

# FileSystem: Recent
$RecentCollection = $Recent | % {

    # Set variables
    $ShortcutFileName = $_.Name
    $ShortcutFilePath = $_.FullName

    # Create shortcut object
    try {
        $Shell = New-Object -ComObject WScript.Shell
        $Shortcut = $Shell.CreateShortcut($ShortcutFilePath)
        $OriginalTarget = $Shortcut.TargetPath
        }
    catch {}

    # Filter out targets that point back to the C:\ drive
    if ($OriginalTarget -like "C:\*") {Return}
    if ($OriginalTarget -like "") {return}

    # Modify existing shortcut
    try {
        $NewTarget = Convert-TargetString -OriginalString $OriginalTarget
        $Shortcut.TargetPath = $NewTarget
        }
    catch {}

    # Save the modified shortcut
    if ($OriginalTarget -ne $NewTarget) {
        try {$Shortcut.Save()}
        catch {}
        "$DateTime // Original: $OriginalTarget  |  New: $NewTarget" | Out-File -FilePath $LogFilePath -Append -Encoding ascii
        $ChangeApplied = $true
    }
    else {$ChangeApplied = $false}

    # Create an object
    Create-TargetItem `
        -DateTime $DateTime `
        -ComputerName $ComputerName `
        -Username $Username `
        -DisplayName $DisplayName `
        -UserDomain $UserDomain `
        -OriginalTarget $OriginalTarget `
        -NewTarget $NewTarget `
        -Type "Shortcut" `
        -FullPath $ShortcutFilePath `
        -ItemName $ShortcutFileName `
        -RegAppVersion $null `
        -RegAppName $null `
        -RegFullValue $null `
        -ChangeApplied $ChangeApplied
    }

# Remove automatic and custom destinations
Remove-Item -Path $AutomaticDestinations -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
Remove-Item -Path $CustomDestinations -Recurse -Force -ErrorAction SilentlyContinue | Out-Null

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
