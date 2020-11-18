[CmdletBinding()]
param ()

# Set variables
    # Destination path -- where do you want this collection script to dump the collected CSV files?
    $DestinationPath = "\\servername\TargetAudit"

    # Where is the CSV file with computer names? CSV file should have just one column, "Name"
    $ComputerListPath = "\\servername\TargetAudit\ComputerLists\Computers.csv"

# Test ComputerListPath
if (!(Test-Path $ComputerListPath)) {Throw "Unable to find Computers.csv. ($ComputerListPath)"}

# Import ComputerListPath as Csv and select the Name variable
$ComputerName = (Import-Csv -Path $ComputerListPath).Name

# Set up count for verbose
$ComputerCount = ($ComputerName | Measure-Object).Count
$ComputerLoop = 1

# Process each computer
$ComputerName | % {

    # Set variables
    $Name = $_

    # Count
    Write-Host "$ComputerLoop of $ComputerCount`: $Name"

    # Test connection
    if (!(Test-Connection -ComputerName $Name -Count 1 -Quiet)) {
        # If computer does not respond, bypass and go to the next computer name in the pipeline
        Return
        }
    
    # Test accessing users folder
    if (!(Test-Path -Path "\\$Name\c`$\Users")) {
        Write-Warning "Unable to access path. (\\$Name\c$\Users)"
        Return
        }

    # Get all TargetAudit csv files
    $Files = Get-Item -Path "\\$Name\c`$\Users\*\AppData\Roaming\TargetAudit-*.csv"

    # Process each file
    $Files | % {
        # Set vars
        $FileName = $_.Name
        $FilePath = $_.FullName

        # If existing file exists at the destination path, delete it
        if (Test-Path "$DestinationPath\$FileName") {
            Remove-Item -Path "$DestinationPath\$FileName" -Force
            }
        
        # Copy file to destination path
        Copy-Item -Path $FilePath -Destination "$DestinationPath\$FileName"

        }

    $ComputerLoop++
    }