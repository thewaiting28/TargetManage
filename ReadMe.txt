Get-Targets.ps1
This script was written to be deployed as a configuration item to all users within the scope of shortcuts and recent docs you want to audit. It will return "true" if successful.

Collect-TargetFiles.ps1
This script is designed to be run from a central server, either manually or via scheduled task. It requires a CSV file of computer names with the "Name" column containing the computer names themselves.

Export-TargetAuditSql.ps1
This script examines a given directory and for each file it finds, it writes the data within that CSV file to a SQL database.

--

SQL database will need to be created, service accounts assigned and rights delegated.

Columns:
DateTime
KeyString
ComputerName
Username
DisplayName
UserDomain
Target
Type
FullPath
ItemName
RegAppVersion
RegAppName
RegFullValue
