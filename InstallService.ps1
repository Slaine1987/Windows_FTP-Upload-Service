# The path to PowerShell
$Binary = (Get-Command Powershell).Source
$ScriptLocation = (Get-Location).Path
$ServicesLocation = "$ScriptLocation\Services.ps1"


# The necessary arguments, including the path to our script
$Arguments = "-ExecutionPolicy Bypass -NoProfile -File $ServicesLocation"


#write-host $Aruments

# Creating the service
.\nssm.exe install FTPUploadService $Binary $Arguments