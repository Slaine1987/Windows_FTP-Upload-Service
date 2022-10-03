#Service to upload files every x minutes

#Parameters
$ScriptLocation = (Get-Location).Path
$WinSCPDll = "$ScriptLocation\WinSCPnet.dll"
$IniLocation = "$ScriptLocation\config.ini"

function Test-EventLog {
    Param(
        [Parameter(Mandatory=$true)]
        [string] $LogName
    )

    [System.Diagnostics.EventLog]::SourceExists($LogName)
}

function InitEventLog(){
    $logFileExists = Test-EventLog("FTPUpload")
    if (! $logFileExists) {
        New-EventLog -LogName "Application" -Source "FTPUpload"
    }
}

function CreateEventLog ($Message) {
    InitEventLog
    Write-EventLog -LogName "Application" -Source "FTPUpload" -EventID 3001 -EntryType Information -Message $Message -RawData 10,20
}

#Test all parameters
If (Test-Path($IniLocation)) { 
    write-host "Ini found"
}else {
    CreateEventLog("$IniLocation not found")
    exit 1 
}

If (Test-Path($WinSCPDll)) { 
    write-host "Dll found"
}else{
    CreateEventLog("$WinSCPDll not found")
    exit 1 
}

$INI = Get-Content $IniLocation

ForEach($Line in $INI)
{
    $SplitArray = $Line.Split("=")
    switch ($SplitArray[0]) {
        "FTPHostname" { $FTPHostname = $SplitArray[1] }
        "FTPUsername" { $FTPUsername = $SplitArray[1] }
        "FTPPassword" { $FTPPassword = $SplitArray[1] }
        "LocalPath" { $LocalPath = $SplitArray[1] }
        "RemotePath" { $RemotePath = $SplitArray[1] }
        "BackupPath" { $RemotePath = $SplitArray[1] }
        "IntervalMinutes" { $IntervalMinutes = $SplitArray[1] }        
    }
}


$LogFileDate = $(((get-date).ToUniversalTime()).ToString("yyyyMMdd"))
$LogFileTime = $(((get-date)).ToString("HH:mm:ss"))
$ItemsUploaded = 0

function StartUpload() {
    
    CreateEventLog("Upload started")
    # Load WinSCP .NET assembly
    Add-Type -Path $WinSCPDll
 
    # Set up session options
    $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
        Protocol = [WinSCP.Protocol]::Ftp
        HostName = $FTPHostname
        UserName = $FTPUserName
        Password = $FTPPassword
    }

 
    $session = New-Object WinSCP.Session
 
    # Connect
    $session.Open($sessionOptions)
                  
    # Upload files, collect results
    $transferResult = $session.PutFiles($LocalPath, $RemotePath)
            
    # Iterate over every transfer
        foreach ($transfer in $transferResult.Transfers)
        {
            # Success or error?
            if ($transfer.Error -eq $Null)
            {
                #Write-Host "Upload of $($transfer.FileName) succeeded, moving to backup"
                #Add-content $Logfile -value "$LogFileTime - Upload of $($transfer.FileName) succeeded, moving to backup"
                CreateEventLog("$LogFileTime - Upload of $($transfer.FileName) succeeded, moving to backup")
                # Upload succeeded, move source file to backup
                Move-Item $transfer.FileName $BackupPath
                $ItemsUploaded += 1
            }
            else
            {
                #Write-Host "Upload of $($transfer.FileName) failed: $($transfer.Error.Message)"
                #Add-content $Logfile -value "$LogFileTime - Upload of $($transfer.FileName) failed: $($transfer.Error.Message)"
                CreateEventLog("$LogFileTime - Upload of $($transfer.FileName) failed: $($transfer.Error.Message)")
            }
        } 
        
        # Disconnect, clean up
        $session.Dispose()        
    
    If ($ItemsUploaded -eq 0) {
            #Add-content $Logfile -value "$LogFileTime - No files found"
            #write-host "No files found"    
            CreateEventLog("Upload completed no files found.")
        
    } else {
        CreateEventLog("Upload completed, $ItemsUploaded files uploaded.")
    }
}

CreateEventLog("FTP Upload services started")
while ($true) {
    StartUpload
    Start-Sleep -Seconds (60*15)
}