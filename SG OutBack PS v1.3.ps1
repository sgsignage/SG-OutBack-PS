# SG OutBack PS (Stands for Outlook Backup PowerShell) v1.3
# Note - this script must be ran as an Administrator in order for the VSS to work!

# *************** New in v1.3: *****************
# 1. Configurations stored in JSON file on remote server!
# **********************************************


# --- 1. Load Master Configuration File (Contains the majority of this script's configuration) ---
$ConfigPath = "\\YOURSERVERNAMEHERE\YOURSHARENAMEHERE\OutBack_Config.json"

try {
    if (!(Test-Path $ConfigPath)) { throw "Master config file not found at $ConfigPath" }
    $Config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
} catch {
    Write-Error "CRITICAL: Failed to load config: $($_.Exception.Message)"
    exit
}


# --- 2. Initialize Variables & Tracking ---
$Workstation = $env:COMPUTERNAME
$Date = Get-Date -Format $Config.Zip_Config.DateFormat
$FilePattern = $Config.Copy7ZipFileToServer_Config.FilePattern -f $Workstation
$ZipFileName = "$Workstation-Outlook-$Date.7z"
$ZipLocalPath = Join-Path $Config.Copy7ZipFileToServer_Config.SourcePath $ZipFileName

$ScriptStatus = "SUCCESS"
$ErrorLog = @()
$TotalStopwatch = [System.Diagnostics.Stopwatch]::StartNew() # We're using this to time the whole process for the email report - Start The Stopwatch!



try {

    # ******************************* SECTION 1 - VSS Activation *********************************
    Write-Host "--- Starting VSS Snapshot ---" -ForegroundColor Cyan
    if (Test-Path $Config.VSS_Config.ShadowLink) { 
        Remove-Item -Path $Config.VSS_Config.ShadowLink -Recurse -Force -Confirm:$false 
    }

    $VSS = Invoke-CimMethod -ClassName Win32_ShadowCopy -MethodName Create -Arguments @{
        Volume = $Config.VSS_Config.SourceDrive
        Context = "ClientAccessible"
    }
    
    if ($VSS.ReturnValue -ne 0) { throw "VSS Snapshot creation failed with code $($VSS.ReturnValue)" }
    
    $VSS_Vol = (Get-CimInstance -ClassName Win32_ShadowCopy -Filter "ID = '$($VSS.ShadowID)'").DeviceObject
    cmd /c mklink /d $Config.VSS_Config.ShadowLink "$VSS_Vol\"

    if (!(Test-Path $Config.VSS_Config.TempDestination)) { New-Item -ItemType Directory -Path $Config.VSS_Config.TempDestination -Force }
    
    # Copy from Shadow Link to local staging
    robocopy "$($Config.VSS_Config.ShadowLink)\$($Config.VSS_Config.SourceFolder)" "$($Config.VSS_Config.TempDestination)" /E /R:3 /W:5 /MT:32
    
    # Cleanup Snapshot
    Remove-Item -Path $Config.VSS_Config.ShadowLink -Recurse -Force -Confirm:$false
    Get-CimInstance -ClassName Win32_ShadowCopy -Filter "ID = '$($VSS.ShadowID)'" | Remove-CimInstance




    # ************************** SECTION 2 - 7-Zip Compression **************************
    Write-Host "--- Starting Compression ---" -ForegroundColor Cyan
    if (!(Test-Path $Config.Zip_Config.ZipExe)) { throw "7-Zip not found at $($Config.Zip_Config.ZipExe)" }

    $7zArgs = "a -t7z `"$ZipLocalPath`" `"$($Config.Zip_Config.SourceDir)`" -mx1 -mmt1 -ms=off"
    $7zProc = Start-Process -FilePath $Config.Zip_Config.ZipExe -ArgumentList $7zArgs -Wait -PassThru -WindowStyle Hidden
    
    if ($7zProc.ExitCode -ne 0) { throw "7-Zip failed with exit code $($7zProc.ExitCode)" }




    # ************************** SECTION 3 - Server Copy & Retention **************************
    Write-Host "--- Managing Server Copy & Retention ---" -ForegroundColor Cyan
    
    if (!(Test-Path $Config.Copy7ZipFileToServer_Config.DestinationPath)) { 
        throw "Server destination path $($Config.Copy7ZipFileToServer_Config.DestinationPath) is unreachable." 
    }

    # A. Execute Copy to Server
    robocopy $Config.Copy7ZipFileToServer_Config.SourcePath $Config.Copy7ZipFileToServer_Config.DestinationPath $ZipFileName /R:3 /W:5
    if ($LASTEXITCODE -ge 8) { throw "Robocopy to server failed with exit code $LASTEXITCODE" }

    # B. Manage Retention on Server
    $ServerFiles = Get-ChildItem -Path $Config.Copy7ZipFileToServer_Config.DestinationPath -Filter $FilePattern | Sort-Object LastWriteTime
    if ($ServerFiles.Count -gt $Config.Copy7ZipFileToServer_Config.MaxFileCount) {
        $NumToDelete = $ServerFiles.Count - $Config.Copy7ZipFileToServer_Config.MaxFileCount
        $ServerFiles | Select-Object -First $NumToDelete | Remove-Item -Force
        Write-Host "Retention: Cleaned up oldest backups." -ForegroundColor Gray
    }

    # C. Final Local Cleanup
    # 1. Delete the temporary .7z file
    if (Test-Path $ZipLocalPath) { Remove-Item $ZipLocalPath -Force }
    
    # 2. Delete the uncompressed staging files in C:\Backups\Outlook
    if (Test-Path $Config.VSS_Config.TempDestination) {
        Get-ChildItem -Path $Config.VSS_Config.TempDestination | Remove-Item -Recurse -Force
        Write-Host "Local Staging: Cleared staging folder." -ForegroundColor Gray
    }

    Write-Host "Backup process successful." -ForegroundColor Green

} catch {
    $ScriptStatus = "FAILURE"
    $ErrorLog += $_.Exception.Message
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
}




# ******************************************** SECTION 4 - Email Report **********************************************
$TotalStopwatch.Stop()
$Duration = "{0:hh\:mm\:ss}" -f $TotalStopwatch.Elapsed

$Em_Subject = "SG Outlook Backup: $ScriptStatus ($Workstation)"

# Using a 'here string' to build the body (@" ... "@)
$Em_Body = @"
SG OutBack PS v1.3 Report
--------------------------------------
Workstation:  $Workstation
Status:       $ScriptStatus
Duration:     $Duration
Date:         $(Get-Date)

Details:
$(if ($ScriptStatus -eq "FAILURE") { "Errors encountered:`n" + ($ErrorLog -join "`n") } else { "The backup was completed and moved to the server successfully." })
--------------------------------------
"@

try {
    $SMTP = New-Object Net.Mail.SmtpClient($Config.Email_Config.SmtpServer, $Config.Email_Config.SmtpPort)
    $SMTP.EnableSsl = $true
    $SMTP.Credentials = New-Object System.Net.NetworkCredential($Config.Email_Config.Username, $Config.Email_Config.Password)
    $Msg = New-Object Net.Mail.MailMessage($Config.Email_Config.From, $Config.Email_Config.To, $Em_Subject, $Em_Body)
    $SMTP.Send($Msg)
    Write-Host "Email report sent." -ForegroundColor Gray
} catch {
    Write-Warning "Could not send email report: $($_.Exception.Message)"
}



# All done!
