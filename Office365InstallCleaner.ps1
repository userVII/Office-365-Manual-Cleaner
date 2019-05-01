<#
    .SYNOPSIS
        Office 365 Manual Cleaner
    .DESCRIPTION
        Testing utility cleaner for when an office full install won't complete
    .NOTES
        Version:  1
        Author:  userVII
        Creation Date:  05-01-2019
        Last Update:  05-01-2019
#>

<############
  Functions
############>
function Kill_Process{
    Write-Host "Attempting to kill known Office Process"
    $processNames = @(
                    "OfficeClickToRun", "OfficeC2RClient", "AppVShNotify", "Setup*",
                    "OUTLOOK", "WORD", "OneDrive", "Lync", "ONENOTEM", "ONENOTE", "EXCEL"
                    )

    foreach($p in $processNames){
        $aliveProcess = Get-Process $p -ErrorAction SilentlyContinue
        if ($aliveProcess) {
            Write-Host "Stopping $p process" -ForegroundColor Yellow
            $aliveProcess.CloseMainWindow() | Out-Null
            Start-Sleep -Seconds 3
            if (!$aliveProcess.HasExited) {
                #using both methods seems to be more reliable
                $aliveProcess | Stop-Process -Force
                Stop-Process -name $p -force
            }
        }else{
            Write-Host "$p wasn't running"
        }
    }
}

function Remove_OneDrive{
    Write-Host "Attempting to remove OneDrive"
    if (Test-Path "$env:systemroot\System32\OneDriveSetup.exe") {
        Start-Process "$env:systemroot\System32\OneDriveSetup.exe" -ArgumentList "/uninstall"
    }
    if (Test-Path "$env:systemroot\SysWOW64\OneDriveSetup.exe") {
        Start-Process "$env:systemroot\SysWOW64\OneDriveSetup.exe" -ArgumentList "/uninstall"
    }
}

function Remove_ScheduledTasks{
    Write-Host "Removing Scheduled Tasks"
    Start-Process schtasks.exe -ArgumentList "/delete /tn /F '\Microsoft\Office\Office Automatic Updates'"
    Start-Process schtasks.exe -ArgumentList "/delete /tn /F '\Microsoft\Office\Office Subscription Maintenance'"
    Start-Process schtasks.exe -ArgumentList "/delete /tn /F '\Microsoft\Office\Office ClickToRun Service Monitor'"
    Start-Process schtasks.exe -ArgumentList "/delete /tn /F '\Microsoft\Office\OfficeTelemetryAgentLogOn2016'"
    Start-Process schtasks.exe -ArgumentList "/delete /tn /F '\Microsoft\Office\OfficeTelemetryAgentFallBack2016'"
}

function Remove_InstallDirectories{
    Write-Host "Removing Install Directories"
    $directoriesToRemove = @(
                            "C:\Program Files\Microsoft Office 15", "C:\Program Files\Microsoft Office 16",
                            "C:\Program Files\Microsoft Office", "C:\Program Files (x86)\Microsoft Office",
                            "C:\Program Files (x86)\Microsoft Office 15", "C:\Program Files (x86)\Microsoft Office 16",
                            "C:\ProgramData\Microsoft\ClickToRun"
                            )
    $filesToRemove = @("C:\ProgramData\Microsoft\Office\ClickToRunPackageLocker")

    foreach($dirRemoval in $directoriesToRemove){
        if(Test-Path $dirRemoval){
            Write-Host "Removing $dirRemoval" -ForegroundColor Yellow
            Remove-Item -LiteralPath $dirRemoval -Force -Recurse
        }else{
            Write-Host "$dirRemoval is already removed"
        }
    }

    foreach($fileRemoval in $filesToRemove){
        if(Test-Path $fileRemoval){
            Write-Host "Removing $fileRemoval" -ForegroundColor Yellow
            Remove-Item -LiteralPath $fileRemoval -Force
        }else{
            Write-Host "$fileRemoval is already removed"
        }
    }
}

function Remove_RegistryItems{
    Write-Host "Removing Registry Items"
    if(Test-Path "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun"){
        Remove-Item -Path "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun" -Recurse -Force
    }

    if(Test-Path "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\AppVISV"){
        Remove-Item -Path "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\AppVISV" -Recurse -Force
    }

    if(Test-Path "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365ProPlusRetail - en-us"){
        Remove-Item -Path "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365ProPlusRetail - en-us" -Recurse -Force
    }

    if(Test-Path "HKEY_CURRENT_USER\Software\Microsoft\Office"){
        Remove-Item -Path "HKEY_CURRENT_USER\Software\Microsoft\Office" -Recurse -Force
    }
}

function Remove_UtilityPrograms{
    Write-Host "Removing Utilities"
    #x64 OS x32 Office
    Start-Process MsiExec.exe -ArgumentList "/X{90160000-008F-0000-1000-0000000FF1CE} /quiet"
    Start-Process MsiExec.exe -ArgumentList "/X{90160000-008C-0000-0000-0000000FF1CE} /quiet"
    Start-Process MsiExec.exe -ArgumentList "/X{90160000-008C-0409-0000-0000000FF1CE} /quiet"
    #x86 OS x86 Office
    Start-Process MsiExec.exe -ArgumentList "/X{90160000-007E-0000-0000-0000000FF1CE} /quiet"
    Start-Process MsiExec.exe -ArgumentList "/X{90160000-008C-0000-0000-0000000FF1CE} /quiet"
    Start-Process MsiExec.exe -ArgumentList "/X{90160000-008C-0409-0000-0000000FF1CE} /quiet"
    #x64 OS x64 Office
    Start-Process MsiExec.exe -ArgumentList "/X{90160000-007E-0000-1000-0000000FF1CE} /quiet"
    Start-Process MsiExec.exe -ArgumentList "/X{90160000-008C-0000-1000-0000000FF1CE} /quiet"
    Start-Process MsiExec.exe -ArgumentList "/X{90160000-008C-0409-1000-0000000FF1CE} /quiet"
}

function main{
    Kill_Process
    Remove_OneDrive
    Remove_ScheduledTasks
    Remove_RegistryItems
    Remove_UtilityPrograms
    Write-Host "Waiting for everything to catch up"
    Start-Sleep -Seconds 5
    Remove_InstallDirectories
}

<############
  Script
############>
main
$pause = Read-Host -Prompt 'Press any button to exit...'
