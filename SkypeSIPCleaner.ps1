<#
    .SYNOPSIS
        Skype SIP Cleaner
    .DESCRIPTION
        Cleans all users on the PC's Skype SIP files. Usually happens when in a domain Skype starts asking for passwords
    .NOTES
        Version:  1
        Author:  userVII
        Creation Date:  05-01-2019
        Last Update:  05-01-2019
#>

$processNames = @(
                    "Lync", "Outlook"
                    )

foreach($p in $processNames){
    $aliveProcess = Get-Process $p -ErrorAction SilentlyContinue
    if ($aliveProcess) {
        Write-Host "Stopping $p process" -ForegroundColor Yellow
        $aliveProcess.CloseMainWindow() | Out-Null
        # If the kill wasn't graceful
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

$users = Get-ChildItem -Path "C:\Users"

foreach($user in $users){
    Write-Host "Working on $user"
    $SIPPattern = "sip_$($user)*"
    $SIPPath = Get-ChildItem -Path "C:\Users\$user\AppData\Local\Microsoft\Office\16.0\Lync\" -Filter $SIPPattern -Recurse | Where-Object {$_.PSIsContainer} | Select-Object -ExpandProperty Fullname

    if(Test-Path $SIPPath){
        Write-Host "Found Skype SIP file"   
        Remove-Item -Path $SIPPath -Force -Recurse
    }

    if(Test-Path "C:\Users\$user\AppData\Local\Microsoft\Office\16.0\Lync\Tracing"){
        Write-Host "Deleting Tracing files"
        Remove-Item "C:\Users\$user\AppData\Local\Microsoft\Office\16.0\Lync\Tracing\*.*" | Where { ! $_.PSIsContainer }
    }
}

Write-Host "Any file deletion errors are probably from the skype plugin and Outlook being open. Should be fine to reopen Skype now on the users PC"
$pause = Read-Host -Prompt "Press any button to exit..."
