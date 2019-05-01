Set-StrictMode -Version 2.0

$host.ui.RawUI.WindowTitle = "AD Picture Fiddler"

$ADPhotoDirectory = "C:\Users\$env:UserName\ADPictureFiddler"
$counter = 0
$usernamelistholder = $NULL

function GetUserNameListNoThumbnail($search){
    return Get-ADUser -Filter * -SearchBase $search -Properties thumbnailPhoto | Where-Object {$_.thumbnailPhoto -eq $Null}
}

function GetAllUserNameList($search){
    return Get-ADUser -Filter * -SearchBase $search
}

function GetUserPhotoFromOffice365($userprincipalname, $userloginname, $loginname, $loginpass){
    $ImageSize = "96x96"
    $DownloadURL = "https://outlook.office365.com/ews/Exchange.asmx/s/GetUserPhoto?email="+$userprincipalname+"&size=HR"+$ImageSize
    $DownloadPath = $ADPhotoDirectory+"\"+$userprincipalname+" "+$ImageSize+".jpg"

    try{
        $WebClient = New-Object System.Net.WebClient
        $WebClient.Credentials = New-Object System.Net.NetworkCredential($loginname,[Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($loginpass)))
        $WebClient.DownloadFile($DownloadURL,$DownloadPath)
        $ADPhoto = ([System.IO.File]::ReadAllBytes($DownloadPath))
        SET-ADUser $userloginname –add @{thumbnailphoto=$ADPhoto}
        
    }catch{
        Write-Host "Unable to get $userprincipalname image. Server returned: $_"
    }

    if(Test-Path $DownloadPath){
        Remove-Item $DownloadPath -Force
    }
}

function PhotoWorker($userlist){
    $ewsloginname = Read-Host 'What is your email address?'
    $ewspassword = Read-Host 'What is your password?' -AsSecureString

    foreach($user in $userlist){    
        GetUserPhotoFromOffice365 $user.UserPrincipalName $user.SamAccountName $ewsloginname $ewspassword
        $counter++
        Write-Progress -Activity "Fiddling with profile picture $($counter) out of $($userlist.Count)" -status "Working on $($user.Name)"  -percentComplete (($counter / $userlist.Count)*100)
    }

    Write-Progress -Activity "Fiddling with profile pictures" -status "Completed. $($counter) out of $($userlist.Count)"
    Write-Host "Completed. $($counter) out of $($userlist.Count)"
    Write-Host ""
    Read-Host -Prompt "Press Enter to continue..."
}
<####################
#
# Start script
#
####################>

$inputsuccess = $false
while($inputsuccess -ne $true){
    Write-Host "Which action would you like to perform?"
    Write-Host ""
    Write-Host "`t1. Update missing pictures"
    Write-Host "`t2. Update all pictures"
    Write-Host "`t3. Exit"
    Write-Host ""

    $input = Read-Host 'Please enter 1-3'
    switch ($input){ 
        "1" {
            $inputsuccess = $true
            New-Item -ItemType Directory -Force -Path $ADPhotoDirectory
            $usersearchbase = 'SEARCHBASE'
            $usernamelistholder = GetUserNameListNoThumbnail $usersearchbase
            PhotoWorker $usernamelistholder
            break
        }
        "2" {
            $inputsuccess = $true
            New-Item -ItemType Directory -Force -Path $ADPhotoDirectory
            $usersearchbase = 'SEARCHBASE'
            $usernamelistholder = GetAllUserNameList $usersearchbase
            PhotoWorker $usernamelistholder
            break
        }
        "3" {
            $inputsuccess = $true
            break
        }

    }
}
