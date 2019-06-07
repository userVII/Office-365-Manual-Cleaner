# Utility scripts for Office 365
A collection of scripts to aid with Office 365 installs

## [Clean_Office365Install.ps1](https://github.com/userVII/Office-365-Utilities/blob/master/Clean_Office365Install.ps1)
Powershell script to manually clean up Office 365 installations. 
We have found sometimes upgrading from a basic install to a Publisher Access install hangs and needs a good clean/uninstall.

The steps for the uninstall were taken from here and adapted to a PowerShell script:

[Microsoft Article on manual uninstalls](https://support.office.com/en-us/article/manually-uninstall-office-4e2904ea-25c8-4544-99ee-17696bb3027b)

## [Clean_SkypeSIPProfiles.ps1](https://github.com/userVII/Office-365-Utilities/blob/master/Clean_SkypeSIPProfiles.ps1)
Small utility to clean off all Skype SIP files on a PC. Sometimes in a domain environment Skype for Business will
start asking for passwords and won't take credentials. This will fix that.

## [Import_Office365PhotoToAD.ps1](https://github.com/userVII/Office-365-Utilities/blob/master/Import_Office365PhotoToAD.ps1)
Powershell script to get an Office 365 image and insert it into the users AD account

Requires the AD module in powershell which you can get by installing administrative tools. 

[Microsoft AD Tools](https://www.microsoft.com/en-us/download/details.aspx?id=45520)

Should just be able to edit the 'SEARCHBASE' to something like "OU=something,DC=company,DC=com" and go to town :)
