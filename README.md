# Powershell-user-tools
set of small GUI Powershell tool to allow user to complete simple tasks

Sync-PublicFolderContacts.ps1:
<#
.SYNOPSIS
 GUI tool to synchronize Outlook public folder contacts to personal contact folder so that the contacts are available on smartphones
 WARNING: if your personal contact contains a folder with the same name as a public folder it wil be removed!
 
.DESCRIPTION
 collects all Outlook public contact folders and displays them in a list, select the folders you want to synchronize to your personal contacts.
 folders that where synchronized before will be selected at startup, to make sure all changes are synchronized the existing folder will be removed and a new copy will be made.
 folders that are no longer selected will be deleted.
 Outlook needs to be installed and running.
 Only tested with Powershell 5.1 and Office 365 v1810 on Windows Server 2016.
 
.NOTES
 Filename: Sync-PublicFolderContacts.ps1
 Version: 1.0
 Author: Sander Smelt
 Creation Date: 30-11-2018
 WARNING: set RemoveExistingFolder to True if you understand the consequences.
 
.LINK
 https://github.com/SanderSmelt/Powershell-user-tools
 
.EXAMPLE
 C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File .\Sync-PublicFolderContacts.ps1
 use this command in a shortcut
 
.EXAMPLE
 .\Sync-PublicFolderContacts.ps1 -Verbose
#>
