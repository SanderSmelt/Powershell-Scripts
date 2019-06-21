# Powershell-Scripts
set of small Powershell GUI tool and Powershell scripts

# Get-DateFromWeekNumber.psm1

NAME

 Get-DateFromWeekNumber

SYNOPSIS

 Get the date of the specified day, week and year

SYNTAX

 Get-DateFromWeekNumber [-week] \<Int32\> [[-Year] \<Int32\>] [[-Day] \<String\>]

DESCRIPTION

 Get the date of the specified day, week and year

PARAMETERS

 -week <Int32>
 
     specifie number of the week between 1 and 52

     Required?                    true
     Position?                    1
     Default value                0
     Accept pipeline input?       false
     Accept wildcard characters?  false

 -Year <Int32>
 
     Specifie the year. the current year is the default.

     Required?                    false
     Position?                    2
     Default value                ((Get-Date).year)
     Accept pipeline input?       false
     Accept wildcard characters?  false

 -Day <String>
 
     Specifie the day. the current day is the default.

     Required?                    false
     Position?                    3
     Default value                ((Get-Date).DayOfWeek)
     Accept pipeline input?       false
     Accept wildcard characters?  false

INPUTS

 None. You cannot pipe objects to get-datefromweeknumber.

OUTPUTS

 System.DateTime object of the requested day

NOTES

 Filename: Get-DateFromWeekNumber.psm1
 
 Version: 1.0
 
 Author: Sander Smelt
 
 Creation Date: 21-06-2019
 

 -------------------------- EXAMPLE 1 --------------------------

 PS>get-datefromweeknumber -week 1
 
 Friday, January 4, 2019 12:43:42 PM

 -------------------------- EXAMPLE 2 --------------------------

 PS>get-datefromweeknumber -week 1 -day "monday"
 
 Monday, December 31, 2018 12:43:42 PM

 -------------------------- EXAMPLE 3 --------------------------

 PS>get-datefromweeknumber -week 10 -day "wednesday" -year 2020
 
 Wednesday, March 4, 2020 12:43:42 PM

RELATED LINKS

 https://github.com/SanderSmelt

# Sync-PublicFolderContacts.ps1:

![Alt text](/Screenshot.png?raw=true)

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

