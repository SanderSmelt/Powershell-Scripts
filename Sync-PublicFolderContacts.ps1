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

[cmdletbinding()]
Param()

#Region Begin Vars{ 
$OutlookExe = "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"
$RemoveExistingFolder = $False
#translations for user text here
$Formtiteltext = "Synchronize public folder contacts"
$ButtonSynchronizeText = "Synchronize"
$ButtonCloseText = "Close"
$LabelInfoText = "Synchronize public folder contacts with personal contacts so they are available on your smartphone.`nSelect the folders with contacts you want to synchronize. Deselected folders will be removed."
$PopupText = "Synchronisation completed"
$PopupTitel = "Info"
#EndRegion Vars }

#Region Begin Loading{
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
Write-Verbose "check if outlook is running if not start it"
if ( -not(get-process outlook | Where-Object {$_.MainWindowTitle -ne ""})){
	& $Outlookexe
}
$Outlook = New-Object -com Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
Write-Verbose "Opening Publicfolders,personal contact folders and deleted items folder"
$PublicFolders = $Namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olPublicFoldersAllPublicFolders)
$ContactFolders = $Namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderContacts)
$DeletedItemsFolders = $Namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderDeletedItems)
#EndRegion Loading }

#Region Begin GUI{ 
$Form = New-Object system.Windows.Forms.Form
$Form.ClientSize = New-Object System.Drawing.Size(380,380)
$Form.text = $FormtitelText
$Form.MaximizeBox = $False
$Form.FormBorderStyle = "Fixed3D"
$Form.Topmost = $True
$Form.Icon = [system.drawing.icon]::ExtractAssociatedIcon($OutlookExe)

$ButtonSynchronize = New-Object system.Windows.Forms.Button
$ButtonSynchronize.text = $ButtonSynchronizeText
$ButtonSynchronize.Size = New-Object System.Drawing.Size(120,30)
$ButtonSynchronize.location = New-Object System.Drawing.Point(250,340)

$ButtonClose = New-Object system.Windows.Forms.Button
$ButtonClose.text = $ButtonCloseText
$ButtonClose.Size = New-Object System.Drawing.Size(120,30)
$ButtonClose.location = New-Object System.Drawing.Point(120,340)

$LabelInfo = New-Object system.Windows.Forms.Label
$LabelInfo.Font = "Microsoft Sans Serif,10"
$LabelInfo.MaximumSize = New-Object System.Drawing.Size(360,0)
$LabelInfo.location = New-Object System.Drawing.Point(10,10)
$LabelInfo.text = $LabelInfoText
$LabelInfo.AutoSize = $True

$CheckedListBoxPublicFolders = New-Object system.Windows.Forms.CheckedListBox
Write-Verbose "Filling Checkboxlist and selecting synched items"
$CheckedListBoxPublicFolders.Items.AddRange(($PublicFolders.folders | Where-Object {$_.defaultitemtype -eq 2} | select addressbookname).addressbookname);
foreach ($Item in ($CheckedListBoxPublicFolders.Items | Where-Object {($ContactFolders.folders | select addressbookname).addressbookname -contains $_ })){
	$CheckedListBoxPublicFolders.SetItemChecked([array]::IndexOf($CheckedListBoxPublicFolders.items,$Item), 1)
}
$CheckedListBoxPublicFolders.Size = New-Object System.Drawing.Size(360,250)
$CheckedListBoxPublicFolders.location = New-Object System.Drawing.Point(10,85)
$CheckedListBoxPublicFolders.CheckOnClick = $True

$Form.controls.AddRange(@($ButtonSynchronize,$ButtonClose,$LabelInfo,$CheckedListBoxPublicFolders))

#Region gui events {
$ButtonSynchronize.add_Click({
	Write-Verbose "Sync checked items"
	$ButtonSynchronize.Enabled = $False
	$ButtonClose.Enabled = $False
	foreach($Item in $CheckedListBoxPublicFolders.CheckedItems){
		if ((($DeletedItemsFolders.folders | select addressbookname).addressbookname -contains $item) -and ($RemoveExistingFolder)){
			$DeletedItemsFolders.folders.Item($Item).delete()
			Write-Verbose "Removing $Item from deleted items folder"
		}
		if ((($ContactFolders.folders | select addressbookname).addressbookname -contains $item) -and ($RemoveExistingFolder)){
			$ContactFolders.folders.Item($Item).delete()
			Write-Verbose "Removing $Item from personal contact folder"
		}
		$PublicFolders.folders.Item($Item).CopyTo($ContactFolders)
		Write-Verbose "Copy $Item to personal contacts folder"
	}
	Write-verbose "Remove unchecked items"	
	foreach ($item in (($CheckedListBoxPublicFolders.Items | Where-Object {$CheckedListBoxPublicFolders.CheckedItems -notcontains $_}) | Where-Object {($DeletedItemsFolders.folders | select addressbookname).addressbookname -contains $_ })){
		if ($RemoveExistingFolder){
			$DeletedItemsFolders.folders.Item($Item).delete()
			Write-Verbose "Removing $Item from deleted items folder"
		}
	}
	foreach ($item in (($CheckedListBoxPublicFolders.Items | Where-Object {$CheckedListBoxPublicFolders.CheckedItems -notcontains $_}) | Where-Object {($ContactFolders.folders | select addressbookname).addressbookname -contains $_ })){
		if ($RemoveExistingFolder){
			$ContactFolders.folders.Item($Item).delete()
			Write-Verbose "Removing $Item from personal contact folder"
		}
	}
	[System.Windows.Forms.MessageBox]::Show($PopupText, $PopupTitel, 0)
	$ButtonSynchronize.Enabled = $True
	$ButtonClose.Enabled = $True
})

$ButtonClose.add_Click({
	$Form.Close()
})
#EndRegion events }
#EndRegion GUI }

$Form.ShowDialog()