if (Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell) {
    Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
	Write-Host "Le module 'Microsoft.Online.SharePoint.PowerShell' est déjà installé et a été importé"
} else {
    Install-Module Microsoft.Online.SharePoint.PowerShell -Confirm:$false -Force
    Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
	Write-Host "Le module 'Microsoft.Online.SharePoint.PowerShell' a été installé et importé"
}

if (Get-Module -ListAvailable -Name SharePointPnPPowerShellOnline) {
    Import-Module SharePointPnPPowerShellOnline
	Write-Host "Le module 'SharePointPnPPowerShellOnline' est déjà installé et a été importé"
} else {
    Install-Module SharePointPnPPowerShellOnline -Confirm:$false -Force
    Import-Module SharePointPnPPowerShellOnline
	Write-Host "Le module 'SharePointPnPPowerShellOnline' a été installé et importé"
}

if (Get-Module -ListAvailable -Name MSOnline) {
    Import-Module MSOnline
	Write-Host "Le module 'MSOnline' est déjà installé et a été importé"
} else {
    Install-Module MSOnline -Confirm:$false -Force
	Import-Module MSOnline
	Write-Host "Le module 'MSOnline' a été installé et importé"
}

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form = New-Object System.Windows.Forms.Form
$Form.ClientSize = '600,250'
$form.MinimumSize = '600, 250'
$form.MaximumSize = '600, 250'
$Form.FormBorderStyle = 'Fixed3D'
$Form.Text = "Transfert de donnees entre OneDrive"

$Textbox_login_office365 = New-Object System.Windows.Forms.TextBox
$Textbox_login_office365.Location = New-Object System.Drawing.Point(160,35)
$Textbox_login_office365.Width = 300
$Label_login_office365 = New-Object System.Windows.Forms.Label
$Label_login_office365.Location = New-Object System.Drawing.Point(30,35)
$Label_login_office365.Width = 120
$Label_login_office365.Text = "Utilisateur :"
$Form.controls.AddRange(@($Label_login_office365,$Textbox_login_office365))

$TextBox_password_office365 = New-Object System.Windows.Forms.TextBox
$TextBox_password_office365.Location = New-Object System.Drawing.Point(160,65)
$TextBox_password_office365.Width = 300
$TextBox_password_office365.PasswordChar = '*'
$Label_password_office365 = New-Object System.Windows.Forms.Label
$Label_password_office365.Location = New-Object System.Drawing.Point(30,65)
$Label_password_office365.Width = 120
$Label_password_office365.Text = "Mot de passe :"
$Form.controls.AddRange(@($Label_password_office365,$Textbox_password_office365))

$GroupBoxLogin = New-Object System.Windows.Forms.GroupBox
$GroupBoxLogin.Location = New-Object System.Drawing.Point(10,10)
$GroupBoxLogin.Width = 470
$GroupBoxLogin.Height = 90
$GroupBoxLogin.Text = "Login/Password de l'administrateur du tenant Office365"

$Textbox_onedrive_origine = New-Object System.Windows.Forms.TextBox
$Textbox_onedrive_origine.Location = New-Object System.Drawing.Point(160,135)
$Textbox_onedrive_origine.Width = 300
$Label_onedrive_origine = New-Object System.Windows.Forms.Label
$Label_onedrive_origine.Location = New-Object System.Drawing.Point(30,135)
$Label_onedrive_origine.Width = 120
$Label_onedrive_origine.Text = "Email d'origine :"
$Form.controls.AddRange(@($Label_onedrive_origine,$Textbox_onedrive_origine))

$TextBox_onedrive_destination = New-Object System.Windows.Forms.TextBox
$TextBox_onedrive_destination.Location = New-Object System.Drawing.Point(160,165)
$TextBox_onedrive_destination.Width = 300
$Label_onedrive_destination = New-Object System.Windows.Forms.Label
$Label_onedrive_destination.Location = New-Object System.Drawing.Point(30,165)
$Label_onedrive_destination.Width = 120
$Label_onedrive_destination.Text = "Email de destination :"
$Form.controls.AddRange(@($Label_onedrive_destination,$Textbox_onedrive_destination))

$GroupBoxOnedrive = New-Object System.Windows.Forms.GroupBox
$GroupBoxOnedrive.Location = New-Object System.Drawing.Point(10,110)
$GroupBoxOnedrive.Width = 470
$GroupBoxOnedrive.Height = 90
$GroupBoxOnedrive.Text = "Emails d'origine et de destination"
$Form.controls.AddRange(@($GroupBoxOnedrive,$GroupBoxLogin))

$Bouton = New-Object System.Windows.Forms.Button
$Bouton.Location = New-Object System.Drawing.Point(490,15)
$Button.ButtonBorderStyle = 'Fixed3D'
$Bouton.Width = '80'
$Bouton.Height = '185'
$Bouton.Text = "Valider"
$Form.controls.Add($Bouton)

$btnSubmit_Click={

	if(($Textbox_login_office365.Text -ne '') -and ($Textbox_password_office365.Text -ne '') -and ($Textbox_onedrive_origine.Text -ne '') -and ($TextBox_onedrive_destination.Text -ne '')) {
		$departinguser = $Textbox_onedrive_origine.Text
		$destinationuser = $TextBox_onedrive_destination.Text
		$globaladmin = $Textbox_login_office365.Text
		#"adm.netoptima@cabrini.fr"
		$password = $Textbox_password_office365.Text
		#"223C!FDj6vGXc%@4FP5k"
		$secstr = New-Object -TypeName System.Security.SecureString
		$password.ToCharArray() | ForEach-Object {$secstr.AppendChar($_)}
		$credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $globaladmin, $secstr
		Connect-MsolService -Credential $credentials
		 
		$InitialDomain = Get-MsolDomain | Where-Object {$_.IsInitial -eq $true}
		  
		$SharePointAdminURL = "https://$($InitialDomain.Name.Split(".")[0])-admin.sharepoint.com"
		  
		$departingUserUnderscore = $departinguser -replace "[^a-zA-Z,0-9]", "_"
		$destinationUserUnderscore = $destinationuser -replace "[^a-zA-Z,0-9]", "_"
		$departingOneDriveSite = "https://$($InitialDomain.Name.Split(".")[0])-my.sharepoint.com/personal/$departingUserUnderscore"
		write-host "Site SharePoint de $departinguser : "$departingOneDriveSite
		$destinationOneDriveSite = "https://$($InitialDomain.Name.Split(".")[0])-my.sharepoint.com/personal/$destinationUserUnderscore"
		write-host "Site SharePoint de $destinationuser : "$destinationOneDriveSite
		
		Write-Host "`nConnecting to SharePoint Online" -ForegroundColor Blue
		Connect-SPOService -Url $SharePointAdminURL -Credential $credentials
		  
		Write-Host "`nAdding $globaladmin as site collection admin on both OneDrive site collections" -ForegroundColor Blue
		# Set current admin as a Site Collection Admin on both OneDrive Site Collections
		Set-SPOUser -Site $departingOneDriveSite -LoginName $globaladmin -IsSiteCollectionAdmin $true
		Set-SPOUser -Site $destinationOneDriveSite -LoginName $globaladmin -IsSiteCollectionAdmin $true
		  
		Write-Host "`nConnecting to $departinguser's OneDrive via SharePoint Online PNP module" -ForegroundColor Blue
		  
		Connect-PnPOnline -Url $departingOneDriveSite -Credentials $credentials
		  
		Write-Host "`nGetting display name of $departinguser" -ForegroundColor Blue
		# Get name of departing user to create folder name.
		$departingOwner = Get-PnPSiteCollectionAdmin | Where-Object {$_.loginname -match $departinguser}
		  
		# If there's an issue retrieving the departing user's display name, set this one.
		if ($departingOwner -contains $null) {
			$departingOwner = @{
				Title = "Departing User"
			}
		}
		  
		# Define relative folder locations for OneDrive source and destination
		$departingOneDrivePath = "/personal/$departingUserUnderscore/Documents"
		$destinationOneDrivePath = "/personal/$destinationUserUnderscore/Documents/$($departingOwner.Title)'s Files"
		$destinationOneDriveSiteRelativePath = "Documents/$($departingOwner.Title)'s Files"
		  
		Write-Host "`nGetting all items from $($departingOwner.Title)" -ForegroundColor Blue
		# Get all items from source OneDrive
		$items = Get-PnPListItem -List Documents -PageSize 1000
		  
		$largeItems = $items | Where-Object {[long]$_.fieldvalues.SMTotalFileStreamSize -ge 261095424 -and $_.FileSystemObjectType -contains "File"}
		if ($largeItems) {
			$largeexport = @()
			foreach ($item in $largeitems) {
				$largeexport += "$(Get-Date) - Size: $([math]::Round(($item.FieldValues.SMTotalFileStreamSize / 1MB),2)) MB Path: $($item.FieldValues.FileRef)"
				Write-Host "File too large to copy: $($item.FieldValues.FileRef)" -ForegroundColor DarkYellow
			}
			$largeexport | Out-file C:\temp\largefiles.txt -Append
			Write-Host "A list of files too large to be copied from $($departingOwner.Title) have been exported to C:\temp\LargeFiles.txt" -ForegroundColor Yellow
		}
		  
		$rightSizeItems = $items | Where-Object {[long]$_.fieldvalues.SMTotalFileStreamSize -lt 261095424 -or $_.FileSystemObjectType -contains "Folder"}
		  
		Write-Host "`nConnecting to $destinationuser via SharePoint PNP PowerShell module" -ForegroundColor Blue
		Connect-PnPOnline -Url $destinationOneDriveSite -Credentials $credentials
		  
		Write-Host "`nFilter by folders" -ForegroundColor Blue
		# Filter by Folders to create directory structure
		$folders = $rightSizeItems | Where-Object {$_.FileSystemObjectType -contains "Folder"}
		  
		Write-Host "`nCreating Directory Structure" -ForegroundColor Blue
		foreach ($folder in $folders) {
			$path = ('{0}{1}' -f $destinationOneDriveSiteRelativePath, $folder.fieldvalues.FileRef).Replace($departingOneDrivePath, '')
			Write-Host "Creating folder in $path" -ForegroundColor Green
			$newfolder = Resolve-PnPFolder -SiteRelativePath $path
		}
		  
		 
		Write-Host "`nCopying Files" -ForegroundColor Blue
		$files = $rightSizeItems | Where-Object {$_.FileSystemObjectType -contains "File"}
		$fileerrors = ""
		foreach ($file in $files) {
			  
			$destpath = ("$destinationOneDrivePath$($file.fieldvalues.FileDirRef)").Replace($departingOneDrivePath, "")
			Write-Host "Copying $($file.fieldvalues.FileLeafRef) to $destpath" -ForegroundColor Green
			$newfile = Copy-PnPFile -SourceUrl $file.fieldvalues.FileRef -TargetUrl $destpath -OverwriteIfAlreadyExists -Force -ErrorVariable errors -ErrorAction SilentlyContinue
			$fileerrors += $errors
		}
		$fileerrors | Out-File E:\temp\fileerrors.txt
		  
		# Remove Global Admin from Site Collection Admin role for both users
		Write-Host "`nRemoving $globaladmin from OneDrive site collections" -ForegroundColor Blue
		Set-SPOUser -Site $departingOneDriveSite -LoginName $globaladmin -IsSiteCollectionAdmin $false
		Set-SPOUser -Site $destinationOneDriveSite -LoginName $globaladmin -IsSiteCollectionAdmin $false
		Write-Host "`nComplete!" -ForegroundColor Green
	}
}

$Bouton.add_Click($btnSubmit_Click)

$Form.ShowDialog()