<#
.SYNOPSIS
Backup copy script.

.DESCRIPTION
Script with gui for making backup copies.
Script uses hash values to detect changes in files. That is because depending on what files are being copied, Windows
might have changed Last Modified Date, Creation Date or Last Accessed Date on file during update. Hash value will tell the truth
about changes. While copying large folders script can be laggy but will recover, eventually.
#>

Add-Type -assembly System.Windows.Forms
Add-Type -AssemblyName PresentationCore,PresentationFramework

#Käyttöliittymän elementtien määrittely alkaa.
$main_form = New-Object System.Windows.Forms.Form
$main_form.SizeGripStyle = 'Hide'
$main_form.StartPosition = 'CenterScreen'
$main_form.Text ='Backup Copy'
$main_form.BackColor = '#002456'
$main_form.ForeColor = 'Black'

$main_form.AutoSize = $true
$main_form.AutosizeMode = 'GrowAndShrink'
$main_form.MaximizeBox = $false

$formWidthSetter = New-Object System.Windows.Forms.label
$formWidthSetter.Location = New-Object System.Drawing.Size(0,0)
$formWidthSetter.Size = New-Object System.Drawing.Size(640,1)
$formWidthSetter.BackColor = '#002456'

$formHeightSetter = New-Object System.Windows.Forms.label
$formHeightSetter.Location = New-Object System.Drawing.Size(0,0)
$formHeightSetter.Size = New-Object System.Drawing.Size(1,450)
$formHeightSetter.BackColor = '#002456'

$filesTargetLabel = New-Object System.Windows.Forms.label
$filesTargetLabel.Location = New-Object System.Drawing.Size(10,280)
$filesTargetLabel.Size = New-Object System.Drawing.Size(300,15)
$filesTargetLabel.BackColor = '#002456'
$filesTargetLabel.ForeColor = 'Yellow'
$filesTargetLabel.Text = "Target: "

$foldersTargetLabel = New-Object System.Windows.Forms.label
$foldersTargetLabel.Location = New-Object System.Drawing.Size(330,280)
$foldersTargetLabel.Size = New-Object System.Drawing.Size(300,15)
$foldersTargetLabel.BackColor = '#002456'
$foldersTargetLabel.ForeColor = 'Yellow'
$foldersTargetLabel.Text = "Target: "

$outputBoxFiles = New-Object System.Windows.Forms.TextBox
$outputBoxFiles.Location = New-Object System.Drawing.Size(10,60)
$outputBoxFiles.Size = New-Object System.Drawing.Size(300,200)
$outputBoxFiles.BackColor = 'Black'
$outputBoxFiles.ForeColor = 'Yellow'
$outputBoxFiles.MultiLine = $True
$outputBoxFiles.ScrollBars = "Vertical"
$outputBoxFiles.ReadOnly = $True

$outputBoxFolders = New-Object System.Windows.Forms.TextBox
$outputBoxFolders.Location = New-Object System.Drawing.Size(330,60)
$outputBoxFolders.Size = New-Object System.Drawing.Size(300,200)
$outputBoxFolders.BackColor = 'Black'
$outputBoxFolders.ForeColor = 'Yellow'
$outputBoxFolders.MultiLine = $True
$outputBoxFolders.ScrollBars = "Vertical"
$outputBoxFolders.ReadOnly = $True

$ButtonAddFiles = New-Object System.Windows.Forms.Button
$ButtonAddFiles.Location = New-Object System.Drawing.Size(95,20)
$ButtonAddFiles.BackColor = '#00249B'
$ButtonAddFiles.ForeColor = 'Yellow'
$ButtonAddFiles.Size = New-Object System.Drawing.Size(130,20)
$ButtonAddFiles.Text = "Add Files To Copy"

$ButtonAddFolders = New-Object System.Windows.Forms.Button
$ButtonAddFolders.Location = New-Object System.Drawing.Size(415,20)
$ButtonAddFolders.BackColor = '#00249B'
$ButtonAddFolders.ForeColor = 'Yellow'
$ButtonAddFolders.Size = New-Object System.Drawing.Size(130,20)
$ButtonAddFolders.Text = "Add Folders To Copy"

$ButtonTargetForFiles = New-Object System.Windows.Forms.Button
$ButtonTargetForFiles.Location = New-Object System.Drawing.Size(95,310)
$ButtonTargetForFiles.BackColor = '#00249B'
$ButtonTargetForFiles.ForeColor = 'Yellow'
$ButtonTargetForFiles.Size = New-Object System.Drawing.Size(130,20)
$ButtonTargetForFiles.Text = "Set Target For Files"

$ButtonTargetForFolders = New-Object System.Windows.Forms.Button
$ButtonTargetForFolders.Location = New-Object System.Drawing.Size(415,310)
$ButtonTargetForFolders.BackColor = '#00249B'
$ButtonTargetForFolders.ForeColor = 'Yellow'
$ButtonTargetForFolders.Size = New-Object System.Drawing.Size(130,20)
$ButtonTargetForFolders.Text = "Set Target For Folders"

$ButtonCopyFiles = New-Object System.Windows.Forms.Button
$ButtonCopyFiles.Location = New-Object System.Drawing.Size(95,340)
$ButtonCopyFiles.BackColor = '#00249B'
$ButtonCopyFiles.ForeColor = 'Yellow'
$ButtonCopyFiles.Size = New-Object System.Drawing.Size(130,20)
$ButtonCopyFiles.Text = "Copy Files"

$ButtonCopyChangedFiles = New-Object System.Windows.Forms.Button
$ButtonCopyChangedFiles.Location = New-Object System.Drawing.Size(95,370)
$ButtonCopyChangedFiles.BackColor = '#00249B'
$ButtonCopyChangedFiles.ForeColor = 'Yellow'
$ButtonCopyChangedFiles.Size = New-Object System.Drawing.Size(130,20)
$ButtonCopyChangedFiles.Text = "Copy Changed Files"

$ButtonCopyFolders = New-Object System.Windows.Forms.Button
$ButtonCopyFolders.Location = New-Object System.Drawing.Size(415,340)
$ButtonCopyFolders.BackColor = '#00249B'
$ButtonCopyFolders.ForeColor = 'Yellow'
$ButtonCopyFolders.Size = New-Object System.Drawing.Size(130,20)
$ButtonCopyFolders.Text = "Copy Folders"

$ButtonCopyChangedFolders = New-Object System.Windows.Forms.Button
$ButtonCopyChangedFolders.Location = New-Object System.Drawing.Size(415,370)
$ButtonCopyChangedFolders.BackColor = '#00249B'
$ButtonCopyChangedFolders.ForeColor = 'Yellow'
$ButtonCopyChangedFolders.Size = New-Object System.Drawing.Size(130,20)
$ButtonCopyChangedFolders.Text = "Copy Changed Folders"

$ButtonReset = New-Object System.Windows.Forms.Button
$ButtonReset.Location = New-Object System.Drawing.Size(255,310)
$ButtonReset.BackColor = '#00249B'
$ButtonReset.ForeColor = 'Yellow'
$ButtonReset.Size = New-Object System.Drawing.Size(130,20)
$ButtonReset.Text = "Reset"

$ButtonHelp = New-Object System.Windows.Forms.Button
$ButtonHelp.Location = New-Object System.Drawing.Size(255,340)
$ButtonHelp.BackColor = '#00249B'
$ButtonHelp.ForeColor = 'Yellow'
$ButtonHelp.Size = New-Object System.Drawing.Size(130,20)
$ButtonHelp.Text = "Help"

$ButtonQuit = New-Object System.Windows.Forms.Button
$ButtonQuit.Location = New-Object System.Drawing.Size(255,370)
$ButtonQuit.BackColor = '#00249B'
$ButtonQuit.ForeColor = 'Yellow'
$ButtonQuit.Size = New-Object System.Drawing.Size(130,20)
$ButtonQuit.Text = "Quit"

$Global:progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Name = 'progressBar'
$Global:progressBar.Value = 0
$progressBar.Maximum = 100
$progressBar.Step = 1 
$progressBar.Style="Continuous"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 600
$System_Drawing_Size.Height = 20
$progressBar.Size = $System_Drawing_Size
$Global:progressBar.Visible = $false
$progressBar.ForeColor = 'Yellow'
$progressBar.BackColor= '#00249B'
$progressBar.Left = 20
$progressBar.Top = 410
#Käyttöliittymän elementtien määrittely päättyy.

#Alustetaan globaaleja muuttujia.
$global:selectedFolders = @()
$global:selectedFiles = @()
$global:targetForFolders = $null
$global:targetForFiles = $null

#Funktio joka tarkistaa onko kopioitavaa ja/tai paikkaa minne kopioidaan. Kutsutaan kopiointipainikkeista.
Function readyForCopyFiles {
	if ($global:selectedFiles.Length -eq 0 -and $global:targetForFiles.Length -eq 0) {
		[System.Windows.MessageBox]::Show('Please select files to copy and target for files.')
	} elseif ($global:selectedFiles.Length -eq 0) {
		[System.Windows.MessageBox]::Show('Please select files to copy.')
	} elseif ($global:targetForFiles.Length -eq 0) {
		[System.Windows.MessageBox]::Show('Please select target for files.')
	} else {
		return
	}
}

#Funktio joka tarkistaa onko kopioitavaa ja/tai paikkaa minne kopioidaan. Kutsutaan kopiointipainikkeista.
Function readyForCopyFolders {
	if ($global:selectedFolders.Length -eq 0 -and $global:targetForFolders.Length -eq 0) {
		[System.Windows.MessageBox]::Show('Please select folders to copy and target for folders.')
	} elseif ($global:selectedFolders.Length -eq 0) {
		[System.Windows.MessageBox]::Show('Please select folders to copy.')
	} elseif ($global:targetForFolders.Length -eq 0) {
		[System.Windows.MessageBox]::Show('Please select target for folders.')
	} else {
		return
	}
}

$ButtonAddFiles.Add_Click({
	$global:selectedFiles = @()
	$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Multiselect = $true
    $openFileDialog.filter = 'All files (*.*)| *.*'
    $openFileDialog.initialDirectory = [System.IO.Directory]::SetCurrentDirectory('C:\') #'
    $openFileDialog.initialDirectory = [System.IO.Directory]::GetCurrentDirectory()
    $openFileDialog.title = 'Select Files to Copy'
	$null = $openFileDialog.ShowDialog()
	$lengthFiles = $openFileDialog.FileNames.Length - 1
	
	For ($i=0; $i -le $lengthFiles; $i++) {
		$fileName = $openFileDialog.FileNames[$i]
		$fileNameLength = $fileName.Length
		
		#Varmistetaan ettei lisätä mitään olematonta.
		if ($fileNameLength -gt 0) {
		$outputBoxFiles.Text += $openFileDialog.FileNames[$i]+"`r`n"
		$global:selectedFiles += $openFileDialog.FileNames[$i]
		}
	}
})

$ButtonTargetForFiles.Add_Click({
	$targetFoldernameFiles = New-Object System.Windows.Forms.folderbrowserdialog
	$targetFoldernameFiles.showdialog()
	$filesTargetLabel.Text = "Target: "+$targetFoldernameFiles.SelectedPath
	$global:targetForFiles = $targetFoldernameFiles.SelectedPath
})

$ButtonCopyFiles.Add_Click({
	readyForCopyFiles
	
	if ($global:selectedFiles.Length -gt 0 -and $global:targetForFiles.Length -gt 0) {
		$Global:progressBar.Visible = $true
		$progress = 100 / $global:selectedFiles.Length
		
		$global:selectedFiles | ForEach-Object {
			Copy-Item $_ -Destination $global:targetForFiles -Force
			$copiedFiles += 1
			$Global:progressBar.Value += $progress
		}
		$Global:progressBar.Visible = $false
		$Global:progressBar.Value = 0
		[System.Windows.MessageBox]::Show("Ready. $copiedFiles files copied into $global:targetForFiles")
	}
})

$ButtonCopyChangedFiles.Add_Click({
	readyForCopyFiles
	
	if ($global:selectedFiles.Length -gt 0 -and $global:targetForFiles.Length -gt 0) {
	$filesTargetHashes = @(Get-ChildItem -Path $global:targetForFiles | Get-FileHash) #Haetaan kohdekansiosta Hash Valuet.
	$hashTaulukko = @()
	$totalFiles = $global:selectedFiles.Length
	
	foreach ($item in $filesTargetHashes) { #Muutetaan Hash Valuet merkkijonoiksi. 
		$stringHash = [String]$item.Hash
		$hashTaulukko += $stringHash
	}
	$Global:progressBar.Visible = $true
	$progress = 100 / $global:selectedFiles.Length
	
	foreach ($file in $global:selectedFiles) {
		$fileHash = Get-Item -Path $file | Get-FileHash #Otetaan kopioitavan tiedoston Hash Value.
		$fileHashString = [String]$fileHash.Hash #Muutetaan kopioitavan tiedoston Hash Value merkkijonoksi.
		
		<#Verrataan -notcontains operaattorilla onko kopioitavaa tiedostoa olemassa, tai onko se muuttunut verrattuna olemassaolevaan.
		Jos on muuttunut, niin ylikirjoitetaan ja jos ei ole olemassa niin lisätään.#>
		if ($hashTaulukko -notcontains $fileHashString) { 	
			Copy-Item $file -Destination $global:targetForFiles -Force
			$changedFiles += 1
			$Global:progressBar.Value += $progress
		}	
	}
	$notChangedFiles = $totalFiles - $changedFiles
	
	if ($changedFiles -gt 0) {
		$Global:progressBar.Visible = $false
		$Global:progressBar.Value = 0
		[System.Windows.MessageBox]::Show("
Ready. $changedFiles files had changes or did not exist in $global:targetForFiles
and were copied into $global:targetForFiles.

$notChangedFiles files had no changes and were not copied.
")
	} else {
		$Global:progressBar.Visible = $false
		[System.Windows.MessageBox]::Show("
Ready. No files had changes or did not exist in $global:targetForFiles.

No files copied.
")
	}
	
	}
})

$ButtonAddFolders.Add_Click({
	$foldername = New-Object System.Windows.Forms.folderbrowserdialog
	$foldername.showdialog()
	$path = $foldername.SelectedPath
	
	#Varmistetaan ettei lisätä mitään olematonta.
	if ($path.Length -gt 0) {
	$outputBoxFolders.Text += $foldername.SelectedPath+"`r`n"
	$global:selectedFolders += $foldername.SelectedPath  
	}
})

$ButtonTargetForFolders.Add_Click({
	$targetFoldername = New-Object System.Windows.Forms.folderbrowserdialog
	$targetFoldername.showdialog()
	$foldersTargetLabel.Text = "Target: "+$targetFoldername.SelectedPath
	$global:targetForFolders = $targetFoldername.SelectedPath
})

$ButtonCopyFolders.Add_Click({
	readyForCopyFolders
	
	if ($global:selectedFolders.Length -gt 0 -and $global:targetForFolders.Length -gt 0) {
		$Global:progressBar.Visible = $true
		$progress = 100 / $global:selectedFolders.Length
		$global:selectedFolders | ForEach-Object {
			Copy-Item $_ -Destination $global:targetForFolders -Recurse -Force
			$copiedFolders += 1
			$Global:progressBar.Value += $progress
		}
		$Global:progressBar.Visible = $false
		$Global:progressBar.Value = 0
		[System.Windows.MessageBox]::Show("Ready. $copiedFolders folders copied into $global:targetForFolders")
	}	
})

$ButtonCopyChangedFolders.Add_Click({
	
	<#Mikäli kohdekansio ei ole tyhjä, muutetaan kohdekansion jokaisesta kansiosta jokaisen tiedoston Hash Value merkkijonoksi ja lisätään taulukkoon.
	Tämän jälkeen jokaisen kopioitavaksi valitun kansion sisällölle tehdään yksitellen sama toiminto ja
	verrataan löytyykö kohdekansion HashValue-taulukosta sama määrä vastaavuuksia kuin kopioitavassa
	kansiossa on tiedostoja. Mikäli ei löydy, kansio kopioidaan. Tässä ei siis kopioida yksittäisiä tiedostoja,
	vaan mikäli yhdessäkin kansion tiedostossa on muutoksia niin koko kansio kopioidaan. Mikäli kohdekansio on tyhjä,
	kopioidaan suoraan.#>
	
	readyForCopyFolders	
	
	if ($global:selectedFolders.Length -gt 0 -and $global:targetForFolders.Length -gt 0) {
	$targetFolderContent = Get-ChildItem $global:targetForFolders -File -Recurse | Get-FileHash
	
	#Funktio jolla käsitellään tilanne jos kohdekansiossa ei ole kansioita.
	Function EmptyTarget {
		$Global:progressBar.Visible = $true
		$progress = 100 / $global:selectedFolders.Length
		
		$global:selectedFolders | ForEach-Object {
			Copy-Item $_ -Destination $global:targetForFolders -Recurse -Force
			$toEmptyTarget += 1
			$Global:progressBar.Value += $progress
		}
		$Global:progressBar.Visible = $false
		$Global:progressBar.Value = 0
		[System.Windows.MessageBox]::Show("$global:targetForFolders was empty. $toEmptyTarget folders copied to $global:targetForFolders")
	}
	
	#Funktio jolla käsitellään tilanne jos kohdekansiossa on kansioita.
	Function NotSoEmptyTarget {
		$totalFolders = $global:selectedFolders.Length
		foreach ($item in $targetFolderContent) {
			$stringHash = [String]$item.Hash
			$folderHashTaulukko += @($stringHash)
		}
		
		foreach ($folder in $global:selectedFolders) {
			$selectedFolderHashes = Get-ChildItem $folder -File -Recurse | Get-FileHash
			
			foreach ($item2 in $selectedFolderHashes) {
				$stringHashSelected = [String]$item2.Hash
				$stringHashSelectedTaulukko += @($stringHashSelected)
			}
			
			$stringHashSelectedTaulukko | ForEach-Object {
				if ($folderHashTaulukko -contains $_) {
				$contains += 1
				}
			}
			
			if ($contains -ne $stringHashSelectedTaulukko.Length) {
				$Global:progressBar.Visible = $true
				$progress = 100 / $global:selectedFolders.Length
				Copy-Item $folder -Destination $global:targetForFolders -Recurse -Force
				$copiedFolders += 1
				$Global:progressBar.Value += $progress	
			}
			
			$stringHashSelectedTaulukko = @()
			$contains = 0
		}
		$notCopiedFolders = $totalFolders - $copiedFolders
		
		if ($copiedFolders-gt 0) {
			$Global:progressBar.Visible = $false
			$Global:progressBar.Value = 0
			[System.Windows.MessageBox]::Show("
Ready. $copiedFolders folders had changes or did not exist in $global:targetForFolders
and were copied into $global:targetForFolders.

$notCopiedFolders folders had no changes and were not copied.
		")
		} else {
		$Global:progressBar.Visible = $false
		[System.Windows.MessageBox]::Show("
Ready. No folders had changes or did not exist in $global:targetForFolders.

No folders copied.
		")
	}
	}
	
	#Tarkistetaan onko kohdekansio tyhjä vai ei, ja kutsutaan tilanteeseen sopivaa funktiota.
	if ($targetFolderContent -eq $null) {
		EmptyTarget
	} else {
		NotSoEmptyTarget
	}
	}
})

$ButtonReset.Add_Click({
	$global:selectedFolders = @()
	$global:selectedFiles = @()
	$global:targetForFolders = $null
	$global:targetForFiles = $null
	$filesTargetLabel.Text = "Target: "
	$foldersTargetLabel.Text = "Target: "
	$outputBoxFiles.Text = ""
	$outputBoxFolders.Text = ""
	$Global:progressBar.Visible = $false
	$Global:progressBar.Value = 0
})

$ButtonHelp.Add_Click({
	[System.Windows.MessageBox]::Show(
'
Add Files To Copy
 - Choose files you want to copy. Accepts files only.

Add Folders To Copy
 - Choose folders you want to copy.
   No multiselect so you must select multiple
   folders individually. Accepts folders only.
   
Set Target For Files
 - Choose where to copy selected files.
	
Set Target For Folders
 - Choose where to copy selected folders.
	
Copy Files
 - Copy selected files even if they exist in
   target folder or have not changed.
   
Copy Changed Files
 - Copy only files that have changed compared to files
   in target folder. Also if file to copy does not exist in 
   target folder, it will be copied.

Copy Folders
 - Copy selected folder/folders even if they exist in
   target folder or have not changed.
   
Copy Changed Folders
 - Copy only folders that have changed compared to folders
   in target folder. In other words, if any file inside 
   folder to copy has changed, whole folder will be copied.   
   Also if folder to copy does not exist in target folder, it 
   will be copied.
   
Reset
 - Resets everything.
 
Help
 - Displays this help.
 
Quit
 - Exit script.
')
})

$ButtonQuit.Add_Click({
	$main_form.Close()
})

#Lisätään käyttöliittymän elementtien toiminnot lomakkeelle.
$main_form.Controls.Add($formWidthSetter)
$main_form.Controls.Add($formHeightSetter)
$main_form.Controls.Add($filesTargetLabel)
$main_form.Controls.Add($foldersTargetLabel)
$main_form.Controls.Add($OutputBoxFiles)
$main_form.Controls.Add($OutputBoxFolders)
$main_form.Controls.Add($ButtonAddFiles)
$main_form.Controls.Add($ButtonAddFolders)
$main_form.Controls.Add($ButtonTargetForFiles)
$main_form.Controls.Add($ButtonTargetForFolders)
$main_form.Controls.Add($ButtonCopyFiles)
$main_form.Controls.Add($ButtonCopyChangedFiles)
$main_form.Controls.Add($ButtonCopyFolders)
$main_form.Controls.Add($ButtonCopyChangedFolders)
$main_form.Controls.Add($ButtonReset)
$main_form.Controls.Add($ButtonHelp)
$main_form.Controls.Add($ButtonQuit)
$main_form.controls.add($progressBar)
$main_form.ShowDialog()