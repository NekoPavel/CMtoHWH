Param(
    [string]$currentUserDirectory
)
function fkarfinder {
    param($pcName)
    function Find-ADObjects($attributes = "distinguishedName") {
        $dc = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext ([System.DirectoryServices.ActiveDirectory.DirectoryContextType]"domain", "gaia");
        $dn = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($dc);
        $ds = New-Object System.DirectoryServices.DirectorySearcher;
        $ds.SearchRoot = $dn.GetDirectoryEntry();
        $ds.SearchScope = "subtree";
        $ds.PageSize = 1024;
        $ds.Filter = "(&(objectCategory=user)(userworkstations=*$pcName*)(cn=F*))";
        $ds.PropertiesToLoad.AddRange($attributes.Split(","))
        $result = $ds.FindAll();
        $ds.Dispose();
        return $result;
    }
    (Find-ADObjects "cn,userworkstations").Properties
}

Add-Type -Assembly System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles();
$Form = New-Object Windows.Forms.Form 
$Form.Width = 900
$Form.Height = 130
$Form.Text = "SFFS av D8BP"

$progressBar = New-Object System.Windows.Forms.ProgressBar
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$runButton = New-Object System.Windows.Forms.Button
$inputTextbox = New-Object System.Windows.Forms.TextBox
$openFileButton = New-Object System.Windows.Forms.Button
$showMoreCheckbox = New-Object System.Windows.Forms.CheckBox
$outputTextbox = New-Object System.Windows.Forms.TextBox

$progressBar.Location = New-Object Drawing.Point(12, 49)
$progressBar.Name = "progressBar"
$progressBar.Size = New-Object Drawing.Size(767, 23)
$progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
$progressBar.TabIndex = 0
$progressBar.Value = 0
$progressBar.ForeColor = [System.Drawing.Color]::FromArgb(6, 176, 37)

$openFileDialog.Filter = "Excel Sheet|*.xlsx"
$openFileDialog.InitialDirectory = "$($currentUserDirectory)\Downloads"

$pathToIcon = Get-Item 'icon.png'
$openFileIcon = [System.Drawing.Image]::FromFile($pathToIcon)
$openFileButton.BackgroundImage = $openFileIcon
$openFileButton.BackgroundImageLayout = [System.Windows.Forms.ImageLayout]::Center
$openFileButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$openFileButton.FlatAppearance.BorderSize = 0
$openFileButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$openFileButton.Location = New-Object System.Drawing.Point(746, 12)
$openFileButton.Name = "openFileButton"
$openFileButton.Size = New-Object System.Drawing.Size(33, 33)
$openFileButton.TabIndex = 2
$openFileButton.UseVisualStyleBackColor = $false
$openFileButton_OnClick = {
    $result = $openFileDialog.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $inputTextbox.Text = $openFileDialog.FileName
    }
}
$openFileButton.Add_Click($openFileButton_OnClick)

$inputTextbox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$inputTextbox.Name = "inputTextbox"
$inputTextbox.Size = New-Object System.Drawing.Size(728, 31)
$inputTextbox.TabIndex = 1
$inputTextbox.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 15.75, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point, 0)
$inputTextbox.Location = New-Object System.Drawing.Point(12, 12)
$inputTextbox.ForeColor = [System.Drawing.Color]::FromArgb(239, 244, 255)
$inputTextbox.BackColor = [System.Drawing.Color]::FromArgb(51, 51, 51)

$showMoreCheckbox.Appearance = [System.Windows.Forms.Appearance]::Button
$showMoreCheckbox.AutoSize = $true
$showMoreCheckbox.BackColor = [System.Drawing.Color]::FromArgb(51, 51, 51)
$showMoreCheckbox.CheckAlign = [System.Drawing.ContentAlignment]::MiddleRight
$showMoreCheckbox.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$showMoreCheckbox.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8.25, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point, 0)
$showMoreCheckbox.Location = New-Object System.Drawing.Point(790, 48)
$showMoreCheckbox.Name = "showMoreCheckbox"
$showMoreCheckbox.Size = New-Object System.Drawing.Size(10, 23)
$showMoreCheckbox.TabIndex = 4
$showMoreCheckbox.Text = "▼ Visa Mer ▼"
$showMoreCheckbox.UseVisualStyleBackColor = $false
$showMoreCheckbox_OnCheckedChanged = {
    if ($showMoreCheckbox.Checked) {
        $Form.Height = 350;
        $outputTextbox.Visible = $true;
        $showMoreCheckbox.Text = "▲ Visa Mer ▲";
    }
    else {
        $Form.Height = 130;
        $outputTextbox.Visible = $false;
        $showMoreCheckbox.Text = "▼ Visa Mer ▼";
    }
}
$showMoreCheckbox.Add_CheckedChanged($showMoreCheckbox_OnCheckedChanged)

$runButton.BackColor = [System.Drawing.Color]::FromArgb(51, 51, 51)
$runButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$runButton.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 15.75, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point, 0)
$runButton.Location = New-Object System.Drawing.Point(785, 12)
$runButton.Name = "runButton"
$runButton.Size = New-Object System.Drawing.Size(94, 31)
$runButton.TabIndex = 3
$runButton.Text = "Starta"
$runButton.UseVisualStyleBackColor = $false
$runButton_onClick = {
    $runButton.Enabled = $false
    $output = (New-Object System.Collections.Concurrent.ConcurrentQueue[PSCustomObject])
    $pathToCsv = '' + $inputTextbox.Text + ''
    $pcList = Import-Excel -Path $pathToCsv -HeaderName 'OldPCName', 'OldUsername', 'OldSerial', 'NewPCName', 'NewUsername', 'NewSerial' -StartRow 3
    $pcListLength = 0
    foreach ($tempObj in $pcList) {
        $pcListLength = $pcListLength + 1
    }
    $progressBar.Maximum = $pcListLength
    $progressBar.Value = 0
    $pcJob = $pcList | ForEach-Object -AsJob -ThrottleLimit 10 -Parallel $sffs
    Start-ThreadJob -ScriptBlock {
        $tempPcJob = $using:pcJob
        $tempOutput = $using:output
        while ($tempPcJob.State -eq "Running" -or $tempOutput.Count -gt 0) {
            if ($tempOutput.Count -gt 0) {
                $tempObj = New-Object -TypeName PSObject
                if ($tempOutput.TryDequeue([ref]$tempObj)) {
                    $tempOutputTextbox = $using:outputTextbox
                    $tempOutputTextbox.Text = $tempObj.Text + $tempOutputTextbox.Text
                    $tempProgress = $using:progressBar
                    $tempProgress.Value = $tempProgress.Value + 1
                    Remove-Variable -Name "tempProgress"
                    Remove-Variable -Name "tempOutputTextbox"
                }
                Remove-Variable -Name "tempObj"
            }
        }
        Remove-Variable -Name "tempPcJob"
        Remove-Variable -Name "tempOutput"
    } -ThrottleLimit 10
    $runButton.Enabled = $true
}
$runButton.Add_Click($runButton_onClick)

$outputTextbox.BackColor = [System.Drawing.Color]::FromArgb(51, 51, 51)
$outputTextbox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$outputTextbox.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 11, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point, 0)
$outputTextbox.ForeColor = [System.Drawing.Color]::FromArgb(239, 244, 255)
$outputTextbox.Location = New-Object System.Drawing.Point(13, 79)
$outputTextbox.Multiline = $true
$outputTextbox.Name = "outputTextbox"
$outputTextbox.ReadOnly = $true
$outputTextbox.Size = New-Object System.Drawing.Size(866, 220)
$outputTextbox.TabIndex = 5
$outputTextbox.Visible = $false
$outputTextbox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
#$outputTextbox.SelectionStart = $outputTextbox.Length
#$outputTextbox.ScrollToCaret()

$Form.AllowDrop = $true
$Form.BackColor = [System.Drawing.Color]::FromArgb(29, 29, 29)
$Form.ForeColor = [System.Drawing.Color]::FromArgb(239, 244, 255)
$Form.Controls.Add($outputTextbox)
$Form.Controls.Add($showMoreCheckbox)
$Form.Controls.Add($openFileButton)
$Form.Controls.Add($inputTextbox)
$Form.Controls.Add($runButton)
$Form.Controls.Add($progressBar)
$Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$Form.MaximizeBox = $false
$Form.ShowIcon = $false
$Form_OnDragDrop = [System.Windows.Forms.DragEventHandler] {
    foreach ($filename in $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)) {
        if ($filename.Substring($filename.Length - 5, 5) -eq ".xlsx") {
            $inputTextbox.Text = $filename
        }
    }
}

$Form.Add_DragDrop($Form_OnDragDrop)
$Form_OnDragOver = {
    #Write-Host "Hovering"
    if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        $_.Effect = 'Copy'
    }
    else {
        $_.Effect = 'None'
    }
}
$Form.Add_DragOver($Form_OnDragOver)

$Form.ShowDialog()

$sffs = {
    #$pc = $_.
    $tempOut = ""
    $urlNew = "http://sysman.sll.se/SysMan/api/Client?name=" + $_.NewPCName + "&take=10&skip=0&type=0&targetActive=1"
    $ResponceNew = Invoke-WebRequest -Uri $urlNew -AllowUnencryptedAuthentication -UseDefaultCredentials -SessionVariable 'Session'
    $idNew = (($ResponceNew.Content | ConvertFrom-Json).result).id
    $urlOld = "http://sysman.sll.se/SysMan/api/Client?name=" + $_.OldPCName + "&take=10&skip=0&type=0&targetActive=1"
    $ResponceOld = Invoke-WebRequest -Uri $urlOld -AllowUnencryptedAuthentication -UseDefaultCredentials -WebSession $Session
    $idOld = (($ResponceOld.Content | ConvertFrom-Json).result).id
    if (!($ResponceNew.Content | ConvertFrom-Json).result) {
        $tempOut = $tempOut + "Nya Datorn: $($_.NewPCName) hittas ej.`r`n"
    }
    elseif (!($ResponceOld.Content | ConvertFrom-Json).result) {
        $tempOut = $tempOut + "Gamla Datorn: $($_.OldPCName) hittas ej.`r`n"
    }
    else {
        $urlOld = "http://sysman.sll.se/SysMan/api/Reporting/Client?clientId=" + $idOld
        $ResponceOld = Invoke-WebRequest -Uri $urlOld -AllowUnencryptedAuthentication -UseDefaultCredentials -WebSession $Session
        $printers = New-Object -TypeName "System.Collections.ArrayList"
        foreach ($printerTemp in ($ResponceOld.Content | ConvertFrom-Json).sysManInformation.installedPrinters) {
            $printer = [PSCustomObject]@{id = $printerTemp.id ; isDefault = $printerTemp.isDefault }
            $printers += $printer
        }
        $idNewList = New-Object -TypeName "System.Collections.ArrayList"
        $idNewList += [int]$idNew
        $requestBody = 
        @{
            targets          = $idNewList
            printers         = $printers
            templateTargetId = $null
        } | ConvertTo-Json -Compress
        Invoke-WebRequest -Method Post -Uri "http://sysman.sll.se/SysMan/api/v2/printer/install" -Body $requestBody -AllowUnencryptedAuthentication -UseDefaultCredentials -WebSession $Session -ContentType "application/json"
        #$NewName = $using:pc.NewPCName
        $tempOut = $tempOut + "Datorn: {0} har fått {1} skrivare från {2}.`r`n" -f $_.NewPCName, $printers.Count, $_.OldPCName

        $fkonto = fkarfinder -pcName $_.OldPCName
        if ($null -ne $fkonto.cn) { 
            $fkonto = ($fkonto.cn -join ', ')
            $tempOut = $tempOut + "Hittat funktionskonto $($fkonto)`r`n"
            if (!((Get-ADUser -Identity $fkonto -Properties *).userWorkstations -like "*$($_.NewPCName)*")) {
                $logOnTo = (Get-ADUser -Identity $fkonto -Properties *).userWorkstations + "," + $_.NewPCName 
                $fkonto | Set-ADUser -LogonWorkstations $logOnTo
            }
            $pcIdentity = Get-ADComputer -Identity $_.NewPCName
            $tempOut = $tempOut + "$($fkonto) fungerar nu på : $($logOnTo)`r`n"
            Add-ADGroupMember -Identity "$($_.NewPCName.Substring(0,3))_Wrk_F-kontoWS_SLLeKlient" -Members $pcIdentity.ObjectGUID
            $tempOut = $tempOut + "Dator tillagd i $($_.NewPCName.Substring(0,3))_Wrk_F-kontoWS_SLLeKlient`r`n"
        }
        else {
            $tempOut = $tempOut + "Dator har inte Funktionskonto`r`n"
        }
    }
    $tempQueue = $using:output
    $tempQueue.Enqueue(<#$pcObject#>[PSCustomObject]@{
            Text = $tempOut
        })
    Remove-Variable -Name $pc
    Remove-Variable -Name $tempOut
    Remove-Variable -Name $tempQueue
    Remove-Variable -Name $printers
    Remove-Variable -Name $printerTemp
    Remove-Variable -Name $pcIdentity
    Remove-Variable -Name $fkonto
    Remove-Variable -Name $idNewList
    Remove-Variable -Name $idNew
    Remove-Variable -Name $idOld
    Remove-Variable -Name $requestBody
    Remove-Variable -Name $ResponceOld
    Remove-Variable -Name $ResponceNew
    Remove-Variable -Name $urlOld
    Remove-Variable -Name $urlNew
    Remove-Variable -Name $printer
    Remove-Variable -Name $Session
    Remove-Variable -Name $logOnTo
}