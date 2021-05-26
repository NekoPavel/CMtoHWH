If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {  
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    Break
}
function localadmininfo {
    param($pcName)

    function Find-ADObjects($domain, $class, $filter, $attributes = "distinguishedName") {
        $dc = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext ([System.DirectoryServices.ActiveDirectory.DirectoryContextType]"domain", $domain)
        $dn = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($dc)

        $ds = New-Object System.DirectoryServices.DirectorySearcher
        $ds.SearchRoot = $dn.GetDirectoryEntry()
        $ds.SearchScope = "subtree"
        $ds.PageSize = 1024
        $ds.Filter = "(&(objectCategory=$class)$filter)"
        $ds.PropertiesToLoad.AddRange($attributes.Split(","))
        $result = $ds.FindAll()
        $ds.Dispose()
        return $result
    }

    $adminRoles = @("CN=Kar_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Kar,OU=HealthCare,DC=gaia,DC=sll,DC=se", "CN=Sos_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Sos,OU=HealthCare,DC=gaia,DC=sll,DC=se", "CN=Lit_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Lit,OU=Administration,DC=gaia,DC=sll,DC=se", "CN=Ita_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Ita,OU=Reference,DC=gaia,DC=sll,DC=se", "CN=Dan_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Dan,OU=HealthCare,DC=gaia,DC=sll,DC=se", "CN=Hsf_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Hsf,OU=Administration,DC=gaia,DC=sll,DC=se", "CN=Lsf_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Lsf,OU=Administration,DC=gaia,DC=sll,DC=se", "CN=Fut_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Fut,OU=PublicTransportation,DC=gaia,DC=sll,DC=se", "CN=Int_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Int,OU=Administration,DC=gaia,DC=sll,DC=se", "CN=Trf_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Trf,OU=Administration,DC=gaia,DC=sll,DC=se", "CN=Sll_Wrk_LocalAdmin_SLLeKlient,OU=Workstation,OU=Groups,OU=Sll,DC=gaia,DC=sll,DC=se", "CN=Ser_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Ser,OU=Administration,DC=gaia,DC=sll,DC=se")
    $adminRolesRegex = [string]::Join('|', $adminRoles)
    $adVarde = (Find-ADObjects "gaia" "computer" "(cn=$pcName)" "cn,MemberOf").Properties
    if ($adVarde.memberof -match $adminRolesRegex) {
        $true
    }
    else {
        $false
    }
}
function fkarfinder {
    param($pcName)

    function Find-ADObjects($attributes = "distinguishedName") {
        $dc = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext ([System.DirectoryServices.ActiveDirectory.DirectoryContextType]"domain", "gaia");
        $dn = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($dc);

        $ds = New-Object System.DirectoryServices.DirectorySearcher;
        $ds.SearchRoot = $dn.GetDirectoryEntry();
        $ds.SearchScope = "subtree";
        $ds.PageSize = 2048;
        $ds.Filter = "(&(objectCategory=user)(userworkstations=*$pcName*)(cn=F*))";
        $ds.PropertiesToLoad.AddRange($attributes.Split(","))
        $result = $ds.FindAll();
        $ds.Dispose();
        return $result;
    }

    (Find-ADObjects "cn,userworkstations").Properties
}
Add-Type -Assembly System.Windows.Forms
$Form = New-Object Windows.Forms.Form 
$Form.Width = 505
$Form.Height = 325
$Form.Text = "HittaDator V2.0"

$infoTextLbl = New-Object Windows.Forms.Label
$datorNamnIn = New-Object System.Windows.Forms.TextBox
$hittaDatorBtn = New-Object System.Windows.Forms.Button
$datorInfoUt = New-Object System.Windows.Forms.TextBox

$infoTextLbl.Location = New-Object Drawing.Point(5, 5)
$infoTextLbl.Text = "Datornamn/Mac:adress"
$infoTextLbl.AutoSize = $true
$infoTextLbl.Visible = $true
$infoTextLbl.ForeColor = [System.Drawing.Color]::FromArgb(239,244,255)
$infoTextLbl.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 15.75, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point, 0)
$infoTextLbl.Size = New-Object System.Drawing.Size(235, 31)
$infoTextLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft

$datorNamnIn.Location = New-Object System.Drawing.Point(240, 5)
$datorNamnIn.Name = "datorNamnIn"
$datorNamnIn.Size = New-Object System.Drawing.Size(185, 31)
$datorNamnIn.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$datorNamnIn.BackColor = [System.Drawing.Color]::FromArgb(51,51,51)
$datorNamnIn.ForeColor = [System.Drawing.Color]::FromArgb(239,244,255)
$datorNamnIn.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 15.75, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point, 0)

$hittaDatorBtn.Location = New-Object System.Drawing.Point(430, 5)
$hittaDatorBtn.Name = "hittaDatorBtn"
$hittaDatorBtn.Size = New-Object System.Drawing.Size(55, 31)
$hittaDatorBtn.TabIndex = 2
$hittaDatorBtn.Text = "Sök"
$hittaDatorBtn.UseVisualStyleBackColor = $true
$hittaDatorBtn_OnClick = {
    $hittaDatorBtn.Enabled = $false
    $Form.AcceptButton = $null
    $continue = $false
    $pc = $datorNamnIn.Text
    $pc = (($pc.ToString()).Trim()).ToLower()
    if ($pc.StartsWith("kar") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("lit") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("sos") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("hsf") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("lsf") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("int") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("fut") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("pnf") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("rev") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("ser") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("sll") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("tka") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("trf") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("tsl") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("pnf") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("ita") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("dan") -and ($pc.length -eq 13)) { $continue = $true }
    elseif (!$pc.Contains(":") -and ($pc.length -eq 12 )) {

        $pc = $pc.insert(2, ":").insert(5, ":").insert(8, ":").insert(11, ":").insert(14, ":")
        $pc = $pc -replace ":", "%3A"
        $continue = $true
    }
    elseif ($pc.Contains(":") -and ($pc.length -eq 17 )) {
        $pc = $pc -replace ":", "%3A"
        $continue = $true
    }
    else { 
        $datorInfoUt.Text = "Fel format på namn, försök igen!" 
        $datorInfoUt.ForeColor = [System.Drawing.Color]::FromArgb(238,17,17)
        $hittaDatorBtn.Enabled = $true
        $Form.AcceptButton = $hittaDatorBtn
    }
    
            
    if ($continue) {    
        $url = "http://sysman.sll.se/SysMan/api/Client?name=" + $pc + "&take=10&skip=0&type=0&targetActive=1"
        $Responce = Invoke-WebRequest -Uri $url -AllowUnencryptedAuthentication -UseDefaultCredentials -SessionVariable 'Session'
        if (!($Responce.Content | ConvertFrom-Json).result) {
            $datorInfoUt.Text = "Datorn: $($pc) kan inte hittas i Sysman, den är med högst sannolikhet inaktiv."
            $datorInfoUt.ForeColor = [System.Drawing.Color]::FromArgb(238,17,17)
            $hittaDatorBtn.Enabled = $true
            $Form.AcceptButton = $hittaDatorBtn
        }
        else {
            $pcName = (($Responce.Content | ConvertFrom-Json).result).name
            $mac = "EJ ANGIVEN"
            $requestBody =
            @{
                UserName     = ""
                ComputerName = $pcName.ToString()
                Id           = "88eeae01-fc85-426c-898d-dae73ec31867"
            } | ConvertTo-Json -Compress
            $macResponce = Invoke-WebRequest -Method Post -Uri "http://sysman.sll.se/SysMan/api/Tool/Run" -Body $requestBody -AllowUnencryptedAuthentication -WebSession $Session -ContentType "application/json"
            $macResponce = ($macResponce | ConvertFrom-Json).result
            foreach ($result in $macResponce) {
                if ((($result -like "*Ethernet*") -and !($result -like "*#2*")  -and !($result -like "*Virtual*") -and !($result -like "*Server Adapter*") -and !($result -like "*Dock*")) -or (($result -like "*GbE*") -and !($result -like "*#2*") -and !($result -like "*USB*")) -or $result -like "*Gigabit*") {
                    [string]$macColon = [string]$result.Substring(0, 17)
                    [string]$mac = [string]([string]$macColon -replace ":", "")
                }
            }
            $model = ((($Responce.Content | ConvertFrom-Json).result).hardwareModel).name
            $id = (($Responce.Content | ConvertFrom-Json).result).id
            $request = "http://sysman.sll.se/SysMan/api/Reporting/Client?clientId=" + $id
            $Responce = (Invoke-WebRequest -Uri $request -AllowUnencryptedAuthentication -WebSession $Session).Content |  ConvertFrom-Json
            $serial = $Responce.serial
            if ($Responce.operatingSystem -like "*7*") {
                $os = "Win7"
            }
            elseif ($Responce.operatingSystem -like "*10*") {
                $os = "Win10"
            }
            else {
                $os = "Okänt"
            }
            $bit = $Responce.processorArchitecture
            $names = $Responce.collections
            $filteredName = "Inte hittad"
            foreach ($name in $names) {
                    
                if (($name.StartsWith("Kar_Wrk_PR") -or $name.StartsWith("Sos_Wrk_PR") -or $name.StartsWith("Dan_Wrk_PR") -or ($name.StartsWith("Lit_Wrk_PR") -or $name.StartsWith("Sll_Wrk_PR") -or $name.StartsWith("Hsf_Wrk_PR") -or $name.StartsWith("Lsf_Wrk_PR") -or $name.StartsWith("Int_Wrk_PR") -or $name.StartsWith("Fut_Wrk_PR") -or $name.StartsWith("Pnf_Wrk_PR") -or $name.StartsWith("Rev_Wrk_PR") -or $name.StartsWith("Ser_Wrk_PR") -or $name.StartsWith("Tka_Wrk_PR") -or $name.StartsWith("Trf_Wrk_PR") -or $name.StartsWith("Tsl_Wrk_PR") -or $name.StartsWith("Pnf_Wrk_PR") -or $name.StartsWith("Ita_Wrk_PR")))) {
                    $filteredName = $name
                    $filteredName = $filteredName.Substring(11)
                        
                }
            }
            if ($filteredName -eq "Kiosk_PC"){
                foreach($name in $names){
                    if ($name.StartsWith("Gai_App_CitrixReceiverVardTerminal")){
                        $filteredName = "Vardterminal"
                    }
                }
            }
            $names = $Responce.installedApplications
            if ($filteredName -eq "Inte hittad") {
                foreach ($name in $names) {
                    if (($name.name.StartsWith("Sll_Wrk_Kar_PR") -or $name.name.StartsWith("Sll_Wrk_Sos_PR") -or $name.name.StartsWith("Sll_Wrk_Dan_PR") -or ($name.name.StartsWith("Sll_Wrk_Lit_PR") -or $name.name.StartsWith("Sll_Wrk_Sll_PR") -or $name.name.StartsWith("Sll_Wrk_Hsf_PR") -or $name.name.StartsWith("Sll_Wrk_Lsf_PR") -or $name.name.StartsWith("Sll_Wrk_Int_PR") -or $name.name.StartsWith("Sll_Wrk_Fut_PR") -or $name.name.StartsWith("Sll_Wrk_Pnf_PR") -or $name.name.StartsWith("Sll_Wrk_Rev_PR") -or $name.name.StartsWith("Sll_Wrk_Ser_PR") -or $name.name.StartsWith("Sll_Wrk_Tka_PR") -or $name.name.StartsWith("Sll_Wrk_Trf_PR") -or $name.name.StartsWith("Sll_Wrk_Tsl_PR") -or $name.name.StartsWith("Sll_Wrk_Pnf_PR") -or $name.name.StartsWith("Sll_Wrk_Ita_PR")))) {
                        $filteredName = $name.name
                        $filteredName = $filteredName.Substring(15)
                    }
                }
                
            }
            if (($filteredName -eq "Kiosk_PC" )-or($filteredName -eq "Inte hittad")){
                foreach($name in $names){
                    if ($names.Name -eq "Kar_Rol_Vardterminal-Kiosk"){
                        $filteredName = "Vardterminal"
                    }
                }
            }
            if (($filteredName -eq "Administrativ_PC" ) -and ($bit -like "*64*")) {
                $filteredName = "Administrativ_PC_64bit"
            }
            $adVarde = fkarfinder -pcName $pcName
            if ($null -ne $adVarde.cn) { $adVarde = ($adVarde.cn -join ', ') } 
            else {
                $adVarde = "Nej"
            }


            $lokaladmin = localadmininfo -pcName $pcName
            if ($localadmin) {
                $lokaladmin = "Ja"
            }
            else {
                $lokaladmin = "Nej"
            }
            $datorInfoUt.Text = "Operativsystem`t= $($os)`r`nModell`t`t= $($model)`r`nRoll`t`t= $($filteredName)`r`nDatornamn`t= $($pcName)`r`nMAC-adress`t= $($mac)`r`nMAC med kolon`t= $($macColon)`r`nSerienummer`t= $($serial)`r`nFunktionskonto`t= $($adVarde)`r`nLokalAdmin`t= $($lokaladmin)"
            $datorInfoUt.ForeColor = [System.Drawing.Color]::FromArgb(239,244,255)
            $datorNamnIn.Clear()
            $hittaDatorBtn.Enabled = $true
            $Form.AcceptButton = $hittaDatorBtn            
        }
    }
}
$hittaDatorBtn.Add_Click($hittaDatorBtn_OnClick)
$hittaDatorBtn.BackColor = [System.Drawing.Color]::FromArgb(51,51,51)
$hittaDatorBtn.ForeColor = [System.Drawing.Color]::FromArgb(239,244,255)
$hittaDatorBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$hittaDatorBtn.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 15.75, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point, 0)

$datorInfoUt.Location = New-Object System.Drawing.Point(5, 41)
$datorInfoUt.Multiline = $true
$datorInfoUt.Name = "datorInfoUt"
$datorInfoUt.ReadOnly = $true
$datorInfoUt.Size = New-Object System.Drawing.Size(480, 240)
$datorInfoUt.BackColor = [System.Drawing.Color]::FromArgb(51,51,51)
$datorInfoUt.ForeColor = [System.Drawing.Color]::FromArgb(239,244,255)
$datorInfoUt.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$datorInfoUt.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 15.75, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point, 0)

$Form.BackColor = [System.Drawing.Color]::FromArgb(29,29,29)
$Form.ForeColor = [System.Drawing.Color]::FromArgb(239,244,255)
$Form.AcceptButton = $hittaDatorBtn
$Form.Controls.Add($datorNamnIn)
$Form.Controls.Add($datorInfoUt)
$Form.Controls.Add($hittaDatorBtn)
$Form.Controls.Add($infoTextLbl)
$Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$Form.MaximizeBox = $false

$Form.ShowDialog()

