if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator"))
{
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    break
}
while ($true) {
    $continue = $false
    $pc = Read-Host "Enter MAC or PC name"
    $pc = ($pc.ToString()).ToLower()
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
    elseif (!$pc.Contains(":") -and ($pc.length -eq 12)) {
        $pc = $pc.insert(2,":").insert(5,":").insert(8,":").insert(11,":").insert(14,":")
        $pc = $pc -replace ":","%3A"
        $continue = $true
    }
    elseif ($pc.Contains(":") -and ($pc.length -eq 17)) {
        $pc = $pc -replace ":","%3A"
        $continue = $true
    }
    else { Write-Host "Wrong format of name, please try again" }

    if ($continue) {
        $url = "http://sysman.sll.se/SysMan/api/Client?name=" + $pc + "&take=10&skip=0&type=0&targetActive=1"
        $Responce = Invoke-WebRequest -Uri $url -UseDefaultCredentials -AllowUnencryptedAuthentication -SessionVariable 'Session'
        if (!($Responce.Content | ConvertFrom-Json).result) {
            Write-Host "Datorn:" $pc "kan inte hittas i Sysman, den är med högst sannolikhet inaktiv."
        }
        else {
            Clear-Host
            $pcName = (($Responce.Content | ConvertFrom-Json).result).Name
            #MAC
            $mac = "EJ ANGIVEN"
            $requestBody =
            @{
                UserName = ""
                ComputerName = $pcName.ToString()
                Id = "88eeae01-fc85-426c-898d-dae73ec31867"
            } | ConvertTo-Json -Compress
            $macResponce = Invoke-WebRequest -Method Post -Uri "http://sysman.sll.se/SysMan/api/Tool/Run" -AllowUnencryptedAuthentication -Body $requestBody -WebSession $Session -ContentType "application/json"
            $macResponce = ($macResponce | ConvertFrom-Json).result
            foreach ($result in $macResponce) {
                if ((($result -like "*Ethernet*") -and !($result -like "*Dock*")) -or (($result -like "*GbE*") -and !($result -like "*USB*")) -or $result -like "*Gigabit*") {
                    $macColon = $result.Substring(0,17)
                    $mac = $macColon -replace ":",""
                }
            }
            #random bs
            $model = ((($Responce.Content | ConvertFrom-Json).result).hardwareModel).Name
            $id = (($Responce.Content | ConvertFrom-Json).result).Id
            $request = "http://sysman.sll.se/SysMan/api/Reporting/Client?clientId=" + $id
            $Responce = (Invoke-WebRequest -Uri $request -AllowUnencryptedAuthentication -WebSession $Session).Content | ConvertFrom-Json
            $latestReboot = $Responce.lastBootTime
            $freeSpace = $Responce.health.osDiskFreeSpace
            if (!$freeSpace) {
                $freeSpace = "Okänt"
            }
            $serial = $Responce.serial
            if ($Responce.operatingSystem -like "*7*") {
                #$os = "1"
                $os = "Win7"
            }
            elseif ($Responce.operatingSystem -like "*10*") {
                #$os = "2"
                $os = "Win10"
            }
            else {
                #$os = "0"
                $os = "Okänt"
            }
            $lastUser = $Responce.lastUser
            $bit = $Responce.processorArchitecture
            $names = $Responce.collections
            $filteredName = "Inte hittad"
            foreach ($name in $names) {

                if (($name.StartsWith("Kar_Wrk_PR") -or $name.StartsWith("Sos_Wrk_PR") -or $name.StartsWith("Dan_Wrk_PR") -or ($name.StartsWith("Lit_Wrk_PR") -or $name.StartsWith("Sll_Wrk_PR") -or $name.StartsWith("Hsf_Wrk_PR") -or $name.StartsWith("Lsf_Wrk_PR") -or $name.StartsWith("Int_Wrk_PR") -or $name.StartsWith("Fut_Wrk_PR") -or $name.StartsWith("Pnf_Wrk_PR") -or $name.StartsWith("Rev_Wrk_PR") -or $name.StartsWith("Ser_Wrk_PR") -or $name.StartsWith("Tka_Wrk_PR") -or $name.StartsWith("Trf_Wrk_PR") -or $name.StartsWith("Tsl_Wrk_PR") -or $name.StartsWith("Pnf_Wrk_PR") -or $name.StartsWith("Ita_Wrk_PR")))) {
                    $filteredName = $name
                    $filteredName = $filteredName.Substring(11)

                }
            }
            $names = $Responce.installedApplications
            if ($filteredName -eq "Inte hittad") {
                foreach ($name in $names) {
                    #$name.name
                    if (($name.Name.StartsWith("Sll_Wrk_Kar_PR") -or $name.Name.StartsWith("Sll_Wrk_Sos_PR") -or $name.Name.StartsWith("Sll_Wrk_Dan_PR") -or ($name.Name.StartsWith("Sll_Wrk_Lit_PR") -or $name.Name.StartsWith("Sll_Wrk_Sll_PR") -or $name.Name.StartsWith("Sll_Wrk_Hsf_PR") -or $name.Name.StartsWith("Sll_Wrk_Lsf_PR") -or $name.Name.StartsWith("Sll_Wrk_Int_PR") -or $name.Name.StartsWith("Sll_Wrk_Fut_PR") -or $name.Name.StartsWith("Sll_Wrk_Pnf_PR") -or $name.Name.StartsWith("Sll_Wrk_Rev_PR") -or $name.Name.StartsWith("Sll_Wrk_Ser_PR") -or $name.Name.StartsWith("Sll_Wrk_Tka_PR") -or $name.Name.StartsWith("Sll_Wrk_Trf_PR") -or $name.Name.StartsWith("Sll_Wrk_Tsl_PR") -or $name.Name.StartsWith("Sll_Wrk_Pnf_PR") -or $name.Name.StartsWith("Sll_Wrk_Ita_PR")))) {
                        $filteredName = $name.Name
                        $filteredName = $filteredName.Substring(15)
                    }
                }

            }
            if (($filteredName -eq "Administrativ_PC") -and ($bit -like "*64*")) {
                $filteredName = "Administrativ_PC_64bit"
            }
            $adVarde = & $PSScriptRoot\fkarfinder.ps1 $pcName
            if ($null -ne $adVarde.cn) { $adVarde = ($adVarde.cn -join ', ') }
            else {
                $adVarde = "NEJ"
            }
            $lokaladmin = & $PSScriptRoot\localadmin.ps1 $pcName
            if ($localadmin) {
                $lokaladmin = "JA"
            }
            else {
                $lokaladmin = "NEJ"
            }
            $request = "http://sysman.sll.se/SysMan/api/monitoring?target=" + $pcName
            $Responce = (Invoke-WebRequest -Uri $request -AllowUnencryptedAuthentication -WebSession $Session).Content|ConvertFrom-Json
            $latestReinstall = $Responce.endDate[0]
            #BIOS
            $requestBody = 
	    	@{
                ComputerName=$pcName.ToString()
                UserName=""
                
                Id="a5733056-bd09-4aae-b7c9-80d18758b218"
            } | ConvertTo-Json -Compress
            $biosResponce = Invoke-WebRequest -Method Post -Uri "http://sysman.sll.se/SysMan/api/Tool/Run" -AllowUnencryptedAuthentication -Body $requestBody  -WebSession $Session -ContentType "application/json"
            $biosResponce = ($biosResponce| ConvertFrom-Json).result
            foreach ($result in $biosResponce){
                if ($result.StartsWith("Version")){
                    $bios = $result.Substring(9)
                    
                }
            }
            #IP
            try {
                $hostEntry= [System.Net.Dns]::GetHostByName($pcName) 
                foreach ($item in $hostEntry.AddressList) {
                    if ($item.AddressFamily -like "InterNetwork") {
                        $ip = $item.IPAddressToString
                        $ping = Test-Connection -IPv4 -Count 1 -Quiet -TargetName $ip
                        if ($ping) {
                            $ping = "Datorn svarar på ping"
                        }
                        else {
                            $ping = "Datorn svarar inte på ping"
                        }
                    }
                }
            }
            catch {
                $ip = "Ej hittad"
            }
            $pcObject = New-Object -TypeName PSObject
            Add-Member -InputObject $pcObject -MemberType NoteProperty -Name Datornamn -Value $pcName
            Add-Member -InputObject $pcObject -MemberType NoteProperty -Name Operativsystem -Value $os
            Add-Member -InputObject $pcObject -MemberType NoteProperty -Name Modell -Value $model
            Add-Member -InputObject $pcObject -MemberType NoteProperty -Name Senaste_ominstallation -Value $latestReinstall
            Add-Member -InputObject $pcObject -MemberType NoteProperty -Name Senaste_omstart -Value $latestReboot
            Add-Member -InputObject $pcObject -MemberType NoteProperty -Name Roll -Value $filteredName
            Add-Member -InputObject $pcObject -MemberType NoteProperty -Name Biosversion -Value $bios
            Add-Member -InputObject $pcObject -MemberType NoteProperty -Name Ip-Adress -Value $ip
            Add-Member -InputObject $pcObject -MemberType NoteProperty -Name Ping_test -Value $ping
            Add-Member -InputObject $pcObject -MemberType NoteProperty -Name MAC_med_kolon -Value $macColon
            Add-Member -InputObject $pcObject -MemberType NoteProperty -Name MAC_utan_kolon -Value $mac
            Add-Member -InputObject $pcObject -MemberType NoteProperty -Name Serienummer -Value $serial
            Add-Member -InputObject $pcObject -MemberType NoteProperty -Name Senaste_användare -Value $lastUser
            Add-Member -InputObject $pcObject -MemberType NoteProperty -Name Ledig_plats -Value $freeSpace
            Add-Member -InputObject $pcObject -MemberType NoteProperty -Name Funktionskonto -Value $adVarde
            Add-Member -InputObject $pcObject -MemberType NoteProperty -Name LokalAdmin -Value $lokaladmin
            
            $pcObject
            Pause
            Clear-Host
        }
    }
}
