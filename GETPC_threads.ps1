param([parameter(Mandatory=$true)][string]$pathToCsv,[parameter(Mandatory=$true)][string]$filename)
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")){
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    break
}
#$excludedModels = @("Virtual Machine","VMware Virtual Platform","Parallels Virtual Platform","OEM")
#$excludedModelsRegex = [string]::Join('|',$excludedModels)
$PCObjects = (New-Object System.Collections.Concurrent.ConcurrentQueue[PSCustomObject])
$findPC = {
    $pc = $_.pcName.ToLower()
    $continue = $false
    if ($pc.StartsWith("kar") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("lit") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("sos") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("hsf") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("lsf") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("int") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("fut") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("rev") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("ser") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("sll") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("tka") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("trf") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("tsl") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("pnf") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("ita") -and ($pc.length -eq 13)) { $continue = $true }
    elseif ($pc.StartsWith("dan") -and ($pc.length -eq 13)) { $continue = $true }
    else {}
    if ($continue) {
        $url = "http://sysman.sll.se/SysMan/api/Client?name=" + $pc + "&take=10&skip=0&type=0&targetActive=1"
        $Responce = Invoke-WebRequest -Uri $url -UseDefaultCredentials -AllowUnencryptedAuthentication -SessionVariable 'Session'
        if (!($Responce.Content | ConvertFrom-Json).result) {}
        elseif (($Responce.Content | ConvertFrom-Json).result) {
            $save = $true
            $pcName = (($Responce.Content | ConvertFrom-Json).result).Name
            $requestBody =
            @{
                UserName = ""
                ComputerName = $pcName.ToString()
                Id = "88eeae01-fc85-426c-898d-dae73ec31867"
            } | ConvertTo-Json -Compress
            $macResponce = Invoke-WebRequest -Method Post -Uri "http://sysman.sll.se/SysMan/api/Tool/Run" -Body $requestBody -WebSession $Session -ContentType "application/json" -AllowUnencryptedAuthentication
            $macResponce = ($macResponce | ConvertFrom-Json).result
            #MAC
            foreach ($result in $macResponce) {
                if ((($result -like "*Ethernet*") -and !($result -like "*Virtual*") -and !($result -like "*Server Adapter*") -and !($result -like "*Dock*")) -or (($result -like "*GbE*")  -and !($result -like "*USB*")) -or $result -like "*Gigabit*") {
                    [string]$macColon = [string]$result.Substring(0,17)
                    [string]$mac = [string]([string]$macColon -replace ":","")
                }
            }
            if (!$mac -or !$macColon) {
                $save = $false
            }
            #Modell
            $model = ((($Responce.Content | ConvertFrom-Json).result).hardwareModel).Name
            
            $id = (($Responce.Content | ConvertFrom-Json).result).Id
            $request = "http://sysman.sll.se/SysMan/api/Reporting/Client?clientId=" + $id
            $Responce = (Invoke-WebRequest -Uri $request -AllowUnencryptedAuthentication -WebSession $Session).Content | ConvertFrom-Json
            #Serienummer
            [string]$serial = [string]$Responce.serial
            if ($serial -like "*O.E.M.*") {
                $save = $false
            }

            #Mer modellsaker
            foreach ($pc_model in $using:pc_modelsList) {
                if ($pc_model.hv_typ -eq $model) {
                    $model = $pc_model.Id
                    $hardware = $pc_model.hv_category
                    if ($hardware -eq "4") {
                        foreach ($aio in $using:aioList) {
                            if ($aio -eq $pcName) {
                                        $hardware = "6"
                                    }
                                }
                        }
                    if ($hardware -eq "6" -and $model -eq "91") {
                        $model = "90"
                    }
                    if ($model -eq "69") {
                        #Om datorn INTE finns i listan så är det en av de nya
                        foreach ($aioSerial in $using:aioSerialList69) {
                            if (!($aioSerial -eq $serial)) {
                                $model = "106"
                            }
                        }
                        # M725s
                    }
                    if ($model -eq "107") {
                        #Om datorn INTE finns i listan så är det en av de nya
                        foreach ($aioSerial in $using:aioSerialList107) {
                            if (!($aioSerial -eq $serial)) {
                                $model = "117"
                            }
                        }
                        #M75s1
                    }

                }
            }
            <#
            if ($model -match $excludedModelsRegex) {
                #$save = $false
                #$model = "match"
            }
            #>
            if ($model -match "\D" -or !$model) {
                $save = $false
                #TODO Make this log
                $model | Out-File -FilePath $PSScriptRoot\unmappedModelsLog.txt -Append
                #This should output to a text file
            }

            #OS
            if ($Responce.operatingSystem -like "*7*") {
                $os = "W7"   
            }
            elseif ($Responce.operatingSystem -like "*10*") {
                $os = "W10"  
            }
            else {
                $save = $false
            }
            #Här händer magin för roller
            $bit = $Responce.processorArchitecture
            $names = $Responce.collections
            $filteredName = "Inte hittad"
            foreach ($name in $names) {   
                if (($name.StartsWith("Kar_Wrk_PR") -or $name.StartsWith("Sos_Wrk_PR") -or ($name.StartsWith("Lit_Wrk_PR") -or $name.StartsWith("Dan_Wrk_PR") -or $name.StartsWith("Sll_Wrk_PR") -or $name.StartsWith("Hsf_Wrk_PR") -or $name.StartsWith("Lsf_Wrk_PR") -or $name.StartsWith("Int_Wrk_PR") -or $name.StartsWith("Fut_Wrk_PR") -or $name.StartsWith("Pnf_Wrk_PR") -or $name.StartsWith("Rev_Wrk_PR") -or $name.StartsWith("Ser_Wrk_PR") -or $name.StartsWith("Tka_Wrk_PR") -or $name.StartsWith("Trf_Wrk_PR") -or $name.StartsWith("Tsl_Wrk_PR") -or $name.StartsWith("Pnf_Wrk_PR") -or $name.StartsWith("Ita_Wrk_PR")))) {
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
                    if (($name.Name.StartsWith("Sll_Wrk_Kar_PR") -or $name.Name.StartsWith("Sll_Wrk_Sos_PR") -or $name.Name.StartsWith("Sll_Wrk_Dan_PR") -or ($name.Name.StartsWith("Sll_Wrk_Lit_PR") -or $name.Name.StartsWith("Sll_Wrk_Sll_PR") -or $name.Name.StartsWith("Sll_Wrk_Hsf_PR") -or $name.Name.StartsWith("Sll_Wrk_Lsf_PR") -or $name.Name.StartsWith("Sll_Wrk_Int_PR") -or $name.Name.StartsWith("Sll_Wrk_Fut_PR") -or $name.Name.StartsWith("Sll_Wrk_Pnf_PR") -or $name.Name.StartsWith("Sll_Wrk_Rev_PR") -or $name.Name.StartsWith("Sll_Wrk_Ser_PR") -or $name.Name.StartsWith("Sll_Wrk_Tka_PR") -or $name.Name.StartsWith("Sll_Wrk_Trf_PR") -or $name.Name.StartsWith("Sll_Wrk_Tsl_PR") -or $name.Name.StartsWith("Sll_Wrk_Pnf_PR") -or $name.Name.StartsWith("Sll_Wrk_Ita_PR")))) {
                        $filteredName = $name.Name
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
            foreach ($pc_role in $using:rolesList) {
                if ($pc_role.role -eq $filteredName) {
                    $filteredName = $pc_role.Id
                }
            }
            if (($filteredName -eq "1") -and ($bit -like "*64*")) {
                $filteredName = "2"
            }
            if ($filteredName -eq "Inte hittad") {
                $save = $false
            }
            #Funktionskonto
            $adVarde = &$PSScriptRoot\fkarfinder.ps1 $pcName
            if ($null -ne $adVarde.cn) { $adVarde = ($adVarde.cn -join ', ') }
            else {
                $adVarde = "N/A"
            }
            $lokaladmin = &$PSScriptRoot\localadmin.ps1 $pcName
            if ($lokaladmin) {
                $lokaladmin = "JA"
            }
            else {
                $lokaladmin = "NEJ"
            }
            if ($save) {
                <#
                $pcObject = New-Object -TypeName PSObject 
                Add-Member -InputObject $pcObject -MemberType NoteProperty -Name Hardvara_G -Value $hardware
                Add-Member -InputObject $pcObject -MemberType NoteProperty -Name Operativsystem_G -Value $os
                Add-Member -InputObject $pcObject -MemberType NoteProperty -Name Modell_G -Value $model
                Add-Member -InputObject $pcObject -MemberType NoteProperty -Name Roll_G -Value $filteredName
                Add-Member -InputObject $pcObject -MemberType NoteProperty -Name Hardvarunamn_G -Value $pcName
                Add-Member -InputObject $pcObject -MemberType NoteProperty -Name MAC_Adress_G -Value $mac
                Add-Member -InputObject $pcObject -MemberType NoteProperty -Name MAC_med_Kolon_G -Value $macColon
                Add-Member -InputObject $pcObject -MemberType NoteProperty -Name Serienummer_G -Value $serial
                Add-Member -InputObject $pcObject -MemberType NoteProperty -Name Funktionskonto_G -Value $adVarde
                Add-Member -InputObject $pcObject -MemberType NoteProperty -Name Lokal_admin_G -Value $lokaladmin
                #>
                $tempQueue = $using:PCObjects
                $tempQueue.Enqueue(<#$pcObject#>[PSCustomObject]@{
                    Hardvara_G = $hardware
                    Operativsystem_G = $os
                    Modell_G = $model
                    Roll_G = $filteredName
                    Hardvarunamn_G = $pcName
                    MAC_Adress_G = $mac
                    MAC_med_Kolon_G = $macColon
                    Serienummer_G = $serial
                    Funktionskonto_G = $adVarde
                    Lokal_admin_G = $lokaladmin
                })

            }
            Remove-Variable -Name "save"
            Remove-Variable -Name "PCObject"
            Remove-Variable -Name "hardware"
            Remove-Variable -Name "tempQueue"
            Remove-Variable -Name "os"
            Remove-Variable -Name "model"
            Remove-Variable -Name "filteredName"
            Remove-Variable -Name "pcName"
            Remove-Variable -Name "mac"
            Remove-Variable -Name "macColon"
            Remove-Variable -Name "serial"
            Remove-Variable -Name "adVarde"
            Remove-Variable -Name "lokaladmin"
            Remove-Variable -Name "bit"
            Remove-Variable -Name "pc_role"
            Remove-Variable -Name "Responce"
            Remove-Variable -Name "id"
            Remove-Variable -Name "request"
            Remove-Variable -Name "pc_model"
            Remove-Variable -Name "requestbody"
            Remove-Variable -Name "result"
            Remove-Variable -Name "url"
            Remove-Variable -Name "continue"
        } 
    } 
}
#$pathToCsv = Read-Host "Enter full csv file path to go thru (Without quotation marks)"
#$filename = Read-Host "FILNAMN?"
Write-Host "Din fil kommer att sparas i" $PSScriptRoot "med namnet" $filename".csv"
$stopWatchTotal = [System.Diagnostics.Stopwatch]::StartNew()
#$stopWatchOuter = [System.Diagnostics.Stopwatch]::StartNew()
$pc_modelsList = Import-Csv -Delimiter ";" -Path $PSScriptRoot\all_pc_models.csv -Header 'id','hv_typ','hv_category'
$rolesList = Import-Csv -Delimiter ";" -Path $PSScriptRoot\roles.csv -Header 'id','role'
$aioList = Import-Csv -Delimiter "," -Path $PSScriptRoot\all_touch_aio.csv -Header 'name'
$aioSerialList69 = Import-Csv -Delimiter "," -Path $PSScriptRoot\all_old_M725s.csv -Header 'serial'
$aioSerialList107 = Import-Csv -Delimiter "," -Path $PSScriptRoot\all_old_M75s1.csv -Header 'serial'
$pcList = Import-Csv -Delimiter ";" -Path $pathToCsv -Header 'pcName' -Encoding UTF8
#$running = $false

#Write-Host "List import" $stopWatchOuter.Elapsed.TotalMilliseconds


$job = $pcList | ForEach-Object -AsJob -ThrottleLimit 48 -Parallel $findPC 
while ($job.State -eq "Running" -or $PCObjects.Count -gt 0) {
    if ($PcObjects.Count -gt 0) {
        $tempObj = New-Object -TypeName PSObject
        if ($PcObjects.TryDequeue([ref]$tempObj)) {
            $tempObj | Export-Csv -Path ($PSScriptRoot + "\" + $filename + ".csv") -NoTypeInformation -Append -Force -Delimiter ";" -Encoding UTF8
        }
        Remove-Variable -Name "tempObj"
    }
}
#Export-Csv -InputObject $job -Path ($PSScriptRoot + "\" + $filename + ".csv") -NoTypeInformation -Append -Force -Delimiter ";" -Encoding UTF8 | Wait-Job -Any | Receive-Job
#$stopWatchOuter.Restart()
#foreach ($object in $PCObjects.ToArray()) { $object | Export-Csv -Path ($PSScriptRoot + "\" + $filename + ".csv") -NoTypeInformation -Append -Force -Delimiter ";" -Encoding UTF8 }
#Write-Host "File saving:" $stopWatchOuter.Elapsed.TotalMilliseconds
Write-Host "Total:" $stopWatchTotal.Elapsed.TotalMinutes
Write-Host "Done"
#Get-Job | Stop-Job
#Get-Job | Remove-Job
#  C:\Users\gaisysd8bp\Desktop\NewScript\test.csv