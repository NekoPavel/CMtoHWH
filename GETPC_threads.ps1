param([parameter(Mandatory = $true)][string]$pathToCsv, [parameter(Mandatory = $true)][string]$filename)
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    break
}

#$excludedModels = @("Virtual Machine","VMware Virtual Platform","Parallels Virtual Platform","OEM")
#$excludedModelsRegex = [string]::Join('|',$excludedModels)
$PCObjects = (New-Object System.Collections.Concurrent.ConcurrentQueue[PSCustomObject])
$ModelsQueue = (New-Object System.Collections.Concurrent.ConcurrentQueue[String])
$findPC = {
    $pc = $_.pcName
    if ($pc -imatch "^[KLSHIFRTPD][AIOSNUELKRT][RTSFVLAN]((LS)|(DS))\d{8}") {
        $url = "http://sysman.sll.se/SysMan/api/Client?name=" + $pc + "&take=1&skip=0&type=0"
        $Responce = Invoke-WebRequest -Uri $url -UseDefaultCredentials -AllowUnencryptedAuthentication -SessionVariable 'Session'
        #if (!($Responce.Content | ConvertFrom-Json).result) {}
        <#else#>
        if (($Responce.Content | ConvertFrom-Json).result) {
            $save = $true
            $pcName = (($Responce.Content | ConvertFrom-Json).result).Name

            $fkarJob = Start-ThreadJob -InputObject $pcName -ScriptBlock {
                $adVarde = (.\AdFind.exe -h gaia -f "(&(objectCategory=user)(userworkstations=*$input*)(cn=F*))" cn userworkstations -nodn -incllike cn)
                if ($adVarde[$advarde.Length-1].Substring(0,1) -ne "0") {
                    $adVarde = ($adVarde[$advarde.Length-3]).Substring(5,8)
                }
                else {
                    $adVarde = "NEJ"
                }
                Write-Output $adVarde
                Remove-Variable -Name "adVarde"
            }

            $adminJob = Start-ThreadJob -InputObject $pcName -ScriptBlock {
                $lokaladmin = ([regex]::match((.\AdFind.exe -h gaia -f "(cn=$input)" cn MemberOf -nodn -incllike MemberOf),"_Wrk_LocalAdmin_SLLeKlient")).Success
                if ($lokaladmin) {
                    $lokaladmin = "JA"
                }
                else {
                    $lokaladmin = "NEJ"
                }
                Write-Output $lokaladmin
                Remove-Variable -Name "lokaladmin"
            }
            $id = (($Responce.Content | ConvertFrom-Json).result).Id
            #Modell
            $model = ((($Responce.Content | ConvertFrom-Json).result).hardwareModel).name

            $request = "http://sysman.sll.se/SysMan/api/client?id=" + $id + "&name=" + $pcName + "&assetTag=" + $pcName
            $Responce = ((Invoke-WebRequest -Uri $request -AllowUnencryptedAuthentication -WebSession $Session).Content | ConvertFrom-Json)
            #NEW MAC
            [string]$macColon = [string]$Responce.macAddress
            [string]$mac = [string]([string]$macColon -replace ":", "")
            #Serienummer
            [string]$serial = [string]$Responce.serialNumber
            
            $request = "http://sysman.sll.se/SysMan/api/Reporting/Client?clientId=" + $id
            $Responce = (Invoke-WebRequest -Uri $request -AllowUnencryptedAuthentication -WebSession $Session).Content | ConvertFrom-Json
            #Serienummer Påfyllning
            if (-not ($serial.Length -gt 0)) {
                [string]$serial = [string]$Responce.serial
            }
            
            if ($serial -like "*O.E.M.*" -or $serial.Length -lt 1) {
                $save = $false
            }

            #Mer modellsaker
            foreach ($pc_model in $using:pc_modelsList) {
                if ($pc_model.hv_typ -ieq $model) {
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
                            if (!($aioSerial -ieq $serial)) {
                                $model = "106"
                            }
                        }
                        # M725s
                    }
                    if ($model -eq "107") {
                        #Om datorn INTE finns i listan så är det en av de nya
                        foreach ($aioSerial in $using:aioSerialList107) {
                            if (!($aioSerial -ieq $serial)) {
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
                
                #$tempModelQueue = $using:ModelsQueue
                #$tempModelQueue.Enqueue("$model : $pcName")
                
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
            
            
            $role = "Inte hittad"

            $tempRole = ($Responce.collections | Where-Object { $_ -match "[A-Z][a-z]{2}_Wrk_PR" })
            if ($tempRole) {
                $role = $tempRole.Substring(11)
            }


            $tempRole = ($Responce.installedApplications | Where-Object { $_.Name -match "Sll_Wrk_[A-Z][a-z]{2}_PR" })
            if ($tempRole) {
                $role = $tempRole.Name.Substring(15)
            }
            
            if (($role -eq "Kiosk_PC" ) -or ($role -eq "Inte hittad")) {
                if (($Responce.installedApplications | Where-Object { $_.Name -eq "Kar_Rol_Vardterminal-Kiosk" })) {
                    $role = "Vardterminal"
                }
                elseif ($Responce.collections -match "Gai_App_CitrixReceiverVardTerminal") {
                    $role = "Vardterminal"
                }
            }
            foreach ($pc_role in $using:rolesList) {
                if ($pc_role.role -ieq $role) {
                    $role = $pc_role.Id
                }
            }
            if (($role -eq "1") -and ($bit -like "*64*")) {
                $role = "2"
            }
            if ($role -ieq "Inte hittad" -or $role.Length -gt 2 -or $role -gt 11) {
                $save = $false
            }
            #Funktionskonto och lokaladmin
            $fkontoResult = Receive-Job -Job $fkarJob -Wait -AutoRemoveJob
            $lokaladmin = Receive-Job -Job $adminJob -Wait -AutoRemoveJob
            
            
            if ($save) {
                $tempQueue = $using:PCObjects
                $tempQueue.Enqueue([PSCustomObject]@{
                        Hardvara_G       = $hardware
                        Operativsystem_G = $os
                        Modell_G         = $model
                        Roll_G           = $role
                        Hardvarunamn_G   = $pcName
                        MAC_Adress_G     = $mac
                        MAC_med_Kolon_G  = $macColon
                        Serienummer_G    = $serial
                        Funktionskonto_G = $fkontoResult
                        Lokal_admin_G    = $lokaladmin
                    })
            }
            else {
                $tempQueue = $using:PCObjects
                $tempQueue.Enqueue([PSCustomObject]@{
                        Hardvara_G = "filler"
                    })
            }
            Remove-Variable -Name "save"
            Remove-Variable -Name "hardware"
            Remove-Variable -Name "tempQueue"
            Remove-Variable -Name "os"
            Remove-Variable -Name "model"
            Remove-Variable -Name "role"
            Remove-Variable -Name "pcName"
            Remove-Variable -Name "mac"
            Remove-Variable -Name "macColon"
            Remove-Variable -Name "serial"
            Remove-Variable -Name "lokaladmin"
            Remove-Variable -Name "bit"
            Remove-Variable -Name "pc_role"
            Remove-Variable -Name "Responce"
            Remove-Variable -Name "id"
            Remove-Variable -Name "request"
            Remove-Variable -Name "pc_model"
            Remove-Variable -Name "url"
            Remove-Variable -Name "continue"
            Remove-Variable -Name "tempModelQueue"
            Remove-Variable -Name "adminJob"
            Remove-Variable -Name "fkarJob"
            Remove-Variable -Name "fkontoResult"
            
        } 
    }
    else {
        $tempQueue = $using:PCObjects
        $tempQueue.Enqueue([PSCustomObject]@{
                Hardvara_G = "filler"
            })
    }
}
#$pathToCsv = Read-Host "Enter full csv file path to go thru (Without quotation marks)"
#$filename = Read-Host "FILNAMN?"
Write-Host "Din fil kommer att sparas i" $PSScriptRoot "med namnet" $filename".csv"
$stopWatchTotal = [System.Diagnostics.Stopwatch]::StartNew()
#$stopWatchOuter = [System.Diagnostics.Stopwatch]::StartNew()
$pc_modelsList = Import-Csv -Delimiter ";" -Path $PSScriptRoot\all_pc_models.csv -Header 'id', 'hv_typ', 'hv_category'
$rolesList = Import-Csv -Delimiter ";" -Path $PSScriptRoot\roles.csv -Header 'id', 'role'
$aioList = Import-Csv -Delimiter "," -Path $PSScriptRoot\all_touch_aio.csv -Header 'name'
$aioSerialList69 = Import-Csv -Delimiter "," -Path $PSScriptRoot\all_old_M725s.csv -Header 'serial'
$aioSerialList107 = Import-Csv -Delimiter "," -Path $PSScriptRoot\all_old_M75s1.csv -Header 'serial'
$pcList = Import-Csv -Delimiter ";" -Path $pathToCsv -Header 'pcName' -Encoding UTF8
#$running = $false

#Write-Host "List import" $stopWatchOuter.Elapsed.TotalMilliseconds

#$totalCount = $pcList.Length
#$currentStep = 0
#Write-Host "$totalCount rader ska köras"
$stopWatchTotal = [System.Diagnostics.Stopwatch]::StartNew()
$job = $pcList | ForEach-Object -AsJob -ThrottleLimit 8 -Parallel $findPC  
while ($job.State -eq "Running" -or $PCObjects.Count -gt 0 -or $ModelsQueue -gt 0) {
    if ($PcObjects.Count -gt 0) {
        $tempObj = New-Object -TypeName PSObject
        if ($PcObjects.TryDequeue([ref]$tempObj)) {
            if ($tempObj.Hardvara_G -ne "filler") {
                $tempObj | Export-Csv -Path ($PSScriptRoot + "\" + $filename + ".csv") -NoTypeInformation -Append -Force -Delimiter ";" -Encoding UTF8
            }
            <#
            $currentStep++
            Clear-Host
            Write-Host "Steg $currentStep av $totalCount"
            $perc=($currentStep/$totalCount)*100
            Write-Host "$perc% färdigt"
            $timeOne = $stopWatchTotal.Elapsed.TotalMinutes/$currentStep
            $secondsOne = $timeOne * 60
            $computersSecond = [math]::round(1.0/$secondsOne)
            $timeLeft = ($timeOne*$totalCount)-$stopWatchTotal.Elapsed.TotalMinutes
            Write-Host "Ungefär $timeLeft minuter kvar ($secondsOne sekunder/dator) ($computersSecond datorer/sekund)"
            Remove-Variable -Name "perc"
            Remove-Variable -Name "timeLeft"
            Remove-Variable -Name "timeOne"
            Remove-Variable -Name "secondsOne"
            Remove-Variable -Name "computersSecond"
            #>
        }
        Remove-Variable -Name "tempObj"
    }
    if ($ModelsQueue.Count -gt 0) {
        $tempModel = ""
        if ($ModelsQueue.TryDequeue([ref]$tempModel)) {
            $tempModel | Out-File -FilePath ($PSScriptRoot + "\unmappedModelsLog.txt") -Append
        }
        Remove-Variable -Name "tempModel"
    }
}

Write-Host "Total:" $stopWatchTotal.Elapsed.TotalMinutes "minuter"
Write-Host "Done"