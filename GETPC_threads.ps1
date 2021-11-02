param([parameter(Mandatory = $true)][string]$pathToCsv, [parameter(Mandatory = $true)][string]$filename)
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    break
}
function Find-ADObjects($domain, $class, $filter, $attributes = "distinguishedName") {
    $dc = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext ([System.DirectoryServices.ActiveDirectory.DirectoryContextType]"domain", $domain);
    $dn = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($dc);
    
    $ds = New-Object System.DirectoryServices.DirectorySearcher;
    $ds.SearchRoot = $dn.GetDirectoryEntry();
    $ds.SearchScope = "subtree";
    $ds.PageSize = 1024;
    $ds.Filter = "(&(objectCategory=$class)$filter)";
    $ds.PropertiesToLoad.AddRange($attributes.Split(","))
    $result = $ds.FindAll();
    $ds.Dispose();
    return $result;
}
function GetLocalAdmin {
    param (
        [string]$computerName
    )
    $adminRoles = @("CN=Pnf_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Pnf,OU=HealthCare,DC=gaia,DC=sll,DC=se", "CN=Kar_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Kar,OU=HealthCare,DC=gaia,DC=sll,DC=se", "CN=Sos_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Sos,OU=HealthCare,DC=gaia,DC=sll,DC=se", "CN=Lit_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Lit,OU=Administration,DC=gaia,DC=sll,DC=se", "CN=Ita_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Ita,OU=Reference,DC=gaia,DC=sll,DC=se", "CN=Dan_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Dan,OU=HealthCare,DC=gaia,DC=sll,DC=se", "CN=Hsf_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Hsf,OU=Administration,DC=gaia,DC=sll,DC=se", "CN=Lsf_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Lsf,OU=Administration,DC=gaia,DC=sll,DC=se", "CN=Fut_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Fut,OU=PublicTransportation,DC=gaia,DC=sll,DC=se", "CN=Int_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Int,OU=Administration,DC=gaia,DC=sll,DC=se", "CN=Trf_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Trf,OU=Administration,DC=gaia,DC=sll,DC=se", "CN=Sll_Wrk_LocalAdmin_SLLeKlient,OU=Workstation,OU=Groups,OU=Sll,DC=gaia,DC=sll,DC=se", "CN=Ser_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Ser,OU=Administration,DC=gaia,DC=sll,DC=se")
    $adminRolesRegex = [string]::Join('|', $adminRoles)
    $adVarde = (Find-ADObjects "gaia" "computer" "(cn=$computerName)" "cn,MemberOf").Properties
    if ($adVarde.memberof -match $adminRolesRegex) {
        $true
    }
    else {
        $false
    }
}
function FindFunkAccount {
    param (
        [string]$computerName
    )
    (Find-ADObjects "gaia" "user" "(userworkstations=*$computerName*)(cn=F*)" "cn,userworkstations").Properties
}


#$excludedModels = @("Virtual Machine","VMware Virtual Platform","Parallels Virtual Platform","OEM")
#$excludedModelsRegex = [string]::Join('|',$excludedModels)
$PCObjects = (New-Object System.Collections.Concurrent.ConcurrentQueue[PSCustomObject])
$ModelsQueue = (New-Object System.Collections.Concurrent.ConcurrentQueue[String])
$findPC = {
    $pc = $_.pcName
    #$continue = $false
    <#
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
    #>
    if ($pc -imatch "^[KLSHIFRTPD][AIOSNUELKRT][RTSFVLAN]((LS)|(DS))\d{8}") {
        $url = "http://sysman.sll.se/SysMan/api/Client?name=" + $pc + "&take=1&skip=0&type=0"
        $Responce = Invoke-WebRequest -Uri $url -UseDefaultCredentials -AllowUnencryptedAuthentication -SessionVariable 'Session'
        #if (!($Responce.Content | ConvertFrom-Json).result) {}
        <#else#>if (($Responce.Content | ConvertFrom-Json).result) {
            $save = $true
            $pcName = (($Responce.Content | ConvertFrom-Json).result).Name
            $id = (($Responce.Content | ConvertFrom-Json).result).Id
            #Modell
            $model = ((($Responce.Content | ConvertFrom-Json).result).hardwareModel).Name
            <#$requestBody =
            @{
                UserName = ""
                ComputerName = $pcName.ToString()
                Id = "88eeae01-fc85-426c-898d-dae73ec31867"
            } | ConvertTo-Json -Compress
            $macResponce = Invoke-WebRequest -Method Post -Uri "http://sysman.sll.se/SysMan/api/Tool/Run" -Body $requestBody -WebSession $Session -ContentType "application/json" -AllowUnencryptedAuthentication
            $macResponce = ($macResponce | ConvertFrom-Json).result
            #MAC
            foreach ($result in $macResponce) {
                if ((($result -like "*Ethernet*") -and !($result -like "*#2*") -and !($result -like "*Virtual*") -and !($result -like "*Server Adapter*") -and !($result -like "*Dock*")) -or (($result -like "*GbE*") -and !($result -like "*#2*") -and !($result -like "*USB*")) -or $result -like "*Gigabit*") {
                    [string]$macColon = [string]$result.Substring(0,17)
                    [string]$mac = [string]([string]$macColon -replace ":","")
                }
            }
            if (!$mac -or !$macColon) {
                $save = $false
            }
            #>
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
                #TODO Make this log
                $tempModelQueue = $using:ModelsQueue
                $tempModelQueue.Enqueue($model)
                
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
            
            
            $filteredName = "Inte hittad"

            $tempName = ($Responce.collections | Where-Object { $_ -match "[A-Z][a-z]{2}_Wrk_PR" })
            if ($tempName) {
                $filteredName = $tempName.Substring(11)
            }


            $tempName = ($Responce.installedApplications | Where-Object { $_.Name -match "Sll_Wrk_[A-Z][a-z]{2}_PR" })
            if ($tempName) {
                $filteredName = $tempName.Name.Substring(15)
            }
            
            if (($filteredName -eq "Kiosk_PC" ) -or ($filteredName -eq "Inte hittad")) {
                if (($Responce.installedApplications | Where-Object { $_.Name -eq "Kar_Rol_Vardterminal-Kiosk" })) {
                    $filteredName = "Vardterminal"
                }
                elseif ($Responce.collections -match "Gai_App_CitrixReceiverVardTerminal") {
                    $filteredName = "Vardterminal"
                }
            }
            foreach ($pc_role in $using:rolesList) {
                if ($pc_role.role -ieq $filteredName) {
                    $filteredName = $pc_role.Id
                }
            }
            if (($filteredName -eq "1") -and ($bit -like "*64*")) {
                $filteredName = "2"
            }
            if ($filteredName -ieq "Inte hittad" -or $filename.Length -gt 1) {
                $save = $false
            }
            #Funktionskonto
            $adVarde = FindFunkAccount($pcName)
            if ($null -ne $adVarde.cn) { $adVarde = ($adVarde.cn -join ', ') }
            else {
                $adVarde = "NEJ"
            }
            $lokaladmin = GetLocalAdmin($pcName)
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
                        Hardvara_G       = $hardware
                        Operativsystem_G = $os
                        Modell_G         = $model
                        Roll_G           = $filteredName
                        Hardvarunamn_G   = $pcName
                        MAC_Adress_G     = $mac
                        MAC_med_Kolon_G  = $macColon
                        Serienummer_G    = $serial
                        Funktionskonto_G = $adVarde
                        Lokal_admin_G    = $lokaladmin
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
            Remove-Variable -Name "tempModelQueue"
        } 
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


$job = $pcList | ForEach-Object -AsJob -ThrottleLimit 48 -Parallel $findPC 
while ($job.State -eq "Running" -or $PCObjects.Count -gt 0 -or $ModelsQueue -gt 0) {
    if ($PcObjects.Count -gt 0) {
        $tempObj = New-Object -TypeName PSObject
        if ($PcObjects.TryDequeue([ref]$tempObj)) {
            $tempObj | Export-Csv -Path ($PSScriptRoot + "\" + $filename + ".csv") -NoTypeInformation -Append -Force -Delimiter ";" -Encoding UTF8
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
#Export-Csv -InputObject $job -Path ($PSScriptRoot + "\" + $filename + ".csv") -NoTypeInformation -Append -Force -Delimiter ";" -Encoding UTF8 | Wait-Job -Any | Receive-Job
#$stopWatchOuter.Restart()
#foreach ($object in $PCObjects.ToArray()) { $object | Export-Csv -Path ($PSScriptRoot + "\" + $filename + ".csv") -NoTypeInformation -Append -Force -Delimiter ";" -Encoding UTF8 }
#Write-Host "File saving:" $stopWatchOuter.Elapsed.TotalMilliseconds
Write-Host "Total:" $stopWatchTotal.Elapsed.TotalMinutes "minuter"
Write-Host "Done"
#Get-Job | Stop-Job
#Get-Job | Remove-Job
#  C:\Users\gaisysd8bp\Desktop\NewScript\test.csv