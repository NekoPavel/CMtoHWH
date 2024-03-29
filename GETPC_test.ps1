$pc_modelsList = Import-Csv -Delimiter ";" -Path $PSScriptRoot\all_pc_models.csv -Header 'id', 'hv_typ', 'hv_category'
$rolesList = Import-Csv -Delimiter ";" -Path $PSScriptRoot\roles.csv -Header 'id', 'role'
$aioList = Import-Csv -Delimiter "," -Path $PSScriptRoot\all_touch_aio.csv -Header 'name'
$aioSerialList69 = Import-Csv -Delimiter "," -Path $PSScriptRoot\all_old_M725s.csv -Header 'serial'
$aioSerialList107 = Import-Csv -Delimiter "," -Path $PSScriptRoot\all_old_M75s1.csv -Header 'serial'

$pc = "KARDS98449028"
if ($pc -imatch "^[KLSHIFRTPD][AIOSNUELKRT][RTSFVLAN]((LS)|(DS))\d{8}") {
    $url = "http://sysman.sll.se/SysMan/api/Client?name=" + $pc + "&take=1&skip=0&type=0"
    $Responce = Invoke-WebRequest -Uri $url -UseDefaultCredentials -AllowUnencryptedAuthentication -SessionVariable 'Session'
    #if (!($Responce.Content | ConvertFrom-Json).result) {}
    <#else#>
    if (($Responce.Content | ConvertFrom-Json).result) {

        $pcName = (($Responce.Content | ConvertFrom-Json).result).Name

        $fkarJob = Start-Job -ScriptBlock { 
            <#function Find-ADObjects($domain, $class, $filter, $attributes = "distinguishedName") {
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
            (Find-ADObjects "gaia" "user" "(userworkstations=*$pcName*)(cn=F*)" "cn,userworkstations").Properties
            #>
            (dsquery * -filter '(&(objectCategory=user)(userworkstations=*KARDS98449028*)(cn=F*))' -attr cn userworkstations).Properties
        }

        $adminJob = Start-Job -ScriptBlock {
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
            $adminRoles = @("CN=Pnf_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Pnf,OU=HealthCare,DC=gaia,DC=sll,DC=se", "CN=Kar_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Kar,OU=HealthCare,DC=gaia,DC=sll,DC=se", "CN=Sos_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Sos,OU=HealthCare,DC=gaia,DC=sll,DC=se", "CN=Lit_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Lit,OU=Administration,DC=gaia,DC=sll,DC=se", "CN=Ita_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Ita,OU=Reference,DC=gaia,DC=sll,DC=se", "CN=Dan_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Dan,OU=HealthCare,DC=gaia,DC=sll,DC=se", "CN=Hsf_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Hsf,OU=Administration,DC=gaia,DC=sll,DC=se", "CN=Lsf_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Lsf,OU=Administration,DC=gaia,DC=sll,DC=se", "CN=Fut_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Fut,OU=PublicTransportation,DC=gaia,DC=sll,DC=se", "CN=Int_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Int,OU=Administration,DC=gaia,DC=sll,DC=se", "CN=Trf_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Trf,OU=Administration,DC=gaia,DC=sll,DC=se", "CN=Sll_Wrk_LocalAdmin_SLLeKlient,OU=Workstation,OU=Groups,OU=Sll,DC=gaia,DC=sll,DC=se", "CN=Ser_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Ser,OU=Administration,DC=gaia,DC=sll,DC=se")
            $adminRolesRegex = [string]::Join('|', $adminRoles)
            $adVarde = (Find-ADObjects "gaia" "computer" "(cn=$pcName)" "cn,MemberOf").Properties
            if ($adVarde.memberof -match $adminRolesRegex) {
                $true
            }
            else {
                $false
            }
            Remove-Variable -Name "adminRoles"
            Remove-Variable -Name "adminRolesRegex"
            Remove-Variable -Name "adVarde"
        }

        $id = (($Responce.Content | ConvertFrom-Json).result).Id
        #Modell
        $model = ((($Responce.Content | ConvertFrom-Json).result).hardwareModel).Name
        
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

        #Mer modellsaker
        foreach ($pc_model in $pc_modelsList) {
            if ($pc_model.hv_typ -ieq $model) {
                $model = $pc_model.Id
                $hardware = $pc_model.hv_category
                if ($hardware -eq "4") {
                    foreach ($aio in $aioList) {
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
                    foreach ($aioSerial in $aioSerialList69) {
                        if (!($aioSerial -eq $serial)) {
                            $model = "106"
                        }
                    }
                    # M725s
                }
                if ($model -eq "107") {
                    #Om datorn INTE finns i listan så är det en av de nya
                    foreach ($aioSerial in $aioSerialList107) {
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
        foreach ($pc_role in $rolesList) {
            if ($pc_role.role -eq $filteredName) {
                $filteredName = $pc_role.Id
            }
        }
        if (($filteredName -eq "1") -and ($bit -like "*64*")) {
            $filteredName = "2"
        }
        if ($filteredName -eq "Inte hittad") {
        }
        #Funktionskonto
        $fkontoResult = Receive-Job -Job $fkarJob -Wait
        if ($null -ne $fkontoResult.cn) { $fkontoResult = ($fkontoResult.cn -join ', ') }
        else {
            $fkontoResult = "NEJ"
        }
        $lokaladmin = Receive-Job -Job $adminJob -Wait
        if ($lokaladmin) {
            $lokaladmin = "JA"
        }
        else {
            $lokaladmin = "NEJ"
        }
        if ($true) {
            Write-Host $hardware
            Write-Host $os
            Write-Host $model
            Write-Host $filteredName
            Write-Host $pcName
            Write-Host $mac
            Write-Host $macColon
            Write-Host $serial
            Write-Host $adVarde
            Write-Host $lokaladmin

            Remove-Variable -Name "hardware"
            Remove-Variable -Name "os"
            Remove-Variable -Name "model"
            Remove-Variable -Name "filteredName"
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
            Remove-Variable -Name "adminJob"
            Remove-Variable -Name "fkarJob"
            Remove-Variable -Name "fkontoResult"
        }
    } 
} 
