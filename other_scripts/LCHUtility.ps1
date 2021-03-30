If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {  
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    Break
}
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

$pathToCsv = Read-Host "Enter full csv file path to go thru (Without quotation marks)"
$pcList = Import-Csv -Delimiter ";" -Path $pathToCsv -Encoding UTF8 -Header 'OldPCName', 'OldUsername', 'OldSerial', 'NewPCName', 'NewUsername', 'NewSerial'
foreach ($pc in $pcList) {
    $urlNew = "http://sysman.sll.se/SysMan/api/Client?name=" + $pc.NewPCName + "&take=10&skip=0&type=0&targetActive=1"
    $ResponceNew = Invoke-WebRequest -Uri $urlNew -AllowUnencryptedAuthentication -UseDefaultCredentials -SessionVariable 'Session'
    $idNew = (($ResponceNew.Content | ConvertFrom-Json).result).id
    $urlOld = "http://sysman.sll.se/SysMan/api/Client?name=" + $pc.OldPCName + "&take=10&skip=0&type=0&targetActive=1"
    $ResponceOld = Invoke-WebRequest -Uri $urlOld -AllowUnencryptedAuthentication -UseDefaultCredentials -WebSession $Session
    $idOld = (($ResponceOld.Content | ConvertFrom-Json).result).id
    if (!($ResponceNew.Content | ConvertFrom-Json).result) {
        Write-Host "Datorn: $($pc.NewPCName) hittas ej."
    }
    elseif (!($ResponceOld.Content | ConvertFrom-Json).result) {
        Write-Host "Datorn: $($pc.OldPCName) hittas ej."
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
        $request = Invoke-WebRequest -Method Post -Uri "http://sysman.sll.se/SysMan/api/v2/printer/install" -Body $requestBody -AllowUnencryptedAuthentication -UseDefaultCredentials -WebSession $Session -ContentType "application/json"
        Write-Host "Datorn: $($pc.NewPCName) har fått $($printers.Count) skrivare från $($pc.OldPCName)."

        $fkonto = fkarfinder -pcName $pc.OldPCName
        if ($null -ne $fkonto.cn) { 
            $fkonto = ($fkonto.cn -join ', ')
            Write-Host "Hittat funktionskonto $($fkonto)"
            $logOnTo = (Get-ADUser -Identity $fkonto -Properties *).userWorkstations + ","+$pc.NewPCName 
            $fkonto | Set-ADUser -LogonWorkstations $logOnTo
            $pcIdentity = Get-ADComputer -Identity $pc.NewPCName
            Write-Host "$($fkonto) fungerar nu på : $($logOnTo)"
            Add-ADGroupMember -Identity "$($pc.NewPCName.Substring(0,3))_Wrk_F-kontoWS_SLLeKlient" -Members $pcIdentity         
            Write-Host "Dator tillagd i $($pc.NewPCName.Substring(0,3))_Wrk_F-kontoWS_SLLeKlient"
        }
        else {
            Write-Host "Dator har inte Funktionskonto"
        }
    }
    

}