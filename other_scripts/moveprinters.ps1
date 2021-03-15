If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {  
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    Break
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
        targets  = $idNewList
        printers = $printers
        templateTargetId = $null
    } | ConvertTo-Json -Compress
    Invoke-WebRequest -Method Post -Uri "http://sysman.sll.se/SysMan/api/v2/printer/install" -Body $requestBody -AllowUnencryptedAuthentication -UseDefaultCredentials -WebSession $Session -ContentType "application/json"
}