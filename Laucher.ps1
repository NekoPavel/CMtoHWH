If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {  
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    Break
}
Write-Host "Kollar om du har PowerShell7"
if (!(Test-Path .\pwsh\pwsh.exe)) {
    Write-Host "PowerShell saknas, laddar ner (Detta kan ta en stund)"
    Copy-Item -Path '\\dfs\Gem$\Lit\IT-Service\G55\HittaPC\pwsh' -Destination .\ -Force -Recurse
}
Write-Host "Kollar om du har startfilen eller om det finns en uppdatering"
if (!(Test-Path ./HittaDator.bat)) {
    Write-Host "Filen saknas, laddar ner"
    Copy-Item -Path '\\dfs\Gem$\Lit\IT-Service\G55\HittaPC\HittaDator.bat' -Destination .\HittaDator.bat -Force
}
elseif(!((Get-FileHash .\HittaDator.bat).Hash -eq (Get-FileHash '\\dfs\Gem$\Lit\IT-Service\G55\HittaPC\HittaDator.bat').Hash))
{
    Write-Host "Uppdatering hittad, laddar ner"
    Copy-Item -Path '\\dfs\Gem$\Lit\IT-Service\G55\HittaPC\HittaDator.bat' -Destination .\HittaDator.bat -Force
}
Write-Host "Kollar om du har scriptet eller om det finns en uppdatering"
if (!(Test-Path .\FindPC.ps1)) {
    Write-Host "Scriptet saknas, laddar ner"
    Copy-Item -Path '\\dfs\Gem$\Lit\IT-Service\G55\HittaPC\FindPC.ps1' -Destination .\FindPC.ps1 -Force
}
elseif(!((Get-FileHash .\FindPC.ps1).Hash -eq (Get-FileHash '\\dfs\Gem$\Lit\IT-Service\G55\HittaPC\FindPC.ps1').Hash))
{
    Write-Host "Uppdatering hittad, laddar ner"
    Copy-Item -Path '\\dfs\Gem$\Lit\IT-Service\G55\HittaPC\FindPC.ps1' -Destination .\FindPC.ps1 -Force
}
Write-Host "Startar Scriptet"
.\HittaDator.bat