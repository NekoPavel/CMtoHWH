Param(
    [Parameter(Mandatory=$false)] [string] $Loc,
    [Parameter(Mandatory=$false)] [string] $currentUserDirectory
)

if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole] 'Administrator')
) {
    Write-Host "Not elevated, restarting..."
    #$Loc = $Loc.Substring(1, $Loc.Length - 2)
    [string]$currentUserDirectory = $ENV:USERNAME.ToString()

    $Arguments = @(
        '-NoProfile',
        '-ExecutionPolicy Bypass',
        '-File',
        "`"$($MyInvocation.MyCommand.Path)`"",
        "\`"$Loc\`"`"",
        "\`"$currentUserDirectory\`""
    )
    Start-Process -FilePath PowerShell.exe -Verb RunAs -ArgumentList $Arguments
    Break
}
else {
    Write-Host "Already elevated, exiting..."
}
if ($Loc.Length -gt 1) {
    Set-Location $Loc.Substring(1, $Loc.Length - 2)
}
Write-Host "Kollar om du har PowerShell7"
if (!(Test-Path .\pwsh\pwsh.exe)) {
    Write-Host "PowerShell saknas, laddar ner (Detta tar en stund)"
    Copy-Item -Path '\\dfs\Gem$\Lit\IT-Service\G55\Charon\pwsh' -Destination .\ -Force -Recurse
}

Write-Host "Kollar om du har RSAT"
if (!(Test-Path .\RSAT.msu)) {
    Write-Host "RSAT saknas, laddar ner (Detta tar en stund)"
    Copy-Item -Path '\\dfs\Gem$\Lit\IT-Service\G55\Charon\RSAT.msu' -Destination .\RSAT.msu -Force 
}

Write-Host "Kollar om du har scriptet eller om det finns en uppdatering"
if (!(Test-Path .\Charon.ps1)) {
    Write-Host "Scriptet saknas, laddar ner"
    Copy-Item -Path '\\dfs\Gem$\Lit\IT-Service\G55\Charon\Charon.ps1' -Destination .\Charon.ps1 -Force
}
elseif (!(Get-FileHash .\Charon.ps1).Hash -eq (Get-FileHash '\\dfs\Gem$\Lit\IT-Service\G55\Charon\Charon.ps1').Hash) {
    Write-Host "Uppdatering hittad, laddar ner"
    Copy-Item -Path '\\dfs\Gem$\Lit\IT-Service\G55\Charon\Charon.ps1' -Destination .\Charon.ps1 -Force
}

Write-Host "Kollar om du har fkar-scriptet eller om det finns en uppdatering"
if (!(Test-Path .\fkarfinder.ps1)) {
    Write-Host "Scriptet saknas, laddar ner"
    Copy-Item -Path '\\dfs\Gem$\Lit\IT-Service\G55\Charon\fkarfinder.ps1' -Destination .\fkarfinder.ps1 -Force
}
elseif (!(Get-FileHash .\fkarfinder.ps1).Hash -eq (Get-FileHash '\\dfs\Gem$\Lit\IT-Service\G55\Charon\fkarfinder.ps1').Hash) {
    Write-Host "Uppdatering hittad, laddar ner"
    Copy-Item -Path '\\dfs\Gem$\Lit\IT-Service\G55\Charon\fkarfinder.ps1' -Destination .\fkarfinder.ps1 -Force
}

Write-Host "Kollar om du har ikonen"
if (!(Test-Path .\icon.png)) {
    Write-Host "Ikonen saknas, laddar ner"
    Copy-Item -Path '\\dfs\Gem$\Lit\IT-Service\G55\Charon\icon.png' -Destination .\icon.png -Force
}
Write-Host "Kollar om RSAT är installerat"
try {
    Get-ADUser $currentUserDirectory.Substring(1,$currentUserDirectory.Length -2) | Out-Null
}
catch {
    
}
if (Get-Module -Name ActiveDirectory) {
    Write-Host "RSAT är redan installerat, fortsätter"
}
else {
    Write-Host "RSAT saknas, installerar. Detta tar en stund"
    $LocRSAT=$Loc+"RSAT.msu`""
    Start-Process -FilePath $LocRSAT -ArgumentList "/quiet /norestart" -Wait
}

Write-Host "Startar Scriptet"
$LocScript = $Loc+"Charon.ps1`""
.\pwsh\pwsh.exe -File $LocScript $currentUserDirectory