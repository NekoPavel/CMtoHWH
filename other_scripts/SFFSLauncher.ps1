Param(
    [string]$Loc,
    [string]$currentUserDirectory
)

$Delay = 0

if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole] 'Administrator')
) {
    Write-Host "Not elevated, restarting in $Delay seconds ..."
    $Loc = Get-Location
    [string]$currentUserDirectory = $ENV:USERPROFILE.ToString()
    Start-Sleep -Seconds $Delay

    $Arguments = @(
        '-NoProfile',
        '-ExecutionPolicy Bypass',
        '-NoExit',
        '-File',
        "`"$($MyInvocation.MyCommand.Path)`"",
        "\`"$Loc\`"",
        "\`"$currentUserDirectory\`""
    )
    Start-Process -FilePath PowerShell.exe -Verb RunAs -ArgumentList $Arguments
    Break
}
else {
    Write-Host "Already elevated, exiting in $Delay seconds..."
    Start-Sleep -Seconds $Delay
}
if ($Loc.Length -gt 1) {
    Set-Location $Loc.Substring(1, $Loc.Length - 2)
}

Write-Host "Kollar om du har PowerShell7"
if (!(Test-Path .\pwsh\pwsh.exe)) {
    Write-Host "PowerShell saknas, laddar ner (Detta tar en stund, vänligen vänta)"
    Copy-Item -Path '\\dfs\Gem$\Lit\IT-Service\G55\SFFS\pwsh' -Destination .\ -Force -Recurse
}

Write-Host "Kollar om du har scriptet eller om det finns en uppdatering"
if (!(Test-Path .\FindPC.ps1)) {
    Write-Host "Scriptet saknas, laddar ner"
    Copy-Item -Path '\\dfs\Gem$\Lit\IT-Service\G55\SFFS\FindPC.ps1' -Destination .\FindPC.ps1 -Force
}
elseif (!(Get-FileHash .\FindPC.ps1).Hash -eq (Get-FileHash '\\dfs\Gem$\Lit\IT-Service\G55\SFFS\FindPC.ps1').Hash) {
    Write-Host "Uppdatering hittad, laddar ner"
    Copy-Item -Path '\\dfs\Gem$\Lit\IT-Service\G55\SFFS\FindPC.ps1' -Destination .\FindPC.ps1 -Force
}

Write-Host "Kollar om du har ikonen"
if (!(Test-Path .\icon.png)) {
    Write-Host "Ikonen saknas, laddar ner"
    Copy-Item -Path '\\dfs\Gem$\Lit\IT-Service\G55\SFFS\icon.ps1' -Destination .\icon.ps1 -Force
}
Write-Host "Startar Scriptet"

.\pwsh\pwsh.exe $Loc\SFFS.ps1 $currentUserDirectory