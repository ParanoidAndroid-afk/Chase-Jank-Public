#Downloads and installs Steam launcher
$TempDir = Join-Path $env:USERPROFILE "Downloads"
$SteamInstaller = Join-Path $TempDir "SteamSetup.exe"
$SteamURL = "https://cdn.akamai.steamstatic.com/client/installer/SteamSetup.exe"
Invoke-WebRequest -Uri $SteamURL -OutFile $SteamInstaller 
Try {
    Start-process -FilePath $SteamInstaller -ArgumentList "/S" -Wait -NoNewWindow
}
Catch {
    Set-Content -Path "$env:USERPROFILE\Downloads\CatchTest.txt" -Value $_.Exception.Message
}