$RemoteDesktop = Get-WmiObject -Class Win32_Product | Where-Object{$_.Name -eq 'Remote Desktop'} | select name
$isInstalled = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | where { $_.DisplayName -eq $RemoteDesktop }) -ne $null
if(-not $isInstalled) {
Write-Host 'Per machine Remote Desktop version is not installed, lates version is going to be downloaded and installed.' -BackgroundColor Red -ForegroundColor Yellow 
}
else {
Write-Host 'Uninstalling current Remote Desktop installation, please wait.' -BackgroundColor Green -ForegroundColor Black
$RemoteDesktop.Uninstall()
Write-Host 'Waiting 60 seconds..'
Start-Sleep -Seconds 60
}

$checkTempFolder = 'C:\temp'
#Test to see if C:\temp folder  exists
if (test-Path -Path $checkTempFolder) {
    Write-host 'Temp folder exists, script continues.' -BackgroundColor Green -ForegroundColor Black 
}
else {
    Write-host 'Temp folder does not exist!' -BackgroundColor Red -ForegroundColor Yellow 
    New-Item -Path 'C:\' -Name 'temp' -ItemType Directory
    Start-Sleep -Seconds 5
    Write-host 'Temp folder created' -BackgroundColor Green -ForegroundColor Black
}
Write-Host 'Downloading current Remote Desktop version, please wait.'
Invoke-WebRequest -Uri 'https://go.microsoft.com/fwlink/?linkid=2068602' -OutFile 'C:\temp\RemoteDesktop.msi'

Write-Host 'Remote Desktop downloaded, application is going to be installed. Please wait.'
Start-Process -FilePath '$env:systemroot\system32\msiexec.exe' -ArgumentList '/i 'C:\temp\RemoteDesktop.msi' /qn ALLUSERS=2 MSIINSTALLPERUSER=1'

Write-Host 'Adding registry, please wait.'
$registryKey1 = Get-Item 'HKLM:\Software\Microsoft\MSRDC'
$registryKey2 = Get-Item 'HKLM:\Software\Microsoft\MSRDC\Policies'
$registryKey3 = Get-ItemProperty 'HKLM:\Software\Microsoft\MSRDC\Policies\AutomaticUpdates'

if(-not $registryKey1){
New-Item -Path 'HKLM:\Software\Microsoft\' -name 'MSRDC'
Write-Host 'MSRDC registry key created.' -BackgroundColor Green -ForegroundColor Black
}
else {
Write-Host 'MSRDC registry key already exists' -BackgroundColor Green -ForegroundColor Black
}
if(-not $registryKey2){
New-Item -Path 'HKLM:\Software\Microsoft\MSRDC\' -name 'Policies'
Write-Host 'Policies registry key created.' -BackgroundColor Green -ForegroundColor Black
}
else {
Write-Host 'Policies registry already exists' -BackgroundColor Green -ForegroundColor Black
}
if(-not $registryKey3){
New-ItemProperty -Path 'HKLM:\Software\Microsoft\MSRDC\Policies\' -name 'AutomaticUpdates' -Value '2'
Write-Host 'AutomaticUpdates registry key created and value set to '2'.' -BackgroundColor Green -ForegroundColor Black
}
else {
Write-Host 'AutomaticUpdates registry already exists' -BackgroundColor Green -ForegroundColor Black
}
Remove-Item -Path 'C:\temp\RemoteDesktop.msi' -Force
Remove-Item -Path 'C:\temp\RemoteDesktop.bat' -Force
Read-Host -Prompt "Press Enter to exit"
