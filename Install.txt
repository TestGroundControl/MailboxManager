## Install File for the Mailbox Manager
## Download Files from Github

## Created by Aaron Haydon 2024
$downloadlocation = "C:\Temp\MailboxManager"
$fileexisits = Test-Path $downloadlocation
if ($fileexisits -eq $False) {
    Invoke-WebRequest -URI "https://github.com/TestGroundControl/MailboxManager/archive/refs/heads/main.zip" -OutFile "C:\Temp\MailboxManager.zip"
    mkdir "C:\Temp\MailboxManager"
    Expand-Archive -Path "C:\Temp\MailboxManager.zip" -DestinationPath $downloadlocation

}
else {
    Write-Host "File Already Exists"
}
if (-not $env:PSModulePath) {
        throw "The environment variable 'PSModulePath' is not set."
    }

    $PSModulePath = $env:PSModulePath -split ";" | Select-Object -First 1
    $path = "$($PSModulePath)\MailboxFunctions"
try {
if( (Test-Path($path)) -eq $false) {
    
        New-Item -Path $path -ItemType Directory -Force
    
        Copy-Item -Path 'C:\Temp\MailboxManager\MailboxManager-main\MailboxFunctions\*' -Destination $path -Recurse -Force
        }
else {
    Write-Host "Module Already Exists"
    Remove-Item -Path $path -Force -Recurse
    
    New-Item -Path $path -ItemType Directory -Force
    
    Copy-Item -Path 'C:\Temp\MailboxManager\MailboxManager-main\MailboxFunctions\*' -Destination $path -Recurse -Force
       
}
}
catch {
    Write-Host "Error: The path '$downloadlocation' does not exist." -ForegroundColor Red
    }

#Remove-Item -Path "C:\Temp\MailboxManager.zip" -Force
#Remove-Item -Path "C:\Temp\MailboxManager" -Force -Recurse



