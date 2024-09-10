Start-Transcript -Path "c:\Temp\Logs\mailboxmanagerlog.txt"
$DebugPreference = "Continue"

# Ensure the required scripts are loaded
Import-Module MailboxFunctions
## Event Handlers

Write-Debug "Importing Event Handlers"
$btnEnableArchive_Click = {
    Enable-Archive
}

$btnAutoExpand_Click = {
    Enable-AutoExpand
}

$ShownHandler = {

    Get-Data
}

$comboMailbox_SelectedChange = {
    Clear-MailboxDetails
    Get-Accounts
    Get-Details
    if ($comboUsers.SelectedItem -ne $null) {
        Get-Accounts
        Get-Access
    }

 }

$comboUsers_SelectedChange = {
    Clear-UserDetails
    Get-Accounts
    Get-Access
    if ($comboMailbox.SelectedItem -ne $null) {
        
        Get-Details
    }
}

$btnRemove_Click = {
    $lblProgress.Visible = $true
    $progressBar.Visible = $true

    Remove-Permissions
}

$btnSubmit_Click = {
    $lblProgress.Visible = $true
    $progressBar.Visible = $true
    DelegateIndivdual
}
$HelpButtonClicked = {
    Start-Process "msedge" "https://groundcontrol.atlassian.net/wiki/spaces/GC/"
}

Write-Debug "Adding Windows Form"
Add-Type -AssemblyName System.Windows.Forms
. (Join-Path $PSScriptRoot 'Mailbox Script.designer.ps1')

Write-Debug "Importing .\mailbox script.resources.ps1"
if (Test-Path .\'mailbox script.resources.ps1') {
    . .\'mailbox script.resources.ps1'
} else {
    Write-Error "'mailbox script.resources.ps1' not found"
}

Write-Debug "Importing .\mailbox script.designer.ps1"

if (Test-Path .\'mailbox script.designer.ps1') {
    . .\'mailbox script.designer.ps1'
} else {
    Write-Error "'mailbox script.designer.ps1' not found"
}

Write-Debug "Connecting to Exchange Online"
Connect-Modules
Write-Debug "Show Display"
$MailboxManager.ShowDialog()
$MailboxManager.Dispose()
Stop-Transcript
