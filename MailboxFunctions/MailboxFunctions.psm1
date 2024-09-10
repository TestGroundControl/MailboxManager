## Function File for Mailbox Manager
## Contains the following Scripts
## - Connect-Modules
## - Add-DataGrid
## - Clear-UserDetails
## - Clear-MailboxDetails
## - Get-Accounts
## - Get-Details
## - Get-Access
## - Enable-AutoExpand
## - Enable-Archive
## - Remove-Permissions
## - DelegateIndivdual
## - Get-MailboxLicenseDetails

## Created by Aaron Haydon 2024


<#
    .SYNOPSIS
    Connect to Exchange Online and Microsoft Graph and install the module if it does not exist

    .DESCRIPTION
    This function will check if the Exchange Online management and Graph module is installed and available. If it is not, it will install the module and then connect to Exchange Online.

    .PARAMETER None
    This cmdlet does not take any parameters.

    .EXAMPLE
    Connect-Modules

    .NOTES
    The module is installed in the user scope, so it will not require administrative access to install.
#>
function Connect-Modules {
    [CmdletBinding()]
    Param()

    # Check if the Exchange Online module is available
    $ExchangeOnlineModule = Get-Module ExchangeOnlineManagement -ListAvailable
    if (-not $ExchangeOnlineModule) {
        Write-Verbose "Important: Exchange Online module is unavailable. It is mandatory to have this module installed in the system to run the script successfully."
        Write-Verbose "Installing Exchange Online module..."
        Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
        Write-Verbose "Exchange Online Module is installed in the machine successfully" -ForegroundColor Magenta
    }

    # Connect to Exchange Online
    Write-Verbose "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$false | Out-Null
    Write-Verbose "Connected to Exchange Online successfully."

    
    $GraphModule = Get-Module Microsoft.Graph.Users -ListAvailable
    if (-not $GraphModule) { 
        Write-Warning "Important: Microsoft Graph module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
        Write-Verbose "Installing Microsoft Graph module..."
        Install-Module -Name Microsoft.Graph.Users -Scope CurrentUser
        Write-Verbose "Microsoft Graph Module is installed in the machine successfully" -ForegroundColor Magenta
        Connect-MgGraph -NoWelcome | Out-Null 
        Write-Verbose "Connected to Microsoft Graph successfully."
    } 
    else { 
        Write-Verbose "Connecting to Microsoft Graph..."
        Connect-MgGraph -NoWelcome | Out-Null 
            
    }
}

function DelegateIndivdual {
    $lblprogress.Visible = $true
    $progressBar.Visible = $true
    $lblprogress.Text = "Delegating Permissions..."
    $user = Get-MgUser -Filter "displayname eq '$($comboUsers.SelectedItem)'"
    $progressBar.Value = 0
    $progressBar.Visible = $true
    $lblProgress.Visible = $true
	
    try {
        if ($chkReqFullAccess.Checked -eq $true) {
    
            If ($chkReqAutoMap.Checked -eq $true) {
                Add-MailboxPermission -Identity $mailbox.Identity -User $user.UserPrincipalName -AccessRights FullAccess -InheritanceType All -AutoMapping $true -Confirm:$false
            }
            else {
                Add-MailboxPermission -Identity $mailbox.Identity -User $user.UserPrincipalName -AccessRights FullAccess -InheritanceType All -AutoMapping $false -Confirm:$false
            }
	
            $progressBar.Value = $progressBar.Value + 30 
        }
        if ($chkReqSendAs.Checked -eq $true) {
            Add-RecipientPermission -Identity $mailbox.Identity -AccessRights SendAs -Trustee $user.UserPrincipalName -Confirm:$false
            $progressBar.Value = $progressBar.Value + 30 
        }
    
    }
    catch [System.Exception] {
        Write-Error "Error processing Mailbox: $($_.Exception.Message)"
        $lblProgress.Text = "Error processing Mailbox: $($_.Exception.Message)"
        $progressBar.ForeColor = [System.Drawing.Color]::Red
        $progressBar.Value = 100
    }
    finally {
        $Permissions = Get-MailboxFolderPermission -Identity $mailbox | Where-Object { $_.User -Like $user.UserPrincipalName }     
        If ($Null -ne $Permissions) {
            $progressBar.Value = 100
            $lblProgress.Text = "Delegation Failed"
            $progressBar.ForeColor = [System.Drawing.Color]::Red
        }
        Else {
            $progressBar.Value = 100
            $lblProgress.Text = "Delegation Successful"
            $progressBar.ForeColor = [System.Drawing.Color]::Green
        }
    }
        

}

Function Get-Details {
    [CmdletBinding()]
    Param(
    )
    # Check if the mailbox has a Archive
    $lblProgress.Text = "Getting Report for $($mailbox.DisplayName)"
    $progressBar.Value = 10
    # If the Mailbox has an Archive Get the Stats

    $lblRecipientType.Text = $mailbox.RecipientTypeDetails
    $SharedMbxReport = Get-MailboxLicenseDetails -mailbox $mailbox
    
    $progressBar.Value = 30
    Write-Host $SharedMbxReport
    try {
        $status = Get-ExoMailbox $mailbox.Identity -Property Archive

    }
    catch {

        Write-Verbose "No Archive Enabled"
    }
    if ("Active" -eq $status.ArchiveStatus) {
        $chkArchive.Checked = $true
        $btnEnableArchive.Enabled = $false
        $MbxStats = Get-ExoMailboxStatistics $mailbox.UserPrincipalName | Select-Object ItemCount, TotalItemSize, TotalDeletedItemSize
        $progressBar.Value = 50
        $lblMbxStats.Text = "Item Count: $($MbxStats.ItemCount) -- Total Item Size: $($MbxStats.TotalItemSize)`nTotal Deleted Item Size: $($MbxStats.TotalDeletedItemSize)`nLicense: Plan 1: $($SharedMbxReport.ExoPlan1License.ToString()) // Plan 2: $($SharedMbxReport.ExoPlan2License.ToString())`nMailbox Over Size: $($SharedMbxReport.MailboxOverSize.ToString())`nLicense Status: $($SharedMbxReport.LicenseStatus)`n"
        Write-Verbose $lblMbxStats.Text
        $progressBar.Value = 50
        $ArchiveStats = Get-ExoMailboxStatistics $mailbox.UserPrincipalName -Archive | Select-Object ItemCount, TotalItemSize, TotalDeletedItemSize, ArchiveQuota
        $AutoExpanding = Get-Mailbox $mailbox.Identity | Select-Object AutoExpandingArchiveEnabled
        $progressBar.Value = 60
        $lblArchiveStats.Text = "Needs License: $($SharedMbxReport.NeedsLicense)`nItem Count: $($ArchiveStats.ItemCount) -- Total Item Size: $($ArchiveStats.TotalItemSize)`nTotal Deleted Item Size: $($ArchiveStats.TotalDeletedItemSize)`nAutoExpanding Archive: $($AutoExpanding.AutoExpandingArchiveEnabled.ToString()) `nArchive Quota: $($ArchiveStats.ArchiveQuota)"
        Write-Verbose $lblArchiveStats.Text
        $progressBar.Value = 90
        
    }
    else {
        $progressBar.Value = 60
        $chkArchive.Checked = $false
        $btnEnableArchive.Enabled = $true
        $btnAutoExpand.Enabled = $true
        $MbxStats = Get-ExoMailboxStatistics $mailbox.Identity | Select-Object ItemCount, TotalItemSize, TotalDeletedItemSize
        $progressBar.Value = 70
        $lblMbxStats.Text = "Item Count: $($MbxStats.ItemCount) -- Total Item Size: $($MbxStats.TotalItemSize)`nTotal Deleted Item Size: $($MbxStats.TotalDeletedItemSize)`nLicense: Plan 1: $($SharedMbxReport.ExoPlan1License.ToString()) // Plan 2: $($ExoPlan2License.ToString())`nMailbox Over Size: $($SharedMbxReport.MailboxOverSize.ToString())`nLicense Status: $($SharedMbxReport.LicenseStatus)"
        Write-Verbose $lblMbxStats.Text
        $progressBar.Value = 90

    }
    if ($AutoExpanding.AutoExpandingArchiveEnabled -eq $true) {
     
        $btnAutoExpand.Enabled = $false

    }

    $lblProgress.Text = "Statitics Retrieved"
    $progressBar.Value = 100
    
        
}

Function Get-Access {


    Add-DataGrid
    $Access = Get-ExoMailboxPermission -Identity $mailbox.Identity -User $User.UserPrincipalName
    $sendAs = Get-EXORecipientPermission -Identity $mailbox.Identity -Trustee $user.UserPrincipalName
    if ($Access.AccessRights -eq 'FullAccess') {
        $chkCurrentFull.Checked = $true
    }
    else {
        $chkCurrentFull.Checked = $false
    }
    if ($sendAs.AccessRights -eq 'SendAs') {
        $chkCurrentSendAs.Checked = $true
    }
    else {
        $chkCurrentSendAs.Checked = $false
    }
    
    
}

Function Add-DataGrid {

    # Step 2: Create a DataTable and Define Columns
    $dataTable = New-Object System.Data.DataTable
    $dataTable.Columns.Add("User", [System.String])
    $dataTable.Columns.Add("AccessRights", [System.String])

    # Step 3: Populate the DataTable with data
    $mailboxPermissions = Get-ExoMailboxPermission -Identity $mailbox.Identity | Select-Object User, AccessRights

    foreach ($permission in $mailboxPermissions) {
        $row = $dataTable.NewRow()
        $row["User"] = $permission.User
        $row["AccessRights"] = $permission.AccessRights -join ", "
        $dataTable.Rows.Add($row)
    }

    # Step 4: Set the DataTable as the DataSource
    $dataGridPermissions.DataSource = $dataTable
}

Function Clear-MailboxDetails {
    $lblArchiveStats.Text = $null
    $lblMbxStats.Text = $null
    $lblRecipientType.Text = $null
    $chkArchive.Checked = $false
    $dataGridPermissions.Rows.Clear()

}

Function Clear-UserDetails {
    $chkCurrentFull.Checked = $false
    $chkCurrentSendAs.Checked = $false
}

Function Get-Data {
    try {
        
    
    $lblProgress.Visible = $true
    $progressBar.Value = 15
    $lblProgress.Text = "Loading mailboxes: $($progressBar.Value)%"
    $mailboxes = Get-ExoMailbox -ResultSize Unlimited
    $progressBar.Value = 30

    $totalMailboxes = $mailboxes.Count
    $currentMailbox = 0
    $mailboxes | ForEach-Object {
    
        $comboMailboxes.Items.Add($_.DisplayName)
        $currentMailbox++
        $progressBar.Value = [math]::Round(($currentMailbox / $totalMailboxes) * 50)
        $lblProgress.Text = "Loading mailboxes: $($progressBar.Value)%"
        
    }
    # Get the users and update the progress bar
    $progressBar.Value = 50
    $lblProgress.Text = "Loading users: $($progressBar.Value)%"
    $users = Get-MgUser -All
    $progressBar.Value = 75
    $totalUsers = $users.Count
    $currentUser = 0
    $users | ForEach-Object {
        $comboUsers.Items.Add($_.DisplayName)
        $currentUser++
        $progressBar.Value = 50 + [math]::Round(($currentUser / $totalUsers) * 50)
        $lblProgress.Text = "Loading users: $($progressBar.Value)%"
    
    }
    $progressBar.Visible = $false
    $lblProgress.Visible = $false
}
catch {
    $lblProgress.Visible = $true
    $lblProgress.ForeColor = "Red"
    $lblProgress.Text = "Error Getting Data $($_.Exception)"
}
}

Function Enable-AutoExpand {
    ##Enable Auto-Expand Archive for selected Mailbox
    try {
        $mailbox | Set-EXOMailbox -AutoExpandingArchive $true
        $lblProgress.ForeColor = "Black"
        $lblProgress.Text = "Auto Expanding Archive Enabled on $($mailbox.DisplayName)"
    }
    catch {
        $lblProgress.ForeColor = "Red"
        $lblProgress.Text = "Error Setting Auto Expand Archive $($_.Exception)"
    }
    
    
}

function Enable-Archive {
    try {
        Enable-Mailbox -Identity $mailbox.Identity -Archive
        $lblProgress.ForeColor = "Black"
        $lblProgress.Text = "Archive Enabled on $($mailbox.DisplayName)"
    }
    catch {
        $lblProgress.ForeColor = "Red"
        $lblProgress.Text = "Error Enabling Archive $($_.Exception)"
    }
    
}

function Get-Accounts {
    $Global:mailbox = Get-EXOMailbox -Identity $comboMailboxes.SelectedItem
    if ($comboUsers.SelectedItem -ne $null) {
        $Global:user = Get-MgUser -Filter "displayname eq '$($comboUsers.SelectedItem)'"
    }
    

}

function Remove-Permissions {
        [CmdletBinding()]
        param ( 
        )
        Get-Accounts
        try {
            if($chkReqFullAccess.Checked -eq $true -and $chkReqFullAccess.Checked -eq $true) {
                Remove-MailboxPermission -Identity $mailbox.Identity -User $user.UserPrincipalName -Confirm:$false
                $lblProgress.Text = "Permissions Removed from $($mailbox.DisplayName) for FullAccess for $($user.DisplayName)"
                $progressBar.Value = 50

                Remove-MailboxPermission -Identity $mailbox.Identity -User $user.UserPrincipalName -AccessRights SendAs -Confirm:$false
                $lblProgress.Text = "Permissions Removed from $($mailbox.DisplayName) for SendAs for $($user.DisplayName)"
                $progressBar.Value = 100
            }
            elseif ($chkReqFullAccess.Checked -eq $true -and $chkReqSendAs.Checked -eq $false) {
                Remove-MailboxPermission -Identity $mailbox.Identity -User $user.UserPrincipalName -Confirm:$false
                $lblProgress.Text = "Permissions Removed from $($mailbox.DisplayName) for FullAccess for $($user.DisplayName)"
                $progressBar.Value = 100
            }
            elseif ($chkReqFullAccess.Checked -eq $false -and $chkReqSendAs.Checked -eq $true) {
                Remove-MailboxPermission -Identity $mailbox.Identity -User $user.UserPrincipalName -AccessRights SendAs -Confirm:$false
                $lblProgress.Text = "Permissions Removed from $($mailbox.DisplayName) for SendAs for $($user.DisplayName)"

            }
            else {
                lblProgress.Text = "Please Select a Permission to Remove"
                lblProgress.ForeColor = "Orange"
            }
        }
        catch {
            $lblProgress.Text = "Error removing Permissions $($_.Exception)"
            $progressBar.ForeColor = "Red"
            $progressBar.Value = 100
        }
        
            
        }
        
# ReportSharedMailboxLicenses.PS1

Function Get-MailboxLicenseDetails {
    [CmdletBinding()]
    param($mailbox)

    # Define some variables
    $SharedMbxReport = [System.Collections.Generic.List[Object]]::new()
    $ExoPlan1 = "9aaf7827-d63c-4b61-89c3-182f06f82e5c"
    $ExoArchiveAddOn = "ee02fd1b-340e-4a4b-b355-4a514e4c8943"
    $ExoPlan2 = "efb87545-963c-4e0d-99df-69c6916d9eb0" # See https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
    $mailboxLimit = 50GB

    Write-Verbose "Creating License report"
    $NeedsLicense = $False; $ArchiveStatus = $Null; $ExoArchiveLicense = $False; $ExoPlan2License = $False; $LicenseStatus = "OK"; $ArchiveStats = $Null
    $mailboxOverSize = $False; $ExoPlan1License = $False; $ArchiveMbxSize = $Null
    
    $MbxStats = Get-ExomailboxStatistics -UserPrincipalName $mailbox.UserPrincipalName
    $MbxSize = [math]::Round(($MbxStats.TotalItemSize.Value.toBytes() / 1GB),5)
    If ($mailbox.ArchiveStatus -ne "Active") { #Mailbox has an archive
        $ArchiveStats = Get-ExoMailboxStatistics -Archive -UserPrincipalName $mailbox.UserPrincipalName
        IF ($ArchiveStats) {       
            $ArchiveMbxSize = [math]::Round(($ArchiveStats.TotalItemSize.Value.toBytes() / 1GB),5)}
    }
    $Licenses = Get-MgUserLicenseDetail -UserId $mailbox.UserPrincipalName | Select-Object -ExpandProperty ServicePlans | Where-Object {$_.ProvisioningStatus -eq "Success"} | Sort ServicePlanId -Unique
    If ($Licenses) { # The mailbox has some licenses
        If ($ExoArchiveAddOn -in $Licenses.ServicePlanId) { $ExoArchiveLicense = $True }
        If ($ExoPlan2 -in $Licenses.ServicePlanId) { $ExoPlan2License = $True }
        If ($ExoPlan1 -in $Licenses.ServicePlanId) { $ExoPlan1License = $True }
 
    }

    # Mailbox has an archive and it doesn't have an Exchange Online Plan 2 license, unless it has Exchange Online Plan 1 and the archive add-on
    If ($mailbox.ArchiveStatus -eq "Active") {
        If ($ExoPlan2License -eq $False) { $NeedsLicense = $True }
        If ($ExoPlan1License -eq $True -and $ExoArchiveLicense -eq $True) { $NeedsLicense = $False }
    }
    # Mailbox is on litigation hold and it doesn't have an Exchange Online Plan 2 license
    If ($mailbox.LitigationHoldEnabled -eq $True -and $ExoPlan2License -eq $False)  { $NeedsLicense = $True }
    # Mailbox is over the 50GB limit for unlicensed shared mailboxes
    If ($MbxStats.TotalItemSize.value -gt $mailboxLimit) { # Exceeds mailbox size for unlicensed shared mailboxes
        $mailboxOverSize = $True
        $NeedsLicense = $True}
        Else { 
            $mailboxOverSize = $False 
            $NeedsLicense = $False
        
        }

    $SharedMbxReport = [PSCustomObject]@{
        ExoPlan1License = $ExoPlan1License
        ExoPlan2License = $ExoPlan2License
        ExoArchiveLicense = $ExoArchiveLicense
        NeedsLicense = $NeedsLicense
        LicenseStatus = $LicenseStatus
        MailboxSize = $MbxSize
        MailboxItems = $MbxStats.ItemCount
        MailboxOverSize = $mailboxOverSize
        Archive = $mailbox.ArchiveStatus
        ArchiveSize = $ArchiveMbxSize
        ArchiveItems = $ArchiveStats.ItemCount
    }
    
        
    Return $SharedMbxReport
    Write-Verbose $SharedMbxReport
}

