
$MailboxManager = New-Object -TypeName System.Windows.Forms.Form
[System.Windows.Forms.ComboBox]$comboUsers = $null
[System.Windows.Forms.ComboBox]$comboMailboxes = $null
[System.Windows.Forms.Label]$lblUser = $null
[System.Windows.Forms.Label]$lblMailbox = $null
[System.Windows.Forms.Label]$lblUsers = $null
[System.Windows.Forms.GroupBox]$grpCurrent = $null
[System.Windows.Forms.CheckBox]$chkCurrentFull = $null
[System.Windows.Forms.CheckBox]$chkCurrentSendAs = $null
[System.Windows.Forms.GroupBox]$grpRequest = $null
[System.Windows.Forms.CheckBox]$chkReqAutoMap = $null
[System.Windows.Forms.CheckBox]$chkReqFullAccess = $null
[System.Windows.Forms.CheckBox]$chkReqSendAs = $null
[System.Windows.Forms.Button]$btnSubmit = $null
[System.Windows.Forms.Button]$btnRemove = $null
[System.Windows.Forms.ProgressBar]$progressBar = $null
[System.Windows.Forms.Label]$lblProgress = $null
[System.Windows.Forms.Label]$Label1 = $null
[System.Windows.Forms.DataGridView]$dataGridPermissions = $null
[System.Windows.Forms.DataGridViewTextBoxColumn]$User = $null
[System.Windows.Forms.DataGridViewTextBoxColumn]$AccessRights = $null
[System.Windows.Forms.Label]$Label2 = $null
[System.Windows.Forms.Label]$lblRecipientType = $null
[System.Windows.Forms.Label]$Label3 = $null
[System.Windows.Forms.Label]$lblMbxStats = $null
[System.Windows.Forms.Label]$lblArchiveStats = $null
[System.Windows.Forms.Button]$btnEnableArchive = $null
[System.Windows.Forms.Button]$btnAutoExpand = $null
[System.Windows.Forms.CheckBox]$chkArchive = $null
[System.Windows.Forms.Label]$Label4 = $null
function InitializeComponent
{
$resources = . (Join-Path $PSScriptRoot 'mailbox script.resources.ps1')
$comboUsers = (New-Object -TypeName System.Windows.Forms.ComboBox)
$comboMailboxes = (New-Object -TypeName System.Windows.Forms.ComboBox)
$lblUser = (New-Object -TypeName System.Windows.Forms.Label)
$lblMailbox = (New-Object -TypeName System.Windows.Forms.Label)
$lblUsers = (New-Object -TypeName System.Windows.Forms.Label)
$grpCurrent = (New-Object -TypeName System.Windows.Forms.GroupBox)
$chkCurrentFull = (New-Object -TypeName System.Windows.Forms.CheckBox)
$chkCurrentSendAs = (New-Object -TypeName System.Windows.Forms.CheckBox)
$grpRequest = (New-Object -TypeName System.Windows.Forms.GroupBox)
$btnSubmit = (New-Object -TypeName System.Windows.Forms.Button)
$btnRemove = (New-Object -TypeName System.Windows.Forms.Button)
$progressBar = (New-Object -TypeName System.Windows.Forms.ProgressBar)
$lblProgress = (New-Object -TypeName System.Windows.Forms.Label)
$Label1 = (New-Object -TypeName System.Windows.Forms.Label)
$dataGridPermissions = (New-Object -TypeName System.Windows.Forms.DataGridView)
$User = (New-Object -TypeName System.Windows.Forms.DataGridViewTextBoxColumn)
$AccessRights = (New-Object -TypeName System.Windows.Forms.DataGridViewTextBoxColumn)
$Label2 = (New-Object -TypeName System.Windows.Forms.Label)
$lblRecipientType = (New-Object -TypeName System.Windows.Forms.Label)
$Label3 = (New-Object -TypeName System.Windows.Forms.Label)
$lblMbxStats = (New-Object -TypeName System.Windows.Forms.Label)
$lblArchiveStats = (New-Object -TypeName System.Windows.Forms.Label)
$btnEnableArchive = (New-Object -TypeName System.Windows.Forms.Button)
$btnAutoExpand = (New-Object -TypeName System.Windows.Forms.Button)
$chkReqSendAs = (New-Object -TypeName System.Windows.Forms.CheckBox)
$chkReqFullAccess = (New-Object -TypeName System.Windows.Forms.CheckBox)
$chkReqAutoMap = (New-Object -TypeName System.Windows.Forms.CheckBox)
$chkArchive = (New-Object -TypeName System.Windows.Forms.CheckBox)
$Label4 = (New-Object -TypeName System.Windows.Forms.Label)
$grpCurrent.SuspendLayout()
$grpRequest.SuspendLayout()
([System.ComponentModel.ISupportInitialize]$dataGridPermissions).BeginInit()
$MailboxManager.SuspendLayout()
#
#comboUsers
#
$comboUsers.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::Suggest
$comboUsers.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems
$comboUsers.FormattingEnabled = $true
$comboUsers.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]494,[System.Int32]154))
$comboUsers.Name = [System.String]'comboUsers'
$comboUsers.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]290,[System.Int32]21))
$comboUsers.Sorted = $true
$comboUsers.TabIndex = [System.Int32]0
$comboUsers.TabStop = $false
$comboUsers.Text = [System.String]'Select User...'
$comboUsers.add_SelectedValueChanged($comboUsers_SelectedChange)
#
#comboMailboxes
#
$comboMailboxes.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::Suggest
$comboMailboxes.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems
$comboMailboxes.FormattingEnabled = $true
$comboMailboxes.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]92,[System.Int32]19))
$comboMailboxes.Name = [System.String]'comboMailboxes'
$comboMailboxes.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]290,[System.Int32]21))
$comboMailboxes.Sorted = $true
$comboMailboxes.TabIndex = [System.Int32]0
$comboMailboxes.TabStop = $false
$comboMailboxes.Text = [System.String]'Select Mailbox...'
$comboMailboxes.add_SelectedValueChanged($comboMailbox_SelectedChange)
#
#lblUser
#
$lblUser.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]9.75,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$lblUser.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]8,[System.Int32]55))
$lblUser.Name = [System.String]'lblUser'
$lblUser.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]78,[System.Int32]21))
$lblUser.Text = [System.String]'User:'
$lblUser.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
#
#lblMailbox
#
$lblMailbox.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]9.75,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$lblMailbox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]8,[System.Int32]17))
$lblMailbox.Name = [System.String]'lblMailbox'
$lblMailbox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]78,[System.Int32]21))
$lblMailbox.TabIndex = [System.Int32]1
$lblMailbox.Text = [System.String]'Mailbox:'
$lblMailbox.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
#
#lblUsers
#
$lblUsers.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]9.75,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$lblUsers.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]423,[System.Int32]154))
$lblUsers.Name = [System.String]'lblUsers'
$lblUsers.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]67,[System.Int32]23))
$lblUsers.TabIndex = [System.Int32]2
$lblUsers.Text = [System.String]'User:'
$lblUsers.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
#
#grpCurrent
#
$grpCurrent.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$grpCurrent.BackColor = [System.Drawing.SystemColors]::GradientActiveCaption
$grpCurrent.Controls.Add($chkCurrentFull)
$grpCurrent.Controls.Add($chkCurrentSendAs)
$grpCurrent.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]9.75,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$grpCurrent.ForeColor = [System.Drawing.SystemColors]::ActiveCaptionText
$grpCurrent.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]423,[System.Int32]189))
$grpCurrent.Name = [System.String]'grpCurrent'
$grpCurrent.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]363,[System.Int32]58))
$grpCurrent.TabIndex = [System.Int32]7
$grpCurrent.TabStop = $false
$grpCurrent.Text = [System.String]'Current Access'
#
#chkCurrentFull
#
$chkCurrentFull.Enabled = $false
$chkCurrentFull.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]9.75,[System.Drawing.FontStyle]::Italic,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$chkCurrentFull.ForeColor = [System.Drawing.SystemColors]::ActiveCaptionText
$chkCurrentFull.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]197,[System.Int32]24))
$chkCurrentFull.Name = [System.String]'chkCurrentFull'
$chkCurrentFull.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]90,[System.Int32]24))
$chkCurrentFull.TabIndex = [System.Int32]1
$chkCurrentFull.Text = [System.String]'Full Access'
$chkCurrentFull.UseVisualStyleBackColor = $true
#
#chkCurrentSendAs
#
$chkCurrentSendAs.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]9.75,[System.Drawing.FontStyle]::Italic,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$chkCurrentSendAs.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]61,[System.Int32]24))
$chkCurrentSendAs.Name = [System.String]'chkCurrentSendAs'
$chkCurrentSendAs.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]104,[System.Int32]24))
$chkCurrentSendAs.TabIndex = [System.Int32]2
$chkCurrentSendAs.Text = [System.String]'Send As'
$chkCurrentSendAs.Enabled = $false
#
#grpRequest
#
$grpRequest.BackColor = [System.Drawing.SystemColors]::GradientActiveCaption
$grpRequest.Controls.Add($chkReqAutoMap)
$grpRequest.Controls.Add($chkReqFullAccess)
$grpRequest.Controls.Add($chkReqSendAs)
$grpRequest.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]9.75,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$grpRequest.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]423,[System.Int32]262))
$grpRequest.Name = [System.String]'grpRequest'
$grpRequest.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]363,[System.Int32]80))
$grpRequest.TabIndex = [System.Int32]25
$grpRequest.TabStop = $false
$grpRequest.Text = [System.String]'Requested Access'
#
#btnSubmit
#
$btnSubmit.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]566,[System.Int32]365))
$btnSubmit.Name = [System.String]'btnSubmit'
$btnSubmit.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]121,[System.Int32]35))
$btnSubmit.TabIndex = [System.Int32]24
$btnSubmit.Text = [System.String]'Submit'
$btnSubmit.add_Click($btnSubmit_Click)
#
#btnRemove
#
$btnRemove.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]566,[System.Int32]420))
$btnRemove.Name = [System.String]'btnRemove'
$btnRemove.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]121,[System.Int32]35))
$btnRemove.TabIndex = [System.Int32]26
$btnRemove.Text = [System.String]'Remove Selected'
$btnRemove.add_Click($btnRemove_Click)
#
#progressBar
#
$progressBar.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]174,[System.Int32]422))
$progressBar.Name = [System.String]'progressBar'
$progressBar.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]386,[System.Int32]23))
$progressBar.TabIndex = [System.Int32]23
$progressBar.ForeColor = [System.Drawing.Color]::LightGreen
#
#lblProgress
#
$lblProgress.BackColor = [System.Drawing.Color]::Transparent
$lblProgress.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]9.75,([System.Drawing.FontStyle][System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic),[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$lblProgress.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]199,[System.Int32]375))
$lblProgress.Name = [System.String]'lblProgress'
$lblProgress.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]361,[System.Int32]44))
$lblProgress.TabIndex = [System.Int32]11
$lblProgress.Text = [System.String]'Processing ... '
$lblProgress.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$lblProgress.Visible = $false
#
#Label1
#
$Label1.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]12,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$Label1.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]494,[System.Int32]15))
$Label1.Name = [System.String]'Label1'
$Label1.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]226,[System.Int32]23))
$Label1.TabIndex = [System.Int32]12
$Label1.Text = [System.String]'Mailbox Permissions'
$Label1.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
#
#dataGridPermissions
#
$dataGridPermissions.BackgroundColor = [System.Drawing.SystemColors]::InactiveCaption
$dataGridPermissions.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
$dataGridPermissions.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$dataGridPermissions.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]423,[System.Int32]47))
$dataGridPermissions.Name = [System.String]'dataGridPermissions'
$dataGridPermissions.ReadOnly = $true
$dataGridPermissions.RowTemplate.ReadOnly = $true
$dataGridPermissions.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$dataGridPermissions.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]361,[System.Int32]90))
$dataGridPermissions.TabIndex = [System.Int32]13
#
#User
# #
# $User.HeaderText = [System.String]'User'
# $User.Name = [System.String]'User'
# $User.Width = [System.Int32]200
# #
# #AccessRights
# #
# $AccessRights.HeaderText = [System.String]'Access Rights'
# $AccessRights.Name = [System.String]'AccessRights'
# #
#Label2
#
$Label2.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]9.75,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$Label2.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]22,[System.Int32]46))
$Label2.Name = [System.String]'Label2'
$Label2.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]64,[System.Int32]23))
$Label2.TabIndex = [System.Int32]14
$Label2.Text = [System.String]'Type:'
#
#lblRecipientType
#
$lblRecipientType.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]92,[System.Int32]48))
$lblRecipientType.Name = [System.String]'lblRecipientType'
$lblRecipientType.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]100,[System.Int32]23))
$lblRecipientType.TabIndex = [System.Int32]15
$lblRecipientType.Text = [System.String]'Shared Mailbox'
#
#Label3
#
$Label3.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]9.75,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$Label3.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]121,[System.Int32]82))
$Label3.Name = [System.String]'Label3'
$Label3.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]125,[System.Int32]23))
$Label3.TabIndex = [System.Int32]16
$Label3.Text = [System.String]'Mailbox Statistics'
#
#lblMbxStats
#
$lblMbxStats.BackColor = [System.Drawing.SystemColors]::ControlLightLight
$lblMbxStats.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]12,[System.Int32]105))
$lblMbxStats.Name = [System.String]'lblMbxStats'
$lblMbxStats.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]378,[System.Int32]107))
$lblMbxStats.TabIndex = [System.Int32]17
#
#lblArchiveStats
#
$lblArchiveStats.BackColor = [System.Drawing.SystemColors]::ControlLightLight
$lblArchiveStats.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]12,[System.Int32]252))
$lblArchiveStats.Name = [System.String]'lblArchiveStats'
$lblArchiveStats.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]378,[System.Int32]107))
$lblArchiveStats.TabIndex = [System.Int32]19
#
#btnEnableArchive
#
$btnEnableArchive.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]31,[System.Int32]365))
$btnEnableArchive.Name = [System.String]'btnEnableArchive'
$btnEnableArchive.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]121,[System.Int32]35))
$btnEnableArchive.TabIndex = [System.Int32]21
$btnEnableArchive.Text = [System.String]'Enable Archive'
$btnEnableArchive.UseVisualStyleBackColor = $true
$btnEnableArchive.add_Click($btnEnableArchive_Click)
#
#btnAutoExpand
#
$btnAutoExpand.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]31,[System.Int32]420))
$btnAutoExpand.Name = [System.String]'btnAutoExpand'
$btnAutoExpand.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]121,[System.Int32]35))
$btnAutoExpand.TabIndex = [System.Int32]22
$btnAutoExpand.Text = [System.String]'Enable Auto Expanding'
$btnAutoExpand.UseVisualStyleBackColor = $true
$btnAutoExpand.add_Click($btnAutoExpand_Click)
#
#chkReqSendAs
#
$chkReqSendAs.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]9.75,[System.Drawing.FontStyle]::Italic,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$chkReqSendAs.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]129,[System.Int32]31))
$chkReqSendAs.Name = [System.String]'chkReqSendAs'
$chkReqSendAs.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]104,[System.Int32]24))
$chkReqSendAs.TabIndex = [System.Int32]0
$chkReqSendAs.Text = [System.String]'Send As'
$chkReqSendAs.UseVisualStyleBackColor = $true
#
#chkReqFullAccess
#
$chkReqFullAccess.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]9.75,[System.Drawing.FontStyle]::Italic,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$chkReqFullAccess.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]19,[System.Int32]31))
$chkReqFullAccess.Name = [System.String]'chkReqFullAccess'
$chkReqFullAccess.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]104,[System.Int32]24))
$chkReqFullAccess.TabIndex = [System.Int32]1
$chkReqFullAccess.Text = [System.String]'Full Access'
$chkReqFullAccess.UseVisualStyleBackColor = $true
#
#chkReqAutoMap
#
$chkReqAutoMap.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]9.75,[System.Drawing.FontStyle]::Italic,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$chkReqAutoMap.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]227,[System.Int32]31))
$chkReqAutoMap.Name = [System.String]'chkReqAutoMap'
$chkReqAutoMap.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]104,[System.Int32]24))
$chkReqAutoMap.TabIndex = [System.Int32]2
$chkReqAutoMap.Text = [System.String]'AutoMapped'
$chkReqAutoMap.UseVisualStyleBackColor = $true
#
#chkEnableArchive
#
$chkArchive.CheckAlign = [System.Drawing.ContentAlignment]::MiddleRight
$chkArchive.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]233,[System.Int32]224))
$chkArchive.Name = [System.String]'chkArchive'
$chkArchive.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]154,[System.Int32]24))
$chkArchive.TabIndex = [System.Int32]26
$chkArchive.Text = [System.String]'Archive Enabled'
$chkArchive.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$chkArchive.UseVisualStyleBackColor = $true
#
#Label4
#
$Label4.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]9.75,([System.Drawing.FontStyle][System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic),[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$Label4.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]119,[System.Int32]227))
$Label4.Name = [System.String]'Label4'
$Label4.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]127,[System.Int32]23))
$Label4.TabIndex = [System.Int32]27
$Label4.Text = [System.String]'Archive Statistics'
#
#MailboxManager
#
$MailboxManager.BackColor = [System.Drawing.SystemColors]::Control
$MailboxManager.ClientSize = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]812,[System.Int32]459))
$MailboxManager.Controls.Add($Label4)
$MailboxManager.Controls.Add($chkArchive)
$MailboxManager.Controls.Add($btnAutoExpand)
$MailboxManager.Controls.Add($btnEnableArchive)
$MailboxManager.Controls.Add($lblArchiveStats)
$MailboxManager.Controls.Add($lblMbxStats)
$MailboxManager.Controls.Add($Label3)
$MailboxManager.Controls.Add($lblRecipientType)
$MailboxManager.Controls.Add($Label2)
$MailboxManager.Controls.Add($dataGridPermissions)
$MailboxManager.Controls.Add($Label1)
$MailboxManager.Controls.Add($lblProgress)
$MailboxManager.Controls.Add($progressBar)
$MailboxManager.Controls.Add($btnSubmit)
$MailboxManager.Controls.Add($btnRemove)
$MailboxManager.Controls.Add($grpRequest)
$MailboxManager.Controls.Add($grpCurrent)
$MailboxManager.Controls.Add($lblUsers)
$MailboxManager.Controls.Add($lblMailbox)
$MailboxManager.Controls.Add($comboMailboxes)
$MailboxManager.Controls.Add($comboUsers)
$MailboxManager.Icon = ([System.Drawing.Icon]$resources.'$this.Icon')
$MailboxManager.KeyPreview = $true
$MailboxManager.add_HelpButtonClicked($HelpButtonClicked)
$MailboxManager.HelpButton = $true
$MailboxManager.MaximizeBox = $false
$MailboxManager.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$MailboxManager.Text = [System.String]'Mailbox Manager'
$grpCurrent.ResumeLayout($false)
$grpRequest.ResumeLayout($false)
([System.ComponentModel.ISupportInitialize]$dataGridPermissions).EndInit()
$MailboxManager.ResumeLayout($false)
$MailboxManager.add_Shown($ShownHandler)
Add-Member -InputObject $MailboxManager -Name comboUsers -Value $comboUsers -MemberType NoteProperty
Add-Member -InputObject $MailboxManager -Name comboMailboxes -Value $comboMailboxes -MemberType NoteProperty
Add-Member -InputObject $MailboxManager -Name lblUser -Value $lblUser -MemberType NoteProperty
Add-Member -InputObject $MailboxManager -Name lblMailbox -Value $lblMailbox -MemberType NoteProperty
Add-Member -InputObject $MailboxManager -Name lblUsers -Value $lblUsers -MemberType NoteProperty
Add-Member -InputObject $MailboxManager -Name grpCurrent -Value $grpCurrent -MemberType NoteProperty
Add-Member -InputObject $MailboxManager -Name chkCurrentFull -Value $chkCurrentFull -MemberType NoteProperty
Add-Member -InputObject $MailboxManager -Name chkCurrentSendAs -Value $chkCurrentSendAs -MemberType NoteProperty
Add-Member -InputObject $MailboxManager -Name grpRequest -Value $grpRequest -MemberType NoteProperty
Add-Member -InputObject $MailboxManager -Name chkReqAutoMap -Value $chkReqAutoMap -MemberType NoteProperty
Add-Member -InputObject $MailboxManager -Name chkReqFullAccess -Value $chkReqFullAccess -MemberType NoteProperty
Add-Member -InputObject $MailboxManager -Name chkReqSendAs -Value $chkReqSendAs -MemberType NoteProperty
Add-Member -InputObject $MailboxManager -Name btnSubmit -Value $btnSubmit -MemberType NoteProperty
Add-Member -InputObject $MailboxManager -Name btnRemove -Value $btnRemove -MemberType NoteProperty
Add-Member -InputObject $MailboxManager -Name progressBar -Value $progressBar -MemberType NoteProperty
Add-Member -InputObject $MailboxManager -Name lblProgress -Value $lblProgress -MemberType NoteProperty
Add-Member -InputObject $MailboxManager -Name Label1 -Value $Label1 -MemberType NoteProperty
Add-Member -InputObject $MailboxManager -Name dataGridPermissions -Value $dataGridPermissions -MemberType NoteProperty
Add-Member -InputObject $MailboxManager -Name User -Value $User -MemberType NoteProperty
Add-Member -InputObject $MailboxManager -Name AccessRights -Value $AccessRights -MemberType NoteProperty
Add-Member -InputObject $MailboxManager -Name Label2 -Value $Label2 -MemberType NoteProperty
Add-Member -InputObject $MailboxManager -Name lblRecipientType -Value $lblRecipientType -MemberType NoteProperty
Add-Member -InputObject $MailboxManager -Name Label3 -Value $Label3 -MemberType NoteProperty
Add-Member -InputObject $MailboxManager -Name lblMbxStats -Value $lblMbxStats -MemberType NoteProperty
Add-Member -InputObject $MailboxManager -Name lblArchiveStats -Value $lblArchiveStats -MemberType NoteProperty
Add-Member -InputObject $MailboxManager -Name btnEnableArchive -Value $btnEnableArchive -MemberType NoteProperty
Add-Member -InputObject $MailboxManager -Name btnAutoExpand -Value $btnAutoExpand -MemberType NoteProperty
Add-Member -InputObject $MailboxManager -Name chkArchive -Value $chkArchive -MemberType NoteProperty
Add-Member -InputObject $MailboxManager -Name Label4 -Value $Label4 -MemberType NoteProperty
}
. InitializeComponent
