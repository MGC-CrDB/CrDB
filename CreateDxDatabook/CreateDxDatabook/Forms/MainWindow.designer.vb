<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainWindow
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer
    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainWindow))
        Me.StatusbarMain = New System.Windows.Forms.StatusStrip()
        Me.MsgArea = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ProgressBar = New System.Windows.Forms.ToolStripProgressBar()
        Me.EntryLmcFile = New System.Windows.Forms.TextBox()
        Me.MainMenubar = New System.Windows.Forms.MenuStrip()
        Me.MenuFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.ItemSelLmcFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ItemEditConfig = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ItemChkDao360 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.ItemExit = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuHelp = New System.Windows.Forms.ToolStripMenuItem()
        Me.ItemAbout = New System.Windows.Forms.ToolStripMenuItem()
        Me.ItemCmdLineArgs = New System.Windows.Forms.ToolStripMenuItem()
        Me.ItemDAO360Error = New System.Windows.Forms.ToolStripMenuItem()
        Me.ItemOtherMdbDir = New System.Windows.Forms.ToolStripMenuItem()
        Me.ItemPropSideFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.ItemSysInfo = New System.Windows.Forms.ToolStripMenuItem()
        Me.BtnSelLmcFile = New System.Windows.Forms.Button()
        Me.BtnCrDxDbook = New System.Windows.Forms.Button()
        Me.EntryPrpFile = New System.Windows.Forms.TextBox()
        Me.BtnSelPsfFile = New System.Windows.Forms.Button()
        Me.ChkCrtDbcFile = New System.Windows.Forms.CheckBox()
        Me.BtnLibManager = New System.Windows.Forms.Button()
        Me.EntryMdbDir = New System.Windows.Forms.TextBox()
        Me.BtnSelMdbDir = New System.Windows.Forms.Button()
        Me.LblLmcFile = New System.Windows.Forms.Label()
        Me.LblMdbDir = New System.Windows.Forms.Label()
        Me.LblPrpFile = New System.Windows.Forms.Label()
        Me.TranscriptList = New CreateDxDatabook.Transcript()
        Me.StatusbarMain.SuspendLayout()
        Me.MainMenubar.SuspendLayout()
        Me.SuspendLayout()
        '
        'StatusbarMain
        '
        Me.StatusbarMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MsgArea, Me.ProgressBar})
        Me.StatusbarMain.Location = New System.Drawing.Point(0, 639)
        Me.StatusbarMain.Name = "StatusbarMain"
        Me.StatusbarMain.Size = New System.Drawing.Size(784, 22)
        Me.StatusbarMain.TabIndex = 4
        Me.StatusbarMain.Text = "StatusStrip1"
        '
        'MsgArea
        '
        Me.MsgArea.AutoSize = False
        Me.MsgArea.Name = "MsgArea"
        Me.MsgArea.Size = New System.Drawing.Size(490, 17)
        Me.MsgArea.Text = "MsgArea"
        Me.MsgArea.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ProgressBar
        '
        Me.ProgressBar.AutoSize = False
        Me.ProgressBar.Name = "ProgressBar"
        Me.ProgressBar.Size = New System.Drawing.Size(250, 16)
        Me.ProgressBar.Step = 1
        Me.ProgressBar.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.ProgressBar.Visible = False
        '
        'EntryLmcFile
        '
        Me.EntryLmcFile.BackColor = System.Drawing.SystemColors.Window
        Me.EntryLmcFile.Enabled = False
        Me.EntryLmcFile.Location = New System.Drawing.Point(97, 35)
        Me.EntryLmcFile.Name = "EntryLmcFile"
        Me.EntryLmcFile.Size = New System.Drawing.Size(573, 20)
        Me.EntryLmcFile.TabIndex = 6
        Me.EntryLmcFile.WordWrap = False
        '
        'MainMenubar
        '
        Me.MainMenubar.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MenuFile, Me.MenuHelp})
        Me.MainMenubar.Location = New System.Drawing.Point(0, 0)
        Me.MainMenubar.Name = "MainMenubar"
        Me.MainMenubar.Size = New System.Drawing.Size(784, 24)
        Me.MainMenubar.TabIndex = 7
        Me.MainMenubar.Text = "MenuStrip1"
        '
        'MenuFile
        '
        Me.MenuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ItemSelLmcFile, Me.ToolStripSeparator2, Me.ItemEditConfig, Me.ToolStripSeparator1, Me.ItemChkDao360, Me.ToolStripSeparator3, Me.ItemExit})
        Me.MenuFile.Name = "MenuFile"
        Me.MenuFile.Size = New System.Drawing.Size(37, 20)
        Me.MenuFile.Text = "File"
        '
        'ItemSelLmcFile
        '
        Me.ItemSelLmcFile.Name = "ItemSelLmcFile"
        Me.ItemSelLmcFile.Size = New System.Drawing.Size(183, 22)
        Me.ItemSelLmcFile.Text = "Select LMC File ..."
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(180, 6)
        '
        'ItemEditConfig
        '
        Me.ItemEditConfig.Name = "ItemEditConfig"
        Me.ItemEditConfig.Size = New System.Drawing.Size(183, 22)
        Me.ItemEditConfig.Text = "Edit Configuration ..."
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(180, 6)
        '
        'ItemChkDao360
        '
        Me.ItemChkDao360.Name = "ItemChkDao360"
        Me.ItemChkDao360.Size = New System.Drawing.Size(183, 22)
        Me.ItemChkDao360.Text = "Check DAO360.dll"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(180, 6)
        '
        'ItemExit
        '
        Me.ItemExit.Name = "ItemExit"
        Me.ItemExit.Size = New System.Drawing.Size(183, 22)
        Me.ItemExit.Text = "Exit"
        '
        'MenuHelp
        '
        Me.MenuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ItemAbout, Me.ItemCmdLineArgs, Me.ItemDAO360Error, Me.ItemOtherMdbDir, Me.ItemPropSideFile, Me.ItemSysInfo})
        Me.MenuHelp.Name = "MenuHelp"
        Me.MenuHelp.Size = New System.Drawing.Size(44, 20)
        Me.MenuHelp.Text = "Help"
        '
        'ItemAbout
        '
        Me.ItemAbout.Name = "ItemAbout"
        Me.ItemAbout.Size = New System.Drawing.Size(180, 22)
        Me.ItemAbout.Text = "About"
        '
        'ItemCmdLineArgs
        '
        Me.ItemCmdLineArgs.Name = "ItemCmdLineArgs"
        Me.ItemCmdLineArgs.Size = New System.Drawing.Size(180, 22)
        Me.ItemCmdLineArgs.Text = "CommandLine Args"
        '
        'ItemDAO360Error
        '
        Me.ItemDAO360Error.Name = "ItemDAO360Error"
        Me.ItemDAO360Error.Size = New System.Drawing.Size(180, 22)
        Me.ItemDAO360Error.Text = "DAO360 Error"
        '
        'ItemOtherMdbDir
        '
        Me.ItemOtherMdbDir.Name = "ItemOtherMdbDir"
        Me.ItemOtherMdbDir.Size = New System.Drawing.Size(180, 22)
        Me.ItemOtherMdbDir.Text = "MDB Directory"
        '
        'ItemPropSideFile
        '
        Me.ItemPropSideFile.Name = "ItemPropSideFile"
        Me.ItemPropSideFile.Size = New System.Drawing.Size(180, 22)
        Me.ItemPropSideFile.Text = "Property Side File"
        '
        'ItemSysInfo
        '
        Me.ItemSysInfo.Name = "ItemSysInfo"
        Me.ItemSysInfo.Size = New System.Drawing.Size(180, 22)
        Me.ItemSysInfo.Text = "System Info"
        '
        'BtnSelLmcFile
        '
        Me.BtnSelLmcFile.Location = New System.Drawing.Point(686, 34)
        Me.BtnSelLmcFile.Name = "BtnSelLmcFile"
        Me.BtnSelLmcFile.Size = New System.Drawing.Size(95, 21)
        Me.BtnSelLmcFile.TabIndex = 8
        Me.BtnSelLmcFile.Text = "Select LMC File"
        Me.BtnSelLmcFile.UseVisualStyleBackColor = True
        '
        'BtnCrDxDbook
        '
        Me.BtnCrDxDbook.Enabled = False
        Me.BtnCrDxDbook.Location = New System.Drawing.Point(538, 126)
        Me.BtnCrDxDbook.Name = "BtnCrDxDbook"
        Me.BtnCrDxDbook.Size = New System.Drawing.Size(132, 21)
        Me.BtnCrDxDbook.TabIndex = 9
        Me.BtnCrDxDbook.Text = "Create Database"
        Me.BtnCrDxDbook.UseVisualStyleBackColor = True
        '
        'EntryPrpFile
        '
        Me.EntryPrpFile.BackColor = System.Drawing.SystemColors.Window
        Me.EntryPrpFile.Enabled = False
        Me.EntryPrpFile.Location = New System.Drawing.Point(97, 95)
        Me.EntryPrpFile.Name = "EntryPrpFile"
        Me.EntryPrpFile.Size = New System.Drawing.Size(573, 20)
        Me.EntryPrpFile.TabIndex = 12
        Me.EntryPrpFile.Visible = False
        Me.EntryPrpFile.WordWrap = False
        '
        'BtnSelPsfFile
        '
        Me.BtnSelPsfFile.Enabled = False
        Me.BtnSelPsfFile.Location = New System.Drawing.Point(686, 94)
        Me.BtnSelPsfFile.Name = "BtnSelPsfFile"
        Me.BtnSelPsfFile.Size = New System.Drawing.Size(95, 21)
        Me.BtnSelPsfFile.TabIndex = 13
        Me.BtnSelPsfFile.Text = "Select Prop File"
        Me.BtnSelPsfFile.UseVisualStyleBackColor = True
        Me.BtnSelPsfFile.Visible = False
        '
        'ChkCrtDbcFile
        '
        Me.ChkCrtDbcFile.AutoSize = True
        Me.ChkCrtDbcFile.Location = New System.Drawing.Point(391, 129)
        Me.ChkCrtDbcFile.Name = "ChkCrtDbcFile"
        Me.ChkCrtDbcFile.Size = New System.Drawing.Size(141, 17)
        Me.ChkCrtDbcFile.TabIndex = 15
        Me.ChkCrtDbcFile.Text = "Create/Update DBC File"
        Me.ChkCrtDbcFile.UseVisualStyleBackColor = True
        '
        'BtnLibManager
        '
        Me.BtnLibManager.Enabled = False
        Me.BtnLibManager.Location = New System.Drawing.Point(686, 126)
        Me.BtnLibManager.Name = "BtnLibManager"
        Me.BtnLibManager.Size = New System.Drawing.Size(95, 21)
        Me.BtnLibManager.TabIndex = 18
        Me.BtnLibManager.Text = "Library Manager"
        Me.BtnLibManager.UseVisualStyleBackColor = True
        '
        'EntryMdbDir
        '
        Me.EntryMdbDir.BackColor = System.Drawing.SystemColors.Window
        Me.EntryMdbDir.Enabled = False
        Me.EntryMdbDir.Location = New System.Drawing.Point(97, 65)
        Me.EntryMdbDir.Name = "EntryMdbDir"
        Me.EntryMdbDir.Size = New System.Drawing.Size(573, 20)
        Me.EntryMdbDir.TabIndex = 20
        Me.EntryMdbDir.Visible = False
        Me.EntryMdbDir.WordWrap = False
        '
        'BtnSelMdbDir
        '
        Me.BtnSelMdbDir.Location = New System.Drawing.Point(686, 64)
        Me.BtnSelMdbDir.Name = "BtnSelMdbDir"
        Me.BtnSelMdbDir.Size = New System.Drawing.Size(95, 21)
        Me.BtnSelMdbDir.TabIndex = 21
        Me.BtnSelMdbDir.Text = "Select MDB Dir"
        Me.BtnSelMdbDir.UseVisualStyleBackColor = True
        Me.BtnSelMdbDir.Visible = False
        '
        'LblLmcFile
        '
        Me.LblLmcFile.AutoSize = True
        Me.LblLmcFile.Location = New System.Drawing.Point(6, 38)
        Me.LblLmcFile.Name = "LblLmcFile"
        Me.LblLmcFile.Size = New System.Drawing.Size(85, 13)
        Me.LblLmcFile.TabIndex = 22
        Me.LblLmcFile.Text = "Library LMC File:"
        '
        'LblMdbDir
        '
        Me.LblMdbDir.AutoSize = True
        Me.LblMdbDir.Location = New System.Drawing.Point(25, 68)
        Me.LblMdbDir.Name = "LblMdbDir"
        Me.LblMdbDir.Size = New System.Drawing.Size(66, 13)
        Me.LblMdbDir.TabIndex = 23
        Me.LblMdbDir.Text = "MDB Folder:"
        Me.LblMdbDir.Visible = False
        '
        'LblPrpFile
        '
        Me.LblPrpFile.AutoSize = True
        Me.LblPrpFile.Location = New System.Drawing.Point(23, 98)
        Me.LblPrpFile.Name = "LblPrpFile"
        Me.LblPrpFile.Size = New System.Drawing.Size(68, 13)
        Me.LblPrpFile.TabIndex = 24
        Me.LblPrpFile.Text = "Property File:"
        '
        'TranscriptList
        '
        Me.TranscriptList.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.TranscriptList.Location = New System.Drawing.Point(0, 153)
        Me.TranscriptList.Name = "TranscriptList"
        Me.TranscriptList.Size = New System.Drawing.Size(784, 486)
        Me.TranscriptList.TabIndex = 19
        '
        'MainWindow
        '
        Me.ClientSize = New System.Drawing.Size(784, 661)
        Me.Controls.Add(Me.LblPrpFile)
        Me.Controls.Add(Me.LblMdbDir)
        Me.Controls.Add(Me.LblLmcFile)
        Me.Controls.Add(Me.BtnSelMdbDir)
        Me.Controls.Add(Me.EntryMdbDir)
        Me.Controls.Add(Me.TranscriptList)
        Me.Controls.Add(Me.BtnLibManager)
        Me.Controls.Add(Me.ChkCrtDbcFile)
        Me.Controls.Add(Me.BtnSelPsfFile)
        Me.Controls.Add(Me.EntryPrpFile)
        Me.Controls.Add(Me.BtnCrDxDbook)
        Me.Controls.Add(Me.BtnSelLmcFile)
        Me.Controls.Add(Me.EntryLmcFile)
        Me.Controls.Add(Me.StatusbarMain)
        Me.Controls.Add(Me.MainMenubar)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MainMenubar
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(800, 700)
        Me.MinimumSize = New System.Drawing.Size(800, 700)
        Me.Name = "MainWindow"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Create DxDatabook - MDB/SQLite"
        Me.StatusbarMain.ResumeLayout(False)
        Me.StatusbarMain.PerformLayout()
        Me.MainMenubar.ResumeLayout(False)
        Me.MainMenubar.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents StatusbarMain As System.Windows.Forms.StatusStrip
    Friend WithEvents MsgArea As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ProgressBar As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents EntryLmcFile As System.Windows.Forms.TextBox
    Friend WithEvents MainMenubar As System.Windows.Forms.MenuStrip
    Friend WithEvents MenuFile As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ItemExit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BtnSelLmcFile As System.Windows.Forms.Button
    Friend WithEvents BtnCrDxDbook As System.Windows.Forms.Button
    Friend WithEvents MenuHelp As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ItemAbout As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ItemSysInfo As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ItemSelLmcFile As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents EntryPrpFile As System.Windows.Forms.TextBox
    Friend WithEvents BtnSelPsfFile As System.Windows.Forms.Button
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ItemEditConfig As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ChkCrtDbcFile As System.Windows.Forms.CheckBox
    Friend WithEvents BtnLibManager As System.Windows.Forms.Button
    Friend WithEvents ItemChkDao360 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ItemCmdLineArgs As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TranscriptList As CreateDxDatabook.Transcript
    Friend WithEvents ItemPropSideFile As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EntryMdbDir As System.Windows.Forms.TextBox
    Friend WithEvents BtnSelMdbDir As System.Windows.Forms.Button
    Friend WithEvents ItemOtherMdbDir As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LblLmcFile As System.Windows.Forms.Label
    Friend WithEvents LblMdbDir As System.Windows.Forms.Label
    Friend WithEvents LblPrpFile As System.Windows.Forms.Label
    Friend WithEvents ItemDAO360Error As System.Windows.Forms.ToolStripMenuItem

End Class
