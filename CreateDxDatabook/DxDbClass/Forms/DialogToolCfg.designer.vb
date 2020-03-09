<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DialogToolCfg
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.BtnCancel = New System.Windows.Forms.Button()
        Me.BtnOK = New System.Windows.Forms.Button()
        Me.ChkEnaPrpImport = New System.Windows.Forms.CheckBox()
        Me.ChkUseOdbcAlias = New System.Windows.Forms.CheckBox()
        Me.CmbOdbcAlias = New System.Windows.Forms.ComboBox()
        Me.ChkDuplPartnos = New System.Windows.Forms.CheckBox()
        Me.CmbFieldDelim = New System.Windows.Forms.ComboBox()
        Me.LblFieldDelim = New System.Windows.Forms.Label()
        Me.ChkUseOtherMdbDir = New System.Windows.Forms.CheckBox()
        Me.ChkDynamicGui = New System.Windows.Forms.CheckBox()
        Me.ChkUseCellPinCount = New System.Windows.Forms.CheckBox()
        Me.ChkDbcInLibDir = New System.Windows.Forms.CheckBox()
        Me.ChkUseCLibSymbols = New System.Windows.Forms.CheckBox()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.BtnCancel, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.BtnOK, 0, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(42, 299)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 29)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'BtnCancel
        '
        Me.BtnCancel.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnCancel.Location = New System.Drawing.Point(76, 3)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(67, 23)
        Me.BtnCancel.TabIndex = 1
        Me.BtnCancel.Text = "Cancel"
        '
        'BtnOK
        '
        Me.BtnOK.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.BtnOK.Location = New System.Drawing.Point(3, 3)
        Me.BtnOK.Name = "BtnOK"
        Me.BtnOK.Size = New System.Drawing.Size(67, 23)
        Me.BtnOK.TabIndex = 0
        Me.BtnOK.Text = "OK"
        '
        'ChkEnaPrpImport
        '
        Me.ChkEnaPrpImport.AutoSize = True
        Me.ChkEnaPrpImport.Location = New System.Drawing.Point(24, 231)
        Me.ChkEnaPrpImport.Name = "ChkEnaPrpImport"
        Me.ChkEnaPrpImport.Size = New System.Drawing.Size(144, 17)
        Me.ChkEnaPrpImport.TabIndex = 1
        Me.ChkEnaPrpImport.Text = "Import Ascii Property  File"
        Me.ChkEnaPrpImport.UseVisualStyleBackColor = True
        '
        'ChkUseOdbcAlias
        '
        Me.ChkUseOdbcAlias.AutoSize = True
        Me.ChkUseOdbcAlias.Location = New System.Drawing.Point(24, 129)
        Me.ChkUseOdbcAlias.Name = "ChkUseOdbcAlias"
        Me.ChkUseOdbcAlias.Size = New System.Drawing.Size(136, 17)
        Me.ChkUseOdbcAlias.TabIndex = 1
        Me.ChkUseOdbcAlias.Text = "Use ODBC Alias (MDB)"
        Me.ChkUseOdbcAlias.UseVisualStyleBackColor = True
        '
        'CmbOdbcAlias
        '
        Me.CmbOdbcAlias.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbOdbcAlias.Enabled = False
        Me.CmbOdbcAlias.FormattingEnabled = True
        Me.CmbOdbcAlias.Location = New System.Drawing.Point(49, 152)
        Me.CmbOdbcAlias.Name = "CmbOdbcAlias"
        Me.CmbOdbcAlias.Size = New System.Drawing.Size(155, 21)
        Me.CmbOdbcAlias.TabIndex = 6
        '
        'ChkDuplPartnos
        '
        Me.ChkDuplPartnos.AutoSize = True
        Me.ChkDuplPartnos.Location = New System.Drawing.Point(24, 37)
        Me.ChkDuplPartnos.Name = "ChkDuplPartnos"
        Me.ChkDuplPartnos.Size = New System.Drawing.Size(159, 17)
        Me.ChkDuplPartnos.TabIndex = 5
        Me.ChkDuplPartnos.Text = "Allow duplicate Partnumbers"
        Me.ChkDuplPartnos.UseVisualStyleBackColor = True
        '
        'CmbFieldDelim
        '
        Me.CmbFieldDelim.Enabled = False
        Me.CmbFieldDelim.FormattingEnabled = True
        Me.CmbFieldDelim.Items.AddRange(New Object() {"tab", ";", "@", "§", "#", "*", "|"})
        Me.CmbFieldDelim.Location = New System.Drawing.Point(143, 254)
        Me.CmbFieldDelim.Name = "CmbFieldDelim"
        Me.CmbFieldDelim.Size = New System.Drawing.Size(60, 21)
        Me.CmbFieldDelim.TabIndex = 18
        Me.CmbFieldDelim.Text = "tab"
        '
        'LblFieldDelim
        '
        Me.LblFieldDelim.AutoSize = True
        Me.LblFieldDelim.Enabled = False
        Me.LblFieldDelim.Location = New System.Drawing.Point(40, 257)
        Me.LblFieldDelim.Name = "LblFieldDelim"
        Me.LblFieldDelim.Size = New System.Drawing.Size(97, 13)
        Me.LblFieldDelim.TabIndex = 19
        Me.LblFieldDelim.Text = "Ascii Field Delimiter"
        '
        'ChkUseOtherMdbDir
        '
        Me.ChkUseOtherMdbDir.AutoSize = True
        Me.ChkUseOtherMdbDir.Location = New System.Drawing.Point(24, 60)
        Me.ChkUseOtherMdbDir.Name = "ChkUseOtherMdbDir"
        Me.ChkUseOtherMdbDir.Size = New System.Drawing.Size(197, 17)
        Me.ChkUseOtherMdbDir.TabIndex = 20
        Me.ChkUseOtherMdbDir.Text = "Use alternate MDB Dir (not LMC Dir)"
        Me.ChkUseOtherMdbDir.UseVisualStyleBackColor = True
        '
        'ChkDynamicGui
        '
        Me.ChkDynamicGui.AutoSize = True
        Me.ChkDynamicGui.Checked = True
        Me.ChkDynamicGui.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkDynamicGui.Location = New System.Drawing.Point(24, 12)
        Me.ChkDynamicGui.Name = "ChkDynamicGui"
        Me.ChkDynamicGui.Size = New System.Drawing.Size(89, 17)
        Me.ChkDynamicGui.TabIndex = 21
        Me.ChkDynamicGui.Text = "Dynamic GUI"
        Me.ChkDynamicGui.UseVisualStyleBackColor = True
        '
        'ChkUseCellPinCount
        '
        Me.ChkUseCellPinCount.AutoSize = True
        Me.ChkUseCellPinCount.Location = New System.Drawing.Point(24, 106)
        Me.ChkUseCellPinCount.Name = "ChkUseCellPinCount"
        Me.ChkUseCellPinCount.Size = New System.Drawing.Size(111, 17)
        Me.ChkUseCellPinCount.TabIndex = 22
        Me.ChkUseCellPinCount.Text = "Use CellPin Count"
        Me.ChkUseCellPinCount.UseVisualStyleBackColor = True
        '
        'ChkDbcInLibDir
        '
        Me.ChkDbcInLibDir.AutoSize = True
        Me.ChkDbcInLibDir.Checked = True
        Me.ChkDbcInLibDir.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkDbcInLibDir.Enabled = False
        Me.ChkDbcInLibDir.Location = New System.Drawing.Point(49, 83)
        Me.ChkDbcInLibDir.Name = "ChkDbcInLibDir"
        Me.ChkDbcInLibDir.Size = New System.Drawing.Size(154, 17)
        Me.ChkDbcInLibDir.TabIndex = 24
        Me.ChkDbcInLibDir.Text = "Store .dbc File in Lib Folder"
        Me.ChkDbcInLibDir.UseVisualStyleBackColor = True
        '
        'ChkUseCLibSymbols
        '
        Me.ChkUseCLibSymbols.AutoSize = True
        Me.ChkUseCLibSymbols.Location = New System.Drawing.Point(24, 179)
        Me.ChkUseCLibSymbols.Name = "ChkUseCLibSymbols"
        Me.ChkUseCLibSymbols.Size = New System.Drawing.Size(157, 17)
        Me.ChkUseCLibSymbols.TabIndex = 25
        Me.ChkUseCLibSymbols.Text = "Use Central Library Symbols"
        Me.ChkUseCLibSymbols.UseVisualStyleBackColor = True
        '
        'DialogToolCfg
        '
        Me.AcceptButton = Me.BtnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnCancel
        Me.ClientSize = New System.Drawing.Size(231, 340)
        Me.Controls.Add(Me.ChkUseCLibSymbols)
        Me.Controls.Add(Me.ChkDbcInLibDir)
        Me.Controls.Add(Me.ChkUseCellPinCount)
        Me.Controls.Add(Me.ChkDynamicGui)
        Me.Controls.Add(Me.ChkUseOtherMdbDir)
        Me.Controls.Add(Me.LblFieldDelim)
        Me.Controls.Add(Me.CmbFieldDelim)
        Me.Controls.Add(Me.CmbOdbcAlias)
        Me.Controls.Add(Me.ChkDuplPartnos)
        Me.Controls.Add(Me.ChkUseOdbcAlias)
        Me.Controls.Add(Me.ChkEnaPrpImport)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DialogToolCfg"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Create DxDatabook Configuration"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents BtnOK As System.Windows.Forms.Button
    Friend WithEvents BtnCancel As System.Windows.Forms.Button
    Friend WithEvents ChkEnaPrpImport As System.Windows.Forms.CheckBox
    Friend WithEvents ChkUseOdbcAlias As System.Windows.Forms.CheckBox
    Friend WithEvents CmbOdbcAlias As System.Windows.Forms.ComboBox
    Friend WithEvents ChkDuplPartnos As System.Windows.Forms.CheckBox
    Friend WithEvents CmbFieldDelim As System.Windows.Forms.ComboBox
    Friend WithEvents LblFieldDelim As System.Windows.Forms.Label
    Friend WithEvents ChkUseOtherMdbDir As System.Windows.Forms.CheckBox
    Friend WithEvents ChkDynamicGui As System.Windows.Forms.CheckBox
    Friend WithEvents ChkUseCellPinCount As System.Windows.Forms.CheckBox
    Friend WithEvents ChkDbcInLibDir As System.Windows.Forms.CheckBox
    Friend WithEvents ChkUseCLibSymbols As Windows.Forms.CheckBox
End Class
