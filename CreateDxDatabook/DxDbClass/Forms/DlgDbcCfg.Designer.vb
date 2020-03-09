<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DlgDbcCfg
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.BtnOverride = New System.Windows.Forms.Button()
        Me.BtnUpdate = New System.Windows.Forms.Button()
        Me.BtnDoNothing = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(121, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(212, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "DBC Config File already exist! Select Option"
        '
        'BtnOverride
        '
        Me.BtnOverride.Location = New System.Drawing.Point(163, 34)
        Me.BtnOverride.Name = "BtnOverride"
        Me.BtnOverride.Size = New System.Drawing.Size(145, 22)
        Me.BtnOverride.TabIndex = 3
        Me.BtnOverride.Text = "Override Existing Config"
        Me.BtnOverride.UseVisualStyleBackColor = True
        '
        'BtnUpdate
        '
        Me.BtnUpdate.Location = New System.Drawing.Point(12, 34)
        Me.BtnUpdate.Name = "BtnUpdate"
        Me.BtnUpdate.Size = New System.Drawing.Size(145, 22)
        Me.BtnUpdate.TabIndex = 1
        Me.BtnUpdate.Text = "Update Existing Config"
        Me.BtnUpdate.UseVisualStyleBackColor = True
        '
        'BtnDoNothing
        '
        Me.BtnDoNothing.Location = New System.Drawing.Point(314, 34)
        Me.BtnDoNothing.Name = "BtnDoNothing"
        Me.BtnDoNothing.Size = New System.Drawing.Size(145, 22)
        Me.BtnDoNothing.TabIndex = 3
        Me.BtnDoNothing.Text = "Skip! Do Nothing"
        Me.BtnDoNothing.UseVisualStyleBackColor = True
        '
        'DlgDbcCfg
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(470, 72)
        Me.Controls.Add(Me.BtnDoNothing)
        Me.Controls.Add(Me.BtnUpdate)
        Me.Controls.Add(Me.BtnOverride)
        Me.Controls.Add(Me.Label1)
        Me.Name = "DlgDbcCfg"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Create DBC Config File"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents BtnOverride As System.Windows.Forms.Button
    Friend WithEvents BtnUpdate As System.Windows.Forms.Button
    Friend WithEvents BtnDoNothing As System.Windows.Forms.Button
End Class
