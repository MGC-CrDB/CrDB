<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSelectApp
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
        Me.LvAppInstances = New System.Windows.Forms.ListView()
        Me.BtnSelect = New System.Windows.Forms.Button()
        Me.BtnCancel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'LvAppInstances
        '
        Me.LvAppInstances.Location = New System.Drawing.Point(8, 9)
        Me.LvAppInstances.Name = "LvAppInstances"
        Me.LvAppInstances.Size = New System.Drawing.Size(678, 155)
        Me.LvAppInstances.TabIndex = 0
        Me.LvAppInstances.UseCompatibleStateImageBehavior = False
        '
        'BtnSelect
        '
        Me.BtnSelect.Enabled = False
        Me.BtnSelect.Location = New System.Drawing.Point(235, 170)
        Me.BtnSelect.Name = "BtnSelect"
        Me.BtnSelect.Size = New System.Drawing.Size(88, 22)
        Me.BtnSelect.TabIndex = 1
        Me.BtnSelect.Text = "Select"
        Me.BtnSelect.UseVisualStyleBackColor = True
        '
        'BtnCancel
        '
        Me.BtnCancel.Location = New System.Drawing.Point(372, 170)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(88, 22)
        Me.BtnCancel.TabIndex = 1
        Me.BtnCancel.Text = "Abort"
        Me.BtnCancel.UseVisualStyleBackColor = True
        '
        'FrmSelectApp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(694, 201)
        Me.ControlBox = False
        Me.Controls.Add(Me.BtnCancel)
        Me.Controls.Add(Me.BtnSelect)
        Me.Controls.Add(Me.LvAppInstances)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximumSize = New System.Drawing.Size(700, 230)
        Me.MinimumSize = New System.Drawing.Size(700, 230)
        Me.Name = "FrmSelectApp"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Select Application Instance"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents LvAppInstances As System.Windows.Forms.ListView
    Friend WithEvents BtnSelect As System.Windows.Forms.Button
    Friend WithEvents BtnCancel As System.Windows.Forms.Button
End Class
