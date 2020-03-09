<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Transcript
    Inherits System.Windows.Forms.UserControl

    'UserControl überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
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

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.MsgList = New System.Windows.Forms.ListView()
        Me.MsgListPopup = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.CopySelToClipboard = New System.Windows.Forms.ToolStripMenuItem()
        Me.CopyAllToClipboard = New System.Windows.Forms.ToolStripMenuItem()
        Me.MsgListPopup.SuspendLayout()
        Me.SuspendLayout()
        '
        'MsgList
        '
        Me.MsgList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.MsgList.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Bold)
        Me.MsgList.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.MsgList.Location = New System.Drawing.Point(0, 0)
        Me.MsgList.Name = "MsgList"
        Me.MsgList.Size = New System.Drawing.Size(621, 524)
        Me.MsgList.TabIndex = 0
        Me.MsgList.UseCompatibleStateImageBehavior = False
        Me.MsgList.View = System.Windows.Forms.View.Details
        '
        'MsgListPopup
        '
        Me.MsgListPopup.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CopySelToClipboard, Me.CopyAllToClipboard})
        Me.MsgListPopup.Name = "MsgListPopup"
        Me.MsgListPopup.RenderMode = System.Windows.Forms.ToolStripRenderMode.System
        Me.MsgListPopup.ShowImageMargin = False
        Me.MsgListPopup.Size = New System.Drawing.Size(175, 70)
        '
        'CopySelToClipboard
        '
        Me.CopySelToClipboard.Name = "CopySelToClipboard"
        Me.CopySelToClipboard.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.C), System.Windows.Forms.Keys)
        Me.CopySelToClipboard.Size = New System.Drawing.Size(174, 22)
        Me.CopySelToClipboard.Text = "Copy Selection"
        '
        'CopyAllToClipboard
        '
        Me.CopyAllToClipboard.Name = "CopyAllToClipboard"
        Me.CopyAllToClipboard.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.T), System.Windows.Forms.Keys)
        Me.CopyAllToClipboard.Size = New System.Drawing.Size(174, 22)
        Me.CopyAllToClipboard.Text = "Copy Transcript"
        '
        'Transcript
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.MsgList)
        Me.Name = "Transcript"
        Me.Size = New System.Drawing.Size(621, 524)
        Me.MsgListPopup.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents MsgList As System.Windows.Forms.ListView
    Friend WithEvents MsgListPopup As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents CopySelToClipboard As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CopyAllToClipboard As System.Windows.Forms.ToolStripMenuItem

End Class
