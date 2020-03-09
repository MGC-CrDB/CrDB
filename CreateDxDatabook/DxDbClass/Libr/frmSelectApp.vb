Imports System.IO
Imports System.Windows.Forms

Imports DxDbClass.LibVxApp

Public Class frmSelectApp
    Protected Apps As List(Of RunningApp)

    Public Shared Function CreateAndShow(ByRef nApps As List(Of RunningApp)) As RunningApp
        Dim Dlg As New frmSelectApp(nApps)
        Dim SelApp As RunningApp = Dlg.ShowWindow()
        If Dlg.DialogResult = Windows.Forms.DialogResult.Cancel Then
            Return Nothing
        End If
        Return SelApp
    End Function

    Public Sub New(ByRef nApps As List(Of RunningApp))
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        Me.Apps = nApps
        Me.AcceptButton = Me.BtnSelect
    End Sub

    Public Function ShowWindow() As RunningApp
        Me.ShowDialog()
        Me.BringToFront()
        If Me.DialogResult = Windows.Forms.DialogResult.OK Then
            For Each Lv As ListViewItem In Me.LvAppInstances.Items()
                If Lv.Selected Then Return CType(Lv.Tag, RunningApp)
            Next
        End If
        Return Nothing
    End Function

    Private Sub FrmSelectApp_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Dim Lv As ListViewItem

        Me.InitLvSelVar()

        For Each App As RunningApp In Me.Apps
            Lv = Me.LvAppInstances.Items.Add(App.Name)
            Lv.SubItems.Add(App.ReleaseName)
            Lv.SubItems.Add(Me.TruncatePath(App.LmcFilePath))
            Lv.Selected = False
            Lv.Tag = App
        Next

        Me.TopMost = True
        Me.BringToFront()
    End Sub

    Protected Sub InitLvSelVar()
        Me.Visible = True

        Me.LvAppInstances.View = View.Details
        Me.LvAppInstances.CheckBoxes = False
        Me.LvAppInstances.MultiSelect = False
        Me.LvAppInstances.FullRowSelect = True
        Me.LvAppInstances.Items.Clear()
        Me.LvAppInstances.Columns.Clear()

        Me.LvAppInstances.Columns.Add("")
        Me.LvAppInstances.Columns.Add("")
        Me.LvAppInstances.Columns.Add("")

        ' Benamung Möglichkeit 2
        For i As Integer = 0 To Me.LvAppInstances.Columns.Count - 1
            Select Case i
                Case 0
                    Me.LvAppInstances.Columns.Item(i).Text = "Name"
                    Me.LvAppInstances.Columns.Item(i).Width = CInt(Me.LvAppInstances.Width * 0.24)
                Case 1
                    Me.LvAppInstances.Columns.Item(i).Text = "Release"
                    Me.LvAppInstances.Columns.Item(i).Width = CInt(Me.LvAppInstances.Width * 0.11)
                Case 2
                    Me.LvAppInstances.Columns.Item(i).Text = "Document"
                    Me.LvAppInstances.Columns.Item(i).Width = CInt(Me.LvAppInstances.Width * 0.62)
            End Select
        Next
    End Sub

    Private Sub LvAppInstances_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles LvAppInstances.SelectedIndexChanged
        If Me.LvAppInstances.SelectedIndices.Count = 0 Then
            Me.BtnSelect.Enabled = False
        Else
            Me.BtnSelect.Enabled = True
        End If
    End Sub

    Private Sub BtnSelect_Click(sender As Object, e As System.EventArgs) Handles BtnSelect.Click
        If Me.LvAppInstances.SelectedIndices.Count = 0 Then Exit Sub
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As System.EventArgs) Handles BtnCancel.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Public Function TruncatePath(InPath As String, Optional ByVal DirLevel As Short = 5) As String
        Dim i As Integer, TruncPath As String = ""
        Dim FileName As String, Count As Integer

        If String.IsNullOrEmpty(InPath) Then Return String.Empty

        FileName = "\" + Path.GetFileName(InPath)
        InPath = Path.GetDirectoryName(InPath)

        Count = 0
        For i = Len(InPath) - 1 To 0 Step -1
            Select Case InPath(i).ToString
                Case ":"
                    Return InPath
                Case "\"
                    TruncPath = "..." + InPath.Substring(i) + FileName
                    Count += 1
            End Select
            If Count = DirLevel Then Exit For
        Next

        Return TruncPath
    End Function
End Class