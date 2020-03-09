Imports System.Drawing

Public Enum DbcAction
    DbcOverride
    DbcUpdate
    DoNothing
End Enum

Public Class DlgDbcCfg
    Private DlgResult As DbcAction

    Public Shared Function CreateAndShow(ByVal ParentLoc As Point, ByVal ParentSize As Size) As DbcAction
        Dim Dlg As New DlgDbcCfg(ParentLoc, ParentSize)
        Dlg.ShowDialog()
        Return Dlg.DlgResult
    End Function

    Public Sub New(ByVal ParentLoc As Point, ByVal ParentSize As Size)
        Me.InitializeComponent()
        Me.Location = Me.SetLocation(ParentLoc, ParentSize)
    End Sub

    Private Sub DlgDbcCfg_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.DlgResult = DbcAction.DoNothing
        Me.TopMost = True
    End Sub

    Private Sub BtnDoNothing_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnDoNothing.Click
        Me.DlgResult = DbcAction.DoNothing
        Me.Hide()
        Me.Close()
    End Sub

    Private Sub BtnOverride_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnOverride.Click
        Me.DlgResult = DbcAction.DbcOverride
        Me.Hide()
        Me.Close()
    End Sub

    Private Sub BtnUpdate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnUpdate.Click
        Me.DlgResult = DbcAction.DbcUpdate
        Me.Hide()
        Me.Close()
    End Sub

    Private Function SetLocation(ByVal pLoc As Point, ByVal pSize As Size) As Point
        Dim dX As Integer = pSize.Width - Me.Width
        Dim dY As Integer = pSize.Height - Me.Height
        Dim MyLoc As New Point(CInt(pLoc.X + (dX / 2)), CInt(pLoc.Y + (dY / 2)))
        Return MyLoc
    End Function
End Class