Imports System.IO
Imports System.Text
Imports Microsoft.VisualBasic.FileIO

<System.ComponentModel.ToolboxItem(True)> _
Public Class Transcript

#Region "Declarations"
    Public Enum MsgType
        Txt
        Nte
        Wrn
        Err
    End Enum

    Public ConsoleMode As Boolean

    Private SelectLines As Boolean
    Private _NumNotes As Long
    Private _NumWarngs As Long
    Private _NumErrors As Long

    Private _LogMsgs As List(Of String)

    Public ReadOnly Property NumNotes() As Long
        Get
            Return Me._NumNotes
        End Get
    End Property

    Public ReadOnly Property NumWarngs() As Long
        Get
            Return Me._NumWarngs
        End Get
    End Property

    Public ReadOnly Property NumErrors() As Long
        Get
            Return Me._NumErrors
        End Get
    End Property

    Public ReadOnly Property LogMsgs() As List(Of String)
        Get
            Return Me._LogMsgs
        End Get
    End Property
#End Region

#Region "Public Methods"
    Private Sub Transcript_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Me.InitControl()
        Me.SelectLines = False
        Me._NumNotes = 0
        Me._NumWarngs = 0
        Me._NumErrors = 0
        Me._LogMsgs = New List(Of String)
        Me.ConsoleMode = False
    End Sub

    Delegate Sub ClearCallback(ByVal ClearLog As Boolean)
    Public Sub Clear(Optional ByVal ClearLog As Boolean = False)
        If Me.MsgList.InvokeRequired Then
            Dim d As New ClearCallback(AddressOf Clear)
            Me.Invoke(d, New Object() {ClearLog})
        Else
            Me.MsgList.Items.Clear()
            If ClearLog Then Me.ClearLog()
        End If
    End Sub

    Public Sub SetBoxStyle(ByVal BrdStyle As BorderStyle)
        Me.BorderStyle = BrdStyle
    End Sub

    Delegate Sub MessageCallback(ByVal [TextStr] As String, ByVal [LeadingSpaces] As Integer, ByVal [MsgType] As MsgType, ByVal AddLogMsg As Boolean)
    Public Sub Message(ByVal TextStr As String, Optional ByVal LeadingSpaces As Integer = 0, Optional ByVal MsgType As MsgType = MsgType.Txt, Optional ByVal AddLogMsg As Boolean = True)
        Dim MsgStr As String, Lspaces As String, LvItem As ListViewItem

        If Me.MsgList.InvokeRequired Then
            Dim d As New MessageCallback(AddressOf Message)
            Me.Invoke(d, New Object() {TextStr, LeadingSpaces, MsgType, AddLogMsg})
        Else
            If LeadingSpaces > 0 Then
                Lspaces = Space([LeadingSpaces])
            Else
                Lspaces = String.Empty
            End If

            Select Case MsgType
                Case MsgType.Txt
                    MsgStr = "# " + Lspaces + TextStr
                Case MsgType.Nte
                    MsgStr = "# " + Lspaces + "NOTE! " + TextStr
                    Me._NumNotes += 1
                Case MsgType.Wrn
                    MsgStr = "# " + Lspaces + "WARNING! " + TextStr
                    Me._NumWarngs += 1
                Case MsgType.Err
                    MsgStr = "# " + Lspaces + "ERROR!! " + TextStr
                    Me._NumErrors += 1
                Case Else
                    MsgStr = "# " + Lspaces + TextStr
            End Select

            If Me.ConsoleMode Then
                Console.WriteLine(MsgStr)
            Else
                LvItem = Me.MsgList.Items.Add(MsgStr)
                Me.MsgList.Items.Item(LvItem.Index).Selected = False
                LvItem.EnsureVisible()
            End If

            If AddLogMsg And Not IsNothing(Me._LogMsgs) Then
                Me._LogMsgs.Add(MsgStr)
            End If
        End If
    End Sub

    Public Sub AddLogMessage(ByVal TextStr As String, Optional ByVal LeadingSpaces As Integer = 0, Optional ByVal MsgType As MsgType = MsgType.Txt)
        Dim MsgStr As String, Lspaces As String
        If Not IsNothing(Me._LogMsgs) Then
            If LeadingSpaces > 0 Then
                Lspaces = Space(LeadingSpaces)
            Else
                Lspaces = String.Empty
            End If

            Select Case MsgType
                Case MsgType.Txt
                    MsgStr = "# " + Lspaces + TextStr
                Case MsgType.Nte
                    MsgStr = "# " + Lspaces + "NOTE! " + TextStr
                    Me._NumNotes += 1
                Case MsgType.Wrn
                    MsgStr = "# " + Lspaces + "WARNING! " + TextStr
                    Me._NumWarngs += 1
                Case MsgType.Err
                    MsgStr = "# " + Lspaces + "ERROR!! " + TextStr
                    Me._NumErrors += 1
                Case Else
                    MsgStr = "# " + Lspaces + TextStr
            End Select

            Me._LogMsgs.Add(MsgStr)
        End If
    End Sub

    Delegate Sub PrintExceptionCallback(ByVal UserMsg As String, ByVal ex As Exception, ByVal Verbose As Boolean)
    Public Sub PrintException(ByVal UserMsg As String, ByVal ex As Exception, Optional ByVal Verbose As Boolean = False)
        Dim MsgVect As ArrayList = VectorizeExMsg(ex)

        If Me.InvokeRequired Then
            Dim d As New PrintExceptionCallback(AddressOf PrintException)
            Me.Invoke(d, New Object() {UserMsg, ex, Verbose})
        Else
            Me.Message("")
            Me.Message(ex.GetType.Name, 0, Transcript.MsgType.Err)
            Me.Message(UserMsg, 5)
            Me.Message(ex.Message, 5)
            If Verbose Then
                Me.Message(MsgVect(0).ToString, 10)
                Me.Message(MsgVect(MsgVect.Count - 1).ToString, 10)
            End If
            Me.Message("")
        End If
    End Sub

    Public Sub WriteLog2File(ByVal FilePath As String)
        Dim sw As StreamWriter, FileEnumerator As Collections.IEnumerator

        Try
            If IsNothing(Me._LogMsgs) Then Exit Sub
            'If FileVect.Count = 0 Then Exit Sub

            ' Remove the existing file
            If My.Computer.FileSystem.FileExists(FilePath) Then
                My.Computer.FileSystem.DeleteFile(FilePath, UIOption.OnlyErrorDialogs, RecycleOption.DeletePermanently)
            End If

            'Loop through the arraylist (Content) and write each line to the file
            sw = New StreamWriter(FilePath)

            FileEnumerator = Me._LogMsgs.GetEnumerator()

            While FileEnumerator.MoveNext()
                sw.WriteLine(FileEnumerator.Current.ToString)
            End While

            sw.Close()
            sw.Dispose()

            Me._LogMsgs.Clear()
        Catch ex As Exception
            Me.PrintException("WriteLog2File()", ex)
        End Try
    End Sub
#End Region

#Region "Private Methods"
    Protected Sub ClearLog()
        Me._NumNotes = 0
        Me._NumWarngs = 0
        Me._NumErrors = 0
        Me._LogMsgs.Clear()
    End Sub

    Protected Sub CopyTranscript()
        Dim Data As New StringBuilder
        Dim Item As ListViewItem

        For L As Integer = 0 To Me.MsgList.Items.Count - 1
            Item = Me.MsgList.Items(L)

            For M As Integer = 0 To Item.SubItems.Count - 1
                Data.Append(Item.SubItems(M).Text)
                If M <> Item.SubItems.Count - 1 Then
                    Data.Append(ControlChars.Tab)
                End If
            Next
            Data.Append(ControlChars.CrLf)
        Next

        Clipboard.SetDataObject(Data.ToString, True)
    End Sub

    Protected Sub CopySelected()
        Dim Data As New StringBuilder
        Dim Item As ListViewItem

        For L As Integer = 0 To Me.MsgList.Items.Count - 1
            Item = Me.MsgList.Items(L)
            If Item.Selected Then
                For M As Integer = 0 To Item.SubItems.Count - 1
                    Data.Append(Item.SubItems(M).Text)
                    If M <> Item.SubItems.Count - 1 Then
                        Data.Append(ControlChars.Tab)
                    End If
                Next
                Data.Append(ControlChars.CrLf)
            End If
        Next

        If Not Data.ToString = "" Then
            Clipboard.SetDataObject(Data.ToString, True)
        End If

        Me.UnselectAll()
    End Sub

    Protected Sub UnselectAll()
        For L As Integer = 0 To Me.MsgList.Items.Count - 1
            Me.MsgList.Items(L).Selected = False
        Next
        Me.MsgList.Select()
        Me.MsgList.Items(Me.MsgList.Items.Count - 1).EnsureVisible()
    End Sub

    Protected Sub InitControl()
        Me.MsgList.Visible = True
        Me.MsgList.FullRowSelect = False
        Me.MsgList.HeaderStyle = ColumnHeaderStyle.Nonclickable
        Me.MsgList.View = View.Details
        Me.MsgList.Columns.Clear()
        Me.MsgList.Items.Clear()
        Me.MsgList.Columns.Add("Transcript", (Me.Size.Width * 2))
    End Sub

    Protected Function VectorizeExMsg(ByVal ex As Exception) As ArrayList
        Dim errVect As New ArrayList, Start As Integer = 0
        For i As Integer = 0 To ex.ToString.Length - 1
            Select Case ex.ToString.Substring(i, 1)
                Case vbCr
                    errVect.Add(ex.ToString.Substring(i - Start, Start).Trim())
                Case vbLf
                    Start = 0
                Case Else
                    Start += 1
                    If i = ex.ToString.Length - 1 Then
                        errVect.Add(ex.ToString.Substring(i - Start).Trim())
                    End If
            End Select
        Next
        Return errVect
    End Function
#End Region

#Region "Control Events"
    Protected Sub MsgList_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles MsgList.KeyDown
        If e.KeyCode = Keys.C And e.Control Then
            If Me.MsgList.SelectedItems.Count > 0 Then
                Me.CopySelected()
            End If
        End If

        If e.KeyCode = Keys.T And e.Control Then
            Me.CopyTranscript()
        End If
    End Sub

    Protected Sub MsgList_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MsgList.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Left Then
            Me.SelectLines = True
        End If
    End Sub

    Protected Sub MsgList_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MsgList.MouseMove
        If Me.SelectLines Then
            Dim lvi As ListViewItem = Me.MsgList.GetItemAt(e.X, e.Y)
            If Not IsNothing(lvi) Then
                lvi.Selected = True
            End If
        End If
    End Sub

    Protected Sub MsgList_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MsgList.MouseUp
        Me.SelectLines = False
        If e.Button = Windows.Forms.MouseButtons.Right Then
            If TypeOf sender Is ListView Then
                Dim LV As ListView = CType(sender, ListView)
                If LV.Items.Count > 0 Then
                    Me.MsgListPopup.Show(Control.MousePosition, ToolStripDropDownDirection.BelowRight)
                End If
            End If
        End If
    End Sub

    Protected Sub MsgListPopup_ItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles MsgListPopup.ItemClicked
        Select Case e.ClickedItem.Name
            Case "CopySelToClipboard"
                If Me.MsgList.SelectedItems.Count > 0 Then
                    Me.CopySelected()
                End If
            Case "CopyAllToClipboard"
                Me.CopyTranscript()
            Case "ToggleRowSelect"
                Me.MsgList.FullRowSelect = Not Me.MsgList.FullRowSelect
        End Select
    End Sub
#End Region

End Class
