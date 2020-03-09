Imports System.IO
Imports System.Threading
Imports MGCPCBPartsEditor
Imports System.Globalization
Imports Microsoft.VisualBasic.FileIO
Imports System.Runtime.Serialization.Formatters.Binary

Namespace DxDB
    Public Class Pdb
        Inherits BaseClass

#Region "Members"
        Private ComVersion As Integer
        Private _LmcFile As String
        Private _Partitions As List(Of PdbPrtn)
#End Region

#Region "Properties"
        Public ReadOnly Property LmcFile() As String
            Get
                Return Me._LmcFile
            End Get
        End Property

        Public ReadOnly Property Partitions() As List(Of PdbPrtn)
            Get
                Return Me._Partitions
            End Get
        End Property

        Public ReadOnly Property Partition(ByVal p As Integer) As PdbPrtn
            Get
                Return Me._Partitions.Item(p)
            End Get
        End Property
#End Region

#Region "Constructor Methods"
        Public Sub New(LmcFile As String, nComVersion As Integer)
            MyBase.New(nComVersion)
            Me.ComVersion = nComVersion
            Me._LmcFile = LmcFile
            Me._Partitions = New List(Of PdbPrtn)
        End Sub

        Public Sub New(CurPdb As Pdb, nComVersion As Integer)
            MyBase.New(nComVersion)
            Me._LmcFile = CurPdb.LmcFile
            Me._Partitions = New List(Of PdbPrtn)
            For Each Prtn As PdbPrtn In CurPdb.Partitions
                Me._Partitions.Add(Prtn)
            Next
        End Sub
#End Region

#Region "Public Methods"
        Public Sub ReadPDB(UseSymbolTable As Boolean, UseCellPins As Boolean)
            Dim LibDir As String, MgcPrtn As Partition
            'Dim PdbEdit As New PartsEditorDlg
            Dim PdbEdit As PartsEditorDlg
            Dim PartsDB As PartsDB = Nothing

            PdbEdit = CType(CreateObject("MGCPCBLibraries.PartsEditorDlg"), PartsEditorDlg)

            ' automation objects
            ' clear Partitions
            Me._Partitions.Clear()

            If IsNothing(PdbEdit) Then Exit Sub
            If Not My.Computer.FileSystem.FileExists(Me.LmcFile) Then Exit Sub
            LibDir = Path.GetDirectoryName(Me.LmcFile)

            If IsNothing(PartsDB) And My.Computer.FileSystem.FileExists(Me.LmcFile) Then
                Try
                    Me.DxDbTranscriptMsg("Opening LMC Parts Database ...", 5)
                    PartsDB = PdbEdit.OpenDatabaseEx(Me.LmcFile, True)
                    Me.DxDbTranscriptMsg("LMC Parts Database opened successfully.", 5)
                    Me.DxDbTranscriptMsg("")
                Catch ex As Exception
                    Me.DxDbSysException("ReadPdbPartsDb(): " + Me.LmcFile, ex, 5)
                    Me.DxDbTranscriptMsg("")
                    Exit Sub
                End Try
            End If

            Me.DxDbStatusbarMsg("Reading LMC Parts Database")
            Me.DxDbTranscriptMsg("Reading LMC Parts Database", 5)
            Me.DxDbProgBarAction(PartsDB.Partitions.Count, ProgbarMode.Init)

            For Each MgcPrtn In PartsDB.Partitions
                Me.DxDbProgBarAction(0, ProgbarMode.Incr)

                Dim pdbPrtn As New PdbPrtn(MgcPrtn.Name, Me.ComVersion)

                Me.AddPrtnEventHandler(pdbPrtn)

                pdbPrtn.ReadPdbPrtn(MgcPrtn, LibDir, PdbEdit.Gui.Unit, UseSymbolTable, UseCellPins)

                Me.Partitions.Add(pdbPrtn)

                Me.RemPrtnEventHandler(pdbPrtn)
            Next
            Me.DxDbProgBarAction(0, ProgbarMode.Close)

            Me.DxDbTranscriptMsg("Reading LMC Parts Database done", 5)
            Me.DxDbTranscriptMsg("")
            Me.DxDbStatusbarMsg("Reading LMC Parts Database done")

            PdbEdit.Quit()
            PdbEdit = Nothing
        End Sub

        Public Sub MergePropsDb(ByRef AsciiPropsDb As SideFileDB)
            Dim i As Integer, Prtn As PdbPrtn

            If AsciiPropsDb.Parts.Count = 0 Then Exit Sub
            Me.DxDbTranscriptMsg("Merging Properties to DxDatabook", 5)

            '--------------------------------------
            ' create empty table for each partition
            Me.DxDbProgBarAction(Me.Partitions.Count, ProgbarMode.Init)
            For i = 0 To Me.Partitions.Count - 1
                ' set tablename to partition name
                Prtn = CType(Me.Partition(i), PdbPrtn)
                Me.AddPrtnEventHandler(Prtn)

                Me.DxDbTranscriptMsg("... processing Partition '" + Path.GetFileNameWithoutExtension(Prtn.Name) + "'", 10)
                Prtn.AddSideFileProps2Parts(AsciiPropsDb)
                Me.DxDbProgBarAction(0, ProgbarMode.Incr)

                Me.RemPrtnEventHandler(Prtn)
            Next

            If AsciiPropsDb.Parts.Count > 0 Then
                Me.DxDbAdd2Logfile("")
                Me.DxDbAdd2Logfile("Unmatched Sidefile Entries:", 10)
                For Each Prt As PdbPart In AsciiPropsDb.Parts
                    Me.DxDbAdd2Logfile(Prt.PartNo, 15)
                Next
            End If

            Me.DxDbProgBarAction(0, ProgbarMode.Close)
            Me.DxDbTranscriptMsg("Merging Properties to DxDatabook done", 5)
            Me.DxDbTranscriptMsg("")
        End Sub
#End Region

#Region "Add EventHandler Methods"
        Protected Sub AddPrtnEventHandler(ByRef PdbPartition As PdbPrtn)
            AddHandler PdbPartition.StatusbarMsg, AddressOf Me.DxDbStatusbarMsg
            AddHandler PdbPartition.ProgBarAction, AddressOf Me.DxDbProgBarAction
            AddHandler PdbPartition.SysException, AddressOf Me.DxDbSysException
            AddHandler PdbPartition.Add2Logfile, AddressOf Me.DxDbAdd2Logfile
            AddHandler PdbPartition.TranscriptMsg, AddressOf Me.DxDbTranscriptMsg
        End Sub

        Protected Sub RemPrtnEventHandler(ByRef PdbPartition As PdbPrtn)
            RemoveHandler PdbPartition.StatusbarMsg, AddressOf Me.DxDbStatusbarMsg
            RemoveHandler PdbPartition.ProgBarAction, AddressOf Me.DxDbProgBarAction
            RemoveHandler PdbPartition.SysException, AddressOf Me.DxDbSysException
            RemoveHandler PdbPartition.Add2Logfile, AddressOf Me.DxDbAdd2Logfile
            RemoveHandler PdbPartition.TranscriptMsg, AddressOf Me.DxDbTranscriptMsg
        End Sub
#End Region

    End Class
End Namespace