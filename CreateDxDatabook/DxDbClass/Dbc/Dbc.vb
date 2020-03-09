Imports System.IO
Imports System.Xml
Imports System.Threading
Imports System.Globalization
Imports Microsoft.VisualBasic.FileIO
Imports DxDbClass.DxDB

Namespace DxDB
    Public Class Dbc
        Inherits XmlFunc

#Region "Declarations, Properties & Variables"
        Private ComVersion As Integer
        Private _DbcFile As String
        Private _LmcFile As String

        Public ReadOnly Property DbcFile() As String
            Get
                Return Me._DbcFile
            End Get
        End Property

        Public ReadOnly Property LmcFile() As String
            Get
                Return Me._LmcFile
            End Get
        End Property
#End Region

#Region "Constructor Methods"
        Public Sub New(nLmcFile As String, nDbcFile As String, nComVersion As Integer, Optional ByVal PartsDb As Pdb = Nothing)
            MyBase.New(nComVersion)

            Me.ComVersion = nComVersion
            MyBase.UseOdbcAlias = False   ' Base Class Member
            MyBase.AllowDuplNos = False   ' Base Class Member
            MyBase.UseSymTable = False    ' Base Class Member

            Me.DbFile = ""
            Me.LmcPdb = Nothing

            Me._LmcFile = nLmcFile
            Me._DbcFile = nDbcFile

            ' set PartsDb if Not nothing
            If IsNothing(PartsDb) Then
                Me.LmcPdb = New Pdb(Me.LmcFile, nComVersion)
            Else
                Me.LmcPdb = New Pdb(PartsDb, nComVersion)
            End If
        End Sub
#End Region

#Region "Public Methods"
        Public Function LoadDbcConfig() As Boolean
            If File.Exists(DbcFile) Then
                Try
                    Me.DbcDoc.Load(Me.DbcFile)
                Catch ex As Exception
                    Return False
                End Try
                Return True
            End If
            Return False
        End Function

        Public Sub CreateDatabaseDbcCfg(nMdbFile As String, UseSymbolTable As Boolean, Optional AllowDuplPartNos As Boolean = False, Optional UseOdbcAlias As Boolean = False, Optional OdbcALias As String = "")
            Dim DbcLibraries As New List(Of XmlElement)
            Dim CCfgLibNode As XmlNode = Nothing

            MyBase.AllowDuplNos = AllowDuplPartNos
            MyBase.UseSymTable = UseSymbolTable

            Me.DbFile = nMdbFile
            If UseOdbcAlias And Not OdbcALias = String.Empty Then
                MyBase.UseOdbcAlias = UseOdbcAlias
                Me.DbFile = OdbcALias
            End If

            ' create template dbc
            If Not Me.CreateEmptyDbcCfg() Then Exit Sub

            '--------------------------------------
            ' get library config node <CConfigLib>
            '--------------------------------------
            If Not Me.GetNode(elemCConfigLib, CCfgLibNode) _
            And IsNothing(CCfgLibNode) Then Exit Sub

            '-----------------------------------
            ' create new XML library partitions
            '-----------------------------------
            DbcLibraries = Me.CreateLibPartitions(Me.LmcPdb.Partitions)

            '----------------------------------------------------
            ' add updated elements to <CConfigLibEntry/CObArray>
            '----------------------------------------------------
            Me.AddCObArrayChildNodes(CCfgLibNode, DbcLibraries)

            '--------------------------------------------------------------
            ' update library count <CObArray CXMLTypedPtrArraySize="(n)"\>
            '--------------------------------------------------------------
            Me.UpdateCObArraySize(CCfgLibNode)
        End Sub

        '------------------------------------------------------------
        ' Me.UpdateDbcConfig()
        '	-> Me.UpdateLibPartitions()
        '	<CConfigLib>/<CObArray>/<CConfigLibEntry>
        '------------------------------------------------------------
        Public Sub UpdateDatabaseDbcCfg(nDbFile As String, UseSymbolTable As Boolean, AllowDuplPartNos As Boolean, UseOdbcAlias As Boolean, OdbcALias As String)
            Dim DbcLibraries As New List(Of XmlElement)
            Dim CCfgLibNode As XmlNode = Nothing

            MyBase.AllowDuplNos = AllowDuplPartNos
            MyBase.UseSymTable = UseSymbolTable

            Me.DbFile = nDbFile
            If UseOdbcAlias And Not OdbcALias = String.Empty Then
                MyBase.UseOdbcAlias = UseOdbcAlias
                MyBase.DbFile = OdbcALias
            End If

            ' check DBC file
            If My.Computer.FileSystem.FileExists(Me._DbcFile) Then
                Me._DbcFile = DbcFile
                Me.DbcDoc.Load(Me.DbcFile)
                Me.DeIndexElements()
            Else
                ' create template dbc
                If Not Me.CreateEmptyDbcCfg() Then Exit Sub
            End If

            '--------------------------------------
            ' get library config node <CConfigLib>
            ' <CConfigLibEntry6 LibraryName="Analog" JoinTable="false" Locked="0" SymbolExpression="" JoinType="0" HorizJoinType="0">
            '--------------------------------------
            If Not Me.GetNode(elemCConfigLib, CCfgLibNode) _
            And IsNothing(CCfgLibNode) Then Exit Sub

            '----------------------------------------------------------------
            ' get/delete existing library entries <CConfigLibEntry(n) ....\>
            ' <CObArray8 CXMLTypedPtrArraySize="8">
            '   <CConfigAttEntry9 FieldName="Part Number" AttName="Part Number" DefaultValue="" ExcludeWhenAnnotating="0" ExcludeWhenLoading="0" m_bNameVisible="0" m_bValueVisible="0" AttType="4" ValueType="1" MagType="7" UnitsString="" ShowUnitsForIEC62="0" AddAsOat="0" Required="0" VerificationKey="0" ValidMag="2047" ValidMagDir="63" FutureUseBool1="0" FutureUseBool2="0" FutureUseStr1="" FutureUseStr2="" MatchCondition="--" AttrUnplace="5" SerialKey="0"/>
            '----------------------------------------------------------------
            DbcLibraries = Me.RemoveCObArrayChildNodes(CCfgLibNode)
            Me.DbcDoc.Save(Me.DbcFile)

            '----------------------------------
            ' updating existing XML partitions
            '----------------------------------
            DbcLibraries = Me.UpdateLibPartitions(DbcLibraries, Me.LmcPdb.Partitions)

            '----------------------------------------------------
            ' add updated elements to <CConfigLibEntry/CObArray>
            '----------------------------------------------------
            Me.AddCObArrayChildNodes(CCfgLibNode, DbcLibraries)

            '--------------------------------------------------------------
            ' update library count <CObArray CXMLTypedPtrArraySize="(n)"\>
            ' <CObArray8 CXMLTypedPtrArraySize="8">
            '--------------------------------------------------------------
            Me.UpdateCObArraySize(CCfgLibNode)
        End Sub

        Public Sub Save()
            Me.ReIndexElements()
            If Me.DbcFile <> String.Empty Then
                Me.DbcDoc.Save(Me.DbcFile)
                Dim Buff As List(Of String) = Me.ReadFileToVect(Me.DbcFile)
                'If Buff(Buff.Count - 1).ToString = "</dbc>" Then
                '    Buff.Add(String.Empty)
                'End If
                Me.WriteTextFile(Me.DbcFile, Buff)
            End If
        End Sub
#End Region

#Region "Protected Methods"
        Protected Function ReadFileToVect(ByVal FilePath As String) As List(Of String)
            Dim FileVect As New List(Of String)
            Dim Fline As String, Sr As IO.StreamReader

            Try
                If My.Computer.FileSystem.FileExists(FilePath) = False Then
                    Return FileVect
                End If

                Sr = New StreamReader(FilePath, System.Text.Encoding.Default)
                Do
                    Fline = Sr.ReadLine()
                    If IsNothing(Fline) Then Exit Do
                    If Fline.ToString <> "" Then
                        FileVect.Add(Fline)
                    End If
                Loop
                Sr.Close()
                Sr.Dispose()

                Return FileVect
            Catch Ex As System.Exception
                Throw New Exception("Exception in ReadFileToVect(): " + Ex.Message, Ex)
            End Try
            Return Nothing
        End Function

        Protected Sub WriteTextFile(ByVal FileName As String, ByVal FileVect As List(Of String))
            Dim sw As System.IO.StreamWriter
            Dim FileEnumerator As System.Collections.IEnumerator

            Try
                If IsNothing(FileVect) Then Exit Sub
                If FileVect.Count = 0 Then Exit Sub

                If My.Computer.FileSystem.FileExists(FileName) Then
                    My.Computer.FileSystem.DeleteFile(FileName) ' Remove the existing file
                End If

                'Loop through the arraylist (Content) and write each line to the file
                sw = New StreamWriter(FileName, False, System.Text.Encoding.Default)

                FileEnumerator = FileVect.GetEnumerator()

                While FileEnumerator.MoveNext()
                    sw.WriteLine(FileEnumerator.Current.ToString)
                End While

                sw.Close()
                sw.Dispose()
            Catch ex As Exception
                Throw New Exception("Exception in WriteTextFile(): " + ex.Message, ex)
            End Try
        End Sub

        Protected Function RemFileExt(ByVal FileName As String) As String
            Dim FileNameNoExt As String = Path.GetFileNameWithoutExtension(FileName)
            Dim DirPath As String = Path.GetDirectoryName(FileName)
            If DirPath = String.Empty Then Return FileNameNoExt
            Return DirPath + "\" + FileNameNoExt
        End Function
#End Region

    End Class
End Namespace

