Imports System.IO
Imports MGCPCBPartsEditor
Imports Microsoft.VisualBasic.FileIO

Namespace DxDB
    Public Class SideFileDB
        Inherits BaseClass

#Region "Declarations"
        Protected PartnoIdentifier() As String = {"PARTNO", "PARTNUMBER", "PART NUMBER", "PN"}

        ' Defines wether the first column is partnumber
        ' or any other property name to be search within
        ' each partnumber property list
        Public Enum SideFileDbMatchKey
            Partnumber
            [Property]
        End Enum

        Protected _PnKeyType As SideFileDbMatchKey
        Protected _PnKeyProp As String

        Protected _Parts As List(Of PdbPart)
        Protected _Attrs As List(Of PrtnAttr)
#End Region

#Region "Properties"
        Public ReadOnly Property PnKeyType As SideFileDbMatchKey
            Get
                Return Me._PnKeyType
            End Get
        End Property

        Public ReadOnly Property PnKeyProp As String
            Get
                Return Me._PnKeyProp
            End Get
        End Property

        Public ReadOnly Property Attrs As List(Of PrtnAttr)
            Get
                Return Me._Attrs
            End Get
        End Property

        Public Property Part(ByVal p As Integer) As PdbPart
            Get
                Return CType(Me._Parts(p), PdbPart)
            End Get
            Set(ByVal value As PdbPart)
                Me._Parts(p) = value
            End Set
        End Property

        Public ReadOnly Property Parts As List(Of PdbPart)
            Get
                Return Me._Parts
            End Get
        End Property
#End Region

#Region "Constructor"
        Public Sub New(nComVersion As Integer)
            MyBase.New(nComVersion)
            Me._PnKeyType = SideFileDbMatchKey.Partnumber
            Me._PnKeyProp = String.Empty
            Me._Parts = New List(Of PdbPart)
            Me._Attrs = New List(Of PrtnAttr)
        End Sub
#End Region

#Region "Public Methods"
        Public Function ReadPropsFile(SideFile As String, Delimiter As String) As Boolean
            Dim i As Integer, PartNo As PdbPart, DbProp As PdbProp
            Dim ReadPnPrps As TextFieldParser, ThisLine As String()
            Dim PropIndex As Integer, DefProp As PrtnAttr

            If Not My.Computer.FileSystem.FileExists(SideFile) Then Return False

            Me.DxDbTranscriptMsg("Reading Partnumber Property Sidefile", 5)
            Me.DxDbTranscriptMsg("... processing File '" + Me.TruncatePath(SideFile) + "'", 10)

            Try
                PartNo = Nothing

                ReadPnPrps = New TextFieldParser(SideFile)
                ReadPnPrps.TextFieldType = FieldType.Delimited
                ReadPnPrps.SetDelimiters(Delimiter)

                ' 1st line is property file header (property names)
                ThisLine = Me.RemoveEmptyArrayFields(ReadPnPrps.ReadFields())
                Me.SetPartNumberIdentifier(ThisLine(0))

                ' 1st column is Partnumber so start with 2nd column
                Me.DxDbAdd2Logfile("... parsing Header Properties", 15)
                For i = 1 To ThisLine.Length - 1
                    If String.IsNullOrEmpty(ThisLine(i)) Then Continue For
                    PropIndex += 1
                    ' create new attribute
                    Dim DbAttr As New PrtnAttr
                    DbAttr.Name = ThisLine(i)
                    DbAttr.Index = PropIndex

                    DbProp = New PdbProp()
                    DbProp.Name = ThisLine(i)
                    DbProp.PdbType = EPDBPropertyType.epdbPropTypeString
                    DbProp.Value = "<val>"
                    DbProp.Index = i

                    Me.DxDbAdd2Logfile("Found Property Name: " + DbProp.Name, 20)
                    Me.AddPropDbAttr(DbProp)
                Next
                Me.DxDbAdd2Logfile("Property Count: " + Me._Attrs.Count.ToString, 20)

                Me.DxDbAdd2Logfile("... parsing Partnumber Lines", 15)
                While Not ReadPnPrps.EndOfData
                    Try
                        'ThisLine = Me.RemoveEmptyArrayFields(ReadPnPrps.ReadFields())
                        ThisLine = ReadPnPrps.ReadFields()
                        PartNo = New PdbPart(ThisLine(0))

                        ' add properties. start at index 1
                        ' to ommit partnumber field
                        For i = 0 To Me.Attrs.Count - 1
                            DefProp = CType(Me.Attrs(i), PrtnAttr)

                            If DefProp.Index > -1 Then
                                ' create property
                                DbProp = New PdbProp()
                                DbProp.Name = DefProp.Name
                                DbProp.PdbType = EPDBPropertyType.epdbPropTypeString
                                DbProp.Value = ThisLine(DefProp.Index)
                                DbProp.Index = DefProp.Index
                                PartNo.Props.Add(DbProp)
                            End If
                        Next
                    Catch ex As Exception
                        Me.DxDbTranscriptMsg("---> Skipping inconsistent Line ...", 5)
                        Me.DxDbSysException("ReadPropsFile()", ex, 5)
                    End Try

                    Me.DxDbAdd2Logfile("Found PartNumber: " + PartNo.PartNo + " Props(" + PartNo.Props.Count.ToString + ")", 20)
                    Me._Parts.Add(PartNo)
                End While

                Me.DxDbTranscriptMsg("Reading Partnumber Property Sidefile done.", 5)
                Me.DxDbTranscriptMsg("")
            Catch ex As System.Exception
                Me.DxDbSysException("ReadPropsFile()", ex, 5)
                Return False
            End Try

            If Me.Parts.Count > 0 Then Return True
            Return False
        End Function

        Public Function FindPartByProperty(ByVal Part As PdbPart, ByRef SideFileProps As List(Of PdbProp)) As Boolean
            Dim MatchinProp As String = String.Empty

            ' PnKeyType=Property search for specific property in partnumber
            For Each Prp As PdbProp In Part.Props
                If Prp.Name.ToUpper = Me.PnKeyProp.ToUpper Then
                    MatchinProp = Prp.Value
                    Exit For
                End If
            Next

            If MatchinProp = String.Empty Then Return False

            Return Me.FindPartByPartnumber(MatchinProp, SideFileProps)
        End Function

        Public Function FindPartByPartnumber(ByVal PartNo As String, ByRef SideFileProps As List(Of PdbProp), Optional ByVal UniquePN As Boolean = False) As Boolean
            Dim i, n As Integer, Tmp As New List(Of PdbProp), Part As PdbPart
            Dim Matches As Integer, PnMatch As Boolean
            Dim DelIdx As New ArrayList

            SideFileProps.Clear()
            Matches = 0
            PnMatch = False

            ' each Partnumber from LMC can occur 
            ' multiple times in the PRP source file
            For i = 0 To Me.Parts.Count - 1
                Part = CType(Me.Parts(i), PdbPart)

                If LCase(Part.PartNo) = LCase(PartNo) Then
                    PnMatch = True
                    DelIdx.Add(i) ' store index to remove later
                End If

                If PnMatch And Not LCase(Part.PartNo) = LCase(PartNo) Then
                    PnMatch = False
                End If

                If PnMatch Then
                    Matches += 1
                    For n = 0 To Part.Props.Count - 1
                        Tmp.Add(Part.Props(n))
                    Next
                End If
            Next

            ' delete all matching partnumbers from PropDB
            ' don't delete if sidefile entry is used for
            ' multiple partnumbers 
            If UniquePN Then Me.DelArrayItems(DelIdx)
            If Matches > 1 Then
                SideFileProps = Me.MakeUniquePropListMergeValues(Tmp)
            Else
                SideFileProps = Tmp
            End If

            If SideFileProps.Count > 0 Then Return True
            Return False
        End Function
#End Region

#Region "Protected Methods"
        Protected Sub AddPropDbAttr(ByVal Prop As PdbProp)
            ' create new attribute
            Dim Attr As New PrtnAttr
            Attr.Name = Prop.Name
            Attr.Index = Prop.Index

            Select Case Prop.PdbType
                Case EPDBPropertyType.epdbPropTypeInt
                    Attr.PdbType = EPDBPropertyType.epdbPropTypeInt
                    Attr.DaoType = DAODataType.dbInteger
                Case EPDBPropertyType.epdbPropTypeReal
                    Attr.PdbType = EPDBPropertyType.epdbPropTypeReal
                    Attr.DaoType = DAODataType.dbDouble
                Case EPDBPropertyType.epdbPropTypeString
                    Attr.PdbType = EPDBPropertyType.epdbPropTypeString
                    Attr.DaoType = DAODataType.dbText
                Case Else
                    Attr.PdbType = EPDBPropertyType.epdbPropTypeString
                    Attr.DaoType = DAODataType.dbText
            End Select

            If Not Me.ContainsAttr(Attr) Then
                Me._Attrs.Add(Attr)
            End If
        End Sub

        Protected Sub SetPartNumberIdentifier(ByVal ColumnTitle As String)
            If Me.PartnoIdentifier.Contains(ColumnTitle.ToUpper) Then
                Exit Sub
            End If
            Me._PnKeyType = SideFileDbMatchKey.Property
            Me._PnKeyProp = ColumnTitle
        End Sub

        Protected Function ContainsAttr(ByVal Attr As PrtnAttr) As Boolean
            For i As Integer = 0 To Me.Attrs.Count - 1
                Dim iAttr As PrtnAttr = CType(Me.Attrs(i), PrtnAttr)
                If iAttr.Name = Attr.Name And iAttr.PdbType = Attr.PdbType Then
                    Return True
                End If
            Next
            Return False
        End Function

        Protected Sub DelArrayItems(ByRef DelIdx As ArrayList)
            Dim i As Integer
            If DelIdx.Count > 0 Then
                DelIdx.Sort()
                For i = DelIdx.Count - 1 To 0 Step -1
                    Me.RemovePart(CInt(DelIdx(i)))
                Next
            End If
            DelIdx.Clear()
        End Sub

        Protected Function MakeUniquePropListMergeValues(ByVal PropMatches As List(Of PdbProp), Optional ByVal ValueSeparator As String = "|") As List(Of PdbProp)
            Dim i, Idx As Integer, Props As New List(Of PdbProp)
            Dim ChkProp As PdbProp, UniqueProp As PdbProp

            For i = 0 To PropMatches.Count - 1
                ChkProp = CType(PropMatches(i), PdbProp)
                Idx = Me.PropExist(ChkProp.Name, Props)
                ' create unique prop list
                If Idx = -1 Then ' add property to list
                    Props.Add(ChkProp)
                Else ' property exist - merge value
                    UniqueProp = CType(Props(Idx), PdbProp)
                    UniqueProp.Value = UniqueProp.Value + ValueSeparator + ChkProp.Value
                    Props(Idx) = UniqueProp
                End If
            Next
            Return Props
        End Function

        Protected Function PropExist(ByVal PropName As String, ByVal Props As List(Of PdbProp)) As Integer
            Dim i As Integer, Prop As PdbProp
            For i = 0 To Props.Count - 1
                Prop = CType(Props(i), PdbProp)
                If PropName = Prop.Name Then Return i
            Next
            Return -1
        End Function

        Protected Sub RemovePart(ByVal Idx As Integer)
            If Idx < 0 Or Idx > Me.Parts.Count - 1 Then
                Exit Sub
            End If
            Me.Parts.RemoveAt(Idx)
        End Sub

        Protected Function RemoveEmptyArrayFields(ByVal InBuff As String()) As String()
            Dim OutBuff As String()
            ReDim OutBuff(-1)

            For Each Val As String In InBuff
                If Not String.IsNullOrEmpty(Val) Then
                    ReDim Preserve OutBuff(OutBuff.Length)
                    OutBuff(OutBuff.Length - 1) = Val
                End If
            Next

            Return OutBuff
        End Function

        Protected Function TruncatePath(ByVal InPath As String) As String
            Dim DirName As String, FileName As String
            DirName = Path.GetFileName(Path.GetDirectoryName(InPath))
            FileName = Path.GetFileName(InPath)
            Return "...\" + DirName + "\" + FileName
        End Function
#End Region

    End Class
End Namespace
