Imports System.IO
Imports System.Data
Imports System.Globalization
Imports System.Data.SQLite
Imports MGCPCBPartsEditor

<Serializable()> _
Public Enum ProgbarMode
    Init = 3
    Incr = 7
    Close = 10
End Enum

<Serializable()> _
Public Enum MsgType
    Txt
    Nte
    Wrn
    Err
End Enum

<Serializable()> _
Public Enum DAODataType
    dbInteger = 3
    dbDouble = 7
    dbText = 10
End Enum

<Serializable()> _
Public Enum DbExtension
    db3
    mdb
End Enum

<Serializable()> _
Public MustInherit Class BaseClass

#Region "Public Events"
    Public Event StatusbarMsg(ByVal Msg As String)
    Public Event ProgBarAction(ByVal Value As Integer, ByVal Mode As ProgbarMode)
    Public Event SysException(ByVal UserMsg As String, ByVal Ex As System.Exception, ByVal LeadingSpaces As Integer)
    Public Event Add2Logfile(ByVal Msg As String, ByVal LeadingSpaces As Integer, ByVal MsgType As MsgType)
    Public Event TranscriptMsg(ByVal Msg As String, ByVal LeadingSpaces As Integer, ByVal MsgType As MsgType, ByVal AddLogMsg As Boolean)

    Public Sub DxDbStatusbarMsg(ByVal Msg As String)
        RaiseEvent StatusbarMsg(Msg)
    End Sub

    Public Sub DxDbProgBarAction(ByVal Value As Integer, ByVal Mode As ProgbarMode)
        RaiseEvent ProgBarAction(Value, Mode)
    End Sub

    Public Sub DxDbSysException(ByVal UserMsg As String, ByVal Ex As System.Exception, ByVal LeadingSpaces As Integer)
        RaiseEvent SysException(UserMsg, Ex, LeadingSpaces)
    End Sub

    Public Sub DxDbAdd2Logfile(ByVal TextStr As String, Optional ByVal LeadingSpaces As Integer = 0, Optional ByVal MsgType As MsgType = MsgType.Txt)
        RaiseEvent Add2Logfile(TextStr, LeadingSpaces, MsgType)
    End Sub

    Public Sub DxDbTranscriptMsg(ByVal Msg As String, Optional ByVal LeadingSpaces As Integer = 0, Optional ByVal MsgType As MsgType = MsgType.Txt, Optional ByVal AddLogMsg As Boolean = True)
        RaiseEvent TranscriptMsg(Msg, LeadingSpaces, MsgType, AddLogMsg)
    End Sub
#End Region

#Region "XML Element Name Constants"
    Protected Const elemCObject As String = "CObject"
    Protected Const elemCConfig As String = "CConfig"
    Protected Const elemCObList As String = "CObList"
    Protected Const elemCConfigSym As String = "CConfigSym"
    Protected Const elemCConfigLib As String = "CConfigLib"
    Protected Const elemCConfigPref As String = "CConfigPref"
    Protected Const elemCConfigScripting As String = "CConfigScripting"
    Protected Const elemCConfigScriptLanguage As String = "CConfigScriptLanguage"

    Protected Const elemCObArray As String = "CObArray"
    Protected Const attrCXMLTypedPtrArraySize As String = "CXMLTypedPtrArraySize"

    Protected Const elemCConfigLibEntry As String = "CConfigLibEntry"

    Protected Const elemCConfigAtt As String = "CConfigAtt"
    Protected Const elemCConfigAttEntry As String = "CConfigAttEntry"

    Protected Const elemCConfigTable As String = "CConfigTable"
    Protected Const elemCConfigTableEntry As String = "CConfigTableEntry"
    Protected Const elemCStringList As String = "CStringList"
#End Region

#Region "MDB Tablename Constants"
    ' symbol table
    Public Const SymTableName As String = "Symbols"
    Public Const SymbolField As String = "Symbol"

    ' field names for partition table
    Public Const DeviceField As String = "Part Number"
    Public Const PartNameField As String = "Part Name"
    Public Const PartLabelField As String = "Part Label"
    Public Const RefDesField As String = "RefPreFix"
    Public Const CellNameField As String = "Cell Name"
    Public Const CellPinsField As String = "Cell Pins"
    Public Const DescriptField As String = "Desc"
    Public Const TypeField As String = "Type"

    Public Const ValueProp As String = "Value"
    Public Const DataSheetProp As String = "Datasheet"
    Public Const IbisModelProp As String = "IBIS"
    Public Const M3DModelProp As String = "3D_Model"

    Public EEFields As List(Of String)
#End Region

#Region "Constructor"
    Public Sub New(ComVersion As Integer)
        'Public EEFields As String() = {DeviceField, PartNameField, PartLabelField, RefDesField _
        '                             , CellNameField, CellPinsField, DescriptField, TypeField}
        Me.EEFields = New List(Of String) From {
            BaseClass.DeviceField,
            BaseClass.PartNameField,
            BaseClass.PartLabelField,
            BaseClass.RefDesField,
            BaseClass.CellNameField,
            BaseClass.CellPinsField,
            BaseClass.DescriptField,
            BaseClass.TypeField
        }
    End Sub
#End Region

End Class

<Serializable()> _
Public Class PdbPart
    Public PartNo As String
    Public RefPref As String
    Public Name As String
    Public Label As String
    Public Desc As String
    Public Type As String
    Public SymCount As Integer
    Public Symbs As List(Of PdbSymb)
    Public Props As List(Of PdbProp)
    Public Cells As List(Of PdbCell)

    Public Sub New(ByVal PartNo As String)
        Me.PartNo = PartNo
        Me.RefPref = ""
        Me.Name = ""
        Me.Label = ""
        Me.Desc = ""
        Me.Type = ""
        Me.SymCount = 0
        Me.Symbs = New List(Of PdbSymb)
        Me.Props = New List(Of PdbProp)
        Me.Cells = New List(Of PdbCell)
    End Sub

    Public Function SymbolExist(ByVal SymName As String) As Boolean
        If Me.Symbs.Count = 0 Then Return False
        For Each Sym As PdbSymb In Me.Symbs
            If Sym.Name.ToLower = SymName.ToLower Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Function GetDefaultCell() As String
        'check Top Cell
        For Each Cell As PdbCell In Me.Cells
            If Cell.Type = EPDBCellReferenceType.epdbCellRefTop Then
                Return Cell.Name
            End If
        Next

        'check Bottom Cell
        For Each Cell As PdbCell In Me.Cells
            If Cell.Type = EPDBCellReferenceType.epdbCellRefBottom Then
                Return Cell.Name
            End If
        Next

        'check Alternate Cell
        For Each Cell As PdbCell In Me.Cells
            If Cell.Type = EPDBCellReferenceType.epdbCellRefAlternate Then
                Return Cell.Name
            End If
        Next

        Return String.Empty
    End Function

    Public Function GetDefaultCellPinCount() As String
        Dim DefCell As String = Me.GetDefaultCell()
        If DefCell = String.Empty Then Return "0" ' String.Empty

        Dim oCell As PdbCell = Nothing
        For Each oCell In Me.Cells
            If oCell.Name = DefCell Then
                Exit For
            End If
        Next

        If Not IsNothing(oCell) Then
            Return oCell.PinNumbers.Count.ToString
        End If

        Return "0"
    End Function
End Class

<Serializable()> _
Public Class PdbProp
    Public Unit As MGCPCBPartsEditor.EPDBUnit
    Public Name As String
    Public Value As String
    Public PdbType As EPDBPropertyType
    Public Index As Integer
    Public PropEx As Exception

    Public Sub New()
        Me.Name = ""
        Me.Value = ""
        Me.PdbType = EPDBPropertyType.epdbPropTypeString
        Me.Index = -1
        Me.PropEx = Nothing
    End Sub

    Public Sub New(MgcProp As String, Type As EPDBPropertyType)
        Me.Unit = EPDBUnit.epdbUnitCurrent
        Me.Name = MgcProp.Trim()
        Me.PdbType = Type
        Me.Index = -1
        Me.PropEx = Nothing
        Me.Value = String.Empty
    End Sub

    Public Sub New(MgcProp As MGCPCBPartsEditor.Property, PrpUnits As EPDBUnit)
        Me.Unit = PrpUnits
        Me.Name = MgcProp.Name.Trim()
        Me.PdbType = MgcProp.Type
        Me.Index = -1
        Me.PropEx = Nothing

        Try
            'set property type
            Select Case Me.PdbType
                Case EPDBPropertyType.epdbPropTypeInt
                    Me.Value = MgcProp.Value(Me.Unit).ToString
                Case EPDBPropertyType.epdbPropTypeReal
                    Me.Value = MgcProp.Value(Me.Unit).ToString
                    Me.NormalizePropVal()
                Case EPDBPropertyType.epdbPropTypeString
                    Me.Value = MgcProp.GetValueString()
            End Select
        Catch ex As Exception
            Me.PropEx = New Exception("Property Name: '" + Me.Name + "' !error processing property value!")
        End Try
    End Sub

    Private Sub NormalizePropVal()
        Dim Value As Double
        If Not IsNumeric(Me.Value) Then
            If Normalize.ProcessValue(Me.Value, Value) Then
                Me.Value = CStr(Value)
            End If
        End If
    End Sub
End Class

<Serializable()> _
Public Class SqlField
    Public Name As String
    Public DbType As TypeAffinity
    Public SqlType As String
    Public SysType As System.Type

    Public Sub New(fName As String, fPdbType As EPDBPropertyType)
        Me.Name = fName

        Me.DbType = TypeAffinity.Text
        Me.SqlType = "TEXT"
        Me.SysType = System.Type.GetType("System.String")
        'Exit Sub

        Select Case fPdbType
            Case EPDBPropertyType.epdbPropTypeInt
                Me.DbType = TypeAffinity.Int64
                'Me.SqlType = "INTEGER"
                Me.SqlType = "NUMERIC"
                Me.SysType = System.Type.GetType("System.Int64")
            Case EPDBPropertyType.epdbPropTypeReal
                Me.DbType = TypeAffinity.Double
                'Me.SqlType = "REAL"
                Me.SqlType = "NUMERIC"
                Me.SysType = System.Type.GetType("System.Double")
            Case EPDBPropertyType.epdbPropTypeString
                Me.DbType = TypeAffinity.Text
                Me.SqlType = "TEXT"
                Me.SysType = System.Type.GetType("System.String")
        End Select
    End Sub
End Class

<Serializable()> _
Public Class PdbSymb
    Public Name As String
    Public PinNames As List(Of String)

    Public Sub New(SymRef As MGCPCBPartsEditor.SymbolReference)
        Me.Name = SymRef.Name
        Me.PinNames = New List(Of String)
        For Each p As String In CType(SymRef.PinNames, Object())
            Me.PinNames.Add(p)
        Next
    End Sub
End Class

<Serializable()> _
Public Class PdbCell
    Public Name As String
    Public Type As MGCPCBPartsEditor.EPDBCellReferenceType
    Public PinNumbers As List(Of String)

    Public Sub New(CelRef As MGCPCBPartsEditor.CellReference)
        Me.Name = CelRef.Name
        Me.Type = CelRef.Type
        Me.PinNumbers = New List(Of String)
        For Each p As String In CType(CelRef.PinNumbers, Object())
            Me.PinNumbers.Add(p)
        Next
    End Sub
End Class

<Serializable()> _
Public Class PrtnAttr
    Public Name As String
    Public PdbType As EPDBPropertyType
    Public DaoType As DAODataType
    Public Index As Integer

    Public Sub New()
        Me.Name = ""
        Me.PdbType = EPDBPropertyType.epdbPropTypeString
        Me.DaoType = DAODataType.dbText
        Me.Index = -1
    End Sub

    Public Sub New(ByVal AttrName As String)
        Me.Name = AttrName.Trim()
        Me.PdbType = EPDBPropertyType.epdbPropTypeString
        Me.DaoType = DAODataType.dbText
        Me.Index = -1
    End Sub
End Class

<Serializable()> _
Public Class Normalize
    Public Shared Function ProcessValue(ByVal ValStr As String, ByRef ValNum As Double) As Boolean
        Dim Val As String = ""
        Dim Magnifier As String = ""

        ' Dezimaltrennzeichen
        Dim sDecChar As String = Normalize.GetDecimalChar(ValStr)

        ValNum = 0

        If Not Normalize.IsValid(ValStr, sDecChar) Then Return False

        For i As Integer = 0 To Len(ValStr) - 1
            If IsNumeric(ValStr.Substring(i, 1)) Then
                Val += ValStr.Substring(i, 1)
            ElseIf ValStr.Substring(i, 1) = sDecChar Then
                Val += String.Format("{0:0.0}", 0).Chars(1)
            Else
                If IsValidMagnifier(ValStr.Substring(i, 1)) AndAlso Magnifier = "" Then
                    Magnifier = ValStr.Substring(i, 1)
                    If IsTrailingMagnifierNum(i, ValStr) Then
                        Val += CStr(sDecChar)
                    End If
                End If
            End If
        Next

        If Not IsValidMagnifier(Magnifier) Then Return False
        If Not IsNumeric(Val) Then Return False
        ValNum = CDbl(Val) * Normalize.GetMagnifierValue(Magnifier)
        Return True
    End Function

    Private Shared Function IsValid(ByVal ValStr As String, ByVal sDecChar As String) As Boolean
        Dim NumCount As Integer, Digit As String, IsMagnifier As Boolean

        If ValStr = "" Then Return False
        Digit = ValStr.Substring(0, 1)
        If Not IsNumeric(Digit) And Not Digit = sDecChar Then Return False

        For i As Integer = 0 To Len(ValStr) - 1
            Digit = CChar(ValStr.Substring(i, 1))
            Select Case Digit
                Case sDecChar
                    NumCount += 1
                Case Else
                    If IsNumeric(ValStr.Substring(i, 1)) Then
                        NumCount += 1
                    Else
                        If Not IsMagnifier And IsValidMagnifier(CStr(Digit)) Then IsMagnifier = True
                    End If
            End Select
        Next

        Return True
    End Function

    Private Shared Function GetMagnifierValue(ByVal Magnifier As String) As Double
        Select Case Magnifier
            Case "f" : Return 0.000000000000001 ' 1e-15;
            Case "p", "P" : Return 0.000000000001    ' 1e-12
            Case "n", "N" : Return 0.000000001       ' 1e-9
            Case "u", "U" : Return 0.000001          ' 1e-6
            Case "m" : Return 0.001                  ' 1e-3
            Case "k", "K" : Return 1000              ' 1e3
            Case "M" : Return 1000000                ' 1e6
            Case "g", "G" : Return 1000000000        ' 1e9
            Case "t", "T" : Return 1000000000000     ' 1e12
            Case "%" : Return 0.01
            Case "R", "F" : Return 1.0
            Case Else : Return 1.0
        End Select
    End Function

    Private Shared Function IsValidMagnifier(ByVal Magnifier As String) As Boolean
        Select Case Magnifier
            Case "", "f", "F", "p", "P", "n", "N", "u", "U", "m", "k", "K", "M", "g", "G", "t", "T", "%"
                Return True
            Case Else
                Return False
        End Select
    End Function

    Private Shared Function IsTrailingMagnifierNum(ByVal StartIdx As Integer, ByVal ValStr As String) As Boolean
        If StartIdx = Len(ValStr) - 1 Then Return False
        For i As Integer = StartIdx + 1 To Len(ValStr) - 1
            If IsNumeric(ValStr.Substring(i, 1)) Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Shared Function GetDecimalChar(ByVal ValStr As String) As String
        Dim sDecChar As Char = String.Format("{0:0.0}", 0).Chars(1)
        For i As Integer = Len(ValStr) - 1 To 0 Step -1
            If ValStr.Substring(i, 1) = "." Or ValStr.Substring(i, 1) = "," Then
                Return ValStr.Substring(i, 1)
            End If
        Next
        Return ""
    End Function
End Class
