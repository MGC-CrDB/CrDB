Imports System.IO
Imports System.Threading
Imports MGCPCBPartsEditor
Imports System.Globalization
Imports Microsoft.VisualBasic.FileIO
Imports System.Runtime.Serialization.Formatters.Binary

Namespace DxDB
    Public Class PdbPrtn
        Inherits BaseClass

#Region "Members"
        Public ReadOnly Name As String
        Protected _Parts As List(Of PdbPart)
        Protected _Attrs As List(Of PrtnAttr)
#End Region

#Region "Properties"
        Public Property Part(ByVal p As Integer) As PdbPart
            Get
                Return CType(Me._Parts(p), PdbPart)
            End Get
            Set(ByVal value As PdbPart)
                Me._Parts(p) = value
            End Set
        End Property

        Public ReadOnly Property Parts() As List(Of PdbPart)
            Get
                Return Me._Parts
            End Get
        End Property

        Public ReadOnly Property Attrs() As List(Of PrtnAttr)
            Get
                Return Me._Attrs
            End Get
        End Property
#End Region

#Region "Constructors"
        Public Sub New(nComVersion As Integer)
            MyBase.New(nComVersion)
            Me.Name = String.Empty
            Me._Parts = New List(Of PdbPart)
            Me._Attrs = New List(Of PrtnAttr)
        End Sub

        Public Sub New(PrtnName As String, nComVersion As Integer)
            MyBase.New(nComVersion)
            Me.Name = PrtnName
            Me._Parts = New List(Of PdbPart)
            Me._Attrs = New List(Of PrtnAttr)
        End Sub
#End Region

#Region "Partition Public Methods"
        Public Sub ReadPdbPrtn(ByVal MgcPrtn As Partition, ByVal LibDir As String, ByVal PdbUnits As MGCPCBPartsEditor.EPDBUnit, UseSymbolTable As Boolean, UseCellPins As Boolean)
            Dim i As Integer, len As Integer, DatashSubstr As String
            Dim PartNo As PdbPart, PartProp As PdbProp
            Dim CelRef As MGCPCBPartsEditor.CellReference
            Dim CelRefs As MGCPCBPartsEditor.CellReferences
            Dim SymRef As MGCPCBPartsEditor.SymbolReference
            Dim SymRefs As MGCPCBPartsEditor.SymbolReferences
            Dim PdbPart As MGCPCBPartsEditor.Part, PnNum As Integer
            Dim MgcProp As MGCPCBPartsEditor.Property
            Dim Cell As PdbCell, Symb As PdbSymb

            If IsNothing(MgcPrtn) Then Exit Sub

            '-------------------------------
            ' create mandatory system fields
            For Each EEField As String In Me.EEFields
                If UseCellPins AndAlso EEField = BaseClass.CellPinsField Then
                    Me.AddDefAttr(EEField, EPDBPropertyType.epdbPropTypeInt, DAODataType.dbInteger) ' FieldName
                    Continue For
                End If
                Me.AddDefAttr(EEField) ' FieldName
            Next

            Try
                '-------------------------------
                ' read props for each partition
                Me.DxDbStatusbarMsg("Reading Partition '" + MgcPrtn.Name + "'")
                Me.DxDbTranscriptMsg("... reading Partition '" + MgcPrtn.Name + "' - Partnumber(s): " + CStr(MgcPrtn.Parts.Count), 10)
                Me.DxDbProgBarAction(MgcPrtn.Parts.Count, ProgbarMode.Init)
                PnNum = 0
                For Each PdbPart In MgcPrtn.Parts()
                    PnNum += 1
                    Me.DxDbProgBarAction(0, ProgbarMode.Incr)
                    Me.DxDbStatusbarMsg("(" & PnNum & "," & MgcPrtn.Parts.Count & ") " & PdbPart.Number)

                    ' create new part
                    PartNo = New PdbPart(PdbPart.Number)

                    ' add ref_prefix
                    PartNo.RefPref = PdbPart.RefDesPrefix

                    ' add part name
                    PartNo.Name = PdbPart.Name

                    ' add part label
                    PartNo.Label = PdbPart.Label

                    ' add part description
                    PartNo.Desc = PdbPart.Description

                    ' add part type
                    PartNo.Type = PdbPart.TypeString

                    ' add top, bottom, alternate cell(s)
                    CelRefs = PdbPart.CellReferences(EPDBCellReferenceType.epdbCellRefAll)
                    For Each CelRef In CelRefs
                        Cell = New PdbCell(CelRef)
                        PartNo.Cells.Add(Cell)
                    Next

                    ' add symbol(s)
                    SymRefs = PdbPart.SymbolReferences()
                    For Each SymRef In SymRefs
                        ' increment symbol counter
                        PartNo.SymCount += 1
                        If Not PartNo.SymbolExist(SymRef.Name) Then
                            Symb = New PdbSymb(SymRef)
                            PartNo.Symbs.Add(Symb)
                        End If
                    Next

                    ' add properties
                    Try
                        For Each MgcProp In PdbPart.Properties
                            PartProp = New PdbProp(MgcProp, PdbUnits)

                            If Not IsNothing(PartProp.PropEx) Then
                                Dim UserMsg As String = "ReadPdbPrtn(), PartNumber: " + PdbPart.Number
                                Me.ProcessMgcPdbExecption(PartProp.PropEx, UserMsg)
                                Continue For
                            End If

                            If UCase(PartProp.Name) = UCase(BaseClass.M3DModelProp) _
                            Or UCase(PartProp.Name) = UCase(BaseClass.IbisModelProp) Then
                                PartProp.Value = LibDir + "\" + PartProp.Value
                                PartProp.Value = Me.ReplaceEx(PartProp.Value, "\\", "\")
                            End If

                            ' process Central Library Datasheets folder 
                            If UCase(PartProp.Name) = UCase(BaseClass.DataSheetProp) Then
                                len = PartProp.Value.Length
                                If len > 10 Then
                                    DatashSubstr = PartProp.Value.Substring(0, 10).ToLower
                                    If DatashSubstr = "datasheets" Then
                                        PartProp.Value = LibDir + "\" + PartProp.Value
                                        PartProp.Value = Me.ReplaceEx(PartProp.Value, "\\", "\")
                                    End If
                                End If
                            End If

                            If Not PartProp.Value = String.Empty Then
                                PartNo.Props.Add(PartProp)
                                Me.AddPrtnAttr(PartProp)
                            End If
                        Next

                        If UseSymbolTable Then
                            PartProp = New PdbProp(BaseClass.SymbolField, EPDBPropertyType.epdbPropTypeString)
                            Me.AddPrtnAttr(PartProp)
                        End If

                    Catch ex As Exception
                        Dim UserMsg As String = "ReadPdbPrtn(), PartNumber: " + PdbPart.Number + " while parsing Property List!"
                        Me.ProcessMgcPdbExecption(ex, UserMsg)
                    End Try

                    ' add PdbPart to Partition
                    Me.Parts.Add(PartNo)
                Next

                Me.DxDbProgBarAction(0, ProgbarMode.Close)
                Me.DxDbStatusbarMsg("Reading Partition done.")
            Catch ex As System.Exception
                Dim UserMsg As String = "ReadPdbPrtn(): " + MgcPrtn.Name
                Me.ProcessMgcPdbExecption(ex, UserMsg)
            End Try
        End Sub

        Public Sub AddSideFileProps2Parts(ByRef AsciiPropsDb As SideFileDB)
            Dim PartNo As PdbPart
            Dim i As Integer, n As Integer
            Dim SideFileProps As New List(Of PdbProp)

            For i = 0 To Me.Parts.Count - 1
                PartNo = Me.Part(i)
                ' add record for each partnumber
                ' ------- this is Maquet stuff ArtNr --------
                If AsciiPropsDb.PnKeyType = SideFileDB.SideFileDbMatchKey.Property Then
                    If AsciiPropsDb.FindPartByProperty(PartNo, SideFileProps) Then
                        Me.DxDbAdd2Logfile("Property match found (" + AsciiPropsDb.PnKeyProp + "): " + Me.Part(i).PartNo, 15)
                        For n = 0 To SideFileProps.Count - 1
                            PartNo.Props.Add(SideFileProps(n))
                            Me.AddProp2Attrs(CType(SideFileProps(n), PdbProp))
                        Next
                    End If
                Else
                    If AsciiPropsDb.FindPartByPartnumber(Me.Part(i).PartNo, SideFileProps) Then
                        Me.DxDbAdd2Logfile("PartNumber match found: " + Me.Part(i).PartNo, 15)
                        For n = 0 To SideFileProps.Count - 1
                            PartNo.Props.Add(SideFileProps(n))
                            Me.AddProp2Attrs(CType(SideFileProps(n), PdbProp))
                        Next
                    End If
                End If
                Me.Part(i) = PartNo
                Me.DxDbStatusbarMsg("(" & i + 1 & "," & Me.Parts.Count & ") " & Me.Part(i).PartNo)
            Next
        End Sub

#End Region

#Region "Read PDB Protected Methods"
        Private Function ChkSymbolExist(ByVal SymName As String, ByVal Symbols As List(Of String)) As Boolean
            If Symbols.Count = 0 Then Return False
            For i As Integer = 0 To Symbols.Count - 1
                If Symbols(i).ToString = SymName Then Return True
            Next
            Return False
        End Function

        Private Function ReplaceEx(ByVal original As String, ByVal pattern As String, ByVal replacement As String) As String
            Dim count, position0, position1 As Integer

            Dim upperString As String = original.ToUpper()
            Dim upperPattern As String = pattern.ToUpper()

            Dim inc As Integer = CInt((original.Length / pattern.Length) * (replacement.Length - pattern.Length))
            Dim Chars((original.Length + Math.Max(0, inc))) As Char

            position1 = upperString.IndexOf(upperPattern, position0)
            While position1 > -1
                For i As Integer = position0 To position1 - 1
                    Chars(count) = original(i)
                    count = count + 1
                Next
                For i As Integer = 0 To replacement.Length - 1
                    Chars(count) = replacement(i)
                    count = count + 1
                Next
                position0 = position1 + pattern.Length
                position1 = upperString.IndexOf(upperPattern, position0)
            End While

            If position0 = 0 Then Return original
            For i As Integer = position0 To original.Length - 1
                Chars(count) = original(i)
                count = count + 1
            Next
            Return New String(Chars, 0, count)
        End Function

        Private Sub AddDefAttr(AttrName As String, Optional PdbType As EPDBPropertyType = EPDBPropertyType.epdbPropTypeString, Optional DaoType As DAODataType = DAODataType.dbText)
            Dim Attr As New PrtnAttr

            Attr.Name = AttrName
            Attr.PdbType = PdbType
            Attr.DaoType = DaoType

            If Not Me.ContainsAttr(Attr) Then
                Me._Attrs.Add(Attr)
            End If
        End Sub

        Private Sub AddPrtnAttr(Prop As PdbProp)
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

        Private Function ContainsAttr(ByVal Attr As PrtnAttr) As Boolean
            For i As Integer = 0 To Me.Attrs.Count - 1
                Dim iAttr As PrtnAttr = CType(Me.Attrs(i), PrtnAttr)
                If iAttr.Name.ToLower = Attr.Name.ToLower And iAttr.PdbType = Attr.PdbType Then
                    Return True
                End If
            Next
            Return False
        End Function
#End Region

#Region "Read additional Property File Protected Methods"
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

        Protected Sub AddProp2Attrs(ByVal Prop As PdbProp)
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
#End Region

#Region "Message & Error Handling"
        Protected Sub ProcessMgcPdbExecption(ByVal SysErr As System.Exception, ByVal UserMsg As String)
            Dim PDBErrCode As Integer
            Dim ErrNum As String
            Dim ErrVect As String(), ErrStr As String

            ErrStr = SysErr.Message
            ErrVect = Split(ErrStr, ":")

            If UBound(ErrVect) = 0 Then
                Me.DxDbSysException(UserMsg, SysErr, 5)
                Exit Sub
            End If

            Try
                ErrNum = ErrVect(1).ToString.Trim.Substring(2)
                PDBErrCode = Integer.Parse(ErrNum, System.Globalization.NumberStyles.AllowHexSpecifier)
            Catch ex As Exception
                Me.DxDbSysException(UserMsg, SysErr, 5)
                Exit Sub
            End Try

            Select Case PDBErrCode
                Case EPDBErrCode.epdbErrCodeBadObject
                    ErrStr = "BadObject(" + ErrVect(1).ToString.Trim + "): Invalid Object. Validation of the object failed."
                Case EPDBErrCode.epdbErrCodeCannotModifyPartsInReservedPartition
                    ErrStr = "CannotModifyPartsInReservedPartition(" + ErrVect(1).ToString.Trim + "): Partition containing the part is reserved."
                Case EPDBErrCode.epdbErrCodeCannotOpenDatabase
                    ErrStr = "CannotOpenDatabase(" + ErrVect(1).ToString.Trim + "): Cannot close the database if automation did not open the database. This occurs if an application, such as Expedition PCB, opens the database."
                Case EPDBErrCode.epdbErrCodeCannotOpenDatabaseForWriting
                    ErrStr = "CannotOpenDatabaseForWriting(" + ErrVect(1).ToString.Trim + "): Cannot open the database for writing when the Part Editor is opened in standalone mode."
                Case EPDBErrCode.epdbErrCodeCannotOpenExternalDatabase
                    ErrStr = "CannotOpenExternalDatabase(" + ErrVect(1).ToString.Trim + "): Cannot open the database when the Part Editor is opened by an application, such as Expedition PCB. Instead use the ActiveDatabaseEx Method."
                Case EPDBErrCode.epdbErrCodeCannotSetPinDefinition
                    ErrStr = "CannotSetPinDefinition(" + ErrVect(1).ToString.Trim + "): Cannot Set PinDefinition."
                Case EPDBErrCode.epdbErrCodeCannotSetPinDefinitionsNotSameType
                    ErrStr = "CannotSetPinDefinitionsNotSameType(" + ErrVect(1).ToString.Trim + "): Cannot Set PinDefinitions Not SameType."
                Case EPDBErrCode.epdbErrCodeCantAccessPartitions
                    ErrStr = "CantAccessPartitions(" + ErrVect(1).ToString.Trim + "): Cant Access Partitions."
                Case EPDBErrCode.epdbErrCodeCantCloseDatabase
                    ErrStr = "CantCloseDatabase(" + ErrVect(1).ToString.Trim + "): Cannot close the database if automation did not open the database. This occurs if an application, such as Expedition PCB, opens the database."
                Case EPDBErrCode.epdbErrCodeCantSetToCurrent
                    ErrStr = "CantSetToCurrent(" + ErrVect(1).ToString.Trim + "): Cannot set the default unit to epdbCurrentUnit."
                Case EPDBErrCode.epdbErrCodeCantSetVisible
                    ErrStr = "CantSetVisible(" + ErrVect(1).ToString.Trim + "): Cannot set to visible unless a database is currently open."
                Case EPDBErrCode.epdbErrCodeCentralLibraryIncorrectOrNotMigrated
                    ErrStr = "CentralLibraryIncorrectOrNotMigrated(" + ErrVect(1).ToString.Trim + "): Central Library Incorrect Or Not Migrated."
                Case EPDBErrCode.epdbErrCodeDBAlreadyLocked
                    ErrStr = "DBAlreadyLocked(" + ErrVect(1).ToString.Trim + "): Database is already locked."
                Case EPDBErrCode.epdbErrCodeDBLockedByAnotherClient
                    ErrStr = "DBLockedByAnotherClient(" + ErrVect(1).ToString.Trim + "): Another automation client has already locked the database."
                Case EPDBErrCode.epdbErrCodeDBNotLocked
                    ErrStr = "DBNotLocked(" + ErrVect(1).ToString.Trim + "): Database is not locked."
                Case EPDBErrCode.epdbErrCodeDBNotOpen
                    ErrStr = "DBNotOpen(" + ErrVect(1).ToString.Trim + "): Database is not open."
                Case EPDBErrCode.epdbErrCodeDuplicatedPinNumbers
                    ErrStr = "DuplicatedPinNumbers(" + ErrVect(1).ToString.Trim + "): Pin number used more than once for a slot."
                Case EPDBErrCode.epdbErrCodeFileNotFound
                    ErrStr = "FileNotFound(" + ErrVect(1).ToString.Trim + "): File not found."
                Case EPDBErrCode.epdbErrCodeGateNameIsNotUnique
                    ErrStr = "GateNameIsNotUnique(" + ErrVect(1).ToString.Trim + "): Gate Name Is Not Unique."
                Case EPDBErrCode.epdbErrCodeGateObjectIsInvalid
                    ErrStr = "GateObjectIsInvalid(" + ErrVect(1).ToString.Trim + "): Gate Object Is Invalid."
                Case EPDBErrCode.epdbErrCodeIncompletePartsSaveFailed
                    ErrStr = "IncompletePartsSaveFailed(" + ErrVect(1).ToString.Trim + "): Incomplete Parts Save Failed."
                Case EPDBErrCode.epdbErrCodeIncorrectCellType
                    ErrStr = "IncorrectCellType(" + ErrVect(1).ToString.Trim + "): Incorrect CellType."
                Case EPDBErrCode.epdbErrCodeIncorrectNumberOfPins
                    ErrStr = "IncorrectNumberOfPins(" + ErrVect(1).ToString.Trim + "): Incorrect Number Of Pins."
                Case EPDBErrCode.epdbErrCodeIncorrectPinCountForGate
                    ErrStr = "IncorrectPinCountForGate(" + ErrVect(1).ToString.Trim + "): Incorrect Pin Count For Gate."
                Case EPDBErrCode.epdbErrCodeIncorrectPinValueType
                    ErrStr = "IncorrectPinValueType(" + ErrVect(1).ToString.Trim + "): Incorrect Pin Value Type."
                Case EPDBErrCode.epdbErrCodeIndexIsOutOfRange
                    ErrStr = "IndexIsOutOfRange(" + ErrVect(1).ToString.Trim + "): Index Is Out Of Range."
                Case EPDBErrCode.epdbErrCodeInternalError
                    ErrStr = "InternalError(" + ErrVect(1).ToString.Trim + "): Internal Error."
                Case EPDBErrCode.epdbErrCodeInvalidAutomationBasicLicense
                    ErrStr = "InvalidAutomationBasicLicense(" + ErrVect(1).ToString.Trim + "): Cannot acquire an Automation Basic license."
                Case EPDBErrCode.epdbErrCodeInvalidParameter
                    ErrStr = "InvalidParameter(" + ErrVect(1).ToString.Trim + "): Invalid parameter(s)."
                Case EPDBErrCode.epdbErrCodeInvalidLibraryManagerLicense
                    ErrStr = "InvalidLibraryManagerLicense(" + ErrVect(1).ToString.Trim + "): Invalid LibraryManager License."
                Case EPDBErrCode.epdbErrCodeInvalidPartType
                    ErrStr = "InvalidPartType(" + ErrVect(1).ToString.Trim + "): Invalid PartType."
                Case EPDBErrCode.epdbErrCodeLicenseStillRequired
                    ErrStr = "LicenseStillRequired(" + ErrVect(1).ToString.Trim + "): License Still Required."
                Case EPDBErrCode.epdbErrCodeMappingsSaveFailed
                    ErrStr = "MappingsSaveFailed(" + ErrVect(1).ToString.Trim + "): Mappings Save Failed."
                Case EPDBErrCode.epdbErrCodeOperationAllowedForLogicalGatesOnly
                    ErrStr = "OperationAllowedForLogicalGatesOnly(" + ErrVect(1).ToString.Trim + "): Operation Allowed For Logical Gates Only."
                Case EPDBErrCode.epdbErrCodeOperationForbiddenForNewPart
                    ErrStr = "OperationForbiddenForNewPart(" + ErrVect(1).ToString.Trim + "): Operation Forbidden For New Part."
                Case EPDBErrCode.epdbErrCodeOperationForbiddenForThisSymbolReference
                    ErrStr = "OperationForbiddenForThisSymbolReference(" + ErrVect(1).ToString.Trim + "): Operation Forbidden For This SymbolReference."
                Case EPDBErrCode.epdbErrCodePartAlreadyCommitted
                    ErrStr = "PartAlreadyCommitted(" + ErrVect(1).ToString.Trim + "): Part is already saved to the database."
                Case EPDBErrCode.epdbErrCodePartDescriptionIsInvalid
                    ErrStr = "PartDescriptionIsInvalid(" + ErrVect(1).ToString.Trim + "): Part description is invalid."
                Case EPDBErrCode.epdbErrCodePartIsNoLongerComplete
                    ErrStr = "PartIsNoLongerComplete(" + ErrVect(1).ToString.Trim + "): Part was complete but changes occured in the pin mappings which changed the status of the part to incomplete. Possible reasons include: slot is not complete (pin name or pin number empty), pin name is not used, pin number is not used."
                Case EPDBErrCode.epdbErrCodePartitionAlreadyExists
                    ErrStr = "PartitionAlreadyExists(" + ErrVect(1).ToString.Trim + "): Partition with the specified name already exists in the Parts Database."
                Case EPDBErrCode.epdbErrCodePartitionIsNotEmpty
                    ErrStr = "PartitionIsNotEmpty(" + ErrVect(1).ToString.Trim + "): Partion is not empty."
                Case EPDBErrCode.epdbErrCodePartitionIsReserved
                    ErrStr = "PartitionIsReserved(" + ErrVect(1).ToString.Trim + "): Partition is reserved by another user."
                Case EPDBErrCode.epdbErrCodePartitionNameIsInvalid
                    ErrStr = "PartitionNameIsInvalid(" + ErrVect(1).ToString.Trim + "): Partition name is invalid."
                Case EPDBErrCode.epdbErrCodePartitionOperationFailed
                    ErrStr = "PartitionOperationFailed(" + ErrVect(1).ToString.Trim + "): Partition creation failed for any other reason."
                Case EPDBErrCode.epdbErrCodePartitionOperationForbiddenInPCBMode
                    ErrStr = "PartitionOperationForbiddenInPCBMode(" + ErrVect(1).ToString.Trim + "): Partition Operation Forbidden In PCB Mode."
                Case EPDBErrCode.epdbErrCodePartLabelIsInvalid
                    ErrStr = "PartLabelIsInvalid(" + ErrVect(1).ToString.Trim + "): Part Label is invalid."
                Case EPDBErrCode.epdbErrCodePartNameIsInvalid
                    ErrStr = "PartNameIsInvalid(" + ErrVect(1).ToString.Trim + "): Part name is invalid."
                Case EPDBErrCode.epdbErrCodePartNumberIsInvalid
                    ErrStr = "PartNumberIsInvalid(" + ErrVect(1).ToString.Trim + "): Part number is invalid."
                Case EPDBErrCode.epdbErrCodePartNumberIsNotUniqueInLMC
                    ErrStr = "PartNumberIsNotUniqueInLMC(" + ErrVect(1).ToString.Trim + "): Part number is not unique in the Central Library."
                Case EPDBErrCode.epdbErrCodePartNumberIsNotUniqueInPDB
                    ErrStr = "PartNumberIsNotUniqueInPDB(" + ErrVect(1).ToString.Trim + "): Part number is not unique in the local library."
                Case EPDBErrCode.epdbErrCodePartRefDesPrefixIsInvalid
                    ErrStr = "PartRefDesPrefixIsInvalid(" + ErrVect(1).ToString.Trim + "): Reference designator prefix is invalid."
                Case EPDBErrCode.epdbErrCodePartUsedInRB
                    ErrStr = "PartUsedInRB(" + ErrVect(1).ToString.Trim + "): Invalid operation. Part is used in a reuseable block."
                Case EPDBErrCode.epdbErrCodePartUsedInTheDesign
                    ErrStr = "PartUsedInTheDesign(" + ErrVect(1).ToString.Trim + "): Invalid operation. Part is used in the design."
                Case EPDBErrCode.epdbErrCodePinNumbersMismatch
                    ErrStr = "PinNumbersMismatch(" + ErrVect(1).ToString.Trim + "): Cell reference do not have the same set of pin number. All cell references must have the same pin numbers."
                Case EPDBErrCode.epdbErrCodePinNumberUsedInLogicalAndPhysicalGates
                    ErrStr = "PinNumberUsedInLogicalAndPhysicalGates(" + ErrVect(1).ToString.Trim + "): PinNumber Used In Logical And Physical Gates."
                Case EPDBErrCode.epdbErrCodePinValueTypeIncorrectForThisGateType
                    ErrStr = "PinValueTypeIncorrectForThisGateType(" + ErrVect(1).ToString.Trim + "): Pin Value Type Incorrect For This Gate Type."
                Case EPDBErrCode.epdbErrCodePropertyNotExists
                    ErrStr = "PropertyNotExists(" + ErrVect(1).ToString.Trim + "): Property is not defined as a common property by Library Manager."
                Case EPDBErrCode.epdbErrCodePropertyValueIsInvalid
                    ErrStr = "PropertyValueIsInvalid(" + ErrVect(1).ToString.Trim + "): Property Value Is Invalid."
                Case EPDBErrCode.epdbErrCodeReadOnlyMode
                    ErrStr = "ReadOnlyMode(" + ErrVect(1).ToString.Trim + "): Database is in read-only mode."
                Case EPDBErrCode.epdbErrCodeSaveFailed
                    ErrStr = "SaveFailed(" + ErrVect(1).ToString.Trim + "): Save Failed."
                Case EPDBErrCode.epdbErrCodeSwapIdentifierTooLong
                    ErrStr = "SwapIdentifierTooLong(" + ErrVect(1).ToString.Trim + "): Swap Identifier Too Long."
                Case EPDBErrCode.epdbErrCodeSymbolReferenceAlreadyExists
                    ErrStr = "SymbolReferenceAlreadyExists(" + ErrVect(1).ToString.Trim + "): Symbol Reference Already Exists."
                Case EPDBErrCode.epdbErrCodeSymbolReferenceObjectIsInvalid
                    ErrStr = "SymbolReferenceObjectIsInvalid(" + ErrVect(1).ToString.Trim + "): Symbol Reference Object Is Invalid."
                Case EPDBErrCode.epdbErrCodeUnableToReserveThePartition
                    ErrStr = "UnableToReserveThePartition(" + ErrVect(1).ToString.Trim + "): Partition with the specified name is already in use in the database."
                Case EPDBErrCode.epdbErrCodeUndefinedPinNameUsed
                    ErrStr = "UndefinedPinNameUsed(" + ErrVect(1).ToString.Trim + "): Slot associated with a symbol reference contains an incorrect pin name. Pin name must be defined in the associated symbol reference."
                Case EPDBErrCode.epdbErrCodeUndefinedPinNumberUsed
                    ErrStr = "UndefinedPinNumberUsed(" + ErrVect(1).ToString.Trim + "): Undefined PinNumber used."
                Case EPDBErrCode.epdbErrCodeUnknownPartType
                    ErrStr = "UnknownPartType(" + ErrVect(1).ToString.Trim + "): Unknown part type."
                Case EPDBErrCode.epdbErrCodeUnknownPropertyType
                    ErrStr = "UnknownPropertyType(" + ErrVect(1).ToString.Trim + "): Cannot determine property type."
            End Select

            Dim nEx As New Exception(ErrStr)
            Me.DxDbSysException(UserMsg, nEx, 5)
        End Sub
#End Region

    End Class
End Namespace
