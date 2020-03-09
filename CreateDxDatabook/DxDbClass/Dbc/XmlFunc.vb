Imports System.IO
Imports System.Xml
Imports DxDbClass.DxDB

Namespace DxDB
    Public MustInherit Class XmlFunc
        Inherits BaseClass

#Region "Members & Constructor"
        ' Options
        Friend AllowDuplNos As Boolean
        Friend UseOdbcAlias As Boolean
        Friend UseSymTable As Boolean

        ' HKP PDB data for DBC creation
        Friend LmcPdb As Pdb
        Friend DbFile As String

        Private ComVersion As Integer

        Private ElementCounter As Integer
        Private SerialKeyCounter As Integer
        Private ReadOnly _DbcDoc As XmlDocument     ' Instance of DBC File

        ' SQL Server Data
        Private Machine As String
        Private TcpPort As String
        Private Login As String
        Private Password As String
        Private Database As String

        Public Sub New(nComVersion As Integer)
            MyBase.New(nComVersion)
            Me.ComVersion = nComVersion
            Me._DbcDoc = New XmlDocument
        End Sub
#End Region

#Region "ParentClass Members & Constructor"
        Public ReadOnly Property DbcDoc() As XmlDocument
            Get
                Return Me._DbcDoc
            End Get
        End Property
#End Region

#Region "Public Methods"
        Public Function CreateEmptyDbcCfg() As Boolean
            Dim Decl As XmlDeclaration
            Dim XmlElem As XmlElement

            Try
                ' create Declaration
                Decl = Me.DbcDoc.CreateXmlDeclaration("1.0", "", "")
                Me.DbcDoc.AppendChild(CType(Decl, XmlDeclaration))

                ' Create Dbc Element
                XmlElem = Me.DbcDoc.CreateElement("dbc")
                XmlElem.SetAttribute("version", "1.0")

                ' add to XML document
                Me.DbcDoc.AppendChild(CType(XmlElem, XmlNode))

                ' add "CObject" Root Element - encloses all sub configs (level 0)
                If Not Me.AddRootObjectChild() Then Return False

                '-----------------------------
                ' Level 1 => CConfig, CObList
                ' Parent Node: CObject
                '-----------------------------

                ' add CConfig element - contains symbol/library definitions (level 1)
                XmlElem = Me.DbcDoc.CreateElement(elemCConfig)
                If Not Me.AddChildNode(elemCObject, XmlElem) Then Return False

                ' add CObList element - (level 1)
                XmlElem = Me.DbcDoc.CreateElement(elemCObList)
                XmlElem.SetAttribute("CPersistList", "0")
                If Not Me.AddChildNode(elemCObject, XmlElem) Then Return False

                '------------------------------------------
                ' Level 2 => CConfigSym, CConfigLib
                '            CConfigPref, CConfigScripting
                ' Parent Node: CConfig
                '------------------------------------------

                ' add CConfigSym element - contains power, port and sheetconn symbols (level 2)
                XmlElem = Me.DbcDoc.CreateElement(elemCConfigSym)
                If Not Me.AddChildNode(elemCConfig, XmlElem) Then Return False
                ' add CConfigSym/CObArray element - contains power, port and sheetconn symbols (level 2)
                Me.AddCObArrayElement(elemCConfigSym)

                ' add CConfigLib element - contains all library partitions (level 2)
                XmlElem = Me.DbcDoc.CreateElement(elemCConfigLib)
                If Not Me.AddChildNode(elemCConfig, XmlElem) Then Return False
                ' add CConfigLib/CObArray element - contains all library partitions (level 2)
                Me.AddCObArrayElement(elemCConfigLib)

                ' add CConfigPref element - contains DxDatabook preferences - has no CObArray element (level 2)
                XmlElem = Me.CreateCConfigPrefElement()
                If Not Me.AddChildNode(elemCConfig, XmlElem) Then Return False

                ' add CConfigScripting element - contains scripting params - has no CObArray element (level 2)
                XmlElem = Me.CreateCConfigScriptingElement()
                If Not Me.AddChildNode(elemCConfig, XmlElem) Then Return False
            Catch ex As Exception
                Return False
            End Try
            Return True
        End Function

        Public Function CreateLibPartitions(ByVal PdbPrtns As List(Of PdbPrtn)) As List(Of XmlElement)
            Dim LibEntry As XmlElement

            ' clone new array nodes w/o child nodes
            Dim LibNodesOk As New List(Of XmlElement)

            ' loop for each library in PdbPrtns
            For Each Prtn As PdbPrtn In PdbPrtns
                ' add ConfigAtt and Attributes
                MyBase.DxDbTranscriptMsg("... creating DBC Library: '" + Prtn.Name + "'", 10, MsgType.Txt, True)
                LibEntry = Me.CreateLibraryElement(Prtn)
                LibNodesOk.Add(CType(LibEntry.Clone, XmlElement))
            Next

            Return LibNodesOk
        End Function

        Public Function UpdateLibPartitions(OldLibraryNodes As List(Of XmlElement), PdbPrtns As List(Of PdbPrtn)) As List(Of XmlElement)
            Dim Prtn As PdbPrtn, XmlAttr As XmlAttribute
            Dim DbcOk As Boolean, LibraryNode As XmlElement

            '--------------------------------
            ' init new/updated LibNode Array
            '--------------------------------
            Dim LibrayNodesUpdated As New List(Of XmlElement)

            '-----------------------------------
            ' loop for each library in PdbPrtns
            '-----------------------------------
            For Each Prtn In PdbPrtns
                '---------------------------------------
                ' check if partition exists in DcbPrtns
                '---------------------------------------
                DbcOk = False
                For Each LibraryNode In OldLibraryNodes
                    XmlAttr = Me.GetAttribute("LibraryName", LibraryNode)
                    If XmlAttr.Value = Prtn.Name Then
                        ' ../CConfigLib/CObArray/CConfigLibEntry
                        MyBase.DxDbTranscriptMsg("... updating DBC Library: '" + XmlAttr.Value + "'", 10, MsgType.Txt, True)
                        LibraryNode = Me.UpdateLibEntry(LibraryNode, Prtn.Attrs, Prtn.Name)
                        If Not IsNothing(LibraryNode) Then
                            LibrayNodesUpdated.Add(CType(LibraryNode.Clone, XmlElement))
                        End If
                        DbcOk = True
                        Exit For
                    End If
                Next

                '------------------------------------
                ' if not found add new pdb Partition
                '------------------------------------
                If Not DbcOk Then
                    ' add ConfigAtt and Attributes
                    LibraryNode = Me.CreateLibraryElement(Prtn)
                    LibrayNodesUpdated.Add(CType(LibraryNode.Clone, XmlElement))
                End If
            Next

            Return LibrayNodesUpdated
        End Function
#End Region

#Region "DBC Create Library Methods"
        Private Function CreateLibEntry(ByVal Partition As PdbPrtn) As XmlElement
            Dim LibEntry As XmlElement, CConfigAtt As XmlElement
            Dim CObArray As XmlElement, CfgAttEntry As XmlNode

            ' create the new lib element
            LibEntry = Me.DbcDoc.CreateElement(elemCConfigLibEntry)
            LibEntry.SetAttribute("LibraryName", Utils.MakeValidDbTableName(Partition.Name))
            LibEntry.SetAttribute("JoinTable", "false")
            LibEntry.SetAttribute("Locked", "0")
            LibEntry.SetAttribute("SymbolExpression", "")
            LibEntry.SetAttribute("JoinType", "0")
            LibEntry.SetAttribute("HorizJoinType", "0")

            CConfigAtt = Me.DbcDoc.CreateElement(elemCConfigAtt)
            CObArray = Me.DbcDoc.CreateElement(elemCObArray)
            If Me.UseSymTable Then
                CConfigAtt.SetAttribute("UseCentralLibrarySymbols", "0")
            Else
                CConfigAtt.SetAttribute("UseCentralLibrarySymbols", "1")
            End If
            CObArray.SetAttribute(attrCXMLTypedPtrArraySize, CStr(Partition.Attrs.Count))

            ' add Attributes
            For Each PdbAttr As PrtnAttr In Partition.Attrs
                CfgAttEntry = Me.CreateLibEntryAttr(PdbAttr)
                CObArray.AppendChild(CfgAttEntry)
            Next

            ' add Attributes
            CConfigAtt.AppendChild(CType(CObArray, XmlNode))
            LibEntry.AppendChild(CType(CConfigAtt, XmlNode))

            Return LibEntry
        End Function

        Private Function CreateLibraryElement(ByVal Partition As PdbPrtn) As XmlElement
            Dim LibEntry As XmlElement, CfgTable As XmlElement

            ' create library entry
            LibEntry = Me.CreateLibEntry(Partition)

            ' create table def
            CfgTable = Me.CreateTableDef(Partition)

            ' add Attributes
            LibEntry.AppendChild(CfgTable)

            Return LibEntry
        End Function

        Private Function CreateTableDef(Partition As PdbPrtn) As XmlElement
            Dim DbTableName As String, Assembly As String
            Dim CObArray As XmlElement, CfgTable As XmlElement
            Dim CfgTableEntry As XmlElement, CStringList As XmlNode
            Dim CfgTableEntryCPDB As XmlElement = Nothing

            CfgTable = Me.DbcDoc.CreateElement(elemCConfigTable)
            CObArray = Me.DbcDoc.CreateElement(elemCObArray)

            If Me.ComVersion > 76 Then
                CObArray.SetAttribute(attrCXMLTypedPtrArraySize, "2")
            Else
                CObArray.SetAttribute(attrCXMLTypedPtrArraySize, "1")
            End If

            ' create table def entry
            CfgTableEntry = Me.DbcDoc.CreateElement(elemCConfigTableEntry)

            DbTableName = Utils.MakeValidDbTableName(Partition.Name)

            Assembly = Me.CreateAssemblySqlite()
            CfgTableEntry.SetAttribute("Assembly", Assembly)
            CfgTableEntry.SetAttribute("TableName", DbTableName)

            If Me.ComVersion > 76 Then
                ' create table def entry
                CfgTableEntryCPDB = Me.DbcDoc.CreateElement(elemCConfigTableEntry)

                Assembly = Me.CreateAssemblyCpdb()
                CfgTableEntryCPDB.SetAttribute("Assembly", Assembly)
                CfgTableEntryCPDB.SetAttribute("TableName", DbTableName)
            End If

            ' create string list
            CStringList = Me.CreateCStringList(Partition)
            CfgTableEntry.AppendChild(CStringList)
            CObArray.AppendChild(CType(CfgTableEntry, XmlNode))

            If Me.ComVersion > 76 AndAlso Not IsNothing(CfgTableEntryCPDB) Then
                CStringList = Me.CreateCStringList(Partition)
                CfgTableEntryCPDB.AppendChild(CStringList)
                CObArray.AppendChild(CType(CfgTableEntryCPDB, XmlNode))
            End If

            ' add table def
            CfgTable.AppendChild(CType(CObArray, XmlNode))

            Return CfgTable
        End Function

        Private Function CreateCStringList(LibAttrs As List(Of XmlElement)) As XmlNode
            Dim ND As XmlNode, EL As XmlElement, i As Integer = 0
            Dim NdAttr As XmlAttribute, LibAttr As XmlAttribute

            ' create string list
            ND = CType(Me.DbcDoc.CreateElement(elemCStringList), XmlNode)

            NdAttr = Me.DbcDoc.CreateAttribute("ListSize")
            NdAttr.Value = CStr(LibAttrs.Count)
            ND.Attributes.Append(NdAttr)

            For Each EL In LibAttrs
                LibAttr = Me.GetAttribute("FieldName", EL)
                NdAttr = Me.DbcDoc.CreateAttribute("ListItem" + CStr(i))
                NdAttr.Value = LibAttr.Value
                ND.Attributes.Append(NdAttr)
                i += 1
            Next

            Return ND
        End Function

        Private Function CreateCStringList(Partition As PdbPrtn) As XmlNode
            Dim ND As XmlNode, i As Integer = 0
            Dim NdAttr As XmlAttribute, PrtnAttr As PrtnAttr

            ' create string list
            ND = CType(Me.DbcDoc.CreateElement(elemCStringList), XmlNode)

            NdAttr = Me.DbcDoc.CreateAttribute("ListSize")
            NdAttr.Value = CStr(Partition.Attrs.Count)
            ND.Attributes.Append(NdAttr)

            For Each PrtnAttr In Partition.Attrs
                NdAttr = Me.DbcDoc.CreateAttribute("ListItem" + CStr(i))
                NdAttr.Value = PrtnAttr.Name
                MyBase.DxDbAdd2Logfile("... adding DBC Library Attribute: '" + PrtnAttr.Name + "'", 15)
                ND.Attributes.Append(NdAttr)
                i += 1
            Next

            Return ND
        End Function

        Private Function CreateAssemblyCpdb() As String
            Dim Assembly As String

            '<CConfigTableEntry21 Assembly="SQLITE;ALIAS=CPDB;HOST=QSQLITE;DB=${THIS_DBC_PATH}/Library.cpd;;IPD=0;" TableName="Analog">
            '  <CStringList22 ListSize="8" ListItem0="Part Number" ListItem1="Part Name" ListItem2="Type" ListItem3="Cell Name" ListItem4="Cell Pins" ListItem5="Desc" ListItem6="Tech" ListItem7="Height"/>
            '</CConfigTableEntry21>

            Assembly = "SQLITE;"
            Assembly += "ALIAS=CPDB;"
            Assembly += "HOST=QSQLITE;"
            Assembly += "DB=${THIS_DBC_PATH}/Library.cpd;;"
            Assembly += "IPD=0;"

            Return Assembly
        End Function

        Private Function CreateAssemblySqlite() As String
            Dim Assembly As String = String.Empty

            ' CConfigTableEntry20 Assembly="SQLITE;
            ' ALIAS=QSQLITE.Connection;HOST=QSQLITE;
            ' DB=D:\Users\Thomass\Documents\Programming\VBasic\EEtoolsVX\_LibraryTools\CreateDxDatabook\TestData\Library\Library.db3;;IPD=0;" TableName="Analog">

            If My.Computer.FileSystem.FileExists(Me.DbFile) Then
                Assembly = "SQLITE;"
                Assembly += "ALIAS=QSQLITE.Connection;"
                Assembly += "HOST=QSQLITE;"
                Assembly += "DB=" + Me.DbFile + ";;"
                Assembly += "IPD=0;"
            Else
                If Me.UseOdbcAlias And Not Me.DbFile = String.Empty Then
                    'Assembly="ODBC;ALIAS=DXDB_SQL;CNS={ODBC;DSN=DXDB_SQL;};DDS={};
                    Assembly = "ODBC;ALIAS=" + Me.DbFile + ";"
                    Assembly += "CNS={ODBC;DSN=" + Me.DbFile + ";};"
                    'Assembly  += "DDS={};"
                Else
                    Assembly = "ODBC;ALIAS=" + Me.DbFile + ";"
                    Assembly = "ODBC;"
                    Assembly += "CNS={ODBC;DSN=" + Me.DbFile + ";"
                    Assembly += "Server=" + Me.Machine + "," + Me.TcpPort + ";"
                    Assembly += "Database=" + Me.Database + ";"
                    Assembly += "Uid=" + Me.Login + ";"
                    Assembly += "Pwd=" + Me.Password + ";"
                    Assembly += "};"
                    'Assembly  += "DDS={};"
                End If
            End If

            Return Assembly
        End Function

        Private Function CreateLibEntryAttr(ByVal Attr As PrtnAttr) As XmlElement
            Dim CfgAttEntry As XmlElement
            CfgAttEntry = Me.DbcDoc.CreateElement(elemCConfigAttEntry)
            CfgAttEntry.SetAttribute("FieldName", Attr.Name)
            CfgAttEntry.SetAttribute("AttName", Attr.Name)
            CfgAttEntry.SetAttribute("DefaultValue", "")

            '-------------------------------
            ' set default Exclude attributes
            Select Case Attr.Name
                Case DeviceField, PartNameField, ValueProp
                    CfgAttEntry.SetAttribute("ExcludeWhenAnnotating", "0")
                    CfgAttEntry.SetAttribute("ExcludeWhenLoading", "0")
                Case SymbolField
                    CfgAttEntry.SetAttribute("ExcludeWhenAnnotating", "1")
                    CfgAttEntry.SetAttribute("ExcludeWhenLoading", "0")
                Case Else
                    CfgAttEntry.SetAttribute("ExcludeWhenAnnotating", "1")
                    CfgAttEntry.SetAttribute("ExcludeWhenLoading", "1")
            End Select

            '----------------------------------
            ' set default Visibility attributes
            Select Case Attr.Name
                Case ValueProp
                    CfgAttEntry.SetAttribute("m_bNameVisible", "0")
                    CfgAttEntry.SetAttribute("m_bValueVisible", "1")
                Case Else
                    CfgAttEntry.SetAttribute("m_bNameVisible", "0")
                    CfgAttEntry.SetAttribute("m_bValueVisible", "0")
            End Select

            '---------------------------------
            ' set default Field Type attribute
            Select Case Attr.Name
                Case DeviceField
                    If Me.AllowDuplNos Then
                        CfgAttEntry.SetAttribute("AttType", "0")
                    Else
                        CfgAttEntry.SetAttribute("AttType", "4")
                    End If
                Case DataSheetProp, M3DModelProp, IbisModelProp
                    CfgAttEntry.SetAttribute("AttType", "2")
                Case SymbolField
                    CfgAttEntry.SetAttribute("AttType", "1")
                Case PartLabelField, RefDesField
                    CfgAttEntry.SetAttribute("AttType", "3")
                Case Else
                    CfgAttEntry.SetAttribute("AttType", "0")
            End Select

            '---------------------------------
            ' set default Value Type attribute
            Select Case Attr.DaoType
                Case DAODataType.dbDouble, DAODataType.dbInteger
                    CfgAttEntry.SetAttribute("ValueType", "2") ' 2 = Real,Integer; 
                Case Else
                    CfgAttEntry.SetAttribute("ValueType", "1") ' 1 = Text; 
            End Select

            '-----------------------------
            ' set default other attributes
            CfgAttEntry.SetAttribute("MagType", "7")
            CfgAttEntry.SetAttribute("UnitsString", "")
            CfgAttEntry.SetAttribute("ShowUnitsForIEC62", "0")
            CfgAttEntry.SetAttribute("AddAsOat", "0")
            CfgAttEntry.SetAttribute("Required", "0")
            CfgAttEntry.SetAttribute("VerificationKey", "0")
            CfgAttEntry.SetAttribute("ValidMag", "2047")
            CfgAttEntry.SetAttribute("ValidMagDir", "63")
            CfgAttEntry.SetAttribute("FutureUseBool1", "0")
            CfgAttEntry.SetAttribute("FutureUseBool2", "0")
            CfgAttEntry.SetAttribute("FutureUseStr1", "")
            CfgAttEntry.SetAttribute("FutureUseStr2", "")
            CfgAttEntry.SetAttribute("MatchCondition", "--")
            If Attr.Name = SymbolField Then
                CfgAttEntry.SetAttribute("MatchCondition", "=")
            End If
            CfgAttEntry.SetAttribute("AttrUnplace", "5")
            CfgAttEntry.SetAttribute("SerialKey", "0")

            Return CfgAttEntry
        End Function
#End Region

#Region "DBC Update Library Methods"
        '------------------------------------------------------------
        ' Me.UpdateLibPartitions()
        '   -> Me.UpdateLibEntry()
        '   <CConfigLibEntry>/CConfigAtt/<CObArray>/CConfigAttEntry
        '   <CConfigLibEntry>/CConfigTable/<CObArray>/CConfigTableEntry
        '------------------------------------------------------------
        Private Function UpdateLibEntry(DbcLibEntry As XmlElement, PrtnAttrs As List(Of PrtnAttr), PrtnName As String) As XmlElement
            Dim LibEntryUpdated As XmlElement, CConfigTableNode As XmlNode
            Dim LibAttrNodes As New List(Of XmlElement), CConfigAttNode As XmlNode

            '---------------------
            ' Clone existing Node
            '---------------------
            LibEntryUpdated = CType(DbcLibEntry.Clone, XmlElement)

            '-------------------------------------------------------------------
            ' update attribute count <CConfigAtt UseCentralLibrarySymbols="0/1">
            '-------------------------------------------------------------------
            Me.UpdateUseCentralLibrarySymbols(LibEntryUpdated)

            '------------------------------------------
            ' Get libentry attribute node <CConfigAtt>
            '------------------------------------------
            CConfigAttNode = Nothing
            If Me.GetNode(elemCConfigAtt, CConfigAttNode, LibEntryUpdated) Then

                '----------------------------------------------------------------
                ' get/delete existing attributes entries <CConfigAttEntry(n) ....\>
                '----------------------------------------------------------------
                LibAttrNodes = Me.RemoveCObArrayChildNodes(CConfigAttNode)

                '----------------------------------
                ' updating existing XML partitions
                '----------------------------------
                LibAttrNodes = Me.UpdateLibPrtnAttrs(LibAttrNodes, PrtnAttrs)

                '----------------------------------------------------
                ' add updated elements to <CConfigLibEntry/CObArray>
                '----------------------------------------------------
                Me.AddCObArrayChildNodes(CConfigAttNode, LibAttrNodes)

                '----------------------------------------------------------------
                ' update attribute count <CObArray CXMLTypedPtrArraySize="(n)"\>
                '----------------------------------------------------------------
                Me.UpdateCObArraySize(CConfigAttNode)
            End If

            '------------------------------------------
            ' Get libentry table node <CConfigTable>
            '------------------------------------------
            CConfigTableNode = Nothing
            If Me.GetNode(elemCConfigTable, CConfigTableNode, LibEntryUpdated) Then
                Me.UpdateLibTableAttrs(CConfigTableNode, LibAttrNodes, PrtnName)
            End If

            Return LibEntryUpdated
        End Function

        Private Sub UpdateLibTableAttrs(CConfigTableNode As XmlNode, LibAttrs As List(Of XmlElement), PrtnName As String)
            Dim Assembly As String, DbTableName As String
            Dim CConfigTableCObArrayNode As XmlNode = Nothing
            Dim tmpElem As XmlElement, tmpNode As XmlNode
            Dim CStringList As XmlNode

            'elemCObArray
            'If Me.GetNode(elemCConfigTableEntry, CConfigTableEntryNode, CConfigTableNode) Then
            If Me.GetNode(elemCObArray, CConfigTableCObArrayNode, CConfigTableNode) Then

                DbTableName = Utils.MakeValidDbTableName(PrtnName)

                tmpElem = CType(CConfigTableCObArrayNode, XmlElement)

                If Me.ComVersion > 76 Then
                    tmpElem.SetAttribute(attrCXMLTypedPtrArraySize, "2")
                Else
                    tmpElem.SetAttribute(attrCXMLTypedPtrArraySize, "1")
                End If

                tmpElem = CType(CConfigTableCObArrayNode.FirstChild, XmlElement)
                tmpNode = CType(CConfigTableCObArrayNode.FirstChild, XmlNode)

                Assembly = Me.CreateAssemblySqlite()
                tmpElem.SetAttribute("Assembly", Assembly)
                tmpElem.SetAttribute("TableName", DbTableName)

                ' create string list
                CStringList = Me.CreateCStringList(LibAttrs)
                tmpNode.ReplaceChild(CStringList, tmpNode.FirstChild)

                If Me.ComVersion > 76 AndAlso CConfigTableCObArrayNode.ChildNodes.Count = 2 Then
                    ' create table def entry
                    tmpElem = CType(CConfigTableCObArrayNode.LastChild, XmlElement)
                    tmpNode = CType(CConfigTableCObArrayNode.LastChild, XmlNode)

                    Assembly = Me.CreateAssemblyCpdb()
                    tmpElem.SetAttribute("Assembly", Assembly)
                    tmpElem.SetAttribute("TableName", DbTableName)

                    ' create string list
                    CStringList = Me.CreateCStringList(LibAttrs)
                    tmpNode.ReplaceChild(CStringList, tmpNode.LastChild)
                End If
            End If
        End Sub

        Private Function UpdateLibPrtnAttrs(ByVal DbcPrtnAttrs As List(Of XmlElement), ByVal PdbPrtnAttrs As List(Of PrtnAttr)) As List(Of XmlElement)
            Dim PdbAttr As PrtnAttr, DbcOk As Boolean
            Dim CfgAttEntry As XmlElement, XmlAttr As XmlAttribute

            '--------------------------------
            ' init new/updated LibNode Array
            '--------------------------------
            Dim PrtnAttrsUpdated As New List(Of XmlElement)

            '---------------------------------------------------
            ' loop for each pdb partition attribute in PdbPrtns
            '---------------------------------------------------
            For Each PdbAttr In PdbPrtnAttrs
                '--------------------------------------
                ' check if attribute exists in DcbPrtn
                '--------------------------------------
                DbcOk = False
                For Each CfgAttEntry In DbcPrtnAttrs
                    XmlAttr = Me.GetAttribute("FieldName", CfgAttEntry)
                    'XmlAttr = Me.GetAttribute("AttName", AttEntry)

                    If XmlAttr.Value = PdbAttr.Name Then
                        If Not IsNothing(CfgAttEntry) Then
                            MyBase.DxDbAdd2Logfile("... preserving DBC Library Attribute: '" + XmlAttr.Value + "'", 15)
                            PrtnAttrsUpdated.Add(CType(CfgAttEntry.Clone, XmlElement))
                        End If
                        DbcOk = True
                        Exit For
                    End If
                Next

                '------------------------------------
                ' if not found add new pdb attribute
                '------------------------------------
                If Not DbcOk Then
                    ' add ConfigAtt and Attributes
                    CfgAttEntry = Me.CreateLibEntryAttr(PdbAttr)
                    If Not IsNothing(CfgAttEntry) Then
                        XmlAttr = Me.GetAttribute("FieldName", CfgAttEntry)
                        MyBase.DxDbAdd2Logfile("... adding DBC Library Attribute: '" + XmlAttr.Value + "'", 15)
                        PrtnAttrsUpdated.Add(CType(CfgAttEntry.Clone, XmlElement))
                    End If
                End If
            Next

            Return PrtnAttrsUpdated
        End Function
#End Region

#Region "Protected XML Base Methods - visible to parent class"
        ' Add CObject child node to the document root element.
        Protected Function AddRootObjectChild() As Boolean
            Dim DbcRootObject As XmlElement
            DbcRootObject = Me.DbcDoc.CreateElement(elemCObject)
            ' add "CObject" Root Element - encloses all sub configs
            DbcRootObject.SetAttribute("overallSchema", "21")
            DbcRootObject.SetAttribute("flag", "DX Databook Overlay Configuration File")
            DbcRootObject.SetAttribute("BaseConfigurationURL", "")
            Dim Nd As XmlNode = Me.DbcDoc.DocumentElement.AppendChild(CType(DbcRootObject, XmlNode))
            If IsNothing(Nd) Then Return False
            Return True
        End Function

        ''' <summary>
        ''' Loops from a specified RootNode recursively to the specified ParentNode and
        ''' appennd a chile node. If RootNode is not specified RootNode is to document root.
        ''' </summary>
        ''' <param name="ParentNodeName">Node where child should be appended</param>
        ''' <param name="DbcChild">Child Node to append</param>
        ''' <param name="RootNode">Node where recursive search starts</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Function AddChildNode(ParentNodeName As String, DbcChild As XmlElement, Optional ByRef RootNode As XmlNode = Nothing) As Boolean
            If IsNothing(RootNode) And IsNothing(Me.DbcDoc) Then Return False
            If IsNothing(RootNode) Then RootNode = CType(Me.DbcDoc.DocumentElement, XmlNode)

            For i As Integer = 0 To RootNode.ChildNodes.Count - 1
                Dim Nd As XmlNode = RootNode.ChildNodes(i)
                If TypeOf Nd Is XmlElement Then
                    If Nd.Name = ParentNodeName Then
                        RootNode.ChildNodes(i).AppendChild(CType(DbcChild, XmlNode))
                        Return True
                    End If
                    If Me.AddChildNode(ParentNodeName, DbcChild, Nd) Then Return True
                End If
            Next
            Return False
        End Function

        ''' <summary>
        ''' Adds a CObArray XML element to any level
        ''' </summary>
        ''' <param name="ParentNode">ParentNode for CObArray</param>
        ''' <param name="ArraySize"> Optional: Defined array size</param>
        ''' <remarks></remarks>
        Protected Sub AddCObArrayElement(ParentNode As String, Optional ArraySize As Integer = 0)
            Dim DbcObj As XmlElement
            DbcObj = Me.DbcDoc.CreateElement(elemCObArray)
            DbcObj.SetAttribute(attrCXMLTypedPtrArraySize, ArraySize.ToString)
            Me.AddChildNode(ParentNode, DbcObj)
        End Sub
#End Region

#Region "Protected XML Base Methods"
        Protected Function GetNode(DbcName As String, ByRef DbcNode As XmlNode, Optional RootNode As XmlNode = Nothing) As Boolean
            Dim Nd As XmlNode

            If IsNothing(RootNode) And IsNothing(Me.DbcDoc) Then Return False

            If IsNothing(RootNode) Then RootNode = CType(Me.DbcDoc.DocumentElement, XmlNode)

            ' loop all following members in XML-file
            DbcNode = Nothing
            For Each Nd In RootNode.ChildNodes
                If TypeOf Nd Is XmlElement Then
                    If Me.GetElementName(Nd.Name) = DbcName Then
                        DbcNode = Nd
                        Return True
                    End If
                    If Me.GetNode(DbcName, DbcNode, Nd) Then Return True
                End If
            Next

            Return False
        End Function

        Protected Function ReNameNode(NodeName As String, Nd As XmlNode, ByRef RootNode As XmlNode) As XmlNode
            Dim eOld As XmlElement, eNew As XmlElement
            Dim Attr As XmlAttribute, cNd As XmlNode

            eOld = CType(Nd, XmlElement)
            eNew = Me.DbcDoc.CreateElement(NodeName)
            ' copy attributes
            For Each Attr In eOld.Attributes
                eNew.SetAttribute(Attr.Name, Attr.Value)
            Next

            ' copy child nodes
            For Each cNd In eOld.ChildNodes
                eNew.AppendChild(cNd.Clone)
            Next

            RootNode.ReplaceChild(CType(eNew, XmlNode), CType(eOld, XmlNode))

            Return eNew
        End Function

        Protected Function GetElement(DbcName As String, ByRef DbcObj As XmlElement, Optional RootNode As XmlNode = Nothing) As Boolean
            Dim Nd As XmlNode

            If IsNothing(RootNode) And IsNothing(Me.DbcDoc) Then Return False

            If IsNothing(RootNode) Then RootNode = CType(Me.DbcDoc.DocumentElement, XmlNode)

            ' loop all following members in XML-file
            For Each Nd In RootNode.ChildNodes
                If TypeOf Nd Is XmlElement Then
                    If Me.GetElementName(Nd.Name) = DbcName Then
                        DbcObj = CType(Nd, XmlElement)
                        Return True
                    End If
                    If Me.GetElement(DbcName, DbcObj, Nd) Then Return True
                End If
            Next

            Return False
        End Function

        Protected Function GetElementName(Name As String) As String
            For i As Integer = Len(Name) To 0 Step -1
                If Not IsNumeric(Name.Substring(i - 1, 1)) Then
                    Return Name.Substring(0, i)
                End If
            Next
            Return Name
        End Function

        Protected Function GetAttribute(AttrName As String, Element As XmlElement) As XmlAttribute
            Dim Attr As XmlAttribute
            For Each Attr In Element.Attributes
                If LCase(Attr.Name) = LCase(AttrName) Then
                    Return Attr
                End If
            Next
            Return Nothing
        End Function

        Protected Function GetAttributeValue(AttrName As String, DefaultValue As String, Element As XmlElement) As String
            Dim Attr As XmlAttribute
            For Each Attr In Element.Attributes
                If LCase(Attr.Name) = LCase(AttrName) Then
                    Return Attr.Value
                End If
            Next
            Return DefaultValue
        End Function

        Protected Sub DeIndexElements(Optional ByRef RootNode As XmlNode = Nothing)
            Dim NodeName As String, Nd As XmlNode

            If IsNothing(RootNode) And IsNothing(Me.DbcDoc) Then Exit Sub

            If IsNothing(RootNode) Then
                RootNode = CType(Me.DbcDoc.DocumentElement, XmlNode)
            End If

            For i As Integer = 0 To RootNode.ChildNodes.Count - 1
                Nd = RootNode.ChildNodes(i)
                If Nd.NodeType = XmlNodeType.Element Then
                    NodeName = Me.GetElementName(Nd.Name)
                    Nd = Me.ReNameNode(NodeName, Nd, RootNode)
                    Me.DeIndexElements(Nd)
                End If
            Next
        End Sub

        Protected Sub ReIndexElements(Optional ByRef RootNode As XmlNode = Nothing)
            Dim NodeName As String, ElementName As String, Nd As XmlNode

            If IsNothing(RootNode) And IsNothing(Me.DbcDoc) Then Exit Sub

            If IsNothing(RootNode) Then
                RootNode = CType(Me.DbcDoc.DocumentElement, XmlNode)
                Me.ElementCounter = 0
                Me.SerialKeyCounter = 0
            End If

            For i As Integer = 0 To RootNode.ChildNodes.Count - 1
                Nd = RootNode.ChildNodes(i)
                If Nd.NodeType = XmlNodeType.Element Then
                    ElementName = Me.GetElementName(Nd.Name)
                    If ElementName = elemCConfigAttEntry Then
                        CType(Nd, XmlElement).SetAttribute("SerialKey", CStr(Me.SerialKeyCounter))
                        Me.SerialKeyCounter += 1
                    End If
                    NodeName = Me.GetElementName(Nd.Name) + CStr(Me.ElementCounter)
                    Nd = Me.ReNameNode(NodeName, Nd, RootNode)
                    Me.ElementCounter += 1
                    Me.ReIndexElements(Nd)
                End If
            Next
        End Sub
#End Region

#Region "Protected DBC Utilitiy Methods"
        ''' <summary>
        ''' Add a list of XML Elements to CObArray of given XML node.
        ''' </summary>
        ''' <param name="CObArrayParent">Parent Node of CObArray</param>
        ''' <param name="XmlElementArray">Arraylist of XML Elements</param>
        Protected Sub AddCObArrayChildNodes(ByRef CObArrayParent As XmlNode, XmlElementArray As List(Of XmlElement))
            Dim CObArray As XmlNode = CObArrayParent.FirstChild
            For i As Integer = 0 To XmlElementArray.Count - 1
                ' <CObArrayParent>/CObArray => First Child
                CObArray.AppendChild(CType(XmlElementArray(i), XmlNode))
            Next
        End Sub

        ''' <summary>
        ''' Update CConfigAtt UseCentralLibrarySymbols
        ''' </summary>
        ''' <param name="LibEntryOk"></param>
        ''' <remarks></remarks>
        Protected Sub UpdateUseCentralLibrarySymbols(ByRef LibEntryOk As XmlElement)
            Dim UseClibSym As XmlElement = CType(LibEntryOk.FirstChild, XmlElement)
            UseClibSym.RemoveAttribute("UseCentralLibrarySymbols")
            If Me.UseSymTable Then
                UseClibSym.SetAttribute("UseCentralLibrarySymbols", "0")
            Else
                UseClibSym.SetAttribute("UseCentralLibrarySymbols", "1")
            End If
            LibEntryOk.ReplaceChild(CType(UseClibSym, XmlNode), LibEntryOk.FirstChild)
        End Sub

        ''' <summary>
        ''' Updates the CObArray element count of given XML node
        ''' </summary>
        ''' <param name="CObArrayParent">Parent Node of CObArray</param>
        Protected Sub UpdateCObArraySize(CObArrayParent As XmlNode)
            Dim CObArray As XmlElement = CType(CObArrayParent.FirstChild, XmlElement)
            ' CObArray is first childnode of CConfigLib. Has to be an XmlElement to change attributes
            CObArray.SetAttribute(attrCXMLTypedPtrArraySize, CStr(CObArray.ChildNodes.Count))
            'CObArrayParent.ReplaceChild(CType(CObArray, XmlNode), CObArrayParent.FirstChild)
        End Sub

        ''' <summary>
        ''' Save current CObArray elements in ArrayList and delete elements.
        ''' </summary>
        ''' <param name="CObArrayParent">Parent Node of CObArray</param>
        ''' <returns></returns>
        Protected Function RemoveCObArrayChildNodes(ByRef CObArrayParent As XmlNode) As List(Of XmlElement)
            Dim ArrayChilds As New List(Of XmlElement)
            Dim CObArray As XmlNode = CObArrayParent.FirstChild

            '---------------------------------------------
            ' save existing CObArray entries in ArrayList
            '---------------------------------------------
            For i As Integer = 0 To CObArray.ChildNodes.Count - 1
                ArrayChilds.Add(CType(CObArray.ChildNodes(i).Clone, XmlElement))
            Next

            '-----------------------------
            ' delete CObArray child nodes
            '-----------------------------
            Me.RemoveAllChilds(CObArray)

            Return ArrayChilds
        End Function
#End Region

#Region "Private XML Base Creation Methods"
        Private Sub RemoveAllChilds(ByRef Parent As XmlNode)
            For i As Integer = Parent.ChildNodes.Count - 1 To 0 Step -1
                Parent.RemoveChild(Parent.ChildNodes(i))
            Next
        End Sub

        Friend Function CreateCConfigPrefElement() As XmlElement
            Dim CConfigPref As XmlElement
            CConfigPref = Me.DbcDoc.CreateElement(elemCConfigPref)
            CConfigPref.SetAttribute("LoadIntoNewWindow", "0")
            CConfigPref.SetAttribute("SymbolAttributesGetLoaded", "0")
            CConfigPref.SetAttribute("ComponentTracking", "1")
            CConfigPref.SetAttribute("AnnotateUndefinedAttributes", "0")
            CConfigPref.SetAttribute("ChangeCompEnabled", "0")
            CConfigPref.SetAttribute("ComponentTrackingZoom", "0")
            CConfigPref.SetAttribute("MaintainAttVisibility", "1")
            CConfigPref.SetAttribute("LibraryAttEnabled", "1")
            CConfigPref.SetAttribute("LibraryAttName", "DXDB_LIBNAME")
            CConfigPref.SetAttribute("AnnotateUndefinedAttributesValue", "")
            CConfigPref.SetAttribute("UseLoadTabAttribute", "0")
            CConfigPref.SetAttribute("LoadTabAttributeName", "")
            CConfigPref.SetAttribute("AnnotateSelected", "0")
            CConfigPref.SetAttribute("AnnotateSelectedPrefix", "ALT_*")
            CConfigPref.SetAttribute("AnnotateRemoveAttributes", "1")
            CConfigPref.SetAttribute("ExcludedAttList", "&quot;Ref Designator&quot;")
            Return CConfigPref
        End Function

        Friend Function CreateCConfigScriptingElement() As XmlElement
            Dim ScriptElem, CfgScriptElem As XmlElement
            ScriptElem = Me.DbcDoc.CreateElement(elemCConfigScripting)
            ScriptElem.SetAttribute("OnAddComponentEnabled", "1")
            ScriptElem.SetAttribute("OnAnnoComponentEnabled", "1")
            ScriptElem.SetAttribute("OnLoadComponentEnabled", "1")
            ScriptElem.SetAttribute("OnSelectComponentEnabled", "1")
            ScriptElem.SetAttribute("OnViewDocumentEnabled", "1")
            ScriptElem.SetAttribute("OnAfterAddComponentEnabled", "1")
            ScriptElem.SetAttribute("OnAfterAnnoComponentEnabled", "1")
            ' add CConfigScripting->CConfigScriptLanguage element
            CfgScriptElem = Me.DbcDoc.CreateElement(elemCConfigScriptLanguage)
            CfgScriptElem.SetAttribute("Filename", "")
            CfgScriptElem.SetAttribute("type", "0")
            CfgScriptElem.SetAttribute("progID", "VBScript")
            ScriptElem.AppendChild(CType(CfgScriptElem, XmlNode))
            Return ScriptElem
        End Function
#End Region

    End Class
End Namespace
