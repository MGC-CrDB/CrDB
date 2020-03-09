Imports System.Globalization
Imports System.Threading

Imports MGCPCBPartsEditor
Imports DxDbClass
Imports System.IO
Imports System.Data.SqlClient

Imports Microsoft.SqlServer.Management
Imports Microsoft.SqlServer.Management.Smo
Imports Microsoft.SqlServer.Management.Common

Namespace DxDB
    Public Class SqlSvr
        Inherits BaseClass

#Region "Constants & Declarations"
        Public Const MaxStrLen As Integer = 128
        Protected PnCount As Integer
        Protected SqlPartsDb As DxDB.Pdb

        ' SQL Server Connection Properties
        Protected SQLServerConnectionValid As Boolean
        Protected SQLServerConnectionString As String
        Protected SQLServerConnection As ServerConnection
        Protected SQLServerDxDbDatabase As String
#End Region

#Region "Public Methods"
        Public Sub New(ByVal DatabaseName As String, ByVal PartsDb As DxDB.Pdb)
            Me.PnCount = 0
            Me.SqlPartsDb = PartsDb
            Me.SQLServerDxDbDatabase = DatabaseName
            Me.SQLServerConnectionValid = False
        End Sub

        Public Function CreateSqlSvrDB(ByVal Server As String, ByVal Port As String, ByVal Login As String, ByVal Passwd As String) As Boolean
            Dim CreateSuccess As Boolean = False
            'DXDB=1433 - ASQLINSTANCE=1434
            If Me.SetSQLServerConnection(Server, Port, Login, Passwd) Then
                CreateSuccess = Me.CreateDatabase()
            End If
            Return CreateSuccess
        End Function

        Public Sub TestSqlSvrConn(ByVal Server As String, ByVal Port As String, ByVal Login As String, ByVal Passwd As String)
            'DXDB=1433 - ASQLINSTANCE=1434
            If Me.SetSQLServerConnection(Server, Port, Login, Passwd, True) Then
                Me.DxDbTranscriptMsg("SQL Server Instance connected successfully.", 5, MsgType.Nte)
                Me.DxDbTranscriptMsg("")
            Else
                Me.DxDbTranscriptMsg("Connect to SQL Server Instance failed!", 5, MsgType.Err)
                Me.DxDbTranscriptMsg("")
            End If
        End Sub

        Public Function SetSQLServerConnection(Machine As String, TcpPort As String, Login As String, Password As String, Optional TestConnection As Boolean = False) As Boolean
            Try
                Me.SQLServerConnectionValid = False
                Me.SQLServerConnectionString = "Data Source=" + Machine + "," + TcpPort + ";Network Library=DBMSSOCN;User ID=" + Login + ";Password=" + Password + ";"

                'Dim SqlCon As New System.Data.SqlClient.SqlConnection(Me.SQLServerConnectionString)
                'Me.SQLServerConnection = New ServerConnection(SqlCon)
                Me.SQLServerConnection = New ServerConnection(Machine + "," + TcpPort, Login, Password)

                Me.SQLServerConnection.Connect()
                Me.SQLServerConnectionValid = Me.SQLServerConnection.IsOpen
                If TestConnection Then Me.ServerDisconnect()

                Return Me.SQLServerConnectionValid
            Catch ex As Exception
                Me.DxDbTranscriptMsg(ex.Message, 5, MsgType.Err)
                Return False
            End Try
        End Function

        Public Function CreateDatabase() As Boolean
            Dim SqlSvr As Server, SqlDB As Database

            Try
                If Me.SqlPartsDb.Partitions.Count = 0 Then Return False

                ' set CultureInfo to "en-US"
                Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US", False)

                ' create server database
                If Not Me.SQLServerConnectionValid Then
                    Me.DxDbTranscriptMsg("Connect to SQL Server Instance failed!", 5, MsgType.Err)
                    Me.DxDbTranscriptMsg("")
                    Return False
                End If

                Me.DxDbTranscriptMsg("Creating SQL Server Database", 5)

                ' create Server object
                SqlSvr = New Server(Me.SQLServerConnection)

                ' check if database exists
                ' delete database
                If Not Me.DeleteSqlServerDb(SqlSvr) Then
                    Me.DxDbTranscriptMsg("Creating SQL Server Database failed!", 5, MsgType.Err)
                    Me.DxDbTranscriptMsg("")
                    Return False
                End If

                ' create server database
                If Not Me.CreateSqlServerDb(SqlSvr) Then
                    Me.DxDbTranscriptMsg("Creating SQL Server Database failed!", 5, MsgType.Err)
                    Me.DxDbTranscriptMsg("")
                    Return False
                End If

                ' get the database object from server
                SqlDB = SqlSvr.Databases(Me.SQLServerDxDbDatabase)

                If IsNothing(SqlDB) Then
                    Me.DxDbTranscriptMsg("Create SQL Server Database failed.", 0, MsgType.Err)
                    Me.DxDbTranscriptMsg("")
                    Return False
                End If

                ' create empty table for all symbols#
                'Me.CreateSymTable(SqlDB)

                ' creating databook
                Me.CreatePartitionTables(SqlDB)

                'Me.SQLServerConnection.Disconnect()
                Me.ServerDisconnect()

                Me.DxDbTranscriptMsg("Creating SQL Server Database done", 5)
                Me.DxDbTranscriptMsg("")
                Me.DxDbTranscriptMsg("Partnumbers added: " + CStr(PnCount), 5)
                Me.DxDbTranscriptMsg("")

                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
#End Region

#Region "Protected Methods"
        Protected Sub ServerDisconnect()
            Me.SQLServerConnection.ForceDisconnected()
            Me.SQLServerConnection.Disconnect()
            Me.SQLServerConnection.Cancel()
            Me.SQLServerConnection = Nothing
        End Sub

        Protected Function DeleteSqlServerDb(ByRef Svr As Server) As Boolean
            Try
                Dim Db As Database = Svr.Databases(Me.SQLServerDxDbDatabase)
                If Not IsNothing(Db) Then Db.Drop()
                Return True
            Catch ex As Exception
                Me.DxDbSysException("DeleteSqlServerDb(): '" + Me.SQLServerDxDbDatabase + "'", ex, 10)
                Me.DxDbTranscriptMsg("")
                Return False
            End Try
        End Function

        Protected Function CreateSqlServerDb(ByRef Svr As Server) As Boolean
            Dim db As Database

            Try
                db = New Database(Svr, Me.SQLServerDxDbDatabase)
                db.Create()

                'Reference the database and display the date when it was created.
                db = Svr.Databases(Me.SQLServerDxDbDatabase)

                If Not IsNothing(db) Then
                    Return True
                End If

                Return False
            Catch ex As Exception
                Me.DxDbSysException("CreateSqlServerDb(): '" + Me.SQLServerDxDbDatabase + "'", ex, 10)
                Me.DxDbTranscriptMsg("")
                Return False
            End Try
        End Function

        Protected Sub CreateSymTable(ByRef SqlDB As Database)
            Dim SymTableDef As Table = Nothing
            Dim TblColumn As Column = Nothing

            If IsNothing(SqlDB) Then Exit Sub

            Me.DxDbAdd2Logfile("... creating Symbol Table '" + BaseClass.SymTableName + "'", 10, MsgType.Txt)

            Try
                ' create symbol table
                ' Define a Table object variable by supplying the parent database and table name in the constructor. 
                SymTableDef = New Table(SqlDB, BaseClass.SymTableName)

                ' create column device
                TblColumn = New Column(SymTableDef, BaseClass.DeviceField, DataType.VarChar(DxDB.SqlSvr.MaxStrLen))
                TblColumn.Nullable = False
                SymTableDef.Columns.Add(TblColumn)

                ' create column symbol
                TblColumn = New Column(SymTableDef, BaseClass.SymbolField, DataType.VarChar(DxDB.SqlSvr.MaxStrLen))
                TblColumn.Nullable = True
                SymTableDef.Columns.Add(TblColumn)

                'Create the table on the instance of SQL Server.
                SymTableDef.Create()

                'Me.CreateSymbTableIndex(SymTableDef)
            Catch ex As Exception
                Me.DxDbSysException("CreateSymTable(): '" + BaseClass.SymTableName + "'", ex, 5)
            End Try
        End Sub

        Protected Sub CreatePartitionTables(ByRef SqlDB As Database)
            Dim i As Integer, PnTableName As String
            Dim DataTblDef As DataTable

            If IsNothing(SqlDB) Then Exit Sub

            '--------------------------------------
            ' create empty table for each partition
            For i = 0 To Me.SqlPartsDb.Partitions.Count - 1
                ' set tablename to partition name
                PnTableName = Me.RemFileExt(Me.SqlPartsDb.Partition(i).Name)
                Me.DxDbTranscriptMsg("... creating Table '" + PnTableName + "'", 10)

                ' create parts DataTable for partition
                DataTblDef = Me.CreatePrtnPartNoTable(SqlDB, PnTableName, Me.SqlPartsDb.Partition(i))

                If Not IsNothing(DataTblDef) Then
                    Me.DxDbTranscriptMsg("... adding " + CStr(Me.SqlPartsDb.Partition(i).Parts.Count) + " Partnumber Record(s)", 15)
                    Me.CreateDataTableEntries(DataTblDef, Me.SqlPartsDb.Partition(i).Parts)
                    Me.SqlBulkCopyDataTable(DataTblDef)
                End If
            Next
        End Sub

        Protected Function CreatePrtnPartNoTable(ByRef SqlDB As Database, ByVal TableName As String, ByVal Partition As DxDB.PdbPrtn) As DataTable
            Dim SqlTableDef As Table
            Dim DaoFields As List(Of DaoField)

            Try
                ' building fieldname database
                DaoFields = Me.CreateFieldsDB(Partition)

                ' create SQL table
                SqlTableDef = Me.CreateSqlPrtnTable(SqlDB, TableName, DaoFields)
                If Not IsNothing(SqlTableDef) Then
                    'Create the table index for partnumber.
                    'Me.CreatePrtnTableIndex(SqlTableDef)
                Else
                    Return Nothing
                End If

                ' create Data table columns
                Return Me.CreateDataTableColumns(TableName, DaoFields)
            Catch Ex As System.Exception
                Me.DxDbSysException("CreatePrtnPartNoTable(): '" + TableName + "'", Ex, 5)
                Return Nothing
            End Try
        End Function

        Protected Sub CreateDataTableEntries(ByRef DataTblDef As DataTable, ByVal PartNos As List(Of PdbPart))
            Dim PartNo As PdbPart = Nothing

            Try
                Me.DxDbProgBarAction(PartNos.Count, ProgbarMode.Init)
                For i As Integer = 0 To PartNos.Count - 1
                    ' add record for each partnumber
                    Me.PnCount += 1
                    Me.DxDbStatusbarMsg("(" & i + 1 & "," & PartNos.Count & ") " & CType(PartNos(i), PdbPart).PartNo)
                    PartNo = CType(PartNos(i), PdbPart)

                    Me.AddDatTablePartEntry(DataTblDef, PartNo)
                    Me.DxDbProgBarAction(0, ProgbarMode.Incr)
                Next
                Me.DxDbProgBarAction(0, ProgbarMode.Close)
            Catch ex As Exception
                Me.DxDbSysException("CreateDataTableEntries(): " + DataTblDef.TableName, ex, 5)
            End Try
        End Sub

        Protected Sub AddDatTablePartEntry(ByRef PartTable As DataTable, ByVal PartNo As PdbPart)
            Dim Props As List(Of PdbProp), Prop As PdbProp
            Dim PnRecord As DataRow

            Prop = New PdbProp()

            Try
                Props = PartNo.Props

                '---------------------------------------
                ' create a new reord for each partnumber
                If Not IsNothing(PartTable) Then
                    PnRecord = PartTable.NewRow()

                    '-----------------------------
                    ' add mandatory PDB properties
                    For Each Col As DataColumn In PnRecord.Table.Columns
                        Select Case Col.ColumnName
                            Case BaseClass.DeviceField ' device 
                                PnRecord(Col.ColumnName) = PartNo.PartNo
                            Case PartNameField ' partname 
                                PnRecord(Col.ColumnName) = PartNo.Name
                            Case PartLabelField ' partlabel 
                                PnRecord(Col.ColumnName) = PartNo.Label
                            Case RefDesField ' refdes 
                                PnRecord(Col.ColumnName) = PartNo.RefPref
                            Case CellNameField ' pkg_type 
                                PnRecord(Col.ColumnName) = Me.GetDefaultCell(PartNo)
                            Case DescriptField ' desc 
                                PnRecord(Col.ColumnName) = PartNo.Desc
                            Case TypeField
                                PnRecord(Col.ColumnName) = PartNo.Type
                        End Select
                    Next

                    '-------------------------
                    ' add other PDB properties
                    For p As Integer = 0 To Props.Count - 1
                        Prop = CType(Props(p), PdbProp)
                        For Each Col As DataColumn In PnRecord.Table.Columns
                            If Prop.Name.ToLower = Col.ColumnName.ToLower Then
                                Select Case Col.DataType.Name
                                    Case "Double"
                                        PnRecord(Col.ColumnName) = CDbl(Prop.Value)
                                    Case "Integer"
                                        PnRecord(Col.ColumnName) = CInt(Prop.Value)
                                    Case "String"
                                        PnRecord(Col.ColumnName) = CStr(Prop.Value)
                                End Select
                            End If
                        Next
                    Next

                    PartTable.Rows.Add(PnRecord)
                End If
            Catch ex As Exception
                Me.DxDbSysException("AddDatTablePartEntry(): PN=" + PartNo.PartNo + ", Prop: " + Prop.Name, ex, 5)
            End Try
        End Sub

        Protected Sub SqlBulkCopyDataTable(ByRef Table2Copy As DataTable)
            Dim SqlCon As SqlConnection, SqlBlkCp As SqlBulkCopy

            Try
                SqlCon = New SqlConnection(Me.SQLServerConnectionString)

                SqlCon.Open()
                SqlCon.ChangeDatabase(Me.SQLServerDxDbDatabase)

                SqlBlkCp = New SqlBulkCopy(SqlCon)
                SqlBlkCp.DestinationTableName = "dbo." + Table2Copy.TableName

                SqlBlkCp.WriteToServer(Table2Copy)

                SqlCon.Close()
                SqlCon.Dispose()
            Catch ex As Exception
                Me.DxDbSysException("SqlBulkCopyDataTable(): '" + Table2Copy.TableName + "'", ex, 5)
            End Try
        End Sub

        Protected Function GetDefaultCell(ByVal PartNo As PdbPart) As String
            If PartNo.TopCell <> "" Then Return PartNo.TopCell
            If PartNo.BotCell <> "" Then Return PartNo.BotCell
            If PartNo.AltCells.Count = 0 Then Return ""
            Return PartNo.AltCells(0).ToString
        End Function

        Protected Sub CreateSymbTableIndex(ByRef SymTableDef As Table)
            Dim PrimKeyIndex As Index
            Dim IndexCol As IndexedColumn

            If IsNothing(SymTableDef) Then Exit Sub

            '------------------------------------------
            ' create primary key index field partnumber
            PrimKeyIndex = New Index(SymTableDef, "PrimKey_" + SymTableDef.Name)
            PrimKeyIndex.IndexKeyType = IndexKeyType.DriPrimaryKey
            PrimKeyIndex.IsClustered = False
            PrimKeyIndex.IgnoreDuplicateKeys = True
            PrimKeyIndex.IsUnique = False

            'Add indexed columns to the index.
            IndexCol = New IndexedColumn(PrimKeyIndex, BaseClass.DeviceField, False)

            'Add indexed columns to the index.
            PrimKeyIndex.IndexedColumns.Add(IndexCol)

            'Create the index on the instance of SQL Server.
            PrimKeyIndex.Create()
        End Sub

        Protected Sub CreatePrtnTableIndex(ByRef PnTableDef As Table)
            Dim PrimKeyIndex As Index
            'Dim UniqueIndex As Index
            Dim IndexCol As IndexedColumn

            If IsNothing(PnTableDef) Then Exit Sub

            '------------------------------------------
            ' create primary key index field partnumber
            PrimKeyIndex = New Index(PnTableDef, "PrimKey_" + PnTableDef.Name)
            PrimKeyIndex.IndexKeyType = IndexKeyType.DriPrimaryKey
            PrimKeyIndex.IsClustered = False
            PrimKeyIndex.IgnoreDuplicateKeys = True
            PrimKeyIndex.IsUnique = False

            'Add indexed columns to the index.
            IndexCol = New IndexedColumn(PrimKeyIndex, BaseClass.DeviceField, False)

            'Add indexed columns to the index.
            PrimKeyIndex.IndexedColumns.Add(IndexCol)

            'Create the index on the instance of SQL Server.
            PrimKeyIndex.Create()

            '-----------------------------------------
            ' create unique key index field partnumber
            'If Not Me.AllowDuplPartNos Then
            '    UniqueIndex = New Index(PnTableDef, "Unique_" + PnTableDef.Name)
            '    UniqueIndex.IndexKeyType = IndexKeyType.DriUniqueKey
            '    UniqueIndex.IsClustered = False

            '    'Add indexed columns to the index.
            '    IndexCol = New IndexedColumn(UniqueIndex, DxDbClass.DeviceField, False)

            '    'Add indexed columns to the index.
            '    UniqueIndex.IndexedColumns.Add(IndexCol)

            '    'Create the index on the instance of SQL Server.
            '    UniqueIndex.Create()
            'End If
        End Sub

        Protected Function CreateSqlPrtnTable(ByRef SqlDB As Database, ByVal TableName As String, ByRef DaoFields As List(Of DaoField)) As Table
            Dim SqlTableDef As Table, SqlTableCol As Column

            Try
                ' create SQL tabledef
                SqlTableDef = New Table(SqlDB, TableName)

                For i As Integer = 0 To DaoFields.Count - 1
                    SqlTableCol = New Column(SqlTableDef, DaoFields(i).Name, DaoFields(i).SqlType)
                    SqlTableCol.Nullable = True

                    If DaoFields(i).Name = BaseClass.DeviceField Then
                        SqlTableCol.Nullable = False
                    End If

                    SqlTableDef.Columns.Add(SqlTableCol)
                Next
                SqlTableDef.Create()

                Return SqlTableDef
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        Protected Function CreateDataTableColumns(ByVal TableName As String, ByRef DaoFields As List(Of DaoField)) As DataTable
            Dim DatTableDef As DataTable = Nothing
            Dim DatTableCol As DataColumn

            Try
                DatTableDef = New DataTable(TableName)

                For i As Integer = 0 To DaoFields.Count - 1
                    DatTableCol = New DataColumn(DaoFields(i).Name, GetType(String))
                    DatTableCol.AllowDBNull = True
                    DatTableCol.Unique = False

                    If DaoFields(i).Name = BaseClass.DeviceField Then
                        DatTableCol.AllowDBNull = False
                        DatTableCol.Unique = True
                    End If

                    Select Case DaoFields(i).DaoType
                        Case DAODataType.dbInteger
                            DatTableCol.DataType = GetType(Integer)
                        Case DAODataType.dbDouble
                            DatTableCol.DataType = GetType(Double)
                    End Select

                    DatTableDef.Columns.Add(DatTableCol)
                Next
                Return DatTableDef
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        Protected Function CreateFieldsDB(ByVal Partition As DxDB.PdbPrtn) As List(Of DaoField)
            Dim i As Integer, FieldDb As New List(Of DaoField), Field As DaoField
            Dim p As Integer, PnPropList As New List(Of PdbProp), Prop As PdbProp

            Me.DxDbAdd2Logfile("... building FieldName Database", 15, MsgType.Txt)

            '------------------------------------------
            ' create unique propname list for partition
            For i = 0 To Partition.Parts.Count - 1
                For p = 0 To Partition.Part(i).Props.Count - 1
                    Dim Prp As PdbProp = CType(Partition.Part(i).Props(p), PdbProp)
                    If Prp.Name = "" Then
                        Me.DxDbAdd2Logfile("Part: " + Partition.Part(i).PartNo + " has empty Property Name ''", 20, MsgType.Err)
                        Continue For
                    End If
                    If Not Me.PropInList(Prp.Name, PnPropList) Then
                        Me.DxDbAdd2Logfile("Found Partition Property: '" + Prp.Name + "'", 20, MsgType.Txt)
                        PnPropList.Add(Partition.Part(i).Props(p))
                    End If
                Next
            Next

            '----------------------
            ' Build Field/Column DB
            FieldDb.Clear()

            '-------------------------------
            ' create mandatory system fields
            For i = 0 To EEFields.Count - 1
                Me.DxDbAdd2Logfile("Adding mandarory Field: '" + EEFields(i) + "'", 20, MsgType.Txt)
                Field = New DaoField(EEFields(i), EPDBPropertyType.epdbPropTypeString) ' FieldName, FieldType
                FieldDb.Add(Field)
            Next

            '------------------------------------
            ' create fields for custom properties
            For i = 0 To PnPropList.Count - 1
                Prop = CType(PnPropList(i), PdbProp)
                If Not Me.FieldExist(Prop.Name, FieldDb) Then
                    Field = New DaoField(Prop.Name, Prop.PdbType) ' FieldName, FieldType
                    Me.DxDbAdd2Logfile("Adding custom Field: '" + Prop.Name + "'", 20, MsgType.Txt)
                    FieldDb.Add(Field)
                End If
            Next

            Return FieldDb
        End Function

        Protected Function PropInList(ByVal PropName As String, ByVal PnPropList As List(Of PdbProp)) As Boolean
            Dim i As Integer, Prop As PdbProp
            For i = 0 To PnPropList.Count - 1
                Prop = CType(PnPropList(i), PdbProp)
                If UCase(Prop.Name) = UCase(PropName) Then Return True
            Next
            Return False
        End Function

        Protected Function FieldExist(ByVal FieldName As String, ByVal FieldDb As List(Of DaoField)) As Boolean
            For i As Integer = 0 To FieldDb.Count - 1
                If FieldName.ToUpper = FieldDb(i).Name.ToUpper Then Return True
            Next
            Return False
        End Function

        Protected Function RemFileExt(ByVal FileName As String) As String
            Dim FileNameNoExt As String = Path.GetFileNameWithoutExtension(FileName)
            Dim DirPath As String = Path.GetDirectoryName(FileName)
            If DirPath = String.Empty Then Return FileNameNoExt
            Return DirPath + "\" + FileNameNoExt
        End Function
#End Region

#Region "DAO Field Class"
        Protected Friend Class DaoField
            Public Name As String
            Public DaoType As DAODataType
            Public SqlType As DataType

            Public Sub New(ByVal fName As String, ByVal fPdbType As EPDBPropertyType)
                Me.Name = fName
                'set field datatype
                Select Case fPdbType
                    Case EPDBPropertyType.epdbPropTypeInt
                        Me.DaoType = DAODataType.dbInteger
                        Me.SqlType = DataType.Int
                    Case EPDBPropertyType.epdbPropTypeReal
                        Me.DaoType = DAODataType.dbDouble
                        Me.SqlType = DataType.Real
                    Case EPDBPropertyType.epdbPropTypeString
                        Me.DaoType = DAODataType.dbText
                        Me.SqlType = DataType.VarChar(DxDB.SqlSvr.MaxStrLen)
                    Case Else
                        Me.DaoType = DAODataType.dbText
                        Me.SqlType = DataType.VarChar(DxDB.SqlSvr.MaxStrLen)
                End Select
            End Sub
        End Class
#End Region

    End Class
End Namespace
