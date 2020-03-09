Imports System.Globalization
Imports System.Threading
Imports System.Data.SQLite

Imports MGCPCBPartsEditor
Imports DxDbClass
Imports System.IO

Namespace DxDB
    Public Class SqliteADO
        Inherits BaseClass

        ' http://zetcode.com/db/sqlitevb/dataset/
        ' http://dhamen.com/index.php/2017/12/create-sqlite-database-programmatically-using-vb-net/
        ' https://www.codeproject.com/Questions/468654/VB-net-and-SQLite-Best-way-to-write-data-to-db

#Region "SQLite Members"
        Private ComVersion As Integer
        Private _DBFile As String
        Private _DBConnect As SQLiteConnection = Nothing
        Private _SqlCommand As SQLiteCommand = Nothing
        Private _SqlDataReader As SQLiteDataReader = Nothing
        Private _SqlDataSet As New DataSet
        Private dbConnString As String
#End Region

#Region "Constants & Declarations"
        Protected PnCount As Integer
        Protected LibLmcFile As String
        Protected SqlFile As String
        Protected LmcPartsDb As DxDB.Pdb
        Protected AllowDuplPartNos As Boolean
        Protected AddCellPinCount As Boolean
        Protected UseSymbolTable As Boolean
#End Region

#Region "Properties"
        Public ReadOnly Property DBFile() As String
            Get
                Return Me._DBFile
            End Get
        End Property

        Public ReadOnly Property DBConnect() As SQLiteConnection
            Get
                Return Me._DBConnect
            End Get
        End Property
#End Region

#Region "Public Methods"
        Public Sub New(LmcFile As String, AllowDplPartNos As Boolean, UseSymTable As Boolean, UseCellPins As Boolean, nComVersion As Integer)
            MyBase.New(nComVersion)

            Me.ComVersion = nComVersion

            Me._DBConnect = New SQLiteConnection()

            Me.PnCount = 0
            Me.LibLmcFile = LmcFile
            Me.SqlFile = String.Empty
            Me.AllowDuplPartNos = AllowDplPartNos
            Me.AddCellPinCount = UseCellPins
            Me.UseSymbolTable = UseSymTable
        End Sub

        Public Function Create(PartsDb As DxDB.Pdb, Optional OtherMdbDir As String = "") As Boolean
            Dim SqlDir As String, SqlFile As String

            Me.LmcPartsDb = PartsDb

            Try
                If Me.LmcPartsDb.Partitions.Count = 0 Then Return False

                ' set CultureInfo to "en-US"
                Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US", False)

                '----------------------
                ' database name/file
                If My.Computer.FileSystem.DirectoryExists(OtherMdbDir) Then
                    SqlDir = OtherMdbDir
                Else
                    SqlDir = Path.GetDirectoryName(Me.LibLmcFile)
                End If

                If Not My.Computer.FileSystem.DirectoryExists(SqlDir) Then
                    Me.DxDbTranscriptMsg("Specified MDB Output Folder does NOT exist.", 0, MsgType.Err)
                    Me.DxDbTranscriptMsg(SqlDir)
                    Me.DxDbTranscriptMsg("")
                    Return False
                End If

                SqlFile = SqlDir + "\" + Path.GetFileNameWithoutExtension(Me.LibLmcFile) + "." + DbExtension.db3.ToString
                If My.Computer.FileSystem.FileExists(SqlFile) Then
                    My.Computer.FileSystem.DeleteFile(SqlFile, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                End If

                Me.CreateDatabase(SqlFile)

                ' create Sqlite database
                Me.CreateSqliteDbTables()

                If My.Computer.FileSystem.FileExists(SqlFile) Then
                    Me.DxDbTranscriptMsg("Partnumbers added: " + CStr(PnCount), 5)
                    Me.DxDbTranscriptMsg("")
                    Me.DxDbTranscriptMsg("Database: '" + Utils.TruncatePath(SqlFile) + "' successfully created.", 5)
                    Me.DxDbTranscriptMsg("")
                Else
                    Me.DxDbTranscriptMsg("Create DxDatabook failed.", 5, MsgType.Err)
                    Me.DxDbTranscriptMsg("")
                End If
            Catch ex As Exception
                Return False
            End Try

            Return True
        End Function
#End Region

#Region "Protected Methods"
        Protected Function CreateDatabase(DatabaseFile As String) As Boolean
            Try
                If File.Exists(DatabaseFile) Then
                    File.Delete(DatabaseFile)
                End If

                SQLiteConnection.CreateFile(DatabaseFile)

                If File.Exists(DatabaseFile) Then
                    Me.OpenDatabase(DatabaseFile)
                End If
            Catch ex As Exception
                Return False
            End Try

            Return True
        End Function

        Protected Sub CreateSqliteDbTables()
            Try
                If Not IsNothing(Me.DBConnect) Then
                    Me.DxDbTranscriptMsg("Creating SQLite Database", 5)

                    ' creating databook
                    Me.CreatePartitionTables()

                    Me.DxDbTranscriptMsg("Creating SQLite Database done", 5)
                    Me.DxDbTranscriptMsg("")
                End If
            Catch ex As Exception
                Me.DxDbSysException("CreateSqliteDb(): '" + Me.DBFile + "'", ex, 10)
            End Try
        End Sub

        Protected Function OpenDatabase(DatabaseFile As String) As Boolean
            Try
                If Me._DBConnect.State = ConnectionState.Closed Then
                    Me.dbConnString = "Data Source=" + DatabaseFile + ";Version=3;New=True;Compress=True;Synchronous=Off"
                    Me._DBConnect.ConnectionString = Me.dbConnString
                    Me._SqlDataSet.Locale = CultureInfo.InvariantCulture
                    Me._DBConnect.Open()
                    Me._SqlDataSet.Reset()
                    Me._DBFile = DatabaseFile
                End If
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

        Protected Sub CreatePartitionTables()
            Dim DxDbLibName As String
            Dim DbTableName As String
            Dim SymTableDef As DataTable
            Dim PrtnTableDef As DataTable
            Dim Partition As DxDB.PdbPrtn
            Dim TblFields As List(Of SqlField)

            If Me.UseSymbolTable Then
                Me.DxDbTranscriptMsg("... creating Table '" + BaseClass.SymTableName + "'", 10)
                SymTableDef = Me.CreateSymbolTable()
            Else
                SymTableDef = Nothing
            End If

            '--------------------------------------
            ' create empty table for each partition
            For Each Partition In Me.LmcPartsDb.Partitions
                ' set tablename to partition name
                DxDbLibName = Utils.RemFileExt(Partition.Name)
                DbTableName = Utils.MakeValidDbTableName(DxDbLibName)

                Me.DxDbTranscriptMsg("... creating Table '" + DbTableName + "' (Library Name=" + DxDbLibName + ")", 10)

                ' building fieldname database
                TblFields = Me.CreateTableColumnList(Partition)
                If TblFields.Count = 0 Then Continue For

                ' create empty table for partition
                PrtnTableDef = Me.CreatePartitionTable(DbTableName, TblFields)

                Me.DxDbTranscriptMsg("... adding " + CStr(Partition.Parts.Count) + " Partnumber Record(s)", 15)
                Me.CreatePartNoRecords(PrtnTableDef, SymTableDef, Partition.Parts)

                Me.UpdateTable(PrtnTableDef)

                If Me.UseSymbolTable Then
                    Me.DxDbTranscriptMsg("... creating Query-View for Table '" + DbTableName + "'", 15)
                    'Me.CreatePartitionTableView(DbTableName)
                End If

                If Not IsNothing(SymTableDef) Then Me.UpdateTable(SymTableDef)
            Next
        End Sub

        Protected Function CreateSymbolTable() As DataTable
            Dim SymTableDef As DataTable

            If IsNothing(Me.DBConnect) Then Return Nothing

            Dim TblFields As New List(Of SqlField) From {
                New SqlField(BaseClass.DeviceField, EPDBPropertyType.epdbPropTypeString),
                New SqlField(BaseClass.SymbolField, EPDBPropertyType.epdbPropTypeString)
            }

            Try
                SymTableDef = Me.CreateSqliteTable(BaseClass.SymTableName, TblFields)
            Catch Ex As System.Exception
                Me.DxDbSysException("CreateSymbolTable(): '" + BaseClass.SymTableName + "'", Ex, 5)
                Return Nothing
            End Try

            Return SymTableDef
        End Function

        Protected Function CreateSqliteTable(TableName As String, TableFields As List(Of SqlField), Optional CloseDB As Boolean = False) As DataTable
            Dim DB3TableDef As DataTable
            Dim SQLCmd As String = String.Empty

            ' create table sql statement (table cannot be empty)
            SQLCmd += "CREATE TABLE IF NOT EXISTS [" + TableName + "] ("

            If Not IsNothing(Me.DBConnect) Then
                Try
                    If Not Me._DBConnect.State = ConnectionState.Open Then
                        Me.DBConnect.Open()
                    End If

                    ' create columns
                    For Each F As SqlField In TableFields

                        If F.Name.ToLower = "part number" Then
                            If Me.UseSymbolTable And TableName = BaseClass.SymTableName Then
                                'SQLCmd += "[SymbolId] INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL"
                                'SQLCmd += ", [" + F.Name + "] " + F.SqlType
                                SQLCmd += "[" + F.Name + "] " + F.SqlType + " NOT NULL"
                            Else
                                If Me.AllowDuplPartNos Then
                                    SQLCmd += "[" + F.Name + "] " + F.SqlType + " NOT NULL"
                                Else
                                    SQLCmd += "[" + F.Name + "] " + F.SqlType + " PRIMARY KEY NOT NULL"
                                End If
                            End If
                            Continue For
                        End If

                        ' other fields
                        ' Datatypes: TEXT INTEGER REAL
                        SQLCmd += ", [" + F.Name + "] " + F.SqlType
                    Next

                    ' create Relationship
                    If Me.UseSymbolTable And Not TableName = BaseClass.SymTableName Then
                        SQLCmd += ", [" + BaseClass.SymbolField + "] TEXT"
                        SQLCmd += ", FOREIGN KEY([" + BaseClass.DeviceField + "]) REFERENCES  " + BaseClass.SymTableName + "([" + BaseClass.DeviceField + "])"
                        'SQLCmd += ", [Symbol_" + BaseClass.DeviceField + "] TEXT FOREIGN KEY([" + BaseClass.DeviceField + "]) REFERENCES  " + BaseClass.SymTableName + "([" + BaseClass.DeviceField + "])"
                        'SQLCmd += ", INNER JOIN [" + BaseClass.SymbolField + "] ON Symbols.[" + BaseClass.DeviceField + "] = Symbol"
                    End If

                    SQLCmd += ");"

                    Using cmd As New SQLiteCommand(SQLCmd, Me._DBConnect)
                        cmd.ExecuteNonQuery()
                    End Using

                    DB3TableDef = Me.GetDataTable(TableName)

                    If CloseDB AndAlso Me._DBConnect.State = ConnectionState.Open Then
                        Me.DBConnect.Close()
                    End If

                    Return DB3TableDef
                Catch ex As Exception
                    Return Nothing
                End Try
            Else
                Return Nothing
            End If
        End Function

        Public Function GetDataTable(TableName As String, Optional CloseDB As Boolean = False) As DataTable
            Dim Table As DataTable = New DataTable()
            Dim SqlCommand As SQLiteCommand
            Dim reader As SQLiteDataReader

            Try
                If Not Me._DBConnect.State = ConnectionState.Open Then
                    Me._DBConnect.Open()
                End If

                SqlCommand = New SQLiteCommand(Me._DBConnect)
                SqlCommand.CommandText = "SELECT * FROM '" + TableName + "'"
                reader = SqlCommand.ExecuteReader()
                Table.Load(reader)
                reader.Close()

                If CloseDB AndAlso Me._DBConnect.State = ConnectionState.Open Then
                    Me.DBConnect.Close()
                End If

                Return Table
            Catch e As Exception
                Throw New Exception(e.Message)
                Return Nothing
            End Try

            Return Table
        End Function

        Public Function UpdateTable(TableDef As DataTable) As Boolean
            If IsNothing(Me.DBConnect) Then
                Me.DxDbTranscriptMsg("UpdateTable: Database Connection does not exists!", 0, MsgType.Err)
                Return False
            End If

            Try
                If Not Me._DBConnect.State = ConnectionState.Open Then
                    Me._DBConnect.Open()
                End If

                Dim SQL As String = "SELECT * FROM '" + TableDef.TableName + "'"
                Using da As New SQLiteDataAdapter(SQL, Me.DBConnect)
                    da.FillLoadOption = LoadOption.PreserveChanges
                    Using cb As New SQLiteCommandBuilder(da) '  Using ds As New DataSet
                        da.Fill(TableDef)
                        da.Update(TableDef)
                    End Using
                End Using


                For Each Row As DataRow In TableDef.Rows
                    Row = Row
                Next
                Me.DBConnect.Close()

                Return True
            Catch ex As Exception
                Me.DxDbTranscriptMsg("UpdateTable: " + ex.Message + "!", 0, MsgType.Err)
                Return False
            End Try

        End Function

        Private Sub CreatePartitionTableView(TableName As String)
            Dim SQLCmd As String = String.Empty
            'CREATE VIEW Q_Analog
            'AS 
            'SELECT
            '    *,
            '    Symbols.Symbol AS Symbol
            'FROM
            '    Analog
            'INNER JOIN Symbols ON Symbols.[Part Number] = Analog.[Part Number]
            'GROUP BY Analog.[Part Number];
            ' create table sql statement (table cannot be empty)

            SQLCmd += "CREATE VIEW IF NOT EXISTS [Q_" + TableName + "] AS SELECT *"
            SQLCmd += ", Symbols.Symbol AS Symbol FROM " + TableName
            SQLCmd += " INNER JOIN " + BaseClass.SymTableName + " ON " + BaseClass.SymTableName + ".[Part Number] = " + TableName + ".[Part Number]"
            SQLCmd += " GROUP BY " + TableName + ".[Part Number];"

            If Not IsNothing(Me.DBConnect) Then
                Using cmd As New SQLiteCommand(SQLCmd, Me._DBConnect)
                    cmd.ExecuteNonQuery()
                End Using
            End If
        End Sub

        Private Sub CreatePartNoRecords(ByRef PrtnTableDef As DataTable, ByRef SymTableDef As DataTable, PartNos As List(Of PdbPart))
            Dim i As Integer = 0, PartNo As PdbPart

            Try
                Me.DxDbProgBarAction(PartNos.Count, ProgbarMode.Init)
                For Each PartNo In PartNos
                    ' add record for each partnumber
                    i += 1 : Me.PnCount += 1
                    Me.DxDbStatusbarMsg("(" & i + 1 & "," & PartNos.Count & ") " & PartNo.PartNo)

                    Me.AddPartNoRecord(PrtnTableDef, SymTableDef, PartNo)
                    Me.DxDbProgBarAction(0, ProgbarMode.Incr)
                Next
                Me.DxDbProgBarAction(0, ProgbarMode.Close)
            Catch Ex As System.Exception
                Me.DxDbSysException("CreatePnRecords(): " + PrtnTableDef.TableName, Ex, 5)
            End Try
        End Sub

        Private Sub AddPartNoRecord(ByRef PrtnTableDef As DataTable, ByRef SymTableDef As DataTable, PartNo As PdbPart)
            Dim CellPins As String, FieldName As String
            Dim Column As DataColumn, DataRow As DataRow
            Dim Props As List(Of PdbProp), Prop As PdbProp

            FieldName = String.Empty

            Try
                Props = PartNo.Props

                '-----------------------------
                ' add entry in partition table
                DataRow = PrtnTableDef.NewRow()

                '-----------------------------
                ' add mandatory PDB properties
                For Each Column In PrtnTableDef.Columns
                    Select Case Column.ColumnName
                        Case BaseClass.DeviceField ' device 
                            DataRow.SetField(Column.ColumnName, PartNo.PartNo)
                        Case BaseClass.PartNameField ' partname 
                            DataRow.SetField(Column.ColumnName, PartNo.Name)
                        Case BaseClass.PartLabelField ' partlabel 
                            DataRow.SetField(Column.ColumnName, PartNo.Label)
                        Case BaseClass.RefDesField ' refdes 
                            DataRow.SetField(Column.ColumnName, PartNo.RefPref)
                        Case BaseClass.CellNameField ' pkg_type 
                            DataRow.SetField(Column.ColumnName, PartNo.GetDefaultCell())
                        Case BaseClass.CellPinsField ' cellpins 
                            CellPins = PartNo.GetDefaultCellPinCount()
                            If CellPins.Count > 255 Then
                                Me.DxDbTranscriptMsg("CellPin count exceeds max field lenght of 255 characters. Value truncated.", 10, MsgType.Wrn)
                                CellPins = CellPins.Substring(0, 255)
                            End If
                            DataRow.SetField(Column.ColumnName, CellPins)
                        Case BaseClass.DescriptField ' desc 
                            DataRow.SetField(Column.ColumnName, PartNo.Desc)
                        Case BaseClass.TypeField
                            DataRow.SetField(Column.ColumnName, PartNo.Type)
                    End Select
                Next

                '-------------------------
                ' add other PDB properties
                Dim Sytype As String
                For Each Prop In Props
                    For Each Column In PrtnTableDef.Columns
                        If Prop.Name.ToLower = Column.ColumnName.ToLower Then
                            Sytype = CType(Column.DataType, System.Type).Name
                            Select Case Sytype
                                Case "Int64"
                                    DataRow.SetField(Column.ColumnName, CInt(Prop.Value))
                                Case "Double"
                                    DataRow.SetField(Column.ColumnName, CDbl(Prop.Value))
                                Case "Single"
                                    DataRow.SetField(Column.ColumnName, CSng(Prop.Value))
                                Case "Decimal"
                                    DataRow.SetField(Column.ColumnName, CDec(Prop.Value))
                                Case "String"
                                    DataRow.SetField(Column.ColumnName, CStr(Prop.Value))
                            End Select
                            Exit For
                        End If
                    Next
                Next

                '--------------------
                ' add row to table
                PrtnTableDef.Rows.Add(DataRow)

                '-----------------------
                ' add PartNumber symbols
                If Me.UseSymbolTable And Not IsNothing(SymTableDef) Then
                    Me.AddPartNoSymbols(SymTableDef, PartNo)
                End If

            Catch Ex As System.Exception
                Me.DxDbSysException("AddPartRecord(): PN: " + PartNo.PartNo + ", Field: '" + FieldName.ToUpper + "'", Ex, 5)
            End Try
        End Sub

        Private Sub AddPartNoSymbols(ByRef SymTableDef As DataTable, PartNo As PdbPart)
            Dim Column As DataColumn, DataRow As DataRow

            For Each Sym As PdbSymb In PartNo.Symbs
                DataRow = SymTableDef.NewRow()

                For Each Column In SymTableDef.Columns
                    Select Case Column.ColumnName
                        Case BaseClass.DeviceField ' device 
                            DataRow.SetField(Column.ColumnName, PartNo.PartNo)
                        Case BaseClass.SymbolField ' partname 
                            DataRow.SetField(Column.ColumnName, Sym.Name)
                    End Select
                Next

                '--------------------
                ' add row to table
                SymTableDef.Rows.Add(DataRow)
            Next
        End Sub

        Private Function CreatePartitionTable(TableName As String, TblFields As List(Of SqlField)) As DataTable
            Dim PrtnTableDef As DataTable = Nothing

            If IsNothing(Me.DBConnect) Then Return Nothing

            Try
                PrtnTableDef = Me.CreateSqliteTable(TableName, TblFields)
            Catch Ex As System.Exception
                Me.DxDbSysException("CreatePrtnTable(): '" + TableName + "'", Ex, 5)
            End Try

            Return PrtnTableDef
        End Function

        Private Function CreateTableColumnList(ByVal Partition As DxDB.PdbPrtn) As List(Of SqlField)
            Dim i As Integer, Columns As New List(Of SqlField), Field As SqlField
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
            If Columns.Count > 0 Then Columns.Clear()

            '-------------------------------
            ' create mandatory system fields
            For Each EEField As String In Me.EEFields
                Me.DxDbAdd2Logfile("Adding mandarory Field: '" + EEField + "'", 20, MsgType.Txt)

                If EEField = BaseClass.CellPinsField AndAlso Me.AddCellPinCount Then
                    Field = New SqlField(EEField, EPDBPropertyType.epdbPropTypeInt) ' FieldName, FieldType
                    Columns.Add(Field)
                    Continue For
                End If

                Field = New SqlField(EEField, EPDBPropertyType.epdbPropTypeString) ' FieldName, FieldType
                Columns.Add(Field)
            Next

            '------------------------------------
            ' create fields for custom properties
            For i = 0 To PnPropList.Count - 1
                Prop = CType(PnPropList(i), PdbProp)
                If Not Me.FieldExist(Prop.Name, Columns) Then
                    Field = New SqlField(Prop.Name, Prop.PdbType) ' FieldName, FieldType
                    Me.DxDbAdd2Logfile("Adding custom Field: '" + Prop.Name + "'", 20, MsgType.Txt)
                    Columns.Add(Field)
                End If
            Next

            Return Columns
        End Function

        Private Function IsValidPartTable(ByVal PartsDb As DxDB.Pdb, ByVal TableName As String) As Boolean
            For i As Integer = 0 To PartsDb.Partitions.Count - 1
                If TableName = PartsDb.Partition(i).Name Then
                    Return True
                End If
            Next
            Return False
        End Function

        Private Function PropInList(ByVal PropName As String, ByVal PnPropList As List(Of PdbProp)) As Boolean
            Dim i As Integer, Prop As PdbProp
            For i = 0 To PnPropList.Count - 1
                Prop = CType(PnPropList(i), PdbProp)
                If UCase(Prop.Name) = UCase(PropName) Then Return True
            Next
            Return False
        End Function

        Private Function FieldExist(ByVal FieldName As String, ByVal FieldDb As List(Of SqlField)) As Boolean
            For i As Integer = 0 To FieldDb.Count - 1
                If FieldName.ToUpper = FieldDb(i).Name.ToUpper Then Return True
            Next
            Return False
        End Function
#End Region

    End Class
End Namespace
