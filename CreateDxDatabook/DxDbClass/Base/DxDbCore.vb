Imports System.IO
Imports System.Reflection
Imports System.Windows.Forms
Imports DxDbClass
Imports LibraryManager

Public Class DxDbCore

    ' https://docs.microsoft.com/en-us/visualstudio/ide/how-to-configure-projects-to-target-platforms?view=vs-2019

#Region "DxDBCore Events"
    Public Event StatusbarMsg(ByVal Msg As String)
    Public Event ProgBarAction(ByVal Value As Integer, ByVal Mode As ProgbarMode)
    Public Event SysException(ByVal UserMsg As String, ByVal Ex As System.Exception, ByVal LeadingSpaces As Integer)
    Public Event Add2Logfile(ByVal Msg As String, ByVal LeadingSpaces As Integer, ByVal MsgType As MsgType)
    Public Event TranscriptMsg(ByVal Msg As String, ByVal LeadingSpaces As Integer, ByVal MsgType As MsgType, ByVal AddLogMsg As Boolean)
#End Region

#Region "Declarations"
    ' Tool Config Class
    Private WithEvents O_CfgXml As XmlCfg

    Private ComVersion As Integer
    Private WithEvents PDB As DxDB.Pdb
    Private WithEvents SFDB As DxDB.SideFileDB
    Private WithEvents DB3 As DxDB.SqliteADO
    Private WithEvents DBC As DxDB.Dbc

    Private LmcFile As String
    Private DbcFile As String
    Private AllowDplPartNos As Boolean
    Private UseSymbolTable As Boolean
    Private DbcCfgInLibDir As Boolean
    Private UseOdbcAlias As Boolean
    Private OdbcAliasName As String
    Private UseCellPins As Boolean
    Private UseAlternateDbFolder As Boolean
    Private UsePrpSideFile As Boolean
    Private PrpFileDelimiter As String
#End Region

#Region "Properties"
    Public ReadOnly Property CfgXml As XmlCfg
        Get
            Return Me.O_CfgXml
        End Get
    End Property

    Public ReadOnly Property PdbValid As Boolean
        Get
            Return (Not IsNothing(Me.PDB))
        End Get
    End Property
#End Region

#Region "Constructor"
    Public Sub New(nComVersion As Integer)
        '-----------------------------
        ' create Config class instance
        ' ini file path
        Dim IniFile As String = My.Application.Info.DirectoryPath
        IniFile = IniFile + "\" + My.Application.Info.AssemblyName + ".ini"

        ' read configuration file
        Me.O_CfgXml = New XmlCfg(IniFile)

        Me.ComVersion = nComVersion
        Me.LmcFile = Me.CfgXml.GetIniCentLibLmcFile()
        Me.AllowDplPartNos = Me.CfgXml.GetIniAllowDuplPartnos()
        Me.UseSymbolTable = (Not Me.CfgXml.GetIniUseCentLinSymbols())
        Me.DbcCfgInLibDir = Me.CfgXml.GetIniDbcInLibDir()
        Me.UseOdbcAlias = Me.CfgXml.GetIniUseOdbcAlias()
        Me.OdbcAliasName = Me.CfgXml.GetIniOdbcAliasName()
        Me.UseCellPins = Me.CfgXml.GetIniUseCellPins()
        Me.UseAlternateDbFolder = Me.CfgXml.GetIniUseOtherMdbDir()

        Me.UsePrpSideFile = Me.CfgXml.GetIniMdbImpPrpEnable()
        Me.PrpFileDelimiter = Me.GetPnPrpFieldDelimiter()

        '----------------------------
        ' create new PartsDB Instance
        Me.PDB = New DxDB.Pdb(Me.LmcFile, nComVersion)

        '----------------------------
        ' create new PartsDB Instance
        Me.SFDB = New DxDB.SideFileDB(Me.ComVersion)

        Me.DB3 = Nothing
        Me.DBC = Nothing
    End Sub
#End Region

#Region "Public Methods"
    Public Sub UpdateConfigParameters()
        Me.LmcFile = Me.CfgXml.GetIniCentLibLmcFile()
        Me.AllowDplPartNos = Me.CfgXml.GetIniAllowDuplPartnos()
        Me.UseSymbolTable = (Not Me.CfgXml.GetIniUseCentLinSymbols())
        Me.DbcCfgInLibDir = Me.CfgXml.GetIniDbcInLibDir()
        Me.UseOdbcAlias = Me.CfgXml.GetIniUseOdbcAlias()
        Me.OdbcAliasName = Me.CfgXml.GetIniOdbcAliasName()
        Me.UseCellPins = Me.CfgXml.GetIniUseCellPins()
        Me.UseAlternateDbFolder = Me.CfgXml.GetIniUseOtherMdbDir()
        Me.UsePrpSideFile = Me.CfgXml.GetIniMdbImpPrpEnable()
        Me.PrpFileDelimiter = Me.GetPnPrpFieldDelimiter()
    End Sub

    Public Sub ReadPartsDatabase()
        Me.PDB.ReadPDB(Me.UseSymbolTable, Me.UseCellPins)
    End Sub

    Public Sub ReadPropertySideFile(SideFile As String)
        If Not IsNothing(Me.PDB) AndAlso Me.UsePrpSideFile Then
            If Me.SFDB.ReadPropsFile(SideFile, Me.PrpFileDelimiter) Then
                Me.PDB.MergePropsDb(Me.SFDB)
            End If
        End If
    End Sub

    Public Function CreateDatabase(Optional AlternateDbFolder As String = "") As Boolean
        Dim DxDbOk As Boolean

        '-----------------------------
        ' creating Sqlite DxD Databook
        Me.DB3 = New DxDB.SqliteADO(Me.LmcFile, Me.AllowDplPartNos, Me.UseSymbolTable, Me.UseCellPins, Me.ComVersion)

        If Me.UseAlternateDbFolder AndAlso Directory.Exists(AlternateDbFolder) Then
            DxDbOk = Me.DB3.Create(Me.PDB, AlternateDbFolder)
        Else
            DxDbOk = Me.DB3.Create(Me.PDB)
        End If

        Me.DB3 = Nothing
        GC.Collect()

        Return DxDbOk
    End Function

    Public Sub ProcessDxDbDbcFile(ByRef ParentForm As Form, Optional AlternateDbFolder As String = "")
        Dim Result As DbcAction
        Dim DbcDir As String, DbFile As String

        ' read hkp data
        Result = DbcAction.DbcOverride

        If Me.UseAlternateDbFolder AndAlso My.Computer.FileSystem.DirectoryExists(AlternateDbFolder) Then
            DbcDir = AlternateDbFolder
            If Me.DbcCfgInLibDir Then
                DbcDir = Utils.GetBaseName(Me.LmcFile)
            End If
        Else
            DbcDir = Utils.GetBaseName(Me.LmcFile)
        End If

        DbFile = DbcDir + "\" + Path.GetFileNameWithoutExtension(Me.LmcFile) + "." + DbExtension.db3.ToString
        Me.DbcFile = Utils.RemFileExt(DbFile) + "." + DbExtension.db3.ToString + ".dbc"
        Me.DbcFile = Utils.RemFileExt(DbFile) + ".dbc"

        If File.Exists(Me.DbcFile) Then
            Result = DlgDbcCfg.CreateAndShow(ParentForm.Location, ParentForm.Size)
        End If

        Me.DBC = New DxDB.Dbc(Me.LmcFile, Me.DbcFile, Me.ComVersion, Me.PDB)

        Select Case Result
            Case DbcAction.DbcOverride
                If File.Exists(Me.DbcFile) Then
                    RaiseEvent TranscriptMsg("Overriding DxDatabook Configuration File", 5, MsgType.Txt, True)
                Else
                    RaiseEvent TranscriptMsg("Writing DxDatabook Configuration File", 5, MsgType.Txt, True)
                End If

                ' read hkp data
                If IsNothing(Me.DBC.LmcPdb) Then
                    RaiseEvent TranscriptMsg("... reading PDB data", 5, MsgType.Txt, False)
                    Me.DBC.LmcPdb.ReadPDB(Me.UseSymbolTable, Me.UseCellPins)
                End If

                RaiseEvent TranscriptMsg("... creating DBC config", 5, MsgType.Txt, False)
                If IsNothing(Me.DBC.DbcDoc) Then
                    RaiseEvent TranscriptMsg("Failed! ... creating DBC config", 5, MsgType.Wrn, False)
                    Exit Sub
                End If

                If Me.UseOdbcAlias And Not Me.OdbcAliasName = XmlCfg.KeyDefs.Undefined.ToString Then
                    Me.DBC.CreateDatabaseDbcCfg(DbFile, Me.UseSymbolTable, Me.AllowDplPartNos, Me.UseOdbcAlias, Me.OdbcAliasName)
                Else
                    Me.DBC.CreateDatabaseDbcCfg(DbFile, Me.UseSymbolTable, Me.AllowDplPartNos)
                End If

                RaiseEvent TranscriptMsg("DBC File: '" + Utils.TruncatePath(Me.DbcFile) + "' successfully created.", 5, MsgType.Txt, True)
                RaiseEvent TranscriptMsg("", 0, MsgType.Txt, True)
            Case DbcAction.DbcUpdate
                RaiseEvent TranscriptMsg("Updating DxDatabook Configuration File", 5, MsgType.Txt, True)

                If Not Me.DBC.LoadDbcConfig() Then
                    RaiseEvent TranscriptMsg("Failed! ... loading DBC config", 10, MsgType.Wrn, False)
                    Exit Sub
                End If

                ' read hkp data
                If IsNothing(Me.DBC.LmcPdb) Then
                    RaiseEvent TranscriptMsg("... reading PDB data", 10, MsgType.Txt, False)
                    Me.DBC.LmcPdb.ReadPDB(Me.UseSymbolTable, Me.UseCellPins)
                End If

                If IsNothing(Me.DBC.DbcDoc) Then
                    RaiseEvent TranscriptMsg("Failed! ... updating DBC config", 10, MsgType.Wrn, False)
                    Exit Sub
                End If

                If Me.UseOdbcAlias And Not Me.OdbcAliasName = XmlCfg.KeyDefs.Undefined.ToString Then
                    Me.DBC.UpdateDatabaseDbcCfg(DbFile, Me.UseSymbolTable, Me.AllowDplPartNos, Me.UseOdbcAlias, Me.OdbcAliasName)
                Else
                    Me.DBC.UpdateDatabaseDbcCfg(DbFile, Me.UseSymbolTable, Me.AllowDplPartNos, False, String.Empty)
                End If

                RaiseEvent TranscriptMsg("DBC File: '" + Utils.TruncatePath(Me.DbcFile) + "' successfully updated.", 5, MsgType.Txt, True)
                RaiseEvent TranscriptMsg("", 0, MsgType.Txt, True)
            Case Else
                Exit Sub
        End Select

        Me.DBC.Save()
    End Sub

    Public Sub UpdateLibraryCPD()
        Dim LibMgr As LibVxApp
        Dim CpdFile As String = Utils.RemFileExt(Me.LmcFile) + ".cpd"
        Dim DbcSave As String = Utils.RemFileExt(Me.DbcFile) + ".sav"

        If Me.ComVersion <= 76 Then Exit Sub

        LibMgr = New LibVxApp()
        'Dim ListLibMgr As List(Of LibraryManagerApp) = LibMgr.GetLibMgrInstances()

        If File.Exists(DbcSave) Then File.Delete(DbcSave)
        File.Copy(Me.DbcFile, DbcSave)

        RaiseEvent TranscriptMsg("Updating Library CPD File", 5, MsgType.Txt, False)
        RaiseEvent TranscriptMsg(Utils.TruncatePath(CpdFile), 10, MsgType.Txt, False)

        If LibMgr.CloseExistingInstance() Then
            LibMgr.CreateApplication(False, Me.LmcFile)
            If LibMgr.IsAppValid Then
                RaiseEvent TranscriptMsg("... updating CPD database", 10, MsgType.Txt, False)
            End If
            LibMgr.Close()
            RaiseEvent TranscriptMsg("Done", 5, MsgType.Txt, False)
        Else
            RaiseEvent TranscriptMsg("Update CPD database failed! Instance of Library Manager already open ...", 10, MsgType.Txt, False)
        End If

        RaiseEvent TranscriptMsg("", 0, MsgType.Txt, False)

        If File.Exists(DbcSave) AndAlso File.Exists(Me.DbcFile) Then File.Delete(Me.DbcFile)
        File.Copy(DbcSave, Me.DbcFile)
        File.Delete(DbcSave)
    End Sub
#End Region

#Region "Get COM Version shared Method"
    Public Shared Function GetClassCOMVersion() As Integer
        Dim AN As AssemblyName = GetType(DxDbCore).Assembly.GetName()
        Dim buff As String() = DxDbCore.SplitNoEmpties(AN.Name, DxDbCore.GetClassRootNamespace())
        Return CInt(buff(0))
    End Function

    Private Shared Function GetClassRootNamespace() As String
        Return GetType(DxDbCore).Namespace
    End Function

    Private Shared Function SplitNoEmpties(Input As String, Delimiter As String) As String()
        Dim arrSplit As String() = Split(Input, Delimiter)
        Return arrSplit.Where(Function(s) Not String.IsNullOrWhiteSpace(s)).ToArray()
    End Function
#End Region

#Region "Private Methods"
    Private Function GetPnPrpFieldDelimiter() As String
        Dim Delim As String = Me.CfgXml.GetIniMdbImpFldDelim()
        If Delim = "TAB" Then Return vbTab
        Return Delim
    End Function
#End Region

#Region "PartsDb, PropsDb, Dao360 Events"
    Private Sub DxDB_StatusbarMsg(ByVal Msg As String) Handles PDB.StatusbarMsg, SFDB.StatusbarMsg, DB3.StatusbarMsg, DBC.StatusbarMsg
        RaiseEvent StatusbarMsg(Msg)
    End Sub

    Private Sub DxDB_ProgressBarAction(ByVal Value As Integer, ByVal Mode As ProgbarMode) Handles PDB.ProgBarAction, SFDB.ProgBarAction, DB3.ProgBarAction, DBC.ProgBarAction
        RaiseEvent ProgBarAction(Value, Mode)
    End Sub

    Private Sub DxDB_SysException(ByVal UserMsg As String, ByVal Ex As System.Exception, ByVal LeadingSpaces As Integer) Handles PDB.SysException, SFDB.SysException, DB3.SysException, DBC.SysException
        RaiseEvent SysException(UserMsg, Ex, LeadingSpaces)
    End Sub

    Private Sub DxDB_TranscriptMsg(ByVal Msg As String, ByVal LeadingSpaces As Integer, ByVal MsgType As MsgType, ByVal AddLogMsg As Boolean) Handles PDB.TranscriptMsg, SFDB.TranscriptMsg, DB3.TranscriptMsg, DBC.TranscriptMsg
        RaiseEvent TranscriptMsg(Msg, LeadingSpaces, MsgType, AddLogMsg)
    End Sub

    Private Sub DxDB_Add2Logfile(ByVal Msg As String, ByVal LeadingSpaces As Integer, ByVal MsgType As MsgType) Handles PDB.Add2Logfile, SFDB.Add2Logfile, DB3.Add2Logfile, DBC.Add2Logfile
        RaiseEvent Add2Logfile(Msg, LeadingSpaces, MsgType)
    End Sub

    Private Sub CfgXml_Notify(Msg As String, LeadingSpaces As Integer, Type As MsgType) Handles O_CfgXml.Notify
        RaiseEvent TranscriptMsg(Msg, LeadingSpaces, Type, False)
    End Sub
#End Region

End Class
