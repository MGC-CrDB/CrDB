Imports System.IO
Imports System.Threading
Imports System.Reflection
Imports System.Globalization
Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Linq.Enumerable
Imports System.Configuration

Imports DxDbClass
Imports DxDbClass.LibVxApp

Public Class MainWindow

#Region "Declarations & Properties"
    Private Const CHARCOUNT As Integer = 80

    Private BatchMode As Boolean
    Private LASTNAVDIR As String

    Private SddActRel As VxRelease
    Private RequiredPlatform As SddOsType
    Private RequiredComVersion As Integer

    Private NUM_ERROR As Integer
    Private NUM_WRNGS As Integer
    Private NUM_NOTES As Integer

    ' Backgroundworker for multi-threading
    Private WithEvents BgWorker As BackgroundWorker

    Private WithEvents DXDB As DxDbCore
    Private WithEvents Dao360 As RegDAO

    Private StartTime As Date
    Private LogFile As String

    ' Location and Height of controls
    Private ChkUseClibSymbolsLoc As Drawing.Point
    Private ChkCrtDbcFileLoc As Drawing.Point
    Private BtnCrDxDbookLoc As Drawing.Point
    Private BtnLibManagerLoc As Drawing.Point
    Private TranscriptListHeight As Integer
    Private TranscriptListLoc As Drawing.Point
    Private LblPrpFileLoc As Drawing.Point
    Private EntryPrpFileLoc As Drawing.Point
    Private BtnSelPrpFileLoc As Drawing.Point
#End Region

#Region "BgWorker Events"
    Private Sub BtnCrDxDbook_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnCrDxDbook.Click
        Dim LibMgr As New LibVxApp()

        If LibMgr.GetExistingInstance() Then
            Me.WriteTranscript("An Instance of Library Manager is already running.", 5, Transcript.MsgType.Wrn)
            Me.WriteTranscript("Close Application and restart ...", 5)
            Me.WriteTranscript("")
            Exit Sub
        End If

        Me.StartTime = Now()
        Me.BtnCrDxDbook.Enabled = False
        Me.BtnLibManager.Enabled = False

        ' create new instance
        Me.BgWorker = New BackgroundWorker
        Me.BgWorker.RunWorkerAsync()
    End Sub

    Private Sub BgWorker_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgWorker.DoWork
        AppSetWaitCursor()
        ' set CultureInfo to "en-US"
        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US", False)
        Me.CreateDxDatabook()
        AppSetDefaultCursor()
    End Sub

    Private Sub BgWorker_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgWorker.RunWorkerCompleted
        Me.BtnCrDxDbook.Enabled = True
        Me.BtnLibManager.Enabled = True
    End Sub
#End Region

#Region "Get COM Version shared Method"
    'Public Function GetClassCOMVersion() As Integer
    '    Dim AN As AssemblyName = GetType(DxDB.Pdb).Assembly.GetName()
    '    Dim buff As String() = Me.SplitNoEmpties(AN.Name, Me.GetClassRootNamespace())
    '    Return CInt(buff(0))
    'End Function

    'Private Function SplitNoEmpties(Input As String, Delimiter As String) As String()
    '    Dim arrSplit As String() = Split(Input, Delimiter)
    '    Return arrSplit.Where(Function(s) Not String.IsNullOrWhiteSpace(s)).ToArray()
    'End Function

    'Private Function GetClassRootNamespace() As String
    '    Dim NS As String = GetType(DxDB.Pdb).Namespace
    '    Dim buff As String() = Me.SplitNoEmpties(NS, ".")
    '    Return buff(0)
    'End Function
#End Region

#Region "Form Events"
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' check application config CreateDxDatabook.exe.config
        Dim AppCfgFile As String = My.Application.Info.DirectoryPath + "\" + My.Application.Info.AssemblyName + ".exe.config"
        If Not File.Exists(AppCfgFile) Then
            DxDbClass.Utils.CreateAppExeConfig(AppCfgFile)
        End If

        Me.RequiredPlatform = SddOsType.win32
        If Environment.Is64BitProcess Then
            Me.RequiredPlatform = SddOsType.win64
        End If

        ' Add any initialization after the InitializeComponent() call.
        Dim VxEnv As New SddVxEnv(True)
        Me.RequiredComVersion = DxDbClass.DxDbCore.GetClassCOMVersion()
        If Not VxEnv.IsValid OrElse Not VxEnv.SetSddEnv(Me.RequiredComVersion) Then
            VxEnv.ShowNoReleaseMsgBox()
            End ' close application
        End If

        Me.SddActRel = VxEnv.ActiveRelease()
        If Me.CheckValidRelease() Then
            VxEnv = Nothing
        Else
            Me.ShowInvalidReleaseMsgBox(Me.RequiredComVersion, Me.RequiredPlatform)
            End ' close application
        End If

        Me.GetInititialControlsLocHeight()

        ' force culture to EN-US to avoid decimal separator problem
        System.Threading.Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture 'New System.Globalization.CultureInfo("en-US")

        '-----------------------------
        ' create new DxDbCore Instance
        Me.DXDB = New DxDbCore(Me.RequiredComVersion)
    End Sub

    Public Function CheckValidRelease() As Boolean
        If Not Me.SddActRel.ComVersion = Me.RequiredComVersion Or Not Me.SddActRel.Platform = Me.RequiredPlatform Then
            Return False
        End If
        Return True
    End Function

    Public Sub ShowInvalidReleaseMsgBox(RequiredComversion As Integer, RequiredPlatform As SddOsType)
        Dim MsgText As String = String.Empty
        MsgText = "Required Release: " + RequiredPlatform.ToString + " ComVersion=" + RequiredComversion.ToString + " does NOT match" + vbNewLine
        MsgText = MsgText + "Active Release: " + Me.SddActRel.Name + " (" + Me.SddActRel.Platform.ToString + ") ComVersion=" + Me.SddActRel.ComVersion.ToString
        MsgBox(MsgText, MsgBoxStyle.Information, My.Application.Info.ProductName + " - v" + My.Application.Info.Version.ToString())
    End Sub

    Private Sub MainWindow_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim RevStr = "Version " + My.Application.Info.Version.ToString
        Me.LASTNAVDIR = Path.GetDirectoryName(My.Application.Info.DirectoryPath)

        Me.SetStatusbarMessage(My.Application.Info.Title + " " + RevStr + " - SDD Release: " + Me.SddActRel.Name + " - SDD Platform: " + Me.SddActRel.Platform.ToString)

        '-con -lmc D:\Users\Thomass\MyDocuments\Programming\VBasic\EEtoolsVX\xCreateDxDatabook\TestData\Library\Library.lmc
        '     -psf D:\Users\Thomass\MyDocuments\Programming\VBasic\EEtoolsVX\xCreateDxDatabook\TestData\AdditionalProps.psf
        '     -mdb D:\Temp

        ' enable/disable other MDB dir file control
        Me.ToggleMdbDirControls(Me.DXDB.CfgXml.GetIniUseOtherMdbDir())

        ' enable/disable additional prop file control
        Me.TogglePropFileControls(Me.DXDB.CfgXml.GetIniMdbImpPrpEnable())

        Me.BatchMode = Me.EvaluateCmdLineArguments(Me.Text, ConsoleColor.Yellow)
        If Me.BatchMode Then
            Me.PrintVersion()

            Me.DXDB.CfgXml.PrintConfigEntries()

            Me.CreateDxDatabook()

            If Me.TranscriptList.ConsoleMode Then
                Me.WriteTranscript("")
                Me.WriteTranscript("Hit any key to continue ...")
                Console.ReadKey(True)
                AppConsole.ForgetConsole()
            End If

            End ' exit application
        End If

        ' look for open LibraryManager and get LMC file.
        Me.SetCentralLibLmcFile()

        Me.SetOtherMdbDirectory()

        Me.SetPropertySideFile()

        If Me.DXDB.CfgXml.GetIniDynamicGui() Then
            Me.SetBtnAndChkAndTranscriptTopPosition()
        End If

        Me.PrintVersion()

        Me.DXDB.CfgXml.PrintConfigEntries()

        Me.Text = My.Application.Info.AssemblyName + " - SQLite"
    End Sub

    Private Sub MainWindow_Shown(sender As Object, e As System.EventArgs) Handles Me.Shown
    End Sub

    Private Sub BtnSelLmcFile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnSelLmcFile.Click
        Me.SelLmcFile()
        If File.Exists(Me.EntryLmcFile.Text) Then
            Me.DXDB.CfgXml.SetIniCentLibLmcPath(Me.EntryLmcFile, True)
        End If
    End Sub

    Private Sub BtnSelMdbDir_Click(sender As Object, e As System.EventArgs) Handles BtnSelMdbDir.Click
        Me.SelMdbDir()
        If My.Computer.FileSystem.DirectoryExists(Me.EntryMdbDir.Text) Then
            Me.DXDB.CfgXml.SetIniOtherMdbDir(Me.EntryMdbDir, True)
        End If
    End Sub

    Private Sub BtnSelPsfFile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnSelPsfFile.Click
        Me.SelPsfFile()
        If My.Computer.FileSystem.FileExists(Me.EntryPrpFile.Text) Then
            Me.DXDB.CfgXml.SetIniPrpSideFilePath(Me.EntryPrpFile)
        End If
    End Sub

    Private Sub BtnLibManager_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnLibManager.Click
        Dim LmcFile As String

        LmcFile = Me.EntryLmcFile.Text
        If Not My.Computer.FileSystem.FileExists(LmcFile) Then
            Me.WriteTranscript("Central Library File '" + LmcFile + "' does not exist!", , Transcript.MsgType.Err)
            Me.WriteTranscript("")
            Exit Sub
        End If

        '------------------------
        ' Invoke DxLibraryManager
        Me.SetStatusbarMessage("Starting Library Manager")
        Me.RunDxLibMgr(LmcFile)

        Me.Refresh()
    End Sub

    Private Sub EntryLmcFile_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles EntryLmcFile.TextChanged
        If My.Computer.FileSystem.FileExists(CType(sender, TextBox).Text) Then
            Me.BtnLibManager.Enabled = True
            Me.BtnCrDxDbook.Enabled = True
        Else
            Me.BtnLibManager.Enabled = False
            Me.BtnCrDxDbook.Enabled = False
        End If
    End Sub

    Private Sub ItemEditConfig_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ItemEditConfig.Click
        Dim CfgSaved As Boolean = False

        If Me.ShowConfigForm(sender, e) Then
            Me.TogglePropFileControls(Me.DXDB.CfgXml.GetIniMdbImpPrpEnable())
            Me.ToggleMdbDirControls(Me.DXDB.CfgXml.GetIniUseOtherMdbDir())

            If Me.DXDB.CfgXml.GetIniDynamicGui() Then
                Me.SetBtnAndChkAndTranscriptTopPosition()
            Else
                Me.SetControlsInitialPosition()
            End If

            CfgSaved = Me.DXDB.CfgXml.SaveConfig()

            ' print config entires
            Me.ClearTranscript()
            Me.DXDB.CfgXml.PrintConfigEntries(CfgSaved)

            Me.Text = (My.Application.Info.AssemblyName + " - SQLite")
        End If
    End Sub

    Private Sub ItemAbout_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ItemAbout.Click
        My.Forms.AboutScreen.ShowDialog()
    End Sub

    Private Sub ItemExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ItemExit.Click
        End
    End Sub

    Private Sub ItemSelLmcFile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ItemSelLmcFile.Click
        Me.BtnSelLmcFile_Click(sender, e)
    End Sub

    Private Sub ItemCmdLineArgs_Click(sender As Object, e As System.EventArgs) Handles ItemCmdLineArgs.Click
        Me.PrintHelpCmdLineArgs()
    End Sub

    Private Sub ItemOtherMdbDir_Click(sender As Object, e As System.EventArgs) Handles ItemOtherMdbDir.Click
        Me.PrintHelpOtherMdbDir()
    End Sub

    Private Sub ItemDAO360Error_Click(sender As Object, e As System.EventArgs) Handles ItemDAO360Error.Click
        Me.PrintDAO360Error()
    End Sub

    Private Sub ItemPropSideFile_Click(sender As Object, e As System.EventArgs) Handles ItemPropSideFile.Click
        Me.PrintHelpPropSideFile()
    End Sub

    Private Sub ItemSysInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ItemSysInfo.Click
        Me.GetMsInfo32()
    End Sub

    Private Sub ItemChkDao360_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ItemChkDao360.Click
        Me.ClearTranscript()
        Me.Dao360 = New RegDAO
        Me.Dao360.ChkDAO360DLL()
        Me.Dao360 = Nothing
    End Sub
#End Region

#Region "Eval Argument Methods"
    Protected Function EvaluateCmdLineArguments(ByVal Title As String, ByVal TextColor As System.ConsoleColor) As Boolean
        Dim Args As New AppArgs

        ' -con -lmc <LmcFile> -prp <PropertyFile> -dbc
        ' -con -lmc D:\Users\Thomass\MyDocuments\Programming\VBasic\EEtoolsVX\xCreateDxDatabook\TestData\Library\Library.lmc
        '      -psf D:\Users\Thomass\MyDocuments\Programming\VBasic\EEtoolsVX\xCreateDxDatabook\TestData\AdditionalProps.psf
        '      -out D:\Users\Thomass\MyDocuments\Programming\VBasic\EEtoolsVX\xCreateDxDatabook\TestData

        For Each Arg As CmdArg In Args.Argv
            Select Case Arg.Name
                Case "-con" ' open console window
                    Me.TranscriptList.ConsoleMode = True
                Case "-lmc" ' path to .lmc file
                    If My.Computer.FileSystem.FileExists(Arg.Value) Then
                        Me.EntryLmcFile.Text = Arg.Value
                    End If
                Case "-psf" ' path to .psf additional property side file
                    If My.Computer.FileSystem.FileExists(Arg.Value) Then
                        Me.EntryPrpFile.Text = Arg.Value
                    End If
                Case "-out" ' path to other MDB directory
                    If My.Computer.FileSystem.DirectoryExists(Arg.Value) Then
                        Me.EntryMdbDir.Text = Arg.Value
                    End If
            End Select
        Next

        If Me.TranscriptList.ConsoleMode Then
            AppConsole.CreateConsole(AppConsole.Mode.Create, Title, TextColor)
        End If

        If Args.Argc > 0 Then Return True

        Return False
    End Function
#End Region

#Region "Form Functions"
    Protected Sub TogglePropFileControls(ByVal Enable As Boolean)
        Me.LblPrpFile.Visible = Enable
        Me.EntryPrpFile.Visible = Enable
        Me.BtnSelPsfFile.Enabled = Enable
        Me.BtnSelPsfFile.Visible = Enable
        If Enable Then
            Dim f As String = Me.DXDB.CfgXml.GetIniMdbImpPrpSideFile()
            If My.Computer.FileSystem.FileExists(f) Then
                Me.EntryPrpFile.Text = f
            End If
        End If
    End Sub

    Protected Sub ToggleMdbDirControls(ByVal Enable As Boolean)
        Me.LblMdbDir.Visible = Enable
        Me.EntryMdbDir.Visible = Enable
        Me.BtnSelMdbDir.Enabled = Enable
        Me.BtnSelMdbDir.Visible = Enable
        If Enable Then
            Dim d As String = Me.DXDB.CfgXml.GetIniOtherMdbDir()
            If My.Computer.FileSystem.DirectoryExists(d) Then
                Me.EntryMdbDir.Text = d
            End If
        End If
    End Sub

    Protected Sub GetInititialControlsLocHeight()
        Me.ChkCrtDbcFileLoc = Me.ChkCrtDbcFile.Location
        Me.BtnCrDxDbookLoc = Me.BtnCrDxDbook.Location
        Me.BtnLibManagerLoc = Me.BtnLibManager.Location
        Me.TranscriptListHeight = Me.TranscriptList.Height
        Me.TranscriptListLoc = Me.TranscriptList.Location
        Me.LblPrpFileLoc = Me.LblPrpFile.Location
        Me.EntryPrpFileLoc = Me.EntryPrpFile.Location
        Me.BtnSelPrpFileLoc = Me.BtnSelPsfFile.Location
    End Sub

    Protected Sub SetControlsInitialPosition()
        Me.ChkCrtDbcFile.Location = Me.ChkCrtDbcFileLoc
        Me.BtnCrDxDbook.Location = Me.BtnCrDxDbookLoc
        Me.BtnLibManager.Location = Me.BtnLibManagerLoc

        Me.TranscriptList.Height = Me.TranscriptListHeight
        Me.TranscriptList.Location = Me.TranscriptListLoc

        ' move PrpFile label, textbox and button to original location
        Me.LblPrpFile.Location = Me.LblPrpFileLoc
        Me.EntryPrpFile.Location = Me.EntryPrpFileLoc
        Me.BtnSelPsfFile.Location = Me.BtnSelPrpFileLoc
    End Sub

    Protected Sub SetBtnAndChkAndTranscriptTopPosition()
        Dim DbFolder As Boolean = Me.DXDB.CfgXml.GetIniUseOtherMdbDir()
        Dim PrpFile As Boolean = Me.DXDB.CfgXml.GetIniMdbImpPrpEnable()

        ' set to roginal location
        If DbFolder And PrpFile Then
            Me.ChkCrtDbcFile.Location = Me.ChkCrtDbcFileLoc
            Me.BtnCrDxDbook.Location = Me.BtnCrDxDbookLoc
            Me.BtnLibManager.Location = Me.BtnLibManagerLoc

            Me.TranscriptList.Height = Me.TranscriptListHeight
            Me.TranscriptList.Location = Me.TranscriptListLoc

            ' move PrpFile label, textbox and button to original location
            Me.LblPrpFile.Location = New Drawing.Point(Me.LblPrpFile.Location.X, Me.LblPrpFileLoc.Y)
            Me.EntryPrpFile.Location = New Drawing.Point(Me.EntryPrpFile.Location.X, Me.EntryPrpFileLoc.Y)
            Me.BtnSelPsfFile.Location = New Drawing.Point(Me.BtnSelPsfFile.Location.X, Me.BtnSelPrpFileLoc.Y)
            Exit Sub
        End If

        If Not DbFolder And Not PrpFile Then
            Me.ChkCrtDbcFile.Location = New Drawing.Point(Me.ChkCrtDbcFile.Location.X, Me.ChkCrtDbcFileLoc.Y - 60)
            Me.BtnCrDxDbook.Location = New Drawing.Point(Me.BtnCrDxDbook.Location.X, Me.BtnCrDxDbookLoc.Y - 60)
            Me.BtnLibManager.Location = New Drawing.Point(Me.BtnLibManager.Location.X, Me.BtnLibManagerLoc.Y - 60)

            Me.TranscriptList.Height = Me.TranscriptListHeight + 75
            Me.TranscriptList.Location = New Drawing.Point(Me.TranscriptList.Location.X, Me.TranscriptListLoc.Y - 60)
            Exit Sub
        End If

        If DbFolder Then
            Me.ChkCrtDbcFile.Location = New Drawing.Point(Me.ChkCrtDbcFile.Location.X, Me.ChkCrtDbcFileLoc.Y - 30)
            Me.BtnCrDxDbook.Location = New Drawing.Point(Me.BtnCrDxDbook.Location.X, Me.BtnCrDxDbookLoc.Y - 30)
            Me.BtnLibManager.Location = New Drawing.Point(Me.BtnLibManager.Location.X, Me.BtnLibManagerLoc.Y - 30)

            Me.TranscriptList.Height = Me.TranscriptListHeight + 30
            Me.TranscriptList.Location = New Drawing.Point(Me.TranscriptList.Location.X, Me.TranscriptListLoc.Y - 30)

            ' move PrpFile label, textbox and button to original location
            Me.LblPrpFile.Location = New Drawing.Point(Me.LblPrpFile.Location.X, Me.LblPrpFileLoc.Y)
            Me.EntryPrpFile.Location = New Drawing.Point(Me.EntryPrpFile.Location.X, Me.EntryPrpFileLoc.Y)
            Me.BtnSelPsfFile.Location = New Drawing.Point(Me.BtnSelPsfFile.Location.X, Me.BtnSelPrpFileLoc.Y)
            Exit Sub
        End If

        If PrpFile Then
            Me.ChkCrtDbcFile.Location = New Drawing.Point(Me.ChkCrtDbcFile.Location.X, Me.ChkCrtDbcFileLoc.Y - 30)
            Me.BtnCrDxDbook.Location = New Drawing.Point(Me.BtnCrDxDbook.Location.X, Me.BtnCrDxDbookLoc.Y - 30)
            Me.BtnLibManager.Location = New Drawing.Point(Me.BtnLibManager.Location.X, Me.BtnLibManagerLoc.Y - 30)

            Me.TranscriptList.Height = Me.TranscriptListHeight + 30
            Me.TranscriptList.Location = New Drawing.Point(Me.TranscriptList.Location.X, Me.TranscriptListLoc.Y - 30)

            ' move PrpFile textbox and button 30 up
            Me.LblPrpFile.Location = New Drawing.Point(Me.LblPrpFile.Location.X, Me.LblPrpFileLoc.Y - 30)
            Me.EntryPrpFile.Location = New Drawing.Point(Me.EntryPrpFile.Location.X, Me.EntryPrpFileLoc.Y - 30)
            Me.BtnSelPsfFile.Location = New Drawing.Point(Me.BtnSelPsfFile.Location.X, Me.BtnSelPrpFileLoc.Y - 30)
            Exit Sub
        End If
    End Sub

    Protected Function ShowConfigForm(ByVal sender As Object, ByVal e As System.EventArgs) As Boolean
        If IsNothing(Me.DXDB.CfgXml) Then Return False
        Dim Res As DialogResult = DialogToolCfg.CreateAndShow(Me.DXDB.CfgXml)
        If Not Res = Windows.Forms.DialogResult.OK Then Return False
        Return True
    End Function

    Delegate Function GetCrtDbcFileOptionCallback() As Boolean
    Protected Function GetCrtDbcFileOption() As Boolean
        If Me.ChkCrtDbcFile.InvokeRequired Then
            Dim d As New GetCrtDbcFileOptionCallback(AddressOf GetCrtDbcFileOption)
            Return CBool(CStr(Me.Invoke(d, New Object() {})))
        Else
            Return Me.ChkCrtDbcFile.Checked
        End If
    End Function

    Delegate Function GetCentLibLmcCallback() As String
    Protected Function GetCentLibLmc() As String
        If Me.EntryLmcFile.InvokeRequired Then
            Dim d As New GetCentLibLmcCallback(AddressOf GetCentLibLmc)
            Return CStr(Me.Invoke(d, New Object() {}))
        Else
            If My.Computer.FileSystem.FileExists(Me.EntryLmcFile.Text) Then
                Return Me.EntryLmcFile.Text
            End If
            Return String.Empty
        End If
    End Function

    Delegate Function GetOtherMdbDirCallback() As String
    Protected Function GetOtherMdbDir() As String
        If Me.EntryMdbDir.InvokeRequired Then
            Dim d As New GetOtherMdbDirCallback(AddressOf GetOtherMdbDir)
            Return CStr(Me.Invoke(d, New Object() {}))
        Else
            If Directory.Exists(Me.EntryMdbDir.Text) Then
                Return Me.EntryMdbDir.Text
            End If
            Return String.Empty
        End If
    End Function

    Delegate Function GetPnPrpFileCallback() As String
    Protected Function GetPnPrpFile() As String
        If Me.EntryPrpFile.InvokeRequired Then
            Dim d As New GetPnPrpFileCallback(AddressOf GetPnPrpFile)
            Return CStr(Me.Invoke(d, New Object() {}))
        Else
            If My.Computer.FileSystem.FileExists(Me.EntryPrpFile.Text) Then
                Return Me.EntryPrpFile.Text
            End If
            Return String.Empty
        End If
    End Function

    Protected Sub SelMdbDir()
        Dim MdbDir As String = Me.SelectAnyFolder()
        If My.Computer.FileSystem.DirectoryExists(MdbDir) Then
            Me.EntryMdbDir.Text = MdbDir
        End If
    End Sub

    Protected Sub SelLmcFile()
        Dim ClibLmcFile As String, FileDir As String
        Dim FileName As String

        FileDir = "" : FileName = ""

        ClibLmcFile = Me.EntryLmcFile.Text
        If My.Computer.FileSystem.FileExists(ClibLmcFile) Then
            FileName = Path.GetFileName(ClibLmcFile)
            FileDir = Path.GetDirectoryName(ClibLmcFile)
        Else
            FileDir = Me.LASTNAVDIR
        End If

        ClibLmcFile = Me.SelectAnyFile("Select Central Library File", "Central Library File", ".lmc", FileDir, FileName)

        If My.Computer.FileSystem.FileExists(ClibLmcFile) Then
            Me.LASTNAVDIR = Path.GetDirectoryName(ClibLmcFile)
            Me.EntryLmcFile.Text = ClibLmcFile.Clone().ToString
            Me.BtnCrDxDbook.Enabled = True
        End If
    End Sub

    Protected Sub SelPsfFile()
        Dim CmsPrpFile As String, FileDir As String
        Dim FileName As String

        FileDir = "" : FileName = ""

        CmsPrpFile = Me.EntryPrpFile.Text
        If My.Computer.FileSystem.FileExists(CmsPrpFile) Then
            FileName = Path.GetFileName(CmsPrpFile)
            FileDir = Path.GetDirectoryName(CmsPrpFile)
        ElseIf My.Computer.FileSystem.FileExists(Me.EntryLmcFile.Text) Then
            FileDir = Path.GetDirectoryName(Me.EntryLmcFile.Text)
        Else
            FileDir = Me.LASTNAVDIR
        End If

        CmsPrpFile = Me.SelectAnyFile("Select Property Side File", "Property Side File", ".psf", FileDir, FileName)

        If My.Computer.FileSystem.FileExists(CmsPrpFile) Then
            Me.LASTNAVDIR = Path.GetDirectoryName(CmsPrpFile)
            Me.EntryPrpFile.Text = CmsPrpFile.Clone().ToString
        End If
    End Sub

    Protected Function VectorizeExMsg(ByVal ex As Exception) As List(Of String)
        Dim errVect As New List(Of String), Start As Integer = 0
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

    Protected Sub GetMsInfo32()
        Dim TmpFile As String, MsInfo32 As String = ""
        Dim RegKey As Microsoft.Win32.RegistryKey
        Dim MyProcess As New Process

        Try
            ' retrieve the path to MSINFO32.EXE
            RegKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Shared Tools\MSInfo")
            If RegKey Is Nothing Then
                Throw New Exception("Registry key not found")
            End If

            MsInfo32 = RegKey.GetValue("Path", "").ToString()
            RegKey.Close()

            If MsInfo32 = "" Then
                Throw New Exception("Registry Path value not found")
            End If

            ' run MsInfo32, output the result to C:\TEMP\REPORT.TXT
            TmpFile = My.Computer.FileSystem.SpecialDirectories.Temp + "\Ms32Info.txt"

            ' assign prog and parameter
            MyProcess.StartInfo.FileName = MsInfo32
            MyProcess.StartInfo.Arguments = ""

            MyProcess.Start()
        Catch Ex As System.Exception
            Me.PrintException(Ex)
        End Try
    End Sub
#End Region

#Region "Print to Transcript Methods"
    Protected Sub PrintException(ByVal ex As Exception, Optional ByVal LeadingSpaces As Integer = 0, Optional ByVal Verbose As Boolean = False)
        Dim MsgVect As List(Of String) = Me.VectorizeExMsg(ex)
        Me.WriteTranscript(ex.GetType.Name, LeadingSpaces, Transcript.MsgType.Err, True)
        Me.WriteTranscript(ex.Message, LeadingSpaces + 5, , True)
        If Verbose Then
            For i As Integer = 0 To MsgVect.Count - 1
                Me.WriteTranscript(MsgVect(i), LeadingSpaces + 5)
                Me.AddLogMessage(MsgVect(i), LeadingSpaces + 5)
            Next
        Else
            Me.WriteTranscript(MsgVect(MsgVect.Count - 1), LeadingSpaces + 5)
            Me.AddLogMessage(MsgVect(MsgVect.Count - 1), LeadingSpaces + 5)
        End If
    End Sub

    Protected Sub PrintVersion()
        Me.ClearTranscript()
        Dim HeaderLine As String = My.Application.Info.Description + " (Version=" + My.Application.Info.Version.ToString + ")"

        Me.WriteTranscript("")
        Me.WriteTranscript(HeaderLine)
        Me.WriteTranscript("")
        Me.WriteTranscript("  Copyright (c) Mentor Graphics Corporation, 2007, All Rights Reserved.")
        Me.WriteTranscript("                       UNPUBLISHED, LICENSED SOFTWARE.")
        Me.WriteTranscript("            CONFIDENTIAL AND PROPRIETARY INFORMATION WHICH IS THE")
        Me.WriteTranscript("          PROPERTY OF MENTOR GRAPHICS CORPORATION OR ITS LICENSORS.")
        Me.WriteTranscript("")
        Me.WriteTranscript("     >> THIS SOFTWARE IS NOT OFFICIALLY SUPPORTED BY MENTOR GRAPHICS <<")
        Me.WriteTranscript("")
    End Sub

    Protected Sub PrintHelpCmdLineArgs()
        Me.ClearTranscript()
        Me.WriteTranscript("")
        Me.WriteTranscript("Usage: " + My.Application.Info.AssemblyName + " -lmc <LmcFile> [-con] [-psf <PropertyFile>] [-out <DbOutputFolder>]")
        Me.WriteTranscript("")
        Me.WriteTranscript("-lmc <LmcFile> Path to library .lmc file. This argument is mandatory", 10)
        Me.WriteTranscript("")
        Me.WriteTranscript("-psf <PropertySideFile> Path to additional property ascii file (*.psf).", 10)
        Me.WriteTranscript("Must be enabled in tool configuration", 30)
        Me.WriteTranscript("")
        Me.WriteTranscript("-out <DbOutputFolder> Path to optional Database File output folder.", 10)
        Me.WriteTranscript("Must be enabled in tool configuration", 30)
        Me.WriteTranscript("")
        Me.WriteTranscript("-con Run the tool in a console window. Silent mode if omitted.", 10)
        Me.WriteTranscript("")
        Me.WriteTranscript("Use commandline mode ONLY to update the MDB Access database.", 0, Transcript.MsgType.Nte)
        Me.WriteTranscript("")
        Me.WriteTranscript("All tool configuration must be done in interactive mode.", 6)
        Me.WriteTranscript("Excute 'File>Edit Configuration ...' to modify tool configuration.", 6)
        Me.WriteTranscript("")
        Me.WriteTranscript("Also create or update Databook configuration (.dbc) has to be done", 6)
        Me.WriteTranscript("in interactive mode.", 6)
        Me.WriteTranscript("")
    End Sub

    Protected Sub PrintHelpOtherMdbDir()
        Me.ClearTranscript()
        Me.WriteTranscript("")
        Me.WriteTranscript("Specify a folder to store the Access MDB File as an alternative")
        Me.WriteTranscript("to the the default library folder.")
        Me.WriteTranscript("The specified folder must have write access for all users.")
        Me.WriteTranscript("")
    End Sub

    Protected Sub PrintDAO360Error()
        Me.ClearTranscript()
        Me.WriteTranscript("")
        Me.WriteTranscript("If the following Error occures ...")
        Me.WriteTranscript("")
        Me.WriteTranscript(">> ERROR!! InvalidCastException")
        Me.WriteTranscript(">> Unable to cast COM object of type 'System.__ComObject' to interface type 'dao.DBEngine'.")
        Me.WriteTranscript("")
        Me.WriteTranscript("This is caused by an invalid Registration of Microsofts DAO360.dll Type Library.")
        Me.WriteTranscript("")
        Me.WriteTranscript("This can be fixed by invoking MenuItem 'File > Check DAO360.dll'")
        Me.WriteTranscript("If this does not solve the Problem you may have no write Permission to the Registry.")
        Me.WriteTranscript("")
    End Sub

    Protected Sub PrintHelpPropSideFile()
        Me.ClearTranscript()
        Me.WriteTranscript("")
        Me.WriteTranscript("Property Side File (*.psf ASCII) Format:")
        Me.WriteTranscript("")
        Me.WriteTranscript("First line defines the property names. The first column defines the matching key.")
        Me.WriteTranscript("If the matching key name (Proprty Name) is one of the following, the importer")
        Me.WriteTranscript("looks for a matching 'Part Number' value: 'PARTNO', 'PARTNUMBER', 'PART NUMBER', 'PN'")
        Me.WriteTranscript("")
        Me.WriteTranscript("Any other name of the first column indicates to the importer to look in the part property")
        Me.WriteTranscript("list for matching key. The 'Part Number' is not treated as the matching key.")
        Me.WriteTranscript("")
        Me.WriteTranscript("Valid column separators are ; @ § # * | or Tab.")
        Me.WriteTranscript("")
        Me.WriteTranscript("Example:")
        Me.WriteTranscript("Header Line -> PART NUMBER|Prop 1 Name|Prop 2 Name")
        Me.WriteTranscript("Data Line   -> 6667559|LM2576S-ADJ|Prop 2 Value")
        Me.WriteTranscript("")
    End Sub

    Protected Sub PrintHeadline(ByVal Text As String, ByVal MaxLen As Integer)
        Dim OutStr As String, TextLen As Integer, FillLen As Integer
        TextLen = Len(Text)
        FillLen = CInt((MaxLen / 2) - 1 - (TextLen / 2))
        OutStr = StrDup(FillLen, "=") + Space(1) + Text + Space(1) + StrDup(FillLen, "=")
        Me.WriteTranscript(StrDup(Len(OutStr), "="))
        Me.WriteTranscript(OutStr)
        Me.WriteTranscript(StrDup(Len(OutStr), "="))
        Me.WriteTranscript("")
    End Sub

    Protected Sub PrintSummary()
        Dim StoppTime As Date = Now()
        Dim Duration As String = Time.FormatTimeSpan(Now.Subtract(Me.StartTime))

        Me.WriteTranscript(" Started: " + Format$(StartTime, "dd.mm.yy") + ", " + Format$(StartTime.ToLocalTime, "HH:mm:ss"))
        Me.WriteTranscript("Finished: " + Format$(StoppTime, "dd.mm.yy") + ", " + Format$(StoppTime.ToLocalTime, "HH:mm:ss"))
        Me.WriteTranscript("Duration: " + Duration)
        Me.WriteTranscript("")

        ' delete logfile
        If My.Computer.FileSystem.FileExists(Me.LogFile) Then
            My.Computer.FileSystem.DeleteFile(Me.LogFile, UIOption.OnlyErrorDialogs, RecycleOption.DeletePermanently)
        End If

        ' write logfile
        Me.TranscriptList.WriteLog2File(Me.LogFile)

        Me.WriteTranscript("Log File: " + Me.TruncatePath(Me.LogFile), , , False)
        Me.WriteTranscript("", , , False)
    End Sub

    Protected Function TruncatePath(ByVal InPath As String) As String
        Dim LibDir As String, LogDir As String, FileName As String
        LibDir = Path.GetFileName(Path.GetDirectoryName(Path.GetDirectoryName(InPath)))
        LogDir = Path.GetFileName(Path.GetDirectoryName(InPath))
        FileName = Path.GetFileName(InPath)
        Return "...\" + LibDir + "\" + LogDir + "\" + FileName
    End Function
#End Region

#Region "Form Function with defined Callback"
    ' This delegate enables asynchronous calls for setting
    ' the text property on a TextBox control.
    Delegate Sub AppSetWaitCursorCallback()
    Delegate Sub AppSetDefaultCursorCallback()
    Delegate Sub InitProgressBarCallback(ByVal MaxCount As Integer)
    Delegate Sub IncrProgressBarCallback()
    Delegate Sub FinishProgressBarCallback()
    Delegate Sub SetStatusbarMessageCallback(ByVal MsgStr As String)

    Public Sub ClearTranscript(Optional ByVal ClearLog As Boolean = True)
        Me.TranscriptList.Clear(ClearLog)
    End Sub

    Public Sub AppSetWaitCursor()
        If Me.InvokeRequired Then
            Dim d As New AppSetWaitCursorCallback(AddressOf AppSetWaitCursor)
            Me.Invoke(d, New Object() {})
        Else
            Me.Cursor = Cursors.WaitCursor
        End If
    End Sub

    Public Sub AppSetDefaultCursor()
        If Me.InvokeRequired Then
            Dim d As New AppSetDefaultCursorCallback(AddressOf AppSetDefaultCursor)
            Me.Invoke(d, New Object() {})
        Else
            Me.Cursor = Cursors.Default
        End If
    End Sub

    Public Sub WriteTranscript(ByVal MsgStr As String, Optional ByVal LeadingSpaces As Integer = 0, Optional ByVal MsgType As Transcript.MsgType = Transcript.MsgType.Txt, Optional ByVal AddLogMsg As Boolean = True)
        Me.TranscriptList.Message(MsgStr, LeadingSpaces, MsgType, AddLogMsg)
    End Sub

    Public Sub AddLogMessage(ByVal TextStr As String, Optional ByVal LeadingSpaces As Integer = 0, Optional ByVal MsgType As Transcript.MsgType = Transcript.MsgType.Txt)
        Me.TranscriptList.AddLogMessage(TextStr, LeadingSpaces, MsgType)
    End Sub

    Public Sub SetStatusbarMessage(ByVal MsgStr As String)
        If Me.StatusbarMain.InvokeRequired Then
            Dim d As New SetStatusbarMessageCallback(AddressOf SetStatusbarMessage)
            Me.Invoke(d, New Object() {[MsgStr]})
        Else
            Me.MsgArea.Text = [MsgStr]
        End If
    End Sub

    Public Sub InitProgressBar(ByVal MaxCount As Integer)
        If MaxCount < 1 Then
            Exit Sub
        End If
        If Me.StatusbarMain.InvokeRequired Then
            Dim d As New InitProgressBarCallback(AddressOf InitProgressBar)
            Me.Invoke(d, New Object() {[MaxCount]})
        Else
            Me.ProgressBar.Step = 1
            Me.ProgressBar.Minimum = 0
            Me.ProgressBar.Maximum = [MaxCount]
            Me.ProgressBar.Value = 0
            Me.ProgressBar.Visible = True
        End If
    End Sub

    Public Sub IncrProgressBar()
        If Me.StatusbarMain.InvokeRequired Then
            Dim d As New IncrProgressBarCallback(AddressOf IncrProgressBar)
            Me.Invoke(d, New Object() {})
        Else
            If Me.ProgressBar.Value = Me.ProgressBar.Maximum Then
                Exit Sub
            End If
            Me.ProgressBar.PerformStep()
        End If
    End Sub

    Public Sub FinishProgressBar()
        If Me.StatusbarMain.InvokeRequired Then
            Dim d As New FinishProgressBarCallback(AddressOf FinishProgressBar)
            Me.Invoke(d, New Object() {})
        Else
            Me.ProgressBar.Value = Me.ProgressBar.Maximum
            Me.ProgressBar.Value = 0
            Me.ProgressBar.Visible = False
        End If
    End Sub
#End Region

#Region "Create DxDatabook"
    Protected Sub CreateDxDatabook()
        Dim DxDbOk As Boolean = False
        Dim AlternateDir As String = Me.GetOtherMdbDir()

        Me.ClearTranscript()

        System.Threading.Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture

        '----------------------
        ' print current library
        Me.WriteTranscript("")
        If Not My.Computer.FileSystem.FileExists(Me.GetCentLibLmc()) Then
            Me.WriteTranscript("Select a valid LMC File ...", , Transcript.MsgType.Err)
            Me.WriteTranscript("")
            Exit Sub
        Else
            Me.WriteTranscript("Central Library: " + Utils.TruncatePath(Me.GetCentLibLmc()))
            Me.WriteTranscript("")
        End If

        '----------------------------
        ' set logile
        Me.LogFile = Path.GetDirectoryName(Me.GetCentLibLmc()) + "\LogFiles\CreateDxDatabook.log"

        '-----------------------
        ' reading Parts Database
        If Not Me.DXDB.PdbValid() Then Exit Sub
        Me.DXDB.UpdateConfigParameters()

        Me.PrintHeadline("Start Creating DxDatabook (PDB Automation)", CHARCOUNT)
        Me.DXDB.CfgXml.PrintConfigEntries()
        Me.DXDB.ReadPartsDatabase()

        '-------------------------------------
        ' reading Ascii Property File PartNoDb
        Me.DXDB.ReadPropertySideFile(Me.GetPnPrpFile())

        '-----------------------------
        ' creating DxDatabook database
        DxDbOk = Me.DXDB.CreateDatabase(AlternateDir)

        '----------------
        ' create DBC File
        If DxDbOk AndAlso Me.GetCrtDbcFileOption() Then
            Me.DXDB.ProcessDxDbDbcFile(Me, AlternateDir)
        End If


        ' updating Library.cpd
        If Me.RequiredComVersion > 76 Then
            Me.DXDB.UpdateLibraryCPD()
        End If

        '--------------
        ' print summary
        Me.PrintSummary()
    End Sub

    Protected Sub RunDxLibMgr(ByVal LmcFile As String)
        Dim AppPath As String
        Dim Platform As String = Me.SddActRel.Platform.ToString

        AppPath = System.Environment.GetEnvironmentVariable("SDD_HOME") + "\lm\" + Platform + "\bin\LibraryManager.exe"
        If Not My.Computer.FileSystem.FileExists(AppPath) Then
            Me.WriteTranscript("File not found '" + AppPath + "'", , Transcript.MsgType.Wrn)
            Me.WriteTranscript("")
            Exit Sub
        End If

        If Not My.Computer.FileSystem.FileExists(LmcFile) Then Exit Sub

        Try
            Call Shell(AppPath + Chr(32) + LmcFile, AppWinStyle.NormalNoFocus, False, 0)
        Catch ex As System.Exception
            Me.PrintException(ex)
        End Try
    End Sub

    Protected Function SelectAnyFile(ByVal Title As String, Optional ByVal FileType As String = "", Optional ByVal FileExt As String = "", Optional ByVal InitialDir As String = "", Optional ByVal FileName As String = "", Optional ByVal FileMustExist As Boolean = True) As String
        Dim FileTypeStr As String, AnyFileBrowser As New OpenFileDialog

        FileTypeStr = "All files (*.*)|*.*"

        AnyFileBrowser.FileName = ""
        AnyFileBrowser.CheckFileExists = FileMustExist

        ' check title
        If Title = "" Then
            AnyFileBrowser.FileName = "Open File"
        Else
            AnyFileBrowser.Title = Title
        End If

        ' check initial directory
        If My.Computer.FileSystem.DirectoryExists(InitialDir) Then
            AnyFileBrowser.InitialDirectory = InitialDir
        Else
            AnyFileBrowser.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        End If

        AnyFileBrowser.FileName = FileName

        If FileType = "" Or FileExt = "" Then
            AnyFileBrowser.Filter = FileTypeStr
            AnyFileBrowser.FilterIndex = 1
        Else
            If FileExt.Substring(0, 1) <> "." Then
                FileExt = "." + FileExt
            End If
            AnyFileBrowser.Filter = FileType + " (*" + FileExt + ")|*" + FileExt + "|" + FileTypeStr
            AnyFileBrowser.FilterIndex = 1
        End If

        If AnyFileBrowser.ShowDialog() = DialogResult.OK Then
            Return AnyFileBrowser.FileName
        Else
            Return ""
        End If
    End Function

    Protected Function SelectAnyFolder() As String
        Dim AnyDirBrowser As New FolderBrowserDialog

        AnyDirBrowser.RootFolder = System.Environment.SpecialFolder.MyComputer

        If AnyDirBrowser.ShowDialog() = DialogResult.OK Then
            Return AnyDirBrowser.SelectedPath
        Else
            Return ""
        End If
    End Function

    Protected Function RemFileExt(ByVal FileName As String) As String
        Dim FileNameNoExt As String = Path.GetFileNameWithoutExtension(FileName)
        Dim DirPath As String = Path.GetDirectoryName(FileName)
        If DirPath = String.Empty Then Return FileNameNoExt
        Return DirPath + "\" + FileNameNoExt
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
#End Region

#Region "Get LMC/PRP File from LibMgr Method"
    Protected Sub SetCentralLibLmcFile()
        Dim LmcFile As String = String.Empty

        ' check for opne Library Manager instance
        LmcFile = DxDbClass.LibVxApp.GetLibMgrLmcFile()

        If String.IsNullOrEmpty(LmcFile) Then
            LmcFile = Me.DXDB.CfgXml.GetIniCentLibLmcFile()
        End If

        If My.Computer.FileSystem.FileExists(LmcFile) Then
            Me.EntryLmcFile.Text = LmcFile
        End If

        Me.DXDB.CfgXml.SetIniCentLibLmcPath(Me.EntryLmcFile, True)
    End Sub

    Protected Sub SetPropertySideFile()
        Dim PrpFile As String = String.Empty

        ' get config file entry
        PrpFile = Me.DXDB.CfgXml.GetIniMdbImpPrpSideFile()

        If String.IsNullOrEmpty(PrpFile) Or Not My.Computer.FileSystem.FileExists(PrpFile) Then
            ' check for open Library Manager instance
            PrpFile = DxDbClass.LibVxApp.GetCentLibPsfFile(Me.EntryLmcFile.Text)
        End If

        If My.Computer.FileSystem.FileExists(PrpFile) Then
            Me.EntryPrpFile.Text = PrpFile
        End If

        Me.DXDB.CfgXml.SetIniPrpSideFilePath(Me.EntryPrpFile)
    End Sub

    Protected Sub SetOtherMdbDirectory()
        Dim MdbDir As String = String.Empty

        MdbDir = Me.DXDB.CfgXml.GetIniOtherMdbDir()

        If My.Computer.FileSystem.DirectoryExists(MdbDir) Then
            Me.EntryMdbDir.Text = MdbDir
        End If

        Me.DXDB.CfgXml.SetIniOtherMdbDir(Me.EntryMdbDir)
    End Sub
#End Region

#Region "PartsDb, PropsDb, Dao360 Events"
    Private Sub DxDB_StatusbarMsg(ByVal Msg As String) Handles DXDB.StatusbarMsg
        Me.SetStatusbarMessage(Msg)
    End Sub

    Private Sub DxDB_ProgressBarAction(ByVal Value As Integer, ByVal Mode As ProgbarMode) Handles DXDB.ProgBarAction
        Value = Value
        Select Case Mode
            Case ProgbarMode.Init
                Me.InitProgressBar(Value)
            Case ProgbarMode.Incr
                Me.IncrProgressBar()
            Case ProgbarMode.Close
                Me.FinishProgressBar()
        End Select
    End Sub

    Private Sub DxDB_SysException(ByVal UserMsg As String, ByVal Ex As System.Exception, ByVal LeadingSpaces As Integer) Handles DXDB.SysException
        Me.WriteTranscript(UserMsg, LeadingSpaces, Transcript.MsgType.Err)
        Me.PrintException(Ex, LeadingSpaces)
    End Sub

    Private Sub DxDB_TranscriptMsg(ByVal Msg As String, ByVal LeadingSpaces As Integer, ByVal MsgType As MsgType, ByVal AddLogMsg As Boolean) Handles DXDB.TranscriptMsg
        Me.WriteTranscript(Msg, LeadingSpaces, CType(MsgType, Transcript.MsgType), AddLogMsg)
    End Sub

    Private Sub DxDB_Add2Logfile(ByVal Msg As String, ByVal LeadingSpaces As Integer, ByVal MsgType As MsgType) Handles DXDB.Add2Logfile
        Me.AddLogMessage(Msg, LeadingSpaces, CType(MsgType, Transcript.MsgType))
    End Sub

    Private Sub Dao_Cfg_Notify(ByVal Msg As String, ByVal LeadingSpaces As Integer, ByVal Type As Transcript.MsgType) Handles Dao360.Notify
        Me.WriteTranscript(Msg, LeadingSpaces, Type)
    End Sub
#End Region

End Class

