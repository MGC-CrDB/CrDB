Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.IO
Imports System.Reflection
Imports System.Threading
Imports LibraryManager

Public Class LibVxApp

#Region "API Methods"
    <DllImport("ole32.dll", ExactSpelling:=True, PreserveSig:=False)>
    Private Shared Function GetRunningObjectTable(reserved As Int32) As IRunningObjectTable
    End Function

    <DllImport("ole32.dll")>
    Private Shared Sub GetRunningObjectTable(ByVal reserved As Integer, <Out> ByRef prot As IRunningObjectTable)
    End Sub

    <DllImport("ole32.dll", CharSet:=CharSet.Unicode, ExactSpelling:=True, PreserveSig:=False)>
    Private Shared Function CreateItemMoniker(lpszDelim As String, lpszItem As String) As IMoniker
    End Function

    <DllImport("ole32.dll")>
    Private Shared Function CreateBindCtx(ByVal reserved As UInteger, <Out> ByRef ppbc As IBindCtx) As Integer
    End Function

    <DllImport("ole32.dll", ExactSpelling:=True, PreserveSig:=False)>
    Private Shared Function CreateBindCtx(reserved As Integer) As IBindCtx
    End Function
#End Region

#Region "Structures, Enumerations"
    Public Enum SddOsType
        win32
        win64
    End Enum

    Public Enum AppErrCode As Integer
        ' application erro codes
        LibAppUndefined = -1
        LibAppNotRunning = 429
        LibAppStartFailed = 767
        LibAppPlatformNotValid = 878
        LibAppReleaseNotValid = 989

        ' document error codes
        LibDocNoLibraryOpen = 2
        LibDocLibraryIsOpen = 4
    End Enum

    Public Structure AppErrMsg
        ' application error
        Const LibAppUndefined As String = "Library Manager Application does not exist!"
        Const LibAppNotRunning As String = "Library Manager Application is Not running!"
        Const LibAppStartFailed As String = "Library Manager Application failed to start!"
        Const LibAppPlatformNotValid As String = "Library Manager Application has wrong platform type!"

        ' document errors
        Const LibDocNoLibraryLoaded As String = "Running Instance of Library Manager found, but No Library open!"
        Const LibDocLibraryIsLoaded As String = "Running Instance of Library Manager found and Library loaded!"
    End Structure
#End Region

#Region "Declarations"
    Private ComVersion As Integer
    Private SddEnv As SddVxEnv

    Private ErrCode As AppErrCode

    ' members exposed by properties
    Private _Releases As List(Of VxRelease)
    Private _Platform As SddOsType
    Private _IsAppValid As Boolean

    Private WithEvents O_App As LibraryManager.LibraryManagerApp
    Private WithEvents O_Lib As LibraryManager.MGCLMLibrary

    Private _PartEditor As MGCPCBPartsEditor.PartsEditorDlg

    ' event objects
    Public Event OnLMCModified()
    Public Event OnQuit()
#End Region

#Region "Class Properties"
    Public ReadOnly Property Releases() As List(Of VxRelease)
        Get
            If Me.SddEnv.IsValid() Then
                Return Me.SddEnv.Releases
            Else
                Return New List(Of VxRelease)
            End If
        End Get
    End Property

    Public ReadOnly Property ReleaseName() As String
        Get
            If Me.SddEnv.IsValid() Then
                Return Me.SddEnv.ActiveRelease().Name
            Else
                Return String.Empty
            End If
        End Get
    End Property

    Public ReadOnly Property ReleaseComVersion() As String
        Get
            If Me.SddEnv.IsValid() Then
                Return Me.SddEnv.ActiveRelease().ComVersion.ToString
            Else
                Return String.Empty
            End If
        End Get
    End Property

    Public ReadOnly Property ReleasePlatform() As String
        Get
            If Me.SddEnv.IsValid() Then
                Return Me.SddEnv.ActiveRelease().Platform.ToString
            Else
                Return String.Empty
            End If
        End Get
    End Property

    Public ReadOnly Property ReleaseSddHome() As String
        Get
            If Me.SddEnv.IsValid() Then
                Return Me.SddEnv.ActiveRelease.SddHome
            Else
                Return String.Empty
            End If
        End Get
    End Property

    Public ReadOnly Property App() As LibraryManager.LibraryManagerApp
        Get
            Return Me.O_App
        End Get
    End Property

    Public ReadOnly Property IsAppValid() As Boolean
        Get
            Return Me._IsAppValid
        End Get
    End Property

    Public Property Visible As Boolean
        Get
            Return Me.O_App.Visible
        End Get
        Set(ByVal value As Boolean)
            Me.O_App.Visible = value
        End Set
    End Property

    Public ReadOnly Property Platform As SddOsType
        Get
            Return Me._Platform
        End Get
    End Property

    Public ReadOnly Property Name() As String
        Get
            Return "Xpedition Library Manager"
        End Get
    End Property

    Public Function IsLibraryOpen() As Boolean
        Try
            Dim ActLib As LibraryManager.MGCLMLibrary
            ActLib = Me.O_App.ActiveLibrary
            Return Not IsNothing(ActLib)
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function LmcDir() As String
        If IsNothing(Me.O_Lib) Then Return String.Empty
        Return Path.GetDirectoryName(Me.O_Lib.FullName)
    End Function

    Public Function LmcFile() As String
        If IsNothing(Me.O_Lib) Then Return String.Empty
        Return Me.O_Lib.FullName
    End Function

    Public ReadOnly Property PartEditor As MGCPCBPartsEditor.PartsEditorDlg
        Get
            If IsNothing(Me._PartEditor) Then
                Dim ProgID As String = "MGCPCBLibraries.PartsEditorDlg." + Me.ComVersion.ToString
                Me._PartEditor = CType(CreateObject(ProgID), MGCPCBPartsEditor.PartsEditorDlg)
            End If
            Return Me._PartEditor
        End Get
    End Property

    Public ReadOnly Property IsLibraryValid As Boolean
        Get
            Select Case Me.ErrCode
                Case AppErrCode.LibDocLibraryIsOpen
                    Return True
                Case Else
                    Return False
            End Select
        End Get
    End Property

    Public ReadOnly Property ErrMsg() As String
        Get
            Select Case Me.ErrCode
                Case AppErrCode.LibDocLibraryIsOpen
                    Return AppErrMsg.LibDocLibraryIsLoaded
                Case AppErrCode.LibAppNotRunning
                    Return AppErrMsg.LibAppNotRunning
                Case AppErrCode.LibAppStartFailed
                    Return AppErrMsg.LibAppStartFailed
                Case AppErrCode.LibAppPlatformNotValid
                    Return AppErrMsg.LibAppPlatformNotValid
                Case AppErrCode.LibAppUndefined
                    Return AppErrMsg.LibAppUndefined
                Case AppErrCode.LibDocNoLibraryOpen
                    Return AppErrMsg.LibDocNoLibraryLoaded
                Case Else
                    Return AppErrMsg.LibAppUndefined
            End Select
        End Get
    End Property
#End Region

#Region "Get COM Version Method"
    Public Function GetClassCOMVersion() As Integer
        Dim buff As String(), ComVer As Integer
        Dim AN As AssemblyName = GetType(DxDbCore).Assembly.GetName()
        buff = Me.SplitNoEmpties(AN.Name, GetType(DxDbCore).Namespace)
        If Integer.TryParse(buff(0), ComVer) Then
            Return ComVer
        Else
            Return -1
        End If
    End Function

    Public Function SplitNoEmpties(Input As String, Delimiter As String) As String()
        Dim arrSplit As String() = Split(Input, Delimiter)
        Return arrSplit.Where(Function(s) Not String.IsNullOrWhiteSpace(s)).ToArray()
    End Function
#End Region

#Region "Class Constructor"
    Public Sub New(Optional GetInstalledReleases As Boolean = False)
        Me.ComVersion = -1

        If Me.ComVersion = -1 Then
            Me.ComVersion = Me.GetClassCOMVersion()
        End If

        Me.ErrCode = AppErrCode.LibAppUndefined
        Me._Platform = SddOsType.win32
        If System.Environment.Is64BitProcess Then
            Me._Platform = SddOsType.win64
        End If

        Me.SddEnv = New SddVxEnv(GetInstalledReleases)
        If Not Me.SddEnv.IsValid() Then
            ' no release found
            Me.ErrCode = AppErrCode.LibAppReleaseNotValid
        End If

        If Me.ComVersion = -1 Then
            Me.ComVersion = Me.SddEnv.ActiveRelease.ComVersion
        End If
    End Sub
#End Region

#Region "Library Manager Events"
    Private Sub App_LMCModified() Handles O_App.LMCModified
        RaiseEvent OnLMCModified()
    End Sub

    Protected Sub App_Quit() Handles O_App.Quit
        RaiseEvent OnQuit()
    End Sub
#End Region

#Region "Public Library Manager Application Methods"
    Public Function GetExistingInstance() As Boolean
        Dim Procs As Process() = Process.GetProcessesByName("LibraryManager")
        Return (Procs.Count > 0)
    End Function

    Public Function CloseExistingInstance() As Boolean
        Dim Proc As Process, Procs As Process()

        Procs = Process.GetProcessesByName("LibraryManager")
        If Procs.Count = 0 Then Return True

        Try
            For Each Proc In Procs
                Proc.Kill()
            Next
        Catch ex As Exception
            Return False
        End Try

        Return True
    End Function

    Public Function ConnectApplication(Optional PrjPath As String = "") As Boolean
        Dim rApps As List(Of RunningApp)
        Dim rApp As RunningApp

        ' members exposed by properties
        Me.O_App = Nothing
        Me._IsAppValid = False

        If Not Me.SddEnv.IsValid() Then
            ' no release found
            Me.ErrCode = AppErrCode.LibAppReleaseNotValid
            Return False
        End If

        rApps = Me.GetRunningApps(PrjPath, Me.SddEnv.Releases)
        If rApps.Count > 0 Then
            If rApps.Count = 1 Then
                rApp = rApps(0)
            Else
                rApp = frmSelectApp.CreateAndShow(rApps)
            End If
        Else
            Me.ErrCode = AppErrCode.LibAppNotRunning
            Return False
        End If

        If IsNothing(rApp) Then Return False

        If Not Me.SddEnv.SetSddEnv(rApp.ComVersion) Then
            ' should never happen! no valid release found
            Me.ErrCode = AppErrCode.LibAppReleaseNotValid
            Return False
        End If

        ' Environment and app ok ...
        Me.ErrCode = AppErrCode.LibDocNoLibraryOpen
        Me.O_App = rApp.Application

        If Not rApp.LmcFilePath = String.Empty Then
            Me._IsAppValid = True
            Me.ErrCode = AppErrCode.LibDocLibraryIsOpen
            Me.O_Lib = rApp.Application.ActiveLibrary
        End If

        '----------------------
        ' dispose rApps
        Me.DisposeRunningApps(rApps)

        Return Me._IsAppValid
    End Function

    Public Function CreateApplication(AppVisible As Boolean, Optional LmcFile As String = "") As Boolean
        ' members exposed by properties
        Me.O_App = Nothing
        Me._IsAppValid = False

        If Not Me.SddEnv.SetSddEnv(Me.ComVersion) OrElse Not Me.SddEnv.IsValid() Then
            Me.ErrCode = AppErrCode.LibAppReleaseNotValid
            Return Me._IsAppValid
        End If

        ' -------------------------------
        ' create new apllication instance
        Dim ProgID As String
        If Not Me.ComVersion = -1 Then
            ProgID = Me.GetAppProgID(Me.ComVersion)
        Else
            ProgID = Me.GetAppProgID(Me.SddEnv.ActiveRelease.ComVersion)
        End If

        Try
            Me.O_App = CType(CreateObject(ProgID), LibraryManager.LibraryManagerApp)
            Me.O_App.Visible = AppVisible
            Me._IsAppValid = True
        Catch ex As Exception
            Me.ErrCode = AppErrCode.LibAppStartFailed
            Return Me._IsAppValid
        End Try

        ' check platform type
        If Not Me._Platform = Me.GetAppPlatform(Me.O_App.FullName) Then
            Me.ErrCode = AppErrCode.LibAppPlatformNotValid
            Return Me._IsAppValid
        End If

        ' load project
        Me.ErrCode = AppErrCode.LibDocNoLibraryOpen
        If File.Exists(LmcFile) Then
            Try
                Me.O_Lib = Me.O_App.OpenLibrary(LmcFile)
                Me._IsAppValid = Not IsNothing(Me.O_Lib)
                If Me._IsAppValid Then
                    Me.ErrCode = AppErrCode.LibDocLibraryIsOpen
                End If
            Catch ex As Exception
                Me._IsAppValid = False
                Me.ErrCode = AppErrCode.LibDocNoLibraryOpen
            End Try
        End If

        Return Me._IsAppValid
    End Function

    Public Sub ShowNoApplicationMsgBox()
        Dim PlatformCOMtxt As String
        If Me.ErrCode = AppErrCode.LibAppReleaseNotValid Then
            Me.SddEnv.ShowNoReleaseMsgBox()
        Else
            If Me.ComVersion = -1 Then
                PlatformCOMtxt = "Platform=" + Me._Platform.ToString
            Else
                PlatformCOMtxt = "Platform=" + Me._Platform.ToString + ", COMVersion=" + CStr(Me.ComVersion)
            End If
            MsgBox("Library Manager (" + PlatformCOMtxt + ")" + vbCrLf + Me.ErrMsg, MsgBoxStyle.Exclamation _
           , My.Application.Info.ProductName + " - v" + My.Application.Info.Version.ToString())
        End If
    End Sub

    Public Sub Close()
        If Not IsNothing(Me.App) AndAlso Me.IsAppValid Then
            Me.DisConnectPartsEditor()
            Me.O_App.Quit()
        End If
        Me.ErrCode = AppErrCode.LibAppUndefined
        Me.ReleaseCOMObject(Me.O_App)
        Me.O_App = Nothing
        Me._IsAppValid = False
    End Sub

    Public Sub Disconnect()
        If Not IsNothing(Me.App) AndAlso Me.IsAppValid Then
            Me.DisConnectPartsEditor()
        End If
        Me.ErrCode = AppErrCode.LibAppUndefined
        Me.ReleaseCOMObject(Me.O_App)
        Me.O_App = Nothing
        Me._IsAppValid = False
    End Sub

    Public Function ActiveLibrary() As LibraryManager.MGCLMLibrary
        If IsNothing(Me.O_App) Then Return Nothing
        Return Me.O_App.ActiveLibrary
    End Function

    Public Function GetMgcToolExePath(ExeName As String) As String
        If Not Me.SddEnv.IsValid() Then Return String.Empty
        Return Me.SddEnv.ActiveRelease.GetMgcToolExePath(ExeName)
    End Function

    Public Function GetLibMgrInstances() As List(Of LibraryManager.LibraryManagerApp)
        Dim runningObjectInst As Object = Nothing
        Dim runningObjectTable As IRunningObjectTable
        Dim monikerEnumerator As IEnumMoniker = Nothing
        Dim LibMgrIns As LibraryManager.LibraryManagerApp
        Dim monikers(1) As IMoniker

        Dim LOs As List(Of Object) = Me.GetRunningLibMgrInstances()

        Dim LibMgrs As New List(Of LibraryManager.LibraryManagerApp)

        runningObjectTable = GetRunningObjectTable(0)
        runningObjectTable.EnumRunning(monikerEnumerator)
        monikerEnumerator.Reset()

        Dim numFetched As IntPtr = New IntPtr()
        While (monikerEnumerator.Next(1, monikers, numFetched) = 0)
            Dim ctx As IBindCtx
            ctx = CreateBindCtx(0)

            Dim runningObjectName As String = ""
            LibMgrIns = Nothing
            monikers(0).GetDisplayName(ctx, Nothing, runningObjectName)

            If (runningObjectName.Equals("")) Then Continue While
            Debug.Print("RunningObjectName: " + runningObjectName)

            runningObjectTable.GetObject(monikers(0), runningObjectInst)

            'Check if object is a Library Manager Object object
            Try
                LibMgrIns = DirectCast(runningObjectInst, LibraryManager.LibraryManagerApp)
            Catch ex As Exception
                ' do nothing
                Debug.Print("GetLibMgrInstances(): " + ex.Message)
            End Try

            If Not IsNothing(LibMgrIns) Then
                If Not LibMgrs.Contains(LibMgrIns) Then LibMgrs.Add(LibMgrIns)
            End If
        End While

        Return LibMgrs
    End Function

    Public Sub PrintRunningObjectTableROT()
        Dim Rot As IRunningObjectTable = Nothing
        Dim enumMoniker As IEnumMoniker = Nothing

        If Not Me.SddEnv.SetSddEnv(Me.ComVersion) Then
            Exit Sub
        End If
        GetRunningObjectTable(0, Rot)

        Dim ProgID As String, ProgIdType As Type
        ProgID = Me.GetAppProgID(Me.SddEnv.ActiveRelease.ComVersion)
        ProgIdType = Type.GetTypeFromProgID(ProgID)
        Console.WriteLine("GUID: {0}", ProgIdType.GUID.ToString.ToUpper)

        Rot.EnumRunning(enumMoniker)
        enumMoniker.Reset()

        ' Array für Moniker-Abfrage erzeugen
        Dim fetched As IntPtr = IntPtr.Zero
        Dim moniker As IMoniker() = New IMoniker(0) {}

        While (enumMoniker.Next(1, moniker, fetched) = 0)
            Dim bindCtx As IBindCtx = Nothing
            CreateBindCtx(0, bindCtx)
            Dim displayName As String = String.Empty
            moniker(0).GetDisplayName(bindCtx, Nothing, displayName)

            ' Bindungsobjekt entsorgen
            Marshal.ReleaseComObject(bindCtx)

            Console.WriteLine("Display Name: {0}", displayName)
        End While
    End Sub
#End Region

#Region "Public Methods"
    Public Shared Function GetLibMgrLmcFile(Optional ComVersion As Integer = -1) As String
        Dim LmcFile As String, ProgID As String
        Dim LibMgr As LibraryManager.LibraryManagerApp

        ProgID = "LibraryManager.Application"
        If Not ComVersion = -1 Then
            ProgID = ProgID + "." + ComVersion.ToString
        End If

        LmcFile = String.Empty

        Try
            LibMgr = CType(GetObject(, ProgID), LibraryManager.LibraryManagerApp)
            If Not IsNothing(LibMgr.ActiveLibrary) Then
                LmcFile = LibMgr.ActiveLibrary.FullName
            End If
            LibMgr = Nothing
            GC.Collect()
        Catch ex As Exception
            Return LmcFile
        End Try

        Return LmcFile
    End Function

    Public Shared Function GetCentLibPsfFile(LmcFile As String) As String
        Dim LibDir As String, PrpFile As String

        PrpFile = String.Empty
        If My.Computer.FileSystem.FileExists(LmcFile) Then
            LibDir = Path.GetDirectoryName(LmcFile)
            PrpFile = LibDir + "\" + Path.GetFileName(LibDir) + ".psf"
            If My.Computer.FileSystem.FileExists(PrpFile) Then
                Return PrpFile
            End If
            Return String.Empty
        End If

        Return String.Empty
    End Function
#End Region

#Region "Protected Library Manager Application Methods"
    Protected Function GetAppProgID(ByVal ComVersion As Integer) As String
        Dim ProgIDExtension As String = String.Empty

        If Not ComVersion = -1 Then
            ProgIDExtension = "." + ComVersion.ToString
        End If
        Return "LibraryManager.Application" + ProgIDExtension
    End Function
#End Region

#Region "Get running Application Methods"
    Protected Function GetRunningApps(PrjPath As String, Releases As List(Of VxRelease)) As List(Of RunningApp)
        Dim ListDxd As List(Of LibraryManager.LibraryManagerApp)
        Dim ReleaseName As String, AppPlatform As SddOsType
        Dim App As RunningApp, RunApps As New List(Of RunningApp)

        ' looup application(s) for each registered release
        ListDxd = Me.GetLibMgrInstances()

        For Each VxRelease As VxRelease In Releases
            If Me._Platform = VxRelease.Platform Then

                For Each DxdApp As LibraryManager.LibraryManagerApp In ListDxd
                    ReleaseName = Me.GetAppReleaseName(DxdApp.FullName)
                    AppPlatform = Me.GetAppPlatform(DxdApp.FullName)
                    If ReleaseName = VxRelease.Name And AppPlatform = VxRelease.Platform Then
                        App = New RunningApp(VxRelease.ComVersion, AppPlatform, ReleaseName, DxdApp)
                        If Me.AddApp(App, PrjPath) Then RunApps.Add(App)
                    End If
                Next
            End If
        Next

        Return RunApps
    End Function

    Private Function GetRunningLibMgrInstances() As List(Of Object)
        Dim ComObject As Object = Nothing
        Dim BindCtx As IBindCtx, DisplayName As String = String.Empty
        Dim ClassIDs As List(Of String) = New List(Of String)

        Dim ProgID As String = "LibraryManager.Application" 'Me.GetAppProgID(Me.SddEnv.ActiveRelease.ComVersion)
        Dim ProgIdType As Type = Type.GetTypeFromProgID(ProgID)
        If ProgIdType IsNot Nothing Then ClassIDs.Add(ProgIdType.GUID.ToString.ToUpper)

        Dim Rot As IRunningObjectTable = Nothing
        GetRunningObjectTable(0, Rot)
        If Rot Is Nothing Then Return New List(Of Object)

        Dim monikerEnumerator As IEnumMoniker = Nothing
        Rot.EnumRunning(monikerEnumerator)
        If monikerEnumerator Is Nothing Then Return New List(Of Object)
        monikerEnumerator.Reset()

        Dim instances As List(Of Object) = New List(Of Object)()
        Dim pNumFetched As IntPtr = New IntPtr()
        Dim monikers As IMoniker() = New IMoniker(0) {}

        While monikerEnumerator.Next(1, monikers, pNumFetched) = 0
            BindCtx = CreateBindCtx(0)
            If BindCtx Is Nothing Then Continue While
            monikers(0).GetDisplayName(BindCtx, Nothing, DisplayName)

            ' Debug.Print("Display Name: {0}", DisplayName)

            For Each clsId As String In ClassIDs
                If DisplayName.ToUpper().IndexOf(clsId) > 0 Then
                    Rot.GetObject(monikers(0), ComObject)
                    If ComObject Is Nothing Then Continue For
                    instances.Add(ComObject)
                    Exit For
                End If
            Next
        End While

        Return instances
    End Function

    Protected Function AddApp(ByVal App As RunningApp, ByVal PrjPath As String) As Boolean
        If Me.ComVersion > 0 And Not Me.ComVersion = App.ComVersion Then
            Return False
        End If

        If PrjPath = String.Empty Then Return True
        If Not File.Exists(PrjPath) Then Return False
        Return (App.LmcFilePath = PrjPath)
    End Function

    Protected Sub DisposeRunningApps(rApps As List(Of RunningApp))
        For Each rApp As RunningApp In rApps
            rApp.Application = Nothing
        Next
        rApps.Clear()
    End Sub

    Protected Function GetAppReleaseName(AppPath As String) As String
        Dim Vect As String() = Split(AppPath, "\SDD_HOME\")
        Return Path.GetFileName(Vect(0))
    End Function

    Protected Function GetAppPlatform(AppPath As String) As SddOsType
        If AppPath.Contains("\win64\") Then
            Return SddOsType.win64
        End If
        Return SddOsType.win32
    End Function

    Private Sub DisConnectPartsEditor()
        If Not IsNothing(Me._PartEditor) Then
            Try
                Me._PartEditor.CloseActiveDatabase()
            Catch ex As Exception
                ' do nothing
            End Try
            Me.ReleaseCOMObject(Me._PartEditor)
        End If
    End Sub

    Protected Sub ReleaseCOMObject(ByVal COMObj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(COMObj)
            COMObj = Nothing
        Catch ex As Exception
            COMObj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    'Public Class RunningApp
    '    Public Name As String
    '    Public Type As EPcbApplicationType
    '    Public ComVersion As Integer
    '    Public ReleaseName As String
    '    Public Platform As SddOsType
    '    Public Document As String
    '    Public Application As MGCPCB.Application

    '    Public Sub New(nName As String, nType As EPcbApplicationType, nComVersion As Integer, nPlatform As SddOsType, nReleaseName As String, nApplication As MGCPCB.Application)
    '        Me.Name = nName
    '        Me.Type = nType
    '        Me.ComVersion = nComVersion
    '        Me.Platform = nPlatform
    '        Me.Document = String.Empty
    '        If Not IsNothing(nApplication.ActiveDocument) Then
    '            Me.Document = nApplication.ActiveDocument.MasterFullName
    '        End If
    '        Me.ReleaseName = nReleaseName
    '        Me.Application = nApplication
    '    End Sub
    'End Class

    Public Class RunningApp
        Public Name As String
        Public ComVersion As Integer
        Public ReleaseName As String
        Public Platform As SddOsType
        Public LmcFilePath As String
        Public Application As LibraryManager.LibraryManagerApp

        Public Sub New(nComVersion As Integer, nPlatform As SddOsType, nReleaseName As String, nApplication As LibraryManager.LibraryManagerApp)
            Me.Name = "LibraryManager (" + nReleaseName + ")"
            Me.ComVersion = nComVersion
            Me.Platform = nPlatform
            Try
                Me.LmcFilePath = nApplication.ActiveLibrary.FullName
            Catch ex As Exception
                Me.LmcFilePath = String.Empty
            End Try
            Me.ReleaseName = nReleaseName
            Me.Application = nApplication
        End Sub
    End Class
#End Region

End Class
