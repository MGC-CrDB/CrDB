Imports System
Imports System.IO
Imports Microsoft.Win32
Imports System.Globalization

Public Class SddVxEnv

#Region "Declarations"
    Public Enum SddOsType
        win32
        win64
    End Enum

    Protected Enum Hkey
        HKEY_CLASSES_ROOT
        HKEY_CURRENT_USER
        HKEY_LOCAL_MACHINE
    End Enum

    Protected Class RegistryRelease
        Public Name As String
        Public Platform As SddOsType
    End Class

    Protected Platform As SddOsType
    Protected LatestRelNum As Double

    Protected _ActiveRelease As VxRelease
    Protected _Releases As List(Of VxRelease)
#End Region

#Region "Properties"
    Public ReadOnly Property IsValid() As Boolean
        Get
            If Me._Releases.Count = 0 Then Return False
            Return (Not IsNothing(Me._ActiveRelease))
        End Get
    End Property

    Public ReadOnly Property ActiveRelease() As VxRelease
        Get
            Return Me._ActiveRelease
        End Get
    End Property

    Public ReadOnly Property Releases() As List(Of VxRelease)
        Get
            Return Me._Releases
        End Get
    End Property
#End Region

#Region "Public Methods"
    Public Sub New(Optional GetInstalledReleases As Boolean = False)
        Me.LatestRelNum = -1

        Me._Releases = New List(Of VxRelease)

        Me.Platform = SddOsType.win32
        If System.Environment.Is64BitProcess Then
            Me.Platform = SddOsType.win64
        End If

        ' always get all registered releases
        If GetInstalledReleases Then
            Me.GetInstalledVxReleases()
        End If

        ' get environment if already set
        Me._ActiveRelease = Me.GetEnvironment()

        If Not IsNothing(Me.ActiveRelease) AndAlso Not Me.ReleaseExists(Me.ActiveRelease) Then
            Me.AddReleaseToList(Me._ActiveRelease)
        End If

        If Not Me.IsValid() Then
            ' no environment set - set active release
            Me._ActiveRelease = Me.SetActiveRelease()
        End If
    End Sub

    Public Function SetSddEnv(Optional ComVersion As Integer = -1) As Boolean
        Dim VXEnvServer As MGCPCBReleaseEnvironmentLib.MGCPCBReleaseEnvServer

        Try
            Me._ActiveRelease = Me.SetActiveRelease(ComVersion)

            If Not Me.IsValid() Then Return False

            Try
                VXEnvServer = New MGCPCBReleaseEnvironmentLib.MGCPCBReleaseEnvServer
            Catch ex As Exception
                VXEnvServer = Nothing
            End Try

            If IsNothing(VXEnvServer) Then Return False

            VXEnvServer.SetEnvironment(Me._ActiveRelease.SddHome)
            VXEnvServer = Nothing
            GC.Collect()

            Return Me.IsValid()
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Sub ShowNoReleaseMsgBox()
        Dim MsgText As String = String.Empty

        Dim Platform As String = "win32"
        If Environment.Is64BitProcess Then
            Platform = "win64"
        End If

        MsgText = "Unable to set EEVX Release Environment!" + vbNewLine + vbNewLine

        MsgText += "Required Release  (" + Platform + ")" + vbNewLine
        MsgText += "either not registered or not installed." + vbNewLine + vbNewLine
        MsgText += "Run 'MGC PCB Release Switcher' or install software."

        MsgBox(MsgText, MsgBoxStyle.Exclamation, My.Application.Info.ProductName + " - v" + My.Application.Info.Version.ToString())
    End Sub
#End Region

#Region "Release Identification Methods"
    Protected Sub GetInstalledVxReleases()
        Dim Release As VxRelease = Nothing
        Dim SddHome As String = String.Empty
        Dim SddVersion As String = String.Empty
        Dim ComVersion As String = String.Empty
        Dim SddPlatform As String = String.Empty
        Dim VxEnvReleases As Object(,) = Nothing
        Dim VXEnvServer As MGCPCBReleaseEnvironmentLib.MGCPCBReleaseEnvServer

        Try
            VXEnvServer = New MGCPCBReleaseEnvironmentLib.MGCPCBReleaseEnvServer
            'VXEnvServer = CType(CreateObject("MGCPCBReleaseEnvironmentLib.MGCPCBReleaseEnvServer"), MGCPCBReleaseEnvironmentLib.MGCPCBReleaseEnvServer)
        Catch ex As Exception
            VXEnvServer = Nothing
        End Try

        If IsNothing(VXEnvServer) Then Exit Sub

        Me.LatestRelNum = -1
        VxEnvReleases = CType(VXEnvServer.GetInstalledReleases(), Object(,))
        For i As Integer = 1 To UBound(VxEnvReleases, 1)
            For j As Integer = 0 To UBound(VxEnvReleases, 2)
                Select Case j
                    Case 0 ' COM_VERSION
                        ComVersion = CStr(VxEnvReleases(i, j))
                    Case 1 ' SDD_HOME
                        SddHome = CStr(VxEnvReleases(i, j))
                    Case 2 ' SDD_PLATFORM
                        SddPlatform = CStr(VxEnvReleases(i, j))
                    Case 3 ' SDD_VERSION
                        SddVersion = CStr(VxEnvReleases(i, j))
                End Select
            Next

            Release = New VxRelease(SddHome, SddVersion, SddPlatform, CInt(ComVersion))

            Me.AddReleaseToList(Release)
        Next

        VXEnvServer = Nothing
        GC.Collect()
    End Sub

    Protected Sub AddReleaseToList(Release As VxRelease)
        If Release.Number > Me.LatestRelNum Then
            Me.LatestRelNum = Release.Number
        End If
        Me._Releases.Add(Release)
    End Sub

    Protected Function GetEnvironment() As VxRelease
        Dim SDD_HOME As String = Environment.GetEnvironmentVariable("SDD_HOME", EnvironmentVariableTarget.Process)
        If IsNothing(SDD_HOME) Then Return Nothing

        Dim SDD_VERSION As String = Environment.GetEnvironmentVariable("SDD_VERSION", EnvironmentVariableTarget.Process)
        Dim PROG_ID_VER As String = Environment.GetEnvironmentVariable("PROG_ID_VER", EnvironmentVariableTarget.Process)
        Dim SDD_PLATFORM As String = Environment.GetEnvironmentVariable("SDD_PLATFORM", EnvironmentVariableTarget.Process)
        Dim EnvRel As VxRelease = New VxRelease(SDD_HOME, SDD_VERSION, SDD_PLATFORM, CInt(PROG_ID_VER))

        Me.LatestRelNum = EnvRel.Number
        Return New VxRelease(SDD_HOME, SDD_VERSION, SDD_PLATFORM, CInt(PROG_ID_VER))
    End Function

    Protected Function ReleaseExists(Rel As VxRelease) As Boolean
        For Each Item As VxRelease In Me.Releases
            If Item.Name = Rel.Name And Item.Platform = Rel.Platform And Item.RootPath = Rel.RootPath Then
                Return True
            End If
        Next
        Return False
    End Function

    Protected Function SetActiveRelease(Optional RequiredComVersion As Integer = -1) As VxRelease
        Try

            If RequiredComVersion > 1 Then
                For Each Rel As VxRelease In Me._Releases
                    If Rel.Platform = Me.Platform And RequiredComVersion = Rel.ComVersion Then
                        Return Rel
                    End If
                Next
            End If

            ' get default release registry
            Dim ActiveRegistryRelease As RegistryRelease = Me.GetActiveRegistryRelease()

            If IsNothing(ActiveRegistryRelease) Then
                For Each Rel As VxRelease In Me._Releases
                    If Rel.Platform = Me.Platform And Rel.Number = Me.LatestRelNum Then
                        Return Rel
                    End If
                Next
            Else
                For Each Rel As VxRelease In Me._Releases
                    If ActiveRegistryRelease.Platform = Rel.Platform And ActiveRegistryRelease.Name = Rel.Name Then
                        Return Rel
                    End If
                Next
            End If

        Catch ex As Exception
            Return Nothing
        End Try
        Return Nothing
    End Function

    Protected Function GetActiveRegistryRelease() As RegistryRelease
        Dim RegRelease As RegistryKey, RegActRel As RegistryRelease

        Try
            If System.Environment.Is64BitProcess Then
                RegRelease = Me.FindRegSubKey(Hkey.HKEY_CURRENT_USER, "SOFTWARE\Mentor Graphics", "Releases", RegistryView.Registry64)
            Else
                RegRelease = Me.FindRegSubKey(Hkey.HKEY_CURRENT_USER, "SOFTWARE\Mentor Graphics", "Releases", RegistryView.Registry32)
            End If

            ' get default release
            If Not IsNothing(RegRelease) Then
                RegActRel = New RegistryRelease()
                RegActRel.Name = Path.GetFileName(Me.GetRegKeyValue(RegRelease, ""))
                RegActRel.Platform = Me.GetActiveReleasePlatform(RegRelease)
                Return RegActRel
            End If
        Catch ex As Exception
            Return Nothing
        End Try

        Return Nothing
    End Function

    Protected Function GetActiveReleasePlatform(RegRelease As RegistryKey) As SddOsType
        Dim PlatformPath As String = Me.GetRegKeyValue(RegRelease, "") + "\win64"
        If Directory.Exists(PlatformPath) Then
            Return SddOsType.win64
        Else
            Return SddOsType.win32
        End If
    End Function
#End Region

#Region "Registry Methods"
    Protected Function GetRegKeyValue(Key As RegistryKey, ValueName As String) As String
        Dim ProcessedVal As String = String.Empty
        Dim VxRelease As String = CStr(Key.GetValue(ValueName))
        If String.IsNullOrEmpty(VxRelease) Then Return Nothing
        For Each Chr As String In VxRelease
            If Asc(Chr) = 0 Then Exit For
            ProcessedVal += Chr
        Next
        Return ProcessedVal.Trim()
    End Function

    Protected Function FindRegSubKey(HKEY As Hkey, Key As String, SubKey As String, RegView As RegistryView) As RegistryKey
        Dim BaseKey As RegistryKey
        Dim KeyStr As String, Key2Find As RegistryKey = Nothing

        KeyStr = Key + "\" + SubKey

        If Key = String.Empty Then KeyStr = SubKey
        If SubKey = String.Empty Then KeyStr = Key
        If KeyStr = String.Empty Then Return Nothing

        Try
            Select Case HKEY
                Case SddVxEnv.Hkey.HKEY_CLASSES_ROOT
                    BaseKey = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.ClassesRoot, RegView)
                Case SddVxEnv.Hkey.HKEY_CURRENT_USER
                    BaseKey = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentUser, RegView)
                Case SddVxEnv.Hkey.HKEY_LOCAL_MACHINE
                    BaseKey = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, RegView)
                Case Else
                    Return Nothing
            End Select

            If IsNothing(BaseKey) Then Return Nothing
            Key2Find = BaseKey.OpenSubKey(KeyStr)
        Catch ex As Exception
            ' do nothing
        End Try

        If IsNothing(Key2Find) Then Return Nothing
        Return Key2Find
    End Function
#End Region

End Class

Public Class VxRelease

#Region "Declarations"
    Protected _Platform As SddVxEnv.SddOsType
    Protected _Name As String
    Protected _Number As Double
    Protected _ComVersion As Integer
    Protected _RootPath As String
    Protected _SddHome As String
#End Region

#Region "Properties"
    Public ReadOnly Property Platform As SddVxEnv.SddOsType
        Get
            Return Me._Platform
        End Get
    End Property

    Public ReadOnly Property Name As String
        Get
            Return Me._Name
        End Get
    End Property

    Public ReadOnly Property Number As Double
        Get
            Return Me._Number
        End Get
    End Property

    Public ReadOnly Property ComVersion As Integer
        Get
            Return Me._ComVersion
        End Get
    End Property

    Public ReadOnly Property RootPath As String
        Get
            Return Me._RootPath
        End Get
    End Property

    Public ReadOnly Property SddHome As String
        Get
            Return Me._SddHome
        End Get
    End Property

    Public ReadOnly Property LmUtil As String
        Get
            Dim Exe As String = Me._SddHome + "\common\" + Me._Platform.ToString + "\_bin\lmutil.exe"
            If My.Computer.FileSystem.FileExists(Exe) Then
                Return Exe
            End If
            Return "<File Not Found!>"
        End Get
    End Property

    Public ReadOnly Property LmTools As String
        Get
            Dim Exe As String = Me._SddHome + "\common\" + Me._Platform.ToString + "\_bin\lmtools.exe"
            If My.Computer.FileSystem.FileExists(Exe) Then
                Return Exe
            End If
            Return "<File Not Found!>"
        End Get
    End Property

    Public ReadOnly Property MgCmd As String
        Get
            Dim Exe As String = Me._SddHome + "\common\" + Me._Platform.ToString + "\bin\mgcmd.exe"
            If My.Computer.FileSystem.FileExists(Exe) Then
                Return Exe
            End If
            Return "<File Not Found!>"
        End Get
    End Property

    Public ReadOnly Property MgLaunch As String
        Get
            Dim Exe As String = Me._SddHome + "\common\" + Me._Platform.ToString + "\bin\mglaunch.exe"
            If My.Computer.FileSystem.FileExists(Exe) Then
                Return Exe
            End If
            Return "<File Not Found!>"
        End Get
    End Property

    Public ReadOnly Property MgInvoke As String
        Get
            Dim Exe As String = Me._SddHome + "\common\" + Me._Platform.ToString + "\bin\mginvoke.exe"
            If My.Computer.FileSystem.FileExists(Exe) Then
                Return Exe
            End If
            Return "<File Not Found!>"
        End Get
    End Property

    Public ReadOnly Property MgcScript As String
        Get
            Dim Exe As String = Me._SddHome + "\common\" + Me._Platform.ToString + "\bin\mgcscript.exe"
            If My.Computer.FileSystem.FileExists(Exe) Then
                Return Exe
            End If
            Return "<File Not Found!>"
        End Get
    End Property

    Public ReadOnly Property ReleaseSwitcher As String
        Get
            Dim Exe As String = Me.SddHome + "\common\" + Me._Platform.ToString + "\_bin\ReleaseSwitcher.exe"
            If My.Computer.FileSystem.FileExists(Exe) Then
                Return Exe
            End If
            Return "<File Not Found!>"
        End Get
    End Property

    Public ReadOnly Property SutHome As String
        Get
            Dim SutDir As String = Me.SddHome + "\EDM-Server\Utilities"
            If My.Computer.FileSystem.DirectoryExists(SutDir) Then
                Return SutDir
            End If
            Return "<Directory Not Found!>"
        End Get
    End Property
#End Region

#Region "Public Methods"
    Public Sub New(ByVal SddHome As String, ByVal SddVersion As String, ByVal SddPlatform As String, ByVal ComVersion As Integer)
        Me._SddHome = SddHome
        Me._ComVersion = ComVersion
        Me._Name = SddVersion
        Me._Number = -1
        Me._RootPath = Path.GetDirectoryName(Path.GetDirectoryName(SddHome))

        Me._Platform = SddVxEnv.SddOsType.win32
        If SddPlatform.ToLower = "win64" Then
            Me._Platform = SddVxEnv.SddOsType.win64
        End If

        Me.SetReleaseNumber()
    End Sub

    Public Function GetMgcToolExePath(ExeName As String) As String
        Dim ToolPath As String
        If String.IsNullOrEmpty(Path.GetExtension(ExeName)) Then
            ExeName += ".exe"
        End If
        ToolPath = Me.SddHome + "/common/" + Me.Platform.ToString + "/bin/" + ExeName
        If File.Exists(ToolPath) Then Return ToolPath
        Return String.Empty
    End Function

    Protected Sub SetReleaseNumber()
        Dim Arr As String() = Split(Me._Name, "EEVX.")
        If Not Arr.Length() = 2 Then Exit Sub
        Me._Number = Me.String2Double(Arr(1).Trim)
    End Sub

    Protected Function String2Double(ByVal Value As String, Optional ByVal Digits As Integer = 3) As Double
        Dim DblVal As Double, DecimalSeparator As String = Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator

        If Not DecimalSeparator = "." Then Value = Value.Replace(".", DecimalSeparator)

        Try
            DblVal = Double.Parse(Value, CultureInfo.CurrentCulture.NumberFormat)
        Catch ex As Exception
            Return -1
        End Try

        Return System.Math.Round(DblVal, Digits)
    End Function
#End Region

End Class
