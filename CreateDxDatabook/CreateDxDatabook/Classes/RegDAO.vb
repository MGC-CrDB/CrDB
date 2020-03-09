Imports Microsoft.Win32

Public Class RegDAO

#Region "TypeLib: {00025E01-0000-0000-C000-000000000046} Regitry Entries"
    ' HKEY_CLASSES_ROOT\TypeLib
    ' {00025E01-0000-0000-C000-000000000046}
    ' {00025E01-0000-0000-C000-000000000046}\5.0         @="Microsoft DAO 3.6 Object Library"
    ' {00025E01-0000-0000-C000-000000000046}\5.0\0
    ' {00025E01-0000-0000-C000-000000000046}\5.0\0\win32 @="C:\\Program Files (x86)\\Common Files\\microsoft shared\\DAO\\dao360.dll"
    ' {00025E01-0000-0000-C000-000000000046}\5.0\FLAGS   @="0"
    ' {00025E01-0000-0000-C000-000000000046}\5.0\HELPDIR @="C:\\Program Files (x86)\\Common Files\\microsoft shared\\DAO\\"

    ' HKEY_CLASSES_ROOT\Wow6432Node\TypeLib
    ' {00025E01-0000-0000-C000-000000000046}
    ' {00025E01-0000-0000-C000-000000000046}\5.0         @="Microsoft DAO 3.6 Object Library"
    ' {00025E01-0000-0000-C000-000000000046}\5.0\0
    ' {00025E01-0000-0000-C000-000000000046}\5.0\0\win32 @="C:\\Program Files (x86)\\Common Files\\microsoft shared\\DAO\\dao360.dll"
    ' {00025E01-0000-0000-C000-000000000046}\5.0\FLAGS   @="0"
    ' {00025E01-0000-0000-C000-000000000046}\5.0\HELPDIR @="C:\\Program Files (x86)\\Common Files\\microsoft shared\\DAO\\"

    ' HKEY_LOCAL_MACHINE\SOFTWARE\Classes\TypeLib
    ' {00025E01-0000-0000-C000-000000000046}
    ' {00025E01-0000-0000-C000-000000000046}\5.0         @="Microsoft DAO 3.6 Object Library"
    ' {00025E01-0000-0000-C000-000000000046}\5.0\0
    ' {00025E01-0000-0000-C000-000000000046}\5.0\0\win32 @="C:\\Program Files (x86)\\Common Files\\microsoft shared\\DAO\\dao360.dll"
    ' {00025E01-0000-0000-C000-000000000046}\5.0\FLAGS   @="0"
    ' {00025E01-0000-0000-C000-000000000046}\5.0\HELPDIR @="C:\\Program Files (x86)\\Common Files\\microsoft shared\\DAO\\"

    ' HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Wow6432Node\TypeLib
    ' {00025E01-0000-0000-C000-000000000046}
    ' {00025E01-0000-0000-C000-000000000046}\5.0         @="Microsoft DAO 3.6 Object Library"
    ' {00025E01-0000-0000-C000-000000000046}\5.0\0
    ' {00025E01-0000-0000-C000-000000000046}\5.0\0\win32 @="C:\\Program Files (x86)\\Common Files\\microsoft shared\\DAO\\dao360.dll"
    ' {00025E01-0000-0000-C000-000000000046}\5.0\FLAGS   @="0"
    ' {00025E01-0000-0000-C000-000000000046}\5.0\HELPDIR @="C:\\Program Files (x86)\\Common Files\\microsoft shared\\DAO\\"

    ' HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Classes\TypeLib
    ' {00025E01-0000-0000-C000-000000000046}
    ' {00025E01-0000-0000-C000-000000000046}\5.0         @="Microsoft DAO 3.6 Object Library"
    ' {00025E01-0000-0000-C000-000000000046}\5.0\0
    ' {00025E01-0000-0000-C000-000000000046}\5.0\0\win32 @="C:\\Program Files (x86)\\Common Files\\microsoft shared\\DAO\\dao360.dll"
    ' {00025E01-0000-0000-C000-000000000046}\5.0\FLAGS   @="0"
    ' {00025E01-0000-0000-C000-000000000046}\5.0\HELPDIR @="C:\\Program Files (x86)\\Common Files\\microsoft shared\\DAO\\"
#End Region

#Region "Declarations"
    Public Event Notify(ByVal Msg As String, ByVal LeadingSpaces As Integer, ByVal Type As Transcript.MsgType)

    Public Enum OSTyp
        Win7
        Other
    End Enum

    Public Enum Hkey
        HKEY_CLASSES_ROOT
        HKEY_LOCAL_MACHINE
    End Enum

    Protected ReadOnly CLSID As New ConstStringArray _
    (New String() {"{00000100-0000-0010-8000-00AA006D2EA4}" _
                 , "{00000101-0000-0010-8000-00AA006D2EA4}" _
                 , "{00000103-0000-0010-8000-00AA006D2EA4}" _
                 , "{00000104-0000-0010-8000-00AA006D2EA4}" _
                 , "{00000105-0000-0010-8000-00AA006D2EA4}" _
                 , "{00000106-0000-0010-8000-00AA006D2EA4}" _
                 , "{00000107-0000-0010-8000-00AA006D2EA4}" _
                 , "{00000108-0000-0010-8000-00AA006D2EA4}" _
                 , "{00000109-0000-0010-8000-00AA006D2EA4}"})

    Protected ReadOnly ClsIds As New ConstStringArray _
    (New String() {"dao.DBEngineClass", "dao.PrivDBEngineClass" _
                 , "dao.TableDefClass", "dao.FieldClass", "dao.IndexClass" _
                 , "dao.GroupClass", "dao.UserClass", "dao.QueryDefClass", "dao.RelationClass"})

    Protected ReadOnly ProgIds As New ConstStringArray _
    (New String() {"DAO.DBEngine.36", "DAO.PrivateDBEngine.36" _
                 , "DAO.TableDef.36", "DAO.Field.36", "DAO.Index.36" _
                 , "DAO.Group.36", "DAO.User.36", "DAO.QueryDef.36", "DAO.Relation.36"})

    Protected ReadOnly DaoLibKey As String = "{00025E01-0000-0000-C000-000000000046}"
    Protected ReadOnly DaoDll As String = "DAO360.dll"
    Protected ReadOnly DaoVer As String = "5.0"

    Protected DaoPath As String
    Protected OsType As OSTyp
    Protected RegOk As Boolean

#End Region

#Region "Public Methods"
    Public Sub New()
        Dim ComPrgFiles As String = System.Environment.GetEnvironmentVariable("CommonProgramFiles")
        Dim ComPrgFilesX86 As String = System.Environment.GetEnvironmentVariable("CommonProgramFiles(x86)")

        If Not IsNothing(ComPrgFilesX86) Then
            Me.OsType = OSTyp.Win7
            Me.DaoPath = ComPrgFilesX86 + "\Microsoft Shared\DAO\" + Me.DaoDll
        Else
            Me.OsType = OSTyp.Other
            Me.DaoPath = ComPrgFiles + "\Microsoft Shared\DAO\" + Me.DaoDll
        End If

        Me.RegOk = False
    End Sub

    Public Sub ChkDAO360DLL()
        RaiseEvent Notify("Checking DAO Shared Library: " + Me.DaoDll, 0, Transcript.MsgType.Txt)

        If Not My.Computer.FileSystem.FileExists(Me.DaoPath) Then
            RaiseEvent Notify("Microsoft Shared Library 'Dao360.dll' does NOT exist!", 5, Transcript.MsgType.Err)
            RaiseEvent Notify("Path: " + Me.DaoPath, 10, Transcript.MsgType.Txt)
            Exit Sub
        End If

        ' reregister DLL 
        Me.RegisterDll()

        ' check registration
        Me.RegOk = Me.ChkDaoTypeLibs()

        If Me.RegOk Then
            RaiseEvent Notify("Shared Library 'Dao360.dll' registered correctly!", 0, Transcript.MsgType.Txt)
            RaiseEvent Notify("", 0, Transcript.MsgType.Txt)
        Else
            RaiseEvent Notify("Microsoft Shared Library 'Dao360.dll' NOT properly registered!", 0, Transcript.MsgType.Err)
            RaiseEvent Notify("Path: " + Me.DaoPath, 5, Transcript.MsgType.Txt)
            RaiseEvent Notify("Contact your System Administrator ... ", 5, Transcript.MsgType.Txt)
            RaiseEvent Notify("", 0, Transcript.MsgType.Txt)
        End If
    End Sub
#End Region

#Region "Check DAO TypeLib Version"
    Protected Function ChkDaoTypeLibs() As Boolean
        Dim RegOk As Boolean = True
        Dim KeySubStr As String = String.Empty

        ' check common locations XP/Windows7
        ' HKEY_CLASSES_ROOT\TypeLib\{00025E01-0000-0000-C000-000000000046}
        KeySubStr = "TypeLib\" + Me.DaoLibKey
        RegOk = Me.ChkDaoTypeLib(Hkey.HKEY_CLASSES_ROOT, KeySubStr)

        ' HKEY_LOCAL_MACHINE\SOFTWARE\Classes\TypeLib\{00025E01-0000-0000-C000-000000000046}
        KeySubStr = "SOFTWARE\Classes\TypeLib\" + Me.DaoLibKey
        RegOk = Me.ChkDaoTypeLib(Hkey.HKEY_LOCAL_MACHINE, KeySubStr)

        If Me.OsType = OSTyp.Win7 Then
            ' HKEY_CLASSES_ROOT\Wow6432Node\TypeLib\{00025E01-0000-0000-C000-000000000046}
            KeySubStr = "Wow6432Node\TypeLib\" + Me.DaoLibKey
            RegOk = Me.ChkDaoTypeLib(Hkey.HKEY_CLASSES_ROOT, KeySubStr)

            ' HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Wow6432Node\TypeLib\{00025E01-0000-0000-C000-000000000046}
            KeySubStr = "SOFTWARE\Classes\Wow6432Node\TypeLib\" + Me.DaoLibKey
            RegOk = Me.ChkDaoTypeLib(Hkey.HKEY_LOCAL_MACHINE, KeySubStr)

            ' HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Classes\TypeLib\{00025E01-0000-0000-C000-000000000046}
            KeySubStr = "SOFTWARE\Wow6432Node\Classes\TypeLib\" + Me.DaoLibKey
            RegOk = Me.ChkDaoTypeLib(Hkey.HKEY_LOCAL_MACHINE, KeySubStr)
        End If

        Return RegOk
    End Function

    Protected Function ChkDaoTypeLib(ByVal HKEY As Hkey, ByVal KeySubStr As String) As Boolean
        Dim RegOk As Boolean = True

        RaiseEvent Notify("... checking " + HKEY.ToString + "\" + KeySubStr, 5, Transcript.MsgType.Txt)
        If Not Me.ChkTypeLib(HKEY, KeySubStr) Then
            RaiseEvent Notify("... register " + HKEY.ToString + "\" + KeySubStr, 5, Transcript.MsgType.Txt)
            Me.RegDaoTypeLib(HKEY, KeySubStr)
            RegOk = Me.ChkTypeLib(HKEY, KeySubStr)
            If Not RegOk Then Me.SendBadRegMessage(HKEY.ToString + "\" + KeySubStr)
        Else
            RaiseEvent Notify("Registration OK!", 10, Transcript.MsgType.Txt)
        End If

        Return RegOk
    End Function

    Protected Function ChkTypeLib(ByVal HKEY As Hkey, ByVal KeySubStr As String) As Boolean
        Dim RegKey As RegistryKey

        RegKey = Me.FindRegSubKey(HKEY, KeySubStr + "\5.0\0", "win32")

        If IsNothing(RegKey) Then Return False

        ' check value (Default)
        Dim RegVal As Object = RegKey.GetValue(String.Empty) ' empty string means (Default)
        RegKey.Close()
        RegKey = Nothing

        If IsNothing(RegVal) Then Return False
        If RegVal.ToString.ToLower = Me.DaoPath.ToLower Then Return True
        Return False
    End Function

    Protected Sub RegDaoTypeLib(ByVal HKEY As Hkey, ByVal KeySubStr As String)
        Dim RegVal As Object
        Dim RegKeyUID As RegistryKey
        Dim RegKeyVer As RegistryKey
        Dim RegKeyZer As RegistryKey
        Dim RegKeyNew As RegistryKey

        ' {00025E01-0000-0000-C000-000000000046}
        ' {00025E01-0000-0000-C000-000000000046}\5.0         @="Microsoft DAO 3.6 Object Library"
        ' {00025E01-0000-0000-C000-000000000046}\5.0\0
        ' {00025E01-0000-0000-C000-000000000046}\5.0\0\win32 @="C:\\Program Files (x86)\\Common Files\\microsoft shared\\DAO\\dao360.dll"
        ' {00025E01-0000-0000-C000-000000000046}\5.0\FLAGS   @="0"
        ' {00025E01-0000-0000-C000-000000000046}\5.0\HELPDIR @="C:\\Program Files (x86)\\Common Files\\microsoft shared\\DAO\\"

        ' {00025E01-0000-0000-C000-000000000046}
        RegKeyUID = Me.FindRegSubKey(HKEY, KeySubStr, "", True)
        If IsNothing(RegKeyUID) Then Exit Sub

        ' {00025E01-0000-0000-C000-000000000046}\5.0 
        RegKeyVer = Me.FindRegSubKey(HKEY, KeySubStr, "5.0", True)
        If IsNothing(RegKeyVer) Then
            RegKeyVer = Me.AddRegSubKey(RegKeyUID, "5.0")
            If IsNothing(RegKeyVer) Then Exit Sub
        End If

        ' check default value @="Microsoft DAO 3.6 Object Library"
        RegVal = Me.AddRegKeyValue(RegKeyVer, "", "Microsoft DAO 3.6 Object Library", RegistryValueKind.String)

        '------------------------------------------------------------------------

        ' {00025E01-0000-0000-C000-000000000046}\5.0\0
        RegKeyZer = Me.FindRegSubKey(HKEY, KeySubStr, "5.0\0", True)
        If IsNothing(RegKeyZer) Then
            RegKeyZer = Me.AddRegSubKey(RegKeyVer, "0")
            If IsNothing(RegKeyZer) Then Exit Sub
        End If

        '------------------------------------------------------------------------

        ' {00025E01-0000-0000-C000-000000000046}\5.0\0\win32
        RegKeyNew = Me.FindRegSubKey(HKEY, KeySubStr, "5.0\0\win32", True)
        If IsNothing(RegKeyNew) Then
            RegKeyNew = Me.AddRegSubKey(RegKeyZer, "win32")
            If IsNothing(RegKeyNew) Then Exit Sub
        End If

        ' check default value @="Microsoft DAO 3.6 Object Library"
        RegVal = Me.AddRegKeyValue(RegKeyNew, "", Me.DaoPath, RegistryValueKind.String)

        '------------------------------------------------------------------------

        ' {00025E01-0000-0000-C000-000000000046}\5.0\FLAGS
        RegKeyNew = Me.FindRegSubKey(HKEY, KeySubStr, "5.0\FLAGS", True)
        If IsNothing(RegKeyNew) Then
            RegKeyNew = Me.AddRegSubKey(RegKeyVer, "FLAGS")
            If IsNothing(RegKeyNew) Then Exit Sub
        End If

        ' check default value @="Microsoft DAO 3.6 Object Library"
        RegVal = Me.AddRegKeyValue(RegKeyNew, "", "0", RegistryValueKind.String)

        '------------------------------------------------------------------------

        ' {00025E01-0000-0000-C000-000000000046}\5.0\FLAGS
        RegKeyNew = Me.FindRegSubKey(HKEY, KeySubStr, "5.0\HELPDIR", True)
        If IsNothing(RegKeyNew) Then
            RegKeyNew = Me.AddRegSubKey(RegKeyVer, "HELPDIR")
            If IsNothing(RegKeyNew) Then Exit Sub
        End If

        ' check default value @="Microsoft DAO 3.6 Object Library"
        RegVal = Me.AddRegKeyValue(RegKeyNew, "", Path.GetDirectoryName(Me.DaoPath), RegistryValueKind.String)
    End Sub
#End Region

#Region "Registry Methods"
    Protected Function AddRegSubKey(ByVal Key As RegistryKey, ByVal SubKey As String) As RegistryKey
        Dim NewKey As RegistryKey
        Try
            NewKey = Key.CreateSubKey(SubKey, RegistryKeyPermissionCheck.Default)
        Catch ex As Exception
            Me.SendAddRegMessage(ex.Message)
            NewKey = Nothing
        End Try
        Return NewKey
    End Function

    Protected Function AddRegKeyValue(ByVal RegKey As RegistryKey, ByVal ValueName As String, ByVal ValueData As String, ByVal Type As RegistryValueKind) As Object
        Dim RegVal As Object
        RegKey.SetValue(ValueName, ValueData, Type)
        RegVal = Me.GetRegKeyValue(RegKey)
        If IsNothing(RegVal) Then Return Nothing
        Return RegVal
    End Function

    Protected Function GetRegKeyValue(ByVal RegKey As RegistryKey) As Object
        Dim RegVal As Object = RegKey.GetValue("") ' empty string means (Default)
        If IsNothing(RegVal) Then Return Nothing
        Return RegVal
    End Function

    Protected Function FindRegSubKey(ByVal HKEY As Hkey, ByVal Key As String, ByVal SubKey As String, Optional ByVal Writable As Boolean = False) As RegistryKey
        Dim KeyStr As String, Key2Find As RegistryKey = Nothing

        KeyStr = Key + "\" + SubKey

        If Key = String.Empty Then KeyStr = SubKey
        If SubKey = String.Empty Then KeyStr = Key
        If KeyStr = String.Empty Then Return Nothing

        Try
            Select Case HKEY
                Case RegDAO.Hkey.HKEY_CLASSES_ROOT
                    Key2Find = My.Computer.Registry.ClassesRoot.OpenSubKey(KeyStr, Writable)
                Case RegDAO.Hkey.HKEY_LOCAL_MACHINE
                    Key2Find = My.Computer.Registry.LocalMachine.OpenSubKey(KeyStr, Writable)
            End Select
        Catch ex As Exception
            Me.SendGetRegMessage(ex.Message)
        End Try

        If IsNothing(Key2Find) Then Return Nothing
        Return Key2Find
    End Function
#End Region

#Region "Protected Methods"
    Protected Sub SendBadRegMessage(ByVal RegKeyStr As String)
        RaiseEvent Notify("Bad or missing Registration!", 5, Transcript.MsgType.Err)
        RaiseEvent Notify(RegKeyStr, 5, Transcript.MsgType.Txt)
        RaiseEvent Notify("", 0, Transcript.MsgType.Txt)
    End Sub

    Protected Sub SendAddRegMessage(ByVal ExMessage As String)
        RaiseEvent Notify("Registry Exception !", 5, Transcript.MsgType.Err)
        RaiseEvent Notify(ExMessage, 5, Transcript.MsgType.Txt)
        RaiseEvent Notify("", 0, Transcript.MsgType.Txt)
    End Sub

    Protected Sub SendGetRegMessage(ByVal ExMessage As String)
        RaiseEvent Notify("Registry Exception !", 5, Transcript.MsgType.Err)
        RaiseEvent Notify(ExMessage, 5, Transcript.MsgType.Txt)
        RaiseEvent Notify("", 0, Transcript.MsgType.Txt)
    End Sub

    Protected Sub RegisterDll()
        Dim RegDll As Process, RegSvr As String
        Dim WinRoot As String = System.Environment.GetEnvironmentVariable("SystemRoot")

        RegSvr = WinRoot + "\System32\regsvr32.exe"

        If Me.OsType = OSTyp.Win7 Then
            RegSvr = WinRoot + "\SysWOW64\regsvr32.exe"
        End If

        RegDll = New System.Diagnostics.Process()
        RegDll.StartInfo.CreateNoWindow = True
        RegDll.StartInfo.UseShellExecute = True

        RegDll.StartInfo.FileName = RegSvr
        RegDll.StartInfo.Arguments = "-s " + Chr(34) + Me.DaoPath + Chr(34)

        Try
            'C:\WINDOWS\System32\regsvr32 "C:\Program Files\Common Files\Microsoft Shared\DAO\dao360.dll"
            'C:\WINDOWS\SysWOW64\regsvr32 "C:\Program Files (x86)\Common Files\Microsoft Shared\DAO\dao360.dll"

            ' start process
            RegDll.Start()

            'Wait until the process passes back an exit code 
            RegDll.WaitForExit()

            'Free resources associated with this process
            RegDll.Close()
        Catch ex As Exception
            RaiseEvent Notify("Unable to register 'Dao360.dll' ...", 5, Transcript.MsgType.Err)
            RaiseEvent Notify(ex.Message, 10, Transcript.MsgType.Txt)
            RaiseEvent Notify("", 0, Transcript.MsgType.Txt)
        Finally
            RegDll = Nothing
        End Try
    End Sub
#End Region

#Region "ConstStringArray"
    Public Class ConstStringArray
        ' protected reference
        Protected Strings() As String

        Public Sub New(ByVal Strings() As String)
            ' must be cloned
            Me.Strings = CType(Strings.Clone(), String())
        End Sub

        Default Public ReadOnly Property Item(ByVal Index As Integer) As String
            Get
                Return Strings(Index)
            End Get
        End Property

        Public ReadOnly Property Count() As Integer
            Get
                Return UBound(Strings) + 1
            End Get
        End Property
    End Class
#End Region

End Class
