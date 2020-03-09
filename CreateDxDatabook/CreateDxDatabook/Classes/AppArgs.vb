Imports System.Runtime.InteropServices

Public Class AppArgs
    Private _Argc As Integer
    Private _Argv As List(Of CmdArg)

    Public ReadOnly Property Argc() As Integer
        Get
            Return Me._Argc
        End Get
    End Property

    Public ReadOnly Property Argv() As List(Of CmdArg)
        Get
            Return Me._Argv
        End Get
    End Property

    Public Sub New()
        Me._Argv = New List(Of CmdArg)
        Me._Argc = Me._Argv.Count
        Me.GetArgs()
    End Sub

    Private Sub GetArgs()
        Dim Arg As CmdArg = Nothing
        For Each param As String In My.Application.CommandLineArgs
            Select Case param.IndexOf("-")
                Case 0
                    If Not IsNothing(Arg) Then Me._Argv.Add(Arg)
                    Arg = New CmdArg
                    Me._Argc += 1
                    Arg.Name = param.ToString.ToLower()
                    Arg.Value = ""
                Case Else
                    Arg.Value = param.ToString
            End Select
        Next param
        If Not IsNothing(Arg) Then Me._Argv.Add(Arg)
    End Sub
End Class

Public Class CmdArg
    Public Name As String
    Public Value As String
End Class

Public Class AppConsole
    <DllImport("kernel32.dll")> Private Shared Function AllocConsole(ByVal dwProcessId As Int32) As Boolean
    End Function

    <DllImport("kernel32.dll")> Private Shared Function AttachConsole(ByVal dwProcessId As Int32) As Boolean
    End Function

    <DllImport("kernel32.dll")> Private Shared Function FreeConsole() As Boolean
    End Function

    Public Enum Mode
        Create
        Attach
    End Enum

    Protected Const ATTACH_PARENT_PROCESS As Integer = -1

    Public Shared Function CreateConsole(ByVal m As Mode, Optional ByVal Title As String = "Console", Optional ByVal TextColor As System.ConsoleColor = ConsoleColor.White) As Boolean
        Dim ConsoleOk As Boolean
        Try
            Select Case m
                Case Mode.Attach
                    ConsoleOk = AttachConsole(ATTACH_PARENT_PROCESS)
                Case Mode.Create
                    ConsoleOk = AllocConsole(ATTACH_PARENT_PROCESS)
                Case Else
                    ConsoleOk = AttachConsole(ATTACH_PARENT_PROCESS)
            End Select

            If ConsoleOk Then
                Console.Title = Title

                Console.ForegroundColor = TextColor
                Console.CursorVisible = False

                Console.BufferWidth = 512
                Console.BufferHeight = 2048
            End If

            Return ConsoleOk
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Shared Function ForgetConsole() As Boolean
        Try
            FreeConsole()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
End Class
