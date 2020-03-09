Public Class Time
    'Public Structure MyTime
    '    Public Hrs As String ' 08
    '    Public Min As String ' 05
    '    Public Sec As String ' 03
    'End Structure

    'Public Structure MyDate
    '    Public YY As String ' Year - 08
    '    Public YS As String ' Year - 2008
    '    Public MM As String ' Month - 04
    '    Public MN As String ' Month - Juni
    '    Public TT As String ' Day - 09
    '    Public TS As String ' Day - Monday
    'End Structure

    Public Shared Function CreateTimeStamp() As String
        Return Format$(Now, "dd.mm.yy") + "_" + Format$(Now, "hh") + "h" + Format$(Now, "mm")
    End Function

    Public Shared Function FormatTimeSpan(ByVal Duration As TimeSpan) As String
        Dim Hours As String, Minutes As String, Seconds As String
        Hours = Format(Duration.Hours, "#.##")
        Minutes = Format(Duration.Minutes, "#.##")
        Seconds = Format(Duration.Seconds, "#.##")
        If Hours = "" Then Hours = "00"
        If Minutes = "" Then Minutes = "00"
        If Seconds = "" Then Seconds = "00"
        If Len(Hours) = 1 Then Hours = "0" + Hours
        If Len(Minutes) = 1 Then Minutes = "0" + Minutes
        If Len(Seconds) = 1 Then Seconds = "0" + Seconds
        Return Hours + ":" + Minutes + ":" + Seconds
    End Function

    Public Shared Function Duration(ByVal StartDateTime As Date, ByVal EndDateTime As Date) As String
        Dim Thrs As Single, Tmins As Single, Tsecs As Integer
        Dim i As Integer, Var_tmp_Hour As Integer, DTcheck As Single

        ' check elapsed days
        DTcheck = DateDiff("d", StartDateTime, EndDateTime)
        If DTcheck < 0 Then Return "00:00:00"

        ' get elapsed hours
        Var_tmp_Hour = Hour(EndDateTime)
        For i = 1 To CInt(DTcheck)
            Var_tmp_Hour = Var_tmp_Hour + 24
        Next

        Thrs = Var_tmp_Hour - Hour(StartDateTime)
        If Thrs < 0 Then Return "00:00:00"

        Tmins = Minute(EndDateTime) - Minute(StartDateTime)
        If Tmins < 0 Then Tmins = Tmins + 60 : Thrs = Thrs - 1

        Tsecs = Second(EndDateTime) - Second(StartDateTime)
        If Tsecs < 0 Then
            Tsecs = Tsecs + 60
            If Tmins > 0 Then
                Tmins = Tmins - 1
            Else
                Tmins = 59
            End If
        End If

        If Format(Thrs & ":" & Tmins & ":" & Tsecs, "HH:mm:ss") = "HH:mm:ss" Then
            Return "00:00:00"
        Else
            Return Format(Thrs & ":" & Tmins & ":" & Tsecs, "HH:mm:ss")
        End If
    End Function
End Class
