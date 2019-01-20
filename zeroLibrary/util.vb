Imports System.IO

Public Module Util
    Public Const ERRORBASE As Integer = 32766

    Public Const SWP_NOMOVE As Short = 2
    Public Const SWP_NOSIZE As Short = 1
    Public Const flags As Boolean = SWP_NOMOVE Or SWP_NOSIZE
    Public Const HWND_TOPMOST As Short = -1
    Public Const HWND_NOTOPMOST As Short = -2
    '
    'This call will see whether or not there are any messages waiting to be executed in the calling thread.
    '이 함수는 Application.DoEvents()와 함께 사용할 것임.
    Public Declare Function GetInputState Lib "user32" () As Int32

    Function StripPath(ByRef FullFilename As String) As String
        '#'***************************** -사용법
        '#'      Begin Block Out
        '#'By: Yong-Yee Kim
        '#'On: 2001-11-26
        '#'*****************************
        '@Private Sub Form_Load()
        '@'Enter here the full path you want to strip from it the file name.
        '@MsgBox StripPath("c:\mydir\myfile.exe")
        '@End Sub
        '#'*****************************
        '#'      End Block Out
        '#'*****************************
        Dim x, ct As Short

        StripPath = FullFilename
        x = InStr(FullFilename, "\")
        '
        Do While x
            ct = x
            x = InStr(ct + 1, FullFilename, "\")
        Loop
        If ct > 0 Then StripPath = Mid(FullFilename, ct + 1)
    End Function

    'This function will check also if the given path is only directory.
    'Insert the following code to your form:

    Function GetExtension(ByVal FileName As String) As String

        Dim PthPos, ExtPos As Integer
        Dim i, j As Integer
        '
        For i = Len(FileName) To 1 Step -1
            If Mid(FileName, i, 1) = "." Then
                ExtPos = i
                For j = Len(FileName) To 1 Step -1
                    If Mid(FileName, j, 1) = "\" Then
                        PthPos = j
                        Exit For
                    End If
                Next j
                Exit For
            End If
        Next i

        If PthPos > ExtPos Then
            'Exit Function
            Return ""
        Else
            If ExtPos = 0 Then Return "" 'Exit Function
            Return Mid(FileName, ExtPos + 1, Len(FileName) - ExtPos)
        End If
    End Function


    '문자열을 단어로 분리하기 위하여 사용되는 구분 글자들
    Private Separators() As Char = {","c, ":"c, "|"c, ";"c, "'"c, "="c}
    '
    Sub GetWords(ByVal TextLine As String, ByRef Word() As String)
        Dim count As Short
        GetWords(TextLine, Word, count)
    End Sub
    '
    Sub GetWords(ByVal TextLine As String, ByRef Words() As String, ByRef WordCount As Short)
        '
        ' 입력받은 문자열을 쉼표와 빈칸에 따라 단어를 나누어 준다.
        ' divide character string into separate words according to period and blank space.
        '
        'Revision History
        '  Written by Yong-Yee Kim, 2005/11/25
        '**********************************************************************

        '1)문자열 시작과 끝의 공백을 제거한다.
        TextLine = TextLine.Trim()
        '
        '2)문자열을 구분자(쉼표)를 통하여 분할한다.
        Dim Tokens() As String = TextLine.Split(Separators)
        '
        '3)얼마나 많은 단어가 존재하는가?
        Dim i As Integer, nums As Integer = 0
        For i = 0 To Tokens.Length - 1
            If Tokens(i).Trim.Length > 0 Then '빈 문자열의 단어는 제외시킴
                nums += 1
            End If
        Next i
        '
        '4)단어수만큼 배열을 할당
        ReDim Words(nums - 1)
        '
        '5)단어를 반환한다.
        WordCount = 0
        For j As Integer = 0 To Tokens.Length - 1
            If Tokens(j).Trim.Length > 0 Then
                Words(WordCount) = Tokens(j).Trim
                WordCount += 1
            End If
        Next j
        '
    End Sub

    ''' <summary>
    ''' 문자열을 단어로 분리한다.
    ''' </summary>
    ''' <param name="TextLine">단어로 분리할 문자열</param>
    ''' <param name="Separators">단어로 구분하기 위한 문자(구분자)</param>
    ''' <returns>단어들의 배열</returns>
    ''' <remarks>구분자를 지정하지 않으면 기본 구분자(쉼표, 콜론, 파이프, 세미콜론, 큰따옴표, 작은따옴표)를 사용한다.</remarks>
    Function GetWords(ByVal TextLine As String, ByVal ParamArray Separators() As Char) As String()
        '
        ' 입력받은 문자열을 쉼표와 빈칸에 따라 단어를 나누어 준다.
        ' divide character string into separate words according to period and blank space.
        '
        'Revision History
        '  Written by Yong-Yee Kim, 2005/11/25
        '**********************************************************************
        '
        '0)지역변수 선언
        Dim MySeparators() As Char = {","c, ":"c, "|"c, ";"c, """"c, "'"c, "="c} '기본 구분자 = 쉼표, 콜론, 파이프, 세미콜론, 큰따옴표, 작은따옴표, 등호
        Dim Words() As String = {}
        '
        '1)문자열 시작과 끝의 공백을 제거한다.
        TextLine = TextLine.Trim()
        '
        '2)문자열을 구분자(쉼표)를 통하여 분할한다.
        If Separators.Length > 0 Then
            MySeparators = Separators
        End If
        Dim Tokens() As String = TextLine.Split(MySeparators)

        '
        '3)얼마나 많은 단어가 존재하는가?
        Dim nums As Integer = 0
        For i As Integer = 0 To Tokens.Length - 1
            If Tokens(i).Trim.Length > 0 Then '빈 문자열의 단어는 제외시킴
                nums += 1
            End If
        Next i
        '
        '4)단어수만큼 배열을 할당
        ReDim Words(nums - 1)
        '
        '5)단어를 반환한다.
        Dim WordCount As Integer = 0
        For j As Integer = 0 To Tokens.Length - 1
            If Tokens(j).Trim.Length > 0 Then
                Words(WordCount) = Tokens(j).Trim
                WordCount += 1
            End If
        Next j
        '
        Return Words
    End Function

    ''' <summary>
    ''' 문자열이 주석이 아니거나 널이 아니면 참값을 반환
    ''' 즉, 실질적인 정보를 갖고 있으면 참
    ''' </summary>
    ''' <param name="Words">문자열</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function NotComment(ByVal Words As String) As Boolean
        Return (Not Words.TrimStart.StartsWith("*")) AndAlso (Not Words.Trim.Equals(""))
    End Function


    ''' <summary>
    ''' 정수로 된 시간을 문자로 반환한다.
    ''' </summary>
    ''' <param name="myTime">단위가 분(minute)인 정수형 시간</param>
    ''' <returns>**시간**분으로 표현되는 문자열</returns>
    ''' <remarks></remarks>
    Function GetTimeToString(ByVal myTime As Integer) As String
        '
        Dim myHour As Integer = myTime \ 60
        Dim myMinute As Integer = myTime Mod 60
        '
        Return myHour.ToString("D1") & "시간 " & myMinute.ToString("D2") & "분"
    End Function

End Module



''' <summary>
''' 경과시간을 계산하는 클래스(calculating elapsed time)
''' </summary>
''' <remarks>
''' </remarks>
Public Class SV_ElapsedTime
    Private _StartTime As Integer
    '
    Public Sub New()
        MyBase.New()
        '
        Call SetStartTime()
    End Sub
    '
    Public Sub SetStartTime()
        _StartTime = Environment.TickCount
    End Sub
    '
    ReadOnly Property MilliSecondsElapsed() As Double
        Get
            Return (Environment.TickCount - _StartTime)
        End Get
    End Property
    '
    ReadOnly Property SecondsElapsed() As Double
        Get
            Return MilliSecondsElapsed / 1000
        End Get
    End Property
    '
    ReadOnly Property MinutesElapsed() As Double
        Get
            Return SecondsElapsed / 60
        End Get
    End Property
End Class
