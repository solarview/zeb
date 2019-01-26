Public Class Envelope
#Region "Instance Variables"
    '이름(유일무이한 이름)
    Private _Name As String
    '외피 유형 => 일반 벽이든 창이 되게 한다. (상속으로 처리)
    '인접면 설정 => 외기든 다른 존이든 설정하게 한다.


    '방위
    '기울기
    '면적
    '개수
    '열관류율
    Protected _Uvalue As Double '벽체에서는 사용자 입력, 창호에서는 계산
    '흡수율
    '인접 존 선택
    Private _InsideZone As Zone
    Private _OutsideZone As Zone

#End Region
    Public Property UValue As Double
        Get
            Return _Uvalue
        End Get
        Set(value As Double)
            _Uvalue = value
        End Set
    End Property
#Region "methods"
    'Protected Function GetUValue()
#End Region

End Class

Public Class Envelopes
#Region "Instance Variables"
    Private _Envelopes As Generic.Dictionary(Of String, Envelope)

#End Region
#Region "Methods"
    Public Sub AddEnvelope(newName As String, NewEnvelope As Envelope)
        '기존에 등록된 벽체의 이름과 중복이 되는지를 검토해야함.
        If _Envelopes.ContainsKey(newName) Then
            MsgBox("다른 이름을 지어주십시오.", "중복된 이름")
            Exit Sub
        End If
        '
        _Envelopes.Add(newName, NewEnvelope)
    End Sub

    Function ContainsKey(key As String) As Boolean
        Return _Envelopes.ContainsKey(key)
    End Function

    Public Sub Import(FileName As String)

        For Each ttt As KeyValuePair(Of String, Envelope) In _Envelopes

        Next
    End Sub
    Public Sub Export(filename As String)

    End Sub
#End Region
End Class

Public Class Wall
    Inherits Envelope


End Class

Public Class Window
    Inherits Envelope
    '창호 길이
    '창호 높이
    '하인방 높이
    '프레임 열관류율
    '유리면적비
    '유리종류
    '차양설치위치
    '차양종류
    '차양색상
    '차양작동방식 -이용일
    '유리열관류율
    '창호 열관류율
    '가시광선 투과율
    '태양열취득률
    '태양열취득률
    '수직 장애물
    '수직 오버행
    '수평 핀(좌측)
    '수평 핀(우측)
    '겨울철 음영감소계수
    '여름철 음영감소계수
#Region "Properties"
    Public Overloads ReadOnly Property UValue As Double
        Get
            '별도로 계산한다.
            '유리열관류율과 프레임열관류율과의 면적가중평균 -> 창호열관유율 계산
            Return _Uvalue
        End Get
    End Property
#End Region

    'test
End Class