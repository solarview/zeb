
''' <summary>
''' 존과 벽체를 편집하는 클래스
''' </summary>
Public Class ZoneEdit
#Region "Instance Variables"
    '존 목록. (주의!!!)별도로 생성하지 않고 프로그램에서 사용하는 것을 참조해서 사용한다.
    Private _Zones As Zones '
    Private _Envelopes As Envelopes

#End Region

#Region "Constructors"
    ''' <summary>
    ''' 편집할 존과 외피 목록을 가지고 시작한다.
    ''' 주의 : 존과 외피 목록 없이 생성되지 않도록 한다.
    ''' </summary>
    ''' <param name="ZoneList"></param>
    ''' <param name="EnvelopeList"></param>
    Sub New(ZoneList As Zones, EnvelopeList As Envelopes)
        _Zones = ZoneList
        _Envelopes = EnvelopeList
    End Sub
#End Region

#Region "Methods"

    Sub AddZone(ZoneName As String)

        If _Zones.ContainsKey(ZoneName) Then '등록된 이름이 있을 경우

        Else '등록된 이름이 없을 경우
        End If
    End Sub

    Sub AddZone(theZone As Zone)
        _Zones.AddZone(theZone.Name, theZone)
    End Sub

    Sub DeleteZone(ZoneName As String)

        If _Zones.ContainsKey(ZoneName) Then '등록된 이름이 있을 경우
            '이 존과 연관된 외피가 있을 경우

            '이 존관 연관된 외피가 없을 경우

        Else '등록된 이름이 없을 경우
        End If
    End Sub

    Sub DeleteEnvelope(EnvelopeName As String)

        If _Envelopes.ContainsKey(EnvelopeName) Then '등록된 이름이 있을 경우
            '이 외피와 연관된 존이 있을 경우

            '이 외피와 연관된 존이 없을 경우

        Else '등록된 이름이 없을 경우
        End If
    End Sub
#End Region

End Class
