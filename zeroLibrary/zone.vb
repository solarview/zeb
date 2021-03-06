﻿Public Class Zone
#Region "존정보입력"

    '1.존정보입력
    '존 이름 = 유일무이한 이름이어야 한다.
    Private _Name As String
    '
    '용도 프로필
    Private _UsageType As zUsageType = zUsageType.대규모사무실사용자입력
    '냉난방여부
    Private _ConditioningMode As zConditioningMode = zConditioningMode.냉방난방
    '공조방식
    Private _HVACMode As zHVACMode = zHVACMode.냉방난방
    '순바닥면적
    Private _A_NF As Double
    '천장고
    Private _h_G As Double = 2.9
    '순 체적
    Private _V As Double
    '열용량
    Private _HeatCapacity As Double
    '야간운전방식
    Private _RunningModeNight As zRunningMode = zRunningMode.운전정지
    '주말운전방식
    Private _RunningModeWeekend As zRunningMode = zRunningMode.운전정지
    '건물 자동화 계수
    Private _AutomationFactor As zAutomationFactor = zAutomationFactor.C건물자동제어
#End Region

#Region "<존의 최하층>"
    '지중 접합 유형
    Private _GroundContactType As zGroundContactType = zGroundContactType.비냉난방지하실
    '냉난방바닥면적 = 최하층 외피 바닥면적
    Private _A As Double
    '지면에 노출된 둘레
    Private _P As Double
    '최하층 벽체 두께
    Private _W As Double



    '최하층 존의 바닥 열관류율
    '존의 벽/바닥 선형열관류율
    '지면 아래 매립 깊이
    '지면 아래 벽체 열관류율
    '지면 위 지하실의 벽체 열관류율
    '지하실 바닥 열관류율
    '지면 위 지하층 지하실의 외벽 높이
    '최하층 지하실 환기횟수
    '최하층 지하실 체적
    '보강단열재 열전도율
    '보강단열 두께
    '수평보강단열 길이
    '수직보강단열 길이
    '지하수깊이
    '지하수유속
#End Region
    '
    Private _zEnvelopes As Envelopes

#Region "Constructors"

#End Region
#Region "Properties"
    Public Property Name As String
        Get
            Return _Name
        End Get
        Set(value As String)
            _Name = value
        End Set
    End Property
    Public Property A_NF As Double
        Get
            Return _A_NF
        End Get
        Set(ByVal value As Double)
            _A_NF = value
        End Set
    End Property
#End Region
    Sub test()
        Dim Abuilding As Building
        Dim Bbuilding As Building
        '
        'default constructor 사용법
        Abuilding = New Building()
        'copy constructor 사용법
        Bbuilding = New Building(Abuilding)
    End Sub
    Public Sub Import(FileName As String)

    End Sub
    Public Sub Export(filename As String)

    End Sub
    Sub AddEnvelope(newEnvelope As Envelope)
        '_zEnvelopes.AddEnvelope()
    End Sub
End Class




''' <summary>
''' 존의 집합체를 나타내는 클래스
''' </summary>
Public Class Zones

#Region "Instance Variables"
    Private _zones As Generic.Dictionary(Of String, Zone) ' key 의 형식이 문자열, value의 형식이 Zone
#End Region

#Region "Constructors"
    Public Sub New()
        _zones = New Generic.Dictionary(Of String, Zone)
    End Sub
#End Region

#Region "Properties"
    ''' <summary>
    ''' 등록된 존을 반환한다.
    ''' Default Property로 선언해서 Zones.Item("존의 이름")과 동일하게 
    ''' Zones("존의 이름")과 같이 사용할 수 있음
    ''' </summary>
    ''' <param name="ZoneName">등록된 존의 이름[string]</param>
    ''' <returns>등록된 존[zone]</returns>
    Default Public ReadOnly Property Item(ZoneName As String) As Zone
        Get
            Return _zones.Item(ZoneName)
        End Get
    End Property
#End Region

#Region "Methods"
    ''' <summary>
    ''' 모든 존의 순바닥면적을 합산하여 반환한다.
    ''' </summary>
    ''' <returns>모든 존의 순바닥면적의 합</returns>
    Public Function A_NF() As Double
        Dim sum As Double = 0
        For Each ttt As KeyValuePair(Of String, Zone) In _zones
            sum += ttt.Value.A_NF 'sum = sum + ttt.Value.A_NF
        Next
        Return sum
    End Function

    ''' <summary>
    ''' 새로운 존을 추가한다.
    ''' </summary>
    ''' <param name="newZoneName">새 존의 이름. 유일무이한 이름이어야 한다.</param>
    ''' <param name="NewZone">새 존의 객체</param>
    Public Sub AddZone(newZoneName As String, NewZone As Zone)
        '기존에 등록된 존의 이름과 중복이 되는지를 검토해야함.
        If _zones.ContainsKey(newZoneName) Then
            MsgBox("다른 이름을 지어주십시오.", "중복된 이름")
            Exit Sub
        End If
        '
        _zones.Add(newZoneName, NewZone)
    End Sub

    Function ContainsKey(key As String) As Boolean
        Return _zones.ContainsKey(key)
    End Function

    Public Sub Import(FileName As String)

        For Each ttt As KeyValuePair(Of String, Zone) In _zones

        Next
    End Sub
    Public Sub Export(filename As String)

    End Sub
#End Region
End Class
