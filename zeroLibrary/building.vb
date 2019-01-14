'<변수 명명 규칙>
' 1)DIN 규정에 나온 변수명을 그대로 사용한다.
' 2)변수명의 마침표(.)는 밑줄(_)로 한다.
' 3)내부변수의 시작은 밑줄(_)로 한다.
'</변수 명명 규칙>

''' <summary>
''' 건물을 정보를 갖는 클래스
''' </summary>
Public Class BuildingInfo
#Region "Instance Variables"
    '1.건물 정보 입력
    '건물 유형
    Private _BuildingType As zBuildingType
    '지역
    Private _Location As zLocationIndex
    '대지 조건
    Private _SoilType As zSoilType
    '
    '3.건물 개요
    '건물이름
    Private _Name As String
    '지역 (이미 위에서 사용한 것이 아닌가?)
    '고도차이
    Private _Altitude As Double
    '건물 주소
    Private _Address As String
    '층수
    Private _n_G As Integer
    '연면적
    Private _TotalFloorArea As Double
    '순바닥면적 net floor area = 모든 존의 바닥면적의 합
    Private _A_NF As Double
    '외피면적
    Private _SurfaceArea As Double
    '외부체적
    Private _V_e As Double
    Private _SurfaceAreaToVolumeRatio As Double
    '기하학적 특성값
    '길이
    Private _L_char As Double
    '너비
    Private _B_char As Double
    '
    '4.계산 결과 (출력변수)
    '난방 에너지 요구량
    Private _EnergyNeedforHeating As Double
    '냉방 에너지 요구량
    Private _EnergyNeedforCooling As Double
    '조명 에너지 요구량
    Private _EnergyNeedforLighting As Double
    '급탕 에너지 요구량
    Private _EnergyNeedforHotWater As Double
    '환기 에너지 요구량
    Private _EnergyNeedforVentilation As Double
    '전체 에너지 요구량
    Private _TotalEnergyNeed As Double
    '
    '다른 클래스와의 관계
    Private _zones As Zones
    '
#End Region

#Region "Constructors"
    Sub New()
        _zones = New Zones()
        SetDefaultValues()
    End Sub

    'copy constructor
    Sub New(otherBuilding As BuildingInfo)
        Me.New()
        '
        '<복제 과정>
        Me._SoilType = otherBuilding._SoilType
        '</복제 과정>
    End Sub
#End Region

#Region "Properties"
    Public Property BuildingType As zBuildingType
        Get
            Return _BuildingType
        End Get
        Set(value As zBuildingType)
            _BuildingType = value
        End Set
    End Property

    Public Property Location As zLocationIndex
        Get
            Return _Location
        End Get
        Set(value As zLocationIndex)
            _Location = value
        End Set
    End Property

    Public Property SiolType As zSoilType

    Public Property Name As String
        Get
            Return _Name
        End Get
        Set(value As String)
            _Name = value
        End Set
    End Property

    Public Property Altitude As Double

    '건물 주소
    Public Property Address As String
    '층수
    Public Property n_G As Integer
    '연면적
    Public Property TotalFloorArea As Double
    '순바닥면적 net floor area = 모든 존의 바닥면적의 합

    ''' <summary>
    ''' 건물에 등록된 모든 존들의 순바닥면적을 반환한다.
    ''' </summary>
    ''' <returns>순바닥면적[double]</returns>
    Public ReadOnly Property A_NF As Double
        Get
            Return _zones.A_NF()
        End Get
    End Property

    '외피면적
    Public Property SurfaceArea As Double
    '외부체적
    Public Property V_e As Double
    '외피면적체적비
    Public Property SurfaceAreaToVolumeRatio As Double
    '기하학적 특성값
    '길이
    Public Property L_char As Double
    '너비
    Public Property B_char As Double
#End Region

#Region "Methods"
    Public Sub SetDefaultValues()
        _BuildingType = zBuildingType.비주거
        _Location = zLocationIndex.서울운영규정
        _SoilType = zSoilType.모래또는자갈
        '
        _Name = "New Building"
        '고도차이
        _Altitude = 0.0
        '건물 주소
        _Address = "서울시"
        '층수
        _n_G = 0.0
        '연면적
        _TotalFloorArea = 0.0
        '순바닥면적 net floor area = 모든 존의 바닥면적의 합
        _A_NF = 0.0
        '외피면적
        _SurfaceArea = 0.0
        '외부체적
        _V_e = 0.0
        _SurfaceAreaToVolumeRatio = 0.0
        '기하학적 특성값
        '길이
        _L_char = 0.0
        '너비
        _B_char = 0.0
    End Sub
    Public Sub Import(FileName As String)

    End Sub
    Public Sub Export(filename As String)

    End Sub
#End Region
    '
End Class
