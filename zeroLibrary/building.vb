'<변수 명명 규칙>
' 1)DIN 규정에 나온 변수명을 그대로 사용한다.
' 2)변수명의 마침표(.)는 밑줄(_)로 한다.
' 3)내부변수의 시작은 밑줄(_)로 한다.
'</변수 명명 규칙>

''' <summary>
''' 건물을 정보를 갖는 클래스
''' </summary>
Public Class Building
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
    'Private _SurfaceAreaToVolumeRatio As Double
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
    ''' <summary>
    ''' Default constructor
    ''' </summary>
    Sub New()
        _zones = New Zones()
        '
        SetDefaultValues()
    End Sub

    'copy constructor
    Sub New(otherBuilding As Building)
        Me.New()
        '
        '<복제 과정>
        Me._SoilType = otherBuilding._SoilType
        '</복제 과정>
    End Sub
    Sub New(txt As String)
        InternalParse(txt)
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
        Get
            Return _SoilType
        End Get
        Set(value As zSoilType)
            _SoilType = value
        End Set
    End Property
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    Public Property Name As String
        Get
            Return _Name
        End Get
        Set(value As String)
            _Name = value
        End Set
    End Property

    ' 고도 차이
    ''' <summary>
    ''' 건물과 기상청의 고도 차이를 입력하거나 출력하는 속성.
    ''' 고도차이 = 건물의 고도 - 기상청의 고도
    ''' </summary>
    ''' <returns>고도 차이[m]</returns>
    Public Property Altitude As Double
        Get
            Return _Altitude
        End Get
        Set(Value As Double)
            _Altitude = Value
        End Set
    End Property

    '건물 주소
    ''' <summary>
    ''' 건물 주소를 입력하거나 출력하는 속성.
    ''' </summary>
    ''' <returns>건물 주소</returns>
    Public Property Address As String
        Get
            Return _Address
        End Get
        Set(sValue As String)
            _Address = sValue
        End Set
    End Property

    '층수
    ''' <summary>
    ''' 층수를 입력하거나 출력하는 속성.
    ''' </summary>
    ''' <returns>층수</returns>
    Public Property n_G As Integer
        Get
            Return _n_G
        End Get
        Set(iValue As Integer)
            _n_G = iValue
        End Set
    End Property

    '연면적
    ''' <summary>
    ''' 건물의 연면적을 입력하거나 출력하는 속성.
    ''' </summary>
    ''' <returns>연면적[m^2]</returns>
    Public Property TotalFloorArea As Double
        Get
            Return _TotalFloorArea
        End Get
        Set(Value As Double)
            _TotalFloorArea = Value
        End Set
    End Property
    '
    '  
    ''' <summary>
    ''' 건물에 등록된 모든 존들의 순바닥면적을 반환한다.
    ''' 순바닥면적 net floor area = 모든 존의 바닥면적의 합
    ''' </summary>
    ''' <returns>순바닥면적[m^2]</returns>
    Public ReadOnly Property A_NF As Double
        Get
            Return _zones.A_NF()
        End Get
    End Property

    '
    ''' <summary>
    ''' 외피면적을 입력하거나 출력하는 속성
    ''' </summary>
    ''' <returns>외피면적</returns>
    Public Property SurfaceArea As Double
        Get
            Return _SurfaceArea
        End Get
        Set(value As Double)
            _SurfaceArea = value
        End Set
    End Property

    '외부체적
    Public Property V_e As Double
        Get
            Return _V_e
        End Get
        Set(value As Double)
            _V_e = value
        End Set
    End Property

    '외피면적체적비
    Public ReadOnly Property SurfaceAreaToVolumeRatio As Double
        Get
            Dim _SurfaceAreaToVolumeRatio As Double
            If _V_e <> 0 Then
                _SurfaceAreaToVolumeRatio = _SurfaceArea / _V_e
            Else
                _SurfaceAreaToVolumeRatio = 0
            End If
            Return _SurfaceAreaToVolumeRatio
        End Get
    End Property

    '기하학적 특성값
    '길이
    Public Property L_char As Double
        Get
            Return _L_char
        End Get
        Set(value As Double)
            _L_char = value
        End Set
    End Property

    '너비
    Public Property B_char As Double
        Get
            Return _B_char
        End Get
        Set(value As Double)
            _B_char = value
        End Set
    End Property
    '
#End Region

#Region "Methods"
    Public Sub SetDefaultValues()
        _BuildingType = zBuildingType.비주거
        _Location = zLocationIndex.서울운영규정
        _SoilType = zSoilType.모래또는자갈
        '
        _Name = "New Building"
        '고도차이
        _Altitude = 0.0 ' 서울고도 - 서울기상청고도
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
        '_SurfaceAreaToVolumeRatio = 0.0
        '기하학적 특성값
        '길이
        _L_char = 0.0
        '너비
        _B_char = 0.0
    End Sub

    Public Sub Import(FileName As String)
        Me.SurfaceArea = 100
        Dim ttt As Double = Me.SurfaceArea
        ttt = Me.SurfaceAreaToVolumeRatio
    End Sub
    Public Sub Export(filename As String)

    End Sub

    Protected Overridable Sub InternalParse(ByVal txt As String)
        Dim words As String() = txt.Split(New Char() {","c})
        '
        '
    End Sub
#End Region
    '
#Region "Shared Methods"
    Public Overloads Shared Function Parse(ByVal txt As String) As Building
        Return New Building(txt)
    End Function
#End Region
    '
End Class
