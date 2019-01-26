Public Class UsageProfile
#Region "Instance Variables"
    Private _Number As Integer
    Private _Name As String
    '
    Private _StartTime As Integer '01.사용시작시각
    Private _EndTime As Integer '02. 사용종료시각
    Private _t_wd_d As Integer '03. 하루사용시간
    Private _d_wd_a As Integer '04. 연간사용일수
    Private _t_Day As Integer '05. 연간낮사용시간
    Private _t_Night As Integer '06. 연간밤사용시간,	
    Private _t_v_op_d As Integer    '07. 냉동/냉방일일운전시간,	
    Private _d_op_a As Integer '08. 냉난방설비운영일수,	
    Private _t_h_op_d As Integer '09. 난방일일운전시간,	
    Private _E_m As Double  '10. 조도,
    Private _h_Ne As Double  '11. 작업면 높이,	
    Private _k_A As Double  '12. 작업영역 감소계수,	
    Private _C_A As Double  '13. 부재율,	
    Private _k As Double '14. 실 계수,
    Private _F_t As Double   '15. 조명운전 감소계수,	
    Private _k_VB As Double  '16. 수직면기준 조명계수,	
    Private _θ_i_h_soll As Double  '17. 난방설정온도,	θi_h_soll	
    Private _θ_i_c_soll As Double  '18. 냉방설정온도,	θi_c_soll
    Private _d_θ_i_NA As Double '19. 허용셋백온도,	Δθi_NA
    Private _θ_i_h_min As Double   '20. 난방최저온도,	θi_h_min
    Private _θ_i_c_max As Double  '21. 냉방최고온도,	θi_c_max
    Private _HumidityRequirements As Boolean  '22. 습도민감도(Humidity requirements)[with tolerance or not],	-
    Private _V_A As Double  '23. 거주환경기준 최소외기도입량,V_A
    Private _V_A_bldg As Double   '24. 건물기준 최소외기 도입량,	V_A_bldg
    Private _c_HVAC As Double  '25. 상대부재율_Relative Abwesenheit RLT,	c_HVAC
    Private _F_HVAC As Double  '26. 냉동/냉방 부분운전계수,	F_HVAC
    Private _q_I_p As Double  '27. 인체내부발열,	qI_p
    Private _q_I_fac As Double '28. 기기내부발열, qI_fac
#End Region

#Region "Constructors"
    '문자열로 사용프로필 생성
    Public Sub New(txt)
        InternalParse(txt)
    End Sub
#End Region

#Region "Properties"
    Public ReadOnly Property Number As Integer
        Get
            Return _Number
        End Get
    End Property

    Public ReadOnly Property Name As String
        Get
            Return _Name
        End Get
    End Property

    Public ReadOnly Property StartTime As Integer '01.사용시작시각
        Get
            Return _StartTime
        End Get
    End Property

    Public ReadOnly Property EndTime As Integer '02. 사용종료시각
        Get
            Return _EndTime
        End Get
    End Property

    Public ReadOnly Property t_wd_d As Integer '03. 하루사용시간
        Get
            Return _t_wd_d
        End Get
    End Property

    Public ReadOnly Property d_wd_a As Integer '04. 연간사용일수
        Get
            Return _d_wd_a
        End Get
    End Property

    Public ReadOnly Property t_Day As Integer '05. 연간낮사용시간
        Get
            Return _t_Day
        End Get
    End Property

    Public ReadOnly Property t_Night As Integer '06. 연간밤사용시간,	
        Get
            Return _t_Night
        End Get
    End Property
    Public ReadOnly Property t_v_op_d As Integer    '07. 냉동/냉방일일운전시간,	
        Get
            Return _t_v_op_d
        End Get
    End Property

    Public ReadOnly Property d_op_a As Integer '08. 냉난방설비운영일수,	
        Get
            Return _d_op_a
        End Get
    End Property

    Public ReadOnly Property t_h_op_d As Integer '09. 난방일일운전시간,	
        Get
            Return _t_h_op_d
        End Get
    End Property

    Public ReadOnly Property E_m As Double  '10. 조도,
        Get
            Return _E_m
        End Get
    End Property

    Public ReadOnly Property h_Ne As Double  '11. 작업면 높이,	
        Get
            Return _h_Ne
        End Get
    End Property

    Public ReadOnly Property k_A As Double  '12. 작업영역 감소계수,	
        Get
            Return _k_A
        End Get
    End Property

    Public ReadOnly Property C_A As Double  '13. 부재율,	
        Get
            Return _C_A
        End Get
    End Property

    Public ReadOnly Property k As Double '14. 실 계수,
        Get
            Return _k
        End Get
    End Property

    Public ReadOnly Property F_t As Double   '15. 조명운전 감소계수,	
        Get
            Return _F_t
        End Get
    End Property

    Public ReadOnly Property k_VB As Double  '16. 수직면기준 조명계수,	
        Get
            Return _k_VB
        End Get
    End Property

    Public ReadOnly Property θ_i_h_soll As Double  '17. 난방설정온도,	θi_h_soll	
        Get
            Return _θ_i_h_soll
        End Get
    End Property

    Public ReadOnly Property θ_i_c_soll As Double  '18. 냉방설정온도,	θi_c_soll
        Get
            Return _θ_i_c_soll
        End Get
    End Property

    Public ReadOnly Property d_θ_i_NA As Double '19. 허용셋백온도,	Δθi_NA
        Get
            Return _d_θ_i_NA
        End Get
    End Property

    Public ReadOnly Property θ_i_h_min As Double   '20. 난방최저온도,	θi_h_min
        Get
            Return _θ_i_h_min
        End Get
    End Property

    Public ReadOnly Property θ_i_c_max As Double  '21. 냉방최고온도,	θi_c_max
        Get
            Return θ_i_c_max
        End Get
    End Property
    Public ReadOnly Property HumidityRequirements As Boolean  '22'22. 습도민감도,	-
        Get
            Return _HumidityRequirements
        End Get
    End Property
    Public ReadOnly Property V_A As Double  '23. 거주환경기준 최소외기도입량,V_A
        Get
            Return _V_A
        End Get
    End Property

    Public ReadOnly Property V_A_bldg As Double   '24. 건물기준 최소외기 도입량,	V_A_bldg
        Get
            Return _V_A_bldg
        End Get
    End Property

    Public ReadOnly Property c_HVAC As Double  '25. 상대부재율_Relative Abwesenheit RLT,	c_HVAC
        Get
            Return _c_HVAC
        End Get
    End Property

    Public ReadOnly Property F_HVAC As Double  '26. 냉동/냉방 부분운전계수,	F_HVAC
        Get
            Return _F_HVAC
        End Get
    End Property

    Public ReadOnly Property q_I_p As Double  '27. 인체내부발열,	qI_p
        Get
            Return _q_I_p
        End Get
    End Property

    Public ReadOnly Property q_I_fac As Double '28. 기기내부발열, qI_fac
        Get
            Return _q_I_fac
        End Get
    End Property
#End Region
#Region "Methods"

    ''' <summary>
    ''' 문자열로 된 기상자료를 분석하여 내부 변수에 할당한다.
    ''' </summary>
    ''' <param name="txt">기상자료[문자열]</param>
    Protected Overridable Sub InternalParse(ByVal txt As String)
        '
        '문자열을 쉼표로 구분하여 단어별로 배열을 만든다.
        Dim Words As String() = txt.Split(New Char() {","c})
        '
        '기상자료 번호
        _Number = Integer.Parse(Words(0)) '= cInt("12")
        '지역 이름
        _Name = Words(1)
        'Comment = words(2)
        _StartTime = Integer.Parse(Words(3)) '01.사용시작시각
        _EndTime = Integer.Parse(Words(4)) '02. 사용종료시각
        _t_wd_d = Integer.Parse(Words(5)) '03. 하루사용시간
        _d_wd_a = Integer.Parse(Words(6)) '04. 연간사용일수
        _t_Day = Integer.Parse(Words(7)) '05. 연간낮사용시간
        _t_Night = Integer.Parse(Words(8)) '06. 연간밤사용시간,	
        _t_v_op_d = Integer.Parse(Words(9))   '07. 냉동/냉방일일운전시간,	
        _d_op_a = Integer.Parse(Words(10)) '08. 냉난방설비운영일수,	
        _t_h_op_d = Integer.Parse(Words(11)) '09. 난방일일운전시간,	
        _E_m = Integer.Parse(Words(12))  '10. 조도,
        _h_Ne = Double.Parse(Words(13))   '11. 작업면 높이,	
        _k_A = Double.Parse(Words(14))  '12. 작업영역 감소계수,	
        _C_A = Double.Parse(Words(15))  '13. 부재율,	
        _k = Double.Parse(Words(16)) '14. 실 계수,
        _F_t = Double.Parse(Words(17))   '15. 조명운전 감소계수,	
        _k_VB = Double.Parse(Words(18))  '16. 수직면기준 조명계수,	
        _θ_i_h_soll = Double.Parse(Words(19))  '17. 난방설정온도,	θi_h_soll	
        _θ_i_c_soll = Double.Parse(Words(20))  '18. 냉방설정온도,	θi_c_soll
        _d_θ_i_NA = Double.Parse(Words(21)) '19. 허용셋백온도,	Δθi_NA
        _θ_i_h_min = Double.Parse(Words(22))   '20. 난방최저온도,	θi_h_min
        _θ_i_c_max = Double.Parse(Words(23))  '21. 냉방최고온도,	θi_c_max
        '_HumidityRequirements = Boolean.Parse(Words(24))  '22. 습도민감도(Humidity requirements)[with tolerance or not],	-
        '엑셀에 있는 문자열 : m. T., o. T., -
        '현재 m. T.을 True로, 나머지는 False로 처리
        If Words(24) = "True" Then
            _HumidityRequirements = True
        Else
            _HumidityRequirements = False
        End If
        _V_A = Double.Parse(Words(25))  '23. 거주환경기준 최소외기도입량,V_A
        _V_A_bldg = Double.Parse(Words(26))   '24. 건물기준 최소외기 도입량,	V_A_bldg
        _c_HVAC = Double.Parse(Words(27))  '25. 상대부재율_Relative Abwesenheit RLT,	c_HVAC
        _F_HVAC = Double.Parse(Words(28)) '26. 냉동/냉방 부분운전계수,	F_HVAC
        _q_I_p = Double.Parse(Words(29))  '27. 인체내부발열,	qI_p
        _q_I_fac = Double.Parse(Words(30)) '28. 기기내부발열, qI_fac
        '
    End Sub
#End Region
    '	

#Region "Shared Methods"
    Public Overloads Shared Function Parse(ByVal txt As String) As UsageProfile
        Return New UsageProfile(txt)
    End Function
#End Region
End Class



Public Class UsageProfiles
#Region "Instance Variables"
    Private _UsageList As Generic.List(Of UsageProfile)
    Private _UsageStrings As Generic.List(Of String)
#End Region
#Region "Constructors"
    Sub New()
        _UsageList = New Generic.List(Of UsageProfile)
        _UsageStrings = New Generic.List(Of String)
        '
        Call Initialize()
    End Sub
#End Region
#Region "Methods"
    Private Sub Initialize()
        With _UsageStrings
            '형식 :.Add("")
            '.Add("재실시간/이용시간,조명,실내환경,내부발열")
            '.Add("01. 사용시작시간,	02. 사용종료시간,	03. 하루사용시간,	04. 연간사용일수,	05. 연간낮사용시간,	06. 연간밤사용시간,	07. 냉동/냉방일일운전시간,	08. 냉난방설비운영일수,	09. 난방일일운전시간,	10. 조도,	11. 작업면 높이,	12. 작업영역 감소계수,	13. 부재율,	14. 실 계수,	15. 조명운전 감소계수,	16. 수직면기준 조명계수,	17. 난방설정온도,	18. 냉방설정온도,	19. 허용셋백온도,	20. 난방최저온도,	21. 냉방최고온도,	22. 습도민감도,	23. 거주환경기준 최소외기도입량,	24. 건물기준 최소외기 도입량,	"25. 상대부재율_Relative Abwesenheit RLT,	26. 냉동/냉방 부분운전계수,	27. 인체내부발열,	28. 기기내부발열")
            '.Add("기호,-,	-,	t_wd_d,	d_wd_a,	t_Day,	t_Night,	t_v_op_d,	d_op_a,	t_h_op_d,	E_m,	h_Ne,	k_A,	C_A,	k,	F_t,	kVB,	θi_h_soll,	θi_c_soll,	Δθi_NA,	θi_h_min,	θi_c_max,	–,	VA,	VA_,bldg,	cHVAC,	FHVAC,	qI_p,	qI_fac")
            '.Add("단위,Time, Time, h/d, d/a, h/a, h/a, h/d, d/a, h/d, lx, m, -, -, -, -, -, °C, °C, K, °C, °C, -, m³/(h ∙ m²), m³/(h ∙ m²), -, -, Wh/(m2 · d), Wh/(m2 · d)")
            .Add("00,주거 용도,00. 주거 용도,0,24,24,365,0,0,24,365,17,0,0,0,0,0,0,0,20,26,4,20,26,0,0,0,0,0,50,0")
            .Add("01,단독 사무실,01. 단독 사무실,7,18,11,250,2543,207,13,250,13,500,0.8,0.84,0.3,0.9,0.7,1,21,24,4,20,26,True,4,2.5,0.3,0.7,30,43")
            .Add("02,그룹 사무실,02. 그룹 사무실,7,18,11,250,2543,207,13,250,13,500,0.8,0.92,0.3,1.25,0.7,1,21,24,4,20,26,True,4,2.5,0.3,0.7,30,43")
            .Add("03,오픈 플랜 사무실,03. 오픈 플랜 사무실,7,18,11,250,2543,207,13,250,13,500,0.8,0.93,0,2.5,1,1,21,24,4,20,26,True,6,2.5,0.2,1,42,60")
            .Add("04,회의실,04. 회의실,7,18,11,250,2543,207,13,250,13,500,0.8,0.93,0.5,1.25,1,1,21,24,4,20,26,True,15,2.5,0.5,0.5,93,8")
            .Add("05,카운터,05. 카운터,7,18,11,250,2543,207,13,250,13,200,0.8,0.87,0,1.5,1,1,21,24,4,20,26,True,2,1.25,0.5,1,35,24")
            .Add("06,소매점,06. 소매점,8,20,12,300,3009,591,14,300,14,300,0.8,0.93,0,2.5,1,1.5,21,24,4,20,26,True,4,2.5,0.5,1,84,24")
            .Add("07,상점,07. 상점,8,20,12,300,3009,591,14,300,14,300,0.8,0.93,0,2.5,1,1,21,24,4,20,26,True,4,2.5,0.5,1,84,-170")
            .Add("08,교실,08. 교실,8,15,7,200,1400,0,9,200,9,300,0.8,0.97,0.25,2,0.9,1,21,24,4,20,26,True,10,2.5,0.25,0.9,100,20")
            .Add("09,강당,09. 강당,8,18,10,150,1408,92,12,150,12,500,0.8,0.92,0.25,2.5,0.7,1,21,24,4,20,26,True,30,2.5,0.5,0.6,420,24")
            .Add("10,침실,10. 침실,0,24,24,365,4407,4353,24,365,24,300,0.8,1,0,1.5,0.5,1,22,24,0,20,26,0,5,2.5,0,0.8,120,24")
            .Add("11,객실,11. 객실,21,8,11,365,743,3272,24,365,24,200,0.8,1,0.6,1.25,0.3,1,21,24,0,20,26,True,3,2,0.5,0.5,70,44")
            .Add("12,구내식당,12. 구내식당,8,15,7,250,1750,0,9,250,9,200,0.8,0.97,0,2.5,1,1,21,24,4,20,26,True,18,2.5,0.7,1,175,10")
            .Add("13,레스토랑,13. 레스토랑,10,0,14,300,2411,1789,16,300,16,200,0.8,1,0,2.5,1,1,21,24,4,20,26,True,18,2.5,0.6,0.7,233,14")
            .Add("14,비주거용 주방,14. 비주거용 주방,10,23,13,300,2411,1489,15,300,15,500,0.8,0.96,0,1.5,1,1,21,24,4,20,26,True,90,0,0,0,56,1800")
            .Add("15,주방용 창고,15. 주방용 창고,10,23,13,300,2411,1489,15,300,15,300,0.8,1,0.5,1.5,1,1,21,24,4,20,26,True,15,0,0,0,56,180")
            .Add("16,비주거용 화장실,16. 비주거용 화장실,7,18,11,250,2543,207,13,250,13,200,0.8,1,0.9,0.8,1,1,21,24,4,20,26,0,15,5,0.7,1,0,0")
            .Add("17,그 외 거주가능실,17. 그 외 거주가능실,7,18,11,250,2543,207,13,250,13,300,0.8,0.93,0.5,1.25,1,1,21,24,4,20,26,True,7,2.5,0.5,0.8,93,0")
            .Add("18,부속실 (거주불가),18. 부속실 (거주불가),7,18,11,250,2543,207,13,250,13,100,0.8,1,0.9,1.5,1,1,21,24,4,20,26,0,0.15,0,0,0,0,0")
            .Add("19,복도,19. 복도,7,18,11,250,2543,207,13,250,13,100,0.2,1,0.8,0.8,1,1,21,24,4,20,26,0,0,0,0,0,0,0")
            .Add("20,저장 및 설비실,20. 저장 및 설비실,7,18,11,250,2543,207,13,250,13,100,0.8,1,0.98,1.5,1,2,21,24,4,20,26,0,0.15,0,0,0,0,0")
            .Add("21,서버실,21. 서버실,0,24,24,365,4407,4353,24,365,24,500,0.8,0.96,0.5,1.5,0.5,1,21,24,4,20,26,0,1.3,0,0,0,14,1800")
            .Add("22,작업실1_서서하는 힘든 작업,22. 작업실1_서서하는 힘든 작업,7,16,9,230,2018,52,10,230,10,300,0.8,0.85,0.1,2.5,0.9,1,15,28,4,15,30,0,3.5,0,0,0,48,280")
            .Add("23,작업실2_일부 서서하는 다소 힘든 작업,23. 작업실2_일부 서서하는 다소 힘든 작업,7,16,9,230,2018,52,10,230,10,400,0.8,0.85,0.1,2.5,0.9,1,17,26,4,15,28,0,2.5,0,0,0,40,280")
            .Add("24,작업실3_조금 서서하는 약간 힘든 작업,24. 작업실3_조금 서서하는 약간 힘든 작업,7,16,9,230,2018,52,10,230,10,500,0.8,0.85,0.1,2.5,0.9,1,20,24,4,18,26,0,1.5,0,0,0,32,280")
            .Add("25,객석,25. 객석,19,23,4,250,59,941,6,250,6,200,0.8,0.97,0,4,1,1,21,24,4,20,26,True,40,5,0.7,1,187,0")
            .Add("26,공연장 로비,26. 공연장 로비,19,23,4,250,59,941,6,250,6,300,0,1,0.5,4,1,1,21,24,4,20,26,True,25,5,0.5,1,88,0")
            .Add("27,공연장 무대,27. 공연장 무대,13,23,10,250,1259,1241,12,250,12,1000,0,0.9,0,2.5,0.6,1,21,24,4,20,26,True,0.3,0,0,0,0,0")
            .Add("28,박람회 및 의회실,28. 박람회 및 의회실,9,18,9,150,1258,92,11,150,11,300,0.8,0.93,0.5,5,1,1,21,24,4,20,26,True,7,2.5,0.7,1,140,12")
            .Add("29,전시장 / 박물관,29. 전시장 / 박물관,10,18,8,250,1846,154,24,365,24,200,0.8,0.88,0,2,1,1,21,24,4,20,26,o. T.,2,2,0.5,1,28,0")
            .Add("30,도서관 열람실,30. 도서관 열람실,8,20,12,300,3009,591,14,300,14,500,0.8,0.88,0,1.5,1,1,21,24,4,20,26,True,8,2.5,0.5,1,168,0")
            .Add("31,도서관 오픈 통로 ,31. 도서관 오픈 통로 ,8,20,12,300,3009,591,14,300,14,200,0.8,1,0,1.5,1,1.7,21,24,4,20,26,True,2,2,0,0,42,0")
            .Add("32,도시관 보관실 및 잡지,32. 도시관 보관실 및 잡지,8,20,12,300,3009,591,14,300,14,100,0.8,1,0.9,1.5,1,2,21,24,4,20,26,True,3,2,0.5,1,0,0")
            .Add("33,체육관 (관람석제외),33. 체육관 (관람석제외),8,23,15,250,2509,1241,17,250,17,300,1,1,0.3,2,1,1,19,24,4,20,26,0,3,1.25,0.5,0.9,63,0")
            .Add("34,주차장1 (사무실 개인),34. 주차장1 (사무실 개인),7,18,11,250,2543,207,13,250,0,75,0.2,1,0.95,2,1,1,21,24,4,20,26,0,8,0,0,0,0,0")
            .Add("35,주차장2 (공공),35. 주차장2 (공공),9,24,15,365,3298,2177,17,365,0,75,0.2,1,0.8,4,1,1,21,24,4,20,26,0,8,0,0,0,0,0")
            .Add("36,사우나 (15m2미만 욕조),36. 사우나 (15m2미만 욕조),10,22,12,365,2933,1447,14,365,14,200,0,1,0,1,1,1,24,0,4,23,0,True,15,0,0,0,58,500")
            .Add("37,피트니스 / 운동실,37. 피트니스 / 운동실,8,23,15,365,3663,1812,17,365,17,300,0,1,0,2,1,1,20,24,4,18,26,True,12,2.5,0.5,0.9,264,24")
            .Add("38,실험실,38. 실험실,7,18,11,250,2543,207,24,250,13,500,1,0.92,0.3,1.25,1,1,22,24,4,20,26,True,25,0,0,0,39,108")
            .Add("39,난방공간과 구별되는 연구 및 조사실,39. 난방공간과 구별되는 연구 및 조사실,7,18,11,250,2543,207,13,250,13,500,0.8,1,0,1.2,1,1,22,24,0,20,26,True,10,2.5,0.3,0.7,82,35")
            .Add("40,특별보호실,40. 특별보호실,0,24,24,365,4407,4353,24,365,24,300,0.8,1,0,1.2,0.8,1,24,24,0,22,26,True,30,0,0,0,112,228")
            .Add("41,일반보호실의 복도,41. 일반보호실의 복도,0,24,24,365,4407,4353,24,365,24,125,0.2,1,0.8,1,1,1,22,24,0,20,26,0,10,0,0,0,0,0")
            .Add("42,병원 및 치료실,42. 병원 및 치료실,8,18,10,250,2346,154,12,250,12,500,0.8,1,0,1.2,1,1,22,24,4,20,26,True,10,2.5,0.3,0.7,53,25")
            .Add("43,회관 / 저장공간, 43. 회관 / 저장공간,0,24,24,365,4407,4353,24,365,24,150,0,1,0.6,2.4,0.4,1.8,12,26,0,12,28,0,1,0,0,0,0,0")
            .Add("45,대규모사무실 (사용자입력),45. 대규모사무실 (사용자입력),9,18,9,250,2543,207,11,250,11,500,0.8,0.93,0,2.5,1,1,20,26,4,20,26,True,6,2.5,0.2,1,55.8,126") '사용자입력부분
            .Add("46,구내식당 (사용자입력),46. 구내식당 (사용자입력),8,15,7,250,1750,0,9,250,9,200,0.8,0.97,0,2.5,1,1,20,26,4,20,26,True,18,2.5,0.7,1,177,10") '사용자입력부분
        End With
        '
        For Each txt As String In _UsageStrings
            _UsageList.Add(UsageProfile.Parse(txt))
        Next
    End Sub
#End Region

End Class