'기상 자료 관련
Public Enum zOrientationType
    수평면 = 0
    북 = 1
    북동 = 2
    동 = 3
    남동 = 4
    남 = 5
    남서 = 6
    서 = 7
    북서 = 8
End Enum

'건물 관련
Public Enum zBuildingType
    주거단독 = 0
    주거다세대 = 1
    비주거 = 2
End Enum

'토양 종류
Public Enum zSoilType
    점토 = 0
    모래또는자갈 = 1
    바위 = 2
End Enum

'위치
Public Enum zLocationIndex
    서울운영규정 = 0
    강릉 = 1
    강화 = 2
    거제 = 3
    거창 = 4
    고흥 = 5
    광주 = 6
    구미 = 7
    군산 = 8
    금산 = 9
    남원 = 10
    남해 = 11
    대구 = 12
    대전 = 13
    동두천 = 14
    동해 = 15
    목포 = 16
    문경 = 17
    밀양 = 18
    보령 = 19
    보은 = 20
    봉화 = 21
    부산 = 22
    부안 = 23
    부여 = 24
    산청 = 25
    서산 = 26
    서울 = 27
    속초 = 28
    수원 = 29
    안동 = 30
    양평 = 31
    여수 = 32
    영덕 = 33
    영동 = 34
    영월 = 35
    영주 = 36
    영천 = 37
    완도 = 38
    울산 = 39
    울진 = 40
    원주 = 41
    의성 = 42
    이천 = 43
    인제 = 44
    인천 = 45
    임실 = 46
    장수 = 47
    장흥 = 48
    전주 = 49
    정읍 = 50
    제주 = 51
    제천 = 52
    진주 = 53
    창원 = 54
    천안 = 55
    철원 = 56
    청주 = 57
    춘천 = 58
    충주 = 59
    태백 = 60
    통영 = 61
    평창 = 62
    포항 = 63
    합천 = 64
    해남 = 65
    홍천 = 66
End Enum


'존 관련
Public Enum zUsageType
    주거용도 = 0 '주거용도
    단독사무실 = 1
    그룹사무실 = 2
    오픈플랜사무실 = 3
    회의실 = 4
    카운터 = 5
    소매점 = 6
    상점 = 7
    교실 = 8
    강당 = 9
    침실 = 10
    객실 = 11
    구내식당 = 12
    레스토랑 = 13
    비주거용주방 = 14
    주방용창고 = 15
    비주거용화장실 = 16
    그외거주가능실 = 17
    부속실거주불가 = 18
    복도 = 19
    저장및설비실 = 20
    서버실 = 21
    작업실1_서서하는힘든작업 = 22
    작업실2_일부서서하는다소힘든작업 = 23
    작업실3_조금서서하는약간힘든작업 = 24
    객석 = 25
    공연장로비 = 26
    공연장무대 = 27
    박람회및의회실 = 28
    전시장박물관 = 29
    도서관열람실 = 30
    도서관오픈통로 = 31
    도시관보관실및잡지 = 32
    체육관관람석제외 = 33
    주차장1사무실개인 = 34
    주차장2공공 = 35
    사우나15m2미만욕조 = 36
    피트니스운동실 = 37
    실험실 = 38
    난방공간과구별되는연구및조사실 = 39
    특별보호실 = 40
    일반보호실의복도 = 41
    병원및치료실 = 42
    회관저장공간 = 43
    대규모사무실사용자입력 = 44 '엑셀 파일 넘버 45
    구내식당사용자입력 = 45 '엑셀 파일 넘버 46
End Enum

'냉난방 여부
Public Enum zConditioningMode
    없음 = 0
    냉방 = 1
    난방 = 2
    냉방난방 = 3
End Enum

'HVAC 시스템
Public Enum zHVACMode
    없음 = 0
    환기 = 1
    냉방 = 2
    난방 = 3
    냉방난방 = 4
End Enum

'감소운전방식
Public Enum zRunningMode
    감소운전 = 0
    운전정지 = 1
End Enum

'건물자동화계수
Public Enum zAutomationFactor
    D수동제어 = 0
    C건물자동제어 = 1
    B실별자동제어 = 2
    A최적자동제어 = 3
End Enum

'지중 접합 유형
Public Enum zGroundContactType
    없음 = 0
    지면맞닿음 = 1
    냉난방지하 = 2
    비냉난방지하실 = 3
End Enum


'급수량기준
Public Enum zWaterSupplyStandard
    용도프로필바닥면적기준 = 0
    사용자입력데이터 = 1
End Enum

'난방 공급방식
Public Enum zHeatingType
    라디에이터 = 0
    바닥난방_열 = 1
    바닥난방_전기 = 2
    전기 = 3
End Enum

'열원방식
Public Enum zHeatSourceType
    보일러 = 0
    지역난방 = 1
    전기 = 2
End Enum

'열교가산치
Public Enum zdU_WB
    일반적인경우 = 0
    내단열인경우 = 1
    DIN4108_2_Annex_A검토 = 2
    DIN4108_2_Annex_B검토 = 3
    직접입력 = 4
End Enum

'DIN 4108에 따른 열교가산치
Public Enum zdU_WB_DIN4108
    모든부위검토 = 0
    일부부위검토 = 1
End Enum

'냉난방공간
Public Enum zHeatingCooling
    냉난방공간 = 0
    비냉난방공간 = 1
End Enum

'공조공간
Public Enum zHVACYN
    공조공간 = 0
    비공조공간 = 1
End Enum

'환기장치
Public Enum zVentilationYN
    환기장치있음 = 0
    환기장치없음 = 1
End Enum

'입력여부
Public Enum zInputYN
    입력완료 = 0
    정보누락 = 1
    입력오류 = 2
End Enum

'실별 컨디션 설정

'외피유형
Public Enum zEnvelopeType
    외피 = 0
    창문 = 1
End Enum

'인접면 설정
Public Enum zAdjZone
    외기 = 0
    비냉난방공간 = 1
    인접존 = 2
    지중 = 3
End Enum

'기울기
Public Enum zInclination
    _90도 = 0
    _60도 = 1
    _45도 = 2
    _30도 = 3
    _0도 = 4
End Enum

'유리종류
'Public Enum zGlassType

'End Enum

'차양설치위치
Public Enum zShadeLocation
    없음 = 0
    외부 = 1
    내부 = 2
End Enum

'외부차양장치
Public Enum zExternalShadeType
    베네치안블라인드_10도슬렛 = 0
    베네치안블라인드_45도슬렛 = 1
    수직천막 = 2
    롤블라인드_완전닫힘 = 3
    롤블라인드_일부닫힘 = 4 '3/4닫힘
End Enum

'내부차양장치
Public Enum zIndoorShadeType
    베네치안블라인드_10도슬렛 = 0
    베네치안블라인드_45도슬렛 = 1
    롤블라인드 = 2
    필름 = 3
End Enum

'차양장치 색상
Public Enum zShadeColor
    백색 = 0
    회색 = 1
End Enum

'제어방식
Public Enum zControlType
    제어안함 = 0
    수동_타이머제어 = 1
    일사자동제어 = 2
End Enum

'수평 방해물 : "외피" 탭에서는 <수직 장애물>로 표현
Public Enum zF_h
    _0 = 0
    _10 = 1
    _20 = 2
    _30 = 3
    _40 = 4
End Enum

'수평 오버행 : "외피" 탭에서는 <수직 오버행>으로 표현
Public Enum zF_o
    _0 = 0
    _30 = 1
    _45 = 2
    _60 = 3
End Enum

'수직 핀 : "외피" 탭에서는 <수평 핀>으로 표현
Public Enum zF_f
    _0 = 0
    _30 = 1
    _45 = 2
    _60 = 3
End Enum

'환기방식
Public Enum zVentType
    전체자연환기 = 0
    전체기계환기 = 1
    부분기계환기 = 2
    에어컨디셔닝 = 3
End Enum

'외기도입여부
Public Enum zOA_YN
    실외로열리는창있음 = 0
    창문없음 = 1
    외기와맞닿은면없음 = 2
    자연환기안함 = 3
End Enum

'기밀테스트
Public Enum zn50Type
    카테고리1 = 0
    카테고리2 = 1
    카테고리3 = 2
    카테고리4 = 3
    측정값사용 = 4
End Enum

'열회수 유형
Public Enum zHeatRecoveryType
    열회수되지않음 = 0
    현열교환 = 1
    전열교환 = 2
End Enum

'존간 환기
Public Enum zVentBetweenZone
    급기 = 0
    흡기 = 1
End Enum

'통기구 여부
Public Enum zVentYN
    있음 = 0
    없음 = 1
End Enum

'그래프 확인
Public Enum zGraphCheck
    전체 = 0
End Enum

'입력방법
Public Enum zInputStandard
    이용용도기준 = 0
    바닥면적기준 = 1
End Enum

