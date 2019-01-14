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

Public Enum zSoilType
    점토
    모래또는자갈
    바위
End Enum

Public Enum zLocationIndex
    서울운영규정 = 0
    강릉
    강화
    거제
    고창
    고흥
    광주
    구미
    군산
    금산
    남원
End Enum

Public Enum UsageType
    주거용도
    단독사무실
    그룹사무실
    오픈플랜사무실
    회의실
    카운터
    소매점
    상점
    교실
    강당
    침실
    객실
    구내식당
    레스토랑
    비주거용주방
    주방용창고
    비주거용화장실
    그외거주가능실
    부속실거주불가
    복도
    저장및설비실
    서버실
    작업실1_서서하는힘든작업
    작업실2_일부서서하는다소힘든작업
    작업실3_조금서서하는약간힘든작업
    객석
    공연장로비
    공연장무대
    박람회및의회실
    전시장박물관
    도서관열람실
    도서관오픈통로
    도시관보관실및잡지
    체육관관람석제외
    주차장1사무실개인
    주차장2공공
    사우나15m2미만욕조
    피트니스운동실
    실험실
    난방공간과구별되는연구및조사실
    특별보호실
    일반보호실의복도
    병원및치료실
    회관저장공간
    대규모사무실사용자입력
    구내식당사용자입력
End Enum

'존 관련
Public Enum zUsageType
    주거용도
    단독사무실
    그룹사무실
    오픈플랜사무실
    회의실
    카운터
    소매점
    상점
    교실
    강당
    침실
    객실
    구내식당
    레스토랑
    비주거용주방
    주방용창고
    비주거용화장실
    그외거주가능실
    부속실거주불가
    복도
    저장및설비실
    서버실
    작업실1_서서하는힘든작업
    작업실2_일부서서하는다소힘든작업
    작업실3_조금서서하는약간힘든작업
    객석
    공연장로비
    공연장무대
    박람회및의회실
    전시장박물관
    도서관열람실
    도서관오픈통로
    도시관보관실및잡지
    체육관관람석제외
    주차장1사무실개인
    주차장2공공
    사우나15m2미만욕조
    피트니스운동실
    실험실
    난방공간과구별되는연구및조사실
    특별보호실
    일반보호실의복도
    병원및치료실
    회관저장공간
    대규모사무실사용자입력
    구내식당사용자입력
End Enum
Public Enum zConditioningMode
    없음
    냉방
    난방
    냉방난방
End Enum
Public Enum zHVACMode
    없음
    환기
    냉방
    난방
    냉방난방
End Enum

'감소운전방식
Public Enum zRunningMode
    감소운전
    운전정지
End Enum

'건물자동화계수
Public Enum zAutomationFactor
    D수동제어
    C건물자동제어
    B실별자동제어
    A최적자동제어
End Enum

'지중 접합 유형
Public Enum zGroundContactType
    없음
    지면맞닿음
    냉난방지하
    비냉난방지하실
End Enum