Public Class Envelope

    '이름
    '외피 유형 => 일반 벽이든 창이 되게 한다. (상속으로 처리)
    '인접면 설정 => 외기든 다른 존이든 설정하게 한다.
    Private _InsideZone As Zone

    '방위
    '기울기
    '면적
    '개수
    '열관류율
    '흡수율
    '인접 존 선택
    Private _OutsideZone As Zone



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
End Class