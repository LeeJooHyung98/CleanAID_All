Attribute VB_Name = "ProgramVersion"
Option Explicit

Public strProgram_Version   As String
Public strProgram_LastEdit  As String


' SP_01011_B_00 수정

Public Function SetProgramVersion()
    Dim MyVersion As String
    
    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.34
    strProgram_LastEdit = "2017.06.22"
'   장부관리 내용 조회 기능 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.33
'    strProgram_LastEdit = "2017.03.29"
'   요일 할인 관리에서 삭제 오류로 추정 되는 부분 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.31-32
'    strProgram_LastEdit = "2017.03.29"
'   업데이트 경로 변경
'   "www.clean-aid.co.kr:8090/cleanaid" => "www.insoftnet.com/upgrade/cleanaid"
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.30
'    strProgram_LastEdit = "2017.03.27"
'   환불 현황 지사별 전체 기능 조회 추가
'   환불 현황 인쇄 기능 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.29
'    strProgram_LastEdit = "2017.02.28"
'   CS사고품 조회관련 화면 기능 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.28
'    strProgram_LastEdit = "2017.02.23"
'   CS사고품 조회관련 화면 기능 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.27
'    strProgram_LastEdit = "2017.01.31"
'   가맹점 의류분류 등록에서 수신일자, 생성일자 기능 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.26
'    strProgram_LastEdit = "2017.01.31"
'   가맹점 의류분류 등록에서 마진을 소수점 2자리까지 등록 가능하도록 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.25
'    strProgram_LastEdit = "2017.01.19"
'  사고품 내역 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.24
'    strProgram_LastEdit = "2017.01.17"
'  사고품 내역 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.23
'    strProgram_LastEdit = "2016.12.13"
'  사고품 내역 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.22
'    strProgram_LastEdit = "2016.12.05"
'  사고품 내역 수정
'  정산 일자 수정이 되어도 바로 검색이 되지 않도록 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.21
'    strProgram_LastEdit = "2016.06.28"
'  가맹점 고객 미출고 현황 기능 추가 (P_03018)
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.20
'    strProgram_LastEdit = "2016.06.14"
'  판매취소 내용 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.19
'    strProgram_LastEdit = "2016.05.23"
'  일일 판매 집계(가맹점) 화면에서 미수금 수금 입금액 표시
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.18
'    strProgram_LastEdit = "2016.05.17"
'  반품현황 화면 수정

'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.17
'    strProgram_LastEdit = "2016.05.10"
'  사고 유형에 비율 추가
'  품목별 접수현황 내용 수정
    
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.16
'    strProgram_LastEdit = "2016.05.02"
'  고객 관리 부분 전체적으로 수정 및 속도 개선
'  기초관리에 점주명 추가및 검색 가능 하도록 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.15
'    strProgram_LastEdit = "2016.04.18"
'  고객정보에서 메모 항목 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.14
'    strProgram_LastEdit = "2016.04.11"
'  가맹점 품목할인 등록 / 가맹점 요일할인 등록에서 이전 적용 내역을 삭제 가능 하도록 수정
' (삭제 처리하면 가맹점에서 로그인시 삭제를 하도록 구성 하였다.)
'  SMS 충전할 경우 이전 일자로 설정되는 문제 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.13
'    strProgram_LastEdit = "2016.04.01"
'  가맹점 할인 현황 Excel 저장 가능 하도록 수정

'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.12
'    strProgram_LastEdit = "2016.03.31"
'  (본사) 대표 품목 등록 내용 기능 개선
'  - 삭제가 처리되지 않는 문제 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.11
'    strProgram_LastEdit = "2016.03.30"
'  (본사) 대표 품목 등록 내용 기능 개선
'  - 저장시 가맹점 가격 자료 자동 생성 여부 설정 가능 하도록 변경
'  - 금액이 변경된 경우 모든 가맹점의 금액을 일괄 변경 하도록 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.10
'    strProgram_LastEdit = "2016.03.24"
'   일일 판매 집계 (가맹점) 추가
'   매출 현황 (가맹점) 추가
'   일부화면 지사 선택 내용 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.09
'    strProgram_LastEdit = "2016.02.22"
'    지사 출고 검품 현황에서 선택된 지사만 처리하도록 처리하였으나 오류가 있어 오류 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.07
'    strProgram_LastEdit = "2016.02.22"
'    사고 유형 분석에서 금액도 같이 표시되도록 수정
'    지사 출고 검품 현황에서 선택된 지사만 처리하도록 수정
    
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.06
'    strProgram_LastEdit = "2016.01.15"
'    지사선택에서 조회 기능 추가(택관리,특정접수현황등)
'    수기 출고 등록 화면 개선
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.05
'    strProgram_LastEdit = "2016.01.12"
'    사고 유형별 분석자료 가맹점 현황 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.04
'    strProgram_LastEdit = "2015.12.24"
'    가맹점 현황에서 검색 조건에 현재 택코드로 검색 가능 하도록 수정(CS팀)
'    지사등록에서 프로그램사용 종료일자를 설정 가능 하도록 변경
    
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.03
'    strProgram_LastEdit = "2015.12.22"
'    사고 유형별 분석 자료 신규 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.02
'    strProgram_LastEdit = "2015.12.09"
'    사고 담당자(SMS)등록에서 해당 가맹점을 선택하여 등록 가능하도록 수정
'    이전 문자 충전 내용 삭제 가능 하도록 수정 (수정 기록이 남기 때문에 가능)
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.01
'    strProgram_LastEdit = "2015.11.12"
'    가맹점 검색에서 대표자명으로 검색이 가능 하도록 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.2.00
'    strProgram_LastEdit = "2015.09.16"
'    '고객별 미수금에서 사용마일리지 내용 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.99
'    strProgram_LastEdit = "2015.08.25"
'    '고객별 미수금에서 사용마일리지 내용 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.98
'    strProgram_LastEdit = "2015.06.29"
'    '월간 매출현황(일별 합계) 에 현 지사의 재고 수량을 표시 하도록 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.97
'    strProgram_LastEdit = "2015.06.23"
'    '문자메시지 발송에서 발송자 전화번호 처리 부분 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.96
'    strProgram_LastEdit = "2015.06.17"
'    '요일할인,할인관리에서 적용 대상이 되는 내용의 색상을 표시하여 구분 가능 하도록 수정
    
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.95
'    strProgram_LastEdit = "2015.05.15"
'    '가맹점 프로그램 사용종료일자 기능 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.94
'    strProgram_LastEdit = "2015.05.15"
'    '가맹점 프로그램 사용종료일자 기능 추가
    
    'SP_01001_01, SP_01001_03 수정
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.93
'    strProgram_LastEdit = "2015.05.15"
'    '가맹점 프로그램 사용종료일자 기능 추가
'    '가맹점 프로그램 최종 접속 일자 확인 기능 추가
    
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.92
'    strProgram_LastEdit = "2015.04.30"
'    'SMS 등록을 2012,2020 가능 하도록 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.91
'    strProgram_LastEdit = "2015.04.14"
'    'SMS 문자에서 가맹점상담자 발송 내역 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.90
'    strProgram_LastEdit = "2015.04.03"
'    'SMS 문자에서 가맹점상담자 등록및 발송 가능 하도록 기능 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.89
'    strProgram_LastEdit = "2015.03.31"
'    '매출관리->월간매출현황(합계)에 평균 단가 합 표시

'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.88
'    strProgram_LastEdit = "2015.03.24"
'    ' 미출고 정리 내역 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.87
'    strProgram_LastEdit = "2015.03.18"
    ' SMS 등록에서 사용자 2004, 2012 만 충전이 가능 하도록 수정
    ' 지사 미출고 정리에서 고객정보및 가맹점 입고 정보 기록
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.86
'    strProgram_LastEdit = "2015.02.11"
    ' SMS 등록에서 수량이 금액보다 클경우 등록이 되지 않고 경고 메시지 출력 하도록 변경

'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.84
'    strProgram_LastEdit = "2015.02.01"
    ' 매장 매출 현황(보고용-가맹점기준) 선택 방법 변경
    ' 매장 매출 현황(보고용-지사기준)   선택 방법 변경
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.83
'    strProgram_LastEdit = "2015.01.21"
    ' 매장 매출 현황(보고용-가맹점기준) 수정
    ' 매장 매출 현황(보고용-지사기준)   수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.82
'    strProgram_LastEdit = "2015.01.19"
    ' 매장 매출 현황(보고용-가맹점기준) 수정
    ' 매장 매출 현황(보고용-지사기준)   수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.81
'    strProgram_LastEdit = "2015.01.16"
    ' 매장 매출 현황(보고용-가맹점기준) 수정
    ' 매장 매출 현황(보고용-지사기준)   신규 생성
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.80
'    strProgram_LastEdit = "2015.01.12"
    ' 재고 조사 현황 기능 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.78
'    strProgram_LastEdit = "2014.12.03"
    ' 재고 조사 현황 기능 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.77
'    strProgram_LastEdit = "2014.11.25"
    '가맹점 품목별 접수현황(특정)
    '  - 가맹점 품목별 접수현황(특정) 가맹점 정보 추가 기능
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.76
'    strProgram_LastEdit = "2014.11.18"
    '가맹점 품목별 접수현황(특정)
    '  - 가맹점 품목별 접수현황(특정) 조회 및 출력 기능 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.75
'    strProgram_LastEdit = "+2014.11.05"
    '년간 매출 현황 수정
    '  - 가맹점 매출 기준으로 조회가 가능 하도록 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.74
'    strProgram_LastEdit = "2014.11.01"
    '협력사 발송 부분이 사라진 버그 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.73
'    strProgram_LastEdit = "2014.10.29"
    ' 품목별 출고 현황에서 매장이 나오지 않는 문제 수정
    
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.72
'    strProgram_LastEdit = "2014.10.16"
    ' SMS 등록시 입금자명 및 등록자 정보 추가
    ' SMS 등록 기간별 조회 화면 추가
    
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.71
'    strProgram_LastEdit = "2014.07.25"
    ' 요일 할인 관리에서 특정 매장을 수정할 경우 오류 내용 수정
    
'
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.70
'    strProgram_LastEdit = "2014.06.30"
'    ' 품목별 입고 현황에서 2가지 범위로 조회 가능 하도록 수정
    
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.68
'    strProgram_LastEdit = "2014.06.03"
    ' 가맹점 판매취소/환불 현황에서 전체 매장 조회 가능 하도록 수정 (1002)
    ' - 조회 쿼리 수정
    ' 사고품 결재 라인에서 과장->팀장 으로 변경
    
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.67
'    strProgram_LastEdit = "2014.06.03"
    ' 월간 사업장 매출현황(new)에 본사 정산분 추가
    ' 특정매장 분석에 로얄티및 카드 수수료 부분 추가
    ' 지사변경시 지사를 선택하여야만 적용되도록 변경
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.66
'    strProgram_LastEdit = "2014.05.22"
    ' 가맹점 마진 수정의 로그를 남겨 조회 가능 하도록 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.65
'    strProgram_LastEdit = "2014.05.16"
    ' 가맹점정보 저장부분 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.64
'    strProgram_LastEdit = "2014.04.17"
    ' 판매취소 환불 현황에서 환불 일자 구분으로 조회 가능하도록 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.63
'    strProgram_LastEdit = "2014.04.09"
    ' 연간 매출현황(월별 합계) 신규 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.62
'    strProgram_LastEdit = "2014.04.08"
    ' 가맹점 정보 저장시 가맹점 코드 자리수 확인
    ' 가맹점 정보 저장시 가맹점 택번호 자리수 확인
    ' 출고검품 엑셀 저장 내용 변경
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.61
'    strProgram_LastEdit = "2014.04.02"
    ' 로얄티신규 출력및 세부 내용 수정

'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.60
'    strProgram_LastEdit = "2014.03.25"
    ' sms발송현황에서 본사 내용 확인 가능 하도록 수정
    '  -- 매장 사고품및 전산실 발송 내역은 확인하지 못함.
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.59
'    strProgram_LastEdit = "2014.03.14"
    ' 매장 매출 현황(보고용)에서 합계 신장율 추가
    ' 크렌즈 겔러리 매출 현형조회 일부 수정
    
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.58
'    strProgram_LastEdit = "2014.03.10"
    ' 크렌즈 겔러리 매출 현형조회 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.57
'    strProgram_LastEdit = "2014.02.27"
    ' 품목별 입고 현황 쿼리 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.56
'    strProgram_LastEdit = "2014.02.27"
    ' 품목별 입고 현황 쿼리 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.55
'    strProgram_LastEdit = "2014.02.18"
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.54
'    strProgram_LastEdit = "2014.01.20"
    ' 품목별 출고 기간 현황에서 비율 내용 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.53
'    strProgram_LastEdit = "2014.01.14"
    ' 품목별 출고 기간 현황에서 비율 내용 추가
    
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.52
'    strProgram_LastEdit = "2013.12.30"
    ' 일일출고 현황에서 인쇄 버튼 클릭시 디세이블 시켜서 이중 클릭 방지
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.50
'    strProgram_LastEdit = "2013.12.12"
    ' 품목별 출고 기간 현황 화면 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.49
'    strProgram_LastEdit = "2013.12.11"
    ' 가맹점 매출혀황에 매출액 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.48
'    strProgram_LastEdit = "2013.11.15"
    ' 요일별 할인관리에서 시작일자와 종료일자가 같으면서 요일이 다른 경우 조회 오류 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.47
'    strProgram_LastEdit = "2013.11.14"
    ' 예상매출 등록 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.46
'    strProgram_LastEdit = "2013.11.07"
    ' 영업사원괄리에서 관리담당자 추가
    '  - 가맹점 정보에서 관리당당자 추가
    '  - 기타 조회 화면에서 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.45
'    strProgram_LastEdit = "2013.10.29"
    ' sms 이마트 내역 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.44
'    strProgram_LastEdit = "2013.10.29"
    ' 반품 현황 내용 추가 등록하여 매장에서 조회 가능 하도록 수정
    ' TB_반품현황 테이블 생성
    ' SP_03009_01 수정
    ' SP_03009_02 수정
    
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.43
'    strProgram_LastEdit = "2013.10.18"
    ' 1. 달성률 순위 조회 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.42   41
'    strProgram_LastEdit = "2013.10.17"
    ' 1. 가맹점 오픈/이동 현황 조회 화면 기능 추가
    ' 2. 사장님 요청 사항 화면 추가
    '   - 예상매출 등록(영업사워녈)
    '   - 매출관리 (달성률 조회)
    '   - 달성률 순위 조회
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.40
'    strProgram_LastEdit = "2013.10.08"
    ' 1. 가맹점 오픈/이동 현황 조회 화면 기능 추가
    
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.39
'    strProgram_LastEdit = "2013.09.23"
    ' 사고품 내역 지사별 조회 가능 하도록 수정
    ' 사용자 등록 화면 수정(지사 등록 가능하도록 )
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.38
'    strProgram_LastEdit = "2013.09.23"
    ' 월간매출 현황 일자별 합계에서 0000(전체) 지사 조회 가능하도록 수정
    ' 일별 매출현황(그래프) 일별, 월별 조회 가능 하도록 수정
    ' 가앰점 매출현황(지사기준)에서 0000(전체) 지사 조회 가능하도록 수정
    
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.37
'    strProgram_LastEdit = "2013.08.29"
    ' 매출현황 부분 타이틀 수정
    ' 매출현황 보고용 비고 내용 저장 가능하도록 수정
    ' 객수 화면 추가
    ' 매출조회 화면 단위 조정가능하도록
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.36
'    strProgram_LastEdit = "2013.08.26"
    ' 메일 작성 내용 지사에서 가능 하도록 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.35
'    strProgram_LastEdit = "2013.07.24"
    ' 매장 매출 현황 (보고용) 작성
    ' 특정매장 분석 선택사항 추가

'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.34
'    strProgram_LastEdit = "2013.07.02"
    ' 품목별 입고 현황 조회 부분 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.33
'    strProgram_LastEdit = "2013.06.10"
    ' 수기 출고시 해당 내용을 tb_입출고에 저장하지 않는 문제
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.32
'    strProgram_LastEdit = "2013.05.06"
    ' 1. 공문 내용 이미지 저장 가능하도록 수정
    
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.31
'    strProgram_LastEdit = "2013.04.15"
    ' 1. 지사 입고 검품 현황 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.30
'    strProgram_LastEdit = "2013.04.12"
    ' 1. 지사 입고 검품 현황 수정
    
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.29
'    strProgram_LastEdit = "2013.03.29"
    ' 1. 가맹점 기간별 매출 현황에서 지사 출고 수량을 출력
    ' 2. 월간매출 현황(일별 합계)에 지사 출고 수량및 재고 출력
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.26
'    strProgram_LastEdit = "2013.03.19"
    ' 1. 매출현황의 합계가 정렬 되면서 틀려지는 문제 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.25
'    strProgram_LastEdit = "2013.03.12"
    ' 1. 매출현황의 합계가 정렬 되면서 틀려지는 문제 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.24
'    strProgram_LastEdit = "2013.01.29"
    ' 1. 매출현황의 합계가 정렬 되면서 틀려지는 문제 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.23
'    strProgram_LastEdit = "2013.01.23"
    ' 1. 매출현황의 합계가 정렬 되면서 틀려지는 문제 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.21
'    strProgram_LastEdit = "2013.01.10"
    ' 1. 가맹점 구분 내용추가, 메모 기능 추가
    
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.01
'    strProgram_LastEdit = "2013.01.03"
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.1.01
'    strProgram_LastEdit = "2012.12.20"
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.99
'    strProgram_LastEdit = "2012.12.17"
    ' 1. 개별 출고 현황
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.99
'    strProgram_LastEdit = "2012.11.03"
    ' 1. 택번호 오류로 인한 매출 조회 내역 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.98
'    strProgram_LastEdit = "2012.10.17"
    ' 1. 미출고 정리에서 해당 기간 매장 조회가 안되는 문제 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.97 - 96
'    strProgram_LastEdit = "2012.09.11"
    ' 1. 기타 관리 에서 택관리 기능 수정
    ' 2. 사고 조회에서 더블클릭시 접수 상세 조회로 이동하도록 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.95
'    strProgram_LastEdit = "2012.09.10 "
    ' 1. 사고  sms 발송자 삭제 기능 추가
    ' 2. 기타 관리 에서 택관리 기능 추가
     
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.94
'    strProgram_LastEdit = "2012.08.30 "
    ' 1. sms 발송 현황 분석 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.93
'    strProgram_LastEdit = "2012.08.18 "
    ' 1. 가맹점 등록 내용 에서 지사 저장시 오류 날 수 있는 부분 수정
    ' 2. 가맹점 등록에서 0000 전체 가맹점 조회시 모든 가맹점이 나오도록 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.91, 1.0.92
'    strProgram_LastEdit = "2012.08.16 "
    ' 1. 고객별 미수금 현황조회 수정, 엑셀 저장되도록 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.90
'    strProgram_LastEdit = "2012.08.09 "
    ' 1. 가맹점 기간별 매출현황(합계)
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.89
'    strProgram_LastEdit = "2012.07.27 "
    ' 1. 품목별 접수 현황에서 전체 조회가 되도록 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.88
'    strProgram_LastEdit = "2012.07.24 "
    ' 1. 가격관리, 요일할인관리, 할인관리에 수신일자 정보 추가
    
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.87
'    strProgram_LastEdit = "2012.07.16 "
   ' 1. 미출고 정리가 안되는 문제 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.86
'    strProgram_LastEdit = "2012.07.11 "
    ' 1. 요일 오류 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.85
'    strProgram_LastEdit = "2012.07.10 "
    ' 1. 가맹점 구분에 크렌즈 겔러리 추가
    ' 2. ExportToExcel 파일 저장할 경우 insert한 내용의 다음줄이 저장되지 않는 문제 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.84
'    strProgram_LastEdit = "2012.07.02 "
    ' 1. 지사 변경이 있을경우 지사를 저장하고 다시 상단의 저장버튼을 눌러서 지사 정보가
    '    다시 이전 지사로 변경되는 버그 수정
    ' 2. 사고품 내역에서 처리일자 관련 수정
    
    'strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.83
    'strProgram_LastEdit = "2012.06.22 "
    ' 1. 마트 인력 협력인 리스트 삭제및 코드 확인 기능 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.82
'    strProgram_LastEdit = "2012.06.22 "
    ' 1. 마트 인력 협력인 리스트 기능 추가
    ' 2. sms 문자 발송 기능 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.81
'    strProgram_LastEdit = "2012.06.13 "
    ' 1. 일일출고 가맹점 현황 프로그램 추가
    ' 2. 미출고 현황에서 판매 취소된 내용이 조회도는 문제 수정
    ' 3. 미출고 정리에서 본사에서 조회가 안되는 문제 수정
    ' 4. 지사에서 출고 처리한 내용을 가맹점에서 수동으로 입고를 잡은 경우
    '    지사 출고정보가 매장의 자료로 업데이트 되면서 출고 자료가 사라지는 문제 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.80
'    strProgram_LastEdit = "2012.06.07 "
    ' 1. 가맹점 오픈/이동 현황 조회 화면 추가
    ' 2. 가맹점 기본정보 수정
    ' 3. 운영일수 표시
    ' 4. 가맹점매출현황(지사기준)에서 종료택번호 관련 오류 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.79
'    strProgram_LastEdit = "2012.06.07 "
    ' 1. 메일 관련 내용 수정
    ' 2. 가맹점 기간별 매출현황(합계) 인쇄 기능 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.78
'    strProgram_LastEdit = "2012.05.24 "
    ' 1. 가격자료 생성시 분류 등록(%)이 먼저 되어 있어야 생성 하도록 수정
    ' 2. 가맹점현황 저장시 폐점일 경우 택,자사 종료일자 설정하도록 수정
    ' 3. 특정일자 조회에서 폐정 현황에의 의하여 조회 되도록 수정
    
    
    'strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.77
    'strProgram_LastEdit = "2012.05.24 "
    ' 1. 특정 매장 분석 자료
    ' 2. 목요세일 -> 요일세일 변경
    ' 3. 로열티 분석에서 5% 추가및 인쇄 변경
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.76
'    strProgram_LastEdit = "2012.05.22 "
    ' 1. 지사코드도 암호화 하여 변경할 수 없도록 설정
    ' 2. sms 조회 관련 내용 수정
    ' 3. 가맹점별 매출집계(특정일) 색상 조정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.75
'    strProgram_LastEdit = "2012.05.21 "
    ' 1. 로그인창에서 최초 서버 설정창 추가
    ' 2. 폐점 내용도 접수 현황에서 조회 가능하도록 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.74
'    strProgram_LastEdit = "2012.05.16 "
    ' 1. 최초 서버 설정창 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.73
'    strProgram_LastEdit = "2012.05.16 "
    ' 1. 셀 지정 하여 복사 가능 하도록 수정
    ' 2. 가맹점 일일 매출 조회시 색상 보기 기능 추가
    ' 3. 폐점한 가맹점 가맹점 리스트에 적색으로 표시 되도록 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.72
'    strProgram_LastEdit = "2012.05.03 "
    ' 1. 사고품 내역 수정 (저장시 처리일자 등록하도록 수정)
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.71
'    strProgram_LastEdit = "2012.05.01 "
    ' 1. 사고품 내역 저장 오류 수정
    ' 2. 매출 조회시 택번호 오류로 인한 런타임 오류 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.70
'    strProgram_LastEdit = "2012.04.26 "
    ' 1. 실시간 조회 오류 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.69
'    strProgram_LastEdit = "2012.04.26 "
    ' 1. 각종 조회 화면에서 단가 출력 오류 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.68
'    strProgram_LastEdit = "2012.04.18 "
    ' 1. 사고 접수 내용 보안및 출력물 수정
    ' 2. 가맹점별 매출 집계(특정일)에 매출이 없어도 매장명이 나오도록 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.67
'    strProgram_LastEdit = "2012.04.12 "
    ' 1. 기타수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.66
'    strProgram_LastEdit = "2012.04.06 "
    ' 1. 각종 매출 조회에서 매출이 없어도 매장은 나타나도록 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.65
'    strProgram_LastEdit = "2012.04.06 "
    ' 1. 실시간 매출 현황  처리
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.64
'    strProgram_LastEdit = "2012.04.05 "
    ' 1. 복사 및 붙여넣기 기능 추가
    ' 2. 월간매출현황(일별합계) 추가
    ' 3. 실시간 매출 현황 일부 처리
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.62
'    strProgram_LastEdit = "2012.04.02 "
    ' 1. 예상매출 입력 내용 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.61
'    strProgram_LastEdit = "2012.03.29 "
    ' 1. 택번호 변경시 적용일자를 수정 하지 않으면 저장이 되지 않도록 수정
    ' 2. 매출 메뉴 정리
    ' 3. 예상 매출 관리 메뉴 추가
    ' 4. 로열티 메뉴 인쇄 기능 추가
    
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.60
'    strProgram_LastEdit = "2012.03.12 "
    ' 1. 품목할인 및 요일을 매장별로 별도 수정 가능 하도록 추가
    ' 2. 본사 대표 품목 저장시 본사 품목을 기준으로 저장 되도록 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.59
'    strProgram_LastEdit = "2012.02.01 "
    ' 1. 일일출고 현황에서 출력 장수 최종 선택한 것으로 출력 되도록 수정
    ' 2. 지사출고 확정을 지을 경우 소요 시간 분석 할 수 있도록 처리
    ' 3. 지사출고 처리일자가 당일보다 이전일 경우 확인 메시지 처리
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.58
'    strProgram_LastEdit = "2012.02.01 "
    ' 1. 매출 관련 지사 이동전 매장이 안보이는 문제 해결
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.57
'    strProgram_LastEdit = "2012.01.02 "
    ' 1. 사고품 관련 내용 수정 (당해 년도만 보이는 문제 수정)
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.56
'    strProgram_LastEdit = "2012.01.02 "
    ' 1. 지사별 자료를 본사만 조회 되도록 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.55
'    strProgram_LastEdit = "2011.12.26 "
    ' 1. 사고품 관리 전화 번호 입력 및 가맹점에서 설정 하도록 변경
    ' 2. P_02005_01 - 가맹점 품목별 접수현황(상세) 기능 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.54
'    strProgram_LastEdit = "2011.12.09 "
    ' 1. 외주 입출고 관리 적용
    '
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.53
'    strProgram_LastEdit = "2011.12.02 "
    ' 1. 사고 접수 보고서 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.52
'    strProgram_LastEdit = "2011.11.09 "
    ' 1. 외주 관련 내용 전면 수정
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.51
'    strProgram_LastEdit = "2011.11.02"
    ' 1. 가맹점 가격자료 검색시 적용일자별로 검색 가능 하도록 수정
    ' 2. 개별 출고 현황에서 순수 택번호로만 조회가 가능하도록 수정
    ' 3. 매장 찾기 기능 추가
    
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision  ' 1.0.50
'    strProgram_LastEdit = "2011.11.02"
    ' 1. 가맹점 관리 내역 수정
    ' 2. 수기 일일 출고
    ' 3. 회원관리
    
    
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision
'    strProgram_LastEdit = "2011.10.27"
    ' 원격 지원 파일이 없을경우  다운로드및 실행
    ' 툴바 종료 버튼 기능
    ' 출고 확정시 종료 버튼 비활성화
    ' 가맹점 등록시 지사 코드
    
'   strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision
'   strProgram_LastEdit = "2011.03.23"
    ' 원격 지원 파일 다운로드및 실행
    ' 툴바 아이콘 추가
    
'    strProgram_Version = "1.0.48"
'    strProgram_LastEdit = "2011.03.23"
    ' P_01001 (가맹점 현황) 바코드택, 지사 현황 삭제 오류 수정
    ' P_02008 (지사 입고검품 현황) 오류 수정
    ' P_03015 (지사출고 검품 현황) 출고 확정중 가맹점 리스트를 클릭할 경우 잘못 처리되는 문제 수정
    

End Function

