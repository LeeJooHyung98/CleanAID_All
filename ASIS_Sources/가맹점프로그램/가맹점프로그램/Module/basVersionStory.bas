Attribute VB_Name = "basVersionStory"
Option Explicit

' 프로그램 버전 정보
Public Program_Version As String
Public Program_LastEdit As String

Public Function Set_ProgramVersion()

'   신용카드 승인현황 Excel 파일로 변환가능 하도록 수정
'   현금영수증 승인현황 Excel 파일로 변환가능 하도록 수정
    
    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.74
    Program_LastEdit = "2018-02-06"
    
'   로그 파일 오류 내용 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.74
'    Program_LastEdit = "2017-03-30"
    
'   출고 결제에서 IC 인증인 경우에도 출고 영수증 출력일 경우 카드 승인 전표가 나오는 문제 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.72
'    Program_LastEdit = "2017-03-29"
'   출고 결제에서 IC 인증인 경우에도 출고 영수증 출력일 경우 카드 승인 전표가 나오는 문제 수정
    
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.71
'    Program_LastEdit = "2017-03-29"
'   요일할인, 할인 정보 수신 업데이트를 품목별로 하도록 처리

'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.68 -70
'    Program_LastEdit = "2017-03-29"
'   업데이트 경로 변경
'   "www.clean-aid.co.kr:8090/cleanaid" => "www.insoftnet.com/upgrade/cleanaid"
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.67
'    Program_LastEdit = "2017-03-07"
    ' 1. 세탁마진을 가맹점에서 임의로변경 했는지의 여부를 확인 하도록 수정
    ' 2. 할인 자료 업데이트시 받은 갯수와 처리 갯수가 모두 같은 경우만 처리되도록 변경

'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.66
'    Program_LastEdit = "2017-03-07"
    ' 1. 세탁마진을 가맹점에서 임의로변경 했는지의 여부를 확인 하도록 수정
    
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.64
'    Program_LastEdit = "2017-02-01"
    ' 1. 의류분류 등록의 순서가 변경되지 않는 문제 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.63
'    Program_LastEdit = "2017-01-31"
    ' 1. 세탁마진,수선마진,외주마진을 소수점 입력 가능하도록 수정
    ' 2. 의류분류 현황에서 소수점 표시되도록 변경
    ' 3. 각종 화면에서 세탁마진,수선마진,외주마진 변경 내용 적용
    ' 4. TB_입출고 저장에서 마진내용을 소수점까지 저장
    ' 5. 일일매출 마감, 일일판매 현황에서 접수집계 내용에서 가맹점 마진을 표시


'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.62
'    Program_LastEdit = "2017-01-25"
    ' 1. 접수에서 수선은 접수하지 못하도록 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.61
'    Program_LastEdit = "2016-11-02"
    ' 1. 출고에서 메모가 1줄로 나오는 문제 해결
    '
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.60
'    Program_LastEdit = "2016-09-28"
    ' 1. 접수를 2곳에서 잡는걸로 설정할 경우 판매 취소한 택을 재 사용할경우 중간에 사용한 번호가
    '    중복으로 잡히는 문제 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.59
'    Program_LastEdit = "2016-09-27"
    ' 1. 080으로 등록된 고객은 문자가 발송되지 않도록 변경

'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.58
'    Program_LastEdit = "2016-09-21"
    ' 1. 마감후 프로그램을 재 실행하지 않을경우 요일 행사가 적용되지 않는 문제 수정
    ' 2. 현금영수증을 소득공제에서 사업자로 변경할 경우 사용자 정보 입력 하도록 변경
    ' 3. 승인시도후 최소한 경우 다시 승인이 안되는 문제 해결
    ' 4. 일일마감에서 가맹점 마진에서 전단위 반올림이 안되어서 몇원 손해보는 문제 수정
    ' 5. 행사용 문자에서 정규 메시지 삽입 ( 최대 55자까지 전송 가능) 하도록 변경
    '        (광고)                         <---- 첫번째 줄에 삽입'
    '        무료수신거부 080-863-5771      <---- 마지막 줄에 삽입'
    
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.57
'    Program_LastEdit = "2016-06-16"
    ' 1. KS4060단말기 사용매장에서 카도로 결제시 보관증미출력을 선택한 경우에도 보관증이 나오는 문제 수정
    ' 2. 할인관련 삭제시 로그 기록 하도록 수정
    ' 3. 본사 자료 삭제시 가맹점에서 매장 할인 정보를 한번더 확인한 후 삭제 하도록 수정
    ' 4. 세탁물 인도문자 쿼리를 프로시저로 변경 처리속도 향상
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.55
'    Program_LastEdit = "2016-05-31"
    ' 1. 접수에서 가격이 간혹 잘못 나오는 버그 수정.
    '     (품목 상세에서 키보드로 조정하여 처리할 경우 최상위 품목이 그대로 나옴...)
    ' 2. 접수, 출고시 고객 조회할 경우 2명 이상이 있을 경우 고객 조회 오류나는 현상 수정
    '     (조회 되고 있을때 엔터키를 치면 2중 조회 처리가 되면서 오류가 나타나는 현상 수정)

'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.54
'    Program_LastEdit = "2016-05-26"
    ' 1. 일일마감에서 카드 수수료 계산시 카드승인 금액 오류 ㅠㅠ
    '    (당일 승인하고 취소하면 취소 금액에는 포함이 되고 승인 금액에는 포함이 안되는 문제 수정)
    
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.53
'    Program_LastEdit = "2016-05-23"
    ' 1. 일일판매 집계에 미수금 수금 정보 표시
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.52
'    Program_LastEdit = "2016-04-27"
    ' 1. 접수시 작업내용에 들어가서 취소한 경우 이전 가격으로 되돌아 오지 않는 문제 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.49
'    Program_LastEdit = "2016-04-21"
    ' 1. IC카드 관련(ks-4060)단말기 사용시 출고결제에서 처리한 경우 영수증이 출력 되는 문제 수정

'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.48
'    Program_LastEdit = "2016-04-19"
    ' 1. 할인내용 수신 완료 처리가 안되는 문제 수정
    ' 2. 출고 화면에서 입고내용 출력하여 물건 찾는 용도로 활요 하도록 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.47
'    Program_LastEdit = "2016-04-11"
    ' 1. 요일할인및 행사 내용을 삭제 처리하도록 수정

'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.46
'    Program_LastEdit = "2016-01-12"
    ' 1. 출고에서 출고 취소시 이전 일자를 암호 입력하면 취소 가능 하도록 수정
    ' 2. 접수 구분에서 할증내역에 보풀제거,털제거 추가

'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.45
'    Program_LastEdit = "2016-01-12"
    ' 1. 사고품 내역에 수축,변형 추가
 

'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.44
'    Program_LastEdit = "2015-12-09"
    ' 1. 신용카드 승인시 종료 버튼이 비활성화 되도록 변경
    '   - 현금영수증 금액이나 신용카드 승인 내용이 있을 경우 비활성화하여 종료 방지
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.43
'    Program_LastEdit = "2015-11-24"
    ' 1. 신용카드 결제시 버튼 더블클릭 방지
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.42
'    Program_LastEdit = "2015-10-29"
    ' 1. 마감시 현금 결제 부부느이 금액이 현금결제+미수금현금결제가 나오는 문제 수정
    ' 2. 미수금 현금 결제 내용이 표시 되지 않는 문제 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.41
'    Program_LastEdit = "2015-10-13"
    ' 1. SMS 발송화면에서 수량을 다시 조회 가능 하도록 수정
    ' 2. 2015-10-16일 부터 시행하는 발신번호 사전 등록제 적용을 위하여 발신 번호를 변경 하지 못하도록 수정
    
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.40
'    Program_LastEdit = "2015-09-30"
    ' 1.IC 카드 리턴값으로 메시지 처리
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.39
'    Program_LastEdit = "2015-09-30"
    ' 1.IC 카드 관련 승인시작 버튼 삽입
    '    - 할부, 거래 구분을 선택 하지 못하는 부분을 적용하기 위하여.
    ' 2. 사고품의 구매일자 필수 입력사항으로 설정


'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.38
'    Program_LastEdit = "2015-09-09"
    ' 1. KSNET OCX 관련 업그레이드 적용 관련 메시지 추가
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.37
'    Program_LastEdit = "2015-09-04"
    ' 1. KSNET 2015-09-01일 버전 3.2.0.7
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.36
'    Program_LastEdit = "2015-09-04"
    ' 1. KSNET 2015-09-01일 버전 3.2.0.7

'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.35
'    Program_LastEdit = "2015-09-02"
    ' 1. 여신법 변경으로 인한 KS4060적용
    ' 2. KS4060 보안 모듈 적용
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.34
'    Program_LastEdit = "2015-08-18"
    ' 1. SMS 발신자 번호 관련 부분 수정

'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.33
'    Program_LastEdit = "2015-07-14"
    ' 1. 세탁물 지연 문자 조회 속도 개선
    ' 1. SMS 관련 부분 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.32
'    Program_LastEdit = "2015-07-14"
    ' 1. 세탁물 인도 문자 조회 속도 개선
    
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.31
'    Program_LastEdit = "2015-06-17"
    
    ' 1. 구두/운동화 세탁 안내 문구 추가
    ' 2. 세탁물 인도 문자 속도 개선


'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.30
'    Program_LastEdit = "2015-06-17"
    ' 1. 최종거래일자에 시간이 안들어가는 문제 수정 (언제 수정했는지 모름.)
    ' 2. 지사 정보를 다운로드 하도록 수정
    ' 3. 지점입고에서 지사 정보 표시및 코멘트 추가
    ' 4. 할인 정보에서 적용 대상이 되는 항목 색상 변경으로 확인 가능하도록 처리
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.29
'    Program_LastEdit = "2015-05-15"
    ' 1. 가맹점 프로그램 사용 종료 일자 기능 추가
    ' 2. 최종 접속 정보 추가
    
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.28
'    Program_LastEdit = "2015-05-06"
    ' 1. 주의대상 파일 복구 명령어 저장
    ' 2. 각종 불필요한 폴더 자료 삭제 처리
    ' 3. 삭제 마일리지 처리후 본사 전송여부를 변경하지 않아서 최종 수정 자료가 본사로 전송 되지 않은 문제 수정
    '    - 마일리지 삭제후 본사 전송여부 'N'로 설정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.27
'    Program_LastEdit = "2015-01-12"
    ' 1. 택번호 수정시 잘못 입력되는 부분을 확인 하도록 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.26
'    Program_LastEdit = "2014-12-12"
    ' 1. 미수금액 처리과정 오류 수정 최종 금액에 판매 취소 분만 차감되는 문제 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.25
'    Program_LastEdit = "2014-12-04"
    ' 1. 마일리지 사용과 미수금액으로 처리한 부분을 판매 취소할 경우 마일리지 금액 만큼 오류 나는 문제 수정
    ' 2. 서버 자료 수신시(재설치) 사고품 내역은 모두 가저 오도록 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.24
'    Program_LastEdit = "2014-11-25"
    ' 1. 세탁환불, 반품환불 오류 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.23
'    Program_LastEdit = "2014-11-25"
    ' 1. 고객 정보창 이동 위치 조정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.22
'    Program_LastEdit = "2014-11-20"
    ' 1. 일일마감에서 쿠폰 금액의 비율을 6:4로 조정 하여 처리
    ' 2. 세탁환불, 반품환불을 할경우 해당 택번호가 판매 취소된 택번호일 경우 일일 마감에 중복 처리되는 문제 수정
    ' 3. 접수- 신규 고객등록시 이름이 반드시 입력 되도록 수정
    ' 4. 동명이 검색될 경우 출고 경고 메시지 문구 추가
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.21
'    Program_LastEdit = "2014-10-24"
    ' 1. 신용카드 매장 보관용 출력 여부를 가맹점에서 결정할 수 있도록 변경

'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.20
'    Program_LastEdit = "2014-10-06"
    ' 1. 반품환불, 세탁환불을 마감이후에 저장할 경우 적용되지 않는 문제 수정
    ' 2. 서버 자료 수신은 이전 지사부터 하도록 변경(최근 지사부터 할 경우 미출고 자료가 생기는 문제가 있었음)
    ' 3. 서버 자료 수신중 입출고 자료만 다시 수신할 수 있도록 변경
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.19
'    Program_LastEdit = "2014-09-23"
    ' 1. 무실동점 사용중지 처리 (100576) - 지사를 이용하고 있지 않아서.

'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.18
'    Program_LastEdit = "2014-07-21"
    ' 1. 일일매출마감에서 반품환불,세탁환불이 있을 경우 해당택번호가 판매 취소된 정보가 있을 경우
    '    판매취소된 내역도 같이 집계가 되는 문제 수정

'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.18
'    Program_LastEdit = "2014-07-21"
    ' 1. 현금영수증 및 신용카드 영수증에 부가세 표시 되도록 수정
    ' 2. SMS 문자 전송을 80자에서 90자로 변경
    ' 3. 로그인시 공지사항 확인 여부 기능 개선

'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.17
'    Program_LastEdit = "2014-04-21"
    ' 1. 마감후 접수 받았을 경우 접수시간 정보는 당일 출력하고 옆에 저장영업일 별도 표시
    ' 2. 출고 화면에서 1```````이미 출고된 부분을 선택시 미불,완불이 변경되는 문제 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.17
'    Program_LastEdit = "2014-04-21"
    ' 1. 2014-04-22일 100570-청라1동점만 선불 현금 결제시만 적용 되도록 수정
    ' 2.세탁물 인도문자에서 반품 현황 색상 변경
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.16
'    Program_LastEdit = "2014-04-16"
    ' 1. 출고취소를 할경우 출고 시간이 초기화 되지 않는 문제 수정


 '   Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.16
 '   Program_LastEdit = "2014-04-15"
    ' 1. 세탁물 인도문자에서 반품 현황이 있는 부분은 기본적으로 전송되지 않도록 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.15
'    Program_LastEdit = "2014-04-03"


    ' 1. 서버 자료 수신 부분에서 연결 오류가 발생하여도 다시 시도 하도록 수정
    '    - 입출고 정보는 수신한 부분 이후 부터 다시 시작 하도록 처리
    '    - 나머지 부분은 그냥 해당 내용을 처음부터 다시 수신하도록 처리


'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.15
'    Program_LastEdit = "2014-04-03"
    ' 1. 부분결제 되어 있는 부분에서 0원인 제품을 판매 취소할 경우 미수금액이 더블로 잡히는 문제 수정
    '    -- 카드 취소후 결제 창이 떠서 미수금액이 다시 잡혀 버렸음
    ' 2. 부분결제일 경우 판매취소시 미수금액이 적을 경우 미수금액을 0원으로 설정
    
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.14
'    Program_LastEdit = "2014-03-31"
    ' 1. 부분결제 되어 있는 부분에서 0원인 제품을 판매 취소할 경우 미수금액이 더블로 잡히는 문제 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.13
'    Program_LastEdit = "2014-03-26"
    ' 1. 2014-04-01일 부터 크렌즈 겔러리 신규고객 10%할인되지 않도록 수정

'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.12
'    Program_LastEdit = "2014-03-10"
    ' 1. 일일마감에서 매장 수익금 내역 수정
    ' 2. 일일마감현황 출력에서 지사정산금액이 10만원이 넘으면 표시가 안되는 문제 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.11
'    Program_LastEdit = "2014-03-10"
    ' 1. 로열티및 카드 수수료 지원 자료를 다운로드 받지 못하는 문제 수정

'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.11
'    Program_LastEdit = "2014-03-03"
    ' 1. 일일매출 마감에 판매취소 후 카드결제 금액이 포함되지 않는 문제 수정
    ' 2. 현금 환불이 있을 경우 미수금액에 포함되는 부분 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.01
'    Program_LastEdit = "2014-02-26"
    ' 1. 일일매출 마감에 매장 수익금액이 마일리지 포함으로 표시되는 문제 수정
    ' 2. 보관증을 1장 출력시 카드전표및 현금 영수증이 1장만 나오는 문제 수정


    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.3.00
'    Program_LastEdit = "2014-02-25"
    ' 1. 일일매출 마감에 로열티 부분 적용
    '  - 로열티1, 로열티2, 수수료 지원
    ' 2. 가맹점 정보에 로열티1, 로열티2, 수수료 지원 정보 등록 하도록 변경
    ' 3. frm일일판매집계 부분 위의 내용과 같이 수정
    
 ' 작업 내용
' 전송 프로그램 수정 (일일마감 자료 전송, 카드 승인자료 취소일자 추가, 가맹점 정보 확인)
' 서버 자료 수신 내용 확익( 일일마감, 카드승인, 가맹점 정보)
' 마감작업시 일일마감_Send 자료 수정
' 로열티.

' LAUNDRY1000 TB_가맹점 변경 지사 뷰테이블 수정
        '로열티여부1, 로열티비율1, 로열티여부2, 로열티비율2, 수수료지원여부, 수수료지원비율
        
    ' LAUNDRY1000 tb_일일마감 변경 지사 뷰테이블 수정
'        alter table tb_일일마감 add 로열티정보1 nvarchar(11)
'        alter table tb_일일마감 add 로열티정보2 nvarchar(11)
'        alter table tb_일일마감 add 수수료정보  nvarchar(11)
'        alter table tb_일일마감 add 반품환불지사금액  int
'        alter table tb_일일마감 add 세탁환불지사금액  int
'
'        alte2r table tb_일일마감 add 카드취소금액  int
'        alter table tb_일일마감 add 카드취소건수  int
'
'        alter table tb_일일마감 add 로열티금액1  int
'        alter table tb_일일마감 add 로열티금액2  int
'        alter table tb_일일마감 add 수수료승인금액   int
'        alter table tb_일일마감 add 수수료취소금액   int
    
    
''    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.83
'    Program_LastEdit = "2014-01-16"
    ' 1. 카드번호 9-12자리 까지 **** 로 표시
    '    - 화면에 표시하는 부분도 카드번호 16자리만 표시
    '    - 이전 자료 모두 일괄 적용
    '    - 유효기간 자료 저장하지 않음
    ' 2. 문자 메시지의 특정 발송 번호 검색 기능 추가
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.82
'    Program_LastEdit = "2013-11-29"
    ' 1. 마감후 접수 받은 품목(다음날)을 판매 취소할 경우 판매 취소 매출이 당일로 잡혀 일마감이 맞지 않는 문제 수정
    ' 2. 최소 마일리지 적용을 본사에서 설정 하도록 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.81
'    Program_LastEdit = "2013-10-29"
    ' 1. 마일리지 최소 사용금액을 3,000->1,000으로 변경(크렌즈 겔러리 제외) 2013-11-01일 부터
    ' 2. 품명란에 색상으로 접수 위치 표시
    ' 3. 신규고객 접수시 고객성명란옆에 "작성후 엔터"문구 추가
    ' 4. 기본고객등급을 A.크 ->C.에로 변경
    ' 5. 반품 현황 내역 조회 기능 추가
    ' 6. 휴대폰 번호를 반드시 입력하도록 수정
    ' 7. 세탁비 환불현황에서 판매 취소 내역도 같이 나오는 문제 수정
    
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.80
'    Program_LastEdit = "2013-10-07"
    ' 1. 2013-10-21일부터 W코드을 크렌즈 겔러리에 나타나지 않도록 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.79
'    Program_LastEdit = "2013-10-07"
    ' 1. 일일매출현황에서 조회및 출력내용이 일일마감과 미수금액이 다르게 나오는 문제 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.78
'    Program_LastEdit = "2013-09-24"
    ' 1. 행사 안내용 문자에서 조회 구분을 추가하여 등록일로도 조회가 가능 하도록 수정
    '    - 등록만 하고 이용을 안하는 고객에게도 발송하기 위하여(수선 고객등록 하는 경우등등)
    ' 2. 본사관리->공지사항 조회 오류 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.78
'    Program_LastEdit = "2013-08-30"
    ' 1. 출고화면에서 출고 내역에서 고객을 조회할 경우 출고및 기타 버튼이 활성화 되지 않는 문제 수정
    
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.77
'    Program_LastEdit = "2013-08-29"
    ' 1. 출고내역에서도 미불완불 수정 가능하도록 수정
    ' 2. 출고 화면에서 출고 내역을 선택할 경우 사용하지 못하는 버튼 비활성화 처리
    ' 3. 2013-09-01일 자연수 130% 적용 하도록 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.76
'    Program_LastEdit = "2013-07-12"
    ' 1. 보관증 재출력시 사업자 정보 출력 되도록 수정
    ' 2. 문자 발송 최종 전송일자 선택 취소 문구 변경
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.75
'    Program_LastEdit = "2013-07-05"
    ' 1. 상표 수정시 삭제되는 부분 코드 보강
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.74
'    Program_LastEdit = "2013-07-03"
    ' 1. 상표 수정시 삭제되는 부분 코드 보강
    ' 2. 당일 마감일 경우 프로그램을 종료한부분 원상복구
    '   - 마감후 매출현황 출력 문제 때문에.
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.73
'    Program_LastEdit = "2013-07-01"
    ' 1. 2013-07-01일 부터 카드, 현금 결제시에만 마일리지 3% 적용 되도록 수정
    '    - 크렌즈 갤러리 제외(1024)
    ' 2. 당일 마감일 경우 프로그램을 종료 한다.
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.72
'    Program_LastEdit = "2013-05-20"
    ' 1. 고객별 미출고 현황 내역 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.70
'    Program_LastEdit = "2013-05-16"
    ' 1. 공지사항 팝업 삭제
    ' 2. 2013-05-15일 부터 팝업 부분 파일로 강제 적용
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.68
'    Program_LastEdit = "2013-05-14"
    ' 1. 접수현황, 매출, 일일마감 부분수정
    '  - 마일리지를 사용한 금액이 판매취소가 발생할 경우 미수금액이 잘못 계산되는 문제 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.67
'    Program_LastEdit = "2013-05-07"
    ' 1. 원격지원 실행 문제 해결
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.66
'    Program_LastEdit = "2013-05-06"
    ' 1. DBUpdate.exe 파일 실행 내용 변경
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.65
'    Program_LastEdit = "2013-05-06"
    ' 1. 가격 적용 순서 변경
    ' -- 기존: 행사->요일->일반
    ' -- 변경: 행사->요일->일반 순으로 적용은 같으나. 행사,요일이 중복될 경우 낮은 금액이 적용 되도록 수정
    ' 2. 접수 화면 상단에 적용된 금액 표시
    ' 3. 메시지 전송시 이미지 파일 가능 하도록 수정
    '   - 관련 TB_공지사항파일 테이블 신규 생성
    '   - 관련 프로그램 모두 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.64
'    Program_LastEdit = "2013-04-18"
    ' 1. SMS에서 ' 입력이 안되도록 수정
    ' 2. 지점입고 출력물 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.63
'    Program_LastEdit = "2013-04-05"
    ' 1. 미불 ,완불 수정 내역 변경
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.62
'    Program_LastEdit = "2013-04-01"
    ' 1. 사고품 접수시 오류 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.61
'    Program_LastEdit = "2013-02-27"
    ' 1. 사고품 접수시 접수일자 수정하지 못하도록 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.60
'    Program_LastEdit = "2013-02-27"
    ' 1. SMS 문자 발송시 크레임구분도 함께 발송하도록 수정

'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.58
'    Program_LastEdit = "2012-11-29"
    ' 1. 코드류 품명 변경 '코드' -> 코트
    ' 2. 기타 버그 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.56
'    Program_LastEdit = "2012-11-15"
    ' 1. 수선관련 내용 삭제
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.55
'    Program_LastEdit = "2012-10-24"
    ' 1. 수선관련 내용 삭제
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.55
'    Program_LastEdit = "2012-10-09"
    ' 1. 크렌즈 갤러리 백화점 전 직원 세탁 세일 건 ( 직원세탁으로 전품목 30%) 버튼을 만듬
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.54
'    Program_LastEdit = "2012-09-20"
    ' 1. 신용카드 서명이 되면 취소 버튼 비활성화 처리
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.53
'    Program_LastEdit = "2012-09-20"
    ' 1. 접수시 완불 후불 처리로인하여 미수금이 잘못 되는 문제 수정
    ' 2. 자연수- 크렌즈 현대 본점은 안되도록 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.52
'    Program_LastEdit = "2012-09-18"
    ' 1. 보관증 출력시 일부 컴퓨터에서 출력되지 않는 문제 수정
    '  (출력이 완료되기 전에 다음 출력 명령이 전달되어 출력되지 않아 딜레이 타임을 설정했다.)
    
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.51
'    Program_LastEdit = "2012-09-05"
    ' 1. 판매 취소시 카드 취소 문제 해결
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.5
'    Program_LastEdit = "2012-08-30"
    ' 1. 판매 취소시 카드 취소 문제 해결
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.04
'    Program_LastEdit = "2012-08-17"
    ' 1. 부분계산일경우 완불 / 미불 계산이 잘못 되는 부분 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.03
'    Program_LastEdit = "2012-08-17"
    ' 1. 무엇이든 물어보세요 기능 추가
    
    '    -- 부분 결제 판매 취소시 미수금액이 잘못 계산 되는 문제가 있음.
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.02
'    Program_LastEdit = "2012-07-11"
    ' 1. 부분결제시 접수에서 완불, 부분, 미불로 등록 되도록 수정
    
    '    -- 부분 결제 판매 취소시 미수금액이 잘못 계산 되는 문제가 있음.
    ' 2. 문자메시지에서 반품 수량 가저오는 오류 수정
    ' 3. 신용카드 승인현황 조회에서 일부 카드번호를 ****로 처리
    ' 4. 출고현황 인쇄 부분 수정
    ' 5. 크렌즈 겔러리 로그인 화면 로고 수정
    ' 6. 사고품에서 판매 취소 내용이 있을 경우 오류 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.2.01
'    Program_LastEdit = "2012-07-04"
    ' 1. 전송관련 테스트로 인한 버전업
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.93
'    Program_LastEdit = "2012-07-04"
    ' 1. DBUpdate.exe 파일 다운로드 받도록 수정
    ' 2. 사고품 저장하지 않고 출력되는 문제 수정
    ' 3. 전송관련 테스트로 인한 버전업
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.92
'    Program_LastEdit = "2012-07-04"
    ' 1. 요일할인 및 행사자료 자료 복구에서 다운로드 받도록 수정
    ' 2. 사고품 저장하지 않고 출력되는 문제 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.91
'    Program_LastEdit = "2012-07-02"
    ' 1. 요일할인 및 행사자료가 처음일경우(오픈매장) 다운로드가 안되는 문제 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.91
'    Program_LastEdit = "2012-07-02"
    ' 1. 요일할인 및 행사자료가 처음일경우(오픈매장) 다운로드가 안되는 문제 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.90
'    Program_LastEdit = "2012-06-29"
    ' 1. 크렌즈 겔러리 자연수 2012-06-29일 적용 하도록 재수정
    ' 2. 최초 요일자료 수신시 오류나는 문제 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.88
'    Program_LastEdit = "2012-06-29"
    ' 1. 크렌즈 겔러리 자연수 2012-07-01일 적용 하도록 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.87
'    Program_LastEdit = "2012-06-27"
    ' 1. 할인관련 내용 추가 수정
    '    - 할증하여 받은 금액은 할인 계산에서 할증 금액을 기준으로 적용
    '    - 한벌 기능이 정화기 적용되지 않는 문제 수정
    '    - 할인 금액이 할증과 차감되어 나오는 문제 수정
    
'[    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.86
 '   Program_LastEdit = "2012-06-25"
 '   ' 1. 일일마감에서 이전내역이 초기화가 안되는 문제 수정(세탁,반품환불 콤보박스, 판매취소,택분실)


'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.85
'    Program_LastEdit = "2012-06-25"
    ' 1. 할인 현황이 있을 경우에만 출력 되도록 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.83
'    Program_LastEdit = "2012-06-19"
    ' 1. 판매취소후 택번호 변경후 다시 같은택을 판매취소할 경우 처음것으로 처리되는 문제 수정
    '    (사용 마일리지가 있을 경우 금액차이로 마일리지가 잘 정산이 안되는 문제가 있어서 해결함)
    ' 2. 현금 영수증 판매 취소시 취소 사유를 등록 하여 처리되도록 수정 (2012-07-01일 부터 적용)
    ' 3. 지점입고에서 폼이 액티브될때 다시 조회 하도록 수정
    ' 4. 1024(크렌즈) 자연수 기능 추가, 인쇄할 경우 한자가 지원이 안되서 water로 출력
    ' 5. 보관증에 정상금액및 할인금액 출력되도록 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.82
'    Program_LastEdit = "2012-05-30"
    ' 1. 미수금 수정이 있을 경우 해당 내용만 나오는 문제 수정
    ' 2. 공지 사항이 있을 경우 해당 공지사항을 확인하지 않을 경우 로그인이 안되도록 수정
    ' 3. 2012-06-01일 부터 공지사항 수신여부를 확인한다.
    

'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.80
'    Program_LastEdit = "2012-05-25"
    ' 1. 현금영수증의 거래자구분출력 내용 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.79
'    Program_LastEdit = "2012-05-18"
    ' 1. 문자메시지 관련하여 작업중 다른 화면으로 전환후 다시 돌아온경우 선택수량및 확인 수량이
    '    모두 0으로 표시되는 문제 수정
    
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.78
'    Program_LastEdit = "2012-05-16"
    ' 1. 승인취소된 내용이 전달 되지 않는 문제 수정(카드,현금영수증)
    ' 2. 판매 취소시 마일리지가 정확히 환원되지 않는 문제 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.77
'    Program_LastEdit = "2012-05-07"
    ' 1. 본사미출고 현황에서 수동입고 잡을 수 있도록 수정
    ' 2. SMS 정보 연결 정보 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.76
'    Program_LastEdit = "2012-05-07"
    ' 1. 당일 출고분만 출고 취소가 가능 하도록 수정
    ' 2. DB관련 SP 추가(복구및 용량 줄이기)
    ' 3. 출고에서 미불,완불 클릭하여 수정하도록 수정
    ' 4. 당일만 상표가 가능하다는 수정 메시지 처리
    ' 5. 당일 오전에 마감시 확인 메시지 출력하도록 수정
    ' 6. 접수 금액이 100만원이 넘을 경우 출력오류가 나는 문제 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.75
'    Program_LastEdit = "2012-04-27"
    ' 1. 미수금 수금시 출고된것부터 완불 처리로 변경
    ' 2. 택번호 재 사용시 최근 5일의 택번호를 확인하여 번호를 증가 하도록 수정
    ' 3. 자료 가저오기에서 미수금 조정 내역 부분 수정
    ' 4. 아동복을 한벌 기능으로 추가 접수시 금액 적용 되어 나오도록 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.74
'    Program_LastEdit = "2012-04-27"
    ' 1. 지점입고 속도 개선
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.73
'    Program_LastEdit = "2012-04-19"
    ' 1. 일일판매 현황 출력 부분 수정
    ' 2. 일일마감을 위한 테이블 인덱스 추가
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.72
'    Program_LastEdit = "2012-04-19"
    ' 1. 사고품의 내용을 공유하여 다른 매장에서 접수받은 내역도 접수에서 처리 되도록 수정
    ' 2. 로그인시 사고 자료 자동 다운로드 되도록 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.71
'    Program_LastEdit = "2012-04-18"
    ' 1. 사고품 접수에서 최초 접수만 문자를 보내도록 수정
    ' 2. 자료 생성시 이전지사 자료 수신 가능 하도록 수정
    
    'Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.70
    'Program_LastEdit = "2012-04-13"
    ' 1. 크랜즈겔러리 첫거래 고객일 경우 10%할인을 사용자가 적용 하도록 수정
    ' 2. 접수증 출력시 일부품목이 가격이 안나오는 문제 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.69
'    Program_LastEdit = "2012-04-05"
    ' 1. 세탁환불, 반품환불, 판매 취소일 경우 사용가능 마일리지로 넘어간 경우 삭제되지 않는 문제 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.68
'    Program_LastEdit = "2012-03-03"
    ' 1. 기본가격, 할인가격, 요일할인의 품목 순서를 가맹점의 기본가격으로 나타나도록 수정
    ' 2. 본사 미출고 현황 출력 내용 조정
    ' 3. 사인패드와 프린터 포트가 같아도 동작 되도록 설정
    ' 4. 미수금 수금시 미불->완불로 수정되도록 수정
    ' 5. 2곳 이상에서 접수를 받을 수 있도록 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.67
'    Program_LastEdit = "2012-03-03"
    ' 1. 가격자료 수신할 경우 미 수신 자료만 다시 수신할 수 있도록 변경
    
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.66
'    Program_LastEdit = "2012-03-03"
    ' 1. 사인패드 인식 문제 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.65
'    Program_LastEdit = "2012-03-02"
    ' 1. 일일마감에서 미수금이 반품환불, 세탁환불이 있을 경우 잘못 나오는 문제 수정
    ' 2. 보관증을 고객용이 먼저 나오도록 수정
    ' 3. 영수증 프린터 및 사인패드 찾기 기능 추가
    ' 4. 카드 승인및 현금 영수증 승인시 장비 컨트롤 변경
    ' 5. 기타 내용 수정
    
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.64
'    Program_LastEdit = "2012-02-20"
    ' 1. 사고품 가맹점,지사, 본사 입력 내용 많이 입력 되도록 수정
    ' 2. 일일매출현황이 미리 조회된 내용이 출력되어 일일 마감 현황이 맞지 않는 문제 수정
    ' 3. 신용카드 현황 및 현금 영수증 현황에 합계 금액 출력
    
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.63
'    Program_LastEdit = "2012-02-17"
    ' 1. 마감할 경우 마감 시간도 저장 하도록 수정
    ' 2. LAUNDRYXXXX 으로 접속 하도록 변경
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.62
'    Program_LastEdit = "2012-01-04"
    ' 1. 삭제 마일리지 적용 되도록 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.61
'    Program_LastEdit = "2012-01-04"
    ' 1. 이마트 메시지 수정 부분 오류
    ' 2. 일일마감시 Laundry1000에 한번만 저장되도록 수정 나머지는 뷰터이블이라 해당 지사로는
    '    저장하지 않아도 됨
    ' 3. 삭제 마일리지 적용 되도록 수정
    ' 4. 문자 메시지 발송 하기전에 서버 연결 시도하도록 수정
    

'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.60
'    Program_LastEdit = "2012-01-04"
'    ' 1. 마이그레이션 일자를 12-12-31일 까지로 변경
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.58
'    Program_LastEdit = "2011-12-26"
    ' 1. 사고 발생시 SMS 문자 전송 하도록 수정


'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.57
'    Program_LastEdit = "2011-12-09"
    ' 1. 일일마감 미수금 회수 금액 오류 수정
    ' 2. 마일리지를 사용하다가 미사용으로 변경하였을 경우 사용가능 마일리지와 누적마일리지가 잘못 ㅈ
    '    저장되는 문제 수정
    ' 3. 행사, 요일할인의 순서를 기본 가격의 순서로 설정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.56
'    Program_LastEdit = "2011-12-02"
    ' 1. 일일마감 카드 금액 다시 조정 (세탁환불, 반품환불 내역 조정)
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.54,55
'    Program_LastEdit = "2011-11-10"
    ' 1. 일일마감 카드 금액 다시 조정, 매출현황 포함
    ' 2. 상표 입력시 마지막 삭제 되는 문제
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.53
'    Program_LastEdit = "2011-11-08"
    ' 1. SMS 조회에서 출고된 내역 제외 (반품환불, 세탁환불 내용 취소)
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.52
'    Program_LastEdit = "2011-11-08"
    ' 1. 위치 조정
    ' 2. DB 생성 파일 다운로드 한다.
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.50
'    Program_LastEdit = "2011-11-02"
    ' 1. 행사금액 -> 요일할인 -> 정상 가격 순으로 처리되도록 변경
    ' 2. 일일마감, 환경설정 화면 위치 변경
    ' 3. 지점입고 속도 향상  IDX_지점입고 인덱스 설정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.49
'    Program_LastEdit = "2011-11-02"
    ' 1. 일일마감 위치 수정
    ' 2. 스프레드 내역 조정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.48
'    Program_LastEdit = "2011-11-01"
    ' 1. 일일마감 수정(마일리지 ㅡㅡ)
    ' 2. 마감화면 단말기로 출력 가능하도록 수정
    ' 3. 공지화면 제목 출력
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.47
'    Program_LastEdit = "2011-11-01"
    ' 1. 출고 화면에서 상표 입력 내용 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.46
'    Program_LastEdit = "2011-10-31"
    ' 1. 출고 화면에서 상표 입력 내용 수정
    ' 2. 일일매출 현황 내용 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.45
'    Program_LastEdit = "2011-10-31"
    ' 1. 출고 화면에서 상표 입력 내용 수정
    ' 2. 마감 화면에서 환불 금액 별도 계산 처리
    ' 3. 일일마감에서 마일리지 문제 수정
    
    
    'Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.44
    'Program_LastEdit = "2011-10-27"
    ' 1. 출고 화면에서 고객 주소로 검색 되도록 수정
    ' 2. 접수/출고 화면에서 고객 조회가 안되는 경우 수정
    ' 3. 접수/출고 고객 이용실적 출력 건수 조정
    ' 4. 무늬에서 기타 추가
    
    
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.43
'    Program_LastEdit = "2011-10-27"
    ' 1. 일일매출현항 같은 고객명은 출력 안하고 라인을 출력 시켜줌
    ' 2. 일일매출현항 금액 부분 대폭 수정
    ' 3. 특정 카드 결제시 카드 자리수 때문에 프로그램이 날라가는 버그 수정
    '    Dinus카드는 14자리, Amex카드는 9/15자리
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.42
'    Program_LastEdit = "2011-10-26"
    ' 1. 이마트 행사 관련 내용 수정
    ' 2. 월간매출 현황 마일리지 오류 부분 수정
    ' 3. 일일마감에서 카드 판매취소가 있을 경우 미수금 처리가 잘못 되는 문제 해결
    ' 4. 일일판매현황 양식 조정


'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.39   '40, 41
'    Program_LastEdit = "2011-10-20"
    ' 1. 일일매출에서 카드 수량및 금액이 잘못 계산되는 문제 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.38
'    Program_LastEdit = "2011-10-11"
    ' 1. 접수에서 택번호가 없을 경우 저장되지 않도록 수정
    ' 2. 오점 이미지 관련 내용수정
    '    - 택번호가 없을 경우 저장되지 않도록 수정
    ' 3. 지점 입고에서 처리가 완료 되면 연결을 종료 한다.



'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.38
'    Program_LastEdit = "2011-10-11"
    ' 1. 마감화면및 일일매출 화면에서 카드금액이 잘못 표시되는 문제 수정
    ' 2. 판매리스트 택번호 정렬 수정
    ' 3. DBUpdate.exe및 AidSupport.exe 파일이 없을 경우 무조건 다운로드 하도록 수정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.37
'    Program_LastEdit = "2011-10-01"
    ' 1. 결제 항목에서 카드 영수증이 출력 되지 않는 문제 수정
    
'    1.0.36
'    Program_LastEdit = "2011-10-01"
    ' 1. 위치 조정
    
'    Program_Version = App.Major & "." & App.Minor & "." & App.Revision   ' 버전 : 1.0.35
'    Program_LastEdit = "2011-10-01"
    ' 1. 쿠폰 사용하도록 수정 (마감 내역, 판매 리스트 내역 적용)
    ' 2. 출력물 위치 수정
    ' 3. 자료 정리에서 미수금 정리 부분 미수금 수정 테이블에 삽입 하도록 수정
    ' 4. 원격 지원 다운로드 하도록 변경및 해당 파일이 있을 경우 해당 파일로 실행 하도록 수정
    ' 5. 판매리스트를 마감해야만 출력 되도록 수정
    
    
End Function
    
    
