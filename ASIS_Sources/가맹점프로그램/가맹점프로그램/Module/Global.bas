Attribute VB_Name = "Global"
Option Explicit

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

' ini File Control에 관한 API라이브러리
Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (dest As Any, ByVal numBytes As Long)

Public Printer_Gb As String ' 프린터 구분 0=도트 1= 잉크젯, 2= 레이저
Public Printer_BO_Gb As String ' 보관증 구분 0=이전 1= 신규
Public Connect_Gb As Boolean ' 전송   구분 0=모뎀 1= 인터넷

Public ADOCon  As ADODB.Connection
Public ADORs   As ADODB.Recordset
Public SUBRs   As ADODB.Recordset
Public Rs      As ADODB.Recordset
Public Query   As String

Public MyDB      As Database
Public m_DBPath  As String
Public tempTagNo As String  '택번호
Public TagNo     As Integer     '택번호
Public strCode   As String    '구분코드
Public intCur    As Integer    '현재커서의 위치
Public chkItem1  As Boolean  '자료유무확인
Public strTel1   As String    '전화1
Public strtel2   As String    '전화2

'Public Err As Error         '에러오브젝트

Public nDayCloseChk As Boolean  '전일 마감을 했는지의 여부를 확인한다.
Public strDayClose  As String   ' 전일 마감여부.
Public tempCol      As Integer   'form1에서 임시사용
Public tempRow      As Integer   'form1에서 임시사용
Public NewRowchk    As Boolean  '자료유무확인
Public chkDaySale   As Boolean   '목요세일여부
Public chkEventSale As Boolean   '행사여부
Public chkPassWord  As Boolean   ' 비빌번호 확인여부 ture 이면 확인한 상태
Public chkServicePassWord As String   ' 세탁 서비스의 여부를 결정한다. 한손님에 한해서
Public chkPricPassWord    As String   ' 금액 입력 비빌번호 확인하면 그 택번호가 저장된다.
Public chkPassInput       As String   ' 입력한 비밀번호
Public chkProgramMode     As String  ' 01 이면 서버모드로 작동 ( 입력 가능)

Public Const ServerMode = "1"              ' 서버

Public ChkInputKey  As Boolean  ' 입고 버튼을 누른경우 전화번호에 커서를 두기위해

Public chkinputflig As String  ' 현재 입력 위치 ' 입고,출고,조회

Public iniFile      As String    ' ini 파일
Public MailCheck    As Boolean

Public iComboList   As Integer

Public bMsgMode     As Boolean   ' 메시지출력 여부
Public strMessage   As String  ' 출력할 메시지

Public Const SET_TITLE_INPUT = "입고중 ..."
Public Const SET_TITLE_OUTPUT = "출고중 ..."
Public Const SET_TITLE_VIEW = "조회 ..."
Public Const SET_TITLE_EXIT = "종료"
Public Const M_COUPON_LANGTH = 8

' 20091008일 체인점 코드에서 지사 코드로 변경
Public Const M_COUPON_KLENZ_CODE = "1024"
'Public Const M_COUPON_KLENZ_CODE = "100273"

Public bUpdatePoing As Boolean

Type TYPE고객정보
    고객번호    As String
    성명        As String
    전화1       As String
    전화2       As String
    휴대폰      As String
    주소        As String
    미수금      As String
    전송구분    As String
    카드번호    As String
    전화번호    As String
    SMS전송여부 As String
    등록일자    As String
End Type

Type TYPE대리점정보
    MasterCode  As String
    StoreCode   As String
    StoreName   As String
    StartDate   As String
    대리점번호 As String
    대리점색상 As String
    대리점명 As String
    수선 As String
    할인시작일 As String
    할인종료일 As String
    일수 As String
    비율 As String
    전화1 As String
    전화2 As String
    전화매장 As String
    전화SMS As String
    전화번호 As String
    목요세일 As String
    수선마진 As String
    운동화마진 As String
    가죽무스탕마진 As String
    카페트마진 As String
    외주운동화마진 As String
    프린터 As String
    일수2 As String
    마일리지여부  As String
    마일리지증가구분    As String
    지정할인비율  As String
    지정할인여부    As String
    삼성카드할인여부    As String
    삼성카드할인비율    As String
    특정할인비율  As String
    특정할인여부    As String
    고가세탁비율    As String
    세탁비환불여부  As String
    고객전화번호모두출력 As String
    SMS_EMART       As String
    
End Type

Type TYPE일일마감정보
    일자 As String
    총점수 As String
    반품수량 As String
    재세탁수량 As String
    수선수량 As String
    총매출액 As String
    본사금액 As String
    대리점금액 As String
    수선금액 As String
    판매구분 As String
    시작택 As String
    종료택 As String
    마감여부 As String
    전송여부 As String
End Type

Public 고객정보     As TYPE고객정보
Public 대리점정보   As TYPE대리점정보
Public 일일마감정보 As TYPE일일마감정보
Public 고객수정     As TYPE고객정보

'컨트롤 상태 열거형
Public Enum CommandFiles
    Backup      ' DB백업
    DBSend      ' DB를 본사에 보넨다
    Restore     ' DB를 복구한다.
    PGDown      ' 프로그램 다운로드
End Enum

Public Type monRecord
    strDay   As String * 2
    intCount As Integer
    dblTptal As Double
End Type

Type chkbill
    strchkdate(100)        As String
    strchkTno(100)         As String
    strchkItem(100)        As String
    lngMoney(100)          As Long
    lngchkRejectmoney(100) As Long
End Type

Type TYPE마일리지정보
    검색여부        As Boolean
    총사용금액      As Double
    잔액            As Double
    최종발생금액    As Double
    발생총누계      As Double
    사용누계        As Double
    미반환마일리지  As Double
End Type

Public userMileage  As TYPE마일리지정보
Public Const NextMileage = 100000   ' 다음 마일리지 발생금액

Public m_CommandTimeOut    As Long
Public M_CompnyMasterName  As String
Public m_SMS_EMART_PASS    As Boolean

'====================================================================================================
' Procedure : ConnectMasterCheck
' DateTime  : 2008-04-15 04:13
' Author    : pds2004
' Last Edit :
' Return Val:
'----------------------------------------------------------------------------------------------------
' 설     명 : 최초 신규 매장코드가 없을 경우 본사에 등록되어있는 내용을 가저온다.
'====================================================================================================
Public Function ConnectMasterCheck(MyHost As ADODB.Connection) As Boolean
    On Error GoTo ConnectMasterCheck_Error
    
    Dim HostConn As String
    
    HostConn = ""
    HostConn = HostConn & "Provider=SQLOLEDB.1;"
    HostConn = HostConn & "Persist Security Info=False;"
    HostConn = HostConn & "User ID=sa;"
    HostConn = HostConn & "Password=;"
    HostConn = HostConn & "Initial Catalog=Laundry;"
    HostConn = HostConn & "Data Source=store.clean-aid.co.kr,8657"
    
    Set MyHost = Nothing
    Set MyHost = New ADODB.Connection

    If MyHost.State = adStateOpen Then MyHost.Close
    
    MyHost.CommandTimeout = 30
    MyHost.Open HostConn

    ConnectMasterCheck = True
    
    On Error GoTo 0
    
    Exit Function

ConnectMasterCheck_Error:

    ConnectMasterCheck = False

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ConnectMasterCheck of Module Global"
End Function

Public Function SendStoreDefaultInfo(sOldDate As String, sOldMstCode As String, sOldTag As String, sOldStoreCode As String, sOldStoreName As String) As Boolean
    Dim sValue(12)  As String
    Dim MyHost      As ADODB.Connection
    
    On Error GoTo SendStoreDefaultInfo_Error

    If Fb대리점정보 = "Error" Then
        MsgBox "대리점 정보가 올바르지 않습니다. 확인후 저장하여 주십시요.", vbInformation, "확인"
        
        frmINIT.Show 1
        End
    End If
    
    If Trim(대리점정보.StoreCode) = "" Then
        MsgBox "대리점 정보가 올바르지 않습니다.", vbCritical, "경고"
        Exit Function
    End If
    
    If ConnectMasterCheck(MyHost) = True Then
        MyHost.BeginTrans
        
        '-------------------------------------------------------------------------------------------
        ' 기존 자료에 종료일자 등록
        '-------------------------------------------------------------------------------------------
        Query = "UPDATE MASTER_TAG_TBL SET END_DT = '" & Format(DateAdd("d", -1, Date), "yyyyMMdd") & "'"
        Query = Query & " WHERE Store_CD  = '" & 대리점정보.StoreCode & "'"
        Query = Query & "   AND END_DT    = '20991231' "
        Query = Query & "   AND START_DT <> '" & 대리점정보.StartDate & "'"
        MyHost.Execute Query
    
        '-------------------------------------------------------------------------------------------
        '
        '-------------------------------------------------------------------------------------------
        Query = "SELECT STORE_CD "
        Query = Query & " FROM MASTER_TAG_TBL "
        Query = Query & " WHERE Store_CD = '" & 대리점정보.StoreCode & "' "
        Query = Query & "   AND START_DT = '" & 대리점정보.StartDate & "' "
        Set SUBRs = New ADODB.Recordset
        SUBRs.Open Query, MyHost, adOpenStatic, adLockOptimistic
        
        If SUBRs.EOF = True Then
            Query = "INSERT INTO  MASTER_TAG_TBL (STORE_CD, START_DT, MASTER_CD, TAG_NB, END_DT) "
            Query = Query & " VALUES ('" & 대리점정보.StoreCode & "', '" & 대리점정보.StartDate & "', "
            Query = Query & " '" & 대리점정보.MasterCode & "', '" & 대리점정보.대리점번호 & "','20991231')"
            MyHost.Execute Query
        Else
            Query = "UPDATE MASTER_TAG_TBL "
            Query = Query & "SET TAG_NB =  '" & 대리점정보.대리점번호 & "', "
            Query = Query & "    END_DT = '20991231', "
            Query = Query & "    MASTER_CD = '" & 대리점정보.MasterCode & "'  "
            Query = Query & " WHERE Store_CD = '" & 대리점정보.StoreCode & "' "
            Query = Query & "   AND START_DT = '" & 대리점정보.StartDate & "' "
            MyHost.Execute Query
        End If
        
        '-------------------------------------------------------------------------------------------
        '
        '-------------------------------------------------------------------------------------------
        Query = "UPDATE STORE_TBL "
        Query = Query & "SET TAG_NB = '" & 대리점정보.대리점번호 & "', "
        Query = Query & "    MASTER_CD = '" & 대리점정보.MasterCode & "',  "
        Query = Query & "    START_DT = '" & 대리점정보.StartDate & "' "
        Query = Query & "WHERE Store_CD = '" & 대리점정보.StoreCode & "' "
        MyHost.Execute Query
    
        
        ' 변경 여부를 문자로 전송한다.
        ' 전송, 메시지타입, 수신번호, 발신번호, 메시지, 지사코드, 대리점코드, 고객코드, 고객성명, 참고5, 참고6
        sValue(0) = "1"
        sValue(1) = "0"
        sValue(2) = "010-9004-4523"
        sValue(3) = "031-522-2025"
        sValue(4) = "[정보변경] " & vbNewLine
        sValue(4) = sValue(4) & Hangul_Mid(대리점정보.대리점명, 1, 12) & vbNewLine
        sValue(4) = sValue(4) & sOldDate & "->" & Replace(대리점정보.StartDate, "-", "") & vbNewLine
        sValue(4) = sValue(4) & sOldStoreCode & "->" & 대리점정보.StoreCode & vbNewLine
        sValue(4) = sValue(4) & sOldMstCode & "->" & 대리점정보.MasterCode & vbNewLine
        sValue(4) = sValue(4) & sOldTag & "->" & 대리점정보.대리점번호 & vbNewLine
        sValue(5) = "9999"
        sValue(6) = "999"
        sValue(7) = " "
        sValue(8) = "Laundry"
        sValue(9) = "999999"
        sValue(10) = "2"
        
        '-------------------------------------------------------------------------------------------
        '
        '-------------------------------------------------------------------------------------------
        Query = "EXEC PRO_SMS_SEND "
        Query = Query & "'" & sValue(0) & "', "
        Query = Query & "'" & sValue(1) & "', "
        Query = Query & "'" & sValue(2) & "', "
        Query = Query & "'" & sValue(3) & "', "
        Query = Query & "'" & sValue(4) & "', "
        Query = Query & "'" & sValue(5) & "', "
        Query = Query & "'" & sValue(6) & "', "
        Query = Query & "'" & sValue(7) & "', "
        Query = Query & "'" & sValue(8) & "', "
        Query = Query & "'" & sValue(9) & "', "
        Query = Query & "'" & sValue(10) & "' "
        
        If 대리점정보.MasterCode <> "9999" Then
            MyHost.Execute Query
        End If
                
        ' 정보 변경 히스토리
        
        sValue(0) = "1"
        sValue(1) = sOldDate
        sValue(2) = sOldMstCode
        sValue(3) = sOldTag
        sValue(4) = sOldStoreCode
        sValue(5) = sOldStoreName
        sValue(6) = Format(대리점정보.StartDate, "@@@@-@@-@@")
        sValue(7) = 대리점정보.MasterCode
        sValue(8) = 대리점정보.대리점번호
        sValue(9) = 대리점정보.StoreCode
        sValue(10) = 대리점정보.StoreName
        sValue(11) = "LAUNDRY"
        sValue(12) = Now
        
        '-------------------------------------------------------------------------------------------
        '
        '-------------------------------------------------------------------------------------------
        Query = "EXEC PRO_STORE_CHANGE "
        Query = Query & "'" & sValue(0) & "', "
        Query = Query & "'" & sValue(1) & "', "
        Query = Query & "'" & sValue(2) & "', "
        Query = Query & "'" & sValue(3) & "', "
        Query = Query & "'" & sValue(4) & "', "
        Query = Query & "'" & sValue(5) & "', "
        Query = Query & "'" & sValue(6) & "', "
        Query = Query & "'" & sValue(7) & "', "
        Query = Query & "'" & sValue(8) & "', "
        Query = Query & "'" & sValue(9) & "', "
        Query = Query & "'" & sValue(10) & "', "
        Query = Query & "'" & sValue(11) & "', "
        Query = Query & "'" & sValue(12) & "' "
        MyHost.Execute Query
    
        MyHost.CommitTrans
    End If

    SUBRs.Close
    Set SUBRs = Nothing
    
    Set MyHost = Nothing

    On Error GoTo 0
    
    Exit Function

SendStoreDefaultInfo_Error:
    MyHost.RollbackTrans
    
    SendStoreDefaultInfo = False
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SendStoreDefaultInfo of Module Global"
    Resume
End Function

Public Function GetMasterStoreFromToDate(ByVal sDate As String, ByRef sMaster As String, ByRef sStore As String) As Boolean
    Dim MyHost  As ADODB.Connection

    GetMasterStoreFromToDate = False
    
    If ConnectMasterCheck(MyHost) = True Then
        '---------------------------------------------------------------------------------------
        ' 해당 기간은 무조건 1개여야 한다. 2개 이상일 경우 첫번째것으로 처리한다.
        '---------------------------------------------------------------------------------------
        Query = "SELECT TOP 1 *  "
        Query = Query & " FROM MASTER_TAG_TBL "
        Query = Query & " WHERE Store_CD = '" & 대리점정보.StoreCode & "' "
        Query = Query & "   AND '" & sDate & "' BETWEEN START_DT AND ISNULL(END_DT,'2099-12-31') "
        Query = Query & " ORDER BY START_DT "
        Set SUBRs = New ADODB.Recordset
        SUBRs.Open Query, MyHost, adOpenStatic, adLockOptimistic
        
        'If SUBRs.State = adStateOpen Then SUBRs.Close
        'SUBRs.CursorLocation = adUseClient
        'SUBRs.Open Query, MyHost, adOpenStatic, adLockBatchOptimistic, adCmdText

        If Not SUBRs.EOF = True Then
            sMaster = SUBRs.Fields("MASTER_CD") & ""
            sStore = SUBRs.Fields("TAG_NB") & ""
            
            GetMasterStoreFromToDate = True
        Else
            sMaster = ""
            sStore = ""
            
            GetMasterStoreFromToDate = False
        End If
    End If
    SUBRs.Close
    Set SUBRs = Nothing
    
    Set MyHost = Nothing
End Function


'====================================================================================================
' Procedure : SendNoSalesData
' DateTime  : 2008-04-15 20:29
' Author    : pds2004
' Last Edit :
' Return Val:
'----------------------------------------------------------------------------------------------------
' 설     명 : "N"로 설정되어 있는 모든 일자의 자료를 다시전송한다..
'====================================================================================================
Public Function SendNoSalesData() As Boolean
    Dim iDay    As Integer
    Dim iTempDay    As String
    
    Dim MyHost  As ADODB.Connection
    
    Dim sData(27)   As String
    Dim Scode(1)    As String
    
    Dim sSendDate() As String
    
    On Error GoTo SendNoSalesData_Error

    If Trim(대리점정보.StoreCode) = "000000" Then
        MsgBox "대리점 정보가 올바르지 않습니다.", vbCritical, "경고"
        Exit Function
    End If
    
    
    If ConnectMasterCheck(MyHost) = True Then
        'N로 설정되어 있는 모든 일자를 구한다.
        Query = "SELECT STORE_CD, SALE_DT "
        Query = Query & " FROM SALE_TBL "
        Query = Query & " WHERE Store_CD = '" & 대리점정보.StoreCode & "' "
        Query = Query & "   AND TRANS_CHK = 'N' "
        Set SUBRs = New ADODB.Recordset
        SUBRs.Open Query, MyHost, adOpenStatic, adLockOptimistic
        
        'If ADORset.State = adStateOpen Then ADORset.Close
        'ADORset.CursorLocation = adUseClient
        'ADORset.Open Query, MyHost, adOpenStatic, adLockBatchOptimistic, adCmdText
        
        iDay = 0
        
        Do While Not SUBRs.EOF
            ReDim Preserve sSendDate(iDay)
            
            sSendDate(iDay) = SUBRs.Fields("SALE_DT") & ""
            iDay = iDay + 1
            
            SUBRs.MoveNext
        Loop
        
        SUBRs.Close
        Set SUBRs = Nothing
        
        ' 없을 경우
        If iDay <= 0 Then
            Exit Function
        End If
        
        For iDay = 0 To UBound(sSendDate)
            '실제 전송일자
            iTempDay = Format(sSendDate(iDay), "@@@@-@@-@@")
            
            ' 전송 정보를 알아온다.
            Erase sData
            
            Call GetSendDataValues(iTempDay, sData)
            
            ' 해당 일자의 자료가 있을 경우만 작업한다.
            If sData(1) <> "" Then
                ' 해당 일자의 지사및 체인점 코드를 알아온다.
                Erase Scode
                
                If GetMasterStoreFromToDate(Format(iTempDay, "yyyyMMdd"), Scode(0), Scode(1)) = False Then
                    ' 본사에서 확인하여 처리할 수 있도록 하기위하여
                    ' 지사 정보가 없더라도 전송 처리한다.
                End If
                
                '---------------------------------------------------------------------------
                '
                '---------------------------------------------------------------------------
                Query = "SELECT STORE_CD, TRANS_CHK "
                Query = Query & " FROM SALE_TBL "
                Query = Query & " WHERE Store_CD = '" & 대리점정보.StoreCode & "' "
                Query = Query & "   AND SALE_DT  = '" & Format(iTempDay, "yyyyMMdd") & "' "
                Set SUBRs = New ADODB.Recordset
                SUBRs.Open Query, MyHost, adOpenStatic, adLockOptimistic
                                
                If SUBRs.EOF = True Then
                    Query = "INSERT INTO  SALE_TBL (SALE_DT, STORE_CD, "
                    Query = Query & " MASTER_CD , TAG_NB, "
                    Query = Query & " START_TAG, END_TAG, "
                    Query = Query & " SALE_AMT, MASTER_AMT, "
                    Query = Query & " STORE_AMT , IN_CNT, "
                    Query = Query & " JAES_CNT, SU_CNT, "
                    Query = Query & " BAN_CNT , OUT_CNT, "
                    Query = Query & " CARD_AMT, CARD_CNT, "
                    Query = Query & " SU_AMT, SALE_CHK, "
                    Query = Query & " CREATE_MIL, USE_MIL, DELETE_MIL,"
                    Query = Query & " TRANS_CHK , TRANS_DT)"
                    
                    Query = Query & " VALUES ('" & sData(1) & "', '" & 대리점정보.StoreCode & "', "
                    Query = Query & " '" & Scode(0) & "', '" & Scode(1) & "',"
                    Query = Query & " '" & sData(11) & "', '" & sData(12) & "', " 'START_TAG, END_TAG
                    Query = Query & " " & sData(6) & ", " & sData(7) & ", "       'SALE_AMT, MASTER_AMT
                    Query = Query & " " & sData(8) & ", " & sData(2) & ", "       'STORE_AMT, IN_CNT
                    Query = Query & " " & sData(4) & ", " & sData(5) & ", "       'JAES_CNT, SU_CNT
                    Query = Query & " " & sData(3) & ", " & "0" & ", "            'BAN_CNT, OUT_CNT
                    Query = Query & " " & sData(18) & ", " & sData(19) & ", "     'CARD_AMT, CARD_CNT
                    Query = Query & " " & sData(9) & ", '" & sData(10) & "', "     'SU_AMT, SALE_CHK
                    Query = Query & " " & sData(15) & ", " & sData(16) & ", " & sData(17) & ", "     'CREATE_MIL, USE_MIL, DELETE_MIL
                    Query = Query & " 'Y', '" & Format(Date, "yyyyMMdd") & "' "    'TRANS_CHK , TRANS_DT
                    Query = Query & " )"
                    MyHost.Execute Query
                Else
                    If SUBRs.Fields("TRANS_CHK") = "N" Then
                        Query = "UPDATE SALE_TBL "
                        Query = Query & " SET MASTER_CD = '" & Scode(0) & "', "
                        Query = Query & " TAG_NB =  '" & Scode(1) & "', "
                        Query = Query & " START_TAG =  '" & sData(11) & "', "
                        Query = Query & " END_TAG =  '" & sData(12) & "', "
                        Query = Query & " SALE_AMT = " & sData(6) & ", "
                        Query = Query & " MASTER_AMT = " & sData(7) & ", "
                        Query = Query & " STORE_AMT = " & sData(8) & ", "
                        Query = Query & " IN_CNT = " & sData(2) & ", "
                        Query = Query & " JAES_CNT = " & sData(4) & ", "
                        Query = Query & " SU_CNT = " & sData(5) & ", "
                        Query = Query & " BAN_CNT = " & sData(3) & ", "
                        Query = Query & " OUT_CNT = " & "0" & ", "
                        Query = Query & " CARD_AMT = " & sData(18) & ", "
                        Query = Query & " CARD_CNT = " & sData(19) & ", "
                        Query = Query & " SU_AMT = " & sData(9) & ", "
                        Query = Query & " SALE_CHK =  '" & sData(10) & "', "
                        Query = Query & " CREATE_MIL = " & sData(15) & ", "
                        Query = Query & " USE_MIL = " & sData(16) & ", "
                        Query = Query & " DELETE_MIL = " & sData(17) & ", "
                        Query = Query & " TRANS_CHK =  'Y', "
                        Query = Query & " TRANS_DT =  '" & Format(Date, "yyyyMMdd") & "'  "
                        Query = Query & " WHERE Store_CD = '" & 대리점정보.StoreCode & "' "
                        Query = Query & "   AND SALE_DT  = '" & Format(iTempDay, "yyyyMMdd") & "' "
                        MyHost.Execute Query
                    End If
                End If
            End If
        Next iDay
    End If

    SUBRs.Close
    Set SUBRs = Nothing
    
    Set MyHost = Nothing
    
    SendNoSalesData = True

    On Error GoTo 0
    
    Exit Function

SendNoSalesData_Error:
    SendNoSalesData = False
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SendNoSalesData of Module Global"
End Function

'====================================================================================================
' Procedure : SendSalesData
' DateTime  : 2008-04-15 20:29
' Author    : pds2004
' Last Edit :
' Return Val:
'----------------------------------------------------------------------------------------------------
' 설     명 : 기본적으로 최근 1주일 내용을 전송한다.
'====================================================================================================
Public Function SendSalesData(ByVal sSendDate As String, Optional iSendDay As Integer = 0) As Boolean
    Dim iDay    As Integer
    Dim iTempDay    As String
    
    Dim MyHost  As ADODB.Connection
    
    Dim sData(27)   As String
    Dim Scode(1)    As String
    
    On Error GoTo SendSalesData_Error

    If Trim(대리점정보.StoreCode) = "000000" Then
        MsgBox "대리점 정보가 올바르지 않습니다.", vbCritical, "경고"

        Exit Function
    End If
        
    If ConnectMasterCheck(MyHost) = True Then
        For iDay = 0 To iSendDay
            '실제 전송일자
            iTempDay = DateAdd("d", iDay, sSendDate)
            
            ' 전송 정보를 알아온다.
            Erase sData
            
            Call GetSendDataValues(iTempDay, sData)
            
            ' 해당 일자의 자료가 있을 경우만 작업한다.
            If sData(1) <> "" And sData(2) <> "0" Then
            
                ' 해당 일자의 지사및 체인점 코드를 알아온다.
                Erase Scode
                
                If GetMasterStoreFromToDate(Format(iTempDay, "yyyyMMdd"), Scode(0), Scode(1)) = False Then
                    ' 본사에서 확인하여 처리할 수 있도록 하기위하여
                    ' 지사 정보가 없도라도 전송 처리한다.
                End If
                
                '--------------------------------------------------------------------------
                '
                '--------------------------------------------------------------------------
                Query = "SELECT STORE_CD, TRANS_CHK "
                Query = Query & " FROM SALE_TBL "
                Query = Query & " WHERE Store_CD = '" & 대리점정보.StoreCode & "' "
                Query = Query & "   AND SALE_DT  = '" & Format(iTempDay, "yyyyMMdd") & "' "
                Set SUBRs = New ADODB.Recordset
                SUBRs.Open Query, MyHost, adOpenStatic, adLockOptimistic
                
                'If ADORset.State = adStateOpen Then ADORset.Close
                'ADORset.CursorLocation = adUseClient
                'ADORset.Open Query, MyHost, adOpenStatic, adLockBatchOptimistic, adCmdText
                
                If SUBRs.EOF = True Then
                    Query = "INSERT INTO  SALE_TBL (SALE_DT, STORE_CD, "
                    Query = Query & " MASTER_CD , TAG_NB, "
                    Query = Query & " START_TAG, END_TAG, "
                    Query = Query & " SALE_AMT, MASTER_AMT, "
                    Query = Query & " STORE_AMT , IN_CNT, "
                    Query = Query & " JAES_CNT, SU_CNT, "
                    Query = Query & " BAN_CNT , OUT_CNT, "
                    Query = Query & " CARD_AMT, CARD_CNT, "
                    Query = Query & " SU_AMT, SALE_CHK, "
                    Query = Query & " CREATE_MIL, USE_MIL, DELETE_MIL,"
                    Query = Query & " RunningCnt , RunningMoney, RunningPer, "
                    Query = Query & " SALERETURN_CNT, SALERETURN_AMT, "
                    Query = Query & " SAMSUNGCARDMEM_CNT , SAMSUNGCARD_CNT, SAMSUNGCARD_AMT, "
                    Query = Query & " TRANS_CHK , TRANS_DT) "
                    
                    Query = Query & " VALUES ('" & sData(1) & "', '" & 대리점정보.StoreCode & "', "
                    Query = Query & " '" & Scode(0) & "', '" & Scode(1) & "',"
                    Query = Query & " '" & sData(11) & "', '" & sData(12) & "', " 'START_TAG, END_TAG
                    Query = Query & " " & sData(6) & ", " & sData(7) & ", "       'SALE_AMT, MASTER_AMT
                    Query = Query & " " & sData(8) & ", " & sData(2) & ", "       'STORE_AMT, IN_CNT
                    Query = Query & " " & sData(4) & ", " & sData(5) & ", "       'JAES_CNT, SU_CNT
                    Query = Query & " " & sData(3) & ", " & "0" & ", "            'BAN_CNT, OUT_CNT
                    Query = Query & " " & sData(18) & ", " & sData(19) & ", "     'CARD_AMT, CARD_CNT
                    Query = Query & " " & sData(9) & ", '" & sData(10) & "', "     'SU_AMT, SALE_CHK
                    Query = Query & " " & sData(15) & ", " & sData(16) & ", " & sData(17) & ", "     'CREATE_MIL, USE_MIL, DELETE_MIL
                    Query = Query & " " & sData(20) & ", " & sData(21) & ", " & sData(22) & ", "     'RunningCnt , RunningMoney, RunningPer, "
                    Query = Query & " " & sData(23) & ", " & sData(24) & ", "      'RunningCnt , RunningMoney, RunningPer, "
                    Query = Query & " " & sData(25) & ", " & sData(26) & ", " & sData(27) & ", "     'RunningCnt , RunningMoney, RunningPer, "
                    Query = Query & " 'Y', '" & Format(Date, "yyyyMMdd") & "' "    'TRANS_CHK , TRANS_DT
                    Query = Query & " )"
                    MyHost.Execute Query
                
                Else
                    If SUBRs.Fields("TRANS_CHK") = "N" Then
                        Query = "UPDATE SALE_TBL "
                        Query = Query & " SET MASTER_CD = '" & Scode(0) & "', "
                        Query = Query & " TAG_NB =  '" & Scode(1) & "', "
                        Query = Query & " START_TAG =  '" & sData(11) & "', "
                        Query = Query & " END_TAG =  '" & sData(12) & "', "
                        Query = Query & " SALE_AMT = " & sData(6) & ", "
                        Query = Query & " MASTER_AMT = " & sData(7) & ", "
                        Query = Query & " STORE_AMT = " & sData(8) & ", "
                        Query = Query & " IN_CNT = " & sData(2) & ", "
                        Query = Query & " JAES_CNT = " & sData(4) & ", "
                        Query = Query & " SU_CNT = " & sData(5) & ", "
                        Query = Query & " BAN_CNT = " & sData(3) & ", "
                        Query = Query & " OUT_CNT = " & "0" & ", "
                        Query = Query & " CARD_AMT = " & sData(18) & ", "
                        Query = Query & " CARD_CNT = " & sData(19) & ", "
                        Query = Query & " SU_AMT = " & sData(9) & ", "
                        Query = Query & " SALE_CHK =  '" & sData(10) & "', "
                        Query = Query & " CREATE_MIL = " & sData(15) & ", "
                        Query = Query & " USE_MIL = " & sData(16) & ", "
                        Query = Query & " DELETE_MIL = " & sData(17) & ", "
                        Query = Query & " RunningCnt = " & sData(20) & ", "
                        Query = Query & " RunningMoney = " & sData(21) & ", "
                        Query = Query & " RunningPer = " & sData(22) & ", "
                        
                        Query = Query & " SALERETURN_CNT = " & sData(23) & ", "
                        Query = Query & " SALERETURN_AMT = " & sData(24) & ", "
                        Query = Query & " SAMSUNGCARDMEM_CNT = " & sData(25) & ", "
                        Query = Query & " SAMSUNGCARD_CNT = " & sData(26) & ", "
                        Query = Query & " SAMSUNGCARD_AMT = " & sData(27) & ", "
                        
                        Query = Query & " TRANS_CHK =  'Y', "
                        Query = Query & " TRANS_DT =  '" & Format(Date, "yyyyMMdd") & "'  "
                        Query = Query & " WHERE Store_CD = '" & 대리점정보.StoreCode & "' "
                        Query = Query & "   AND SALE_DT = '" & Format(iTempDay, "yyyyMMdd") & "' "
                        MyHost.Execute Query
                    End If
                End If
            End If
        Next iDay
    End If
    SUBRs.Close
    Set SUBRs = Nothing
    
    Set MyHost = Nothing
    
    SendSalesData = True

    On Error GoTo 0
    
    Exit Function

SendSalesData_Error:
    SendSalesData = False
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SendSalesData of Module Global"
End Function

Private Function GetSendDataValues(ByVal sSendData As String, ByRef sData() As String) As Boolean
    Dim dblRatio    As Double
    Dim dblMil      As Double
    Dim StoreMoney  As Double
    Dim MasterMoney As Double
        
    On Error GoTo Err_Rtn
    
    GetSendDataValues = False
        
    ' 대리점 코드를 Check한다.
    'Set rsTempTb = MyDB.OpenRecordset("SELECT * FROM 일일마감 WHERE 일자 = '" & Format(sSendData, "yyyyMMdd") & "' ")
    
    Query = "SELECT * FROM 일일마감 WHERE 일자 = '" & Format(sSendData, "yyyyMMdd") & "' "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    ' 대리점코드가 없으면 종료를 한다.
    If Not SUBRs.EOF Then
        dblMil = IIf(IsNull(SUBRs.Fields("사용마일리지")) = True, 0, SUBRs.Fields("사용마일리지"))

'        If Format(sSendData, "yyyyMMdd") >= "20080501" Then
'            ' 2009-08-27일
'            ' 이부분에서 오류가 있음
'            ' 본사및 대리점매출에서 해당 마일리지 비율을 제외한 금액을 전송하려고 하였으나
'            ' 비율에 0이 들어가는 오류로 본사 금액은 마일리지가 포함된금액이 그래로 전송이 되며
'            ' 체인점 금액에서 마일리지가 차감된 금액으로 전송하도록 되어 있음. ㅡㅡ
'            ' 임으로 바꿀수 없어서 본사 처리루틴에서 일괄 적용 하기로 처리함
'            MasterMoney = SUBRs.Fields("본사금액") - (dblMil * dblRatio)
'            StoreMoney = SUBRs.Fields("대리점금액") - (dblMil * (1 - dblRatio))
'        Else
'            MasterMoney = SUBRs.Fields("본사금액")
'            StoreMoney = SUBRs.Fields("대리점금액")
'        End If
    
        MasterMoney = SUBRs.Fields("본사금액")
        StoreMoney = SUBRs.Fields("대리점금액")
    
    
        sData(0) = 대리점정보.StoreCode & ""
        sData(1) = Trim(SUBRs!일자 & "")
        sData(2) = IIf(IsNumeric(SUBRs!총점수 & ""), SUBRs!총점수, "0")
        sData(3) = IIf(IsNumeric(SUBRs!반품수량 & ""), SUBRs!반품수량, "0")
        sData(4) = IIf(IsNumeric(SUBRs!재세탁수량 & ""), SUBRs!재세탁수량, "0")
        sData(5) = IIf(IsNumeric(SUBRs!수선수량 & ""), SUBRs!수선수량, "0")
        sData(6) = IIf(IsNumeric(SUBRs!총매출액 & ""), SUBRs!총매출액, "0")
        sData(7) = MasterMoney
        sData(8) = StoreMoney
        sData(9) = IIf(IsNumeric(SUBRs!수선금액 & ""), SUBRs!수선금액, "0")
        sData(10) = SUBRs!판매구분 & ""
        sData(11) = SUBRs!시작택 & ""
        sData(12) = SUBRs!종료택 & ""
        sData(13) = SUBRs!마감여부 & ""
        sData(14) = SUBRs!전송여부 & ""
        sData(15) = IIf(IsNumeric(SUBRs!발생마일리지 & ""), SUBRs!발생마일리지, "0")
        sData(16) = IIf(IsNumeric(SUBRs!사용마일리지 & ""), SUBRs!사용마일리지, "0")
        sData(17) = IIf(IsNumeric(SUBRs!삭제마일리지 & ""), SUBRs!삭제마일리지, "0")
        sData(18) = IIf(IsNumeric(SUBRs!카드금액 & ""), SUBRs!카드금액, "0")
        sData(19) = IIf(IsNumeric(SUBRs!카드건수 & ""), SUBRs!카드건수, "0")
        
        sData(20) = IIf(IsNumeric(SUBRs!운동화건수 & ""), SUBRs!운동화건수, "0")
        sData(21) = IIf(IsNumeric(SUBRs!운동화금액 & ""), SUBRs!운동화금액, "0")
        sData(22) = Val(대리점정보.외주운동화마진 & "")
        
        sData(23) = IIf(IsNumeric(SUBRs!세탁비환불건수 & ""), SUBRs!세탁비환불건수, "0")
        sData(24) = IIf(IsNumeric(SUBRs!세탁비환불금액 & ""), SUBRs!세탁비환불금액, "0")
        
        sData(25) = IIf(IsNumeric(SUBRs!삼성카드할인고객수 & ""), SUBRs!삼성카드할인고객수, "0")
        sData(26) = IIf(IsNumeric(SUBRs!삼성카드할인건수 & ""), SUBRs!삼성카드할인건수, "0")
        sData(27) = IIf(IsNumeric(SUBRs!삼성카드할인금액 & ""), SUBRs!삼성카드할인금액, "0")
        
        
        ' 마감되지 않았을 경우 전송하지 않는다.
        If sData(13) <> "Y" Then sData(2) = "0"
        
    Else
        sData(0) = 대리점정보.StoreCode & ""
        
        If 대리점정보.StartDate <= Format(sSendData, "yyyyMMdd") Then
            sData(1) = Format(sSendData, "yyyyMMdd")
        Else
            sData(1) = ""
        End If
        
        sData(2) = "0"
        sData(3) = "0"
        sData(4) = "0"
        sData(5) = "0"
        sData(6) = "0"
        sData(7) = "0"
        sData(8) = "0"
        sData(9) = "0"
        sData(10) = ""
        sData(11) = ""
        sData(12) = ""
        sData(13) = ""
        sData(14) = ""
        sData(15) = "0"
        sData(16) = "0"
        sData(17) = "0"
        sData(18) = "0"
        sData(19) = "0"
        sData(20) = "0"
        sData(21) = "0"
        sData(22) = "0"
    
        sData(23) = "0"
        sData(24) = "0"
        sData(25) = "0"
        sData(26) = "0"
        sData(27) = "0"
    End If
        
    SUBRs.Close
    Set SUBRs = Nothing
    
    GetSendDataValues = True
    
    Exit Function
    
Err_Rtn:
    GetSendDataValues = False
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetSendDataValues of Module Global"
End Function

Private Function GetSendDataCouponValues(ByVal sSendData As String, ByRef sData() As String) As Boolean
    Dim dblRatio    As Double
    Dim dblMil      As Double
    Dim StoreMoney  As Double
    Dim MasterMoney As Double
    
    
    On Error GoTo Err_Rtn
    
    GetSendDataCouponValues = False
    
    ' 대리점 코드를 Check한다.
    'Set SUBRs = MyDB.OpenRecordset("SELECT * FROM 쿠폰자료 WHERE 접수일자 = '" & Format(sSendData, "yyyyMMdd") & "' ")
            
    '----------------------------------------------------------------------------
    '
    '----------------------------------------------------------------------------
    Query = "SELECT * FROM 쿠폰자료"
    Query = Query & " WHERE 접수일자 = '" & Format(sSendData, "yyyyMMdd") & "' "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    ' 자료가 없으면 종료를 한다.
    If Not SUBRs.EOF Then
        sData(0) = 대리점정보.StoreCode & ""
        sData(1) = Trim(SUBRs!접수일자 & "")
        sData(2) = Trim(SUBRs!쿠폰번호 & "")
        sData(3) = IIf(IsNumeric(SUBRs!쿠폰단가 & ""), SUBRs!쿠폰단가, "0")
        sData(4) = IIf(IsNumeric(SUBRs!쿠폰금액 & ""), SUBRs!쿠폰금액, "0")
        sData(5) = Trim(SUBRs!고객번호 & "")
        sData(6) = Trim(SUBRs!고객이름 & "")
        sData(7) = IIf(IsNumeric(SUBRs!접수금액 & ""), SUBRs!접수금액, "0")
        sData(8) = Trim(SUBRs!택번호 & "")
    
    Else
        sData(0) = 대리점정보.StoreCode & ""
        
        If 대리점정보.StartDate <= Format(sSendData, "yyyyMMdd") Then
            sData(1) = Format(sSendData, "yyyyMMdd")
        Else
            sData(1) = ""
        End If
        
        sData(2) = ""
        sData(3) = "0"
        sData(4) = "0"
        sData(5) = ""
        sData(6) = ""
        sData(7) = "0"
        sData(8) = "0"
    End If
    
    SUBRs.Close
    Set SUBRs = Nothing
        
    GetSendDataCouponValues = True
    
    Exit Function
    
Err_Rtn:
    GetSendDataCouponValues = False
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetSendDataCouponValues of Module Global"
End Function


Public Function GetMailConvert(sData As String, bMode As String) As String
    Select Case bMode
        Case "READ"
            GetMailConvert = Replace(sData, "~&^", "'")
        
          '  sTemp = Replace(sTemp, vbNewLine, "")
          '  sTemp = Replace(sTemp, Chr(8), "")
        
        Case "SAVE"
            GetMailConvert = Replace(sData, "'", "~&^")
        
        Case Else
            GetMailConvert = sData
    End Select
End Function


'====================================================================================================
' Procedure : GetMailData
' DateTime  : 2008-04-15 20:29
' Author    : pds2004
' Last Edit :
' Return Val:
'----------------------------------------------------------------------------------------------------
' 설     명 : "N"로 설정되어 있는 모든 일자의 자료를 다시전송한다..
'====================================================================================================
Public Function GetMailData() As Boolean
    Dim MyHost  As ADODB.Connection
    Dim sTemp       As String
    
    On Error GoTo GetMailData_Error

    If Trim(대리점정보.StoreCode) = "000000" Then
        MsgBox "대리점 정보가 올바르지 않습니다.", vbCritical, "경고"

        Exit Function
    End If
    
    If ConnectMasterCheck(MyHost) = True Then
        'N로 설정되어 있는 모든 일자를 구한다.
        Query = "SELECT * FROM Mail_ALL"
        Query = Query & " WHERE SendChk = '1'"
        Query = Query & "   AND AgencyCode = '" & 대리점정보.StoreCode & "' "
        Set SUBRs = New ADODB.Recordset
        SUBRs.Open Query, MyHost, adOpenStatic, adLockOptimistic
        
        'If ADORset.State = adStateOpen Then ADORset.Close
        '
        'ADORset.CursorLocation = adUseClient
        'ADORset.Open Query, MyHost, adOpenStatic, adLockBatchOptimistic, adCmdText
        
        Do While Not SUBRs.EOF
            Query = "SELECT * FROM 메일"
            Query = Query & " WHERE 메일일자 = '" & SUBRs.Fields("MailDate") & "' "
            Query = Query & "   AND 메일번호 = " & SUBRs.Fields("MailNo") & " "
            Query = Query & "   AND 송수신구분 = '2' "
            Set Rs = New ADODB.Recordset
            Rs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
            
            sTemp = GetMailConvert(SUBRs.Fields("MailDesc"), "SAVE")
                        
            'Set rsTempTb = MyDB.OpenRecordset(Query)
            
            If Not Rs.EOF Then
                Query = "UPDATE 메일 SET 메일내역 = '" & sTemp & "', "
                Query = Query & " 조회시작일 = '" & SUBRs.Fields("MailFrom") & "', "
                Query = Query & " 조회종료일 = '" & SUBRs.Fields("MailTo") & "',  "
                Query = Query & " 수신여부 = 'N',  "
                Query = Query & " 수신일자 = ' '  "
                Query = Query & " WHERE 메일일자 = '" & SUBRs.Fields("MailDate") & "' "
                Query = Query & "   AND 메일번호 = " & SUBRs.Fields("MailNo") & " "
                Query = Query & "   AND 송수신구분 = '2' "
                ADOCon.Execute Query
            Else
                Query = "INSERT INTO  메일(송수신구분,메일일자,메일번호,메일내역,조회시작일,조회종료일,수신여부,수신일자,전송구분) "
                Query = Query & " VALUES('" & SUBRs.Fields("MailType") & "', '" & SUBRs.Fields("MailDate") & "', "
                Query = Query & " " & SUBRs.Fields("MailNo") & ", '" & sTemp & "', '" & SUBRs.Fields("MailFrom") & "', "
                Query = Query & " '" & SUBRs.Fields("MailTo") & "', 'N',' ','N')  "
                ADOCon.Execute Query
            End If
            Rs.Close
            Set Rs = Nothing
            
            '---------------------------------------------------------------------
            '
            '---------------------------------------------------------------------
            Query = "UPDATE Mail_ALL SET SendChk = '2'"
            Query = Query & " WHERE SendChk    = '1'"
            Query = Query & "   AND AgencyCode = '" & 대리점정보.StoreCode & "' "
            Query = Query & "   AND MailNo     = " & SUBRs.Fields("MailNo") & " "
            MyHost.Execute Query
            
            SUBRs.MoveNext
        Loop
        
        SUBRs.Close
        Set SUBRs = Nothing
    End If
    
    Set MyHost = Nothing
    
    GetMailData = True

    On Error GoTo 0
    
    Exit Function

GetMailData_Error:
    GetMailData = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetMailData of Module Global"
End Function

Public Function dayTagchk(ByVal txtTagChk As String) As Boolean
    ' 전달된 택번호가 입출고에 있을 경우 False를 리턴한다.
    ' 판매 취소일 경우는 False로 리턴한다.
    Dim strDayChk As String
    Dim chkRow As Integer
    
    '입출고 table 에 자료가 중복된경우 체크
    strDayChk = txtTagChk
    
    If Not IsTagNum(strDayChk) Then
        MsgBox " 전달된 택번호가 올바르지 않습니다. ", "택번호 오류"
        Exit Function
    End If
    
    Query = "SELECT 번호 "
    Query = Query & "FROM 입출고 "
    Query = Query & "WHERE 번호 = '" & Trim(strDayChk) & "' "
    Query = Query & "AND   입고일 = '" & Format(Date, "yyyymmdd") & "' "
    Query = Query & "AND   ( 판매취소 <> 'Y'  OR 판매취소  IS NULL ) "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If SUBRs.RecordCount < 1 Then
        dayTagchk = True        ' 같은 번호가 없는경우
    Else
        dayTagchk = False       ' 같은 번호가 있는경우
    End If
    
    SUBRs.Close
    Set SUBRs = Nothing
End Function

Public Function tagChk(ByVal txttag As String) As Boolean
    '화면상에서 중복된 자료 체크

    Dim strTag As String
    Dim i As Long
    Dim lastRow1 As Integer
    Dim strTag2 As String
    
    strTag = Trim(txttag)
    
    If Not IsTagNum(strTag) Then
        MsgBox " 전달된 택번호가 올바르지 않습니다. ", "택번호 오류"
        tagChk = False
        Exit Function
    End If
    
    lastRow1 = GetSpreadLine(frm접수.sprGrid) - 1
    
    ' 마지막 라인이 2줄보다 작을 경우 입력한 내용은 1줄이므로 비교할 필요가 없다.
    If lastRow1 < 1 Then
        tagChk = True
        Exit Function
    End If
    For i = 1 To lastRow1
        strTag2 = GetSpreadText(frm접수.sprGrid, i, 2)
        If Len(strTag2) >= 1 Then
            If strTag = Trim(strTag2) Then
                tagChk = False
                Exit Function
            End If
        End If
    Next i
    
    tagChk = True
    
End Function

Public Function f_newRowChk()
' 색상이 입력 되어 있는지 검사한다.
' 입력시 True 리턴

    frm접수.sprGrid.Row = frm접수.sprGrid.ActiveRow
    frm접수.sprGrid.Col = 3
    
    If Len(frm접수.sprGrid.Text) > 1 Then
        NewRowchk = True
    Else
        NewRowchk = False
    End If
End Function

Public Sub dataDisplayForm2(ByVal strCode As String)
    Dim strBlank1 As String
    Dim strBlank2 As String
    Dim i As Integer
    
    strBlank1 = "확"
    
    '----------------------------------------------------------------------------
    Query = "SELECT 고객번호, 성명, 전화1, 전화2, 주소, 미수금 "
    Query = Query & " FROM 고객정보 "
    Query = Query & " WHERE 고객번호 = '" & strCode & "' "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If SUBRs.RecordCount < 1 Then
        SUBRs.Close
        Set SUBRs = Nothing
        
        MsgBox " 일치하는 회원이 없읍니다  다시입력요망", vbInformation, "확인"
        
        Exit Sub
    End If
    
    frm출고.txtTel(0).Text = SUBRs!전화1
    frm출고.txtTel(1).Text = SUBRs!전화2
    frm출고.txtCode.Text = SUBRs!고객번호
    frm출고.txtName.Text = SUBRs!성명
    frm출고.Text1 = SUBRs!주소
    frm출고.txtMisu.Value = Format(SUBRs!미수금, "###,##0")
    
    SUBRs.Close
    Set SUBRs = Nothing
    
    '----------------------------------------------------------------------------
    '
    '----------------------------------------------------------------------------
    Query = "SELECT 입고일, 품명, 번호, 색상, 내용, 금액, 상태, 본출, 상표, 확인, 세트Key, 세트구분 "
    Query = Query & "FROM 입출고 "
    Query = Query & "WHERE 고객번호 = '" & Trim(frm출고.txtCode.Text) & "' "
    Query = Query & "AND (확인 IS NULL OR 확인 <> '" & Trim(strBlank1) & "')"
    Query = Query & "AND (판매취소 IS NULL OR 판매취소 <> 'Y') " ' 2002/11/22 일 추가
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If SUBRs.EOF Then
        SUBRs.Close
        Set SUBRs = Nothing
    
        Exit Sub
    End If
    
    With frm출고.sprChul
        .MaxRows = 20
        .Row = 1
        
        Do Until SUBRs.EOF
            .Col = 1: .Value = Mid(SUBRs!입고일, 1, 4) & "-" & Mid(SUBRs!입고일, 5, 2) & "-" & Mid(SUBRs!입고일, 7, 2)
            
            .Col = 2
            
            If IsNull(SUBRs!본출) = True Then
                .Value = ""
            Else
                .Value = SUBRs!본출
            End If
            
            .Col = 3:  .Value = SUBRs!품명 & ""
            .Col = 4:  .Value = SUBRs!번호 & ""
            .Col = 5:  .Value = SUBRs!색상 & ""
            .Col = 6:  .Value = SUBRs!내용 & ""
            .Col = 7:  .Value = SUBRs!금액 & ""
            .Col = 8:  .Value = SUBRs!상태 & ""
            .Col = 9:  .Value = SUBRs!상표 & ""
            .Col = 10: .Value = SUBRs!확인 & ""
            
            .Col = 12: .Value = SUBRs!세트Key & ""
            .Col = 13: .Value = SUBRs!세트구분 & ""
            
            If .Row >= .MaxRows Then
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                
                .Action = ActionInsertRow
                .RowHeight(.MaxRows) = .RowHeight(1) ' 마지막 라인의 높이를 맞춘다.
            End If
            
            .Row = .Row + 1
            
            SUBRs.MoveNext
        Loop
    End With
    
    SUBRs.Close
    Set SUBRs = Nothing
End Sub

Public Function PrNumSet(Num As Variant, cnt As Integer)
    Dim Num1 As Double
    Dim Str As String
    
    Num1 = Val(Num)
    Str = "                   " & Format(Num1, "#,##0")
    PrNumSet = Right(Str, cnt)
End Function

Public Sub G_CashKeyUp(Text2 As Control, KeyCode As Integer)
    Dim iCurSelstart  As Integer
    Dim iTextEnd_flag As Integer
    Dim iTextLen1     As Integer
    Dim iTextLen2     As Integer
    Dim buff          As String
    
    iCurSelstart = Text2.SelStart
    
    If (KeyCode >= vbKey0 And KeyCode <= vbKey9 Or _
        KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9 Or _
        KeyCode = vbKeyBack Or KeyCode = vbKeyDelete) Then
        
        '정상적인 키입력과 수정하려는 키입력을 구분하기 위해 플래그를 설치
        '-----------------------------------------------------------------
        If iCurSelstart = LenB(Text2.Text) Then
            iTextEnd_flag = True
        Else
            iTextEnd_flag = False
        End If
        
        iTextLen1 = LenB(Text2.Text)                        ' 포맷형식으로 수정되기 전의 길이
        
        buff = Format(Text2.Text, "###############")
        
        If Left(buff, 1) = "," Then buff = Mid(buff, 2)     ' ",###"일경우 "###"로
        
        Text2.Text = Format(buff, "###,###,###,###,###")
        
        If iTextEnd_flag = True Then
            Text2.SelStart = LenB(Text2.Text)
        Else
            iTextLen2 = LenB(Text2.Text) '포맷형식으로 수정된 길이
            '길이차이가 1만큼난다는 것은 자릿수가 하나 더 늘었다는 것을 의미
            '---------------------------------------------------------------
            If iTextLen2 - iTextLen1 = 1 Then
                Text2.SelStart = iCurSelstart + 1
            Else
                If iCurSelstart >= 4 Then
                    Text2.SelStart = iCurSelstart + 1
                Else
                    Text2.SelStart = iCurSelstart
                End If
            End If
        End If
    End If
    
    Set Text2 = Nothing
End Sub

Public Function G_CashKeyPress(KeyAscii As Integer) As Integer
    If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyReturn Or _
        KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then
    Else
        KeyAscii = 0
    End If
    
    G_CashKeyPress = KeyAscii
End Function

Sub Tag_Load()
    Dim strDate As String
    
    '--------------------------------------------------------------
    '
    '--------------------------------------------------------------
    Query = "SELECT 할인시작일 "
    Query = Query & " FROM 대리점정보 "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If SUBRs.EOF Or SUBRs.BOF Then
        frmMain.TagNo.Caption = "0" + "-" + "001"
    ElseIf IsNull(SUBRs(0)) = True Then
        frmMain.TagNo.Caption = "0" + "-" + "001"
    ElseIf SUBRs!할인시작일 = "000000" Then
        frmMain.TagNo.Caption = "0" + "-" + "001"
    Else
        ' 신프로그램으로 변경시 앞으로 사용할 택번호가 출력되게 변경했기 때문에
        ' 기존 구프로그램은 사용한 택번호가 찍히므로 구 택번호에 판매정보가
        ' 신프로그램 처음 사용시 첫 접수자로 변경되는 것을 막기위해 추가.
        ' 2002.12.06
        
        ' 현재 접수일자를 확인한다.
        If DayCloseCheck(Format(Date, "yyyymmdd")) = True Then
            strDate = Format(DateAdd("d", 1, Date), "yyyymmdd")
        Else
            strDate = Format(Date, "yyyymmdd")
        End If
        
        SUBRs.Close
        Set SUBRs = Nothing
        
        '--------------------------------------------------------------
        '
        '--------------------------------------------------------------
        Query = ""
        Query = Query & "SELECT * FROM  입출고 "
        Query = Query & "         WHERE 입고일 = '" & strDate & "'"
        Query = Query & "         AND   번호   = '" & SUBRs(0) & "'"
        Query = Query & "         AND   판매취소 <> 'Y' "
        Set SUBRs = New ADODB.Recordset
        SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
        
        If Rs.RecordCount > 0 Then
            frmMain.TagNo.Caption = GetTagNum(SUBRs(0), "+")
            
            Call Tag_Save
        Else
            frmMain.TagNo.Caption = SUBRs(0)
        End If
    End If
    
    SUBRs.Close
    Set SUBRs = Nothing
End Sub

Sub Main()
    Dim strTemp As String

    ' 환경 설정 파일의 이름을 설정한다.
    iniFile = App.Path & "\Laundry.ini"
    
        ' 프로그램 버전을 설정한다.
    SetProgramVersion


    ' 프로그램 실행 모드를 확인한다.
    SetProgramMode
    
    chkProgramMode = GetIniStr("RUNMODE", "ProgramMode", "", iniFile)
    
    If chkProgramMode = "" Then chkProgramMode = ServerMode
    
    M_CompnyMasterName = Trim(GetIniStr("RUNMODE", "ProgramName", "", iniFile)) ' 수정일시 20090428
    
    If M_CompnyMasterName = "" Then M_CompnyMasterName = "(주)크린에이드"
    
    'DB의 위치를 설정
    If chkProgramMode = ServerMode Then
        m_DBPath = App.Path & "\DB\Laundry.mdb"
    Else
        m_DBPath = Replace(GetIniStr("RUNMODE", "DBPath", "", iniFile), vbNullChar, "")
        m_DBPath = m_DBPath & "\Laundry.mdb"
    End If
    
    ' 메시지 출력 여부 설정
    bMsgMode = False
    
    ' DB 연결
    If Not Db_Connect Then
        If chkProgramMode <> ServerMode Then
            MsgBox "클라이언트로 실행될경우 서버의 DB를 연결하셔야 합니다.", vbCritical, "DB 오류"
            
            If MsgBox("만일 [조회->출고모드전환] 버튼을 클릭한경우 다시 [입고모드]로 전환 하시겠습니까?" & vbLf & _
                        "[예]를 선택한 경우 프로그램을 다시 시작하십시요 ", vbInformation + vbYesNo, "확인") = vbYes Then
                Call SetIniStr("RUNMODE", "ProgramMode", "1", iniFile)
            End If
        End If
        
        End
    End If
    
    ' DB 컬럼 확인및 추가
    Call DatabaseCheck(False)
    
        
    'Directory 체크
    DirectoryCheck
    ' Files 체크
    strTemp = FilesCheck
    If Val(Mid(strTemp, 1, 1)) > 0 Then
        MsgBox "실행에 필요한 파일이 없어서 일부기능이 정상적으로 실행되지 않을수 있습니다" & Chr(13) & Chr(10) & _
            "오류파일:" & Mid(strTemp, 3), vbCritical, "확인"
    End If
    
    chkinputflig = "메뉴"
    
    
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+  기본 실행시
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If UCase(App.EXEName) = "LAUNDRY" Then
        ' 기존 파일이 있을 경우 삭제 한다.
        If Dir(App.Path & "\Laundry_up.exe") <> "" Then Kill App.Path & "\Laundry_up.exe"
        If Dir(App.Path & "\Laundry_run.exe") <> "" Then Kill App.Path & "\Laundry_run.exe"
        If Dir(App.Path & "\laundry_OLD.exe") <> "" Then Kill App.Path & "\laundry_OLD.exe"
        If Dir(App.Path & "\ok.ok") <> "" Then Kill App.Path & "\ok.ok"
    
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+  업그레이드시
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    ElseIf UCase(App.EXEName) = "LAUNDRYUP" Then
        
        Call ProgramUpgrade
'        ' 폴더를 확인한다.
'        If Dir(Trim(App.Path & "\CleanPrg"), vbDirectory) = "" Then
'           MkDir App.Path & "\CleanPrg"
'        End If
'        ' 기존 화일이 있을경우 지운다
'        FileCopy App.Path & "\laundry.exe", App.Path & "\CleanPrg\Laundry_OLD.exe"
'        DoEvents
'        If Dir(App.Path & "\laundry.exe") <> "" Then Kill App.Path & "\laundry.exe"
'        DoEvents
'        If Dir(App.Path & "\laundry_RUN.exe") <> "" Then FileCopy App.Path & "\laundry_RUN.exe", App.Path & "\Laundry.exe"
'        DoEvents
    
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+  프로그램 복원시
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    ElseIf UCase(App.EXEName) = "LAUNDRY_OLD" Then
        '  이전 프로그램으로 실행 했을경우 자신을 신프로그램으로 복사한다.
        On Error GoTo Err_Chk
        ' 기존 화일이 있을경우 지운다
        Delay (5)
        If Dir(App.Path & "\Laundry.exe") <> "" Then Kill App.Path & "\Laundry.exe"
        DoEvents
        FileCopy App.Path & "\laundry_OLD.exe", App.Path & "\laundry.exe"
        DoEvents
        On Error GoTo 0

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+  프로그램이 정해진 이름으로 시작하지 않을경우
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Else
        MsgBox "프로그램은 Laundry.exe 파일로 실행되어야 합니다.", vbInformation, "실행오류"
        End
    End If
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    ' 2008.11.05 서버이전 작업으로 인한 일괄 변경
    If Format(Date, "yyyyMMdd") <= "20081120" Then
        ' SQL DB 접속
        strTemp = "UPDATE 대리점정보 SET "
        strTemp = strTemp & " ServerIP = 'store.clean-aid.co.kr,8657', "
        strTemp = strTemp & " ServerDB = 'Laundry', "
        strTemp = strTemp & " ServerUser = 'sa', "
        strTemp = strTemp & " ServerPass = ' ' "
        ADOCon.Execute strTemp
        
        ' 파일 전송
        Call SetIniStr("Connect", "RemoteIP", "web.clean-aid.co.kr", iniFile)
    End If
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    Load frmBase
    frmBase.Top = -50
    frmBase.Left = -50
    frmBase.Width = 15500
    frmBase.Height = 12000
    frmBase.Show
    
    frmMain.Caption = M_CompnyMasterName
    TitleSet (M_CompnyMasterName)
    
    If Fb대리점정보 = "Error" Then
        MsgBox "대리점 정보가 올바르지 않습니다. 확인후 저장하여 주십시요.", vbInformation, "확인"
        frmINIT.Show 1
        
        End
    End If
    
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    ' 2009.05.05 프로그램 업그레이드 작업시 적용
    ' 이전 지사 내용 기록
    If Format(Date, "yyyyMMdd") <= "20090508" Then
        Select Case 대리점정보.StoreCode
            Case "100240", "100068", "100074", "100261"
                Call SetIniStr("Store", "OldMstCode", "1021", iniFile)
                
            Case "100091", "100115", "100195", "100197", "100143"
                Call SetIniStr("Store", "OldMstCode", "1019", iniFile)
                
            Case Else
        End Select
    End If
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

    
    ' 명품세탁 / 할인행사 기간을 확인한다.
    If Format(Date, "YYYY-MM-DD") <= "2006-11-20" Then
    
        If Check_일반할인 = False Then
            MsgBox "명품 세탁 할인 정보 설정 오류 " & Err.Description, vbInformation, "확인"
            End
        End If
    
    
        If Check_명품세탁할인 = False Then
            MsgBox "명품 세탁 할인 정보 설정 오류 " & Err.Description, vbInformation, "확인"
            End
        End If
    
        ' 대리점 정보를 다시 읽는다.
        If Fb대리점정보 = "Error" Then
            MsgBox "대리점 정보가 올바르지 않습니다. 확인후 저장하여 주십시요.", vbInformation, "확인"
            frmINIT.Show 1
            End
        End If
    
    End If
    
    ' 울산
    If Format(Date, "YYYY-MM-DD") <= "2007-07-11" Then
    
        If Check_일반할인_20070711 = False Then
            MsgBox "명품 세탁 할인 정보 설정 오류 " & Err.Description, vbInformation, "확인"
            End
        End If
        
        ' 대리점 정보를 다시 읽는다.
        If Fb대리점정보 = "Error" Then
            MsgBox "대리점 정보가 올바르지 않습니다. 확인후 저장하여 주십시요.", vbInformation, "확인"
            frmINIT.Show 1
            End
        End If
    
    End If
    
    ' 1008-028 이마트 공항점
    If Format(Date, "YYYY-MM-DD") <= "2007-10-31" Then
    
        If Check_일반할인_20071031 = False Then
            MsgBox "명품 세탁 할인 정보 설정 오류 " & Err.Description, vbInformation, "확인"
            End
        End If
        
        ' 대리점 정보를 다시 읽는다.
        If Fb대리점정보 = "Error" Then
            MsgBox "대리점 정보가 올바르지 않습니다. 확인후 저장하여 주십시요.", vbInformation, "확인"
            frmINIT.Show 1
            End
        End If
    
    End If
    
    ' 1008-028 이마트 공항점
    If Format(Date, "YYYY-MM-DD") <= "2007-09-15" Then
    
        If Check_일반할인_20070915 = False Then
            MsgBox "2007-09-15 할인 정보 설정 오류 " & Err.Description, vbInformation, "확인"
            End
        End If
        
        ' 대리점 정보를 다시 읽는다.
        If Fb대리점정보 = "Error" Then
            MsgBox "대리점 정보가 올바르지 않습니다. 확인후 저장하여 주십시요.", vbInformation, "확인"
            frmINIT.Show 1
            End
        End If
    
    End If
    
        ' 1007-"004", "015", "011", "034"
    If Format(Date, "YYYY-MM-DD") <= "2007-10-25" Then
    
        If Check_일반할인_20071025 = False Then
            MsgBox "2007-10-25 할인 정보 설정 오류 " & Err.Description, vbInformation, "확인"
            End
        End If
        
        ' 대리점 정보를 다시 읽는다.
        If Fb대리점정보 = "Error" Then
            MsgBox "대리점 정보가 올바르지 않습니다. 확인후 저장하여 주십시요.", vbInformation, "확인"
            frmINIT.Show 1
            End
        End If
    
    End If
    
    
        ' 1001  205 구미점
    If Format(Date, "YYYY-MM-DD") <= "2007-10-17" Then
    
        If Check_일반할인_20071017 = False Then
            MsgBox "2007-10-17 할인 정보 설정 오류 " & Err.Description, vbInformation, "확인"
            End
        End If
        
        ' 대리점 정보를 다시 읽는다.
        If Fb대리점정보 = "Error" Then
            MsgBox "대리점 정보가 올바르지 않습니다. 확인후 저장하여 주십시요.", vbInformation, "확인"
            frmINIT.Show 1
            End
        End If
    
    End If
    
        ' 1001  042 비산점, 234 학성점
    If Format(Date, "YYYY-MM-DD") <= "2007-10-18" Then
    
        If Check_일반할인_20071018 = False Then
            MsgBox "2007-10-18 할인 정보 설정 오류 " & Err.Description, vbInformation, "확인"
            End
        End If
        
        ' 대리점 정보를 다시 읽는다.
        If Fb대리점정보 = "Error" Then
            MsgBox "대리점 정보가 올바르지 않습니다. 확인후 저장하여 주십시요.", vbInformation, "확인"
            frmINIT.Show 1
            End
        End If
    
    End If
    
    
        ' 이마트 매장 할인
    If Format(Date, "YYYY-MM-DD") <= "2007-11-14" Then
    
        If Check_이마트할인_20071114 = False Then
            MsgBox "2007-11-14 할인 정보 설정 오류 " & Err.Description, vbInformation, "확인"
            End
        End If
        
        ' 대리점 정보를 다시 읽는다.
        If Fb대리점정보 = "Error" Then
            MsgBox "대리점 정보가 올바르지 않습니다. 확인후 저장하여 주십시요.", vbInformation, "확인"
            frmINIT.Show 1
            End
        End If
    
    End If
    
        ' 이마트 매장 할인
    If Format(Date, "YYYY-MM-DD") <= "2007-12-16" Then
    
        If Check_일반할인_20071206 = False Then
            MsgBox "2007-12-05 할인 정보 설정 오류 " & Err.Description, vbInformation, "확인"
            End
        End If
        
        ' 대리점 정보를 다시 읽는다.
        If Fb대리점정보 = "Error" Then
            MsgBox "대리점 정보가 올바르지 않습니다. 확인후 저장하여 주십시요.", vbInformation, "확인"
            frmINIT.Show 1
            End
        End If
    
    End If
    
        ' 이마트 매장 할인
    If Format(Date, "YYYY-MM-DD") <= "2008-03-26" Then
    
        If Check_일반할인_20080320 = False Then
            MsgBox "2008-03-20 할인 정보 설정 오류 " & Err.Description, vbInformation, "확인"
            End
        End If
        
        ' 대리점 정보를 다시 읽는다.
        If Fb대리점정보 = "Error" Then
            MsgBox "대리점 정보가 올바르지 않습니다. 확인후 저장하여 주십시요.", vbInformation, "확인"
            frmINIT.Show 1
            End
        End If
    
    End If
    
    
    ' 이마트 매장 할인
    If Format(Date, "YYYY-MM-DD") <= "2008-11-12" Then
    
        If Check_이마트할인_20081112 = False Then
            MsgBox "2008-11-12 할인 정보 설정 오류 " & Err.Description, vbInformation, "확인"
            End
        End If
        
        ' 대리점 정보를 다시 읽는다.
        If Fb대리점정보 = "Error" Then
            MsgBox "대리점 정보가 올바르지 않습니다. 확인후 저장하여 주십시요.", vbInformation, "확인"
            frmINIT.Show 1
            End
        End If
    
    End If
    
    
    
    ' 일반 매장 할인
    If Format(Date, "YYYY-MM-DD") <= "2008-11-10" Then
    
        If Check_일반할인_20081110 = False Then
            MsgBox "2008-11-10 일반 매장 할인 정보 설정 오류 " & Err.Description, vbInformation, "확인"
            End
        End If
        
        ' 대리점 정보를 다시 읽는다.
        If Fb대리점정보 = "Error" Then
            MsgBox "대리점 정보가 올바르지 않습니다. 확인후 저장하여 주십시요.", vbInformation, "확인"
            frmINIT.Show 1
            End
        End If
    
    End If
    
    
    If Dir(App.Path & "\20090615_ERR.txt") = "" Then
        
        ' 이전 자료를 모두 지운다.
        If Not Dir(App.Path & "\20090615.txt") = "" Then
            Kill App.Path & "\20090615.txt"
        End If
        ADOCon.Execute "DELETE FROM 할인정보 "
        
        Dim FHandle As Integer
        
        FHandle = FreeFile
        Open App.Path & "\20090615_ERR.txt" For Append As FHandle
        Print #FHandle, Now
        Close #FHandle
    
    End If
    
    ' 일반 매장 할인
    If Format(Date, "YYYY-MM-DD") <= "2009-06-15" Then

        If Check_일반할인_20090615 = False Then
            MsgBox "2009-06-08 ~ 2009-06-15 일반 매장 할인 정보 설정 오류 " & Err.Description, vbInformation, "확인"
            End
        End If

        ' 대리점 정보를 다시 읽는다.
        If Fb대리점정보 = "Error" Then
            MsgBox "대리점 정보가 올바르지 않습니다. 확인후 저장하여 주십시요.", vbInformation, "확인"
            frmINIT.Show 1
            End
        End If

    End If
    
    If M_COUPON_KLENZ_CODE <> 대리점정보.MasterCode Then
        Call Store_Sale_Check
    End If
    
    
    ' pds2004 수정
    ' 기본 프린터를 1장으로 등록한다. 중산 유니트에서 기본적으로 2장이 출력되도록 변경을 오구하였기 때문에
    ' 기본 1장으로 선택처리한다. 대리점 정보에서 2장으로 변경 가능 하다.
    If IsNumeric(GetIniStr("Printer", "Count", "", iniFile)) = False Then
        Call SetIniStr("Printer", "Count", "1", iniFile)
    End If
    
    
    
    Exit Sub

Err_Chk:
    Call Delay(5)
    Resume
End Sub

Public Function Db_Connect() As Boolean
    Dim msg As String
   
    On Error GoTo ErrRtn

    Set ADOCon = New ADODB.Connection

    With ADOCon
        '.ConnectionString = "Provider=SQLOLEDB;Persist Security Info=False;User ID=sa;Password=4867591;Initial Catalog=" & Database & ";Data Source=" & DB_Server
        .ConnectionString = "Provider=SQLOLEDB;Persist Security Info=False;User ID=sa;Password=4867591;Initial Catalog=CleanAID;Data Source=XNOTE\SQLEXPRESS"
        .CursorLocation = adUseClient
        .ConnectionTimeout = 30
        .CommandTimeout = 0
        .Open
    End With
   
    Db_Connect = True
   
    Exit Function
   
ErrRtn:
    Db_Connect = False
    
    msg = "오류 # " & Str(Err.Number) & " : " & Err.Description & vbLf & "프로그램을 종료합니다."
    MsgBox msg, vbCritical, "오류"
End Function

Sub FormChk()
    Dim frm As Form
    
    For Each frm In Forms
        If Not (frm.Name = "frmMain" Or frm.Name = "frmBase") Then
            
            Unload frm
        End If
    Next frm
    
    TitleSet (M_CompnyMasterName)
End Sub

Sub FormChkEsc()
    Dim Response As Integer
    Dim frm As Form
    
    If UCase(chkinputflig) = "메뉴" Then
        Response = MsgBox("프로그램을 종료 하시겠습니까?" & Space(10), vbInformation + vbYesNo + vbDefaultButton1, "종료 확인")
        If Response <> vbYes Then Exit Sub
        End
    End If
    
    For Each frm In Forms
        If Not (frm.Name = "frmMain" Or frm.Name = "frmBase") Then
            Unload frm
            chkinputflig = "메뉴"
        End If
    Next frm

    TitleSet (M_CompnyMasterName)
End Sub

Sub TitleSet(txt As String)
    frmMain.Title.Caption = txt
    
    Select Case txt
        ' 입고중
        Case SET_TITLE_INPUT
            frmMain.Command1(1).BackColor = "&H00EBBF76"
            frmMain.Command1(0).BackColor = vbButtonFace '"&H00C0C0C0"
            frmMain.Command1(2).BackColor = vbButtonFace '"&H00C0C0C0"
        ' 출고중
        Case SET_TITLE_OUTPUT
            frmMain.Command1(0).BackColor = "&H00EBBF76"
            frmMain.Command1(2).BackColor = vbButtonFace '"&H00C0C0C0"
            frmMain.Command1(1).BackColor = vbButtonFace '"&H00C0C0C0"
        ' 조회
        Case SET_TITLE_VIEW
            frmMain.Command1(2).BackColor = "&H00EBBF76"
            frmMain.Command1(0).BackColor = vbButtonFace '"&H00C0C0C0"
            frmMain.Command1(1).BackColor = vbButtonFace '"&H00C0C0C0"
        ' 종료
        Case SET_TITLE_EXIT
            frmMain.Command1(0).BackColor = vbButtonFace '"&H00C0C0C0"
            frmMain.Command1(1).BackColor = vbButtonFace '"&H00C0C0C0"
            frmMain.Command1(2).BackColor = vbButtonFace '"&H00C0C0C0"

    End Select
End Sub

Sub Tag_Save()
    Query = "Update 대리점정보 "
    Query = Query & "Set 할인시작일 = '" & Trim(frmMain.TagNo.Caption) & "'"
    ADOCon.Execute Query
End Sub

Function KeyChk(key As Integer) As Integer
    Dim imsgflg As VbMsgBoxResult

    Select Case key
        Case vbKeyF5 '입고
'            chkinputflig = "입고중" '현재 상태..
            If chkinputflig = "입고중" Then
                If Len(GetSpreadText(frm접수.sprGrid, 1, 1)) > 0 Then
                    If MsgBox("입고 작업을 취소 하시겠습니까?", vbInformation + vbYesNo, "종료 확인") <> vbYes Then
                        Exit Function
                    End If
                End If
            End If
            FormChk
            ChkInputKey = True
            frm접수.Show
            If chkDaySale = True Then
               TitleSet SET_TITLE_INPUT & "(목요SALE)"
            Else
               TitleSet SET_TITLE_INPUT
            End If
        
        Case vbKeyF6 '출고
'            chkinputflig = "출고중" '현재 상태..
            If chkinputflig = "입고중" Then
                If Len(GetSpreadText(frm접수.sprGrid, 1, 1)) > 0 Then
                    If MsgBox("입고 작업을 취소 하시겠습니까?", vbInformation + vbYesNo, "종료 확인") <> vbYes Then
                        Exit Function
                    End If
                End If
            End If
            FormChk
            frm출고.Show
            TitleSet SET_TITLE_OUTPUT
        
        Case vbKeyF7 '조회
 '           chkinputflig = "조회중" '현재 상태..
            If chkinputflig = "입고중" Then
                If Len(GetSpreadText(frm접수.sprGrid, 1, 1)) > 0 Then
                    If MsgBox("입고 작업을 취소 하시겠습니까?", vbInformation + vbYesNo, "종료 확인") <> vbYes Then
                        Exit Function
                    End If
                End If
            End If
            
            FormChk
            
            frm조회.Show
            TitleSet SET_TITLE_VIEW
        
        Case vbKeyF8 '조회
        
        Case vbKeyEscape '종료
            TitleSet SET_TITLE_EXIT
                    
            If chkinputflig = "입고중" Then
                If Len(GetSpreadText(frm접수.sprGrid, 1, 1)) > 0 Then
                    If MsgBox("입고 작업을 취소 하시겠습니까?     ", vbInformation + vbYesNo, "종료 확인") <> vbYes Then
                        Exit Function
                    End If
                End If
            End If

'            If TypeName(frmMain.ActiveForm) = "frm접수" Then
'                If MsgBox("입고 프로그램을 끝내시겠습니까?", vbYesNo) = vbNo Then
'                    chkinputflig = "입고중"
'                   Exit Function
'                End If
'            End If
                
            FormChkEsc
            frmMain.Command1(1).Enabled = True
            frmMain.Command1(0).Enabled = True
            frmMain.Command1(2).Enabled = True
            
            ' 서버모드가 아닐경우 입고를 할 수 없게 한다.
            If chkProgramMode <> ServerMode Then frmMain.Command1(1).Enabled = False
    End Select
End Function

Public Function GetIniStr(SectionName As String, LineName As String, defValue As String, iniFile As String) As String
    Dim retStr As String * 256
    Dim Result As Integer
    
    Result = GetPrivateProfileString(SectionName, LineName, defValue, retStr, Len(retStr), iniFile)
    
    GetIniStr = Left(retStr, Result)
End Function

Public Function SetIniStr(SectionName As String, LineName As String, defValue As String, iniFile As String) As String
    Dim Result As Integer
    
    Result = WritePrivateProfileString(SectionName, LineName, defValue, iniFile)
End Function

'+------------------------------------------------------
'+
'+ 2002/11/22
'+
'+루틴설명
'+  1. 택번호를 전달받아 입출고에 있을경우 Ture 를 리턴한다
'+------------------------------------------------------
Public Function TagNoCheck(sTagNo As String) As Boolean
    Query = "Select 번호 "
    Query = Query & "From 입출고 "
    Query = Query & "Where 번호 = '" & sTagNo & "' "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If Not SUBRs.EOF Then
        If SUBRs!번호 = sTagNo Then
            TagNoCheck = True
        Else
            TagNoCheck = False
        End If
    Else
        TagNoCheck = False
    End If
    
    SUBRs.Close
    Set SUBRs = Nothing
End Function

'+------------------------------------------------------
'+
'+ 2002/11/17
'+
'+루틴설명
'+  1. 전달 날짜가 마감 처리 되었는지 확인한다.
'+  2. 리턴값 =>    마감되었으면 : Ture
'+                  미마감시     : False
'+------------------------------------------------------
Public Function DayCloseCheck(sDate As String) As Boolean
    Query = "SELECT 마감여부 "
    Query = Query & " FROM 일일마감 "
    Query = Query & " WHERE 일자 = '" & sDate & "' "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    If SUBRs.EOF Then
        DayCloseCheck = False
    Else
        If SUBRs!마감여부 = "Y" Then
            DayCloseCheck = True
        Else
            DayCloseCheck = False
        End If
    End If
    SUBRs.Close
    Set SUBRs = Nothing
End Function

'+------------------------------------------------------
'+
'+ 2002/11/17
'+
'+루틴설명
'+  1. 정보를 전역번수에 저장한다.
'+  2. 대리점정보을 검색하지 못한경우 'Error'값을 넘긴다.
'+------------------------------------------------------
Public Function Fb대리점정보() As String
    'Dim rsTempTb   As DAO.Recordset

    On Error GoTo Err_Rtn
    
    ' 대리점 코드를 Check한다.
    'Set rsTempTb = MyDB.OpenRecordset("Select * From 대리점정보")
    
    Query = "SELECT * FROM 대리점정보"
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    ' 대리점코드가 없으면 종료를 한다.
    If Not SUBRs.EOF Then
        대리점정보.StoreCode = SUBRs!StoreCode & ""
        대리점정보.StoreName = Trim(SUBRs!StoreName & "")
        대리점정보.StartDate = Trim(SUBRs!StartDate & "")
    
        대리점정보.대리점명 = SUBRs!대리점명 & ""
        대리점정보.대리점번호 = SUBRs!대리점번호 & ""
        대리점정보.대리점색상 = SUBRs!대리점색상 & ""
        대리점정보.목요세일 = SUBRs!목요세일 & ""
        대리점정보.비율 = SUBRs!비율 & ""
        대리점정보.수선 = SUBRs!수선 & ""
        대리점정보.수선마진 = SUBRs!수선마진 & ""
        대리점정보.운동화마진 = SUBRs!운동화마진 & ""
        대리점정보.가죽무스탕마진 = SUBRs!가죽무스탕마진 & ""
        대리점정보.카페트마진 = SUBRs!카페트마진 & ""
        대리점정보.외주운동화마진 = SUBRs!외주운동화마진 & ""
        대리점정보.일수 = SUBRs!일수 & ""
        대리점정보.일수2 = SUBRs!일수2 & ""
        대리점정보.전화1 = SUBRs!전화1 & ""
        대리점정보.전화2 = SUBRs!전화2 & ""
        대리점정보.전화매장 = SUBRs!telStore & ""
        대리점정보.전화SMS = SUBRs!telSMS & ""
        대리점정보.프린터 = SUBRs!프린터 & ""
        대리점정보.할인시작일 = SUBRs!할인시작일 & ""
        대리점정보.할인종료일 = SUBRs!할인종료일 & ""
        대리점정보.마일리지여부 = SUBRs!마일리지여부 & ""
        대리점정보.마일리지증가구분 = SUBRs!마일리지증가구분 & ""
        
        대리점정보.지정할인여부 = SUBRs!지정할인여부 & ""
        대리점정보.지정할인비율 = SUBRs!지정할인비율 & ""
        
        If 대리점정보.지정할인여부 = "Y" Then
            If SUBRs!지정할인시작일 & "" <= Format(Date, "yyyyMMdd") And SUBRs!지정할인종료일 & "" >= Format(Date, "yyyyMMdd") Then
                대리점정보.지정할인여부 = "Y"
            Else
                대리점정보.지정할인여부 = "N"
            End If
        End If
        
        대리점정보.특정할인여부 = SUBRs!특정할인여부 & ""
        대리점정보.특정할인비율 = SUBRs!특정할인비율 & ""
        
        If 대리점정보.특정할인여부 = "Y" Then
            If SUBRs!특정할인시작일 & "" <= Format(Date, "yyyyMMdd") And SUBRs!특정할인종료일 & "" >= Format(Date, "yyyyMMdd") Then
                대리점정보.특정할인여부 = "Y"
            Else
                대리점정보.특정할인여부 = "N"
            End If
        End If
        
        대리점정보.고가세탁비율 = SUBRs!고가세탁비율 & ""
        대리점정보.세탁비환불여부 = SUBRs!세탁비환불여부 & ""
        대리점정보.고객전화번호모두출력 = GetIniStr("Printer", "TelPrint", "0", iniFile)
        대리점정보.MasterCode = GetIniStr("Connect", "MstCode", "", iniFile)
        대리점정보.SMS_EMART = SUBRs!SMS_EMART & ""
        
' 1차 행사 기간
'        If Format(Date, "yyyyMMdd") >= "20090915" And Format(Date, "yyyyMMdd") <= "20091031" Then
' 2차 행사 기간
        If Format(Date, "yyyyMMdd") >= "20091115" And Format(Date, "yyyyMMdd") <= "20091231" Then
            대리점정보.삼성카드할인여부 = "Y"
            대리점정보.삼성카드할인비율 = "10"
        Else
            대리점정보.삼성카드할인여부 = "N"
            대리점정보.삼성카드할인비율 = "0"
        End If
    
    Else
        Fb대리점정보 = "Error"
        MsgBox "등록된 대리점 정보가 없습니다.", vbExclamation, "오류"
    End If
    
    SUBRs.Close
    Set SUBRs = Nothing
    
    Exit Function
    
Err_Rtn:
    Dim Str As String
    
    Fb대리점정보 = "Error"
    
    Str = "DB 정보가 올바르지 않습니다. DB 정보를 현버전으로 변경하시겠습니까?"
    
    If MsgBox(Str, vbCritical + vbYesNo, "DB 변경 확인") = vbYes Then
        Call DatabaseCheck(True)
        
        Exit Function
    End If
    
    MsgBox Err.Number & "->" & Err.Description, vbInformation, "Error"
End Function

'+------------------------------------------------------
'+
'+ 2002/11/17
'+
'+루틴설명
'+  1. 고객정보를 전역번수에 저장한다.
'+  2. 고객정보을 검색하지 못한경우 'Error'값을 넘긴다.
'+------------------------------------------------------
Public Function Fb고객정보(TempCode As String) As String
    On Error GoTo Err_Rtn
    
    ' 대리점 코드를 Check한다.
    Query = "SELECT * FROM 고객정보 "
    Query = Query & " WHERE 고객번호 = '" & TempCode & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
    ' 대리점코드가 없으면 종료를 한다.
    If Rs.EOF Then
        Fb고객정보 = "Error"
        
        고객정보.전화번호 = "Error"
        'MsgBox "등록된 고객 정보가 없습니다.", vbExclamation, "오류"
    Else
        고객정보.고객번호 = Rs!고객번호 & ""
        고객정보.성명 = Rs!성명 & ""
        고객정보.전화1 = Rs!전화1 & ""
        고객정보.전화2 = Rs!전화2 & ""
        고객정보.주소 = Rs!주소 & ""
        고객정보.미수금 = Rs!미수금 & ""
        고객정보.전송구분 = Rs!전송구분 & ""
        고객정보.카드번호 = Rs!카드번호 & ""
        고객정보.전화번호 = Rs!전화1 & "-" & Rs!전화2 & ""
        고객정보.휴대폰 = Rs!휴대폰 & ""
        고객정보.SMS전송여부 = Rs!SMSSendYN & ""
        고객정보.등록일자 = Rs!등록일자 & ""
    End If
    Rs.Close
    Set Rs = Nothing
    
    Exit Function
    
Err_Rtn:
    Fb고객정보 = "Error"
    
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Function

'+------------------------------------------------------
'+
'+ 2002/11/17
'+
'+루틴설명
'+  1. tempDate의 일자의 마감 정보를 전역번수에 저장한다
'+  2. 마감여부를 검색하지 못한경우 'Error'값을 넘긴다.
'+------------------------------------------------------
Public Function Fb일일마감정보(tempDate As String) As String
    On Error GoTo Err_Rtn
    
    If IsDate(tempDate) = False Then
        MsgBox "Fb일일마감정보 함수에 전달된 날짜를 확인 하십시요.", vbExclamation, "오류"
        Fb일일마감정보 = "Error"
        
        Exit Function
    End If
        
    
    'Query = "SELECT * FROM 일일마감 "
    'Query = Query & " WHERE 일자 = '" & tempDate & "'"
    
    ' 대리점 코드를 Check한다.
    'Set SUBRs = MyDB.OpenRecordset("Select * From 일일마감")
    
    Query = "SELECT * FROM 일일마감 "
    Query = Query & " WHERE 일자 = '" & tempDate & "'"
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    ' 대리점코드가 없으면 종료를 한다.
    If Not SUBRs.EOF Then
        일일마감정보.일자 = SUBRs!일자
        일일마감정보.총점수 = SUBRs!총점수
        일일마감정보.반품수량 = SUBRs!반품수량
        일일마감정보.재세탁수량 = SUBRs!재세탁수량
        일일마감정보.수선수량 = SUBRs!수선수량
        일일마감정보.총매출액 = SUBRs!총매출액
        일일마감정보.본사금액 = SUBRs!본사금액
        일일마감정보.대리점금액 = SUBRs!대리점금액
        일일마감정보.수선금액 = SUBRs!수선금액
        일일마감정보.판매구분 = SUBRs!판매구분
        일일마감정보.시작택 = SUBRs!시작택
        일일마감정보.종료택 = SUBRs!종료택
        일일마감정보.마감여부 = SUBRs!마감여부
        일일마감정보.전송여부 = SUBRs!전송여부
    Else
        Fb일일마감정보 = "Error"
        MsgBox "등록된 대리점 정보가 없습니다.", vbExclamation, "오류"
    End If
    
    SUBRs.Close
    Set SUBRs = Nothing
    
    Exit Function
    
Err_Rtn:
    Fb일일마감정보 = "Error"
    MsgBox Err.Number & "->" & Err.Description, vbInformation, "Error"
End Function
 
'+------------------------------------------------------
'+
'+ 2002/11/16
'+
'+루틴설명
'+  1. 고객번호를 가져 오기 위한 루틴
'+     TempName = 고객성명
'+     TempTel1 = 고객전화번호 중 국번
'+     TempTel2 = 고객전화번호 중 나머지 4자리
'+  2. 고객번호를 검색하지 못한경우 'Error'값을 넘긴다.
'+------------------------------------------------------
Public Function Fb고객번호(tempName As String, TempTel1 As String, TempTel2 As String) As String
    On Error GoTo Err_Rtn
    
    Query = "SELECT * FROM 고객정보 "
    Query = Query & " WHERE 성명    = '" & tempName & "' "
    Query = Query & "   AND 전화1   = '" & TempTel1 & "' "
    Query = Query & "   AND 전화2   = '" & TempTel2 & "' "
    Set Rs = New ADODB.Recordset
    Rs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If Rs.RecordCount <= 0 Then
        Fb고객번호 = "Error"
    Else
        Fb고객번호 = Rs!고객번호
    End If
    
    Rs.Close
    Set Rs = Nothing
    
    Exit Function
    
Err_Rtn:
    Fb고객번호 = "Error"
    
    MsgBox Err.Number & "->" & Err.Description, vbInformation, "Error"
End Function

'+------------------------------------------------------
'+
'+ 2005/10/31
'+
'+루틴설명
'+  1. 해당 고객의 해당일자 수금액을 리턴한다.
'+     sCode = 고객코드
'+     sDate = 일자
'+------------------------------------------------------
Public Function Fb수금액(Scode As String, sDate As String) As Double
    On Error GoTo Err_Rtn
    
    Query = "SELECT SUM(금액) AS 수금액 "
    Query = Query & " FROM    미수회수정보 "
    Query = Query & " WHERE   일자    = '" & sDate & "' "
    Query = Query & "   AND   고객코드   = '" & Scode & "' "
    Set Rs = New ADODB.Recordset
    Rs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If Rs.RecordCount <= 0 Then
        Fb수금액 = 0
    Else
        Fb수금액 = Val(Rs!수금액 & "")
    End If
    
    Rs.Close
    Set Rs = Nothing
    
    Exit Function
    
Err_Rtn:
    Fb수금액 = 0
    
    MsgBox Err.Number & "->" & Err.Description, vbInformation, "Error"
End Function

    
Public Sub FpSpreedText(tempSp As Object, SpRow As Long, SpCol As Long, TempText As Variant)
'+------------------------------------------------------
'+
'+ 2002/12/03
'+
'+루틴설명
'+  1. 전달된 spreed에 문자를 기록한다
'+------------------------------------------------------
    Dim nCol As Integer
    Dim nRow As Integer
    
    With tempSp
        nCol = .ActiveCol
        nRow = .ActiveRow
        .Col = SpCol
        .Row = SpRow
        .Text = TempText & ""
        .Col = nCol
        .Row = nRow
    End With
End Sub

Public Function GetSpreadText(tempSp As Object, SpRow As Long, SpCol As Long) As Variant
'+------------------------------------------------------
'+
'+ 2002/12/03
'+
'+루틴설명
'+  1. 원하는 위치의 문자를 리턴한다
'+------------------------------------------------------
    Dim nCol As Integer
    Dim nRow As Integer
    With tempSp
        nCol = .ActiveCol
        nRow = .ActiveRow
        If nCol < 1 Or nCol > .MaxCols Then nCol = .MaxCols
        If nRow < 1 Or nRow > .MaxRows Then nRow = .MaxRows
        .Col = SpCol
        .Row = SpRow
        GetSpreadText = .Text
        .Col = nCol
        .Row = nRow
    End With
End Function

Public Function GetSpreadLine(tempSp As Object) As Integer
'+------------------------------------------------------
'+
'+ 2003/01/16
'+
'+루틴설명
'+  1. 전달된 Spread의 마지막 데이타의  다음 위치를 리턴한다
'+------------------------------------------------------
    Dim i As Integer
    Dim nCol As Integer
    Dim nRow As Integer
    With tempSp
        nCol = .ActiveCol ' 이전 위치 보관
        nRow = .ActiveRow
        If nCol < 1 Or nCol > .MaxCols Then nCol = .MaxCols
        If nRow < 1 Or nRow > .MaxRows Then nRow = .MaxRows
        .Col = 1
        For i = 1 To .MaxRows
            .Row = i
            If Len(.Text) < 1 Then
                Exit For
            End If
        Next i
        .Col = nCol ' 이전 위치 설정
        .Row = nRow
    End With
    GetSpreadLine = i
End Function


Public Sub MoveWindow(myForm As Form)
    myForm.Top = (Screen.Height - myForm.Height) / 2
    myForm.Left = (Screen.Width - myForm.Width) / 2
End Sub


Public Function SetDataBaseTable() As Boolean
    Dim daoDB   As Database        'Access DB
    Dim rs01    As Recordset
    Dim cnt     As Integer
    
    On Error GoTo dbError
    SetDataBaseTable = False
    cnt = 0
    
RE_CHECK:
    Set daoDB = Workspaces(0).OpenDatabase(m_DBPath)
    Query = "SELECT * FROM DataBaseUpdate "
    Set rs01 = MyDB.OpenRecordset(Query)
    
    SetDataBaseTable = True
    Exit Function
    
dbError:
    ' 테이블이 없을 경우
    If Err.Number = 3078 And cnt < 3 Then
        Query = "CREATE TABLE DataBaseUpdate "
        Query = Query & "(DGubun Char(3) Not Null, "
        Query = Query & "DDate Char(8)    Not Null, "
        Query = Query & "DMemo char(50)   Null)"
        daoDB.Execute Query
        
        cnt = cnt + 1
        GoSub RE_CHECK
    End If

End Function

Public Sub DatabaseCheck(bProcMode As Boolean)
' DB가 기존 DB일 경우 필드를 추가한다.
' bMode = true  : 전체를 무조건 실행한다.
' bMode = false : 업그레이드가 안된 내용만 실행한다.

    Dim strValus As String
    Dim i As Integer
    Dim bMode As Boolean
    
    Dim daoDB   As Database        'Access DB
    Dim rs01    As Recordset

    On Error GoTo Err_Rtn
    
    If SetDataBaseTable = False Then Exit Sub
    
    Set daoDB = Workspaces(0).OpenDatabase(m_DBPath)
    
    For i = 1 To 48 ' 업데이드 갯수
    
        ' 업그레이드 안된 경우에 실행한다.
        If bProcMode = False Then
            Query = "SELECT DGubun FROM DataBaseUpdate "
            Query = Query & " WHERE DGubun = '" & Format(i, "000") & "'"
            Set rs01 = MyDB.OpenRecordset(Query)
            If rs01.RecordCount <= 0 Then
                If LaunDryDataBaseUpdate(i) Then
                    bMode = True
                End If
            End If
        
        ' 무조건 실행한다. (조회 ->DB 관리에서 실행)
        Else
            
            If i >= 15 Then
                If i = 15 Then MyDB.Close
                If LaunDryDataBaseUpdate(i) Then
                    bMode = True
                End If
            End If
        End If
    Next i
    
    If bMode Then
        bMsgMode = True
        strMessage = "이전 버전의 DataBase를 현 버전에 맞게 수정 하였습니다.     "
    End If
Err_Rtn:
    Exit Sub
    
End Sub


Public Function LaunDryDataBaseUpdate(Count As Integer) As Boolean
    On Error GoTo dbError
    
    LaunDryDataBaseUpdate = False
    
    Select Case Count
        Case 1
            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD 일수2 Char(2) Null "
            ADOCon.Execute Query
        
            Query = "ALTER TABLE 입출고 "
            Query = Query & "ADD 판매취소 Char(1) Null "
            
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 입출고 "
            Query = Query & "ADD 입고예정일 Char(8) Null "
            
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 일일마감 "
            Query = Query & "ADD 마감여부 Char(1) Null "
            
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 일일마감 "
            Query = Query & "ADD 전송여부 Char(1) Null "
            
            ADOCon.Execute Query
            
            Query = "CREATE TABLE 메일 "
            Query = Query & "(송수신구분 Char(1) Not Null, "
            Query = Query & "메일일자 Char(8)   Not Null, "
            Query = Query & "메일번호 Int       Not Null, "
            Query = Query & "메일내역 Text      Null, "
            Query = Query & "전송구분 Char(1)   Null)"
            
            ADOCon.Execute Query
        
            Query = "ALTER TABLE 고객정보 "
            Query = Query & "ADD 전송구분 Char(1) Null "
            
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 고객정보 "
            Query = Query & "ADD 카드번호 Char(6) Null "
            
            ADOCon.Execute Query
            
            Query = "DROP TABLE 보관증 "
            ADOCon.Execute Query
            
            Query = "CREATE TABLE 보관증 "
            Query = Query & "(일련번호 Int Not Null, "
            Query = Query & "고객전화 Char(9) Not Null, "
            Query = Query & "성명 Char(10) NUll, "
            Query = Query & "접수일 Char(10) NUll, "
            Query = Query & "인도예정일 Char(4) NUll, "
            Query = Query & "택번호 Char(5) NUll, "
            Query = Query & "품명 Char(15) NUll, "
            Query = Query & "색상 Char(4) NUll, "
            Query = Query & "금액 Char(7) NUll, "
            Query = Query & "내용 Char(4) NUll, "
            Query = Query & "합계 Char(3) NUll, "
            Query = Query & "합계금액 Char(7) NUll, "
            Query = Query & "수령액 Char(7) NUll, "
            Query = Query & "잔액 Char(7) NUll, "
            Query = Query & "대리점명 Char(15) NUll, "
            Query = Query & "대리점전화 Char(9) NUll, "
            Query = Query & "상표 Char(20) NUll) "
            
            ADOCon.Execute Query
            GoSub SUB_OK
            
        Case 2
        
            Query = "CREATE TABLE 보관증 "
            Query = Query & "(일련번호 Int  Not Null, "
            Query = Query & "고객전화 Char(9) Not Null, "
            Query = Query & "성명 Char(10) NUll, "
            Query = Query & "접수일 Char(10) NUll, "
            Query = Query & "인도예정일 Char(4) NUll, "
            Query = Query & "택번호 Char(5) NUll, "
            Query = Query & "품명 Char(15) NUll, "
            Query = Query & "색상 Char(4) NUll, "
            Query = Query & "금액 Char(7) NUll, "
            Query = Query & "내용 Char(4) NUll, "
            Query = Query & "합계 Char(3) NUll, "
            Query = Query & "합계금액 Char(7) NUll, "
            Query = Query & "수령액 Char(7) NUll, "
            Query = Query & "잔액 Char(7) NUll, "
            Query = Query & "대리점명 Char(15) NUll, "
            Query = Query & "대리점전화 Char(9) NUll, "
            Query = Query & "상표 Char(20) NUll) "
    
            ADOCon.Execute Query
            GoSub SUB_OK
    
        Case 3
        
            Query = "DROP TABLE 사고품 "
            ADOCon.Execute Query
            
            Query = "CREATE TABLE 사고품 "
            Query = Query & "(일련번호 Int PRIMARY KEY Not Null, "
            Query = Query & "접수일 Char(8) Not NUll, "
            Query = Query & "성명 Char(10) Not NUll, "
            Query = Query & "고객전화 Char(14) Not Null, "
            Query = Query & "주소 Char(50) NUll, "
            Query = Query & "휴대폰 Char(14) NUll, "
            Query = Query & "품명 Char(15) NUll, "
            Query = Query & "상표 Char(15) NUll, "
            Query = Query & "구입일자 Char(8) NUll, "
            Query = Query & "색상 Char(4) NUll, "
            Query = Query & "구입처 Char(15) NUll, "
            Query = Query & "최초택번호 Char(5) NUll, "
            Query = Query & "최종택번호 Char(5) NUll, "
            Query = Query & "구입형태 Char(15) NUll, "
            Query = Query & "최초입고일 Char(10) NUll, "
            Query = Query & "최종입고일 Char(10) NUll, "
            Query = Query & "구입가격 Char(7) NUll, "
            Query = Query & "사고접수일 Char(8) NUll, "
            Query = Query & "사고종류 Char(50) NUll, "
            Query = Query & "사고내용 Char(50) NUll, "
            Query = Query & "사고의견 Char(50) NUll, "
            Query = Query & "보상금액 Char(10) NUll, "
            Query = Query & "합의금액 Char(10) NUll, "
            Query = Query & "처리유무 Char(10) NUll) "
    
            ADOCon.Execute Query
            GoSub SUB_OK
        
        Case 4
            ' 운동화 마진
            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD 운동화마진 Char(3) Null "
            ADOCon.Execute Query
            GoSub SUB_OK
            
        Case 5
            ' 본사에서 대리점으로 출고된 내역을 알기위하여 신규로 추가
            Query = "CREATE TABLE 본사입고 ("
            Query = Query & "작업일자     Char(8)     Null, "
            Query = Query & "본사출고일   Char(8)     Null, "
            Query = Query & "입고일자     char(8)     Null, "
            Query = Query & "구분         char(1)     Null, "
            Query = Query & "택번호       Char(4)   Null)"
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD 가죽무스탕마진 Char(3) Null "
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD 카페트마진 Char(3) Null "
            ADOCon.Execute Query
            GoSub SUB_OK
        
        Case 6
            Query = "ALTER TABLE 고객정보 "
            Query = Query & "ADD 휴대폰 Char(14) Null "
            ADOCon.Execute Query
            GoSub SUB_OK
            
        Case 7
            
'            Query = "DROP TABLE 마일리지현황 "
'            ADOCon.Execute Query
'
'            Query = "DROP TABLE 마일리지스토리 "
'            ADOCon.Execute Query

            Query = "CREATE TABLE 마일리지현황 "
            Query = Query & "(고객번호    Char(6) Null, "
            Query = Query & "총사용금액   Int Not Null, "
            Query = Query & "마일리지     Int Not Null, "
            Query = Query & "최종발생금액 Int Not Null, "
            Query = Query & "발생누계 Int Not Null, "
            Query = Query & "사용마일리지 Int Not Null, "
            Query = Query & "최종거래일자 char(8) Null, "
            Query = Query & "전송여부     Char(1) Null) "
            ADOCon.Execute Query
        
            Query = "CREATE TABLE 마일리지스토리 "
            Query = Query & "(발생일자    char(14) Not Null, "
            Query = Query & "고객번호     char(6) Null, "
            Query = Query & "발생마일리지 Int Not Null, "
            Query = Query & "사용마일리지 Int Not Null, "
            Query = Query & "삭제마일리지 Int Not Null, "
            Query = Query & "보관증       Int Not Null, "
            Query = Query & "전송여부     Char(1) Null) "
            ADOCon.Execute Query
            GoSub SUB_OK
        
        Case 8
            
            Query = "ALTER TABLE 마일리지스토리 "
            Query = Query & "ADD 반환마일리지 Int       Not Null "
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 마일리지현황 "
            Query = Query & "ADD 미반환마일리지 Int       Not Null "
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 일일마감 "
            Query = Query & "ADD 발생마일리지 Int       Not Null "
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 일일마감 "
            Query = Query & "ADD 사용마일리지 Int       Not Null "
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 일일마감 "
            Query = Query & "ADD 삭제마일리지 Int       Not Null "
            ADOCon.Execute Query
        
            Query = "ALTER TABLE 입출고 "
            Query = Query & "ADD 환불일자 char(14)     Null "
            ADOCon.Execute Query
            GoSub SUB_OK
        
        Case 9
            Query = "CREATE TABLE 미수회수정보 ("
            Query = Query & "일자         Char(8)     NOT Null, "
            Query = Query & "고객코드     Char(6)     NOT Null, "
            Query = Query & "시간         Char(6)     NOT Null, "
            Query = Query & "금액         INT         Null, "
            Query = Query & "비고         Char(30)    Null)"
            ADOCon.Execute Query
            GoSub SUB_OK
        
        Case 10
            Query = "ALTER TABLE 마일리지현황 "
            Query = Query & "ADD 전송일자 char(8) Null "
            ADOCon.Execute Query
            Query = "ALTER TABLE 마일리지스토리 "
            Query = Query & "ADD 전송일자 char(8) Null "
            ADOCon.Execute Query
            GoSub SUB_OK
        
        Case 11
            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD 마일리지여부 Char(1) Null "
            ADOCon.Execute Query
            GoSub SUB_OK
            
        Case 12
            Query = "ALTER TABLE 보관증 "
            Query = Query & "ADD 마일리지 Char(10)    Null"
            ADOCon.Execute Query
            GoSub SUB_OK
            
        Case 13
            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD 보관증종류 Char(1)    Null"
            ADOCon.Execute Query
            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD 특정할인여부 Char(1)    Null"
            ADOCon.Execute Query
            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD 특정할인비율 Char(3)    Null"
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 보관증 "
            Query = Query & "ADD 미수합계 Char(10)    Null"
            ADOCon.Execute Query
            Query = "ALTER TABLE 보관증 "
            Query = Query & "ADD 전일미수 Char(10)    Null"
            ADOCon.Execute Query
            Query = "ALTER TABLE 보관증 "
            Query = Query & "ADD 수금액 Char(10)    Null"
            ADOCon.Execute Query
            Query = "ALTER TABLE 보관증 "
            Query = Query & "ADD 누적마일리지 Char(10)    Null"
            ADOCon.Execute Query
            Query = "ALTER TABLE 보관증 "
            Query = Query & "ADD 마일리지잔액 Char(10)    Null"
            ADOCon.Execute Query
            GoSub SUB_OK
            
            
        Case 14
            Query = "CREATE INDEX idx_Primary"
            Query = Query & " ON 본사입고"
            Query = Query & " (작업일자, 본사출고일, 입고일자,구분,택번호)"
            Query = Query & " WITH PRIMARY"
            ADOCon.Execute Query
            GoSub SUB_OK

 
        Case 15
            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD 고가세탁비율 Char(3)    Null"
            ADOCon.Execute Query
            
            Query = "UPDATE 대리점정보 SET "
            Query = Query & " 고가세탁비율 = '300' "
            ADOCon.Execute Query
            GoSub SUB_OK
            
        Case 16
            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD 마일리지검사일자 Char(20)    Null"
            ADOCon.Execute Query
            
            Query = "UPDATE 대리점정보 SET "
            Query = Query & " 마일리지검사일자 = '' "
            ADOCon.Execute Query
            GoSub SUB_OK
            
        Case 17
            Query = "ALTER TABLE 입출고 "
            Query = Query & "ADD 수선금액 Int       Not Null "
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD 마일리지증가구분 Char(1)    Null"
            ADOCon.Execute Query
            
            Query = "UPDATE 대리점정보 SET "
            Query = Query & " 마일리지증가구분 = '0' "
            ADOCon.Execute Query
            GoSub SUB_OK
            
            
        Case 18
            Query = " CREATE TABLE 보관가격 "
            Query = Query & "(보관월 Char(2)  Not Null, "
            Query = Query & "아이템수 Double     Not Null , "
            Query = Query & "보관개월수 Double   Not Null, "
            Query = Query & "보관가격 Double     Not Null )"
            ADOCon.Execute Query


            Query = " CREATE TABLE 보관리스트 "
            Query = Query & "(KeyCode     Char(14)    Not Null, "
            Query = Query & "MemRecord    Char(2)     Not Null, "
            Query = Query & "InputNumber  Char(20)    Not Null, "
            Query = Query & "InputDate    Char(17)    Not Null, "
            Query = Query & "InputID      Char(20)    Null, "
            Query = Query & "InputName    Char(20)    Not Null, "
            Query = Query & "EMail        Char(40)    Null, "
            Query = Query & "UserCode     Char(2)     Not Null, "
            Query = Query & "UserNumber   Char(13)    Not Null, "
            Query = Query & "StoreCode    Char(20)    Null, "
            Query = Query & "SaleGubunCode Char(2)    Not Null, "
            Query = Query & "SaleEndDate  Char(8)     Not Null, "
            Query = Query & "Price        Char(8)     Not Null, "
            Query = Query & "DevTimeCode  Char(2)     Not Null, "
            Query = Query & "ItemCount    Char(6)     Not Null, "
            Query = Query & "StatsFlag    Char(1)     Not Null) "
            ADOCon.Execute Query


            Query = " CREATE TABLE 보관상품리스트 "
            Query = Query & "(KeyCode     Char(14)    Not Null, "
            Query = Query & "ItemRecord   Char(2)     Not Null, "
            Query = Query & "ItemIndex    Char(6)     Not Null , "
            Query = Query & "InputDate    Char(8)     Not Null, "
            Query = Query & "Tag          Char(20)    Not Null, "
            Query = Query & "GoodsCode    Char(16)    Not Null, "
            Query = Query & "SizeGubun    Char(2)     Not Null, "
            Query = Query & "SizeCode     Char(2)     Not Null, "
            Query = Query & "Color        Char(10)    Not Null, "
            Query = Query & "BrandName    Char(20)    Not Null, "
            Query = Query & "BuyPrice     Char(10)    Null, "
            Query = Query & "BuyDate      Char(8)     Null, "
            Query = Query & "ASGubun      Char(2)     Null, "
            Query = Query & "BleCount     Char(3)     Not Null, "
            Query = Query & "StatsFlag    Char(1)     Not Null) "
            ADOCon.Execute Query

            Query = " CREATE TABLE 보관하자리스트 "
            Query = Query & "(KeyCode     Char(14)    Not Null, "
            Query = Query & "InputDate    Char(8)     Not Null, "
            Query = Query & "ItemIndex    Char(6)     Not Null, "
            Query = Query & "ItemCount    Char(2)     Not Null, "
            Query = Query & "ItemRemark   Char(50)    Null, "
            Query = Query & "StatsFlag    Char(1)     Not Null) "
            ADOCon.Execute Query

'            Call Fn_보관가격표생성

            GoSub SUB_OK
            
        Case 19
            Query = "ALTER TABLE 보관증 "
            Query = Query & "ADD 카드금액 Int       Not Null "
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 일일마감 "
            Query = Query & "ADD 카드금액 Int       Not Null "
            ADOCon.Execute Query
            
            Query = " CREATE TABLE 카드금액 "
            Query = Query & "(결재일자 Char(8)    Not Null, "
            Query = Query & "접수시간 Char(6)     Not Null, "
            Query = Query & "고객번호 Char(6)     Not Null , "
            Query = Query & "금액     Double      Not Null )"
            ADOCon.Execute Query
            
            GoSub SUB_OK
            
        Case 20
            Query = "ALTER TABLE 일일마감 "
            Query = Query & "ADD 카드건수 Int       Not Null "
            ADOCon.Execute Query
            
            GoSub SUB_OK
            
        Case 21
            Query = "ALTER TABLE 입출고 "
            Query = Query & "ADD 본출일자 Char(8) Null"
            ADOCon.Execute Query
        
            Query = "ALTER TABLE 입출고 "
            Query = Query & "ADD 본출입고구분 Char(8) Null"
            ADOCon.Execute Query
        
            Query = "UPDATE 입출고 SET "
            Query = Query & " 본출일자 = ' ', 본출입고구분 = ' ' "
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD ServerIP Char(20)    Null"
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD ServerDB Char(20)    Null"
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD ServerUser Char(20)    Null"
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD ServerPass Char(20)    Null"
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD TimeOut Int      Null"
            ADOCon.Execute Query
            
            Query = "UPDATE 대리점정보 SET "
            Query = Query & " ServerIP = 'store.clean-aid.co.kr,8657', ServerDB = 'Laundry', ServerUser = 'sa', ServerPass = '', TimeOut = 30 "
            ADOCon.Execute Query
            GoSub SUB_OK
            
        Case 22
            Query = "UPDATE 입출고 SET "
            Query = Query & " 본출일자 = '20070613', 본출입고구분 = '일괄' "
            Query = Query & " WHERE 본출 = '出' "
            ADOCon.Execute Query
            GoSub SUB_OK
            
            
        Case 23
            Query = " CREATE TABLE 문자발송문 "
            Query = Query & "(순번 Char(2)    Null, "
            Query = Query & "내용 Char(200)     Null )"
            ADOCon.Execute Query
            
            GoSub SUB_OK
            
        Case 24
            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD StoreCode Char(6)    Null"
            ADOCon.Execute Query

            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD StoreName Char(50)    Null"
            ADOCon.Execute Query

            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD StartDate Char(10)    Null"
            ADOCon.Execute Query

            Query = "UPDATE 대리점정보 SET "
            Query = Query & " StoreCode = '000000', StoreName = '000000', StartDate = ' ' "
            ADOCon.Execute Query
            
            GoSub SUB_OK
            
        Case 25
            Query = "ALTER TABLE 고객정보 "
            Query = Query & "ADD SMSSendYN Char(1)    Null"
            ADOCon.Execute Query

            Query = "UPDATE 고객정보 SET "
            Query = Query & " SMSSendYN = 'Y'"
            ADOCon.Execute Query

            GoSub SUB_OK
            
            
        Case 26
            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD TelStore Char(20)    Null"
            ADOCon.Execute Query

            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD TelSMS Char(20)    Null"
            ADOCon.Execute Query
            
            Query = "UPDATE 대리점정보 SET "
            Query = Query & " TelStore = 전화1 + '-' + 전화2 ,TelSMS = 전화1 + '-' + 전화2 "
            ADOCon.Execute Query
            
            GoSub SUB_OK
    
        Case 27
            
            ADOCon.Execute "ALTER TABLE 대리점정보 DROP COLUMN ServerIP "
            ADOCon.Execute "ALTER TABLE 대리점정보  ADD ServerIP Char(50)    Null"
            GoSub SUB_OK
        
        Case 28
            ADOCon.Execute "ALTER TABLE 대리점정보 DROP COLUMN ServerIP "
            ADOCon.Execute "ALTER TABLE 대리점정보  ADD ServerIP Char(50)    Null"
            GoSub SUB_OK
            
        Case 29
            GoSub SUB_OK
        
        Case 30
            ' ADOCon.Execute "drop  TABLE 쿠폰자료 "
            
            Query = " CREATE TABLE 쿠폰자료 "
            Query = Query & "(접수일자    Char(8)     Null, "
            Query = Query & " 대리점코드  Char(6)     Null,"
            Query = Query & " 쿠폰번호    Char(8)     Null,"
            Query = Query & " 쿠폰단가    Int         Null,"
            Query = Query & " 쿠폰금액    int         Null,"
            Query = Query & " 고객번호    Char(6)     Null,"
            Query = Query & " 고객이름    Char(30)    Null,"
            Query = Query & " 접수금액    int         Null,"
            Query = Query & " 전송여부    Char(1)     Null,"
            Query = Query & " 전송일자    Char(8)     Null )"
            ADOCon.Execute Query
            
            Query = "CREATE INDEX idx_Coupon_Primary"
            Query = Query & " ON 쿠폰자료"
            Query = Query & " (접수일자, 대리점코드, 쿠폰번호 )"
            Query = Query & " WITH PRIMARY"
            ADOCon.Execute Query
            GoSub SUB_OK
    
        Case 31
            ADOCon.Execute "ALTER TABLE 고객정보 ALTER COLUMN 미수금 CHAR(10)"
            GoSub SUB_OK
    
        Case 32
            ADOCon.Execute "ALTER TABLE 보관증  ADD CouponCnt Char(5)    Null"
            ADOCon.Execute "ALTER TABLE 보관증  ADD CouponNumber Char(100)    Null"
            ADOCon.Execute "ALTER TABLE 보관증  ADD CouponMoney Char(10)    Null"
            GoSub SUB_OK
        
        Case 33
            ADOCon.Execute "ALTER TABLE 쿠폰자료  ADD 택번호 Char(3)    Null"
            GoSub SUB_OK
            
        Case 34
            Query = "ALTER TABLE 고객정보 "
            Query = Query & "ADD 등록일자 Char(8)    Null"
            ADOCon.Execute Query

            Query = "UPDATE 고객정보 SET "
            Query = Query & " 등록일자 = '19000101'"
            Query = Query & " WHERE 등록일자 IS NULL "
            ADOCon.Execute Query

            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD SMS_EMART Char(1)    Null"
            ADOCon.Execute Query

            Query = "UPDATE 대리점정보 SET "
            Query = Query & " SMS_EMART = 'N' "
            ADOCon.Execute Query
            
            Query = "UPDATE 대리점정보 SET "  ' 29개 이마트 기본 설정
            Query = Query & " SMS_EMART = 'Y' "
            Query = Query & " WHERE StoreCode IN('100012','100015','100028','100029','100084','100008',"
            Query = Query & "'100011','100123','100170','100009','100023','100126','100031','100038',"
            Query = Query & "'100056','100022','100222','100264','100128','100027','100032','100211',"
            Query = Query & "'100259','100240','100143','100197','100200','100030','100120','100021')  "
            ADOCon.Execute Query
            
            GoSub SUB_OK
            
        Case 35
            ADOCon.Execute "ALTER TABLE 사고품  ADD 전송구분 Char(1)    Null"
            
            Query = "UPDATE 사고품 SET "
            Query = Query & " 전송구분 = 'N'"
            Query = Query & " WHERE 전송구분 IS NULL "
            ADOCon.Execute Query
            
            GoSub SUB_OK
    
        Case 36
            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD 쿠폰할인여부 Char(1)    Null"
            ADOCon.Execute Query
            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD 쿠폰할인비율 Char(3)    Null"
            ADOCon.Execute Query
    
            GoSub SUB_OK
    
        Case 37
            Query = "ALTER TABLE 입출고 "
            Query = Query & "ADD 본사전송구분 Char(1)    Null"
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD 외주운동화마진 Char(3) Null "
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 입출고 "
            Query = Query & "ADD 외주운동화마진 int Null "
            ADOCon.Execute Query
            
            GoSub SUB_OK
    
        Case 38
            Query = "ALTER TABLE 입출고 "
            Query = Query & "ADD 세탁비환불일자 Char(20)    Null"
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD 세탁비환불여부 Char(1) Null "
            ADOCon.Execute Query
            
            Query = "UPDATE 대리점정보 SET "
            Query = Query & " 세탁비환불여부 = 'N'"
            Query = Query & " WHERE 세탁비환불여부 IS NULL "
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 일일마감 "
            Query = Query & "ADD 세탁비환불건수 int Null "
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 일일마감 "
            Query = Query & "ADD 세탁비환불금액 int Null "
            ADOCon.Execute Query
            
            Query = "UPDATE 일일마감 SET "
            Query = Query & " 세탁비환불건수 = 0, "
            Query = Query & " 세탁비환불금액 = 0 "
            Query = Query & " WHERE 세탁비환불건수 IS NULL "
            ADOCon.Execute Query
            
            GoSub SUB_OK
    
        Case 39
            Query = "ALTER TABLE 메일 "
            Query = Query & "ADD 조회시작일 Char(8)    Null"
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 메일 "
            Query = Query & "ADD 조회종료일 Char(8)    Null"
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 메일 "
            Query = Query & "ADD 수신여부 Char(1)    Null"
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 메일 "
            Query = Query & "ADD 수신일자 Char(200)    Null"
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 일일마감 "
            Query = Query & "ADD 운동화건수 int Null "
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 일일마감 "
            Query = Query & "ADD 운동화금액 int Null "
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 일일마감 "
            Query = Query & "ADD 운동화비율 int Null "
            ADOCon.Execute Query
            
            Query = "UPDATE 일일마감 SET "
            Query = Query & " 운동화건수 = 0, "
            Query = Query & " 운동화금액 = 0, "
            Query = Query & " 운동화비율 = 0 "
            Query = Query & " WHERE 운동화건수 IS NULL "
            ADOCon.Execute Query
            
             ' 미전송 자료가 있어서 재 전송 처리함.
            If Format(Date, "yyyyMMdd") >= "20090901" And Format(Date, "yyyyMMdd") <= "20090910" Then
                Query = "UPDATE 사고품 SET "
                Query = Query & " 전송구분 = 'N' "
                ADOCon.Execute Query
            End If
            
            'alter table 메일 add column 메일내역 memo
            GoSub SUB_OK
            
        Case 40
            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD 특정할인시작일 Char(8) "
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD 특정할인종료일 Char(8) "
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD 쿠폰할인시작일 Char(8) "
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 대리점정보 "
            Query = Query & "ADD 쿠폰할인종료일 Char(8) "
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 일일마감 "
            Query = Query & "ADD 삼성카드할인건수 int Null "
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 일일마감 "
            Query = Query & "ADD 삼성카드할인금액 int Null "
            ADOCon.Execute Query
            
            Query = "UPDATE 일일마감 SET "
            Query = Query & " 삼성카드할인건수 = 0, "
            Query = Query & " 삼성카드할인금액 = 0 "
            Query = Query & " WHERE 삼성카드할인건수 IS NULL "
            ADOCon.Execute Query
            
            Query = "UPDATE 대리점정보 SET "
            Query = Query & " 세탁비환불여부 = 'Y'"
            ADOCon.Execute Query
            
            Dim rsTempTb   As DAO.Recordset
            Set rsTempTb = MyDB.OpenRecordset("Select 특정할인여부, 쿠폰할인여부 From 대리점정보")
            
            If rsTempTb.Fields("특정할인여부") & "" = "Y" Then
                Query = "UPDATE 대리점정보 SET "
                Query = Query & " 특정할인시작일 = '20090101', "
                Query = Query & " 특정할인종료일 = '20990101'  "
                Query = Query & " WHERE 특정할인시작일 IS NULL "
                ADOCon.Execute Query
            Else
                Query = "UPDATE 대리점정보 SET "
                Query = Query & " 특정할인시작일 = '20090101', "
                Query = Query & " 특정할인종료일 = '20090101'  "
                Query = Query & " WHERE 특정할인시작일 IS NULL "
                ADOCon.Execute Query
            End If
    
            If rsTempTb.Fields("쿠폰할인여부") & "" = "Y" Then
                Query = "UPDATE 대리점정보 SET "
                Query = Query & " 쿠폰할인시작일 = '20090101', "
                Query = Query & " 쿠폰할인종료일 = '20990101' "
                Query = Query & " WHERE 쿠폰할인시작일 IS NULL "
                ADOCon.Execute Query
            Else
                Query = "UPDATE 대리점정보 SET "
                Query = Query & " 쿠폰할인시작일 = '20090101', "
                Query = Query & " 쿠폰할인종료일 = '20090101' "
                Query = Query & " WHERE 쿠폰할인시작일 IS NULL "
                ADOCon.Execute Query
            End If
            
            GoSub SUB_OK
        
        Case 41
            ADOCon.Execute "ALTER TABLE 메일 DROP COLUMN 메일내역 "
            
            Query = "ALTER TABLE 메일 "
            Query = Query & "ADD 메일내역 memo "
            ADOCon.Execute Query
            
            GoSub SUB_OK
            
            
        Case 42
            Query = "ALTER TABLE 일일마감 "
            Query = Query & "ADD 삼성카드할인고객수 int Null "
            ADOCon.Execute Query
            
            Query = "UPDATE 일일마감 SET "
            Query = Query & " 삼성카드할인고객수 = 0 "
            Query = Query & " WHERE 삼성카드할인고객수 IS NULL "
            ADOCon.Execute Query
    
            GoSub SUB_OK
    
    
        Case 43
            Query = "ALTER TABLE 대리점정보 ADD 지정할인여부 Char(1) "
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 대리점정보 ADD 지정할인비율 Char(3) "
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 대리점정보 ADD 지정할인시작일 Char(8) "
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 대리점정보 ADD 지정할인종료일 Char(8) "
            ADOCon.Execute Query
            
            Query = "UPDATE 대리점정보 SET "
            Query = Query & " 지정할인여부 = N, "
            Query = Query & " 지정할인비율 = 0, "
            Query = Query & " 지정할인시작일 = 20090101, "
            Query = Query & " 지정할인종료일 = 20090101 "
            Query = Query & " WHERE 지정할인여부 IS NULL "
            ADOCon.Execute Query
            
            GoSub SUB_OK
            
        Case 44
            Query = " CREATE TABLE 세트응모번호 "
            Query = Query & "(응모번호        Char(8)     Null, "
            Query = Query & " 세트Key         Char(20)    Null,"
            Query = Query & " 일자            Char(8)     Null,"
            Query = Query & " 고객코드        Char(7)     Null,"
            Query = Query & " 고객명          Char(10)    Null,"
            Query = Query & " 고객전화번호    Char(15)    Null,"
            Query = Query & " 휴대폰번호      Char(15)    Null,"
            Query = Query & " SendDate        Char(8)     Null) "
            ADOCon.Execute Query
    
            Query = " CREATE TABLE 세트상품정보 "
            Query = Query & "(접수일자        Char(8)     Null, "
            Query = Query & " 세트Key         Char(20)     Null,"
            Query = Query & " 고객코드        Char(7)     Null,"
            Query = Query & " 고객명          Char(10)    Null,"
            Query = Query & " 고객전화번호    Char(15)    Null,"
            Query = Query & " 휴대폰번호      Char(15)    Null,"
            
            Query = Query & " 정상금액        INT    Null,"
            Query = Query & " 세트금액        INT    Null,"
            Query = Query & " 세트할인금액    INT    Null,"
            Query = Query & " 에누리할인금액  INT    Null,"
            Query = Query & " 적용합계금액    INT    Null,"
            
            Query = Query & " 세트2           INT    Null,"
            Query = Query & " 세트3           INT    Null,"
            Query = Query & " 세트4           INT    Null,"
            Query = Query & " 세트5           INT    Null,"
            Query = Query & " 세트6           INT    Null,"
            Query = Query & " 세트7           INT    Null,"
            Query = Query & " 세트8           INT    Null,"
            Query = Query & " 세트9           INT    Null,"
            Query = Query & " 세트10          INT    Null,"
            Query = Query & " SendDate        Char(8)     Null) "
            
            ADOCon.Execute Query
    
    
            Query = "ALTER TABLE 입출고 ADD 세트Key   Char(20) ":            ADOCon.Execute Query
            Query = "ALTER TABLE 입출고 ADD 세트구분  Char(4) ":             ADOCon.Execute Query
            Query = "ALTER TABLE 입출고 ADD 세트금액1 Char(10) ":            ADOCon.Execute Query
            Query = "ALTER TABLE 입출고 ADD 세트금액2 Char(10) ":            ADOCon.Execute Query
            Query = "ALTER TABLE 입출고 ADD 정상가격  Char(10) ":            ADOCon.Execute Query

            Query = "ALTER TABLE 보관증 ADD 세트Key   Char(20) ":            ADOCon.Execute Query
            Query = "ALTER TABLE 보관증 ADD 세트구분  Char(4) ":             ADOCon.Execute Query
            Query = "ALTER TABLE 보관증 ADD 세트금액1 Char(10) ":            ADOCon.Execute Query
            Query = "ALTER TABLE 보관증 ADD 세트금액2 Char(10) ":            ADOCon.Execute Query
            Query = "ALTER TABLE 보관증 ADD 정상가격  Char(10) ":            ADOCon.Execute Query
            Query = "ALTER TABLE 보관증 ADD 상품코드  Char(4) ":            ADOCon.Execute Query
            
            Query = "ALTER TABLE 보관증 ADD 전체정상금액 Char(10) ":             ADOCon.Execute Query
            Query = "ALTER TABLE 보관증 ADD 전체세트금액1 Char(10) ":            ADOCon.Execute Query
            Query = "ALTER TABLE 보관증 ADD 전체세트할인 Char(10) ":             ADOCon.Execute Query
            Query = "ALTER TABLE 보관증 ADD 전체세트에누리할인 Char(10) ":       ADOCon.Execute Query

            GoSub SUB_OK
    
        Case 45
            Query = "UPDATE 대리점정보 SET "
            Query = Query & " 특정할인여부 = 'Y', "
            Query = Query & " 특정할인비율 = '100', "
            Query = Query & " 특정할인시작일 = '20091211', "
            Query = Query & " 특정할인종료일 = '20091231' "
            ADOCon.Execute Query
            
            GoSub SUB_OK
    
        Case 46
            Query = "ALTER TABLE 세트상품정보 ADD 무료세탁권수   INT    Null "
            ADOCon.Execute Query
            
            Query = "UPDATE 세트상품정보 SET "
            Query = Query & " 무료세탁권수 = 0 "
            Query = Query & " WHERE 무료세탁권수 IS NULL "
            ADOCon.Execute Query
            
            GoSub SUB_OK
    
        Case 47
            Query = "ALTER TABLE 참조코드 ADD 출력순번   Char(4)    Null "
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 할인정보 ADD 출력순번   Char(4)    Null "
            ADOCon.Execute Query
            
            Query = "ALTER TABLE 목요세일 ADD 출력순번   Char(4)    Null "
            ADOCon.Execute Query
            
            Query = "UPDATE 참조코드 SET  출력순번 = '' WHERE 출력순번 IS NULL "
            ADOCon.Execute Query
            Query = "UPDATE 할인정보 SET  출력순번 = '' WHERE 출력순번 IS NULL "
            ADOCon.Execute Query
            Query = "UPDATE 목요세일 SET  출력순번 = '' WHERE 출력순번 IS NULL "
            ADOCon.Execute Query
            
            GoSub SUB_OK
    
        Case 48
            Query = "UPDATE 대리점정보 SET "
            Query = Query & " 특정할인여부 = 'Y', "
            Query = Query & " 특정할인비율 = '100', "
            Query = Query & " 특정할인시작일 = '20091211', "
            Query = Query & " 특정할인종료일 = '20100131' "
            ADOCon.Execute Query
            
            GoSub SUB_OK
    
    End Select
    
    LaunDryDataBaseUpdate = True
    Exit Function
    
SUB_OK:
    
    Query = "DELETE FROM DataBaseUpdate "
    Query = Query & " WHERE DGubun = '" & Format(Count, "000") & "'"
    ADOCon.Execute Query
    
    Query = "INSERT INTO DataBaseUpdate "
    Query = Query & " (DGubun, DDate, DMemo ) "
    Query = Query & " VALUES ( "
    Query = Query & "'" & Format(Count, "000") & "', "
    Query = Query & "'" & Format(Date, "yyyyMMdd") & "', "
    Query = Query & "'" & " " & "') "
    ADOCon.Execute Query
    
    Return
    
dbError:
'   이미 있는 필드는 오류가 난다.
'   MsgBox Err.Number & "->" & Err.Description, vbInformation, "Error"
    Resume Next
    Return

End Function

Public Function ShowFolderSize(filespec) As Double
    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    ShowFolderSize = f.Size
End Function


'+------------------------------------------------------
'+
'+ 2003/01/21
'+
'+루틴설명
'+
'+  1. strTag로 전달된 택번호를 가지고  mode에 맞게 증가 하거나 감소한다.
'+     증가나 감소한경우 그택이 판매 취소된것인지 확인한다.
'+     중간택을 삭제한 경우 다음택번호가 사라지는 것을 방지하기 위하여
'+     다음택 사용여부 확인하여 사용중이면 그 다음 택번호증가
'+  2. Mode 값이 "R"일 경우 택 번호 Check를 무시한다.
'+  3. 전달값
'+     strTag :   "1-234"   5자리 전달
'+     Mode   :   "R"       DB에서 새로 읽는다.
'+     Mode   :   "+"       택번호를 1 증가한다
'+     Mode   :   "-"       택번호를 1 감소한다
'+
'+------------------------------------------------------
Public Function GetTagNum(strTag As String, Mode As String) As String
    Dim bTagDayCheck  As Boolean
    Dim bTagFormCheck As Boolean
    Dim TagNum        As String
    
    ' 전달된 Mode  검사
    If Len(Mode) <> 1 And Not (Mode = "+" Or Mode = "-") Then
        MsgBox " 전달된 Mode가 올바르지 않습니다. ", "택번호 오류"
        Exit Function
    End If
    
    ' DB에 저장된 택번호를 읽는다.
    Select Case Mode
        Case "R"
            Query = "Select 할인시작일 "
            Query = Query & "From 대리점정보 "
            Set SUBRs = New ADODB.Recordset
            SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
            
            If SUBRs.EOF Or SUBRs.BOF Then
                SUBRs.Close
                
                TagNum = "0" + "-" + "000"
                frmMain.TagNo.Caption = TagNum
                GetTagNum = TagNum
                Exit Function
                
            ElseIf IsNull(SUBRs(0)) = True Then
                SUBRs.Close
                
                TagNum = "0" + "-" + "000"
                frmMain.TagNo.Caption = TagNum
                GetTagNum = TagNum
                Exit Function
            Else
                ' 형식 오류를 막기위해 2002.12.06일 val추가
                TagNum = Format(Val(Mid(SUBRs(0), 1, 1)), "@") & "-" & Format(Val(Mid(SUBRs(0), 3, 3)), "000")
                SUBRs.Close
                
                GetTagNum = TagNum
                Exit Function
            End If
    End Select
    
    ' 전달된 택 번호 검사
    If Not IsTagNum(strTag) Then
        MsgBox " 전달된 택번호가 올바르지 않습니다. ", vbInformation, "택번호 오류"
        Exit Function
    End If

    bTagDayCheck = False
    bTagFormCheck = False
    
    Do
        ' DB에서 오늘 사용한 것을 확인 (같은 번호가 있을 경우 False 리턴)
        bTagDayCheck = dayTagchk(strTag)
        
        
        ' 접수 그리드에서 사용한 것을 확인 (같은 번호가 있을 경우 False 리턴)
        If bTagDayCheck = True Then bTagFormCheck = tagChk(strTag)
        
        ' DB 및 접수 그리드에 없을 경우 해당 택번호를 사용한다.
        If bTagDayCheck = True And bTagFormCheck = True Then Exit Do
        
        ' 어느 한굿이라도 사용중일 경우 택번호를 증가한다.
        strTag = GetChangeTagNumber(strTag, Mode)
    Loop
    
    GetTagNum = strTag
End Function

'--------------------------------------------------------------------------------------------------------------
' Procedure : GetChangeTagNumber
' DateTime  : 2006-10-13 14:15
' Author    : pds2004
' Purpose   : 전달된 택번호의 다음이나 이전을 무조건 리턴한다
'--------------------------------------------------------------------------------------------------------------
Public Function GetChangeTagNumber(ByVal sTag As String, ByVal Mode As String) As String
    Dim nTagNo1 As Integer
    Dim nTagNo2 As Integer
    Dim TagNum As String
    
    GetChangeTagNumber = ""
    ' 기본 변수에 저장
    nTagNo1 = CInt(Mid(sTag, 1, 1))       ' int
    nTagNo2 = CInt(Mid(sTag, 3, 3))       ' int
    
    Select Case Mode
        Case "+"
            ' 다음 택번호 부여
            If nTagNo2 >= 999 Then
                nTagNo1 = nTagNo1 + 1
                nTagNo2 = 0
            Else
                nTagNo2 = nTagNo2 + 1
            End If
            
            If nTagNo1 > 9 Then
                nTagNo1 = 0
            End If
            
            GetChangeTagNumber = Format(nTagNo1, "0") & "-" & Format(nTagNo2, "000")
            Exit Function
            
        Case "-"
            If nTagNo2 <= 0 Then
                nTagNo1 = nTagNo1 - 1
                nTagNo2 = 999
            Else
                nTagNo2 = nTagNo2 - 1
            End If
            
            If nTagNo1 < 0 Then
                nTagNo1 = 9
            End If
            
            GetChangeTagNumber = Format(nTagNo1, "0") & "-" & Format(nTagNo2, "000")
            
            Exit Function
        
        Case Else
            MsgBox " 전달된 Mode가 올바르지 않습니다. ", "택번호 오류"
            Exit Function
    End Select
End Function

'+------------------------------------------------------
'+
'+ 2003/01/22
'+
'+루틴설명
'+  1. strTag로 전달된 택번호를 택 번호 규격인지를 검사한다.
'+  2. 전달값
'+     strTag :   "1-234"   5자리 전달
'+  3. 리턴값
'+     True :       택 번호 규격에 맞을 경우
'+     Falss:       택 번호 규격에 맞지 않을 경우
'+
'+------------------------------------------------------
Public Function IsTagNum(strTag) As Boolean
    Dim bTag As Boolean

    bTag = True
    If Len(strTag) <> 5 Or Mid(strTag, 2, 1) <> "-" Then
        bTag = False
    End If
    IsTagNum = bTag
End Function

'+------------------------------------------------------
'+
'+ 2003/02/03
'+
'+루틴설명
'+  1. strPass로 전달된 비밀번호의 유효성을 검사한다
'+  2. 전달값
'+     strPass :   "05????????????"   앞 2자리는 유효 일자
'+                                       2자리 다음은 비빌번호
'+                                       ( 일자 * 1544 )
'+  3. 리턴값
'+     "OK" 정상
'+     -1 :         임의 수정한 경우
'+     -3 :         입력한 내용이 틀린 경우
'+
'+------------------------------------------------------
Public Function IsCodePassWord(strPass) As String
    Dim nday As Double
    Dim dPass As Double
    Dim strTemp As String
    
    If IsNumeric(strPass) = False Then
        MsgBox "전달된 본사확인코드의 형식이 정확하지 않습니다.", vbInformation, "입력오류"
        IsCodePassWord = "-1"
        Exit Function
    End If
    
    ' 일자 * 1544
    If Val(Format(Date, "dd")) * 2025 = CStr(strPass) Then
        IsCodePassWord = "OK"
    Else
        IsCodePassWord = "-3"
    End If
End Function

Public Function IsServicePassWord(strPass As String, strMCode As String) As String
'+------------------------------------------------------
'+
'+ 2003/02/03
'+
'+루틴설명
'+  1. IsPricPassWord 전달된 비밀번호의 유효성을 검사한다
'+  2. 전달값
'+     strPass :    전달된 확인 번호
'+     strTag  :    전달된 고객번호 ( 승인 고객번호 )
'+  3. 리턴값
'+     "-1"    :    오류 값
'+     -3 :         입력한 내용이 틀린 경우
'+     "고객번호":    허가받은 고객번호를 리턴한다.
'+------------------------------------------------------
Dim nday As Double
Dim dPass As Double
Dim strTemp As String

    If Len(strMCode) <= 0 Then
        MsgBox "전달된 고객번호 형식이 정확하지 않습니다.", vbInformation, "입력오류"
        IsServicePassWord = "-1"
        Exit Function
    End If
    
    strTemp = Trim(strMCode)
    
    ' 오늘의 일자를 구한다.
    nday = Format(Date, "dd")
    dPass = Left(Format(nday * Val(strTemp) * 1544, "@@@@@@@@"), 8)
    
    If Val(strPass) = dPass Then
        IsServicePassWord = strMCode
    Else
        IsServicePassWord = "-3"
    End If
    
End Function

Public Function IsPricPassWord(strPass As String, strTag As String) As String
'+------------------------------------------------------
'+
'+ 2003/02/03
'+
'+루틴설명
'+  1. IsPricPassWord 전달된 비밀번호의 유효성을 검사한다
'+  2. 전달값
'+     strPass :    전달된 확인 번호
'+     strTag  :    전달된 택번호 ( 승인 택번호 )
'+  3. 리턴값
'+     "-1"    :    오류 값
'+     -3 :         입력한 내용이 틀린 경우
'+     "택번호":    허가받은 택번호를 리턴한다.
'+------------------------------------------------------
    Dim nday As Double
    Dim dPass As Double
    Dim strTemp As String

    If Not IsTagNum(strTag) Then
        MsgBox "전달된 택번호 형식이 정확하지 않습니다.", vbInformation, "입력오류"
        IsPricPassWord = "-1"
        Exit Function
    End If
    
    strTemp = Mid(strTag, 3, Len(strTag))
    
    ' 오늘의 일자를 구한다.
    nday = Format(Date, "dd")
    dPass = nday * Val(strTemp) * 1544
    
    If Val(strPass) = dPass Then
        IsPricPassWord = strTag
    Else
        IsPricPassWord = "-3"
    End If
End Function

Public Function IsPassWord(strPass) As String
'+------------------------------------------------------
'+
'+ 2003/02/03
'+
'+루틴설명
'+  1. strPass로 전달된 비밀번호의 유효성을 검사한다
'+  2. 전달값
'+     strPass :   "05????????????"   앞 2자리는 유효 일자
'+                                       2자리 다음은 비빌번호
'+                                       ( 일자 * 365 * 7079 )
'+  3. 리턴값
'+     앞 2자리를 리턴한다. ( 사용기간 )
'+     -1 :         임의 수정한 경우
'+     -3 :         입력한 내용이 틀린 경우
'+
'+------------------------------------------------------
    Dim nday As Double
    Dim dPass As Double
    Dim strTemp As String

    If Mid(strPass, 1, 1) < "0" Or Mid(strPass, 1, 1) > "9" Or Mid(strPass, 2, 1) < "0" Or Mid(strPass, 2, 1) > "9" Then
        MsgBox "전달된 본사확인코드의 형식이 정확하지 않습니다.", vbInformation, "입력오류"
        IsPassWord = "-1"
        Exit Function
    End If
    strTemp = Mid(strPass, 3, Len(strPass) - 2)
    ' 오늘의 일자를 구한다.
    nday = Format(Date, "dd")
    dPass = nday * Val(Format(Date, "mm")) * 7079
    If strTemp = dPass Then
        IsPassWord = Mid(strPass, 1, 2)
    Else
        IsPassWord = "-3"
    End If
    
End Function

Public Function IsPassREGRead() As String
'+------------------------------------------------------
'+
'+ 2003/02/03
'+
'+루틴설명
'+  1. 레지스터리의 PassWord저장 내용 임의 변경을 확인한다.
'+  2. 리턴값
'+     -1 :         임의 수정한 경우
'+     -2 :         유효 기간이 만료된 경우
'+     -3 :         입력한 내용이 틀린 경우
'+     일자 :       완료 일자를 리턴한다.
'+
'+------------------------------------------------------
Dim strPass As String
Dim strtemp1 As String
Dim strtemp2 As String
Dim strtemp3 As String

'----------------------------------------------------------------------------------------
' 개발자에게 짜증나기 때문에 개발자는 화면이 안뜨도록 하기 위하여 다음 5중을 추가한다.
    If GetSetting("Laundry_Zi", "PassWord", "Administrator", strtemp1) = "ALL" Then
' Administrator 코드 저장
' SaveSetting "Laundry_Zi", "PassWord", "Administrator", "ALL"
        IsPassREGRead = ""
        Exit Function
    End If
'----------------------------------------------------------------------------------------
   
    strtemp1 = GetSetting("Laundry_Zi", "PassWord", "PassVal1", strtemp1)
    strtemp2 = GetSetting("Laundry_Zi", "PassWord", "PassVal2", strtemp2)
    strtemp3 = GetSetting("Laundry_Zi", "PassWord", "PassVal3", strtemp3)
    
    If Len(strtemp1) <= 0 Or Len(strtemp2) <= 0 Or Len(strtemp3) <= 0 Then
        IsPassREGRead = "-1"
        Exit Function
    End If

    
    strPass = Mid(strtemp1, 1, 2)
    ' 일자 확인
    If strPass < "00" And strPass > "99" Then
        IsPassREGRead = "-1"
        Exit Function
    End If
    
    
    ' 시작 일자를 확인한다.
    strtemp2 = Val(strtemp2) / Val(Format(Date, "mm")) / 12
    ' 저장비밀번호에서 시작일을 구한다.
    strPass = Mid(strtemp1, 3, Len(strtemp1)) / Val(Format(Date, "mm")) / 7079 '1544
    
    If strPass <> strtemp2 Then
        IsPassREGRead = "-1"
        Exit Function
    Else
        '확인코드의 유효 기간을 확인한다.
        strtemp3 = Val(strtemp3) / Val(Format(Date, "mm")) / 12
        If strtemp3 > Format(Date, "yyyymmdd") Then
            IsPassREGRead = strtemp3
            Exit Function
        Else
            IsPassREGRead = "-2"
        End If
    End If

End Function

Public Function IsPassREGSave(temp1 As String) As Boolean
'+------------------------------------------------------
'+
'+ 2003/02/03
'+
'+루틴설명
'+  temp1 : 유효 일자
'+  temp2 : 비밀번호
'+  1. 레지스터리에 PassWord의 내용을 저장한다.
'+  2. 리턴값
'+     True :       정상인 경우
'+     Falss:       저장 오류인 경우
'+
'+------------------------------------------------------
    Dim nday As Integer
    Dim dPass As Double
    Dim strTemp As String

    If Mid(temp1, 1, 2) < "00" Or Mid(temp1, 1, 2) > "99" Then
        MsgBox "전달된 비밀번호 형식이 정확하지 않습니다.", vbInformation, "입력오류"
        IsPassREGSave = False
        Exit Function
    End If
    
    ' 확인코드 저장
    SaveSetting "Laundry_Zi", "PassWord", "PassVal1", temp1
    ' 확인 코드의 시작일 저장
    strTemp = Format(Date, "dd") * Val(Format(Date, "mm")) * 12
    SaveSetting "Laundry_Zi", "PassWord", "PassVal2", strTemp
    ' 확인 코드의 유효기간을 "yyyymmdd"형식으로 저장
    strTemp = Format(DateAdd("d", Val(Mid(temp1, 1, 2)), Date), "yyyymmdd")
    strTemp = strTemp * Val(Format(Date, "mm")) * 12
    SaveSetting "Laundry_Zi", "PassWord", "PassVal3", strTemp
    IsPassREGSave = True

End Function

Public Function IsEventPassWord(strPass) As String
'+------------------------------------------------------
'+
'+ 2003/08/29
'+
'+루틴설명
'+  1. strPass로 전달된 비밀번호의 유효성을 검사한다
'+  2. 전달값
'+     strPass :   "05????????????"   앞 2자리는 유효 일자
'+                                       2자리 다음은 비빌번호
'+                                       ( 일자 * 365 * 1544 )
'+  3. 리턴값
'+     앞 2자리를 리턴한다. ( 사용기간 )
'+     -1 :         임의 수정한 경우
'+     -3 :         입력한 내용이 틀린 경우
'+
'+------------------------------------------------------
    Dim nday As Double
    Dim intMM As Integer
    Dim dPass As Double
    Dim strTemp As String
    
    If IsNull(Mid(strPass, 1, 2)) Then
        MsgBox "전달된 본사확인코드의 형식이 정확하지 않습니다.", vbInformation, "입력오류"
        IsEventPassWord = "-1"
        Exit Function
    End If
        
    strTemp = Mid(strPass, 3, Len(strPass) - 2)
    
    ' 오늘의 일자를 구한다.
    nday = Val(Format(Date, "dd"))
    intMM = Val(Format(Date, "mm"))
    dPass = nday * intMM * 1544
    
    If strTemp = CStr(dPass) Then
        IsEventPassWord = Mid(strPass, 1, 2)
    Else
        IsEventPassWord = "-3"
    End If
    
End Function

Public Function IsEventPassREGSave(temp1 As String) As Boolean
'+------------------------------------------------------
'+
'+ 2003/08/29
'+
'+루틴설명
'+  temp1 : 유효 일자
'+  temp2 : 비밀번호
'+  1. 레지스터리에 PassWord의 내용을 저장한다.
'+  2. 리턴값
'+     True :       정상인 경우
'+     Falss:       저장 오류인 경우
'+
'+------------------------------------------------------
    Dim intMM As Integer
    Dim dPass As Double
    Dim strTemp As String

    If Mid(temp1, 1, 2) < "00" Or Mid(temp1, 1, 2) > "99" Then
        MsgBox "전달된 비밀번호 형식이 정확하지 않습니다.", vbInformation, "입력오류"
        IsEventPassREGSave = False
        Exit Function
    End If
    
    intMM = Val(Format(Date, "mm"))
    ' 확인코드 저장
    SaveSetting "Laundry_Zi", "PassWord", "EventPassVal1", temp1
    
    ' 확인 코드의 시작일 저장
    strTemp = Format(Date, "dd") * intMM * 12
    SaveSetting "Laundry_Zi", "PassWord", "EventPassVal2", strTemp
    
    ' 확인 코드의 유효기간을 "yyyymmdd"형식으로 저장
    strTemp = Format(DateAdd("d", Val(Mid(temp1, 1, 2)), Date), "yyyymmdd")
    strTemp = strTemp * intMM * 12
    SaveSetting "Laundry_Zi", "PassWord", "EventPassVal3", strTemp
    
    IsEventPassREGSave = True

End Function

'+------------------------------------------------------
'+
'+ 2003/08/29
'+
'+루틴설명
'+  1. 레지스터리의 PassWord저장 내용 임의 변경을 확인한다.
'+  2. 리턴값
'+     -1 :         임의 수정한 경우
'+     -2 :         유효 기간이 만료된 경우
'+     -3 :         입력한 내용이 틀린 경우
'+     일자 :       완료 일자를 리턴한다.
'+
'+------------------------------------------------------
Public Function IsEventPassREGRead() As String
    Dim strPass As String
    Dim strtemp1 As String
    Dim strtemp2 As String
    Dim strtemp3 As String
    Dim intMM    As Integer
   
    strtemp1 = GetSetting("Laundry_Zi", "PassWord", "EventPassVal1", strtemp1)
    strtemp2 = GetSetting("Laundry_Zi", "PassWord", "EventPassVal2", strtemp2)
    strtemp3 = GetSetting("Laundry_Zi", "PassWord", "EventPassVal3", strtemp3)
    
    If Len(strtemp1) <= 0 Or Len(strtemp2) <= 0 Or Len(strtemp3) <= 0 Then
        IsEventPassREGRead = "-1"
        chkEventSale = False
        Exit Function
    End If
    
    intMM = Val(Format(Date, "MM"))

    
    strPass = Mid(strtemp1, 1, 2)
    ' 일자 확인
    If strPass < "00" And strPass > "99" Then
        IsEventPassREGRead = "-1"
        chkEventSale = False
        Exit Function
    End If
    
    
    ' 시작 일자를 확인한다.
    strtemp2 = Val(strtemp2) / intMM / 12
    
    ' 저장비밀번호에서 시작일을 구한다.
    strPass = Mid(strtemp1, 3, Len(strtemp1)) / intMM / 1544
    
    If strPass <> strtemp2 Then
        IsEventPassREGRead = "-1"
        Exit Function
    Else
        '확인코드의 유효 기간을 확인한다.
        strtemp3 = Val(strtemp3) / intMM / 12
        If strtemp3 > Format(Date, "yyyymmdd") Then
            IsEventPassREGRead = strtemp3
            Exit Function
        Else
            IsEventPassREGRead = "-2"
            chkEventSale = False
        End If
    End If

End Function

'+------------------------------------------------------
'+
'+ 2003/02/07
'+
'+루틴설명
'+  프로그램 실행에 필요한 기본 디렉토리를 확인하여 없을 경우 생성한다.
'+  1. 리턴값
'+     Counter :    생성한 디렉토리 수 를 리턴한다
'+
'+------------------------------------------------------
'+  DB          - db 저장 폴더
'+  Image       - 기본 이미지 저장 파일
'+  RecvData    - 수신자료 저장 폴더
'+  BackData    - 임시 저장 폴더
'+  prg         - 프로그램 업그레이드시 필요
Public Function DirectoryCheck() As Integer
    Dim nCount As Integer


    If Dir(Trim(App.Path & "\DB"), vbDirectory) = "" Then
       MkDir App.Path & "\DB"
       nCount = nCount + 1
    End If
    If Dir(Trim(App.Path & "\Image"), vbDirectory) = "" Then
       MkDir App.Path & "\Image"
       nCount = nCount + 1
    End If
    If Dir(Trim(App.Path & "\RecvData"), vbDirectory) = "" Then
       MkDir App.Path & "\RecvData"
       nCount = nCount + 1
    End If
    If Dir(Trim(App.Path & "\BackData"), vbDirectory) = "" Then
       MkDir App.Path & "\BackData"
       nCount = nCount + 1
    End If
    If Dir(Trim(App.Path & "\CleanPrg"), vbDirectory) = "" Then
       MkDir App.Path & "\CleanPrg"
       nCount = nCount + 1
    End If
    If Dir(Trim(App.Path & "\Internet"), vbDirectory) = "" Then
       MkDir App.Path & "\Internet"
       nCount = nCount + 1
    End If
    
    DirectoryCheck = nCount
End Function

'+------------------------------------------------------
'+
'+ 2003/02/07
'+
'+루틴설명
'+  프로그램 실행에 필요한 기본 파일을 확인하여 없을 경우 생성한다.
'+  1. 리턴값
'+     FilesCheck :    없는 파일의 수 를 리턴한다 (1자리)
'+     FilesCheck :    없는 파일이름을 리턴한다 (누적)
'+
'+------------------------------------------------------
'+  DB          - db 저장 폴더
'+  Image       - 기본 이미지 저장 파일
'+  RecvData    - 수신자료 저장 폴더
'+  BackData    - 임시 저장 폴더
'+  prg         - 프로그램 업그레이드시 필요
Public Function FilesCheck() As String
    Dim nCount As Integer
    Dim strFile As String
    Dim strChkFiles
    Dim acount As Integer
    Dim i As Integer
    
    strChkFiles = Array(m_DBPath, "LAUNDRY.ini", "pkunzip.exe", "pkzip.exe", "Restore.bat", "PGDOWN.bat", "Backup.bat", "arj.exe")
    strFile = ","
    
    For i = 1 To UBound(strChkFiles)
        If Dir(Trim(App.Path & "\" & strChkFiles(i)), vbDirectory) = "" Then
            strFile = strFile & strChkFiles(i) & ","
            nCount = nCount + 1
        End If
    Next i
    
    '리턴값 설정
    If nCount > 9 Then nCount = 9
    FilesCheck = Format(nCount, "0") & Mid(strFile, 1, Len(strFile) - 1)
    
End Function

Public Sub Delay(ByVal utime As Single)
    Dim starttime As Single

    starttime = Timer
 
    Do
        'DoEvents
    Loop While Timer < starttime + utime
End Sub

Public Function GetCustomNo() As String
    Dim CustomNo As String

    Query = "SELECT max(고객번호) FROM 고객정보 "
    Query = Query & " WHERE 고객번호 Like '" & Mid(CStr(Year(Date)), 3, 2) & "%' "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic

    If IsNull(SUBRs(0)) Then
        CustomNo = Mid(CStr(Year(Date)), 3, 2) & "0000"
    Else
        CustomNo = Mid(CStr(Year(Date)), 3, 2) & Right("0000" & CStr(CDbl(SUBRs(0)) + 1), 4)
    End If
    SUBRs.Close
    Set SUBRs = Nothing
    
    GetCustomNo = CustomNo
End Function

Public Sub SetProgramMode()
'+------------------------------------------------------
'+
'+ 2003/02/13
'+
'+루틴설명
'+  프로그램이 어떤 모드로 실행될지를 결정한다.
'+  laundry.ini 파일에서 읽어서 결정한다.
'+  [RUNMODE]
'+  ProgramMode = 1
'+  ' 1이면 서버            - 입고기능 가능
'+  ' 1이아니면 클라이언트  - 입고기능 불가
'+  DBPath = "G:"           ' 프로그램 실행 모드를 확인한다.
    Dim Filename As String
    
    Filename = Dir(iniFile)
    
    ' Laundry.ini 파일이 없을 경우 종료 한다.
    If Filename = "" Then
        End
    Else
        chkProgramMode = GetIniStr("RUNMODE", "ProgramMode", "", iniFile)
        If chkProgramMode = "" Then
            ' 파일은 있지만 서버 설정이 없을 경우 서버로 자동 설정한다.
            chkProgramMode = ServerMode
            ' 기본 내용을 출력한다.
            Open iniFile For Append As #1
            Print #1, "  "
            Print #1, "[RUNMODE]"
            Print #1, "ProgramMode = 1"
            Print #1, "' 1이면 서버        - 입고기능 가능"
            Print #1, "' 1이아니면 클라이언트  - 입고기능 불가"
            Print #1, "DBPath = ""G:"""
            Close
        End If
    End If
End Sub

Public Function GetBo_Day() As String
'+------------------------------------------------------
'+
'+ 2003/03/10
'+
'+루틴설명
'+  보관증을 보존할 일자를 리턴한다.
'+
'+------------------------------------------------------
Dim strTemp As String

    If Val(GetSetting("Laundry_Zi", "Printer", "Bo_Day", strTemp)) = 0 Then
        strTemp = 0
    Else
        strTemp = Val(GetSetting("Laundry_Zi", "Printer", "Bo_Day", strTemp))
    End If
    
    If Val(strTemp & "") > 30 Then
        strTemp = "30"
    ElseIf Val(strTemp & "") <= 0 Then
        strTemp = "30"
    End If

    GetBo_Day = Format(strTemp, "00")

End Function
'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : GetColWidth
' 작  성  자  : IT21
' 작  성  일  : 2000.06.07
' 파 라 미 터 : New_App  -
'               New_Form -
'               New_SS   -
' 리  턴  값  : Boolean
' 비      고  : 스프레드의 Column의 길이를 불러온다.
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Function GetColWidth(ByVal New_App As String, ByVal New_Form As String, ByVal New_SS As Object) As Boolean
    Dim Col As Long

On Error GoTo GetColWidth_Err:
    
    GetColWidth = True
    For Col = 1 To New_SS.MaxCols
        New_SS.ColWidth(Col) = GetSetting(New_App, New_Form, New_SS.Name + "_ColWidth_" + CStr(Col), CStr(New_SS.ColWidth(Col)))
    Next Col
    
    Exit Function

GetColWidth_Err:
    GetColWidth = False
    Resume Next

End Function

'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : SaveColWidth
' 작  성  자  : IT21
' 작  성  일  : 2000.06.07
' 파 라 미 터 : New_App  -
'               New_Form -
'               New_SS   -
' 리  턴  값  : Boolean
' 비      고  : 스프레드의 Column의 길이를 저장한다.
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Function SaveColWidth(ByVal New_App As String, ByVal New_Form As String, ByVal New_SS As Object) As Boolean
    Dim Col As Long

On Error GoTo SaveColWidth_Err:
    
    SaveColWidth = True
    For Col = 1 To New_SS.MaxCols
        SaveSetting New_App, New_Form, New_SS.Name + "_ColWidth_" + CStr(Col), CStr(New_SS.ColWidth(Col))
    Next Col

    Exit Function

SaveColWidth_Err:
    SaveColWidth = False
    Resume Next
End Function


Public Sub ClearArray(fpArr As Variant)
' 2차원 배열을 초기화 한다.
    Dim iCnt As Integer
    Dim jCnt As Integer
    For iCnt = LBound(fpArr, 1) To UBound(fpArr, 1)
        If Len(fpArr(iCnt, 1)) < 1 Then
            Exit For
        End If
        For jCnt = LBound(fpArr, 2) To UBound(fpArr, 2)
            fpArr(iCnt, jCnt) = Empty
        Next jCnt
    Next iCnt
    

End Sub


Public Sub CrateCommondFiles(Mode As CommandFiles, Optional SW As Boolean = False)
    Dim HFandle As Integer
    
    
    If SW = False Then
        '일경우에만 있는지 확인하여 없을 경우만 만듣다.

        Select Case Mode
            Case Backup
                If Dir("backup.bat", vbDirectory) <> "" Then Exit Sub
            Case Restore
                If Dir("Restore.bat", vbDirectory) <> "" Then Exit Sub
            Case PGDown
                If Dir("PGDown.bat", vbDirectory) <> "" Then Exit Sub
            Case DBSend
                If Dir("DBSend.bat", vbDirectory) <> "" Then Exit Sub
        End Select
    End If
    
    ' 무조건 만든다.
    HFandle = FreeFile
    Select Case Mode
        Case Backup
            Open App.Path & "\Backup.Bat" For Output As #HFandle
            Print #HFandle, "@echo off"
            Print #HFandle, "c:\Laundry\arj a -v1440 -y c:\Laundry\db\BackData c:\Laundry\DB\LAUNDRY.MDB"
            Print #HFandle, "COPY c:\Laundry\bACKUP.BAT c:\Laundry\BACKUP.OK"
            Close #HFandle
        Case Restore
            Open App.Path & "\Restore.Bat" For Output As #HFandle
            Print #HFandle, "@echo off"
            Print #HFandle, "c:\Laundry\arj e -v1440 -y c:\Laundry\db\BackData c:\Laundry\db\"
            Print #HFandle, "COPY c:\Laundry\bACKUP.BAT c:\Laundry\BACKUP.OK"
            Close #HFandle


        Case PGDown
        Case DBSend
            Open App.Path & "\DBSend.Bat" For Output As #HFandle
            Print #HFandle, "@echo off"
            Print #HFandle, "c:"
            Print #HFandle, "cd c:\Laundry"
            Print #HFandle, "Copy c:\Laundry\DB\Laundry.MDB C:\Laundry\db\Laundry1.MDB"
            Print #HFandle, "C:\Laundry\Pkzip.exe c:\Laundry\db\" & 대리점정보.대리점번호 & ".zip C:\Laundry\DB\Laundry1.MDB"
            Print #HFandle, "COPY c:\Laundry\DBSend.BAT c:\Laundry\DBSend.OK"
            Close #HFandle
    End Select
        
        
End Sub

Public Function GetStatus(RasStatus As Long) As String
    Dim StatusString As String

    Select Case RasStatus
        Case RASCS_OpenPort
            StatusString = "포트를 OPEN 하는 중 입니다..."
        Case RASCS_PortOpened
            StatusString = "포트가 OPEN 되었습니다."
        Case RASCS_ConnectDevice
            StatusString = "디바이스에 연결하는 중입니다..."
        Case RASCS_DeviceConnected
            StatusString = "디바이스에 연결되었습니다."
        Case RASCS_AllDevicesConnected
            StatusString = "모든 디바이스에 연결되었습니다."
        Case RASCS_Authenticate
            StatusString = "사용자를 인증하고 있습니다..."
        Case RASCS_AuthNotify
            StatusString = "AuthNotify"
        Case RASCS_AuthRetry
            StatusString = "사용자 인증 재시도중..."
        Case RASCS_AuthCallback
            StatusString = "AuthCallback"
        Case RASCS_AuthChangePassword
            StatusString = "비밀번호가 잘못되었습니다."
        Case RASCS_AuthProject
            StatusString = "AuthProject"
        Case RASCS_AuthLinkSpeed
            StatusString = "AuthLinkSpeed"
        Case RASCS_AuthAck
            StatusString = "AuthAck"
        Case RASCS_ReAuthenticate
            StatusString = "ReAuthenticate"
        Case RASCS_Authenticated
            StatusString = "사용자가 인증되었습니다."
        Case RASCS_PrepareForCallback
            StatusString = "PrepareForCallback"
        Case RASCS_WaitForModemReset
            StatusString = "WaitForModemReset"
        Case RASCS_WaitForCallback
            StatusString = "WaitForCallback"
        Case RASCS_Projected
            StatusString = "Projected"
        Case RASCS_StartAuthentication
            StatusString = "사용자 인증을 시작하는 중 입니다."
        Case RASCS_CallbackComplete
            StatusString = "CallbackComplete"
        Case RASCS_LogonNetwork
            StatusString = "네트워크에 로그온이 되었습니다."
        Case RASCS_Interactive
            StatusString = "상호네트워크 Checking 중..."
        Case RASCS_RetryAuthentication
            StatusString = "사용자 인증을 다시 하시기 바랍니다."
        Case RASCS_CallbackSetByCaller
            StatusString = "CallbackSetByCaller"
        Case RASCS_PasswordExpired
            StatusString = "비밀번호가 만료되었습니다."
        Case RASCS_Connected
            StatusString = "본사에 접속이 되었습니다."
        Case RASCS_Disconnected
            StatusString = "본사에 접속이 되지 않았습니다."
        Case 0
            StatusString = "본사에 접속이 되지 않았습니다."
        Case RASBASE
            StatusString = "접속 중..."
        Case Else
            StatusString = "RAS Error " & RasStatus
    End Select
    
    GetStatus = StatusString

End Function

Public Function ProgramUpgrade() As Boolean
'+++++++++++++++++++++++++++++++++++++++++++++++++
' 작성일 : 2003/04/15
' 작성자 : pds2004 (박대선)
' 설  명 : 프로그램을 업그레이드 한다.
' 리턴값 : 0 정상실행
'          1 업그레이드 성공
'          -1 업그레이드 실패
'+++++++++++++++++++++++++++++++++++++++++++++++++
Dim PauseTime, Start, Finish, TotalTime
Dim SourceFile As String
Dim DestinationFile As String


SourceFile = App.Path & "\" & App.EXEName & ".exe" ' 판매재고관리UP.EXE
DestinationFile = App.Path & "\" & Mid(App.EXEName, 1, Len(App.EXEName) - 2) & ".exe" '판매재고관리.EXE

If UCase(Right(App.EXEName, 2)) = "UP" Then
' 업그레이드 파일일경우]
    On Error Resume Next
    PauseTime = 10                  ' 기간을 지정합니다.
    Start = Timer                   ' 시작 시간을 지정합니다.
    Do
        Finish = Timer              ' 종료 시간을 지정합니다.
        TotalTime = Finish - Start  ' 전체 시간을 계산합니다.
        If Dir(DestinationFile, vbDirectory) <> "" Then
            If Timer > Start + PauseTime Then
                MsgBox " 정상적으로 업그레이드 되지 않았습니다." & vbLf & "다시 시도 합니다.", vbCritical, "오류"
                On Error GoTo kill_end
                Kill DestinationFile
            End If
        Else
            Exit Do
        End If
        DoEvents                    ' 다른 프로시저로 넘깁니다.
        Kill DestinationFile        ' 원본 파일을 삭제한다.
    Loop
    
kill_end:
    If Dir(DestinationFile, vbDirectory) = "" Then
        FileCopy SourceFile, DestinationFile
        MsgBox " 업그레이드 완료. " & vbLf & vbLf & " 프로그램을 다시 시작 하여 주십시요.     ", vbInformation, "확인"
        
        Call Fb대리점정보
        Call SendProgramVersion

        End
        
        Exit Function
    Else
        MsgBox " 정상적으로 업그레이드 되지 않았습니다.", vbCritical, "오류"
        ProgramUpgrade = False
        Exit Function
    End If
    
Else
    ProgramUpgrade = True
End If
    
End Function

Public Function Fnc_UserMileage(UserCode As String) As Boolean
' 회원 코드를 전송 받아 마일리지 정보를 리턴한다.
    Dim UserMoney   As Long
    Dim EndMoney    As Long
    Dim EndMileage  As Long
    
    On Error GoTo Err_Rtn
    
RE_CHECK:
    Fnc_UserMileage = False
    
    '-----------------------------------------------------------
    '
    '-----------------------------------------------------------
    Query = "SELECT * FROM 마일리지현황 "
    Query = Query & " WHERE 고객번호 = '" & UserCode & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If Not Rs.EOF Or Not Rs.BOF Then Rs.MoveLast
    
    If Rs.RecordCount = 1 Then
        Fnc_UserMileage = True
        
        With userMileage
            .검색여부 = True
            .총사용금액 = Val(Rs!총사용금액 & "")
            .잔액 = Val(Rs!마일리지 & "")
            .최종발생금액 = Val(Rs!최종발생금액 & "")
            .발생총누계 = Val(Rs!발생누계 & "")
            .사용누계 = Val(Rs!사용마일리지 & "")
            .미반환마일리지 = Val(Rs!미반환마일리지 & "")
        End With
        
'   FormMileageCheck 폼에서 일괄 처리하도록 변경
'        ' 이전 프로그램 오류로 인하여 사용금액이 -금액이 들어가 있고. 그것으로 인하여 오동작하는것을 방지
'        If userMileage.잔액 < 0 Then
'            Query = "SELECT SUM(금액) AS UserMoney FROM 입출고 WHERE 고객번호 = '" & UserCode & "' AND 입고일 >= '20050701' "
'            Set RS2 = MyDB.OpenRecordset(Query)
'            UserMoney = RS2.Fields("UserMoney")
'            RS2.Close
'
'            Query = "SELECT SUM(금액) AS UserMoney FROM 입출고 WHERE 고객번호 = '" & UserCode & "' AND 입고일 >= '20050701' AND 판매취소 = 'R'"
'            Set RS2 = MyDB.OpenRecordset(Query)
'            UserMoney = UserMoney - Val(RS2.Fields("UserMoney") & "")
'
'            If UserMoney < 100000 Then
'                EndMoney = 0:           EndMileage = 0
'            End If
'            If UserMoney >= 100000 Then
'                EndMoney = 100000:      EndMileage = 3000
'            End If
'            If UserMoney >= 200000 Then
'                EndMoney = 200000:      EndMileage = EndMileage + 4000
'            End If
'            If UserMoney >= 300000 Then
'                EndMoney = 300000:      EndMileage = EndMileage + 5000
'            End If
'            If UserMoney >= 400000 Then
'                EndMoney = 400000:      EndMileage = EndMileage + 6000
'            End If
'            If UserMoney >= 500000 Then
'                EndMoney = 500000:      EndMileage = EndMileage + 7000
'            End If
'
'            Query = " UPDATE 마일리지현황 SET "
'            Query = Query & " 총사용금액 = " & UserMoney & ", "
'            Query = Query & " 마일리지 =  " & EndMileage & ", "
'            Query = Query & " 최종발생금액 =" & EndMoney & ", "
'            Query = Query & " 발생누계 = " & EndMileage & ", "
'            Query = Query & " 사용마일리지 = 0, "
'            Query = Query & " 미반환마일리지 = 0, "
'            Query = Query & " 전송여부 = 'N' "
'            Query = Query & " WHERE 고객번호 = '" & UserCode & "'"
'            ADOCon.Execute Query
'
'            GoSub RE_CHECK
'        End If
'
        
    Else
        Fnc_UserMileage = False
        
        With userMileage
            .검색여부 = False
            .총사용금액 = 0
            .잔액 = 0
            .최종발생금액 = 0
            .발생총누계 = 0
            .사용누계 = 0
            .미반환마일리지 = 0
        End With
    End If
    
    Rs.Close
    Set Rs = Nothing
    
    Exit Function

Err_Rtn:
    Fnc_UserMileage = False
    
    MsgBox Err.Description, vbInformation, "확인"
End Function

Public Function Fnc_MileagePoint(UserMoney As Double, LastMoney As Double, UserCode As String) As Double
' 총사용금액,  최종발생금액,  고객번호
' 사용금액이 전달되면 거기에 해당하는 마일리지 금액이 전달된다.
' 한번에 여러 단계가 발생 할수 있다.

' NextMileage = 100,000

    Dim TempMileage As Double
    Dim j As Double
    Dim kk As Double
    
    If LastMoney < 0 Then Exit Function
    ' 마일리지 시작 값을 구한다.
    kk = ((LastMoney \ NextMileage) + 1) * NextMileage
    
    TempMileage = 0
    
    If 대리점정보.마일리지증가구분 = "0" Then
        For j = kk To UserMoney Step NextMileage
        
            '   500,000 이상일 경우
            If j >= (NextMileage * 5) Then
                TempMileage = TempMileage + 7000
                
            '   400,000 이상일 경우
            ElseIf j >= (NextMileage * 4) Then
                TempMileage = TempMileage + 6000
            
            '   300,000 이상일 경우
            ElseIf j >= (NextMileage * 3) Then
                TempMileage = TempMileage + 5000
            
            '   200,000 이상일 경우
            ElseIf j >= (NextMileage * 2) Then
                TempMileage = TempMileage + 4000
            
            '   100,000 이상일 경우
            ElseIf j >= (NextMileage * 1) Then
                TempMileage = TempMileage + 3000
            
            '   100,000 미만일 경우
            ElseIf j < NextMileage Then
                TempMileage = TempMileage + 0
            End If
        Next j
    
    ElseIf 대리점정보.마일리지증가구분 = "1" Then
        For j = kk To UserMoney Step NextMileage
        
'            '   500,000 이상일 경우
'            If j >= (NextMileage * 5) Then
'                TempMileage = TempMileage + 7000
'
'            '   400,000 이상일 경우
'            ElseIf j >= (NextMileage * 4) Then
'                TempMileage = TempMileage + 6000
'
'            '   300,000 이상일 경우
'            ElseIf j >= (NextMileage * 3) Then
'                TempMileage = TempMileage + 5000
'
'            '   200,000 이상일 경우
'            ElseIf j >= (NextMileage * 2) Then
'                TempMileage = TempMileage + 4000
            
            '   100,000 이상일 경우
            If j >= (NextMileage * 1) Then
                TempMileage = TempMileage + 3000
            
            '   100,000 미만일 경우
            ElseIf j < NextMileage Then
                TempMileage = TempMileage + 0
            End If
        Next j
    
    End If
    
    Fnc_MileagePoint = TempMileage
    
End Function

' 전달된 회원의 전체 매출및 마일리지 관련 내용을 정리한다.
' 마일리지 사용금액이 남아 있을 경우 해당 마일리지를 삭제하고 그렇지 않을 경우
' 남아 있는 마일리지만 삭제한후 나중에 마일리지 발생할경우 삭제한 마일리지 만큼만 발생 처리하여 준다.
Public Function Fnc_MileageEdit(UserCode As String, UserDate As String, UserMoney As Double) As Boolean
'   1. 해당 이용실적을 삭제 한다. ( 이용실적/ 마일리지 현황)
'   1. 최종 마일리지사용 금액에서 해당 금액을 차감후 마일리지 변동 여부 확인
'   2. 마일리지가 변동될경우 마일리지 잔액이 남아 있는지 확인하여 남아 있을경우 바로 차감한다.
'   3. 차감할 금액의 일부만 남아 있을경우 해당 금액을 차감한후 미반환마일리지에 차감하지 못한 잔액을 기록한다.

' ******* 문제점 취소된 물건이 100,000만원 이상인 경우 정확한 계산이 안된다. *****
    
    Dim CnMileage As Double
    
    On Error GoTo Err_Rtn
    
    Fnc_MileageEdit = True
    
' 이쪽에서 적용하지 않는다.( 마일리지 사용하지 않는 매장을 위하여)
'    Query = "UPDATE 이용실적 SET "
'    Query = Query & " 이용금액 = (이용금액 - " & UserMoney & ") "
'    Query = Query & " WHERE 고객번호 = '" & UserCode & "'"
'    Query = Query & "   AND 연도 = '" & Left(UserDate, 4) & "'"
'    ADOCon.Execute Query
    
    ' 현재 고객의 마일리지 정보를 가저온다.
    Call Fnc_UserMileage(UserCode)
    
    ' 금액을 환불하였어도 마일리지에 영향이 없을 경우
    If (userMileage.총사용금액 - UserMoney) >= userMileage.최종발생금액 Then
        Query = " UPDATE 마일리지현황 SET "
        Query = Query & " 총사용금액 = (총사용금액 - " & UserMoney & "), "
        Query = Query & " 전송여부 = 'N' "
        Query = Query & " WHERE 고객번호 = '" & UserCode & "'"
        ADOCon.Execute Query
        
        Exit Function
    
    ' 판매 취소하여 -가 나오는 경우
    ElseIf (userMileage.총사용금액 - UserMoney) < 0 Then
        Query = " UPDATE 마일리지현황 SET "
        Query = Query & " 총사용금액 = 0, "
        Query = Query & " 최종발생금액 = 0, "
        Query = Query & " 전송여부 = 'N' "
        Query = Query & " WHERE 고객번호 = '" & UserCode & "'"
        ADOCon.Execute Query
        
        Exit Function
    
    
    ' 마일리지를 반환해야 할경우
    Else
        ' 반환할 마일리지를 계산한다.
        If 대리점정보.마일리지증가구분 = "0" Then
            CnMileage = IIf((userMileage.최종발생금액 / NextMileage) >= 5, 7000, (((userMileage.최종발생금액 / NextMileage) * 1000) + 2000))
        ElseIf 대리점정보.마일리지증가구분 = "1" Then
            CnMileage = 3000
        End If
        
        '마일리지 잔액이 남아 있는경우
        If userMileage.잔액 >= CnMileage Then
            Query = " UPDATE 마일리지현황 SET "
            Query = Query & " 총사용금액 = (총사용금액 - " & UserMoney & "), "
            Query = Query & " 마일리지 = (마일리지 - " & CnMileage & "), "
            Query = Query & " 발생누계 = (발생누계 - " & CnMileage & "), "
            Query = Query & " 최종발생금액 = (최종발생금액 - " & NextMileage & "), "
            Query = Query & " 전송여부 = 'N' "
            Query = Query & " WHERE 고객번호 = '" & UserCode & "'"
            ADOCon.Execute Query
            
            Query = "INSERT INTO 마일리지스토리 (발생일자, 고객번호, 발생마일리지, 사용마일리지, 삭제마일리지, 반환마일리지, 보관증, 전송여부)"
            Query = Query & " VALUES ('" & Format(Now, "yyyymmddhhmmss") & "', '" & UserCode & "', 0, 0, 0, " & CnMileage & ", '0', 'N') "
            ADOCon.Execute Query
            
            Exit Function
        
        ' 잔액이 부족하거나 없는경우
        Else
            Query = " UPDATE 마일리지현황 SET "
            Query = Query & " 총사용금액 = (총사용금액 - " & UserMoney & "), "
            Query = Query & " 마일리지 =  0, "
            Query = Query & " 최종발생금액 = (최종발생금액 - " & NextMileage & "), "
            Query = Query & " 발생누계 = (발생누계 - " & userMileage.잔액 & "), "
            Query = Query & " 미반환마일리지 = " & (CnMileage - userMileage.잔액) & ", "
            Query = Query & " 전송여부 = 'N' "
            Query = Query & " WHERE 고객번호 = '" & UserCode & "'"
            ADOCon.Execute Query
            
            Query = "INSERT INTO 마일리지스토리 (발생일자, 고객번호, 발생마일리지, 사용마일리지, 삭제마일리지, 반환마일리지, 보관증, 전송여부)"
            Query = Query & "VALUES ('" & Format(Now, "yyyymmddhhmmss") & "', '" & UserCode & "', 0, 0, 0, " & (CnMileage - userMileage.잔액) & ", '0', 'N') "
            ADOCon.Execute Query
            
            Exit Function
        End If
    End If
    
    Exit Function

Err_Rtn:
    Fnc_MileageEdit = False
    
    MsgBox Err.Description & Space(10), vbCritical, "오류"
End Function

'전달된 고객의 미수금을 적용한다.
Public Function Fnc_MiSuEdit(UserCode As String, UserMiSu As Double, Mode As String) As Double
' DB의 미수금은 텍스트 타입이다 ㅡㅡ
' 이전 미수금액을 구한다음 그걸 이용한다

    Dim TempMiSu As Double
    
    Fnc_MiSuEdit = -1
    
    '-----------------------------------------------------------------
    Query = "SELECT 미수금 FROM 고객정보 "
    Query = Query & "WHERE 고객번호 = '" & UserCode & "' "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If Not SUBRs.EOF Or Not SUBRs.BOF Then
        If SUBRs!미수금 <> "" Then TempMiSu = CLng(SUBRs!미수금 & "")
    End If
    SUBRs.Close
    Set SUBRs = Nothing
    
    If UCase(Mode) = "ADD" Then
        TempMiSu = TempMiSu + UserMiSu
    
    ElseIf UCase(Mode) = "DELETE" Then
        If (TempMiSu - UserMiSu) <= 0 Then
            TempMiSu = 0
        Else
            TempMiSu = TempMiSu - UserMiSu
        End If
    End If
    
    '-----------------------------------------------------------------
    Query = "UPDATE 고객정보 "
    Query = Query & " SET 미수금 = '" & CStr(TempMiSu) & "' "
    Query = Query & " WHERE 고객번호 = '" & UserCode & "' "
    ADOCon.Execute Query

    Fnc_MiSuEdit = TempMiSu

End Function


'전달된 고객의 미수금의 히스토리를 저장한다..
Public Function Fnc_MiSuHiStory(UserCode As String, UserMiSu As Double) As Boolean
    Query = "INSERT INTO 미수회수정보 (일자, 고객코드, 시간, 금액, 비고) "
    Query = Query & "VALUES('" & Format(Date, "yyyyMMdd") & "','" & UserCode & "','" & Format(Time, "hhmmss") & "',"
    Query = Query & UserMiSu & ",' ')"
    ADOCon.Execute Query
End Function

' 마일리지 마감 ( 100일동안 이용 실적이 없을 경우 마일리지 삭제 )
Public Function Fnc_MileageLastUserDelete() As Long
    Dim TempCnt As Long
    
    TempCnt = 0
    Fnc_MileageLastUserDelete = 0
    
    Query = "SELECT * FROM 마일리지현황 "
    Query = Query & " WHERE 최종거래일자 < '" & Format(DateAdd("d", -100, Date), "yyyymmdd") & "'"
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    Do While Not SUBRs.EOF
        Query = "UPDATE 마일리지현황 SET "
        Query = Query & " 마일리지 = 0 ,"
        Query = Query & " 전송여부 = 'N' "
        Query = Query & " WHERE 고객번호 = '" & SUBRs!고객번호 & "'"
        ADOCon.Execute Query
        
        Query = "INSERT INTO 마일리지스토리 (발생일자, 고객번호, 발생마일리지, 사용마일리지, 삭제마일리지, 보관증, 전송여부)"
        Query = Query & " VALUES ('" & Format(Now, "yyyymmddhhmmss") & "', '" & SUBRs!고객번호 & "', 0, "
        Query = Query & " 0, " & SUBRs!마일리지 & ", ' ', 'N') "
        ADOCon.Execute Query
        
        TempCnt = TempCnt + 1
        
        SUBRs.MoveNext
    Loop
    SUBRs.Close
    Set SUBRs = Nothing
    
    Fnc_MileageLastUserDelete = TempCnt
End Function
    
Public Function LOG_HP_SAVE(ByVal sDescription As String) As Boolean
    Dim FHandle As Integer
    Dim sText   As String
    
    If Dir(App.Path & "\Logs", vbDirectory) = "" Then MkDir "Logs"
    
    sText = Now & " : " & sDescription
    
    FHandle = FreeFile
    Open App.Path & "\Logs\" & Format(Date, "yyyyMMdd") & "_LOG_HP_CHANGE.Txt" For Append As FHandle

    Print #FHandle, sText
    Close #FHandle
    Exit Function

End Function

Public Function ERR_SAVE(ByVal sDescription As String) As Boolean
    Dim FHandle As Integer
    Dim sText   As String
    
'    ' 로그 파일을 생성하지 않는다.
'    If Dir(App.Path & "\NOLOG.TXT", vbDirectory) <> "" Then Exit Function
'
'    If Dir(App.Path & "\Logs", vbDirectory) = "" Then MkDir "Logs"
'
'    sText = Now & " : " & sDescription
'
'    FHandle = FreeFile
'    Open App.Path & "\Logs\" & Format(Date, "yyyyMMdd") & "_ERR_MSG.Txt" For Append As FHandle
'
'    Print #FHandle, sText
'    Close #FHandle
    Exit Function

End Function

Public Function FTP_LOG_SAVE(ByVal sDescription As String) As Boolean
    Dim FHandle As Integer
    Dim sText   As String
    
    sText = Now & " : " & sDescription
    
    FHandle = FreeFile
    Open App.Path & "\" & "FTP_MSG.Txt" For Append As FHandle

    Print #FHandle, sText
    Close #FHandle
    Exit Function

End Function

Public Function CheckUpgrade() As Boolean
    Dim sData(3) As String
    Dim RecentVersion As String
    Dim RecentMemo As String
    Dim MyVersion   As String
    
    CheckUpgrade = False
    '웹상에서 최신버젼이 얼마인지 받아옵니다.
    
    sData(0) = GetSetting("Laundry_Zi", "UpDate", "Url", "")
    sData(1) = GetSetting("Laundry_Zi", "UpDate", "Name", "")
    sData(2) = GetSetting("Laundry_Zi", "UpDate", "Fold", "")
    sData(3) = GetSetting("Laundry_Zi", "UpDate", "FoldName", "")
    
    'url 변경하여 무조건 레지스트리에 저장함.20090115
    
    If sData(0) <> "http://www.clean-aid.co.kr:8090/laundry/" Then
        sData(0) = "http://www.clean-aid.co.kr:8090/laundry/"
        sData(1) = "laundry.exe"
        sData(2) = App.Path & "\"
        sData(3) = "laundryUP.exe"
        SaveSetting "Laundry_Zi", "UpDate", "Url", sData(0)
        SaveSetting "Laundry_Zi", "UpDate", "Name", sData(1)
        SaveSetting "Laundry_Zi", "UpDate", "Fold", sData(2)
        SaveSetting "Laundry_Zi", "UpDate", "FoldName", sData(3)
        
'        MsgBox "프로그램을 다시 시작하십시요     ", vbCritical, "확인"
'        End
    End If
    
    If Trim(sData(0)) = "" Then
        sData(0) = "http://www.clean-aid.co.kr:8090/laundry/"
        sData(1) = "laundry.exe"
        sData(2) = App.Path & "\"
        sData(3) = "laundryUP.exe"
        SaveSetting "Laundry_Zi", "UpDate", "Url", sData(0)
        SaveSetting "Laundry_Zi", "UpDate", "Name", sData(1)
        SaveSetting "Laundry_Zi", "UpDate", "Fold", sData(2)
        SaveSetting "Laundry_Zi", "UpDate", "FoldName", sData(3)
    End If
    
    
    If Right(sData(0), 1) <> "/" Then sData(0) = sData(0) & "/"
    If Right(sData(2), 1) <> "\" Then sData(2) = sData(2) & "\"
    '프로그램 이상 행업.....20090115
    RecentVersion = OpenURL(sData(0) & "Ver.txt", 1000)
    RecentMemo = OpenURL(sData(0) & "memo.txt", 1000)
    
    'UpGrade
    MyVersion = GetSetting("Laundry_Zi", "UpDate", "VerSion", "")
    '현재 버젼이 최신버젼이면 메세지 출력 및 Exit Sub
    If Val(MyVersion) < Val(RecentVersion) Then
        CheckUpgrade = True
    End If

End Function

Public Function CheckUpgradeDBSend() As Boolean
    Dim sData(3) As String
    Dim RecentVersion As String
    Dim RecentMemo As String
    Dim MyVersion   As String
    
    CheckUpgradeDBSend = False
    '웹상에서 최신버젼이 얼마인지 받아옵니다.
    
    sData(0) = GetSetting("Laundry_DB", "UpDate", "Url", "")
    sData(1) = GetSetting("Laundry_DB", "UpDate", "Name", "")
    sData(2) = GetSetting("Laundry_DB", "UpDate", "Fold", "")
    sData(3) = GetSetting("Laundry_DB", "UpDate", "FoldName", "")
    
    'url 변경하여 무조건 레지스트리에 저장함.20090115
    
    If sData(0) <> "http://www.clean-aid.co.kr:8090/laundry/" Then
        sData(0) = "http://www.clean-aid.co.kr:8090/laundry/"
        sData(1) = "DBSend.exe"
        sData(2) = App.Path & "\"
        sData(3) = "DBSend.exe"
        SaveSetting "Laundry_DB", "UpDate", "Url", sData(0)
        SaveSetting "Laundry_DB", "UpDate", "Name", sData(1)
        SaveSetting "Laundry_DB", "UpDate", "Fold", sData(2)
        SaveSetting "Laundry_DB", "UpDate", "FoldName", sData(3)
        
'        MsgBox "프로그램을 다시 시작하십시요     ", vbCritical, "확인"
'        End
    End If
    
    If Trim(sData(0)) = "" Then
        sData(0) = "http://www.clean-aid.co.kr:8090/laundry/"
        sData(1) = "DBSend.exe"
        sData(2) = App.Path & "\"
        sData(3) = "DBSend.exe"
        SaveSetting "Laundry_DB", "UpDate", "Url", sData(0)
        SaveSetting "Laundry_DB", "UpDate", "Name", sData(1)
        SaveSetting "Laundry_DB", "UpDate", "Fold", sData(2)
        SaveSetting "Laundry_DB", "UpDate", "FoldName", sData(3)
    End If
    
    
    If Right(sData(0), 1) <> "/" Then sData(0) = sData(0) & "/"
    If Right(sData(2), 1) <> "\" Then sData(2) = sData(2) & "\"
    
    RecentVersion = OpenURL(sData(0) & "DBSendVer.txt", 1000)
    RecentMemo = OpenURL(sData(0) & "DBSendmemo.txt", 1000)
    
    'UpGrade
    MyVersion = GetSetting("Laundry_DB", "UpDate", "VerSion", "")
    '현재 버젼이 최신버젼이면 메세지 출력 및 Exit Sub
    If Val(MyVersion) < Val(RecentVersion) Then
        CheckUpgradeDBSend = True
    End If

End Function

Private Sub Fn_보관가격표생성()
    On Error GoTo Err_Rtn

    Query = " INSERT INTO 보관가격(보관월, 아이템수, 보관개월수, 보관가격) VALUES('01', 15,   9,  69000) "
    ADOCon.Execute Query
    
    Query = " INSERT INTO 보관가격(보관월, 아이템수, 보관개월수, 보관가격) VALUES('01', 15,   10, 71000) "
    ADOCon.Execute Query
    
    Query = " INSERT INTO 보관가격(보관월, 아이템수, 보관개월수, 보관가격) VALUES('01', 15,   11, 74000) "
    ADOCon.Execute Query
    
    Query = " INSERT INTO 보관가격(보관월, 아이템수, 보관개월수, 보관가격) VALUES('01', 15,   12, 77000) "
    ADOCon.Execute Query

    Exit Sub
    
Err_Rtn:
    MsgBox Err.Description, vbCritical, "확인"
End Sub

Public Function IsJuminNum(ByVal strJuminNum As String) As Boolean
    Dim iSum As Integer
    Dim iRe As Integer
    
    On Error GoTo Wrong_Number
   
    strJuminNum = Replace(strJuminNum, "-", "")
   
    If Len(strJuminNum) <> 13 Then
        IsJuminNum = False
        Exit Function
    End If
   
    If CInt(Mid(strJuminNum, 3, 2)) < 0 Or CInt(Mid(strJuminNum, 3, 2)) > 12 Or _
      CInt(Mid(strJuminNum, 5, 2)) < 0 Or CInt(Mid(strJuminNum, 5, 2)) > 31 Or _
      CInt(Mid(strJuminNum, 7, 1)) < 0 Or CInt(Mid(strJuminNum, 7, 1)) > 4 Then
         IsJuminNum = False
         Exit Function
    End If
   
    iSum = CInt(Mid(strJuminNum, 1, 1)) * 2 + _
           CInt(Mid(strJuminNum, 2, 1)) * 3 + _
           CInt(Mid(strJuminNum, 3, 1)) * 4 + _
           CInt(Mid(strJuminNum, 4, 1)) * 5 + _
           CInt(Mid(strJuminNum, 5, 1)) * 6 + _
           CInt(Mid(strJuminNum, 6, 1)) * 7 + _
           CInt(Mid(strJuminNum, 7, 1)) * 8 + _
           CInt(Mid(strJuminNum, 8, 1)) * 9 + _
           CInt(Mid(strJuminNum, 9, 1)) * 2 + _
           CInt(Mid(strJuminNum, 10, 1)) * 3 + _
           CInt(Mid(strJuminNum, 11, 1)) * 4 + _
           CInt(Mid(strJuminNum, 12, 1)) * 5

    iSum = iSum Mod 11
    iRe = 11 - iSum
   
    If iRe > 9 Then
        iRe = iRe Mod 10
    End If

    iSum = CInt(Mid(strJuminNum, 13, 1))

    If iSum = iRe Then
        IsJuminNum = True
    Else
        IsJuminNum = False
    End If
   
    Exit Function
   
Wrong_Number:
   IsJuminNum = False
End Function

'--------------------------------------------------------------------------------------------------------------
' Procedure : CheckSPointCard
' DateTime  : 2007-02-28 10:40
' Author    : pds2004
' Purpose   : S.Point 카드의 할인 대리점 인지의 여부를 확인한다.
'       일산지사(1004)      043:은평, 005:신월
'       천안유니트(1007)    021:평택, 022:서수원
'       수지유니트(1006)    045:구성
'       춘천지사(1002)      044:원주
'       인천지사(1003)      322:동천
'       안산지사(1005)      011:고잔
'       경산지사(1001)      355:해운대, 245:연재, 015:만촌, 223:월배, 038:칠성, 042:비산, 205:구미, 234:학성, 141:경산
'--------------------------------------------------------------------------------------------------------------
Public Function CheckSPointCard() As Boolean
    Dim sCompanyCode    As String
    Dim sStoreCode      As String
    
    CheckSPointCard = False
    
    ' 지사 코드/ 대리점 코드 설정
    sCompanyCode = 대리점정보.MasterCode
    sStoreCode = 대리점정보.대리점번호
    
    Select Case sCompanyCode
        ' 일산 지사
        Case "1004"
            ' 은평
            If sStoreCode = "043" Then
                CheckSPointCard = True: Exit Function
            ' 신월
            ElseIf sStoreCode = "005" Then
                CheckSPointCard = True: Exit Function
            End If
            
        ' 천안유니트
        Case "1007"
            ' 평택
            If sStoreCode = "021" Then
                CheckSPointCard = True: Exit Function
            ' 서수원
            ElseIf sStoreCode = "022" Then
                CheckSPointCard = True: Exit Function
            End If
        
        ' 수지유니트
        Case "1006"
            If sStoreCode = "045" Then
                CheckSPointCard = True: Exit Function
            End If
    
        ' 춘천 유니트
        Case "1002"
            If sStoreCode = "044" Then
                CheckSPointCard = True: Exit Function
            End If
    
        ' 인천지사
        Case "1003"
            If sStoreCode = "322" Then
                CheckSPointCard = True: Exit Function
            End If
    
        ' 안산지사
        Case "1005"
            If sStoreCode = "011" Then
                CheckSPointCard = True: Exit Function
            End If
    
        ' 경산지사
        Case "1001"
            ' 해운대
            If sStoreCode = "355" Then
                CheckSPointCard = True: Exit Function
            ' 연재
            ElseIf sStoreCode = "245" Then
                CheckSPointCard = True: Exit Function
            ' 만촌
            ElseIf sStoreCode = "015" Then
                CheckSPointCard = True: Exit Function
            ' 월배
            ElseIf sStoreCode = "223" Then
                CheckSPointCard = True: Exit Function
            ' 칠성
            ElseIf sStoreCode = "038" Then
                CheckSPointCard = True: Exit Function
            ' 비산
            ElseIf sStoreCode = "042" Then
                CheckSPointCard = True: Exit Function
            ' 구미
            ElseIf sStoreCode = "205" Then
                CheckSPointCard = True: Exit Function
            ' 학성
            ElseIf sStoreCode = "234" Then
                CheckSPointCard = True: Exit Function
            ' 경산
            ElseIf sStoreCode = "141" Then
                CheckSPointCard = True: Exit Function
            End If

        Case Else
                CheckSPointCard = False: Exit Function
    End Select
    
End Function



'--------------------------------------------------------------------------------------------------------------
' Procedure : Check_지정할인_20070406
' DateTime  : 2007-04-06 10:40
' Author    : pds2004
' Purpose   : 행사기간의 할인 대리점 인지의 여부를 확인한다.
'       경산지사(1001)      355:해운대
'--------------------------------------------------------------------------------------------------------------
Public Function Check_지정할인_20070406() As Boolean
    Dim sCompanyCode    As String
    Dim sStoreCode      As String
    
    Check_지정할인_20070406 = False
    
    ' 지사 코드/ 대리점 코드 설정
    sCompanyCode = 대리점정보.MasterCode
    sStoreCode = 대리점정보.대리점번호
    
    Select Case sCompanyCode
        ' 경산지사
        Case "1001"
            ' 해운대
            If sStoreCode = "355" Then
                Check_지정할인_20070406 = True: Exit Function
            End If
        
        Case Else
            Check_지정할인_20070406 = False: Exit Function
    End Select
End Function

Public Function f_dryPrice(txt As String) As Long
' 전달된 코드의 가격을 DB에서 읽어온다.
    Dim strDateChk As String
    Dim dblPrice   As Double
  
    On Error GoTo Err_Dry
      
    strDateChk = Format(Date, "yyyymmdd")
    
    Query = "SELECT 품명,가격 "
    Query = Query & "FROM 할인정보 "
    Query = Query & "WHERE 시작일 <= '" & strDateChk & "' "
    Query = Query & "AND   종료일 >= '" & strDateChk & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If Rs.RecordCount < 1 Then
        If chkDaySale = True Then
            Query = "SELECT 가격 "
            Query = Query & "FROM 목요세일 "
            Query = Query & "WHERE 구분코드 = '" & Trim(txt) & "'"
            Query = Query & "ORDER BY 가격 DESC "
        Else
            Query = "SELECT 가격 "
            Query = Query & "FROM 참조코드 "
            Query = Query & "WHERE 구분코드 = '" & Trim(txt) & "'"
            Query = Query & "ORDER BY 가격 DESC "
        End If
    Else
        Query = "SELECT 가격 "
        Query = Query & "FROM 할인정보 "
        Query = Query & "WHERE 구분코드 = '" & Trim(txt) & "' "
        Query = Query & "AND 시작일 <= '" & strDateChk & "' "
        Query = Query & "AND   종료일 >= '" & strDateChk & "'"
        Query = Query & "ORDER BY 가격 DESC "
    End If
    Rs.Close
    Set Rs = Nothing
    
    '
    Set Rs = New ADODB.Recordset
    Rs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
          
    If Rs.EOF = True Then
        Rs.Clone
        Set Rs = Nothing
        
        '-------------------------------------------------------
        '
        '-------------------------------------------------------
        Query = "SELECT 가격 "
        Query = Query & "FROM 참조코드 "
        Query = Query & "WHERE 구분코드 = '" & Trim(txt) & "'"
        Query = Query & "ORDER BY 가격 DESC "
        Set Rs = New ADODB.Recordset
        Rs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
        
        If Rs.EOF = True Then
            f_dryPrice = 0
            
            Rs.Close
            Set Rs = Nothing
            
            Exit Function
        End If
    End If
    
    dblPrice = Rs!가격
    
    Rs.Close
    Set Rs = Nothing
    
    
    ' 고객 정보를 얻어 온다.
    Call Fb고객정보(frm접수.txtCode)
    
    '크렌즈 겔러리 최초 손님일 경우 10% 자동 DC 적용처리
    If 대리점정보.MasterCode = M_COUPON_KLENZ_CODE Then
        ' 당일 등록일 경우, 반복적으로 계산 버튼을 클릭한경우의 처리를 위하여(한번만 처리되도록 하기 위하여)
        If 고객정보.등록일자 = Format(Date, "yyyyMMdd") Then
            ' 10원단위를 절사 한다.
            'dblPrice = CDbl(Int(CDbl((CCur(Rs!가격) * 0.9) / 100)) * 100)
            
            dblPrice = CDbl(Int(CDbl((dblPrice * 0.9) / 100)) * 100)
        End If
    End If
    
    f_dryPrice = dblPrice
    
    Exit Function
          
Err_Dry:
    f_dryPrice = 0
    Exit Function
End Function


'====================================================================================================
' Procedure : CheckMobileNumber
' DateTime  : 07-01-18 01:50
' Author    : BlueNice
' Last Edit :
' Return Val:
'----------------------------------------------------------------------------------------------------
' 설     명 : 전달된 번호가 휴대폰번호 인지를 확인한다.
'====================================================================================================
Public Function CheckMobileNumber(ByVal sNumber As String, ByRef sTel() As String) As Boolean
    Dim sTemp   As String
    Dim sLen    As Integer
    On Error GoTo CheckMobileNumber_Error

    CheckMobileNumber = False

    sTemp = Trim(sNumber)
    sTemp = Replace(sTemp, "-", "")
    sTemp = Replace(sTemp, ")", "")
    sTemp = Replace(sTemp, "/", "")
    If Left(sTemp, 2) <> "01" Or Len(sTemp) <= 9 Then Exit Function
    
    sLen = Len(sTemp)
    
    ' 0164401234
    If sLen = 10 Then
        sTel(0) = Left(sTemp, 3):   sTel(1) = Mid(sTemp, 4, 3): sTel(2) = Mid(sTemp, 7, 4)
        CheckMobileNumber = True
        
    '01190044523
    ElseIf sLen = 11 Then
        sTel(0) = Left(sTemp, 3):   sTel(1) = Mid(sTemp, 4, 4): sTel(2) = Mid(sTemp, 8, 4)
        CheckMobileNumber = True
        
    Else
        CheckMobileNumber = False
    End If

    On Error GoTo 0
    Exit Function

CheckMobileNumber_Error:
    CheckMobileNumber = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CheckMobileNumber of Form frmSMS"

    
End Function
'====================================================================================================
' Procedure : CheckTelNumber
' DateTime  : 07-01-18 01:50
' Author    : BlueNice
' Last Edit :
' Return Val:
'----------------------------------------------------------------------------------------------------
' 설     명 : 전달된 번호가 전화번호 인지를 확인한다.
'====================================================================================================
Public Function CheckTelNumber(ByVal sNumber As String, ByRef sTel() As String) As Boolean
    Dim sTemp   As String
    Dim sLen    As Integer
    On Error GoTo CheckTelNumber_Error

    CheckTelNumber = False

    sTemp = Trim(sNumber)
    sTemp = Replace(sTemp, "-", "")
    sTemp = Replace(sTemp, ")", "")
    sTemp = Replace(sTemp, "/", "")
    If Len(sTemp) <= 6 Then Exit Function
    
    sLen = Len(sTemp)
    
    ' 216 1234
    If sLen = 7 Then
        sTel(0) = "":   sTel(1) = Left(sTemp, 3): sTel(2) = Right(sTemp, 4)
        CheckTelNumber = True
        
    '2345 1234
    ElseIf sLen = 8 Then
        sTel(0) = "":   sTel(1) = Left(sTemp, 4): sTel(2) = Right(sTemp, 4)
        CheckTelNumber = True
        
    '2345 1234
    ElseIf sLen >= 9 And sLen <= 12 Then
        sTel(2) = Right(sTemp, 4)
        sTel(1) = Mid(Right(sTemp, 8), 1, 4)
        sTel(0) = Replace(sTemp, sTel(1) & sTel(2), "")
        CheckTelNumber = True
        
    Else
        CheckTelNumber = False
    End If

    On Error GoTo 0
    Exit Function

CheckTelNumber_Error:
    CheckTelNumber = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CheckTelNumber of Form frmSMS"

    
End Function


Public Function Check_일반할인_20070711() As Boolean
    Dim sMstCode    As String
    Dim sDay        As String
    Dim dblPrice    As Double
    Dim FHandle     As Integer
    

    On Error GoTo Check_일반할인_Error
    Check_일반할인_20070711 = False
    
    ' 지사 코드/ 대리점 코드 설정
    sMstCode = 대리점정보.MasterCode
    
    If sMstCode = "1001" And 대리점정보.대리점번호 = "234" Then
        sDay = Format(Date, "YYYY-MM-DD")
        ' 해당 기간동안에 한번만 실행한다.
        If sDay >= "2007-07-06" And sDay <= "2007-07-11" Then
        
            ' 20061101.TXT 파일이 없을 경우만 실행한다.
            ' 처음 실행할 경우 해당 파일이 없어서 실해되고 2번째부터는 실행되지 않느다.
            ' 해당 기간이 지나면 자동으로 프로그램에서 사용하지 않는다.
            If Dir(App.Path & "\20070711.TXT", vbDirectory) = "" Then
                ' 다음 이중 실행되지 않도록 파일을 생성한다.
                FHandle = FreeFile
                Open App.Path & "\20070711.TXT" For Append As FHandle
                Print #FHandle, Now
                
                Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드"
                Set SUBRs = New ADODB.Recordset
                SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
                
                ' 이전 자료를 모두 지운다.
                Query = "DELETE FROM 할인정보 WHERE 시작일 = '20070706' AND 종료일 = '20070711' "
                ADOCon.Execute Query
                
                Do While Not SUBRs.EOF
                    If IsNumeric(SUBRs.Fields("가격")) = True Then
                        Print #FHandle, "["; SUBRs.Fields("구분코드"); ":"; SUBRs.Fields("품명"); Tab; Tab; ":"; SUBRs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(SUBRs.Fields("가격"))) * 0.9)); "]"; Tab; "["; CStr((CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.9) * 0.01) * 100)); "]"
                    
                        dblPrice = 0
                        dblPrice = (CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.9) * 0.01) * 100)
                        
                        Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                        Query = Query & " VALUES ('20070706', '20070711', '" & SUBRs.Fields("구분코드") & "', '"
                        Query = Query & SUBRs.Fields("품명") & "', '" & CStr(dblPrice) & "', '1') "
                        ADOCon.Execute Query
                    Else
                        Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                        Query = Query & "[" & SUBRs.Fields("구분코드") & ":" & SUBRs.Fields("품명") & ":" & SUBRs.Fields("가격") & "]"
                        MsgBox Query, vbCritical, "경고"
                    End If
                    
                    SUBRs.MoveNext
                Loop
                
                SUBRs.Close
                Set SUBRs = Nothing
                
                Close #FHandle
            End If
        End If
    End If
    
    Check_일반할인_20070711 = True

    On Error GoTo 0
    
    Exit Function

Check_일반할인_Error:
    Resume
    Check_일반할인_20070711 = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Check_일반할인_20070711 of Module Global"
End Function

Public Function Check_일반할인_20071031() As Boolean
    Dim sMstCode    As String
    Dim sDay        As String
    Dim dblPrice    As Double
    Dim FHandle     As Integer

    On Error GoTo Check_일반할인_Error
    
    Check_일반할인_20071031 = False
    
    ' 지사 코드/ 대리점 코드 설정
    sMstCode = 대리점정보.MasterCode
    
    If sMstCode = "1008" And 대리점정보.대리점번호 = "028" Then
        sDay = Format(Date, "YYYY-MM-DD")
        ' 해당 기간동안에 수요일,목요일만 실행한다.
        If sDay >= "2007-07-25" And sDay <= "2007-10-31" Then
            ' 수요일 목요일만 할인한다.
            If Weekday(sDay) = 4 Or Weekday(sDay) = 5 Then
                sDay = Format(sDay, "yyyyMMdd")
                ' 해당 일자에 처음 실행할 경우 해당 파일이 없어서 실해되고 2번째부터는 실행되지 않느다.
                ' 해당 기간이 지나면 자동으로 프로그램에서 사용하지 않는다.
                If Dir(App.Path & "\" & sDay & ".TXT", vbDirectory) = "" Then
                    ' 다음 이중 실행되지 않도록 파일을 생성한다.
                    FHandle = FreeFile
                    Open App.Path & "\" & sDay & ".TXT" For Append As FHandle
                    Print #FHandle, Now
                    
                    Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드"
                    Set SUBRs = New ADODB.Recordset
                    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
                    
                    ' 이전 자료를 모두 지운다.
                    Query = "DELETE FROM 할인정보 WHERE 시작일 = '" & sDay & "' AND 종료일 = '" & sDay & "' "
                    ADOCon.Execute Query
                    
                    Do While Not SUBRs.EOF
                        If IsNumeric(SUBRs.Fields("가격")) = True Then
                            Print #FHandle, "["; SUBRs.Fields("구분코드"); ":"; SUBRs.Fields("품명"); Tab; Tab; ":"; SUBRs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(SUBRs.Fields("가격"))) * 0.8)); "]"; Tab; "["; CStr((CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.8) * 0.01) * 100)); "]"
                        
                            dblPrice = 0
                            dblPrice = (CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.8) * 0.01) * 100)
                            
                            Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                            Query = Query & " VALUES ('" & sDay & "', '" & sDay & "', '" & SUBRs.Fields("구분코드") & "', '"
                            Query = Query & SUBRs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                            ADOCon.Execute Query
                        Else
                            Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                            Query = Query & "[" & SUBRs.Fields("구분코드") & ":" & SUBRs.Fields("품명") & ":" & SUBRs.Fields("가격") & "]"
                            MsgBox Query, vbCritical, "경고"
                        End If
                        
                        SUBRs.MoveNext
                    Loop
                    
                    SUBRs.Close
                    Set SUBRs = Nothing
                    
                    Close #FHandle
                End If
            End If
        End If
    End If
    
    Check_일반할인_20071031 = True

    On Error GoTo 0
    
    Exit Function

Check_일반할인_Error:
    Resume
    
    Check_일반할인_20071031 = False
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Check_일반할인_20071031 of Module Global"
End Function

Public Function Check_일반할인_20070915() As Boolean
    Dim sMstCode    As String
    Dim sDay        As String
    Dim dblPrice    As Double
    Dim FHandle     As Integer

    On Error GoTo Check_일반할인_Error
    
    Check_일반할인_20070915 = False
    
    ' 지사 코드/ 대리점 코드 설정
    sMstCode = 대리점정보.MasterCode
    
    If sMstCode = "1012" And 대리점정보.대리점번호 = "023" Then
        sDay = Format(Date, "YYYY-MM-DD")
        
        ' 해당 기간동안에 수요일,목요일만 실행한다.
        If sDay >= "2007-09-03" And sDay <= "2007-09-15" Then
                
            ' 해당 일자에 처음 실행할 경우 해당 파일이 없어서 실해되고 2번째부터는 실행되지 않느다.
            ' 해당 기간이 지나면 자동으로 프로그램에서 사용하지 않는다.
            If Dir(App.Path & "\20070915.TXT", vbDirectory) = "" Then
                ' 다음 이중 실행되지 않도록 파일을 생성한다.
                FHandle = FreeFile
                Open App.Path & "\" & "\20070915.TXT" For Append As FHandle
                Print #FHandle, Now
                
                Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 WHERE left(구분코드,1) = 'a' "
                Set SUBRs = New ADODB.Recordset
                SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
                
                ' 이전 자료를 모두 지운다.
                Query = "DELETE FROM 할인정보 WHERE 시작일 = '20070903' AND 종료일 = '20070915' "
                ADOCon.Execute Query
                
                Do While Not SUBRs.EOF
                    If IsNumeric(SUBRs.Fields("가격")) = True Then
                        Print #FHandle, "["; SUBRs.Fields("구분코드"); ":"; SUBRs.Fields("품명"); Tab; Tab; ":"; SUBRs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(SUBRs.Fields("가격"))) * 0.8)); "]"; Tab; "["; CStr((CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.8) * 0.01) * 100)); "]"
                    
                        dblPrice = 0
                        dblPrice = (CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.8) * 0.01) * 100)
                        
                        Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                        Query = Query & " VALUES ('20070903', '20070915', '" & SUBRs.Fields("구분코드") & "', '"
                        Query = Query & SUBRs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                        ADOCon.Execute Query
                    Else
                        Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                        Query = Query & "[" & SUBRs.Fields("구분코드") & ":" & SUBRs.Fields("품명") & ":" & SUBRs.Fields("가격") & "]"
                        MsgBox Query, vbCritical, "경고"
                    End If
                    
                    SUBRs.MoveNext
                Loop
                
                SUBRs.Close
                Set SUBRs = Nothing
                
                Close #FHandle
            End If
        End If
    End If
    
    If sMstCode = "1012" And 대리점정보.대리점번호 = "300" Then
        sDay = Format(Date, "YYYY-MM-DD")
        ' 해당 기간동안에 수요일,목요일만 실행한다.
        If sDay >= "2007-09-03" And sDay <= "2007-09-15" Then
                
            ' 해당 일자에 처음 실행할 경우 해당 파일이 없어서 실해되고 2번째부터는 실행되지 않느다.
            ' 해당 기간이 지나면 자동으로 프로그램에서 사용하지 않는다.
            If Dir(App.Path & "\20070915.TXT", vbDirectory) = "" Then
                ' 다음 이중 실행되지 않도록 파일을 생성한다.
                FHandle = FreeFile
                Open App.Path & "\" & "\20070915.TXT" For Append As FHandle
                Print #FHandle, Now
                
                Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 WHERE ( left(구분코드,1) = 'a' OR  left(구분코드,1) = 'K')  "
                Set SUBRs = New ADODB.Recordset
                SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
                
                ' 이전 자료를 모두 지운다.
                Query = "DELETE FROM 할인정보 WHERE 시작일 = '20070903' AND 종료일 = '20070915' "
                ADOCon.Execute Query
                
                Do While Not SUBRs.EOF
                    If IsNumeric(SUBRs.Fields("가격")) = True Then
                        Print #FHandle, "["; SUBRs.Fields("구분코드"); ":"; SUBRs.Fields("품명"); Tab; Tab; ":"; SUBRs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(SUBRs.Fields("가격"))) * 0.8)); "]"; Tab; "["; CStr((CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.8) * 0.01) * 100)); "]"
                    
                        If Left(SUBRs.Fields("구분코드") & "", 1) = "a" Then
                            dblPrice = 0
                            dblPrice = (CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.8) * 0.01) * 100)
                        ElseIf Left(SUBRs.Fields("구분코드") & "", 1) = "k" Then
                            dblPrice = 0
                            dblPrice = (CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.7) * 0.01) * 100)
                        End If
                        
                        Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                        Query = Query & " VALUES ('20070903', '20070915', '" & SUBRs.Fields("구분코드") & "', '"
                        Query = Query & SUBRs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                        ADOCon.Execute Query
                    Else
                        Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                        Query = Query & "[" & SUBRs.Fields("구분코드") & ":" & SUBRs.Fields("품명") & ":" & SUBRs.Fields("가격") & "]"
                        MsgBox Query, vbCritical, "경고"
                    End If
                    SUBRs.MoveNext
                Loop
                
                SUBRs.Close
                Close #FHandle
            End If
        End If
    End If
    
    Check_일반할인_20070915 = True

    On Error GoTo 0
    
    Exit Function

Check_일반할인_Error:
    Resume
    
    Check_일반할인_20070915 = False
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Check_일반할인_20070915 of Module Global"
End Function

Public Function Check_일반할인() As Boolean
    Dim sDay    As String
    Dim dblPrice    As Double
    Dim SUBRs          As Recordset
    Dim FHandle     As Integer
    
    On Error GoTo Check_일반할인_Error
    Check_일반할인 = False
    
    
    If Check_할인대상확인 = True Then
        sDay = Format(Date, "YYYY-MM-DD")
        ' 해당 기간동안에 한번만 실행한다.
        If sDay >= "2006-11-01" And sDay <= "2006-11-12" Then
        
            ' 20061101.TXT 파일이 없을 경우만 실행한다.
            ' 처음 실행할 경우 해당 파일이 없어서 실해되고 2번째부터는 실행되지 않느다.
            ' 해당 기간이 지나면 자동으로 프로그램에서 사용하지 않는다.
            If Dir(App.Path & "\20061101.TXT", vbDirectory) = "" Then
                ' 다음 이중 실행되지 않도록 파일을 생성한다.
                FHandle = FreeFile
                Open App.Path & "\20061101.TXT" For Append As FHandle
                Print #FHandle, Now
                
                Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드"
                Set SUBRs = New ADODB.Recordset
                SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
                
                ' 이전 자료를 모두 지운다.
                Query = "DELETE FROM 할인정보 WHERE 시작일 = '20061101' AND 종료일 = '20061112' "
                ADOCon.Execute Query
                
                Do While Not SUBRs.EOF
                    If IsNumeric(SUBRs.Fields("가격")) = True Then
                        Print #FHandle, "["; SUBRs.Fields("구분코드"); ":"; SUBRs.Fields("품명"); Tab; Tab; ":"; SUBRs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(SUBRs.Fields("가격"))) * 0.9)); "]"; Tab; "["; CStr((CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.9) * 0.01) * 100)); "]"
                    
                        dblPrice = 0
                        dblPrice = (CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.9) * 0.01) * 100)
                        
                        Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                        Query = Query & " VALUES ('20061101', '20061112', '" & SUBRs.Fields("구분코드") & "', '"
                        Query = Query & SUBRs.Fields("품명") & "', '" & CStr(dblPrice) & "', '1') "
                        ADOCon.Execute Query
                    Else
                        Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                        Query = Query & "[" & SUBRs.Fields("구분코드") & ":" & SUBRs.Fields("품명") & ":" & SUBRs.Fields("가격") & "]"
                        MsgBox Query, vbCritical, "경고"
                    End If
                    SUBRs.MoveNext
                Loop
                
                SUBRs.Close
                Set SUBRs = Nothing
                
                Close #FHandle
            End If
        End If
    End If
    
    Check_일반할인 = True

    On Error GoTo 0
    
    Exit Function

Check_일반할인_Error:
    Resume
    Check_일반할인 = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Check_일반할인 of Module Global"
End Function

Public Function Check_명품세탁할인() As Boolean
    Dim sDay    As String

    On Error GoTo Check_명품세탁할인_Error
    
    Check_명품세탁할인 = False
    
    
    If Check_할인대상확인 = True Then
        sDay = Format(Date, "YYYY-MM-DD")
        
        If sDay >= "2006-11-01" And sDay <= "2006-11-12" Then
            Query = "UPDATE 대리점정보 SET "
            Query = Query & " 고가세탁비율 = 210 "
            ADOCon.Execute Query
                
        Else
            Query = "UPDATE 대리점정보 SET "
            Query = Query & " 고가세탁비율 = 300 "
            ADOCon.Execute Query
        End If
    End If
    
    Check_명품세탁할인 = True

    On Error GoTo 0
    
    Exit Function

Check_명품세탁할인_Error:
    Check_명품세탁할인 = False
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Check_명품세탁할인 of Module Global"
End Function


Private Function Check_할인대상확인() As Boolean
    Dim sMstCode    As String
    Dim sStoreCode  As String
    
    Dim sCode_1001  As String
    Dim sCode_1002  As String
    Dim sCode_1003  As String
    Dim sCode_1004  As String
    Dim sCode_1005  As String
    Dim sCode_1006  As String
    Dim sCode_1007  As String
    Dim sCode_1008  As String
    Dim sCode_1009  As String

    On Error GoTo Check_할인대상확인_Error
    
    Check_할인대상확인 = False
    
    ' 지사 코드/ 대리점 코드 설정
    sMstCode = 대리점정보.MasterCode
    sStoreCode = 대리점정보.대리점번호
    
    
    '------------- 대리점 설정 ----------------
    ' 경산 지사
    sCode_1001 = "278,038,015,223,009,055,234"
    ' 춘천 지사
    sCode_1002 = "044"
    ' 인천 지사
    sCode_1003 = "759"
    ' 일산 지사
    sCode_1004 = "007,005"
    ' 안산 유니트
    sCode_1005 = "011,055"
    ' 수지 유니트
    sCode_1006 = "045"
    ' 천안 유니트
    sCode_1007 = "021,022"
    ' 중산 유니트
    sCode_1008 = ""
    ' 자인 유니트
    sCode_1009 = "141,205,042"
    
    ' 경산 지사
    If sMstCode = "1001" And InStr(sCode_1001, sStoreCode) > 0 Then
        Check_할인대상확인 = True
        Exit Function
        
    ' 춘전 지사
    ElseIf sMstCode = "1002" And InStr(sCode_1002, sStoreCode) > 0 Then
        Check_할인대상확인 = True
        Exit Function
    
    ' 인천 지사
    ElseIf sMstCode = "1003" And InStr(sCode_1003, sStoreCode) > 0 Then
        Check_할인대상확인 = True
        Exit Function
    
    ' 일산 지사
    ElseIf sMstCode = "1004" And InStr(sCode_1004, sStoreCode) > 0 Then
        Check_할인대상확인 = True
        Exit Function
    
    ' 안산 유니트
    ElseIf sMstCode = "1005" And InStr(sCode_1005, sStoreCode) > 0 Then
        Check_할인대상확인 = True
        Exit Function
    
    ' 수지 유니트
    ElseIf sMstCode = "1006" And InStr(sCode_1006, sStoreCode) > 0 Then
        Check_할인대상확인 = True
        Exit Function
    
    ' 천안 유니트
    ElseIf sMstCode = "1007" And InStr(sCode_1007, sStoreCode) > 0 Then
        Check_할인대상확인 = True
        Exit Function
    
    ' 중산 유니트
    ElseIf sMstCode = "1008" And InStr(sCode_1008, sStoreCode) > 0 Then
        Check_할인대상확인 = True
        Exit Function
    
    ' 자인 유니트
    ElseIf sMstCode = "1009" And InStr(sCode_1009, sStoreCode) > 0 Then
        Check_할인대상확인 = True
        Exit Function
    
    End If
    
    Check_할인대상확인 = False
    On Error GoTo 0
    Exit Function

Check_할인대상확인_Error:
    Check_할인대상확인 = False

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Check_할인대상확인 of Module Global"
End Function

Public Function Check_일반할인_20071025() As Boolean
    Dim sMstCode    As String
    Dim sDay        As String
    Dim dblPrice    As Double
    Dim FHandle     As Integer
    

    On Error GoTo Check_일반할인_Error
    Check_일반할인_20071025 = False
    
    ' 지사 코드/ 대리점 코드 설정
    sMstCode = 대리점정보.MasterCode
    
    If sMstCode = "1007" Then
        Select Case 대리점정보.대리점번호
            Case "004", "015", "011", "034"
                If Dir(App.Path & "\200710255.TXT", vbDirectory) <> "" Then
                    Kill App.Path & "\200710255.TXT"
                End If
            
                    
                sDay = Format(Date, "YYYY-MM-DD")
                If sDay >= "2007-10-10" And sDay <= "2007-10-25" Then
                        
                    ' 해당 일자에 처음 실행할 경우 해당 파일이 없어서 실해되고 2번째부터는 실행되지 않느다.
                    ' 해당 기간이 지나면 자동으로 프로그램에서 사용하지 않는다.
                    If Dir(App.Path & "\200710255_01.TXT", vbDirectory) = "" Then
                        ' 다음 이중 실행되지 않도록 파일을 생성한다.
                        FHandle = FreeFile
                        Open App.Path & "\" & "\200710255_01.TXT" For Append As FHandle
                        Print #FHandle, Now
                        
                        Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                        Query = Query & " WHERE not ( left(구분코드,3) = 'm00'  or  left(구분코드,3) = 'm01' or  left(구분코드,1) = 'a' or  left(구분코드,1) = 'v'  )"
                        Set SUBRs = New ADODB.Recordset
                        SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
                        
                        ' 이전 자료를 모두 지운다.
                        Query = "DELETE FROM 할인정보 WHERE 시작일 = '20071010' AND 종료일 = '20071025' "
                        ADOCon.Execute Query
                        
                        Do While Not SUBRs.EOF
                            If IsNumeric(SUBRs.Fields("가격")) = True Then
                                Print #FHandle, "["; SUBRs.Fields("구분코드"); ":"; SUBRs.Fields("품명"); Tab; Tab; ":"; SUBRs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(SUBRs.Fields("가격"))) * 0.8)); "]"; Tab; "["; CStr((CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.8) * 0.01) * 100)); "]"
                            
                                dblPrice = 0
                                dblPrice = (CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.8) * 0.01) * 100)
                                
                                Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                                Query = Query & " VALUES ('20071010', '20071025', '" & SUBRs.Fields("구분코드") & "', '"
                                Query = Query & SUBRs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                                ADOCon.Execute Query
                            Else
                                Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                                Query = Query & "[" & SUBRs.Fields("구분코드") & ":" & SUBRs.Fields("품명") & ":" & SUBRs.Fields("가격") & "]"
                                MsgBox Query, vbCritical, "경고"
                            End If
                            SUBRs.MoveNext
                        Loop
                        
                        SUBRs.Close
                        
                        Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                        Query = Query & " WHERE   ( left(구분코드,3) = 'm00'  or  left(구분코드,3) = 'm01' or  left(구분코드,1) = 'a' or  left(구분코드,1) = 'v'  )"
                        Set SUBRs = MyDB.OpenRecordset(Query)
                        
                        Do While Not SUBRs.EOF
                            If IsNumeric(SUBRs.Fields("가격")) = True Then
                                Print #FHandle, "["; SUBRs.Fields("구분코드"); ":"; SUBRs.Fields("품명"); Tab; Tab; ":"; SUBRs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(SUBRs.Fields("가격"))) * 0.8)); "]"; Tab; "["; CStr((CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.8) * 0.01) * 100)); "]"
                            
                                dblPrice = Val(CStr(SUBRs.Fields("가격")))
                                
                                Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                                Query = Query & " VALUES ('20071010', '20071025', '" & SUBRs.Fields("구분코드") & "', '"
                                Query = Query & SUBRs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                                ADOCon.Execute Query
                            Else
                                Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                                Query = Query & "[" & SUBRs.Fields("구분코드") & ":" & SUBRs.Fields("품명") & ":" & SUBRs.Fields("가격") & "]"
                                MsgBox Query, vbCritical, "경고"
                            End If
                            SUBRs.MoveNext
                        Loop
                        
                        SUBRs.Close
                        
                        Close #FHandle
                    End If
                End If
            Case Else
            
        End Select
    End If
    
    
    Check_일반할인_20071025 = True

    On Error GoTo 0
    Exit Function

Check_일반할인_Error:
    Resume
    Check_일반할인_20071025 = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Check_일반할인_20071025 of Module Global"
End Function

Public Function Check_일반할인_20071017() As Boolean
    Dim sMstCode    As String
    Dim sDay        As String
    Dim dblPrice    As Double
    Dim SUBRs          As Recordset
    Dim FHandle     As Integer
    

    On Error GoTo Check_일반할인_Error
    Check_일반할인_20071017 = False
    
    ' 지사 코드/ 대리점 코드 설정
    sMstCode = 대리점정보.MasterCode
    
    If sMstCode = "1001" Then
        Select Case 대리점정보.대리점번호
            Case "205"
                ' 205 구미점
                    
                sDay = Format(Date, "YYYY-MM-DD")
                If sDay >= "2007-10-11" And sDay <= "2007-10-17" Then
                        
                    ' 해당 일자에 처음 실행할 경우 해당 파일이 없어서 실해되고 2번째부터는 실행되지 않느다.
                    ' 해당 기간이 지나면 자동으로 프로그램에서 사용하지 않는다.
                    If Dir(App.Path & "\20071017.TXT", vbDirectory) = "" Then
                        ' 다음 이중 실행되지 않도록 파일을 생성한다.
                        FHandle = FreeFile
                        Open App.Path & "\" & "\20071017.TXT" For Append As FHandle
                        Print #FHandle, Now
                        
                        ' 전품목 20% 할인
                        Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                        Set SUBRs = New ADODB.Recordset
                        SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
                        
                        ' 이전 자료를 모두 지운다.
                        Query = "DELETE FROM 할인정보 WHERE 시작일 = '20071011' AND 종료일 = '20071017' "
                        ADOCon.Execute Query
                        
                        Do While Not SUBRs.EOF
                            If IsNumeric(SUBRs.Fields("가격")) = True Then
                                Print #FHandle, "["; SUBRs.Fields("구분코드"); ":"; SUBRs.Fields("품명"); Tab; Tab; ":"; SUBRs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(SUBRs.Fields("가격"))) * 0.8)); "]"; Tab; "["; CStr((CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.8) * 0.01) * 100)); "]"
                            
                                dblPrice = 0
                                dblPrice = (CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.8) * 0.01) * 100)
                                
                                Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                                Query = Query & " VALUES ('20071011', '20071017', '" & SUBRs.Fields("구분코드") & "', '"
                                Query = Query & SUBRs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                                ADOCon.Execute Query
                            Else
                                Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                                Query = Query & "[" & SUBRs.Fields("구분코드") & ":" & SUBRs.Fields("품명") & ":" & SUBRs.Fields("가격") & "]"
                                MsgBox Query, vbCritical, "경고"
                            End If
                            SUBRs.MoveNext
                        Loop
                        
                        SUBRs.Close
                        
                        Close #FHandle
                    End If
                End If
            Case Else
            
        End Select
    End If
    
    
    Check_일반할인_20071017 = True

    On Error GoTo 0
    Exit Function

Check_일반할인_Error:
    Resume
    Check_일반할인_20071017 = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Check_일반할인_20071017 of Module Global"
End Function


Public Function Check_일반할인_20071018() As Boolean
    Dim sMstCode    As String
    Dim sDay        As String
    Dim dblPrice    As Double
    Dim FHandle     As Integer

    On Error GoTo Check_일반할인_Error
    Check_일반할인_20071018 = False
    
    ' 지사 코드/ 대리점 코드 설정
    sMstCode = 대리점정보.MasterCode
    
    If sMstCode = "1001" Then
        Select Case 대리점정보.대리점번호
            Case "042", "234"
                ' 042 비산점, 234 학성점
                    
                sDay = Format(Date, "YYYY-MM-DD")
                If sDay >= "2007-10-11" And sDay <= "2007-10-18" Then
                        
                    ' 해당 일자에 처음 실행할 경우 해당 파일이 없어서 실해되고 2번째부터는 실행되지 않느다.
                    ' 해당 기간이 지나면 자동으로 프로그램에서 사용하지 않는다.
                    If Dir(App.Path & "\20071018.TXT", vbDirectory) = "" Then
                        ' 다음 이중 실행되지 않도록 파일을 생성한다.
                        FHandle = FreeFile
                        Open App.Path & "\" & "\20071018.TXT" For Append As FHandle
                        Print #FHandle, Now
                        
                        ' 전품목 20% 할인
                        Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                        Set SUBRs = New ADODB.Recordset
                        SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
                        
                        ' 이전 자료를 모두 지운다.
                        Query = "DELETE FROM 할인정보 WHERE 시작일 = '20071011' AND 종료일 = '20071018' "
                        ADOCon.Execute Query
                        
                        Do While Not SUBRs.EOF
                            If IsNumeric(SUBRs.Fields("가격")) = True Then
                                Print #FHandle, "["; SUBRs.Fields("구분코드"); ":"; SUBRs.Fields("품명"); Tab; Tab; ":"; SUBRs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(SUBRs.Fields("가격"))) * 0.8)); "]"; Tab; "["; CStr((CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.8) * 0.01) * 100)); "]"
                            
                                dblPrice = 0
                                dblPrice = (CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.8) * 0.01) * 100)
                                
                                Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                                Query = Query & " VALUES ('20071011', '20071018', '" & SUBRs.Fields("구분코드") & "', '"
                                Query = Query & SUBRs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                                ADOCon.Execute Query
                            Else
                                Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                                Query = Query & "[" & SUBRs.Fields("구분코드") & ":" & SUBRs.Fields("품명") & ":" & SUBRs.Fields("가격") & "]"
                                MsgBox Query, vbCritical, "경고"
                            End If
                            SUBRs.MoveNext
                        Loop
                        
                        SUBRs.Close
                        
                        Close #FHandle
                    End If
                End If
            Case Else
            
        End Select
    End If
    
    
    Check_일반할인_20071018 = True

    On Error GoTo 0
    Exit Function

Check_일반할인_Error:
    Resume
    Check_일반할인_20071018 = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Check_일반할인_20071018 of Module Global"
End Function



'--------------------------------------------------------------------------------------------------------------
' Procedure : Check_이마트할인_20071114
' DateTime  : 2007-10-31
' Author    : pds2004
' Purpose   : 이마트 할인 여부   할인기간 2007-11-01 ~ 2007-11-14일 까지
'       본사(1000)          067: 양주점
'       경산지사(1001)      015:만촌, 223:월배, 245:연재, 355:해운대,  234:학성, 141:경산,  038:칠성, 197:상주점 205:구미,  042:비산
'       춘천지사(1002)      044:원주
'       인천지사(1003)      141:동천
'       일산지사(1004)      043:은평, 010:신월, 007:파주점
'       안산지사(1005)      011:고잔, 055:시화점
'       천안지사(1007)      021:평택, 022:서수원
'       남부지사(1011)      029:수지점
'--------------------------------------------------------------------------------------------------------------
Public Function Check_이마트할인대상확인_20071114() As Boolean
    Dim sCompanyCode    As String
    Dim sStoreCode      As String
    
    Check_이마트할인대상확인_20071114 = False
    
    ' 지사 코드/ 대리점 코드 설정
    sCompanyCode = 대리점정보.MasterCode
    sStoreCode = 대리점정보.대리점번호
    
    Select Case sCompanyCode
'       본사(1000)          067: 양주점
        Case "1000"
            If sStoreCode = "067" Then
                Check_이마트할인대상확인_20071114 = True: Exit Function
            End If
        
'       경산지사(1001)      015:만촌, 223:월배, 245:연재, 355:해운대,  234:학성, 141:경산,  038:칠성, 197:상주점 205:구미,  042:비산
        Case "1001"
            Select Case sStoreCode
                Case "015", "223", "245", "355", "234", "141", "038", "197", "205", "042"
                Check_이마트할인대상확인_20071114 = True: Exit Function
            End Select
            
'       춘천지사(1002)      044:원주
        Case "1002"
            If sStoreCode = "044" Then
                Check_이마트할인대상확인_20071114 = True: Exit Function
            End If

'       인천지사(1003)      141:동천
        Case "1003"
            Select Case sStoreCode
                Case "141", "060"
                Check_이마트할인대상확인_20071114 = True: Exit Function
            End Select
            
'       일산지사(1004)      043:은평, 010:신월, 007:파주점
        Case "1004"
            Select Case sStoreCode
                Case "043", "010", "007"
                Check_이마트할인대상확인_20071114 = True: Exit Function
            End Select
            
'       안산지사(1005)      011:고잔, 055:시화점
        Case "1005"
            Select Case sStoreCode
                Case "011", "055"
                Check_이마트할인대상확인_20071114 = True: Exit Function
            End Select
        
'       천안지사(1007)      021:평택, 022:서수원
        Case "1007"
            Select Case sStoreCode
                Case "021", "022"
                Check_이마트할인대상확인_20071114 = True: Exit Function
            End Select

'       지사(1008)      028:
        Case "1008"
            Select Case sStoreCode
                Case "028"
                Check_이마트할인대상확인_20071114 = True: Exit Function
            End Select

'       남부지사(1011)      029:수지점
        Case "1011"
            If sStoreCode = "029" Then
                Check_이마트할인대상확인_20071114 = True: Exit Function
            End If
        
        Case Else
                Check_이마트할인대상확인_20071114 = False: Exit Function
    End Select
    
End Function



'--------------------------------------------------------------------------------------------------------------
' Procedure : Check_이마트할인_20081112
' DateTime  : 2008-11-12
' Author    : pds2004
' Purpose   : 이마트 할인 여부   할인기간 2008-10-30 ~ 2008-11-12일 까지
'1004 일산지사  002
'1004 일산지사  010
'1005 안산지사  171
'1007 천안지사  021
'1006 용인지사  015
'1003 인천지사  322
'1003 인천지사  045
'1004 일산지사  007
'1011 남부지사  014
'1019 마루산에이드  067
'1003 인천지사  166
'1001 경산지사  197
'1001 경산지사  042
'1015 부산지사  020
'1001 경산지사  077
'1001 경산지사  066
'1001 경산지사  060
'1001 경산지사  070
'1001 경산지사  205
'1015 부산지사  245
'1015 부산지사  055
'1002 춘천지사  044
'1019 마루산에이드  057
'1019 마루산에이드  068
'1016 동작지사      002
'1007 천안지사      075
'--------------------------------------------------------------------------------------------------------------
Public Function Check_이마트할인대상확인_20081112() As Boolean
    Dim sCompanyCode    As String
    Dim sStoreCode      As String
    
    Check_이마트할인대상확인_20081112 = False
    
    ' 지사 코드/ 대리점 코드 설정
    sCompanyCode = 대리점정보.MasterCode
    sStoreCode = 대리점정보.대리점번호
    
'1019 마루산에이드  067
'1019 마루산에이드  057
'1019 마루산에이드  068
    
    

    
    Select Case sCompanyCode
'       경산지사(1001)
        Case "1001"
            Select Case sStoreCode
                Case "197", "042", "077", "066", "060", "070", "205"
                Check_이마트할인대상확인_20081112 = True: Exit Function
            End Select
            
'       춘천지사(1002)
        Case "1002"
            If sStoreCode = "044" Then
                Check_이마트할인대상확인_20081112 = True: Exit Function
            End If

'       인천지사(1003)
        Case "1003"
            Select Case sStoreCode
                Case "045", "322", "166"
                Check_이마트할인대상확인_20081112 = True: Exit Function
            End Select
            
'       일산지사(1004)
        Case "1004"
            Select Case sStoreCode
                Case "002", "010", "007"
                Check_이마트할인대상확인_20081112 = True: Exit Function
            End Select
            
'       안산지사(1005)
        Case "1005"
            Select Case sStoreCode
                Case "171"
                Check_이마트할인대상확인_20081112 = True: Exit Function
            End Select

'       용인지사(1006)
        Case "1006"
            Select Case sStoreCode
                Case "015"
                Check_이마트할인대상확인_20081112 = True: Exit Function
            End Select
        
'       천안지사(1007)
        Case "1007"
            Select Case sStoreCode
                Case "021", "075"
                Check_이마트할인대상확인_20081112 = True: Exit Function
            End Select

'       남부지사(1011)
        Case "1011"
            Select Case sStoreCode
                Case "014"
                Check_이마트할인대상확인_20081112 = True: Exit Function
            End Select
            
'       부산지사(1015)
        Case "1015"
            Select Case sStoreCode
                Case "020", "245", "055"
                Check_이마트할인대상확인_20081112 = True: Exit Function
            End Select
            
'       동작지사(1016)
        Case "1016"
            Select Case sStoreCode
                Case "002"
                Check_이마트할인대상확인_20081112 = True: Exit Function
            End Select
            
'       마루산에이드(1019)
        Case "1019"
            Select Case sStoreCode
                Case "067", "057", "068"
                Check_이마트할인대상확인_20081112 = True: Exit Function
            End Select
            
        
        Case Else
                Check_이마트할인대상확인_20081112 = False: Exit Function
    End Select
    
End Function


'--------------------------------------------------------------------------------------------------------------
' Procedure : Check_일반매장할인대상확인_20081110
' DateTime  : 2008-11-10
' Author    : pds2004
' Purpose   : 할인 여부   할인기간 2008-11-01 ~ 2008-11-10일 까지

'---------- 제외 매장 (아래 코드가 행사 하지 않는다 ) -------
'1002( 춘천지사 )        007
'1018( 부천서부지사 )    028
'1018( 부천서부지사 )    030
'1018( 부천서부지사 )    023
'1018( 부천서부지사 )    001
'1018( 부천서부지사 )    010
'1018( 부천서부지사 )    003
'1018( 부천서부지사 )    007
'1018( 부천서부지사 )    024
'1018( 부천서부지사 )    006
'1018( 부천서부지사 )    020
'--------------------------------------------------------------------------------------------------------------
Public Function Check_일반매장할인대상확인_20081110() As Boolean
    Dim sCompanyCode    As String
    Dim sStoreCode      As String
    
    Check_일반매장할인대상확인_20081110 = True
    
    ' 지사 코드/ 대리점 코드 설정
    sCompanyCode = 대리점정보.MasterCode
    sStoreCode = 대리점정보.대리점번호
    
  
  
  ' 만일 이마트의 행사와 연관이 있을 경우 제외 매장으로 처리한다.
  If Check_이마트할인대상확인_20081112 = True Then
    Check_일반매장할인대상확인_20081110 = False
    Exit Function
  End If
  
  
    ' 리턴 값이 반대로 리턴한다.
    ' 해당 매장이 세일을 하지 않기 때문에 해당되면 False를 러턴한다.
    
    
    Select Case sCompanyCode

'       춘천지사(1002)
        Case "1002"
            If sStoreCode = "007" Then
                Check_일반매장할인대상확인_20081110 = False: Exit Function
            End If


'       부산지사(1015)
        Case "1015"
            Select Case sStoreCode
                Case "300", "141", "028"
                Check_일반매장할인대상확인_20081110 = False: Exit Function
            End Select
            
'       부천서부지사(1018)
        Case "1018"
            Select Case sStoreCode
                Case "028", "030", "023", "001", "010", "003", "007", "024", "006", "020"
                Check_일반매장할인대상확인_20081110 = False: Exit Function
            End Select
            
        
        Case Else
                Check_일반매장할인대상확인_20081110 = True: Exit Function
    End Select
    
End Function


'--------------------------------------------------------------------------------------------------------------
' Procedure : Check_일반매장할인대상확인_20090615
' DateTime  : 2009-06-04
' Author    : pds2004
' Purpose   : 할인 여부   할인기간 2009-06-08 ~ 2009-06-15일 까지

'---------- 제외 매장 (아래 코드가 행사 하지 않는다 ) -------
'1002( 춘천지사 ) 모든 매장
'1003( 인천지사 ) 모든 매장
'1017( 속초지사 )

'--------------------------------------------------------------------------------------------------------------
Public Function Check_일반매장할인대상확인_20090615() As Boolean
    Dim sCompanyCode    As String
    Dim sStoreCode      As String
    
    
    Check_일반매장할인대상확인_20090615 = False
    
    ' 지사 코드/ 대리점 코드 설정
    sCompanyCode = 대리점정보.MasterCode
    sStoreCode = 대리점정보.StoreCode
    
    ' 리턴 값이 반대로 리턴한다.
    ' 해당 매장이 세일을 하지 않기 때문에 해당되면 False를 러턴한다.
    
    Select Case sCompanyCode

        '춘천지사(1002), 인천지사(1003), 속초지사(1017) 제외
        Case "1002", "1003", "1017"
            Check_일반매장할인대상확인_20090615 = False: Exit Function


        '경산지사  (행사 매장)
        Case "1001"
            Select Case sStoreCode
                Case "100012", "100015", "100028", "100029", "100047", "100084", "100085", "100093", "100112", "100144", "100163", "100177", "100189", "100209", "100248"
                Check_일반매장할인대상확인_20090615 = True: Exit Function
            End Select

        '일산지사  (행사 매장) "100126 제외
        Case "1004"
            Select Case sStoreCode
                Case "100009", "100054", "100127", "100137", "100140", "100145", "100169", "100218", "100242", "100260", "100271"
                Check_일반매장할인대상확인_20090615 = True: Exit Function
            End Select

        '안산지사  (행사 매장)
        Case "1005"
            Select Case sStoreCode
                Case "100024", "100031", "100243"
                Check_일반매장할인대상확인_20090615 = True: Exit Function
            End Select

        '용인지사  (행사 매장)
        Case "1006"
            Select Case sStoreCode
                Case "100018", "100039", "100044", "100053", "100056", "100148", "100207", "100212"
                Check_일반매장할인대상확인_20090615 = True: Exit Function
            End Select
            
        '천안지사  (행사 매장)
        Case "1007"
            Select Case sStoreCode
                Case "100022", "100043", "100058", "100106", "100111", "100125", "100131", "100191", "100202", "100222", "100235", "100252", "100263", "100266"
                Check_일반매장할인대상확인_20090615 = True: Exit Function
            End Select

        '중산지사  (행사 매장)
        Case "1008"
            Select Case sStoreCode
                Case "100055", "100060", "100061", "100062", "100092", "100118", "100138", "100154", "100178", "100180", "100192", "100194", "100205", "100264"
                Check_일반매장할인대상확인_20090615 = True: Exit Function
            End Select

        '분당지사  (행사 매장)
        Case "1010"
            Select Case sStoreCode
                Case "100136", "100168", "100173"
                Check_일반매장할인대상확인_20090615 = True: Exit Function
            End Select

        '남부지사  (행사 매장)
        Case "1011"
            Select Case sStoreCode
                Case "100003", "100006", "100007", "100014", "100069", "100075", "100097", "100104", "100109", "100128", "100139", "100142", "100156", "100214"
                Check_일반매장할인대상확인_20090615 = True: Exit Function
            End Select

        '신갈지사  (행사 매장)
        Case "1013"
            Select Case sStoreCode
                Case "100040", "100071", "100113", "100124", "100183", "100236", "100238", "100244", "100245", "100249", "100257"
                Check_일반매장할인대상확인_20090615 = True: Exit Function
            End Select

        '부산지사  (행사 매장)
        Case "1015"
            Select Case sStoreCode
                Case "100027", "100032", "100171", "100185", "100227", "100228", "100247", "100255"
                Check_일반매장할인대상확인_20090615 = True: Exit Function
            End Select

        '동작지사  (행사 매장)
        Case "1016"
            Select Case sStoreCode
                Case "100080", "100098", "100133", "100172", "100174", "100175", "100182", "100211", "100216", "100256"
                Check_일반매장할인대상확인_20090615 = True: Exit Function
            End Select

        '서부지사  (행사 매장)
        Case "1018"
            Select Case sStoreCode
                Case "100094", "100099", "100100", "100107", "100108", "100129", "100186", "100219", "100224", "100250", "100251", "100258"
                Check_일반매장할인대상확인_20090615 = True: Exit Function
            End Select

        '1사업장  (행사 매장)
        Case "1019"
            Select Case sStoreCode
                Case "100042", "100051", "100052", "100068", "100074", "100077", "100181", "100203", "100208", "100223", "100230", "100240", "100261"
                Check_일반매장할인대상확인_20090615 = True: Exit Function
            End Select

        '문경지사  (행사 매장) "100030",
        Case "1020"
            Select Case sStoreCode
                Case "100120", "100122", "100176", "100220", "100232", "100233"
                Check_일반매장할인대상확인_20090615 = True: Exit Function
            End Select

        '2사업장  (행사 매장)
        Case "1021"
            Select Case sStoreCode
                Case "100025", "100091", "100114", "100115", "100119", "100130", "100143", "100195", "100197", "100200", "100234", "100270", "100274"
                Check_일반매장할인대상확인_20090615 = True: Exit Function
            End Select

        '울산지사  (행사 매장)
        Case "1022"
            Select Case sStoreCode
                Case "100021", "100226", "100246"
                Check_일반매장할인대상확인_20090615 = True: Exit Function
            End Select
        
        '도곡지사  (행사 매장)
        Case "1023"
            Select Case sStoreCode
                Case "100004", "100016", "100268"
                Check_일반매장할인대상확인_20090615 = True: Exit Function
            End Select
        
        
        Case Else
                Check_일반매장할인대상확인_20090615 = False: Exit Function
    End Select
    
End Function

Public Function Check_이마트할인_20071114() As Boolean
    Dim sDay        As String
    Dim dblPrice    As Double
    Dim FHandle     As Integer
    
    Dim sStartDate  As String
    Dim sEndDate    As String
    
    

    On Error GoTo Check_일반할인_Error
    Check_이마트할인_20071114 = False
    
    
    sStartDate = "20071101"
    sEndDate = "20071114"
    
    If Check_이마트할인대상확인_20071114 = True Then
                
        sDay = Format(Date, "YYYY-MM-DD")
        If sDay >= Format(sStartDate, "@@@@-@@-@@") And sDay <= Format(sEndDate, "@@@@-@@-@@") Then
                
            ' 해당 일자에 처음 실행할 경우 해당 파일이 없어서 실해되고 2번째부터는 실행되지 않느다.
            ' 해당 기간이 지나면 자동으로 프로그램에서 사용하지 않는다.
            If Dir(App.Path & "\" & sEndDate & ".TXT", vbDirectory) = "" Then
                ' 다음 이중 실행되지 않도록 파일을 생성한다.
                FHandle = FreeFile
                Open App.Path & "\" & sEndDate & ".TXT" For Append As FHandle
                Print #FHandle, Now
                
'                f코드(상의류) > 30% 할인
'                g코드(하의류) > 30% 할인
'                r코드(스커트류) > 30% 할인
                Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                Query = Query & " WHERE (  left(구분코드,1) = 'f' or left(구분코드,1) = 'g' or  left(구분코드,1) = 'r'  )"
                Set SUBRs = New ADODB.Recordset
                SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
                
                ' 이전 자료를 모두 지운다.
                Query = "DELETE FROM 할인정보 WHERE 시작일 = '" & sStartDate & "' AND 종료일 = '" & sEndDate & "' "
                ADOCon.Execute Query
                
                Do While Not SUBRs.EOF
                    If IsNumeric(SUBRs.Fields("가격")) = True Then
                        Print #FHandle, "["; SUBRs.Fields("구분코드"); ":"; SUBRs.Fields("품명"); Tab; Tab; ":"; SUBRs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(SUBRs.Fields("가격"))) * 0.8)); "]"; Tab; "["; CStr((CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.8) * 0.01) * 100)); "]"
                    
                        dblPrice = 0
                        dblPrice = (CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.7) * 0.01) * 100)        ' 30% 할인을 적용한다.
                        
                        Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                        Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & SUBRs.Fields("구분코드") & "', '"
                        Query = Query & SUBRs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                        ADOCon.Execute Query
                    Else
                        Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                        Query = Query & "[" & SUBRs.Fields("구분코드") & ":" & SUBRs.Fields("품명") & ":" & SUBRs.Fields("가격") & "]"
                        MsgBox Query, vbCritical, "경고"
                    End If
                    SUBRs.MoveNext
                Loop
                SUBRs.Close
                
                
'               나머지 품목도 할인 코드에 넣어주어야 정상 동작된다.
                Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                Query = Query & " WHERE not ( left(구분코드,1) = 'f' or left(구분코드,1) = 'g' or  left(구분코드,1) = 'r'  )"
                Set SUBRs = New ADODB.Recordset
                SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
                
                Do While Not SUBRs.EOF
                    If IsNumeric(SUBRs.Fields("가격")) = True Then
                        Print #FHandle, "["; SUBRs.Fields("구분코드"); ":"; SUBRs.Fields("품명"); Tab; Tab; ":"; SUBRs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(SUBRs.Fields("가격"))) * 0.8)); "]"; Tab; "["; CStr((CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.8) * 0.01) * 100)); "]"
                    
                        dblPrice = Val(CStr(SUBRs.Fields("가격")))
                        
                        Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                        Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & SUBRs.Fields("구분코드") & "', '"
                        Query = Query & SUBRs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                        ADOCon.Execute Query
                    Else
                        Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                        Query = Query & "[" & SUBRs.Fields("구분코드") & ":" & SUBRs.Fields("품명") & ":" & SUBRs.Fields("가격") & "]"
                        MsgBox Query, vbCritical, "경고"
                    End If
                    SUBRs.MoveNext
                Loop
                
                SUBRs.Close
                Set SUBRs = Nothing
                
                Close #FHandle
            End If
        End If
    End If
    
    Check_이마트할인_20071114 = True

    On Error GoTo 0
    
    Exit Function

Check_일반할인_Error:
    Resume
    
    Check_이마트할인_20071114 = False
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Check_이마트할인_20071114 of Module Global"
End Function




Public Function Check_일반할인_20071206() As Boolean
    Dim sMstCode    As String
    Dim sDay        As String
    Dim dblPrice    As Double
    Dim FHandle     As Integer
    
    On Error GoTo Check_일반할인_Error
    
    Check_일반할인_20071206 = False
    
    ' 지사 코드/ 대리점 코드 설정
    sMstCode = 대리점정보.MasterCode
    
    If sMstCode = "1001" Then
        Select Case 대리점정보.대리점번호
            Case "355"
                
                sDay = Format(Date, "YYYY-MM-DD")
                If sDay >= "2007-12-06" And sDay <= "2007-12-16" Then
                        
                    ' 해당 일자에 처음 실행할 경우 해당 파일이 없어서 실해되고 2번째부터는 실행되지 않느다.
                    ' 해당 기간이 지나면 자동으로 프로그램에서 사용하지 않는다.
                    If Dir(App.Path & "\20071206_01.TXT", vbDirectory) = "" Then
                        ' 다음 이중 실행되지 않도록 파일을 생성한다.
                        FHandle = FreeFile
                        Open App.Path & "\" & "\20071206_01.TXT" For Append As FHandle
                        Print #FHandle, Now
                        
                        Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                        Query = Query & " WHERE   ( left(구분코드,1) = 'f' or  left(구분코드,1) = 'g' or  left(구분코드,1) = 'r'  )"
                        Set SUBRs = New ADODB.Recordset
                        SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
                        
                        ' 이전 자료를 모두 지운다.
                        Query = "DELETE FROM 할인정보 WHERE 시작일 = '20071010' AND 종료일 = '20071025' "
                        ADOCon.Execute Query
                        
                        Do While Not SUBRs.EOF
                            If IsNumeric(SUBRs.Fields("가격")) = True Then
                                Print #FHandle, "["; SUBRs.Fields("구분코드"); ":"; SUBRs.Fields("품명"); Tab; Tab; ":"; SUBRs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(SUBRs.Fields("가격"))) * 0.7)); "]"; Tab; "["; CStr((CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.7) * 0.01) * 100)); "]"
                            
                                dblPrice = 0
                                dblPrice = (CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.7) * 0.01) * 100)
                                
                                Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                                Query = Query & " VALUES ('20071206', '20071216', '" & SUBRs.Fields("구분코드") & "', '"
                                Query = Query & SUBRs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                                ADOCon.Execute Query
                            Else
                                Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                                Query = Query & "[" & SUBRs.Fields("구분코드") & ":" & SUBRs.Fields("품명") & ":" & SUBRs.Fields("가격") & "]"
                                MsgBox Query, vbCritical, "경고"
                            End If
                            
                            SUBRs.MoveNext
                        Loop
                        
                        SUBRs.Close
                        
                        Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                        Query = Query & " WHERE  not  ( left(구분코드,1) = 'f' or  left(구분코드,1) = 'g' or  left(구분코드,1) = 'r'  )"
                        Set SUBRs = New ADODB.Recordset
                        SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
                        
                        Do While Not SUBRs.EOF
                            If IsNumeric(SUBRs.Fields("가격")) = True Then
                                Print #FHandle, "["; SUBRs.Fields("구분코드"); ":"; SUBRs.Fields("품명"); Tab; Tab; ":"; SUBRs.Fields("가격"); "]"; Tab; "["; CStr(SUBRs.Fields("가격")); "]"; Tab; "["; CStr(SUBRs.Fields("가격")); "]"
                            
                                dblPrice = Val(CStr(SUBRs.Fields("가격")))
                                
                                Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                                Query = Query & " VALUES ('20071206', '20071216', '" & SUBRs.Fields("구분코드") & "', '"
                                Query = Query & SUBRs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                                ADOCon.Execute Query
                            Else
                                Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                                Query = Query & "[" & SUBRs.Fields("구분코드") & ":" & SUBRs.Fields("품명") & ":" & SUBRs.Fields("가격") & "]"
                                MsgBox Query, vbCritical, "경고"
                            End If
                            SUBRs.MoveNext
                        Loop
                        
                        SUBRs.Close
                        
                        Close #FHandle
                    End If
                End If
            Case Else
            
        End Select
    End If
    
    
    Check_일반할인_20071206 = True

    On Error GoTo 0
    
    Exit Function

Check_일반할인_Error:
    Resume
    Check_일반할인_20071206 = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Check_일반할인_20071206 of Module Global"
End Function


Public Function Check_일반할인_20080320() As Boolean
    Dim sMstCode    As String
    Dim sDay        As String
    Dim dblPrice    As Double
    Dim FHandle     As Integer

    On Error GoTo Check_일반할인_Error
    
    Check_일반할인_20080320 = False
    
    ' 지사 코드/ 대리점 코드 설정
    sMstCode = 대리점정보.MasterCode
    
    If sMstCode = "1002" Then
        Select Case 대리점정보.대리점번호
            Case "044"
                
                sDay = Format(Date, "YYYY-MM-DD")
                If sDay >= "2008-03-20" And sDay <= "2008-03-26" Then
                        
                    ' 해당 일자에 처음 실행할 경우 해당 파일이 없어서 실해되고 2번째부터는 실행되지 않느다.
                    ' 해당 기간이 지나면 자동으로 프로그램에서 사용하지 않는다.
                    If Dir(App.Path & "\20080320_01.TXT", vbDirectory) = "" Then
                        ' 다음 이중 실행되지 않도록 파일을 생성한다.
                        FHandle = FreeFile
                        Open App.Path & "\" & "\20080320_01.TXT" For Append As FHandle
                        Print #FHandle, Now
                        
                        Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                        Set SUBRs = New ADODB.Recordset
                        SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
                        
                        ' 이전 자료를 모두 지운다.
                        Query = "DELETE FROM 할인정보 WHERE 시작일 = '20080320' AND 종료일 = '20080326' "
                        ADOCon.Execute Query
                        
                        Do While Not SUBRs.EOF
                            If IsNumeric(SUBRs.Fields("가격")) = True Then
                                Print #FHandle, "["; SUBRs.Fields("구분코드"); ":"; SUBRs.Fields("품명"); Tab; Tab; ":"; SUBRs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(SUBRs.Fields("가격"))) * 0.8)); "]"; Tab; "["; CStr((CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.8) * 0.01) * 100)); "]"
                            
                                dblPrice = 0
                                dblPrice = (CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.8) * 0.01) * 100)
                                
                                Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                                Query = Query & " VALUES ('20080320', '20080326', '" & SUBRs.Fields("구분코드") & "', '"
                                Query = Query & SUBRs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                                ADOCon.Execute Query
                            Else
                                Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                                Query = Query & "[" & SUBRs.Fields("구분코드") & ":" & SUBRs.Fields("품명") & ":" & SUBRs.Fields("가격") & "]"
                                MsgBox Query, vbCritical, "경고"
                            End If
                            SUBRs.MoveNext
                        Loop
                        
                        SUBRs.Close
                        Set SUBRs = Nothing
                        
                        Close #FHandle
                    End If
                End If
            Case Else
            
        End Select
    End If
       
    Check_일반할인_20080320 = True

    On Error GoTo 0
    
    Exit Function

Check_일반할인_Error:
    Resume
    Check_일반할인_20080320 = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Check_일반할인_20080320 of Module Global"
End Function


'이마트행사품목: 상의류( f코드 ), 하의류( g코드 ), 스커트류( r코드 )만 20% 할인 행사를 진행하며 나머지 품목은 제외 함
'> 기간은 11.1~11.15일까지
Public Function Check_이마트할인_20081112() As Boolean
    Dim sDay        As String
    Dim dblPrice    As Double
    Dim FHandle     As Integer
    
    Dim sStartDate  As String
    Dim sEndDate    As String
    
    On Error GoTo Check_일반할인_Error
    
    Check_이마트할인_20081112 = False
    
    sStartDate = "20081030"
    sEndDate = "20081112"
    
    If Check_이마트할인대상확인_20081112 = True Then
                
        
        
        ' 이전 잘못된 내용을 지운다.
        Call ADOCon.Execute("delete  From 할인정보")
        
        If Dir(App.Path & "\20081112.TXT", vbDirectory) <> "" Then
            Kill App.Path & "\20081112.TXT"
        End If
        
        sDay = Format(Date, "YYYY-MM-DD")
        If sDay >= Format(sStartDate, "@@@@-@@-@@") And sDay <= Format(sEndDate, "@@@@-@@-@@") Then
                
            ' 해당 일자에 처음 실행할 경우 해당 파일이 없어서 실해되고 2번째부터는 실행되지 않느다.
            ' 해당 기간이 지나면 자동으로 프로그램에서 사용하지 않는다.
            If Dir(App.Path & "\" & sEndDate & ".TXT", vbDirectory) = "" Then
                ' 다음 이중 실행되지 않도록 파일을 생성한다.
                FHandle = FreeFile
                Open App.Path & "\" & sEndDate & ".TXT" For Append As FHandle
                Print #FHandle, Now
                
'                f코드(상의류) > 20% 할인
'                g코드(하의류) > 20% 할인
'                r코드(스커트류) > 20% 할인
                Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                Query = Query & " WHERE (  left(구분코드,1) = 'f' or left(구분코드,1) = 'g' or  left(구분코드,1) = 'r'  )"
                Set SUBRs = New ADODB.Recordset
                SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
                
                ' 이전 자료를 모두 지운다.
                Query = "DELETE FROM 할인정보 WHERE 시작일 = '" & sStartDate & "' AND 종료일 = '" & sEndDate & "' "
                ADOCon.Execute Query
                
                Do While Not SUBRs.EOF
                    If IsNumeric(SUBRs.Fields("가격")) = True Then
                        Print #FHandle, "["; SUBRs.Fields("구분코드"); ":"; SUBRs.Fields("품명"); Tab; Tab; ":"; SUBRs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(SUBRs.Fields("가격"))) * 0.8)); "]"; Tab; "["; CStr((CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.8) * 0.01) * 100)); "]"
                    
                        dblPrice = 0
                        dblPrice = (CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.8) * 0.01) * 100)        ' 20% 할인을 적용한다.
                        
                        Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                        Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & SUBRs.Fields("구분코드") & "', '"
                        Query = Query & SUBRs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                        ADOCon.Execute Query
                    Else
                        Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                        Query = Query & "[" & SUBRs.Fields("구분코드") & ":" & SUBRs.Fields("품명") & ":" & SUBRs.Fields("가격") & "]"
                        MsgBox Query, vbCritical, "경고"
                    End If
                    
                    SUBRs.MoveNext
                Loop
                SUBRs.Close
                
                
'               나머지 품목도 할인 코드에 넣어주어야 정상 동작된다.
                Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                Query = Query & " WHERE not ( left(구분코드,1) = 'f' or left(구분코드,1) = 'g' or  left(구분코드,1) = 'r'  )"
                Set SUBRs = New ADODB.Recordset
                SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
                
                Do While Not SUBRs.EOF
                    If IsNumeric(SUBRs.Fields("가격")) = True Then
                        Print #FHandle, "["; SUBRs.Fields("구분코드"); ":"; SUBRs.Fields("품명"); Tab; Tab; ":"; SUBRs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(SUBRs.Fields("가격"))) * 0.8)); "]"; Tab; "["; CStr((CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.8) * 0.01) * 100)); "]"
                    
                        dblPrice = Val(CStr(SUBRs.Fields("가격")))
                        
                        Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                        Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & SUBRs.Fields("구분코드") & "', '"
                        Query = Query & SUBRs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                        ADOCon.Execute Query
                    Else
                        Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                        Query = Query & "[" & SUBRs.Fields("구분코드") & ":" & SUBRs.Fields("품명") & ":" & SUBRs.Fields("가격") & "]"
                        MsgBox Query, vbCritical, "경고"
                    End If
                    SUBRs.MoveNext
                Loop
                SUBRs.Close
                
                
                Close #FHandle
            End If
        End If
    End If
    
    
    Check_이마트할인_20081112 = True

    On Error GoTo 0
    Exit Function

Check_일반할인_Error:
    Resume
    Check_이마트할인_20081112 = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Check_이마트할인_20081112 of Module Global"
End Function

'일반매장 행사 제외품목: 운동화,구두류( a코드 전체 ), 침구류( k코드 전체 ), 와이셔츠( m코드 중 m00, m01 코드 만 제외 바람 )
'기간은  11.1 ~ 11.10일까지
Public Function Check_일반할인_20081110() As Boolean
    Dim sDay        As String
    Dim dblPrice    As Double
    Dim FHandle     As Integer
    
    Dim sStartDate  As String
    Dim sEndDate    As String
    
    

    On Error GoTo Check_일반할인_Error
    Check_일반할인_20081110 = False
    
    
    sStartDate = "20081101"
    sEndDate = "20081110"
    
    If Check_일반매장할인대상확인_20081110 = True Then
                
        sDay = Format(Date, "YYYY-MM-DD")
        If sDay >= Format(sStartDate, "@@@@-@@-@@") And sDay <= Format(sEndDate, "@@@@-@@-@@") Then
                
            ' 해당 일자에 처음 실행할 경우 해당 파일이 없어서 실해되고 2번째부터는 실행되지 않느다.
            ' 해당 기간이 지나면 자동으로 프로그램에서 사용하지 않는다.
            If Dir(App.Path & "\" & sEndDate & ".TXT", vbDirectory) = "" Then
                ' 다음 이중 실행되지 않도록 파일을 생성한다.
                FHandle = FreeFile
                Open App.Path & "\" & sEndDate & ".TXT" For Append As FHandle
                Print #FHandle, Now
                
'                운동화,구두류( a코드 전체 ), 침구류( k코드 전체 ), 와이셔츠( m코드 중 m00, m01 코드 만 제외 바람 )
                Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                Query = Query & " WHERE not ( left(구분코드,3) = 'm00'  or  left(구분코드,3) = 'm01' or  left(구분코드,1) = 'a' or  left(구분코드,1) = 'k'  )"
                Set SUBRs = New ADODB.Recordset
                SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
                
                
                ' 이전 자료를 모두 지운다.
                Query = "DELETE FROM 할인정보 WHERE 시작일 = '" & sStartDate & "' AND 종료일 = '" & sEndDate & "' "
                ADOCon.Execute Query
                
                Do While Not SUBRs.EOF
                    If IsNumeric(SUBRs.Fields("가격")) = True Then
                        Print #FHandle, "["; SUBRs.Fields("구분코드"); ":"; SUBRs.Fields("품명"); Tab; Tab; ":"; SUBRs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(SUBRs.Fields("가격"))) * 0.8)); "]"; Tab; "["; CStr((CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.8) * 0.01) * 100)); "]"
                    
                        dblPrice = 0
                        dblPrice = (CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.8) * 0.01) * 100)        ' 20% 할인을 적용한다.
                        
                        Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                        Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & SUBRs.Fields("구분코드") & "', '"
                        Query = Query & SUBRs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                        ADOCon.Execute Query
                    Else
                        Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                        Query = Query & "[" & SUBRs.Fields("구분코드") & ":" & SUBRs.Fields("품명") & ":" & SUBRs.Fields("가격") & "]"
                        MsgBox Query, vbCritical, "경고"
                    End If
                    SUBRs.MoveNext
                Loop
                SUBRs.Close
                
                
'               나머지 품목도 할인 코드에 넣어주어야 정상 동작된다.
                Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                Query = Query & " WHERE ( left(구분코드,3) = 'm00'  or  left(구분코드,3) = 'm01' or  left(구분코드,1) = 'a' or  left(구분코드,1) = 'k'  )"
                Set SUBRs = New ADODB.Recordset
                SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
                
                Do While Not SUBRs.EOF
                    If IsNumeric(SUBRs.Fields("가격")) = True Then
                        Print #FHandle, "["; SUBRs.Fields("구분코드"); ":"; SUBRs.Fields("품명"); Tab; Tab; ":"; SUBRs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(SUBRs.Fields("가격"))) * 0.8)); "]"; Tab; "["; CStr((CLng((Val(CStr(SUBRs.Fields("가격"))) * 0.8) * 0.01) * 100)); "]"
                    
                        dblPrice = Val(CStr(SUBRs.Fields("가격")))
                        
                        Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                        Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & SUBRs.Fields("구분코드") & "', '"
                        Query = Query & SUBRs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                        ADOCon.Execute Query
                    Else
                        Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                        Query = Query & "[" & SUBRs.Fields("구분코드") & ":" & SUBRs.Fields("품명") & ":" & SUBRs.Fields("가격") & "]"
                        MsgBox Query, vbCritical, "경고"
                    End If
                    SUBRs.MoveNext
                Loop
                SUBRs.Close
                
                
                Close #FHandle
            End If
        End If
    End If
    
    
    Check_일반할인_20081110 = True

    On Error GoTo 0
    
    Exit Function

Check_일반할인_Error:
    Resume
    Check_일반할인_20081110 = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Check_일반할인_20081110 of Module Global"
End Function


'일반매장 행사 품목: 스웨터류(u), 코트류(i),점퍼류(d),이불류(k),가죽(b),모피류(n)
'기간은  2009.06.08 ~ 2009.06.15일까지
Public Function Check_일반할인_20090615() As Boolean
    Dim sDay        As String
    Dim dblPrice    As Double
    Dim FHandle     As Integer
    
    Dim sStartDate  As String
    Dim sEndDate    As String
    
    Dim nPercent    As Single
    
    
    On Error GoTo Check_일반할인_Error
    Check_일반할인_20090615 = False
    
    nPercent = 0.8  ' 실제 받을 금액 20% 일경우 0.8 입력
    sStartDate = "20090608"
    sEndDate = IIf(대리점정보.StoreCode <> "100084", "20090615", "20090617")  '10084 이마트경산점만 17일까지 세일
    
    If Check_일반매장할인대상확인_20090615 = True Then
                
        sDay = Format(Date, "YYYY-MM-DD")
        If sDay >= Format(sStartDate, "@@@@-@@-@@") And sDay <= Format(sEndDate, "@@@@-@@-@@") Then
                
            ' 해당 일자에 처음 실행할 경우 해당 파일이 없어서 실해되고 2번째부터는 실행되지 않는다.
            ' 해당 기간이 지나면 자동으로 프로그램에서 사용하지 않는다.
            If Dir(App.Path & "\" & sEndDate & ".TXT", vbDirectory) = "" Then
                ' 다음 이중 실행되지 않도록 파일을 생성한다.
                FHandle = FreeFile
                Open App.Path & "\" & sEndDate & ".TXT" For Append As FHandle
                Print #FHandle, Now
                
                If 대리점정보.MasterCode = "1020" Then
                    '일반매장 행사 품목: 스웨터류(u), 코트류(i),점퍼류(d),이불류(k)
                    Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                    Query = Query & " WHERE left(구분코드,1) = 'u'  or  left(구분코드,1) = 'i' or  left(구분코드,1) = 'd' or  left(구분코드,1) = 'k' "
                Else
                    '일반매장 행사 품목: 스웨터류(u), 코트류(i),점퍼류(d),이불류(k),가죽(b),모피류(n)
                    Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                    Query = Query & " WHERE left(구분코드,1) = 'u'  or  left(구분코드,1) = 'i' or  left(구분코드,1) = 'd' or  left(구분코드,1) = 'k' "
                    Query = Query & "       or  left(구분코드,1) = 'b'  or  left(구분코드,1) = 'n' "
                End If
                Set SUBRs = New ADODB.Recordset
                SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
                
                ' 이전 자료를 모두 지운다.
                Query = "DELETE FROM 할인정보 " 'WHERE 시작일 = '" & sStartDate & "' AND 종료일 = '" & sEndDate & "' "
                ADOCon.Execute Query
                
                Do While Not SUBRs.EOF
                    If IsNumeric(SUBRs.Fields("가격")) = True Then
                        Print #FHandle, "["; SUBRs.Fields("구분코드"); ":"; SUBRs.Fields("품명"); Tab; Tab; ":"; SUBRs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(SUBRs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(SUBRs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"
                    
                        dblPrice = 0
                        dblPrice = (CLng((Val(CStr(SUBRs.Fields("가격"))) * nPercent) * 0.01) * 100)        '  할인을 적용한다.
                        
                        Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                        Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & SUBRs.Fields("구분코드") & "', '"
                        Query = Query & SUBRs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                        ADOCon.Execute Query
                    Else
                        Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                        Query = Query & "[" & SUBRs.Fields("구분코드") & ":" & SUBRs.Fields("품명") & ":" & SUBRs.Fields("가격") & "]"
                        MsgBox Query, vbCritical, "경고"
                    End If
                    SUBRs.MoveNext
                Loop
                SUBRs.Close
                
                
                If 대리점정보.MasterCode = "1020" Then
    '               나머지 품목도 할인 코드에 넣어주어야 정상 동작된다.
                    Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                    Query = Query & " WHERE not(left(구분코드,1) = 'u'  or  left(구분코드,1) = 'i' or  left(구분코드,1) = 'd' or  left(구분코드,1) = 'k' ) "
                Else
    '               나머지 품목도 할인 코드에 넣어주어야 정상 동작된다.
                    Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                    Query = Query & " WHERE not(left(구분코드,1) = 'u'  or  left(구분코드,1) = 'i' or  left(구분코드,1) = 'd' or  left(구분코드,1) = 'k' "
                    Query = Query & "       or  left(구분코드,1) = 'b'  or  left(구분코드,1) = 'n') "
                End If
                Set SUBRs = New ADODB.Recordset
                SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
                
                Do While Not SUBRs.EOF
                    If IsNumeric(SUBRs.Fields("가격")) = True Then
                        Print #FHandle, "["; SUBRs.Fields("구분코드"); ":"; SUBRs.Fields("품명"); Tab; Tab; ":"; SUBRs.Fields("가격"); "]"; Tab; "["; CStr(SUBRs.Fields("가격")); "]"; Tab; "["; CStr(SUBRs.Fields("가격")); "]"
                    
                        dblPrice = Val(CStr(SUBRs.Fields("가격")))
                        
                        Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                        Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & SUBRs.Fields("구분코드") & "', '"
                        Query = Query & SUBRs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                        ADOCon.Execute Query
                    Else
                        Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                        Query = Query & "[" & SUBRs.Fields("구분코드") & ":" & SUBRs.Fields("품명") & ":" & SUBRs.Fields("가격") & "]"
                        MsgBox Query, vbCritical, "경고"
                    End If
                    SUBRs.MoveNext
                Loop
                SUBRs.Close
                Set SUBRs = Nothing
                
                Close #FHandle
            End If
        End If
    End If
    
    
    Check_일반할인_20090615 = True

    On Error GoTo 0
    Exit Function

Check_일반할인_Error:
    Resume
    Check_일반할인_20090615 = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Check_일반할인_20090615 of Module Global"
End Function

'====================================================================================================
' Procedure : SetTableDefaultSendData
' DateTime  : 2008-05-03 01:11
' Author    : pds2004
' Last Edit :
' Return Val:
'----------------------------------------------------------------------------------------------------
' 설     명 : 최초 아무런 자료가 없을 경우 초기 값을 저장한다.
'====================================================================================================
Public Function SetTableDefaultSendData() As Boolean
    Dim MyHost  As ADODB.Connection
    
    On Error GoTo SendTableDateCheck_Error
    
    SetTableDefaultSendData = False
    
    If Trim(대리점정보.StoreCode) = "000000" Then
        MsgBox "대리점 정보가 올바르지 않습니다.", vbCritical, "경고"
        Exit Function
    End If
    
    If ConnectMasterCheck(MyHost) = True Then
        ' 최초 아무것도 없을 경우 초기 기준일자를 입력한다.
        Query = "SELECT STORE_CD FROM 가맹점전송정보 WHERE STORE_CD = '" & 대리점정보.StoreCode & "' "
        Set SUBRs = New ADODB.Recordset
        SUBRs.Open Query, MyHost, adOpenStatic, adLockOptimistic
        
        If SUBRs.EOF = True Then
            Query = "INSERT INTO 가맹점전송정보(STORE_CD, "
            Query = Query & " 고객정보, 고객정보기준일자, 고객정보전송일자, "
            Query = Query & " 대리점정보, 대리점정보기준일자, 대리점정보전송일자, "
            Query = Query & " 마일리지스토리, 마일리지스토리기준일자, 마일리지스토리전송일자, "
            Query = Query & " 마일리지현황, 마일리지현황기준일자, 마일리지현황전송일자, "
            Query = Query & " 목요세일, 목요세일기준일자, 목요세일전송일자, "
            Query = Query & " 미수회수정보, 미수회수정보기준일자, 미수회수정보전송일자, "
            Query = Query & " 입출고, 입출고기준일자, 입출고전송일자, "
            Query = Query & " 참조코드, 참조코드기준일자, 참조코드전송일자, "
            Query = Query & " 할인정보, 할인정보기준일자, 할인정보전송일자) "
            Query = Query & " VALUES('" & 대리점정보.StoreCode & "',"
            Query = Query & " 'N','2008-01-01','', "
            Query = Query & " 'N','2008-01-01','', "
            Query = Query & " 'N','2008-01-01','', "
            Query = Query & " 'N','2008-01-01','', "
            Query = Query & " 'N','2008-01-01','', "
            Query = Query & " 'N','2008-01-01','', "
            Query = Query & " 'N','2008-01-01','', "
            Query = Query & " 'N','2008-01-01','', "
            Query = Query & " 'N','2008-01-01','') "
            MyHost.Execute Query
        End If
    End If
    
    Set MyHost = Nothing
    Set SUBRs = Nothing
    
    SetTableDefaultSendData = True
    
    On Error GoTo 0
    
    Exit Function

SendTableDateCheck_Error:
    
    Set MyHost = Nothing
    Set SUBRs = Nothing

    SetTableDefaultSendData = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SetTableDefaultSendData of Module Global"
End Function

'====================================================================================================
' Procedure : SendTableDateSave
' DateTime  : 2008-05-03 01:11
' Author    : pds2004
' Last Edit :
' Return Val:
'----------------------------------------------------------------------------------------------------
' 설     명 : 전달된 테이블에 전송 "Y"를 설정한다.
'====================================================================================================
Public Function SendProgramVersion() As Boolean
    Dim MyHost  As ADODB.Connection
    
    On Error GoTo SendProgramVersion_Error
    
    If ConnectMasterCheck(MyHost) = False Then
        Set MyHost = Nothing
        Exit Function
    End If
    
    SendProgramVersion = True
    
    Query = "UPDATE 가맹점대리점정보 SET Ver = '" & Program_Version & "'  "
    Query = Query & " WHERE StoreCode = '" & 대리점정보.StoreCode & "' "
    MyHost.Execute Query
    
    Set MyHost = Nothing
    
    On Error GoTo 0
    
    Exit Function

SendProgramVersion_Error:

    Set MyHost = Nothing
    SendProgramVersion = False
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SendProgramVersion of Module Global"
End Function

'====================================================================================================
' Procedure : SendTableDateSave
' DateTime  : 2008-05-03 01:11
' Author    : pds2004
' Last Edit :
' Return Val:
'----------------------------------------------------------------------------------------------------
' 설     명 : 전달된 테이블에 전송 "Y"를 설정한다.
'====================================================================================================
Private Function SendTableDateSave(ByVal sTableName As String, ByRef MyHost As ADODB.Connection) As Boolean
    On Error GoTo SendTableDateSave_Error
        
    SendTableDateSave = True
    
    sTableName = Replace(sTableName, "가맹점", "")
    
    Query = "UPDATE 가맹점전송정보 SET " & sTableName & "전송일자 = '" & Format(Date, "yyyy-MM-dd") & "', "
    Query = Query & sTableName & " = 'Y'"
    Query = Query & " WHERE STORE_CD = '" & 대리점정보.StoreCode & "' "
    MyHost.Execute Query
    
    On Error GoTo 0
    
    Exit Function

SendTableDateSave_Error:

    SendTableDateSave = False
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SendTableDateSave of Module Global"
End Function
 
'====================================================================================================
' Procedure : SendTableDateCheck
' DateTime  : 2008-05-03 01:11
' Author    : pds2004
' Last Edit :
' Return Val:
'----------------------------------------------------------------------------------------------------
' 설     명 : 기준일자와 전송일자를 확인하여 전송일자 이후의 일자를 리턴한다.
'             기준일자 : 2008-01-01 전송일자 : '' 일경우 2008-01-01리턴
'             기준일자 : 2008-01-01 전송일자 : '2008-04-20' 일경우 2008-04-21을 리턴
'====================================================================================================
Public Function SendTableDateCheck(ByVal sTableName As String, ByRef MyHost As ADODB.Connection, ByRef SendYN As String) As String
    On Error GoTo SendTableDateCheck_Error
    
    SendYN = "N"
    SendTableDateCheck = "2008-01-01"
    
    sTableName = Replace(sTableName, "가맹점", "")
    
    ' 최초 아무것도 없을 경우 초기 기준일자를 입력한다.
    Query = "SELECT " & sTableName & "전송일자, " & sTableName
    Query = Query & " FROM 가맹점전송정보"
    Query = Query & " WHERE STORE_CD = '" & 대리점정보.StoreCode & "' "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, MyHost, adOpenStatic, adLockOptimistic
        
    If Not SUBRs.EOF Then
        SendYN = IIf(Trim(SUBRs.Fields(1) & "") <> "Y", "N", "Y")
        
        If Trim(SUBRs.Fields(0) & "") = "" Then
            SendTableDateCheck = "2008-08-01"
        Else
            SendTableDateCheck = Trim(SUBRs.Fields(0) & "")
        End If
    End If
    SUBRs.Close
    Set SUBRs = Nothing

    On Error GoTo 0
    
    Exit Function

SendTableDateCheck_Error:

    SendTableDateCheck = False
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SendTableDateCheck of Module Global"
End Function
 
'====================================================================================================
' Procedure : SendTableData
' DateTime  : 2008-05-03 01:03
' Author    : pds2004
' Last Edit :
' Return Val:
'----------------------------------------------------------------------------------------------------
' 설     명 : 각종테이블의 자료를 본사 SQL 서버에 저장한다.
'====================================================================================================
Public Function SendTableData(objPrBar As Object, Optional DaySendYN As Boolean = False) As Boolean
    Dim sSendData   As String
    Dim sYN         As String
    Dim MyHost      As ADODB.Connection
    
    On Error GoTo SendTableData_Error

    If Trim(대리점정보.StoreCode) = "000000" Then
        MsgBox "대리점 정보가 올바르지 않습니다.", vbCritical, "경고"
        Exit Function
    End If
    
    If ConnectMasterCheck(MyHost) = False Then
        Set MyHost = Nothing
        Exit Function
    End If
    
    ' 가맹점 가맹점대리점정보를 전송한다.
    sSendData = SendTableDateCheck("가맹점대리점정보", MyHost, sYN)
    
    If sYN = "N" Then Call SendTable_대리점정보(sSendData, MyHost, objPrBar)
    
    ' 가맹점 가맹점참조코드를 전송한다.
    sSendData = SendTableDateCheck("가맹점참조코드", MyHost, sYN)
    
    If sYN = "N" Then Call SendTable_참조코드(sSendData, MyHost, objPrBar)
    
    '사고품 정보를 전송한다.
    ' 사고품은 입력된 내용이 있을 경우 매일 전송하여야 하기 때문에 "Y" 일경우도전송한다.
    sSendData = SendTableDateCheck("사고품", MyHost, sYN)
    
    Call SendTable_사고품(sSendData, MyHost, objPrBar)
    
    ' 마감 자료를 전송한다.
    ' 기본 30일을 확인한다. 가맹점전송정보에 마감일자 관련 테이블이 없어서 30일 기준으로 처리한다.
    ' DaySendYN = true 일경우 당일날 물건을 받기 위하여 본사연결 화면에 들오롤경우 당일 매출이 발생하지
    ' 않는 내용이 저장되어 다음날 적용되지 않는 문제가 발생하여 처리함.
    Call SendSalesData(Format(DateAdd("d", -7, Date), "yyyy-MM-dd"), IIf(DaySendYN = True, 6, 7))
    
    ' 쿠폰정보 본사 전송
    Call SendTable_쿠폰정보(Format(DateAdd("d", -7, Date), "yyyy-MM-dd"), 7)
    
    '입출고 정보를 전송한다.
    '입출고는 입력된 내용이 있을 경우 매일 전송하여야 하기 때문에 "Y" 일경우도전송한다.
    sSendData = SendTableDateCheck("입출고", MyHost, sYN)
    
    Call SendTable_입출고(sSendData, MyHost, objPrBar)
    
    '메시지 확인 내용을 전송한다.
    Call SendTable_메시지(sSendData, MyHost, objPrBar)
    
    '세트 상품 관련 내용을 전송한다.
    If Format(Date, "yyyyMMdd") <= "20100131" Then Call SendTable_세트응모번호(sSendData, MyHost, objPrBar)
    
    Call SendTable_세트상품정보(sSendData, MyHost, objPrBar)
    
    Set MyHost = Nothing
    
    SendTableData = True

    On Error GoTo 0
    
    Exit Function

SendTableData_Error:

    SendTableData = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SendTableData of Module Global"

End Function

'====================================================================================================
' Procedure : SendTable_고객정보
' DateTime  : 2008-05-03 01:42
' Author    : pds2004
' Last Edit :
' Return Val:
'----------------------------------------------------------------------------------------------------
' 설     명 : 전달된 일자부터 현재일자까지의 고객정보를 본사 SQL 서버에 저장한다.
'====================================================================================================
Private Function SendTable_고객정보(ByVal sStartDate As String, ByRef MyHost As Object, ByRef objPrBar As Object) As Long
    Dim aa As String
        
    On Error GoTo SendTable_고객정보_Error
    
    If IsDate(sStartDate) = False Then
        MsgBox "전달된 일자가 올바르지 않습니다.  [" & sStartDate & "]", vbCritical, "경고"
        
        Exit Function
    End If
     
     SendTable_고객정보 = 0
     
     aa = "시작시간: " & CStr(Now) & vbNewLine
     
    '---------------------------------------------------------
    ' 대리점 코드를 Check한다.
    '---------------------------------------------------------
    Query = "SELECT    고객번호"
    Query = Query & ", 성명"
    Query = Query & ", 전화1"
    Query = Query & ", 전화2"
    Query = Query & ", 주소"
    Query = Query & ", 미수금"
    Query = Query & ", 전송구분"
    Query = Query & ", 카드번호"
    Query = Query & ", 휴대폰"
    Query = Query & " FROM 고객정보"
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    ' 미전송 자료가 없으면 종료를 한다.
    If Not SUBRs.EOF Then
        SUBRs.MoveLast
        
        objPrBar.MAX = SUBRs.RecordCount
        
        SUBRs.MoveFirst
    End If
    
    ' 고객정보가 없으면 종료를 한다.
    Do While Not SUBRs.EOF
        Query = "INSERT INTO 가맹점고객정보 (STORE_CD, 고객번호, 성명, 전화1, 전화2, 주소, 미수금, "
        Query = Query & " 전송구분, 카드번호, 휴대폰, TRANS_CHK, TRANS_DT) "
        Query = Query & " VALUES('" & 대리점정보.StoreCode & "','" & SUBRs.Fields("고객번호") & "', "
        Query = Query & " '" & SUBRs.Fields("성명") & "','" & SUBRs.Fields("전화1") & "', "
        Query = Query & " '" & SUBRs.Fields("전화2") & "','" & Trim(Replace(SUBRs.Fields("주소"), Chr(1), "")) & "', "
        Query = Query & " '" & SUBRs.Fields("미수금") & "','" & SUBRs.Fields("전송구분") & "', "
        Query = Query & " '" & SUBRs.Fields("카드번호") & "','" & SUBRs.Fields("휴대폰") & "', "
        Query = Query & " 'Y','" & Format(Date, "yyyyMMdd") & "') "
        MyHost.Execute Query
        
        SendTable_고객정보 = SendTable_고객정보 + 1
        
        objPrBar.Value = IIf(objPrBar.MAX <= objPrBar.Value, 1, objPrBar.Value + 1)
        
        SUBRs.MoveNext
    Loop
    
    SUBRs.Close
    Set SUBRs = Nothing
    
    aa = aa & "종료시간: " & CStr(Now) & vbNewLine
    aa = aa & "전송수량: " & CStr(SendTable_고객정보) & vbNewLine
     
    'MsgBox aa
    
    On Error GoTo 0
    
    Exit Function

SendTable_고객정보_Error:

    SendTable_고객정보 = False
    
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SendTable_고객정보 of Module Global"
    
    Resume Next
End Function


'====================================================================================================
' Procedure : SendTable_대리점정보
' DateTime  : 2008-05-03 01:42
' Author    : pds2004
' Last Edit :
' Return Val:
'----------------------------------------------------------------------------------------------------
' 설     명 : 전달된 일자부터 현재일자까지의 고객정보를 본사 SQL 서버에 저장한다.
'====================================================================================================
Private Function SendTable_대리점정보(ByVal sStartDate As String, ByRef MyHost As Object, ByRef objPrBar As Object) As Long
    On Error GoTo SendTable_대리점정보_Error
    
    If IsDate(sStartDate) = False Then
        MsgBox "전달된 일자가 올바르지 않습니다.  [" & sStartDate & "]", vbCritical, "경고"
        Exit Function
    End If
     
     SendTable_대리점정보 = 0
     'aa = "시작시간: " & CStr(Now) & vbNewLine
     
    '----------------------------------------------------------------------
    ' 대리점 코드를 Check한다.
    '----------------------------------------------------------------------
    Query = "SELECT * FROM 대리점정보"
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    
    ' 고객정보가 없으면 종료를 한다.
    If Not SUBRs.EOF Then
        '----------------------------------------------------------------------
        Query = "DELETE FROM 가맹점대리점정보"
        Query = Query & " WHERE StoreCode = '" & 대리점정보.StoreCode & "' "
        MyHost.Execute Query
        
    
        Query = "INSERT INTO 가맹점대리점정보 (":
        Query = Query & "StoreCode, ":
        Query = Query & "StoreName, ":
        Query = Query & "StartDate, ":
        Query = Query & "대리점번호, ":
        Query = Query & "대리점색상, ":
        Query = Query & "대리점명, ":
        Query = Query & "수선, ":
        Query = Query & "할인시작일, ":
        Query = Query & "할인종료일, ":
        Query = Query & "일수, ":
        Query = Query & "비율, ":
        Query = Query & "전화1, ":
        Query = Query & "전화2, ":
        Query = Query & "목요세일, ":
        Query = Query & "수선마진, ":
        Query = Query & "프린터, ":
        Query = Query & "일수2, ":
        Query = Query & "운동화마진, ":
        Query = Query & "가죽무스탕마진, ":
        Query = Query & "카페트마진, ":
        Query = Query & "마일리지여부, ":
        Query = Query & "보관증종류, ":
        Query = Query & "특정할인여부, ":
        Query = Query & "특정할인비율, ":
        Query = Query & "고가세탁비율, ":
        Query = Query & "마일리지검사일자, ":
        Query = Query & "마일리지증가구분, ":
        Query = Query & "ServerIP, ":
        Query = Query & "ServerDB, ":
        Query = Query & "ServerUser, ":
        Query = Query & "ServerPass, ":
        Query = Query & "TimeOut, ":
        Query = Query & "TRANS_CHK, ":
        Query = Query & "TRANS_DT) ":
        Query = Query & " VALUES("
        Query = Query & "'" & SUBRs.Fields("StoreCode") & "', "
        Query = Query & "'" & SUBRs.Fields("StoreName") & "', "
        Query = Query & "'" & SUBRs.Fields("StartDate") & "', "
        Query = Query & "'" & SUBRs.Fields("대리점번호") & "', "
        Query = Query & "'" & SUBRs.Fields("대리점색상") & "', "
        Query = Query & "'" & SUBRs.Fields("대리점명") & "', "
        Query = Query & "'" & SUBRs.Fields("수선") & "', "
        Query = Query & "'" & SUBRs.Fields("할인시작일") & "', "
        Query = Query & "'" & SUBRs.Fields("할인종료일") & "', "
        Query = Query & "'" & SUBRs.Fields("일수") & "', "
        Query = Query & "'" & SUBRs.Fields("비율") & "', "
        Query = Query & "'" & SUBRs.Fields("전화1") & "', "
        Query = Query & "'" & SUBRs.Fields("전화2") & "', "
        Query = Query & "'" & SUBRs.Fields("목요세일") & "', "
        Query = Query & "'" & SUBRs.Fields("수선마진") & "', "
        Query = Query & "'" & SUBRs.Fields("프린터") & "', "
        Query = Query & "'" & SUBRs.Fields("일수2") & "', "
        Query = Query & "'" & SUBRs.Fields("운동화마진") & "', "
        Query = Query & "'" & SUBRs.Fields("가죽무스탕마진") & "', "
        Query = Query & "'" & SUBRs.Fields("카페트마진") & "', "
        Query = Query & "'" & SUBRs.Fields("마일리지여부") & "', "
        Query = Query & "'" & SUBRs.Fields("보관증종류") & "', "
        Query = Query & "'" & SUBRs.Fields("특정할인여부") & "', "
        Query = Query & "'" & SUBRs.Fields("특정할인비율") & "', "
        Query = Query & "'" & SUBRs.Fields("고가세탁비율") & "', "
        Query = Query & "'" & SUBRs.Fields("마일리지검사일자") & "', "
        Query = Query & "'" & SUBRs.Fields("마일리지증가구분") & "', "
        Query = Query & "'" & SUBRs.Fields("ServerIP") & "', "
        Query = Query & "'" & SUBRs.Fields("ServerDB") & "', "
        Query = Query & "'" & SUBRs.Fields("ServerUser") & "', "
        Query = Query & "'" & SUBRs.Fields("ServerPass") & "', "
        Query = Query & "'" & SUBRs.Fields("TimeOut") & "', "
        Query = Query & "'" & "Y" & "', "
        Query = Query & "'" & Format(Date, "yyyyMMdd") & "') "
        MyHost.Execute Query & Query
        
        Call SendTableDateSave("가맹점대리점정보", MyHost)
    End If
    
    SUBRs.Close
    Set SUBRs = Nothing
    
    ' aa = aa & "종료시간: " & CStr(Now) & vbNewLine
    ' aa = aa & "전송수량: " & CStr(SendTable_대리점정보) & vbNewLine
    ' MsgBox aa
    
    On Error GoTo 0
    
    Exit Function

SendTable_대리점정보_Error:

    SendTable_대리점정보 = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SendTable_대리점정보 of Module Global"
    Resume Next
End Function

'====================================================================================================
' Procedure : SendTable_참조코드
' DateTime  : 2008-05-03 01:42
' Author    : pds2004
' Last Edit :
' Return Val:
'----------------------------------------------------------------------------------------------------
' 설     명 : 전달된 일자부터 현재일자까지의 고객정보를 본사 SQL 서버에 저장한다.
'====================================================================================================
Private Function SendTable_참조코드(ByVal sStartDate As String, ByRef MyHost As Object, ByRef objPrBar As Object) As Long
    On Error GoTo SendTable_참조코드_Error
    
    If IsDate(sStartDate) = False Then
        MsgBox "전달된 일자가 올바르지 않습니다.  [" & sStartDate & "]", vbCritical, "경고"
        Exit Function
    End If
     
     SendTable_참조코드 = 0
     'aa = "시작시간: " & CStr(Now) & vbNewLine
     
    '--------------------------------------------------------------------
    ' 대리점 코드를 Check한다.
    '--------------------------------------------------------------------
    Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드"
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    '--------------------------------------------------------------------
    '
    '--------------------------------------------------------------------
    Query = "DELETE FROM 가맹점참조코드 "
    Query = Query & " WHERE STORE_CD = '" & 대리점정보.StoreCode & "' "
    MyHost.Execute Query
    
    ' 미전송 자료가 없으면 종료를 한다.
    If Not SUBRs.EOF Then
        SUBRs.MoveLast
        objPrBar.MAX = SUBRs.RecordCount
        SUBRs.MoveFirst
    End If
    
    
    ' 고객정보가 없으면 종료를 한다.
    Do While Not SUBRs.EOF
        Query = "INSERT INTO 가맹점참조코드 ("
        Query = Query & "STORE_CD, ":
        Query = Query & "구분코드, ":
        Query = Query & "품명, ":
        Query = Query & "가격, ":
        Query = Query & "TRANS_CHK, ":
        Query = Query & "TRANS_DT) ":
        Query = Query & "VALUES("
        Query = Query & "'" & 대리점정보.StoreCode & "', "
        Query = Query & "'" & SUBRs.Fields("구분코드") & "', "
        Query = Query & "'" & SUBRs.Fields("품명") & "', "
        Query = Query & "" & SUBRs.Fields("가격") & ", "
        Query = Query & "'" & "Y" & "', "
        Query = Query & "'" & Format(Date, "yyyyMMdd") & "') "
        MyHost.Execute Query
        
        SUBRs.MoveNext
        
        If objPrBar.Value < objPrBar.MAX Then
            objPrBar.Value = objPrBar.Value + 1
        End If
    Loop
    
    Call SendTableDateSave("가맹점참조코드", MyHost)
    
    SUBRs.Close
    Set SUBRs = Nothing
    
    ' aa = aa & "종료시간: " & CStr(Now) & vbNewLine
    ' aa = aa & "전송수량: " & CStr(SendTable_대리점정보) & vbNewLine
    ' MsgBox aa
    
    On Error GoTo 0
    
    Exit Function

SendTable_참조코드_Error:
    SendTable_참조코드 = False
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SendTable_참조코드 of Module Global"
    
    Resume Next
End Function

'====================================================================================================
' Procedure : SendSalesData
' DateTime  : 2008-04-15 20:29
' Author    : pds2004
' Last Edit :
' Return Val:
'----------------------------------------------------------------------------------------------------
' 설     명 : 기본적으로 최근 1주일 내용을 전송한다.
'====================================================================================================
Public Function SendTable_쿠폰정보(ByVal sSendDate As String, Optional iSendDay As Integer = 0) As Boolean
    Dim iDay    As Integer
    Dim iTempDay    As String
    
    Dim MyHost  As ADODB.Connection
    
    Dim sData(23)   As String
    Dim Scode(1)    As String
    
    On Error GoTo SendTable_쿠폰정보_Error

    If Trim(대리점정보.StoreCode) = "000000" Then
        MsgBox "대리점 정보가 올바르지 않습니다.", vbCritical, "경고"
        
        'Set ADORset = Nothing
        Exit Function
    End If
        
    If ConnectMasterCheck(MyHost) = True Then
        For iDay = 0 To iSendDay
            '실제 전송일자
            iTempDay = DateAdd("d", iDay, sSendDate)
                
            ' 해당 일자의 지사및 체인점 코드를 알아온다.
            Erase Scode
            
            If GetMasterStoreFromToDate(Format(iTempDay, "yyyyMMdd"), Scode(0), Scode(1)) = False Then
                ' 본사에서 확인하여 처리할 수 있도록 하기위하여
                ' 지사 정보가 없도라도 전송 처리한다.
            End If
                
            'Set rsTempTb = MyDB.OpenRecordset("SELECT * FROM 쿠폰자료 WHERE 접수일자 = '" & Format(iTempDay, "yyyyMMdd") & "' ")
        
            Query = "SELECT * FROM 쿠폰자료 WHERE 접수일자 = '" & Format(iTempDay, "yyyyMMdd") & "' "
            Set SUBRs = New ADODB.Recordset
            SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
        
            ' 자료가 없으면 종료를 한다.
            Do While Not SUBRs.EOF
                sData(0) = 대리점정보.StoreCode & ""
                sData(1) = Trim(SUBRs!접수일자 & "")
                sData(2) = Trim(SUBRs!쿠폰번호 & "")
                sData(3) = IIf(IsNumeric(SUBRs!쿠폰단가 & ""), SUBRs!쿠폰단가, "0")
                sData(4) = IIf(IsNumeric(SUBRs!쿠폰금액 & ""), SUBRs!쿠폰금액, "0")
                sData(5) = Trim(SUBRs!고객번호 & "")
                sData(6) = Trim(SUBRs!고객이름 & "")
                sData(7) = IIf(IsNumeric(SUBRs!접수금액 & ""), SUBRs!접수금액, "0")
                sData(8) = Trim(SUBRs!택번호 & "")
                
                '------------------------------------------------------------------------
                '
                '------------------------------------------------------------------------
                Query = "SELECT * "
                Query = Query & " FROM CouponUseDataALL "
                Query = Query & " WHERE StoreCode = '" & 대리점정보.StoreCode & "' "
                Query = Query & "   AND SaleDate  = '" & Format(iTempDay, "yyyyMMdd") & "' "
                Query = Query & "   AND Number    = '" & Trim(SUBRs!쿠폰번호 & "") & "' "
                Set Rs = New ADODB.Recordset
                Rs.Open Query, MyHost, adOpenStatic, adLockOptimistic
                
                'If ADORset.State = adStateOpen Then ADORset.Close
                '
                'ADORset.CursorLocation = adUseClient
                'ADORset.Open Query, MyHost, adOpenStatic, adLockBatchOptimistic, adCmdText
                
                If Rs.EOF = True Then
                    Query = "INSERT INTO  CouponUseDataALL (SaleDate, StoreCode, MASTER_CD, "
                    Query = Query & " Number , Cost, "
                    Query = Query & " Money, CustNum, "
                    Query = Query & " CustName, SaleMoney, StoreTag, "
                    Query = Query & " SendYN, sEndDate  )"
                    Query = Query & " VALUES ('" & sData(1) & "', '" & 대리점정보.StoreCode & "', '" & Scode(0) & "', "
                    Query = Query & " '" & sData(2) & "', " & sData(3) & ", " 'Number, Cost
                    Query = Query & " " & sData(4) & ", '" & sData(5) & "', "       'Money, CustNum
                    Query = Query & " '" & sData(6) & "', " & sData(7) & ", '" & sData(8) & "', "  'CustName, SaleMoney, StoreTag"
                    Query = Query & " 'Y', '" & Format(Date, "yyyyMMdd") & "' "    'SendYN , sEndDate
                    Query = Query & " )"
                    MyHost.Execute Query
                Else
                    If Rs.Fields("SendYN") = "N" Then
                        Query = "UPDATE CouponUseDataALL "
                        Query = Query & " SET Number = '" & sData(2) & "', "
                        Query = Query & " Cost =  " & sData(3) & ", "
                        Query = Query & " Money =  " & sData(4) & ", "
                        Query = Query & " CustNum =  '" & sData(5) & "', "
                        Query = Query & " CustName = '" & sData(6) & "', "
                        Query = Query & " SaleMoney = " & sData(7) & ", "
                        Query = Query & " StoreTag = '" & sData(8) & "', "
                        Query = Query & " SendYN =  'Y', "
                        Query = Query & " sEndDate =  '" & Format(Date, "yyyyMMdd") & "'  "
                        Query = Query & " WHERE StoreCode = '" & 대리점정보.StoreCode & "' "
                        Query = Query & "   AND MASTER_CD = '" & Scode(0) & "' "
                        Query = Query & "   AND SaleDate = '" & Format(iTempDay, "yyyyMMdd") & "' "
                        MyHost.Execute Query
                    End If
                End If
                Rs.Close
                Set Rs = Nothing
                
                SUBRs.MoveNext
            Loop
                
            SUBRs.Close
            Set SUBRs = Nothing
        Next iDay
    End If
    
    Set MyHost = Nothing
    
    SendTable_쿠폰정보 = True

    On Error GoTo 0
    
    Exit Function

SendTable_쿠폰정보_Error:
    SendTable_쿠폰정보 = False
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SendTable_쿠폰정보 of Module Global"
End Function

Public Function GetSelectSpread(MySS As Object, nCol As Long) As Long
    Dim lRow    As Long
    Dim nSelCnt As Long
    
    nSelCnt = 0
    
    For lRow = 1 To MySS.MaxRows
        MySS.Col = nCol
        MySS.Row = lRow
        
        If MySS.Value = 1 Then nSelCnt = nSelCnt + 1
    Next lRow
    
    GetSelectSpread = nSelCnt
End Function


'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : Sort_Select
' 작  성  자  : IT21
' 작  성  일  : 2000.06.26
' 파 라 미 터 : MySpread   - Spread Control
'               nSortOrder - Sort 방향
'               lRow       - 시작열
' 비      고  : 대리점명을 리턴한다.
'-----------------------------------------------------------------------------------------------------------------------------------------
'Public Sub Sort_Select(MySpread As Object, nSortOrder As Integer, iCol As Long)
'
'    On Error GoTo Err_Rtn
'
'    MySpread.Row = 1
'    MySpread.Col = 1
'    MySpread.Row2 = MySpread.MaxRows
'    MySpread.Col2 = MySpread.MaxCols
'    MySpread.SortBy = 0
'    MySpread.SortKey(1) = iCol
'
'    If iCol = MySpread.MaxCols Then
'        MySpread.SortKey(2) = 1
'    Else
'        MySpread.SortKey(2) = iCol + 1
'    End If
'
'    If iCol + 1 = MySpread.MaxCols Then
'        MySpread.SortKey(3) = 1
'    Else
'        MySpread.SortKey(3) = iCol + 2
'    End If
'
'    MySpread.SortKeyOrder(1) = nSortOrder
'    MySpread.SortKeyOrder(2) = nSortOrder
'    MySpread.SortKeyOrder(3) = nSortOrder
'    MySpread.Action = 25
'
'Err_Rtn:
'
'End Sub
 
Public Sub Sort_Select(MySpread As Object, nSortOrder As Integer, lRow As Long)

On Error GoTo Err_Rtn
    MySpread.Row = lRow
    MySpread.Col = 1
    MySpread.Row2 = MySpread.MaxRows
    MySpread.Col2 = MySpread.MaxCols
    MySpread.SortBy = 0
    MySpread.SortKey(1) = MySpread.ActiveCol
    
    If MySpread.ActiveCol = MySpread.MaxCols Then
        MySpread.SortKey(2) = 1
    Else
        MySpread.SortKey(2) = MySpread.ActiveCol + 1
    End If
    
    If MySpread.ActiveCol + 1 = MySpread.MaxCols Then
        MySpread.SortKey(3) = 1
    Else
        MySpread.SortKey(3) = MySpread.ActiveCol + 2
    End If
    
    MySpread.SortKeyOrder(1) = nSortOrder
    MySpread.SortKeyOrder(2) = nSortOrder
    MySpread.SortKeyOrder(3) = nSortOrder
    MySpread.Action = 25
    
    Exit Sub
    
Err_Rtn:
    MsgBox "정렬 할 항목을 선택후 작업하세요!", vbOKOnly
End Sub


Public Function GetCouponMoney(ByVal strCouponNumbar As String) As Double
    Dim dblMoney As Double
    
    Select Case Left(strCouponNumbar, 2)
        Case "02"
            ' 모피행사 2009-12-31일가지
            dblMoney = 50000
        
        Case "01"
            ' LG 타운젠트 행사
            dblMoney = 5000
            
        Case "00"
            ' 크랜즈겔러리
            dblMoney = 33000
            
        Case Else
            dblMoney = 5000
    End Select
    
    GetCouponMoney = dblMoney
    
End Function

Public Function GetCouponCost(ByVal strCouponNumbar As String) As Double
    Dim dblMoney As Double
    
    Select Case Left(strCouponNumbar, 2)
    
        Case "02"
            ' 모피행사 2009-12-31일가지
            dblMoney = 0
        
        Case "01"
            ' LG 타운젠트 행사
            dblMoney = 3000
            
        Case "00"
            ' 크랜즈겔러리
            dblMoney = 0
        
        Case Else
            dblMoney = 5000
    End Select
    
    GetCouponCost = dblMoney
End Function


' 현재 전달된 일자의 쿠폰 금액을 가저온다.
Public Function GetCouponSaleTotalMoney(ByVal strDate As String, ByRef CouponCnt As Integer) As Double
    Dim dblMoney As Double
    
    dblMoney = 0
    
    Query = "SELECT 쿠폰번호 FROM 쿠폰자료 "
    Query = Query & " WHERE 접수일자 = '" & strDate & "' "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If Not SUBRs.EOF Then
        SUBRs.MoveLast
        
        CouponCnt = CStr(SUBRs.RecordCount)
        
        SUBRs.MoveFirst
        
        Do Until SUBRs.EOF
            dblMoney = dblMoney + GetCouponMoney(CStr(SUBRs.Fields("쿠폰번호")) & "")
            
            SUBRs.MoveNext
        Loop
    End If
    SUBRs.Close
    Set SUBRs = Nothing
    
    GetCouponSaleTotalMoney = dblMoney
End Function


' 현재 전달된 일자의 쿠폰 금액을 가저온다.
Public Function GetCouponSaleTotalMoney2(ByVal strDate As String) As Double
    Dim sTemp     As String
    Dim Index     As Integer
    Dim dblMoney  As Double
    Dim dCost     As Double
    Dim sSale     As Double
    Dim sNum(99)  As Integer

    dblMoney = 0
    
    Erase sNum
    
    Query = "SELECT 쿠폰번호  FROM 쿠폰자료 "
    Query = Query & " WHERE 접수일자 = '" & strDate & "' "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    Do Until SUBRs.EOF
        sTemp = Left(SUBRs.Fields("쿠폰번호") & "", 2)
        
        sNum(Val(sTemp)) = sNum(Val(sTemp)) + 1
        
        SUBRs.MoveNext
    Loop
    SUBRs.Close
    Set SUBRs = Nothing
    
    For Index = 0 To 99
        If sNum(Index) > 0 Then
            dCost = GetCouponCost(Format(Index, "00"))
            sSale = GetCouponMoney(Format(Index, "00"))
            
            dblMoney = dblMoney + (sNum(Index) * (sSale - dCost))
        End If
    Next Index
    
    GetCouponSaleTotalMoney2 = dblMoney
End Function

Public Function CheckCouponNumber(sCouponNum As String) As Integer
    ' 정상 0
    CheckCouponNumber = -1

    ' 입력 형태 검사
    If Len(sCouponNum) <> M_COUPON_LANGTH Then
        CheckCouponNumber = -1
        
        Exit Function
        
    ' 크렌즈겔러리 모피 행사 50,000 2009-12-31일까지
    ElseIf Left(sCouponNum, 2) = "02" And 대리점정보.MasterCode <> M_COUPON_KLENZ_CODE Then
        CheckCouponNumber = -1
        
        Exit Function
        
    ElseIf Left(sCouponNum, 2) = "00" And 대리점정보.MasterCode <> M_COUPON_KLENZ_CODE Then
        CheckCouponNumber = -1
        
        Exit Function
        
    ElseIf Left(sCouponNum, 2) = "01" And 대리점정보.MasterCode = M_COUPON_KLENZ_CODE Then
        CheckCouponNumber = -1
        
        Exit Function
    
    ' 유효기간 검사 오류
    ElseIf Left(sCouponNum, 2) = "01" And Format(Date, "yyyyMMdd") > "20090831" Then
        CheckCouponNumber = -2
        
        Exit Function
    End If
    
    CheckCouponNumber = 0
End Function

Public Function ReadSendTextMessage(cboSendText As Object) As Boolean
    cboSendText.Clear
    
    If 대리점정보.SMS_EMART = "Y" Then
        cboSendText.AddItem "고객님 세탁물이 도착했습니다. 인수 부탁 드립니다. 즐겁고 행복한 하루 되세요."
        cboSendText.AddItem "고객님 세탁물을 보관 중입니다. 인수 부탁 드립니다. 즐겁고 행복한 하루 되세요."
        cboSendText.AddItem "고객님 세탁물을 보관 중입니다. 보관 중 먼지가 쌓일 수 있으니 빠른 인수 바랍니다."
        cboSendText.AddItem "고객님 크린에이드를 이용해 주셔서 감사 드립니다. 즐겁고 행복한 하루 되세요."
        cboSendText.AddItem "고객님 금일 크린에이드 이용중에 불편 드렸던 점 진심으로 사과 드립니다."
        cboSendText.AddItem "고객님 접수하신 세탁물이 다소 지연되고 있습니다. 도착하는 대로 연락 드리겠습니다"
        cboSendText.AddItem "고객님 접수하신 세탁물이 반품되어 확인이 필요합니다. 내방 부탁 드립니다."
        cboSendText.AddItem "고객님 죄송합니다.^^ 고객님께 메시지전송이 오류로 발송되었습니다. 감사합니다."
        
'        cboSendText.AddItem "고객님 세탁물이 도착했습니다. 조속한 시일 내에 인수 부탁 드립니다."
'        cboSendText.AddItem "고객님 크린에이드를 이용해 주신 점 감사 드립니다.즐겁고 행복한 하루 되세요."
'        cboSendText.AddItem "고객님 금일 크린에이드 이용 중에 불편 드렸던 점 진심으로 사과 드립니다."
'        cboSendText.AddItem "고객님 접수하신 세탁물이 다소 지연되고 있습니다.도착하는 대로 연락 드리겠습니다."
'        cboSendText.AddItem "고객님 접수하신 세탁물이 반품되어 확인이 필요합니다. 내방 부탁 드립니다."
    Else
        Query = " SELECT * "
        Query = Query & "  FROM 문자발송문 "
        Query = Query & " ORDER BY 순번 "
        Set SUBRs = New ADODB.Recordset
        SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
        If SUBRs.EOF Then
            SUBRs.Close
            Set SUBRs = Nothing
            
            Query = "INSERT INTO  문자발송문 VALUES('01', '" & "전할 메시지를 입력하여 주십시요" & "') "
            ADOCon.Execute Query
        
            '--------------------------------------------
            Query = " SELECT * "
            Query = Query & "  FROM 문자발송문 "
            Query = Query & " ORDER BY 순번 "
            Set SUBRs = New ADODB.Recordset
            SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
        End If
        
        While Not SUBRs.EOF
            cboSendText.AddItem Trim(SUBRs.Fields("내용") & "")
            cboSendText.ItemData(cboSendText.ListCount - 1) = Val(SUBRs.Fields("순번") & "")
            
            SUBRs.MoveNext
        Wend
        
        SUBRs.Close
        Set SUBRs = Nothing
    End If
        
    ' 마지막 문구 선택
    If cboSendText.ListCount > 0 Then
        cboSendText.ListIndex = 0
    End If
End Function

'====================================================================================================
' Procedure : SendTable_참조코드
' DateTime  : 2008-05-03 01:42
' Author    : pds2004
' Last Edit :
' Return Val:
'----------------------------------------------------------------------------------------------------
' 설     명 : 전달된 일자부터 현재일자까지의 고객정보를 본사 SQL 서버에 저장한다.
'====================================================================================================
Public Function SendTable_사고품(ByVal sStartDate As String, ByRef MyHost As Object, ByRef objPrBar As Object) As Long
    On Error GoTo SendTable_사고품_Error
    
    If IsDate(sStartDate) = False Then
        MsgBox "전달된 일자가 올바르지 않습니다.  [" & sStartDate & "]", vbCritical, "경고"
        Exit Function
    End If
     
     SendTable_사고품 = 0
    
    ' 미전송 자료를 구한다.
    Query = "SELECT * FROM 사고품 WHERE 전송구분 <> 'Y' OR  접수일 >= '" & Format(sStartDate, "yyyyMMdd") & "' "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    ' 미전송 자료가 없으면 종료를 한다.
    If Not SUBRs.EOF Then
        SUBRs.MoveLast
        
        objPrBar.MAX = SUBRs.RecordCount
        
        SUBRs.MoveFirst
    End If
    
    Do While Not SUBRs.EOF
        Query = "EXEC PRO_P_06004_01 "
        Query = Query & "'" & Replace(SUBRs.Fields("접수일"), vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("일련번호"), vbNullChar, "") & "', "
        Query = Query & "'" & 대리점정보.StoreCode & "', "
        Query = Query & "'" & 대리점정보.MasterCode & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("성명"), vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("고객전화"), vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("주소"), Chr(2), "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("휴대폰"), vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("품명"), vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("상표"), vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("구입일자"), vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("색상"), vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("구입처"), vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("최초택번호"), vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("최종택번호"), vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("구입형태"), vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("최초입고일"), vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("최종입고일"), vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("구입가격"), vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("사고접수일"), vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("사고종류"), vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("사고내용"), vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("사고의견"), vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("보상금액"), vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("합의금액"), vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("처리유무"), vbNullChar, "") & "' "
        MyHost.Execute Query
        
        Query = "UPDATE 사고품 SET "
        Query = Query & " 전송구분 = 'Y'  "
        Query = Query & " WHERE 일련번호 = " & SUBRs.Fields("일련번호")
        ADOCon.Execute Query
        
        If objPrBar.Value < objPrBar.MAX Then
            objPrBar.Value = objPrBar.Value + 1
        End If
        
        SUBRs.MoveNext
    Loop
    
    Call SendTableDateSave("사고품", MyHost)
    
    SUBRs.Close
    Set SUBRs = Nothing
    
    On Error GoTo 0
    
    Exit Function

SendTable_사고품_Error:

    SendTable_사고품 = 0
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SendTable_사고품 of Module Global"
    
    Resume Next
End Function



'====================================================================================================
' Procedure : SendTable_입출고
' DateTime  : 2008-05-03 01:42
' Author    : pds2004
' Last Edit :
' Return Val:
'----------------------------------------------------------------------------------------------------
' 설     명 : 전달된 일자부터 현재일자까지의 입출고 내요을 본사 SQL 서버에 저장한다.
'====================================================================================================
Public Function SendTable_입출고(ByVal sStartDate As String, ByRef MyHost As Object, ByRef objPrBar As Object) As Long
    On Error GoTo SendTable_입출고_Error
    
    If IsDate(sStartDate) = False Then
        MsgBox "전달된 일자가 올바르지 않습니다.  [" & sStartDate & "]", vbCritical, "경고"
        Exit Function
    End If
     
     SendTable_입출고 = 0
    
    '-------------------------------------------------------------------------------------------------------------
    ' 미전송 자료를 구한다. ' 최초 너무 많은 자료 전송을 막기 위하여 2009-08-01일을 기준으로 정한다.
    '-------------------------------------------------------------------------------------------------------------
    'Query = "SELECT * FROM 입출고 "
    'Query = Query & " WHERE (iif( 본사전송구분  is null , 'N',본사전송구분) <> 'Y'  and 입고일 >= '2009-08-01')  "
    'Query = Query & "   AND  입고일 >= '" & Format(sStartDate, "yyyyMMdd") & "' "
    'Query = Query & "   AND ( 코드 LIKE 'a%' OR  세탁비환불일자 <> '') "
    'Query = Query & "    OR  left(세탁비환불일자,8) = '" & Format(sStartDate, "yyyyMMdd") & "' "
    
    Query = "SELECT * FROM 입출고"
    Query = Query & " WHERE (CASE 본사전송구분 WHEN Null THEN 'N' ELSE 본사전송구분 END) <> 'Y'"
    Query = Query & "   AND  입고일 >= '" & Format(sStartDate, "yyyyMMdd") & "' "
    Query = Query & "   AND ( 코드 LIKE 'a%' OR  세탁비환불일자 <> '') "
    Query = Query & "    OR SUBSTRING(세탁비환불일자,1,8) = '" & Format(sStartDate, "yyyyMMdd") & "' "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    ' 미전송 자료가 없으면 종료를 한다.
    If Not SUBRs.EOF Then
        SUBRs.MoveLast
        
        objPrBar.MAX = SUBRs.RecordCount
        
        SUBRs.MoveFirst
    End If
    
    Do While Not SUBRs.EOF
        Query = "EXEC PRO_A_00005 "
        Query = Query & "'" & 대리점정보.StoreCode & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("입고일") & "", vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("번호") & "", vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("고객번호") & "", vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("코드") & "", vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("품명") & "", vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("색상") & "", vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("내용") & "", vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("금액") & "", vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("상표") & "", vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("본출") & "", vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("상태") & "", vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("확인") & "", vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("출고일") & "", vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("판매취소") & "", vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("입고예정일") & "", vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("환불일자") & "", vbNullChar, "") & "', "
        Query = Query & "" & Replace(Val(SUBRs.Fields("수선금액") & ""), vbNullChar, "") & ", "
        Query = Query & "'" & Replace(SUBRs.Fields("본출일자") & "", vbNullChar, "") & "', "
        Query = Query & "'" & Replace(SUBRs.Fields("본출입고구분") & "", vbNullChar, "") & "', "
        Query = Query & " " & Replace(Val(SUBRs.Fields("외주운동화마진") & ""), vbNullChar, "") & ", "
        Query = Query & "'" & Replace(SUBRs.Fields("세탁비환불일자") & "", vbNullChar, "") & "'  "
        MyHost.Execute Query
        
        '-------------------------------------------------------------------------
        '
        '-------------------------------------------------------------------------
        Query = "UPDATE 입출고 SET 본사전송구분 = 'Y'  "
        Query = Query & " WHERE 입고일 = '" & SUBRs.Fields("입고일") & "' "
        Query = Query & "   AND 번호   = '" & SUBRs.Fields("번호") & "'"
        ADOCon.Execute Query
        
        If objPrBar.Value < objPrBar.MAX Then
            objPrBar.Value = objPrBar.Value + 1
        End If
        
        SUBRs.MoveNext
    Loop
    
    Call SendTableDateSave("입출고", MyHost)
    
    SUBRs.Close
    Set SUBRs = Nothing
    
    On Error GoTo 0
    
    Exit Function

SendTable_입출고_Error:
    SendTable_입출고 = 0
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SendTable_입출고 of Module Global"
    
    Resume Next
End Function


'====================================================================================================
' Procedure : SendTable_입출고
' DateTime  : 2008-05-03 01:42
' Author    : pds2004
' Last Edit :
' Return Val:
'----------------------------------------------------------------------------------------------------
' 설     명 : 전달된 일자부터 현재일자까지의 입출고 내요을 본사 SQL 서버에 저장한다.
'====================================================================================================
Public Function SendTable_메시지(ByVal sStartDate As String, ByRef MyHost As Object, ByRef objPrBar As Object) As Long
    On Error GoTo SendTable_메시지_Error
    
    If IsDate(sStartDate) = False Then
        MsgBox "전달된 일자가 올바르지 않습니다.  [" & sStartDate & "]", vbCritical, "경고"
        Exit Function
    End If
     
    SendTable_메시지 = 0
    
    ' 미전송 자료를 구한다.
    ' 최초 너무 많은 자료 전송을 막기 위하여 2009-08-01일을 기준으로 정한다.
    
    '------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------
    Query = "SELECT * FROM 메일 "
    Query = Query & " WHERE IIF(ISNULL(전송구분) ,'N',전송구분) <> 'Y' "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    ' 미전송 자료가 없으면 종료를 한다.
    If Not SUBRs.EOF Then
        SUBRs.MoveLast
        
        objPrBar.MAX = SUBRs.RecordCount
        
        SUBRs.MoveFirst
    End If
    
    Do While Not SUBRs.EOF
        ' MailType, MailDate, MailNo, AgencyCode, MailFrom, MailTo, MailDesc, ReadChk, ReadDate, SendChk
        
        '----------------------------------------------------------------------------
        '
        '----------------------------------------------------------------------------
        Query = "UPDATE Mail_ALL SET "
        Query = Query & " ReadChk = '" & SUBRs.Fields("수신여부") & "', "
        Query = Query & " ReadDate = '" & SUBRs.Fields("수신일자") & "' "
        Query = Query & " WHERE MailType   = '2' "
        Query = Query & "   AND MailDate   = '" & SUBRs.Fields("메일일자") & "' "
        Query = Query & "   AND MailNo     = " & Val(SUBRs.Fields("메일번호")) & " "
        Query = Query & "   AND AgencyCode = '" & 대리점정보.StoreCode & "' "
        MyHost.Execute Query
        
        '----------------------------------------------------------------------------
        '
        '----------------------------------------------------------------------------
        Query = "UPDATE 메일 SET 전송구분 = 'Y' "
        Query = Query & " WHERE 메일일자 = '" & SUBRs.Fields("메일일자") & "" & "' "
        Query = Query & "   AND 메일번호 =  " & SUBRs.Fields("메일번호") & "" & " "
        ADOCon.Execute Query
    
        If objPrBar.Value < objPrBar.MAX Then
            objPrBar.Value = objPrBar.Value + 1
        End If
        
        SUBRs.MoveNext
    Loop
    
    SUBRs.Close
    Set SUBRs = Nothing
    
    On Error GoTo 0
    
    Exit Function

SendTable_메시지_Error:
    SendTable_메시지 = 0
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SendTable_메시지 of Module Global"
End Function


'====================================================================================================
' Procedure : SendTable_세트상품정보
' DateTime  : 2008-05-03 01:42
' Author    : pds2004
' Last Edit :
' Return Val:
'----------------------------------------------------------------------------------------------------
' 설     명 : 전달된 일자부터 현재일자까지의 입출고 내요을 본사 SQL 서버에 저장한다.
'====================================================================================================
Public Function SendTable_세트상품정보(ByVal sStartDate As String, ByRef MyHost As Object, ByRef objPrBar As Object) As Long
    Dim sValue(22) As String
    
    On Error GoTo SendTable_Error
    
    If IsDate(sStartDate) = False Then
        MsgBox "전달된 일자가 올바르지 않습니다.  [" & sStartDate & "]", vbCritical, "경고"
        Exit Function
    End If
     
     SendTable_세트상품정보 = 0
    
    Query = "SELECT * FROM 세트상품정보 "
    Query = Query & " WHERE TRIM(IIF(ISNULL(SendDate) ,'',SendDate)) = '' "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    ' 미전송 자료가 없으면 종료를 한다.
    If Not SUBRs.EOF Then
        SUBRs.MoveLast
        objPrBar.MAX = SUBRs.RecordCount
        SUBRs.MoveFirst
    End If
    
    Do While Not SUBRs.EOF
        sValue(0) = 대리점정보.StoreCode
        sValue(1) = SUBRs.Fields("접수일자") & ""
        sValue(2) = SUBRs.Fields("세트Key") & ""
        sValue(3) = SUBRs.Fields("고객코드") & ""
        sValue(4) = SUBRs.Fields("고객명") & ""
        sValue(5) = SUBRs.Fields("고객전화번호") & ""
        sValue(6) = SUBRs.Fields("휴대폰번호") & ""
        
        sValue(7) = Replace(SUBRs.Fields("정상금액") & "", ",", "")
        sValue(8) = Replace(SUBRs.Fields("세트금액") & "", ",", "")
        sValue(9) = Replace(SUBRs.Fields("세트할인금액") & "", ",", "")
        sValue(10) = Replace(SUBRs.Fields("에누리할인금액") & "", ",", "")
        sValue(11) = Replace(SUBRs.Fields("적용합계금액") & "", ",", "")
        
        sValue(12) = Val(Replace(SUBRs.Fields("세트2") & "", ",", ""))
        sValue(13) = Val(Replace(SUBRs.Fields("세트3") & "", ",", ""))
        sValue(14) = Val(Replace(SUBRs.Fields("세트4") & "", ",", ""))
        sValue(15) = Val(Replace(SUBRs.Fields("세트5") & "", ",", ""))
        sValue(16) = Val(Replace(SUBRs.Fields("세트6") & "", ",", ""))
        sValue(17) = Val(Replace(SUBRs.Fields("세트7") & "", ",", ""))
        sValue(18) = Val(Replace(SUBRs.Fields("세트8") & "", ",", ""))
        sValue(19) = Val(Replace(SUBRs.Fields("세트9") & "", ",", ""))
        sValue(20) = Val(Replace(SUBRs.Fields("세트10") & "", ",", ""))
        sValue(21) = Val(Replace(SUBRs.Fields("무료세탁권수") & "", ",", ""))
        sValue(22) = 대리점정보.MasterCode
        
        Query = "EXEC PRO_GROUPGOODS_SEND "
        Query = Query & "'" & sValue(0) & "', "
        Query = Query & "'" & sValue(1) & "', "
        Query = Query & "'" & sValue(2) & "', "
        Query = Query & "'" & sValue(3) & "', "
        Query = Query & "'" & sValue(4) & "', "
        Query = Query & "'" & sValue(5) & "', "
        Query = Query & "'" & sValue(6) & "', "
        Query = Query & "" & sValue(7) & ", "
        Query = Query & "" & sValue(8) & ", "
        Query = Query & "" & sValue(9) & ", "
        Query = Query & "" & sValue(10) & ", "
        Query = Query & "" & sValue(11) & ", "
        Query = Query & "" & sValue(12) & ", "
        Query = Query & "" & sValue(13) & ", "
        Query = Query & "" & sValue(14) & ", "
        Query = Query & "" & sValue(15) & ", "
        Query = Query & "" & sValue(16) & ", "
        Query = Query & "" & sValue(17) & ", "
        Query = Query & "" & sValue(18) & ", "
        Query = Query & "" & sValue(19) & ", "
        Query = Query & "" & sValue(20) & ", "
        Query = Query & "" & sValue(21) & ", "
        Query = Query & "'" & sValue(22) & "' "
        MyHost.Execute Query
        
        '----------------------------------------------------------------------------
        '
        '----------------------------------------------------------------------------
        Query = "UPDATE 세트상품정보 SET "
        Query = Query & " SendDate       = '" & Format(Date, "yyyyMMdd") & "'"
        Query = Query & " WHERE 접수일자 = '" & SUBRs.Fields("접수일자") & "" & "' "
        Query = Query & "   AND 세트Key  = '" & SUBRs.Fields("세트Key") & "" & "' "
        ADOCon.Execute Query
    
        If objPrBar.Value < objPrBar.MAX Then
            objPrBar.Value = objPrBar.Value + 1
        End If
        
        SUBRs.MoveNext
    Loop
    
    SUBRs.Close
    Set SUBRs = Nothing
    
    On Error GoTo 0
    
    Exit Function

SendTable_Error:

    SendTable_세트상품정보 = 0
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SendTable_세트상품정보 of Module Global"
End Function


'====================================================================================================
' Procedure : SendTable_입출고
' DateTime  : 2008-05-03 01:42
' Author    : pds2004
' Last Edit :
' Return Val:
'----------------------------------------------------------------------------------------------------
' 설     명 : 전달된 일자부터 현재일자까지의 입출고 내요을 본사 SQL 서버에 저장한다.
'====================================================================================================
Public Function SendTable_세트응모번호(ByVal sStartDate As String, ByRef MyHost As Object, ByRef objPrBar As Object) As Long
    Dim sValue(7) As String
    
    On Error GoTo SendTable_Error
    
    If IsDate(sStartDate) = False Then
        MsgBox "전달된 일자가 올바르지 않습니다.  [" & sStartDate & "]", vbCritical, "경고"
        Exit Function
    End If
     
    SendTable_세트응모번호 = 0
    
    '--------------------------------------------------------------------------------
    '
    '--------------------------------------------------------------------------------
    Query = "SELECT * FROM 세트응모번호 "
    Query = Query & " WHERE TRIM(IIF(ISNULL(SendDate) ,'',SendDate)) = '' "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
        
    ' 미전송 자료가 없으면 종료를 한다.
    If Not SUBRs.EOF Then
        SUBRs.MoveLast
        
        objPrBar.MAX = rsTB.RecordCount
        
        SUBRs.MoveFirst
    End If
    
    Do While Not SUBRs.EOF
        sValue(0) = 대리점정보.StoreCode
        sValue(1) = SUBRs.Fields("응모번호") & ""
        sValue(2) = SUBRs.Fields("세트Key") & ""
        sValue(3) = SUBRs.Fields("일자") & ""
        sValue(4) = SUBRs.Fields("고객코드") & ""
        sValue(5) = SUBRs.Fields("고객명") & ""
        sValue(6) = SUBRs.Fields("고객전화번호") & ""
        sValue(7) = SUBRs.Fields("휴대폰번호") & ""
        
        Query = "EXEC PRO_GROUPGOODS_NUMBER_SEND "
        Query = Query & "'" & sValue(0) & "', "
        Query = Query & "'" & sValue(1) & "', "
        Query = Query & "'" & sValue(2) & "', "
        Query = Query & "'" & sValue(3) & "', "
        Query = Query & "'" & sValue(4) & "', "
        Query = Query & "'" & sValue(5) & "', "
        Query = Query & "'" & sValue(6) & "', "
        Query = Query & "'" & sValue(7) & "'  "
        MyHost.Execute Query

        Query = "UPDATE 세트응모번호 SET SendDate = '" & Format(Date, "yyyyMMdd") & "'"
        Query = Query & " WHERE 응모번호 = '" & SUBRs.Fields("응모번호") & "" & "' "
        ADOCon.Execute Query
    
        If objPrBar.Value < objPrBar.MAX Then
            objPrBar.Value = objPrBar.Value + 1
        End If
        
        SUBRs.MoveNext
    Loop
    SUBRs.Close
    Set SUBRs = Nothing
    
    On Error GoTo 0
    
    Exit Function

SendTable_Error:
    SendTable_세트응모번호 = 0
    
    Call Error_Msg("SendTable_세트응모번호", Err.Source, Err.Number, Err.Description)
    
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SendTable_세트응모번호 of Module Global"
End Function

'---------------------------------------------------------------------------
' 함수명 : Error_Msg
' 기  능 :
' 설  명 :
'---------------------------------------------------------------------------
Public Sub Error_Msg(strEvent As String, strSource As String, strNumber As String, strDescription As String)
              Err_Msg = "발생위치 : " & strEvent & vbNewLine & vbNewLine
    Err_Msg = Err_Msg & "오류소스 : " & strSource & vbNewLine & vbNewLine
    Err_Msg = Err_Msg & "오류번호 : " & strNumber & vbNewLine & vbNewLine
    Err_Msg = Err_Msg & "오류내용 : " & strDescription
    
    MsgBox Err_Msg, vbCritical, "오류"
    Screen.MousePointer = 0
End Sub

