Attribute VB_Name = "basLKPOS"
Option Explicit
'Method return value
Public Const LK_SUCCESS As Integer = 0
Public Const LK_CREATE_ERROR As Integer = 1
Public Const LK_NOT_OPENED As Integer = 2
Public Const LK_FAIL As Integer = -1

' Printer Status flag
Public Const LK_STS_NORMAL As Integer = 0
Public Const LK_STS_COVEROPEN As Integer = 1
Public Const LK_STS_PAPERNEAREMPTY As Integer = 2
Public Const LK_STS_PAPEREMPTY As Integer = 4
Public Const LK_STS_POWEROFF As Integer = 8

' Alignment Code
Public Const LK_ALIGNMENT_LEFT As Integer = 0
Public Const LK_ALIGNMENT_CENTER As Integer = 1
Public Const LK_ALIGNMENT_RIGHT As Integer = 2

' Bitmap Size
Public Const LK_BITMAP_NORMAL As Integer = 0
Public Const LK_BITMAP_WIDTH_DOUBLE As Integer = 1
Public Const LK_BITMAP_HEIGHT_DOUBLE As Integer = 2
Public Const LK_BITMAP_WIDTH_HEIGHT_DOUBLE As Integer = 3

' Bitmap Image Mode
Public Const LK_BITMAP_NO_DITHER As Integer = 0
Public Const LK_BITMAP_ERROR_DIFFUSION As Integer = 1
Public Const LK_BITMAP_ORDERED_DITHER As Integer = 2

'Text Attribute
'Font Attribute default value : not Bold, FontA, not Underline, not reverse
Public Const LK_FNT_DEFAULT As Integer = 0
Public Const LK_FNT_FONTB As Integer = 1
Public Const LK_FNT_BOLD As Integer = 8
Public Const LK_FNT_UNDERLINE As Integer = 128

' Text Size Attribute
Public Const LK_TXT_1WIDTH As Integer = 0
Public Const LK_TXT_2WIDTH As Integer = 16
Public Const LK_TXT_3WIDTH As Integer = 32
Public Const LK_TXT_4WIDTH As Integer = 48
Public Const LK_TXT_5WIDTH As Integer = 64
Public Const LK_TXT_6WIDTH As Integer = 80
Public Const LK_TXT_7WIDTH As Integer = 96
Public Const LK_TXT_8WIDTH As Integer = 112

Public Const LK_TXT_1HEIGHT As Integer = 0
Public Const LK_TXT_2HEIGHT As Integer = 1
Public Const LK_TXT_3HEIGHT As Integer = 2
Public Const LK_TXT_4HEIGHT As Integer = 3
Public Const LK_TXT_5HEIGHT As Integer = 4
Public Const LK_TXT_6HEIGHT As Integer = 5
Public Const LK_TXT_7HEIGHT As Integer = 6
Public Const LK_TXT_8HEIGHT As Integer = 7

' Barcode
Public Const LK_BCS_PDF417 As Integer = 200
Public Const LK_BCS_MAXICODE As Integer = 201
Public Const LK_BCS_QRCODE As Integer = 202
Public Const LK_BCS_DATAMATRIX As Integer = 203

Public Const LK_BCS_UPCA As Integer = 101
Public Const LK_BCS_UPCE As Integer = 102
Public Const LK_BCS_EAN8 As Integer = 103
Public Const LK_BCS_EAN13 As Integer = 104
Public Const LK_BCS_JAN8 As Integer = 105
Public Const LK_BCS_JAN13 As Integer = 106
Public Const LK_BCS_ITF As Integer = 107
Public Const LK_BCS_Codabar As Integer = 108
Public Const LK_BCS_Code39 As Integer = 109
Public Const LK_BCS_Code93 As Integer = 110
Public Const LK_BCS_Code128 As Integer = 111
Public Const LK_BCS_3OF5 As Integer = 112

' Barcode text position
Public Const LK_HRI_TEXT_NONE As Integer = 0
Public Const LK_HRI_TEXT_ABOVE As Integer = 1
Public Const LK_HRI_TEXT_BELOW As Integer = 2


' Cash Drawer Status
Public Const LK_CD_STS_CLOSED As Integer = 0
Public Const LK_CD_STS_OPENED As Integer = 1

' Cash Drawer Pin Connector
Public Const LK_CD_PIN_TWO As Integer = 2
Public Const LK_CD_PIN_FIVE As Integer = 5

 
Public Declare Function OpenPort Lib "LKPosPrinter.dll" (ByVal PortName As String, ByVal BaudRate As Long) As Long
Public Declare Function ClosePort Lib "LKPosPrinter.dll" () As Long
Public Declare Function OpenTcpip Lib "LKPosPrinter.dll" (ByVal IP As String, ByVal PortNumber As Long) As Long
Public Declare Function CloseTcpip Lib "LKPosPrinter.dll" () As Long
Public Declare Function PrintBitmap Lib "LKPosPrinter.dll" (ByVal BitmapFile As String, ByVal Alignment As Long, ByVal options As Long, ByVal Brightness As Long, ByVal ImageMode As Long) As Long
Public Declare Function PrintBitmapTcpip Lib "LKPosPrinter.dll" (ByVal BitmapFile As String, ByVal Alignment As Long, ByVal options As Long, ByVal Brightness As Long, ByVal ImageMode As Long) As Long
Public Declare Function PrintString Lib "LKPosPrinter.dll" (ByVal Data As String) As Long
Public Declare Function PrintStringTcpip Lib "LKPosPrinter.dll" (ByVal Data As String) As Long
Public Declare Function PrintText Lib "LKPosPrinter.dll" (ByVal Data As String, ByVal Alignment As Long, ByVal options As Long, ByVal TextSize As Long) As Long
Public Declare Function PrintTextTcpip Lib "LKPosPrinter.dll" (ByVal Data As String, ByVal Alignment As Long, ByVal options As Long, ByVal TextSize As Long) As Long
Public Declare Function PrintNormal Lib "LKPosPrinter.dll" (ByVal Data As String) As Long
Public Declare Function PrintNormalTcpip Lib "LKPosPrinter.dll" (ByVal Data As String) As Long
Public Declare Function PrintBarCode Lib "LKPosPrinter.dll" (ByVal Data As String, ByVal Symbology As Long, ByVal Height As Long, ByVal Width As Long, ByVal Alignment As Long, ByVal TextPosition As Long) As Long
Public Declare Function PrintBarCodeTcpip Lib "LKPosPrinter.dll" (ByVal Data As String, ByVal Symbology As Long, ByVal Height As Long, ByVal Width As Long, ByVal Alignment As Long, ByVal TextPosition As Long) As Long
Public Declare Function PrinterSts Lib "LKPosPrinter.dll" () As Long
Public Declare Function PrinterStsTcpip Lib "LKPosPrinter.dll" () As Long
Public Declare Function OpenDrawer Lib "LKPosPrinter.dll" (ByVal DrawerPinNum As Long, ByVal PulseOnTime As Long, ByVal PulseOffTime As Long) As Long
Public Declare Function OpenDrawerTcpip Lib "LKPosPrinter.dll" (ByVal DrawerPinNum As Long, ByVal PulseOnTime As Long, ByVal PulseOffTime As Long) As Long
Public Declare Function DrawerSts Lib "LKPosPrinter.dll" () As Long
Public Declare Function DrawerStsTcpip Lib "LKPosPrinter.dll" () As Long
Public Declare Function CutPaper Lib "LKPosPrinter.dll" () As Long
Public Declare Function CutPaperTcpip Lib "LKPosPrinter.dll" () As Long
Public Declare Function PrintStart Lib "LKPosPrinter.dll" () As Long
Public Declare Function PrintStop Lib "LKPosPrinter.dll" () As Long

Dim lResult    As Long

Dim ESC        As String * 1
Dim fDate      As String
Dim LPT_No     As String
Dim BaudRate   As String

'---------------------------------------------------------------------------------------
Private strMaxLng       As String
Private strTempStr      As String

Private Page_Count      As Integer  ' 보관증에 출력될 상품의 총 갯수
Private sPage_count     As Integer  ' 보관증의  전체 페이지수
Private Page_Item_Count As Integer  ' 한페이지에 출력될 상품의 갯수

Private iLine     As Integer
Private iLine2    As Integer
Private GRD_TOT   As Integer
Private GRD_S_TOT As Integer
Private iPage     As Integer
Private m         As Integer
Private Sub_TOT   As Integer

'*****************************************************************************
'기    능 : 한글을 2Byte로 처리하여 Mid함수로 처리한다.
'인    수 : sInStr As String Mid처리할 스트링
'           iStart As Integer 시작위치
'           iCnt As Integer Mid할 스트링 숫
'리 턴 값 : Mid된 결과 스트링
'사 용 예 : strTemp = gfMid("무궁화꽃이피었습니다",3,6)
'*****************************************************************************
Public Function MidH(sInStr As String, iStart As Integer, iCnt As Integer) As String
    MidH = StrConv(MidB(StrConv(sInStr, vbFromUnicode), iStart, iCnt), vbUnicode)
End Function

'*****************************************************************************
'기 능    : 한글을 2Byte로 처리하여 Len함수로 처리한다.
'인 수    : strString As String Length를 구할 스트링
'리 턴 값 : 2바이트로 처리된 Length
'사 용 예 : intLen = gfLen("무궁화꽃이피었습니다")
'*****************************************************************************
Public Function LenH(strString As String) As Integer
    LenH = LenB(StrConv(strString, vbFromUnicode))
End Function

'내용 부분 출력
Private Sub LKPOS_Center()
    Dim PrintData1 As String
    Dim PrintData2 As String
    Dim PrintData3 As String
    Dim PrintData4 As String
    Dim PrintData5 As String
    
    m = 0 ' 보관증 출력 라인 초기화
    
    If (iLine + Page_Item_Count) > Page_Count Then
        Sub_TOT = Page_Count
    Else
        Sub_TOT = iLine + Page_Item_Count - 1
    End If
    
    PrintNormal "==========================================" + vbCrLf
    PrintNormal "택번호  의류            색상 작업     금액" + vbCrLf
                '1234567 123456789012345 1234 12345 1234567
    PrintNormal "------------------------------------------" + vbCrLf
    
    For i = iLine To Sub_TOT
        m = m + 1
        
        PrintData1 = Format(Right(CStr(FPArray(i, 1)), 6), "00-0000") & " " '택번호
        PrintData2 = CStr(FPArray(i, 2)) '의류명
        PrintData3 = CStr(FPArray(i, 3)) '색상
        PrintData4 = CStr(FPArray(i, 5)) '내용
        PrintData5 = CStr(FPArray(i, 4)) '금액
        
        If Trim(PrintData2) = "" Then Exit For '택번호(PrintData1) 는 수선인 경우 없다.
        
        If LenH(PrintData2) > 15 Then
            PrintData2 = MidH(PrintData2, 1, 15) & " "                          '의류명
        Else
            PrintData2 = PrintData2 + String(16 - LenH(PrintData2), " ")        '의류명
        End If
        
        If LenH(PrintData3) > 4 Then
            PrintData3 = MidH(PrintData3, 1, 4) & " "                           '색상
        Else
            PrintData3 = Trim(PrintData3) + String(5 - LenH(PrintData3), " ")   '색상
        End If
        
        If LenH(PrintData4) > 5 Then
            PrintData4 = MidH(PrintData4, 1, 5) & " "                           '내용
        Else
            PrintData4 = Trim(PrintData4) + String(6 - LenH(PrintData4), " ")   '내용
        End If
        
        If LenH(PrintData5) > 7 Then
            PrintData5 = MidH(PrintData5, 1, 7)                                 '금액
        Else
            PrintData5 = String(7 - LenH(PrintData5), " ") + Trim(PrintData5)   '금액
        End If
        
        If Trim(PrintData1) = "" Then
            PrintNormal "수선    " + PrintData2 + PrintData3 + PrintData4 + PrintData5 + vbCrLf
        Else
            PrintNormal PrintData1 + PrintData2 + PrintData3 + PrintData4 + PrintData5 + vbCrLf
        End If
        
        If Trim(FPArray(i, 6)) <> "" Then
            PrintNormal "        ->" + Trim(FPArray(i, 6)) + vbCrLf '상표
        End If
    Next i
    
    PrintNormal "------------------------------------------" + vbCrLf
End Sub

Private Sub LKPOS_GropGoodsInfo()
    ' 세트 내역이 없을경우 출력하지 않는다.
    If 세트상품정보.d세트수량합계 <= 0 Then Exit Sub
    
    If Format(Date, "YYYY-MM-DD") <= "20091231" Then
        PrintNormal "경품추첨은 당사 홈페이지 " & Chr(34) & "경품이벤트 참여하기" & Chr(34) & "에 "
        PrintNormal "응모하신 고객분에 한하여 추첨합니다. 12월 31일까지" + vbCrLf
        
        '택번호
        'strTempStr = strMaxLng
        'RSet strTempStr = Format(세트상품정보.d최종수령액, "#,#0")
            
        If m_세트응모번호수량 = 1 Then
            PrintNormal "  경품응모번호 : " & m_세트응모번호(0) & Space(15) & " 증정매수 : " & Format(세트상품정보.d무료세탁권수량, "@@") & " 장" + vbCrLf
        ElseIf m_세트응모번호수량 = 2 Then
            PrintNormal "  경품응모번호 : " & m_세트응모번호(0) & ", " & m_세트응모번호(1) & Space(5) & " 증정매수 : " & Format(세트상품정보.d무료세탁권수량, "@@") & " 장" + vbCrLf
        End If
    End If
    
    strTempStr = strMaxLng
    RSet strTempStr = Format(세트상품정보.d전체금액, "#,#0")
    PrintNormal "세트할인전금액 : " & strTempStr + " 원" + vbCrLf      '세트할인전금액
    
    strTempStr = strMaxLng '"123456789"
    RSet strTempStr = Format(세트상품정보.d세트할인금액, "#,#0")
    PrintNormal "  세트기본할인 : " & strTempStr + " 원" + vbCrLf          '세트기본할인
    
    strTempStr = strMaxLng
    RSet strTempStr = Format(세트상품정보.d전체할인금액, "#,#0")
    PrintNormal " 세트할인 금액 : " & strTempStr + " 원" + vbCrLf         '세트할인 금액
    
    strTempStr = strMaxLng '"123456789"
    RSet strTempStr = Format(세트상품정보.d에누리할인금액, "#,#0")
    PrintNormal "   에누리 할인 : " & strTempStr + " 원" + vbCrLf           '에누리  할인
    
    strTempStr = strMaxLng
    RSet strTempStr = Format(세트상품정보.d최종수령액, "#,#0")
    PrintNormal "세트할인후금액 : " & strTempStr + " 원" + vbCrLf        '세트할인후금액
    
    strTempStr = "2:" & 세트상품정보.d2세트수량 & ",3:" & 세트상품정보.d3세트수량 & "," & _
                 "4:" & 세트상품정보.d4세트수량 & ",5:" & 세트상품정보.d4세트수량 & "," & _
                 "빅:" & 세트상품정보.d5세트수량
    PrintNormal "구성 : " & strTempStr + vbCrLf
                
    'strTempStr = "123456789"
    'RSet strTempStr = Format(세트상품정보.d세트금액, "#,#0")
    'PrintNormal "세트품목금액: " & strTempStr + vbCrLf
    
    PrintNormal "------------------------------------------" + vbCrLf
End Sub

Private Sub LKPOS_Bottom()
    ' 마지막 장일경우 전체 합계및 금액 출력
    If iPage = sPage_count Or sPage_count = 1 Then
        PrintNormal "      접    수 : "
        PrintNormal ESC + "|bC" + ESC + "|2C" + String(8 - Len(Trim(CStr(FPrtBottom.Sum))), " ") + Trim(FPrtBottom.Sum) + " 점" + vbCrLf
        
        PrintNormal "      금    액 : "
        PrintNormal ESC + "|bC" + ESC + "|2C" + String(8 - Len(Trim(CStr(FPrtBottom.Account0))), " ") + Trim(FPrtBottom.Account0) + " 원" + vbCrLf
        PrintNormal "------------------------------------------" + vbCrLf
        PrintNormal "      전일미수 : " + String(10 - Len(CStr(FPrtBottom.OldDayMisu)), " ") + FPrtBottom.OldDayMisu + " 원" + vbCrLf
        PrintNormal "  사용마일리지 : " + String(10 - Len(FPrtBottom.MilUser), " ") + FPrtBottom.MilUser + " 원" + vbCrLf
        
        PrintNormal "      미수합계 : " + String(10 - Len(CStr(FPrtBottom.MiSuTotal)), " ") + FPrtBottom.MiSuTotal + " 원" + vbCrLf
        PrintNormal "  마일리지잔액 : " + String(10 - Len(FPrtBottom.MilMoney), " ") + FPrtBottom.MilMoney + " 원" + vbCrLf
                        
        If 가맹점정보.마일리지여부 = "Y" Then
            PrintNormal "  누적마일리지 : " + String(10 - Len(CStr(FPrtBottom.MilAddMoney)), " ") + FPrtBottom.MilAddMoney + " 원" + vbCrLf
            PrintNormal "      수 령 액 : " + String(10 - Len(Trim(CStr(FPrtBottom.Account1))), " ") + Trim(FPrtBottom.Account1) + " 원" + vbCrLf
        Else
            PrintNormal "      수 령 액 : " + String(10 - Len(Trim(CStr(FPrtBottom.Account1))), " ") + Trim(FPrtBottom.Account1) + " 원" + vbCrLf
        End If
    
        PrintNormal "      적립금액 : " + String(10 - Len(CStr(FPrtBottom.SuGumMonye)), " ") + FPrtBottom.SuGumMonye + " 원" + vbCrLf
        PrintNormal "      잔    액 : " + String(10 - Len(CStr(FPrtBottom.Account2)), " ") + FPrtBottom.Account2 + " 원" + vbCrLf
    End If
            
    PrintNormal "------------------------------------------" + vbCrLf
    PrintNormal "가맹점명 : " + Trim(FPrtBottom.DName) + "   쿠폰 : " + FPrtBottom.CouponMoney + vbCrLf
    PrintNormal "전화번호 : " + Trim(FPrtBottom.DTel) + vbCrLf
    
    'PrintNormal "전화번호 : " + FPrtBottom.DTel + " (" + CStr(iPage) & "/" & CStr(sPage_count) + ")" + vbCrLf ' 페이지/전체 페이지
End Sub

'출력값 초기화
Private Sub LKPOS_Init(prtNum As String, prtTel As String)
    With FPrtBottom
        .Account0 = ""
        .Account1 = ""
        .Account2 = ""
        .DName = ""
        .DTel = ""
        .MilAddMoney = ""
        .MilMoney = ""
        .MilUser = ""
        .MiSuTotal = ""
        .OldDayMisu = ""
        .SuGumMonye = ""
        .Sum = ""
        .CouponCnt = ""
        .CouponMoney = ""
        .CouponNum = ""
    End With
    
    '--------------------------------------------------------------
    ' 보관증 출력 상단 자료 초기화
    '--------------------------------------------------------------
    Query = "SELECT * FROM TB_보관증"
    Query = Query & " WHERE 일련번호 =  " & Val(prtNum)
    Query = Query & "   AND 고객전화 = '" & prtTel & "'"
    Query = Query & " ORDER BY 택번호"
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If SUBRs.RecordCount > 0 Then
        SUBRs.MoveLast
        Page_Count = SUBRs.RecordCount
        SUBRs.MoveFirst
    Else
        SUBRs.Close
        Set SUBRs = Nothing
        
        Debug.Print "보관증 출력 없음. (오류)"
        Exit Sub
    End If

    FPrtTop.PrtNo = Format(Date, "MMDD") & "-" & SUBRs!일련번호
    
    '2009-04-02일 다시 수정  20090113 수정사항
    If 가맹점정보.고객전화번호모두출력 = "0" Then
        FPrtTop.Tel = Right("***************" & Right(Trim(SUBRs!고객전화), 4), Len(Trim(SUBRs!고객전화)))
    Else
        FPrtTop.Tel = SUBRs!고객전화 & ""                  '
    End If
    
    FPrtTop.Name = SUBRs!성명 & ""                         '
    FPrtTop.Addr = SUBRs!주소 & ""                         ' 고객 주소임
    FPrtTop.Date = SUBRs!접수일자 & ""                     '
    FPrtTop.RTime = SUBRs!접수시간 & ""                    '
    FPrtTop.Date2 = Format(SUBRs!인도예정일, "YYYY-MM-DD") '
    
    ' 전화 번호의 국번이 3자리일 경우 오른쪽 "@@@ "로 전달되는 것을 방지 하기위하여 trim 사용
    'FPrtTop.Code = Fun_고객코드(FPrtTop.Name, Left(SUBRs!고객전화, InStr(SUBRs!고객전화, "-") - 1), Right(Trim(SUBRs!고객전화), 4))
    FPrtTop.Code = Fun_고객코드(FPrtTop.Name, SUBRs!고객전화)
    
    Call Get_고객정보(FPrtTop.Code)
    
    '2009-04-02일 다시 수정  20090113 수정사항
    If 가맹점정보.고객전화번호모두출력 = "0" Then
        FPrtTop.HpTel = Right("***************" & Right(Trim(고객정보.휴대전화), 4), Len(Trim(고객정보.휴대전화)))
    Else
        FPrtTop.HpTel = 고객정보.휴대전화 & ""
    End If
    
    ' 보관증 출력 하단 자료 초기화
    strMaxLng = "1234567890"
    
    With FPrtBottom
        .Sum = strMaxLng
        RSet .Sum = RTrim(SUBRs!합계)
        .Account0 = strMaxLng
        RSet .Account0 = Format(SUBRs!합계금액, "#,##0")
        
        .Account1 = strMaxLng & "12345"
        
        If Val(CStr(SUBRs!마일리지)) = 0 Then
            RSet .Account1 = Format(Val(CStr(SUBRs!수령액)), "#,##0")
        Else
            RSet .Account1 = Format(Val(CStr(SUBRs!수령액)), "#,##0") & "/" & Format(Val(CStr(SUBRs!마일리지)), "#,##0")
        End If
        
        .Account2 = strMaxLng
        RSet .Account2 = Format(SUBRs!잔액, "#,#0")
    
        .MiSuTotal = strMaxLng
        RSet .MiSuTotal = Format(SUBRs!미수합계, "#,#0") 'Format(고객정보.미수금액, "#,#0")
        
        .OldDayMisu = strMaxLng
        RSet .OldDayMisu = Format(SUBRs!전일미수, "#,#0") '고객정보.미수금액 - SUBRs!잔액
        
        .SuGumMonye = strMaxLng
        RSet .SuGumMonye = Format(SUBRs!수금액, "#,#0")
    
    ' 사용마일리지, 마일리지 잔액, 누적 마일리지
        .MilMoney = strMaxLng
        RSet .MilMoney = Format(SUBRs!마일리지잔액, "#,#0") ' Format(마일리지.마일리지, "#,##0")
        
        .MilUser = strMaxLng
        RSet .MilUser = Format(SUBRs!마일리지, "#,##0")
        
        .MilAddMoney = strMaxLng
        RSet .MilAddMoney = Format(GetMileageMoneyToPoint(SUBRs!누적마일리지 & ""), "#,#0")
        
        ' 20090529일 수정전 원문..
        ' 수정 이유 : 누적마일리지 내용 출력을 최종 발생 금액에 해당하는 비율의 포인트로 출력 하도록 변경
        'RSet .MilAddMoney = Format(SUBRs!누적마일리지, "#,#0") 'Format(마일리지.총사용금액, "#,##0")
                    
        .DName = 가맹점정보.가맹점명 & ""
        .DTel = 가맹점정보.전화매장 & ""
        
        .CouponCnt = Format(SUBRs!CouponCnt, "#,#0")
        .CouponNum = Format(SUBRs!CouponNumber, "#,#0")
        .CouponMoney = Format(SUBRs!CouponMoney, "#,#0")
    End With
    
' 보관증 출력 중간 자료 초기화
    For i = 1 To 500
        FPArray(i, 1) = SUBRs!택번호 & ""
        FPArray(i, 2) = SUBRs!의류명 & ""
        FPArray(i, 3) = SUBRs!색상 & ""
        FPArray(i, 4) = Format(SUBRs!금액, "#,#0") & ""
        FPArray(i, 5) = SUBRs!내용 & ""
        FPArray(i, 6) = SUBRs!상표 & ""

        SUBRs.MoveNext

        If SUBRs.EOF = True Then
            Exit For
        End If
    Next i
    
    ' 세트 상품의 내역을 가저온다.
    SUBRs.MoveFirst
    
    ZeroMemory 세트상품정보, Len(세트상품정보)
    
''    '--------------------------------------------------------------------
''    '
''    '--------------------------------------------------------------------
''    Query = "SELECT * FROM TB_세트상품정보 "
''    Query = Query & " WHERE 세트Key = '" & SUBRs!세트Key & "' "
''    Set Rs = New ADODB.Recordset
''    Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
''
''    If Not Rs.EOF Then
''        With 세트상품정보
''            .d세트Key = Rs.Fields("세트Key") & ""
''
''            .d2세트수량 = Val(Rs.Fields("세트2") & "")
''            .d3세트수량 = Val(Rs.Fields("세트3") & "")
''            .d4세트수량 = Val(Rs.Fields("세트4") & "")
''            .d5세트수량 = Val(Rs.Fields("세트5") & "")
''            .d6세트수량 = Val(Rs.Fields("세트6") & "")
''
''            .d세트수량합계 = .d2세트수량 + .d3세트수량 + .d4세트수량 + .d5세트수량 + .d6세트수량
''            .d무료세탁권수량 = (.d2세트수량 * 1) + _
''                             (.d3세트수량 * 2) + _
''                             (.d4세트수량 * 3) + _
''                             (.d5세트수량 * 4) + _
''                             (.d6세트수량 * 5)
''
''            .d전체금액 = Val(Rs.Fields("정상금액") & "")
''            .d세트금액 = Val(Rs.Fields("세트금액") & "")
''
''            .d세트할인금액 = Val(Rs.Fields("세트할인금액") & "")
''            .d에누리할인금액 = Val(Rs.Fields("에누리할인금액") & "")
''            .d전체할인금액 = .d세트할인금액 + .d에누리할인금액
''            .d최종수령액 = Val(Rs.Fields("적용합계금액") & "")
''        End With
''     End If
''    Rs.Close
''    Set Rs = Nothing
    
    m_세트응모번호수량 = 0
    
''    '--------------------------------------------------------------------
''    '
''    '--------------------------------------------------------------------
''    Query = "SELECT * FROM TB_세트응모번호 "
''    Query = Query & " WHERE 세트Key = '" & CStr(SUBRs!세트Key & "") & "' "
''    Set Rs = New ADODB.Recordset
''    Rs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
''
''    If Rs.RecordCount > 0 Then
''        Rs.MoveLast
''        ReDim m_세트응모번호(Rs.RecordCount - 1)
''        Rs.MoveFirst
''
''        Do While Not Rs.EOF
''            m_세트응모번호(m_세트응모번호수량) = Rs.Fields("응모번호") & ""
''            m_세트응모번호수량 = m_세트응모번호수량 + 1
''
''            Rs.MoveNext
''        Loop
''    End If
''    Rs.Close
''    Set Rs = Nothing
    
    SUBRs.Close
    Set SUBRs = Nothing
End Sub

' 타이틀 부분 출력
Private Sub LKPOS_Title()
    PrintNormal ESC + "|bC" + ESC + "|cA" + ESC + "|4C" + "세탁물 보관증" + vbLf
    PrintNormal "" + vbCrLf
    
    If 가맹점정보.지사코드 <> M_COUPON_KLENZ_CODE Then
        If Format(Date, "YYYY-MM-DD") >= "20091207" And Format(Date, "YYYY-MM-DD") <= "20091231" Then
            PrintNormal "★★ 세트세탁서비스 출시기념 이벤트 2009-12-11 ~ 12-31일까지 ★★" + vbCrLf
            PrintNormal "1.세탁물 10%할인 2.경품이벤트 3.세트 세탁 접수시 무료 세탁권 증정" + vbCrLf
        
        ElseIf Format(Date, "YYYY-MM-DD") >= "20100101" Then
            PrintNormal "★★ 세트세탁서비스 출시★★" + vbCrLf
            PrintNormal "세트세탁 접수시 7 ~ 3% 할인서비스 제공" + vbCrLf
        End If
    End If
        
    PrintNormal "전표번호 : " + FPrtTop.PrtNo + vbCrLf                                   ' 전표번호
    PrintNormal "고객번호 : " + FPrtTop.Code + vbCrLf                                    ' 고객 번호
    
    PrintNormal "성    명 : "                                                            ' 고객 성명
    PrintNormal ESC + "|bC" + ESC + "|2C" + FPrtTop.Name + vbLf
    
    PrintNormal "전화번호 : "                                                            ' 고객 전화번호
    PrintNormal ESC + "|bC" + ESC + "|2C" + FPrtTop.Tel + vbLf
    
    PrintNormal "휴 대 폰 : "                                                            ' 고객 휴대전화
    PrintNormal ESC + "|bC" + ESC + "|2C" + FPrtTop.HpTel + vbLf
    
    PrintNormal "주    소 : " + FPrtTop.Addr + vbCrLf                                    ' 주소 (손님)
    PrintNormal "접수일자 : " + CStr(FPrtTop.Date) + " " + CStr(FPrtTop.RTime) + vbCrLf  ' 접수일자 + 접수시간
    PrintNormal "예정일자 : " + CStr(FPrtTop.Date2) + vbCrLf                             ' 인도 연도
End Sub

Public Function LKPOS_Print(prtNum As String, prtTel As String)
        
    On Error GoTo ErrRtn
    
''    '---------------------------------------------------------------------------------
''    ' OpenPort
''    '---------------------------------------------------------------------------------
''    i = 0
''
''    Do
''        lResult = OpenPort("USB", 19200)
''        DoEvents
''
''        If lResult <> 0 Then
''            i = i + 1
''
''            If i > 3 Then
''                        Query = "프린터를 점검해주세요." & vbNewLine & vbNewLine
''                Query = Query & "- 프린터에 전원이 켜져 있는지요?" & vbNewLine
''                Query = Query & "- 컴퓨터와 프린터가 정상적으로 연결되어 있는지요?" & vbNewLine
''                Query = Query & "- 프린터에 용지가 있는지요?" & vbNewLine
''
''                MsgBox Query, vbCritical, "오류"
''
''                Exit Function
''            End If
''        End If
''    Loop Until lResult = 0

    '---------------------------------------------------------------------------------
    ' PrinterSts
    '---------------------------------------------------------------------------------
    lResult = PrinterSts

    Select Case lResult
        Case LK_STS_NORMAL         'MsgBox "No Error"
        Case LK_STS_PAPERNEAREMPTY 'MsgBox "Paper Near Empty"

        Case LK_STS_COVEROPEN      'MsgBox "Cover Open"
            MsgBox "프린터 뚜껑이 열려 있습니다.", vbCritical, "오류"

            'lResult = ClosePort
            Exit Function

        Case LK_STS_PAPEREMPTY     'MsgBox "Paper Empty"
            MsgBox "프린터 용지가 없습니다. 용지를 넣어주세요.", vbCritical, "오류"

            'lResult = ClosePort
            Exit Function

        Case LK_STS_POWEROFF       'MsgBox "Power Off"
            MsgBox "프린터 전원이 꺼져 있습니다.", vbCritical, "오류"

            'lResult = ClosePort
            Exit Function
    End Select

    ESC = Chr(&H1B)

    PrintStart

    ' 사용 값들을 초기화 한다.
    iPage = 0
    iLine = 0
    iLine2 = 0
    GRD_TOT = 0
    GRD_S_TOT = 0
    
    Erase FPArray
    
    Page_Item_Count = GetPrtItemCount("보관증")     ' 보관증에 출력될 상품 갯수
    
    Call LKPOS_Init(prtNum, prtTel) '전체 출력 갯수및 출력 내용 변수에 초기화 'GoSub Print_Value_Init
    
    If (Page_Count <= 0) Then
        Exit Function
    End If

    '----------------------------------------------------
    ' 세트 관련 최종 출력 내용의 4칸을 할당 하여 세트 내용을 출력한다.
    If FPrtTop.Date <= "2009-12-31" Then
        If 세트상품정보.d세트수량합계 > 0 Then Page_Count = Page_Count + 6
    Else
        If 세트상품정보.d세트수량합계 > 0 Then Page_Count = Page_Count + 3
    End If
    '----------------------------------------------------
    
    ' 전체 출력 페이지 구하기
    If (Page_Count Mod Page_Item_Count) <> 0 Then
        sPage_count = Int(Page_Count / Page_Item_Count) + 1
    Else
        sPage_count = Int(Page_Count / Page_Item_Count)
    End If
    
    '전체 페이지 까지 반복.
    For iPage = 1 To sPage_count
        ' 첫번째 장이나 마지막 장일경우
        If (iPage = sPage_count) Or (sPage_count = 1) Then
            iLine = iLine2 + 1
            iLine2 = Page_Count   ' frmINPUT.ListView1.ListItems.Count
            
            Call LKPOS_Title         'GoSub Print_Title
            Call LKPOS_Center        'GoSub Print_Center
            Call LKPOS_GropGoodsInfo '세트 상품 관련 내용 출력 'GoSub Print_GropGoodsINFO
            Call LKPOS_Bottom        'GoSub Print_Bottom
        Else
            ' 중간 페이지 일 경우
            iLine = iLine2 + 1
            iLine2 = iLine2 + Page_Item_Count
            
            Call LKPOS_Title  'GoSub Print_Title
            Call LKPOS_Center 'GoSub Print_Center
            Call LKPOS_Bottom 'GoSub Print_Bottom
        End If
    Next iPage


    PrintNormal ESC + "|fP" '커팅

    PrintStop

    'lResult = ClosePort

    Screen.MousePointer = 0
    
    Exit Function
    
ErrRtn:
    'lResult = ClosePort
    
    MsgBox " 프린터를 확인해 주십시요 ! " & vbNewLine & vbNewLine & Err.Description, vbCritical, "출력오류발생"
End Function
