Attribute VB_Name = "basKSNET"
Option Explicit

Public Declare Function ReqAppr Lib "KSNet_ADSL.dll" Alias "RequestApproval" ( _
ByVal ipAddr As String, ByVal sPort As Integer, _
ByVal sMedia As Integer, _
ByVal RequestMsg As String, ByVal RequestLen As Integer, _
ByVal sRecvMsg As String, ByVal timeout As Integer, _
ByVal options As Integer) As Long

Public Account_Form    As String  '접수결제인지, 출고결제인지 체크
Public KS7500_PortOpen As Boolean 'KS7500 Port Open 여부



'Public Sub 카드결제_Report(KS7500i As KSNetPOSLib, Spread As fpSpread, iPaper As Integer)
'    On Error GoTo ErrRtn
'
'    Dim 카드번호 As String
'    Dim 거래일시 As String
'    Dim 카드결제 As Double
'
'    ' 가맹젘 카드전표 출력 안함일 경우 처리하지 않는다.
'    If 가맹점정보.가맹점카드영수증출력YN = "N" And iPaper = 2 Then Exit Sub
'
'
'    With Spread
'        For i = 1 To .MaxRows
'            .Row = i
'
'            If iPaper = 1 Then
'                Call KS7500i.PrintString("    신용승인(고객용)", 2)
'            Else
'                Call KS7500i.PrintString("    신용승인(가맹점용)", 2)
'            End If
'
'            Call KS7500i.LineFeed(1)
'            Call KS7500i.PrintString("-----------------------------------------------", 1)
'
'            Query = "SELECT    ISNULL(가맹점명, '')     AS 가맹점명"
'            Query = Query & ", ISNULL(사업자번호, '')   AS 사업자번호"
'            Query = Query & ", ISNULL(대표자명, '')     AS 대표자명"
'            Query = Query & ", ISNULL(매장전화번호, '') AS 매장전화번호"
'            Query = Query & ", ISNULL(사업장주소, '')   AS 사업장주소"
'            Query = Query & " FROM TB_기본정보"
'            Set ADORs = New ADODB.Recordset
'            ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'            If ADORs.EOF Then
'                Call KS7500i.PrintString("상 호 명 : ", 1)
'                Call KS7500i.PrintString("사업자No : ", 1)
'                Call KS7500i.PrintString("대 표 자 : ", 1)
'                Call KS7500i.PrintString("전화번호 : ", 1)
'                Call KS7500i.PrintString("주    소 : ", 1)
'
'            Else
'                Call KS7500i.PrintString("상 호 명 : " & ADORs!가맹점명, 1)
'                Call KS7500i.PrintString("사업자No : " & ADORs!사업자번호, 1)
'                Call KS7500i.PrintString("대 표 자 : " & ADORs!대표자명, 1)
'                Call KS7500i.PrintString("전화번호 : " & ADORs!매장전화번호, 1)
'                Call KS7500i.PrintString("주    소 : " & ADORs!사업장주소, 1)
'            End If
'            ADORs.Close
'            Set ADORs = Nothing
'
'            Call KS7500i.PrintString("-----------------------------------------------", 1)
'
'            .Col = 3: 거래일시 = "거래일시 : " & Format(.Text, "2000/00/00") + " "
'            .Col = 4: 거래일시 = 거래일시 & Format(.Text, "00:00")
'
'            .Col = 11: 카드번호 = Left(.Text, 4) & "-"
'                       카드번호 = 카드번호 & Mid(.Text, 5, 4) & "-"
'                       카드번호 = 카드번호 & "****-"
'                       카드번호 = 카드번호 & Mid(.Text, 13, 4)
'
'            Call KS7500i.PrintString(거래일시, 1)
'            Call KS7500i.PrintString("카드번호 : " & 카드번호, 1)
'
'            .Col = 8: Call KS7500i.PrintString("카드종류 : " & .Text, 1)
'            .Col = 2: Call KS7500i.PrintString("승인번호 : " & .Text, 1)
'
'            .Col = 5
'            If .Text = "00" Then
'                Call KS7500i.PrintString("할부개월 : 일시불", 1)
'            Else
'                Call KS7500i.PrintString("할부개월 : " & CStr(.Value) & " 개월", 1)
'            End If
'
''            .Col = 6: Call KS7500i.PrintString("결제금액 : " & Format(.Text, "#,##0") & "원", 1)
'            .Col = 6: Call KS7500i.PrintString("승인금액 : " & Format(.Text, "#,##0") & "원", 1)
'
'            .Col = 6: 카드결제 = CDbl(.Text)
'
'            Call KS7500i.PrintString("과세금액 : " + Format(카드결제 - (카드결제 - (카드결제 / 1.1)), "#,##0") + "원", 1)
'            Call KS7500i.PrintString("부가세액 : " + Format(카드결제 - (카드결제 / 1.1), "#,##0") + "원", 1)
'            Call KS7500i.PrintString("승인금액 : " + Format(카드결제, "#,##0") + "원", 2)
'
'            Call KS7500i.PrintString("-----------------------------------------------", 1)
'            Call KS7500i.PrintString("(서명)", 1)
'            Call KS7500i.PrintImage(m_SignFileName, 5)
'            Call KS7500i.LineFeed(1)
'            Call KS7500i.CutPaper
'        Next i
'    End With
'
'    Exit Sub
'
'ErrRtn:
'    Call Error_Msg("", Err.Source, Err.Number, Err.description)
'
'    Screen.MousePointer = 0
'End Sub

'Public Sub 현금영수증_Report(KS7500i As KSNetPOSLib, Spread As fpSpread, iPaper As Integer)
'    On Error GoTo ErrRtn
'
'    Dim 카드번호 As String
'    Dim 승인일자 As String
'    Dim 카드결제 As Double
'
'    With Spread
'        .Col = 1
'        .Row = 1
'
'        If .Text = "" Then Exit Sub
'
'        ' 가맹젘 카드전표 출력 안함일 경우 처리하지 않는다.
'        If 가맹점정보.가맹점카드영수증출력YN = "N" And iPaper = 2 Then Exit Sub
'
'
'        If iPaper = 1 Then
'            Call KS7500i.PrintString("현금영수증(고객용)", 2)
'        Else
'            Call KS7500i.PrintString("현금영수증(가맹점용)", 2)
'        End If
'
'        Call KS7500i.LineFeed(1)
'        Call KS7500i.PrintString("-----------------------------------------------", 1)
'
'        Query = "SELECT    ISNULL(가맹점명, '')     AS 가맹점명"
'        Query = Query & ", ISNULL(사업자번호, '')   AS 사업자번호"
'        Query = Query & ", ISNULL(대표자명, '')     AS 대표자명"
'        Query = Query & ", ISNULL(매장전화번호, '') AS 매장전화번호"
'        Query = Query & ", ISNULL(사업장주소, '')   AS 사업장주소"
'        Query = Query & " FROM TB_기본정보"
'        Set ADORs = New ADODB.Recordset
'        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'        If ADORs.EOF Then
'            Call KS7500i.PrintString("상 호 명 : ", 1)
'            Call KS7500i.PrintString("사업자No : ", 1)
'            Call KS7500i.PrintString("대 표 자 : ", 1)
'            Call KS7500i.PrintString("전화번호 : ", 1)
'            Call KS7500i.PrintString("주    소 : ", 1)
'
'        Else
'            Call KS7500i.PrintString("상 호 명 : " & ADORs!가맹점명, 1)
'            Call KS7500i.PrintString("사업자No : " & ADORs!사업자번호, 1)
'            Call KS7500i.PrintString("대 표 자 : " & ADORs!대표자명, 1)
'            Call KS7500i.PrintString("전화번호 : " & ADORs!매장전화번호, 1)
'            Call KS7500i.PrintString("주    소 : " & ADORs!사업장주소, 1)
'        End If
'        ADORs.Close
'        Set ADORs = Nothing
'
'        Call KS7500i.PrintString("-----------------------------------------------", 1)
'
'        .Row = 6: 카드번호 = Trim(.Text) & "" '사용자정보
'
'        If Len(카드번호) = 13 Then
'            '주민번호
'            카드번호 = Left(카드번호, Len(카드번호) - 7) & "*******"
'
'        ElseIf Len(카드번호) = 19 Then
'            '현금영수증 카드
'            카드번호 = Left(카드번호, Len(카드번호) - 7) & "******"
'        Else
'            '휴대폰
'            If Len(카드번호) > 4 Then
'                카드번호 = Left(카드번호, Len(카드번호) - 4) & "****"
'            End If
'        End If
'
'        Call KS7500i.PrintString("구 매 자 : " + 카드번호, 1)
'
'        .Row = 4
'        ' KS7500,7050     모듈은 0, 소득공제용, 1.지출증빙용
'        ' KS4060 보안인증 모듈은 1, 소득공제용, 2.지출증빙용
'        If 가맹점정보.CAT단말기종류 <> "KS4060 보안인증" Then
'            Call KS7500i.PrintString("거래구분 : " & IIf(.Text = "0", "소득공제용", "지출증빙용"), 1)
'        Else
'            Call KS7500i.PrintString("거래구분 : " & IIf(.Text = "1", "소득공제용", "지출증빙용"), 1)
'        End If
'
'
'        .Row = 2: 승인일자 = "승인일자 : " + Format(.Text, "2000/00/00") + " "
'        .Row = 3: 승인일자 = 승인일자 + Format(.Text, "00:00")
'
'        Call KS7500i.PrintString(승인일자, 1)
'
'        .Row = 1: Call KS7500i.PrintString("승인번호 : " & .Text, 1)
''        .Row = 5: Call KS7500i.PrintString("승인금액 : " & .Text & "원", 1)
'
'        .Row = 5: 카드결제 = CDbl(.Text)
'
'        Call KS7500i.PrintString("과세금액 : " + Format(카드결제 - (카드결제 - (카드결제 / 1.1)), "#,##0") + "원", 1)
'        Call KS7500i.PrintString("부가세액 : " + Format(카드결제 - (카드결제 / 1.1), "#,##0") + "원", 1)
'        Call KS7500i.PrintString("승인금액 : " + Format(카드결제, "#,##0") + "원", 2)
'
'        Call KS7500i.PrintString("-----------------------------------------------", 1)
'        .Row = 10: Call KS7500i.PrintString(.Text, 1)
'        .Row = 11: Call KS7500i.PrintString(.Text, 1)
'        Call KS7500i.LineFeed(1)
'        Call KS7500i.CutPaper
'    End With
'
'    Exit Sub
'
'ErrRtn:
'    Call Error_Msg("", Err.Source, Err.Number, Err.description)
'
'    Screen.MousePointer = 0
'End Sub

'
'Public Sub 신용카드취소_Report(KS7500i As KSNetPOSLib, 승인번호 As String, 승인일자 As String, 승인시간 As String)
'    On Error GoTo ErrRtn
'
'    Dim CommPort As String
'    Dim BaudRate As String
'    Dim 카드번호 As String
'
'    CommPort = GetIniStr("VAN", "KS7500_CommPort", "", iniFile)
'    BaudRate = GetIniStr("VAN", "KS7500_BaudRate", "", iniFile)
'
'    '--------------------------------------------------------------------
'    ' 1. 신용카드승인 취소 내역 출력
'    '--------------------------------------------------------------------
'    Query = "SELECT * FROM TB_신용카드승인"
'    Query = Query & " WHERE 승인번호 = '" & 승인번호 & "'"
'    Query = Query & "   AND 승인일자 = '" & 승인일자 & "'"
'    Query = Query & "   AND 승인시간 = '" & 승인시간 & "'"
'    Query = Query & "   AND SUBSTRING(메시지2,1,2) = '취소'"
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    If Not ADORs.EOF Then
'        Do
'            Rtn = KS7500i.CheckPort(CInt(CommPort), CLng(BaudRate))
'            DoEvents
'
'            If Rtn < 0 Then
'                i = i + 1
'
'                If i > 3 Then
'                    MsgBox "카드단말기 장치가 연결되어 있지 않습니다", vbCritical, "오류"
'
'                    Exit Do
'                End If
'            Else
'                Call KS7500i.SetConfig("", Rtn, CLng(BaudRate))    '첫번째 인자는 "" 로 넣어 준다.
'
'                KS7500i.InitPrint
'
'                For i = 1 To 2
'                    If i = 1 Then
'                        Call KS7500i.PrintString("  신용승인취소(가맹점용)", 2)
'                    Else
'                        Call KS7500i.PrintString("  신용승인취소(고객용)", 2)
'                    End If
'
'                    Call KS7500i.LineFeed(1)
'
'                    Call KS7500i.PrintString("-----------------------------------------------", 1)
'
'                    Query = "SELECT    ISNULL(가맹점명, '')     AS 가맹점명"
'                    Query = Query & ", ISNULL(사업자번호, '')   AS 사업자번호"
'                    Query = Query & ", ISNULL(대표자명, '')     AS 대표자명"
'                    Query = Query & ", ISNULL(매장전화번호, '') AS 매장전화번호"
'                    Query = Query & ", ISNULL(사업장주소, '')   AS 사업장주소"
'                    Query = Query & " FROM TB_기본정보"
'                    Set SUBRs = New ADODB.Recordset
'                    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'                    If SUBRs.EOF Then
'                        Call KS7500i.PrintString("상 호 명 : ", 1)
'                        Call KS7500i.PrintString("사업자No : ", 1)
'                        Call KS7500i.PrintString("대 표 자 : ", 1)
'                        Call KS7500i.PrintString("전화번호 : ", 1)
'                        Call KS7500i.PrintString("주    소 : ", 1)
'
'                    Else
'                        Call KS7500i.PrintString("상 호 명 : " + SUBRs!가맹점명, 1)
'                        Call KS7500i.PrintString("사업자No : " + SUBRs!사업자번호, 1)
'                        Call KS7500i.PrintString("대 표 자 : " + SUBRs!대표자명, 1)
'                        Call KS7500i.PrintString("전화번호 : " + SUBRs!매장전화번호, 1)
'                        Call KS7500i.PrintString("주    소 : " + SUBRs!사업장주소, 1)
'                    End If
'                    SUBRs.Close
'                    Set SUBRs = Nothing
'
'                    Call KS7500i.PrintString("-----------------------------------------------", 1)
'                    Call KS7500i.PrintString("거래일시 : " + Format(ADORs!승인일자, "2000/00/00") + " " + Format(ADORs!승인시간, "00:00"), 1)
'
'                    카드번호 = Left(ADORs!카드번호, 4) & "-"
'                    카드번호 = 카드번호 & Mid(ADORs!카드번호, 5, 4) & "-"
'                    카드번호 = 카드번호 & "****-"
'                    카드번호 = 카드번호 & Mid(ADORs!카드번호, 13, 4)
'
'                    Call KS7500i.PrintString("카드번호 : " + 카드번호, 1)
'                    Call KS7500i.PrintString("카드종류 : " + ADORs!카드종류명, 1)
'                    Call KS7500i.PrintString("승인번호 : " + ADORs!승인번호, 1)
'
'                    If ADORs!할부기간 = "00" Then
'                        Call KS7500i.PrintString("할부개월 : 일시불", 1)
'                    Else
'                        Call KS7500i.PrintString("할부개월 : " + CInt(ADORs!할부기간) + " 개월", 1)
'                    End If
'
'                    Call KS7500i.PrintString("과세금액 : " + Format(CDbl(ADORs!결제금액) - (CDbl(ADORs!결제금액) - (CDbl(ADORs!결제금액) / 1.1)), "#,##0") + "원", 1)
'                    Call KS7500i.PrintString("부가세액 : " + Format(CDbl(ADORs!결제금액) - (CDbl(ADORs!결제금액) / 1.1), "#,##0") + "원", 1)
'                    Call KS7500i.PrintString("승인금액 : " + Format(CDbl(ADORs!결제금액), "#,##0") + "원", 2)
'                    Call KS7500i.PrintString("-----------------------------------------------", 1)
'                    Call KS7500i.LineFeed(1)
'                    Call KS7500i.CutPaper
'                Next i
'
'                KS7500i.ClosePort
'                DoEvents
'            End If
'        Loop Until Rtn > 0
'    End If
'    ADORs.Close
'    Set ADORs = Nothing
'
'    Exit Sub
'
'ErrRtn:
'    Call Error_Msg("", Err.Source, Err.Number, Err.description)
'
'    Screen.MousePointer = 0
'End Sub

'
'Public Sub 현금영수증취소_Report(승인번호 As String, 승인일자 As String, 승인시간 As String)
''    On Error GoTo ErrRtn
''
''    Dim CommPort As String
''    Dim BaudRate As String
''    Dim 카드번호 As String
''
''    CommPort = GetIniStr("VAN", "KS7500_CommPort", "", iniFile)
''    BaudRate = GetIniStr("VAN", "KS7500_BaudRate", "", iniFile)
''
''    '--------------------------------------------------------------------
''    ' 1. TB_현금영수증 취소 내역 출력
''    '--------------------------------------------------------------------
''    Query = "SELECT * FROM TB_현금영수증"
''    Query = Query & " WHERE 승인번호 = '" & 승인번호 & "'"
''    Query = Query & "   AND 승인일자 = '" & 승인일자 & "'"
''    Query = Query & "   AND 승인시간 = '" & 승인시간 & "'"
''    Query = Query & "   AND SUBSTRING(메시지2,1,2) = '취소'"
''    Set ADORs = New ADODB.Recordset
''    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
''
''    If Not ADORs.EOF Then
''        Do
''            Rtn = KS7500i.CheckPort(CInt(CommPort), CLng(BaudRate))
''            DoEvents
''
''            If Rtn < 0 Then
''                i = i + 1
''
''                If i > 3 Then
''                    MsgBox "카드단말기 장치가 연결되어 있지 않습니다", vbCritical, "오류"
''
''                    Exit Do
''                End If
''            Else
''                Call KS7500i.SetConfig("", Rtn, CLng(BaudRate))    '첫번째 인자는 "" 로 넣어 준다.
''
''                KS7500i.InitPrint
''
''                For i = 1 To IIf(가맹점정보.가맹점카드영수증출력YN = "Y", 2, 1)
''                    If i = 1 Then
''                        Call KS7500i.PrintString("현금영수증취소(고객용)", 4)
''                    Else
''                        Call KS7500i.PrintString("현금영수증취소(가맹점용)", 4)
''                    End If
''
''                    Call KS7500i.LineFeed(1)
''
''                    Call KS7500i.PrintString("-----------------------------------------------", 1)
''
''                    Query = "SELECT    ISNULL(가맹점명, '')     AS 가맹점명"
''                    Query = Query & ", ISNULL(사업자번호, '')   AS 사업자번호"
''                    Query = Query & ", ISNULL(대표자명, '')     AS 대표자명"
''                    Query = Query & ", ISNULL(매장전화번호, '') AS 매장전화번호"
''                    Query = Query & ", ISNULL(사업장주소, '')   AS 사업장주소"
''                    Query = Query & " FROM TB_기본정보"
''                    Set SUBRs = New ADODB.Recordset
''                    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
''
''                    If SUBRs.EOF Then
''                        Call KS7500i.PrintString("상 호 명 : ", 1)
''                        Call KS7500i.PrintString("사업자No : ", 1)
''                        Call KS7500i.PrintString("대 표 자 : ", 1)
''                        Call KS7500i.PrintString("전화번호 : ", 1)
''                        Call KS7500i.PrintString("주    소 : ", 1)
''
''                    Else
''                        Call KS7500i.PrintString("상 호 명 : " + SUBRs!가맹점명, 1)
''                        Call KS7500i.PrintString("사업자No : " + SUBRs!사업자번호, 1)
''                        Call KS7500i.PrintString("대 표 자 : " + SUBRs!대표자명, 1)
''                        Call KS7500i.PrintString("전화번호 : " + SUBRs!매장전화번호, 1)
''                        Call KS7500i.PrintString("주    소 : " + SUBRs!사업장주소, 1)
''                    End If
''                    SUBRs.Close
''                    Set SUBRs = Nothing
''
''                    Call KS7500i.PrintString("-----------------------------------------------", 1)
''
''                    카드번호 = Trim(ADORs!사용자정보) & "" '사용자정보
''
''                    If Len(카드번호) = 13 Then
''                        '주민번호
''                        카드번호 = Left(카드번호, Len(카드번호) - 7) & "*******"
''
''                    ElseIf Len(카드번호) = 19 Then
''                        '현금영수증 카드
''                        카드번호 = Left(카드번호, Len(카드번호) - 7) & "******"
''                    Else
''                        '휴대폰
''                        If Len(카드번호) > 4 Then
''                            카드번호 = Left(카드번호, Len(카드번호) - 4) & "****"
''                        End If
''                    End If
''
''                    Call KS7500i.PrintString("구 매 자 : " + 카드번호, 1)
''
''                    ' KS7500,7050     모듈은 0, 소득공제용, 1.지출증빙용
''                    ' KS4060 보안인증 모듈은 1, 소득공제용, 2.지출증빙용
''                    If 가맹점정보.CAT단말기종류 <> "KS4060 보안인증" Then
''                        Call KS7500i.PrintString("거래구분 : " & IIf(ADORs!소득구분 = "0", "소득공제용", "지출증빙용"), 1)
''                    Else
''                        Call KS7500i.PrintString("거래구분 : " & IIf(ADORs!소득구분 = "1", "소득공제용", "지출증빙용"), 1)
''                    End If
''
''                    Call KS7500i.PrintString("승인일자 : " + Format(ADORs!승인일자, "2000/00/00") + " " + Format(ADORs!승인시간, "00:00"), 1)
''                    Call KS7500i.PrintString("승인번호 : " + ADORs!승인번호, 1)
''
''
''                    Call KS7500i.PrintString("과세금액 : " + Format(CDbl(ADORs!총금액) - (CDbl(ADORs!총금액) - (CDbl(ADORs!총금액) / 1.1)), "#,##0") + "원", 1)
''                    Call KS7500i.PrintString("부가세액 : " + Format(CDbl(ADORs!총금액) - (CDbl(ADORs!총금액) / 1.1), "#,##0") + "원", 1)
''                    Call KS7500i.PrintString("승인금액 : " + Format(ADORs!총금액, "#,##0") + "원", 2)
''
''                    Call KS7500i.PrintString("-----------------------------------------------", 1)
''                    Call KS7500i.PrintString(ADORs!국세청1, 1)
''                    Call KS7500i.PrintString(ADORs!국세청2, 1)
''                    Call KS7500i.LineFeed(1)
''                    Call KS7500i.CutPaper
''                Next i
''
''                KS7500i.ClosePort
''                DoEvents
''            End If
''        Loop Until Rtn > 0
''    End If
''    ADORs.Close
''    Set ADORs = Nothing
'
'    Exit Sub
'
'ErrRtn:
'    Call Error_Msg("", Err.Source, Err.Number, Err.description)
'
'    Screen.MousePointer = 0
'End Sub

'
'Public Sub 승인취소_Report(KS7500i As KSNetPOSLib, 고객코드 As String, 접수번호 As Long)
'    On Error GoTo ErrRtn
'
'    Dim CommPort As String
'    Dim BaudRate As String
'    Dim 카드번호 As String
'
'    Dim 거래일시 As String
'    Dim 승인일자 As String
'
'    '--------------------------------------------------------------------
'    ' 1. 신용카드승인 취소 내역 출력
'    '--------------------------------------------------------------------
'    Query = "SELECT * FROM TB_신용카드승인"
'    Query = Query & " WHERE 고객코드 = '" & 고객코드 & "'"
'    Query = Query & "   AND 승인일자 = '" & Format(Date, "YYMMDD") & "'"
'
'    If 접수번호 = 0 Then
'        Query = Query & "   AND 접수번호 = 0"
'    Else
'        Query = Query & "   AND 접수번호 = " & 접수번호
'    End If
'
'    Query = Query & "   AND SUBSTRING(메시지2,1,2) = '취소'"
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    If Not ADORs.EOF Then
'        Do Until ADORs.EOF
'            For i = 1 To IIf(가맹점정보.가맹점카드영수증출력YN = "Y", 2, 1)
'                If i = 1 Then
'                    Call KS7500i.PrintString("신용승인취소(고객용)", 2)
'                Else
'                    Call KS7500i.PrintString("신용승인취소(가맹점용)", 2)
'                End If
'
'                KS7500i.LineFeed (1)
'
'                Call KS7500i.PrintString("-----------------------------------------------", 1)
'
'                '------------------------------------------------------------------------
'                ' TB_기본정보
'                '------------------------------------------------------------------------
'                Query = "SELECT    ISNULL(가맹점명, '')     AS 가맹점명"
'                Query = Query & ", ISNULL(사업자번호, '')   AS 사업자번호"
'                Query = Query & ", ISNULL(대표자명, '')     AS 대표자명"
'                Query = Query & ", ISNULL(매장전화번호, '') AS 매장전화번호"
'                Query = Query & ", ISNULL(사업장주소, '')   AS 사업장주소"
'                Query = Query & " FROM TB_기본정보"
'                Set SUBRs = New ADODB.Recordset
'                SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'                If SUBRs.EOF Then
'                    Call KS7500i.PrintString("상 호 명 : ", 1)
'                    Call KS7500i.PrintString("사업자No : ", 1)
'                    Call KS7500i.PrintString("대 표 자 : ", 1)
'                    Call KS7500i.PrintString("전화번호 : ", 1)
'                    Call KS7500i.PrintString("주    소 : ", 1)
'
'                Else
'                    Call KS7500i.PrintString("상 호 명 : " + SUBRs!가맹점명, 1)
'                    Call KS7500i.PrintString("사업자No : " + SUBRs!사업자번호, 1)
'                    Call KS7500i.PrintString("대 표 자 : " + SUBRs!대표자명, 1)
'                    Call KS7500i.PrintString("전화번호 : " + SUBRs!매장전화번호, 1)
'                    Call KS7500i.PrintString("주    소 : " + SUBRs!사업장주소, 1)
'                End If
'                SUBRs.Close
'                Set SUBRs = Nothing
'
'                Call KS7500i.PrintString("-----------------------------------------------", 1)
'                Call KS7500i.PrintString("거래일시 : " + Format(ADORs!승인일자, "2000/00/00") + " " + Format(ADORs!승인시간, "00:00"), 1)
'
'                카드번호 = Left(ADORs!카드번호, 4) & "-"
'                카드번호 = 카드번호 & Mid(ADORs!카드번호, 5, 4) & "-"
'                카드번호 = 카드번호 & "****-"
'                카드번호 = 카드번호 & Mid(ADORs!카드번호, 13, 4)
'
'                Call KS7500i.PrintString("카드번호 : " + 카드번호, 1)
'                Call KS7500i.PrintString("카드종류 : " + ADORs!카드종류명, 1)
'                Call KS7500i.PrintString("승인번호 : " + ADORs!승인번호, 1)
'
'                If ADORs!할부기간 = "00" Then
'                    Call KS7500i.PrintString("할부개월 : 일시불", 1)
'                Else
'                    Call KS7500i.PrintString("할부개월 : " + CInt(ADORs!할부기간) + " 개월", 1)
'                End If
'
'                Call KS7500i.PrintString("결제금액 : " + CStr(ADORs!결제금액) + "원", 2)
'                Call KS7500i.PrintString("-----------------------------------------------", 1)
'                Call KS7500i.PrintString("(서명)", 1)
'                Call KS7500i.PrintImage(m_SignFileName, 5)
'
'                KS7500i.LineFeed (1)
'
'                KS7500i.CutPaper
'            Next i
'
'            ADORs.MoveNext
'        Loop
'    End If
'    ADORs.Close
'    Set ADORs = Nothing
'
'    '--------------------------------------------------------------------
'    ' 2. 현금영수증 취소 내역 출력
'    '--------------------------------------------------------------------
'    Query = "SELECT * FROM TB_현금영수증"
'    Query = Query & " WHERE 고객코드 = '" & 고객코드 & "'"
'    Query = Query & "   AND 승인일자 = '" & Format(Date, "YYMMDD") & "'"
'
'    If 접수번호 = 0 Then
'        Query = Query & "   AND 접수번호 = 0"
'    Else
'        Query = Query & "   AND 접수번호 = " & 접수번호
'    End If
'
'    Query = Query & "   AND SUBSTRING(메시지2,1,2) = '취소'"
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    If Not ADORs.EOF Then
'        Do Until ADORs.EOF
'            For i = 1 To IIf(가맹점정보.가맹점카드영수증출력YN = "Y", 2, 1)
'                If i = 1 Then
'                    Call KS7500i.PrintString("현금영수증취소(고객용)", 2)
'                Else
'                    Call KS7500i.PrintString("현금영수증취소(가맹점용)", 2)
'                End If
'
'                KS7500i.LineFeed (1)
'
'                Call KS7500i.PrintString("-----------------------------------------------", 1)
'
'                '---------------------------------------------------------------------
'                ' TB_기본정보
'                '---------------------------------------------------------------------
'                Query = "SELECT    ISNULL(가맹점명, '')     AS 가맹점명"
'                Query = Query & ", ISNULL(사업자번호, '')   AS 사업자번호"
'                Query = Query & ", ISNULL(대표자명, '')     AS 대표자명"
'                Query = Query & ", ISNULL(매장전화번호, '') AS 매장전화번호"
'                Query = Query & ", ISNULL(사업장주소, '')   AS 사업장주소"
'                Query = Query & " FROM TB_기본정보"
'                Set SUBRs = New ADODB.Recordset
'                SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'                If SUBRs.EOF Then
'                    Call KS7500i.PrintString("상 호 명 : ", 1)
'                    Call KS7500i.PrintString("사업자No : ", 1)
'                    Call KS7500i.PrintString("대 표 자 : ", 1)
'                    Call KS7500i.PrintString("전화번호 : ", 1)
'                    Call KS7500i.PrintString("주    소 : ", 1)
'
'                Else
'                    Call KS7500i.PrintString("상 호 명 : " + SUBRs!가맹점명, 1)
'                    Call KS7500i.PrintString("사업자No : " + SUBRs!사업자번호, 1)
'                    Call KS7500i.PrintString("대 표 자 : " + SUBRs!대표자명, 1)
'                    Call KS7500i.PrintString("전화번호 : " + SUBRs!매장전화번호, 1)
'                    Call KS7500i.PrintString("주    소 : " + SUBRs!사업장주소, 1)
'                End If
'                SUBRs.Close
'                Set SUBRs = Nothing
'
'                Call KS7500i.PrintString("-----------------------------------------------", 1)
'
'                카드번호 = Trim(ADORs!사용자정보) & "" '사용자정보
'
'                If Len(카드번호) = 13 Then
'                    '주민번호
'                    카드번호 = Left(카드번호, Len(카드번호) - 7) & "*******"
'
'                ElseIf Len(카드번호) = 19 Then
'                    '현금영수증 카드
'                    카드번호 = Left(카드번호, Len(카드번호) - 7) & "******"
'                Else
'                    '휴대폰
'                    If Len(카드번호) > 4 Then
'                        카드번호 = Left(카드번호, Len(카드번호) - 4) & "****"
'                    End If
'                End If
'
'                Call KS7500i.PrintString("구 매 자 : " + 카드번호, 1)
'
'                If ADORs!소득구분 = "0" Then
'                    Call KS7500i.PrintString("소득구분 : 소득공제", 1)
'                Else
'                    Call KS7500i.PrintString("소득구분 : 비소득공제", 1)
'                End If
'
'                Call KS7500i.PrintString("승인일자 : " + Format(ADORs!승인일자, "2000/00/00") + " " + Format(ADORs!승인시간, "00:00"), 1)
'                Call KS7500i.PrintString("승인번호 : " + ADORs!승인번호, 1)
'                Call KS7500i.PrintString("승인금액 : " + CStr(ADORs!총금액), 2)
'                Call KS7500i.PrintString("-----------------------------------------------", 1)
'                Call KS7500i.PrintString(ADORs!국세청1, 1)
'                Call KS7500i.PrintString(ADORs!국세청2, 1)
'
'                KS7500i.LineFeed (1)
'
'                KS7500i.CutPaper
'            Next i
'
'            ADORs.MoveNext
'        Loop
'    End If
'    ADORs.Close
'    Set ADORs = Nothing
'
'    Exit Sub
'
'ErrRtn:
'    Call Error_Msg("", Err.Source, Err.Number, Err.description)
'
'    Screen.MousePointer = 0
'End Sub

Public Function Check_KS7500() As Boolean
    Query = "SELECT * FROM TB_기본정보"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    If ADORs.EOF Then
        Check_KS7500 = False
    Else
        If (Trim(ADORs!사업자번호) = "") Or (Trim(ADORs!단말기번호) = "") Or (Trim(ADORs!VAN_IP) = "") Or (Trim(ADORs!VAN_PORT) = "") Then
            Check_KS7500 = False
            
            Exit Function
        End If
    End If
    ADORs.Close
    Set ADORs = Nothing

    Check_KS7500 = True
    
    Exit Function
    
ErrRtn:
    Check_KS7500 = False
End Function

Public Function PhoneNo_Asterisk(전화번호 As String) As String
    Dim PhoneLen As Integer
    Dim iPos     As Long
    
    PhoneLen = Len(전화번호)
    
    Select Case Len(전화번호)
        Case Is <= 4
            PhoneNo_Asterisk = 전화번호 & ""
            
        Case 5 To 7
            PhoneNo_Asterisk = String(PhoneLen - 4, "*") & Right(전화번호, 4) & ""
            
        Case 8 To 10
            iPos = InStr(1, 전화번호, "-")
        
            If iPos > 0 Then
                If iPos < 4 Then
                    PhoneNo_Asterisk = String(PhoneLen - 4, "*") & Right(전화번호, 4) & ""
                Else
                    PhoneNo_Asterisk = String(PhoneLen - (iPos - 1), "*") & Mid(전화번호, iPos + 1, PhoneLen - iPos) & ""
                End If
            Else
                PhoneNo_Asterisk = String(PhoneLen - 4, "*") & Right(전화번호, 4) & ""
            End If
        
        Case 11
            PhoneNo_Asterisk = Left(전화번호, 3) & String(4, "*") & Right(전화번호, 4) & ""
        
        Case 12
            PhoneNo_Asterisk = Left(전화번호, 4) & String(3, "*") & Right(전화번호, 5) & ""
            
        Case 13
            PhoneNo_Asterisk = Left(전화번호, 4) & String(4, "*") & Right(전화번호, 5) & ""
        
        Case Else
            PhoneNo_Asterisk = String(PhoneLen - 4, "*") & Right(전화번호, 4) & ""
    End Select
End Function


Public Sub 신용카드재발행_Report(iRow As Long, strTitle As String, sGubun As Integer)
    On Error GoTo ErrRtn
    
   
    Dim tmp      As String
    Dim 이전미수 As String
    Dim 접수수량 As Integer
    Dim 접수금액 As String
    
    Dim 현금결제 As String
    Dim 카드결제 As Double
    
    Dim 마일리지 As String
    Dim 카드번호 As String
    
    Dim 받은금액 As String
    Dim 거스름돈 As String
    
    Dim 전화번호     As String
        
    Dim 거래일시 As String
    Dim Print_Msg As String

   
    If strTitle = "승인" Then
        strTitle = "신용승인"
    Else
        strTitle = "신용승인취소"
    End If
    
    With frm신용카드승인.sprGrid
        .Row = iRow
        
        If sGubun = 1 Then
            Print_Msg = Print_Msg & PrintString(strTitle & "(가맹점용)", 4, True)
        Else
            Print_Msg = Print_Msg & PrintString(strTitle & "(고객용)", 4, True)
        End If
        
        Print_Msg = Print_Msg & PrintString("(재발행)", 1, True)
        Print_Msg = Print_Msg & PrintString("-----------------------------------------------", 1, True)
        
        Query = "SELECT    ISNULL(가맹점명, '')     AS 가맹점명"
        Query = Query & ", ISNULL(사업자번호, '')   AS 사업자번호"
        Query = Query & ", ISNULL(대표자명, '')     AS 대표자명"
        Query = Query & ", ISNULL(매장전화번호, '') AS 매장전화번호"
        Query = Query & ", ISNULL(사업장주소, '')   AS 사업장주소"
        Query = Query & " FROM TB_기본정보"
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        If ADORs.EOF Then
            Print_Msg = Print_Msg & PrintString("상 호 명 : ", 1, True)
            Print_Msg = Print_Msg & PrintString("사업자No : ", 1, True)
            Print_Msg = Print_Msg & PrintString("대 표 자 : ", 1, True)
            Print_Msg = Print_Msg & PrintString("전화번호 : ", 1, True)
            Print_Msg = Print_Msg & PrintString("주    소 : ", 1, True)
        Else
            Print_Msg = Print_Msg & PrintString("상 호 명 : " + ADORs!가맹점명, 1, True)
            Print_Msg = Print_Msg & PrintString("사업자No : " + ADORs!사업자번호, 1, True)
            Print_Msg = Print_Msg & PrintString("대 표 자 : " + ADORs!대표자명, 1, True)
            Print_Msg = Print_Msg & PrintString("전화번호 : " + ADORs!매장전화번호, 1, True)
            Print_Msg = Print_Msg & PrintString("주    소 : " + ADORs!사업장주소, 1, True)
        End If
        ADORs.Close
        Set ADORs = Nothing
        
        Print_Msg = Print_Msg & PrintString("-----------------------------------------------", 1)
        
        .Col = 8: 거래일시 = "거래일시 : " & Format(.Text, "2000/00/00") & " "
        .Col = 9: 거래일시 = 거래일시 & Format(.Text, "00:00")
        
        Print_Msg = Print_Msg & PrintString(거래일시, 1, True)
        
        '            .Col = 16: 카드번호 = Left(.Text, 4) & "-"
        '                       카드번호 = 카드번호 & Mid(.Text, 5, 4) & "-"
        '                       카드번호 = 카드번호 & "****-"
        '                       카드번호 = 카드번호 & Mid(.Text, 13, 4)
        
        .Col = 16: 카드번호 = Left(.Text, 19)
        
        Print_Msg = Print_Msg & PrintString("카드번호 : " + 카드번호, 1, True)
        
        .Col = 13: Print_Msg = Print_Msg & PrintString("카드종류 : " + .Text, 1, True)
        .Col = 7:  Print_Msg = Print_Msg & PrintString("승인번호 : " + .Text, 1, True)
        
        .Col = 10
        If .Text = "00" Then
        Print_Msg = Print_Msg & PrintString("할부개월 : 일시불", 1, True)
        Else
        Print_Msg = Print_Msg & PrintString("할부개월 : " + CStr(.Value) + " 개월", 1, True)
        End If
        
        .Col = 11: 카드결제 = CDbl(.Text)
        
        Print_Msg = Print_Msg & PrintString("과세금액 : " + Format(카드결제 - (카드결제 - (카드결제 / 1.1)), "#,##0") + "원", 1, True)
        Print_Msg = Print_Msg & PrintString("부가세액 : " + Format(카드결제 - (카드결제 / 1.1), "#,##0") + "원", 1, True)
        Print_Msg = Print_Msg & PrintString("승인금액 : " + Format(카드결제, "#,##0") + "원", 4, True)
        Print_Msg = Print_Msg & PrintString("-----------------------------------------------", 1)
        
        Print_Msg = Print_Msg & PrintLineFeed(4)
        Print_Msg = Print_Msg & PrintCut
        
        
        
    End With
    
    Call frmKicc.Card_Print(Print_Msg)
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Public Sub 현금영수증재발행_Report(iRow As Long, strTitle As String, sGubun As Integer)
    On Error GoTo ErrRtn

    
    Dim Print_Msg As String
    Dim tmp      As String
    Dim 이전미수 As String
    Dim 접수수량 As Integer
    Dim 접수금액 As String

    Dim 현금결제 As String
    Dim 카드결제 As Double


    Dim 마일리지 As String
    Dim 카드번호 As String

    Dim 받은금액 As String
    Dim 거스름돈 As String

    Dim 전화번호     As String

    Dim 승인일자 As String

    
    If strTitle = "승인" Then
        strTitle = "현금영수증"
    Else
        strTitle = "현금영수증취소"
    End If



    With frm현금영수증승인.sprGrid
        .Row = iRow

        If sGubun = 1 Then
            Print_Msg = Print_Msg & PrintString(strTitle & "(가맹점용)", 4, True)
        Else
            Print_Msg = Print_Msg & PrintString(strTitle & "(고객용)", 4, True)
        End If

        Print_Msg = Print_Msg & PrintString("(재발행)", 1, True)
        Print_Msg = Print_Msg & PrintString("-----------------------------------------------", 1)

        Query = "SELECT    ISNULL(가맹점명, '')     AS 가맹점명"
        Query = Query & ", ISNULL(사업자번호, '')   AS 사업자번호"
        Query = Query & ", ISNULL(대표자명, '')     AS 대표자명"
        Query = Query & ", ISNULL(매장전화번호, '') AS 매장전화번호"
        Query = Query & ", ISNULL(사업장주소, '')   AS 사업장주소"
        Query = Query & " FROM TB_기본정보"
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

        If ADORs.EOF Then
            Print_Msg = Print_Msg & PrintString("상 호 명 : ", 1, True)
            Print_Msg = Print_Msg & PrintString("사업자No : ", 1, True)
            Print_Msg = Print_Msg & PrintString("대 표 자 : ", 1, True)
            Print_Msg = Print_Msg & PrintString("전화번호 : ", 1, True)
            Print_Msg = Print_Msg & PrintString("주    소 : ", 1, True)

        Else
            Print_Msg = Print_Msg & PrintString("상 호 명 : " + ADORs!가맹점명, 1, True)
            Print_Msg = Print_Msg & PrintString("사업자No : " + ADORs!사업자번호, 1, True)
            Print_Msg = Print_Msg & PrintString("대 표 자 : " + ADORs!대표자명, 1, True)
            Print_Msg = Print_Msg & PrintString("전화번호 : " + ADORs!매장전화번호, 1, True)
            Print_Msg = Print_Msg & PrintString("주    소 : " + ADORs!사업장주소, 1, True)
        End If
        ADORs.Close
        Set ADORs = Nothing

        Print_Msg = Print_Msg & PrintString("-----------------------------------------------", 1)

        .Col = 13: 카드번호 = Trim(.Text) & "" '사용자정보

        If Len(카드번호) = 13 Then
            '주민번호
            카드번호 = Left(카드번호, Len(카드번호) - 7) & "*******"

        ElseIf Len(카드번호) = 19 Then
            '현금영수증 카드
            카드번호 = Left(카드번호, Len(카드번호) - 7) & "******"
        Else
            '휴대폰
            If Len(카드번호) > 4 Then
                카드번호 = Left(카드번호, Len(카드번호) - 4) & "****"
            End If
        End If

        Print_Msg = Print_Msg & PrintString("구 매 자 : " + 카드번호, 1, True)

        .Col = 11
        ' KS7500,7050     모듈은 0, 소득공제용, 1.지출증빙용
        ' KS4060 보안인증 모듈은 1, 소득공제용, 2.지출증빙용
        If 가맹점정보.CAT단말기종류 <> "KS4060 보안인증" Then
            Print_Msg = Print_Msg & PrintString("거래구분 : " & IIf(.Text = "0", "소득공제용", "지출증빙용"), 1, True)
        Else
            Print_Msg = Print_Msg & PrintString("거래구분 : " & IIf(.Text = "1", "소득공제용", "지출증빙용"), 1, True)
        End If

        .Col = 8: 승인일자 = "승인일자 : " & Format(.Text, "2000/00/00") + " "
        .Col = 9: 승인일자 = 승인일자 & Format(.Text, "00:00")

        Print_Msg = Print_Msg & PrintString(승인일자, 1, True)

        .Col = 7: Print_Msg = Print_Msg & PrintString("승인번호 : " + .Text, 1, True)
'            .Col = 12: Call PrintString("승인금액 : " + .Text, 1)

        .Col = 12: 카드결제 = CDbl(.Text)

        Print_Msg = Print_Msg & PrintString("과세금액 : " + Format(카드결제 - (카드결제 - (카드결제 / 1.1)), "#,##0") + "원", 1, True)
        Print_Msg = Print_Msg & PrintString("부가세액 : " + Format(카드결제 - (카드결제 / 1.1), "#,##0") + "원", 1, True)
        Print_Msg = Print_Msg & PrintString("승인금액 : " + Format(카드결제, "#,##0") + "원", 4, True)

        Print_Msg = Print_Msg & PrintString("-----------------------------------------------", 1)
        .Col = 17: Print_Msg = Print_Msg & PrintString(.Text, 1)
        .Col = 18: Print_Msg = Print_Msg & PrintString(.Text, 1)

        Print_Msg = Print_Msg & PrintLineFeed(4)
        Print_Msg = Print_Msg & PrintCut
    End With
    
    Call frmKicc.Card_Print(Print_Msg)
    Exit Sub
    
ErrRtn:
    
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
    Screen.MousePointer = 0
End Sub

