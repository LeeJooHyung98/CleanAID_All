Attribute VB_Name = "basSale"
Option Explicit

Public Function Get_세탁금액(ClothCode As String, sGoodsStats As String, Optional Internet As String = "") As Long
    Dim ADORs       As ADODB.RecordSet
    
    On Error GoTo ErrRtn
    
    Set ADORs = New ADODB.RecordSet
    Set ADORs = Get_의류정보(ClothCode, sGoodsStats, Internet)
    
    If ADORs.EOF Then
        ADORs.Close:    Set ADORs = Nothing
        Get_세탁금액 = -1
        Exit Function
    End If
     
    Get_세탁금액 = ADORs!금액 & ""
    ADORs.Close:    Set ADORs = Nothing
    
    Exit Function
    
ErrRtn:
    Get_세탁금액 = -1
End Function


Public Function Get_세탁금액_20130423이전(ClothCode As String) As Long
    Dim 요일 As String
        
    On Error GoTo ErrRtn
    
    
    '--------------------------------------------------------
    ' TB_할인정보
    '--------------------------------------------------------
    Query = "SELECT TOP 1 시작일자"
    Query = Query & ", ISNULL(할인금액,0) AS 할인금액"
    Query = Query & " FROM TB_할인정보"
    Query = Query & " WHERE 시작일자 <= '" & Format(Date, "YYYY-MM-DD") & "'"
    Query = Query & "   AND 종료일자 >= '" & Format(Date, "YYYY-MM-DD") & "'"
    Query = Query & "   AND 의류코드  = '" & ClothCode & "'"
    Query = Query & " ORDER BY 시작일자 DESC, 종료일자 ASC "
    Set Rs = New ADODB.RecordSet
    Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If Rs.EOF Then
        Rs.Close
        Set Rs = Nothing
    Else
        Get_세탁금액_20130423이전 = Rs!할인금액 & ""
        
        Rs.Close:   Set Rs = Nothing
        Exit Function
    End If
        
    요일 = Weekday(Date)
    If chkDaySale = True Then
        '--------------------------------------------------------
        ' TB_요일할인
        '--------------------------------------------------------
        Query = "SELECT  TOP 1  시작일자"
        Query = Query & ", ISNULL(할인금액,0) AS 할인금액"
        Query = Query & " FROM TB_요일할인"
        Query = Query & " WHERE 시작일자 <= '" & Format(Date, "YYYY-MM-DD") & "'"
        Query = Query & "   AND 종료일자 >= '" & Format(Date, "YYYY-MM-DD") & "'"
        Query = Query & "   AND 요일      = '" & 요일 & "'"
        Query = Query & "   AND 의류코드  = '" & ClothCode & "'"
        Query = Query & " ORDER BY 시작일자 DESC, 종료일자 ASC "
        Set Rs = New ADODB.RecordSet
        Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        If Rs.EOF Then
            Rs.Close
            Set Rs = Nothing
        Else
            Get_세탁금액_20130423이전 = Rs!할인금액 & ""
            
            Rs.Close:   Set Rs = Nothing
            Exit Function
        End If
    End If
    

    '--------------------------------------------------------
    ' TB_의류
    '--------------------------------------------------------
    Query = "SELECT ISNULL(금액,0) FROM TB_의류"
    Query = Query & " WHERE 의류코드  = '" & ClothCode & "'"
    Set Rs = New ADODB.RecordSet
    Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If Rs.EOF Then
        Get_세탁금액_20130423이전 = -1
    Else
        Get_세탁금액_20130423이전 = Rs(0) & ""
    End If
    Rs.Close:   Set Rs = Nothing
    Exit Function
    
ErrRtn:
    Get_세탁금액_20130423이전 = -1
End Function

Public Function Get_세탁정상금액(ClothCode As String) As Long
    
    On Error GoTo ErrRtn
    
    '--------------------------------------------------------
    ' TB_의류
    '--------------------------------------------------------
    Query = "SELECT ISNULL(금액,0) FROM TB_의류"
    Query = Query & " WHERE 의류코드  = '" & ClothCode & "'"
    Set Rs = New ADODB.RecordSet
    Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If Rs.EOF Then
        Get_세탁정상금액 = -1
    Else
        Get_세탁정상금액 = Rs(0) & ""
    End If
    Rs.Close:   Set Rs = Nothing
    Exit Function
    
ErrRtn:
    Get_세탁정상금액 = -1
End Function


Public Function Get_의류정보(ClothCode As String, sGoodsStats As String, Optional Internet As String = "") As ADODB.RecordSet
    Dim ADORs1  As ADODB.RecordSet      ' 행사
    Dim ADORs2  As ADODB.RecordSet      ' 요일
    Dim ADORs3  As ADODB.RecordSet      ' 일반
    Dim bChekc(3) As Boolean
    Dim 요일     As String
    Dim Data_EOF As Boolean
    
    Dim 시작일자 As String
    
    ' 행사 -> 요일가격 -> 일반 가격 순으로 처리한다.
    ' 행사,요일이 중복될 경우 낮은 가격으로 처리한다.
    ' 해당 일자에서 시작일자 및 종료 일자가 빠른것을 우선 처리한다.
    
    Set ADORs1 = New ADODB.RecordSet
    Set ADORs2 = New ADODB.RecordSet
    Set ADORs3 = New ADODB.RecordSet
    
    If Internet <> "" Then
        ' 해당 조건이 아닐 경우 일반 금액을 적용
        '---------------------------------------------------------------------------
        ' TB_의류
        '---------------------------------------------------------------------------
        Query = "SELECT    의류명, 금액, 의류코드, 순서 FROM TB_의류_인터넷"
        If Len(ClothCode) = 2 Then
            Query = Query & " WHERE SUBSTRING(의류코드,1,2) = '" & Left(ClothCode, 2) & "'"
        Else
            Query = Query & " WHERE 의류코드 = '" & ClothCode & "'"
        End If
        Query = Query & " ORDER BY 순서 ASC, 의류코드 ASC"
        
        ADORs3.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        Set Get_의류정보 = New ADODB.RecordSet
        Set Get_의류정보 = ADORs3
        sGoodsStats = "인터넷 가격 적용"
        
        Set ADORs1 = Nothing:   Set ADORs2 = Nothing:   Set ADORs3 = Nothing
        Exit Function
    End If
    

    
    
    bChekc(1) = False: bChekc(2) = False: bChekc(3) = False
    
    
    
    Data_EOF = True  ' 행사, 요일 자료가 없는 것으로 기본 값을 가진다.
    시작일자 = "2000-01-01"
    '--------------------------------------------------------------------------
    ' TB_할인정보 - 시작일자
    '--------------------------------------------------------------------------
            Query = "SELECT ISNULL(MAX(시작일자), '2000-01-01') AS 시작일자"
    Query = Query & " FROM TB_할인정보"
    Query = Query & " WHERE  '" & Format(Date, "YYYY-MM-DD") & "' BETWEEN 시작일자 AND 종료일자 "
    ADORs1.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If Not ADORs1.EOF Then 시작일자 = ADORs1!시작일자 & ""
    ADORs1.Close

    '---------------------------------------------------------------------------
    ' TB_할인정보 - 시작일자는 같으면서 종료일자가 틀린 할인정보가 저장된 경우
    '               종료일자가 빠른 것으로 처리한다.
    '---------------------------------------------------------------------------
    Query = "SELECT    A.의류명"
    Query = Query & ", A.할인금액 AS 금액"
    Query = Query & ", A.의류코드"
    Query = Query & ", A.순서"
    Query = Query & " FROM TB_할인정보 AS A RIGHT OUTER JOIN ("
    Query = Query & "                                        SELECT 의류코드"
    Query = Query & "                                             , 시작일자"
    Query = Query & "                                             , MIN(종료일자) AS 종료일자"
    Query = Query & "                                        FROM TB_할인정보"
    Query = Query & "                                        WHERE 시작일자 = '" & Format(시작일자, "YYYY-MM-DD") & "'"
    Query = Query & "                                        GROUP BY 의류코드, 시작일자"
    Query = Query & "                                       ) AS B ON A.의류코드 = B.의류코드"
    Query = Query & "                                             AND A.시작일자 = B.시작일자"
    Query = Query & "                                             AND A.종료일자 = B.종료일자"
    Query = Query & " WHERE '" & Format(Date, "YYYY-MM-DD") & "' BETWEEN A.시작일자 AND  A.종료일자 "
    
    If Len(ClothCode) = 2 Then
        Query = Query & "   AND SUBSTRING(A.의류코드,1,2) = '" & Left(ClothCode, 2) & "'"
    Else
        Query = Query & "   AND A.의류코드 = '" & ClothCode & "'"
    End If
    
    Query = Query & " ORDER BY A.순서 ASC, A.의류코드 ASC"
    ADORs1.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    ' 해당 자료 여부
    If Not ADORs1.EOF Then bChekc(1) = True
    
    ' 요일을 확인한다.
    If chkDaySale = True Then
        요일 = Weekday(Date)
        
        '--------------------------------------------------------------------------
        ' TB_요일할인 - 시작일자
        '--------------------------------------------------------------------------
                Query = "SELECT ISNULL(MAX(시작일자), '2000-01-01') AS 시작일자"
        Query = Query & " FROM TB_요일할인"
        Query = Query & " WHERE  '" & Format(Date, "YYYY-MM-DD") & "' BETWEEN 시작일자 AND 종료일자 "
        Query = Query & "   AND 요일       = '" & 요일 & "'"
        Set ADORs2 = New ADODB.RecordSet
        ADORs2.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        If ADORs2.EOF Then
            시작일자 = "2000-01-01"
        Else
            시작일자 = ADORs2!시작일자 & ""
        End If
        ADORs2.Close
        
        '---------------------------------------------------------------------------
        ' TB_요일할인
        '---------------------------------------------------------------------------
        Query = "SELECT    A.의류명"
        Query = Query & ", A.할인금액 AS 금액"
        Query = Query & ", A.의류코드"
        Query = Query & ", A.순서"
        Query = Query & " FROM TB_요일할인 AS A RIGHT OUTER JOIN ("
        Query = Query & "                                        SELECT 의류코드"
        Query = Query & "                                             , 시작일자"
        Query = Query & "                                             , MIN(종료일자) AS 종료일자"
        Query = Query & "                                        FROM TB_요일할인"
        Query = Query & "                                        WHERE 시작일자 = '" & Format(시작일자, "YYYY-MM-DD") & "'"
        Query = Query & "                                          AND 요일     = '" & 요일 & "'"
        Query = Query & "                                        GROUP BY 의류코드, 시작일자"
        Query = Query & "                                       ) AS B ON A.의류코드 = B.의류코드"
        Query = Query & "                                             AND A.시작일자 = B.시작일자"
        Query = Query & "                                             AND A.종료일자 = B.종료일자"
        Query = Query & " WHERE '" & Format(Date, "YYYY-MM-DD") & "' BETWEEN A.시작일자 AND  A.종료일자 "
        Query = Query & "   AND A.요일       = '" & 요일 & "'"
        
        If Len(ClothCode) = 2 Then
            Query = Query & "   AND SUBSTRING(A.의류코드,1,2) = '" & ClothCode & "'"
        Else
            Query = Query & "   AND A.의류코드 = '" & ClothCode & "'"
        End If
        
        Query = Query & " ORDER BY A.순서 ASC, A.의류코드 ASC"
        ADORs2.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        ' 해당 자료 여부
        If Not ADORs2.EOF Then bChekc(2) = True
    
    End If
    
    '행사적용: 행사가 있고 요일이 아닐경우
    If bChekc(1) And bChekc(2) = False Then
        Set Get_의류정보 = New ADODB.RecordSet
        Set Get_의류정보 = ADORs1
        sGoodsStats = "행사가격 적용"
        
        Set ADORs1 = Nothing:   Set ADORs2 = Nothing:   Set ADORs3 = Nothing
        Exit Function
        
    '요일행사적용:  행사가 없고 요일일 경우
    ElseIf bChekc(1) = False And bChekc(2) = True Then
        ' 요일 행사는 체크가 되어 있으나 해당 일자의 가격이 없을 경우 기본 가격으로 한다.
        If Not ADORs2.EOF Then
            Set Get_의류정보 = New ADODB.RecordSet
            Set Get_의류정보 = ADORs2
            sGoodsStats = "요일행사가격 적용"
        
            Set ADORs1 = Nothing:   Set ADORs2 = Nothing:   Set ADORs3 = Nothing
            Exit Function
        End If
    
    '행사,요일중 적은 금액일 경우 적용
    ElseIf bChekc(1) And bChekc(2) Then
        If Not ADORs1.EOF And Not ADORs2.EOF Then
            Set Get_의류정보 = New ADODB.RecordSet
            ' 행사 금액이 적거나 같은경우
            If ADORs1.Fields("금액") <= ADORs2.Fields("금액") Then
                Set Get_의류정보 = ADORs1
                sGoodsStats = "행사가격 적용"
            
            ' 요일 금액이 적은 경우
            ElseIf ADORs1.Fields("금액") > ADORs2.Fields("금액") Then
                Set Get_의류정보 = ADORs2
                sGoodsStats = "요일행사가격 적용"
            End If
        
            Set ADORs1 = Nothing:   Set ADORs2 = Nothing:   Set ADORs3 = Nothing
            Exit Function
        End If
    End If
        
        
    ' 해당 조건이 아닐 경우 일반 금액을 적용
    '---------------------------------------------------------------------------
    ' TB_의류
    '---------------------------------------------------------------------------
    Query = "SELECT    의류명, 금액, 의류코드, 순서 FROM TB_의류"
    If Len(ClothCode) = 2 Then
        Query = Query & " WHERE SUBSTRING(의류코드,1,2) = '" & Left(ClothCode, 2) & "'"
    Else
        Query = Query & " WHERE 의류코드 = '" & ClothCode & "'"
    End If
    Query = Query & " ORDER BY 순서 ASC, 의류코드 ASC"
    
    ADORs3.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    Set Get_의류정보 = New ADODB.RecordSet
    Set Get_의류정보 = ADORs3
    sGoodsStats = "기본가격 적용"
    
    Set ADORs1 = Nothing:   Set ADORs2 = Nothing:   Set ADORs3 = Nothing
    Exit Function

ErrRtn:
    Set Get_의류정보 = Nothing
    Set ADORs1 = Nothing:   Set ADORs2 = Nothing:   Set ADORs3 = Nothing

End Function



Public Function Get_의류정보_20130423이전(ClothCode As String) As Boolean
    Dim 요일     As String
    Dim Data_EOF As Boolean
    
    Dim 시작일자 As String
    
    ' 행사 -> 요일가격 -> 일반 가격 순으로 처리한다.
    ' 해당 일자에서 시작일자 및 종료 일자가 빠른것을 우선 처리한다.
    
    Data_EOF = True  ' 행사, 요일 자료가 없는 것으로 기본 값을 가진다.
    '--------------------------------------------------------------------------
    ' TB_할인정보 - 시작일자
    '--------------------------------------------------------------------------
            Query = "SELECT ISNULL(MAX(시작일자), '2000-01-01') AS 시작일자"
    Query = Query & " FROM TB_할인정보"
    Query = Query & " WHERE (시작일자 <= '" & Format(Date, "YYYY-MM-DD") & "'"
    Query = Query & "   AND  종료일자 >= '" & Format(Date, "YYYY-MM-DD") & "')"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF Then
        시작일자 = "2000-01-01"
    Else
        시작일자 = ADORs!시작일자 & ""
    End If
    ADORs.Close:    Set ADORs = Nothing

    '---------------------------------------------------------------------------
    ' TB_할인정보 - 시작일자는 같으면서 종료일자가 틀린 할인정보가 저장된 경우
    '               종료일자가 빠른 것으로 처리한다.
    '---------------------------------------------------------------------------
    Query = "SELECT    A.의류명"
    Query = Query & ", A.할인금액 AS 금액"
    Query = Query & ", A.의류코드"
    Query = Query & ", A.순서"
    Query = Query & " FROM TB_할인정보 AS A RIGHT OUTER JOIN ("
    Query = Query & "                                        SELECT 의류코드"
    Query = Query & "                                             , 시작일자"
    Query = Query & "                                             , MIN(종료일자) AS 종료일자"
    Query = Query & "                                        FROM TB_할인정보"
    Query = Query & "                                        WHERE 시작일자 = '" & Format(시작일자, "YYYY-MM-DD") & "'"
    Query = Query & "                                        GROUP BY 의류코드, 시작일자"
    Query = Query & "                                       ) AS B ON A.의류코드 = B.의류코드"
    Query = Query & "                                             AND A.시작일자 = B.시작일자"
    Query = Query & "                                             AND A.종료일자 = B.종료일자"
    Query = Query & " WHERE (A.시작일자 <= '" & Format(Date, "YYYY-MM-DD") & "'"
    Query = Query & "   AND  A.종료일자 >= '" & Format(Date, "YYYY-MM-DD") & "')"
    
    If Len(ClothCode) = 2 Then
        Query = Query & "   AND SUBSTRING(A.의류코드,1,2) = '" & Left(ClothCode, 2) & "'"
    Else
        Query = Query & "   AND A.의류코드 = '" & ClothCode & "'"
    End If
    
    Query = Query & " ORDER BY A.순서 ASC, A.의류코드 ASC"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
                        
    If ADORs.EOF Then
        ADORs.Close:    Set ADORs = Nothing
        
        If chkDaySale = True Then
            요일 = Weekday(Date)
            
            '--------------------------------------------------------------------------
            ' TB_요일할인 - 시작일자
            '--------------------------------------------------------------------------
                    Query = "SELECT ISNULL(MAX(시작일자), '2000-01-01') AS 시작일자"
            Query = Query & " FROM TB_요일할인"
            Query = Query & " WHERE (시작일자 <= '" & Format(Date, "YYYY-MM-DD") & "'"
            Query = Query & "   AND  종료일자 >= '" & Format(Date, "YYYY-MM-DD") & "')"
            Query = Query & "   AND 요일       = '" & 요일 & "'"
            Set ADORs = New ADODB.RecordSet
            ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            If ADORs.EOF Then
                시작일자 = "2000-01-01"
            Else
                시작일자 = ADORs!시작일자 & ""
            End If
            ADORs.Close
            Set ADORs = Nothing
            
            '---------------------------------------------------------------------------
            ' TB_요일할인
            '---------------------------------------------------------------------------
            Query = "SELECT    A.의류명"
            Query = Query & ", A.할인금액 AS 금액"
            Query = Query & ", A.의류코드"
            Query = Query & ", A.순서"
            Query = Query & " FROM TB_요일할인 AS A RIGHT OUTER JOIN ("
            Query = Query & "                                        SELECT 의류코드"
            Query = Query & "                                             , 시작일자"
            Query = Query & "                                             , MIN(종료일자) AS 종료일자"
            Query = Query & "                                        FROM TB_요일할인"
            Query = Query & "                                        WHERE 시작일자 = '" & Format(시작일자, "YYYY-MM-DD") & "'"
            Query = Query & "                                          AND 요일     = '" & 요일 & "'"
            Query = Query & "                                        GROUP BY 의류코드, 시작일자"
            Query = Query & "                                       ) AS B ON A.의류코드 = B.의류코드"
            Query = Query & "                                             AND A.시작일자 = B.시작일자"
            Query = Query & "                                             AND A.종료일자 = B.종료일자"
            Query = Query & " WHERE (A.시작일자 <= '" & Format(Date, "YYYY-MM-DD") & "'"
            Query = Query & "   AND  A.종료일자 >= '" & Format(Date, "YYYY-MM-DD") & "')"
            Query = Query & "   AND A.요일       = '" & 요일 & "'"
            
            If Len(ClothCode) = 2 Then
                Query = Query & "   AND SUBSTRING(A.의류코드,1,2) = '" & ClothCode & "'"
            Else
                Query = Query & "   AND A.의류코드 = '" & ClothCode & "'"
            End If
            
            Query = Query & " ORDER BY A.순서 ASC, A.의류코드 ASC"
            Set ADORs = New ADODB.RecordSet
            ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
                    
            If ADORs.EOF Then
                ADORs.Close
                Set ADORs = Nothing
                            
                Data_EOF = True
            Else
                ' 요일 행사가 있을 경우
                Data_EOF = False
            End If
        End If
    Else
        ' 행사 자료가 있을 경우
        Data_EOF = False
    End If
        
    'TB_요일할인 - 데이터가 없는 경우
    If Data_EOF = True Then

        '---------------------------------------------------------------------------
        ' TB_의류
        '---------------------------------------------------------------------------
        Query = "SELECT    의류명"
        Query = Query & ", 금액"
        Query = Query & ", 의류코드"
        Query = Query & ", 순서"
        Query = Query & " FROM TB_의류"
        
        If Len(ClothCode) = 2 Then
            Query = Query & " WHERE SUBSTRING(의류코드,1,2) = '" & Left(ClothCode, 2) & "'"
        Else
            Query = Query & " WHERE 의류코드 = '" & ClothCode & "'"
        End If
        
        Query = Query & " ORDER BY 순서 ASC, 의류코드 ASC"
        Set ADORs = New ADODB.RecordSet
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        If ADORs.EOF Then
            ADORs.Close
            Set ADORs = Nothing
            
            Get_의류정보_20130423이전 = False
            Exit Function
        End If
    End If
        
    Get_의류정보_20130423이전 = True
    Exit Function
    
ErrRtn:
    Get_의류정보_20130423이전 = False
End Function

'  이마트 할인
'> 기간은 2011-10-27 ~ 2011-11-30
Public Function Action_지정할인_코드확인(sCode As String) As String
    Dim bAction As String
    
    Dim varTemp     As Variant
    
    On Error GoTo Action_지정할인_코드확인_Error
    
    Select Case LCase(sCode)
        Case "m000", "v000"
            Action_지정할인_코드확인 = "A"
            Exit Function
    
        Case "a000" To "a999"
            Action_지정할인_코드확인 = "A"
            Exit Function
        
        Case "s000" To "s999"
            Action_지정할인_코드확인 = "A"
            Exit Function
        
        Case "l000" To "l999"
            Action_지정할인_코드확인 = "A"
            Exit Function
        
        Case "n000" To "n999"
            Action_지정할인_코드확인 = "A"
            Exit Function
        
        Case "o000" To "o999"
            Action_지정할인_코드확인 = "A"
            Exit Function
        
        Case "p000" To "p999"
            Action_지정할인_코드확인 = "A"
            Exit Function
        
        Case "w000" To "w999"
            Action_지정할인_코드확인 = "A"
            Exit Function
        
        Case "x000" To "x999"
            Action_지정할인_코드확인 = "A"
            Exit Function
            
    End Select
    
    On Error GoTo 0
    Exit Function

Action_지정할인_코드확인_Error:
    Action_지정할인_코드확인 = "Z"
    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure Action_지정할인_코드확인 of Module Sale"
End Function



Public Function 행사관련자료()
    
'    If Format(Date, "yyyy-MM-dd") >= "2011-10-27" And Format(Date, "yyyy-MM-dd") <= "2011-11-30" Then
'        Call E마트행사_20111027
'    End If
    

End Function

Public Function Get_반품수량(sSDate As String, sEDate As String, sCustCode As String) As Long
    Dim SSQL    As String
    Dim Any_Rs  As ADODB.RecordSet
    
    On Error GoTo ERR_RTN

    '-----------------------------------------------------------------------------------
    ' 반품수량
    '-----------------------------------------------------------------------------------
    Query = "SELECT    COUNT(택번호)     AS 수량"
    Query = Query & ", MAX(지사출고일자) AS 본사출고일"
    Query = Query & " FROM TB_입출고"
    Query = Query & " WHERE (접수일자 >= '" & sSDate & "'"
    Query = Query & "   AND  접수일자 <= '" & sEDate & "')"
    Query = Query & "   AND  고객코드   = '" & sCustCode & "'"
    Query = Query & "   AND  지사출고상태 = '2' "
    Query = Query & "   AND (출고일자 = '' OR 출고일자 IS NULL)"
    
    Set Any_Rs = New ADODB.RecordSet
    Any_Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If Any_Rs.EOF = False Then
        Get_반품수량 = Any_Rs!수량 & ""
    End If
    Any_Rs.Close:    Set Any_Rs = Nothing
    Exit Function
    
ERR_RTN:

End Function


Public Function Get_요일행사여부() As Boolean
    Dim SSQL    As String
    Dim Any_Rs  As ADODB.RecordSet
    
    On Error GoTo ERR_RTN

    '-------------------------------------------------------------
    ' 요일세일 - 일,월,화,수,목,금,토
    '-------------------------------------------------------------
    SSQL = "SELECT 요일할인 FROM TB_기본정보"
    Set Any_Rs = New ADODB.RecordSet
    Any_Rs.Open SSQL, ADOCon, adOpenForwardOnly, adLockReadOnly

    If Not Any_Rs.EOF Then
        i = Weekday(Date)
        
        If Mid(Any_Rs(0), i, 1) = "1" Then
            Get_요일행사여부 = True
        Else
            Get_요일행사여부 = False
        End If
    End If
    
    Any_Rs.Close:    Set Any_Rs = Nothing
    Exit Function
    
ERR_RTN:
    Get_요일행사여부 = False

End Function



