Attribute VB_Name = "Sale"
Option Explicit

'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'★★                                                                                                                      ★★
'★★                                                                                                                      ★★
'★★            크랜즈 갤러리 행사에서 제외 시킬것                                                                        ★★
'★★                                                                                                                      ★★
'★★                                                                                                                      ★★
'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★


Public Sub Store_Sale_Check()
    
    '----------------------------------------------------------------------------------------------------------------------
    ' 이마트 매장 할인
    If Format(Date, "YYYY-MM-DD") >= "2009-07-30" And Format(Date, "YYYY-MM-DD") <= "2009-08-05" Then
    
        If Check_이마트할인_20090730 = False Then
            MsgBox "2009-07-30 ~ 2009-08-05 할인 정보 설정 오류 " & Err.Description, vbInformation, "확인"
            End
        End If
        
        ' 대리점 정보를 다시 읽는다.
        If Fb대리점정보 = "Error" Then
            MsgBox "대리점 정보가 올바르지 않습니다. 확인후 저장하여 주십시요.", vbInformation, "확인"
            frmINIT.Show 1
            End
        End If
    
    End If
    '----------------------------------------------------------------------------------------------------------------------
    
    
    '----------------------------------------------------------------------------------------------------------------------
    ' 2009.10.29 - 2009.11.04 이마트 매장 할인
    If Format(Date, "YYYY-MM-DD") >= "2009-10-29" And Format(Date, "YYYY-MM-DD") <= "2009-11-04" Then
    
        If Check_이마트할인_20091029 = False Then
            MsgBox "2009-10-29 ~ 2009-11-04 할인 정보 설정 오류 " & Err.Description, vbInformation, "확인"
            End
        End If
        
        ' 대리점 정보를 다시 읽는다.
        If Fb대리점정보 = "Error" Then
            MsgBox "대리점 정보가 올바르지 않습니다. 확인후 저장하여 주십시요.", vbInformation, "확인"
            frmINIT.Show 1
            End
        End If
    
    End If
    '----------------------------------------------------------------------------------------------------------------------
    
    '----------------------------------------------------------------------------------------------------------------------
    ' 2009.11.09 - 2009.11.14 전매장 빼빼로데이 할인 행사 실시 (이마트 매장 제외)
    If Format(Date, "YYYY-MM-DD") >= "2009-11-09" And Format(Date, "YYYY-MM-DD") <= "2009-11-14" Then
    
        If Action_빼빼로행사_20091109 = False Then
            MsgBox "2009-10-29 ~ 2009-11-04 할인 정보 설정 오류 " & Err.Description, vbInformation, "확인"
            End
        End If
        
        ' 대리점 정보를 다시 읽는다.
        If Fb대리점정보 = "Error" Then
            MsgBox "대리점 정보가 올바르지 않습니다. 확인후 저장하여 주십시요.", vbInformation, "확인"
            frmINIT.Show 1
            End
        End If
    
    End If
    '----------------------------------------------------------------------------------------------------------------------
    
    '----------------------------------------------------------------------------------------------------------------------
    ' 2009.12.11 - 2009.12.31 전매장 세트 상품 행사 내용
    If Format(Date, "YYYY-MM-DD") >= "2009-12-11" And Format(Date, "YYYY-MM-DD") <= "2009-12-31" Then

        If Action_세트행사_20091211 = False Then
            MsgBox "2009-12-11 ~ 2009-12-31 할인 정보 설정 오류 " & Err.Description, vbInformation, "확인"
            End
        End If

        ' 대리점 정보를 다시 읽는다.
        If Fb대리점정보 = "Error" Then
            MsgBox "대리점 정보가 올바르지 않습니다. 확인후 저장하여 주십시요.", vbInformation, "확인"
            frmINIT.Show 1
            End
        End If

    End If
    '----------------------------------------------------------------------------------------------------------------------

    '----------------------------------------------------------------------------------------------------------------------
    ' 2009.12.13 - 2009.12.31 전매장 세트 상품 행사 내용  -- 매화주공점
    If Format(Date, "YYYY-MM-DD") >= "2009-12-13" And Format(Date, "YYYY-MM-DD") <= "2009-12-31" Then

        If Action_세트행사_20091213 = False Then
            MsgBox "2009-12-13 ~ 2009-12-31 할인 정보 설정 오류 " & Err.Description, vbInformation, "확인"
            End
        End If

        ' 대리점 정보를 다시 읽는다.
        If Fb대리점정보 = "Error" Then
            MsgBox "대리점 정보가 올바르지 않습니다. 확인후 저장하여 주십시요.", vbInformation, "확인"
            frmINIT.Show 1
            End
        End If

    End If
    '----------------------------------------------------------------------------------------------------------------------
    
    '----------------------------------------------------------------------------------------------------------------------
    ' 2009.12.16 - 2009.12.31 전매장 세트 상품 행사 내용  -- 매화주공점
    If Format(Date, "YYYY-MM-DD") >= "2009-12-16" And Format(Date, "YYYY-MM-DD") <= "2009-12-31" Then

        If Action_세트행사_20091216 = False Then
            MsgBox "2009-12-16 ~ 2009-12-31 할인 정보 설정 오류 " & Err.Description, vbInformation, "확인"
            End
        End If

        ' 대리점 정보를 다시 읽는다.
        If Fb대리점정보 = "Error" Then
            MsgBox "대리점 정보가 올바르지 않습니다. 확인후 저장하여 주십시요.", vbInformation, "확인"
            frmINIT.Show 1
            End
        End If

    End If
    '----------------------------------------------------------------------------------------------------------------------

    '----------------------------------------------------------------------------------------------------------------------
    ' 2009.12.17 - 2009.12.31 전매장 세트 상품 행사 내용  -- 송화점, 쌈지공원점
    If Format(Date, "YYYY-MM-DD") >= "2009-12-17" And Format(Date, "YYYY-MM-DD") <= "2009-12-31" Then

        If Action_세트행사_20091217 = False Then
            MsgBox "2009-12-17 ~ 2009-12-31 할인 정보 설정 오류 " & Err.Description, vbInformation, "확인"
            End
        End If

        ' 대리점 정보를 다시 읽는다.
        If Fb대리점정보 = "Error" Then
            MsgBox "대리점 정보가 올바르지 않습니다. 확인후 저장하여 주십시요.", vbInformation, "확인"
            frmINIT.Show 1
            End
        End If

    End If
    '----------------------------------------------------------------------------------------------------------------------
    
    '----------------------------------------------------------------------------------------------------------------------
    ' 2009.12.18 - 2009.12.31 전매장 세트 상품 행사 내용  -- 대치은마점, 시흥장곡점
    If Format(Date, "YYYY-MM-DD") >= "2009-12-18" And Format(Date, "YYYY-MM-DD") <= "2009-12-31" Then

        If Action_세트행사_20091218 = False Then
            MsgBox "2009-12-18 ~ 2009-12-31 할인 정보 설정 오류 " & Err.Description, vbInformation, "확인"
            End
        End If

        ' 대리점 정보를 다시 읽는다.
        If Fb대리점정보 = "Error" Then
            MsgBox "대리점 정보가 올바르지 않습니다. 확인후 저장하여 주십시요.", vbInformation, "확인"
            frmINIT.Show 1
            End
        End If

    End If
    '----------------------------------------------------------------------------------------------------------------------

    '----------------------------------------------------------------------------------------------------------------------
    ' 2010.03.02 - 2010.03.15 전매장 세트 상품 행사 내용  -- (이마트 제외)
    If Format(Date, "YYYY-MM-DD") >= "2010-03-02" And Format(Date, "YYYY-MM-DD") <= "2010-03-15" Then

        If Action_세일행사_20100302("20100302", "20100315") = False Then
            MsgBox "2010-03-02 ~ 2010-03-15 할인 정보 설정 오류 " & Err.Description, vbInformation, "확인"
            End
        End If

        ' 대리점 정보를 다시 읽는다.
        If Fb대리점정보 = "Error" Then
            MsgBox "대리점 정보가 올바르지 않습니다. 확인후 저장하여 주십시요.", vbInformation, "확인"
            frmINIT.Show 1
            End
        End If

    End If
    '----------------------------------------------------------------------------------------------------------------------

    '----------------------------------------------------------------------------------------------------------------------
    ' 2010.03.04 - 2010.03.17  이마트 매장 세일 행사
    If Format(Date, "YYYY-MM-DD") >= "2010-03-04" And Format(Date, "YYYY-MM-DD") <= "2010-03-17" Then

        If Action_세일행사_20100304("20100304", "20100317") = False Then
            MsgBox "2010-03-04 ~ 2010-03-17 할인 정보 설정 오류 " & Err.Description, vbInformation, "확인"
            End
        End If

        ' 대리점 정보를 다시 읽는다.
        If Fb대리점정보 = "Error" Then
            MsgBox "대리점 정보가 올바르지 않습니다. 확인후 저장하여 주십시요.", vbInformation, "확인"
            frmINIT.Show 1
            End
        End If

    End If
    '----------------------------------------------------------------------------------------------------------------------

    '----------------------------------------------------------------------------------------------------------------------
    ' 2010.03.02 - 2010.03.15 전매장 세트 상품 행사 내용  -- (이마트 제외)
    If Format(Date, "YYYY-MM-DD") >= "2010-03-02" And Format(Date, "YYYY-MM-DD") <= "2010-03-15" Then

        If Action_세일행사_20100302_01("20100302", "20100315") = False Then
            MsgBox "2010-03-02 ~ 2010-03-15 할인 정보 설정 오류 " & Err.Description, vbInformation, "확인"
            End
        End If

        ' 대리점 정보를 다시 읽는다.
        If Fb대리점정보 = "Error" Then
            MsgBox "대리점 정보가 올바르지 않습니다. 확인후 저장하여 주십시요.", vbInformation, "확인"
            frmINIT.Show 1
            End
        End If

    End If
    '----------------------------------------------------------------------------------------------------------------------

    '----------------------------------------------------------------------------------------------------------------------
    ' 2010.03.02 - 2010.03.15 전매장 세트 상품 행사 내용  -- (이마트 제외)
    If Format(Date, "YYYY-MM-DD") >= "2010-03-02" And Format(Date, "YYYY-MM-DD") <= "2010-03-15" Then

        If Action_세일행사_20100302_02("20100302", "20100315") = False Then
            MsgBox "2010-03-02 ~ 2010-03-15 할인 정보 설정 오류 " & Err.Description, vbInformation, "확인"
            End
        End If

        ' 대리점 정보를 다시 읽는다.
        If Fb대리점정보 = "Error" Then
            MsgBox "대리점 정보가 올바르지 않습니다. 확인후 저장하여 주십시요.", vbInformation, "확인"
            frmINIT.Show 1
            End
        End If

    End If
    '----------------------------------------------------------------------------------------------------------------------


End Sub

'--------------------------------------------------------------------------------------------------------------
' Procedure : Check_이마트할인대상확인_20090730
' DateTime  : 2009-07-27
' Author    : pds2004
' Purpose   : 이마트 할인 여부   할인기간 2009-07-30 ~ 2009-08-05일 까지 전품목 20%
' 100211 이마트 황학점은 제외.... 조과장이 별도로 만들어서 적용함.(기간이 달라서)
'--------------------------------------------------------------------------------------------------------------
Public Function Check_이마트할인대상확인_20090730() As Boolean
    Dim sCompanyCode    As String
    Dim sStoreCode      As String
    
    Check_이마트할인대상확인_20090730 = False
    
    
    Select Case 대리점정보.StoreCode
            '원주점,    신월점,  동인천점, 만촌점,   월배점
        Case "100008", "100009", "100011", "100012", "100015"
                Check_이마트할인대상확인_20090730 = True: Exit Function
                
            '울산학성점, 평택점,  은평점,  연제점,   비산점
        Case "100021", "100022", "100023", "100027", "100028"
                Check_이마트할인대상확인_20090730 = True: Exit Function
                
            '칠성점,    구미점,  시화점,   해운대점,   고잔점
        Case "100029", "100030", "100031", "100032", "100038"
                Check_이마트할인대상확인_20090730 = True: Exit Function
                
            '서수원점,  경산점,  상주점,   공항점,   파주점
        Case "100056", "100084", "100120", "100123", "100126"
                Check_이마트할인대상확인_20090730 = True: Exit Function
                
            '수지점,   양주점,   송림점,   도농점,   하남점
        Case "100128", "100143", "100170", "100197", "100200"
                Check_이마트할인대상확인_20090730 = True: Exit Function
                
            '안성점,    성수점,  부평점,    목동점
        Case "100222", "100240", "100259", "100264"
                Check_이마트할인대상확인_20090730 = True: Exit Function
                
            '테스트 매장(개발자)
        Case "999999"
                Check_이마트할인대상확인_20090730 = True: Exit Function
                
                
        Case Else
                Check_이마트할인대상확인_20090730 = False: Exit Function
    End Select
    
End Function


'이마트행사품목: 상의류( f코드 ), 하의류( g코드 ), 스커트류( r코드 )만 20% 할인 행사를 진행하며 나머지 품목은 제외 함
'> 기간은 11.1~11.15일까지
Public Function Check_이마트할인_20090730() As Boolean
    Dim sDay        As String
    Dim dblPrice    As Double

    Dim FHandle     As Integer
    
    Dim sStartDate  As String
    Dim sEndDate    As String
    
    On Error GoTo Check_일반할인_Error
    
    Check_이마트할인_20090730 = False
        
    sStartDate = "20090730"
    sEndDate = "20090805"
    
    If Check_이마트할인대상확인_20090730 = True Then
        sDay = Format(Date, "YYYY-MM-DD")
        
        If sDay >= Format(sStartDate, "@@@@-@@-@@") And sDay <= Format(sEndDate, "@@@@-@@-@@") Then
            ' 해당 일자에 처음 실행할 경우 해당 파일이 없어서 실해되고 2번째부터는 실행되지 않느다.
            ' 해당 기간이 지나면 자동으로 프로그램에서 사용하지 않는다.
            If Dir(App.Path & "\" & sEndDate & ".TXT", vbDirectory) = "" Then
                ' 다음 이중 실행되지 않도록 파일을 생성한다.
                FHandle = FreeFile
                Open App.Path & "\" & sEndDate & ".TXT" For Append As FHandle
                Print #FHandle, Now
                
                Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                Set Rs = New ADODB.Recordset
                Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
                
                ' 이전 자료를 모두 지운다.
                Query = "DELETE FROM 할인정보 WHERE 시작일 = '" & sStartDate & "' AND 종료일 = '" & sEndDate & "' "
                ADOCon.Execute Query
                
                Do While Not Rs.EOF
                    If IsNumeric(Rs.Fields("가격")) = True Then
                        Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * 0.8)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * 0.8) * 0.01) * 100)); "]"
                    
                        dblPrice = 0
                        dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * 0.8) * 0.01) * 100)        ' 20% 할인을 적용한다.
                        
                        '------------------------------------------------------------------------------------------------------
                        Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                        Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                        Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                        ADOCon.Execute Query
                    Else
                        Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                        Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                        MsgBox Query, vbCritical, "경고"
                    End If
                    
                    Rs.MoveNext
                Loop
                Rs.Close
                Set Rs = Nothing
                
''               나머지 품목도 할인 코드에 넣어주어야 정상 동작된다.
'                Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
'                Query = Query & " WHERE not ( left(구분코드,1) = 'f' or left(구분코드,1) = 'g' or  left(구분코드,1) = 'r'  )"
'                Set rs = MyDB.OpenRecordset(Query)
'
'                Do While Not rs.EOF
'                    If IsNumeric(rs.Fields("가격")) = True Then
'                        Print #FHandle, "["; rs.Fields("구분코드"); ":"; rs.Fields("품명"); Tab; Tab; ":"; rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(rs.Fields("가격"))) * 0.8)); "]"; Tab; "["; CStr((CLng((Val(CStr(rs.Fields("가격"))) * 0.8) * 0.01) * 100)); "]"
'
'                        dblPrice = Val(CStr(rs.Fields("가격")))
'
'                        Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
'                        Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & rs.Fields("구분코드") & "', '"
'                        Query = Query & rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
'                        ADOCon.Execute Query
'                    Else
'                        Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
'                        Query = Query & "[" & rs.Fields("구분코드") & ":" & rs.Fields("품명") & ":" & rs.Fields("가격") & "]"
'                        MsgBox Query, vbCritical, "경고"
'                    End If
'                    rs.MoveNext
'                Loop
'                rs.Close
                                
                Close #FHandle
            End If
        End If
    End If
    
    Check_이마트할인_20090730 = True

    On Error GoTo 0
    
    Exit Function

Check_일반할인_Error:
    Resume
    Check_이마트할인_20090730 = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Check_이마트할인_20090730 of Module Sale"
End Function


'--------------------------------------------------------------------------------------------------------------
' Procedure : Check_이마트할인대상확인_20091029
' DateTime  : 2009-10-23
' Author    : pds2004
' Purpose   : 이마트 할인 여부   할인기간 2009-10-29 ~ 2009-11-04일 까지 전품목 30%
'--------------------------------------------------------------------------------------------------------------
Public Function Check_이마트할인대상확인_20091029() As Boolean
    Dim sCompanyCode    As String
    Dim sStoreCode      As String
    
    Check_이마트할인대상확인_20091029 = False
    
    
    Select Case 대리점정보.StoreCode
            '동인천점   연재점   경산점    고잔점    공항점
        Case "100011", "100027", "100084", "100038", "100123"
                Check_이마트할인대상확인_20091029 = True: Exit Function
                
            '구미점     도농점    만촌점    목동점    부평점
        Case "100030", "100197", "100012", "100264", "100259"
                Check_이마트할인대상확인_20091029 = True: Exit Function
                
            '비산점     상주점    서수원    성수점    송림점
        Case "100028", "100120", "100056", "100240", "100170"
                Check_이마트할인대상확인_20091029 = True: Exit Function
                
            '수지점     시화점    신월점    안성점    양주점
        Case "100128", "100031", "100009", "100222", "100143"
                Check_이마트할인대상확인_20091029 = True: Exit Function
                
            '울산학성점 원주점    월배점    은평점   청계천점
        Case "100021", "100008", "100015", "100023", "100211"
                Check_이마트할인대상확인_20091029 = True: Exit Function
                
            '칠성점     파주점    평택점    하남점   해운대점
        Case "100029", "100126", "100022", "100200", "100032"
                Check_이마트할인대상확인_20091029 = True: Exit Function
                
            '테스트 매장(개발자)
        Case "999999"
                Check_이마트할인대상확인_20091029 = True: Exit Function
                
                
        Case Else
                Check_이마트할인대상확인_20091029 = False: Exit Function
    End Select
    
End Function



'이마트행사품목: 전품목 30% 할인 행사 제외 품목 없음
'> 기간은 2009.10.29 - 2009.11.04일 까지
Public Function Check_이마트할인_20091029() As Boolean
    Dim sDay        As String
    Dim dblPrice    As Double

    Dim FHandle     As Integer
    Dim nPercent    As Single
    
    Dim sStartDate  As String
    Dim sEndDate    As String
        

    On Error GoTo Check_일반할인_Error
    
    Check_이마트할인_20091029 = False
    
    
    sStartDate = "20091029"
    sEndDate = "20091104"
    nPercent = 0.7          '<- 30% 할인
    
    If Check_이마트할인대상확인_20091029 = True Then
        
        sDay = Format(Date, "YYYY-MM-DD")
        If sDay >= Format(sStartDate, "@@@@-@@-@@") And sDay <= Format(sEndDate, "@@@@-@@-@@") Then
                
            ' 해당 일자에 처음 실행할 경우 해당 파일이 없어서 실해되고 2번째부터는 실행되지 않는다.
            ' 해당 기간이 지나면 자동으로 프로그램에서 사용하지 않는다.
            If Dir(App.Path & "\" & sEndDate & ".TXT", vbDirectory) = "" Then
                ' 다음 이중 실행되지 않도록 파일을 생성한다.
                FHandle = FreeFile
                Open App.Path & "\" & sEndDate & ".TXT" For Append As FHandle
                Print #FHandle, Now
                
                '--------------------------------------------------------
                '
                '--------------------------------------------------------
                Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                Set Rs = New ADODB.Recordset
                Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
                
                '--------------------------------------------------------
                ' 이전 자료를 모두 지운다.
                '--------------------------------------------------------
                Query = "DELETE FROM 할인정보"
                Query = Query & " WHERE 시작일 = '" & sStartDate & "'"
                Query = Query & "   AND 종료일 = '" & sEndDate & "' "
                ADOCon.Execute Query
                
                Do While Not Rs.EOF
                    If IsNumeric(Rs.Fields("가격")) = True Then
                        Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"
                    
                        dblPrice = 0
                        dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)
                        
                        Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                        Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                        Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                        ADOCon.Execute Query
                    Else
                        Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                        Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                        MsgBox Query, vbCritical, "경고"
                    End If
                    
                    Rs.MoveNext
                Loop
                Rs.Close
                Set Rs = Nothing
                
''               나머지 품목도 할인 코드에 넣어주어야 정상 동작된다.
'                Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
'                Query = Query & " WHERE not ( left(구분코드,1) = 'f' or left(구분코드,1) = 'g' or  left(구분코드,1) = 'r'  )"
'                Set rs = MyDB.OpenRecordset(Query)
'
'                Do While Not rs.EOF
'                    If IsNumeric(rs.Fields("가격")) = True Then
'                        Print #FHandle, "["; rs.Fields("구분코드"); ":"; rs.Fields("품명"); Tab; Tab; ":"; rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"
'
'                        dblPrice = Val(CStr(rs.Fields("가격")))
'
'                        Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
'                        Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & rs.Fields("구분코드") & "', '"
'                        Query = Query & rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
'                        ADOCon.Execute Query
'                    Else
'                        Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
'                        Query = Query & "[" & rs.Fields("구분코드") & ":" & rs.Fields("품명") & ":" & rs.Fields("가격") & "]"
'                        MsgBox Query, vbCritical, "경고"
'                    End If
'                    rs.MoveNext
'                Loop
'                rs.Close
                
                
                Close #FHandle
            End If
        End If
    End If
    
    
    Check_이마트할인_20091029 = True

    On Error GoTo 0
    Exit Function

Check_일반할인_Error:
    Resume
    Check_이마트할인_20091029 = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Check_이마트할인_20091029 of Module Sale"
End Function



'--------------------------------------------------------------------------------------------------------------
' Procedure : Check_이마트할인제외확인_20091109
' DateTime  : 2009-10-23
' Author    : pds2004
' Purpose   : 행사 기간중 이마트는 제외됨 행사기간 2009-11-09 ~ 2009-11-14일
'             상의류,하의류 한벌 점수할 경우 m00, m01,m20 코드는 무료로 세탁
'--------------------------------------------------------------------------------------------------------------
Public Function Check_이마트할인제외확인_20091109() As Boolean
    Dim sCompanyCode    As String
    Dim sStoreCode      As String
    
    Dim bMode           As Boolean
    
    bMode = False
    
    
    Select Case 대리점정보.StoreCode
            '동인천점   연재점   경산점    고잔점    공항점
        Case "100011", "100027", "100084", "100038", "100123"
                bMode = True
                
            '구미점     도농점    만촌점    목동점    부평점
        Case "100030", "100197", "100012", "100264", "100259"
                bMode = True
                
            '비산점     상주점    서수원    성수점    송림점
        Case "100028", "100120", "100056", "100240", "100170"
                bMode = True
                
            '수지점     시화점    신월점    안성점    양주점
        Case "100128", "100031", "100009", "100222", "100143"
                bMode = True
                
            '울산학성점 원주점    월배점    은평점   청계천점
        Case "100021", "100008", "100015", "100023", "100211"
                bMode = True
                
            '칠성점     파주점    평택점    하남점   해운대점
        Case "100029", "100126", "100022", "100200", "100032"
                bMode = True
                
        
        Case Else
                bMode = False
    End Select
    
    Check_이마트할인제외확인_20091109 = Not bMode
End Function


'  행사품목: 빼빼로 데이 할인 행사
'> 기간은 2009.11.09 - 2009.11.14일 까지
Public Function Action_빼빼로행사_20091109() As Boolean
    Dim sDay        As String
    Dim dblPrice    As Double
    Dim FHandle     As Integer
    Dim nPercent    As Single
    
    Dim sStartDate  As String
    Dim sEndDate    As String
    
    On Error GoTo Check_일반할인_Error
    
    Action_빼빼로행사_20091109 = False
    
    sStartDate = "20091109"
    sEndDate = "20091114"
    
    'nPercent = 0.7          '<- 30% 할인
    
    ' 행사 제외 이마트일 경우 False 가 날라온다.
    If Check_이마트할인제외확인_20091109 = True Then
        sDay = Format(Date, "YYYY-MM-DD")
        
        If sDay >= Format(sStartDate, "@@@@-@@-@@") And sDay <= Format(sEndDate, "@@@@-@@-@@") Then
            If Dir(App.Path & "\" & sStartDate & ".TXT", vbDirectory) = "" Then
                ' 다음 이중 실행되지 않도록 파일을 생성한다.
                FHandle = FreeFile
                Open App.Path & "\" & sStartDate & ".TXT" For Append As FHandle
                Print #FHandle, Now
            
                '--------------------------------------------------------
                ' 지정 할인의 행사 내용을 적용 시킨다.
                '--------------------------------------------------------
                Query = "UPDATE 대리점정보 SET "
                Query = Query & "    지정할인여부   = 'Y', "
                Query = Query & "    지정할인비율     = '100', "
                Query = Query & "    지정할인시작일     = '20091109',  "
                Query = Query & "    지정할인종료일     = '20091114'   "
                ADOCon.Execute Query
                
                Close #FHandle
            End If
        End If
    End If
        
    Action_빼빼로행사_20091109 = True

    On Error GoTo 0
    Exit Function

Check_일반할인_Error:
    Resume
    Action_빼빼로행사_20091109 = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Action_빼빼로행사_20091109 of Module Sale"
End Function


'  행사품목: 빼빼로 데이 할인 행사
'> 기간은 2009.11.09 - 2009.11.14일 까지
Public Function Action_빼빼로행사_코드확인(MySpr As fpSpread, nActRow As Long) As String
    Dim bAction As String
    
    Dim bSaleCode1  As Boolean
    Dim bSaleCode2  As Boolean
    Dim lRow        As Long
    Dim varTemp     As Variant
    
    On Error GoTo Action_빼빼로행사_코드확인_Error
    
    bAction = "Z"
    bSaleCode1 = False:        bSaleCode2 = False
    
    Call MySpr.GetText(7, nActRow, varTemp)
    
    If InStr("m00, m01, m20", LCase(CStr(varTemp))) <= 0 Then
        Action_빼빼로행사_코드확인 = "A"
        Exit Function
    End If

    
    For lRow = 1 To MySpr.MaxRows
        Call MySpr.GetText(7, lRow, varTemp)
        
        If LCase(CStr(varTemp)) = "" Then
            Exit For
        ElseIf LCase(CStr(varTemp)) >= "f00" And LCase(CStr(varTemp)) <= "f30" Then
            bSaleCode1 = True
        ElseIf LCase(CStr(varTemp)) >= "g00" And LCase(CStr(varTemp)) <= "g35" Then
            bSaleCode2 = True
        End If
    
        If bSaleCode1 = True And bSaleCode2 = True Then
            bAction = ""
            Exit For
        End If
    Next lRow
        
    Action_빼빼로행사_코드확인 = bAction


    On Error GoTo 0
    
    Exit Function

Action_빼빼로행사_코드확인_Error:
    Action_빼빼로행사_코드확인 = "Z"
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Action_빼빼로행사_코드확인 of Module Sale"
End Function


'--------------------------------------------------------------------------------------------------------------
' Procedure : Check_세트행사할인대상확인_20091211
' DateTime  : 2009-12-09
' Author    : pds2004
' Purpose   : 세트상품 할인 여부   할인기간 2009-12-11 ~ 2009-12-31일 까지 전품목 10%
' 100211 이마트 황학점은 제외.... 조과장이 별도로 만들어서 적용함.(기간이 달라서)
'--------------------------------------------------------------------------------------------------------------
Public Function Check_세트행사할인대상확인_20091211() As Boolean
    Dim sCompanyCode    As String
    Dim sStoreCode      As String
    
    Check_세트행사할인대상확인_20091211 = False
    
    If 대리점정보.MasterCode = M_COUPON_KLENZ_CODE Then
        Check_세트행사할인대상확인_20091211 = False
        Exit Function
    End If
    
    Select Case 대리점정보.StoreCode
            '현대캐피탈점, 매화주공점, 송화점, 쌈지공원점, 목동사거리점, 대치은마점, 시흥장곡점, 평촌대원점, 동천레미안점
        Case "100280", "100320", "100321", "100322", "100323", "100324", "100325", "100326", "100328"
            Check_세트행사할인대상확인_20091211 = False:    Exit Function
                
        ' 할인 대상으로 처리됨
        Case Else
            Check_세트행사할인대상확인_20091211 = True:   Exit Function
            
    End Select
    
End Function

'--------------------------------------------------------------------------------
'세트 상품 행사로 인한 전품목 10% 할인
'> 기간은 2009.12.11 - 2009.12.31일 까지
Public Function Action_세트행사_20091211() As Boolean
    Dim sDay        As String
    Dim dblPrice    As Double
    Dim FHandle     As Integer
    Dim nPercent    As Single

    Dim sStartDate  As String
    Dim sEndDate    As String



    On Error GoTo Check_할인_Error
    Action_세트행사_20091211 = False


    sStartDate = "20091211"
    sEndDate = "20091231"
    nPercent = 0.9          '<- 10% 할인

    If Check_세트행사할인대상확인_20091211 = True Then

        sDay = Format(Date, "YYYY-MM-DD")
        If sDay >= Format(sStartDate, "@@@@-@@-@@") And sDay <= Format(sEndDate, "@@@@-@@-@@") Then

            ' 해당 일자에 처음 실행할 경우 해당 파일이 없어서 실해되고 2번째부터는 실행되지 않는다.
            ' 해당 기간이 지나면 자동으로 프로그램에서 사용하지 않는다.
            If Dir(App.Path & "\" & sEndDate & ".TXT", vbDirectory) = "" Then
                ' 다음 이중 실행되지 않도록 파일을 생성한다.
                FHandle = FreeFile
                Open App.Path & "\" & sEndDate & ".TXT" For Append As FHandle
                Print #FHandle, Now

                Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                Query = Query & " WHERE NOT ( left(구분코드,1) = 'a' or 구분코드 = 'm00' or 구분코드 = 'm01'  or 구분코드 = 'm02'  )"
                Set Rs = New ADODB.Recordset
                Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly


                ' 이전 자료를 모두 지운다.
                Query = "DELETE FROM 할인정보 WHERE 시작일 = '" & sStartDate & "' AND 종료일 = '" & sEndDate & "' "
                ADOCon.Execute Query

                Do While Not Rs.EOF
                    If IsNumeric(Rs.Fields("가격")) = True Then
                        Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"

                        dblPrice = 0
                        dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)

                        Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                        Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                        Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                        ADOCon.Execute Query
                    Else
                        Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                        Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                        MsgBox Query, vbCritical, "경고"
                    End If
                    Rs.MoveNext
                Loop
                Rs.Close


'               나머지 품목도 할인 코드에 넣어주어야 정상 동작된다.
                Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                Query = Query & " WHERE ( left(구분코드,1) = 'a' or 구분코드 = 'm00' or 구분코드 = 'm01'  or 구분코드 = 'm02'  ) "
                Set Rs = New ADODB.Recordset
                Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

                Do While Not Rs.EOF
                    If IsNumeric(Rs.Fields("가격")) = True Then
                        Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"

                        dblPrice = Val(CStr(Rs.Fields("가격")))

                        Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                        Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                        Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                        ADOCon.Execute Query
                    Else
                        Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                        Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                        MsgBox Query, vbCritical, "경고"
                    End If
                    
                    Rs.MoveNext
                Loop
                Rs.Close

                Close #FHandle
            End If
        End If
    End If
    
    Action_세트행사_20091211 = True

    On Error GoTo 0
    Exit Function

Check_할인_Error:
    Resume
    Action_세트행사_20091211 = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Action_세트행사_20091211 of Module Sale"
End Function

'--------------------------------------------------------------------------------------------------------------
' Procedure : Check_세트행사할인대상확인_20091211
' DateTime  : 2009-12-09
' Author    : pds2004
' Purpose   : 세트상품 할인 여부   할인기간 2009-12-11 ~ 2009-12-31일 까지 전품목 10%
' 100211 이마트 황학점은 제외.... 조과장이 별도로 만들어서 적용함.(기간이 달라서)
'--------------------------------------------------------------------------------------------------------------
Public Function Check_세트행사할인대상확인_20091213() As Boolean
    Dim sCompanyCode    As String
    Dim sStoreCode      As String
    
    Check_세트행사할인대상확인_20091213 = False
    
    If 대리점정보.MasterCode = M_COUPON_KLENZ_CODE Then
        Check_세트행사할인대상확인_20091213 = False
        Exit Function
    End If
    
    Select Case 대리점정보.StoreCode
            ' 매화주공점
        Case "100320"
            Check_세트행사할인대상확인_20091213 = True:    Exit Function
                
        ' 할인 대상으로 처리됨
        Case Else
            Check_세트행사할인대상확인_20091213 = False:   Exit Function
            
    End Select
    
End Function
'--------------------------------------------------------------------------------
'세트 상품 행사로 인한 전품목 10% 할인
'> 기간은 2009.12.13 - 2009.12.31일 까지
Public Function Action_세트행사_20091213() As Boolean
    Dim sDay        As String
    Dim dblPrice    As Double
    Dim Rs          As Recordset
    Dim FHandle     As Integer
    Dim nPercent    As Single

    Dim sStartDate  As String
    Dim sEndDate    As String



    On Error GoTo Check_할인_Error
    Action_세트행사_20091213 = False


    sStartDate = "20091213"
    sEndDate = "20091231"
    nPercent = 0.9          '<- 10% 할인

    If Check_세트행사할인대상확인_20091213 = True Then

        sDay = Format(Date, "YYYY-MM-DD")
        If sDay >= Format(sStartDate, "@@@@-@@-@@") And sDay <= Format(sEndDate, "@@@@-@@-@@") Then

            ' 해당 일자에 처음 실행할 경우 해당 파일이 없어서 실해되고 2번째부터는 실행되지 않는다.
            ' 해당 기간이 지나면 자동으로 프로그램에서 사용하지 않는다.
            If Dir(App.Path & "\" & sEndDate & ".TXT", vbDirectory) = "" Then
                ' 다음 이중 실행되지 않도록 파일을 생성한다.
                FHandle = FreeFile
                Open App.Path & "\" & sEndDate & ".TXT" For Append As FHandle
                Print #FHandle, Now

                Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                Query = Query & " WHERE NOT ( left(구분코드,1) = 'a' or 구분코드 = 'm00' or 구분코드 = 'm01'  or 구분코드 = 'm02'  )"
                Set Rs = New ADODB.Recordset
                Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

                ' 이전 자료를 모두 지운다.
                Query = "DELETE FROM 할인정보 WHERE 시작일 = '" & sStartDate & "' AND 종료일 = '" & sEndDate & "' "
                ADOCon.Execute Query

                Do While Not Rs.EOF
                    If IsNumeric(Rs.Fields("가격")) = True Then
                        Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"

                        dblPrice = 0
                        dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)

                        Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                        Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                        Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                        ADOCon.Execute Query
                    Else
                        Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                        Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                        MsgBox Query, vbCritical, "경고"
                    End If
                    Rs.MoveNext
                Loop
                Rs.Close


'               나머지 품목도 할인 코드에 넣어주어야 정상 동작된다.
                Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                Query = Query & " WHERE ( left(구분코드,1) = 'a' or 구분코드 = 'm00' or 구분코드 = 'm01'  or 구분코드 = 'm02'  ) "
                Set Rs = New ADODB.Recordset
                Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

                Do While Not Rs.EOF
                    If IsNumeric(Rs.Fields("가격")) = True Then
                        Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"

                        dblPrice = Val(CStr(Rs.Fields("가격")))

                        Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                        Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                        Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                        ADOCon.Execute Query
                    Else
                        Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                        Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                        MsgBox Query, vbCritical, "경고"
                    End If
                    Rs.MoveNext
                Loop
                Rs.Close


                Close #FHandle
            End If
        End If
    End If
    
    Action_세트행사_20091213 = True

    On Error GoTo 0
    Exit Function

Check_할인_Error:
    Resume
    Action_세트행사_20091213 = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Action_세트행사_20091213 of Module Sale"
End Function



'--------------------------------------------------------------------------------------------------------------
' Procedure : Check_세트행사할인대상확인_20091216
' DateTime  : 2009-12-09
' Author    : pds2004
' Purpose   : 세트상품 할인 여부   할인기간 2009-12-16 ~ 2009-12-31일 까지 전품목 10%
' 100211 이마트 황학점은 제외.... 조과장이 별도로 만들어서 적용함.(기간이 달라서)
'--------------------------------------------------------------------------------------------------------------
Public Function Check_세트행사할인대상확인_20091216() As Boolean
    Dim sCompanyCode    As String
    Dim sStoreCode      As String
    
    Check_세트행사할인대상확인_20091216 = False
    
    If 대리점정보.MasterCode = M_COUPON_KLENZ_CODE Then
        Check_세트행사할인대상확인_20091216 = False
        Exit Function
    End If
    
    Select Case 대리점정보.StoreCode
            ' 송화점, 쌈지공원점
        Case "100321", "100322"
            Check_세트행사할인대상확인_20091216 = True:    Exit Function
                
        ' 할인 대상으로 처리됨
        Case Else
            Check_세트행사할인대상확인_20091216 = False:   Exit Function
            
    End Select
    
End Function
'--------------------------------------------------------------------------------
'세트 상품 행사로 인한 전품목 10% 할인
'> 기간은 2009.12.16 - 2009.12.31일 까지
Public Function Action_세트행사_20091216() As Boolean
    Dim sDay        As String
    Dim dblPrice    As Double
    Dim FHandle     As Integer
    Dim nPercent    As Single

    Dim sStartDate  As String
    Dim sEndDate    As String

    On Error GoTo Check_할인_Error
    Action_세트행사_20091216 = False


    sStartDate = "20091216"
    sEndDate = "20091231"
    nPercent = 0.9          '<- 10% 할인

    If Check_세트행사할인대상확인_20091216 = True Then

        sDay = Format(Date, "YYYY-MM-DD")
        
        If sDay >= Format(sStartDate, "@@@@-@@-@@") And sDay <= Format(sEndDate, "@@@@-@@-@@") Then

            ' 해당 일자에 처음 실행할 경우 해당 파일이 없어서 실해되고 2번째부터는 실행되지 않는다.
            ' 해당 기간이 지나면 자동으로 프로그램에서 사용하지 않는다.
            If Dir(App.Path & "\" & sEndDate & ".TXT", vbDirectory) = "" Then
                ' 다음 이중 실행되지 않도록 파일을 생성한다.
                FHandle = FreeFile
                Open App.Path & "\" & sEndDate & ".TXT" For Append As FHandle
                Print #FHandle, Now

                Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                Query = Query & " WHERE NOT ( left(구분코드,1) = 'a' or 구분코드 = 'm00' or 구분코드 = 'm01'  or 구분코드 = 'm02'  )"
                Set Rs = New ADODB.Recordset
                Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

                ' 이전 자료를 모두 지운다.
                Query = "DELETE FROM 할인정보 WHERE 시작일 = '" & sStartDate & "' AND 종료일 = '" & sEndDate & "' "
                ADOCon.Execute Query

                Do While Not Rs.EOF
                    If IsNumeric(Rs.Fields("가격")) = True Then
                        Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"

                        dblPrice = 0
                        dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)

                        Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                        Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                        Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                        ADOCon.Execute Query
                    Else
                        Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                        Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                        MsgBox Query, vbCritical, "경고"
                    End If
                    Rs.MoveNext
                Loop
                Rs.Close


'               나머지 품목도 할인 코드에 넣어주어야 정상 동작된다.
                Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                Query = Query & " WHERE ( left(구분코드,1) = 'a' or 구분코드 = 'm00' or 구분코드 = 'm01'  or 구분코드 = 'm02'  ) "
                Set Rs = MyDB.OpenRecordset(Query)

                Do While Not Rs.EOF
                    If IsNumeric(Rs.Fields("가격")) = True Then
                        Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"

                        dblPrice = Val(CStr(Rs.Fields("가격")))

                        Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                        Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                        Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                        ADOCon.Execute Query
                    Else
                        Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                        Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                        MsgBox Query, vbCritical, "경고"
                    End If
                    Rs.MoveNext
                Loop
                Rs.Close


                Close #FHandle
            End If
        End If
    End If
    
    Action_세트행사_20091216 = True

    On Error GoTo 0
    Exit Function

Check_할인_Error:
    Resume
    Action_세트행사_20091216 = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Action_세트행사_20091216 of Module Sale"
End Function




'--------------------------------------------------------------------------------------------------------------
' Procedure : Check_세트행사할인대상확인_20091217
' DateTime  : 2009-12-09
' Author    : pds2004
' Purpose   : 세트상품 할인 여부   할인기간 2009-12-17 ~ 2009-12-31일 까지 전품목 10%
' 100211 이마트 황학점은 제외.... 조과장이 별도로 만들어서 적용함.(기간이 달라서)
'--------------------------------------------------------------------------------------------------------------
Public Function Check_세트행사할인대상확인_20091217() As Boolean
    Dim sCompanyCode    As String
    Dim sStoreCode      As String
    
    Check_세트행사할인대상확인_20091217 = False
    
    If 대리점정보.MasterCode = M_COUPON_KLENZ_CODE Then
        Check_세트행사할인대상확인_20091217 = False
        Exit Function
    End If
    
    Select Case 대리점정보.StoreCode
            ' 목동사거리점
        Case "100323"
            Check_세트행사할인대상확인_20091217 = True:    Exit Function
                
        ' 할인 대상으로 처리됨
        Case Else
            Check_세트행사할인대상확인_20091217 = False:   Exit Function
            
    End Select
    
End Function
'--------------------------------------------------------------------------------
'세트 상품 행사로 인한 전품목 10% 할인
'> 기간은 2009.12.17 - 2009.12.31일 까지
Public Function Action_세트행사_20091217() As Boolean
    Dim sDay        As String
    Dim dblPrice    As Double
    Dim FHandle     As Integer
    Dim nPercent    As Single

    Dim sStartDate  As String
    Dim sEndDate    As String

    On Error GoTo Check_할인_Error
    
    Action_세트행사_20091217 = False


    sStartDate = "20091217"
    sEndDate = "20091231"
    nPercent = 0.9          '<- 10% 할인

    If Check_세트행사할인대상확인_20091217 = True Then

        sDay = Format(Date, "YYYY-MM-DD")
        If sDay >= Format(sStartDate, "@@@@-@@-@@") And sDay <= Format(sEndDate, "@@@@-@@-@@") Then

            ' 해당 일자에 처음 실행할 경우 해당 파일이 없어서 실해되고 2번째부터는 실행되지 않는다.
            ' 해당 기간이 지나면 자동으로 프로그램에서 사용하지 않는다.
            If Dir(App.Path & "\" & sEndDate & ".TXT", vbDirectory) = "" Then
                ' 다음 이중 실행되지 않도록 파일을 생성한다.
                FHandle = FreeFile
                Open App.Path & "\" & sEndDate & ".TXT" For Append As FHandle
                Print #FHandle, Now
                
                '-----------------------------------------------------------
                '
                '-----------------------------------------------------------
                Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                Query = Query & " WHERE NOT ( left(구분코드,1) = 'a'"
                Query = Query & "    OR 구분코드 = 'm00'"
                Query = Query & "    OR 구분코드 = 'm01'"
                Query = Query & "    OR 구분코드 = 'm02')"
                Set Rs = New ADODB.Recordset
                Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

                ' 이전 자료를 모두 지운다.
                Query = "DELETE FROM 할인정보 WHERE 시작일 = '" & sStartDate & "' AND 종료일 = '" & sEndDate & "' "
                ADOCon.Execute Query

                Do While Not Rs.EOF
                    If IsNumeric(Rs.Fields("가격")) = True Then
                        Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"

                        dblPrice = 0
                        dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)

                        Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                        Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                        Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                        ADOCon.Execute Query
                    Else
                        Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                        Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                        MsgBox Query, vbCritical, "경고"
                    End If
                    Rs.MoveNext
                Loop
                Rs.Close


'               나머지 품목도 할인 코드에 넣어주어야 정상 동작된다.
                Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                Query = Query & " WHERE ( left(구분코드,1) = 'a' or 구분코드 = 'm00' or 구분코드 = 'm01'  or 구분코드 = 'm02'  ) "
                Set Rs = New ADODB.Recordset
                Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

                Do Until Rs.EOF
                    If IsNumeric(Rs.Fields("가격")) = True Then
                        Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"

                        dblPrice = Val(CStr(Rs.Fields("가격")))

                        Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                        Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                        Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                        ADOCon.Execute Query
                    Else
                        Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                        Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                        MsgBox Query, vbCritical, "경고"
                    End If
                    
                    Rs.MoveNext
                Loop
                Rs.Close
                Set Rs = Nothing

                Close #FHandle
            End If
        End If
    End If
    
    Action_세트행사_20091217 = True

    On Error GoTo 0
    Exit Function

Check_할인_Error:
    Resume
    Action_세트행사_20091217 = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Action_세트행사_20091217 of Module Sale"
End Function



'--------------------------------------------------------------------------------------------------------------
' Procedure : Check_세트행사할인대상확인_20091218
' DateTime  : 2009-12-09
' Author    : pds2004
' Purpose   : 세트상품 할인 여부   할인기간 2009-12-18 ~ 2009-12-31일 까지 전품목 10%
'
'--------------------------------------------------------------------------------------------------------------
Public Function Check_세트행사할인대상확인_20091218() As Boolean
    Dim sCompanyCode    As String
    Dim sStoreCode      As String
    
    Check_세트행사할인대상확인_20091218 = False
    
    If 대리점정보.MasterCode = M_COUPON_KLENZ_CODE Then
        Check_세트행사할인대상확인_20091218 = False
        Exit Function
    End If
    
    Select Case 대리점정보.StoreCode
            ' 대치은마점, 시흥장곡점
        Case "100324", "100325"
            Check_세트행사할인대상확인_20091218 = True:    Exit Function
                
        ' 할인 대상으로 처리됨
        Case Else
            Check_세트행사할인대상확인_20091218 = False:   Exit Function
            
    End Select
    
End Function
'--------------------------------------------------------------------------------
'세트 상품 행사로 인한 전품목 10% 할인
'> 기간은 2009.12.18 - 2009.12.31일 까지
Public Function Action_세트행사_20091218() As Boolean
    Dim sDay        As String
    Dim dblPrice    As Double
    Dim FHandle     As Integer
    Dim nPercent    As Single

    Dim sStartDate  As String
    Dim sEndDate    As String

    On Error GoTo Check_할인_Error
    
    Action_세트행사_20091218 = False


    sStartDate = "20091218"
    sEndDate = "20091231"
    nPercent = 0.9          '<- 10% 할인

    If Check_세트행사할인대상확인_20091218 = True Then
        sDay = Format(Date, "YYYY-MM-DD")
        
        If sDay >= Format(sStartDate, "@@@@-@@-@@") And sDay <= Format(sEndDate, "@@@@-@@-@@") Then

            ' 해당 일자에 처음 실행할 경우 해당 파일이 없어서 실해되고 2번째부터는 실행되지 않는다.
            ' 해당 기간이 지나면 자동으로 프로그램에서 사용하지 않는다.
            If Dir(App.Path & "\" & sEndDate & ".TXT", vbDirectory) = "" Then
                ' 다음 이중 실행되지 않도록 파일을 생성한다.
                FHandle = FreeFile
                Open App.Path & "\" & sEndDate & ".TXT" For Append As FHandle
                Print #FHandle, Now

                Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                Query = Query & " WHERE NOT ( left(구분코드,1) = 'a' or 구분코드 = 'm00' or 구분코드 = 'm01'  or 구분코드 = 'm02'  )"
                Set Rs = New ADODB.Recordset
                Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

                ' 이전 자료를 모두 지운다.
                Query = "DELETE FROM 할인정보 WHERE 시작일 = '" & sStartDate & "' AND 종료일 = '" & sEndDate & "' "
                ADOCon.Execute Query

                Do While Not Rs.EOF
                    If IsNumeric(Rs.Fields("가격")) = True Then
                        Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"

                        dblPrice = 0
                        dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)

                        Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                        Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                        Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                        ADOCon.Execute Query
                    Else
                        Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                        Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                        MsgBox Query, vbCritical, "경고"
                    End If
                    Rs.MoveNext
                Loop
                Rs.Close


'               나머지 품목도 할인 코드에 넣어주어야 정상 동작된다.
                Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
                Query = Query & " WHERE ( left(구분코드,1) = 'a' or 구분코드 = 'm00' or 구분코드 = 'm01'  or 구분코드 = 'm02'  ) "
                Set Rs = New ADODB.Recordset
                Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

                Do While Not Rs.EOF
                    If IsNumeric(Rs.Fields("가격")) = True Then
                        Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"

                        dblPrice = Val(CStr(Rs.Fields("가격")))

                        Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                        Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                        Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                        ADOCon.Execute Query
                    Else
                        Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                        Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                        MsgBox Query, vbCritical, "경고"
                    End If
                    Rs.MoveNext
                Loop
                Rs.Close


                Close #FHandle
            End If
        End If
    End If
    
    Action_세트행사_20091218 = True

    On Error GoTo 0
    Exit Function

Check_할인_Error:
    Resume
    Action_세트행사_20091218 = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Action_세트행사_20091218 of Module Sale"
End Function


'--------------------------------------------------------------------------------------------------------------
' Procedure : Check_이마트할인제외확인_20100302
' DateTime  : 2010-02-03
' Author    : pds2004
' Purpose   : 행사 기간중 이마트는 제외됨 행사기간 2010-03-02 ~ 2010-03-15일
'--------------------------------------------------------------------------------------------------------------
Public Function Check_이마트할인제외확인_20100302() As Boolean
    Dim sCompanyCode    As String
    Dim sStoreCode      As String
    
    Dim bMode           As Boolean
    
    bMode = True
    
    
    Select Case 대리점정보.StoreCode
            '동인천점   연재점   경산점    고잔점    공항점
        Case "100011", "100027", "100084", "100038", "100123"
                bMode = False
                
            '구미점     도농점    만촌점    목동점    부평점
        Case "100030", "100197", "100012", "100264", "100259"
                bMode = False
                
            '비산점     상주점    서수원    성수점    송림점
        Case "100028", "100120", "100056", "100240", "100170"
                bMode = False
                
            '수지점     시화점    신월점    안성점    양주점
        Case "100128", "100031", "100009", "100222", "100143"
                bMode = False
                
            '울산학성점 원주점    월배점    은평점   청계천점
        Case "100021", "100008", "100015", "100023", "100211"
                bMode = False
                
            '칠성점     파주점    평택점    하남점   해운대점
        Case "100029", "100126", "100022", "100200", "100032"
                bMode = False
                
            '제천점
        Case "100331"
                bMode = False
                
        
        Case Else
                bMode = True
    End Select
    
    Check_이마트할인제외확인_20100302 = bMode
End Function

'--------------------------------------------------------------------------------
'새몸맞이 세일 행사
'> 기간은 2010.03.02 - 2010.03.15일 까지
Public Function Action_세일행사_20100302(sStartDate As String, sEndDate As String) As Boolean
    Dim dblPrice    As Double
    Dim FHandle     As Integer
    Dim nPercent    As Single
    
    On Error GoTo Check_할인_Error
    Action_세일행사_20100302 = False

    If Check_이마트할인제외확인_20100302 = True Then
        ' 해당 일자에 처음 실행할 경우 해당 파일이 없어서 실해되고 2번째부터는 실행되지 않는다.
        ' 해당 기간이 지나면 자동으로 프로그램에서 사용하지 않는다.
        If Dir(App.Path & "\" & sEndDate & ".TXT", vbDirectory) = "" Then
            ' 다음 이중 실행되지 않도록 파일을 생성한다.
            FHandle = FreeFile
            Open App.Path & "\" & sEndDate & ".TXT" For Append As FHandle
            Print #FHandle, Now

            
            ' 이전 자료를 모두 지운다.
            Query = "DELETE FROM 할인정보 WHERE 시작일 = '" & sStartDate & "' AND 종료일 = '" & sEndDate & "' "
            ADOCon.Execute Query
            
            '------------------------------------------------------------------------------------------------------------------
            ' 코드류(j), 점퍼류(d), 원피스류(t)
            '------------------------------------------------------------------------------------------------------------------
            nPercent = 0.75          '<- 25% 할인
            
            Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
            Query = Query & " WHERE  left(구분코드,1) = 'i' or left(구분코드,1) = 'd' or left(구분코드,1) = 't' "
            Set Rs = New ADODB.Recordset
            Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            Do While Not Rs.EOF
                If IsNumeric(Rs.Fields("가격")) = True Then
                    Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"; Tab; "["; CStr((1 - nPercent) * 100); "%]"

                    dblPrice = 0
                    dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)

                    Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                    Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                    Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                    ADOCon.Execute Query
                Else
                    Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                    Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                    MsgBox Query, vbCritical, "경고"
                End If
                Rs.MoveNext
            Loop
            Rs.Close
            
            
            '------------------------------------------------------------------------------------------------------------------
            ' 상의류(f)
            '------------------------------------------------------------------------------------------------------------------
            nPercent = 0.9          '<- 10% 할인   f17 직원조끼
            Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
            Query = Query & " WHERE  left(구분코드,1) = 'f' and 구분코드= 'f17'"
            Set Rs = New ADODB.Recordset
            Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            Do Until Rs.EOF
                If IsNumeric(Rs.Fields("가격")) = True Then
                    Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"; Tab; "["; CStr((1 - nPercent) * 100); "%]"

                    dblPrice = 0
                    dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)

                    Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                    Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                    Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                    ADOCon.Execute Query
                Else
                    Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                    Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                    MsgBox Query, vbCritical, "경고"
                End If
                Rs.MoveNext
            Loop
            Rs.Close
                            
            '------------------------------------------------------------------------------------------------------------------
            ' 나머지 상의류
            '------------------------------------------------------------------------------------------------------------------
            nPercent = 0.8          '<- 20% 할인   f17 직원조끼 제외
            Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
            Query = Query & " WHERE  left(구분코드,1) = 'f' and 구분코드<> 'f17'"
            Set Rs = New ADODB.Recordset
            Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            Do While Not Rs.EOF
                If IsNumeric(Rs.Fields("가격")) = True Then
                    Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"; Tab; "["; CStr((1 - nPercent) * 100); "%]"

                    dblPrice = 0
                    dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)

                    Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                    Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                    Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                    ADOCon.Execute Query
                Else
                    Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                    Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                    MsgBox Query, vbCritical, "경고"
                End If
                Rs.MoveNext
            Loop
            Rs.Close
            
            '------------------------------------------------------------------------------------------------------------------
            ' Y셔츠/T셔츠류(m)
            '------------------------------------------------------------------------------------------------------------------
            nPercent = 1          '<- 0% 할인   m00,m01,m20 제외
            Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
            Query = Query & " WHERE  left(구분코드,1) = 'm' and ( 구분코드= 'm00' or 구분코드= 'm01' or 구분코드= 'm20')"
            Set Rs = New ADODB.Recordset
            Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            Do While Not Rs.EOF
                If IsNumeric(Rs.Fields("가격")) = True Then
                    Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"; Tab; "["; CStr((1 - nPercent) * 100); "%]"

                    dblPrice = 0
                    dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)

                    Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                    Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                    Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                    ADOCon.Execute Query
                Else
                    Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                    Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                    MsgBox Query, vbCritical, "경고"
                End If
                Rs.MoveNext
            Loop
            Rs.Close
                            
            '------------------------------------------------------------------------------------------------------------------
            ' 나머지 상의류
            '------------------------------------------------------------------------------------------------------------------
            nPercent = 0.8          '<- 20% 할인   m00,m01,m20 제외한 나머지 품목
            Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
            Query = Query & " WHERE  left(구분코드,1) = 'm' and NOT ( 구분코드= 'm00' or 구분코드= 'm01' or 구분코드= 'm20')"
            Set Rs = New ADODB.Recordset
            Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            Do While Not Rs.EOF
                If IsNumeric(Rs.Fields("가격")) = True Then
                    Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"; Tab; "["; CStr((1 - nPercent) * 100); "%]"

                    dblPrice = 0
                    dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)

                    Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                    Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                    Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                    ADOCon.Execute Query
                Else
                    Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                    Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                    MsgBox Query, vbCritical, "경고"
                End If
                Rs.MoveNext
            Loop
            Rs.Close
            
            '------------------------------------------------------------------------------------------------------------------
            ' 스웨터류(u), 블라우스류(e)
            '------------------------------------------------------------------------------------------------------------------
            nPercent = 0.8          '<- 20% 할인
            
            Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
            Query = Query & " WHERE  left(구분코드,1) = 'u' or left(구분코드,1) = 'e' "
            Set Rs = New ADODB.Recordset
            Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            Do While Not Rs.EOF
                If IsNumeric(Rs.Fields("가격")) = True Then
                    Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"; Tab; "["; CStr((1 - nPercent) * 100); "%]"

                    dblPrice = 0
                    dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)

                    Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                    Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                    Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                    ADOCon.Execute Query
                Else
                    Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                    Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                    MsgBox Query, vbCritical, "경고"
                End If
                Rs.MoveNext
            Loop
            Rs.Close
            
            '------------------------------------------------------------------------------------------------------------------
            ' 이블류(k), 커텐류(j), 하의류(g), 스커트류(r), 커버류(z), 인형류(c), 한복류(h), 아동복류(q)
            '------------------------------------------------------------------------------------------------------------------
            nPercent = 0.85          '<- 15% 할인
            
            Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
            Query = Query & " WHERE  left(구분코드,1) = 'k' or left(구분코드,1) = 'j' or left(구분코드,1) = 'g' or left(구분코드,1) = 'r' or left(구분코드,1) = 'z' "
            Query = Query & " or     left(구분코드,1) = 'c' or left(구분코드,1) = 'h' or left(구분코드,1) = 'q'  "
            Set Rs = New ADODB.Recordset
            Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            Do While Not Rs.EOF
                If IsNumeric(Rs.Fields("가격")) = True Then
                    Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"; Tab; "["; CStr((1 - nPercent) * 100); "%]"

                    dblPrice = 0
                    dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)

                    Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                    Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                    Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                    ADOCon.Execute Query
                Else
                    Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                    Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                    MsgBox Query, vbCritical, "경고"
                End If
                Rs.MoveNext
            Loop
            Rs.Close
            
            '------------------------------------------------------------------------------------------------------------------
            ' 조끼(y)
            '------------------------------------------------------------------------------------------------------------------
            nPercent = 0.9          '<- 10% 할인
            
            Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
            Query = Query & " WHERE  left(구분코드,1) = 'y'  "
            Set Rs = New ADODB.Recordset
            Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            Do While Not Rs.EOF
                If IsNumeric(Rs.Fields("가격")) = True Then
                    Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"; Tab; "["; CStr((1 - nPercent) * 100); "%]"

                    dblPrice = 0
                    dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)

                    Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                    Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                    Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                    ADOCon.Execute Query
                Else
                    Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                    Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                    MsgBox Query, vbCritical, "경고"
                End If
                Rs.MoveNext
            Loop
            Rs.Close
            
            '------------------------------------------------------------------------------------------------------------------
            ' 넥타이(v)
            '------------------------------------------------------------------------------------------------------------------
            nPercent = 1          '<- 0% 할인   v00,v01,v02,v03,v04,v05,v06,v09,v10,v14,v15,v16,v17,v18 품목
            Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
            Query = Query & " WHERE  left(구분코드,1) = 'v' and "
            Query = Query & "        (구분코드= 'v00' or 구분코드= 'v01' or 구분코드= 'v02' or 구분코드= 'v03' or 구분코드= 'v04' or 구분코드= 'v05' "
            Query = Query & "      or 구분코드= 'v06' or 구분코드= 'v09' or 구분코드= 'v10' or 구분코드= 'v14' or 구분코드= 'v15' or 구분코드= 'v16' "
            Query = Query & "      or 구분코드= 'v17' or 구분코드= 'v18')"
            Set Rs = New ADODB.Recordset
            Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            Do While Not Rs.EOF
                If IsNumeric(Rs.Fields("가격")) = True Then
                    Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"; Tab; "["; CStr((1 - nPercent) * 100); "%]"

                    dblPrice = 0
                    dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)

                    Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                    Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                    Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                    ADOCon.Execute Query
                Else
                    Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                    Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                    MsgBox Query, vbCritical, "경고"
                End If
                Rs.MoveNext
            Loop
            Rs.Close
                            
            '------------------------------------------------------------------------------------------------------------------
            ' 나머지 상의류
            '------------------------------------------------------------------------------------------------------------------
            nPercent = 0.9          '<- 10% 할인   v00,v01,v02,v03,v04,v05,v06,v09,v10,v14,v15,v16,v17,v18을 제외한 나머지 품목
            Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
            Query = Query & " WHERE  left(구분코드,1) = 'v' and "
            Query = Query & "   NOT  (구분코드= 'v00' or 구분코드= 'v01' or 구분코드= 'v02' or 구분코드= 'v03' or 구분코드= 'v04' or 구분코드= 'v05' "
            Query = Query & "      or 구분코드= 'v06' or 구분코드= 'v09' or 구분코드= 'v10' or 구분코드= 'v14' or 구분코드= 'v15' or 구분코드= 'v16' "
            Query = Query & "      or 구분코드= 'v17' or 구분코드= 'v18')"
            Set Rs = New ADODB.Recordset
            Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            Do While Not Rs.EOF
                If IsNumeric(Rs.Fields("가격")) = True Then
                    Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"; Tab; "["; CStr((1 - nPercent) * 100); "%]"

                    dblPrice = 0
                    dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)

                    Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                    Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                    Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                    ADOCon.Execute Query
                Else
                    Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                    Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                    MsgBox Query, vbCritical, "경고"
                End If
                Rs.MoveNext
            Loop
            Rs.Close
            
            '------------------------------------------------------------------------------------------------------------------
            ' 가죽류(b)
            '------------------------------------------------------------------------------------------------------------------
            nPercent = 0.75          '<- 25% 할인   b27 인조가죽상의
            Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
            Query = Query & " WHERE  left(구분코드,1) = 'b' and 구분코드= 'b27' "
            Set Rs = New ADODB.Recordset
            Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            Do While Not Rs.EOF
                If IsNumeric(Rs.Fields("가격")) = True Then
                    Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"; Tab; "["; CStr((1 - nPercent) * 100); "%]"

                    dblPrice = 0
                    dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)

                    Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                    Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                    Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                    ADOCon.Execute Query
                Else
                    Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                    Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                    MsgBox Query, vbCritical, "경고"
                End If
                Rs.MoveNext
            Loop
            Rs.Close
                            
            '------------------------------------------------------------------------------------------------------------------
            ' 나머지 가죽류
            '------------------------------------------------------------------------------------------------------------------
            nPercent = 0.9          '<- 10% 할인   b27을 제외한 나머지 품목
            Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
            Query = Query & " WHERE  left(구분코드,1) = 'b' and 구분코드<> 'b27' "
            Set Rs = New ADODB.Recordset
            Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            Do While Not Rs.EOF
                If IsNumeric(Rs.Fields("가격")) = True Then
                    Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"; Tab; "["; CStr((1 - nPercent) * 100); "%]"

                    dblPrice = 0
                    dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)

                    Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                    Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                    Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                    ADOCon.Execute Query
                Else
                    Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                    Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                    MsgBox Query, vbCritical, "경고"
                End If
                Rs.MoveNext
            Loop
            Rs.Close
            
            '------------------------------------------------------------------------------------------------------------------
            ' 무스탕류(n)
            '------------------------------------------------------------------------------------------------------------------
            nPercent = 0.75          '<- 25% 할인   n25,n26,n27,n47,n48,n49,n50 품목
            Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
            Query = Query & " WHERE  left(구분코드,1) = 'n' and "
            Query = Query & "        (구분코드= 'n25' or 구분코드= 'n26' or 구분코드= 'n27' or 구분코드= 'n47' or 구분코드= 'n48' or 구분코드= 'n49' "
            Query = Query & "      or 구분코드= 'n50' )"
            Set Rs = New ADODB.Recordset
            Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            Do While Not Rs.EOF
                If IsNumeric(Rs.Fields("가격")) = True Then
                    Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"; Tab; "["; CStr((1 - nPercent) * 100); "%]"

                    dblPrice = 0
                    dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)

                    Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                    Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                    Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                    ADOCon.Execute Query
                Else
                    Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                    Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                    MsgBox Query, vbCritical, "경고"
                End If
                Rs.MoveNext
            Loop
            Rs.Close
                            
            '------------------------------------------------------------------------------------------------------------------
            ' 나머지 무스탕류
            '------------------------------------------------------------------------------------------------------------------
            nPercent = 0.9          '<- 10% 할인   n25,n26,n27,n47,n48,n49,n50 품목 제외
            Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
            Query = Query & " WHERE  left(구분코드,1) = 'n' and "
            Query = Query & "   NOT  (구분코드= 'n25' or 구분코드= 'n26' or 구분코드= 'n27' or 구분코드= 'n47' or 구분코드= 'n48' or 구분코드= 'n49' "
            Query = Query & "      or 구분코드= 'n50' )"
            Set Rs = New ADODB.Recordset
            Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            Do While Not Rs.EOF
                If IsNumeric(Rs.Fields("가격")) = True Then
                    Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"; Tab; "["; CStr((1 - nPercent) * 100); "%]"

                    dblPrice = 0
                    dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)

                    Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                    Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                    Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                    ADOCon.Execute Query
                Else
                    Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                    Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                    MsgBox Query, vbCritical, "경고"
                End If
                Rs.MoveNext
            Loop
            Rs.Close
            
            '------------------------------------------------------------------------------------------------------------------
            ' 카페트(x), 운동화(a)
            '------------------------------------------------------------------------------------------------------------------
            nPercent = 1          '<- 0% 할인   행사에서 제외
            Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
            Query = Query & " WHERE  left(구분코드,1) = 'x' or left(구분코드,1) = 'a' "
            Set Rs = New ADODB.Recordset
            Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            Do While Not Rs.EOF
                If IsNumeric(Rs.Fields("가격")) = True Then
                    Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"; Tab; "["; CStr((1 - nPercent) * 100); "%]"

                    dblPrice = 0
                    dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)

                    Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                    Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                    Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                    ADOCon.Execute Query
                Else
                    Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                    Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                    MsgBox Query, vbCritical, "경고"
                End If
                Rs.MoveNext
            Loop
            Rs.Close
            
            '------------------------------------------------------------------------------------------------------------------
            ' 소품기타(s)
            '------------------------------------------------------------------------------------------------------------------
            nPercent = 0.75          '<- 25% 할인   s20,s21,s22,s23,s27,s28,s29,s30,s35,s36,s37 품목
            Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
            Query = Query & " WHERE  left(구분코드,1) = 's' and "
            Query = Query & "        (구분코드= 's20' or 구분코드= 's21' or 구분코드= 's22' or 구분코드= 's23' or 구분코드= 's27' or 구분코드= 's28' "
            Query = Query & "      or 구분코드= 's29' or 구분코드= 's30' or 구분코드= 's35' or 구분코드= 's36' or 구분코드= 's37' )"
            Set Rs = New ADODB.Recordset
            Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            Do While Not Rs.EOF
                If IsNumeric(Rs.Fields("가격")) = True Then
                    Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"; Tab; "["; CStr((1 - nPercent) * 100); "%]"

                    dblPrice = 0
                    dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)

                    Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                    Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                    Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                    ADOCon.Execute Query
                Else
                    Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                    Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                    MsgBox Query, vbCritical, "경고"
                End If
                Rs.MoveNext
            Loop
            Rs.Close
            
            
            nPercent = 0.8          '<- 20% 할인   s25,s32,s34 품목
            Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
            Query = Query & " WHERE  left(구분코드,1) = 's' and "
            Query = Query & "        (구분코드= 's25' or 구분코드= 's32' or 구분코드= 's34' )"
            Set Rs = New ADODB.Recordset
            Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            Do While Not Rs.EOF
                If IsNumeric(Rs.Fields("가격")) = True Then
                    Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"; Tab; "["; CStr((1 - nPercent) * 100); "%]"

                    dblPrice = 0
                    dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)

                    Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                    Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                    Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                    ADOCon.Execute Query
                Else
                    Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                    Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                    MsgBox Query, vbCritical, "경고"
                End If
                Rs.MoveNext
            Loop
            Rs.Close
            
            
            nPercent = 0.85          '<- 15% 할인   s26,s33 품목
            Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
            Query = Query & " WHERE  left(구분코드,1) = 's' and "
            Query = Query & "        (구분코드= 's26' or 구분코드= 's33' )"
            Set Rs = New ADODB.Recordset
            Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            Do While Not Rs.EOF
                If IsNumeric(Rs.Fields("가격")) = True Then
                    Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"; Tab; "["; CStr((1 - nPercent) * 100); "%]"

                    dblPrice = 0
                    dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)

                    Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                    Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                    Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                    ADOCon.Execute Query
                Else
                    Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                    Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                    MsgBox Query, vbCritical, "경고"
                End If
                Rs.MoveNext
            Loop
            Rs.Close
            
            
            nPercent = 0.9          '<- 10% 할인   s24,s31 품목
            Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
            Query = Query & " WHERE  left(구분코드,1) = 's' and "
            Query = Query & "        (구분코드= 's24' or 구분코드= 's31' )"
            Set Rs = New ADODB.Recordset
            Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            Do While Not Rs.EOF
                If IsNumeric(Rs.Fields("가격")) = True Then
                    Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"; Tab; "["; CStr((1 - nPercent) * 100); "%]"

                    dblPrice = 0
                    dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)

                    Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                    Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                    Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                    ADOCon.Execute Query
                Else
                    Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                    Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                    MsgBox Query, vbCritical, "경고"
                End If
                Rs.MoveNext
            Loop
            Rs.Close
            
                            
            ' 나머지 소품기타
            nPercent = 1          '<- 0% 할인 s20,s21,s22,s23,s27,s28,s29,s30,s35,s36,s37 s25,s32,s34 s26,s33 s24,s31 품목 제외
            Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
            Query = Query & " WHERE  left(구분코드,1) = 's' and "
            Query = Query & "   NOT  (구분코드= 's20' or 구분코드= 's21' or 구분코드= 's22' or 구분코드= 's23' or 구분코드= 's27' or 구분코드= 's28' "
            Query = Query & "      or 구분코드= 's29' or 구분코드= 's30' or 구분코드= 's35' or 구분코드= 's36' or 구분코드= 's37' "
            Query = Query & "      or 구분코드= 's25' or 구분코드= 's32' or 구분코드= 's34' "
            Query = Query & "      or 구분코드= 's26' or 구분코드= 's33'  "
            Query = Query & "      or 구분코드= 's24' or 구분코드= 's31' )"
            Set Rs = New ADODB.Recordset
            Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            Do While Not Rs.EOF
                If IsNumeric(Rs.Fields("가격")) = True Then
                    Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"; Tab; "["; CStr((1 - nPercent) * 100); "%]"

                    dblPrice = 0
                    dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)

                    Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                    Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                    Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                    ADOCon.Execute Query
                Else
                    Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                    Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                    MsgBox Query, vbCritical, "경고"
                End If
                Rs.MoveNext
            Loop
            Rs.Close
            '------------------------------------------------------------------------------------------------------------------
            
            Close #FHandle
        End If
    End If
    
    Action_세일행사_20100302 = True

    On Error GoTo 0
    Exit Function

Check_할인_Error:
'    Resume
    Action_세일행사_20100302 = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Action_세일행사_20100302 of Module Sale"
End Function


'--------------------------------------------------------------------------------
'새몸맞이 세일 행사
'> 기간은 2010.03.04 - 2010.03.17일 까지
Public Function Action_세일행사_20100304(sStartDate As String, sEndDate As String) As Boolean
    Dim dblPrice    As Double
    Dim Rs          As Recordset
    Dim FHandle     As Integer
    Dim nPercent    As Single
    
    On Error GoTo Check_할인_Error
    Action_세일행사_20100304 = False


    ' 03-02일자 제외된 이마트 매장만 실시
    If Check_이마트할인제외확인_20100302 = False Then

            
        ' 해당 일자에 처음 실행할 경우 해당 파일이 없어서 실해되고 2번째부터는 실행되지 않는다.
        ' 해당 기간이 지나면 자동으로 프로그램에서 사용하지 않는다.
        If Dir(App.Path & "\" & sEndDate & ".TXT", vbDirectory) = "" Then
            ' 다음 이중 실행되지 않도록 파일을 생성한다.
            FHandle = FreeFile
            Open App.Path & "\" & sEndDate & ".TXT" For Append As FHandle
            Print #FHandle, Now
            
            ' 이전 자료를 모두 지운다.
            Query = "DELETE FROM 할인정보 WHERE 시작일 = '" & sStartDate & "' AND 종료일 = '" & sEndDate & "' "
            ADOCon.Execute Query
            
            '------------------------------------------------------------------------------------------------------------------
            ' 전품목
            nPercent = 0.8          '<- 20% 할인
            
            Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
            Set Rs = New ADODB.Recordset
            Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            Do While Not Rs.EOF
                If IsNumeric(Rs.Fields("가격")) = True Then
                    Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"; Tab; "["; CStr((1 - nPercent) * 100); "%]"

                    dblPrice = 0
                    dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)

                    Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                    Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                    Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                    ADOCon.Execute Query
                Else
                    Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                    Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                    MsgBox Query, vbCritical, "경고"
                End If
                Rs.MoveNext
            Loop
            Rs.Close
            
            Close #FHandle
        End If

    End If
    
    Action_세일행사_20100304 = True

    On Error GoTo 0
    Exit Function

Check_할인_Error:
    Resume
    Action_세일행사_20100304 = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Action_세일행사_20100304 of Module Sale"
End Function


'--------------------------------------------------------------------------------
'새몸맞이 세일 행사
'> 기간은 2010.03.02 - 2010.03.15일 까지
Public Function Action_세일행사_20100302_01(sStartDate As String, sEndDate As String) As Boolean
    Dim dblPrice    As Double
    Dim Rs          As Recordset
    Dim FHandle     As Integer
    Dim nPercent    As Single
    
    On Error GoTo Check_할인_Error
    Action_세일행사_20100302_01 = False



    If Check_이마트할인제외확인_20100302 = True Then


        ' 해당 일자에 처음 실행할 경우 해당 파일이 없어서 실해되고 2번째부터는 실행되지 않는다.
        ' 해당 기간이 지나면 자동으로 프로그램에서 사용하지 않는다.
        If Dir(App.Path & "\" & sEndDate & "_01.TXT", vbDirectory) = "" Then
            ' 다음 이중 실행되지 않도록 파일을 생성한다.
            FHandle = FreeFile
            Open App.Path & "\" & sEndDate & "_01.TXT" For Append As FHandle
            Print #FHandle, Now

            
            ' 이전 자료를 모두 지운다.
            Query = "DELETE FROM 할인정보 WHERE 시작일 = '" & sStartDate & "' AND 종료일 = '" & sEndDate & "'   and "
            Query = Query & "        (구분코드= 'v00' or 구분코드= 'v01' or 구분코드= 'v02' or 구분코드= 'v03' or 구분코드= 'v04' or 구분코드= 'v05' "
            Query = Query & "      or 구분코드= 'v06' or 구분코드= 'v09' or 구분코드= 'v10' or 구분코드= 'v14' or 구분코드= 'v15' or 구분코드= 'v16' "
            Query = Query & "      or 구분코드= 'v17' or 구분코드= 'v18')"
            ADOCon.Execute Query
            
  

            '------------------------------------------------------------------------------------------------------------------
            ' 넥타이(v)
            nPercent = 0.9          '<- 10% 할인   v00,v01,v02,v03,v04,v05,v06,v09,v10,v14,v15,v16,v17,v18 품목
            Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
            Query = Query & " WHERE  left(구분코드,1) = 'v' and "
            Query = Query & "        (구분코드= 'v00' or 구분코드= 'v01' or 구분코드= 'v02' or 구분코드= 'v03' or 구분코드= 'v04' or 구분코드= 'v05' "
            Query = Query & "      or 구분코드= 'v06' or 구분코드= 'v09' or 구분코드= 'v10' or 구분코드= 'v14' or 구분코드= 'v15' or 구분코드= 'v16' "
            Query = Query & "      or 구분코드= 'v17' or 구분코드= 'v18')"
            Set Rs = New ADODB.Recordset
            Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            Do While Not Rs.EOF
                If IsNumeric(Rs.Fields("가격")) = True Then
                    Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"; Tab; "["; CStr((1 - nPercent) * 100); "%]"

                    dblPrice = 0
                    dblPrice = (CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)

                    Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                    Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                    Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                    ADOCon.Execute Query
                Else
                    Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                    Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                    MsgBox Query, vbCritical, "경고"
                End If
                
                Rs.MoveNext
            Loop
            Rs.Close
 
            
            Close #FHandle
        End If
    End If
    
    Action_세일행사_20100302_01 = True

    On Error GoTo 0
    
    Exit Function

Check_할인_Error:
'    Resume
    Action_세일행사_20100302_01 = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Action_세일행사_20100302_01 of Module Sale"
End Function


'--------------------------------------------------------------------------------
'새몸맞이 세일 행사
'> 기간은 2010.03.02 - 2010.03.15일 까지
Public Function Action_세일행사_20100302_02(sStartDate As String, sEndDate As String) As Boolean
    Dim dblPrice    As Double
    Dim Rs          As Recordset
    Dim FHandle     As Integer
    Dim nPercent    As Single
    
    On Error GoTo Check_할인_Error
    Action_세일행사_20100302_02 = False



    If Check_이마트할인제외확인_20100302 = True Then


        ' 해당 일자에 처음 실행할 경우 해당 파일이 없어서 실해되고 2번째부터는 실행되지 않는다.
        ' 해당 기간이 지나면 자동으로 프로그램에서 사용하지 않는다.
        If Dir(App.Path & "\" & sEndDate & "_02.TXT", vbDirectory) = "" Then
            ' 다음 이중 실행되지 않도록 파일을 생성한다.
            FHandle = FreeFile
            Open App.Path & "\" & sEndDate & "_02.TXT" For Append As FHandle
            Print #FHandle, Now

            
            ' 이전 자료를 모두 지운다.
            Query = "DELETE FROM 할인정보 WHERE 시작일 = '" & sStartDate & "' AND 종료일 = '" & sEndDate & "'   and  left(구분코드,1) = 'a' "
            ADOCon.Execute Query
            
  

            '------------------------------------------------------------------------------------------------------------------
            ' 운동화(v)
            nPercent = 1
            Query = "SELECT 구분코드, 품명, 가격 FROM 참조코드 "
            Query = Query & " WHERE  left(구분코드,1) = 'a'  "
            Set Rs = New ADODB.Recordset
            Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            Do While Not Rs.EOF
                If IsNumeric(Rs.Fields("가격")) = True Then
                    Print #FHandle, "["; Rs.Fields("구분코드"); ":"; Rs.Fields("품명"); Tab; Tab; ":"; Rs.Fields("가격"); "]"; Tab; "["; CStr((Val(CStr(Rs.Fields("가격"))) * nPercent)); "]"; Tab; "["; CStr((CLng((Val(CStr(Rs.Fields("가격"))) * nPercent) * 0.01) * 100)); "]"; Tab; "["; CStr((1 - nPercent) * 100); "%]"

                    dblPrice = 0
                    dblPrice = Val(Rs.Fields("가격"))

                    Query = "INSERT INTO 할인정보 (시작일, 종료일, 구분코드, 품명, 가격, 비율 ) "
                    Query = Query & " VALUES ('" & sStartDate & "', '" & sEndDate & "', '" & Rs.Fields("구분코드") & "', '"
                    Query = Query & Rs.Fields("품명") & "', '" & CStr(dblPrice) & "', '2') "
                    ADOCon.Execute Query
                Else
                    Query = "등록정 상품코드가 올바르지 않습니다. 확인하여 주십시요" & vbLf & vbLf
                    Query = Query & "[" & Rs.Fields("구분코드") & ":" & Rs.Fields("품명") & ":" & Rs.Fields("가격") & "]"
                    MsgBox Query, vbCritical, "경고"
                End If
                Rs.MoveNext
            Loop
            Rs.Close
 
            
            Close #FHandle
        End If
    End If
    
    Action_세일행사_20100302_02 = True

    On Error GoTo 0
    Exit Function

Check_할인_Error:
'    Resume
    Action_세일행사_20100302_02 = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Action_세일행사_20100302_02 of Module Sale"
End Function
