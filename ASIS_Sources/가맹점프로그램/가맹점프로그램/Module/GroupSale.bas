Attribute VB_Name = "GroupSale"
Option Explicit

'할인 상태
Public Enum uGSS_Enum
    s2세트할인 = 2
    s3세트할인 = 3
    s4세트할인 = 4
    s5세트할인 = 5
    s6세트할인 = 6
End Enum
Dim GSStats As uGSS_Enum

Type GroupSalePencent_TYPE
    s2세트할인률 As Single
    s3세트할인률 As Single
    s4세트할인률 As Single
    s5세트할인률 As Single
    s6세트할인률 As Single
End Type
Dim sPencent As GroupSalePencent_TYPE

Type GroupSaleGoodsCode_Type
    Set2Code1 As String
    Set2Code2 As String
    Set3Code1 As String
    Set4Code1 As String
    Set5Code1 As String
    Set6Code1 As String
End Type
Dim gsgCode As GroupSaleGoodsCode_Type

Type GroupSaleItem_TYPE
    Tag         As String   '택번호
    Code        As String   '상품 코드
    PriceOrg    As Double   '원 금액
    PriceEnd    As Double   '할인후 금액
End Type

Type GroupSaleMain_Type
    SaleGubun       As uGSS_Enum
    SalePercent     As GroupSalePencent_TYPE
    Goods()         As GroupSaleItem_TYPE
End Type

Dim GSetINFO()    As GroupSaleMain_Type

Type GroupSaleMoney_TYPE
    d세트Key            As String
    d2세트수량          As Integer
    d3세트수량          As Integer
    d4세트수량          As Integer
    d5세트수량          As Integer
    d6세트수량          As Integer
    d세트수량합계       As Integer
    d무료세탁권수량     As Integer
    d전체금액           As Double
    d세트금액           As Double
    d세트할인금액       As Double
    d에누리할인금액     As Double
    d전체할인금액       As Double
    d최종수령액         As Double
End Type
Public m_GSGMoney As GroupSaleMoney_TYPE

Public m_세트응모번호수량 As Integer
Public m_세트응모번호() As String

Public Const m_SetCol1 = 10    ' A~Z
Public Const m_SetCol2 = 11    ' ex. 6-01, 5-01, 5-02
Public Const m_SetCol3 = 12    ' 세트 할인률을 기준으로 계산한 금액(10원단위 포함)
Public Const m_SetCol4 = 13    ' 세트 원단위 포함한 합계 금액에서 원단위 절사후 다시 계산한 금액
Public Const m_SetCol5 = 14    ' 원가격 기록

Public Const m_SetDefPrice = 5  ' 가격 자료
Public Const m_SetDefCode = 7   ' 상품 코드

Public Sub DefSet(sMode As String)
    ZeroMemory m_GSGMoney, Len(m_GSGMoney)
    
    sPencent.s2세트할인률 = 0.03
    sPencent.s3세트할인률 = 0.04
    sPencent.s4세트할인률 = 0.05
    sPencent.s5세트할인률 = 0.06
    sPencent.s6세트할인률 = 0.07
 
    gsgCode.Set2Code1 = "f00,f01,f02,f03,f04,f05,f06,f07,f08,f14,f15,f20,f22,f30"
    gsgCode.Set2Code2 = "g00,g04,g05,g07,g15,g16,g17,g23,g24,g28,g35,r00,r01,r02,r03,r04,r05,r06,r07,r08,r09,r10,r11,r12,r13,r14,r15,r16,r17,r18,r19,r20,r21,r22,r23,r24,r25,r26,r27,r28,r29,r30,r31,r32,r33,r34,r35,r36,r37,r38,r39,r40,r41,r42,r43,r44,r45,r46,r47,r48,r49,r50,r51,r52,r53,r54,r55,r56,r57,r58,r59,r60,r61,r62,r63,r64,r65,r66,r67,r68,r69,r70,r71,r72,r73,r74,r75,r76,r77,r78,r79,r80,r81,r82,r83,r84,r85,r86,r87,r88,r89,r90,r91,r92,r93,r94,r95,r96,r97,r98,r99"
    gsgCode.Set3Code1 = "i00,i02,i05,i06,i07,i10,i11,i12,i16,i19,i20,i26"
    gsgCode.Set4Code1 = "i01,i04,i08,i09,i14,i15,i17,i18,i21,i23,i25"
    gsgCode.Set5Code1 = "d00,d01,d02,d03,d04,d05,d06,d07,d08,d09,d10,d11,d12,d13,d14,d15,d16,d17,d18,d19,d20,d21,d22,d23,d24,d25,d26,d27,d28,d29,d30,d31,d32,d33,d34,d35,d36,d37,d38,d39,d40,d41,d42,d43,d44,d45,d46,d47,d48,d49,d50,d51,d52,d53,d54,d55,d56,d57,d58,d59,d60,d61,d62,d63,d64,d65,d66,d67,d68,d69,d70,d71,d72,d73,d74,d75,d76,d77,d78,d79,d80,d81,d82,d83,d84,d85,d86,d87,d88,d89,d90,d91,d92,d93,d94,d95,d96,d97,d98,d99"
    gsgCode.Set6Code1 = "t00,t01,t02,t03,t04,t05,t06,t07,t08,t09,t10,t11,t12,t13,t14,t15,t16,t17,t18,t19,t20,t21,t22,t23,t24,t25,t26,t27,t28,t29,t30,t31,t32,t33,t34,t35,t36,t37,t38,t39,t40,t41,t42,t43,t44,t45,t46,t47,t48,t49,t50,t51,t52,t53,t54,t55,t56,t57,t58,t59,t60,t61,t62,t63,t64,t65,t66,t67,t68,t69,t70,t71,t72,t73,t74,t75,t76,t77,t78,t79,t80,t81,t82,t83,t84,t85,t86,t87,t88,t89,t90,t91,t92,t93,t94,t95,t96,t97,t98,t99"
End Sub

Public Function Check세트상품확인(MySpr As fpSpread) As Boolean
    Call DefSet("")
    
    Call GetSetSpreedDef(MySpr)      ' 스프레드의 내용을 초기화 한다.
    Call GroupSaleCheck(MySpr)       ' 1. 해당 세트 구성 상품을 확인하고 A~Z까지 기록하여 세트 구성 내용을 확인한다. (Col = m_SetCol1, Col = m_SetCol2)
    Call GroupSalePriceCheck(MySpr)  ' 2. 해당 세트 구성 상품의 가격을 기록한다. (Col = m_SetCol3 , Col = m_SetCol5)
    Call GroupSalePriceCheck2(MySpr) ' 3. 전체 금액을 계산하여 다시 십원단위 절사하여 다시 계산한다. (Col = m_SetCol4)
End Function


'------------------------------------------------------------------------------------------
' 전달된 내용의 세트 상품이 있을 경우 해당하는 문자를 기록한다. (nSetStr : A~Z)
Private Sub GroupSaleCheck(MySpr As fpSpread)
    Dim lRow    As Long
    Dim nSetStr As Integer
    
    '전달된 내용의 세트 상품이 있을 경우 해당하는 문자를 기록한다. (nSetStr : A~Z)
    For nSetStr = 65 To 90      ' A~Z
        ' 2세트중 첫번째 상품이 있는지 확인한다.
        If GetSetGoodsCheck(MySpr, gsgCode.Set2Code1, nSetStr) = False Then
        
        ' 2세트중 두번째 상품이 있는지 확인한다.
        ElseIf GetSetGoodsCheck(MySpr, gsgCode.Set2Code2, nSetStr) = False Then
            
            ' 2번째 상품이 없다는 것은 첫번째 상품만 표시가 되었다는 말이기 때문에
            ' 첫번째 상품에 해당 표시가 있는 부분을 삭제 하여 준다.
            For lRow = 1 To MySpr.MaxRows
                MySpr.Row = lRow:  MySpr.Col = m_SetCol1
                
                '기존 세트에 설정되지 않는 내역만을 세트 확인한다.
                If Trim(MySpr.Value) = Chr(nSetStr) Then
                    MySpr.Text = "" 'nSetStr
                    nSetStr = 90
                    Exit For
                End If
            Next lRow
        
        ElseIf GetSetGoodsCheck(MySpr, gsgCode.Set3Code1, nSetStr) = False Then
            ' 3세트 상품이 있는지 확인한다.
        
        ElseIf GetSetGoodsCheck(MySpr, gsgCode.Set4Code1, nSetStr) = False Then
            ' 4세트 상품이 있는지 확인한다.
        
        ElseIf GetSetGoodsCheck(MySpr, gsgCode.Set5Code1, nSetStr) = False Then
            ' 5세트 상품이 있는지 확인한다.
        
        ElseIf GetSetGoodsCheck(MySpr, gsgCode.Set6Code1, nSetStr) = False Then
            ' 6세트 상품이 있는지 확인한다.
        End If
    Next nSetStr
    
    '---------------------------------------------------------------------------------
    'A~Z까지 기록된 수를 구하여 세트 구성을 확정한다.
    Dim sSetCode    As String
    
    Dim nSetCount(65 To 90) As Integer
    Dim nSetAZCnt(2 To 6) As Integer
    
    ' 해당 세트의 수를 구한다.
    ' A: nSetCount(65) = 4 세트 구성
    ' B: nSetCount(66) = 3 세트 구성
    ' C: nSetCount(67) = 3 세트 구성
    ' D: nSetCount(68) = 2 세트 구성
    
    Erase nSetCount
    
    With MySpr
        For nSetStr = 65 To 90  ' A~Z
            For lRow = 1 To .MaxRows
                .Row = lRow
                
                .Col = m_SetCol1: sSetCode = Trim(.Value)
                
                If sSetCode = Chr(nSetStr) Then
                    nSetCount(Asc(sSetCode)) = nSetCount(Asc(sSetCode)) + 1
                End If
            Next lRow
        Next nSetStr
    End With
   
    
    '---------------------------------------------------------------------------------
    'A~Z까지 기록된 수를 구하여 세트 구성을 확정한다.
    With MySpr
        nSetAZCnt(uGSS_Enum.s2세트할인) = 1
        nSetAZCnt(uGSS_Enum.s3세트할인) = 1
        nSetAZCnt(uGSS_Enum.s4세트할인) = 1
        nSetAZCnt(uGSS_Enum.s5세트할인) = 1
        nSetAZCnt(uGSS_Enum.s6세트할인) = 1
        
        ' 해당 세트의 수를 확인하여 순번을 정한다.
        For nSetStr = 65 To 90      ' A~Z
            ' 6세트 할인
            If nSetCount(nSetStr) = uGSS_Enum.s6세트할인 Then
                Call SetGoodsINOF_DATA(MySpr, nSetCount, nSetStr, s6세트할인, nSetAZCnt(uGSS_Enum.s6세트할인))
                nSetAZCnt(uGSS_Enum.s6세트할인) = nSetAZCnt(uGSS_Enum.s6세트할인) + 1
            
            ' 5세트 할인
            ElseIf nSetCount(nSetStr) = uGSS_Enum.s5세트할인 Then
                Call SetGoodsINOF_DATA(MySpr, nSetCount, nSetStr, s5세트할인, nSetAZCnt(uGSS_Enum.s5세트할인))
                nSetAZCnt(uGSS_Enum.s5세트할인) = nSetAZCnt(uGSS_Enum.s5세트할인) + 1
            
            ' 4세트 할인
            ElseIf nSetCount(nSetStr) = uGSS_Enum.s4세트할인 Then
                Call SetGoodsINOF_DATA(MySpr, nSetCount, nSetStr, s4세트할인, nSetAZCnt(uGSS_Enum.s4세트할인))
                nSetAZCnt(uGSS_Enum.s4세트할인) = nSetAZCnt(uGSS_Enum.s4세트할인) + 1
            
            ' 3세트 할인
            ElseIf nSetCount(nSetStr) = uGSS_Enum.s3세트할인 Then
                Call SetGoodsINOF_DATA(MySpr, nSetCount, nSetStr, s3세트할인, nSetAZCnt(uGSS_Enum.s3세트할인))
                nSetAZCnt(uGSS_Enum.s3세트할인) = nSetAZCnt(uGSS_Enum.s3세트할인) + 1
            
            ' 2세트 할인
            ElseIf nSetCount(nSetStr) = uGSS_Enum.s2세트할인 Then
                Call SetGoodsINOF_DATA(MySpr, nSetCount, nSetStr, s2세트할인, nSetAZCnt(uGSS_Enum.s2세트할인))
                nSetAZCnt(uGSS_Enum.s2세트할인) = nSetAZCnt(uGSS_Enum.s2세트할인) + 1
            End If
        Next nSetStr
        
        ' 전세 수량을 계산하여 기록한다.
        m_GSGMoney.d2세트수량 = nSetAZCnt(uGSS_Enum.s2세트할인) - 1
        m_GSGMoney.d3세트수량 = nSetAZCnt(uGSS_Enum.s3세트할인) - 1
        m_GSGMoney.d4세트수량 = nSetAZCnt(uGSS_Enum.s4세트할인) - 1
        m_GSGMoney.d5세트수량 = nSetAZCnt(uGSS_Enum.s5세트할인) - 1
        m_GSGMoney.d6세트수량 = nSetAZCnt(uGSS_Enum.s6세트할인) - 1
        m_GSGMoney.d세트수량합계 = (m_GSGMoney.d2세트수량 + m_GSGMoney.d3세트수량 + m_GSGMoney.d4세트수량 + m_GSGMoney.d5세트수량 + m_GSGMoney.d6세트수량)
        
        m_GSGMoney.d무료세탁권수량 = 0
        
'        m_GSGMoney.d무료세탁권수량 = (m_GSGMoney.d2세트수량 * 1) + _
'                                     (m_GSGMoney.d3세트수량 * 2) + _
'                                     (m_GSGMoney.d4세트수량 * 3) + _
'                                     (m_GSGMoney.d5세트수량 * 4) + _
'                                     (m_GSGMoney.d6세트수량 * 5)
        
    End With
End Sub

'------------------------------------------------------------------------------------------
' 전달된 내용의 세트 상품이 있을 경우 해당하는 문자를 기록한다. (nSetStr : A~Z)
Private Sub GetSetSpreedDef(MySpr As fpSpread)
    Dim lRow As Long
    
    With MySpr
        ' 2세트중 첫번째 상품이 있는지 확인한다.
        For lRow = 1 To .MaxRows
            .Row = lRow
        
            .Col = m_SetCol1: .Text = ""
            .Col = m_SetCol2: .Text = ""
            .Col = m_SetCol3: .Text = ""
            .Col = m_SetCol4: .Text = ""
        Next lRow
    End With
End Sub


'------------------------------------------------------------------------------------------
' 전달된 내용의 세트 상품이 있을 경우 해당하는 문자를 기록한다. (nSetStr : A~Z)
Private Function GetSetGoodsCheck(MySpr As fpSpread, sSetCode As String, nSetStr As Integer) As Boolean
    Dim lRow As Long
    Dim strCode As String
    
    GetSetGoodsCheck = False
                
    With MySpr
        ' 2세트중 첫번째 상품이 있는지 확인한다.
        For lRow = 1 To .MaxRows
            .Row = lRow
            .Col = m_SetCol1
            
            '기존 세트에 설정되지 않는 내역만을 세트 확인한다.
            If Trim(.Value) = "" Then
                .Col = m_SetDefCode: strCode = .Value
                
                ' 마지막 제품일 경우
                If strCode = "" Then Exit For
                
                ' 2세트 상품의 구성이 있을 경우
                If InStr(sSetCode, strCode) > 0 Then
                    .Col = m_SetCol1
                    .Text = Chr(nSetStr)
                    GetSetGoodsCheck = True
                    Exit For
                End If
            End If
        Next lRow
    End With
End Function

Private Sub SetGoodsINOF_DATA(MySpr As fpSpread, nSetCount() As Integer, nSetStr As Integer, gStats As uGSS_Enum, nCnt As Integer)
    Dim lRow      As Long
    Dim sSetCode  As String
    
    With MySpr
        If nSetCount(nSetStr) = gStats Then
            For lRow = 1 To .MaxRows
                .Row = lRow: .Col = m_SetDefCode ' 마지막 상품 확인
                
                If .Text = "" Then Exit For
                
                .Col = m_SetCol1: sSetCode = Trim(.Value)
                
                If sSetCode = Chr(nSetStr) Then
                    .Col = m_SetCol2
                    .Text = CStr(gStats) & "-" & Format(nCnt, "00")
                End If
            Next lRow
        End If
    End With
End Sub

'------------------------------------------------------------------------------------------
'  기록된 값을 가지고 해당 세트의 수를 리턴한다.
Private Function GetSetGoodsCount(nSetData() As Integer, nSelect As Integer) As Integer
    Dim Index   As Integer
    Dim nCnt    As Integer
    
    nCnt = 0
    
    For Index = Chr(65) To Chr(90)  ' A~Z
        If nSetData(Index) = nSelect Then nCnt = nCnt + 1
    Next Index
    
    GetSetGoodsCount = nCnt
End Function


'------------------------------------------------------------------------------------------
'  2. 해당 세트 구성 상품의 가격을 기록한다.
Private Sub GroupSalePriceCheck(MySpr As fpSpread)
    Dim lRow        As Long
    Dim nPercent    As Single
    Dim sSetCode    As String
    Dim dblPrice    As Double
    
    With MySpr
    
        '원가격을 보관한다.
        For lRow = 1 To .MaxRows
            .Row = lRow
            
            .Col = m_SetDefCode
            If .Value = "" Then Exit For
            
            .Col = m_SetDefPrice: dblPrice = Val(Replace(Trim(.Value), ",", ""))
            
            m_GSGMoney.d전체금액 = m_GSGMoney.d전체금액 + dblPrice
                        
            .Col = m_SetCol5: .Text = dblPrice
        Next lRow
        
        For lRow = 1 To .MaxRows
            .Row = lRow
            .Col = m_SetDefCode
            If .Value = "" Then Exit For
            
            .Col = m_SetCol2: sSetCode = Trim(.Value) ' 세트 상품인지 확인한다.
            
            If sSetCode <> "" Then
                nPercent = Choose(Val(Left(sSetCode, 1)), 0, sPencent.s2세트할인률, sPencent.s3세트할인률, sPencent.s4세트할인률, sPencent.s5세트할인률, sPencent.s6세트할인률)
                
                .Col = m_SetDefPrice: dblPrice = Val(Replace(Trim(.Value), ",", "")) ' 가격
                .Col = m_SetCol3:     .Text = CLng(dblPrice * (1 - nPercent))        '
                .Col = m_SetCol4:     .Text = CStr(CLng(dblPrice * (1 - nPercent)))  '
                
                m_GSGMoney.d세트금액 = m_GSGMoney.d세트금액 + CLng(dblPrice * (1 - nPercent))
                m_GSGMoney.d세트할인금액 = m_GSGMoney.d세트할인금액 + (dblPrice - CLng(dblPrice * (1 - nPercent))) ' 전체 금액과 세트금액의 차이를 구한다.
            End If
        Next lRow
    End With
End Sub

'------------------------------------------------------------------------------------------
'  2. 해당 세트 구성 상품의 가격을 기록한다.
Private Sub GroupSalePriceCheck2(MySpr As fpSpread)
    Dim lRow        As Long
    Dim sSetCode    As String
    
    Dim dblTotalPrice   As Double
    
    Dim LBigRow         As Long
    Dim dblBigPrice     As Double
    Dim dblCutMoney     As Double
    Dim dbltempPrice    As Double
    
    dblTotalPrice = 0
    
    ' 전체 금액을 계산한다.
    With MySpr
        For lRow = 1 To .MaxRows
            .Row = lRow
            .Col = m_SetDefCode
            
            ' 마지막 상품 확인
            If .Text = "" Then Exit For
            
            ' 세트 상품인지 확인한다.
            .Col = m_SetCol2: sSetCode = Trim(Replace(.Value, ",", ""))
            
            If sSetCode <> "" Then
                .Col = m_SetCol3: dblTotalPrice = dblTotalPrice + Trim(Replace(.Value, ",", ""))
                
                '원단위 절사를 위하여
                If dblBigPrice <= Trim(Replace(.Value, ",", "")) Then
                    LBigRow = lRow
                    dblBigPrice = Trim(Replace(.Value, ",", ""))
                End If
                
            Else
                .Col = m_SetDefPrice: dblTotalPrice = dblTotalPrice + Trim(Replace(.Value, ",", "")) ' 평상시 가격
            End If
        Next lRow
        
        ' 10원단위 금액을 구한다.
        dblCutMoney = dblTotalPrice Mod 100
        m_GSGMoney.d에누리할인금액 = dblCutMoney
        
        ' 가장 큰 금액에서 원단위 절사 합의함
        .Row = LBigRow
        .Col = m_SetCol4: .Value = Val(.Value) - dblCutMoney
        
        With m_GSGMoney
            .d전체할인금액 = .d에누리할인금액 + .d세트할인금액
            .d최종수령액 = .d전체금액 - .d전체할인금액
        End With
        
        ' 가격을 표시해 준다.
        For lRow = 1 To .MaxRows
            .Row = lRow
            .Col = m_SetDefCode
            
            ' 마지막 상품 확인
            If .Text = "" Then Exit For
            
            ' 세트 상품인지 확인한다.
            .Col = m_SetCol2
            If .Value <> "" Then
                .Col = m_SetCol4:     dbltempPrice = Val(Replace(Trim(.Value), ",", ""))
                .Col = m_SetDefPrice: .Text = Format(dbltempPrice, "#,##0")
            End If
        Next lRow
    End With
End Sub


Public Function GetGroupGoodsKeyNumber(uMember As TYPE고객정보) As String
    Dim sKeyNumber  As String

    On Error GoTo ErrRtn
    
    GetGroupGoodsKeyNumber = ""
    
    
    Query = "SELECT MAX(응모번호) FROM 세트응모번호 "
    Query = Query & " WHERE LEN(응모번호) = 8 AND LEFT(응모번호,1) = '7' "
    Set Rs = New ADODB.Recordset
    Rs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If Rs.RecordCount < 1 Or IsNull(Rs.Fields(0)) Then
        sKeyNumber = "7" & Mid(대리점정보.StoreCode, 4, 3) & "0001"
        
        GoSub SUB_UPDATE
    End If
    
    If Not IsNull(Rs.Fields(0)) Then
        sKeyNumber = CStr(Val(Rs.Fields(0) & "") + 1)
    End If
    
'    If Not IsNull(rs.Fields(0)) Then
'        If Left(rs.Fields(0), 3) <> Mid(대리점정보.StoreCode, 4, 3) Then
'            sKeyNumber = Mid(대리점정보.StoreCode, 4, 3) & "10001"
'        Else
'            sKeyNumber = CStr(Val(rs.Fields(0) & "") + 1)
'        End If
'    End If
    
SUB_UPDATE:
    Query = "INSERT INTO 세트응모번호(응모번호, 세트Key, 일자, 고객코드, 고객명, 고객전화번호, 휴대폰번호, SendDate)"
    Query = Query & " VALUES('" & sKeyNumber & "', '" & m_GSGMoney.d세트Key & "', '" & Format(Date, "yyyymmdd") & "','" & uMember.고객번호 & "', "
    Query = Query & " '" & uMember.성명 & "', '" & uMember.전화번호 & "','" & uMember.휴대폰 & "', ' ') "
    ADOCon.Execute Query
    
    Exit Function

ErrRtn:
    GetGroupGoodsKeyNumber = ""
    
    Call Error_Msg("GetGroupGoodsKeyNumber", Err.Source, Err.Number, Err.Description)
End Function

'
'Public Function GetGroupGoodsKeyNumber(uMember As TYPE고객정보) As String
'    Dim sKeyNumber  As String
'    Dim Query        As String
'    Dim rs          As Recordset
'
'    On Error GoTo GetGroupGoodsKeyNumber_Error
'    GetGroupGoodsKeyNumber = ""
'
'
'    Query = "SELECT MAX(응모번호) FROM 세트응모번호 "
'    Set rs = MyDB.OpenRecordset(Query)
'
'    If rs.RecordCount < 1 Or IsNull(rs.Fields(0)) Then
'        sKeyNumber = Mid(대리점정보.StoreCode, 3, 3) & "00001"
'        GoSub SUB_UPDATE
'    End If
'
'    If Not IsNull(rs.Fields(0)) Then
'        sKeyNumber = CStr(Val(rs.Fields(0) & "") + 1)
'    End If
'
''    If Not IsNull(rs.Fields(0)) Then
''        If Left(rs.Fields(0), 3) <> Mid(대리점정보.StoreCode, 4, 3) Then
''            sKeyNumber = Mid(대리점정보.StoreCode, 4, 3) & "10001"
''        Else
''            sKeyNumber = CStr(Val(rs.Fields(0) & "") + 1)
''        End If
''    End If
'
'SUB_UPDATE:
'    Query = "INSERT INTO 세트응모번호(응모번호, 세트Key, 일자, 고객코드, 고객명, 고객전화번호, 휴대폰번호, SendDate)"
'    Query = Query & " VALUES('" & sKeyNumber & "', '" & m_GSGMoney.d세트Key & "', '" & Format(Date, "yyyymmdd") & "','" & uMember.고객번호 & "', "
'    Query = Query & " '" & uMember.성명 & "', '" & uMember.전화번호 & "','" & uMember.휴대폰 & "', ' ') "
'    ADOCon.Execute Query
'    Exit Function
'
'GetGroupGoodsKeyNumber_Error:
'    GetGroupGoodsKeyNumber = ""
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetGroupGoodsKeyNumber_Error of Module GroupSale"
'
'End Function
 
Public Function SaveGroupGoodsINFO(uMember As TYPE고객정보, uGSGoods As GroupSaleMoney_TYPE) As Boolean
    Query = " SELECT * FROM 세트상품정보 "
    Query = Query & " WHERE 세트Key = '" & uGSGoods.d세트Key & "' "
    Set Rs = New ADODB.Recordset
    Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If Rs.EOF Then
        If uGSGoods.d세트수량합계 > 0 Then
            Query = "INSERT INTO 세트상품정보(접수일자, 세트Key, 고객코드, "
            Query = Query & " 고객명 , 고객전화번호, 휴대폰번호, 정상금액, "
            Query = Query & " 세트금액, 세트할인금액, 에누리할인금액, 적용합계금액, "
            Query = Query & " 세트2, 세트3, 세트4, 세트5, "
            Query = Query & " 세트6, 세트7, 세트8, 세트9, "
            Query = Query & " 세트10, 무료세탁권수, SendDate)"
            
            Query = Query & " VALUES('" & Format(Date, "yyyymmdd") & "', '" & uGSGoods.d세트Key & "','" & uMember.고객번호 & "', "
            Query = Query & " '" & uMember.성명 & "', '" & uMember.전화번호 & "','" & uMember.휴대폰 & "', " & uGSGoods.d전체금액 & ", "
            Query = Query & " " & uGSGoods.d세트금액 & ", " & uGSGoods.d세트할인금액 & ", " & uGSGoods.d에누리할인금액 & ", " & uGSGoods.d최종수령액 & ", "
            Query = Query & " " & uGSGoods.d2세트수량 & ", " & uGSGoods.d3세트수량 & ", " & uGSGoods.d4세트수량 & ", " & uGSGoods.d5세트수량 & ", "
            Query = Query & " " & uGSGoods.d6세트수량 & ", 0,0,0,0, " & uGSGoods.d무료세탁권수량 & ",' ') "
            ADOCon.Execute Query
        End If
    Else
        Query = "UPDATE 세트상품정보 SET"
        Query = Query & " 고객코드 = '" & uMember.고객번호 & "', "
        Query = Query & " 고객명 = '" & uMember.성명 & "', "
        Query = Query & " 고객전화번호 = '" & uMember.전화번호 & "', "
        Query = Query & " 휴대폰번호 = '" & uMember.휴대폰 & "', "
        Query = Query & " 정상금액 = " & uGSGoods.d전체금액 & ", "
        Query = Query & " 세트금액 = " & uGSGoods.d세트금액 & ", "
        Query = Query & " 세트할인금액 = " & uGSGoods.d세트할인금액 & ", "
        Query = Query & " 에누리할인금액 = " & uGSGoods.d에누리할인금액 & ", "
        Query = Query & " 적용합계금액 = " & uGSGoods.d최종수령액 & ", "
        Query = Query & " 세트2 = " & uGSGoods.d2세트수량 & ", "
        Query = Query & " 세트3 = " & uGSGoods.d3세트수량 & ", "
        Query = Query & " 세트4 = " & uGSGoods.d4세트수량 & ", "
        Query = Query & " 세트5 = " & uGSGoods.d5세트수량 & ", "
        Query = Query & " 세트6 = " & uGSGoods.d6세트수량 & ", "
        Query = Query & " 세트7 = 0, "
        Query = Query & " 세트8 = 0, "
        Query = Query & " 세트9 = 0, "
        Query = Query & " 세트10 = 0, "
        Query = Query & " 무료세탁권수 = " & uGSGoods.d무료세탁권수량 & ", "
        Query = Query & " SendDate = ' ' "
        Query = Query & " WHERE 접수일자 = '" & Format(Date, "yyyymmdd") & "'  "
        Query = Query & "   AND 세트Key = '" & uGSGoods.d세트Key & "' "
        ADOCon.Execute Query
    End If
    Rs.Close
    Set Rs = Nothing
End Function

