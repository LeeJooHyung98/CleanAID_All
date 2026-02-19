Attribute VB_Name = "mod_Function"
Option Explicit

' 고객 정보 조회 함수
Public Function GetCustomer(Search As String)


End Function

' 상품 대분류 정보 조회 함수
Public Function GetGoodsTitle()
    '----------------------------------------------------------
    ' TB_의류분류
    '----------------------------------------------------------
    Query = "SELECT    의류분류코드"
    Query = Query & ", 의류분류명"
    Query = Query & ", 순서"
    Query = Query & " FROM TB_의류분류"
    Query = Query & " ORDER BY 순서"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    Dim LoopI As Integer
    LoopI = 0
    Do Until ADORs.EOF
           
        If UCase(Left(ADORs!의류분류코드, 1)) = "W" And 가맹점정보.지사코드 = "1024" Then

        End If
           
        frmAccept.Goods(LoopI).Caption = ADORs!의류분류명
        frmAccept.Goods(LoopI).Tag = ADORs!의류분류코드
        LoopI = LoopI + 1
        
        ADORs.MoveNext
    Loop
    ADORs.Close
    Set ADORs = Nothing

End Function

' 상품 중분류 정보 조회 함수
Public Function GetGoodsSub(Search As String) As ADODB.RecordSet
    '----------------------------------------------------------
    ' TB_의류분류
    '----------------------------------------------------------
    Query = "SELECT 의류코드, 의류명, 금액 FROM TB_의류 WHERE 적용일자 = (SELECT MAX(적용일자) FROM TB_의류) AND SUBSTRING(의류코드,1,2) = '" & Search & "' ORDER BY 순서,의류코드"

    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    Set GetGoodsSub = ADORs.Clone
    ADORs.Close
    Set ADORs = Nothing
End Function
