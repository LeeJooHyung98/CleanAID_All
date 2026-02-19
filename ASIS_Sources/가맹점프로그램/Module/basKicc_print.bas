Attribute VB_Name = "basKicc_Print"
Option Explicit

Public Function PrintCustomer(전화번호출력 As String, 고객명 As String, 집전화 As String, 휴대전화 As String, 주소 As String)
    Dim ESC      As String * 1
    Dim 전화번호 As String
    Dim Print_Msg As String
    ESC = Chr(&H1B)
    Print_Msg = Print_Msg & "고 객 명 : "
    Print_Msg = Print_Msg & ESC + "!" + Chr$(32)             'Selects double-height mode
    Print_Msg = Print_Msg & 고객명 + Chr$(&HA)
    Print_Msg = Print_Msg & ESC + "!" + Chr$(0)              'Cancels double-height mode
    
    '--------------------------------------------------------------------------------------------------------
    Print_Msg = Print_Msg & "전화번호 : "
    Print_Msg = Print_Msg & ESC + "!" + Chr$(32)             'Selects double-height mode
        
    If 전화번호출력 = "Y" Then
        Print_Msg = Print_Msg & Trim(집전화) + Chr$(&HA)
    Else
        전화번호 = PhoneNo_Asterisk(Trim(집전화))
        
        Print_Msg = Print_Msg & 전화번호 + Chr$(&HA)
    End If
    
    Print_Msg = Print_Msg & ESC + "!" + Chr$(0)              'Cancels double-height mode
    
    '--------------------------------------------------------------------------------------------------------
    Print_Msg = Print_Msg & "휴대전화 : "
    Print_Msg = Print_Msg & ESC + "!" + Chr$(32)             'Selects double-height mode
        
    If 전화번호출력 = "Y" Then
        Print_Msg = Print_Msg & Trim(휴대전화) + Chr$(&HA)
    Else
        전화번호 = PhoneNo_Asterisk(Trim(휴대전화))
        
        Print_Msg = Print_Msg & 전화번호 + Chr$(&HA)
    End If
    
    Print_Msg = Print_Msg & ESC + "!" + Chr$(0)              'Cancels double-height mode
    Print_Msg = Print_Msg & "주    소 : " + 주소 + Chr$(&HA)
    PrintCustomer = Print_Msg
End Function


Public Function PrintDouble(출력내용 As String)
    Dim ESC      As String * 1
    Dim 전화번호 As String
    Dim ReturnMsg As String
    ESC = Chr(&H1B)
    
    ReturnMsg = ReturnMsg & ESC + "!" + Chr$(32)             'Selects double-height mode
    ReturnMsg = ReturnMsg & 출력내용
    ReturnMsg = ReturnMsg & ESC + "!" + Chr$(0)              'Cancels double-height mode
    PrintDouble = ReturnMsg
End Function


Public Function PrintNormal(출력내용 As String) As String
    Dim ESC      As String * 1
    Dim 전화번호 As String
    Dim ReturnMsg As String
    
    ESC = Chr(&H1B)
    ReturnMsg = ReturnMsg & ESC + "!" + Chr$(0)              'Cancels double-height mode
    ReturnMsg = ReturnMsg & 출력내용
    PrintNormal = ReturnMsg
End Function

Public Function PrintHeight(출력내용 As String) As String
    Dim ESC      As String * 1
    Dim 전화번호 As String
    Dim ReturnMsg As String
    ESC = Chr(&H1B)
    ReturnMsg = ReturnMsg & ESC + "!" + Chr$(16)             'Selects double-height mode
    ReturnMsg = ReturnMsg & 출력내용
    ReturnMsg = ReturnMsg & ESC + "!" + Chr$(0)              'Cancels double-height mode
    PrintHeight = ReturnMsg
End Function

Public Function PrintTitle(출력내용 As String) As String
    Dim ESC      As String * 1
    Dim 전화번호 As String
    Dim ReturnMsg As String
    
    ESC = Chr(&H1B)
    ReturnMsg = ReturnMsg & ESC + "!" + Chr$(0)              'Specifies font A (ESC !)
    ReturnMsg = ReturnMsg & ESC + "a" + Chr$(1)              'Specifies a centered printing position (ESC a)
    ReturnMsg = ReturnMsg & ESC + "!" + Chr$(128)             'Selects double-height mode
    ReturnMsg = ReturnMsg & ESC + "!" + Chr$(48)            'Selects double-height mode
    
    ReturnMsg = ReturnMsg & 출력내용 + Chr$(&HA)
   
    ReturnMsg = ReturnMsg & ESC + "!" + Chr$(0)              'Cancels double-height mode
    ReturnMsg = ReturnMsg & Chr$(&HA)                        'Line feeding (LF)
    ReturnMsg = ReturnMsg & ESC + "a" + Chr$(0)              'Selects the left print position (ESC a)
    PrintTitle = ReturnMsg
    
End Function

Public Function PrintTitle2(출력내용 As String) As String
    Dim ESC      As String * 1
    Dim 전화번호 As String
    Dim ReturnMsg As String
    ESC = Chr(&H1B)
    ReturnMsg = ReturnMsg & ESC + "!" + Chr$(0)              'Specifies font A (ESC !)
    ReturnMsg = ReturnMsg & ESC + "a" + Chr$(1)              'Specifies a centered printing position (ESC a)
    ReturnMsg = ReturnMsg & ESC + "!" + Chr$(16)             'Selects double-height mode
    ReturnMsg = ReturnMsg & 출력내용
    ReturnMsg = ReturnMsg & ESC + "!" + Chr$(0)              'Cancels double-height mode
    ReturnMsg = ReturnMsg & Chr$(&HA)                        'Line feeding (LF)
    ReturnMsg = ReturnMsg & ESC + "a" + Chr$(0)              'Selects the left print position (ESC a)
    PrintTitle2 = ReturnMsg
End Function

'Chr(&H1D) & Chr(&H21) & Chr(&H1) & "    [신용 구매]     " & Chr(&H1D) & Chr(&H21) & Chr(&H0)

Public Function PrintString(출력내용 As String, Size As Integer, Optional NewLine As Boolean = False) As String
    Dim ESC      As String * 1
    Dim 전화번호 As String
    Dim ReturnMsg As String
    ESC = Chr(&H1B)
    Select Case Size
    Case 1 ' 일반
        ReturnMsg = ReturnMsg & PrintNormal(출력내용)
    Case 4 ' 가로세로 더블
        ReturnMsg = ReturnMsg & PrintDouble(출력내용)
    Case 6 ' 세로 더블
        ReturnMsg = ReturnMsg & PrintHeight(출력내용)
    End Select
    If NewLine Then
        ReturnMsg = ReturnMsg & PrintLineFeed
    End If
    PrintString = ReturnMsg
End Function

Public Function PrintLineFeed(Optional LineCount As Integer = 1) As String
    Dim LoopI As Integer
    Dim ReturnMsg As String
    For LoopI = 1 To LineCount
        ReturnMsg = ReturnMsg & Chr$(&HA)                        'Line feeding (LF)
    Next LoopI
    PrintLineFeed = ReturnMsg
End Function

Public Function PrintCut() As String
    PrintCut = Chr(&H1B) & "i" & Chr(&HD) & Chr(&HA) & "<C>" 'Feeds paper & cut &
    
End Function

Public Function SetMessage(Send_Type As Approve_Type, Money As String, Optional Receipt_Method As String = "", Optional RequestAcceptNumber As String = "", Optional AcceptDate As String = "", Optional Pos_Approve As String) As String
    Dim Request_Type As String
    Dim WCC As String
    Dim CardNumber As String
    Dim ReceiptType As String
    Dim InputDate As String
    Dim AcceptNumber As String
    Dim InputMoney As String
    Dim ServiceMoney As String
    Dim VAT As String
    Dim Pos_Number As String
    
    ' 현금영수증 일 경우 00 : 개인, 01 : 사업자
    ReceiptType = Space(2)
    
    Select Case Send_Type
    Case Credit_Approve
        Request_Type = "D1"
    Case Credit_Cancel_Today
        Request_Type = "D2"
    Case Credit_Cancel_Prev_Day
        Request_Type = "D4"
    Case Cash_Approve
        Request_Type = "B1"
    Case Cash_Cancel_Today
        Request_Type = "B2"
    Case Cash_Cancel_Prev_Day
        Request_Type = "B3"
    End Select
    
    If Receipt_Method <> "" Then
        ReceiptType = Receipt_Method
    End If
    
    WCC = Space(1)
    
    CardNumber = Space(40)
    If AcceptDate <> "" Then
        InputDate = AcceptDate
        'InputDate = Space(6)
    Else
        InputDate = Space(6)
    End If
    
    AcceptNumber = RPad(12, RequestAcceptNumber)
    InputMoney = RPad(8, Money)
    ServiceMoney = RPad(8, "0")
    VAT = RPad(8, "0")
    If Pos_Approve <> "" Then
        Pos_Number = RPad(20, Pos_Approve)
    Else
        Pos_Number = Space(20)
    End If
    
    SetMessage = Request_Type & WCC & CardNumber & ReceiptType & InputDate & AcceptNumber & InputMoney & ServiceMoney & VAT & Pos_Number
End Function


Public Function lPad(Num, Str)
    Do While Len(Str) < Num
        Str = " " & Str
    Loop
    lPad = Str
End Function

Public Function RPad(Num, Str)
    Do While Len(Str) < Num
        Str = Str & " "
    Loop
    RPad = Str
End Function


