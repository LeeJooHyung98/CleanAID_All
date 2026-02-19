Attribute VB_Name = "pds_Module"
Option Explicit

Public Const MASTER_OFFICE_CODE As String = "1000"

Public Sub SubBottonEnable(cbo As Object, sBtnIndex As String)
    Dim i As Integer
    
    For i = 0 To Len(sBtnIndex) - 1
        cbo(i).Enabled = Val(Mid(sBtnIndex, i + 1, 1))
    Next i
    
End Sub



'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : OrderComboAdd
' 작  성  자  : pds2004
' 작  성  일  : 2010.10.20
' 파 라 미 터 : Control  - Combo Box Control Object
' 비      고  : Combo Box에 외주 업체 내역을 Add한다.
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Sub OrderComboAdd(Control As Object)
    Dim Rs As ADODB.Recordset
    Dim sValue() As String
    
    Dim Err_Num As Long
    Dim Err_Dec As String
    
    Control.Clear
    
    ReDim sValue(1)
    
    sValue(0) = "0"
    sValue(1) = Store.Code
    
    Set Rs = New ADODB.Recordset
    Set Rs = ExecPro("SP_M_07000_01", sValue(), Err_Num, Err_Dec)

    Control.AddItem ""

    Do While Not Rs.EOF
        Control.AddItem "[" & Rs!외주코드 & "] " & Rs!외주명
        
        Rs.MoveNext
    Loop
    Rs.Close
    
End Sub


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

Public Sub DateCheckAdd(spdView As fpSpread, sSDate As String, sEDate As String)
    Dim nRow        As Long
    Dim nStartRow   As Long
    Dim varDate     As Variant
    '-----------------------------------------------------------------------------------------------------
    '첫번째 라인에 일자를 확인한다.
    Call spdView.GetText(1, 1, varDate)
    If varDate = "" Then varDate = sEDate
    Call SubDateADD(spdView, 0, sSDate, CStr(varDate))
    
    nStartRow = 1
SUB_RTN:
    ' 최종 일자가 없는 것을 삽입한다.
    For nRow = nStartRow To spdView.MaxRows - 1
        ' 다음 날자를 확인한다.
        Call spdView.GetText(1, nRow + 1, varDate)
        
        spdView.Row = nRow
        spdView.Col = 1
        If Format(DateAdd("d", 1, spdView.Text), "yyyy-MM-dd") < CStr(varDate) Then
    
            '첫번째 라인에 일자를 확인한다.
            Call SubDateADD(spdView, nRow, spdView.Text, CStr(varDate))
            nStartRow = nRow
            GoSub SUB_RTN
        End If
    Next nRow

    ' 마지막 일자에서 조회종료일까지
    Call spdView.GetText(1, spdView.MaxRows, varDate)
    If IsDate(varDate) = False Then varDate = sEDate
    Call SubDateADD(spdView, -1, CStr(varDate), sEDate)
    '-----------------------------------------------------------------------------------------------------

    
End Sub

Private Sub SubDateADD(spdView As fpSpread, nRow As Long, sSDate As String, sEDate As String)
    ' 최종 일자가 없는 것을 삽입한다.
    Dim nCnt    As Integer
    Dim nRow2   As Long
    Dim nCol    As Long
    
    ' 첫번째 라인에 일자가 적을경우
    If nRow = 0 Then
        spdView.Row = 1
        spdView.Col = 1
        
        If spdView.Text > sSDate Then
            For nCnt = 1 To DateDiff("d", sSDate, sEDate)
                spdView.MaxRows = spdView.MaxRows + 1
                spdView.Row = nCnt
                spdView.Action = ActionInsertRow
                
                '---------------------------------------------------------------------------------------
                ' Spread의 버그가 있음
                ' 왜 그러는지는 모리지만 ActionInsertRow 이후 ExportToExcel함수에서 저장할경우
                ' 다음줄이ActionInsertRow 이후 자료를 추가한 col까지만 저장되는 문제가 있음
                For nCol = 1 To spdView.MaxCols
                    spdView.Col = nCol: spdView.Text = "test"
                Next nCol
                spdView.Col = -1:   spdView.Action = ActionClearText
                '---------------------------------------------------------------------------------------
                
                spdView.Col = 1
                spdView.Text = Format(DateAdd("d", nCnt - 1, sSDate), "yyyy-mm-dd")
                
                spdView.Col = 2
                spdView.Text = ExecWeekDay(DateAdd("d", nCnt - 1, sSDate))
                
                spdView.Col = -1
                spdView.BackColor = vbWhite
                
            Next nCnt
        End If
        
    ' 마지막 라인일 경우
    ElseIf nRow < 0 Then
        spdView.Row = nRow
        spdView.Col = 1
        
        For nCnt = 1 To DateDiff("d", sSDate, sEDate)
            spdView.MaxRows = spdView.MaxRows + 1
            
            spdView.Row = spdView.MaxRows
            spdView.Col = 1
            spdView.Text = Format(DateAdd("d", nCnt, sSDate), "yyyy-mm-dd")
            
            spdView.Col = 2
            spdView.Text = ExecWeekDay(DateAdd("d", nCnt, sSDate))
            
            spdView.Col = -1
            spdView.BackColor = vbWhite
            
        Next nCnt
    
    ' 2
    Else
        spdView.Row = nRow
        spdView.Col = 1
        
        For nCnt = 1 To DateDiff("d", sSDate, sEDate) - 1
            
            spdView.MaxRows = spdView.MaxRows + 1
            
            spdView.Row = nRow + nCnt
            spdView.Action = ActionInsertRow
            
            '---------------------------------------------------------------------------------------
            ' Spread의 버그가 있음
            ' 왜 그러는지는 모리지만 ActionInsertRow 이후 ExportToExcel함수에서 저장할경우
            ' 다음줄이ActionInsertRow 이후 자료를 추가한 col까지만 저장되는 문제가 있음
            For nCol = 1 To spdView.MaxCols
                spdView.Col = nCol: spdView.Text = "test"
            Next nCol
            spdView.Col = -1:   spdView.Action = ActionClearText
            '---------------------------------------------------------------------------------------
            
            spdView.Col = 1
            spdView.Text = Format(DateAdd("d", nCnt, sSDate), "yyyy-mm-dd")
            
            spdView.Col = 2
            spdView.Text = ExecWeekDay(DateAdd("d", nCnt, sSDate))
            
            spdView.Col = -1
            spdView.BackColor = vbWhite
        
        Next nCnt
    
    End If

End Sub


'--------------------------------------------------------------------------------------------------------------
' Procedure : GetCheckCount
' DateTime  : 2007-06-28 12:22
' Author    : pds2004
' Purpose   : 선택된 내용의 가맹점 코드 리스트를  리턴한다.
'--------------------------------------------------------------------------------------------------------------
Public Function GetSelectMasterCodeList(ByRef MySpread As Object, ByVal LCol As Long) As String
    Dim nRow        As Long
    Dim vText       As Variant
    Dim SelCodeList As String
    
    SelCodeList = "":       GetSelectMasterCodeList = ""
    
    With MySpread
        If .MaxRows <= 0 Then Exit Function
        
        For nRow = 1 To MySpread.MaxRows
            Call .GetText(LCol, nRow, vText)
            If CStr(vText) = "1" Then
                Call .GetText(LCol + 1, nRow, vText)
                
                ' 지사 이외의 내용을 제외한다.(전체 등등)
                If IsNumeric(vText) = True Then
                    SelCodeList = SelCodeList & CStr(vText) + ","
                End If
            End If
        Next nRow
        
        ' 마지막 ,를 삭제한다.
        If Len(SelCodeList) > 4 Then SelCodeList = Mid(SelCodeList, 1, Len(SelCodeList) - 1)
    End With
    
    
    GetSelectMasterCodeList = SelCodeList

End Function



'====================================================================================================
' Procedure : GetCheckSMSSendTel
' DateTime  : 15-06-15
' Author    : pds2004
' Last Edit :
' Return Val:
'----------------------------------------------------------------------------------------------------
' 설     명 : 문자 메시지 발송 번호 확인
'      전기통신사업법 제84조에 의하여 문자 발신 번호는 반드시 입력 하여야 합니다."
'      주요내용
'      - 발신번호 없이 문자 전송 불가
'      - 발신번호는 수신자가 실제 발신(통화)이 가능한 번호만 허용
'      - 일반번호의 경우 지역번호(02,031등)를 앞자리에 포함한 번호만 허용
'      - 대표번호는 8자리만 입력 허용되며,내선번호 포함불가
'       - 030 번호의 경우 12자리까지 허용
'====================================================================================================
Public Function GetCheckSMSSendTel(ByVal sNumber As String, ByRef sTel() As String, Optional s지역번호확인 As Boolean = True) As Boolean
    
    Dim strTel As String
    Dim Tel_Check  As String
    
    On Error GoTo GetCheckSMSSendTel_Error
    
    GetCheckSMSSendTel = False
    strTel = Trim(Replace(sNumber, "-", ""))
    strTel = Replace(strTel, ")", "")

    ' 기본 사항 확인
    If IsNumeric(strTel) = False Then Exit Function
    If s지역번호확인 = True And Len(strTel) <= 7 Then Exit Function
    If Len(strTel) < 7 Then Exit Function
    
    If Len(strTel) = 7 Then
        Tel_Check = Format(strTel, "-000-0000")
    ElseIf Len(strTel) = 8 Then
        Tel_Check = Format(strTel, "0000-0000")
    ElseIf Len(strTel) = 9 Then
        Tel_Check = Format(strTel, "00-000-0000")
    ElseIf Len(strTel) = 10 Then
         If Left(strTel, 2) = "02" Then
             Tel_Check = Format(strTel, "00-0000-0000")
         Else
             Tel_Check = Format(strTel, "000-000-0000")
         End If
    ElseIf Len(strTel) = 11 Then
        Tel_Check = Format(strTel, "000-0000-0000")
    ElseIf Len(strTel) = 12 Then
         If Left(strTel, 3) = "030" Then
             Tel_Check = Format(strTel, "000-0000-0000")
         End If
        
    ElseIf Len(strTel) >= 13 Then
        
        Tel_Check = ""
    End If
    
    sTel = Split(Tel_Check, "-")
    GetCheckSMSSendTel = True
    
    Exit Function

GetCheckSMSSendTel_Error:
    GetCheckSMSSendTel = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetCheckSMSSendTel of Form pds_Module"

End Function


 

Sub SearchString(Ky As Integer)
  
  Dim s As String, l As Long
  Dim cbo As ComboBox

  If TypeOf Screen.ActiveControl Is ComboBox Then
     Set cbo = Screen.ActiveControl
     If cbo.Style < 2 Then 'ComboBox의 Style이 2이하에서만 입력 가능
        s = Left(cbo.Text, cbo.SelStart) & Chr(Ky) '검색 문자열을 만든다.
        If Left(s, 1) <> "[" Then s = "[" & s
        Debug.Print s
        
        If IsNumeric(Mid(s, 2, 4)) = True Then
      
            l = SendMessage(cbo.hwnd, CB_FINDSTRING, -1, ByVal s) '문자열을 검색한다
              
              If l <> CB_ERR Then '콤보박스에서 찾은 것이 있으면
                 With cbo ' ListIndex를 설정하고 선택한다
                    .ListIndex = l
                    .Text = .List(l)
                    .SelStart = Len(s)
                    .SelLength = Len(.Text)
                 End With
                 Ky = 0
              End If
        End If
     End If
  End If
End Sub

Sub SearchString_한글(Ky As Integer)
  
  Dim s As String, l As Long, i As Integer
  Dim cbo As ComboBox

  If TypeOf Screen.ActiveControl Is ComboBox Then
     Set cbo = Screen.ActiveControl
     If cbo.Style < 2 Then 'ComboBox의 Style이 2이하에서만 입력 가능
        s = cbo.Text   '검색 문자열을 만든다.
        
        Debug.Print s
        
            
        For i = 0 To cbo.ListCount
            Debug.Print cbo.List(i) & " : " & Mid(s, 1, Len(s)) & " : " & InStr(cbo.List(i), Mid(s, 1, Len(s)))
            
            If InStr(cbo.List(i), Mid(s, 1, Len(s))) > 0 Then
                cbo.ListIndex = i
                
'                cbo.SelStart = InStr(cbo.List(i), Mid(s, 1, Len(s))) + 1
'                cbo.SelLength = Len(cbo.Text)
                
                Exit Sub
        
            End If
        Next i
        
        If i > cbo.ListCount Then
            cbo.SelStart = 0
            cbo.SelLength = 10000
        End If
            
     End If
  End If
End Sub


 

Sub SearchString_ORG(Ky As Integer)
  
  Dim s As String, l As Long
  Dim cbo As ComboBox

  If TypeOf Screen.ActiveControl Is ComboBox Then
     Set cbo = Screen.ActiveControl
     If cbo.Style < 2 Then 'ComboBox의 Style이 2이하에서만 입력 가능
        s = Left(cbo.Text, cbo.SelStart) & Chr(Ky) '검색 문자열을 만든다.
            If Left(s, 1) <> "[" Then s = "[" & s
          l = SendMessage(cbo.hwnd, CB_FINDSTRING, -1, ByVal s) '문자열을 검색한다
            If l <> CB_ERR Then '콤보박스에서 찾은 것이 있으면
               With cbo ' ListIndex를 설정하고 선택한다
                  .ListIndex = l
                  .Text = .List(l)
                  .SelStart = Len(s)
                  .SelLength = Len(.Text)
               End With
               Ky = 0
            End If
         End If
  End If
End Sub

