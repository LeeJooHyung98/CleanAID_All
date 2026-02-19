Attribute VB_Name = "Printer1"
Option Explicit

' Ãâ·Â¿¡ ÇÊ¿äÇÑ º¯¼ö
Public Company_Name             As String
Public strMsg                   As String
Public ProssCount               As Integer      ' Ãâ·Â¹°ÀÇ ÃÑ °¹¼ö
Public TempPrint1               As Recordset    ' Ãâ·ÂÇÒ DB
Public TempPrint2               As Recordset    ' Ãâ·ÂÇÒ DB
Public TempBool                 As Boolean      ' ´ÙÀ½ ¶óÀÎ Ãâ·Â
Public TempMoney1                As Double   ' ´©°è ±Ý¾×À» °è»ê
Public TempMoney2               As Double   ' ´©°è ±Ý¾×À» °è»ê
Public TempMoney3               As Double   ' ´©°è ±Ý¾×À» °è»ê
Public ReView                  As Boolean   ' ¹Ì¸® º¸±â ¼³Á¤ True=>¹Ù·Î ÀÎ¼â
Public FHandle                 As Integer   ' ÀÎ¼âÇÒ ÆÄÀÏÀÇ ÇÚµé
Public TextData(20)            As String    ' ÀÎ¼âÇÒ ³»¿ëÀ» ÀÓ½Ã ÀúÀåÇÑ´Ù.
Public hhh(60)                 As String    ' ¾ç½Ä
Public Title(20)                As String
Public StartDate               As String    ' ÀÎ¼â ½ÃÀÛÀÏ
Public starttime               As String    ' ÀÎ¼â ½ÃÀÛ ½Ã°£
Public strFileName             As String    ' ÀÎ¼âÇÒ È­ÀÏ ÀÌ¸§
Public PageCnt                  As Integer   ' ÀÎ¼âÇÒ ÆäÀÌÁö ¹× ¶óÀÎ¼ö
Public LineCnt                  As Integer
Public PRINT_LINE_COUNT         As Integer


Type FPTop
    Tel As String       ' ÀüÈ­¹øÈ£
    Name As String      ' ÀÌ¸§
    Date As Date        ' ¿À´Ã ³¯Â¥.
    Date2 As String     ' Ãâ°í ¿¹Á¤ÀÏ
End Type

Type FPBottom
    Sum As Variant
    Account0 As Variant
    Account1 As Variant
    Account2 As Variant
    Tel As String
    Name As Variant
    Addr As Variant
End Type

'''''''''''''''''''''''''''''''''''''
''''' pds2004
'''''''''''''''''''''''''''''''''''''
' ±âº» ¿©¹é
Public Prt_Top As Integer
Public Prt_Left As Integer
Public Prt_Height As Integer
Public Top_Margin As Integer    ' ÇØ´ç ´ë»óÀÇ Å¾ ±âº» À§Ä¡ º¯°æ
Public Left_Margin As Integer   ' ÇØ´ç ´ë»óÀÇ ÁÂÃø ±âº» À§Ä¡ º¯°æ
Public Text_Height As Integer   ' TextÀÇ ³ôÀÌ


Public Type PrintPoint
    x As Integer
    y As Integer
End Type

' º¸°üÁõ Ãâ·Â »ó´Ü
Type FPrint_Top
    PrtNo   As String       ' ÀüÇ¥¹øÈ£
    Tel     As String       ' ÀüÈ­¹øÈ£
    Name    As String       ' ÀÌ¸§
    Addr    As String       ' °í°´ ÁÖ¼Ò
    Date    As Date         ' Á¢¼öÀÏ
    Date2   As String       ' Ãâ°í ¿¹Á¤ÀÏ
    Code    As String       ' °í°´¹øÈ£
    HpTel   As String       ' ÈÞ´ëÆù ¹øÈ£
End Type

' º¸°üÁõ Ãâ·Â ÇÏ´Ü
Type FPrint_Bottom
    Sum         As String   ' Á¡¼ö
    Counter     As String  '
    Account0    As String   ' ±Ý¾×
    Account1    As String   ' ¼ö·É¾×
    Account2    As String   ' ÀÜ¾×
    DName       As String   ' ´ë¸®Á¡¸í
    DTel        As String   ' ´ë¸®Á¡ ÀüÈ­¹øÈ£
    OldDayMisu  As String   ' ÀüÀÏ ¹Ì¼ö
    MiSuTotal   As String   ' ¹Ì¼ö ÇÕ°è
    MilUser     As String   ' »ç¿ë ¸¶ÀÏ¸®Áö
    MilMoney    As String   ' ¸¶ÀÏ¸®Áö ÀÜ¾×
    MilAddMoney As String   ' ´©Àû ¸¶ÀÏ¸®Áö
    SuGumMonye  As String   ' ¼ö±Ý¾×
    CouponCnt   As String
    CouponNum   As String
    CouponMoney As String
End Type

' ÀÏÀÏ¸ÅÃâ ÇöÇ× Ãâ·Â
Type FPDayTop
    Compnay As String
    Title As String
    sDay As String
    TagNum As String
    Tel As String
    Name As String
    PName As String
    PAccount As String
    PColor As String
    PTemp As String
    PTemp2 As String
    Flag As String
End Type

Public FPrtTop As FPrint_Top
Public FPrtBottom As FPrint_Bottom
Public FPDayPrint As FPDayTop
Public PrtPoint As PrintPoint   ' ±âº» ÁÂÇ¥
Public PrtPoint2 As PrintPoint  ' ¶óÀÎ°£°Ý
Public PrtPoint3 As PrintPoint  ' ¼Õ´Ô¿ë
Public PrtPoint4 As PrintPoint  ' ¿©¹é
'''''''''''''''''''''''''''''''''''''
Public FPArray(1 To 1000, 1 To 6) As Variant
Public FPTop As FPTop
Public FPBottom As FPBottom


Public Page_Count As Integer
' Ãâ·ÂÇÒ Ç×¸ñÀÇ ÃÑ °¹¼ö
Private iRowCount As Integer
'Public Printer_Top              As Double ' À§ ¿©¹é
'Public Printer_Left             As Double ' ÁÂ ¿©¹é
'Public Printer_Height           As Double ' À§ °ø¹é

Public Const Printer_Top = 30       ' À§ ¿©¹é
Public Const Printer_Left = 30      ' ÁÂ ¿©¹é
Public Const Printer_Height = 30    ' À§ °ø¹é



Dim sSEQ As String
'

' ¹®¼­ ÆíÁý±â ½ÇÇà
Public Sub EDIT_Text(strTitle As String)
    Call Shell("notepad.exe " & strTitle, vbNormalFocus)
End Sub

Public Function PrinterCheck() As Boolean
    On Error GoTo Err_Rtn
    
    Dim printer_name As String
  
    Dim x As Printer
    
    For Each x In Printers
        printer_name = x.DeviceName
    Next

    If Printer.DeviceName = "" Then
        MsgBox "ÇÁ¸°ÅÍ¸¦ ¼³Ä¡ÇØ ÁÖ½Ê½Ã¿ä!", vbInformation, "È®ÀÎ"
        PrinterCheck = False
        Exit Function
    End If
    
    PrinterCheck = True
    
    Exit Function
  
Err_Rtn:
    PrinterCheck = False
    
    Exit Function
End Function

'If InStr(1, ppp$, "") > 0 Then
'È­ÀÏÀ» ÇÁ¸°ÅÍ·Î  Ãâ·Â ÇÑ´Ù.
'*****************************************************************
Public Sub FileToPrint(strFileName As String, Ãâ·Â¹æÇâ As Integer, bView As Boolean)
    Dim ppp As String

    On Error GoTo Error_Handle
    
    If bView Then
        ' ¹Ì¸® º¸±âÀÌ¸é
        EDIT_Text (strFileName)
    Else
        ' ÀÎ¼â
        FHandle = FreeFile
        Printer.FontName = "±¼¸²Ã¼"
''           Printer.ShowPrinter
        Printer.Orientation = Ãâ·Â¹æÇâ
        Open strFileName For Input As #FHandle
        Do
            Line Input #FHandle, ppp$
            ' »õ·Î¿î ÆäÀÌÁö È®ÀÎ
            If Left$(ppp$, 1) = "" Then
                Printer.NewPage

            Else
                ' Å¸ÀÌÆ²ÀÎÁö È®ÀÎÇÑ´Ù.
                If InStr(1, ppp$, "") > 0 Then
                    Printer.FontSize = 17
                    Mid(ppp$, InStr(1, ppp$, ""), 1) = Space(1)
                Else
                ' º¸Åë ÀÚ·áÀÏ °æ¿ì
                    Printer.FontSize = 10
                End If
                Printer.Print ppp$
            End If
        Loop Until EOF(FHandle)
        Printer.EndDoc
        Close #FHandle
    End If
    Exit Sub

    
'Error Ã³¸®ºÎ
Error_Handle:
    Close #FHandle
    strMsg = "Error Number : " & CStr(Err.Number) & Chr(10) & Chr(13) & _
        "Error Description : " & Err.Description
    MsgBox strMsg, 16, "Error Message!"
    Printer.KillDoc
    Resume Next
End Sub
'
'Function Page_Title()
'    Dim Query As String
'    Const gg As Integer = 2500
'    Screen.MousePointer = 13
'
'    'array ÃÊ±âÈ­
'    For iCnt1 = 1 To 100
'        For iCnt2 = 1 To 5
'            FPArray(iCnt1, iCnt2) = ""
'        Next iCnt2
'    Next iCnt1
'
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    '»ó´Ü
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    FPrtTop.Name = frmÁ¢¼ö.txtName.Text
'    FPrtTop.Date = Str(Date)
'    FPrtTop.Tel = frmÁ¢¼ö.txtTEL(0).Text & "-" & frmÁ¢¼ö.txtTEL(1).Text
'    'Ãâ°í¿¹Á¤ÀÏ
'    If frmÁ¢¼ö.Option1 Then
'        FPrtTop.Date2 = CStr(Month(Date + 3)) & "  " & CStr(Day(Date + 3))
'    ElseIf frmÁ¢¼ö.Option2 Then
'        FPrtTop.Date2 = CStr(Month(Date + 4)) & "  " & CStr(Day(Date + 4))
'    Else
'        FPrtTop.Date2 = CStr(Month(Date + 5)) & "  " & CStr(Day(Date + 5))
'    End If
'
'    Query = "SELECT * "
'    Query = Query & "FROM ´ë¸®Á¡Á¤º¸ "
'
'    Set rsAgent = MyDB.OpenRecordset(Query)
'
'    If rsAgent.RecordCount < 1 Then
'        Debug.Print ("Ãâ·Â ´ë¸®Á¡ Á¤º¸ ´ë¸®Á¡¸í, ÀüÈ­ ¹øÈ£ ºÎÁ·")
'    Else
'        FPrtTop.DName = rsAgent!´ë¸®Á¡¸í
'        FPrtTop.DTel = rsAgent!ÀüÈ­1 & "-" & rsAgent!ÀüÈ­2
'    End If
'
'    On Error GoTo printError
'
'    'Call Title_BOX
'    Printer.FontName = "¹ÙÅÁ"
'    Printer.Font.Bold = True
'    Printer.Font.Size = 10
'    Printer.CurrentY = Printer_Top - 10
'    Printer.CurrentX = 600 + Printer_Left
'    Printer.Print "Page :  " & Printer.Page & "/" & Page_Count
'
'    ' °í°´ ÀüÈ­¹øÈ£
'    Printer.FontName = "¹ÙÅÁ"
'    Printer.Font.Bold = True
'    Printer.Font.Size = 10
'    Printer.CurrentY = Printer_Top + Printer_Height * 19.3
'    Printer.CurrentX = 10500 + Printer_Left
'    Printer.Print FPrtTop.Tel
'
'    ' °í°´ ¼º¸í
'    Printer.FontName = "¹ÙÅÁ"
'    Printer.Font.Bold = True
'    Printer.Font.Size = 10
'    Printer.CurrentY = Printer_Top + Printer_Height * 21.205
'    Printer.CurrentX = 7700 + Printer_Left
'    Printer.Print FPrtTop.Name
'
'    ' ´ë¸®Á¡¸í
'    Printer.FontName = "¹ÙÅÁ"
'    Printer.Font.Bold = True
'    Printer.Font.Size = 10
'    Printer.CurrentY = Printer_Top + Printer_Height * 21.205
'    Printer.CurrentX = 10500 + Printer_Left
'    Printer.Print FPrtTop.DName
'
'    ' Á¢¼öÀÏ
'    Printer.FontName = "¹ÙÅÁ"
'    Printer.Font.Bold = True
'    Printer.Font.Size = 10
'    Printer.CurrentY = Printer_Top + Printer_Height * 21.205
'    Printer.CurrentX = 10500 + Printer_Left
'    Printer.Print FPrtTop.Date
'
'    ' ´ë¸®Á¡ ÀüÈ­¹øÈ£
'    Printer.FontName = "¹ÙÅÁ"
'    Printer.Font.Bold = True
'    Printer.Font.Size = 10
'    Printer.CurrentY = Printer_Top + Printer_Height * 21.205
'    Printer.CurrentX = 10500 + Printer_Left
'    Printer.Print FPrtTop.DTel
'
'    ' ÀÎµµ ¿¬µµ
'    Printer.FontName = "¹ÙÅÁ"
'    Printer.Font.Bold = True
'    Printer.Font.Size = 10
'    Printer.CurrentY = Prinr_Top + Printer_Height * 21.205
'    Printer.CurrentX = 10500 + Printer_Left
'    Printer.Print FPrtTop.Date2
'
'    Screen.MousePointer = 0
'    Exit Function
'
'printError:
'    MsgBox " ÇÁ¸°ÅÍ¸¦ È®ÀÎÇØ ÁÖ½Ê½Ã¿ä ! " & VBA.Err.Number, vbCritical, "Ãâ·Â¿À·ù¹ß»ý"
'End Function
'Function Page_Printer(Start_Line As Integer, End_Line As Integer, ST_I As Integer)
'    Dim i%
'    Dim ll%
'    Dim stk As String
'    Dim PUMst As String
'    DY = Printer_Height
'    YS = 1900 + Printer_Top
'    ll = 0
'    SUB_TOT = 0
'    SUB_S_TOT = 0
'
'    On Error GoTo printError
'
'    Printer.FontName = "¹ÙÅÁ"
'    Printer.Font.Size = 10   ' 8
'    Printer.Font.Bold = True
'    For i = Start_Line To End_Line
'        Screen.MousePointer = 11
'        ll = ll + 1
'
'        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        ' ¼Õ´Ô¿ë
'        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        'ÅÃ¹øÈ£
'        Printer.CurrentX = 500 + Printer_Left
'        Printer.CurrentY = YS + DY * ll
'        frmÁ¢¼ö.sprGrid.Col = 2
'        Printer.Print frmÁ¢¼ö.sprGrid.Value
'
'        'Ç°¸í
'        Printer.CurrentX = 700 + Printer_Left
'        Printer.CurrentY = YS + DY * ll
'        frmÁ¢¼ö.sprGrid.Col = 1
'        Printer.Print frmÁ¢¼ö.sprGrid.Value
'
'        '»ö»ó
'        Printer.CurrentX = 900 + Printer_Left
'        Printer.CurrentY = YS + DY * ll
'        frmÁ¢¼ö.sprGrid.Col = 3
'        Printer.Print frmÁ¢¼ö.sprGrid.Value
'
'        '±Ý¾×
'        Printer.CurrentX = 1100 + Printer_Left
'        Printer.CurrentY = YS + DY * ll
'        frmÁ¢¼ö.sprGrid.Col = 5
'        Printer.Print frmÁ¢¼ö.sprGrid.Value
'
'        '³»¿ë
'        Printer.CurrentX = 1300 + Printer_Left
'        Printer.CurrentY = YS + DY * ll
'        frmÁ¢¼ö.sprGrid.Col = 4
'        Printer.Print frmÁ¢¼ö.sprGrid.Value
'        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        ' º¸°ü¿ë
'        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        'ÅÃ¹øÈ£
'        Printer.CurrentX = 500 + Printer_Left
'        Printer.CurrentY = YS + DY * ll
'        frmÁ¢¼ö.sprGrid.Col = 2
'        Printer.Print frmÁ¢¼ö.sprGrid.Value
'
'        'Ç°¸í
'        Printer.CurrentX = 700 + Printer_Left
'        Printer.CurrentY = YS + DY * ll
'        frmÁ¢¼ö.sprGrid.Col = 1
'        Printer.Print frmÁ¢¼ö.sprGrid.Value
'
'        '»ö»ó
'        Printer.CurrentX = 900 + Printer_Left
'        Printer.CurrentY = YS + DY * ll
'        frmÁ¢¼ö.sprGrid.Col = 3
'        Printer.Print frmÁ¢¼ö.sprGrid.Value
'
'        '±Ý¾×
'        Printer.CurrentX = 1100 + Printer_Left
'        Printer.CurrentY = YS + DY * ll
'        frmÁ¢¼ö.sprGrid.Col = 5
'        Printer.Print frmÁ¢¼ö.sprGrid.Value
'
'        '³»¿ë
'        Printer.CurrentX = 1300 + Printer_Left
'        Printer.CurrentY = YS + DY * ll
'        frmÁ¢¼ö.sprGrid.Col = 4
'        Printer.Print frmÁ¢¼ö.sprGrid.Value
'
'    Next i
'Screen.MousePointer = 0
'Exit Function
'
'printError:
'    MsgBox " ÇÁ¸°ÅÍ¸¦ È®ÀÎÇØ ÁÖ½Ê½Ã¿ä ! " & VBA.Err.Number, vbCritical, "Ãâ·Â¿À·ù¹ß»ý"
'End Function
'Function Page_Bottom(bEndflag As Boolean)
'    Dim i%
'    Dim ll%
'    Dim stk As String
'    Dim PUMst As String
'
'    DY = Printer_Height
'    YS = 1900 + Printer_Top
'    ll = 0
'    SUB_TOT = 0
'    SUB_S_TOT = 0
'
'    ' ÇÏ´Ü Ãâ·Â°ª ÃÊ±âÈ­
'    FPrtBottom.Counter = iRowCount
'    FPrtBottom.Addr = frmÁ¢¼ö.txtAddress
'    FPrtBottom.Name = frmMain.StatusBar1.Panels(2).Text
'    FPrtBottom.Tel = frmMain.StatusBar1.Panels(5).Text
'    FPrtBottom.Account0 = frm°áÁ¦.Label6
'    FPrtBottom.Account1 = frm°áÁ¦.txtMoney.Text
'    FPrtBottom.Account2 = frm°áÁ¦.Label3
'
'
'    On Error GoTo printError
'    ' Ãâ·Â ÇüÅÂ Á¤ÀÇ
'    Printer.FontName = "¹ÙÅÁ"
'    Printer.Font.Size = 10   ' 8
'    Printer.Font.Bold = True
'
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ' ¼Õ´Ô¿ë
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ' ¸¶Áö¸· ÀåÀÏ°æ¿ì ÀüÃ¼ ÇÕ°è¹× ±Ý¾× Ãâ·Â
'    If bEndflag Then
'        'ÇÕ°è
'        Printer.CurrentX = 500 + Printer_Left
'        Printer.CurrentY = YS + DY * ll
'        frmÁ¢¼ö.sprGrid.Col = 2
'        Printer.Print FPrtBottom.Counter
'
'        '
'        Printer.CurrentX = 700 + Printer_Left
'        Printer.CurrentY = YS + DY * ll
'        frmÁ¢¼ö.sprGrid.Col = 1
'        Printer.Print FPrtBottom.Account0
'
'        '
'        Printer.CurrentX = 900 + Printer_Left
'        Printer.CurrentY = YS + DY * ll
'        frmÁ¢¼ö.sprGrid.Col = 3
'        Printer.Print FPrtBottom.Account1
'
'        '
'        Printer.CurrentX = 1100 + Printer_Left
'        Printer.CurrentY = YS + DY * ll
'        frmÁ¢¼ö.sprGrid.Col = 5
'        Printer.Print FPrtBottom.Account2
'    End If
'
'    '
'    Printer.CurrentX = 1300 + Printer_Left
'    Printer.CurrentY = YS + DY * ll
'    frmÁ¢¼ö.sprGrid.Col = 4
'    Printer.Print FPrtBottom.Addr
'
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ' º¸°ü¿ë
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    'ÇÕ°è
'    ' ¸¶Áö¸· ÀåÀÏ°æ¿ì ÀüÃ¼ ÇÕ°è¹× ±Ý¾× Ãâ·Â
'    If dendflag Then
'        'ÇÕ°è
'        Printer.CurrentX = 500 + Printer_Left
'        Printer.CurrentY = YS + DY * ll
'        frmÁ¢¼ö.sprGrid.Col = 2
'        Printer.Print FPrtBottom.Counter
'
'        '
'        Printer.CurrentX = 700 + Printer_Left
'        Printer.CurrentY = YS + DY * ll
'        frmÁ¢¼ö.sprGrid.Col = 1
'        Printer.Print FPrtBottom.Account0
'
'        '
'        Printer.CurrentX = 900 + Printer_Left
'        Printer.CurrentY = YS + DY * ll
'        frmÁ¢¼ö.sprGrid.Col = 3
'        Printer.Print FPrtBottom.Account1
'
'        '
'        Printer.CurrentX = 1100 + Printer_Left
'        Printer.CurrentY = YS + DY * ll
'        frmÁ¢¼ö.sprGrid.Col = 5
'        Printer.Print FPrtBottom.Account2
'    End If
'    '
'    Printer.CurrentX = 1300 + Printer_Left
'    Printer.CurrentY = YS + DY * ll
'    frmÁ¢¼ö.sprGrid.Col = 4
'    Printer.Print FPrtBottom.Addr
'
'    Screen.MousePointer = 0
'    Exit Function
'
'printError:
'    MsgBox " ÇÁ¸°ÅÍ¸¦ È®ÀÎÇØ ÁÖ½Ê½Ã¿ä ! " & VBA.Err.Number, vbCritical, "Ãâ·Â¿À·ù¹ß»ý"
'
'End Function
Sub FormPrint()
    Dim temp As String
    Dim iCnt As Integer
    Dim iPage As Integer
    Dim iITEM As Integer
    Dim iLine As Integer
    Dim iTxt As String
    Dim htmp As String
    Dim iCnt1 As Integer
    
    For iCnt = LBound(FPArray, 1) To UBound(FPArray, 1)
        If Len(FPArray(iCnt, 1)) < 1 Then
            Exit For
        End If
    Next iCnt
    
    iPage = Fix((iCnt - 2) / 10) + 1
      
    iLine = 1
    iITEM = 0
    On Error GoTo printError

    Open "LPT1" For Output As #1
    'Open App.Path & "\PPP.TXT" For Output As #1

    For iCnt = 1 To iPage
        '¿ëÁö Å©±â ¼³Á¤
        temp = Chr(27) & "C" & Chr(32)
        'ÅÇ¼³Á¤
        temp = temp & Chr(27) & "D" & Chr(7) & Chr(24) & Chr(0)
        'ÀüÈ­¹øÈ£,ÀÌ¸§
        
        '20090113¼öÁ¤
        'temp = temp & vbTab & FPTop.Tel & vbTab & FPTop.Name & vbCr & vbCr & vbLf & vbLf
        temp = temp & vbTab & Right("***************" & Right(Trim(FPTop.Tel), 4), Len(Trim(FPTop.Tel))) & vbTab & FPTop.Name & vbCr & vbCr & vbLf & vbLf
        
        temp = temp & vbTab & FPTop.Date & "  /  " & Left(FPTop.Date2, 2) & "¿ù" & Right(FPTop.Date2, 2) & "ÀÏ" & vbLf & vbLf
        temp = temp & vbTab & FPBottom.Addr & vbLf
        Print #1, temp
            
        
        'Tab Àç¼³Á¤
        Print #1, Chr(27) & "D" & Chr(6) & Chr(18) & Chr(0)
        'Print #1, Chr(27) & "D" & Chr(17) & Chr(22) & Chr(28) & Chr(0)
        
        For iCnt1 = 1 To 10
            If Len(Trim(FPArray(iLine, 1))) = 0 Then
                Print #1, ""
            Else
                temp = ""
                'Tag_no
                temp = temp & FPArray(iLine, 1)
                'Ç°¸í
                temp = temp & vbTab & funLeft(Trim(FPArray(iLine, 2)), 12)
                '»ö»ó
                temp = temp & vbTab & Trim(FPArray(iLine, 3))
                '±Ý¾×
                temp = temp & vbTab & PrNumSet(FPArray(iLine, 4), 6)
                '³»¿ë
                temp = temp & FPArray(iLine, 5)
                
                Print #1, temp
                iITEM = iITEM + 1
                iLine = iLine + 1
            End If
            
        Next iCnt1
        Print #1, ""
        htmp = String(7, " ")
        RSet htmp = Format(FPBottom.Account1, "#,#0")
        
        temp = ""
        If iCnt = iPage Then
            
            'ÇÕ°è
            temp = vbTab & "    " & CStr(iITEM)
            '±Ý¾×
            temp = temp & vbTab & vbTab & "    " & FPBottom.Account0
            temp = temp & vbCr & vbLf & vbLf & vbTab & htmp
            'ÀÜ¾×
            temp = temp & vbTab & vbTab & "    " & Format(FPBottom.Account2, "#,#0")
            '´ë¸®Á¡¸í
            temp = temp & vbCr & vbLf & vbLf & "        TEL:" & frmMain.StatusBar1.Panels(5)
            Print #1, temp
            temp = "        " & frmMain.StatusBar1.Panels(2) & "   " & iCnt & "/" & iPage
            Print #1, temp
        Else
            '´ë¸®Á¡¸í
            temp = vbLf & vbLf & vbLf & vbCr & vbLf & "        TEL:" & frmMain.StatusBar1.Panels(5)
            Print #1, temp
            temp = "        " & frmMain.StatusBar1.Panels(2) & "   " & iCnt & "/" & iPage
            Print #1, temp
        End If
        
        Print #1, Chr(12)
    Next iCnt
    Close #1
    Exit Sub
    
printError:
 '   Close #1
    MsgBox " ÇÁ¸°ÅÍ¸¦ È®ÀÎÇØ ÁÖ½Ê½Ã¿ä ! " & VBA.Err.Number, vbCritical, "Ãâ·Â¿À·ù¹ß»ý"
    
End Sub

Sub FormPrintTest()
    Dim iCnt As Integer
    FPTop.Name = "È«±æµ¿"
    FPTop.Date = "1997-10-11"
    FPTop.Date2 = "11-11"
    FPTop.Tel = "477-1211"
    iCnt = 1
    For iCnt = 1 To 2
    
        FPArray(iCnt, 1) = "0-668"
        FPArray(iCnt, 2) = "1¾Æ°¡Á×"
        FPArray(iCnt, 3) = "ÃÊ·Ï"
        FPArray(iCnt, 4) = "20300"
        FPArray(iCnt, 5) = "µåÇÏ"
    Next iCnt
    
    FPArray(iCnt, 1) = "0-668"
    FPArray(iCnt, 2) = "·¹ÀÌ½º´Þ(¸°)¾Æ°¡Á×"
    FPArray(iCnt, 3) = "ÃÊ·Ï"
    FPArray(iCnt, 4) = "20300"
    FPArray(iCnt, 5) = "µåÇÏ"

    FPBottom.Account0 = "10000"
    FPBottom.Account1 = "2000"
    FPBottom.Account2 = "2000"
    FPBottom.Name = "È«ÀÍÁ¡"
    FPBottom.Sum = "2000"
    FPBottom.Tel = "477-0000"
    FPBottom.Addr = "¿ëÀÎ½Ã ¼öÁöÀ¾ ½ÅºÀ¸®"

    FormPrint
End Sub

Function funLeft(ByVal txt As String, ByVal Length As Integer) As String
    Dim iCnt As Integer
    Dim TrimCnt0, TrimCnt1 As Integer
    Dim iLoop As Integer
    
    iCnt = Len(txt)
    TrimCnt0 = 0
    TrimCnt1 = 0
    
    For iLoop = 1 To iCnt
        If Asc(Mid(txt, iCnt, 1)) > 0 Then
            TrimCnt1 = TrimCnt1 + 1
        Else
            TrimCnt1 = TrimCnt1 + 2
        End If
        If TrimCnt1 > Length Then
            funLeft = MidB(txt, 1, TrimCnt0)
            Exit Function
        Else
            TrimCnt0 = TrimCnt1
        End If
            
    Next iLoop
    funLeft = txt
End Function

Function PrNumSet(Num As Variant, cnt As Integer)

    Dim Num1 As Double
    Dim Str As String

    Num1 = Val(Num)
    Str = "                           " & Format(Num1, "#,##0")
    PrNumSet = Right(Str, cnt)
End Function

Function SetPrtPoint(pt1 As PrintPoint, pt2 As PrintPoint, pt3 As PrintPoint)
    Printer.CurrentX = pt1.x + pt2.x + pt3.x
    Printer.CurrentY = pt1.y + pt2.y + pt3.y
    
    Debug.Print "X => " & Printer.CurrentX; "  , Y => " & Printer.CurrentY
    If Printer.CurrentX > 190 Then
        Debug.Print "X => " & Printer.CurrentX
    End If
    If Printer.CurrentY > 150 Then
        Debug.Print "y => " & Printer.CurrentY
    End If
    
End Function

Function GetPrtStartPoint(strTemp As String) As Integer
    Select Case UCase(strTemp)
        Case "TOP"
            If Val(GetSetting("Laundry_Zi", "Printer", "Top", strTemp)) = 0 Then
                GetPrtStartPoint = 25
                ' ·¹Áö½ºÅÍ¸®¿¡ °ªÀÌ ¾øÀ» °æ¿ì °æ°í ¸Þ½ÃÁö Ãâ·Â
                bMsgMode = True
                strMessage = "´ë¸®Á¡ Á¤º¸¼öÁ¤¿¡¼­ ÇÁ¸°ÅÍÁ¤º¸°¡ µî·ÏµÇ¾îÀÖÁö ¾Ê½À´Ï´Ù." & Chr(10) & Chr(13) & Chr(10) & Chr(13) & "´ë¸®Á¡ Á¤º¸¸¦ È®ÀÎÇÏ¿© ÁÖ½Ê½Ã¿ä."
            Else
                GetPrtStartPoint = Val(GetSetting("Laundry_Zi", "Printer", "Top", strTemp))
            End If
        Case "LEFT"
            If Val(GetSetting("Laundry_Zi", "Printer", "Left", strTemp)) = 0 Then
                GetPrtStartPoint = 1
            Else
                GetPrtStartPoint = Val(GetSetting("Laundry_Zi", "Printer", "Left", strTemp))
            End If
        Case "HEIGHT"
            If Val(GetSetting("Laundry_Zi", "Printer", "Height", strTemp)) = 0 Then
                GetPrtStartPoint = 4
            Else
                GetPrtStartPoint = Val(GetSetting("Laundry_Zi", "Printer", "Height", strTemp))
            End If
        Case Else
            MsgBox "GetPrtStartPoint -> ±âº» ¿©¹é À§Ä¡ ¾ò±â ¿À·ù", vbInformation, "Error"
    End Select
    
    ' ·¹Áö½ºÅÍ¸®¿¡ °ªÀÌ ¾øÀ» °æ¿ì °æ°í ¸Þ½ÃÁö Ãâ·Â
End Function

Function GetPrtItemCount(Index As String)
    Select Case UCase(Index)
        Case "º¸°üÁõ"
            GetPrtItemCount = IIf(Printer_BO_Gb = "0", 10, 11)
        Case "ÀÏÀÏ¸ÅÃâÇöÈ²"
            GetPrtItemCount = 50
        Case "¿ùº°¸ÅÃâÇöÈ²"
            GetPrtItemCount = 31
        Case Else
            Debug.Print Index & " GetPrtItemCount ÇÔ¼ö =>ÆäÀÌÁö´ç Ãâ·Â ¿À·ù"
    End Select
End Function

Function GetPrtGubun() As Integer
    ' À×Å©Á¬ÀÏ °æ¿ì ¿©¹é
    ' 0 = µµÆ® ÇÁ¸°ÅÍ , 1= À×Å©Á¬, 3= ±âÅ¸
    
    Query = "SELECT ÇÁ¸°ÅÍ FROM ´ë¸®Á¡Á¤º¸"
    Set Rs = New ADODB.Recordset
    Rs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If Rs.RecordCount > 0 Then
        GetPrtGubun = Val(Rs!ÇÁ¸°ÅÍ & "")
    Else
        GetPrtGubun = 0
    End If
    
    Rs.Close
    Set Rs = Nothing
End Function

Function GetPrtBOGubun() As Integer
    ' 0 = ±âÁ¸ º¸°üÁõ , 1= ½Å±Ô º¸°üÁõ
    
    Query = "SELECT º¸°üÁõÁ¾·ù FROM ´ë¸®Á¡Á¤º¸"
    Set Rs = New ADODB.Recordset
    Rs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If Rs.RecordCount > 0 Then
        GetPrtBOGubun = Val(Rs!º¸°üÁõÁ¾·ù & "")
    Else
        GetPrtBOGubun = 0
    End If
    Rs.Close
    Set Rs = Nothing
End Function


Function subinkPrint2(cdPrt As CommonDialog, prtNum As String, prtTel As String)
    Dim Page_Count As Integer       ' º¸°üÁõ¿¡ Ãâ·ÂµÉ »óÇ°ÀÇ ÃÑ °¹¼ö
    Dim sPage_count As Integer      ' º¸°üÁõÀÇ  ÀüÃ¼ ÆäÀÌÁö¼ö
    Dim Page_Item_Count As Integer  ' ÇÑÆäÀÌÁö¿¡ Ãâ·ÂµÉ »óÇ°ÀÇ °¹¼ö

    Dim dXOffSet As Integer
    Dim dYOffSet As Integer
    
    Dim tmpKEY2 As String
    Dim tmpKEY
    Dim tmpCOD1 '(1 To tmpListCNT)
    Dim tmpAC1 '(1 To tmpListCNT)
    Dim tmpCOD2 '(1 To tmpListCNT)
    Dim tmpAC2 '(1 To tmpListCNT)

    Dim tmpSUSUN '(1 To tmpListCNT)
    Dim tmpCOL  As Long '(1 To tmpListCNT)

    Dim tmpBI1 '(1 To tmpListCNT)
    Dim tmpBIS '(1 To tmpListCNT)

    Dim tmpMON  As Long '(1 To tmpListCNT) As Long
    Dim tmpVAL  As Long
    
    Dim S_Line As Integer
    Dim L_Line As Integer
    Dim GRD_TOT As Integer
    Dim GRD_S_TOT As Integer
    Dim L_Page As Integer
    Dim i As Integer
    Dim j As Integer
    Dim ll As Integer
    Dim SUB_TOT As Integer
    
    
    ''''''''''''''''
    On Error GoTo printError
    '''''''''''''''
Print_Start:

    cdPrt.Flags = cdlPDHidePrintToFile
    'CommonDialog1.Action = 5

    ' »ç¿ë °ªµéÀ» ÃÊ±âÈ­ ÇÑ´Ù.
    L_Page = 0
    S_Line = 0
    L_Line = 0
    GRD_TOT = 0
    GRD_S_TOT = 0
    
    Page_Item_Count = GetPrtItemCount("º¸°üÁõ") ' º¸°üÁõ¿¡ Ãâ·ÂµÉ »óÇ° °¹¼ö
   
   
    Printer.ScaleMode = vbPixels            ' ÇÁ¸°ÅÍÀÇ ½ºÄÉÀÏ ¸ðµå¸¦ ¹Ð¸®¹ÌÅÍ ´ÜÀ§·Î
    Printer.Width = 2000
    Printer.Height = 1485
'        Printer.ScaleWidth = 2000
'        Printer.ScaleHeight = 1485
    Printer.FontName = "±¼¸²"
    Printer.Font.Bold = True
    Printer.Font.Size = 9
    
    'ÀüÃ¼ Ãâ·Â °¹¼ö¹× Ãâ·Â ³»¿ë º¯¼ö¿¡ ÃÊ±âÈ­
    GoSub Print_Value_Init
    
    If (Page_Count <= 0) Then
        Exit Function
    End If

    ' ÀüÃ¼ Ãâ·Â ÆäÀÌÁö ±¸ÇÏ±â
    If (Page_Count Mod Page_Item_Count) <> 0 Then
        sPage_count = Int(Page_Count / Page_Item_Count) + 1
    Else
        sPage_count = Int(Page_Count / Page_Item_Count)
    End If
    
    'ÀüÃ¼ ÆäÀÌÁö ±îÁö ¹Ýº¹.
    For L_Page = 1 To sPage_count
        ' Ã¹¹øÂ° ÀåÀÌ³ª ¸¶Áö¸· ÀåÀÏ°æ¿ì
        If L_Page = sPage_count Or sPage_count = 1 Then
            S_Line = L_Line + 1
            L_Line = Page_Count   ' frmINPUT.ListView1.ListItems.Count
            'À×Å©Á¬
            GoSub Print_Title
            GoSub Print_Center
            GoSub Print_Bottom
            Printer.EndDoc
            Exit For
        Else
        ' Áß°£ ÆäÀÌÁö ÀÏ °æ¿ì
            S_Line = L_Line + 1
            L_Line = L_Line + Page_Item_Count
            'À×Å©Á¬
            GoSub Print_Title
            GoSub Print_Center
            GoSub Print_Bottom
            Printer.NewPage
        End If
    Next L_Page

    ''''''''''''''''
    'On Error Resume Next
    Screen.MousePointer = 0
    Exit Function
    
'-------------------------------------------------------------------------------
'--   Ãâ·Â°ª ÃÊ±âÈ­
'-------------------------------------------------------------------------------
Print_Value_Init:
    
' º¸°üÁõ Ãâ·Â »ó´Ü ÀÚ·á ÃÊ±âÈ­
    Query = "SELECT * "
    Query = Query & "FROM º¸°üÁõ "
    Query = Query & "WHERE ÀÏ·Ã¹øÈ£ = " & Val(prtNum) & " "
    Query = Query & "AND °í°´ÀüÈ­ = '" & prtTel & "' "
    Query = Query & "ORDER BY ÅÃ¹øÈ£"
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If SUBRs.RecordCount > 0 Then
        SUBRs.MoveLast
        Page_Count = SUBRs.RecordCount
        SUBRs.MoveFiSUBRst
    Else
        SUBRs.Close
        Debug.Print "º¸°üÁõ Ãâ·Â ¾øÀ½. (¿À·ù)"
        Return
    End If


    FPrtTop.PrtNo = Format(Date, "MMDD") & "-" & SUBRs!ÀÏ·Ã¹øÈ£
    FPrtTop.Tel = SUBRs!°í°´ÀüÈ­
    FPrtTop.Name = SUBRs!¼º¸í
    FPrtTop.Addr = SUBRs!´ë¸®Á¡¸í '°í°´ ÁÖ¼ÒÀÓ
    FPrtTop.Date = SUBRs!Á¢¼öÀÏ
    
    If SUBRs!Á¢¼öÀÏ > (Format(Date, "YYYY") & "-" & Format(SUBRs!ÀÎµµ¿¹Á¤ÀÏ, "00-00")) Then
        FPrtTop.Date2 = Format(DateAdd("yyyy", 1, Date), "YYYY") & "-" & Format(SUBRs!ÀÎµµ¿¹Á¤ÀÏ, "00-00")
    Else
        FPrtTop.Date2 = Format(Date, "YYYY") & "-" & Format(SUBRs!ÀÎµµ¿¹Á¤ÀÏ, "00-00")
    End If
    
    FPrtTop.Code = Fb°í°´¹øÈ£(FPrtTop.Name, Left(FPrtTop.Tel, InStr(FPrtTop.Tel, "-") - 1), Right(FPrtTop.Tel, 4))
    
    ' º¸°üÁõ Ãâ·Â ÇÏ´Ü ÀÚ·á ÃÊ±âÈ­
    FPrtBottom.Sum = SUBRs!ÇÕ°è
    FPrtBottom.Account0 = SUBRs!ÇÕ°è±Ý¾×
    FPrtBottom.Account1 = Format(Val(CStr(SUBRs!¼ö·É¾×)) + Val(CStr(SUBRs!¸¶ÀÏ¸®Áö)), "#,##0")
    FPrtBottom.Account2 = Format(SUBRs!ÀÜ¾×, "#,#0")
    
    '----------------------------------------------------------------
    '
    '----------------------------------------------------------------
    Query = "SELECT * FROM ´ë¸®Á¡Á¤º¸ "
    Set Rs = MyDB.OpenRecordset(Query)
    
    If Rs.RecordCount < 1 Then
        Debug.Print ("Ãâ·Â ´ë¸®Á¡ Á¤º¸ ´ë¸®Á¡¸í, ÀüÈ­ ¹øÈ£ ºÎÁ·")
    Else
        FPrtBottom.DName = Rs!´ë¸®Á¡¸í
        FPrtBottom.DTel = Rs!telStore & ""
    End If
    Rs.Close
    Set Rs = Nothing
    '-----------------------------------------------------------------
    
' º¸°üÁõ Ãâ·Â Áß°£ ÀÚ·á ÃÊ±âÈ­
    For i = 1 To 1000
        FPArray(i, 1) = SUBRs!ÅÃ¹øÈ£
        FPArray(i, 2) = SUBRs!Ç°¸í
        FPArray(i, 3) = SUBRs!»ö»ó
        FPArray(i, 4) = Format(SUBRs!±Ý¾×, "#,#0")
        FPArray(i, 5) = SUBRs!³»¿ë
        FPArray(i, 6) = SUBRs!»óÇ¥

        SUBRs.MoveNext

        If SUBRs.EOF = True Then
            Exit For
        End If
    Next i

    SUBRs.Close
    Set SUBRs = Nothing
    
    Return
'-------------------------------------------------------------------------------
'--   Å¸ÀÌÆ² ºÎºÐ Ãâ·Â
'-------------------------------------------------------------------------------
Print_Title:
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '  ´ë¸®Á¡ º¸°ü¿ë
    
    ' ±âº» ¿©¹éÀ» °¡Àú¿Â´Ù
    PrtPoint4 = GetPrtPoint("¿©¹é")
    
    For j = 0 To 1
        If j = 1 Then
            PrtPoint2 = GetPrtPoint("¼Õ´Ô¿ë")
        Else
            PrtPoint2.x = 0
            PrtPoint2.y = 0
        End If
        
        ' ÀüÇ¥ ¹øÈ£
        PrtPoint = GetPrtPoint("PRTNO")
        SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
        Printer.Print FPrtTop.PrtNo
        ' °í°´ ÀüÈ­¹øÈ£
        PrtPoint = GetPrtPoint("GTEL")
        SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
        Printer.Print FPrtTop.Tel
        ' °í°´ ¼º¸í
        PrtPoint = GetPrtPoint("GNAME")
        SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
        Printer.Print FPrtTop.Name
        ' ÁÖ¼Ò (¼Õ´Ô)
        PrtPoint = GetPrtPoint("ADDR")
        SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
        Printer.Print FPrtTop.Addr
        ' Á¢¼öÀÏ
        PrtPoint = GetPrtPoint("DATE")
        SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
        Printer.Print FPrtTop.Date
        ' °í°´ ¹øÈ£
        PrtPoint = GetPrtPoint("CODE")
        SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
        Printer.Print FPrtTop.Code
        ' ÀÎµµ ¿¬µµ
        PrtPoint = GetPrtPoint("DATE2")
        SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
        Printer.Print FPrtTop.Date2
    Next j
    Return
'-------------------------------------------------------------------------------
'--   ³»¿ë ºÎºÐ Ãâ·Â
'-------------------------------------------------------------------------------
Print_Center:
    
    
    ll = 0 ' º¸°üÁõ Ãâ·Â ¶óÀÎ ÃÊ±âÈ­
    If (S_Line + Page_Item_Count) > Page_Count Then
        SUB_TOT = Page_Count
    Else
        SUB_TOT = S_Line + Page_Item_Count - 1
    End If
    
    For i = S_Line To SUB_TOT
        ll = ll + 1
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' º¸°ü¿ë
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' ±âº» ¿©¹éÀ» °¡Àú¿Â´Ù
        PrtPoint4 = GetPrtPoint("¿©¹é")
        PrtPoint3 = GetPrtPoint("NEXT_LINE")
        For j = 0 To 1
            If j = 1 Then
                PrtPoint2 = GetPrtPoint("¼Õ´Ô¿ë")
            Else
                PrtPoint2.x = 0
                PrtPoint2.y = 0
            End If
        
            'ÅÃ¹øÈ£
            PrtPoint = GetPrtPoint("TAGNUM")
            Printer.CurrentY = PrtPoint4.y + PrtPoint2.y + PrtPoint.y + (PrtPoint3.y * (ll - 1))
            Printer.CurrentX = PrtPoint4.x + PrtPoint2.x + PrtPoint.x + PrtPoint3.x
            Printer.Print FPArray(i, 1)
            
            'Ç°¸í
            PrtPoint = GetPrtPoint("PNAME")
            Printer.CurrentY = PrtPoint4.y + PrtPoint2.y + PrtPoint.y + (PrtPoint3.y * (ll - 1))
            Printer.CurrentX = PrtPoint4.x + PrtPoint2.x + PrtPoint.x + PrtPoint3.x
            Printer.Print FPArray(i, 2)
            
            '»ö»ó
            PrtPoint = GetPrtPoint("PCOLOR")
            Printer.CurrentY = PrtPoint4.y + PrtPoint2.y + PrtPoint.y + (PrtPoint3.y * (ll - 1))
            Printer.CurrentX = PrtPoint4.x + PrtPoint2.x + PrtPoint.x + PrtPoint3.x
            Printer.Print FPArray(i, 3)
            
            '±Ý¾×
            PrtPoint = GetPrtPoint("PACCOUNT")
            Printer.CurrentY = PrtPoint4.y + PrtPoint2.y + PrtPoint.y + (PrtPoint3.y * (ll - 1))
            Printer.CurrentX = PrtPoint4.x + PrtPoint2.x + PrtPoint.x + PrtPoint3.x
            Printer.Print FPArray(i, 4)
            
            '³»¿ë
            PrtPoint = GetPrtPoint("PTEMP")
            Printer.CurrentY = PrtPoint4.y + PrtPoint2.y + PrtPoint.y + (PrtPoint3.y * (ll - 1))
            Printer.CurrentX = PrtPoint4.x + PrtPoint2.x + PrtPoint.x + PrtPoint3.x
            Printer.Print FPArray(i, 5)
            
            '»óÇ¥
            PrtPoint = GetPrtPoint("BRAND")
            Printer.CurrentY = PrtPoint4.y + PrtPoint2.y + PrtPoint.y + (PrtPoint3.y * (ll - 1))
            Printer.CurrentX = PrtPoint4.x + PrtPoint2.x + PrtPoint.x + PrtPoint3.x
            Printer.Print FPArray(i, 6)
        Next j
    Next i
    Return

'-------------------------------------------------------------------------------
'--   ³¡ ºÎºÐ Ãâ·Â
'-------------------------------------------------------------------------------
Print_Bottom:
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' º¸°ü¿ë
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ' ±âº» ¿©¹éÀ» °¡Àú¿Â´Ù
    PrtPoint4 = GetPrtPoint("¿©¹é")
    For j = 0 To 1
        If j = 1 Then
            PrtPoint2 = GetPrtPoint("¼Õ´Ô¿ë")
        Else
            PrtPoint2.x = 0
            PrtPoint2.y = 0
        End If
        
        ' ¸¶Áö¸· ÀåÀÏ°æ¿ì ÀüÃ¼ ÇÕ°è¹× ±Ý¾× Ãâ·Â
        If L_Page = sPage_count Or sPage_count = 1 Then
            ' Á¡¼ö
            PrtPoint = GetPrtPoint("SUM")
            SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
            Printer.Print FPrtBottom.Sum
            '±Ý¾×
            PrtPoint = GetPrtPoint("ACCOUNT0")
            SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
            Printer.Print FPrtBottom.Account0
            ' ¼ö·É¾×
            PrtPoint = GetPrtPoint("ACCOUNT1")
            SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
            Printer.Print FPrtBottom.Account1
            'ÀÜ¾×
            PrtPoint = GetPrtPoint("ACCOUNT2")
            SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
            Printer.Print FPrtBottom.Account2
        End If
        
        ' ´ë¸®Á¡¸í
        PrtPoint = GetPrtPoint("DNAME")
        SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
        Printer.Print FPrtBottom.DName
        ' ´ë¸®Á¡ ÀüÈ­¹øÈ£
        PrtPoint = GetPrtPoint("DTEL")
        SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
        Printer.Print FPrtBottom.DTel
        ' ÆäÀÌÁö/ÀüÃ¼ ÆäÀÌÁö
        PrtPoint = GetPrtPoint("PAGE")
        SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
        Printer.Print L_Page & "/" & sPage_count
    Next j
    
    Return

'-------------------------------------------------------------------------------
'--   ÀÎ¼âÁß ¿À·ù ½ÇÇà ºÎºÐ
'-------------------------------------------------------------------------------
printError:
    MsgBox " ÇÁ¸°ÅÍ¸¦ È®ÀÎÇØ ÁÖ½Ê½Ã¿ä ! " & VBA.Err.Number, vbCritical, "Ãâ·Â¿À·ù¹ß»ý"
    
End Function


Function subinkPrintMM(cdPrt As CommonDialog, prtNum As String, prtTel As String)
' ±âº»¼³Á¤ 25,1,5

    Dim strMaxLng   As String
    Dim strTempStr  As String
    
    Dim Page_Count As Integer       ' º¸°üÁõ¿¡ Ãâ·ÂµÉ »óÇ°ÀÇ ÃÑ °¹¼ö
    Dim sPage_count As Integer      ' º¸°üÁõÀÇ  ÀüÃ¼ ÆäÀÌÁö¼ö
    Dim Page_Item_Count As Integer  ' ÇÑÆäÀÌÁö¿¡ Ãâ·ÂµÉ »óÇ°ÀÇ °¹¼ö

    Dim dXOffSet As Integer
    Dim dYOffSet As Integer
    
    Dim tmpKEY2 As String
    Dim tmpKEY
    Dim tmpCOD1 '(1 To tmpListCNT)
    Dim tmpAC1 '(1 To tmpListCNT)
    Dim tmpCOD2 '(1 To tmpListCNT)
    Dim tmpAC2 '(1 To tmpListCNT)

    Dim tmpSUSUN '(1 To tmpListCNT)
    Dim tmpCOL  As Long '(1 To tmpListCNT)

    Dim tmpBI1 '(1 To tmpListCNT)
    Dim tmpBIS '(1 To tmpListCNT)

    Dim tmpMON  As Long '(1 To tmpListCNT) As Long
    Dim tmpVAL  As Long
    
    Dim S_Line As Integer
    Dim L_Line As Integer
    Dim GRD_TOT As Integer
    Dim GRD_S_TOT As Integer
    Dim L_Page As Integer
    Dim i As Integer
    Dim j As Integer
    Dim ll As Integer
    Dim SUB_TOT As Integer
    
    Dim zz As PrintPoint
    
    ' ±âº» ÇÁ¸°ÅÍ°¡ ¾øÀ» °æ¿ì
    If Not PrinterCheck Then Exit Function
        
   
    ''''''''''''''''
    On Error GoTo printError
    '''''''''''''''
Print_Start:

    cdPrt.Flags = cdlPDHidePrintToFile
    'CommonDialog1.Action = 5

    ' »ç¿ë °ªµéÀ» ÃÊ±âÈ­ ÇÑ´Ù.
    L_Page = 0
    S_Line = 0
    L_Line = 0
    GRD_TOT = 0
    GRD_S_TOT = 0
    
    Erase FPArray
    
    Page_Item_Count = GetPrtItemCount("º¸°üÁõ")     ' º¸°üÁõ¿¡ Ãâ·ÂµÉ »óÇ° °¹¼ö
   
    ' À×Å©Á¬ ÇÁ¸°ÅÍ
    If Printer_Gb = "1" Then
        Printer.ScaleMode = vbMillimeters           ' ÇÁ¸°ÅÍÀÇ ½ºÄÉÀÏ ¸ðµå¸¦ ¹Ð¸®¹ÌÅÍ ´ÜÀ§·Î
        Printer.Width = 19 * 567
        Printer.Height = 15 * 567
        Printer.FontName = "±¼¸²Ã¼"
        Printer.Font.Bold = True
        Printer.Font.Size = 9
        Printer.DrawWidth = 1
    
    ' ·¹ÀÌÀú ÇÁ¸°ÅÍ
    ElseIf Printer_Gb = "2" Then
        Printer.ScaleMode = vbMillimeters           ' ÇÁ¸°ÅÍÀÇ ½ºÄÉÀÏ ¸ðµå¸¦ ¹Ð¸®¹ÌÅÍ ´ÜÀ§·Î
        Printer.FontName = "±¼¸²Ã¼"
        Printer.Font.Bold = True
        Printer.Font.Size = 9
        Printer.DrawWidth = 1
    
    End If

    'ÀüÃ¼ Ãâ·Â °¹¼ö¹× Ãâ·Â ³»¿ë º¯¼ö¿¡ ÃÊ±âÈ­
    GoSub Print_Value_Init
    If (Page_Count <= 0) Then
        Exit Function
    End If

    '----------------------------------------------------
    ' ¼¼Æ® °ü·Ã ÃÖÁ¾ Ãâ·Â ³»¿ëÀÇ 4Ä­À» ÇÒ´ç ÇÏ¿© ¼¼Æ® ³»¿ëÀ» Ãâ·ÂÇÑ´Ù.
    If FPrtTop.Date <= "2009-12-31" Then
        If m_GSGMoney.d¼¼Æ®¼ö·®ÇÕ°è > 0 Then Page_Count = Page_Count + 6
    Else
        If m_GSGMoney.d¼¼Æ®¼ö·®ÇÕ°è > 0 Then Page_Count = Page_Count + 3
    End If
    '----------------------------------------------------
    
    ' ÀüÃ¼ Ãâ·Â ÆäÀÌÁö ±¸ÇÏ±â
    If (Page_Count Mod Page_Item_Count) <> 0 Then
        sPage_count = Int(Page_Count / Page_Item_Count) + 1
    Else
        sPage_count = Int(Page_Count / Page_Item_Count)
    End If
    
    'ÀüÃ¼ ÆäÀÌÁö ±îÁö ¹Ýº¹.
    For L_Page = 1 To sPage_count
    
        ' Ã¹¹øÂ° ÀåÀÌ³ª ¸¶Áö¸· ÀåÀÏ°æ¿ì
        If L_Page = sPage_count Or sPage_count = 1 Then
            S_Line = L_Line + 1
            L_Line = Page_Count   ' frmINPUT.ListView1.ListItems.Count
            'À×Å©Á¬
            GoSub Print_Title
            GoSub Print_Center
            
            ' ¼¼Æ® »óÇ° °ü·Ã ³»¿ë Ãâ·Â
            GoSub Print_GropGoodsINFO
            
            GoSub Print_Bottom
            Printer.EndDoc
            Exit For
        Else
        ' Áß°£ ÆäÀÌÁö ÀÏ °æ¿ì
            S_Line = L_Line + 1
            L_Line = L_Line + Page_Item_Count
            'À×Å©Á¬
            GoSub Print_Title
            GoSub Print_Center
            
            GoSub Print_Bottom
            Printer.NewPage
        End If
    Next L_Page

    ''''''''''''''''
    'On Error Resume Next
    Screen.MousePointer = 0
    Exit Function
    
'-------------------------------------------------------------------------------
'--   Ãâ·Â°ª ÃÊ±âÈ­
'-------------------------------------------------------------------------------
Print_Value_Init:
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
    ' º¸°üÁõ Ãâ·Â »ó´Ü ÀÚ·á ÃÊ±âÈ­
    '--------------------------------------------------------------
    Query = "SELECT * FROM º¸°üÁõ "
    Query = Query & " WHERE ÀÏ·Ã¹øÈ£ = " & Val(prtNum) & " "
    Query = Query & "   AND °í°´ÀüÈ­ = '" & prtTel & "' "
    Query = Query & " ORDER BY ÅÃ¹øÈ£"
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If SUBRs.RecordCount > 0 Then
        SUBRs.MoveLast
        Page_Count = SUBRs.RecordCount
        SUBRs.MoveFirst
    Else
        SUBRs.Close
        Set SUBRs = Nothing
        
        Debug.Print "º¸°üÁõ Ãâ·Â ¾øÀ½. (¿À·ù)"
        Return
    End If


    FPrtTop.PrtNo = Format(Date, "MMDD") & "-" & SUBRs!ÀÏ·Ã¹øÈ£
    
    '2009-04-02ÀÏ ´Ù½Ã ¼öÁ¤  20090113 ¼öÁ¤»çÇ×
    If ´ë¸®Á¡Á¤º¸.°í°´ÀüÈ­¹øÈ£¸ðµÎÃâ·Â = "0" Then
        FPrtTop.Tel = Right("***************" & Right(Trim(SUBRs!°í°´ÀüÈ­), 4), Len(Trim(SUBRs!°í°´ÀüÈ­)))
    Else
        FPrtTop.Tel = SUBRs!°í°´ÀüÈ­
    End If
    
    FPrtTop.Name = SUBRs!¼º¸í
    FPrtTop.Addr = SUBRs!´ë¸®Á¡¸í '°í°´ ÁÖ¼ÒÀÓ
    FPrtTop.Date = SUBRs!Á¢¼öÀÏ
    
    If SUBRs!Á¢¼öÀÏ > (Format(Date, "YYYY") & "-" & Format(SUBRs!ÀÎµµ¿¹Á¤ÀÏ, "00-00")) Then
        FPrtTop.Date2 = Format(DateAdd("yyyy", 1, Date), "YYYY") & "-" & Format(SUBRs!ÀÎµµ¿¹Á¤ÀÏ, "00-00")
    Else
        FPrtTop.Date2 = Format(Date, "YYYY") & "-" & Format(SUBRs!ÀÎµµ¿¹Á¤ÀÏ, "00-00")
    End If
    
    ' ÀüÈ­ ¹øÈ£ÀÇ ±¹¹øÀÌ 3ÀÚ¸®ÀÏ °æ¿ì ¿À¸¥ÂÊ "@@@ "·Î Àü´ÞµÇ´Â °ÍÀ» ¹æÁö ÇÏ±âÀ§ÇÏ¿© trim »ç¿ë
    FPrtTop.Code = Fb°í°´¹øÈ£(FPrtTop.Name, Left(SUBRs!°í°´ÀüÈ­, InStr(SUBRs!°í°´ÀüÈ­, "-") - 1), Right(Trim(SUBRs!°í°´ÀüÈ­), 4))
    
    Call Fb°í°´Á¤º¸(FPrtTop.Code)
    
    '2009-04-02ÀÏ ´Ù½Ã ¼öÁ¤  20090113 ¼öÁ¤»çÇ×
    If ´ë¸®Á¡Á¤º¸.°í°´ÀüÈ­¹øÈ£¸ðµÎÃâ·Â = "0" Then
        FPrtTop.HpTel = Right("***************" & Right(Trim(°í°´Á¤º¸.ÈÞ´ëÆù), 4), Len(Trim(°í°´Á¤º¸.ÈÞ´ëÆù)))
    Else
        FPrtTop.HpTel = °í°´Á¤º¸.ÈÞ´ëÆù
    End If
    'FPrtTop.Tel = Right("***************" & Right(SUBRs!°í°´ÀüÈ­, 4), Len(SUBRs!°í°´ÀüÈ­))
    
    
' º¸°üÁõ Ãâ·Â ÇÏ´Ü ÀÚ·á ÃÊ±âÈ­
    strMaxLng = "1234567890"
    
    With FPrtBottom
        .Sum = strMaxLng
        RSet .Sum = RTrim(SUBRs!ÇÕ°è)
        .Account0 = strMaxLng
        RSet .Account0 = RTrim(SUBRs!ÇÕ°è±Ý¾×)
        
        .Account1 = strMaxLng & "12345"
        If Val(CStr(SUBRs!¸¶ÀÏ¸®Áö)) = 0 Then
            RSet .Account1 = Format(Val(CStr(SUBRs!¼ö·É¾×)), "#,##0")
        Else
            RSet .Account1 = Format(Val(CStr(SUBRs!¼ö·É¾×)), "#,##0") & "/" & Format(Val(CStr(SUBRs!¸¶ÀÏ¸®Áö)), "#,##0")
        End If
        
        .Account2 = strMaxLng
        RSet .Account2 = Format(SUBRs!ÀÜ¾×, "#,#0")
    
        .MiSuTotal = strMaxLng
        RSet .MiSuTotal = Format(SUBRs!¹Ì¼öÇÕ°è, "#,#0") 'Format(°í°´Á¤º¸.¹Ì¼ö±Ý, "#,#0")
        .OldDayMisu = strMaxLng
        RSet .OldDayMisu = Format(SUBRs!ÀüÀÏ¹Ì¼ö, "#,#0") '°í°´Á¤º¸.¹Ì¼ö±Ý - SUBRs!ÀÜ¾×
        .SuGumMonye = strMaxLng
        RSet .SuGumMonye = Format(SUBRs!¼ö±Ý¾×, "#,#0")
    
    ' »ç¿ë¸¶ÀÏ¸®Áö, ¸¶ÀÏ¸®Áö ÀÜ¾×, ´©Àû ¸¶ÀÏ¸®Áö
        .MilMoney = strMaxLng
        RSet .MilMoney = Format(SUBRs!¸¶ÀÏ¸®ÁöÀÜ¾×, "#,#0") ' Format(userMileage.ÀÜ¾×, "#,##0")
        .MilUser = strMaxLng
        RSet .MilUser = Format(SUBRs!¸¶ÀÏ¸®Áö, "#,##0")
        
        .MilAddMoney = strMaxLng
        RSet .MilAddMoney = Format(GetMileageMoneyToPoint(SUBRs!´©Àû¸¶ÀÏ¸®Áö & ""), "#,#0")
        ' 20090529ÀÏ ¼öÁ¤Àü ¿ø¹®..
        ' ¼öÁ¤ ÀÌÀ¯ : ´©Àû¸¶ÀÏ¸®Áö ³»¿ë Ãâ·ÂÀ» ÃÖÁ¾ ¹ß»ý ±Ý¾×¿¡ ÇØ´çÇÏ´Â ºñÀ²ÀÇ Æ÷ÀÎÆ®·Î Ãâ·Â ÇÏµµ·Ï º¯°æ
        'RSet .MilAddMoney = Format(SUBRs!´©Àû¸¶ÀÏ¸®Áö, "#,#0") 'Format(userMileage.ÃÑ»ç¿ë±Ý¾×, "#,##0")
                    
        .DName = ´ë¸®Á¡Á¤º¸.´ë¸®Á¡¸í
        .DTel = ´ë¸®Á¡Á¤º¸.ÀüÈ­¸ÅÀå
        
        .CouponCnt = Format(SUBRs!CouponCnt, "#,#0")
        .CouponNum = Format(SUBRs!CouponNumber, "#,#0")
        .CouponMoney = Format(SUBRs!CouponMoney, "#,#0")
    End With
    
' º¸°üÁõ Ãâ·Â Áß°£ ÀÚ·á ÃÊ±âÈ­
    For i = 1 To 500
        FPArray(i, 1) = SUBRs!ÅÃ¹øÈ£ & ""
        FPArray(i, 2) = SUBRs!Ç°¸í & ""
        FPArray(i, 3) = SUBRs!»ö»ó & ""
        FPArray(i, 4) = Format(SUBRs!±Ý¾×, "#,#0") & ""
        FPArray(i, 5) = SUBRs!³»¿ë & ""
        FPArray(i, 6) = SUBRs!»óÇ¥ & ""

        SUBRs.MoveNext

        If SUBRs.EOF = True Then
            Exit For
        End If
    Next i
    
    ' ¼¼Æ® »óÇ°ÀÇ ³»¿ªÀ» °¡Àú¿Â´Ù.
    SUBRs.MoveFirst
    
    ZeroMemory m_GSGMoney, Len(m_GSGMoney)
    
    '--------------------------------------------------------------------
    '
    '--------------------------------------------------------------------
    Query = "SELECT * FROM ¼¼Æ®»óÇ°Á¤º¸ "
    Query = Query & "WHERE ¼¼Æ®Key = '" & CStr(SUBRs!¼¼Æ®Key & "") & "' "
    Set Rs = New ADODB.Recordset
    Rs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If Rs.RecordCount > 0 Then
        With m_GSGMoney
            .d¼¼Æ®Key = Rs.Fields("¼¼Æ®Key") & ""
            
            .d2¼¼Æ®¼ö·® = Val(Rs.Fields("¼¼Æ®2") & "")
            .d3¼¼Æ®¼ö·® = Val(Rs.Fields("¼¼Æ®3") & "")
            .d4¼¼Æ®¼ö·® = Val(Rs.Fields("¼¼Æ®4") & "")
            .d5¼¼Æ®¼ö·® = Val(Rs.Fields("¼¼Æ®5") & "")
            .d6¼¼Æ®¼ö·® = Val(Rs.Fields("¼¼Æ®6") & "")
            
            .d¼¼Æ®¼ö·®ÇÕ°è = .d2¼¼Æ®¼ö·® + .d3¼¼Æ®¼ö·® + .d4¼¼Æ®¼ö·® + .d5¼¼Æ®¼ö·® + .d6¼¼Æ®¼ö·®
            .d¹«·á¼¼Å¹±Ç¼ö·® = (.d2¼¼Æ®¼ö·® * 1) + _
                             (.d3¼¼Æ®¼ö·® * 2) + _
                             (.d4¼¼Æ®¼ö·® * 3) + _
                             (.d5¼¼Æ®¼ö·® * 4) + _
                             (.d6¼¼Æ®¼ö·® * 5)
            
            .dÀüÃ¼±Ý¾× = Val(Rs.Fields("Á¤»ó±Ý¾×") & "")
            .d¼¼Æ®±Ý¾× = Val(Rs.Fields("¼¼Æ®±Ý¾×") & "")
            
            .d¼¼Æ®ÇÒÀÎ±Ý¾× = Val(Rs.Fields("¼¼Æ®ÇÒÀÎ±Ý¾×") & "")
            .d¿¡´©¸®ÇÒÀÎ±Ý¾× = Val(Rs.Fields("¿¡´©¸®ÇÒÀÎ±Ý¾×") & "")
            .dÀüÃ¼ÇÒÀÎ±Ý¾× = .d¼¼Æ®ÇÒÀÎ±Ý¾× + .d¿¡´©¸®ÇÒÀÎ±Ý¾×
            .dÃÖÁ¾¼ö·É¾× = Val(Rs.Fields("Àû¿ëÇÕ°è±Ý¾×") & "")
        End With
     End If
    Rs.Close
    Set Rs = Nothing
    
    m_¼¼Æ®ÀÀ¸ð¹øÈ£¼ö·® = 0
    
    '--------------------------------------------------------------------
    '
    '--------------------------------------------------------------------
    Query = "SELECT * FROM ¼¼Æ®ÀÀ¸ð¹øÈ£ "
    Query = Query & " WHERE ¼¼Æ®Key = '" & CStr(SUBRs!¼¼Æ®Key & "") & "' "
    Set Rs = New ADODB.Recordset
    Rs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If Rs.RecordCount > 0 Then
        Rs.MoveLast:    ReDim m_¼¼Æ®ÀÀ¸ð¹øÈ£(Rs.RecordCount - 1)
        Rs.MoveFirst
        
        Do While Not Rs.EOF
            m_¼¼Æ®ÀÀ¸ð¹øÈ£(m_¼¼Æ®ÀÀ¸ð¹øÈ£¼ö·®) = Rs.Fields("ÀÀ¸ð¹øÈ£") & ""
            m_¼¼Æ®ÀÀ¸ð¹øÈ£¼ö·® = m_¼¼Æ®ÀÀ¸ð¹øÈ£¼ö·® + 1
            
            Rs.MoveNext
        Loop
    End If
    Rs.Close
    Set Rs = Nothing
    
    SUBRs.Close
    Set SUBRs = Nothing
    
    Return
'-------------------------------------------------------------------------------
'--   Å¸ÀÌÆ² ºÎºÐ Ãâ·Â
'-------------------------------------------------------------------------------
Print_Title:
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '  ´ë¸®Á¡ º¸°ü¿ë

    For j = 0 To 1
        If j = 1 Then
            PrtPoint2 = GetPrtPointMM("¼Õ´Ô¿ë")
        Else
            PrtPoint2.x = 0
            PrtPoint2.y = 0
        End If
        
        PrtPoint4 = GetPrtPointMM("¿©¹é")                ' ¼³Á¤ÇÑ ¿©¹éÀ» °¡Áö°í ¿Â´Ù.
        
        
        If ´ë¸®Á¡Á¤º¸.MasterCode <> M_COUPON_KLENZ_CODE Then
            If Format(Date, "yyyyMMdd") >= "20091207" And Format(Date, "yyyyMMdd") <= "20091231" Then   '--
                
                PrtPoint.x = 0: PrtPoint.y = 0
                
                zz = GetPrtPointMM("¿©¹é")
                zz.y = zz.y - 7
                
                Printer.FontName = "±¼¸²Ã¼"
                Printer.Font.Bold = True
                Printer.Font.Size = 7
                
                SetPrtPoint PrtPoint, PrtPoint2, zz
                Printer.Print "¡Ú¡Ú ¼¼Æ®¼¼Å¹¼­ºñ½º Ãâ½Ã±â³ä ÀÌº¥Æ® 2009-12-11 ~ 12-31ÀÏ±îÁö ¡Ú¡Ú"
                
                zz.y = zz.y + 3
                SetPrtPoint PrtPoint, PrtPoint2, zz
                Printer.Print "1.¼¼Å¹¹° 10%ÇÒÀÎ 2.°æÇ°ÀÌº¥Æ® 3.¼¼Æ® ¼¼Å¹ Á¢¼ö½Ã ¹«·á ¼¼Å¹±Ç ÁõÁ¤"
            
                Printer.FontName = "±¼¸²Ã¼"
                Printer.Font.Bold = True
                Printer.Font.Size = 9
                
            ElseIf Format(Date, "yyyyMMdd") >= "20100101" Then   '--
                
                PrtPoint.x = 0: PrtPoint.y = 0
                
                zz = GetPrtPointMM("¿©¹é")
                zz.y = zz.y - 7
                
                Printer.FontName = "±¼¸²Ã¼"
                Printer.Font.Bold = True
                Printer.Font.Size = 7
                
                SetPrtPoint PrtPoint, PrtPoint2, zz
                Printer.Print "¡Ú¡Ú ¼¼Æ®¼¼Å¹¼­ºñ½º Ãâ½Ã¡Ú¡Ú"
                
                zz.y = zz.y + 3
                SetPrtPoint PrtPoint, PrtPoint2, zz
                Printer.Print "¼¼Æ®¼¼Å¹ Á¢¼ö½Ã 7 ~ 3% ÇÒÀÎ¼­ºñ½º Á¦°ø"
            
                Printer.FontName = "±¼¸²Ã¼"
                Printer.Font.Bold = True
                Printer.Font.Size = 9
            End If
        End If
        
        ' ÀüÇ¥ ¹øÈ£
        If Printer_BO_Gb = "0" Then
            PrtPoint = GetPrtPointMM("PRTNO")
            SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
            Printer.Print FPrtTop.PrtNo
        End If
        If Printer_BO_Gb = "1" Then
            PrtPoint = GetPrtPointMM("HPTEL")
            SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
            Printer.Print FPrtTop.HpTel
        End If
        
        ' °í°´ ÀüÈ­¹øÈ£
        PrtPoint = GetPrtPointMM("GTEL")
        SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
        Printer.Print FPrtTop.Tel
        ' °í°´ ¼º¸í
        PrtPoint = GetPrtPointMM("GNAME")
        SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
        Printer.Print FPrtTop.Name
        ' ÁÖ¼Ò (¼Õ´Ô)
        PrtPoint = GetPrtPointMM("ADDR")
        SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
        Printer.Print FPrtTop.Addr
        ' Á¢¼öÀÏ
        PrtPoint = GetPrtPointMM("DATE")
        SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
        Printer.Print FPrtTop.Date
        ' °í°´ ¹øÈ£
        PrtPoint = GetPrtPointMM("CODE")
        SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
        Printer.Print FPrtTop.Code
        ' ÀÎµµ ¿¬µµ
        PrtPoint = GetPrtPointMM("DATE2")
        SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
        Printer.Print FPrtTop.Date2
    Next j
    Return
'-------------------------------------------------------------------------------
'--   ³»¿ë ºÎºÐ Ãâ·Â
'-------------------------------------------------------------------------------
Print_Center:
    
    
    ll = 0 ' º¸°üÁõ Ãâ·Â ¶óÀÎ ÃÊ±âÈ­
    If (S_Line + Page_Item_Count) > Page_Count Then
        SUB_TOT = Page_Count
    Else
        SUB_TOT = S_Line + Page_Item_Count - 1
    End If
    
    ' ±âº» ¶óÀÎ´ç °£°ÝÀ» °¡Àú¿Â´Ù
    PrtPoint3 = GetPrtPoint("NEXT_LINE")
    PrtPoint4 = GetPrtPoint("¿©¹é")
    For i = S_Line To SUB_TOT
        ll = ll + 1
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' º¸°ü¿ë
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Ã¹ÁÙÀº Áõ°¡ ÇÏÁö ¾Ê´Â´Ù
        If (ll - 1) Then
            PrtPoint4.y = PrtPoint4.y + PrtPoint3.y + IIf((i Mod 2), 1, 0)
        End If
        
        For j = 0 To 1
            If j = 1 Then
                PrtPoint2 = GetPrtPointMM("¼Õ´Ô¿ë")
            Else
                PrtPoint2.x = 0
                PrtPoint2.y = 0
            End If
            
        
            'ÅÃ¹øÈ£
            PrtPoint = GetPrtPointMM("TAGNUM")
            SetPrtPoint PrtPoint2, PrtPoint, PrtPoint4
            SetPrtPoint PrtPoint2, PrtPoint, PrtPoint4
            Printer.Print FPArray(i, 1)
            
            'Ç°¸í
            PrtPoint = GetPrtPointMM("PNAME")
            SetPrtPoint PrtPoint2, PrtPoint, PrtPoint4
            Printer.Print FPArray(i, 2)
            
            '»ö»ó
            PrtPoint = GetPrtPointMM("PCOLOR")
            SetPrtPoint PrtPoint2, PrtPoint, PrtPoint4
            Printer.Print FPArray(i, 3)
            
            '±Ý¾×
            PrtPoint = GetPrtPointMM("PACCOUNT")
            SetPrtPoint PrtPoint2, PrtPoint, PrtPoint4
            Printer.Print FPArray(i, 4)
            
            '³»¿ë
            PrtPoint = GetPrtPointMM("PTEMP")
            SetPrtPoint PrtPoint2, PrtPoint, PrtPoint4
            Printer.Print FPArray(i, 5)
            
            '»óÇ¥
            PrtPoint = GetPrtPointMM("BRAND")
            SetPrtPoint PrtPoint2, PrtPoint, PrtPoint4
            Printer.Print FPArray(i, 6)
        Next j
    Next i
    
    Return
    
'-------------------------------------------------------------------------------
'--   ¼¼Æ® »óÇ° ºÎºÐÀ» Ãâ·Â ÇÑ´Ù.
'-------------------------------------------------------------------------------
Print_GropGoodsINFO:
    
    ' ¼¼Æ® ³»¿ªÀÌ ¾øÀ»°æ¿ì Ãâ·ÂÇÏÁö ¾Ê´Â´Ù.
    If m_GSGMoney.d¼¼Æ®¼ö·®ÇÕ°è <= 0 Then Return
    
    ' ±âº» ¶óÀÎ´ç °£°ÝÀ» °¡Àú¿Â´Ù
    PrtPoint3 = GetPrtPoint("NEXT_LINE")
    PrtPoint4 = GetPrtPoint("¿©¹é")
    
    If Format(Date, "yyyyMMdd") <= "20091231" Then
        PrtPoint4.y = 44
        
        For j = 0 To 1
            If j = 1 Then
                PrtPoint2 = GetPrtPointMM("¼Õ´Ô¿ë")
            Else
                PrtPoint2 = GetPrtPointMM("º¸°ü¿ë")
            End If
            
            'ÅÃ¹øÈ£
            SetPrtPoint PrtPoint2, GetPrtPointMM("TAGNUM"), PrtPoint4
            Printer.Print "°æÇ°ÃßÃ·Àº ´ç»ç È¨ÆäÀÌÁö " & Chr(34) & "°æÇ°ÀÌº¥Æ® Âü¿©ÇÏ±â" & Chr(34) & "¿¡"
        Next j
        
        PrtPoint4.y = 48
        For j = 0 To 1
            If j = 1 Then
                PrtPoint2 = GetPrtPointMM("¼Õ´Ô¿ë")
            Else
                PrtPoint2 = GetPrtPointMM("º¸°ü¿ë")
            End If
            
            SetPrtPoint PrtPoint2, GetPrtPointMM("TAGNUM"), PrtPoint4
            Printer.Print "ÀÀ¸ðÇÏ½Å °í°´ºÐ¿¡ ÇÑÇÏ¿© ÃßÃ·ÇÕ´Ï´Ù. 12¿ù 31ÀÏ±îÁö"
        Next j
        
        PrtPoint4.y = 53
        
        For j = 0 To 1
            If j = 1 Then
                PrtPoint2 = GetPrtPointMM("¼Õ´Ô¿ë")
            Else
                PrtPoint2 = GetPrtPointMM("º¸°ü¿ë")
            End If
            
            'ÅÃ¹øÈ£
            SetPrtPoint PrtPoint2, GetPrtPointMM("TAGNUM"), PrtPoint4
'            strTempStr = strMaxLng
'            RSet strTempStr = Format(m_GSGMoney.dÃÖÁ¾¼ö·É¾×, "#,#0")
            
            If m_¼¼Æ®ÀÀ¸ð¹øÈ£¼ö·® = 1 Then
                Printer.Print "°æÇ°ÀÀ¸ð¹øÈ£: " & m_¼¼Æ®ÀÀ¸ð¹øÈ£(0) & Space(15) & "ÁõÁ¤¸Å¼ö: " & Format(m_GSGMoney.d¹«·á¼¼Å¹±Ç¼ö·®, "@@") & " Àå"
            ElseIf m_¼¼Æ®ÀÀ¸ð¹øÈ£¼ö·® = 2 Then
                Printer.Print "°æÇ°ÀÀ¸ð¹øÈ£: " & m_¼¼Æ®ÀÀ¸ð¹øÈ£(0) & ", " & m_¼¼Æ®ÀÀ¸ð¹øÈ£(1) & Space(5) & "ÁõÁ¤¸Å¼ö: " & Format(m_GSGMoney.d¹«·á¼¼Å¹±Ç¼ö·®, "@@") & " Àå"
            End If
        Next j
    End If
    
    PrtPoint4.y = 53
    PrtPoint4.y = PrtPoint4.y + PrtPoint3.y
    
    For j = 0 To 1
        If j = 1 Then
            PrtPoint2 = GetPrtPointMM("¼Õ´Ô¿ë")
        Else
            PrtPoint2 = GetPrtPointMM("º¸°ü¿ë")
        End If
        
        'ÅÃ¹øÈ£
        SetPrtPoint PrtPoint2, GetPrtPointMM("TAGNUM"), PrtPoint4
        
        strTempStr = strMaxLng
        RSet strTempStr = Format(m_GSGMoney.dÀüÃ¼±Ý¾×, "#,#0")
        Printer.Print "¼¼Æ®ÇÒÀÎÀü±Ý¾×: " & strTempStr
        
        '±Ý¾×
        SetPrtPoint PrtPoint2, GetPrtPointMM("PACCOUNT"), PrtPoint4
        strTempStr = "123456789"
        RSet strTempStr = Format(m_GSGMoney.d¼¼Æ®ÇÒÀÎ±Ý¾×, "#,#0")
        Printer.Print "¼¼Æ®±âº»ÇÒÀÎ: " & strTempStr
    Next j
    
            
    PrtPoint4.y = PrtPoint4.y + PrtPoint3.y + 1
    For j = 0 To 1
        If j = 1 Then
            PrtPoint2 = GetPrtPointMM("¼Õ´Ô¿ë")
        Else
            PrtPoint2 = GetPrtPointMM("º¸°ü¿ë")
        End If
        'ÅÃ¹øÈ£
        SetPrtPoint PrtPoint2, GetPrtPointMM("TAGNUM"), PrtPoint4
        strTempStr = strMaxLng
        RSet strTempStr = Format(m_GSGMoney.dÀüÃ¼ÇÒÀÎ±Ý¾×, "#,#0")
        Printer.Print "¼¼Æ®ÇÒÀÎ  ±Ý¾×: " & strTempStr
        
        '±Ý¾×
        SetPrtPoint PrtPoint2, GetPrtPointMM("PACCOUNT"), PrtPoint4
        strTempStr = "123456789"
        RSet strTempStr = Format(m_GSGMoney.d¿¡´©¸®ÇÒÀÎ±Ý¾×, "#,#0")
        Printer.Print "¿¡´©¸®  ÇÒÀÎ: " & strTempStr
    Next j
            
    PrtPoint4.y = PrtPoint4.y + PrtPoint3.y
    For j = 0 To 1
        If j = 1 Then
            PrtPoint2 = GetPrtPointMM("¼Õ´Ô¿ë")
        Else
            PrtPoint2 = GetPrtPointMM("º¸°ü¿ë")
        End If
        
        'ÅÃ¹øÈ£
        SetPrtPoint PrtPoint2, GetPrtPointMM("TAGNUM"), PrtPoint4
        strTempStr = strMaxLng
        
        RSet strTempStr = Format(m_GSGMoney.dÃÖÁ¾¼ö·É¾×, "#,#0")
        
        Printer.Print "¼¼Æ®ÇÒÀÎÈÄ±Ý¾×: " & strTempStr
        
        '±Ý¾×
        SetPrtPoint PrtPoint2, GetPrtPointMM("PACCOUNT"), PrtPoint4
        
        strTempStr = "2:" & m_GSGMoney.d2¼¼Æ®¼ö·® & ",3:" & m_GSGMoney.d3¼¼Æ®¼ö·® & "," & _
                     "4:" & m_GSGMoney.d4¼¼Æ®¼ö·® & ",5:" & m_GSGMoney.d4¼¼Æ®¼ö·® & "," & _
                     "ºò:" & m_GSGMoney.d5¼¼Æ®¼ö·®
        Printer.Print "±¸¼º: " & strTempStr
        
        
'        strTempStr = "123456789"
'        RSet strTempStr = Format(m_GSGMoney.d¼¼Æ®±Ý¾×, "#,#0")
'        Printer.Print "¼¼Æ®Ç°¸ñ±Ý¾×: " & strTempStr
        
    Next j
    
    Return

'-------------------------------------------------------------------------------
'--   ³¡ ºÎºÐ Ãâ·Â
'-------------------------------------------------------------------------------
Print_Bottom:
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' º¸°ü¿ë
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    For j = 0 To 1
        If j = 1 Then
            PrtPoint2 = GetPrtPointMM("¼Õ´Ô¿ë")
        Else
            PrtPoint2.x = 0
            PrtPoint2.y = 0
        End If
        
        PrtPoint4 = GetPrtPointMM("¿©¹é")                ' ¼³Á¤ÇÑ ¿©¹éÀ» °¡Áö°í ¿Â´Ù.
        ' ¸¶Áö¸· ÀåÀÏ°æ¿ì ÀüÃ¼ ÇÕ°è¹× ±Ý¾× Ãâ·Â
        If L_Page = sPage_count Or sPage_count = 1 Then
            ' Á¡¼ö
            PrtPoint = GetPrtPointMM("SUM")
            SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
            Printer.Print FPrtBottom.Sum
            '±Ý¾×
            PrtPoint = GetPrtPointMM("ACCOUNT0")
            SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
            Printer.Print FPrtBottom.Account0
            ' ¼ö·É¾×
            PrtPoint = GetPrtPointMM("ACCOUNT1")
            SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
            Printer.Print FPrtBottom.Account1
            'ÀÜ¾×
            PrtPoint = GetPrtPointMM("ACCOUNT2")
            SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
            Printer.Print FPrtBottom.Account2
        
            '¸¶ÀÏ¸®Áö
            If Val(FPrtBottom.MilMoney) > 0 Then
                PrtPoint = GetPrtPointMM("MILEAGE")
                SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
                'Printer.Print "¸¶ÀÏ¸®ÁöÀÜ¾× : " & FPrtBottom.MilMoney
                Printer.Print FPrtBottom.MilMoney
            End If
            
            If Printer_BO_Gb = "1" Then
                ' ÀüÀÏ ¹Ì¼ö
                PrtPoint = GetPrtPointMM("OLDMISU")
                SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
                Printer.Print FPrtBottom.OldDayMisu
                ' ¹Ì¼ö ÇÕ°è
                PrtPoint = GetPrtPointMM("MISUMONEY")
                SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
                Printer.Print FPrtBottom.MiSuTotal
                ' ¼ö±Ý¾×
                PrtPoint = GetPrtPointMM("SUGUMONEY")
                SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
                Printer.Print FPrtBottom.SuGumMonye
                ' »ç¿ë¸¶ÀÏ¸®Áö
                PrtPoint = GetPrtPointMM("USERMILEAGE")
                SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
                Printer.Print FPrtBottom.MilUser
                
                ' ¸¶ÀÏ¸®Áö ÀÜ¾×
            If Val(FPrtBottom.MilMoney) > 0 Then
                    PrtPoint = GetPrtPointMM("MILEAGE")
                    SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
                    Printer.Print FPrtBottom.MilMoney
            End If
            
            If ´ë¸®Á¡Á¤º¸.¸¶ÀÏ¸®Áö¿©ºÎ = "Y" Then
                ' ´©Àû ¸¶ÀÏ¸®Áö
                PrtPoint = GetPrtPointMM("ADDMILEAGE")
                SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
                Printer.Print FPrtBottom.MilAddMoney
            End If
            
                ' º¸°üÁõ ¿À·ù ¼öÁ¤
                PrtPoint = GetPrtPointMM("ADDMILEAGETITLE")
                SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
                Printer.Print "Àû¸³"
            End If
        End If
        
        ' ´ë¸®Á¡¸í
        PrtPoint = GetPrtPointMM("DNAME")
        SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
        Printer.Print FPrtBottom.DName & "   ÄíÆù:" & FPrtBottom.CouponMoney
        ' ´ë¸®Á¡ ÀüÈ­¹øÈ£
        PrtPoint = GetPrtPointMM("DTEL")
        SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
        Printer.Print FPrtBottom.DTel
        ' ÆäÀÌÁö/ÀüÃ¼ ÆäÀÌÁö
        PrtPoint = GetPrtPointMM("PAGE")
        SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
        Printer.Print L_Page & "/" & sPage_count
        
    Next j
Return
'-------------------------------------------------------------------------------
'--   ÀÎ¼âÁß ¿À·ù ½ÇÇà ºÎºÐ
'-------------------------------------------------------------------------------
printError:
    MsgBox " ÇÁ¸°ÅÍ¸¦ È®ÀÎÇØ ÁÖ½Ê½Ã¿ä ! " & vbNewLine & vbNewLine & Err.Description, vbCritical, "Ãâ·Â¿À·ù¹ß»ý"
    Resume Next
End Function

Function GetPrtPoint(prtIndex As String) As PrintPoint
    Select Case UCase(prtIndex)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' »ó´Ü °ü·Ã À§Ä¡
        Case "PRTNO"            ' ÀüÇ¥ ¹øÈ£
            GetPrtPoint.x = 230
            GetPrtPoint.y = 320
        Case "GTEL"             ' °í°´ ÀüÈ­ ¹øÈ£
            GetPrtPoint.x = 230
            GetPrtPoint.y = 380
        Case "GNAME"            ' °í°´ ¼º¸í
            GetPrtPoint.x = 760
            GetPrtPoint.y = 380
        Case "ADDR"             ' ÁÖ¼Ò
            GetPrtPoint.x = 230
            GetPrtPoint.y = 450
        Case "DATE"             ' Á¢¼öÀÏ
            GetPrtPoint.x = 760
            GetPrtPoint.y = 450
        Case "CODE"             ' °í°´ ¹øÈ£
            GetPrtPoint.x = 230
            GetPrtPoint.y = 520
        Case "DATE2"            ' ÀÎµµ ÀÏ
            GetPrtPoint.x = 760
            GetPrtPoint.y = 520
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Áß°£ °ü·Ã À§Ä¡
        Case "NEXT_LINE"        ' Áß°£ ³»¿ëÀ» Ãâ·ÂÇÒ¶§  ´ÙÀ½ ¶óÀÎ°úÀÇ °Å¸®
            GetPrtPoint.x = 0
            GetPrtPoint.y = Prt_Height
        Case "TAGNUM"           ' ÅÃ ¹øÈ£
            GetPrtPoint.x = 30
            GetPrtPoint.y = 650
        Case "PNAME"            ' »óÇ°¸í
            GetPrtPoint.x = 150
            GetPrtPoint.y = 650
        Case "PCOLOR"           ' Ä®¶ó
            GetPrtPoint.x = 450
            GetPrtPoint.y = 650
        Case "PACCOUNT"         ' ±Ý¾×
            GetPrtPoint.x = 580
            GetPrtPoint.y = 650
        Case "PTEMP"            ' ³»¿ë
            GetPrtPoint.x = 760
            GetPrtPoint.y = 650
        Case "BRAND"            ' »óÇ¥
            GetPrtPoint.x = 860
            GetPrtPoint.y = 650
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ÇÏ´Ü °ü·Ã À§Ä¡
        Case "SUM"              ' ÇÕ°è °Ç¼ö
            GetPrtPoint.x = 400
            GetPrtPoint.y = 1165
        Case "ACCOUNT0"         ' ÇÕ°è ±Ý¾×
            GetPrtPoint.x = 800
            GetPrtPoint.y = 1165
        Case "ACCOUNT1"         ' ¼ö·É¾×
            GetPrtPoint.x = 800
            GetPrtPoint.y = 1225
        Case "ACCOUNT2"         ' ÀÜ¾×
            GetPrtPoint.x = 800
            GetPrtPoint.y = 1295
        Case "DNAME"            '´ë¸®Á¡ ¸í
            GetPrtPoint.x = 175
            GetPrtPoint.y = 1360
        Case "DTEL"             ' ´ë¸®Á¡ ÀüÈ­¹øÈ£
            GetPrtPoint.x = 170
            GetPrtPoint.y = 1420
        Case "PAGE"             ' Ãâ·Â ÆäÀÌÁö 1/2
            GetPrtPoint.x = 900
            GetPrtPoint.y = 1500
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ±âÅ¸ °ü·Ã À§Ä¡
        Case "¿©¹é"             ' Ãâ·ÂÇÒ ÆäÀÌÁöÀ§ À§ÂÊ ¿©¹é
            GetPrtPoint.x = Prt_Left
            GetPrtPoint.y = Prt_Top
        Case "¼Õ´Ô¿ë"           ' ¼Õ´Ô¿ë Ãâ·Â ½ÃÀÛ À§Ä¡
            GetPrtPoint.x = 1125
            GetPrtPoint.y = 0
        Case "º¸°ü¿ë"           ' º¸°ü¿ë Ãâ·Â ½ÃÀÛ À§Ä¡( ¹Ì»ç¿ë)
            GetPrtPoint.x = 0
            GetPrtPoint.y = 0
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' "Ãâ·Â ³»¿ë À§Ä¡ ¿À·ù
        Case Else               ' ±âÅ¸
            GetPrtPoint.x = 0
            GetPrtPoint.y = 0
            Debug.Print (UCase(prtIndex) & "Ãâ·Â À§Ä¡ ¿À·ù")
    End Select
End Function

Public Function GetPrtPointMM(prtIndex As String) As PrintPoint

    ' ÀÌÀü º¸°üÁõ Ãâ·Â
    If Printer_BO_Gb = "0" Then
        Select Case UCase(prtIndex)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' »ó´Ü °ü·Ã À§Ä¡
            Case "PRTNO"            ' ÀüÇ¥ ¹øÈ£
                GetPrtPointMM.x = 15
                GetPrtPointMM.y = 0
            Case "GTEL"             ' °í°´ ÀüÈ­ ¹øÈ£
                GetPrtPointMM.x = 15
                GetPrtPointMM.y = 5
            Case "GNAME"            ' °í°´ ¼º¸í
                GetPrtPointMM.x = 60
                GetPrtPointMM.y = 5
            Case "ADDR"             ' ÁÖ¼Ò
                GetPrtPointMM.x = 15
                GetPrtPointMM.y = 11
            Case "DATE"             ' Á¢¼öÀÏ
                GetPrtPointMM.x = 60
                GetPrtPointMM.y = 11
            Case "CODE"             ' °í°´ ¹øÈ£
                GetPrtPointMM.x = 15
                GetPrtPointMM.y = 17
            Case "DATE2"            ' ÀÎµµ ÀÏ
                GetPrtPointMM.x = 60
                GetPrtPointMM.y = 17
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Áß°£ °ü·Ã À§Ä¡
            Case "NEXT_LINE"        ' Áß°£ ³»¿ëÀ» Ãâ·ÂÇÒ¶§  ´ÙÀ½ ¶óÀÎ°úÀÇ °Å¸®
                GetPrtPointMM.x = 0
                GetPrtPointMM.y = Prt_Height
            Case "TAGNUM"           ' ÅÃ ¹øÈ£
                GetPrtPointMM.x = 0
                GetPrtPointMM.y = 27
            Case "PNAME"            ' »óÇ°¸í
                GetPrtPointMM.x = 10
                GetPrtPointMM.y = 27
            Case "PCOLOR"           ' Ä®¶ó
                GetPrtPointMM.x = 35
                GetPrtPointMM.y = 27
            Case "PACCOUNT"         ' ±Ý¾×
                GetPrtPointMM.x = 46
                GetPrtPointMM.y = 27
            Case "PTEMP"            ' ³»¿ë
                GetPrtPointMM.x = 60
                GetPrtPointMM.y = 27
            Case "BRAND"            ' »óÇ¥
                GetPrtPointMM.x = 69
                GetPrtPointMM.y = 27
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' ÇÏ´Ü °ü·Ã À§Ä¡
            Case "SUM"              ' ÇÕ°è °Ç¼ö
                GetPrtPointMM.x = 19
                GetPrtPointMM.y = 72
            Case "ACCOUNT0"         ' ÇÕ°è ±Ý¾×
                GetPrtPointMM.x = 60
                GetPrtPointMM.y = 72
            Case "ACCOUNT1"         ' ¼ö·É¾×
                GetPrtPointMM.x = 52
                GetPrtPointMM.y = 78
            Case "ACCOUNT2"         ' ÀÜ¾×
                GetPrtPointMM.x = 60
                GetPrtPointMM.y = 84
            Case "DNAME"            '´ë¸®Á¡ ¸í
                GetPrtPointMM.x = 14
                GetPrtPointMM.y = 89
            Case "DTEL"             ' ´ë¸®Á¡ ÀüÈ­¹øÈ£
                GetPrtPointMM.x = 14
                GetPrtPointMM.y = 93
            Case "PAGE"             ' Ãâ·Â ÆäÀÌÁö 1/2
                GetPrtPointMM.x = 75
                GetPrtPointMM.y = 99
            Case "MILEAGE"          ' ¸¶ÀÏ¸®Áö ÀÜ¾×
                GetPrtPointMM.x = 60
                GetPrtPointMM.y = 85
            
            
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' ±âÅ¸ °ü·Ã À§Ä¡
            Case "¿©¹é"             ' Ãâ·ÂÇÒ ÆäÀÌÁöÀ§ À§ÂÊ ¿©¹é
                GetPrtPointMM.x = Prt_Left
                GetPrtPointMM.y = Prt_Top
            Case "¼Õ´Ô¿ë"           ' ¼Õ´Ô¿ë Ãâ·Â ½ÃÀÛ À§Ä¡
                GetPrtPointMM.x = 95
                GetPrtPointMM.y = 0
            Case "º¸°ü¿ë"           ' º¸°ü¿ë Ãâ·Â ½ÃÀÛ À§Ä¡
                GetPrtPointMM.x = 0
                GetPrtPointMM.y = 0
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' "Ãâ·Â ³»¿ë À§Ä¡ ¿À·ù
            Case Else               ' ±âÅ¸
                GetPrtPointMM.x = 0
                GetPrtPointMM.y = 0
                Debug.Print (UCase(prtIndex) & "Ãâ·Â À§Ä¡ ¿À·ù")
        End Select
    
    ElseIf Printer_BO_Gb = "1" Then
    
        Select Case UCase(prtIndex)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' »ó´Ü °ü·Ã À§Ä¡
            Case "CODE"             ' °í°´ ¹øÈ£
                GetPrtPointMM.x = 16
                GetPrtPointMM.y = 0
            Case "GTEL"             ' °í°´ ÀüÈ­ ¹øÈ£
                GetPrtPointMM.x = 16
                GetPrtPointMM.y = 5
            Case "GNAME"            ' °í°´ ¼º¸í
                GetPrtPointMM.x = 60
                GetPrtPointMM.y = 5
            Case "HPTEL"            ' ÈÞ´ëÆù ¹øÈ£
                GetPrtPointMM.x = 16
                GetPrtPointMM.y = 10
            Case "DATE"             ' Á¢¼öÀÏ
                GetPrtPointMM.x = 60
                GetPrtPointMM.y = 10
            Case "ADDR"             ' ÁÖ¼Ò
                GetPrtPointMM.x = 16
                GetPrtPointMM.y = 15
            Case "DATE2"            ' ÀÎµµ ÀÏ
                GetPrtPointMM.x = 60
                GetPrtPointMM.y = 15
            Case "PRTNO"            ' ÀüÇ¥ ¹øÈ£
                GetPrtPointMM.x = 16
                GetPrtPointMM.y = 0
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Áß°£ °ü·Ã À§Ä¡
            Case "NEXT_LINE"        ' Áß°£ ³»¿ëÀ» Ãâ·ÂÇÒ¶§  ´ÙÀ½ ¶óÀÎ°úÀÇ °Å¸®
                GetPrtPointMM.x = 0
                GetPrtPointMM.y = Prt_Height
            Case "TAGNUM"           ' ÅÃ ¹øÈ£
                GetPrtPointMM.x = 0
                GetPrtPointMM.y = 25
            Case "PNAME"            ' »óÇ°¸í
                GetPrtPointMM.x = 11
                GetPrtPointMM.y = 25
            Case "PCOLOR"           ' Ä®¶ó
                GetPrtPointMM.x = 33
                GetPrtPointMM.y = 25
            Case "PACCOUNT"         ' ±Ý¾×
                GetPrtPointMM.x = 46
                GetPrtPointMM.y = 25
            Case "PTEMP"            ' ³»¿ë
                GetPrtPointMM.x = 58
                GetPrtPointMM.y = 25
            Case "BRAND"            ' »óÇ¥
                GetPrtPointMM.x = 69
                GetPrtPointMM.y = 25
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' ÇÏ´Ü °ü·Ã À§Ä¡
            Case "SUM"              ' ÇÕ°è °Ç¼ö
                GetPrtPointMM.x = 19
                GetPrtPointMM.y = 75
            Case "ACCOUNT0"         ' ÇÕ°è ±Ý¾×
                GetPrtPointMM.x = 60
                GetPrtPointMM.y = 75
            Case "OLDMISU"          ' ÀüÀÏ ¹Ì¼ö
                GetPrtPointMM.x = 19
                GetPrtPointMM.y = 80
            Case "MISUMONEY"          ' ¹Ì¼ö ÇÕ°è
                GetPrtPointMM.x = 19
                GetPrtPointMM.y = 85
            Case "SUGUMONEY"          ' ¼ö±Ý¾×
                GetPrtPointMM.x = 19
                GetPrtPointMM.y = 95
            Case "ACCOUNT1"         ' ¼ö·É¾×
                GetPrtPointMM.x = 52
                GetPrtPointMM.y = 90
            Case "ACCOUNT2"         ' ÀÜ¾×
                GetPrtPointMM.x = 60
                GetPrtPointMM.y = 95
            Case "DNAME"            '´ë¸®Á¡ ¸í
                GetPrtPointMM.x = 19
                GetPrtPointMM.y = 100
            Case "DTEL"             ' ´ë¸®Á¡ ÀüÈ­¹øÈ£
                GetPrtPointMM.x = 19
                GetPrtPointMM.y = 105
            Case "PAGE"             ' Ãâ·Â ÆäÀÌÁö 1/2
                GetPrtPointMM.x = 50
                GetPrtPointMM.y = 105
            Case "USERMILEAGE"         ' »ç¿ë ¸¶ÀÏ¸®Áö
                GetPrtPointMM.x = 60
                GetPrtPointMM.y = 80
            Case "MILEAGE"             ' ¸¶ÀÏ¸®Áö ÀÜ¾×
                GetPrtPointMM.x = 60
                GetPrtPointMM.y = 85
            Case "ADDMILEAGE"          ' ´©Àû ¸¶ÀÏ¸®Áö
                GetPrtPointMM.x = 19
                GetPrtPointMM.y = 90
            Case "ADDMILEAGETITLE"     ' ´©Àû ¸¶ÀÏ¸®Áö
                GetPrtPointMM.x = 0
                GetPrtPointMM.y = 90
            
            
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' ±âÅ¸ °ü·Ã À§Ä¡
            Case "¿©¹é"             ' Ãâ·ÂÇÒ ÆäÀÌÁöÀ§ À§ÂÊ ¿©¹é
                GetPrtPointMM.x = Prt_Left
                GetPrtPointMM.y = Prt_Top
            Case "¼Õ´Ô¿ë"           ' ¼Õ´Ô¿ë Ãâ·Â ½ÃÀÛ À§Ä¡
                GetPrtPointMM.x = 95
                GetPrtPointMM.y = 0
            Case "º¸°ü¿ë"           ' º¸°ü¿ë Ãâ·Â ½ÃÀÛ À§Ä¡
                GetPrtPointMM.x = 0
                GetPrtPointMM.y = 0
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' "Ãâ·Â ³»¿ë À§Ä¡ ¿À·ù
            Case Else               ' ±âÅ¸
                GetPrtPointMM.x = 0
                GetPrtPointMM.y = 0
                Debug.Print (UCase(prtIndex) & "Ãâ·Â À§Ä¡ ¿À·ù")
        End Select
        
    
    
    End If
    
End Function

Function PrintPointDisplay()
    Dim px As Integer
    Dim py As Integer

'   Printer.PaperSize = vbPRPSA4
    Printer.Width = 19 * 567
    Printer.Height = 15 * 567
    Printer.ScaleMode = vbMillimeters           ' ÇÁ¸°ÅÍÀÇ ½ºÄÉÀÏ ¸ðµå¸¦ ¹Ð¸®¹ÌÅÍ ´ÜÀ§·Î
    Printer.FontName = "±¼¸²Ã¼"
    Printer.Font.Bold = True
    Printer.Font.Size = 9
    
'    MsgBox Str(Printer.ScaleWidth) & "," & Str(Printer.ScaleHeight)
'    Exit Function

    For px = 0 To Int(Printer.ScaleWidth) - 5 Step 5
        For py = 0 To Int(Printer.ScaleHeight) - 5 Step 5
            Printer.CurrentX = px
            Printer.CurrentY = py
            If (px Mod 50) = 0 And (py Mod 50) = 0 Then
                Printer.Print Str(px) & "." & Str(py)
            Else
                Printer.Print "."
            End If
        Next py
    Next px
    
    Printer.EndDoc
    
    Exit Function

prt_err:
    MsgBox " ÇÁ¸°ÅÍ¸¦ È®ÀÎÇØ ÁÖ½Ê½Ã¿ä ! " & VBA.Err.Number, vbCritical, "Ãâ·Â¿À·ù¹ß»ý"
    
    
End Function

Function subDayListPrint(cdPrt As CommonDialog, prtDay As String, bView As Boolean)
    Dim i As Long
    Dim kk As Long
    Dim FHandle As Integer                  ' ÀÎ¼âÇÒ ÆÄÀÏÀÇ ÇÚµé
    Dim ProssCount As Integer               ' ÀüÃ¼ ÆäÀÌÁö¿¡¼­ Ãâ·ÂµÉ ÃÑ ¾ÆÀÌÅÛ ÃÑ °¹¼ö
    Dim Prt_Total_Page_count As Integer     ' Ãâ·ÂµÉ ÀüÃ¼ ÆäÀÌÁö¼ö
    Dim PRINT_LINE_COUNT As Integer         ' ÇÑÆäÀÌÁö´ç Ãâ·ÂµÉ ¾ÆÀÌÅÛ °¹¼ö
    Dim PageCnt As Integer                  ' ÇöÀç Ãâ·ÂÁßÀÎ ÆäÀÌÁö
    Dim LineCnt As Integer                  ' ÇöÀç Ãâ·ÂÁßÀÎ ¶óÀÎ
    Dim strFileName As String               ' Ãâ·Â ÆÄÀÏ¸í
    Dim TextData(20) As String              ' ÀÎ¼âÇÒ ³»¿ëÀ» ÀÓ½Ã ÀúÀåÇÑ´Ù.
    Dim hhh(60) As String                   ' ¾ç½ÄÀ» ÀúÀåÇÑ´Ù.
    Dim dblReturnMoney  As Double
    Dim dblQNPrice(4)   As Double
    Dim dblMilPrice(4)   As Double
    Dim dblCardMoney    As Double       ' Ä«µå±Ý¾×
    Dim dblCardCount    As Double       ' Ä«µå°Ç¼ö
    Dim dblCouponCnt    As Double       ' ÄíÆù°Ç¼ö
    Dim strCouponNumber As String       ' ÄíÆù¹øÈ£
    Dim dblSaleReturnCnt    As Double       ' ¼¼Å¹È¯ºÒ°Ç¼ö
    Dim dblSaleReturnMoney  As Double       ' ¼¼Å¹È¯ºÒ±Ý¾×
    
    
    ' µ¿ÀÏ ÅÃ¹øÈ£,ÀüÈ­¹øÈ£,ÀÌ¸§ÀÌ Ãâ·Â µÇ´Â°ÍÀ» ¸·±â À§ÇÑº¯¼ö
    Dim tempTag As String                   ' µ¿ÀÏ ÅÃ¹øÈ£ Ãâ·ÂÀ» ¸·±â À§ÇÑ º¯¼ö
    Dim tempPhone As String                 ' µ¿ÀÏ ÀüÈ­¹øÈ£ Ãâ·ÂÀ» ¸·±â À§ÇÑ º¯¼ö
    Dim tempName As String                  ' µ¿ÀÏ ÀÌ¸§ Ãâ·ÂÀ» ¸·±â À§ÇÑ º¯¼ö
    
    ' ¸¶Áö¸· ÀüÃ¼ ÇÕ°è Ãâ·ÂÀ» À§ÇÑ º¯¼ö
    Dim iTemp As Integer
    Dim iTotal As Integer                   ' ÃÑÁ¡¼ö
    Dim iSub1 As Integer                    ' Àç¼¼Å¹
    Dim iSub2 As Integer                    ' ¹ÝÇ°
    Dim iSub3 As Integer                    ' ¼ö¼±
    Dim iSub4 As Integer                    ' »ç°íÇ°
    Dim iSub5 As Integer                    ' Àç´Ù¸²Áú
    Dim iTotalMoney As Long                 ' ¸ÅÃâ¾×
    Dim iSub1Money As Long                  ' º»»ç
    Dim iSub2Money As Long                  ' ´ë¸®Á¡
    Dim iSub3Money As Long                  ' ¼ö¼±ºñ¿ë
    Dim iSub4Money As Long                  ' ¿îµ¿È­ ±Ý¾×
    
    Dim dblRatio As Double                  ' º»»ç ¸¶Áø
    
    Dim rsReprint As DAO.Recordset
    Dim rsRep As DAO.Recordset
    Dim strMsg As String
    
    Dim nCouponTotal        As Integer          ' ÄíÆù ¼ö·®
    Dim nCouponTotalMoney   As Double           ' ÄíÆù ±Ý¾×
    Dim nCouponTotalMoney2  As Double           ' ÄíÆù °è»êÀ» À§ÇÑ ±Ý¾×
    Screen.MousePointer = 13
    
    On Error GoTo Error_Rtn
    
    prtDay = Replace(prtDay, "-", "")
    
    If Not IsDate(Format(prtDay, "0000-00-00")) Then
        MsgBox "ÀÏÀÚ°¡ Àß¸ø ÀÔ·Â µÇ¾ú½À´Ï´Ù. È®ÀÎ ÇÏ½Ê½Ã¿ä.", vbInformation, "Error"
        Screen.MousePointer = 1
        Exit Function
    End If
    
    If Not DayCloseCheck(prtDay) Then
        MsgBox " ÀÏÀÏ¸ÅÃâ¸¶°¨À» ÇÏ½ÅÈÄ¿¡ Ãâ·ÂÇÏ¼¼¿ä..! ", vbInformation, "ÀÏÀÏ¸ÅÃâ¸¶°¨"
        Screen.MousePointer = 1
        Exit Function
    End If
    
    ' Ãâ·Â ÀÚ·áÀÇ ¾çÀ» ±¸ÇÑ´Ù., ÄÚµå DB¸¦ ¿ÀÇÂ ÇÑ´Ù.
    ' Ãâ·Â ³»¿ëÀÌ ¾øÀ» °æ¿ì ¹Ù·Î Á¾·á ÇÑ´Ù.
    GoSub Print_ProssCount
    ' Ãâ·Â ¾ç½ÄÀ» ÃÊ±âÈ­ ÇÑ´Ù
    GoSub Print_Init
    ' Ãâ·ÂÇÒ ÆÄÀÏÀ» ¿ÀÇÂ ÇÑ´Ù.
    GoSub Print_FileOpen
    ' ¾ç½ÄÀÇ Å¸ÀÌÆ²À» È­ÀÏ¿¡ Ãâ·Â ÇÑ´Ù.
    GoSub Print_Head
    ' ¹Ýº¹µÇ´Â ºÎºÐÀ» È­ÀÏ¿¡ Ãâ·Â ÇÑ´Ù.
    GoSub Print_Next
    ' ¸¶Áö¸· ºÎºÐÀ» È­ÀÏ¿¡ Ãâ·Â ÇÑ´Ù.
    ' Ãâ·ÂÇÑ ÆÄÀÏÀ» ´Ý´Â´Ù.
    GoSub Print_Bottom
    
    Screen.MousePointer = 0
    ' È­ÀÏÀ» ÇÁ¸°ÅÍ·Î Ãâ·ÂÇÑ´Ù.
    Call FileToPrint(strFileName, 1, bView)
    Exit Function
    
'Ãâ·ÂÇÒ ÀÚ·áÀÇ Ä«¿îÅÍ¸¦ È®ÀÎÇÑ´Ù.
Print_ProssCount:
    '--------------------------------------------------------------
    '
    '--------------------------------------------------------------
    Query = "SELECT * FROM ÀÔÃâ°í "
    Query = Query & " WHERE ÀÔ°íÀÏ='" & prtDay & "' "
    Query = Query & "   AND ( ÆÇ¸ÅÃë¼Ò IS NULL OR ÆÇ¸ÅÃë¼Ò <> 'Y') "
    Query = Query & " ORDER BY ¹øÈ£"
    Set Rs = New ADODB.Recordset
    Rs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If Rs.RecordCount > 0 Then
        Rs.MoveLast
        ProssCount = Rs.RecordCount
        Rs.MoveFirst
    Else
        Rs.Close
        MsgBox "ÆÇ¸ÅµÈ ³»¿ëÀÌ ¾ø½À´Ï´Ù.", vbInformation, "È®ÀÎ"
        Screen.MousePointer = 0
        Exit Function
    End If
    Rs.Close
    Set Rs = Nothing
    
    Return


' Ãâ·ÂÇÒ ¾ç½ÄÀ» ÃÊ±âÈ­ ÇÑ´Ù.
Print_Init:
    ' Ãâ·Â ÆÄÀÏ¸í
    strFileName = App.Path & "\PrintRep.txt"
     
    hhh$(1) = "                    @@@@@@@@@@@@@@ ÀÏÀÏ¸ÅÃâÇöÈ²                                                                  "
    hhh$(2) = "                  ¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬                                                               "
    hhh$(3) = "  ÀÏ ÀÚ : !!!!/!!/!!  @@@@@@                                                                                      "
    hhh$(4) = " ¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬ "
    hhh$(5) = "  ÅÃ¹øÈ£  ÀüÈ­¹øÈ£  ¼º   ¸í      Ç°        ¸í   ±Ý   ¾×  »ö»ó  ³»  ¿ë   »ó              Ç¥     »óÅÂ ÀÎ¼öÀÚ ÀÎ¼öÀÏ "
    hhh$(6) = " ¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬ "
    hhh$(7) = "  @@@@@  @@@@@@@@@  @@@@@@@@@@ @@@@@@@@@@@@@@@  ###,###  @@@@  @@@@@@  @@@@@@@@@@@@@@@@@@@@@@@ @@@@@              "
    hhh$(8) = "                                                                                                                  "
    hhh$(9) = "                                                                                           Page  : ### / ###      "
    hhh$(10) = ""
    hhh$(11) = " ¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬ "
    hhh$(12) = " ¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬ "
    hhh$(13) = "                                                                                                                  "
    hhh$(14) = "  ÃÑÁ¡¼ö:  ##,### Á¡      Àç¼¼Å¹:  ### Á¡     ¹ÝÇ°:  ### Á¡   ¼ö¼±:  ### Á¡   »ç°íÇ°:  ### Á¡    Àç´Ù¸²Áú: ### Á¡ "
    hhh$(15) = "                                                                                                                  "
    hhh$(16) = "  ¸ÅÃâ¾×:  #,###,### ¿ø        º»  »ç:  #,###,### ¿ø        ´ë¸®Á¡:  #,###,### ¿ø        ¼ö¼±ºñ¿ë:   #,###,### ¿ø "
    hhh$(32) = "  ¼ö±Ý¾×:  #,###,### ¿ø        ¿Ï  ºÒ:  #,###,### ¿ø        ¹Ì¼ö±Ý:  #,###,### ¿ø        ¹ÝÇ°È¯ºÒ:   #,###,### ¿ø "
    hhh$(40) = "  Ä«µå±Ý¾×:#,###,### ¿ø      Ä«µå°Ç¼ö:  #,###,### °Ç                             ºÒ·®¼¼Å¹È¯ºÒ±Ý¾×:   #,###,### ¿ø "
    hhh$(37) = "                                                                                                                  "
    hhh$(38) = "  ¹ÝÇ°È¯ºÒÁö»çÃ»±¸±Ý¾×:   #,###,### ¿ø        ¼¼Å¹È¯ºÒÁö»çÃ»±¸±Ý¾×:  #,###,### ¿ø                                 "
    hhh$(17) = "                                                                                                                  "
    hhh$(18) = " ¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬ "
    hhh$(19) = "  @@@@@  @@@@@@@@@  @@@@@@@@@@ @@@@@@@@@@@@@@@        @  @@@@  @@@@@@  @@@@@@@@@@@@@@@@@@@@@@@ @@@@@              "
    hhh$(20) = "                                                                                                                  "
    hhh$(21) = "  ´© ¶ô ÅÃ: @@ Á¡ ( @@@@@ - @@@@@ ) ÅÃ¹øÈ£: @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ "
    hhh$(22) = "                                            @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ "
    hhh$(23) = "                                                                                                                  "
    hhh$(24) = "  Àç»ç¿ëÅÃ: @@ Á¡                   ÅÃ¹øÈ£: @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ "
    hhh$(25) = "                                            @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ "
    hhh$(26) = "                                                                                                                  "
    hhh$(33) = "  ¹Ý Ç° ÅÃ: @@ Á¡                   ÅÃ¹øÈ£: @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ "
    hhh$(34) = "                                            @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ "
    hhh$(35) = "                                                                                                                  "
    hhh$(27) = " ¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬ "
    hhh$(28) = "  ¹ß»ý¸¶ÀÏ¸®Áö:  ###,##0 ¿ø      »ç¿ë¸¶ÀÏ¸®Áö:    ###,##0 ¿ø           »èÁ¦¸¶ÀÏ¸®Áö:    ###,###,##0 ¿ø            "
    hhh$(29) = "                                                                                                                  "
    hhh$(30) = "  ÀÔ ±Ý ÃÑ ¾×: #,###,##0 ¿ø         °øÀå ¸¶Áø:  #,###,##0 ¿ø            ´ë¸®Á¡ ¸¶Áø:    ###,###,##0 ¿ø            "
    hhh$(31) = " ¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬ "
    hhh$(39) = "  º¸ °ü ±Ý ¾×: #,###,##0 ¿ø      º»  »ç :  #,###,##0 ¿ø      Ã¼ÀÎÁ¡ :  #,###,##0 ¿ø     º¸°ü ¾÷Ã¼:  #,###,##0 ¿ø  "
    hhh$(41) = "  ÄíÆù°Ç¼ö:#,###,### °Ç      ÄíÆù¹øÈ£:  @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ " '73
    hhh$(42) = "                                        @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ "
    hhh$(43) = "  Äí Æù ´Ü °¡: #,###,##0 ¿ø      »ç¿ëÄíÆù±Ý¾×:    ###,##0 ¿ø      ÄíÆùÀû¿ëÀü ¸ÅÃâ¾×:    ###,###,##0 ¿ø            "
    hhh$(44) = "                                     ¸¶ÀÏ¸®Áö:    ###,##0 ¿ø         ÃÖÁ¾ ½Ç ÀÔ±Ý¾×:    ###,###,##0 ¿ø            "
    hhh$(45) = "  ºÒ·®¼¼Å¹È¯ºÒ°Ç¼ö: #,### °Ç   ºÒ·®¼¼Å¹³»¿ë :  @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ " '73
    hhh$(46) = "                                               @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ "
    hhh$(47) = "  »ï¼ºÄ«µå ÇÒÀÎ Á¤º¸ [  °í°´¼ö : #,### ¸í   »óÇ°¼ö : #,### Á¡  ÇÒÀÎ±Ý¾×   ###,##0 ¿ø  ]                           "
    hhh$(48) = "  ¼¼Æ®»óÇ° Á¤º¸    W2Á¾¼¼Æ® #,##0 °Ç,  W3Á¾¼¼Æ® #,##0 °Ç, W4Á¾¼¼Æ® #,##0 °Ç, W5Á¾¼¼Æ® #,##0 °Ç, W6Á¾¼¼Æ® #,##0 °Ç "
    hhh$(49) = "                   ¹«·á¼¼Å¹±Ç¼ö: #,##0 Àå   °í°´¼ö: #,##0 ¸í  ¼¼Æ®ÇÒÀÎ±Ý¾×: ###,##0 ¿ø ¿¡´©¸®ÇÒÀÎ±Ý¾×: ###,##0 ¿ø "
    
    ' ÆäÀÌÁö ¹× ¶óÀÎÀ» ÃÊ±âÈ­ ÇÑ´Ù.
    PageCnt = 0:  LineCnt = 0
     ' ÇÑÆäÀÌÁö´ç Ãâ·ÂµÉ ¾ÆÀÌÅÛ °¹¼ö
    PRINT_LINE_COUNT = GetPrtItemCount("ÀÏÀÏ¸ÅÃâÇöÈ²")
    
    ' ÀüÃ¼ Ãâ·Â ÆäÀÌÁö ±¸ÇÏ±â
    If (ProssCount Mod PRINT_LINE_COUNT) <> 0 Then
        Prt_Total_Page_count = Int(ProssCount / PRINT_LINE_COUNT) + 1
    Else
        Prt_Total_Page_count = Int(ProssCount / PRINT_LINE_COUNT)
    End If
    
    ' »ç¿ëº¯¼ö ÃÊ±âÈ­.....
    tempTag = "":       tempPhone = "":     tempName = ""
    iTotal = 0:         iSub1 = 0:          iSub2 = 0:          iSub3 = 0:      iSub4 = 0:  iSub5 = 0
    iTotalMoney = 0:    iSub1Money = 0:     iSub2Money = 0:     iSub3Money = 0

    '---------------------------------------------
    ' ¸¶Áø ±âÁØÀº ´ë¸®Á¡, º»»ç´Â 1-´ë¸®Á¡ºñÀ²
    '---------------------------------------------
    Query = "SELECT * FROM ´ë¸®Á¡Á¤º¸ "
    Set Rs = New ADODB.Recordset
    Rs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If Not IsNull(Rs!ºñÀ²) Then
        dblRatio = (CDbl(Rs!ºñÀ²)) / 100
        dblRatio = 1 - dblRatio
    End If
    Rs.Close
    Set Rs = Nothing
    
    Return
    
    
' Ãâ·ÂÇÒ ÆÄÀÏÀ» ¿ÀÇÂ ÇÑ´Ù.
Print_FileOpen:
    FHandle = FreeFile
    Open strFileName For Output As #FHandle
    Return
  
'¹Ýº¹µÇ´Â Å¸ÀÌÆ²À» Ãâ·Â ÇÑ´Ù.
Print_Head:
    PageCnt = PageCnt + 1
    LineCnt = 0
    Print #FHandle, hhh$(8) '¿©¹é Ãâ·Â
    Print #FHandle, hhh$(8) '¿©¹é Ãâ·Â
    Print #FHandle, hhh$(8) '¿©¹é Ãâ·Â
    Print #FHandle, hhh$(8) '¿©¹é Ãâ·Â
    Print #FHandle, hhh$(8) '¿©¹é Ãâ·Â
    Print #FHandle, hhh$(8) '¿©¹é Ãâ·Â
    TextData$(1) = ´ë¸®Á¡Á¤º¸.´ë¸®Á¡¸í
    Print #FHandle, Line_Format(hhh$(1), TextData(), 1)
    Print #FHandle, hhh$(2)
    Print #FHandle, hhh$(8)
    'ÀÏÀÚ/ ÆäÀÌÁö Ãâ·Â
    TextData$(1) = prtDay
    TextData$(2) = WeekdayName(Weekday(Format(prtDay, "0000-00-00")), False)
    Print #FHandle, Line_Format(hhh$(3), TextData(), 2)
    Print #FHandle, hhh$(4)
    Print #FHandle, hhh$(5)
    Print #FHandle, hhh$(6)
    Return
    
' Áß°£ÀÇ ¹Ýº¹ ºÎºÐÀÇ ÀÚ·á¸¦ Ãâ·Â ÇÑ´Ù.
Print_Next:
    ' ´ÙÀ½¿¡ Ãâ·ÂÇÒ ÀÚ·á°¡ ¾øÀ»¶§ ±îÁú
    '--------------------------------------------------------------
    '
    '--------------------------------------------------------------
    Query = "SELECT * FROM ÀÔÃâ°í "
    Query = Query & " WHERE ÀÔ°íÀÏ='" & prtDay & "' "
    Query = Query & "   AND ( ÆÇ¸ÅÃë¼Ò IS NULL OR ÆÇ¸ÅÃë¼Ò <> 'Y') "
    Query = Query & " ORDER BY ¹øÈ£"
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
        
    If SUBRs.RecordCount <= 0 Then
        SUBRs.Close
        Set SUBRs = Nothing
        
        MsgBox "ÆÇ¸ÅµÈ ³»¿ëÀÌ ¾ø½À´Ï´Ù.", vbInformation, "È®ÀÎ"
        
        Exit Function
    End If
    
    Do Until SUBRs.EOF
        If Fb°í°´Á¤º¸(SUBRs!°í°´¹øÈ£) = "Error" Then
            MsgBox "°í°´Á¤º¸ ¿À·ù ÀÔ´Ï´Ù ", vbInformation, "Error"
            GoSub Error_Rtn
        End If
    
        TextData$(1) = Space(1):    TextData$(2) = Space(1):    TextData$(3) = Space(1):
        TextData$(4) = Space(1):    TextData$(5) = Space(1):    TextData$(6) = Space(1):
        TextData$(7) = Space(1):    TextData$(8) = Space(1):    TextData$(9) = Space(1):
        
        If tempTag <> SUBRs!¹øÈ£ Then
            TextData$(1) = SUBRs!¹øÈ£
        End If
        
        If tempPhone <> °í°´Á¤º¸.ÀüÈ­¹øÈ£ Then
            ' ÀüÈ­¹øÈ£°¡ ¹Ù²î¸é ´ÙÀ½¿¡ µ¿¸í ÀÌÀÎÀÌ ³ª¿Ã¼öµµ ÀÖ±â ¶§¹®¿¡ ÀÌ¸§À» ÃÊ±âÈ­ ÇÑ´Ù
            TextData$(2) = °í°´Á¤º¸.ÀüÈ­¹øÈ£
            tempName = ""
        End If
        
        If tempName <> Hangul_Mid(°í°´Á¤º¸.¼º¸í & Space(10), 1, 10) Then
            TextData$(3) = Hangul_Mid(°í°´Á¤º¸.¼º¸í & Space(10), 1, 10)
        End If
        
        tempTag = SUBRs!¹øÈ£
        tempPhone = °í°´Á¤º¸.ÀüÈ­¹øÈ£
        tempName = Hangul_Mid(°í°´Á¤º¸.¼º¸í & Space(10), 1, 10)
        TextData$(4) = Hangul_Mid(SUBRs!Ç°¸í & Space(10), 1, 20)
        ' ÀüÃ¼ ±Ý¾×À» ±¸ÇÑ´Ù.
        iTotalMoney = iTotalMoney + SUBRs!±Ý¾×
        
        '¿îµ¿È­ ±Ý¾×À» ±¸ÇÑ´Ù.
        If UCase(SUBRs!ÄÚµå) >= "A00" And UCase(SUBRs!ÄÚµå) <= "A99" Then
            iSub4Money = iSub4Money + SUBRs!±Ý¾×
        End If
        
        TextData$(5) = SUBRs!±Ý¾×
        TextData$(6) = Hangul_Mid(SUBRs!»ö»ó & Space(10), 1, 8)
        
        ' Àç¼¼Å¹,¹ÝÇ°,¼ö¼± ¼ö·® ±¸ÇÏ±â
        If InStr(SUBRs!³»¿ë, "Àç") > 0 Then iSub1 = iSub1 + 1
        If InStr(SUBRs!³»¿ë, "¹Ý") > 0 Then iSub2 = iSub2 + 1
        If InStr(SUBRs!³»¿ë, "¼ö") > 0 Then iSub3 = iSub3 + 1
        If InStr(SUBRs!³»¿ë, "µå»ç") > 0 Then iSub4 = iSub4 + 1
        'If InStr(SUBRs!³»¿ë, "µåÀç") > 0 Then iSub5 = iSub5 + 1
        
        If InStr(SUBRs!³»¿ë, "¼ö") > 0 Then
            iSub3Money = iSub3Money + Val(SUBRs!±Ý¾×)
        End If
        
        TextData$(7) = Hangul_Mid(SUBRs!³»¿ë & Space(10), 1, 8)
        TextData$(8) = Hangul_Mid(SUBRs!»óÇ¥ & Space(10), 1, 22)
        TextData$(9) = "¿ÏºÒ"
        If SUBRs!»óÅÂ = "Ú±" Then TextData$(9) = "¹ÌºÒ"
        
        ' ±Ý¾×ÀÌ 0¿øÀÏ °æ¿ì ¹®ÀÚ·Î Ã³¸®ÇÏ¿© 0À» Ãâ·Â ½ÃÅ²´Ù.
        If SUBRs!±Ý¾× Then
            Print #FHandle, Line_Format(hhh$(7), TextData(), 9)
        Else
            Print #FHandle, Line_Format(hhh$(19), TextData(), 9)
        End If
        
        ' ¶óÀÎÀ» Áõ°¡ ½ÃÅ²´Ù.
        LineCnt = LineCnt + 1
        
        If PRINT_LINE_COUNT < LineCnt Then
            Print #FHandle, hhh$(11)
            TextData$(1) = PageCnt
            TextData$(2) = Prt_Total_Page_count
            Print #FHandle, Line_Format(hhh$(9), TextData(), 2)
            Print #FHandle, hhh$(10)
            GoSub Print_Head
            LineCnt = 0
        End If
        
        ' ¶óÀÎÀ» È®ÀÎÈÄ ÁöÁ¤µÈ ¶óÀÎ ÀÎ¼â½Ã ´ÙÀ½ ÆäÀÌÁö·Î ÀÌµ¿ ÇÑ´Ù.
        ' ³ª¸ÓÁö¸¦ ¹ÝÈ¯ÇÑ´Ù.
        SUBRs.MoveNext
    Loop
    SUBRs.Close
    Set SUBRs = Nothing
        
    Return
   
   
' ¸¶Áö¸· ºÎºÐÀ» Ãâ·Â ÇÑ´Ù.
Print_Bottom:

    Print #FHandle, hhh$(12)
    TextData(1) = CStr(ProssCount)
    TextData(2) = CStr(iSub1)
    TextData(3) = CStr(iSub2)
    TextData(4) = CStr(iSub3)
    TextData(5) = CStr(iSub4)
    TextData(6) = CStr(iSub5)
    Print #FHandle, Line_Format(hhh$(14), TextData(), 6)
    
    '-------------------------------------------------------------------
    '
    '-------------------------------------------------------------------
    Query = "SELECT * FROM ÀÏÀÏ¸¶°¨ "
    Query = Query & " WHERE ÀÏÀÚ='" & prtDay & "' "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    TextData(1) = 0: TextData(2) = 0: TextData(3) = 0: TextData(4) = 0
    
    If SUBRs.RecordCount > 0 Then
        TextData(1) = SUBRs.Fields("ÃÑ¸ÅÃâ¾×")
        TextData(2) = SUBRs.Fields("º»»ç±Ý¾×")
        TextData(3) = SUBRs.Fields("´ë¸®Á¡±Ý¾×")
        TextData(4) = SUBRs.Fields("¼ö¼±±Ý¾×")
    End If
    
    Print #FHandle, Line_Format(hhh$(16), TextData(), 4)
    
    '  ¼ö±Ý¾×/¿ÏºÒ / ¹ÌºÒ / ¹ÝÇ°È¯ºÒ
    TextData(1) = 0: TextData(2) = 0: TextData(3) = 0: TextData(4) = 0
    
    '------------------------------------------------------------------
    '
    '------------------------------------------------------------------
    Query = "SELECT SUM(±Ý¾×) AS ¼ö±Ý¾× FROM ¹Ì¼öÈ¸¼öÁ¤º¸ "
    Query = Query & " WHERE ÀÏÀÚ = '" & prtDay & "' "
    Set Rs = New ADODB.Recordset
    Rs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If Not Rs.EOF Then
        TextData(1) = Val(Rs!¼ö±Ý¾× & "")
    End If
    Rs.Close
    Set Rs = Nothing
    
    '------------------------------------------------------------
    '
    '------------------------------------------------------------
    Query = "SELECT SUM(±Ý¾×) AS ¿ÏºÒ FROM ÀÔÃâ°í "
    Query = Query & " WHERE ÀÔ°íÀÏ = '" & prtDay & "'"
    Query = Query & "   AND »óÅÂ   = 'èÇ'"
    Query = Query & "   AND (ÆÇ¸ÅÃë¼Ò IS NULL OR ÆÇ¸ÅÃë¼Ò <> 'Y') "
    Set Rs = New ADODB.Recordset
    Rs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If Not Rs.EOF Then
        TextData(2) = Val(Rs!¿ÏºÒ & "")
    End If
    Rs.Close
    Set Rs = Nothing
    
    '------------------------------------------------------------
    '
    '------------------------------------------------------------
    Query = "SELECT SUM(±Ý¾×) AS ¹ÌºÒ FROM ÀÔÃâ°í "
    Query = Query & " WHERE ÀÔ°íÀÏ = '" & prtDay & "'"
    Query = Query & "   AND »óÅÂ   = 'Ú±'"
    Query = Query & "   AND (ÆÇ¸ÅÃë¼Ò IS NULL OR ÆÇ¸ÅÃë¼Ò <> 'Y') "
    Set Rs = New ADODB.Recordset
    Rs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If Not Rs.EOF Then
        TextData(3) = Val(Rs!¹ÌºÒ & "")
    End If
    Rs.Close
    Set Rs = Nothing
    
    '------------------------------------------------------------
    '
    '------------------------------------------------------------
    Query = "SELECT SUM(±Ý¾×) AS È¯ºÒ FROM ÀÔÃâ°í "
    Query = Query & " WHERE LEFT(È¯ºÒÀÏÀÚ,8) = '" & prtDay & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If Not Rs.EOF Then
        TextData(4) = Val(Rs!È¯ºÒ & "")
        dblReturnMoney = Val(Rs!È¯ºÒ & "")
    End If
    Rs.Close
    Set Rs = Nothing
    
    
    Print #FHandle, Line_Format(hhh$(32), TextData(), 4)
    
    ' Ä«µå ±Ý¾×À» °¡Àú¿Â´Ù.
    If IsNull(SUBRs!Ä«µå±Ý¾×) = True Then
        dblCardMoney = 0
    ElseIf IsNumeric(SUBRs!Ä«µå±Ý¾×) = True Then
        dblCardMoney = Val(SUBRs!Ä«µå±Ý¾×)
    End If
    
    ' Ä«µå ±Ý¾×°Ç¼ö °¡Àú¿Â´Ù.
    If IsNull(SUBRs!Ä«µå°Ç¼ö) = True Then
        dblCardCount = 0
    ElseIf IsNumeric(SUBRs!Ä«µå°Ç¼ö) = True Then
        dblCardCount = Val(SUBRs!Ä«µå°Ç¼ö)
    End If
    
    ' ¼¼Å¹ºñ È¯ºÒ °Ç¼ö
    If IsNull(SUBRs!¼¼Å¹ºñÈ¯ºÒ°Ç¼ö) = True Then
        dblSaleReturnCnt = 0
    ElseIf IsNumeric(SUBRs!¼¼Å¹ºñÈ¯ºÒ°Ç¼ö) = True Then
        dblSaleReturnCnt = Val(SUBRs!¼¼Å¹ºñÈ¯ºÒ°Ç¼ö)
    End If
    
    ' ¼¼Å¹ºñ È¯ºÒ ±Ý¾×
    If IsNull(SUBRs!¼¼Å¹ºñÈ¯ºÒ±Ý¾×) = True Then
        dblSaleReturnMoney = 0
    ElseIf IsNumeric(SUBRs!¼¼Å¹ºñÈ¯ºÒ±Ý¾×) = True Then
        dblSaleReturnMoney = Val(SUBRs!¼¼Å¹ºñÈ¯ºÒ±Ý¾×)
    End If
    
    TextData(1) = dblCardMoney
    TextData(2) = dblCardCount
    TextData(3) = dblSaleReturnMoney
    Print #FHandle, Line_Format(hhh$(40), TextData(), 3)
    
'
    ' ¹ÝÇ° È¯ºÒ±Ý¾×¹× ¼¼Å¹ºñ È®ºÒ ±Ý¾× Ãâ·Â
    Print #FHandle, hhh$(17)
    TextData(1) = dblReturnMoney * dblRatio
    TextData(2) = dblSaleReturnMoney * dblRatio
    
    Print #FHandle, Line_Format(hhh$(38), TextData(), 2)
    Print #FHandle, hhh$(17)
    Print #FHandle, hhh$(18)
    
    ' ¸¶ÀÏ¸®Áö °ü·Ã
    TextData(1) = 10: TextData(2) = 0: TextData(3) = 0: TextData(4) = 0
    
    dblMilPrice(0) = 0: dblMilPrice(1) = 0: dblMilPrice(2) = 0: dblMilPrice(3) = 0
    
    If SUBRs.RecordCount > 0 Then
        dblMilPrice(0) = Val(SUBRs.Fields("¹ß»ý¸¶ÀÏ¸®Áö") & "")
        dblMilPrice(1) = Val(SUBRs.Fields("»ç¿ë¸¶ÀÏ¸®Áö") & "")
        dblMilPrice(2) = Val(SUBRs.Fields("»èÁ¦¸¶ÀÏ¸®Áö") & "")
    
        TextData(1) = CStr(dblMilPrice(0))
        TextData(2) = CStr(dblMilPrice(1))
        TextData(3) = CStr(dblMilPrice(2))
    End If
    
    Print #FHandle, Line_Format(hhh$(28), TextData(), 3)
    
    '-----------------------------------------------------------------------------
    ' ÄíÆù °ü·Ã ³»¿ë Ãâ·Â
    nCouponTotalMoney = GetCouponSaleTotalMoney(prtDay, nCouponTotal)
    
    If ´ë¸®Á¡Á¤º¸.MasterCode = M_COUPON_KLENZ_CODE Then
        TextData(1) = GetCouponCost("00")
    Else
        TextData(1) = GetCouponCost("01")
    End If
    
    TextData(2) = CStr(nCouponTotalMoney)
    TextData(3) = SUBRs.Fields("ÃÑ¸ÅÃâ¾×")
    
    Print #FHandle, Line_Format(hhh$(43), TextData(), 3)
    
    Print #FHandle, hhh$(29)
    
    ' ¾Æ·¡ ÀüÃ¼ ¸ÅÃâ¿¡ °üÇÏ¿© Ã³¸®ÇÏ±â À§ÇÏ¿©
    ' ÄíÆùÀÌ ¿©·¯ Á¾·ù »ç¿ëµÇ¾úÀ» °æ¿ì °è»ê Ã³¸®ÇÏ¿© ±Ý¾×À» °¡Àú¿Â´Ù.
    nCouponTotalMoney2 = GetCouponSaleTotalMoney2(prtDay)
    '-----------------------------------------------------------------------------
    
    '-----------------------------------------------------------------------------
    ' ÃÑ¸ÅÃâ¾×, ¸¶ÀÏ¸®Áö, ÃÖÁ¾½ÇÀÔ±Ý¾× Ãâ·Â
    TextData(1) = dblMilPrice(1)
    TextData(2) = SUBRs.Fields("ÃÑ¸ÅÃâ¾×") - dblMilPrice(1)
    Print #FHandle, Line_Format(hhh$(44), TextData(), 2)
    
    '-----------------------------------------------------------------------------
    
    
'    '-----------------------------------------------------------------------------
'    ' Äí¿£¼Öºê »ç¿ë ¾ÈÇÔ
'    TextData(1) = 0: TextData(2) = 0: TextData(3) = 0: TextData(4) = 0
'    dblQNPrice(0) = 0: dblQNPrice(1) = 0: dblQNPrice(2) = 0: dblQNPrice(3) = 0
'    dblQNPrice(0) = QN_Day_Info(prtDay, dblQNPrice(1), dblQNPrice(2), dblQNPrice(3), dblQNPrice(4))
'    If dblQNPrice(0) > 0 Then
'        TextData(1) = CStr(dblQNPrice(0))
'        TextData(2) = CStr(dblQNPrice(2))
'        TextData(3) = CStr(dblQNPrice(3))
'        TextData(4) = CStr(dblQNPrice(4))
'    End If
'    Print #FHandle, Line_Format(hhh$(39), TextData(), 4)
'    Print #FHandle, hhh$(29)
'    '-----------------------------------------------------------------------------
    
    TextData(1) = 0: TextData(2) = 0: TextData(3) = 0: TextData(4) = 0
    ' 2009.05.19ÀÏ Ãâ·Â ³»¿ë º¯°æ
    TextData(1) = SUBRs.Fields("ÃÑ¸ÅÃâ¾×") - nCouponTotalMoney2 - dblMilPrice(1)
    TextData(2) = SUBRs.Fields("º»»ç±Ý¾×") - (nCouponTotalMoney2 * dblRatio) - (dblMilPrice(1) * dblRatio)
    TextData(3) = SUBRs.Fields("´ë¸®Á¡±Ý¾×") - (nCouponTotalMoney2 * (1 - dblRatio)) - (dblMilPrice(1) * (1 - dblRatio))
    
   
'   TextData(1) = SUBRs.Fields("ÃÑ¸ÅÃâ¾×") - (nCouponTotal * 2000) - dblMilPrice(1)
'   TextData(2) = SUBRs.Fields("º»»ç±Ý¾×") - ((nCouponTotal * 2000) * dblRatio) - (dblMilPrice(1) * dblRatio)
'   TextData(3) = SUBRs.Fields("´ë¸®Á¡±Ý¾×") - ((nCouponTotal * 2000) * (1 - dblRatio)) - (dblMilPrice(1) * (1 - dblRatio))
    
    Print #FHandle, Line_Format(hhh$(30), TextData(), 3)
    Print #FHandle, hhh$(31)
    
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '+  »ï¼ºÄ«µå °ü·Ã ³»¿ë Ãâ·Â
    TextData(1) = 0: TextData(2) = 0: TextData(3) = 0: TextData(4) = 0
    TextData(1) = SUBRs.Fields("»ï¼ºÄ«µåÇÒÀÎ°í°´¼ö") & ""
    TextData(2) = SUBRs.Fields("»ï¼ºÄ«µåÇÒÀÎ°Ç¼ö") & ""
    TextData(3) = SUBRs.Fields("»ï¼ºÄ«µåÇÒÀÎ±Ý¾×") & ""
    
    Print #FHandle, Line_Format(hhh$(47), TextData(), 3)
    'Print #FHandle, hhh$(31)
    SUBRs.Close
    Set SUBRs = Nothing
    
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '+ ¼¼Æ® »óÇ° ³»¿ë Ãâ·Â
'   hhh$(48) = "  ¼¼Æ®»óÇ° Á¤º¸    W2Á¾¼¼Æ® #,##0 °Ç,  W3Á¾¼¼Æ® #,##0 °Ç, W4Á¾¼¼Æ® #,##0 °Ç, W5Á¾¼¼Æ® #,##0 °Ç, W6Á¾¼¼Æ® #,##0 °Ç "
'   hhh$(49) = "                   ¹«·á¼¼Å¹±Ç¼ö: #,##0 Àå   °í°´¼ö: #,##0 ¸í  ¼¼Æ®ÇÒÀÎ±Ý¾×: ###,##0 ¿ø ¿¡´©¸®ÇÒÀÎ±Ý¾×: ###,##0 ¿ø "
    
    If ´ë¸®Á¡Á¤º¸.MasterCode <> M_COUPON_KLENZ_CODE Then
        Query = "SELECT count(°í°´ÄÚµå)        as Cnt1"
        Query = Query & ", sum(¼¼Æ®ÇÒÀÎ±Ý¾×)   as Cnt2"
        Query = Query & ", sum(¿¡´©¸®ÇÒÀÎ±Ý¾×) as Cnt3"
        Query = Query & ", sum(¹«·á¼¼Å¹±Ç¼ö)   as Cnt4"
        Query = Query & ", SUM(¼¼Æ®2) AS WS2"
        Query = Query & ", SUM(¼¼Æ®3) AS WS3"
        Query = Query & ", SUM(¼¼Æ®4) AS WS4"
        Query = Query & ", SUM(¼¼Æ®5) AS WS5"
        Query = Query & ", SUM(¼¼Æ®6) AS WS6  "
        Query = Query & " FROM ¼¼Æ®»óÇ°Á¤º¸ "
        Query = Query & " WHERE Á¢¼öÀÏÀÚ = '" & prtDay & "' "
        Set Rs = New ADODB.Recordset
        Rs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
        
        If Rs.RecordCount >= 1 Then
            TextData(1) = 0: TextData(2) = 0: TextData(3) = 0: TextData(4) = 0: TextData(5) = 0
            TextData(1) = Val(Rs.Fields("WS2") & "")
            TextData(2) = Val(Rs.Fields("WS3") & "")
            TextData(3) = Val(Rs.Fields("WS4") & "")
            TextData(4) = Val(Rs.Fields("WS5") & "")
            TextData(5) = Val(Rs.Fields("WS6") & "")
            Print #FHandle, Line_Format(hhh$(48), TextData(), 5)
        
            TextData(1) = 0: TextData(2) = 0: TextData(3) = 0: TextData(4) = 0: TextData(5) = 0
            TextData(1) = Val(Rs.Fields("Cnt4") & "")
            TextData(2) = Val(Rs.Fields("Cnt1") & "")
            TextData(3) = Val(Rs.Fields("Cnt2") & "")
            TextData(4) = Val(Rs.Fields("Cnt3") & "")
            Print #FHandle, Line_Format(hhh$(49), TextData(), 4)
        End If
        Rs.Close
        Set Rs = Nothing
    End If
    
    
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '+        ' ´©¶ôTAG CHECK
    Dim strSTag As String       '½ÃÀÛ ÅÃ¹øÈ£
    Dim strETag As String       'Á¾·á ÅÃ¹øÈ£
    Dim strTempTag As String    '´©¶ôÅÃÀ» ÀúÀå
    Dim sMemTagNo As String     'ÅÃ¹øÈ£ °Ë»ç½Ã ÇÊ¿ä
    
    '½ÃÀÛ-Á¾·áÅÃ¹øÈ£ ±¸ÇÏ±â
    Query = "SELECT MIN(¹øÈ£) AS STAG, MAX(¹øÈ£) AS ETAG "
    Query = Query & "FROM ÀÔÃâ°í "
    Query = Query & "WHERE ÀÔ°íÀÏ = '" & prtDay & "' "
    Query = Query & "AND   (ÆÇ¸ÅÃë¼Ò IS NULL OR ÆÇ¸ÅÃë¼Ò <> 'Y') "
    Set Rs = New ADODB.Recordset
    Rs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If Rs.RecordCount < 1 Then
        strSTag = ""
        strETag = ""
    Else
        strSTag = IIf(IsNull(Rs!sTag), "", Mid(Rs!sTag, 1, 1) & Mid(Rs!sTag, 3, 3))
        strETag = IIf(IsNull(Rs!ETAG), "", Mid(Rs!ETAG, 1, 1) & Mid(Rs!ETAG, 3, 3))
    End If
    Rs.Close
    Set Rs = Nothing
    
    '----------------------------------------------------------------
    ' ´©¶ô TAG ±¸ÇÏ±â
    '----------------------------------------------------------------
    Query = "SELECT ¹øÈ£ "
    Query = Query & "FROM ÀÔÃâ°í "
    Query = Query & "WHERE ÀÔ°íÀÏ = '" & prtDay & "' "
    Query = Query & "AND   (ÆÇ¸ÅÃë¼Ò IS NULL OR ÆÇ¸ÅÃë¼Ò <> 'Y') "
    Query = Query & "ORDER BY ¹øÈ£ "
    Set Rs = New ADODB.Recordset
    Rs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If Val(strETag) - Val(strSTag) < 5000 Then ' ÀÏÀÏ ÃÖ°í ÆÇ¸Å ¼ö·® 5000Á¡
        Do While Not Rs.EOF
            sMemTagNo = Left(Rs!¹øÈ£, 1) & Right(Rs!¹øÈ£, 3)
            
            Rs.MoveNext
            
            If Rs.EOF Then
                Exit Do
            Else
                Do While Val(sMemTagNo) + 1 <> Val(Left(Rs!¹øÈ£, 1) & Right(Rs!¹øÈ£, 3))
                    If Val(sMemTagNo) + 1 < 1000 Then
                        sMemTagNo = Val(sMemTagNo) + 1
                        strTempTag = strTempTag + "0" & Format(Replace(sMemTagNo, "-", ""), "-@@@, ")
                    Else
                        sMemTagNo = Val(sMemTagNo) + 1
                        strTempTag = strTempTag + Format(Replace(sMemTagNo, "-", ""), "@-@@@, ")
                    End If
                Loop
            End If
        Loop
    End If
    Rs.Close
    Set Rs = Nothing
    
    
    ' ´©¶ôÅÃ Ãâ·Â
    Print #FHandle, hhh$(20)
    TextData(1) = Format(Val(Len(strTempTag) / 7), "#0")
    TextData(2) = Format(strSTag, "@-@@@")
    TextData(3) = Format(strETag, "@-@@@")
    If Len(strTempTag) Then strTempTag = Mid(strTempTag, 1, Len(strTempTag) - 2)    ' ¸¶Áö¸· ","Áö¿ì±â
    TextData(4) = Mid(strTempTag, 1, 70)
    Print #FHandle, Line_Format(hhh$(21), TextData(), 4)
    strTempTag = Mid(strTempTag, 71, Len(strTempTag))
    
    Do While Len(strTempTag) > 4
        TextData(1) = Mid(strTempTag, 1, 70)
        Print #FHandle, Line_Format(hhh$(22), TextData(), 1)
        strTempTag = Mid(strTempTag, 71, Len(strTempTag))
    Loop
    Print #FHandle, hhh$(23)
    
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '+        ' Àç»ç¿ë AG CHECK
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    strTempTag = ""
    
    Query = "SELECT ¹øÈ£ FROM ÀÔÃâ°í "
    Query = Query & " WHERE ÀÔ°íÀÏ   = '" & prtDay & "' "
    Query = Query & "   AND ÆÇ¸ÅÃë¼Ò = 'R' "
    Query = Query & " ORDER BY ¹øÈ£ "
    Set Rs = New ADODB.Recordset
    Rs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    Do While Not Rs.EOF
        strTempTag = strTempTag + Format(Replace(Rs!¹øÈ£, "-", ""), "@-@@@, ")
        
        Rs.MoveNext
    Loop
    Rs.Close
    Set Rs = Nothing
    
    '
    TextData(1) = Format(Val(Len(strTempTag) / 7), "#0")
    
    If Len(strTempTag) Then
        strTempTag = Mid(strTempTag, 1, Len(strTempTag) - 2)    ' ¸¶Áö¸· ","Áö¿ì±â
    End If
    
    TextData(2) = Mid(strTempTag, 1, 70)
    
    Print #FHandle, Line_Format(hhh$(24), TextData(), 2)
    
    strTempTag = Mid(strTempTag, 71, Len(strTempTag))
    
    Do While Len(strTempTag) > 4
        TextData(1) = Mid(strTempTag, 1, 70)
        Print #FHandle, Line_Format(hhh$(25), TextData(), 1)
        strTempTag = Mid(strTempTag, 71, Len(strTempTag))
    Loop
    Print #FHandle, hhh$(26)
    
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '+            ' ¹ÝÇ°ÅÃ Ãâ·Â
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    strTempTag = ""
    
    Query = "SELECT ¹øÈ£ FROM ÀÔÃâ°í "
    Query = Query & " WHERE LEFT(È¯ºÒÀÏÀÚ,8) ='" & prtDay & "' "
    Query = Query & " ORDER BY ¹øÈ£ "
    Set Rs = New ADODB.Recordset
    Rs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    Do While Not Rs.EOF
        strTempTag = strTempTag + Format(Replace(Rs!¹øÈ£, "-", ""), "@-@@@, ")
        
        Rs.MoveNext
    Loop
    Rs.Close
    Set Rs = Nothing
    
    TextData(1) = Format(Val(Len(strTempTag) / 7), "#0")
    If Len(strTempTag) Then strTempTag = Mid(strTempTag, 1, Len(strTempTag) - 2)    ' ¸¶Áö¸· ","Áö¿ì±â
    TextData(2) = Mid(strTempTag, 1, 70)
    Print #FHandle, Line_Format(hhh$(33), TextData(), 2)
    strTempTag = Mid(strTempTag, 71, Len(strTempTag))
    
    Do While Len(strTempTag) > 4
        TextData(1) = Mid(strTempTag, 1, 70)
        Print #FHandle, Line_Format(hhh$(34), TextData(), 1)
        strTempTag = Mid(strTempTag, 71, Len(strTempTag))
    Loop
    
    Print #FHandle, hhh$(26)
    
    
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '+            ' ¼¼Å¹ºñÈ¯ºÒ ÅÃ ¹øÈ£ Ãâ·Â
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    i = 0
    
    strTempTag = ""
    
    Query = "SELECT ÀÔ°íÀÏ, ¹øÈ£, °í°´¹øÈ£ "
    Query = Query & "FROM ÀÔÃâ°í "
    Query = Query & "WHERE LEFT(¼¼Å¹ºñÈ¯ºÒÀÏÀÚ,8) ='" & prtDay & "' "
    Query = Query & "ORDER BY ¹øÈ£ "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    Do While Not SUBRs.EOF
        Call Fb°í°´Á¤º¸(SUBRs!°í°´¹øÈ£ & "")
        
        If i = 0 Then
            TextData(1) = CStr(dblSaleReturnCnt)
            TextData(2) = Format(SUBRs!ÀÔ°íÀÏ, "@@@@-@@-@@") & "  " & Format(Replace(SUBRs!¹øÈ£, "-", ""), "@-@@@") & "  " & °í°´Á¤º¸.¼º¸í & " ( " & °í°´Á¤º¸.ÈÞ´ëÆù & " )"
            
            Print #FHandle, Line_Format(hhh$(45), TextData(), 2)
        Else
            TextData(1) = Format(SUBRs!ÀÔ°íÀÏ, "@@@@-@@-@@") & "  " & Format(Replace(SUBRs!¹øÈ£, "-", ""), "@-@@@") & "  " & °í°´Á¤º¸.¼º¸í & " ( " & °í°´Á¤º¸.ÈÞ´ëÆù & " )"
            Print #FHandle, Line_Format(hhh$(46), TextData(), 1)
        End If
        
        i = i + 1
        
        SUBRs.MoveNext
    Loop
    SUBRs.Close
    
    ' --------------------------------------------------------------------------------------
    ' ÄíÆù ¼ö·®
    ' --------------------------------------------------------------------------------------
    Dim nCPrtCnt        As Integer
    Dim nMaxPrtCnt      As Integer
    
    nMaxPrtCnt = 5
    
    Query = "SELECT ÄíÆù¹øÈ£, °í°´ÀÌ¸§  FROM ÄíÆùÀÚ·á "
    Query = Query & " WHERE Á¢¼öÀÏÀÚ = '" & prtDay & "' "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If Not SUBRs.EOF Then
        SUBRs.MoveLast
        nCouponTotal = CStr(SUBRs.RecordCount)
        SUBRs.MoveFirst
        
        TextData(1) = CStr(nCouponTotal)
        
        Do Until SUBRs.EOF
            nCPrtCnt = nCPrtCnt + 1
            
            strCouponNumber = strCouponNumber & CStr(SUBRs.Fields("ÄíÆù¹øÈ£")) & "(" & Trim(CStr(SUBRs.Fields("°í°´ÀÌ¸§"))) & "),"
            
            If nCPrtCnt = nMaxPrtCnt Then
                strCouponNumber = Left(strCouponNumber, Len(strCouponNumber) - 1)

                If nCPrtCnt <= nCouponTotal Then
                    TextData(2) = strCouponNumber
                    Print #FHandle, Line_Format(hhh$(41), TextData(), 2)
                Else
                    TextData(1) = strCouponNumber
                    Print #FHandle, Line_Format(hhh$(42), TextData(), 1)
                End If

                strCouponNumber = ""
                nCPrtCnt = 0
            End If
            
            SUBRs.MoveNext
        Loop

        If strCouponNumber <> "" Then
            ' 9°³ ¹Ì¸¸ÀÌ¿©¼­ ÃÖÃÊ ÀÎ¼âµÉ°æ¿ì
            If nCouponTotal < nMaxPrtCnt Then
                strCouponNumber = Left(strCouponNumber, Len(strCouponNumber) - 1)
                TextData(2) = strCouponNumber
                Print #FHandle, Line_Format(hhh$(41), TextData(), 2)
    
            ' 9°³ ÀÌ»óÀÌ¸ç ³ª¸ÓÁö Ãâ·ÂÀÏ °æ¿ì
            ElseIf nCouponTotal > nMaxPrtCnt Then
                strCouponNumber = Left(strCouponNumber, Len(strCouponNumber) - 1)
                TextData(1) = strCouponNumber
                Print #FHandle, Line_Format(hhh$(42), TextData(), 1)
            End If
        End If
    End If
    SUBRs.Close
    
    ' --------------------------------------------------------------------------------------
    ' --------------------------------------------------------------------------------------
    Print #FHandle, hhh$(35)
    Print #FHandle, hhh$(27)
    
    If Format(Date, "yyyyMMdd") <= "20090831" And ´ë¸®Á¡Á¤º¸.MasterCode <> M_COUPON_KLENZ_CODE Then
        Print #FHandle, Space(10) & M_CompnyMasterName & " LG Å¸¿îÁ¨Æ® Çà»ç ±â°£Àº 8¿ù 31ÀÏ ±îÁö ÀÔ´Ï´Ù."
    End If
     
     ' ÆäÀÌÁö ºÎºÐ Ãâ·Â
    TextData(1) = PageCnt
    TextData(2) = Prt_Total_Page_count
    Print #FHandle, Line_Format(hhh$(9), TextData(), 2)
    Close #FHandle
    
    Return

'Error Ã³¸®ºÎ
Error_Rtn:
'    Dim strMsg As String
    Close #FHandle
    strMsg = "Error Number : " & CStr(Err.Number) & Chr(10) & Chr(13) & _
        "Error Description : " & Err.Description
    MsgBox strMsg, 16, "Error Message!"
    Resume Next
End Function

Function subMonthListPrint(cdPrt As CommonDialog, prtMonth As String)
    
    Dim i As Long
    Dim kk As Long
    Dim FHandle As Integer                  ' ÀÎ¼âÇÒ ÆÄÀÏÀÇ ÇÚµé
    Dim ProssCount As Integer         ' ÀüÃ¼ ÆäÀÌÁö¿¡¼­ Ãâ·ÂµÉ ÃÑ ¾ÆÀÌÅÛ ÃÑ °¹¼ö
    Dim Prt_Total_Page_count As Integer     ' Ãâ·ÂµÉ ÀüÃ¼ ÆäÀÌÁö¼ö
    Dim PRINT_LINE_COUNT As Integer          ' ÇÑÆäÀÌÁö´ç Ãâ·ÂµÉ ¾ÆÀÌÅÛ °¹¼ö
    Dim PageCnt As Integer                  ' ÇöÀç Ãâ·ÂÁßÀÎ ÆäÀÌÁö
    Dim LineCnt As Integer                  ' ÇöÀç Ãâ·ÂÁßÀÎ ¶óÀÎ
    Dim strFileName As String               ' Ãâ·Â ÆÄÀÏ¸í
    Dim TextData(20) As String              ' ÀÎ¼âÇÒ ³»¿ëÀ» ÀÓ½Ã ÀúÀåÇÑ´Ù.
    Dim hhh(60) As String                   ' ¾ç½ÄÀ» ÀúÀåÇÑ´Ù.

    Dim BottomValue1    As Integer          ' ÃÑÁ¡¼ö
    Dim BottomValue2    As Integer          ' ÃÑ¼öÀü
    Dim BottomValue3    As Integer          ' ÃÑ¹ÝÇ°
    Dim BottomValue4    As Integer          ' Àç¼¼Å¹
    Dim BottomValue5    As Double          ' ÃÑ±Ý¾×
    
    Dim strQuery As String
    Dim Rs As Recordset
    Dim rsReprint As DAO.Recordset
    Dim strMsg As String
    
    Screen.MousePointer = 13
    
    On Error GoTo Error_Rtn
    
    ' Ãâ·Â ÀÚ·áÀÇ ¾çÀ» ±¸ÇÑ´Ù., ÄÚµå DB¸¦ ¿ÀÇÂ ÇÑ´Ù.
    ' Ãâ·Â ³»¿ëÀÌ ¾øÀ» °æ¿ì ¹Ù·Î Á¾·á ÇÑ´Ù.
    GoSub Print_ProssCount
    ' Ãâ·Â ¾ç½ÄÀ» ÃÊ±âÈ­ ÇÑ´Ù
    GoSub Print_Init
    ' Ãâ·ÂÇÒ ÆÄÀÏÀ» ¿ÀÇÂ ÇÑ´Ù.
    GoSub Print_FileOpen
    ' ¾ç½ÄÀÇ Å¸ÀÌÆ²À» È­ÀÏ¿¡ Ãâ·Â ÇÑ´Ù.
    GoSub Print_Head
    ' ¹Ýº¹µÇ´Â ºÎºÐÀ» È­ÀÏ¿¡ Ãâ·Â ÇÑ´Ù.
    GoSub Print_Next
    ' ¸¶Áö¸· ºÎºÐÀ» È­ÀÏ¿¡ Ãâ·Â ÇÑ´Ù.
    ' Ãâ·ÂÇÑ ÆÄÀÏÀ» ´Ý´Â´Ù.
    GoSub Print_Bottom
    
    Screen.MousePointer = 0
    ' È­ÀÏÀ» ÇÁ¸°ÅÍ·Î Ãâ·ÂÇÑ´Ù.
    Call FileToPrint(strFileName, 1, False)
    Exit Function
    
'Ãâ·ÂÇÒ ÀÚ·áÀÇ Ä«¿îÅÍ¸¦ È®ÀÎÇÑ´Ù.
Print_ProssCount:
    '-------------------------------------------------------------
    '
    '-------------------------------------------------------------
    Query = "SELECT * FROM ÀÏÀÏ¸¶°¨ "
    Query = Query & " WHERE Mid(ÀÏÀÚ, 1, 6) = '" & prtMonth & "' "
    Query = Query & " ORDER BY ÀÏÀÚ"
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If SUBRs.RecordCount > 0 Then
        SUBRs.MoveLast
        ProssCount = SUBRs.RecordCount
        SUBRs.MoveFirst
    Else
        SUBRs.Close
        Set SUBRs = Nothing
        
        MsgBox "Ãâ·ÂÇÒ ³»¿ëÀÌ ¾ø½À´Ï´Ù.", vbInformation, "È®ÀÎ"
        Exit Function
    End If
    SUBRs.Close
    Set SUBRs = Nothing
    
    Return

' Ãâ·ÂÇÒ ¾ç½ÄÀ» ÃÊ±âÈ­ ÇÑ´Ù.
Print_Init:
    ' Ãâ·Â ÆÄÀÏ¸í
    strFileName = App.Path & "\PrintRep.txt"
    
    hhh$(1) = "                        @@@@ ³â @@ ¿ù ¸ÅÃâÇöÈ²                                                                  "
    hhh$(2) = "                    ¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬                                                                      "
    hhh$(3) = "                                                                                              ÀÏ ÀÚ : !!!!/!!/!!  "
    hhh$(4) = " ¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬ "
    hhh$(5) = "      ÀÔ°íÀÏÀÚ            ÃÑÁ¡¼ö              ¼ö  ¼±              ¹Ý  Ç°              Àç¼¼Å¹            ±Ý  ¾×    "
    hhh$(6) = " ¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦¬ "
    hhh$(7) = "         @@               #,###               #,###               #,###               #,###          ###,###,###  "
    hhh$(8) = "   @@@@  ¦¢ @@@@@@@@@  ¦¢  @@@@@@  ¦¢ @@@@@@@@@@@@ ¦¢  ###,### ¦¢ @@@@@@@@ ¦¢ @@@@@@@@ ¦¢ @@@@@@@@@@ ¦¢@@@@@@@@@@ "
    hhh$(9) = "                                                                                                                  "
    hhh$(10) = "                                                                                              Page  : ### / ###   "
    hhh$(11) = " ¦¬¦¬¦¬¦¬¦º¦¬¦¬¦¬¦¬¦¬¦¬¦º¦¬¦¬¦¬¦¬¦¬¦º¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦º¦¬¦¬¦¬¦¬¦¬¦º¦¬¦¬¦¬¦¬¦¬¦º¦¬¦¬¦¬¦¬¦¬¦º¦¬¦¬¦¬¦¬¦¬¦¬¦º¦¬¦¬¦¬¦¬¦¬ "
    hhh$(12) = " ¦¬¦¬¦¬¦¬¦á¦¬¦¬¦¬¦¬¦¬¦¬¦á¦¬¦¬¦¬¦¬¦¬¦á¦¬¦¬¦¬¦¬¦¬¦¬¦¬¦á¦¬¦¬¦¬¦¬¦¬¦á¦¬¦¬¦¬¦¬¦¬¦á¦¬¦¬¦¬¦¬¦¬¦á¦¬¦¬¦¬¦¬¦¬¦¬¦á¦¬¦¬¦¬¦¬¦¬ "
    hhh$(13) = "                                                                                                                  "
    hhh$(14) = "      ÇÕ°è :              #,### Á¡            #,### Á¡            #,### Á¡            #,### Á¡       ###,###,###  "
    hhh$(15) = ""
    hhh$(16) = " - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
    
    ' ÆäÀÌÁö ¹× ¶óÀÎÀ» ÃÊ±âÈ­ ÇÑ´Ù.
    PageCnt = 0:  LineCnt = 0
     ' ÇÑÆäÀÌÁö´ç Ãâ·ÂµÉ ¾ÆÀÌÅÛ °¹¼ö
    PRINT_LINE_COUNT = GetPrtItemCount("¿ùº°¸ÅÃâÇöÈ²")
    ' ÀüÃ¼ Ãâ·Â ÆäÀÌÁö ±¸ÇÏ±â
    If (ProssCount Mod PRINT_LINE_COUNT) <> 0 Then
        Prt_Total_Page_count = Int(ProssCount / PRINT_LINE_COUNT) + 1
    Else
        Prt_Total_Page_count = Int(ProssCount / PRINT_LINE_COUNT)
    End If
    Return
    
' Ãâ·ÂÇÒ ÆÄÀÏÀ» ¿ÀÇÂ ÇÑ´Ù.
Print_FileOpen:
    FHandle = FreeFile
    Open strFileName For Output As #FHandle
    Return
  
'¹Ýº¹µÇ´Â Å¸ÀÌÆ²À» Ãâ·Â ÇÑ´Ù.
Print_Head:
    PageCnt = PageCnt + 1
    LineCnt = 0
    Print #FHandle, hhh$(13) '¿©¹é Ãâ·Â
    Print #FHandle, hhh$(13) '¿©¹é Ãâ·Â
    Print #FHandle, hhh$(13) '¿©¹é Ãâ·Â
    Print #FHandle, hhh$(13) '¿©¹é Ãâ·Â
    Print #FHandle, hhh$(13) '¿©¹é Ãâ·Â
    Print #FHandle, hhh$(13) '¿©¹é Ãâ·Â
    TextData$(1) = Mid(prtMonth, 1, 4)
    TextData$(2) = Mid(prtMonth, 5, 2)
    Print #FHandle, Line_Format(hhh$(1), TextData(), 2)
    Print #FHandle, hhh$(2)
    
    TextData$(1) = Date
    Print #FHandle, Line_Format(hhh$(3), TextData(), 1)
    Print #FHandle, hhh$(4)
    Print #FHandle, hhh$(5)
    Print #FHandle, hhh$(6)
    Return
    
' Áß°£ÀÇ ¹Ýº¹ ºÎºÐÀÇ ÀÚ·á¸¦ Ãâ·Â ÇÑ´Ù.
Print_Next:
    ' ´ÙÀ½¿¡ Ãâ·ÂÇÒ ÀÚ·á°¡ ¾øÀ»¶§ ±îÁú
    '---------------------------------------------------------------
    Query = "SELECT * FROM ÀÏÀÏ¸¶°¨ "
    Query = Query & " WHERE Mid(ÀÏÀÚ, 1, 6) = '" & prtMonth & "' "
    Query = Query & " ORDER BY ÀÏÀÚ"
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
        
    If SUBRs.RecordCount <= 0 Then
        SUBRs.Close
        Set SUBRs = Nothing
        
        MsgBox "¸ÅÃâ ³»¿ëÀÌ ¾ø½À´Ï´Ù.", vbInformation, "È®ÀÎ"
        Exit Function
    End If
    
    Do Until SUBRs.EOF
        TextData$(1) = Space(1):    TextData$(2) = Space(1):    TextData$(3) = Space(1):
        TextData$(4) = Space(1):    TextData$(5) = Space(1):    TextData$(6) = Space(1):
        
        TextData$(1) = Format(Format(SUBRs!ÀÏÀÚ, "0000-00-00"), "dd")
        TextData$(2) = SUBRs!ÃÑÁ¡¼ö
        TextData$(3) = SUBRs!¼ö¼±¼ö·®
        TextData$(4) = SUBRs!¹ÝÇ°¼ö·®
        TextData$(5) = SUBRs!Àç¼¼Å¹¼ö·®
        TextData$(6) = SUBRs!ÃÑ¸ÅÃâ¾×
        
        BottomValue1 = BottomValue1 + SUBRs!ÃÑÁ¡¼ö
        BottomValue2 = BottomValue2 + SUBRs!¼ö¼±¼ö·®
        BottomValue3 = BottomValue3 + SUBRs!¹ÝÇ°¼ö·®
        BottomValue4 = BottomValue4 + SUBRs!Àç¼¼Å¹¼ö·®
        BottomValue5 = BottomValue5 + SUBRs!ÃÑ¸ÅÃâ¾×
        Print #FHandle, Line_Format(hhh$(7), TextData(), 6)
        
        ' ¶óÀÎÀ» Áõ°¡ ½ÃÅ²´Ù.
        LineCnt = LineCnt + 1
        
        If PRINT_LINE_COUNT < LineCnt Then
            Print #FHandle, hhh$(6)
            TextData$(1) = PageCnt
            TextData$(2) = Prt_Total_Page_count
            Print #FHandle, Line_Format(hhh$(10), TextData(), 2)
            Print #FHandle, hhh$(13)
            
            GoSub Print_Head
            
            LineCnt = 0
        End If
        
        ' ¶óÀÎÀ» È®ÀÎÈÄ ÁöÁ¤µÈ ¶óÀÎ ÀÎ¼â½Ã ´ÙÀ½ ÆäÀÌÁö·Î ÀÌµ¿ ÇÑ´Ù.
        ' ³ª¸ÓÁö¸¦ ¹ÝÈ¯ÇÑ´Ù.
        
        SUBRs.MoveNext
        
        If (LineCnt Mod 5) = 0 Then
            If Not SUBRs.EOF Then Print #FHandle, hhh$(16)
        End If
    Loop
    SUBRs.Close
    Set SUBRs = Nothing
    
    Return
   
   
' ¸¶Áö¸· ºÎºÐÀ» Ãâ·Â ÇÑ´Ù.
Print_Bottom:

    Print #FHandle, hhh$(6)
    TextData(1) = BottomValue1
    TextData(2) = BottomValue2
    TextData(3) = BottomValue3
    TextData(4) = BottomValue4
    TextData(5) = BottomValue5
    Print #FHandle, Line_Format(hhh$(14), TextData(), 5)
    Print #FHandle, hhh$(6)
    Close #FHandle
    Return

'Error Ã³¸®ºÎ
Error_Rtn:
'    Dim strMsg As String
    Close #FHandle
    strMsg = "Error Number : " & CStr(Err.Number) & Chr(10) & Chr(13) & _
        "Error Description : " & Err.Description
    MsgBox strMsg, 16, "Error Message!"
    Resume Next
End Function

' »ç°íÁ¢¼ö º¸°í¼­¸¦ Ãâ·ÂÇÑ´Ù
Public Function PrintSagoReport(cdPrt As CommonDialog, prtNum As String) As Boolean
    PrintSagoReport = True
    
    ' ±âº» ÇÁ¸°ÅÍ°¡ ¾øÀ» °æ¿ì
    If Not PrinterCheck Then
        PrintSagoReport = False
        Exit Function
    End If
        
    ''''''''''''''''
    On Error GoTo printError
    '''''''''''''''
    
    Query = " SELECT * FROM »ç°íÇ° WHERE ÀÏ·Ã¹øÈ£ = " & Val(prtNum) & " "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    
Print_Start:
    Prt_Top = 5
    Prt_Left = 15

    cdPrt.Flags = cdlPDHidePrintToFile

    Printer.ScaleMode = vbMillimeters           ' ÇÁ¸°ÅÍÀÇ ½ºÄÉÀÏ ¸ðµå¸¦ ¹Ð¸®¹ÌÅÍ ´ÜÀ§·Î
    
    Printer.FontName = "±¼¸²Ã¼"
    Printer.Font.Bold = True
    Printer.Font.Size = 9
    
    Printer.Font.Size = "18"
    PrintText 55, 0, "»ç °í Á¢ ¼ö º¸ °í ¼­"
    Printer.DrawWidth = 12
    PrintLine 50, 8, 125, 8
    
    
    ' °áÀç ¾ç½Ä ÀÛ¼º
    GoSub PrintApproval
    
    ' ±âº»»çÇ× ¾ç½Ä ÀÛ¼º
    GoSub PrintDefault
    
    ' ÇÇÇØ°ü·Ã»çÇ×
    GoSub PrintDamage
    
    ' ´ë¸®Á¡ ±âÀç
    GoSub AgencyWrite
    
    ' º»»ç ±âÀç
    GoSub CompanyWrite
    
    'Ãâ·Â ÇÑ´Ù.
    SUBRs.Close
    Set SUBRs = Nothing
    
    Printer.EndDoc
    
    Exit Function
    
printError:
    PrintSagoReport = False
    Printer.EndDoc
    Exit Function
    
' °áÀç ¾ç½Ä ÀÎ¼â
PrintApproval:

    Top_Margin = 0: Left_Margin = 0
    
    Printer.Font.Size = 9
    Printer.DrawWidth = 7
    PrintRect 105, 18, 175, 37  '¿Ü°¢ Æ²
    PrintLine 110, 24, 175, 24  '¼öÆò ¶óÀÎ
    
    PrintLine 110, 18, 110, 37  ' °áÀç
    PrintText 106, 20, "°á"
    PrintText 106, 32, "Àç"
    PrintText 113, 19, "°æ  ¸®"
    
    PrintLine 126, 18, 126, 37
    PrintText 129, 19, "´ã  ´ç"
    
    PrintLine 142, 18, 142, 37
    PrintText 145, 19, "Â÷  Àå"
    
    PrintLine 158, 18, 158, 37
    PrintText 161, 19, "»ç  Àå"
    Return
    
' ±âº»»çÇ× ¾ç½Ä ÀÛ¼º
PrintDefault:
    
    Top_Margin = 0: Left_Margin = 0
    
    Printer.Font.Size = 9
    Printer.DrawWidth = 7
    Text_Height = Printer.TextHeight("aa")
    
    PrintRect 0, 63, 180, 91    '¿Ü°¢ Æ²
    PrintLine 60, 70, 180, 70   '¼öÆò ¶óÀÎ
    PrintLine 0, 77, 180, 77    '¼öÆò ¶óÀÎ
    PrintLine 60, 84, 180, 84   '¼öÆò ¶óÀÎ
    
    PrintLine 30, 63, 30, 91    '¼öÁ÷ ¶óÀÎ
    PrintLine 60, 63, 60, 91    '¼öÁ÷ ¶óÀÎ
    PrintLine 90, 63, 90, 91    '¼öÁ÷ ¶óÀÎ
    PrintLine 120, 70, 120, 77    '¼öÁ÷ ¶óÀÎ
    PrintLine 150, 70, 150, 77    '¼öÁ÷ ¶óÀÎ
    PrintLine 120, 84, 120, 91    '¼öÁ÷ ¶óÀÎ
    PrintLine 150, 84, 150, 91    '¼öÁ÷ ¶óÀÎ
    
    PrintText 0, 58, "¢Á ±âº»»çÇ×"
    PrintText 6, 68, "´ë ¸® Á¡ ¸í"
    PrintText 69, 65, "ÁÖ    ¼Ò"
    PrintText 69, 72, "¼º    ¸í"
    PrintText 129, 72, "Àü    È­"
    PrintText 6, 82, "¼Ò ºñ ÀÚ ¸í"
    PrintText 69, 79, "ÁÖ    ¼Ò"
    PrintText 69, 86, "Àü    È­"
    PrintText 129, 86, "ÇÚ µå Æù"
    ' °ª Ãâ·Â
    PrintText 35, 68, ´ë¸®Á¡Á¤º¸.´ë¸®Á¡¸í
    PrintText 95, 65, "" 'ÁÖ¼Ò
    PrintText 95, 72, "" '¼º¸í
    PrintText 155, 72, ´ë¸®Á¡Á¤º¸.ÀüÈ­¸ÅÀå
    PrintText 35, 82, SUBRs!¼º¸í
    PrintText 95, 79, SUBRs!ÁÖ¼Ò
    PrintText 95, 86, SUBRs!°í°´ÀüÈ­
    PrintText 155, 86, SUBRs!ÈÞ´ëÆù
    Return
    
' ÇÇÇØ°ü·Ã»çÇ×
PrintDamage:

    Top_Margin = 0: Left_Margin = 0
    Printer.Font.Size = 9
    Printer.DrawWidth = 7
    Text_Height = Printer.TextHeight("a")
    
    PrintRect 0, 105, 180, 140      '¿Ü°¢ Æ²
    PrintLine 0, 112, 180, 112      '¼öÆò ¶óÀÎ
    PrintLine 0, 119, 180, 119      '¼öÆò ¶óÀÎ
    PrintLine 0, 126, 180, 126      '¼öÆò ¶óÀÎ
    PrintLine 0, 133, 180, 133      '¼öÆò ¶óÀÎ
    
    PrintLine 35, 105, 35, 140      '¼öÁ÷ ¶óÀÎ
    PrintLine 85, 105, 85, 140      '¼öÁ÷ ¶óÀÎ
    PrintLine 120, 105, 120, 140      '¼öÁ÷ ¶óÀÎ
    PrintLine 135, 119, 135, 133      '¼öÁ÷ ¶óÀÎ
    PrintLine 150, 119, 150, 133      '¼öÁ÷ ¶óÀÎ
    PrintLine 165, 119, 165, 133      '¼öÁ÷ ¶óÀÎ
    
    PrintText 0, 100, "¢Á ÇÇÇØ°ü·Ã»çÇ×"
    PrintText 10, 107, "Ç°    ¸ñ"
    PrintText 95, 107, "»ó    Ç¥"
    PrintText 10, 114, "±¸ ÀÔ ÀÏ"
    PrintText 95, 114, "»ö    »ó"
    PrintText 10, 121, "±¸ ÀÔ Ã³"
    PrintText 95, 121, "ÅÃ ¹ø È£"
    PrintText 10, 128, "±¸ÀÔÇüÅÂ"
    PrintText 95, 128, "ÀÔ °í ÀÏ"
    PrintText 10, 135, "±¸ÀÔ°¡°Ý"
    PrintText 94, 135, "»ç°íÁ¢¼öÀÏ"
    PrintText 122, 121, "ÃÖÃÊÅÃ"
    PrintText 152, 121, "ÃÖÁ¾ÅÃ"
    PrintText 122, 128, "ÃÖÃÊÀÏ"
    PrintText 152, 128, "ÃÖÁ¾ÀÏ"
    PrintText 20, 145, "¾È³» : ¸ÕÀú ÀúÈñ " & M_CompnyMasterName & "¸¦ ÀÌ¿ëÇØ ÁÖ½Å ¼ÒºñÀÚ²² Áø½ÉÀ¸·Î °¨»çÀÇ ¸»¾¸À» µå¸®¸ç"
    PrintText 20, 149, "       º¸´Ù Á¤È®ÇÑ ÇÇÇØº¸»óÀ» À§ÇÏ¿© °¡´ÉÇÑ »ó¼¼ÇÏ°Ô ÁöÀçÇÏ¿© ÁÖ½Ã±â ¹Ù¶ó¸ç ÇãÀ§"
    PrintText 20, 153, "       ±âÀç´Â ºÒÀÌÀÍÀ» ¹ÞÀ» ¼ö µµ ÀÖ½À´Ï´Ù."
    ' °ª Ãâ·Â
    PrintText 40, 107, SUBRs!Ç°¸í
    PrintText 125, 107, SUBRs!»óÇ¥
    PrintText 40, 114, SUBRs!±¸ÀÔÀÏÀÚ
    PrintText 125, 114, SUBRs!»ö»ó
    PrintText 40, 121, SUBRs!±¸ÀÔÃ³
    PrintText 40, 128, SUBRs!±¸ÀÔÇüÅÂ
    PrintText 40, 135, Format(SUBRs!±¸ÀÔ°¡°Ý, "#,##0")
    PrintText 125, 135, SUBRs!»ç°íÁ¢¼öÀÏ
    PrintText 137, 121, SUBRs!ÃÖÃÊÅÃ¹øÈ£
    PrintText 167, 121, SUBRs!ÃÖÁ¾ÅÃ¹øÈ£
    PrintText 136, 128, SUBRs!ÃÖÃÊÀÔ°íÀÏ
    PrintText 166, 128, SUBRs!ÃÖÁ¾ÀÔ°íÀÏ
    Return

' ´ë¸®Á¡ ±âÀç
AgencyWrite:

    Top_Margin = 0: Left_Margin = 0
    Printer.Font.Size = 9
    Printer.DrawWidth = 7
    Text_Height = Printer.TextHeight("a")
    
    PrintRect 0, 170, 180, 191      '¿Ü°¢ Æ²
    PrintLine 0, 177, 180, 177      '¼öÆò ¶óÀÎ
    PrintLine 0, 184, 180, 184      '¼öÆò ¶óÀÎ
    
    PrintLine 43, 170, 43, 191      '¼öÁ÷ ¶óÀÎ

    PrintText 0, 165, "¢Á ´ë¸®Á¡ ±âÀç"
    PrintText 13, 172, "»ç°íÀÇ Á¾·ù"
    PrintText 13, 179, "»ç°íÀÇ ³»¿ë"
    PrintText 4, 186, "¼ÒºñÀÚÀÇ°ß ¹× ¿ä±¸»çÇ×"
    ' °ª Ãâ·Â
    PrintText 50, 172, SUBRs!»ç°íÁ¾·ù
    PrintText 50, 179, SUBRs!»ç°í³»¿ë
    PrintText 50, 186, SUBRs!»ç°íÀÇ°ß
    Return

' º»»ç ±âÀç
CompanyWrite:

    Top_Margin = 0: Left_Margin = 0
    Printer.Font.Size = 9
    Printer.DrawWidth = 7
    Text_Height = Printer.TextHeight("a")
    
    PrintRect 0, 205, 180, 226      '¿Ü°¢ Æ²
    PrintLine 0, 212, 180, 212      '¼öÆò ¶óÀÎ
    PrintLine 0, 219, 180, 219      '¼öÆò ¶óÀÎ
    
    PrintLine 30, 205, 30, 226      '¼öÁ÷ ¶óÀÎ
    PrintLine 60, 205, 60, 226      '¼öÁ÷ ¶óÀÎ
    PrintLine 90, 205, 90, 226      '¼öÁ÷ ¶óÀÎ
    PrintLine 120, 205, 120, 226      '¼öÁ÷ ¶óÀÎ
    PrintLine 150, 205, 150, 226      '¼öÁ÷ ¶óÀÎ

    PrintText 0, 200, "¢Á º»»ç ±âÀç"
    PrintText 6, 207, "Á¦ Á¶ È¸ »ç"
    PrintText 66, 207, "Àü È­ ¹ø È£"
    PrintText 126, 207, "ÆÇ ¸Å ÀÏ ÀÚ"
    PrintText 6, 214, "Àç °í Çö È²"
    PrintText 66, 214, "´ã       ´ç"
    PrintText 126, 214, "ÆÇ ¸Å ±Ý ¾×"
    PrintText 6, 221, "º¸ »ó ºñ À²"
    PrintText 65, 221, "º¸»ó»êÁ¤±Ý¾×"
    PrintText 126, 221, "ÇÕ ÀÇ ³» ¿ë"
    ' °ª Ãâ·Â
    PrintText 33, 207, "" ' Á¦Á¶È¸»ç
    PrintText 93, 207, "" ' ÀüÈ­¹øÈ£
    PrintText 153, 207, "" 'ÆÇ¸ÅÀÏÀÚ
    PrintText 33, 214, ""  'Àç°íÇöÈ²
    PrintText 93, 214, ""  '´ã´ç
    PrintText 153, 214, "" 'ÆÇ¸Å±Ý¾×
    PrintText 33, 221, "" 'º¸»óºñÀ²
    PrintText 93, 221, "" 'º¸»ó»êÁ¤±Ý¾×
    PrintText 153, 221, "" 'ÇÕÀÇ³»¿ë
    Return

End Function

Public Sub PrintRect(spX As Integer, spY As Integer, epX As Integer, epY As Integer)
        
    ' ¿©¹éÀ» Àû¿ë ½ÃÅ²´Ù.
    spX = spX + Prt_Left + Left_Margin:    spY = spY + Prt_Top + Top_Margin
    epX = epX + Prt_Left + Left_Margin:    epY = epY + Prt_Top + Top_Margin

    Printer.DrawWidth = 6
    Printer.DrawStyle = vbSolid
    Printer.Line (spX, spY)-(epX, epY), , B

End Sub

Public Sub PrintLine(spX As Integer, spY As Integer, epX As Integer, epY As Integer)
        
    ' ¿©¹éÀ» Àû¿ë ½ÃÅ²´Ù.
    spX = spX + Prt_Left + Left_Margin:    spY = spY + Prt_Top + Top_Margin
    epX = epX + Prt_Left + Left_Margin:    epY = epY + Prt_Top + Top_Margin

    Printer.DrawWidth = 6
    Printer.DrawStyle = vbSolid
    Printer.Line (spX, spY)-(epX, epY)

End Sub

Public Sub PrintText(spX As Integer, spY As Integer, msg As String)
        
    ' ¿©¹éÀ» Àû¿ë ½ÃÅ²´Ù.
    spX = spX + Prt_Left + Left_Margin:  spY = spY + Prt_Top + Top_Margin
    
    Printer.CurrentX = spX
    Printer.CurrentY = spY
    Printer.Print msg

End Sub



Public Function Print¹ÌÃâ°íÇöÈ²(ObjRSet As Object, prtNum As Integer, Title As String) As Boolean
    
'    Query = "SELECT DISTINCTROW P.ÀÔ°íÀÏ, (P1.ÀüÈ­1+'-'+ P1.ÀüÈ­2)  AS ÀüÈ­¹øÈ£ , P1.¼º¸í, P.Ç°¸í, "
'    Query = Query & " P.¹øÈ£, P.»ö»ó, P.³»¿ë, P.±Ý¾×, P.»óÅÂ, P.»óÇ¥ "
'    Query = Query & " FROM °í°´Á¤º¸ AS P1, ÀÔÃâ°í AS P "
'    Query = Query & " WHERE (P.ÀÔ°íÀÏ BETWEEN '" & strFromD & "' AND '" & strToD & "') "
'    Query = Query & " AND   (P1.°í°´¹øÈ£ = P.°í°´¹øÈ£ AND P.È®ÀÎ <> 'È®') "
'    Query = Query & " ORDER BY P1.°í°´¹øÈ£, P.ÀÔ°íÀÏ, P.¹øÈ£ "
    
    Dim TotProssCnt As Long
    Dim DefLineSpage As Integer
    Dim DefPointTop     As Integer
    Dim TotalPage   As Long
    
    Print¹ÌÃâ°íÇöÈ² = True
    
    ' ±âº» ÇÁ¸°ÅÍ°¡ ¾øÀ» °æ¿ì
    If Not PrinterCheck Then
        Print¹ÌÃâ°íÇöÈ² = False
        Exit Function
    End If
        
        
    ''''''''''''''''
    On Error GoTo printError
    '''''''''''''''
    
Print_Start:
    Prt_Top = 5
    Prt_Left = 10
    LineCnt = 0
    PageCnt = 1
    TotProssCnt = 0
    PRINT_LINE_COUNT = 45
    DefLineSpage = 5
    DefPointTop = 30
    
    ' ÀüÃ¼ ÆäÀÌÁö ¼ö¸¦ ±¸ÇÑ´Ù.
    TotalPage = Round((ObjRSet.RecordCount / PRINT_LINE_COUNT) + IIf((ObjRSet.RecordCount Mod PRINT_LINE_COUNT) = 0, 0, 0.5))
    
    Printer.ScaleMode = vbMillimeters           ' ÇÁ¸°ÅÍÀÇ ½ºÄÉÀÏ ¸ðµå¸¦ ¹Ð¸®¹ÌÅÍ ´ÜÀ§·Î
    
    
    ' ±âº»»çÇ× ¾ç½Ä ÀÛ¼º
    GoSub PrintDefault
    
    Do Until ObjRSet.EOF
     
        ' ¶óÀÎÀ» Áõ°¡ ½ÃÅ²´Ù.
        LineCnt = LineCnt + 1
        TotProssCnt = TotProssCnt + 1
        
        PrintText 0, (LineCnt * DefLineSpage) + DefPointTop, Format(TotProssCnt, "@@@@")
        PrintText 10, (LineCnt * DefLineSpage) + DefPointTop, Format(ObjRSet.Fields("ÀÔ°íÀÏ"), "@@@@-@@-@@")
        PrintText 30, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("ÀüÈ­¹øÈ£")
        PrintText 50, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("¼º¸í")
        PrintText 70, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("Ç°¸í")
        PrintText 95, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("¹øÈ£")
        PrintText 107, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("»ö»ó")
        PrintText 117, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("³»¿ë")
        PrintText 127, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("±Ý¾×")
        PrintText 140, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("»óÅÂ")
        PrintText 150, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("»óÇ¥")
    
        If PRINT_LINE_COUNT <= LineCnt Then
            
            PageCnt = PageCnt + 1
            PrintLine 0, 260, 180, 260      '¼öÆò ¶óÀÎ
            PrintText 150, 262, "[" & ´ë¸®Á¡Á¤º¸.´ë¸®Á¡¹øÈ£ & "] " & ´ë¸®Á¡Á¤º¸.´ë¸®Á¡¸í
            
'            Printer.EndDoc
'            Exit Function
            Printer.NewPage
            GoSub PrintDefault
            LineCnt = 0
        End If
        ObjRSet.MoveNext
    Loop
    
    PageCnt = PageCnt + 1
    PrintLine 0, 260, 180, 260      '¼öÆò ¶óÀÎ
    PrintText 150, 262, "[" & ´ë¸®Á¡Á¤º¸.´ë¸®Á¡¹øÈ£ & "] " & ´ë¸®Á¡Á¤º¸.´ë¸®Á¡¸í
    
    'Ãâ·Â ÇÑ´Ù.
    ObjRSet.Close
    
    Printer.EndDoc
    
    Exit Function
    
printError:
    Print¹ÌÃâ°íÇöÈ² = False
    Printer.EndDoc
    Exit Function
    
    
' ±âº»»çÇ× ¾ç½Ä ÀÛ¼º
PrintDefault:
    
    Top_Margin = 0: Left_Margin = 0
    
    Printer.FontName = "±¼¸²Ã¼"
    Printer.Font.Bold = True
    Printer.Font.Size = 9
    
    Printer.Font.Size = "18"
    PrintText 55, 8, Title
    Printer.DrawWidth = 12
    PrintLine 50, 16, 125, 16
    
    Printer.Font.Size = 9
    Printer.DrawWidth = 7
    Text_Height = Printer.TextHeight("aa")
    
    PrintLine 0, 25, 180, 25   '¼öÆò ¶óÀÎ
    PrintLine 0, 32, 180, 32    '¼öÆò ¶óÀÎ
    PrintText 160, 20, CStr(PageCnt) & " / " & CStr(TotalPage) & " Page"
    
    
    PrintText 0, 27, "¼ø¹ø"
    PrintText 10, 27, "ÀÔ°íÀÏÀÚ"
    PrintText 30, 27, "ÀüÈ­¹øÈ£"
    PrintText 50, 27, "¼º   ¸í"
    PrintText 70, 27, "Ç°    ¸í"
    PrintText 95, 27, "¹ø  È£"
    PrintText 107, 27, "»ö»ó"
    PrintText 117, 27, "³»¿ë"
    PrintText 127, 27, "±Ý   ¾×"
    PrintText 140, 27, "»óÅÂ"
    PrintText 150, 27, "»ó    Ç¥"
    Return
    
' ÇÇÇØ°ü·Ã»çÇ×
PrintDamage:

    Top_Margin = 0: Left_Margin = 0
    Printer.Font.Size = 9
    Printer.DrawWidth = 7
    Text_Height = Printer.TextHeight("a")
    
    PrintRect 0, 105, 180, 140      '¿Ü°¢ Æ²
    PrintLine 0, 112, 180, 112      '¼öÆò ¶óÀÎ
    PrintLine 0, 119, 180, 119      '¼öÆò ¶óÀÎ
    PrintLine 0, 126, 180, 126      '¼öÆò ¶óÀÎ
    PrintLine 0, 133, 180, 133      '¼öÆò ¶óÀÎ
    
    PrintLine 35, 105, 35, 140      '¼öÁ÷ ¶óÀÎ
    PrintLine 85, 105, 85, 140      '¼öÁ÷ ¶óÀÎ
    PrintLine 120, 105, 120, 140      '¼öÁ÷ ¶óÀÎ
    PrintLine 135, 119, 135, 133      '¼öÁ÷ ¶óÀÎ
    PrintLine 150, 119, 150, 133      '¼öÁ÷ ¶óÀÎ
    PrintLine 165, 119, 165, 133      '¼öÁ÷ ¶óÀÎ

    Return
End Function



Function Print_QN_MM(ByVal KeyCodeTime As String)
' ±âº»¼³Á¤ 25,1,5

    Dim Page_Count As Integer       ' º¸°üÁõ¿¡ Ãâ·ÂµÉ »óÇ°ÀÇ ÃÑ °¹¼ö
    Dim sPage_count As Integer      ' º¸°üÁõÀÇ  ÀüÃ¼ ÆäÀÌÁö¼ö
    Dim Page_Item_Count As Integer  ' ÇÑÆäÀÌÁö¿¡ Ãâ·ÂµÉ »óÇ°ÀÇ °¹¼ö

    Dim dXOffSet As Integer
    Dim dYOffSet As Integer
    
    Dim tmpKEY2 As String
    Dim tmpKEY
    Dim tmpCOD1 '(1 To tmpListCNT)
    Dim tmpAC1 '(1 To tmpListCNT)
    Dim tmpCOD2 '(1 To tmpListCNT)
    Dim tmpAC2 '(1 To tmpListCNT)

    Dim tmpSUSUN '(1 To tmpListCNT)
    Dim tmpCOL  As Long '(1 To tmpListCNT)

    Dim tmpBI1 '(1 To tmpListCNT)
    Dim tmpBIS '(1 To tmpListCNT)

    Dim tmpMON  As Long '(1 To tmpListCNT) As Long
    Dim tmpVAL  As Long
    
    Dim S_Line As Integer
    Dim L_Line As Integer
    Dim GRD_TOT As Integer
    Dim GRD_S_TOT As Integer
    Dim L_Page As Integer
    Dim i As Integer
    Dim j As Integer
    Dim ll As Integer
    Dim SUB_TOT As Integer
    
    ' ±âº» ÇÁ¸°ÅÍ°¡ ¾øÀ» °æ¿ì
    If Not PrinterCheck Then Exit Function
        
   
    ''''''''''''''''
    On Error GoTo printError
    '''''''''''''''
Print_Start:

    
    'CommonDialog1.Action = 5

    ' »ç¿ë °ªµéÀ» ÃÊ±âÈ­ ÇÑ´Ù.
    L_Page = 0
    S_Line = 0
    L_Line = 0
    GRD_TOT = 0
    GRD_S_TOT = 0
    
    Page_Item_Count = GetPrtItemCount("º¸°üÁõ")     ' º¸°üÁõ¿¡ Ãâ·ÂµÉ »óÇ° °¹¼ö
   
    ' À×Å©Á¬ ÇÁ¸°ÅÍ
    If Printer_Gb = "1" Then
        Printer.ScaleMode = vbMillimeters           ' ÇÁ¸°ÅÍÀÇ ½ºÄÉÀÏ ¸ðµå¸¦ ¹Ð¸®¹ÌÅÍ ´ÜÀ§·Î
        Printer.Width = 19 * 567
        Printer.Height = 15 * 567
        Printer.FontName = "±¼¸²Ã¼"
        Printer.Font.Bold = True
        Printer.Font.Size = 9
        Printer.DrawWidth = 1
    
    ' ·¹ÀÌÀú ÇÁ¸°ÅÍ
    ElseIf Printer_Gb = "2" Then
        Printer.ScaleMode = vbMillimeters           ' ÇÁ¸°ÅÍÀÇ ½ºÄÉÀÏ ¸ðµå¸¦ ¹Ð¸®¹ÌÅÍ ´ÜÀ§·Î
        Printer.FontName = "±¼¸²Ã¼"
        Printer.Font.Bold = True
        Printer.Font.Size = 9
        Printer.DrawWidth = 1
    
    End If

    'ÀüÃ¼ Ãâ·Â °¹¼ö¹× Ãâ·Â ³»¿ë º¯¼ö¿¡ ÃÊ±âÈ­
    GoSub Print_Value_Init
    
    If (Page_Count <= 0) Then
        Exit Function
    End If

    ' ÀüÃ¼ Ãâ·Â ÆäÀÌÁö ±¸ÇÏ±â
    If (Page_Count Mod Page_Item_Count) <> 0 Then
        sPage_count = Int(Page_Count / Page_Item_Count) + 1
    Else
        sPage_count = Int(Page_Count / Page_Item_Count)
    End If
    
    'ÀüÃ¼ ÆäÀÌÁö ±îÁö ¹Ýº¹.
    For L_Page = 1 To sPage_count
        ' Ã¹¹øÂ° ÀåÀÌ³ª ¸¶Áö¸· ÀåÀÏ°æ¿ì
        If L_Page = sPage_count Or sPage_count = 1 Then
            S_Line = L_Line + 1
            L_Line = Page_Count   ' frmINPUT.ListView1.ListItems.Count
            'À×Å©Á¬
            GoSub Print_Title
            GoSub Print_Center
            GoSub Print_Bottom
            Printer.EndDoc
            Exit For
        Else
        ' Áß°£ ÆäÀÌÁö ÀÏ °æ¿ì
            S_Line = L_Line + 1
            L_Line = L_Line + Page_Item_Count
            'À×Å©Á¬
            GoSub Print_Title
            GoSub Print_Center
            GoSub Print_Bottom
            Printer.NewPage
        End If
    Next L_Page

    ''''''''''''''''
    'On Error Resume Next
    
    Screen.MousePointer = 0
    
    Exit Function
    
'-------------------------------------------------------------------------------
'--   Ãâ·Â°ª ÃÊ±âÈ­
'-------------------------------------------------------------------------------
Print_Value_Init:
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
    End With
    
    '---------------------------------------------------------
    ' º¸°üÁõ Ãâ·Â »ó´Ü ÀÚ·á ÃÊ±âÈ­
    '---------------------------------------------------------
    Query = "SELECT * FROM º¸°ü¸®½ºÆ® "
    Query = Query & " WHERE KeyCode = '" & KeyCodeTime & "' "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic

    FPrtTop.Code = SUBRs!UserCode
    
    Call Fb°í°´Á¤º¸(FPrtTop.Code)
    
    FPrtTop.HpTel = °í°´Á¤º¸.ÈÞ´ëÆù
    
    FPrtTop.PrtNo = Format(Date, "MMDD") & "-" & SUBRs!InputNumber
    FPrtTop.Tel = °í°´Á¤º¸.ÀüÈ­1 & "-" & °í°´Á¤º¸.ÀüÈ­2
    FPrtTop.Name = SUBRs!InputName
    FPrtTop.Addr = °í°´Á¤º¸.ÁÖ¼Ò
    FPrtTop.Date = Format(Left(SUBRs!InputDate, 8), "@@@@-@@-@@")
    FPrtTop.Date2 = Format(SUBRs!SaleEndDate, "@@@@-@@-@@")
    
    ' º¸°üÁõ Ãâ·Â ÇÏ´Ü ÀÚ·á ÃÊ±âÈ­
    Dim strMaxLng   As String
    
    strMaxLng = "1234567890"
    
    With FPrtBottom
        .Sum = strMaxLng
        RSet .Sum = Format(Val(SUBRs!Price), "#,##0")
        
        .Account0 = strMaxLng
        RSet .Account0 = Format(Val(SUBRs!Price), "#,##0")
        
        .DName = ´ë¸®Á¡Á¤º¸.´ë¸®Á¡¸í
        .DTel = ´ë¸®Á¡Á¤º¸.ÀüÈ­¸ÅÀå
    End With
    SUBRs.Close
    Set SUBRs = Nothing
    
    '-------------------------------------------------------
    ' º¸°üÁõ Ãâ·Â Áß°£ ÀÚ·á ÃÊ±âÈ­
    '-------------------------------------------------------
    Query = "SELECT * FROM º¸°ü»óÇ°¸®½ºÆ® "
    Query = Query & " WHERE KeyCode = '" & KeyCodeTime & "' "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If SUBRs.RecordCount > 0 Then
        SUBRs.MoveLast
        Page_Count = SUBRs.RecordCount
        SUBRs.MoveFirst
    Else
        SUBRs.Close
        Debug.Print "º¸°üÁõ Ãâ·Â ¾øÀ½. (¿À·ù)"
        Return
    End If
    
    If SUBRs.RecordCount > 0 Then
        For i = 1 To 500
            FPArray(i, 1) = SUBRs!Tag
            FPArray(i, 2) = GetGoodsName(SUBRs!GoodsCode)
            FPArray(i, 3) = SUBRs!Color
            FPArray(i, 4) = Format(0, "#,#0")
            FPArray(i, 5) = "º¸°ü¼­ºñ½º"
            FPArray(i, 6) = SUBRs!BrandName
    
            SUBRs.MoveNext
    
            If SUBRs.EOF = True Then
                Exit For
            End If
        Next i
    End If
    SUBRs.Close
    Set SUBRs = Nothing
    
    Return
'-------------------------------------------------------------------------------
'--   Å¸ÀÌÆ² ºÎºÐ Ãâ·Â
'-------------------------------------------------------------------------------
Print_Title:
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '  ´ë¸®Á¡ º¸°ü¿ë

    For j = 0 To 1
        If j = 1 Then
            PrtPoint2 = GetPrtPointMM("¼Õ´Ô¿ë")
        Else
            PrtPoint2.x = 0
            PrtPoint2.y = 0
        End If
        
        PrtPoint4 = GetPrtPointMM("¿©¹é")                ' ¼³Á¤ÇÑ ¿©¹éÀ» °¡Áö°í ¿Â´Ù.
        
        ' ÀüÇ¥ ¹øÈ£
        If Printer_BO_Gb = "0" Then
            PrtPoint = GetPrtPointMM("PRTNO")
            SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
            Printer.Print FPrtTop.PrtNo
        End If
        If Printer_BO_Gb = "1" Then
            PrtPoint = GetPrtPointMM("HPTEL")
            SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
            Printer.Print FPrtTop.HpTel
        End If
        ' °í°´ ÀüÈ­¹øÈ£
        PrtPoint = GetPrtPointMM("GTEL")
        SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
        Printer.Print FPrtTop.Tel
        ' °í°´ ¼º¸í
        PrtPoint = GetPrtPointMM("GNAME")
        SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
        Printer.Print FPrtTop.Name
        ' ÁÖ¼Ò (¼Õ´Ô)
        PrtPoint = GetPrtPointMM("ADDR")
        SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
        Printer.Print FPrtTop.Addr
        ' Á¢¼öÀÏ
        PrtPoint = GetPrtPointMM("DATE")
        SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
        Printer.Print FPrtTop.Date
        ' °í°´ ¹øÈ£
        PrtPoint = GetPrtPointMM("CODE")
        SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
        Printer.Print FPrtTop.Code
        ' ÀÎµµ ¿¬µµ
        PrtPoint = GetPrtPointMM("DATE2")
        SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
        Printer.Print FPrtTop.Date2
    Next j
    Return
'-------------------------------------------------------------------------------
'--   ³»¿ë ºÎºÐ Ãâ·Â
'-------------------------------------------------------------------------------
Print_Center:
    
    
    ll = 0 ' º¸°üÁõ Ãâ·Â ¶óÀÎ ÃÊ±âÈ­
    If (S_Line + Page_Item_Count) > Page_Count Then
        SUB_TOT = Page_Count
    Else
        SUB_TOT = S_Line + Page_Item_Count - 1
    End If
    
    ' ±âº» ¶óÀÎ´ç °£°ÝÀ» °¡Àú¿Â´Ù
    PrtPoint3 = GetPrtPoint("NEXT_LINE")
    PrtPoint4 = GetPrtPoint("¿©¹é")
    For i = S_Line To SUB_TOT
        ll = ll + 1
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' º¸°ü¿ë
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Ã¹ÁÙÀº Áõ°¡ ÇÏÁö ¾Ê´Â´Ù
        If (ll - 1) Then
            If (i Mod 2) Then
                PrtPoint4.y = PrtPoint4.y + PrtPoint3.y + 1
            Else
                PrtPoint4.y = PrtPoint4.y + PrtPoint3.y
            End If
        End If
        
        For j = 0 To 1
            If j = 1 Then
                PrtPoint2 = GetPrtPointMM("¼Õ´Ô¿ë")
            Else
                PrtPoint2.x = 0
                PrtPoint2.y = 0
            End If
            
        
            'ÅÃ¹øÈ£
            PrtPoint = GetPrtPointMM("TAGNUM")
            SetPrtPoint PrtPoint2, PrtPoint, PrtPoint4
            SetPrtPoint PrtPoint2, PrtPoint, PrtPoint4
            Printer.Print FPArray(i, 1)
            
            'Ç°¸í
            PrtPoint = GetPrtPointMM("PNAME")
            SetPrtPoint PrtPoint2, PrtPoint, PrtPoint4
            Printer.Print FPArray(i, 2)
            
            '»ö»ó
            PrtPoint = GetPrtPointMM("PCOLOR")
            SetPrtPoint PrtPoint2, PrtPoint, PrtPoint4
            Printer.Print FPArray(i, 3)
            
            '±Ý¾×
            PrtPoint = GetPrtPointMM("PACCOUNT")
            SetPrtPoint PrtPoint2, PrtPoint, PrtPoint4
            Printer.Print FPArray(i, 4)
            
            '³»¿ë
            PrtPoint = GetPrtPointMM("PTEMP")
            SetPrtPoint PrtPoint2, PrtPoint, PrtPoint4
            Printer.Print FPArray(i, 5)
            
            '»óÇ¥
            PrtPoint = GetPrtPointMM("BRAND")
            SetPrtPoint PrtPoint2, PrtPoint, PrtPoint4
            Printer.Print FPArray(i, 6)
        Next j
    Next i
    Return

'-------------------------------------------------------------------------------
'--   ³¡ ºÎºÐ Ãâ·Â
'-------------------------------------------------------------------------------
Print_Bottom:
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' º¸°ü¿ë
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    For j = 0 To 1
        If j = 1 Then
            PrtPoint2 = GetPrtPointMM("¼Õ´Ô¿ë")
        Else
            PrtPoint2.x = 0
            PrtPoint2.y = 0
        End If
        
        PrtPoint4 = GetPrtPointMM("¿©¹é")                ' ¼³Á¤ÇÑ ¿©¹éÀ» °¡Áö°í ¿Â´Ù.
        ' ¸¶Áö¸· ÀåÀÏ°æ¿ì ÀüÃ¼ ÇÕ°è¹× ±Ý¾× Ãâ·Â
        If L_Page = sPage_count Or sPage_count = 1 Then
            ' Á¡¼ö
            PrtPoint = GetPrtPointMM("SUM")
            SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
            Printer.Print FPrtBottom.Sum
            '±Ý¾×
            PrtPoint = GetPrtPointMM("ACCOUNT0")
            SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
            Printer.Print FPrtBottom.Account0
            ' ¼ö·É¾×
            PrtPoint = GetPrtPointMM("ACCOUNT1")
            SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
            Printer.Print FPrtBottom.Account1
            'ÀÜ¾×
            PrtPoint = GetPrtPointMM("ACCOUNT2")
            SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
            Printer.Print FPrtBottom.Account2
        
            '¸¶ÀÏ¸®Áö
            If Val(FPrtBottom.MilMoney) > 0 Then
                PrtPoint = GetPrtPointMM("MILEAGE")
                SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
                'Printer.Print "¸¶ÀÏ¸®ÁöÀÜ¾× : " & FPrtBottom.MilMoney
                Printer.Print FPrtBottom.MilMoney
            End If
            
            If Printer_BO_Gb = "1" Then
                ' ÀüÀÏ ¹Ì¼ö
                PrtPoint = GetPrtPointMM("OLDMISU")
                SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
                Printer.Print FPrtBottom.OldDayMisu
                ' ¹Ì¼ö ÇÕ°è
                PrtPoint = GetPrtPointMM("MISUMONEY")
                SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
                Printer.Print FPrtBottom.MiSuTotal
                ' ¼ö±Ý¾×
                PrtPoint = GetPrtPointMM("SUGUMONEY")
                SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
                Printer.Print FPrtBottom.SuGumMonye
                ' »ç¿ë¸¶ÀÏ¸®Áö
                PrtPoint = GetPrtPointMM("USERMILEAGE")
                SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
                Printer.Print FPrtBottom.MilUser
                
            If Val(FPrtBottom.MilMoney) > 0 Then
                    ' ¸¶ÀÏ¸®Áö ÀÜ¾×
                    PrtPoint = GetPrtPointMM("MILEAGE")
                    SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
                    Printer.Print FPrtBottom.MilMoney
            End If
            
            If ´ë¸®Á¡Á¤º¸.¸¶ÀÏ¸®Áö¿©ºÎ = "Y" Then
                ' ´©Àû ¸¶ÀÏ¸®Áö
                PrtPoint = GetPrtPointMM("ADDMILEAGE")
                SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
                Printer.Print FPrtBottom.MilAddMoney
            End If
            
                ' º¸°üÁõ ¿À·ù ¼öÁ¤
                PrtPoint = GetPrtPointMM("ADDMILEAGETITLE")
                SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
                Printer.Print "Àû¸³"
            End If
            
        End If
        ' ´ë¸®Á¡¸í
        PrtPoint = GetPrtPointMM("DNAME")
        SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
        Printer.Print FPrtBottom.DName
        ' ´ë¸®Á¡ ÀüÈ­¹øÈ£
        PrtPoint = GetPrtPointMM("DTEL")
        SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
        Printer.Print FPrtBottom.DTel
        ' ÆäÀÌÁö/ÀüÃ¼ ÆäÀÌÁö
        PrtPoint = GetPrtPointMM("PAGE")
        SetPrtPoint PrtPoint, PrtPoint2, PrtPoint4
        Printer.Print L_Page & "/" & sPage_count
        
    Next j
Return
'-------------------------------------------------------------------------------
'--   ÀÎ¼âÁß ¿À·ù ½ÇÇà ºÎºÐ
'-------------------------------------------------------------------------------
printError:
    MsgBox Err.Description & Space(10), vbCritical
    'MsgBox " ÇÁ¸°ÅÍ¸¦ È®ÀÎÇØ ÁÖ½Ê½Ã¿ä ! " & VBA.Err.Number, vbCritical, "Ãâ·Â¿À·ù¹ß»ý"
    Resume
End Function


Public Function GetGoodsName(ByVal Scode As String) As String
    GetGoodsName = ""
    
    Query = "SELECT Ç°¸í FROM ÂüÁ¶ÄÚµå "
    Query = Query & " WHERE ±¸ºÐÄÚµå = '" & Scode & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If Rs.RecordCount = 1 Then
        GetGoodsName = Rs!Ç°¸í & ""
    End If
    
    Rs.Close
    Set Rs = Nothing
End Function

'--------------------------------------------------------------------------------------------------------------
' Procedure : QN_Day_Info
' DateTime  : 2006-11-07 12:59
' Author    : pds2004
' Purpose   : Àü´ÞµÈ ÀÏÀÚÀÇ ÀüÃ¼ º¸°ü ¼­ºñ½º ÇÕ°è ±Ý¾×À» ¸®ÅÏÇÑ´Ù.
'--------------------------------------------------------------------------------------------------------------
Private Function QN_Day_Info(ByVal sDate As String, ByRef dCleanTotal As Double, ByRef dBonSaTotal As Double, ByRef dStoreTotal As Double, ByRef dMasterTotal As Double) As Double
    On Error GoTo QN_Day_Info_Error

    ' ÃÊ±âÈ­
    dCleanTotal = 0:    dBonSaTotal = 0: dStoreTotal = 0

    Query = " SELECT SUM(Price) AS TotalPrice FROM º¸°ü¸®½ºÆ® "
    Query = Query & " WHERE LEFT(InputDate,8) = '" & sDate & "' "
    Set Rs = New ADODB.Recordset
    Rs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If Rs.RecordCount <= 0 Then
        Rs.Close
        Set Rs = Nothing
        
        Exit Function
    End If
    
    If IsNull(Rs.Fields("TotalPrice")) = True Then
        QN_Day_Info = 0
    Else
        QN_Day_Info = CDbl(Rs.Fields("TotalPrice"))
    End If
    Rs.Close
    Set Rs = Nothing
    
    dMasterTotal = (QN_Day_Info * 0.8)
    dCleanTotal = (QN_Day_Info * 0.2)
    dBonSaTotal = (dCleanTotal * 0.5)
    dStoreTotal = (dCleanTotal * 0.5)

    On Error GoTo 0
    
    Exit Function

QN_Day_Info_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure QN_Day_Info of Module Printer1"

End Function

Public Function GetMileageMoneyToPoint(ByVal nMoney As Double) As Double
    Dim nTemp   As Double

    If ´ë¸®Á¡Á¤º¸.¸¶ÀÏ¸®ÁöÁõ°¡±¸ºÐ = "0" Then
    ' 100,000¿ø = 3,000¿ø   200,000¿ø = 4,000¿ø    300,000¿ø = 5,000¿ø
    ' 400,000¿ø = 6,000¿ø   500,000¿ø = 7,000¿ø
        
        nTemp = nMoney - ((nMoney \ NextMileage) * NextMileage)
        
        ' ´ÙÀ½ ¹ß»ýÇÒ ±Ý¾×¿¡ ÇØ´çÇÏ´Â ³»¿ëÀ» ±¸ÇÏ±â ¶§¹®
        If nMoney > 400000 Then
            GetMileageMoneyToPoint = (nTemp * (7000 / NextMileage))
            Exit Function
            
        ElseIf nMoney > 300000 Then
            GetMileageMoneyToPoint = (nTemp * (6000 / NextMileage))
            Exit Function
            
        ElseIf nMoney > 200000 Then
            GetMileageMoneyToPoint = (nTemp * (5000 / NextMileage))
            Exit Function
            
        ElseIf nMoney > 100000 Then
            GetMileageMoneyToPoint = (nTemp * (4000 / NextMileage))
            Exit Function
            
        ElseIf nMoney < 100000 Then
            GetMileageMoneyToPoint = (nTemp * (3000 / NextMileage))
            Exit Function
            
        End If
        
    
    
    ElseIf ´ë¸®Á¡Á¤º¸.¸¶ÀÏ¸®ÁöÁõ°¡±¸ºÐ = "1" Then
    ' 100,000¿ø ´ÜÀ§·Î ¸Å¹ø 3,000¿ø¾¿ Áõ°¡ÇÑ´Ù.
        nTemp = nMoney - ((nMoney \ NextMileage) * NextMileage)
        GetMileageMoneyToPoint = (nTemp * (3000 / NextMileage))
        Exit Function
    
    End If


End Function



Public Function Print´ëºÐ·ùÇöÈ²(ObjRSet As Object, prtNum As Integer, Title As String) As Boolean
    
'    Query = "SELECT DISTINCTROW P.ÀÔ°íÀÏ, M.ÈÞ´ëÆù, (M.ÀüÈ­1+'-'+ M.ÀüÈ­2)  AS ÀüÈ­¹øÈ£ , M.¼º¸í, P.Ç°¸í, "
'    Query = Query & " P.¹øÈ£, P.»ö»ó, P.³»¿ë, P.±Ý¾×, P.»óÅÂ, P.»óÇ¥ "
'    Query = Query & " FROM °í°´Á¤º¸ AS M, ÀÔÃâ°í AS P "
'    Query = Query & " WHERE (P.ÀÔ°íÀÏ BETWEEN '" & Format(DTPicker1(0).Value, "yyyyMMdd") & "' "
'    Query = Query & " AND '" & Format(DTPicker1(1).Value, "yyyyMMdd") & "') "
'    Query = Query & " AND   (M.°í°´¹øÈ£ = P.°í°´¹øÈ£ AND P.È®ÀÎ <> 'È®') "
'    Query = Query & " AND   (P.ÆÇ¸ÅÃë¼Ò <> 'Y') "
'    Query = Query & " AND   (P.ÄÚµå LIKE '" & Left(cboGroup.Text, 1) & "%') "
'    Query = Query & " ORDER BY P.ÀÔ°íÀÏ,  P.¹øÈ£ "
    
    Dim TotProssCnt As Long
    Dim DefLineSpage As Integer
    Dim DefPointTop     As Integer
    Dim TotalPage   As Long
    
    Print´ëºÐ·ùÇöÈ² = True
    
    ' ±âº» ÇÁ¸°ÅÍ°¡ ¾øÀ» °æ¿ì
    If Not PrinterCheck Then
        Print´ëºÐ·ùÇöÈ² = False
        Exit Function
    End If
        
        
    ''''''''''''''''
    On Error GoTo printError
    '''''''''''''''
    
Print_Start:
    Prt_Top = 5
    Prt_Left = 10
    LineCnt = 0
    PageCnt = 1
    TotProssCnt = 0
    PRINT_LINE_COUNT = 45
    DefLineSpage = 5
    DefPointTop = 30
    
    ' ÀüÃ¼ ÆäÀÌÁö ¼ö¸¦ ±¸ÇÑ´Ù.
    TotalPage = Round((ObjRSet.RecordCount / PRINT_LINE_COUNT) + IIf((ObjRSet.RecordCount Mod PRINT_LINE_COUNT) = 0, 0, 0.5))
    
    Printer.ScaleMode = vbMillimeters           ' ÇÁ¸°ÅÍÀÇ ½ºÄÉÀÏ ¸ðµå¸¦ ¹Ð¸®¹ÌÅÍ ´ÜÀ§·Î
    
    
    ' ±âº»»çÇ× ¾ç½Ä ÀÛ¼º
    GoSub PrintDefault
    
    Do Until ObjRSet.EOF
     
        ' ¶óÀÎÀ» Áõ°¡ ½ÃÅ²´Ù.
        LineCnt = LineCnt + 1
        TotProssCnt = TotProssCnt + 1
        
        PrintText 0, (LineCnt * DefLineSpage) + DefPointTop, Format(TotProssCnt, "@@@@")
        PrintText 10, (LineCnt * DefLineSpage) + DefPointTop, Format(ObjRSet.Fields("ÀÔ°íÀÏ"), "@@@@-@@-@@")
        PrintText 30, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("ÀüÈ­¹øÈ£")
        PrintText 50, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("¼º¸í")
        PrintText 70, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("Ç°¸í")
        PrintText 95, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("¹øÈ£")
        PrintText 107, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("»ö»ó")
        PrintText 117, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("³»¿ë")
        PrintText 127, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("±Ý¾×")
        PrintText 140, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("»óÅÂ")
        PrintText 150, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("»óÇ¥")
    
        If PRINT_LINE_COUNT <= LineCnt Then
            
            PageCnt = PageCnt + 1
            PrintLine 0, 260, 180, 260      '¼öÆò ¶óÀÎ
            PrintText 150, 262, "[" & ´ë¸®Á¡Á¤º¸.´ë¸®Á¡¹øÈ£ & "] " & ´ë¸®Á¡Á¤º¸.´ë¸®Á¡¸í
            
'            Printer.EndDoc
'            Exit Function
            Printer.NewPage
            GoSub PrintDefault
            LineCnt = 0
        End If
        ObjRSet.MoveNext
    Loop
    
    PageCnt = PageCnt + 1
    PrintLine 0, 260, 180, 260      '¼öÆò ¶óÀÎ
    PrintText 150, 262, "[" & ´ë¸®Á¡Á¤º¸.´ë¸®Á¡¹øÈ£ & "] " & ´ë¸®Á¡Á¤º¸.´ë¸®Á¡¸í
    
    'Ãâ·Â ÇÑ´Ù.
    ObjRSet.Close
    Printer.EndDoc
    Exit Function
    
printError:
    Print´ëºÐ·ùÇöÈ² = False
    Printer.EndDoc
    Exit Function
    
    
' ±âº»»çÇ× ¾ç½Ä ÀÛ¼º
PrintDefault:
    
    Top_Margin = 0: Left_Margin = 0
    
    Printer.FontName = "±¼¸²Ã¼"
    Printer.Font.Bold = True
    Printer.Font.Size = 9
    
    Printer.Font.Size = "18"
    PrintText 55, 8, Title
    Printer.DrawWidth = 12
    PrintLine 50, 16, 125, 16
    
    Printer.Font.Size = 9
    Printer.DrawWidth = 7
    Text_Height = Printer.TextHeight("aa")
    
    PrintLine 0, 25, 180, 25   '¼öÆò ¶óÀÎ
    PrintLine 0, 32, 180, 32    '¼öÆò ¶óÀÎ
    PrintText 160, 20, CStr(PageCnt) & " / " & CStr(TotalPage) & " Page"
    
    
    PrintText 0, 27, "¼ø¹ø"
    PrintText 10, 27, "ÀÔ°íÀÏÀÚ"
    PrintText 30, 27, "ÀüÈ­¹øÈ£"
    PrintText 50, 27, "¼º   ¸í"
    PrintText 70, 27, "Ç°    ¸í"
    PrintText 95, 27, "¹ø  È£"
    PrintText 107, 27, "»ö»ó"
    PrintText 117, 27, "³»¿ë"
    PrintText 127, 27, "±Ý   ¾×"
    PrintText 140, 27, "»óÅÂ"
    PrintText 150, 27, "»ó    Ç¥"
    Return
    
' ÇÇÇØ°ü·Ã»çÇ×
PrintDamage:

    Top_Margin = 0: Left_Margin = 0
    Printer.Font.Size = 9
    Printer.DrawWidth = 7
    Text_Height = Printer.TextHeight("a")
    
    PrintRect 0, 105, 180, 140      '¿Ü°¢ Æ²
    PrintLine 0, 112, 180, 112      '¼öÆò ¶óÀÎ
    PrintLine 0, 119, 180, 119      '¼öÆò ¶óÀÎ
    PrintLine 0, 126, 180, 126      '¼öÆò ¶óÀÎ
    PrintLine 0, 133, 180, 133      '¼öÆò ¶óÀÎ
    
    PrintLine 35, 105, 35, 140      '¼öÁ÷ ¶óÀÎ
    PrintLine 85, 105, 85, 140      '¼öÁ÷ ¶óÀÎ
    PrintLine 120, 105, 120, 140      '¼öÁ÷ ¶óÀÎ
    PrintLine 135, 119, 135, 133      '¼öÁ÷ ¶óÀÎ
    PrintLine 150, 119, 150, 133      '¼öÁ÷ ¶óÀÎ
    PrintLine 165, 119, 165, 133      '¼öÁ÷ ¶óÀÎ


    Return



End Function



Public Function Print´ëºÒ·ùÇöÈ²¿ÜÁÖ(ObjRSet As Object, prtNum As Integer, Title As String) As Boolean
    
'    Query = "SELECT DISTINCTROW P.ÀÔ°íÀÏ, M.ÈÞ´ëÆù, (M.ÀüÈ­1+'-'+ M.ÀüÈ­2)  AS ÀüÈ­¹øÈ£ , M.¼º¸í, P.Ç°¸í, "
'    Query = Query & " P.¹øÈ£, P.»ö»ó, P.³»¿ë, P.±Ý¾×, P.»óÅÂ, P.»óÇ¥ "
'    Query = Query & " FROM °í°´Á¤º¸ AS M, ÀÔÃâ°í AS P "
'    Query = Query & " WHERE (P.ÀÔ°íÀÏ BETWEEN '" & Format(DTPicker1(0).Value, "yyyyMMdd") & "' "
'    Query = Query & " AND '" & Format(DTPicker1(1).Value, "yyyyMMdd") & "') "
'    Query = Query & " AND   (M.°í°´¹øÈ£ = P.°í°´¹øÈ£ AND P.È®ÀÎ <> 'È®') "
'    Query = Query & " AND   (P.ÆÇ¸ÅÃë¼Ò <> 'Y') "
'    Query = Query & " AND   (P.ÄÚµå LIKE '" & Left(cboGroup.Text, 1) & "%') "
'    Query = Query & " ORDER BY P.ÀÔ°íÀÏ,  P.¹øÈ£ "
    
    Dim TotProssCnt As Long
    Dim DefLineSpage As Integer
    Dim DefPointTop     As Integer
    Dim TotalPage   As Long
    Dim nMoneyTotal As Long
    
    Print´ëºÒ·ùÇöÈ²¿ÜÁÖ = True
    
    ' ±âº» ÇÁ¸°ÅÍ°¡ ¾øÀ» °æ¿ì
    If Not PrinterCheck Then
        Print´ëºÒ·ùÇöÈ²¿ÜÁÖ = False
        Exit Function
    End If
        
        
    ''''''''''''''''
    On Error GoTo printError
    '''''''''''''''
    
Print_Start:

    nMoneyTotal = 0
    
    Prt_Top = 5
    Prt_Left = 10
    LineCnt = 0
    PageCnt = 1
    TotProssCnt = 0
    PRINT_LINE_COUNT = 45
    DefLineSpage = 5
    DefPointTop = 30
    
    ' ÀüÃ¼ ÆäÀÌÁö ¼ö¸¦ ±¸ÇÑ´Ù.
    TotalPage = Round((ObjRSet.RecordCount / PRINT_LINE_COUNT) + IIf((ObjRSet.RecordCount Mod PRINT_LINE_COUNT) = 0, 0, 0.5))
    
    Printer.ScaleMode = vbMillimeters           ' ÇÁ¸°ÅÍÀÇ ½ºÄÉÀÏ ¸ðµå¸¦ ¹Ð¸®¹ÌÅÍ ´ÜÀ§·Î
    
    
    ' ±âº»»çÇ× ¾ç½Ä ÀÛ¼º
    GoSub PrintDefault
    
    Do Until ObjRSet.EOF
     
        ' ¶óÀÎÀ» Áõ°¡ ½ÃÅ²´Ù.
        LineCnt = LineCnt + 1
        TotProssCnt = TotProssCnt + 1
        
'        PrintText 0, (LineCnt * DefLineSpage) + DefPointTop, Format(TotProssCnt, "@@@@")
'        PrintText 10, (LineCnt * DefLineSpage) + DefPointTop, Format(ObjRSet.Fields("ÀÔ°íÀÏ"), "@@@@-@@-@@")
'        PrintText 30, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("ÀüÈ­¹øÈ£")
'        PrintText 50, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("¼º¸í")
'        PrintText 70, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("Ç°¸í")
'        PrintText 95, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("¹øÈ£")
'        PrintText 107, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("»ö»ó")
'        PrintText 117, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("³»¿ë")
'        PrintText 127, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("±Ý¾×")
'        PrintText 140, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("»óÅÂ")
'        PrintText 150, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("»óÇ¥")
    
        PrintText 0, (LineCnt * DefLineSpage) + DefPointTop, Format(TotProssCnt, "@@@@")
        PrintText 10, (LineCnt * DefLineSpage) + DefPointTop, Format(ObjRSet.Fields("ÀÔ°íÀÏ"), "@@@@-@@-@@")
        PrintText 30, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("¹øÈ£")
        PrintText 45, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("Ç°¸í")
        PrintText 85, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("»ö»ó")
        PrintText 105, (LineCnt * DefLineSpage) + DefPointTop, Format(ObjRSet.Fields("±Ý¾×"), "#,##0")
        PrintText 120, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("»óÇ¥")
    
        nMoneyTotal = nMoneyTotal + Val(ObjRSet.Fields("±Ý¾×"))
        
        If PRINT_LINE_COUNT <= LineCnt Then
            
            PageCnt = PageCnt + 1
            PrintLine 0, 260, 180, 260      '¼öÆò ¶óÀÎ
            PrintText 150, 262, "[" & ´ë¸®Á¡Á¤º¸.´ë¸®Á¡¹øÈ£ & "] " & ´ë¸®Á¡Á¤º¸.´ë¸®Á¡¸í
            
'            Printer.EndDoc
'            Exit Function
            Printer.NewPage
            GoSub PrintDefault
            LineCnt = 0
        End If
        ObjRSet.MoveNext
    Loop
    
    LineCnt = LineCnt + 1
    PrintLine 0, (LineCnt * DefLineSpage) + DefPointTop, 180, (LineCnt * DefLineSpage) + DefPointTop      '¼öÆò ¶óÀÎ
    
    LineCnt = LineCnt + 1
    PrintText 0, (LineCnt * DefLineSpage) + DefPointTop, "ÇÕ°è±Ý¾×: " & Format(nMoneyTotal, "#,##0") & "¿ø"
    PrintText 45, (LineCnt * DefLineSpage) + DefPointTop, "¸ÅÀå±Ý¾×: " & Format(nMoneyTotal * (1 - (Val(´ë¸®Á¡Á¤º¸.¿ÜÁÖ¿îµ¿È­¸¶Áø) / 100)), "#,##0") & "¿ø"
    PrintText 85, (LineCnt * DefLineSpage) + DefPointTop, "¿ÜÁÖ±Ý¾×: " & Format(nMoneyTotal * (Val(´ë¸®Á¡Á¤º¸.¿ÜÁÖ¿îµ¿È­¸¶Áø) / 100), "#,##0") & "¿ø"
    
  
    
    PageCnt = PageCnt + 1
    PrintLine 0, 260, 180, 260      '¼öÆò ¶óÀÎ
    PrintText 150, 262, "[" & ´ë¸®Á¡Á¤º¸.´ë¸®Á¡¹øÈ£ & "] " & ´ë¸®Á¡Á¤º¸.´ë¸®Á¡¸í
    
    'Ãâ·Â ÇÑ´Ù.
    ObjRSet.Close
    Printer.EndDoc
    Exit Function
    
printError:
    Print´ëºÒ·ùÇöÈ²¿ÜÁÖ = False
    Printer.EndDoc
    Exit Function
    
    
' ±âº»»çÇ× ¾ç½Ä ÀÛ¼º
PrintDefault:
    
    Top_Margin = 0: Left_Margin = 0
    
    Printer.FontName = "±¼¸²Ã¼"
    Printer.Font.Bold = True
    Printer.Font.Size = 9
    
    Printer.Font.Size = "18"
    PrintText 55, 8, Title
    Printer.DrawWidth = 12
    PrintLine 50, 16, 125, 16
    
    Printer.Font.Size = 9
    Printer.DrawWidth = 7
    Text_Height = Printer.TextHeight("aa")
    
    PrintLine 0, 25, 180, 25   '¼öÆò ¶óÀÎ
    PrintLine 0, 32, 180, 32    '¼öÆò ¶óÀÎ
    
    
    PrintText 0, 20, "[" & ´ë¸®Á¡Á¤º¸.´ë¸®Á¡¹øÈ£ & "] " & ´ë¸®Á¡Á¤º¸.´ë¸®Á¡¸í
    PrintText 160, 20, CStr(PageCnt) & " / " & CStr(TotalPage) & " Page"
    
    
    PrintText 0, 27, "¼ø¹ø"
    PrintText 10, 27, "ÀÔ°íÀÏÀÚ"
    PrintText 30, 27, "¹ø  È£"
    PrintText 45, 27, "Ç°    ¸í"
    PrintText 85, 27, "»ö»ó"
    PrintText 105, 27, "±Ý¾×"
    PrintText 120, 27, "»ó    Ç¥"
    PrintText 150, 27, "ºñ    °í"
    Return
    
' ÇÇÇØ°ü·Ã»çÇ×
PrintDamage:

    Top_Margin = 0: Left_Margin = 0
    Printer.Font.Size = 9
    Printer.DrawWidth = 7
    Text_Height = Printer.TextHeight("a")
    
    PrintRect 0, 105, 180, 140      '¿Ü°¢ Æ²
    PrintLine 0, 112, 180, 112      '¼öÆò ¶óÀÎ
    PrintLine 0, 119, 180, 119      '¼öÆò ¶óÀÎ
    PrintLine 0, 126, 180, 126      '¼öÆò ¶óÀÎ
    PrintLine 0, 133, 180, 133      '¼öÆò ¶óÀÎ
    
    PrintLine 35, 105, 35, 140      '¼öÁ÷ ¶óÀÎ
    PrintLine 85, 105, 85, 140      '¼öÁ÷ ¶óÀÎ
    PrintLine 120, 105, 120, 140      '¼öÁ÷ ¶óÀÎ
    PrintLine 135, 119, 135, 133      '¼öÁ÷ ¶óÀÎ
    PrintLine 150, 119, 150, 133      '¼öÁ÷ ¶óÀÎ
    PrintLine 165, 119, 165, 133      '¼öÁ÷ ¶óÀÎ

    Return
End Function

Public Function Print¼¼Å¹ºñÈ¯ºÒÇöÈ²(ObjRSet As Object, prtNum As Integer, Title As String, sDateFrom As String, sDateTo As String) As Boolean
'    Query = "SELECT DISTINCTROW P.¼¼Å¹ºñÈ¯ºÒÀÏÀÚ, P.ÀÔ°íÀÏ, M.ÈÞ´ëÆù, (M.ÀüÈ­1+'-'+ M.ÀüÈ­2)  AS ÀüÈ­¹øÈ£ , M.¼º¸í, P.Ç°¸í, "
'    Query = Query & " P.¹øÈ£, P.»ö»ó, P.³»¿ë, P.±Ý¾×, P.»óÅÂ, P.»óÇ¥ "
'    Query = Query & " FROM °í°´Á¤º¸ AS M, ÀÔÃâ°í AS P "
'    Query = Query & " WHERE (LEFT(P.¼¼Å¹ºñÈ¯ºÒÀÏÀÚ,8) BETWEEN '" & Format(DTPicker1(0).Value, "yyyyMMdd") & "' "
'    Query = Query & " AND '" & Format(DTPicker1(1).Value, "yyyyMMdd") & "') "
'    Query = Query & " AND   (M.°í°´¹øÈ£ = P.°í°´¹øÈ£ ) "
'    Query = Query & " AND   (P.¼¼Å¹ºñÈ¯ºÒÀÏÀÚ <> '' ) "
'    Query = Query & " ORDER BY P.ÀÔ°íÀÏ, M.¼º¸í, P.¹øÈ£ "
    
    Dim TotProssCnt As Long
    Dim DefLineSpage As Integer
    Dim DefPointTop     As Integer
    Dim TotalPage   As Long
    
    Dim strMaxLng   As String
    Dim nTotalMoney As Double
    
    Print¼¼Å¹ºñÈ¯ºÒÇöÈ² = True
    
    ' ±âº» ÇÁ¸°ÅÍ°¡ ¾øÀ» °æ¿ì
    If Not PrinterCheck Then
        Print¼¼Å¹ºñÈ¯ºÒÇöÈ² = False
        Exit Function
    End If
        
        
    ''''''''''''''''
    On Error GoTo printError
    '''''''''''''''
    
Print_Start:
    Prt_Top = 5
    Prt_Left = 10
    LineCnt = 0
    PageCnt = 1
    TotProssCnt = 0
    PRINT_LINE_COUNT = 45
    DefLineSpage = 5
    DefPointTop = 30
    
    nTotalMoney = 0
    
    ' ÀüÃ¼ ÆäÀÌÁö ¼ö¸¦ ±¸ÇÑ´Ù.
    TotalPage = Round((ObjRSet.RecordCount / PRINT_LINE_COUNT) + IIf((ObjRSet.RecordCount Mod PRINT_LINE_COUNT) = 0, 0, 0.5))
    
    Printer.ScaleMode = vbMillimeters           ' ÇÁ¸°ÅÍÀÇ ½ºÄÉÀÏ ¸ðµå¸¦ ¹Ð¸®¹ÌÅÍ ´ÜÀ§·Î
    
    
    ' ±âº»»çÇ× ¾ç½Ä ÀÛ¼º
    GoSub PrintDefault
    
    Do Until ObjRSet.EOF
        ' ¶óÀÎÀ» Áõ°¡ ½ÃÅ²´Ù.
        LineCnt = LineCnt + 1
        TotProssCnt = TotProssCnt + 1

        PrintText 0, (LineCnt * DefLineSpage) + DefPointTop, Format(TotProssCnt, "@@@@")
        PrintText 10, (LineCnt * DefLineSpage) + DefPointTop, Format(Left(ObjRSet.Fields("¼¼Å¹ºñÈ¯ºÒÀÏÀÚ"), 8), "@@@@-@@-@@")
        PrintText 30, (LineCnt * DefLineSpage) + DefPointTop, Format(ObjRSet.Fields("ÀÔ°íÀÏ"), "@@@@-@@-@@")
        PrintText 50, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("ÀüÈ­¹øÈ£")
        PrintText 70, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("¼º¸í")
        PrintText 90, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("Ç°¸í")
        PrintText 125, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("¹øÈ£")
        PrintText 145, (LineCnt * DefLineSpage) + DefPointTop, ObjRSet.Fields("»ö»ó")
        
        
        nTotalMoney = nTotalMoney + Val(ObjRSet.Fields("±Ý¾×"))
        strMaxLng = "1234567890"
        RSet strMaxLng = Format(Val(ObjRSet.Fields("±Ý¾×")), "#,##0")
        PrintText 160, (LineCnt * DefLineSpage) + DefPointTop, strMaxLng
    
        If PRINT_LINE_COUNT <= LineCnt Then
            PageCnt = PageCnt + 1
            PrintLine 0, 260, 180, 260      '¼öÆò ¶óÀÎ
            PrintText 150, 262, "[" & ´ë¸®Á¡Á¤º¸.´ë¸®Á¡¹øÈ£ & "] " & ´ë¸®Á¡Á¤º¸.´ë¸®Á¡¸í
            
'            Printer.EndDoc
'            Exit Function
            Printer.NewPage
            GoSub PrintDefault
            LineCnt = 0
        End If
        ObjRSet.MoveNext
    Loop
    
    LineCnt = LineCnt + 1
    PrintLine 0, (LineCnt * DefLineSpage) + DefPointTop, 180, (LineCnt * DefLineSpage) + DefPointTop      '¼öÆò ¶óÀÎ
    
    LineCnt = LineCnt + 1
    PrintText 0, (LineCnt * DefLineSpage) + DefPointTop, "ºÒ·® ¼¼Å¹ºñ È¯ºÒ ÇÕ°è±Ý¾×: " & Format(nTotalMoney, "#,##0") & " ¿ø"
    PrintText 80, (LineCnt * DefLineSpage) + DefPointTop, "Áö»ç ºÎ´ã±Ý: " & Format(nTotalMoney * (1 - (Val(´ë¸®Á¡Á¤º¸.ºñÀ²) / 100)), "#,##0") & " ¿ø"
    PrintText 140, (LineCnt * DefLineSpage) + DefPointTop, "¸ÅÀå ºÎ´ã±Ý: " & Format(nTotalMoney * (Val(´ë¸®Á¡Á¤º¸.ºñÀ²) / 100), "#,##0") & " ¿ø"
    
    
    Dim sStrTemp(1 To 5) As String
    sStrTemp(1) = "¨ç  ºÒ·® ¼¼Å¹ È¯ºÒ±Ý Ã»±¸¾×Àº Áö»ç ½ÂÀÎ ÈÄ 3ÀÏ ÈÄ °áÁ¦ ÇØµå¸®°Ú½À´Ï´Ù."
    sStrTemp(2) = "¨è  ºÒ·® ¼¼Å¹ È¯ºÒÀº 1È¸ Àç ¼¼Å¹ ÈÄ¿¡µµ Ç°Áú »óÅÂ°¡ ºÒ·®ÇÑ °æ¿ì¿¡ È¯ºÒÇÏ´Â Á¦µµ ÀÔ´Ï´Ù."
    sStrTemp(3) = "¨é  ºÒ·® ¼¼Å¹ È¯ºÒ±ÝÀ» Áö»ç¿¡ Ã»±¸ÇÏ´Â °æ¿ì ÇØ´ç ¹ÙÄÚµåÅÃÀ» º» ¸®½ºÆ®¿¡ ºÎÂø ¹Ù¶ø´Ï´Ù."
    sStrTemp(4) = "¨ê  ºÒ·® ¼¼Å¹ È¯ºÒ±Ý Áö»ç Ã»±¸ °¡´ÉÀÏÀº È¯ºÒÀÏ·ÎºÎÅÍ 7ÀÏ ÀÌ³»¿¡ ¼­¸éÀ¸·Î Ã»±¸ÇØ¾ß ÇÕ´Ï´Ù."
    sStrTemp(5) = "¨ë  ºÒ·® ¼¼Å¹ È¯ºÒ °í°´ºÐÀÇ ¿¬¶ôÃ³, ÅÃ¹øÈ£, Á¢¼öÀÏÀÚ¸¦ ±â·ÏÇØ¾ß È¯ºÒÀÌ °¡´ÉÇÕ´Ï´Ù."
    
    PrintLine 0, 232, 180, 232      '¼öÆò ¶óÀÎ
    PrintText 0, 235, sStrTemp(1)
    PrintText 0, 240, sStrTemp(2)
    PrintText 0, 245, sStrTemp(3)
    PrintText 0, 250, sStrTemp(4)
    PrintText 0, 255, sStrTemp(5)
    
    PageCnt = PageCnt + 1
    PrintLine 0, 260, 180, 260      '¼öÆò ¶óÀÎ
    PrintText 150, 262, "[" & ´ë¸®Á¡Á¤º¸.´ë¸®Á¡¹øÈ£ & "] " & ´ë¸®Á¡Á¤º¸.´ë¸®Á¡¸í
    
    'Ãâ·Â ÇÑ´Ù.
    ObjRSet.Close
    Printer.EndDoc
    
    Exit Function
    
printError:
    Print¼¼Å¹ºñÈ¯ºÒÇöÈ² = False
    Printer.EndDoc
    Exit Function
    
    
' ±âº»»çÇ× ¾ç½Ä ÀÛ¼º
PrintDefault:
    
    Top_Margin = 0: Left_Margin = 0
    
    Printer.FontName = "±¼¸²Ã¼"
    Printer.Font.Bold = True
    Printer.Font.Size = 9
    
    Printer.Font.Size = "18"
    PrintText 55, 8, Title
    Printer.DrawWidth = 12
    PrintLine 50, 16, 125, 16
    
    Printer.Font.Size = 9
    Printer.DrawWidth = 7
    Text_Height = Printer.TextHeight("aa")
    
    PrintLine 0, 25, 180, 25   '¼öÆò ¶óÀÎ
    PrintLine 0, 32, 180, 32    '¼öÆò ¶óÀÎ
    PrintText 160, 20, CStr(PageCnt) & " / " & CStr(TotalPage) & " Page"
    
    PrintText 0, 20, "[" & ´ë¸®Á¡Á¤º¸.´ë¸®Á¡¹øÈ£ & "] " & ´ë¸®Á¡Á¤º¸.´ë¸®Á¡¸í & "°Ë»ö ÀÏÀÚ : " & sDateFrom & "~" & sDateTo
    
    
    
    PrintText 0, 27, "¼ø¹ø"
    PrintText 10, 27, "È¯ºÒÀÏÀÚ"
    PrintText 30, 27, "ÀÔ°íÀÏÀÚ"
    PrintText 50, 27, "ÀüÈ­¹øÈ£"
    PrintText 70, 27, "¼º   ¸í"
    PrintText 90, 27, "Ç°    ¸í"
    PrintText 125, 27, "¹ø  È£"
    PrintText 145, 27, "»ö»ó"
    PrintText 160, 27, "±Ý   ¾×"
    Return
    
' ÇÇÇØ°ü·Ã»çÇ×
PrintDamage:

    Top_Margin = 0: Left_Margin = 0
    Printer.Font.Size = 9
    Printer.DrawWidth = 7
    Text_Height = Printer.TextHeight("a")
    
    PrintRect 0, 105, 180, 140      '¿Ü°¢ Æ²
    PrintLine 0, 112, 180, 112      '¼öÆò ¶óÀÎ
    PrintLine 0, 119, 180, 119      '¼öÆò ¶óÀÎ
    PrintLine 0, 126, 180, 126      '¼öÆò ¶óÀÎ
    PrintLine 0, 133, 180, 133      '¼öÆò ¶óÀÎ
    
    PrintLine 35, 105, 35, 140      '¼öÁ÷ ¶óÀÎ
    PrintLine 85, 105, 85, 140      '¼öÁ÷ ¶óÀÎ
    PrintLine 120, 105, 120, 140      '¼öÁ÷ ¶óÀÎ
    PrintLine 135, 119, 135, 133      '¼öÁ÷ ¶óÀÎ
    PrintLine 150, 119, 150, 133      '¼öÁ÷ ¶óÀÎ
    PrintLine 165, 119, 165, 133      '¼öÁ÷ ¶óÀÎ

    Return
End Function

