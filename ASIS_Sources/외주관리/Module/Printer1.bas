Attribute VB_Name = "Printer1"
Option Explicit
Public Type USER_TYPE_RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Public Type PRINT_DATA_TYPE
    MagRect     As USER_TYPE_RECT            ' Ãâ·ÂÇÒ ÁÂÇ¥
    
    LeftLooper  As Integer              ' ÇÑ ¶óÀÎ¿¡ ¿©·¯°³ Ãâ·ÂµÉ°æ¿ì ÇöÀç Ãâ·Â °¹¼ö
    LeftMaxLooper   As Integer          ' ÇÑ ¶óÀÎ¿¡ Ãâ·ÂÇÒ ÀüÃ¼ °¹¼ö
    LeftNextSpace   As Integer          ' ÇÑ ¶óÀÎ¿¡ ¿©·¯°³ Ãâ·ÂµÉ°æ¿ì ´ÙÀ½ Ãâ·Â À§Ä¡
    
    MaxCount        As Integer      ' Ãâ·ÂÇÒ ÀüÃ¼ °Ç ¼ö
    MaxPage         As Integer      ' Ãâ·ÂÇÒ ÀüÃ¼ ÆäÀÌÁö ¼ö
    ProcCnt         As Integer      ' ÇöÀçÃâ·Â ÁßÀÎ °Ç¼ö
    ProcPageCnt     As Integer      ' ÇöÀç Ãâ·ÂÁßÀÎ ÆäÀÌÁö ¼ö

    PageProcMaxCnt   As Integer      ' ÇÑÆäÀÌÁö´ç ÀüÃ¼ Ãâ·Â¼ö ( Details ¼ö )
    PageProcCnt      As Integer      ' ÇÑÆäÀÌÁö¿¡¼­ DetailsÀÇ ÇöÀç Ãâ·ÂÁßÀÎ ¼ö
    NextLineSpace    As Integer      ' DetailsÀÇ ´ÙÀ½ ¶óÀÎ¿¡ Ãâ·ÂµÉ ¿©¹é
End Type


' Ãâ·Â¿¡ ÇÊ¿äÇÑ º¯¼ö
Public bMsgMode As Boolean
Public strMessage As String
Public Type PrintParamTYPE
    Param() As String
End Type
Public PrtParam As PrintParamTYPE

Public Type RECT_TYPE
    Left    As Long
    Top     As Long
    Right   As Long
    Botton  As Long
End Type


' ¹Ì¸®º¸±â
'¹Ì¸®º¸±â¿¡¼­ ÀÎ¼â


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
Public strFilename             As String    ' ÀÎ¼âÇÒ È­ÀÏ ÀÌ¸§
Public PageCnt, LineCnt        As Integer   ' ÀÎ¼âÇÒ ÆäÀÌÁö ¹× ¶óÀÎ¼ö
Public PRINT_LINE_COUNT         As Integer


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


Type PrintPoint
    x As Integer
    y As Integer
End Type


Public PrtPoint As PrintPoint   ' ±âº» ÁÂÇ¥
Public PrtPoint2 As PrintPoint  ' ¶óÀÎ°£°Ý
Public PrtPoint3 As PrintPoint  ' ¼Õ´Ô¿ë
Public PrtPoint4 As PrintPoint  ' ¿©¹é
'''''''''''''''''''''''''''''''''''''
Public FPArray(1 To 100, 1 To 6) As Variant


Public Page_Count As Integer
' Ãâ·ÂÇÒ Ç×¸ñÀÇ ÃÑ °¹¼ö
Private intRowCount As Integer
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

Public Function Fn_PrinterCheck() As Boolean

On Error GoTo ERR_RTN
  Dim printer_name As String
  
  Dim x As Printer
    
      For Each x In Printers
          printer_name = x.DeviceName
      Next

      If Printer.DeviceName = "" Then
        MsgBox "ÇÁ¸°ÅÍ¸¦ ¼³Ä¡ÇØ ÁÖ½Ê½Ã¿ä!", vbInformation, "È®ÀÎ"
        Fn_PrinterCheck = False
        Exit Function
        
    End If
    Fn_PrinterCheck = True
    Exit Function
  
ERR_RTN:
    Fn_PrinterCheck = False
  Exit Function
    
End Function

'If InStr(1, ppp$, "") > 0 Then
'È­ÀÏÀ» ÇÁ¸°ÅÍ·Î  Ãâ·Â ÇÑ´Ù.
'*****************************************************************
Public Sub FileToPrint(strFilename$, Ãâ·Â¹æÇâ As Integer, bView As Boolean)
Dim ppp$

    On Error GoTo Error_Handle
    If bView Then
        ' ¹Ì¸® º¸±âÀÌ¸é
            EDIT_Text (strFilename)
    Else
        ' ÀÎ¼â
        FHandle = FreeFile
        Printer.FontName = "±¼¸²Ã¼"
''           Printer.ShowPrinter
        Printer.Orientation = Ãâ·Â¹æÇâ
        Open strFilename For Input As #FHandle
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


Function funLeft(ByVal Txt As String, ByVal Length As Integer) As String
    Dim iCnt As Integer
    Dim TrimCnt0, TrimCnt1 As Integer
    Dim iLoop As Integer
    
    iCnt = Len(Txt)
    TrimCnt0 = 0
    TrimCnt1 = 0
    For iLoop = 1 To iCnt
        If Asc(Mid(Txt, iCnt, 1)) > 0 Then
            TrimCnt1 = TrimCnt1 + 1
        Else
            TrimCnt1 = TrimCnt1 + 2
        End If
        If TrimCnt1 > Length Then
            funLeft = MidB(Txt, 1, TrimCnt0)
            Exit Function
        Else
            TrimCnt0 = TrimCnt1
        End If
            
    Next iLoop
    funLeft = Txt
End Function
Function PrNumSet(Num As Variant, cnt As Integer)

    Dim Num1 As Double
    Dim Str As String

    Num1 = Val(Num)
    Str = "                           " & Format(Num1, "#,##0")
    PrNumSet = Right(Str, cnt)
End Function

Public Sub PrintRect(PView As Object, spX As Long, spY As Long, _
                     epX As Long, epY As Long, Optional DrawWidth As Integer = 2)
        
    ' ¿©¹éÀ» Àû¿ë ½ÃÅ²´Ù.
'    spX = spX + Prt_Left + Left_Margin:    spY = spY + Prt_Top + Top_Margin
'    epX = epX + Prt_Left + Left_Margin:    epY = epY + Prt_Top + Top_Margin

    PView.DrawWidth = DrawWidth
    PView.DrawStyle = vbSolid
    PView.Line (spX, spY)-(epX, epY), , B

End Sub

Public Sub PrintLine(PView As Object, spX As Long, spY As Long, epX As Long, _
                     epY As Long, Optional DrawWidth As Integer = 2)
        
    ' ¿©¹éÀ» Àû¿ë ½ÃÅ²´Ù.
'    spX = spX + Prt_Left + Left_Margin:    spY = spY + Prt_Top + Top_Margin
'    epX = epX + Prt_Left + Left_Margin:    epY = epY + Prt_Top + Top_Margin

    PView.DrawWidth = DrawWidth
    PView.DrawStyle = vbSolid
    PView.Line (spX, spY)-(epX, epY)

End Sub

Public Sub PrintText(PView As Object, spX As Long, spY As Long, MSG As String)
        
    ' ¿©¹éÀ» Àû¿ë ½ÃÅ²´Ù.
'    spX = spX + Prt_Left + Left_Margin:  spY = spY + Prt_Top + Top_Margin
    
    PView.CurrentX = spX
    PView.CurrentY = spY
    PView.Print MSG

End Sub

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

 


Public Sub PrnRefresh(PView As Object, SL_Type As Single)
    PView.ScaleMode = vbMillimeters
    PView.Width = 210 * (1440 / 25.4) * SL_Type
    PView.Height = 295 * (1440 / 25.4) * SL_Type
    PView.Scale (0, 0)-(210, 295)

End Sub


