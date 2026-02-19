VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CISNTaxPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 공급자용 자료
Private Type Company_Info_TYPE
    CompanyNumber       As String
    CompanyName         As String
    DaePyo              As String
    Address1            As String
    Address2            As String
    UpTage              As String
    UpJong              As String
End Type

' 리스트 자료
Private Type TaxListData_Info_TYPE
    Month       As String
    Day         As String
    Item        As String
    Spec        As String
    Qty         As String
    DanGa       As String
    Net         As String
    Vat         As String
    Remark      As String
End Type



Private Const ERROR_BASE = vbObjectError + 3000
Private Const ERROR_PRINTER_NOT_FOUND = ERROR_BASE + 1
Private Const ERROR_FILE_NOT_FOUND = ERROR_BASE + 2
Private Const ERROR_ORDER = ERROR_BASE + 3


'이 이벤트를 발생시키려면 다음과 같은 구문을 가지는 RaiseEvent를 사용하십시오.
'RaiseEvent NotPrinter[(arg1, arg2, ... , argn)]
Public Event ErrorPrinter(ErrNum As Integer, ErrDec As String)
Public Event PrintPage(mode As String, PrtPage As Long, TotalPage As Long)

'--------------------------------------------------------------------------------------------------------------
' Procedure : Fn_PrinterCheck
' DateTime  : 2007-01-02 12:49
' Author    : pds2004
' Purpose   : 기본 프린터 찾기
'--------------------------------------------------------------------------------------------------------------
Private Function Fn_PrinterCheck() As Boolean

    On Error GoTo ERR_RTN
    Dim printer_name As String
  
    Dim x As Printer
    
    For Each x In Printers
        printer_name = x.DeviceName
    Next

    If Printer.DeviceName = "" Then
        GoSub ERR_RTN
        Exit Function
    End If
    
    Fn_PrinterCheck = True
    Exit Function
  
ERR_RTN:
    Fn_PrinterCheck = False
    RaiseEvent ErrorPrinter(ERROR_PRINTER_NOT_FOUND, "기본 프린터를 찾을 수 없습니다.")
End Function

Private Sub PrnRefresh(PView As Object, SL_Type As Single)
    Dim PageWidth   As Integer
    Dim PageHeight  As Integer
    
    ' 기본 A4 양식으로 생각하고 그린다.
    ' Scale(0,0)-(210,295) 로 설정하고 그릴경우  좌표가 저 범위를 넘어서면 안된다.
    ' 해당 비율로 적용되어 그려진다.
    Dim DefWidth    As Integer
    Dim DefHeight   As Integer
    DefWidth = 210
    DefHeight = 295

    If UCase(TypeName(PView)) = "PICTUREBOX" Then
        PView.Cls
        PageWidth = 210
        PageHeight = 295
        PView.Width = PageWidth * (1440 / 25.4) * SL_Type
        PView.Height = PageHeight * (1440 / 25.4) * SL_Type
        
    Else
        Debug.Print Printer.PaperSize
        
        Select Case Printer.PaperSize
            Case vbPRPSLetter
                PageWidth = 220
                PageHeight = 280
            
            Case vbPRPSLetter
                PageWidth = 220
                PageHeight = 280
            
            Case vbPRPSA4
                PageWidth = 210
                PageHeight = 295
                
            Case Else
                PageWidth = 220
                PageHeight = 280
        End Select
    End If
    
    PView.ScaleMode = vbMillimeters
    PView.Scale (0, 0)-(DefWidth, DefHeight)
End Sub

Private Sub PrintRect(PView As Object, Def As PRINT_DATA_TYPE, spX As Long, spY As Long, _
                     epX As Long, epY As Long, Optional DrawWidth As Integer = 2)
    PView.FillStyle = 1
    PView.DrawWidth = DrawWidth
    PView.DrawStyle = vbSolid
    PView.Line (Def.MagRect.Left + spX, Def.MagRect.Top + spY)-(Def.MagRect.Left + epX, Def.MagRect.Top + epY), , B
End Sub
Private Sub PrintText(PView As Object, Def As PRINT_DATA_TYPE, spX As Long, spY As Long, MSG As String)
    PView.CurrentX = Def.MagRect.Left + spX
    PView.CurrentY = Def.MagRect.Top + spY
    PView.Print MSG
End Sub

Private Sub PrintLine(PView As Object, Def As PRINT_DATA_TYPE, spX As Long, spY As Long, epX As Long, _
                     epY As Long, Optional DrawWidth As Integer = 2)
    
    PView.DrawWidth = DrawWidth
    PView.DrawStyle = vbSolid
    PView.Line (Def.MagRect.Left + spX, Def.MagRect.Top + spY)-(Def.MagRect.Left + epX, Def.MagRect.Top + epY)

End Sub

 
Public Sub TaxForm_Print(ByRef pb As Object, sf As Single, wdLine As Integer)
'Dim x1, y1 As Single
'Dim intItem As Integer
'Dim i           As Integer
'Dim TmpStr      As String
'Dim AllSpace    As Integer
'Dim TNET(11)   As String
'Dim TVAT(11)   As String
'Dim kk         As Integer
'Dim PointSizeX As Integer
'Dim PointSizeY As Integer
'Dim PointSize As Integer
'
'    On Error Resume Next
'    AllSpace = 11
'    For i = 0 To Len(Trim(Str(Tax_NET))) - 1
'        TNET(i) = Mid(Trim(Str(Tax_NET)), (Len(Trim(Str(Tax_NET))) - i), 1)
'    Next i
'
'    For i = 0 To Len(Trim(Str(Tax_VAT))) - 1
'        TVAT(i) = Mid(Trim(Str(Tax_VAT)), (Len(Trim(Str(Tax_VAT))) - i), 1)
'    Next i
'
'
'For intItem = 1 To 2
'    If intItem = 1 Then
'        x1 = 30
'        y1 = 20
'        If Option_Printer_Color.Value Then
'            LineColor = vbBlue
'        Else
'            LineColor = vbBlack
'        End If
'    Else
'        x1 = 30
'        y1 = y1 + 140
'        If Option_Printer_Color.Value Then
'            LineColor = vbRed
'        Else
'            LineColor = vbBlack
'        End If
'    End If
'
'
'    ' 계산서 도장을 출력할 경우
'    If Dir(App.Path & "\image\Stamp.jpg", vbNormal) <> "" Then
'        Set imgStamp.Picture = LoadPicture(App.Path & "\image\Stamp.jpg")
'        Set StampImg = imgStamp
'        '도장 인쇄
'        pb.PaintPicture StampImg.Picture, x1 + 63, y1 + 18, 18, 18
'    End If
'
'
'    '먼저 굵은 선들을 그리고나서 가는선들을 그린다.
'    '테두리선을 그린다.
'    pb.DrawStyle = vbSolid
'    pb.DrawWidth = wdLine * 2
'
'    '----------------------------------------------------------------------------------
'    ' 모퉁이의 절취선을 표시한다.
'    PointSizeX = 13
'    PointSizeY = 9
'    PointSize = 1
'    Call Display_Point(pb, x1 - PointSizeX - 5, y1 - PointSizeY, PointSize)
'    Call Display_Point(pb, x1 + 153 + PointSizeX, y1 - PointSizeY, PointSize)
'    Call Display_Point(pb, x1 - PointSizeX - 5, y1 + 111 + PointSizeY, PointSize)
'    Call Display_Point(pb, x1 + 153 + PointSizeX, y1 + 111 + PointSizeY, PointSize)
'    '----------------------------------------------------------------------------------
'
'
'
'    pb.Line (x1, y1)-(x1 + 153, y1), LineColor             '위 끝
'    pb.Line (x1 + 153, y1)-(x1 + 153, y1 + 111), LineColor            '오른쪽 끝
'    pb.Line (x1 + 153, y1 + 111)-(x1, y1 + 111), LineColor             '아래끝
'    pb.Line (x1, y1 + 111)-(x1, y1), LineColor                         '왼쪽끝
'
'    '공급자등록번호 굵은 테두리
'    pb.Line (x1 + 5, y1 + 15)-(x1 + 76, y1 + 15), LineColor        '위
'    pb.Line (x1 + 76, y1 + 15)-(x1 + 76, y1 + 31), LineColor       '오른쪽
'    pb.Line (x1 + 76, y1 + 31)-(x1 + 5, y1 + 31), LineColor        '아래
'    pb.Line (x1 + 5, y1 + 31)-(x1 + 5, y1 + 15), LineColor         '왼쪽
'    '공급받는자 등록번호 테두리
'    pb.Line (x1 + 153, y1 + 15)-(x1 + 81, y1 + 15), LineColor      '위
'    pb.Line (x1 + 81, y1 + 15)-(x1 + 81, y1 + 23), LineColor       '왼쪽
'    pb.Line (x1 + 81, y1 + 23)-(x1 + 153, y1 + 23), LineColor      '아래
'    pb.Line (x1 + 123, y1 + 9)-(x1 + 123, y1 + 15), LineColor     '중간선 1
'    pb.Line (x1 + 129, y1 + 9)-(x1 + 129, y1 + 23), LineColor  '   중간선 2
'
'    '책번호, 일련번호의 왼쪽선
'    pb.Line (x1 + 111, y1)-(x1 + 111, y1 + 23), LineColor          '왼쪽
'    pb.Line (x1 + 115, y1 + 15)-(x1 + 115, y1 + 23), LineColor     '중간선 1
'    pb.Line (x1 + 125, y1 + 15)-(x1 + 125, y1 + 23), LineColor     '중간선 2
'
'    '금액란 굵은선
'    pb.Line (x1, y1 + 47)-(x1 + 133, y1 + 47), LineColor           '위쪽
'    pb.Line (x1 + 133, y1 + 47)-(x1 + 133, y1 + 66), LineColor     '오른쪽
'    pb.Line (x1 + 133, y1 + 66)-(x1, y1 + 66), LineColor           '아래쪽
'
'    pb.Line (x1 + 37, y1 + 52)-(x1 + 37, y1 + 66), LineColor       '중간선 1
'    pb.Line (x1 + 52, y1 + 52)-(x1 + 52, y1 + 66), LineColor       '중간선 2
'    pb.Line (x1 + 67, y1 + 52)-(x1 + 67, y1 + 66), LineColor       '중간선 3
'    pb.Line (x1 + 88, y1 + 52)-(x1 + 88, y1 + 66), LineColor       '중간선 4
'    pb.Line (x1 + 103, y1 + 52)-(x1 + 103, y1 + 66), LineColor     '중간선 5
'    pb.Line (x1 + 118, y1 + 52)-(x1 + 118, y1 + 66), LineColor     '중간선 6
'
'    pb.DrawWidth = wdLine
'    '책번호 일련번호한의 가는선들
'    pb.Line (x1 + 111, y1 + 9)-(x1 + 153, y1 + 9), LineColor       '가로
'    pb.Line (x1 + 129, y1)-(x1 + 129, y1 + 9), LineColor           '세로
'    pb.Line (x1 + 117, y1 + 9)-(x1 + 117, y1 + 15), LineColor      '세로
'    pb.Line (x1 + 135, y1 + 9)-(x1 + 135, y1 + 15), LineColor      '세로
'    pb.Line (x1 + 141, y1 + 9)-(x1 + 141, y1 + 15), LineColor      '세로
'    pb.Line (x1 + 147, y1 + 9)-(x1 + 147, y1 + 15), LineColor      '세로
'    pb.Line (x1 + 124.5, y1 + 12)-(x1 + 127.5, y1 + 12), LineColor '대쉬
'
'    '공급자 란의 가는선들
'    pb.Line (x1, y1 + 15)-(x1 + 5, y1 + 15), LineColor             '가로
'    pb.Line (x1 + 5, y1 + 23)-(x1 + 76, y1 + 23), LineColor        '가로
'    pb.Line (x1 + 5, y1 + 39)-(x1 + 76, y1 + 39), LineColor        '가로
'    pb.Line (x1 + 5, y1 + 31)-(x1 + 5, y1 + 47), LineColor         '세로
'    pb.Line (x1 + 20, y1 + 15)-(x1 + 20, y1 + 66), LineColor       '세로
'    pb.Line (x1 + 45, y1 + 23)-(x1 + 45, y1 + 31), LineColor       '세로
'    pb.Line (x1 + 54, y1 + 23)-(x1 + 54, y1 + 31), LineColor       '세로
'    pb.Line (x1 + 41, y1 + 39)-(x1 + 41, y1 + 47), LineColor       '세로
'    pb.Line (x1 + 46, y1 + 39)-(x1 + 46, y1 + 47), LineColor       '세로
'    pb.Line (x1 + 76, y1 + 31)-(x1 + 76, y1 + 47), LineColor       '세로
'
'    '공급받는자 란의 가는선
'    pb.Line (x1 + 76, y1 + 15)-(x1 + 81, y1 + 15), LineColor       '가로
'    pb.Line (x1 + 81, y1 + 31)-(x1 + 153, y1 + 31), LineColor      '"
'    pb.Line (x1 + 81, y1 + 39)-(x1 + 153, y1 + 39), LineColor      '"
'    pb.Line (x1 + 133, y1 + 47)-(x1 + 153, y1 + 47), LineColor     '"
'    pb.Line (x1 + 81, y1 + 23)-(x1 + 81, y1 + 47), LineColor       '세로
'    pb.Line (x1 + 96, y1 + 15)-(x1 + 96, y1 + 47), LineColor       '"
'    pb.Line (x1 + 101, y1 + 15)-(x1 + 101, y1 + 23), LineColor     '"
'    pb.Line (x1 + 106, y1 + 15)-(x1 + 106, y1 + 23), LineColor     '"
'    pb.Line (x1 + 120, y1 + 15)-(x1 + 120, y1 + 23), LineColor     '"
'    pb.Line (x1 + 133, y1 + 15)-(x1 + 133, y1 + 23), LineColor     '"
'    pb.Line (x1 + 138, y1 + 15)-(x1 + 138, y1 + 23), LineColor     '"
'    pb.Line (x1 + 143, y1 + 15)-(x1 + 143, y1 + 23), LineColor     '"
'    pb.Line (x1 + 148, y1 + 15)-(x1 + 148, y1 + 23), LineColor     '"
'    pb.Line (x1 + 124, y1 + 23)-(x1 + 124, y1 + 31), LineColor     '"
'    pb.Line (x1 + 133, y1 + 23)-(x1 + 133, y1 + 31), LineColor     '"
'    pb.Line (x1 + 118, y1 + 39)-(x1 + 118, y1 + 47), LineColor     '"
'    pb.Line (x1 + 123, y1 + 39)-(x1 + 123, y1 + 47), LineColor     '"
'    pb.Line (x1 + 112, y1 + 19)-(x1 + 114, y1 + 19), LineColor     '대쉬
'    pb.Line (x1 + 126, y1 + 19)-(x1 + 128, y1 + 19), LineColor     '"
'
'    '금액란의 가는선들
'    pb.Line (x1, y1 + 52)-(x1 + 153, y1 + 52), LineColor           '가로
'    pb.Line (x1, y1 + 56)-(x1 + 133, y1 + 56), LineColor           '"
'    pb.Line (x1 + 133, y1 + 66)-(x1 + 153, y1 + 66), LineColor     '"
'
'    pb.Line (x1 + 7, y1 + 52)-(x1 + 7, y1 + 66), LineColor         '세로
'    pb.Line (x1 + 13, y1 + 52)-(x1 + 13, y1 + 66), LineColor       '"
'    For i = 27 To 77 Step 5
'        pb.Line (x1 + i, y1 + 52)-(x1 + i, y1 + 66), LineColor     '"
'    Next
'    pb.Line (x1 + 82, y1 + 47)-(x1 + 82, y1 + 66), LineColor       '"
'    pb.Line (x1 + 83, y1 + 47)-(x1 + 83, y1 + 66), LineColor       '"
'    For i = 88 To 128 Step 5
'        pb.Line (x1 + i, y1 + 52)-(x1 + i, y1 + 66), LineColor     '"
'    Next
'
'    '주문 리스트 란의 가는선들
'    pb.Line (x1, y1 + 71)-(x1 + 153, y1 + 71), LineColor
'    pb.Line (x1, y1 + 78.5)-(x1 + 153, y1 + 78.5), LineColor
'    pb.Line (x1, y1 + 86)-(x1 + 153, y1 + 86), LineColor
'    pb.Line (x1, y1 + 93.5)-(x1 + 153, y1 + 93.5), LineColor
'    pb.Line (x1, y1 + 101)-(x1 + 153, y1 + 101), LineColor
'    pb.Line (x1, y1 + 105)-(x1 + 115, y1 + 105), LineColor
'
'    pb.Line (x1 + 5, y1 + 71)-(x1 + 5, y1 + 101), LineColor
'    pb.Line (x1 + 10, y1 + 66)-(x1 + 10, y1 + 101), LineColor
'    pb.Line (x1 + 45, y1 + 66)-(x1 + 45, y1 + 101), LineColor
'    pb.Line (x1 + 58, y1 + 66)-(x1 + 58, y1 + 101), LineColor
'    pb.Line (x1 + 73, y1 + 66)-(x1 + 73, y1 + 101), LineColor
'    pb.Line (x1 + 90, y1 + 66)-(x1 + 90, y1 + 101), LineColor
'    pb.Line (x1 + 120, y1 + 66)-(x1 + 120, y1 + 101), LineColor
'    pb.Line (x1 + 140, y1 + 66)-(x1 + 140, y1 + 101), LineColor
'    pb.Line (x1 + 26, y1 + 101)-(x1 + 26, y1 + 111), LineColor
'    pb.Line (x1 + 48, y1 + 101)-(x1 + 48, y1 + 111), LineColor
'    pb.Line (x1 + 71, y1 + 101)-(x1 + 71, y1 + 111), LineColor
'    pb.Line (x1 + 93, y1 + 101)-(x1 + 93, y1 + 111), LineColor
'    pb.Line (x1 + 115, y1 + 101)-(x1 + 115, y1 + 111), LineColor
'
'    '글자들 쓰기
'
'    pb.Font = "명조체"
'    pb.FontSize = 20 * sf
'    pb.FontBold = True
'    pb.ForeColor = LineColor
'    pb.CurrentX = x1 + 15: pb.CurrentY = y1 + 4
'    pb.Print "세금계산서"
'
'    pb.Font = "돋움체"
'    pb.FontSize = 10 * sf
'    pb.FontBold = False
'    pb.CurrentX = x1 + 54: pb.CurrentY = y1 + 6
'
'    If intItem = 1 Then
'        pb.Print "(공급받는자보관용)"
'    Else
'        pb.Print "(공급자 보관용)"
'    End If
'
'pb.FontSize = 8 * sf
'
'    pb.CurrentX = x1 + 1: pb.CurrentY = y1 - 3.5:       pb.Print "[별지 제11호 서식] (96. 3. 30 개정)"
'    pb.CurrentX = x1 + 142: pb.CurrentY = y1 - 3.5
'    If intItem = 1 Then
'        pb.Print "(청색)"
'    Else
'        pb.Print "(적색)"
'    End If
'    pb.Circle (x1 + 72.5, y1 + 26.1), 1.7 * sf
'    pb.CurrentX = x1 + 71: pb.CurrentY = y1 + 25:       pb.Print "인"
'    pb.Circle (x1 + 149.5, y1 + 26.1), 1.7 * sf
'    pb.CurrentX = x1 + 148: pb.CurrentY = y1 + 25:       pb.Print "인"
'
'pb.FontSize = 5 * sf
'    pb.CurrentX = x1 + 46: pb.CurrentY = y1 + 28:       pb.Print "(대표자)"
'
'    pb.CurrentX = x1 + 125: pb.CurrentY = y1 + 28:     pb.Print "(대표자)"
'    'pb.CurrentX = x1 + 127.7: pb.CurrentY = y1 + 28:     pb.Print "표"
'    'pb.CurrentX = x1 + 128.9: pb.CurrentY = y1 + 28:     pb.Print "자)"
'
'    pb.CurrentX = x1 + 20.5: pb.CurrentY = y1 + 53:        pb.Print "공란수"
'
'pb.FontSize = 9 * sf
'
'    pb.CurrentX = x1 + 95: pb.CurrentY = y1 + 3:        pb.Print "책 번 호"
'    pb.CurrentX = x1 + 95: pb.CurrentY = y1 + 10:       pb.Print "일련번호"
'    pb.CurrentX = x1 + 125: pb.CurrentY = y1 + 3:       pb.Print "권"
'    pb.CurrentX = x1 + 149: pb.CurrentY = y1 + 3:       pb.Print "호"
'
'    pb.CurrentX = x1 + 1: pb.CurrentY = y1 + 20:        pb.Print "공"
'    pb.CurrentX = x1 + 1: pb.CurrentY = y1 + 29:        pb.Print "급"
'    pb.CurrentX = x1 + 1: pb.CurrentY = y1 + 39:        pb.Print "자"
'
'
'    pb.CurrentX = x1 + 6.5: pb.CurrentY = y1 + 17.5:    pb.Print "등록번호"
'    pb.CurrentX = x1 + 6.5: pb.CurrentY = y1 + 23.5:    pb.Print "상    호"
'    pb.CurrentX = x1 + 6.5: pb.CurrentY = y1 + 27:      pb.Print "(법인명)"
'    pb.CurrentX = x1 + 6.5: pb.CurrentY = y1 + 32:      pb.Print "사 업 장"
'    pb.CurrentX = x1 + 6.5: pb.CurrentY = y1 + 35.5:    pb.Print "주    소"
'    pb.CurrentX = x1 + 6.5: pb.CurrentY = y1 + 41:      pb.Print "업    태"
'    pb.CurrentX = x1 + 42: pb.CurrentY = y1 + 40:       pb.Print "종"
'    pb.CurrentX = x1 + 42: pb.CurrentY = y1 + 43.5:     pb.Print "목"
'    pb.CurrentX = x1 + 46.5: pb.CurrentY = y1 + 24:     pb.Print "성명"
'
'
'
'    pb.CurrentX = x1 + 77: pb.CurrentY = y1 + 20:        pb.Print "공"
'    pb.CurrentX = x1 + 77: pb.CurrentY = y1 + 25:        pb.Print "급"
'    pb.CurrentX = x1 + 77: pb.CurrentY = y1 + 30:        pb.Print "받"
'    pb.CurrentX = x1 + 77: pb.CurrentY = y1 + 35:        pb.Print "는"
'    pb.CurrentX = x1 + 77: pb.CurrentY = y1 + 40:        pb.Print "자"
'
'    pb.CurrentX = x1 + 82.5: pb.CurrentY = y1 + 17.5:    pb.Print "등록번호"
'    pb.CurrentX = x1 + 82.5: pb.CurrentY = y1 + 23.5:    pb.Print "상    호"
'    pb.CurrentX = x1 + 82.5: pb.CurrentY = y1 + 27:      pb.Print "(법인명)"
'    pb.CurrentX = x1 + 82.5: pb.CurrentY = y1 + 32:      pb.Print "사 업 장"
'    pb.CurrentX = x1 + 82.5: pb.CurrentY = y1 + 35.5:    pb.Print "주    소"
'    pb.CurrentX = x1 + 82.5: pb.CurrentY = y1 + 41:      pb.Print "업    태"
'    pb.CurrentX = x1 + 119: pb.CurrentY = y1 + 40:       pb.Print "종"
'    pb.CurrentX = x1 + 119: pb.CurrentY = y1 + 43.5:     pb.Print "목"
'    pb.CurrentX = x1 + 125.5: pb.CurrentY = y1 + 24:     pb.Print "성명"
'    '-----------------------------------------------------------------------------------------------
'    '상단내용출력
'    pb.ForeColor = vbBlack
'    pb.CurrentX = x1 + 112: pb.CurrentY = y1 + 3:       pb.Print Format(Trim(Book_Gun), "@@@@@@@")
'    pb.CurrentX = x1 + 130: pb.CurrentY = y1 + 3:       pb.Print Format(Trim(Book_Ho), "@@@@@@@@@@@")
'
'    pb.CurrentX = x1 + 114: pb.CurrentY = y1 + 10:      pb.Print Mid(Serial, 1, 1)
'    pb.CurrentX = x1 + 119: pb.CurrentY = y1 + 10:      pb.Print Mid(Serial, 2, 1)
'
'    pb.CurrentX = x1 + 132: pb.CurrentY = y1 + 10:       pb.Print Mid(SerialNo, 1, 1)
'    pb.CurrentX = x1 + 137: pb.CurrentY = y1 + 10:       pb.Print Mid(SerialNo, 2, 1)
'    pb.CurrentX = x1 + 143: pb.CurrentY = y1 + 10:       pb.Print Mid(SerialNo, 3, 1)
'    pb.CurrentX = x1 + 149: pb.CurrentY = y1 + 10:       pb.Print Mid(SerialNo, 4, 1)
'
'    '공급자
'    pb.CurrentX = x1 + 21.5: pb.CurrentY = y1 + 17.5:    pb.Print Format(Company_No, "000-00-00000")
'    pb.CurrentX = x1 + 20.5: pb.CurrentY = y1 + 25:      pb.Print Company_SangHo
'    pb.CurrentX = x1 + 55.5: pb.CurrentY = y1 + 25:      pb.Print Company_DaePyo
'    pb.CurrentX = x1 + 20.5: pb.CurrentY = y1 + 32:      pb.Print Company_Addr
'    pb.CurrentX = x1 + 20.5: pb.CurrentY = y1 + 35.5:    pb.Print Company_Addr2
'    If Hangul_Length(Company_UpTae) > 14 Then
'        pb.CurrentX = x1 + 20.5: pb.CurrentY = y1 + 40:      pb.Print Hangul_Mid(Company_UpTae, 1, 14)
'        pb.CurrentX = x1 + 20.5: pb.CurrentY = y1 + 43.5:    pb.Print Hangul_Mid(Company_UpTae, 15, 14)
'    Else
'        pb.CurrentX = x1 + 20.5: pb.CurrentY = y1 + 41:      pb.Print Company_UpTae
'    End If
'    If Hangul_Length(Company_UpJong) > 14 Then
'        pb.CurrentX = x1 + 47: pb.CurrentY = y1 + 40.5:        pb.Print Hangul_Mid(Company_UpJong, 1, 14)
'        pb.CurrentX = x1 + 47: pb.CurrentY = y1 + 43.5:     pb.Print Hangul_Mid(Company_UpJong, 15, 14)
'    Else
'        pb.CurrentX = x1 + 47: pb.CurrentY = y1 + 42:        pb.Print Company_UpJong
'    End If
'
'    '공급받는자
'    pb.CurrentX = x1 + 98: pb.CurrentY = y1 + 17.5:      pb.Print Mid(Kure_No, 1, 1)
'    pb.CurrentX = x1 + 103: pb.CurrentY = y1 + 17.5:      pb.Print Mid(Kure_No, 2, 1)
'    pb.CurrentX = x1 + 108: pb.CurrentY = y1 + 17.5:      pb.Print Mid(Kure_No, 3, 1)
'
'    pb.CurrentX = x1 + 117: pb.CurrentY = y1 + 17.5:      pb.Print Mid(Kure_No, 4, 1)
'    pb.CurrentX = x1 + 122: pb.CurrentY = y1 + 17.5:      pb.Print Mid(Kure_No, 5, 1)
'
'    pb.CurrentX = x1 + 130.5: pb.CurrentY = y1 + 17.5:      pb.Print Mid(Kure_No, 6, 1)
'    pb.CurrentX = x1 + 135: pb.CurrentY = y1 + 17.5:      pb.Print Mid(Kure_No, 7, 1)
'    pb.CurrentX = x1 + 140: pb.CurrentY = y1 + 17.5:      pb.Print Mid(Kure_No, 8, 1)
'    pb.CurrentX = x1 + 145: pb.CurrentY = y1 + 17.5:      pb.Print Mid(Kure_No, 9, 1)
'    pb.CurrentX = x1 + 150: pb.CurrentY = y1 + 17.5:      pb.Print Mid(Kure_No, 10, 1)
'
'
'
'    pb.CurrentX = x1 + 97: pb.CurrentY = y1 + 25:        pb.Print Kure_SangHo
'    pb.CurrentX = x1 + 135: pb.CurrentY = y1 + 25:       pb.Print Kure_DaePyo
'    pb.CurrentX = x1 + 97: pb.CurrentY = y1 + 32:        pb.Print Kure_Addr
'    pb.CurrentX = x1 + 97: pb.CurrentY = y1 + 35.5:      pb.Print Kure_Addr2
'
'    'pb.CurrentX = x1 + 97: pb.CurrentY = y1 + 41:        pb.Print Kure_UpTae
'
'    'pb.CurrentX = x1 + 124: pb.CurrentY = y1 + 40:       pb.Print Kure_UpJong
'   'pb.CurrentX = x1 + 124: pb.CurrentY = y1 + 43.5:     pb.Print "목"
'
'    If Hangul_Length(Kure_UpTae) > 14 Then
'        pb.CurrentX = x1 + 97: pb.CurrentY = y1 + 40.5:      pb.Print Hangul_Mid(Kure_UpTae, 1, 14)
'        pb.CurrentX = x1 + 97: pb.CurrentY = y1 + 43.5:    pb.Print Hangul_Mid(Kure_UpTae, 15, 14)
'    Else
'        pb.CurrentX = x1 + 97: pb.CurrentY = y1 + 41:      pb.Print Kure_UpTae
'    End If
'    If Hangul_Length(Kure_UpJong) > 14 Then
'        pb.CurrentX = x1 + 124: pb.CurrentY = y1 + 40.5:        pb.Print Hangul_Mid(Kure_UpJong, 1, 14)
'        pb.CurrentX = x1 + 124: pb.CurrentY = y1 + 43.5:     pb.Print Hangul_Mid(Kure_UpJong, 15, 14)
'    Else
'        pb.CurrentX = x1 + 124: pb.CurrentY = y1 + 42:        pb.Print Kure_UpJong
'    End If
'
'
'    '작성및 곡급가액 세액 비고 출력
'    pb.ForeColor = LineColor
'    pb.CurrentX = x1 + 5: pb.CurrentY = y1 + 48:         pb.Print "작   성"
'    pb.CurrentX = x1 + 30: pb.CurrentY = y1 + 48:        pb.Print "공      급      가      액"
'    pb.CurrentX = x1 + 97: pb.CurrentY = y1 + 48:        pb.Print "세            액"
'    pb.CurrentX = x1 + 138: pb.CurrentY = y1 + 48:       pb.Print "비   고"
'    pb.CurrentX = x1 + 2: pb.CurrentY = y1 + 52.5:         pb.Print "년"
'    pb.CurrentX = x1 + 8.5: pb.CurrentY = y1 + 52.5:         pb.Print "월"
'    pb.CurrentX = x1 + 15: pb.CurrentY = y1 + 52.5:        pb.Print "일"
'
'    pb.CurrentX = x1 + 28: pb.CurrentY = y1 + 52.5:        pb.Print "백"
'    pb.CurrentX = x1 + 33: pb.CurrentY = y1 + 52.5:        pb.Print "십"
'    pb.CurrentX = x1 + 38: pb.CurrentY = y1 + 52.5:        pb.Print "억"
'    pb.CurrentX = x1 + 43: pb.CurrentY = y1 + 52.5:        pb.Print "천"
'    pb.CurrentX = x1 + 48: pb.CurrentY = y1 + 52.5:        pb.Print "백"
'    pb.CurrentX = x1 + 53: pb.CurrentY = y1 + 52.5:        pb.Print "십"
'    pb.CurrentX = x1 + 58: pb.CurrentY = y1 + 52.5:        pb.Print "만"
'    pb.CurrentX = x1 + 63: pb.CurrentY = y1 + 52.5:        pb.Print "천"
'    pb.CurrentX = x1 + 68: pb.CurrentY = y1 + 52.5:        pb.Print "백"
'    pb.CurrentX = x1 + 73: pb.CurrentY = y1 + 52.5:        pb.Print "십"
'    pb.CurrentX = x1 + 78: pb.CurrentY = y1 + 52.5:        pb.Print "일"
'
'    pb.CurrentX = x1 + 84: pb.CurrentY = y1 + 52.5:        pb.Print "십"
'    pb.CurrentX = x1 + 89: pb.CurrentY = y1 + 52.5:        pb.Print "억"
'    pb.CurrentX = x1 + 94: pb.CurrentY = y1 + 52.5:        pb.Print "천"
'    pb.CurrentX = x1 + 99: pb.CurrentY = y1 + 52.5:        pb.Print "백"
'    pb.CurrentX = x1 + 104: pb.CurrentY = y1 + 52.5:        pb.Print "십"
'    pb.CurrentX = x1 + 109: pb.CurrentY = y1 + 52.5:        pb.Print "만"
'    pb.CurrentX = x1 + 114: pb.CurrentY = y1 + 52.5:        pb.Print "천"
'    pb.CurrentX = x1 + 119: pb.CurrentY = y1 + 52.5:        pb.Print "백"
'    pb.CurrentX = x1 + 124: pb.CurrentY = y1 + 52.5:        pb.Print "십"
'    pb.CurrentX = x1 + 129: pb.CurrentY = y1 + 52.5:        pb.Print "일"
'
'    '내용
'    pb.ForeColor = vbBlack
'    kk = 59
'    pb.CurrentX = x1 + 0.5: pb.CurrentY = y1 + kk:        pb.Print Tax_YYYY
'    pb.CurrentX = x1 + 8.5: pb.CurrentY = y1 + kk:       pb.Print Tax_MM
'    pb.CurrentX = x1 + 15: pb.CurrentY = y1 + kk:       pb.Print Tax_DD
'    pb.CurrentX = x1 + 22: pb.CurrentY = y1 + kk:        pb.Print Tax_BlankCnt
'
'
'    pb.CurrentX = x1 + 28: pb.CurrentY = y1 + kk:        pb.Print TNET(10)
'    pb.CurrentX = x1 + 33: pb.CurrentY = y1 + kk:        pb.Print TNET(9)
'    pb.CurrentX = x1 + 38: pb.CurrentY = y1 + kk:        pb.Print TNET(8)
'    pb.CurrentX = x1 + 43: pb.CurrentY = y1 + kk:        pb.Print TNET(7)
'    pb.CurrentX = x1 + 48: pb.CurrentY = y1 + kk:        pb.Print TNET(6)
'    pb.CurrentX = x1 + 53: pb.CurrentY = y1 + kk:        pb.Print TNET(5)
'    pb.CurrentX = x1 + 58: pb.CurrentY = y1 + kk:        pb.Print TNET(4)
'    pb.CurrentX = x1 + 63: pb.CurrentY = y1 + kk:        pb.Print TNET(3)
'    pb.CurrentX = x1 + 68: pb.CurrentY = y1 + kk:        pb.Print TNET(2)
'    pb.CurrentX = x1 + 73: pb.CurrentY = y1 + kk:        pb.Print TNET(1)
'    pb.CurrentX = x1 + 78: pb.CurrentY = y1 + kk:        pb.Print TNET(0)
'
'    pb.CurrentX = x1 + 84: pb.CurrentY = y1 + kk:        pb.Print TVAT(9)
'    pb.CurrentX = x1 + 89: pb.CurrentY = y1 + kk:        pb.Print TVAT(8)
'    pb.CurrentX = x1 + 94: pb.CurrentY = y1 + kk:        pb.Print TVAT(7)
'    pb.CurrentX = x1 + 99: pb.CurrentY = y1 + kk:        pb.Print TVAT(6)
'    pb.CurrentX = x1 + 104: pb.CurrentY = y1 + kk:       pb.Print TVAT(5)
'    pb.CurrentX = x1 + 109: pb.CurrentY = y1 + kk:       pb.Print TVAT(4)
'    pb.CurrentX = x1 + 114: pb.CurrentY = y1 + kk:       pb.Print TVAT(3)
'    pb.CurrentX = x1 + 119: pb.CurrentY = y1 + kk:       pb.Print TVAT(2)
'    pb.CurrentX = x1 + 124: pb.CurrentY = y1 + kk:       pb.Print TVAT(1)
'    pb.CurrentX = x1 + 129: pb.CurrentY = y1 + kk:       pb.Print TVAT(0)
'    '---------------------------------------------------------------------------
'    pb.ForeColor = LineColor
'    pb.CurrentX = x1 + 1: pb.CurrentY = y1 + 67.5:          pb.Print "월 일"
'    pb.CurrentX = x1 + 15: pb.CurrentY = y1 + 67.5:          pb.Print "품             목"
'    pb.CurrentX = x1 + 48: pb.CurrentY = y1 + 67.5:          pb.Print "규 격"
'    pb.CurrentX = x1 + 61: pb.CurrentY = y1 + 67.5:          pb.Print "수  량"
'    pb.CurrentX = x1 + 77: pb.CurrentY = y1 + 67.5:          pb.Print "단   가"
'    pb.CurrentX = x1 + 93: pb.CurrentY = y1 + 67.5:          pb.Print "공  급  가  액"
'    pb.CurrentX = x1 + 123: pb.CurrentY = y1 + 67.5:          pb.Print "세     액"
'    pb.CurrentX = x1 + 142: pb.CurrentY = y1 + 67.5:          pb.Print "비 고"
'
'    pb.ForeColor = vbBlack
'    kk = 74
'    For i = 0 To 3
'        pb.CurrentX = x1 + 1: pb.CurrentY = y1 + kk:         pb.Print Format(Tax_ListMM(i), "00")
'        pb.CurrentX = x1 + 6: pb.CurrentY = y1 + kk:         pb.Print Format(Tax_ListDD(i), "00")
'        pb.CurrentX = x1 + 12: pb.CurrentY = y1 + kk:        pb.Print Hangul_Mid(Tax_ListItem(i), 1, 20)
'        pb.CurrentX = x1 + 48: pb.CurrentY = y1 + kk:        pb.Print Tax_ListSpec(i) 'Format(Tax_ListSpec(i), "@@@@@")
'        pb.CurrentX = x1 + 61: pb.CurrentY = y1 + kk:        pb.Print Format(Tax_ListQty(i), "@@@@@@")
'        pb.CurrentX = x1 + 73.5: pb.CurrentY = y1 + kk:        pb.Print Format(Tax_ListDan(i), "@@@@@@@@@")
'        pb.CurrentX = x1 + 93: pb.CurrentY = y1 + kk:        pb.Print Format(Tax_ListNet(i), "@@@@@@@@@@@@@@")
'        pb.CurrentX = x1 + 123: pb.CurrentY = y1 + kk:       pb.Print Format(Tax_ListVat(i), "@@@@@@@@@")
'        pb.CurrentX = x1 + 141: pb.CurrentY = y1 + kk:       pb.Print Hangul_Mid(Tax_ListRem(i), 1, 6)
'        kk = kk + 7.4
'    Next i
'
'    pb.ForeColor = LineColor
'    pb.CurrentX = x1 + 3: pb.CurrentY = y1 + 101.5:           pb.Print "합 계 금 액"
'    pb.CurrentX = x1 + 30: pb.CurrentY = y1 + 101.5:          pb.Print "현   금"
'    pb.CurrentX = x1 + 53: pb.CurrentY = y1 + 101.5:          pb.Print "수   표"
'    pb.CurrentX = x1 + 75: pb.CurrentY = y1 + 101.5:          pb.Print "어   음"
'    pb.CurrentX = x1 + 96: pb.CurrentY = y1 + 101.5:          pb.Print "외상미수금"
'    pb.CurrentX = x1 + 119: pb.CurrentY = y1 + 104:           pb.Print "이 금액을      함."
'    pb.CurrentX = x1 + 135: pb.CurrentY = y1 + 102.5:         pb.Print "영수"
'    pb.CurrentX = x1 + 135: pb.CurrentY = y1 + 106.5:         pb.Print "청구"
'
'    pb.ForeColor = vbBlack
'    kk = 106
'    pb.CurrentX = x1 + 3: pb.CurrentY = y1 + kk:           pb.Print Tax_Total
'    pb.CurrentX = x1 + 30: pb.CurrentY = y1 + kk:          pb.Print Tax_Cash
'    pb.CurrentX = x1 + 53: pb.CurrentY = y1 + kk:          pb.Print Tax_SuPyo
'    pb.CurrentX = x1 + 75: pb.CurrentY = y1 + kk:          pb.Print Tax_Draft
'    pb.CurrentX = x1 + 96: pb.CurrentY = y1 + kk:          pb.Print Tax_Misu
'
'    pb.DrawWidth = 2
'    If Tax_Flag = "1" Then
'        pb.Circle (x1 + 138, y1 + 104), 3 * sf      '영수
'        'pb.CurrentX = x1 + 135: pb.CurrentY = y1 + 102.5:         pb.Print "===="
'    Else
'        pb.Circle (x1 + 138, y1 + 108), 3 * sf     '청구
'        'pb.CurrentX = x1 + 135: pb.CurrentY = y1 + 106.5:         pb.Print "===="
'    End If
'
'    pb.ForeColor = LineColor
'    If intItem = 1 Then
'        pb.DrawStyle = vbDash
'        pb.DrawWidth = 1
'        pb.Line (x1 - 10, y1 + 120)-(x1 + 160, y1 + 120)
'    End If
'pb.FontSize = 8 * sf
'    pb.CurrentX = x1 + 1: pb.CurrentY = y1 + 112:       pb.Print "22226-28131 96. 2 27 개정"
'    pb.CurrentX = x1 + 100: pb.CurrentY = y1 + 112:     pb.Print "182mm X 128mm 인쇄용지특급 34g ㎡"
'    Next
'
End Sub

