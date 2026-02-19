VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCAIDPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Dim m_P             As PRINT_DATA_TYPE
Dim m_PrtParam      As PrintParamTYPE
Private Const LPT_PORT     As String = "LPT1"

Private Const ERROR_BASE = vbObjectError + 2000
Private Const ERROR_PRINTER_NOT_FOUND = ERROR_BASE + 1
Private Const ERROR_FILE_NOT_FOUND = ERROR_BASE + 2
Private Const ERROR_ORDER = ERROR_BASE + 3


'이 이벤트를 발생시키려면 다음과 같은 구문을 가지는 RaiseEvent를 사용하십시오.
'RaiseEvent NotPrinter[(arg1, arg2, ... , argn)]
Public Event ErrorPrinter(ErrNum As Integer, ErrDec As String)
Public Event PrintPage(mode As String, PrtPage As Long, TotalPage As Long)

Public Property Get PaperSize() As Integer
    PaperSize = Printer.PaperSize
End Property


Public Function Fn_PrinterCheck() As Boolean

    On Error GoTo ERR_RTN
    Dim printer_name As String
  
    Dim x As Printer
    
    For Each x In Printers
        printer_name = x.DeviceName
    Next

    If Printer.DeviceName = "" Then
        Fn_PrinterCheck = False
        Exit Function
    End If
    
    Fn_PrinterCheck = True
    Exit Function
  
ERR_RTN:
    Fn_PrinterCheck = False

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
        
        PanelsMsg CStr(Printer.PaperSize)
    
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

'*************************************************************************************************
' 기본 결재 라인을 그린다.
Private Function Def_Approval(PView As Object, Optional DrawWidth As Integer = 2) As Integer
'                                                        ┌─┬───┬───┬───┬───┐
'                                                        │결│ 경리 │ 담당 │ 차장 │ 사장 │
'                                                        │  ├───┼───┼───┼───┤
'                                                        │재│      │      │      │      │
'                                                        └─┴───┴───┴───┴───┘

    Dim x1 As Long, y1 As Long

    x1 = 0: y1 = 0
    PView.DrawWidth = DrawWidth
    PView.DrawStyle = vbSolid
    
    PView.Line (130 + x1, 25)-(195 + x1, 43 + y1), , B
    PView.Line (135 + x1, 25)-(135 + x1, 43 + y1), , B
    PView.Line (135 + x1, 30)-(195 + x1, 30 + y1), , B
    PView.Line (150 + x1, 25)-(150 + x1, 43 + y1), , B
    PView.Line (165 + x1, 25)-(165 + x1, 43 + y1), , B
    PView.Line (180 + x1, 25)-(180 + x1, 43 + y1), , B
    
    PView.Font.Underline = False
    PView.Font.Bold = False:    PView.Font.Size = 9 * iSL_Type

    PrintText PView, m_P, 121 + x1, 20 + y1, "결"
    PrintText PView, m_P, 121 + x1, 27 + y1, "재"
    PrintText PView, m_P, 127 + x1, 16 + y1, "경  리"
    PrintText PView, m_P, 142 + x1, 16 + y1, "담  당"
    PrintText PView, m_P, 157 + x1, 16 + y1, "차  장"
    PrintText PView, m_P, 172 + x1, 16 + y1, "사  장"

End Function

'---------------------------------------------------------------------------------------------
' 레포트 하단을 출력한다.
Private Sub Def_Page_Footer(PView As Object)
    Dim lMSpace As Long
'                                                                           크린에이드
    PView.Font.Bold = True
    lMSpace = m_P.MagRect.Bottom - 20
    PrintLine PView, m_P, 0, lMSpace, m_P.MagRect.Right, lMSpace     '수평 라인
    lMSpace = m_P.MagRect.Bottom - 18
    PrintText PView, m_P, m_P.MagRect.Right - 15 - PView.TextWidth(Store.Name), lMSpace, Store.Name

End Sub

'*************************************************************************************************
' 해당 출력 양식의 모든 출력 건수를 알아온다.
Private Function Fnc_GetMaxPageCount(sCode As String, ByRef Def As PRINT_DATA_TYPE) As Integer
    Dim FHandel     As Integer
    Dim FilePath    As String
    Dim MaxCount    As Integer
    Dim MaxPage     As Integer
    Dim Str         As String
    
    Fnc_GetMaxPageCount = 0
    
    Select Case UCase(sCode)
    
        '
        Case "PRT_04001_01"
            FilePath = GetIniStr("REPORT", "TempPath", "", sIniFile) & "\Temp.txt"
            If Dir(FilePath, vbNormal) = "" Then
                RaiseEvent ErrorPrinter(ERROR_FILE_NOT_FOUND, "프린트에 필요한 파일을 찾을수 없습니다.")
                Exit Function
            End If
            
            FHandel = FreeFile
            Open FilePath For Input As #FHandel
            Do While Not EOF(FHandel)
                Line Input #FHandel, Str
                If Str <> "" Then MaxCount = MaxCount + 1
            Loop
            Close #FHandel
            
            MaxPage = Round((MaxCount / Def.PageProcMaxCnt) + 0.5)
            Def.MaxPage = MaxPage
            Fnc_GetMaxPageCount = MaxPage
            Exit Function
        Case "PRT_03001_01"
            Debug.Print PrtParam.Param(0)
            Def.MaxPage = 1
            Fnc_GetMaxPageCount = 1
        
        Case Else
        
    End Select
    
End Function

' 일일수금 종합 관리 대장
Public Function PRT_04001_01(PView As Object, ViewPage As Integer) As Integer
    Dim ViewCnt As Integer

    On Error GoTo PRT_04001_01_Error
    
    PRT_04001_01 = 0
    
    ' 용지 설정
    Call PrnRefresh(PView, iSL_Type)
    
    ' 전달된 페이지의 유효성을 확인한다.
    m_P.PageProcMaxCnt = 44         ' 한페이지당 30개 출력
    Call Fnc_GetMaxPageCount("PRT_04001_01", m_P)   ' m_p.MaxPage가 설정된다.
    If ViewPage > m_P.MaxPage Then ViewPage = m_P.MaxPage
    
    ' 모든 페이지를 인쇄할 경우
    If ViewPage = 0 And UCase(TypeName(PView)) = "PRINTER" Then
        For ViewCnt = 1 To m_P.MaxPage
            Call PRT_04001_VIEW(PView, ViewCnt)
            RaiseEvent PrintPage("PRINTER", Val(ViewCnt), Val(m_P.MaxPage))
            If ViewCnt <> m_P.MaxPage Then PView.NewPage
        Next ViewCnt
        PView.EndDoc
    Else
        Call PRT_04001_VIEW(PView, ViewPage)
        If UCase(TypeName(PView)) = "PRINTER" Then
            RaiseEvent PrintPage("PRINTER", Val(ViewCnt), Val(m_P.MaxPage))
            PView.EndDoc
        Else
            RaiseEvent PrintPage("SCREEN", Val(ViewCnt), Val(m_P.MaxPage))
        End If
    End If
    Exit Function
    
PRT_04001_01_Error:
    PView.EndDoc
    RaiseEvent ErrorPrinter(ERROR_FILE_NOT_FOUND, Err.Description)

End Function

Private Function PRT_04001_VIEW(PView As Object, ViewPage As Integer) As Integer
' 일일수금 종합 관리 대장
    
    On Error GoTo PRT_04001_VIEW_Error
    
    Dim vData       As Variant
    Dim FHandel     As Integer
    Dim Str         As String
    Dim lMSpace     As Long
    Dim TempLng     As Long
    
    PRT_04001_VIEW = 0

    ' 출력하기 위한 기본 자료를 가저온다.
    GoSub SUB_INIT_VAL
    
    GoSub SUB_PAGE_HEADER
    
    GoSub SUB_DETAILS_PRINTER
    
    GoSub SUB_REPORT_FOOTER
    
    Call Def_Page_Footer(PView)
    Exit Function
    
'---------------------------------------------------------------------------------------------
' 출력하기 위한 기본 자료를 가저온다.
SUB_INIT_VAL:

    m_P.MagRect.Left = 10:          m_P.MagRect.Top = 10
    m_P.MagRect.Right = 200:        m_P.MagRect.Bottom = 285
    
    m_P.MaxCount = 0             ' 출력할 전체 건수
    m_P.ProcCnt = 0             ' 현재출력 중인 건수
    m_P.PageProcCnt = 0         ' 한페이지에서 Details의 현재 출력중인 수
    
    m_P.NextLineSpace = 5       ' Details의 다음 라인에 출력될 여백
    m_P.ProcPageCnt = ViewPage
    
    If iSL_Type = 0 Then iSL_Type = 1
    Return

'---------------------------------------------------------------------------------------------
' 페이지 상단을 출력한다.
SUB_PAGE_HEADER:
'                                    일일수금 종합 관리 대장
'                                 -----------------------------
'                                                        ┌─┬───┬───┬───┬───┐
'                                                        │확│ 경리 │ 담당 │ 차장 │ 사장 │
'                                                        │  ├───┼───┼───┼───┤
'                                                        │인│      │      │      │      │
'   수금일자:                   날씨:                    └─┴───┴───┴───┴───┘
'   ------------------------------------------------------------------------------------------
'   순위 대리점명       수금액   입고량   출고량    단가    재세탁    수선   반품    택번호
'   ------------------------------------------------------------------------------------------
    
    PView.FontName = "굴림체"
    PView.Font.Bold = True:               PView.Font.Underline = True
    PView.Font.Size = 16 * iSL_Type
    PrintText PView, m_P, 60, 5, " 일일수금 종합 관리 대장 "
    
    PView.Font.Underline = False
    PView.Font.Bold = True:    PView.Font.Size = 9 * iSL_Type
    PrintText PView, m_P, 2, 30, PrtParam.Param(1)
    PrintText PView, m_P, 50, 30, PrtParam.Param(2)
    
    Call Def_Approval(PView)
    
    PView.Font.Bold = False:    PView.Font.Size = 9
    PrintLine PView, m_P, 0, 35, 0, 35 '수평 라인
    PrintText PView, m_P, 2, 37, "순위"
    PrintText PView, m_P, 13, 37, "대리점명"
    PrintText PView, m_P, 50, 37, "수금액"
    PrintText PView, m_P, 65, 37, "입고량"
    PrintText PView, m_P, 80, 37, "출고량"
    PrintText PView, m_P, 100, 37, "단가"
    PrintText PView, m_P, 115, 37, "재세탁"
    PrintText PView, m_P, 132, 37, "수선"
    PrintText PView, m_P, 145, 37, "반품"
    PrintText PView, m_P, 160, 37, "택번호"
    PrintLine PView, m_P, 0, 42, 0, 42 '수평 라인
    Return
    
'---------------------------------------------------------------------------------------------
' 레포트 하단을 출력한다.
SUB_REPORT_FOOTER:
'   ------------------------------------------------------------------------------------------
'   수금액 :        단  가:         월수금:         일수금:
'   입고량 :        출고량:         재세탁:         수  선:         반  품:
'   ------------------------------------------------------------------------------------------
    
    ' 마지막장에만 출력한다.
    If m_P.MaxPage <> m_P.ProcPageCnt Then Return
    
    lMSpace = (m_P.MagRect.Top + 47) + ((m_P.PageProcCnt + 1) * m_P.NextLineSpace)
    PrintLine PView, m_P, 0, lMSpace, 0, lMSpace '수평 라인
    
    lMSpace = lMSpace + 3
    PrintText PView, m_P, 5, lMSpace, PrtParam.Param(3)
    PrintText PView, m_P, 50, lMSpace, PrtParam.Param(4)
    PrintText PView, m_P, 90, lMSpace, PrtParam.Param(5)
    PrintText PView, m_P, 135, lMSpace, PrtParam.Param(6)
    
    lMSpace = lMSpace + 5
    PrintText PView, m_P, 5, lMSpace, PrtParam.Param(7)
    PrintText PView, m_P, 40, lMSpace, PrtParam.Param(8)
    PrintText PView, m_P, 70, lMSpace, PrtParam.Param(9)
    PrintText PView, m_P, 100, lMSpace, PrtParam.Param(10)
    PrintText PView, m_P, 140, lMSpace, PrtParam.Param(11)
    
    lMSpace = lMSpace + m_P.NextLineSpace
    PrintLine PView, m_P, 0, lMSpace, 0, lMSpace '수평 라인
    Return
    
    
SUB_DETAILS_PRINTER:
    ' 내용을 출력한다.
    
    FHandel = FreeFile
    Open GetIniStr("REPORT", "TempPath", "", sIniFile) & "\Temp.txt" For Input As #FHandel
    Do While Not EOF(FHandel)
    
        Line Input #FHandel, Str
        vData = Split(Str, "|")
        If UBound(vData) = 11 Then
            
            ' 출력중인 전체 수
            m_P.ProcCnt = m_P.ProcCnt + 1
            
            If m_P.ProcCnt > ((m_P.ProcPageCnt - 1) * m_P.PageProcMaxCnt) Then
                
                ' Details 출력
                TempLng = 45
                PrintText PView, m_P, 2, TempLng + (m_P.PageProcCnt * m_P.NextLineSpace), CStr(vData(0))
                PrintText PView, m_P, 13, TempLng + (m_P.PageProcCnt * m_P.NextLineSpace), CStr(vData(1)) & CStr(vData(2))
                PrintText PView, m_P, 42, TempLng + (m_P.PageProcCnt * m_P.NextLineSpace), CStr(vData(3))
                PrintText PView, m_P, 63, TempLng + (m_P.PageProcCnt * m_P.NextLineSpace), CStr(vData(4))
                PrintText PView, m_P, 76, TempLng + (m_P.PageProcCnt * m_P.NextLineSpace), CStr(vData(5))
                PrintText PView, m_P, 88, TempLng + (m_P.PageProcCnt * m_P.NextLineSpace), CStr(vData(6))
                PrintText PView, m_P, 112, TempLng + (m_P.PageProcCnt * m_P.NextLineSpace), CStr(vData(7))
                PrintText PView, m_P, 125, TempLng + (m_P.PageProcCnt * m_P.NextLineSpace), CStr(vData(8))
                PrintText PView, m_P, 138, TempLng + (m_P.PageProcCnt * m_P.NextLineSpace), CStr(vData(9))
                PrintText PView, m_P, 160, TempLng + (m_P.PageProcCnt * m_P.NextLineSpace), CStr(vData(10))
    
                m_P.PageProcCnt = m_P.PageProcCnt + 1
                If m_P.PageProcCnt >= m_P.PageProcMaxCnt Then Exit Do
            End If
            
        End If
        
    Loop
    
    If m_P.PageProcCnt < m_P.PageProcMaxCnt Then
    PrintLine PView, m_P, 0, 41 + ((m_P.PageProcCnt + 1) * m_P.NextLineSpace), _
                          0, 41 + ((m_P.PageProcCnt + 1) * m_P.NextLineSpace)   '수평 라인
    End If
    Close #FHandel
    Return
    

PRT_04001_VIEW_Error:
    Resume
    Close #FHandel
    m_Error.ErrorMsg = "Error " & Err.Number & " (" & Err.Description & ") in procedure PRT_04001_VIEW of Form CCAIDPrinter"

    If m_Error.VisibleMSG Then
        MsgBox m_Error.ErrorMsg, vbInformation, "확인 [Error]"
    End If

    If m_Error.SaveLog Then
        Call ProgramErrorLogWrite(m_Error)
    End If

    If m_Error.ResumeMode Then
        Resume
    End If


End Function


' 일일수금 종합 관리 대장
Public Function PRT_04001_01_MASTER(PView As Object, ViewPage As Integer) As Integer
    Dim ViewCnt As Integer

    On Error GoTo PRT_04001_01_Error
    
    PRT_04001_01_MASTER = 0
    
    ' 용지 설정
    Call PrnRefresh(PView, iSL_Type)
    
    ' 전달된 페이지의 유효성을 확인한다.
    m_P.PageProcMaxCnt = 44         ' 한페이지당 30개 출력
    Call Fnc_GetMaxPageCount("PRT_04001_01", m_P)   ' m_p.MaxPage가 설정된다.
    If ViewPage > m_P.MaxPage Then ViewPage = m_P.MaxPage
    
    ' 모든 페이지를 인쇄할 경우
    If ViewPage = 0 And UCase(TypeName(PView)) = "PRINTER" Then
        For ViewCnt = 1 To m_P.MaxPage
            Call PRT_04001_MASTER_VIEW(PView, ViewCnt)
            RaiseEvent PrintPage("PRINTER", Val(ViewCnt), Val(m_P.MaxPage))
            If ViewCnt <> m_P.MaxPage Then PView.NewPage
        Next ViewCnt
        PView.EndDoc
    Else
        Call PRT_04001_MASTER_VIEW(PView, ViewPage)
        If UCase(TypeName(PView)) = "PRINTER" Then
            RaiseEvent PrintPage("PRINTER", Val(ViewCnt), Val(m_P.MaxPage))
            PView.EndDoc
        Else
            RaiseEvent PrintPage("SCREEN", Val(ViewCnt), Val(m_P.MaxPage))
        End If
    End If
    Exit Function
    
PRT_04001_01_Error:
    PView.EndDoc
    RaiseEvent ErrorPrinter(ERROR_FILE_NOT_FOUND, Err.Description)

End Function

Private Function PRT_04001_MASTER_VIEW(PView As Object, ViewPage As Integer) As Integer
' 일일수금 종합 관리 대장
    
    On Error GoTo PRT_04001_VIEW_Error
    
    Dim vData       As Variant
    Dim FHandel     As Integer
    Dim Str         As String
    Dim lMSpace     As Long
    Dim TempLng     As Long
    
    PRT_04001_MASTER_VIEW = 0

    ' 출력하기 위한 기본 자료를 가저온다.
    GoSub SUB_INIT_VAL
    
    GoSub SUB_PAGE_HEADER
    
    GoSub SUB_DETAILS_PRINTER
    
    GoSub SUB_REPORT_FOOTER
    
    Call Def_Page_Footer(PView)
    Exit Function
    
'---------------------------------------------------------------------------------------------
' 출력하기 위한 기본 자료를 가저온다.
SUB_INIT_VAL:

    m_P.MagRect.Left = 10:          m_P.MagRect.Top = 10
    m_P.MagRect.Right = 200:        m_P.MagRect.Bottom = 285
    
    m_P.MaxCount = 0             ' 출력할 전체 건수
    m_P.ProcCnt = 0             ' 현재출력 중인 건수
    m_P.PageProcCnt = 0         ' 한페이지에서 Details의 현재 출력중인 수
    
    m_P.NextLineSpace = 5       ' Details의 다음 라인에 출력될 여백
    m_P.ProcPageCnt = ViewPage
    
    If iSL_Type = 0 Then iSL_Type = 1
    Return

'---------------------------------------------------------------------------------------------
' 페이지 상단을 출력한다.
SUB_PAGE_HEADER:
'                                    일일수금 종합 관리 대장
'                                 -----------------------------
'                                                        ┌─┬───┬───┬───┬───┐
'                                                        │확│ 경리 │ 담당 │ 차장 │ 사장 │
'                                                        │  ├───┼───┼───┼───┤
'                                                        │인│      │      │      │      │
'   수금일자:                   날씨:                    └─┴───┴───┴───┴───┘
'   ------------------------------------------------------------------------------------------
'   순위 대리점명       수금액   입고량   출고량    단가    재세탁    수선   반품    택번호
'   ------------------------------------------------------------------------------------------
    
    PView.FontName = "굴림체"
    PView.Font.Bold = True:               PView.Font.Underline = True
    PView.Font.Size = 16 * iSL_Type
    PrintText PView, m_P, 60, 5, " 일일수금 종합 관리 대장 "
    
    PView.Font.Underline = False
    PView.Font.Bold = True:    PView.Font.Size = 9 * iSL_Type
    PrintText PView, m_P, 2, 30, PrtParam.Param(1)
    PrintText PView, m_P, 50, 30, PrtParam.Param(2)
    
    Call Def_Approval(PView)
    
    PView.Font.Bold = False:    PView.Font.Size = 9
    PrintLine PView, m_P, 0, 35, 180, 35 '수평 라인
    PrintText PView, m_P, 2, 37, "순위"
    PrintText PView, m_P, 13, 37, "대리점명"
    PrintText PView, m_P, 50, 37, "수금액"
    PrintText PView, m_P, 65, 37, "입고량"
    PrintText PView, m_P, 80, 37, "출고량"
    PrintText PView, m_P, 100, 37, "단가"
    PrintText PView, m_P, 115, 37, "재세탁"
    PrintText PView, m_P, 132, 37, "카드금액"
    PrintText PView, m_P, 148, 37, "반품"
    PrintText PView, m_P, 160, 37, "택번호"
    PrintLine PView, m_P, 0, 42, 180, 42 '수평 라인
    Return
    
'---------------------------------------------------------------------------------------------
' 레포트 하단을 출력한다.
SUB_REPORT_FOOTER:
'   ------------------------------------------------------------------------------------------
'   수금액 :        단  가:         월수금:         일수금:
'   입고량 :        출고량:         재세탁:         수  선:         반  품:
'   ------------------------------------------------------------------------------------------
    
    ' 마지막장에만 출력한다.
    If m_P.MaxPage <> m_P.ProcPageCnt Then Return
    
    lMSpace = (m_P.MagRect.Top + 47) + ((m_P.PageProcCnt + 1) * m_P.NextLineSpace)
    PrintLine PView, m_P, 0, lMSpace, 180, lMSpace '수평 라인
    
    lMSpace = lMSpace + 3
    PrintText PView, m_P, 5, lMSpace, PrtParam.Param(3)
    PrintText PView, m_P, 40, lMSpace, PrtParam.Param(4)
    PrintText PView, m_P, 70, lMSpace, PrtParam.Param(5)
    PrintText PView, m_P, 110, lMSpace, PrtParam.Param(6)
    PrintText PView, m_P, 150, lMSpace, PrtParam.Param(12)
    
    lMSpace = lMSpace + 5
    PrintText PView, m_P, 5, lMSpace, PrtParam.Param(7)
    PrintText PView, m_P, 40, lMSpace, PrtParam.Param(8)
    PrintText PView, m_P, 70, lMSpace, PrtParam.Param(9)
    PrintText PView, m_P, 100, lMSpace, PrtParam.Param(10)
    PrintText PView, m_P, 140, lMSpace, PrtParam.Param(11)
    
    lMSpace = lMSpace + m_P.NextLineSpace
    PrintLine PView, m_P, 0, lMSpace, 180, lMSpace '수평 라인
    Return
    
    
SUB_DETAILS_PRINTER:
    ' 내용을 출력한다.
    
    FHandel = FreeFile
    Open GetIniStr("REPORT", "TempPath", "", sIniFile) & "\Temp.txt" For Input As #FHandel
    Do While Not EOF(FHandel)
    
        Line Input #FHandel, Str
        vData = Split(Str, "|")
        If UBound(vData) = 13 Then
            
            ' 출력중인 전체 수
            m_P.ProcCnt = m_P.ProcCnt + 1
            
            If m_P.ProcCnt > ((m_P.ProcPageCnt - 1) * m_P.PageProcMaxCnt) Then
                
                ' Details 출력
                TempLng = 45
                PrintText PView, m_P, 2, TempLng + (m_P.PageProcCnt * m_P.NextLineSpace), CStr(vData(0))
                PrintText PView, m_P, 13, TempLng + (m_P.PageProcCnt * m_P.NextLineSpace), CStr(vData(1)) & CStr(vData(2))
                PrintText PView, m_P, 42, TempLng + (m_P.PageProcCnt * m_P.NextLineSpace), CStr(vData(3))
                PrintText PView, m_P, 63, TempLng + (m_P.PageProcCnt * m_P.NextLineSpace), CStr(vData(4))
                PrintText PView, m_P, 76, TempLng + (m_P.PageProcCnt * m_P.NextLineSpace), CStr(vData(5))
                PrintText PView, m_P, 88, TempLng + (m_P.PageProcCnt * m_P.NextLineSpace), CStr(vData(6))
                PrintText PView, m_P, 112, TempLng + (m_P.PageProcCnt * m_P.NextLineSpace), CStr(vData(9))
                PrintText PView, m_P, 130, TempLng + (m_P.PageProcCnt * m_P.NextLineSpace), CStr(vData(7))
                PrintText PView, m_P, 142, TempLng + (m_P.PageProcCnt * m_P.NextLineSpace), CStr(vData(11))
                PrintText PView, m_P, 160, TempLng + (m_P.PageProcCnt * m_P.NextLineSpace), CStr(vData(12))
    
                m_P.PageProcCnt = m_P.PageProcCnt + 1
                If m_P.PageProcCnt >= m_P.PageProcMaxCnt Then Exit Do
            End If
            
        End If
        
    Loop
    
    If m_P.PageProcCnt < m_P.PageProcMaxCnt Then
    PrintLine PView, m_P, 0, 41 + ((m_P.PageProcCnt + 1) * m_P.NextLineSpace), _
                          0, 41 + ((m_P.PageProcCnt + 1) * m_P.NextLineSpace)   '수평 라인
    End If
    Close #FHandel
    Return
    

PRT_04001_VIEW_Error:
    Resume
    Close #FHandel
    m_Error.ErrorMsg = "Error " & Err.Number & " (" & Err.Description & ") in procedure PRT_04001_VIEW of Form CCAIDPrinter"

    If m_Error.VisibleMSG Then
        MsgBox m_Error.ErrorMsg, vbInformation, "확인 [Error]"
    End If

    If m_Error.SaveLog Then
        Call ProgramErrorLogWrite(m_Error)
    End If

    If m_Error.ResumeMode Then
        Resume
    End If


End Function

' 일일 출고 현황(윈도우 방식)
Public Function PRT_03001_01_WIN(PView As Object, ViewPage As Integer) As Integer
    Dim ViewCnt As Integer

    On Error GoTo PRT_03001_01_Error
    
    PRT_03001_01_WIN = 0
    
    ' 용지 설정
    'Call PrnRefresh(PView, iSL_Type)
    
    ' 전달된 페이지의 유효성을 확인한다.
    m_P.PageProcMaxCnt = 44         ' 한페이지당 30개 출력
    Call Fnc_GetMaxPageCount("PRT_03001_01", m_P)   ' m_p.MaxPage가 설정된다.
    If ViewPage > m_P.MaxPage Then ViewPage = m_P.MaxPage
    
    ' 모든 페이지를 인쇄할 경우
    If ViewPage = 0 And UCase(TypeName(PView)) = "PRINTER" Then
        For ViewCnt = 1 To m_P.MaxPage
            Call PRT_03001_VIEW(PView, ViewCnt)
            RaiseEvent PrintPage("PRINTER", Val(ViewCnt), Val(m_P.MaxPage))
            If ViewCnt <> m_P.MaxPage Then PView.NewPage
        Next ViewCnt
        PView.EndDoc
    Else
        Call PRT_03001_VIEW(PView, ViewPage)
        If UCase(TypeName(PView)) = "PRINTER" Then
            RaiseEvent PrintPage("PRINTER", Val(ViewCnt), Val(m_P.MaxPage))
            PView.EndDoc
        Else
            RaiseEvent PrintPage("SCREEN", Val(ViewCnt), Val(m_P.MaxPage))
        End If
    End If
    Exit Function
    
PRT_03001_01_Error:
    PView.EndDoc
    RaiseEvent ErrorPrinter(ERROR_FILE_NOT_FOUND, Err.Description)

End Function

' 일일 출고 현황 (도트 방식)
Public Function PRT_03001_01_DOT(PView As Object, ViewPage As Integer) As Integer
    Dim ViewCnt As Integer

    On Error GoTo PRT_03001_01_Error
    
    PRT_03001_01_DOT = 0
    
    ' 용지 설정
    'Call PrnRefresh(PView, iSL_Type)
    
    ' 전달된 페이지의 유효성을 확인한다.
    m_P.PageProcMaxCnt = 44         ' 한페이지당 30개 출력
    Call Fnc_GetMaxPageCount("PRT_03001_01", m_P)   ' m_p.MaxPage가 설정된다.
    If ViewPage > m_P.MaxPage Then ViewPage = m_P.MaxPage
    
    ' 모든 페이지를 인쇄할 경우
    If ViewPage = 0 And UCase(TypeName(PView)) = "PRINTER" Then
        For ViewCnt = 1 To m_P.MaxPage
            Call PRT_03001_DOT_VIEW(PView, ViewCnt)
            RaiseEvent PrintPage("PRINTER", Val(ViewCnt), Val(m_P.MaxPage))
        Next ViewCnt
    
    ' 미리보기는 윈도우 방식과 동일하게 보여준다.
    Else
        Call PRT_03001_VIEW(PView, ViewPage)
        If UCase(TypeName(PView)) = "PRINTER" Then
            RaiseEvent PrintPage("PRINTER", Val(ViewCnt), Val(m_P.MaxPage))
            PView.EndDoc
        Else
            RaiseEvent PrintPage("SCREEN", Val(ViewCnt), Val(m_P.MaxPage))
        End If
    End If
    Exit Function
    
PRT_03001_01_Error:
    PView.EndDoc
    RaiseEvent ErrorPrinter(ERROR_FILE_NOT_FOUND, Err.Description)

End Function

Public Function PRT_03001_VIEW(PView As Object, ViewPage As Integer) As Boolean
' 일일 출고 현황의 출고증을 인쇄 한다.
    
    ' 기본 프린터가 없을 경우
    On Error GoTo PRT_P03001_Error
    
    Dim StoreCode As String     ' 현재 출력중인 대리점 코드
    Dim StoreCount As Integer   ' 현재 출력중인 대리점의 총수 ( 택수)
    Dim StoreSoCnt As Integer
    Dim StoreJaeCnt As Integer
    Dim StoreBanCnt As Integer
    Dim StoreSuCnt As Integer
    Dim StoreLostCnt As Integer
    Dim SelectCount()  As String   '선택된 대리점 정보
    Dim SelectCode()  As String   '선택된 대리점 정보
    Dim i           As Integer
    Dim iSelectCode As Integer

    PRT_03001_VIEW = False
    If Not Fn_PrinterCheck Then Exit Function
        
    ' 출력하기 위한 기본 자료를 가저온다.
    GoSub SUB_INIT_VAL
    
    ' 내용을 출력한다.
    GoSub SUB_PRINTER
    
    ' 마무리 한다.
    GoSub SUB_END_PRINT
    
    Exit Function
    
'---------------------------------------------------------------------------------------------
' 출력하기 위한 기본 자료를 가저온다.
SUB_INIT_VAL:

    m_P.MagRect.Left = 15
    m_P.MagRect.Top = 15
    m_P.MagRect.Right = 180
    m_P.MagRect.Bottom = 280
    
    iLeftLooper = 0     ' 한 라인에 여러개 출력될경우 현재 출력 갯수
    iLeftMaxLooper = 8  ' 한 라인에 출력할 전체 갯수
    iLeftNextSpace = ((m_P.MagRect.Right - m_P.MagRect.Left) / 7.5) ' 한 라인에 여러개 출력될경우 다음 출력 위치
    
    iTotalCnt = 0       ' 출력할 전체 건수
    iProcCnt = 0        ' 현재출력 중인 건수
    iPageTotProcCnt = 25 ' 한페이지당 전체 출력수 ( Details 수 )
    iPageProcCnt = 0    ' 한페이지에서 Details의 현재 출력중인 수
    
    iProcPageCnt = 0    ' 현재 출력중인 페이지 수
    iNextLineSpace = 8  ' Details의 다음 라인에 출력될 여백
    
    ' 선택된 대리점명을 구한다. ( 대리점명 ^ 택수량| )
    ReDim SelectCount(UBound(Split(PrtParam.Param(5), "|")))
    SelectCount = Split(PrtParam.Param(5), "|")
    
    ReDim SelectCode(UBound(Split(PrtParam.Param(6), "|")))
    SelectCode = Split(PrtParam.Param(6), "|")
    
    If UCase(TypeName(PView)) = "PICTUREBOX" Then
        Call PrnRefresh(PView, iSL_Type)
        PView.Cls
        
    ElseIf UCase(TypeName(PView)) = "PRINTER" Then
        Call PrnRefresh(PView, iSL_Type)

    End If

    Return

    
'---------------------------------------------------------------------------------------------
' 페이지 상단을 출력한다.
SUP_PAGE_HEADER:
'                                                                   ┌─┬────┬───┐
'      가락시영점(038) 물품 명세서                                  │확│ 본  사 │      │
'   ---------------------------------                               │  ├────┼───┤
'                                                                   │인│ 대리점 │      │
'                                                                   └─┴────┴───┘
'   2004년 11월 30일                                                              PAGE:     1
'   ------------------------------------------------------------------------------------------
'   택번호      택번호      택번호      택번호      택번호      택번호      택번호      택번호
'   ------------------------------------------------------------------------------------------
    
    Dim iLooper As Integer

    PView.FontName = "굴림체"
    PView.Font.Bold = True:               PView.Font.Underline = True
    PView.Font.Size = 20 * iSL_Type
    PrintText PView, m_P, 0, 5, Trim(RS01!AgencyName) & " (" & RS01!AgencyCode & ") 물품명세서"
    
    PrintRect PView, m_P, 115, 0, 180, 22
    PrintLine PView, m_P, 125, 0, 125, 22
    PrintLine PView, m_P, 145, 0, 145, 22
    PrintLine PView, m_P, 125, 11, 180, 11
    
    PView.Font.Bold = False:    PView.Font.Size = 10.5
    PView.Font.Underline = False
    PrintText PView, m_P, 118, 4, "확"
    PrintText PView, m_P, 118, 15, "인"
    PrintText PView, m_P, 128, 4, "본    사"
    PrintText PView, m_P, 128, 15, "대 리 점"
    
    PView.Font.Underline = False
    PView.Font.Bold = True:    PView.Font.Size = 14 * iSL_Type
    PrintText PView, m_P, 2, 28, Format(PrtParam.Param(2), "@@@@년 @@월 @@일")
    
    PrintText PView, m_P, 150, 28, "PAGE:" & Format(iProcPageCnt, "@@@@@")
    PrintLine PView, m_P, 0, 35, 180, 35  '수평 라인
    PrintLine PView, m_P, 0, 44, 180, 44    '수평 라인
    
    PView.Font.Bold = False:    PView.Font.Size = 14
    For iLooper = 1 To iLeftMaxLooper
        PrintText PView, m_P, (iLeftNextSpace * (iLooper - 1)) + 2, 38, "택번호"
    Next iLooper
    
    Return
    
'---------------------------------------------------------------------------------------------
' 레포트 하단을 출력한다.
SUB_REPORT_FOOTER:
'   ------------------------------------------------------------------------------------------
'   소품:    7   재세탁:    2   반품    0   수선:   0   택분실: 0   합계    37
'   ------------------------------------------------------------------------------------------
    Dim lMSpace As Long
    
    PView.Font.Bold = True:    PView.Font.Size = 12.5
    lMSpace = 47 + ((iPageProcCnt + 1) * iNextLineSpace)
    PrintLine PView, m_P, 0, lMSpace, 0, lMSpace
    lMSpace = lMSpace + 3
    PrintText PView, m_P, 0, lMSpace, "소품:" & Format(StoreSoCnt, "@@@@")
    PrintText PView, m_P, 30, lMSpace, "재세탁:" & Format(StoreJaeCnt, "@@@@")
    PrintText PView, m_P, 60, lMSpace, "반품:" & Format(StoreBanCnt, "@@@@")
    PrintText PView, m_P, 90, lMSpace, "수선:" & Format(StoreSuCnt, "@@@@")
    PrintText PView, m_P, 120, lMSpace, "택분실:" & Format(StoreLostCnt, "@@@@")
    PrintText PView, m_P, 160, lMSpace, "합계:" & Format(StoreCount, "@@@@")
    lMSpace = lMSpace + iNextLineSpace
    PrintLine PView, m_P, 0, lMSpace, 0, lMSpace
    
    StoreCount = 0:     StoreSoCnt = 0
    StoreJaeCnt = 0:    StoreBanCnt = 0
    StoreSuCnt = 0:     StoreLostCnt = 0

    Return
    
    
SUB_PRINTER:
    ' 내용을 출력한다.
    Dim bForExit    As Boolean
    ReDim sValue(1) As String
    
    For iSelectCode = 0 To UBound(SelectCode())
    
        sValue(0) = PrtParam.Param(2)
        sValue(1) = "(A.AgencyCode = '" & SelectCode(iSelectCode) & "') "
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("ChulgoTag_New", sValue(), Err_Num, Err_Dec)
        If Err_Num <> 0 Then
            If UCase(TypeName(PView)) = "PRINTER" Then PView.EndDoc
            RaiseEvent ErrorPrinter("1", Err_Dec)
            Exit Function
        End If

        
        iLeftNextSpace = ((m_P.MagRect.Right - m_P.MagRect.Left) / 7.5)
        
        Do While Not RS01.EOF
            
            If RS01!SendChk = "0" Then
                        
                iLeftLooper = iLeftLooper + 1
                
                ' Page Header 출력
                If (iPageProcCnt Mod iPageTotProcCnt) = 0 And iLeftLooper = 1 Then
                    ' 현재 출력중인 페이지 수
                    iProcPageCnt = iProcPageCnt + 1
                    ' 한페이지당 현재 출려 라인 수
                    iPageProcCnt = 0
                    StoreCode = RS01!AgencyCode
                    
                    GoSub SUP_PAGE_HEADER
                    
                End If
                    
                ' 출력중인 전체 수
                iProcCnt = iProcCnt + 1
                StoreCount = StoreCount + 1:    StoreSoCnt = StoreSoCnt + Val(RS01!SoCnt)
                StoreJaeCnt = StoreJaeCnt + Val(RS01!JaeCnt):  StoreBanCnt = StoreBanCnt + Val(RS01!BanCnt)
                StoreSuCnt = StoreSuCnt + Val(RS01!SuCnt):    StoreLostCnt = StoreLostCnt + Val(RS01!LostCnt)
                
                ' Details 출력
                PView.Font.Bold = False:    PView.Font.Size = 14
                PrintText PView, m_P, (iLeftNextSpace * (iLeftLooper - 1)), _
                                 47 + ((iPageProcCnt) * iNextLineSpace), RS01!Tag
                                
                ' "0"일의 값으로 마지막일 경우 모든 처리를 하고 빠저 나간다.
                If RS01.EOF Then bForExit = True
                If bForExit Then
                    PrintLine PView, m_P, 0, 47 + ((iPageProcCnt + 1) * iNextLineSpace), _
                                     0, 47 + ((iPageProcCnt + 1) * iNextLineSpace)
                    ' Report Footer 출력
                    GoSub SUB_REPORT_FOOTER
                    Call Def_Page_Footer(PView)
                    Exit Do
                End If
                
                
'    '           If Trim(RS01!Tag) = "9-867" Then MsgBox RS01!Tag
                ' 1 Page 이상일경우 마지막 부분에 라인을 표시한다.
                If ((iPageProcCnt + 1) Mod iPageTotProcCnt) = 0 And iLeftLooper = iLeftMaxLooper Then
                    PrintLine PView, m_P, 0, 47 + ((iPageProcCnt + 1) * iNextLineSpace), _
                                     m_P.MagRect.Right, 47 + ((iPageProcCnt + 1) * iNextLineSpace)

                    Call Def_Page_Footer(PView)

                    If UCase(TypeName(PView)) = "PICTUREBOX" Then
                        bForExit = True:    If bForExit Then Exit For

                    Else
                        PView.NewPage
                    End If

                End If

                ' 라인을 증가 한다.
                If iLeftLooper = iLeftMaxLooper Then
                    iLeftLooper = 0
                    iPageProcCnt = iPageProcCnt + 1
                End If
            End If
            
            RS01.MoveNext

            If bForExit Then Exit Do
        Loop
        
        ' 하나의 대리점 출력이 종료 되었음.
        ' bForExit 값이 True일경우는 "0"으로 출력 되었기 때문에 출력하면 안된다.
        If Not bForExit Then
            PrintLine PView, m_P, 0, 47 + ((iPageProcCnt + 1) * iNextLineSpace), _
                             m_P.MagRect.Right, 47 + ((iPageProcCnt + 1) * iNextLineSpace)
            ' Report Footer 출력
            GoSub SUB_REPORT_FOOTER
            Call Def_Page_Footer(PView)
        End If
            
        If UCase(TypeName(PView)) = "PICTUREBOX" Then
            bForExit = True:    If bForExit Then Exit For

        Else
            PView.NewPage
            iLeftLooper = 0     ' 한 라인에 여러개 출력될경우 현재 출력 갯수
            
            iTotalCnt = 0       ' 출력할 전체 건수
            iProcCnt = 0        ' 현재출력 중인 건수
            iPageProcCnt = 0    ' 한페이지에서 Details의 현재 출력중인 수
            
            iProcPageCnt = 0    ' 현재 출력중인 페이지 수
            StoreCount = 0      ' 한 거래처당 출려수
            
        End If
        
    Next iSelectCode
    Return
    
' 마무리 한다.
SUB_END_PRINT:
    RS01.Close
    Return

    On Error GoTo 0
    Exit Function

PRT_P03001_Error:

    m_Error.ErrorMsg = "Error " & Err.Number & " (" & Err.Description & ") in procedure PRT_03001_VIEW of Form CCAIDPrinter"

    If m_Error.VisibleMSG Then
        MsgBox m_Error.ErrorMsg, vbInformation, "확인 [Error]"
    End If

    If m_Error.SaveLog Then
        Call ProgramErrorLogWrite(m_Error)
    End If

    If m_Error.ResumeMode Then
        Resume
    End If


End Function


Public Function PRT_03001_DOT_VIEW(PView As Object, ViewPage As Integer) As Boolean
' 일일 출고 현황의 출고증을 인쇄 한다.
    
    ' 기본 프린터가 없을 경우
    On Error GoTo PRT_P03001_Error
    
    Dim StoreCode As String     ' 현재 출력중인 대리점 코드
    Dim StoreCount As Integer   ' 현재 출력중인 대리점의 총수 ( 택수)
    Dim StoreSoCnt As Integer
    Dim StoreJaeCnt As Integer
    Dim StoreBanCnt As Integer
    Dim StoreSuCnt As Integer
    Dim StoreLostCnt As Integer
    Dim SelectCount()  As String   '선택된 대리점 정보
    Dim SelectCode()  As String   '선택된 대리점 정보
    Dim i           As Integer
    Dim iSelectCode As Integer
    Dim PRT_TAG_DATA    As String
    Dim Findex As Integer
    
    Dim PRT_L   As Integer
    Dim PRT_LINE_NUMBER As Integer
    Dim PRT_LINE_MAX_NUMBER As Integer
    PRT_LINE_MAX_NUMBER = 66
    PRT_L = 5                   ' 기본 좌측 여백
    
    Dim LINE_DISPLAY As String
    LINE_DISPLAY = Space(PRT_L) & "------------------------------------------------------------------------------------------"

    PRT_03001_DOT_VIEW = False
    If Not Fn_PrinterCheck Then Exit Function
    
    ' 출력하기 위한 기본 자료를 가저온다.
    GoSub SUB_INIT_VAL
    
    ' 내용을 출력한다.
    GoSub SUB_PRINTER
    
    ' 마무리 한다.
    GoSub SUB_END_PRINT
    
    Close #Findex
    Exit Function
    
'---------------------------------------------------------------------------------------------
' 출력하기 위한 기본 자료를 가저온다.
SUB_INIT_VAL:

    Findex = FreeFile
    Open LPT_PORT For Output As #Findex
    
    m_P.MagRect.Left = 15
    m_P.MagRect.Top = 15
    m_P.MagRect.Right = 180
    m_P.MagRect.Bottom = 280
    
    PRT_LINE_NUMBER = 0 ' 현재 라인을 출력하는 라인수
    
    iLeftLooper = 0     ' 한 라인에 여러개 출력될경우 현재 출력 갯수
    iLeftMaxLooper = 8  ' 한 라인에 출력할 전체 갯수
    iLeftNextSpace = ((m_P.MagRect.Right - m_P.MagRect.Left) / 7.5) ' 한 라인에 여러개 출력될경우 다음 출력 위치
    
    iTotalCnt = 0       ' 출력할 전체 건수
    iProcCnt = 0        ' 현재출력 중인 건수
    iPageTotProcCnt = 30 ' 한페이지당 전체 출력수 ( Details 수 )
    iPageProcCnt = 0    ' 한페이지에서 Details의 현재 출력중인 수
    
    iProcPageCnt = 0    ' 현재 출력중인 페이지 수
    iNextLineSpace = 6  ' Details의 다음 라인에 출력될 여백
    
    ' 선택된 대리점명을 구한다. ( 대리점명 ^ 택수량| )
    ReDim SelectCount(UBound(Split(PrtParam.Param(5), "|")))
    SelectCount = Split(PrtParam.Param(5), "|")
    
    ReDim SelectCode(UBound(Split(PrtParam.Param(6), "|")))
    SelectCode = Split(PrtParam.Param(6), "|")
    
    If UCase(TypeName(PView)) = "PICTUREBOX" Then
        Call PrnRefresh(PView, iSL_Type)
        PView.Cls
        
    ElseIf UCase(TypeName(PView)) = "PRINTER" Then
    ' 아무런 동작을 하지 않는다.
    End If

    Return

    
'---------------------------------------------------------------------------------------------
' 페이지 상단을 출력한다.
SUP_PAGE_HEADER:
'                                                                   ┌─┬────┬───┐
'      가락시영점(038) 물품 명세서                                  │확│ 본  사 │      │
'   ---------------------------------                               │  ├────┼───┤
'                                                                   │인│ 대리점 │      │
'                                                                   └─┴────┴───┘
'   2004년 11월 30일                                                              PAGE:     1
'   ------------------------------------------------------------------------------------------
'   택번호      택번호      택번호      택번호      택번호      택번호      택번호      택번호
'   ------------------------------------------------------------------------------------------
    
    Print #Findex, Space(PRT_L + 1)
    Print #Findex, Space(PRT_L + 1)
    Print #Findex, Space(PRT_L + 1)
    Print #Findex, Space(PRT_L + 10) & Trim(RS01!AgencyName) & " (" & RS01!AgencyCode & ") 물품명세서"
    Print #Findex, Space(PRT_L + 9) & String(Len(Trim(RS01!AgencyName) & " (" & RS01!AgencyCode & ") 물품명세서") + 10, "-")
    Print #Findex, Space(PRT_L + 62) & "┌─┬────┬─────┐"
    Print #Findex, Space(PRT_L + 62) & "│확│ 본  사 │          │"
    Print #Findex, Space(PRT_L + 62) & "│  ├────┼─────┤"
    Print #Findex, Space(PRT_L + 62) & "│인│ 대리점 │          │"
    Print #Findex, Space(PRT_L + 62) & "└─┴────┴─────┘"
    Print #Findex, Space(PRT_L + 1)
    Print #Findex, Space(PRT_L + 1) & Format(PrtParam.Param(2), "@@@@년 @@월 @@일") & _
                                            Space(63) & "PAGE:" & Format(iProcPageCnt, "@@@@@")
    Print #Findex, LINE_DISPLAY
    Print #Findex, Space(PRT_L + 0); "택번호      택번호      택번호      택번호      택번호      택번호      택번호      택번호"
    Print #Findex, LINE_DISPLAY
    
    PRT_LINE_NUMBER = 0 ' 현재 라인을 출력하는 라인수
    PRT_LINE_NUMBER = PRT_LINE_NUMBER + 15
    Return
    
'---------------------------------------------------------------------------------------------
' 레포트 하단을 출력한다.
SUB_REPORT_FOOTER:
'   ------------------------------------------------------------------------------------------
'   소품:    7   재세탁:    2   반품    0   수선:   0   택분실: 0   합계    37
'   ------------------------------------------------------------------------------------------
    Print #Findex, Space(PRT_L + 0) & " 소품:" & Format(StoreSoCnt, "@@@@@") & Space(5) & "재세탁:" & Format(StoreJaeCnt, "@@@@@") & Space(5) & _
                    "반품:" & Format(StoreBanCnt, "@@@@@") & Space(5) & "수선:" & Format(StoreSuCnt, "@@@@@") & Space(5) & _
                    "택분실:" & Format(StoreLostCnt, "@@@@@") & Space(5) & "합계:" & Format(StoreCount, "@@@@@"); ""
    
    PRT_LINE_NUMBER = PRT_LINE_NUMBER + 1
    StoreCount = 0:     StoreSoCnt = 0
    StoreJaeCnt = 0:    StoreBanCnt = 0
    StoreSuCnt = 0:     StoreLostCnt = 0
    Return
    
'---------------------------------------------------------------------------------------------
' 레포트 하단을 출력한다.
SUB_PAGE_FOOTER:
'   ------------------------------------------------------------------------------------------
'                                                                          크린에이드
    Do
        Print #Findex, Space(PRT_L + 0)
        PRT_LINE_NUMBER = PRT_LINE_NUMBER + 1
        ' 7줄을 더 찍어야 한기 때문에
        If PRT_LINE_NUMBER >= (PRT_LINE_MAX_NUMBER - 7) Then Exit Do
    Loop
    
    Print #Findex, LINE_DISPLAY
    Print #Findex, Space(PRT_L + 80) & "크린에이드"
    Print #Findex, Space(PRT_L + 0)
    Print #Findex, Space(PRT_L + 0)
    Print #Findex, Space(PRT_L + 0)
    Print #Findex, Space(PRT_L + 0)
    Print #Findex, Space(PRT_L + 0)
    Return

    
SUB_PRINTER:
    ' 내용을 출력한다.
    Dim bForExit    As Boolean
    ReDim sValue(1) As String
    
    For iSelectCode = 0 To UBound(SelectCode())
    
        sValue(0) = PrtParam.Param(2)
        sValue(1) = "(A.AgencyCode = '" & SelectCode(iSelectCode) & "') "
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("ChulgoTag_New", sValue(), Err_Num, Err_Dec)
        
        If Err_Num <> 0 Then
            Close #Findex
            RaiseEvent ErrorPrinter("1", Err_Dec)
            Exit Function
        End If
        
        PView.Font.Bold = False:    PView.Font.Size = 10.5
        iLeftNextSpace = ((m_P.MagRect.Right - m_P.MagRect.Left) / 7.5)
        
        Do While Not RS01.EOF
            
            If RS01!SendChk = "0" Then
                        
                iLeftLooper = iLeftLooper + 1
                
                ' Page Header 출력
                If (iPageProcCnt Mod iPageTotProcCnt) = 0 And iLeftLooper = 1 Then
                    ' 현재 출력중인 페이지 수
                    iProcPageCnt = iProcPageCnt + 1
                    ' 한페이지당 현재 출려 라인 수
                    iPageProcCnt = 0
                    StoreCode = RS01!AgencyCode
                    
                    GoSub SUP_PAGE_HEADER
                    
                End If
                    
                ' 출력중인 전체 수
                iProcCnt = iProcCnt + 1
                StoreCount = StoreCount + 1:    StoreSoCnt = StoreSoCnt + Val(RS01!SoCnt)
                StoreJaeCnt = StoreJaeCnt + Val(RS01!JaeCnt):  StoreBanCnt = StoreBanCnt + Val(RS01!BanCnt)
                StoreSuCnt = StoreSuCnt + Val(RS01!SuCnt):    StoreLostCnt = StoreLostCnt + Val(RS01!LostCnt)
                
                ' Details 출력
                PRT_TAG_DATA = PRT_TAG_DATA & RS01!Tag & IIf(InStr(RS01!Tag, "반") > 0, Space(4), Space(5))
                                
                ' "0"일의 값으로 마지막일 경우 모든 처리를 하고 빠저 나간다.
                If RS01.EOF Then bForExit = True
                If bForExit Then
                    PRT_TAG_DATA = RTrim(PRT_TAG_DATA)
                    Print #Findex, Space(PRT_L + 0) & PRT_TAG_DATA:   PRT_TAG_DATA = "":      PRT_LINE_NUMBER = PRT_LINE_NUMBER + 1
                    
                    GoSub SUB_REPORT_FOOTER
                    GoSub SUB_PAGE_FOOTER
                    Exit Do
                End If
                
                
'    '           If Trim(RS01!Tag) = "9-867" Then MsgBox RS01!Tag
                ' 1 Page 이상일경우 마지막 부분에 라인을 표시한다.
                If ((iPageProcCnt + 1) Mod iPageTotProcCnt) = 0 And iLeftLooper = iLeftMaxLooper Then
                    
                    PRT_TAG_DATA = RTrim(PRT_TAG_DATA)
                    Print #Findex, Space(PRT_L + 0) & PRT_TAG_DATA:   PRT_TAG_DATA = "":      PRT_LINE_NUMBER = PRT_LINE_NUMBER + 1
                    
                    GoSub SUB_PAGE_FOOTER
                    ' 다음 시작 페이지를 찍는다.
                    GoSub SUP_PAGE_HEADER
                End If

                ' 라인을 증가 한다.
                If iLeftLooper = iLeftMaxLooper Then
                    PRT_TAG_DATA = RTrim(PRT_TAG_DATA)
                    Print #Findex, Space(PRT_L + 0) & PRT_TAG_DATA: PRT_TAG_DATA = "":      PRT_LINE_NUMBER = PRT_LINE_NUMBER + 1
                    
                    iLeftLooper = 0
                    iPageProcCnt = iPageProcCnt + 1
                End If
            End If
            
            RS01.MoveNext

            If bForExit Then Exit Do
        Loop
        
        ' 마지막 라인이 안찍히고 종료되기 때문에 마지막 라인을 찍어준다.
        If PRT_TAG_DATA <> "" Then
            Print #Findex, Space(PRT_L + 0) & PRT_TAG_DATA:   PRT_TAG_DATA = "":      PRT_LINE_NUMBER = PRT_LINE_NUMBER + 1
        End If
        
        
        ' 하나의 대리점 출력이 종료 되었음.
        ' bForExit 값이 True일경우는 "0"으로 출력 되었기 때문에 출력하면 안된다.
        If Not bForExit Then
            Print #Findex, LINE_DISPLAY
            PRT_LINE_NUMBER = PRT_LINE_NUMBER + 1
            GoSub SUB_REPORT_FOOTER
            GoSub SUB_PAGE_FOOTER
        End If
            
        iLeftLooper = 0     ' 한 라인에 여러개 출력될경우 현재 출력 갯수
        
        iTotalCnt = 0       ' 출력할 전체 건수
        iProcCnt = 0        ' 현재출력 중인 건수
        iPageProcCnt = 0    ' 한페이지에서 Details의 현재 출력중인 수
        
        iProcPageCnt = 0    ' 현재 출력중인 페이지 수
        StoreCount = 0      ' 한 거래처당 출려수
    
    Next iSelectCode
    Return
    
' 마무리 한다.
SUB_END_PRINT:
    RS01.Close
    Return

    On Error GoTo 0
    Exit Function

PRT_P03001_Error:

    m_Error.ErrorMsg = "Error " & Err.Number & " (" & Err.Description & ") in procedure PRT_03001_DOT_VIEW of Form CCAIDPrinter"

    If m_Error.VisibleMSG Then
        MsgBox m_Error.ErrorMsg, vbInformation, "확인 [Error]"
    End If

    If m_Error.SaveLog Then
        Call ProgramErrorLogWrite(m_Error)
    End If

    If m_Error.ResumeMode Then
        Resume
    End If


End Function

'정상 인쇄
Private Sub DOT_PrintString(HPRT As Integer, pStr As String)
    Print #HPRT, Chr$(&H1B); "!"; Chr$(&H1); pStr;
End Sub

'2H 확대
Private Sub DOT_2HPrintString(HPRT As Integer, pStr As String)
    Print #HPRT, Chr$(&H1B); "!"; Chr$(&H10); pStr;
End Sub

'2W 확대
Private Sub DOT_2WPrintString(HPRT As Integer, pStr As String)
    Print #HPRT, Chr$(&H1B); "!"; Chr$(&H20); pStr;
End Sub

'2HW 확대
Private Sub DOT_2HWPrintString(HPRT As Integer, pStr As String)
    Print #HPRT, Chr$(&H1B); "!"; Chr$(&H30); pStr;
End Sub

' 신규 페이지
Private Sub DOT_PrintCrlf(HPRT As Integer)
    Print #HPRT, Chr$(&HA);
End Sub

Private Sub DOT_PrinterFeed(HPRT As Integer, pfeed As Integer)
    Dim i As Integer
    For i = 0 To pfeed
        Print #HPRT, Chr$(&HA)
    Next i
End Sub

'연결점 1개남긴후 cut
Private Sub DOT_PartialCut(HPRT As Integer)
    Print #HPRT, Chr$(&H1B); "m";
End Sub

'구동펄스 발생
Private Sub DOT_PulseRise(HPRT As Integer)
    On Error Resume Next
'    If Left(SECTION_JIJUM.jijum_kumgo_open, 1) = "Y" Then
'        Call gFunc_LptOpen
        Print #HPRT, Chr$(&H1B); "p"; Chr$(0); Chr$(2); Chr$(10);
'        Call gProc_LptClose
'    End If
End Sub



Public Function PRT_06006_VIEW(PView As Object, sDate As String, sStoreCode As String, sNumber As String) As Boolean
' 사고접수 보고서를 출력한다
    Dim RS01 As ADODB.Recordset
    ReDim sValue(3)
    
    PRT_06006_VIEW = True

    ' 기본 프린터가 없을 경우
    If Not Fn_PrinterCheck Then
        PRT_06006_VIEW = False
        Exit Function
    End If

    ''''''''''''''''
    On Error GoTo printError
    '''''''''''''''

    
    sValue(0) = "0"
    sValue(1) = sDate
    sValue(2) = sStoreCode
    sValue(3) = sNumber
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_06006_02", sValue(), Err_Num, Err_Dec)
    
    If RS01.EOF Then
        Exit Function
    End If

Print_Start:
    Prt_Top = 5
    Prt_Left = 15

    m_P.MagRect.Left = 15:          m_P.MagRect.Top = 5
    m_P.MagRect.Right = 200:        m_P.MagRect.Bottom = 285
    
    m_P.MaxCount = 0             ' 출력할 전체 건수
    m_P.ProcCnt = 0             ' 현재출력 중인 건수
    m_P.PageProcCnt = 0         ' 한페이지에서 Details의 현재 출력중인 수
    
    m_P.NextLineSpace = 5       ' Details의 다음 라인에 출력될 여백
    'm_P.ProcPageCnt = ViewPage
    
    If iSL_Type = 0 Then iSL_Type = 1
    
    
    Printer.ScaleMode = vbMillimeters           ' 프린터의 스케일 모드를 밀리미터 단위로

    Printer.FontName = "굴림체"
    Printer.Font.Bold = True
    Printer.Font.Size = 9

    Printer.Font.Size = "18"
    PrintText Printer, m_P, 55, 0, "사 고 접 수 보 고 서"
    Printer.DrawWidth = 12
    PrintLine Printer, m_P, 50, 8, 125, 8


    ' 결재 양식 작성
    GoSub PrintApproval

    ' 기본사항 양식 작성
    GoSub PrintDefault

    ' 피해관련사항
    GoSub PrintDamage

    ' 대리점 기재
    GoSub AgencyWrite

    ' 본사 기재
    'GoSub CompanyWrite

    '출력 한다.
    RS01.Close
    Printer.EndDoc
    Exit Function

printError:
    PRT_06006_VIEW = False
    Printer.EndDoc
    RaiseEvent ErrorPrinter(Err.Number, Err.Description)
    Exit Function

' 결재 양식 인쇄
PrintApproval:

    Top_Margin = 0: Left_Margin = 0

    Printer.Font.Size = 9
    Printer.DrawWidth = 7
    PrintRect Printer, m_P, 89, 18, 175, 37    '외각 틀
    PrintLine Printer, m_P, 94, 24, 175, 24   '수평 라인

    PrintLine Printer, m_P, 94, 18, 94, 37  ' 결재
    PrintText Printer, m_P, 90, 20, "결"
    PrintText Printer, m_P, 90, 32, "재"
    PrintText Printer, m_P, 97, 19, "담  당"
    
    PrintLine Printer, m_P, 110, 18, 110, 37
    PrintText Printer, m_P, 113, 19, "대  리"

    PrintLine Printer, m_P, 126, 18, 126, 37
    PrintText Printer, m_P, 129, 19, "부  장"

    PrintLine Printer, m_P, 142, 18, 142, 37
    PrintText Printer, m_P, 145, 19, "이  사"

    PrintLine Printer, m_P, 158, 18, 158, 37
    PrintText Printer, m_P, 161, 19, "사  장"
    
    Return

' 기본사항 양식 작성
PrintDefault:

    Top_Margin = 0: Left_Margin = 0

    Printer.Font.Size = 9
    Printer.DrawWidth = 7
    Text_Height = Printer.TextHeight("aa")

    PrintRect Printer, m_P, 0, 63, 180, 91    '외각 틀
    PrintLine Printer, m_P, 60, 70, 180, 70   '수평 라인
    PrintLine Printer, m_P, 0, 77, 180, 77    '수평 라인
    PrintLine Printer, m_P, 60, 84, 180, 84   '수평 라인

    PrintLine Printer, m_P, 30, 63, 30, 91    '수직 라인
    PrintLine Printer, m_P, 60, 63, 60, 91    '수직 라인
    PrintLine Printer, m_P, 90, 63, 90, 91    '수직 라인
    'PrintLine Printer, m_P, 120, 70, 120, 77    '수직 라인
    'PrintLine Printer, m_P, 150, 70, 150, 77    '수직 라인
    PrintLine Printer, m_P, 120, 84, 120, 91    '수직 라인
    PrintLine Printer, m_P, 150, 84, 150, 91    '수직 라인

    PrintText Printer, m_P, 0, 58, "⊙ 기본사항"
    PrintText Printer, m_P, 6, 68, "대 리 점 명"
    PrintText Printer, m_P, 69, 65, "지    사"
    PrintText Printer, m_P, 69, 72, "전    화"
    PrintText Printer, m_P, 129, 72, " "
    PrintText Printer, m_P, 6, 82, "소 비 자 명"
    PrintText Printer, m_P, 69, 79, "주    소"
    PrintText Printer, m_P, 69, 86, "전    화"
    PrintText Printer, m_P, 129, 86, "핸 드 폰"
    ' 값 출력
    PrintText Printer, m_P, 35, 65, "(" & RS01!대리점코드 & ")"
    PrintText Printer, m_P, 35, 70, RS01!대리점명
    
    PrintText Printer, m_P, 95, 65, "[" & RS01!사업장코드 & "]  " & RS01!지사명
    PrintText Printer, m_P, 95, 72, RS01!대리점전화번호
    PrintText Printer, m_P, 155, 72, ""
    
    PrintText Printer, m_P, 35, 82, RS01!고객명
    PrintText Printer, m_P, 95, 79, RS01!고객주소
    PrintText Printer, m_P, 95, 86, RS01!전화번호
    PrintText Printer, m_P, 155, 86, RS01!휴대폰번호
    Return
    

' 피해관련사항
PrintDamage:

    Top_Margin = 0: Left_Margin = 0
    Printer.Font.Size = 9
    Printer.DrawWidth = 7
    Text_Height = Printer.TextHeight("a")

    PrintRect Printer, m_P, 0, 105, 180, 140      '외각 틀
    PrintLine Printer, m_P, 0, 112, 180, 112      '수평 라인
    PrintLine Printer, m_P, 0, 119, 180, 119      '수평 라인
    PrintLine Printer, m_P, 0, 126, 180, 126      '수평 라인
    PrintLine Printer, m_P, 0, 133, 180, 133      '수평 라인

    PrintLine Printer, m_P, 35, 105, 35, 140      '수직 라인
    PrintLine Printer, m_P, 85, 105, 85, 140      '수직 라인
    PrintLine Printer, m_P, 120, 105, 120, 140      '수직 라인
    PrintLine Printer, m_P, 135, 119, 135, 133      '수직 라인
    PrintLine Printer, m_P, 150, 119, 150, 133      '수직 라인
    PrintLine Printer, m_P, 165, 119, 165, 133      '수직 라인

    PrintText Printer, m_P, 0, 100, "⊙ 피해관련사항"
    PrintText Printer, m_P, 10, 107, "품    목"
    PrintText Printer, m_P, 95, 107, "상    표"
    PrintText Printer, m_P, 10, 114, "구 입 일"
    PrintText Printer, m_P, 95, 114, "색    상"
    PrintText Printer, m_P, 10, 121, "구 입 처"
    PrintText Printer, m_P, 95, 121, "택 번 호"
    PrintText Printer, m_P, 10, 128, "구입형태"
    PrintText Printer, m_P, 95, 128, "입 고 일"
    PrintText Printer, m_P, 10, 135, "구입가격"
    PrintText Printer, m_P, 94, 135, "사고접수일"
    PrintText Printer, m_P, 122, 121, "최초택"
    PrintText Printer, m_P, 152, 121, "최종택"
    PrintText Printer, m_P, 122, 128, "최초일"
    PrintText Printer, m_P, 152, 128, "최종일"
    
    ' 값 출력
    PrintText Printer, m_P, 40, 107, RS01!품명
    PrintText Printer, m_P, 125, 107, RS01!브랜드
    PrintText Printer, m_P, 40, 114, RS01!구입일자
    PrintText Printer, m_P, 125, 114, RS01!색상
    PrintText Printer, m_P, 40, 121, RS01!구입처
    PrintText Printer, m_P, 40, 128, RS01!구입형태
    PrintText Printer, m_P, 40, 135, Format(RS01!구입가격, "#,##0")
    
    PrintText Printer, m_P, 125, 135, RS01!사고접수일
    
    PrintText Printer, m_P, 137, 121, RS01!최초택번호
    PrintText Printer, m_P, 167, 121, RS01!최종택번호
    PrintText Printer, m_P, 136, 128, RS01!최초입고일
    PrintText Printer, m_P, 166, 128, RS01!최종입고일
    Return
    

' 대리점 기재
AgencyWrite:

    Top_Margin = 0: Left_Margin = 0
    Printer.Font.Size = 9
    Printer.DrawWidth = 7
    Text_Height = Printer.TextHeight("a")

    PrintRect Printer, m_P, 0, 170, 180, 212      '외각 틀
    PrintLine Printer, m_P, 0, 177, 180, 177      '수평 라인
    PrintLine Printer, m_P, 0, 184, 180, 184      '수평 라인
    PrintLine Printer, m_P, 0, 191, 180, 191      '수평 라인
    PrintLine Printer, m_P, 0, 198, 180, 198      '수평 라인
    PrintLine Printer, m_P, 0, 205, 180, 205      '수평 라인

    PrintLine Printer, m_P, 43, 170, 43, 212      '수직 라인

    PrintText Printer, m_P, 0, 165, "⊙ 대리점 기재"
    PrintText Printer, m_P, 13, 172, "사고의 종류"
    PrintText Printer, m_P, 13, 179, "사고의 내용"
    PrintText Printer, m_P, 4, 186, "소비자의견 및 요구사항"
    PrintText Printer, m_P, 10, 193, "보상 산정 금액"
    PrintText Printer, m_P, 13, 200, "합 의 금 액"
    PrintText Printer, m_P, 13, 207, "처 리 여 부"
    
    ' 값 출력
    PrintText Printer, m_P, 50, 172, RS01!사고종류
    PrintText Printer, m_P, 50, 179, RS01!사고내용
    PrintText Printer, m_P, 50, 186, RS01!사고의견
    
    PrintText Printer, m_P, 50, 193, RS01!보상금액
    PrintText Printer, m_P, 50, 200, RS01!합의금액
    PrintText Printer, m_P, 50, 207, RS01!처리유무
    
    Return

' 본사 기재
CompanyWrite:

    Top_Margin = 0: Left_Margin = 0
    Printer.Font.Size = 9
    Printer.DrawWidth = 7
    Text_Height = Printer.TextHeight("a")

    PrintRect Printer, m_P, 0, 205, 180, 226      '외각 틀
    PrintLine Printer, m_P, 0, 212, 180, 212      '수평 라인
    PrintLine Printer, m_P, 0, 219, 180, 219      '수평 라인

    PrintLine Printer, m_P, 30, 205, 30, 226      '수직 라인
    PrintLine Printer, m_P, 60, 205, 60, 226      '수직 라인
    PrintLine Printer, m_P, 90, 205, 90, 226      '수직 라인
    PrintLine Printer, m_P, 120, 205, 120, 226      '수직 라인
    PrintLine Printer, m_P, 150, 205, 150, 226      '수직 라인

    PrintText Printer, m_P, 0, 200, "⊙ 본사 기재"
    PrintText Printer, m_P, 6, 207, "제 조 회 사"
    PrintText Printer, m_P, 66, 207, "전 화 번 호"
    PrintText Printer, m_P, 126, 207, "판 매 일 자"
    PrintText Printer, m_P, 6, 214, "재 고 현 황"
    PrintText Printer, m_P, 66, 214, "담       당"
    PrintText Printer, m_P, 126, 214, "판 매 금 액"
    PrintText Printer, m_P, 6, 221, "보 상 비 율"
    PrintText Printer, m_P, 65, 221, "보상산정금액"
    PrintText Printer, m_P, 126, 221, "합 의 내 용"
    ' 값 출력
    PrintText Printer, m_P, 33, 207, "" ' 제조회사
    PrintText Printer, m_P, 93, 207, "" ' 전화번호
    PrintText Printer, m_P, 153, 207, "" '판매일자
    PrintText Printer, m_P, 33, 214, "" '재고현황
    PrintText Printer, m_P, 93, 214, "" '담당
    PrintText Printer, m_P, 153, 214, "" '판매금액
    PrintText Printer, m_P, 33, 221, "" '보상비율
    PrintText Printer, m_P, 93, 221, "" '보상산정금액
    PrintText Printer, m_P, 153, 221, "" '합의내용
    Return


End Function



Public Function PRT_06007_VIEW(PView As Object, sDate As String, sNumber As String) As Boolean
' 사고접수 보고서를 출력한다
    Dim RS01 As ADODB.Recordset
    ReDim sValue(2)
    
    PRT_06007_VIEW = True

    ' 기본 프린터가 없을 경우
    If Not Fn_PrinterCheck Then
        PRT_06007_VIEW = False
        Exit Function
    End If

    ''''''''''''''''
    On Error GoTo printError
    '''''''''''''''

    
    sValue(0) = "0"
    sValue(1) = sDate
    sValue(2) = sNumber
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_06007_02", sValue(), Err_Num, Err_Dec)
    
    If RS01.EOF Then
        Exit Function
    End If

Print_Start:
    Prt_Top = 20
    Prt_Left = 15

    m_P.MagRect.Left = 15:          m_P.MagRect.Top = 20
    m_P.MagRect.Right = 200:        m_P.MagRect.Bottom = 285
    
    m_P.MaxCount = 0             ' 출력할 전체 건수
    m_P.ProcCnt = 0             ' 현재출력 중인 건수
    m_P.PageProcCnt = 0         ' 한페이지에서 Details의 현재 출력중인 수
    
    m_P.NextLineSpace = 5       ' Details의 다음 라인에 출력될 여백
    'm_P.ProcPageCnt = ViewPage
    
    If iSL_Type = 0 Then iSL_Type = 1
    
    
    Printer.ScaleMode = vbMillimeters           ' 프린터의 스케일 모드를 밀리미터 단위로

    Printer.FontName = "굴림체"
    Printer.Font.Bold = True
    Printer.Font.Size = 9

    Printer.Font.Size = "18"
    PrintText Printer, m_P, 55, 0, "사 고 접 수 보 고 서"
    Printer.DrawWidth = 12
    PrintLine Printer, m_P, 50, 8, 125, 8


    ' 결재 양식 작성
     ReDim sValue(8)
    sValue(0) = "담 당"
    sValue(1) = "주 임"
    sValue(2) = "대 리"
    sValue(3) = "과 장"
    sValue(4) = "차 장"
    sValue(5) = "부 장"
    sValue(6) = "이 사"
    sValue(7) = "상 무"
    sValue(8) = "대 표"
    Call Def_Approval2(Printer, sValue)
    
    ' 접수 내역
    GoSub PrintDefault

    ' 품목 정보
    GoSub PrintGoodsInfo

    ' 구입자 정보
    GoSub CustInfo

    ' 본사 기재
    GoSub CompanyWrite
    
    ' 대리점 의견
    GoSub StoreInfo

    '출력 한다.
    RS01.Close
    Printer.EndDoc
    Exit Function

printError:
    PRT_06007_VIEW = False
    Printer.EndDoc
    RaiseEvent ErrorPrinter(Err.Number, Err.Description)
    Resume
    Exit Function


' 기본사항 양식 작성
PrintDefault:

    Top_Margin = 0: Left_Margin = 0

    Printer.Font.Size = 9
    Printer.DrawWidth = 7
    Text_Height = Printer.TextHeight("aa")

    PrintRect Printer, m_P, 0, 53, 180, 74    '외각 틀
    PrintLine Printer, m_P, 0, 60, 180, 60    '1 수평 라인
    PrintLine Printer, m_P, 0, 67, 180, 67    '2 수평 라인
    
    PrintLine Printer, m_P, 35, 53, 35, 74     '수직 라인
    PrintLine Printer, m_P, 85, 53, 85, 74     '수직 라인
    PrintLine Printer, m_P, 120, 53, 120, 74   '수직 라인
    
    PrintText Printer, m_P, 0, 48, "⊙ 접수내역"
    PrintText Printer, m_P, 10, 55, "가 맹 점 명"
    PrintText Printer, m_P, 95, 55, "전 화 번 호"
    PrintText Printer, m_P, 10, 62, "사 업 장 명"
    PrintText Printer, m_P, 95, 62, "전 화 번 호"
    PrintText Printer, m_P, 10, 69, "담 당 자 명"
    
    ' 값 출력
    PrintText Printer, m_P, 40, 55, "(" & RS01!대리점코드 & ")  " & RS01!대리점명
    PrintText Printer, m_P, 125, 55, RS01!대리점전화번호
    
    
    PrintText Printer, m_P, 40, 62, "[" & RS01!사업장코드 & "]  " & RS01!사업장명
    PrintText Printer, m_P, 125, 62, ""
    
    PrintText Printer, m_P, 40, 69, RS01!담당자명
    Return
    

' 품목정보
PrintGoodsInfo:

    Top_Margin = 0: Left_Margin = 0
    Printer.Font.Size = 9
    Printer.DrawWidth = 7
    Text_Height = Printer.TextHeight("a")

    PrintRect Printer, m_P, 0, 85, 180, 127      '외각 틀
    
    PrintLine Printer, m_P, 0, 92, 180, 92      '수평 라인
    PrintLine Printer, m_P, 0, 99, 180, 99      '수평 라인
    PrintLine Printer, m_P, 0, 106, 180, 106    '수평 라인
    PrintLine Printer, m_P, 0, 113, 180, 113    '수평 라인
    PrintLine Printer, m_P, 0, 120, 180, 120    '수평 라인
    
    PrintLine Printer, m_P, 35, 85, 35, 127     '수직 라인
    PrintLine Printer, m_P, 85, 85, 85, 127     '수직 라인
    PrintLine Printer, m_P, 120, 85, 120, 127   '수직 라인
    
    
    PrintText Printer, m_P, 0, 80, "⊙ 품목정보"
    PrintText Printer, m_P, 10, 87, "입 고 일 자"
    PrintText Printer, m_P, 95, 87, "택  번  호"
    PrintText Printer, m_P, 10, 94, "출 고 일 자"
    PrintText Printer, m_P, 95, 94, "인 도 일 자"
    PrintText Printer, m_P, 10, 101, "품       목"
    PrintText Printer, m_P, 95, 101, "브  랜  드"
    PrintText Printer, m_P, 10, 108, "색       상"
    PrintText Printer, m_P, 95, 108, ""
    PrintText Printer, m_P, 10, 115, "구 입 일 자"
    PrintText Printer, m_P, 94, 115, "구  입  처"
    PrintText Printer, m_P, 10, 122, "구 입 형 태"
    PrintText Printer, m_P, 94, 122, "구 입 가 격"
    
    ' 값 출력
    PrintText Printer, m_P, 40, 87, Format(RS01!입고일자, "@@@@-@@-@@")
    PrintText Printer, m_P, 125, 87, RS01!택번호
    PrintText Printer, m_P, 40, 94, Format(RS01!출고일자, "@@@@-@@-@@")
    PrintText Printer, m_P, 125, 94, Format(RS01!인도일자, "@@@@-@@-@@")
    PrintText Printer, m_P, 40, 101, RS01!품목
    PrintText Printer, m_P, 125, 101, RS01!브랜드
    PrintText Printer, m_P, 40, 108, RS01!색상
    PrintText Printer, m_P, 125, 108, ""
    PrintText Printer, m_P, 40, 115, Format(RS01!구입일자, "@@@@-@@-@@")
    PrintText Printer, m_P, 125, 115, RS01!구입처
    PrintText Printer, m_P, 40, 122, RS01!구입형태
    PrintText Printer, m_P, 125, 122, Format(RS01!구입가격, "#,##0")
    Return
    

' 구입자 정보
CustInfo:

    Top_Margin = 0: Left_Margin = 0
    Printer.Font.Size = 9
    Printer.DrawWidth = 7
    Text_Height = Printer.TextHeight("a")

    PrintRect Printer, m_P, 0, 138, 180, 180      '외각 틀
    PrintLine Printer, m_P, 0, 145, 180, 145      '수평 라인
    PrintLine Printer, m_P, 0, 152, 180, 152      '수평 라인
    PrintLine Printer, m_P, 0, 159, 180, 159      '수평 라인
    PrintLine Printer, m_P, 0, 166, 180, 166      '수평 라인
    PrintLine Printer, m_P, 0, 173, 180, 173      '수평 라인
    
    PrintLine Printer, m_P, 35, 138, 35, 180      '수직 라인
    PrintLine Printer, m_P, 85, 138, 85, 145      '수직 라인
    PrintLine Printer, m_P, 120, 138, 120, 145    '수직 라인
    PrintLine Printer, m_P, 85, 152, 85, 166      '수직 라인
    PrintLine Printer, m_P, 120, 152, 120, 166    '수직 라인
    
    PrintText Printer, m_P, 0, 133, "⊙ 고객 정보"
    PrintText Printer, m_P, 10, 140, "성    명"
    PrintText Printer, m_P, 95, 140, "전 화 번 호"
    PrintText Printer, m_P, 10, 147, "고 객 의 견"
    PrintText Printer, m_P, 10, 154, "처 리 일 자"
    PrintText Printer, m_P, 95, 154, "휴대폰번호"
    PrintText Printer, m_P, 10, 161, "크레임구분"
    PrintText Printer, m_P, 95, 161, "보 상 구 분"
    PrintText Printer, m_P, 10, 168, "보 상 금 액"
    PrintText Printer, m_P, 10, 175, "제 품 정 보"
    
    ' 값 출력
    PrintText Printer, m_P, 40, 140, RS01!고객성명
    PrintText Printer, m_P, 125, 140, RS01!고객전화번호
    PrintText Printer, m_P, 40, 147, RS01!고객주소
    PrintText Printer, m_P, 40, 154, Format(RS01!처리일자, "@@@@-@@-@@")
    PrintText Printer, m_P, 125, 154, RS01!휴대폰번호
    PrintText Printer, m_P, 40, 161, RS01!크레임구분
    PrintText Printer, m_P, 125, 161, RS01!보상구분
    PrintText Printer, m_P, 40, 168, Format(RS01!보상금액, "#,##0")
    PrintText Printer, m_P, 40, 175, RS01!비고
    
    Return

' 보상산정기준
CompanyWrite:

    Top_Margin = 0: Left_Margin = 0
    Printer.Font.Size = 9
    Printer.DrawWidth = 7
    Text_Height = Printer.TextHeight("a")

    PrintRect Printer, m_P, 0, 190, 180, 218      '외각 틀
    PrintLine Printer, m_P, 0, 197, 180, 197      '수평 라인
    PrintLine Printer, m_P, 0, 204, 180, 204      '수평 라인
    PrintLine Printer, m_P, 0, 211, 180, 211      '수평 라인
    
    PrintLine Printer, m_P, 35, 190, 35, 218     '수직 라인
    PrintLine Printer, m_P, 85, 190, 85, 218     '수직 라인
    PrintLine Printer, m_P, 120, 190, 120, 218   '수직 라인
    
    PrintText Printer, m_P, 0, 185, "⊙ 보상산정기준"
    PrintText Printer, m_P, 10, 192, "품       목"
    PrintText Printer, m_P, 95, 192, "용       도"
    PrintText Printer, m_P, 10, 199, "소       재"
    PrintText Printer, m_P, 95, 199, "내 용 연 수"
    PrintText Printer, m_P, 10, 206, "경 과 일 수"
    PrintText Printer, m_P, 95, 206, "환 산 일 수"
    PrintText Printer, m_P, 10, 213, "배 상 비 율"
    PrintText Printer, m_P, 95, 213, "보상산정금액"
    
    ' 값 출력
    PrintText Printer, m_P, 40, 192, "[" & RS01!클레임품목코드 & "] " & RS01!클레임품목명칭
    PrintText Printer, m_P, 125, 192, "[" & RS01!클레임용도코드 & "] " & RS01!클레임용도명칭
    PrintText Printer, m_P, 40, 199, "[" & RS01!클레임소재코드 & "] " & RS01!클레임소재명칭
    PrintText Printer, m_P, 125, 199, RS01!내용연수
    PrintText Printer, m_P, 40, 206, RS01!경과일수
    PrintText Printer, m_P, 125, 206, RS01!환산일수
    PrintText Printer, m_P, 40, 213, RS01!배상일수  '배상비율
    PrintText Printer, m_P, 125, 213, Format(RS01!배상금액, "#,##0")
    Return
 
' 담당자의견
StoreInfo:

    Top_Margin = 0: Left_Margin = 0
    Printer.Font.Size = 9
    Printer.DrawWidth = 7
    Text_Height = Printer.TextHeight("a")

    PrintRect Printer, m_P, 0, 230, 180, 251      '외각 틀
    PrintLine Printer, m_P, 0, 237, 180, 237      '수평 라인
    PrintLine Printer, m_P, 0, 244, 180, 244      '수평 라인
    
    PrintLine Printer, m_P, 35, 230, 35, 251      '수직 라인
    
    
    PrintText Printer, m_P, 0, 225, "⊙ 담당자 의견"
    PrintText Printer, m_P, 10, 232, "의 견 1"
    PrintText Printer, m_P, 10, 239, "의 견 2"
    PrintText Printer, m_P, 10, 246, "결제 계좌"
    
    ' 값 출력
    PrintText Printer, m_P, 40, 232, RS01!담당자의견1
    PrintText Printer, m_P, 40, 239, RS01!담당자의견2
    PrintText Printer, m_P, 40, 246, RS01!담당자의견3
    
    Return

End Function




'*************************************************************************************************
' 기본 결재 라인을 그린다.
Private Function Def_Approval2(PView As Object, sValue As Variant) As Boolean
'                                                        ┌─┬───┬───┬───┬───┐
'                                                        │결│ 경리 │ 담당 │ 차장 │ 사장 │
'                                                        │  ├───┼───┼───┼───┤
'                                                        │재│      │      │      │      │
'                                                        └─┴───┴───┴───┴───┘

    Dim x1 As Long, y1 As Long
    Dim x2 As Long, y2 As Long
    
    Dim nHeight As Long
    Dim nWidth  As Long
    
    Dim nCnt        As Integer
    Dim nTotalCnt   As Integer
    
    x2 = 175
    y2 = 37
    nWidth = 13
    nHeight = 20
    nTotalCnt = UBound(sValue) + 1
    
    x1 = x2 - ((nWidth * nTotalCnt) + 5)  ' 5 는 결재 명칭
    y1 = y2 - nHeight
    
    Printer.Font.Size = 9
    Printer.DrawWidth = 7
    
    PrintRect Printer, m_P, x1, y1, x2, y2              '외각 틀
    PrintLine Printer, m_P, x1 + 5, y1 + 5, x2, y1 + 5  '수평 라인
    
    PrintText Printer, m_P, x1 + 1, y1 + 5, "결"
    PrintText Printer, m_P, x1 + 1, y1 + 13, "재"
    
    For nCnt = 0 To nTotalCnt - 1
        PrintLine Printer, m_P, x1 + 5 + (nWidth * nCnt), y1, x1 + 5 + (nWidth * nCnt), y2
        PrintText Printer, m_P, x1 + 5 + 2 + (nWidth * nCnt), y1 + 1, CStr(sValue(nCnt))
    Next nCnt
End Function
