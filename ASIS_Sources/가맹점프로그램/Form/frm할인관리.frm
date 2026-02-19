VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm할인관리 
   Caption         =   "할인현황"
   ClientHeight    =   10230
   ClientLeft      =   2895
   ClientTop       =   3210
   ClientWidth     =   15540
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form11"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10230
   ScaleWidth      =   15540
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   10230
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15540
      _ExtentX        =   27411
      _ExtentY        =   18045
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm할인관리.frx":0000
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   15510
         _ExtentX        =   27358
         _ExtentY        =   741
         _Version        =   262144
         Font3D          =   3
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "      할인 현황"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm할인관리.frx":0092
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm할인관리.frx":02B8
            Top             =   -15
            Width           =   765
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   750
         Left            =   15
         TabIndex        =   2
         Top             =   450
         Width           =   15510
         _ExtentX        =   27358
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSOption optSale 
            Height          =   345
            Index           =   0
            Left            =   105
            TabIndex        =   7
            Top             =   210
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   609
            _Version        =   262144
            Font3D          =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "품목할인"
         End
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   630
            Left            =   8340
            TabIndex        =   3
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 조회(&F)"
            Appearance      =   6
            Picture         =   "frm할인관리.frx":0E82
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   3
            Left            =   9885
            TabIndex        =   4
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 엑셀(&E)"
            Appearance      =   6
            Picture         =   "frm할인관리.frx":157C
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   13170
            TabIndex        =   5
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frm할인관리.frx":1CF6
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   4
            Left            =   11430
            TabIndex        =   6
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 출력(&P)"
            Appearance      =   6
            Picture         =   "frm할인관리.frx":2D88
         End
         Begin Threed.SSOption optSale 
            Height          =   345
            Index           =   1
            Left            =   1815
            TabIndex        =   8
            Top             =   210
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   609
            _Version        =   262144
            Font3D          =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "요일할인"
         End
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   9000
         Left            =   15
         TabIndex        =   9
         Top             =   1215
         Width           =   4665
         _Version        =   524288
         _ExtentX        =   8229
         _ExtentY        =   15875
         _StockProps     =   64
         AutoCalc        =   0   'False
         BackColorStyle  =   1
         ColsFrozen      =   4
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
         DInformActiveRowChange=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         MaxCols         =   4
         MaxRows         =   200
         OperationMode   =   1
         Protect         =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frm할인관리.frx":3482
         UserResize      =   1
         VisibleCols     =   4
         VisibleRows     =   50
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread sprList 
         Height          =   9000
         Left            =   4695
         TabIndex        =   10
         Top             =   1215
         Width           =   10830
         _Version        =   524288
         _ExtentX        =   19103
         _ExtentY        =   15875
         _StockProps     =   64
         AutoCalc        =   0   'False
         BackColorStyle  =   1
         ColsFrozen      =   4
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
         DInformActiveRowChange=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         MaxCols         =   9
         MaxRows         =   200
         MoveActiveOnFocus=   0   'False
         OperationMode   =   1
         Protect         =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frm할인관리.frx":3B00
         UserResize      =   1
         VisibleCols     =   6
         VisibleRows     =   50
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "frm할인관리"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim 시작일자 As String
Dim 종료일자 As String
Dim 요일     As String

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 3: Call Export_Excel(frmMain.cdgExcel, sprGrid)
        
        Case 4:
            Rtn = MsgBox("출력 미리보기를 하시겠습니까?", vbQuestion + vbYesNo, "출력")
            
            If Rtn = vbYes Then
                Call Data_Print(True)
            Else
                Call Data_Print(False)
            End If
            
        Case 5: Unload Me
    End Select
End Sub

Private Sub Data_Print(Print_PreView As Boolean)
    On Error GoTo ErrRtn
    
    If sprGrid.MaxRows = 0 Then Exit Sub

    If Dir(AppPath & "XML", vbDirectory) = "" Then
        MkDir AppPath & "XML"
    End If

    Open AppPath & "XML\할인현황.XML" For Output As #1
    
    Print #1, "<?xml version=""1.0"" encoding=""EUC-KR""?>"
    Print #1, "<root>"
    
          XML = "    <조건>"
          
    If optSale(0).Value = True Then
        XML = XML & "        <검색조건>품목할인 현황</검색조건>"
    Else
        XML = XML & "        <검색조건>요일할인 현황</검색조건>"
    End If
    
    XML = XML & "        <가맹점>크린에이드 - " & Func_Replace(가맹점정보.가맹점명) & "</가맹점>"
    XML = XML & "   </조건>"
    Print #1, XML
    
    With sprList
        For i = 1 To .MaxRows
            .Row = i
            
                             XML = "    <Data>"
            .Col = 1:  XML = XML & "        <시작일자>" & .Text & "</시작일자>"
            .Col = 2:  XML = XML & "        <종료일자>" & .Text & "</종료일자>"
            .Col = 3:  XML = XML & "        <요일>" & .Text & "</요일>"
            .Col = 4:  XML = XML & "        <의류코드>" & Func_Replace(.Text) & "</의류코드>"
            .Col = 5:  XML = XML & "        <의류명>" & Func_Replace(.Text) & "</의류명>"
            .Col = 6:  XML = XML & "        <금액>" & .Text & "</금액>"
            .Col = 7:  XML = XML & "        <할인금액>" & .Text & "</할인금액>"
            .Col = 8:  XML = XML & "        <할인율>" & .Text & "</할인율>"
            .Col = 9:  XML = XML & "        <순서>" & Func_Replace(.Text) & "</순서>"
                       XML = XML & "   </Data>"
                       Print #1, XML
        Next i
        
        Print #1, "</root>"
        Close #1
    End With
    
    If Print_PreView = True Then
        With rpt할인현황
            .dc.FileURL = AppPath & "XML\할인현황.XML"
            .Show 1
        End With
    Else
        With rpt할인현황
            .dc.FileURL = AppPath & "XML\할인현황.XML"
            .PrintReport False
        End With
        
        Unload rpt할인현황
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub Form_Load()
    With sprGrid
        .MaxRows = 0
        .RowHeight(-1) = 14
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeSingle
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
    
    With sprList
        .MaxRows = 0
        .RowHeight(-1) = 14
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeSingle
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    cmdBtn(5).Left = Me.Width - cmdBtn(5).Width - 200
End Sub

Private Sub optSale_Click(Index As Integer, Value As Integer)
    sprList.MaxRows = 0
    
    If Index = 0 Then
        Call 품목할인_Display
    Else
        Call 요일할인_Display
    End If
End Sub

Private Sub 품목할인_Display()
    Dim sDate   As String
    
    sDate = Format(Date, "yyyy-MM-dd")
    
    Query = "SELECT DISTINCT 시작일자, 종료일자"
    Query = Query & " FROM TB_할인정보"
    Query = Query & " ORDER BY 시작일자 desc, 종료일자"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = ADORs!시작일자 & ""
            .Col = 2: .Text = ADORs!종료일자 & ""
            .Col = 3: .Text = ""
            .Col = 4: .Text = ""
            
            ' 적용 대상일 경우
            If ADORs!시작일자 & "" <= sDate And ADORs!종료일자 & "" >= sDate Then
                .Col = -1
                .BackColor = vbGreen
            End If
            
        
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        .ReDraw = True
    End With
End Sub

Private Sub 요일할인_Display()
    Dim sDate   As String
    
    sDate = Format(Date, "yyyy-MM-dd")
    
    Query = "SELECT DISTINCT 시작일자, 종료일자, 요일"
    Query = Query & " FROM TB_요일할인"
    Query = Query & " ORDER BY 시작일자 desc, 종료일자, 요일"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = ADORs!시작일자 & ""       '
            .Col = 2: .Text = ADORs!종료일자 & ""       '
            .Col = 3: .Text = ADORs!요일 & ""           '
            .Col = 4: .Text = Fun_Week(ADORs!요일) & "" '
        
            ' 적용 대상일 경우
            If ADORs!시작일자 & "" <= sDate And ADORs!종료일자 & "" >= sDate Then
                .Col = -1
                .BackColor = vbGreen
            End If
        
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        .ReDraw = True
    End With
End Sub

Private Sub sprGrid_Click(ByVal Col As Long, ByVal Row As Long)
    If Row <= 0 Then Exit Sub
    
    With sprGrid
        .Row = Row
        .Col = 1: 시작일자 = Trim(.Text) & ""
        .Col = 2: 종료일자 = Trim(.Text) & ""
        .Col = 3: 요일 = Trim(.Text) & ""
    End With
    
    Call Data_Display
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn
    
    If optSale(0).Value = True Then
        Query = "SELECT    A.의류코드"
        Query = Query & ", A.의류명"
        Query = Query & ", B.금액"
        Query = Query & ", A.할인금액"
        Query = Query & ", A.할인율"
        Query = Query & ", A.순서"
        Query = Query & " FROM TB_할인정보 AS A LEFT OUTER JOIN TB_의류 AS B ON A.의류코드 = B.의류코드"
        Query = Query & " WHERE (A.시작일자 = '" & 시작일자 & "'"
        Query = Query & "   AND  A.종료일자 = '" & 종료일자 & "')"
        Query = Query & " ORDER BY A.의류코드 ASC"
    Else
        Query = "SELECT    A.의류코드"
        Query = Query & ", A.의류명"
        Query = Query & ", B.금액"
        Query = Query & ", A.할인금액"
        Query = Query & ", A.할인율"
        Query = Query & ", A.순서"
        Query = Query & " FROM TB_요일할인 AS A LEFT OUTER JOIN TB_의류 AS B ON A.의류코드 = B.의류코드"
        Query = Query & " WHERE (A.시작일자 = '" & 시작일자 & "'"
        Query = Query & "   AND  A.종료일자 = '" & 종료일자 & "')"
        Query = Query & "   AND  A.요일 = '" & 요일 & "'"
        Query = Query & " ORDER BY A.의류코드 ASC"
    End If
    
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprList
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = 시작일자 & ""        '
            .Col = 2: .Text = 종료일자 & ""        '
            .Col = 3: .Text = Fun_Week(요일) & ""  '
            .Col = 4: .Text = ADORs!의류코드 & ""  '
            .Col = 5: .Text = ADORs!의류명 & ""    '
            .Col = 6: .Text = ADORs!금액 & ""      '
            .Col = 7: .Text = ADORs!할인금액 & ""  '
            .Col = 8: .Text = ADORs!할인율 & ""    '
            .Col = 9: .Text = ADORs!순서 & ""      '
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub SSPanel1_Click()
    Dim sMsg As String
    Dim Query   As String
    
    sMsg = InputBox("암호를 입력 하여 주십시요", "암호 확인")
    If sMsg <> "cleanaid1996" Then Exit Sub
    
    
    sMsg = InputBox("삭제할 시작일자를 입력 하여 주십시요.", "삭제 시작일자")
    
    If IsDate(sMsg) = False Then Exit Sub
    
    
    
    If optSale(0).Value = True Then
        Query = "DELETE  FROM TB_할인정보"
        Query = Query & " WHERE 시작일자 = '" & sMsg & "'"
    Else
        Query = "DELETE  FROM TB_요일할인"
        Query = Query & " WHERE 시작일자 = '" & sMsg & "'"
    End If
    
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
End Sub
