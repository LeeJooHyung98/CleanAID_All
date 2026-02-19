VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_01004_C 
   Caption         =   "본사 가맹점 품목할인 관리"
   ClientHeight    =   12345
   ClientLeft      =   240
   ClientTop       =   3435
   ClientWidth     =   16215
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_01004_C.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12345
   ScaleWidth      =   16215
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   12345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   21775
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_01004_C.frx":058A
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   1
         Top             =   540
         Width           =   16185
         _ExtentX        =   28549
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1125
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   "cboOffice"
            Top             =   75
            Width           =   3405
         End
         Begin XtremeSuiteControls.ProgressBar ProgressBar2 
            Height          =   330
            Left            =   9900
            TabIndex        =   19
            Top             =   390
            Visible         =   0   'False
            Width           =   5325
            _Version        =   851970
            _ExtentX        =   9393
            _ExtentY        =   582
            _StockProps     =   93
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label lblProgress 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "할인품목 저장:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   1
            Left            =   8610
            TabIndex        =   20
            Top             =   480
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.Label lblTitle 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "지 사 명:"
            Height          =   225
            Index           =   24
            Left            =   0
            TabIndex        =   13
            Top             =   135
            Width           =   1065
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   900
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   255
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "본사 가맹점 품목할인 관리 (P_01004_C)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_01004_C.frx":069C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Left            =   8610
         TabIndex        =   3
         Top             =   15
         Width           =   7590
         _ExtentX        =   13388
         _ExtentY        =   900
         _Version        =   262144
         Font3D          =   3
         ForeColor       =   192
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureBackgroundStyle=   2
         PictureBackground=   "P_01004_C.frx":089E
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   4
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "종료"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_01004_C.frx":0AA0
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   5
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "엑셀"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_01004_C.frx":103A
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   6
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "인쇄"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_01004_C.frx":15D4
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   7
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "취소"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_01004_C.frx":1B6E
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   8
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "삭제"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_01004_C.frx":2108
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   9
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "저장"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_01004_C.frx":26A2
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   10
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "신규"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_01004_C.frx":2C3C
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   11
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "조회"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_01004_C.frx":31D6
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   4650
         Left            =   15
         TabIndex        =   14
         Top             =   1335
         Width           =   3750
         _Version        =   524288
         _ExtentX        =   6615
         _ExtentY        =   8202
         _StockProps     =   64
         BackColorStyle  =   1
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         MaxCols         =   2
         Protect         =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "P_01004_C.frx":3770
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdList 
         Height          =   6330
         Left            =   15
         TabIndex        =   15
         Top             =   6000
         Width           =   3750
         _Version        =   524288
         _ExtentX        =   6615
         _ExtentY        =   11165
         _StockProps     =   64
         BackColorStyle  =   1
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         MaxCols         =   3
         Protect         =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "P_01004_C.frx":3C90
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panCaption 
         Height          =   390
         Index           =   2
         Left            =   3780
         TabIndex        =   16
         Top             =   1335
         Width           =   12420
         _ExtentX        =   21908
         _ExtentY        =   688
         _Version        =   262144
         Font3D          =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " 가맹점 품목할인 현황"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_01004_C.frx":428F
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
      End
      Begin FPSpreadADO.fpSpread sprClass 
         Height          =   10590
         Left            =   3780
         TabIndex        =   17
         Top             =   1740
         Width           =   3585
         _Version        =   524288
         _ExtentX        =   6324
         _ExtentY        =   18680
         _StockProps     =   64
         BackColorStyle  =   1
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         MaxCols         =   2
         ScrollBars      =   2
         SpreadDesigner  =   "P_01004_C.frx":46F1
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdView2 
         Height          =   10590
         Left            =   7380
         TabIndex        =   18
         Top             =   1740
         Width           =   8820
         _Version        =   524288
         _ExtentX        =   15558
         _ExtentY        =   18680
         _StockProps     =   64
         BackColorStyle  =   1
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         MaxCols         =   6
         ScrollBars      =   2
         SpreadDesigner  =   "P_01004_C.frx":4C13
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_01004_C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim 가맹점코드   As String
Dim 시작일자     As String
Dim 종료일자     As String

Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cboOffice_Click()
    Call Data_Display(Mid(cboOffice.Text, 2, 4))
End Sub

Private Sub Data_Display(지사코드 As String)
    On Error GoTo ErrRtn

    ReDim sValue(2)
    
    sValue(0) = 지사코드 'HeadOffice
    sValue(1) = Format(Date, "YYYY-MM-DD")
    sValue(2) = Format(Date, "YYYY-MM-DD")
        
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01001_00", sValue(), Err_Num, Err_Dec)
    
    With spdView
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            If (RS01!종료일자 = "") Or (RS01!종료일자 = "2099-12-31") Then '
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                
                .Col = 1: .Text = RS01!가맹점코드 & "" ' 1
                .Col = 2: .Text = RS01!가맹점명 & ""   ' 2
            End If
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
    
    Call Data_Display3
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

'---------------------------------------------------------------------
' SP_01004_00 - TB_가맹점할인
'---------------------------------------------------------------------
Private Sub Data_Display2(가맹점코드 As String)
    ReDim sValue(0)
    
    sValue(0) = 가맹점코드
        
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01004_00", sValue(), Err_Num, Err_Dec)
    
    With spdList
        .MaxRows = 0
        .Redraw = False
                    
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = RS01!시작일자 & ""
            .Col = 2: .Text = RS01!종료일자 & ""
            .Col = 3: .Text = RS01!할인율 & ""
            
            ' 적용 대상일 경우
            If RS01!시작일자 & "" <= Format(Date, "yyyy-MM-dd") And RS01!종료일자 & "" >= Format(Date, "yyyy-MM-dd") Then
                .Col = -1
                .BackColor = vbGreen
            End If
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .SortKey(1) = 1
        .SortKeyOrder(1) = SortKeyOrderDescending
        .Sort -1, -1, -1, -1, SortByRow
            
        .Redraw = True
    End With
End Sub

Private Sub Data_Display3()
    '-------------------------------------------------------------
    ' TB_의류분류
    '-------------------------------------------------------------
    ReDim sValue(0)

    sValue(0) = "0"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_00013", sValue(), Err_Num, Err_Dec)

    With sprClass
        .MaxRows = 0
        .Redraw = False
                    
        '-전체-
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        
        .Col = 1: .Text = ""
        .Col = 2: .Text = "-전체-"
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = RS01!의류분류코드 & ""
            .Col = 2: .Text = RS01!의류분류명 & ""
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
End Sub


Private Sub Data_Display4(가맹점코드 As String, 시작일자 As String, 의류분류코드 As String)
    ReDim sValue(2)

    sValue(0) = 가맹점코드
    sValue(1) = 시작일자
    sValue(2) = 의류분류코드
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01004_06", sValue(), Err_Num, Err_Dec)

    With spdView2
        .MaxRows = 0
        .Redraw = False
                    
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = RS01!품목코드 & ""
            .Col = 2: .Text = RS01!품목명 & ""
            .Col = 3: .Text = RS01!정상가격 & ""
            .Col = 4: .Text = RS01!할인가격 & ""
            .Col = 5: .Text = RS01!비율 & ""
            .Col = 6: .Text = RS01!수신정보 & ""
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
End Sub

Private Sub cboOffice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SearchString_한글 KeyAscii
'    Else
       ' SearchString KeyAscii
    End If

End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: 'Call Data_Display5   ' 조회
        Case 1: 'Call DataAdd        ' 신규
        Case 2:  Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
        Case 4: 'Call DataCancel     ' 취소
        Case 5: 'Call DataPrint      ' 인쇄
        Case 6: Call Export_Excel(P_00000.cdgExcel, spdView2)      ' 엑셀
        Case 7: Unload Me           ' 종료
    End Select
    
'    Me.MousePointer = 0
    
    Exit Sub
    
ErrRtn:
    Me.MousePointer = 0
    
    If Err.Number = "0" Then
        
    ElseIf Err.Number = "91" Then
        End
    Else
        Resume Next
    End If
End Sub

Private Sub Form_Activate()
    '본사 여부
    If HeadOffice = MASTER_OFFICE_CODE Then
        cboOffice.Locked = False
    Else
        cboOffice.Locked = True
    End If

    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    With spdView
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

    With spdList
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
    
    With sprClass
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
    
    With spdView2
        .MaxRows = 0
        .RowHeight(-1) = 14
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        '.OperationMode = OperationModeSingle
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
    
    
    Call Get_지사리스트(cboOffice)
    
    Dim i As Integer
    
    With cboOffice
        For i = 0 To .ListCount - 1
            If Mid(.List(i), 2, 4) = HeadOffice Then
                .ListIndex = i
                
                Exit For
            End If
        Next i
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_01004_C_Flag = False
End Sub


Public Sub DataSave()
    Dim iRow       As Long
    
    On Error GoTo ErrRtn
    
    lblProgress(1).Visible = True
    
    ProgressBar2.Value = 0
    ProgressBar2.Min = 0
    ProgressBar2.Max = spdView2.DataRowCnt
    ProgressBar2.Visible = True
    
    Set RS01 = New ADODB.Recordset
    
    For iRow = 1 To spdView2.DataRowCnt
        spdView2.Row = iRow
        spdView2.Col = 1
        
        ProgressBar2.Value = iRow
        DoEvents
       
       '-------------------------------------------------------------------
        ' TB_가맹점할인 저장 - SP_01004_A_09
        '-------------------------------------------------------------------
        ReDim sValue(7)
        
        sValue(0) = 가맹점코드                             ' 0 가매점코드
        sValue(1) = 시작일자                               ' 1 시작일자
        sValue(2) = 종료일자                               ' 2 종료일자
        
        spdView2.Col = 1:   spdView2.Row = iRow:    sValue(3) = Trim(spdView2.Text) & ""                ' 3 의류코드
        spdView2.Col = 2:   spdView2.Row = iRow:    sValue(4) = Trim(spdView2.Text) & ""                ' 4 의류명
        spdView2.Col = 3:   spdView2.Row = iRow:    sValue(5) = Trim(spdView2.Text) & ""                ' 5 정상가격
        spdView2.Col = 4:   spdView2.Row = iRow:    sValue(6) = Trim(spdView2.Text) & ""                ' 6 할인가격
        spdView2.Col = 5:   spdView2.Row = iRow:    sValue(7) = Trim(spdView2.Text) & ""                ' 7 할인률
        
        Call ExecPro("SP_01004_A_04", sValue(), Err_Num, Err_Dec)
    
        If Err_Num <> 0 Then
            MsgBox "[" & Err_Num & "] " & Err_Dec
            
            ProgressBar2.Visible = False
            Exit Sub
        End If
        
    Next iRow
        
    lblProgress(1).Visible = False
    ProgressBar2.Visible = False
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub
Private Sub DataCancel()

End Sub

'-------------------------------------------------------------------
' 할인 현황 Spread
'-------------------------------------------------------------------
Private Sub spdList_Click(ByVal Col As Long, ByVal Row As Long)
    If Row <= 0 Then Exit Sub
    
    Call 의류_Display
End Sub

Private Sub spdList_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Call spdList_Click(NewCol, NewRow)
End Sub

'-------------------------------------------------------------------
' 가맹점 Spread
'-------------------------------------------------------------------
Private Sub spdView_Click(ByVal Col As Long, ByVal Row As Long)
    
    If Row <= 0 Then Exit Sub
    
    spdView.Row = Row
    spdView.Col = 1: 가맹점코드 = spdView.Text & ""
    
    Call Data_Display2(가맹점코드)
    
    Call 의류_Display
End Sub

Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Call spdView_Click(NewCol, NewRow)
End Sub

'-------------------------------------------------------------------
' 품목분류 Spread
'-------------------------------------------------------------------
Private Sub sprClass_Click(ByVal Col As Long, ByVal Row As Long)
    If Row <= 0 Then Exit Sub

    Call 의류_Display
End Sub

Private Sub sprClass_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Call sprClass_Click(NewCol, NewRow)
End Sub

Private Sub 의류_Display()
    Dim 의류분류코드 As String
    
    If spdList.ActiveRow <= 0 Then
        spdView2.MaxRows = 0
        Exit Sub
    End If
    
    If sprClass.ActiveRow <= 0 Then
        spdView2.MaxRows = 0
        Exit Sub
    End If
    
    spdList.Row = spdList.ActiveRow
    spdList.Col = 1: 시작일자 = spdList.Text & ""
    
    spdList.Row = spdList.ActiveRow
    spdList.Col = 2: 종료일자 = spdList.Text & ""
    
    sprClass.Row = sprClass.ActiveRow
    sprClass.Col = 1: 의류분류코드 = sprClass.Text & ""
    
    Call Data_Display4(가맹점코드, 시작일자, 의류분류코드)
End Sub
