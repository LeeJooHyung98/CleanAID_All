VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm마일리지 
   BorderStyle     =   1  '단일 고정
   Caption         =   "마일리지"
   ClientHeight    =   4575
   ClientLeft      =   540
   ClientTop       =   2730
   ClientWidth     =   11205
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm마일리지.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   11205
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   8070
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   0
      PaneTree        =   "frm마일리지.frx":0A02
      Begin Threed.SSPanel SSPanel 
         Height          =   570
         Index           =   1
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   11205
         _ExtentX        =   19764
         _ExtentY        =   1005
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   0
            Left            =   975
            TabIndex        =   2
            Top             =   120
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   63963139
            CurrentDate     =   40279
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   1
            Left            =   2685
            TabIndex        =   3
            Top             =   120
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   63963139
            CurrentDate     =   40279
         End
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   465
            Left            =   4380
            TabIndex        =   6
            Top             =   60
            Width           =   1305
            _Version        =   851970
            _ExtentX        =   2302
            _ExtentY        =   820
            _StockProps     =   79
            Caption         =   " 조회(&F)"
            Appearance      =   6
            Picture         =   "frm마일리지.frx":0A54
         End
         Begin XtremeSuiteControls.PushButton btnExit 
            Height          =   465
            Left            =   9855
            TabIndex        =   7
            Top             =   60
            Width           =   1305
            _Version        =   851970
            _ExtentX        =   2302
            _ExtentY        =   820
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frm마일리지.frx":0FEE
         End
         Begin VB.Label lblCode 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "0"
            ForeColor       =   &H00808080&
            Height          =   180
            Left            =   7290
            TabIndex        =   8
            Top             =   210
            Width           =   90
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "~"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   2475
            TabIndex        =   5
            Top             =   180
            Width           =   120
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "발생일자:"
            Height          =   225
            Index           =   2
            Left            =   60
            TabIndex        =   4
            Top             =   195
            Width           =   870
         End
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   3990
         Left            =   0
         TabIndex        =   9
         Top             =   585
         Width           =   11205
         _Version        =   524288
         _ExtentX        =   19764
         _ExtentY        =   7038
         _StockProps     =   64
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
         MaxCols         =   9
         ScrollBars      =   2
         SpreadDesigner  =   "frm마일리지.frx":1588
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "frm마일리지"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub cmdList_Click()
    On Error GoTo ErrRtn
    
    With sprGrid
        .MaxRows = 0
        .ReDraw = False

        '-------------------------------------------------------------------------------------
        ' 이전누계
        '-------------------------------------------------------------------------------------
        Query = "SELECT TOP 1 "
        Query = Query & "  ISNULL(누적마일리지,0)"
        Query = Query & ", ISNULL(사용가능마일리지,0)"
        Query = Query & " FROM TB_매출"
        Query = Query & " WHERE 고객코드 = '" & lblCode.Caption & "'"
        Query = Query & "   AND 매출일자 < '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "'"
        Query = Query & "   AND (발생마일리지 <> 0 OR 사용마일리지 <> 0)"
        Query = Query & " ORDER BY 매출일자 DESC, 매출시간 DESC"
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        If Not ADORs.EOF Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
    
            .Col = 1: .Text = "이전누계"
            .Col = 2: .Text = ""
            .Col = 3: .Text = ""
            .Col = 4: .Text = ""
            .Col = 5: .Text = ""
            .Col = 6: .Text = ""
            .Col = 7: .Text = ""
            .Col = 8: .Text = ADORs(0) & ""
            .Col = 9: .Text = ADORs(1) & ""
        End If
        ADORs.Close
        Set ADORs = Nothing
        
        
        '-------------------------------------------------------------------------------------
        ' 마일리지 - 판매취소 등이 있기 때문에 "매출일자, 매출시간 순으로 보여줘야 한다."
        '-------------------------------------------------------------------------------------
        Query = "SELECT     매출일자"
        Query = Query & " , 접수금액"
        Query = Query & " , 현금입금"
        Query = Query & " , 카드입금"
        Query = Query & " , 발생마일리지"
        Query = Query & " , 사용마일리지"
        Query = Query & " , 누적마일리지"
        Query = Query & " , 사용가능마일리지"
        Query = Query & " FROM TB_매출"
        Query = Query & " WHERE 고객코드   = '" & lblCode.Caption & "'"
        Query = Query & "   AND (매출일자 >= '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "'"
        Query = Query & "   AND  매출일자 <= '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "')"
        Query = Query & "   AND (발생마일리지 <> 0 OR 사용마일리지 <> 0)"
        Query = Query & " ORDER BY 매출일자, 매출시간"
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows

            .Col = 1: .Text = ADORs!매출일자 & ""
            .Col = 2: .Text = ADORs!접수금액 & ""
            .Col = 3: .Text = ADORs!현금입금 & ""
            .Col = 4: .Text = ADORs!카드입금 & ""
            .Col = 5: .Text = ""
            .Col = 6: .Text = ADORs!사용마일리지 & ""
            .Col = 7: .Text = ADORs!발생마일리지 & ""
            .Col = 8: .Text = ADORs!누적마일리지 & ""
            .Col = 9: .Text = ADORs!사용가능마일리지 & ""

            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing

        .ReDraw = True
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
    
    Screen.MousePointer = 0
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
        .OperationMode = OperationModeExtended
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With

    dtpDay(0).Value = Format(Date, "YYYY-MM-DD")
    dtpDay(1).Value = Format(Date, "YYYY-MM-DD")
    
    
    
    
End Sub
