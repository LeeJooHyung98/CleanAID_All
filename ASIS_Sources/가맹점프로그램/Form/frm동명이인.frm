VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm동명이인 
   BorderStyle     =   1  '단일 고정
   Caption         =   "고객 검색"
   ClientHeight    =   6165
   ClientLeft      =   2745
   ClientTop       =   5820
   ClientWidth     =   11160
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form30"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   11160
   StartUpPosition =   1  '소유자 가운데
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   6165
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   10874
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm동명이인.frx":0000
      Begin Threed.SSPanel SSPanel1 
         Height          =   555
         Left            =   15
         TabIndex        =   3
         Top             =   5595
         Width           =   11130
         _ExtentX        =   19632
         _ExtentY        =   979
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   480
            Index           =   0
            Left            =   30
            TabIndex        =   4
            Top             =   30
            Visible         =   0   'False
            Width           =   1365
            _Version        =   851970
            _ExtentX        =   2408
            _ExtentY        =   847
            _StockProps     =   79
            Caption         =   "선택"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   480
            Index           =   1
            Left            =   9735
            TabIndex        =   5
            Top             =   30
            Width           =   1365
            _Version        =   851970
            _ExtentX        =   2408
            _ExtentY        =   847
            _StockProps     =   79
            Caption         =   "취소"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   390
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   11130
         _ExtentX        =   19632
         _ExtentY        =   688
         _Version        =   262144
         Font3D          =   3
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "    고객 검색"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm동명이인.frx":0072
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image1 
            Height          =   240
            Left            =   60
            Picture         =   "frm동명이인.frx":0381
            Top             =   75
            Width           =   240
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   5160
         Left            =   15
         TabIndex        =   0
         Top             =   420
         Width           =   11130
         _Version        =   524288
         _ExtentX        =   19632
         _ExtentY        =   9102
         _StockProps     =   64
         BackColorStyle  =   1
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         MaxCols         =   5
         MaxRows         =   5
         ScrollBars      =   2
         SpreadDesigner  =   "frm동명이인.frx":090B
         HighlightHeaders=   1
         HighlightStyle  =   1
         ScrollBarStyle  =   2
      End
   End
End
Attribute VB_Name = "frm동명이인"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SELECTCODE As String

Public Sub DataDisplay(Query As String)
    i = 0
    
    Set Rs = New ADODB.Recordset
    Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With spdView
        .MaxRows = 0
        
        Do Until Rs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = Rs!성명 & ""
            .Col = 2: .Text = Rs!전화번호 & ""
            .Col = 3: .Text = Rs!휴대전화 & ""
            .Col = 4: .Text = Rs!주소 & ""
            .Col = 5: .Text = Rs!고객코드 & ""
            
            Rs.MoveNext
        Loop
        Rs.Close
        Set Rs = Nothing
    End With
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 0: Call spdView_DblClick(spdView.ActiveCol, spdView.ActiveRow)
                        
        Case 1
            고객정보.전화번호 = "Error"
            SELECTCODE = "CANCEL"
            
            Unload Me
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        고객정보.전화번호 = "Error"
        SELECTCODE = "CANCEL"
        
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    SELECTCODE = ""
    
    With spdView
        .MaxRows = 0
        .RowHeight(-1) = 18
        
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
    
    'Me.Left = (Screen.Width - Me.Width) / 2
    'Me.Top = (Screen.Height - Me.Height) / 2
End Sub

Private Sub spdView_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim tempStr As String
    
    If Len(GetSpreadText(spdView, Row, 5)) <= 0 Then Exit Sub
    
    tempStr = Get_고객정보(GetSpreadText(spdView, Row, 5))
    
    If tempStr <> "Error" Then
        SELECTCODE = 고객정보.고객코드
        
        Unload Me
    End If
End Sub

Private Sub spdView_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call spdView_DblClick(spdView.ActiveCol, spdView.ActiveRow)
    End If
End Sub
