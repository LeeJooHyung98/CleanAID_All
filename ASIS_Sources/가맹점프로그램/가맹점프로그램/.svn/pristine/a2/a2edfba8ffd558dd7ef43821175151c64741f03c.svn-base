VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm품목집계 
   BorderStyle     =   1  '단일 고정
   Caption         =   "품목별 접수집계"
   ClientHeight    =   6165
   ClientLeft      =   2745
   ClientTop       =   5820
   ClientWidth     =   6255
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form30"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   6255
   StartUpPosition =   1  '소유자 가운데
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   6165
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   10874
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm품목집계.frx":0000
      Begin Threed.SSPanel SSPanel1 
         Height          =   555
         Left            =   15
         TabIndex        =   1
         Top             =   5595
         Width           =   6225
         _ExtentX        =   10980
         _ExtentY        =   979
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton btnExit 
            Height          =   465
            Left            =   4875
            TabIndex        =   3
            Top             =   45
            Width           =   1305
            _Version        =   851970
            _ExtentX        =   2302
            _ExtentY        =   820
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frm품목집계.frx":0052
         End
      End
      Begin FPSpreadADO.fpSpread sprCloth 
         Height          =   5565
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   6225
         _Version        =   524288
         _ExtentX        =   10980
         _ExtentY        =   9816
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BackColorStyle  =   1
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
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         MaxCols         =   4
         MaxRows         =   30
         Protect         =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frm품목집계.frx":05EC
         VisibleCols     =   3
         VisibleRows     =   30
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "frm품목집계"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    With sprCloth
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
End Sub
