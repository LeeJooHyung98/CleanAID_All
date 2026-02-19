VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm접수오점표시 
   BorderStyle     =   1  '단일 고정
   Caption         =   "오점 표시"
   ClientHeight    =   10485
   ClientLeft      =   7065
   ClientTop       =   3240
   ClientWidth     =   11025
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm접수오점표시.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10485
   ScaleWidth      =   11025
   StartUpPosition =   1  '소유자 가운데
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   10485
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   18494
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   0
      PaneTree        =   "frm접수오점표시.frx":058A
      Begin Threed.SSPanel SSPanel4 
         Height          =   7200
         Left            =   0
         TabIndex        =   34
         Top             =   2745
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   12700
         _Version        =   262144
         BackColor       =   16777215
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "치마"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   180
            Index           =   2
            Left            =   45
            TabIndex        =   37
            Top             =   3210
            Width           =   390
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "바지"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   180
            Index           =   1
            Left            =   45
            TabIndex        =   36
            Top             =   1575
            Width           =   390
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "상의"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   180
            Index           =   0
            Left            =   45
            TabIndex        =   35
            Top             =   90
            Width           =   390
         End
         Begin VB.Image imgClass 
            Height          =   1125
            Index           =   2
            Left            =   120
            Picture         =   "frm접수오점표시.frx":063C
            Stretch         =   -1  'True
            Top             =   3435
            Width           =   1125
         End
         Begin VB.Image imgClass 
            Height          =   1125
            Index           =   1
            Left            =   150
            Picture         =   "frm접수오점표시.frx":0F6C
            Stretch         =   -1  'True
            Top             =   1860
            Width           =   1125
         End
         Begin VB.Image imgClass 
            Height          =   1125
            Index           =   0
            Left            =   120
            Picture         =   "frm접수오점표시.frx":220F
            Stretch         =   -1  'True
            Top             =   255
            Width           =   1125
         End
      End
      Begin VB.PictureBox picCapture 
         Appearance      =   0  '평면
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   7200
         Left            =   1410
         Picture         =   "frm접수오점표시.frx":2F9C
         ScaleHeight     =   7200
         ScaleWidth      =   9615
         TabIndex        =   32
         Top             =   2745
         Width           =   9615
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   525
         Left            =   0
         TabIndex        =   21
         Top             =   9960
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   926
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CommandButton cmdStop 
            Caption         =   "STOP"
            Height          =   375
            Left            =   2190
            TabIndex        =   24
            Top             =   60
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.ListBox lstDevices 
            Height          =   240
            Left            =   3300
            TabIndex        =   23
            Top             =   135
            Visible         =   0   'False
            Width           =   1440
         End
         Begin XtremeSuiteControls.PushButton btnCapture 
            Height          =   450
            Left            =   8385
            TabIndex        =   22
            Top             =   30
            Width           =   1395
            _Version        =   851970
            _ExtentX        =   2461
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "사진찍기"
            ForeColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Picture         =   "frm접수오점표시.frx":68D4
         End
         Begin XtremeSuiteControls.PushButton btnSave 
            Height          =   450
            Left            =   9795
            TabIndex        =   25
            Top             =   30
            Width           =   1170
            _Version        =   851970
            _ExtentX        =   2064
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 저장"
            ForeColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Picture         =   "frm접수오점표시.frx":69AB
         End
         Begin XtremeSuiteControls.PushButton cmdStart 
            Height          =   450
            Left            =   45
            TabIndex        =   26
            Top             =   30
            Visible         =   0   'False
            Width           =   1740
            _Version        =   851970
            _ExtentX        =   3069
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 캠으로 찍기"
            ForeColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Picture         =   "frm접수오점표시.frx":73BD
         End
         Begin VB.Label lblWebCAM_No 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4770
            TabIndex        =   27
            Top             =   105
            Visible         =   0   'False
            Width           =   120
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   2205
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   3889
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox cboStain 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7665
            TabIndex        =   29
            Text            =   "cboStain"
            Top             =   90
            Width           =   1995
         End
         Begin FPSpreadADO.fpSpread sprGrid 
            Height          =   1695
            Left            =   7665
            TabIndex        =   30
            Top             =   435
            Width           =   2910
            _Version        =   524288
            _ExtentX        =   5133
            _ExtentY        =   2990
            _StockProps     =   64
            BackColorStyle  =   1
            DAutoCellTypes  =   0   'False
            DAutoHeadings   =   0   'False
            DAutoSave       =   0   'False
            DisplayRowHeaders=   0   'False
            EditModePermanent=   -1  'True
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
            MaxCols         =   2
            MaxRows         =   200
            ScrollBars      =   2
            SpreadDesigner  =   "frm접수오점표시.frx":762F
            UserResize      =   1
            VisibleCols     =   2
            VisibleRows     =   30
            Appearance      =   1
            HighlightHeaders=   1
            HighlightStyle  =   1
            ScrollBarStyle  =   2
         End
         Begin XtremeSuiteControls.PushButton btnADD 
            Height          =   345
            Left            =   9690
            TabIndex        =   31
            Top             =   60
            Width           =   885
            _Version        =   851970
            _ExtentX        =   1561
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   " 추가"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm접수오점표시.frx":7C1F
         End
         Begin FPSpreadADO.fpSpread sprColor 
            Height          =   2040
            Left            =   840
            TabIndex        =   33
            Top             =   90
            Width           =   3000
            _Version        =   524288
            _ExtentX        =   5292
            _ExtentY        =   3598
            _StockProps     =   64
            BackColorStyle  =   1
            DAutoCellTypes  =   0   'False
            DAutoHeadings   =   0   'False
            DAutoSave       =   0   'False
            DisplayColHeaders=   0   'False
            DisplayRowHeaders=   0   'False
            EditModePermanent=   -1  'True
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
            MaxCols         =   4
            MaxRows         =   4
            ScrollBars      =   0
            SpreadDesigner  =   "frm접수오점표시.frx":8631
            UserResize      =   1
            VisibleCols     =   1
            VisibleRows     =   4
            Appearance      =   1
            HighlightHeaders=   1
            HighlightStyle  =   1
            ScrollBarStyle  =   2
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "색상표:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   1
            Left            =   -30
            TabIndex        =   38
            Top             =   150
            Width           =   810
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "오점내용:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   0
            Left            =   6600
            TabIndex        =   28
            Top             =   150
            Width           =   1020
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   510
         Left            =   0
         TabIndex        =   1
         Top             =   2220
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   900
         _Version        =   262144
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton btnTool 
            Height          =   420
            Index           =   0
            Left            =   45
            TabIndex        =   4
            Top             =   45
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frm접수오점표시.frx":8BB0
         End
         Begin VB.PictureBox picColorL 
            Appearance      =   0  '평면
            BackColor       =   &H80000008&
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   6345
            ScaleHeight     =   285
            ScaleWidth      =   300
            TabIndex        =   2
            Top             =   45
            Width           =   330
         End
         Begin VB.PictureBox picColorR 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   6435
            ScaleHeight     =   285
            ScaleWidth      =   300
            TabIndex        =   3
            Top             =   135
            Width           =   330
         End
         Begin XtremeSuiteControls.PushButton btnTool 
            Height          =   420
            Index           =   1
            Left            =   465
            TabIndex        =   5
            Top             =   45
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frm접수오점표시.frx":91DA
         End
         Begin XtremeSuiteControls.PushButton btnTool 
            Height          =   420
            Index           =   2
            Left            =   885
            TabIndex        =   6
            Top             =   45
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frm접수오점표시.frx":9804
         End
         Begin XtremeSuiteControls.PushButton btnTool 
            Height          =   420
            Index           =   3
            Left            =   1305
            TabIndex        =   7
            Top             =   45
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frm접수오점표시.frx":9E2E
         End
         Begin XtremeSuiteControls.PushButton btnTool 
            Height          =   420
            Index           =   4
            Left            =   1725
            TabIndex        =   8
            Top             =   45
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frm접수오점표시.frx":A458
         End
         Begin XtremeSuiteControls.PushButton btnTool 
            Height          =   420
            Index           =   5
            Left            =   2145
            TabIndex        =   9
            Top             =   45
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frm접수오점표시.frx":AA82
         End
         Begin XtremeSuiteControls.PushButton btnTool 
            Height          =   420
            Index           =   6
            Left            =   2565
            TabIndex        =   10
            Top             =   45
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frm접수오점표시.frx":B0AC
         End
         Begin XtremeSuiteControls.PushButton btnTool 
            Height          =   420
            Index           =   7
            Left            =   2985
            TabIndex        =   11
            Top             =   45
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frm접수오점표시.frx":B6D6
         End
         Begin XtremeSuiteControls.PushButton btnTool 
            Height          =   420
            Index           =   8
            Left            =   3405
            TabIndex        =   12
            Top             =   45
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frm접수오점표시.frx":BD00
         End
         Begin XtremeSuiteControls.PushButton btnTool 
            Height          =   420
            Index           =   9
            Left            =   3825
            TabIndex        =   13
            Top             =   45
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frm접수오점표시.frx":C32A
         End
         Begin XtremeSuiteControls.PushButton btnTool 
            Height          =   420
            Index           =   10
            Left            =   4245
            TabIndex        =   14
            Top             =   45
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frm접수오점표시.frx":C954
         End
         Begin XtremeSuiteControls.PushButton btnTool 
            Height          =   420
            Index           =   11
            Left            =   4665
            TabIndex        =   15
            Top             =   45
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frm접수오점표시.frx":CF7E
         End
         Begin XtremeSuiteControls.PushButton btnTool 
            Height          =   420
            Index           =   12
            Left            =   5085
            TabIndex        =   16
            Top             =   45
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frm접수오점표시.frx":D5A8
         End
         Begin XtremeSuiteControls.PushButton btnTool 
            Height          =   420
            Index           =   15
            Left            =   5505
            TabIndex        =   17
            Top             =   45
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frm접수오점표시.frx":DBD2
         End
         Begin VB.Label lblTag 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "#"
            ForeColor       =   &H00808080&
            Height          =   180
            Left            =   9330
            TabIndex        =   20
            Top             =   60
            Width           =   90
         End
         Begin VB.Label lblDate 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "#"
            ForeColor       =   &H00808080&
            Height          =   180
            Left            =   9330
            TabIndex        =   19
            Top             =   300
            Width           =   90
         End
      End
   End
End
Attribute VB_Name = "frm접수오점표시"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const WM_CAP As Integer = &H400

Const WM_CAP_DRIVER_CONNECT    As Long = WM_CAP + 10
Const WM_CAP_DRIVER_DISCONNECT As Long = WM_CAP + 11
Const WM_CAP_EDIT_COPY         As Long = WM_CAP + 30

Const WM_CAP_SET_PREVIEW     As Long = WM_CAP + 50
Const WM_CAP_SET_PREVIEWRATE As Long = WM_CAP + 52
Const WM_CAP_SET_SCALE       As Long = WM_CAP + 53

Const WS_CHILD     As Long = &H40000000
Const WS_VISIBLE   As Long = &H10000000
Const SWP_NOMOVE   As Long = &H2
Const SWP_NOSIZE   As Integer = 1
Const SWP_NOZORDER As Integer = &H4
Const HWND_BOTTOM  As Integer = 1

Dim iDevice As Long  ' Current device ID
Dim hHwnd   As Long ' Handle to preview window

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hndw As Long) As Boolean
Private Declare Function capCreateCaptureWindowA Lib "avicap32.dll" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Integer, ByVal hwndParent As Long, ByVal nID As Long) As Long
Private Declare Function capGetDriverDescriptionA Lib "avicap32.dll" (ByVal wDriver As Long, ByVal lpszName As String, ByVal cbName As Long, ByVal lpszVer As String, ByVal cbVer As Long) As Boolean

'Paint
'---------------------------------------------------
Private blnSquareFill As Boolean
Private blnCircleFill As Boolean
Private blnElipseFill As Boolean

Private mudtTool  As UDT_Tool
Private mudtPoint As UDT_Point

'48 Color
Private Const ColorValues = "&HFFFFFF&HC0C0FF&HC0E0FF&HC0FFFF&HC0FFC0&HFFFFC0&HFFC0C0&HFFC0FF&HE0E0E0&H8080FF&H80C0FF&H80FFFF&H80FF80&HFFFF80&HFF8080&HFF80FF&HC0C0C0&H0000FF&H0080FF&H00FFFF&H00FF00&HFFFF00&HFF0000&HFF00FF&H808080&H0000C0&H0040C0&H00C0C0&H00C000&HC0C000&HC00000&HC000C0&H404040&H000080&H004080&H008080&H008000&H808000&H800000&H800080" & "&H000000&H000040&H404080&H004040&H004000&H404000&H400000&H400040"

Private Sub cmdFormat_Click()

End Sub

Private Sub cmdSource_Click()

End Sub

Private Sub btnADD_Click()
    With sprGrid
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 1: .Text = Trim(cboStain.Text) & ""
    End With
End Sub

Private Sub btnCapture_Click()
    On Error Resume Next
    
    Dim DPath As String
    Dim bm    As Image
    Dim strFile As String
    Dim SaveFile As String
    
    '
    DPath = AppPath + "Capture"
    
    If Dir(DPath, vbDirectory) = "" Then
        MkDir DPath
    End If
    
'    ' Copy image to clipboard
'    SendMessage hHwnd, WM_CAP_EDIT_COPY, 0, 0
'    ClosePreviewWindow
'
'    picCapture.Picture = Clipboard.GetData
    SaveFile = AppPath & "capture\" & lblTag.Caption & ".jpg"
    If Dir(SaveFile) <> "" Then
        Kill SaveFile
    End If
    
    strFile = Dir(AppPath & "capture.exe")
    
    
    If strFile <> "" Then
        Shell AppPath & "capture.exe " & lblTag.Caption, vbNormalFocus
    End If
    On Error Resume Next
    While True
        DoEvents
        
        If Dir(SaveFile) <> "" Then
            If Dir(SaveFile & "_complete") <> "" Then
                picCapture.Picture = LoadPicture(SaveFile)
                On Error GoTo 0
                Kill SaveFile
                Kill SaveFile & "_complete"
                Exit Sub
            End If
        End If
        Sleep (1000)
    Wend
    
End Sub

Public Function ShellWS(ByRef Command As String, Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus, _
                                                 Optional ByVal WaitOnReturn As Boolean) As Long
    #Const Referenced = True
    #If Not Referenced Then
        ShellWS = CreateObject("WScript.Shell").Run(Command, WindowStyle, WaitOnReturn)
    #Else
        With New WshShell
            ShellWS = .Run(Command, WindowStyle, WaitOnReturn)
        End With
    #End If         'Adapted from "Best Shell & Wait (No API's!)" by Matthew Roberts
End Function

Private Sub btnSave_Click()
    On Error Resume Next

    Dim DPath    As String
    Dim FileName As String
    Dim bm       As Image
    
    Dim 오류내용 As String
    
    '
    DPath = AppPath + "Capture"

    If Dir(DPath, vbDirectory) = "" Then
        MkDir DPath
    End If

    FileName = Format(lblDate.Caption, "YYYYMMDD") & lblTag.Caption & ".JPG"
    
    SavePicture picCapture.Image, AppPath & "Capture\" & FileName

    frm접수.imgCapture.Picture = LoadPicture(App.Path & "\Capture\" & FileName)

    오류내용 = ""
    
    With sprGrid
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1: 오류내용 = 오류내용 & .Text & " "
        Next i
        
        오류내용 = Trim(오류내용)
    End With
    
    frm접수.sprGrid.Row = frm접수.sprGrid.ActiveRow
    frm접수.sprGrid.Col = 19: frm접수.sprGrid.Text = 오류내용 & ""
    frm접수.sprGrid.Col = 21: frm접수.sprGrid.Text = FileName & ""
    
    Unload Me
End Sub

Private Sub btnTool_Click(Index As Integer)
    'set UDT tool to clear previous tool
    SetAllFalse mudtTool
    
    'set the width in case it chaged with the previous tool
    picCapture.DrawWidth = CurentWidth
    
    'select tool and attributes
    With mudtTool
        Select Case Index
            Case 0: .Point = True     'single point
            Case 1: .Line = True      'straight line
            Case 2: .FreeLine = True  'freehand line
            
            Case 3      'circle
                picCapture.FillStyle = vbFSTransparent
                blnCircleFill = False
                .Circle = True
                
            Case 4      'filled Circle
                picCapture.FillStyle = intFillStyle
                blnCircleFill = True
                .Circle = True
                
            Case 5      'elipse
                picCapture.FillStyle = vbFSTransparent
                blnElipseFill = False
                .Elipse = True
                
            Case 6      'elipse fill
                picCapture.FillStyle = intFillStyle
                blnElipseFill = True
                .Elipse = True
                
            Case 7      'empty rectangle
                blnSquareFill = False
                picCapture.FillStyle = vbFSTransparent
                .Square = True
                
            Case 8      'filled rectangle
                blnSquareFill = True
                picCapture.FillStyle = intFillStyle
                picCapture.FillColor = lngRColor
                .Square = True
                
            Case 9: .Fan = True     'fan effect
            Case 10: .ColorPicker = True    'colorpicker tool
            Case 11: .Caligraphy = True     'caligraphy tool
            Case 12
                .Filler = True         'filler tool
                picCapture.FillStyle = intFillStyle
                
            Case 15     'eraser
                picCapture.DrawWidth = CurentWidth + 5
                .Eraser = True
                
            Case 20: picCapture.BackColor = lngRColor    'fill background with current selected color
            Case 21: RandomizeBackground Me    'fill background with random color dots
        End Select
    End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    Me.Top = ((frmMain.Height - Me.Height) / 2) + frmMain.Top
    Me.Left = ((frmMain.Width - Me.Width) / 2) + frmMain.Left
    
    With sprGrid
        .RowHeight(-1) = 14
        
        .MaxRows = 0
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeNormal
    End With

    '
    lngLColor = vbRed
    lngRColor = vbRed
    
    picColorL.BackColor = lngLColor
    picColorR.BackColor = lngRColor
    
    CurentWidth = 2
    
    picCapture.DrawWidth = CurentWidth
    picCapture.FillColor = vbRed
    picCapture.FillStyle = vbSolid

    'set UDT tool to clear previous tool
    SetAllFalse mudtTool
        
    'select tool and attributes
    picCapture.FillStyle = intFillStyle
    blnCircleFill = True
    
    With mudtTool
        .Circle = True
                
        '.FreeLine = True  'freehand line
    End With
    
    '--------------------------------------------------
    Call LoadDeviceList
    
    If lstDevices.ListCount > 0 Then
        lstDevices.Selected(0) = True
    Else
        cmdStart.Enabled = False
        lstDevices.AddItem ("No Device Available")
    End If
    
    'Call cmdStart_Click
    
    '-----------------------------------------------------------
    ' TB_오점내용
    '-----------------------------------------------------------
    Query = "SELECT * FROM TB_오점내용"
    Query = Query & " ORDER BY 오점내용 ASC"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With cboStain
        .Clear
        
        Do Until ADORs.EOF
            .AddItem Trim(ADORs!오점내용) & ""
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
    
    Dim iRow         As Integer
    Dim iCol         As Integer
    Dim ColorPallete As String
    
    iRow = 1
    iCol = 0
    
' 4 * 4 = 16
    For i = 15 To 0 Step -1
        ColorPallete = QBColor(i)

        If iCol >= 4 Then
            iRow = iRow + 1
            iCol = 1
        Else
            iCol = iCol + 1
        End If

        sprColor.Row = iRow
        sprColor.Col = iCol: sprColor.BackColor = ColorPallete

        sprColor.CellTag = i & ""
        'sprColor.Text = i & ""
    Next i
    
' 6 * 6 = 48
'    For i = 1 To 48
'        If i = 1 Then
'            ColorPallete = Mid(ColorValues, 1, 8)
'        Else
'            ColorPallete = Mid(ColorValues, (i - 1) * 8 + 1, 8)
'        End If
'
'        If iCol > 6 Then
'            iRow = iRow + 1
'            iCol = 1
'        Else
'            iCol = iCol + 1
'        End If
'
'        sprColor.Row = iRow
'        sprColor.Col = iCol: sprColor.BackColor = ColorPallete
'    Next i
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub ButtonPlus1_Click(Index As Integer)

End Sub

Private Sub cmdStart_Click()
    iDevice = lstDevices.ListIndex
    OpenPreviewWindow
End Sub

Private Sub cmdStop_Click()
    ClosePreviewWindow
    
    picCapture.Cls
    
    cmdStop.Enabled = False
    btnSave.Enabled = False
    cmdStart.Enabled = True
End Sub

Private Sub cmdTool_Click(Index As Integer)
    'set UDT tool to clear previous tool
    SetAllFalse mudtTool
    
    'set the width in case it chaged with the previous tool
    picCapture.DrawWidth = CurentWidth
    
    'select tool and attributes
    With mudtTool
        Select Case Index
            Case 0: .Point = True     'single point
            Case 1: .Line = True      'straight line
            Case 2: .FreeLine = True  'freehand line
            Case 3      'circle
                picCapture.FillStyle = vbFSTransparent
                blnCircleFill = False
                .Circle = True
            Case 4      'filled Circle
                picCapture.FillStyle = intFillStyle
                blnCircleFill = True
                .Circle = True
            Case 5      'elipse
                picCapture.FillStyle = vbFSTransparent
                blnElipseFill = False
                .Elipse = True
            Case 6      'elipse fill
                picCapture.FillStyle = intFillStyle
                blnElipseFill = True
                .Elipse = True
            Case 7      'empty rectangle
                blnSquareFill = False
                picCapture.FillStyle = vbFSTransparent
                .Square = True
            Case 8      'filled rectangle
                blnSquareFill = True
                picCapture.FillStyle = intFillStyle
                picCapture.FillColor = lngRColor
                .Square = True
                
            Case 9: .Fan = True     'fan effect
            Case 10: .ColorPicker = True    'colorpicker tool
            Case 11: .Caligraphy = True     'caligraphy tool
            Case 12
                .Filler = True         'filler tool
                picCapture.FillStyle = intFillStyle
            Case 15     'eraser
                picCapture.DrawWidth = CurentWidth + 5
                .Eraser = True
            Case 20: picCapture.BackColor = lngRColor    'fill background with current selected color
            Case 21: RandomizeBackground Me    'fill background with random color dots
        End Select
    End With
End Sub

Private Sub LoadDeviceList()
    Dim strName As String
    Dim strVer  As String
    Dim iReturn As Boolean
    Dim x       As Long
    
    x = 0
    strName = Space(100)
    strVer = Space(100)

    ' Load name of all available devices into lstDevices
    Do
        iReturn = capGetDriverDescriptionA(x, strName, 100, strVer, 100) ' Get Driver name and version
        
        If iReturn Then
            lstDevices.AddItem Trim(strName)  ' If there was a device add device name to the list
        End If
        
        x = x + 1
    Loop Until iReturn = False
End Sub

Private Sub OpenPreviewWindow()
    ' Open Preview window in picturebox
    hHwnd = capCreateCaptureWindowA(iDevice, WS_VISIBLE Or WS_CHILD, 0, 0, 640, 480, picCapture.hWnd, 0)

    ' Connect to device
    If SendMessage(hHwnd, WM_CAP_DRIVER_CONNECT, iDevice, 0) Then
        SendMessage hHwnd, WM_CAP_SET_SCALE, True, 0 'Set the preview scale
        SendMessage hHwnd, WM_CAP_SET_PREVIEWRATE, 66, 0 'Set the preview rate in milliseconds
        SendMessage hHwnd, WM_CAP_SET_PREVIEW, True, 0 'Start previewing the image from the camera

        ' Resize window to fit in picturebox
        'SetWindowPos hHwnd, HWND_BOTTOM, 0, 0, picCapture.ScaleWidth, picCapture.ScaleHeight, SWP_NOMOVE Or SWP_NOZORDER
        btnSave.Enabled = True
        cmdStop.Enabled = True
        cmdStart.Enabled = False
    Else
        DestroyWindow hHwnd ' Error connecting to device close window

        btnSave.Enabled = False
    End If
 End Sub

Private Sub ClosePreviewWindow()
    SendMessage hHwnd, WM_CAP_DRIVER_DISCONNECT, iDevice, 0 ' Disconnect from device
    DestroyWindow hHwnd ' close window
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdStop.Enabled Then
        Call ClosePreviewWindow
    End If
End Sub

Private Sub imgClass_Click(Index As Integer)
    Select Case Index
        Case 0: picCapture.Picture = LoadPicture(AppPath & "\image\Shirt-01.gif")
        Case 1: picCapture.Picture = LoadPicture(AppPath & "\image\Paint-01.gif")
        Case 2: picCapture.Picture = LoadPicture(AppPath & "\image\Skirt-01.gif")
    End Select
End Sub

'picCapture
Private Sub picCapture_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'check to see which tool is selected and act accordingly
   If Button = vbLeftButton Then
        lngOutColor = lngLColor
   ElseIf Button = vbRightButton Then
        lngOutColor = lngRColor
   End If
   
   With mudtPoint
        .msglX = x
        .msglY = y
   End With
   
   'start to draw the objects
   With mudtTool
        If .Line Then
            InitiateLine mudtPoint
            
        ElseIf .Point Then
            blnModified = True
            picCapture.PSet (mudtPoint.msglX, mudtPoint.msglY), lngOutColor
            
        ElseIf .Circle Then
            InitiateCircle mudtPoint, blnCircleFill, Me
            
        ElseIf .Elipse Then
            InitiateElipse mudtPoint, blnElipseFill, Me
            
        ElseIf .Square Then
            InitiateSquare mudtPoint, blnSquareFill, Me
            
        ElseIf .Eraser Then
            picCapture.PSet (x, y), picCapture.BackColor
            
        ElseIf .FreeLine Then
            InitiateFreeLine mudtPoint
            
        ElseIf .Fan Then
            InitiateFan mudtPoint
            
        ElseIf .ColorPicker Then
            GetColorPicker mudtPoint, Button, Me
            
        ElseIf .Caligraphy Then
            InitiateCaligraphy mudtPoint, Me
            
        ElseIf .Filler Then
            picCapture.FillColor = lngOutColor
            DoFiller mudtPoint, picCapture.hdc
        End If
   End With
End Sub

Private Sub picCapture_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Or Button = vbRightButton Then
        With mudtPoint
            .msglX = x
            .msglY = y
        End With
       
        'use live drawing features here
        With mudtTool
            If .FreeLine Then
                DrawFreeLine mudtPoint, Me
                
            ElseIf .Fan Then
                DrawFan mudtPoint, Me
                
            ElseIf .Eraser Then
                picCapture.PSet (x, y), picCapture.BackColor
                
            ElseIf .Line Then
                DrawLine mudtPoint, Me
                
            ElseIf .Circle Then
                DrawCircle mudtPoint, Me
                
            ElseIf .Elipse Then
                DrawElipse mudtPoint, Me
                
            ElseIf .Square Then
                DrawSquare mudtPoint, Me
                
            ElseIf .Caligraphy Then
                DrawCaligraphy mudtPoint, Me
            End If
        End With
    End If
End Sub

Private Sub picCapture_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    picCapture.DrawMode = 13
    
   With mudtPoint
        .msglX = x
        .msglY = y
   End With
   
    'draw the final objects
    With mudtTool
        If .Line Then
            FinalizeLine mudtPoint, Me
        ElseIf .Circle Then
            FinalizeCircle mudtPoint, Me
        ElseIf .Elipse Then
            FinalizeElipse mudtPoint, Me
        ElseIf .Square Then
            FinalizeSquare mudtPoint, Me
        End If
    End With
End Sub

Private Sub sprColor_Click(ByVal Col As Long, ByVal Row As Long)
    If Row <= 0 Then Exit Sub
    
    mudtTool.ColorPicker = True
    
    With sprColor
        .Row = Row
        .Col = Col
        
        lngLColor = .BackColor
        lngRColor = .BackColor
        iQBColor = .CellTag & ""
    End With
    
    picColorL.BackColor = lngLColor
    picColorR.BackColor = lngRColor
    
    'set UDT tool to clear previous tool
    SetAllFalse mudtTool
    
    'set the width in case it chaged with the previous tool
    picCapture.DrawWidth = CurentWidth
    
    'select tool and attributes
    mudtTool.FreeLine = True
End Sub

Private Sub sprGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    If Row <= 0 Then Exit Sub
    
    If Col = 2 Then
        Call sprGrid.DeleteRows(Row, 1)
                    
        sprGrid.MaxRows = sprGrid.MaxRows - 1
    End If
End Sub
