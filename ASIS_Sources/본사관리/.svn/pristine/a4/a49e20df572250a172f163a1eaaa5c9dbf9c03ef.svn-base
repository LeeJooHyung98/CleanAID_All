VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_02008 
   Caption         =   "지사 입고검품 현황"
   ClientHeight    =   10185
   ClientLeft      =   3000
   ClientTop       =   3345
   ClientWidth     =   16380
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_02008.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10185
   ScaleWidth      =   16380
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   10185
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16380
      _ExtentX        =   28893
      _ExtentY        =   17965
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_02008.frx":058A
      Begin Threed.SSPanel SSPanel 
         Height          =   570
         Index           =   1
         Left            =   5310
         TabIndex        =   25
         Top             =   1740
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   1005
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   390
            Index           =   2
            Left            =   1770
            TabIndex        =   26
            Top             =   105
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   688
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   16777215
            Format          =   65077248
            CurrentDate     =   36686
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "지사입고 처리일자:"
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
            Left            =   30
            TabIndex        =   29
            Top             =   210
            Width           =   1680
         End
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   1
         Top             =   540
         Width           =   16350
         _ExtentX        =   28840
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1245
            Locked          =   -1  'True
            Style           =   2  '드롭다운 목록
            TabIndex        =   12
            Top             =   60
            Width           =   2850
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1245
            TabIndex        =   13
            Top             =   405
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            Format          =   65077248
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   14
            Top             =   405
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "입고일자"
            BorderWidth     =   0
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   7
            Left            =   60
            TabIndex        =   15
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "지    사"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   4425
            TabIndex        =   16
            Top             =   405
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            Format          =   65077248
            CurrentDate     =   36686
         End
         Begin VB.Label Label 
            Alignment       =   2  '가운데 맞춤
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
            Height          =   255
            Index           =   0
            Left            =   4110
            TabIndex        =   17
            Top             =   465
            Width           =   300
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   8745
         _ExtentX        =   15425
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
         Caption         =   " 지사 입고검품 현황 (P_02008)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_02008.frx":069C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   0
         Left            =   8775
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
         PictureBackground=   "P_02008.frx":089E
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
            Picture         =   "P_02008.frx":0AA0
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
            Picture         =   "P_02008.frx":103A
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
            Picture         =   "P_02008.frx":15D4
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
            Picture         =   "P_02008.frx":1B6E
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
            Picture         =   "P_02008.frx":2108
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
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_02008.frx":26A2
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
            Picture         =   "P_02008.frx":2C3C
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
            Picture         =   "P_02008.frx":31D6
         End
      End
      Begin XtremeSuiteControls.TabControl TabControl 
         Height          =   7845
         Left            =   5310
         TabIndex        =   18
         Top             =   2325
         Width           =   11055
         _Version        =   851970
         _ExtentX        =   19500
         _ExtentY        =   13838
         _StockProps     =   68
         Appearance      =   3
         Color           =   64
         PaintManager.BoldSelected=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         PaintManager.ButtonMargin=   "2,3,2,3"
         ItemCount       =   2
         Item(0).Caption =   " 가맹점 - 접수현황 "
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "TabControlPage(0)"
         Item(1).Caption =   " 지사 - PDA 스캔 입고현황 "
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "TabControlPage(1)"
         Begin XtremeSuiteControls.TabControlPage TabControlPage 
            Height          =   7365
            Index           =   1
            Left            =   -69970
            TabIndex        =   19
            Top             =   450
            Visible         =   0   'False
            Width           =   10995
            _Version        =   851970
            _ExtentX        =   19394
            _ExtentY        =   12991
            _StockProps     =   1
            Page            =   1
            Begin FPSpreadADO.fpSpread spdView1 
               Height          =   8235
               Index           =   1
               Left            =   60
               TabIndex        =   20
               Top             =   570
               Width           =   10875
               _Version        =   524288
               _ExtentX        =   19182
               _ExtentY        =   14526
               _StockProps     =   64
               BackColorStyle  =   1
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
               MaxCols         =   5
               MaxRows         =   35
               ScrollBars      =   2
               SpreadDesigner  =   "P_02008.frx":3770
               UserResize      =   1
               Appearance      =   1
               HighlightHeaders=   1
               HighlightStyle  =   1
            End
            Begin XtremeSuiteControls.PushButton cmdBtn 
               Height          =   450
               Index           =   9
               Left            =   75
               TabIndex        =   28
               Top             =   75
               Width           =   3105
               _Version        =   851970
               _ExtentX        =   5477
               _ExtentY        =   794
               _StockProps     =   79
               Caption         =   " PDA 스캔 - 지사입고 확정"
               ForeColor       =   -2147483640
               BackColor       =   -2147483636
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
               Picture         =   "P_02008.frx":3D96
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage 
            Height          =   7365
            Index           =   0
            Left            =   30
            TabIndex        =   21
            Top             =   450
            Width           =   10995
            _Version        =   851970
            _ExtentX        =   19394
            _ExtentY        =   12991
            _StockProps     =   1
            Page            =   0
            Begin FPSpreadADO.fpSpread spdView1 
               Height          =   8235
               Index           =   0
               Left            =   60
               TabIndex        =   22
               Top             =   570
               Width           =   10875
               _Version        =   524288
               _ExtentX        =   19182
               _ExtentY        =   14526
               _StockProps     =   64
               BackColorStyle  =   1
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
               MaxCols         =   5
               MaxRows         =   35
               ScrollBars      =   2
               SpreadDesigner  =   "P_02008.frx":4490
               UserResize      =   1
               Appearance      =   1
               HighlightHeaders=   1
               HighlightStyle  =   1
            End
            Begin XtremeSuiteControls.PushButton cmdBtn 
               Height          =   450
               Index           =   8
               Left            =   75
               TabIndex        =   27
               Top             =   75
               Width           =   1995
               _Version        =   851970
               _ExtentX        =   3519
               _ExtentY        =   794
               _StockProps     =   79
               Caption         =   " 지사입고 확정"
               ForeColor       =   -2147483640
               BackColor       =   -2147483636
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
               Picture         =   "P_02008.frx":4AE6
            End
         End
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   390
         Index           =   2
         Left            =   5310
         TabIndex        =   23
         Top             =   1335
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   688
         _Version        =   262144
         Font3D          =   3
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " 지사입고 현황"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_02008.frx":5080
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.ProgressBar ProgressBar 
            Height          =   270
            Left            =   5910
            TabIndex        =   24
            Top             =   45
            Visible         =   0   'False
            Width           =   3270
            _Version        =   851970
            _ExtentX        =   5768
            _ExtentY        =   476
            _StockProps     =   93
            Scrolling       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   8385
         Left            =   15
         TabIndex        =   30
         Top             =   1335
         Width           =   5280
         _Version        =   524288
         _ExtentX        =   9313
         _ExtentY        =   14790
         _StockProps     =   64
         BackColorStyle  =   1
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
         MaxCols         =   4
         MaxRows         =   35
         ScrollBars      =   2
         SpreadDesigner  =   "P_02008.frx":54E2
         UserResize      =   1
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   435
         Index           =   3
         Left            =   15
         TabIndex        =   31
         Top             =   9735
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   767
         _Version        =   262144
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   345
            Index           =   0
            Left            =   3165
            TabIndex        =   32
            Top             =   45
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
            _ExtentY        =   609
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   192
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.76
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            ReadOnly        =   -1  'True
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   2
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   16
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "접수수량:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   2040
            TabIndex        =   34
            Top             =   135
            Width           =   1080
         End
         Begin VB.Label Label 
            BackStyle       =   0  '투명
            Caption         =   "점"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   4125
            TabIndex        =   33
            Top             =   135
            Width           =   255
         End
      End
   End
End
Attribute VB_Name = "P_02008"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub SPR_Resize()
    On Error GoTo ErrRtn
    
    spdView1(0).Width = Me.Width - 5610
    spdView1(0).Height = Me.Height - 3900

    spdView1(1).Width = Me.Width - 5610
    spdView1(1).Height = Me.Height - 3900

    Exit Sub
    
ErrRtn:

End Sub

Private Sub cboOffice_Click()
    If cboOffice.ListIndex < 0 Then Exit Sub
    
    Call Data_Display
End Sub

'-----------------------------------------------------------------
'
'-----------------------------------------------------------------
Private Sub Data_Display()
    On Error GoTo ErrRtn
    
    txtNum(0).Value = 0
    
    ReDim sValue(2)
    
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_02008_00", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_02008_00", sValue(), Err_Num, Err_Dec)
    End If
        
    With spdView
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = RS01!가맹점코드 & ""
            .Col = 2: .Text = RS01!가맹점명 & ""
            .Col = 3: .Text = RS01!접수수량 & ""
            .Col = 4: .Text = RS01!입고수량 & ""
            
            txtNum(0).Value = txtNum(0).Value + RS01!접수수량
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
    
        .Redraw = True
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: Call Data_Display    ' 조회
        Case 1: 'Call DataAdd        ' 신규
        Case 2: 'Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
        Case 4: 'Call DataCancel     ' 취소
        Case 5: 'Call DataPrint      ' 인쇄
        Case 6: Call Export_Excel(P_00000.cdgExcel, spdView)      ' 엑셀
        Case 7: Unload Me            ' 종료
        Case 8: Call Data_Update(0)  '
        Case 9: Call Data_Update(1)  '
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

Private Sub Data_Update(Idx As Integer)
    Dim i      As Long
    Dim j      As Long
    
    Dim 가맹점코드 As String
    Dim 택번호     As String
    Dim Query      As String
    
    If spdView.ActiveRow <= 0 Then Exit Sub
    
    spdView.Row = spdView.ActiveRow
    spdView.Col = 1: 가맹점코드 = spdView.Text & ""
    
    With spdView1(Idx)
        For j = 1 To .MaxRows
            .Row = j
            .Col = 1: 택번호 = Replace(.Text, "-", "") & ""
            
            '------------------------------------------------------------------------------------------
            ' TB_입출고
            '------------------------------------------------------------------------------------------
            Query = "UPDATE TB_입출고 SET 지사입고일자 = '" & Format(dtInput(2).Value, "YYYY-MM-DD") & " " & Format(Time, "hh:mm:ss") & "'"
            Query = Query & " WHERE 가맹점코드 = '" & 가맹점코드 & "'"
            Query = Query & "   AND 택번호     = '" & 택번호 & "'"
            Query = Query & "   AND 접수일자 BETWEEN '" & Format(dtInput(0).Value, "YYYY-MM-DD") & "'  AND '" & Format(dtInput(1).Value, "YYYY-MM-DD") & "'"
            ADOCon.Execute Query
        Next j
                
                
        If Idx = 1 Then
            '-------------------------------------------------------------------------------------
            ' 2) SCANINPUT_LOG_TB에 저장하기
            '-------------------------------------------------------------------------------------
            ReDim sValue(3)
            
            For i = 1 To .MaxRows
                'ProgressBar.Value = (i / spdView1(idx)) * 100
                DoEvents
                
                .Row = i
                .Col = 4:  sValue(0) = Trim(.Text) & ""             '1
                .Col = 5:  sValue(1) = Trim(.Text) & ""             '2
                .Col = 1:  sValue(2) = Replace(.Text, "-", "") & "" '3
                           sValue(3) = 가맹점코드                   '4
                                
                Call ExecPro("SP_02008_03", sValue(), Err_Num, Err_Dec)
                
            Next i
            
'            SP_02008_03 에서 삭제
'            '-------------------------------------------------------------------------------------
'            ' 3) SCANOUTPUT_TB 삭제하기
'            '-------------------------------------------------------------------------------------
'            Query = "DELETE FROM SCANINPUT_TB"
'            Query = Query & " WHERE STORE_CD = '" & 가맹점코드 & "'"
'            Query = Query & "   AND SUBSTRING(SCAN_DATE,1,10) >= '" & Format(dtInput(0).Value, "YYYY-MM-DD") & "'"
'            Query = Query & "   AND SUBSTRING(SCAN_DATE,1,10) <= '" & Format(dtInput(1).Value, "YYYY-MM-DD") & "'"
'            ADOCon.Execute Query
            
            Call Data_Display2(가맹점코드)
        End If
    End With
    
    ' 좌측 내용 다시 조회
    Call cmdBtn_Click(0)
End Sub

Private Sub dtInput_Change(Index As Integer)
'    dtInput(Index).Enabled = False
'
'    Call Data_Display
'
'    dtInput(Index).Enabled = True
End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(5).Enabled = True
    
    '본사 여부
    If HeadOffice = MASTER_OFFICE_CODE Then
        cboOffice.Locked = False
        
        cmdBtn(8).Enabled = False '본사에서는 지사입고 확정은 못함
        cmdBtn(9).Enabled = False '본사에서는 지사입고 확정은 못함
    Else
        cboOffice.Locked = True
    End If
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
    If P_02008_Flag = False Then
        dtInput(0).Value = Date
        dtInput(1).Value = Date
        
        P_02008_Flag = True
    End If
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
    
    Dim i As Integer
    
    For i = 0 To 1
        With spdView1(i)
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
    Next i
    
    Call SPR_Resize
    
    dtInput(0).Value = Date
    dtInput(1).Value = Date
    dtInput(2).Value = Date

    '
    Call Get_지사리스트(cboOffice)
    
    With cboOffice
        For i = 0 To .ListCount - 1
            If Mid(.List(i), 2, 4) = HeadOffice Then
                .ListIndex = i
                
                Exit For
            End If
        Next i
    End With
        
'    If P_02008_Flag = False Then
'        dtInput(0).Value = Date
'        dtInput(1).Value = Date
'
'        '
'        Call Get_지사리스트(cboOffice)
'
'        With cboOffice
'            For i = 0 To .ListCount - 1
'                If Mid(.List(i), 2, 4) = HeadOffice Then
'                    .ListIndex = i
'
'                    Exit For
'                End If
'            Next i
'        End With
'
'        P_02008_Flag = True
'    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Resize()
    Call SPR_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_02008_Flag = True
End Sub

Private Sub Data_Display2(가맹점코드 As String)
    On Error GoTo ErrRtn
    
    ReDim sValue(2)
    
    Screen.MousePointer = vbHourglass
    
    '------------------------------------------------------------
    ' 지사 입고처리 - SP_02008_01
    '------------------------------------------------------------
    sValue(0) = 가맹점코드 'Mid(cboInput.Text, 2, 6)
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
        
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_02008_01", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_02008_01", sValue(), Err_Num, Err_Dec)
    End If
            
    With spdView1(0)
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = Left(RS01!택번호, 3) & "-" & Mid(RS01!택번호, 4, 2) & "-" & Mid(RS01!택번호, 6, 4) '
            .Col = 2: .Text = RS01!의류코드 & ""                                  '
            .Col = 3: .Text = RS01!의류명 & ""                                    '
            .Col = 4: .Text = RS01!접수일자 & ""                                  '
            .Col = 5: .Text = RS01!SCAN_DATE & ""                                 '
        
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
    
    '------------------------------------------------------------
    ' 지사 입고처리 - SP_02008_02
    '------------------------------------------------------------
    sValue(0) = 가맹점코드 'Mid(cboInput.Text, 2, 6)
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
        
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_02008_02", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_02008_02", sValue(), Err_Num, Err_Dec)
    End If
            
    With spdView1(1)
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = Left(RS01!TAG_NO, 3) & "-" & Mid(RS01!TAG_NO, 4, 2) & "-" & Mid(RS01!TAG_NO, 6, 4) '
            .Col = 2: .Text = RS01!의류코드 & ""  '
            .Col = 3: .Text = RS01!의류명 & ""    '
            .Col = 4: .Text = RS01!SCAN_DATE & "" '
            .Col = 5: .Text = RS01!PDA_NO & ""    '
        
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
        
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrRtn:
    Screen.MousePointer = vbDefault

End Sub

Private Sub spdView_Click(ByVal Col As Long, ByVal Row As Long)
    Dim 가맹점코드 As String
    
    If Row <= 0 Then Exit Sub
    
    spdView.Row = Row
    spdView.Col = 1: 가맹점코드 = spdView.Text & ""
    
    Call Data_Display2(가맹점코드)
End Sub

