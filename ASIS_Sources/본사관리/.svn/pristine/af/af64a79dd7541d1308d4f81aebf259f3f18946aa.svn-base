VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_01003_A 
   Caption         =   "(본사) 대표 품목 등록"
   ClientHeight    =   11715
   ClientLeft      =   3615
   ClientTop       =   1980
   ClientWidth     =   15930
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_01003_A.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11715
   ScaleWidth      =   15930
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   11715
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15930
      _ExtentX        =   28099
      _ExtentY        =   20664
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_01003_A.frx":058A
      Begin XtremeSuiteControls.TabControl TabControl 
         Height          =   9960
         Left            =   2970
         TabIndex        =   28
         Top             =   1740
         Width           =   5685
         _Version        =   851970
         _ExtentX        =   10028
         _ExtentY        =   17568
         _StockProps     =   68
         Appearance      =   3
         Color           =   64
         PaintManager.Position=   2
         PaintManager.BoldSelected=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         ItemCount       =   2
         Item(0).Caption =   " 품목 "
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "TabControlPage(0)"
         Item(1).Caption =   " 가맹점 "
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "TabControlPage(1)"
         Begin XtremeSuiteControls.TabControlPage TabControlPage 
            Height          =   9570
            Index           =   1
            Left            =   -69970
            TabIndex        =   30
            Top             =   30
            Visible         =   0   'False
            Width           =   5625
            _Version        =   851970
            _ExtentX        =   9922
            _ExtentY        =   16880
            _StockProps     =   1
            Page            =   1
            Begin FPSpreadADO.fpSpread sprBranch 
               Height          =   9165
               Left            =   45
               TabIndex        =   32
               Top             =   45
               Width           =   5550
               _Version        =   524288
               _ExtentX        =   9790
               _ExtentY        =   16166
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
               MaxCols         =   2
               ScrollBars      =   2
               SpreadDesigner  =   "P_01003_A.frx":06BC
               Appearance      =   1
               HighlightHeaders=   1
               HighlightStyle  =   1
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage 
            Height          =   9570
            Index           =   0
            Left            =   30
            TabIndex        =   29
            Top             =   30
            Width           =   5625
            _Version        =   851970
            _ExtentX        =   9922
            _ExtentY        =   16880
            _StockProps     =   1
            Page            =   0
            Begin FPSpreadADO.fpSpread spdView 
               Height          =   9165
               Left            =   45
               TabIndex        =   31
               Top             =   45
               Width           =   5550
               _Version        =   524288
               _ExtentX        =   9790
               _ExtentY        =   16166
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
               MaxCols         =   5
               ScrollBars      =   2
               SpreadDesigner  =   "P_01003_A.frx":0BE0
               Appearance      =   1
               HighlightHeaders=   1
               HighlightStyle  =   1
            End
         End
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   780
         Index           =   2
         Left            =   15
         TabIndex        =   12
         Top             =   540
         Width           =   15900
         _ExtentX        =   28046
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSFrame SSFrame1 
            Height          =   735
            Left            =   6540
            TabIndex        =   37
            Top             =   30
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   1296
            _Version        =   262144
            Begin XtremeSuiteControls.CheckBox chkSaveAction 
               Height          =   225
               Left            =   240
               TabIndex        =   40
               Top             =   60
               Width           =   3705
               _Version        =   851970
               _ExtentX        =   6535
               _ExtentY        =   397
               _StockProps     =   79
               Caption         =   " 저장시 가맹점 가격 자료 자동 "
               ForeColor       =   255
               UseVisualStyle  =   -1  'True
            End
            Begin MSComCtl2.DTPicker dtInput 
               Height          =   315
               Left            =   1185
               TabIndex        =   38
               Top             =   330
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   556
               _Version        =   393216
               Format          =   63438848
               CurrentDate     =   36686
            End
            Begin VB.Label lblTitle 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "적용일자:"
               Height          =   225
               Index           =   24
               Left            =   90
               TabIndex        =   39
               Top             =   390
               Width           =   1065
            End
         End
         Begin XtremeSuiteControls.ProgressBar ProgressBar 
            Height          =   330
            Left            =   1170
            TabIndex        =   33
            Top             =   60
            Width           =   5325
            _Version        =   851970
            _ExtentX        =   9393
            _ExtentY        =   582
            _StockProps     =   93
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Scrolling       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ProgressBar ProgressBar2 
            Height          =   330
            Left            =   1170
            TabIndex        =   34
            Top             =   420
            Visible         =   0   'False
            Width           =   5325
            _Version        =   851970
            _ExtentX        =   9393
            _ExtentY        =   582
            _StockProps     =   93
            ForeColor       =   16777215
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
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label lblProgress 
            BackStyle       =   0  '투명
            Caption         =   "모든 작업은 당일 처리하셔야 합니다."
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   165
            Index           =   6
            Left            =   10980
            TabIndex        =   47
            Top             =   570
            Width           =   5700
         End
         Begin VB.Label lblProgress 
            BackStyle       =   0  '투명
            Caption         =   "최종 작업자가 저장시 가맹점 가격자료 생성을 하여 주시면 됩니다."
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   165
            Index           =   5
            Left            =   10980
            TabIndex        =   46
            Top             =   330
            Width           =   5700
         End
         Begin VB.Label lblProgress 
            BackStyle       =   0  '투명
            Caption         =   "PC 일자가 같은 컴퓨터에서 분할 작업이 가능합니다."
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   165
            Index           =   4
            Left            =   10980
            TabIndex        =   45
            Top             =   90
            Width           =   5220
         End
         Begin VB.Label lblProgress 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "가맹점 저장:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   0
            TabIndex        =   36
            Top             =   150
            Width           =   1140
         End
         Begin VB.Label lblProgress 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "품목 저장:"
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
            Left            =   0
            TabIndex        =   35
            Top             =   510
            Visible         =   0   'False
            Width           =   1140
         End
      End
      Begin FPSpreadADO.fpSpread sprList 
         Height          =   7815
         Left            =   8670
         TabIndex        =   1
         Top             =   3885
         Width           =   7245
         _Version        =   524288
         _ExtentX        =   12779
         _ExtentY        =   13785
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
         MaxCols         =   8
         SpreadDesigner  =   "P_01003_A.frx":122B
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   8295
         _ExtentX        =   14631
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
         Caption         =   " (본사) 대표 품목 등록 (P_01003_A)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_01003_A.frx":1945
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   0
         Left            =   8325
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
         PictureBackground=   "P_01003_A.frx":1B47
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
            Picture         =   "P_01003_A.frx":1D49
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
            Picture         =   "P_01003_A.frx":22E3
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
            Picture         =   "P_01003_A.frx":287D
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
            Picture         =   "P_01003_A.frx":2E17
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
            Picture         =   "P_01003_A.frx":33B1
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
            Picture         =   "P_01003_A.frx":394B
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
            Picture         =   "P_01003_A.frx":3EE5
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
            Picture         =   "P_01003_A.frx":447F
         End
      End
      Begin FPSpreadADO.fpSpread sprClass 
         Height          =   9960
         Left            =   15
         TabIndex        =   13
         Top             =   1740
         Width           =   2940
         _Version        =   524288
         _ExtentX        =   5186
         _ExtentY        =   17568
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
         MaxCols         =   2
         ScrollBars      =   2
         SpreadDesigner  =   "P_01003_A.frx":4A19
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panCaption 
         Height          =   390
         Index           =   2
         Left            =   15
         TabIndex        =   14
         Top             =   1335
         Width           =   8640
         _ExtentX        =   15240
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
         Caption         =   " 대표 품목 현황"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_01003_A.frx":4F35
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   2130
         Index           =   1
         Left            =   8670
         TabIndex        =   15
         Top             =   1740
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   3757
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtData 
            Height          =   315
            Index           =   1
            Left            =   1245
            TabIndex        =   18
            Top             =   780
            Width           =   3615
         End
         Begin VB.ComboBox cboGubun 
            Height          =   315
            Left            =   1245
            Style           =   2  '드롭다운 목록
            TabIndex        =   17
            Top             =   60
            Width           =   1575
         End
         Begin VB.TextBox txtData 
            Height          =   315
            Index           =   0
            Left            =   1245
            TabIndex        =   16
            Top             =   420
            Width           =   1575
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   36
            Left            =   60
            TabIndex        =   19
            Top             =   420
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "코    드"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   38
            Left            =   60
            TabIndex        =   20
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "구    분"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   60
            TabIndex        =   21
            Top             =   780
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "품 목 명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin CSTextLibCtl.sidbEdit txtMoney 
            Height          =   315
            Left            =   1245
            TabIndex        =   22
            Top             =   1140
            Width           =   1140
            _Version        =   262145
            _ExtentX        =   2011
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Insert          =   0   'False
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   2
            StartText.y     =   2
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
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   3
            Left            =   60
            TabIndex        =   23
            Top             =   1140
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "금    액"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnDEL 
            Height          =   450
            Left            =   2250
            TabIndex        =   24
            Top             =   1605
            Width           =   1050
            _Version        =   851970
            _ExtentX        =   1852
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 삭제"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_01003_A.frx":5397
         End
         Begin XtremeSuiteControls.PushButton btnEDIT 
            Height          =   450
            Left            =   1155
            TabIndex        =   25
            Top             =   1605
            Width           =   1050
            _Version        =   851970
            _ExtentX        =   1852
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 수정"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_01003_A.frx":5931
         End
         Begin XtremeSuiteControls.PushButton btnADD 
            Height          =   450
            Left            =   60
            TabIndex        =   26
            Top             =   1605
            Width           =   1050
            _Version        =   851970
            _ExtentX        =   1852
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 추가"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_01003_A.frx":5ECB
         End
         Begin CSTextLibCtl.sidbEdit txtOldMoney 
            Height          =   315
            Left            =   3720
            TabIndex        =   41
            Top             =   1140
            Width           =   1140
            _Version        =   262145
            _ExtentX        =   2011
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Insert          =   0   'False
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   2
            StartText.y     =   2
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
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   4
            Left            =   2550
            TabIndex        =   42
            Top             =   1140
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "이전금액"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin VB.Label lblProgress 
            BackStyle       =   0  '투명
            Caption         =   "모든 가맹점의 가격이 일괄 변경 됩니다."
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   3
            Left            =   2940
            TabIndex        =   44
            Top             =   360
            Width           =   5220
         End
         Begin VB.Label lblProgress 
            BackStyle       =   0  '투명
            Caption         =   "금액이 변경된 경우 가맹점 가격자료 자동 생성시"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   2
            Left            =   2940
            TabIndex        =   43
            Top             =   120
            Width           =   5220
         End
      End
      Begin Threed.SSPanel panCaption 
         Height          =   390
         Index           =   0
         Left            =   8670
         TabIndex        =   27
         Top             =   1335
         Width           =   7245
         _ExtentX        =   12779
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
         Caption         =   " 품목 추가 및 변경"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_01003_A.frx":6465
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
      End
   End
End
Attribute VB_Name = "P_01003_A"
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
    
    spdView.Height = Me.Height - 2680
    sprBranch.Height = Me.Height - 2680

    Exit Sub
    
ErrRtn:

End Sub

Private Sub btnADD_Click()
    Call 품목_Rtn("추가")
End Sub

Private Sub btnDEL_Click()
    Call 품목_Rtn("삭제")
End Sub

Private Sub btnEDIT_Click()
    Call 품목_Rtn("수정")
End Sub

Private Sub 품목_Rtn(작업구분 As String)
    If txtData(0).Text = "" Then
        
        Exit Sub
    End If
    
    If txtData(1).Text = "" Then
        
        Exit Sub
    End If
    
    With sprList
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        
        .Col = 1: .Text = cboGubun.Text & ""
        .Col = 2: .Text = 작업구분
        
        Select Case 작업구분
            Case "추가": .ForeColor = vbBlack
            Case "수정": .ForeColor = vbBlue
            Case "삭제": .ForeColor = vbRed
        End Select
        
        .Col = 3: .Text = LCase(txtData(0).Text) & "" '모두 소문자로 변환...
        .Col = 4: .Text = txtData(1).Text & ""        '
        .Col = 5: .Text = txtMoney.Value & ""         '
        .Col = 8: .Text = txtOldMoney.Value & ""         '
    End With
    
    txtData(0).Text = ""
    txtData(0).Text = ""
    txtMoney.Value = 0
    txtOldMoney.Value = 0
End Sub

 

Private Sub chkSaveAction_Click()
    dtInput.Enabled = chkSaveAction.Value
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: 'Call Data_Display   ' 조회
        Case 1: Call DataAdd        ' 신규
        
        Case 2:
            Call DataSave       ' 저장
        
        Case 3: Call DataDelete     ' 삭제
        Case 4: Call DataCancel     ' 취소
        Case 5: 'Call DataPrint      ' 인쇄
        Case 6: Call Export_Excel(P_00000.cdgExcel, spdView)      ' 엑셀
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
    cmdBtn(0).Enabled = True        '조회
    cmdBtn(1).Enabled = True        '신규
    cmdBtn(2).Enabled = True        '저장
    cmdBtn(3).Enabled = False       '삭제
    cmdBtn(4).Enabled = False       '취소
    cmdBtn(5).Enabled = False       '인쇄
    cmdBtn(6).Enabled = False       'Screen
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
End Sub

'Private Sub spdDisplay(RS As Recordset)
'    Call fpSpread_Display(spdView, RS)
'End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
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
        .OperationMode = OperationModeNormal
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With

    With sprBranch
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

    With cboGubun
        .Clear
        .AddItem "품목분류"
        .AddItem "품목"
        
        .ListIndex = -1
    End With
    
    chkSaveAction.Value = xtpChecked
    
    Call SPR_Resize
    
    If P_01003_A_Flag = False Then
        dtInput.Value = Date
        
        Call 가맹점_Display
        
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
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Resize()
    Call SPR_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_01003_A_Flag = False
End Sub

Private Sub DataAdd()
    spdView.MaxRows = spdView.MaxRows + 1
    
    spdView.Row = spdView.MaxRows
    spdView.Col = 1
    spdView.Action = ActionActiveCell
    spdView.Lock = False
    
    spdView.SetFocus
End Sub

'Private Sub DataSave_20120307_OLD()
'    On Error GoTo ErrRtn
'
'    Dim i         As Long
'    Dim Cloth_Cnt As Long
'
'    Dim Query     As String
'
'    cmdBtn(2).Enabled = False
'    DoEvents
'
'    For i = 1 To sprList.MaxRows
'        sprList.Row = i
'        sprList.Col = 1
'
'        If sprList.Text = "품목" Then
'            sprList.Col = 2
'
'            Select Case sprList.Text
'                Case "추가", "수정"
'                    '------------------------------------------------------
'                    ' TB_의류 - SP_01003_01
'                    '------------------------------------------------------
'                    ReDim sValue(2)
'
'                    sprList.Col = 3: sValue(0) = sprList.Text  '코드
'                    sprList.Col = 4: sValue(1) = sprList.Text  '품명
'                    sprList.Col = 5: sValue(2) = sprList.Value '금액
'
'                    Call ExecPro("SP_01003_01", sValue(), Err_Num, Err_Dec)
'
'                Case "삭제"
'                    '------------------------------------------------------
'                    ' TB_의류 - SP_01003_01
'                    '------------------------------------------------------
'                    ReDim sValue(0)
'
'                    sprList.Row = sprList.ActiveRow
'                    sprList.Col = 3: sValue(0) = sprList.Text '코드
'
'                    Call ExecPro("SP_01003_02", sValue(), Err_Num, Err_Dec)
'            End Select
'
'        Else
'            sprList.Col = 2
'
'            Select Case sprList.Text
'                Case "추가", "수정"
'                    '------------------------------------------------------
'                    ' TB_의류분류 - SP_01003_A_00
'                    '------------------------------------------------------
'                    ReDim sValue(4)
'
'                    sprList.Col = 3: sValue(0) = sprList.Text        '코드
'                    sprList.Col = 4: sValue(1) = Trim(sprList.Text)  '품명
'                                     sValue(2) = 0                   '
'                                     sValue(3) = 0                   '
'                                     sValue(4) = 0                   '
'
'                    Call ExecPro("SP_01003_A_00", sValue(), Err_Num, Err_Dec)
'
'                Case "삭제"
'                    '------------------------------------------------------
'                    ' TB_의류분류 - SP_01003_A_01
'                    '------------------------------------------------------
'                    ReDim sValue(0)
'
'                    sprList.Row = sprList.ActiveRow
'                    sprList.Col = 3: sValue(0) = sprList.Text '코드
'
'                    Call ExecPro("SP_01003_A_01", sValue(), Err_Num, Err_Dec)
'            End Select
'        End If
'    Next i
'
'    '==========================================================================
'    Dim 가맹점코드 As String
'    Dim 의류코드   As String
'
'    ProgressBar.Value = 0
'    ProgressBar.Min = 0
'    ProgressBar.Max = 100
'
'    ProgressBar2.Value = 0
'    ProgressBar2.Min = 0
'    ProgressBar2.Max = 100
'
'    For i = 1 To sprBranch.MaxRows
'        sprBranch.Row = i
'        sprBranch.Col = 1: 가맹점코드 = sprBranch.Text & ""
'        sprBranch.Col = 2: ProgressBar.Text = sprBranch.Text & ""
'
'        ProgressBar.Value = (i / sprBranch.MaxRows) * 100
'        DoEvents
'
'        Call Data_Display(가맹점코드)   '가맹점 의류 보여주기
'
'        Query = "DELETE FROM TB_가맹점의류"
'        Query = Query & " WHERE 가맹점코드 = '" & 가맹점코드 & "'"
'        Query = Query & "   AND 적용일자   = '" & Format(dtInput.Value, "YYYY-MM-DD") & "'"
'        ADOCon.Execute Query
'        DoEvents
'
'        With spdView
'            For Cloth_Cnt = 1 To .MaxRows
'                .Row = Cloth_Cnt
'
'                ProgressBar2.Value = (Cloth_Cnt / .MaxRows) * 100
'
'                .Col = 2: ProgressBar2.Text = .Text
'                DoEvents
'
'                '-------------------------------------------------------------------
'                ' TB_가맹점의류 저장 - SP_01011_02
'                '-------------------------------------------------------------------
'                ReDim sValue(5)
'
'                          sValue(0) = 가맹점코드 & ""                     '1 가맹점코드
'                          sValue(1) = Format(dtInput.Value, "YYYY-MM-DD") '2 적용일자
'                .Col = 1: sValue(2) = Trim(.Text) & ""                    '3 의류코드
'                .Col = 2: sValue(3) = Trim(.Text) & ""                    '4 의류명
'                .Col = 3: sValue(4) = .Value & ""                         '5 금액
'                .Col = 4: sValue(5) = Trim(.Text) & ""                    '6 순서
'
'                Call ExecPro("SP_01011_02", sValue(), Err_Num, Err_Dec)
'            Next Cloth_Cnt
'        End With
'
'        With sprList
'            For Cloth_Cnt = 1 To .MaxRows
'                .Row = Cloth_Cnt
'                .Col = 1
'
'                If .Text = "품목" Then
'                    .Col = 2
'
'                    Select Case .Text
'                        Case "추가", "수정"
'                            .Col = 3
'                            Query = "SELECT * FROM TB_가맹점의류"
'                            Query = Query & " WHERE 가맹점코드 = '" & 가맹점코드 & "'"
'                            Query = Query & "   AND 적용일자   = '" & Format(dtInput.Value, "YYYY-MM-DD") & "'"
'                            Query = Query & "   AND 의류코드   = '" & .Text & "'"
'                            Set RS01 = New ADODB.Recordset
'                            RS01.Open Query, ADOCon, adOpenDynamic, adLockOptimistic
'
'                            If RS01.EOF Then RS01.AddNew
'
'                                      RS01!가맹점코드 = 가맹점코드 & ""                        ' 1
'                                      RS01!적용일자 = Format(dtInput.Value, "YYYY-MM-DD") & "" ' 2
'                            .Col = 3: RS01!의류코드 = Trim(.Text) & ""                         ' 3
'                            .Col = 4: RS01!의류명 = Trim(.Text) & ""                           ' 4
'                            .Col = 5: RS01!금액 = .Value                                       ' 5
'                                      RS01!순서 = ""                                           ' 6
'                                      RS01!사용여부 = "Y"                                      ' 7
'
'                            RS01.Update
'
'
'                            RS01.Close
'                            Set RS01 = Nothing
'
'                        Case "삭제"
'                            .Col = 3
'                            Query = "DELETE FROM TB_가맹점의류"
'                            Query = Query & " WHERE 가맹점코드 = '" & 가맹점코드 & "'"
'                            Query = Query & "   AND 적용일자   = '" & Format(dtInput.Value, "YYYY-MM-DD") & "'"
'                            Query = Query & "   AND 의류코드   = '" & .Text & "'"
'                            ADOCon.Execute Query
'                    End Select
'
'                Else '의류분류
'                    .Col = 2
'
'                    Select Case .Text
'                        Case "추가"
'                        Case "수정"
'                        Case "삭제"
'                    End Select
'                End If
'            Next Cloth_Cnt
'        End With
'    Next i
'
'    ProgressBar.Value = 0
'    ProgressBar.Text = ""
'
'    ProgressBar2.Value = 0
'    ProgressBar2.Text = ""
'
'    If Err_Num = 0 Then
'        MsgBox "해당되는 데이터가 정상적으로 저장이 되었습니다.", vbInformation
'    End If
'
'    cmdBtn(2).Enabled = True
'
'    Exit Sub
'
'ErrRtn:
'    cmdBtn(2).Enabled = True
'
'    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
'End Sub


Private Sub DataSave()
    On Error GoTo ErrRtn
    
    Dim dMoney(1)   As Double
    
    Dim i         As Long
    Dim Cloth_Cnt As Long
    
    Dim Query     As String
    
    If sprList.DataRowCnt <= 0 Then Exit Sub
    
    If chkSaveAction.Value = xtpChecked Then
        If MsgBox("가맹점 자료 적용일자 " & Format(dtInput.Value, "YYYY-MM-DD") & "일 입니다. 계속 작업하시겠습니까?", vbInformation + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    cmdBtn(2).Enabled = False
    DoEvents
    
    For i = 1 To sprList.MaxRows
        sprList.Row = i
        sprList.Col = 1
        
        If sprList.Text = "품목" Then
            sprList.Col = 2
            
            Select Case sprList.Text
                Case "추가", "수정"
                    '------------------------------------------------------
                    ' TB_의류 - SP_01003_01
                    '------------------------------------------------------
                    ReDim sValue(3)
                    
                    sprList.Col = 3: sValue(0) = sprList.Text  '코드
                    sprList.Col = 4: sValue(1) = sprList.Text  '품명
                    sprList.Col = 5: sValue(2) = sprList.Value '금액
                    
                    ' 금액이 변동된 경우 변경된 금액으로 전체 매장을 설정 하기 위하여 작업 일자를
                    ' 넣고 해당 프로시저에서 작업일자가 오늘인것에 한하여 해당 작업을 하도록 처리한다.
                                     dMoney(0) = sprList.Value '신규 금액
                    sprList.Col = 8: dMoney(1) = sprList.Value '이전 금액
                    If dMoney(0) <> dMoney(1) Then sValue(3) = Format(Date, "yyyy-MM-dd")
                    
                    Call ExecPro("SP_01003_01", sValue(), Err_Num, Err_Dec)
                
                Case "삭제"
                    '------------------------------------------------------
                    ' TB_의류 - SP_01003_01
                    '------------------------------------------------------
                    ReDim sValue(0)
                    
                    sprList.Col = 3: sValue(0) = sprList.Text '코드
                    
                    Call ExecPro("SP_01003_02", sValue(), Err_Num, Err_Dec)
            End Select
            
        Else
            sprList.Col = 2
            
            Select Case sprList.Text
                Case "추가", "수정"
                    '------------------------------------------------------
                    ' TB_의류분류 - SP_01003_A_00
                    '------------------------------------------------------
                    ReDim sValue(4)
                    
                    sprList.Col = 3: sValue(0) = sprList.Text        '코드
                    sprList.Col = 4: sValue(1) = Trim(sprList.Text)  '품명
                                     sValue(2) = 0                   '
                                     sValue(3) = 0                   '
                                     sValue(4) = 0                   '
                                     
                    Call ExecPro("SP_01003_A_00", sValue(), Err_Num, Err_Dec)
                
                Case "삭제"
                    '------------------------------------------------------
                    ' TB_의류분류 - SP_01003_A_01
                    '------------------------------------------------------
                    ReDim sValue(0)
                    
                    sprList.Col = 3: sValue(0) = sprList.Text '코드
                    
                    Call ExecPro("SP_01003_A_01", sValue(), Err_Num, Err_Dec)
            End Select
        End If
    Next i
    
    
        
    '==========================================================================
    
    If chkSaveAction.Value = xtpChecked Then
        Dim 가맹점코드 As String
        Dim 의류코드   As String
        
        ProgressBar.Value = 0
        ProgressBar.Min = 0
        ProgressBar.Max = sprBranch.MaxRows
        
        ProgressBar2.Value = 0
        ProgressBar2.Min = 0
        ProgressBar2.Max = 100
        
        For i = 1 To sprBranch.MaxRows
            sprBranch.Row = i
            sprBranch.Col = 1: 가맹점코드 = sprBranch.Text & ""
            sprBranch.Col = 2: ProgressBar.Text = sprBranch.Text & ""
            
            ProgressBar.Value = i
            DoEvents
            
            If Trim(가맹점코드) <> "" And Len(가맹점코드) = 6 Then
                
                '------------------------------------------------------
                ' TB_가맹점의류
                '------------------------------------------------------
                ReDim sValue(2)
                
                sValue(0) = 가맹점코드                              ' 가맹점 코드
                sValue(1) = Format(dtInput.Value, "YYYY-MM-DD")     ' 적용일자
                sValue(2) = Format(Date, "YYYY-MM-DD")              ' 작업일자
                
                Call ExecPro("SP_01011_06", sValue(), Err_Num, Err_Dec)
                If Err_Num <> 0 Then
                    MsgBox 가맹점코드 & " 가맹점 저장중 오류가 발생 하였습니다. " & vbNewLine & Err_Dec, vbCritical
                End If
            End If
        Next i
        
        ProgressBar.Value = 0
        ProgressBar.Text = ""
        
        ProgressBar2.Value = 0
        ProgressBar2.Text = ""
    
    End If
    
    If Err_Num = 0 Then
        MsgBox "해당되는 데이터가 정상적으로 저장이 되었습니다.", vbInformation
    End If
    
    cmdBtn(2).Enabled = True
    
    Exit Sub
    
ErrRtn:
    cmdBtn(2).Enabled = True
    
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub
Private Sub DataDelete()
'    If MsgBox("해당되는 데이터를 삭제하시겠습니까?", vbYesNo + vbInformation + vbDefaultButton2, "데이터 삭제") = vbYes Then
'
'        ReDim sValue(0)
'
'        spdView.Row = spdView.ActiveRow
'        spdView.Col = 1
'        sValue(0) = spdView.Text
'
'        Call ExecPro("SP_01003_02_A", sValue(), Err_Num, Err_Dec)
'
'        If Err_Num = 0 Then
'            spdView.Row = spdView.ActiveRow
'            spdView.Action = SS_ACTION_DELETE_ROW
''            spdView.MaxRows = spdView.MaxRows - 1
'
'            MsgBox "해당되는 데이터가 정상적으로 삭제가 되었습니다.", vbInformation
'        End If
'    End If
End Sub

Private Sub DataCancel()
'    Call Data_Display
End Sub

Private Sub spdView_Click(ByVal Col As Long, ByVal Row As Long)
    If Row <= 0 Then Exit Sub
    
    spdView.Row = Row
    
    cboGubun.ListIndex = 1                                 '구분
    spdView.Col = 1: txtData(0).Text = spdView.Text & "" '품목분류코드
    spdView.Col = 2: txtData(1).Text = spdView.Text & "" '품목분류명
    spdView.Col = 3: txtMoney.Value = spdView.Value & "" '금액
                     txtOldMoney.Value = txtMoney.Value ' 이전금액

End Sub

'
Private Sub sprClass_Click(ByVal Col As Long, ByVal Row As Long)
    Dim 의류분류코드 As String
    
    If Row <= 0 Then Exit Sub
    
    sprClass.Row = Row
    sprClass.Col = 1: 의류분류코드 = sprClass.Text & ""
    
    If 의류분류코드 = "" Then
        txtData(0).Text = ""
        txtData(1).Text = ""
        txtMoney.Value = 0
    Else
        cboGubun.ListIndex = 0                                 '구분
        sprClass.Col = 1: txtData(0).Text = sprClass.Text & "" '품목분류코드
        sprClass.Col = 2: txtData(1).Text = sprClass.Text & "" '품목분류명
                          txtMoney.Value = 0                   '
    End If
    
    Call 의류_Display(의류분류코드)
End Sub

Private Sub sprClass_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Call sprClass_Click(NewCol, NewRow)
End Sub

Private Sub 의류_Display(의류분류코드 As String)
    ReDim sValue(0)
                       '
    sValue(0) = 의류분류코드

    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01003_00", sValue(), Err_Num, Err_Dec)

    With spdView
        .MaxRows = 0
        .Redraw = False

        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows

            .Col = 1: .Text = RS01!코드 & ""
            .Col = 2: .Text = RS01!품목명 & ""
            .Col = 3: .Text = RS01!단가 & ""

            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing

        .Redraw = True
    End With
End Sub

Private Sub Data_Display(가맹점코드 As String)
    On Error GoTo ErrRtn

    ReDim sValue(0)
                       '
    sValue(0) = 가맹점코드

    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01003_03", sValue(), Err_Num, Err_Dec)

    With spdView
        .MaxRows = 0
        .Redraw = False

        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows

            .Col = 1: .Text = RS01!의류코드 & ""
            .Col = 2: .Text = RS01!의류명 & ""
            .Col = 3: .Text = RS01!금액 & ""

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
Private Sub 가맹점_Display()
    ReDim sValue(0)
    
    sValue(0) = "0"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_00003_01", sValue(), Err_Num, Err_Dec)

    With sprBranch
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = RS01(0) & ""  '2
            .Col = 2: .Text = RS01(1) & ""  '3
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
End Sub

Private Sub sprList_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Call sprList.DeleteRows(Row, 1)
    
    sprList.MaxRows = sprList.MaxRows - 1
End Sub
