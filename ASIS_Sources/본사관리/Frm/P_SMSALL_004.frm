VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_SMSALL_4 
   Caption         =   "3큐브"
   ClientHeight    =   12450
   ClientLeft      =   3465
   ClientTop       =   2940
   ClientWidth     =   16500
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_SMSALL_004.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12450
   ScaleWidth      =   16500
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   12450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16500
      _ExtentX        =   29104
      _ExtentY        =   21960
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_SMSALL_004.frx":058A
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   11085
         Left            =   15
         TabIndex        =   1
         Top             =   1350
         Width           =   5730
         _Version        =   524288
         _ExtentX        =   10107
         _ExtentY        =   19553
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
         ScrollBars      =   2
         SpreadDesigner  =   "P_SMSALL_004.frx":065C
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   795
         Left            =   15
         TabIndex        =   2
         Top             =   540
         Width           =   16470
         _ExtentX        =   29051
         _ExtentY        =   1402
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.CommandButton cmdAllCheck 
            Caption         =   "전체 선택"
            Height          =   315
            Left            =   9480
            TabIndex        =   4
            Top             =   240
            Width           =   1305
         End
         Begin XtremeSuiteControls.CheckBox chkSelect 
            Height          =   195
            Index           =   1
            Left            =   12420
            TabIndex        =   3
            Tag             =   "유통매장"
            Top             =   300
            Width           =   1185
            _Version        =   851970
            _ExtentX        =   2090
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "유통매장"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox chkSelect 
            Height          =   195
            Index           =   0
            Left            =   11010
            TabIndex        =   5
            Tag             =   "일반매장"
            Top             =   300
            Width           =   1185
            _Version        =   851970
            _ExtentX        =   2090
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "일반매장"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox chkSelect 
            Height          =   195
            Index           =   2
            Left            =   13830
            TabIndex        =   6
            Tag             =   "이마트"
            Top             =   300
            Width           =   1185
            _Version        =   851970
            _ExtentX        =   2090
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "이마트"
            UseVisualStyle  =   -1  'True
         End
         Begin Threed.SSOption optGubun 
            Height          =   240
            Index           =   0
            Left            =   1680
            TabIndex        =   42
            Top             =   450
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   423
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "3일"
            Value           =   -1
         End
         Begin Threed.SSOption optGubun 
            Height          =   240
            Index           =   1
            Left            =   2535
            TabIndex        =   43
            Top             =   450
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   423
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "3주(21일)"
         End
         Begin Threed.SSOption optGubun 
            Height          =   240
            Index           =   2
            Left            =   3885
            TabIndex        =   44
            Top             =   450
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   423
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "3개월(90일)"
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1590
            TabIndex        =   46
            Top             =   90
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Format          =   64487424
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   47
            Top             =   90
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "기 준 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   48
            Top             =   420
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "검 색 조 건"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox chkSelect 
            Height          =   195
            Index           =   3
            Left            =   15030
            TabIndex        =   49
            Tag             =   "폐점"
            Top             =   300
            Width           =   1185
            _Version        =   851970
            _ExtentX        =   2090
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "폐점"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   7
         Top             =   15
         Width           =   8865
         _ExtentX        =   15637
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
         Caption         =   " 3큐브 (P_SMSALL_4)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_SMSALL_004.frx":0C15
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   0
         Left            =   8895
         TabIndex        =   8
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
         PictureBackground=   "P_SMSALL_004.frx":0E17
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   9
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
            Picture         =   "P_SMSALL_004.frx":1019
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   10
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "화면"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_SMSALL_004.frx":15B3
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   11
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
            Picture         =   "P_SMSALL_004.frx":1B4D
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   12
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
            Picture         =   "P_SMSALL_004.frx":20E7
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   13
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
            Picture         =   "P_SMSALL_004.frx":2681
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   14
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
            Picture         =   "P_SMSALL_004.frx":2C1B
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   15
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
            Picture         =   "P_SMSALL_004.frx":31B5
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   16
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
            Picture         =   "P_SMSALL_004.frx":374F
         End
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   840
         Index           =   1
         Left            =   5760
         TabIndex        =   17
         Top             =   11595
         Width           =   10725
         _ExtentX        =   18918
         _ExtentY        =   1482
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   14
            Left            =   60
            TabIndex        =   18
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   582
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "전체매출액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   15
            Left            =   2340
            TabIndex        =   19
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   582
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "지사매출액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   16
            Left            =   4620
            TabIndex        =   20
            Top             =   375
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   582
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "입고 수량"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   17
            Left            =   4620
            TabIndex        =   21
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   582
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "가맹점매출액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   18
            Left            =   9180
            TabIndex        =   22
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   582
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "카드 금액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   19
            Left            =   6900
            TabIndex        =   23
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   582
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "수선 금액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   20
            Left            =   9180
            TabIndex        =   24
            Top             =   375
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   582
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "카드 건수"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   21
            Left            =   6900
            TabIndex        =   25
            Top             =   375
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   582
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "수선 수량"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   22
            Left            =   11460
            TabIndex        =   26
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   582
            _Version        =   262144
            BackColor       =   12632319
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "반품 수량"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   23
            Left            =   11460
            TabIndex        =   27
            Top             =   375
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   582
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "재세탁수량"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   24
            Left            =   60
            TabIndex        =   28
            Top             =   375
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   582
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "전체 단가"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   25
            Left            =   2340
            TabIndex        =   29
            Top             =   375
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   582
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "지사 단가"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   330
            Index           =   0
            Left            =   1200
            TabIndex        =   30
            Top             =   60
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
            _ExtentY        =   582
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
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
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   330
            Index           =   10
            Left            =   1200
            TabIndex        =   31
            Top             =   375
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
            _ExtentY        =   582
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
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
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   330
            Index           =   1
            Left            =   3480
            TabIndex        =   32
            Top             =   60
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
            _ExtentY        =   582
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
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
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   330
            Index           =   11
            Left            =   3480
            TabIndex        =   33
            Top             =   375
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
            _ExtentY        =   582
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
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
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   330
            Index           =   2
            Left            =   5760
            TabIndex        =   34
            Top             =   60
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
            _ExtentY        =   582
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
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
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   330
            Index           =   7
            Left            =   8040
            TabIndex        =   35
            Top             =   60
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
            _ExtentY        =   582
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
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
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   330
            Index           =   4
            Left            =   10320
            TabIndex        =   36
            Top             =   60
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
            _ExtentY        =   582
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
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
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   330
            Index           =   3
            Left            =   5760
            TabIndex        =   37
            Top             =   375
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
            _ExtentY        =   582
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
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
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   330
            Index           =   6
            Left            =   8040
            TabIndex        =   38
            Top             =   375
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
            _ExtentY        =   582
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
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
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   330
            Index           =   5
            Left            =   10320
            TabIndex        =   39
            Top             =   375
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
            _ExtentY        =   582
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
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
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   330
            Index           =   8
            Left            =   12600
            TabIndex        =   40
            Top             =   60
            Width           =   930
            _Version        =   262145
            _ExtentX        =   1640
            _ExtentY        =   582
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
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
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   330
            Index           =   9
            Left            =   12600
            TabIndex        =   41
            Top             =   375
            Width           =   930
            _Version        =   262145
            _ExtentX        =   1640
            _ExtentY        =   582
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
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
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
      End
      Begin FPSpreadADO.fpSpread spdView1 
         Height          =   10230
         Left            =   5760
         TabIndex        =   45
         Top             =   1350
         Width           =   10725
         _Version        =   524288
         _ExtentX        =   18918
         _ExtentY        =   18045
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
         MaxCols         =   15
         SpreadDesigner  =   "P_SMSALL_004.frx":3CE9
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_SMSALL_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String
Dim sCheck  As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub chkSelect_Click(Index As Integer)
    Dim nRow    As Long
    
    With spdView
        For nRow = 1 To .MaxRows
            .Row = nRow
            .Col = 4
            
            ' 현재의 가맹점 종류와 선택 가맹점의 종류가 같을 경우
            If .Text = chkSelect(Index).Tag Then
                sCheck = "N"
                .Row = nRow
                .Col = 2
                .Action = ActionActiveCell
                .Value = chkSelect(Index).Value
            End If
        
        Next nRow
    End With

End Sub

Private Sub cmdAllCheck_Click()
    Dim nRow    As Long
    
    With spdView
        '.EventEnabled(EventButtonClicked) = False
        
        For nRow = 1 To .MaxRows
            .Row = nRow
            .Col = 3
            
            If .Text = "" Then
                
                .Row = nRow
                .Col = 2
                .Action = ActionActiveCell
                .Value = IIf(cmdAllCheck.Caption = "전체 선택", "1", "0")
            Else
                Exit For
            End If
        
        Next nRow
        '.EventEnabled(EventButtonClicked) = True
    End With
        
    cmdAllCheck.Caption = IIf(cmdAllCheck.Caption = "전체 선택", "전체 취소", "전체 선택")
        
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 0: Call Data_Display           ' 조회
        Case 1:                ' 신규
        Case 2:                 ' 저장
        Case 3:            ' 삭제
        Case 4:            ' 취소
        Case 5:            ' 인쇄
        Case 6:            ' 화면
        Case 7: Unload Me           ' 종료
        
        Case Else
            '
    End Select

End Sub

Private Sub Form_Activate()
    Dim nRow    As Long
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"

    Call SubBottonEnable(cmdBtn, "10000001")
    
    If P_SMSALL_4_Flag = False Then
        Screen.MousePointer = vbHourglass
        dtInput(0).Value = Date
        
        
        ReDim sValue(1)
        
        sValue(0) = "0"
        sValue(1) = IIf(Store.Code = "1000", "%", Store.Code & "%")
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_M_SMSALL_004_00", sValue(), Err_Num, Err_Dec)
        
        spdView.MaxRows = RS01.RecordCount
        Call spdDisplay(RS01)
        
        For nRow = 1 To spdView.DataRowCnt
            spdView.Row = nRow: spdView.Col = 4
            If InStr(spdView.Text, "폐점") > 0 Then
                spdView.Col = -1
                spdView.ForeColor = vbRed
            End If
        Next nRow
        
        
        
        P_SMSALL_4_Flag = True
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub spdDisplay(Rs As ADODB.Recordset)

    Call fpSpread_Display(spdView, Rs)
    
End Sub

Private Sub Form_Load()
    dtInput(0).Value = Date
End Sub

Private Sub spdView_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    
    spdView.EventEnabled(EventButtonClicked) = False
    
    
    If Row = spdView.ActiveRow Then
                    
        Dim nRow    As Long
        ReDim sValue(2)
        
        If Col = 2 Then
            spdView.Row = spdView.ActiveRow
            spdView.Col = Col
            If spdView.Value = False Then
                spdView.Col = 2
                spdView.Text = ""
            
                ' 선택 내용이 지사일 경우 해당 체인점을 모두 선택 시킨다.
                spdView.Col = 1
                sValue(2) = Mid(spdView.Text, 2, 6)
                If Mid(sValue(2), 5, 1) = "]" Then
                    
                    sValue(2) = Left(sValue(2), 4)
                    For nRow = 1 To spdView.MaxRows
                        spdView.Row = nRow
                        spdView.Col = 3
                        If spdView.Text = sValue(2) Then
                            spdView.Col = 2
                            spdView.Value = "0"
                        
                        End If
                    Next nRow
                End If
        
            Else

                
                spdView.Row = Row
                spdView.Col = 2: spdView.Text = "1"
                
                
                ' 선택 내용이 지사일 경우 해당 체인점을 모두 선택 시킨다.
                spdView.Col = 1: sValue(2) = Mid(spdView.Text, 2, 6)
                If Mid(sValue(2), 5, 1) = "]" Then
                    sValue(2) = Left(sValue(2), 4)
                    
                    ' 해당 지사의 매장일 경우
                    For nRow = 1 To spdView.MaxRows
                        spdView.Row = nRow
                        spdView.Col = 3
                        If spdView.Text = sValue(2) Then
                            spdView.Col = 2: spdView.Value = "1"
                        End If
                    Next nRow
                
                End If
            End If
        End If
    End If
    
    spdView.EventEnabled(EventButtonClicked) = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cmdBtn(0).Enabled = False
    cmdBtn(1).Enabled = False
    cmdBtn(2).Enabled = False
    cmdBtn(3).Enabled = False
    cmdBtn(4).Enabled = False
    cmdBtn(5).Enabled = False
    cmdBtn(6).Enabled = False
    
    P_SMSALL_4_Flag = False
End Sub

Public Sub DataSave()

End Sub

Public Sub DataAdd()

End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim nRow As Long
    Dim SSQL2   As String
    
    ReDim sValue(2)
    
    SSQL2 = ""
    '-------------------------------------------------------------------------------
    ' 선택된 체인점의 내용만을 구해서 쿼리한다.
    For nRow = 0 To spdView.MaxRows
        spdView.Col = 2:  spdView.Row = nRow
        ' 매장코드만 적용한다.
        If spdView.Text = "1" Then
            spdView.Col = 1
            If IsNumeric(Mid(spdView.Text, 6, 1)) Then
                        
                spdView.Col = 1
                SSQL2 = SSQL2 & "'" & Mid(spdView.Text, 2, 6) & "', "
            End If
        End If
    Next nRow
    
    SSQL2 = Trim(SSQL2)
    If Len(SSQL2) > 3 Then
        sValue(0) = Mid(SSQL2, 1, Len(SSQL2) - 1)           ' 마지막 ,을 제거한다.
    End If
        '-------------------------------------------------------------------------------
    
    Select Case True
        Case optGubun(0).Value: sValue(1) = Format(DateAdd("d", -3, dtInput(0).Value), "YYYY-MM-DD")
        Case optGubun(1).Value: sValue(1) = Format(DateAdd("d", -21, dtInput(0).Value), "YYYY-MM-DD")
        Case optGubun(2).Value: sValue(1) = Format(DateAdd("d", -90, dtInput(0).Value), "YYYY-MM-DD")
    End Select
    
    If sValue(0) = "" Then Exit Sub
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_SMSALL_004_01", sValue(), Err_Num, Err_Dec)
    
    With spdView1
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = RS01!가맹점정보 & ""
            .Col = 2:  .Text = RS01!고객코드 & ""
            .Col = 3:  .Text = RS01!고객등급코드 & ""
            .Col = 4:  .Text = Trim(RS01!성명 & "")
            .Col = 5:  .Text = Trim(RS01!전화번호 & "")
            .Col = 6:  .Text = Trim(RS01!휴대전화 & "")
            .Col = 7:  .Text = Trim(RS01!등록일자 & "")
            .Col = 8:  .Text = Trim(RS01!최종거래일자 & "")
            .Col = 9: .Text = RS01!이용횟수 & ""
            .Col = 10: .Text = RS01!총접수금액 & ""
            .Col = 11: .Text = RS01!누적마일리지 & ""
            .Col = 12: .Text = RS01!사용가능마일리지 & ""
            .Col = 13: .Text = RS01!미수금액 & ""
            .Col = 14: .Text = RS01!SMS & ""
            .Col = 15: .Text = IIf(RS01!SMS & "" = "N", 0, 1)
            
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


