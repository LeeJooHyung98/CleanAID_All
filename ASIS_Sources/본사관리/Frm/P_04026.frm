VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_04026 
   Caption         =   "[전사업장] 특정매장 분석"
   ClientHeight    =   12450
   ClientLeft      =   690
   ClientTop       =   1995
   ClientWidth     =   20370
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_04026.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12450
   ScaleWidth      =   20370
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   12450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   21960
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_04026.frx":058A
      Begin Threed.SSPanel panInput 
         Height          =   795
         Left            =   15
         TabIndex        =   1
         Top             =   540
         Width           =   20340
         _ExtentX        =   35878
         _ExtentY        =   1402
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.CommandButton cmdAllCheck 
            Caption         =   "전체 선택"
            Height          =   315
            Left            =   8280
            TabIndex        =   2
            Top             =   360
            Width           =   1305
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            TabIndex        =   3
            Top             =   405
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   63635456
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   4
            Top             =   405
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "매 출 기 간"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   4815
            TabIndex        =   5
            Top             =   405
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   63635456
            CurrentDate     =   36686
         End
         Begin XtremeSuiteControls.CheckBox chkSelect 
            Height          =   195
            Index           =   1
            Left            =   11190
            TabIndex        =   42
            Tag             =   "유통매장"
            Top             =   150
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
            Left            =   9690
            TabIndex        =   43
            Tag             =   "일반매장"
            Top             =   150
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
            Left            =   12600
            TabIndex        =   44
            Tag             =   "이마트"
            Top             =   150
            Width           =   945
            _Version        =   851970
            _ExtentX        =   1667
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "이마트"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox chkSelect 
            Height          =   195
            Index           =   3
            Left            =   13740
            TabIndex        =   45
            Tag             =   "크렌즈"
            Top             =   150
            Width           =   1095
            _Version        =   851970
            _ExtentX        =   1931
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "크렌즈"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox chkSelect2 
            Height          =   195
            Index           =   0
            Left            =   9690
            TabIndex        =   46
            Tag             =   "폐점"
            Top             =   450
            Width           =   1515
            _Version        =   851970
            _ExtentX        =   2672
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "폐점 포함"
            ForeColor       =   255
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox chkSelect 
            Height          =   195
            Index           =   4
            Left            =   14880
            TabIndex        =   47
            Tag             =   "유니트샵"
            Top             =   150
            Width           =   1185
            _Version        =   851970
            _ExtentX        =   2090
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "유니트샵"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label Label1 
            Caption         =   "~"
            Height          =   225
            Left            =   4620
            TabIndex        =   6
            Top             =   465
            Width           =   225
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
         Caption         =   " 특정매장 분석 (P_04026)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_04026.frx":065C
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
         Width           =   11460
         _ExtentX        =   20214
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
         PictureBackground=   "P_04026.frx":085E
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
            Picture         =   "P_04026.frx":0A60
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   10
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
            Picture         =   "P_04026.frx":0FFA
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   11
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
            Picture         =   "P_04026.frx":1594
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   12
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
            Picture         =   "P_04026.frx":1B2E
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   13
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
            Picture         =   "P_04026.frx":20C8
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   14
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
            Picture         =   "P_04026.frx":2662
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   15
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
            Picture         =   "P_04026.frx":2BFC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   49
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
            Picture         =   "P_04026.frx":3196
         End
      End
      Begin FPSpreadADO.fpSpread spdView2 
         Height          =   10230
         Left            =   5760
         TabIndex        =   16
         Top             =   1350
         Width           =   14595
         _Version        =   524288
         _ExtentX        =   25744
         _ExtentY        =   18045
         _StockProps     =   64
         BackColorStyle  =   1
         ColsFrozen      =   2
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   37
         SpreadDesigner  =   "P_04026.frx":3730
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   840
         Index           =   1
         Left            =   5760
         TabIndex        =   17
         Top             =   11595
         Width           =   14595
         _ExtentX        =   25744
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
            Caption         =   "반품 금액"
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
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   11085
         Left            =   15
         TabIndex        =   48
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
         MaxCols         =   5
         ScrollBars      =   2
         SpreadDesigner  =   "P_04026.frx":4C10
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_04026"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String
Dim bChkFlag    As Boolean

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub chkSelect_Click(Index As Integer)
    Dim nRow    As Long
    Dim vText   As Variant
    
    With spdView
        For nRow = 1 To .MaxRows
            .GetText 4, nRow, vText
            
            ' 현재의 가맹점 종류와 선택 가맹점의 종류가 같을 경우
            If CStr(vText) = chkSelect(Index).Tag Then
                .SetText 2, nRow, CVar(chkSelect(Index).Value)
                
                ' 폐점 여부 다시 확인
                .GetText 5, nRow, vText
                If CStr(vText) = chkSelect2(0).Tag Then
                    If chkSelect2(0).Value = xtpUnchecked Then
                        .SetText 2, nRow, "0"
                    End If
                End If
                
                
            End If
        
        Next nRow
    End With
End Sub

' 제외 취소
Private Sub chkSelect2_Click(Index As Integer)
    Dim Idx     As Long
    Dim nRow    As Long
    Dim vText   As Variant
    
    With spdView
        For nRow = 1 To .MaxRows
            .GetText 5, nRow, vText
            
            ' 현재의 가맹점 종류와 선택 가맹점의 종류가 같을 경우
            If CStr(vText) = chkSelect2(Index).Tag Then
            
                ' 매장 구분을 가저온다.
                .GetText 4, nRow, vText
                For Idx = 0 To 4
                    If chkSelect(Idx).Tag = CStr(vText) Then
                        
                        ' 선택된 구분에 포함일 경우
                        If chkSelect(Idx).Value = xtpChecked Then
                            .SetText 2, nRow, CVar(chkSelect2(Index).Value)
                            
                        ' 선택이 아닐경우 무조건 미처리
                        Else
                            .SetText 2, nRow, "0"
                        End If
                            
                    End If
                Next Idx
            End If
        
        Next nRow
    End With
End Sub

Private Sub cmdAllCheck_Click()
    Dim vText   As Variant
    Dim nRow    As Long
    
    With spdView
        '.EventEnabled(EventButtonClicked) = False
        
        For nRow = 1 To .MaxRows
            .GetText 3, nRow, vText
            
            If CStr(vText) = "" Then
                .SetText 2, nRow, CVar(IIf(cmdAllCheck.Caption = "전체 선택", "1", "0"))
            Else
                Exit For
            End If
        
        Next nRow
        '.EventEnabled(EventButtonClicked) = True
    End With
            
    If cmdAllCheck.Caption = "전체 선택" Then
        cmdAllCheck.Caption = "전체 취소"
        chkSelect(0).Value = xtpChecked
        chkSelect(1).Value = xtpChecked
        chkSelect(2).Value = xtpChecked
        chkSelect(3).Value = xtpChecked
        chkSelect(4).Value = xtpChecked
        chkSelect2(0).Value = xtpChecked
    
    Else
        cmdAllCheck.Caption = "전체 선택"
        chkSelect(0).Value = xtpUnchecked
        chkSelect(1).Value = xtpUnchecked
        chkSelect(2).Value = xtpUnchecked
        chkSelect(3).Value = xtpUnchecked
        chkSelect(4).Value = xtpUnchecked
        chkSelect2(0).Value = xtpUnchecked
    
    End If

End Sub

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 0: Call Data_Display           ' 조회
        Case 1:                ' 신규
        Case 2:                 ' 저장
        Case 3:            ' 삭제
        Case 4:            ' 취소
        Case 5:            ' 인쇄
        Case 6: Call Export_Excel(P_00000.cdgExcel, spdView2)      ' 엑셀
        Case 7: Unload Me           ' 종료
        
        Case Else
            '
    End Select

End Sub

Private Sub ComboBox1_KeyDown(KeyCode As Integer, Shift As Integer)

End Sub

Private Sub Form_Activate()
    Dim nRow    As Long
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"

    Call SubBottonEnable(cmdBtn, "10000011")
    
    If P_04026_Flag = False Then
        Screen.MousePointer = vbHourglass
        dtInput(0).Value = Date
        dtInput(1).Value = Date
        
        ReDim sValue(1)
        
        sValue(0) = "0"
        sValue(1) = IIf(Store.Code = "1000", "%", Store.Code & "%")
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_04026_00", sValue(), Err_Num, Err_Dec)
        
        spdView.MaxRows = RS01.RecordCount
        Call spdDisplay(RS01)
        
        For nRow = 1 To spdView.DataRowCnt
            spdView.Row = nRow: spdView.Col = 5
            If InStr(spdView.Text, "폐점") > 0 Then
                spdView.Col = -1
                spdView.ForeColor = vbRed
            End If
        Next nRow
        
        
        
        P_04026_Flag = True
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub spdDisplay(Rs As ADODB.Recordset)

    Call fpSpread_Display(spdView, Rs)
    
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
    
    P_04026_Flag = False
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
            If IsNumeric(Mid(spdView.Text, 6, 1)) Then '<- 매장만, 지사제외
                        
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
    
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    If sValue(0) = "" Then Exit Sub
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_04026_01", sValue(), Err_Num, Err_Dec)
    
    With spdView2
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = RS01!마감일자 & ""              ' 1
            .Col = 2:  .Text = ExecWeekDay(RS01!마감일자) & "" ' 2
            .Col = 3:  .Text = RS01!가맹점코드 & ""            ' 3
            .Col = 4:  .Text = RS01!가맹점명 & ""              ' 4
            .Col = 5:  .Text = RS01!지사금액 & ""              ' 5
            .Col = 6:  .Text = RS01!가맹점금액 & ""            ' 6
            .Col = 7:  .Text = RS01!접수수량 & ""              ' 7
            .Col = 8:  .Text = RS01!출고수량 & ""              ' 8
                        
            If Len(RS01!시작택번호) = 9 Then
                .Col = 9:  .Text = Format(RS01!시작택번호, "000-00-0000") & ""         ' 9
            Else
                .Col = 9:  .Text = RS01!시작택번호 & ""        ' 9
            End If
            
            If Len(RS01!종료택번호) = 9 Then
                .Col = 10: .Text = Format(RS01!종료택번호, "000-00-0000") & ""         '10
            Else
                .Col = 10: .Text = RS01!종료택번호 & ""        '10
            End If
            
            .Col = 11
            Select Case RS01!판매구분
                Case "1": .Text = "세일"       '11
                Case "2": .Text = "요일"       '
                Case "3": .Text = "정상"      '
            End Select

            .Col = 12: .Text = RS01!접수금액 & ""                   '12
            .Col = 13: .Text = RS01!현금입금 + RS01!카드금액 & ""   '13
            
            If RS01!접수수량 = 0 Then
                .Col = 14: .Text = 0 & ""   '14
                .Col = 15: .Text = 0 & ""   '15
                .Col = 16: .Text = 0 & ""   '16
            Else
                .Col = 14: .Text = RS01!접수금액 / RS01!접수수량 & ""   '14
                .Col = 15: .Text = RS01!지사금액 / RS01!접수수량 & ""   '15
                .Col = 16: .Text = RS01!가맹점금액 / RS01!접수수량 & "" '16
            End If
            
            .Col = 17: .Text = RS01!로열티금액1 & ""                   '17
            .Col = 18: .Text = RS01!로열티금액2 & ""                   '17
            .Col = 19: .Text = RS01!지사차감후 & ""                   '17
            .Col = 20: .Text = RS01!수수료승인금액 & ""                   '17
            .Col = 21: .Text = RS01!수수료취소금액 & ""                   '17
            .Col = 22: .Text = RS01!수수료지원금액 & ""                   '17
            
            .Col = 23: .Text = RS01!현금입금 & ""                   '17
            .Col = 24: .Text = RS01!카드금액 & ""                   '18
            .Col = 25: .Text = RS01!카드건수 & ""                   '19
            .Col = 26: .Text = RS01!쿠폰금액 & ""                   '20
            .Col = 27: .Text = RS01!쿠폰건수 & ""                   '21
            .Col = 28: .Text = RS01!발생마일리지 & ""               '22
            .Col = 29: .Text = RS01!사용마일리지 & ""               '23
            .Col = 30: .Text = RS01!삭제마일리지 & ""               '24
            .Col = 31: .Text = RS01!반품환불금액 & ""               '25
            .Col = 32: .Text = RS01!반품환불건수 & ""               '26
            .Col = 33: .Text = RS01!세탁환불금액 & ""               '27
            .Col = 34: .Text = RS01!세탁환불건수 & ""               '28
            .Col = 35: .Text = RS01!재세탁수량 & ""                 '29
            .Col = 36: .Text = RS01!수선금액 & ""                   '30
            .Col = 37: .Text = RS01!수선수량 & ""                   '31
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        If .MaxRows > 0 Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Row = .Row
            .Row2 = .Row
            .Col = 1
            .Col2 = .MaxCols
            .BlockMode = True
            .BackColor = &HC0FFC0
            .BlockMode = False
        
            .Col = 4:  .Text = "합계"
            
            .Col = 5:  .Formula = "SUM(E1:E" & .MaxRows - 1 & ")"
            .Col = 6:  .Formula = "SUM(F1:F" & .MaxRows - 1 & ")"
            .Col = 7:  .Formula = "SUM(G1:G" & .MaxRows - 1 & ")"
            .Col = 8:  .Formula = "SUM(H1:H" & .MaxRows - 1 & ")"
            
            .Col = 12: .Formula = "SUM(L1:L" & .MaxRows - 1 & ")"
            .Col = 13: .Formula = "SUM(M1:M" & .MaxRows - 1 & ")"
            
            .Col = 14: .Formula = "SUM(N1:N" & .MaxRows - 1 & ") / " & .MaxRows - 1 & ""
            .Col = 15: .Formula = "SUM(O1:O" & .MaxRows - 1 & ") / " & .MaxRows - 1 & ""
            .Col = 16: .Formula = "SUM(P1:P" & .MaxRows - 1 & ") / " & .MaxRows - 1 & ""
            
            .Col = 17: .Formula = "SUM(Q1:Q" & .MaxRows - 1 & ")"
            .Col = 18: .Formula = "SUM(R1:R" & .MaxRows - 1 & ")"
            .Col = 19: .Formula = "SUM(S1:S" & .MaxRows - 1 & ")"
            .Col = 20: .Formula = "SUM(T1:T" & .MaxRows - 1 & ")"
            .Col = 21: .Formula = "SUM(U1:U" & .MaxRows - 1 & ")"
            .Col = 22: .Formula = "SUM(V1:V" & .MaxRows - 1 & ")"
            .Col = 23: .Formula = "SUM(W1:W" & .MaxRows - 1 & ")"
            
            .Col = 24: .Formula = "SUM(X1:X" & .MaxRows - 1 & ")"
            .Col = 25: .Formula = "SUM(Y1:Y" & .MaxRows - 1 & ")"
            .Col = 26: .Formula = "SUM(Z1:Z" & .MaxRows - 1 & ")"
            .Col = 27: .Formula = "SUM(AA1:AA" & .MaxRows - 1 & ")"
            .Col = 28: .Formula = "SUM(AB1:AB" & .MaxRows - 1 & ")"
            .Col = 29: .Formula = "SUM(AC1:AC" & .MaxRows - 1 & ")"
            .Col = 30: .Formula = "SUM(AD1:AD" & .MaxRows - 1 & ")"
            .Col = 31: .Formula = "SUM(AE1:AE" & .MaxRows - 1 & ")"
            
            .Col = 32: .Formula = "SUM(AF1:AF" & .MaxRows - 1 & ")"
            .Col = 33: .Formula = "SUM(AG1:AG" & .MaxRows - 1 & ")"
            .Col = 34: .Formula = "SUM(AH1:AH" & .MaxRows - 1 & ")"
            .Col = 35: .Formula = "SUM(AI1:AI" & .MaxRows - 1 & ")"
            .Col = 36: .Formula = "SUM(AJ1:AJ" & .MaxRows - 1 & ")"
            .Col = 37: .Formula = "SUM(AK1:AK" & .MaxRows - 1 & ")"
            
            '------------------------------------------------------
            
            .Col = 12:  txtNum(0).Value = .Value  '전체매출액
            .Col = 14: txtNum(10).Value = .Value '전체단가
            .Col = 15: txtNum(11).Value = .Value '지사단가

            .Col = 5: txtNum(1).Value = .Value   '지사매출
            .Col = 6: txtNum(2).Value = .Value   '가맹점매출

            .Col = 7: txtNum(3).Value = .Value   '입고수량

            .Col = 36: txtNum(7).Value = .Value   '수선금액
            .Col = 37: txtNum(6).Value = .Value   '수선수량

            .Col = 24: txtNum(4).Value = .Value   '카드금액
            .Col = 25: txtNum(5).Value = .Value   '카드수량

            .Col = 31: txtNum(8).Value = .Value   '반품수량
            .Col = 35: txtNum(9).Value = .Value   '재세탁수량
        End If
        
        '.Row = 1:   .Col = -1:  oldColor = .BackColor
            
        ' 누락된 요일을 설정한다.
'        Call DateCheckAdd(spdView, Format(dtInput(0).Value, "YYYY-MM-DD"), Format(dtInput(1).Value, "YYYY-MM-DD"))
        .Redraw = True
    End With
        
    
    
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub


