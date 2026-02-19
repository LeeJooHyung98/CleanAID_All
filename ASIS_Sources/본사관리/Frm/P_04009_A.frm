VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_04009_A 
   Caption         =   "[전사업장]월간 사업장 매출현황"
   ClientHeight    =   9915
   ClientLeft      =   2475
   ClientTop       =   6300
   ClientWidth     =   16395
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_04009_A.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9915
   ScaleWidth      =   16395
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   9915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16395
      _ExtentX        =   28919
      _ExtentY        =   17489
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_04009_A.frx":058A
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   1
         Top             =   540
         Width           =   16365
         _ExtentX        =   28866
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1245
            Style           =   2  '드롭다운 목록
            TabIndex        =   43
            Top             =   60
            Width           =   3060
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   2
            Top             =   420
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "수금년월"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   8
            Left            =   60
            TabIndex        =   3
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "지 사 명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   285
            Index           =   14
            Left            =   2475
            TabIndex        =   4
            Top             =   435
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   503
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "~"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
         End
         Begin XtremeSuiteControls.DateTimePicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1245
            TabIndex        =   41
            Top             =   420
            Width           =   1215
            _Version        =   851970
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   68
            CustomFormat    =   "yyyy-MM"
            Format          =   3
            UpDown          =   -1  'True
         End
         Begin XtremeSuiteControls.DateTimePicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   2730
            TabIndex        =   42
            Top             =   420
            Width           =   1215
            _Version        =   851970
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   68
            CustomFormat    =   "yyyy-MM"
            Format          =   3
            UpDown          =   -1  'True
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   5
         Top             =   15
         Width           =   8760
         _ExtentX        =   15452
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
         Caption         =   " #"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_04009_A.frx":063C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   1
         Left            =   8790
         TabIndex        =   6
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
         PictureBackground=   "P_04009_A.frx":083E
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   7
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
            Picture         =   "P_04009_A.frx":0A40
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   8
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
            Picture         =   "P_04009_A.frx":0FDA
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   9
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
            Picture         =   "P_04009_A.frx":1574
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   10
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
            Picture         =   "P_04009_A.frx":1B0E
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   11
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
            Picture         =   "P_04009_A.frx":20A8
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   12
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
            Picture         =   "P_04009_A.frx":2642
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   13
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
            Picture         =   "P_04009_A.frx":2BDC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   14
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
            Picture         =   "P_04009_A.frx":3176
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   7800
         Left            =   15
         TabIndex        =   15
         Top             =   1335
         Width           =   16365
         _Version        =   524288
         _ExtentX        =   28866
         _ExtentY        =   13758
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
         MaxCols         =   27
         SpreadDesigner  =   "P_04009_A.frx":3710
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   750
         Index           =   0
         Left            =   15
         TabIndex        =   16
         Top             =   9150
         Width           =   16365
         _ExtentX        =   28866
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   0
            Left            =   60
            TabIndex        =   17
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
            Index           =   1
            Left            =   2340
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
            Caption         =   "지사매출액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   3
            Left            =   4620
            TabIndex        =   19
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
            Index           =   4
            Left            =   4620
            TabIndex        =   20
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
            Index           =   5
            Left            =   9180
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
            Caption         =   "카드 금액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   6
            Left            =   6900
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
            Caption         =   "수선 금액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   7
            Left            =   9180
            TabIndex        =   23
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
            Index           =   9
            Left            =   6900
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
            Caption         =   "수선 수량"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   10
            Left            =   11460
            TabIndex        =   25
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
            Caption         =   "반품 수량"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   11
            Left            =   11460
            TabIndex        =   26
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
            Index           =   12
            Left            =   60
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
            Caption         =   "전체 단가"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   13
            Left            =   2340
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
            Caption         =   "지사 단가"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   330
            Index           =   0
            Left            =   1200
            TabIndex        =   29
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
            TabIndex        =   30
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
            TabIndex        =   31
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
            TabIndex        =   32
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
            TabIndex        =   33
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
            Index           =   4
            Left            =   10320
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
            Index           =   3
            Left            =   5760
            TabIndex        =   36
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
            Index           =   5
            Left            =   10320
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
            Index           =   8
            Left            =   12600
            TabIndex        =   39
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
            TabIndex        =   40
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
   End
End
Attribute VB_Name = "P_04009_A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01, RS02 As ADODB.Recordset
Dim strSql As String
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String
 
Private Sub cboOffice_Click()
    Call Data_Display
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: Call Data_Display   ' 조회
        Case 1: 'Call DataAdd        ' 신규
        Case 2: 'Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
        Case 4: 'Call DataCancel     ' 취소
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

Private Sub dtInput_Change(Index As Integer)
'    Call Data_Display
End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(5).Enabled = True
    cmdBtn(6).Enabled = True
        
    '본사 여부
    If HeadOffice = MASTER_OFFICE_CODE Then
        cboOffice.Locked = False
    Else
        cboOffice.Locked = True
    End If
        
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
        
'        dtInput.Value = Format(Date, "yyyy-mm")
'
'        Call Get_지사리스트(cboInput(0))
'
'        ReDim sValue(3)
'
'        cboInput(0).ListIndex = 1
'        sValue(0) = "1"
'        sValue(1) = ""
'        sValue(2) = ""
'        sValue(3) = ""
'
'
'        Set RS01 = New ADODB.Recordset
'        Set RS01 = ExecPro("SP_04009_00_ALL", sValue(), Err_Num, Err_Dec)
'
'        spdView.MaxCols = RS01.Fields.Count
'        spdView.MaxRows = RS01.RecordCount
'
'        Call spdDisplay
''       Call fpSpread_Display(spdView, RS01)
'        Call GetColWidth(REG_App, Me.Name, spdView)
        
'        P_04009_Flag = True
'    End If
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
    
    dtInput(0).Value = Format(Date, "yyyy-mm")
    dtInput(1).Value = Format(Date, "yyyy-mm")
       
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
    
'    Call Master_tblComboAdd(cboInput(0))
'
'    ReDim sValue(3)
'
'    cboInput(0).ListIndex = 1
'    sValue(0) = "1"
'    sValue(1) = ""
'    sValue(2) = ""
'    sValue(3) = ""
'
'
'    Set RS01 = New ADODB.Recordset
'    Set RS01 = ExecPro("SP_04009_00_ALL", sValue(), Err_Num, Err_Dec)
'
'    spdView.MaxCols = RS01.Fields.Count
'    spdView.MaxRows = RS01.RecordCount
'
'    Call spdDisplay
'    Call fpSpread_Display(spdView, RS01)
'    Call GetColWidth(REG_App, Me.Name, spdView)
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_04009_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i As Integer
    
    ReDim sValue(3)
    
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    sValue(1) = ""
    sValue(2) = Format(dtInput(0).Value, "YYYY-MM-01")
    sValue(3) = Format(dtInput(1).Value, "YYYY-MM-31")
    
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_04001_01", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_04001_01", sValue(), Err_Num, Err_Dec)
    End If
    
    With spdView
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = RS01!가맹점코드 & ""               ' 1
            .Col = 2:  .Text = RS01!가맹점명 & ""                 ' 2
            .Col = 3:  .Text = RS01!영업일수 & ""                 ' 3
            .Col = 4:  .Text = RS01!지사금액 & ""                 ' 4
            .Col = 5:  .Text = RS01!가맹점금액 & ""               ' 5
            .Col = 6:  .Text = RS01!접수수량 & ""                 ' 6
            .Col = 7:  .Text = RS01!출고수량 & ""                 ' 7
            .Col = 8:  .Text = RS01!접수금액 & ""                 ' 8
            .Col = 9:  .Text = RS01!현금입금 + RS01!카드금액 & "" ' 9
            
            If RS01!접수수량 = 0 Then
                .Col = 10: .Text = 0 & ""   '10
                .Col = 11: .Text = 0 & ""   '11
                .Col = 12: .Text = 0 & ""   '12
            Else
                .Col = 10: .Text = RS01!접수금액 / RS01!접수수량 & ""   '10
                .Col = 11: .Text = RS01!지사금액 / RS01!접수수량 & ""   '11
                .Col = 12: .Text = RS01!가맹점금액 / RS01!접수수량 & "" '12
            End If
            
            .Col = 13: .Text = RS01!현금입금 & ""                 '10
            .Col = 14: .Text = RS01!카드금액 & ""                 '11
            .Col = 15: .Text = RS01!카드건수 & ""                 '12
            .Col = 16: .Text = RS01!쿠폰금액 & ""                 '13
            .Col = 17: .Text = RS01!쿠폰건수 & ""                 '14
            .Col = 18: .Text = RS01!발생마일리지 & ""             '15
            .Col = 19: .Text = RS01!사용마일리지 & ""             '16
            .Col = 20: .Text = RS01!삭제마일리지 & ""             '17
            .Col = 21: .Text = RS01!반품환불금액 & ""             '18
            .Col = 22: .Text = RS01!반품환불건수 & ""             '19
            .Col = 23: .Text = RS01!세탁환불금액 & ""             '20
            .Col = 24: .Text = RS01!세탁환불건수 & ""             '21
            .Col = 25: .Text = RS01!재세탁수량 & ""               '22
            .Col = 26: .Text = RS01!수선금액 & ""                 '23
            .Col = 27: .Text = RS01!수선수량 & ""                 '24
                        
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
        
            .Col = 2:  .Text = "합계"
            .Col = 4:  .Formula = "SUM(D1:D" & .MaxRows - 1 & ")"
            .Col = 5:  .Formula = "SUM(E1:E" & .MaxRows - 1 & ")"
            .Col = 6:  .Formula = "SUM(F1:F" & .MaxRows - 1 & ")"
            .Col = 7:  .Formula = "SUM(G1:G" & .MaxRows - 1 & ")"
            .Col = 8:  .Formula = "SUM(H1:H" & .MaxRows - 1 & ")"
            .Col = 9:  .Formula = "SUM(I1:I" & .MaxRows - 1 & ")"
            
            .Col = 10: .Formula = "SUM(J1:J" & .MaxRows - 1 & ") / " & .MaxRows - 1 & " "
            .Col = 11: .Formula = "SUM(K1:K" & .MaxRows - 1 & ") / " & .MaxRows - 1 & " "
            .Col = 12: .Formula = "SUM(L1:L" & .MaxRows - 1 & ") / " & .MaxRows - 1 & " "
            
            .Col = 13: .Formula = "SUM(M1:M" & .MaxRows - 1 & ")"
            .Col = 14: .Formula = "SUM(N1:N" & .MaxRows - 1 & ")"
            .Col = 15: .Formula = "SUM(O1:O" & .MaxRows - 1 & ")"
            .Col = 16: .Formula = "SUM(P1:P" & .MaxRows - 1 & ")"
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
            
            
            .Col = 8:  txtNum(0).Value = .Value  '전체매출액
            .Col = 10: txtNum(10).Value = .Value '전체단가
            .Col = 11: txtNum(11).Value = .Value '지사단가
            
            .Col = 4: txtNum(1).Value = .Value   '지사매출
            .Col = 5: txtNum(2).Value = .Value   '가맹점매출
            .Col = 6: txtNum(3).Value = .Value   '입고수량
            
            .Col = 26: txtNum(7).Value = .Value   '수선금액
            .Col = 27: txtNum(6).Value = .Value   '수선수량
            
            .Col = 14: txtNum(4).Value = .Value   '카드금액
            .Col = 15: txtNum(5).Value = .Value   '카드수량
            
            .Col = 21: txtNum(8).Value = .Value   '반품수량
            .Col = 25: txtNum(9).Value = .Value   '재세탁수량
        End If
        
        .Redraw = True
    End With
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

''private Sub Data_Display()
''
''    ReDim sValue(3)
''
''    Dim i As Integer
''
''    sValue(0) = "0"
''    sValue(1) = MidH(cboInput(0).Text, 2, 4)
''    sValue(2) = Format(Left(dtInput(0).Value, 7) & "-01", "YYYY-MM-DD")
''    sValue(3) = Left(dtInput(1).Value, 4) & Mid(dtInput(1).Value, 6, 2) & "31"
''
''    Set RS01 = New ADODB.Recordset
''
''    Set RS01 = ExecPro("SP_04009_00_ALL", sValue(), Err_Num, Err_Dec)
''
''    spdView.MaxRows = 0
''    spdView.MaxCols = RS01.Fields.Count  '순번
''    spdView.MaxRows = RS01.RecordCount
''
''    Call spdDisplay
''    Call GetColWidth(REG_App, Me.Name, spdView)
''
''    For i = 0 To 9
''        txtInput(i).Text = 0
''    Next i
''    With spdView
''        .Redraw = False
''        For i = 1 To RS01.RecordCount
''            .Row = i
''
''            .Col = 1:    .Text = RS01!가맹점
''            If RS01!상태 = "Y" Then
''                .Col = 2:    .Text = RS01!상태 & ":개점"
''            Else
''                .Col = 2:    .Text = RS01!상태 & ":폐점"
''            End If
''
''            .Col = 3:    .Text = RS01!택번호
''            .Col = 4:    .Text = RS01!영업일수
''            .Col = 5:    .Text = RS01!전체매출액
''
''            txtInput(0).Text = txtInput(0).Text + RS01!전체매출액
''
''            If RS01!입고수량 = 0 Then
''                .Col = 6:    .Text = Format(RS01!입고수량, "##,##0")
''                .Col = 8:    .Text = Format(RS01!입고수량, "##,##0")
''            Else
''                .Col = 6:    .Text = Format(RS01!전체매출액 / RS01!입고수량, "##,##0")
''                .Col = 8:    .Text = Format(RS01!사업장매출 / RS01!입고수량, "##,##0")
''            End If
''
''            .Col = 7:    .Text = RS01!사업장매출
''            txtInput(1).Text = txtInput(1).Text + RS01!사업장매출
''
''            .Col = 9:    .Text = RS01!가맹점매출
''            txtInput(2).Text = txtInput(2).Text + RS01!가맹점매출
''
''            .Col = 10:   .Text = RS01!입고수량
''            txtInput(3).Text = txtInput(3).Text + RS01!입고수량
''
''            .Col = 11:   .Text = RS01!카드금액
''            txtInput(4).Text = txtInput(4).Text + RS01!카드금액
''
''            .Col = 12:   .Text = RS01!카드건수
''            txtInput(5).Text = txtInput(5).Text + RS01!카드건수
''
''            .Col = 13:   .Text = RS01!재세탁수량
''            txtInput(9).Text = txtInput(9).Text + RS01!재세탁수량
''
''            .Col = 14:   .Text = RS01!수선수량
''            txtInput(6).Text = txtInput(6).Text + RS01!수선수량
''
''            .Col = 15:   .Text = RS01!수선금액
''            txtInput(7).Text = txtInput(7).Text + RS01!수선금액
''
''            .Col = 16:   .Text = RS01!반품수량
''            txtInput(8).Text = txtInput(8).Text + RS01!반품수량
''
''            .Col = 17:   .Text = RS01!출고수량
''            .Col = 18:   .Text = RS01!발생마일리지
''            .Col = 19:   .Text = RS01!사용마일리지
''            .Col = 20:   .Text = RS01!삭제마일리지
''            RS01.MoveNext
''        Next i
''        .Redraw = True
''    End With
''
''    If txtInput(3).Text = 0 Then
''        txtInput(10).Text = 0
''        txtInput(11).Text = 0
''    Else
''        txtInput(10).Text = Format(txtInput(0).Text / txtInput(3).Text, "#,##0")
''        txtInput(11).Text = Format(txtInput(1).Text / txtInput(3).Text, "#,##0")
''    End If
''
''
''    For i = 0 To 9
''        txtInput(i).Text = Format(txtInput(i).Text, "###,###,##0")
''    Next i
''
''    RS01.Close
''
''
'''    spdView.AutoCalc = True
'''
'''    spdView.MaxRows = spdView.MaxRows + 1
'''    spdView.Row = spdView.MaxRows
''
'''    spdView.RowHidden = True
''
'''    spdView.Col = 3
'''    spdView.Formula = "SUM(C1:C" & spdView.MaxRows - 1 & ")"
'''    txtInput(0).Text = spdView.Text
'''
'''    spdView.Col = 4
'''    spdView.Formula = "SUM(D1:D" & spdView.MaxRows - 1 & ")"
'''    txtInput(1).Text = spdView.Text
'''
'''    spdView.Col = 5
'''    spdView.Formula = "SUM(E1:E" & spdView.MaxRows - 1 & ")"
'''    txtInput(2).Text = spdView.Text
'''
'''    spdView.Col = 6
'''    spdView.Formula = "SUM(F1:F" & spdView.MaxRows - 1 & ")"
'''    txtInput(3).Text = spdView.Text
'''
'''    spdView.Col = 7
'''    spdView.Formula = "SUM(G1:G" & spdView.MaxRows - 1 & ")"
'''    txtInput(4).Text = spdView.Text
'''
'''    spdView.Col = 8
'''    spdView.Formula = "SUM(H1:H" & spdView.MaxRows - 1 & ")"
'''    txtInput(5).Text = spdView.Text
'''
'''    spdView.Col = 9
'''    spdView.Formula = "C" & spdView.MaxRows & " / D" & spdView.MaxRows & ""
'''    txtInput(6).Text = spdView.Text
''
''End Sub
'
'
'Public Sub Data_Display_MasterCode(sCode As String)
'    ReDim sValue(1)
'    Dim i As Integer
'
'    sValue(0) = "0"
'    sValue(1) = Format(dtInput.Value, "yyyymm")
'
'    Set RS01 = New ADODB.Recordset
'    strSql = ""
'    strSql = strSql + "SELECT "
'    strSql = strSql + "             A.AgencyName                            '대리점명', "
'    strSql = strSql + "             S.Amount                                '입금액', "
'    strSql = strSql + "             S.ISu                                   '입고수량', "
'    strSql = strSql + "             S.CSu                                   '출고수량', "
'    strSql = strSql + "             S.JSu                                   '재세탁수량', "
'    strSql = strSql + "             S.SSu                                   '수선수량', "
'    strSql = strSql + "             S.BSu                                   '반품수량', "
'    strSql = strSql + "             CASE WHEN S.Amount = 0 THEN S.Amount ELSE S.Amount / S.ISu END      '단가', "
'    strSql = strSql + "             SubString(S.STag, 1, 1) + '-' + SubString(S.STag, 2, 3)         '시작택', "
'    strSql = strSql + "             SubString(S.ETag, 1, 1) + '-' + SubString(S.ETag, 2, 3)         '종료택' "
'
'    strSql = strSql + "    FROM    SugeumMSTTotal   S (NOLOCK), "
'    strSql = strSql + "    MasterAgencyCT A(NOLOCK) "
'    strSql = strSql + "       Where A.AgencyCode = s.AgencyCode "
'    strSql = strSql + "       AND A.MasterCode = '" & sCode & "' "
'    strSql = strSql + "       AND S.MasterCode = '" & sCode & "' "
'    strSql = strSql + "       AND A.MasterCode = s.MasterCode "
'    strSql = strSql + "       AND S.SYear     =   SubString('" + sValue(1) + "', 1, 4) "
'    strSql = strSql + "       AND S.SMonth    =   SubString('" + sValue(1) + "', 5, 2) "
'    strSql = strSql + "       ORDER BY    S.Amount DESC "
'
'     Call SqlDataValue(RS01, strSql)
'
'    'Set RS01 = ExecPro("SP_04009_00", sValue(), Err_Num, Err_Dec)
'
'    spdView.MaxRows = 0
'    spdView.MaxCols = RS01.Fields.Count + 1 '순번
'    spdView.MaxRows = RS01.RecordCount
'
'    Call spdDisplay
'    Call GetColWidth(REG_App, Me.Name, spdView)
'
'    With spdView
'        .ReDraw = False
'        For i = 1 To RS01.RecordCount
'            .Row = i
'            .Col = 1:            .Text = CStr(i)
'            .Col = 2:            .Text = RS01!대리점명
'            .Col = 3:            .Text = RS01!입금액
'            .Col = 4:            .Text = RS01!입고수량
'            .Col = 5:            .Text = RS01!출고수량
'            .Col = 6:            .Text = RS01!재세탁수량
'            .Col = 7:            .Text = RS01!수선수량
'            .Col = 8:            .Text = RS01!반품수량
'            .Col = 9:            .Text = RS01!단가
'            .Col = 10:            .Text = RS01!시작택
'            .Col = 11:            .Text = RS01!종료택
'            RS01.MoveNext
'        Next i
'        .ReDraw = True
'    End With
'
'    RS01.Close
'
'
'    spdView.AutoCalc = True
'
'    spdView.MaxRows = spdView.MaxRows + 1
'    spdView.Row = spdView.MaxRows
'
''    spdView.RowHidden = True
'
'    spdView.Col = 3
'    spdView.Formula = "SUM(C1:C" & spdView.MaxRows - 1 & ")"
'    txtInput(0).Text = spdView.Text
'
'    spdView.Col = 4
'    spdView.Formula = "SUM(D1:D" & spdView.MaxRows - 1 & ")"
'    txtInput(1).Text = spdView.Text
'
'    spdView.Col = 5
'    spdView.Formula = "SUM(E1:E" & spdView.MaxRows - 1 & ")"
'    txtInput(2).Text = spdView.Text
'
'    spdView.Col = 6
'    spdView.Formula = "SUM(F1:F" & spdView.MaxRows - 1 & ")"
'    txtInput(3).Text = spdView.Text
'
'    spdView.Col = 7
'    spdView.Formula = "SUM(G1:G" & spdView.MaxRows - 1 & ")"
'    txtInput(4).Text = spdView.Text
'
'    spdView.Col = 8
'    spdView.Formula = "SUM(H1:H" & spdView.MaxRows - 1 & ")"
'    txtInput(5).Text = spdView.Text
'
'    spdView.Col = 9
'    spdView.Formula = "C" & spdView.MaxRows & " / D" & spdView.MaxRows & ""
'    txtInput(6).Text = spdView.Text
'End Sub

Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'    If NewRow <> -1 Then
'        spdView.Row = Row
'        spdView.Col = -1
'        spdView.BackColor = vbWhite
'
'        spdView.Row = NewRow
'        spdView.Col = -1
'        spdView.BackColor = glbYellow
'    End If

'    With spdView
'        If NewRow <> -1 Then
'            .Row = Row
'            If (Row Mod 2) = 0 Then
'                .Col = -1
'                .BackColor = glbGray
'            Else
'                .Col = -1
'                .BackColor = vbWhite
'            End If
'
'            .Row = NewRow
'            .Col = -1
'            .BackColor = glbYellow
'        End If
'    End With
End Sub

Public Sub DataPrint()
'    Dim ReportFP As String
'    Dim ReportFile As String
'    Dim sData As String
'
'    ReportFP = GetIniStr("REPORT", "FilePath", "", m_iniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    Call PrintDesc
'
'    Dim i As Integer
'    For i = 0 To 30
'        P_00000.crPrint.Formulas(i) = ""
'    Next
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'    P_00000.crPrint.Formulas(0) = "월 = '" & Format(dtInput(0).Value, "yyyy-mm") & "'"
'    P_00000.crPrint.Formulas(1) = "사업장 = '" & Trim(cboInput(0).Text) & "'"
'
'
'    sData = Space(15) & LeftH(" 합         계" & Space(28), 28)
'    sData = sData & RightH(Space(13) & Format(txtInput(0).Text, "#,##0"), 13)
'    sData = sData & RightH(Space(14) & Format(txtInput(1).Text, "#,##0"), 14)
'    sData = sData & RightH(Space(13) & Format(txtInput(2).Text, "#,##0"), 13)
'    sData = sData & RightH(Space(9) & Format(txtInput(3).Text, "#,##0"), 9)
'    sData = sData & RightH(Space(13) & Format(txtInput(4).Text, "#,##0"), 13)
'    sData = sData & RightH(Space(9) & Format(txtInput(5).Text, "#,##0"), 9)
'
'
'    P_00000.crPrint.Formulas(2) = "합계 = '" & sData & "'"
'
'    Set RS01 = New ADODB.Recordset
'    ReDim sValue(0)
'    sValue(0) = "0"
'    Set RS01 = ExecPro("SP_A_0000", sValue(), Err_Num, Err_Dec)
'    P_00000.crPrint.Formulas(3) = "출력시간 = '" & RS01!DB_DATE & "'"
'
'    Call ReportPrint(ReportFile, "1")
End Sub

Public Sub DataScreen()
'    Dim ReportFP As String
'    Dim ReportFile As String
'    Dim sData As String
'
'    ReportFP = GetIniStr("REPORT", "FilePath", "", m_iniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    Call PrintDesc
'
'    Dim i As Integer
'    For i = 0 To 30
'        P_00000.crPrint.Formulas(i) = ""
'    Next
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'    P_00000.crPrint.Formulas(0) = "월 = '" & Format(dtInput(0).Value, "yyyy-mm") & "'"
'    P_00000.crPrint.Formulas(1) = "사업장 = '" & Trim(cboInput(0).Text) & "'"
'
'
'    sData = Space(15) & LeftH(" 합         계" & Space(28), 28)
'    sData = sData & RightH(Space(13) & Format(txtInput(0).Text, "#,##0"), 13)
'    sData = sData & RightH(Space(14) & Format(txtInput(1).Text, "#,##0"), 14)
'    sData = sData & RightH(Space(13) & Format(txtInput(2).Text, "#,##0"), 13)
'    sData = sData & RightH(Space(9) & Format(txtInput(3).Text, "#,##0"), 9)
'    sData = sData & RightH(Space(13) & Format(txtInput(4).Text, "#,##0"), 13)
'    sData = sData & RightH(Space(9) & Format(txtInput(5).Text, "#,##0"), 9)
'
'
'    P_00000.crPrint.Formulas(2) = "합계 = '" & sData & "'"
'
'    Set RS01 = New ADODB.Recordset
'    ReDim sValue(0)
'    sValue(0) = "0"
'    Set RS01 = ExecPro("SP_A_0000", sValue(), Err_Num, Err_Dec)
'    P_00000.crPrint.Formulas(3) = "출력시간 = '" & RS01!DB_DATE & "'"
'
'    Call ReportPrint(ReportFile, "2")
End Sub

Private Sub PrintDesc()
    Dim i As Integer
    Dim TempText As String
    Dim TempFP As String
    Dim TempFile As String
    Dim FHandel As Integer
    
    FHandle = FreeFile
    
    TempFP = GetIniStr("REPORT", "TempPath", "", m_iniFile)
    TempFile = TempFP & "\Temp.txt"
    
    
    Open TempFile For Output As #FHandle
    
    TempText = ""
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        
        spdView.Col = 1
        TempText = LeftH(spdView.Text & Space(32), 32)
        spdView.Col = 3
        TempText = TempText & LeftH(spdView.Text & Space(3), 3)
        spdView.Col = 4
        TempText = TempText & RightH(Space(8) & spdView.Text, 8)
        spdView.Col = 5
        TempText = TempText & RightH(Space(14) & spdView.Text, 13)
        spdView.Col = 7
        TempText = TempText & RightH(Space(14) & spdView.Text, 14)
        spdView.Col = 9
        TempText = TempText & RightH(Space(13) & spdView.Text, 13)
        spdView.Col = 10
        TempText = TempText & RightH(Space(9) & spdView.Text, 9)
        spdView.Col = 11
        TempText = TempText & RightH(Space(13) & spdView.Text, 13)
        spdView.Col = 12
        TempText = TempText & RightH(Space(9) & spdView.Text, 9)
        
        Print #FHandle, TempText
    Next i
    
    Close #FHandle
End Sub
