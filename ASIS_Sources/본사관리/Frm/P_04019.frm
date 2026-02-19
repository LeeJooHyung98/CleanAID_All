VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_04019 
   Caption         =   "가맹점 기간별 매출현황 (합계)"
   ClientHeight    =   10425
   ClientLeft      =   6015
   ClientTop       =   4440
   ClientWidth     =   16425
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_04019.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10425
   ScaleWidth      =   16425
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   10425
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16425
      _ExtentX        =   28972
      _ExtentY        =   18389
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_04019.frx":058A
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   1
         Top             =   540
         Width           =   16395
         _ExtentX        =   28919
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1245
            TabIndex        =   2
            Text            =   "cboOffice"
            Top             =   60
            Width           =   2880
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   4425
            TabIndex        =   3
            Top             =   420
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   556
            _Version        =   393216
            Format          =   56885248
            CurrentDate     =   39826
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1245
            TabIndex        =   4
            Top             =   420
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            _Version        =   393216
            Format          =   56885248
            CurrentDate     =   37140
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   4140
            TabIndex        =   5
            Top             =   420
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
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
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   8
            Left            =   60
            TabIndex        =   6
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
            Height          =   315
            Index           =   14
            Left            =   60
            TabIndex        =   7
            Top             =   420
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "기    간"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox CheckBox1 
            Height          =   225
            Left            =   7530
            TabIndex        =   44
            Top             =   450
            Width           =   3135
            _Version        =   851970
            _ExtentX        =   5530
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "지사출고 수량 집계"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   8
         Top             =   15
         Width           =   8790
         _ExtentX        =   15505
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
         PictureBackground=   "P_04019.frx":063C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   1
         Left            =   8820
         TabIndex        =   9
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
         PictureBackground=   "P_04019.frx":083E
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   10
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
            Picture         =   "P_04019.frx":0A40
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   11
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
            Picture         =   "P_04019.frx":0FDA
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   12
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
            Picture         =   "P_04019.frx":1574
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   13
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
            Picture         =   "P_04019.frx":1B0E
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   14
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
            Picture         =   "P_04019.frx":20A8
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   15
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
            Picture         =   "P_04019.frx":2642
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   16
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
            Picture         =   "P_04019.frx":2BDC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   17
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
            Picture         =   "P_04019.frx":3176
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   8310
         Left            =   15
         TabIndex        =   18
         Top             =   1335
         Width           =   16395
         _Version        =   524288
         _ExtentX        =   28919
         _ExtentY        =   14658
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
         SpreadDesigner  =   "P_04019.frx":3710
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   750
         Index           =   0
         Left            =   15
         TabIndex        =   19
         Top             =   9660
         Width           =   16395
         _ExtentX        =   28919
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   0
            Left            =   60
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
            Caption         =   "전체매출액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   1
            Left            =   2340
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
            Caption         =   "지사매출액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   3
            Left            =   4620
            TabIndex        =   22
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
            Caption         =   "가맹점매출액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   5
            Left            =   9180
            TabIndex        =   24
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
            Caption         =   "수선 금액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   7
            Left            =   9180
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
            Caption         =   "카드 건수"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   9
            Left            =   6900
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
            Caption         =   "수선 수량"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   10
            Left            =   11460
            TabIndex        =   28
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
            Caption         =   "재세탁수량"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   12
            Left            =   60
            TabIndex        =   30
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
            TabIndex        =   31
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
            Index           =   10
            Left            =   1200
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
            Index           =   1
            Left            =   3480
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
            Index           =   11
            Left            =   3480
            TabIndex        =   35
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
            Index           =   7
            Left            =   8040
            TabIndex        =   37
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
            TabIndex        =   38
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
            Index           =   6
            Left            =   8040
            TabIndex        =   40
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
            TabIndex        =   41
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
            TabIndex        =   42
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
            TabIndex        =   43
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
Attribute VB_Name = "P_04019"
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

Private Sub cboInput_Change(Index As Integer)
'    Select Case Index
'        Case 0
'            Call Data_Display
'    End Select
End Sub

Private Sub cboOffice_Click()
    Call Data_Display
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
        Case 0: Call Data_Display   ' 조회
        Case 1: 'Call DataAdd        ' 신규
        Case 2: 'Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
        Case 4: 'Call DataCancel     ' 취소
        Case 5: Call DataPrint      ' 인쇄
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

'    spdView.Row = 0
'    spdView.Col = 1:    spdView.Text = "가맹점"
'    spdView.Col = 2:    spdView.Text = "상태"
'    spdView.Col = 3:    spdView.Text = "택번호"
'    spdView.Col = 4:    spdView.Text = "영업일수"
'    spdView.Col = 5:    spdView.Text = "전체매출액"
'    spdView.Col = 6:    spdView.Text = "매출단가"
'    spdView.Col = 7:    spdView.Text = "사업장매출"
'    spdView.Col = 8:    spdView.Text = "사업장단가"
'    spdView.Col = 9:    spdView.Text = "가맹점매출"
'    spdView.Col = 10:   spdView.Text = "입고수량"
'    spdView.Col = 11:   spdView.Text = "카드금액"
'    spdView.Col = 12:   spdView.Text = "카드건수"
'    spdView.Col = 13:   spdView.Text = "재세탁수량"
'    spdView.Col = 14:   spdView.Text = "수선수량"
'    spdView.Col = 15:   spdView.Text = "수선금액"
'    spdView.Col = 16:   spdView.Text = "반품수량"
'    spdView.Col = 17:   spdView.Text = "출고수량"
'    spdView.Col = 18:   spdView.Text = "발생마일리지"
'    spdView.Col = 19:   spdView.Text = "사용마일리지"
'    spdView.Col = 20:   spdView.Text = "삭제마일리지"


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
'        .OperationMode = OperationModeSingle
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
    
    dtInput(0).Value = Format(Date, "YYYY-MM-DD")
    dtInput(1).Value = Format(Date, "YYYY-MM-DD")
       
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
    
'     Call Master_tblComboAdd(cboInput(0))
'
'     ReDim sValue(3)
'
'     cboInput(0).ListIndex = 1
'     sValue(0) = "1"
'     sValue(1) = ""
'     sValue(2) = ""
'     sValue(3) = ""
'
'
'     Set RS01 = New ADODB.Recordset
'     Set RS01 = ExecPro("SP_04009_00_ALL", sValue(), Err_Num, Err_Dec)
'
'     spdView.MaxCols = RS01.Fields.Count
'     spdView.MaxRows = RS01.RecordCount
'
'     Call spdDisplay
''       Call fpSpread_Display(spdView, RS01)
'     Call GetColWidth(REG_App, Me.Name, spdView)
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

'Private Sub Form_Load()
'    dtInput.Value = Format(Date, "yyyy-mm")
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_04019_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i As Integer
    Dim nRow    As Long
    Dim vText   As Variant
    
    ReDim sValue(3)
    
    sValue(0) = IIf(Mid(cboOffice.Text, 2, 4) = "0000", "%", Mid(cboOffice.Text, 2, 4))
    sValue(1) = ""
    sValue(2) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(3) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(HeadOffice) = False Then Exit Sub
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_04019_01", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_04019_01", sValue(), Err_Num, Err_Dec)
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
            .Col = 7:  .Text = "0" ' RS01!출고수량 & ""                 ' 7
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
        
        ReDim sValue(4)
        
        sValue(0) = Mid(cboOffice.Text, 2, 4)
        If sValue(0) = "0000" Then sValue(0) = "%"

        sValue(1) = "%"
        sValue(2) = Format(dtInput(0).Value, "YYYY-MM-DD")
        sValue(3) = Format(dtInput(1).Value, "YYYY-MM-DD")
        sValue(4) = "STORE_SUM"
        
        If CheckBox1.Value = xtpChecked Then
            
            If HeadOffice = MASTER_OFFICE_CODE Then
                If DBOpen_Master(HeadOffice) = False Then Exit Sub
                
                Set RS01 = New ADODB.Recordset
                Set RS01 = ExecProMaster("SP_04001_B_01", sValue(), Err_Num, Err_Dec)
            Else
                Set RS01 = New ADODB.Recordset
                Set RS01 = ExecPro("SP_04001_B_01", sValue(), Err_Num, Err_Dec)
            End If
        
        
            Do While Not RS01.EOF
                ' 지사출고 수량 출력
                For nRow = 1 To .MaxRows
                    .GetText 1, nRow, vText
                    If CStr(vText) = RS01.Fields(1) Then
                        .SetText 7, nRow, CVar(RS01.Fields(2))
                        Exit For
                    End If
                Next nRow
                
                RS01.MoveNext
            Loop
            RS01.Close
            Set RS01 = Nothing
        End If
        
        
        
        If .MaxRows > 0 Then

            ' 합계 출력
            Dim nCol    As Long
            Dim dblCnt(4)   As Double
            For nCol = 4 To .MaxCols
                Select Case nCol
                    Case 4: dblCnt(2) = SpreadSum(spdView, 2, nCol)
                    Case 5: dblCnt(3) = SpreadSum(spdView, -1, nCol)
                    Case 8: dblCnt(1) = SpreadSum(spdView, -1, nCol)
                    Case 6:  dblCnt(0) = SpreadSum(spdView, -1, nCol)
                    Case 10: .SetText nCol, .MaxRows, CVar(dblCnt(1) / dblCnt(0))
                    Case 11: .SetText nCol, .MaxRows, CVar(dblCnt(2) / dblCnt(0))
                    Case 12: .SetText nCol, .MaxRows, CVar(dblCnt(3) / dblCnt(0))
                    Case Else: Call SpreadSum(spdView, -1, nCol)
                End Select
            Next nCol
    
'            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Row = .Row
            .Row2 = .Row
            .Col = 1
            .Col2 = .MaxCols
            .BlockMode = True
            .BackColor = &HC0FFC0
            .BlockMode = False
        
'            .Col = 3:  .Text = "합계"
'            .Col = 4:  .Formula = "SUM(D1:D" & .MaxRows - 1 & ")"
'            .Col = 5:  .Formula = "SUM(E1:E" & .MaxRows - 1 & ")"
'            .Col = 6:  .Formula = "SUM(F1:F" & .MaxRows - 1 & ")"
'            .Col = 7:  .Formula = "SUM(G1:G" & .MaxRows - 1 & ")"
'            .Col = 8:  .Formula = "SUM(H1:H" & .MaxRows - 1 & ")"
'            .Col = 9:  .Formula = "SUM(I1:I" & .MaxRows - 1 & ")"
'
'            .Col = 10: .Formula = "SUM(H1:H" & .MaxRows - 1 & ") /  " & .MaxRows - 1
'            .Col = 11: .Formula = "SUM(D1:D" & .MaxRows - 1 & ") /  " & .MaxRows - 1
'            .Col = 12: .Formula = "SUM(E1:E" & .MaxRows - 1 & ") /  " & .MaxRows - 1
'
'            .Col = 13: .Formula = "SUM(M1:M" & .MaxRows - 1 & ")"
'            .Col = 14: .Formula = "SUM(N1:N" & .MaxRows - 1 & ")"
'            .Col = 15: .Formula = "SUM(O1:O" & .MaxRows - 1 & ")"
'            .Col = 16: .Formula = "SUM(P1:P" & .MaxRows - 1 & ")"
'            .Col = 17: .Formula = "SUM(Q1:Q" & .MaxRows - 1 & ")"
'            .Col = 18: .Formula = "SUM(R1:R" & .MaxRows - 1 & ")"
'            .Col = 19: .Formula = "SUM(S1:S" & .MaxRows - 1 & ")"
'            .Col = 20: .Formula = "SUM(T1:T" & .MaxRows - 1 & ")"
'            .Col = 21: .Formula = "SUM(U1:U" & .MaxRows - 1 & ")"
'            .Col = 22: .Formula = "SUM(V1:V" & .MaxRows - 1 & ")"
'            .Col = 23: .Formula = "SUM(W1:W" & .MaxRows - 1 & ")"
'            .Col = 24: .Formula = "SUM(X1:X" & .MaxRows - 1 & ")"
'
'            .Col = 25: .Formula = "SUM(Y1:Y" & .MaxRows - 1 & ")"
'            .Col = 26: .Formula = "SUM(Z1:Z" & .MaxRows - 1 & ")"
'            .Col = 27: .Formula = "SUM(AA1:AA" & .MaxRows - 1 & ")"
            
            
            .Col = 8:  txtNum(0).Value = .Value  '전체매출액
            .Col = 10: txtNum(10).Value = Val(.Value) '전체단가
            .Col = 11: txtNum(11).Value = Val(.Value) '지사단가
            
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

'private Sub Data_Display()
'    ReDim sValue(3)
'
'    Dim i As Integer
'
'    If dtInput.Value > dtInput1.Value Then
'        MsgBox "기간을 확인 하세요", vbInformation, "오류"
'        dtInput1.SetFocus
'        Exit Sub
'    End If
'
'    sValue(0) = "0"
'    sValue(1) = MidH(cboInput(0).Text, 2, 4)
'    sValue(2) = Format(dtInput.Value, "YYYY-MM-DD")
'    sValue(3) = Format(dtInput1.Value, "YYYY-MM-DD")  'Left(dtInput.Value, 4) & Mid(dtInput.Value, 6, 2) & "31"
'
'
'    Set RS01 = New ADODB.Recordset
'
'    Set RS01 = ExecPro("SP_04009_00_ALL", sValue(), Err_Num, Err_Dec)
'
'    spdView.MaxRows = 0
'    spdView.MaxCols = RS01.Fields.Count  '순번
'    spdView.MaxRows = RS01.RecordCount
'
'    Call spdDisplay
'    Call GetColWidth(REG_App, Me.Name, spdView)
'
'    For i = 0 To 9
'        txtInput(i).Text = 0
'    Next i
'    With spdView
'        .Redraw = False
'        For i = 1 To RS01.RecordCount
'            .Row = i
'
'            .Col = 1:    .Text = RS01!가맹점
'            If RS01!상태 = "Y" Then
'                .Col = 2:    .Text = RS01!상태 & ":개점"
'            Else
'                .Col = 2:    .Text = RS01!상태 & ":폐점"
'            End If
'            .Col = 3:    .Text = RS01!택번호
'            .Col = 4:    .Text = RS01!영업일수
'            .Col = 5:    .Text = RS01!전체매출액
'            txtInput(0).Text = txtInput(0).Text + RS01!전체매출액
'            If RS01!입고수량 = 0 Then
'                .Col = 6:    .Text = Format(RS01!입고수량, "##,##0")
'                .Col = 8:    .Text = Format(RS01!입고수량, "##,##0")
'            Else
'                .Col = 6:    .Text = Format(RS01!전체매출액 / RS01!입고수량, "##,##0")
'                .Col = 8:    .Text = Format(RS01!사업장매출 / RS01!입고수량, "##,##0")
'            End If
'            .Col = 7:    .Text = RS01!사업장매출
'            txtInput(1).Text = txtInput(1).Text + RS01!사업장매출
'            .Col = 9:    .Text = RS01!가맹점매출
'            txtInput(2).Text = txtInput(2).Text + RS01!가맹점매출
'            .Col = 10:   .Text = RS01!입고수량
'            txtInput(3).Text = txtInput(3).Text + RS01!입고수량
'            .Col = 11:   .Text = RS01!카드금액
'            txtInput(4).Text = txtInput(4).Text + RS01!카드금액
'            .Col = 12:   .Text = RS01!카드건수
'            txtInput(5).Text = txtInput(5).Text + RS01!카드건수
'            .Col = 13:   .Text = RS01!재세탁수량
'            txtInput(9).Text = txtInput(9).Text + RS01!재세탁수량
'            .Col = 14:   .Text = RS01!수선수량
'            txtInput(6).Text = txtInput(6).Text + RS01!수선수량
'            .Col = 15:   .Text = RS01!수선금액
'            txtInput(7).Text = txtInput(7).Text + RS01!수선금액
'            .Col = 16:   .Text = RS01!반품수량
'            txtInput(8).Text = txtInput(8).Text + RS01!반품수량
'            .Col = 17:   .Text = RS01!출고수량
'            .Col = 18:   .Text = RS01!발생마일리지
'            .Col = 19:   .Text = RS01!사용마일리지
'            .Col = 20:   .Text = RS01!삭제마일리지
'            RS01.MoveNext
'        Next i
'        .Redraw = True
'    End With
'
'    If txtInput(3).Text = 0 Then
'        txtInput(10).Text = 0
'        txtInput(11).Text = 0
'    Else
'        txtInput(10).Text = Format(txtInput(0).Text / txtInput(3).Text, "#,##0")
'        txtInput(11).Text = Format(txtInput(1).Text / txtInput(3).Text, "#,##0")
'    End If
'
'
'    For i = 0 To 9
'        txtInput(i).Text = Format(txtInput(i).Text, "###,###,##0")
'    Next i
'
'    RS01.Close
'
'
''    spdView.AutoCalc = True
''
''    spdView.MaxRows = spdView.MaxRows + 1
''    spdView.Row = spdView.MaxRows
'
''    spdView.RowHidden = True
'
''    spdView.Col = 3
''    spdView.Formula = "SUM(C1:C" & spdView.MaxRows - 1 & ")"
''    txtInput(0).Text = spdView.Text
''
''    spdView.Col = 4
''    spdView.Formula = "SUM(D1:D" & spdView.MaxRows - 1 & ")"
''    txtInput(1).Text = spdView.Text
''
''    spdView.Col = 5
''    spdView.Formula = "SUM(E1:E" & spdView.MaxRows - 1 & ")"
''    txtInput(2).Text = spdView.Text
''
''    spdView.Col = 6
''    spdView.Formula = "SUM(F1:F" & spdView.MaxRows - 1 & ")"
''    txtInput(3).Text = spdView.Text
''
''    spdView.Col = 7
''    spdView.Formula = "SUM(G1:G" & spdView.MaxRows - 1 & ")"
''    txtInput(4).Text = spdView.Text
''
''    spdView.Col = 8
''    spdView.Formula = "SUM(H1:H" & spdView.MaxRows - 1 & ")"
''    txtInput(5).Text = spdView.Text
''
''    spdView.Col = 9
''    spdView.Formula = "C" & spdView.MaxRows & " / D" & spdView.MaxRows & ""
''    txtInput(6).Text = spdView.Text
'End Sub
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

'Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
''    If NewRow <> -1 Then
''        spdView.Row = Row
''        spdView.Col = -1
''        spdView.BackColor = vbWhite
''
''        spdView.Row = NewRow
''        spdView.Col = -1
''        spdView.BackColor = glbYellow
''    End If
'
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
'End Sub


Private Sub DataPrint()
    On Error GoTo ErrRtn
    Dim 지사명      As String
    Dim 지사코드    As String
    
    Dim 택번호      As String
    Dim XML         As String
    Dim i           As Integer
    Dim FileNumber
        
    If spdView.DataRowCnt <= 0 Then Exit Sub
    
        
    지사코드 = Mid(cboOffice.Text, 2, 4)
    지사명 = Mid(cboOffice.Text, 7)

    FileNumber = FreeFile
    Open App.Path & "\XML\P_04019.XML" For Output As #FileNumber
    
    Print #FileNumber, "<?xml version=""1.0"" encoding=""EUC-KR""?>"
    Print #FileNumber, "<root>"
    
          XML = "    <HEADERDATA>" & vbLf
    XML = XML & "        <타이틀>" & Func_Replace(지사명) & " (" & 지사코드 & ") 가맹점 기간별 매출현황(합계)</타이틀>" & vbLf
    XML = XML & "        <지사>" & "검색 기간 : " & Format(dtInput(0).Value, "yyyy-MM-dd") & " ~ " & Format(dtInput(1).Value, "yyyy-MM-dd") & "</지사>" & vbLf
    XML = XML & "   </HEADERDATA>" & vbLf
    Print #FileNumber, XML
    
    With spdView
        
        For i = 1 To .DataRowCnt
            .Row = i
            XML = "    <DATA>" & vbLf
            .Col = 2: XML = XML & "        <가맹점명>" & Trim(.Text) & "</가맹점명>" & vbLf
            .Col = 3: XML = XML & "        <일수>" & Trim(.Text) & "</일수>" & vbLf
            .Col = 4: XML = XML & "        <지사>" & Trim(.Text) & "</지사>" & vbLf
            .Col = 5: XML = XML & "        <가맹점>" & Trim(.Text) & "</가맹점>" & vbLf
            .Col = 6: XML = XML & "        <접수수량>" & Trim(.Text) & "</접수수량>" & vbLf
            .Col = 7: XML = XML & "        <출고수량>" & Trim(.Text) & "</출고수량>" & vbLf
            .Col = 8: XML = XML & "        <매출액>" & Trim(.Text) & "</매출액>" & vbLf
            .Col = 9: XML = XML & "        <입금액>" & Trim(.Text) & "</입금액>" & vbLf
            
            .Col = 13: XML = XML & "       <현금>" & Trim(.Text) & "</현금>" & vbLf
            .Col = 14: XML = XML & "       <카드매출액>" & Trim(.Text) & "</카드매출액>" & vbLf
            .Col = 15: XML = XML & "       <카드건수>" & Trim(.Text) & "</카드건수>" & vbLf
            .Col = 19: XML = XML & "       <사용>" & Trim(.Text) & "</사용>" & vbLf
            .Col = 21: XML = XML & "       <반품금액>" & Trim(.Text) & "</반품금액>" & vbLf
            .Col = 22: XML = XML & "       <반품건수>" & Trim(.Text) & "</반품건수>" & vbLf
            .Col = 23: XML = XML & "       <세탁금액>" & Trim(.Text) & "</세탁금액>" & vbLf
            .Col = 24: XML = XML & "       <세탁건수>" & Trim(.Text) & "</세탁건수>" & vbLf
            XML = XML & "   </DATA>" & vbLf
            Print #FileNumber, XML
        Next i
        
        Print #FileNumber, "</root>" & vbLf
        Close #FileNumber
    End With
    
    With rpt_P_04019
        .documentName = "가맹점 기간별 매출현황(합계)"
        .dc.FileURL = App.Path & "\XML\P_04019.XML"
        .PrintReport False
        
        '.Show 1
    End With

    Unload rpt월간사업장매출현황
    
    Exit Sub

ErrRtn:
    MsgBox Err.Description, vbInformation, "오류"
    Screen.MousePointer = 0
End Sub


Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If NewRow <> -1 Then
        With spdView
            If NewRow <> -1 Then
                .Row = Row
                If (Row Mod 2) = 0 Then
                    .Col = -1
                    .BackColor = vbWhite
                Else
                    .Col = -1
                    .BackColor = vbWhite
                End If
                
                .Row = NewRow
                .Col = -1
                .BackColor = glbYellow
            End If
        End With
    End If

End Sub

