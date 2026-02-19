VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_04009_R1 
   Caption         =   "[전사업장]기간별 불량세탁 환불현황"
   ClientHeight    =   9645
   ClientLeft      =   7740
   ClientTop       =   3990
   ClientWidth     =   16350
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_04009_R1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9645
   ScaleWidth      =   16350
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   9645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16350
      _ExtentX        =   28840
      _ExtentY        =   17013
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_04009_R1.frx":058A
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   1
         Top             =   540
         Width           =   16320
         _ExtentX        =   28787
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1245
            Style           =   2  '드롭다운 목록
            TabIndex        =   2
            Top             =   60
            Width           =   3285
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1245
            TabIndex        =   3
            Top             =   405
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            Format          =   21430273
            CurrentDate     =   37140
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   4
            Top             =   405
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "조회년월"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   8
            Left            =   60
            TabIndex        =   5
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "지 사 명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   3015
            TabIndex        =   6
            Top             =   405
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy-MM"
            Format          =   21430273
            CurrentDate     =   37140
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   14
            Left            =   2760
            TabIndex        =   7
            Top             =   390
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   11.25
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
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   8
         Top             =   15
         Width           =   8715
         _ExtentX        =   15372
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
         PictureBackground=   "P_04009_R1.frx":063C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   1
         Left            =   8745
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
         PictureBackground=   "P_04009_R1.frx":083E
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
            Picture         =   "P_04009_R1.frx":0A40
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
            Picture         =   "P_04009_R1.frx":0FDA
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
            Picture         =   "P_04009_R1.frx":1574
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
            Picture         =   "P_04009_R1.frx":1B0E
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
            Picture         =   "P_04009_R1.frx":20A8
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
            Picture         =   "P_04009_R1.frx":2642
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
            Picture         =   "P_04009_R1.frx":2BDC
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
            Picture         =   "P_04009_R1.frx":3176
         End
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   750
         Index           =   0
         Left            =   15
         TabIndex        =   18
         Top             =   8880
         Width           =   16320
         _ExtentX        =   28787
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   0
            Left            =   45
            TabIndex        =   19
            Top             =   45
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   262144
            BackColor       =   12648384
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
            Left            =   2640
            TabIndex        =   20
            Top             =   45
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   262144
            BackColor       =   12648384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "지사 매출액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   3
            Left            =   7830
            TabIndex        =   21
            Top             =   45
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   262144
            BackColor       =   12648384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "접수 수량"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   4
            Left            =   5235
            TabIndex        =   22
            Top             =   45
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   262144
            BackColor       =   12648384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "가맹점 매출액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   5
            Left            =   5235
            TabIndex        =   23
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   262144
            BackColor       =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "환불 금액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   7
            Left            =   7830
            TabIndex        =   24
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   262144
            BackColor       =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "환불  건수"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   12
            Left            =   45
            TabIndex        =   25
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   262144
            BackColor       =   12648384
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
            Left            =   2640
            TabIndex        =   26
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   262144
            BackColor       =   12648384
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
            Left            =   1365
            TabIndex        =   27
            Top             =   45
            Width           =   1290
            _Version        =   262145
            _ExtentX        =   2275
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
            Left            =   3960
            TabIndex        =   28
            Top             =   45
            Width           =   1290
            _Version        =   262145
            _ExtentX        =   2275
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
            Left            =   6555
            TabIndex        =   29
            Top             =   45
            Width           =   1290
            _Version        =   262145
            _ExtentX        =   2275
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
            Left            =   9150
            TabIndex        =   30
            Top             =   45
            Width           =   1290
            _Version        =   262145
            _ExtentX        =   2275
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
            Left            =   1365
            TabIndex        =   31
            Top             =   360
            Width           =   1290
            _Version        =   262145
            _ExtentX        =   2275
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
            Left            =   3960
            TabIndex        =   32
            Top             =   360
            Width           =   1290
            _Version        =   262145
            _ExtentX        =   2275
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
            Left            =   6555
            TabIndex        =   33
            Top             =   360
            Width           =   1290
            _Version        =   262145
            _ExtentX        =   2275
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
            Left            =   9150
            TabIndex        =   34
            Top             =   360
            Width           =   1290
            _Version        =   262145
            _ExtentX        =   2275
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
         Height          =   7530
         Left            =   15
         TabIndex        =   35
         Top             =   1335
         Width           =   16320
         _Version        =   524288
         _ExtentX        =   28787
         _ExtentY        =   13282
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
         MaxCols         =   13
         SpreadDesigner  =   "P_04009_R1.frx":3710
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_04009_R1"
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

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(5).Enabled = True
    cmdBtn(6).Enabled = True
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
        
    '본사 여부
    If HeadOffice = MASTER_OFFICE_CODE Then
        cboOffice.Locked = False
    Else
        cboOffice.Locked = True
    End If
        
        
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

'    spdView.MaxRows = 24
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
'    spdView.Col = 11:   spdView.Text = "환불수량"
'    spdView.Col = 12:   spdView.Text = "환불금액"

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
    
    dtInput(0).Value = Format(Date, "YYYY-MM-01")
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
'    'Set RS01 = New ADODB.Recordset
'    'Set RS01 = ExecPro("SP_04009_R0", sValue(), Err_Num, Err_Dec)
'
'    'spdView.MaxCols = RS01.Fields.Count
'    'spdView.MaxRows = RS01.RecordCount
'
'    Call spdDisplay
'    '       Call fpSpread_Display(spdView, RS01)
'    Call GetColWidth(REG_App, Me.Name, spdView)
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

'Private Sub Form_Load()
'    dtInput.Value = Format(Date, "yyyy-mm")
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_04009_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i As Integer
    
    For i = 0 To 7
        txtNum(i).Value = 0
    Next i
    
    ReDim sValue(3)
    
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    sValue(1) = ""
    sValue(2) = Format(dtInput(0).Value, "YYYY-MM-01")
    sValue(3) = Format(dtInput(1).Value, "YYYY-MM-31")
    
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_04009_R01", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_04009_R01", sValue(), Err_Num, Err_Dec)
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
            
            '.Col = 6:  .Text = RS01!접수수량 & ""                 ' 6
            '.Col = 7:  .Text = RS01!출고수량 & ""                 ' 7
            
            .Col = 6:  .Text = RS01!접수금액 & ""                 ' 8
            .Col = 7:  .Text = RS01!현금입금 + RS01!카드입금 & "" ' 9
            
            If RS01!접수수량 = 0 Then
                .Col = 8:  .Text = 0 & ""   '10
                .Col = 9:  .Text = 0 & ""   '11
                .Col = 10: .Text = 0 & ""   '12
            Else
                .Col = 8:  .Text = RS01!접수금액 / RS01!접수수량 & ""   '10
                .Col = 9:  .Text = RS01!지사금액 / RS01!접수수량 & ""   '11
                .Col = 10: .Text = RS01!가맹점금액 / RS01!접수수량 & "" '12
            End If
            
            .Col = 11: .Text = RS01!세탁환불건수 & ""                 '10
            .Col = 12: .Text = RS01!세탁환불금액 & ""                 '11
            .Col = 13: .Text = ""                  '12
                        
            txtNum(0).Value = txtNum(0).Value + RS01!접수금액
            txtNum(1).Value = txtNum(1).Value + RS01!지사금액
            txtNum(2).Value = txtNum(2).Value + RS01!가맹점금액
            
            txtNum(3).Value = txtNum(3).Value + RS01!접수수량
                        
            txtNum(6).Value = txtNum(6).Value + RS01!세탁환불금액
            txtNum(7).Value = txtNum(6).Value + RS01!세탁환불건수
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        If .MaxRows > 0 Then
            ' 합계 출력
            Dim nCol    As Long
            Dim dblCnt(4)   As Double
            For nCol = 4 To .MaxCols
                Select Case nCol
                    Case 4:  Call SpreadSum(spdView, 2, nCol)
                    Case Else: Call SpreadSum(spdView, -1, nCol)
                End Select
            Next nCol
        End If
'        If .MaxRows > 0 Then
'            .MaxRows = .MaxRows + 1
'            .Row = .MaxRows
'
'            .Row = .Row
'            .Row2 = .Row
'            .Col = 1
'            .Col2 = .MaxCols
'            .BlockMode = True
'            .BackColor = &HC0FFC0
'            .BlockMode = False
'
'            .Col = 1:  .Text = "합계"
'            .Col = 4:  .Formula = "SUM(D1:D" & .MaxRows - 1 & ")"
'            .Col = 5:  .Formula = "SUM(E1:E" & .MaxRows - 1 & ")"
'            .Col = 6:  .Formula = "SUM(F1:F" & .MaxRows - 1 & ")"
'            .Col = 7:  .Formula = "SUM(G1:G" & .MaxRows - 1 & ")"
'
'            '.Col = 8:  .Formula = "SUM(H1:H" & .MaxRows - 1 & ")"
'            '.Col = 9:  .Formula = "SUM(I1:I" & .MaxRows - 1 & ")"
'            '.Col = 10: .Formula = "SUM(J1:J" & .MaxRows - 1 & ") / SUM(F1:F" & .MaxRows - 1 & ")"
'
'            .Col = 11: .Formula = "SUM(K1:K" & .MaxRows - 1 & ")"
'            .Col = 12: .Formula = "SUM(L1:L" & .MaxRows - 1 & ")"
'        End If
'
        .Redraw = True
    End With
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

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

    With spdView
        If NewRow <> -1 Then
            .Row = Row
            If (Row Mod 2) = 0 Then
                .Col = -1
                .BackColor = glbGray
            Else
                .Col = -1
                .BackColor = vbWhite
            End If
            
            .Row = NewRow
            .Col = -1
            .BackColor = glbYellow
        End If
    End With
End Sub

Private Sub spdView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        'PopupMenu P_00000.PopUp
    End If
End Sub

Private Sub DataPrint()
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

Private Sub DataScreen()
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
