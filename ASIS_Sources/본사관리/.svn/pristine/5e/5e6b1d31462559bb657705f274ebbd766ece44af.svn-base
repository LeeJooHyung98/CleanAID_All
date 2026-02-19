VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_06012 
   Caption         =   "사고 처리 접수 (CS팀)"
   ClientHeight    =   10260
   ClientLeft      =   2220
   ClientTop       =   4935
   ClientWidth     =   15270
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_06012.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10260
   ScaleWidth      =   15270
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   10260
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   18098
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_06012.frx":058A
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   1
         Top             =   540
         Width           =   15240
         _ExtentX        =   26882
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   0
            Left            =   8925
            Style           =   2  '드롭다운 목록
            TabIndex        =   2
            Top             =   60
            Width           =   5355
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   6
            Left            =   1800
            TabIndex        =   3
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Format          =   64290816
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   675
            Index           =   0
            Left            =   60
            TabIndex        =   4
            Top             =   60
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   1191
            _Version        =   262144
            Caption         =   "사고접수일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   19
            Left            =   5715
            TabIndex        =   5
            Top             =   60
            Width           =   3180
            _ExtentX        =   5609
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "사고접수일자/접수번호/매장명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   390
            Index           =   8
            Left            =   5700
            TabIndex        =   70
            Top             =   390
            Width           =   3210
            _Version        =   851970
            _ExtentX        =   5662
            _ExtentY        =   688
            _StockProps     =   79
            Caption         =   "매장 찾기"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_06012.frx":061C
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   7
            Left            =   1800
            TabIndex        =   71
            Top             =   420
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Format          =   64290816
            CurrentDate     =   36686
         End
         Begin XtremeSuiteControls.PushButton cmdRefresh 
            Height          =   330
            Left            =   4950
            TabIndex        =   74
            Top             =   60
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   582
            _StockProps     =   79
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_06012.frx":0BB6
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   6
         Top             =   15
         Width           =   7620
         _ExtentX        =   13441
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
         Caption         =   " 사고 처리 접수  (CS팀)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_06012.frx":1150
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   0
         Left            =   7650
         TabIndex        =   7
         Top             =   15
         Width           =   7605
         _ExtentX        =   13414
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
         PictureBackground=   "P_06012.frx":1352
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   8
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
            Picture         =   "P_06012.frx":1554
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   9
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "화면"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_06012.frx":1AEE
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
            Appearance      =   6
            Picture         =   "P_06012.frx":2088
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
            Appearance      =   6
            Picture         =   "P_06012.frx":2622
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
            Appearance      =   6
            Picture         =   "P_06012.frx":2BBC
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
            Appearance      =   6
            Picture         =   "P_06012.frx":3156
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
            Appearance      =   6
            Picture         =   "P_06012.frx":36F0
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
            Appearance      =   6
            Picture         =   "P_06012.frx":3C8A
         End
      End
      Begin Threed.SSPanel panDetail 
         Height          =   8910
         Left            =   15
         TabIndex        =   16
         Top             =   1335
         Width           =   15240
         _ExtentX        =   26882
         _ExtentY        =   15716
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin XtremeSuiteControls.GroupBox GroupBox 
            Height          =   3600
            Index           =   4
            Left            =   120
            TabIndex        =   53
            Top             =   5070
            Width           =   15060
            _Version        =   851970
            _ExtentX        =   26564
            _ExtentY        =   6350
            _StockProps     =   79
            Caption         =   "※ 보상 산정 기준"
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
            Appearance      =   6
            BorderStyle     =   1
            Begin VB.TextBox txtInput 
               Height          =   1635
               Index           =   9
               Left            =   1560
               MultiLine       =   -1  'True
               TabIndex        =   80
               Top             =   1770
               Width           =   12735
            End
            Begin VB.ComboBox cboInput 
               Height          =   315
               Index           =   5
               Left            =   5160
               Style           =   2  '드롭다운 목록
               TabIndex        =   79
               Top             =   1050
               Width           =   2055
            End
            Begin VB.ComboBox cboInput 
               Height          =   315
               Index           =   4
               Left            =   1560
               Style           =   2  '드롭다운 목록
               TabIndex        =   78
               Top             =   1050
               Width           =   2025
            End
            Begin VB.ComboBox cboInput 
               Height          =   315
               Index           =   6
               Left            =   8760
               Style           =   2  '드롭다운 목록
               TabIndex        =   77
               Top             =   1050
               Width           =   2055
            End
            Begin VB.ComboBox cboInput 
               Height          =   315
               Index           =   7
               ItemData        =   "P_06012.frx":4224
               Left            =   1575
               List            =   "P_06012.frx":4226
               Style           =   2  '드롭다운 목록
               TabIndex        =   56
               Top             =   285
               Width           =   3750
            End
            Begin VB.ComboBox cboInput 
               Height          =   315
               Index           =   8
               Left            =   6945
               Style           =   2  '드롭다운 목록
               TabIndex        =   55
               Top             =   285
               Width           =   2055
            End
            Begin VB.ComboBox cboInput 
               Height          =   315
               Index           =   9
               Left            =   11625
               Style           =   2  '드롭다운 목록
               TabIndex        =   54
               Top             =   285
               Width           =   2655
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   20
               Left            =   105
               TabIndex        =   57
               Top             =   285
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "품    목"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   23
               Left            =   5475
               TabIndex        =   58
               Top             =   285
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "용    도"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   24
               Left            =   10155
               TabIndex        =   59
               Top             =   285
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "소    재"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   28
               Left            =   10905
               TabIndex        =   60
               Top             =   675
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "내 용 연 수"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   29
               Left            =   105
               TabIndex        =   61
               Top             =   645
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "경 과 일 수"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   30
               Left            =   3705
               TabIndex        =   62
               Top             =   645
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "환 산 일 수"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   31
               Left            =   7305
               TabIndex        =   63
               Top             =   645
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "배 상 비 율"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   32
               Left            =   10905
               TabIndex        =   64
               Top             =   1035
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "규정 보상금"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin CSTextLibCtl.sidbEdit sidbEdit 
               Height          =   315
               Index           =   1
               Left            =   12375
               TabIndex        =   65
               Top             =   660
               Width           =   1875
               _Version        =   262145
               _ExtentX        =   3307
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 999,999,999,999"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
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
               StartText.y     =   4
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   13
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
            Begin CSTextLibCtl.sidbEdit sidbEdit 
               Height          =   315
               Index           =   2
               Left            =   1575
               TabIndex        =   66
               Top             =   645
               Width           =   2010
               _Version        =   262145
               _ExtentX        =   3545
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 999,999,999,999"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
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
               StartText.y     =   4
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   13
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
            Begin CSTextLibCtl.sidbEdit sidbEdit 
               Height          =   315
               Index           =   3
               Left            =   5175
               TabIndex        =   67
               Top             =   645
               Width           =   2010
               _Version        =   262145
               _ExtentX        =   3545
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 999,999,999,999"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
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
               StartText.y     =   4
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   13
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
            Begin CSTextLibCtl.sidbEdit sidbEdit 
               Height          =   315
               Index           =   4
               Left            =   8775
               TabIndex        =   68
               Top             =   645
               Width           =   2010
               _Version        =   262145
               _ExtentX        =   3545
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 999,999,999,999"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
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
               StartText.y     =   4
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   13
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
            Begin CSTextLibCtl.sidbEdit sidbEdit 
               Height          =   315
               Index           =   5
               Left            =   12375
               TabIndex        =   69
               Top             =   1035
               Width           =   1875
               _Version        =   262145
               _ExtentX        =   3307
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 999,999,999,999"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
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
               StartText.y     =   4
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   13
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
               Index           =   12
               Left            =   90
               TabIndex        =   81
               Top             =   1050
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "사고 형태"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   14
               Left            =   3690
               TabIndex        =   82
               Top             =   1050
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "보 상 구 분"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   15
               Left            =   90
               TabIndex        =   83
               Top             =   1410
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "실제보상금액"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   16
               Left            =   90
               TabIndex        =   84
               Top             =   1770
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "심의결과"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   21
               Left            =   3690
               TabIndex        =   85
               Top             =   1410
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "옷도착일"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtInput 
               Height          =   315
               Index           =   5
               Left            =   5160
               TabIndex        =   86
               Top             =   1410
               Width           =   3075
               _ExtentX        =   5424
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   64290816
               CurrentDate     =   36684
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   18
               Left            =   7290
               TabIndex        =   87
               Top             =   1050
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "처 리 구 분"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin CSTextLibCtl.sidbEdit sidbEdit 
               Height          =   345
               Index           =   6
               Left            =   1560
               TabIndex        =   88
               Top             =   1410
               Width           =   1995
               _Version        =   262145
               _ExtentX        =   3519
               _ExtentY        =   609
               _StockProps     =   125
               Text            =   " 999,999,999,999"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
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
               StartText.y     =   5
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   13
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
               Index           =   35
               Left            =   8370
               TabIndex        =   93
               Top             =   1410
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "처리완료일"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtInput 
               Height          =   315
               Index           =   8
               Left            =   9840
               TabIndex        =   94
               Top             =   1410
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   64290816
               CurrentDate     =   36684
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox 
            Height          =   2490
            Index           =   1
            Left            =   120
            TabIndex        =   30
            Top             =   2370
            Width           =   14955
            _Version        =   851970
            _ExtentX        =   26379
            _ExtentY        =   4392
            _StockProps     =   79
            Caption         =   "※ 품목 정보"
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
            Appearance      =   6
            BorderStyle     =   1
            Begin FPSpreadADO.fpSpread spdView 
               Height          =   1830
               Left            =   11460
               TabIndex        =   90
               Top             =   330
               Visible         =   0   'False
               Width           =   8175
               _Version        =   524288
               _ExtentX        =   14420
               _ExtentY        =   3228
               _StockProps     =   64
               BackColorStyle  =   1
               EditEnterAction =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
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
               SpreadDesigner  =   "P_06012.frx":4228
               HighlightHeaders=   1
               HighlightStyle  =   1
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   12
               Left            =   6945
               MaxLength       =   20
               TabIndex        =   36
               Top             =   285
               Width           =   3165
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   1
               Left            =   1575
               TabIndex        =   35
               Top             =   1005
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   2
               Left            =   6975
               TabIndex        =   34
               Top             =   1005
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   3
               Left            =   1575
               TabIndex        =   33
               Top             =   1365
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   4
               Left            =   6975
               TabIndex        =   32
               Top             =   1725
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   5
               Left            =   1575
               TabIndex        =   31
               Top             =   2085
               Width           =   3735
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   5
               Left            =   105
               TabIndex        =   37
               Top             =   285
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "입고일자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   6
               Left            =   105
               TabIndex        =   38
               Top             =   1005
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "품    목"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   7
               Left            =   5505
               TabIndex        =   39
               Top             =   1005
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "브 랜 드"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtInput 
               Height          =   315
               Index           =   1
               Left            =   1575
               TabIndex        =   40
               Top             =   285
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   64290816
               CurrentDate     =   36686
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   3
               Left            =   5505
               TabIndex        =   41
               Top             =   285
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "택 번 호"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   4
               Left            =   105
               TabIndex        =   42
               Top             =   645
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "지사출고일자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtInput 
               Height          =   315
               Index           =   2
               Left            =   1575
               TabIndex        =   43
               Top             =   645
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   64290816
               CurrentDate     =   36686
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   8
               Left            =   5505
               TabIndex        =   44
               Top             =   645
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "고객출고일자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtInput 
               Height          =   315
               Index           =   3
               Left            =   6975
               TabIndex        =   45
               Top             =   645
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   64290816
               CurrentDate     =   36686
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   9
               Left            =   105
               TabIndex        =   46
               Top             =   1365
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "색    상"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   10
               Left            =   105
               TabIndex        =   47
               Top             =   1725
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "구입일자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtInput 
               Height          =   315
               Index           =   4
               Left            =   1575
               TabIndex        =   48
               Top             =   1725
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   64290816
               CurrentDate     =   36686
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   25
               Left            =   5505
               TabIndex        =   49
               Top             =   1725
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "구 입 처"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   26
               Left            =   105
               TabIndex        =   50
               Top             =   2085
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "구입형태"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   27
               Left            =   5505
               TabIndex        =   51
               Top             =   2085
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "구입가격"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin CSTextLibCtl.sidbEdit sidbEdit 
               Height          =   330
               Index           =   0
               Left            =   6975
               TabIndex        =   52
               Top             =   2070
               Width           =   3720
               _Version        =   262145
               _ExtentX        =   6562
               _ExtentY        =   582
               _StockProps     =   125
               Text            =   " 999,999,999,999"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
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
               StartText.y     =   4
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   13
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
            Begin XtremeSuiteControls.PushButton btnTAG 
               Height          =   315
               Left            =   10140
               TabIndex        =   89
               Top             =   270
               Width           =   630
               _Version        =   851970
               _ExtentX        =   1111
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "검색"
               Appearance      =   6
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox 
            Height          =   2580
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   135
            Width           =   15075
            _Version        =   851970
            _ExtentX        =   26591
            _ExtentY        =   4551
            _StockProps     =   79
            Caption         =   "※ 기본 정보"
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
            Appearance      =   6
            BorderStyle     =   1
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   10
               Left            =   6960
               TabIndex        =   91
               Top             =   990
               Width           =   3735
            End
            Begin VB.ComboBox cboInput 
               Height          =   315
               Index           =   1
               Left            =   6990
               Style           =   2  '드롭다운 목록
               TabIndex        =   76
               Top             =   600
               Width           =   3330
            End
            Begin VB.ComboBox cboOffice 
               Height          =   315
               Left            =   1590
               TabIndex        =   75
               Text            =   "cboOffice"
               Top             =   660
               Width           =   3330
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   6
               Left            =   1575
               TabIndex        =   25
               Top             =   1005
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   8
               Left            =   1575
               TabIndex        =   24
               Top             =   1770
               Width           =   9135
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   7
               Left            =   1575
               TabIndex        =   23
               Top             =   1365
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   11
               Left            =   6975
               TabIndex        =   22
               Top             =   1365
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Enabled         =   0   'False
               Height          =   315
               Index           =   0
               Left            =   6975
               TabIndex        =   18
               Top             =   225
               Width           =   3735
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   1
               Left            =   5505
               TabIndex        =   19
               Top             =   225
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "접수번호"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   2
               Left            =   5505
               TabIndex        =   20
               Top             =   615
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "대 리 점 명"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   34
               Left            =   105
               TabIndex        =   21
               Top             =   675
               Width           =   1470
               _ExtentX        =   2593
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "지 사 정 보"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   11
               Left            =   105
               TabIndex        =   26
               Top             =   1035
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "성    명"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   13
               Left            =   105
               TabIndex        =   27
               Top             =   1755
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "주    소"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   22
               Left            =   105
               TabIndex        =   28
               Top             =   1395
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "전화번호"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   33
               Left            =   5505
               TabIndex        =   29
               Top             =   1365
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "핸드폰 번호"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtInput 
               Height          =   315
               Index           =   0
               Left            =   1605
               TabIndex        =   72
               Top             =   300
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   556
               _Version        =   393216
               Format          =   64290816
               CurrentDate     =   36686
            End
            Begin Threed.SSPanel panCaption 
               Height          =   345
               Index           =   36
               Left            =   120
               TabIndex        =   73
               Top             =   300
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   609
               _Version        =   262144
               Caption         =   "사고접수일자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   17
               Left            =   5490
               TabIndex        =   92
               Top             =   990
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "고객코드"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
         End
         Begin MSComDlg.CommonDialog cdPicture 
            Left            =   13380
            Top             =   3870
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            DialogTitle     =   "사고 제품 이미지파일 선택"
         End
      End
   End
End
Attribute VB_Name = "P_06012"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim RS02 As ADODB.Recordset

Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Dim sPictureFile As String

Private Sub btnTAG_Click()
    On Error GoTo ErrRtn
    
    If txtInput(0).Text = "" Then
        MsgBox "접수번호를 생성하여 주십시요(신규)...... ", vbInformation
        Exit Sub
    End If
    
    If txtInput(12).Text = "" Then Exit Sub
    
    '-----------------------------------------------------------------------------------
    ' TB_고객
    '-----------------------------------------------------------------------------------
    ReDim sValue(3)
    
    sValue(0) = Mid(cboOffice.Text, 2, 4)                       ' 지사코드
    sValue(1) = Mid(cboInput(1).Text, 2, 6)                     ' 가맹점코드
    sValue(2) = ""                                              ' 접수일자(전체를 조회 한다.)
    sValue(3) = Replace(txtInput(12).Text, "-", "")             ' 택번호
    
    Set RS02 = New ADODB.Recordset
    Set RS02 = ExecPro("SP_M_06012_99", sValue(), Err_Num, Err_Dec)
    
    
    With spdView
        .Top = 600
        .Left = 2610
        
        .Visible = True
        .SetFocus
        
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS02.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = Format(RS02!택번호, "000-00-0000") & ""
            .Col = 2:  .Text = RS02!성명 & ""
            .Col = 3:  .Text = RS02!휴대전화 & ""
            .Col = 4:  .Text = RS02!접수일자 & ""
            .Col = 5:  .Text = RS02!의류명 & ""
            
            RS02.MoveNext
        Loop
        
        .Redraw = True
    End With
    
    RS02.Close
    Set RS02 = Nothing
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)

    Screen.MousePointer = 0
End Sub

Private Sub cboInput_Click(Index As Integer)
    Select Case Index
        Case 0: Call Data_Display("", "")
        
        Case 7, 8, 9
            If cboInput(7).Text <> "" And cboInput(8).Text <> "" And cboInput(9).Text <> "" Then
                ReDim sValue(3)
                
                sValue(0) = "0"
                sValue(1) = Mid(cboInput(7).Text, 2, 3)
                sValue(2) = Mid(cboInput(8).Text, 2, 3)
                sValue(3) = Mid(cboInput(9).Text, 2, 3)
                
                Set RS02 = New ADODB.Recordset
                Set RS02 = ExecPro("SP_M_06001_96", sValue(), Err_Num, Err_Dec)
        
                If RS02.RecordCount = 0 Then
                    sidbEdit(1).Text = ""
                    Exit Sub
                Else
                    sidbEdit(1).Text = RS02!내용연수 & ""
                End If
            End If
    End Select
End Sub
 
Private Sub cboOffice_Click()
    If cboOffice.ListIndex < 0 Then Exit Sub
    
    '-----------------------------------------------------------------
    '
    '-----------------------------------------------------------------
    cboInput(1).Clear
    
    ReDim sValue(2)
    
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01001_00", sValue(), Err_Num, Err_Dec)
    
    Do Until RS01.EOF
        cboInput(1).AddItem "[" & RS01!가맹점코드 & "] " & RS01!가맹점명
        RS01.MoveNext
    Loop
    RS01.Close
    Set RS01 = Nothing
    
    If cboInput(1).ListCount > 0 Then cboInput(1).ListIndex = 0
End Sub


Private Sub cboOffice_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        SearchString_한글 KeyAscii
'    Else
       ' SearchString KeyAscii
    End If

End Sub

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 0: Call Data_Display("", "")          ' 조회
        Case 1: Call Data_New                            ' 신규
        Case 2: Call DataSave               ' 저장
        Case 3:                     ' 삭제
        Case 4:                     ' 취소
        Case 5: Call DataPrint      ' 인쇄
        Case 6:                     ' 엑셀
        Case 7: Unload Me           ' 종료
        Case 8: StoreFind           ' 매장 찾기
        
        
        Case Else
            '
    End Select


End Sub
Private Sub StoreFind()
    
    cmdBtn(8).Tag = ""
    
    Set P_01001_A1.m_FormObj = Me
    P_01001_A1.Show vbModal
    
    
    ' 조회에서 매장 코드가 설정된 경우 해당 매장이 선택 되도록 한다.
    DoEvents
    If Trim(cmdBtn(8).Tag) = "" Then Exit Sub
    Call CboDataReSet(cmdBtn(8).Tag)

End Sub
Private Sub DataPrint()
    On Error GoTo ErrRtn
    
    Dim Query       As String
    Dim XML         As String
    
    Dim FileNumber
            
    FileNumber = FreeFile
    
    Open App.Path & "\XML\사고접수.XML" For Output As #FileNumber
    
    Print #FileNumber, "<?xml version=""1.0"" encoding=""EUC-KR""?>"
    Print #FileNumber, "<root>"
    
          XML = ""
        
    Query = "SELECT * FROM TB_가맹점"
    Query = Query & " WHERE 가맹점코드 = '" & Mid(txtInput(18).Text, 2, 6) & "'"
    Set RS02 = New ADODB.Recordset
    RS02.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If RS02.EOF Then
        XML = XML & "    <가맹점명></가맹점명>"
        XML = XML & "    <가맹점주소></가맹점주소>"
        XML = XML & "    <가맹점전화번호></가맹점전화번호>"
    Else
        XML = XML & "    <가맹점명>" & Func_Replace(RS02!가맹점명) & "</가맹점명>"
        XML = XML & "    <가맹점주소>" & Func_Replace(RS02!사업장주소) & "</가맹점주소>"
        XML = XML & "    <가맹점전화번호>" & Func_Replace(RS02!매장전화번호) & "</가맹점전화번호>"
    End If
    RS02.Close: Set RS02 = Nothing
    
    XML = XML & "    <지사정보>" & Func_Replace(txtInput(10).Text) & "</지사정보>"
    
    XML = XML & "    <소비자명>" & Func_Replace(txtInput(6).Text) & "</소비자명>"
    XML = XML & "    <소비자주소>" & Func_Replace(txtInput(8).Text) & "</소비자주소>"
    XML = XML & "    <소비자전화번호>" & Func_Replace(txtInput(7).Text) & "</소비자전화번호>"
    XML = XML & "    <소비자휴대전화>" & Func_Replace(txtInput(11).Text) & "</소비자휴대전화>"
    
    XML = XML & "    <품목>" & Func_Replace(txtInput(1).Text) & "</품목>"
    XML = XML & "    <상표>" & Func_Replace(txtInput(2).Text) & "</상표>"
    XML = XML & "    <구입일자>" & Format(dtInput(4).Value, "YYYY-MM-DD") & "</구입일자>"
    XML = XML & "    <색상>" & Func_Replace(txtInput(3).Text) & "</색상>"
    XML = XML & "    <구입처>" & Func_Replace(txtInput(4).Text) & "</구입처>"
    XML = XML & "    <최초택번호>" & Func_Replace(txtInput(19).Text) & "</최초택번호>"
    XML = XML & "    <구입형태>" & Func_Replace(txtInput(5).Text) & "</구입형태>"
    XML = XML & "    <최초입고일>" & Format(dtInput(1).Value, "YYYY-MM-DD") & "</최초입고일>"
    XML = XML & "    <구입가격>" & sidbEdit(0).Text & "</구입가격>"
    XML = XML & "    <사고접수일>" & Format(dtInput(0).Value, "yyyy-MM-dd") & "</사고접수일>"
    
    
    XML = XML & "    <크레임구분>" & Func_Replace(cboInput(4).Text) & "</크레임구분>"
    XML = XML & "    <보상구분>" & Func_Replace(cboInput(5).Text) & "</보상구분>"
    XML = XML & "    <보상금액>" & IIf(sidbEdit(6).Value = 0, "", sidbEdit(6).Text) & "</보상금액>"
    XML = XML & "    <보상제품정보>" & Func_Replace(txtInput(9).Text) & "</보상제품정보>"
    
    XML = XML & "    <보상품목>" & Func_Replace(cboInput(7).Text) & "</보상품목>"
    XML = XML & "    <보상용도>" & Func_Replace(cboInput(8).Text) & "</보상용도>"
    XML = XML & "    <보상소재>" & Func_Replace(cboInput(9).Text) & "</보상소재>"
    XML = XML & "    <내용연수>" & Func_Replace(sidbEdit(1).Text) & "</내용연수>"
    XML = XML & "    <경과일수>" & Func_Replace(sidbEdit(2).Text) & "</경과일수>"
    XML = XML & "    <환산일수>" & Func_Replace(sidbEdit(3).Text) & "</환산일수>"
    XML = XML & "    <배상비율>" & Func_Replace(sidbEdit(4).Text) & "</배상비율>"
    XML = XML & "    <보상산정금액>" & Func_Replace(sidbEdit(5).Text) & "</보상산정금액>"
    
'    XML = XML & "    <가맹점의견>" & Func_Replace(RichTextBox(0).Text) & "</가맹점의견>"
'    XML = XML & "    <지사의견>" & Func_Replace(RichTextBox(1).Text) & "</지사의견>"
'    XML = XML & "    <본사의견>" & Func_Replace(RichTextBox(2).Text) & "</본사의견>"
    
    Print #FileNumber, XML
    
    Print #FileNumber, "</root>"
    Close #FileNumber
        
    rpt사고접수.dc.FileURL = App.Path & "\XML\사고접수.XML"
    rpt사고접수.Show 1
    
    'rpt사고접수.PrintReport False
    'Unload rpt사고접수
    
    Exit Sub

ErrRtn:
    MsgBox Err.Description, vbInformation, "오류"
    Screen.MousePointer = 0
End Sub

Private Sub cmdRefresh_Click()
    Call CboDataReSet("%")
End Sub

 

Private Sub dtInput_Change(Index As Integer)
    If Index = 6 Or Index = 7 Then
        Call CboDataReSet("%")
    
    ' 입고일자가 바뀌면 해당입고일의 Tag번호를 읽어온다.
    ElseIf Index = 1 Then
'        ReDim sValue(1)
'
'        sValue(0) = Mid(cboInput(0).Text, 2, 3)
'        sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
'
'        Set RS01 = New ADODB.Recordset
'        Set RS01 = ExecPro("SP_06001_03", sValue(), Err_Num, Err_Dec)
'
'        cboInput(3).Clear
'
'        Do While Not RS01.EOF
'            cboInput(3).AddItem RS01!택번호
'
'            RS01.MoveNext
'        Loop
    End If
End Sub

Private Sub Form_Activate()
    
    If Store.Code = MASTER_OFFICE_CODE Then
        Call SubBottonEnable(cmdBtn, "11100111")
    Else
        Call SubBottonEnable(cmdBtn, "10000111")
    
    End If
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    If P_06012_Flag = False Then
        dtInput(0).Value = Date
        dtInput(1).Value = Date
        dtInput(2).Value = Date
        dtInput(3).Value = Date
        dtInput(4).Value = Date
        dtInput(5).Value = Date
        dtInput(6).Value = DateAdd("m", -1, Date)
        dtInput(7).Value = Date
        
        dtInput(1).Value = ""
        dtInput(2).Value = ""
        dtInput(3).Value = ""
        dtInput(4).Value = ""
        dtInput(5).Value = ""
        
        
        
        ' Combo BOX의 내역을 채운다.
        Call ComboAdd
        Call CboDataReSet("%")
        
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
        
        
        P_06012_Flag = True
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_06012_Flag = False
End Sub

Public Sub Data_Display(sStoreCode As String, sSEQ As String)
    On Error GoTo ErrRtn
    
    Dim RS01 As New ADODB.Recordset
    Dim sValue(2) As String
    Dim i As Integer
    Dim SSQL    As String
    
    If Trim(cboInput(0).Text) = "" Then Exit Sub
    
    If sStoreCode <> "" And sSEQ <> "" Then
        sValue(0) = "0"
        sValue(1) = sStoreCode
        sValue(2) = sSEQ
    
    Else
        sValue(0) = "0"
        sValue(1) = Trim(Mid(Trim(CStr(Split(cboInput(0).Text, "/")(2))), 2, 6))
        sValue(2) = CStr(Split(cboInput(0).Text, "/")(1))
    End If
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_06012_04", sValue(), Err_Num, Err_Dec)
    If RS01.EOF Then Exit Sub
    
    txtInput(11).Text = ""   '접수정보
    
    If Not IsNull(RS01!일련번호) Then txtInput(0).Text = RS01!일련번호 Else txtInput(0).Text = ""   '일련번호
    If Not IsNull(RS01!사고접수일자) Then dtInput(0).Value = RS01!사고접수일자
    
    If Not IsNull(RS01!지사코드) Then
        For i = 0 To cboOffice.ListCount - 1
            If RS01!지사코드 = Mid(cboOffice.List(i), 2, 4) Then
                cboOffice.ListIndex = i
                Exit For
            End If
        Next i
    End If
    
    If Not IsNull(RS01!가맹점코드) Then
        For i = 0 To cboInput(1).ListCount - 1
            If RS01!가맹점코드 = Mid(cboInput(1).List(i), 2, 6) Then
                cboInput(1).ListIndex = i
                Exit For
            End If
        Next i
    End If
    
    If Not IsNull(RS01!고객코드) Then txtInput(10).Text = RS01!고객코드
    If Not IsNull(RS01!성명) Then txtInput(6).Text = RS01!성명
    
    If Not IsNull(RS01!전화번호) Then txtInput(7).Text = RS01!전화번호 Else txtInput(7).Text = ""
    If Not IsNull(RS01!휴대전화) Then txtInput(11).Text = RS01!휴대전화 Else txtInput(11).Text = ""
    If Not IsNull(RS01!주소) Then txtInput(8).Text = RS01!주소 Else txtInput(8).Text = ""
    
    If Trim(RS01!입고일자) <> "" Then dtInput(1).Value = Format(RS01!입고일자, "YYYY-MM-DD") Else dtInput(1).Value = ""
    If Not IsNull(RS01!택번호) Then txtInput(12).Text = RS01!택번호 Else txtInput(12).Text = ""
    If Trim(RS01!지사출고일자) <> "" Then dtInput(2).Value = Format(RS01!지사출고일자, "YYYY-MM-DD") Else dtInput(2).Value = ""
    If Trim(RS01!고객출고일자) <> "" Then dtInput(3).Value = Format(RS01!고객출고일자, "YYYY-MM-DD") Else dtInput(3).Value = ""
    If Not IsNull(RS01!의류명) Then txtInput(1).Text = RS01!의류명 Else txtInput(1).Text = ""
    If Not IsNull(RS01!상표) Then txtInput(2).Text = RS01!상표 Else txtInput(2).Text = ""
    If Trim(RS01!색상) <> "" Then txtInput(3).Text = RS01!색상 Else txtInput(3).Text = ""
    
    
    
    If Not IsNull(RS01!구입일자) Then dtInput(4).Value = Format(RS01!구입일자, "YYYY-MM-DD") Else dtInput(4).Value = ""
    If Trim(RS01!구입처) <> "" Then txtInput(4).Text = RS01!구입처 Else txtInput(4).Text = ""
    If Trim(RS01!구입형태) <> "" Then txtInput(5).Text = RS01!구입형태 Else txtInput(5).Text = ""
    If Trim(RS01!구입가격) <> "" Then sidbEdit(0).Value = RS01!구입가격 Else sidbEdit(0).Value = 0
    

    If Not IsNull(RS01!품목) Then
        For i = 0 To cboInput(7).ListCount - 1
            If RS01!품목 = Mid(cboInput(7).List(i), 2, 3) Then
                cboInput(7).ListIndex = i
                Exit For
            End If
        Next i
    Else
        cboInput(7).ListIndex = -1
    End If
    
    If Not IsNull(RS01!용도) Then
        For i = 0 To cboInput(8).ListCount - 1
            If RS01!용도 = Mid(cboInput(8).List(i), 2, 3) Then
                cboInput(8).ListIndex = i
                Exit For
            End If
        Next i
    Else
        cboInput(8).ListIndex = -1
    End If
    
    If Not IsNull(RS01!소재) Then
        For i = 0 To cboInput(9).ListCount - 1
            If RS01!소재 = Mid(cboInput(9).List(i), 2, 3) Then
                cboInput(9).ListIndex = i
                Exit For
            End If
        Next i
    Else
        cboInput(9).ListIndex = -1
    End If
    
    If Not IsNull(RS01!내용연수) Then sidbEdit(1).Value = RS01!내용연수 Else sidbEdit(1).Value = ""
    If Not IsNull(RS01!경과일수) Then sidbEdit(2).Value = RS01!경과일수 Else sidbEdit(2).Value = ""
    If Not IsNull(RS01!환산일수) Then sidbEdit(3).Value = RS01!환산일수 Else sidbEdit(3).Value = ""
    If Not IsNull(RS01!배상비율) Then sidbEdit(4).Value = RS01!배상비율 Else sidbEdit(4).Value = ""
    If Not IsNull(RS01!배상금액) Then sidbEdit(5).Value = RS01!배상금액 Else sidbEdit(5).Value = ""
    
    
    
    If Not IsNull(RS01!크레임구분) Then
        For i = 0 To cboInput(4).ListCount - 1
            If Trim(RS01!크레임구분) = cboInput(4).List(i) Then
                cboInput(4).ListIndex = i
                Exit For
            End If
        Next i
    Else
        cboInput(4).ListIndex = -1
    End If
    
    If Not IsNull(RS01!보상구분) Then
        For i = 0 To cboInput(5).ListCount - 1
            If Trim(RS01!보상구분) = cboInput(5).List(i) Then
                cboInput(5).ListIndex = i
                Exit For
            End If
        Next i
    Else
        cboInput(5).ListIndex = -1
    End If
    
    If Not IsNull(RS01!처리구분) Then
        For i = 0 To cboInput(6).ListCount - 1
            If Trim(RS01!처리구분) = cboInput(5).List(i) Then
                cboInput(6).ListIndex = i
                Exit For
            End If
        Next i
    Else
        cboInput(6).ListIndex = -1
    End If
    
    
    If Not IsNull(RS01!보상금액) Then sidbEdit(6).Value = RS01!보상금액 Else sidbEdit(6).Value = ""
    If Trim(RS01!옷도착일) <> "" Then
        dtInput(5).Value = Format(RS01!옷도착일, "yyyy-MM-dd")
    Else
        dtInput(5).Value = Date
        dtInput(5).Value = ""
    End If
    
    If Trim(RS01!처리일자) <> "" Then
        dtInput(8).Value = Format(RS01!처리일자, "yyyy-MM-dd")
    Else
        dtInput(8).Value = Date
        dtInput(8).Value = ""
    End If
    
    
    If Not IsNull(RS01!심의결과) Then txtInput(9).Text = RS01!심의결과 Else txtInput(9).Text = ""
    

'    If Not IsNull(RS01!이미지) Then
'        pctPicture.Picture = LoadPicture(RS01!이미지)
'    End If
    RS01.Close
    Set RS01 = Nothing

    Exit Sub
    
ErrRtn:
'    Resume
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Public Sub DataDelete()
    If MsgBox("해당되는 사고내역을 삭제하시겠습니까?", vbYesNo + vbInformation + vbDefaultButton2, "데이터 삭제") = vbYes Then
        ReDim sValeu(1)
        
        sValue(0) = Format(dtInput(0).Value, "YYYY-MM-DD")                        ' 접수일자
        sValue(1) = txtInput(0).Text                                            ' 접수번호
        
        Call ExecPro("SP_06001_05", sValue(), Err_Num, Err_Dec)
        
        If Err_Num = 0 Then
            MsgBox "해당되는 데이터가 삭제 되었습니다.", vbInformation
            Call DataClear
            Exit Sub
        End If
    End If
End Sub

Private Sub ComboAdd()

    ' Call AgencyComboAdd(cboInput(0))

    ReDim sValue(1)
    
    sValue(0) = "0"
    
    '------------------------------------------------------------------------
    ' 사고 담당자
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_06001_90", sValue(), Err_Num, Err_Dec)

    Do While Not RS01.EOF
        cboInput(1).AddItem "[" & RS01!담당자코드 & "] " & RS01!담당자명
        
        RS01.MoveNext
    Loop
    RS01.Close

    '------------------------------------------------------------------------
    ' 크래임 구분
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_06012_91", sValue(), Err_Num, Err_Dec)

    Do While Not RS01.EOF
        ' 탈색, 파손, 이염, 분실, 기타
        'cboInput(4).AddItem "[" & RS01!코드 & "] " & RS01!내용
        cboInput(4).AddItem RS01!내용 & ""
        RS01.MoveNext
    Loop
    RS01.Close

    '------------------------------------------------------------------------
    '보상구분
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_06001_92", sValue(), Err_Num, Err_Dec)

    Do While Not RS01.EOF
        ' 수선, 물품이도후 일부보상, 현금, 제품, 복구
        'cboInput(5).AddItem "[" & RS01!코드 & "] " & RS01!내용
        cboInput(5).AddItem RS01!내용 & ""
        RS01.MoveNext
    Loop
    RS01.Close
    
    '------------------------------------------------------------------------
    '처리구분
    cboInput(6).AddItem "[001] 접수"
    cboInput(6).AddItem "[002] 진행중"
    cboInput(6).AddItem "[003] 처리완료"
    
    '------------------------------------------------------------------------
    ' 사고품 품목
    ReDim sValue(1)
    
    sValue(0) = "0"
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_06001_93", sValue(), Err_Num, Err_Dec)

    Do While Not RS01.EOF
        cboInput(7).AddItem "[" & RS01!품목코드 & "] " & RS01!품목명
        
        RS01.MoveNext
    Loop
    
    '------------------------------------------------------------------------
    ' 사고품 용도
    ReDim sValue(1)
    
    sValue(0) = "0"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_06001_94", sValue(), Err_Num, Err_Dec)

    Do While Not RS01.EOF
        cboInput(8).AddItem "[" & RS01!용도코드 & "] " & RS01!용도명
        
        RS01.MoveNext
    Loop

    '------------------------------------------------------------------------
    ' 사고품 용도
    ReDim sValue(1)
    
    sValue(0) = "0"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_06001_95", sValue(), Err_Num, Err_Dec)

    Do While Not RS01.EOF
        cboInput(9).AddItem "[" & RS01!소재코드 & "] " & RS01!소재명
        
        RS01.MoveNext
    Loop
End Sub

Public Sub DataSave()

    If DefCheck() = False Then Exit Sub


    If MsgBox("해당되는 내역을 저장하시겠습니까?", vbYesNo + vbInformation, "데이터 저장") = vbYes Then
        ReDim sValue(36)
        
        sValue(0) = txtInput(0).Text                                ' 일련번호
        sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")          ' 접수일자
        sValue(2) = Mid(cboInput(1).Text, 2, 6)                     ' 가맹점코드
        sValue(3) = Mid(cboOffice.Text, 2, 4)                       ' 지사코드
        
        sValue(4) = txtInput(10).Text                               ' 고객코드
        sValue(5) = txtInput(6).Text                                ' 성명
        sValue(6) = txtInput(7).Text                                ' 전화번호
        sValue(7) = txtInput(11).Text                               ' 휴대전화
        sValue(8) = Replace(txtInput(8).Text, "'", " ")             ' 주소
        
        sValue(9) = Format(dtInput(1).Value, "YYYY-MM-DD")          ' 입고일
        sValue(10) = txtInput(12).Text                              ' 택번호
        sValue(11) = Format(dtInput(2).Value, "YYYY-MM-DD")         ' 지사출고일
        sValue(12) = Format(dtInput(3).Value, "YYYY-MM-DD")         ' 고객출고일자
        sValue(13) = txtInput(1).Text                               ' 의류명
        sValue(14) = Replace(txtInput(2).Text, "'", " ")            ' 상표
        sValue(15) = Replace(txtInput(3).Text, "'", " ")            ' 색상
        
        sValue(16) = Format(dtInput(4).Value, "YYYY-MM-DD")         ' 구입일자
        sValue(17) = Replace(txtInput(4).Text, "'", " ")            ' 구입처
        sValue(18) = Replace(txtInput(5).Text, "'", " ")            ' 구입형태
        sValue(19) = sidbEdit(0).Value                              ' 구입가격
        
        sValue(20) = cboInput(7).Text                               ' 품목
        sValue(21) = cboInput(8).Text                               ' 용도
        sValue(22) = cboInput(9).Text                               ' 소재
        sValue(23) = sidbEdit(1).Value                              ' 내용연수
        sValue(24) = sidbEdit(2).Value                              ' 경과일수
        sValue(25) = sidbEdit(3).Value                              ' 환산일수
        sValue(26) = sidbEdit(4).Value                              ' 배상비율
        sValue(27) = sidbEdit(5).Value                              ' 배상금액
        
        sValue(28) = cboInput(4).Text                               ' 크레임구분
        sValue(29) = cboInput(5).Text                               ' 보상구분
        sValue(30) = cboInput(6).Text                               ' 처리구분
        sValue(31) = sidbEdit(6).Value                              ' 보상금액
        sValue(32) = Format(dtInput(5).Value, "YYYY-MM-DD")         ' 옷도착일
        sValue(33) = Format(dtInput(8).Value, "YYYY-MM-DD")         ' 처리일자
        sValue(34) = txtInput(9).Text                               ' 심의결과
        sValue(35) = USERNAME                                       ' 등록자명
        sValue(36) = USERNAME                                       ' 최종수정자
        
        
        
        Call ExecPro("SP_M_06012_01", sValue(), Err_Num, Err_Dec)
    
        If Err_Num = 0 Then
            MsgBox "해당되는 데이터가 정상적으로 저장이 되었습니다.", vbInformation
        
            ' 해당 기간의 매장의 내용을 다시 설정한다.
            Call CboDataReSet("%")

        Else
            MsgBox "[" & Err_Num & "] " & Err_Dec
            Exit Sub
        End If
    End If
End Sub

Public Sub Data_New()
    Dim i As Integer
    
    ReDim sValue(0)
    
'    dtInput(0).Value = Date
    
    sValue(0) = Format(dtInput(0).Value, "YYYY-MM-DD")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_06012_03", sValue(), Err_Num, Err_Dec)
    txtInput(0).Text = Val(RS01!접수번호)
    
    
    ' TEXT BOX 초기화
    For i = 1 To txtInput.Count - 1
        txtInput(i).Text = ""
    Next i
    
    ' Combo BOX 초기화
    For i = 0 To cboInput.Count - 1
        Select Case i
            Case 2, 3
            
            Case Else
                cboInput(i).ListIndex = -1
        End Select
    Next i
    
    ' sidbEdit BOX 초기화
    For i = 0 To sidbEdit.Count - 1
        sidbEdit(i).Text = ""
    Next i
    
    ' 일자Combo BOX 초기화
    For i = 1 To dtInput.Count - 1
        dtInput(i).Value = Date
        
        If dtInput(i).CheckBox = True Then dtInput(i).Value = ""
    Next i
End Sub


Private Sub sidbEdit_Change(Index As Integer)
    Select Case Index
        Case 1
            Call ClaimClac
    End Select

End Sub

Private Sub ClaimClac()
    If sidbEdit(1).Text = "0" Then
        Exit Sub
    End If

    If sidbEdit(1).Text = "" Then
        MsgBox "내용연수를 입력하십시요...", vbInformation
        txtInput(13).SetFocus
        Exit Sub
    End If
    
    If sidbEdit(0).Value <= "0" Then
        MsgBox "구입금액을 입력하십시요...", vbInformation
        sidbEdit(0).SetFocus
        Exit Sub
    End If
    
    If dtInput(4).Value = "" Then
        MsgBox " 구입일자를 등록하십시요...", vbInformation
        dtInput(4).SetFocus
        Exit Sub
    End If
    
    If dtInput(5).Value = "" Then
        MsgBox "처리일자를 등록하십시요...", vbInformation
        dtInput(5).SetFocus
        Exit Sub
    End If
    
    If sidbEdit(1).Text <> "" And sidbEdit(0).Text <> 0 And dtInput(4).Value <> "" And _
       Val(sidbEdit(1).Text) <> 0 Then
        Dim iDay As Integer
        Dim hDay As Integer
        Dim bRate As Integer

        ' 실제경과일수 계산 (구입일자 - 입고일자)
        iDay = dtInput(1).Value - dtInput(4).Value
        sidbEdit(2).Text = iDay

        ' 환산경과일수
        hDay = iDay / Val(sidbEdit(1).Text)
        sidbEdit(3).Text = hDay

        ' 배상비율 계산
        If hDay < 15 Then
            bRate = 95
        ElseIf hDay >= 15 And hDay < 45 Then
            bRate = 85
        ElseIf hDay >= 45 And hDay < 90 Then
            bRate = 70
        ElseIf hDay >= 90 And hDay < 135 Then
            bRate = 60
        ElseIf hDay >= 135 And hDay < 180 Then
            bRate = 50
        ElseIf hDay >= 180 And hDay < 225 Then
            bRate = 45
        ElseIf hDay >= 225 And hDay < 270 Then
            bRate = 40
        ElseIf hDay >= 270 And hDay < 315 Then
            bRate = 35
        ElseIf hDay >= 315 And hDay < 360 Then
            bRate = 30
        ElseIf hDay >= 360 Then
            bRate = 20
        End If

        sidbEdit(4).Text = bRate

        sidbEdit(5).Text = sidbEdit(0).Value * (bRate * 0.01)
    End If
End Sub

Private Sub DataClear()
    Dim i As Integer

    ' TEXT BOX 초기화
    For i = 1 To txtInput.Count - 1
        txtInput(i).Text = ""
    Next i
    
    ' Combo BOX 초기화
    For i = 0 To cboInput.Count - 1
        cboInput(i).ListIndex = -1
    Next i
'
'    ' MaskEdit BOX 초기화
'    For i = 0 To mskInput.Count - 1
'        mskInput(i).Text = ""
'    Next i
    
    ' 일자Combo BOX 초기화
    For i = 1 To dtInput.Count - 1
        dtInput(i).Value = Date
        dtInput(i).Value = ""
    Next i
End Sub


Private Sub CboDataReSet(StoreCode As String)
        ReDim sValue(3)
        
        sValue(0) = "0"                              '
        sValue(1) = Format(dtInput(6).Value, "yyyy-mm-dd")
        sValue(2) = Format(dtInput(7).Value, "yyyy-mm-dd")
        sValue(3) = StoreCode
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_M_06012_02", sValue(), Err_Num, Err_Dec)
        
        cboInput(0).Clear
        
        Do While Not RS01.EOF
            cboInput(0).AddItem Format(RS01!접수일자, "YYYY-MM-DD") & " / " & RS01!접수번호 & " / " & RS01!매장명
        
            RS01.MoveNext
        Loop
        
        If cboInput(0).ListCount > 0 Then cboInput(0).ListIndex = 0

End Sub

Private Sub spdView_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    On Error GoTo ErrRtn
    Dim RS01 As ADODB.Recordset
    Dim vText As Variant
    
    With spdView
    
        .GetText 1, Row, vText
        txtInput(12).Text = vText   '택번호
        txtInput(12).Tag = txtInput(12).Text
                                     
        .GetText 4, Row, vText
        dtInput(1).Value = Trim(vText)                          '접수일자
        
        .Visible = False
        
        If txtInput(12).Text = "" Then
            txtInput(12).SetFocus
            Exit Sub
        End If
        
        '-------------------------------------------------------------
        ' TB_입출고
        '-------------------------------------------------------------
        ReDim sValue(3)
        
        sValue(0) = Mid(cboOffice.Text, 2, 4)                       ' 지사코드
        sValue(1) = Mid(cboInput(1).Text, 2, 6)                     ' 가맹점코드
        sValue(2) = Format(dtInput(1).Value, "yyyy-MM-dd")          ' 접수일자
        sValue(3) = Replace(txtInput(12).Text, "-", "")             ' 택번호
        
        Set RS02 = New ADODB.Recordset
        Set RS02 = ExecPro("SP_M_06012_99", sValue(), Err_Num, Err_Dec)
        
        If Not RS02.EOF Then
            txtInput(10).Text = RS02!고객코드 & ""                      ' 5
            txtInput(6).Text = RS02!성명 & ""                          ' 6
            txtInput(7).Text = RS02!전화번호 & ""                      ' 7
            txtInput(11).Text = RS02!휴대전화 & ""                      ' 8
            txtInput(8).Text = RS02!주소 & ""                          ' 9
            
            txtInput(12).Text = RS02!택번호 & ""                        '12
            dtInput(1).Value = Format(RS02!접수일자, "YYYY-MM-DD")     '11
            dtInput(2).Value = Format(RS02!지사출고일자, "YYYY-MM-DD") '13
            dtInput(3).Value = Format(RS02!출고일자, "YYYY-MM-DD")     '13
                   
            txtInput(1).Text = RS02!의류명 & ""                        '15
            txtInput(2).Text = RS02!상표 & ""                          '16
            txtInput(3).Text = RS02!색상 & ""                          '17
        End If
        RS02.Close
        Set RS02 = Nothing
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)

    Screen.MousePointer = 0
End Sub



Private Function DefCheck() As Boolean

    DefCheck = False
    
    If txtInput(0).Text = "" Then
        MsgBox "접수번호를 생성하여 주십시요(신규)...... ", vbInformation
        Exit Function
    End If

    If sidbEdit(1).Text = "" Then
        MsgBox "내용연수를 입력하십시요...", vbInformation
        txtInput(13).SetFocus
        Exit Function
    End If
    
    If sidbEdit(0).Value <= "0" Then
        MsgBox "구입금액을 입력하십시요...", vbInformation
        sidbEdit(0).SetFocus
        Exit Function
    End If
    
    If dtInput(4).Value = "" Then
        MsgBox " 구입일자를 등록하십시요...", vbInformation
        dtInput(4).SetFocus
        Exit Function
    End If
    
    
    If Trim(cboInput(4).Text) = "" Then
        MsgBox "사고 형태를 등록하십시요...", vbInformation
        cboInput(4).SetFocus
        Exit Function
    End If
    
    If dtInput(5).Value = "" Then
        MsgBox "처리일자를 등록하십시요...", vbInformation
        dtInput(5).SetFocus
        Exit Function
    End If
    
    If sidbEdit(1).Text <> "" And sidbEdit(0).Text <> 0 And dtInput(4).Value <> "" And _
       Val(sidbEdit(1).Text) <> 0 Then
        Dim iDay As Integer
        Dim hDay As Integer
        Dim bRate As Integer

        ' 실제경과일수 계산 (구입일자 - 입고일자)
        iDay = dtInput(1).Value - dtInput(4).Value
        sidbEdit(2).Text = iDay

        ' 환산경과일수
        hDay = iDay / Val(sidbEdit(1).Text)
        sidbEdit(3).Text = hDay

        ' 배상비율 계산
        If hDay < 15 Then
            bRate = 95
        ElseIf hDay >= 15 And hDay < 45 Then
            bRate = 85
        ElseIf hDay >= 45 And hDay < 90 Then
            bRate = 70
        ElseIf hDay >= 90 And hDay < 135 Then
            bRate = 60
        ElseIf hDay >= 135 And hDay < 180 Then
            bRate = 50
        ElseIf hDay >= 180 And hDay < 225 Then
            bRate = 45
        ElseIf hDay >= 225 And hDay < 270 Then
            bRate = 40
        ElseIf hDay >= 270 And hDay < 315 Then
            bRate = 35
        ElseIf hDay >= 315 And hDay < 360 Then
            bRate = 30
        ElseIf hDay >= 360 Then
            bRate = 20
        End If

        sidbEdit(4).Text = bRate

        sidbEdit(5).Text = sidbEdit(0).Value * (bRate * 0.01)
    End If
    
    DefCheck = True
    
End Function

