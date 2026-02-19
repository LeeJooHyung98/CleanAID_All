VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_06001 
   Caption         =   "사고 처리 접수"
   ClientHeight    =   9900
   ClientLeft      =   4065
   ClientTop       =   2865
   ClientWidth     =   15270
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9900
   ScaleWidth      =   15270
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   9900
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   17463
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_06001.frx":0000
      Begin XtremeSuiteControls.TabControl TabControl 
         Height          =   2145
         Left            =   15
         TabIndex        =   72
         Top             =   7740
         Width           =   15240
         _Version        =   851970
         _ExtentX        =   26882
         _ExtentY        =   3784
         _StockProps     =   68
         ItemCount       =   4
         Item(0).Caption =   "[사 고 처 리 정 보]"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "TabControlPage(0)"
         Item(1).Caption =   "[대 리 점 / 고 객 의 견]"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "TabControlPage(1)"
         Item(2).Caption =   "[지 사 의 견]"
         Item(2).ControlCount=   1
         Item(2).Control(0)=   "TabControlPage(2)"
         Item(3).Caption =   " [본 사 의 견]"
         Item(3).ControlCount=   1
         Item(3).Control(0)=   "TabControlPage1"
         Begin XtremeSuiteControls.TabControlPage TabControlPage1 
            Height          =   1785
            Left            =   -69970
            TabIndex        =   103
            Top             =   330
            Visible         =   0   'False
            Width           =   15180
            _Version        =   851970
            _ExtentX        =   26776
            _ExtentY        =   3149
            _StockProps     =   1
            Page            =   3
            Begin RichTextLib.RichTextBox RichTextBox 
               Height          =   1665
               Index           =   2
               Left            =   30
               TabIndex        =   104
               Top             =   60
               Width           =   15105
               _ExtentX        =   26644
               _ExtentY        =   2937
               _Version        =   393217
               Enabled         =   -1  'True
               TextRTF         =   $"P_06001.frx":00B2
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage 
            Height          =   1785
            Index           =   2
            Left            =   -69970
            TabIndex        =   75
            Top             =   330
            Visible         =   0   'False
            Width           =   15180
            _Version        =   851970
            _ExtentX        =   26776
            _ExtentY        =   3149
            _StockProps     =   1
            Page            =   2
            Begin RichTextLib.RichTextBox RichTextBox 
               Height          =   1665
               Index           =   1
               Left            =   30
               TabIndex        =   76
               Top             =   60
               Width           =   15105
               _ExtentX        =   26644
               _ExtentY        =   2937
               _Version        =   393217
               Enabled         =   -1  'True
               TextRTF         =   $"P_06001.frx":0148
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage 
            Height          =   1785
            Index           =   1
            Left            =   -69970
            TabIndex        =   74
            Top             =   330
            Visible         =   0   'False
            Width           =   15180
            _Version        =   851970
            _ExtentX        =   26776
            _ExtentY        =   3149
            _StockProps     =   1
            Page            =   1
            Begin RichTextLib.RichTextBox RichTextBox 
               Height          =   1665
               Index           =   0
               Left            =   30
               TabIndex        =   77
               Top             =   60
               Width           =   15105
               _ExtentX        =   26644
               _ExtentY        =   2937
               _Version        =   393217
               Enabled         =   -1  'True
               TextRTF         =   $"P_06001.frx":01DE
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage 
            Height          =   1785
            Index           =   0
            Left            =   30
            TabIndex        =   73
            Top             =   330
            Width           =   15180
            _Version        =   851970
            _ExtentX        =   26776
            _ExtentY        =   3149
            _StockProps     =   1
            Page            =   0
            Begin VB.ComboBox cboInput 
               Height          =   315
               Index           =   6
               Left            =   12720
               Style           =   2  '드롭다운 목록
               TabIndex        =   78
               Top             =   240
               Width           =   1875
            End
            Begin VB.ComboBox cboInput 
               Height          =   315
               Index           =   4
               Left            =   1860
               Style           =   2  '드롭다운 목록
               TabIndex        =   81
               Top             =   210
               Width           =   3735
            End
            Begin VB.ComboBox cboInput 
               Height          =   315
               Index           =   5
               Left            =   7260
               Style           =   2  '드롭다운 목록
               TabIndex        =   80
               Top             =   210
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   9
               Left            =   1860
               MaxLength       =   50
               TabIndex        =   79
               Top             =   930
               Width           =   12735
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   12
               Left            =   390
               TabIndex        =   82
               Top             =   210
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "크레임 구분"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   14
               Left            =   5790
               TabIndex        =   83
               Top             =   210
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
               Left            =   390
               TabIndex        =   84
               Top             =   570
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "보 상 금 액"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   16
               Left            =   390
               TabIndex        =   85
               Top             =   930
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "비    고"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   21
               Left            =   5790
               TabIndex        =   86
               Top             =   570
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "처 리 일 자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtInput 
               Height          =   315
               Index           =   5
               Left            =   7260
               TabIndex        =   87
               Top             =   570
               Width           =   3765
               _ExtentX        =   6641
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   71892992
               CurrentDate     =   36684
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   18
               Left            =   11250
               TabIndex        =   88
               Top             =   240
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
               Left            =   1860
               TabIndex        =   95
               Top             =   570
               Width           =   3705
               _Version        =   262145
               _ExtentX        =   6535
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
         End
      End
      Begin Threed.SSPanel panInput 
         Height          =   510
         Left            =   15
         TabIndex        =   1
         Top             =   585
         Width           =   15240
         _ExtentX        =   26882
         _ExtentY        =   900
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   0
            Left            =   8040
            Style           =   2  '드롭다운 목록
            TabIndex        =   2
            Top             =   60
            Width           =   5355
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            TabIndex        =   3
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   71892992
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   4
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "접 수 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   19
            Left            =   4755
            TabIndex        =   5
            Top             =   60
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "접수일자 / 접수번호 / 매장명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   555
         Left            =   15
         TabIndex        =   6
         Top             =   15
         Width           =   7620
         _ExtentX        =   13441
         _ExtentY        =   979
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   255
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
         Caption         =   "사고 처리 접수 (P_06001)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_06001.frx":0274
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   555
         Index           =   0
         Left            =   7650
         TabIndex        =   7
         Top             =   15
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   979
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
         PictureBackground=   "P_06001.frx":0476
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
            Picture         =   "P_06001.frx":0678
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
            Picture         =   "P_06001.frx":0C12
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
            Picture         =   "P_06001.frx":11AC
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
            Picture         =   "P_06001.frx":1746
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
            Picture         =   "P_06001.frx":1CE0
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
            Picture         =   "P_06001.frx":227A
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
            Picture         =   "P_06001.frx":2814
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
            Picture         =   "P_06001.frx":2DAE
         End
      End
      Begin Threed.SSPanel panDetail 
         Height          =   6615
         Left            =   15
         TabIndex        =   16
         Top             =   1110
         Width           =   15240
         _ExtentX        =   26882
         _ExtentY        =   11668
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin Threed.SSFrame SSFrame6 
            Height          =   1095
            Left            =   60
            TabIndex        =   17
            Top             =   5310
            Width           =   14955
            _ExtentX        =   26379
            _ExtentY        =   1931
            _Version        =   262144
            Caption         =   "[보 상 산 정 기 준]"
            Begin VB.ComboBox cboInput 
               Height          =   315
               Index           =   9
               Left            =   8970
               Style           =   2  '드롭다운 목록
               TabIndex        =   20
               Top             =   300
               Width           =   1875
            End
            Begin VB.ComboBox cboInput 
               Height          =   315
               Index           =   8
               Left            =   5370
               Style           =   2  '드롭다운 목록
               TabIndex        =   19
               Top             =   300
               Width           =   1875
            End
            Begin VB.ComboBox cboInput 
               Height          =   315
               Index           =   7
               ItemData        =   "P_06001.frx":3348
               Left            =   1770
               List            =   "P_06001.frx":334A
               Style           =   2  '드롭다운 목록
               TabIndex        =   18
               Top             =   300
               Width           =   1875
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   20
               Left            =   300
               TabIndex        =   21
               Top             =   300
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
               Left            =   3900
               TabIndex        =   22
               Top             =   300
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
               Left            =   7500
               TabIndex        =   23
               Top             =   300
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
               Left            =   11100
               TabIndex        =   24
               Top             =   300
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
               Left            =   300
               TabIndex        =   25
               Top             =   660
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
               Left            =   3900
               TabIndex        =   26
               Top             =   660
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
               Left            =   7500
               TabIndex        =   27
               Top             =   660
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
               Left            =   11100
               TabIndex        =   28
               Top             =   660
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "배 상 금 액"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin CSTextLibCtl.sidbEdit sidbEdit 
               Height          =   345
               Index           =   1
               Left            =   12570
               TabIndex        =   90
               Top             =   300
               Width           =   1875
               _Version        =   262145
               _ExtentX        =   3307
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
            Begin CSTextLibCtl.sidbEdit sidbEdit 
               Height          =   345
               Index           =   2
               Left            =   1740
               TabIndex        =   91
               Top             =   660
               Width           =   1875
               _Version        =   262145
               _ExtentX        =   3307
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
            Begin CSTextLibCtl.sidbEdit sidbEdit 
               Height          =   345
               Index           =   3
               Left            =   5430
               TabIndex        =   92
               Top             =   630
               Width           =   1875
               _Version        =   262145
               _ExtentX        =   3307
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
            Begin CSTextLibCtl.sidbEdit sidbEdit 
               Height          =   345
               Index           =   4
               Left            =   8970
               TabIndex        =   93
               Top             =   660
               Width           =   1875
               _Version        =   262145
               _ExtentX        =   3307
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
            Begin CSTextLibCtl.sidbEdit sidbEdit 
               Height          =   345
               Index           =   5
               Left            =   12570
               TabIndex        =   94
               Top             =   690
               Width           =   1875
               _Version        =   262145
               _ExtentX        =   3307
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
         End
         Begin Threed.SSFrame SSFrame4 
            Height          =   1485
            Left            =   60
            TabIndex        =   29
            Top             =   1230
            Width           =   11235
            _ExtentX        =   19817
            _ExtentY        =   2619
            _Version        =   262144
            Caption         =   "[고 객 정 보]"
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   17
               Left            =   7170
               MaxLength       =   14
               TabIndex        =   33
               Top             =   1020
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   7
               Left            =   1770
               MaxLength       =   15
               TabIndex        =   32
               Top             =   1020
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   8
               Left            =   1770
               MaxLength       =   40
               TabIndex        =   31
               Top             =   660
               Width           =   9135
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   6
               Left            =   1770
               MaxLength       =   10
               TabIndex        =   30
               Top             =   300
               Width           =   3735
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   11
               Left            =   300
               TabIndex        =   34
               Top             =   300
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
               Left            =   300
               TabIndex        =   35
               Top             =   660
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
               Left            =   300
               TabIndex        =   36
               Top             =   1020
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "전 화 번 호"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   33
               Left            =   5700
               TabIndex        =   37
               Top             =   1020
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "핸드폰 번호"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
         End
         Begin Threed.SSFrame SSFrame2 
            Height          =   2475
            Left            =   60
            TabIndex        =   38
            Top             =   2760
            Width           =   11235
            _ExtentX        =   19817
            _ExtentY        =   4366
            _Version        =   262144
            Caption         =   "[품 목 정 보]"
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   5
               Left            =   1710
               MaxLength       =   10
               TabIndex        =   45
               Top             =   2040
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   4
               Left            =   7110
               MaxLength       =   20
               TabIndex        =   44
               Top             =   1680
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   3
               Left            =   1710
               MaxLength       =   10
               TabIndex        =   43
               Top             =   1320
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   2
               Left            =   7110
               MaxLength       =   20
               TabIndex        =   42
               Top             =   960
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   1
               Left            =   1710
               MaxLength       =   20
               TabIndex        =   41
               Top             =   960
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   19
               Left            =   7080
               MaxLength       =   20
               TabIndex        =   40
               Top             =   240
               Width           =   3165
            End
            Begin VB.CommandButton cmdTag 
               Caption         =   "..."
               Height          =   315
               Left            =   10200
               TabIndex        =   39
               Top             =   240
               Width           =   615
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   5
               Left            =   240
               TabIndex        =   46
               Top             =   240
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "입 고 일 자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   6
               Left            =   240
               TabIndex        =   47
               Top             =   960
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
               Left            =   5640
               TabIndex        =   48
               Top             =   960
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "브  랜  드"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtInput 
               Height          =   315
               Index           =   1
               Left            =   1710
               TabIndex        =   49
               Top             =   240
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   71892992
               CurrentDate     =   36686
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   3
               Left            =   5640
               TabIndex        =   50
               Top             =   240
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "택  번  호"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   4
               Left            =   240
               TabIndex        =   51
               Top             =   600
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "출 고 일 자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtInput 
               Height          =   315
               Index           =   2
               Left            =   1710
               TabIndex        =   52
               Top             =   600
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   71892992
               CurrentDate     =   36686
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   8
               Left            =   5640
               TabIndex        =   53
               Top             =   600
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "인 도 일 자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtInput 
               Height          =   315
               Index           =   3
               Left            =   7110
               TabIndex        =   54
               Top             =   600
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   71892992
               CurrentDate     =   36686
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   9
               Left            =   240
               TabIndex        =   55
               Top             =   1320
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
               Left            =   240
               TabIndex        =   56
               Top             =   1680
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "구 입 일 자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtInput 
               Height          =   315
               Index           =   4
               Left            =   1710
               TabIndex        =   57
               Top             =   1680
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   71892992
               CurrentDate     =   36686
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   25
               Left            =   5640
               TabIndex        =   58
               Top             =   1680
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "구  입  처"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   26
               Left            =   240
               TabIndex        =   59
               Top             =   2040
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "구 입 형 태"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   27
               Left            =   5640
               TabIndex        =   60
               Top             =   2040
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "구 입 가 격"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin CSTextLibCtl.sidbEdit sidbEdit 
               Height          =   345
               Index           =   0
               Left            =   7140
               TabIndex        =   89
               Top             =   2040
               Width           =   3675
               _Version        =   262145
               _ExtentX        =   6482
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
         End
         Begin Threed.SSFrame SSFrame1 
            Height          =   1095
            Left            =   60
            TabIndex        =   61
            Top             =   60
            Width           =   11235
            _ExtentX        =   19817
            _ExtentY        =   1931
            _Version        =   262144
            Caption         =   "[접 수 내 역]"
            Begin VB.ComboBox cboInput 
               Height          =   315
               Index           =   1
               Left            =   1710
               Style           =   2  '드롭다운 목록
               TabIndex        =   65
               Top             =   660
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   0
               Left            =   1710
               TabIndex        =   64
               Top             =   300
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   18
               Left            =   7080
               TabIndex        =   63
               Top             =   300
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   10
               Left            =   7080
               TabIndex        =   62
               Top             =   690
               Width           =   3735
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   1
               Left            =   240
               TabIndex        =   66
               Top             =   300
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "접 수 번 호"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   2
               Left            =   5640
               TabIndex        =   67
               Top             =   300
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
               Index           =   17
               Left            =   240
               TabIndex        =   68
               Top             =   660
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "담당자명"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   34
               Left            =   5640
               TabIndex        =   69
               Top             =   690
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "지 사 정 보"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
         End
         Begin Threed.SSFrame SSFrame3 
            Height          =   2325
            Left            =   11460
            TabIndex        =   70
            Top             =   2850
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   4101
            _Version        =   262144
            Caption         =   "사고 제품 이미지"
            Begin VB.PictureBox pctPicture 
               BackColor       =   &H8000000E&
               Height          =   1545
               Left            =   150
               ScaleHeight     =   1485
               ScaleWidth      =   3315
               TabIndex        =   71
               Top             =   240
               Width           =   3375
            End
            Begin Threed.SSCommand cmdSubBtn 
               Height          =   375
               Index           =   0
               Left            =   780
               TabIndex        =   96
               Top             =   1830
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   661
               _Version        =   262144
               Caption         =   "이미지추가"
            End
            Begin Threed.SSCommand cmdSubBtn 
               Height          =   405
               Index           =   1
               Left            =   2100
               TabIndex        =   97
               Top             =   1800
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   714
               _Version        =   262144
               Caption         =   "이미지제거"
            End
         End
         Begin MSComDlg.CommonDialog cdPicture 
            Left            =   11340
            Top             =   4980
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            DialogTitle     =   "사고 제품 이미지파일 선택"
         End
         Begin Threed.SSFrame SSFrame5 
            Height          =   2325
            Left            =   11460
            TabIndex        =   98
            Top             =   60
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   4101
            _Version        =   262144
            Caption         =   "[ 승 인 내 역 ]"
            Begin VB.CommandButton cmdApply 
               Caption         =   "본사 접수 승인"
               Height          =   585
               Index           =   1
               Left            =   210
               TabIndex        =   102
               Top             =   1620
               Width           =   3165
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   11
               Left            =   210
               Locked          =   -1  'True
               TabIndex        =   100
               Top             =   570
               Width           =   3165
            End
            Begin VB.CommandButton cmdApply 
               Caption         =   "지사 접수 승인"
               Height          =   585
               Index           =   0
               Left            =   210
               TabIndex        =   99
               Top             =   990
               Width           =   3165
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   35
               Left            =   210
               TabIndex        =   101
               Top             =   270
               Width           =   3150
               _ExtentX        =   5556
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "가맹점 작성정보"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
         End
      End
   End
End
Attribute VB_Name = "P_06001"
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

Private Sub cboInput_Click(Index As Integer)
    Select Case Index
'        Case 3
'            ReDim sValue(2)
'
'            sValue(0) = Mid(cboInput(0).Text, 2, 3)
'            sValue(1) = Format(dtInput(1).Value, "YYYY-MM-DD")
'            sValue(2) = Mid(cboInput(3).Text, 1, 1) & Mid(cboInput(3).Text, 3, 3)
'
'            Set RS02 = New ADODB.Recordset
'            Set RS02 = ExecPro("PRO_P_06001_07", sValue(), Err_Num, Err_Dec)
'
'            If Err_Num = 0 Then
'                If RS02.RecordCount = 0 Then
'
'                Else
'                    If Not IsNull(RS02!품명) Then txtInput(1).Text = RS02!품명
'                    If Not IsNull(RS02!브랜드) Then txtInput(2).Text = RS02!브랜드
'                    If Not IsNull(RS02!색상) Then txtInput(3).Text = RS02!색상
'                End If
'            Else
'                MsgBox "[" & Err_Num & "] " & Err_Dec
'                Exit Sub
'            End If
'        Case 6
'            dtInput(0).Value = Format(Mid(cboInput(6).Text, 1, 10), "YYYY-MM-DD")
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

Private Sub cmdApply_Click(Index As Integer)
    Dim sMSG As String
    
    ReDim sValue(3)

    sValue(2) = txtInput(0).Text                ' 일련번호
    sValue(3) = Mid(txtInput(18).Text, 2, 6)    ' 가맹점 코드
    
    M_SP_NAME = "SP_M_06001_02"

    ' 본사 승인일 경우만 처리한다.
    If Store.Code = MASTER_CODE And Index = 1 Then
        
        ' 승인 처리
        If InStr(cmdApply(1).Caption, "미승인") > 0 Then
            sMSG = "해당 내용을 접수 승인 처리 하시 겠습니까?"
            If MsgBox(sMSG, vbYesNo + vbInformation + vbDefaultButton2, "확인") = vbNo Then Exit Sub
            
            sValue(0) = "1"                             ' 0.지사 1.본사
            sValue(1) = "Y"                             ' Y.승인/N.취소
            
            Set RS02 = New ADODB.Recordset
            Set RS02 = ExecPro(M_SP_NAME, sValue(), Err_Num, Err_Dec)
            cmdApply(1).Caption = "본사 :" & RS02!승인일시
            RS02.Close:     Set RS02 = Nothing
            Exit Sub

        ' 승인 취소 처리
        Else
                
            sMSG = "해당 내용을 접수 승인 취소 처리 하시 겠습니까?"
            If MsgBox(sMSG, vbYesNo + vbInformation + vbDefaultButton2, "확인") = vbNo Then Exit Sub
            
            sValue(0) = "1"                             ' 0.지사 1.본사
            sValue(1) = "N"                             ' Y.승인/N.취소
            
            Set RS02 = New ADODB.Recordset
            Set RS02 = ExecPro(M_SP_NAME, sValue(), Err_Num, Err_Dec)
            cmdApply(1).Caption = "본사 미승인"
            RS02.Close:     Set RS02 = Nothing
            Exit Sub
        
        End If
    
    ' 지사일 경우만 처리한다.
    Else
        
    
        ' 승인 처리
        If InStr(cmdApply(0).Caption, "미승인") > 0 Then
            sMSG = "해당 내용을 접수 승인 처리 하시 겠습니까?"
            If MsgBox(sMSG, vbYesNo + vbInformation + vbDefaultButton2, "확인") = vbNo Then Exit Sub
            
            sValue(0) = "0"                             ' 0.지사 1.본사
            sValue(1) = "Y"                             ' Y.승인/N.취소
            
            Set RS02 = New ADODB.Recordset
            Set RS02 = ExecPro(M_SP_NAME, sValue(), Err_Num, Err_Dec)
            cmdApply(0).Caption = "지사 :" & RS02!승인일시
            RS02.Close:     Set RS02 = Nothing
            Exit Sub

        ' 승인 취소 처리
        Else
                
            sMSG = "해당 내용을 접수 승인 취소 처리 하시 겠습니까?"
            If MsgBox(sMSG, vbYesNo + vbInformation + vbDefaultButton2, "확인") = vbNo Then Exit Sub
            
            sValue(0) = "0"                             ' 0.지사 1.본사
            sValue(1) = "N"                             ' Y.승인/N.취소
            
            Set RS02 = New ADODB.Recordset
            Set RS02 = ExecPro(M_SP_NAME, sValue(), Err_Num, Err_Dec)
            cmdApply(0).Caption = "지사 미승인"
            RS02.Close:     Set RS02 = Nothing
            Exit Sub
        
        End If
    
    End If
    
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 0: Call Data_Display           ' 조회
        Case 1:                             ' 신규
        Case 2: Call DataSave               ' 저장
        
        Case 3:            ' 삭제
        Case 4:            ' 취소
        Case 5:            ' 인쇄
        Case 6:            ' 화면
        Case 7: Unload Me           ' 종료
        
        Case Else
            '
    End Select


End Sub

Private Sub cmdSubBtn_Click(Index As Integer)
    Select Case Index
        Case 0
            cdPicture.Action = 1
            pctPicture.Picture = LoadPicture(cdPicture.FileName)
            sPictureFile = cdPicture.FileName
        Case 1
            pctPicture.Picture = LoadPicture("")
            sPictureFile = ""
    End Select
End Sub

Private Sub dtInput_Change(Index As Integer)
    If Index = 0 Then
        ReDim sValue(2)
        
        sValue(0) = "0"
        sValue(1) = Format(dtInput(0).Value, "YYYY")
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("PRO_P_06001_00", sValue(), Err_Num, Err_Dec)
        
        cboInput(6).Clear
        
        Do While Not RS01.EOF
            cboInput(6).AddItem Format(RS01!접수일자, "YYYY-MM-DD") & " / " & RS01!접수번호 & " / " & RS01!매장명
        
            RS01.MoveNext
        Loop
    
    ' 입고일자가 바뀌면 해당입고일의 Tag번호를 읽어온다.
    ElseIf Index = 1 Then
        ReDim sValue(1)
        
        sValue(0) = Mid(cboInput(0).Text, 2, 3)
        sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("PRO_P_06001_03", sValue(), Err_Num, Err_Dec)
        
        cboInput(3).Clear
        
        Do While Not RS01.EOF
            cboInput(3).AddItem RS01!택번호
            
            RS01.MoveNext
        Loop
    End If
End Sub

Private Sub Form_Activate()
    
    If Store.Code = MASTER_CODE Then
        Call SubBottonEnable(cmdBtn, "11110111")
    Else
        Call SubBottonEnable(cmdBtn, "10000111")
        cmdApply(1).Enabled = False
    End If
    
End Sub

Private Sub Form_Load()

    If P_06001_Flag = False Then
        dtInput(0).Value = Date
        dtInput(1).Value = ""
        dtInput(2).Value = ""
        dtInput(3).Value = ""
        dtInput(4).Value = ""
        dtInput(5).Value = ""
        
        ' Combo BOX의 내역을 채운다.
        Call ComboAdd
            
        ReDim sValue(2)
        
        sValue(0) = "0"
        sValue(1) = Format(dtInput(0).Value, "yyyy")
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_M_06001_00", sValue(), Err_Num, Err_Dec)
        
        cboInput(0).Clear
        
        Do While Not RS01.EOF
            cboInput(0).AddItem Format(RS01!접수일자, "YYYY-MM-DD") & " / " & RS01!접수번호 & " / " & RS01!매장명
        
            RS01.MoveNext
        Loop
        
        RS01.Close
        Set RS01 = Nothing
        
        P_06001_Flag = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_06001_Flag = False
End Sub

Public Sub Data_Display()
    Dim i As Integer
    
    ReDim sValue(2)
    
    If Trim(cboInput(0).Text) = "" Then Exit Sub
    
    sValue(0) = "0"
    sValue(1) = Trim(Mid(Trim(CStr(Split(cboInput(0).Text, "/")(2))), 2, 6))
    sValue(2) = CStr(Split(cboInput(0).Text, "/")(1))
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_06001_01", sValue(), Err_Num, Err_Dec)
    
    If RS01.EOF Then
        Exit Sub
    End If
    
    If Not IsNull(RS01!접수일자) Then txtInput(11).Text = RS01!접수일자 Else txtInput(11).Text = ""   '접수정보
    If Not IsNull(RS01!지사승인) Then
        cmdApply(0).Caption = IIf(UCase(RS01!지사승인) = "Y", "지사 :" & RS01!지사승인일시, "지사 미승인")
    Else
        cmdApply(0).Caption = "지사 미승인"
    End If
    If Not IsNull(RS01!본사승인) Then
        cmdApply(1).Caption = IIf(UCase(RS01!본사승인) = "Y", "본사 :" & RS01!본사승인일시 & "", "본사 미승인")
    Else
        cmdApply(1).Caption = "본사 미승인"
    End If
    
    If Not IsNull(RS01!일련번호) Then txtInput(0).Text = RS01!일련번호 Else txtInput(0).Text = ""   '일련번호
    txtInput(18).Text = Trim(CStr(Split(cboInput(0).Text, "/")(2)))                                 '가맹점 코드/ 명칭
    
    If Not IsNull(RS01!담당자명) Then                                                             ' 담당자명
        For i = 0 To cboInput(1).ListCount - 1
            If Trim(RS01!담당자명) = Trim(cboInput(1).List(i)) Then
                cboInput(1).ListIndex = i
                Exit For
            End If
        Next i
    Else
        cboInput(1).ListIndex = -1
    End If
    If Not IsNull(RS01!지사코드) Then txtInput(10).Text = RS01!지사코드 Else txtInput(10).Text = ""
    
    
    If Not IsNull(RS01!고객코드) Then
        txtInput(6).Text = "[" & RS01!고객코드 & "] " & RS01!성명
    Else
        txtInput(6).Text = ""
    End If
    If Not IsNull(RS01!전화번호) Then txtInput(7).Text = RS01!전화번호 Else txtInput(7).Text = ""
    If Not IsNull(RS01!휴대폰) Then txtInput(17).Text = RS01!휴대폰 Else txtInput(17).Text = ""
    If Not IsNull(RS01!주소) Then txtInput(8).Text = RS01!주소 Else txtInput(8).Text = ""
    
    If Trim(RS01!입고일) <> "" Then dtInput(1).Value = Format(RS01!입고일, "####-##-##") Else dtInput(1).Value = ""
    If Not IsNull(RS01!택번호) Then txtInput(19).Text = RS01!택번호 Else txtInput(19).Text = ""
    If Trim(RS01!출고일) <> "" Then dtInput(2).Value = Format(RS01!출고일, "####-##-##") Else dtInput(2).Value = ""
    If Trim(RS01!인도일자) <> "" Then dtInput(3).Value = Format(RS01!인도일자, "####-##-##") Else dtInput(3).Value = ""
    If Not IsNull(RS01!의류명) Then txtInput(1).Text = RS01!의류명 Else txtInput(1).Text = ""
    If Not IsNull(RS01!상표) Then txtInput(2).Text = RS01!상표 Else txtInput(2).Text = ""
    
    
    If Trim(RS01!색상) <> "" Then txtInput(3).Text = RS01!색상 Else txtInput(3).Text = ""
    If Not IsNull(RS01!구입일자) Then dtInput(4).Value = Format(RS01!구입일자, "####-##-##") Else dtInput(4).Value = ""
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
    If Trim(RS01!처리일자) <> "" Then dtInput(5).Value = Format(RS01!처리일자, "####-##-##") Else dtInput(5).Value = ""
    
    
    If Not IsNull(RS01!비고) Then txtInput(9).Text = RS01!비고 Else txtInput(9).Text = ""
    
    If Not IsNull(RS01!대리점의견) Then RichTextBox(0).Text = RS01!대리점의견 Else RichTextBox(0).Text = ""
    If Not IsNull(RS01!지사의견) Then RichTextBox(1).Text = RS01!지사의견 Else RichTextBox(1).Text = ""
    If Not IsNull(RS01!본사의견) Then RichTextBox(2).Text = RS01!본사의견 Else RichTextBox(2).Text = ""

'    If Not IsNull(RS01!이미지) Then
'        pctPicture.Picture = LoadPicture(RS01!이미지)
'    End If
End Sub

Public Sub DataDelete()
    If MsgBox("해당되는 사고내역을 삭제하시겠습니까?", vbYesNo + vbInformation + vbDefaultButton2, "데이터 삭제") = vbYes Then
        ReDim sValeu(1)
        
        sValue(0) = Format(dtInput(0).Value, "YYYY-MM-DD")                        ' 접수일자
        sValue(1) = txtInput(0).Text                                            ' 접수번호
        
        Call ExecPro("PRO_P_06001_05", sValue(), Err_Num, Err_Dec)
        
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
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_06001_90", sValue(), Err_Num, Err_Dec)

    Do While Not RS01.EOF
        cboInput(1).AddItem "[" & RS01!담당자코드 & "] " & RS01!담당자명
        
        RS01.MoveNext
    Loop
    RS01.Close

    ' 크래임 구분
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_06001_91", sValue(), Err_Num, Err_Dec)

    Do While Not RS01.EOF
        ' 탈색, 파손, 이염, 분실, 기타
        'cboInput(4).AddItem "[" & RS01!코드 & "] " & RS01!내용
        cboInput(4).AddItem RS01!내용 & ""
        RS01.MoveNext
    Loop
    RS01.Close

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
    
  
    
    ReDim sValue(1)
    
    sValue(0) = "0"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_06001_93", sValue(), Err_Num, Err_Dec)

    Do While Not RS01.EOF
        cboInput(7).AddItem "[" & RS01!품목코드 & "] " & RS01!품목명
        
        RS01.MoveNext
    Loop
    

    ReDim sValue(1)
    
    sValue(0) = "0"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_06001_94", sValue(), Err_Num, Err_Dec)

    Do While Not RS01.EOF
        cboInput(8).AddItem "[" & RS01!용도코드 & "] " & RS01!용도명
        
        RS01.MoveNext
    Loop


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
    If MsgBox("해당되는 내역을 저장하시겠습니까?", vbYesNo + vbInformation, "데이터 저장") = vbYes Then
        ReDim sValue(37)
        
        sValue(0) = txtInput(0).Text                                ' 일련번호
        sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")          ' 접수일자
        sValue(2) = Mid(txtInput(18).Text, 2, 6)                    ' 가맹점코드
        sValue(3) = txtInput(10).Text                               ' 지사코드
        sValue(4) = cboInput(1).Text                                ' 담당자명
        
        sValue(5) = Mid(txtInput(6).Text, 2, 6)                     ' 고객코드
        sValue(6) = Trim(Mid(txtInput(6).Text, 9))                  ' 성명
        
        sValue(7) = txtInput(7).Text                                ' 전화번호
        sValue(8) = txtInput(17).Text                               ' 휴대폰
        sValue(9) = Replace(txtInput(8).Text, "'", " ")             ' 주소
        
        sValue(10) = Format(dtInput(1).Value, "YYYY-MM-DD")         ' 입고일
        sValue(11) = txtInput(19).Text                              ' 택번호
        sValue(12) = Format(dtInput(2).Value, "YYYY-MM-DD")         ' 출고일
        sValue(13) = Format(dtInput(3).Value, "YYYY-MM-DD")         ' 인도일자
        sValue(14) = txtInput(1).Text                               ' 의류명
        sValue(15) = Replace(txtInput(2).Text, "'", " ")            ' 상표
        
        sValue(16) = Replace(txtInput(3).Text, "'", " ")            ' 색상
        sValue(17) = Format(dtInput(4).Value, "YYYY-MM-DD")         ' 구입일자
        sValue(18) = Replace(txtInput(4).Text, "'", " ")            ' 구입처
        sValue(19) = Replace(txtInput(5).Text, "'", " ")            ' 구입형태
        sValue(20) = sidbEdit(0).Value                              ' 구입가격
        
        sValue(21) = cboInput(7).Text                               ' 품목
        sValue(22) = cboInput(8).Text                               ' 용도
        sValue(23) = cboInput(9).Text                               ' 소재
        sValue(24) = sidbEdit(1).Value                              ' 내용연수
        sValue(25) = sidbEdit(2).Value                              ' 경과일수
        sValue(26) = sidbEdit(3).Value                              ' 환산일수
        sValue(27) = sidbEdit(4).Value                              ' 배상비율
        sValue(28) = sidbEdit(5).Value                              ' 배상금액
        
        sValue(29) = cboInput(4).Text                               ' 크레임구분
        sValue(30) = cboInput(5).Text                               ' 보상구분
        sValue(31) = cboInput(6).Text                               ' 처리구분
        sValue(32) = sidbEdit(6).Value                              ' 보상금액
        sValue(33) = Format(dtInput(5).Value, "YYYY-MM-DD")         ' 처리일자
        sValue(34) = txtInput(9).Text                               ' 비고
        sValue(35) = Replace(RichTextBox(0).Text, "'", " ")         ' 대리점의견
        sValue(36) = Replace(RichTextBox(1).Text, "'", " ")         ' 지사의견
        sValue(37) = Replace(RichTextBox(2).Text, "'", " ")         ' 본사의견
        
        Call ExecPro("SP_M_06001_04", sValue(), Err_Num, Err_Dec)
    
        If Err_Num = 0 Then
            MsgBox "해당되는 데이터가 정상적으로 저장이 되었습니다.", vbInformation
        
            ReDim sValue(2)
            
            sValue(0) = "0"
            sValue(1) = Format(dtInput(0).Value, "YYYY")
            
            Set RS01 = New ADODB.Recordset
            Set RS01 = ExecPro("SP_M_06001_00", sValue(), Err_Num, Err_Dec)
            
            cboInput(0).Clear
            
            Do While Not RS01.EOF
                cboInput(0).AddItem Format(RS01!접수일자, "YYYY-MM-DD") & " / " & RS01!접수번호 & " / " & RS01!매장명
            
                RS01.MoveNext
            Loop
        Else
            MsgBox "[" & Err_Num & "] " & Err_Dec
            Exit Sub
        End If
    End If
End Sub

Public Sub DataAdd()
    Dim i As Integer
    
    ReDim sValue(0)
    
'    dtInput(0).Value = Date
    
    sValue(0) = Format(dtInput(0).Value, "YYYY-MM-DD")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("PRO_P_06001_02", sValue(), Err_Num, Err_Dec)
    
    If RS01.RecordCount = 0 Or IsNull(RS01!접수번호) Then
        txtInput(0).Text = "0001"
    Else
        txtInput(0).Text = Right("0000" & Val(RS01!접수번호) + 1, 4)
    End If
    
    ' TEXT BOX 초기화
    For i = 1 To txtInput.Count - 1
        txtInput(i).Text = ""
    Next i
    
    ' Combo BOX 초기화
    For i = 0 To cboInput.Count - 1
        cboInput(i).ListIndex = -1
    Next i
    
    ' MaskEdit BOX 초기화
    For i = 0 To mskInput.Count - 1
        mskInput(i).Text = ""
    Next i
    
    ' 일자Combo BOX 초기화
    For i = 1 To dtInput.Count - 1
        dtInput(i).Value = Date
        dtInput(i).Value = ""
    Next i
End Sub

Private Sub spdView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu P_00000.PopUp
    End If
End Sub

Private Sub txtInput_Change(Index As Integer)
    Select Case Index
        Case 13
            Call ClaimClac
    End Select
End Sub

Private Sub ClaimClac()
    If txtInput(13).Text = "0" Then
        Exit Sub
    End If

    If txtInput(13).Text = "" Then
        MsgBox "내용연수를 입력하십시요...", vbInformation
        txtInput(13).SetFocus
        Exit Sub
    End If
    
    If mskInput(0).ClipText = "" Then
        MsgBox "구입금액을 입력하십시요...", vbInformation
        mskInput(0).SetFocus
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
    
    If txtInput(13).Text <> "" And mskInput(0).ClipText <> 0 And dtInput(4).Value <> "" And _
       Val(txtInput(13).Text) <> 0 Then
        Dim iDay As Integer
        Dim hDay As Integer
        Dim bRate As Integer
        
        ' 실제경과일수 계산 (구입일자 - 입고일자)
        iDay = dtInput(1).Value - dtInput(4).Value
        txtInput(14).Text = iDay
        
        ' 환산경과일수
        hDay = iDay / Val(txtInput(13).Text)
        txtInput(15).Text = hDay
        
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
        
        txtInput(16).Text = bRate
        
        mskInput(2).Text = mskInput(0).ClipText * (bRate * 0.01)
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
    
    ' MaskEdit BOX 초기화
    For i = 0 To mskInput.Count - 1
        mskInput(i).Text = ""
    Next i
    
    ' 일자Combo BOX 초기화
    For i = 1 To dtInput.Count - 1
        dtInput(i).Value = Date
        dtInput(i).Value = ""
    Next i
End Sub
