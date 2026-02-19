VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_SMSALL_6 
   Caption         =   " SMS 발송"
   ClientHeight    =   12450
   ClientLeft      =   3870
   ClientTop       =   3135
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
   Icon            =   "P_SMSALL_006.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
      PaneTree        =   "P_SMSALL_006.frx":058A
      Begin XtremeSuiteControls.TabControl TabControl1 
         Height          =   11085
         Left            =   4080
         TabIndex        =   23
         Top             =   1350
         Width           =   12405
         _Version        =   851970
         _ExtentX        =   21881
         _ExtentY        =   19553
         _StockProps     =   68
         Appearance      =   3
         Color           =   4
         PaintManager.BoldSelected=   -1  'True
         PaintManager.OneNoteColors=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         PaintManager.LargeIcons=   -1  'True
         ItemCount       =   4
         SelectedItem    =   3
         Item(0).Caption =   "지사장"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "TabControlPage1"
         Item(1).Caption =   "가맹점"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "TabControlPage2"
         Item(2).Caption =   "협력사"
         Item(2).ControlCount=   2
         Item(2).Control(0)=   "TabControlPage3"
         Item(2).Control(1)=   "TabControlPage4"
         Item(3).Caption =   "상담자"
         Item(3).ControlCount=   3
         Item(3).Control(0)=   "TabControlPage5"
         Item(3).Control(1)=   "Command1"
         Item(3).Control(2)=   "DTPicker2"
         Begin VB.CommandButton Command1 
            Caption         =   "이후 발송자 선택 취소"
            Height          =   345
            Left            =   5310
            TabIndex        =   50
            Tag             =   "1"
            Top             =   150
            Width           =   2715
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage5 
            Height          =   10455
            Left            =   30
            TabIndex        =   44
            Top             =   600
            Width           =   12345
            _Version        =   851970
            _ExtentX        =   21775
            _ExtentY        =   18441
            _StockProps     =   1
            Page            =   4
            Begin SSSplitter.SSSplitter SSSplitter2 
               Height          =   10455
               Left            =   0
               TabIndex        =   45
               Top             =   0
               Width           =   12345
               _ExtentX        =   21775
               _ExtentY        =   18441
               _Version        =   262144
               AutoSize        =   1
               PaneTree        =   "P_SMSALL_006.frx":067C
               Begin FPSpreadADO.fpSpread spdView 
                  Height          =   10395
                  Index           =   4
                  Left            =   30
                  TabIndex        =   46
                  Top             =   30
                  Width           =   12285
                  _Version        =   524288
                  _ExtentX        =   21669
                  _ExtentY        =   18336
                  _StockProps     =   64
                  BackColorStyle  =   1
                  EditEnterAction =   4
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
                  MaxCols         =   10
                  SpreadDesigner  =   "P_SMSALL_006.frx":06AE
                  HighlightHeaders=   1
                  HighlightStyle  =   1
               End
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage4 
            Height          =   10455
            Left            =   -69970
            TabIndex        =   43
            Top             =   600
            Visible         =   0   'False
            Width           =   12345
            _Version        =   851970
            _ExtentX        =   21775
            _ExtentY        =   18441
            _StockProps     =   1
            Page            =   3
            Begin SSSplitter.SSSplitter SSSplitter3 
               Height          =   10455
               Left            =   0
               TabIndex        =   47
               Top             =   0
               Width           =   12345
               _ExtentX        =   21775
               _ExtentY        =   18441
               _Version        =   262144
               AutoSize        =   1
               PaneTree        =   "P_SMSALL_006.frx":0D92
               Begin FPSpreadADO.fpSpread spdView 
                  Height          =   10395
                  Index           =   2
                  Left            =   30
                  TabIndex        =   48
                  Top             =   30
                  Width           =   12285
                  _Version        =   524288
                  _ExtentX        =   21669
                  _ExtentY        =   18336
                  _StockProps     =   64
                  BackColorStyle  =   1
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
                  MaxCols         =   7
                  SpreadDesigner  =   "P_SMSALL_006.frx":0DC4
                  HighlightHeaders=   1
                  HighlightStyle  =   1
               End
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage3 
            Height          =   10455
            Left            =   -69970
            TabIndex        =   26
            Top             =   600
            Visible         =   0   'False
            Width           =   12345
            _Version        =   851970
            _ExtentX        =   21775
            _ExtentY        =   18441
            _StockProps     =   1
            Page            =   2
            Begin SSSplitter.SSSplitter SSSplitter1 
               Height          =   10455
               Index           =   0
               Left            =   0
               TabIndex        =   35
               Top             =   0
               Width           =   12345
               _ExtentX        =   21775
               _ExtentY        =   18441
               _Version        =   262144
               AutoSize        =   1
               PaneTree        =   "P_SMSALL_006.frx":13DD
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage2 
            Height          =   10455
            Left            =   -69970
            TabIndex        =   25
            Top             =   600
            Visible         =   0   'False
            Width           =   12345
            _Version        =   851970
            _ExtentX        =   21775
            _ExtentY        =   18441
            _StockProps     =   1
            Page            =   1
            Begin SSSplitter.SSSplitter SSSplitter1 
               Height          =   10455
               Index           =   2
               Left            =   0
               TabIndex        =   38
               Top             =   0
               Width           =   12345
               _ExtentX        =   21775
               _ExtentY        =   18441
               _Version        =   262144
               AutoSize        =   1
               PaneTree        =   "P_SMSALL_006.frx":140F
               Begin FPSpreadADO.fpSpread spdView 
                  Height          =   10395
                  Index           =   1
                  Left            =   30
                  TabIndex        =   39
                  Top             =   30
                  Width           =   4155
                  _Version        =   524288
                  _ExtentX        =   7329
                  _ExtentY        =   18336
                  _StockProps     =   64
                  BackColorStyle  =   1
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
                  MaxCols         =   3
                  SpreadDesigner  =   "P_SMSALL_006.frx":1461
                  HighlightHeaders=   1
                  HighlightStyle  =   1
               End
               Begin FPSpreadADO.fpSpread spdView 
                  Height          =   10395
                  Index           =   3
                  Left            =   4275
                  TabIndex        =   40
                  Top             =   30
                  Width           =   8040
                  _Version        =   524288
                  _ExtentX        =   14182
                  _ExtentY        =   18336
                  _StockProps     =   64
                  BackColorStyle  =   1
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
                  MaxCols         =   7
                  SpreadDesigner  =   "P_SMSALL_006.frx":19AD
                  HighlightHeaders=   1
                  HighlightStyle  =   1
               End
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage1 
            Height          =   10455
            Left            =   -69970
            TabIndex        =   24
            Top             =   600
            Visible         =   0   'False
            Width           =   12345
            _Version        =   851970
            _ExtentX        =   21775
            _ExtentY        =   18441
            _StockProps     =   1
            Page            =   0
            Begin SSSplitter.SSSplitter SSSplitter1 
               Height          =   10455
               Index           =   1
               Left            =   0
               TabIndex        =   36
               Top             =   0
               Width           =   12345
               _ExtentX        =   21775
               _ExtentY        =   18441
               _Version        =   262144
               AutoSize        =   1
               PaneTree        =   "P_SMSALL_006.frx":1FEF
               Begin FPSpreadADO.fpSpread spdView 
                  Height          =   10395
                  Index           =   0
                  Left            =   30
                  TabIndex        =   37
                  Top             =   30
                  Width           =   12285
                  _Version        =   524288
                  _ExtentX        =   21669
                  _ExtentY        =   18336
                  _StockProps     =   64
                  BackColorStyle  =   1
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
                  MaxCols         =   6
                  SpreadDesigner  =   "P_SMSALL_006.frx":2021
                  HighlightHeaders=   1
                  HighlightStyle  =   1
               End
            End
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   3750
            TabIndex        =   49
            Top             =   180
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd hh:mm"
            Format          =   64225281
            CurrentDate     =   41936
         End
      End
      Begin Threed.SSPanel panInput 
         Height          =   795
         Left            =   15
         TabIndex        =   1
         Top             =   540
         Width           =   16470
         _ExtentX        =   29051
         _ExtentY        =   1402
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin Threed.SSPanel SSPanel_Search 
            Height          =   495
            Index           =   1
            Left            =   10230
            TabIndex        =   30
            Top             =   210
            Visible         =   0   'False
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   873
            _Version        =   262144
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin XtremeSuiteControls.CheckBox chkSelect 
               Height          =   195
               Index           =   1
               Left            =   1530
               TabIndex        =   31
               Tag             =   "2"
               Top             =   150
               Width           =   1185
               _Version        =   851970
               _ExtentX        =   2090
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "유통매장"
               UseVisualStyle  =   -1  'True
               Value           =   1
            End
            Begin XtremeSuiteControls.CheckBox chkSelect 
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   32
               Tag             =   "1"
               Top             =   150
               Width           =   1185
               _Version        =   851970
               _ExtentX        =   2090
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "일반매장"
               UseVisualStyle  =   -1  'True
               Value           =   1
            End
            Begin XtremeSuiteControls.CheckBox chkSelect 
               Height          =   195
               Index           =   2
               Left            =   2940
               TabIndex        =   33
               Tag             =   "3"
               Top             =   150
               Width           =   1185
               _Version        =   851970
               _ExtentX        =   2090
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "이마트"
               UseVisualStyle  =   -1  'True
               Value           =   1
            End
            Begin XtremeSuiteControls.CheckBox chkSelect 
               Height          =   195
               Index           =   3
               Left            =   4140
               TabIndex        =   34
               Tag             =   "4"
               Top             =   150
               Width           =   1185
               _Version        =   851970
               _ExtentX        =   2090
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "크렌즈"
               UseVisualStyle  =   -1  'True
            End
         End
         Begin VB.CommandButton cmdAllCheck 
            Caption         =   "전체 선택"
            Height          =   465
            Left            =   5850
            TabIndex        =   2
            Top             =   240
            Width           =   1335
         End
         Begin Threed.SSPanel SSPanel_Search 
            Height          =   495
            Index           =   0
            Left            =   7350
            TabIndex        =   27
            Top             =   210
            Visible         =   0   'False
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   873
            _Version        =   262144
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin XtremeSuiteControls.CheckBox chkMaster 
               Height          =   195
               Index           =   0
               Left            =   240
               TabIndex        =   28
               Tag             =   "이마트"
               Top             =   150
               Width           =   1185
               _Version        =   851970
               _ExtentX        =   2090
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "영업중"
               UseVisualStyle  =   -1  'True
               Value           =   1
            End
            Begin XtremeSuiteControls.CheckBox chkMaster 
               Height          =   195
               Index           =   1
               Left            =   1440
               TabIndex        =   29
               Tag             =   "폐점"
               Top             =   150
               Width           =   1185
               _Version        =   851970
               _ExtentX        =   2090
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "폐점"
               UseVisualStyle  =   -1  'True
            End
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  '투명
            Caption         =   "보내는 사람 번호는 http://sms.uplus.co.kr/에 등록된 번호만 입력 가능 합니다. 반드시 전산실에 먼저 등록 하여 주십시요."
            ForeColor       =   &H000000FF&
            Height          =   585
            Index           =   33
            Left            =   90
            TabIndex        =   51
            Top             =   120
            Width           =   5430
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   3
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
         Caption         =   " SMS 발송 (P_SMSALL_6)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_SMSALL_006.frx":261C
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
         TabIndex        =   4
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
         PictureBackground=   "P_SMSALL_006.frx":281E
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   5
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
            Picture         =   "P_SMSALL_006.frx":2A20
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   6
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
            Picture         =   "P_SMSALL_006.frx":2FBA
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   7
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
            Picture         =   "P_SMSALL_006.frx":3554
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   8
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
            Picture         =   "P_SMSALL_006.frx":3AEE
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   9
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
            Picture         =   "P_SMSALL_006.frx":4088
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   10
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
            Picture         =   "P_SMSALL_006.frx":4622
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   11
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
            Picture         =   "P_SMSALL_006.frx":4BBC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   12
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
            Picture         =   "P_SMSALL_006.frx":5156
         End
      End
      Begin Threed.SSPanel pnlSMSMsg 
         Height          =   2475
         Index           =   0
         Left            =   15
         TabIndex        =   13
         Top             =   1350
         Width           =   4050
         _ExtentX        =   7144
         _ExtentY        =   4366
         _Version        =   262144
         BackColor       =   16777215
         PictureFrames   =   1
         Picture         =   "P_SMSALL_006.frx":56F0
         BevelOuter      =   0
         PictureAlignment=   7
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtSMS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            Height          =   1590
            Left            =   1050
            MultiLine       =   -1  'True
            TabIndex        =   14
            Top             =   570
            Width           =   1875
         End
         Begin VB.Label lblNum 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   1575
            TabIndex        =   16
            Top             =   270
            Width           =   105
         End
         Begin VB.Label lbl_SMS 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   180
            Left            =   2220
            TabIndex        =   15
            Top             =   270
            Width           =   105
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   1860
         Left            =   15
         TabIndex        =   17
         Top             =   3840
         Width           =   4050
         _ExtentX        =   7144
         _ExtentY        =   3281
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         PictureAlignment=   7
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   1440
            TabIndex        =   42
            Top             =   450
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   64225283
            CurrentDate     =   41936
         End
         Begin VB.TextBox txtRecvTel 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   18
            Top             =   45
            Width           =   2535
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   8
            Left            =   1440
            TabIndex        =   19
            Top             =   1050
            Width           =   2520
            _Version        =   851970
            _ExtentX        =   4445
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 문자메시지 보내기"
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
            Picture         =   "P_SMSALL_006.frx":13892
         End
         Begin XtremeSuiteControls.PushButton cmdBtnC 
            Height          =   630
            Index           =   9
            Left            =   150
            TabIndex        =   22
            Top             =   1050
            Width           =   1260
            _Version        =   851970
            _ExtentX        =   2222
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   "초기화"
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
         End
         Begin VB.Label lblTitle 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  '투명
            Caption         =   "예약 발송"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   330
            TabIndex        =   41
            Top             =   525
            Width           =   1170
         End
         Begin VB.Image Image 
            Height          =   240
            Index           =   2
            Left            =   90
            Picture         =   "P_SMSALL_006.frx":13F8C
            Top             =   60
            Width           =   240
         End
         Begin VB.Label lblTitle 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  '투명
            Caption         =   "보내는 사람"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   20
            Top             =   120
            Width           =   1170
         End
      End
      Begin FPSpreadADO.fpSpread spdUser 
         Height          =   6720
         Left            =   15
         TabIndex        =   21
         Top             =   5715
         Width           =   4050
         _Version        =   524288
         _ExtentX        =   7144
         _ExtentY        =   11853
         _StockProps     =   64
         BackColorStyle  =   1
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
         MaxCols         =   1
         SpreadDesigner  =   "P_SMSALL_006.frx":14316
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_SMSALL_6"
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

 

Private Sub cmdAllCheck_Click()
    Dim nRow    As Long
    Dim bMode   As Integer
    Dim vText   As Variant
    Dim sTel(2) As String
    
    bMode = IIf(cmdAllCheck.Caption = "전체 선택", "1", "0")
    cmdAllCheck.Caption = IIf(cmdAllCheck.Caption = "전체 선택", "전체 취소", "전체 선택")
    
    ' 지사장
    If TabControl1.SelectedItem = 0 Then
        
        With spdView(0)
            .EventEnabled(EventButtonClicked) = False
            
            For nRow = 1 To .MaxRows
                .GetText 6, nRow, vText
                
                If CheckMobileNumber(CStr(vText), sTel) = True Then
                    .SetText 1, nRow, CVar(bMode)
                End If
            
            Next nRow
            .EventEnabled(EventButtonClicked) = True
        End With
        
    ' 가맹점
    ElseIf TabControl1.SelectedItem = 1 Then
        With spdView(3)
            .EventEnabled(EventButtonClicked) = False
            
            For nRow = 1 To .MaxRows
                .GetText 6, nRow, vText
                
                If CheckMobileNumber(CStr(vText), sTel) = True Then
                    .SetText 1, nRow, CVar(bMode)
                End If
            
            Next nRow
            .EventEnabled(EventButtonClicked) = True
        End With
        
    ' 협력사
    ElseIf TabControl1.SelectedItem = 2 Then
        
        With spdView(2)
            .EventEnabled(EventButtonClicked) = False
            
            For nRow = 1 To .MaxRows
                .GetText 6, nRow, vText
                
                If CheckMobileNumber(CStr(vText), sTel) = True Then
                    .SetText 1, nRow, CVar(bMode)
                End If
            
            Next nRow
            .EventEnabled(EventButtonClicked) = True
        End With
    
    
    ' 상담사
    ElseIf TabControl1.SelectedItem = 3 Then
        
        With spdView(4)
            .EventEnabled(EventButtonClicked) = False
            
            For nRow = 1 To .MaxRows
                .GetText 6, nRow, vText
                
                If CheckMobileNumber(CStr(vText), sTel) = True Then
                    .SetText 1, nRow, CVar(bMode)
                End If
            
            Next nRow
            .EventEnabled(EventButtonClicked) = True
        End With
    End If
        
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 0
            If TabControl1.SelectedItem = 0 Then
                Call Data_Display_Master
                
            ElseIf TabControl1.SelectedItem = 1 Then
                Call Data_Display_Store
                
            ElseIf TabControl1.SelectedItem = 2 Then
                Call Data_Display_ETC
            
            ElseIf TabControl1.SelectedItem = 3 Then
                Call Data_Display_COU
            End If
        Case 1:                 ' 신규
        Case 2:                 ' 저장
        Case 3: ' 삭제
        Case 4:            ' 취소
        Case 5:            ' 인쇄
        Case 6:            ' 화면
        Case 7: Unload Me  ' 종료
        Case 8: SendSMS    ' sms 발송
        
        Case Else
            '
    End Select

End Sub

Private Sub cmdBtnC_Click(Index As Integer)
    Dim nRow    As Long
    
    For nRow = 1 To spdUser.MaxRows
        spdUser.SetText 1, nRow, CVar("")
    Next nRow
    
End Sub

Private Sub Command1_Click()
    Dim nRow    As Integer
    Dim vText   As Variant
    
    With spdView(4)
        For nRow = 1 To .MaxRows
            .GetText 10, nRow, vText
            
            If Mid(CStr(vText), 1, 10) >= Format(DTPicker2.Value, "yyyy-MM-dd") Then
                .SetText 1, nRow, "0"
            End If
        Next nRow
        
    End With

End Sub

Private Sub Form_Activate()
    Dim nRow    As Long
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"

    Call SubBottonEnable(cmdBtn, "10000001")
    
    
    
    If P_SMSALL_6_Flag = False Then
        Screen.MousePointer = vbHourglass
        
        DTPicker1.CheckBox = True
        DTPicker1.Value = Now
        
        txtRecvTel.Text = Trim(GetIniStr("Order Setting", "P_SMSALL_6", "031-522-2000", m_iniFile))
        txtSMS.Text = Trim(GetIniStr("Order Setting", "P_SMSALL_6_S", " ", m_iniFile))
        
        ReDim sValue(0)
        
        sValue(0) = "0"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_M_SMSALL_006_03", sValue(), Err_Num, Err_Dec)
        
        spdView(1).MaxRows = RS01.RecordCount
        Call fpSpread_Display(spdView(1), RS01)
        
        P_SMSALL_6_Flag = True
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Form_Load()

    TabControl1.SelectedItem = 0
    SSPanel_Search(0).Visible = True
    SSPanel_Search(1).Move SSPanel_Search(0).Left, SSPanel_Search(0).Top
    
    DTPicker2.Value = Date

    With spdView(0)
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        .Col = 1
        .Row = 0
        .CellType = CellTypeCheckBox
        .BackColor = &HE0E0E0
        .TypeCheckCenter = True
        .Value = False
    
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeNormal
        
        'Init the User Sort
        '.UserColAction = UserColActionSort
    End With

    With spdView(1)
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        .Col = 1
        .Row = 0
        .CellType = CellTypeCheckBox
        .BackColor = &HE0E0E0
        .TypeCheckCenter = True
        .Value = False
    
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeNormal
        
        'Init the User Sort
        '.UserColAction = UserColActionSort
    End With

    With spdView(2)
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        .Col = 1
        .Row = 0
        .CellType = CellTypeCheckBox
        .BackColor = &HE0E0E0
        .TypeCheckCenter = True
        .Value = False
    
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeNormal
        
        'Init the User Sort
        '.UserColAction = UserColActionSort
    End With

    With spdView(3)
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        .Col = 1
        .Row = 0
        .CellType = CellTypeCheckBox
        .BackColor = &HE0E0E0
        .TypeCheckCenter = True
        .Value = False
    
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeNormal
        
        'Init the User Sort
        '.UserColAction = UserColActionSort
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    cmdBtn(0).Enabled = False
    cmdBtn(1).Enabled = False
    cmdBtn(2).Enabled = False
    cmdBtn(3).Enabled = False
    cmdBtn(4).Enabled = False
    cmdBtn(5).Enabled = False
    cmdBtn(6).Enabled = False
    
    Call SetIniStr("Order Setting", "P_SMSALL_6", Trim(txtRecvTel.Text), m_iniFile)
    Call SetIniStr("Order Setting", "P_SMSALL_6_S", Trim(txtSMS.Text), m_iniFile)
    
    P_SMSALL_6_Flag = False
End Sub

Public Sub DataSave()

End Sub

Public Sub DataAdd()

End Sub

Private Sub Data_Display_Master()
    On Error GoTo ErrRtn

    Dim nRow As Long
    
    ReDim sValue(0)
    Dim sTel(2) As String
    
    If chkMaster(0).Value = xtpChecked And chkMaster(1).Value = xtpChecked Then
        sValue(0) = "%"
    ElseIf chkMaster(0).Value = xtpChecked Then
        sValue(0) = "Y"
    ElseIf chkMaster(1).Value = xtpChecked Then
        sValue(0) = "N"
    End If
  
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_SMSALL_006_00", sValue(), Err_Num, Err_Dec)
    
    With spdView(0)
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = IIf(CheckMobileNumber(CStr(Trim(RS01!휴대전화 & "")), sTel) = True, "1", "0")
            .Col = 2:  .Text = RS01!지사코드 & ""
            .Col = 3:  .Text = RS01!지사명 & ""
            .Col = 4:  .Text = RS01!지사상태 & ""
            .Col = 5:  .Text = Trim(RS01!지사장명 & "")
            .Col = 6:  .Text = Trim(RS01!휴대전화 & "")
            
            RS01.MoveNext
        Loop
        RS01.Close: Set RS01 = Nothing
        .Redraw = True
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Data_Display_Store()
    On Error GoTo ErrRtn

    Dim nRow As Long
    Dim sTel(2) As String
    
    ReDim sValue(1)
    
    ' 선택된 지사 코드를 리턴한다.
    sValue(0) = GetSelectMasterCodeList(spdView(1), 1)
  
    If sValue(0) = "" Then Exit Sub
    
    sValue(1) = ""
    If chkSelect(0).Value = xtpChecked Then sValue(1) = sValue(1) & "1,"
    If chkSelect(1).Value = xtpChecked Then sValue(1) = sValue(1) & "2,"
    If chkSelect(2).Value = xtpChecked Then sValue(1) = sValue(1) & "3,"
    If chkSelect(3).Value = xtpChecked Then sValue(1) = sValue(1) & "4,"
    ' 마지막 ,를 삭제한다.
    If Len(sValue(1)) > 1 Then sValue(1) = Mid(sValue(1), 1, Len(sValue(1)) - 1)
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_SMSALL_006_04", sValue(), Err_Num, Err_Dec)
    
    With spdView(3)
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = IIf(CheckMobileNumber(CStr(Trim(RS01!휴대전화번호 & "")), sTel) = True, "1", "0")
            .Col = 2:  .Text = RS01!지사코드 & ""
            .Col = 3:  .Text = Trim(RS01!가맹점코드 & "")
            .Col = 4:  .Text = Trim(RS01!가맹점명 & "")
            .Col = 5:  .Text = Trim(RS01!대표자명 & "")
            .Col = 6:  .Text = Trim(RS01!휴대전화번호 & "")
            .Col = 5:  .Text = Trim(RS01!매장전화번호 & "")
            
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

Private Sub Data_Display_ETC()
    On Error GoTo ErrRtn

    Dim nRow As Long
    Dim sTel(2) As String
    
    ReDim sValue(0)
    
    sValue(0) = Store.Code
  
    If sValue(0) = "" Then Exit Sub
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_SMSALL_006_02", sValue(), Err_Num, Err_Dec)
    
    
    With spdView(2)
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = IIf(CheckMobileNumber(CStr(Trim(RS01!연락처 & "")), sTel) = True, "1", "0")
            .Col = 2:  .Text = RS01!담당자코드 & ""
            .Col = 3:  .Text = RS01!구분 & ""
            .Col = 4:  .Text = RS01!매장명 & ""
            .Col = 5:  .Text = Trim(RS01!성명 & "")
            .Col = 6:  .Text = Trim(RS01!연락처 & "")
            .Col = 7:  .Text = Trim(RS01!비고 & "")
            
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


Private Sub Data_Display_COU()
    On Error GoTo ErrRtn

    Dim nRow As Long
    Dim sTel(2) As String
    
    ReDim sValue(1)
    
    sValue(0) = Store.Code
    sValue(1) = "02"
  
    If sValue(0) = "" Then Exit Sub
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_SMSALL_005_00", sValue(), Err_Num, Err_Dec)
    
    
    With spdView(4)
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = IIf(CheckMobileNumber(CStr(Trim(RS01!연락처 & "")), sTel) = True, "1", "0")
            .Col = 2:  .Text = RS01!코드 & ""
            .Col = 3:  .Text = RS01!지역 & ""
            .Col = 4:  .Text = RS01!최초상담 & ""
            .Col = 5:  .Text = Trim(RS01!성명 & "")
            .Col = 6:  .Text = Trim(RS01!연락처 & "")
            .Col = 7:  .Text = Trim(RS01!점포상황주소 & "")
            .Col = 8:  .Text = Trim(RS01!전화상담자 & "")
            .Col = 9:  .Text = Trim(RS01!비고 & "")
            .Col = 10:  .Text = Trim(RS01!최종전송일자 & "")
            
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

Private Sub SendSMS()
    Dim dLng        As Long
    Dim sRecvTel() As String
    Dim Query       As String

    ' 설정 문자 저장
    Call SetIniStr("Order Setting", "P_SMSALL_6_S", Trim(txtSMS.Text), m_iniFile)
    
    ' 문자 메시지 길이 확인
    dLng = CheckSendMessageLangth
    
    If dLng <= 0 Or dLng > m_SMS_Lng Then Exit Sub
    
    ' 발신자 번호 확인
    If GetCheckSMSSendTel(txtRecvTel.Text, sRecvTel, True) = False Then
        MsgBox txtRecvTel.Text & " 보내는 사람 전화 번호로를 확인하여 주십시요.", vbInformation, "확인"
        
        txtRecvTel.SetFocus
        txtRecvTel.SelStart = 0: txtRecvTel.SelLength = Len(txtRecvTel.Text)
        Exit Sub
    End If
    
    ' 최종 확인 메시지
    Query = "메시지를 전송 하시겠습니까? "
    Rtn = MsgBox(Query, vbInformation + vbYesNo + vbDefaultButton2, "확인")
       
    If Rtn = vbNo Then Exit Sub


    ' 개별 보내는 내용이 있을 경우
    If spdUser.DataRowCnt > 0 Then
        Call SendSMS_USER
        Exit Sub
        
    ' 지사장
    ElseIf TabControl1.SelectedItem = 0 Then
    
        Call SendSMS_Master
        Exit Sub
        
    ' 가맹점
    ElseIf TabControl1.SelectedItem = 1 Then
    
        Call SendSMS_Store
        Exit Sub
        
    ' 협력사
    ElseIf TabControl1.SelectedItem = 2 Then
    
        Call SendSMS_ETC
        Exit Sub
    
    ' 상담자
    ElseIf TabControl1.SelectedItem = 3 Then
    
        Call SendSMS_COU
        Exit Sub
    End If

End Sub


'--------------------------------------------------------------------------------------------------------------
' Procedure : SendSMS
' DateTime  : 2007-05-06 23:16
' Author    : pds2004
' Purpose   : SMS 문자 메시지 전송
'--------------------------------------------------------------------------------------------------------------
Private Function SendSMS_USER() As Boolean
    Dim nRow        As Long
    Dim sValue(11)  As String
    Dim vTemp       As Variant
    Dim nSendCnt    As Long
    Dim Query       As String
    Dim sRecvTel(2) As String
    Dim sSendTel(2) As String
    
    
    
    On Error GoTo ErrRtn
     
    nSendCnt = 0
    
    If Not IsNull(DTPicker1.Value) Then
         If Format(DTPicker1.Value, "yyyy-MM-dd hh:mm:ss") <= Now Then
            MsgBox "예약 전송 시간을 확인하여 주십시요.", vbInformation, "확인"
            Exit Function
         End If
    End If
    
    
    With spdUser
        For nRow = 1 To .MaxRows
            .GetText 1, nRow, vTemp
            If CheckMobileNumber(CStr(vTemp), sSendTel) = True Then
                           
                sValue(0) = "1"                      '전송
                sValue(1) = "0"                      '메시지타입
                sValue(2) = sSendTel(0) & sSendTel(1) & sSendTel(2) '수신번호
                sValue(3) = Trim(txtRecvTel.Text)    '발신번호
                sValue(4) = Trim(txtSMS.Text)        '메시지
                sValue(5) = Store.Code               '지사코드
                sValue(6) = "000"                    '가맹점코드
                sValue(7) = " "                      '고객코드
                sValue(8) = " "                      '고객성명
                sValue(9) = "000000"                 '참고5
                sValue(10) = "1"                     '참고6
                sValue(11) = IIf(IsNull(DTPicker1.Value), "", Format(DTPicker1.Value, "yyyy-MM-dd hh:mm:ss"))
                
                Call ExecPro("PRO_SMS_SEND_MASTER", sValue(), Err_Num, Err_Dec)
                
                If Err_Num <> 0 Then
                    MsgBox Err_Dec, vbCritical, "오류"
                    Exit Function
                End If
                
                nSendCnt = nSendCnt + 1
            
            End If
        Next nRow
        
    End With
    
    If nSendCnt > 0 Then MsgBox "[" & CStr(nSendCnt) & "] 건을 발송 하였습니다.", vbInformation, "확인"
    
    On Error GoTo 0
    Exit Function

ErrRtn:
    ' 최종 남은 수량을 설정한다.
    
    DoEvents
    
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Function


'--------------------------------------------------------------------------------------------------------------
' Procedure : SendSMS
' DateTime  : 2007-05-06 23:16
' Author    : pds2004
' Purpose   : SMS 문자 메시지 전송
'--------------------------------------------------------------------------------------------------------------
Private Function SendSMS_Master() As Boolean
    Dim nRow        As Long
    Dim dLng        As Long
    Dim sValue(11)  As String
    Dim vTemp       As Variant
    Dim nSendCnt    As Long
    Dim Query       As String
    Dim sRecvTel(2) As String
    Dim sSendTel(2) As String
    
    On Error GoTo ErrRtn
    
    nSendCnt = 0
    If Not IsNull(DTPicker1.Value) Then
         If Format(DTPicker1.Value, "yyyy-MM-dd hh:mm:ss") <= Now Then
            MsgBox "예약 전송 시간을 확인하여 주십시요.", vbInformation, "확인"
            Exit Function
         End If
    End If
    
    
    With spdView(0)
        For nRow = 1 To .MaxRows
            
            ' 전송 구분일 경우 전송 처리한다.
            .GetText 1, nRow, vTemp
            
            If CStr(vTemp) = "1" Then
                .GetText 6, nRow, vTemp
                If CheckMobileNumber(CStr(vTemp), sSendTel) = True Then
                               
                    sValue(0) = "1"                      '전송
                    sValue(1) = "0"                      '메시지타입
                    sValue(2) = sSendTel(0) & sSendTel(1) & sSendTel(2) '수신번호
                    sValue(3) = Trim(txtRecvTel.Text)    '발신번호
                    sValue(4) = Trim(txtSMS.Text)        '메시지
                    sValue(5) = Store.Code               '지사코드
                    sValue(6) = "000"                    '가맹점코드
                    sValue(7) = " "                      '고객코드
                    
                    .GetText 5, nRow, vTemp
                    sValue(8) = CStr(vTemp)              '고객성명
                    
                    sValue(9) = "000000"                 '참고5
                    sValue(10) = "1"                     '참고6
                    sValue(11) = IIf(IsNull(DTPicker1.Value), "", Format(DTPicker1.Value, "yyyy-MM-dd hh:mm:ss"))
                    
                    Call ExecPro("PRO_SMS_SEND_MASTER", sValue(), Err_Num, Err_Dec)
                    
                    If Err_Num <> 0 Then
                        MsgBox Err_Dec, vbCritical, "오류"
                        Exit Function
                    End If
                    
                    nSendCnt = nSendCnt + 1
                    
                    Call .SetText(1, nRow, CVar("0"))
            
                End If
            End If
        Next nRow
        
    End With
    
    If nSendCnt > 0 Then MsgBox "[" & CStr(nSendCnt) & "] 건을 발송 하였습니다.", vbInformation, "확인"
    
    On Error GoTo 0
    Exit Function

ErrRtn:
    ' 최종 남은 수량을 설정한다.
    
    DoEvents
    
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Function

'--------------------------------------------------------------------------------------------------------------
' Procedure : SendSMS
' DateTime  : 2007-05-06 23:16
' Author    : pds2004
' Purpose   : SMS 문자 메시지 전송
'--------------------------------------------------------------------------------------------------------------
Private Function SendSMS_Store() As Boolean
    Dim nRow        As Long
    Dim dLng        As Long
    Dim sValue(11)  As String
    Dim vTemp       As Variant
    Dim nSendCnt    As Long
    Dim Query       As String
    Dim sRecvTel(2) As String
    Dim sSendTel(2) As String
    
    On Error GoTo ErrRtn
    
    nSendCnt = 0
    If Not IsNull(DTPicker1.Value) Then
         If Format(DTPicker1.Value, "yyyy-MM-dd hh:mm:ss") <= Now Then
            MsgBox "예약 전송 시간을 확인하여 주십시요.", vbInformation, "확인"
            Exit Function
         End If
    End If
    
    
    With spdView(3)
        For nRow = 1 To .MaxRows
            
            ' 전송 구분일 경우 전송 처리한다.
            .GetText 1, nRow, vTemp
            If CStr(vTemp) = "1" Then
                .GetText 6, nRow, vTemp
                If CheckMobileNumber(CStr(vTemp), sSendTel) = True Then
                               
                    sValue(0) = "1"                      '전송
                    sValue(1) = "0"                      '메시지타입
                    sValue(2) = sSendTel(0) & sSendTel(1) & sSendTel(2) '수신번호
                    sValue(3) = Trim(txtRecvTel.Text)    '발신번호
                    sValue(4) = Trim(txtSMS.Text)        '메시지
                    sValue(5) = Store.Code               '지사코드
                    sValue(6) = "000"                    '가맹점코드
                    sValue(7) = " "                      '고객코드
                    
                    .GetText 5, nRow, vTemp
                    sValue(8) = CStr(vTemp)              '고객성명
                    
                    sValue(9) = "000000"                 '참고5
                    sValue(10) = "1"                     '참고6
                    sValue(11) = IIf(IsNull(DTPicker1.Value), "", Format(DTPicker1.Value, "yyyy-MM-dd hh:mm:ss"))
                    
                    Call ExecPro("PRO_SMS_SEND_MASTER", sValue(), Err_Num, Err_Dec)
                    
                    If Err_Num <> 0 Then
                        MsgBox Err_Dec, vbCritical, "오류"
                        Exit Function
                    End If
                    
                    nSendCnt = nSendCnt + 1
                    
                    Call .SetText(1, nRow, CVar("0"))
            
                End If
            End If
        Next nRow
        
    End With
    
    If nSendCnt > 0 Then MsgBox "[" & CStr(nSendCnt) & "] 건을 발송 하였습니다.", vbInformation, "확인"
    
    On Error GoTo 0
    Exit Function

ErrRtn:
    ' 최종 남은 수량을 설정한다.
    
    DoEvents
    
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Function
'--------------------------------------------------------------------------------------------------------------
' Procedure : SendSMS
' DateTime  : 2007-05-06 23:16
' Author    : pds2004
' Purpose   : SMS 문자 메시지 전송
'--------------------------------------------------------------------------------------------------------------
Private Function SendSMS_ETC() As Boolean
    Dim nRow        As Long
    Dim dLng        As Long
    Dim sValue(11)  As String
    Dim vTemp       As Variant
    Dim nSendCnt    As Long
    Dim Query       As String
    Dim sRecvTel(2) As String
    Dim sSendTel(2) As String
    
    On Error GoTo ErrRtn
     
    nSendCnt = 0
    If Not IsNull(DTPicker1.Value) Then
         If Format(DTPicker1.Value, "yyyy-MM-dd hh:mm:ss") <= Now Then
            MsgBox "예약 전송 시간을 확인하여 주십시요.", vbInformation, "확인"
            Exit Function
         End If
    End If
    
    With spdView(2)
        For nRow = 1 To .MaxRows
            ' 전송 구분일 경우 전송 처리한다.
            .GetText 1, nRow, vTemp
            If CStr(vTemp) = "1" Then
            
                .GetText 6, nRow, vTemp
                If CheckMobileNumber(CStr(vTemp), sSendTel) = True Then
                    
                    sValue(0) = "1"                      '전송
                    sValue(1) = "0"                      '메시지타입
                    sValue(2) = sSendTel(0) & sSendTel(1) & sSendTel(2) '수신번호
                    sValue(3) = Trim(txtRecvTel.Text)    '발신번호
                    sValue(4) = Trim(txtSMS.Text)        '메시지
                    sValue(5) = Store.Code               '지사코드
                    sValue(6) = "000"                    '가맹점코드
                    sValue(7) = " "                      '고객코드
                    
                    .GetText 5, nRow, vTemp
                    sValue(8) = CStr(vTemp)              '고객성명
                    
                    sValue(9) = "000000"                 '참고5
                    sValue(10) = "1"                     '참고6
                    sValue(11) = IIf(IsNull(DTPicker1.Value), "", Format(DTPicker1.Value, "yyyy-MM-dd hh:mm:ss"))
                    
                    Call ExecPro("PRO_SMS_SEND_MASTER", sValue(), Err_Num, Err_Dec)
                    
                    If Err_Num <> 0 Then
                        MsgBox Err_Dec, vbCritical, "오류"
                        Exit Function
                    End If
                    
                    nSendCnt = nSendCnt + 1
                    Call .SetText(1, nRow, CVar("0"))
            
                End If
            End If
        Next nRow
        
    End With
    
    If nSendCnt > 0 Then MsgBox "[" & CStr(nSendCnt) & "] 건을 발송 하였습니다.", vbInformation, "확인"
    
    On Error GoTo 0
    Exit Function

ErrRtn:
    ' 최종 남은 수량을 설정한다.
    
    DoEvents
    
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Function

'--------------------------------------------------------------------------------------------------------------
' Procedure : SendSMS
' DateTime  : 2007-05-06 23:16
' Author    : pds2004
' Purpose   : SMS 문자 메시지 전송
'--------------------------------------------------------------------------------------------------------------
Private Function SendSMS_COU() As Boolean
    Dim nRow        As Long
    Dim dLng        As Long
    Dim sValue(11)  As String
    Dim vTemp       As Variant
    Dim nSendCnt    As Long
    Dim Query       As String
    Dim sRecvTel(2) As String
    Dim sSendTel(2) As String
    
    On Error GoTo ErrRtn
     
    nSendCnt = 0
    If Not IsNull(DTPicker1.Value) Then
         If Format(DTPicker1.Value, "yyyy-MM-dd hh:mm:ss") <= Now Then
            MsgBox "예약 전송 시간을 확인하여 주십시요.", vbInformation, "확인"
            Exit Function
         End If
    End If
    
    With spdView(4)
        For nRow = 1 To .MaxRows
            ' 전송 구분일 경우 전송 처리한다.
            .GetText 1, nRow, vTemp
            If CStr(vTemp) = "1" Then
            
                .GetText 6, nRow, vTemp
                If CheckMobileNumber(CStr(vTemp), sSendTel) = True Then
                    
                    sValue(0) = "1"                      '전송
                    sValue(1) = "0"                      '메시지타입
                    sValue(2) = sSendTel(0) & sSendTel(1) & sSendTel(2) '수신번호
                    sValue(3) = Trim(txtRecvTel.Text)    '발신번호
                    sValue(4) = Trim(txtSMS.Text)        '메시지
                    sValue(5) = Store.Code               '지사코드
                    sValue(6) = "000"                    '가맹점코드
                    sValue(7) = " "                      '고객코드
                    
                    .GetText 5, nRow, vTemp
                    sValue(8) = CStr(vTemp)              '고객성명
                    
                    sValue(9) = "000000"                 '참고5
                    sValue(10) = "11"                     '참고6
                    sValue(11) = IIf(IsNull(DTPicker1.Value), "", Format(DTPicker1.Value, "yyyy-MM-dd hh:mm:ss"))
                    
                    Call ExecPro("PRO_SMS_SEND_MASTER", sValue(), Err_Num, Err_Dec)
                    
                    If Err_Num <> 0 Then
                        MsgBox Err_Dec, vbCritical, "오류"
                        Exit Function
                    End If
                    
                    nSendCnt = nSendCnt + 1
                    Call .SetText(1, nRow, CVar("0"))
            
                End If
            End If
        Next nRow
        
    End With
    
    If nSendCnt > 0 Then MsgBox "[" & CStr(nSendCnt) & "] 건을 발송 하였습니다.", vbInformation, "확인"
    
    On Error GoTo 0
    Exit Function

ErrRtn:
    ' 최종 남은 수량을 설정한다.
    
    DoEvents
    
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Function

Private Sub spdView_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)

Dim nRow As Long
Dim vText As Variant
Dim sTel(2) As String
Dim varGet As Variant

 
    
    ' 지사
    If Index = 0 And Row = 0 And Col = 1 Then
        With spdView(0)
            '.EventEnabled(EventButtonClicked) = False
            
            .GetText 1, 0, varGet
            .SetText 1, 0, CVar(IIf(CStr(varGet) = "1", "0", "1"))
            
            For nRow = 1 To .MaxRows
                .GetText 6, nRow, vText
                
                If CheckMobileNumber(CStr(vText), sTel) = True Then
                    .SetText 1, nRow, CVar(IIf(CStr(varGet) = "1", "0", "1"))
                End If
            
            Next nRow
            '.EventEnabled(EventButtonClicked) = True
        End With
        Exit Sub
        
    ' 가맹점
    ElseIf Index = 1 And Row = 0 And Col = 1 Then
        With spdView(1)
            '.EventEnabled(EventButtonClicked) = False
            
            .GetText 1, 0, varGet
            
            .Col = 1
            .Col2 = 1
            .Row = 0
            .Row2 = .MaxRows
            .BlockMode = True
            .Value = IIf(CStr(varGet) = "1", "0", "1")
            .BlockMode = False
            
            '.EventEnabled(EventButtonClicked) = True
        End With
        Exit Sub
    
    ' 가맹점
    ElseIf Index = 3 And Row = 0 And Col = 1 Then
        With spdView(3)
            .EventEnabled(EventButtonClicked) = False
            
            .GetText 1, 0, varGet
            .SetText 1, 0, CVar(IIf(CStr(varGet) = "1", "0", "1"))
            
            For nRow = 1 To .MaxRows
                .GetText 6, nRow, vText
                
                If CheckMobileNumber(CStr(vText), sTel) = True Then
                    .SetText 1, nRow, CVar(IIf(CStr(varGet) = "1", "0", "1"))
                End If
            
            Next nRow
            .EventEnabled(EventButtonClicked) = True
        End With
        Exit Sub
    
    ' 협력사
    ElseIf Index = 2 And Row = 0 And Col = 1 Then
        With spdView(2)
            .EventEnabled(EventButtonClicked) = False
            
            .GetText 1, 0, varGet
            .SetText 1, 0, CVar(IIf(CStr(varGet) = "1", "0", "1"))
            
            For nRow = 1 To .MaxRows
                .GetText 6, nRow, vText
                
                If CheckMobileNumber(CStr(vText), sTel) = True Then
                    .SetText 1, nRow, CVar(IIf(CStr(varGet) = "1", "0", "1"))
                End If
            
            Next nRow
            .EventEnabled(EventButtonClicked) = True
        End With
        Exit Sub
    
    End If

End Sub

Private Sub TabControl1_BeforeItemClick(ByVal Item As XtremeSuiteControls.ITabControlItem, Cancel As Variant)
    Select Case Item.Index
        Case 0
            SSPanel_Search(0).Visible = True
            SSPanel_Search(1).Visible = False
            SSPanel_Search(0).ZOrder 0
            
        Case 1
            SSPanel_Search(0).Visible = False
            SSPanel_Search(1).Visible = True
            SSPanel_Search(1).ZOrder 0
            
        Case 2
            SSPanel_Search(0).Visible = False
            SSPanel_Search(1).Visible = False
            
        Case 3
            SSPanel_Search(0).Visible = False
            SSPanel_Search(1).Visible = False
            
    End Select

End Sub

Private Sub txtSMS_Change()
    lbl_SMS.Tag = CStr(LenB(StrConv(txtSMS.Text, vbFromUnicode)))
    lbl_SMS.Caption = lbl_SMS.Tag & "자"
    Debug.Print lbl_SMS.Tag & "자"
    
    If LenB(StrConv(txtSMS.Text, vbFromUnicode)) > m_SMS_Lng Then
        lbl_SMS.BackColor = vbRed
        MsgBox "작성된 메시지가 " & CStr(m_SMS_Lng) & " 자 이상 입니다. " & CStr(m_SMS_Lng) & "자 이상은 전송할 수 없습니다.", vbCritical, "확인"
        Exit Sub
    Else
        lbl_SMS.BackColor = Me.BackColor
    End If

End Sub


Private Function CheckSendMessageLangth() As Integer
    If IsNumeric(lbl_SMS.Tag) = False Then
        CheckSendMessageLangth = 0
        txtSMS.SetFocus
        MsgBox "전송할 메시지를 입력 하여 주십시요..  [" & CStr(Val(lbl_SMS.Tag)) & "자]", vbInformation, "확인"
        Exit Function
        
    ElseIf Val(lbl_SMS.Tag) > m_SMS_Lng Then
        CheckSendMessageLangth = Val(lbl_SMS.Tag)
        txtSMS.SetFocus
        MsgBox "전송할 메시지를 확인 하여 주십시요..  [" & CStr(Val(lbl_SMS.Tag)) & "자]", vbInformation, "확인"
        Exit Function
    Else
        CheckSendMessageLangth = Val(lbl_SMS.Tag)
        Exit Function
    End If
End Function




 
