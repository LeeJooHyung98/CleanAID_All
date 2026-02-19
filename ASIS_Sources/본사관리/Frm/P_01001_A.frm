VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_01001_A 
   Caption         =   "가맹점 등록"
   ClientHeight    =   11010
   ClientLeft      =   105
   ClientTop       =   3045
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_01001_A.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   11010
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   19420
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_01001_A.frx":058A
      Begin Threed.SSPanel SSPanel1 
         Height          =   4605
         Left            =   15
         TabIndex        =   2
         Top             =   6390
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   8123
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   12
            Left            =   1740
            MaxLength       =   3
            TabIndex        =   11
            Top             =   4230
            Width           =   675
         End
         Begin VB.TextBox txtInput 
            Height          =   345
            Index           =   11
            Left            =   60
            MaxLength       =   255
            TabIndex        =   10
            Top             =   3840
            Width           =   6195
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   10
            Left            =   1740
            MaxLength       =   50
            TabIndex        =   9
            Top             =   3465
            Width           =   1185
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   9
            Left            =   1740
            MaxLength       =   50
            TabIndex        =   8
            Top             =   3090
            Width           =   4515
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   5
            Left            =   1740
            MaxLength       =   50
            TabIndex        =   7
            Top             =   1950
            Width           =   975
         End
         Begin VB.TextBox txtInput 
            Enabled         =   0   'False
            Height          =   315
            Index           =   4
            Left            =   2775
            MaxLength       =   6
            TabIndex        =   6
            Top             =   1590
            Width           =   2685
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   3
            Left            =   1740
            MaxLength       =   50
            TabIndex        =   5
            Top             =   1590
            Width           =   1005
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   2
            Left            =   1740
            MaxLength       =   50
            TabIndex        =   4
            Top             =   450
            Width           =   3735
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   1
            Left            =   1740
            MaxLength       =   6
            TabIndex        =   3
            Top             =   60
            Width           =   1335
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   60
            TabIndex        =   12
            Top             =   60
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "가맹점코드"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   13
            Top             =   450
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "가맹점명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   32
            Left            =   60
            TabIndex        =   14
            Top             =   840
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "가맹점상태"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   3
            Left            =   60
            TabIndex        =   15
            Top             =   1590
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "사 업 장"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   4
            Left            =   60
            TabIndex        =   16
            Top             =   1950
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "택 번 호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   25
            Left            =   60
            TabIndex        =   17
            Top             =   1230
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "시작일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1740
            TabIndex        =   18
            Top             =   1230
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   56885248
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   5
            Left            =   60
            TabIndex        =   19
            Top             =   2340
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "오픈일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   6
            Left            =   60
            TabIndex        =   20
            Top             =   2730
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "폐점일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   7
            Left            =   60
            TabIndex        =   21
            Top             =   3090
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "전화번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   8
            Left            =   60
            TabIndex        =   22
            Top             =   3480
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "우편번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   9
            Left            =   60
            TabIndex        =   23
            Top             =   4230
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "가맹점마진"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   3
            Left            =   1740
            TabIndex        =   24
            Top             =   2325
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   56885248
            CurrentDate     =   36686
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   4
            Left            =   1740
            TabIndex        =   25
            Top             =   2730
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   56885248
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   10
            Left            =   2445
            TabIndex        =   26
            Top             =   4245
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "%"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   33
            Left            =   1740
            TabIndex        =   27
            Top             =   840
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   262144
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   1
               Left            =   2100
               TabIndex        =   28
               Top             =   30
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "폐점"
            End
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   0
               Left            =   360
               TabIndex        =   29
               Top             =   30
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "개점"
            End
         End
         Begin Threed.SSFrame SSFrame1 
            Height          =   4440
            Index           =   1
            Left            =   6585
            TabIndex        =   30
            Top             =   105
            Width           =   6675
            _ExtentX        =   11774
            _ExtentY        =   7832
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
            Caption         =   "사업장 택 사용현황"
            Begin VB.CommandButton cmdSub 
               Caption         =   "삭제"
               Height          =   345
               Index           =   1
               Left            =   5730
               TabIndex        =   35
               Top             =   3975
               Width           =   825
            End
            Begin VB.CommandButton cmdSub 
               Caption         =   "적용"
               Height          =   345
               Index           =   0
               Left            =   5730
               TabIndex        =   34
               Top             =   3585
               Width           =   825
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   8
               Left            =   1800
               MaxLength       =   50
               TabIndex        =   33
               Top             =   3615
               Width           =   795
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   6
               Left            =   1800
               MaxLength       =   6
               TabIndex        =   32
               Top             =   3255
               Width           =   1200
            End
            Begin VB.TextBox txtInput 
               Enabled         =   0   'False
               Height          =   315
               Index           =   7
               Left            =   3030
               MaxLength       =   50
               TabIndex        =   31
               Top             =   3255
               Width           =   2475
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   36
               Left            =   120
               TabIndex        =   36
               Top             =   3270
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "사업장"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   37
               Left            =   120
               TabIndex        =   37
               Top             =   3630
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "택번호"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   38
               Left            =   120
               TabIndex        =   38
               Top             =   2910
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "시작일자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   39
               Left            =   120
               TabIndex        =   39
               Top             =   3990
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "종료일자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtInput 
               Height          =   315
               Index           =   1
               Left            =   1800
               TabIndex        =   40
               Top             =   2910
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   56885248
               CurrentDate     =   36686
            End
            Begin MSComCtl2.DTPicker dtInput 
               Height          =   315
               Index           =   2
               Left            =   1800
               TabIndex        =   41
               Top             =   3975
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   56885248
               CurrentDate     =   36686
            End
            Begin FPSpreadADO.fpSpread spdView1 
               Height          =   2535
               Left            =   120
               TabIndex        =   42
               Top             =   300
               Width           =   6465
               _Version        =   524288
               _ExtentX        =   11404
               _ExtentY        =   4471
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
               ScrollBars      =   2
               SpreadDesigner  =   "P_01001_A.frx":063C
               Appearance      =   1
               HighlightHeaders=   1
               HighlightStyle  =   1
            End
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   5040
         Left            =   15
         TabIndex        =   1
         Top             =   1335
         Width           =   15210
         _Version        =   524288
         _ExtentX        =   26829
         _ExtentY        =   8890
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
         SpreadDesigner  =   "P_01001_A.frx":0AAF
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   43
         Top             =   540
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            Style           =   2  '드롭다운 목록
            TabIndex        =   46
            Top             =   60
            Width           =   2775
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            TabIndex        =   45
            Top             =   420
            Width           =   1395
         End
         Begin XtremeSuiteControls.PushButton cmdFind 
            Height          =   660
            Left            =   5610
            TabIndex        =   44
            Top             =   60
            Width           =   1545
            _Version        =   851970
            _ExtentX        =   2725
            _ExtentY        =   1164
            _StockProps     =   79
            Caption         =   " 가맹점 찾기"
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
            Picture         =   "P_01001_A.frx":0F23
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   47
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "가맹점코드"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   11
            Left            =   60
            TabIndex        =   48
            Top             =   420
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "가맹점구분"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   49
         Top             =   15
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   900
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
         Caption         =   " #"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_01001_A.frx":1935
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Left            =   7635
         TabIndex        =   50
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
         PictureBackground=   "P_01001_A.frx":1B37
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   51
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
            Picture         =   "P_01001_A.frx":1D39
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   52
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
            Picture         =   "P_01001_A.frx":22D3
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   53
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
            Picture         =   "P_01001_A.frx":286D
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   54
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
            Picture         =   "P_01001_A.frx":2E07
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   55
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
            Picture         =   "P_01001_A.frx":33A1
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   56
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
            Picture         =   "P_01001_A.frx":393B
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   57
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
            Picture         =   "P_01001_A.frx":3ED5
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   58
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
            Picture         =   "P_01001_A.frx":446F
         End
      End
   End
End
Attribute VB_Name = "P_01001_A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Dim sPrintOption As String

Public Sub Data_Display()
    ReDim sValue(3)
    
    Dim i As Integer
    
    txtInput(1).Enabled = False
    
    sValue(0) = "0"
    sValue(1) = txtInput(0).Text & "%"
    sValue(2) = Mid(cboInput(0).Text, 2, 4)
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01001_00_ALL", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxRows = 0
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    'Call spdDisplay(RS01)
    Call fpSpread_Display(spdView, RS01)
    Call GetColWidth(REG_App, Me.Name, spdView)
    
    If RS01.RecordCount = 0 Then
        For i = 1 To txtInput.Count - 1
            txtInput(i).Text = ""
        Next i
        
        dtInput(0).Value = Date
        dtInput(1).Value = Date
        dtInput(2).Value = Date
        dtInput(3).Value = Date
        dtInput(4).Value = Date
        
        dtInput(2).Value = ""
        dtInput(4).Value = ""
        
        spdView1.MaxRows = 0
    Else
        Call Data_Display2(spdView.ActiveRow)
    End If
End Sub

'Private Sub spdDisplay(RS As ADODB.Recordset)
'
'    Call fpSpread_Display(spdView, RS)
'
'End Sub

Private Sub cmdPrint_Click()
    Call DataScreen2
    'panPrint.Visible = False
End Sub

Private Sub cboInput_Change(Index As Integer)
    If Index = 0 Then
        txtInput(0).Text = ""
    End If
End Sub

Private Sub cboInput_Click(Index As Integer)
    If Index = 0 Then
        Call Data_Display
    End If
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: Call Data_Display   ' 조회
        Case 1: Call DataAdd        ' 신규
        Case 2: Call DataSave       ' 저장
        Case 3: Call DataDelete     ' 삭제
        Case 4: Call DataCancel     ' 취소
        Case 5: 'Call DataPrint      ' 인쇄
        Case 6: 'Call DataScreen     '
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

Private Sub cmdFind_Click()
    P_01001_A1.Show vbModal
    
    Call Data_Display
End Sub

Private Sub cmdSub_Click(Index As Integer)
    Select Case Index
        Case 0: Call DataSubSave '저장
        Case 1: Call DataSubDel '삭제
    End Select
End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(1).Enabled = True
    cmdBtn(2).Enabled = True
    cmdBtn(3).Enabled = False
    cmdBtn(4).Enabled = False
    cmdBtn(5).Enabled = False
    cmdBtn(6).Enabled = False
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
    panInput.Caption = ""
    DoEvents
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
        
        
        .ColsFrozen = 2  '틀고정
        .Row = -1
    
        .Col = 1
        .ColWidth(1) = 6
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter

        .Col = 2
        .ColWidth(2) = 14
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft

        .Col = 3
        .ColWidth(3) = 5
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
    
        .Col = 4
        .ColWidth(4) = 12
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
    
        .Col = 5
        .ColWidth(5) = 5
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
    
        .Col = 6
        .ColWidth(6) = 16
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
        
        .Col = 7
        .ColWidth(7) = 5
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
         
        .Col = 8
        .ColWidth(8) = 10
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 9
        .ColWidth(9) = 10
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 10
        .ColWidth(10) = 14
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
        
        .Col = 11
        .ColWidth(11) = 8
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 12
        .ColWidth(12) = 44
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
   
        .Col = 13
        .ColWidth(13) = 5
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignRight
    End With

    With spdView1
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
    
    
        .ColsFrozen = 2  '틀고정
        .Row = -1
    
        
        .Col = 1
        .ColWidth(1) = 8
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter

        .Col = 2
        .ColWidth(2) = 5
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter

        .Col = 3
        .ColWidth(3) = 18
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
    
        .Col = 4
        .ColWidth(4) = 6
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
    
        .Col = 5
        .ColWidth(5) = 10
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
    End With

    Call Master_tblComboAdd(cboInput(0))
    
    If P_01001_A_Flag = False Then
        ' Combo BOX의 내역을 채운다.
        'Call ComboAdd
        Call Data_Display
        
'        ReDim sValue(2)
'
'        sValue(0) = "1"
'
'        Set RS01 = New ADODB.Recordset
'        Set RS01 = ExecPro("SP_01001_00_ALL", sValue(), Err_Num, Err_Dec)
'
'        spdView.MaxCols = RS01.Fields.Count
'        spdView.MaxRows = RS01.RecordCount
'
'        Call spdDisplay(RS01)
'        Call GetColWidth(REG_App, Me.Name, spdView)
        
        P_01001_A_Flag = True
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_01001_A_Flag = False
End Sub

Private Sub panCaption_Click(Index As Integer)
'    If Index = 0 Then
'        P_01001_A1.Show vbModal
'        Call Data_Display
'    End If
End Sub

Private Sub spdView_Click(ByVal Col As Long, ByVal Row As Long)
    If Row <> 0 Then
        Call Data_Display2(Row)
    End If
End Sub

Private Sub Data_Display2(iRow As Long)
    On Error GoTo ErrRtn
    
    Dim i As Integer

    With spdView
        .Row = iRow '.ActiveRow
        .Col = 1: txtInput(1).Text = .Text & ""
        .Col = 2: txtInput(2).Text = .Text & ""
        
        .Col = 3
        If .Text = "Y" Then
            optSelect(0).Value = True
        Else
            optSelect(1).Value = True
        End If
        
        .Col = 4
        If .Text = "" Then
            dtInput(0).Value = ""
        Else
            dtInput(0).Value = .Text & ""
        End If
        
        .Col = 5: txtInput(3).Text = .Text & ""
        .Col = 6: txtInput(4).Text = .Text & ""
        .Col = 7: txtInput(5).Text = .Text & ""
        
        .Col = 8
        If .Text = "" Then
            dtInput(3).Value = ""
        Else
            dtInput(3).Value = .Text & ""
        End If
        
        .Col = 9
        If Trim(.Text) = "2099-12-31" Or Trim(.Text) = "" Then
            dtInput(4).Value = ""
        Else
            dtInput(4).Value = .Text & "" 'Left(.Text, 4) & "-" & Mid(.Text, 5, 2) & "-" & Right(.Text, 2)
        End If
        
        .Col = 10: txtInput(9).Text = .Text & ""
        .Col = 11: txtInput(10).Text = .Text & ""
        .Col = 12: txtInput(11).Text = .Text & ""
        .Col = 13: txtInput(12).Text = .Text & ""
    End With
    
    Call Data_Display3
    
    'spdView.Row = spdView.ActiveRow
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Data_Display3()
    On Error GoTo ErrRtn
    
    ReDim sValue(0)
    
    sValue(0) = txtInput(1).Text
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01001_01_ALL", sValue(), Err_Num, Err_Dec)
    
    spdView1.MaxRows = 0
    spdView1.MaxCols = RS01.Fields.Count
    spdView1.MaxRows = RS01.RecordCount
    Call fpSpread_Display(spdView1, RS01)
        
    'Call spdDisplay(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView1)
    
    If spdView1.MaxRows > 0 Then
        Call spdView1_Click(1, 1)
    Else
        dtInput(1).Value = Now()
        txtInput(6).Text = ""
        txtInput(7).Text = ""
        txtInput(8).Text = ""
        dtInput(2).Value = ""
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Public Sub DataAdd()
    Dim i As Integer
    
    'spdView.MaxRows = spdView.MaxRows + 1
    
    'spdView.Row = spdView.MaxRows
    'spdView.Move
    
    txtInput(1).Enabled = True
    
    For i = 1 To txtInput.Count - 1
    
        txtInput(i).Text = ""
        
    Next i
    
    dtInput(0).Value = Date
    dtInput(1).Value = Date
    dtInput(2).Value = Date
    dtInput(3).Value = Date
    dtInput(4).Value = Date
    
    dtInput(2).Value = ""
    dtInput(4).Value = ""
    
    'dtInput(2).Value = ""
    
    spdView1.MaxRows = 0
    txtInput(1).SetFocus
End Sub

Public Sub DataCancel()
    'Call Data_Display2
End Sub

Public Sub DataDelete()
'    If MsgBox("해당되는 대리점코드를 삭제하시겠습니까?", vbYesNo + vbInformation + vbDefaultButton2, "데이터 삭제") = vbYes Then
'
'        ReDim sValue(1)
'
'        sValue(0) = txtInput(1).Text
'        sValue(1) = Mid(cboInput(3).Text, 2, 4)
'
'        Call ExecPro("SP_01001_02_MASTER", sValue(), Err_Num, Err_Dec)
'
'        If Err_Num = 0 Then
'            spdView.Row = spdView.ActiveRow
'            spdView.Action = ActionDeleteRow
'
'            MsgBox "해당되는 데이터가 정상적으로 삭제가 되었습니다.", vbInformation
'        End If
'    End If
End Sub

Private Sub ComboAdd()
    ReDim sValue(1)
    
    sValue(0) = "0"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_00001", sValue(), Err_Num, Err_Dec)

    Do While Not RS01.EOF
        cboInput(0).AddItem "[" & RS01!담당자코드 & "] " & RS01!담당자명
        
        RS01.MoveNext
    Loop

    sValue(0) = "0"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_00002", sValue(), Err_Num, Err_Dec)

    Do While Not RS01.EOF
        cboInput(1).AddItem "[" & RS01!기사코드 & "] " & RS01!기사명
        
        RS01.MoveNext
    Loop
    
    cboInput(2).AddItem "[0] 해당없음"
    cboInput(2).AddItem "[5] 목요일"
    cboInput(2).AddItem "[6] 금요일"
    cboInput(2).AddItem "[7] 토요일"
    cboInput(2).AddItem "[1] 일요일"
    cboInput(2).AddItem "[2] 월요일"
    cboInput(2).AddItem "[3] 화요일"
    cboInput(2).AddItem "[4] 수요일"
End Sub

Public Sub DataSave()
    If MsgBox("해당되는 내역을 저장하시겠습니까?", vbYesNo + vbInformation, "데이터 저장") = vbYes Then
    
        ReDim sValue(11)
        
        PanelsMsg ""
        
        sValue(0) = txtInput(1).Text                                    ' 가맹점코드
        sValue(1) = txtInput(2).Text                                    ' 가맹점명
        If optSelect(0).Value Then                                      ' 가맹점상태
            sValue(2) = "Y"
        Else
            sValue(2) = "N"
        End If
        sValue(3) = Format(dtInput(0).Value, "YYYY-MM-DD")                ' 시작일자
        sValue(4) = txtInput(3).Text                                    ' 사업장
        sValue(5) = txtInput(5).Text                                    ' 택번호
        
        sValue(6) = Format(dtInput(3).Value, "YYYY-MM-DD")                ' 오픈일자
        
        If IsNull(dtInput(4).Value) Then
            sValue(7) = "20991231"                                      ' 폐점일자
        Else
            sValue(7) = Format(dtInput(4).Value, "YYYY-MM-DD")
        End If
        'sValue(7) = Format(dtInput(4).Value, "YYYY-MM-DD")                ' 폐점일자
        
        sValue(8) = txtInput(9).Text                                    ' 전화번호
        sValue(9) = txtInput(10).Text                                   ' 우편번호
        sValue(10) = txtInput(11).Text                                  ' 주소
        sValue(11) = txtInput(12).Text                                  ' 마진
        
        Call ExecPro("SP_01001_03_ALL", sValue(), Err_Num, Err_Dec)
        
        Call SaveChangeInof ' 변경 정보를 저장한다.
        
        If Err_Num = 0 Then
            
            If txtInput(1).Enabled Then
                MsgBox "신규 데이터가 정상적으로 저장이 되었습니다." & Chr(13) & "가맹점 기준으로 다시 조회합니다.", vbInformation
                txtInput(0).Text = txtInput(1).Text
                'cboInput(0).ListIndex = "[" & txtInput(3).Text & "] " & Trim(txtInput(4).Text)
                Call Data_Display
                'txtInput(1).Enabled = False
            Else
                'spdView.ActiveRow
                With spdView
'                    .Row = .ActiveRow
'                    If txtInput(1).Enabled Then
'                    .Co1 = 1
'                    .Text = txtInput(1).Text

                    .Col = 2: .Text = txtInput(2).Text
                    
                    .Col = 3
                    If optSelect(0).Value Then
                        .Text = "Y"
                    Else
                        .Text = "N"
                    End If
                    
                    .Col = 4:  .Text = Format(dtInput(0).Value, "YYYY-MM-DD")
                    .Col = 5:  .Text = txtInput(3).Text
                    .Col = 6:  .Text = txtInput(4).Text
                    .Col = 7:  .Text = txtInput(5).Text
                    .Col = 8:  .Text = Format(dtInput(3).Value, "YYYY-MM-DD")
                    .Col = 9:  .Text = Format(dtInput(4).Value, "YYYY-MM-DD")
                    .Col = 10: .Text = txtInput(9).Text
                    .Col = 11: .Text = txtInput(10).Text
                    .Col = 12: .Text = txtInput(11).Text
                    .Col = 13: .Text = txtInput(12).Text
                End With
                
                Call Data_Display3
                
                MsgBox "해당되는 데이터가 정상적으로 저장이 되었습니다.", vbInformation
            End If
        End If

        'txtInput(1).Enabled = False
    End If
End Sub

Public Sub DataSubSave()
    If MsgBox("해당되는 내역을 저장하시겠습니까?", vbYesNo + vbInformation, "데이터 저장") = vbYes Then
    
        ReDim sValue(4)
        
        PanelsMsg ""
        
        sValue(0) = txtInput(1).Text                                    ' 가맹점코드
        sValue(1) = Format(dtInput(1).Value, "YYYY-MM-DD")                ' 시작일자
        sValue(2) = txtInput(6).Text                                    ' 사업장
        sValue(3) = txtInput(8).Text                                    ' 택번호
        
        If IsNull(dtInput(2).Value) Then
            sValue(4) = "20991231"                                      ' 종료일자
        Else
            sValue(4) = Format(dtInput(2).Value, "YYYY-MM-DD")
        End If
        
        Call ExecPro("SP_01001_04_ALL", sValue(), Err_Num, Err_Dec)
        
        
        ' 변경 정보를 저장한다.
        ReDim sValue(12)
        
        With spdView1
            .Row = .ActiveRow
            .Col = 1:   sValue(1) = Left(.Text, 4) & "-" & Mid(.Text, 5, 2) & "-" & Right(.Text, 2)
            .Col = 2:   sValue(2) = .Text
            .Col = 4:   sValue(3) = .Text
            sValue(4) = txtInput(1).Text
            sValue(5) = txtInput(2).Text
        End With
        
        sValue(0) = "1"
        sValue(6) = dtInput(1).Value
        sValue(7) = txtInput(6).Text
        sValue(8) = txtInput(8).Text
        sValue(9) = txtInput(1).Text
        sValue(10) = txtInput(2).Text
        sValue(11) = UserID & ": " & USERNAME
        sValue(12) = Now & " EDIT"
        
        Call ExecPro("SP_STORE_CHANGE", sValue(), Err_Num, Err_Dec)
        
        
        If Err_Num = 0 Then
            MsgBox "해당되는 데이터가 정상적으로 저장이 되었습니다.", vbInformation
        End If
        
        Call Data_Display3
    End If
End Sub

Public Sub DataSubDel()
    If MsgBox("해당되는 내역을 삭제하시겠습니까?", vbYesNo + vbInformation, "데이터 저장") = vbYes Then
    
        ReDim sValue(1)
        
        PanelsMsg ""
        If Trim(txtInput(1).Text) = "" Or Trim(Format(dtInput(1).Value, "YYYY-MM-DD")) = "" Then
            MsgBox "해당되는 데이터가 삭제하지 못했습니다.", vbInformation
            Call Data_Display3
            Exit Sub
        End If
        
        sValue(0) = txtInput(1).Text                                    ' 가맹점코드
        sValue(1) = Format(dtInput(1).Value, "YYYY-MM-DD")                ' 시작일자

        
        Call ExecPro("SP_01001_05_ALL", sValue(), Err_Num, Err_Dec)
        
        If Err_Num = 0 Then
            MsgBox "해당되는 데이터가 정상적으로 삭제 되었습니다.", vbInformation
        End If
        
        Call Data_Display3
    End If
End Sub

Public Sub DataPrint()

End Sub

Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If NewRow <= 0 Then
        '
    Else
        Call Data_Display2(NewRow)
    End If
    
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
'
'    If Row > 0 Then
'        Call Data_Display2(NewRow)
'    End If
End Sub


Private Sub spdView1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    With spdView1
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


Public Sub DataScreen()
    'panPrint.Visible = True
    
    sPrintOption = "2"
End Sub

Private Sub DataScreen2()
'    Dim ReportFP As String
'    Dim ReportFile As String
'
'    ReportFP = GetIniStr("REPORT", "FilePath", "", sIniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    Dim i As Integer
'    For i = 0 To 30
'        P_00000.crPrint.Formulas(i) = ""
'    Next
'
'    P_00000.crPrint.StoredProcParam(0) = "0"
'
'    If optPrint(0).Value = True Then
'        P_00000.crPrint.StoredProcParam(1) = "0"
'    ElseIf optPrint(1).Value = True Then
'        P_00000.crPrint.StoredProcParam(1) = "1"
'    ElseIf optPrint(2).Value = True Then
'        P_00000.crPrint.StoredProcParam(1) = "2"
'    End If
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'
'    If sPrintOption = "2" Then
'        Call ReportPrint(ReportFile, "2")
'    ElseIf sPrintOption = "1" Then
'        Call ReportPrint(ReportFile, "1")
'    End If
End Sub

Private Sub spdView1_Click(ByVal Col As Long, ByVal Row As Long)

    With spdView1
'        .ActiveRow = Row
'        .ActiveCol = Col

        If Row = 0 And Col = 0 Then
            If MsgBox("해당되는 내역을 저장하시겠습니까?", vbYesNo + vbInformation, "데이터 저장") = vbYes Then
                ReDim sValue(4)
                
                PanelsMsg ""
                
                sValue(0) = txtInput(1).Text                                    ' 가맹점코드
                sValue(1) = InputBox("시작일자를 YYYY-MM-DD 형식으로 입력하여 주십시요.", "시작일자 입력")
                
                If Not (Len(sValue(1)) = 8 And IsDate(Format(sValue(1), "@@@@-@@-@@")) = True) Then
                    MsgBox "일자를 확인하여 주십시요.....", vbInformation, "일자 입력 오류"
                    Exit Sub
                End If
                
                sValue(2) = txtInput(6).Text                                    ' 사업장
                sValue(3) = txtInput(8).Text                                    ' 택번호
                sValue(4) = ""                                      ' 종료일자
            
            
                Call ExecPro("SP_01001_04_ALL", sValue(), Err_Num, Err_Dec)
                
                If Err_Num = 0 Then
                    MsgBox "해당되는 데이터가 정상적으로 저장이 되었습니다.", vbInformation
                End If
                
                Call Data_Display3
            End If
        
            Exit Sub
        
        ElseIf Row = 0 Then
            Exit Sub
        End If

        .Col = Col
        .Row = Row

        .Col = 1: dtInput(1).Value = .Text & "" 'Left(.Text, 4) & "-" & Mid(.Text, 5, 2) & "-" & Right(.Text, 2)
        .Col = 2: txtInput(6).Text = .Text & ""
        .Col = 3: txtInput(7).Text = .Text & ""
        .Col = 4: txtInput(8).Text = .Text & ""
        
        .Col = 5
        
        If Trim(.Text) = "2099-12-31" Or Trim(.Text) = "" Then
            dtInput(2).Value = ""
        Else
            dtInput(2).Value = .Text & "" 'Left(.Text, 4) & "-" & Mid(.Text, 5, 2) & "-" & Right(.Text, 2)
        End If
    End With
End Sub

Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case Index
            Case 1
                If Len(Trim(txtInput(Index).Text)) <> 6 Then
                    MsgBox "가맹점코드는 6자리로 구성 하여야 합니다", vbInformation
                    txtInput(Index).SetFocus
                Else
                    ReDim sValue(0)
                    sValue(0) = txtInput(Index).Text
                    Set RS01 = New ADODB.Recordset
                    Set RS01 = ExecPro("SP_A_0003", sValue(), Err_Num, Err_Dec)
                    
                    If RS01.RecordCount = 0 Then
                        txtInput(Index + 1).Text = ""
                        SendKeys "{TAB}"
                    Else
                        'txtInput(Index + 1).Text = RS01!가맹점명
                        'txtInput(Index).Text
                        MsgBox "가맹점코드[" & txtInput(Index).Text & "]는 " & RS01!가맹점명 & "으로 등록 되어 있습니다." & Chr(13) & "확인후 등록 바랍니다.", vbInformation
                        txtInput(Index).Text = ""
                        txtInput(Index).SetFocus
                    End If
                End If
                'SendKeys "{TAB}"
                
            Case 3, 6
                ReDim sValue(0)
                sValue(0) = txtInput(Index).Text
                Set RS01 = New ADODB.Recordset
                Set RS01 = ExecPro("PRO_A_0002", sValue(), Err_Num, Err_Dec)
                
                If RS01.RecordCount = 0 Then
                    txtInput(Index + 1).Text = ""
                    txtInput(Index).SetFocus
                Else
                    txtInput(Index + 1).Text = RS01!사업장명
                    SendKeys "{TAB}"
                End If
                
            Case Else
                SendKeys "{TAB}"
        End Select
        'SendKeys "{TAB}"
    End If
End Sub

Private Sub SaveChangeInof()
        ' 정보 변경 히스토리
    ReDim sValue(12)
    
    With spdView
        .Row = .ActiveRow
        .Col = 4:   sValue(1) = Left(.Text, 4) & "-" & Mid(.Text, 5, 2) & "-" & Right(.Text, 2)
        .Col = 5:   sValue(2) = .Text
        .Col = 7:   sValue(3) = .Text
        .Col = 1:   sValue(4) = .Text
        .Col = 2:   sValue(5) = .Text
    End With
    
    sValue(0) = "1"
    
    sValue(6) = dtInput(0).Value
    sValue(7) = txtInput(3).Text
    sValue(8) = txtInput(5).Text
    sValue(9) = txtInput(1).Text
    sValue(10) = txtInput(2).Text
    sValue(11) = UserID & ": " & USERNAME
    sValue(12) = Now
        
    Call ExecPro("SP_STORE_CHANGE", sValue(), Err_Num, Err_Dec)
End Sub
