VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_04014 
   Caption         =   "점별 실적"
   ClientHeight    =   11250
   ClientLeft      =   960
   ClientTop       =   2610
   ClientWidth     =   16320
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_04014.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11250
   ScaleWidth      =   16320
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   11250
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16320
      _ExtentX        =   28787
      _ExtentY        =   19844
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_04014.frx":058A
      Begin Threed.SSPanel SSPanel 
         Height          =   1905
         Index           =   0
         Left            =   15
         TabIndex        =   2
         Top             =   9330
         Width           =   16290
         _ExtentX        =   28734
         _ExtentY        =   3360
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   0
            Left            =   1665
            Locked          =   -1  'True
            TabIndex        =   44
            Top             =   645
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   1
            Left            =   2985
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   645
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   2
            Left            =   4305
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   645
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   3
            Left            =   5625
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   645
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   4
            Left            =   6945
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   645
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   5
            Left            =   8265
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   645
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   6
            Left            =   9585
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   645
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   7
            Left            =   10905
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   645
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   8
            Left            =   12225
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   645
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   9
            Left            =   13545
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   645
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   10
            Left            =   1665
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   945
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   11
            Left            =   2985
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   945
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   12
            Left            =   4305
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   945
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   13
            Left            =   5625
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   945
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   14
            Left            =   6945
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   945
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   15
            Left            =   8265
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   945
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   16
            Left            =   9585
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   945
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   17
            Left            =   10905
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   945
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   18
            Left            =   12225
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   945
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   19
            Left            =   13545
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   945
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   20
            Left            =   1665
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   1245
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   21
            Left            =   2985
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   1245
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   22
            Left            =   4305
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   1245
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   23
            Left            =   5625
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   1245
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   24
            Left            =   6945
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   1245
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   25
            Left            =   8265
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   1245
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   26
            Left            =   9585
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   1245
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   27
            Left            =   10905
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   1245
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   28
            Left            =   12225
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   1245
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   29
            Left            =   13545
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   1245
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   30
            Left            =   1665
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   1545
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   31
            Left            =   2985
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   1545
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   32
            Left            =   4305
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   1545
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   33
            Left            =   5625
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   1545
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   34
            Left            =   6945
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   1545
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   35
            Left            =   8265
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   1545
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   36
            Left            =   9585
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   1545
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   37
            Left            =   10905
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   1545
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   38
            Left            =   12225
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   1545
            Width           =   1340
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   39
            Left            =   13545
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   1545
            Width           =   1340
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   8
            Left            =   45
            TabIndex        =   13
            Top             =   1545
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "총  합  계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   21
            Left            =   45
            TabIndex        =   24
            Top             =   1245
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "할인매장 계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   15
            Left            =   13545
            TabIndex        =   45
            Top             =   345
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "누  계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   14
            Left            =   12225
            TabIndex        =   46
            Top             =   345
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "월  계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   13
            Left            =   10905
            TabIndex        =   47
            Top             =   345
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "누  계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   12
            Left            =   9585
            TabIndex        =   48
            Top             =   345
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "월  계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   11
            Left            =   8265
            TabIndex        =   49
            Top             =   345
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "누  계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   10
            Left            =   6945
            TabIndex        =   50
            Top             =   345
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "월  계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   9
            Left            =   5625
            TabIndex        =   51
            Top             =   345
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "누  계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   7
            Left            =   45
            TabIndex        =   52
            Top             =   945
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "백 화 점 계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   6
            Left            =   45
            TabIndex        =   53
            Top             =   645
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "대 리 점 계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   5
            Left            =   4305
            TabIndex        =   54
            Top             =   345
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "월  계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   4
            Left            =   2985
            TabIndex        =   55
            Top             =   345
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "누  계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   3
            Left            =   1665
            TabIndex        =   56
            Top             =   345
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "월  계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   16
            Left            =   12225
            TabIndex        =   57
            Top             =   45
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "전년대비실적율"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   17
            Left            =   9585
            TabIndex        =   58
            Top             =   45
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "목표대비달성율"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   18
            Left            =   6945
            TabIndex        =   59
            Top             =   45
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "금  년  실  적"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   19
            Left            =   4305
            TabIndex        =   60
            Top             =   45
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "금  년  목  표"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   20
            Left            =   1665
            TabIndex        =   61
            Top             =   45
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "전  년  실  적"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   7980
         Left            =   15
         TabIndex        =   1
         Top             =   1335
         Width           =   16290
         _Version        =   524288
         _ExtentX        =   28734
         _ExtentY        =   14076
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
         SpreadDesigner  =   "P_04014.frx":063C
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   62
         Top             =   540
         Width           =   16290
         _ExtentX        =   28734
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   6480
            TabIndex        =   63
            Top             =   60
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   556
            _Version        =   262144
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   64
               Top             =   30
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "전  체"
               Value           =   -1
            End
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   1
               Left            =   1440
               TabIndex        =   65
               Top             =   30
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "대리점"
            End
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   2
               Left            =   2700
               TabIndex        =   66
               Top             =   30
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "백화점"
            End
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   3
               Left            =   3960
               TabIndex        =   67
               Top             =   30
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "할인매장"
            End
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Left            =   1680
            TabIndex        =   68
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   63307776
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   69
            Top             =   60
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "기 준 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   4860
            TabIndex        =   70
            Top             =   60
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "구    분"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   71
         Top             =   15
         Width           =   8685
         _ExtentX        =   15319
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
         PictureBackground=   "P_04014.frx":0AB5
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   1
         Left            =   8715
         TabIndex        =   72
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
         PictureBackground=   "P_04014.frx":0CB7
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   73
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
            Picture         =   "P_04014.frx":0EB9
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   74
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
            Picture         =   "P_04014.frx":1453
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   75
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
            Picture         =   "P_04014.frx":19ED
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   76
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
            Picture         =   "P_04014.frx":1F87
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   77
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
            Picture         =   "P_04014.frx":2521
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   78
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
            Picture         =   "P_04014.frx":2ABB
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   79
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
            Picture         =   "P_04014.frx":3055
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   80
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
            Picture         =   "P_04014.frx":35EF
         End
      End
   End
End
Attribute VB_Name = "P_04014"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

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
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    If P_04014_Flag = False Then
        dtInput.Value = Date
        
        ReDim sValue(2)
        
        sValue(0) = "1"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_04014_00", sValue(), Err_Num, Err_Dec)
        
        spdView.MaxCols = RS01.Fields.Count
        spdView.MaxRows = RS01.RecordCount
        
        Call spdDisplay(RS01)
        Call GetColWidth(REG_App, Me.Name, spdView)
        
        P_04014_Flag = True
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub spdDisplay(Rs As ADODB.Recordset)
    
    Call fpSpread_Display(spdView, Rs)
    
    spdView.ColsFrozen = 1 '틀고정
    
    spdView.Row = -1
    
    spdView.Col = 1
    spdView.ColWidth(1) = 20
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft

    spdView.Col = 2
    spdView.ColWidth(2) = 8
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 1
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 3
    spdView.ColWidth(3) = 8
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 4
    spdView.ColWidth(4) = 8
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 5
    spdView.ColWidth(5) = 8
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 6
    spdView.ColWidth(6) = 8
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 7
    spdView.ColWidth(7) = 8
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 8
    spdView.ColWidth(8) = 8
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 9
    spdView.ColWidth(9) = 8
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 2
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 10
    spdView.ColWidth(10) = 8
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 2
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 11
    spdView.ColWidth(11) = 8
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 2
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 12
    spdView.ColWidth(12) = 8
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 2
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 13
    spdView.ColHidden = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_04014_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i As Integer
    Dim nDTotal(5) As Long
    Dim nBTotal(5) As Long
    Dim nCTotal(5) As Long
    Dim nTTotal(5) As Long
    
    ReDim sValue(2)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput.Value, "YYYY-MM-DD")
    
    If optSelect(0).Value = True Then
        sValue(2) = "1"
    ElseIf optSelect(1).Value = True Then
        sValue(2) = "2"
    ElseIf optSelect(2).Value = True Then
        sValue(2) = "3"
    ElseIf optSelect(3).Value = True Then
        sValue(2) = "4"
    End If
        
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_04014_00", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    Call spdDisplay(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView)
    
    For i = 1 To spdView.MaxRows
        spdView.AutoCalc = True
        
        spdView.Row = i
        spdView.Col = 9:  spdView.Formula = "G" & i & "/ E" & i & " * 100"
        spdView.Col = 10: spdView.Formula = "H" & i & "/ F" & i & " * 100"
        spdView.Col = 11: spdView.Formula = "(G" & i & " - C" & i & ") / C" & i & " * 100"
        spdView.Col = 12: spdView.Formula = "(H" & i & " - D" & i & ") / D" & i & " * 100"
        
        ReDim sValue(2)
        
        sValue(0) = "0"
        sValue(1) = Format(dtInput.Value, "yyyy")
        spdView.Col = 1
        sValue(2) = Mid(spdView.Text, 2, 3)
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_04014_01", sValue(), Err_Num, Err_Dec)
        
        If RS01.RecordCount <> 0 Then
            Select Case Format(dtInput.Value, "mm")
                Case "01"
                    spdView.Col = 5
                    spdView.Value = RS01!수량01
                    spdView.Col = 6
                    spdView.Value = RS01!금액01
                Case "02"
                    spdView.Col = 5
                    spdView.Value = RS01!수량02
                    spdView.Col = 6
                    spdView.Value = RS01!금액02
                Case "03"
                    spdView.Col = 5
                    spdView.Value = RS01!수량03
                    spdView.Col = 6
                    spdView.Value = RS01!금액03
                Case "04"
                    spdView.Col = 5
                    spdView.Value = RS01!수량04
                    spdView.Col = 6
                    spdView.Value = RS01!금액04
                Case "05"
                    spdView.Col = 5
                    spdView.Value = RS01!수량05
                    spdView.Col = 6
                    spdView.Value = RS01!금액05
                Case "06"
                    spdView.Col = 5
                    spdView.Value = RS01!수량06
                    spdView.Col = 6
                    spdView.Value = RS01!금액06
                Case "07"
                    spdView.Col = 5
                    spdView.Value = RS01!수량07
                    spdView.Col = 6
                    spdView.Value = RS01!금액07
                Case "08"
                    spdView.Col = 5
                    spdView.Value = RS01!수량08
                    spdView.Col = 6
                    spdView.Value = RS01!금액08
                Case "09"
                    spdView.Col = 5
                    spdView.Value = RS01!수량09
                    spdView.Col = 6
                    spdView.Value = RS01!금액09
                Case "10"
                    spdView.Col = 5: spdView.Value = RS01!수량10
                    spdView.Col = 6: spdView.Value = RS01!금액10
                Case "11"
                    spdView.Col = 5: spdView.Value = RS01!수량11
                    spdView.Col = 6: spdView.Value = RS01!금액11
                Case "12"
                    spdView.Col = 5: spdView.Value = RS01!수량12
                    spdView.Col = 6: spdView.Value = RS01!금액12
            End Select
        End If
        
        spdView.Col = 13
        If spdView.Text = "1" Then
            spdView.Col = 3: nDTotal(0) = nDTotal(0) + spdView.Value
            spdView.Col = 4: nDTotal(1) = nDTotal(1) + spdView.Value
            spdView.Col = 5: nDTotal(2) = nDTotal(2) + spdView.Value
            spdView.Col = 6: nDTotal(3) = nDTotal(3) + spdView.Value
            spdView.Col = 7: nDTotal(4) = nDTotal(4) + spdView.Value
            spdView.Col = 8: nDTotal(5) = nDTotal(5) + spdView.Value
            
        ElseIf spdView.Text = "2" Then
            spdView.Col = 3: nBTotal(0) = nBTotal(0) + spdView.Value
            spdView.Col = 4: nBTotal(1) = nBTotal(1) + spdView.Value
            spdView.Col = 5: nBTotal(2) = nBTotal(2) + spdView.Value
            spdView.Col = 6: nBTotal(3) = nBTotal(3) + spdView.Value
            spdView.Col = 7: nBTotal(4) = nBTotal(4) + spdView.Value
            spdView.Col = 8: nBTotal(5) = nBTotal(5) + spdView.Value
            
        ElseIf spdView.Text = "3" Then
            spdView.Col = 3: nCTotal(0) = nCTotal(0) + spdView.Value
            spdView.Col = 4: nCTotal(1) = nCTotal(1) + spdView.Value
            spdView.Col = 5: nCTotal(2) = nCTotal(2) + spdView.Value
            spdView.Col = 6: nCTotal(3) = nCTotal(3) + spdView.Value
            spdView.Col = 7: nCTotal(4) = nBTotal(4) + spdView.Value
            spdView.Col = 8: nCTotal(5) = nCTotal(5) + spdView.Value
        End If
        
        spdView.Col = 3: nTTotal(0) = nTTotal(0) + spdView.Value
        spdView.Col = 4: nTTotal(1) = nTTotal(1) + spdView.Value
        spdView.Col = 5: nTTotal(2) = nTTotal(2) + spdView.Value
        spdView.Col = 6: nTTotal(3) = nTTotal(3) + spdView.Value
        spdView.Col = 7: nTTotal(4) = nTTotal(4) + spdView.Value
        spdView.Col = 8: nTTotal(5) = nTTotal(5) + spdView.Value
    Next i
    
    ' 대리점계
    txtInput(0).Text = Format(nDTotal(0), "#,##0")
    txtInput(1).Text = Format(nDTotal(1), "#,##0")
    txtInput(2).Text = Format(nDTotal(2), "#,##0")
    txtInput(3).Text = Format(nDTotal(3), "#,##0")
    txtInput(4).Text = Format(nDTotal(4), "#,##0")
    txtInput(5).Text = Format(nDTotal(5), "#,##0")
    If nDTotal(4) = "0" Or nDTotal(2) = "0" Then txtInput(6).Text = 0 Else txtInput(6).Text = Format(nDTotal(4) / nDTotal(2) * 100, "#,##0.00")
    If nDTotal(5) = "0" Or nDTotal(3) = "0" Then txtInput(7).Text = 0 Else txtInput(7).Text = Format(nDTotal(5) / nDTotal(3) * 100, "#,##0.00")
    If nDTotal(4) = "0" Or nDTotal(0) = "0" Then txtInput(8).Text = 0 Else txtInput(8).Text = Format((nDTotal(4) - nDTotal(0)) / nDTotal(0) * 100, "#,##0.00")
    If nDTotal(5) = "0" Or nDTotal(1) = "0" Then txtInput(9).Text = 0 Else txtInput(9).Text = Format((nDTotal(5) - nDTotal(1)) / nDTotal(1) * 100, "#,##0.00")
    ' 백화점계
    txtInput(10).Text = Format(nBTotal(0), "#,##0")
    txtInput(11).Text = Format(nBTotal(1), "#,##0")
    txtInput(12).Text = Format(nBTotal(2), "#,##0")
    txtInput(13).Text = Format(nBTotal(3), "#,##0")
    txtInput(14).Text = Format(nBTotal(4), "#,##0")
    txtInput(15).Text = Format(nBTotal(5), "#,##0")
    If nBTotal(4) = "0" Or nBTotal(2) = "0" Then txtInput(16).Text = 0 Else txtInput(16).Text = Format(nBTotal(4) / nBTotal(2) * 100, "#,##0.00")
    If nBTotal(5) = "0" Or nBTotal(3) = "0" Then txtInput(17).Text = 0 Else txtInput(17).Text = Format(nBTotal(5) / nBTotal(3) * 100, "#,##0.00")
    If nBTotal(4) = "0" Or nBTotal(0) = "0" Then txtInput(18).Text = 0 Else txtInput(18).Text = Format((nBTotal(4) - nBTotal(0)) / nBTotal(0) * 100, "#,##0.00")
    If nBTotal(5) = "0" Or nBTotal(1) = "0" Then txtInput(19).Text = 0 Else txtInput(19).Text = Format((nBTotal(5) - nBTotal(1)) / nBTotal(1) * 100, "#,##0.00")
    ' 할인매장계
    txtInput(20).Text = Format(nCTotal(0), "#,##0")
    txtInput(21).Text = Format(nCTotal(1), "#,##0")
    txtInput(22).Text = Format(nCTotal(2), "#,##0")
    txtInput(23).Text = Format(nCTotal(3), "#,##0")
    txtInput(24).Text = Format(nCTotal(4), "#,##0")
    txtInput(25).Text = Format(nCTotal(5), "#,##0")
    If nCTotal(4) = "0" Or nCTotal(2) = "0" Then txtInput(26).Text = 0 Else txtInput(26).Text = Format(nCTotal(4) / nCTotal(2) * 100, "#,##0.00")
    If nCTotal(5) = "0" Or nCTotal(3) = "0" Then txtInput(27).Text = 0 Else txtInput(27).Text = Format(nCTotal(5) / nCTotal(3) * 100, "#,##0.00")
    If nCTotal(4) = "0" Or nCTotal(0) = "0" Then txtInput(28).Text = 0 Else txtInput(28).Text = Format((nCTotal(4) - nCTotal(0)) / nCTotal(0) * 100, "#,##0.00")
    If nCTotal(5) = "0" Or nCTotal(1) = "0" Then txtInput(29).Text = 0 Else txtInput(29).Text = Format((nCTotal(5) - nCTotal(1)) / nCTotal(1) * 100, "#,##0.00")
    ' 총계
    txtInput(30).Text = Format(nTTotal(0), "#,##0")
    txtInput(31).Text = Format(nTTotal(1), "#,##0")
    txtInput(32).Text = Format(nTTotal(2), "#,##0")
    txtInput(33).Text = Format(nTTotal(3), "#,##0")
    txtInput(34).Text = Format(nTTotal(4), "#,##0")
    txtInput(35).Text = Format(nTTotal(5), "#,##0")
    If nTTotal(4) = "0" Or nTTotal(2) = "0" Then txtInput(36).Text = 0 Else txtInput(36).Text = Format(nTTotal(4) / nTTotal(2) * 100, "#,##0.00")
    If nTTotal(5) = "0" Or nTTotal(3) = "0" Then txtInput(37).Text = 0 Else txtInput(37).Text = Format(nTTotal(5) / nTTotal(3) * 100, "#,##0.00")
    If nTTotal(4) = "0" Or nTTotal(0) = "0" Then txtInput(38).Text = 0 Else txtInput(38).Text = Format((nTTotal(4) - nTTotal(0)) / nTTotal(0) * 100, "#,##0.00")
    If nTTotal(5) = "0" Or nTTotal(1) = "0" Then txtInput(39).Text = 0 Else txtInput(39).Text = Format((nTTotal(5) - nTTotal(1)) / nTTotal(1) * 100, "#,##0.00")
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub spdView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        'PopupMenu P_00000.PopUp
    End If
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
'    If optSelect(0).Value = True Then
'        P_00000.crPrint.Formulas(0) = "매장구분 = '전  체'"
'    ElseIf optSelect(1).Value = True Then
'        P_00000.crPrint.Formulas(0) = "매장구분 = '대리점'"
'    ElseIf optSelect(2).Value = True Then
'        P_00000.crPrint.Formulas(0) = "매장구분 = '백화점'"
'    ElseIf optSelect(3).Value = True Then
'        P_00000.crPrint.Formulas(0) = "매장구분 = '할인매장'"
'    End If
'
'    P_00000.crPrint.Formulas(1) = "기준일자 = '" & Format(dtInput.Value, "YYYY-MM-DD") & "'"
'
'    sData = Right(Space(10) & txtInput(0).Text, 10)
'    sData = sData & Right(Space(11) & txtInput(1).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(2).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(3).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(4).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(5).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(6).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(7).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(8).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(9).Text, 11)
'
'    P_00000.crPrint.Formulas(2) = "합계01 = '" & sData & "'"
'
'    sData = Right(Space(10) & txtInput(10).Text, 10)
'    sData = sData & Right(Space(11) & txtInput(11).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(12).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(13).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(14).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(15).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(16).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(17).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(18).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(19).Text, 11)
'
'    P_00000.crPrint.Formulas(3) = "합계02 = '" & sData & "'"
'
'    sData = Right(Space(10) & txtInput(20).Text, 10)
'    sData = sData & Right(Space(11) & txtInput(21).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(22).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(23).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(24).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(25).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(26).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(27).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(28).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(29).Text, 11)
'
'    P_00000.crPrint.Formulas(4) = "합계03 = '" & sData & "'"
'
'    sData = Right(Space(10) & txtInput(30).Text, 10)
'    sData = sData & Right(Space(11) & txtInput(31).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(32).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(33).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(34).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(35).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(36).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(37).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(38).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(39).Text, 11)
'
'    P_00000.crPrint.Formulas(5) = "합계04 = '" & sData & "'"
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
'    If optSelect(0).Value = True Then
'        P_00000.crPrint.Formulas(0) = "매장구분 = '전  체'"
'    ElseIf optSelect(1).Value = True Then
'        P_00000.crPrint.Formulas(0) = "매장구분 = '대리점'"
'    ElseIf optSelect(2).Value = True Then
'        P_00000.crPrint.Formulas(0) = "매장구분 = '백화점'"
'    ElseIf optSelect(3).Value = True Then
'        P_00000.crPrint.Formulas(0) = "매장구분 = '할인매장'"
'    End If
'
'    P_00000.crPrint.Formulas(1) = "기준일자 = '" & Format(dtInput.Value, "YYYY-MM-DD") & "'"
'
'    sData = Right(Space(10) & txtInput(0).Text, 10)
'    sData = sData & Right(Space(11) & txtInput(1).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(2).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(3).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(4).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(5).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(6).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(7).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(8).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(9).Text, 11)
'
'    P_00000.crPrint.Formulas(2) = "합계01 = '" & sData & "'"
'
'    sData = Right(Space(10) & txtInput(10).Text, 10)
'    sData = sData & Right(Space(11) & txtInput(11).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(12).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(13).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(14).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(15).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(16).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(17).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(18).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(19).Text, 11)
'
'    P_00000.crPrint.Formulas(3) = "합계02 = '" & sData & "'"
'
'    sData = Right(Space(10) & txtInput(20).Text, 10)
'    sData = sData & Right(Space(11) & txtInput(21).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(22).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(23).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(24).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(25).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(26).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(27).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(28).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(29).Text, 11)
'
'    P_00000.crPrint.Formulas(4) = "합계03 = '" & sData & "'"
'
'    sData = Right(Space(10) & txtInput(30).Text, 10)
'    sData = sData & Right(Space(11) & txtInput(31).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(32).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(33).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(34).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(35).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(36).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(37).Text, 11)
'    sData = sData & Right(Space(11) & txtInput(38).Text, 12)
'    sData = sData & Right(Space(11) & txtInput(39).Text, 11)
'
'    P_00000.crPrint.Formulas(5) = "합계04 = '" & sData & "'"
'
'    Call ReportPrint(ReportFile, "2")
End Sub

Private Sub PrintDesc()
    Dim i As Integer
    Dim TempText As String
    Dim TempFP As String
    Dim TempFile As String
    
    TempFP = GetIniStr("REPORT", "TempPath", "", m_iniFile)
    TempFile = TempFP & "\Temp.txt"
    
    Open TempFile For Output As #1
    
    TempText = ""
    
    For i = 1 To spdView.MaxRows - 1
        spdView.Row = i
        
        spdView.Col = 1
        TempText = LeftH(Mid(spdView.Text, 7) & Space(16), 16)
        spdView.Col = 2
        TempText = TempText & RightH(Space(6) & spdView.Text, 6) & Space(2)
        spdView.Col = 3
        TempText = TempText & RightH(Space(9) & spdView.Text, 9) & Space(2)
        spdView.Col = 4
        TempText = TempText & RightH(Space(9) & spdView.Text, 9) & Space(2)
        spdView.Col = 5
        TempText = TempText & RightH(Space(9) & spdView.Text, 9) & Space(2)
        spdView.Col = 6
        TempText = TempText & RightH(Space(9) & spdView.Text, 9) & Space(2)
        spdView.Col = 7
        TempText = TempText & RightH(Space(9) & spdView.Text, 9) & Space(2)
        spdView.Col = 8
        TempText = TempText & RightH(Space(9) & spdView.Text, 9) & Space(2)
        spdView.Col = 9
        TempText = TempText & RightH(Space(9) & spdView.Text, 9) & Space(2)
        spdView.Col = 10
        TempText = TempText & RightH(Space(9) & spdView.Text, 9) & Space(2)
        spdView.Col = 11
        TempText = TempText & RightH(Space(9) & spdView.Text, 9) & Space(2)
        spdView.Col = 12
        TempText = TempText & RightH(Space(9) & spdView.Text, 9)
        
        Print #1, TempText
    Next i
    
    Close #1
End Sub

