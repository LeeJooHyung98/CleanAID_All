VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form frm사고품1 
   Caption         =   "사고품 관리"
   ClientHeight    =   12210
   ClientLeft      =   750
   ClientTop       =   3210
   ClientWidth     =   17730
   ControlBox      =   0   'False
   LinkTopic       =   "Form31"
   MDIChild        =   -1  'True
   ScaleHeight     =   12210
   ScaleWidth      =   17730
   WindowState     =   2  '최대화
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   315
      TabIndex        =   64
      Top             =   9615
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   556
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog cdPrt 
      Left            =   10815
      Top             =   75
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   795
      Left            =   45
      TabIndex        =   48
      Top             =   120
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   1402
      _Version        =   262144
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComCtl2.DTPicker dtInput 
         Height          =   345
         Index           =   0
         Left            =   8040
         TabIndex        =   27
         Top             =   315
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   109117440
         CurrentDate     =   36686
      End
      Begin VB.CommandButton Command1 
         Caption         =   "인쇄"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   4
         Left            =   5400
         MaskColor       =   &H00FFFFFF&
         Style           =   1  '그래픽
         TabIndex        =   59
         Top             =   135
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "삭제"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   3
         Left            =   4065
         MaskColor       =   &H00FFFFFF&
         Style           =   1  '그래픽
         TabIndex        =   25
         Top             =   135
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "저장"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   2
         Left            =   2730
         MaskColor       =   &H00FFFFFF&
         Style           =   1  '그래픽
         TabIndex        =   24
         Top             =   135
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "입력"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   0
         Left            =   60
         MaskColor       =   &H00FFFFFF&
         Style           =   1  '그래픽
         TabIndex        =   22
         Top             =   135
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "조회"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   1
         Left            =   1395
         MaskColor       =   &H00FFFFFF&
         Style           =   1  '그래픽
         TabIndex        =   23
         Top             =   135
         Width           =   1335
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   345
         Index           =   14
         Left            =   6795
         TabIndex        =   49
         Top             =   315
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   609
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "접수일자"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1275
      Left            =   45
      TabIndex        =   29
      Top             =   2610
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   2249
      _Version        =   262144
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "[ 고객정보 ]"
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   345
         Index           =   1
         Left            =   1980
         TabIndex        =   0
         Top             =   585
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtInput 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   8985
         MaxLength       =   14
         TabIndex        =   3
         Top             =   750
         Width           =   2640
      End
      Begin VB.TextBox txtInput 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   5160
         MaxLength       =   14
         TabIndex        =   2
         Top             =   750
         Width           =   2640
      End
      Begin VB.TextBox txtInput 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   5160
         MaxLength       =   50
         TabIndex        =   1
         Top             =   330
         Width           =   6465
      End
      Begin VB.TextBox txtInput 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   100
         Left            =   1950
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   330
         Width           =   1875
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   360
         Index           =   1
         Left            =   3855
         TabIndex        =   44
         Top             =   330
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "주   소"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   360
         Index           =   2
         Left            =   3855
         TabIndex        =   45
         Top             =   750
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   635
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "전   화"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   360
         Index           =   3
         Left            =   7830
         TabIndex        =   46
         Top             =   750
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   635
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "휴대폰"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   780
         Index           =   181
         Left            =   465
         TabIndex        =   47
         Top             =   330
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   1376
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "소비자명"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   2415
      Left            =   45
      TabIndex        =   28
      Top             =   4710
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   4260
      _Version        =   262144
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "[ 피해관련사항 ]"
      Begin VB.TextBox txtInput 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   12
         Left            =   10050
         MaxLength       =   10
         TabIndex        =   13
         Top             =   1470
         Width           =   1500
      End
      Begin VB.TextBox txtInput 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   7815
         MaxLength       =   10
         TabIndex        =   12
         Top             =   1470
         Width           =   1530
      End
      Begin VB.TextBox txtInput 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   10335
         MaxLength       =   5
         TabIndex        =   10
         Top             =   1110
         Width           =   1215
      End
      Begin VB.TextBox txtInput 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   13
         Left            =   2235
         MaxLength       =   7
         TabIndex        =   14
         Top             =   1830
         Width           =   3735
      End
      Begin VB.TextBox txtInput 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   2235
         MaxLength       =   15
         TabIndex        =   11
         Top             =   1470
         Width           =   3735
      End
      Begin VB.TextBox txtInput 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   7815
         MaxLength       =   5
         TabIndex        =   9
         Top             =   1110
         Width           =   1140
      End
      Begin VB.TextBox txtInput 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   2235
         MaxLength       =   15
         TabIndex        =   8
         Top             =   1110
         Width           =   3735
      End
      Begin VB.TextBox txtInput 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   7815
         MaxLength       =   4
         TabIndex        =   7
         Top             =   750
         Width           =   3735
      End
      Begin VB.TextBox txtInput 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   7815
         MaxLength       =   15
         TabIndex        =   5
         Top             =   390
         Width           =   3735
      End
      Begin VB.TextBox txtInput 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   2235
         MaxLength       =   15
         TabIndex        =   4
         Top             =   390
         Width           =   3735
      End
      Begin MSComCtl2.DTPicker dtInput 
         Height          =   360
         Index           =   1
         Left            =   2235
         TabIndex        =   6
         Top             =   750
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   109117440
         CurrentDate     =   36686
      End
      Begin MSComCtl2.DTPicker dtInput 
         Height          =   345
         Index           =   3
         Left            =   7815
         TabIndex        =   15
         Top             =   1830
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   109117440
         CurrentDate     =   36686
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   345
         Index           =   4
         Left            =   420
         TabIndex        =   34
         Top             =   390
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   609
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "품       목"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   345
         Index           =   5
         Left            =   6030
         TabIndex        =   35
         Top             =   390
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   609
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "상        표"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   330
         Index           =   151
         Left            =   420
         TabIndex        =   36
         Top             =   765
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   582
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "구 입 일 자"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   345
         Index           =   6
         Left            =   6030
         TabIndex        =   37
         Top             =   765
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   609
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "색        상"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   330
         Index           =   7
         Left            =   420
         TabIndex        =   38
         Top             =   1125
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   582
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "구  입  처"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   315
         Index           =   8
         Left            =   6030
         TabIndex        =   39
         Top             =   1140
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   556
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "최초TAG"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   330
         Index           =   10
         Left            =   420
         TabIndex        =   40
         Top             =   1485
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   582
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "구 입 형 태"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   315
         Index           =   11
         Left            =   6030
         TabIndex        =   41
         Top             =   1485
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   556
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "최초입고일"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   345
         Index           =   13
         Left            =   420
         TabIndex        =   42
         Top             =   1845
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   609
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "구 입 가 격"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   345
         Index           =   171
         Left            =   6030
         TabIndex        =   43
         Top             =   1830
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   609
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "사고 접수일"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   330
         Index           =   9
         Left            =   8985
         TabIndex        =   58
         Top             =   1125
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "최종TAG"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   360
         Index           =   12
         Left            =   9360
         TabIndex        =   63
         Top             =   1470
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   635
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "최종"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   2280
      Left            =   45
      TabIndex        =   26
      Top             =   7215
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   4022
      _Version        =   262144
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "[ 대리점 기재 ]"
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9810
         TabIndex        =   21
         Top             =   1560
         Width           =   1740
      End
      Begin VB.TextBox txtInput 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   18
         Left            =   5985
         MaxLength       =   10
         TabIndex        =   20
         Top             =   1545
         Width           =   1785
      End
      Begin VB.TextBox txtInput 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   17
         Left            =   2340
         MaxLength       =   10
         TabIndex        =   19
         Top             =   1545
         Width           =   1935
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   360
         Index           =   14
         Left            =   495
         TabIndex        =   31
         Top             =   360
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   635
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "사고의 종류"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin VB.TextBox txtInput 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   16
         Left            =   2325
         MaxLength       =   50
         TabIndex        =   18
         Top             =   1110
         Width           =   9225
      End
      Begin VB.TextBox txtInput 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   15
         Left            =   2325
         MaxLength       =   50
         TabIndex        =   17
         Top             =   735
         Width           =   9225
      End
      Begin VB.TextBox txtInput 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   14
         Left            =   2310
         MaxLength       =   50
         TabIndex        =   16
         Top             =   360
         Width           =   9225
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   360
         Index           =   15
         Left            =   495
         TabIndex        =   32
         Top             =   735
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   635
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "사고의 내용"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   360
         Index           =   16
         Left            =   495
         TabIndex        =   33
         Top             =   1110
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   635
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "소비자 의견"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   360
         Index           =   17
         Left            =   495
         TabIndex        =   60
         Top             =   1530
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   635
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "보상산정금액"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   360
         Index           =   18
         Left            =   4290
         TabIndex        =   61
         Top             =   1545
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   635
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "합의금액"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   360
         Index           =   221
         Left            =   7815
         TabIndex        =   62
         Top             =   1560
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   635
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "처리유뮤"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   6075
      Left            =   3615
      TabIndex        =   50
      Top             =   945
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   10716
      _Version        =   262144
      BackColor       =   13160660
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   5535
         Left            =   90
         TabIndex        =   57
         Top             =   465
         Width           =   11550
         _Version        =   524288
         _ExtentX        =   20373
         _ExtentY        =   9763
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "frm사고품1.frx":0000
      End
      Begin VB.TextBox txtInput 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   30
         Left            =   1365
         TabIndex        =   54
         Top             =   60
         Width           =   1770
      End
      Begin VB.TextBox txtInput 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   31
         Left            =   4500
         TabIndex        =   52
         Top             =   60
         Width           =   2160
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   360
         Index           =   0
         Left            =   75
         TabIndex        =   51
         Top             =   60
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "소비자명"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   360
         Index           =   141
         Left            =   3180
         TabIndex        =   53
         Top             =   60
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   635
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "전   화"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   345
         Index           =   0
         Left            =   6705
         TabIndex        =   55
         Top             =   90
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   609
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "접수일자"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin MSComCtl2.DTPicker dtInput 
         Height          =   345
         Index           =   4
         Left            =   8040
         TabIndex        =   56
         Top             =   90
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   109117440
         CurrentDate     =   37742
      End
   End
End
Attribute VB_Name = "frm사고품1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iEditMode As Integer    ' 0 = 정상, 1 = 입력, 2 = 수정
Dim iReadRecord As Integer  ' 0 = 읽은 레코드 없음, 숫자 = 읽은 레코드 번호
Dim iErrorCheck As Integer  ' 1 = 정상
Dim strErrMsg As String     ' 각종 메시지 출력

'+------------------------------------------------------
'+
'+ 2003/03/22
'+
'+루틴설명
'+
'+  1. 사고 폼의 내용을 지운다.
'+
'+------------------------------------------------------
Private Sub TextBoxClear()
    For i = 1 To 1
        MaskEdBox1(i).Text = ""
    Next i
    
    For i = 1 To 18
        txtInput(i).Text = ""
    Next i
    
    For i = 30 To 31
        txtInput(i).Text = ""
    Next i
End Sub

'+------------------------------------------------------
'+
'+ 2003/03/22
'+
'+  - 전달값
'+      iReadRecord : 현재 읽은 레코드의 값
'+  - 리턴값
'+      0           : 저장 오류
'+      1           : 저장 성공
'+      3           : 필수 입력 내용 미입력시
'+      4           : 기타 오류
'+  - 루틴설명
'+      1. 모드및 필수 항목 확인
'+      2. 신규 저장과 수정을 구분하여 저장
'+
'+------------------------------------------------------
Private Function SaveData(iReadRecord) As Integer
    Dim strMsg As String
    Dim strDate(4) As String
    Dim lDBCount As Long

    On Error GoTo ErrRtn
    
    ' 모드를 확인한다.
    If iEditMode <> 1 And iEditMode <> 2 Then
        SaveData = 0
        Exit Function
    End If

    '필수 입력항목 검사
    If Len(MaskEdBox1(1).Text) <= 0 Then
        strMsg = Replace(SSPanel2(18).Caption, " ", "") & "이 입력되지 않았습니다."
        GoTo InputErr
    End If
    '필수 입력항목 검사
    For i = 1 To 16
        If Len(txtInput(i).Text) <= 0 Then
            strMsg = Replace(SSPanel2(i).Caption, " ", "") & " (가)이 입력되지 않았습니다."
            GoTo InputErr
        End If
    Next i
    
    ' 일자를 확인한다.
    strDate(0) = Format(dtInput(0).Value, "YYYY-MM-DD")
    strDate(1) = Format(dtInput(1).Value, "YYYY-MM-DD")
    strDate(3) = Format(dtInput(3).Value, "YYYY-MM-DD")
    
    If strDate(0) < strDate(3) Then
        MsgBox "접수일이 사고접수일보다 적을수 없습니다.", vbInformation, "확인"
        SaveData = 0
        Exit Function
        
    ElseIf strDate(0) < strDate(1) Then
        MsgBox "접수일이 구입일자보다 적을수 없습니다.", vbInformation, "확인"
        SaveData = 0
        Exit Function
        
    ElseIf strDate(1) > strDate(3) Then
        MsgBox "구입일자가  사고접수일보다 클수 없습니다.", vbInformation, "확인"
        SaveData = 0
        Exit Function
    End If
    
    For i = 1 To 1
        MaskEdBox1(i).Text = Replace(MaskEdBox1(i).Text, "'", "")
    Next i
    
    For i = 1 To 18
        txtInput(i).Text = Replace(txtInput(i).Text, "'", "")
    Next i
    
    If iReadRecord = 0 Then
        ' 신규 입력일 경우
        ' 저장
        Query = "SELECT  일련번호 FROM TB_사고품 ORDER BY 일련번호 DESC"
        Set SUBRs = New ADODB.Recordset
        SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        If SUBRs.EOF Then
            lDBCount = 1
        Else
            lDBCount = CDbl(SUBRs!일련번호) + 1
        End If
        SUBRs.Close
        
        Query = "INSERT INTO TB_사고품(일련번호, 접수일자, 성명, 고객전화, 주소, 휴대전화, "
        Query = Query & " 의류명, 상표, 구입일자, 색상, 구입처, 최초택번호, 최종택번호, 구입형태, 최초입고일, 최종입고일, "
        Query = Query & " 구입가격, 사고접수일, 사고종류, 사고내용, 사고의견, 보상금액, 합의금액, 처리유무 ) "
        Query = Query & "VALUES ('" & lDBCount & "', "
        Query = Query & "'" & Trim(strDate(0)) & "', "
        Query = Query & "'" & Trim(MaskEdBox1(1).Text) & "', "
        Query = Query & "'" & Trim(txtInput(2).Text) & "', "
        Query = Query & "'" & Trim(txtInput(1).Text) & "', "
        Query = Query & "'" & Trim(txtInput(3).Text) & "', "
        Query = Query & "'" & Trim(txtInput(4).Text) & "', "
        Query = Query & "'" & Trim(txtInput(5).Text) & "', "
        Query = Query & "'" & Trim(strDate(1)) & "', "
        Query = Query & "'" & Trim(txtInput(6).Text) & "', "
        Query = Query & "'" & Trim(txtInput(7).Text) & "', "
        Query = Query & "'" & Trim(txtInput(8).Text) & "', "
        Query = Query & "'" & Trim(txtInput(9).Text) & "', "
        Query = Query & "'" & Trim(txtInput(10).Text) & "', "
        Query = Query & "'" & Trim(txtInput(11).Text) & "', "
        Query = Query & "'" & Trim(txtInput(12).Text) & "', "
        Query = Query & "'" & Trim(Replace(txtInput(13).Text, ",", "")) & "', "
        Query = Query & "'" & Trim(strDate(3)) & "', "
        Query = Query & "'" & Trim(txtInput(14).Text) & "', "
        Query = Query & "'" & Trim(txtInput(15).Text) & "', "
        Query = Query & "'" & Trim(txtInput(16).Text) & "', "
        Query = Query & "'" & Trim(txtInput(17).Text) & "', "
        Query = Query & "'" & Trim(txtInput(18).Text) & "', "
        Query = Query & "'" & Trim(Combo1.Text) & "') "
        
        ADOCon.Execute Query
        SaveData = 1
        Exit Function
        
    ElseIf iEditMode = 2 Then
        ' 수정일 경우
        Query = "UPDATE  사고품 SET "
        Query = Query & " 접수일자 = '" & Trim(strDate(0)) & "', "
        Query = Query & " 성명 = '" & Trim(MaskEdBox1(1).Text) & "', "
        Query = Query & " 고객전화 = '" & Trim(txtInput(2).Text) & "', "
        Query = Query & " 주소 = '" & Trim(txtInput(1).Text) & "', "
        Query = Query & " 휴대전화 = '" & Trim(txtInput(3).Text) & "', "
        Query = Query & " 의류명 = '" & Trim(txtInput(4).Text) & "', "
        Query = Query & " 상표 = '" & Trim(txtInput(5).Text) & "', "
        Query = Query & " 구입일자 = '" & Trim(strDate(1)) & "', "
        Query = Query & " 색상 = '" & Trim(txtInput(6).Text) & "', "
        Query = Query & " 구입처 = '" & Trim(txtInput(7).Text) & "', "
        Query = Query & " 최초택번호 = '" & Trim(txtInput(8).Text) & "', "
        Query = Query & " 최종택번호 = '" & Trim(txtInput(9).Text) & "', "
        Query = Query & " 구입형태 = '" & Trim(txtInput(10).Text) & "', "
        Query = Query & " 최초입고일 = '" & Trim(txtInput(11).Text) & "', "
        Query = Query & " 최종입고일 = '" & Trim(txtInput(12).Text) & "', "
        Query = Query & " 구입가격 = '" & Trim(txtInput(13).Text) & "', "
        Query = Query & " 사고접수일 = '" & Trim(strDate(3)) & "', "
        Query = Query & " 사고종류 = '" & Trim(txtInput(14).Text) & "', "
        Query = Query & " 사고내용 = '" & Trim(txtInput(15).Text) & "', "
        Query = Query & " 사고의견 = '" & Trim(txtInput(16).Text) & "',  "
        Query = Query & " 보상금액 = '" & Trim(txtInput(17).Text) & "',  "
        Query = Query & " 합의금액 = '" & Trim(txtInput(18).Text) & "',  "
        Query = Query & " 처리유무 = '" & Trim(Combo1.Text) & "'  "
        Query = Query & " WHERE 일련번호 = " & iReadRecord & ""
        
        ADOCon.Execute Query
        SaveData = 1
        Exit Function
    End If
    
    
InputErr:
' 오류 발생시
    MsgBox strMsg, vbInformation, "입력확인"
    SaveData = 3
    Exit Function
ErrRtn:
    ' 오류 발생시
    MsgBox Err.Number & ", " & Err.Description, vbInformation, "입력확인"
    SaveData = 4
End Function

Private Function DeleteData(iReadRecord) As Integer
'+------------------------------------------------------
'+
'+ 2003/03/22
'+
'+  - 전달값
'+      iReadRecord : 현재 읽은 레코드의 값
'+  - 리턴값
'+      0           : 삭제 오류
'+      1           : 삭제 성공
'+      4           : 기타 오류
'+  - 루틴설명
'+      1. 모드 항목 확인
'+      2. iReadRecord 와 일치하는 자료 삭제
'+
'+------------------------------------------------------
Dim strMsg As String
    On Error GoTo deleteErr

    ' 모드를 확인한다.
    If iReadRecord = 0 Then
        DeleteData = 0
        Exit Function
    End If
    
    ' 삭제
    If iReadRecord <> 0 Then
        Query = "  DELETE  FROM TB_사고품 "
        Query = Query & " WHERE 일련번호 = " & iReadRecord & ""
        ADOCon.Execute Query
        
        DeleteData = 1
        Exit Function
    End If
    
    
deleteErr:
    ' 오류 발생시
    MsgBox Err.Number & ", " & Err.Source, vbInformation, "입력확인"
    DeleteData = 4
    Exit Function
End Function

'+------------------------------------------------------
'+
'+ 2003/03/22
'+
'+  - 전달값
'+
'+  - 리턴값
'+
'+  - 루틴설명
'+      1. 조회 화면을 초기화 한다
'+
'+------------------------------------------------------
Private Sub SpreadClear()
    Dim j As Integer
    
    Query = " SELECT * FROM TB_사고품 "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    If Not SUBRs.BOF Then SUBRs.MoveFirst
    
    With fpSpread1
        .ReDraw = False
        .RowHeight(0) = 20
        .MaxRows = 0
        .Row = -1
        .MaxCols = SUBRs.Fields.Count

        For i = 1 To SUBRs.Fields.Count
            .Col = i
            .ColWidth(i) = 10
            .CellType = CellTypeStaticText
            .TypeVAlign = TypeVAlignCenter
            .TypeHAlign = TypeHAlignCenter
            Select Case i
                Case 1
                    '일련번호
                    .ColWidth(i) = 5
                Case 2
                    '접수일자
                    .ColWidth(i) = 10
                Case 3
                    '성명
                    .ColWidth(i) = 8
                Case 4
                    '고객전화
                    .ColWidth(i) = 15
                Case 5
                    '주소
                    .ColWidth(i) = 30
                    .TypeHAlign = TypeHAlignLeft
                Case 6
                    '고객전화
                    .ColWidth(i) = 15
                Case 17
                    .ColWidth(i) = 20
                    .TypeHAlign = TypeHAlignLeft
                Case 19
                    .ColWidth(i) = 30
                    .TypeHAlign = TypeHAlignLeft
                Case 20
                    .ColWidth(i) = 30
                    .TypeHAlign = TypeHAlignLeft
            End Select
            .SetText i, 0, SUBRs.Fields.Item(i - 1).Name
        Next i
        
        .ReDraw = True
    End With
    SUBRs.Close
           
End Sub

'+------------------------------------------------------
'+ 2003/03/22
'+  - 전달값
'+  - 리턴값
'+  - 루틴설명
'+      1. 입력된 내용으로 조회하여 스프레드에 출력한다.
'+------------------------------------------------------
Private Sub DisplaySpread()
    Dim j As Integer

    Query = " SELECT * FROM TB_사고품 WHERE "
    
    If Len(txtInput(30).Text) > 0 Then
        Query = Query & " 성명 LIKE '%" & txtInput(30).Text & "%'  AND  "
    ElseIf Len(txtInput(31).Text) > 0 Then
        Query = Query & " 고객전화 LIKE '%" & txtInput(31).Text & "%'  AND  "
    End If
    
    Query = Query & " ( 접수일자 >= '" & Format(dtInput(4).Value, "YYYY-MM-DD") & "' AND 접수일자 <= '" & Format(dtInput(0).Value, "YYYY-MM-DD") & "') "
    Query = Query & " ORDER BY 일련번호 "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    
    If Not SUBRs.BOF Then SUBRs.MoveLast
    If SUBRs.RecordCount <= 0 Then
        MsgBox "해당 내용이 없습니다.", vbInformation, "확인"
        Exit Sub
    End If
    
    With fpSpread1
        .ReDraw = False
        SUBRs.MoveFirst
        For i = 1 To SUBRs.RecordCount
            If .MaxRows < SUBRs.RecordCount Then
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Action = ActionInsertRow
                .RowHeight(.MaxRows) = .RowHeight(0) ' 마지막 라인의 높이를 맞춘다.
            End If

            For j = 1 To SUBRs.Fields.Count
                If j = 2 Or j = 9 Or j = 18 Then
                    .SetText j, i, Format(Trim(SUBRs.Fields(j - 1)), "YYYY-MM-DD")
                ElseIf j = 17 Or j = 22 Or j = 23 Then
                    .SetText j, i, Format(Trim(SUBRs.Fields(j - 1)), "#,###")
                Else
                .SetText j, i, Trim(SUBRs.Fields(j - 1))
                End If
            Next j
            .Row = .Row + 1
            SUBRs.MoveNext
        Next i
        .ReDraw = True
    End With
    
    SUBRs.Close

End Sub

'+------------------------------------------------------
'+
'+ 2003/03/22
'+
'+  - 전달값
'+      1. 사고품의 일련 번호
'+  - 리턴값
'+
'+  - 루틴설명
'+      1. 전달된 일련번호의 내용을 폼에 출력 한다.
'+
'+------------------------------------------------------
Private Sub DisplayFromData(strCode As String)
    Dim j As Integer

    Query = " SELECT * FROM TB_사고품 WHERE 일련번호 = " & strCode & ""
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If SUBRs.BOF Or SUBRs.EOF Then
        Exit Sub
    End If
    
    iReadRecord = CInt(strCode)
    
    MaskEdBox1(1).Text = SUBRs!성명
    txtInput(1).Text = SUBRs!주소 & ""
    txtInput(2).Text = SUBRs!고객전화 & ""
    txtInput(3).Text = SUBRs!휴대전화 & ""
    txtInput(4).Text = SUBRs!의류명 & ""
    txtInput(5).Text = SUBRs!상표 & ""
    If IsDate(Format(SUBRs!구입일자, "YYYY-MM-DD")) = True Then
        dtInput(1).Value = Format(SUBRs!구입일자, "YYYY-MM-DD")
    End If
    txtInput(6).Text = SUBRs!색상 & ""
    txtInput(7).Text = SUBRs!구입처 & ""
    txtInput(8).Text = SUBRs!최초택번호 & ""
    txtInput(9).Text = SUBRs!최종택번호 & ""
    txtInput(10).Text = SUBRs!구입형태 & ""
    txtInput(11).Text = SUBRs!최초입고일 & ""
    txtInput(12).Text = SUBRs!최종입고일 & ""
    txtInput(13).Text = SUBRs!구입가격 & ""
    dtInput(3).Value = Format(SUBRs!사고접수일 & "", "YYYY-MM-DD")
    txtInput(14).Text = SUBRs!사고종류 & ""
    txtInput(15).Text = SUBRs!사고내용 & ""
    txtInput(16).Text = SUBRs!사고의견 & ""
    txtInput(17).Text = SUBRs!보상금액 & ""
    txtInput(18).Text = SUBRs!합의금액 & ""
    Combo1.Text = SUBRs!처리유무 & ""

    SUBRs.Close
End Sub


Private Sub Combo1_Change()
    If iReadRecord <> 0 Then
        iEditMode = 2
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim nDelCount As Integer
    
    Select Case Index
    '+------------------------------------------------------
    '+
    '+ 2003/03/11
    '+
    '+루틴설명
    '+  1. 입력
    '+  2. 사고의 내용을 입력할 수 있도록 화면을 초기화 한다.
    '+
    '+------------------------------------------------------
        Case 0
            If iEditMode = 1 Then
                If MsgBox("입력을 취소하시겠습니까?", vbInformation + vbYesNo, "확인") = vbYes Then
                    TextBoxClear
                End If
            End If
            iReadRecord = 0
            iEditMode = 1
            SSPanel3.Visible = False
            MaskEdBox1(1).SetFocus
            Exit Sub
    '+------------------------------------------------------
    '+
    '+ 2003/03/11
    '+
    '+루틴설명
    '+  1. 조회
    '+  2. 사고의 내용을 입력할 수 있도록 화면을 초기화 한다.
    '+
    '+------------------------------------------------------
        Case 1
            If iEditMode = 1 Then
                If MsgBox("입력을 취소하시겠습니까?", vbInformation + vbYesNo, "확인") = vbNo Then
                    Exit Sub
                End If
            End If
            iEditMode = 0
            iReadRecord = 0
            TextBoxClear
            SSPanel3.Visible = True
            SpreadClear
            DisplaySpread
            
            
        
    '+------------------------------------------------------
    '+
    '+ 2003/03/11
    '+
    '+루틴설명
    '+  1. 저장
    '+  2. 사고의 내용을 입력할 수 있도록 화면을 초기화 한다.
    '+
    '+------------------------------------------------------
        Case 2
            SSPanel3.Visible = False
            ' 입력 내용이 없을 경우
            If iEditMode = 0 Then Exit Sub
            ' 신규와 수정을 구분한다. (iReadRecord )
            iErrorCheck = SaveData(iReadRecord)
            
            If iErrorCheck = 1 Then
                iReadRecord = 0
                iEditMode = 0
                
                ' 저장된 자료를 전송한다.
                Call SSPanel2_Click(17)
                
                MsgBox "자료가 정상 저장되었습니다.", vbInformation, "확인"
                TextBoxClear
                Exit Sub
                
            ElseIf iErrorCheck = 2 Then
                MsgBox "자료 저장중 오류가 발생했습니다.", vbCritical, vbInformation, "확인"
                Exit Sub
            ElseIf iErrorCheck = 3 Or iErrorCheck = 4 Then
            
            End If
    '+------------------------------------------------------
    '+
    '+ 2003/03/11
    '+
    '+루틴설명
    '+  1. 삭제
    '+  2. 사고의 내용을 입력할 수 있도록 화면을 초기화 한다.
    '+
    '+------------------------------------------------------
        Case 3
            If SSPanel3.Visible = False Then
                If iReadRecord = 0 Then Exit Sub
                If MsgBox("선택된 내용을 삭제하시겠습니까?, 복원 불가능", vbInformation + vbYesNo, "확인") = vbYes Then
                    iErrorCheck = DeleteData(iReadRecord)
                    If iErrorCheck = 1 Then
                        MsgBox "자료가 정상적으로 삭제 되었습니다", vbInformation, "확인"
                        TextBoxClear
                        Exit Sub
                    End If
                End If
            Else
                ' 조회에서 삭제일 경우
                nDelCount = 0
                For i = 1 To fpSpread1.MaxRows
                    fpSpread1.Row = i
                    fpSpread1.Col = 1
                    If fpSpread1.BackColor = vbYellow Then
                        iErrorCheck = DeleteData(CInt(fpSpread1.Text))
                        If iErrorCheck = 1 Then
                            nDelCount = nDelCount + 1
                        Else
                            MsgBox "[ " & fpSpread1.Text & "번 ]의 자료가 삭제중 오류가 발생했습니다.", vbCritical, vbInformation, "확인"
                        End If
                    End If
                Next i
                If nDelCount > 0 Then
                    SpreadClear
                    DisplaySpread
                    MsgBox "[ " & nDelCount & "건 ]의 자료를 삭제했습니다."
                End If
            End If
    '+------------------------------------------------------
    '+
    '+ 2003/03/11
    '+
    '+루틴설명
    '+  1. 삭제
    '+  2. 사고의 내용을 입력할 수 있도록 화면을 초기화 한다.
    '+
    '+------------------------------------------------------
        Case 4
        ' 인쇄
        fpSpread1.Row = 1
        fpSpread1.Col = 1
        i = 0
        If iEditMode <> 0 Then
            MsgBox "조회 기능에서만 인쇄가 가능합니다.", vbInformation, "확인"
            Exit Sub
        End If
        
        Do While fpSpread1.Text <> ""
            If fpSpread1.BackColor = vbYellow Then
                i = i + 1
                If i = 1 Then
                    If vbYes <> MsgBox("선택된 내역을 인쇄 하시겠습니까?", vbInformation + vbYesNo, "확인") Then Exit Sub
                End If
                Call PrintSagoReport(cdPrt, fpSpread1.Text)
            End If
            If fpSpread1.Row = fpSpread1.MaxRows Then Exit Do
            fpSpread1.Row = fpSpread1.Row + 1
        Loop
        If i <= 0 Then MsgBox "인쇄할 내역을 선택하십시요.", vbInformation, "확인"
    End Select

End Sub

Private Sub dtInput_Change(Index As Integer)
    If iReadRecord <> 0 Then
        iEditMode = 2
    End If
End Sub

Private Sub Form_Load()
    'TitleSet "사고품 작성"
    iEditMode = 0
    SSPanel3.Visible = False
    dtInput(0).Value = Date
    dtInput(1).Value = Date
    dtInput(3).Value = Date
    dtInput(4).Value = DateAdd("yyyy", -1, Date)
    
    Combo1.AddItem "완결"
    Combo1.AddItem "보류"
    Combo1.AddItem "미결"
    Combo1.AddItem "처리중"
    Combo1.AddItem "기타"
End Sub

Private Sub MaskEdBox1_GotFocus(Index As Integer)
    Dim hiMe As Long
    
    Toggle_Check = True
    
    ' //KEYCODE 123 번은 펑션키12번(F12)
    ' //특정키를 입력하려면 아래 KEYCODE만 바꿔주면됨
    If Toggle_Check = True Then
        ' // 한글로 바꾸기
        hiMe = ImmGetContext(MaskEdBox1(1).hwnd)
        ImmSetConversionStatus hiMe, IME_HANGUL, IME_NONE
        Toggle_Check = False
    Else
        ' // 영어로 바꾸기
        hiMe = ImmGetContext(MaskEdBox1(1).hwnd)
        ImmSetConversionStatus hiMe, IME_ENGLISH, IME_NONE
        Toggle_Check = True
    End If
    
    Select Case Index
        Case 1
            MaskEdBox1(1).BackColor = "&H0080FF80"
            txtInput(100).BackColor = "&H0080FF80"
    End Select
End Sub

Private Sub MaskEdBox1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 14 Then
        ' 조회에서 소비자명일 경우
        If KeyCode = vbKeyReturn Then
            DisplaySpread
        End If
    End If
    
    If KeyCode = vbKeyReturn Then
        SendKeys "{tab}"
    Else
        If iReadRecord = 0 Then
            iEditMode = 1
        Else
            iEditMode = 2
        End If
    End If
End Sub

Private Sub MaskEdBox1_LostFocus(Index As Integer)
    Select Case Index
        Case 1
            MaskEdBox1(1).BackColor = "&H00FFFFFF"
            txtInput(100).BackColor = "&H00FFFFFF"
    End Select
End Sub

Private Sub SSPanel2_Click(Index As Integer)
    Dim sSendData   As String
    Dim sYN         As String
    
    If Trim(가맹점정보.가맹점코드) = "000000" Then
        MsgBox "가맹점 정보가 올바르지 않습니다.", vbCritical, "경고"
        Exit Sub
    End If
    
    If Server_Connection(HostCon) = False Then
        Set HostCon = Nothing
        Exit Sub
    End If


    If Index = 17 Then
        sSendData = SendTable_DateCheck("사고품", HostCon, sYN)
        
        'Call SendTable_사고품(sSendData, HostCon, frmMain.ProgressBar1)
        Call SendTable_사고품(sSendData, HostCon, ProgressBar1)
    End If
    
    Set HostCon = Nothing
    
End Sub

Private Sub txtInput_GotFocus(Index As Integer)
    Dim hiMe As Long

    txtInput(Index).BackColor = "&H0080FF80"
    
    Toggle_Check = True
    
    ' //KEYCODE 123 번은 펑션키12번(F12)
    ' //특정키를 입력하려면 아래 KEYCODE만 바꿔주면됨
    If Toggle_Check = True Then
        ' // 한글로 바꾸기
        hiMe = ImmGetContext(txtInput(Index).hwnd)
        ImmSetConversionStatus hiMe, IME_HANGUL, IME_NONE
        Toggle_Check = False
    Else
        ' // 영어로 바꾸기
        hiMe = ImmGetContext(txtInput(Index).hwnd)
        ImmSetConversionStatus hiMe, IME_ENGLISH, IME_NONE
        Toggle_Check = True
    End If
    
    If Index = 100 Then
    ' 소비자명의 뒷부분선택시
        MaskEdBox1(1).BackColor = "&H0080FF80"
        MaskEdBox1(1).SetFocus
    End If
End Sub

Private Sub txtInput_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 30 Or Index = 31 Then
        ' 조회에서 소비자명, 전화 번호 일 경우
        If KeyCode = vbKeyReturn Then
            SpreadClear
            DisplaySpread
            Exit Sub
        End If
    End If
    
    '소비자의견에서 엔터시 저장여부 확인
    If KeyCode = vbKeyReturn And Index = 18 Then
        If MsgBox("내용을 저장하시겠습니까 ?", vbYesNo + vbInformation, "저장") = vbYes Then
            Call Command1_Click(2)
        End If
    ElseIf KeyCode = vbKeyReturn Then
        SendKeys "{tab}"
    Else
        If iReadRecord = 0 And (Index <> 30 And Index <> 31) Then
            iEditMode = 1
        Else
            iEditMode = 2
        End If
    End If
End Sub

Private Sub txtInput_LostFocus(Index As Integer)
    txtInput(Index).BackColor = "&H00FFFFFF"
End Sub

Private Sub fpSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
    ReDim sValue(2)
    
    sValue(0) = "0"
    
    fpSpread1.Row = fpSpread1.ActiveRow
    fpSpread1.Col = 1
    
    sValue(1) = fpSpread1.Text
    
    DisplayFromData CStr(sValue(1))
    
    SSPanel3.Visible = False

End Sub

Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If NewRow <> -1 Then
        
        fpSpread1.Row = NewRow
        fpSpread1.Col = -1
        If fpSpread1.BackColor = vbYellow Then
            fpSpread1.BackColor = vbWhite
        Else
            fpSpread1.BackColor = vbYellow
        End If
    End If

End Sub
