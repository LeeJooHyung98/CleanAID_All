VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMask32.ocx"
Begin VB.Form P_06007 
   Caption         =   "[전사업장] 사고처리 접수"
   ClientHeight    =   12105
   ClientLeft      =   705
   ClientTop       =   3450
   ClientWidth     =   17190
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_06007.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12105
   ScaleWidth      =   17190
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   12105
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17190
      _ExtentX        =   30321
      _ExtentY        =   21352
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_06007.frx":058A
      Begin Threed.SSPanel panInput 
         Height          =   765
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   17160
         _ExtentX        =   30268
         _ExtentY        =   1349
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   6
            Left            =   8040
            Style           =   2  '드롭다운 목록
            TabIndex        =   2
            Top             =   60
            Width           =   6315
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
            Format          =   64749568
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
         Begin MSComDlg.CommonDialog cdPicture 
            Left            =   105
            Top             =   330
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            DialogTitle     =   "사고 제품 이미지파일 선택"
         End
      End
      Begin Threed.SSPanel panDetail 
         Height          =   11295
         Left            =   15
         TabIndex        =   6
         Top             =   795
         Width           =   17160
         _ExtentX        =   30268
         _ExtentY        =   19923
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin Threed.SSFrame SSFrame6 
            Height          =   1095
            Left            =   120
            TabIndex        =   7
            Top             =   6660
            Width           =   14955
            _ExtentX        =   26379
            _ExtentY        =   1931
            _Version        =   262144
            Caption         =   "보 상 산 정 기 준"
            Begin VB.ComboBox cboInput 
               Height          =   315
               Index           =   7
               Left            =   1920
               Style           =   2  '드롭다운 목록
               TabIndex        =   14
               Top             =   300
               Width           =   1875
            End
            Begin VB.ComboBox cboInput 
               Height          =   315
               Index           =   8
               Left            =   5520
               Style           =   2  '드롭다운 목록
               TabIndex        =   13
               Top             =   300
               Width           =   1875
            End
            Begin VB.ComboBox cboInput 
               Height          =   315
               Index           =   9
               Left            =   9120
               Style           =   2  '드롭다운 목록
               TabIndex        =   12
               Top             =   300
               Width           =   1875
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   13
               Left            =   12720
               TabIndex        =   11
               Top             =   300
               Width           =   1875
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   14
               Left            =   1920
               TabIndex        =   10
               Top             =   660
               Width           =   1875
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   15
               Left            =   5520
               TabIndex        =   9
               Top             =   660
               Width           =   1875
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   16
               Left            =   9120
               TabIndex        =   8
               Top             =   660
               Width           =   1875
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   20
               Left            =   300
               TabIndex        =   15
               Top             =   300
               Width           =   1635
               _ExtentX        =   2884
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
               TabIndex        =   16
               Top             =   300
               Width           =   1635
               _ExtentX        =   2884
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
               TabIndex        =   17
               Top             =   300
               Width           =   1635
               _ExtentX        =   2884
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
               TabIndex        =   18
               Top             =   300
               Width           =   1635
               _ExtentX        =   2884
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
               TabIndex        =   19
               Top             =   660
               Width           =   1635
               _ExtentX        =   2884
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
               TabIndex        =   20
               Top             =   660
               Width           =   1635
               _ExtentX        =   2884
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
               TabIndex        =   21
               Top             =   660
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "배 상 비 율"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSMask.MaskEdBox mskInput 
               Height          =   315
               Index           =   2
               Left            =   12720
               TabIndex        =   22
               Top             =   660
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "#,##0"
               PromptChar      =   "_"
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   32
               Left            =   11100
               TabIndex        =   23
               Top             =   660
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "보상산정금액"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
         End
         Begin Threed.SSCommand cmdSubBtn 
            Height          =   435
            Index           =   0
            Left            =   12480
            TabIndex        =   24
            Top             =   6060
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   767
            _Version        =   262144
            Caption         =   "이미지추가"
         End
         Begin Threed.SSFrame SSFrame5 
            Height          =   735
            Left            =   120
            TabIndex        =   25
            Top             =   7860
            Width           =   14955
            _ExtentX        =   26379
            _ExtentY        =   1296
            _Version        =   262144
            Caption         =   "담 당 자 의 견"
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   10
               Left            =   300
               MaxLength       =   80
               TabIndex        =   28
               Top             =   300
               Width           =   4755
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   11
               Left            =   5040
               MaxLength       =   80
               TabIndex        =   27
               Top             =   300
               Width           =   4755
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   12
               Left            =   9780
               MaxLength       =   80
               TabIndex        =   26
               Top             =   300
               Width           =   4755
            End
         End
         Begin Threed.SSFrame SSFrame4 
            Height          =   2535
            Left            =   120
            TabIndex        =   29
            Top             =   3960
            Width           =   11235
            _ExtentX        =   19817
            _ExtentY        =   4471
            _Version        =   262144
            Caption         =   "구 입 자 정 보"
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   6
               Left            =   1920
               MaxLength       =   10
               TabIndex        =   37
               Top             =   300
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   8
               Left            =   1920
               MaxLength       =   40
               TabIndex        =   36
               Top             =   660
               Width           =   9135
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   7
               Left            =   7320
               MaxLength       =   15
               TabIndex        =   35
               Top             =   300
               Width           =   3735
            End
            Begin VB.ComboBox cboInput 
               Height          =   315
               Index           =   4
               Left            =   1920
               Style           =   2  '드롭다운 목록
               TabIndex        =   34
               Top             =   1380
               Width           =   3735
            End
            Begin VB.ComboBox cboInput 
               Height          =   315
               Index           =   5
               Left            =   7320
               Style           =   2  '드롭다운 목록
               TabIndex        =   33
               Top             =   1380
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   9
               Left            =   1920
               MaxLength       =   50
               TabIndex        =   31
               Top             =   2100
               Width           =   9075
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   17
               Left            =   7320
               MaxLength       =   14
               TabIndex        =   30
               Top             =   1020
               Width           =   3735
            End
            Begin MSMask.MaskEdBox mskInput 
               Height          =   315
               Index           =   1
               Left            =   1920
               TabIndex        =   32
               Top             =   1740
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "#,##0"
               PromptChar      =   "_"
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   11
               Left            =   300
               TabIndex        =   38
               Top             =   300
               Width           =   1635
               _ExtentX        =   2884
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
               TabIndex        =   39
               Top             =   660
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "고 객 의 견"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   22
               Left            =   5700
               TabIndex        =   40
               Top             =   300
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "전 화 번 호"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   21
               Left            =   300
               TabIndex        =   41
               Top             =   1020
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "처 리 일 자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtInput 
               Height          =   315
               Index           =   5
               Left            =   1920
               TabIndex        =   42
               Top             =   1020
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   64749568
               CurrentDate     =   36684
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   12
               Left            =   300
               TabIndex        =   43
               Top             =   1380
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "크레임구분"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   14
               Left            =   5700
               TabIndex        =   44
               Top             =   1380
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "보 상 구 분"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   15
               Left            =   300
               TabIndex        =   45
               Top             =   1740
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "보 상 금 액"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   16
               Left            =   300
               TabIndex        =   46
               Top             =   2100
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "제 품 정 보"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   33
               Left            =   5700
               TabIndex        =   47
               Top             =   1020
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "휴대폰 번호"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
         End
         Begin Threed.SSFrame SSFrame2 
            Height          =   2475
            Left            =   120
            TabIndex        =   48
            Top             =   1320
            Width           =   11235
            _ExtentX        =   19817
            _ExtentY        =   4366
            _Version        =   262144
            Caption         =   "품 목 정 보"
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   1
               Left            =   1860
               MaxLength       =   20
               TabIndex        =   55
               Top             =   960
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   2
               Left            =   7260
               MaxLength       =   20
               TabIndex        =   54
               Top             =   960
               Width           =   3735
            End
            Begin VB.ComboBox cboInput 
               Height          =   315
               Index           =   3
               Left            =   7260
               TabIndex        =   53
               Top             =   240
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   3
               Left            =   1860
               MaxLength       =   10
               TabIndex        =   52
               Top             =   1320
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   4
               Left            =   7260
               MaxLength       =   20
               TabIndex        =   51
               Top             =   1680
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   5
               Left            =   1860
               MaxLength       =   10
               TabIndex        =   50
               Top             =   2040
               Width           =   3735
            End
            Begin MSMask.MaskEdBox mskInput 
               Height          =   315
               Index           =   0
               Left            =   7260
               TabIndex        =   49
               Top             =   2040
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "#,##0"
               PromptChar      =   "_"
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   5
               Left            =   240
               TabIndex        =   56
               Top             =   240
               Width           =   1635
               _ExtentX        =   2884
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
               TabIndex        =   57
               Top             =   960
               Width           =   1635
               _ExtentX        =   2884
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
               TabIndex        =   58
               Top             =   960
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "브  랜  드"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtInput 
               Height          =   315
               Index           =   1
               Left            =   1860
               TabIndex        =   59
               Top             =   240
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   64749568
               CurrentDate     =   36686
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   3
               Left            =   5640
               TabIndex        =   60
               Top             =   240
               Width           =   1635
               _ExtentX        =   2884
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
               TabIndex        =   61
               Top             =   600
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "출 고 일 자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtInput 
               Height          =   315
               Index           =   2
               Left            =   1860
               TabIndex        =   62
               Top             =   600
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   64749568
               CurrentDate     =   36686
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   8
               Left            =   5640
               TabIndex        =   63
               Top             =   600
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "인 도 일 자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtInput 
               Height          =   315
               Index           =   3
               Left            =   7260
               TabIndex        =   64
               Top             =   600
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   64749568
               CurrentDate     =   36686
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   9
               Left            =   240
               TabIndex        =   65
               Top             =   1320
               Width           =   1635
               _ExtentX        =   2884
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
               TabIndex        =   66
               Top             =   1680
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "구 입 일 자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtInput 
               Height          =   315
               Index           =   4
               Left            =   1860
               TabIndex        =   67
               Top             =   1680
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   64749568
               CurrentDate     =   36686
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   25
               Left            =   5640
               TabIndex        =   68
               Top             =   1680
               Width           =   1635
               _ExtentX        =   2884
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
               TabIndex        =   69
               Top             =   2040
               Width           =   1635
               _ExtentX        =   2884
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
               TabIndex        =   70
               Top             =   2040
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "구 입 가 격"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
         End
         Begin Threed.SSFrame SSFrame1 
            Height          =   1095
            Left            =   120
            TabIndex        =   71
            Top             =   120
            Width           =   11235
            _ExtentX        =   19817
            _ExtentY        =   1931
            _Version        =   262144
            Caption         =   "접 수 내 역"
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   0
               Left            =   1860
               Locked          =   -1  'True
               TabIndex        =   75
               Top             =   300
               Width           =   3735
            End
            Begin VB.ComboBox cboInput 
               Height          =   315
               Index           =   2
               Left            =   7260
               Style           =   2  '드롭다운 목록
               TabIndex        =   74
               Top             =   660
               Width           =   3735
            End
            Begin VB.ComboBox cboInput 
               Height          =   315
               Index           =   1
               Left            =   1860
               Style           =   2  '드롭다운 목록
               TabIndex        =   73
               Top             =   660
               Width           =   3735
            End
            Begin VB.ComboBox cboInput 
               Height          =   315
               Index           =   0
               Left            =   7260
               Style           =   2  '드롭다운 목록
               TabIndex        =   72
               Top             =   300
               Width           =   3735
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   1
               Left            =   240
               TabIndex        =   76
               Top             =   300
               Width           =   1635
               _ExtentX        =   2884
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
               TabIndex        =   77
               Top             =   300
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "사업장 명칭"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   17
               Left            =   240
               TabIndex        =   78
               Top             =   660
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "담당자명"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   18
               Left            =   5640
               TabIndex        =   79
               Top             =   660
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "체인점 명칭"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
         End
         Begin Threed.SSFrame SSFrame3 
            Height          =   5835
            Left            =   11460
            TabIndex        =   80
            Top             =   120
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   10292
            _Version        =   262144
            Caption         =   "사고 제품 이미지"
            Begin VB.PictureBox pctPicture 
               BackColor       =   &H8000000E&
               Height          =   5355
               Left            =   120
               ScaleHeight     =   5295
               ScaleWidth      =   3315
               TabIndex        =   81
               Top             =   300
               Width           =   3375
            End
         End
         Begin Threed.SSCommand cmdSubBtn 
            Height          =   435
            Index           =   1
            Left            =   13800
            TabIndex        =   82
            Top             =   6060
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   767
            _Version        =   262144
            Caption         =   "이미지제거"
         End
      End
   End
End
Attribute VB_Name = "P_06007"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents CPrt    As CCAIDPrinter
Attribute CPrt.VB_VarHelpID = -1

Dim RS01 As ADODB.Recordset
Dim RS02 As ADODB.Recordset

Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Dim sPictureFile As String

Private Sub cboInput_Click(Index As Integer)
    Select Case Index
        Case 0
            Dim sCode As String
            sCode = Trim(Mid(Trim(cboInput(0)) & Space(10), 2, 4))
    
            Call Get_가맹점리스트(cboInput(2), sCode)
        
        Case 3
        
            ReDim sValue(2)
            
            sValue(0) = Mid(cboInput(0).Text, 2, 3)
            sValue(1) = Format(dtInput(1).Value, "YYYY-MM-DD")
            sValue(2) = Mid(cboInput(3).Text, 1, 1) & Mid(cboInput(3).Text, 3, 3)
            
            Set RS02 = New ADODB.Recordset
            Set RS02 = ExecPro("SP_06001_07", sValue(), Err_Num, Err_Dec)
            
            If Err_Num = 0 Then
                If RS02.RecordCount = 0 Then
                
                Else
                    If Not IsNull(RS02!품명) Then txtInput(1).Text = RS02!품명
                    If Not IsNull(RS02!브랜드) Then txtInput(2).Text = RS02!브랜드
                    If Not IsNull(RS02!색상) Then txtInput(3).Text = RS02!색상
                End If
            Else
                MsgBox "[" & Err_Num & "] " & Err_Dec
                Exit Sub
            End If
        Case 6
            dtInput(0).Value = Format(Mid(cboInput(6).Text, 1, 10), "YYYY-MM-DD")
        Case 7, 8, 9
            If cboInput(7).Text <> "" And cboInput(8).Text <> "" And cboInput(9).Text <> "" Then
                ReDim sValue(3)
                
                sValue(0) = "0"
                sValue(1) = Mid(cboInput(7).Text, 2, 3)
                sValue(2) = Mid(cboInput(8).Text, 2, 3)
                sValue(3) = Mid(cboInput(9).Text, 2, 3)
                
                Set RS02 = New ADODB.Recordset
                Set RS02 = ExecPro("SP_06001_06", sValue(), Err_Num, Err_Dec)
        
                If RS02.RecordCount = 0 Then
                    txtInput(13).Text = ""
                    Exit Sub
                Else
                    txtInput(13).Text = RS02!내용연수
                End If
            End If
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

Private Sub CPrt_ErrorPrinter(ErrNum As Integer, ErrDec As String)
    MsgBox Err.Description
End Sub

Private Sub dtInput_Change(Index As Integer)
    If Index = 0 Then
        ReDim sValue(2)
        
        sValue(0) = "0"
        sValue(1) = Format(dtInput(0).Value, "yyyy")
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_06007_00", sValue(), Err_Num, Err_Dec)
        
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
        Set RS01 = ExecPro("SP_06001_03", sValue(), Err_Num, Err_Dec)
        
        cboInput(3).Clear
        
        Do While Not RS01.EOF
            cboInput(3).AddItem RS01!택번호
            
            RS01.MoveNext
        Loop
    End If
End Sub

Private Sub Form_Activate()

'    cmdBtn(0).Enabled = True
'    cmdBtn(1).Enabled = True
'    cmdBtn(2).Enabled = True
'    cmdBtn(3).Enabled = True
'    cmdBtn(4).Enabled = True
'    cmdBtn(5).Enabled = True
'
'    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    If P_06007_Flag = False Then
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
        Set RS01 = ExecPro("SP_06007_00", sValue(), Err_Num, Err_Dec)
        
        cboInput(6).Clear
        
        Do While Not RS01.EOF
            cboInput(6).AddItem Format(RS01!접수일자, "YYYY-MM-DD") & " / " & RS01!접수번호 & " / " & RS01!매장명
        
            RS01.MoveNext
        Loop
        
        P_06007_Flag = True
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_06007_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i As Integer
    
    ReDim sValue(2)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Trim(Mid(cboInput(6).Text, 13, 6))
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_06007_01", sValue(), Err_Num, Err_Dec)
    
    If RS01.EOF Then
        Exit Sub
    End If
    
    If Not IsNull(RS01!접수번호) Then txtInput(0).Text = RS01!접수번호 Else txtInput(1).Text = ""
    
    If Not IsNull(RS01!지사코드) Then
        For i = 0 To cboInput(0).ListCount - 1
            If RS01!지사코드 = Mid(cboInput(0).List(i), 2, 4) Then
                cboInput(0).ListIndex = i
                Exit For
            End If
        Next i
    Else
        cboInput(0).ListIndex = -1
    End If
    
    If Not IsNull(RS01!담당자코드) Then
        For i = 0 To cboInput(1).ListCount - 1
            If RS01!담당자코드 = Mid(cboInput(1).List(i), 2, 3) Then
                cboInput(1).ListIndex = i
                Exit For
            End If
        Next i
    Else
        cboInput(1).ListIndex = -1
    End If
    
    If Not IsNull(RS01!대리점코드) Then
        For i = 0 To cboInput(2).ListCount - 1
            If RS01!대리점코드 = Mid(cboInput(2).List(i), 2, 6) Then
                cboInput(2).ListIndex = i
                Exit For
            End If
        Next i
    Else
        cboInput(2).ListIndex = -1
    End If
    
    If Trim(RS01!입고일자) <> "" Then dtInput(1).Value = Format(RS01!입고일자, "####-##-##") Else dtInput(1).Value = ""
    
    If Trim(RS01!택번호) <> "" Then
        For i = 0 To cboInput(3).ListCount - 1
            If Format(RS01!택번호, "@-@@@") = cboInput(3).List(i) Then
                cboInput(3).ListIndex = i
            End If
            
            If i = cboInput(3).ListCount Then
                cboInput(3).Text = Format(RS01!택번호, "@-@@@")
            End If
        Next i
    Else
        cboInput(3).Text = Format(RS01!택번호, "@-@@@")
    End If
    
    If i = cboInput(3).ListCount Then
        cboInput(3).Text = Format(RS01!택번호, "@-@@@")
    End If
    
    If Trim(RS01!출고일자) <> "" Then dtInput(2).Value = Format(RS01!출고일자, "####-##-##") Else dtInput(2).Value = ""
    If Trim(RS01!인도일자) <> "" Then dtInput(3).Value = Format(RS01!인도일자, "####-##-##") Else dtInput(3).Value = ""
    If Not IsNull(RS01!품명) Then txtInput(1).Text = RS01!품명 Else txtInput(1).Text = ""
    If Not IsNull(RS01!브랜드) Then txtInput(2).Text = RS01!브랜드 Else txtInput(2).Text = ""
    If Not IsNull(RS01!색상) Then txtInput(3).Text = RS01!색상 Else txtInput(3).Text = ""
    If Trim(RS01!구입일자) <> "" Then dtInput(4).Value = Format(RS01!구입일자, "####-##-##") Else dtInput(4).Value = ""
    If Not IsNull(RS01!구입처) Then txtInput(4).Text = RS01!구입처 Else txtInput(4).Text = ""
    If Not IsNull(RS01!구입형태) Then txtInput(5).Text = RS01!구입형태 Else txtInput(5).Text = ""
    If Not IsNull(RS01!구입가격) Then mskInput(0).Text = RS01!구입가격 Else mskInput(0).Text = ""
    
    If Not IsNull(RS01!고객성명) Then txtInput(6).Text = RS01!고객성명 Else txtInput(6).Text = ""
    If Not IsNull(RS01!고객전화번호) Then txtInput(7).Text = RS01!고객전화번호 Else txtInput(7).Text = ""
    If Not IsNull(RS01!고객주소) Then txtInput(8).Text = RS01!고객주소 Else txtInput(8).Text = ""
    If Trim(RS01!처리일자) <> "" Then dtInput(5).Value = Format(RS01!처리일자, "####-##-##") Else dtInput(5).Value = ""
    
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
    
    If Not IsNull(RS01!보상금액) Then mskInput(1).Text = RS01!보상금액 Else mskInput(1).Text = ""
    If Not IsNull(RS01!비고) Then txtInput(9).Text = RS01!비고 Else txtInput(9).Text = ""
    
    If Not IsNull(RS01!대리점의견1) Then txtInput(10).Text = RS01!대리점의견1 Else txtInput(10).Text = ""
    If Not IsNull(RS01!대리점의견2) Then txtInput(11).Text = RS01!대리점의견2 Else txtInput(11).Text = ""
    If Not IsNull(RS01!대리점의견3) Then txtInput(12).Text = RS01!대리점의견3 Else txtInput(12).Text = ""
    
    If Not IsNull(RS01!핸드폰번호) Then txtInput(17).Text = RS01!핸드폰번호 Else txtInput(17).Text = ""
    
    If Not IsNull(RS01!클레임품목) Then
        For i = 0 To cboInput(7).ListCount - 1
            If RS01!클레임품목 = Mid(cboInput(7).List(i), 2, 3) Then
                cboInput(7).ListIndex = i
                Exit For
            End If
        Next i
    Else
        cboInput(7).ListIndex = -1
    End If
    
    If Not IsNull(RS01!클레임용도) Then
        For i = 0 To cboInput(8).ListCount - 1
            If RS01!클레임용도 = Mid(cboInput(8).List(i), 2, 3) Then
                cboInput(8).ListIndex = i
                Exit For
            End If
        Next i
    Else
        cboInput(8).ListIndex = -1
    End If
    
    If Not IsNull(RS01!클레임소재) Then
        For i = 0 To cboInput(9).ListCount - 1
            If RS01!클레임소재 = Mid(cboInput(9).List(i), 2, 3) Then
                cboInput(9).ListIndex = i
                Exit For
            End If
        Next i
    Else
        cboInput(9).ListIndex = -1
    End If
    
    If Not IsNull(RS01!내용연수) Then txtInput(13).Text = RS01!내용연수 Else txtInput(13).Text = ""
    If Not IsNull(RS01!경과일수) Then txtInput(14).Text = RS01!경과일수 Else txtInput(14).Text = ""
    If Not IsNull(RS01!환산일수) Then txtInput(15).Text = RS01!환산일수 Else txtInput(15).Text = ""
    If Not IsNull(RS01!배상일수) Then txtInput(16).Text = RS01!배상일수 Else txtInput(16).Text = ""
    If Not IsNull(RS01!배상금액) Then mskInput(2).Text = RS01!배상금액 Else mskInput(2).Text = ""
    
'    If Not IsNull(RS01!이미지) Then
'        pctPicture.Picture = LoadPicture(RS01!이미지)
'    End If
    
    Exit Sub
    
ErrRtn:
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
    Call AgencyComboAdd(cboInput(0))

    ReDim sValue(1)
    
    sValue(0) = "0"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_00001", sValue(), Err_Num, Err_Dec)

    Do While Not RS01.EOF
        cboInput(1).AddItem "[" & RS01!담당자코드 & "] " & RS01!담당자명
        
        RS01.MoveNext
    Loop


    Call Master_tblComboAdd(cboInput(0))


    sValue(0) = "0"
'
'    Set RS01 = New ADODB.Recordset
'    Set RS01 = ExecPro("SP_00002", sValue(), Err_Num, Err_Dec)
'
'    Do While Not RS01.EOF
'        cboInput(2).AddItem "[" & RS01!기사코드 & "] " & RS01!기사명
'
'        RS01.MoveNext
'    Loop
    
    cboInput(4).AddItem "탈색"
    cboInput(4).AddItem "파손"
    cboInput(4).AddItem "이염"
    cboInput(4).AddItem "분실"
    cboInput(4).AddItem "기타"
    
    cboInput(5).AddItem "수선"
    cboInput(5).AddItem "물품인도후 일부보상"
    cboInput(5).AddItem "현금"
    cboInput(5).AddItem "제품"
    cboInput(5).AddItem "복구"
    
    ReDim sValue(1)
    
    sValue(0) = "0"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_00008", sValue(), Err_Num, Err_Dec)

    Do While Not RS01.EOF
        cboInput(7).AddItem "[" & RS01!품목코드 & "] " & RS01!품목명
        
        RS01.MoveNext
    Loop
    

    ReDim sValue(1)
    
    sValue(0) = "0"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_00009", sValue(), Err_Num, Err_Dec)

    Do While Not RS01.EOF
        cboInput(8).AddItem "[" & RS01!용도코드 & "] " & RS01!용도명
        
        RS01.MoveNext
    Loop

    ReDim sValue(1)
    
    sValue(0) = "0"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_00010", sValue(), Err_Num, Err_Dec)

    Do While Not RS01.EOF
        cboInput(9).AddItem "[" & RS01!소재코드 & "] " & RS01!소재명
        
        RS01.MoveNext
    Loop
End Sub

Public Sub DataSave()
    If MsgBox("해당되는 내역을 저장하시겠습니까?", vbYesNo + vbInformation, "데이터 저장") = vbYes Then
        ReDim sValue(35)
        
        sValue(0) = Format(dtInput(0).Value, "YYYY-MM-DD")                        ' 접수일자
        sValue(1) = txtInput(0).Text                                            ' 접수번호
        sValue(2) = Mid(cboInput(0).Text, 2, 4)                                 ' 지사코드
        sValue(3) = Mid(cboInput(1).Text, 2, 3)                                 ' 담당자코드
        sValue(4) = Mid(cboInput(2).Text, 2, 6)                                 ' 대리점코드 6자리
        sValue(5) = Mid(cboInput(3).Text, 1, 1) & Mid(cboInput(3).Text, 3, 3)   ' 택번호
        sValue(6) = Format(dtInput(1).Value, "YYYY-MM-DD")                        ' 입고일자
        sValue(7) = Format(dtInput(2).Value, "YYYY-MM-DD")                        ' 출고일자
        sValue(8) = Format(dtInput(3).Value, "YYYY-MM-DD")                        ' 인도일자
        sValue(9) = Replace(txtInput(1).Text, "'", " ")                         ' 품목
        sValue(10) = Replace(txtInput(2).Text, "'", " ")                        ' 브랜드
        sValue(11) = Replace(txtInput(3).Text, "'", " ")                        ' 색상
        sValue(12) = Format(dtInput(4).Value, "YYYY-MM-DD")                       ' 구입일자
        sValue(13) = Replace(txtInput(4).Text, "'", " ")                        ' 구입처
        sValue(14) = Replace(txtInput(5).Text, "'", " ")                        ' 구입형태
        If mskInput(0).ClipText = "" Then
            sValue(15) = 0
        Else
            sValue(15) = mskInput(0).ClipText
        End If
        sValue(16) = Replace(txtInput(6).Text, "'", " ")                        ' 고객성명
        sValue(17) = Replace(txtInput(7).Text, "'", " ")                        ' 고객전화번호
        sValue(18) = Replace(txtInput(8).Text, "'", " ")                        ' 고객주소
        sValue(19) = Replace(cboInput(4).Text, "'", " ")                        ' 크레임구분
        sValue(20) = Replace(cboInput(5).Text, "'", " ")                        ' 보상구분
        sValue(21) = Format(dtInput(5).Value, "YYYY-MM-DD")                       ' 처리일자
        If mskInput(1).ClipText = "" Then
            sValue(22) = 0
        Else
            sValue(22) = mskInput(1).ClipText
        End If
        sValue(23) = Replace(txtInput(9).Text, "'", " ")                        ' 비고
        sValue(24) = Replace(txtInput(10).Text, "'", " ")                       ' 대리점의견1
        sValue(25) = Replace(txtInput(11).Text, "'", " ")                       ' 대리점의견2
        sValue(26) = Replace(txtInput(12).Text, "'", " ")                       ' 대리점의견3
        sValue(27) = Replace(txtInput(17).Text, "'", " ")                       ' 핸드폰번호
        sValue(28) = Mid(cboInput(7).Text, 2, 3)                                ' 품목
        sValue(29) = Mid(cboInput(8).Text, 2, 3)                                ' 용도
        sValue(30) = Mid(cboInput(9).Text, 2, 3)                                ' 소재
        
        If txtInput(13).Text = "" Then
            sValue(31) = "0"
        Else
            sValue(31) = Replace(txtInput(13).Text, "'", " ")                   ' 내용연수
        End If
        
        If txtInput(14).Text = "" Then
            sValue(32) = "0"
        Else
            sValue(32) = Replace(txtInput(14).Text, "'", " ")                   ' 경과일수
        End If
        
        If txtInput(15).Text = "" Then
            sValue(33) = "0"
        Else
            sValue(33) = Replace(txtInput(15).Text, "'", " ")                   ' 경과일수
        End If
        
        If txtInput(16).Text = "" Then
            sValue(34) = "0"
        Else
            sValue(34) = Replace(txtInput(16).Text, "'", " ")                   ' 경과일수
        End If
        
        If mskInput(2).ClipText = "" Then
            sValue(35) = 0
        Else
            sValue(35) = Replace(mskInput(2).ClipText, "'", " ")
        End If
        
        Call ExecPro("SP_06007_04", sValue(), Err_Num, Err_Dec)
    
        If Err_Num = 0 Then
            MsgBox "해당되는 데이터가 정상적으로 저장이 되었습니다.", vbInformation
        
            ReDim sValue(2)
            
            sValue(0) = "0"
            sValue(1) = Format(dtInput(0).Value, "YYYY")
            
            Set RS01 = New ADODB.Recordset
            Set RS01 = ExecPro("SP_06007_00", sValue(), Err_Num, Err_Dec)
            
            cboInput(6).Clear
            
            Do While Not RS01.EOF
                cboInput(6).AddItem Format(RS01!접수일자, "YYYY-MM-DD") & " / " & RS01!접수번호 & " / " & RS01!매장명
            
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
    Set RS01 = ExecPro("SP_06001_02", sValue(), Err_Num, Err_Dec)
    
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

Public Sub DataPrint()

    If MsgBox("해당 내용을 출력 하시겠습니까?", vbInformation + vbYesNo, "출력 확인") = vbYes Then
    
        ReDim sValue(1)
        sValue(0) = Format(dtInput(0).Value, "YYYY-MM-DD")
        sValue(1) = Trim(Mid(cboInput(6).Text, 13, 6))
        
        Set CPrt = New CCAIDPrinter
        Call CPrt.PRT_06007_VIEW(Printer, sValue(0), sValue(1))
    End If

End Sub

