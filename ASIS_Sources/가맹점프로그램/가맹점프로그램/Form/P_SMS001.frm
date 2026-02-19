VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form P_SMS001 
   ClientHeight    =   12735
   ClientLeft      =   1965
   ClientTop       =   4560
   ClientWidth     =   17460
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
   LinkTopic       =   "Form23"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12735
   ScaleWidth      =   17460
   WindowState     =   2  '최대화
   Begin Threed.SSPanel pnlSend 
      Height          =   4965
      Index           =   0
      Left            =   3780
      TabIndex        =   27
      Top             =   8835
      Visible         =   0   'False
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   8758
      _Version        =   262144
      BackColor       =   13160660
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   2
      BevelInner      =   1
      RoundedCorners  =   0   'False
      Begin VB.TextBox txtMobile 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   9
         Left            =   5130
         TabIndex        =   42
         Top             =   4410
         Width           =   2985
      End
      Begin VB.TextBox txtMobile 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   8
         Left            =   5130
         TabIndex        =   41
         Top             =   4005
         Width           =   2985
      End
      Begin VB.TextBox txtMobile 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   7
         Left            =   5130
         TabIndex        =   40
         Top             =   3600
         Width           =   2985
      End
      Begin VB.TextBox txtMobile 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   6
         Left            =   5130
         TabIndex        =   39
         Top             =   3195
         Width           =   2985
      End
      Begin VB.TextBox txtMobile 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   5130
         TabIndex        =   38
         Top             =   2790
         Width           =   2985
      End
      Begin VB.TextBox txtMobile 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   5130
         TabIndex        =   37
         Top             =   2385
         Width           =   2985
      End
      Begin VB.TextBox txtMobile 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   5130
         TabIndex        =   36
         Top             =   1980
         Width           =   2985
      End
      Begin VB.TextBox txtMobile 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   5130
         TabIndex        =   35
         Top             =   1575
         Width           =   2985
      End
      Begin VB.TextBox txtMobile 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   5130
         TabIndex        =   34
         Top             =   1170
         Width           =   2985
      End
      Begin VB.TextBox txtMobile 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   5130
         TabIndex        =   33
         Top             =   795
         Width           =   2985
      End
      Begin VB.TextBox txtRecvTel 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1830
         TabIndex        =   31
         Top             =   4410
         Width           =   2955
      End
      Begin VB.TextBox txtSend2 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3225
         Left            =   1830
         MultiLine       =   -1  'True
         TabIndex        =   28
         Top             =   780
         Width           =   2955
      End
      Begin Threed.SSCommand cmdSend 
         Height          =   915
         Left            =   8310
         TabIndex        =   44
         Top             =   450
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1614
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "문자 보내기"
         ButtonStyle     =   2
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "받는 사람"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   5130
         TabIndex        =   43
         Top             =   420
         Width           =   2985
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "보내는 사람"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   1830
         TabIndex        =   32
         Top             =   4020
         Width           =   2955
      End
      Begin VB.Label lblLan2 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4170
         TabIndex        =   30
         Top             =   420
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "전송 메시지"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2130
         TabIndex        =   29
         Top             =   480
         Width           =   1740
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  '투명하지 않음
         Height          =   375
         Left            =   1830
         Top             =   420
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "발송자취소"
      Height          =   345
      Index           =   4
      Left            =   5280
      TabIndex        =   67
      Top             =   1560
      Width           =   1215
   End
   Begin Threed.SSPanel pnlDetView 
      Height          =   4605
      Left            =   435
      TabIndex        =   53
      Top             =   8415
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   8123
      _Version        =   262144
      BackColor       =   13160660
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   2
      BevelInner      =   1
      RoundedCorners  =   0   'False
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   60
         TabIndex        =   56
         Top             =   60
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   767
         _Version        =   262144
         ForeColor       =   -2147483633
         BackColor       =   16711680
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "상 세 내 역"
         RoundedCorners  =   0   'False
      End
      Begin VB.CommandButton Command1 
         Caption         =   "닫기"
         Height          =   435
         Left            =   6570
         TabIndex        =   55
         Top             =   60
         Width           =   1425
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   4020
         Left            =   60
         TabIndex        =   54
         Top             =   510
         Width           =   7935
         _Version        =   524288
         _ExtentX        =   13996
         _ExtentY        =   7091
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowUserFormulas=   -1  'True
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
         DInformActiveRowChange=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         MaxRows         =   300
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "P_SMS001.frx":0000
         VisibleCols     =   2
         VisibleRows     =   50
         AppearanceStyle =   0
      End
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "발송일자순"
      Height          =   405
      Index           =   3
      Left            =   10530
      TabIndex        =   63
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "이름순"
      Height          =   405
      Index           =   2
      Left            =   9300
      TabIndex        =   62
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "전체취소"
      Height          =   345
      Index           =   1
      Left            =   5280
      TabIndex        =   61
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "전체선택"
      Height          =   345
      Index           =   0
      Left            =   5280
      TabIndex        =   60
      Top             =   840
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Height          =   585
      Left            =   60
      TabIndex        =   57
      Top             =   1320
      Width           =   5205
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1530
         TabIndex        =   58
         Top             =   150
         Width           =   3585
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "보내는 사람"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   60
         TabIndex        =   59
         Top             =   150
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "사용 정보"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1215
      Left            =   60
      TabIndex        =   4
      Top             =   90
      Width           =   5205
      Begin Threed.SSCommand cmdBtn 
         Height          =   435
         Index           =   2
         Left            =   4020
         TabIndex        =   66
         Top             =   720
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   767
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "설정"
         ButtonStyle     =   2
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "전송후 수량"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   4
         Left            =   2580
         TabIndex        =   11
         Top             =   330
         Width           =   1410
      End
      Begin VB.Label lblSMS 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   4020
         TabIndex        =   10
         Top             =   270
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "남은수량"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   13
         Left            =   225
         TabIndex        =   8
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "선택수량"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   3
         Left            =   150
         TabIndex        =   7
         Top             =   780
         Width           =   1020
      End
      Begin VB.Label lblSMS 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   6
         Top             =   750
         Width           =   1125
      End
      Begin VB.Label lblSMS 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1320
         TabIndex        =   5
         Top             =   270
         Width           =   1125
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1125
      Left            =   1800
      TabIndex        =   3
      Top             =   3360
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   1984
      _Version        =   262144
      ForeColor       =   16777215
      BackColor       =   16711680
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "서버에 연결중 입니다. 잠시만 기다려 주십시요..."
      BevelInner      =   1
      FloodColor      =   16777215
      RoundedCorners  =   0   'False
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   5280
      TabIndex        =   0
      Top             =   90
      Width           =   6495
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   345
         Left            =   3480
         TabIndex        =   51
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
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
         Format          =   56360961
         CurrentDate     =   39596
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1470
         TabIndex        =   50
         Top             =   240
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   661
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
         Format          =   56360961
         CurrentDate     =   39596
      End
      Begin Threed.SSCommand cmdBtn 
         Height          =   465
         Index           =   0
         Left            =   5340
         TabIndex        =   52
         Top             =   210
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   820
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "조회"
         ButtonStyle     =   2
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "접수일자"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   300
         Width           =   1020
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   4005
      Left            =   11265
      TabIndex        =   12
      Top             =   3480
      Visible         =   0   'False
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   7064
      _Version        =   262144
      BackColor       =   13160660
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
      BorderWidth     =   2
      BevelInner      =   1
      RoundedCorners  =   0   'False
      Begin VB.TextBox txtServer 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   3
         Left            =   2250
         TabIndex        =   25
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox txtServer 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   2
         Left            =   2250
         TabIndex        =   23
         Top             =   1770
         Width           =   2895
      End
      Begin VB.TextBox txtServer 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   1
         Left            =   2250
         TabIndex        =   21
         Top             =   1260
         Width           =   2895
      End
      Begin VB.TextBox txtServer 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   0
         Left            =   2250
         TabIndex        =   15
         Top             =   750
         Width           =   2895
      End
      Begin Threed.SSCommand cmdSvr 
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   2910
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   873
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "초기화"
         ButtonStyle     =   2
      End
      Begin Threed.SSCommand cmdSvr 
         Height          =   495
         Index           =   1
         Left            =   2700
         TabIndex        =   17
         Top             =   2910
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   873
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "연결 확인"
         ButtonStyle     =   2
      End
      Begin Threed.SSCommand cmdSvr 
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   3420
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   873
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "저장"
         ButtonStyle     =   2
      End
      Begin Threed.SSCommand cmdSvr 
         Height          =   495
         Index           =   3
         Left            =   2700
         TabIndex        =   19
         Top             =   3420
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   873
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "닫기"
         ButtonStyle     =   2
      End
      Begin VB.Label Label4 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "비밀번호"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   660
         TabIndex        =   24
         Top             =   2400
         Width           =   1035
      End
      Begin VB.Label Label4 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "사용자 이름"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   465
         TabIndex        =   22
         Top             =   1860
         Width           =   1425
      End
      Begin VB.Label Label4 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "서버 DB명"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   615
         TabIndex        =   20
         Top             =   1350
         Width           =   1185
      End
      Begin VB.Label Label4 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "서버 IP"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   705
         TabIndex        =   14
         Top             =   840
         Width           =   945
      End
      Begin VB.Label Label4 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "본사 서버 연결 정보 설정"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   1080
         TabIndex        =   13
         Top             =   180
         Width           =   3105
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  '투명하지 않음
         Height          =   525
         Index           =   0
         Left            =   60
         Top             =   60
         Width           =   5265
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  '투명하지 않음
         Height          =   465
         Index           =   1
         Left            =   120
         Top             =   750
         Width           =   2055
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  '투명하지 않음
         Height          =   465
         Index           =   2
         Left            =   120
         Top             =   1260
         Width           =   2055
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  '투명하지 않음
         Height          =   465
         Index           =   3
         Left            =   120
         Top             =   1770
         Width           =   2055
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  '투명하지 않음
         Height          =   465
         Index           =   4
         Left            =   120
         Top             =   2280
         Width           =   2055
      End
   End
   Begin FPSpreadADO.fpSpread SS 
      Height          =   4890
      Left            =   60
      TabIndex        =   1
      Top             =   3060
      Width           =   11745
      _Version        =   524288
      _ExtentX        =   20717
      _ExtentY        =   8625
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowUserFormulas=   -1  'True
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      DAutoCellTypes  =   0   'False
      DAutoHeadings   =   0   'False
      DAutoSave       =   0   'False
      DAutoSizeCols   =   0
      DInformActiveRowChange=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   10
      MaxRows         =   300
      MoveActiveOnFocus=   0   'False
      Protect         =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "P_SMS001.frx":060B
      VisibleCols     =   2
      VisibleRows     =   50
      AppearanceStyle =   0
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   1125
      Left            =   60
      TabIndex        =   45
      Top             =   1920
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1984
      _Version        =   262144
      BackColor       =   13160660
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   2
      BevelInner      =   1
      RoundedCorners  =   0   'False
      Begin VB.CommandButton cmdSendTextSave 
         Caption         =   "삭제"
         Height          =   495
         Index           =   2
         Left            =   11280
         TabIndex        =   70
         Top             =   570
         Width           =   405
      End
      Begin VB.CommandButton cmdSendTextSave 
         Caption         =   "추가"
         Height          =   495
         Index           =   0
         Left            =   10440
         TabIndex        =   69
         Top             =   570
         Width           =   405
      End
      Begin VB.CommandButton cmdSendTextSave 
         Caption         =   "수정"
         Height          =   495
         Index           =   1
         Left            =   10860
         TabIndex        =   68
         Top             =   570
         Width           =   405
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "변경 암호"
         Height          =   465
         Left            =   10440
         TabIndex        =   65
         Top             =   90
         Width           =   1245
      End
      Begin VB.TextBox txtSMS 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1530
         TabIndex        =   49
         Text            =   "Text1"
         Top             =   180
         Width           =   8865
      End
      Begin VB.ComboBox cboSendText 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "P_SMS001.frx":0D88
         Left            =   1530
         List            =   "P_SMS001.frx":0D8A
         Style           =   2  '드롭다운 목록
         TabIndex        =   46
         Top             =   600
         Width           =   8895
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "전송 메시지"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   6
         Left            =   90
         TabIndex        =   48
         Top             =   240
         Width           =   1410
      End
      Begin VB.Label lbl_SMS 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H80000000&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Top             =   600
         Width           =   1245
      End
   End
   Begin Threed.SSCommand cmdBtn 
      Height          =   525
      Index           =   1
      Left            =   10140
      TabIndex        =   2
      Top             =   840
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   926
      _Version        =   262144
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "보내기"
      ButtonStyle     =   2
   End
   Begin Threed.SSCommand cmdBtn 
      Height          =   525
      Index           =   4
      Left            =   8475
      TabIndex        =   26
      Top             =   840
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   926
      _Version        =   262144
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "개별 발송"
      ButtonStyle     =   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "정렬방법"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   8160
      TabIndex        =   64
      Top             =   1530
      Width           =   1020
   End
End
Attribute VB_Name = "P_SMS001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim rs01 As DAO.Recordset

Dim m_Host_DataBase     As ADODB.Connection
Dim m_Connect           As Boolean
Dim FORM_SMS001_ACTIVATE    As Boolean
Dim bCountFlag          As Boolean
Dim bSSChangeFlag       As Boolean
Dim bSSChangeFlag2      As Boolean



Private Sub cboSendText_Click()
    If cboSendText.ListIndex >= 0 Then
        txtSMS.Text = cboSendText.Text
        txtSend2.Text = cboSendText.Text
    End If
        
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        ' 조회
        Case 0
            Call DataDisplay
            Exit Sub
        ' 발송
        Case 1
            Call SendSMS
            Exit Sub
        ' 설정
        Case 2
            SSPanel2.ZOrder 0
            SSPanel2.Visible = Not SSPanel2.Visible
            Exit Sub
            
        ' 개별 문자 메시지
        Case 4
            
            If cmdBtn(Index).Caption = "개별 발송" Then
                cmdBtn(Index).Caption = "그룹 발송"
                cmdBtn(1).Visible = False
            Else
                cmdBtn(Index).Caption = "개별 발송"
                cmdBtn(1).Visible = True
            End If
            
            pnlSend(0).ZOrder 0
            pnlSend(0).Visible = Not pnlSend(0).Visible
            
        Case Else
        
    End Select
End Sub


Private Sub cmdChange_Click()
'+------------------------------------------------------
'+ 2003/02/11 수정
'+
'+루틴설명      - 비밀번호확인
'+  1. 암호를 확인하여 암호 규칙에 맞으면 화면을 종료한다.
'+  2. 레지스터리에 저장한다.
'+
'+------------------------------------------------------
    Dim strPass As String
    Dim bPass   As Boolean
    
    ' 입력 확인
    bPass = False
    
    strPass = InputBox("암호를 입력하여 주십시요", "SMS 암호")
    If Len(strPass) <= 0 Then
        Exit Sub
    End If
    
'   기본 디폴드 암호.. ( 프로그램 셋팅/설치를 위한 암호 )
    If UCase(strPass) = "DUDTJSGH" Then
        bPass = True
    Else
        ' 비밀번호 확인
        strPass = IsPassWord(strPass)
        If strPass = "-1" Or strPass = "-3" Then
            If strPass = "-3" Then MsgBox "입력한 내용이 정확하지 않습니다.", vbCritical, "입력오류"
            Exit Sub
        Else
            bPass = True
        End If
    End If
    
    ' 암호를 입력 하였으면
    If bPass = True Then
        m_SMS_EMART_PASS = True
        cmdBtn(4).Enabled = True
        txtSMS.Enabled = True
    End If

End Sub

Private Sub cmdSelect_Click(Index As Integer)
  Dim lRow    As Long
    Dim sTel(2) As String
    
    Select Case Index
    
        Case 2
        ' 이름순
            SS.Col = 1
            cmdSelect(Index).Tag = IIf(cmdSelect(Index).Tag = "1", "2", "1")
            Call Sort_Select(SS, Val(cmdSelect(Index).Tag), 1)
    
        Case 3
        ' 발송일자순
            SS.Col = 9
            cmdSelect(Index).Tag = IIf(cmdSelect(Index).Tag = "1", "2", "1")
            Call Sort_Select(SS, Val(cmdSelect(Index).Tag), 9)
        
        Case 1
        ' 전체 취소
            For lRow = 1 To SS.MaxRows
                SS.Col = 1: SS.Row = lRow
                
                If SS.Value = 1 Then
                    SS.SetText 1, lRow, "0"
                End If
            Next lRow
            
            
            lblSMS(1).Caption = CStr(GetSelectSpread(SS, 1))
            Exit Sub
            
        Case 4
        ' 발송자 취소
            For lRow = 1 To SS.MaxRows
                SS.Col = 9: SS.Row = lRow
                
                If Trim(SS.Text) <> "" Then
                    SS.SetText 1, lRow, "0"
                End If
            Next lRow
            
            
            lblSMS(1).Caption = CStr(GetSelectSpread(SS, 1))
            Exit Sub
            
        Case 0
            For lRow = 1 To SS.MaxRows
                SS.Col = 4: SS.Row = lRow
                
                ' 휴대폰 번호가 있을경우
                If SS.Text <> "" Then
                    SS.Col = 10: SS.Row = lRow
                    ' 해당 회원의 정보를 얻어온다.
                    If Fun_고객정보(SS.Text) <> "Error" Then
                
                        If CheckMobileNumber(고객정보.휴대폰, sTel) = True Then
                            If 고객정보.SMS전송여부 = "N" Then
                                SS.Col = -1: SS.Row = lRow
                                SS.BackColor = vbRed
                            Else
                                SS.SetText 1, lRow, "1"
                                ' 선택 수량 누적
                                bCountFlag = False
                                'lblSMS(1).Caption = Format(Val(Replace(lblSMS(1).Caption, ",", "")) + 1, "#,##0")
                            End If
                        End If
                    End If
                End If
            Next lRow
            lblSMS(1).Caption = CStr(GetSelectSpread(SS, 1))
            Exit Sub
    End Select
    
End Sub

Private Sub cmdSend_Click()
    Dim nIndex       As Integer
    Dim nSendCount   As Integer
    Dim sSendTel(2)  As String
    Dim sRecvTel(2)  As String
    Dim sValue(10)   As String
    Dim lRow         As Long
    
    On Error GoTo ErrRtn
    
    If Val(Replace(lblSMS(0).Caption, ",", "")) <= 0 Then
        MsgBox "사용 가능 여부및 수량을 확인 하여 주십시요.", vbInformation, "확인"
        Exit Sub
    End If
    
    ' 전송 시간 확인
    If "18:00" < Format(Time, "hh:mm") And 가맹점정보.SMS_EMART = "Y" Then
        MsgBox "이마트에서는 18:00 이후에는 문자 메시지 발송을 할 수 없습니다.", vbInformation, "확인"
        Exit Sub
    End If
    
    If Val(lblLan2.Tag) > 80 Then
        MsgBox "작성된 메시지가 80자 이상 입니다. 80자 이상은 전송할 수 없습니다.", vbCritical, "확인"
        Exit Sub
    ElseIf Val(lblLan2.Tag) <= 0 Then
        MsgBox "메시지를 확인하여 주십시요.", vbInformation, "확인"
        Exit Sub
    End If
    
    ' 입력 전화 번호 확인
    nSendCount = 0
    For nIndex = 0 To txtMobile.Count - 1
        If txtMobile(nIndex).Text <> "" Then
            If CheckMobileNumber(txtMobile(nIndex).Text, sSendTel) = False Then
                MsgBox txtMobile(nIndex).Text & " 전화 번호로는 문자 메시지를 보낼수 없습니다.", vbInformation, "확인"
                txtMobile(nIndex).SetFocus
                txtMobile(nIndex).SelStart = 0: txtMobile(nIndex).SelLength = Len(txtMobile(nIndex).Text)
                Exit Sub
            Else
                nSendCount = nSendCount + 1
            End If
        End If
    Next nIndex
    
    ' 발신자 번호 확인
    If CheckTelNumber(txtRecvTel.Text, sRecvTel) = False Then
        MsgBox txtRecvTel.Text & " 보내는 사람 전화 번호로를 확인하여 주십시요.", vbInformation, "확인"
        txtRecvTel.SetFocus
        txtRecvTel.SelStart = 0: txtRecvTel.SelLength = Len(txtRecvTel.Text)
        Exit Sub
    End If
    
    ' 잔여수량보다 선택 수량이 더 많을 경우
    If Val(Replace(lblSMS(0).Caption, ",", "")) < nSendCount Then
        MsgBox "SMS 잔여 수량보다 더 많이 선택되었습니다." & vbLf & vbLf & " 전송 수량을 조절 하여 주십시요.", vbInformation, "확인"
        Exit Sub
    End If
    
    SSPanel1.ZOrder 0
    SSPanel1.Visible = True
    SSPanel1.Caption = "메시지를 전송 중 입니다. 잠시만 기다려 주십시요."
    Screen.MousePointer = vbHourglass
    DoEvents
    nSendCount = 0
    
    For nIndex = 0 To txtMobile.Count - 1
            
        ' 입력된 전화 번호가 있을 경우
        If txtMobile(nIndex).Text <> "" Then
            ' 전화 번호 검사및 휴대폰 번호 분리
            If CheckMobileNumber(txtMobile(nIndex).Text, sSendTel) = False Then
                MsgBox txtMobile(nIndex).Text & " 전화 번호로는 문자 메시지를 보낼수 없습니다.", vbInformation, "확인"
                txtMobile(nIndex).SetFocus
                txtMobile(nIndex).SelStart = 0: txtMobile(nIndex).SelLength = Len(txtMobile(nIndex).Text)
                Exit Sub
                
            ' 발송 가능한 번호일 경우 발송 한다.
            Else
                ' 전송, 메시지타입, 수신번호, 발신번호, 메시지, 지사코드, 대리점코드, 고객코드, 고객성명, 참고5, 참고6
                sValue(0) = "1"
                sValue(1) = "0"
                sValue(2) = txtMobile(nIndex).Text
                sValue(3) = Trim(txtRecvTel.Text)
                sValue(4) = Trim(txtSend2.Text)
                sValue(5) = 가맹점정보.지사코드
                sValue(6) = 가맹점정보.택코드
                sValue(7) = " "
                sValue(8) = " "
                sValue(9) = 가맹점정보.가맹점코드
                sValue(10) = "2"
                
                Query = "EXEC PRO_SMS_SEND "
                Query = Query & "'" & sValue(0) & "', "
                Query = Query & "'" & sValue(1) & "', "
                Query = Query & "'" & sValue(2) & "', "
                Query = Query & "'" & sValue(3) & "', "
                Query = Query & "'" & sValue(4) & "', "
                Query = Query & "'" & sValue(5) & "', "
                Query = Query & "'" & sValue(6) & "', "
                Query = Query & "'" & sValue(7) & "', "
                Query = Query & "'" & sValue(8) & "', "
                Query = Query & "'" & sValue(9) & "', "
                Query = Query & "'" & sValue(10) & "' "
                
                Debug.Print Query
                
                If Dir(App.Path & "\NO_SMS.DAT", vbNormal) = "" Then
                    m_Host_DataBase.Execute Query
                    nSendCount = nSendCount + 1
                End If
            
            End If
        End If
    
    Next nIndex
    
    ' 최종 남은 수량을 설정한다.
    Call SetUseSMSCount

    SSPanel1.Visible = False
    MsgBox CStr(nSendCount) & " 건 발송 되었습니다.   ", vbInformation, "확인"
    Screen.MousePointer = vbDefault
    
    On Error GoTo 0
    Exit Sub

ErrRtn:
    Screen.MousePointer = vbDefault
    SSPanel1.Visible = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdSend_Click of Form frmSMSSend"
    
End Sub

Private Sub cmdSendTextSave_Click(Index As Integer)
    
    If 가맹점정보.SMS_EMART = "Y" And m_SMS_EMART_PASS = False Then
        MsgBox "이마트 매장에서는 수정 불가능 합니다.", vbInformation, "확인"
        Exit Sub
    End If
    
    Select Case Index
        Case 0  ' 추가
            Query = "INSERT INTO TB_문자발송문 VALUES('" & Format(cboSendText.ListCount + 1, "00") & "', '" & txtSMS.Text & "') "
            ADOCon.Execute Query
        
        Case 1  ' 수정
            Query = "UPDATE TB_문자발송문 SET "
            Query = Query & " 내용 = '" & txtSMS.Text & "' "
            Query = Query & " WHERE 순번 = '" & Format(cboSendText.ListIndex + 1, "00") & "'  "
            ADOCon.Execute Query
        
        Case 2  ' 삭제
            Query = "DELETE  FROM TB_문자발송문 "
            Query = Query & " WHERE 순번 = '" & Format(cboSendText.ListIndex + 1, "00") & "'  "
            ADOCon.Execute Query
        Case Else
    End Select
    
    Call ReadSendTextMessage(cboSendText)
End Sub

Private Sub cmdSvr_Click(Index As Integer)
    Select Case Index
        Case 3
            SSPanel2.Visible = False
            
        Case 2
            If SaveConnectData = True Then
                MsgBox "저장 완료     ", vbInformation
            Else
                MsgBox "저장 실패     ", vbCritical
            End If
            
        Case 1
            ' 저장을 먼저 처리하낟.
            Call SaveConnectData
            
            If CheckConnect = True Then
                MsgBox " 연결 완료    ", vbInformation
            End If
            
        Case 0
            Call DefaultServerSetting
            
    End Select

End Sub

 
Private Sub Command1_Click()
    pnlDetView.Visible = False
End Sub

Private Sub Form_Activate()
    On Error GoTo ErrRtn

    If FORM_SMS001_ACTIVATE = True Then Exit Sub
    FORM_SMS001_ACTIVATE = True
    
    DTPicker1.Value = DateAdd("d", -2, Date)
    DTPicker2.Value = Date
    
    Text1.Text = 가맹점정보.전화SMS
    
    DoEvents
    ' 기본 설정으로 본사에 연결한다.
    Call DefaultServerSetting
    If CheckConnect = True Then
        ' 최종 남은 수량을 설정한다.
        Call SetUseSMSCount
    End If

    On Error GoTo 0
    Exit Sub

ErrRtn:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Activate of Form P_SMS001"

End Sub

Private Sub Form_Load()
    
    SSPanel1.Visible = False
    pnlDetView.Visible = False
    
    txtRecvTel.Text = 가맹점정보.전화SMS
    Text1.Text = 가맹점정보.전화SMS
    
    Call ReadSendTextMessage(cboSendText)
    
    'TitleSet "세탁물 인도 문자"
    
    
    cmdBtn(4).Enabled = IIf(가맹점정보.SMS_EMART = "Y", False, True)
    cmdBtn(4).Enabled = IIf(가맹점정보.SMS_EMART = "N", True, m_SMS_EMART_PASS)
    txtSMS.Enabled = cmdBtn(4).Enabled
    
    cmdChange.Enabled = Not cmdBtn(4).Enabled

End Sub


Private Sub Form_Unload(Cancel As Integer)
    FORM_SMS001_ACTIVATE = False
    
End Sub

Private Sub lblSMS_Change(Index As Integer)
    ' 선택 수량이 변경될 경우 전송후 수량을 수정하여 준다.
    If Index = 1 Then
        If bCountFlag = True Then Exit Sub
        Dim nCount1  As Integer
        Dim nCount2  As Integer
        
        nCount1 = Val(Replace(lblSMS(0).Caption, ",", ""))
        nCount2 = Val(Replace(lblSMS(1).Caption, ",", ""))
        
        lblSMS(2).Caption = Format(nCount1 - nCount2, "#,##0")
        bCountFlag = True
    End If
End Sub
 
Private Sub fpSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    If bSSChangeFlag = True Then Exit Sub
    If bSSChangeFlag2 = True Then Exit Sub
    
    If Col = 1 Then
        
        Dim sTel(2) As String
        
        ' 선택을 해지 하면 선택 수량에서 -1를 해준다.
        If ButtonDown = 0 Then
            bCountFlag = False
            lblSMS(1).Caption = Val(Replace(lblSMS(1), ",", "")) - 1
            Exit Sub
     
        Else
            
            bSSChangeFlag2 = True
            SS.Row = Row:   SS.Col = 4
            
            If CheckMobileNumber(SS.Text, sTel) = True Then
                bCountFlag = False
                lblSMS(1).Caption = Val(Replace(lblSMS(1), ",", "")) + 1
            Else
                MsgBox "선택된 전화번호로는 문자 메시지를 보낼수 없습니다.", vbInformation, "확인"
                SS.Row = Row:   SS.Col = 1
                SS.Value = 0
            End If
            bSSChangeFlag2 = False
        End If
    End If
End Sub
 
Private Sub fpSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
    SS.Col = 10
    SS.Row = Row
    Call DisplayDetailData(SS.Text)
End Sub

Private Sub fpSpread1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu frmMain.PopUp
    End If
End Sub

Private Sub txtMobile_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{Tab}"
    End If
End Sub

Private Sub txtSend2_Change()
    lblLan2.Tag = CStr(LenB(StrConv(txtSend2.Text, vbFromUnicode)))
    lblLan2.Caption = lblLan2.Tag & "자"
    
    If LenB(StrConv(txtSend2.Text, vbFromUnicode)) > 80 Then
        lblLan2.BackColor = vbRed
        MsgBox "작성된 메시지가 80자 이상 입니다. 80자 이상은 전송할 수 없습니다.", vbCritical, "확인"
        Exit Sub
    Else
        lblLan2.BackColor = &HC0C0FF
        Exit Sub
    End If
End Sub

Private Sub txtSMS_Change()
    lbl_SMS.Tag = CStr(LenB(StrConv(txtSMS.Text, vbFromUnicode)))
    lbl_SMS.Caption = lbl_SMS.Tag & "자"
    Debug.Print lbl_SMS.Tag & "자"
    
    If LenB(StrConv(txtSMS.Text, vbFromUnicode)) > 80 Then
        lbl_SMS.BackColor = vbRed
        MsgBox "작성된 메시지가 80자 이상 입니다. 80자 이상은 전송할 수 없습니다.", vbCritical, "확인"
        Exit Sub
    Else
        lbl_SMS.BackColor = Me.BackColor
    End If

End Sub


Private Sub DefaultServerSetting()
    ' 기본 설정 정보가 없을 경우
    On Error GoTo ErrRtn
    
    txtServer(0).Text = "store.clean-aid.co.kr,8657"
    txtServer(1).Text = "Laundry"
    txtServer(2).Text = "sa"
    txtServer(3).Text = ""
    m_CommandTimeOut = 30
    
'
'    Query = "SELECT * FROM TB_기본정보 "
'    Set RS01 = MyDB.OpenRecordset(Query)
'
'    If RS01.RecordCount > 0 Then
'        txtServer(0).Text = Trim(RS01.Fields("ServerIP") & "")
'        txtServer(1).Text = Trim(RS01.Fields("ServerDB") & "")
'        txtServer(2).Text = Trim(RS01.Fields("ServerUser") & "")
'        txtServer(3).Text = Trim(RS01.Fields("ServerPass") & "")
'        m_CommandTimeOut = Val(Trim(RS01.Fields("TimeOut") & ""))
'    Else
'        txtServer(0).Text = "store.clean-aid.co.kr,8657"
'        txtServer(1).Text = "Laundry"
'        txtServer(2).Text = "sa"
'        txtServer(3).Text = ""
'        m_CommandTimeOut = 30
'    End If
    

    On Error GoTo 0
    Exit Sub

ErrRtn:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DefaultServerSetting of Form P_SMS001"
End Sub


'-------------------------------------------------------------------------------------
' 내용 구분의 정보를 저장 한다.
'-------------------------------------------------------------------------------------
Private Function SaveConnectData() As Boolean
    
    On Error GoTo ErrRtn
    Dim Query    As String
    
    SaveConnectData = False
    
    txtServer(0).Text = Trim(txtServer(0).Text)
    txtServer(1).Text = Trim(txtServer(1).Text)
    txtServer(2).Text = Trim(txtServer(2).Text)
    txtServer(3).Text = Trim(txtServer(3).Text)
    
    Query = "UPDATE TB_기본정보 SET "
    Query = Query & " ServerIP = ' " & txtServer(0).Text & "', "
    Query = Query & " ServerDB = ' " & txtServer(1).Text & "', "
    Query = Query & " ServerUser = ' " & txtServer(2).Text & "', "
    Query = Query & " ServerPass = ' " & txtServer(3).Text & "' "
    ADOCon.Execute Query
    
    SaveConnectData = True
    
    On Error GoTo 0
    Exit Function

ErrRtn:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SaveConnectData of Form P_SMS001"
End Function


Private Function CheckConnect() As Boolean
    On Error GoTo ErrRtn
    
    Dim HostConn    As String
    
    HostConn = ""
    HostConn = HostConn & "Provider=SQLOLEDB.1;"
    HostConn = HostConn & "Persist Security Info=False;"
    HostConn = HostConn & "User ID=" & txtServer(2).Text & ";"
    HostConn = HostConn & "Password=" & txtServer(3).Text & ";"
    HostConn = HostConn & "Initial Catalog=" & txtServer(1).Text & ";"
    HostConn = HostConn & "Data Source=" & txtServer(0).Text
    m_CommandTimeOut = IIf(m_CommandTimeOut = 0, 30, m_CommandTimeOut)

    Set m_Host_DataBase = Nothing
    Set m_Host_DataBase = New ADODB.Connection
    
    SSPanel1.ZOrder 0
    SSPanel1.Visible = True
    SSPanel1.Caption = "서버에 연결중 입니다. 잠시만 기다려 주십시요..."
    
    If m_Host_DataBase.State = adStateOpen Then m_Host_DataBase.Close
    
    m_Host_DataBase.ConnectionTimeout = 10
    m_Host_DataBase.CommandTimeout = m_CommandTimeOut
    m_Host_DataBase.Open HostConn
    
    SSPanel1.ZOrder 0
    SSPanel1.Visible = False
    
    m_Connect = True
    CheckConnect = True
    
    On Error GoTo 0
    Exit Function

ErrRtn:
    SSPanel1.ZOrder 0:  SSPanel1.Visible = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CheckConnect of Form P_SMS001"
End Function



'--------------------------------------------------------------------------------------------------------------
' Procedure : SetUseSMSCount
' DateTime  : 2007-06-10 12:26
' Author    : pds2004
' Purpose   : 문자 메시지의 잔여 수량을 가저온다.
'--------------------------------------------------------------------------------------------------------------
Private Sub SetUseSMSCount()
    Dim bResult     As Boolean
    Dim Query        As String
    Dim ADORset     As New ADODB.Recordset
    
    ' 연결되어 있지 않을 경우 다시한번 연결을 시도한다.
    On Error GoTo ErrRtn

    lblSMS(0).Caption = "0"
    lblSMS(1).Caption = "0"
    lblSMS(2).Caption = "0"


    If m_Connect = False Then
        If CheckConnect = False Then
            MsgBox "본사와의 연결 설정을 확인하여 주십시요", vbInformation, "확인"
            ' 설정 화면을 활성화 한다.
            Call cmdBtn_Click(2)
            Exit Sub
        End If
    End If
    
    Query = "EXEC PRO_SMS_STORE_001_01 '0', '" & 가맹점정보.가맹점코드 & "' "

    ADORset.CursorLocation = adUseClient
    ADORset.Open Query, m_Host_DataBase, adOpenStatic, adLockBatchOptimistic, adCmdText
    
    If ADORset.EOF = False Then
        If ADORset.RecordCount > 0 Then
            lblSMS(0).Caption = Format(ADORset.Fields("잔여수량") & "", "#,##0")
        End If
    End If
    ADORset.Close:  Set ADORset = Nothing

    On Error GoTo 0
    Exit Sub

ErrRtn:
    Set ADORset = Nothing
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SetUseSMSCount of Form P_SMS001"
    
End Sub


Private Sub DataDisplay()

    Dim FindCount   As Integer
    Dim sTel(2) As String
    Dim lRow        As Long
    Dim bResult     As Boolean
    Dim sData(1)    As String

    On Error GoTo ErrRtn
    
    Screen.MousePointer = vbHourglass
    
    SS.MaxRows = 0
    SS.MaxCols = 10
    
    SS.ColWidth(10) = 0
    lblSMS(1).Caption = "0"
    bSSChangeFlag = True

    
    '--------------------------------------------------------------------------------
    ' 해당 일자에 입고된 고객 번호 내역을 구한다.
    '--------------------------------------------------------------------------------
    Query = " SELECT 고객코드  "
    Query = Query & "  FROM TB_입출고 "
    Query = Query & " WHERE 접수일자 >= '" & Format(DTPicker1.Value, "YYYY-MM-DD") & "' "
    Query = Query & "   AND 접수일자 <= '" & Format(DTPicker2.Value, "YYYY-MM-DD") & "' "
    Query = Query & "   AND 본출 = '出' "
    Query = Query & "   AND 출고일자 is null "
    Query = Query & "   AND 판매취소 <>  'Y' "
    Query = Query & " GROUP BY 고객코드  "
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    lRow = 0
    If ADORs.RecordCount > 0 Then
        ADORs.MoveFirst
        'SS.ReDraw = False
        While Not ADORs.EOF
        
            ' 위에서 조회된 내용을 기준으로 해당 일자에 미 입고된 내용이 있는지 확인한다.
            ' 없을 경우에만 문자 메시지를 전송 처리하기 위하여 리스트에 표시한다.
            
            Query = " SELECT * FROM TB_입출고 "
            Query = Query & " WHERE 접수일자 >= '" & Format(DTPicker1.Value, "YYYY-MM-DD") & "' "
            Query = Query & "   AND 접수일자 <= '" & Format(DTPicker2.Value, "YYYY-MM-DD") & "' "
            Query = Query & "   AND NOT ( 본출 = '出' OR 본출 = '反') "
            Query = Query & "   AND 고객코드 = '" & ADORs.Fields("고객코드") & "" & "' "
            Query = Query & "   AND 판매취소 <>  'Y' "
            Set SUBRs = New ADODB.Recordset
            SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
            
            'Set TempRSet2 = MyDB.OpenRecordset(Query)
            
            ' 미입고된 내용이 없을 경우
            If SUBRs.EOF Then
                SUBRs.Close
                Set SUBRs = Nothing
                
                ' 해당 회원의 정보를 얻어온다.
                If Fun_고객정보(ADORs.Fields("고객코드") & "") <> "Error" Then
                
                    If SS.MaxRows = lRow Then
                        SS.MaxRows = SS.MaxRows + 1
                        SS.RowHeight(SS.MaxRows) = 20
                    End If
                    
                    If SS.MaxRows > 15 Then SS.ReDraw = False

                                
                    lRow = lRow + 1
                    SS.SetText 2, lRow, 고객정보.성명
                    SS.SetText 3, lRow, 고객정보.전화번호
                    SS.SetText 4, lRow, 고객정보.휴대폰
                                        
                    '-----------------------------------------------------------------------------------
                    ' 미출고된 내역의 정보를 구한다. 시작텍과 종료텍의 수량을 구한다.
                    '-----------------------------------------------------------------------------------
                    Query = " SELECT Min(택번호) As 시작번호, Max(택번호) AS 종료번호, Count(택번호) AS 수량,"
                    Query = Query & " MAX(본출일자) AS 본사출고일 FROM TB_입출고 "
                    Query = Query & " WHERE 접수일자 >= '" & Format(DTPicker1.Value, "YYYY-MM-DD") & "' "
                    Query = Query & "   AND 접수일자 <= '" & Format(DTPicker2.Value, "YYYY-MM-DD") & "' "
                    Query = Query & "   AND (본출 = '出'  OR 본출 = '反') "     ' 본사 출고 상태인것
                    Query = Query & "   AND 확인 <> '확' "    ' 미출고 상태인것
                    Query = Query & "   AND 고객코드 = '" & ADORs.Fields("고객코드") & "" & "' "
                    Query = Query & "   AND 판매취소 <>  'Y' "
                    Set SUBRs = New ADODB.Recordset
                    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
                    
                    'Set TempRSet3 = MyDB.OpenRecordset(Query)
                    
                    If SUBRs.EOF = False Then
                        SS.SetText 5, lRow, Trim(SUBRs.Fields("시작번호") & "")
                        SS.SetText 6, lRow, Trim(SUBRs.Fields("종료번호") & "")
                        SS.SetText 7, lRow, CStr(SUBRs.Fields("수량") & "")
                        
                        If Trim(고객정보.휴대폰) = "" Then
                            SS.SetText 9, lRow, ""
                        Else
                            SS.SetText 9, lRow, GetLastUseDate(고객정보.휴대폰) '-- 시간이 너무 걸리네 ㅡㅡ
                        End If
                        
                        fpSpread1.SetText 10, lRow, 고객정보.고객코드
                    End If
                    SUBRs.Close
                    Set SUBRs = Nothing
                    
                    '-----------------------------------------------------------------------------------
                    ' 반품수량
                    '-----------------------------------------------------------------------------------
                    Query = " SELECT Count(택번호) AS 수량,"
                    Query = Query & " MAX(본출일자) AS 본사출고일 FROM TB_입출고 "
                    Query = Query & " WHERE 접수일자 >= '" & Format(DTPicker1.Value, "YYYY-MM-DD") & "' "
                    Query = Query & "   AND 접수일자 <= '" & Format(DTPicker2.Value, "YYYY-MM-DD") & "' "
                    Query = Query & "   AND 본출 = '反' "     ' 본사 출고 상태인것
                    Query = Query & "   AND 확인 <> '확' "    ' 미출고 상태인것
                    Query = Query & "   AND 고객코드 = '" & ADORs.Fields("고객코드") & "" & "' "
                    Query = Query & "   AND 판매취소 <>  'Y' "
                    Set SUBRs = New ADODB.Recordset
                    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
                    
                    If SUBRs.EOF = False Then
                        fpSpread1.SetText 8, lRow, CStr(SUBRs.Fields("수량") & "")
                    End If
                    
                    If CheckMobileNumber(고객정보.휴대폰, sTel) = True Then
                        If 고객정보.SMS전송여부 = "N" Then
                            fpSpread1.Col = -1: fpSpread1.Row = lRow
                            fpSpread1.BackColor = vbRed
                        Else
                            fpSpread1.SetText 1, lRow, "1"
                            ' 선택 수량 누적
                            bCountFlag = False
                            lblSMS(1).Caption = Format(Val(Replace(lblSMS(1).Caption, ",", "")) + 1, "#,##0")
                        End If
                    End If
                    SUBRs.Close
                    Set SUBRs = Nothing
                End If
            Else
                SUBRs.Close
                Set SUBRs = Nothing
            End If
            
            ADORs.MoveNext
        Wend
    End If
    ADORs.Close
    Set ADORs = Nothing
    
    fpSpread1.ReDraw = True
    bSSChangeFlag = False
    Screen.MousePointer = vbDefault
    
    On Error GoTo 0
    Exit Sub

ErrRtn:
    fpSpread1.ReDraw = True
    bSSChangeFlag = False
    Screen.MousePointer = vbDefault
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DataDisplay of Form P_SMS001"

End Sub



'--------------------------------------------------------------------------------------------------------------
' Procedure : SendSMS
' DateTime  : 2007-05-06 23:16
' Author    : pds2004
' Purpose   : SMS 문자 메시지 전송
'--------------------------------------------------------------------------------------------------------------
Private Function SendSMS() As Boolean
    Dim sFlag   As Boolean
    Dim lRow    As Long
    Dim dLng    As Long
    Dim sValue(10)   As String
    Dim vTemp   As Variant
    Dim bResult     As Boolean
    Dim Query    As String
    Dim sRecvTel(2)     As String
    Dim ADORset     As New ADODB.Recordset
    
    
    On Error GoTo ErrRtn
    
    ' 문자 메시지 길이 확인
    dLng = CheckSendMessageLangth
    If dLng <= 0 Or dLng > 80 Then Exit Function
    
    ' 전송 시간 확인
    If "18:00" < Format(Time, "hh:mm") And 가맹점정보.SMS_EMART = "Y" Then
        MsgBox "이마트에서는 18:00 이후에는 문자 메시지 발송을 할 수 없습니다.", vbInformation, "확인"
        Exit Function
    End If
    
    ' 잔여 수량 확인
    If Val(Replace(lblSMS(2).Caption, ",", "")) < 0 Then
        MsgBox "잔여 수량보다 선택수량이 더 큼니다. 선택 수량을 조절하여 주십시요.", vbInformation, "확인"
        Exit Function
    End If

    If m_Connect = False Then
        If CheckConnect = False Then
            MsgBox "본사와의 연결 설정을 확인하여 주십시요", vbInformation, "확인"
            ' 설정 화면을 활성화 한다.
            Call cmdBtn_Click(2)
            Exit Function
        End If
    End If
    
    ' 최종 확인 메시지
    If MsgBox("본 출고 자료는 정확 하지 않을수 있습니다!!!!" & vbNewLine & vbNewLine & "가맹점에서 반드시 확인후 문자 메시지를 발송바랍니다." & vbNewLine & vbNewLine & "메시지를 전송 하시겠습니까? ", vbCritical + vbYesNo + vbDefaultButton2, "확인") = vbNo Then
        Exit Function
    End If
    
    ' 발신자 번호 확인
    If CheckTelNumber(Text1.Text, sRecvTel) = False Then
        MsgBox Text1.Text & " 보내는 사람 전화 번호로를 확인하여 주십시요.", vbInformation, "확인"
        Text1.SetFocus
        Text1.SelStart = 0: Text1.SelLength = Len(Text1.Text)
        Exit Function
    End If
    
    
    For lRow = 1 To fpSpread1.MaxRows
        Call fpSpread1.GetText(1, lRow, vTemp)
        sFlag = IIf(CStr(vTemp) = "1", True, False)
        
        ' 전송 구분일 경우 전송 처리한다.
        If sFlag = True Then
            ' 전송, 메시지타입, 수신번호, 발신번호, 메시지, 지사코드, 대리점코드, 고객코드, 고객성명, 참고5, 참고6
            sValue(0) = "1"
            sValue(1) = "0"
            Call fpSpread1.GetText(4, lRow, vTemp):    sValue(2) = CStr(vTemp)
            sValue(3) = Trim(Text1.Text)
            sValue(4) = Trim(txtSMS.Text)
            sValue(5) = 가맹점정보.지사코드
            sValue(6) = 가맹점정보.택코드
            
            Call fpSpread1.GetText(10, lRow, vTemp):    sValue(7) = CStr(vTemp)
            Call fpSpread1.GetText(2, lRow, vTemp):    sValue(8) = CStr(vTemp)
            
            sValue(9) = 가맹점정보.가맹점코드
            sValue(10) = "1"
            
            
            Query = "EXEC PRO_SMS_SEND "
            Query = Query & "'" & sValue(0) & "', "
            Query = Query & "'" & sValue(1) & "', "
            Query = Query & "'" & sValue(2) & "', "
            Query = Query & "'" & sValue(3) & "', "
            Query = Query & "'" & sValue(4) & "', "
            Query = Query & "'" & sValue(5) & "', "
            Query = Query & "'" & sValue(6) & "', "
            Query = Query & "'" & sValue(7) & "', "
            Query = Query & "'" & sValue(8) & "', "
            Query = Query & "'" & sValue(9) & "', "
            Query = Query & "'" & sValue(10) & "' "
            
            Debug.Print Query
            If Dir(App.Path & "\NO_SMS.DAT", vbNormal) = "" Then
                m_Host_DataBase.Execute Query
            
                vTemp = "0"
                Call fpSpread1.SetText(1, lRow, vTemp)
            
            End If
            
            
        End If
    
    Next lRow
    
    ' 최종 남은 수량을 설정한다.
    Call SetUseSMSCount
    
    Set ADORset = Nothing
    
    On Error GoTo 0
    Exit Function

ErrRtn:
    ' 최종 남은 수량을 설정한다.
    Call SetUseSMSCount
    DoEvents
    
    Set ADORset = Nothing
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SendSMS of Form V_SMS001"
End Function


Private Function CheckSendMessageLangth() As Integer
    If IsNumeric(lbl_SMS.Tag) = False Then
        CheckSendMessageLangth = 0
        txtSMS.SetFocus
        MsgBox "전송할 메시지를 입력 하여 주십시요..  [" & CStr(Val(lbl_SMS.Tag)) & "자]", vbInformation, "확인"
        Exit Function
        
    ElseIf Val(lbl_SMS.Tag) >= 80 Then
        CheckSendMessageLangth = Val(lbl_SMS.Tag)
        txtSMS.SetFocus
        MsgBox "전송할 메시지를 확인 하여 주십시요..  [" & CStr(Val(lbl_SMS.Tag)) & "자]", vbInformation, "확인"
        Exit Function
    Else
        CheckSendMessageLangth = Val(lbl_SMS.Tag)
        Exit Function
    End If
End Function



'Private Function ReadSendTextMessage() As Boolean
'    Dim TempRSet    As Recordset
'    Dim Query    As String
'
'
'    cboSendText.Clear
'
'    Query = " SELECT * "
'    Query = Query & "  FROM TB_문자발송문 "
'    Query = Query & " ORDER BY 순번 "
'    Set TempRSet = MyDB.OpenRecordset(Query)
'
'    If TempRSet.EOF Then
'        cboSendText.AddItem "전할 메시지를 입력하여 주십시요"
'    End If
'
'
'    While Not TempRSet.EOF
'        cboSendText.AddItem Trim(TempRSet.Fields("내용") & "")
'        TempRSet.MoveNext
'    Wend
'
'    ' 마지막 문구 선택
'    If cboSendText.ListCount >= 0 Then
'        cboSendText.ListIndex = cboSendText.ListCount - 1
'
'    End If
'
'    TempRSet.Close
'End Function


Private Sub DisplayDetailData(ByVal sCustCode As String)
    Dim lRow    As Long
    
    fpSpread1.MaxRows = 0
    lRow = 0
    
    Query = " SELECT * "
    Query = Query & "  FROM TB_입출고 "
    Query = Query & " WHERE 접수일자 >= '" & Format(DTPicker1.Value, "YYYY-MM-DD") & "' "
    Query = Query & "   AND 접수일자 <= '" & Format(DTPicker2.Value, "YYYY-MM-DD") & "' "
    Query = Query & "   AND (본출 = '出'  OR 본출 = '反') "     ' 본사 출고 상태인것
    Query = Query & "   AND 확인 <> '확' "    ' 미출고 상태인것
    'Query = Query & "   AND 출고일자 is null "
    Query = Query & "   AND 판매취소 <>  'Y' "
    Query = Query & "   AND 고객코드 = '" & sCustCode & "'"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic

    While Not ADORs.EOF
        If fpSpread1.MaxRows = lRow Then
            fpSpread1.MaxRows = fpSpread1.MaxRows + 1
            fpSpread1.RowHeight(fpSpread1.MaxRows) = 20
        End If

        With fpSpread1
            .Row = lRow + 1
            .Col = 1:   .Text = Format(ADORs.Fields("접수일자") & "", "YYYY-MM-DD")
            .Col = 2:   .Text = ADORs.Fields("의류명") & ""
            .Col = 3:   .Text = ADORs.Fields("택번호") & ""
            .Col = 4:   .Text = Format(ADORs.Fields("본출일자") & "", "YYYY-MM-DD")
            .Col = 5:   .Text = ADORs.Fields("본출") & ""
            
        End With
        
        lRow = lRow + 1
        
        ADORs.MoveNext
    Wend
    
    ADORs.Close
    Set ADORs = Nothing
    
    pnlDetView.Visible = True
End Sub


'--------------------------------------------------------------------------------------------------------------
' Procedure : GetLastUseDate
' DateTime  : 2008-06-7
' Author    : pds2004
' Purpose   : 최종 사용일자를 리턴한다.
'--------------------------------------------------------------------------------------------------------------
Private Function GetLastUseDate(ByVal CustPhone As String) As String
    Dim bResult     As Boolean
    Dim Query        As String
    Dim ADORset     As New ADODB.Recordset
    
    ' 연결되어 있지 않을 경우 다시한번 연결을 시도한다.
    On Error GoTo ErrRtn
 

    If m_Connect = False Then
        If CheckConnect = False Then
            MsgBox "본사와의 연결 설정을 확인하여 주십시요", vbInformation, "확인"
            ' 설정 화면을 활성화 한다.
            Call cmdBtn_Click(2)
            Exit Function
        End If
    End If
    
    Query = "EXEC PRO_SMS_STORE_001_05   '" & 가맹점정보.가맹점코드 & "', '" & CustPhone & "' "

    ADORset.CursorLocation = adUseClient
    ADORset.Open Query, m_Host_DataBase, adOpenStatic, adLockBatchOptimistic, adCmdText
    If ADORset.EOF = False Then
        If ADORset.RecordCount > 0 Then
            GetLastUseDate = ADORset.Fields("전송정보") & ""
        End If
    End If
    ADORset.Close:  Set ADORset = Nothing

    On Error GoTo 0
    Exit Function

ErrRtn:
    Set ADORset = Nothing
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetLastUseDate of Form P_SMS001"
    
End Function

Private Sub txtSMS_KeyPress(KeyAscii As Integer)
    If 가맹점정보.SMS_EMART = "Y" And m_SMS_EMART_PASS = False Then
        KeyAscii = 0
    End If

End Sub
