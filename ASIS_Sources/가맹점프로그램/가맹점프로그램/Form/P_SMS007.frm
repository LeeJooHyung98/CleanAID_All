VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form P_SMS007 
   ClientHeight    =   7965
   ClientLeft      =   930
   ClientTop       =   5940
   ClientWidth     =   11850
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
   ScaleHeight     =   7965
   ScaleWidth      =   11850
   WindowState     =   2  '최대화
   Begin Threed.SSPanel pnlSend 
      Height          =   5295
      Index           =   0
      Left            =   30
      TabIndex        =   40
      Top             =   2640
      Visible         =   0   'False
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9340
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
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
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
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   44
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
         TabIndex        =   41
         Top             =   780
         Width           =   2955
      End
      Begin Threed.SSCommand cmdSend 
         Height          =   915
         Left            =   8310
         TabIndex        =   57
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
         Caption         =   "문자 보네기"
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
         TabIndex        =   56
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
         TabIndex        =   45
         Top             =   4020
         Width           =   2955
      End
      Begin VB.Label lblLan2 
         BackColor       =   &H00FFC0FF&
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
         TabIndex        =   43
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
         TabIndex        =   42
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
   Begin Threed.SSPanel SSPanel4 
      Height          =   1125
      Left            =   30
      TabIndex        =   58
      Top             =   1530
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
         TabIndex        =   66
         Top             =   570
         Width           =   405
      End
      Begin VB.CommandButton cmdSendTextSave 
         Caption         =   "수정"
         Height          =   495
         Index           =   1
         Left            =   10860
         TabIndex        =   64
         Top             =   570
         Width           =   405
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
         TabIndex        =   63
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
         Left            =   1530
         Style           =   2  '드롭다운 목록
         TabIndex        =   60
         Top             =   600
         Width           =   8865
      End
      Begin VB.CommandButton cmdSendTextSave 
         Caption         =   "추가"
         Height          =   495
         Index           =   0
         Left            =   10440
         TabIndex        =   59
         Top             =   570
         Width           =   405
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
         TabIndex        =   62
         Top             =   240
         Width           =   1410
      End
      Begin VB.Label lbl_SMS 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '투명
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
         Left            =   10440
         TabIndex        =   61
         Top             =   180
         Width           =   1245
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   4005
      Left            =   3570
      TabIndex        =   20
      Top             =   2490
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
         TabIndex        =   33
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
         TabIndex        =   31
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
         TabIndex        =   29
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
         TabIndex        =   23
         Top             =   750
         Width           =   2895
      End
      Begin Threed.SSCommand cmdSvr 
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   24
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
         TabIndex        =   25
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
         TabIndex        =   26
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
         TabIndex        =   27
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
         TabIndex        =   32
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
         TabIndex        =   30
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
         TabIndex        =   28
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
         TabIndex        =   22
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
         TabIndex        =   21
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
      TabIndex        =   11
      Top             =   90
      Width           =   5205
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   270
         Width           =   1125
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1125
      Left            =   1800
      TabIndex        =   10
      Top             =   3150
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
      Left            =   5430
      TabIndex        =   0
      Top             =   60
      Width           =   4725
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Index           =   0
         Left            =   1470
         TabIndex        =   1
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   375
         Index           =   0
         Left            =   2550
         TabIndex        =   2
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   375
         Index           =   0
         Left            =   3270
         TabIndex        =   3
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##"
         PromptChar      =   "_"
      End
      Begin Threed.SSCommand cmdBtn 
         Height          =   405
         Index           =   3
         Left            =   4080
         TabIndex        =   38
         Top             =   210
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
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
         Caption         =   "?"
         ButtonStyle     =   2
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검색 일자"
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
         TabIndex        =   16
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  '단일 고정
         Caption         =   "일"
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
         Left            =   3630
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  '단일 고정
         Caption         =   "월"
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
         Left            =   2910
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  '단일 고정
         Caption         =   "년"
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
         Left            =   2190
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
   End
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   5100
      Left            =   60
      TabIndex        =   7
      Top             =   2700
      Width           =   11745
      _Version        =   524288
      _ExtentX        =   20717
      _ExtentY        =   8996
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
      MaxCols         =   7
      MaxRows         =   300
      MoveActiveOnFocus=   0   'False
      Protect         =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "P_SMS007.frx":0000
      VisibleCols     =   2
      VisibleRows     =   50
      AppearanceStyle =   0
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   2355
      Left            =   1470
      TabIndex        =   34
      Top             =   810
      Visible         =   0   'False
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   4154
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "본사 입고 이전에 출고처리된 고객 제외"
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
         Left            =   330
         TabIndex        =   65
         Top             =   1920
         Width           =   4755
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "미 입고된 품목이 한개라도 있을 경우 조회되지 않음"
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
         Left            =   330
         TabIndex        =   37
         Top             =   1500
         Width           =   6300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "이전 품목이 모두 입고된 고객만 조회됨"
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
         Left            =   330
         TabIndex        =   36
         Top             =   1020
         Width           =   4755
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "검색일자를 기준으로 본사에서 입고된 고객중 당일 접수분을 제외한"
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
         Left            =   330
         TabIndex        =   35
         Top             =   540
         Width           =   8085
      End
   End
   Begin Threed.SSCommand cmdBtn 
      Height          =   615
      Index           =   0
      Left            =   10230
      TabIndex        =   8
      Top             =   180
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1085
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
   Begin Threed.SSCommand cmdBtn 
      Height          =   615
      Index           =   1
      Left            =   9600
      TabIndex        =   9
      Top             =   870
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   1085
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
      Height          =   615
      Index           =   2
      Left            =   5430
      TabIndex        =   19
      Top             =   870
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   1085
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
   Begin Threed.SSCommand cmdBtn 
      Height          =   615
      Index           =   4
      Left            =   7515
      TabIndex        =   39
      Top             =   870
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   1085
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
End
Attribute VB_Name = "P_SMS007"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
            
        ' 검색 조건보기
        Case 3
            SSPanel3.ZOrder 0
            SSPanel3.Visible = Not SSPanel3.Visible
            
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


Private Sub cmdSend_Click()
    Dim nIndex  As Integer
    Dim nSendCount  As Integer
    Dim sSendTel(2)     As String
    Dim sRecvTel(2)     As String
    Dim sValue(8) As String
    Dim lRow        As Long
    
    On Error GoTo ErrRtn
    
    If Val(Replace(lblSMS(0).Caption, ",", "")) <= 0 Then
        MsgBox "사용 가능 여부및 수량을 확인 하여 주십시요.", vbInformation, "확인"
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
                MsgBox txtMobile(nIndex).Text & " 전화 번호로는 문자 메시지를 보넬수 없습니다.", vbInformation, "확인"
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
                MsgBox txtMobile(nIndex).Text & " 전화 번호로는 문자 메시지를 보넬수 없습니다.", vbInformation, "확인"
                txtMobile(nIndex).SetFocus
                txtMobile(nIndex).SelStart = 0: txtMobile(nIndex).SelLength = Len(txtMobile(nIndex).Text)
                Exit Sub
                
            ' 발송 가능한 번호일 경우 발송 한다.
            Else
                ' 전송, 메시지타입, 수신번호, 발신번호, 메시지, 지사코드, 대리점코드, 고객코드, 고객성명, 참고5, 참고6
                sValue(0) = "1"
                sValue(1) = "0"
                sValue(2) = txtMobile(nIndex).Text
                sValue(3) = txtRecvTel.Text
                sValue(4) = Trim(txtSend2.Text)
                sValue(5) = 가맹점정보.지사코드
                sValue(6) = 가맹점정보.택코드
                sValue(7) = " "
                sValue(8) = " "
                
                Query = "EXEC PRO_SMS_SEND "
                Query = Query & "'" & sValue(0) & "', "
                Query = Query & "'" & sValue(1) & "', "
                Query = Query & "'" & sValue(2) & "', "
                Query = Query & "'" & sValue(3) & "', "
                Query = Query & "'" & sValue(4) & "', "
                Query = Query & "'" & sValue(5) & "', "
                Query = Query & "'" & sValue(6) & "', "
                Query = Query & "'" & sValue(7) & "', "
                Query = Query & "'" & sValue(8) & "' "
                
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
            Query = "DELETE  FROM 문자발송문 "
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

 
Private Sub Form_Activate()
    On Error GoTo ErrRtn

    If FORM_SMS001_ACTIVATE = True Then Exit Sub
    FORM_SMS001_ACTIVATE = True
    
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

    MaskEdBox1(0).Text = Format(Date, "yyyy")
    MaskEdBox2(0).Text = Format(Date, "mm")
    MaskEdBox3(0).Text = Format(Date, "dd")
    
    
    txtRecvTel.Text = 가맹점정보.전화번호
    
    Call ReadSendTextMessage(cboSendText)
    
    'TitleSet "문자 메시지 전송"
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

Private Sub MaskEdBox1_GotFocus(Index As Integer)
    MaskEdBox1(Index).SelStart = 0
    MaskEdBox1(Index).SelLength = Len(MaskEdBox1(Index).Text)
End Sub

Private Sub MaskEdBox2_GotFocus(Index As Integer)
    MaskEdBox2(Index).SelStart = 0
    MaskEdBox2(Index).SelLength = Len(MaskEdBox2(Index).Text)
End Sub

Private Sub MaskEdBox3_GotFocus(Index As Integer)
    MaskEdBox3(Index).SelStart = 0
    MaskEdBox3(Index).SelLength = Len(MaskEdBox3(Index).Text)
End Sub

Private Sub SS_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
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
            fpSpread1.Row = Row:   fpSpread1.Col = 4
            
            If CheckMobileNumber(fpSpread1.Text, sTel) = True Then
                bCountFlag = False
                lblSMS(1).Caption = Val(Replace(lblSMS(1), ",", "")) + 1
            Else
                MsgBox "선택된 전화번호로는 문자 메시지를 보넬수 없습니다.", vbInformation, "확인"
                fpSpread1.Row = Row:   fpSpread1.Col = 1
                fpSpread1.Value = 0
            End If
            bSSChangeFlag2 = False
        End If
    End If
End Sub
 
Private Sub txtSend2_Change()
    lblLan2.Tag = CStr(LenB(StrConv(txtSend2.Text, vbFromUnicode)))
    lblLan2.Caption = lblLan2.Tag & "자"
End Sub

Private Sub txtSMS_Change()
    lbl_SMS.Tag = CStr(LenB(StrConv(txtSMS.Text, vbFromUnicode)))
    lbl_SMS.Caption = lbl_SMS.Tag & "자"
    Debug.Print lbl_SMS.Tag & "자"

End Sub

Private Sub DefaultServerSetting()
    ' 기본 설정 정보가 없을 경우
    On Error GoTo ErrRtn
        
    Query = "SELECT * FROM TB_기본정보 "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    If SUBRs.EOF Then
        txtServer(0).Text = "store.clean-aid.co.kr,8657"
        txtServer(1).Text = "Laundry"
        txtServer(2).Text = "sa"
        txtServer(3).Text = ""
        m_CommandTimeOut = 30
    Else
        txtServer(0).Text = Trim(SUBRs!ServerIP) & ""
        txtServer(1).Text = Trim(SUBRs!ServerDB) & ""
        txtServer(2).Text = Trim(SUBRs!ServerUser) & ""
        txtServer(3).Text = Trim(SUBRs!ServerPass) & ""
        m_CommandTimeOut = Val(Trim(SUBRs!TimeOut)) & ""
    End If
    SUBRs.Close
    Set SUBRs = Nothing

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
    
    Query = "EXEC PRO_SMS_001_01 '0', '" & 가맹점정보.지사코드 & "', '" & 가맹점정보.택코드 & "' "

    ADORset.CursorLocation = adUseClient
    ADORset.Open Query, m_Host_DataBase, adOpenStatic, adLockBatchOptimistic, adCmdText
    
    If ADORset.EOF = False Then
        If ADORset.RecordCount > 0 Then
            lblSMS(0).Caption = Format(ADORset!잔여수량 & "", "#,##0")
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
    fpSpread1.MaxRows = 0: fpSpread1.MaxCols = 7
    lblSMS(1).Caption = "0"
    bSSChangeFlag = True

    '-------------------------------------------------------------------------------
    ' 해당 일자에 입고된 고객 번호 내역을 구한다.
    '-------------------------------------------------------------------------------
    Query = " SELECT ISNULL(고객코드,'') AS 고객코드"
    Query = Query & "  FROM TB_입출고 "
    Query = Query & " WHERE 본출일자 = '" & MaskEdBox1(0).Text & MaskEdBox2(0).Text & MaskEdBox3(0).Text & "'"
    Query = Query & "   AND 본출     = '出'"
    Query = Query & "   AND 출고일자 IS NULL"
    Query = Query & "   AND 판매취소 <>  'Y' "
    Query = Query & " GROUP BY 고객코드  "
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    lRow = 0
    Do Until ADORs.EOF
        '-------------------------------------------------------------------------------
        ' 해당 일자의 이전 내용을 포함하여  미입고된 내용이 있는지 확인한다.
        ' 당일날 입고된 내용은 제외 시킨다.
        '-------------------------------------------------------------------------------
        Query = " SELECT * FROM TB_입출고 "
        Query = Query & " WHERE 접수일자 <= '" & MaskEdBox1(0).Text & MaskEdBox2(0).Text & MaskEdBox3(0).Text & "' "
        Query = Query & "   AND NOT ( 본출 = '出' OR 본출 = '反') "
        Query = Query & "   AND 고객코드 = '" & ADORs!고객코드 & "'"
        Query = Query & "   AND 판매취소 <>  'Y' "
        Set SUBRs = New ADODB.Recordset
        SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        ' 미입고된 내용이 없을 경우
        If SUBRs.EOF Then
            ' 해당 회원의 정보를 얻어온다.
            If Fun_고객정보(ADORs!고객코드) <> "Error" Then
            
                If fpSpread1.MaxRows = lRow Then
                    fpSpread1.MaxRows = fpSpread1.MaxRows + 1
                    fpSpread1.RowHeight(fpSpread1.MaxRows) = 20
                End If
                            
                lRow = lRow + 1
                fpSpread1.SetText 2, lRow, 고객정보.고객코드
                fpSpread1.SetText 3, lRow, 고객정보.성명
                fpSpread1.SetText 4, lRow, 고객정보.휴대폰
                
                '-------------------------------------------------------------------------------
                ' 미출고된 내역의 정보를 구한다. 시작텍과 종료텍의 수량을 구한다.
                '-------------------------------------------------------------------------------
                Query = " SELECT   Min(택번호)   As 시작번호"
                Query = Query & ", Max(택번호)   AS 종료번호"
                Query = Query & ", Count(택번호) AS 수량"
                Query = Query & " FROM TB_입출고"
                Query = Query & " WHERE 접수일자 <= '" & MaskEdBox1(0).Text & MaskEdBox2(0).Text & MaskEdBox3(0).Text & "' "
                Query = Query & "   AND (본출 = '出'  OR 본출 = '反') "     ' 본사 출고 상태인것
                Query = Query & "   AND 확인 <> '확' "    ' 미출고 상태인것
                Query = Query & "   AND 고객코드 = '" & ADORs!고객코드 & "' "
                Query = Query & "   AND 판매취소 <>  'Y' "
                Set Rs = New ADODB.Recordset
                Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
                
                If Rs.EOF Then
                    fpSpread1.SetText 5, lRow, Trim(Rs!시작번호) & ""
                    fpSpread1.SetText 6, lRow, Trim(Rs!종료번호) & ""
                    fpSpread1.SetText 7, lRow, CStr(Rs!수량) & ""
                End If
                Rs.Close
                Set Rs = Nothing
                
                If CheckMobileNumber(고객정보.휴대폰, sTel) = True Then
                    fpSpread1.SetText 1, lRow, "1"
                    ' 선택 수량 누적
                    bCountFlag = False
                    lblSMS(1).Caption = Format(Val(Replace(lblSMS(1).Caption, ",", "")) + 1, "#,##0")
                End If
            End If
        End If
        SUBRs.Close
        Set SUBRs = Nothing
        
        ADORs.MoveNext
    Loop
    ADORs.Close
    Set ADORs = Nothing
    
    bSSChangeFlag = False
    Screen.MousePointer = vbDefault
    
    Exit Sub

ErrRtn:
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
    Dim sValue(8)   As String
    Dim vTemp   As Variant
    Dim bResult     As Boolean
   
    On Error GoTo ErrRtn
    
    ' 문자 메시지 길이 확인
    dLng = CheckSendMessageLangth
    
    If dLng <= 0 Or dLng > 80 Then Exit Function
    
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
    
    
    For lRow = 1 To fpSpread1.MaxRows
        Call fpSpread1.GetText(1, lRow, vTemp)
        
        sFlag = IIf(CStr(vTemp) = "1", True, False)
        
        ' 전송 구분일 경우 전송 처리한다.
        If sFlag = True Then
            ' 전송, 메시지타입, 수신번호, 발신번호, 메시지, 지사코드, 대리점코드, 고객코드, 고객성명, 참고5, 참고6
            sValue(0) = "1"
            sValue(1) = "0"
            
            Call fpSpread1.GetText(4, lRow, vTemp):    sValue(2) = CStr(vTemp)
            
            sValue(3) = 가맹점정보.전화번호
            sValue(4) = Trim(txtSMS.Text)
            sValue(5) = 가맹점정보.지사코드
            sValue(6) = 가맹점정보.택코드
            
            Call fpSpread1.GetText(2, lRow, vTemp):    sValue(7) = CStr(vTemp)
            Call fpSpread1.GetText(3, lRow, vTemp):    sValue(8) = CStr(vTemp)
            
            Query = "EXEC PRO_SMS_SEND "
            Query = Query & "'" & sValue(0) & "', "
            Query = Query & "'" & sValue(1) & "', "
            Query = Query & "'" & sValue(2) & "', "
            Query = Query & "'" & sValue(3) & "', "
            Query = Query & "'" & sValue(4) & "', "
            Query = Query & "'" & sValue(5) & "', "
            Query = Query & "'" & sValue(6) & "', "
            Query = Query & "'" & sValue(7) & "', "
            Query = Query & "'" & sValue(8) & "' "
            
            Debug.Print Query
            
            If Dir(App.Path & "\NO_SMS.DAT", vbNormal) = "" Then
                m_Host_DataBase.Execute Query
            End If
        End If
    
    Next lRow
    
    ' 최종 남은 수량을 설정한다.
    Call SetUseSMSCount
    
    Exit Function

ErrRtn:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SendSMS of Form V_SMS001"
End Function


Private Function CheckSendMessageLangth() As Integer
    If IsNumeric(lbl_SMS.Tag) = False Then
        CheckSendMessageLangth = 0
        txtSMS.SetFocus
        MsgBox "전송할 메시지를 입력 하여 주십시요", vbInformation, "확인"
        Exit Function
        
    ElseIf Val(lbl_SMS.Tag) >= 80 Then
        CheckSendMessageLangth = Val(lbl_SMS.Tag)
        txtSMS.SetFocus
        MsgBox "전송할 메시지를 입력 하여 주십시요", vbInformation, "확인"
        Exit Function
    Else
        CheckSendMessageLangth = Val(lbl_SMS.Tag)
        Exit Function
    End If
End Function

'
'
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
'        cboSendText.AddItem Trim(TempRSet!내용 & "")
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


