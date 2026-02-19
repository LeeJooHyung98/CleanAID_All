VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form frm본사전송 
   Caption         =   "본사접속"
   ClientHeight    =   7770
   ClientLeft      =   5025
   ClientTop       =   3525
   ClientWidth     =   9120
   ControlBox      =   0   'False
   LinkTopic       =   "Form22"
   LockControls    =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   9120
   Begin CleanAID.ctlFileTransfer FTC 
      Left            =   7410
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      ReceiveDirPath  =   "D:\InSoftNet\백상\백상매장\Internet"
      RemotePort      =   0
      LocalPort       =   0
      Version         =   1
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   7905
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "localhost"
   End
   Begin VB.TextBox txtPassWord 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   8  '영문
      Left            =   5145
      TabIndex        =   28
      Top             =   1680
      Width           =   2505
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7890
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "본사확인"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7695
      MaskColor       =   &H8000000F&
      TabIndex        =   26
      Top             =   1680
      Width           =   1305
   End
   Begin Threed.SSPanel panMsg 
      Height          =   675
      Left            =   60
      TabIndex        =   2
      Top             =   960
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   1191
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
      BevelWidth      =   2
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7440
      Top             =   120
   End
   Begin MSMask.MaskEdBox txtDate 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1680
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####-##-##"
      PromptChar      =   "_"
   End
   Begin Threed.SSPanel SSPanel6 
      Height          =   375
      Left            =   60
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   262144
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "일자"
      BevelWidth      =   2
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1470
      Left            =   60
      TabIndex        =   5
      Top             =   2100
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   2593
      _Version        =   262144
      BackColor       =   -2147483632
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
      BevelWidth      =   2
      RoundedCorners  =   0   'False
      Begin ComctlLib.ProgressBar pgbStatus 
         Height          =   330
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   8715
         _ExtentX        =   15372
         _ExtentY        =   582
         _Version        =   327682
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlMsg 
         Height          =   795
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   1402
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
         BevelWidth      =   2
         RoundedCorners  =   0   'False
      End
   End
   Begin Threed.SSPanel pnlSet 
      Height          =   1575
      Left            =   60
      TabIndex        =   8
      Top             =   3600
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   2778
      _Version        =   262144
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
      BevelWidth      =   2
      RoundedCorners  =   0   'False
      Begin VB.Timer Timer_FTC 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   2385
         Top             =   210
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "도움말"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   7290
         Picture         =   "frm본사전송.frx":0000
         Style           =   1  '그래픽
         TabIndex        =   11
         Top             =   630
         Width           =   1575
      End
      Begin VB.TextBox txtDir 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1590
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   675
         Width           =   5595
      End
      Begin VB.TextBox txtDir2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1590
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1050
         Width           =   5595
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   495
         Left            =   150
         TabIndex        =   12
         Top             =   120
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   873
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "환경설정"
         BevelWidth      =   2
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   345
         Index           =   8
         Left            =   120
         TabIndex        =   13
         Top             =   675
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "전송작업경로"
         BevelWidth      =   2
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   345
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "수신작업경로"
         BevelWidth      =   2
         RoundedCorners  =   0   'False
      End
   End
   Begin VB.PictureBox RasDial 
      Height          =   480
      Left            =   8475
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   29
      Top             =   120
      Width           =   1200
   End
   Begin Threed.SSCommand cmdBtn 
      Height          =   795
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1402
      _Version        =   262144
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "본사접속"
      ButtonStyle     =   2
   End
   Begin Threed.SSCommand cmdBtn 
      Height          =   795
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   60
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1402
      _Version        =   262144
      Enabled         =   0   'False
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "접속종료"
      ButtonStyle     =   2
   End
   Begin Threed.SSCommand cmdIpgoSend 
      Height          =   780
      Left            =   60
      TabIndex        =   15
      Top             =   5220
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   1376
      _Version        =   262144
      Enabled         =   0   'False
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "입고자료보내기"
      ButtonStyle     =   2
      BevelWidth      =   3
   End
   Begin Threed.SSCommand cmdSaleRecv 
      Height          =   780
      Left            =   6060
      TabIndex        =   16
      Top             =   5220
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   1376
      _Version        =   262144
      Enabled         =   0   'False
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "할인자료받기"
      ButtonStyle     =   2
      BevelWidth      =   3
   End
   Begin Threed.SSCommand cmdDaySale 
      Height          =   780
      Left            =   60
      TabIndex        =   17
      Top             =   6060
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   1376
      _Version        =   262144
      ForeColor       =   0
      Enabled         =   0   'False
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "목요세일자료받기"
      ButtonStyle     =   2
      BevelWidth      =   3
   End
   Begin Threed.SSCommand cmdChulgoRecv 
      Height          =   780
      Left            =   3060
      TabIndex        =   18
      Top             =   5220
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   1376
      _Version        =   262144
      Enabled         =   0   'False
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "출고자료받기"
      ButtonStyle     =   2
      BevelWidth      =   3
   End
   Begin Threed.SSCommand cmdPriceRecv 
      Height          =   780
      Left            =   3060
      TabIndex        =   19
      Top             =   6060
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   1376
      _Version        =   262144
      Enabled         =   0   'False
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "가격자료받기"
      ButtonStyle     =   2
      BevelWidth      =   3
   End
   Begin Threed.SSCommand cmdRepair 
      Height          =   780
      Left            =   6060
      TabIndex        =   20
      Top             =   6060
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   1376
      _Version        =   262144
      Enabled         =   0   'False
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "수선자료받기"
      ButtonStyle     =   2
      BevelWidth      =   3
   End
   Begin Threed.SSCommand cmdBtn 
      Height          =   795
      Index           =   2
      Left            =   3300
      TabIndex        =   21
      Top             =   60
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1402
      _Version        =   262144
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "종  료"
      ButtonStyle     =   2
   End
   Begin Threed.SSCommand cmdDBSend 
      Height          =   780
      Left            =   60
      TabIndex        =   22
      Top             =   6900
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   1376
      _Version        =   262144
      Enabled         =   0   'False
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "DB 본사전송"
      ButtonStyle     =   2
      BevelWidth      =   3
   End
   Begin Threed.SSCommand cmdCustSend 
      Height          =   780
      Left            =   3060
      TabIndex        =   23
      Top             =   6900
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   1376
      _Version        =   262144
      Enabled         =   0   'False
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "고객자료보내기"
      ButtonStyle     =   2
      BevelWidth      =   3
   End
   Begin Threed.SSCommand cmdPGRecv 
      Height          =   780
      Left            =   6060
      TabIndex        =   24
      Top             =   6900
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   1376
      _Version        =   262144
      Enabled         =   0   'False
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "이전출고자료받기"
      ButtonStyle     =   2
      BevelWidth      =   3
   End
   Begin Threed.SSCommand cmdBtn 
      Height          =   795
      Index           =   3
      Left            =   4950
      TabIndex        =   27
      Top             =   60
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   1402
      _Version        =   262144
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "프로그램종료"
      ButtonStyle     =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "본사확인코드 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3465
      TabIndex        =   25
      Top             =   1740
      Width           =   1620
   End
End
Attribute VB_Name = "frm본사전송"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'
''Dim daoDB As Database        'Access DB
''Dim GS_RS As Recordset
''Dim daoRS As Recordset
''Dim daoRS1 As Recordset
''Dim daoQD As QueryDef
''Dim daoQD1 As QueryDef
'
'Dim SendMode As Integer         ' 1: Modem,  2, Internet
'Dim SendFlag As LaundrySendFlag ' 현재 어떤 작업을 하였는지
'Dim g_AgencyCode As String
'Dim b_OldDataRecv As Boolean    ' 이전 자료 받는지의 여부
'
'Dim FileTotalCount As Integer   ' 본사에서 한번에 수신할 파일수
'Dim FileProCount As Integer     ' 현재 받고 있는 수
'Dim strFileList As String       ' 수신할 파일이름을 보관하고 있음
'Dim NextPross As Boolean        ' 인터넷 작업시 다음 작업 작업 여부
'Dim GS_ISQL As String
'Dim St_Data(5) As String
'Dim strSendPath As String       ' 서버로 전송한 자료가 보관될 위치
'Dim strRecvPath As String       ' 서버에 수신할 자료가 보관되어 있는 위치
'Dim strPrgPath As String        ' 서버에 수신할 프로그래이 보관되어 있는 위치
'Dim strSendData As String       ' 서버에 전송할 메시지
'Dim PauseTime, Start, Finish, TotalTime ' 서버의 전송을 기다리는데 필요
'Dim strNewVersion As String     ' 신규 프로그램이 있을경우
'Dim SendFile    As String       ' 서버에 전송할 파일이름
'Dim Ret As Boolean
'Dim FormActivate As Boolean
'Dim m_MstCode   As String
'
''+------------------------------------------------------
''+ 2003/02/11 수정
''+
''+루틴설명
''+  index - 0 :     본사접속
''+  index - 1 :     접속종료
''+  index - 2 :     종료
''+  index - 3 :     프로그램종료
''+------------------------------------------------------
'Private Sub cmdBtn_Click(Index As Integer)
'    Dim lStatus As Long
'
'    ' 본사확인코드 메시지를 죽이기 위하여 타이머를 저지한다.
'    Timer2.Enabled = False
'    pnlMsg.Caption = ""
'    pnlMsg.BackColor = &H8000000F
'
'    Select Case Index
'        Case 0
'            ' 원격다이알을 정의한다.
'            RasDial.EntryName = GetIniStr("SERVER", "RasDialName", "", iniFile)
'            RasDial.UserName = GetIniStr("SERVER", "UserID", "", iniFile)
'            RasDial.Password = GetIniStr("SERVER", "Password", "", iniFile)
'
'            ' 본사에 전화를 건다.
'            panMsg.Caption = "본사에 전화를 거는 중입니다..."
'
'            lStatus = RasDial.Dial
''            panMsg.Caption = "Dial returned Status = " & Format(lStatus)
'
'            If lStatus = 0 Then
'                panMsg.Caption = "정상적으로 연결이 되었습니다..."
'            ElseIf lStatus = 600 Then
'                panMsg.Caption = "정상적으로 연결이 되었습니다..."
'            Else
'                If lStatus = 678 Then
'                    panMsg.Caption = "전화 받는 컴퓨터에서 응답하지 않습니다..."
'                ElseIf lStatus = 720 Then
'                    panMsg.Caption = "전화 접속 네트워킹에서 서버 종류 설정에 지정한 호환 가능한 네트워크 프로토콜 세트를 교섭할 수 없습니다..."
'                ElseIf lStatus = 602 Then
'                    panMsg.Caption = "통화중 입니다... "
'                End If
'
'                RasDial.HangUp
'
'                ' 원격다이알을 정의한다.
'                RasDial.EntryName = GetIniStr("SERVER", "RasDialName2", "", iniFile)
'                RasDial.UserName = GetIniStr("SERVER", "UserID", "", iniFile)
'                RasDial.Password = GetIniStr("SERVER", "Password", "", iniFile)
'
'                ' 본사에 전화를 건다.
'                panMsg.Caption = "[" & lStatus & "] 본사2에 전화를 거는 중입니다..."
'
'                lStatus = RasDial.Dial
'
'                If lStatus = 0 Then
'                    panMsg.Caption = "정상적으로 연결이 되었습니다..."
'                ElseIf lStatus = 600 Then
'                    panMsg.Caption = "정상적으로 연결이 되었습니다..."
'                ElseIf lStatus = 678 Then
'                    panMsg.Caption = "전화 받는 컴퓨터에서 응답하지 않습니다..."
'                    RasDial.HangUp
'                    Exit Sub
'                ElseIf lStatus = 602 Then
'                    panMsg.Caption = "통화중 입니다... "
'                    RasDial.HangUp
'                    Exit Sub
'                Else
'                    panMsg.Caption = "본사연결에 실패하였습니다."
'                    RasDial.HangUp
'                    Exit Sub
'                End If
'            End If
'
'            If RasDial.AsyncMode = False Then
'                Select Case RasDial.RasStatus
'                    Case RASCS_Connected
'                        Timer1.Enabled = True
'                    Case RASCS_Disconnected
'                        RasDial.HangUp
'                        MsgBox "본사와의 연결이 종료 되었습니다.", vbInformation
'                End Select
'            End If
'        Case 1
'            ' 본사와 접속을 종료한다.
'            RasDial.HangUp
'            MsgBox "본사와의 연결이 종료 되었습니다.", vbInformation
'
'            cmdBtn(1).Enabled = False       ' 접속종료
'
'            Call ButtonEnable(False)
'
'        Case 2
'            ' 신규 대리점 정보가 있을 경우
'            If 가맹점정보.가맹점코드 <> "000000" And panMsg.Tag <> "ERROR" Then
'
'                pnlMsg.Caption = "본사 자료 전송중 입니다. 잠시만 기다려 주십시요."
'                DoEvents
'
'                ' 각종 테이블의 자료를 전송한다.
'                pgbStatus.Visible = True
'
'                ' 프로그램의 버전을 설정한다.
'                Call SendProgramVersion
'
'                ' 본사에서 'N'를 설정한 자룔를 다시 전송한다.
'                Call SendNoSalesData
'
'                Call SendTable_Data(pgbStatus, True)
'
'                pnlMsg.Caption = "메일 자료를 수신중 입니다. 잠시만 기다려 주십시요."
'                DoEvents
'
'                Call GetMailData
'
'                pgbStatus.Visible = False
'            End If
'
'            Unload Me ' 본사전송 종료
'
'        Case 3
'            ' 프로그램 종료
'            RasDial.HangUp
'            Unload Me
'            End
'    End Select
'End Sub
'
''+------------------------------------------------------
''+ 2003/02/11 수정
''+
''+루틴설명      - 고객자료보내기
''+
''+------------------------------------------------------
'Private Sub cmdCustSend_Click()
'    ' 이전 자료 받기 때문에 모두 해준다
'    g_AgencyCode = 가맹점정보.택코드
'    m_MstCode = 가맹점정보.지사코드
'
'    Call ButtonEnable(False)
'
'    pgbStatus.Value = 0
'
'    If SendMode = 1 Then
'        ' 모뎀
'        Ret = SendCust
'    Else
'        ' 인터넷
'        pnlMsg.Caption = ""
'        Ret = SendCust
'    End If
'
'    Call ButtonEnable(True)
'End Sub
'
'Private Sub cmdDBSend_Click()
''+------------------------------------------------------
''+ 2003/02/11 수정
''+
''+루틴설명      - DB 본사전송
''+
''+------------------------------------------------------
'    Dim SendFile As String
'
'    On Error GoTo ErrRtn
'
'    ' 이전 자료 받기 때문에 모두 해준다
'    g_AgencyCode = 가맹점정보.택코드
'    m_MstCode = 가맹점정보.지사코드
'
'
'    pnlMsg.Caption = "본사에 DB를 전송 준비중 입니다. 최대 3분까지 기다려 주십시요."
'    Call ButtonEnable(False)
'    pgbStatus.Value = 0
'
'    ' DB를 닫는다.
'    'daoDB.Close
'    'Set daoDB = Nothing
'
'    ' Bat 화일이 있는지 확인한다.
'    Call CrateCommondFiles(DBSend, True)
'
''    DoEvents
'    '  이전 "ok.ok" 화일을 지운다
'    If Not Dir("C:\Laundry\DBSend.OK") = "" Then Kill "C:\Laundry\DBSend.OK"
'
'    Shell App.Path & "\DBSend.BAT", vbHide
'    pnlMsg.Caption = "파일을 압축하는중입니다. 최대 3분 까지 기다려 주십시요."
'
'    PauseTime = 180                 ' 기간을 지정합니다.
'    Start = Timer                   ' 시작 시간을 지정합니다.
'
'    Do While Dir("C:\Laundry\DBSend.ok") = ""
'        Finish = Timer              ' 종료 시간을 지정합니다.
'        TotalTime = Finish - Start  ' 전체 시간을 계산합니다.
'        If Timer > Start + PauseTime Then
'            GoTo ErrRtn
'        End If
''        DoEvents                    ' 다른 프로시저로 넘깁니다.
'    Loop
'
'    pnlMsg.Caption = "압축 완료."
'
'
'    ' 본사에 파일을 전송한다.
'    pnlMsg.Caption = "DB파일을 본사에 전송중 입니다."
'
'    SendFile = "C:\Laundry\DB\" & g_AgencyCode & ".zip"
'    SendFlag = lauSendDB
'
'    Call InterNetSendFiles(SendFile)
'
'    ' DB 연결
'    'If Not DB_Connect Then
'    '    End
'    'End If
'
'    'Set daoDB = Workspaces(0).OpenDatabase(App.Path & "\DB\Laundry.MDB")
'
'    Call ButtonEnable(True)
'    Exit Sub
'
'ErrRtn:
'    pnlMsg.Caption = "파일 전송에 실패하였습니다."
'    Call ButtonEnable(True)
'
'    'Set daoDB = Workspaces(0).OpenDatabase(App.Path & "\DB\Laundry.MDB")
'End Sub
'
'Private Sub cmdPGRecv_Click()
'    Dim Scode       As String
'    Dim sMstCode    As String
'    Dim sDate       As String
'
'    Scode = Trim(GetIniStr("Store", "OldCode", "", iniFile))
'
'    If Scode = "" Then
'        MsgBox "등록된 이전 체인점 코드가 없습니다.", vbInformation, "확인"
'        Exit Sub
'    End If
'
'    sMstCode = Trim(GetIniStr("Store", "OldMstCode", "", iniFile))
'
'    If sMstCode = "" Then sMstCode = 가맹점정보.지사코드
'
'    sDate = Trim(GetIniStr("Store", "OldDate", "", iniFile))
'
'    If sDate = "" Or Not IsDate(sDate) Or sDate < Date Then
'        MsgBox "이전 일자의 자료를 받을 수 있는 기간이 경과 하였습니다. " & vbNewLine & "[" & sDate & "]", vbInformation, "확인"
'        Exit Sub
'    End If
'
'    ' 이전 자료 받기 때문에 모두 해준다
'    'g_AgencyCode = 가맹점정보.택코드
'    g_AgencyCode = Scode
'    m_MstCode = sMstCode
'    b_OldDataRecv = True
'
'    ' 클라이언트 소켓을 종료시킴
'    Winsock1.Close
'    FTC.RemoteClose
'
'    If Winsock1.State <> sckClosed Then Winsock1.Close
'    Do
'        DoEvents
'        If Winsock1.State = sckListening Then
'            ' 클라이언트 소켓을 종료시킴
'            Winsock1.Close
'            Exit Do
'
'        End If
'
'        If Winsock1.State = sckError Or Winsock1.State = sckClosed Then
'            Exit Do
'        End If
'    Loop
'
'    Call ConnectMainServer(g_AgencyCode)
'    DoEvents
'    DoEvents
'
'    Call cmdChulgoRecv_Click
'    DoEvents
'    DoEvents
'
''    ' 클라이언트 소켓을 종료시킴
''    Winsock1.Close
''    FTC.RemoteClose
''    If Winsock1.State <> sckClosed Then Winsock1.Close
''    Do
''        DoEvents
''        If Winsock1.State = sckListening Then
''            ' 클라이언트 소켓을 종료시킴
''            Winsock1.Close
''            Exit Do
''
''        End If
''
''        If Winsock1.State = sckError Or Winsock1.State = sckClosed Then
''            Exit Do
''        End If
''    Loop
''
''    ' 파일 수신 완료후 g_AgencyCode에 등록되어 있는 내용으로 작업을 하기 때문에
''    Call ConnectMainServer(가맹점정보.택코드)
'    b_OldDataRecv = False
'
'End Sub
'
'Private Sub Command1_Click()
''+------------------------------------------------------
''+ 2003/02/11 수정
''+
''+루틴설명      - 비밀번호확인
''+  1. 암호를 확인하여 암호 규칙에 맞으면 화면을 종료한다.
''+  2. 레지스터리에 저장한다.
''+
''+------------------------------------------------------
'    Dim strPass As String
'    ' 입력 확인
'    If Len(txtPassWord.Text) <= 0 Then
'        Exit Sub
'    End If
'
''   기본 디폴드 암호.. ( 프로그램 셋팅/설치를 위한 암호 )
'    If UCase(txtPassWord.Text) = "DUDTJSGH" Then
'        chkPassWord = True
'        Unload Me
'        Exit Sub
'    End If
'
'    ' 비밀번호 확인
'    strPass = IsPassWord(txtPassWord.Text)
'
'    If strPass = "-1" Or strPass = "-3" Then
'        chkPassWord = False
'        txtPassWord.SelStart = 0: txtPassWord.SelLength = Len(txtPassWord.Text)
'
'        If strPass = "-3" Then MsgBox "입력한 내용이 정확하지 않습니다.", vbInformation, "입력오류"
'
'        txtPassWord.Text = ""
'        txtPassWord.SetFocus
'        Exit Sub
'    Else
'        If Not IsPassREGSave(txtPassWord.Text) Then
'            chkPassWord = False
'            MsgBox "입력한 내용이 레지스터리에 저장되지 않았습니다.", vbInformation, "저장오류"
'            Exit Sub
'        Else
'            chkPassWord = True
'            Unload Me
'        End If
'    End If
'
'End Sub
'
''+------------------------------------------------------
''+ 2003/02/11
''+
''+루틴설명
''+
''+  1. 조회와 본사확인코드를 구분하기 위하여
''+  2. 본사확인코드시 포커스를 본사확인코드 부분에 맞추기 위하여
''+
''+------------------------------------------------------
'Private Sub Form_Activate()
'    Dim sDate   As String
'
'    ' 오류 발생시
'    ' 인터넷이 안돼는 PC에서 인터넷으로 설정한 경우 오류가 발생한다.
'    On Error GoTo ErrRtn
'
'    If FormActivate Then Exit Sub
'
'    FormActivate = True
'
'    sDate = Trim(GetIniStr("Store", "OldDate", "", iniFile))
'
'    If sDate = "" Or Not IsDate(sDate) Or sDate < Date Then
'        cmdPGRecv.Enabled = False
'    End If
'
'
'    If Not chkPassWord Then
'        txtPassWord.SetFocus
'        'pnlMsg.Caption = frmMain.Title.Tag & " 미전송된 자료를 전송하여 주십시요 "
'        pnlMsg.Caption = " 미전송된 자료를 전송하여 주십시요 "
'
'        Timer2.Enabled = True
'    End If
'
''    DoEvents
'
'    If SendMode = 2 Then
'
'    Call ConnectMainServer(g_AgencyCode)
'
'    End If
'    Exit Sub
'
'ErrRtn:
'    MsgBox "ERROR NUMBER : " & Err.Number & vbLf & vbLf & _
'            "ERROR DESCR. : " & Err.Description & vbLf & vbLf & _
'            "인터넷 설정을 확인하여 주십시요.", vbInformation, "확인"
'End Sub
'
'Private Sub ConnectMainServer(ByVal sConnectStoreCode As String)
'
'    ' 소켓상태가 연결인 경우는 재처리 안함
'    If Winsock1.State <> sckConnected Then
'
'        panMsg.Caption = "본사와 연결중 입니다. 잠시만 기다려 주십시요."
'
'        ' 프로그램 종료 버튼을 잠시 중지 시킨다.
'        cmdBtn(3).Enabled = False
'        Call ButtonEnable(False)
'
'        ' 클라이언트 소켓을 종료시킴
'        Winsock1.Close
'
'        ' 클라이언트 소켓이 종료할 때까지 기다림
'        Do While Winsock1.State <> sckClosed
'            DoEvents
'        Loop
'
'        ' 클라이언트에서 서버에 연결을 시도함
'        Winsock1.RemoteHost = Fn_GetRemoteIP
'        Winsock1.RemotePort = Fn_GetMsgRemotePort
'        Winsock1.Connect
'
'        ' 클라이언트에서 서버에 연결이 완료할 때까지 기다림
'        Do While Winsock1.State <> sckConnected
'            DoEvents
'            If Winsock1.State = sckError Then
'                panMsg.Caption = "본사 연결 오류"
'                panMsg.Tag = "ERROR"
'                cmdBtn(3).Enabled = True
'                Exit Sub
'            End If
'        Loop
'        ' 다시 활성화 한다.
'        cmdBtn(3).Enabled = True
'
'    End If
'
'
'    ' 전송할 메시지 작성
'    strSendData = S_STA & "|" & "CLEANAID" & "|" & m_MstCode & "|" & _
'                  "STORE_CONNECT" & "|" & _
'                   m_MstCode & ";" & sConnectStoreCode & ";" & 가맹점정보.가맹점명 & "|" & _
'                  S_END
'
'    '        strSendData = S_STA & "|" & S_MYIP & "=" & GetIPAddress & "|" & _
'    '                      S_CUSTCODE & "=" & g_AgencyCode & "|" & _
'    '                      S_CUSTNAME & "=" & 가맹점정보.가맹점명 & "|" & _
'    '                      S_MYFILEPORT & "=" & Fn_GetFileLocatPort & "|" & _
'    '                      S_END
'
'    ' 소켓상태가 연결인 경우만 데이타를 보냄
'    If Winsock1.State = sckClosed Or Winsock1.State = sckError Then
'        pnlMsg.Caption = "본사와 연결되지 않았습니다."
'        Exit Sub
'    End If
'
'    ' 소켓상태가 연결인 경우만 데이타를 보냄
'    FileTotalCount = -1
'    If Winsock1.State = sckConnected Then
'        ' 데이타를 서버에 보냄
'        Winsock1.SendData strSendData
'    End If
'
'End Sub
'
'Private Sub Form_Load()
'    Dim St As String
'    Dim strDBPath   As String
'
'    Me.Left = (Screen.Width - Me.Width) / 2
'    Me.Top = (Screen.Height - Me.Height) / 2
'
'    txtDate.Text = Format(Date, "YYYY-MM-DD")
'
'    m_MstCode = 가맹점정보.지사코드
'
'    '기본을 모뎀으로 한다.
'    If GetSetting("Laundry_Zi", "Connect", "Type", "True") Then
'        SendMode = 1
'        strSendPath = GetIniStr("DIRPATH", "DataPath", "", iniFile)
'        strRecvPath = GetIniStr("DIRPATH", "RecvPath", "", iniFile)
'        strPrgPath = GetIniStr("DIRPATH", "PrgPath", "", iniFile)
'
'        txtDir.Text = strRecvPath
'        txtDir2.Text = strSendPath
'
'        Connect_Gb = False
'    Else
'        ' 인터넷
'        SendMode = 2
'        SSPanel4(8).Caption = "Server IP "
'        SSPanel4(3).Caption = "Client IP "
'
'        txtDir.Text = Fn_GetRemoteIP
'        txtDir2.Text = GetIPAddress
'
'        ' 인터넷일경우 파일을 다른곳에서 복사하여 모뎀의 함수를 공통으로 사용하기 위하여
'        ' 인터넷으로 받은 파일이 위치한 경로로 변경하여 준다.
'        strSendPath = App.Path & "\RecvData"
'
'        Connect_Gb = True
'        cmdBtn(0).Enabled = False
'        cmdBtn(1).Enabled = False
'        Timer1.Enabled = True
'    End If
'
'    'Set daoDB = Workspaces(0).OpenDatabase(m_DBPath)
'    'Set daoRS = daoDB.OpenRecordset("Select 택코드 FROM TB_기본정보")
'
'    g_AgencyCode = 가맹점정보.택코드
'
'    If g_AgencyCode = "" Then
'        Call ButtonEnable(False)
'    End If
'
'End Sub
'
'Private Sub FTC_ChangeState(ByVal NewState As TransferState)
'    Select Case NewState
'        Case ftcReady
'            ' 본사에 연결되면
'            ' 파일이 전송중일때 중간에 일시적으로 전송 대기 모드가 됨
'            ' 그렇지 않으면 인터넷 전송때문에 순서대로 안들어옴
'            If FileProCount >= FileTotalCount Then
'                panMsg.Caption = "본사와 연결 되었습니다. "
'                Call ButtonEnable(True)
'            End If
'        Case Else
'
'    End Select
'
'End Sub
'
'Private Sub FTC_Error(ByVal Number As Long, Description As String)
'    Select Case Number
'        Case 10061
'            panMsg.Caption = "파일 서버에 연결 오류" & vbLf & _
'                             "잠시후 다시 시도하여 주십시요."
'            Call ButtonEnable(False)
'        Case Else
'            panMsg.Caption = "[" & CStr(Number) & "]" & Description
'    End Select
'End Sub
'
'Private Sub FTC_ReceiveProgress(ByVal ReceiveSize As Long)
'    On Error GoTo ErrRtn
'
'    pgbStatus.Value = ReceiveSize
'
'    Exit Sub
'
'ErrRtn:
'
'End Sub
'
'Private Sub RasDial_Status(ByVal RasConnState As Long, ByVal dwError As Long)
'    If dwError > 0 Then
'        panMsg.Caption = GetStatus(dwError)
'    Else
'        panMsg.Caption = GetStatus(RasConnState)
'    End If
'
'    Select Case RasConnState
'        Case RASCS_Connected
'            Timer1.Enabled = True
'        Case RASCS_Disconnected
'            RasDial.HangUp
'            MsgBox "본사와의 연결이 종료 되었습니다.", vbInformation
'    End Select
'End Sub
'
'Private Function GetStatus(RasStatus As Long) As String
'    Dim StatusString As String
'
'    Select Case RasStatus
'        Case RASCS_OpenPort:            StatusString = "포트를 OPEN 하는 중 입니다..."
'        Case RASCS_PortOpened:          StatusString = "포트가 OPEN 되었습니다."
'        Case RASCS_ConnectDevice:       StatusString = "디바이스에 연결하는 중입니다..."
'        Case RASCS_DeviceConnected:     StatusString = "디바이스에 연결되었습니다."
'        Case RASCS_AllDevicesConnected: StatusString = "모든 디바이스에 연결되었습니다."
'        Case RASCS_Authenticate:        StatusString = "사용자를 인증하고 있습니다..."
'        Case RASCS_AuthNotify:          StatusString = "AuthNotify"
'        Case RASCS_AuthRetry:           StatusString = "사용자 인증 재시도중..."
'        Case RASCS_AuthCallback:        StatusString = "AuthCallback"
'        Case RASCS_AuthChangePassword:  StatusString = "비밀번호가 잘못되었습니다."
'        Case RASCS_AuthProject:         StatusString = "AuthProject"
'        Case RASCS_AuthLinkSpeed:       StatusString = "AuthLinkSpeed"
'        Case RASCS_AuthAck:             StatusString = "AuthAck"
'        Case RASCS_ReAuthenticate:      StatusString = "ReAuthenticate"
'        Case RASCS_Authenticated:       StatusString = "사용자가 인증되었습니다."
'        Case RASCS_PrepareForCallback:  StatusString = "PrepareForCallback"
'        Case RASCS_WaitForModemReset:   StatusString = "WaitForModemReset"
'        Case RASCS_WaitForCallback:     StatusString = "WaitForCallback"
'        Case RASCS_Projected:           StatusString = "Projected"
'        Case RASCS_StartAuthentication: StatusString = "사용자 인증을 시작하는 중 입니다."
'        Case RASCS_CallbackComplete:    StatusString = "CallbackComplete"
'        Case RASCS_LogonNetwork:        StatusString = "네트워크에 로그온이 되었습니다."
'        Case RASCS_Interactive:         StatusString = "상호네트워크 Checking 중..."
'        Case RASCS_RetryAuthentication: StatusString = "사용자 인증을 다시 하시기 바랍니다."
'        Case RASCS_CallbackSetByCaller: StatusString = "CallbackSetByCaller"
'        Case RASCS_PasswordExpired:     StatusString = "비밀번호가 만료되었습니다."
'        Case RASCS_Connected:           StatusString = "본사에 접속이 되었습니다."
'        Case RASCS_Disconnected:        StatusString = "본사에 접속이 되지 않았습니다."
'        Case 0:                         StatusString = "본사에 접속이 되지 않았습니다."
'        Case RASBASE:                   StatusString = "접속 중..."
'        Case Else:                      StatusString = "RAS Error " & RasStatus
'    End Select
'
'    GetStatus = StatusString
'End Function
'
'Private Sub Timer_FTC_Timer()
'    '타이머를 죽인다.
'    Timer_FTC.Enabled = False
'
'    Select Case UCase(Timer_FTC.Tag)
'
'        Case "RECVMAIL"
'            Call MailData
'            NextPross = True
'            Exit Sub
'
'        Case "CHULGO"
'            Call ButtonEnable(False)
'            Call ChulgoData
'            Call ButtonEnable(True)
'            NextPross = False
'            Exit Sub
'
'        Case "SALEDATA"
'            Call ButtonEnable(False)
'            Call SaleData
'            Call ButtonEnable(True)
'            NextPross = False
'            Exit Sub
'
'        Case "DAYSALEDATA"
'            Call ButtonEnable(False)
'            Call DaySaleData
'            Call ButtonEnable(True)
'            NextPross = False
'            Exit Sub
'
'        Case "PRICEDATA"
'            Call ButtonEnable(False)
'            Call PriceData
'            Call ButtonEnable(True)
'            NextPross = False
'            Exit Sub
'
'        Case "REPAIRDATA"
'            Call ButtonEnable(False)
'            Call RepairData
'            Call ButtonEnable(True)
'            NextPross = False
'            Exit Sub
'
'        Case "PROGRAM"
'            Call ButtonEnable(False)
'            Call ProgramUpgrade
'            Call ButtonEnable(True)
'            NextPross = False
'            Exit Sub
'
'    End Select
'
'
'End Sub
'
'Private Sub Timer1_Timer()
'    Timer1.Enabled = False
'
'    If Connect_Gb = False Then
'        ' 모뎀일 경우에만
'        RasDial.MinimizeRas
'        MsgBox RasDial.EntryName & "에 연결이 되었습니다.", vbInformation
'
'        cmdBtn(1).Enabled = True        ' 접속종료
'
'        Call ButtonEnable(True)
'    End If
'End Sub
'
'Private Sub Timer2_Timer()
''+------------------------------------------------------
''+ 2003/02/11
''+
''+       본사 확인 코드 메시지 출력을 위하여
''+
''+------------------------------------------------------
'    If pnlMsg.BackColor = vbButtonFace Then '&HC0C0C0 Then
'        pnlMsg.BackColor = &HFFC0FF
'    Else
'        pnlMsg.BackColor = vbButtonFace '&HC0C0C0
'    End If
'End Sub
'
'
''+------------------------------------------------------
''+ 2003/02/11 수정
''+
''+루틴설명      - 입고 자료 보내기
''+  1. 일자의 내용을 본사에 전송한다.
''+  2. 매일이 있을 경우 매일도 같이 전송한다.
''+
''+------------------------------------------------------
'Private Sub cmdIpgoSend_Click()
'    ' 이전 자료 받기 때문에 모두 해준다
'    g_AgencyCode = 가맹점정보.택코드
'    m_MstCode = 가맹점정보.지사코드
'
'    Call ButtonEnable(False)
'
'    pgbStatus.Value = 0
'
'    If SendMode = 2 And FTC.State <> ftcReady Then
'        pnlMsg.Caption = "잠시후 다시 시도하여 주십시요."
'        Exit Sub
'    End If
'
'    Ret = SendMail
'
'    DoEvents
'    Ret = SendCust
'
'    DoEvents
'    If 가맹점정보.마일리지여부 = "Y" Then
'        Ret = SendMileageData
'    End If
'
''    DoEvents
''    ret = QN_Data
'
'    DoEvents
'    GoSub SUB_FTC_CHACK
'    Ret = SendCoupon
'
'    DoEvents
'
'    GoSub SUB_FTC_CHACK
'
'    Ret = InputData
'
'
'    Call ButtonEnable(True)
'    Exit Sub
'
'SUB_FTC_CHACK:
'    ' 인터넷 전송 부분
'    DoEvents
'    PauseTime = 10                 ' 기간을 지정합니다.
'    Start = Timer                   ' 시작 시간을 지정합니다.
'
'    Do While FTC.State <> ftcReady
'        Finish = Timer              ' 종료 시간을 지정합니다.
'        TotalTime = Finish - Start  ' 전체 시간을 계산합니다.
'
'        If Timer > Start + PauseTime Then
'            panMsg.Caption = "SUB_FTC_CHACK 실패"
'            Exit Sub
'        End If
'        DoEvents                    ' 다른 프로시저로 넘깁니다.
'    Loop
'    Return
'
'
'End Sub
'
''+------------------------------------------------------
''+ 2003/02/11 수정
''+
''+루틴설명      - 출고자료 보내기
''+  1. 일자의 내용을 본사에서 내려 받는다.
''+  2. 매일이 있을 경우 매일도 같이 내려 받느다.
''+
''+------------------------------------------------------
'Private Sub cmdChulgoRecv_Click()
'
'    ' 수신이 완료될때 까지 다시 누르는것을 방지하기 위하여
'    ' 활성화는 Timer_FTC에서 시킨다.
'
'
'    ' 이전 자료 받기 때문에 모두 해준다
'    If b_OldDataRecv = False Then
'        g_AgencyCode = 가맹점정보.택코드
'        m_MstCode = 가맹점정보.지사코드
'    End If
'
'    Call ButtonEnable(False)
'    pgbStatus.Value = 0
'
'    If SendMode = 1 Then
'        ' 모뎀
'        Ret = ChulgoData
'        Ret = MailData
'        Call ButtonEnable(True)
'    Else
'        ' 인터넷
'        ' 먼저 파일 리스트를 받아 수신할 파일의 수를 구한다.
'        ' 해당 파일의 수만큼 파일을 수신한다.
'        NextPross = False
'        pnlMsg.Caption = ""
'        FileProCount = 0
'
'        Call ButtonEnable(False)
'        SendFlag = lauRecvMail
'        Ret = InterNetFileRequest(lauRecvMail)
'
'        ' Timer_FTC에서 메일이 완료될때까지 기다린다.
'        Do While NextPross = False
'            DoEvents                    ' 다른 프로시저로 넘깁니다.
'        Loop
'        SendFlag = lauChulGo
'        Call ButtonEnable(False)
'        Ret = InterNetFileRequest(lauChulGo)
'
'        ' 이전 자료 받기일 경우 출고 자료를 받지 못하도록 수정
'        ' 다시 나갔다 와서 기존 입고 자료를 받을 수 있도록 수정
'        If b_OldDataRecv = False Then
'            Call ButtonEnable(True)
'        Else
'            cmdPGRecv.Enabled = True
'        End If
'    End If
'
'
'End Sub
'
''+------------------------------------------------------
''+ 2003/02/11 수정
''+
''+루틴설명      - 목요세일자료받기
''+  1. 일자의 내용을 본사에서 내려 받는다.
''+  2. 매일이 있을 경우 매일도 같이 내려 받느다.
''+
''+------------------------------------------------------
'Private Sub cmdDaySale_Click()
'    ' 이전 자료 받기 때문에 모두 해준다
'    g_AgencyCode = 가맹점정보.택코드
'    m_MstCode = 가맹점정보.지사코드
'
'    Call ButtonEnable(False)
'
'    pgbStatus.Value = 0
'
'    If SendMode = 1 Then
'        ' 모뎀
'        Ret = DaySaleData
'    Else
'        ' 인터넷
'        ' 먼저 파일 리스트를 받아 수신할 파일의 수를 구한다.
'        ' 해당 파일의 수만큼 파일을 수신한다.
'        NextPross = False
'        pnlMsg.Caption = ""
'        FileProCount = 0
'        SendFlag = lauDaySaleData
'        Ret = InterNetFileRequest(lauDaySaleData)
'    End If
'    Call ButtonEnable(True)
'
'End Sub
'
'Private Sub cmdExit_Click()
'    Unload Me
'End Sub
'
''+------------------------------------------------------
''+ 2003/02/11 수정
''+
''+루틴설명      - 금액자료받기
''+  1. 일자의 내용을 본사에서 내려 받는다.
''+  2. 매일이 있을 경우 매일도 같이 내려 받느다.
''+
''+------------------------------------------------------
'Private Sub cmdPriceRecv_Click()
'    ' 이전 자료 받기 때문에 모두 해준다
'    g_AgencyCode = 가맹점정보.택코드
'    m_MstCode = 가맹점정보.지사코드
'
'    Call ButtonEnable(False)
'    pgbStatus.Value = 0
'
'    If SendMode = 1 Then
'        ' 모뎀
'        Ret = PriceData
'    Else
'        ' 인터넷
'        ' 먼저 파일 리스트를 받아 수신할 파일의 수를 구한다.
'        ' 해당 파일의 수만큼 파일을 수신한다.
'        NextPross = False
'        pnlMsg.Caption = ""
'        FileProCount = 0
'        SendFlag = lauPriceData
'        Ret = InterNetFileRequest(lauPriceData)
'    End If
'    Call ButtonEnable(True)
'
'End Sub
'
''+------------------------------------------------------
''+ 2003/02/11 수정
''+
''+루틴설명      - 수선자료받기
''+  1. 일자의 내용을 본사에서 내려 받는다.
''+  2. 매일이 있을 경우 매일도 같이 내려 받느다.
''+
''+------------------------------------------------------
'Private Sub cmdRepair_Click()
'    ' 이전 자료 받기 때문에 모두 해준다
'    g_AgencyCode = 가맹점정보.택코드
'    m_MstCode = 가맹점정보.지사코드
'
'    Call ButtonEnable(False)
'    pgbStatus.Value = 0
'
'    If SendMode = 1 Then
'        ' 모뎀
'        Ret = RepairData
'    Else
'        ' 인터넷
'        ' 먼저 파일 리스트를 받아 수신할 파일의 수를 구한다.
'        ' 해당 파일의 수만큼 파일을 수신한다.
'        NextPross = False
'        pnlMsg.Caption = ""
'        FileProCount = 0
'        SendFlag = lauRepairData
'        Ret = InterNetFileRequest(lauRepairData)
'    End If
'    Call ButtonEnable(True)
'
'End Sub
'
''+------------------------------------------------------
''+ 2003/02/11 수정
''+
''+루틴설명      - 할인자료받기
''+  1. 일자의 내용을 본사에서 내려 받는다.
''+  2. 매일이 있을 경우 매일도 같이 내려 받느다.
''+
''+------------------------------------------------------
'Private Sub cmdSaleRecv_Click()
'    ' 이전 자료 받기 때문에 모두 해준다
'    g_AgencyCode = 가맹점정보.택코드
'    m_MstCode = 가맹점정보.지사코드
'
'    Call ButtonEnable(False)
'    pgbStatus.Value = 0
'
'    If SendMode = 1 Then
'        ' 모뎀
'        Ret = SaleData
'    Else
'        ' 인터넷
'        ' 먼저 파일 리스트를 받아 수신할 파일의 수를 구한다.
'        ' 해당 파일의 수만큼 파일을 수신한다.
'        NextPross = False
'        pnlMsg.Caption = ""
'        FileProCount = 0
'        SendFlag = lauSaleData
'        Ret = InterNetFileRequest(lauSaleData)
'
'    End If
'    Call ButtonEnable(True)
'
'End Sub
'
'Private Sub cmdSave_Click()
'    Dim msg As String
'
'    If SendMode = 1 Then
'        msg = "전송작업경로 [기본값 : \\CleanAid\CleanData]" & vbLf
'        msg = msg & "수신작업경로 [기본값 : \\CleanAid\CleanRecv]" & vbLf
'        msg = msg & vbLf & vbLf & "[2005/06/20일 현재]"
'    ElseIf SendMode = 2 Then
'        msg = "전송작업경로 [" & M_CompnyMasterName & " 기본값 : web.clean-aid.co.kr ]" & vbLf
'        msg = msg & "수신작업경로 [체인점 기본값 : 체인점 IP 자동설정]" & vbLf
'        msg = msg & "MsgRemotePort=8607" & vbLf
'        msg = msg & "FileLocalPort=8629" & vbLf
'        msg = msg & "FileRemotePort=8627" & vbLf
'        msg = msg & vbLf & vbLf & "[2005/06/20일 현재]"
'    End If
'
'    MsgBox msg, vbInformation, "도움말"
'End Sub
'
'Private Function SendMileageData() As Boolean
'    Dim g_IpDate As String
'    Dim g_Count As Integer
'    Dim g_STag As String
'    Dim g_ETag As String
'    Dim g_Amount As Currency
'    Dim g_Jae As Integer
'    Dim g_Su As Integer
'    Dim g_Ban As Integer
'    Dim g_Ma As Integer
'    Dim g_SDate As String
'    Dim g_PAmt As Currency
'    Dim tmpStr As String
'    Dim Filename As String
'    Dim temp As String
'    Dim FHandle As Integer
'    SendMileageData = False
'
'    If 가맹점정보.마일리지여부 <> "Y" Then Exit Function
'
'    ' 일자를 Check한다.
'    If Not IsDate(txtDate.Text) Then
'        MsgBox "일자가 잘못 입력되었습니다.", vbExclamation, "확인"
'        Exit Function
'    End If
'
'    ' 현재일자를 g_IpDate에 넣는다.
'    g_IpDate = Format(txtDate.Text, "YYYY-MM-DD")
'
'    ' 마감 여부를 확인한다.
'    If Not Fun_일일마감여부(g_IpDate) Then Exit Function
'
'    On Error GoTo Err_Loop
'
'    pnlMsg.Caption = "마일리지자료 확인중..!"
'
'    ' 마일리지 파일 생성 "G20050314-002-1.DAT"
'    Filename = "G" & CStr(g_IpDate) & "-" & g_AgencyCode & "-1.DAT"
'
'    If UCase(Dir(App.Path & "\Internet\" & Filename)) = UCase(Filename) Then
'        Kill App.Path & "\Internet" & "\" & Filename
'    End If
'
'    FHandle = FreeFile
'    Open App.Path & "\Internet" & "\" & Filename _
'        For Output Lock Read Write As #FHandle
'
'    Print #FHandle, "마일리지스토리|발생일자|고객코드|발생마일리지|사용마일리지|삭제마일리지|보관증|전송여부|"
'    Print #FHandle, "마일리지현황|고객코드|총사용금액|마일리지|최종발생금액|발생누계|사용마일리지|최종거래일자|전송여부|"
'
'    '-----------------------------------------------------------------------------
'    '
'    '-----------------------------------------------------------------------------
'    Query = "SELECT * FROM TB_마일리지스토리 "
'    Query = Query & " WHERE 전송여부 = 'N'"
'    Query = Query & "    OR 전송일자 = '" & Format(Date, "YYYY-MM-DD") & "'"
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
'
'    'Set daoQD = daoDB.CreateQueryDef("", Query)
'    'Set daoRS = daoQD.OpenRecordset()
'
'    If Not ADORs.EOF Then
'        pnlMsg.Caption = ""
'
'        ADORs.MoveLast
'        pgbStatus.MAX = ADORs.RecordCount
'        pgbStatus.Min = 0
'        pgbStatus.Value = 0
'        pnlMsg.Caption = "마일리지자료 생성중..!" & Chr(13) & "총 " & pgbStatus.MAX & " 건"
'        ADORs.MoveFirst
'
'        Do While Not ADORs.EOF
'            If pgbStatus.Value < pgbStatus.MAX Then pgbStatus.Value = pgbStatus.Value + 1
'
'            tmpStr = "마일리지스토리" & "|"
'            tmpStr = tmpStr & ADORs!발생일자 & "|"
'            tmpStr = tmpStr & ADORs!고객코드 & "|"
'            tmpStr = tmpStr & CStr(ADORs!발생마일리지) & "|"
'            tmpStr = tmpStr & CStr(ADORs!사용마일리지) & "|"
'            tmpStr = tmpStr & CStr(ADORs!삭제마일리지) & "|"
'            tmpStr = tmpStr & CStr(ADORs!보관증) & "|"
'            tmpStr = tmpStr & ADORs!전송여부 & "|"
'
'            Print #FHandle, tmpStr
'
'            ADORs.MoveNext
'        Loop
'    End If
'    ADORs.Close
'    Set ADORs = Nothing
'
'    '-----------------------------------------------------------------------------
'    '
'    '-----------------------------------------------------------------------------
'    Query = "SELECT * FROM TB_마일리지현황 "
'    Query = Query & " WHERE 전송여부 = 'N'"
'    Query = Query & "    OR 전송일자 = '" & Format(Date, "YYYY-MM-DD") & "'"
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
'
'    'Set daoQD = daoDB.CreateQueryDef("", Query)
'    'Set daoRS = daoQD.OpenRecordset()
'
'    If Not ADORs.EOF Then
'        ADORs.MoveLast
'        pgbStatus.MAX = ADORs.RecordCount
'        pgbStatus.Min = 0
'        pgbStatus.Value = 0
'        pnlMsg.Caption = "마일리지자료 생성중..!" & Chr(13) & "총 " & pgbStatus.MAX & " 건"
'        ADORs.MoveFirst
'
'        Do While Not ADORs.EOF
'            If pgbStatus.Value < pgbStatus.MAX Then pgbStatus.Value = pgbStatus.Value + 1
'
'            tmpStr = "마일리지현황" & "|"
'            tmpStr = tmpStr & ADORs!고객코드 & "|"
'            tmpStr = tmpStr & CStr(ADORs!총사용금액) & "|"
'            tmpStr = tmpStr & CStr(ADORs!마일리지) & "|"
'            tmpStr = tmpStr & CStr(ADORs!최종발생금액) & "|"
'            tmpStr = tmpStr & CStr(ADORs!발생누계) & "|"
'            tmpStr = tmpStr & CStr(ADORs!사용마일리지) & "|"
'            tmpStr = tmpStr & ADORs!최종거래일자 & "|"
'            tmpStr = tmpStr & ADORs!전송여부 & "|"
'
'            Print #FHandle, tmpStr
'
'            ADORs.MoveNext
'        Loop
'    End If
'    ADORs.Close
'    Set ADORs = Nothing
'
'    Close #FHandle
'
'    pnlMsg.Caption = "마일리지자료 생성이 완료되었습니다." & Chr(13) & "총 " & pgbStatus.MAX & " 건"
'
'    ' 인터넷 전송 부분
'    PauseTime = 10                 ' 기간을 지정합니다.
'    Start = Timer                   ' 시작 시간을 지정합니다.
'
'    Do While FTC.State <> ftcReady
'        Finish = Timer              ' 종료 시간을 지정합니다.
'        TotalTime = Finish - Start  ' 전체 시간을 계산합니다.
'
'        If Timer > Start + PauseTime Then
'            panMsg.Caption = "전송 실패"
'            Exit Do
'        End If
'
'        DoEvents                    ' 다른 프로시저로 넘깁니다.
'    Loop
'
'    ' 본사에 파일을 전송한다.
'    SendFile = App.Path & "\Internet" & "\" & Filename
'    SendFlag = lauMileage
'    SendMileageData = InterNetSendFiles(SendFile)
'
'    Exit Function
'
'End_Loop:
'    Exit Function
'
'Err_Loop:
'
'    pnlMsg.Caption = "입고자료 전송이 취소되었습니다."
'
'    MsgBox "작업오류 입니다." & Chr(13) & _
'           "오류코드 : " & VBA.Err.Number & Chr(13) & _
'           "오류설명 : " & VBA.Err.Description, vbExclamation, "오류"
'    Resume End_Loop
'End Function
'
'Private Function InputData() As Boolean
'    Dim dblMil  As Double
'    Dim dblRatio As Double
'    Dim g_IpDate As String
'    Dim g_Count As Integer
'    Dim g_STag As String
'    Dim g_ETag As String
'    Dim g_Amount As Currency
'    Dim g_Jae As Integer
'    Dim g_Su As Integer
'    Dim g_Ban As Integer
'    Dim g_Ma As Integer
'    Dim g_SDate As String
'    Dim g_PAmt As Currency
'    Dim tmpStr As String
'    Dim Filename As String
'    Dim temp As String
'    Dim FHandle As Integer
'
'    InputData = False
'
'    ' 일자를 Check한다.
'    If Not IsDate(txtDate.Text) Then
'        MsgBox "일자가 잘못 입력되었습니다.", vbExclamation, "입력오류"
'        Exit Function
'    End If
'
'    ' 현재일자를 g_IpDate에 넣는다.
'    g_IpDate = Format(txtDate.Text, "YYYY-MM-DD")
'
'    ' 마감 여부를 확인한다.
'    If Not Fun_일일마감여부(g_IpDate) Then
'        MsgBox "일일마감을 하신후에 전송하세요..!", vbInformation, "자료없음"
'        Exit Function
'    End If
'
'    On Error GoTo Err_Loop
'
'    pnlMsg.Caption = "입고자료 확인중..!"
'
'    ' 입고파일를 생성한다.
'    ' 입고파일형식은 '현재날짜-매장코드-1.DAT'로 생성된다.
'    ' 메일파일를 생성한다.
'    Filename = CStr(g_IpDate) & "-" & g_AgencyCode & "-1.DAT"
'
'    If UCase(Dir(App.Path & "\Internet\" & Filename)) = UCase(Filename) Then
'        Kill App.Path & "\Internet" & "\" & Filename
'    End If
'
'    FHandle = 126 'FreeFile
'
'    Open App.Path & "\Internet" & "\" & Filename _
'        For Output Lock Read Write As #FHandle
'
'    '---------------------------------------------------------------
'    ' Text문서(현재날짜-매장코드-1.DAT)에 추가될 내역을 Select한다.
'    ' 입출고내역을 읽어온다.
'    '---------------------------------------------------------------
'    Query = "SELECT   A.접수일자 as IpDate, "
'    Query = Query & " A.고객코드 as CustCode, "
'    Query = Query & " A.의류코드 as GoodsCode, "
'    Query = Query & " A.택번호 as  TagNo, "
'    Query = Query & " A.색상 as Color, "
'    Query = Query & " A.내용 as Worked, "
'    Query = Query & " A.금액 as Amount, "
'    Query = Query & " A.상표 as Label, "
'    Query = Query & " A.결제여부 as Status, "
'    Query = Query & " B.성명 as CustName, "
'    Query = Query & " B.전화번호 as TelNo, "
'    Query = Query & " A.판매취소 as CancelSale "
'    Query = Query & " FROM TB_입출고 AS A LEFT OUTER JOIN TB_고객정보 AS B  ON A.고객코드 = B.고객코드 "
'    Query = Query & " WHERE A.접수일자 = '" & g_IpDate & "'"
'    ' 2003/03/02 본사 프로그램과 연동을 위하여 판매 취소 전송에서 제외
'    ' 체인점 프로그램에서 판매 취소한 택을 재 사용하기 때문에 이 기능은 사실상 무의미
'    Query = Query & "   AND (A.판매취소 IS NULL OR A.판매취소 <> 'Y') "
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
'
'    If ADORs.EOF Then
'        ADORs.Close
'        Set ADORs = Nothing
'
'        MsgBox "입고자료가 없습니다.", vbInformation, "자료없음"
'        pnlMsg.Caption = ""
'
'        GoTo End_Loop
'    End If
'
'    '----------------------------------------------------------------
'    ' 일일마감 CHECK
'    '----------------------------------------------------------------
'    Query = "SELECT * FROM TB_일일마감 WHERE 일자 = '" & g_IpDate & "'"
'    Set SUBRs = New ADODB.Recordset
'    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
'
'    If SUBRs.BOF Or SUBRs.EOF Then
'       SUBRs.Close
'       Set SUBRs = Nothing
'
'       MsgBox "일일마감을 하신후에 전송하세요..!", vbInformation, "자료없음"
'       pnlMsg.Caption = ""
'       GoTo End_Loop
'    End If
'
'    pnlMsg.Caption = "입고자료 생성중..!" & Chr(13) & "총 " & pgbStatus.MAX & " 건"
'    dblRatio = 1 - (CDbl(가맹점정보.비율)) / 100
'    dblMil = IIf(IsNull(SUBRs.Fields("사용마일리지")) = True, 0, SUBRs.Fields("사용마일리지"))
'
'    tmpStr = "일일마감|"
'    temp = "1234"
'    RSet temp = SUBRs!총점수
'
'    tmpStr = tmpStr & temp & "|"
'    temp = "1234"
'    RSet temp = SUBRs!반품수량
'
'    tmpStr = tmpStr & temp & "|"
'    temp = "1234"
'    RSet temp = SUBRs!재세탁수량
'
'    tmpStr = tmpStr & temp & "|"
'    temp = "1234"
'    RSet temp = SUBRs!수선수량
'
'    tmpStr = tmpStr & temp & "|"
'    temp = "12345678"
'    RSet temp = SUBRs!총매출액
'
'    tmpStr = tmpStr & temp & "|"
'    temp = "12345678"
'    RSet temp = SUBRs.Fields("본사금액") - (dblMil * dblRatio)
'
'    tmpStr = tmpStr & temp & "|"
'    temp = "12345678"
'    RSet temp = SUBRs.Fields("가맹점금액") - (dblMil * (1 - dblRatio))
'
'    tmpStr = tmpStr & temp & "|"
'    tmpStr = tmpStr & SUBRs!판매구분 & "|" & SUBRs!시작택 & "|" & SUBRs!종료택 & "|"
'
'    ' pds2004 2007-05-28 카드금액, 건수 추가
'    temp = "12345678"
'    If IsNull(SUBRs!카드금액) = True Then
'        RSet temp = "0"
'    Else
'        RSet temp = SUBRs!카드금액
'    End If
'
'    tmpStr = tmpStr & temp & "|"
'
'    temp = "12345678"
'    If IsNull(SUBRs!카드건수) = True Then
'        RSet temp = "0"
'    Else
'        RSet temp = SUBRs!카드건수
'    End If
'    tmpStr = tmpStr & temp & "|"
'
'    Print #FHandle, tmpStr
'    SUBRs.Close
'
'    ADORs.MoveLast
'
'    pgbStatus.MAX = ADORs.RecordCount
'    pgbStatus.Min = 0
'    pgbStatus.Value = 0
'
'    ADORs.MoveFirst
'
'    g_STag = Left(ADORs!TagNo, 1) & Right(ADORs!TagNo, 3)
'    g_ETag = Left(ADORs!TagNo, 1) & Right(ADORs!TagNo, 3)
'
'    Do While Not ADORs.EOF
'        If pgbStatus.Value < pgbStatus.MAX Then
'            pgbStatus.Value = pgbStatus.Value + 1
'        End If
'
'        tmpStr = ADORs!Ipdate & "|"
'        tmpStr = tmpStr & g_AgencyCode & "|"
'        tmpStr = tmpStr & Left(ADORs!TagNo, 1) & Right(ADORs!TagNo, 3) & "|"
'        tmpStr = tmpStr & ADORs!CustCode & "|"
'        tmpStr = tmpStr & ADORs!CustName & "|"
'        tmpStr = tmpStr & ADORs!TelNo & "|"
'        tmpStr = tmpStr & ADORs!GoodsCode & "|"
'        tmpStr = tmpStr & ADORs!Color & "|"
'        tmpStr = tmpStr & ADORs!Worked & "|"
'        tmpStr = tmpStr & Val(ADORs!Amount)
'        tmpStr = tmpStr & "|"
'        tmpStr = tmpStr & ADORs!Label & "|"
'        If ADORs!Status = "완불" Then
'            tmpStr = tmpStr & "1" & "|"
'        Else
'            tmpStr = tmpStr & "0" & "|"
'        End If
'
'        tmpStr = tmpStr & "|"
'        tmpStr = tmpStr & "|"
'        tmpStr = tmpStr & "" & "|" 'ADORs!CancelSale & "|"
'        '20091027일 수정 재새탁이나 다른 기호가 들어갈경우 문제가됨
'        ' 본사에서 출고 명세서에 찍히지 않는문제 점이 있음
'
'        Print #FHandle, tmpStr
'
'        ADORs.MoveNext
'    Loop
'    ADORs.Close
'    Set ADORs = Nothing
'
'    Close #FHandle
'
'    pnlMsg.Caption = "입고자료 생성이 완료되었습니다." & Chr(13) & "총 " & pgbStatus.MAX & " 건"
'
'
'    DoEvents
'    PauseTime = 10                 ' 기간을 지정합니다.
'    Start = Timer                   ' 시작 시간을 지정합니다.
'
'    Do While FTC.State <> ftcReady
'        Finish = Timer              ' 종료 시간을 지정합니다.
'        TotalTime = Finish - Start  ' 전체 시간을 계산합니다.
'
'        If Timer > Start + PauseTime Then
'            InputData = False
'            panMsg.Caption = "전송 실패"
'            Exit Function
'        End If
'        DoEvents                    ' 다른 프로시저로 넘깁니다.
'    Loop
'
'    ' 본사에 파일을 전송한다.
'    SendFile = App.Path & "\Internet" & "\" & Filename
'    SendFlag = lauInput
'    InputData = InterNetSendFiles(SendFile)
'
'    Exit Function
'
'End_Loop:
'    Close #FHandle
'
'
'    Exit Function
'
'Err_Loop:
'    Close #FHandle
'
'    pnlMsg.Caption = "입고자료 전송이 취소되었습니다."
'
'    MsgBox "작업오류 입니다." & Chr(13) & _
'           "오류코드 : " & VBA.Err.Number & Chr(13) & _
'           "오류설명 : " & VBA.Err.Description, vbExclamation, "오류"
'    'Resume
'    Resume End_Loop
'End Function
''
'
'Private Function QN_Data() As Boolean
'    Dim dblMil  As Double
'    Dim dblRatio As Double
'    Dim g_IpDate As String
'    Dim g_Count As Integer
'    Dim g_STag As String
'    Dim g_ETag As String
'    Dim g_Amount As Currency
'    Dim g_Jae As Integer
'    Dim g_Su As Integer
'    Dim g_Ban As Integer
'    Dim g_Ma As Integer
'    Dim g_SDate As String
'    Dim g_PAmt As Currency
'    Dim tmpStr As String
'    Dim Filename As String
'    Dim temp As String
'    Dim FHandle As Integer
'
'    QN_Data = False
'
'    ' 일자를 Check한다.
'    If Not IsDate(txtDate.Text) Then
'        MsgBox "일자가 잘못 입력되었습니다.", vbExclamation, "입력오류"
'        Exit Function
'    End If
'
'    ' 현재일자를 g_IpDate에 넣는다.
'    g_IpDate = Format(txtDate.Text, "YYYY-MM-DD")
'
'    ' 마감 여부를 확인한다.
'    If Not Fun_일일마감여부(g_IpDate) Then
'        Exit Function
'    End If
'
'    On Error GoTo Err_Loop
'
'    pnlMsg.Caption = "보관 서비스 자료 확인중..!"
'
'    GoSub SUB_LIST1 ' 보관 리스트 자료를 생성한다.
'
'    GoSub SUB_LIST2 ' 보관 상품 리스트 자료를 생성한다.
'
'    GoSub SUB_LIST3 ' 보관 하자 리스트 자료를 생성한다.
'
'    Close #FHandle
'
'    pnlMsg.Caption = "보관 서비스 자료 생성이 완료되었습니다." & Chr(13) & "총 " & pgbStatus.MAX & " 건"
'
'    ADORs.Close
'    Set ADORs = Nothing
'
'    DoEvents
'    PauseTime = 10                 ' 기간을 지정합니다.
'    Start = Timer                   ' 시작 시간을 지정합니다.
'
'    Do While FTC.State <> ftcReady
'        Finish = Timer              ' 종료 시간을 지정합니다.
'        TotalTime = Finish - Start  ' 전체 시간을 계산합니다.
'
'        If Timer > Start + PauseTime Then
'            QN_Data = False
'            panMsg.Caption = "전송 실패"
'            Exit Function
'        End If
'
'        DoEvents                    ' 다른 프로시저로 넘깁니다.
'    Loop
'
'    ' 본사에 파일을 전송한다.
'    SendFile = App.Path & "\Internet" & "\" & Filename
'    SendFlag = lauQNData
'    QN_Data = InterNetSendFiles(SendFile)
'
'    Exit Function
'
'End_Loop:
'    Exit Function
'
'Err_Loop:
'    If FHandle > 0 Then Close #FHandle
'
'    pnlMsg.Caption = "입고자료 전송이 취소되었습니다."
'
'    MsgBox "작업오류 입니다." & Chr(13) & _
'           "오류코드 : " & VBA.Err.Number & Chr(13) & _
'           "오류설명 : " & VBA.Err.Description, vbExclamation, "오류"
'    Resume End_Loop
'
'' 보관 리스트 자료를 생성한다.
'SUB_LIST1:
'    '-------------------------------------------------------------------
'    '
'    '-------------------------------------------------------------------
'    Query = "SELECT KeyCode, MemRecord, InputNumber, InputDate, InputID, InputName, "
'    Query = Query & " EMail, UserCode, UserNumber, 가맹점코드, SaleGubunCode, SaleEndDate, "
'    Query = Query & " Price, DevTimeCode, ItemCount "
'    Query = Query & "  FROM TB_보관리스트 "
'    Query = Query & " WHERE SUBSTRING(InputDate,1,10) = '" & g_IpDate & "' "
'    Query = Query & "   AND StatsFlag <> 'C' "
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
'
'    If ADORs.EOF Then
'        pnlMsg.Caption = "보관 서비스 자료가 없습니다."
'
'        GoTo End_Loop
'    End If
'
'    ADORs.MoveLast
'
'    pgbStatus.MAX = ADORs.RecordCount
'    pgbStatus.Min = 0
'    pgbStatus.Value = 0
'    ADORs.MoveFirst
'    pnlMsg.Caption = "보관 서비스 자료 생성중..!" & Chr(13) & "총 " & pgbStatus.MAX & " 건"
'
'    ' 입고파일를 생성한다.
'    ' 입고파일형식은 '현재날짜-매장코드-1.DAT'로 생성된다.
'    ' 메일파일를 생성한다.
'    Filename = "Q" & CStr(g_IpDate) & "-" & g_AgencyCode & "-1.DAT"
'
'    If UCase(Dir(App.Path & "\Internet\" & Filename)) = UCase(Filename) Then
'        Kill App.Path & "\Internet" & "\" & Filename
'    End If
'
'    FHandle = FreeFile
'    Open App.Path & "\Internet" & "\" & Filename For Output Lock Read Write As #FHandle
'
'    Do While Not ADORs.EOF
'        If pgbStatus.Value < pgbStatus.MAX Then
'            pgbStatus.Value = pgbStatus.Value + 1
'        End If
'
'        tmpStr = "보관리스트" & "|"
'        tmpStr = tmpStr & ADORs!KeyCode & "|"
'        tmpStr = tmpStr & ADORs!MemRecord & "|"
'        tmpStr = tmpStr & ADORs!InputNumber & "|"
'        tmpStr = tmpStr & ADORs!InputDate & "|"
'        tmpStr = tmpStr & ADORs!InputID & "|"
'        tmpStr = tmpStr & ADORs!InputName & "|"
'        tmpStr = tmpStr & ADORs!EMail & "|"
'        tmpStr = tmpStr & ADORs!UserCode & "|"
'        tmpStr = tmpStr & ADORs!UserNumber & "|"
'        tmpStr = tmpStr & ADORs!가맹점코드 & "|"
'        tmpStr = tmpStr & ADORs!SaleGubunCode & "|"
'        tmpStr = tmpStr & ADORs!SaleEndDate & "|"
'        tmpStr = tmpStr & ADORs!Price & "|"
'        tmpStr = tmpStr & ADORs!DevTimeCode & "|"
'        tmpStr = tmpStr & ADORs!ItemCount & "|"
'
'        Print #FHandle, tmpStr
'
'        ADORs.MoveNext
'    Loop
'    Return
'
'' 보관 상품 리스트 자료를 생성한다.
'SUB_LIST2:
'
'    '-----------------------------------------------------------------
'    '
'    '-----------------------------------------------------------------
'    Query = "SELECT KeyCode, ItemRecord, ItemIndex, InputDate, TAG, GoodsCode, SizeGubun, "
'    Query = Query & " SizeCode, Color, BrandName, BuyPrice, BuyDate, ASGubun, BleCount "
'    Query = Query & "  FROM TB_보관상품리스트 "
'    Query = Query & " WHERE InputDate = '" & g_IpDate & "' "
'    Query = Query & "   AND StatsFlag <> 'C' "
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
'
'    If ADORs.EOF Then
'        pnlMsg.Caption = "보관 상품 서비스 자료가 없습니다."
'        GoTo End_Loop
'    End If
'
'    ADORs.MoveLast
'
'    pgbStatus.MAX = ADORs.RecordCount
'    pgbStatus.Min = 0
'    pgbStatus.Value = 0
'    ADORs.MoveFirst
'    pnlMsg.Caption = "보관 상품 서비스 자료 생성중..!" & Chr(13) & "총 " & pgbStatus.MAX & " 건"
'
'    Do While Not ADORs.EOF
'        If pgbStatus.Value < pgbStatus.MAX Then
'            pgbStatus.Value = pgbStatus.Value + 1
'        End If
'
'        tmpStr = "보관상품리스트" & "|"
'        tmpStr = tmpStr & ADORs!KeyCode & "|"
'        tmpStr = tmpStr & ADORs!ItemRecord & "|"
'        tmpStr = tmpStr & ADORs!ItemIndex & "|"
'        tmpStr = tmpStr & ADORs!InputDate & "|"
'        tmpStr = tmpStr & ADORs!Tag & "|"
'        tmpStr = tmpStr & ADORs!GoodsCode & "|"
'        tmpStr = tmpStr & ADORs!SizeGubun & "|"
'        tmpStr = tmpStr & ADORs!SizeCode & "|"
'        tmpStr = tmpStr & ADORs!Color & "|"
'        tmpStr = tmpStr & ADORs!BrandName & "|"
'        tmpStr = tmpStr & ADORs!BuyPrice & "|"
'        tmpStr = tmpStr & ADORs!BuyDate & "|"
'        tmpStr = tmpStr & ADORs!ASGubun & "|"
'        tmpStr = tmpStr & ADORs!BleCount & "|"
'
'        Print #FHandle, tmpStr
'
'        ADORs.MoveNext
'    Loop
'    Return
'
'' 보관 하자리스트 자료를 생성한다.
'SUB_LIST3:
'    '-----------------------------------------------------------------
'    '
'    '-----------------------------------------------------------------
'    Query = "SELECT KeyCode, InputDate, ItemIndex, ItemCount, ItemRemark "
'    Query = Query & "  FROM TB_보관하자리스트 "
'    Query = Query & " WHERE InputDate = '" & g_IpDate & "' "
'    Query = Query & "   AND StatsFlag <> 'C' "
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
'
'    If ADORs.EOF Then
'        pnlMsg.Caption = "보관 하자 서비스 자료가 없습니다."
'        Return
'    End If
'
'    ADORs.MoveLast
'
'    pgbStatus.MAX = ADORs.RecordCount
'    pgbStatus.Min = 0
'    pgbStatus.Value = 0
'    ADORs.MoveFirst
'    pnlMsg.Caption = "보관 하자 서비스 자료 생성중..!" & Chr(13) & "총 " & pgbStatus.MAX & " 건"
'
'    Do While Not ADORs.EOF
'        If pgbStatus.Value < pgbStatus.MAX Then
'            pgbStatus.Value = pgbStatus.Value + 1
'        End If
'
'        tmpStr = "보관하자리스트" & "|"
'        tmpStr = tmpStr & ADORs!KeyCode & "|"
'        tmpStr = tmpStr & ADORs!InputDate & "|"
'        tmpStr = tmpStr & ADORs!ItemIndex & "|"
'        tmpStr = tmpStr & ADORs!ItemCount & "|"
'        tmpStr = tmpStr & ADORs!ItemRemark & "|"
'
'        Print #FHandle, tmpStr
'
'        ADORs.MoveNext
'    Loop
'
'    Return
'
'End Function
'
''+------------------------------------------------------
''+ 2002/00/00
''+
''+루틴설명      - 출고자료받기
''+
''+ 전달 형태
''+  파일명 : Down일자.dat                           ex. Down20030222.Dat
''+  내  용 : 일보일자 택번호 (1출,2반,E오류 여부)   ex. 20030218 1263 2
''+           인터넷일경우 다른곳에서 이미 파일을 다운로드 한 상태이다.
''+------------------------------------------------------
'Private Function ChulgoData() As Boolean
'
'    Dim g_Count         As Long
'    Dim l_TotalCount    As Long
'    Dim Filename        As String
'    Dim St              As String
'    Dim FHandel         As Integer
'    Dim FHandel2        As Integer
'
'    pnlMsg.Caption = "출고자료 확인중..!"
'
'    On Error GoTo Err_Loop
'
'    l_TotalCount = 0
'
'    FHandel = FreeFile
'    Open App.Path & "\RecvData\Chulgo.Dat" For Output As #FHandel
'    Close #FHandel
'    DoEvents
'
'    Filename = Dir(strSendPath & "\Down" & g_AgencyCode & "*")
'
'    If Filename = "" Then
'        pnlMsg.Caption = ""
'        MsgBox "출고자료가 없습니다." & Space(10), vbInformation, "확인"
'        Exit Function
'    End If
'
'    Do
'        ' 모뎀일 경우에만 복사한다
'        ' 인터넷일경우 다른곳에서 복사한다.
'        If SendMode = 1 Then FileCopy strSendPath & "\" & Filename, App.Path & "\RecvData\" & Filename
'
'        '파일의 총 가운터 수를 구한다.
'        g_Count = 0
'        FHandel = FreeFile
'        Open App.Path & "\RecvData\" & Filename For Input As #FHandel
'        Do While Not EOF(1)
'            Line Input #FHandel, St
'            g_Count = g_Count + 1
'            l_TotalCount = l_TotalCount + 1
'        Loop
'        DoEvents
'        Close #FHandel
'
'        Filename = Dir
'    Loop While Not Filename = ""
'
'
'    pgbStatus.MAX = l_TotalCount * 2 ' 2가지 업무로 나누어지기 때문에 * 2를 해준다.
'    pgbStatus.Min = 0
'    pgbStatus.Value = 0
'
'
'    Filename = Dir(strSendPath & "\Down" & g_AgencyCode & "*")
'
'    Do
'
'        FHandel = FreeFile
'        Open App.Path & "\RecvData\" & Filename For Input As #FHandel
'
'        FHandel2 = FreeFile
'        Open App.Path & "\RecvData\Chulgo.Dat" For Append As #FHandel2
'
'        pnlMsg.Caption = "[" & Format(Mid(Filename, 8, 8), "YYYY-MM-DD") & "일]" & _
'                     " 본사 출고 자료를 처리 중입니다." & vbLf & "잠시만 기다려 주십시요."
'        DoEvents
'
'
'        Do While Not EOF(FHandel)
'            Line Input #FHandel, St
'            Print #FHandel2, St
'
'            If pgbStatus.MAX > pgbStatus.Value Then pgbStatus.Value = pgbStatus.Value + 1
'
'            St_Data(0) = Format(Date, "YYYY-MM-DD")   ' 작업일자
'            St_Data(1) = Mid(Filename, 8, 8)        ' 본사출고일
'            St_Data(2) = Mid(St, 2, 8)              ' 입고일자
'            St_Data(3) = Mid(St, 11, 4)             ' 택번호
'            St_Data(4) = Mid(St, 16, 1)             ' 구분
'
'            Query = "SELECT * FROM TB_본사입고  "
'            Query = Query & " WHERE 작업일자   =  '" & St_Data(0) & "' "
'            Query = Query & "   AND 본사출고일 =  '" & St_Data(1) & "'"
'            Query = Query & "   AND 입고일자   =  '" & St_Data(2) & "'"
'            Query = Query & "   AND 택번호     =  '" & St_Data(3) & "'"
'            Query = Query & "   AND 구분       =  '" & St_Data(4) & "'"
'            Set ADORs = New ADODB.Recordset
'            ADORs.Open Query, ADOCon, adOpenDynamic, adLockOptimistic
'
'            'Set GS_RS = Nothing
'            'Set GS_RS = daoDB.OpenRecordset(Query)
'
'            If ADORs.RecordCount > 0 Then
'                '
'            Else
''                Query = "INSERT INTO TB_본사입고(작업일자,본사출고일,입고일자,택번호,구분) "
''                Query = Query & "VALUES('" & St_Data(0) & "','" & St_Data(1) & "','" & St_Data(2) & "','" & St_Data(3) & "','" & St_Data(4) & "')"
''                daoDB.Execute Query
''
'                ADORs.AddNew
'                ADORs.Fields("작업일자") = St_Data(0)
'                ADORs.Fields("본사출고일") = St_Data(1)
'                ADORs.Fields("입고일자") = St_Data(2)
'                ADORs.Fields("택번호") = St_Data(3)
'                ADORs.Fields("구분") = St_Data(4)
'                ADORs.Update
'            End If
'
'            ADORs.Close
'        Loop
'
'        Close
'
'        Filename = Dir
'    Loop While Not Filename = ""
'
'    'Workspaces(0).BeginTrans
'
'    pnlMsg.Caption = "입출고 자료에 업데이트중.....!" & Chr(13) & "총 " & l_TotalCount & " 건"
'
'    FHandel = FreeFile
'    Open App.Path & "\RecvData\Chulgo.Dat" For Input As #FHandel
'
'    ' pds2004 수정 2007-06-10 문자메시지 관련 내용
'    'Set daoQD = daoDB.CreateQueryDef("", "UPDATE TB_입출고 SET 본출 = ChulGu, 본출일자 = ChulDate, 본출입고구분 =  ChulStats WHERE 택번호 = TagNo ") ' 2003.01.16일 수정 AND 접수일자 = IpDate")
'
'    ' 택번호가 1회전시 이전 택은 무조건 출고 한것으로 인정한다.
'
'    Do While Not EOF(1)
'        If pgbStatus.Value < pgbStatus.MAX Then
'            pgbStatus.Value = pgbStatus.Value + 1
'        End If
'
'        Line Input #FHandel, St
'
'''        If Right(St, 1) = "2" Then
'''            daoQD("ChulGu") = "出"
'''        ElseIf Right(St, 1) = "3" Then
'''            daoQD("ChulGu") = "反"
'''        Else
'''            daoQD("ChulGu") = "E"
'''        End If
'''
'''        ' pds2004 수정 2007-06-10 문자메시지 관련 내용
'''        daoQD("ChulDate") = Format(Date, "YYYY-MM-DD")
'''        daoQD("ChulStats") = "자동"
'''
'''' 2003.01.16        daoQD("IpDate") = Mid(St, 2, 8)
'''        daoQD("TagNo") = Mid(St, 11, 1) & "-" & Mid(St, 12, 3)
'''        daoQD.Execute
'
'        '-----------------------------------------------------------------------------
'        '
'        '-----------------------------------------------------------------------------
'        Query = "UPDATE TB_입출고 SET "
'
'        If Right(St, 1) = "2" Then
'            Query = Query & " 본출         = '出'"
'        ElseIf Right(St, 1) = "3" Then
'            Query = Query & " 본출         = '反'"
'        Else
'            Query = Query & " 본출         = 'E'"
'        End If
'
'        Query = Query & "        , 본출일자     = '" & Format(Date, "YYYY-MM-DD") & "'"
'        Query = Query & "        , 본출입고구분 = '자동'"
'        Query = Query & " WHERE 택번호 = '" & Mid(St, 11, 1) & "-" & Mid(St, 12, 3) & "'"
'        ADOCon.Execute Query
'    Loop
'
'    'Workspaces(0).CommitTrans
'
'    pnlMsg.Caption = "출고자료 수신이 완료되었습니다." & Chr(13) & "총 " & l_TotalCount & " 건"
'    ChulgoData = True
'
'End_Loop:
'    Close
'    Kill App.Path & "\RecvData\*.*"
'    Exit Function
'
'Err_Loop:
''    Workspaces(0).Rollback
'    pnlMsg.Caption = "출고자료 수신이 취소되었습니다."
'
'    MsgBox "작업오류 입니다." & Chr(13) & _
'           "오류코드 : " & VBA.Err.Number & Chr(13) & _
'           "오류설명 : " & VBA.Err.Description, vbExclamation, "오류"
'Resume
'    Resume End_Loop
'End Function
'
'Private Function SaleData() As Boolean
'    Dim g_SaleDate As String
'    Dim g_Count As Integer
'    Dim Filename As String
'    Dim St As String
'
'    If Not IsDate(txtDate.Text) Then
'        MsgBox "일자가 잘못 입력되었습니다.", vbExclamation, "입력오류"
'        Exit Function
'    End If
'
'    g_SaleDate = Format(txtDate.Text, "YYYY-MM-DD")
'
'    pnlMsg.Caption = "할인자료 확인중..!"
'
'    On Error GoTo Err_Loop
'
'    Filename = Dir(strSendPath & "\Sale" & g_AgencyCode & ".DAT")
'
'    If Filename = "" Then
'        pnlMsg.Caption = ""
'        MsgBox "할인자료가 없습니다.", vbInformation, "확인"
'        Exit Function
'    Else
'        ' 모뎀일 경우에만 복사한다
'        ' 인터넷일경우 다른곳에서 복사한다.
'        If SendMode = 1 Then FileCopy strSendPath & "\" & Filename, App.Path & "\RecvData\" & Filename
'
'        Open App.Path & "\RecvData\" & Filename For Input As #1
'        Open App.Path & "\RecvData\Sale.Dat" For Output As #2
'
'        Do While Not EOF(1)
'            Line Input #1, St
'            Print #2, St
'            g_Count = g_Count + 1
'        Loop
'
'        Close
'
'        pgbStatus.MAX = g_Count
'        pgbStatus.Min = 0
'        pgbStatus.Value = 0
'    End If
'
'    Workspaces(0).BeginTrans
'
'    pnlMsg.Caption = "기존 자료를 삭제중..!"
'    'daoDB.Execute "DELETE FROM TB_할인정보"
'    ADOCon.Execute "DELETE FROM TB_할인정보"
'
'    pnlMsg.Caption = "할인자료 수신중..!" & Chr(13) & "총 " & g_Count & " 건"
'
'    Open App.Path & "\RecvData\Sale.Dat" For Input As #1
'
'    'Set daoQD = daoDB.CreateQueryDef("", "INSERT INTO TB_할인정보 VALUES (시작일, 종료일, 의류코드, 의류명, 금액, 비율, 출력순번)")
'
'    Do While Not EOF(1)
'        If pgbStatus.Value < pgbStatus.MAX Then
'            pgbStatus.Value = pgbStatus.Value + 1
'        End If
'
'        Line Input #1, St
'
'''        daoQD("시작일") = Mid(St, 2, 8)
'''        daoQD("종료일") = Mid(St, 11, 8)
'''        daoQD("의류코드") = Mid(St, 20, 3)
'''        daoQD("의류명") = Mid(St, 36)
'''        daoQD("금액") = Mid(St, 24, 8)
'''        daoQD("비율") = Mid(St, 33, 1)
'''        daoQD("출력순번") = " "
'''        daoQD.Execute
'
'        Query = "INSERT INTO TB_할인정보 VALUES ("
'        Query = Query & "  '" & Mid(St, 2, 8) & "'"
'        Query = Query & ", '" & Mid(St, 11, 8) & "'"
'        Query = Query & ", '" & Mid(St, 20, 3) & "'"
'        Query = Query & ", '" & Mid(St, 36) & "'"
'        Query = Query & ", '" & Mid(St, 24, 8) & "'"
'        Query = Query & ", '" & Mid(St, 33, 1) & "'"
'        Query = Query & ", ' ')"
'        ADOCon.Execute Query
'    Loop
'
'    pnlMsg.Caption = "할인자료 수신이 완료되었습니다." & Chr(13) & "총 " & g_Count & " 건"
'    SaleData = True
'
'    'daoQD.Close
'
'    Workspaces(0).CommitTrans
'
'End_Loop:
'    Close
'
'    On Error GoTo ERR_FILE
'    If Dir(App.Path & "\RecvData\*.*") <> "" Then
'        Kill App.Path & "\RecvData\*.*"
'    End If
'    Exit Function
'
'Err_Loop:
'    Workspaces(0).ROLLBACK
'    pnlMsg.Caption = "할인자료 수신이 취소되었습니다."
'
'    MsgBox "작업오류 입니다." & Chr(13) & _
'           "오류코드 : " & VBA.Err.Number & Chr(13) & _
'           "오류설명 : " & VBA.Err.Description, vbExclamation, "오류"
'    Resume End_Loop
'
'ERR_FILE:
'    pnlMsg.Caption = "파일삭제중 오류가 밸생했습니다."
'    MsgBox "작업오류 입니다." & Chr(13) & _
'           "오류코드 : " & VBA.Err.Number & Chr(13) & _
'           "오류설명 : " & VBA.Err.Description, vbExclamation, "오류"
'    Resume Next
'End Function
'
'
'' 수정 안한내용임 수정하여 사용
'Private Function SendData_세탁환불() As Boolean
'    Dim g_SaleDate As String
'    Dim g_Count As Integer
'    Dim Filename As String
'    Dim St As String
'
'    If Not IsDate(txtDate.Text) Then
'        MsgBox "일자가 잘못 입력되었습니다.", vbExclamation, "입력오류"
'        Exit Function
'    End If
'    g_SaleDate = Format(txtDate.Text, "YYYY-MM-DD")
'
'    pnlMsg.Caption = "세탁 환불 자료 확인중..!"
'
'    On Error GoTo Err_Loop
'
'    Filename = Dir(strSendPath & "\Sale" & g_AgencyCode & ".DAT")
'
'    If Filename = "" Then
'        pnlMsg.Caption = ""
'        MsgBox "할인자료가 없습니다.", vbInformation, "확인"
'        Exit Function
'    Else
'        ' 모뎀일 경우에만 복사한다
'        ' 인터넷일경우 다른곳에서 복사한다.
'        If SendMode = 1 Then FileCopy strSendPath & "\" & Filename, App.Path & "\RecvData\" & Filename
'
'        Open App.Path & "\RecvData\" & Filename For Input As #1
'        Open App.Path & "\RecvData\Sale.Dat" For Output As #2
'
'        Do While Not EOF(1)
'            Line Input #1, St
'            Print #2, St
'            g_Count = g_Count + 1
'        Loop
'
'        Close
'
'        pgbStatus.MAX = g_Count
'        pgbStatus.Min = 0
'        pgbStatus.Value = 0
'    End If
'
'    'Workspaces(0).BeginTrans
'
'    pnlMsg.Caption = "기존 자료를 삭제중..!"
'    'daoDB.Execute "DELETE FROM TB_할인정보"
'    ADOCon.Execute "DELETE FROM TB_할인정보"
'
'    pnlMsg.Caption = "할인자료 수신중..!" & Chr(13) & "총 " & g_Count & " 건"
'
'    Open App.Path & "\RecvData\Sale.Dat" For Input As #1
'
'    'Set daoQD = daoDB.CreateQueryDef("", "INSERT INTO TB_할인정보 VALUES (시작일, 종료일, 의류코드, 의류명, 금액, 비율)")
'
'    Do While Not EOF(1)
'        If pgbStatus.Value < pgbStatus.MAX Then
'            pgbStatus.Value = pgbStatus.Value + 1
'        End If
'
'        Line Input #1, St
'
'''        daoQD("시작일") = Mid(St, 2, 8)
'''        daoQD("종료일") = Mid(St, 11, 8)
'''        daoQD("의류코드") = Mid(St, 20, 3)
'''        daoQD("의류명") = Mid(St, 36)
'''        daoQD("금액") = Mid(St, 24, 8)
'''        daoQD("비율") = Mid(St, 33, 1)
'''        daoQD.Execute
'
'        Query = "INSERT INTO TB_할인정보 VALUES ("
'        Query = Query & "  '" & Mid(St, 2, 8) & "'"
'        Query = Query & ", '" & Mid(St, 11, 8) & "'"
'        Query = Query & ", '" & Mid(St, 20, 3) & "'"
'        Query = Query & ", '" & Mid(St, 36) & "'"
'        Query = Query & ", '" & Mid(St, 24, 8) & "'"
'        Query = Query & ", '" & Mid(St, 33, 1) & "')"
'        ADOCon.Execute Query
'    Loop
'
'    pnlMsg.Caption = "할인자료 수신이 완료되었습니다." & Chr(13) & "총 " & g_Count & " 건"
'    SendData_세탁환불 = True
'
'    'daoQD.Close
'    'Workspaces(0).CommitTrans
'
'End_Loop:
'    Close
'
'    On Error GoTo ERR_FILE
'
'    If Dir(App.Path & "\RecvData\*.*") <> "" Then
'        Kill App.Path & "\RecvData\*.*"
'    End If
'    Exit Function
'
'Err_Loop:
'    'Workspaces(0).ROLLBACK
'    pnlMsg.Caption = "할인자료 수신이 취소되었습니다."
'
'    MsgBox "작업오류 입니다." & Chr(13) & _
'           "오류코드 : " & VBA.Err.Number & Chr(13) & _
'           "오류설명 : " & VBA.Err.Description, vbExclamation, "오류"
'    Resume End_Loop
'
'ERR_FILE:
'    pnlMsg.Caption = "파일삭제중 오류가 밸생했습니다."
'    MsgBox "작업오류 입니다." & Chr(13) & _
'           "오류코드 : " & VBA.Err.Number & Chr(13) & _
'           "오류설명 : " & VBA.Err.Description, vbExclamation, "오류"
'    Resume Next
'End Function
'
'Private Function PriceData() As Boolean
'    Dim g_SaleDate As String
'    Dim g_Count As Integer
'    Dim Filename As String
'    Dim St As String
'
'    If Not IsDate(txtDate.Text) Then
'        MsgBox "일자가 잘못 입력되었습니다.", vbExclamation, "입력오류"
'        Exit Function
'    End If
'
'    g_SaleDate = Format(txtDate.Text, "YYYY-MM-DD")
'
'    pnlMsg.Caption = "금액자료 확인중..!"
'
'    On Error GoTo Err_Loop
'
'    Filename = Dir(strSendPath & "\????????" & g_AgencyCode & ".DAT")
'
'    If Filename = "" Then
'        pnlMsg.Caption = ""
'        MsgBox "금액자료가 없습니다.", vbInformation, "확인"
'        Exit Function
'    Else
'        pgbStatus.MAX = 400
'        pgbStatus.Min = 0
'        pgbStatus.Value = 0
'
'        pnlMsg.Caption = "금액자료 수신중..!"
'
'        ' 모뎀일 경우에만 복사한다
'        ' 인터넷일경우 다른곳에서 복사한다.
'        If SendMode = 1 Then FileCopy strSendPath & "\" & Filename, App.Path & "\RecvData\" & Filename
'
'        Open App.Path & "\RecvData\" & Filename For Input As #1
'        Open App.Path & "\BackData\" & Filename For Output As #2
'
'        Do While Not EOF(1)
'            Line Input #1, St
'            Print #2, St
'            g_Count = g_Count + 1
'
'            If pgbStatus.Value = 400 Then
'               pgbStatus.Value = 0
'            End If
'
'            pgbStatus.Value = pgbStatus.Value + 1
'        Loop
'
'        pgbStatus.Value = 400
'
'        pnlMsg.Caption = "금액자료 수신이 완료되었습니다." & Chr(13) & "총 " & g_Count & " 건"
'    End If
'
'End_Loop:
'    Close
'    On Error GoTo ERR_FILE
'
'    If Dir(App.Path & "\RecvData\*.*") <> "" Then
'        Kill App.Path & "\RecvData\*.*"
'    End If
'
'    Exit Function
'
'Err_Loop:
'    pnlMsg.Caption = "금액자료 수신이 취소되었습니다."
'
'    MsgBox "작업오류 입니다." & Chr(13) & _
'           "오류코드 : " & VBA.Err.Number & Chr(13) & _
'           "오류설명 : " & VBA.Err.Description, vbExclamation, "오류"
'    Resume End_Loop
'
'ERR_FILE:
'    pnlMsg.Caption = "파일삭제중 오류가 밸생했습니다."
'
'    MsgBox "작업오류 입니다." & Chr(13) & _
'           "오류코드 : " & VBA.Err.Number & Chr(13) & _
'           "오류설명 : " & VBA.Err.Description, vbExclamation, "오류"
'    Resume Next
'End Function
'
'Private Function DaySaleData() As Boolean
'    Dim g_SaleDate As String
'    Dim g_Count As Integer
'    Dim Filename As String
'    Dim St As String
'
'    If Not IsDate(txtDate.Text) Then
'        MsgBox "일자가 잘못 입력되었습니다.", vbExclamation, "입력오류"
'        Exit Function
'    End If
'    g_SaleDate = Format(txtDate.Text, "YYYY-MM-DD")
'
'    pnlMsg.Caption = "목요세일자료 확인중..!"
'
'    On Error GoTo Err_Loop
'
'    Filename = Dir(strSendPath & "\D????????" & g_AgencyCode & ".DAT")
'
'    If Filename = "" Then
'        pnlMsg.Caption = ""
'        MsgBox "목요세일자료가 없습니다.", vbInformation, "확인"
'        Exit Function
'    Else
'        pgbStatus.MAX = 400
'        pgbStatus.Min = 0
'        pgbStatus.Value = 0
'
'        pnlMsg.Caption = "목요세일자료 수신중..!"
'
'        ' 모뎀일 경우에만 복사한다
'        ' 인터넷일경우 다른곳에서 복사한다.
'        If SendMode = 1 Then FileCopy strSendPath & "\" & Filename, App.Path & "\RecvData\" & Filename
'
'        Open App.Path & "\RecvData\" & Filename For Input As #1
'        Open App.Path & "\BackData\" & Filename For Output As #2
'
'        Do While Not EOF(1)
'            Line Input #1, St
'            Print #2, St
'            g_Count = g_Count + 1
'            If pgbStatus.Value = 400 Then
'               pgbStatus.Value = 0
'            End If
'            pgbStatus.Value = pgbStatus.Value + 1
'        Loop
'        pgbStatus.Value = 400
'
'        pnlMsg.Caption = "목요세일자료 수신이 완료되었습니다." & Chr(13) & "총 " & g_Count & " 건"
'    End If
'
'End_Loop:
'    Close
'    On Error GoTo ERR_FILE
'
'    If Dir(App.Path & "\RecvData\*.*") <> "" Then
'        Kill App.Path & "\RecvData\*.*"
'    End If
'
'    Exit Function
'
'Err_Loop:
'    pnlMsg.Caption = "목요세일자료 수신이 취소되었습니다."
'
'    MsgBox "작업오류 입니다." & Chr(13) & _
'           "오류코드 : " & VBA.Err.Number & Chr(13) & _
'           "오류설명 : " & VBA.Err.Description, vbExclamation, "오류"
'    Resume End_Loop
'ERR_FILE:
'    pnlMsg.Caption = "파일삭제중 오류가 밸생했습니다."
'
'    MsgBox "작업오류 입니다." & Chr(13) & _
'           "오류코드 : " & VBA.Err.Number & Chr(13) & _
'           "오류설명 : " & VBA.Err.Description, vbExclamation, "오류"
'    Resume Next
'End Function
'
'Private Sub Form_Unload(Cancel As Integer)
'
'    'daoDB.Close
'
'    FormActivate = False
'
'    If SendMode = 1 Then
'        RasDial.HangUp
'    ElseIf SendMode = 2 Then
'        ' 클라이언트 소켓을 종료시킴
'        Winsock1.Close
'        FTC.RemoteClose
'
'        If Winsock1.State <> sckClosed Then Winsock1.Close
'
'        Do
'            DoEvents
'            If Winsock1.State = sckListening Then
'                ' 클라이언트 소켓을 종료시킴
'                Winsock1.Close
'                Exit Do
'            End If
'
'            If Winsock1.State = sckError Or Winsock1.State = sckClosed Then
'                Exit Do
'            End If
'        Loop
'    End If
'
'End Sub
'
'Private Function RepairData() As Boolean
'    Dim g_Count As Integer
'    Dim Filename As String
'    Dim St As String
'
'    If Not IsDate(txtDate.Text) Then
'        MsgBox "일자가 잘못 입력되었습니다.", vbExclamation, "입력오류"
'        Exit Function
'    End If
'
'    pnlMsg.Caption = "수선자료 확인중..!"
'
'    On Error GoTo Err_Loop
'
'    Filename = Dir(strSendPath & "\R????????" & ".DAT")
'
'    If Filename = "" Then
'        pnlMsg.Caption = ""
'        MsgBox "수선자료가 없습니다.", vbInformation, "확인"
'        Exit Function
'    Else
'        pgbStatus.MAX = 400
'        pgbStatus.Min = 0
'        pgbStatus.Value = 0
'
'        pnlMsg.Caption = "수선자료 수신중..!"
'
'        ' 모뎀일 경우에만 복사한다
'        ' 인터넷일경우 다른곳에서 복사한다.
'        If SendMode = 1 Then FileCopy strSendPath & "\" & Filename, App.Path & "\RecvData\" & Filename
'
'        Open App.Path & "\RecvData\" & Filename For Input As #1
'        Open App.Path & "\BackData\" & Filename For Output As #2
'
'        Do While Not EOF(1)
'            Line Input #1, St
'            Print #2, St
'            g_Count = g_Count + 1
'            If pgbStatus.Value = 400 Then
'               pgbStatus.Value = 0
'            End If
'            pgbStatus.Value = pgbStatus.Value + 1
'        Loop
'        pgbStatus.Value = 400
'        pnlMsg.Caption = "수선자료 수신이 완료되었습니다." & Chr(13) & "총 " & g_Count & " 건"
'    End If
'
'End_Loop:
'    Close
'    On Error GoTo ERR_FILE
'    Do While Dir(App.Path & "\RecvData\*.*") <> ""
'        Kill App.Path & "\RecvData\*.*"
'    Loop
'    Exit Function
'
'Err_Loop:
'    pnlMsg.Caption = "수선자료 수신이 취소되었습니다."
'    MsgBox "작업오류 입니다." & Chr(13) & _
'           "오류코드 : " & VBA.Err.Number & Chr(13) & _
'           "오류설명 : " & VBA.Err.Description, vbExclamation, "오류"
'    Resume End_Loop
'ERR_FILE:
'    pnlMsg.Caption = "파일삭제중 오류가 밸생했습니다."
'    MsgBox "작업오류 입니다." & Chr(13) & _
'           "오류코드 : " & VBA.Err.Number & Chr(13) & _
'           "오류설명 : " & VBA.Err.Description, vbExclamation, "오류"
'    Resume Next
'End Function
'
'Private Function MailData() As Boolean
'    Dim g_Count As Integer
'    Dim Filename As String
'    Dim St As String
'    Dim sData(100, 2) As String
'
'    pnlMsg.Caption = "메일자료 확인중..!"
'
'    On Error GoTo Err_Loop
'
'    Filename = Dir(strSendPath & "\M*." & g_AgencyCode)
'
'    If Filename = "" Then
'        pnlMsg.Caption = ""
'        MsgBox "메일자료가 없습니다.", vbInformation, "확인"
'
'        Exit Function
'    Else
'        Do
'            ' 모뎀일 경우에만 복사한다
'            ' 인터넷일경우 다른곳에서 복사한다.
'            If SendMode = 1 Then FileCopy strSendPath & "\" & Filename, App.Path & "\RecvData\" & Filename
'
'            Open App.Path & "\RecvData\" & Filename For Input As #1
'            Open App.Path & "\RecvData\Mail.Dat" For Output As #2
'
'            Do While Not EOF(1)
'                Line Input #1, St
'                Print #2, St
'                g_Count = g_Count + 1
'            Loop
'
'            Close
'            Filename = Dir
'        Loop While Not Filename = ""
'
'        pgbStatus.MAX = g_Count
'        pgbStatus.Min = 0
'        pgbStatus.Value = 0
'    End If
'
'    'Workspaces(0).BeginTrans
'
'    pnlMsg.Caption = "메일자료를 적용중 입니다...!" & Chr(13) & "총 " & g_Count & " 건"
'
'    Open App.Path & "\RecvData\Mail.Dat" For Input As #1
'
'    Do While Not EOF(1)
'        Line Input #1, St
'
'        ' 처음의 데이터가 일자이면
'        If IsDate(Format(Mid(St, 1, 8), "####-##-##")) = True Then
'            sData(i, 0) = Mid(St, 1, 8)
'            sData(i, 1) = Mid(St, 10, 1)
'            sData(i, 2) = Mid(St, 12) & Chr(13)
'             i = i + 1
'        Else
'            sData(i - 1, 2) = sData(i - 1, 2) + St & Chr(13)
'        End If
'    Loop
'
'    For i = 0 To 100
'        If sData(i, 0) = "" Then
'            Exit For
'        End If
'
'        Query = "INSERT INTO TB_메일 "
'        Query = Query & "VALUES('2', "
'        Query = Query & "'" & sData(i, 0) & "', "
'        Query = Query & sData(i, 1) & ", "
'        Query = Query & "'" & Replace(sData(i, 2), "'", "~") & "', "
'        Query = Query & "'')"
'        ADOCon.Execute Query
'
'        'daoDB.Execute Query
'    Next i
'
'    'Workspaces(0).CommitTrans
'
'    pnlMsg.Caption = "메일자료 수신이 완료되었습니다." & Chr(13) & "총 " & g_Count & " 건"
'    MailData = True
'
'End_Loop:
'    Close
'
'    On Error GoTo ERR_FILE
'
'    If Dir(App.Path & "\RecvData\*.*") <> "" Then
'        Kill App.Path & "\RecvData\*.*"
'    End If
'    Exit Function
'
'Err_Loop:
''    Workspaces(0).Rollback
'
'    pnlMsg.Caption = "메일자료 수신이 취소되었습니다."
'
'    MsgBox "작업오류 입니다." & Chr(13) & _
'           "오류코드 : " & VBA.Err.Number & Chr(13) & _
'           "오류설명 : " & VBA.Err.Description, vbExclamation, "오류"
'
'    Resume End_Loop
'
'ERR_FILE:
'    pnlMsg.Caption = "파일삭제중 오류가 밸생했습니다."
'
'    MsgBox "작업오류 입니다." & Chr(13) & _
'           "오류코드 : " & VBA.Err.Number & Chr(13) & _
'           "오류설명 : " & VBA.Err.Description, vbExclamation, "오류"
'    Resume Next
'
'End Function
'
'Private Function SendMail() As Boolean
'    Dim g_IpDate As String
'    Dim g_Count As Integer
'    Dim g_STag As String
'    Dim g_ETag As String
'    Dim g_Amount As Currency
'    Dim g_Jae As Integer
'    Dim g_Su As Integer
'    Dim g_Ban As Integer
'    Dim g_Ma As Integer
'    Dim g_SDate As String
'    Dim g_PAmt As Currency
'    Dim tmpStr As String
'    Dim Filename As String
'    Dim temp As String
'    Dim FHandle As Integer
'
'    ' 일자를 Check한다.
'    If Not IsDate(txtDate.Text) Then
'        MsgBox "일자가 잘못 입력되었습니다.", vbExclamation, "입력오류"
'        Exit Function
'    End If
'
'    g_IpDate = Format(txtDate.Text, "YYYY-MM-DD")
'
'    On Error GoTo Err_Loop
'
'    pnlMsg.Caption = "메일자료 확인중..!"
'
'    '---------------------------------------------------------
'    ' 메일내역을 읽어온다.
'    '---------------------------------------------------------
'    Query = "SELECT    메일일자"
'    Query = Query & ", 메일번호"
'    Query = Query & ", 메일내역"
'    Query = Query & " FROM TB_메일 "
'    Query = Query & " WHERE 송수신구분 = '1' "
'    Query = Query & "   AND 전송구분  <> 'Y' "
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
'
'    If ADORs.EOF Then
'        pnlMsg.Caption = "메일자료가 없습니다."
'        GoTo End_Loop
'    End If
'
'    pnlMsg.Caption = "메일자료 생성중..!" & Chr(13) & "총 " & pgbStatus.MAX & " 건"
'
'    ' 메일파일를 생성한다.
'    Filename = "M" & CStr(g_IpDate) & "-" & g_AgencyCode & "-1.DAT"
'    If UCase(Dir(App.Path & "\Internet\" & Filename)) = UCase(Filename) Then
'        Kill App.Path & "\Internet" & "\" & Filename
'    End If
'
'    FHandle = FreeFile
'
'    Open App.Path & "\Internet" & "\" & Filename For Output Lock Read Write As #FHandle
'
'    Do While Not ADORs.EOF
'        pgbStatus.MAX = ADORs.RecordCount
'        pgbStatus.Min = 0
'        pgbStatus.Value = 0
'
'        tmpStr = "1|"
'        tmpStr = tmpStr & ADORs!메일일자 & "|"
'        tmpStr = tmpStr & g_AgencyCode & "|"
'        tmpStr = tmpStr & ADORs!메일번호 & "|"
'        tmpStr = tmpStr & ADORs!메일내역 & "|"
'        tmpStr = tmpStr & "|"
'
'        Print #FHandle, tmpStr
'
'        tmpStr = "UPDATE TB_메일 "
'        tmpStr = tmpStr & "SET 전송구분 = 'Y' "
'        tmpStr = tmpStr & "WHERE 메일일자 = '" & ADORs!메일일자 & "' "
'        tmpStr = tmpStr & "AND   메일번호 = " & ADORs!메일번호 & " "
'
'        ADOCon.Execute tmpStr
'
'        ADORs.MoveNext
'    Loop
'    ADORs.Close
'    Set ADORs = Nothing
'
'    Close #FHandle
'    pnlMsg.Caption = "메일자료 생성이 완료되었습니다." & Chr(13) & "총 " & pgbStatus.MAX & " 건"
'
'
'    ' 본사에 파일을 전송한다.
'    SendFile = App.Path & "\Internet" & "\" & Filename
'    SendFlag = lauSendMail
'    SendMail = InterNetSendFiles(SendFile)
'
'End_Loop:
'    Close #FHandle
'    Exit Function
'
'Err_Loop:
'    pnlMsg.Caption = "메일자료 전송이 취소되었습니다."
'
'    MsgBox "작업오류 입니다." & Chr(13) & _
'           "오류코드 : " & VBA.Err.Number & Chr(13) & _
'           "오류설명 : " & VBA.Err.Description, vbExclamation, "오류"
'
'    Resume End_Loop
'End Function
'
'
'Private Function InterNetFileRequest(Mode As LaundrySendFlag) As Boolean
''인터넷에서 파일을 수신을 요청한다.
'    Dim g_SaleDate As String
'    Dim g_Count As Integer
'    Dim Filename As String
'    Dim strMsg As String
'    Dim MsgFile0 As String
'    Dim MsgTimeOut As String
'    Dim MsgEnd As String
'    Dim St As String
'
'    InterNetFileRequest = False
'
'    If Not IsDate(txtDate.Text) Then
'        MsgBox "일자가 잘못 입력되었습니다.", vbExclamation, "입력오류"
'        Exit Function
'    End If
'
'    g_SaleDate = Format(txtDate.Text, "YYYY-MM-DD")
'
'    Select Case Mode
'        Case lauRecvMail
'            strMsg = "메일 자료"
'            Filename = "M*." & g_AgencyCode
'
'        Case lauChulGo
'            strMsg = "출고 자료"
'            Filename = "Down" & g_AgencyCode & "*"
'
'        Case lauDaySaleData
'            strMsg = "목요세일 자료"
'            Filename = "\D????????" & g_AgencyCode & ".DAT"
'
'        Case lauSaleData
'            strMsg = "할인 자료"
'            Filename = "Sale" & g_AgencyCode & ".DAT"
'
'        Case lauPriceData
'            strMsg = "금액 자료"
'            Filename = "\????????" & g_AgencyCode & ".DAT"
'
'        Case lauRepairData
'            strMsg = "수선 자료"
'            Filename = "\R????????" & ".DAT"
'
'        Case lauProgram
'            strMsg = "프로그램 자료"
'            Filename = "Sale" & g_AgencyCode & ".DAT"
'    End Select
'
'    On Error GoTo ERR_END
'
'    pnlMsg.Caption = strMsg & "를 확인하는 중입니다."
'
'    ' 전송할 메시지 작성
''   S_STA       : 메시지의 시작을 의미한다.
''   CLEANAID    : 프로그램을 사용하고 있는 회사를 의미한다.
''   1001        : 프로그램을 사용하고 있는 회사중 각각의 코드를 의미한다.(지사등등)
''   OK_CREATE_CHULGO_DATA : 출고파일을 올바르게 생성했다.
''   DATA        : 생성한 파일 이름이 전달 된다.
''   S_END       : 메시지의 종료를 의미한다.
'
'    strSendData = S_STA & "|" & "CLEANAID" & "|" & 가맹점정보.지사코드 & "|" & _
'                  "GET_RECV_FILE_COUNT" & "|" & _
'                   Filename & "|" & _
'                  S_END
'
'
'    ' 소켓상태가 연결인 경우만 데이타를 보냄
'    If Winsock1.State = sckClosed Or Winsock1.State = sckError Then
'        pnlMsg.Caption = "본사와 연결되지 않았습니다."
'        Exit Function
'    End If
'
'    ' 소켓상태가 연결인 경우만 데이타를 보냄
'    FileTotalCount = -1
'
'    If Winsock1.State = sckConnected Then
'        ' 데이타를 서버에 보냄
'        Winsock1.SendData CStr(strSendData)
'    End If
'
'    ' Winsock1_DataArrival 이쪽으로 파일 리스트가 날라온다.
'    PauseTime = 10                  ' 기간을 지정합니다.
'    Start = Timer                   ' 시작 시간을 지정합니다.
'
'    Do While FileTotalCount = -1
'        Finish = Timer              ' 종료 시간을 지정합니다.
'        TotalTime = Finish - Start  ' 전체 시간을 계산합니다.
'
'        If Timer > Start + PauseTime Then
'            pnlMsg.Caption = strMsg & " 요청에 응답하지 않습니다."
'            If Mode = lauRecvMail Then NextPross = True
'
'            Exit Function
'        End If
'        DoEvents                    ' 다른 프로시저로 넘깁니다.
'    Loop
'
'    If FileTotalCount = 0 Then
'        pnlMsg.Caption = "수신할 " & strMsg & "가 없습니다."
'        If Mode = lauRecvMail Then NextPross = True
'        Call Delay(2)
'        Exit Function
'    End If
'
'    pnlMsg.Caption = strMsg & "를 수신중 입니다."
'
''    DoEvents
'    strSendData = S_STA & "|" & "CLEANAID" & "|" & 가맹점정보.지사코드 & "|" & _
'                  "GET_ALL_FILE" & "|" & _
'                   Filename & "|" & _
'                  S_END
'
''    strSendData = S_STA & "|" & S_MYIP & "=" & GetIPAddress & "|" & _
''                  S_CUSTCODE & "=" & g_AgencyCode & "|" & _
''                  S_CUSTNAME & "=" & 가맹점정보.가맹점명 & "|" & _
''                  S_MYFILEPORT & "=" & Fn_GetFileLocatPort & "|" & _
''                  S_GETALLFILE & "=" & FileName & "|" & _
''                  S_GETFILE & "=" & "ALL" & "|" & _
''                  S_END
'
'    If Winsock1.State = sckConnected Then
'        ' 데이타를 서버에 보냄
'        Winsock1.SendData strSendData
'    End If
'
''    DoEvents
'    InterNetFileRequest = True
'
'    Exit Function
'
'ERR_END:
'
'End Function
'
'Private Function SendCust() As Boolean
'    Dim g_IpDate As String
'    Dim g_Count As Integer
'    Dim g_STag As String
'    Dim g_ETag As String
'    Dim g_Amount As Currency
'    Dim g_Jae As Integer
'    Dim g_Su As Integer
'    Dim g_Ban As Integer
'    Dim g_Ma As Integer
'    Dim g_SDate As String
'    Dim g_PAmt As Currency
'    Dim tmpStr As String
'    Dim Filename As String
'    Dim temp As String
'    Dim SendFile As String
'    Dim FHandel As Integer
'
'
'    ' 일자를 Check한다.
'    If Not IsDate(txtDate.Text) Then
'        MsgBox "일자가 잘못 입력되었습니다.", vbExclamation, "입력오류"
'        Exit Function
'    End If
'
'    g_IpDate = Format(txtDate.Text, "YYYY-MM-DD")
'
'    On Error GoTo Err_Loop
'
'    pnlMsg.Caption = "고객자료 확인중..!"
'
'
'    '---------------------------------------------------------
'    ' 메일내역을 읽어온다.
'    '---------------------------------------------------------
'    Query = "SELECT    고객코드"
'    Query = Query & ", 성명"
'    Query = Query & ", 전화번호"
'    Query = Query & ", 주소"
'    Query = Query & " FROM TB_고객정보 "
'    Query = Query & " WHERE 전송구분 = 'N' "
'    Query = Query & "    OR 전송구분 Is Null "
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
'
'    If ADORs.EOF Then
'        pnlMsg.Caption = "전송할 고객자료가 없습니다."
'        GoTo End_Loop
'    End If
'
'    pnlMsg.Caption = "고객자료 생성중..!" & Chr(13) & "총 " & ADORs.RecordCount & " 건"
'
'    ' 고객파일를 생성한다.
'    Filename = "C" & CStr(g_IpDate) & "-" & g_AgencyCode & "-1.DAT"
'    If UCase(Dir(App.Path & "\Internet\" & Filename)) = UCase(Filename) Then
'        Kill App.Path & "\Internet" & "\" & Filename
'    End If
'
'    FHandel = FreeFile
'    Open App.Path & "\Internet" & "\" & Filename For Output Lock Read Write As #FHandel
'
'    Do While Not ADORs.EOF
'        pgbStatus.MAX = ADORs.RecordCount
'        pgbStatus.Min = 0
'        pgbStatus.Value = 0
'
'        tmpStr = g_AgencyCode & "|"
'        tmpStr = tmpStr & ADORs!고객코드 & "|"
'        tmpStr = tmpStr & ADORs!성명 & "|"
'        tmpStr = tmpStr & ADORs!전화번호 & "|"
'        tmpStr = tmpStr & ADORs!주소 & "|"
'
'        Print #FHandel, tmpStr
'
'        Query = "UPDATE TB_고객정보 "
'        Query = Query & "SET 전송구분 = 'Y' "
'        Query = Query & "WHERE 고객코드 = '" & ADORs!고객코드 & "'"
'
'        ADOCon.Execute Query
'
'        ADORs.MoveNext
'    Loop
'    ADORs.Close
'    Set ADORs = Nothing
'
'    SendCust = True
'
'    pnlMsg.Caption = "고객자료 생성이 완료되었습니다." & Chr(13) & "총 " & pgbStatus.MAX & " 건"
'
'    Close #FHandel
'
'    ' 인터넷 전송 부분
''    DoEvents
'    PauseTime = 10                 ' 기간을 지정합니다.
'    Start = Timer                   ' 시작 시간을 지정합니다.
'
'    Do While FTC.State <> ftcReady
'        Finish = Timer              ' 종료 시간을 지정합니다.
'        TotalTime = Finish - Start  ' 전체 시간을 계산합니다.
'
'        If Timer > Start + PauseTime Then
'            panMsg.Caption = "전송 실패"
'            Exit Do
'        End If
'        DoEvents                    ' 다른 프로시저로 넘깁니다.
'    Loop
'
'    SendFile = App.Path & "\Internet" & "\" & Filename
'    SendFlag = lauSendCust
'    SendCust = InterNetSendFiles(SendFile)
'
'End_Loop:
'    Close #FHandel
'    Exit Function
'
'Err_Loop:
'    pnlMsg.Caption = "고객자료 전송이 취소되었습니다."
'
'    MsgBox "작업오류 입니다." & Chr(13) & _
'           "오류코드 : " & VBA.Err.Number & Chr(13) & _
'           "오류설명 : " & VBA.Err.Description, vbExclamation, "오류"
'    Resume End_Loop
'End Function
'
'
'Private Function SendCoupon() As Boolean
'    Dim g_IpDate As String
'    Dim g_Count As Integer
'    Dim g_STag As String
'    Dim g_ETag As String
'    Dim g_Amount As Currency
'    Dim g_Jae As Integer
'    Dim g_Su As Integer
'    Dim g_Ban As Integer
'    Dim g_Ma As Integer
'    Dim g_SDate As String
'    Dim g_PAmt As Currency
'    Dim tmpStr As String
'    Dim Filename As String
'    Dim temp As String
'    Dim SendFile As String
'    Dim FHandel As Integer
'
'    g_IpDate = Format(txtDate.Text, "YYYY-MM-DD")
'
'
'    On Error GoTo Err_Loop
'
'    panMsg.Caption = "쿠폰 자료 확인중..!"
'
'    '--------------------------------------------------------
'    ' 메일내역을 읽어온다.
'    '--------------------------------------------------------
'    Query = "SELECT * FROM TB_쿠폰자료 "
'    Query = Query & "WHERE ( 전송여부 = 'N' "
'    Query = Query & "     OR 전송여부 Is Null) "
'    Query = Query & "  OR 전송일자 = '" & g_IpDate & "' "
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
'
'    If ADORs.EOF Then
'        panMsg.Caption = "전송할 쿠폰 자료가 없습니다."
'        GoTo End_Loop
'    End If
'
'    ADORs.MoveLast
'    panMsg.Caption = "쿠폰자료 생성중..!" & Chr(13) & "총 " & ADORs.RecordCount & " 건"
'    ADORs.MoveFirst
'    ' 쿠폰파일을 생성한다.
'    Filename = "P" & CStr(g_IpDate) & "-" & g_AgencyCode & "-1.DAT"
'    If UCase(Dir(App.Path & "\Internet\" & Filename)) = UCase(Filename) Then
'        Kill App.Path & "\Internet" & "\" & Filename
'    End If
'    FHandel = 125 'FreeFile
'    Open App.Path & "\Internet" & "\" & Filename _
'        For Output Lock Read Write As #FHandel
'
'    Do While Not ADORs.EOF
'
'        tmpStr = g_AgencyCode & "|"
'        tmpStr = tmpStr & ADORs!접수일자 & "|"
'        tmpStr = tmpStr & ADORs!대리점코드 & "|"
'        tmpStr = tmpStr & ADORs!쿠폰번호 & "|"
'        tmpStr = tmpStr & ADORs!쿠폰단가 & "|"
'        tmpStr = tmpStr & ADORs!쿠폰금액 & "|"
'        tmpStr = tmpStr & ADORs!고객코드 & "|"
'        tmpStr = tmpStr & ADORs!고객이름 & "|"
'        tmpStr = tmpStr & ADORs!접수금액 & "|"
'        tmpStr = tmpStr & ADORs!택번호 & "|"
'
'        Print #FHandel, tmpStr
'
'        '---------------------------------------------------------------------
'        '
'        '---------------------------------------------------------------------
'        Query = "UPDATE TB_쿠폰자료 "
'        Query = Query & " SET 전송여부 = 'Y', "
'        Query = Query & "     전송일자 = '" & Format(Date, "YYYY-MM-DD") & "' "
'        Query = Query & " WHERE 접수일자 = '" & ADORs!접수일자 & "'"
'        Query = Query & "   AND 대리점코드 = '" & ADORs!대리점코드 & "'"
'        Query = Query & "   AND 쿠폰번호 = '" & ADORs!쿠폰번호 & "'"
'        ADOCon.Execute Query
'
'        ADORs.MoveNext
'    Loop
'
'    SendCoupon = True
'
'    ADORs.Close
'    Set ADORs = Nothing
'
'    Close #FHandel
'
'    DoEvents
'
'    ' 인터넷 전송 부분
''    DoEvents
'    PauseTime = 10                 ' 기간을 지정합니다.
'    Start = Timer                   ' 시작 시간을 지정합니다.
'
'    Do While FTC.State <> ftcReady
'        Finish = Timer              ' 종료 시간을 지정합니다.
'        TotalTime = Finish - Start  ' 전체 시간을 계산합니다.
'
'        If Timer > Start + PauseTime Then
'            panMsg.Caption = "전송 실패"
'            Exit Do
'        End If
'
'        DoEvents                    ' 다른 프로시저로 넘깁니다.
'    Loop
'
'    SendFile = App.Path & "\Internet" & "\" & Filename
'    SendFlag = lauSendCoupoon
'    SendCoupon = InterNetSendFiles(SendFile)
'
'End_Loop:
'    Close #FHandel
'    Exit Function
'
'Err_Loop:
'    panMsg.Caption = "쿠폰자료 전송이 취소되었습니다."
'
'    MsgBox "작업오류 입니다." & Chr(13) & _
'           "오류코드 : " & VBA.Err.Number & Chr(13) & _
'           "오류설명 : " & VBA.Err.Description, vbExclamation, "오류"
'    Resume End_Loop
'End Function
'
'Private Function InterNetProgramUpgrade() As Boolean
'    Dim Filename As String
'    Dim strVersion As String
'
'    On Error GoTo ERR_END
'
'    InterNetProgramUpgrade = False
'    strNewVersion = ""
'    strMsg = "프로그램 자료"
'    Filename = App.EXEName
'    pnlMsg.Caption = strMsg & "를 확인하는 중입니다."
'
'    ' 전송할 메시지 작성
'    strSendData = S_STA & "|" & S_MYIP & "=" & GetIPAddress & "|" & _
'                  S_CUSTCODE & "=" & g_AgencyCode & "|" & _
'                  S_CUSTNAME & "=" & 가맹점정보.가맹점명 & "|" & _
'                  S_MYFILEPORT & "=" & Fn_GetFileLocatPort & "|" & _
'                  S_GETPROGRAMVERSION & "=" & Filename & "|" & _
'                  S_END
'
'    ' 소켓상태가 연결인 경우만 데이타를 보냄
'    If Winsock1.State = sckClosed Or Winsock1.State = sckError Then
'        pnlMsg.Caption = "본사와 연결되지 않았습니다."
'        Exit Function
'    End If
'
'    ' 소켓상태가 연결인 경우만 데이타를 보냄
'    If Winsock1.State = sckConnected Then
'        ' 데이타를 서버에 보냄
'        Winsock1.SendData strSendData
'    End If
'
'    ' Winsock1_DataArrival 이쪽으로 파일 리스트가 날라온다.
'    PauseTime = 10                  ' 기간을 지정합니다.
'    Start = Timer                   ' 시작 시간을 지정합니다.
'
'    Do While strNewVersion = ""
'        Finish = Timer              ' 종료 시간을 지정합니다.
'        TotalTime = Finish - Start  ' 전체 시간을 계산합니다.
'
'        If Timer > Start + PauseTime Then
'            pnlMsg.Caption = strMsg & " 요청에 응답하지 않습니다."
'            Exit Function
'        End If
'        DoEvents                    ' 다른 프로시저로 넘깁니다.
'    Loop
'
'    If strNewVersion <= Program_Version Then
'        pnlMsg.Caption = "업그레이드할 " & strMsg & "가 없습니다."
'        Call ButtonEnable(True)
'        Exit Function
'    End If
'
'    pnlMsg.Caption = strMsg & "를 수신중 입니다."
'    DoEvents
'
'    strSendData = S_STA & "|" & S_MYIP & "=" & GetIPAddress & "|" & _
'                  S_CUSTCODE & "=" & g_AgencyCode & "|" & _
'                  S_CUSTNAME & "=" & 가맹점정보.가맹점명 & "|" & _
'                  S_MYFILEPORT & "=" & Fn_GetFileLocatPort & "|" & _
'                  S_GETPROGRAMFILE & "=" & Filename & "_UP.EXE" & "|" & _
'                  S_GETFILE & "=" & "ALL" & "|" & _
'                  S_END
'
'    If Winsock1.State = sckConnected Then
'        ' 데이타를 서버에 보냄
'        Winsock1.SendData strSendData
'    End If
'
'    DoEvents
'    InterNetProgramUpgrade = True
'    Exit Function
'
'ERR_END:
'
'End Function
'
'Private Function ProgramUpgrade() As Boolean
'    Dim Filename As String
'    Dim strVersion As String
'    Dim strNewVersion As String
'
'On Error GoTo Err_Loop
'
'    '  이전 "ok.ok" 화일을 지운다
'    If Not Dir(App.Path & "\OK.OK") = "" Then
'        Kill App.Path & "\OK.OK"
'    End If
'
'
'    '대리점 코드
'    Filename = Dir(strSendPath & "\Laundry_UP.exe")
'
'    If Filename = "" Then
'        pnlMsg.Caption = "업그레이드된 내용이 없습니다"
'        Exit Function
'    Else
'
'        ' 모뎀일 경우에만 복사한다
'        ' 인터넷일경우 다른곳에서 복사한다.
'        If SendMode = 1 Then
'
''            FileCopy strPrgPath & "\version.ini", App.Path & "\version.ini"
'
'            Shell App.Path & "\PGDOWN.BAT", vbHide
'
'            pgbStatus.MAX = Int(ShowFolderSize(strPrgPath & "\Laundry_UP.EXE"))
'            Do While Dir(App.Path & "\OK.OK") = ""
'                pgbStatus.Value = pgbStatus.Value + 1
'
'                If pgbStatus.Value = pgbStatus.MAX Then
'                    pgbStatus.Value = 0
'                End If
'
'                DoEvents
'            Loop
'            pgbStatus.Value = pgbStatus.MAX
'
'        ElseIf SendMode = 2 Then
'            ' 인터넷일경우 다른 폴더에 받아진다.
'            Call FileCopy(strSendPath & "\Laundry_UP.exe", App.Path & "\Laundry_UP.exe")
'
'        End If
'
'        Shell App.Path & "\Laundry_UP.EXE", vbNormalFocus
'        End
'    End If
'
'End_Loop:
'    Exit Function
'
'Err_Loop:
'    pnlMsg.Caption = "프로그램 Update가 취소되었습니다."
'
'    MsgBox "작업오류 입니다." & Chr(13) & _
'           "오류코드 : " & VBA.Err.Number & Chr(13) & _
'           "오류설명 : " & VBA.Err.Description, vbExclamation, "오류"
'    Resume End_Loop
'End Function
'
'Private Sub txtPassWord_KeyDown(KeyCode As Integer, Shift As Integer)
'    If (KeyCode < 48 Or KeyCode > 57) And (KeyCode < 96 Or KeyCode > 105) Then
'        If txtPassWord.PasswordChar <> "*" Then
'            txtPassWord.PasswordChar = "*"
'        End If
'    Else
'        txtPassWord.PasswordChar = ""
'    End If
'
'    If KeyCode = vbKeyReturn Then
'        Command1_Click
'    End If
'End Sub
'
'Private Sub FTC_SendComplete(ByVal FileSize As Long)
'
'    Debug.Print Now & "FTC_SendComplete" & CStr(SendFlag)
'
'    ' 체인점 -> 본사 전송 완료
'    Select Case SendFlag
'        Case lauInput
'
'            cmdBtn(2).Enabled = True
'            chkPassWord = True
'
'            '-----------------------------------------------------------
'            '
'            '-----------------------------------------------------------
'            Query = "UPDATE TB_일일마감 "
'            Query = Query & " SET 전송여부 = '*' "
'            Query = Query & " WHERE 일자 = '" & txtDate.ClipText & "' "
'            ADOCon.Execute Query
'
'            pnlMsg.Caption = "입고자료 전송이 완료되었습니다." & Chr(13) & "총 " & pgbStatus.MAX & " 건"
'
'            MsgBox "입고자료 전송 완료" & Space(10), vbInformation, "확인"
'
'            Debug.Print Now & "입고자료 전송 완료" & CStr(SendFlag)
'
'        Case lauQNData
'            pnlMsg.Caption = "보관서비스 자료 전송이 완료되었습니다." & Chr(13) & "총 " & pgbStatus.MAX & " 건"
'
'        Case lauSendMail
'            pnlMsg.Caption = "메일자료 전송이 완료되었습니다." & Chr(13) & "총 " & pgbStatus.MAX & " 건"
'
'        Case lauSendDB
'            pnlMsg.Caption = "DB자료 전송이 완료되었습니다." & Chr(13) & "총 " & "1 건"
'
'        Case lauSendCust
'            pnlMsg.Caption = "고객자료 전송이 완료되었습니다." & Chr(13) & "총 " & pgbStatus.MAX & " 건"
'
'        Case lauSendCoupoon
'            pnlMsg.Caption = "쿠폰자료 전송이 완료되었습니다." & Chr(13) & "총 " & pgbStatus.MAX & " 건"
'
'        Case lauMileage
'            If 가맹점정보.마일리지여부 = "Y" Then
'                '------------------------------------------------------------------------------
'                ' 마일리지스토리
'                '------------------------------------------------------------------------------
'                Query = "UPDATE TB_마일리지스토리 SET"
'                Query = Query & " 전송여부 = 'Y', "
'                Query = Query & " 전송일자 = '" & Format(Date, "YYYY-MM-DD") & "' "
'                Query = Query & " WHERE 전송여부 = 'N' OR 전송일자 = '" & Format(Date, "YYYY-MM-DD") & "'"
'                ADOCon.Execute Query
'
'                '------------------------------------------------------------------------------
'                ' 마일리지현황
'                '------------------------------------------------------------------------------
'                Query = "UPDATE TB_마일리지현황 SET"
'                Query = Query & " 전송여부 = 'Y', "
'                Query = Query & " 전송일자 = '" & Format(Date, "YYYY-MM-DD") & "' "
'                Query = Query & " WHERE 전송여부 = 'N'"
'                Query = Query & "    OR 전송일자 = '" & Format(Date, "YYYY-MM-DD") & "'"
'                ADOCon.Execute Query
'
'                pnlMsg.Caption = "마일리지자료 전송이 완료되었습니다." & Chr(13) & "총 " & pgbStatus.MAX & " 건"
'            End If
'    End Select
'End Sub
'
'Private Sub FTC_ReceiveComplete(ByVal FileSize As Long)
''    Call ButtonEnable(True)
'    DoEvents
'
'    ' 본사 -> 체인점 수신 완료
'    Select Case SendFlag
'
'        Case lauRecvMail
'        ' 메일 자료 수신
'            FileProCount = FileProCount + 1
'            If FileProCount >= FileTotalCount Then
'                ' 받은 메일을 처리한다.
'                pnlMsg.Caption = "메일자료 수신이 완료 되었습니다." & vbLf & "처리를 위하여 잠시만 기다려 주십시요"
'                DoEvents
'                Timer_FTC.Tag = "RecvMail"
'                Timer_FTC.Enabled = True
'                Exit Sub
'            End If
'
'        Case lauChulGo
'        ' 출고 자료 수신
'            FileProCount = FileProCount + 1
'            If FileProCount >= FileTotalCount Then
'                ' 받은 출고자료를 처리한다.
'                pnlMsg.Caption = "출고자료 수신이 완료 되었습니다." & vbLf & "처리를 위하여 잠시만 기다려 주십시요"
'                DoEvents
'                Timer_FTC.Tag = "CHULGO"
'                Timer_FTC.Enabled = True
'                Exit Sub
'            End If
'
'        Case lauSaleData
'        ' 할인자료
'            FileProCount = FileProCount + 1
'            If FileProCount >= FileTotalCount Then
'                ' 받은 할인자료을 처리한다.
'                pnlMsg.Caption = "할인자료 수신이 완료 되었습니다." & vbLf & "처리를 위하여 잠시만 기다려 주십시요"
'                DoEvents
'                Timer_FTC.Tag = "SaleData"
'                Timer_FTC.Enabled = True
'                Exit Sub
'            End If
'
'        Case lauDaySaleData
'        ' 세일자료
'            FileProCount = FileProCount + 1
'            If FileProCount >= FileTotalCount Then
'                ' 받은 할인자료을 처리한다.
'                pnlMsg.Caption = "세일자료 수신이 완료 되었습니다." & vbLf & "처리를 위하여 잠시만 기다려 주십시요"
'                DoEvents
'                Timer_FTC.Tag = "DaySaleData"
'                Timer_FTC.Enabled = True
'                Exit Sub
'            End If
'
'        Case lauPriceData
'        ' 금액자료
'            FileProCount = FileProCount + 1
'            If FileProCount >= FileTotalCount Then
'                ' 받은 금액자료 처리한다.
'                pnlMsg.Caption = "금액자료 수신이 완료 되었습니다." & vbLf & "처리를 위하여 잠시만 기다려 주십시요"
'                DoEvents
'                Timer_FTC.Tag = "PriceData"
'                Timer_FTC.Enabled = True
'                Exit Sub
'            End If
'
'        Case lauProgram
'        ' 프로그램 업그레이드
'            ' 받은 금액표을 처리한다.
'            pnlMsg.Caption = "적용중 입니다. 잠시후 프로그램이 재시작 됩니다."
'            DoEvents
'            Timer_FTC.Tag = "Program"
'            Timer_FTC.Enabled = True
'            Exit Sub
'
'    End Select
'
'End Sub
'
'Private Sub FTC_ReceiveStart(Filename As String, ByVal FileSize As Long, Overwrite As Boolean, Cancel As Boolean)
'
'    ' 기존 파일이 있을경우 덮어 씨운다.
'    Overwrite = True
'
'    Call ButtonEnable(False)
'    pgbStatus.MAX = FTC.ReceiveFileSize
'
'    Select Case SendFlag
'        Case lauRecvMail
'            pnlMsg.Caption = Filename & "메일자료를 받는중 입니다."
'            Exit Sub
'
'        Case lauChulGo
'            pnlMsg.Caption = "[" & Format(Mid(Filename, 8, 8), "YYYY-MM-DD") & "일]" & "출로자료를 받는중입니다."
'            Exit Sub
'
'        Case lauProgram
'            pnlMsg.Caption = "프로그램을 받는중 입니다. 잠시만 기다려 주십시요."
'            Exit Sub
'
'    End Select
'End Sub
'
'
'Private Sub Winsock1_Close()
'        panMsg.Caption = "본사와 연결이 종료 되었습니다. "
'        ButtonEnable (False)
'End Sub
'
'Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'' 클라이언트에서 데이타가 수신되면
'    Dim work As String
'    Dim varFlgVal   As Variant
'    Dim varValue As Variant
'    Dim strVal As String
'
'    Winsock1.GetData work, vbString
'
'    ' --------------------------------------------------------------
'    ' 수신한 데이터가 미리 정한 형식에 맞는지를 확인한다.
'
'    varValue = Split(work, "|")
'    If UBound(varValue) < 2 Then
'        ' 최소 3개보다는 많아야 한다.
'        ' 잘못 수신된 데이타를 저장한다.
'        Exit Sub
'    End If
'    '수신된 처음과 마지막을 확인한다.
'    If CStr(varValue(0)) <> S_STA Then
'        Exit Sub
'    ElseIf CStr(varValue(UBound(varValue))) <> S_END Then
'        Exit Sub
'    End If
'
'    ' --------------------------------------------------------------
'    ' 수신한 데이터를 처리한다.
'    Select Case CStr(varValue(3))
'
'        ' 본사에서 수신할 파일 갯수가 날라온다.
'        Case "SEND_ALLFILE_COUNT"
'            FileTotalCount = CStr(varValue(4))
'            Exit Sub
'
'        Case "HOST_FILE_PORT"
'            If UBound(Split(CStr(varValue(4)), ";")) <> 1 Then
'                panMsg.Caption = "파일 서버에 연결할 정보가 올바르지 않습니다."
'            Else
'                FTC.LocalPort = Val(Fn_GetFileLocatPort)
'                FTC.ReceiveDirPath = App.Path & "\RecvData"
'                FTC.RemoteHost = CStr(Split(CStr(varValue(4)), ";")(0))
'                FTC.RemotePort = CStr(Split(CStr(varValue(4)), ";")(1))
'                FTC.RemoteConnect
'            End If
'        Case Else
'
'    End Select
'
'End Sub
'
'Private Sub ButtonEnable(bFlag As Boolean)
'    Dim sDate   As String
'
'    cmdIpgoSend.Enabled = bFlag     ' 입고자료보내기
'    cmdChulgoRecv.Enabled = bFlag   ' 출고자료받기
'    cmdSaleRecv.Enabled = bFlag     ' 할인자료받기
'    cmdDaySale.Enabled = bFlag      ' 목요세일자료받기
'    cmdPriceRecv.Enabled = bFlag    ' 금액자료받기
'    cmdRepair.Enabled = bFlag       ' 수선자료받기
'    cmdDBSend.Enabled = bFlag       ' DB본사전송
'    cmdCustSend.Enabled = bFlag     ' 고객자료 보내기
'
'    sDate = Trim(GetIniStr("Store", "OldDate", "", iniFile))
'
'    If sDate = "" Or Not IsDate(sDate) Or sDate < Date Then
'        cmdPGRecv.Enabled = False
'    Else
'        cmdPGRecv.Enabled = bFlag       ' 프로그램 업그레이드
'    End If
'
'
'' 클라이언트 모드에서의 일부기능 제한
'    If Trim(chkProgramMode) = "2" Then
'
'        cmdDBSend.Enabled = False
'    End If
'End Sub
'
'
'Private Function InterNetSendFiles(SendFile As String) As Boolean
'' sendFile = 전체 경로이다.
'' strRecvPath = 는 서버의 수신 경로
'    Dim Filename As String
'
'    Filename = Mid(SendFile, InStrRev(SendFile, "\") + 1, Len(SendFile))
'
'    Select Case SendFlag
'        Case lauQNData:      pnlMsg.Caption = "보관 서비스 자료를 전송중 입니다."
'        Case lauInput:       pnlMsg.Caption = "입고자료를 전송중 입니다."
'        Case lauSendMail:    pnlMsg.Caption = "메일자료 전송중 입니다."
'        Case lauSendDB:      pnlMsg.Caption = "DB자료 전송중 입니다."
'        Case lauSendCust:    pnlMsg.Caption = "고객자료 전송중 입니다."
'        Case lauMileage:     pnlMsg.Caption = "마일리지자료 전송중 입니다."
'        Case lauSendCoupoon: pnlMsg.Caption = "쿠폰자료 전송중 입니다."
'    End Select
'
'    If Dir(SendFile, vbDirectory) <> "" Then
'        If SendMode = 1 Then
'            ' 모뎀 전송 부분
'            Call FileCopy(SendFile, strRecvPath & "\" & Filename)
'            DoEvents
'            DoEvents
'            Call FileCopy(SendFile, "\\CleanAid\Sugum\" & m_MstCode & "_" & Filename)
'
'            InterNetSendFiles = True
'
'            Select Case SendFlag
'                Case lauInput
'                    pnlMsg.Caption = "입고자료 전송이 완료되었습니다." & Chr(13) & "총 " & pgbStatus.MAX & " 건"
'
'                    cmdBtn(2).Enabled = True
'                    chkPassWord = True
'
'                    Query = "UPDATE TB_일일마감 "
'                    Query = Query & "SET 전송여부 = '*' "
'                    Query = Query & "WHERE 일자 = '" & txtDate.ClipText & "' "
'                    ADOCon.Execute Query
'
'                    MsgBox "입고자료 전송 완료" & Space(10), vbInformation, "확인"
'
'                Case lauQNData
'                    pnlMsg.Caption = "보관 서비스 자료 전송이 완료되었습니다." & Chr(13) & "총 " & pgbStatus.MAX & " 건"
'
'                Case lauSendMail
'                    pnlMsg.Caption = "메일자료 전송이 완료되었습니다." & Chr(13) & "총 " & pgbStatus.MAX & " 건"
'
'                Case lauSendDB
'                    pnlMsg.Caption = "DB자료 전송이 완료되었습니다." & Chr(13) & "총 " & "1 건"
'
'                Case lauSendCust
'                    pnlMsg.Caption = "고객자료 전송이 완료되었습니다." & Chr(13) & "총 " & pgbStatus.MAX & " 건"
'
'                Case lauMileage
'                    If 가맹점정보.마일리지여부 = "Y" Then
'                        Query = "UPDATE TB_마일리지스토리 "
'                        Query = Query & " SET 전송여부 = 'Y' "
'                        Query = Query & " 전송일자 = '" & Format(Date, "YYYY-MM-DD") & "' "
'                        Query = Query & " WHERE 전송여부 = 'N' OR 전송일자 = '" & Format(Date, "YYYY-MM-DD") & "'"
'                        ADOCon.Execute Query
'
'                        Query = "UPDATE TB_마일리지현황 "
'                        Query = Query & " SET 전송여부 = 'Y' "
'                        Query = Query & " 전송일자 = '" & Format(Date, "YYYY-MM-DD") & "' "
'                        Query = Query & " WHERE 전송여부 = 'N' OR 전송일자 = '" & Format(Date, "YYYY-MM-DD") & "'"
'                        ADOCon.Execute Query
'
'                        pnlMsg.Caption = "마일리지자료 전송이 완료되었습니다." & Chr(13) & "총 " & pgbStatus.MAX & " 건"
'                    End If
'            End Select
'
'        ElseIf SendMode = 2 Then
'            ' 인터넷 전송 부분
'            DoEvents
'            PauseTime = 10                 ' 기간을 지정합니다.
'            Start = Timer                   ' 시작 시간을 지정합니다.
'
'            Do While FTC.State <> ftcReady
'                Finish = Timer              ' 종료 시간을 지정합니다.
'                TotalTime = Finish - Start  ' 전체 시간을 계산합니다.
'
'                If Timer > Start + PauseTime Then
'                    InterNetSendFiles = False
'                    panMsg.Caption = "전송 실패"
'                    Exit Do
'                End If
'                DoEvents                    ' 다른 프로시저로 넘깁니다.
'            Loop
'
'            FTC.SendFile SendFile, 10
'            InterNetSendFiles = True
'            DoEvents
'
'        End If
'    Else
'        InterNetSendFiles = False
'    End If
'
'End Function
'
