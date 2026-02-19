VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12045
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9540
   ScaleWidth      =   12045
   StartUpPosition =   3  'Windows 기본값
   Begin XtremeSuiteControls.TabControl TabControl 
      Height          =   10215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5340
      _Version        =   851970
      _ExtentX        =   9419
      _ExtentY        =   18018
      _StockProps     =   68
      Appearance      =   3
      Color           =   16
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.ButtonMargin=   "0,3,0,3"
      ItemCount       =   2
      Item(0).Caption =   " 발송 정보 "
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "TabControlPage(0)"
      Item(1).Caption =   " 개별 문자발송 "
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "TabControlPage1"
      Begin XtremeSuiteControls.TabControlPage TabControlPage1 
         Height          =   9735
         Left            =   -69970
         TabIndex        =   1
         Top             =   450
         Visible         =   0   'False
         Width           =   5280
         _Version        =   851970
         _ExtentX        =   9313
         _ExtentY        =   17171
         _StockProps     =   1
         BackColor       =   16777215
         Page            =   1
         Begin VB.TextBox txtServer 
            Height          =   345
            Index           =   0
            Left            =   1320
            TabIndex        =   5
            Top             =   6075
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.TextBox txtServer 
            Height          =   345
            Index           =   1
            Left            =   1320
            TabIndex        =   4
            Top             =   6465
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.TextBox txtServer 
            Height          =   345
            Index           =   2
            Left            =   1320
            TabIndex        =   3
            Top             =   6855
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.TextBox txtServer 
            Height          =   345
            Index           =   3
            Left            =   1320
            TabIndex        =   2
            Top             =   7245
            Visible         =   0   'False
            Width           =   2895
         End
         Begin XtremeSuiteControls.PushButton cmdSvr 
            Height          =   450
            Index           =   0
            Left            =   1320
            TabIndex        =   6
            Top             =   7650
            Visible         =   0   'False
            Width           =   1125
            _Version        =   851970
            _ExtentX        =   1984
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "초기화"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdSvr 
            Height          =   450
            Index           =   1
            Left            =   2475
            TabIndex        =   7
            Top             =   7650
            Visible         =   0   'False
            Width           =   1125
            _Version        =   851970
            _ExtentX        =   1984
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "연결 확인"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdSvr 
            Height          =   450
            Index           =   2
            Left            =   3630
            TabIndex        =   8
            Top             =   7650
            Visible         =   0   'False
            Width           =   1125
            _Version        =   851970
            _ExtentX        =   1984
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "저장"
            UseVisualStyle  =   -1  'True
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   5610
            Index           =   1
            Left            =   0
            TabIndex        =   9
            Top             =   60
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   9895
            _Version        =   262144
            BackColor       =   16777215
            PictureFrames   =   1
            Picture         =   "Form1.frx":0000
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.TextBox txtRecvTel 
               Height          =   375
               Left            =   1455
               TabIndex        =   11
               Top             =   2670
               Width           =   1410
            End
            Begin VB.TextBox txtSend2 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  '없음
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1920
               Left            =   375
               MultiLine       =   -1  'True
               TabIndex        =   10
               Top             =   630
               Width           =   2430
            End
            Begin XtremeSuiteControls.PushButton cmdSend 
               Height          =   630
               Left            =   540
               TabIndex        =   12
               Top             =   4635
               Width           =   2145
               _Version        =   851970
               _ExtentX        =   3784
               _ExtentY        =   1111
               _StockProps     =   79
               Caption         =   " 문자보내기"
               Appearance      =   6
               Picture         =   "Form1.frx":3D252
            End
            Begin VB.Label lblTitle 
               Alignment       =   1  '오른쪽 맞춤
               BackColor       =   &H00FFC0C0&
               BackStyle       =   0  '투명
               Caption         =   "보내는 사람:"
               Height          =   225
               Index           =   8
               Left            =   210
               TabIndex        =   14
               Top             =   2760
               Width           =   1200
            End
            Begin VB.Label lblLan2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  '투명
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   300
               Left            =   2655
               TabIndex        =   13
               Top             =   195
               Width           =   135
            End
         End
         Begin Threed.SSFrame SSFrame 
            Height          =   4425
            Index           =   1
            Left            =   3135
            TabIndex        =   15
            Top             =   120
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   7805
            _Version        =   262144
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "받는 사람"
            Begin VB.TextBox txtMobile 
               Height          =   375
               Index           =   9
               Left            =   105
               TabIndex        =   25
               Top             =   3960
               Width           =   1830
            End
            Begin VB.TextBox txtMobile 
               Height          =   375
               Index           =   8
               Left            =   105
               TabIndex        =   24
               Top             =   3555
               Width           =   1830
            End
            Begin VB.TextBox txtMobile 
               Height          =   375
               Index           =   7
               Left            =   105
               TabIndex        =   23
               Top             =   3150
               Width           =   1830
            End
            Begin VB.TextBox txtMobile 
               Height          =   375
               Index           =   6
               Left            =   105
               TabIndex        =   22
               Top             =   2745
               Width           =   1830
            End
            Begin VB.TextBox txtMobile 
               Height          =   375
               Index           =   5
               Left            =   105
               TabIndex        =   21
               Top             =   2340
               Width           =   1830
            End
            Begin VB.TextBox txtMobile 
               Height          =   375
               Index           =   4
               Left            =   105
               TabIndex        =   20
               Top             =   1935
               Width           =   1830
            End
            Begin VB.TextBox txtMobile 
               Height          =   375
               Index           =   3
               Left            =   105
               TabIndex        =   19
               Top             =   1530
               Width           =   1830
            End
            Begin VB.TextBox txtMobile 
               Height          =   375
               Index           =   2
               Left            =   105
               TabIndex        =   18
               Top             =   1125
               Width           =   1830
            End
            Begin VB.TextBox txtMobile 
               Height          =   375
               Index           =   1
               Left            =   105
               TabIndex        =   17
               Top             =   720
               Width           =   1830
            End
            Begin VB.TextBox txtMobile 
               Height          =   375
               Index           =   0
               Left            =   105
               TabIndex        =   16
               Top             =   315
               Width           =   1830
            End
         End
         Begin VB.Label lblTitle 
            Alignment       =   1  '오른쪽 맞춤
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  '투명
            Caption         =   "서버 IP :"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   29
            Top             =   6150
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Label lblTitle 
            Alignment       =   1  '오른쪽 맞춤
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  '투명
            Caption         =   "서버 DB :"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   28
            Top             =   6540
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Label lblTitle 
            Alignment       =   1  '오른쪽 맞춤
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  '투명
            Caption         =   "사용자 이름:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   27
            Top             =   6945
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Label lblTitle 
            Alignment       =   1  '오른쪽 맞춤
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  '투명
            Caption         =   "비밀번호 :"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   26
            Top             =   7305
            Visible         =   0   'False
            Width           =   1140
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage 
         Height          =   9735
         Index           =   0
         Left            =   30
         TabIndex        =   30
         Top             =   450
         Width           =   5280
         _Version        =   851970
         _ExtentX        =   9313
         _ExtentY        =   17171
         _StockProps     =   1
         BackColor       =   16777215
         Page            =   0
         Begin VB.ComboBox cboSendText 
            Height          =   300
            Left            =   60
            Style           =   2  '드롭다운 목록
            TabIndex        =   31
            Top             =   90
            Width           =   5145
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   5610
            Index           =   2
            Left            =   1065
            TabIndex        =   32
            Top             =   975
            Width           =   3090
            _ExtentX        =   5450
            _ExtentY        =   9895
            _Version        =   262144
            BackColor       =   16777215
            PictureFrames   =   1
            Picture         =   "Form1.frx":3D94C
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.TextBox Text1 
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
               Left            =   330
               TabIndex        =   34
               Top             =   3045
               Width           =   2520
            End
            Begin VB.TextBox txtSMS 
               Appearance      =   0  '평면
               BorderStyle     =   0  '없음
               Height          =   1890
               Left            =   390
               MultiLine       =   -1  'True
               TabIndex        =   33
               Top             =   645
               Width           =   2400
            End
            Begin XtremeSuiteControls.PushButton cmdBtn 
               Height          =   630
               Index           =   1
               Left            =   525
               TabIndex        =   35
               Top             =   4695
               Width           =   2145
               _Version        =   851970
               _ExtentX        =   3784
               _ExtentY        =   1111
               _StockProps     =   79
               Caption         =   " 보내기"
               Appearance      =   6
               Picture         =   "Form1.frx":7AB9E
            End
            Begin VB.Label lblTitle 
               Alignment       =   1  '오른쪽 맞춤
               BackColor       =   &H00FFC0C0&
               BackStyle       =   0  '투명
               Caption         =   "보내는 사람:"
               Height          =   255
               Index           =   0
               Left            =   270
               TabIndex        =   37
               Top             =   2760
               Width           =   1170
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
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   2700
               TabIndex        =   36
               Top             =   300
               Width           =   120
            End
         End
         Begin XtremeSuiteControls.PushButton cmdSendTextSave 
            Height          =   390
            Index           =   2
            Left            =   1170
            TabIndex        =   38
            Top             =   465
            Width           =   540
            _Version        =   851970
            _ExtentX        =   952
            _ExtentY        =   688
            _StockProps     =   79
            Caption         =   "삭제"
            Appearance      =   6
         End
         Begin XtremeSuiteControls.PushButton cmdSendTextSave 
            Height          =   390
            Index           =   1
            Left            =   615
            TabIndex        =   39
            Top             =   465
            Width           =   540
            _Version        =   851970
            _ExtentX        =   952
            _ExtentY        =   688
            _StockProps     =   79
            Caption         =   "수정"
            Appearance      =   6
         End
         Begin XtremeSuiteControls.PushButton cmdSendTextSave 
            Height          =   390
            Index           =   0
            Left            =   60
            TabIndex        =   40
            Top             =   465
            Width           =   540
            _Version        =   851970
            _ExtentX        =   952
            _ExtentY        =   688
            _StockProps     =   79
            Caption         =   "추가"
            Appearance      =   6
         End
         Begin XtremeSuiteControls.PushButton cmdChange 
            Height          =   390
            Left            =   4185
            TabIndex        =   41
            Top             =   465
            Width           =   990
            _Version        =   851970
            _ExtentX        =   1746
            _ExtentY        =   688
            _StockProps     =   79
            Caption         =   "암호변경"
            Appearance      =   6
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

