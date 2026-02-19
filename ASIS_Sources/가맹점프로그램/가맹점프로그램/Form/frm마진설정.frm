VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form frm마진설정 
   Caption         =   "마진 설정"
   ClientHeight    =   7695
   ClientLeft      =   8085
   ClientTop       =   2985
   ClientWidth     =   6690
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   11.25
      Charset         =   129
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm마진설정.frx":0000
   LinkTopic       =   "Form31"
   LockControls    =   -1  'True
   ScaleHeight     =   7695
   ScaleWidth      =   6690
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdSave 
      Caption         =   "저장"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   4320
      TabIndex        =   24
      Top             =   4890
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   " [기타 설정] "
      Height          =   2775
      Index           =   1
      Left            =   330
      TabIndex        =   20
      Top             =   1590
      Width           =   6015
      Begin VB.ComboBox cboReturn 
         Height          =   345
         ItemData        =   "frm마진설정.frx":08CA
         Left            =   2310
         List            =   "frm마진설정.frx":08D4
         Style           =   2  '드롭다운 목록
         TabIndex        =   40
         Top             =   2010
         Width           =   1935
      End
      Begin VB.ComboBox cboCoupon 
         Height          =   345
         ItemData        =   "frm마진설정.frx":08E4
         Left            =   2340
         List            =   "frm마진설정.frx":08EE
         Style           =   2  '드롭다운 목록
         TabIndex        =   33
         Top             =   1170
         Width           =   1935
      End
      Begin VB.ComboBox cboMilAdd 
         Enabled         =   0   'False
         Height          =   345
         ItemData        =   "frm마진설정.frx":08FE
         Left            =   4320
         List            =   "frm마진설정.frx":0908
         Style           =   2  '드롭다운 목록
         TabIndex        =   32
         Top             =   330
         Width           =   1545
      End
      Begin VB.ComboBox cboSale 
         Height          =   345
         ItemData        =   "frm마진설정.frx":0926
         Left            =   2340
         List            =   "frm마진설정.frx":0930
         Style           =   2  '드롭다운 목록
         TabIndex        =   25
         Top             =   750
         Width           =   1935
      End
      Begin VB.ComboBox cboMil 
         Height          =   345
         ItemData        =   "frm마진설정.frx":0940
         Left            =   2340
         List            =   "frm마진설정.frx":094A
         Style           =   2  '드롭다운 목록
         TabIndex        =   23
         Top             =   330
         Width           =   1935
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   375
         Index           =   17
         Left            =   180
         TabIndex        =   21
         Top             =   330
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "마일리지 사용"
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   375
         Index           =   1
         Left            =   180
         TabIndex        =   26
         Top             =   750
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "특정할인 사용"
         RoundedCorners  =   0   'False
      End
      Begin MSMask.MaskEdBox mskSale 
         Height          =   345
         Left            =   4320
         TabIndex        =   27
         Top             =   750
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   375
         Index           =   2
         Left            =   180
         TabIndex        =   29
         Top             =   1590
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "고가세탁 비율"
         RoundedCorners  =   0   'False
      End
      Begin MSMask.MaskEdBox mskLuxury 
         Height          =   345
         Left            =   2340
         TabIndex        =   30
         Top             =   1590
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   375
         Index           =   3
         Left            =   180
         TabIndex        =   34
         Top             =   1170
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "쿠폰할인 사용"
         RoundedCorners  =   0   'False
      End
      Begin MSMask.MaskEdBox mskCoupon 
         Height          =   345
         Left            =   4320
         TabIndex        =   35
         Top             =   1170
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   375
         Index           =   6
         Left            =   150
         TabIndex        =   41
         Top             =   2010
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "세탁비환불 사용"
         RoundedCorners  =   0   'False
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  '단일 고정
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   5520
         TabIndex        =   36
         Top             =   1170
         Width           =   300
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  '단일 고정
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   3540
         TabIndex        =   31
         Top             =   1590
         Width           =   300
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  '단일 고정
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   5520
         TabIndex        =   28
         Top             =   750
         Width           =   300
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " [설정 변경 암호] "
      Height          =   1155
      Index           =   0
      Left            =   330
      TabIndex        =   17
      Top             =   330
      Width           =   5985
      Begin VB.TextBox txtPassWord 
         Height          =   375
         IMEMode         =   3  '사용 못함
         Left            =   2340
         PasswordChar    =   "*"
         TabIndex        =   19
         Top             =   405
         Width           =   2220
      End
      Begin VB.CommandButton cmdPass 
         Caption         =   "확인"
         Height          =   495
         Left            =   4785
         TabIndex        =   18
         Top             =   360
         Width           =   945
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   375
         Index           =   0
         Left            =   150
         TabIndex        =   22
         Top             =   420
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "본사 확인 코드"
         RoundedCorners  =   0   'False
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "종료"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   4320
      TabIndex        =   16
      Top             =   6270
      Width           =   2055
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3060
      Left            =   330
      TabIndex        =   3
      Top             =   4500
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   5398
      _Version        =   262144
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " [마진 설정] "
      Begin MSMask.MaskEdBox mskRatio 
         Height          =   375
         Left            =   2175
         TabIndex        =   0
         Top             =   300
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   375
         Index           =   4
         Left            =   150
         TabIndex        =   4
         Top             =   300
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   661
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "세탁마진"
         RoundedCorners  =   0   'False
      End
      Begin MSMask.MaskEdBox mskSRatio 
         Height          =   375
         Left            =   2175
         TabIndex        =   2
         Top             =   1170
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   375
         Index           =   8
         Left            =   150
         TabIndex        =   5
         Top             =   735
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   661
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "운동화마진"
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   375
         Index           =   11
         Left            =   150
         TabIndex        =   6
         Top             =   1170
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   661
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "수선마진"
         RoundedCorners  =   0   'False
      End
      Begin MSMask.MaskEdBox mskSports 
         Height          =   375
         Left            =   2175
         TabIndex        =   1
         Top             =   735
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskGa 
         Height          =   360
         Left            =   2175
         TabIndex        =   10
         Top             =   1605
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   375
         Index           =   12
         Left            =   150
         TabIndex        =   11
         Top             =   1605
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   661
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "가죽/무스탕"
         RoundedCorners  =   0   'False
      End
      Begin MSMask.MaskEdBox mskCar 
         Height          =   360
         Left            =   2175
         TabIndex        =   13
         Top             =   2040
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   375
         Index           =   13
         Left            =   150
         TabIndex        =   14
         Top             =   2040
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   661
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "카페트 마진"
         RoundedCorners  =   0   'False
      End
      Begin MSMask.MaskEdBox mskOut 
         Height          =   360
         Left            =   2175
         TabIndex        =   37
         Top             =   2610
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   375
         Index           =   5
         Left            =   150
         TabIndex        =   38
         Top             =   2610
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   661
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "외주 운동화마진"
         RoundedCorners  =   0   'False
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  '단일 고정
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   3375
         TabIndex        =   39
         Top             =   2610
         Width           =   300
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  '단일 고정
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   3375
         TabIndex        =   15
         Top             =   2040
         Width           =   300
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  '단일 고정
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   3375
         TabIndex        =   12
         Top             =   1605
         Width           =   300
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  '단일 고정
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   3375
         TabIndex        =   9
         Top             =   300
         Width           =   300
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  '단일 고정
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3375
         TabIndex        =   8
         Top             =   1170
         Width           =   300
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  '단일 고정
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3375
         TabIndex        =   7
         Top             =   735
         Width           =   300
      End
   End
End
Attribute VB_Name = "frm마진설정"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnPassOK As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'+------------------------------------------------------
'+ 2003/08/29 수정
'+
'+루틴설명      - 비밀번호확인
'+  1. 암호를 확인하여 암호 규칙에 맞으면 화면을 설정한다.
'+  2. 레지스터리에 저장한다.
'+
'+------------------------------------------------------
Private Sub cmdPass_Click()
    Dim strPass As String
    
    ' 입력 확인
    blnPassOK = False
    If Len(txtPassWord.Text) <= 0 Then Exit Sub
    
    
'   기본 디폴드 암호.. ( 프로그램 셋팅/설치를 위한 암호 )
    If UCase(txtPassWord.Text) = "DUDTJSGH" Then
        blnPassOK = True
    
    Else
        ' 비밀번호 확인
        strPass = IsSportsPassWord(txtPassWord.Text)
        If strPass = "-1" Or strPass = "-3" Then
            blnPassOK = False
            txtPassWord.SelStart = 0: txtPassWord.SelLength = Len(txtPassWord.Text)
            If strPass = "-3" Then MsgBox "입력한 내용이 정확하지 않습니다.", vbInformation, "입력오류"
            txtPassWord.Text = ""
            txtPassWord.SetFocus
            Exit Sub
        End If
        blnPassOK = True
        
    End If
    
    If blnPassOK = True Then
        Call ButtonEnabled(True)
        txtPassWord.Text = ""
        mskRatio.SetFocus
        Exit Sub
    End If

End Sub

'+------------------------------------------------------
'+
'+ 2003/02/03
'+
'+루틴설명
'+  1. strPass로 전달된 비밀번호의 유효성을 검사한다
'+  2. 전달값
'+     strPass :   "05????????????"   앞 2자리는 유효 일자
'+                                       2자리 다음은 비빌번호
'+                                       ( 일자 * 365 * 1544 )
'+  3. 리턴값
'+     앞 2자리를 리턴한다. ( 사용기간 )
'+     -1 :         임의 수정한 경우
'+     -3 :         입력한 내용이 틀린 경우
'+
'+------------------------------------------------------
Private Function IsSportsPassWord(strPass) As String
    Dim nday As Double
    Dim intMM As Integer
    Dim dPass As Double
    Dim strTemp As String

    If Not IsNumeric(Mid(strPass, 1, 2)) Then
        MsgBox "전달된 본사확인코드의 형식이 정확하지 않습니다.", vbInformation, "입력오류"
        IsSportsPassWord = "-1"
        Exit Function
    End If
'    strPass = Mid(strPass, 3, Len(strPass) - 2)
    ' 오늘의 일자를 구한다.
    nday = Val(Format(Date, "dd"))
    intMM = Val(Format(Date, "mm"))
    dPass = nday * intMM * 1544
    If strPass = dPass Then
        IsSportsPassWord = Mid(strPass, 1, 2)
    Else
        IsSportsPassWord = "-3"
    End If
    
End Function

Private Sub cmdSave_Click()
    Dim msg As String
    
    On Error GoTo ErrRtn
    
    msg = "[0 ~ 100] 사이의 숫자만입력이 가능합니다."
    
    If Not IsNumeric(mskRatio.ClipText) Or Val(mskRatio.ClipText) < 0 Or Val(mskRatio.ClipText) > 100 Then
        mskRatio.SelStart = 0:  mskRatio.SelLength = 100:   mskRatio.SetFocus
        MsgBox msg, vbInformation, "확인"
        Exit Sub
    End If
    
    If Not IsNumeric(mskSRatio.ClipText) Or Val(mskSRatio.ClipText) < 0 Or Val(mskSRatio.ClipText) > 100 Then
        mskSRatio.SelStart = 0:  mskSRatio.SelLength = 100:   mskSRatio.SetFocus
        MsgBox msg, vbInformation, "확인"
        Exit Sub
    End If
    
    If Not IsNumeric(mskSports.ClipText) Or Val(mskSports.ClipText) < 0 Or Val(mskSports.ClipText) > 100 Then
        mskSports.SelStart = 0:  mskSports.SelLength = 100:   mskSports.SetFocus
        MsgBox msg, vbInformation, "확인"
        Exit Sub
    End If
    
    If Not IsNumeric(mskGa.ClipText) Or Val(mskGa.ClipText) < 0 Or Val(mskGa.ClipText) > 100 Then
        mskGa.SelStart = 0:  mskGa.SelLength = 100:   mskGa.SetFocus
        MsgBox msg, vbInformation, "확인"
        Exit Sub
    End If
    
    If Not IsNumeric(mskCar.ClipText) Or Val(mskCar.ClipText) < 0 Or Val(mskCar.ClipText) > 100 Then
        mskCar.SelStart = 0:  mskCar.SelLength = 100:   mskCar.SetFocus
        MsgBox msg, vbInformation, "확인"
        Exit Sub
    End If
    
    If Not IsNumeric(mskOut.ClipText) Or Val(mskOut.ClipText) < 0 Or Val(mskOut.ClipText) > 100 Then
        mskOut.SelStart = 0:  mskOut.SelLength = 100:   mskOut.SetFocus
        MsgBox msg, vbInformation, "확인"
        Exit Sub
    End If


    Query = "UPDATE TB_기본정보 "
    Query = Query & "SET 비율       = '" & mskRatio.ClipText & "', "
    Query = Query & "    수선마진   = '" & mskSRatio.ClipText & "', "
    Query = Query & "    운동화마진     = '" & mskSports.ClipText & "', "
    Query = Query & "    가죽무스탕마진 = '" & mskGa.ClipText & "', "
    Query = Query & "    카페트마진     = '" & mskCar.ClipText & "', "
    Query = Query & "    외주운동화마진 = '" & mskOut.ClipText & "', "
    
    Query = Query & "    마일리지여부   = '" & IIf(Trim(cboMil.Text) = "예", "Y", "N") & "', "
    Query = Query & "    마일리지증가구분   = '" & IIf(cboMilAdd.ListIndex = 0, "0", "1") & "', "
    Query = Query & "    세탁비환불여부   = '" & IIf(Trim(cboReturn.Text) = "예", "Y", "N") & "', "
    Query = Query & "    지정할인여부   = '" & IIf(Trim(cboSale.Text) = "예", "Y", "N") & "', "
    Query = Query & "    지정할인비율     = '" & mskSale.ClipText & "',  "
    Query = Query & "    특정할인여부   = '" & IIf(Trim(cboCoupon.Text) = "예", "Y", "N") & "', "
    Query = Query & "    특정할인비율     = '" & mskCoupon.ClipText & "',  "
    Query = Query & "    고가세탁비율     = '" & mskLuxury.ClipText & "'  "
    ADOCon.Execute Query
    
    
    frm환경설정.txtRatio.Text = mskRatio.ClipText
    frm환경설정.txtSRatio.Text = mskSRatio.ClipText
    frm환경설정.txtSports.Text = mskSports.ClipText
    frm환경설정.txtGa.Text = mskGa.ClipText
    frm환경설정.txtCar.Text = mskCar.ClipText
    frm환경설정.txtOut.Text = mskOut.ClipText
    
    frm환경설정.cboMil.ListIndex = IIf(Trim(cboMil.Text) = "예", 0, 1)
    frm환경설정.cboMilAdd.ListIndex = cboMilAdd.ListIndex
    frm환경설정.cboSale.ListIndex = IIf(Trim(cboSale.Text) = "예", 0, 1)
    frm환경설정.txtSale.Text = mskSale.ClipText
    frm환경설정.cboCoupon.ListIndex = IIf(Trim(cboCoupon.Text) = "예", 0, 1)
    frm환경설정.txtCoupon.Text = mskCoupon.ClipText
    frm환경설정.txtLuxury.Text = mskLuxury.ClipText
        
    MsgBox "대리점 정보가 변경되어 프로그램을 다시 시작하셔야 합니다.", vbInformation, "확인"
    
    End
    
ErrRtn:
    MsgBox "[저장중 오류]" & Err.Description, vbInformation, "확인"
    Exit Sub
End Sub

Private Sub Form_Activate()
    txtPassWord.SetFocus
End Sub

Private Sub Form_Load()
    MoveWindow Me
    
    Call ButtonEnabled(False)
    
    If Fun_대리점정보 <> "Error" Then
        mskRatio.Text = 대리점정보.비율
        mskSports.Text = 대리점정보.운동화마진
        mskSRatio.Text = 대리점정보.수선마진
        mskGa.Text = 대리점정보.가죽무스탕마진
        mskCar.Text = 대리점정보.카페트마진
        mskOut.Text = 대리점정보.외주운동화마진
        
        
        mskSale.Text = 대리점정보.지정할인비율
        mskCoupon.Text = 대리점정보.특정할인비율
        mskLuxury.Text = 대리점정보.고가세탁비율
        cboMil.ListIndex = IIf(대리점정보.마일리지여부 = "Y", 0, 1)
        cboMilAdd.ListIndex = IIf(대리점정보.마일리지증가구분 = "0", 0, 1)
        cboSale.ListIndex = IIf(대리점정보.지정할인여부 = "Y", 0, 1)
        cboCoupon.ListIndex = IIf(대리점정보.특정할인여부 = "Y", 0, 1)
        cboReturn.ListIndex = IIf(대리점정보.세탁비환불여부 = "Y", 0, 1)
    End If
End Sub

Private Sub txtPassWord_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode < 48 Or KeyCode > 57) And (KeyCode < 96 Or KeyCode > 105) Then
        If txtPassWord.PasswordChar <> "*" Then
            txtPassWord.PasswordChar = "*"
        End If
    Else
        txtPassWord.PasswordChar = ""
    End If
    
    If KeyCode = vbKeyReturn Then
        cmdPass_Click
    End If
End Sub

Private Sub ButtonEnabled(bMode As Boolean)
    mskRatio.Enabled = bMode
    mskSports.Enabled = bMode
    mskSRatio.Enabled = bMode
    mskGa.Enabled = bMode
    mskCar.Enabled = bMode
    cmdSave.Enabled = bMode
    cboMil.Enabled = bMode
    cboMilAdd.Enabled = bMode
    cboSale.Enabled = bMode
    mskSale.Enabled = bMode
    cboCoupon.Enabled = bMode
    mskCoupon.Enabled = bMode
    mskLuxury.Enabled = bMode
    mskOut.Enabled = bMode
    cboReturn.Enabled = bMode
End Sub
