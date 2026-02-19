VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form frm보관서비스 
   Caption         =   "보관 서비스"
   ClientHeight    =   7155
   ClientLeft      =   2595
   ClientTop       =   4365
   ClientWidth     =   11850
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   11850
   WindowState     =   2  '최대화
   Begin VB.Frame Frame3 
      Height          =   3555
      Left            =   7500
      TabIndex        =   27
      Top             =   3690
      Width           =   4305
      Begin VB.CommandButton Command1 
         Caption         =   "적용"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3120
         TabIndex        =   39
         Top             =   3030
         Width           =   1095
      End
      Begin VB.TextBox txtHaData 
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
         Index           =   4
         Left            =   1110
         TabIndex        =   37
         Top             =   2550
         Width           =   3105
      End
      Begin VB.TextBox txtHaData 
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
         Index           =   3
         Left            =   1080
         TabIndex        =   35
         Top             =   2100
         Width           =   3105
      End
      Begin VB.TextBox txtHaData 
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
         Index           =   2
         Left            =   1080
         TabIndex        =   33
         Top             =   1650
         Width           =   3105
      End
      Begin VB.TextBox txtHaData 
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
         Index           =   1
         Left            =   1080
         TabIndex        =   31
         Top             =   1170
         Width           =   3105
      End
      Begin VB.TextBox txtHaData 
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
         Index           =   0
         Left            =   1050
         TabIndex        =   29
         Top             =   660
         Width           =   3105
      End
      Begin Threed.SSPanel ssPanel3 
         Height          =   390
         Index           =   10
         Left            =   60
         TabIndex        =   28
         Top             =   150
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   688
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
         Caption         =   "하자 내용 입력"
         BevelWidth      =   2
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel ssPanel3 
         Height          =   390
         Index           =   11
         Left            =   90
         TabIndex        =   30
         Top             =   660
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   688
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
         Caption         =   "01"
         BevelWidth      =   2
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel ssPanel3 
         Height          =   390
         Index           =   12
         Left            =   120
         TabIndex        =   32
         Top             =   1170
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   688
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
         Caption         =   "02"
         BevelWidth      =   2
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel ssPanel3 
         Height          =   390
         Index           =   13
         Left            =   120
         TabIndex        =   34
         Top             =   1650
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   688
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
         Caption         =   "03"
         BevelWidth      =   2
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel ssPanel3 
         Height          =   390
         Index           =   14
         Left            =   120
         TabIndex        =   36
         Top             =   2100
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   688
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
         Caption         =   "04"
         BevelWidth      =   2
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel ssPanel3 
         Height          =   390
         Index           =   15
         Left            =   150
         TabIndex        =   38
         Top             =   2550
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   688
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
         Caption         =   "05"
         BevelWidth      =   2
         RoundedCorners  =   0   'False
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "보관서비스 - 필수 입력 내용"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   150
      TabIndex        =   12
      Top             =   1320
      Width           =   11565
      Begin VB.CommandButton cmdAction 
         Caption         =   "접수 취소"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   9450
         Style           =   1  '그래픽
         TabIndex        =   45
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "접수 수정"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   9450
         Style           =   1  '그래픽
         TabIndex        =   44
         Top             =   690
         Width           =   2055
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "보관 접수"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   9450
         Style           =   1  '그래픽
         TabIndex        =   43
         Top             =   180
         Width           =   2055
      End
      Begin VB.TextBox txtEMail 
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
         Left            =   5730
         TabIndex        =   24
         Top             =   840
         Width           =   3375
      End
      Begin VB.ComboBox cboDevTime 
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
         ItemData        =   "frm보관서비스.frx":0000
         Left            =   6360
         List            =   "frm보관서비스.frx":000D
         Style           =   2  '드롭다운 목록
         TabIndex        =   22
         Top             =   1290
         Width           =   2745
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1860
         TabIndex        =   20
         Top             =   1260
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   62324737
         CurrentDate     =   39024
      End
      Begin VB.ComboBox cboSaleDate 
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
         ItemData        =   "frm보관서비스.frx":002C
         Left            =   1890
         List            =   "frm보관서비스.frx":004E
         Style           =   2  '드롭다운 목록
         TabIndex        =   18
         Top             =   810
         Width           =   1965
      End
      Begin VB.TextBox txtUserNumber 
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
         Left            =   6540
         TabIndex        =   16
         Top             =   390
         Width           =   2565
      End
      Begin VB.ComboBox cboUserGubun 
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
         ItemData        =   "frm보관서비스.frx":00AC
         Left            =   1890
         List            =   "frm보관서비스.frx":00B6
         Style           =   2  '드롭다운 목록
         TabIndex        =   14
         Top             =   360
         Width           =   1965
      End
      Begin Threed.SSPanel ssPanel3 
         Height          =   390
         Index           =   2
         Left            =   150
         TabIndex        =   13
         Top             =   360
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   688
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
         Caption         =   "내/외국인"
         BevelWidth      =   2
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel ssPanel3 
         Height          =   390
         Index           =   4
         Left            =   3990
         TabIndex        =   15
         Top             =   390
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   688
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
         Caption         =   "주민등록번호"
         BevelWidth      =   2
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel ssPanel3 
         Height          =   390
         Index           =   6
         Left            =   150
         TabIndex        =   17
         Top             =   810
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   688
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
         Caption         =   "보관 기간"
         BevelWidth      =   2
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel ssPanel3 
         Height          =   390
         Index           =   7
         Left            =   150
         TabIndex        =   19
         Top             =   1260
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   688
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
         Caption         =   "만료예정일"
         BevelWidth      =   2
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel ssPanel3 
         Height          =   390
         Index           =   8
         Left            =   3990
         TabIndex        =   21
         Top             =   1290
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   688
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
         Caption         =   "배송요청시간"
         BevelWidth      =   2
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel ssPanel3 
         Height          =   390
         Index           =   9
         Left            =   3990
         TabIndex        =   23
         Top             =   840
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   688
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
         Caption         =   "이메일"
         BevelWidth      =   2
         RoundedCorners  =   0   'False
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1260
      Left            =   135
      TabIndex        =   4
      Top             =   0
      Width           =   11625
      Begin VB.ComboBox cboView 
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
         ItemData        =   "frm보관서비스.frx":00D2
         Left            =   8550
         List            =   "frm보관서비스.frx":00DC
         Style           =   2  '드롭다운 목록
         TabIndex        =   41
         Top             =   240
         Width           =   3045
      End
      Begin VB.CommandButton Command2 
         Caption         =   "전체 선택"
         Height          =   495
         Left            =   10410
         TabIndex        =   40
         Top             =   690
         Width           =   1125
      End
      Begin MSMask.MaskEdBox mskName 
         Height          =   390
         Index           =   0
         Left            =   4770
         TabIndex        =   5
         Top             =   255
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   688
         _Version        =   393216
         BackColor       =   14737632
         PromptInclude   =   0   'False
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
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskCode 
         Height          =   390
         Left            =   1635
         TabIndex        =   3
         Top             =   735
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   688
         _Version        =   393216
         BackColor       =   14737632
         PromptInclude   =   0   'False
         MaxLength       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskTEL 
         Height          =   390
         Index           =   0
         Left            =   1635
         TabIndex        =   0
         Top             =   270
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   688
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "9999"
         PromptChar      =   " "
      End
      Begin Threed.SSPanel ssPanel3 
         Height          =   390
         Index           =   0
         Left            =   225
         TabIndex        =   6
         Top             =   255
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   688
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
         Caption         =   "전화번호"
         BevelWidth      =   2
         BorderWidth     =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel ssPanel3 
         Height          =   390
         Index           =   1
         Left            =   3390
         TabIndex        =   7
         Top             =   255
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   688
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
         Caption         =   "성   명"
         BevelWidth      =   2
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel ssPanel3 
         Height          =   390
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   735
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   688
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
         Caption         =   "고객번호"
         BevelWidth      =   2
         RoundedCorners  =   0   'False
      End
      Begin MSMask.MaskEdBox mskTEL 
         Height          =   390
         Index           =   1
         Left            =   2460
         TabIndex        =   1
         Top             =   270
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   688
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "9999"
         PromptChar      =   " "
      End
      Begin Threed.SSPanel ssPanel3 
         Height          =   390
         Index           =   5
         Left            =   3390
         TabIndex        =   9
         Top             =   690
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   688
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
         Caption         =   "보관수량"
         BevelWidth      =   2
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel ssPanel3 
         Height          =   390
         Index           =   16
         Left            =   6810
         TabIndex        =   42
         Top             =   240
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   688
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
         Caption         =   "접수현황"
         BevelWidth      =   2
         RoundedCorners  =   0   'False
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "원"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   7740
         TabIndex        =   26
         Top             =   690
         Width           =   360
      End
      Begin VB.Label lblMoney 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   390
         Left            =   6120
         TabIndex        =   25
         Top             =   690
         Width           =   1605
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "개"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   5730
         TabIndex        =   11
         Top             =   690
         Width           =   360
      End
      Begin VB.Label lblTotalCount 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   390
         Left            =   4785
         TabIndex        =   10
         Top             =   690
         Width           =   915
      End
   End
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   3945
      Left            =   150
      TabIndex        =   2
      Top             =   3120
      Width           =   11580
      _Version        =   524288
      _ExtentX        =   20426
      _ExtentY        =   6959
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      DAutoCellTypes  =   0   'False
      DAutoHeadings   =   0   'False
      DAutoSave       =   0   'False
      DAutoSizeCols   =   0
      DInformActiveRowChange=   0   'False
      EditEnterAction =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   16
      MaxRows         =   30
      Protect         =   0   'False
      SpreadDesigner  =   "frm보관서비스.frx":00F8
      UnitType        =   2
      UserResize      =   1
      VisibleCols     =   8
      VisibleRows     =   30
      AppearanceStyle =   0
   End
End
Attribute VB_Name = "frm보관서비스"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private KeyCodeTime As String
Dim m_ActionMode    As String
Dim m_ActionMode2   As String
Dim m_MstCode   As String

Private Sub DisplayTotalCount()
    Dim nRow    As Long
    
    lblTotalCount.Caption = "0"
    
    With fpSpread1
        
        For nRow = 1 To .MaxRows
            .Row = nRow
            .Col = 1
            If .Value = 1 Then
                lblTotalCount.Caption = Val(lblTotalCount.Caption) + 1
            End If
        Next nRow
    End With
    
End Sub


'--------------------------------------------------------------------------------------------------------------
' Procedure : cboSaleDate_Click
' DateTime  : 2006-11-04 01:31
' Author    : pds2004
' Purpose   : 보관계월수를 선택하면 만료 예정일자를 출력해준다.
'               보관 금액 계산
'--------------------------------------------------------------------------------------------------------------
Private Sub cboSaleDate_Click()
    Dim sDate   As String
    
    On Error GoTo cboSaleDate_Click_Error

    sDate = DateAdd("M", Val(Left(cboSaleDate.Text, 2)), Date)
    DTPicker1.Value = sDate
    
    If Val(lblTotalCount.Caption) <= 0 Then Exit Sub
    
    ' 기본 접수 단위 확인
    If Val(lblTotalCount.Caption) < 10 Then
        MsgBox "최소 10벌부터 접수가 가능합니다.", vbInformation, "확인"
        Exit Sub
    ElseIf Val(lblTotalCount.Caption) > 70 Then
        MsgBox "최대 70벌까지 접수가 가능합니다.", vbInformation, "확인"
        Exit Sub
    
    End If
    
    
    Query = "SELECT TOP 1 보관금액 FROM 보관금액 "
    Query = Query & " WHERE 보관월 = '" & Month(Date) & "' "
    Query = Query & "   AND 아이템수 > " & Val(lblTotalCount.Caption) & " "
    Query = Query & "   AND 보관개월수 = " & Val(Left(cboSaleDate.Text, 2)) & " "
    Query = Query & " ORDER BY 아이템수 ASC"
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If SUBRs.EOF Then
        MsgBox "해당 자료의 금액이 없습니다. " & vbLf & vbLf & _
                "본사에서 금액 자료를 다운로드 하여 주십시요.", vbInformation, "확인"
                
        lblTotalCount.Caption = "0"
        lblMoney.Caption = "0"
        
        Exit Sub
    
    Else
        lblMoney.Caption = CStr(Val(SUBRs.Fields("보관금액") & ""))
    End If
    SUBRs.Close
    Set SUBRs = Nothing

    Exit Sub

cboSaleDate_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cboSaleDate_Click of Form Form37"

End Sub

Private Sub cboView_Click()
    Dim sKeyCode As String
    
    sKeyCode = cboView.Text
    
    If InStr(sKeyCode, "보관 접수") <= 0 Then
        cboView.Tag = ""
        cmdAction(0).Enabled = False
        cmdAction(1).Enabled = True
        cmdAction(2).Enabled = True
        
        sKeyCode = Replace(sKeyCode, " ", "")
        sKeyCode = Replace(sKeyCode, "-", "")
        sKeyCode = Replace(sKeyCode, ":", "")
        
        Call Display_INFO(sKeyCode)
        
        
    ' 보관 접수
    Else
        cboView.Tag = ""
        fpSpread1.MaxRows = 0
        cmdAction(0).Enabled = True
        cmdAction(1).Enabled = False
        cmdAction(2).Enabled = False
    
    End If
    
End Sub

Private Sub cmdAction_Click(Index As Integer)
    Dim varTemp As Variant

    Dim sMSG    As String
    Dim nCount  As Long
    
    Dim nTotalCount As Integer
    
    ' 보관 접수
    If Index = 0 Then
        m_ActionMode = "ADD"
        m_ActionMode2 = ""
        
    ' 접수 수정
    ElseIf Index = 1 Then
        m_ActionMode = "EDIT"
        m_ActionMode2 = ""
        
    ' 접수 취소
    ElseIf Index = 2 Then
        m_ActionMode = "DELETE"
        
        ' 선택되지 않은 수를 구한다.
        nCount = GetSpreadSelectCount(False)
        If nCount <> 0 And nCount <= 9 Then
            sMSG = "선택된 항목을 취소할 경우 남은 상품이 10개 미만 이기때문에 취소할수 없습니다."
            MsgBox sMSG, vbInformation
            Exit Sub
        End If
        
        sMSG = "선택된 품목을 삭제 하시겠습니까?"
        If MsgBox(sMSG, vbInformation + vbYesNo, "확인") = vbNo Then Exit Sub
        
        ' 1. 먼저 선택된 내용을 삭제하고
        ' 2. EDIT 모드로 다시 저장한다.
    
        ' 저장 기준 설정
        KeyCodeTime = Replace(cboView.Text, " ", "")
        KeyCodeTime = Replace(KeyCodeTime, "-", "")
        KeyCodeTime = Replace(KeyCodeTime, ":", "")
        
            
        ' 모든 상품을 취소할 경우
        If nCount = 0 Then
            Query = " DELETE FROM TB_보관리스트 "
            Query = Query & " WHERE KeyCode = '" & KeyCodeTime & "' "
            ADOCon.Execute Query
            
            Query = " DELETE FROM 보관상품리스트 "
            Query = Query & " WHERE KeyCode = '" & KeyCodeTime & "' "
            ADOCon.Execute Query
    
            Query = " DELETE FROM TB_보관하자리스트 "
            Query = Query & " WHERE KeyCode = '" & KeyCodeTime & "' "
            ADOCon.Execute Query
            
            Unload Me
            Exit Sub
            
        Else
            ' 선택된 상품별 리스트및 하자 리스트를 모두삭제하고 "EDIT"모드로 다시 수정한다.
            Call Delete보관상품리스트
            
            ' 선택을 반전 시킨다.
            Call Command2_Click
            
            m_ActionMode2 = "DELETE"
            m_ActionMode = "EDIT"
        End If
    End If
 
    ' 전체 수량을 다시 구한다.
    Call DisplayTotalCount
 
    ' 일일 마감 여부를 확인한다.
    Call fpSpread1.GetText(4, 1, varTemp)
    If Fun_일일마감여부(Format(CStr(varTemp), "YYYY-MM-DD")) = True Then
        MsgBox "일일마감이 되었으므로 더이상  보관 서비스를 할수 없습니다.", vbInformation
        Exit Sub
    End If

    ' 필수 입력 사항 확인 (나머지는 콤보 박스로 해놓아서 무조건 선택되게 되어있다.)
    If Len(Trim(txtUserNumber.Text)) <= 0 Then
        MsgBox "주민 등록 번호를 입력하여 주십시요.", vbInformation, "확인"
        txtUserNumber.SetFocus
        Exit Sub
    ElseIf IsJuminNum(txtUserNumber.Text) = False Then
        MsgBox "입력한 주민 등록 번호가 올바르지 않습니다." & vbLf & vbLf & "주민 등록번호를 확인하여 주십시요.", vbInformation, "확인"
        txtUserNumber.SelStart = 0
        txtUserNumber.SelLength = Len(txtUserNumber.Text)
        txtUserNumber.SetFocus
        Exit Sub
    End If
    
    ' 기본 접수 단위 확인
    nTotalCount = Val(lblTotalCount.Caption)
    ' 기본 접수 단위 확인
    If Val(lblTotalCount.Caption) < 10 Then
        MsgBox "최소 10벌부터 접수가 가능합니다.", vbInformation, "확인"
        Exit Sub
    ElseIf Val(lblTotalCount.Caption) > 70 Then
        MsgBox "최대 70벌까지 접수가 가능합니다.", vbInformation, "확인"
        Exit Sub
    
    End If
    
    Query = "SELECT TOP 1 보관금액 FROM TB_보관금액 "
    Query = Query & " WHERE 보관월 = '" & Month(Date) & "' "
    Query = Query & "   AND 아이템수 > " & Val(lblTotalCount.Caption) & " "
    Query = Query & "   AND 보관개월수 = " & Val(Left(cboSaleDate.Text, 2)) & " "
    Query = Query & " ORDER BY 아이템수 ASC"
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If SUBRs.EOF = True Then
        MsgBox "해당 자료의 금액이 없습니다. " & vbLf & vbLf & _
                "본사에서 금액 자료를 다운로드 하여 주십시요.", vbInformation, "확인"
        lblTotalCount.Caption = "0"
        lblMoney.Caption = "0"
        Exit Sub
    
    Else
        lblMoney.Caption = CStr(Val(SUBRs.Fields("보관금액") & ""))
    End If
    SUBRs.Close
    
    ' 필수 입력 내용확인 (스프래드에서)
    If Save보관상품리스트_입력확인 = False Then Exit Sub
    
    
    ' 보관 접수
    If m_ActionMode = "ADD" Then
        sMSG = "[" & CStr(lblTotalCount.Caption) & "]점 " & Format(Val(lblMoney.Caption), "#,##0") & "원에 보관 서비스를 접수 하시겠습니까?"
        If MsgBox(sMSG, vbInformation + vbYesNo, "확인") = vbNo Then Exit Sub
    
        ' 저장 기준 설정
        KeyCodeTime = Format(Now, "YYYY-MM-DD hh:mm:ss")

    ' 접수 수정
    ElseIf m_ActionMode = "EDIT" Then
        If m_ActionMode2 <> "DELETE" Then
            sMSG = "[" & CStr(lblTotalCount.Caption) & "] " & CStr(lblMoney.Caption) & "원에 보관 서비스를 수정 하시겠습니까?"
            If MsgBox(sMSG, vbInformation + vbYesNo, "확인") = vbNo Then Exit Sub
        End If
        ' 저장 기준 설정
        KeyCodeTime = Replace(cboView.Text, " ", "")
        KeyCodeTime = Replace(KeyCodeTime, "-", "")
        KeyCodeTime = Replace(KeyCodeTime, ":", "")
    
    End If
    
    cmdAction(0).Enabled = False
    cmdAction(1).Enabled = False
    cmdAction(2).Enabled = False
    
    
    If Save보관리스트 = False Then
        MsgBox "기본 정보를 저장하는중 오류가 발생하였습니다.", vbInformation
        Command2.Enabled = True
        Exit Sub
    End If
    
    If Save보관상품리스트 = False Then
        MsgBox "상품 정보를 저장하는중 오류가 발생하였습니다.", vbInformation
        Command2.Enabled = True
        Exit Sub
    End If
    
    ' 접수증 프린트
    Bill_Printer = CStr(GetPrtGubun)
    'Printer_BO_Gb = CStr(GetPrtBOGubun)
    
    If m_ActionMode = "ADD" Then
        Call Print_QN_MM(KeyCodeTime)
        MsgBox "접수 완료          ", vbInformation
        
    ElseIf m_ActionMode = "EDIT" Or m_ActionMode = "DELETE" Then
        MsgBox "수정 완료          ", vbInformation
    End If
    
    Command2.Enabled = True
    Unload Me
    
End Sub

Private Sub Command1_Click()
    Dim nCount  As Long
    Dim nRow    As Long
    
    If txtHaData(0).Tag <> "" And txtHaData(1).Tag <> "" Then
        fpSpread1.Col = Val(txtHaData(0).Tag)
        fpSpread1.Row = Val(txtHaData(1).Tag)
        fpSpread1.CellType = CellTypeEdit
        fpSpread1.CellType = CellTypeComboBox
        
        nCount = 0
        For nRow = 0 To 4
            If Trim(txtHaData(nRow).Text) <> "" Then
                fpSpread1.TypeComboBoxString = txtHaData(nRow).Text
                nCount = nCount + 1
            End If
        Next nRow
        
        fpSpread1.Col = Val(txtHaData(0).Tag) - 1
        fpSpread1.Row = Val(txtHaData(1).Tag)
        fpSpread1.Text = CStr(nCount)
    End If
    
    Frame3.Visible = False
End Sub

 
Private Sub Command2_Click()
    Dim nRow As Long
    
    For nRow = 1 To fpSpread1.MaxRows
        fpSpread1.Row = nRow
        fpSpread1.Col = 1
        
        fpSpread1.Value = IIf(fpSpread1.Value = 0, 1, 0)
    Next nRow

End Sub

 

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyChk (KeyCode)
 '  If KeyCode = vbKeyReturn Then
 '     SendKeys "{Tab}"
 '     KeyCode = 0
 '  End If
End Sub

Private Sub Form_Load()
    'FormChk
    'TitleSet "보관 서비스"
    
    cboDevTime.ListIndex = 0
    cboUserGubun.ListIndex = 0
    cboSaleDate.ListIndex = 0
    
    Frame3.Visible = False
    
    fpSpread1.ColWidth(15) = 0
    fpSpread1.ColWidth(16) = 0
    
    ' 접수 현황
    cboView.Clear
    cboView.AddItem " 보관 접수 "
    Call SetComboView(cboView)
    
    cmdAction(1).Enabled = False
    cmdAction(2).Enabled = False
    
    m_MstCode = 가맹점정보.지사코드
    
End Sub

 
Private Sub Frame3_Click()
    Frame3.Visible = False
End Sub

Private Sub mskName_GotFocus(Index As Integer)
    Dim hiMe As Long
    
    Toggle_Check = True
    '//KEYCODE 123 번은 펑션키12번(F12)
    '//특정키를 입력하려면 아래 KEYCODE만 바꿔주면됨
    If Toggle_Check = True Then
        '// 한글로 바꾸기
        hiMe = ImmGetContext(mskName(0).hwnd)
        ImmSetConversionStatus hiMe, IME_HANGUL, IME_NONE
        Toggle_Check = False
    Else
        '// 영어로 바꾸기
        hiMe = ImmGetContext(mskName(0).hwnd)
        ImmSetConversionStatus hiMe, IME_ENGLISH, IME_NONE
        Toggle_Check = True
    End If
    
    mskName(0).SelStart = 0
    mskName(0).SelLength = 10
End Sub

Private Sub mskTEL_GotFocus(Index As Integer)
    mskTEL(Index).SelStart = 0
    mskTEL(Index).SelLength = Len(mskTEL(Index).Text)
End Sub

Private Sub mskTEL_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub mskTEL_LostFocus(Index As Integer)
    Dim Query2 As String
    Dim strBlank1 As String
    Dim strBlank2 As String
    
    
    Select Case Index
        Case 1
            Query = "SELECT    고객코드"
            Query = Query & ", 성명"
            Query = Query & ", 주소"
            Query = Query & ", 전화번호"
            Query = Query & ", 휴대폰 "
            
            Query = Query & " FROM TB_고객정보 "
            Query = Query & " WHERE 전화번호        = '" & mskTEL(1).ClipText & "'"
            Query = Query & "    OR RIGHT(휴대폰,4) = '" & mskTEL(1).ClipText & "' )"
            
            'If mskTEL(0).ClipText <> "" Then
            '    Query = Query & " FROM TB_고객정보 "
            '    Query = Query & " WHERE ( 전화번호 = '" & mskTEL(0).ClipText & "' "
            '    Query = Query & "    OR LEFT(RIGHT(TRIM(휴대폰),9),4) = '" & mskTEL(0).ClipText & "') "
            '    Query = Query & "   AND ( 전화2 = '" & mskTEL(1).ClipText & "' "
            '    Query = Query & "    OR RIGHT(TRIM(휴대폰),4) = '" & mskTEL(1).ClipText & "') "
            'Else
            '    Query = Query & "FROM TB_고객정보 "
            '    Query = Query & "WHERE ( 전화2 = '" & mskTEL(1).ClipText & "'"
            '    Query = Query & "    OR  RIGHT(TRIM(휴대폰),4) = '" & mskTEL(1).ClipText & "' )"
            'End If
            
            Set ADORs = New ADODB.Recordset
            ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            If Not ADORs.EOF Or Not ADORs.BOF Then
                ADORs.MoveLast
            End If
            
            If ADORs.EOF Then
                MsgBox "일치하는 전화번호가 없습니다." & Chr(10) & Chr(10) & "다시 입력하세요!"
                mskTEL(0).SetFocus
                Exit Sub
            End If
            
            
            ' 신규 회원이면
            If ADORs.RecordCount = 1 Then
                If "Error" = Get_고객정보(ADORs!고객코드) Then
                    MsgBox "일치하는 전화번호가 없습니다." & Chr(10) & Chr(10) & "다시 입력하세요!"
                    mskTEL(0).SetFocus
                    Exit Sub
                End If

            ElseIf ADORs.RecordCount > 1 Then
                '뿌리고 입력대기상태
                frm동명이인.DataDisplay Query
                frm동명이인.Show 1
                If frm동명이인.SELECTCODE = "CANCEL" Then
                    mskTEL(1).SetFocus
                    Exit Sub
                End If
            End If
            ADORs.Close
            Set ADORs = Nothing
            
            
            mskTEL(0).Text = 고객정보.전화번호
            mskCode.Text = 고객정보.고객코드
            mskName(0).Text = 고객정보.성명
            
            
            strBlank1 = "확"
            
            '---------------------------------------------------------------------------
            '
            '---------------------------------------------------------------------------
            Query2 = "SELECT     의류코드"
            Query2 = Query2 & ", 의류명"
            Query2 = Query2 & ", 택번호"
            Query2 = Query2 & ", 접수일자"
            Query2 = Query2 & ", 색상"
            Query2 = Query2 & ", 상표"
            Query2 = Query2 & ", 내용"
            Query2 = Query2 & " FROM TB_입출고 "
            Query2 = Query2 & " WHERE 고객코드 = '" & Trim(mskCode.Text) & "' "
            Query2 = Query2 & "   AND 확인    <> '" & strBlank1 & "' "
            Query2 = Query2 & "   AND 접수일자 = '" & Format(Date, "YYYY-MM-DD") & "' "
            Query2 = Query2 & "   AND (판매취소 IS NULL OR 판매취소 <> 'Y') "
            Query2 = Query2 & " ORDER BY 택번호 "
            Set ADORs = New ADODB.Recordset
            ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            Do While Not ADORs.EOF
                i = i + 1
                fpSpread1.MaxRows = i
                fpSpread1.Row = i
                
                fpSpread1.Col = 1:  fpSpread1.Value = 0
                fpSpread1.Col = 2:  fpSpread1.Text = ADORs!의류명 & ""
                fpSpread1.Col = 3:  fpSpread1.Text = ADORs!택번호 & ""
                fpSpread1.Col = 4:  fpSpread1.Text = ADORs!접수일자 & ""
                fpSpread1.Col = 5:  fpSpread1.Text = ADORs!색상 & ""
                fpSpread1.Col = 6:  fpSpread1.Text = ADORs!내용 & ""
                fpSpread1.Col = 9:  fpSpread1.Text = ADORs!상표 & ""
                
                fpSpread1.Col = 15: fpSpread1.Text = ADORs!의류코드 & ""
                
                Call SetComboSpread(7, i, "SIZE_GUBUN") ' 사이즈 코드를 등록한다.
                
                Call SetComboSpread(12, i, "AS") ' AS 코드를 등록한다.
                
                ADORs.MoveNext
            Loop
            ADORs.Close
            Set ADORs = Nothing
    End Select
End Sub

Private Sub txtUserNumber_LostFocus()
    If Len(Trim(txtUserNumber.Text)) > 0 Then
        If IsJuminNum(txtUserNumber.Text) = False Then
            MsgBox "입력한 주민 등록 번호가 올바르지 않습니다." & vbLf & vbLf & "주민 등록번호를 확인하여 주십시요.", vbInformation, "확인"
            txtUserNumber.Text = ""
            txtUserNumber.SetFocus
        End If
    End If
End Sub


Private Sub fpSpread1_Click(ByVal Col As Long, ByVal Row As Long)
    ' 하자 내용일 경우
    If Col = 14 Then
        Dim varTemp     As Variant
        Dim sData(5)    As String
        Dim ii          As Integer
        
        txtHaData(0).Text = ""
        txtHaData(1).Text = ""
        txtHaData(2).Text = ""
        txtHaData(3).Text = ""
        txtHaData(4).Text = ""
        
        fpSpread1.Col = Col
        fpSpread1.Row = Row
        
        txtHaData(0).Tag = Col
        txtHaData(1).Tag = Row
        
        varTemp = fpSpread1.TypeComboBoxList
        If Right(varTemp, 1) = Chr(9) Then varTemp = Left(varTemp, Len(varTemp) - 1)
        varTemp = Split(CStr(varTemp), Chr(9))
        
        ii = 0
        For i = UBound(varTemp) To 0 Step -1
            txtHaData(ii).Text = CStr(varTemp(i))
            ii = ii + 1
        Next i
        
        Frame3.Visible = True
    
    
    End If
End Sub

Private Sub fpSpread1_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
    On Error GoTo ErrRtn

    Dim sData As String
    
    If Col <> 7 Then Exit Sub
    
    With fpSpread1
        .Row = Row:     .Col = Col
        sData = Left(Trim(.Text), 2)
        
        
        ' 해당 내용의 사이즈를 등록하낟.
        Call SetComboSpread(Col + 1, Row, sData)
    
    End With
    
    Exit Sub
ErrRtn:
    MsgBox Err.Description, vbInformation, "확인"

End Sub

Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    
    With fpSpread1
        .Row = .ActiveRow
        .Col = 1
        .Value = 1
        
        ' 전체 수량을 다시 출력한다.
        Call DisplayTotalCount

    End With

End Sub

'--------------------------------------------------------------------------------------------------------------
' Procedure : GetKeyRecordIndex
' DateTime  : 2006-11-04 01:52
' Author    : pds2004
' Purpose   : 전달된 테이플의 KeyCode의 순번을 리턴한다.(최정순번+1)을 리턴
'--------------------------------------------------------------------------------------------------------------
Private Function GetKeyRecordIndex(ByVal sTableName As String, ByVal sKeyCode As String) As String
    Dim sNextIndex  As String

    On Error GoTo GetKeyRecordIndex_Error
    
    If sTableName = "보관리스트" Then
        Query = " SELECT MAX(InputNumber) AS MaxCount "
        Query = Query & "  FROM TB_보관리스트 "
        Query = Query & " WHERE SUBSTRING(InputDate,1,10)  = '" & sKeyCode & "' "
        Set SUBRs = New ADODB.Recordset
        SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        If SUBRs.EOF = True Then
            sNextIndex = "0001"
        ElseIf SUBRs.Fields("MaxCount") & "" = "" Then
            sNextIndex = "0001"
        Else
            sNextIndex = Format(Val(SUBRs.Fields("MaxCount")) + 1, "0000")
        End If
        SUBRs.Close
        
    ElseIf sTableName = "보관상품리스트" Then
        Query = " SELECT MAX(ItemIndex) AS MaxCount "
        Query = Query & "  FROM TB_보관상품리스트 "
        Query = Query & " WHERE KeyCode = '" & sKeyCode & "' "
        Set SUBRs = New ADODB.Recordset
        SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

        If SUBRs.EOF = True Then
            sNextIndex = "000001"
        ElseIf SUBRs.Fields("MaxCount") & "" = "" Then
            sNextIndex = "000001"
        Else
            sNextIndex = Format(Val(SUBRs.Fields("MaxCount")) + 1, "000000")
        End If
        SUBRs.Close
    
    ElseIf sTableName = "보관하자리스트" Then
        Query = " SELECT MAX(ItemIndex) AS MaxCount "
        Query = Query & "  FROM TB_보관하자리스트 "
        Query = Query & " WHERE KeyCode = '" & sKeyCode & "' "
        Set SUBRs = New ADODB.Recordset
        SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

        If SUBRs.EOF = True Then
            sNextIndex = "000001"
        ElseIf SUBRs.Fields("MaxCount") & "" = "" Then
            sNextIndex = "000001"
        Else
            sNextIndex = Format(Val(SUBRs.Fields("MaxCount")) + 1, "000000")
        End If
        SUBRs.Close
    End If
    
    GetKeyRecordIndex = sNextIndex

    Exit Function

GetKeyRecordIndex_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetKeyRecordIndex of Form Form37"
    Resume
End Function


'--------------------------------------------------------------------------------------------------------------
' Procedure : Save보관리스트
' DateTime  : 2006-11-04 01:48
' Author    : pds2004
' Purpose   : 보관 리스트에 저장한다. 성공할 경우 True를 리턴한다.
'--------------------------------------------------------------------------------------------------------------
Private Function Save보관리스트() As Boolean
    On Error GoTo Save보관리스트_Error
    
    Dim Query   As String
    Dim Query2   As String
    Dim sNextIndex  As String
    Dim sMoney  As String
    
    Save보관리스트 = False
    
    If m_ActionMode = "ADD" Then
        sNextIndex = GetKeyRecordIndex("보관리스트", KeyCodeTime)
        If Len(sNextIndex) <> 4 Or Not IsNumeric(sNextIndex) Then
            MsgBox "보관리스트 순번 증가 오류 입니다.", vbInformation, "확인"
            Exit Function
        End If
        
        sMoney = lblMoney.Caption
        
        Query = " INSERT INTO TB_보관리스트 (":     Query2 = " VALUES ( "
        Query = Query & " KeyCode,":                Query2 = Query2 & " '" & KeyCodeTime & "', "
        Query = Query & " MemRecord,":              Query2 = Query2 & " '" & "FO" & "', "
        Query = Query & " InputNumber,":            Query2 = Query2 & " '" & sNextIndex & "', "
        Query = Query & " InputDate,":              Query2 = Query2 & " '" & Format(Now, "YYYY-MM-DD hh:mm:ss000") & "', "
        Query = Query & " InputID,":                Query2 = Query2 & " '" & Trim(mskCode.Text) & "', "
        Query = Query & " InputName,":              Query2 = Query2 & " '" & Trim(mskName(0).Text) & "', "
        Query = Query & " EMail,":                  Query2 = Query2 & " '" & Trim(txtEMail.Text) & "', "
        Query = Query & " UserCode,":               Query2 = Query2 & " '" & Left(cboUserGubun.Text, 2) & "', "
        Query = Query & " UserNumber,":             Query2 = Query2 & " '" & Trim(Replace(txtUserNumber.Text, ",", "")) & "', "
        Query = Query & " 가맹점코드,":              Query2 = Query2 & " '" & Trim(m_MstCode) & "," & Trim(가맹점정보.택코드) & "', "
        Query = Query & " SaleGubunCode,":          Query2 = Query2 & " '" & Left(cboSaleDate.Text, 2) & "', "
        Query = Query & " SaleEndDate,":            Query2 = Query2 & " '" & Format(DTPicker1.Value, "YYYY-MM-DD") & "', "
        Query = Query & " Price,":                  Query2 = Query2 & " '" & Format(sMoney, "00000000") & "', "
        Query = Query & " DevTimeCode,":            Query2 = Query2 & " '" & Left(cboDevTime.Text, 2) & "', "
        Query = Query & " ItemCount,":              Query2 = Query2 & " '" & Trim(lblTotalCount.Caption) & "', "
        Query = Query & " StatsFlag)":              Query2 = Query2 & " '" & " " & "') "
        Query = Query & Query2
        ADOCon.Execute Query
        
    ElseIf m_ActionMode = "EDIT" Then
    
        sNextIndex = cboView.Tag
        sMoney = lblMoney.Caption
        
        
        Query = " UPDATE TB_보관리스트 SET "
        Query = Query & " EMail         = '" & Trim(txtEMail.Text) & "', "
        Query = Query & " UserCode      = '" & Left(cboUserGubun.Text, 2) & "', "
        Query = Query & " UserNumber    = '" & Trim(Replace(txtUserNumber.Text, ",", "")) & "', "
        Query = Query & " SaleGubunCode = '" & Left(cboSaleDate.Text, 2) & "', "
        Query = Query & " SaleEndDate   = '" & Format(DTPicker1.Value, "YYYY-MM-DD") & "', "
        Query = Query & " Price         = '" & Format(sMoney, "00000000") & "', "
        Query = Query & " DevTimeCode   = '" & Left(cboDevTime.Text, 2) & "', "
        Query = Query & " ItemCount     = '" & Trim(lblTotalCount.Caption) & "' "
        Query = Query & " WHERE KeyCode     = '" & KeyCodeTime & "' "
        Query = Query & "   AND InputNumber = '" & sNextIndex & "' "
        ADOCon.Execute Query
    
    End If
    Save보관리스트 = True
    
    Exit Function

Save보관리스트_Error:

    Save보관리스트 = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Save보관리스트 of Form Form37"

End Function

'--------------------------------------------------------------------------------------------------------------
' Procedure : Save보관상품리스트
' DateTime  : 2006-11-04 01:48
' Author    : pds2004
' Purpose   : 보관 상품 리스트에 저장한다. 성공할 경우 True를 리턴한다.
'--------------------------------------------------------------------------------------------------------------
Private Function Save보관상품리스트_입력확인() As Boolean
    On Error GoTo ErrRtn
    
    Dim sNextIndex  As String
    Dim sMoney  As String
    
    Dim nRow    As Long
    Dim sData(10) As String
    
    Save보관상품리스트_입력확인 = False
    For nRow = 1 To fpSpread1.MaxRows
        
        fpSpread1.Row = nRow:       fpSpread1.Col = 1
        
        If fpSpread1.Value = 1 Then
    
            ' 자료 저장을 위하여 각종 자료를 구한다.
            fpSpread1.Row = nRow
            fpSpread1.Col = 3:      sData(0) = Replace(fpSpread1.Text, "-", "")     ' 택번호
            fpSpread1.Col = 15:     sData(1) = fpSpread1.Text       ' 상품코드
            fpSpread1.Col = 7:      sData(2) = fpSpread1.Text       ' SIZE구분
            fpSpread1.Col = 8:      sData(3) = Left(fpSpread1.Text, 2)      ' SIZE구분2
            fpSpread1.Col = 5:      sData(4) = fpSpread1.Text       ' 색상
            fpSpread1.Col = 9:      sData(5) = fpSpread1.Text       ' 브랜드 명
            fpSpread1.Col = 10:     sData(6) = fpSpread1.Text       ' 구입가격
            fpSpread1.Col = 11:     sData(7) = fpSpread1.Text       ' 구입일자
            fpSpread1.Col = 12:     sData(8) = Left(fpSpread1.Text, 2)      ' AS여부
            fpSpread1.Col = 13:     sData(9) = fpSpread1.Text       ' 하자개수
            
            If sData(9) = "" Then sData(9) = "0"
            
            If sData(0) = "" Then
                MsgBox "[" & CStr(nRow) & "]줄의 택번호가 입력되어 있지 않습니다. 확인하여 주십시요.", vbInformation
                Exit Function
            End If
            
            If sData(1) = "" Then
                MsgBox "[" & CStr(nRow) & "]줄의 상품코드가 입력되어 있지 않습니다. 확인하여 주십시요.", vbInformation
                Exit Function
            End If
            If sData(2) = "" Then
'                MsgBox "[" & CStr(nRow) & "]줄의 SIZE 구분이 입력되어 있지 않습니다. 확인하여 주십시요.", vbInformation
'                Exit Function
            End If
            If sData(3) = "" Then
'                MsgBox "[" & CStr(nRow) & "]줄의 SIZE 구분2가 입력되어 있지 않습니다. 확인하여 주십시요.", vbInformation
'                Exit Function
            End If
            If sData(4) = "" Then
                MsgBox "[" & CStr(nRow) & "]줄의 색상이 입력되어 있지 않습니다. 확인하여 주십시요.", vbInformation
                Exit Function
            End If
            If sData(5) = "" Then
                MsgBox "[" & CStr(nRow) & "]줄의 브랜드가 입력되어 있지 않습니다. 확인하여 주십시요.", vbInformation
                Exit Function
            End If
            If sData(6) = "" Then
'                MsgBox "[" & CStr(nRow) & "]줄의 구입가격이 입력되어 있지 않습니다. 확인하여 주십시요.", vbInformation
'                Exit Function
            End If
            If sData(7) = "" Then
'                MsgBox "[" & CStr(nRow) & "]줄의 구입일자가 입력되어 있지 않습니다. 확인하여 주십시요.", vbInformation
'                Exit Function
            End If
            If sData(8) = "" Then
'                MsgBox "[" & CStr(nRow) & "]줄의 AS 가능여부가 입력되어 있지 않습니다. 확인하여 주십시요.", vbInformation
'                Exit Function
            End If
            If sData(9) = "" Then
'                MsgBox "[" & CStr(nRow) & "]줄의 하자개수가 입력되어 있지 않습니다. 확인하여 주십시요.", vbInformation
'                Exit Function
            End If
            

            sData(0) = "": sData(1) = "": sData(2) = "": sData(0) = "": sData(4) = "": sData(5) = ""
            sData(6) = "": sData(7) = "": sData(8) = "": sData(9) = ""
        End If
    Next
    
    Save보관상품리스트_입력확인 = True
    
    On Error GoTo 0
    Exit Function

ErrRtn:
    Save보관상품리스트_입력확인 = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Save보관상품리스트_입력확인 of Form Form37"
End Function

'--------------------------------------------------------------------------------------------------------------
' Procedure : Save보관상품리스트
' DateTime  : 2006-11-04 01:48
' Author    : pds2004
' Purpose   : 보관 상품 리스트에 저장한다. 성공할 경우 True를 리턴한다.
'--------------------------------------------------------------------------------------------------------------
Private Function Save보관상품리스트() As Boolean
    On Error GoTo ErrRtn
    
    Dim Query   As String
    Dim Query2   As String
    Dim sNextIndex  As String
    Dim sMoney  As String
    
    Dim nRow    As Long
    Dim sData(10) As String
    
    Save보관상품리스트 = False
    For nRow = 1 To fpSpread1.MaxRows
        
        fpSpread1.Row = nRow:       fpSpread1.Col = 1
        
        If fpSpread1.Value = 1 Then
    
            ' 자료 저장을 위하여 각종 자료를 구한다.
            fpSpread1.Row = nRow
            fpSpread1.Col = 3:      sData(0) = Replace(fpSpread1.Text, "-", "")     ' 택번호
            fpSpread1.Col = 15:     sData(1) = fpSpread1.Text                       ' 상품코드
            fpSpread1.Col = 7:      sData(2) = fpSpread1.Text                       ' SIZE구분
            fpSpread1.Col = 8:      sData(3) = Left(fpSpread1.Text, 2)              ' SIZE구분2
            fpSpread1.Col = 5:      sData(4) = fpSpread1.Text                       ' 색상
            fpSpread1.Col = 9:      sData(5) = fpSpread1.Text                       ' 브랜드 명
            fpSpread1.Col = 10:     sData(6) = Replace(fpSpread1.Text, ",", "")                ' 구입가격
            fpSpread1.Col = 11:     sData(7) = Format(fpSpread1.Text, "YYYY-MM-DD")   ' 구입일자
            fpSpread1.Col = 12:     sData(8) = Left(fpSpread1.Text, 2)              ' AS여부
            fpSpread1.Col = 13:     sData(9) = fpSpread1.Text                       ' 하자 개수
            If sData(9) = "" Then sData(9) = "0"
            
            If m_ActionMode = "ADD" Then
                
                ' 품목 순번을 구해온다.
                sNextIndex = GetKeyRecordIndex("보관상품리스트", KeyCodeTime)
                If Len(sNextIndex) <> 6 Or Not IsNumeric(sNextIndex) Then
                    MsgBox "보관상품리스트 순번 증가 오류 입니다.", vbInformation, "확인"
                    Exit Function
                End If
                
                Query = " INSERT INTO TB_보관상품리스트 (":    Query2 = " VALUES ( "
                Query = Query & " KeyCode,":                Query2 = Query2 & " '" & KeyCodeTime & "', "
                Query = Query & " ItemRecord,":             Query2 = Query2 & " '" & "CI" & "', "
                Query = Query & " ItemIndex,":              Query2 = Query2 & " '" & sNextIndex & "', "
                Query = Query & " InputDate,":              Query2 = Query2 & " '" & Format(Date, "YYYY-MM-DD") & "', "
                Query = Query & " TAG,":                    Query2 = Query2 & " '" & sData(0) & "', "
                Query = Query & " GoodsCode,":              Query2 = Query2 & " '" & sData(1) & "', "
                Query = Query & " SizeGubun,":              Query2 = Query2 & " '" & sData(2) & "', "
                Query = Query & " SizeCode,":               Query2 = Query2 & " '" & sData(3) & "', "
                Query = Query & " Color,":                  Query2 = Query2 & " '" & sData(4) & "', "
                Query = Query & " BrandName,":              Query2 = Query2 & " '" & sData(5) & "', "
                Query = Query & " BuyPrice,":               Query2 = Query2 & " '" & sData(6) & "', "
                Query = Query & " BuyDate,":                Query2 = Query2 & " '" & sData(7) & "', "
                Query = Query & " ASGubun,":                Query2 = Query2 & " '" & sData(8) & "', "
                Query = Query & " BleCount,":               Query2 = Query2 & " '" & sData(9) & "', "
                Query = Query & " StatsFlag)":              Query2 = Query2 & " '" & " " & "') "
                Query = Query & Query2
                ADOCon.Execute Query
            
                If sData(9) <> "0" Then
                    Save보관하자리스트 nRow, 14, sNextIndex
                End If
            
            ElseIf m_ActionMode = "EDIT" Then
                
                fpSpread1.Row = nRow
                fpSpread1.Col = 16:      sNextIndex = fpSpread1.Text
                
                Query = " UPDATE TB_보관상품리스트  SET "
                Query = Query & " SizeGubun = '" & sData(2) & "', "
                Query = Query & " SizeCode = '" & sData(3) & "', "
                Query = Query & " BrandName = '" & sData(5) & "', "
                Query = Query & " BuyPrice = '" & sData(6) & "', "
                Query = Query & " BuyDate = '" & sData(7) & "', "
                Query = Query & " ASGubun = '" & sData(8) & "', "
                Query = Query & " BleCount = '" & sData(9) & "' "
                Query = Query & " WHERE KeyCode = '" & KeyCodeTime & "'  "
                Query = Query & "   AND ItemIndex = '" & sNextIndex & "'  "
                ADOCon.Execute Query
                
                If sData(9) <> "0" Then
                    Save보관하자리스트 nRow, 14, sNextIndex
                End If
            
            
            End If
            
            sData(0) = "": sData(1) = "": sData(2) = "": sData(0) = "": sData(4) = "": sData(5) = ""
            sData(6) = "": sData(7) = "": sData(8) = "": sData(9) = ""
        End If
    Next
    
    Save보관상품리스트 = True
    
    Exit Function

ErrRtn:

    Save보관상품리스트 = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Save보관상품리스트 of Form Form37"
    Resume
End Function


'--------------------------------------------------------------------------------------------------------------
' Procedure : Delete보관상품리스트
' DateTime  : 2006-11-04 01:48
' Author    : pds2004
' Purpose   : 보관 상품 리스트에 해당 내용을 삭제한다. 성공할 경우 True를 리턴한다.
'--------------------------------------------------------------------------------------------------------------
Private Function Delete보관상품리스트() As Boolean
    On Error GoTo ErrRtn
    
    Dim Query   As String
    Dim Query2   As String
    Dim sNextIndex  As String
    Dim sMoney  As String
    
    Dim nRow    As Long
    Dim sData(10) As String
    
    Delete보관상품리스트 = False
    For nRow = 1 To fpSpread1.MaxRows
        
        fpSpread1.Row = nRow:       fpSpread1.Col = 1
        
        If fpSpread1.Value = 1 Then
            
            If m_ActionMode = "DELETE" Then
                
                ' 품목 순번을 구해온다.
                fpSpread1.Row = nRow
                fpSpread1.Col = 16:      sNextIndex = fpSpread1.Text
                
                Query = " UPDATE TB_보관상품리스트 SET "
                Query = Query & " StatsFlag = 'C' "
                Query = Query & " WHERE KeyCode = '" & KeyCodeTime & "' "
                Query = Query & "   AND ItemIndex = '" & sNextIndex & "' "
                ADOCon.Execute Query
                
                Query = " UPDATE TB_보관하자리스트 SET "
                Query = Query & " StatsFlag = 'C' "
                Query = Query & " WHERE KeyCode = '" & KeyCodeTime & "'  "
                Query = Query & "   AND ItemIndex = '" & sNextIndex & "' "
                ADOCon.Execute Query
            End If
        End If
    Next
    
    Delete보관상품리스트 = True
    
    Exit Function

ErrRtn:

    Delete보관상품리스트 = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Delete보관상품리스트 of Form Form37"
    Resume
End Function

Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If Col = 1 Then
        Debug.Print fpSpread1.Value
    End If

End Sub


'--------------------------------------------------------------------------------------------------------------
' Procedure : Save보관리스트
' DateTime  : 2006-11-04 01:48
' Author    : pds2004
' Purpose   : 보관 하자리스트에 저장한다. 성공할 경우 True를 리턴한다.
'--------------------------------------------------------------------------------------------------------------
Private Function Save보관하자리스트(ByVal nRow As Long, ByVal nCol As Long, ByVal sNextIndex As String) As Boolean
    On Error GoTo ErrRtn
    
    Dim Query   As String
    Dim Query2   As String
    Dim varTemp As Variant
    
    fpSpread1.Col = nCol:   fpSpread1.Row = nRow
    varTemp = fpSpread1.TypeComboBoxList
    If Right(varTemp, 1) = Chr(9) Then varTemp = Left(varTemp, Len(varTemp) - 1)
    varTemp = Split(CStr(varTemp), Chr(9))
    
    For i = 0 To UBound(varTemp)
        Save보관하자리스트 = False
        
        If m_ActionMode = "EDIT" Then
            Query = " DELETE FROM TB_보관하자리스트 "
            Query = Query & " WHERE KeyCode = '" & KeyCodeTime & "'  "
            Query = Query & "   AND ItemIndex = '" & sNextIndex & "' "
            
            ADOCon.Execute Query
        End If
        
        Query = " INSERT INTO TB_보관하자리스트 (":    Query2 = " VALUES ( "
        Query = Query & " KeyCode,":                Query2 = Query2 & " '" & KeyCodeTime & "', "
        Query = Query & " InputDate,":              Query2 = Query2 & " '" & Format(Date, "YYYY-MM-DD") & "', "
        Query = Query & " ItemIndex,":              Query2 = Query2 & " '" & sNextIndex & "', "
        Query = Query & " ItemCount,":              Query2 = Query2 & " '" & Format(i + 1, "00") & "', "
        Query = Query & " ItemRemark,":             Query2 = Query2 & " '" & CStr(varTemp(UBound(varTemp) - i)) & "', "
        Query = Query & " StatsFlag)":              Query2 = Query2 & " '" & " " & "') "
        Query = Query & Query2
        ADOCon.Execute Query
    
    
    Next i
    Save보관하자리스트 = True
    
    On Error GoTo 0
    Exit Function

ErrRtn:

    Save보관하자리스트 = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Save보관리스트 of Form Form37"

End Function



'--------------------------------------------------------------------------------------------------------------
' Procedure : SetComboView
' DateTime  : 2006-11-09 14:30
' Author    : pds2004
' Purpose   : 접수된 내용을 설정한다.
'--------------------------------------------------------------------------------------------------------------
Public Function SetComboView(ByRef cboView As ComboBox) As Boolean

    Dim bResult As Boolean

    On Error GoTo ErrRtn

    Query = "SELECT KeyCode FROM TB_보관리스트 "
    Query = Query & " WHERE SUBSTRING(KeyCode,1,10) = '" & Format(Date, "YYYY-MM-DD") & "' "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    Do While Not SUBRs.EOF
        cboView.AddItem Format(SUBRs.Fields("KeyCode"), "YYYY-MM-DD @@:@@:@@")
        SUBRs.MoveNext
    Loop
    SUBRs.Close
    
    
    SetComboView = True

    On Error GoTo 0
    Exit Function

ErrRtn:
    SetComboView = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SetComboView of Form Form37"

End Function

'--------------------------------------------------------------------------------------------------------------
' Procedure : ComboSelectText
' DateTime  : 2006-11-09 14:59
' Author    : pds2004
' Purpose   : 전달된 문자의 길이만큼 왼쪽에서 검색하여 처음으로 같은것을 선택하낟.
'--------------------------------------------------------------------------------------------------------------
Private Function ComboSelectText(ByRef cboView As ComboBox, ByVal sSelectText As String) As Integer
    Dim iSelLen As Integer
    
    iSelLen = Len(sSelectText)
    
    For i = 0 To cboView.ListCount
        If Left(cboView.List(i), iSelLen) = sSelectText Then
            cboView.ListIndex = i
            ComboSelectText = cboView.ListIndex
            Exit Function
        End If
    Next i
    
End Function

Private Function ComboSpreadSelectText(ByVal Col As Long, ByVal Row As Long, ByVal sSelectText As String) As String
        
        Dim varTemp     As Variant
        Dim sData(5)    As String
        Dim ii          As Integer
        Dim nSelLen     As Integer
        
        fpSpread1.Col = Col
        fpSpread1.Row = Row
        
        varTemp = fpSpread1.TypeComboBoxList
        If Right(varTemp, 1) = Chr(9) Then varTemp = Left(varTemp, Len(varTemp) - 1)
        varTemp = Split(CStr(varTemp), Chr(9))
        
        ii = 0
        nSelLen = Len(sSelectText)
        For i = UBound(varTemp) To 0 Step -1
            
            If Left(CStr(varTemp(i)), nSelLen) = sSelectText Then
                fpSpread1.Text = CStr(varTemp(i))
                Exit Function
            End If
            ii = ii + 1
        Next i
        

End Function

'--------------------------------------------------------------------------------------------------------------
' Procedure : SetComboSpread
' DateTime  : 2006-11-09 15:27
' Author    : pds2004
' Purpose   : 전달된 시트의 해당 열에 해당 내용을 추가한다.
'--------------------------------------------------------------------------------------------------------------
Public Function SetComboSpread(ByVal nCol As Long, ByVal nRow As Long, ByVal sMode As String) As Boolean

    Dim bResult As Boolean


    On Error GoTo SetComboSpread_Error

    With fpSpread1
        .Col = nCol
        .Row = nRow
                
        .CellType = 1 '
        .CellType = 8 ' SS_CELL_TYPE_COMBOBOX

        If UCase(sMode) = "SIZE_GUBUN" Then
            .TypeComboBoxString = "C3"
            .TypeComboBoxString = "C2"
            .TypeComboBoxString = "C1"
            .TypeComboBoxString = "F1"
            
        ElseIf UCase(sMode) = "F1" Then
            .TypeComboBoxString = "05. 88"
            .TypeComboBoxString = "04. 77"
            .TypeComboBoxString = "03. 66"
            .TypeComboBoxString = "02. 55"
            .TypeComboBoxString = "01. 44"

        ElseIf UCase(sMode) = "C1" Then
            .TypeComboBoxString = "07. XXL"
            .TypeComboBoxString = "06. XL"
            .TypeComboBoxString = "05. L"
            .TypeComboBoxString = "04. M"
            .TypeComboBoxString = "03. S"
            .TypeComboBoxString = "02. XS"
            .TypeComboBoxString = "01. XXS"
        
        ElseIf UCase(sMode) = "C2" Then
            .TypeComboBoxString = "99. 43 Inch 이상"
            .TypeComboBoxString = "43. 43 Inch"
            .TypeComboBoxString = "42. 42 Inch"
            .TypeComboBoxString = "41. 41 Inch"
            .TypeComboBoxString = "40. 40 Inch"
            .TypeComboBoxString = "39. 39 Inch"
            .TypeComboBoxString = "38. 38 Inch"
            .TypeComboBoxString = "37. 37 Inch"
            .TypeComboBoxString = "36. 36 Inch"
            .TypeComboBoxString = "35. 35 Inch"
            .TypeComboBoxString = "34. 34 Inch"
            .TypeComboBoxString = "33. 33 Inch"
            .TypeComboBoxString = "32. 32 Inch"
            .TypeComboBoxString = "31. 31 Inch"
            .TypeComboBoxString = "30. 30 Inch"
            .TypeComboBoxString = "29. 29 Inch"
            .TypeComboBoxString = "28. 28 Inch"
            .TypeComboBoxString = "27. 27 Inch"
            .TypeComboBoxString = "01. 26 Inch 이하"

        ElseIf UCase(sMode) = "C3" Then
            .TypeComboBoxString = "06. 110"
            .TypeComboBoxString = "05. 105"
            .TypeComboBoxString = "04. 100"
            .TypeComboBoxString = "03. 95"
            .TypeComboBoxString = "02. 90"
            .TypeComboBoxString = "01. 85"
        
        
        ElseIf UCase(sMode) = "AS" Then
            .TypeComboBoxString = "01. 가능"
            .TypeComboBoxString = "00. 불가능"
        
        End If
        
    End With

    SetComboSpread = bResult

    On Error GoTo 0
    Exit Function

SetComboSpread_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SetComboSpread of Form Form37"

End Function


Private Function SetComboSpread_Bel(ByVal nCol As Long, ByVal nRow As Long, ByVal sKeyCode As String, ByVal sItemIndex As String) As Integer
    Dim nCount  As Long
    
    Query = "SELECT * FROM TB_보관하자리스트 "
    Query = Query & " WHERE KeyCode = '" & sKeyCode & "' "
    Query = Query & "   AND ItemIndex = '" & sItemIndex & "' "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    fpSpread1.Col = nCol
    fpSpread1.Row = nRow
    fpSpread1.CellType = CellTypeEdit
    fpSpread1.CellType = CellTypeComboBox
    nCount = 0
    
    Do While Not SUBRs.EOF
        nCount = nCount + 1
        fpSpread1.TypeComboBoxString = Trim(SUBRs.Fields("ItemRemark") & "")
        SUBRs.MoveNext
    Loop
    SUBRs.Close
    
    
    fpSpread1.Text = CStr(nCount)

End Function

'--------------------------------------------------------------------------------------------------------------
' Procedure : Display_Cust
' DateTime  : 2006-11-09 14:49
' Author    : pds2004
' Purpose   : 전달된 내용의 기준 정보를 출력한다.
'--------------------------------------------------------------------------------------------------------------
Public Function Display_INFO(ByVal sKeyCode As String) As Boolean
    Dim bResult As Boolean
    Dim nRow    As Long
    
    On Error GoTo ErrRtn
    
    Query = "SELECT * FROM TB_보관리스트 "
    Query = Query & " WHERE KeyCode = '" & sKeyCode & "' "
    Query = Query & "   AND StatsFlag <> 'C' "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If SUBRs.EOF = True Then
        SUBRs.Close
        Set SUBRs = Nothing
        
        Display_INFO = False
        Exit Function
    End If
    
    Call Get_고객정보(Trim(SUBRs.Fields("InputID")))
    
    If 고객정보.고객코드 <> "Error" Then
        mskCode.Text = 고객정보.고객코드
        mskTEL(0).Text = 고객정보.전화번호
        mskName(0).Text = 고객정보.성명
    End If
    
    cboView.Tag = SUBRs.Fields("InputNumber") & ""
    
    Call ComboSelectText(cboUserGubun, SUBRs.Fields("UserCode") & "")
    Call ComboSelectText(cboSaleDate, SUBRs.Fields("SaleGubunCode") & "")
    Call ComboSelectText(cboDevTime, SUBRs.Fields("DevTimeCode") & "")
    
    txtUserNumber.Text = SUBRs.Fields("UserNumber") & ""
    DTPicker1.Value = Format(SUBRs.Fields("SaleEndDate") & "", "YYYY-MM-DD")
    txtEMail.Text = SUBRs.Fields("EMail") & ""
    
    SUBRs.Close
    Set SUBRs = Nothing
    
    '------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------
    Query = "SELECT * FROM TB_보관상품리스트 "
    Query = Query & " WHERE KeyCode = '" & sKeyCode & "' "
    Query = Query & "   AND StatsFlag <> 'C' "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With fpSpread1
        nRow = 1
        Do Until SUBRs.EOF
            .MaxRows = nRow
            .Row = nRow
            
            .Col = 2:  .Text = GetGoodsName(SUBRs.Fields("GoodsCode") & "")
            .Col = 3:  .Text = Format(SUBRs.Fields("TAG") & "", "@@-@@@")
            .Col = 4:  .Text = Format(Left(SUBRs.Fields("InputDate") & "", 8), "YYYY-MM-DD")
            .Col = 5:  .Text = SUBRs.Fields("Color") & ""
            .Col = 6:  .Text = ""
            
            Call SetComboSpread(7, nRow, "SIZE_GUBUN")
            .Col = 7:  .Text = SUBRs.Fields("SizeGubun") & ""
            
            Call SetComboSpread(8, nRow, SUBRs.Fields("SizeGubun") & "")
            Call ComboSpreadSelectText(8, nRow, SUBRs.Fields("SizeCode") & "")
            
            .Col = 9:  .Text = SUBRs.Fields("BrandName") & ""
            .Col = 10:  .Text = SUBRs.Fields("BuyPrice") & ""
            .Col = 11:  .Text = SUBRs.Fields("BuyDate") & ""
            
            Call SetComboSpread(12, nRow, "AS")
            Call ComboSpreadSelectText(12, nRow, SUBRs.Fields("ASGubun") & "")
            
            .Col = 13:  .Text = SUBRs.Fields("BleCount") & ""
            
            If Val(SUBRs.Fields("BleCount") & "") > 0 Then
                Call SetComboSpread_Bel(14, nRow, SUBRs.Fields("KeyCode") & "", SUBRs.Fields("ItemIndex") & "")
            End If
            
            .Col = 15:  .Text = SUBRs.Fields("GoodsCode") & ""
            .Col = 16:  .Text = SUBRs.Fields("ItemIndex") & ""
        
            nRow = nRow + 1
            
            SUBRs.MoveNext
        Loop
        SUBRs.Close
        Set SUBRs = Nothing
    End With
    
    Display_INFO = bResult

    Exit Function

ErrRtn:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Display_INFO of Form Form37"
    Resume
End Function

'--------------------------------------------------------------------------------------------------------------
' Procedure : GetSpreadSelectCount
' DateTime  : 2006-11-11 03:57
' Author    : pds2004
' Purpose   : bSelectMode = true 선택된 갯수를 리턴한다.
'             bSelectMode = False 선택되지 않은 수를 리턴한다.
'--------------------------------------------------------------------------------------------------------------
Private Function GetSpreadSelectCount(bSelectMode As Boolean) As Long
    Dim nRow    As Long
    Dim nSelect As Long
    
    With fpSpread1
        nSelect = 0
        
        For nRow = 1 To .MaxRows
            .Row = nRow
            .Col = 1
            If bSelectMode = True And .Value = 1 Then
                nSelect = nSelect + 1
            ElseIf bSelectMode = False And .Value = 0 Then
                nSelect = nSelect + 1
            End If
        Next nRow
    End With
    
    GetSpreadSelectCount = nSelect

End Function
