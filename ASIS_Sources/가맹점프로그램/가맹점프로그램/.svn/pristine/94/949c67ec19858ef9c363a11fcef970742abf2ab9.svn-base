VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frmINIT 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "대리점 정보수정"
   ClientHeight    =   8355
   ClientLeft      =   5085
   ClientTop       =   2895
   ClientWidth     =   8220
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
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8355
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8220
      _ExtentX        =   14499
      _ExtentY        =   14737
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frmINIT.frx":0000
      Begin Threed.SSPanel SSPanel3 
         Height          =   660
         Left            =   15
         TabIndex        =   1
         Top             =   7680
         Width           =   8190
         _ExtentX        =   14446
         _ExtentY        =   1164
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   1455
            Top             =   15
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin XtremeSuiteControls.PushButton cmdSave 
            Height          =   570
            Left            =   45
            TabIndex        =   2
            Top             =   45
            Width           =   1260
            _Version        =   851970
            _ExtentX        =   2222
            _ExtentY        =   1005
            _StockProps     =   79
            Caption         =   "저장"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdCancel 
            Height          =   570
            Left            =   6885
            TabIndex        =   3
            Top             =   45
            Width           =   1260
            _Version        =   851970
            _ExtentX        =   2222
            _ExtentY        =   1005
            _StockProps     =   79
            Caption         =   "취소"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.TabControl TabControl1 
         Height          =   7650
         Left            =   15
         TabIndex        =   4
         Top             =   15
         Width           =   8190
         _Version        =   851970
         _ExtentX        =   14446
         _ExtentY        =   13494
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   10
         Color           =   32
         PaintManager.Layout=   5
         PaintManager.Position=   1
         PaintManager.ShowIcons=   -1  'True
         PaintManager.LargeIcons=   -1  'True
         ItemCount       =   4
         Item(0).Caption =   " 기본정보 "
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "TabControlPage1"
         Item(1).Caption =   " 마    진 "
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "TabControlPage2"
         Item(2).Caption =   " 프 린 트 "
         Item(2).ControlCount=   1
         Item(2).Control(0)=   "TabControlPage3"
         Item(3).Caption =   " 접    속 "
         Item(3).ControlCount=   1
         Item(3).Control(0)=   "TabControlPage4"
         Begin XtremeSuiteControls.TabControlPage TabControlPage4 
            Height          =   7590
            Left            =   -68890
            TabIndex        =   81
            Top             =   30
            Visible         =   0   'False
            Width           =   7050
            _Version        =   851970
            _ExtentX        =   12435
            _ExtentY        =   13388
            _StockProps     =   1
            BackColor       =   16777215
            Page            =   3
            Begin VB.TextBox txtSMSUserPass 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   2280
               TabIndex        =   124
               Top             =   3195
               Width           =   4125
            End
            Begin VB.TextBox txtSMSUserName 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   2280
               TabIndex        =   123
               Top             =   2760
               Width           =   4125
            End
            Begin VB.TextBox txtSMSDBName 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   2280
               TabIndex        =   122
               Top             =   2325
               Width           =   4125
            End
            Begin VB.TextBox txtSMSIPAddress 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   2280
               TabIndex        =   121
               Top             =   1890
               Width           =   4125
            End
            Begin VB.TextBox txtFilePort 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   3525
               TabIndex        =   120
               Top             =   1455
               Width           =   1215
            End
            Begin VB.TextBox txtMsgPort 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   2280
               TabIndex        =   119
               Top             =   1455
               Width           =   1215
            End
            Begin VB.TextBox txtIPAddress 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   2280
               TabIndex        =   118
               Top             =   1020
               Width           =   4125
            End
            Begin VB.TextBox txtTel2 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   3105
               TabIndex        =   117
               Top             =   585
               Width           =   1665
            End
            Begin VB.TextBox txtTel1 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   2280
               TabIndex        =   116
               Top             =   585
               Width           =   795
            End
            Begin Threed.SSPanel SSPanel2 
               Height          =   420
               Left            =   2280
               TabIndex        =   82
               Top             =   150
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   741
               _Version        =   262144
               BackColor       =   16777215
               BorderWidth     =   0
               BevelOuter      =   1
               BevelInner      =   2
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin VB.OptionButton OptConnect 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "인터넷"
                  BeginProperty Font 
                     Name            =   "맑은 고딕"
                     Size            =   11.25
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   1
                  Left            =   1095
                  TabIndex        =   84
                  Top             =   45
                  Value           =   -1  'True
                  Width           =   1005
               End
               Begin VB.OptionButton OptConnect 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "모뎀"
                  BeginProperty Font 
                     Name            =   "맑은 고딕"
                     Size            =   11.25
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   0
                  Left            =   75
                  TabIndex        =   83
                  Top             =   45
                  Width           =   855
               End
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   17
               Left            =   150
               TabIndex        =   85
               Top             =   150
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "전송방법 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":0052
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   18
               Left            =   150
               TabIndex        =   86
               Top             =   585
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "전화번호 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":0394
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   19
               Left            =   150
               TabIndex        =   87
               Top             =   1020
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "파일 서버 IP "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":06D6
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   20
               Left            =   150
               TabIndex        =   88
               Top             =   1455
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "파일 서버 Port "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":0A18
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   21
               Left            =   150
               TabIndex        =   89
               Top             =   1890
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "SMS 서버 IP "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":0D5A
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   22
               Left            =   150
               TabIndex        =   90
               Top             =   2325
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "SMS DB 이름 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":109C
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   23
               Left            =   150
               TabIndex        =   91
               Top             =   2760
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "SMS 사용자명 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":13DE
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   24
               Left            =   150
               TabIndex        =   92
               Top             =   3195
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "SMS 사용자 암호 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":1720
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage3 
            Height          =   7590
            Left            =   -68890
            TabIndex        =   5
            Top             =   30
            Visible         =   0   'False
            Width           =   7050
            _Version        =   851970
            _ExtentX        =   12435
            _ExtentY        =   13388
            _StockProps     =   1
            BackColor       =   16777215
            Page            =   2
            Begin VB.TextBox Text4 
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   2280
               MaxLength       =   4
               TabIndex        =   11
               Text            =   "0"
               ToolTipText     =   "기본값 4"
               Top             =   2325
               Width           =   1065
            End
            Begin VB.ComboBox cboBo 
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               ItemData        =   "frmINIT.frx":1A62
               Left            =   2280
               List            =   "frmINIT.frx":1A6C
               Style           =   2  '드롭다운 목록
               TabIndex        =   10
               Top             =   585
               Width           =   2535
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   2280
               MaxLength       =   4
               TabIndex        =   9
               Text            =   "0"
               ToolTipText     =   "기본값은 25"
               Top             =   1020
               Width           =   1065
            End
            Begin VB.TextBox Text3 
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   2280
               MaxLength       =   4
               TabIndex        =   8
               Text            =   "0"
               ToolTipText     =   "기본값 4"
               Top             =   1890
               Width           =   1065
            End
            Begin VB.TextBox Text2 
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   2280
               MaxLength       =   4
               TabIndex        =   7
               Text            =   "0"
               ToolTipText     =   "기본값 1"
               Top             =   1455
               Width           =   1065
            End
            Begin VB.ComboBox cboPrint 
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               ItemData        =   "frmINIT.frx":1A8A
               Left            =   2280
               List            =   "frmINIT.frx":1A97
               Style           =   2  '드롭다운 목록
               TabIndex        =   6
               Top             =   150
               Width           =   2535
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   11
               Left            =   150
               TabIndex        =   12
               Top             =   150
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               PictureFrames   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "frmINIT.frx":1AC6
               Caption         =   "         프린터 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":24D8
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   9
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   12
               Left            =   150
               TabIndex        =   13
               Top             =   585
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               PictureFrames   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "frmINIT.frx":281A
               Caption         =   "    보관증 형태 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":322C
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   9
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   13
               Left            =   150
               TabIndex        =   14
               Top             =   1020
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               PictureFrames   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "frmINIT.frx":356E
               Caption         =   "      위쪽 여백 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":3B08
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   9
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   14
               Left            =   150
               TabIndex        =   15
               Top             =   1455
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               PictureFrames   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "frmINIT.frx":3E4A
               Caption         =   "      왼쪽 여백 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":43E4
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   9
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   15
               Left            =   150
               TabIndex        =   16
               Top             =   1890
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               PictureFrames   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "frmINIT.frx":4726
               Caption         =   "        줄 간격 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":4CC0
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   9
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   16
               Left            =   150
               TabIndex        =   17
               Top             =   2325
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "출력장수 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":5002
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin XtremeSuiteControls.PushButton Command1 
               Height          =   930
               Index           =   0
               Left            =   5430
               TabIndex        =   18
               Top             =   135
               Width           =   1500
               _Version        =   851970
               _ExtentX        =   2646
               _ExtentY        =   1640
               _StockProps     =   79
               Caption         =   "좌표 출력"
               UseVisualStyle  =   -1  'True
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "(재출력에서는 무조건 1장 출력됨)"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3960
               TabIndex        =   19
               Top             =   2325
               Width           =   2985
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage2 
            Height          =   7590
            Left            =   -68890
            TabIndex        =   20
            Top             =   30
            Visible         =   0   'False
            Width           =   7050
            _Version        =   851970
            _ExtentX        =   12435
            _ExtentY        =   13388
            _StockProps     =   1
            BackColor       =   16777215
            Page            =   1
            Begin Threed.SSPanel SSPanel5 
               Height          =   30
               Index           =   0
               Left            =   165
               TabIndex        =   113
               Top             =   2805
               Width           =   6765
               _ExtentX        =   11933
               _ExtentY        =   53
               _Version        =   262144
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin VB.ComboBox cboMilAdd 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               ItemData        =   "frmINIT.frx":5344
               Left            =   4005
               List            =   "frmINIT.frx":534E
               Style           =   2  '드롭다운 목록
               TabIndex        =   26
               Top             =   2895
               Width           =   1545
            End
            Begin VB.ComboBox cboMil 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               ItemData        =   "frmINIT.frx":536C
               Left            =   2265
               List            =   "frmINIT.frx":5376
               Style           =   2  '드롭다운 목록
               TabIndex        =   25
               Top             =   2895
               Width           =   1710
            End
            Begin VB.ComboBox cboSale 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               ItemData        =   "frmINIT.frx":5386
               Left            =   2265
               List            =   "frmINIT.frx":5390
               Style           =   2  '드롭다운 목록
               TabIndex        =   24
               Top             =   3345
               Width           =   1710
            End
            Begin VB.ComboBox cboCoupon 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               ItemData        =   "frmINIT.frx":53A0
               Left            =   2265
               List            =   "frmINIT.frx":53AA
               Style           =   2  '드롭다운 목록
               TabIndex        =   23
               Top             =   4380
               Width           =   1710
            End
            Begin VB.ComboBox cboReturn 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               ItemData        =   "frmINIT.frx":53BA
               Left            =   2265
               List            =   "frmINIT.frx":53C4
               Style           =   2  '드롭다운 목록
               TabIndex        =   22
               Top             =   5850
               Width           =   1545
            End
            Begin XtremeSuiteControls.PushButton Command1 
               Height          =   930
               Index           =   1
               Left            =   5415
               TabIndex        =   21
               Top             =   150
               Width           =   1500
               _Version        =   851970
               _ExtentX        =   2646
               _ExtentY        =   1640
               _StockProps     =   79
               Caption         =   "설정 변경"
               UseVisualStyle  =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   2
               Left            =   150
               TabIndex        =   27
               Top             =   150
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "세탁 마진 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":53D4
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   0
               Left            =   150
               TabIndex        =   28
               Top             =   585
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "운동화 마진 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":5716
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   1
               Left            =   150
               TabIndex        =   29
               Top             =   1020
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "수선 마진 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":5A58
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   3
               Left            =   150
               TabIndex        =   30
               Top             =   1455
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "가죽/무스탕 마진 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":5D9A
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   4
               Left            =   150
               TabIndex        =   31
               Top             =   1890
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "카페트 마진 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":60DC
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   5
               Left            =   150
               TabIndex        =   32
               Top             =   2325
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "외주 운동화 마진 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":641E
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   6
               Left            =   150
               TabIndex        =   33
               Top             =   2895
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "마일리지 사용 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":6760
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   7
               Left            =   150
               TabIndex        =   34
               Top             =   3345
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "지정할인 사용 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":6AA2
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   8
               Left            =   150
               TabIndex        =   35
               Top             =   4380
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "특정할인 사용 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":6DE4
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   9
               Left            =   150
               TabIndex        =   36
               Top             =   5400
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "고가세탁 비율 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":7126
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   10
               Left            =   150
               TabIndex        =   37
               Top             =   5850
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "세탁비환불 사용 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":7468
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin CSTextLibCtl.sidbEdit txtRatio 
               Height          =   405
               Left            =   2265
               TabIndex        =   38
               Top             =   150
               Width           =   1020
               _Version        =   262145
               _ExtentX        =   1799
               _ExtentY        =   714
               _StockProps     =   125
               Text            =   " 0"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Insert          =   0   'False
               Modified        =   -1  'True
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   " 0"
               StartText.x     =   3
               StartText.y     =   4
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   18
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               Justification   =   2
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txtSports 
               Height          =   405
               Left            =   2265
               TabIndex        =   39
               Top             =   585
               Width           =   1020
               _Version        =   262145
               _ExtentX        =   1799
               _ExtentY        =   714
               _StockProps     =   125
               Text            =   " 0"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Insert          =   0   'False
               Modified        =   -1  'True
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   " 0"
               StartText.x     =   3
               StartText.y     =   4
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   18
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               Justification   =   2
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txtSRatio 
               Height          =   405
               Left            =   2265
               TabIndex        =   40
               Top             =   1020
               Width           =   1020
               _Version        =   262145
               _ExtentX        =   1799
               _ExtentY        =   714
               _StockProps     =   125
               Text            =   " 0"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Insert          =   0   'False
               Modified        =   -1  'True
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   " 0"
               StartText.x     =   3
               StartText.y     =   4
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   18
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               Justification   =   2
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txtGa 
               Height          =   405
               Left            =   2265
               TabIndex        =   41
               Top             =   1455
               Width           =   1020
               _Version        =   262145
               _ExtentX        =   1799
               _ExtentY        =   714
               _StockProps     =   125
               Text            =   " 0"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Insert          =   0   'False
               Modified        =   -1  'True
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   " 0"
               StartText.x     =   3
               StartText.y     =   4
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   18
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               Justification   =   2
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txtCar 
               Height          =   405
               Left            =   2265
               TabIndex        =   42
               Top             =   1890
               Width           =   1020
               _Version        =   262145
               _ExtentX        =   1799
               _ExtentY        =   714
               _StockProps     =   125
               Text            =   " 0"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Insert          =   0   'False
               Modified        =   -1  'True
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   " 0"
               StartText.x     =   3
               StartText.y     =   4
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   18
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               Justification   =   2
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txtOut 
               Height          =   405
               Left            =   2265
               TabIndex        =   43
               Top             =   2325
               Width           =   1020
               _Version        =   262145
               _ExtentX        =   1799
               _ExtentY        =   714
               _StockProps     =   125
               Text            =   " 0"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Insert          =   0   'False
               Modified        =   -1  'True
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   " 0"
               StartText.x     =   3
               StartText.y     =   4
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   18
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               Justification   =   2
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               Undo            =   0
               Data            =   0
            End
            Begin MSComCtl2.DTPicker dtpSaleStart 
               Height          =   420
               Left            =   2265
               TabIndex        =   44
               Top             =   3795
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   741
               _Version        =   393216
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   54788097
               CurrentDate     =   40066
            End
            Begin MSComCtl2.DTPicker dtpSaleEnd 
               Height          =   420
               Left            =   4005
               TabIndex        =   45
               Top             =   3795
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   741
               _Version        =   393216
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   54788097
               CurrentDate     =   40066
            End
            Begin MSComCtl2.DTPicker dtpCouponStart 
               Height          =   420
               Left            =   2265
               TabIndex        =   46
               Top             =   4830
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   741
               _Version        =   393216
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   54788097
               CurrentDate     =   40066
            End
            Begin MSComCtl2.DTPicker dtpCouponEnd 
               Height          =   420
               Left            =   4005
               TabIndex        =   47
               Top             =   4830
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   741
               _Version        =   393216
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   54788097
               CurrentDate     =   40066
            End
            Begin CSTextLibCtl.sidbEdit txtSale 
               Height          =   405
               Left            =   4005
               TabIndex        =   110
               Top             =   3345
               Width           =   1260
               _Version        =   262145
               _ExtentX        =   2222
               _ExtentY        =   714
               _StockProps     =   125
               Text            =   " 0"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Insert          =   0   'False
               Modified        =   -1  'True
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   " 0"
               StartText.x     =   3
               StartText.y     =   4
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   18
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               Justification   =   2
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txtCoupon 
               Height          =   405
               Left            =   4005
               TabIndex        =   111
               Top             =   4380
               Width           =   1260
               _Version        =   262145
               _ExtentX        =   2222
               _ExtentY        =   714
               _StockProps     =   125
               Text            =   " 0"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Insert          =   0   'False
               Modified        =   -1  'True
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   " 0"
               StartText.x     =   3
               StartText.y     =   4
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   18
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               Justification   =   2
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txtLuxury 
               Height          =   420
               Left            =   2265
               TabIndex        =   112
               Top             =   5400
               Width           =   1215
               _Version        =   262145
               _ExtentX        =   2143
               _ExtentY        =   741
               _StockProps     =   125
               Text            =   " 0"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.26
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Insert          =   0   'False
               Modified        =   -1  'True
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   " 0"
               StartText.x     =   3
               StartText.y     =   5
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   18
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               Justification   =   2
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               Undo            =   0
               Data            =   0
            End
            Begin Threed.SSPanel SSPanel5 
               Height          =   30
               Index           =   1
               Left            =   165
               TabIndex        =   114
               Top             =   4290
               Width           =   6765
               _ExtentX        =   11933
               _ExtentY        =   53
               _Version        =   262144
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel SSPanel5 
               Height          =   30
               Index           =   2
               Left            =   165
               TabIndex        =   115
               Top             =   5310
               Width           =   6765
               _ExtentX        =   11933
               _ExtentY        =   53
               _Version        =   262144
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "~"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   9
               Left            =   3810
               TabIndex        =   58
               Top             =   3765
               Width           =   165
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "~"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   10
               Left            =   3810
               TabIndex        =   57
               Top             =   4860
               Width           =   165
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   8
               Left            =   3345
               TabIndex        =   56
               Top             =   2370
               Width           =   195
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   3345
               TabIndex        =   55
               Top             =   195
               Width           =   195
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   1
               Left            =   3345
               TabIndex        =   54
               Top             =   1065
               Width           =   195
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   2
               Left            =   3345
               TabIndex        =   53
               Top             =   630
               Width           =   195
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   3
               Left            =   3345
               TabIndex        =   52
               Top             =   1500
               Width           =   195
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   4
               Left            =   3345
               TabIndex        =   51
               Top             =   1935
               Width           =   195
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   7
               Left            =   5325
               TabIndex        =   50
               Top             =   4425
               Width           =   195
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   6
               Left            =   3570
               TabIndex        =   49
               Top             =   5445
               Width           =   195
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   5
               Left            =   5325
               TabIndex        =   48
               Top             =   3390
               Width           =   195
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage1 
            Height          =   7590
            Left            =   1110
            TabIndex        =   59
            Top             =   30
            Width           =   7050
            _Version        =   851970
            _ExtentX        =   12435
            _ExtentY        =   13388
            _StockProps     =   1
            BackColor       =   16777215
            Page            =   0
            Begin VB.TextBox txtTelSMS 
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   2265
               TabIndex        =   109
               Top             =   5190
               Width           =   2505
            End
            Begin VB.TextBox txtTelStore 
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   2265
               TabIndex        =   108
               Top             =   4755
               Width           =   2505
            End
            Begin Threed.SSPanel SSPanel4 
               Height          =   420
               Left            =   2265
               TabIndex        =   105
               Top             =   4245
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   741
               _Version        =   262144
               BackColor       =   16777215
               BorderWidth     =   0
               BevelOuter      =   1
               BevelInner      =   2
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin VB.OptionButton optJa 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "대리점"
                  BeginProperty Font 
                     Name            =   "맑은 고딕"
                     Size            =   11.25
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   1
                  Left            =   1395
                  TabIndex        =   107
                  Top             =   45
                  Width           =   1005
               End
               Begin VB.OptionButton optJa 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "본사"
                  BeginProperty Font 
                     Name            =   "맑은 고딕"
                     Size            =   11.25
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   0
                  Left            =   105
                  TabIndex        =   106
                  Top             =   45
                  Value           =   -1  'True
                  Width           =   855
               End
            End
            Begin Threed.SSPanel SSPanel1 
               Height          =   420
               Left            =   2265
               TabIndex        =   102
               Top             =   3810
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   741
               _Version        =   262144
               BackColor       =   16777215
               BorderWidth     =   0
               BevelOuter      =   1
               BevelInner      =   2
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin VB.OptionButton OptSu 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "본사"
                  BeginProperty Font 
                     Name            =   "맑은 고딕"
                     Size            =   11.25
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   0
                  Left            =   105
                  TabIndex        =   104
                  Top             =   45
                  Value           =   -1  'True
                  Width           =   855
               End
               Begin VB.OptionButton OptSu 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "대리점"
                  BeginProperty Font 
                     Name            =   "맑은 고딕"
                     Size            =   11.25
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   1
                  Left            =   1395
                  TabIndex        =   103
                  Top             =   45
                  Width           =   1035
               End
            End
            Begin VB.TextBox txtColor 
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   2265
               TabIndex        =   101
               Top             =   2850
               Width           =   2505
            End
            Begin VB.TextBox txtName 
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   2265
               TabIndex        =   100
               Top             =   2415
               Width           =   2505
            End
            Begin VB.TextBox txtNo 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   2265
               TabIndex        =   99
               Top             =   1980
               Width           =   2505
            End
            Begin VB.TextBox txtMstCode 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   2265
               TabIndex        =   98
               Top             =   1545
               Width           =   2505
            End
            Begin VB.TextBox txtStoreName 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   2265
               TabIndex        =   97
               Top             =   570
               Width           =   2505
            End
            Begin VB.TextBox txtStoreCode 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   2265
               TabIndex        =   96
               Top             =   135
               Width           =   2505
            End
            Begin VB.CheckBox chkSMSEMART 
               Caption         =   "이마트 SMS"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4935
               TabIndex        =   66
               Top             =   2625
               Width           =   1575
            End
            Begin VB.CheckBox chkTelPrt 
               Caption         =   "고객 전화번호 모두 출력"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   585
               Left            =   4935
               TabIndex        =   65
               Top             =   1830
               Value           =   1  '확인
               Width           =   1785
            End
            Begin VB.ComboBox Cmb_DaySale 
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               ItemData        =   "frmINIT.frx":77AA
               Left            =   2265
               List            =   "frmINIT.frx":77C6
               TabIndex        =   64
               Top             =   3285
               Width           =   2505
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00FFFFFF&
               Caption         =   "이전 코드 내용"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1785
               Left            =   135
               TabIndex        =   60
               Top             =   5685
               Width           =   6810
               Begin VB.TextBox txtOldCode 
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "맑은 고딕"
                     Size            =   11.25
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   1
                  Left            =   2250
                  MaxLength       =   3
                  TabIndex        =   62
                  Top             =   795
                  Width           =   1125
               End
               Begin VB.TextBox txtOldCode 
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "맑은 고딕"
                     Size            =   11.25
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   0
                  Left            =   2250
                  MaxLength       =   4
                  TabIndex        =   61
                  Top             =   345
                  Width           =   1125
               End
               Begin MSComCtl2.DTPicker dtpOldDate 
                  Height          =   420
                  Left            =   2250
                  TabIndex        =   63
                  Top             =   1245
                  Width           =   1740
                  _ExtentX        =   3069
                  _ExtentY        =   741
                  _Version        =   393216
                  Enabled         =   0   'False
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "맑은 고딕"
                     Size            =   11.25
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Format          =   54788097
                  CurrentDate     =   39553
               End
               Begin Threed.SSPanel Label5 
                  Height          =   405
                  Index           =   37
                  Left            =   120
                  TabIndex        =   93
                  Top             =   345
                  Width           =   2100
                  _ExtentX        =   3704
                  _ExtentY        =   714
                  _Version        =   262144
                  CaptionStyle    =   1
                  ForeColor       =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "지점코드 "
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frmINIT.frx":7815
                  BorderWidth     =   0
                  Alignment       =   4
                  PictureAlignment=   2
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel Label5 
                  Height          =   405
                  Index           =   38
                  Left            =   120
                  TabIndex        =   94
                  Top             =   1245
                  Width           =   2100
                  _ExtentX        =   3704
                  _ExtentY        =   714
                  _Version        =   262144
                  CaptionStyle    =   1
                  ForeColor       =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "적용일자 "
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frmINIT.frx":7B57
                  BorderWidth     =   0
                  Alignment       =   4
                  PictureAlignment=   2
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel Label5 
                  Height          =   405
                  Index           =   39
                  Left            =   120
                  TabIndex        =   95
                  Top             =   795
                  Width           =   2100
                  _ExtentX        =   3704
                  _ExtentY        =   714
                  _Version        =   262144
                  CaptionStyle    =   1
                  ForeColor       =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "대리점 번호 "
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frmINIT.frx":7E99
                  BorderWidth     =   0
                  Alignment       =   4
                  PictureAlignment=   2
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   25
               Left            =   135
               TabIndex        =   67
               Top             =   135
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "가맹점 코드 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":81DB
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   26
               Left            =   135
               TabIndex        =   68
               Top             =   570
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "가맹점 명칭 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":851D
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   27
               Left            =   135
               TabIndex        =   69
               Top             =   1005
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "적용일자 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":885F
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlClear 
               Height          =   405
               Left            =   135
               TabIndex        =   70
               Top             =   1545
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "지점코드 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":8BA1
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   29
               Left            =   135
               TabIndex        =   71
               Top             =   1980
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "대리점 번호 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":8EE3
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   30
               Left            =   135
               TabIndex        =   72
               Top             =   2415
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "대리점 이름 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":9225
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   31
               Left            =   135
               TabIndex        =   73
               Top             =   2850
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "대리점 색상 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":9567
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   32
               Left            =   135
               TabIndex        =   74
               Top             =   3285
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "목요 세일 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":98A9
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   33
               Left            =   135
               TabIndex        =   75
               Top             =   3810
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "수선 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":9BEB
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   34
               Left            =   135
               TabIndex        =   76
               Top             =   4245
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "짜집기 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":9F2D
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   35
               Left            =   135
               TabIndex        =   77
               Top             =   4755
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "매장 전화번호 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":A26F
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel Label5 
               Height          =   405
               Index           =   36
               Left            =   135
               TabIndex        =   78
               Top             =   5190
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   714
               _Version        =   262144
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "문자발신전화 "
               PictureBackgroundStyle=   2
               PictureBackground=   "frmINIT.frx":A5B1
               BorderWidth     =   0
               Alignment       =   4
               PictureAlignment=   2
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin XtremeSuiteControls.PushButton cmdChange 
               Height          =   930
               Left            =   5415
               TabIndex        =   79
               Top             =   150
               Width           =   1500
               _Version        =   851970
               _ExtentX        =   2646
               _ExtentY        =   1640
               _StockProps     =   79
               Caption         =   "정보 변경"
               UseVisualStyle  =   -1  'True
            End
            Begin MSComCtl2.DTPicker dtpStart 
               Height          =   420
               Left            =   2265
               TabIndex        =   80
               Top             =   1005
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   741
               _Version        =   393216
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "맑은 고딕"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   54788097
               CurrentDate     =   39553
            End
         End
      End
   End
End
Attribute VB_Name = "frmINIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bchk As Boolean
Dim S_Gu As String
Dim J_Gu As String

Private Function BlankChk() As Boolean
    BlankChk = False
    
    If Trim(txtNo.Text) = "" Then
        txtNo.SetFocus
    ElseIf Trim(txtColor.Text) = "" Then
        txtColor.SetFocus
    ElseIf Trim(txtName.Text) = "" Then
        txtName.SetFocus
    ElseIf Trim(txtRatio.Text) = "" Then
        txtNo.SetFocus
    ElseIf Trim(txtTel1.Text) = "" Then
        txtTel1.SetFocus
    ElseIf Trim(txtTel2.Text) = "" Then
        txtTel2.SetFocus
    
    ElseIf Trim(txtStoreCode.Text) = "" Then
        If txtStoreCode.Enabled = True Then txtStoreCode.SetFocus
    ElseIf Trim(txtStoreName.Text) = "" Then
        If txtStoreName.Enabled = True Then txtStoreName.SetFocus
    Else
        BlankChk = True
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'+------------------------------------------------------
'+ 2003/02/11 수정
'+
'+루틴설명      - 비밀번호확인
'+  1. 암호를 확인하여 암호 규칙에 맞으면 화면을 종료한다.
'+  2. 레지스터리에 저장한다.
'+
'+------------------------------------------------------
Private Sub cmdChange_Click()
    Dim strPass As String
    
    ' 입력 확인
    
    strPass = InputBox("암호를 입력하여 주십시요", "변경 암호")
    
    If Len(strPass) <= 0 Then
        Exit Sub
    End If
    
'   기본 디폴드 암호.. ( 프로그램 셋팅/설치를 위한 암호 )
    If UCase(strPass) = "DUDTJSGH" Then
        chkPassWord = True
        txtMstCode.Enabled = True
        txtNo.Enabled = True
        txtStoreCode.Enabled = True
        txtStoreName.Enabled = True
        dtpStart.Enabled = True
        
        txtOldCode(0).Enabled = True
        txtOldCode(1).Enabled = True
        dtpOldDate.Enabled = True
        chkTelPrt.Enabled = True
        chkSMSEMART.Enabled = True
        
        Exit Sub
    End If
    
    ' 비밀번호 확인
    strPass = IsCodePassWord(strPass)
    
    If strPass = "-1" Or strPass = "-3" Then
        If strPass = "-3" Then MsgBox "입력한 내용이 정확하지 않습니다.", vbCritical, "입력오류"
        Exit Sub
    Else
        txtMstCode.Enabled = True
        txtNo.Enabled = True
        txtStoreCode.Enabled = True
        txtStoreName.Enabled = True
        dtpStart.Enabled = True
    
        txtOldCode(0).Enabled = True
        txtOldCode(1).Enabled = True
        dtpOldDate.Enabled = True
        
        chkTelPrt.Enabled = True
        chkSMSEMART.Enabled = True
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim strPass As String
    
    Select Case Index
        Case 0: Call PrintPointDisplay
        Case 1
        
            strPass = InputBox("암호를 입력하여 주십시요", "변경 암호")
            
            If Len(strPass) <= 0 Then
                Exit Sub
            End If
            
            '기본 디폴드 암호.. ( 프로그램 셋팅/설치를 위한 암호 )
            If UCase(strPass) = "DUDTJSGH" Then
                Call ButtonEnabled(True)
                Exit Sub
            End If
            ' 비밀번호 확인
            strPass = IsSportsPassWord(strPass)
            If strPass = "-1" Or strPass = "-3" Then
                If strPass = "-3" Then MsgBox "입력한 내용이 정확하지 않습니다.", vbCritical, "입력오류"
                Exit Sub
            Else
                Call ButtonEnabled(True)
            End If

        Case Else
    
    End Select
End Sub


Private Sub ButtonEnabled(bMode As Boolean)
    txtRatio.Enabled = bMode
    txtSports.Enabled = bMode
    txtSRatio.Enabled = bMode
    txtGa.Enabled = bMode
    txtCar.Enabled = bMode
    cmdSave.Enabled = bMode
    cboMil.Enabled = bMode
    cboMilAdd.Enabled = bMode
    cboSale.Enabled = bMode
    txtSale.Enabled = bMode
    dtpSaleStart.Enabled = bMode
    dtpSaleEnd.Enabled = bMode
    cboCoupon.Enabled = bMode
    txtCoupon.Enabled = bMode
    dtpCouponStart.Enabled = bMode
    dtpCouponEnd.Enabled = bMode
    txtLuxury.Enabled = bMode
    txtOut.Enabled = bMode
    cboReturn.Enabled = bMode
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
    Dim nday    As Double
    Dim intMM   As Integer
    Dim dPass   As Double
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

Private Sub Form_Load()
    Dim strCode As String
    
    MoveWindow Me
    
    Query = "SELECT * FROM TB_대리점정보 "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If SUBRs.EOF Or SUBRs.BOF Then
        SUBRs.Close
        Set SUBRs = Nothing
        
        Exit Sub
    Else
        txtNo.Text = SUBRs!대리점번호 & ""
        txtColor.Text = SUBRs!대리점색상 & ""
        txtName.Text = SUBRs!대리점명 & ""
        J_Gu = Trim(SUBRs!할인종료일) & ""
        
        If J_Gu = "1" Then
            optJa(0).Value = True
            optJa(1).Value = False
        ElseIf J_Gu = "2" Then
            optJa(0).Value = False
            optJa(1).Value = True
        Else
            J_Gu = "1"
            optJa(0).Value = True
            optJa(1).Value = False
        End If
        
        S_Gu = Trim(SUBRs!수선) & ""
        
        If S_Gu = "1" Then
            OptSu(0).Value = True
            OptSu(1).Value = False
            J_Gu = "1"
            optJa(0).Value = True
            optJa(1).Value = False
            optJa(0).Enabled = False
            optJa(1).Enabled = False
        ElseIf S_Gu = "2" Then
            OptSu(0).Value = False
            OptSu(1).Value = True
            optJa(0).Enabled = True
            optJa(1).Enabled = True
        End If
        
        txtRatio.Text = IIf(IsNull(SUBRs!비율), "30", SUBRs!비율) & ""
        txtSRatio.Text = IIf(IsNull(SUBRs!수선마진), "30", SUBRs!수선마진) & ""
        txtSports.Text = IIf(IsNull(SUBRs!운동화마진), "40", SUBRs!운동화마진) & ""
        txtGa.Text = IIf(IsNull(SUBRs!가죽무스탕마진), "40", SUBRs!가죽무스탕마진) & ""
        txtCar.Text = IIf(IsNull(SUBRs!카페트마진), "40", SUBRs!카페트마진) & ""
        txtOut.Text = IIf(IsNull(SUBRs!외주운동화마진), "0", SUBRs!외주운동화마진) & ""
        
        If IsNull(SUBRs!마일리지여부) Then
            cboMil.ListIndex = 1
        Else
            cboMil.ListIndex = IIf(SUBRs!마일리지여부 = "Y", 0, 1)
        End If
        
        If IsNull(SUBRs!마일리지증가구분) Then
            cboMilAdd.ListIndex = 0
        Else
            cboMilAdd.ListIndex = IIf(SUBRs!마일리지증가구분 <> "1", 0, 1)
        End If
        
        
        If IsNull(SUBRs!지정할인여부) Then
            cboSale.ListIndex = 1
        Else
            cboSale.ListIndex = IIf(SUBRs!지정할인여부 = "Y", 0, 1)
        End If
        
        txtSale.Text = IIf(IsNull(SUBRs!지정할인비율), "20", SUBRs!지정할인비율)
        dtpSaleStart.Value = IIf(IsNull(SUBRs!지정할인시작일), "2009-01-01", Format(SUBRs!지정할인시작일, "YYYY-MM-DD"))
        dtpSaleEnd.Value = IIf(IsNull(SUBRs!지정할인종료일), "2009-01-01", Format(SUBRs!지정할인종료일, "YYYY-MM-DD"))
                
        If IsNull(SUBRs!특정할인여부) Then
            cboCoupon.ListIndex = 1
        Else
            cboCoupon.ListIndex = IIf(SUBRs!특정할인여부 = "Y", 0, 1)
        End If
        
        txtCoupon.Text = IIf(IsNull(SUBRs!특정할인비율), "30", SUBRs!특정할인비율)
        dtpCouponStart.Value = IIf(IsNull(SUBRs!특정할인시작일), "2009-01-01", Format(SUBRs!특정할인시작일, "YYYY-MM-DD"))
        dtpCouponEnd.Value = IIf(IsNull(SUBRs!특정할인종료일), "2009-01-01", Format(SUBRs!특정할인종료일, "YYYY-MM-DD"))
        
        txtLuxury.Text = IIf(IsNull(SUBRs!고가세탁비율), "300", SUBRs!고가세탁비율)
        
        If IsNull(SUBRs!세탁비환불여부) Then
            cboReturn.ListIndex = 1
        Else
            cboReturn.ListIndex = IIf(SUBRs!세탁비환불여부 = "Y", 0, 1)
        End If
        
        txtStoreCode.Text = IIf(IsNull(SUBRs!StoreCode), " ", SUBRs!StoreCode)
        txtStoreName.Text = IIf(IsNull(SUBRs!StoreName), " ", SUBRs!StoreName)
        dtpStart.Value = IIf(IsDate(Format(SUBRs!StartDate, "YYYY-MM-DD")), Format(SUBRs!StartDate, "YYYY-MM-DD"), "1990-01-01")
        
        
        Select Case SUBRs!목요세일
            Case "1": Cmb_DaySale.Text = "일요일"
            Case "2": Cmb_DaySale.Text = "월요일"
            Case "3": Cmb_DaySale.Text = "화요일"
            Case "4": Cmb_DaySale.Text = "수요일"
            Case "5": Cmb_DaySale.Text = "목요일"
            Case "6": Cmb_DaySale.Text = "금요일"
            Case "7": Cmb_DaySale.Text = "토요일"
            Case Else: Cmb_DaySale.Text = "해당없음"
        End Select
        
        txtTel1.Text = SUBRs!전화1 & ""
        txtTel2.Text = SUBRs!전화2 & ""
        
        txtTelStore.Text = SUBRs!telStore & ""
        txtTelSMS.Text = SUBRs!telSMS & ""
        
        If IsNull(SUBRs!프린터) Then
            cboPrint.ListIndex = 0
        ElseIf SUBRs!프린터 >= "0" And cboPrint.ListCount > SUBRs!프린터 Then
            cboPrint.ListIndex = SUBRs!프린터
        Else
            cboPrint.ListIndex = 0
        End If
        
        If IsNull(SUBRs!보관증종류) Then
            cboBo.ListIndex = 0
        ElseIf SUBRs!보관증종류 >= "0" And cboPrint.ListCount > SUBRs!보관증종류 Then
            cboBo.ListIndex = SUBRs!보관증종류
        Else
            cboBo.ListIndex = 0
        End If
        
        If IsNull(SUBRs.Fields("ServerIP")) = True Then
            txtSMSIPAddress.Text = "store.clean-aid.co.kr,8657"
        Else
            txtSMSIPAddress.Text = Trim(SUBRs.Fields("ServerIP") & "")
        End If
        
        If IsNull(SUBRs.Fields("ServerDB")) = True Then
            txtSMSDBName.Text = "Laundry"
        Else
            txtSMSDBName.Text = Trim(SUBRs.Fields("ServerDB") & "")
        End If
        
        If IsNull(SUBRs.Fields("ServerUser")) = True Then
            txtSMSUserName.Text = "sa"
        Else
            txtSMSUserName.Text = Trim(SUBRs.Fields("ServerUser") & "")
        End If
        
        If IsNull(SUBRs.Fields("ServerPass")) = True Then
            txtSMSUserPass.Text = ""
        Else
            txtSMSUserPass.Text = Trim(SUBRs.Fields("ServerPass") & "")
        End If
        
        If IsNull(SUBRs.Fields("TimeOut")) = True Then
            m_CommandTimeOut = 30
        Else
            m_CommandTimeOut = Val(Trim(SUBRs.Fields("TimeOut") & ""))
        End If
        
        If IsNull(SUBRs.Fields("SMS_EMART")) = True Then
            chkSMSEMART.Value = 0
        Else
            chkSMSEMART.Value = IIf(SUBRs.Fields("SMS_EMART") & "" = "Y", 1, 0)
        End If
    End If
    SUBRs.Close
    Set SUBRs = Nothing
    
    Text1.Text = GetPrtStartPoint("TOP")
    Text2.Text = GetPrtStartPoint("LEFT")
    Text3.Text = GetPrtStartPoint("HEIGHT")
    
    
    Text4.Text = GetIniStr("Printer", "Count", "", iniFile)
    chkTelPrt.Value = GetIniStr("Printer", "TelPrint", "0", iniFile)
    
    '기본을 모뎀으로 한다.
    If GetSetting("Laundry_Zi", "Connect", "Type", "True") Then
        OptConnect(0).Value = True
    Else
        OptConnect(1).Value = True
    End If
    
    ' 지점 코드
    txtMstCode.Text = GetIniStr("Connect", "MstCode", "", iniFile)
    txtIPAddress.Text = GetIniStr("Connect", "RemoteIP", "", iniFile)
    txtMsgPort.Text = GetIniStr("Connect", "MsgRemotePort", "", iniFile)
    txtFilePort.Text = GetIniStr("Connect", "FileRemotePort", "", iniFile)
    
    If txtMsgPort.Text = "" Then txtMsgPort.Text = "8607"
    If txtFilePort.Text = "" Then txtFilePort.Text = "8602"
    
    txtOldCode(0).Text = GetIniStr("Store", "OldMstCode", txtMstCode.Text, iniFile)
    txtOldCode(1).Text = GetIniStr("Store", "OldCode", "", iniFile)
        
    
    dtpOldDate.Tag = GetIniStr("Store", "OldDate", "", iniFile)
    
    If IsDate(dtpOldDate.Tag) = False Then dtpOldDate.Tag = Date
    
    dtpOldDate.Value = dtpOldDate.Tag
    
    ' 변경 내용을 처리하기 위하여..
    txtMstCode.Tag = txtMstCode.Text
    txtStoreCode.Tag = txtStoreCode.Text
    txtStoreName.Tag = txtStoreName.Text
    dtpStart.Tag = dtpStart.Value
    txtNo.Tag = txtNo.Text
End Sub

Private Sub pnlClear_Click()
    If InputBox("행사 내용 삭제를 위하여 암호를 입력하여 주십시요", "변경 암호") = "2025" Then
       ' 이전 자료를 모두 지운다.
       ADOCon.Execute "DELETE FROM TB_할인정보 "
       
       MsgBox "행사 관련 내용 삭제 완료", vbInformation
    End If
End Sub

Private Sub txtCoupon_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57, vbKeyBack
        
        Case Else
            KeyAscii = 0
            Exit Sub
    End Select
End Sub

Private Sub txtCoupon_LostFocus()
    If IsNumeric(txtCoupon.Text) = False Then
        MsgBox "숫자만 입력 가능 합니다."
        txtCoupon.SelStart = 0: txtCoupon.SelLength = 3
        txtCoupon.SetFocus
        Exit Sub
    End If
    
    If Val(txtCoupon.Text) > 100 Then
        MsgBox "100 보다 큰수는 입력할 수 없습니다.", vbInformation, "확인"
        txtCoupon.Text = "0"
        txtCoupon.SelStart = 0: txtCoupon.SelLength = 3
        txtCoupon.SetFocus
        Exit Sub
    End If

End Sub

Private Sub txtMstCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57, vbKeyBack
        
            
        Case Else
            KeyAscii = 0
            Exit Sub
    End Select
End Sub

Private Sub txtNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57, vbKeyBack
        
        Case Else
            KeyAscii = 0
            Exit Sub
    End Select
End Sub

Private Sub txtSale_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57, vbKeyBack
        
        Case Else
            KeyAscii = 0
            Exit Sub
    End Select
End Sub

Private Sub txtSale_LostFocus()
    If IsNumeric(txtSale.Text) = False Then
        MsgBox "숫자만 입력 가능 합니다."
        txtSale.SelStart = 0: txtSale.SelLength = 3
        txtSale.SetFocus
        Exit Sub
    End If
    
    If Val(txtSale.Text) > 100 Then
        MsgBox "100 보다 큰수는 입력할 수 없습니다.", vbInformation, "확인"
        txtSale.Text = "0"
        txtSale.SelStart = 0: txtSale.SelLength = 3
        txtSale.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtStoreCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57, vbKeyBack
            dtpStart.Value = Date
        
            
        Case Else
            KeyAscii = 0
            Exit Sub
    End Select

End Sub

Private Sub OptConnect_Click(Index As Integer)
    Dim strValue As String
    
    If Index = 1 Then
        ' 인터넷을 선택했을 경우 기존 설정 사항이 없을 경우 만든다.
    
        strValue = GetIniStr("Connect", "RemoteIP", "", iniFile)
        
        If strValue = "" Then
            ' RemoteIP=61.77.137.104    ' 본사 서버의 IP
            ' FileRemotePort = 8627     ' 상대에게 파일을 전송해줄 포트 ( 본사 파일 포트 )
            ' FileLocalPort = 8629      ' 본사로 부터 전송 받기         ( 클라이언트 파일 포트 )
            ' MsgRemotePort = 8607      ' 서버가 메시지를 기다리는 포트 ( 본사 메시지 포트 )
            ' MsgLocalPort =            ' 메시지를 주고 받을 포트       ( 클라이언트 메시지 포트 - 자동 할당)
        
            strValue = "web.clean-aid.co.kr"
            Call SetIniStr("Connect", "RemoteIP", strValue, iniFile)
            
            strValue = "8627"
            Call SetIniStr("Connect", "FileRemotePort", strValue, iniFile)
            
            strValue = "8629"
            Call SetIniStr("Connect", "FileLocalPort", strValue, iniFile)
            
            strValue = "8607"
            Call SetIniStr("Connect", "MsgRemotePort", strValue, iniFile)
        End If
    End If
End Sub

Private Sub optJa_Click(Index As Integer)
    If Index = 0 Then
        J_Gu = "1"
    Else
        J_Gu = "2"
    End If
End Sub

Private Sub OptPrinter_Click(Index As Integer)
    If Index Then
        Text1.Enabled = True
        Text2.Enabled = True
        Text3.Enabled = True
    Else
        Text1.Enabled = False
        Text2.Enabled = False
        Text3.Enabled = False
    End If
End Sub

Private Sub OptSu_Click(Index As Integer)
    If Index = 0 Then
        S_Gu = "1"
        optJa(0).Value = True
        optJa(1).Value = False
        optJa(0).Enabled = False
        optJa(1).Enabled = False
    Else
        S_Gu = "2"
        optJa(0).Enabled = True
        optJa(1).Enabled = True
    End If
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrRtn
    
    Dim strAgentCode As String
    Dim strDaySale   As String
    Dim Msg          As String
    
    If BlankChk = False Then Exit Sub
    
    txtStoreCode.Text = Trim(txtStoreCode.Text)
    
    If Len(txtStoreCode.Text) <> 6 Then
        MsgBox "가맹점코드 입력에러", vbInformation, "확인"
        
        Exit Sub
    End If
    
    strAgentCode = Trim(txtNo.Text)
    
    If Len(strAgentCode) <> 3 Then
        MsgBox "대리점코드 입력에러", vbInformation, "확인"
        Exit Sub
    End If
    
    Msg = "[0 ~ 100] 사이의 숫자만입력이 가능합니다."
    
    If txtRatio.Value < 0 Or txtRatio.Value > 100 Then
        txtRatio.SetFocus
        
        MsgBox Msg, vbInformation, "확인"
        Exit Sub
    End If
    
    If txtSRatio.Value < 0 Or txtSRatio.Value > 100 Then
        txtSRatio.SetFocus
        MsgBox Msg, vbInformation, "확인"
        Exit Sub
    End If
    
    If txtSports.Value < 0 Or txtSports.Value > 100 Then
        txtSports.SetFocus
        MsgBox Msg, vbInformation, "확인"
        Exit Sub
    End If
    
    If txtGa.Value < 0 Or txtGa.Value > 100 Then
        txtGa.SetFocus
        MsgBox Msg, vbInformation, "확인"
        Exit Sub
    End If
    
    If txtCar.Value < 0 Or txtCar.Value > 100 Then
        txtCar.SetFocus
        MsgBox Msg, vbInformation, "확인"
        Exit Sub
    End If
    
    If txtOut.Value < 0 Or txtOut.Value > 100 Then
        txtOut.SetFocus
        MsgBox Msg, vbInformation, "확인"
        Exit Sub
    End If
    
    If Format(dtpSaleStart.Value, "YYYY-MM-DD") > Format(dtpSaleEnd.Value, "YYYY-MM-DD") Then
        MsgBox "특정할인 일자를 확인하여 주십시요.", vbInformation, "확인"
        Exit Sub
    End If
    
    If Format(dtpCouponStart.Value, "YYYY-MM-DD") > Format(dtpCouponEnd.Value, "YYYY-MM-DD") Then
        MsgBox "특정할인 일자를 확인하여 주십시요.", vbInformation, "확인"
        Exit Sub
    End If
    
    txtSMSIPAddress.Text = Trim(txtSMSIPAddress.Text)
    txtSMSDBName.Text = Trim(txtSMSDBName.Text)
    txtSMSUserName.Text = Trim(txtSMSUserName.Text)
    txtSMSUserPass.Text = Trim(txtSMSUserPass.Text)
    
    Select Case Cmb_DaySale.Text
        Case "일요일": strDaySale = "1"
        Case "월요일": strDaySale = "2"
        Case "화요일": strDaySale = "3"
        Case "수요일": strDaySale = "4"
        Case "목요일": strDaySale = "5"
        Case "금요일": strDaySale = "6"
        Case "토요일": strDaySale = "7"
        Case Else
            strDaySale = "0"
    End Select
        
    Printer_Gb = cboPrint.ItemData(cboPrint.ListIndex)
    Printer_BO_Gb = cboBo.ItemData(cboBo.ListIndex)
    
    '----------------------------------------------------------------------------------------------
    '
    '----------------------------------------------------------------------------------------------
    Query = "UPDATE TB_대리점정보 "
    Query = Query & "SET 대리점번호 = '" & strAgentCode & "', "
    Query = Query & "    대리점색상 = '" & txtColor.Text & "', "
    Query = Query & "    대리점명   = '" & txtName.Text & "', "
    Query = Query & "    StoreCode  = '" & txtStoreCode.Text & "', "
    Query = Query & "    StoreName  = '" & txtStoreName.Text & "', "
    Query = Query & "    StartDate  = '" & Format(dtpStart.Value, "YYYY-MM-DD") & "', "
    Query = Query & "    수선       = '" & S_Gu & "', "
    Query = Query & "    할인종료일 = '" & J_Gu & "', "
    Query = Query & "    전화1      = '" & txtTel1.Text & "', "
    Query = Query & "    전화2      = '" & txtTel2.Text & "', "
    Query = Query & "    TelStore      = '" & txtTelStore.Text & "', "
    Query = Query & "    TelSMS      = '" & txtTelSMS.Text & "', "
    Query = Query & "    목요세일   = '" & strDaySale & "', "
    
    Query = Query & "    비율       = '" & txtRatio.Text & "', "
    Query = Query & "    수선마진   = '" & txtSRatio.Text & "', "
    Query = Query & "    운동화마진     = '" & txtSports.Text & "', "
    Query = Query & "    가죽무스탕마진 = '" & txtGa.Text & "', "
    Query = Query & "    카페트마진     = '" & txtCar.Text & "', "
    Query = Query & "    외주운동화마진 = '" & txtOut.Text & "', "
    
    Query = Query & "    마일리지여부   = '" & IIf(Trim(cboMil.Text) = "예", "Y", "N") & "', "
    Query = Query & "    마일리지증가구분   = '" & IIf(cboMilAdd.ListIndex = 0, "0", "1") & "', "
    
    
    Query = Query & "    지정할인여부   = '" & IIf(Trim(cboSale.Text) = "예", "Y", "N") & "', "
    Query = Query & "    지정할인비율     = '" & txtSale.Text & "',  "
    Query = Query & "    지정할인시작일     = '" & Format(dtpSaleStart.Value, "YYYY-MM-DD") & "',  "
    Query = Query & "    지정할인종료일     = '" & Format(dtpSaleEnd.Value, "YYYY-MM-DD") & "',  "
    
    Query = Query & "    특정할인여부   = '" & IIf(Trim(cboCoupon.Text) = "예", "Y", "N") & "', "
    Query = Query & "    특정할인비율     = '" & txtCoupon.Text & "',  "
    Query = Query & "    특정할인시작일     = '" & Format(dtpCouponStart.Value, "YYYY-MM-DD") & "',  "
    Query = Query & "    특정할인종료일     = '" & Format(dtpCouponEnd.Value, "YYYY-MM-DD") & "',  "
    
'    Query = Query & "    쿠폰할인여부   = '" & IIf(Trim(cboCoupon.Text) = "예", "Y", "N") & "', "
'    Query = Query & "    쿠폰할인비율     = '" & txtCoupon.Text & "',  "
'    Query = Query & "    쿠폰할인시작일     = '" & Format(dtpCouponStart.Value, "YYYY-MM-DD") & "',  "
'    Query = Query & "    쿠폰할인종료일     = '" & Format(dtpCouponEnd.Value, "YYYY-MM-DD") & "',  "
    
    Query = Query & "    고가세탁비율     = '" & txtLuxury.Text & "',  "
    Query = Query & "    세탁비환불여부   = '" & IIf(Trim(cboReturn.Text) = "예", "Y", "N") & "', "
    
    Query = Query & "    ServerIP = ' " & txtSMSIPAddress.Text & "', "
    Query = Query & "    ServerDB = ' " & txtSMSDBName.Text & "', "
    Query = Query & "    ServerUser = ' " & txtSMSUserName.Text & "', "
    Query = Query & "    ServerPass = ' " & txtSMSUserPass.Text & "', "
    Query = Query & "    보관증종류     = '" & Printer_BO_Gb & "', "
    Query = Query & "    SMS_EMART     = '" & IIf(chkSMSEMART.Value = 1, "Y", "N") & "', "
    Query = Query & "    프린터     = '" & Printer_Gb & "'"
    ADOCon.Execute Query
    
    SaveSetting "Laundry_Zi", "Printer", "Top", Text1.Text
    SaveSetting "Laundry_Zi", "Printer", "Left", Text2.Text
    SaveSetting "Laundry_Zi", "Printer", "Height", Text3.Text
    
    SaveSetting "Laundry_Zi", "Connect", "Type", IIf(OptConnect(0).Value, "True", "False")
                
    Text4.Text = Trim(Text4.Text)
    Text4.Text = IIf(IsNumeric(Text4.Text) = True, Text4.Text, "1")
    
    Call SetIniStr("Printer", "Count", Text4.Text, iniFile)
    Call SetIniStr("Printer", "TelPrint", CStr(chkTelPrt.Value), iniFile)
    
    Call SetIniStr("Connect", "MstCode", txtMstCode.Text, iniFile)
    Call SetIniStr("Connect", "RemoteIP", txtIPAddress.Text, iniFile)
    Call SetIniStr("Connect", "MsgRemotePort", txtMsgPort.Text, iniFile)
    Call SetIniStr("Connect", "FileRemotePort", txtFilePort.Text, iniFile)
    
    Call SetIniStr("Store", "OldMstCode", txtOldCode(0).Text, iniFile)
    Call SetIniStr("Store", "OldCode", txtOldCode(1).Text, iniFile)
    
    Call SetIniStr("Store", "OldDate", dtpOldDate.Value, iniFile)
    
    '이전 내용의 자료가 변경되었을 경우 전송하도록 처리한다.
    If txtNo.Tag <> txtNo.Text Or txtMstCode.Tag <> txtMstCode.Text Or txtStoreCode.Tag <> txtStoreCode.Text Or txtStoreName.Tag <> txtStoreName.Text Or dtpStart.Tag <> dtpStart.Value Then
        ' 정보를 저장한다.
        Call SendStoreDefaultInfo(dtpStart.Tag, txtMstCode.Tag, txtNo.Tag, txtStoreCode.Tag, txtStoreName.Tag)
    End If

    MsgBox "프로그램을 다시 시작하십시요     ", vbCritical, "확인"
    
    End
    
    Exit Sub
    
ErrRtn:
    Resume Next
End Sub

'Private Sub TabStrip1_Click()
'    If TabStrip1.SelectedItem.Index = 1 Then
'        frmDef.Visible = True
'        frmDef.ZOrder 0
'
'    ElseIf TabStrip1.SelectedItem.Index = 2 Then
'        frmMaJin.Visible = True
'        frmMaJin.ZOrder 0
'
'    ElseIf TabStrip1.SelectedItem.Index = 3 Then
'        frmPrint.Visible = True
'        frmPrint.ZOrder 0
'    End If
'End Sub
 
