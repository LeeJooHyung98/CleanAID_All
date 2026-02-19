VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm설문조사3 
   BackColor       =   &H00D9E5E9&
   BorderStyle     =   1  '단일 고정
   Caption         =   "설문조사"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm설문조사3.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9450
   StartUpPosition =   1  '소유자 가운데
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   5745
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   10134
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm설문조사3.frx":08CA
      Begin Threed.SSPanel SSPanel 
         Height          =   570
         Index           =   0
         Left            =   15
         TabIndex        =   3
         Top             =   5160
         Width           =   9420
         _ExtentX        =   16616
         _ExtentY        =   1005
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   45
            TabIndex        =   0
            Top             =   60
            Width           =   1275
            _Version        =   851970
            _ExtentX        =   2249
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            BackColor       =   -2147483633
            Appearance      =   6
            Picture         =   "frm설문조사3.frx":095C
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   8100
            TabIndex        =   4
            Top             =   60
            Width           =   1260
            _Version        =   851970
            _ExtentX        =   2222
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 저장(&S)"
            BackColor       =   -2147483633
            Appearance      =   6
            Picture         =   "frm설문조사3.frx":136E
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   405
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   9420
         _ExtentX        =   16616
         _ExtentY        =   714
         _Version        =   262144
         Font3D          =   3
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "   설문조사 - 2011년 2분기 가맹점 운영 평가 시험 문제"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm설문조사3.frx":1D80
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image1 
            Height          =   240
            Left            =   60
            Picture         =   "frm설문조사3.frx":21E2
            Top             =   60
            Width           =   240
         End
      End
      Begin XtremeSuiteControls.TabControl TabControl1 
         Height          =   4050
         Left            =   15
         TabIndex        =   5
         Top             =   1095
         Width           =   9420
         _Version        =   851970
         _ExtentX        =   16616
         _ExtentY        =   7144
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   4
         Color           =   8
         PaintManager.Position=   2
         PaintManager.BoldSelected=   -1  'True
         PaintManager.HotTracking=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         PaintManager.ButtonMargin=   "2,3,2,3"
         ItemCount       =   10
         Item(0).Caption =   "1번"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "TabControlPage1"
         Item(1).Caption =   "2번"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "TabControlPage2"
         Item(2).Caption =   "3번"
         Item(2).ControlCount=   1
         Item(2).Control(0)=   "TabControlPage3"
         Item(3).Caption =   "4번"
         Item(3).ControlCount=   1
         Item(3).Control(0)=   "TabControlPage4"
         Item(4).Caption =   "5번"
         Item(4).ControlCount=   1
         Item(4).Control(0)=   "TabControlPage5"
         Item(5).Caption =   "6번"
         Item(5).ControlCount=   1
         Item(5).Control(0)=   "TabControlPage6"
         Item(6).Caption =   "7번"
         Item(6).ControlCount=   1
         Item(6).Control(0)=   "TabControlPage7"
         Item(7).Caption =   "8번"
         Item(7).ControlCount=   1
         Item(7).Control(0)=   "TabControlPage8"
         Item(8).Caption =   "9번"
         Item(8).ControlCount=   1
         Item(8).Control(0)=   "TabControlPage9"
         Item(9).Caption =   "10번"
         Item(9).ControlCount=   1
         Item(9).Control(0)=   "TabControlPage10"
         Begin XtremeSuiteControls.TabControlPage TabControlPage10 
            Height          =   3555
            Left            =   -69970
            TabIndex        =   6
            Top             =   30
            Visible         =   0   'False
            Width           =   9360
            _Version        =   851970
            _ExtentX        =   16510
            _ExtentY        =   6271
            _StockProps     =   1
            BackColor       =   16777215
            Page            =   9
            Begin Threed.SSOption optCheck10 
               Height          =   285
               Index           =   0
               Left            =   165
               TabIndex        =   7
               Top             =   540
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "1. 세탁물을 작업자가 판단해서 세탁한다."
            End
            Begin Threed.SSOption optCheck10 
               Height          =   285
               Index           =   1
               Left            =   165
               TabIndex        =   8
               Top             =   885
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "2. 오염물질에 따라 세탁을 진행한다."
            End
            Begin Threed.SSOption optCheck10 
               Height          =   285
               Index           =   2
               Left            =   165
               TabIndex        =   9
               Top             =   1230
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "3. 의류에 부착된 세탁표기대로 세탁한다."
            End
            Begin Threed.SSOption optCheck10 
               Height          =   285
               Index           =   3
               Left            =   165
               TabIndex        =   10
               Top             =   1575
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "4. 드라이클리닝 후 구분해서 물세탁을 한다."
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "10.다음중 올바른 세탁은?"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   9
               Left            =   135
               TabIndex        =   11
               Top             =   195
               Width           =   2745
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage9 
            Height          =   3555
            Left            =   -69970
            TabIndex        =   12
            Top             =   30
            Visible         =   0   'False
            Width           =   9360
            _Version        =   851970
            _ExtentX        =   16510
            _ExtentY        =   6271
            _StockProps     =   1
            BackColor       =   16777215
            Page            =   8
            Begin Threed.SSOption optCheck9 
               Height          =   285
               Index           =   0
               Left            =   165
               TabIndex        =   13
               Top             =   540
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "1. 텍은 가맹점 번호만 일치하면 순서없이 부착한다."
            End
            Begin Threed.SSOption optCheck9 
               Height          =   285
               Index           =   1
               Left            =   165
               TabIndex        =   14
               Top             =   885
               Width           =   7905
               _ExtentX        =   13944
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "2. 오염제거텍은 반드시 오염을 지운다고 고객과 약속한 세탁물만 부착한다."
            End
            Begin Threed.SSOption optCheck9 
               Height          =   285
               Index           =   2
               Left            =   165
               TabIndex        =   15
               Top             =   1230
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "3. 빨간텍은 메시지 텍이다."
            End
            Begin Threed.SSOption optCheck9 
               Height          =   285
               Index           =   3
               Left            =   165
               TabIndex        =   16
               Top             =   1575
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "4. 텍을 부착하기 어려울때는 옷핀을 활용한다."
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "9.텍 부착 내용중 맞는 것은?"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   8
               Left            =   135
               TabIndex        =   17
               Top             =   195
               Width           =   3090
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage8 
            Height          =   3555
            Left            =   -69970
            TabIndex        =   18
            Top             =   30
            Visible         =   0   'False
            Width           =   9360
            _Version        =   851970
            _ExtentX        =   16510
            _ExtentY        =   6271
            _StockProps     =   1
            BackColor       =   16777215
            Page            =   7
            Begin Threed.SSOption optCheck8 
               Height          =   285
               Index           =   0
               Left            =   165
               TabIndex        =   19
               Top             =   540
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "1. 사고시 신속히 무조건 처리한다."
            End
            Begin Threed.SSOption optCheck8 
               Height          =   285
               Index           =   1
               Left            =   165
               TabIndex        =   20
               Top             =   885
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "2. 복구가 가능하지만 고객의 입장에서 새옷으로 구입해준다."
            End
            Begin Threed.SSOption optCheck8 
               Height          =   285
               Index           =   2
               Left            =   165
               TabIndex        =   21
               Top             =   1230
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "3. 복구가 불가능한 세탁물은 책임소재를 가려 배상한다."
            End
            Begin Threed.SSOption optCheck8 
               Height          =   285
               Index           =   3
               Left            =   165
               TabIndex        =   22
               Top             =   1575
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "4. 배상은 고객 구입가 기준으로 배상한다."
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "8.세탁물 사고처리 방법 중 옳은것은?"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   7
               Left            =   135
               TabIndex        =   23
               Top             =   195
               Width           =   3990
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage7 
            Height          =   3555
            Left            =   -69970
            TabIndex        =   24
            Top             =   30
            Visible         =   0   'False
            Width           =   9360
            _Version        =   851970
            _ExtentX        =   16510
            _ExtentY        =   6271
            _StockProps     =   1
            BackColor       =   16777215
            Page            =   6
            Begin Threed.SSOption optCheck7 
               Height          =   285
               Index           =   0
               Left            =   165
               TabIndex        =   25
               Top             =   540
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "1. 일반 세탁물은 수거백에 담아 입고한다."
            End
            Begin Threed.SSOption optCheck7 
               Height          =   285
               Index           =   1
               Left            =   165
               TabIndex        =   26
               Top             =   885
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "2. 급자 세탁물은 비닐백에 담아  입고한다."
            End
            Begin Threed.SSOption optCheck7 
               Height          =   285
               Index           =   2
               Left            =   165
               TabIndex        =   27
               Top             =   1230
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "3. 와이셔츠는 수거백에 담아 입고한다.(드라이용 Y셔츠 제외)"
            End
            Begin Threed.SSOption optCheck7 
               Height          =   285
               Index           =   3
               Left            =   165
               TabIndex        =   28
               Top             =   1575
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "4. 운동화류는비닐백에 담아 입고한다. "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "7. 세탁물 입고시 유의사항이 아닌것은?"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   6
               Left            =   135
               TabIndex        =   29
               Top             =   195
               Width           =   4215
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage6 
            Height          =   3555
            Left            =   -69970
            TabIndex        =   30
            Top             =   30
            Visible         =   0   'False
            Width           =   9360
            _Version        =   851970
            _ExtentX        =   16510
            _ExtentY        =   6271
            _StockProps     =   1
            BackColor       =   16777215
            Page            =   5
            Begin Threed.SSOption optCheck6 
               Height          =   285
               Index           =   0
               Left            =   165
               TabIndex        =   31
               Top             =   540
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "1. 페인트 오염"
            End
            Begin Threed.SSOption optCheck6 
               Height          =   285
               Index           =   1
               Left            =   165
               TabIndex        =   32
               Top             =   885
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "2. 부분적으로 가죽이 부착된 의류"
            End
            Begin Threed.SSOption optCheck6 
               Height          =   285
               Index           =   2
               Left            =   165
               TabIndex        =   33
               Top             =   1230
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "3. 장식이 부착된 의류"
            End
            Begin Threed.SSOption optCheck6 
               Height          =   285
               Index           =   3
               Left            =   165
               TabIndex        =   34
               Top             =   1575
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "4. 청바지류"
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "6. 다음중 할증을 받지않는 것은?"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   5
               Left            =   135
               TabIndex        =   35
               Top             =   195
               Width           =   3540
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage5 
            Height          =   3555
            Left            =   -69970
            TabIndex        =   36
            Top             =   30
            Visible         =   0   'False
            Width           =   9360
            _Version        =   851970
            _ExtentX        =   16510
            _ExtentY        =   6271
            _StockProps     =   1
            BackColor       =   16777215
            Page            =   4
            Begin Threed.SSOption optCheck5 
               Height          =   465
               Index           =   0
               Left            =   165
               TabIndex        =   37
               Top             =   540
               Width           =   9015
               _ExtentX        =   15901
               _ExtentY        =   820
               _Version        =   262144
               CaptionStyle    =   1
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "1. 보관증은 물품인수시 꼭 지참하시고 물품에 하자가 없을시 보관증을 반납후 물품을 인      수한다."
            End
            Begin Threed.SSOption optCheck5 
               Height          =   465
               Index           =   1
               Left            =   165
               TabIndex        =   38
               Top             =   1035
               Width           =   8505
               _ExtentX        =   15002
               _ExtentY        =   820
               _Version        =   262144
               CaptionStyle    =   1
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "2. 세탁물 사고품에 대해서는 한국소비자원의 규정을 토대로 피해보상 처리를 한다."
            End
            Begin Threed.SSOption optCheck5 
               Height          =   465
               Index           =   2
               Left            =   165
               TabIndex        =   39
               Top             =   1530
               Width           =   7500
               _ExtentX        =   13229
               _ExtentY        =   820
               _Version        =   262144
               CaptionStyle    =   1
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "3. 10개월 이내에 찾아가지 않은 물품에 대해서는 책임을 지지 않습니다."
            End
            Begin Threed.SSOption optCheck5 
               Height          =   465
               Index           =   3
               Left            =   165
               TabIndex        =   40
               Top             =   2025
               Width           =   8520
               _ExtentX        =   15028
               _ExtentY        =   820
               _Version        =   262144
               CaptionStyle    =   1
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "4. 세탁물 인수후에는 반드시 비닐 커버를 벗기시어 통풍이 잘되는 곳에서 보관한다"
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "5. 세탁물보관증의 고객 유의사항이 아닌것은?"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   4
               Left            =   135
               TabIndex        =   41
               Top             =   195
               Width           =   4890
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage4 
            Height          =   3555
            Left            =   -69970
            TabIndex        =   42
            Top             =   30
            Visible         =   0   'False
            Width           =   9360
            _Version        =   851970
            _ExtentX        =   16510
            _ExtentY        =   6271
            _StockProps     =   1
            BackColor       =   16777215
            Page            =   3
            Begin Threed.SSOption optCheck4 
               Height          =   285
               Index           =   0
               Left            =   165
               TabIndex        =   43
               Top             =   540
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "1. 주머니를 확인한다."
            End
            Begin Threed.SSOption optCheck4 
               Height          =   285
               Index           =   1
               Left            =   165
               TabIndex        =   44
               Top             =   885
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "2. 하자가 있을시에는 고객에게 안내한다."
            End
            Begin Threed.SSOption optCheck4 
               Height          =   285
               Index           =   2
               Left            =   165
               TabIndex        =   45
               Top             =   1230
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "3. 하자가 있는 경우 기록하는 습관을 갖는다."
            End
            Begin Threed.SSOption optCheck4 
               Height          =   285
               Index           =   3
               Left            =   165
               TabIndex        =   46
               Top             =   1575
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "4. 오염은 세탁후 제거가 되므로 그대로 입고한다."
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "4. 세탁물 검품시 주의사항이 아닌것은?"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   3
               Left            =   135
               TabIndex        =   47
               Top             =   195
               Width           =   4215
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage3 
            Height          =   3555
            Left            =   -69970
            TabIndex        =   48
            Top             =   30
            Visible         =   0   'False
            Width           =   9360
            _Version        =   851970
            _ExtentX        =   16510
            _ExtentY        =   6271
            _StockProps     =   1
            BackColor       =   16777215
            Page            =   2
            Begin Threed.SSOption optCheck3 
               Height          =   285
               Index           =   0
               Left            =   165
               TabIndex        =   49
               Top             =   540
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "1. 품목. 텍번호. 색상. 내용, 상표를 정확히 입력한다."
            End
            Begin Threed.SSOption optCheck3 
               Height          =   285
               Index           =   1
               Left            =   165
               TabIndex        =   50
               Top             =   885
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "2. 품목. 텍번호. 색상, 구입장소를 정확히 입력한다."
            End
            Begin Threed.SSOption optCheck3 
               Height          =   285
               Index           =   2
               Left            =   165
               TabIndex        =   51
               Top             =   1230
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "3. 품목. 텍번호. 색상. 고객 인상착의를 정확히 입력한다."
            End
            Begin Threed.SSOption optCheck3 
               Height          =   285
               Index           =   3
               Left            =   165
               TabIndex        =   52
               Top             =   1575
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "4. 품목. 텍번호. 색상. 구입가격을 정확히 입력한다."
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "3. 일반 세탁물 접수시 유의사항이 맞게 나열 한것은?"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   135
               TabIndex        =   53
               Top             =   195
               Width           =   5700
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage2 
            Height          =   3555
            Left            =   -69970
            TabIndex        =   54
            Top             =   30
            Visible         =   0   'False
            Width           =   9360
            _Version        =   851970
            _ExtentX        =   16510
            _ExtentY        =   6271
            _StockProps     =   1
            BackColor       =   16777215
            Page            =   1
            Begin Threed.SSOption optCheck2 
               Height          =   285
               Index           =   0
               Left            =   165
               TabIndex        =   55
               Top             =   540
               Width           =   7000
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "1. 유용성(기름성분)얼룩제거에 용이하다."
            End
            Begin Threed.SSOption optCheck2 
               Height          =   285
               Index           =   1
               Left            =   165
               TabIndex        =   56
               Top             =   885
               Width           =   7000
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "2. 섬유안정성이 좋다."
            End
            Begin Threed.SSOption optCheck2 
               Height          =   285
               Index           =   2
               Left            =   165
               TabIndex        =   57
               Top             =   1230
               Width           =   7000
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "3. 시간과능률이 좋다."
            End
            Begin Threed.SSOption optCheck2 
               Height          =   285
               Index           =   3
               Left            =   165
               TabIndex        =   58
               Top             =   1575
               Width           =   7000
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "4. 수용성 얼룩제거에 용이하다."
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "2. 드라이클리닝의 장점이 아닌것은?"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   135
               TabIndex        =   59
               Top             =   195
               Width           =   3870
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage1 
            Height          =   3555
            Left            =   30
            TabIndex        =   60
            Top             =   30
            Width           =   9360
            _Version        =   851970
            _ExtentX        =   16510
            _ExtentY        =   6271
            _StockProps     =   1
            BackColor       =   16777215
            Page            =   0
            Begin Threed.SSOption optCheck1 
               Height          =   285
               Index           =   0
               Left            =   165
               TabIndex        =   61
               Top             =   540
               Width           =   7000
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "1. 물세탁"
            End
            Begin Threed.SSOption optCheck1 
               Height          =   285
               Index           =   1
               Left            =   165
               TabIndex        =   62
               Top             =   885
               Width           =   7000
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "2. 다림질"
            End
            Begin Threed.SSOption optCheck1 
               Height          =   285
               Index           =   2
               Left            =   165
               TabIndex        =   63
               Top             =   1230
               Width           =   7000
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "3. 스파팅(용제처리)"
            End
            Begin Threed.SSOption optCheck1 
               Height          =   285
               Index           =   3
               Left            =   165
               TabIndex        =   64
               Top             =   1575
               Width           =   7000
               _ExtentX        =   12356
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "4. 검품"
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "1. 세탁은 크게 드라이크리닝과 (       ) 으로 이루어져있다."
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   135
               TabIndex        =   65
               Top             =   195
               Width           =   6660
            End
         End
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   645
         Index           =   1
         Left            =   15
         TabIndex        =   66
         Top             =   435
         Width           =   9420
         _ExtentX        =   16616
         _ExtentY        =   1138
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "1번 부터 10번까지 모두 체크한 후 저장 버튼을 클릭해 주세요."
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
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   67
            Top             =   195
            Width           =   7620
         End
      End
   End
End
Attribute VB_Name = "frm설문조사3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn

    
    
    Select Case Index
        Case 0
            Select Case True
                Case optCheck1(0).Value:
                Case optCheck1(1).Value:
                Case optCheck1(2).Value:
                Case optCheck1(3).Value:
                Case Else:
                    MsgBox "'1번문제'를 체크해주세요.", vbInformation, "확인"
                    TabControl1.SelectedItem = 0
                    Exit Sub
            End Select
        
            Select Case True
                Case optCheck2(0).Value:
                Case optCheck2(1).Value:
                Case optCheck2(2).Value:
                Case optCheck2(3).Value:
                Case Else:
                    MsgBox "'2번문제'를 체크해주세요.", vbInformation, "확인"
                    TabControl1.SelectedItem = 1
                    Exit Sub
            End Select
        
            Select Case True
                Case optCheck3(0).Value:
                Case optCheck3(1).Value:
                Case optCheck3(2).Value:
                Case optCheck3(3).Value:
                Case Else:
                    MsgBox "'3번문제'를 체크해주세요.", vbInformation, "확인"
                    TabControl1.SelectedItem = 2
                    Exit Sub
            End Select
        
            Select Case True
                Case optCheck4(0).Value:
                Case optCheck4(1).Value:
                Case optCheck4(2).Value:
                Case optCheck4(3).Value:
                Case Else:
                    MsgBox "'4번문제'를 체크해주세요.", vbInformation, "확인"
                    TabControl1.SelectedItem = 3
                    Exit Sub
            End Select
        
            Select Case True
                Case optCheck5(0).Value:
                Case optCheck5(1).Value:
                Case optCheck5(2).Value:
                Case optCheck5(3).Value:
                Case Else:
                    MsgBox "'6번문제'를 체크해주세요.", vbInformation, "확인"
                    TabControl1.SelectedItem = 4
                    Exit Sub
            End Select
        
            Select Case True
                Case optCheck6(0).Value:
                Case optCheck6(1).Value:
                Case optCheck6(2).Value:
                Case optCheck6(3).Value:
                Case Else:
                    MsgBox "'6번문제'를 체크해주세요.", vbInformation, "확인"
                    TabControl1.SelectedItem = 5
                    Exit Sub
            End Select
        
            Select Case True
                Case optCheck7(0).Value:
                Case optCheck7(1).Value:
                Case optCheck7(2).Value:
                Case optCheck7(3).Value:
                Case Else:
                    MsgBox "'7번문제'를 체크해주세요.", vbInformation, "확인"
                    TabControl1.SelectedItem = 6
                    Exit Sub
            End Select
        
            Select Case True
                Case optCheck8(0).Value:
                Case optCheck8(1).Value:
                Case optCheck8(2).Value:
                Case optCheck8(3).Value:
                Case Else:
                    MsgBox "'8번문제'를 체크해주세요.", vbInformation, "확인"
                    TabControl1.SelectedItem = 7
                    Exit Sub
            End Select
        
            Select Case True
                Case optCheck9(0).Value:
                Case optCheck9(1).Value:
                Case optCheck9(2).Value:
                Case optCheck9(3).Value:
                Case Else:
                    MsgBox "'9번문제'를 체크해주세요.", vbInformation, "확인"
                    TabControl1.SelectedItem = 8
                    Exit Sub
            End Select
        
            Select Case True
                Case optCheck10(0).Value:
                Case optCheck10(1).Value:
                Case optCheck10(2).Value:
                Case optCheck10(3).Value:
                Case Else:
                    MsgBox "'10번문제'를 체크해주세요.", vbInformation, "확인"
                    TabControl1.SelectedItem = 9
                    Exit Sub
            End Select
            
            Query = "SELECT * FROM TB_설문조사"
            Query = Query & " WHERE KEY_DATE   = '110706'"
            Query = Query & "   AND 가맹점코드 = '" & 가맹점정보.가맹점코드 & "'"
            Set ADORs = New ADODB.Recordset
            ADORs.Open Query, HostCon, adOpenDynamic, adLockOptimistic
        
            If ADORs.EOF Then ADORs.AddNew
            
            ADORs!KEY_DATE = "110706"                   '
            ADORs!가맹점코드 = 가맹점정보.가맹점코드 & ""
            ADORs!지사코드 = 가맹점정보.지사코드 & ""
            
            Select Case True
                Case optCheck1(0).Value: ADORs!항목1 = 1
                Case optCheck1(1).Value: ADORs!항목1 = 2
                Case optCheck1(2).Value: ADORs!항목1 = 3
                Case optCheck1(3).Value: ADORs!항목1 = 4
                Case Else: ADORs!항목1 = 0
            End Select
        
            Select Case True
                Case optCheck2(0).Value: ADORs!항목2 = 1
                Case optCheck2(1).Value: ADORs!항목2 = 2
                Case optCheck2(2).Value: ADORs!항목2 = 3
                Case optCheck2(3).Value: ADORs!항목2 = 4
                Case Else: ADORs!항목2 = 0
            End Select
        
            Select Case True
                Case optCheck3(0).Value: ADORs!항목3 = 1
                Case optCheck3(1).Value: ADORs!항목3 = 2
                Case optCheck3(2).Value: ADORs!항목3 = 3
                Case optCheck3(3).Value: ADORs!항목3 = 4
                Case Else: ADORs!항목3 = 0
            End Select
        
            Select Case True
                Case optCheck4(0).Value: ADORs!항목4 = 1
                Case optCheck4(1).Value: ADORs!항목4 = 2
                Case optCheck4(2).Value: ADORs!항목4 = 3
                Case optCheck4(3).Value: ADORs!항목4 = 4
                Case Else: ADORs!항목4 = 0
            End Select
        
            Select Case True
                Case optCheck5(0).Value: ADORs!항목5 = 1
                Case optCheck5(1).Value: ADORs!항목5 = 2
                Case optCheck5(2).Value: ADORs!항목5 = 3
                Case optCheck5(3).Value: ADORs!항목5 = 4
                Case Else: ADORs!항목5 = 0
            End Select
            
            Select Case True
                Case optCheck6(0).Value: ADORs!항목6 = 1
                Case optCheck6(1).Value: ADORs!항목6 = 2
                Case optCheck6(2).Value: ADORs!항목6 = 3
                Case optCheck6(3).Value: ADORs!항목6 = 4
                Case Else: ADORs!항목6 = 0
            End Select
        
            Select Case True
                Case optCheck7(0).Value: ADORs!항목7 = 1
                Case optCheck7(1).Value: ADORs!항목7 = 2
                Case optCheck7(2).Value: ADORs!항목7 = 3
                Case optCheck7(3).Value: ADORs!항목7 = 4
                Case Else: ADORs!항목7 = 0
            End Select
        
            Select Case True
                Case optCheck8(0).Value: ADORs!항목8 = 1
                Case optCheck8(1).Value: ADORs!항목8 = 2
                Case optCheck8(2).Value: ADORs!항목8 = 3
                Case optCheck8(3).Value: ADORs!항목8 = 4
                Case Else: ADORs!항목8 = 0
            End Select
        
            Select Case True
                Case optCheck9(0).Value: ADORs!항목9 = 1
                Case optCheck9(1).Value: ADORs!항목9 = 2
                Case optCheck9(2).Value: ADORs!항목9 = 3
                Case optCheck9(3).Value: ADORs!항목9 = 4
                Case Else: ADORs!항목9 = 0
            End Select
        
            Select Case True
                Case optCheck10(0).Value: ADORs!항목10 = 1
                Case optCheck10(1).Value: ADORs!항목10 = 2
                Case optCheck10(2).Value: ADORs!항목10 = 3
                Case optCheck10(3).Value: ADORs!항목10 = 4
                Case Else: ADORs!항목10 = 0
            End Select
            
            ADORs!설문일자 = Format(Now, "YYYY-MM-DD hh:mm:ss")
            
            ADORs.Update
            
            ADORs.Close
            Set ADORs = Nothing
            
            Unload Me
            
        Case 1:
            Unload Me
    End Select

    Exit Sub

ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)

    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    If Server_Connection(HostCon, "LAUNDRY1000") = True Then
        Query = "SELECT   항목1"
        Query = Query & ",항목2"
        Query = Query & ",항목3"
        Query = Query & ",항목4"
        Query = Query & ",항목5"
        Query = Query & ",항목6"
        Query = Query & ",항목7"
        Query = Query & ",항목8"
        Query = Query & ",항목9"
        Query = Query & ",항목10"
        Query = Query & " FROM TB_설문조사"
        Query = Query & " WHERE KEY_DATE   = '110706'"
        Query = Query & "   AND 가맹점코드 = '" & 가맹점정보.가맹점코드 & "'"
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, HostCon, adOpenForwardOnly, adLockReadOnly
                
        If Not ADORs.EOF Then
            '항목1
            If IsNull(ADORs!항목1) Then
                optCheck1(0).Value = False
                optCheck1(1).Value = False
                optCheck1(2).Value = False
                optCheck1(3).Value = False
            Else
                Select Case ADORs!항목1
                    Case "1": optCheck1(0).Value = True
                    Case "2": optCheck1(1).Value = True
                    Case "3": optCheck1(2).Value = True
                    Case "4": optCheck1(3).Value = True
                End Select
            End If
            
            '항목2
            If IsNull(ADORs!항목2) Then
                optCheck2(0).Value = False
                optCheck2(1).Value = False
                optCheck2(2).Value = False
                optCheck2(3).Value = False
            Else
                Select Case ADORs!항목2
                    Case "1": optCheck2(0).Value = True
                    Case "2": optCheck2(1).Value = True
                    Case "3": optCheck2(2).Value = True
                    Case "4": optCheck2(3).Value = True
                End Select
            End If
            
            '항목3
            If IsNull(ADORs!항목3) Then
                optCheck3(0).Value = False
                optCheck3(1).Value = False
                optCheck3(2).Value = False
                optCheck3(3).Value = False
            Else
                Select Case ADORs!항목3
                    Case "1": optCheck3(0).Value = True
                    Case "2": optCheck3(1).Value = True
                    Case "3": optCheck3(2).Value = True
                    Case "4": optCheck3(3).Value = True
                End Select
            End If
            
            '항목4
            If IsNull(ADORs!항목4) Then
                optCheck4(0).Value = False
                optCheck4(1).Value = False
                optCheck4(2).Value = False
                optCheck4(3).Value = False
            Else
                Select Case ADORs!항목4
                    Case "1": optCheck4(0).Value = True
                    Case "2": optCheck4(1).Value = True
                    Case "3": optCheck4(2).Value = True
                    Case "4": optCheck4(3).Value = True
                End Select
            End If
            
            '항목5
            If IsNull(ADORs!항목5) Then
                optCheck5(0).Value = False
                optCheck5(1).Value = False
                optCheck5(2).Value = False
                optCheck5(3).Value = False
            Else
                Select Case ADORs!항목5
                    Case "1": optCheck5(0).Value = True
                    Case "2": optCheck5(1).Value = True
                    Case "3": optCheck5(2).Value = True
                    Case "4": optCheck5(3).Value = True
                End Select
            End If
        
            '항목6
            If IsNull(ADORs!항목6) Then
                optCheck6(0).Value = False
                optCheck6(1).Value = False
                optCheck6(2).Value = False
                optCheck6(3).Value = False
            Else
                Select Case ADORs!항목6
                    Case "1": optCheck6(0).Value = True
                    Case "2": optCheck6(1).Value = True
                    Case "3": optCheck6(2).Value = True
                    Case "4": optCheck6(3).Value = True
                End Select
            End If
        
            '항목7
            If IsNull(ADORs!항목7) Then
                optCheck7(0).Value = False
                optCheck7(1).Value = False
                optCheck7(2).Value = False
                optCheck7(3).Value = False
            Else
                Select Case ADORs!항목7
                    Case "1": optCheck7(0).Value = True
                    Case "2": optCheck7(1).Value = True
                    Case "3": optCheck7(2).Value = True
                    Case "4": optCheck7(3).Value = True
                End Select
            End If
        
            '항목8
            If IsNull(ADORs!항목8) Then
                optCheck8(0).Value = False
                optCheck8(1).Value = False
                optCheck8(2).Value = False
                optCheck8(3).Value = False
            Else
                Select Case ADORs!항목8
                    Case "1": optCheck8(0).Value = True
                    Case "2": optCheck8(1).Value = True
                    Case "3": optCheck8(2).Value = True
                    Case "4": optCheck8(3).Value = True
                End Select
            End If
        
            '항목9
            If IsNull(ADORs!항목9) Then
                optCheck9(0).Value = False
                optCheck9(1).Value = False
                optCheck9(2).Value = False
                optCheck9(3).Value = False
            Else
                Select Case ADORs!항목9
                    Case "1": optCheck9(0).Value = True
                    Case "2": optCheck9(1).Value = True
                    Case "3": optCheck9(2).Value = True
                    Case "4": optCheck9(3).Value = True
                End Select
            End If
        
            '항목10
            If IsNull(ADORs!항목10) Then
                optCheck10(0).Value = False
                optCheck10(1).Value = False
                optCheck10(2).Value = False
                optCheck10(3).Value = False
            Else
                Select Case ADORs!항목10
                    Case "1": optCheck10(0).Value = True
                    Case "2": optCheck10(1).Value = True
                    Case "3": optCheck10(2).Value = True
                    Case "4": optCheck10(3).Value = True
                End Select
            End If
        End If
        ADORs.Close
        Set ADORs = Nothing
    End If
    
    Exit Sub

ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)

    Screen.MousePointer = 0
End Sub

'Private Sub optCheck1_Click(Index As Integer, Value As Integer)
'    TabControl1.SelectedItem = 1
'End Sub
'
'Private Sub optCheck2_Click(Index As Integer, Value As Integer)
'    TabControl1.SelectedItem = 2
'End Sub
'
'Private Sub optCheck3_Click(Index As Integer, Value As Integer)
'    TabControl1.SelectedItem = 3
'End Sub
'
'Private Sub optCheck4_Click(Index As Integer, Value As Integer)
'    TabControl1.SelectedItem = 4
'End Sub
'
'Private Sub optCheck5_Click(Index As Integer, Value As Integer)
'    TabControl1.SelectedItem = 5
'End Sub
'
'Private Sub optCheck6_Click(Index As Integer, Value As Integer)
'    TabControl1.SelectedItem = 6
'End Sub
'
'Private Sub optCheck7_Click(Index As Integer, Value As Integer)
'    TabControl1.SelectedItem = 7
'End Sub
'
'Private Sub optCheck8_Click(Index As Integer, Value As Integer)
'    TabControl1.SelectedItem = 8
'End Sub
'
'Private Sub optCheck9_Click(Index As Integer, Value As Integer)
'    TabControl1.SelectedItem = 9
'End Sub
'
