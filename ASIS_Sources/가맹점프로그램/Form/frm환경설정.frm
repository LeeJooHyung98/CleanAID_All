VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{83FD3014-2044-4BA5-9B6C-F0A2482D9C0C}#1.0#0"; "KICCPOSIEX.OCX"
Begin VB.Form frm환경설정 
   BorderStyle     =   1  '단일 고정
   Caption         =   "가맹점 정보 - 환경설정"
   ClientHeight    =   9315
   ClientLeft      =   6405
   ClientTop       =   3960
   ClientWidth     =   13755
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm환경설정.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   13755
   Begin KiccPosIE.KiccPosIEX KiccPosOCX 
      Height          =   585
      Left            =   9660
      TabIndex        =   151
      Top             =   2235
      Visible         =   0   'False
      Width           =   825
      BF0C            =   ""
      Bmp             =   ""
      CardNo          =   ""
      CashNo          =   ""
      CommType        =   1
      Connected       =   0   'False
      Emv             =   ""
      EmvLen          =   0
      MasterClaimerText=   ""
      MasterOfferText =   ""
      PIN             =   ""
      SeqNo           =   ""
      Sign            =   ""
      SignLen         =   0
      TID             =   ""
      RfFlag          =   ""
      VAK             =   ""
      VisaClaimerText =   ""
      VisaOfferText   =   ""
      ErrMsg          =   ""
      ResMsg          =   ""
      RcvData         =   ""
      TRNO            =   ""
      Data            =   ""
      CVER            =   ""
      MVER            =   ""
      PVER            =   ""
      TMTransCount    =   0
      TMOnLineCount   =   0
      EBTransCount    =   0
      Alignment       =   2
      AutoSize        =   0   'False
      BevelInner      =   0
      BevelOuter      =   0
      BorderStyle     =   0
      Caption         =   ""
      Color           =   16777215
      Ctl3D           =   -1  'True
      UseDockManager  =   -1  'True
      DockSite        =   0   'False
      DragCursor      =   -12
      Object.DragMode        =   0
      Enabled         =   -1  'True
      FullRepaint     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   0   'False
      ParentColor     =   0   'False
      ParentCtl3D     =   -1  'True
      Object.Visible         =   -1  'True
      DoubleBuffered  =   -1  'True
      Cursor          =   0
      Protocol        =   0
      JcbClaimerText  =   ""
      JcbOfferText    =   ""
      DccTextVer      =   "00"
      CardHash        =   "$"
      SignAD          =   "0000"
      HandleValue     =   66544
      MemberShip      =   ""
      MemberShipHex   =   ""
      TCPSVCPort      =   0
      TCPSVCActive    =   0   'False
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9315
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13755
      _ExtentX        =   24262
      _ExtentY        =   16431
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm환경설정.frx":0A02
      Begin XtremeSuiteControls.TabControl TabControl1 
         Height          =   8085
         Left            =   15
         TabIndex        =   2
         Top             =   1215
         Width           =   13725
         _Version        =   851970
         _ExtentX        =   24209
         _ExtentY        =   14261
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
         Appearance      =   3
         Color           =   16
         PaintManager.BoldSelected=   -1  'True
         PaintManager.HotTracking=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         PaintManager.ButtonMargin=   "3,4,3,4"
         ItemCount       =   4
         SelectedItem    =   2
         Item(0).Caption =   " 기본정보 "
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "TabControlPage1"
         Item(1).Caption =   " 마진율 및 적립금액 "
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "TabControlPage2"
         Item(2).Caption =   " 카드 단말기(프린터) "
         Item(2).ControlCount=   1
         Item(2).Control(0)=   "TabControlPage3"
         Item(3).Caption =   " 문자서비스(SMS) "
         Item(3).ControlCount=   1
         Item(3).Control(0)=   "TabControlPage4"
         Begin XtremeSuiteControls.TabControlPage TabControlPage4 
            Height          =   7575
            Left            =   -69970
            TabIndex        =   3
            Top             =   480
            Visible         =   0   'False
            Width           =   13665
            _Version        =   851970
            _ExtentX        =   24104
            _ExtentY        =   13361
            _StockProps     =   1
            BackColor       =   16777215
            Page            =   3
            Begin Threed.SSCheck chkSMSEMART 
               Height          =   330
               Left            =   1515
               TabIndex        =   33
               Top             =   1785
               Width           =   3105
               _ExtentX        =   5477
               _ExtentY        =   582
               _Version        =   262144
               BackColor       =   12648447
               PictureFrames   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "frm환경설정.frx":0A74
               Caption         =   " 문자서비스(SMS)"
            End
            Begin VB.TextBox txtSMSIPAddress 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1425
               Locked          =   -1  'True
               TabIndex        =   7
               Top             =   120
               Width           =   4125
            End
            Begin VB.TextBox txtSMSDBName 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1425
               Locked          =   -1  'True
               TabIndex        =   6
               Top             =   510
               Width           =   4125
            End
            Begin VB.TextBox txtSMSUserName 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1425
               Locked          =   -1  'True
               TabIndex        =   5
               Top             =   900
               Width           =   4125
            End
            Begin VB.TextBox txtSMSUserPass 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               IMEMode         =   3  '사용 못함
               Left            =   1425
               Locked          =   -1  'True
               PasswordChar    =   "*"
               TabIndex        =   4
               Top             =   1290
               Width           =   4125
            End
            Begin XtremeSuiteControls.PushButton cmdSMSTest 
               Height          =   450
               Left            =   3990
               TabIndex        =   66
               Top             =   2250
               Width           =   1560
               _Version        =   851970
               _ExtentX        =   2752
               _ExtentY        =   794
               _StockProps     =   79
               Caption         =   "테스트 발송"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
            End
            Begin VB.Shape Shape 
               BackColor       =   &H00C0FFFF&
               BackStyle       =   1  '투명하지 않음
               BorderColor     =   &H00C0C000&
               Height          =   450
               Index           =   0
               Left            =   1410
               Shape           =   4  '둥근 사각형
               Top             =   1725
               Width           =   4140
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "SMS 서버 IP:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   7
               Left            =   285
               TabIndex        =   11
               Top             =   180
               Width           =   1080
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "SMS  DB:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   8
               Left            =   645
               TabIndex        =   10
               Top             =   585
               Width           =   720
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "SMS ID:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   9
               Left            =   735
               TabIndex        =   9
               Top             =   960
               Width           =   630
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "SMS  암호:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   10
               Left            =   465
               TabIndex        =   8
               Top             =   1365
               Width           =   900
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage3 
            Height          =   7575
            Left            =   30
            TabIndex        =   12
            Top             =   480
            Width           =   13665
            _Version        =   851970
            _ExtentX        =   24104
            _ExtentY        =   13361
            _StockProps     =   1
            BackColor       =   16777215
            Page            =   2
            Begin Threed.SSFrame SSFrame_KS4060 
               Height          =   2235
               Left            =   6330
               TabIndex        =   144
               Top             =   2520
               Visible         =   0   'False
               Width           =   6735
               _ExtentX        =   11880
               _ExtentY        =   3942
               _Version        =   262144
               BackColor       =   16777215
               Caption         =   "[ KSNET KS4060 설정 방법 ]"
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "5. 특수->2->5. 사인패드 공유 -> 예 설정"
                  Height          =   195
                  Index           =   4
                  Left            =   360
                  TabIndex        =   149
                  Top             =   1710
                  Width           =   4095
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "4. 특수->2->3. 거래결과 전송 -> 예 설정"
                  Height          =   195
                  Index           =   3
                  Left            =   360
                  TabIndex        =   148
                  Top             =   1380
                  Width           =   4095
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "3. 특수->2->2. PC 연동 -> 예 설정"
                  Height          =   195
                  Index           =   2
                  Left            =   360
                  TabIndex        =   147
                  Top             =   1050
                  Width           =   3465
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "2. 특수->2->1. POS 프린터 사용 -> 예 설정 (여백 0 설정)"
                  Height          =   195
                  Index           =   1
                  Left            =   360
                  TabIndex        =   146
                  Top             =   720
                  Width           =   5775
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "1. 최신 버전으로 업그레이드 한다. (확인:입력->특수)"
                  Height          =   195
                  Index           =   0
                  Left            =   360
                  TabIndex        =   145
                  Top             =   390
                  Width           =   5355
               End
            End
            Begin VB.ComboBox cboKSCAT 
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               ItemData        =   "frm환경설정.frx":14A3
               Left            =   7500
               List            =   "frm환경설정.frx":14A5
               Style           =   2  '드롭다운 목록
               TabIndex        =   142
               Top             =   150
               Width           =   1725
            End
            Begin XtremeSuiteControls.PushButton btnSignPad 
               Height          =   705
               Left            =   11070
               TabIndex        =   117
               Top             =   1320
               Width           =   1800
               _Version        =   851970
               _ExtentX        =   3175
               _ExtentY        =   1244
               _StockProps     =   79
               Caption         =   "싸인패드 테스트"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
            End
            Begin XtremeSuiteControls.PushButton btnReport 
               Height          =   705
               Left            =   11070
               TabIndex        =   116
               Top             =   510
               Width           =   1815
               _Version        =   851970
               _ExtentX        =   3201
               _ExtentY        =   1244
               _StockProps     =   79
               Caption         =   "영수증 테스트 출력"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
            End
            Begin VB.TextBox txtAddress 
               Height          =   360
               Left            =   1440
               TabIndex        =   108
               Top             =   2070
               Width           =   7785
            End
            Begin VB.TextBox txtChairman 
               Height          =   360
               Left            =   1440
               TabIndex        =   107
               Top             =   1680
               Width           =   4125
            End
            Begin VB.TextBox txtVAN 
               Height          =   360
               Index           =   0
               Left            =   1440
               TabIndex        =   98
               Top             =   120
               Width           =   4125
            End
            Begin VB.TextBox txtVAN 
               Height          =   360
               Index           =   1
               Left            =   1440
               TabIndex        =   97
               Top             =   510
               Width           =   4125
            End
            Begin VB.TextBox txtVAN 
               Height          =   360
               Index           =   2
               Left            =   1440
               TabIndex        =   96
               Top             =   1290
               Width           =   2490
            End
            Begin VB.TextBox txtVAN 
               Height          =   360
               Index           =   3
               Left            =   1440
               TabIndex        =   95
               Top             =   900
               Width           =   4140
            End
            Begin VB.ComboBox cboKS7500CommPort 
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   7500
               Style           =   2  '드롭다운 목록
               TabIndex        =   94
               Top             =   510
               Width           =   1725
            End
            Begin VB.ComboBox cboKS7500BaudRate 
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               ItemData        =   "frm환경설정.frx":14A7
               Left            =   7500
               List            =   "frm환경설정.frx":14A9
               Style           =   2  '드롭다운 목록
               TabIndex        =   93
               Top             =   900
               Width           =   1725
            End
            Begin VB.ComboBox cboSignPadCommPort 
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   7500
               Style           =   2  '드롭다운 목록
               TabIndex        =   92
               Top             =   1290
               Width           =   1725
            End
            Begin VB.ComboBox cboSignPadBaudRate 
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   7500
               Style           =   2  '드롭다운 목록
               TabIndex        =   91
               Top             =   1680
               Width           =   1725
            End
            Begin Threed.SSCheck chkTelPrt 
               Height          =   330
               Left            =   1575
               TabIndex        =   35
               Top             =   2655
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   582
               _Version        =   262144
               BackColor       =   12648447
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "고객 전화번호 모두 출력"
               Value           =   1
            End
            Begin CSTextLibCtl.silgEdit txtPaper 
               Height          =   360
               Left            =   1440
               TabIndex        =   89
               Top             =   3570
               Width           =   675
               _Version        =   262145
               _ExtentX        =   1191
               _ExtentY        =   635
               _StockProps     =   125
               Text            =   " 0"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "@굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Insert          =   0   'False
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   "0"
               Text            =   " 0"
               StartText.x     =   3
               StartText.y     =   4
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   15
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               Justification   =   1
               BorderStyle     =   0
               Undo            =   1
               Data            =   0
            End
            Begin CSTextLibCtl.silgEdit txtPaper2 
               Height          =   360
               Left            =   1440
               TabIndex        =   112
               Top             =   3960
               Width           =   675
               _Version        =   262145
               _ExtentX        =   1191
               _ExtentY        =   635
               _StockProps     =   125
               Text            =   " 0"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "@굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Insert          =   0   'False
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   "0"
               Text            =   " 0"
               StartText.x     =   3
               StartText.y     =   4
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   15
               CurNumDataChars =   0
               MaxDataChars    =   0
               FirstDataPos    =   0
               CurPos          =   0
               MaxLen          =   0
               DataReadOnly    =   0   'False
               Mask            =   ""
               Justification   =   1
               BorderStyle     =   0
               Undo            =   1
               Data            =   0
            End
            Begin XtremeSuiteControls.PushButton btnSignPadFind 
               Height          =   705
               Left            =   9270
               TabIndex        =   126
               Top             =   1320
               Width           =   1770
               _Version        =   851970
               _ExtentX        =   3122
               _ExtentY        =   1244
               _StockProps     =   79
               Caption         =   "싸인패드 찾기"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
            End
            Begin XtremeSuiteControls.PushButton btnReportFind 
               Height          =   705
               Left            =   9255
               TabIndex        =   127
               Top             =   510
               Width           =   1770
               _Version        =   851970
               _ExtentX        =   3122
               _ExtentY        =   1244
               _StockProps     =   79
               Caption         =   "영수증 프린터 찾기"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
            End
            Begin Threed.SSCheck chkStoreCardPrt 
               Height          =   330
               Left            =   1575
               TabIndex        =   141
               Top             =   3000
               Width           =   4155
               _ExtentX        =   7329
               _ExtentY        =   582
               _Version        =   262144
               BackColor       =   12648447
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "신용카드/현금영수증 보관용 전표 출력 안함"
               Value           =   1
            End
            Begin VB.Label Label_KSNET 
               BackStyle       =   0  '투명
               Caption         =   "Label4"
               Height          =   405
               Left            =   9300
               TabIndex        =   150
               Top             =   60
               Width           =   4275
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "단말기 종류:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   39
               Left            =   6360
               TabIndex        =   143
               Top             =   240
               Width           =   1080
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "- 도 입력하세요."
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   2
               Left            =   4050
               TabIndex        =   115
               Top             =   1380
               Width           =   1440
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "장"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   1
               Left            =   2205
               TabIndex        =   114
               Top             =   4050
               Width           =   180
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "출고 영수증:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   225
               TabIndex        =   113
               Top             =   4050
               Width           =   1170
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "장"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   0
               Left            =   2205
               TabIndex        =   111
               Top             =   3660
               Width           =   180
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "대표자명:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   -45
               TabIndex        =   110
               Top             =   1755
               Width           =   1425
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "사업장 주소:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   -45
               TabIndex        =   109
               Top             =   2145
               Width           =   1425
            End
            Begin VB.Shape Shape 
               BackColor       =   &H00C0FFFF&
               BackStyle       =   1  '투명하지 않음
               BorderColor     =   &H00C0C000&
               Height          =   840
               Index           =   1
               Left            =   1440
               Shape           =   4  '둥근 사각형
               Top             =   2580
               Width           =   4440
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "VAN IP:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   17
               Left            =   750
               TabIndex        =   106
               Top             =   180
               Width           =   630
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "VAN PORT:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   21
               Left            =   570
               TabIndex        =   105
               Top             =   585
               Width           =   810
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "사업자번호:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   22
               Left            =   390
               TabIndex        =   104
               Top             =   1350
               Width           =   990
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "단말기일련번호:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   23
               Left            =   30
               TabIndex        =   103
               Top             =   975
               Width           =   1350
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "단말기 속도:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   25
               Left            =   6360
               TabIndex        =   102
               Top             =   1005
               Width           =   1080
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "단말기 포트:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   26
               Left            =   6360
               TabIndex        =   101
               Top             =   600
               Width           =   1080
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "사인패드 속도:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   27
               Left            =   6180
               TabIndex        =   100
               Top             =   1785
               Width           =   1260
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "사인패드 포트:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   28
               Left            =   6180
               TabIndex        =   99
               Top             =   1380
               Width           =   1260
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "접수 영수증:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   6
               Left            =   225
               TabIndex        =   90
               Top             =   3660
               Width           =   1170
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage2 
            Height          =   7575
            Left            =   -69970
            TabIndex        =   13
            Top             =   480
            Visible         =   0   'False
            Width           =   13665
            _Version        =   851970
            _ExtentX        =   24104
            _ExtentY        =   13361
            _StockProps     =   1
            BackColor       =   16777215
            Page            =   1
            Begin Threed.SSFrame SSFrame1 
               Height          =   1395
               Left            =   6330
               TabIndex        =   128
               Top             =   2310
               Width           =   7215
               _ExtentX        =   12726
               _ExtentY        =   2461
               _Version        =   262144
               BackColor       =   16777215
               Caption         =   "로얄티 관련"
               Begin VB.ComboBox cboRovalty 
                  Height          =   315
                  Index           =   2
                  ItemData        =   "frm환경설정.frx":14AB
                  Left            =   1920
                  List            =   "frm환경설정.frx":14B5
                  Style           =   2  '드롭다운 목록
                  TabIndex        =   134
                  Top             =   1005
                  Width           =   1365
               End
               Begin VB.ComboBox cboRovalty 
                  Height          =   315
                  Index           =   0
                  ItemData        =   "frm환경설정.frx":14C5
                  Left            =   1920
                  List            =   "frm환경설정.frx":14CF
                  Style           =   2  '드롭다운 목록
                  TabIndex        =   133
                  Top             =   240
                  Width           =   1365
               End
               Begin VB.TextBox txtRovalty 
                  Alignment       =   1  '오른쪽 맞춤
                  Height          =   315
                  Index           =   0
                  Left            =   3420
                  TabIndex        =   132
                  Top             =   210
                  Width           =   1515
               End
               Begin VB.TextBox txtRovalty 
                  Alignment       =   1  '오른쪽 맞춤
                  Height          =   315
                  Index           =   2
                  Left            =   3420
                  TabIndex        =   131
                  Top             =   990
                  Width           =   1515
               End
               Begin VB.TextBox txtRovalty 
                  Alignment       =   1  '오른쪽 맞춤
                  Height          =   315
                  Index           =   1
                  Left            =   3420
                  TabIndex        =   130
                  Top             =   600
                  Width           =   1515
               End
               Begin VB.ComboBox cboRovalty 
                  Height          =   315
                  Index           =   1
                  ItemData        =   "frm환경설정.frx":14DF
                  Left            =   1920
                  List            =   "frm환경설정.frx":14E9
                  Style           =   2  '드롭다운 목록
                  TabIndex        =   129
                  Top             =   630
                  Width           =   1365
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "% (ex. 5.5 전체 매출분의 적용 비율)"
                  Height          =   195
                  Index           =   4
                  Left            =   4950
                  TabIndex        =   140
                  Top             =   300
                  Width           =   3675
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "% (ex. 0.75)"
                  Height          =   195
                  Index           =   2
                  Left            =   4950
                  TabIndex        =   139
                  Top             =   1065
                  Width           =   1260
               End
               Begin VB.Label Label2 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "로얄티 적용 1:"
                  Height          =   255
                  Index           =   38
                  Left            =   90
                  TabIndex        =   138
                  Top             =   285
                  Width           =   1785
               End
               Begin VB.Label Label2 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "수수료 지원 적용:"
                  Height          =   255
                  Index           =   37
                  Left            =   90
                  TabIndex        =   137
                  Top             =   1065
                  Width           =   1785
               End
               Begin VB.Label Label2 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "로얄티 적용 2:"
                  Height          =   255
                  Index           =   36
                  Left            =   90
                  TabIndex        =   136
                  Top             =   645
                  Width           =   1785
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "% (ex. 1.5 매장 매출분의 적용 비율)"
                  Height          =   195
                  Index           =   1
                  Left            =   4950
                  TabIndex        =   135
                  Top             =   690
                  Width           =   3675
               End
            End
            Begin Threed.SSPanel SSPanel1 
               Height          =   2730
               Left            =   6315
               TabIndex        =   46
               Top             =   4110
               Width           =   7245
               _ExtentX        =   12779
               _ExtentY        =   4815
               _Version        =   262144
               BackColor       =   16777215
               Enabled         =   0   'False
               BorderWidth     =   0
               BevelOuter      =   1
               BevelInner      =   2
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin VB.ComboBox cboSale 
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   11.25
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  ItemData        =   "frm환경설정.frx":14F9
                  Left            =   1635
                  List            =   "frm환경설정.frx":1503
                  Locked          =   -1  'True
                  Style           =   2  '드롭다운 목록
                  TabIndex        =   49
                  Top             =   105
                  Width           =   1215
               End
               Begin VB.ComboBox cboCoupon 
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   11.25
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  ItemData        =   "frm환경설정.frx":1513
                  Left            =   1635
                  List            =   "frm환경설정.frx":151D
                  Locked          =   -1  'True
                  Style           =   2  '드롭다운 목록
                  TabIndex        =   48
                  Top             =   555
                  Width           =   1215
               End
               Begin VB.ComboBox cboReturn 
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   11.25
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  ItemData        =   "frm환경설정.frx":152D
                  Left            =   1635
                  List            =   "frm환경설정.frx":1537
                  Locked          =   -1  'True
                  Style           =   2  '드롭다운 목록
                  TabIndex        =   47
                  Top             =   1425
                  Width           =   1215
               End
               Begin MSComCtl2.DTPicker dtpSaleStart 
                  Height          =   345
                  Left            =   3675
                  TabIndex        =   50
                  Top             =   105
                  Width           =   1620
                  _ExtentX        =   2858
                  _ExtentY        =   609
                  _Version        =   393216
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  CheckBox        =   -1  'True
                  CustomFormat    =   "yyyy-MM-dd"
                  Format          =   56033281
                  CurrentDate     =   40066
               End
               Begin MSComCtl2.DTPicker dtpSaleEnd 
                  Height          =   345
                  Left            =   5535
                  TabIndex        =   51
                  Top             =   105
                  Width           =   1620
                  _ExtentX        =   2858
                  _ExtentY        =   609
                  _Version        =   393216
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  CheckBox        =   -1  'True
                  CustomFormat    =   "yyyy-MM-dd"
                  Format          =   56033281
                  CurrentDate     =   40066
               End
               Begin MSComCtl2.DTPicker dtpCouponStart 
                  Height          =   345
                  Left            =   3675
                  TabIndex        =   52
                  Top             =   555
                  Width           =   1620
                  _ExtentX        =   2858
                  _ExtentY        =   609
                  _Version        =   393216
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  CheckBox        =   -1  'True
                  CustomFormat    =   "yyyy-MM-dd"
                  Format          =   56033281
                  CurrentDate     =   40066
               End
               Begin MSComCtl2.DTPicker dtpCouponEnd 
                  Height          =   345
                  Left            =   5535
                  TabIndex        =   53
                  Top             =   555
                  Width           =   1620
                  _ExtentX        =   2858
                  _ExtentY        =   609
                  _Version        =   393216
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  CheckBox        =   -1  'True
                  CustomFormat    =   "yyyy-MM-dd"
                  Format          =   56033281
                  CurrentDate     =   40066
               End
               Begin CSTextLibCtl.sidbEdit txtSale 
                  Height          =   345
                  Left            =   2880
                  TabIndex        =   54
                  Top             =   105
                  Width           =   495
                  _Version        =   262145
                  _ExtentX        =   873
                  _ExtentY        =   609
                  _StockProps     =   125
                  Text            =   " 0"
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   9.76
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BorderEffect    =   2
                  DataProperty    =   2
                  ReadOnly        =   -1  'True
                  FocusSelect     =   -1  'True
                  Insert          =   0   'False
                  Modified        =   -1  'True
                  HideSelection   =   -1  'True
                  RawData         =   ""
                  Text            =   " 0"
                  StartText.x     =   3
                  StartText.y     =   3
                  FirstVisPos     =   0
                  HiAnchor        =   0
                  HiNew           =   0
                  CaretHeight     =   16
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
                  Height          =   345
                  Left            =   2880
                  TabIndex        =   55
                  Top             =   555
                  Width           =   495
                  _Version        =   262145
                  _ExtentX        =   873
                  _ExtentY        =   609
                  _StockProps     =   125
                  Text            =   " 0"
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   9.76
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BorderEffect    =   2
                  DataProperty    =   2
                  ReadOnly        =   -1  'True
                  FocusSelect     =   -1  'True
                  Insert          =   0   'False
                  Modified        =   -1  'True
                  HideSelection   =   -1  'True
                  RawData         =   ""
                  Text            =   " 0"
                  StartText.x     =   3
                  StartText.y     =   3
                  FirstVisPos     =   0
                  HiAnchor        =   0
                  HiNew           =   0
                  CaretHeight     =   16
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
                  Height          =   345
                  Left            =   1635
                  TabIndex        =   56
                  Top             =   990
                  Width           =   1035
                  _Version        =   262145
                  _ExtentX        =   1826
                  _ExtentY        =   609
                  _StockProps     =   125
                  Text            =   " 0"
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "@굴림체"
                     Size            =   9.76
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BorderEffect    =   2
                  DataProperty    =   2
                  ReadOnly        =   -1  'True
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
                  CaretHeight     =   15
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
               Begin VB.Label Label2 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "특정할인 사용:"
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   33
                  Left            =   75
                  TabIndex        =   65
                  Top             =   630
                  Width           =   1515
               End
               Begin VB.Label Label2 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "고가세탁 비율:"
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   32
                  Left            =   75
                  TabIndex        =   64
                  Top             =   1065
                  Width           =   1515
               End
               Begin VB.Label Label2 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "세탁비환불 사용:"
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   31
                  Left            =   75
                  TabIndex        =   63
                  Top             =   1515
                  Width           =   1515
               End
               Begin VB.Label Label2 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "지정할인 사용:"
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   29
                  Left            =   75
                  TabIndex        =   62
                  Top             =   180
                  Width           =   1515
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "~"
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
                  Left            =   5340
                  TabIndex        =   61
                  Top             =   195
                  Width           =   120
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "~"
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
                  Index           =   10
                  Left            =   5340
                  TabIndex        =   60
                  Top             =   615
                  Width           =   120
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "%"
                  Height          =   195
                  Index           =   7
                  Left            =   3405
                  TabIndex        =   59
                  Top             =   630
                  Width           =   105
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "%"
                  Height          =   195
                  Index           =   6
                  Left            =   2715
                  TabIndex        =   58
                  Top             =   1065
                  Width           =   105
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "%"
                  Height          =   195
                  Index           =   5
                  Left            =   3405
                  TabIndex        =   57
                  Top             =   180
                  Width           =   105
               End
            End
            Begin Threed.SSFrame SSFrame6 
               Height          =   2025
               Left            =   6315
               TabIndex        =   39
               Top             =   165
               Width           =   7245
               _ExtentX        =   12779
               _ExtentY        =   3572
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "마일리지"
               Begin VB.ComboBox cboMileage 
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   11.25
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  ItemData        =   "frm환경설정.frx":1547
                  Left            =   1965
                  List            =   "frm환경설정.frx":1551
                  Locked          =   -1  'True
                  Style           =   2  '드롭다운 목록
                  TabIndex        =   67
                  Top             =   285
                  Width           =   1500
               End
               Begin CSTextLibCtl.sidbEdit txtMileage 
                  Height          =   345
                  Index           =   2
                  Left            =   1875
                  TabIndex        =   44
                  Top             =   1185
                  Width           =   1230
                  _Version        =   262145
                  _ExtentX        =   2170
                  _ExtentY        =   609
                  _StockProps     =   125
                  Text            =   " 0"
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "@굴림체"
                     Size            =   9.76
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BorderEffect    =   2
                  DataProperty    =   2
                  ReadOnly        =   -1  'True
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
                  CaretHeight     =   15
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
               Begin CSTextLibCtl.sidbEdit txtMileage 
                  Height          =   345
                  Index           =   0
                  Left            =   2565
                  TabIndex        =   40
                  Top             =   720
                  Width           =   990
                  _Version        =   262145
                  _ExtentX        =   1746
                  _ExtentY        =   609
                  _StockProps     =   125
                  Text            =   " 0"
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "@굴림체"
                     Size            =   9.76
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BorderEffect    =   2
                  DataProperty    =   2
                  ReadOnly        =   -1  'True
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
                  CaretHeight     =   15
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
               Begin CSTextLibCtl.sidbEdit txtMileage 
                  Height          =   345
                  Index           =   1
                  Left            =   4110
                  TabIndex        =   41
                  Top             =   720
                  Width           =   1005
                  _Version        =   262145
                  _ExtentX        =   1773
                  _ExtentY        =   609
                  _StockProps     =   125
                  Text            =   " 0"
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "@굴림체"
                     Size            =   9.76
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BorderEffect    =   2
                  DataProperty    =   2
                  ReadOnly        =   -1  'True
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
                  CaretHeight     =   15
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
               Begin VB.Label Label2 
                  BackStyle       =   0  '투명
                  Caption         =   "마일리지 사용여부:"
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   16
                  Left            =   135
                  TabIndex        =   68
                  Top             =   375
                  Width           =   1830
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "적립된 마일리지는             원 이상 적립 되어야 사용이 가능합니다."
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   180
                  Index           =   5
                  Left            =   135
                  TabIndex        =   45
                  Top             =   1260
                  Width           =   6780
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "마일리지 적립은 이용요금           원에           원 적립 됩니다."
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   180
                  Index           =   30
                  Left            =   135
                  TabIndex        =   43
                  Top             =   810
                  Width           =   6540
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "마일리지 금액은 '100'일이 지나면 자동삭제됩니다."
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   180
                  Index           =   4
                  Left            =   135
                  TabIndex        =   42
                  Top             =   1680
                  Width           =   4755
               End
            End
            Begin Threed.SSFrame SSFrame2 
               Height          =   7425
               Index           =   2
               Left            =   135
               TabIndex        =   36
               Top             =   135
               Width           =   6150
               _ExtentX        =   10848
               _ExtentY        =   13097
               _Version        =   262144
               Font3D          =   3
               BackColor       =   16777215
               PictureFrames   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "frm환경설정.frx":1561
               Caption         =   " 마진율"
               Begin FPSpreadADO.fpSpread sprMargin 
                  Height          =   6975
                  Left            =   90
                  TabIndex        =   38
                  Top             =   345
                  Width           =   5955
                  _Version        =   524288
                  _ExtentX        =   10504
                  _ExtentY        =   12303
                  _StockProps     =   64
                  BorderStyle     =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  GrayAreaBackColor=   16777215
                  MaxCols         =   5
                  ScrollBars      =   2
                  SpreadDesigner  =   "frm환경설정.frx":1AFB
                  HighlightHeaders=   1
                  HighlightStyle  =   1
               End
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage1 
            Height          =   7575
            Left            =   -69970
            TabIndex        =   14
            Top             =   480
            Visible         =   0   'False
            Width           =   13665
            _Version        =   851970
            _ExtentX        =   24104
            _ExtentY        =   13361
            _StockProps     =   1
            BackColor       =   16777215
            Page            =   0
            Begin VB.ComboBox cboComputer 
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   6780
               Style           =   2  '드롭다운 목록
               TabIndex        =   124
               Top             =   7065
               Width           =   1110
            End
            Begin VB.TextBox txtBackupFolder 
               Height          =   645
               Left            =   1545
               MultiLine       =   -1  'True
               TabIndex        =   121
               Top             =   6390
               Width           =   6345
            End
            Begin VB.ComboBox cboMode 
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1545
               Style           =   2  '드롭다운 목록
               TabIndex        =   119
               Top             =   7065
               Width           =   1335
            End
            Begin Threed.SSFrame SSFrame 
               Height          =   3675
               Left            =   7035
               TabIndex        =   74
               Top             =   135
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   6482
               _Version        =   262144
               BackColor       =   16777215
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "세트상품 요일"
               Begin Threed.SSCheck chkWeek 
                  Height          =   255
                  Index           =   0
                  Left            =   255
                  TabIndex        =   75
                  Top             =   360
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   450
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
                  Caption         =   "일요일"
               End
               Begin Threed.SSCheck chkWeek 
                  Height          =   255
                  Index           =   1
                  Left            =   255
                  TabIndex        =   76
                  Top             =   765
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   450
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
                  Caption         =   "월요일"
               End
               Begin Threed.SSCheck chkWeek 
                  Height          =   255
                  Index           =   2
                  Left            =   255
                  TabIndex        =   77
                  Top             =   1170
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   450
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
                  Caption         =   "화요일"
               End
               Begin Threed.SSCheck chkWeek 
                  Height          =   255
                  Index           =   3
                  Left            =   255
                  TabIndex        =   78
                  Top             =   1575
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   450
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
                  Caption         =   "수요일"
               End
               Begin Threed.SSCheck chkWeek 
                  Height          =   255
                  Index           =   4
                  Left            =   255
                  TabIndex        =   79
                  Top             =   2010
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   450
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
                  Caption         =   "목요일"
               End
               Begin Threed.SSCheck chkWeek 
                  Height          =   255
                  Index           =   5
                  Left            =   255
                  TabIndex        =   80
                  Top             =   2415
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   450
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
                  Caption         =   "금요일"
               End
               Begin Threed.SSCheck chkWeek 
                  Height          =   255
                  Index           =   6
                  Left            =   255
                  TabIndex        =   81
                  Top             =   2835
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   450
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
                  Caption         =   "토요일"
               End
            End
            Begin VB.TextBox txtPWD 
               Height          =   360
               Left            =   1545
               TabIndex        =   72
               Top             =   3330
               Width           =   3345
            End
            Begin Threed.SSFrame SSFrame3 
               Height          =   3675
               Left            =   5430
               TabIndex        =   32
               Top             =   135
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   6482
               _Version        =   262144
               BackColor       =   16777215
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "요일 세일"
               Begin Threed.SSCheck chkSale 
                  Height          =   255
                  Index           =   0
                  Left            =   210
                  TabIndex        =   82
                  Top             =   360
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   450
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
                  Caption         =   "일요일"
               End
               Begin Threed.SSCheck chkSale 
                  Height          =   255
                  Index           =   1
                  Left            =   210
                  TabIndex        =   83
                  Top             =   765
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   450
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
                  Caption         =   "월요일"
               End
               Begin Threed.SSCheck chkSale 
                  Height          =   255
                  Index           =   2
                  Left            =   210
                  TabIndex        =   84
                  Top             =   1170
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   450
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
                  Caption         =   "화요일"
               End
               Begin Threed.SSCheck chkSale 
                  Height          =   255
                  Index           =   3
                  Left            =   210
                  TabIndex        =   85
                  Top             =   1575
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   450
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
                  Caption         =   "수요일"
               End
               Begin Threed.SSCheck chkSale 
                  Height          =   255
                  Index           =   4
                  Left            =   210
                  TabIndex        =   86
                  Top             =   2010
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   450
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
                  Caption         =   "목요일"
               End
               Begin Threed.SSCheck chkSale 
                  Height          =   255
                  Index           =   5
                  Left            =   210
                  TabIndex        =   87
                  Top             =   2415
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   450
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
                  Caption         =   "금요일"
               End
               Begin Threed.SSCheck chkSale 
                  Height          =   255
                  Index           =   6
                  Left            =   210
                  TabIndex        =   88
                  Top             =   2835
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   450
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
                  Caption         =   "토요일"
               End
            End
            Begin VB.TextBox txtStartDate 
               Alignment       =   2  '가운데 맞춤
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1545
               Locked          =   -1  'True
               TabIndex        =   31
               Top             =   990
               Width           =   1350
            End
            Begin XtremeSuiteControls.PushButton btnStoreInfoDownload 
               Height          =   375
               Left            =   2700
               TabIndex        =   30
               Top             =   135
               Width           =   2175
               _Version        =   851970
               _ExtentX        =   3836
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "가맹점 정보 내려받기"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
            End
            Begin VB.TextBox txtStoreCode 
               Alignment       =   2  '가운데 맞춤
               BackColor       =   &H0080FFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   1545
               TabIndex        =   0
               Top             =   135
               Width           =   1125
            End
            Begin VB.TextBox txtStoreName 
               Height          =   360
               Left            =   1545
               Locked          =   -1  'True
               TabIndex        =   20
               Top             =   600
               Width           =   3345
            End
            Begin VB.TextBox txtMstCode 
               Alignment       =   2  '가운데 맞춤
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1545
               Locked          =   -1  'True
               TabIndex        =   19
               Top             =   1380
               Width           =   1350
            End
            Begin VB.TextBox txtNo 
               Alignment       =   2  '가운데 맞춤
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1545
               Locked          =   -1  'True
               TabIndex        =   18
               Top             =   1770
               Width           =   1350
            End
            Begin VB.TextBox txtColor 
               Alignment       =   2  '가운데 맞춤
               Height          =   360
               Left            =   1545
               Locked          =   -1  'True
               TabIndex        =   17
               Top             =   2160
               Width           =   1350
            End
            Begin VB.TextBox txtTelStore 
               Height          =   360
               Left            =   1545
               Locked          =   -1  'True
               TabIndex        =   16
               Top             =   2550
               Width           =   3345
            End
            Begin VB.TextBox txtTelSMS 
               Height          =   360
               Left            =   1545
               Locked          =   -1  'True
               TabIndex        =   15
               Top             =   2940
               Width           =   3345
            End
            Begin Threed.SSFrame SSFrame2 
               Height          =   3675
               Index           =   1
               Left            =   8925
               TabIndex        =   34
               Top             =   135
               Width           =   4620
               _ExtentX        =   8149
               _ExtentY        =   6482
               _Version        =   262144
               Font3D          =   3
               BackColor       =   16777215
               PictureFrames   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "frm환경설정.frx":2150
               Caption         =   " 택코드 내역"
               Begin FPSpreadADO.fpSpread sprTagHist 
                  Height          =   3120
                  Left            =   195
                  TabIndex        =   37
                  Top             =   345
                  Width           =   4215
                  _Version        =   524288
                  _ExtentX        =   7435
                  _ExtentY        =   5503
                  _StockProps     =   64
                  BorderStyle     =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  GrayAreaBackColor=   16777215
                  MaxCols         =   3
                  ScrollBars      =   2
                  SpreadDesigner  =   "frm환경설정.frx":275A
                  HighlightHeaders=   1
                  HighlightStyle  =   1
               End
            End
            Begin XtremeSuiteControls.PushButton btnServer 
               Height          =   525
               Left            =   12240
               TabIndex        =   118
               Top             =   6960
               Width           =   1305
               _Version        =   851970
               _ExtentX        =   2302
               _ExtentY        =   926
               _StockProps     =   79
               Caption         =   " DB 설정"
               Appearance      =   6
               Picture         =   "frm환경설정.frx":2C6F
            End
            Begin XtremeSuiteControls.PushButton btnBackupFolder 
               Height          =   375
               Left            =   7920
               TabIndex        =   123
               Top             =   6390
               Width           =   1470
               _Version        =   851970
               _ExtentX        =   2593
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "백업폴더 설정"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "2대 이상 컴퓨터 접수여부:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   35
               Left            =   4470
               TabIndex        =   125
               Top             =   7155
               Width           =   2250
            End
            Begin XtremeSuiteControls.CommonDialog CommonDialog1 
               Left            =   1635
               Top             =   5310
               _Version        =   851970
               _ExtentX        =   423
               _ExtentY        =   423
               _StockProps     =   4
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "DB 백업 폴더:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   34
               Left            =   60
               TabIndex        =   122
               Top             =   6465
               Width           =   1425
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "MODE :"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   3
               Left            =   945
               TabIndex        =   120
               Top             =   7155
               Width           =   540
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "비밀번호:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   24
               Left            =   60
               TabIndex        =   73
               Top             =   3405
               Width           =   1425
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "가맹점 코드:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   11
               Left            =   60
               TabIndex        =   28
               Top             =   240
               Width           =   1425
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "가맹점명:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   12
               Left            =   60
               TabIndex        =   27
               Top             =   675
               Width           =   1425
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "적용일자:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   13
               Left            =   60
               TabIndex        =   26
               Top             =   1065
               Width           =   1425
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "지사코드:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   14
               Left            =   60
               TabIndex        =   25
               Top             =   1470
               Width           =   1425
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "택(TAG) 코드:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   15
               Left            =   60
               TabIndex        =   24
               Top             =   1860
               Width           =   1425
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "택(TAG) 색상:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   18
               Left            =   60
               TabIndex        =   23
               Top             =   2235
               Width           =   1425
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "문자발신 전화:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   19
               Left            =   60
               TabIndex        =   22
               Top             =   3015
               Width           =   1425
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "매장 전화번호:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   20
               Left            =   60
               TabIndex        =   21
               Top             =   2625
               Width           =   1425
            End
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   29
         Top             =   15
         Width           =   13725
         _ExtentX        =   24209
         _ExtentY        =   741
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   4194304
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "      환경설정"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm환경설정.frx":3209
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm환경설정.frx":342F
            Top             =   -15
            Width           =   765
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   750
         Left            =   15
         TabIndex        =   69
         Top             =   450
         Width           =   13725
         _ExtentX        =   24209
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdUpdate 
            Height          =   630
            Left            =   45
            TabIndex        =   70
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 저장(&S)"
            Appearance      =   6
            Picture         =   "frm환경설정.frx":3FF9
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   4020
            TabIndex        =   71
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frm환경설정.frx":508B
         End
      End
   End
End
Attribute VB_Name = "frm환경설정"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bchk As Boolean
Dim S_Gu As String
Dim J_Gu As String

Dim KSCAT_GUBUN         As String
Dim KS7500_CommPort  As String
Dim KS7500_BaudRate  As String

Dim SignPad_CommPort As String
Dim SignPad_BaudRate As String

Private Sub btnBackupFolder_Click()
    On Error GoTo ErrRtn

    With CommonDialog1
        '.Filter = "MDB (*.mdb)|*.mdb|전체 (*.*)|*.*"

        .ShowBrowseFolder

        If .FileName <> "" Then
            txtBackupFolder.Text = .FileName & ""
        End If
    End With

    Exit Sub

ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub btnReport_Click()
    On Error GoTo ErrRtn
    
    Dim CommPort As String
    Dim BaudRate As String
    
    Dim tmp      As String
    Dim 이전미수 As String
    Dim 접수수량 As Integer
    Dim 접수금액 As String
    
    Dim 현금결제 As String
    Dim 카드결제 As String
    
    Dim 카드번호 As String
    
    Dim 받은금액 As String
    Dim 거스름돈 As String
    
    Dim ESC      As String
    Dim sE       As String
    Dim PrintMsg As String
    CommPort = cboKS7500CommPort.Text
    BaudRate = cboKS7500BaudRate.Text
    
    Call SetIniStr("VAN", "KSCAT_GUBUN", cboKSCAT.Text, iniFile)     '
    Call SetIniStr("VAN", "KS7500_CommPort", cboKS7500CommPort.Text, iniFile)     '
    Call SetIniStr("VAN", "KS7500_BaudRate", cboKS7500BaudRate.Text, iniFile)     '
    Call SetIniStr("VAN", "SignPad_CommPort", cboSignPadCommPort.Text, iniFile)   '
    Call SetIniStr("VAN", "SignPad_BaudRate", cboSignPadBaudRate.Text, iniFile)   '
    
    Unload frmKicc
   

    PrintMsg = ""
    If 가맹점정보.지사코드 = M_COUPON_KLENZ_CODE Then '크렌즈갤러리
        PrintMsg = PrintMsg & PrintTitle2("크렌즈갤러리 - 세탁물 접수증(테스트)")
    Else
        PrintMsg = PrintMsg & PrintTitle2("크린에이드 - 세탁물 접수증(테스트)")
    End If
    
    PrintMsg = PrintMsg & PrintLineFeed
    
    PrintMsg = PrintMsg & PrintString("상 호 명 : " + txtStoreName.Text, 1, True)
    PrintMsg = PrintMsg & PrintString("전화번호 : " + txtTelStore.Text, 1, True)
    PrintMsg = PrintMsg & PrintString("주    소 : " + txtAddress.Text, 1, True)

    ' 48자 출력              12345678901234567890123456789012345678901234567890
    PrintMsg = PrintMsg & PrintString("==============================================", 1, True)
    PrintMsg = PrintMsg & PrintString("접수일자 : " + Format(Now, "YYYY년 MM월 DD일 AM/PM hh:mm"), 1, True)
    PrintMsg = PrintMsg & PrintString("찾을날짜 : " + Format(Date, "YYYY년 MM월 DD일"), 1, True)
    PrintMsg = PrintMsg & PrintString("고객코드 : " + "000000", 1, True)

   
    PrintMsg = PrintMsg & PrintCustomer("Y", "크린에이드", txtTelStore.Text, txtTelStore.Text, txtStoreName.Text)
    
    
    
    PrintMsg = PrintMsg & PrintString("==============================================", 1, True)
    PrintMsg = PrintMsg & PrintString("택번호  의류/상표         작업   색상     금액", 1, True)
    PrintMsg = PrintMsg & PrintString("----------------------------------------------", 1, True)
    PrintMsg = PrintMsg & PrintString("00-0001 정장상의          세     흰색    1,000", 1, True)
    PrintMsg = PrintMsg & PrintString("00-0002 정장상의          세     흰색    1,000", 1, True)
    PrintMsg = PrintMsg & PrintString("00-0003 정장상의          세     흰색    1,000", 1, True)
    PrintMsg = PrintMsg & PrintString("00-0004 정장상의          세     흰색    1,000", 1, True)
    PrintMsg = PrintMsg & PrintString("00-0005 정장상의          세     흰색    1,000", 1, True)
    PrintMsg = PrintMsg & PrintString("00-0006 정장상의          세     흰색    1,000", 1, True)
    PrintMsg = PrintMsg & PrintString("00-0007 정장상의          세     흰색    1,000", 1, True)
    PrintMsg = PrintMsg & PrintString("00-0008 정장상의          세     흰색    1,000", 1, True)
    PrintMsg = PrintMsg & PrintString("00-0009 정장상의          세     흰색    1,000", 1, True)
    PrintMsg = PrintMsg & PrintString("00-0010 정장상의          세     흰색    1,000", 1, True)
    
    접수수량 = 10000
    접수금액 = "10000"
    이전미수 = "0"
    현금결제 = "5000"
    카드결제 = "5000"

    받은금액 = "10000"
    거스름돈 = "5000"
    
    PrintMsg = PrintMsg & PrintString("----------------------------------------------", 1, True)
    PrintMsg = PrintMsg & PrintString(String(24, " ") + "이전미수 : " + String(9 - LenH(이전미수), " ") + 이전미수 + "원", 1, True)
    PrintMsg = PrintMsg & PrintString("접수수량 : " + String(8 - LenH(CStr(접수수량)), " ") + CStr(접수수량) + "점 / 접수금액 : " + String(9 - LenH(접수금액), " ") + 접수금액 + "원", 1, True)

    PrintMsg = PrintMsg & PrintString(String(24, " ") + "받은금액 : " + String(9 - LenH(받은금액), " ") + 받은금액 + "원", 1, True)
    PrintMsg = PrintMsg & PrintString(String(24, " ") + "거스름돈 : " + String(9 - LenH(거스름돈), " ") + 거스름돈 + "원", 1, True)

    PrintMsg = PrintMsg & PrintString(String(24, " ") + "현금결제 : " + String(9 - LenH(현금결제), " ") + 현금결제 + "원", 1, True)
    PrintMsg = PrintMsg & PrintString(String(24, " ") + "카드결제 : " + String(9 - LenH(카드결제), " ") + 카드결제 + "원", 1, True)
    PrintMsg = PrintMsg & PrintString("==============================================", 1, True)
    
    PrintMsg = PrintMsg & PrintLineFeed

    PrintMsg = PrintMsg & PrintString("※ 인도예정일은 세탁물의 오염정도에 따라 다소", 1, True)
    PrintMsg = PrintMsg & PrintString("   지연될 수 있습니다.", 1, True)
    
    PrintMsg = PrintMsg & PrintLineFeed(4)
    
    PrintMsg = PrintMsg & PrintCut

    Call frmKicc.Card_Print(PrintMsg)
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0

End Sub

Private Sub btnReportFind_Click()
    Dim nPort   As Integer
    
    For nPort = 1 To cboKS7500CommPort.ListCount
    
        Dim sE As String
        Dim sD As String
        Dim Ret As Long
        Rtn = KiccPosOCX.Open(nPort, CLng(cboKS7500BaudRate.Text), sE)   ' 선택한 포트가 연결되어 있는지 체크
        
        sD = "PO"
        Rtn = KiccPosOCX.ReqCmd(&HFD, 0, 0, sD, sE)
        KiccPosOCX.Close
        
        If Rtn = 0 Then
            MsgBox "영수증 프린터 포트는 " & CStr(nPort) & " 입니다."
            Exit Sub
        End If

    Next nPort
    
    MsgBox "영수증 프린터를 찾을수 없습니다."

End Sub

Private Sub btnServer_Click()
    frm서버.Show 1
End Sub

Private Sub btnSignPad_Click()
'    Dim CommPort As String
'    Dim BaudRate As String
'    Dim mesg1       As String
'    Dim mesg2       As String
'    Dim mesg3       As String
'    Dim mesg4       As String
'
'    On Error GoTo ErrRtn
'
'    CommPort = cboSignPadCommPort.Text & ""
'    BaudRate = cboSignPadBaudRate.Text & ""
'
'
'
'    Call KSSignpad.SetComPort(CInt(CommPort), CLng(BaudRate)) '입력한 포트를 설정해준다
'
'    Rtn = KSSignpad.CheckPort '설정한 포트 연결 상태 확인
'
'    If Rtn < 0 Then
'        MsgBox "싸인패드 장치가 연결되어 있지 않습니다.", vbCritical, "확인"
'    Else
'        mesg1 = "싸인패드"                        '메세지는 출력하고자 하는 값을 넣어준다.
'        mesg2 = "정상 연결"
'        mesg3 = "SignPad OK"
'        mesg4 = ""
'
'        Call KSSignpad.SetMinSignPixel(10)                       '최소 픽셀수 설정 : 10개정도로 셋팅한다.(서명을 하지 않고 넘어가는 경우를 방지하기 위한 처리)
'        Call KSSignpad.SetReqSignTimeout(3)                      '전저서명을 싸인패드에 입력 후 입력된 시간 만큼 기다린 후 OnRecvSignData() 이벤트 발생
'
'        Rtn = KSSignpad.SignComReqA1(mesg1, mesg2, mesg3, mesg4) '전자서명 입력 요청
'            MsgBox "싸인패드 장치가 연결되어 있습니다.", vbInformation, "확인"
'
'            KSSignpad.SignComReqA0
'            KSSignpad.ClosePort
'    End If
    Exit Sub
    
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub btnSignPadFind_Click()
'    Dim nPort   As Integer
'
'    Screen.MousePointer = vbHourglass
'
'    For nPort = cboSignPadCommPort.ListCount To 1 Step -1
'
'        Call KSSignpad.SetComPort(CInt(nPort), CLng(cboSignPadBaudRate.Text))  '입력한 포트를 설정해준다
'
'        Rtn = KSSignpad.CheckPort '설정한 포트 연결 상태 확인
'
'        If Rtn > 0 Then
'            Screen.MousePointer = vbDefault
'            MsgBox "사인패드 포트는 " & CStr(Rtn) & " 입니다. [" & Rtn & "]"
'            KSSignpad.ClosePort
'            Exit Sub
'        End If
'
'    Next nPort
'
'    Screen.MousePointer = vbDefault

End Sub

'
Private Sub btnStoreInfoDownload_Click()
    Dim strTemp As String

    On Error GoTo ErrRtn
    
    If Server_Connection(HostCon, "LAUNDRY1000") = False Then Exit Sub
        
    '------------------------------------------------------------------------
    ' TB_기본정보
    '------------------------------------------------------------------------
    Query = "SELECT 지사코드"
    Query = Query & ", 가맹점코드"
    Query = Query & ", 가맹점명"
    Query = Query & ", 대표자명"
    Query = Query & ", 사업자번호"
    Query = Query & ", 업태"
    Query = Query & ", 종목"
    Query = Query & ", 우편번호"
    Query = Query & ", 주소"
    Query = Query & ", 적용일자"
    Query = Query & ", 택코드"
    Query = Query & ", 택색상"
    Query = Query & ", 수선"
    Query = Query & ", 택번호"
    Query = Query & ", 접수번호"
    Query = Query & ", 비율"
    Query = Query & ", ISNULL(요일할인,'0000000') AS 요일할인"
    Query = Query & ", ISNULL(세트상품세일, '0000000') AS 세트상품세일"
    Query = Query & ", 프린터"
    Query = Query & ", 세탁소요일"
    Query = Query & ", 매장전화번호"
    Query = Query & ", 문자발신전화"
    Query = Query & ", 휴대전화번호"
    Query = Query & ", SMS_IP"
    Query = Query & ", SMS_DB"
    Query = Query & ", SMS_ID"
    Query = Query & ", SMS_PWD"
    Query = Query & ", TimeOut"
    Query = Query & ", ISNULL(SMS_EMART,'N')                AS SMS_EMART"
    Query = Query & ", 외주마진"
    
    Query = Query & ", ISNULL(특정할인여부,'N')             AS 특정할인여부"
    Query = Query & ", ISNULL(특정할인비율,'30')            AS 특정할인비율"
    Query = Query & ", ISNULL(특정할인시작일, '2009-01-01') AS 특정할인시작일"
    Query = Query & ", ISNULL(특정할인종료일, '2009-01-01') AS 특정할인종료일"
    
    Query = Query & ", ISNULL(지정할인여부,'N')             AS 지정할인여부"
    Query = Query & ", ISNULL(지정할인비율,'20')            AS 지정할인비율"
    Query = Query & ", ISNULL(지정할인시작일, '2009-01-01') AS 지정할인시작일"
    Query = Query & ", ISNULL(지정할인종료일, '2009-01-01') AS 지정할인종료일"
    
    Query = Query & ", ISNULL(고가세탁비율, '300')          AS 고가세탁비율"
    Query = Query & ", ISNULL(세탁환불여부,'N')             AS 세탁환불여부"
    Query = Query & ", 마일리지여부"
    Query = Query & ", 기준금액"
    Query = Query & ", 적립마일리지"
    Query = Query & ", 최소마일리지"
    Query = Query & ", 가맹점구분"
    Query = Query & ", 담당자코드"
    Query = Query & ", 기사코드"
    Query = Query & ", 계약일자"
    Query = Query & ", 해지일자"
    Query = Query & ", 이전가맹점코드"
    Query = Query & ", 가맹점상태"
    Query = Query & ", 비밀번호"
    
    Query = Query & ", VAN_IP"     '
    Query = Query & ", VAN_PORT"   '
    Query = Query & ", 사업자번호" '
    Query = Query & ", 단말기번호" '
    Query = Query & ", 대표자명"   '
    Query = Query & ", 사업장주소" '
    Query = Query & ", 전산사용료" '
    Query = Query & " FROM TB_가맹점"
    Query = Query & " WHERE 가맹점코드 = '" & txtStoreCode.Text & "'"
    Set SUBRs = New ADODB.RecordSet
    SUBRs.Open Query, HostCon, adOpenForwardOnly, adLockReadOnly
    
    If SUBRs.EOF Then
        SUBRs.Close
        Set SUBRs = Nothing
        
        Query = "가맹점코드 : " & txtStoreCode.Text & vbNewLine & vbNewLine
        Query = Query & "본사로 부터 가맹점 정보를 내려받지 못하였습니다." & vbNewLine
        Query = Query & "본사 담당자에게 문의하세요." & vbNewLine
        
        MsgBox Query, vbCritical, "확인"
        
        Exit Sub
    End If
    
    txtStoreCode.Text = Trim(SUBRs!가맹점코드) & ""                   ' 1
    txtStoreName.Text = Trim(SUBRs!가맹점명) & ""                     ' 2
    txtStartDate.Text = Format(SUBRs!적용일자, "YYYY-MM-DD") & ""     ' 3
    txtMstCode.Text = Trim(SUBRs!지사코드) & ""                       ' 4
    txtNo.Text = Trim(SUBRs!택코드) & ""                              ' 5
    txtColor.Text = Trim(SUBRs!택색상) & ""                           ' 6
     
    If (Trim(SUBRs!요일할인) = "") Or (Len(SUBRs!요일할인) < 7) Then
        chkSale(0).Value = 0
        chkSale(1).Value = 0
        chkSale(2).Value = 0
        chkSale(3).Value = 0
        chkSale(4).Value = 0
        chkSale(5).Value = 0
        chkSale(6).Value = 0
    Else
        chkSale(0).Value = Mid(SUBRs!요일할인, 1, 1)
        chkSale(1).Value = Mid(SUBRs!요일할인, 2, 1)
        chkSale(2).Value = Mid(SUBRs!요일할인, 3, 1)
        chkSale(3).Value = Mid(SUBRs!요일할인, 4, 1)
        chkSale(4).Value = Mid(SUBRs!요일할인, 5, 1)
        chkSale(5).Value = Mid(SUBRs!요일할인, 6, 1)
        chkSale(6).Value = Mid(SUBRs!요일할인, 7, 1)
    End If
    
    If (Trim(SUBRs!세트상품세일) = "") Or (Len(SUBRs!요일할인) < 7) Then
        chkWeek(0).Value = 0
        chkWeek(1).Value = 0
        chkWeek(2).Value = 0
        chkWeek(3).Value = 0
        chkWeek(4).Value = 0
        chkWeek(5).Value = 0
        chkWeek(6).Value = 0
    Else
        chkWeek(0).Value = Mid(SUBRs!세트상품세일, 1, 1)
        chkWeek(1).Value = Mid(SUBRs!세트상품세일, 2, 1)
        chkWeek(2).Value = Mid(SUBRs!세트상품세일, 3, 1)
        chkWeek(3).Value = Mid(SUBRs!세트상품세일, 4, 1)
        chkWeek(4).Value = Mid(SUBRs!세트상품세일, 5, 1)
        chkWeek(5).Value = Mid(SUBRs!세트상품세일, 6, 1)
        chkWeek(6).Value = Mid(SUBRs!세트상품세일, 7, 1)
    End If
    
    txtTelStore.Text = Trim(SUBRs!매장전화번호) & ""                  ' 8
    txtTelSMS.Text = Trim(SUBRs!문자발신전화) & ""                    ' 9
            
    txtMileage(0).Value = SUBRs!기준금액 & ""                         '10
    txtMileage(1).Value = SUBRs!적립마일리지 & ""                     '11
    txtMileage(2).Value = SUBRs!최소마일리지 & ""                     '12
    
    cboSale.ListIndex = IIf(SUBRs!지정할인여부 = "Y", 0, 1)           '
    txtSale.Text = SUBRs!지정할인비율 & ""
    dtpSaleStart.Value = Format(SUBRs!지정할인시작일, "YYYY-MM-DD")   '
    dtpSaleEnd.Value = Format(SUBRs!지정할인종료일, "YYYY-MM-DD")     '
            
    cboCoupon.ListIndex = IIf(SUBRs!특정할인여부 = "Y", 0, 1)         '
    txtCoupon.Text = SUBRs!특정할인비율 & ""                          '
    dtpCouponStart.Value = Format(SUBRs!특정할인시작일, "YYYY-MM-DD") '
    dtpCouponEnd.Value = Format(SUBRs!특정할인종료일, "YYYY-MM-DD")   '
    
    txtLuxury.Text = SUBRs!고가세탁비율 & ""                          '
    
    cboReturn.ListIndex = IIf(SUBRs!세탁환불여부 = "Y", 0, 1)         '
                
    '----------------------------------------------------------------------
    txtPaper.Value = GetIniStr("Printer", "Paper", "", iniFile)       '영수증 출력 장수
    txtPaper2.Value = GetIniStr("Printer", "Paper2", "", iniFile)     '영수증 출력 장수
    
    strTemp = GetIniStr("Printer", "TelPrint", "Y", iniFile)          '전화번호 출력여부
    chkTelPrt.Value = IIf(strTemp = "Y", True, False)
    
    strTemp = GetIniStr("Printer", "CardStorePrint", "Y", iniFile)    ' 카드 가맹점 영수증 출력 여부 Y 출력
    chkStoreCardPrt.Value = IIf(strTemp = "Y", False, True)
    
    '----------------------------------------------------------------------
    
    txtSMSIPAddress.Text = Trim(SUBRs!SMS_IP & "")       '
    txtSMSDBName.Text = Trim(SUBRs!SMS_DB & "")          '
    txtSMSUserName.Text = Trim(SUBRs!SMS_ID & "")        '
    txtSMSUserPass.Text = Trim(SUBRs!SMS_PWD & "")       '
    m_CommandTimeOut = Val(Trim(SUBRs!timeout & ""))     '
    chkSMSEMART.Value = IIf(SUBRs!SMS_EMART = "Y", -1, 0) '
    
    txtPWD.Text = "0" '기본적인 암호는 '0' 이다.
    
    txtVAN(0).Text = Trim(SUBRs!VAN_IP) & ""             '
    txtVAN(1).Text = Trim(SUBRs!VAN_PORT) & ""           '
    txtVAN(2).Text = Trim(SUBRs!사업자번호) & ""         '
    txtVAN(3).Text = Trim(SUBRs!단말기번호) & ""         '
    txtAddress.Text = Trim(SUBRs!사업장주소) & ""        '
    txtChairman.Text = Trim(SUBRs!대표자명) & ""         '
    
    SUBRs.Close
    Set SUBRs = Nothing
    
    Rtn = MsgBox("본사로 부터 가맹점 정보를 내려받았습니다." & vbNewLine & vbNewLine & "적용하시려면 '예' 버튼을 클릭하십시요.", vbQuestion + vbYesNo + vbDefaultButton2, "확인")
    
    If Rtn = vbYes Then
        ADOCon.Execute "DELETE FROM TB_기본정보"

        cmdUpdate_Click
        
        MsgBox "가맹점 정보를 저장 하였습니다." & vbNewLine & vbNewLine & "프로그램을 종료합니다.", vbInformation, "확인"
        
        End
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
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

 

Private Sub cboKSCAT_Click()
    If cboKSCAT.Text = "KS4060 보안인증" Then
        SSFrame_KS4060.Visible = True
        Label2(27).Visible = True
        Label2(28).Visible = True
        cboSignPadCommPort.Visible = True
        cboSignPadBaudRate.Visible = True
        btnSignPadFind.Visible = True
        btnSignPad.Visible = True
    ElseIf cboKSCAT.Text = "KICC" Then
        Label2(27).Visible = False
        Label2(28).Visible = False
        cboSignPadCommPort.Visible = False
        cboSignPadBaudRate.Visible = False
        btnSignPadFind.Visible = False
        btnSignPad.Visible = False
        btnReportFind.Visible = True
    ElseIf cboKSCAT.Text = "KICC Comm 모듈" Then
        Label2(27).Visible = False
        Label2(28).Visible = False
        cboSignPadCommPort.Visible = False
        cboSignPadBaudRate.Visible = False
        btnSignPadFind.Visible = False
        btnSignPad.Visible = False
        btnReportFind.Visible = False
    Else
        SSFrame_KS4060.Visible = False
        Label2(27).Visible = True
        Label2(28).Visible = True
        cboSignPadCommPort.Visible = True
        cboSignPadBaudRate.Visible = True
        btnSignPadFind.Visible = True
        btnSignPad.Visible = True
    End If
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 5: Unload Me
    End Select
End Sub

Private Sub cmdSMSTest_Click()
    On Error GoTo ErrRtn
    
    Dim HostConn As String
    
    HostConn = "Provider=SQLOLEDB.1;Persist Security Info=False;"
    HostConn = HostConn & "User ID=" & txtSMSUserName.Text & ";"
    HostConn = HostConn & "Password=" & txtSMSUserPass.Text & ";"
    HostConn = HostConn & "Initial Catalog=" & txtSMSDBName.Text & ";"
    HostConn = HostConn & "Data Source=" & txtSMSIPAddress.Text
    
    Set HostCon = Nothing
    Set HostCon = New ADODB.Connection

    If HostCon.State = adStateOpen Then HostCon.Close
    
    HostCon.ConnectionTimeout = 10
    HostCon.CommandTimeout = 30
    HostCon.Open HostConn

    MsgBox "문자서비스(SMS) 발송 가능합니다.", vbInformation, "확인"
    
    HostCon.Close
    Set HostCon = Nothing
    
    Exit Sub
    
ErrRtn:
    MsgBox "문자서비스(SMS) 발송이 불가능합니다.", vbCritical, "확인"
End Sub

Private Sub Form_Activate()
    Dim sFullFileName   As String
    Dim sVer            As String
    
    On Error GoTo ERR_RTN
    
    sFullFileName = "C:\windows\system32\KiccPosIEX.ocx"
    If Dir(sFullFileName, vbNormal) <> "" Then
        sVer = CheckKSNET_OCX_Ver(sFullFileName)
        Label_KSNET.Caption = sFullFileName & vbNewLine & "Ver:" & Left(sVer, 10)
        Exit Sub
    End If
    
    sFullFileName = "C:\windows\SysWOW64\KiccPosIEX.ocx"
    If Dir(sFullFileName, vbNormal) <> "" Then
        sVer = CheckKSNET_OCX_Ver(sFullFileName)
        Label_KSNET.Caption = sFullFileName & vbNewLine & "Ver:" & Left(sVer, 10)
        Exit Sub
    End If
    
    sFullFileName = "C:\cleanaid\KiccPosIEX.ocx"
    If Dir(sFullFileName, vbNormal) <> "" Then
        sVer = CheckKSNET_OCX_Ver(sFullFileName)
        Label_KSNET.Caption = sFullFileName & vbNewLine & "Ver:" & Left(sVer, 10)
    End If
    Exit Sub
    
ERR_RTN:
    MsgBox sFullFileName & "->" & Err.description
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    Me.Top = ((frmMain.Height - Me.Height) / 2) + frmMain.Top
    Me.Left = ((frmMain.Width - Me.Width) / 2) + frmMain.Left
  
    With sprTagHist
        .MaxRows = 0
        .RowHeight(-1) = 14
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeRow
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
    
    With sprMargin
        .MaxRows = 0
        .RowHeight(-1) = 14
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeRow
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
    
    TabControl1.SelectedItem = 0
    
   '-----------------------------------------------------------
    Dim sComputer As String
    
    With cboComputer
        .Clear
        
        .AddItem "아니오"
        .AddItem "예"
        
        .ListIndex = 0
    End With
    
    sComputer = GetIniStr("DB", "DualComputer", "N", iniFile) '
    
    If sComputer = "N" Then
        cboComputer.Text = "아니오"
    Else
        cboComputer.Text = "예"
    End If
    
    '-----------------------------------------------------------
    With cboMileage
        .Clear
        
        .AddItem "사용"
        .AddItem "미사용"
        
        .ListIndex = 0
    End With
        
    '--------------------------------
    ' 단말기 종료
    '--------------------------------
    With cboKSCAT
        .Clear
        .AddItem "KS7500i, KS7050i"
        .AddItem "KS4060 보안인증"
        .AddItem "KICC"
        .AddItem "KICC Comm 모듈"
    End With
    '--------------------------------
    ' KS7500
    '--------------------------------
    With cboKS7500CommPort
        .Clear
        
        For i = 1 To 10
            .AddItem i
        Next i
    End With
        
    With cboKS7500BaudRate
        .Clear
        .AddItem "9600"
        .AddItem "19200"
        .AddItem "38400"
        .AddItem "57600"
        .AddItem "115200"
    End With
        
    '--------------------------------
    ' SignPad
    '--------------------------------
    With cboSignPadCommPort
        .Clear
        
        For i = 1 To 10
            .AddItem i
        Next i
    End With
        
    With cboSignPadBaudRate
        .Clear
        .AddItem "9600"
        .AddItem "19200"
        .AddItem "38400"
        .AddItem "57600"
    End With
        
   '---------------------------------------------------------------------
    Dim sMode As String
    
    With cboMode
        .Clear
        .AddItem "REAL"
        .AddItem "TEST"
    End With
        
    sMode = GetIniStr("RUNMODE", "MODE", "", iniFile) '
    
    If sMode = "REAL" Then
        cboMode.Text = "REAL"
    Else
        cboMode.Text = "TEST"
    End If
    '---------------------------------------------------------------------
    KSCAT_GUBUN = GetIniStr("VAN", "KSCAT_GUBUN", "", iniFile) '

    KS7500_CommPort = GetIniStr("VAN", "KS7500_CommPort", "", iniFile) '
    KS7500_BaudRate = GetIniStr("VAN", "KS7500_BaudRate", "", iniFile) '
    
    SignPad_CommPort = GetIniStr("VAN", "SignPad_CommPort", "", iniFile) '
    SignPad_BaudRate = GetIniStr("VAN", "SignPad_BaudRate", "", iniFile) '
        
    If KSCAT_GUBUN <> "" Then
        cboKSCAT.Text = KSCAT_GUBUN
    Else
        cboKSCAT.ListIndex = 0
    End If
        
    If KS7500_CommPort <> "" Then
        cboKS7500CommPort.Text = KS7500_CommPort
    End If
        
    If KS7500_BaudRate <> "" Then
        cboKS7500BaudRate.Text = KS7500_BaudRate
    End If
        
    If SignPad_CommPort <> "" Then
        cboSignPadCommPort.Text = SignPad_CommPort
    End If
        
    If SignPad_BaudRate <> "" Then
        cboSignPadBaudRate.Text = SignPad_BaudRate
    End If
        
    Call 기본정보_Display
    Call 마진율_Display
    
    '[BACKUP]
    txtBackupFolder.Text = GetIniStr("BACKUP", "PATH", "", iniFile)  '
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub 마진율_Display()
    On Error GoTo ErrRtn
    
    Query = "SELECT * FROM TB_의류분류"
    Query = Query & " ORDER BY 의류분류코드 ASC"
    Set SUBRs = New ADODB.RecordSet
    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    With sprMargin
        .MaxRows = 0
        .ReDraw = False
                
        Do Until SUBRs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = SUBRs!의류분류코드 & ""
            .Col = 2: .Text = SUBRs!의류분류명 & ""
            .Col = 3: .Text = SUBRs!세탁마진 & ""
            .Col = 4: .Text = SUBRs!외주마진 & ""
            .Col = 5: .Text = SUBRs!수선마진 & ""
            
            SUBRs.MoveNext
        Loop
        SUBRs.Close
        Set SUBRs = Nothing
        
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub 기본정보_Display()
    On Error GoTo ErrRtn
    
    Dim strTemp As String
    
    Query = "SELECT    지사코드"
    Query = Query & ", 가맹점코드"
    Query = Query & ", 가맹점명"
    Query = Query & ", 가맹점구분"
    Query = Query & ", 적용일자"
    Query = Query & ", 택코드"
    Query = Query & ", 택색상"
    Query = Query & ", 택번호"
    Query = Query & ", 접수번호"
    Query = Query & ", ISNULL(요일할인, '0000000') AS 요일할인"
    Query = Query & ", ISNULL(세트상품세일,'0000000') AS 세트상품세일"
    Query = Query & ", 세탁소요일"
    Query = Query & ", SMS_IP"
    Query = Query & ", SMS_DB"
    Query = Query & ", SMS_ID"
    Query = Query & ", SMS_PWD"
    Query = Query & ", ISNULL(TimeOut,'30') AS TIMEOUT"
    Query = Query & ", 프로그램버전"
    Query = Query & ", 매장전화번호"
    Query = Query & ", 문자발신전화"
    Query = Query & ", SMS_EMART"
    
    'Query = Query & ", 수선"
    'Query = Query & ", 일수"
    'Query = Query & ", ISNULL(비율,30) AS 비율"
    'Query = Query & ", 전화1"
    'Query = Query & ", 전화2"
    'Query = Query & ", ISNULL(수선마진,30) AS 수선마진"
    'Query = Query & ", 프린터"
    'Query = Query & ", ISNULL(운동화마진,40) AS 운동화마진"
    'Query = Query & ", ISNULL(가죽무스탕마진,40) AS 가죽무스탕마진"
    'Query = Query & ", ISNULL(카페트마진,40) AS 카페트마진"
    'Query = Query & ", 보관증종류"
    'Query = Query & ", 마일리지검사일자"
    'Query = Query & ", 마일리지증가구분"
    
    Query = Query & ", 마일리지여부"
    Query = Query & ", ISNULL(고가세탁비율, 300) AS 고가세탁비율"
    Query = Query & ", ISNULL(외주마진,0) AS 외주마진"
    Query = Query & ", 세탁환불여부"
    
    Query = Query & ", 특정할인여부"
    Query = Query & ", ISNULL(특정할인비율, 30) AS 특정할인비율"
    Query = Query & ", ISNULL(특정할인시작일, '2009-01-01') AS 특정할인시작일"
    Query = Query & ", ISNULL(특정할인종료일, '2009-01-01') AS 특정할인종료일"
    
    Query = Query & ", 쿠폰할인여부"
    Query = Query & ", 쿠폰할인비율"
    Query = Query & ", 쿠폰할인시작일"
    Query = Query & ", 쿠폰할인종료일"
    
    Query = Query & ", 지정할인여부"
    Query = Query & ", ISNULL(지정할인비율, 20) AS 지정할인비율"
    Query = Query & ", ISNULL(지정할인시작일, '2009-01-01') AS 지정할인시작일"
    Query = Query & ", ISNULL(지정할인종료일, '2009-01-01') AS 지정할인종료일"
    
    Query = Query & ", 비밀번호"
    
    Query = Query & ", 기준금액"
    Query = Query & ", 적립마일리지"
    Query = Query & ", 최소마일리지"
    Query = Query & ", VAN_IP"
    Query = Query & ", VAN_PORT"
    Query = Query & ", 사업자번호"
    Query = Query & ", 단말기번호"
    Query = Query & ", 사업장주소"
    Query = Query & ", 대표자명"
    
    Query = Query & ", 로열티여부1"
    Query = Query & ", 로열티비율1"
    Query = Query & ", 로열티여부2"
    Query = Query & ", 로열티비율2"
    Query = Query & ", 수수료지원여부"
    Query = Query & ", 수수료지원비율"
    
    
    Query = Query & " FROM TB_기본정보"
    Set SUBRs = New ADODB.RecordSet
    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If Not SUBRs.EOF Then
        txtStoreCode.Text = Trim(SUBRs!가맹점코드) & ""                   ' 1
        txtStoreName.Text = Trim(SUBRs!가맹점명) & ""                     ' 2
        txtStartDate.Text = Format(SUBRs!적용일자, "YYYY-MM-DD") & ""     ' 3
        txtMstCode.Text = Trim(SUBRs!지사코드) & ""                       ' 4
        txtNo.Text = Trim(SUBRs!택코드) & ""                              ' 5
        txtColor.Text = Trim(SUBRs!택색상) & ""                           ' 6
        txtTelStore.Text = Trim(SUBRs!매장전화번호) & ""                  ' 7
        txtTelSMS.Text = Trim(SUBRs!문자발신전화) & ""                    ' 8
        
        txtPWD.Text = SUBRs!비밀번호 & ""                                 '
        
        '================================================================
        
        If SUBRs!마일리지여부 = "Y" Then
            cboMileage.Text = "사용"                                      ' 9
        Else
            cboMileage.Text = "미사용"                                    ' 9
        End If
        
        txtMileage(0).Value = SUBRs!기준금액 & ""                         '10
        txtMileage(1).Value = SUBRs!적립마일리지 & ""                     '11
        txtMileage(2).Value = SUBRs!최소마일리지 & ""                     '12
        
        '--------------------------------------------------------------------
        If IsNull(SUBRs!지정할인여부) Then
            cboSale.ListIndex = 1                                         '13
        Else
            cboSale.ListIndex = IIf(SUBRs!지정할인여부 = "Y", 0, 1)       '13
        End If
        
        txtSale.Text = SUBRs!지정할인비율 & ""                            '14
        dtpSaleStart.Value = Format(SUBRs!지정할인시작일, "YYYY-MM-DD")   '15
        dtpSaleEnd.Value = Format(SUBRs!지정할인종료일, "YYYY-MM-DD")     '16
        
        '--------------------------------------------------------------------
        
        If IsNull(SUBRs!특정할인여부) Then
            cboCoupon.ListIndex = 1                                       '17
        Else
            cboCoupon.ListIndex = IIf(SUBRs!특정할인여부 = "Y", 0, 1)     '17
        End If
        
        txtCoupon.Text = SUBRs!특정할인비율 & ""                          '18
        dtpCouponStart.Value = Format(SUBRs!특정할인시작일, "YYYY-MM-DD") '19
        dtpCouponEnd.Value = Format(SUBRs!특정할인종료일, "YYYY-MM-DD")   '20
        
        txtLuxury.Text = SUBRs!고가세탁비율 & ""                          '21
        
        If IsNull(SUBRs!세탁환불여부) Then
            cboReturn.ListIndex = 1                                       '22
        Else
            cboReturn.ListIndex = IIf(SUBRs!세탁환불여부 = "Y", 0, 1)   '22
        End If
                        
        If (Trim(SUBRs!요일할인) = "") Or (Len(Trim(SUBRs!요일할인)) < 7) Then
            chkSale(0).Value = 0
            chkSale(1).Value = 0
            chkSale(2).Value = 0
            chkSale(3).Value = 0
            chkSale(4).Value = 0
            chkSale(5).Value = 0
            chkSale(6).Value = 0
        Else
            chkSale(0).Value = Mid(SUBRs!요일할인, 1, 1)
            chkSale(1).Value = Mid(SUBRs!요일할인, 2, 1)
            chkSale(2).Value = Mid(SUBRs!요일할인, 3, 1)
            chkSale(3).Value = Mid(SUBRs!요일할인, 4, 1)
            chkSale(4).Value = Mid(SUBRs!요일할인, 5, 1)
            chkSale(5).Value = Mid(SUBRs!요일할인, 6, 1)
            chkSale(6).Value = Mid(SUBRs!요일할인, 7, 1)
        End If
        
        If (Trim(SUBRs!세트상품세일) = "") Or (Len(Trim(SUBRs!세트상품세일)) < 7) Then
            chkWeek(0).Value = 0
            chkWeek(1).Value = 0
            chkWeek(2).Value = 0
            chkWeek(3).Value = 0
            chkWeek(4).Value = 0
            chkWeek(5).Value = 0
            chkWeek(6).Value = 0
        Else
            chkWeek(0).Value = Mid(SUBRs!세트상품세일, 1, 1)
            chkWeek(1).Value = Mid(SUBRs!세트상품세일, 2, 1)
            chkWeek(2).Value = Mid(SUBRs!세트상품세일, 3, 1)
            chkWeek(3).Value = Mid(SUBRs!세트상품세일, 4, 1)
            chkWeek(4).Value = Mid(SUBRs!세트상품세일, 5, 1)
            chkWeek(5).Value = Mid(SUBRs!세트상품세일, 6, 1)
            chkWeek(6).Value = Mid(SUBRs!세트상품세일, 7, 1)
        End If
        
        '----------------------------------------------------------------------
                
        txtPaper.Value = GetIniStr("Printer", "Paper", "", iniFile) '영수증 출력 장수
        txtPaper2.Value = GetIniStr("Printer", "Paper2", "", iniFile) '영수증 출력 장수
        
        strTemp = GetIniStr("Printer", "TelPrint", "Y", iniFile)    '전화번호 출력여부
        chkTelPrt.Value = IIf(strTemp = "Y", True, False)
        
        strTemp = GetIniStr("Printer", "CardStorePrint", "Y", iniFile) ' 가맹점 카드 영수증 출력 여부 Y 출력
        chkStoreCardPrt.Value = IIf(strTemp = "Y", False, True)
        
        
        '----------------------------------------------------------------------
                
        txtSMSIPAddress.Text = Trim(SUBRs!SMS_IP) & ""    '25
        txtSMSDBName.Text = Trim(SUBRs!SMS_DB) & ""       '26
        txtSMSUserName.Text = Trim(SUBRs!SMS_ID) & ""   '27
        txtSMSUserPass.Text = Trim(SUBRs!SMS_PWD) & ""   '28
        m_CommandTimeOut = Val(Trim(SUBRs!timeout) & "")    '29
        
        If IsNull(SUBRs.Fields("SMS_EMART")) = True Then
            chkSMSEMART.Value = -1                                          '30
        Else
            chkSMSEMART.Value = IIf(SUBRs.Fields("SMS_EMART") = "Y", -1, 0) '30
        End If
        
        txtVAN(0).Text = SUBRs!VAN_IP & ""
        txtVAN(1).Text = SUBRs!VAN_PORT & ""
        txtVAN(2).Text = SUBRs!사업자번호 & ""
        txtVAN(3).Text = SUBRs!단말기번호 & ""
        
        txtChairman.Text = SUBRs!대표자명 & ""
        txtAddress.Text = SUBRs!사업장주소 & ""
        
        
        
        '--------------------------------------------------------------------
        If IsNull(SUBRs!로열티여부1) Then
            cboRovalty(0).ListIndex = 1
        Else
            cboRovalty(0).ListIndex = IIf(SUBRs!로열티여부1 = "Y", 0, 1)
        End If
        
        txtRovalty(0).Text = SUBRs!로열티비율1 & ""
        
        '--------------------------------------------------------------------
        If IsNull(SUBRs!로열티여부2) Then
            cboRovalty(1).ListIndex = 1
        Else
            cboRovalty(1).ListIndex = IIf(SUBRs!로열티여부2 = "Y", 0, 1)
        End If
        
        txtRovalty(1).Text = SUBRs!로열티비율2 & ""
        
        '--------------------------------------------------------------------
        If IsNull(SUBRs!수수료지원여부) Then
            cboRovalty(2).ListIndex = 1
        Else
            cboRovalty(2).ListIndex = IIf(SUBRs!수수료지원여부 = "Y", 0, 1)
        End If
        
        txtRovalty(2).Text = SUBRs!수수료지원비율 & ""
    
    End If
    SUBRs.Close
    Set SUBRs = Nothing
        
    ' 변경 내용을 처리하기 위하여..
    txtMstCode.Tag = txtMstCode.Text
    txtStoreCode.Tag = txtStoreCode.Text
    txtStoreName.Tag = txtStoreName.Text
    txtStartDate.Tag = txtStartDate.Text
    txtNo.Tag = txtNo.Text
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    cmdBtn(5).Left = Me.Width - cmdBtn(5).Width - 200
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
            txtStartDate.Text = Format(Date, "YYYY-MM-DD")
        
            
        Case Else
            KeyAscii = 0
            Exit Sub
    End Select

End Sub

Private Sub cmdUpdate_Click()
    On Error GoTo ErrRtn
    
    
    Dim strCheck   As String
        
    If Trim(txtPWD.Text) = "" Then
        txtPWD.Text = "0" '초기암호 '0'
    End If
    
    '-----------------------------------------------------------------------------------
    ' TB_기본정보
    '-----------------------------------------------------------------------------------
    Query = "SELECT * FROM TB_기본정보"
    Query = Query & " WHERE 가맹점코드 = '" & txtStoreCode.Text & "'"
    Set SUBRs = New ADODB.RecordSet
    SUBRs.Open Query, ADOCon, adOpenDynamic, adLockOptimistic
    
    If SUBRs.EOF Then SUBRs.AddNew

    SUBRs!지사코드 = Trim(txtMstCode.Text) & ""                                  ' 1
    SUBRs!가맹점코드 = txtStoreCode.Text & ""                                    ' 2
    SUBRs!가맹점명 = txtStoreName.Text & ""                                      ' 3
    SUBRs!적용일자 = Format(txtStartDate.Text, "YYYY-MM-DD") & ""                ' 4
    SUBRs!택코드 = txtNo.Text & ""                                               ' 5
    SUBRs!택색상 = txtColor.Text & ""                                            ' 6
    SUBRs!매장전화번호 = txtTelStore.Text & ""                                   ' 7
    SUBRs!문자발신전화 = txtTelSMS.Text & ""                                     ' 8
    
    '------------------------------------------------------------------------------------
    ' 요일할인
    '------------------------------------------------------------------------------------
    If chkSale(0).Value = 0 Then strCheck = "0" Else strCheck = "1"
    If chkSale(1).Value = 0 Then strCheck = strCheck & "0" Else strCheck = strCheck & "1"
    If chkSale(2).Value = 0 Then strCheck = strCheck & "0" Else strCheck = strCheck & "1"
    If chkSale(3).Value = 0 Then strCheck = strCheck & "0" Else strCheck = strCheck & "1"
    If chkSale(4).Value = 0 Then strCheck = strCheck & "0" Else strCheck = strCheck & "1"
    If chkSale(5).Value = 0 Then strCheck = strCheck & "0" Else strCheck = strCheck & "1"
    If chkSale(6).Value = 0 Then strCheck = strCheck & "0" Else strCheck = strCheck & "1"
    
    SUBRs!요일할인 = strCheck & ""                                               ' 9
    
    '------------------------------------------------------------------------------------
    ' 세트상품세일
    '------------------------------------------------------------------------------------
    If chkWeek(0).Value = 0 Then strCheck = "0" Else strCheck = "1"
    If chkWeek(1).Value = 0 Then strCheck = strCheck & "0" Else strCheck = strCheck & "1"
    If chkWeek(2).Value = 0 Then strCheck = strCheck & "0" Else strCheck = strCheck & "1"
    If chkWeek(3).Value = 0 Then strCheck = strCheck & "0" Else strCheck = strCheck & "1"
    If chkWeek(4).Value = 0 Then strCheck = strCheck & "0" Else strCheck = strCheck & "1"
    If chkWeek(5).Value = 0 Then strCheck = strCheck & "0" Else strCheck = strCheck & "1"
    If chkWeek(6).Value = 0 Then strCheck = strCheck & "0" Else strCheck = strCheck & "1"
    
    SUBRs!세트상품세일 = strCheck & ""                                           '10
    
    SUBRs!마일리지여부 = IIf(Trim(cboMileage.Text) = "사용", "Y", "N") & ""      '11
    SUBRs!기준금액 = txtMileage(0).Value                                         '12
    SUBRs!적립마일리지 = txtMileage(1).Value                                     '13
    SUBRs!최소마일리지 = txtMileage(2).Value                                     '14
    
    SUBRs!지정할인여부 = IIf(Trim(cboSale.Text) = "예", "Y", "N") & ""           '15
    SUBRs!지정할인비율 = txtSale.Text & ""                                       '16
    SUBRs!지정할인시작일 = Format(dtpSaleStart.Value, "YYYY-MM-DD") & ""         '17
    SUBRs!지정할인종료일 = Format(dtpSaleEnd.Value, "YYYY-MM-DD") & ""           '18
    
    SUBRs!특정할인여부 = IIf(Trim(cboCoupon.Text) = "예", "Y", "N") & ""         '19
    SUBRs!특정할인비율 = txtCoupon.Text & ""                                     '20
    SUBRs!특정할인시작일 = Format(dtpCouponStart.Value, "YYYY-MM-DD") & ""       '21
    SUBRs!특정할인종료일 = Format(dtpCouponEnd.Value, "YYYY-MM-DD") & ""         '22

'    SUBRS!   쿠폰할인여부   =  IIf(Trim(cboCoupon.Text) = "예", "Y", "N") & ""
'    SUBRS!   쿠폰할인비율     =  txtCoupon.Text & ""
'    SUBRS!   쿠폰할인시작일     =  Format(dtpCouponStart.Value, "YYYY-MM-DD") & ""
'    SUBRS!   쿠폰할인종료일     =  Format(dtpCouponEnd.Value, "YYYY-MM-DD") & ""

    SUBRs!고가세탁비율 = txtLuxury.Text & ""                                     '23
    SUBRs!세탁환불여부 = IIf(Trim(cboReturn.Text) = "예", "Y", "N") & ""         '24

    SUBRs!SMS_IP = txtSMSIPAddress.Text & ""                                     '25
    SUBRs!SMS_DB = txtSMSDBName.Text & ""                                        '26
    SUBRs!SMS_ID = txtSMSUserName.Text & ""                                      '27
    SUBRs!SMS_PWD = txtSMSUserPass.Text & ""                                     '28

    'SUBRS!보관증종류       =  Printer_BO_Gb & ""                                '
    'SUBRs!프린터 = Bill_Printer & ""                                            '

    SUBRs!SMS_EMART = IIf(chkSMSEMART.Value = -1, "Y", "N") & ""                 '29
    SUBRs!VAN_IP = txtVAN(0).Text & ""                                           '
    SUBRs!VAN_PORT = txtVAN(1).Text & ""                                         '
    SUBRs!사업자번호 = txtVAN(2).Text & ""                                       '
    SUBRs!단말기번호 = txtVAN(3).Text & ""                                       '
    SUBRs!비밀번호 = Trim(txtPWD.Text) & ""                                      '
    
    SUBRs!대표자명 = txtChairman.Text & ""                                       '
    SUBRs!사업장주소 = txtAddress.Text & ""                                      '
    
    SUBRs!CAT단말기종류 = cboKSCAT.Text                                   '
    
    SUBRs!본사전송여부 = "N"                                                     ' 본사전송여부
    SUBRs!본사전송일자 = ""                                                      ' 본사전송일자
    
    SUBRs.Update
    
    SUBRs.Close
    Set SUBRs = Nothing
    
    
    '--------------------------------------------------------------------------------------
    '
    '--------------------------------------------------------------------------------------
    Call SetIniStr("VAN", "KSCAT_GUBUN", cboKSCAT.Text, iniFile)     '
    Call SetIniStr("VAN", "KS7500_CommPort", cboKS7500CommPort.Text, iniFile)     '
    Call SetIniStr("VAN", "KS7500_BaudRate", cboKS7500BaudRate.Text, iniFile)     '
    Call SetIniStr("VAN", "SignPad_CommPort", cboSignPadCommPort.Text, iniFile)   '
    Call SetIniStr("VAN", "SignPad_BaudRate", cboSignPadBaudRate.Text, iniFile)   '
    
    Call SetIniStr("Printer", "Paper", txtPaper.Value, iniFile)                   '
    Call SetIniStr("Printer", "Paper2", txtPaper2.Value, iniFile)                 '
    
    If chkTelPrt.Value = True Then
        Call SetIniStr("Printer", "TelPrint", "Y", iniFile)                       '
    Else
        Call SetIniStr("Printer", "TelPrint", "N", iniFile)                       '
    End If

    ' 출력 안함
    If chkStoreCardPrt.Value = True Then
        ' 선택되면 출력을 안한다.
        Call SetIniStr("Printer", "CardStorePrint", "N", iniFile)                       '
    Else
        Call SetIniStr("Printer", "CardStorePrint", "Y", iniFile)                       '
    End If

    Call SetIniStr("RUNMODE", "MODE", cboMode.Text, iniFile)                      'REAL or TEST

    Call SetIniStr("BACKUP", "PATH", txtBackupFolder.Text, iniFile)               'BACKUP

    '접수 컴퓨터
    If cboComputer.Text = "아니오" Then
        Call SetIniStr("DB", "DualComputer", "N", iniFile)
    Else
        Call SetIniStr("DB", "DualComputer", "Y", iniFile)
    End If
    
    MsgBox "프로그램을 다시 시작하십시요.", vbCritical, "확인"
    
    ADOCon.Close
    Set ADOCon = Nothing
    
    End
    
    Exit Sub

ErrRtn:
    Resume Next
End Sub
 
