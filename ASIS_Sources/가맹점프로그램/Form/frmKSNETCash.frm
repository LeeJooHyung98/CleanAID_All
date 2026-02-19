VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frmKSNETCash 
   BorderStyle     =   1  '단일 고정
   Caption         =   "현금영수증"
   ClientHeight    =   7515
   ClientLeft      =   5715
   ClientTop       =   4590
   ClientWidth     =   5025
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   5025
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7515
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   13256
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frmKSNETCash.frx":0000
      Begin XtremeSuiteControls.TabControl TabControl 
         Height          =   5910
         Left            =   15
         TabIndex        =   5
         Top             =   960
         Width           =   4995
         _Version        =   851970
         _ExtentX        =   8811
         _ExtentY        =   10425
         _StockProps     =   68
         Appearance      =   2
         Color           =   16
         ItemCount       =   2
         Item(0).Caption =   "현금영수증"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "TabControlPage1"
         Item(1).Caption =   "승인/취소 정보"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "TabControlPage2"
         Begin XtremeSuiteControls.TabControlPage TabControlPage2 
            Height          =   5565
            Left            =   -69970
            TabIndex        =   7
            Top             =   315
            Visible         =   0   'False
            Width           =   4935
            _Version        =   851970
            _ExtentX        =   8705
            _ExtentY        =   9816
            _StockProps     =   1
            Page            =   1
            Begin Threed.SSPanel pnlNum 
               Height          =   315
               Left            =   975
               TabIndex        =   8
               Top             =   435
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   556
               _Version        =   262144
               Font3D          =   1
               BackColor       =   16777215
               Caption         =   "0"
               BorderWidth     =   0
               BevelOuter      =   0
               BevelInner      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCustomCode 
               Height          =   315
               Left            =   975
               TabIndex        =   9
               Top             =   90
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   556
               _Version        =   262144
               Font3D          =   1
               ForeColor       =   8421504
               BackColor       =   16777215
               BorderWidth     =   0
               BevelOuter      =   0
               BevelInner      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlApprovalNo 
               Height          =   315
               Left            =   3630
               TabIndex        =   10
               Top             =   105
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   556
               _Version        =   262144
               Font3D          =   1
               BackColor       =   16777215
               BorderWidth     =   0
               BevelOuter      =   0
               BevelInner      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlApprovalDay 
               Height          =   315
               Left            =   3630
               TabIndex        =   11
               Top             =   435
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   556
               _Version        =   262144
               Font3D          =   1
               BackColor       =   16777215
               BorderWidth     =   0
               BevelOuter      =   0
               BevelInner      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin FPSpreadADO.fpSpread sprGrid 
               Height          =   4005
               Left            =   105
               TabIndex        =   16
               Top             =   1290
               Width           =   4740
               _Version        =   524288
               _ExtentX        =   8361
               _ExtentY        =   7064
               _StockProps     =   64
               BackColorStyle  =   1
               DisplayColHeaders=   0   'False
               EditModePermanent=   -1  'True
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
               GridSolid       =   0   'False
               MaxCols         =   1
               MaxRows         =   12
               RowHeaderDisplay=   0
               ScrollBars      =   0
               SpreadDesigner  =   "frmKSNETCash.frx":0072
               HighlightHeaders=   1
               HighlightStyle  =   1
            End
            Begin Threed.SSPanel pnlApprovalTime 
               Height          =   315
               Left            =   3630
               TabIndex        =   30
               Top             =   780
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   556
               _Version        =   262144
               Font3D          =   1
               BackColor       =   16777215
               BorderWidth     =   0
               BevelOuter      =   0
               BevelInner      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin VB.Label Label1 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "승인시간:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   1
               Left            =   2520
               TabIndex        =   31
               Top             =   855
               Width           =   1080
            End
            Begin VB.Label Label1 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "고객코드:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   5
               Left            =   60
               TabIndex        =   15
               Top             =   150
               Width           =   885
            End
            Begin VB.Label Label1 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "접수번호:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   6
               Left            =   60
               TabIndex        =   14
               Top             =   510
               Width           =   885
            End
            Begin VB.Label Label1 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "승인번호:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   7
               Left            =   2520
               TabIndex        =   13
               Top             =   165
               Width           =   1080
            End
            Begin VB.Label Label1 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "승인일자:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   8
               Left            =   2520
               TabIndex        =   12
               Top             =   510
               Width           =   1080
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage1 
            Height          =   5565
            Left            =   30
            TabIndex        =   6
            Top             =   315
            Width           =   4935
            _Version        =   851970
            _ExtentX        =   8705
            _ExtentY        =   9816
            _StockProps     =   1
            Page            =   0
            Begin VB.ComboBox cboCancel 
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
               Left            =   1200
               Style           =   2  '드롭다운 목록
               TabIndex        =   34
               Top             =   540
               Width           =   2760
            End
            Begin VB.ComboBox cboGubun 
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
               Left            =   1200
               Style           =   2  '드롭다운 목록
               TabIndex        =   18
               Top             =   150
               Width           =   2760
            End
            Begin VB.TextBox txtUserInfo 
               BackColor       =   &H0080FFFF&
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   14.25
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   660
               IMEMode         =   3  '사용 못함
               Left            =   1200
               MultiLine       =   -1  'True
               TabIndex        =   17
               Top             =   2385
               Width           =   3600
            End
            Begin CSTextLibCtl.silgEdit txtMoney 
               Height          =   405
               Left            =   1200
               TabIndex        =   19
               Top             =   1920
               Width           =   3600
               _Version        =   262145
               _ExtentX        =   6350
               _ExtentY        =   714
               _StockProps     =   125
               Text            =   " 0"
               ForeColor       =   0
               BackColor       =   12648447
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.25
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
               RawData         =   "0"
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
               Justification   =   1
               BorderStyle     =   0
               Undo            =   1
               Data            =   0
            End
            Begin Threed.SSPanel SSPanel1 
               Height          =   1785
               Left            =   1200
               TabIndex        =   20
               Top             =   3090
               Width           =   3645
               _ExtentX        =   6429
               _ExtentY        =   3149
               _Version        =   262144
               BackColor       =   16777215
               BorderWidth     =   0
               BevelOuter      =   1
               BevelInner      =   2
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin XtremeSuiteControls.PushButton cmdBtn 
                  Height          =   1710
                  Index           =   1
                  Left            =   30
                  TabIndex        =   36
                  Top             =   15
                  Width           =   3570
                  _Version        =   851970
                  _ExtentX        =   6297
                  _ExtentY        =   3016
                  _StockProps     =   79
                  Caption         =   " 승인 시작(&R)"
                  ForeColor       =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   15.75
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Appearance      =   6
                  Picture         =   "frmKSNETCash.frx":06A5
               End
            End
            Begin Threed.SSOption optGubun 
               Height          =   300
               Index           =   0
               Left            =   1200
               TabIndex        =   25
               Top             =   1035
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   529
               _Version        =   262144
               Font3D          =   3
               ForeColor       =   255
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "소득공제용"
               Value           =   -1
            End
            Begin Threed.SSOption optGubun 
               Height          =   300
               Index           =   1
               Left            =   1200
               TabIndex        =   26
               Top             =   1440
               Width           =   1560
               _ExtentX        =   2752
               _ExtentY        =   529
               _Version        =   262144
               Font3D          =   3
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "지출증빙용"
            End
            Begin XtremeSuiteControls.PushButton cmdBtn 
               Height          =   780
               Index           =   0
               Left            =   2850
               TabIndex        =   33
               Top             =   990
               Width           =   1950
               _Version        =   851970
               _ExtentX        =   3440
               _ExtentY        =   1376
               _StockProps     =   79
               Caption         =   " 카드사용(&C)"
               ForeColor       =   0
               Enabled         =   0   'False
               Appearance      =   6
               Picture         =   "frmKSNETCash.frx":10B7
            End
            Begin VB.Label Label1 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "취소구분:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   9
               Left            =   60
               TabIndex        =   35
               Top             =   615
               Width           =   1080
            End
            Begin VB.Label Label1 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "메 시 지:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   10
               Left            =   30
               TabIndex        =   32
               Top             =   4980
               Width           =   1080
            End
            Begin VB.Label lblMessage1 
               AutoSize        =   -1  'True
               Caption         =   "#"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   180
               Left            =   1200
               TabIndex        =   29
               Top             =   4965
               Width           =   105
            End
            Begin VB.Label lblMessage2 
               AutoSize        =   -1  'True
               Caption         =   "#"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   180
               Left            =   1200
               TabIndex        =   28
               Top             =   5265
               Width           =   105
            End
            Begin VB.Label Label1 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "거래구분:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   0
               Left            =   60
               TabIndex        =   24
               Top             =   225
               Width           =   1080
            End
            Begin VB.Label Label1 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "사용자정보:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   4
               Left            =   60
               TabIndex        =   23
               Top             =   2460
               Width           =   1080
            End
            Begin VB.Label Label1 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "총금액:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   2
               Left            =   60
               TabIndex        =   22
               Top             =   2010
               Width           =   1080
            End
            Begin VB.Label Label1 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "사인패드:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   3
               Left            =   60
               TabIndex        =   21
               Top             =   3165
               Width           =   1080
            End
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   615
         Left            =   15
         TabIndex        =   1
         Top             =   6885
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   1085
         _Version        =   262144
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   480
            Index           =   2
            Left            =   3690
            TabIndex        =   2
            Top             =   45
            Width           =   1260
            _Version        =   851970
            _ExtentX        =   2222
            _ExtentY        =   847
            _StockProps     =   79
            Caption         =   " 취소(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frmKSNETCash.frx":1AC9
         End
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   930
         Index           =   0
         Left            =   15
         TabIndex        =   3
         Top             =   15
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   1640
         _Version        =   262144
         BackColor       =   4210752
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Label lblErrMsg 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "#"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   4785
            TabIndex        =   27
            Top             =   660
            Width           =   90
         End
         Begin VB.Label lblMsg 
            BackStyle       =   0  '투명
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   690
            Left            =   600
            TabIndex        =   4
            Top             =   135
            Width           =   4275
         End
         Begin VB.Image Image 
            Height          =   360
            Index           =   1
            Left            =   105
            Picture         =   "frmKSNETCash.frx":24DB
            Top             =   120
            Width           =   360
         End
      End
   End
End
Attribute VB_Name = "frmKSNETCash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iFlag            As Integer ' 신용승인-1 , 신용취소-2, 현금영수증-3, 현금영수증취소-4 를 구분하기 위한 플래그

Dim sStatus          As String

Dim 사업자번호       As String
Dim 단말기번호       As String
Dim VAN_IP           As String
Dim VAN_PORT         As String

Dim KS7500_CommPort  As String
Dim KS7500_BaudRate  As String

Dim SignPad_CommPort As String
Dim SignPad_BaudRate As String

Dim 현금결제         As String

'---------------------------------------------------------------------------------------
' 아래 함수는 승인요청 후 응답전문을 분석하기 위해서 만든 함수 입니다
' (작성한 함수 이므로 필요에 맞게 만들어 사용하시기 바랍니다)
'---------------------------------------------------------------------------------------
Public Function MyMid(Str As String, startposition As Integer, Num As Integer) As String
    Dim i      As Integer
    Dim chlen  As Integer
    Dim result As String
        
    For i = 1 To Len(Str)
        If Asc(Mid$(Str, i, 1)) < 0 Then
            chlen = chlen + 2
        Else
            chlen = chlen + 1
        End If
        
        If (chlen >= startposition) And (chlen <= startposition + Num - 1) Then
            result = result + Mid(Str, i, 1)
        End If
     Next i
    
    MyMid = result
End Function

        
'Public Sub 현금영수증승인요청_IC_Rtn()
'
'    Dim sRtnCode    As String
'    Dim Rtn     As Integer
'    Dim telegram As String
'
'    cmdBtn(2).Enabled = False
'    SSPanel2.Caption = "   단말기의 종료 버튼으로 취소"
'
'    KSNet_Dongle.EnableLogging
'    KSNet_Dongle.ClosePort
'    DoEvents
'
'    '설정 및 연결
'    Rtn = KSNet_Dongle.SetComPortEncPCPOS(CInt(KS7500_CommPort), CLng(KS7500_BaudRate))
'    If Rtn < 0 Then
'        lblMsg.Caption = "단말기 장치가 연결되어 있지 않습니다!" & vbNewLine & "전산실: 031)522-2025 연락 주세요."
'        cmdBtn(2).Tag = "0"
'        cmdBtn(2).Enabled = True
'        Exit Sub
'
'    Else
'        lblMsg.Caption = "현금영수증카드를 사용 한다면 단말기에 카드리딩 한다"
'    End If
'
'
'    '// 부가세, 면세금액, 봉사료 항목은 최대 2개까지만 조합해서 사용 가능
'    telegram = ""
'    telegram = telegram & IIf(cboGubun.Text = "현금승인", "C1", "D1")               ' 전문코드 (C1:현금영수증승인, D1:현금영수증취소)
'    telegram = telegram & "00"                                      ' 할부개월 (00=일시불거래, 초기값 space)
'    telegram = telegram & IIf(optGubun(0).Value = True, "1", "2")   ' 현금영수증 소비자 구분 (1.소비자 소득공제, 2.사업자 지출증빙, 3.자진발급)
'    telegram = telegram & Mid(cboCancel.Text, 2, 1)                 ' 현금영수증 취소 사유 (1.거래취소, 2:오류발급취소, 3:기타, space:단말기에서 입력)
'    telegram = telegram & Right("000000000" & IIf(iFlag = "3", CStr(txtMoney.Value), Replace(Spread_GetData(sprGrid, 5, 1, False), ",", "")), 9) ' 총 승인금액
'    telegram = telegram & "000000000"                               ' 부가세
'    telegram = telegram & "000000000"                               ' 면세금액
'    telegram = telegram & "000000000"                               ' 봉사료
'    telegram = telegram & Right(Space(6) & IIf(iFlag = "3", "", Spread_GetData(sprGrid, 2, 1, True)), 6)      ' 원거래일시 승인시 space, 취소시에도 space면 단말기에서 입력
'    telegram = telegram & Right(Space(9) & IIf(iFlag = "3", "", Spread_GetData(sprGrid, 1, 1, True)), 9)      ' 원거래승인번호  승인시 space, 취소시에도 space면 단말기에서 입력
'    telegram = telegram & Right("0000" & pnlNum.Caption, 4)                                                      ' 거래일련번호  승인시 space, 취소시에도 space면 단말기에서 입력
'    telegram = telegram & Chr(3)                    '// ETX
'
'    telegram = Chr(2) & Right("0000" & CStr(Len(telegram)), 4) & telegram
'
'    sRtnCode = KSNet_Dongle.EncPCPOSWrite(telegram, Len(telegram), 1)
'
'        '// 이 후 OnRecvEncPCPOS에서 처리
'
'
'    If sRtnCode <> "00" Then
'
'        ' 단말기 단독 거래의 결과를 전송 받지 않기 위하여.
'        ' 단독 거래중 포트가 오픈되어 있으면 결과 값을 전송 받아 버림.
'
'        cmdBtn(1).Enabled = True
'        cmdBtn(2).Enabled = True
'        cmdBtn(2).Tag = "0"
'        DoEvents
'
'        KSNet_Dongle.EnableLogging
'        KSNet_Dongle.ClosePort
'        lblMsg.Caption = "단말기를 확인하여 주십시요." & vbNewLine & "카드를 제거 후 다시 시도 하여 주십시요."
'
'        With lblMessage1
'            Select Case sRtnCode
'                'Case "00": .Caption = "정상 승인"  '<-- 정상 상징적 처리
'                Case "22": .Caption = "암호화 오류"
'                Case "21": .Caption = "S/W 유효성 오류"
'                Case "40": .Caption = "타임 아웃"
'                Case "50": .Caption = "카드 미 입력(IC 미 삽입)"
'                Case "60": .Caption = "2nd Generation 에러 카드 거절"
'                Case "90": .Caption = "리더 상태 변경 실패"
'                Case "91": .Caption = "리더 인증 코드 불일치"
'
'                Case "01": .Caption = "chip 미 응답"
'                Case "02": .Caption = "application 미 존재"
'                Case "03": .Caption = "chip 데이터 읽기 실패"
'                Case "04": .Caption = "mandatory 데이터 미 포함"
'                Case "05": .Caption = "CVM 커맨드 응답실패"
'                Case "06": .Caption = "EMV 커맨드 오 설정"
'                Case "07": .Caption = "터미널(리더)오 동작"
'
'                Case "30": .Caption = "chip block"
'                Case "31": .Caption = "application block"
'                Case "32": .Caption = "카드 자체 block"
'
'                Case "11": .Caption = "키 유효기간 지남"
'                Case "12": .Caption = "암호화 키 생성 실패"
'                Case "13": .Caption = "이미 암호화 키 있음"
'                Case "14": .Caption = "KEY 유효성 검증 오류"
'                Case "15": .Caption = "IPEK KEY 없음"
'                Case "16": .Caption = "사용될 IPEK 의 년도 Data 없음"
'
'                Case "ZA": .Caption = "STX 수신 오류"
'                Case "ZB": .Caption = "ETX 수신 오류"
'                Case "ZC": .Caption = "LRC 오류"
'                Case "ZD": .Caption = "단말기 mode 오류"
'                Case "ZE": .Caption = "함수 인자 값 오류"
'                Case "ZF": .Caption = "시리얼포트 설정 하지 오류"
'                Case "ZG": .Caption = "시리얼포트가 열려 있지 오류"
'                Case "ZH": .Caption = "테이터 생성 실패"
'                Case "ZI": .Caption = "데이터 송신 실패"
'                Case "ZJ": .Caption = "테이터 수신 실패"
'                Case "ZK": .Caption = "데이터 송수신 대기 시간 초과"
'
'            End Select
'        End With
'    End If
'End Sub


Public Sub 현금영수증승인요청_IC_Start()

    cmdBtn(2).Enabled = True
    SSPanel2.Caption = "   단말기의 종료 버튼으로 취소"
    
    
    Dim sD As String
    Dim sE As String
    Dim ReturnValue As Long


    
        
    sD = SetMessage(IIf(iFlag = "3", Cash_Approve, Cash_Cancel_Today), IIf(iFlag = "3", CStr(txtMoney.Value), Replace(Spread_GetData(sprGrid, 5, 1, False), ",", "")), IIf(optGubun(0).Value = True, "00", "01"), IIf(iFlag = "3", "", Spread_GetData(sprGrid, 1, 1, True)), IIf(iFlag = "3", "", Spread_GetData(sprGrid, 2, 1, True)))
    Call frmKicc.Card_Approve(sD, Me.Name)
    
    
End Sub
            
'-------------------------------------------------------------------------------
' 함수명 : 현금영수증승인요청_Rtn
'
'
'-------------------------------------------------------------------------------
Public Sub 현금영수증승인요청_Rtn(Gbn As String)
    'MsgBox ("현금영수증 승인요청을 시작합니다")
    
    
    txtUserInfo.Text = ""
    lblErrMsg.Caption = ""
    
    If (사업자번호 = "" Or Len(사업자번호) <> 10) Then
        lblMsg.Caption = "사업자번호가 올바르지 않습니다. 취소 버튼을 클릭하세요."
        
        Exit Sub
    End If
    
    If (단말기번호 = "") Then
        lblMsg.Caption = "단말기번호가 올바르지 않습니다. 취소 버튼을 클릭하세요."
        Exit Sub
    End If
    
    If (VAN_IP = "") Or (VAN_PORT = "") Then
        lblMsg.Caption = "승인서버 정보가 올바르지 않습니다. 취소 버튼을 클릭하세요."
        
        Exit Sub
    End If
        
        
    iFlag = Gbn ' 현금영수증임을 나타내는 플래그 셋팅
    
    TabControl.SelectedItem = 0
    If iFlag = "3" Then
        cboGubun.Text = "현금승인"
        TabControlPage1.BackColor = &H8000000F
        optGubun(0).BackColor = &H8000000F
        optGubun(1).BackColor = &H8000000F
        cboGubun.Enabled = False
        cboCancel.Enabled = False
        
    Else
        cboGubun.Text = "현금취소"
        TabControlPage1.BackColor = &HFFC0FF
        optGubun(0).BackColor = &HFFC0FF
        optGubun(1).BackColor = &HFFC0FF
        cboCancel.Enabled = True
        cboCancel.ListIndex = 0
    
        cboGubun.Enabled = False
    End If

    '----------------------------------------------------------------------------------------------------------------------------------
    ' 보안이 적용된 단말기를 사용할 경우
    '----------------------------------------------------------------------------------------------------------------------------------
'    If 가맹점정보.CAT단말기종류 = "KICC" Then
        lblMsg.Caption = "반드시 승인 시작 번튼을 누른 후  IC 카드를 삽입 하여 주십시요."
        Exit Sub
'    Else
'        MsgBox "지원하지 않는 단말기 입니다." & vbCrLf & "단말기 설정을 확인하여 주십시요"
'    End If

End Sub

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 0

        
        Case 1:
            cmdBtn(1).Enabled = False
            cmdBtn(2).Tag = "60"
        
            Call 현금영수증승인요청_IC_Start
    
        Case 2:
            Dim sD As String
            Dim sE As String
            sD = "TM"
            Call frmKicc.Card_Approve(sD, Me.Name)
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Me.Top = ((frmMain.Height - Me.Height) / 2) + frmMain.Top
    Me.Left = ((frmMain.Width - Me.Width) / 2) + frmMain.Left
    
    With cboGubun
        .Clear
        .AddItem "현금승인"
        .AddItem "현금취소"
        
        .ListIndex = 0
    End With
    
    With cboCancel
        .Clear
        .AddItem "01.거래취소"
        .AddItem "02.오류발급취소"
        .AddItem "03.기타"
        
        .ListIndex = 0
    End With
    
    lblMessage1.Caption = ""
    lblMessage2.Caption = ""
        
    '-------------------------------------------------------------------
    '
    '-------------------------------------------------------------------
    Query = "SELECT * FROM TB_기본정보"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    If ADORs.EOF Then
        사업자번호 = ""
        단말기번호 = ""
        
        VAN_IP = ""
        VAN_PORT = ""
    Else
        사업자번호 = Trim(Replace(ADORs!사업자번호, "-", "")) & "" '
        단말기번호 = Trim(ADORs!단말기번호) & ""                   '
        
        VAN_IP = ADORs!VAN_IP & ""                                 '
        VAN_PORT = ADORs!VAN_PORT & ""                             '
    End If
    ADORs.Close
    Set ADORs = Nothing
    
    KS7500_CommPort = GetIniStr("VAN", "KS7500_CommPort", "", iniFile) '
    KS7500_BaudRate = GetIniStr("VAN", "KS7500_BaudRate", "", iniFile) '
    
    SignPad_CommPort = GetIniStr("VAN", "SignPad_CommPort", "", iniFile) '
    SignPad_BaudRate = GetIniStr("VAN", "SignPad_BaudRate", "", iniFile) '
End Sub


'Private Sub KS7500_OnReadMSR(ByVal Data As String)
'    lblMessage1.Caption = ""
'    lblMessage2.Caption = ""
'
'    'MsgBox ("카드정보를 수신하였습니다!")
'
'    If iFlag > 2 Then
'        txtUserInfo.Text = Data
'
'        '현금영수증 카드를 통한 거래이기 때문에 싸인패드(핀패드장비)를 초기화 시킨후 포트를 닫는다
'
'        '싸인패드 초기화
'        KSSignpad.SignComReqAC
'        KSSignpad.ClosePort
'
'        lblMsg.Caption = "현금영수증 승인/취소 요청을 시작합니다!" '현금영수증카드를 이용한 리딩시에 승인처리를 시작한다.
'
'        ' 현금영수증 승인/취소 을 하기 위한 승인전문 구성 작업 진행
'
'        Dim SendTmp As String        '요청전문을 구성하기 위한 변수
'        Dim ReqLen As Integer        '요청전문의 길이
'        Dim RecvMsg As String * 1024 '응답전문을 수신하기 위한 변수
'
'        ' 특별한 설명이 붙지 않은 값들은 아래와 같이 그대로 사용한다.
'        ' 전문 작성 (204전문 작성)
'
'        If iFlag = 3 Then
'            '현금영수증승인요청 일 경우 전문 구성
'
'            현금결제 = txtMoney.Value
'            현금결제 = Trim(현금결제)
'
'            SendTmp = Chr(2)                                               ' STX
'            SendTmp = SendTmp & "bq"                                       ' 거래구분
'            SendTmp = SendTmp & 단말기번호                                 ' "DPT0TEST05" 단말기번호 (Test용)(현재는들어있는 단말기 번호는 테스트 번호 이며 실제 사용할 경우
'                                                                           ' KSNET을 통해서 실제단말번호를 받아야 한다.영업팀에 문의 해야함. 승인과 거래시 단말기 번호는 동일해야함)
'            SendTmp = SendTmp & "        "                                 ' 업체정보
'            SendTmp = SendTmp & "000000"                                   ' 거래일련번호(전문번호) (일자별로 유니크한 값들을 입력하면되며 사용하지 않을 경우 "000000"을 사용해도 됨)
'            SendTmp = SendTmp & "0"                                        ' 거래유형(0 으로 사용한다)
'            SendTmp = SendTmp & "S"                                        ' POS Entry Mode (7500 단말기를 통한 현금영수증 카드 리딩이므로 'S' 를 사용한다)
'            SendTmp = SendTmp & txtUserInfo.Text                           ' 현금영수증 번호
'            SendTmp = SendTmp & Chr(28)                                    ' FS (0x1c)
'            SendTmp = SendTmp & Format(현금결제, "000000000")              ' 공급금액(총금액 = 세금 + 봉사료 + 공급금액) 이어야 한다.(반드시 맞춰 준다)
'            SendTmp = SendTmp & "000000000"                                ' 세금
'            SendTmp = SendTmp & "000000000"                                ' 봉사료
'            SendTmp = SendTmp & Format(현금결제, "000000000")              ' 총금액  ..현금영수증 테스트 금액은 700,7000원 입니다.
'
'            If optGubun(0).Value = True Then
'                SendTmp = SendTmp & "0"                                    ' 거래구분자(개인이 사용한 경우는 0 , 사용자가 사용하는 경우 1)
'            Else
'                SendTmp = SendTmp & "1"                                    ' 거래구분자(개인이 사용한 경우는 0 , 사용자가 사용하는 경우 1)
'            End If
'
'            '=========그대로 사용 ====================
'            SendTmp = SendTmp & " "                                        ' 포인트 POS Entry Mode
'            SendTmp = SendTmp & "                                     "    ' 포인트 Track II
'            SendTmp = SendTmp & Chr(28)                                    ' FS (0x1c)
'            SendTmp = SendTmp & "      "                                   ' 상품코드 ~
'            SendTmp = SendTmp & "  "                                       ' 가맹점 사용ID
'            SendTmp = SendTmp & "                              "           ' 가맹점사용필드
'            SendTmp = SendTmp & "                                                                   " ' 여유필드
'            SendTmp = SendTmp & Chr(3) & Chr(13)  ' FS ,ETX, CR
'             '===========================  여기까지
'
'        Else
'            '현금영수증취소요청 일 경우
'
'            SendTmp = Chr(2)                                         ' STX
'            SendTmp = SendTmp & "bs"                                 ' 거래구분
'            SendTmp = SendTmp & 단말기번호                           ' "DPT0TEST05" 단말기번호 (Test용)(현재는들어있는 단말기 번호는 테스트 번호 이며 실제 사용할 경우
'                                                                     ' KSNET을 통해서 실제단말번호를 받아야 한다.영업팀에 문의 해야함. 승인과 거래시 단말기 번호는 동일해야함)
'            SendTmp = SendTmp & "        "                           ' 업체정보
'            SendTmp = SendTmp & "000000"                             ' 거래일련번호(전문번호) (일자별로 유니크한 값들을 입력하면되며 사용하지 않을 경우 "000000"을 사용해도 됨)
'            SendTmp = SendTmp & "0"                                  ' 거래유형(0으로 사용)
'            SendTmp = SendTmp & "S"                                  ' POS Entry Mode (7500 단말기를 통한 현금영수증 카드 리딩이므로 'S' 를 사용한다)
'            SendTmp = SendTmp & txtUserInfo.Text                     ' 현금영수증(사용자정보) 번호
'            SendTmp = SendTmp & Chr(28)                              ' FS (0x1c)
'            SendTmp = SendTmp & Format(Spread_GetData(sprGrid, 5, 1, True), "000000000") ' 공급금액(총금액 = 세금 + 봉사료 + 공급금액) 이어야 한다.(반드시 맞춰 준다)
'            SendTmp = SendTmp & "000000000"                                              ' 세금
'            SendTmp = SendTmp & "000000000"                                              ' 봉사료
'            SendTmp = SendTmp & Format(Spread_GetData(sprGrid, 5, 1, True), "000000000") ' 총금액   (취소일 경우 사용자정보와 금액,현금승인번호,거래일자가 일치 해야 한다)
'            SendTmp = SendTmp & Spread_GetData(sprGrid, 4, 1, True)                      ' 거래구분자(개인이 사용한 경우는 0 , 사용자가 사용하는 경우 1)
'            SendTmp = SendTmp & Spread_GetData(sprGrid, 1, 1, True)  ' 원거래현금승인번호(9자리)
'            SendTmp = SendTmp & "            "                       ' 원거래포인트승인번호(space 12자리로 사용)
'            SendTmp = SendTmp & Spread_GetData(sprGrid, 2, 1, True)  ' 거래일자(6자리)- YYMMDD
'            '=========그대로 사용 ====================
'            SendTmp = SendTmp & " "                                       ' 포인트 POS Entry Mode
'            SendTmp = SendTmp & "                                     "   ' 포인트 Track II
'            SendTmp = SendTmp & Chr(28)                                   ' FS (0x1c)
'            SendTmp = SendTmp & "      "                                  ' 상품코드 ~
'            SendTmp = SendTmp & "  "                                      ' 가맹점 사용ID
'            SendTmp = SendTmp & "                              "          ' 가맹점사용필드
'            SendTmp = SendTmp & "                                                             " ' 여유필드
'            SendTmp = SendTmp & Left(Trim(Left(cboCancel.Text, 2)) & "01", 2)
'            SendTmp = SendTmp & Chr(3) & Chr(13)  ' FS ,ETX, CR
'        End If
'
'        ' STX~CR까지의 길이를 전문의 가장 앞에 4바이트 크기의 문자열로 붙입니다.
'        ' 즉, 전문의 총 길이가 475이면  "0475"를 요청전문의 맨 앞에 붙이면 됩니다.
'
'        '전체요청전문의 길이를 구한다.
'        ReqLen = Len(SendTmp)
'        SendTmp = Format(ReqLen, "0000") & SendTmp
'
'        ' 요청전문을 ADSL모듈을 통해 KSNET으로 전송합니다.
'        ' 주의사항 : 3번째 파라미터(이름 : Media)는 전자서명 신용승인/취소 일때만 4입니다.
'        ' 현금영수증, 포인트, 직불 등등에서 Media값은 2입니다.
'
'        'Rtn = ReqAppr("210.181.28.116", 9531, 2, SendTmp, Len(SendTmp), RecvMsg, 15, 0)   '뒤에 숫자 2개는 는 그대로 사용 하시기 바랍니다(변경하지 마세요)
'        Rtn = ReqAppr(VAN_IP, CInt(VAN_PORT), 2, SendTmp, Len(SendTmp), RecvMsg, 15, 0)    '뒤에 숫자 2개는 는 그대로 사용 하시기 바랍니다(변경하지 마세요)
'
'        If Rtn > 0 Then
'            lblMsg.Caption = "KSNET으로 현금영수증 승인 요청/취소 성공(통신성공)!!" ' 성공했다면 통신은 성공하였기 때문에 승인 성공/거절을 구분하여 처리한다.
'
'            sStatus = MyMid(RecvMsg, 32, 1)            ' 상태
'
'            With sprGrid
'                .Col = 1
'
'                .Row = 1:  .Text = MyMid(RecvMsg, 76, 9) & ""   '승인번호(거절코드)
'                .Row = 2:  .Text = MyMid(RecvMsg, 33, 6) & ""   '승인일자
'                .Row = 3:  .Text = MyMid(RecvMsg, 39, 4) & ""   '승인시간
'
'                If optGubun(0).Value = True Then
'                    .Row = 4:  .Text = "0"                      '
'                Else
'                    .Row = 4:  .Text = "1"                      '
'                End If
'
'                .Row = 5:  .Text = txtMoney.Value               '결제금액
'                .Row = 6:  .Text = txtUserInfo.Text             '사용자정보
'                .Row = 7:  .Text = MyMid(RecvMsg, 43, 16) & ""  '메시지1
'                .Row = 8:  .Text = MyMid(RecvMsg, 59, 16) & ""  '메시지2
'                .Row = 9:  .Text = MyMid(RecvMsg, 75, 1) & ""   '소득/비소득 구분
'                .Row = 10: .Text = MyMid(RecvMsg, 85, 20) & ""  '국세청1
'                .Row = 11: .Text = MyMid(RecvMsg, 105, 20) & "" '국세청2
'                .Row = 12: .Text = Left(Trim(Left(cboCancel.Text, 2)) & "01", 2) & "" '취소사유
'            End With
'
'            lblMessage1.Caption = MyMid(RecvMsg, 43, 16) & ""  '메시지1
'            lblMessage2.Caption = MyMid(RecvMsg, 59, 16) & ""  '메시지2
'            DoEvents
'
'            If sStatus = "O" Then
'                lblMsg.Caption = "KSNET으로 승인 요청/취소 성공(통신성공)" ' 성공했다면 통신은 성공하였기 때문에 승인 성공/거절을 구분하여 처리한다.
'
'                KS7500.ClosePort '카드정보 수신 후 포트를 닫는다
'
'
'                If iFlag = "3" Then
'
'                    ' 승인된 현금 영수증 자료를 SQL에 저장한다.
'                    Call Save_현금영수증승인정보
'
'
'
'                    Select Case Account_Form
'                        Case "접수"
'                            Call Move_승인자료정보(frm접수결제.sprCash)
'
'                        Case "출고"
'                            Call Move_승인자료정보(frm출고결제.sprCash)
'
'                        Case "판매취소2"
'                            Call Move_승인자료정보(frm판매취소.sprCash)
'
'                    End Select
'
'                Else
'
'                    ' 현금영수증 취소 정보를 SQL에 저장
'                    Call Save_현금영수증취소정보("S")
'
'
'                    Select Case Account_Form
'                        Case "접수"
'                            With frm접수결제.sprCash
'                                .Col = 1
'
'                                For i = 1 To .MaxRows
'                                    .Row = i: .Text = ""
'                                Next i
'                            End With
'
'                        Case "접수2"   '취소
'
'                            Call 현금영수증취소_Report(frm현금영수증승인.KS7500i, pnlApprovalNo.Caption, pnlApprovalDay.Caption, pnlApprovalTime.Caption)
'
'                            frm현금영수증승인.Data_Display
'
'                        Case "출고"
'                            With frm출고결제.sprCash
'                                .Col = 1
'
'                                For i = 1 To .MaxRows
'                                    .Row = i: .Text = ""
'                                Next i
'                            End With
'
'                        Case "판매취소"
'                            With frm판매취소결제.sprCash
'                                .Col = 1
'
'                                For i = 1 To .MaxRows
'                                    .Row = i: .Text = ""
'                                Next i
'                            End With
'
'                        Case "판매취소2"
'                            With frm판매취소.sprCash
'                                .Col = 1
'
'                                For i = 1 To .MaxRows
'                                    .Row = i: .Text = ""
'                                Next i
'                            End With
'                    End Select
'                End If
'
'
'                Unload Me
'            Else
'                lblMsg.Caption = "승인 요청/취소 거절 되었습니다! 거절 메세지를 확인하세요"
'            End If
'        Else
'            ' 통신에러가 발생한 경우 이므로 에러 처리를 한다.
'            lblMsg.Caption = "승인 요청/취소 실패(통신에러가 발생 하였습니다!! 다시 거래 하시기 바랍니다"
'            lblErrMsg.Caption = RecvMsg
'        End If
'    End If
'
'    KS7500.ClosePort '카드정보 수신 후 포트를 닫는다
'End Sub

'---------------------------------------------------------------------------------------------------------------
' KS4060을 사용하여 사인패드로 처리한 경우 여기서 처리가된다.
' 보안 모드가 적용이 되어서 처리가 되는 경우임
'(이미 CAT단말기에서 승인 처리가 된 정보가 넘어옴.)
'---------------------------------------------------------------------------------------------------------------
Private Sub KSNet_Dongle_OnRecvEncPCPOS(ByVal errCode As String, ByVal Data As String, ByVal dataLen As Long)
' 승인시 발송 내역 [ ] 내용은 제외 하고 처리할것.
'errCode [00]
'data[0276A100  000001004000000091000000000000000000150819115812401152652    71112429            05신한카드        05신한카드        신한카드        OK: 01152652    전자서명전표                                                                    1544-7000           0001356297******08527]
'dataLen [282]
        
'errCode [00]
'data[0276A100  000001004000000091000000000000000000150819115812401152652    71112429            05신한카드        05신한카드        신한카드        OK: 01152652    전자서명전표                                                                    1544-7000           0001356297******08527]
'dataLen [282]
'errCode [00]
'data[0276B100  000001004000000091000000000000000000150819120116401152652    71112429            05신한카드        05신한카드        신한카드        취소OK: 01152652전자서명전표                                                                    1544-7000           0001356297******0852y]
'dataLen [282]
        
'        taResponse.Text = taResponse.Text & "errCode[" & errCode & "]" & vbNewLine
'        taResponse.Text = taResponse.Text & "data[" & data & "]" & vbNewLine
'        taResponse.Text = taResponse.Text & "dataLen[" & CStr(dataLen) & "]" & vbNewLine
    Dim sHelpDesk   As String
        
    On Error GoTo ERR_RTN
    
        
    If errCode <> "00" Then
        
        ' 통신에러가 발생한 경우 이므로 에러 처리를 한다.
        lblMsg.Caption = "승인 요청/취소 실패 " & vbNewLine & "(다시 거래 하시기 바랍니다)"
        
        With lblMessage1
            Select Case errCode
'               Case "00": .Caption = "정상 승인"  '<-- 정상 상징적 처리
                Case "22": .Caption = "암호화 오류"
                Case "21": .Caption = "S/W 유효성 오류"
                Case "40": .Caption = "타임 아웃"
                Case "50": .Caption = "카드 미 입력(IC 미 삽입)"
                Case "60": .Caption = "2nd Generation 에러 카드 거절"
                Case "90": .Caption = "리더 상태 변경 실패"
                Case "91": .Caption = "리더 인증 코드 불일치"
                
                Case "01": .Caption = "chip 미 응답"
                Case "02": .Caption = "application 미 존재"
                Case "03": .Caption = "chip 데이터 읽기 실패"
                Case "04": .Caption = "mandatory 데이터 미 포함"
                Case "05": .Caption = "CVM 커맨드 응답실패"
                Case "06": .Caption = "EMV 커맨드 오 설정"
                Case "07": .Caption = "터미널(리더)오 동작"
                
                Case "30": .Caption = "chip block"
                Case "31": .Caption = "application block"
                Case "32": .Caption = "카드 자체 block"
                
                Case "11": .Caption = "키 유효기간 지남"
                Case "12": .Caption = "암호화 키 생성 실패"
                Case "13": .Caption = "이미 암호화 키 있음"
                Case "14": .Caption = "KEY 유효성 검증 오류"
                Case "15": .Caption = "IPEK KEY 없음"
                Case "16": .Caption = "사용될 IPEK 의 년도 Data 없음"
                
                Case "ZA": .Caption = "STX 수신 오류"
                Case "ZB": .Caption = "ETX 수신 오류"
                Case "ZC": .Caption = "LRC 오류"
                Case "ZD": .Caption = "단말기 mode 오류"
                Case "ZE": .Caption = "함수 인자 값 오류"
                Case "ZF": .Caption = "시리얼포트 설정 하지 오류"
                Case "ZG": .Caption = "시리얼포트가 열려 있지 오류"
                Case "ZH": .Caption = "테이터 생성 실패"
                Case "ZI": .Caption = "데이터 송신 실패"
                Case "ZJ": .Caption = "테이터 수신 실패"
                Case "ZK": .Caption = "데이터 송수신 대기 시간 초과"
            
            End Select
        End With
        
        ' 오류가 나서 다시 시도할 경우 아무런 처리를 하지 않고 바로 단말기에서 처리가 되는지 확인 할것.
        cmdBtn(1).Enabled = True
        cmdBtn(2).Enabled = True
        Exit Sub
    End If
    
    ' 3. 거절 응답 : NE (HelpDesk에 사유 입력됨)
    If MyMid(Data, 6, 2) = "NE" Then
    
        sHelpDesk = Trim(MyMid(Data, 241, 20))
        Select Case sHelpDesk
            
            Case "단말기종료키누름"
                cmdBtn(1).Enabled = True
                cmdBtn(2).Enabled = True
                Unload Me
        
                Exit Sub
                
            Case Else
                lblMsg.Caption = sHelpDesk
                cmdBtn(1).Enabled = True
                cmdBtn(2).Enabled = True
                DoEvents
            
        
        End Select
                
                
        Exit Sub
    
    End If
    
            
    
        
    lblMsg.Caption = "현금영수증 승인 요청/취소 성공(통신성공)!!" ' 성공했다면 통신은 성공하였기 때문에 승인 성공/거절을 구분하여 처리한다.
    
    sStatus = MyMid(Data, 61, 12) & ""   ' 승인번호, 거절시 오류코드, 없으면 space           ' 상태
                                        ' 사용자 정보 부족으로 오류 발생시 space가 넘어 온것을 확인
    
    With sprGrid
        .Col = 1
            
        Debug.Print "전문코드:" & MyMid(Data, 6, 2) & ""    ' 3. 거절응답 : NE
        Debug.Print "승인번호:" & MyMid(Data, 61, 9) & ""    ' 승인번호(거절코드) CAT에서부터 12자리로 되어 있음 SQL에는 nvarchar(10)으로 되어 있음 ㅡㅡ
        Debug.Print "승인일자:" & MyMid(Data, 48, 6) & ""   ' 승인일자
        Debug.Print "승인시간:" & MyMid(Data, 54, 6) & ""   ' 승인시간
        
        Debug.Print "소비자 구분:" & MyMid(Data, 10, 1) & ""   ' 현금 영수증 소비자 구분 (POS에서 전송된 데이터 그래도 리턴(1.소득공제, 2.지출증빙 3.자진발급)
        
        Debug.Print "결제금액:" & Val(MyMid(Data, 12, 9))   ' 결제금액
        Debug.Print "사용자정보:" & Left(MyMid(Data, 265, 19), (InStr(MyMid(Data, 265, 19), Chr(3)) - 1)) & "" ' 사용자정보 (전체 카드번호 중 1-12자리를 ***** 표시하여 전달
        Debug.Print "메시지1:" & MyMid(Data, 129, 16) & ""  ' 메시지1
        Debug.Print "메시지2:" & MyMid(Data, 145, 16) & ""  ' 메시지2
        Debug.Print "소득/비소득 구분:" & MyMid(Data, 10, 1) & ""   ' 소득/비소득 구분
        Debug.Print "국세청1:" & MyMid(Data, 161, 20) & ""  ' 국세청1
        Debug.Print "국세청2:" & MyMid(Data, 201, 20) & "" ' 국세청2
        Debug.Print "HelpDesk:" & MyMid(Data, 241, 20) & "" ' HelpDesk
        
        .Row = 1:  .Text = MyMid(Data, 61, 9) & ""   ' 승인번호(거절코드) CAT에서부터 12자리로 되어 있음 SQL에는 nvarchar(10)으로 되어 있음 ㅡㅡ
        .Row = 2:  .Text = MyMid(Data, 48, 6) & ""   ' 승인일자
        .Row = 3:  .Text = MyMid(Data, 54, 6) & ""   ' 승인시간
        
        .Row = 4:  .Text = MyMid(Data, 10, 1) & ""   ' 현금 영수증 소비자 구분 (POS에서 전송된 데이터 그래도 리턴(1.소득공제, 2.지출증빙 3.자진발급)
        
        .Row = 5:  .Text = Val(MyMid(Data, 12, 9))   ' 결제금액
        .Row = 6:  .Text = Left(MyMid(Data, 265, 19), (InStr(MyMid(Data, 265, 19), Chr(3)) - 1)) & ""  ' 사용자정보 (전체 카드번호 중 1-12자리를 ***** 표시하여 전달
        .Row = 7:  .Text = MyMid(Data, 129, 16) & ""  ' 메시지1
        .Row = 8:  .Text = MyMid(Data, 145, 16) & ""  ' 메시지2
        .Row = 9:  .Text = MyMid(Data, 10, 1) & ""   ' 소득/비소득 구분
        .Row = 10: .Text = MyMid(Data, 161, 20) & ""  ' 국세청1
        .Row = 11: .Text = MyMid(Data, 201, 20) & "" ' 국세청2
    End With
    
    lblMessage1.Caption = MyMid(Data, 129, 16) & ""  '메시지1
    lblMessage2.Caption = MyMid(Data, 145, 16) & ""  '메시지2
    DoEvents
    
     
    ' 현금 영수증 승인 번호는 9자리로 확인됨
    If Len(Trim(sStatus)) = 9 Then
        lblMsg.Caption = "KSNET으로 승인 요청/취소 성공(통신성공)" ' 성공했다면 통신은 성공하였기 때문에 승인 성공/거절을 구분하여 처리한다.
        
        If iFlag = "3" Then
        
            ' 승인된 현금 영수증 자료를 SQL에 저장한다.
            Call Save_현금영수증승인정보

            
            Select Case Account_Form
                Case "접수"
                    Call Move_승인자료정보(frm접수결제.sprCash)
                
                Case "출고"
                    Call Move_승인자료정보(frm출고결제.sprCash)
                        
                Case "판매취소2"
                    Call Move_승인자료정보(frm판매취소.sprCash)
            End Select
        Else
            ' 현금영수증 취소 정보를 SQL에 저장
            Call Save_현금영수증취소정보("K")
            
            Select Case Account_Form
                Case "접수"
                    With frm접수결제.sprCash
                        .Col = 1
                        
                        For i = 1 To .MaxRows
                            .Row = i: .Text = ""
                        Next i
                    End With
                
                Case "접수2"   '취소
'                        Call 현금영수증취소_Report(frm현금영수증승인.KS7500i, _
'                                                   Spread_GetData(sprGrid, 1, 1, True), _
'                                                   Spread_GetData(sprGrid, 2, 1, True), _
'                                                   Spread_GetData(sprGrid, 3, 1, True))
                    
                    '"KS4060 보안인증" 는 해당 단말기 에서 바로 출력 처리를 한다.
'                    If 가맹점정보.CAT단말기종류 <> "KS4060 보안인증" Then
'                        Call 현금영수증취소_Report(frm현금영수증승인.KS7500i, pnlApprovalNo.Caption, pnlApprovalDay.Caption, pnlApprovalTime.Caption)
'                    End If
                    
                    frm현금영수증승인.Data_Display
                
                Case "출고"
                    With frm출고결제.sprCash
                        .Col = 1
                        
                        For i = 1 To .MaxRows
                            .Row = i: .Text = ""
                        Next i
                    End With
                
                Case "판매취소"
                    With frm판매취소결제.sprCash
                        .Col = 1
                        
                        For i = 1 To .MaxRows
                            .Row = i: .Text = ""
                        Next i
                    End With
            
                    ' 2014-10-24일 추가.. 현금영수증 판매 취소한 경우 취소 영수증이 안나와서
                    결제취소여부 = True
            
                Case "판매취소2"
                    With frm판매취소.sprCash
                        .Col = 1
                        
                        For i = 1 To .MaxRows
                            .Row = i: .Text = ""
                        Next i
                    End With
            End Select
        End If
        
        Unload Me
    End If

ERR_RTN:
    lblMsg.Caption = Err.description

End Sub

 
'---------------------------------------------------------------------------------------------------------------
' KS7500i, KS7050i를 사용하여 사인패드로 처리한 경우 여기서 처리가된다.
'---------------------------------------------------------------------------------------------------------------
'Private Sub KSSignpad_OnRecvPinData(ByVal Data As String)
'    lblMessage1.Caption = ""
'    lblMessage2.Caption = ""
'
'    lblMsg.Caption = "핀패드 데이터를 수신하였습니다!"
'
'    KS7500.ClosePort    '싸인패드(핀패드)를 통한 입력이므로 7500i 단말기의 포트를 닫아 준다.
'    KSSignpad.ClosePort '싸인패드의 포트도 닫아준다.
'
'    Dim Pindata As String '핀데이터를 저장하기 위한 변수
'
'    Pindata = Trim(Data)
'    Pindata = Pindata & Space(37 - Len(Pindata)) '길이 - 37
'
'    txtUserInfo.Text = Pindata
'
'    ' 현금영수증 승인/취소 을 하기 위한 승인전문 구성 작업 진행
'
'    Dim SendTmp As String        '요청전문을 구성하기 위한 변수
'    Dim ReqLen As Integer        '요청전문의 길이
'    Dim RecvMsg As String * 1024 '응답전문을 수신하기 위한 변수
'
'    ' 특별한 설명이 붙지 않은 값들은 아래와 같이 그대로 사용한다.
'    ' 전문 작성 (204전문 작성)
'
'    If iFlag = 3 Then
'        '현금영수증승인요청 일 경우 전문 구성
'
'        현금결제 = txtMoney.Value
'        현금결제 = Trim(현금결제)
'
'        SendTmp = Chr(2)                                  ' STX
'        SendTmp = SendTmp & "bq"                          ' 거래구분
'        SendTmp = SendTmp & 단말기번호                    ' "DPT0TEST05" 단말기번호 (Test용)(현재는들어있는 단말기 번호는 테스트 번호 이며 실제 사용할 경우
'                                                          ' KSNET을 통해서 실제단말번호를 받아야 한다.영업팀에 문의 해야함. 승인과 거래시 단말기 번호는 동일해야함)
'        SendTmp = SendTmp & "        "                    ' 업체정보
'        SendTmp = SendTmp & "000000"                      ' 거래일련번호(전문번호) (일자별로 유니크한 값들을 입력하면되며 사용하지 않을 경우 "000000"을 사용해도 됨)
'        SendTmp = SendTmp & "0"                           ' 거래유형(0 으로 사용한다)
'        SendTmp = SendTmp & "K"                           ' POS Entry Mode (핀패드를 통한 입력이므로 'K' 를 사용한다)
'        SendTmp = SendTmp & Pindata                       ' 현금영수증 번호
'        SendTmp = SendTmp & Chr(28)                       ' FS (0x1c)
'        SendTmp = SendTmp & Format(현금결제, "000000000") ' 공급금액(총금액 = 세금 + 봉사료 + 공급금액) 이어야 한다.(반드시 맞춰 준다)
'        SendTmp = SendTmp & "000000000"                   ' 세금
'        SendTmp = SendTmp & "000000000"                   ' 봉사료
'        SendTmp = SendTmp & Format(현금결제, "000000000") ' 총금액
'
'        'SendTmp = SendTmp & "0"                           ' 거래구분자(개인이 사용한 경우는 0 , 사용자가 사용하는 경우 1)
'        '                                                  ' 여기서는 개인이라고 가정하고 0을 사용
'        If optGubun(0).Value = True Then
'            SendTmp = SendTmp & "0"                        ' 거래구분자(개인이 사용한 경우는 0 , 사용자가 사용하는 경우 1)
'        Else
'            SendTmp = SendTmp & "1"                        ' 거래구분자(개인이 사용한 경우는 0 , 사용자가 사용하는 경우 1)
'        End If
'
'        '=========그대로 사용 ====================
'        SendTmp = SendTmp & " "                                        ' 포인트 POS Entry Mode
'        SendTmp = SendTmp & "                                     "    ' 포인트 Track II
'        SendTmp = SendTmp & Chr(28)                                    ' FS (0x1c)
'        SendTmp = SendTmp & "      "                                   ' 상품코드 ~
'        SendTmp = SendTmp & "  "                                       ' 가맹점 사용ID
'        SendTmp = SendTmp & "                              "           ' 가맹점사용필드
'        SendTmp = SendTmp & "                                                                   " ' 여유필드
'        SendTmp = SendTmp & Chr(3) & Chr(13)                           ' FS ,ETX, CR
'         '===========================  여기까지
'    Else
'        '현금영수증취소요청 일 경우
'
'        SendTmp = Chr(2)                         ' STX
'        SendTmp = SendTmp & "bs"                 ' 거래구분
'        SendTmp = SendTmp & 단말기번호           ' "DPT0TEST05" 단말기번호 (Test용)(현재는들어있는 단말기 번호는 테스트 번호 이며 실제 사용할 경우
'                                                 ' KSNET을 통해서 실제단말번호를 받아야 한다.영업팀에 문의 해야함. 승인과 거래시 단말기 번호는 동일해야함)
'        SendTmp = SendTmp & "        "           ' 업체정보
'        SendTmp = SendTmp & "000000"             ' 거래일련번호(전문번호) (일자별로 유니크한 값들을 입력하면되며 사용하지 않을 경우 "000000"을 사용해도 됨)
'        SendTmp = SendTmp & "0"                  ' 거래유형(0으로 사용)
'        SendTmp = SendTmp & "K"                  ' POS Entry Mode (핀패드를 통한 입력이므로 'K' 를 사용한다)
'        SendTmp = SendTmp & Pindata              ' 현금영수증(사용자정보) 번호
'        SendTmp = SendTmp & Chr(28)              ' FS (0x1c)
'        SendTmp = SendTmp & Format(Spread_GetData(sprGrid, 5, 1, True), "000000000")  ' 공급금액(총금액 = 세금 + 봉사료 + 공급금액) 이어야 한다.(반드시 맞춰 준다)
'        SendTmp = SendTmp & "000000000"                                               ' 세금
'        SendTmp = SendTmp & "000000000"                                               ' 봉사료
'        SendTmp = SendTmp & Format(Spread_GetData(sprGrid, 5, 1, True), "000000000")  ' 총금액   (취소일 경우 사용자정보와 금액,현금승인번호,거래일자가 일치 해야 한다)
'        SendTmp = SendTmp & Spread_GetData(sprGrid, 4, 1, True)                       ' 거래구분자(개인이 사용한 경우는 0 , 사용자가 사용하는 경우 1)
'                                                                                      ' 여기서는 개인이라고 가정하고 0을 사용
'        SendTmp = SendTmp & Spread_GetData(sprGrid, 1, 1, True)                       ' 원거래현금승인번호(9자리)
'        SendTmp = SendTmp & "            "                                            ' 원거래포인트승인번호(space 12자리로 사용)
'        SendTmp = SendTmp & Spread_GetData(sprGrid, 2, 1, True)                       ' 거래일자(6자리)- YYMMDD
'        '=========그대로 사용 ====================
'        SendTmp = SendTmp & " "                  ' 포인트 POS Entry Mode
'        SendTmp = SendTmp & "                                     "    ' 포인트 Track II
'        SendTmp = SendTmp & Chr(28)              ' FS (0x1c)
'        SendTmp = SendTmp & "      "             ' 상품코드 ~
'        SendTmp = SendTmp & "  "                 ' 가맹점 사용ID
'        SendTmp = SendTmp & "                              " ' 가맹점사용필드
'        SendTmp = SendTmp & "                                                             " ' 여유필드
'         '                   1234567890123456789012345678901234567890123456789012345678901234567
'        SendTmp = SendTmp & Left(Trim(Left(cboCancel.Text, 2)) & "01", 2)
'        SendTmp = SendTmp & Chr(3) & Chr(13)  ' FS ,ETX, CR
'    End If
'
'
'    ' STX~CR까지의 길이를 전문의 가장 앞에 4바이트 크기의 문자열로 붙입니다.
'    ' 즉, 전문의 총 길이가 475이면  "0475"를 요청전문의 맨 앞에 붙이면 됩니다.
'
'    '전체요청전문의 길이를 구한다.
'    ReqLen = Len(SendTmp)
'    SendTmp = Format(ReqLen, "0000") & SendTmp
'
'    ' 요청전문을 ADSL모듈을 통해 KSNET으로 전송합니다.
'    ' 주의사항 : 3번째 파라미터(이름 : Media)는 전자서명 신용승인/취소 일때만 4입니다.
'    ' 현금영수증, 포인트, 직불 등등에서 Media값은 2입니다.
'    'Rtn = ReqAppr("210.181.28.116", 9531, 2, SendTmp, Len(SendTmp), RecvMsg, 15, 0)   '뒤에 숫자 2개는 는 그대로 사용 하시기 바랍니다(변경하지 마세요)
'    Rtn = ReqAppr(VAN_IP, CInt(VAN_PORT), 2, SendTmp, Len(SendTmp), RecvMsg, 15, 0)    '뒤에 숫자 2개는 는 그대로 사용 하시기 바랍니다(변경하지 마세요)
'
'    If Rtn > 0 Then
'        lblMsg.Caption = "KSNET으로 현금영수증 승인 요청/취소 성공(통신성공)!!" ' 성공했다면 통신은 성공하였기 때문에 승인 성공/거절을 구분하여 처리한다.
'
'        sStatus = MyMid(RecvMsg, 32, 1)            ' 상태
'
'        With sprGrid
'            .Col = 1
'
'            .Row = 1:  .Text = MyMid(RecvMsg, 76, 9) & ""   ' 승인번호(거절코드)
'            .Row = 2:  .Text = MyMid(RecvMsg, 33, 6) & ""   ' 승인일자
'            .Row = 3:  .Text = MyMid(RecvMsg, 39, 4) & ""   ' 승인시간
'
'            If optGubun(0).Value = True Then
'                .Row = 4:  .Text = "0"                      ' 소득, 비소득
'            Else
'                .Row = 4:  .Text = "1"                      ' 소득, 비소득
'            End If
'
'            .Row = 5:  .Text = txtMoney.Value               ' 결제금액
'            .Row = 6:  .Text = txtUserInfo.Text             ' 사용자정보
'            .Row = 7:  .Text = MyMid(RecvMsg, 43, 16) & ""  ' 메시지1
'            .Row = 8:  .Text = MyMid(RecvMsg, 59, 16) & ""  ' 메시지2
'            .Row = 9:  .Text = MyMid(RecvMsg, 75, 1) & ""   ' 소득/비소득 구분
'            .Row = 10: .Text = MyMid(RecvMsg, 85, 20) & ""  ' 국세청1
'            .Row = 11: .Text = MyMid(RecvMsg, 105, 20) & "" ' 국세청2
'        End With
'
'        lblMessage1.Caption = MyMid(RecvMsg, 43, 16) & ""  '메시지1
'        lblMessage2.Caption = MyMid(RecvMsg, 59, 16) & ""  '메시지2
'        DoEvents
'
'        If sStatus = "O" Then
'            lblMsg.Caption = "KSNET으로 승인 요청/취소 성공(통신성공)" ' 성공했다면 통신은 성공하였기 때문에 승인 성공/거절을 구분하여 처리한다.
'
'            If iFlag = "3" Then
'
'                ' 승인된 현금 영수증 자료를 SQL에 저장한다.
'                Call Save_현금영수증승인정보
'
'
'                Select Case Account_Form
'                    Case "접수"
'                        Call Move_승인자료정보(frm접수결제.sprCash)
'
'                    Case "출고"
'                        Call Move_승인자료정보(frm출고결제.sprCash)
'
'                    Case "판매취소2"
'                        Call Move_승인자료정보(frm판매취소.sprCash)
'                End Select
'            Else
'                ' 현금영수증 취소 정보를 SQL에 저장
'                Call Save_현금영수증취소정보("K")
'
'                Select Case Account_Form
'                    Case "접수"
'                        With frm접수결제.sprCash
'                            .Col = 1
'
'                            For i = 1 To .MaxRows
'                                .Row = i: .Text = ""
'                            Next i
'                        End With
'
'                    Case "접수2"   '취소
''                        Call 현금영수증취소_Report(frm현금영수증승인.KS7500i, _
''                                                   Spread_GetData(sprGrid, 1, 1, True), _
''                                                   Spread_GetData(sprGrid, 2, 1, True), _
''                                                   Spread_GetData(sprGrid, 3, 1, True))
'
'                        Call 현금영수증취소_Report(frm현금영수증승인.KS7500i, pnlApprovalNo.Caption, pnlApprovalDay.Caption, pnlApprovalTime.Caption)
'
'                        frm현금영수증승인.Data_Display
'
'                    Case "출고"
'                        With frm출고결제.sprCash
'                            .Col = 1
'
'                            For i = 1 To .MaxRows
'                                .Row = i: .Text = ""
'                            Next i
'                        End With
'
'                    Case "판매취소"
'                        With frm판매취소결제.sprCash
'                            .Col = 1
'
'                            For i = 1 To .MaxRows
'                                .Row = i: .Text = ""
'                            Next i
'                        End With
'
'                        ' 2014-10-24일 추가.. 현금영수증 판매 취소한 경우 취소 영수증이 안나와서
'                        결제취소여부 = True
'
'                    Case "판매취소2"
'                        With frm판매취소.sprCash
'                            .Col = 1
'
'                            For i = 1 To .MaxRows
'                                .Row = i: .Text = ""
'                            Next i
'                        End With
'                End Select
'            End If
'
'            Unload Me
'        Else
'            lblMsg.Caption = "승인 요청/취소 거절 되었습니다! 거절 메세지를 확인하세요"
'        End If
'    Else
'        ' 통신에러가 발생한 경우 이므로 에러 처리를 한다.
'        lblMsg.Caption = "승인 요청/취소 실패(통신에러가 발생 하였습니다!! 다시 거래 하시기 바랍니다"
'
'        lblMessage1.Caption = RecvMsg
'    End If
'End Sub


Private Sub Save_현금영수증승인정보()
' sprGrid에 등록되어 있는 현금영수증 승인 정보를 저장한다.
    Dim Query       As String
    

    Query = "INSERT INTO TB_현금영수증 ("
    Query = Query & "  승인번호"   ' 1
    Query = Query & ", 승인일자"   ' 2
    Query = Query & ", 승인시간"   ' 3
    Query = Query & ", 거래유형"   ' 4 <- 실제는 '거래자구분'
    Query = Query & ", 입력방법"   ' 5
    Query = Query & ", 사용자정보" ' 6
    Query = Query & ", 총금액"     ' 7
    Query = Query & ", 메시지1"    ' 8
    Query = Query & ", 메시지2"    ' 9
    Query = Query & ", 소득구분"   '10
    Query = Query & ", 국세청1"    '11
    Query = Query & ", 국세청2"    '12
    Query = Query & ", 가맹점코드" '13
    Query = Query & ", 지사코드"   '14
    Query = Query & ", 고객코드"   '15
    Query = Query & ", 접수번호"   '16
    Query = Query & ", 단말기번호" '17
    Query = Query & ", 거래구분"   '18
    Query = Query & ", 상태"       '19
    Query = Query & ") VALUES ("
    Query = Query & "  '" & Spread_GetData(sprGrid, 1, 1, True) & "'"  ' 1 승인번호
    Query = Query & ", '" & Spread_GetData(sprGrid, 2, 1, True) & "'"  ' 2 승인일자
    Query = Query & ", '" & Spread_GetData(sprGrid, 3, 1, True) & "'"  ' 3 승인시간
    Query = Query & ", '" & Spread_GetData(sprGrid, 4, 1, True) & "'"  ' 4 거래유형 <- 실제는 '거래자구분'
    Query = Query & ", 'K'"                                            ' 5 입력방법
    Query = Query & ", '" & Spread_GetData(sprGrid, 6, 1, True) & "'"  ' 6 사용자정보
    Query = Query & ", '" & Spread_GetData(sprGrid, 5, 1, True) & "'"  ' 7 총금액
    Query = Query & ", '" & Spread_GetData(sprGrid, 7, 1, True) & "'"  ' 8 메시지1
    Query = Query & ", '" & Spread_GetData(sprGrid, 8, 1, True) & "'"  ' 9 메시지2
    Query = Query & ", '" & Spread_GetData(sprGrid, 9, 1, True) & "'"  '10 소득구분
    Query = Query & ", '" & Spread_GetData(sprGrid, 10, 1, True) & "'" '11 국세청1
    Query = Query & ", '" & Spread_GetData(sprGrid, 11, 1, True) & "'" '12 국세청2
    Query = Query & ", '" & 가맹점정보.가맹점코드 & "'"                '13 가맹점코드
    Query = Query & ", '" & 가맹점정보.지사코드 & "'"                  '13 지사코드
    Query = Query & ", '" & pnlCustomCode.Caption & "'"                '13 고객코드
    Query = Query & ",  " & pnlNum.Caption & ""                        '14 접수번호
    Query = Query & ", '" & 단말기번호 & "'"                           '15 단말기번호
    Query = Query & ", 'bq'"                                           '16 거래구분
    Query = Query & ", 'O'"                                            '17 상태
    Query = Query & ")"
    ADOCon.Execute Query

End Sub

Private Sub Save_현금영수증취소정보(sKeyMode As String)
' sprGrid에 등록되어 있는 현금영수증 승인 정보를 저장한다.
    Dim Query       As String
    
    'TB_현금영수증 - 취소
    Query = "UPDATE TB_현금영수증 SET "
    Query = Query & "  거래유형     = '" & Spread_GetData(sprGrid, 4, 1, True) & "'"  '<- 실제는 '거래자구분'
    Query = Query & ", 입력방법     = '" & sKeyMode & "'"
    Query = Query & ", 사용자정보   = '" & Spread_GetData(sprGrid, 6, 1, True) & "'"  '
    Query = Query & ", 총금액       = '" & Spread_GetData(sprGrid, 5, 1, True) & "'"  '
    Query = Query & ", 메시지1      = '" & Spread_GetData(sprGrid, 7, 1, True) & "'"  '
    Query = Query & ", 메시지2      = '" & Spread_GetData(sprGrid, 8, 1, True) & "'"  '
    Query = Query & ", 소득구분     = '" & Spread_GetData(sprGrid, 9, 1, True) & "'"  '
    Query = Query & ", 국세청1      = '" & Spread_GetData(sprGrid, 10, 1, True) & "'" '
    Query = Query & ", 국세청2      = '" & Spread_GetData(sprGrid, 11, 1, True) & "'" '
    Query = Query & ", 취소승인번호 = '" & Spread_GetData(sprGrid, 1, 1, True) & "'"  '
    Query = Query & ", 취소일자     = '" & Spread_GetData(sprGrid, 2, 1, True) & "'"  '
    Query = Query & ", 취소시간     = '" & Spread_GetData(sprGrid, 3, 1, True) & "'"  '
    Query = Query & ", 취소사유     = '" & Left(Trim(Left(cboCancel.Text, 2)) & "01", 2) & "'"  '
    Query = Query & ", 본사전송여부 = 'N'"  '
    Query = Query & " WHERE 승인번호 = '" & pnlApprovalNo.Caption & "'"
    Query = Query & "   AND 승인일자 = '" & pnlApprovalDay.Caption & "'"
    Query = Query & "   AND 승인시간 = '" & pnlApprovalTime.Caption & "'"
    ADOCon.Execute Query


End Sub

Private Sub Move_승인자료정보(MyObj As Object)
    With MyObj
        .Col = 1
    
        .Row = 1:  .Text = Spread_GetData(sprGrid, 1, 1, True)   '승인번호
        .Row = 2:  .Text = Spread_GetData(sprGrid, 2, 1, True)   '승인일자
        .Row = 3:  .Text = Spread_GetData(sprGrid, 3, 1, True)   '승인시간
        .Row = 4:  .Text = Spread_GetData(sprGrid, 4, 1, True)   '거래유형 '<- 실제는 '거래자구분'
        .Row = 5:  .Text = Spread_GetData(sprGrid, 5, 1, True)   '총금액
        .Row = 6:  .Text = Spread_GetData(sprGrid, 6, 1, True)   '사용자정보
        .Row = 7:  .Text = Spread_GetData(sprGrid, 7, 1, True)   '메시지1
        .Row = 8:  .Text = Spread_GetData(sprGrid, 8, 1, True)   '메시지2
        .Row = 9:  .Text = Spread_GetData(sprGrid, 9, 1, True)   '소득구분
        .Row = 10: .Text = Spread_GetData(sprGrid, 10, 1, True)  '국세청1
        .Row = 11: .Text = Spread_GetData(sprGrid, 11, 1, True)  '국세청2
    End With

End Sub

Private Sub optGubun_Click(Index As Integer, Value As Integer)
    Call 현금영수증승인요청_Rtn(CStr(iFlag))
End Sub


Public Sub ReceiveMsg(msg As String)
    Dim sHelpDesk   As String


    Dim TempString As String
    TempString = msg
    Debug.Print "전문코드 : "
    Debug.Print MyMid(TempString, 1, 2)     ' 전문코드
    Debug.Print "응답코드 : "
    Debug.Print MyMid(TempString, 3, 4)     ' 응답코드
    
    Debug.Print "TID : "
    Debug.Print MyMid(TempString, 7, 8)     ' TID
    Debug.Print "WCC : "
    Debug.Print MyMid(TempString, 15, 1)    ' WCC
    Debug.Print "카드번호 : "
    Debug.Print MyMid(TempString, 16, 40)   ' 카드번호
    Debug.Print "할부/현금/권종 : "
    Debug.Print MyMid(TempString, 56, 2)    ' 할부/현금/권종
    Debug.Print "금액 : "
    Debug.Print MyMid(TempString, 58, 8)    ' 금액
    Debug.Print "봉사료 : "
    Debug.Print MyMid(TempString, 66, 8)    ' 봉사료
    Debug.Print "VAT : "
    Debug.Print MyMid(TempString, 74, 8)    ' VAT
    Debug.Print "승인번호 : "
    Debug.Print MyMid(TempString, 82, 12)   ' 승인번호
    Debug.Print "승인일시 : "
    Debug.Print MyMid(TempString, 94, 12)   ' 승인일시
    Debug.Print "발급사코드 : "
    Debug.Print MyMid(TempString, 106, 3)   ' 발급사코드
    Debug.Print "카드사명 : "
    Debug.Print MyMid(TempString, 109, 20)  ' 카드사명
    Debug.Print "가맹점코드 : "
    Debug.Print MyMid(TempString, 129, 12)  ' 가맹점코드
    Debug.Print "매입사코드 : "
    Debug.Print MyMid(TempString, 141, 3)   ' 매입사코드
    Debug.Print "매입사명 : "
    Debug.Print MyMid(TempString, 144, 20)  ' 매입사명
    Debug.Print "POS거래번호 : "
    Debug.Print MyMid(TempString, 164, 20)  ' POS거래번호
    Debug.Print "DSC 구분 : "
    Debug.Print MyMid(TempString, 184, 1)   ' DSC 구분
    Debug.Print "전자서명 : "
    Debug.Print MyMid(TempString, 185, 1)   ' 전자서명
    
    If MyMid(TempString, 3, 4) <> "0000" Then Exit Sub

    On Error GoTo ERR_RTN
    
    lblMsg.Caption = "현금영수증 승인 요청/취소 성공(통신성공)!!" ' 성공했다면 통신은 성공하였기 때문에 승인 성공/거절을 구분하여 처리한다.
    
    sStatus = MyMid(TempString, 82, 12)
    
    With sprGrid
        .Col = 1
            
       
        .Row = 1:  .Text = MyMid(TempString, 82, 12) & ""   '승인번호(거절코드)
        .Row = 2:  .Text = MyMid(TempString, 94, 6) & ""    '승인일자
        .Row = 3:  .Text = MyMid(TempString, 100, 6) & ""   '승인시간
        
        .Row = 4:  .Text = MyMid(TempString, 57, 1) & ""   ' 현금 영수증 소비자 구분 (POS에서 전송된 데이터 그래도 리턴(00.개인소득공제, 01.사업자)
        
        .Row = 5:  .Text = Val(MyMid(TempString, 58, 8))     '결제금액
        .Row = 6:  .Text = MyMid(TempString, 16, 16)        '카드번호 (전체 카드번호 중 1-12자리를 ***** 표시하여 전달
        .Row = 7:  .Text = ""  ' 메시지1
        .Row = 8:  .Text = "OK"  ' 메시지2
        .Row = 9:  .Text = ""   ' 소득/비소득 구분
        .Row = 10: .Text = ""  ' 국세청1
        .Row = 11: .Text = "" ' 국세청2
    End With
    
    lblMessage1.Caption = "OK"  '메시지1
    lblMessage2.Caption = ""  '메시지2
    DoEvents
    
     
    ' 현금 영수증 승인 번호는 9자리로 확인됨
    If Len(Trim(sStatus)) > 0 Then
        lblMsg.Caption = "KICC으로 승인 요청/취소 성공(통신성공)" ' 성공했다면 통신은 성공하였기 때문에 승인 성공/거절을 구분하여 처리한다.
        
        If iFlag = "3" Then
        
            ' 승인된 현금 영수증 자료를 SQL에 저장한다.
            Call Save_현금영수증승인정보

            
            Select Case Account_Form
                Case "접수"
                    Call Move_승인자료정보(frm접수결제.sprCash)
                
                Case "출고"
                    Call Move_승인자료정보(frm출고결제.sprCash)
                        
                Case "판매취소2"
                    Call Move_승인자료정보(frm판매취소.sprCash)
            End Select
        Else
            ' 현금영수증 취소 정보를 SQL에 저장
            Call Save_현금영수증취소정보("K")
            
            Select Case Account_Form
                Case "접수"
                    With frm접수결제.sprCash
                        .Col = 1
                        For i = 1 To .MaxRows
                            .Row = i: .Text = ""
                        Next i
                    End With
                
                Case "접수2"   '취소
                    frm현금영수증승인.Data_Display
                
                Case "출고"
                    With frm출고결제.sprCash
                        .Col = 1
                        
                        For i = 1 To .MaxRows
                            .Row = i: .Text = ""
                        Next i
                    End With
                
                Case "판매취소"
                    With frm판매취소결제.sprCash
                        .Col = 1
                        
                        For i = 1 To .MaxRows
                            .Row = i: .Text = ""
                        Next i
                    End With
            
                    ' 2014-10-24일 추가.. 현금영수증 판매 취소한 경우 취소 영수증이 안나와서
                    결제취소여부 = True
            
                Case "판매취소2"
                    With frm판매취소.sprCash
                        .Col = 1
                        
                        For i = 1 To .MaxRows
                            .Row = i: .Text = ""
                        Next i
                    End With
            End Select
        End If
        Unload Me
    End If

ERR_RTN:
    lblMsg.Caption = Err.description
End Sub
