VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frmKSNET2 
   BorderStyle     =   1  '단일 고정
   Caption         =   "카드결제"
   ClientHeight    =   7260
   ClientLeft      =   10110
   ClientTop       =   3600
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   5145
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7260
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   12806
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frmKSNET2.frx":0000
      Begin XtremeSuiteControls.TabControl TabControl 
         Height          =   5655
         Left            =   15
         TabIndex        =   1
         Top             =   960
         Width           =   5115
         _Version        =   851970
         _ExtentX        =   9022
         _ExtentY        =   9975
         _StockProps     =   68
         Appearance      =   2
         Color           =   16
         ItemCount       =   2
         Item(0).Caption =   "카드결제"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "TabControlPage(0)"
         Item(1).Caption =   "승인/취소 정보"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "TabControlPage(1)"
         Begin XtremeSuiteControls.TabControlPage TabControlPage 
            Height          =   5310
            Index           =   1
            Left            =   -69970
            TabIndex        =   2
            Top             =   315
            Visible         =   0   'False
            Width           =   4950
            _Version        =   851970
            _ExtentX        =   8731
            _ExtentY        =   9366
            _StockProps     =   1
            Page            =   1
            Begin FPSpreadADO.fpSpread sprGrid 
               Height          =   4350
               Left            =   105
               TabIndex        =   3
               Top             =   1335
               Width           =   4740
               _Version        =   524288
               _ExtentX        =   8361
               _ExtentY        =   7673
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
               MaxRows         =   15
               RowHeaderDisplay=   0
               ScrollBars      =   0
               SpreadDesigner  =   "frmKSNET2.frx":0072
               HighlightHeaders=   1
               HighlightStyle  =   1
            End
            Begin Threed.SSPanel pnlNum 
               Height          =   315
               Left            =   975
               TabIndex        =   4
               Top             =   420
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
               TabIndex        =   5
               Top             =   75
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
               TabIndex        =   6
               Top             =   75
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
               TabIndex        =   7
               Top             =   420
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
            Begin Threed.SSPanel pnlApprovalTime 
               Height          =   315
               Left            =   3630
               TabIndex        =   8
               Top             =   765
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
               Caption         =   "고객코드:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   5
               Left            =   60
               TabIndex        =   13
               Top             =   135
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
               TabIndex        =   12
               Top             =   495
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
               TabIndex        =   11
               Top             =   150
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
               TabIndex        =   10
               Top             =   495
               Width           =   1080
            End
            Begin VB.Label Label1 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "승인시간:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   9
               Left            =   2520
               TabIndex        =   9
               Top             =   840
               Width           =   1080
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage 
            Height          =   5310
            Index           =   0
            Left            =   30
            TabIndex        =   14
            Top             =   315
            Width           =   5055
            _Version        =   851970
            _ExtentX        =   8916
            _ExtentY        =   9366
            _StockProps     =   1
            BackColor       =   16761087
            Page            =   0
            Begin VB.ComboBox cboGubun 
               Enabled         =   0   'False
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
               ItemData        =   "frmKSNET2.frx":0686
               Left            =   1200
               List            =   "frmKSNET2.frx":0688
               Style           =   2  '드롭다운 목록
               TabIndex        =   17
               Top             =   150
               Width           =   2760
            End
            Begin VB.ComboBox cboMonth 
               BackColor       =   &H00C0FFFF&
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
               TabIndex        =   16
               Top             =   555
               Width           =   2760
            End
            Begin VB.TextBox txtCardNum 
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
               Height          =   675
               IMEMode         =   3  '사용 못함
               Left            =   1200
               MultiLine       =   -1  'True
               TabIndex        =   15
               Top             =   1425
               Width           =   3630
            End
            Begin Threed.SSPanel SSPanel1 
               Height          =   1785
               Left            =   1200
               TabIndex        =   18
               Top             =   2160
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
                  Index           =   0
                  Left            =   30
                  TabIndex        =   19
                  Top             =   30
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
                  Picture         =   "frmKSNET2.frx":068A
               End
            End
            Begin CSTextLibCtl.silgEdit txtMoney 
               Height          =   405
               Left            =   1200
               TabIndex        =   20
               Top             =   960
               Width           =   3630
               _Version        =   262145
               _ExtentX        =   6403
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
            Begin VB.Label Label1 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "사인패드:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   3
               Left            =   60
               TabIndex        =   28
               Top             =   2265
               Width           =   1080
            End
            Begin VB.Label Label1 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "거래구분:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   0
               Left            =   60
               TabIndex        =   27
               Top             =   240
               Width           =   1080
            End
            Begin VB.Label Label1 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "카드번호:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   4
               Left            =   60
               TabIndex        =   26
               Top             =   1485
               Width           =   1080
            End
            Begin VB.Label Label1 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "할부기간:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   1
               Left            =   60
               TabIndex        =   25
               Top             =   630
               Width           =   1080
            End
            Begin VB.Label Label1 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "결제금액:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   2
               Left            =   60
               TabIndex        =   24
               Top             =   1050
               Width           =   1080
            End
            Begin VB.Label lblMessage1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
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
               Left            =   1260
               TabIndex        =   23
               Top             =   4020
               Width           =   105
            End
            Begin VB.Label lblMessage2 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
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
               Left            =   1260
               TabIndex        =   22
               Top             =   4320
               Width           =   105
            End
            Begin VB.Label Label1 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "메 시 지:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   10
               Left            =   60
               TabIndex        =   21
               Top             =   4020
               Width           =   1080
            End
         End
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   930
         Left            =   15
         TabIndex        =   29
         Top             =   15
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   1640
         _Version        =   262144
         BackColor       =   4210752
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Label lblMsg 
            BackStyle       =   0  '투명
            Caption         =   "#"
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
            TabIndex        =   31
            Top             =   135
            Width           =   4275
         End
         Begin VB.Image Image 
            Height          =   360
            Index           =   1
            Left            =   105
            Picture         =   "frmKSNET2.frx":109C
            Top             =   120
            Width           =   360
         End
         Begin VB.Label lblErrMsg 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "#"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   4785
            TabIndex        =   30
            Top             =   660
            Width           =   90
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   615
         Left            =   15
         TabIndex        =   32
         Top             =   6630
         Width           =   5115
         _ExtentX        =   9022
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
            TabIndex        =   33
            Top             =   45
            Width           =   1260
            _Version        =   851970
            _ExtentX        =   2222
            _ExtentY        =   847
            _StockProps     =   79
            Caption         =   " 취소(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frmKSNET2.frx":211E
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   495
            Index           =   1
            Left            =   60
            TabIndex        =   34
            Top             =   60
            Width           =   1380
            _Version        =   851970
            _ExtentX        =   2434
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   " 재전송"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frmKSNET2.frx":2B30
         End
      End
   End
End
Attribute VB_Name = "frmKSNET2"
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

Dim m_접수일자      As String

Dim KS7500_CommPort  As String
Dim KS7500_BaudRate  As String
Dim KS7500_Delay     As Integer

Dim SignPad_CommPort As String
Dim SignPad_BaudRate As String
Dim ApproveValue As String
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

Public Sub 신용카드승인요청_IC_Start()

    cmdBtn(2).Enabled = True
    SSPanel2.Caption = "   단말기의 종료 버튼으로 취소"
    
    
    Dim sD As String
    Dim sE As String
    Dim ReturnValue As Long
    Dim Cancel_Prev_B As Boolean
    
    If Right(Format(Now, "yyyyMMdd"), 6) <> pnlApprovalDay.Caption Then
        Cancel_Prev_B = True
    End If
        
    sD = SetMessage(IIf(iFlag = "1", Credit_Approve, IIf(Cancel_Prev_B, Credit_Cancel_Prev_Day, Credit_Cancel_Today)), IIf(iFlag = "1", CStr(txtMoney.Value), Replace(Spread_GetData(sprGrid, 5, 1, False), ",", "")), Left(cboMonth.Text, 2), IIf(iFlag = "1", "", Spread_GetData(sprGrid, 1, 1, True)), IIf(iFlag = "1", "", IIf(Cancel_Prev_B, Spread_GetData(sprGrid, 2, 1, True), "")), ApproveValue)
    
    Call frmKicc.Card_Approve(sD, Me.Name)

End Sub

'-------------------------------------------------------------------------------
' 함수명 : 신용카드승인요청_Rtn
'
'
'-------------------------------------------------------------------------------
Public Sub 신용카드승인요청_Rtn(Gbn As String)
    Dim Rtn As Integer
    Dim sRtn    As String
    
    cmdBtn(0).Visible = False
    txtCardNum.Text = ""
    txtCardNum.Tag = ""
    
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
    
    iFlag = Gbn ' 신용승인요청임을 나타내는 플래그 셋팅
    
    cboGubun.ListIndex = IIf(iFlag = 1, 0, 1)
    TabControlPage(0).BackColor = IIf(iFlag = 1, &H8000000F, &HFFC0FF) '&H00FFC0FF&
    TabControl.SelectedItem = 0

'    If 가맹점정보.CAT단말기종류 = "KICC" Then
        cmdBtn(0).Visible = True
        lblMsg.Caption = "반드시 승인 시작 버튼을 누른 후  IC 카드를 삽입 하여 주십시요."
'        Exit Sub
'    Else
'        MsgBox "지원하지 않는 단말기 입니다." & vbCrLf & "단말기 설정을 확인하여 주십시요"
'    End If
    
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 0
            cmdBtn(0).Enabled = False
            lblMessage1.Caption = ""
            lblMessage2.Caption = ""
            Call 신용카드승인요청_IC_Start
            Exit Sub
        Case 1
            Call frmKicc.Card_Request("F1", Me.Name)
            
        Case 2:
            Dim sD As String
            Dim sE As String
            sD = "TM"
            'KiccPosOCX.ReqCmd &HFD, 0, 0, sD, sE
            Call frmKicc.Card_Approve(sD, Me.Name)
            
            Unload Me
        'Case 1
            'Call ReceiveMsg(Text1.Text)
    End Select
End Sub

Private Sub Form_Activate()
    If Get_일일마감여부(Format(Date, "YYYY-MM-DD")) = True Then
        m_접수일자 = Format(DateAdd("d", 1, Date), "YYYY-MM-DD")
    Else
        m_접수일자 = Format(Date, "YYYY-MM-DD")
    End If
    ApproveValue = Format(Now, "YYYYMMDDhhmmss")
End Sub

Private Sub Form_Load()
    Me.Top = ((frmMain.Height - Me.Height) / 2) + frmMain.Top
    Me.Left = ((frmMain.Width - Me.Width) / 2) + frmMain.Left
    

    With sprGrid
        .RowHeight(-1) = 14
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeNormal
    End With
    
    With cboGubun
        .Clear
        .AddItem "신용승인"
        .AddItem "신용취소"
        
        .ListIndex = 0
    End With
    
    With cboMonth
        .Clear
        .AddItem "00 일시불"
        
        For i = 2 To 36
            .AddItem Format(i, "00") & " 개월"
        Next i
        
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
    
    KS7500_Delay = 1
    If IsNumeric(GetIniStr("VAN", "KS7500_Delay", "", iniFile)) = True Then
        KS7500_Delay = Val(GetIniStr("VAN", "KS7500_Delay", "", iniFile))
    Else
        Call SetIniStr("VAN", "KS7500_Delay", "1", iniFile)    '
    End If
    
    SignPad_CommPort = GetIniStr("VAN", "SignPad_CommPort", "", iniFile) '
    SignPad_BaudRate = GetIniStr("VAN", "SignPad_BaudRate", "", iniFile) '
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'KiccPosOCX.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'KiccPosOCX.Close
End Sub

Private Sub 매출취소_Rtn()
    '-------------------------------------------------------------------------------------------------
    ' TB_매출 - 결제취소를 하였으므로 매출에서
    '-------------------------------------------------------------------------------------------------
    Dim iSEQ As Long

    Query = "SELECT ISNULL(MAX(일련번호),0) + 1"
    Query = Query & " FROM TB_매출"
    Query = Query & " WHERE 고객코드 = '" & pnlCustomCode.Caption & "'"
    Query = Query & "   AND 접수번호 =  " & pnlNum.Caption              '기본은 0 이고, 판매취소에서 결제할때는 접수번호 가져옴...
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    iSEQ = ADORs(0)

    ADORs.Close
    Set ADORs = Nothing

    '-----------------------------------------------------------
    Query = "SELECT * FROM TB_매출"
    Query = Query & " WHERE 고객코드 = '" & pnlCustomCode.Caption & "'"
    Query = Query & "   AND 접수번호 =  " & pnlNum.Caption             '기본은 0 이고, 판매취소에서 결제할때는 접수번호 가져옴...
    Query = Query & "   AND 일련번호 =  " & iSEQ
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenDynamic, adLockOptimistic

    If ADORs.EOF Then ADORs.AddNew

    ADORs!지사코드 = 가맹점정보.지사코드                    '
    ADORs!가맹점코드 = 가맹점정보.가맹점코드                '

    ADORs!고객코드 = pnlCustomCode.Caption & ""             ' 1
    ADORs!접수번호 = pnlNum.Caption & ""                    ' 2
    ADORs!일련번호 = iSEQ                                   ' 3
    ADORs!매출일자 = Format(Date, "YYYY-MM-DD") & ""        ' 4
    ADORs!매출시간 = Format(Now, "hh:mm:ss")                ' 5
    ADORs!적요 = "[신용카드 승인취소]"                      ' 6
    ADORs!접수금액 = 0                                      ' 7
    ADORs!입금합계 = txtMoney.Value * -1                    ' 8
    ADORs!현금입금 = 0                                      ' 9
    ADORs!카드입금 = txtMoney.Value * -1                    '10
    ADORs!쿠폰입금 = 0                                      '
    ADORs!쿠폰번호 = ""                                     '
    ADORs!사용마일리지 = 0                                  '
    ADORs!세트할인 = 0                                      '
    ADORs!에누리 = 0                                        '
    ADORs!접수수량 = 0                                      '12
    ADORs!반품수량 = 0                                      '13
    ADORs!발생마일리지 = 0                                  '
    ADORs!누적마일리지 = 0                                  '
    ADORs!사용가능마일리지 = 0                              '
    ADORs!이전미수금 = 0                                    '
    ADORs!본사전송여부 = ""                                 '

    ADORs.Update

    ADORs.Close
    Set ADORs = Nothing
End Sub

Private Sub sprGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    If Row <= 0 Then Exit Sub
        
    If Col = 1 Then
        Rtn = MsgBox("카드승인 취소를 하시겠습니까?", vbQuestion + vbYesNo + vbDefaultButton2, "승인취소")
    
        If Rtn = vbNo Then Exit Sub

        sprGrid.Row = Row
        sprGrid.Col = 6: txtMoney.Value = sprGrid.Value                     '
        
        sprGrid.Col = 5
        If sprGrid.Text = "일시불" Then
            cboMonth.ListIndex = 0
        Else
            For i = 1 To cboMonth.ListCount - 1
                If Left(cboMonth.List(i), 2) = Left(sprGrid.Text, 2) Then
                    cboMonth.ListIndex = i
                    
                    Exit For
                End If
            Next i
        End If
        
        Call 신용카드승인요청_Rtn("2")
    End If
End Sub

Private Sub txtMoney_Change()
    If txtMoney.ReadOnly = False Then
        If txtMoney.Tag = "" Then Exit Sub
        
        '잔액보다 많은 금액을 카드결제할수 없음...
        If txtMoney.Value > CLng(txtMoney.Tag) Then
            txtMoney.Value = CLng(txtMoney.Tag)
        End If
    End If
End Sub

Public Sub ReceiveMsg(msg As String)
    Dim TempString As String
    TempString = msg
    
    If MyMid(TempString, 3, 4) <> "0000" Then
        lblMessage1.Caption = "오류코드 : " & MyMid(TempString, 3, 4)
        lblMessage2.Caption = MyMid(TempString, 16, 40)
        cmdBtn(0).Enabled = True
        Dim sD As String
        Dim sE As String
        sD = "TM"
        Call frmKicc.Card_Approve(sD, Me.Name)
        Exit Sub
    End If
    
    If Trim(ApproveValue) <> Trim(MyMid(TempString, 164, 20)) Then Exit Sub
    
    With sprGrid
        .Col = 1
        .Row = 1:  .Text = MyMid(TempString, 82, 12) & ""   '승인번호(거절코드)
        .Row = 2:  .Text = MyMid(TempString, 94, 6) & ""    '승인일자
        .Row = 3:  .Text = MyMid(TempString, 100, 6) & ""   '승인시간
        .Row = 4:  .Text = MyMid(TempString, 56, 2) & ""    '할부기간
        .Row = 5:  .Text = Val(MyMid(TempString, 58, 8))    '결제금액
        .Row = 6:  .Text = MyMid(TempString, 106, 3) & ""   '발급사코드
        .Row = 7:  .Text = MyMid(TempString, 109, 20) & ""  '카드종류명
        .Row = 8:  .Text = MyMid(TempString, 141, 3) & ""   '매입사코드
        .Row = 9:  .Text = MyMid(TempString, 144, 20) & ""  '매입사명
        .Row = 10: .Text = MyMid(TempString, 16, 16)        '카드번호 (전체 카드번호 중 1-12자리를 ***** 표시하여 전달
        .Row = 11: .Text = "" 'MyMid(TempString, 129, 16) & ""  '메시지1
        .Row = 12: .Text = "OK" & ""  '메시지2
    End With

    sStatus = MyMid(TempString, 82, 12)
    
    ' 신용 승인 번호는 8자리로 확인됨
    If Len(Trim(sStatus)) > 0 Then
    
        lblMsg.Caption = "KICC으로 승인 요청/취소 성공(통신성공)" ' 성공했다면 통신은 성공하였기 때문에 승인 성공/거절을 구분하여 처리한다.
        
        If iFlag = "1" Then
        
            Query = "INSERT INTO TB_신용카드승인 ("
            Query = Query & "  승인번호"   ' 1
            Query = Query & ", 승인일자"   ' 2
            Query = Query & ", 승인시간"   ' 3
            Query = Query & ", 할부기간"   ' 4
            Query = Query & ", 결제금액"   ' 5
            Query = Query & ", 발급사코드" ' 6
            Query = Query & ", 카드종류명" ' 7
            Query = Query & ", 매입사코드" ' 8
            Query = Query & ", 매입사명"   ' 9
            Query = Query & ", 카드번호"   '10
            Query = Query & ", 메시지1"    '11
            Query = Query & ", 메시지2"    '12
            Query = Query & ", 가맹점코드" '13
            Query = Query & ", 지사코드"   '14
            Query = Query & ", 고객코드"   '15
            Query = Query & ", 접수번호"   '16
            Query = Query & ", 단말기번호" '17
            Query = Query & ", 거래구분"   '18
            Query = Query & ", 상태"       '19
            Query = Query & ", 기타메모"   '20
            Query = Query & ") VALUES ("
            Query = Query & "  '" & Spread_GetData(sprGrid, 1, 1, True) & "'"            ' 1 승인번호
            Query = Query & ", '" & Spread_GetData(sprGrid, 2, 1, True) & "'"            ' 2 승인일자
            Query = Query & ", '" & Spread_GetData(sprGrid, 3, 1, True) & "'"            ' 3 승인시간
            Query = Query & ", '" & Spread_GetData(sprGrid, 4, 1, True) & "'"            ' 4 할부기간
            Query = Query & ", '" & Spread_GetData(sprGrid, 5, 1, True) & "'"            ' 5 결제금액
            Query = Query & ", '" & Spread_GetData(sprGrid, 6, 1, True) & "'"            ' 6 발급사코드
            Query = Query & ", '" & Spread_GetData(sprGrid, 7, 1, True) & "'"            ' 7 발급사명
            Query = Query & ", '" & Spread_GetData(sprGrid, 8, 1, True) & "'"            ' 8 매입사코드
            Query = Query & ", '" & Spread_GetData(sprGrid, 9, 1, True) & "'"            ' 9 매입사명
            Query = Query & ", '" & Spread_GetData(sprGrid, 10, 1, True) & "'"           '10 카드번호
            Query = Query & ", '" & Spread_GetData(sprGrid, 11, 1, True) & "'"           '11 메시지1
            Query = Query & ", '" & Spread_GetData(sprGrid, 12, 1, True) & "'"           '12 메시지2
            Query = Query & ", '" & 가맹점정보.가맹점코드 & "'"                          '13 가맹점코드
            Query = Query & ", '" & 가맹점정보.지사코드 & "'"                            '14 지사코드
            Query = Query & ", '" & pnlCustomCode.Caption & "'"                          '15 고객코드
            Query = Query & ",  " & pnlNum.Caption & ""                                  '16 접수번호
            Query = Query & ", '" & 단말기번호 & "'"                                     '17 단말기번호
            Query = Query & ", 'NA'"                                                     '18 거래구분
            Query = Query & ", 'O'"                                                      '19 상태
            Query = Query & ", '" & Account_Form & "'"                                   '20 승인폼
            Query = Query & ")"
            ADOCon.Execute Query
        Else
            'Query = "UPDATE TB_신용카드승인 SET 메시지2 = '" & Spread_GetData(sprGrid, 12, 1, True) & "'"
            'Query = Query & " WHERE 승인번호 = '" & Spread_GetData(sprGrid, 1, 1, True) & "'"
            'Query = Query & "   AND 승인일자 = '" & Spread_GetData(sprGrid, 2, 1, True) & "'"
            'Query = Query & "   AND 승인시간 = '" & Spread_GetData(sprGrid, 3, 1, True) & "'"
            
            If Get_일일마감여부(Format(Date, "YYYY-MM-DD")) = True Then
                
                m_접수일자 = Format(DateAdd("d", 1, Date), "YYYY-MM-DD")
            Else
                m_접수일자 = Format(Date, "YYYY-MM-DD")
            End If
                        
                        
            
            Query = "UPDATE TB_신용카드승인 SET 메시지2 = '취소' "
            Query = Query & ", 취소일자 = '" & Format(Now, "yyyy-MM-dd hh:mm:ss") & " " & m_접수일자 & "' "
            Query = Query & ", 기타메모 = isnull(기타메모,'')  + '" & Account_Form & "' "
            Query = Query & ", 본사전송여부 = 'N' "
            Query = Query & " WHERE 승인번호 = '" & pnlApprovalNo.Caption & "'"
            Query = Query & "   AND 승인일자 = '" & pnlApprovalDay.Caption & "'"
            Query = Query & "   AND 승인시간 = '" & pnlApprovalTime.Caption & "'"
            ADOCon.Execute Query
        End If
        
        Dim 카드결제금액 As Long
        
        카드결제금액 = 0
        
        Select Case Account_Form
            Case "접수"
                    If iFlag = "1" Then
                        '승인
                        With frm접수결제.sprCard
                            .MaxRows = .MaxRows + 1
                            .Row = .MaxRows
                        
                            .Col = 2:  .Text = Spread_GetData(sprGrid, 1, 1, True)   '승인번호
                            .Col = 3:  .Text = Spread_GetData(sprGrid, 2, 1, True)   '승인일자
                            .Col = 4:  .Text = Spread_GetData(sprGrid, 3, 1, True)   '승인시간
                            .Col = 5:  .Text = Spread_GetData(sprGrid, 4, 1, True)   '할부기간
                            .Col = 6:  .Text = Spread_GetData(sprGrid, 5, 1, True)   '결제금액
                            .Col = 7:  .Text = Spread_GetData(sprGrid, 6, 1, True)   '발급사코드
                            .Col = 8:  .Text = Spread_GetData(sprGrid, 7, 1, True)   '발급사명
                            .Col = 9:  .Text = Spread_GetData(sprGrid, 8, 1, True)   '매입사코드
                            .Col = 10: .Text = Spread_GetData(sprGrid, 9, 1, True)   '매입사명
                            .Col = 11: .Text = Spread_GetData(sprGrid, 10, 1, True)  '카드번호
                            .Col = 12: .Text = Spread_GetData(sprGrid, 11, 1, True)  '메시지1
                            .Col = 13: .Text = Spread_GetData(sprGrid, 12, 1, True)  '메시지2
                            
                            For i = 1 To .MaxRows
                                .Row = i
                                .Col = 6: 카드결제금액 = 카드결제금액 + .Value
                            Next i
                        End With
                        
                        frm접수결제.txtCard.Value = 카드결제금액
                        
                    Else
                        '취소
                        With frm접수결제.sprCard
                            .Row = .ActiveRow
                            .DeleteRows .Row, 1
                            
                            .MaxRows = .MaxRows - 1
                        End With
                        
                        frm접수결제.txtCard.Value = frm접수결제.txtCard.Value - txtMoney.Value
                    End If
                    
            Case "접수2" '신용카드승인 취소
'                        Call 신용카드취소_Report(frm신용카드승인.KS7500i, _
'                                                 Spread_GetData(sprGrid, 1, 1, True), _
'                                                 Spread_GetData(sprGrid, 2, 1, True), _
'                                                 Spread_GetData(sprGrid, 3, 1, True))
                    
                    Call 매출취소_Rtn
                    
                    frm신용카드승인.Data_Display
                    
            Case "출고"
                If iFlag = "1" Then
                    With frm출고결제.sprCard
                        .MaxRows = .MaxRows + 1
                        .Row = .MaxRows
                    
                        .Col = 2:  .Text = Spread_GetData(sprGrid, 1, 1, True)   '승인번호
                        .Col = 3:  .Text = Spread_GetData(sprGrid, 2, 1, True)   '승인일자
                        .Col = 4:  .Text = Spread_GetData(sprGrid, 3, 1, True)   '승인시간
                        .Col = 5:  .Text = Spread_GetData(sprGrid, 4, 1, True)   '할부기간
                        .Col = 6:  .Text = Spread_GetData(sprGrid, 5, 1, True)   '결제금액
                        .Col = 7:  .Text = Spread_GetData(sprGrid, 6, 1, True)   '발급사코드
                        .Col = 8:  .Text = Spread_GetData(sprGrid, 7, 1, True)   '발급사명
                        .Col = 9:  .Text = Spread_GetData(sprGrid, 8, 1, True)   '매입사코드
                        .Col = 10: .Text = Spread_GetData(sprGrid, 9, 1, True)   '매입사명
                        .Col = 11: .Text = Spread_GetData(sprGrid, 10, 1, True)  '카드번호
                        .Col = 12: .Text = Spread_GetData(sprGrid, 11, 1, True)  '메시지1
                        .Col = 13: .Text = Spread_GetData(sprGrid, 12, 1, True)  '메시지2
                        
                        For i = 1 To .MaxRows
                            .Row = i
                            .Col = 6: 카드결제금액 = 카드결제금액 + .Value
                        Next i
                    End With
                    
                    frm출고결제.txtCard.Value = 카드결제금액
                    
                Else
                    '취소
                    With frm출고결제.sprCard
                        .Row = .ActiveRow
                        .DeleteRows .Row, 1
                        
                        .MaxRows = .MaxRows - 1
                    End With
                    
                    Call 매출취소_Rtn
                    
                    frm출고결제.txtCard.Value = frm출고결제.txtCard.Value - txtMoney.Value
                End If
                
            Case "판매취소" 'frm판매취소결제
                '취소
                With frm판매취소결제.sprCard
                    .Row = .ActiveRow
                    .DeleteRows .Row, 1
                    
                    .MaxRows = .MaxRows - 1
                End With
                
                Call 매출취소_Rtn
                
                결제취소여부 = True
                
            Case "판매취소2"
                If iFlag = "1" Then
                    With frm판매취소.sprCard
                        .MaxRows = .MaxRows + 1
                        .Row = .MaxRows
                    
                        .Col = 2:  .Text = Spread_GetData(sprGrid, 1, 1, True)   '승인번호
                        .Col = 3:  .Text = Spread_GetData(sprGrid, 2, 1, True)   '승인일자
                        .Col = 4:  .Text = Spread_GetData(sprGrid, 3, 1, True)   '승인시간
                        .Col = 5:  .Text = Spread_GetData(sprGrid, 4, 1, True)   '할부기간
                        .Col = 6:  .Text = Spread_GetData(sprGrid, 5, 1, True)   '결제금액
                        .Col = 7:  .Text = Spread_GetData(sprGrid, 6, 1, True)   '발급사코드
                        .Col = 8:  .Text = Spread_GetData(sprGrid, 7, 1, True)   '발급사명
                        .Col = 9:  .Text = Spread_GetData(sprGrid, 8, 1, True)   '매입사코드
                        .Col = 10: .Text = Spread_GetData(sprGrid, 9, 1, True)   '매입사명
                        .Col = 11: .Text = Spread_GetData(sprGrid, 10, 1, True)  '카드번호
                        .Col = 12: .Text = Spread_GetData(sprGrid, 11, 1, True)  '메시지1
                        .Col = 13: .Text = Spread_GetData(sprGrid, 12, 1, True)  '메시지2
                        
                        For i = 1 To .MaxRows
                            .Row = i
                            .Col = 6: 카드결제금액 = 카드결제금액 + .Value
                        Next i
                    End With
                    
                    frm판매취소.txtCard.Value = 카드결제금액
                    
                Else
                    '취소
                    With frm판매취소.sprCard
                        .Row = .ActiveRow
                        .DeleteRows .Row, 1
                        
                        .MaxRows = .MaxRows - 1
                    End With
                    
                    Call 매출취소_Rtn

                End If
        End Select
        Unload Me
    End If
End Sub

