VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm스타밴코리아 
   BorderStyle     =   1  '단일 고정
   Caption         =   "카드결제"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4830
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   4830
   StartUpPosition =   2  '화면 가운데
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   4320
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   7620
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm스타밴코리아.frx":0000
      Begin Threed.SSPanel SSPanel2 
         Height          =   675
         Left            =   15
         TabIndex        =   7
         Top             =   3630
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   1191
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton btnOK 
            Height          =   585
            Left            =   45
            TabIndex        =   8
            Top             =   45
            Width           =   1275
            _Version        =   851970
            _ExtentX        =   2249
            _ExtentY        =   1032
            _StockProps     =   79
            Caption         =   "승인 요청"
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
            Picture         =   "frm스타밴코리아.frx":0052
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   585
            Index           =   2
            Left            =   3495
            TabIndex        =   10
            Top             =   45
            Width           =   1260
            _Version        =   851970
            _ExtentX        =   2222
            _ExtentY        =   1032
            _StockProps     =   79
            Caption         =   " 닫기"
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
            Picture         =   "frm스타밴코리아.frx":092C
         End
      End
      Begin Threed.SSPanel pnlBack 
         Height          =   3600
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   6350
         _Version        =   262144
         BackColor       =   16777215
         PictureFrames   =   1
         Picture         =   "frm스타밴코리아.frx":19BE
         PictureBackgroundStyle=   1
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
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
            Left            =   990
            Style           =   2  '드롭다운 목록
            TabIndex        =   13
            Top             =   1065
            Width           =   1410
         End
         Begin VB.ComboBox cboMonth 
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
            Left            =   990
            Style           =   2  '드롭다운 목록
            TabIndex        =   11
            Top             =   2385
            Width           =   1410
         End
         Begin VB.TextBox txtData 
            Appearance      =   0  '평면
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
            Height          =   705
            IMEMode         =   3  '사용 못함
            Index           =   0
            Left            =   975
            MultiLine       =   -1  'True
            PasswordChar    =   "*"
            TabIndex        =   0
            Top             =   1545
            Width           =   3750
         End
         Begin CSTextLibCtl.silgEdit txtMoney 
            Height          =   405
            Left            =   975
            TabIndex        =   3
            Top             =   2865
            Width           =   1425
            _Version        =   262145
            _ExtentX        =   2514
            _ExtentY        =   714
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   255
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
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Insert          =   0   'False
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "0"
            Text            =   " 0"
            StartText.x     =   2
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
            Undo            =   1
            Data            =   0
         End
         Begin XtremeSuiteControls.PushButton btnCardRead 
            Height          =   420
            Left            =   3075
            TabIndex        =   9
            Top             =   2295
            Width           =   1635
            _Version        =   851970
            _ExtentX        =   2884
            _ExtentY        =   741
            _StockProps     =   79
            Caption         =   "카드 다시 읽기"
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
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "거래구분:"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   0
            Left            =   105
            TabIndex        =   14
            Top             =   1155
            Width           =   840
         End
         Begin VB.Label lblTerminal 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "99999999"
            ForeColor       =   &H00C0C0C0&
            Height          =   180
            Left            =   3885
            TabIndex        =   12
            Top             =   195
            Width           =   720
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "카드번호:"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   4
            Left            =   105
            TabIndex        =   6
            Top             =   1650
            Width           =   840
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "할    부:"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   1
            Left            =   105
            TabIndex        =   5
            Top             =   2490
            Width           =   840
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "금    액:"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   2
            Left            =   105
            TabIndex        =   4
            Top             =   2985
            Width           =   840
         End
      End
   End
End
Attribute VB_Name = "frm스타밴코리아"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCardRead_Click()
    txtData(0).Text = ""
    txtData(0).SetFocus
End Sub

Private Sub btnOK_Click()
''    Dim lsPosReq As String * 1024
''    Dim lsPosRep As String * 1024
''    Dim lnResult As Integer
''
''    Dim strMoney As String
''
''    '020010999999900000Axxxxxxxxxxxxxxxx=xxxxxxxxxxxxxxxxxxxx  1004                                                    00
''
''    '-------------------------------------------------------------------
''    strMoney = txtMoney.Value
''
''    If cboGubun.Text = "신용승인" Then
''        Query = "0200"                                         '전문구분
''    Else
''        Query = "0420"                                         '전문구분
''    End If
''
''    Query = Query & "10"                                                '업무구분
''    Query = Query & Format(lblTerminal.Caption, "00000000")             '단말기번호
''    Query = Query & "0000"                                              '전표번호
''    Query = Query & "A"                                                 'WCC
''    Query = Query & txtData(0).Text                                     '카드정보
''    Query = Query & Format(cboMonth.ItemData(cboMonth.ListIndex), "00") '할부개월
''    Query = Query & Trim(strMoney) & Space(12 - Len(strMoney))          '거래금액
''    Query = Query & Space(12)                                           '봉사료
''    Query = Query & Space(12)                                           '세금
''    Query = Query & Space(12)                                           '원승인번호
''    Query = Query & Format(Date, "YYYYMMDD")                            '원거래일자
''    Query = Query & "00" & Space(16)                                    '비밀번호
''    Query = Query & Space(1)                                            'Payback 거래구분
''    Query = Query & Space(6)                                            '동글정보
''
''    lsPosReq = Query
''
''    lnResult = svk_POS(lsPosReq, lsPosRep)
''
''    'txtData(1).Text = lsPosRep
''
''    'MsgBox "Return Code : " + CStr(lnResult)
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 2: Unload Me
    End Select
End Sub

Private Sub Form_Load()
    With cboGubun
        .Clear
        .AddItem "신용승인"
        .AddItem "신용취소"
        
        .ListIndex = 0
    End With
    
    With cboMonth
        .Clear
        .AddItem "일시불": .ItemData(.NewIndex) = 0
        
        For i = 2 To 36
            .AddItem i & "개월": .ItemData(.NewIndex) = i
        Next i
        
        .ListIndex = 0
    End With
End Sub

Private Sub txtData_Change(Index As Integer)
    If Len(txtData(0).Text) >= 37 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtData_GotFocus(Index As Integer)
    txtData(Index).SelStart = 0
    txtData(Index).SelLength = Len(txtData(Index).Text)
End Sub
