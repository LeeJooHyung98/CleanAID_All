VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm문자특정번호 
   AutoRedraw      =   -1  'True
   Caption         =   "문자 특정번호 현황"
   ClientHeight    =   12390
   ClientLeft      =   4965
   ClientTop       =   2595
   ClientWidth     =   15840
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
   LinkTopic       =   "Form23"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12390
   ScaleWidth      =   15840
   WindowState     =   2  '최대화
   Begin Threed.SSPanel pnlProg 
      Height          =   1215
      Left            =   600
      TabIndex        =   9
      Top             =   1770
      Visible         =   0   'False
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   2143
      _Version        =   262144
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
      Picture         =   "frm문자특정번호.frx":0000
      Caption         =   " "
      BorderWidth     =   0
      BevelOuter      =   1
      BevelInner      =   2
      Alignment       =   4
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   12390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15840
      _ExtentX        =   27940
      _ExtentY        =   21855
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm문자특정번호.frx":2FCB
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   11160
         Left            =   15
         TabIndex        =   1
         Top             =   1215
         Width           =   15810
         _Version        =   524288
         _ExtentX        =   27887
         _ExtentY        =   19685
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowUserFormulas=   -1  'True
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
         DInformActiveRowChange=   0   'False
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
         MaxCols         =   6
         MaxRows         =   300
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frm문자특정번호.frx":303D
         VisibleCols     =   2
         VisibleRows     =   50
         HighlightHeaders=   1
         HighlightStyle  =   1
         ScrollBarStyle  =   2
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   750
         Left            =   15
         TabIndex        =   2
         Top             =   450
         Width           =   15810
         _ExtentX        =   27887
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtData 
            Height          =   315
            Left            =   1110
            ScrollBars      =   2  '수직
            TabIndex        =   14
            Top             =   390
            Width           =   1545
         End
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   630
            Left            =   4560
            TabIndex        =   5
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 조회(&F)"
            Appearance      =   6
            Picture         =   "frm문자특정번호.frx":3688
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   3
            Left            =   6075
            TabIndex        =   6
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 엑셀(&E)"
            Appearance      =   6
            Picture         =   "frm문자특정번호.frx":3D82
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   9165
            TabIndex        =   7
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frm문자특정번호.frx":44FC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   4
            Left            =   7620
            TabIndex        =   8
            Top             =   60
            Visible         =   0   'False
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 출력(&P)"
            Appearance      =   6
            Picture         =   "frm문자특정번호.frx":558E
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   0
            Left            =   1125
            TabIndex        =   10
            Top             =   45
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM"
            Format          =   56950785
            UpDown          =   -1  'True
            CurrentDate     =   39596
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   1
            Left            =   2910
            TabIndex        =   11
            Top             =   60
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM"
            Format          =   56950785
            UpDown          =   -1  'True
            CurrentDate     =   39596
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "(4자리 또는 전체)"
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
            Left            =   2760
            TabIndex        =   15
            Top             =   480
            Width           =   1530
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "~"
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
            Left            =   2730
            TabIndex        =   13
            Top             =   120
            Width           =   90
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "휴대폰번호:"
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
            Left            =   90
            TabIndex        =   12
            Top             =   420
            Width           =   990
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "전송일자:"
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
            Left            =   195
            TabIndex        =   3
            Top             =   105
            Width           =   840
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   4
         Top             =   15
         Width           =   15810
         _ExtentX        =   27887
         _ExtentY        =   741
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "      문자 특정번호 현황"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm문자특정번호.frx":5C88
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm문자특정번호.frx":5EAE
            Top             =   -15
            Width           =   765
         End
      End
   End
End
Attribute VB_Name = "frm문자특정번호"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_ServerInfo(4)      As String
Dim SMSCon      As ADODB.Connection
Dim m_Connect            As Boolean
Dim FORM_SMS005_ACTIVATE As Boolean
Dim sMasterCode          As String


Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 3: Call Export_Excel(frmMain.cdgExcel, sprGrid(1))
        Case 4:
        Case 5: Unload Me
    End Select
End Sub

 
 

Private Sub cmdList_Click()
    Call GetData_View
End Sub

Private Sub Form_Activate()
    On Error GoTo ErrRtn

    If FORM_SMS005_ACTIVATE = True Then Exit Sub
    
    FORM_SMS005_ACTIVATE = True
    
    DoEvents
  
    On Error GoTo 0
    Exit Sub

ErrRtn:
    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure Form_Activate of Form P_SMS001"

End Sub

Private Sub Form_Load()
    
    m_Connect = False
    
    With sprGrid
        .MaxRows = 0
        .RowHeight(-1) = 14
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeSingle
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
    
    
    dtpDay(0).Value = Format(DateAdd("d", -10, Date), "YYYY-MM-DD")
    dtpDay(1).Value = Format(Date, "YYYY-MM-DD")
    
    
    sMasterCode = 가맹점정보.지사코드
    
    'TitleSet "일자별 발송 현황"
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    
    cmdBtn(5).Left = (Me.Width - 200) - cmdBtn(5).Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FORM_SMS005_ACTIVATE = False
End Sub
 
'Private Sub MaskEdBox1_GotFocus(Index As Integer)
'    MaskEdBox1(Index).SelStart = 0
'    MaskEdBox1(Index).SelLength = Len(MaskEdBox1(Index).Text)
'End Sub
'
'Private Sub MaskEdBox2_GotFocus(Index As Integer)
'    MaskEdBox2(Index).SelStart = 0
'    MaskEdBox2(Index).SelLength = Len(MaskEdBox2(Index).Text)
'End Sub

Private Function CheckConnect() As Boolean
    On Error GoTo ErrRtn
    
    Dim HostConn    As String
    
    Call DefaultServerSetting
    
    HostConn = ""
    HostConn = HostConn & "Provider=SQLOLEDB.1;"
    HostConn = HostConn & "Persist Security Info=False;"
    HostConn = HostConn & "User ID=" & m_ServerInfo(2) & ";"
    HostConn = HostConn & "Password=" & m_ServerInfo(3) & ";"
    HostConn = HostConn & "Initial Catalog=" & m_ServerInfo(1) & ";"
    HostConn = HostConn & "Data Source=" & m_ServerInfo(0)
    m_CommandTimeOut = IIf(m_CommandTimeOut = 0, 30, m_CommandTimeOut)

    Set SMSCon = Nothing
    Set SMSCon = New ADODB.Connection
    
    If SMSCon.State = adStateOpen Then SMSCon.Close
    
    SMSCon.ConnectionTimeout = 10
    SMSCon.CommandTimeout = m_CommandTimeOut
    SMSCon.Open HostConn
    
    m_Connect = True
    CheckConnect = True
    
    On Error GoTo 0
    
    Exit Function

ErrRtn:
    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure CheckConnect of Form P_SMS001"
End Function



'--------------------------------------------------------------------------------------------------------------
' Procedure : GetData1
' DateTime  : 2007-01-08 22:39
' Author    : pds2004
' Purpose   :
'--------------------------------------------------------------------------------------------------------------
Private Sub GetData_View()
    Dim bResult As Boolean
    Dim lRow    As Long
    
    On Error GoTo ErrRtn
    
    pnlProg.Visible = True
    DoEvents
    
    ' 본사 연결 확인
    If m_Connect = False Then
        If CheckConnect = False Then
            MsgBox "본사와의 연결 설정을 확인하여 주십시요", vbInformation, "확인"
            ' 설정 화면을 활성화 한다.
            Call cmdBtn_Click(2)
            
            pnlProg.Visible = False
            Exit Sub
        End If
    End If
    
    '------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------
    Query = " EXEC PRO_SMS_STORE_004_01  '0', '" & 가맹점정보.가맹점코드 & "', "
    Query = Query & "'%" & Trim(txtData.Text) & "%',  "
    Query = Query & "'" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "',  "
    Query = Query & "'" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "'  "
    
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, SMSCon, adOpenForwardOnly, adLockReadOnly
    
    'ADORset.CursorLocation = adUseClient
    'ADORset.Open Query, SMSCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    
    
    
    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = ADORs(0) & ""
            .Col = 2: .Text = ADORs(1) & ""
            .Col = 3: .Text = ADORs(2) & ""
            .Col = 4: .Text = ADORs(3) & ""
            .Col = 5: .Text = ADORs(4) & ""
            .Col = 6: .Text = ADORs(5) & ""
            
            If Mid(CStr(ADORs(5)), 1, 2) & "" <> "06" Then
                .Col = -1: .BackColor = vbRed
            End If
            
            
            ADORs.MoveNext
        Loop
        
        .ReDraw = True
        
        ADORs.Close
        Set ADORs = Nothing
    End With
    
    pnlProg.Visible = False
    Exit Sub

ErrRtn:
    m_Connect = False
    pnlProg.Visible = False
    
    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure GetData_View of Form P_SMS002"
End Sub

  

Private Sub DefaultServerSetting()
    ' 기본 설정 정보가 없을 경우
    On Error GoTo ErrRtn
    
    Query = "SELECT * FROM TB_기본정보 "
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic

    If ADORs.RecordCount > 0 Then
        m_ServerInfo(0) = Trim(ADORs.Fields("SMS_IP") & "")
        m_ServerInfo(1) = Trim(ADORs.Fields("SMS_DB") & "")
        m_ServerInfo(2) = Trim(ADORs.Fields("SMS_ID") & "")
        m_ServerInfo(3) = Trim(ADORs.Fields("SMS_PWD") & "")
        m_CommandTimeOut = Val(Trim(ADORs.Fields("TimeOut") & ""))
    Else
        m_ServerInfo(0) = "115.89.220.5,8657"
        m_ServerInfo(1) = "Laundry1000"
        m_ServerInfo(2) = "sa"
        m_ServerInfo(3) = "cleanaid1996!@#"
        m_CommandTimeOut = 30
    End If
    ADORs.Close
    Set ADORs = Nothing

    On Error GoTo 0
    
    Exit Sub

ErrRtn:
    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure DefaultServerSetting of Form P_SMS001"
End Sub

Private Sub txtData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call GetData_View
    End If
End Sub
