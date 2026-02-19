VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form P_SMS002 
   AutoRedraw      =   -1  'True
   ClientHeight    =   12390
   ClientLeft      =   3420
   ClientTop       =   3885
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
   Begin Threed.SSPanel SSPanel1 
      Height          =   1125
      Left            =   585
      TabIndex        =   9
      Top             =   1125
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   1984
      _Version        =   262144
      ForeColor       =   0
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "맑은 고딕"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "서버에 연결중 입니다. 잠시만 기다려 주십시요..."
      FloodColor      =   16777215
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
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
      PaneTree        =   "P_SMS002.frx":0000
      Begin Threed.SSPanel SSPanel2 
         Height          =   630
         Left            =   15
         TabIndex        =   3
         Top             =   15
         Width           =   15810
         _ExtentX        =   27887
         _ExtentY        =   1111
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   540
            Index           =   0
            Left            =   2280
            TabIndex        =   4
            Top             =   45
            Width           =   1245
            _Version        =   851970
            _ExtentX        =   2196
            _ExtentY        =   952
            _StockProps     =   79
            Caption         =   " 조회"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Picture         =   "P_SMS002.frx":0072
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   495
            Left            =   645
            TabIndex        =   5
            Top             =   75
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM"
            Format          =   62783491
            UpDown          =   -1  'True
            CurrentDate     =   40279
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "발송 합계"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   13
            Left            =   12810
            TabIndex        =   8
            Top             =   150
            Width           =   1050
         End
         Begin VB.Label lblSMS 
            Alignment       =   1  '오른쪽 맞춤
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  '단일 고정
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   13935
            TabIndex        =   7
            Top             =   75
            Width           =   1125
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "일자"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   75
            TabIndex        =   6
            Top             =   135
            Width           =   480
         End
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   11715
         Index           =   0
         Left            =   15
         TabIndex        =   1
         Top             =   660
         Width           =   3210
         _Version        =   524288
         _ExtentX        =   5662
         _ExtentY        =   20664
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowUserFormulas=   -1  'True
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
         DInformActiveRowChange=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   3
         MaxRows         =   300
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "P_SMS002.frx":0A84
         VisibleCols     =   2
         VisibleRows     =   50
         HighlightHeaders=   1
         HighlightStyle  =   1
         ScrollBarStyle  =   2
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   11715
         Index           =   1
         Left            =   3240
         TabIndex        =   2
         Top             =   660
         Width           =   12585
         _Version        =   524288
         _ExtentX        =   22199
         _ExtentY        =   20664
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowUserFormulas=   -1  'True
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
         DInformActiveRowChange=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   6
         MaxRows         =   300
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "P_SMS002.frx":1064
         VisibleCols     =   2
         VisibleRows     =   50
         HighlightHeaders=   1
         HighlightStyle  =   1
         ScrollBarStyle  =   2
      End
   End
End
Attribute VB_Name = "P_SMS002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_ServerInfo(4)      As String
Dim m_Host_DataBase      As ADODB.Connection
Dim m_Connect            As Boolean
Dim FORM_SMS002_ACTIVATE As Boolean
Dim sMasterCode          As String


Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        ' 조회
        Case 0
            Call GetData_View
            Exit Sub
     
        Case Else
        
    End Select
End Sub

Private Sub dtpDay_Change()
    Call GetData_View
End Sub

Private Sub Form_Activate()
    On Error GoTo ErrRtn

    If FORM_SMS002_ACTIVATE = True Then Exit Sub
    
    FORM_SMS002_ACTIVATE = True
    
    DoEvents
  
    On Error GoTo 0
    Exit Sub

ErrRtn:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Activate of Form P_SMS001"

End Sub

Private Sub Form_Load()
    
    SSPanel1.Visible = False

    'MaskEdBox1(0).Text = Format(Date, "yyyy")
    'MaskEdBox2(0).Text = Format(Date, "mm")
    
    dtpDay.Value = Format(Date, "YYYY-MM")
    
    fpSpread1(0).ColWidth(3) = 0
    
    sMasterCode = 가맹점정보.지사코드
    
    'TitleSet "일자별 발송 현황"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    FORM_SMS002_ACTIVATE = False
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

    Set m_Host_DataBase = Nothing
    Set m_Host_DataBase = New ADODB.Connection
    
    SSPanel1.ZOrder 0
    SSPanel1.Visible = True
    
    If m_Host_DataBase.State = adStateOpen Then m_Host_DataBase.Close
    
    m_Host_DataBase.ConnectionTimeout = 10
    m_Host_DataBase.CommandTimeout = m_CommandTimeOut
    m_Host_DataBase.Open HostConn
    
    SSPanel1.ZOrder 0
     SSPanel1.Visible = False
    
    m_Connect = True
    CheckConnect = True
    
    On Error GoTo 0
    
    Exit Function

ErrRtn:
    SSPanel1.ZOrder 0:  SSPanel1.Visible = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CheckConnect of Form P_SMS001"
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
    
    ' 본사 연결 확인
    If m_Connect = False Then
        If CheckConnect = False Then
            MsgBox "본사와의 연결 설정을 확인하여 주십시요", vbInformation, "확인"
            ' 설정 화면을 활성화 한다.
            Call cmdBtn_Click(2)
            Exit Sub
        End If
    End If
    
    fpSpread1(0).MaxRows = 0
    
    '------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------
    Query = " EXEC PRO_SMS_STORE_002_01  '0', '" & 가맹점정보.가맹점코드 & "', "
    Query = Query & "'" & Format(dtpDay.Value, "YYYYMM") & "'  "
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, m_Host_DataBase, adOpenForwardOnly, adLockReadOnly
    
    'ADORset.CursorLocation = adUseClient
    'ADORset.Open Query, m_Host_DataBase, adOpenStatic, adLockBatchOptimistic, adCmdText
    
    With fpSpread1(0)
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = ADORs.Fields(0) & ""
            .Col = 2: .Text = ADORs.Fields(1) & ""
            .Col = 3: .Text = ADORs.Fields(2) & ""
            
            If ADORs.Fields(2) & "" = "1" Then
                .Col = -1: .BackColor = vbRed
            End If
            
            ADORs.MoveNext
        Loop
        
        .ReDraw = True
        
        ADORs.Close
        Set ADORs = Nothing
    End With
    
    ' 합계 출력
    Call DataTotal
    
    Exit Sub

ErrRtn:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetData_View of Form P_SMS002"
End Sub


Private Sub DataTotal()
    Dim lRow    As Long
    Dim varTemp As Variant
    Dim LCount    As Long
    
    LCount = 0
    
    For lRow = 1 To fpSpread1(0).MaxRows
        Call fpSpread1(0).GetText(2, lRow, varTemp)
        
        LCount = LCount + Val(Replace(CStr(varTemp), ",", ""))
    Next lRow
    
    lblSMS(0).Caption = Format(LCount, "#,##0")
End Sub

Private Sub fpSpread1_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
    Dim varTemp As Variant
    Dim sMode   As Variant
    
    ' 좌측 그리드를 클릭한 경우 해당 일자의 세부 내역을 조회 한다.
    If Index = 0 Then
        Call fpSpread1(0).GetText(1, Row, varTemp)
        Call fpSpread1(0).GetText(3, Row, sMode)
        
        If IsDate(CStr(varTemp)) = True Then
            Call GetData_ViewDetailed(CStr(varTemp), CStr(sMode))
        End If
    End If
End Sub


'--------------------------------------------------------------------------------------------------------------
' Procedure : GetData_ViewDetailed
' DateTime  : 2007-01-08 22:39
' Author    : pds2004
' Purpose   :
'--------------------------------------------------------------------------------------------------------------
Private Sub GetData_ViewDetailed(ByVal sDate As String, ByVal sMode As String)
    Dim bResult As Boolean
    Dim lRow    As Long
    
    On Error GoTo ErrRtn

    ' 본사 연결 확인
    If m_Connect = False Then
        If CheckConnect = False Then
            MsgBox "본사와의 연결 설정을 확인하여 주십시요", vbInformation, "확인"
            ' 설정 화면을 활성화 한다.
            Call cmdBtn_Click(2)
            Exit Sub
        End If
    End If
    
    sDate = Replace(Replace(sDate, "-", ""), "/", "")
    
    '-----------------------------------------------------------------------------------
    '
    '-----------------------------------------------------------------------------------
    Query = " EXEC PRO_SMS_STORE_002_02  '0', '" & 가맹점정보.가맹점코드 & "', "
    Query = Query & "'" & sDate & "', '" & sMode & "' "
    
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, m_Host_DataBase, adOpenForwardOnly, adLockReadOnly
    
    'ADORset.CursorLocation = adUseClient
    'ADORset.Open Query, m_Host_DataBase, adOpenStatic, adLockBatchOptimistic, adCmdText
    
    With fpSpread1(1)
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = ADORs.Fields(0) & ""
            .Col = 2: .Text = ADORs.Fields(1) & ""
            .Col = 3: .Text = ADORs.Fields(2) & ""
            .Col = 4: .Text = ADORs.Fields(3) & ""
            .Col = 5: .Text = ADORs.Fields(4) & ""
            .Col = 6: .Text = ADORs.Fields(5) & ""
            
            If Left(Trim(ADORs.Fields(5) & ""), 2) <> "06" Then
                .Col = -1: .BackColor = vbGreen
            End If
            
            ADORs.MoveNext
        Loop
        
        .ReDraw = True
        
        ADORs.Close
        Set ADORs = Nothing
    End With
    
    Exit Sub

ErrRtn:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetData_ViewDetailed of Form P_SMS002"
End Sub

Private Sub DefaultServerSetting()
    ' 기본 설정 정보가 없을 경우
    On Error GoTo ErrRtn
    
    Query = "SELECT * FROM TB_기본정보 "
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic

    If ADORs.RecordCount > 0 Then
        m_ServerInfo(0) = Trim(ADORs.Fields("ServerIP") & "")
        m_ServerInfo(1) = Trim(ADORs.Fields("ServerDB") & "")
        m_ServerInfo(2) = Trim(ADORs.Fields("ServerUser") & "")
        m_ServerInfo(3) = Trim(ADORs.Fields("ServerPass") & "")
        m_CommandTimeOut = Val(Trim(ADORs.Fields("TimeOut") & ""))
    Else
        m_ServerInfo(0) = "store.clean-aid.co.kr,8657"
        m_ServerInfo(1) = "Laundry"
        m_ServerInfo(2) = "sa"
        m_ServerInfo(3) = ""
        m_CommandTimeOut = 30
    End If
    ADORs.Close
    Set ADORs = Nothing

    On Error GoTo 0
    
    Exit Sub

ErrRtn:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DefaultServerSetting of Form P_SMS001"
End Sub
