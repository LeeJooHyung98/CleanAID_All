VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm문자일별현황 
   AutoRedraw      =   -1  'True
   Caption         =   "문자 일별 현황"
   ClientHeight    =   12390
   ClientLeft      =   3630
   ClientTop       =   2400
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
      TabIndex        =   10
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
      Picture         =   "frm문자일별현황.frx":0000
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
      PaneTree        =   "frm문자일별현황.frx":2FCB
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   11160
         Index           =   0
         Left            =   15
         TabIndex        =   1
         Top             =   1215
         Width           =   5010
         _Version        =   524288
         _ExtentX        =   8837
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
         MaxCols         =   4
         MaxRows         =   300
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frm문자일별현황.frx":305D
         VisibleCols     =   2
         VisibleRows     =   50
         HighlightHeaders=   1
         HighlightStyle  =   1
         ScrollBarStyle  =   2
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   11160
         Index           =   1
         Left            =   5040
         TabIndex        =   2
         Top             =   1215
         Width           =   10785
         _Version        =   524288
         _ExtentX        =   19024
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
         SpreadDesigner  =   "frm문자일별현황.frx":36E5
         VisibleCols     =   2
         VisibleRows     =   50
         HighlightHeaders=   1
         HighlightStyle  =   1
         ScrollBarStyle  =   2
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   750
         Left            =   15
         TabIndex        =   3
         Top             =   450
         Width           =   15810
         _ExtentX        =   27887
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin CSTextLibCtl.silgEdit txtCount 
            Height          =   330
            Left            =   915
            TabIndex        =   13
            Top             =   390
            Width           =   1140
            _Version        =   262145
            _ExtentX        =   2011
            _ExtentY        =   582
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "0"
            Text            =   " 0"
            StartText.x     =   3
            StartText.y     =   4
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   14
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            Undo            =   1
            Data            =   0
         End
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   630
            Left            =   3210
            TabIndex        =   6
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 조회(&F)"
            Appearance      =   6
            Picture         =   "frm문자일별현황.frx":3DE1
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   3
            Left            =   5745
            TabIndex        =   7
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 엑셀(&E)"
            Appearance      =   6
            Picture         =   "frm문자일별현황.frx":44DB
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   8835
            TabIndex        =   8
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frm문자일별현황.frx":4C55
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   4
            Left            =   7290
            TabIndex        =   9
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 출력(&P)"
            Appearance      =   6
            Picture         =   "frm문자일별현황.frx":5CE7
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Left            =   915
            TabIndex        =   11
            Top             =   45
            Width           =   1140
            _ExtentX        =   2011
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
            Format          =   64356355
            UpDown          =   -1  'True
            CurrentDate     =   39596
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "발송합계:"
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
            Left            =   45
            TabIndex        =   12
            Top             =   450
            Width           =   840
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
            Left            =   45
            TabIndex        =   4
            Top             =   105
            Width           =   840
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   5
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
         Caption         =   "      문자 일별 현황"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm문자일별현황.frx":63E1
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm문자일별현황.frx":6607
            Top             =   -15
            Width           =   765
         End
      End
   End
End
Attribute VB_Name = "frm문자일별현황"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_ServerInfo(4)      As String
Dim SMSCon      As ADODB.Connection
Dim m_Connect            As Boolean
Dim FORM_SMS002_ACTIVATE As Boolean
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
    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure Form_Activate of Form P_SMS001"

End Sub

Private Sub Form_Load()
    
    m_Connect = False
    
    For i = 0 To 1
        With sprGrid(i)
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
    Next i
    
    dtpDay.Value = Format(Date, "YYYY-MM")
    
    sMasterCode = 가맹점정보.지사코드
    
    'TitleSet "일자별 발송 현황"
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    
    cmdBtn(5).Left = (Me.Width - 200) - cmdBtn(5).Width
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
    Dim sMsg    As String
    
    On Error GoTo ErrRtn
    
    txtCount.Value = 0
    
    pnlProg.Visible = True
    DoEvents
    
    sMsg = "111"
    
    ' 본사 연결 확인
    If m_Connect = False Then
        sMsg = "222"
        If CheckConnect = False Then
            MsgBox "본사와의 연결 설정을 확인하여 주십시요", vbInformation, "확인"
            ' 설정 화면을 활성화 한다.
            Call cmdBtn_Click(2)
            
            pnlProg.Visible = False
            Exit Sub
        End If
    End If
    
    sMsg = "333"
    
    '------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------
    Query = " EXEC PRO_SMS_STORE_002_01  '0', '" & 가맹점정보.가맹점코드 & "', "
    Query = Query & "'" & Format(dtpDay.Value, "YYYYMM") & "'  "
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, SMSCon, adOpenForwardOnly, adLockReadOnly
    
    'ADORset.CursorLocation = adUseClient
    'ADORset.Open Query, SMSCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    sMsg = "444"
    
    With sprGrid(0)
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = ADORs(0) & ""
            .Col = 2: .Text = ADORs(1) & ""
            .Col = 3: .Text = ""
            .Col = 4: .Text = ADORs(2) & ""
            
            If ADORs(2) & "" = "1" Then
                .Col = -1: .BackColor = vbRed
            End If
            
    sMsg = "555"
            txtCount.Value = txtCount.Value + ADORs(1)
            
            ADORs.MoveNext
        Loop
        
        .ReDraw = True
        
        ADORs.Close
        Set ADORs = Nothing
    End With
    
    sMsg = "666"
    
    pnlProg.Visible = False
    Exit Sub

ErrRtn:
    m_Connect = False
    pnlProg.Visible = False
    sMsg = "777"
    
    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure GetData_View of Form P_SMS002" & sMsg & Query
End Sub

Private Sub sprGrid_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
    Dim varTemp As Variant
    Dim sMode   As Variant
    
    ' 좌측 그리드를 클릭한 경우 해당 일자의 세부 내역을 조회 한다.
    If Index = 0 Then
        sprGrid(0).Row = Row
        sprGrid(0).Col = 1: varTemp = sprGrid(0).Text
        sprGrid(0).Col = 4: sMode = sprGrid(0).Text
        
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
    Query = "EXEC PRO_SMS_STORE_002_02"
    Query = Query & "  '0'"
    Query = Query & ", '" & 가맹점정보.가맹점코드 & "'"
    Query = Query & ", '" & sDate & "'"
    Query = Query & ", '" & sMode & "'"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, SMSCon, adOpenForwardOnly, adLockReadOnly
    
    'ADORset.CursorLocation = adUseClient
    'ADORset.Open Query, SMSCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    
    With sprGrid(1)
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
            
            If Left(Trim(ADORs(5) & ""), 2) <> "06" Then
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
    m_Connect = False
    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure GetData_ViewDetailed of Form P_SMS002"
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
