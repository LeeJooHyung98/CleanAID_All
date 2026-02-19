VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm문자월별현황 
   AutoRedraw      =   -1  'True
   Caption         =   "문자 월별 현황"
   ClientHeight    =   12390
   ClientLeft      =   5115
   ClientTop       =   1575
   ClientWidth     =   16050
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
   ScaleWidth      =   16050
   WindowState     =   2  '최대화
   Begin Threed.SSPanel pnlProg 
      Height          =   1215
      Left            =   4335
      TabIndex        =   15
      Top             =   1605
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
      Picture         =   "frm문자월별현황.frx":0000
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
      Width           =   16050
      _ExtentX        =   28310
      _ExtentY        =   21855
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm문자월별현황.frx":2FCB
      Begin Threed.SSPanel SSPanel 
         Height          =   480
         Left            =   15
         TabIndex        =   11
         Top             =   11895
         Width           =   4260
         _ExtentX        =   7514
         _ExtentY        =   847
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin CSTextLibCtl.silgEdit txtCount 
            Height          =   375
            Index           =   0
            Left            =   1530
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   45
            Width           =   1215
            _Version        =   262145
            _ExtentX        =   2143
            _ExtentY        =   661
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
            DataProperty    =   2
            ReadOnly        =   -1  'True
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "0"
            Text            =   " 0"
            StartText.x     =   2
            StartText.y     =   5
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
            Undo            =   1
            Data            =   0
         End
         Begin CSTextLibCtl.silgEdit txtCount 
            Height          =   375
            Index           =   1
            Left            =   2775
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   45
            Width           =   1215
            _Version        =   262145
            _ExtentX        =   2143
            _ExtentY        =   661
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
            DataProperty    =   2
            ReadOnly        =   -1  'True
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "0"
            Text            =   " 0"
            StartText.x     =   2
            StartText.y     =   5
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
            Undo            =   1
            Data            =   0
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "발송합계 :"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   600
            TabIndex        =   13
            Top             =   90
            Width           =   825
         End
      End
      Begin FPSpreadADO.fpSpread sprList 
         Height          =   11160
         Left            =   4290
         TabIndex        =   1
         Top             =   1215
         Width           =   11745
         _Version        =   524288
         _ExtentX        =   20717
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
         SpreadDesigner  =   "frm문자월별현황.frx":307D
         VisibleCols     =   2
         VisibleRows     =   50
         HighlightHeaders=   1
         HighlightStyle  =   1
         ScrollBarStyle  =   2
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   16020
         _ExtentX        =   28258
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
         Caption         =   "      문자 월별 현황"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm문자월별현황.frx":373E
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm문자월별현황.frx":3964
            Top             =   -15
            Width           =   765
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   750
         Left            =   15
         TabIndex        =   3
         Top             =   450
         Width           =   16020
         _ExtentX        =   28258
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   630
            Left            =   3210
            TabIndex        =   4
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 조회(&F)"
            Appearance      =   6
            Picture         =   "frm문자월별현황.frx":452E
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   3
            Left            =   5745
            TabIndex        =   5
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 엑셀(&E)"
            Appearance      =   6
            Picture         =   "frm문자월별현황.frx":4C28
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   8835
            TabIndex        =   6
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frm문자월별현황.frx":53A2
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   4
            Left            =   7290
            TabIndex        =   7
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 출력(&P)"
            Appearance      =   6
            Picture         =   "frm문자월별현황.frx":6434
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Left            =   915
            TabIndex        =   8
            Top             =   45
            Width           =   1125
            _ExtentX        =   1984
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
            CustomFormat    =   "yyyy"
            Format          =   59506691
            UpDown          =   -1  'True
            CurrentDate     =   39596
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
            TabIndex        =   9
            Top             =   105
            Width           =   840
         End
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   10665
         Left            =   15
         TabIndex        =   10
         Top             =   1215
         Width           =   4260
         _Version        =   524288
         _ExtentX        =   7514
         _ExtentY        =   18812
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
         MaxCols         =   3
         MaxRows         =   300
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frm문자월별현황.frx":6B2E
         VisibleCols     =   2
         VisibleRows     =   50
         HighlightHeaders=   1
         HighlightStyle  =   1
         ScrollBarStyle  =   2
      End
   End
End
Attribute VB_Name = "frm문자월별현황"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_ServerInfo(4)     As String
Dim m_Host_DataBase     As ADODB.Connection
Dim m_Connect           As Boolean
Dim FORM_SMS003_ACTIVATE    As Boolean
Dim sMasterCode        As String


Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 3: Call Export_Excel(frmMain.cdgExcel, sprList)
        Case 4:
        Case 5: Unload Me
    End Select
End Sub
 
Private Sub cmdList_Click()
    Call Data_Display
End Sub

Private Sub Form_Activate()
    On Error GoTo ErrRtn

    If FORM_SMS003_ACTIVATE = True Then Exit Sub
    
    FORM_SMS003_ACTIVATE = True
    
    DoEvents
  
    On Error GoTo 0
    Exit Sub

ErrRtn:
    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure Form_Activate of Form P_SMS003"

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
    
    With sprList
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

    dtpDay.Year = Format(Date, "YYYY")
    
    sMasterCode = 가맹점정보.지사코드
    
    'TitleSet "일자별 발송 현황"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    cmdBtn(5).Left = (Me.Width - 200) - cmdBtn(5).Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FORM_SMS003_ACTIVATE = False
End Sub
 
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
    
    If m_Host_DataBase.State = adStateOpen Then m_Host_DataBase.Close
    
    m_Host_DataBase.ConnectionTimeout = 10
    m_Host_DataBase.CommandTimeout = m_CommandTimeOut
    m_Host_DataBase.Open HostConn
    
    m_Connect = True
    
    CheckConnect = True
    
    Exit Function

ErrRtn:
    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure CheckConnect of Form P_SMS003"
End Function

'--------------------------------------------------------------------------------------------------------------
' Procedure : GetData1
' DateTime  : 2007-01-08 22:39
' Author    : pds2004
' Purpose   :
'--------------------------------------------------------------------------------------------------------------
Private Sub Data_Display()
    Dim sMonth As String
    
    Dim Total_Num1 As Long
    Dim Total_Num2 As Long
    
    On Error GoTo ErrRtn
    
    txtCount(0).Value = 0
    txtCount(1).Value = 0
    'txtCount(2).Value = 0
   
    ' 본사 연결 확인
    If m_Connect = False Then
        If CheckConnect = False Then
            MsgBox "본사와의 연결 설정을 확인하여 주십시요", vbInformation, "확인"
            ' 설정 화면을 활성화 한다.
            Call cmdBtn_Click(2)
            Exit Sub
        End If
    End If
    
    '-----------------------------------------------------------------------------------
    ' PRO_SMS_STORE_003_01
    '-----------------------------------------------------------------------------------
    Query = "EXEC PRO_SMS_STORE_003_01"
    Query = Query & "  '0'"
    Query = Query & ", '" & 가맹점정보.가맹점코드 & "'"
    Query = Query & ", '" & Format(dtpDay.Value, "YYYY") & "'"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, m_Host_DataBase, adOpenForwardOnly, adLockReadOnly
    
    'ADORset.CursorLocation = adUseClient
    'ADORset.Open Query, m_Host_DataBase, adOpenStatic, adLockBatchOptimistic, adCmdText
    
    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        sMonth = ""
        
        Do Until ADORs.EOF
            If sMonth = "" Or sMonth <> ADORs(0) Then
                .MaxRows = .MaxRows + 1
            End If
            
            .Row = .MaxRows
            .Col = 1: .Text = ADORs(0) & "월"
            
            If ADORs(1) > 0 Then .Col = 3: .Text = ADORs(1) & ""
            'If ADORs(2) > 0 Then .Col = 4: .Text = ADORs(2) & ""
            
            sMonth = ADORs(0)
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
            
        For i = 1 To .MaxRows
            .Row = i
            .Col = 3: Total_Num1 = .Value
            '.Col = 4: Total_Num2 = .Value
            
            .Col = 2: .Value = Total_Num1 ' + Total_Num2
        
            txtCount(1).Value = txtCount(1).Value + Total_Num1
            'txtCount(2).Value = txtCount(2).Value + Total_Num2
        Next i
        
        txtCount(0).Value = txtCount(1).Value ' + txtCount(2).Value
        
        .ReDraw = True
    End With
        
    Exit Sub

ErrRtn:
    m_Connect = False
    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure Data_Display of Form P_SMS003"
End Sub


'--------------------------------------------------------------------------------------------------------------
' Procedure : GetData_View2
' DateTime  : 2007-01-08 22:39
' Author    : pds2004
' Purpose   :
'--------------------------------------------------------------------------------------------------------------
Private Sub GetData_View2(ByVal sDate As String)
    Dim ADORset As New ADODB.Recordset
    
    Dim sDay       As String
    
    Dim Total_Num1 As Long
    Dim Total_Num2 As Long
    Dim Total_Num3 As Long
    
    'On Error GoTo ErrRtn
    
    ' 본사 연결 확인
    If m_Connect = False Then
        If CheckConnect = False Then
            MsgBox "본사와의 연결 설정을 확인하여 주십시요", vbInformation, "확인"
            ' 설정 화면을 활성화 한다.
            Call cmdBtn_Click(2)
            Exit Sub
        End If
    End If
    
    pnlProg.Visible = True
    DoEvents
    
    '---------------------------------------------------------------------------
    '
    '---------------------------------------------------------------------------
    Query = "EXEC PRO_SMS_STORE_002_01 "
    Query = Query & "  '0'"
    Query = Query & ", '" & 가맹점정보.가맹점코드 & "'"
    Query = Query & ", '" & sDate & "'"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, m_Host_DataBase, adOpenForwardOnly, adLockReadOnly
    
    'ADORset.CursorLocation = adUseClient
    'ADORset.Open Query, m_Host_DataBase, adOpenStatic, adLockBatchOptimistic, adCmdText
    
    With sprList
        .MaxRows = 0
        .ReDraw = False
        
        sDay = ""
        
        Do Until ADORs.EOF
            If sDay = "" Or sDay <> ADORs(0) Then
                .MaxRows = .MaxRows + 1
            End If
        
            .Row = .MaxRows
            
            .Col = 1: .Text = ADORs(0) & ""
            .Col = 2: .Value = ADORs(1) & "" '성공
            .Col = 3: .Value = ADORs(1) & "" '성공
            .Col = 4: .Text = ""
            
            'Select Case Left(ADORs(3), 2)
            '    Case "06": .Col = 3: .Value = ADORs(1) & "" '성공
            '    Case "00": .Col = 5: .Value = ADORs(1) & "" '대기
            '    Case Else: .Col = 4: .Value = ADORs(1) & "" '실패
            'End Select
            
            sDay = ADORs(0)
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        'For i = 1 To .MaxRows
        '    .Row = i
        '    .Col = 3: Total_Num1 = IIf(.Value = "", 0, .Value)
        '    .Col = 4: Total_Num2 = IIf(.Value = "", 0, .Value)
        '    .Col = 5: Total_Num3 = IIf(.Value = "", 0, .Value)
        '
        '    .Col = 2: .Value = Total_Num1 + Total_Num2 + Total_Num3
        'Next i
        
        .ReDraw = True
    End With
    
    pnlProg.Visible = False
    
    Exit Sub

ErrRtn:
    pnlProg.Visible = False
    
    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure GetData_View2 of Form P_SMS003"
End Sub

Private Sub sprGrid_Click(ByVal Col As Long, ByVal Row As Long)
    Dim varTemp As Variant
    
    If Row <= 0 Then
        sprList.MaxRows = 0
        
        Exit Sub
    End If
    
    sprGrid.Enabled = False
    DoEvents
    
    sprGrid.Row = Row
    sprGrid.Col = 1: varTemp = CStr(Replace(CStr(sprGrid.Text), "월", ""))
    
    Call GetData_View2(Format(dtpDay.Value, "YYYY") & CStr(varTemp))
    
    sprGrid.Enabled = True
End Sub

Private Sub sprGrid_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
''    Call sprGrid_Click(NewCol, NewRow)
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
    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure DefaultServerSetting of Form P_SMS003"
End Sub
