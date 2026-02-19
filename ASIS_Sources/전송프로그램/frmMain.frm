VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Auto DBUpdate"
   ClientHeight    =   6300
   ClientLeft      =   14025
   ClientTop       =   8010
   ClientWidth     =   5415
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
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   5415
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   6300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   11113
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frmMain.frx":08CA
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   900
         _Version        =   262144
         BackColor       =   16777215
         PictureFrames   =   1
         Picture         =   "frmMain.frx":095C
         PictureBackgroundStyle=   1
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Timer Timer1 
            Interval        =   10
            Left            =   2715
            Top             =   45
         End
         Begin XtremeSuiteControls.PushButton btnHide 
            Height          =   420
            Left            =   4815
            TabIndex        =   14
            Top             =   45
            Width           =   495
            _Version        =   851970
            _ExtentX        =   873
            _ExtentY        =   741
            _StockProps     =   79
            Appearance      =   6
            Picture         =   "frmMain.frx":1E6A
         End
         Begin VB.Label lblVer 
            BackStyle       =   0  '투명
            Height          =   255
            Left            =   3690
            TabIndex        =   24
            Top             =   165
            Width           =   915
         End
         Begin VB.Image Image4 
            Height          =   360
            Left            =   105
            Picture         =   "frmMain.frx":287C
            Top             =   75
            Width           =   360
         End
         Begin VB.Label lblOffice 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "1019"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   345
            Left            =   780
            TabIndex        =   15
            Top             =   90
            Width           =   720
         End
         Begin XtremeSuiteControls.TrayIcon TrayIcon 
            Left            =   3225
            Top             =   210
            _Version        =   851970
            _ExtentX        =   423
            _ExtentY        =   423
            _StockProps     =   16
            Text            =   "Auto DBUpdate"
            Picture         =   "frmMain.frx":2D46
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   2100
         Left            =   15
         TabIndex        =   2
         Top             =   540
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   3704
         _Version        =   262144
         BackColor       =   16777215
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkCloseDataOnly 
            BackColor       =   &H00FFFFFF&
            Caption         =   "마감자료만"
            Height          =   225
            Left            =   1980
            TabIndex        =   25
            Top             =   885
            Width           =   1575
         End
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00FFFFFF&
            Caption         =   "기간 전송:"
            Height          =   225
            Left            =   720
            TabIndex        =   19
            Top             =   540
            Width           =   1200
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   300
            Index           =   0
            Left            =   1935
            TabIndex        =   17
            Top             =   495
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   529
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   21364739
            UpDown          =   -1  'True
            CurrentDate     =   44927
         End
         Begin VB.CheckBox chkTotal 
            BackColor       =   &H00FFFFFF&
            Caption         =   "전체 전송"
            Height          =   225
            Left            =   720
            TabIndex        =   4
            Top             =   885
            Width           =   1155
         End
         Begin Threed.SSOption optOrder 
            Height          =   255
            Index           =   0
            Left            =   3855
            TabIndex        =   3
            Top             =   1740
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   450
            _Version        =   262144
            BackColor       =   16777215
            Caption         =   "ASC"
         End
         Begin MSComCtl2.UpDown UpDown 
            Height          =   345
            Left            =   1350
            TabIndex        =   5
            Top             =   1680
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin CSTextLibCtl.silgEdit txtInterval 
            Height          =   360
            Left            =   915
            TabIndex        =   6
            Top             =   1680
            Width           =   435
            _Version        =   262145
            _ExtentX        =   767
            _ExtentY        =   635
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            ReadOnly        =   -1  'True
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "10"
            Text            =   " 10"
            StartText.x     =   3
            StartText.y     =   5
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   13
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   1
            Undo            =   1
            Data            =   10
         End
         Begin CSTextLibCtl.silgEdit txtCount 
            Height          =   360
            Left            =   2295
            TabIndex        =   7
            Top             =   1680
            Width           =   495
            _Version        =   262145
            _ExtentX        =   873
            _ExtentY        =   635
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
            StartText.y     =   5
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   13
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   1
            Undo            =   1
            Data            =   0
         End
         Begin Threed.SSOption optOrder 
            Height          =   255
            Index           =   1
            Left            =   4605
            TabIndex        =   8
            Top             =   1740
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   450
            _Version        =   262144
            BackColor       =   16777215
            Caption         =   "DESC"
            Value           =   -1
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   300
            Index           =   1
            Left            =   3705
            TabIndex        =   18
            Top             =   495
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   529
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   21364739
            UpDown          =   -1  'True
            CurrentDate     =   45291
         End
         Begin VB.Label lblTAGCode 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "999"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   180
            Left            =   4080
            TabIndex        =   21
            ToolTipText     =   "택코드..."
            Top             =   135
            Width           =   315
         End
         Begin VB.Label Label 
            Alignment       =   2  '가운데 맞춤
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
            Height          =   180
            Left            =   3465
            TabIndex        =   20
            Top             =   540
            Width           =   225
         End
         Begin VB.Label lblServer 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "#"
            ForeColor       =   &H00808080&
            Height          =   180
            Left            =   735
            TabIndex        =   16
            Top             =   1260
            Width           =   90
         End
         Begin VB.Image Image3 
            Height          =   285
            Left            =   210
            Picture         =   "frmMain.frx":32E0
            Top             =   90
            Width           =   285
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   0
            X1              =   720
            X2              =   5300
            Y1              =   405
            Y2              =   405
         End
         Begin VB.Label lblBranch 
            BackStyle       =   0  '투명
            Caption         =   "가맹점"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   165
            Left            =   735
            TabIndex        =   13
            Top             =   135
            Width           =   2475
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   1
            X1              =   720
            X2              =   5300
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Label Label3 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "Interval:"
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   12
            Top             =   1770
            Width           =   795
         End
         Begin VB.Image imgConnect 
            Height          =   360
            Left            =   4980
            Picture         =   "frmMain.frx":371E
            Top             =   1125
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Image imgNotConnect 
            Height          =   360
            Left            =   4980
            Picture         =   "frmMain.frx":3E08
            Top             =   1125
            Width           =   360
         End
         Begin VB.Label Label3 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "전송:"
            Height          =   195
            Index           =   1
            Left            =   1740
            TabIndex        =   11
            Top             =   1770
            Width           =   510
         End
         Begin VB.Label lblCode 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "999999"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   180
            Left            =   4620
            TabIndex        =   10
            ToolTipText     =   "가맹점코드..."
            Top             =   135
            Width           =   630
         End
         Begin VB.Label Label3 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "전송순서:"
            Height          =   195
            Index           =   2
            Left            =   2910
            TabIndex        =   9
            Top             =   1770
            Width           =   855
         End
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   3630
         Left            =   750
         TabIndex        =   22
         Top             =   2655
         Width           =   4650
         _Version        =   524288
         _ExtentX        =   8202
         _ExtentY        =   6403
         _StockProps     =   64
         BackColorStyle  =   1
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
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
         MaxCols         =   3
         MaxRows         =   10
         ScrollBars      =   2
         ShadowColor     =   14737632
         SpreadDesigner  =   "frmMain.frx":44F2
         TextTip         =   1
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread sprOffice 
         Height          =   3630
         Left            =   15
         TabIndex        =   23
         Top             =   2655
         Width           =   720
         _Version        =   524288
         _ExtentX        =   1270
         _ExtentY        =   6403
         _StockProps     =   64
         BackColorStyle  =   1
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
         DisplayRowHeaders=   0   'False
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
         MaxRows         =   10
         ScrollBars      =   0
         ShadowColor     =   14737632
         SpreadDesigner  =   "frmMain.frx":4A89
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "mnuTray"
      Visible         =   0   'False
      Begin VB.Menu mnuTrayShow 
         Caption         =   "숨기기/보여주기"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "종료"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Minimized As Boolean
Dim m_bTotal    As Boolean  ' 전체전송
Dim m_bDay      As Boolean  ' 일자전송
Dim m_sOrderBy  As String   ' 전송방법


Dim sValue() As String

Dim Before_Data(1 To 6) As String
Dim Cnt    As Integer

Private Sub btnHide_Click()
    Minimized = True

    Me.Hide
End Sub

Private Function CleanAid_Connect() As Boolean
    Dim sServer   As String
    Dim sDatabase As String
    Dim sID       As String
    Dim sPWD      As String
    
    On Error GoTo ErrRtn

    sServer = Get_Decrypt(GetIniStr("DB", "SERVER", "", iniFile), "")    '
    sDatabase = Get_Decrypt(GetIniStr("DB", "DATABASE", "", iniFile), "") '
    sID = Get_Decrypt(GetIniStr("DB", "ID", "", iniFile), "")             '
    sPWD = Get_Decrypt(GetIniStr("DB", "PWD", "", iniFile), "")            '
    

    
    Set ADOConCleanAid = New ADODB.Connection

    With ADOConCleanAid
        .ConnectionString = "Provider=SQLOLEDB;Persist Security Info=False;User ID=" & sID & ";Password=" & sPWD & ";Initial Catalog=" & sDatabase & ";Data Source=" & sServer
        .CursorLocation = adUseClient
        .ConnectionTimeout = 10
        .CommandTimeout = 30
        .Open
    End With
   
    CleanAid_Connect = True
   
    Exit Function
   
ErrRtn:
    If sprGrid.MaxRows <= 0 Then sprGrid.MaxRows = 1
    sprGrid.Row = sprGrid.MaxRows
    sprGrid.Col = 1: sprGrid.Text = Err.Description
    
    CleanAid_Connect = False
End Function

Private Sub chkDay_Click()
    m_bDay = IIf(chkDay.Value = 1, True, False)
End Sub

Private Sub chkTotal_Click()
    m_bTotal = IIf(chkTotal.Value = 1, True, False)
End Sub

Private Sub Image3_Click()
    Cnt = 10
    Timer1_Timer
End Sub

Private Sub optOrder_Click(Index As Integer, Value As Integer)
    If optOrder(0).Value Then
        m_sOrderBy = " ASC "
    Else
        m_sOrderBy = " DESC "
    End If

End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    If App.PrevInstance = True Then End
    
    If Right(App.Path, 1) = "\" Then
        AppPath = App.Path
    Else
        AppPath = App.Path & "\"
    End If
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    m_sOrderBy = "ASC"
    
    Me.lblVer = "Ver " & App.Major & "-" & App.Minor & "-" & App.Revision
    
    iniFile = AppPath & "\CleanAID.ini" ' 환경 설정 파일의 이름을 설정한다.
    
    ' 신규 업데이트 위치를 수정한다.
    'Call SetIniStr("UPDATE", "URL", "www.clean-aid.co.kr:8090/cleanaid", iniFile)
    
    With sprOffice
        .MaxRows = 0
        .RowHeight(-1) = 14
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeNormal
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
    
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
        .OperationMode = OperationModeNormal
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
        
    Me.Caption = "(주)크린에이드 - Backup System (Ver " & App.Major & "." & App.Minor & "." & App.Revision & ")"
    
    Call 가맹점_Display
    
    Call TrayIcon.ShowBalloonTip(10, "", "", xtpToolTipIconNone)
    
    lblServer.Caption = Get_Decrypt(GetIniStr("SERVER", "SERVER", "", iniFile), "")     '
    
    ' 바로 전송을 시작하도록 하기 위하여
    Cnt = 100
    
    Minimized = True
    
    Me.Hide
    
    Exit Sub

ErrRtn:
    End
End Sub


Private Sub sprGrid_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As FPSpreadADO.TextTipFetchMultilineConstants, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim vText As Variant
    
    sprGrid.GetText Col, Row, vText
    TipText = CStr(vText)

End Sub

Private Sub Timer1_Timer()
    
    Dim iCount     As Integer
    Dim strServer  As String
    Dim strVersion As String
    Dim 접수일자   As String
    
    Dim wmi As Object
    Dim processes, process
    Dim sQuery As String
    
    
    On Error GoTo ERR_RTN
    
    Timer1.Interval = 65535
    
    Cnt = Cnt + 1
    Debug.Print Cnt
    
    

    
    If txtInterval.Value <= Cnt Then
        DoEvents
        
        
        '프로그램 자동업데이트
        If Auto_Update = True Then
            Cnt = 0
            Exit Sub
        End If
        
        ' 다시
        Timer1.Enabled = False
        
        
        ' 3시~4시 사이에 가맹점프로그램 강제 종료 및 전송프로그램 종료 (ver 1.1.15) 2023-09-21
        If Format(Now, "hhmm") >= "0300" And Format(Now, "hhmm") < "0330" Then

            Set wmi = GetObject("winmgmts:")
            sQuery = "select * from win32_process where name = 'CleanAid.exe'"
            Set processes = wmi.execquery(sQuery)

            For Each process In processes
                process.Terminate
            Next

            Set wmi = Nothing

            End
        End If
        
        
        ' 지사 변경이 있을 경우 해당 지사쪽으로 자료 전송을 하기 위하여
        ' 지사 정보를 다시 설정한다.
        Call 가맹점_Display

        ' 이전 서버에 자료를 전송한다.
        ' 해당 일자에 한번만 전송을 한다. 마감자료, 입출고 자료 전송
'        If GetIniStr("OLD_SERVER", Format(Date, "yyyy-MM-dd"), "N", iniFile) <> "Y" Then
'            If OLD_Server_Send = True Then
'                Call SetIniStr("OLD_SERVER", Format(Date, "yyyy-MM-dd"), "Y", iniFile)
'            End If
'        End If

        imgConnect.Visible = False
        imgNotConnect.Visible = True
        sprGrid.MaxRows = 0
        DoEvents
                    
        '---------------------------------------------------------
        ' 2. 공지사항 수신
        '---------------------------------------------------------
        If NewServer_Connection(ADONewServer, "LAUNDRY1000") = True Then
            imgConnect.Visible = True
            imgNotConnect.Visible = False
            
            Call Data_Download      '데이터 수신 (공지사항)
            
            ADONewServer.Close
            Set ADONewServer = Nothing
        End If
                    
        '---------------------------------------------------------
        ' 3. 데이터 전송
        '---------------------------------------------------------
        For iRow = 1 To sprOffice.MaxRows
            sprOffice.Row = iRow
            sprOffice.Col = 1: lblOffice.Caption = sprOffice.Text & ""
            
            If chkCloseDataOnly.Value = 1 Then '마감만
            
                    '---------------------------------------------------------
                    ' LAUNDRY10000
                    '---------------------------------------------------------
                    strServer = "LAUNDRY1000"
                        
                    If NewServer_Connection(ADONewServer, strServer) = True Then
                        imgConnect.Visible = True
                        imgNotConnect.Visible = False
                        
                        txtCount.Value = txtCount.Value + 1
                        DoEvents
                        
                        If CleanAid_Connect = True Then
                                        
                            sprGrid.MaxRows = sprGrid.MaxRows + 1
                            sprGrid.Row = sprGrid.MaxRows:    sprGrid.Col = 1:   sprGrid.Text = "일일마감 정보"
                            Call Send_일일마감(strServer)
                        
                            ADOConCleanAid.Close
                            Set ADOConCleanAid = Nothing
                        End If
                        
                        ADONewServer.Close
                        Set ADONewServer = Nothing
                    End If
            Else
            
                If lblOffice.Caption <> "0000" Then
                    strServer = "LAUNDRY" & lblOffice.Caption
                    
                    If NewServer_Connection(ADONewServer, strServer) = True Then
                        imgConnect.Visible = True
                        imgNotConnect.Visible = False
                        
                        txtCount.Value = txtCount.Value + 1
                        DoEvents
                                        
                        Call Data_Update(strServer) '데이터 전송
                        
                        ADONewServer.Close
                        Set ADONewServer = Nothing
                    End If
                    
                    Cnt = 0
                    
                    imgConnect.Visible = False
                    imgNotConnect.Visible = True
                    DoEvents
                                
                    '---------------------------------------------------------
                    ' LAUNDRY10000
                    '---------------------------------------------------------
                    strServer = "LAUNDRY1000"
                        
                    If NewServer_Connection(ADONewServer, strServer) = True Then
                        imgConnect.Visible = True
                        imgNotConnect.Visible = False
                        
                        txtCount.Value = txtCount.Value + 1
                        DoEvents
                                        
                        Call Data_Update(strServer) '데이터 전송
                        
                        ADONewServer.Close
                        Set ADONewServer = Nothing
                    End If
                End If
            End If
        Next iRow
        
        Cnt = 0
        
        imgConnect.Visible = False
        imgNotConnect.Visible = True
        DoEvents
        
        Timer1.Enabled = True
        
        chkDay.Value = 0
        dtpDay(0).Value = ""
        dtpDay(1).Value = ""
        chkTotal.Value = 0
        chkCloseDataOnly.Value = 0
        
        DoEvents
    End If
    Exit Sub
    
ERR_RTN:
    sprGrid.Row = sprGrid.MaxRows
    sprGrid.Col = 2: sprGrid.Text = Err.Description
    Timer1.Enabled = True

End Sub

Public Sub 마감정보_Send(FinishDate As String)
    Dim iDay        As Integer
    Dim iTempDay    As String
    Dim ADORset     As New ADODB.Recordset
    Dim sData(39)   As String
    Dim sCode(1)    As String
        
    sCode(1) = lblTAGCode.Caption & ""
    
    Query = "SELECT * FROM TB_일일마감"
    Query = Query & " WHERE 마감일자 = '" & Format(FinishDate, "YYYY-MM-DD") & "'"
    Query = Query & "   AND 가맹점코드 = '" & lblCode.Caption & "'"
    Query = Query & "   AND 지사코드 = '" & lblOffice.Caption & "'"
    Set SubRs = New ADODB.Recordset
    SubRs.Open Query, ADOConCleanAid, adOpenForwardOnly, adLockReadOnly
    
    If SubRs.EOF Then
        SubRs.Close
        Set SubRs = Nothing
        
        sData(0) = lblOffice.Caption & ""         '
        sData(1) = Format(FinishDate, "YYYYMMDD") '
        sData(2) = "0"
        sData(3) = "0"
        sData(4) = "0"
        sData(5) = "0"
        sData(6) = "0"
        sData(7) = "0"
        sData(8) = "0"
        sData(9) = "0"
        sData(10) = ""
        sData(11) = ""
        sData(12) = ""
        sData(13) = ""
        sData(14) = ""
        sData(15) = "0"
        sData(16) = "0"
        sData(17) = "0"
        sData(18) = "0"
        sData(19) = "0"
        sData(20) = "0"
        sData(21) = "0"
        sData(22) = "0"
    
        sData(23) = "0"
        sData(24) = "0"
        sData(25) = "0"
        sData(26) = "0"
        sData(27) = "0"
    
        sData(28) = "0"
        sData(29) = "0"
        sData(30) = "0"
        sData(31) = "0"
        sData(32) = "0"
        sData(33) = "0"
        sData(34) = "0"
        sData(35) = "0"
        sData(36) = "0"
    
        sData(37) = "0"
        sData(38) = "0"
        sData(39) = "0"
    Else
        sData(0) = lblOffice.Caption & ""            ' 대리점정보.StoreCode
        sData(1) = Format(SubRs!마감일자, "YYYYMMDD")  '
        sData(2) = SubRs!접수수량 & ""                 ' 총점수
        sData(3) = SubRs!반품수량 & ""                 '
        sData(4) = SubRs!재세탁수량 & ""               '
        sData(5) = SubRs!수선수량 & ""                 '
        sData(6) = SubRs!접수금액 & ""                 ' 총매출액
        sData(7) = SubRs!지사금액 & ""                  ' MasterMoney
        sData(8) = SubRs!가맹점금액 & ""                ' StoreMoney
        sData(9) = SubRs!수선금액 & ""                 '
        sData(10) = SubRs!판매구분 & ""                '
        sData(11) = SubRs!시작택번호 & ""              ' 시작택
        sData(12) = SubRs!종료택번호 & ""              ' 종료택
        sData(13) = SubRs!마감여부 & ""                '
        sData(14) = SubRs!본사전송여부 & ""            ' 전송여부
        sData(15) = SubRs!발생마일리지 & ""            '
        sData(16) = SubRs!사용마일리지 & ""            '
        sData(17) = SubRs!삭제마일리지 & ""            '
        sData(18) = SubRs!카드금액                     '
        sData(19) = SubRs!카드건수                     '
        
        sData(20) = SubRs!운동화건수                   '
        sData(21) = SubRs!운동화금액                   '
        sData(22) = SubRs!운동화비율 & ""               ' 대리점정보.외주운동화마진
        
        sData(23) = SubRs!세탁환불건수                 ' 세탁비환불건수
        sData(24) = SubRs!세탁환불금액                 ' 세탁비환불금액
        
        sData(25) = SubRs!삼성카드할인고객수           '
        sData(26) = SubRs!삼성카드할인건수             '
        sData(27) = SubRs!삼성카드할인금액             '
        
        sData(28) = SubRs!명품세탁건수                 '
        sData(29) = SubRs!명품세탁금액                 '
        sData(30) = SubRs!명품세탁비율                 '
        
        sData(31) = SubRs!명품염색건수                 '
        sData(32) = SubRs!명품염색금액                 '
        sData(33) = SubRs!명품염색비율                 '
        
        sData(34) = 0                                   ' SUBRS!일반가죽건수
        sData(35) = 0                                   ' SUBRS!일반가죽금액
        sData(36) = 0                                   ' SUBRS!일반가죽비율
        
        sData(37) = 0                                   ' SUBRS!가죽건수
        sData(38) = 0                                   ' SUBRS!가죽금액
        sData(39) = 0                                   ' SUBRS!가죽비율
         
        ' 마감되지 않았을 경우 전송하지 않는다.
        If sData(13) <> "Y" Then sData(2) = "0"
        
        SubRs.Close
        Set SubRs = Nothing
    End If

    If sData(6) = "0" Then Exit Sub '접수금액이 0 원이면 전송하지 않는다...
    
    '---------------------------------------------------------------------------
    '
    '---------------------------------------------------------------------------
    Query = "SELECT STORE_CD, TRANS_CHK"
    Query = Query & " FROM SALE_TBL"
    Query = Query & " WHERE Store_CD = '" & lblCode.Caption & "'"           ' 대리점정보.StoreCode
    Query = Query & "   AND SALE_DT  = '" & Format(FinishDate, "yyyyMMdd") & "'"  ' Format(iTempDay, "yyyyMMdd")
    
    If ADORset.State = adStateOpen Then ADORset.Close
    
    ADORset.CursorLocation = adUseClient
    ADORset.Open Query, ADOOldServer, adOpenStatic, adLockBatchOptimistic, adCmdText
    
    If ADORset.EOF = True Then
        Query = "INSERT INTO  SALE_TBL ("
        Query = Query & "  SALE_DT"            '1
        Query = Query & ", STORE_CD"           '2
        Query = Query & ", MASTER_CD"          '3
        Query = Query & ", TAG_NB"             '4
        Query = Query & ", START_TAG"          '5
        Query = Query & ", END_TAG"            '6
        Query = Query & ", SALE_AMT"           '7
        Query = Query & ", MASTER_AMT"         '8
        Query = Query & ", STORE_AMT"          '9
        Query = Query & ", IN_CNT"             '10
        Query = Query & ", JAES_CNT"           '11
        Query = Query & ", SU_CNT"             '12
        Query = Query & ", BAN_CNT"            '13
        Query = Query & ", OUT_CNT"            '14
        Query = Query & ", CARD_AMT"           '15
        Query = Query & ", CARD_CNT"           '16
        Query = Query & ", SU_AMT"             '17
        Query = Query & ", SALE_CHK"           '18
        Query = Query & ", CREATE_MIL"         '19
        Query = Query & ", USE_MIL"            '20
        Query = Query & ", DELETE_MIL"         '21
        Query = Query & ", RunningCnt"         '22
        Query = Query & ", RunningMoney"       '23
        Query = Query & ", RunningPer"         '24
        Query = Query & ", SALERETURN_CNT"     '25
        Query = Query & ", SALERETURN_AMT"     '26
        Query = Query & ", SAMSUNGCARDMEM_CNT" '27
        Query = Query & ", SAMSUNGCARD_CNT"    '28
        Query = Query & ", SAMSUNGCARD_AMT"    '29
        
        ' 20100525 명품관련추가 2줄 추가
        ' 20100722 명품관련추가 1줄 추가
        Query = Query & ", LuxuryLAU_CNT"      '30
        Query = Query & ", LuxuryLAU_AMT"      '31
        Query = Query & ", LuxuryLAU_PER"      '32
        Query = Query & ", LuxuryDYE_CNT"      '33
        Query = Query & ", LuxuryDYE_AMT"      '34
        Query = Query & ", LuxuryDYE_PER"      '35
        Query = Query & ", LuxuryDEF_CNT"      '36
        Query = Query & ", LuxuryDEF_AMT"      '37
        Query = Query & ", LuxuryDEF_PER"      '38
        Query = Query & ", Leather_CNT"        '39
        Query = Query & ", Leather_AMT"        '40
        Query = Query & ", Leather_PER"        '41
        Query = Query & ", TRANS_CHK"          '42
        Query = Query & ", TRANS_DT)"          '43
        
        Query = Query & " VALUES ("
        Query = Query & "  '" & sData(1) & "'"                  ' 1 마감일자
        Query = Query & ", '" & lblCode.Caption & "'"           ' 2 가맹점코드
        Query = Query & ", '" & sData(0) & "'"                  ' 3 지사코드 sCode(0)
        Query = Query & ", '" & sCode(1) & "'"                  ' 4 택코드
        Query = Query & ", '" & Right(sData(11), 4) & "'"       ' 5 START_TAG
        Query = Query & ", '" & Right(sData(11), 4) & "'"       ' 6 END_TAG
        
        Query = Query & ", " & sData(6)                         ' 7 SALE_AMT
        Query = Query & ", " & sData(7)                         ' 8 MASTER_AMT
        
        Query = Query & ", " & sData(8)                         ' 9 STORE_AMT
        Query = Query & ", " & sData(2)                         '10 IN_CNT
        
        Query = Query & ", " & sData(4)                         '11 JAES_CNT
        Query = Query & ", " & sData(5)                         '12 SU_CNT
        
        Query = Query & ", " & sData(3)                         '13 BAN_CNT
        Query = Query & ", " & "0"                              '14 OUT_CNT
        
        Query = Query & ", " & sData(18)                        '15 CARD_AMT
        Query = Query & ", " & sData(19)                        '16 CARD_CNT
        
        Query = Query & ", " & sData(9)                         '17 SU_AMT
        Query = Query & ", '" & sData(10) & "'"                 '18 SALE_CHK
        
        Query = Query & ", " & sData(15)                        '19 CREATE_MIL
        Query = Query & ", " & sData(16)                        '20 USE_MIL
        Query = Query & ", " & sData(17)                        '21 DELETE_MIL
        
        Query = Query & ", " & sData(20)                        '22 RunningCnt
        Query = Query & ", " & sData(21)                        '23 RunningMoney
        Query = Query & ", " & sData(22)                        '24 RunningPer
        
        Query = Query & ", " & sData(23)                        '25 세탁환불 관련
        Query = Query & ", " & sData(24)                        '26
        
        Query = Query & ", " & sData(25)                        '27 삼성카드 관련
        Query = Query & ", " & sData(26)                        '28
        Query = Query & ", " & sData(27)                        '29
        
        Query = Query & ", " & sData(28)                        '30
        Query = Query & ", " & sData(29)                        '31
        Query = Query & ", " & sData(30)                        '32 명품세탁 관련
        
        Query = Query & ", " & sData(31)                        '33 명품염색 관련
        Query = Query & ", " & sData(32)                        '34
        Query = Query & ", " & sData(33)                        '35
        
        Query = Query & ", " & sData(34)                        '36 일반가죽 관련
        Query = Query & ", " & sData(35)                        '37
        Query = Query & ", " & sData(36)                        '38
        
        Query = Query & ", " & sData(37)                        '39 L가죽 관련
        Query = Query & ", " & sData(38)                        '40
        Query = Query & ", " & sData(39)                        '41
        
        Query = Query & ", 'Y'"                                 '42 TRANS_CHK
        Query = Query & ", '" & Format(Date, "yyyyMMdd") & "'"  '43 TRANS_DT
        Query = Query & ")"
        ADOOldServer.Execute Query
    
    Else
        If ADORset.Fields("TRANS_CHK") = "N" Then
            Query = "UPDATE SALE_TBL SET"
            Query = Query & "  MASTER_CD    = '" & sData(0) & "'"            'sCode(0)
            Query = Query & ", TAG_NB       = '" & sCode(1) & "'"
            Query = Query & ", START_TAG    = '" & Right(sData(11), 4) & "'"
            Query = Query & ", END_TAG      = '" & Right(sData(12), 4) & "'"
            Query = Query & ", SALE_AMT     =  " & sData(6)
            Query = Query & ", MASTER_AMT   =  " & sData(7)
            Query = Query & ", STORE_AMT    =  " & sData(8)
            Query = Query & ", IN_CNT       =  " & sData(2)
            Query = Query & ", JAES_CNT     =  " & sData(4)
            Query = Query & ", SU_CNT       =  " & sData(5)
            Query = Query & ", BAN_CNT      =  " & sData(3)
            Query = Query & ", OUT_CNT      =  " & "0"
            Query = Query & ", CARD_AMT     =  " & sData(18)
            Query = Query & ", CARD_CNT     =  " & sData(19)
            Query = Query & ", SU_AMT       =  " & sData(9)
            Query = Query & ", SALE_CHK     = '" & sData(10) & "'"
            Query = Query & ", CREATE_MIL   =  " & sData(15)
            Query = Query & ", USE_MIL      =  " & sData(16)
            Query = Query & ", DELETE_MIL   =  " & sData(17)
            Query = Query & ", RunningCnt   =  " & sData(20)
            Query = Query & ", RunningMoney =  " & sData(21)
            Query = Query & ", RunningPer   =  " & sData(22)
            
            Query = Query & ", SALERETURN_CNT     = " & sData(23)
            Query = Query & ", SALERETURN_AMT     = " & sData(24)
            Query = Query & ", SAMSUNGCARDMEM_CNT = " & sData(25)
            Query = Query & ", SAMSUNGCARD_CNT    = " & sData(26)
            Query = Query & ", SAMSUNGCARD_AMT    = " & sData(27)
            
            ' 20100525 명품관련추가 6줄 추가
            Query = Query & ", LuxuryLAU_CNT = " & sData(28)
            Query = Query & ", LuxuryLAU_AMT = " & sData(29)
            Query = Query & ", LuxuryLAU_PER = " & sData(30)
            Query = Query & ", LuxuryDYE_CNT = " & sData(31)
            Query = Query & ", LuxuryDYE_AMT = " & sData(32)
            Query = Query & ", LuxuryDYE_PER = " & sData(33)
            
            ' 20100722 일반가죽 관련추가 3줄 추가
            Query = Query & ", LuxuryDEF_CNT = " & sData(34)
            Query = Query & ", LuxuryDEF_AMT = " & sData(35)
            Query = Query & ", LuxuryDEF_PER = " & sData(36)
            
            ' 20101015 L가죽 관련추가 3줄 추가
            Query = Query & ", Leather_CNT = " & sData(37)
            Query = Query & ", Leather_AMT = " & sData(38)
            Query = Query & ", Leather_PER = " & sData(39)
            
            Query = Query & ", TRANS_CHK =  'Y'"
            Query = Query & ", TRANS_DT =  '" & Format(Date, "yyyyMMdd") & "'  "
            Query = Query & " WHERE Store_CD = '" & lblCode.Caption & "' "          '대리점정보.StoreCode
            Query = Query & "   AND SALE_DT  = '" & Format(FinishDate, "yyyyMMdd") & "' " 'Format(iTempDay, "yyyyMMdd")
            ADOOldServer.Execute Query
        End If
    End If
End Sub

Private Sub Data_Download()
    Dim Total_Cnt As Long
    Dim nCnt       As Long
    
    On Error Resume Next
    
    ReDim sValue(4)
    
    If CleanAid_Connect = False Then Exit Sub
    
    '===========================================================
    ' 1. 공지사항
    '===========================================================
    sprGrid.MaxRows = sprGrid.MaxRows + 1
    sprGrid.Row = sprGrid.MaxRows
    sprGrid.Col = 1: sprGrid.Text = "공지사항"

    nCnt = 0
        
    sValue(0) = lblCode.Caption
    
    Set ADORs = New ADODB.Recordset
    Set ADORs = ExecPro("SP_SE_00001_SEL", sValue(), Err_Num, Err_Dec)
    
    Do Until ADORs.EOF
        nCnt = nCnt + 1
        
        sprGrid.Row = sprGrid.MaxRows
        sprGrid.Col = 2: sprGrid.Text = Format(nCnt, "#,##0") '현재
        DoEvents
        
        '---------------------------------------------------------
        '
        '---------------------------------------------------------
        Query = "SELECT * FROM TB_공지사항"
        Query = Query & " WHERE 가맹점코드 = '" & ADORs!가맹점코드 & "'"
        Query = Query & "   AND 공지구분   = '" & ADORs!공지구분 & "'"
        Query = Query & "   AND 작성일자   = '" & ADORs!작성일자 & "'"
        Query = Query & "   AND 문서번호   =  " & ADORs!문서번호
        Set SubRs = New ADODB.Recordset
        SubRs.Open Query, ADOConCleanAid, adOpenDynamic, adLockOptimistic
    
        If SubRs.EOF Then SubRs.AddNew
        
        SubRs!가맹점코드 = ADORs!가맹점코드 & ""     '1
        SubRs!공지구분 = ADORs!공지구분 & ""         '2
        SubRs!작성일자 = ADORs!작성일자 & ""         '3
        SubRs!문서번호 = ADORs!문서번호 & ""         '4
        SubRs!시작일자 = ADORs!시작일자 & ""         '5
        SubRs!종료일자 = ADORs!종료일자 & ""         '6
        SubRs!공지내용 = ADORs!공지내용 & ""         '7
        SubRs!수신여부 = ADORs!수신여부 & ""         '8
        SubRs!수신일자 = ADORs!수신일자 & ""         '9
        SubRs!파일명 = ADORs!파일명 & ""             '10
        
        SubRs.Update
        
        SubRs.Close
        Set SubRs = Nothing
        
        ' 서버에 다운로드 내용을 설정한다.
        sValue(0) = lblCode.Caption
        sValue(1) = ADORs!공지구분 & ""
        sValue(2) = ADORs!작성일자 & ""
        sValue(3) = ADORs!문서번호 & ""
        Call ExecPro("SP_SE_00001_INS", sValue(), Err_Num, Err_Dec)
        
        ADORs.MoveNext
    Loop
    ADORs.Close
    Set ADORs = Nothing
        
    '-----------------------------------------------
        
    '===========================================================
    ' 2. 수동입고취소 (ver 1.1.14   2023-09-20)
    '===========================================================


    nCnt = 0

    sValue(0) = lblCode.Caption

    Set ADORs = New ADODB.Recordset
    Set ADORs = ExecPro("SP_SE_00017_INCANCEL", sValue(), Err_Num, Err_Dec)
    
    'Text2.Text = Err_Dec
    If ADORs.EOF = False Then
        sprGrid.MaxRows = sprGrid.MaxRows + 1
        sprGrid.Row = sprGrid.MaxRows
        sprGrid.Col = 1: sprGrid.Text = "입고오류 보정"
    End If
    
    Do Until ADORs.EOF
        nCnt = nCnt + 1

        sprGrid.Row = sprGrid.MaxRows
        sprGrid.Col = 2: sprGrid.Text = Format(nCnt, "#,##0") '현재
        DoEvents

        '---------------------------------------------------------
        '
        '---------------------------------------------------------
        Query = "SELECT * FROM TB_입출고"
        Query = Query & " WHERE 가맹점코드 = '" & ADORs!가맹점코드 & "" & "'"
        Query = Query & "   AND 접수일자   = '" & ADORs!접수일자 & "" & "'"
        Query = Query & "   AND 택번호   = '" & ADORs!택번호 & "" & "'"
        Query = Query & "   AND 접수번호   =  " & ADORs!접수번호 & ""
        Set SubRs = New ADODB.Recordset
        SubRs.Open Query, ADOConCleanAid, adOpenDynamic, adLockOptimistic
        
        'Text1.Text = Query
        
        If Not SubRs.EOF Then

            SubRs!가맹점입고일자 = ""
            SubRs!가맹점입고구분 = ""
            SubRs!본사전송여부 = ""

            SubRs.Update

            SubRs.Close
            Set SubRs = Nothing

        ' 서버에 업데이트 한다.
            sValue(0) = ADORs!가맹점코드 & ""
            sValue(1) = ADORs!접수일자 & ""
            sValue(2) = ADORs!택번호 & ""
            sValue(3) = ADORs!접수번호 & ""
            sValue(4) = ADORs!처리일자 & ""
            Call ExecPro("SP_SE_00017_INCANCEL_UPDATE", sValue(), Err_Num, Err_Dec)
            
        End If
        ADORs.MoveNext
    Loop


    ADORs.Close
    Set ADORs = Nothing
    
    '-----------------------------------------------
    
    '===========================================================
    ' 3. 마일리지 복구 (ver 1.1.18   2024-06-12)
    '===========================================================


    nCnt = 0

    sValue(0) = lblCode.Caption

    Set ADORs = New ADODB.Recordset
    Set ADORs = ExecPro("SP_SE_00018_MILEAGE", sValue(), Err_Num, Err_Dec)
    
    'Text2.Text = Err_Dec
    If ADORs.EOF = False Then
        sprGrid.MaxRows = sprGrid.MaxRows + 1
        sprGrid.Row = sprGrid.MaxRows
        sprGrid.Col = 1: sprGrid.Text = "마일리지 오류 보정"
    End If
    
    
    Do Until ADORs.EOF
        nCnt = nCnt + 1

        sprGrid.Row = sprGrid.MaxRows
        sprGrid.Col = 2: sprGrid.Text = Format(nCnt, "#,##0") '현재
        DoEvents
        
        ' 전체 마일리지 삭제 처리
        If UCase(ADORs!고객코드) = "ALL" And ADORs!접수번호 = 0 And ADORs!복구마일리지 = 0 Then
            
            sValue(0) = ADORs!가맹점코드 & ""
            sValue(1) = ADORs!고객코드 & ""
            sValue(2) = ADORs!접수번호 & ""
            sValue(3) = ADORs!처리일자 & ""
            
            Call Set_마일리지삭제
            
            ' 서버에 업데이트 한다.

                Call ExecPro("SP_SE_00018_MILEAGE_UPDATE", sValue(), Err_Num, Err_Dec)
        Else
        
            '---------------------------------------------------------
            '
            '---------------------------------------------------------
            Query = "SELECT * FROM TB_매출 a"
            Query = Query & " WHERE 가맹점코드 = '" & ADORs!가맹점코드 & "" & "'"
            Query = Query & "   AND 고객코드   = '" & ADORs!고객코드 & "" & "'"
            Query = Query & "   AND 접수번호   = '" & ADORs!접수번호 & "" & "'"
            Query = Query & "   AND 일련번호   =  (SELECT MIN(일련번호) FROM TB_매출 WHERE 가맹점코드 = a.가맹점코드 AND 고객코드 = a.고객코드 AND 접수번호 = a.접수번호 AND 접수금액 < 0) "
            Set SubRs = New ADODB.Recordset
            SubRs.Open Query, ADOConCleanAid, adOpenDynamic, adLockOptimistic
            
            'Text1.Text = Query
            
            If Not SubRs.EOF Then
    
                SubRs!사용마일리지 = ADORs!복구마일리지 * -1
                SubRs!본사전송여부 = ""
    
                SubRs.Update
    
                SubRs.Close
                Set SubRs = Nothing
                
                
                Query = "SELECT * FROM TB_고객정보 a"
                Query = Query & " WHERE 가맹점코드 = '" & ADORs!가맹점코드 & "" & "'"
                Query = Query & "   AND 고객코드   = '" & ADORs!고객코드 & "" & "'"
                Set SubRs = New ADODB.Recordset
                SubRs.Open Query, ADOConCleanAid, adOpenDynamic, adLockOptimistic
                            
                If Not SubRs.EOF Then
                    SubRs!사용가능마일리지 = SubRs!사용가능마일리지 + (ADORs!복구마일리지)
                    SubRs!본사전송여부 = ""
        
                    SubRs.Update
        
                    SubRs.Close
                    Set SubRs = Nothing
                End If
                
                
    
            ' 서버에 업데이트 한다.
                sValue(0) = ADORs!가맹점코드 & ""
                sValue(1) = ADORs!고객코드 & ""
                sValue(2) = ADORs!접수번호 & ""
                sValue(3) = ADORs!처리일자 & ""
                Call ExecPro("SP_SE_00018_MILEAGE_UPDATE", sValue(), Err_Num, Err_Dec)
                
            End If
        End If
        ADORs.MoveNext
    Loop


    ADORs.Close
    Set ADORs = Nothing
    
    
    '===========================================================
    ' 4. 마감지사 변경 (ver 1.1.19   2024-09-02)
    '===========================================================

    nCnt = 0

    sValue(0) = lblCode.Caption

    Set ADORs = New ADODB.Recordset
    Set ADORs = ExecPro("SP_SE_00019_CLOSEBRANCHCHANGE", sValue(), Err_Num, Err_Dec)
    
    'Text2.Text = Err_Dec
    
    If ADORs.EOF = False Then
        sprGrid.MaxRows = sprGrid.MaxRows + 1
        sprGrid.Row = sprGrid.MaxRows
        sprGrid.Col = 1: sprGrid.Text = "마감 지사 변경"
    End If
    
    Do Until ADORs.EOF
        nCnt = nCnt + 1

        sprGrid.Row = sprGrid.MaxRows
        sprGrid.Col = 2: sprGrid.Text = Format(nCnt, "#,##0") '현재
        DoEvents

        '---------------------------------------------------------
        '
        '---------------------------------------------------------
        Query = "SELECT * FROM TB_일일마감"
        Query = Query & " WHERE 가맹점코드 = '" & ADORs!가맹점코드 & "" & "'"
        Query = Query & "   AND 마감일자   = '" & ADORs!마감일자 & "" & "'"
        Query = Query & "   AND 지사코드   = '" & ADORs!지사코드 & "" & "'"
        
        Set SubRs = New ADODB.Recordset
        SubRs.Open Query, ADOConCleanAid, adOpenDynamic, adLockOptimistic
        
        'Text1.Text = Query
        
        If Not SubRs.EOF Then

            SubRs!지사코드 = ADORs!변경지사코드
            SubRs!본사전송여부 = ""

            SubRs.Update

            SubRs.Close
            Set SubRs = Nothing

        ' 서버에 업데이트 한다.
            sValue(0) = ADORs!가맹점코드 & ""
            sValue(1) = ADORs!지사코드 & ""
            sValue(2) = ADORs!마감일자 & ""
            sValue(3) = ADORs!처리일자 & ""
            Call ExecPro("SP_SE_00019_CLOSEBRANCHCHANGE_UPDATE", sValue(), Err_Num, Err_Dec)
            
        End If
        ADORs.MoveNext
    Loop


    ADORs.Close
    Set ADORs = Nothing
    
    
    '-----------------------------------------------
    
    ADOConCleanAid.Close
    Set ADOConCleanAid = Nothing
    
    DoEvents
    Exit Sub
    
    
ERR_RTN:
    sprGrid.Row = sprGrid.MaxRows
    sprGrid.Col = 2: sprGrid.Text = Err.Description

End Sub

Public Sub Set_마일리지삭제()
    On Error GoTo ErrRtn
    
    Dim iSEQ         As Long
    Dim 삭제마일리지 As Long
    
    Query = "SELECT * FROM TB_고객정보"
    Query = Query & " WHERE 사용가능마일리지 > 0 or 누적마일리지 > 0 "
    Set ADORs2 = New ADODB.Recordset
    ADORs2.Open Query, ADOConCleanAid, adOpenForwardOnly, adLockReadOnly

    Do Until ADORs2.EOF
        DoEvents
        
        Query = "SELECT ISNULL(MAX(일련번호),0) + 1"
        Query = Query & " FROM TB_매출"
        Query = Query & " WHERE 고객코드 = '" & ADORs!고객코드 & "'"
        Query = Query & "   AND 접수번호 = 0"
        Set SubRs = New ADODB.Recordset
        SubRs.Open Query, ADOConCleanAid, adOpenForwardOnly, adLockReadOnly
                        
        iSEQ = SubRs(0)
        SubRs.Close:    Set SubRs = Nothing
            
        삭제마일리지 = ADORs2!사용가능마일리지 + ADORs2!누적마일리지
        
        ADOConCleanAid.BeginTrans
        
        Query = "INSERT INTO TB_매출 ( 고객코드"          ' 1
        Query = Query & "            , 접수번호"          ' 2
        Query = Query & "            , 일련번호"          ' 3
        Query = Query & "            , 매출일자"          ' 4
        Query = Query & "            , 매출시간"          ' 5
        Query = Query & "            , 적요"              ' 6
        Query = Query & "            , 접수금액"          ' 7
        Query = Query & "            , 입금합계"          ' 8
        Query = Query & "            , 현금입금"          ' 9
        Query = Query & "            , 카드입금"          '10
        Query = Query & "            , 쿠폰입금"          '11
        Query = Query & "            , 쿠폰번호"          '12
        Query = Query & "            , 사용마일리지"      '13
        Query = Query & "            , 세트할인"          '14
        Query = Query & "            , 에누리"            '15
        Query = Query & "            , 접수수량"          '16
        Query = Query & "            , 반품수량"          '17
        Query = Query & "            , 발생마일리지"      '18
        Query = Query & "            , 누적마일리지"      '19
        Query = Query & "            , 사용가능마일리지"  '20
        Query = Query & "            , 삭제마일리지"      '21
        Query = Query & "            , 가맹점코드"        '22
        Query = Query & "            , 지사코드"          '23
        Query = Query & "            , 본사전송여부"      '
        Query = Query & "            ) VALUES ("
        Query = Query & "              '" & ADORs!고객코드 & "'"             ' 1
        Query = Query & "            , 0"                                    ' 2
        Query = Query & "            ,  " & iSEQ                             ' 3
        Query = Query & "            , '" & Format(Date, "YYYY-MM-DD") & "'" ' 4
        Query = Query & "            , '" & Format(Time, "hh:mm:ss") & "'"   ' 5
        Query = Query & "            , '[마일리지 삭제]'"                    ' 6
        Query = Query & "            , 0"                                    ' 7
        Query = Query & "            , 0"                                    ' 8
        Query = Query & "            , 0"                                    ' 9
        Query = Query & "            , 0"                                    '10
        Query = Query & "            , 0"                                    '11
        Query = Query & "            , 0"                                    '12
        Query = Query & "            , 0"                                    '13
        Query = Query & "            , 0"                                    '14
        Query = Query & "            , 0"                                    '15
        Query = Query & "            , 0"                                    '16
        Query = Query & "            , 0"                                    '17
        Query = Query & "            , 0"                                    '18
        Query = Query & "            , 0"                                    '19
        Query = Query & "            , 0"                                    '20
        Query = Query & "            ,  " & 삭제마일리지                     '21
        Query = Query & "            , '" & lblCode.Caption & "'"      '22
        Query = Query & "            , '" & lblOffice.Caption & "'"        '23
        Query = Query & "            , '')"
        ADOConCleanAid.Execute Query
        
        ' 고객정보를 수정한다.
        Query = "UPDATE TB_고객정보 SET 사용가능마일리지 = 0, 누적마일리지 = 0, 본사전송여부 = 'N' WHERE 고객코드 = '" & ADORs!고객코드 & "'"
        ADOConCleanAid.Execute Query
        
        ADOConCleanAid.CommitTrans
        
        ADORs2.MoveNext
    Loop
    
    ADORs2.Close
    Set ADORs2 = Nothing

    Exit Sub
    
ErrRtn:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Set_마일리지삭제 of frmMain"
    ADOConCleanAid.RollbackTrans

End Sub

Private Sub 가맹점_Display()
    If CleanAid_Connect = False Then Exit Sub
        
    Query = "SELECT * FROM TB_기본정보"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOConCleanAid, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF Then
        lblBranch.Caption = "테스트"
        lblCode.Caption = "999999"
        lblOffice.Caption = "9999"
        lblTAGCode.Caption = "999"
    Else
        lblBranch.Caption = Trim(ADORs!가맹점명) & ""
        lblCode.Caption = Trim(ADORs!가맹점코드) & ""
        lblOffice.Caption = Trim(ADORs!지사코드) & ""
        lblTAGCode.Caption = Trim(ADORs!택코드) & ""
    End If
    ADORs.Close
    Set ADORs = Nothing
    
    ' 2012-10-24일 수정
    ' 모두 전송하였을 경우 현재의지사만을 등록 한다.
    
    Query = "SELECT DISTINCT 지사코드 FROM TB_입출고 WHERE 본사전송여부 <> 'Y'"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOConCleanAid, adOpenForwardOnly, adLockReadOnly
    
    With sprOffice
        .MaxRows = 0
        .ReDraw = False
        
        If ADORs.EOF Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = lblOffice.Caption
 
        
        Else
            Do Until ADORs.EOF
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                
                .Col = 1: .Text = ADORs!지사코드 & ""
                
                ADORs.MoveNext
            Loop
            ADORs.Close
        End If
        Set ADORs = Nothing
        
        .ReDraw = True
    End With
 
    ADOConCleanAid.Close
    Set ADOConCleanAid = Nothing
End Sub

Private Sub UpDown_DownClick()
    If txtInterval.Value <= 0 Then
        txtInterval.Value = 0
    Else
        txtInterval.Value = txtInterval.Value - 1
    End If
End Sub

Private Sub UpDown_UpClick()
    If txtInterval.Value >= 60 Then
        txtInterval.Value = 60
    Else
        txtInterval.Value = txtInterval.Value + 1
    End If
End Sub



Private Sub Data_Update(LaundryDB As String)
    
    On Error GoTo ERR_RTN
    
    ' CleanAid 연결이 안되면 바로 종료 한다.
    If CleanAid_Connect = False Then Exit Sub
    
    If LaundryDB = "LAUNDRY1000" Then
        With sprGrid
            '===========================================================
            ' 1. 접속 정보 전송
            '===========================================================
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows:    .Col = 1:   .Text = "접속 정보"
            Call Send_접속정보(LaundryDB)
            
            '===========================================================
            ' 1. 체인점 정보
            '===========================================================
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows:    .Col = 1:   .Text = "가맹점 정보"
            Call Send_가맹점정보(LaundryDB)
            
            '===========================================================
            ' 2. 고객
            '===========================================================
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows:    .Col = 1:   .Text = "고객 정보"
            Call Send_고객정보(LaundryDB)
            
            '===========================================================
            ' 5.TB_일일마감
            '===========================================================
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows:    .Col = 1:   .Text = "일일마감 정보"
            Call Send_일일마감(LaundryDB)
            
            '===========================================================
            ' 11. 사고품내역
            '===========================================================
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows:    .Col = 1:   .Text = "사고품내역 정보"
            Call Send_사고품(LaundryDB)
        End With
    
    '******************************************************************************
    '*
    '* 해당 지사 LAUNDRYXXXX DB에만 저장하고 LAUNDRY1000 (본사)에는 저장하지 않는다.
    '*
    '******************************************************************************
    Else
        With sprGrid
            '===========================================================
            ' * 본사에서 환불확정처리건을 수신
            '===========================================================
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows:    .Col = 1:   .Text = "환불확정처리"
            Call Recv_환불확정처리
        
            '===========================================================
            ' 3.TB_입출고
            '===========================================================
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows:    .Col = 1:   .Text = "입출고 정보"
            Call Send_입출고
            
            '-----------------------------------------------------------
            ' 4.TB_매출
            '-----------------------------------------------------------
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows:    .Col = 1:   .Text = "매출 정보"
            Call Send_매출
    
            '===========================================================
            ' 6. 현금영수증
            '===========================================================
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows:    .Col = 1:   .Text = "현금영수증 정보"
            Call Send_현금영수증(LaundryDB)
            
            '===========================================================
            ' 7. 신용카드승인
            '===========================================================
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows:    .Col = 1:   .Text = "신용카드승인 정보"
            Call Send_신용카드승인(LaundryDB)

            '===========================================================
            ' 8. 이용실적
            '===========================================================
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows:    .Col = 1:   .Text = "이용실적 정보"
            Call Send_이용실적(LaundryDB)
            
            '===========================================================
            ' 9. 가맹점입금
            '===========================================================
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows:    .Col = 1:   .Text = "가맹점입금 정보"
            Call Send_가맹점입금(LaundryDB)
            
            '===========================================================
            ' 10. 부자재주문
            '===========================================================
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows:    .Col = 1:   .Text = "부자재주문 정보"
            Call Send_부자재주문(LaundryDB)
            
            '===========================================================
            ' 11. 쿠폰자료 전송
            '===========================================================
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows:    .Col = 1:   .Text = "쿠폰자료 정보"
            Call Send_쿠폰자료(LaundryDB)
        
            '===========================================================
            ' 12. 미수금수정 내역
            '===========================================================
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows:    .Col = 1:   .Text = "미수금수정 정보"
            Call Send_미수금수정(LaundryDB)
        
        End With
    End If
    
    ADOConCleanAid.Close
    Set ADOConCleanAid = Nothing
    
    DoEvents
    Exit Sub

ERR_RTN:
    sprGrid.Row = sprGrid.MaxRows
    sprGrid.Col = 2: sprGrid.Text = Err.Description

End Sub

Private Sub TrayIcon_DblClick()
    If (Minimized) Then
        Call MinimizeToTray
    End If
End Sub

Private Sub TrayIcon_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Button = 2) Then Me.PopupMenu mnuTray
End Sub

Private Sub MinimizeToTray()
    If Not Minimized Then
        TrayIcon.MinimizeToTray Me.hwnd
        Minimized = True
    Else
        TrayIcon.MaximizeFromTray Me.hwnd
        Minimized = False
    End If
    
End Sub

Private Sub mnuExit_Click()
    
    Unload Me
End Sub

Private Sub mnuTrayShow_Click()
    Call MinimizeToTray
End Sub

Private Function Auto_Update() As Boolean
    Dim strVersion As String
    Dim InfoText   As String
    
    Auto_Update = False
    
    strVersion = Format(App.Major, "00") & "." & Format(App.Minor, "00") & "." & Format(App.Revision, "00") '
    
    Load frmUpdateCheck
    'frmUpdateCheck.Show
    
    frmUpdateCheck.lblVersion.Caption = "Ver. " & strVersion & ""
    'frmUpdateCheck.Refresh
    
    i = frmUpdateCheck.SmartUpdateX.GetInfo 'Geturl 에서 버전 정보화일을 읽어옵니다
  
    Debug.Print "UPDATE"
    
    If i = 1 Then '접속 성공 - 업데이트 정보 가저옴
        InfoText = frmUpdateCheck.SmartUpdateX.GetInfoText '서버에서 가저온 IniFileName 모든 내용을 보여줍니다
     
        Open AppPath & "Update.ini" For Output As #1
    
        Print #1, InfoText
        Close #1
        
        If strVersion < frmUpdateCheck.SmartUpdateX.ReadInfo("VER", "INFO2") Then
            frmUpdateCheck.lblNewVersion.Caption = "New Ver. " & frmUpdateCheck.SmartUpdateX.ReadInfo("VER", "INFO2") & ""
            
            Timer1.Enabled = False
            Auto_Update = True
            DoEvents
            
            Call frmUpdateCheck.SmartUpdate 'DBUpdate.EXE 업데이트
        Else
            Unload frmUpdateCheck
        End If
    Else
        Unload frmUpdateCheck
    End If
End Function


'====================================================================================================
' Procedure : SendTable_입출고
' DateTime  : 2008-05-03 01:42
' Author    : pds2004
' Last Edit :
' Return Val:
'----------------------------------------------------------------------------------------------------
' 설     명 : 전달된 일자부터 현재일자까지의 입출고 내요을 본사 SQL 서버에 저장한다.
'====================================================================================================
Private Sub SendTable_입출고(ByVal sStartDate As String)
    Dim 택번호       As String
    
    Dim 지사출고일자 As String
    Dim 반품환불일자 As String
    
    Dim 본출         As String
    Dim 상태         As String
    Dim 확인         As String
    
    On Error GoTo ERR_RTN
         
    Query = "SELECT * FROM TB_입출고"
    Query = Query & " WHERE (접수일자 >= '" & Format(sStartDate, "YYYY-MM-DD") & "')"
    Query = Query & "   AND ((SUBSTRING(의류코드,1,1) IN ('a','o','p','w','b','n','l') OR 세탁환불일자 <> '')"
    Query = Query & "    OR SUBSTRING(세탁환불일자,1,10) = '" & Format(sStartDate, "YYYY-MM-DD") & "')"
    Query = Query & "   AND 가맹점코드 = '" & lblCode.Caption & "'"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOConCleanAid, adOpenForwardOnly, adLockReadOnly
    
    Do Until ADORs.EOF
        택번호 = Format(Right(ADORs!택번호, 4), "0-000")
        
        If Trim(ADORs!지사출고일자) = "" Then
            지사출고일자 = ""
        Else
            지사출고일자 = Format(Left(ADORs!지사출고일자, 10), "YYYYMMDD")
        End If
        
        If Trim(ADORs!반품환불일자) = "" Then
            반품환불일자 = ""
        Else
            반품환불일자 = Replace(ADORs!반품환불일자, "-", "")
            반품환불일자 = Replace(반품환불일자, ":", "")
            반품환불일자 = Replace(반품환불일자, " ", "")
        End If
        
        If Trim(ADORs!지사출고일자) = "" Then
            본출 = ""
        Else
            본출 = "出"
        End If
        
        If Trim(ADORs!출고일자) = "" Then
            상태 = "未"
            확인 = ""
        Else
            상태 = "完"
            확인 = "확"
        End If
        
        Query = "EXEC PRO_A_00005"
        Query = Query & "  '" & lblCode.Caption & "'"                    ' 1
        Query = Query & ", '" & Format(ADORs!접수일자, "YYYYMMDD") & "'" ' 2
        Query = Query & ", '" & 택번호 & "'"                             ' 3
        Query = Query & ", '" & ADORs!고객코드 & "'"                     ' 4
        Query = Query & ", '" & ADORs!의류코드 & "'"                     ' 5
        Query = Query & ", '" & ADORs!의류명 & "'"                       ' 6
        Query = Query & ", '" & ADORs!색상 & "'"                         ' 7
        Query = Query & ", '" & ADORs!내용 & "'"                         ' 8
        Query = Query & ", '" & ADORs!금액 & "'"                         ' 9
        Query = Query & ", '" & ADORs!상표 & "'"                         '10
        Query = Query & ", '" & 본출 & "'"                               '11 ADORs!본출
        Query = Query & ", '" & 상태 & "'"                               '12 상태
        Query = Query & ", '" & 확인 & "'"                               '13 ADORs!확인
        Query = Query & ", '" & Format(ADORs!출고일자, "YYYYMMDD") & "'" '14
        Query = Query & ", '" & Trim(ADORs!판매취소) & "'"               '15
        Query = Query & ", '" & Format(ADORs!예정일자, "YYYYMMDD") & "'" '16 입고예정일
        Query = Query & ", '" & 반품환불일자 & "'"                       '17 환불일자
        Query = Query & ",  " & ADORs!수선금액                           '18
        Query = Query & ", '" & 지사출고일자 & "'"                       '19 본출일자
        Query = Query & ", '" & Trim(ADORs!가맹점입고구분) & "'"         '20 본출입고구분
        Query = Query & ",  " & ADORs!외주마진                           '21 외주운동화마진
        Query = Query & ", '" & ADORs!세탁환불일자 & "'"                 '22
        Query = Query & ", '" & lblOffice.Caption & "'"                  '23
        ADOConCleanAid.Execute Query
        
        ADORs.MoveNext
    Loop
    
    ADORs.Close
    Set ADORs = Nothing
    
    Exit Sub

ERR_RTN:
    sprGrid.Row = sprGrid.MaxRows
    sprGrid.Col = 2: sprGrid.Text = Err.Description

End Sub


Private Function OLD_Server_Send() As Boolean
    Dim 전송여부   As String
    Dim 마감일자   As String
    Dim 접수일자   As String
    
    On Error GoTo ERR_RTN
    OLD_Server_Send = False
        
    If ConnectOldServerCheck(ADOOldServer) = True Then
        For iRow = 1 To sprOffice.MaxRows
            sprOffice.Row = iRow
            sprOffice.Col = 1: lblOffice.Caption = sprOffice.Text & ""
        
            '---------------------------------------------------------
            ' 1. 마감정보 송신
            '---------------------------------------------------------
            ' 마감일자 = "2010-12-15"
             
             마감일자 = Format(DateAdd("d", -10, Date), "yyyy-MM-dd")
            If CleanAid_Connect = True Then
                imgConnect.Visible = True
                imgNotConnect.Visible = False
                DoEvents
                
                Do Until 마감일자 > Format(Date, "YYYY-MM-DD")
                    Call 마감정보_Send(마감일자)    '
                
                    마감일자 = Format(DateAdd("d", 1, 마감일자), "YYYY-MM-DD")
                Loop
            End If
            ADOConCleanAid.Close
            Set ADOConCleanAid = Nothing
            '---------------------------------------------------------
            
            '---------------------------------------------------------
            ' 2. 입출고 송신
            '---------------------------------------------------------
            전송여부 = GetIniStr("WORK", "INOUT", "", iniFile)
            
            If 전송여부 = "" Then
                '접수일자 = "2010-10-01"
                 접수일자 = Format(DateAdd("d", -5, Date), "yyyy-MM-dd")
            Else
                접수일자 = Format(Date, "YYYY-MM-DD")
            End If
            
            If CleanAid_Connect = True Then
                imgConnect.Visible = True
                imgNotConnect.Visible = False
                DoEvents
                
                Do Until 접수일자 > Format(Date, "YYYY-MM-DD")
                    Call SendTable_입출고(접수일자) '
                
                    접수일자 = Format(DateAdd("d", 1, 접수일자), "YYYY-MM-DD")
                Loop
            End If
            ADOConCleanAid.Close
            Set ADOConCleanAid = Nothing
            
            If 전송여부 = "" Then Call SetIniStr("WORK", "INOUT", "Y", iniFile)
        
        Next iRow
        
        ADOOldServer.Close
        Set ADOOldServer = Nothing
    End If
    OLD_Server_Send = True
    Exit Function
    
ERR_RTN:
    sprGrid.Row = sprGrid.MaxRows
    sprGrid.Col = 2: sprGrid.Text = Err.Description
    
    OLD_Server_Send = False
End Function

Private Function Send_접속정보(LaundryDB As String) As Boolean
    ReDim sValue(24)

    On Error GoTo ERR_RTN
    Send_접속정보 = False
        
    sprGrid.Row = sprGrid.MaxRows
    sprGrid.Col = 2: sprGrid.Text = Format(1, "#,##0") '현재
    DoEvents
        
    sValue(0) = lblCode.Caption
    sValue(1) = App.Major & "." & App.Minor & "." & App.Revision
    sValue(2) = m_id
    
    
        
    Call ExecPro("SP_SE_00016_INS", sValue(), Err_Num, Err_Dec)
    
    Send_접속정보 = False
    Exit Function
    
ERR_RTN:
    Send_접속정보 = True
    sprGrid.Row = sprGrid.MaxRows
    sprGrid.Col = 2: sprGrid.Text = Err.Description
End Function



Private Function Send_가맹점정보(LaundryDB As String) As Boolean
    Dim nCnt    As Long
    
    On Error GoTo ERR_RTN
    Send_가맹점정보 = False
    nCnt = 0
    
    Query = "SELECT * FROM TB_기본정보"
    Query = Query & " WHERE (본사전송여부 <> 'Y' OR 본사전송여부 IS NULL)"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOConCleanAid, adOpenForwardOnly, adLockReadOnly
    
    Do Until ADORs.EOF
    
    
        If Trim(ADORs!가맹점코드 & "") <> "" And Len(Trim(ADORs!가맹점코드 & "")) = 6 Then
        
            nCnt = nCnt + 1
            
            sprGrid.Row = sprGrid.MaxRows
            sprGrid.Col = 2: sprGrid.Text = Format(nCnt, "#,##0") '현재
            
            lblBranch.Caption = Trim(ADORs!가맹점명) & "" '
            lblCode.Caption = Trim(ADORs!가맹점코드) & "" '
            lblTAGCode.Caption = Trim(ADORs!택코드) & ""  '
            DoEvents
            
            '---------------------------------------------------------
            '
            '---------------------------------------------------------
            Query = "SELECT * FROM TB_가맹점"
            Query = Query & " WHERE 가맹점코드 = '" & lblCode.Caption & "'"
            Set SubRs = New ADODB.Recordset
            SubRs.Open Query, ADONewServer, adOpenDynamic, adLockOptimistic
        
            If SubRs.EOF Then SubRs.AddNew
            ' 해당 가맹점에서 수정할 수 있는 내용만 전송을 한다.
            
'            SubRs!지사코드 = ADORs!지사코드 & ""             '1
'            SubRs!가맹점코드 = ADORs!가맹점코드 & ""         '2
'            SubRs!가맹점명 = ADORs!가맹점명 & ""             '3
'            SubRs!가맹점구분 = ADORs!가맹점구분 & ""         '4
'            SubRs!적용일자 = ADORs!적용일자 & ""             '5
'            SubRs!택코드 = ADORs!택코드 & ""                 '6
'            SubRs!택색상 = ADORs!택색상 & ""                 '7
             SubRs!택번호 = ADORs!택번호 & ""                 '8
             SubRs!접수번호 = ADORs!접수번호 & ""             '9
'            SubRs!요일할인 = ADORs!요일할인 & ""             '10
'            SubRs!세트상품세일 = ADORs!세트상품세일 & ""     '11
'            SubRs!세탁소요일 = ADORs!세탁소요일 & ""         '12
'            SubRs!SMS_IP = ADORs!SMS_IP & ""                 '13
'            SubRs!SMS_DB = ADORs!SMS_DB & ""                 '14
'            SubRs!SMS_ID = ADORs!SMS_ID & ""                 '15
'            SubRs!SMS_PWD = ADORs!SMS_PWD & ""               '16
'            SubRs!TimeOut = ADORs!TimeOut & ""               '17
            
            SubRs!프로그램버전 = ADORs!프로그램버전 & ""                           '18 가맹점 프로그램 버전
            SubRs!업데이트버전 = App.Major & "." & App.Minor & "." & App.Revision  '   DBUPDATE 프로그램 버전
            
'            SubRs!매장전화번호 = ADORs!매장전화번호 & ""     '19
'            SubRs!문자발신전화 = ADORs!문자발신전화 & ""     '20
'            SubRs!휴대전화번호 = ADORs!휴대전화번호 & ""     '21
'            SubRs!SMS_EMART = ADORs!SMS_EMART & ""           '22
'            SubRs!외주마진 = ADORs!외주마진 & ""             '23
'            SubRs!특정할인여부 = ADORs!특정할인여부 & ""     '24
'            SubRs!특정할인비율 = ADORs!특정할인비율 & ""     '25
'            SubRs!특정할인시작일 = ADORs!특정할인시작일 & "" '26
'            SubRs!특정할인종료일 = ADORs!특정할인종료일 & "" '27
'
'            SubRs!쿠폰할인여부 = ADORs!쿠폰할인여부 & ""     '28
'            SubRs!쿠폰할인비율 = ADORs!쿠폰할인비율 & ""     '29
'            SubRs!쿠폰할인시작일 = ADORs!쿠폰할인시작일 & "" '30
'            SubRs!쿠폰할인종료일 = ADORs!쿠폰할인종료일 & "" '31
'
'            SubRs!지정할인여부 = ADORs!지정할인여부 & ""     '32
'            SubRs!지정할인비율 = ADORs!지정할인비율 & ""     '33
'            SubRs!지정할인시작일 = ADORs!지정할인시작일 & "" '34
'            SubRs!지정할인종료일 = ADORs!지정할인종료일 & "" '35
'
'            SubRs!고가세탁비율 = ADORs!고가세탁비율 & ""     '36
'            SubRs!세탁환불여부 = ADORs!세탁환불여부 & ""     '37
'            SubRs!마일리지여부 = ADORs!마일리지여부 & ""     '38
'            SubRs!기준금액 = ADORs!기준금액 & ""             '39
'            SubRs!적립마일리지 = ADORs!적립마일리지 & ""     '40
'            SubRs!최소마일리지 = ADORs!최소마일리지 & ""     '41

            SubRs!VAN_IP = ADORs!VAN_IP & ""                 '42
            SubRs!VAN_PORT = ADORs!VAN_PORT & ""             '43
            SubRs!사업자번호 = ADORs!사업자번호 & ""         '44
            SubRs!단말기번호 = ADORs!단말기번호 & ""         '45
            SubRs!사업장주소 = ADORs!사업장주소 & ""         '46
            SubRs!대표자명 = ADORs!대표자명 & ""             '47
            SubRs!업태 = ADORs!업태 & ""                     '48
            SubRs!종목 = ADORs!종목 & ""                     '49
            SubRs!우편번호 = ADORs!우편번호 & ""             '50
            SubRs!주소 = ADORs!주소 & ""                     '51
            
'            SubRs!담당자코드 = ADORs!담당자코드 & ""         '52
'            SubRs!기사코드 = ADORs!기사코드 & ""             '53
'            SubRs!계약일자 = ADORs!계약일자 & ""             '54
'            SubRs!해지일자 = ADORs!해지일자 & ""             '55
'            SubRs!이전가맹점코드 = ADORs!이전가맹점코드 & "" '56
'            SubRs!가맹점상태 = ADORs!가맹점상태 & ""         '57
            SubRs!비밀번호 = ADORs!비밀번호 & ""             '58
            SubRs!본사수신일자 = ADORs!본사수신일자 & ""     '59
            SubRs!본사전송여부 = "Y"                                '60 ADORs!본사전송여부
            SubRs!본사전송일자 = Format(Now, "YYYY-MM-DD hh:mm:ss") '61 ADORs!본사전송일자
            
            SubRs.Update
            SubRs.Close:    Set SubRs = Nothing
            
            '-------------------------------------------------------------------------------------------
            ' TB_기본정보 - 전송여부
            '-------------------------------------------------------------------------------------------
            Query = "UPDATE TB_기본정보 SET 본사전송여부 = 'Y'"
            Query = Query & "             , 본사전송일자 = '" & Format(Now, "YYYY-MM-DD hh:mm:ss") & "'"
            ADOConCleanAid.Execute Query
        End If
        
        ADORs.MoveNext
    Loop
    
    ADORs.Close:    Set ADORs = Nothing
    Send_가맹점정보 = True
    Exit Function
    
ERR_RTN:
    Send_가맹점정보 = False
    sprGrid.Row = sprGrid.MaxRows
    sprGrid.Col = 2: sprGrid.Text = Err.Description

End Function

Private Function Send_고객정보(LaundryDB As String) As Boolean
    Dim nCnt As Long
    ReDim sValue(29)
    
    On Error GoTo ERR_RTN
    Send_고객정보 = False
    nCnt = 0
    
'    If (m_bTotal = True) Or (m_bDay = True) Then
'        Query = "SELECT * FROM TB_고객정보 "
'        Query = Query & "   where 가맹점코드 = '" & lblCode.Caption & "'"
'    Else
'        Query = "SELECT * FROM TB_고객정보"
'        Query = Query & " WHERE (본사전송여부 <> 'Y' OR 본사전송여부 IS NULL)"
'        Query = Query & "   AND 가맹점코드 = '" & lblCode.Caption & "'"
'    End If

    If (m_bTotal = True) Then
        Query = "SELECT * FROM TB_고객정보 "
        Query = Query & "   where 가맹점코드 = '" & lblCode.Caption & "'"
    ElseIf (m_bDay = True) Then
        Query = "SELECT * FROM TB_고객정보"
            Query = Query & "   WHERE (등록일자 >= '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "'"
            Query = Query & "   AND  등록일자 <= '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "')"
            Query = Query & "   AND 가맹점코드 = '" & lblCode.Caption & "'"
                    
        Else
            Query = "SELECT * FROM TB_고객정보"
            Query = Query & " WHERE (본사전송여부 <> 'Y' OR 본사전송여부 IS NULL)"
            Query = Query & "   AND 가맹점코드 = '" & lblCode.Caption & "'"
        End If

    Query = Query & " ORDER BY 고객코드 " & m_sOrderBy
    
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOConCleanAid, adOpenForwardOnly, adLockReadOnly
    
    Do Until ADORs.EOF
        nCnt = nCnt + 1
        
        sprGrid.Row = sprGrid.MaxRows
        sprGrid.Col = 2: sprGrid.Text = Format(nCnt, "#,##0") '현재
        DoEvents
        
        '---------------------------------------------------------
        ' 고객정보
        '---------------------------------------------------------
        sValue(0) = ADORs!지사코드 & ""             '1
        sValue(1) = ADORs!가맹점코드 & ""           '2
        sValue(2) = ADORs!고객코드 & ""             '3
        sValue(3) = ADORs!성명 & ""                 '4
        sValue(4) = ADORs!전화번호 & ""             '5
        sValue(5) = ADORs!휴대전화 & ""             '6
        sValue(6) = ADORs!주소 & ""                 '7
        sValue(7) = ADORs!카드번호 & ""             '8
        sValue(8) = ADORs!문자발송여부 & ""         '9
        sValue(9) = ADORs!등록일자 & ""             '10
        sValue(10) = ADORs!메모 & ""                '11
        sValue(11) = ADORs!고객등급코드 & ""        '12
        sValue(12) = ADORs!수정일자 & ""            '13
        sValue(13) = ADORs!접수금액 & ""            '14
        sValue(14) = ADORs!현금입금 & ""            '15
        sValue(15) = ADORs!카드입금 & ""            '16
        sValue(16) = ADORs!사용마일리지 & ""        '17
        sValue(17) = ADORs!쿠폰금액 & ""            '18
        sValue(18) = ADORs!세트할인 & ""            '19
        sValue(19) = ADORs!에누리 & ""              '20
        sValue(20) = ADORs!미수금액 & ""            '21
        sValue(21) = ADORs!이용횟수 & ""            '22
        sValue(22) = ADORs!총접수금액 & ""          '23
        sValue(23) = ADORs!누적마일리지 & ""        '24
        sValue(24) = ADORs!사용가능마일리지 & ""    '25
        sValue(25) = IIf(ADORs!삭제 = False, 0, 1)  ' 26
        sValue(26) = ADORs!본사전송여부 & ""        '27
        sValue(27) = ADORs!최종거래일자 & ""        '28
        sValue(28) = ADORs!초기미수금 & ""          '29
        sValue(29) = ADORs!이전고객 & ""            '30
        
        Call ExecPro("SP_SE_00002_INS", sValue(), Err_Num, Err_Dec)
        
        If LaundryDB = "LAUNDRY1000" And Err_Num = 0 Then
            Query = "UPDATE TB_고객정보 SET 본사전송여부 = 'Y'"
            Query = Query & " WHERE 고객코드 = '" & ADORs!고객코드 & "'"
            ADOConCleanAid.Execute Query
        End If
        
        ADORs.MoveNext
    Loop
    ADORs.Close:    Set ADORs = Nothing
    Send_고객정보 = True
    Exit Function
    
ERR_RTN:
    Send_고객정보 = False
    sprGrid.Row = sprGrid.MaxRows
    sprGrid.Col = 2: sprGrid.Text = Err.Description

End Function

Private Function Send_일일마감(LaundryDB As String) As Boolean
    Dim nCnt    As Long
    ReDim sValue(58)
    
    On Error GoTo ERR_RTN
    Send_일일마감 = False
    nCnt = 0
    
    Query = "SELECT * FROM TB_일일마감"
    Query = Query & " WHERE 가맹점코드 = '" & lblCode.Caption & "'"
    
    '2024-09-02 지사변경 전에 마감을 하지 않은 경우 지사 변경처리를 하기 위해 지사코드를 제외시킴
    'Query = Query & " WHERE 지사코드 = '" & lblOffice.Caption & "'"
    'Query = Query & "   AND 가맹점코드 = '" & lblCode.Caption & "'"
    
    If m_bDay = False Then
        Query = Query & "   AND (본사전송여부 <> 'Y' OR 본사전송여부 IS NULL)"
    Else
        Query = Query & "   AND (마감일자 >= '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "'"
        Query = Query & "   AND  마감일자 <= '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "')"
    End If
    Query = Query & " ORDER BY 마감일자 " & m_sOrderBy
    
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOConCleanAid, adOpenForwardOnly, adLockReadOnly
    
    Do Until ADORs.EOF
        nCnt = nCnt + 1
        
        sprGrid.Row = sprGrid.MaxRows
        sprGrid.Col = 2: sprGrid.Text = Format(nCnt, "#,##0") '현재
        DoEvents
        
        
        sValue(0) = ADORs!지사코드 & ""                         '1
        sValue(1) = ADORs!가맹점코드 & ""                       '2
        sValue(2) = Format(ADORs!마감일자, "YYYY-MM-DD") & ""   '3
        sValue(3) = ADORs!접수금액 & ""                         '4
        sValue(4) = ADORs!접수수량 & ""                         '5
        sValue(5) = ADORs!출고수량 & ""                         '6
        sValue(6) = ADORs!반품수량 & ""                         '7
        sValue(7) = ADORs!재세탁수량 & ""                       '8
        sValue(8) = ADORs!수선금액 & ""                         '9
        sValue(9) = ADORs!수선수량 & ""                         '10
        sValue(10) = ADORs!판매구분 & ""                        '11
        sValue(11) = ADORs!시작택번호 & ""                      '12
        sValue(12) = ADORs!종료택번호 & ""                      '13
        sValue(13) = ADORs!쿠폰금액 & ""                        '14
        sValue(14) = ADORs!쿠폰건수 & ""                        '15
        sValue(15) = ADORs!발생마일리지 & ""                    '16
        sValue(16) = ADORs!사용마일리지 & ""                    '17
        sValue(17) = ADORs!삭제마일리지 & ""                    '18
        sValue(18) = ADORs!현금입금 & ""                        '19
        sValue(19) = ADORs!카드금액 & ""                        '20
        sValue(20) = ADORs!카드건수 & ""                        '21
        sValue(21) = ADORs!반품환불금액 & ""                    '22
        sValue(22) = ADORs!반품환불건수 & ""                    '23
        sValue(23) = ADORs!세탁환불금액 & ""                    '24
        sValue(24) = ADORs!세탁환불건수 & ""                    '25
        sValue(25) = ADORs!삼성카드할인금액 & ""                '26
        sValue(26) = ADORs!삼성카드할인건수 & ""                '27
        sValue(27) = ADORs!삼성카드할인고객수 & ""              '28
        sValue(28) = ADORs!근무자명 & ""                        '29
        sValue(29) = ADORs!지사금액 & ""                        '30
        sValue(30) = ADORs!가맹점금액 & ""                      '31
        sValue(31) = ADORs!운동화금액 & ""                      '32
        sValue(32) = ADORs!운동화건수 & ""                      '33
        sValue(33) = ADORs!운동화비율 & ""                      '34
        sValue(34) = ADORs!카페트금액 & ""                      '35
        sValue(35) = ADORs!카페트건수 & ""                      '36
        sValue(36) = ADORs!명품세탁금액 & ""                    '37
        sValue(37) = ADORs!명품세탁건수 & ""                    '38
        sValue(38) = ADORs!명품세탁비율 & ""                    '39
        sValue(39) = ADORs!명품염색금액 & ""                    '40
        sValue(40) = ADORs!명품염색건수 & ""                    '41
        sValue(41) = ADORs!명품염색비율 & ""                    '42
        sValue(42) = ADORs!로열티정보1 & ""                    '42
        sValue(43) = ADORs!로열티정보2 & ""                    '42
        sValue(44) = ADORs!수수료정보 & ""                    '42
        sValue(45) = ADORs!반품환불지사금액 & ""               '42
        sValue(46) = ADORs!세탁환불지사금액 & ""               '42
        sValue(47) = ADORs!카드취소금액 & ""                   '42
        sValue(48) = ADORs!카드취소건수 & ""                   '42
        sValue(49) = ADORs!로열티금액1 & ""                    '42
        
        sValue(50) = ADORs!로열티금액2 & ""                    '42
        sValue(51) = ADORs!수수료승인금액 & ""                 '42
        sValue(52) = ADORs!수수료취소금액 & ""                 '42
        
        sValue(53) = ADORs!미수카드건수 & ""                    '42
        sValue(54) = ADORs!미수카드금액 & ""                    '42
        sValue(55) = ADORs!미수현금수금금액 & ""                '42
        
        sValue(56) = ADORs!마감여부 & ""                        '43
        sValue(57) = ""                                         '44
        sValue(58) = ADORs!전산사용료 & ""

        
        Call ExecPro("SP_SE_00003_INS_NEW2", sValue(), Err_Num, Err_Dec)
        
        If LaundryDB = "LAUNDRY1000" And Err_Num = 0 Then
            '----------------------------------------------------------
            ' 일일정산 Update
            '----------------------------------------------------------
            Query = "UPDATE TB_일일마감 SET 본사전송여부 = 'Y'"
            Query = Query & " WHERE 마감일자   = '" & Format(ADORs!마감일자, "YYYY-MM-DD") & "'"
            Query = Query & "   AND 가맹점코드 = '" & ADORs!가맹점코드 & "'"
            ADOConCleanAid.Execute Query
        End If
        
        ADORs.MoveNext
    Loop
    ADORs.Close:    Set ADORs = Nothing
    Send_일일마감 = True
    Exit Function
    
ERR_RTN:
    Send_일일마감 = False
    sprGrid.Row = sprGrid.MaxRows
    sprGrid.Col = 2: sprGrid.Text = Err.Description

End Function

        

Private Function Send_사고품(LaundryDB As String) As Boolean
    Dim nCnt    As Long
    ReDim sValue(41)

    On Error GoTo ERR_RTN
    Send_사고품 = False
    nCnt = 0
        
    Query = "SELECT * FROM TB_사고품내역"
    Query = Query & "   WHERE 가맹점코드 = '" & lblCode.Caption & "'"
    If m_bTotal = False Then Query = Query & " AND (본사전송여부 <> 'Y' OR 본사전송여부 IS NULL)"
    Query = Query & " ORDER BY 사고접수일자 " & m_sOrderBy
    
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOConCleanAid, adOpenForwardOnly, adLockReadOnly
    
    Do Until ADORs.EOF
        nCnt = nCnt + 1
        
        sprGrid.Row = sprGrid.MaxRows
        sprGrid.Col = 2: sprGrid.Text = Format(nCnt, "#,##0") '현재
        DoEvents
        
        sValue(0) = ADORs!지사코드 & ""         ' 1
        sValue(1) = ADORs!가맹점코드 & ""       ' 2
        sValue(2) = ADORs!일련번호 & ""         ' 3
        sValue(3) = ADORs!사고접수일자 & ""     ' 4
        sValue(4) = ADORs!담당자명 & ""         ' 5
        sValue(5) = ADORs!고객코드 & ""         ' 6
        sValue(6) = ADORs!성명 & ""             ' 7
        sValue(7) = ADORs!전화번호 & ""         ' 8
        sValue(8) = ADORs!휴대전화 & ""         ' 9
        sValue(9) = ADORs!주소 & ""             '10
        sValue(10) = ADORs!접수일자 & ""        '11
        sValue(11) = ADORs!택번호 & ""          '12
        sValue(12) = ADORs!출고일자 & ""        '13
        sValue(13) = ADORs!인도일자 & ""        '14
        sValue(14) = ADORs!의류명 & ""          '15
        sValue(15) = ADORs!상표 & ""            '16
        sValue(16) = ADORs!색상 & ""            '17
        sValue(17) = ADORs!구입일자 & ""        '18
        sValue(18) = ADORs!구입처 & ""          '19
        sValue(19) = ADORs!구입형태 & ""        '20
        sValue(20) = ADORs!구입가격 & ""        '21
        sValue(21) = ADORs!품목 & ""            '22
        sValue(22) = ADORs!용도 & ""            '23
        sValue(23) = ADORs!소재 & ""            '24
        sValue(24) = ADORs!내용연수 & ""        '25
        sValue(25) = ADORs!경과일수 & ""        '26
        sValue(26) = ADORs!환산일수 & ""        '27
        sValue(27) = ADORs!배상비율 & ""        '28
        sValue(28) = ADORs!배상금액 & ""        '29
        sValue(29) = ADORs!크레임구분 & ""      '30
        sValue(30) = ADORs!보상구분 & ""        '31
        sValue(31) = ADORs!처리구분 & ""        '32
        sValue(32) = ADORs!보상금액 & ""        '33
        sValue(33) = ADORs!처리일자 & ""        '34
        sValue(34) = ADORs!비고 & ""            '35
        sValue(35) = ADORs!가맹점의견 & ""      '36
        sValue(36) = ADORs!지사의견 & ""        '37
        sValue(37) = ADORs!본사의견 & ""        '38
        sValue(38) = ADORs!지사승인 & ""        '39
        sValue(39) = ADORs!지사승인일시 & ""    '40
        sValue(40) = ADORs!본사승인 & ""        '41
        sValue(41) = ADORs!본사승인일시 & ""    '42
      
        Call ExecPro("SP_SE_00004_INS", sValue(), Err_Num, Err_Dec)
        
        
        If LaundryDB = "LAUNDRY1000" And Err_Num = 0 Then
            '----------------------------------------------------------
            ' TB_사고품내역 Update
            '----------------------------------------------------------
            Query = "UPDATE TB_사고품내역 SET 본사전송여부 = 'Y'"
            Query = Query & ", 본사전송일자 = '" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "' "
            Query = Query & " WHERE 가맹점코드 = '" & lblCode.Caption & "'"
            Query = Query & "   AND 일련번호   =  " & ADORs!일련번호
            ADOConCleanAid.Execute Query
            '----------------------------------------------------------
        End If
        
        ADORs.MoveNext
    Loop
    ADORs.Close:    Set ADORs = Nothing
    Send_사고품 = True
    Exit Function
    
ERR_RTN:
    Send_사고품 = False
    sprGrid.Row = sprGrid.MaxRows
    sprGrid.Col = 2: sprGrid.Text = Err.Description

End Function

'===========================================================
' * 본사에서 환불확정처리건을 수신
'===========================================================
Private Function Recv_환불확정처리() As Boolean
    ReDim sValue(0)
        
    On Error GoTo ERR_RTN
    Recv_환불확정처리 = False
    
    sValue(0) = lblCode.Caption
    Set SubRs = New ADODB.Recordset
    Set SubRs = ExecPro("SP_SE_00005_SEL", sValue(), Err_Num, Err_Dec)
    
    If Err_Num = 0 Then
        Do Until SubRs.EOF
            Query = "UPDATE TB_입출고 SET 환불확정일자 = '" & SubRs!환불확정일자 & "'"
            Query = Query & " WHERE 택번호   = '" & SubRs!택번호 & "'"
            Query = Query & "   AND 가맹점코드 = '" & lblCode.Caption & "'"
            Query = Query & "   AND 접수일자 = '" & Format(SubRs!접수일자, "YYYY-MM-DD") & "'"
            ADOConCleanAid.Execute Query
            
            SubRs.MoveNext
        Loop
    Else
        sprGrid.Row = sprGrid.MaxRows:  sprGrid.Col = 2: sprGrid.Text = Err.Description
    End If
    
    SubRs.Close:    Set SubRs = Nothing
    Recv_환불확정처리 = True
    Exit Function
    
ERR_RTN:
    Recv_환불확정처리 = False
    sprGrid.Row = sprGrid.MaxRows
    sprGrid.Col = 2: sprGrid.Text = Err.Description

End Function

Private Function Send_입출고() As Boolean
    Dim nCnt    As Long
    
    On Error GoTo ERR_RTN
    Send_입출고 = False
    nCnt = 0
    
    Query = "SELECT * FROM TB_입출고"
    Query = Query & " WHERE 지사코드 = '" & lblOffice.Caption & "'"
    Query = Query & "   AND 가맹점코드 = '" & lblCode.Caption & "'"
        
    If m_bDay = False Then
        Query = Query & "   AND (본사전송여부 <> 'Y' OR 본사전송여부 IS NULL)"
    Else
        Query = Query & "   AND (접수일자 >= '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "'"
        Query = Query & "   AND  접수일자 <= '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "')"
    End If
    Query = Query & " ORDER BY 택번호 " & m_sOrderBy
    
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOConCleanAid, adOpenForwardOnly, adLockReadOnly
            
    Do Until ADORs.EOF
        nCnt = nCnt + 1
        
        sprGrid.Row = sprGrid.MaxRows
        sprGrid.Col = 2: sprGrid.Text = Format(nCnt, "#,##0") '현재
        DoEvents
        
        ReDim sValue(50)
        
        sValue(0) = ADORs!지사코드 & ""         '1
        sValue(1) = ADORs!가맹점코드 & ""       '2
        sValue(2) = ADORs!접수일자 & ""         '3
        sValue(3) = ADORs!택번호 & ""           '4
        sValue(4) = ADORs!접수번호 & ""         '5
        sValue(5) = ADORs!고객코드 & ""         '6
        sValue(6) = ADORs!의류코드 & ""         '7
        sValue(7) = ADORs!의류명 & ""           '8
        sValue(8) = ADORs!색상 & ""             '9
        sValue(9) = ADORs!무늬 & ""             '10
        sValue(10) = ADORs!내용 & ""            '11
        sValue(11) = ADORs!금액 & ""            '12
        sValue(12) = ADORs!상표 & ""            '13
        sValue(13) = ADORs!결제여부 & ""        '14
        sValue(14) = ADORs!예정일자 & ""        '15
        sValue(15) = ADORs!출고일자 & ""        '16
        sValue(16) = ADORs!판매취소 & ""        '17
        sValue(17) = ADORs!판매취소일자 & ""    '18
        sValue(18) = ADORs!반품환불일자 & ""    '19
        sValue(19) = ADORs!세탁환불일자 & ""    '20
        sValue(20) = ADORs!환불사유 & ""        '21
        sValue(21) = ADORs!수선금액 & ""        '22
        sValue(22) = ADORs!세탁마진 & ""        '23
        sValue(23) = ADORs!외주마진 & ""        '24
        sValue(24) = ADORs!수선마진 & ""        '25
        sValue(25) = ADORs!세트Key & ""         '26
        sValue(26) = ADORs!세트구분 & ""        '27
        sValue(27) = ADORs!세트금액1 & ""       '28
        sValue(28) = ADORs!세트금액2 & ""       '29
        sValue(29) = ADORs!정상금액 & ""        '30
        sValue(30) = ADORs!마일리지 & ""        '31
        sValue(31) = ADORs!가맹점출고일자 & ""  '32
        sValue(32) = ADORs!가맹점입고일자 & ""  '33
        sValue(33) = ADORs!가맹점입고구분 & ""  '34
        sValue(34) = ADORs!부모택번호 & ""      '35
        sValue(35) = ADORs!근무자명 & ""        '36
        sValue(36) = ADORs!미입고사유 & ""      '37
        sValue(37) = ADORs!오점내용 & ""        '38
        sValue(38) = ADORs!접수시간 & ""        '39
        sValue(39) = ADORs!출고시간 & ""        '40
        sValue(40) = ADORs!의류금액 & ""        '41
        sValue(41) = ADORs!지사입고일자 & ""    '42
        sValue(42) = ADORs!지사출고일자 & ""    '43
        sValue(43) = ADORs!지사출고물품 & ""    '44
        sValue(44) = ADORs!지사출고상태 & ""    '45
        sValue(45) = ""                         '46 ADORs!본사전송여부
        
        'SubRs!지사출고예정 = ADORs!지사출고예정 & ""     '47
        'SubRs!지사출고번호 = ADORs!지사출고번호 & ""     '48
        
        Call ExecPro("SP_SE_00006_INS", sValue(), Err_Num, Err_Dec)
        
        If Err_Num = 0 Then
    '        '----------------------------------------------------------
    '        ' 오점이미지가 있을 경우 저장한다.
    '        '----------------------------------------------------------
    '        If (Trim(ADORs!오점이미지) = "") Or (IsNull(ADORs!오점이미지)) Then
    '        Else
    '            Query = "SELECT * FROM TB_입출고"
    '            Query = Query & " WHERE 가맹점코드 = '" & lblCode.Caption & "'"
    '            Query = Query & "   AND 접수일자   = '" & ADORs!접수일자 & "'"
    '            Query = Query & "   AND 택번호     = '" & ADORs!택번호 & "'"
    '            Query = Query & "   AND 접수번호   =  " & ADORs!접수번호
    '            Set SubRs = New ADODB.Recordset
    '            'SubRs.Open Query, ADONewServer, adOpenKeyset, adLockOptimistic
    '            'SubRs.Open Query, ADONewServer, adOpenDynamic, adLockOptimistic
    '            SubRs.CursorType = adOpenKeyset
    '            SubRs.LockType = adLockOptimistic
    '            SubRs.Open Query, ADONewServer
    '
    '
    '            If Not SubRs.BOF Then
    '                SubRs.Fields("오점이미지").AppendChunk ADORs.Fields("오점이미지").GetChunk(1024)
    ''                SubRs.Fields("오점이미지").AppendChunk ADORs.Fields("오점이미지").GetChunk(Len(ADORs.Fields("오점이미지").Value))
    '                SubRs.Update
    '            End If
    '
    '            SubRs.Close:    Set SubRs = Nothing
    '        End If
        
            '----------------------------------------------------------
            ' ParentList_TB Update
            '----------------------------------------------------------
            If (ADORs!부모택번호 = "") Or IsNull(ADORs!부모택번호) Then
                '
            Else
                If Trim(ADORs!부모택번호) <> "" Then
                    
                    ReDim sValue(3)
                    
                    sValue(0) = ADORs!가맹점코드 & ""
                    sValue(1) = ADORs!접수일자 & ""
                    sValue(2) = ADORs!부모택번호 & ""
                    sValue(3) = ADORs!접수번호 & ""
                
                    Call ExecPro("SP_SE_00007_INS", sValue(), Err_Num, Err_Dec)
                End If
            End If
                        
            If Err_Num = 0 Then
                '----------------------------------------------------------
                ' TB_접수 Update
                '----------------------------------------------------------
                Query = "UPDATE TB_입출고 SET 본사전송여부 = 'Y'"
                Query = Query & " WHERE 가맹점코드 = '" & ADORs!가맹점코드 & "'"
                Query = Query & "   AND 접수일자   = '" & ADORs!접수일자 & "'"
                Query = Query & "   AND 택번호   = '" & ADORs!택번호 & "'"
                Query = Query & "   AND 접수번호 =  " & ADORs!접수번호
                ADOConCleanAid.Execute Query
                '----------------------------------------------------------
            Else
                Call ERR_SAVE("Send_입출고 SP_SE_00007_INS " & Err_Dec)
            
            End If
        
        Else
        
            sprGrid.Row = sprGrid.MaxRows
            sprGrid.Col = 2: sprGrid.Text = Err_Dec
        
            Call ERR_SAVE("Send_입출고 SP_SE_00006_INS " & Err_Dec)
        End If
        
        
        ADORs.MoveNext
    Loop
    ADORs.Close:    Set ADORs = Nothing
    Send_입출고 = True
    Exit Function

ERR_RTN:
    ADORs.Close:    Set ADORs = Nothing '2024.11.08 추가
    
    Send_입출고 = False
    sprGrid.Row = sprGrid.MaxRows
    sprGrid.Col = 2: sprGrid.Text = Err.Description

    Call ERR_SAVE("Send_입출고 ERR_RTN " & Err.Description)
    
    'Resume Next    '2024.11.08 삭제
    On Error GoTo 0 '2024.11.08 추가
            
End Function

Private Function Send_매출() As Boolean
    Dim nCnt    As Long
    ReDim sValue(24)

    On Error GoTo ERR_RTN
    Send_매출 = False
    nCnt = 0
    
    Query = "SELECT * FROM TB_매출"
    Query = Query & " WHERE 지사코드 = '" & lblOffice.Caption & "'"
    Query = Query & "   AND 가맹점코드 = '" & lblCode.Caption & "'"
        
    If m_bDay = False Then
        Query = Query & "   AND (본사전송여부 <> 'Y' OR 본사전송여부 IS NULL)"
    Else
        Query = Query & "   AND (매출일자 >= '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "'"
        Query = Query & "   AND  매출일자 <= '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "')"
    End If
    Query = Query & " ORDER BY 매출일자 " & m_sOrderBy
    
    
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOConCleanAid, adOpenForwardOnly, adLockReadOnly
            
    Do Until ADORs.EOF
        nCnt = nCnt + 1
        
        sprGrid.Row = sprGrid.MaxRows
        sprGrid.Col = 2: sprGrid.Text = Format(nCnt, "#,##0") '현재
        DoEvents
        
        sValue(0) = ADORs!지사코드 & ""             '1
        sValue(1) = ADORs!가맹점코드 & ""           '2
        sValue(2) = ADORs!고객코드 & ""             '3
        sValue(3) = ADORs!접수번호 & ""             '4
        sValue(4) = ADORs!일련번호 & ""             '5
        sValue(5) = ADORs!매출일자 & ""             '6
        sValue(6) = ADORs!매출시간 & ""             '7
        sValue(7) = ADORs!적요 & ""                 '8
        sValue(8) = ADORs!접수금액 & ""             '9
        sValue(9) = ADORs!입금합계 & ""             '10
        sValue(10) = ADORs!현금입금 & ""            '11
        sValue(11) = ADORs!카드입금 & ""            '12
        sValue(12) = ADORs!쿠폰입금 & ""            '13
        sValue(13) = ADORs!쿠폰번호 & ""            '14
        sValue(14) = ADORs!사용마일리지 & ""        '15
        sValue(15) = ADORs!세트할인 & ""            '16
        sValue(16) = ADORs!에누리 & ""              '17
        sValue(17) = ADORs!접수수량 & ""            '18
        sValue(18) = ADORs!반품수량 & ""            '19
        sValue(19) = ADORs!발생마일리지 & ""        '20
        sValue(20) = ADORs!누적마일리지 & ""        '21
        sValue(21) = ADORs!사용가능마일리지 & ""    '22
        sValue(22) = ADORs!삭제마일리지 & ""        '23
        sValue(23) = ADORs!이전미수금 & ""          '24
        sValue(24) = ""                             '25 ADORs!본사전송여부
        
        Call ExecPro("SP_SE_00008_INS", sValue(), Err_Num, Err_Dec)
        
        If Err_Num = 0 Then
            '----------------------------------------------------------
            ' TB_접수 Update
            '----------------------------------------------------------
            Query = "UPDATE TB_매출 SET 본사전송여부 = 'Y'"
            Query = Query & " WHERE 고객코드   = '" & ADORs!고객코드 & "'"
            Query = Query & "   AND 접수번호   =  " & ADORs!접수번호
            Query = Query & "   AND 일련번호   =  " & ADORs!일련번호
            Query = Query & "   AND 가맹점코드 = '" & ADORs!가맹점코드 & "'"
            ADOConCleanAid.Execute Query
            '----------------------------------------------------------
        End If
        
        ADORs.MoveNext
    Loop
    ADORs.Close:    Set ADORs = Nothing
    Send_매출 = False
    Exit Function
    
ERR_RTN:
    Send_매출 = True
    sprGrid.Row = sprGrid.MaxRows
    sprGrid.Col = 2: sprGrid.Text = Err.Description

End Function

Private Function Send_현금영수증(LaundryDB As String) As Boolean
    Dim nCnt    As Long
    ReDim sValue(19)
    
    On Error GoTo ERR_RTN
    Send_현금영수증 = False
    nCnt = 0
    
    Query = "SELECT * FROM TB_현금영수증"
    Query = Query & " WHERE 지사코드 = '" & lblOffice.Caption & "'"
    Query = Query & "   AND 가맹점코드 = '" & lblCode.Caption & "'"
    If m_bTotal = False Then
        Query = Query & " AND (본사전송여부 <> 'Y' OR 본사전송여부 IS NULL)"
    End If
    Query = Query & " ORDER BY 승인일자 " & m_sOrderBy
    
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOConCleanAid, adOpenForwardOnly, adLockReadOnly
    
    Do Until ADORs.EOF
        nCnt = nCnt + 1
        
        sprGrid.Row = sprGrid.MaxRows
        sprGrid.Col = 2: sprGrid.Text = Format(nCnt, "#,##0") '현재
        DoEvents
        
        sValue(0) = ADORs!지사코드 & ""
        sValue(1) = ADORs!가맹점코드 & ""
        sValue(2) = ADORs!고객코드 & ""
        sValue(3) = ADORs!접수번호 & ""
        sValue(4) = ADORs!승인번호 & ""
        sValue(5) = ADORs!승인일자 & ""
        sValue(6) = ADORs!승인시간 & ""
        sValue(7) = ADORs!거래유형 & ""
        sValue(8) = ADORs!입력방법 & ""
        sValue(9) = ADORs!사용자정보 & ""
        sValue(10) = ADORs!총금액 & ""
        sValue(11) = ADORs!메시지1 & ""
        sValue(12) = ADORs!메시지2 & ""
        sValue(13) = ADORs!소득구분 & ""
        sValue(14) = ADORs!국세청1 & ""
        sValue(15) = ADORs!국세청2 & ""
        sValue(16) = ADORs!단말기번호 & ""
        sValue(17) = ADORs!거래구분 & ""
        sValue(18) = ADORs!상태 & ""
        sValue(19) = "" ' ADORs!본사전송여부
        
        Call ExecPro("SP_SE_00009_INS", sValue(), Err_Num, Err_Dec)
        
        If Err_Num = 0 Then
            '----------------------------------------------------------
            ' TB_현금영수증 Update
            '----------------------------------------------------------
            Query = "UPDATE TB_현금영수증 SET 본사전송여부 = 'Y'"
            Query = Query & " WHERE 가맹점코드 = '" & ADORs!가맹점코드 & "'"
            Query = Query & "   AND 승인번호 = '" & ADORs!승인번호 & "'"
            Query = Query & "   AND 승인일자 = '" & ADORs!승인일자 & "'"
            ADOConCleanAid.Execute Query
            '----------------------------------------------------------
        End If
        
        ADORs.MoveNext
    Loop
    ADORs.Close:    Set ADORs = Nothing
    Send_현금영수증 = True
    Exit Function

ERR_RTN:
    Send_현금영수증 = False
    sprGrid.Row = sprGrid.MaxRows
    sprGrid.Col = 2: sprGrid.Text = Err.Description
End Function

Private Function Send_신용카드승인(LaundryDB As String) As Boolean
    Dim nCnt    As Long
    ReDim sValue(21)
    
    On Error GoTo ERR_RTN
    Send_신용카드승인 = False
    nCnt = 0
    
    Query = "SELECT * FROM TB_신용카드승인"
    Query = Query & " WHERE 지사코드 = '" & lblOffice.Caption & "'"
    Query = Query & "   AND 가맹점코드 = '" & lblCode.Caption & "'"
    
    If m_bTotal = False Then Query = Query & " AND (본사전송여부 <> 'Y' OR 본사전송여부 IS NULL)"
    Query = Query & " ORDER BY 승인일자 " & m_sOrderBy
    
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOConCleanAid, adOpenForwardOnly, adLockReadOnly
    
    Do Until ADORs.EOF
        nCnt = nCnt + 1
        
        sprGrid.Row = sprGrid.MaxRows
        sprGrid.Col = 2: sprGrid.Text = Format(nCnt, "#,##0") '현재
        DoEvents
    
        
        sValue(0) = ADORs!지사코드 & ""
        sValue(1) = ADORs!가맹점코드 & ""
        sValue(2) = ADORs!고객코드 & ""
        sValue(3) = ADORs!접수번호 & ""
        sValue(4) = ADORs!승인번호 & ""
        sValue(5) = ADORs!승인일자 & ""
        sValue(6) = ADORs!승인시간 & ""
        sValue(7) = ADORs!할부기간 & ""
        sValue(8) = ADORs!결제금액 & ""
        sValue(9) = ADORs!발급사코드 & ""
        sValue(10) = ADORs!카드종류명 & ""
        sValue(11) = ADORs!매입사코드 & ""
        sValue(12) = ADORs!매입사명 & ""
        sValue(13) = ADORs!카드번호 & ""
        sValue(14) = ADORs!메시지1 & ""
        sValue(15) = ADORs!메시지2 & ""
        sValue(16) = ADORs!가맹점번호 & ""
        sValue(17) = ADORs!단말기번호 & ""
        sValue(18) = ADORs!거래구분 & ""
        sValue(19) = ADORs!상태 & ""
        sValue(20) = ADORs!취소일자 & ""
        sValue(21) = ADORs!기타메모 & ""
        
        Call ExecPro("SP_SE_00010_INS_NEW", sValue(), Err_Num, Err_Dec)
        
        If Err_Num = 0 Then
            '----------------------------------------------------------
            ' TB_신용카드승인 Update
            '----------------------------------------------------------
            Query = "UPDATE TB_신용카드승인 SET 본사전송여부 = 'Y'"
            Query = Query & " WHERE 가맹점코드 = '" & ADORs!가맹점코드 & "'"
            Query = Query & "   AND 승인번호 = '" & ADORs!승인번호 & "'"
            Query = Query & "   AND 승인일자 = '" & ADORs!승인일자 & "'"
            ADOConCleanAid.Execute Query
        End If
        
        ADORs.MoveNext
    Loop
    ADORs.Close:    Set ADORs = Nothing
    Send_신용카드승인 = True
    Exit Function
    
ERR_RTN:
    Send_신용카드승인 = False
    sprGrid.Row = sprGrid.MaxRows
    sprGrid.Col = 2: sprGrid.Text = Err.Description

End Function

Private Function Send_이용실적(LaundryDB As String) As Boolean
    Dim nCnt    As Long
    ReDim sValue(5)
    
    On Error GoTo ERR_RTN
    Send_이용실적 = False
    nCnt = 0
        
    Query = "SELECT * FROM TB_이용실적"
    Query = Query & "   WHERE 가맹점코드 = '" & lblCode.Caption & "'"

    If m_bTotal = False Then
        Query = Query & " AND (본사전송여부 <> 'Y' OR 본사전송여부 IS NULL)"
    End If
    Query = Query & " ORDER BY 고객코드 " & m_sOrderBy
    
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOConCleanAid, adOpenForwardOnly, adLockReadOnly
    
    Do Until ADORs.EOF
        nCnt = nCnt + 1
        
        sprGrid.Row = sprGrid.MaxRows
        sprGrid.Col = 2: sprGrid.Text = Format(nCnt, "#,##0") '현재
        DoEvents
    
        
        sValue(0) = ADORs!가맹점코드 & ""
        sValue(1) = ADORs!고객코드 & ""
        sValue(2) = ADORs!연도 & ""
        sValue(3) = ADORs!이용횟수 & ""
        sValue(4) = ADORs!이용금액 & ""
        sValue(5) = ""  ' ADORs!본사전송여부
        
        Call ExecPro("SP_SE_00011_INS", sValue(), Err_Num, Err_Dec)
        
        If Err_Num = 0 Then
            '----------------------------------------------------------
            ' TB_이용실적 Update
            '----------------------------------------------------------
            Query = "UPDATE TB_이용실적 SET 본사전송여부 = 'Y'"
            Query = Query & " WHERE 가맹점코드 = '" & ADORs!가맹점코드 & "'"
            Query = Query & "   AND 고객코드   = '" & ADORs!고객코드 & "'"
            Query = Query & "   AND 연도       = '" & ADORs!연도 & "'"
            ADOConCleanAid.Execute Query
        End If
        
        ADORs.MoveNext
    Loop
    ADORs.Close:    Set ADORs = Nothing
    Send_이용실적 = True
    Exit Function
    
ERR_RTN:
    Send_이용실적 = False
    sprGrid.Row = sprGrid.MaxRows
    sprGrid.Col = 2: sprGrid.Text = Err.Description
End Function

Private Function Send_가맹점입금(LaundryDB As String) As Boolean
    Dim nCnt    As Long
    ReDim sValue(10)

    On Error GoTo ERR_RTN
    Send_가맹점입금 = False
    nCnt = 0
        
    Query = "SELECT * FROM TB_가맹점입금"
    Query = Query & "  WHERE 지사코드 = '" & lblOffice.Caption & "'"
    Query = Query & "    AND 가맹점코드 = '" & lblCode.Caption & "'"

    If m_bTotal = False Then
        Query = Query & " AND (본사전송여부 <> 'Y' OR 본사전송여부 IS NULL)"
    End If
    Query = Query & " ORDER BY 입금일자 " & m_sOrderBy
    
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOConCleanAid, adOpenForwardOnly, adLockReadOnly
    
    Do Until ADORs.EOF
        nCnt = nCnt + 1
        
        sprGrid.Row = sprGrid.MaxRows
        sprGrid.Col = 2: sprGrid.Text = Format(nCnt, "#,##0") '현재
        DoEvents
    
        sValue(0) = ADORs!지사코드 & ""
        sValue(1) = ADORs!가맹점코드 & ""
        sValue(2) = ADORs!입금일자 & ""
        sValue(3) = ADORs!배송기사코드 & ""
        sValue(4) = ADORs!배송기사명 & ""
        sValue(5) = ADORs!입금액 & ""
        sValue(6) = ADORs!비고 & ""
        sValue(7) = ADORs!경리담당자 & ""
        sValue(8) = ADORs!입금확정 & ""
        sValue(9) = ADORs!확정일자 & ""
        sValue(10) = ""                     '본사전송여부
        
        Call ExecPro("SP_SE_00012_INS", sValue(), Err_Num, Err_Dec)
        
        If Err_Num = 0 Then
            '----------------------------------------------------------
            ' TB_가맹점입금 Update
            '----------------------------------------------------------
            Query = "UPDATE TB_가맹점입금 SET 본사전송여부 = 'Y'"
            Query = Query & " WHERE 가맹점코드 = '" & ADORs!가맹점코드 & "'"
            Query = Query & "   AND 고객코드   = '" & ADORs!고객코드 & "'"
            Query = Query & "   AND 연도       = '" & ADORs!연도 & "'"
            ADOConCleanAid.Execute Query
            '----------------------------------------------------------
        End If
        
        ADORs.MoveNext
    Loop
    ADORs.Close:    Set ADORs = Nothing
    Send_가맹점입금 = True
    Exit Function
    
ERR_RTN:
    Send_가맹점입금 = False
    sprGrid.Row = sprGrid.MaxRows
    sprGrid.Col = 2: sprGrid.Text = Err.Description
    
End Function

Private Function Send_부자재주문(LaundryDB As String) As Boolean
    Dim nCnt    As Long
    ReDim sValue(16)

    On Error GoTo ERR_RTN
    Send_부자재주문 = False
    nCnt = 0
        
    Query = "SELECT * FROM TB_부자재주문"
    Query = Query & " WHERE 지사코드 = '" & lblOffice.Caption & "'"
    Query = Query & "   AND 가맹점코드 = '" & lblCode.Caption & "'"

    If m_bTotal = False Then
        Query = Query & " AND (본사전송여부 <> 'Y' OR 본사전송여부 IS NULL)"
    End If
    
    Query = Query & " ORDER BY 주문일자 " & m_sOrderBy
    
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOConCleanAid, adOpenForwardOnly, adLockReadOnly
    
    Do Until ADORs.EOF
        nCnt = nCnt + 1
        
        sprGrid.Row = sprGrid.MaxRows
        sprGrid.Col = 2: sprGrid.Text = Format(nCnt, "#,##0") '현재
        DoEvents
        
        sValue(0) = ADORs!지사코드 & ""
        sValue(1) = ADORs!가맹점코드 & ""
        sValue(2) = ADORs!주문코드 & ""
        sValue(3) = ADORs!주문일자 & ""
        sValue(4) = ADORs!부자재코드 & ""
        sValue(5) = ADORs!부자재명 & ""
        sValue(6) = ADORs!규격 & ""
        sValue(7) = ADORs!수량 & ""
        sValue(8) = ADORs!단가 & ""
        sValue(9) = ADORs!공급가액 & ""
        sValue(10) = ADORs!세액 & ""
        sValue(11) = ADORs!합계금액 & ""
        sValue(12) = ADORs!비고 & ""
        sValue(13) = ADORs!출고일자 & ""
        sValue(14) = ADORs!입고확정 & ""
        sValue(15) = ADORs!확정일자 & ""
        sValue(16) = ""  ' ADORs!본사전송여부
      
        Call ExecPro("SP_SE_00013_INS", sValue(), Err_Num, Err_Dec)
        
        If Err_Num = 0 Then
            '----------------------------------------------------------
            ' TB_부자재주문 Update
            '----------------------------------------------------------
            Query = "UPDATE TB_부자재주문 SET 본사전송여부 = 'Y'"
            Query = Query & " WHERE 가맹점코드 = '" & lblCode.Caption & "'"
            Query = Query & "   AND 주문코드   = '" & ADORs!주문코드 & "'"
            ADOConCleanAid.Execute Query
            '----------------------------------------------------------
        End If
        
        ADORs.MoveNext
    Loop
    ADORs.Close:    Set ADORs = Nothing
    Send_부자재주문 = True
    Exit Function
    
ERR_RTN:
    Send_부자재주문 = False
    sprGrid.Row = sprGrid.MaxRows
    sprGrid.Col = 2: sprGrid.Text = Err.Description
    
End Function

Private Function Send_쿠폰자료(LaundryDB As String) As Boolean
    Dim nCnt    As Long
    ReDim sValue(11)
        
    On Error GoTo ERR_RTN
    Send_쿠폰자료 = False
    nCnt = 0
    
    Query = "SELECT * FROM TB_쿠폰자료"
    Query = Query & " WHERE 지사코드 = '" & lblOffice.Caption & "'"
    Query = Query & "   AND 가맹점코드 = '" & lblCode.Caption & "'"

    If m_bTotal = False Then
        Query = Query & " AND (본사전송여부 <> 'Y' OR 본사전송여부 IS NULL)"
    End If
    Query = Query & " ORDER BY 접수일자 " & m_sOrderBy
    
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOConCleanAid, adOpenForwardOnly, adLockReadOnly
    
    Do Until ADORs.EOF
        nCnt = nCnt + 1
        
        sprGrid.Row = sprGrid.MaxRows
        sprGrid.Col = 2: sprGrid.Text = Format(nCnt, "#,##0") '현재
        DoEvents
        
        sValue(0) = ADORs!가맹점코드 & ""
        sValue(1) = ADORs!접수일자 & ""
        sValue(2) = ADORs!쿠폰번호 & ""
        sValue(3) = ADORs!쿠폰단가 & ""
        sValue(4) = ADORs!쿠폰금액 & ""
        sValue(5) = ADORs!고객코드 & ""
        sValue(6) = ADORs!고객이름 & ""
        sValue(7) = ADORs!접수금액 & ""
        sValue(8) = ADORs!택번호 & ""
        sValue(9) = ADORs!지사코드 & ""
        sValue(10) = ""
        
        Call ExecPro("SP_SE_00014_INS", sValue(), Err_Num, Err_Dec)
        
        If Err_Num = 0 Then
            '----------------------------------------------------------
            ' TB_쿠폰자료 Update
            '----------------------------------------------------------
            Query = "UPDATE TB_쿠폰자료 SET 본사전송여부 = 'Y'"
            Query = Query & " WHERE 가맹점코드 = '" & ADORs!가맹점코드 & "'"
            Query = Query & "   AND 접수일자   = '" & ADORs!접수일자 & "'"
            Query = Query & "   AND 쿠폰번호   = '" & ADORs!쿠폰번호 & "'"
            ADOConCleanAid.Execute Query
            '----------------------------------------------------------
        End If
        
        ADORs.MoveNext
    Loop
    ADORs.Close:    Set ADORs = Nothing
    Send_쿠폰자료 = True
    Exit Function
    
ERR_RTN:
    Send_쿠폰자료 = False
    sprGrid.Row = sprGrid.MaxRows
    sprGrid.Col = 2: sprGrid.Text = Err.Description
    
End Function


Private Function Send_미수금수정(LaundryDB As String) As Boolean
    Dim nCnt    As Long
    ReDim sValue(6)
        
    On Error GoTo ERR_RTN
    Send_미수금수정 = False
    nCnt = 0
    
    Query = "SELECT * FROM TB_미수금수정"
    Query = Query & " WHERE 지사코드 = '" & lblOffice.Caption & "'"
    Query = Query & "   AND 가맹점코드 = '" & lblCode.Caption & "'"
    
    If m_bTotal = False Then
        Query = Query & " AND (본사전송여부 <> 'Y' OR 본사전송여부 IS NULL)"
    End If
    Query = Query & " ORDER BY 수정일자 " & m_sOrderBy
    
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOConCleanAid, adOpenForwardOnly, adLockReadOnly
    
    Do Until ADORs.EOF
        nCnt = nCnt + 1
        
        sprGrid.Row = sprGrid.MaxRows
        sprGrid.Col = 2: sprGrid.Text = Format(nCnt, "#,##0") '현재
        DoEvents
        
        sValue(0) = ADORs!지사코드 & ""
        sValue(1) = ADORs!가맹점코드 & ""
        sValue(2) = ADORs!고객코드 & ""
        sValue(3) = ADORs!수정일자 & ""
        sValue(4) = ADORs!수정미수금 & ""
        sValue(5) = ADORs!이전미수금 & ""
        sValue(6) = ADORs!내용 & ""
        
        
        Call ExecPro("SP_SE_00015_INS", sValue(), Err_Num, Err_Dec)
        
        If Err_Num = 0 Then
            '----------------------------------------------------------
            ' TB_쿠폰자료 Update
            '----------------------------------------------------------
            Query = "UPDATE TB_미수금수정 SET 본사전송여부 = 'Y'"
            Query = Query & " WHERE 가맹점코드 = '" & ADORs!가맹점코드 & "'"
            Query = Query & "   AND 수정일자   = '" & ADORs!수정일자 & "'"
            Query = Query & "   AND 고객코드   = '" & ADORs!고객코드 & "'"
            ADOConCleanAid.Execute Query
            '----------------------------------------------------------
        End If
        
        ADORs.MoveNext
    Loop
    ADORs.Close:    Set ADORs = Nothing
    
    Send_미수금수정 = True
    Exit Function
    
ERR_RTN:
    Send_미수금수정 = False
    sprGrid.Row = sprGrid.MaxRows
    sprGrid.Col = 2: sprGrid.Text = Err.Description
    
End Function

