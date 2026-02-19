VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.MDIForm Main_Form 
   Appearance      =   0  '평면
   BackColor       =   &H8000000F&
   ClientHeight    =   9240
   ClientLeft      =   1620
   ClientTop       =   4110
   ClientWidth     =   15585
   Icon            =   "Main_Form.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   NegotiateToolbars=   0   'False
   ScrollBars      =   0   'False
   WindowState     =   2  '최대화
   Begin ComctlLib.ProgressBar ProgressBar1 
      Align           =   1  '위 맞춤
      Height          =   405
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   15585
      _ExtentX        =   27490
      _ExtentY        =   714
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   0
      Top             =   1800
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   8805
      Width           =   15585
      _ExtentX        =   27490
      _ExtentY        =   767
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   6175
            MinWidth        =   6175
            Text            =   " 크 린 에 이 드 Ver 1.00"
            TextSave        =   " 크 린 에 이 드 Ver 1.00"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Object.Width           =   3528
            MinWidth        =   3528
            TextSave        =   "2010-04-11"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Object.Width           =   2999
            MinWidth        =   2999
            TextSave        =   "오전 9:32"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   8819
            MinWidth        =   8819
            Picture         =   "Main_Form.frx":030A
            Text            =   "000-0000"
            TextSave        =   "000-0000"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  '위 맞춤
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   405
      Width           =   15585
      _ExtentX        =   27490
      _ExtentY        =   1931
      _Version        =   262144
      BackColor       =   16777215
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureBackground=   "Main_Form.frx":0624
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      Begin XtremeSuiteControls.PushButton Command1 
         Height          =   930
         Index           =   1
         Left            =   75
         TabIndex        =   5
         Top             =   75
         Width           =   1650
         _Version        =   851970
         _ExtentX        =   2910
         _ExtentY        =   1640
         _StockProps     =   79
         Caption         =   "입고 (F5)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin Threed.SSPanel Title 
         Height          =   930
         Left            =   5205
         TabIndex        =   2
         Top             =   75
         Width           =   4020
         _ExtentX        =   7091
         _ExtentY        =   1640
         _Version        =   262144
         Font3D          =   2
         ForeColor       =   16711680
         BackStyle       =   1
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "미출고현황"
         BorderWidth     =   0
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton Command1 
         Height          =   930
         Index           =   0
         Left            =   1785
         TabIndex        =   6
         Top             =   75
         Width           =   1650
         _Version        =   851970
         _ExtentX        =   2910
         _ExtentY        =   1640
         _StockProps     =   79
         Caption         =   "출고 (F6)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton Command1 
         Height          =   930
         Index           =   2
         Left            =   3495
         TabIndex        =   7
         Top             =   75
         Width           =   1650
         _Version        =   851970
         _ExtentX        =   2910
         _ExtentY        =   1640
         _StockProps     =   79
         Caption         =   "조회 (F7)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         Picture         =   "Main_Form.frx":80B52
      End
      Begin XtremeSuiteControls.PushButton Command1 
         Height          =   930
         Index           =   3
         Left            =   10860
         TabIndex        =   4
         Top             =   75
         Width           =   1650
         _Version        =   851970
         _ExtentX        =   2910
         _ExtentY        =   1640
         _StockProps     =   79
         Caption         =   "종료 (ESC)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         Picture         =   "Main_Form.frx":8142C
      End
      Begin XtremeSuiteControls.PushButton TagNo 
         Height          =   930
         Left            =   9285
         TabIndex        =   8
         Top             =   75
         Width           =   1515
         _Version        =   851970
         _ExtentX        =   2672
         _ExtentY        =   1640
         _StockProps     =   79
         Caption         =   "0-001"
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Menu PopUp 
      Caption         =   "POPup"
      Visible         =   0   'False
      Begin VB.Menu m_99998 
         Caption         =   "오름차순"
      End
      Begin VB.Menu m_99999 
         Caption         =   "내림차순"
      End
   End
End
Attribute VB_Name = "Main_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bMAIN_ACTIVATE As Boolean

Private Sub F5_Click()
    Call KeyChk(vbKeyF5)
End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 1 ' 입고
            If chkinputflig = "입고중" Then
                If Len(GetSpreadText(frm접수.sprGrid, 1, 1)) > 0 Then
                    If MsgBox("입고 작업을 취소 하시겠습니까?", vbInformation + vbYesNo, "종료 확인") = vbYes Then
                        chkinputflig = "출고중" '현재 상태..
                    Else
                        Exit Sub
                    End If
                End If
            End If
            Call KeyChk(vbKeyF5)
            
        Case 0 ' 출고
            If chkinputflig = "입고중" Then
                frm접수.sprGrid.Col = 1
                frm접수.sprGrid.Row = 1
                If Len(frm접수.sprGrid.Text) > 0 Then
                    If MsgBox("입고 작업을 취소 하시겠습니까?", vbInformation + vbYesNo, "종료 확인") = vbYes Then
                        chkinputflig = "출고중" '현재 상태..
                    Else
                        Exit Sub
                    End If
                End If
            End If
            Call KeyChk(vbKeyF6)
            
        Case 2 ' 조회
            If chkinputflig = "입고중" Then
                frm접수.sprGrid.Col = 1
                frm접수.sprGrid.Row = 1
                If Len(frm접수.sprGrid.Text) > 0 Then
                    If MsgBox("입고 작업을 취소 하시겠습니까?", vbInformation + vbYesNo, "종료 확인") = vbYes Then
                        chkinputflig = "조회중" '현재 상태..
                    Else
                        Exit Sub
                    End If
                End If
            End If
            Call KeyChk(vbKeyF7)
            
        Case 3 ' 종료
'            If chkinputflig = "입고중" Then
'                frm접수.sprGrid.Col = 1
'                frm접수.sprGrid.Row = 1
'                If Len(frm접수.sprGrid.Text) > 0 Then
'                    If MsgBox("입고 작업을 취소 하시겠습니까?     ", vbInformation + vbYesNo, "종료 확인") = vbYes Then
'                    Else
'                        Exit Sub
'                    End If
'                End If
'            End If
            Call KeyChk(vbKeyEscape)
            Main_Form.Command1(1).Enabled = True
            Main_Form.Command1(0).Enabled = True
            Main_Form.Command1(2).Enabled = True

            ' 서버모드가 아닐경우 입고를 할 수 없게 한다.
            If chkProgramMode <> ServerMode Then Main_Form.Command1(1).Enabled = False
    End Select
End Sub

Private Sub MDIForm_Activate()
    Dim i As Integer
    
    'On Error GoTo Err_Rtn
    
    ' 인터넷을 사용할 경우 업데이트 내용이 있을 경우 확인하여 처리한다.

    If bMAIN_ACTIVATE = False Then
        bMAIN_ACTIVATE = True
            
        DoEvents
        
        ' 프로그램의 버전을 설정한다.
        'If Format(Date, "yyyyMMdd") <= "20090610" Then
        '    Call SendProgramVersion
        'End If

        If GetSetting("Laundry_Zi", "UpDate", "Auto", "Y") = "Y" Then
            If GetSetting("Laundry_Zi", "Connect", "Type", "True") = False Then
                FormUpdateCheck.Show 1
            End If
        End If
        
        ' 기본 정보를 가저온다.
        If 대리점정보.StoreCode = "000000" Then
            FormStoreDefaultINFO.Show 1
        End If
        
        ' 신규 대리점 정보가 있을 경우
        Call Fb대리점정보
        
'        ' 시작일이 일주일 전일경우 까지 계속 전송한다.
'        ' 별도의 작업을 처리하기에는 무리가 있어 시작일부터 일주일간 계속전송한다.
'        ' 시작일날 개점을 하지않을 경우가 있을수 있어 이렇게 처리한다.
'        If 대리점정보.StartDate >= Format(DateAdd("d", -7, Date), "yyyyMMdd") Then
'            Call SendStoreDefaultInfo
'        End If
'
'        If 대리점정보.StoreCode <> "000000" Then
'            Call SendSalesData("2008-01-01", DateDiff("d", "2008-01-01", Date))
'        End If
        
    End If
    
    ' 서버모드가 아닐경우 입고를 할 수 없게 한다.
    If chkProgramMode <> ServerMode Then Command1(1).Enabled = False
    
    ' 전일 마감이 안되었을 경우 마감을 해야 사용할 수 있게함
    ' nDayCloseChk는 form6에서 마감되었을 경우 True
    ' 최근 일주일의 내용을 검사하여 마감여부를 확인한다.
    If Not nDayCloseChk Then
        strDayClose = "" '마감일자를 초기화 한다.
        
        For i = 1 To 7
            strDayClose = Format(DateAdd("d", -i, Date), "yyyymmdd")
            
            If Not DayCloseCheck(strDayClose) Then
                Query = "SELECT 입고일 FROM 입출고 WHERE 입고일 = '" & strDayClose & "'"
                Set ADORs = New ADODB.Recordset
                ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
                
                'Set Myrec = MyDB.OpenRecordset(Query)
                
                If ADORs.RecordCount >= 1 Then
                    Command1(0).Enabled = False
                    Command1(1).Enabled = False
                    Command1(2).Enabled = False
                    Command1(3).Enabled = False
                    
                    ADORs.Close
                    Set ADORs = Nothing
                    
                    chkinputflig = "조회중"
                    
                    Form6_OLD.Show
                    Form6_OLD.CmdMagam(1).Enabled = False
                    
                    bMsgMode = True
                    strMessage = "[ " & Format(strDayClose, "@@@@-@@-@@") & " ]일자을 마감하여 주십시요"
                    
                    Exit For
                End If
                ADORs.Close
                Set ADORs = Nothing
            End If
        Next i
        
        ' 일주일동안 마감이 없을 경우 참으로 한다.
        nDayCloseChk = True
    End If
    
    ' 출력할 메시지가 있을 경우 출력 한다.
    If bMsgMode Then
        MsgBox strMessage, vbInformation, "Laundry - 메시지"
        bMsgMode = False
    End If
    
    '------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------
    Query = "SELECT 마일리지검사일자 FROM 대리점정보 "
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If Trim(ADORs.Fields("마일리지검사일자") & "") = "" Then
        ADORs.Close
        Set ADORs = Nothing
        
        FormMileageCheck.Show
    End If
    
    If MailCheck = False Then
        ' 편지내역을 CHECK한다.
        Query = "SELECT Count(*) As 수량 "
        Query = Query & "FROM 메일 "
        Query = Query & "WHERE 송수신구분 = '2' "
        Query = Query & "AND   메일일자 = '" & Format(Date, "yyyymmdd") & "' "
        
        MailCheck = True
    End If
    
    '
    Call 참조코드_운동화_추가
    
    Timer1.Enabled = True
    
    Exit Sub
    
Err_Rtn:
    MsgBox "프로그램이 정상적으로 동작하지 않을 수 있습니다." & vbLf & vbLf & Err.Description, vbCritical, "확인"
    
   
End Sub

Private Sub MDIForm_Load()
    Dim strDate    As String
    Dim strWeekDay As Integer
    Dim strWD      As String
    Dim sTag       As String
    
    If App.PrevInstance = True Then
        Call ActivatePrevInstance(Me, M_CompnyMasterName)    '이전 실행 처리
        
        Exit Sub
    End If

    bMAIN_ACTIVATE = False
    
    '-------------------------------------------------------------------------
    If Screen.Width / Screen.TwipsPerPixelX > 1024 Then
        Main_Form.WindowState = 0
        
        Me.Width = 15480  'Screen.TwipsPerPixelX * 1024
        Me.Height = 11190 'Screen.TwipsPerPixelY * 768
    Else
        Main_Form.WindowState = 2
    End If
    
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
  
    ' 기본 프린터 여백정보 얻기
    Prt_Top = GetPrtStartPoint("TOP")
    Prt_Left = GetPrtStartPoint("LEFT")
    Prt_Height = GetPrtStartPoint("HEIGHT")
    
    sTag = GetTagNum("", "R")
    Main_Form.TagNo.Caption = GetTagNum(sTag, "+")
    
    'Tag_Load  '택번호
        
    '마감일자를 초기화 한다.
    nDayCloseChk = False
    strDayClose = ""

    chkinputflig = "메뉴"
    
    strDate = Format(Date, "yyyymmdd")
            
    '-------------------------------------------------------------
    '
    '-------------------------------------------------------------
    Query = "SELECT * "
    Query = Query & " FROM 할인정보 "
    Query = Query & " WHERE 시작일 <= '" & strDate & "' "
    Query = Query & "   AND 종료일 >= '" & strDate & "' "
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    'Set Myrec = MyDB.OpenRecordset(Query)
    
    If ADORs.RecordCount > 0 Then
        ADORs.Close
        Set ADORs = Nothing
        
        chkDaySale = False
        
        Exit Sub
    End If
    ADORs.Close
    Set ADORs = Nothing
    
    '-------------------------------------------------------------
    '
    '-------------------------------------------------------------
    Query = "SELECT 목요세일 "
    Query = Query & "FROM 대리점정보 "
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
        
    If ADORs.EOF Then
        strWeekDay = 0
    Else
        strWeekDay = Val(ADORs(0))
    End If
    ADORs.Close
    Set ADORs = Nothing
    
    If strWeekDay = Weekday(Date) Then
        chkDaySale = True
    Else
        chkDaySale = False
    End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    
    Unload frm작업
'    Tag_Save '택번호 저장
End Sub

Private Sub MDIForm_Resize()
    Command1(3).Left = Me.Width - Command1(3).Width - 200
    
    If Me.WindowState = vbMinimized Then
        Unload frm작업
        Unload frm색상표
        Unload frm의류
    End If
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
    bMAIN_ACTIVATE = False

End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As ComctlLib.Panel)
    If Panel.Index = 4 Then
        Dim sStr    As String
        'Dim Myrec As DAO.Recordset
        
        sStr = InputBox("관리자외 사용금지", "암호입력")
        
        If sStr = "dudtjsgh" Or sStr = "cleanaid" Or sStr = 대리점정보.전화2 Then
            'Set Myrec = MyDB.OpenRecordset("SELECT MAX(DGubun) from  DataBaseUpdate")
                        
            '------------------------------------------------------------------------------------
            Query = "SELECT MAX(DGubun) from  DataBaseUpdate"
            Set SUBRs = New ADODB.Recordset
            SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
            
            ADOCon.Execute "DELETE FROM DataBaseUpdate WHERE DGubun = '" & SUBRs.Fields(0) & "' "
            
            SUBRs.Close
            Set SUBRs = Nothing
        End If
    End If
End Sub

Private Sub TagNo_Click()
    frmTag.Show
End Sub

Private Sub Timer1_Timer()
    Dim sDate As String
    Dim dblFlage As Double
    Dim bSendFlage As Boolean
    Dim strRegCheck As String
    
    On Error GoTo Err_Handle
    
    '------------------------------------------------------------------------------------
    ' 메시지 관리 내용 추가
    '------------------------------------------------------------------------------------
    Query = "SELECT * FROM 메일 "
    Query = Query & " WHERE 송수신구분 = '2' "
    Query = Query & "   AND 조회시작일 <= '" & Format(Date, "yyyyMMdd") & "' "
    Query = Query & "   AND 조회종료일 >= '" & Format(Date, "yyyyMMdd") & "' "
'    Query = Query & "  AND not (수신일자 like '%" & Format(Date, "yyyy-MM-dd") & "%') "
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    Dim frm As Form
    
    Do Until ADORs.EOF
        If InStr(ADORs.Fields("수신일자") & "", Format(Date, "yyyy-MM-dd")) <= 0 Then
            If chkinputflig = "입고중" Then
                If frm접수.RichTextBox1.Tag <> "VIEW" Then
                    frm접수.Label6(1).Caption = "   " & Format(ADORs.Fields("조회시작일"), "@@@@-@@-@@") & " ~ " & Format(ADORs.Fields("조회종료일"), "@@@@-@@-@@")
                    
                    frm접수.Label6(2).Caption = ADORs.Fields("메일일자") & ""                         '
                    frm접수.Label6(3).Caption = ADORs.Fields("메일번호") & ""                         '
                    frm접수.RichTextBox1.Text = GetMailConvert(ADORs.Fields("메일내역") & "", "READ") '
                    frm접수.RichTextBox1.Tag = "VIEW"                                                 '
                    frm접수.pnlMessage.Visible = True                                                 '
                    frm접수.pnlMessage.ZOrder 0                                                       '
                End If
            End If
        End If
        
        ADORs.MoveNext
    Loop
    ADORs.Close
    Set ADORs = Nothing
    
    If dblFlage < 10 Then
        Query = "SELECT MAX(일자) AS MDATE "
        Query = Query & " FROM 일일마감 "
        Query = Query & " WHERE 마감여부  = 'Y'"
        Query = Query & "   AND len(일자) = 8 "
        Query = Query & "   AND 일자      < '" & Format(Date, "yyyyMMdd") & "' "
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
        
        'Set Myrec = MyDB.OpenRecordset(Query)
        
        If ADORs.EOF Then
            ADORs.Close
            Set ADORs = Nothing
            
            bSendFlage = True
            Title.ToolTipText = ""
            
            ' 바탕
            Title.BackColor = vbButtonFace '&HC0C0C0
            Exit Sub
            
        ElseIf ADORs.RecordCount <= 0 Then
            ADORs.Close
            Set ADORs = Nothing
            
            ' 최초 자료일 경우 암호를 확인한 것으로 한다.
            chkPassWord = True
            Exit Sub
            
        Else
            If IsNull(ADORs!MDATE) Then
                ADORs.Close
                Set ADORs = Nothing
                
                chkPassWord = True
                
                Exit Sub
            End If
            
            sDate = ADORs!MDATE & ""
        End If
        ADORs.Close
        Set ADORs = Nothing
        
        '---------------------------------------------------------
        '
        '---------------------------------------------------------
        Query = "SELECT 전송여부 "
        Query = Query & " FROM 일일마감 "
        Query = Query & " WHERE 일자 = '" & sDate & "' "
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
        
        If ADORs!전송여부 = "*" Then
            ADORs.Close
            Set ADORs = Nothing
            
            Title.ToolTipText = ""
            Title.Tag = ""
            
            bSendFlage = True
            chkPassWord = True
            
            Exit Sub
            
        Else
            Title.ToolTipText = "[" & Format(sDate, "0000-00-00") & "]일자의 데이터가 전송이 안되었습니다."
            Title.Tag = Format(sDate, "0000-00-00")
        End If
        ADORs.Close
        Set ADORs = Nothing
    End If
    
    dblFlage = dblFlage + 10
    
    If dblFlage >= 300 Then dblFlage = 0
    
    If bSendFlage = False Then
        If Title.BackColor = &HFF0000 Then
            Title.BackColor = vbButtonFace '&HC0C0C0 ' 바탕
        Else
            Title.BackColor = &HFF0000 ' 파랑
        End If
        
'        Timer1.Enabled = False
'        If MsgBox("[" & sDate & "]일자의 데이터가 전송이 안되었습니다.", vbInformation) = vbOK Then
'            Timer1.Enabled = True
'        End If
    End If
    
    ' 이미 암호를 확인 받았을 경우 취소.
    ' 프로그램을 시작했을 경우는 Fasle 이므로 한번은 실행 된다.
    If chkPassWord Then Exit Sub
    
    ' 일요일은 나타나지 않도록 한다.
    If Weekday(Date) = 1 Then Exit Sub
    
    ' 본사확인 코드의 유효 기간을 확인한다.
    strRegCheck = IsPassREGRead
    
    If strRegCheck = "-1" Or strRegCheck = "-2" Then
        If strRegCheck = "-1" Then
            ' 최초나 임의로 레지스터리를 기록 했을 경우
            MsgBox "일일 마감후 전송되지 않았습니다. 전송되지 않은 자료를 재전송 하시면 계속 사용할 수 있습니다. " & vbLf & vbLf & "재전송이 어려운 경우 본사에서 본사 확인 코드를 받으셔야 합니다." & vbLf, vbInformation, "자료 재전송"
        ElseIf strRegCheck = "-2" Then
            ' 유효 기간이 지간 경우
            MsgBox "본사확인의 유효기간이 만료 되었습니다.. 다시 확이 받으십시요.", vbInformation, "본사 확인 코드 입력"
        Else
            MsgBox "일일 마감후 전송되지 않았습니다. 재전송 하시거나, 본사에서 본사 확인 코드를 받으셔야 합니다." & vbLf & vbLf & "본사에 확인 바랍니다.", vbInformation, "본사 확인 코드 입력"
        End If
    Else
        chkPassWord = True
    End If
        
    If Not chkPassWord Then
        Timer1.Enabled = False
        
        Load Form22_OLD
        Form22_OLD.cmdBtn(2).Enabled = False
        Form22_OLD.Show 1
        Timer1.Enabled = True
        Exit Sub
    End If
    
    Exit Sub
    
Err_Handle:
    If Err.Number = "3061" Then
        Timer1.Enabled = False
        Exit Sub
    End If
End Sub


'--------------------------------------------------------------------------------------------------------------
' Procedure : 참조코드_운동화_추가
' DateTime  : 2007-04-13 03:31
' Author    : pds2004
' Purpose   : 운동화가 빠저있는 매장이 있다고 해서 강재로 등록
'--------------------------------------------------------------------------------------------------------------
Private Sub 참조코드_운동화_추가()
    On Error GoTo Error_Rtn

    If Format(Date, "yyyyMMdd") <= "20070414" Then
        Query = "SELECT 구분코드 FROM 참조코드 WHERE 구분코드 = 'a00' "
        Set SUBRs = New ADODB.Recordset
        SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
        
        'Set Myrec = MyDB.OpenRecordset(Query)
        
        If SUBRs.EOF Then
            Query = "INSERT INTO 참조코드(구분코드,품명,가격) "
            Query = Query & " VALUES('a00','운동화(우)', '1500')"
            ADOCon.Execute Query
        End If
    End If

    On Error GoTo 0
    
    Exit Sub

Error_Rtn:
    'sendErrormessage
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure 참조코드_운동화_추가 of Form Main_Form"

End Sub

Private Sub Title_Click()
'    Exit Sub
'
'    ' 프로그램의 버전을 설정한다.
'    Call SendProgramVersion
'
'
'    ' 기본 자료를 설정한다.
'    Call SetTableDefaultSendData
'
'    ' 본사에서 'N'를 설정한 자룔를 다시 전송한다.
'    Call SendNoSalesData
'
'    ' 각종 테이블의 자료를 전송한다.
'    ProgressBar1.Visible = True
'    Call SendTableData(ProgressBar1)
'    ProgressBar1.Visible = False
'
'    ' 마감 자료를 전송한다.
'    Call SendSalesData(Format(DateAdd("d", -7, Date), "yyyy-MM-dd"), 7)
End Sub

Private Sub m_99998_Click()
    Call Sort_Select(Me.ActiveForm.ActiveControl, 1, 1)
End Sub

Private Sub m_99999_Click()
    Call Sort_Select(Me.ActiveForm.ActiveControl, 2, 1)
End Sub
