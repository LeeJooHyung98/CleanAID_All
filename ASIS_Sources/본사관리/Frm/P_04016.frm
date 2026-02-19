VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{B6C10482-FB89-11D4-93C9-006008A7EED4}#1.0#0"; "TeeChart5.ocx"
Begin VB.Form P_04016 
   Caption         =   "일별 매출현황 (그래프)"
   ClientHeight    =   10635
   ClientLeft      =   3810
   ClientTop       =   1305
   ClientWidth     =   16080
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_04016.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10635
   ScaleWidth      =   16080
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   10635
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16080
      _ExtentX        =   28363
      _ExtentY        =   18759
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_04016.frx":058A
      Begin TeeChart.TChart TChart1 
         Height          =   9285
         Left            =   5655
         TabIndex        =   23
         Top             =   1335
         Width           =   10410
         Base64          =   $"P_04016.frx":063C
         Begin TeeChart.ChartPageNavigator ChartPageNavigator1 
            Height          =   345
            Left            =   1200
            Negotiate       =   -1  'True
            OleObjectBlob   =   "P_04016.frx":0DCC
            TabIndex        =   24
            Top             =   90
            Width           =   1200
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   9285
         Left            =   15
         TabIndex        =   1
         Top             =   1335
         Width           =   5625
         _Version        =   524288
         _ExtentX        =   9922
         _ExtentY        =   16378
         _StockProps     =   64
         BackColorStyle  =   1
         EditEnterAction =   5
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
         MaxCols         =   5
         ScrollBars      =   2
         SpreadDesigner  =   "P_04016.frx":0E1D
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   2
         Top             =   540
         Width           =   16050
         _ExtentX        =   28310
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Left            =   1245
            Style           =   2  '드롭다운 목록
            TabIndex        =   17
            Top             =   405
            Width           =   3420
         End
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1245
            Locked          =   -1  'True
            TabIndex        =   16
            Text            =   "cboOffice"
            Top             =   60
            Width           =   3420
         End
         Begin Threed.SSOption optGubun 
            Height          =   360
            Index           =   0
            Left            =   6660
            TabIndex        =   14
            Top             =   45
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   635
            _Version        =   262144
            Caption         =   "일별"
            Value           =   -1
         End
         Begin Threed.SSPanel pnlTitle 
            Height          =   315
            Left            =   5430
            TabIndex        =   3
            Top             =   405
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "수금년월"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSOption optGubun 
            Height          =   360
            Index           =   2
            Left            =   7635
            TabIndex        =   15
            Top             =   45
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   635
            _Version        =   262144
            Caption         =   "월별"
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   18
            Top             =   405
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "가맹점명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   7
            Left            =   60
            TabIndex        =   19
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "지    사"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin XtremeSuiteControls.DateTimePicker dtInput 
            Height          =   315
            Left            =   6600
            TabIndex        =   20
            Top             =   405
            Width           =   1170
            _Version        =   851970
            _ExtentX        =   2064
            _ExtentY        =   556
            _StockProps     =   68
            CustomFormat    =   "yyyy-MM"
            Format          =   3
            UpDown          =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   5430
            TabIndex        =   21
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "조    건"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton cmdRefresh 
            Height          =   330
            Left            =   4695
            TabIndex        =   22
            Top             =   390
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   582
            _StockProps     =   79
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_04016.frx":145E
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   4
         Top             =   15
         Width           =   8445
         _ExtentX        =   14896
         _ExtentY        =   900
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   255
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
         Caption         =   " #"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_04016.frx":19F8
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   1
         Left            =   8475
         TabIndex        =   5
         Top             =   15
         Width           =   7590
         _ExtentX        =   13388
         _ExtentY        =   900
         _Version        =   262144
         Font3D          =   3
         ForeColor       =   192
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
         PictureBackgroundStyle=   2
         PictureBackground=   "P_04016.frx":1BFA
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   6
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "종료"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_04016.frx":1DFC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   7
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "엑셀"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_04016.frx":2396
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   8
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "인쇄"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_04016.frx":2930
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   9
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "취소"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_04016.frx":2ECA
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   10
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "삭제"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_04016.frx":3464
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   11
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "저장"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_04016.frx":39FE
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   12
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "신규"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_04016.frx":3F98
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   13
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "조회"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_04016.frx":4532
         End
      End
   End
End
Attribute VB_Name = "P_04016"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cboOffice_Click()
    If cboOffice.ListIndex < 0 Then Exit Sub
    
    '-----------------------------------------------------------------
    '
    '-----------------------------------------------------------------
    cboInput.Clear
    
    ReDim sValue(2)
    
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    
    If optGubun(0).Value = True Then
        sValue(1) = Format(dtInput.Value, "YYYY-MM-01")
        sValue(2) = Format(dtInput.Value, "YYYY-MM-31")
        
    Else
        sValue(1) = Format(dtInput.Value, "YYYY-01-01")
        sValue(2) = Format(dtInput.Value, "YYYY-12-31")
    End If
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01001_00", sValue(), Err_Num, Err_Dec)
    
    cboInput.AddItem "[000000] 전체"
    
    Do Until RS01.EOF
        cboInput.AddItem "[" & RS01!가맹점코드 & "] " & RS01!가맹점명
        
        RS01.MoveNext
    Loop
    RS01.Close
    Set RS01 = Nothing
    
    If cboInput.ListCount > 0 Then cboInput.ListIndex = 0
End Sub

Private Sub cboOffice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SearchString_한글 KeyAscii
'    Else
       ' SearchString KeyAscii
    End If


End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0:    ' 조회
        
            Select Case True
                Case optGubun(0).Value: Call Data_Display
                'Case optGubun(1).Value: Call Data_Display1
                Case optGubun(2).Value: Call Data_Display2
                    
            End Select
    
        Case 1: 'Call DataAdd        ' 신규
        Case 2: 'Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
        Case 4: 'Call DataCancel     ' 취소
        Case 5: Call DataPrint      ' 인쇄
        Case 6: Call Export_Excel(P_00000.cdgExcel, spdView)      ' 엑셀
        Case 7: Unload Me           ' 종료
    End Select
    
'    Me.MousePointer = 0
    
    Exit Sub
    
ErrRtn:
    Me.MousePointer = 0
    
    If Err.Number = "0" Then
        
    ElseIf Err.Number = "91" Then
        End
    Else
        Resume Next
    End If
End Sub

Private Sub DataPrint()
    TChart1.Printer.ShowPreview
End Sub

Private Sub cmdRefresh_Click()
    cboOffice_Click
End Sub

Private Sub dtInput_Change()
    Call Data_Display
End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(5).Enabled = True
    cmdBtn(6).Enabled = True
    
    '본사 여부
    If HeadOffice = MASTER_OFFICE_CODE Then
        cboOffice.Locked = False
    Else
        cboOffice.Locked = True
    End If
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    With spdView
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

    dtInput.Value = Format(Date, "YYYY-MM")
        
    Call Get_지사리스트(cboOffice)
    
    Dim i As Integer
    
    With cboOffice
        For i = 0 To .ListCount - 1
            If Mid(.List(i), 2, 4) = HeadOffice Then
                .ListIndex = i
                
                Exit For
            End If
        Next i
    End With
    
    
''    Call ChartInit
''
''    ReDim sValue(1)
''
''    sValue(0) = "1"
''
''    Set RS01 = New ADODB.Recordset
''    Set RS01 = ExecPro("SP_04016_00", sValue(), Err_Num, Err_Dec)
''
''    spdView.MaxCols = RS01.Fields.Count
''    spdView.MaxRows = RS01.RecordCount
''
''    Call spdDisplay(RS01)
''    Call GetColWidth(REG_App, Me.Name, spdView)
''
''    Call optSelect_Click(0, True)
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i As Integer

    ReDim sValue(2)

    Screen.MousePointer = vbHourglass

    sValue(0) = Mid(cboOffice.Text, 2, 4)           '0
    If sValue(0) = "0000" Then sValue(0) = "%"

    If Mid(cboInput.Text, 2, 6) = "000000" Then
        sValue(1) = ""                              '1
    Else
        sValue(1) = Mid(cboInput.Text, 2, 6)        '1
    End If

    sValue(2) = Format(dtInput.Value, "YYYY-MM-DD") '2

    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(HeadOffice) = False Then

            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_04016_00", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_04016_00", sValue(), Err_Num, Err_Dec)
    End If

    With spdView
        .MaxRows = 0
        .Redraw = False

        TChart1.Series(0).Clear
        TChart1.Series(1).Clear

        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows

            .Col = 1: .Text = Right(RS01!마감일자, 2)

            .Col = 2: .Text = RS01!전년접수금액 & ""
            .Col = 3: .Text = RS01!금년접수금액 & ""

            .Col = 4: .Text = RS01!전년평균단가 & ""
            .Col = 5: .Text = RS01!금년평균단가 & ""

            TChart1.Series(0).Add RS01!전년접수금액, Right(RS01!마감일자, 2), vbRed
            TChart1.Series(1).Add RS01!금년접수금액, Right(RS01!마감일자, 2), vbBlue

            RS01.MoveNext
        Loop

        RS01.Close
        Set RS01 = Nothing


        Call SpreadSum(spdView, 1, 2)
        Call SpreadSum(spdView, -1, 3)
'
'        If .MaxRows > 0 Then
'            .MaxRows = .MaxRows + 1
'            .Row = .MaxRows
'
'            .Col = 1: .Text = "합계"
'            .Col = 2: .Formula = "SUM(B1:B" & .MaxRows - 1 & ")"
'            .Col = 3: .Formula = "SUM(C1:C" & .MaxRows - 1 & ")"
'        End If

        .Redraw = True
    End With

    ChartPageNavigator1.ChartLink = TChart1.ChartLink
    ChartPageNavigator1.EnableButtons

    Screen.MousePointer = vbDefault
    Exit Sub

ErrRtn:
    Screen.MousePointer = vbDefault
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Data_Display1()
    On Error GoTo ErrRtn

    Dim i As Integer

    ReDim sValue(3)

    Screen.MousePointer = vbHourglass
    sValue(0) = Mid(cboOffice.Text, 2, 4)

    If Mid(cboInput.Text, 2, 6) = "000000" Then
        sValue(1) = ""
    Else
        sValue(1) = Mid(cboInput.Text, 2, 6)
    End If

    sValue(2) = Format(dtInput.Value, "YYYY-MM-DD")

    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(HeadOffice) = False Then

            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_04016_01", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_04016_01", sValue(), Err_Num, Err_Dec)
    End If

    With spdView
        .MaxRows = 0
        .Redraw = False

        TChart1.Series(0).Clear
        TChart1.Series(1).Clear

        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows

            .Col = 1: .Text = Right(RS01!마감일자, 2)

            .Col = 2: .Text = RS01!전년접수금액 & ""
            .Col = 3: .Text = RS01!금년접수금액 & ""

            .Col = 4: .Text = RS01!전년평균단가 & ""
            .Col = 5: .Text = RS01!금년평균단가 & ""

            TChart1.Series(0).Add RS01!전년접수금액, Right(RS01!마감일자, 2), vbRed
            TChart1.Series(1).Add RS01!금년접수금액, Right(RS01!마감일자, 2), vbBlue

            RS01.MoveNext
        Loop

        RS01.Close
        Set RS01 = Nothing


        Call SpreadSum(spdView, 1, 2)
        Call SpreadSum(spdView, -1, 3)

'        If .MaxRows > 0 Then
'            .MaxRows = .MaxRows + 1
'            .Row = .MaxRows
'
'            .Col = 1: .Text = "합계"
'            .Col = 2: .Formula = "SUM(B1:B" & .MaxRows - 1 & ")"
'            .Col = 3: .Formula = "SUM(C1:C" & .MaxRows - 1 & ")"
'        End If

        .Redraw = True
    End With

    ChartPageNavigator1.ChartLink = TChart1.ChartLink
    ChartPageNavigator1.EnableButtons
    Screen.MousePointer = vbDefault

    Exit Sub

ErrRtn:
    Screen.MousePointer = vbDefault
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Data_Display2()
    On Error GoTo ErrRtn

    Dim i As Integer

    ReDim sValue(2)

    Screen.MousePointer = vbHorizontal
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    If sValue(0) = "0000" Then sValue(0) = "%"

    If Mid(cboInput.Text, 2, 6) = "000000" Then
        sValue(1) = ""
    Else
        sValue(1) = Mid(cboInput.Text, 2, 6)
    End If

    sValue(2) = Format(dtInput.Value, "YYYY-MM-DD")

    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(HeadOffice) = False Then

            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_04016_02", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_04016_02", sValue(), Err_Num, Err_Dec)
    End If

    With spdView
        .MaxRows = 0
        .Redraw = False

        TChart1.Series(0).Clear
        TChart1.Series(1).Clear

        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows

            .Col = 1: .Text = Right(RS01!마감월, 2)

            .Col = 2: .Text = RS01!전년접수금액 & ""
            .Col = 3: .Text = RS01!금년접수금액 & ""

            .Col = 4: .Text = RS01!전년평균단가 & ""
            .Col = 5: .Text = RS01!금년평균단가 & ""

            TChart1.Series(0).Add RS01!전년접수금액, Right(RS01!마감월, 2), vbRed
            TChart1.Series(1).Add RS01!금년접수금액, Right(RS01!마감월, 2), vbBlue

            RS01.MoveNext
        Loop

        RS01.Close
        Set RS01 = Nothing

        Call SpreadSum(spdView, 1, 2)
        Call SpreadSum(spdView, -1, 3)

'        If .MaxRows > 0 Then
'            .MaxRows = .MaxRows + 1
'            .Row = .MaxRows
'
'            .Col = 1: .Text = "합계"
'            .Col = 2: .Formula = "SUM(B1:B" & .MaxRows - 1 & ")"
'            .Col = 3: .Formula = "SUM(C1:C" & .MaxRows - 1 & ")"
'        End If

        .Redraw = True
    End With

    ChartPageNavigator1.ChartLink = TChart1.ChartLink
    ChartPageNavigator1.EnableButtons
    Screen.MousePointer = vbDefault

    Exit Sub

ErrRtn:
    Screen.MousePointer = vbDefault
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub ChartInit()

End Sub

Private Sub optSelect_Click(Index As Integer, Value As Integer)
    Select Case Index
        Case 0
            cboInput.Clear
            
            cboInput.AddItem "막대형 / 그림 그래프"
            cboInput.AddItem "꺽은선형"
            cboInput.AddItem "영역형"
            cboInput.AddItem "단계"
            cboInput.AddItem "혼합형"
            cboInput.AddItem "원형"
            cboInput.AddItem "XY(분산형)"
        Case 1
            cboInput.Clear
            
            cboInput.AddItem "막대형(열)"
            cboInput.AddItem "꺽은선형(테입프)"
            cboInput.AddItem "영역형"
            cboInput.AddItem "단계"
            cboInput.AddItem "혼합형"
    End Select
    
    cboInput.ListIndex = 0
End Sub

Private Sub cboInput_Click()
    Select Case True
        Case optGubun(0).Value
            pnlTitle.Caption = "조회년월"
            Call Data_Display
        
        Case optGubun(1).Value
            pnlTitle.Caption = "조회년도"
            Call Data_Display1
            
        Case optGubun(2).Value
            pnlTitle.Caption = "조회년도"
            Call Data_Display2
    End Select
End Sub

Private Sub optGubun_Click(Index As Integer, Value As Integer)
    Select Case Index
        Case 0
            pnlTitle.Caption = "조회년월"
            Call Data_Display
        
        Case 1:
            pnlTitle.Caption = "조회년도"
            Call Data_Display1
            
        Case 2:
            pnlTitle.Caption = "조회년도"
            Call Data_Display2
    End Select
End Sub

Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If NewRow <> -1 Then
        With spdView
            If NewRow <> -1 Then
                .Row = Row
                If (Row Mod 2) = 0 Then
                    .Col = -1
                    .BackColor = vbWhite
                Else
                    .Col = -1
                    .BackColor = vbWhite
                End If
                
                .Row = NewRow
                .Col = -1
                .BackColor = glbYellow
            End If
        End With
    End If

End Sub
