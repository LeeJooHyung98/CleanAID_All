VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_04025 
   Caption         =   "실시간 접수금액 확인"
   ClientHeight    =   12255
   ClientLeft      =   1455
   ClientTop       =   2775
   ClientWidth     =   16890
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_04025.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12255
   ScaleWidth      =   16890
   WindowState     =   2  '최대화
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7050
      Top             =   690
   End
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   12255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16890
      _ExtentX        =   29792
      _ExtentY        =   21616
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_04025.frx":058A
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   1
         Top             =   540
         Width           =   16860
         _ExtentX        =   29739
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1245
            Locked          =   -1  'True
            Style           =   2  '드롭다운 목록
            TabIndex        =   15
            Top             =   60
            Width           =   2850
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Left            =   1245
            TabIndex        =   2
            Top             =   420
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            Format          =   21430272
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   3
            Top             =   420
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "접수일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "지    사"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit1 
            Height          =   285
            Left            =   4500
            TabIndex        =   19
            Top             =   120
            Width           =   435
            _Version        =   851970
            _ExtentX        =   767
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin VB.CheckBox Check1 
            Caption         =   "    분후 새로 고침"
            Height          =   285
            Left            =   4260
            TabIndex        =   20
            Top             =   120
            Width           =   2355
         End
         Begin VB.Label Label_time 
            Caption         =   "Label1"
            Height          =   255
            Left            =   4500
            TabIndex        =   21
            Top             =   480
            Visible         =   0   'False
            Width           =   2685
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   4
         Top             =   15
         Width           =   9255
         _ExtentX        =   16325
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
         Caption         =   " 실시간 접수금액 확인 (P_04025)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_04025.frx":065C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Left            =   9285
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
         PictureBackground=   "P_04025.frx":085E
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
            Picture         =   "P_04025.frx":0A60
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
            Picture         =   "P_04025.frx":0FFA
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
            Picture         =   "P_04025.frx":1594
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
            Picture         =   "P_04025.frx":1B2E
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
            Picture         =   "P_04025.frx":20C8
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
            Picture         =   "P_04025.frx":2662
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
            Picture         =   "P_04025.frx":2BFC
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
            Picture         =   "P_04025.frx":3196
         End
      End
      Begin FPSpreadADO.fpSpread spdView1 
         Height          =   10905
         Left            =   10725
         TabIndex        =   14
         Top             =   1335
         Width           =   6150
         _Version        =   524288
         _ExtentX        =   10848
         _ExtentY        =   19235
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
         MaxCols         =   11
         SpreadDesigner  =   "P_04025.frx":3730
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdViewM 
         Height          =   10905
         Left            =   15
         TabIndex        =   17
         Top             =   1335
         Width           =   5055
         _Version        =   524288
         _ExtentX        =   8916
         _ExtentY        =   19235
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
         MaxCols         =   4
         MaxRows         =   35
         ScrollBars      =   2
         SpreadDesigner  =   "P_04025.frx":3ECD
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdViewS 
         Height          =   10905
         Left            =   5085
         TabIndex        =   18
         Top             =   1335
         Width           =   5625
         _Version        =   524288
         _ExtentX        =   9922
         _ExtentY        =   19235
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
         MaxCols         =   4
         MaxRows         =   35
         ScrollBars      =   2
         SpreadDesigner  =   "P_04025.frx":44A3
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_04025"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim RS02 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cboOffice_Click()
    Call Data_Display_Master
    Call Data_Display
End Sub

Private Sub Check1_Click()
    Timer1.Tag = Now
    Timer1.Enabled = IIf(Check1.Value = 1, True, False)
    
    Label_time.Visible = Timer1.Enabled
    Label_time.Caption = CStr(DateDiff("s", Timer1.Tag, Now)) & "초후 새로고침"
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0
            Call Data_Display_Master   ' 조회
            Call Data_Display
            Call spdViewS_Click(1, 1)
        Case 1: 'Call DataAdd        ' 신규
        Case 2: 'Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
        Case 4: 'Call DataCancel     ' 취소
        Case 5: 'Call DataPrint      ' 인쇄
        Case 6: Call Export_Excel(P_00000.cdgExcel, spdView1)      ' 엑셀
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

Private Sub dtInput_Change()
    dtInput.Enabled = False
    DoEvents
    
    Call cmdBtn_Click(0)
    
    dtInput.Enabled = True
    dtInput.SetFocus
End Sub

Private Sub FlatEdit1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then FlatEdit1_LostFocus
End Sub

Private Sub FlatEdit1_LostFocus()
    
    If IsNumeric(FlatEdit1.Text) = False Then
        FlatEdit1.Text = ""
        Exit Sub
    End If

ERR_RTN:

End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(5).Enabled = True
    cmdBtn(6).Enabled = True
        
    cboOffice.Locked = True
    cboOffice.Enabled = False
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"

    If P_04025_Flag = False Then
        dtInput.Value = Date
        
        P_04025_Flag = True
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    dtInput.Value = Date
    
    With spdViewM
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
        '.UserColAction = UserColActionSort
    End With
    
    With spdViewS
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
        '.UserColAction = UserColActionSort
    End With
    
    With spdView1
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
    
    Call Get_지사리스트(cboOffice)
    
    Dim i As Integer
    
    With cboOffice
        For i = 0 To .ListCount - 1
            If HeadOffice = "1000" Then
                .ListIndex = 0
                Exit For
            End If
            If Mid(.List(i), 2, 4) = HeadOffice Then
                .ListIndex = i
                
                Exit For
            End If
        Next i
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_04025_Flag = False
End Sub

Private Sub Data_Display()
    Dim lAmt As Long
    
    On Error GoTo ErrRtn
    
    '-------------------------------------------------------------
    ' SP_02002_00
    '-------------------------------------------------------------
    ReDim sValue(1)
    
    Screen.MousePointer = vbHourglass
    spdViewM.Row = spdViewM.ActiveRow
    spdViewM.Col = 1
    
    sValue(0) = spdViewM.Text
    sValue(1) = Format(dtInput.Value, "YYYY-MM-DD")
        
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(sValue(0)) = False Then Exit Sub

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_04025_00", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_04025_00", sValue(), Err_Num, Err_Dec)
    End If
    
    With spdViewS
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = Trim(RS01!가맹점코드) & ""
            .Col = 2: .Text = Trim(RS01!가맹점명) & ""
            .Col = 3: .Text = RS01!금액 & ""
            .Col = 4: .Text = RS01!건수 & ""
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        If .MaxRows > 0 Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = -1: .BackColor = &HD8FCFE
            
            .Col = 1: .Text = "합계"
            .Col = 3: .Formula = "SUM(C1:C" & .MaxRows - 1 & ")"
            .Col = 4: .Formula = "SUM(D1:D" & .MaxRows - 1 & ")"
        End If
        
        .Redraw = True
    End With
    
    Screen.MousePointer = vbDefault
    Exit Sub
        
ErrRtn:
    Screen.MousePointer = vbDefault
    dtInput.Enabled = True
    
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub


Private Sub Data_Display_Master()
    Dim lAmt As Long
    
    On Error GoTo ErrRtn
    
    '-------------------------------------------------------------
    ' SP_02002_00
    '-------------------------------------------------------------
    ReDim sValue(1)
    Screen.MousePointer = vbHourglass
    
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    sValue(1) = Format(dtInput.Value, "YYYY-MM-DD")
        
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(HeadOffice) = False Then Exit Sub

        Set RS02 = New ADODB.Recordset
        Set RS02 = ExecProMaster("SP_04025_01", sValue(), Err_Num, Err_Dec)
    Else
        Set RS02 = New ADODB.Recordset
        Set RS02 = ExecPro("SP_04025_01", sValue(), Err_Num, Err_Dec)
    End If
    
    With spdViewM
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS02.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = Trim(RS02!지사코드) & ""
            .Col = 2: .Text = Trim(RS02!지사명) & ""
            .Col = 3: .Text = RS02!금액 & ""
            .Col = 4: .Text = RS02!건수 & ""
            
            RS02.MoveNext
        Loop
        RS02.Close
        Set RS01 = Nothing
        
        If .MaxRows > 0 Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = -1: .BackColor = &HD8FCFE
            
            .Col = 1: .Text = "합계"
            .Col = 3: .Formula = "SUM(C1:C" & .MaxRows - 1 & ")"
            .Col = 4: .Formula = "SUM(D1:D" & .MaxRows - 1 & ")"
        End If
        
        .Redraw = True
    End With
    
    Screen.MousePointer = vbDefault
    Exit Sub
        
ErrRtn:
    Screen.MousePointer = vbDefault
    dtInput.Enabled = True
    
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Public Sub DataPrint()
'    Dim ReportFP As String
'    Dim ReportFile As String
'
'    ReportFP = GetIniStr("REPORT", "FilePath", "", m_iniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    Call PrintDesc
'
'    Dim i As Integer
'
'    For i = 0 To 30
'        P_00000.crPrint.Formulas(i) = ""
'    Next
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'    P_00000.crPrint.Formulas(0) = "입고일자 = '" & Format(dtInput.Value, "YYYY-MM-DD") & "'"
'
'    Call ReportPrint(ReportFile, "1")
End Sub

Public Sub DataScreen()
'    Dim ReportFP As String
'    Dim ReportFile As String
'
'    ReportFP = GetIniStr("REPORT", "FilePath", "", m_iniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    Call PrintDesc
'
'    Dim i As Integer
'
'    For i = 0 To 30
'        P_00000.crPrint.Formulas(i) = ""
'    Next
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'    P_00000.crPrint.Formulas(0) = "입고일자 = '" & Format(dtInput.Value, "YYYY-MM-DD") & "'"
'
'    Call ReportPrint(ReportFile, "2")
End Sub

Private Sub PrintDesc()
'    Dim i, j As Integer
'
'    Dim TempText As String
'    Dim TempFP As String
'    Dim TempFile As String
'
'    TempFP = GetIniStr("REPORT", "TempPath", "", m_iniFile)
'    TempFile = TempFP & "\Temp.txt"
'
'    Open TempFile For Output As #1
'
'    TempText = ""
'
'    For j = 0 To 2
'        For i = 1 To spdView(j).MaxRows
'            spdView(j).Row = i
'
'            spdView(j).Col = 1: TempText = LeftH(spdView(j).Text & Space(20), 20)
'            spdView(j).Col = 2: TempText = TempText & RightH(Space(9) & spdView(j).Text, 9)
'            spdView(j).Col = 3: TempText = TempText & RightH(Space(9) & spdView(j).Text, 9)
'            spdView(j).Col = 4: TempText = TempText & RightH(Space(9) & spdView(j).Text, 9) & Space(4)
'
'            If spdView(j).Text = "" Then
'                Close #1
'                Exit For
'            End If
'
'            i = i + 1
'            spdView(j).Row = i
'
'            spdView(j).Col = 1: TempText = TempText & LeftH(spdView(j).Text & Space(20), 20)
'            spdView(j).Col = 2: TempText = TempText & RightH(Space(9) & spdView(j).Text, 9)
'            spdView(j).Col = 3: TempText = TempText & RightH(Space(9) & spdView(j).Text, 9)
'            spdView(j).Col = 4: TempText = TempText & RightH(Space(9) & spdView(j).Text, 9)
'
'            Print #1, TempText
'
'            If spdView(j).Text = "" Then
'                Exit For
'            End If
'        Next i
'    Next j
'    Close #1
End Sub

Public Sub DataAdd()

End Sub

Private Sub spdViewM_Click(ByVal Col As Long, ByVal Row As Long)
    Call Data_Display
End Sub

Private Sub spdViewS_Click(ByVal Col As Long, ByVal Row As Long)
    Dim 가맹점코드 As String
    Dim 지사코드    As String
    
    On Error GoTo ErrRtn
    
    If Row <= 0 Then Exit Sub
    
    spdViewS.Row = Row
    spdViewS.Col = 1: 가맹점코드 = spdViewS.Text & ""
    
    If 가맹점코드 = "합계" Then
        spdView1.MaxRows = 0
        Exit Sub
    End If
    
    spdViewM.Row = spdViewM.ActiveRow
    spdViewM.Col = 1
    
    지사코드 = spdViewM.Text
    
    '-------------------------------------------------------------
    ' SP_02002_00
    '-------------------------------------------------------------
    ReDim sValue(1)
    
    sValue(0) = 가맹점코드
    sValue(1) = Format(dtInput.Value, "YYYY-MM-DD")
        
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(지사코드) = False Then Exit Sub

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_02002_01", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_02002_01", sValue(), Err_Num, Err_Dec)
    End If
            
    With spdView1
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = RS01!접수일자 & ""
            .Col = 2:  .Text = RS01!접수시간 & ""
            
            .Col = 3:  .Text = Format(RS01!택번호, "000-00-0000") & ""
            .Col = 4:  .Text = Trim(RS01!의류명) & ""
            .Col = 5:  .Text = Trim(RS01!색상) & ""
            .Col = 6:  .Text = Trim(RS01!내용) & ""
            .Col = 7:  .Text = RS01!금액 & ""
            .Col = 8:  .Text = RS01!의류금액 & ""
            .Col = 9:  .Text = RS01!차액 & ""
            .Col = 10: .Text = RS01!수선금액 & ""
            .Col = 11: .Text = RS01!전화번호 & ""
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
    
    Exit Sub
    
ErrRtn:

    dtInput.Enabled = True
End Sub


Private Sub UpDown1_DownClick()

End Sub

Private Sub UpDown1_UpClick()

End Sub

Private Sub Timer1_Timer()

    Debug.Print Now
    
    If Trim(Timer1.Tag) = "" Then Exit Sub
    If DateDiff("n", Timer1.Tag, Now) < Val(FlatEdit1.Text) Then
        Label_time.Caption = CStr(DateDiff("s", Timer1.Tag, Now)) & "초후 새로고침"
        Exit Sub
    End If

    Debug.Print Now & "   Run"
    Timer1.Tag = Now

    Screen.MousePointer = vbHourglass
    Call Data_Display_Master
    Call Data_Display
    Call spdViewS_Click(1, 1)
    Debug.Print Now

    Screen.MousePointer = vbDefault
End Sub
