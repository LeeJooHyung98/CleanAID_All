VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_03007 
   Caption         =   "미출고 정리"
   ClientHeight    =   12555
   ClientLeft      =   -615
   ClientTop       =   2460
   ClientWidth     =   17355
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_03007.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12555
   ScaleWidth      =   17355
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   12555
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17355
      _ExtentX        =   30612
      _ExtentY        =   22146
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_03007.frx":058A
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   11205
         Left            =   15
         TabIndex        =   1
         Top             =   1335
         Width           =   17325
         _Version        =   524288
         _ExtentX        =   30559
         _ExtentY        =   19764
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
         MaxCols         =   15
         SpreadDesigner  =   "P_03007.frx":061C
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   2
         Top             =   540
         Width           =   17325
         _ExtentX        =   30559
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1245
            Locked          =   -1  'True
            Style           =   2  '드롭다운 목록
            TabIndex        =   19
            Top             =   60
            Width           =   3330
         End
         Begin VB.ComboBox cboInput 
            Height          =   315
            Left            =   1245
            Style           =   2  '드롭다운 목록
            TabIndex        =   3
            Top             =   420
            Width           =   3330
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   6135
            TabIndex        =   4
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Format          =   64159744
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   5
            Top             =   420
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
            Index           =   2
            Left            =   4950
            TabIndex        =   6
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "접수일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   9375
            TabIndex        =   7
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Format          =   64159744
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   10
            Left            =   60
            TabIndex        =   20
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "지    사"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin XtremeSuiteControls.ProgressBar ProgressBar 
            Height          =   330
            Left            =   9360
            TabIndex        =   21
            Top             =   405
            Visible         =   0   'False
            Width           =   3030
            _Version        =   851970
            _ExtentX        =   5345
            _ExtentY        =   582
            _StockProps     =   93
            Scrolling       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton cmdRefresh 
            Height          =   330
            Left            =   4605
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
            Picture         =   "P_03007.frx":0E94
         End
         Begin VB.Label Label1 
            Alignment       =   2  '가운데 맞춤
            Caption         =   "~"
            Height          =   255
            Left            =   9135
            TabIndex        =   8
            Top             =   120
            Width           =   255
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   9
         Top             =   15
         Width           =   9720
         _ExtentX        =   17145
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
         Caption         =   " 미출고 정리 (P_03007) - 지사출고일자 기준"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_03007.frx":142E
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   1
         Left            =   9750
         TabIndex        =   10
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
         PictureBackground=   "P_03007.frx":1630
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   11
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
            Picture         =   "P_03007.frx":1832
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   12
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
            Picture         =   "P_03007.frx":1DCC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   13
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
            Picture         =   "P_03007.frx":2366
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   14
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
            Picture         =   "P_03007.frx":2900
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   15
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
            Picture         =   "P_03007.frx":2E9A
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   16
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
            Picture         =   "P_03007.frx":3434
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   17
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
            Picture         =   "P_03007.frx":39CE
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   18
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
            Picture         =   "P_03007.frx":3F68
         End
      End
   End
End
Attribute VB_Name = "P_03007"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cboInput_Click()
    Call Data_Display
End Sub

Private Sub cboOffice_Click()
    If cboOffice.ListIndex < 0 Then Exit Sub
    
    '-----------------------------------------------------------------
    '
    '-----------------------------------------------------------------
    cboInput.Clear
    
    ReDim sValue(2)
    
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01001_00", sValue(), Err_Num, Err_Dec)
    
    Do Until RS01.EOF
        cboInput.AddItem "[" & RS01!가맹점코드 & "] " & RS01!가맹점명
        RS01.MoveNext
    Loop
    RS01.Close
    Set RS01 = Nothing
    
    If cboInput.ListCount > 0 Then cboInput.ListIndex = 0
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: Call Data_Display   ' 조회
        Case 1: 'Call DataAdd        ' 신규
        Case 2: 'Call DataSave       ' 저장
        Case 3: Call DataDelete     ' 삭제
        Case 4: 'Call DataCancel     ' 취소
        Case 5: 'Call DataPrint      ' 인쇄
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

Private Sub cmdRefresh_Click()
    cboOffice_Click
End Sub

Private Sub dtInput_Change(Index As Integer)
'    Call Data_Display
End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(3).Enabled = True
    cmdBtn(5).Enabled = True
    cmdBtn(6).Enabled = True
    
    '본사 여부
    If HeadOffice = MASTER_OFFICE_CODE Then
        cboOffice.Locked = False
    Else
        cboOffice.Locked = True
    End If
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")" & " - 지사출고일자 기준"
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    With spdView
        .MaxRows = 500
        .RowHeight(-1) = 14
                
        .Col = 11:  .ColMerge = MergeAlways
        .Col = 12:  .ColMerge = MergeRestricted
        .Col = 13: .ColMerge = MergeRestricted
                
                
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

    dtInput(0).Value = Date
    dtInput(1).Value = Date
    
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
    
'    If P_03007_Flag = False Then
'''        Call AgencyComboAdd(cboInput)
'''
'''        ReDim sValue(3)
'''
'''        sValue(0) = "1"
'''
'''        Set RS01 = New ADODB.Recordset
'''        Set RS01 = ExecPro("SP_03007_00", sValue(), Err_Num, Err_Dec)
'''
'''        spdView.MaxCols = RS01.Fields.Count
'''        spdView.MaxRows = RS01.RecordCount
'''
'''        Call spdDisplay(RS01)
'''        Call GetColWidth(REG_App, Me.Name, spdView)
'
'        P_03007_Flag = True
'    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_03007_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    If cboInput.Text = "" Then
        MsgBox "가맹점을 선택하십시오", vbInformation
        Exit Sub
    End If
    
    ReDim sValue(2)
    
    sValue(0) = Format(dtInput(0).Value, "YYYY-MM-DD") '
    sValue(1) = Format(dtInput(1).Value, "YYYY-MM-DD") '
    sValue(2) = Mid(cboInput.Text, 2, 6)               '
        
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_03007_00", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_03007_00", sValue(), Err_Num, Err_Dec)
    End If
    
    
    With spdView
        .MaxRows = 0
        .Redraw = False
        
        Do While Not RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = "0"                    '
            .Col = 2:  .Text = RS01!접수일자 & ""     '
            .Col = 3:  .Text = Format(RS01!택번호, "000-00-0000") & ""      '
            .Col = 4:  .Text = RS01!의류코드 & ""     '
            .Col = 5:  .Text = RS01!의류명 & ""       '
            .Col = 6:  .Text = RS01!색상 & ""         '
            .Col = 7:  .Text = RS01!무늬 & ""         '
            .Col = 8:  .Text = RS01!내용 & ""         '
            .Col = 9:  .Text = RS01!금액 & ""         '
            .Col = 10: .Text = RS01!상표 & ""         '
            .Col = 11: .Text = RS01!고객명 & ""     '
            .Col = 12: .Text = RS01!전화번호 & ""     '
            .Col = 13: .Text = RS01!휴대폰번호 & ""     '
            .Col = 14: .Text = RS01!가맹점입고일자 & "" '
            .Col = 15: .Text = RS01!고객출고일자 & "" '
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub


Private Sub DataDelete()
    Dim 접수일자 As String
    
    Dim i As Integer
    ReDim sValue(5)
    
    ProgressBar.Value = 0
    ProgressBar.Min = 0
    ProgressBar.Max = 100
    ProgressBar.Visible = True
    
    For i = 1 To spdView.MaxRows
        ProgressBar.Value = (i / spdView.MaxRows) * 100
        DoEvents
        
        spdView.Row = i
        spdView.Col = 1
        If spdView.Text = "1" Then
            spdView.Col = 2: 접수일자 = spdView.Text
            
            If Weekday(DateAdd("d", 3, 접수일자)) = 0 Then
                sValue(0) = Format(DateAdd("d", 4, 접수일자), "YYYY-MM-DD") ' 0
            Else
                sValue(0) = Format(DateAdd("d", 3, 접수일자), "YYYY-MM-DD") ' 0
            End If
            
                              sValue(1) = 접수일자                          ' 1
                              sValue(2) = Mid(cboInput.Text, 2, 6)          ' 2
            spdView.Col = 3:  sValue(3) = Replace(spdView.Text, "-", "")    ' 3
            spdView.Col = 4:  sValue(4) = spdView.Text & ""                 ' 4
            spdView.Col = 12: sValue(5) = ""                  ' 5
            
            If HeadOffice = MASTER_OFFICE_CODE Then
                If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub
                Call ExecProMaster("SP_03007_01", sValue(), Err_Num, Err_Dec)
            Else
                Call ExecPro("SP_03007_01", sValue(), Err_Num, Err_Dec)
            End If
            
            If Err_Num <> 0 Then
                MsgBox "[" & Err_Num & "] " & Err_Dec
                Exit Sub
            End If
'
'            spdView.Col = -1
'            spdView.Action = ActionDeleteRow
'            spdView.MaxRows = spdView.MaxRows - 1
        End If
    Next i
    
    ProgressBar.Visible = False
    
    Call Data_Display
    
    MsgBox "미출고 정리가 완료되었습니다.", vbInformation
End Sub

Private Sub DataPrint()

End Sub

Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If NewRow <> -1 Then
        spdView.Row = Row
        spdView.Col = -1
        spdView.BackColor = vbWhite
        
        spdView.Row = NewRow
        spdView.Col = -1
        spdView.BackColor = glbYellow
    End If
End Sub

