VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_03018 
   Caption         =   "가맹점 고객 미출고 현황"
   ClientHeight    =   8805
   ClientLeft      =   990
   ClientTop       =   3465
   ClientWidth     =   16395
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_03018.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8805
   ScaleWidth      =   16395
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8805
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16395
      _ExtentX        =   28919
      _ExtentY        =   15531
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_03018.frx":058A
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   1
         Top             =   540
         Width           =   16365
         _ExtentX        =   28866
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1245
            Locked          =   -1  'True
            TabIndex        =   13
            Text            =   "cboOffice"
            Top             =   60
            Width           =   3420
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   6525
            TabIndex        =   14
            Top             =   60
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   68812800
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   5340
            TabIndex        =   15
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "접수일자"
            BorderWidth     =   0
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   7
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
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   9825
            TabIndex        =   17
            Top             =   60
            Width           =   3030
            _ExtentX        =   5345
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   68812800
            CurrentDate     =   36686
         End
         Begin VB.Label Label 
            Alignment       =   2  '가운데 맞춤
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
            Height          =   255
            Left            =   9510
            TabIndex        =   18
            Top             =   120
            Width           =   300
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   8760
         _ExtentX        =   15452
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
         Caption         =   "가맹점 고객 미출고 현황 (P_03018)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_03018.frx":063C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Left            =   8790
         TabIndex        =   3
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
         PictureBackground=   "P_03018.frx":083E
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   4
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
            Picture         =   "P_03018.frx":0A40
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   5
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
            Picture         =   "P_03018.frx":0FDA
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   6
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
            Picture         =   "P_03018.frx":1574
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   7
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
            Picture         =   "P_03018.frx":1B0E
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   8
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
            Picture         =   "P_03018.frx":20A8
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   9
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
            Picture         =   "P_03018.frx":2642
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   10
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
            Picture         =   "P_03018.frx":2BDC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   11
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
            Picture         =   "P_03018.frx":3176
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   7455
         Index           =   0
         Left            =   15
         TabIndex        =   12
         Top             =   1335
         Width           =   5700
         _Version        =   524288
         _ExtentX        =   10054
         _ExtentY        =   13150
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
         MaxCols         =   3
         ScrollBars      =   2
         SpreadDesigner  =   "P_03018.frx":3710
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   7455
         Index           =   1
         Left            =   5730
         TabIndex        =   19
         Top             =   1335
         Width           =   10650
         _Version        =   524288
         _ExtentX        =   18785
         _ExtentY        =   13150
         _StockProps     =   64
         BackColorStyle  =   1
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
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
         MaxCols         =   23
         Protect         =   0   'False
         SpreadDesigner  =   "P_03018.frx":3CE0
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_03018"
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
    If cboOffice.ListIndex < 0 Then Exit Sub
    
    '-----------------------------------------------------------------
    '
    '-----------------------------------------------------------------
    Screen.MousePointer = vbHourglass
    
    ReDim sValue(2)
    
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    If IsNull(dtInput(0).Value) Then
        sValue(1) = "2001-01-01"
        sValue(2) = Format(Date, "YYYY-MM-DD")
    Else
        sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
        sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    End If
    
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_03018_01", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_03018_01", sValue(), Err_Num, Err_Dec)
    End If
    
 
    With spdView(0)
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = RS01!가맹점코드 & ""
            .Col = 2: .Text = RS01!가맹점명 & ""
            .Col = 3: .Text = RS01!미출고수량 & ""
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
        
    Call SpreadSum(spdView(0), 2, 3)
    Screen.MousePointer = vbDefault


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
        Case 0: Call Data_Display   ' 조회
        Case 1: 'Call DataAdd        ' 신규
        Case 2: 'Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
        Case 4: 'Call DataCancel     ' 취소
        Case 5: 'Call DataPrint      ' 인쇄
        Case 6
            Call Export_Excel(P_00000.cdgExcel, spdView(1))              ' 엑셀
                
                
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

 
'Private Sub dtInput_Change(Index As Integer)
'    Call Data_Display
'End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    'cmdBtn(5).Enabled = True
    cmdBtn(6).Enabled = True
    
    '본사 여부
    If HeadOffice = MASTER_OFFICE_CODE Then
        cboOffice.Locked = False
    Else
        cboOffice.Locked = True
    End If
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"

    If P_02005_Flag = False Then
'        Call GoodsComboAdd(cboInput(0))
'        Call GoodsComboAdd(cboInput(1))
'        Call AgencyComboAdd(cboInput(2))
'
'        dtInput(0).Value = Date
'        dtInput(1).Value = Date
'
'        ReDim sValue(6)
'
'        sValue(0) = "1"
'
'        Set RS01 = New ADODB.Recordset
'        Set RS01 = ExecPro("SP_02005_00", sValue(), Err_Num, Err_Dec)
'
'        spdView.MaxCols = RS01.Fields.Count
'        spdView.MaxRows = RS01.RecordCount
'
'        Call spdDisplay(RS01)
'        Call GetColWidth(REG_App, Me.Name, spdView)
        
        P_02005_Flag = True
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    With spdView(0)
        .MaxRows = 0
        .RowHeight(-1) = 14
        
        .Col = 2: .ColMerge = MergeAlways
        .Col = 3: .ColMerge = MergeRestricted
        .Col = 4: .ColMerge = MergeRestricted
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        '.OperationMode = OperationModeSingle
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With

    With spdView(1)
        .MaxRows = 0
        .RowHeight(-1) = 14
        
        .Col = 2: .ColMerge = MergeAlways
        .Col = 3: .ColMerge = MergeRestricted
        .Col = 4: .ColMerge = MergeRestricted
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
'        .OperationMode = OperationModeSingle
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
    
    If P_02005_Flag = False Then
        dtInput(0).Value = Date
        dtInput(1).Value = Date

        '
        Call Get_지사리스트(cboOffice, False)
        
        Dim i As Integer
        
        cboOffice.AddItem "[0000] 전체", 0
        With cboOffice
            For i = 0 To .ListCount - 1
                If Mid(.List(i), 2, 4) = HeadOffice Then
                    .ListIndex = i
                    
                    Exit For
                End If
            Next i
        End With

        P_02005_Flag = True
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_02005_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    
    '-------------------------------------------------------------
    ' SP_02005_00
    '-------------------------------------------------------------
    ReDim sValue(2)
    
    Screen.MousePointer = vbHourglass
        
            
    sValue(0) = spdView(0).Tag
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")

    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_03018_02", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_03018_02", sValue(), Err_Num, Err_Dec)
    End If
    
     With spdView(1)
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = Format(RS01!택번호, "000-00-0000") & ""        ' 1
            .Col = 2:  .Text = RS01!의류코드 & ""       ' 6
            .Col = 3:  .Text = RS01!의류명 & ""         ' 7
            .Col = 4:  .Text = RS01!색상 & ""           ' 8
            .Col = 5:  .Text = RS01!내용 & ""           ' 9
            .Col = 6:  .Text = RS01!상표 & ""           '10
            .Col = 7:  .Text = RS01!금액 & ""           '11
            
            .Col = 8:  .Text = RS01!접수번호 & ""       ' 2
            .Col = 9:  .Text = RS01!성명 & ""           ' 3
            .Col = 10: .Text = RS01!전화번호 & ""       ' 4
            .Col = 11: .Text = RS01!휴대전화 & ""         ' 5
            
            .Col = 12: .Text = RS01!접수일자 & ""       '12
            .Col = 13: .Text = RS01!접수시간 & ""       '12
            
            .Col = 14: .Text = RS01!가맹점출고일자 & "" '13
            .Col = 15: .Text = RS01!가맹점입고일자 & "" '14
            .Col = 16: .Text = RS01!지사입고일자 & ""   '15
            .Col = 17: .Text = RS01!지사출고일자 & ""   '16
            
            .Col = 18: .Text = RS01!부모택번호 & ""     '18
            .Col = 19: .Text = RS01!반품환불일자 & ""   '19
            .Col = 20: .Text = RS01!세탁환불일자 & ""   '20
            .Col = 21: .Text = RS01!판매취소일자 & ""   '21
            .Col = 22: .Text = RS01!환불사유 & ""       '22
            .Col = 23: .Text = RS01!오점내용 & ""       '23


            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
        
    Screen.MousePointer = vbDefault
        
    Exit Sub
    
ErrRtn:
    Screen.MousePointer = vbDefault
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub
 
Private Sub DataPrint()
'    Dim ReportFP As String
'    Dim ReportFile As String
'
'    ReportFP = GetIniStr("REPORT", "FilePath", "", m_iniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    Call PrintDesc
'
'    Dim i As Integer
'    For i = 0 To 30
'        P_00000.crPrint.Formulas(i) = ""
'    Next
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'    P_00000.crPrint.Formulas(0) = "입고일자1 = '" & Format(dtInput(0).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(1) = "입고일자2 = '" & Format(dtInput(1).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(2) = "순위구분 = '" & IIf(optSelect(0).Value = True, "금액", "수량") & "'"
'    P_00000.crPrint.Formulas(3) = "대리점명 = '" & cboInput(2).Text & "'"
'    P_00000.crPrint.Formulas(4) = "품목명1 = '" & cboInput(0).Text & "'"
'    P_00000.crPrint.Formulas(5) = "품목명2 = '" & cboInput(1).Text & "'"
'
'    P_00000.crPrint.Formulas(6) = "수량합계 = '" & txtNum(0).Text & "'"
'    P_00000.crPrint.Formulas(7) = "금액합계 = '" & txtNum(1).Text & "'"
'    P_00000.crPrint.Formulas(8) = "점유율(단위)수량 = '" & txtNum(2).Text & "'"
'    P_00000.crPrint.Formulas(9) = "점유율(단위)금액 = '" & txtNum(3).Text & "'"
'    P_00000.crPrint.Formulas(10) = "점유율(전체)수량 = '" & txtNum(4).Text & "'"
'    P_00000.crPrint.Formulas(11) = "점유율(전체)금액 = '" & txtNum(5).Text & "'"
'
'    Call ReportPrint(ReportFile, "1")
End Sub

 
Private Sub DataScreen()
'    Dim ReportFP As String
'    Dim ReportFile As String
'
'    ReportFP = GetIniStr("REPORT", "FilePath", "", m_iniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    Call PrintDesc
'
'    Dim i As Integer
'    For i = 0 To 30
'        P_00000.crPrint.Formulas(i) = ""
'    Next
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'    P_00000.crPrint.Formulas(0) = "입고일자1 = '" & Format(dtInput(0).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(1) = "입고일자2 = '" & Format(dtInput(1).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(2) = "순위구분 = '" & IIf(optSelect(0).Value = True, "금액", "수량") & "'"
'    P_00000.crPrint.Formulas(3) = "대리점명 = '" & cboInput(2).Text & "'"
'    P_00000.crPrint.Formulas(4) = "품목명1 = '" & cboInput(0).Text & "'"
'    P_00000.crPrint.Formulas(5) = "품목명2 = '" & cboInput(1).Text & "'"
'
'    P_00000.crPrint.Formulas(6) = "수량합계 = '" & txtNum(0).Text & "'"
'    P_00000.crPrint.Formulas(7) = "금액합계 = '" & txtNum(1).Text & "'"
'    P_00000.crPrint.Formulas(8) = "점유율(단위)수량 = '" & txtNum(2).Text & "'"
'    P_00000.crPrint.Formulas(9) = "점유율(단위)금액 = '" & txtNum(3).Text & "'"
'    P_00000.crPrint.Formulas(10) = "점유율(전체)수량 = '" & txtNum(4).Text & "'"
'    P_00000.crPrint.Formulas(11) = "점유율(전체)금액 = '" & txtNum(5).Text & "'"
'
'    Call ReportPrint(ReportFile, "2")
End Sub

Private Sub PrintDesc()
'    Dim i As Integer
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
'    For i = 1 To spdView.MaxRows - 1
'        spdView.Row = i
'
'        TempText = Left(i & Space(3), 3)
'
'        spdView.Col = 1
'        TempText = TempText & LeftH(Mid(spdView.Text, 7) & Space(12), 12)
'        spdView.Col = 2
'        TempText = TempText & RightH(Space(6) & spdView.Text, 6) & Space(1)
'        spdView.Col = 3
'        TempText = TempText & RightH(Space(10) & spdView.Text, 10) & Space(4)
'        spdView.Col = 4
'        TempText = TempText & RightH(Space(10) & spdView.Text, 10) & Space(5)
'        spdView.Col = 5
'        TempText = TempText & RightH(Space(10) & spdView.Text, 10) & Space(5)
'        spdView.Col = 6
'        TempText = TempText & RightH(Space(10) & spdView.Text, 10) & Space(5)
'        spdView.Col = 7
'        TempText = TempText & RightH(Space(10) & spdView.Text, 10) & Space(5)
'        spdView.Col = 8
'        TempText = TempText & RightH(Space(10) & spdView.Text, 10)
'
'        Print #1, TempText
'        TempText = ""
'    Next i
'
'    Close #1
End Sub

 

Private Sub spdView_DblClick(Index As Integer, ByVal Col As Long, ByVal Row As Long)
    If Index = 0 Then
        
        
        Dim vText       As Variant
        
        spdView(Index).GetText 1, Row, vText
        
        If CStr(vText) <> "" Then
            spdView(Index).Tag = CStr(vText)
            Call Data_Display
        End If
    End If
End Sub
