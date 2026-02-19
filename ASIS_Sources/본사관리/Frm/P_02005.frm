VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_02005 
   Caption         =   "품목별 입고현황"
   ClientHeight    =   8805
   ClientLeft      =   6450
   ClientTop       =   4770
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
   Icon            =   "P_02005.frx":0000
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
      PaneTree        =   "P_02005.frx":058A
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
         Begin VB.ComboBox cboInput 
            Height          =   315
            Left            =   1245
            Style           =   2  '드롭다운 목록
            TabIndex        =   15
            Top             =   405
            Width           =   3420
         End
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1245
            Locked          =   -1  'True
            Style           =   2  '드롭다운 목록
            TabIndex        =   14
            Top             =   60
            Width           =   3420
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   6525
            TabIndex        =   16
            Top             =   60
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            Format          =   56557568
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   17
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
            Index           =   2
            Left            =   5340
            TabIndex        =   18
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
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   9705
            TabIndex        =   20
            Top             =   60
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            Format          =   56557568
            CurrentDate     =   36686
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
            Picture         =   "P_02005.frx":063C
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
            Left            =   9390
            TabIndex        =   21
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
         Caption         =   " 품목별 입고현황 (P_02005)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_02005.frx":0BD6
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
         PictureBackground=   "P_02005.frx":0DD8
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
            Picture         =   "P_02005.frx":0FDA
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
            Picture         =   "P_02005.frx":1574
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
            Picture         =   "P_02005.frx":1B0E
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
            Picture         =   "P_02005.frx":20A8
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
            Picture         =   "P_02005.frx":2642
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
            Picture         =   "P_02005.frx":2BDC
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
            Picture         =   "P_02005.frx":3176
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
            Picture         =   "P_02005.frx":3710
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   7455
         Left            =   15
         TabIndex        =   12
         Top             =   1335
         Width           =   7245
         _Version        =   524288
         _ExtentX        =   12779
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
         MaxCols         =   6
         ScrollBars      =   2
         SpreadDesigner  =   "P_02005.frx":3CAA
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdView1 
         Height          =   7455
         Left            =   7275
         TabIndex        =   13
         Top             =   1335
         Width           =   9105
         _Version        =   524288
         _ExtentX        =   16060
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
         MaxCols         =   6
         SpreadDesigner  =   "P_02005.frx":440B
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_02005"
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
    
    If Mid(cboOffice.Text, 2, 4) = "0000" Then
        cboInput.Enabled = False
        Data_Display
        Exit Sub
        sValue(0) = ""
    Else
        cboInput.Enabled = True
        sValue(0) = Mid(cboOffice.Text, 2, 4)
    End If
    
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01001_00", sValue(), Err_Num, Err_Dec)
    
    With cboInput
        .AddItem "[000000] 전체": .ItemData(cboInput.NewIndex) = "0000"
        
        Do Until RS01.EOF
            'If (RS01!종료일자 = "") Or (RS01!종료일자 = "2099-12-31") Then '가맹점이 현 지사에서 관리중...
                .AddItem "[" & RS01!가맹점코드 & "] " & RS01!가맹점명: .ItemData(cboInput.NewIndex) = RS01!지사코드
            'End If
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        If .ListCount > 0 Then .ListIndex = 0
    End With
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
            If cmdBtn(6).Tag = "0" Then
                Call Export_Excel(P_00000.cdgExcel, spdView)              ' 엑셀
            Else
                Call Export_Excel(P_00000.cdgExcel, spdView1)              ' 엑셀
            
            End If
                
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

'Private Sub dtInput_Change(Index As Integer)
'    Call Data_Display
'End Sub

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
    
    With spdView
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

    With spdView1
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

    Dim i         As Long
    
    Dim 수량합계  As Double
    Dim 금액합계  As Double
    
    Dim 접수수량  As Double
    Dim 접수금액  As Double
    
    Dim 지사코드  As String
    
    '-------------------------------------------------------------
    ' SP_02005_00
    '-------------------------------------------------------------
    ReDim sValue(3)
    
    Screen.MousePointer = vbHourglass
        
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    sValue(1) = Mid(cboInput.Text, 2, 6)
    sValue(2) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(3) = Format(dtInput(1).Value, "YYYY-MM-DD")

'    If HeadOffice = MASTER_OFFICE_CODE Then
'        'If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub
'
'        '지사코드 = Format(cboInput.ItemData(cboInput.ListIndex), "0000")
'        지사코드 = Mid(cboOffice.Text, 2, 4)
'
'        If 지사코드 = "0000" Then
'            If DBOpen_Master("1000") = False Then Exit Sub
'
'            Set RS01 = New ADODB.Recordset
'            Set RS01 = ExecProMaster("SP_02005_03", sValue(), Err_Num, Err_Dec)
'
'        Else
'
'            If DBOpen_Master(지사코드) = False Then Exit Sub
'
'            Set RS01 = New ADODB.Recordset
'            Set RS01 = ExecProMaster("SP_02005_00", sValue(), Err_Num, Err_Dec)
'        End If
'    Else
'        Set RS01 = New ADODB.Recordset
'        Set RS01 = ExecPro("SP_02005_00", sValue(), Err_Num, Err_Dec)
'    End If
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_02005_03", sValue(), Err_Num, Err_Dec)

    수량합계 = 0
    금액합계 = 0
    
    With spdView
        .MaxRows = 0
        .Redraw = False
        
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        
        .Col = 1: .Text = ""
        .Col = 2: .Text = "-전체-"
        .Col = 3: .Text = 0
        .Col = 4: .Text = 0
        
        .Row = .Row
        .Row2 = .Row
        .Col = 1
        .Col2 = .MaxCols
        .BlockMode = True
        .BackColor = &H80FFFF
        .BlockMode = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = RS01!의류분류코드 & ""
            .Col = 2: .Text = RS01!의류분류명 & ""
            .Col = 3: .Text = RS01!접수수량 & ""
            .Col = 4: .Text = RS01!접수금액 & ""
            
            수량합계 = 수량합계 + RS01!접수수량
            금액합계 = 금액합계 + RS01!접수금액
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Row = 1
        .Col = 3: .Text = 수량합계 & ""
        .Col = 4: .Text = 금액합계 & ""
        
        '-------------------------------------------
        '
        '-------------------------------------------
        For i = 2 To .MaxRows
            .Row = i
            .Col = 3: 접수수량 = .Value
            .Col = 4: 접수금액 = .Value
        
            If 수량합계 = 0 Then
                .Col = 5: .Text = 0
            Else
                .Col = 5: .Text = (접수수량 / 수량합계) * 100
            End If
            
            If 금액합계 = 0 Then
                .Col = 6: .Text = 0
            Else
                .Col = 6: .Text = (접수금액 / 금액합계) * 100
            End If
        Next i
        
        .Redraw = True
    End With
        
    Screen.MousePointer = vbDefault
        
    Exit Sub
    
ErrRtn:
    Screen.MousePointer = vbDefault
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Data_Display2(의류분류코드 As String, 수량합계 As Double, 금액합계 As Double)
    Dim i         As Long
        
    Dim 지사코드  As String
    
    '-------------------------------------------------------------
    ' SP_02005_00
    '-------------------------------------------------------------
    ReDim sValue(3)

    sValue(0) = Mid(cboInput.Text, 2, 6)
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    sValue(3) = 의류분류코드

    If HeadOffice = MASTER_OFFICE_CODE Then
        'If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub

        
        지사코드 = Mid(cboOffice.Text, 2, 4)
        If 지사코드 = "0000" Then
            ReDim sValue(4)
        
            sValue(0) = 지사코드
            sValue(1) = Mid(cboInput.Text, 2, 6)
            sValue(2) = Format(dtInput(0).Value, "YYYY-MM-DD")
            sValue(3) = Format(dtInput(1).Value, "YYYY-MM-DD")
            sValue(4) = 의류분류코드
                    
            If DBOpen_Master("1000") = False Then Exit Sub
    
            Set RS01 = New ADODB.Recordset
            Set RS01 = ExecProMaster("SP_02005_04_NEW", sValue(), Err_Num, Err_Dec)
        
        Else
        
            If DBOpen_Master(지사코드) = False Then Exit Sub
    
            Set RS01 = New ADODB.Recordset
            Set RS01 = ExecProMaster("SP_02005_01", sValue(), Err_Num, Err_Dec)
        End If
    
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_02005_01", sValue(), Err_Num, Err_Dec)
    End If
    
    With spdView1
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = RS01!의류코드 & ""
            .Col = 2: .Text = RS01!의류명 & ""
            .Col = 3: .Text = RS01!접수수량 & ""
            .Col = 4: .Text = RS01!접수금액 & ""
            
            If 수량합계 = 0 Then
                .Col = 5: .Text = 0
            Else
                .Col = 5: .Text = (RS01!접수수량 / 수량합계) * 100
            End If
            
            If 금액합계 = 0 Then
                .Col = 6: .Text = 0
            Else
                .Col = 6: .Text = (RS01!접수금액 / 금액합계) * 100
            End If
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
    
    Call SpreadSum(spdView1, 2, 3)
    Call SpreadSum(spdView1, -1, 4)
    Call SpreadSum(spdView1, -1, 5)
    Call SpreadSum(spdView1, -1, 6)
    
    Debug.Print "Data_Display2" & "  " & Now
    
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

Private Sub spdView_Click(ByVal Col As Long, ByVal Row As Long)
    Dim 의류분류코드 As String
    
    Dim 수량합계  As Double
    Dim 금액합계  As Double
    
    cmdBtn(6).Tag = "0"
    
    If Row <= 0 Then Exit Sub
    
    spdView.Row = Row
    spdView.Col = 1: 의류분류코드 = spdView.Text & ""
    spdView.Col = 3: 수량합계 = spdView.Value & ""
    spdView.Col = 4: 금액합계 = spdView.Value & ""
    
'    spdView.Enabled = False
    Call Data_Display2(의류분류코드, 수량합계, 금액합계)
'    spdView.Enabled = True
End Sub

Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    'Call spdView_Click(NewCol, NewRow)
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
    Dim i As Integer
    Dim TempText As String
    Dim TempFP As String
    Dim TempFile As String
    
    TempFP = GetIniStr("REPORT", "TempPath", "", m_iniFile)
    TempFile = TempFP & "\Temp.txt"
    
    Open TempFile For Output As #1
    
    TempText = ""
    
    For i = 1 To spdView.MaxRows - 1
        spdView.Row = i
        
        TempText = Left(i & Space(3), 3)
        
        spdView.Col = 1
        TempText = TempText & LeftH(Mid(spdView.Text, 7) & Space(12), 12)
        spdView.Col = 2
        TempText = TempText & RightH(Space(6) & spdView.Text, 6) & Space(1)
        spdView.Col = 3
        TempText = TempText & RightH(Space(10) & spdView.Text, 10) & Space(4)
        spdView.Col = 4
        TempText = TempText & RightH(Space(10) & spdView.Text, 10) & Space(5)
        spdView.Col = 5
        TempText = TempText & RightH(Space(10) & spdView.Text, 10) & Space(5)
        spdView.Col = 6
        TempText = TempText & RightH(Space(10) & spdView.Text, 10) & Space(5)
        spdView.Col = 7
        TempText = TempText & RightH(Space(10) & spdView.Text, 10) & Space(5)
        spdView.Col = 8
        TempText = TempText & RightH(Space(10) & spdView.Text, 10)
        
        Print #1, TempText
        TempText = ""
    Next i
    
    Close #1
End Sub

Private Sub spdView1_Click(ByVal Col As Long, ByVal Row As Long)
    cmdBtn(6).Tag = "1"

End Sub
