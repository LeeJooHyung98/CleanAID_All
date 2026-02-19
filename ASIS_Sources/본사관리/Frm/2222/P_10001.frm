VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_10001 
   Caption         =   " 고객 조회"
   ClientHeight    =   12195
   ClientLeft      =   1725
   ClientTop       =   2520
   ClientWidth     =   16890
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12195
   ScaleWidth      =   16890
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   12195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16890
      _ExtentX        =   29792
      _ExtentY        =   21511
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_10001.frx":0000
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
         Begin VB.ComboBox cboGubun 
            Height          =   315
            Left            =   10665
            Style           =   2  '드롭다운 목록
            TabIndex        =   22
            Top             =   60
            Width           =   2700
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   1
            Left            =   6135
            TabIndex        =   19
            Top             =   420
            Width           =   2895
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   0
            Left            =   6135
            TabIndex        =   18
            Top             =   60
            Width           =   2895
         End
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1605
            Style           =   2  '드롭다운 목록
            TabIndex        =   16
            Top             =   60
            Width           =   2850
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Left            =   1605
            TabIndex        =   2
            Top             =   420
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            Format          =   68747264
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   3
            Top             =   420
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "최종이용일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   17
            Top             =   60
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "지    사"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   9120
            TabIndex        =   23
            Top             =   60
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "고객등급"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin VB.Label lblTitle 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "전화번호:"
            Height          =   225
            Index           =   3
            Left            =   4620
            TabIndex        =   21
            Top             =   465
            Width           =   1470
         End
         Begin VB.Label lblTitle 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "성    명:"
            Height          =   225
            Index           =   2
            Left            =   4620
            TabIndex        =   20
            Top             =   120
            Width           =   1470
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
         Caption         =   " 고객 조회 (P_10001)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_10001.frx":00B2
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
         PictureBackground=   "P_10001.frx":02B4
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
            Picture         =   "P_10001.frx":04B6
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
            Caption         =   "화면"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_10001.frx":0A50
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
            Picture         =   "P_10001.frx":0FEA
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
            Picture         =   "P_10001.frx":1584
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
            Picture         =   "P_10001.frx":1B1E
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
            Picture         =   "P_10001.frx":20B8
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
            Picture         =   "P_10001.frx":2652
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
            Picture         =   "P_10001.frx":2BEC
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10845
         Left            =   15
         TabIndex        =   14
         Top             =   1335
         Width           =   4410
         _Version        =   524288
         _ExtentX        =   7779
         _ExtentY        =   19129
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
         MaxRows         =   35
         ScrollBars      =   2
         SpreadDesigner  =   "P_10001.frx":3186
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdView1 
         Height          =   10845
         Left            =   4440
         TabIndex        =   15
         Top             =   1335
         Width           =   12435
         _Version        =   524288
         _ExtentX        =   21934
         _ExtentY        =   19129
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
         SpreadDesigner  =   "P_10001.frx":3702
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_10001"
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

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: Call Data_Display   ' 조회
        Case 1: ' Call DataAdd        ' 신규
        Case 2: 'Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
        Case 4: 'Call DataCancel     ' 취소
        Case 5: 'Call DataPrint      ' 인쇄
        Case 6: 'Call DataScreen     '
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
    
    Call Data_Display
    
    dtInput.Enabled = True
    dtInput.SetFocus
End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(5).Enabled = True
    cmdBtn(6).Enabled = True
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"

    If P_10001_Flag = False Then
        dtInput.Value = Date
        
        P_10001_Flag = True
    End If
End Sub

Private Sub Form_Load()
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
        .OperationMode = OperationModeSingle
        
        'Init the User Sort
        .UserColAction = UserColActionSort
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
    
    Call MasterComboAdd(cboOffice)
    Call MemberGubunAdd(cboGubun)
    
    Dim I As Integer
    
 
    With cboOffice
        For I = 0 To .ListCount - 1
            If Mid(.List(I), 2, 4) = HeadOffice Then
                .ListIndex = I
                
                Exit For
            End If
        Next I
    End With
    
     cboOffice.Enabled = IIf(Store.Code = MASTER_CODE, True, False)
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_10001_Flag = False
End Sub

Private Sub Data_Display()
    Dim lAmt As Long
    
    On Error GoTo ErrRtn
    
    '-------------------------------------------------------------
    ' SP_02002_00
    '-------------------------------------------------------------
    ReDim sValue(5)
    
    sValue(0) = "0"
    sValue(1) = Trim(Mid(cboOffice.Text, 2, 4)) + "%"
    sValue(2) = Trim(txtInput(0).Text) + "%"
    sValue(3) = Trim(txtInput(1).Text) + "%"
    sValue(4) = Trim(Mid(cboGubun.Text, 2, 1)) + "%"
    sValue(5) = Format(dtInput.Value, "YYYY-MM-DD")
        
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_10001_00", sValue(), Err_Num, Err_Dec)
    
    With spdView
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = Trim(RS01!코드) & ""
            .Col = 2: .Text = Trim(RS01!가맹점명) & ""
            .Col = 3: .Text = RS01!고객수 & ""
            
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

Public Sub DataPrint()
    Dim ReportFP As String
    Dim ReportFile As String
    
    ReportFP = GetIniStr("REPORT", "FilePath", "", sIniFile)
    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
    
    Call PrintDesc
    
    Dim I As Integer
    
    For I = 0 To 30
        P_00000.crPrint.Formulas(I) = ""
    Next
    
    P_00000.crPrint.WindowTitle = Me.Caption
    P_00000.crPrint.Formulas(0) = "입고일자 = '" & Format(dtInput.Value, "YYYY-MM-DD") & "'"
    
    Call ReportPrint(ReportFile, "1")
End Sub

Public Sub DataScreen()
    Dim ReportFP As String
    Dim ReportFile As String
    
    ReportFP = GetIniStr("REPORT", "FilePath", "", sIniFile)
    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
    
    Call PrintDesc
    
    Dim I As Integer
    
    For I = 0 To 30
        P_00000.crPrint.Formulas(I) = ""
    Next
    
    P_00000.crPrint.WindowTitle = Me.Caption
    P_00000.crPrint.Formulas(0) = "입고일자 = '" & Format(dtInput.Value, "YYYY-MM-DD") & "'"
    
    Call ReportPrint(ReportFile, "2")
End Sub

Private Sub PrintDesc()
'    Dim i, j As Integer
'
'    Dim TempText As String
'    Dim TempFP As String
'    Dim TempFile As String
'
'    TempFP = GetIniStr("REPORT", "TempPath", "", sIniFile)
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
    spdView.MaxRows = spdView.MaxRows + 1
    
    spdView.Row = spdView.MaxRows
    spdView.Col = 1
    spdView.Action = ActionActiveCell
    
    spdView.SetFocus
End Sub

Private Sub spdView_Click(ByVal Col As Long, ByVal Row As Long)
    Dim 가맹점코드 As String
    
    On Error GoTo ErrRtn
    
    If Row <= 0 Then Exit Sub
    
    spdView.Row = Row
    spdView.Col = 1: 가맹점코드 = spdView.Text & ""
    
    '-------------------------------------------------------------
    ' SP_10001_01
    '-------------------------------------------------------------
    ReDim sValue(5)
    
    sValue(0) = "0"
    sValue(1) = 가맹점코드
    sValue(2) = Trim(txtInput(0).Text) + "%"
    sValue(3) = "%" + Trim(txtInput(1).Text)
    sValue(4) = Trim(Mid(cboGubun.Text, 2, 1)) + "%"
    sValue(5) = Format(dtInput.Value, "YYYY-MM-DD")
        
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_10001_01", sValue(), Err_Num, Err_Dec)
    
    With spdView1
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows


            
            .Col = 1: .Text = RS01!고객코드 & ""
            .Col = 2: .Text = Trim(RS01!고객등급코드) & ""
            .Col = 3: .Text = Trim(RS01!성명) & ""
            .Col = 4: .Text = Trim(RS01!전화번호) & ""
            .Col = 5: .Text = RS01!휴대폰번호 & ""
            .Col = 6: .Text = RS01!최종이용일자 & ""
            .Col = 7: .Text = RS01!이용횟수 & ""
            .Col = 8: .Text = RS01!총접수금액 & ""
            .Col = 9: .Text = RS01!누적마일리지 & ""
            
            .Col = 10: .Text = RS01!사용가능마일리지 & ""
            .Col = 11: .Text = RS01!미수금 & ""
            .Col = 12: .Text = RS01!SMS & ""
            .Col = 13: .Text = RS01!등록일자 & ""
            .Col = 14: .Text = RS01!상태 & ""
            .Col = 15: .Text = RS01!주소 & ""
            
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

Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    'Call spdView_Click(NewCol, NewRow)
End Sub
