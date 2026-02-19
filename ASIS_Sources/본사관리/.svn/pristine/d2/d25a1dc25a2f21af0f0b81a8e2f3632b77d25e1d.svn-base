VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_04012 
   Caption         =   "점별 월별 매출현황"
   ClientHeight    =   9720
   ClientLeft      =   735
   ClientTop       =   2565
   ClientWidth     =   16005
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_04012.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9720
   ScaleWidth      =   16005
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   9720
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16005
      _ExtentX        =   28231
      _ExtentY        =   17145
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_04012.frx":058A
      Begin Threed.SSPanel SSPanel 
         Height          =   765
         Index           =   0
         Left            =   15
         TabIndex        =   2
         Top             =   8940
         Width           =   15975
         _ExtentX        =   28178
         _ExtentY        =   1349
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   5
            Left            =   11985
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   405
            Width           =   735
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   4
            Left            =   9585
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   405
            Width           =   735
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   3
            Left            =   7185
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   405
            Width           =   735
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   2
            Left            =   4785
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   405
            Width           =   735
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   1
            Left            =   1485
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   405
            Width           =   1635
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   0
            Left            =   1485
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   45
            Width           =   1635
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   6
            Left            =   14385
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   405
            Width           =   735
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   3345
            TabIndex        =   10
            Top             =   405
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "단가금액평균"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   3
            Left            =   45
            TabIndex        =   11
            Top             =   405
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "입고수량합계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   4
            Left            =   5745
            TabIndex        =   12
            Top             =   405
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "출고수량합계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   5
            Left            =   8145
            TabIndex        =   13
            Top             =   405
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "반 품 합 계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   6
            Left            =   10545
            TabIndex        =   14
            Top             =   405
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "재세탁합계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   7
            Left            =   12945
            TabIndex        =   15
            Top             =   405
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "수 선 합 계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   8
            Left            =   45
            TabIndex        =   16
            Top             =   45
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "수금액합계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   7590
         Left            =   15
         TabIndex        =   1
         Top             =   1335
         Width           =   15975
         _Version        =   524288
         _ExtentX        =   28178
         _ExtentY        =   13388
         _StockProps     =   64
         BackColorStyle  =   1
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         SpreadDesigner  =   "P_04012.frx":063C
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   17
         Top             =   540
         Width           =   15975
         _ExtentX        =   28178
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Left            =   4260
            Style           =   2  '드롭다운 목록
            TabIndex        =   18
            Top             =   60
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Left            =   1680
            TabIndex        =   19
            Top             =   60
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy"
            Format          =   64159747
            UpDown          =   -1  'True
            CurrentDate     =   37140
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   20
            Top             =   60
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "수 금 년"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   2640
            TabIndex        =   21
            Top             =   60
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "대 리 점 명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   22
         Top             =   15
         Width           =   8370
         _ExtentX        =   14764
         _ExtentY        =   900
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   255
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
         Caption         =   " #"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_04012.frx":0AB5
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   1
         Left            =   8400
         TabIndex        =   23
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
         PictureBackground=   "P_04012.frx":0CB7
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   24
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
            Picture         =   "P_04012.frx":0EB9
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   25
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
            Picture         =   "P_04012.frx":1453
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   26
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
            Picture         =   "P_04012.frx":19ED
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   27
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
            Picture         =   "P_04012.frx":1F87
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   28
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
            Picture         =   "P_04012.frx":2521
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   29
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
            Picture         =   "P_04012.frx":2ABB
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   30
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
            Picture         =   "P_04012.frx":3055
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   31
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
            Picture         =   "P_04012.frx":35EF
         End
      End
   End
End
Attribute VB_Name = "P_04012"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: Call Data_Display   ' 조회
        Case 1: 'Call DataAdd        ' 신규
        Case 2: 'Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
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

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(5).Enabled = True
    cmdBtn(6).Enabled = True
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    If P_04012_Flag = False Then
        Call AgencyComboAdd(cboInput)
        
        dtInput.Value = Date
        
        ReDim sValue(2)
        
        sValue(0) = "1"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_04012_00", sValue(), Err_Num, Err_Dec)
        
        spdView.MaxCols = RS01.Fields.Count
        spdView.MaxRows = RS01.RecordCount
        
        Call spdDisplay(RS01)
        Call GetColWidth(REG_App, Me.Name, spdView)
        
        P_04012_Flag = True
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub spdDisplay(Rs As ADODB.Recordset)
    
    Call fpSpread_Display(spdView, Rs)
    
    spdView.ColsFrozen = 1 '틀고정
    
    spdView.Row = -1
    
    spdView.Col = 1
    spdView.ColWidth(1) = 10
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter

    spdView.Col = 2
    spdView.ColWidth(2) = 12
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
    
    spdView.Col = 3
    spdView.ColWidth(3) = 12
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 4
    spdView.ColWidth(4) = 12
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 5
    spdView.ColWidth(5) = 10
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 6
    spdView.ColWidth(6) = 10
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 7
    spdView.ColWidth(7) = 10
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 8
    spdView.ColWidth(8) = 10
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_04012_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    ReDim sValue(2)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput.Value, "yyyy")
    sValue(2) = Mid(cboInput.Text, 2, 6)
        
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_04012_00", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    Call spdDisplay(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView)
    
    spdView.AutoCalc = True
    spdView.MaxRows = spdView.MaxRows + 1
    
    spdView.Row = spdView.MaxRows
    spdView.RowHidden = True
    
    spdView.Col = 2
    spdView.Formula = "SUM(B1:B" & spdView.MaxRows - 1 & ")"
    txtInput(0).Text = spdView.Text
    spdView.Col = 3
    spdView.Formula = "SUM(C1:C" & spdView.MaxRows - 1 & ")"
    txtInput(1).Text = spdView.Text
    spdView.Col = 4
    spdView.Formula = "SUM(D1:D" & spdView.MaxRows - 1 & ") / " & spdView.MaxRows - 1
    txtInput(2).Text = spdView.Text
    spdView.Col = 5
    spdView.Formula = "SUM(E1:E" & spdView.MaxRows - 1 & ")"
    txtInput(3).Text = spdView.Text
    spdView.Col = 6
    spdView.Formula = "SUM(F1:F" & spdView.MaxRows - 1 & ")"
    txtInput(4).Text = spdView.Text
    spdView.Col = 7
    spdView.Formula = "SUM(G1:G" & spdView.MaxRows - 1 & ")"
    txtInput(5).Text = spdView.Text
    spdView.Col = 8
    spdView.Formula = "SUM(H1:H" & spdView.MaxRows - 1 & ")"
    txtInput(6).Text = spdView.Text
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
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

Private Sub spdView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        'PopupMenu P_00000.PopUp
    End If
End Sub

Public Sub DataPrint()
'    Dim ReportFP As String
'    Dim ReportFile As String
'    Dim sData As String
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
'    P_00000.crPrint.Formulas(0) = "년도 = '" & Format(dtInput.Value, "yyyy") & "'"
'    P_00000.crPrint.Formulas(1) = "대리점 = '" & Trim(Mid(cboInput.Text, 7)) & "'"
'
'    sData = Space(7) & RightH(Space(11) & txtInput(0).Text, 11) & Space(1)
'    sData = sData & RightH(Space(11) & txtInput(1).Text, 11) & Space(1)
'    sData = sData & RightH(Space(11) & txtInput(2).Text, 11) & Space(1)
'    sData = sData & RightH(Space(11) & txtInput(3).Text, 11) & Space(2)
'    sData = sData & RightH(Space(11) & txtInput(4).Text, 11) & Space(2)
'    sData = sData & RightH(Space(11) & txtInput(5).Text, 11) & Space(2)
'    sData = sData & RightH(Space(11) & txtInput(6).Text, 11) & Space(1)
'
'    P_00000.crPrint.Formulas(2) = "합계 = '" & sData & "'"
'
'    Call ReportPrint(ReportFile, "1")
End Sub

Public Sub DataScreen()
'    Dim ReportFP As String
'    Dim ReportFile As String
'    Dim sData As String
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
'    P_00000.crPrint.Formulas(0) = "년도 = '" & Format(dtInput.Value, "yyyy") & "'"
'    P_00000.crPrint.Formulas(1) = "대리점 = '" & Trim(Mid(cboInput.Text, 7)) & "'"
'
'    sData = Space(7) & RightH(Space(11) & txtInput(0).Text, 11) & Space(1)
'    sData = sData & RightH(Space(11) & txtInput(1).Text, 11) & Space(1)
'    sData = sData & RightH(Space(11) & txtInput(2).Text, 11) & Space(1)
'    sData = sData & RightH(Space(11) & txtInput(3).Text, 11) & Space(2)
'    sData = sData & RightH(Space(11) & txtInput(4).Text, 11) & Space(2)
'    sData = sData & RightH(Space(11) & txtInput(5).Text, 11) & Space(2)
'    sData = sData & RightH(Space(11) & txtInput(6).Text, 11) & Space(1)
'
'    P_00000.crPrint.Formulas(2) = "합계 = '" & sData & "'"
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
        
        spdView.Col = 1
        TempText = "   " & LeftH(spdView.Text & Space(10), 10)
        spdView.Col = 2
        TempText = TempText & RightH(Space(11) & spdView.Text, 11) & Space(1)
        spdView.Col = 3
        TempText = TempText & RightH(Space(11) & spdView.Text, 11) & Space(1)
        spdView.Col = 4
        TempText = TempText & RightH(Space(11) & spdView.Text, 11) & Space(1)
        spdView.Col = 5
        TempText = TempText & RightH(Space(11) & spdView.Text, 11) & Space(1)
        spdView.Col = 6
        TempText = TempText & RightH(Space(11) & spdView.Text, 11) & Space(1)
        spdView.Col = 7
        TempText = TempText & RightH(Space(11) & spdView.Text, 11) & Space(1)
        spdView.Col = 8
        TempText = TempText & RightH(Space(11) & spdView.Text, 11) & Space(1)
        spdView.Col = 9
        TempText = TempText & RightH(Space(11) & spdView.Text, 11) & Space(1)
        spdView.Col = 10
        TempText = TempText & RightH(Space(11) & spdView.Text, 11) & Space(1)
        spdView.Col = 11
        TempText = TempText & RightH(Space(11) & spdView.Text, 11) & Space(1)
        
        Print #1, TempText
    Next i
    
    Close #1
End Sub
