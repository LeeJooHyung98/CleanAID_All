VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form P_06004 
   Caption         =   "사고처리 내역"
   ClientHeight    =   8640
   ClientLeft      =   615
   ClientTop       =   6225
   ClientWidth     =   16650
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_06004.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8640
   ScaleWidth      =   16650
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8640
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16650
      _ExtentX        =   29369
      _ExtentY        =   15240
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_06004.frx":058A
      Begin Threed.SSPanel SSPanel1 
         Height          =   420
         Left            =   15
         TabIndex        =   16
         Top             =   8205
         Width           =   16620
         _ExtentX        =   29316
         _ExtentY        =   741
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   0
            Left            =   1530
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   60
            Width           =   2115
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   1
            Left            =   5310
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   60
            Width           =   2115
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   2
            Left            =   9090
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   60
            Width           =   2115
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   3
            Left            =   12870
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   60
            Width           =   2115
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   7
            Left            =   60
            TabIndex        =   21
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "총  건  수"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   8
            Left            =   3840
            TabIndex        =   22
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "제품금액계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   9
            Left            =   7620
            TabIndex        =   23
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "보 상 건 수"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   10
            Left            =   11400
            TabIndex        =   24
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "보상금액계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel panInput 
         Height          =   765
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   16620
         _ExtentX        =   29316
         _ExtentY        =   1349
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            Style           =   2  '드롭다운 목록
            TabIndex        =   4
            Top             =   420
            Width           =   3735
         End
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   2
            Left            =   7020
            Style           =   2  '드롭다운 목록
            TabIndex        =   3
            Top             =   420
            Width           =   3735
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   4
            Left            =   12570
            TabIndex        =   2
            Top             =   420
            Width           =   1155
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   1530
            TabIndex        =   5
            Top             =   60
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   262144
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   0
               Left            =   420
               TabIndex        =   6
               Top             =   30
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "접수일자"
               Value           =   -1
            End
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   7020
            TabIndex        =   7
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   64487424
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   5550
            TabIndex        =   8
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "검 색 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   10335
            TabIndex        =   9
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   64487424
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   10
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "검 색 기 준"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   3
            Left            =   60
            TabIndex        =   11
            Top             =   420
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "지 사 명 칭"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   5
            Left            =   5550
            TabIndex        =   12
            Top             =   420
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "대 리 점 명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   11
            Left            =   11100
            TabIndex        =   13
            Top             =   420
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "접 수 번 호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin VB.Label Label1 
            Alignment       =   2  '가운데 맞춤
            Caption         =   "~"
            Height          =   195
            Left            =   10035
            TabIndex        =   14
            Top             =   120
            Width           =   255
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   7395
         Left            =   15
         TabIndex        =   15
         Top             =   795
         Width           =   16620
         _Version        =   524288
         _ExtentX        =   29316
         _ExtentY        =   13044
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
         SpreadDesigner  =   "P_06004.frx":05FC
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_06004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cboInput_Click(Index As Integer)
    Dim sCode As String

    If Index = 0 Then
        sCode = Trim(Mid(Trim(cboInput(0)) & Space(10), 2, 4))

        Call Get_가맹점리스트(cboInput(2), sCode)
    End If

End Sub

Private Sub Form_Activate()
'    cmdBtn(0).Enabled = True
'    cmdBtn(5).Enabled = True
'    cmdBtn(6).Enabled = True
'
'    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn

    If P_06004_Flag = False Then
        dtInput(0).Value = Date
        dtInput(1).Value = Date
        
        Call ComboAdd
        
        ReDim sValue(8)
        
        sValue(0) = "1"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_06002_00", sValue(), Err_Num, Err_Dec)
        
        spdView.MaxCols = RS01.Fields.Count
        spdView.MaxRows = RS01.RecordCount
        
        Call spdDisplay(RS01)
        Call GetColWidth(REG_App, Me.Name, spdView)
        
        Call Master_tblComboAdd(cboInput(0))

        
        P_06004_Flag = True
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
    spdView.CellType = CellTypeDate
    spdView.TypeDateCentury = True
    spdView.TypeDateFormat = TypeDateFormatYYMMDD
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft
    
    spdView.Col = 2
    spdView.ColWidth(2) = 8
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter
    
    spdView.Col = 3
    spdView.ColWidth(3) = 20
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft

    spdView.Col = 4
    spdView.ColWidth(4) = 7
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter

    spdView.Col = 5
    spdView.ColWidth(5) = 20
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft
    
    spdView.Col = 6
    spdView.ColWidth(6) = 10
    spdView.CellType = CellTypeDate
    spdView.TypeDateCentury = True
    spdView.TypeDateFormat = TypeDateFormatYYMMDD
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft
    
    spdView.Col = 7
    spdView.ColWidth(7) = 10
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft
    
    spdView.Col = 8
    spdView.ColWidth(8) = 12
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 9
    spdView.ColWidth(9) = 12
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 10
    spdView.ColWidth(10) = 10
    spdView.CellType = CellTypeDate
    spdView.TypeDateCentury = True
    spdView.TypeDateFormat = TypeDateFormatYYMMDD
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_06004_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i As Integer
    
    ReDim sValue(5)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    sValue(3) = "9999" 'Trim(Mid(cboInput(0).Text, 2, 4))
    sValue(4) = Trim(Mid(cboInput(2).Text, 2, 6))
    
    
    If optSelect(0).Value = True Then
        sValue(5) = "1"
    ElseIf optSelect(1).Value = True Then
        sValue(5) = "2"
    End If
        
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_06004_00", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    Call spdDisplay(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView)

    txtInput(0).Text = ""
    txtInput(1).Text = ""
    txtInput(2).Text = ""
    txtInput(3).Text = ""
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        
        spdView.Col = 8: txtInput(1).Text = Val(txtInput(1).Text) + spdView.Value
        spdView.Col = 9: txtInput(3).Text = Val(txtInput(3).Text) + spdView.Value
        
        spdView.Col = 9
        If spdView.Value <> "" And spdView.Value <> 0 Then
            txtInput(2).Text = Val(txtInput(2).Text) + 1
        End If
    Next i
    
    txtInput(0).Text = Format(spdView.MaxRows, "#,##0")
    txtInput(1).Text = Format(txtInput(1).Text, "#,##0")
    txtInput(2).Text = Format(txtInput(2).Text, "#,##0")
    txtInput(3).Text = Format(txtInput(3).Text, "#,##0")
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub ComboAdd()
    Call AgencyComboAdd(cboInput(0))

    ReDim sValue(1)
    
    sValue(0) = "0"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_00001", sValue(), Err_Num, Err_Dec)

    cboInput(2).AddItem ""

    Do While Not RS01.EOF
        cboInput(2).AddItem "[" & RS01!담당자코드 & "] " & RS01!담당자명
        
        RS01.MoveNext
    Loop

    sValue(0) = "0"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_00002", sValue(), Err_Num, Err_Dec)

    

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



Public Sub DataPrint()
'    Dim ReportFP As String
'    Dim ReportFile As String
'
'    ReportFP = GetIniStr("REPORT", "FilePath", "", m_iniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    Call PrintDesc
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'
'    Dim i As Integer
'    For i = 0 To 30
'        P_00000.crPrint.Formulas(i) = ""
'    Next
'
'    If optSelect(0).Value = True Then
'        P_00000.crPrint.Formulas(0) = "검색기준 = '접수일자'"
'    ElseIf optSelect(1).Value = True Then
'        P_00000.crPrint.Formulas(0) = "검색기준 = '입고일자'"
'    End If
'    P_00000.crPrint.Formulas(1) = "기사명 = '" & Mid(cboInput(3).Text, 7) & "'"
'    P_00000.crPrint.Formulas(2) = "담당자명 = '" & Mid(cboInput(2).Text, 7) & "'"
'    P_00000.crPrint.Formulas(3) = "대리점명 = '" & Mid(cboInput(0).Text, 7) & "'"
'    P_00000.crPrint.Formulas(4) = "보상건수 = '" & txtInput(2).Text & "'"
'    P_00000.crPrint.Formulas(5) = "접수일자1 = '" & Format(dtInput(0).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(6) = "접수일자2 = '" & Format(dtInput(1).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(7) = "크레임명 = '" & cboInput(1).Text & "'"
'    P_00000.crPrint.Formulas(8) = "크레임총건수 = '" & txtInput(0).Text & "'"
'    P_00000.crPrint.Formulas(9) = "합계금액 = '" & txtInput(3).Text & "'"
'    P_00000.crPrint.Formulas(10) = "제품금액 = '" & txtInput(1).Text & "'"
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
'    P_00000.crPrint.WindowTitle = Me.Caption
'
'    Dim i As Integer
'    For i = 0 To 30
'        P_00000.crPrint.Formulas(i) = ""
'    Next
'
'    If optSelect(0).Value = True Then
'        P_00000.crPrint.Formulas(0) = "검색기준 = '접수일자'"
'    ElseIf optSelect(1).Value = True Then
'        P_00000.crPrint.Formulas(0) = "검색기준 = '입고일자'"
'    End If
'    P_00000.crPrint.Formulas(1) = "기사명 = '" & Mid(cboInput(3).Text, 7) & "'"
'    P_00000.crPrint.Formulas(2) = "담당자명 = '" & Mid(cboInput(2).Text, 7) & "'"
'    P_00000.crPrint.Formulas(3) = "대리점명 = '" & Mid(cboInput(0).Text, 7) & "'"
'    P_00000.crPrint.Formulas(4) = "보상건수 = '" & txtInput(2).Text & "'"
'    P_00000.crPrint.Formulas(5) = "접수일자1 = '" & Format(dtInput(0).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(6) = "접수일자2 = '" & Format(dtInput(1).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(7) = "크레임명 = '" & cboInput(1).Text & "'"
'    P_00000.crPrint.Formulas(8) = "크레임총건수 = '" & txtInput(0).Text & "'"
'    P_00000.crPrint.Formulas(9) = "합계금액 = '" & txtInput(3).Text & "'"
'    P_00000.crPrint.Formulas(10) = "제품금액 = '" & txtInput(1).Text & "'"
'
'    Call ReportPrint(ReportFile, "2")
End Sub

Private Sub PrintDesc()
    Dim i As Integer
    Dim TempText As String
    Dim TempFP As String
    Dim TempFile As String
    
    TempFP = GetIniStr("REPORT", "TempPath", "", m_iniFile)
    If CheckDirectory(TempFP, True) = False Then
        Exit Sub
    End If
    TempFile = TempFP & "\Temp.txt"
    Open TempFile For Output As #1
    
    TempText = ""
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        
        spdView.Col = 1
        TempText = LeftH(spdView.Text & Space(12), 12)
        spdView.Col = 2
        TempText = TempText & LeftH(spdView.Text & Space(6), 6)
        spdView.Col = 3
        TempText = TempText & LeftH(spdView.Text & Space(14), 14)
        spdView.Col = 7
        TempText = TempText & LeftH(spdView.Text & Space(8), 8)
        spdView.Col = 4
        TempText = TempText & LeftH(spdView.Text & Space(6), 6)
        spdView.Col = 5
        TempText = TempText & LeftH(spdView.Text & Space(12), 12)
        spdView.Col = 6
        TempText = TempText & LeftH(spdView.Text & Space(12), 12)
        spdView.Col = 8
        TempText = TempText & RightH(Space(9) & spdView.Text, 9) & Space(1)
        spdView.Col = 9
        TempText = TempText & RightH(Space(9) & spdView.Text, 9) & Space(1)
        spdView.Col = 10
        TempText = TempText & LeftH(spdView.Text & Space(12), 12)
        
        Print #1, TempText
    Next i
    
    Close #1
End Sub
