VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form P_05011 
   Caption         =   "보관 서비스 상세 접수 현황"
   ClientHeight    =   12000
   ClientLeft      =   1485
   ClientTop       =   3285
   ClientWidth     =   17340
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_05011.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12000
   ScaleWidth      =   17340
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   12000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17340
      _ExtentX        =   30586
      _ExtentY        =   21167
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_05011.frx":058A
      Begin Threed.SSPanel panInput 
         Height          =   765
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   17310
         _ExtentX        =   30533
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
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   4770
            TabIndex        =   2
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   67895296
            CurrentDate     =   36686
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            TabIndex        =   3
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   67895296
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   5
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
            Index           =   2
            Left            =   60
            TabIndex        =   6
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "검 색 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin VB.Label Label1 
            Alignment       =   2  '가운데 맞춤
            Caption         =   "~"
            Height          =   195
            Left            =   4530
            TabIndex        =   7
            Top             =   120
            Width           =   195
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   11190
         Left            =   15
         TabIndex        =   8
         Top             =   795
         Width           =   17310
         _Version        =   524288
         _ExtentX        =   30533
         _ExtentY        =   19738
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
         SpreadDesigner  =   "P_05011.frx":05DC
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_05011"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub Form_Activate()
'    cmdBtn(0).Enabled = True
'    cmdBtn(5).Enabled = False
'    cmdBtn(6).Enabled = False
'
'    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    If P_05010_Flag = False Then
        Call ComboAdd

        dtInput(0).Value = Date
        dtInput(1).Value = Date

        ReDim sValue(3)

        sValue(0) = "1"

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_05011_01", sValue(), Err_Num, Err_Dec)

        spdView.MaxCols = RS01.Fields.Count
        spdView.MaxRows = RS01.RecordCount

        Call spdDisplay(RS01)
        Call GetColWidth(REG_App, Me.Name, spdView)

        P_05010_Flag = True
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
    spdView.ColWidth(1) = 20
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft
    
    spdView.Col = 2
    spdView.ColWidth(2) = 10
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter
    
    spdView.Col = 3
    spdView.ColWidth(3) = 15
    spdView.CellType = CellTypeDate
    spdView.TypeDateCentury = True
    spdView.TypeDateFormat = TypeDateFormatYYMMDD
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter
    
    spdView.Col = 4
    spdView.ColWidth(4) = 20
    spdView.CellType = CellTypeDate
    spdView.TypeDateCentury = True
    spdView.TypeDateFormat = TypeDateFormatYYMMDD
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter
    
    spdView.Col = 5
    spdView.ColWidth(5) = 10
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft
    
    spdView.Col = 6
    spdView.ColWidth(6) = 15
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter

    spdView.Col = 7
    spdView.ColWidth(7) = 10
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter

    spdView.Col = 8
    spdView.ColWidth(8) = 10
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft

    spdView.Col = 9
    spdView.ColWidth(9) = 16
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft

    spdView.Col = 10
    spdView.ColWidth(9) = 16
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 11
    spdView.ColWidth(9) = 16
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter

    spdView.Col = 12
    spdView.ColWidth(9) = 16
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.CellType = CellTypeFloat
    spdView.TypeVAlign = TypeVAlignCenter
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_05010_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i As Integer
    Dim iTotal(3) As Long

    ReDim sValue(3)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    sValue(3) = Mid(cboInput(0).Text, 2, 3) & "%"
    
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_05011_01", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    If RS01.RecordCount = 0 Then
        MsgBox "해당되는 데이터가 존재하지 않습니다.", vbInformation
        Exit Sub
    End If
    
    Call spdDisplay(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView)
        
    Exit Sub
    
ErrRtn:
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
'    For i = 0 To 30
'        P_00000.crPrint.Formulas(i) = ""
'    Next
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'    P_00000.crPrint.Formulas(0) = "일자1 = '" & Format(dtInput(0).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(1) = "일자2 = '" & Format(dtInput(1).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(2) = "대리점 = '" & cboInput(0).Text & "'"
'    P_00000.crPrint.Formulas(3) = "검색구분 = '입고일자'"
'
'    Call ReportPrint(ReportFile, "1")
End Sub

Private Sub spdView_DblClick(ByVal Col As Long, ByVal Row As Long)
    ' 접수일자
End Sub

Private Sub spdView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        'PopupMenu P_00000.PopUp
    End If
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
'    For i = 0 To 30
'        P_00000.crPrint.Formulas(i) = ""
'    Next
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'    P_00000.crPrint.Formulas(0) = "일자1 = '" & Format(dtInput(0).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(1) = "일자2 = '" & Format(dtInput(1).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(2) = "대리점 = '" & cboInput(0).Text & "'"
'    If optSelect(0).Value = True Then
'        P_00000.crPrint.Formulas(3) = "검색구분 = '입고일자'"
'    Else
'        P_00000.crPrint.Formulas(3) = "검색구분 = '출고일자'"
'    End If
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
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        
        spdView.Col = 1         ' 대리점명
        TempText = LeftH(spdView.Text & Space(14), 14)
        spdView.Col = 2         ' 택번호
        TempText = TempText & LeftH(spdView.Text & Space(7), 7)
        spdView.Col = 3         ' 입고일자
        TempText = TempText & LeftH(spdView.Text & Space(12), 12)
        spdView.Col = 4         ' 입고일자
        TempText = TempText & LeftH(spdView.Text & Space(12), 12)
        spdView.Col = 5         ' 품목명
        TempText = TempText & LeftH(spdView.Text & Space(16), 16)
        spdView.Col = 6         ' 색상
        TempText = TempText & LeftH(spdView.Text & Space(6), 6)
        spdView.Col = 7         ' 구분
        TempText = TempText & LeftH(spdView.Text & Space(6), 6)
        spdView.Col = 8         ' 접수내용
        TempText = TempText & LeftH(spdView.Text & Space(10), 10)
        spdView.Col = 9         ' 처리내용
        TempText = TempText & LeftH(spdView.Text & Space(10), 10)
        
        Print #1, TempText
    Next i
    
    Close #1
End Sub

Private Sub ComboAdd()
    Call AgencyComboAdd(cboInput(0))


End Sub
