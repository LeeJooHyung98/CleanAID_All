VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form P_02004_2 
   Caption         =   "제품별 입고현황"
   ClientHeight    =   8115
   ClientLeft      =   735
   ClientTop       =   1785
   ClientWidth     =   16380
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
   ScaleHeight     =   8115
   ScaleWidth      =   16380
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8115
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16380
      _ExtentX        =   28893
      _ExtentY        =   14314
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_02004_2.frx":0000
      Begin Threed.SSPanel SSPanel1 
         Height          =   480
         Left            =   15
         TabIndex        =   13
         Top             =   7620
         Width           =   16350
         _ExtentX        =   28840
         _ExtentY        =   847
         _Version        =   262144
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSPanel panCaption 
            Height          =   345
            Index           =   1
            Left            =   60
            TabIndex        =   14
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   609
            _Version        =   262144
            Caption         =   "수 량 합 계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   345
            Index           =   3
            Left            =   3450
            TabIndex        =   15
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   609
            _Version        =   262144
            Caption         =   "금 액 합 계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   345
            Index           =   0
            Left            =   1530
            TabIndex        =   16
            Top             =   60
            Width           =   1050
            _Version        =   262145
            _ExtentX        =   1852
            _ExtentY        =   609
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   3
            StartText.y     =   5
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   13
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   345
            Index           =   1
            Left            =   4920
            TabIndex        =   17
            Top             =   60
            Width           =   1425
            _Version        =   262145
            _ExtentX        =   2514
            _ExtentY        =   609
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   3
            StartText.y     =   5
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   13
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
      End
      Begin Threed.SSPanel panInput 
         Height          =   765
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   16350
         _ExtentX        =   28840
         _ExtentY        =   1349
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            Style           =   2  '드롭다운 목록
            TabIndex        =   6
            Top             =   420
            Width           =   3015
         End
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   1
            Left            =   4830
            Style           =   2  '드롭다운 목록
            TabIndex        =   3
            Top             =   420
            Width           =   2955
         End
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   2
            Left            =   9570
            Style           =   2  '드롭다운 목록
            TabIndex        =   2
            Top             =   420
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   4830
            TabIndex        =   4
            Top             =   60
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   55377920
            CurrentDate     =   36686
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            TabIndex        =   5
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   55377920
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   8100
            TabIndex        =   7
            Top             =   420
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            BackColor       =   16777215
            Caption         =   "대 리 점 명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   8
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            BackColor       =   16777215
            Caption         =   "입 고 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   7
            Left            =   60
            TabIndex        =   9
            Top             =   420
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            BackColor       =   16777215
            Caption         =   "품  목  명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin VB.Label Label1 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "~"
            Height          =   255
            Left            =   4530
            TabIndex        =   11
            Top             =   120
            Width           =   255
         End
         Begin VB.Label Label2 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "~"
            Height          =   255
            Left            =   4530
            TabIndex        =   10
            Top             =   480
            Width           =   255
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   6810
         Left            =   15
         TabIndex        =   12
         Top             =   795
         Width           =   16350
         _Version        =   524288
         _ExtentX        =   28840
         _ExtentY        =   12012
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
         SpreadDesigner  =   "P_02004_2.frx":0072
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_02004_2"
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

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(5).Enabled = True
    cmdBtn(6).Enabled = True
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"

    If P_02005_Flag = False Then
        Call GoodsComboAdd(cboInput(0))
        Call GoodsComboAdd(cboInput(1))
        Call AgencyComboAdd(cboInput(2))
        
        dtInput(0).Value = Date
        dtInput(1).Value = Date
        
        ReDim sValue(6)
        
        sValue(0) = "1"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_02004_02", sValue(), Err_Num, Err_Dec)
        
        spdView.MaxCols = RS01.Fields.Count
        spdView.MaxRows = RS01.RecordCount
        
        'Call spdDisplay(RS01)
        Call fpSpread_Display(spdView, RS01)
        Call GetColWidth("백상", Me.Name, spdView)
        
        P_02005_Flag = True
    End If
End Sub

'Private Sub spdDisplay(Rs As ADODB.Recordset)
'    Call fpSpread_Display(spdView, Rs)
'End Sub

Private Sub Form_Load()
    With spdView
        .ColsFrozen = 1 '틀고정
        .Row = -1
        
        .Col = 1
        .ColWidth(1) = 16
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
        
        .Col = 2
        .ColWidth(2) = 8
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 3
        .ColWidth(3) = 12
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 0
        .TypeVAlign = TypeVAlignCenter
        
        .Col = 4
        .ColWidth(4) = 16
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 0
        .TypeVAlign = TypeVAlignCenter
    
        .Col = 5
        .ColWidth(5) = 16
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 2
        .TypeVAlign = TypeVAlignCenter
    
        .Col = 6
        .ColWidth(6) = 16
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignRight
    
        .Col = 7
        .ColWidth(7) = 16
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignRight
    
        .Col = 8
        .ColWidth(8) = 16
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignRight
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_02005_Flag = False
End Sub

Public Sub Data_Display()
    Dim i As Integer
    Dim j As Integer
    Dim lTemp(3) As Single
    Dim sTemp    As String
    
    ReDim sValue(6)
    
    sTemp = ""
    sValue(0) = "0"
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    If cboInput(0).Text = "" Then sValue(3) = "000" Else sValue(3) = Mid(cboInput(0).Text, 2, 3)
    If cboInput(0).Text = "" Then sValue(4) = "ZZZ" Else sValue(4) = Mid(cboInput(1).Text, 2, 3)
    
    sValue(5) = Mid(cboInput(2).Text, 2, 3) & "%"
    
    If sValue(3) > sValue(4) Then
        MsgBox "품목선택이 조전에 맞지 않습니다.", vbInformation
        Exit Sub
    End If
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_02004_02", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    'Call spdDisplay(RS01)
    Call fpSpread_Display(spdView, RS01)
    
    ReDim sValue(5)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    If cboInput(0).Text = "" Then sValue(3) = "000" Else sValue(3) = Mid(cboInput(0).Text, 2, 3)
    If cboInput(0).Text = "" Then sValue(4) = "ZZZ" Else sValue(4) = Mid(cboInput(1).Text, 2, 3)
    
    sValue(5) = Mid(cboInput(2).Text, 2, 3) & "%"

    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_02004_03", sValue(), Err_Num, Err_Dec)
    
    ReDim sValue(5)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    If cboInput(0).Text = "" Then sValue(3) = "000" Else sValue(3) = Mid(cboInput(0).Text, 2, 3)
    If cboInput(0).Text = "" Then sValue(4) = "ZZZ" Else sValue(4) = Mid(cboInput(1).Text, 2, 3)
    
    sValue(5) = Mid(cboInput(2).Text, 2, 3) & "%"

    Set RS02 = New ADODB.Recordset
    Set RS02 = ExecPro("SP_02004_04", sValue(), Err_Num, Err_Dec)
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        
        spdView.Col = 4
        lTemp(0) = Val(spdView.Value) / RS02!총금액 * 100
        
        spdView.Col = 5
        spdView.Text = lTemp(0)
    Next i
    
    For i = 1 To RS01.RecordCount
        For j = 1 To spdView.MaxRows
            spdView.Row = j
            spdView.Col = 2
            If Left(spdView.Value, 1) = RS01!상품코드 Then
                If sTemp <> RS01!상품코드 Then
                    sTemp = RS01!상품코드
                    spdView.Col = 6:  spdView.Text = Format(RS01!총수량, "#,##0")
                    spdView.Col = 7:  spdView.Text = Format(RS01!총금액, "#,##0")
                    spdView.Col = 8:  spdView.Text = Format(Val(RS01!총금액) / RS02!총금액 * 100, "#,##0.00")
                    spdView.Col = -1: spdView.BackColor = &HD8FCFE
                End If
            End If
        Next j
        RS01.MoveNext
    Next i
    
    spdView.AutoCalc = True
    
    spdView.MaxRows = spdView.MaxRows + 1
    spdView.Row = spdView.MaxRows
    
    spdView.RowHidden = True
    
    txtNum(0).Text = Format(RS02!총수량, "#,##0")
    txtNum(1).Text = Format(RS02!총금액, "#,##0")

End Sub

Public Sub DataPrint()
    Dim ReportFP As String
    Dim ReportFile As String
    
    ReportFP = GetIniStr("REPORT", "FilePath", "", sIniFile)
    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
    
    Call PrintDesc
    
    Dim i As Integer
    For i = 0 To 30
        P_00000.crPrint.Formulas(i) = ""
    Next
    
    P_00000.crPrint.WindowTitle = Me.Caption
    P_00000.crPrint.Formulas(0) = "입고일자1 = '" & Format(dtInput(0).Value, "YYYY-MM-DD") & "'"
    P_00000.crPrint.Formulas(1) = "입고일자2 = '" & Format(dtInput(1).Value, "YYYY-MM-DD") & "'"
    P_00000.crPrint.Formulas(2) = "순위구분 = '" & IIf(optSelect(0).Value = True, "금액", "수량") & "'"
    P_00000.crPrint.Formulas(3) = "대리점명 = '" & cboInput(2).Text & "'"
    P_00000.crPrint.Formulas(4) = "품목명1 = '" & cboInput(0).Text & "'"
    P_00000.crPrint.Formulas(5) = "품목명2 = '" & cboInput(1).Text & "'"
    
    P_00000.crPrint.Formulas(6) = "수량합계 = '" & txtNum(0).Text & "'"
    P_00000.crPrint.Formulas(7) = "금액합계 = '" & txtNum(1).Text & "'"
    P_00000.crPrint.Formulas(8) = "점유율(단위)수량 = '" & txtNum(2).Text & "'"
    P_00000.crPrint.Formulas(9) = "점유율(단위)금액 = '" & txtNum(3).Text & "'"
    P_00000.crPrint.Formulas(10) = "점유율(전체)수량 = '" & txtNum(4).Text & "'"
    P_00000.crPrint.Formulas(11) = "점유율(전체)금액 = '" & txtNum(5).Text & "'"
    
    Call ReportPrint(ReportFile, "1")
End Sub

Private Sub spdView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        'PopupMenu P_00000.PopUp
    End If
End Sub

Public Sub DataScreen()
    Dim ReportFP As String
    Dim ReportFile As String
    
    ReportFP = GetIniStr("REPORT", "FilePath", "", sIniFile)
    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
    
    Call PrintDesc
    
    Dim i As Integer
    For i = 0 To 30
        P_00000.crPrint.Formulas(i) = ""
    Next
    
    P_00000.crPrint.WindowTitle = Me.Caption
    P_00000.crPrint.Formulas(0) = "입고일자1 = '" & Format(dtInput(0).Value, "YYYY-MM-DD") & "'"
    P_00000.crPrint.Formulas(1) = "입고일자2 = '" & Format(dtInput(1).Value, "YYYY-MM-DD") & "'"
    P_00000.crPrint.Formulas(2) = "순위구분 = '" & IIf(optSelect(0).Value = True, "금액", "수량") & "'"
    P_00000.crPrint.Formulas(3) = "대리점명 = '" & cboInput(2).Text & "'"
    P_00000.crPrint.Formulas(4) = "품목명1 = '" & cboInput(0).Text & "'"
    P_00000.crPrint.Formulas(5) = "품목명2 = '" & cboInput(1).Text & "'"
    
    P_00000.crPrint.Formulas(6) = "수량합계 = '" & txtNum(0).Text & "'"
    P_00000.crPrint.Formulas(7) = "금액합계 = '" & txtNum(1).Text & "'"
    P_00000.crPrint.Formulas(8) = "점유율(단위)수량 = '" & txtNum(2).Text & "'"
    P_00000.crPrint.Formulas(9) = "점유율(단위)금액 = '" & txtNum(3).Text & "'"
    P_00000.crPrint.Formulas(10) = "점유율(전체)수량 = '" & txtNum(4).Text & "'"
    P_00000.crPrint.Formulas(11) = "점유율(전체)금액 = '" & txtNum(5).Text & "'"
    
    Call ReportPrint(ReportFile, "2")
End Sub

Private Sub PrintDesc()
    Dim i As Integer
    Dim TempText As String
    Dim TempFP As String
    Dim TempFile As String
    
    TempFP = GetIniStr("REPORT", "TempPath", "", sIniFile)
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
