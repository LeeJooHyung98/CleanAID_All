VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form P_02002_01 
   Caption         =   "규정금액 CHECK"
   ClientHeight    =   11040
   ClientLeft      =   1725
   ClientTop       =   2520
   ClientWidth     =   15825
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
   ScaleHeight     =   11040
   ScaleWidth      =   15825
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   11040
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15825
      _ExtentX        =   27914
      _ExtentY        =   19473
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_02002_01.frx":0000
      Begin Threed.SSPanel SSPanel1 
         Height          =   465
         Left            =   15
         TabIndex        =   7
         Top             =   10560
         Width           =   15795
         _ExtentX        =   27861
         _ExtentY        =   820
         _Version        =   262144
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSPanel panCaption 
            Height          =   345
            Index           =   0
            Left            =   45
            TabIndex        =   8
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   609
            _Version        =   262144
            Caption         =   "금액합계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   345
            Index           =   1
            Left            =   3480
            TabIndex        =   9
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   609
            _Version        =   262144
            Caption         =   "판매금액합계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   345
            Index           =   3
            Left            =   6990
            TabIndex        =   10
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   609
            _Version        =   262144
            Caption         =   "차액합계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   345
            Index           =   0
            Left            =   1515
            TabIndex        =   11
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
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   345
            Index           =   1
            Left            =   4950
            TabIndex        =   12
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
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   345
            Index           =   2
            Left            =   8460
            TabIndex        =   13
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
         Width           =   15795
         _ExtentX        =   27861
         _ExtentY        =   1349
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Left            =   1530
            Style           =   2  '드롭다운 목록
            TabIndex        =   2
            Top             =   405
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Left            =   1530
            TabIndex        =   3
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   54919168
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   4
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
            TabIndex        =   5
            Top             =   405
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            BackColor       =   16777215
            Caption         =   "대 리 점 명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   9750
         Index           =   1
         Left            =   15
         TabIndex        =   6
         Top             =   795
         Width           =   15795
         _Version        =   524288
         _ExtentX        =   27861
         _ExtentY        =   17198
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
         SpreadDesigner  =   "P_02002_01.frx":0072
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_02002_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub dtInput_Change()
    Call Data_Display
End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(5).Enabled = True
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
    If P_02002_01_Flag = False Then
        dtInput.Value = P_02002.dtInput.Value
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_02002_02", sValue(), Err_Num, Err_Dec)
        
        spdView(1).MaxCols = RS01.Fields.Count
        spdView(1).MaxRows = RS01.RecordCount
        
        'Call spdDisplay2(RS01)
        Call fpSpread_Display(spdView(1), RS01)
        Call GetColWidth(REG_App, Me.Name, spdView(1))
        
        P_02002_01_Flag = True
    End If
End Sub

'Private Sub spdDisplay2(Rs As ADODB.Recordset)
'
'    Call fpSpread_Display(spdView(1), Rs)
'
'End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    With spdView(1)
        .ColsFrozen = 1 '틀고정
        .Row = -1
        
        .Col = 1                                  ' 택번호
        .ColWidth(1) = 8
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 2                                  ' 품명
        .ColWidth(2) = 20
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
        
        .Col = 3                                  ' 색상
        .ColWidth(3) = 8
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 4                                  ' 내용
        .ColWidth(4) = 8
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
        
        .Col = 5                                  ' 금액
        .ColWidth(5) = 12
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 0
        .TypeVAlign = TypeVAlignCenter
        
        .Col = 6                                  ' 판매금액
        .ColWidth(6) = 12
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 0
        .TypeVAlign = TypeVAlignCenter
    
        .Col = 7                                  ' 차액
        .ColWidth(7) = 12
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 0
        .TypeVAlign = TypeVAlignCenter
        
        .Col = 8                                  ' 수선
        .ColWidth(8) = 12
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 0
        .TypeVAlign = TypeVAlignCenter
    
        .Col = 9                                  ' 전화번호
        .ColWidth(9) = 10
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
    End With
    
    
    dtInput.Value = P_02002.dtInput.Value
    
    i = P_02002.ActiveControl.Index
    
    ReDim sValue(3)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput.Value, "YYYY-MM-DD")
    
    P_02002.spdView(i).Row = P_02002.spdView(i).ActiveRow
    P_02002.spdView(i).Col = 1
    
    sValue(2) = Mid(P_02002.spdView(i).Text, 2, 3)
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_02002_02", sValue(), Err_Num, Err_Dec)
    
    spdView(1).MaxCols = RS01.Fields.Count
    spdView(1).MaxRows = RS01.RecordCount
    
    'Call spdDisplay2(RS01)
    Call fpSpread_Display(spdView(1), RS01)
    Call GetColWidth(REG_App, Me.Name, spdView(1))

    spdView(1).AutoCalc = True
    
    spdView(1).MaxRows = spdView(1).MaxRows + 1
    spdView(1).Row = spdView(1).MaxRows
    
    spdView(1).RowHidden = True
    
    spdView(1).Col = 5: spdView(1).Formula = "SUM(E1:E" & spdView(1).MaxRows - 1 & ")"
                        txtNum(0).Text = spdView(1).Text
    spdView(1).Col = 6: spdView(1).Formula = "SUM(F1:F" & spdView(1).MaxRows - 1 & ")"
                        txtNum(1).Text = spdView(1).Text
    spdView(1).Col = 7: spdView(1).Formula = "SUM(G1:G" & spdView(1).MaxRows - 1 & ")"
                        txtNum(2).Text = spdView(1).Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_02002_01_Flag = True
End Sub

Public Sub Data_Display()
    Dim x As String
    
    ReDim sValue(2)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput.Value, "YYYY-MM-DD")
    sValue(2) = Mid(cboInput.Text, 2, 3)
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_02002_02", sValue(), Err_Num, Err_Dec)
    
    spdView(1).MaxCols = RS01.Fields.Count
    spdView(1).MaxRows = RS01.RecordCount
    
    'Call spdDisplay2(RS01)
    Call fpSpread_Display(spdView(1), RS01)
    Call GetColWidth(REG_App, Me.Name, spdView(1))
    
    spdView(1).AutoCalc = True
    
    spdView(1).MaxRows = spdView(1).MaxRows + 1
    spdView(1).Row = spdView(1).MaxRows
    
    spdView(1).RowHidden = True
    
    spdView(1).Col = 5: spdView(1).Formula = "SUM(E1:E" & spdView(1).MaxRows - 1 & ")"
                        txtNum(0).Text = spdView(1).Text
    spdView(1).Col = 6: spdView(1).Formula = "SUM(F1:F" & spdView(1).MaxRows - 1 & ")"
                        txtNum(1).Text = spdView(1).Text
    spdView(1).Col = 7: spdView(1).Formula = "SUM(G1:G" & spdView(1).MaxRows - 1 & ")"
                        txtNum(2).Text = spdView(1).Text
End Sub

Private Sub spdView_LeaveCell(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If NewRow <> -1 Then
        spdView(Index).Row = Row
        spdView(Index).Col = -1
        spdView(Index).BackColor = vbWhite
        
        spdView(Index).Row = NewRow
        spdView(Index).Col = -1
        spdView(Index).BackColor = glbYellow
    End If
End Sub

Private Sub spdView_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        'PopupMenu P_00000.PopUp
    End If
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
    P_00000.crPrint.Formulas(0) = "입고일자 = '" & Format(dtInput.Value, "YYYY-MM-DD") & "'"
    P_00000.crPrint.Formulas(1) = "대리점명 = '" & cboInput.Text & "'"
    P_00000.crPrint.Formulas(2) = "금액합계 = '" & txtNum(0).Text & "'"
    P_00000.crPrint.Formulas(3) = "판매금액합계 = '" & txtNum(1).Text & "'"
    P_00000.crPrint.Formulas(4) = "차액합계 = '" & txtNum(2).Text & "'"
    
    Call ReportPrint(ReportFile, "1")
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
    P_00000.crPrint.Formulas(0) = "입고일자 = '" & Format(dtInput.Value, "YYYY-MM-DD") & "'"
    P_00000.crPrint.Formulas(1) = "대리점명 = '" & cboInput.Text & "'"
    P_00000.crPrint.Formulas(2) = "금액합계 = '" & txtNum(0).Text & "'"
    P_00000.crPrint.Formulas(3) = "판매금액합계 = '" & txtNum(1).Text & "'"
    P_00000.crPrint.Formulas(4) = "차액합계 = '" & txtNum(2).Text & "'"
    
    Call ReportPrint(ReportFile, "2")
End Sub

Private Sub PrintDesc()
    Dim i, j As Integer
    
    Dim TempText As String
    Dim TempFP As String
    Dim TempFile As String
    
    TempFP = GetIniStr("REPORT", "TempPath", "", sIniFile)
    TempFile = TempFP & "\Temp.txt"
    
    Open TempFile For Output As #1
    
    TempText = ""
    
    For i = 1 To spdView(1).MaxRows
        spdView(1).Row = i
        
        spdView(1).Col = 1: TempText = Space(2) & LeftH(spdView(1).Text & Space(8), 8)
        spdView(1).Col = 2: TempText = TempText & LeftH(spdView(1).Text & Space(20), 20)
        spdView(1).Col = 3: TempText = TempText & LeftH(spdView(1).Text & Space(8), 8)
        spdView(1).Col = 4: TempText = TempText & LeftH(spdView(1).Text & Space(12), 12)
        spdView(1).Col = 5: TempText = TempText & RightH(Space(9) & spdView(1).Text, 9)
        spdView(1).Col = 6: TempText = TempText & RightH(Space(9) & spdView(1).Text, 9)
        spdView(1).Col = 7: TempText = TempText & RightH(Space(9) & spdView(1).Text, 9)
        spdView(1).Col = 8: TempText = TempText & RightH(Space(9) & spdView(1).Text, 9)
        spdView(1).Col = 9: TempText = TempText & LeftH(spdView(1).Text & Space(10), 10)
        
        Print #1, TempText
    Next i
    
    Close #1
End Sub
