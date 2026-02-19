VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_04003 
   Caption         =   "세금계산서 출력"
   ClientHeight    =   12045
   ClientLeft      =   1110
   ClientTop       =   5415
   ClientWidth     =   16470
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_04003.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12045
   ScaleWidth      =   16470
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   12045
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16470
      _ExtentX        =   29051
      _ExtentY        =   21246
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_04003.frx":058A
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10695
         Left            =   15
         TabIndex        =   1
         Top             =   1335
         Width           =   16440
         _Version        =   524288
         _ExtentX        =   28998
         _ExtentY        =   18865
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
         MaxCols         =   8
         SpreadDesigner  =   "P_04003.frx":061C
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   2
         Top             =   540
         Width           =   16440
         _ExtentX        =   28998
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   1
            Left            =   2235
            TabIndex        =   4
            Top             =   420
            Width           =   735
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   0
            Left            =   1245
            TabIndex        =   3
            Top             =   420
            Width           =   735
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   8265
            TabIndex        =   5
            Top             =   60
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            Format          =   61014016
            CurrentDate     =   37140
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   7080
            TabIndex        =   6
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "발행일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   7
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "집계년월"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   8
            Top             =   420
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "책 번 호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin XtremeSuiteControls.DateTimePicker dtInput 
            Height          =   330
            Left            =   1245
            TabIndex        =   20
            Top             =   60
            Width           =   1140
            _Version        =   851970
            _ExtentX        =   2011
            _ExtentY        =   582
            _StockProps     =   68
            CustomFormat    =   "yyyy-MM"
            Format          =   3
            UpDown          =   -1  'True
            CurrentDate     =   40544
         End
         Begin VB.Label Label1 
            Alignment       =   2  '가운데 맞춤
            Caption         =   "-"
            Height          =   135
            Left            =   1980
            TabIndex        =   9
            Top             =   480
            Width           =   255
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   10
         Top             =   15
         Width           =   8835
         _ExtentX        =   15584
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
         Caption         =   " #"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_04003.frx":0D31
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   1
         Left            =   8865
         TabIndex        =   11
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
         PictureBackground=   "P_04003.frx":0F33
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   12
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
            Picture         =   "P_04003.frx":1135
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   13
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
            Picture         =   "P_04003.frx":16CF
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   14
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
            Picture         =   "P_04003.frx":1C69
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   15
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
            Picture         =   "P_04003.frx":2203
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   16
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
            Picture         =   "P_04003.frx":279D
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   17
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
            Picture         =   "P_04003.frx":2D37
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   18
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
            Picture         =   "P_04003.frx":32D1
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   19
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
            Picture         =   "P_04003.frx":386B
         End
      End
   End
End
Attribute VB_Name = "P_04003"
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
        Case 2: Call DataSave       ' 저장
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

Private Sub Command1_Click()
    Dim nRow   As Long
    For nRow = 1 To spdView.MaxRows
        spdView.Col = 8:    spdView.Row = nRow
        spdView.Value = Not spdView.Value
    Next nRow
End Sub


Private Sub dtInput_Click()
    Dim sDate   As String
    
    sDate = DateAdd("M", 1, Format(dtInput.Value, "yyyy-mm" & "-01"))
    sDate = DateAdd("D", -1, Format(sDate, "YYYY-MM-DD"))
    DTPicker1.Value = sDate
End Sub

Private Sub dtInput_LostFocus()
    Dim sDate   As String
    
    sDate = DateAdd("M", 1, Format(dtInput.Value, "yyyy-mm" & "-01"))
    sDate = DateAdd("D", -1, Format(sDate, "YYYY-MM-DD"))
    DTPicker1.Value = sDate
End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(2).Enabled = True
    cmdBtn(3).Enabled = True
    cmdBtn(5).Enabled = True
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    If P_04003_Flag = False Then
        Dim sDate   As String
        
        dtInput.Value = Format(Date, "yyyy-mm")
        sDate = DateAdd("M", 1, Format(dtInput.Value, "yyyy-mm" & "-01"))
        sDate = DateAdd("D", -1, Format(sDate, "YYYY-MM-DD"))
        DTPicker1.Value = sDate
        
        ReDim sValue(2)
        
        sValue(0) = "1"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_04003_01", sValue(), Err_Num, Err_Dec)
        
        spdView.MaxCols = RS01.Fields.Count
        spdView.MaxRows = RS01.RecordCount
        
        Call spdDisplay(RS01)
        Call GetColWidth(REG_App, Me.Name, spdView)
        
        P_04003_Flag = True
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
    spdView.ColWidth(2) = 20
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft
    
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
    spdView.ColWidth(5) = 12
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 6
    spdView.ColWidth(6) = 6
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter

    spdView.Col = 7
    spdView.ColWidth(7) = 6
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter
    
    spdView.Col = 8
    spdView.ColWidth(8) = 6
    spdView.CellType = CellTypeCheckBox
    spdView.Value = False
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_04003_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i As Integer
    Dim bNo1 As Integer
    Dim bNo2 As Integer
    
    '----------------------------------------------------------------
    ' SP_04003_00
    '----------------------------------------------------------------
    ReDim sValue(0)
    
    sValue(0) = Format(dtInput.Value, "YYYY")
        
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_04003_00", sValue(), Err_Num, Err_Dec)
    
    If RS01.EOF Or IsNull(RS01!번호) Then
        txtInput(0).Text = 1
        txtInput(1).Text = 1
    Else
        If Val(Mid(RS01!번호, 7, 2)) > 49 Then
            txtInput(0).Text = Val(Mid(RS01!번호, 5, 2)) + 1
            txtInput(1).Text = 1
        Else
            txtInput(0).Text = Val(Mid(RS01!번호, 5, 2))
            txtInput(1).Text = Val(Mid(RS01!번호, 7, 2)) + 1
        End If
    End If
    
    '----------------------------------------------------------------
    ' SP_04003_01
    '----------------------------------------------------------------
    ReDim sValue(0)
    
    sValue(0) = Format(dtInput.Value, "YYYY-MM")
        
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_04003_01", sValue(), Err_Num, Err_Dec)
    
    With spdView
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = RS01!가맹점코드 & ""
            .Col = 2: .Text = RS01!가맹점명 & ""
            .Col = 3: .Text = RS01!공급가액 & ""
            .Col = 4: .Text = RS01!세액 & ""
            .Col = 5: .Text = RS01!합계금액 & ""
            .Col = 6: .Text = RS01!권 & ""
            .Col = 7: .Text = RS01!호 & ""
                        
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
    
    'spdView.MaxCols = RS01.Fields.Count
    'spdView.MaxRows = RS01.RecordCount
    
    'Call spdDisplay(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView)
    
    bNo1 = Val(txtInput(0).Text)
    bNo2 = Val(txtInput(1).Text)
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        spdView.Col = 6
            
        If Trim(spdView.Text) = "" Then
            spdView.Col = 6: spdView.Text = Right("00" & bNo1, 2)
            spdView.Col = 7: spdView.Text = Right("00" & bNo2, 2)
            
            bNo2 = bNo2 + 1
            If bNo2 > 50 Then
                bNo1 = bNo1 + 1
                bNo2 = 1
            End If
        End If
    Next i
    
    spdView.AutoCalc = True
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        
        spdView.Col = 4: spdView.Formula = "C" & i & " * 0.1"
        spdView.Col = 5: spdView.Formula = "SUM(C" & i & ":D" & i & ")"
    Next i
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Public Sub DataSave()
    Dim i As Integer
    Dim nCount  As Integer
    
    ReDim sValue(0)
    
    
    
    If MsgBox("발행일 : " & Format(DTPicker1.Value, "YYYY-MM-DD") & _
                "일자로 전송 하시겠습니까 ?", vbInformation + vbYesNo) = vbNo Then Exit Sub
    
            
    Set RS01 = New ADODB.Recordset
    
    sValue(0) = Format(dtInput.Value, "yyyymm")
        
    Call ExecPro("SP_04003_03", sValue(), Err_Num, Err_Dec)
    
    ReDim sValue(5)
    
    nCount = 0
    For i = 1 To spdView.MaxRows
        ReDim sValue(5)
            
        spdView.Row = i
                         sValue(0) = Format(dtInput.Value, "YYYY-MM")
        spdView.Col = 1: sValue(1) = spdView.Text
        spdView.Col = 3: sValue(2) = spdView.Value
        spdView.Col = 4: sValue(3) = spdView.Value
        spdView.Col = 5: sValue(4) = spdView.Value
        spdView.Col = 6: sValue(5) = spdView.Text
        
        Call ExecPro("SP_04003_02", sValue(), Err_Num, Err_Dec)
        
        spdView.Col = 8
        If Err_Num = 0 And spdView.Value = True Then
            ' exec  SP_04003_06 '1','200611','20061231','0001','001'
            
            '----------------------------------------------------------------
            ' SP_04003_06
            '----------------------------------------------------------------
            ReDim sValue(4)
            
            sValue(0) = "1"
            sValue(1) = Format(dtInput.Value, "YYYY-MM")
            sValue(2) = Format(DTPicker1.Value, "YYYY-MM-DD")
            sValue(3) = "0001"
            
            spdView.Col = 1
            sValue(4) = spdView.Text
            
            Set RS01 = ExecPro("SP_04003_06", sValue(), Err_Num, Err_Dec)
            
            If Left(CStr(RS01.Fields(0)), 2) <> "OK" Then
                  MsgBox CStr(RS01.Fields(0)), vbInformation, "확인"
            End If
            
            RS01.Close
            nCount = nCount + 1
            
        ElseIf Err_Num <> 0 Then
            MsgBox Err_Dec, vbInformation, "확인"
            Err_Num = 0
        End If
    
    Next i
    
    If nCount = 0 Then
        MsgBox "전송할 체인점을 먼저 선택하여 주십시요.", vbInformation, "확인"
        Exit Sub
    Else
        MsgBox CStr(nCount) & " 건을 전송하였습니다.", vbInformation, "확인"
        Exit Sub
    End If
End Sub

Public Sub DataDelete()
    If MsgBox("세금계산서내역을 삭제하시겠습니까?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        ReDim sValue(0)
        
        sValue(0) = Format(dtInput.Value, "yyyymm")
        
        Call ExecPro("SP_04003_03", sValue(), Err_Num, Err_Dec)
        
        If Err_Num = 0 Then
            MsgBox "정상적으로 삭제되었습니다.", vbInformation
        End If
    End If
End Sub

Public Sub DataScreen()
    
'    Load P_PRTSCREEN
'    P_PRTSCREEN.Show

End Sub

Public Sub DataPrint()
'    Dim i, ii As Integer
'    Dim ReportFP As String
'    Dim ReportFile As String
'    Dim sData As String
'
'    sValue(0) = "0"
'    sValue(1) = Format(dtInput.Value, "yyyymm")
'
'    Set RS01 = New ADODB.Recordset
'    Set RS01 = ExecPro("SP_04003_04", sValue(), Err_Num, Err_Dec)
'
'    If RS01!레코드건수 = 0 Then
'        MsgBox "해당데이터를 저장한 후에 출력을 하십시오", vbInformation
'        Exit Sub
'    End If
'
'    ReportFP = GetIniStr("REPORT", "FilePath", "", m_iniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    Call PrintDesc
'
'    For i = 0 To 30
'        P_00000.crPrint.Formulas(i) = ""
'    Next
'
'    P_00000.crPrint.SelectionFormula = ""
'    P_00000.crPrint.Formulas(0) = "월 = '" & Format(dtInput.Value, "mm") & "'"
'    P_00000.crPrint.Formulas(1) = "일 = '" & Format(DateAdd("d", -1, DateAdd("m", 1, Format(dtInput.Value, "yyyy-mm") & "-01")), "dd") & "'"
'    P_00000.crPrint.Formulas(2) = "Pdate = '" & Format(DateAdd("d", -1, DateAdd("m", 1, Format(dtInput.Value, "yyyy-mm") & "-01")), "YYYY-MM-DD") & "'"
'
'    Call ReportPrint(ReportFile, "2")
End Sub

Private Sub PrintDesc()
    Dim i As Integer
    Dim TempText As String
    Dim TempFP As String
    Dim TempFile As String
    
    Dim iDanga As Long
    Dim Gong As Integer
    
    On Error GoTo FileError:
    
    TempFP = GetIniStr("REPORT", "TempPath", "", m_iniFile)
    TempFile = TempFP & "\Temp.txt"
    
    Open TempFile For Output As #1
    
    ReDim sValue(1)
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        spdView.Col = 8
        If spdView.Value = True Then
            spdView.Col = 1
            
            sValue(0) = "0"
            sValue(1) = spdView.Text
                
            Set RS01 = New ADODB.Recordset
            Set RS01 = ExecPro("SP_04003_05", sValue(), Err_Num, Err_Dec)
            
            TempText = ""
            TempText = TempText & LeftH(IIf(IsNull(RS01!FullName), " ", RS01!FullName) & Space(30), 30)
            TempText = TempText & LeftH(RS01!Representation & Space(12), 12)
            TempText = TempText & LeftH(RS01!EnterpriseNo & Space(12), 12)
            TempText = TempText & LeftH(RS01!EnterpriseTemp1 & Space(12), 12)
            TempText = TempText & LeftH(RS01!EnterpriseTemp2 & Space(12), 12)
            TempText = TempText & LeftH(RS01!Address & Space(44), 44)
            
            spdView.Col = 3
            TempText = TempText & RightH(Space(10) & spdView.Value, 10)
            spdView.Col = 4
            TempText = TempText & RightH(Space(9) & spdView.Value, 9)
            spdView.Col = 5
            TempText = TempText & RightH(Space(10) & spdView.Value, 10)
            spdView.Col = 6
            TempText = TempText & RightH(Space(2) & spdView.Value, 2)
            spdView.Col = 7
            TempText = TempText & RightH(Space(2) & spdView.Value, 2)
            
            Gong = 0
            
            spdView.Col = 3
            If spdView.Value < 1000000000 Then
               Gong = 1
            End If
            If spdView.Value < 100000000 Then
               Gong = 2
            End If
            If spdView.Value < 10000000 Then
               Gong = 3
            End If
            If spdView.Value < 1000000 Then
               Gong = 4
            End If
            If spdView.Value < 100000 Then
               Gong = 5
            End If
            If spdView.Value < 10000 Then
               Gong = 6
            End If
            If spdView.Value < 1000 Then
               Gong = 7
            End If
            If spdView.Value < 100 Then
               Gong = 8
            End If
            If spdView.Value < 10 Then
               Gong = 9
            End If
            TempText = TempText & Gong
            
            Print #1, TempText
        End If
    Next i
    
    Close #1
    
FileError:
    If Err.Number = 55 Then
        Resume Next
    End If
End Sub

