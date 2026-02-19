VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form P_05006 
   Caption         =   "재세탁 관리"
   ClientHeight    =   12060
   ClientLeft      =   2700
   ClientTop       =   1920
   ClientWidth     =   16140
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_05006.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12060
   ScaleWidth      =   16140
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   12060
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16140
      _ExtentX        =   28469
      _ExtentY        =   21273
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_05006.frx":058A
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10815
         Left            =   15
         TabIndex        =   1
         Top             =   1230
         Width           =   16110
         _Version        =   524288
         _ExtentX        =   28416
         _ExtentY        =   19076
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
         SpreadDesigner  =   "P_05006.frx":05FC
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   2
         Top             =   435
         Width           =   16110
         _ExtentX        =   28416
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Left            =   6210
            Style           =   2  '드롭다운 목록
            TabIndex        =   3
            Top             =   60
            Width           =   2775
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Left            =   1530
            TabIndex        =   4
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   21430272
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   4740
            TabIndex        =   5
            Top             =   60
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
            Caption         =   "입 고 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   405
         Left            =   15
         TabIndex        =   7
         Top             =   15
         Width           =   16110
         _ExtentX        =   28416
         _ExtentY        =   714
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
         Caption         =   " 재세탁 관리 (P_05006)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_05006.frx":0A64
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
End
Attribute VB_Name = "P_05006"
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
'    cmdBtn(2).Enabled = True
'    cmdBtn(5).Enabled = True
'    cmdBtn(6).Enabled = True
'
'    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    If P_05006_Flag = False Then
        dtInput.Value = Date
        
        Call AgencyComboAdd(cboInput)
        
        ReDim sValue(2)
        
        sValue(0) = "1"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_05006_00", sValue(), Err_Num, Err_Dec)
        
        spdView.MaxCols = RS01.Fields.Count
        spdView.MaxRows = RS01.RecordCount
        
        Call spdDisplay(RS01)
        Call GetColWidth(REG_App, Me.Name, spdView)
        
        P_05006_Flag = True
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
    spdView.ColWidth(2) = 8
    spdView.CellType = CellTypePic
    spdView.TypePicMask = "9-999"
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter
    
    spdView.Col = 3
    spdView.ColWidth(3) = 20
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft
    
    spdView.Col = 4
    spdView.ColWidth(4) = 8
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter
    
    spdView.Col = 5
    spdView.ColWidth(5) = 14
    spdView.CellType = CellTypeComboBox
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft

    spdView.Col = 6
    spdView.ColWidth(6) = 23
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft

    spdView.Col = 7
    spdView.ColWidth(7) = 23
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft

    Call ComboAdd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_05006_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i As Integer
    
    ReDim sValue(2)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput.Value, "YYYY-MM-DD")
    sValue(2) = Mid(cboInput.Text, 2, 6) & "%"
        
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_05006_00", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    Call spdDisplay(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView)
    
    Do While Not RS01.EOF
        i = i + 1
        
        spdView.Row = i
        spdView.Col = 5
        If Not IsNull(RS01!구분) Then spdView.Text = RS01!구분
        
        RS01.MoveNext
    Loop
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub spdView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        'PopupMenu P_00000.PopUp
    End If
End Sub

Private Sub ComboAdd()
    Dim sItem As String

    sItem = "오염" & Chr(9)
    sItem = sItem & "이염" & Chr(9)
    sItem = sItem & "수축" & Chr(9)
    sItem = sItem & "늘어짐" & Chr(9)
    sItem = sItem & "경화" & Chr(9)
    sItem = sItem & "파손" & Chr(9)
    sItem = sItem & "탈색" & Chr(9)
    sItem = sItem & "기타" & Chr(9)

    spdView.Row = -1
    spdView.Col = 5
    spdView.TypeComboBoxList = sItem
End Sub

Public Sub DataSave()
    Dim i As Integer
    
    ReDim sValue(5)
    
    For i = 1 To spdView.MaxRows
        sValue(0) = Format(dtInput.Value, "YYYY-MM-DD")
    
        spdView.Row = i
        spdView.Col = 1: sValue(1) = Mid(spdView.Text, 2, 3)
        spdView.Col = 2: sValue(2) = spdView.Value
        spdView.Col = 5: sValue(3) = spdView.Text
        spdView.Col = 6: sValue(4) = spdView.Text
        spdView.Col = 7: sValue(5) = spdView.Text
        
        Call ExecPro("SP_05006_01", sValue(), Err_Num, Err_Dec)
    Next i
    
    If Err_Num = 0 Then
        MsgBox "해당내역이 정상적으로 저장이 되었습니다.", vbInformation
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
'    Dim i As Integer
'    For i = 0 To 30
'        P_00000.crPrint.Formulas(i) = ""
'    Next
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'    P_00000.crPrint.Formulas(0) = "일자1 = '" & Format(dtInput.Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(1) = "대리점 = '" & cboInput.Text & "'"
'    P_00000.crPrint.Formulas(2) = "합계수량 = '" & spdView.MaxRows & "'"
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
'    Dim i As Integer
'    For i = 0 To 30
'        P_00000.crPrint.Formulas(i) = ""
'    Next
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'    P_00000.crPrint.Formulas(0) = "일자1 = '" & Format(dtInput.Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(1) = "대리점 = '" & cboInput.Text & "'"
'    P_00000.crPrint.Formulas(2) = "합계수량 = '" & spdView.MaxRows & "'"
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
        
        spdView.Col = 1
        TempText = LeftH(spdView.Text & Space(16), 16)
        spdView.Col = 2
        TempText = TempText & LeftH(spdView.Text & Space(6), 6)
        spdView.Col = 3
        TempText = TempText & LeftH(spdView.Text & Space(19), 19)
        spdView.Col = 4
        TempText = TempText & LeftH(spdView.Text & Space(6), 6)
        spdView.Col = 5
        TempText = TempText & LeftH(spdView.Text & Space(8), 8)
        spdView.Col = 6
        TempText = TempText & LeftH(spdView.Text & Space(24), 24)
        spdView.Col = 7
        TempText = TempText & LeftH(spdView.Text & Space(24), 24)
        
        Print #1, TempText
    Next i
    
    Close #1
End Sub

