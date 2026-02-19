VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "SSA3D30.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form P_02010 
   Caption         =   "TAG번호 CHECK"
   ClientHeight    =   7860
   ClientLeft      =   1275
   ClientTop       =   1920
   ClientWidth     =   6870
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
   MDIChild        =   -1  'True
   ScaleHeight     =   7860
   ScaleWidth      =   6870
   WindowState     =   2  '최대화
   Begin Threed.SSPanel panMain 
      Align           =   1  '위 맞춤
      Height          =   9555
      Left            =   0
      TabIndex        =   3
      Top             =   435
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   16854
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   9375
         Left            =   60
         TabIndex        =   4
         Top             =   60
         Width           =   15195
         _Version        =   196608
         _ExtentX        =   26802
         _ExtentY        =   16536
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "P_02010.frx":0000
      End
   End
   Begin Threed.SSPanel panInput 
      Align           =   1  '위 맞춤
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   767
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin MSComCtl2.DTPicker dtInput 
         Height          =   315
         Index           =   0
         Left            =   1680
         TabIndex        =   1
         Top             =   60
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   24444928
         CurrentDate     =   36686
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   2
         Left            =   60
         TabIndex        =   2
         Top             =   60
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "입 고 일 자"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
End
Attribute VB_Name = "P_02010"
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
    P_00000.cmdBtn(0).Enabled = True
    P_00000.cmdBtn(5).Enabled = True
    P_00000.cmdBtn(6).Enabled = True
    
    P_00000.panProgramID = Me.Name
    P_00000.panProgramName = Me.Caption
End Sub

Private Sub Form_Load()
    dtInput(0).Value = Date

    ReDim sValue(2)
    
    sValue(0) = "1"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("PRO_P_02010_00", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    Call spdDisplay(RS01)
    Call GetColWidth("백상", Me.Name & "A", spdView)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_00000.cmdBtn(0).Enabled = False
    P_00000.cmdBtn(1).Enabled = False
    P_00000.cmdBtn(2).Enabled = False
    P_00000.cmdBtn(3).Enabled = False
    P_00000.cmdBtn(4).Enabled = False
    P_00000.cmdBtn(5).Enabled = False
    P_00000.cmdBtn(6).Enabled = False

    Call SaveColWidth("백상", Me.Name & "A", spdView)

    P_00000.panProgramID = ""
    P_00000.panProgramName = ""
End Sub

Public Sub DataDisplay()
    Dim i As Integer
    Dim lAmt As Long
    
    ReDim sValue(2)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput(0).Value, "yyyymmdd")
        
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("PRO_P_02010_00", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    Call spdDisplay(RS01)
    Call GetColWidth("백상", Me.Name, spdView)
End Sub

Private Sub spdDisplay(RS As ADODB.Recordset)
    Set spdView.DataSource = RS
    
    spdView.ColsFrozen = 1 '틀고정
    
    spdView.Row = -1
    
    spdView.Col = 1
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft
    
    spdView.Col = 2
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter

    spdView.Col = 3
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter
End Sub

Public Sub DataPrint()
    Dim ReportFP As String
    Dim ReportFile As String
    
    ReportFP = GetIniStr("REPORT", "FilePath", "", sIniFile)
    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
    
    Call PrintDesc
    
    P_00000.crPrint.WindowTitle = Me.Caption
    P_00000.crPrint.Formulas(0) = "입고일자 = '" & Format(dtInput(0).Value, "yyyy-mm-dd") & "'"
    
    Call ReportPrint(ReportFile, "1")
End Sub

Private Sub spdView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu P_00000.PopUp
    End If
End Sub

Public Sub DataScreen()
    Dim ReportFP As String
    Dim ReportFile As String
    
    ReportFP = GetIniStr("REPORT", "FilePath", "", sIniFile)
    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
    
    Call PrintDesc
    
    P_00000.crPrint.WindowTitle = Me.Caption
    P_00000.crPrint.Formulas(0) = "입고일자 = '" & Format(dtInput(0).Value, "yyyy-mm-dd") & "'"
    
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
        
        spdView.Col = 2
        If Mid(spdView.Text, 1, 10) = Format(dtInput(0).Value, "yyyy-mm-dd") Then
            spdView.Col = 1
            TempText = TempText & "    " & LeftH(spdView.Text & Space(20), 20)
        Else
            spdView.Col = 1
            TempText = TempText & "   *" & LeftH(spdView.Text & Space(20), 20)
        End If
        
        spdView.Col = 2
        TempText = TempText & "   " & spdView.Text
        spdView.Col = 3
        TempText = TempText & "   " & spdView.Text
        
        If i Mod 2 = 0 Then
            Print #1, TempText
            TempText = ""
        End If
    Next i
    
    Close #1
End Sub

