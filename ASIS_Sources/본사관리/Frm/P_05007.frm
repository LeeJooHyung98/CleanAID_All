VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMask32.ocx"
Begin VB.Form P_05007 
   Caption         =   "재세탁 관리 현황"
   ClientHeight    =   12465
   ClientLeft      =   1545
   ClientTop       =   3585
   ClientWidth     =   17490
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_05007.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12465
   ScaleWidth      =   17490
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   12465
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17490
      _ExtentX        =   30850
      _ExtentY        =   21987
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_05007.frx":058A
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   11220
         Left            =   15
         TabIndex        =   1
         Top             =   1230
         Width           =   17460
         _Version        =   524288
         _ExtentX        =   30797
         _ExtentY        =   19791
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
         SpreadDesigner  =   "P_05007.frx":05FC
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   2
         Top             =   435
         Width           =   17460
         _ExtentX        =   30798
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   1
            Left            =   9570
            Style           =   2  '드롭다운 목록
            TabIndex        =   4
            Top             =   60
            Width           =   1695
         End
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            Style           =   2  '드롭다운 목록
            TabIndex        =   3
            Top             =   420
            Width           =   3015
         End
         Begin MSMask.MaskEdBox mskInput 
            Height          =   315
            Left            =   6690
            TabIndex        =   5
            Top             =   420
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "#-###"
            PromptChar      =   "_"
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   4770
            TabIndex        =   6
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   63307776
            CurrentDate     =   36686
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            TabIndex        =   7
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   63307776
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   8
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
            TabIndex        =   9
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "검 색 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   7
            Left            =   8100
            TabIndex        =   10
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "구    분"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   9
            Left            =   5220
            TabIndex        =   11
            Top             =   420
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "택  번  호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   8
            Left            =   8100
            TabIndex        =   12
            Top             =   420
            Width           =   3180
            _ExtentX        =   5609
            _ExtentY        =   556
            _Version        =   262144
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   13
               Top             =   30
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "입고일자"
               Value           =   -1
            End
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   1
               Left            =   1800
               TabIndex        =   14
               Top             =   30
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "출고일자"
            End
         End
         Begin VB.Label Label1 
            Alignment       =   2  '가운데 맞춤
            Caption         =   "~"
            Height          =   195
            Left            =   4530
            TabIndex        =   15
            Top             =   120
            Width           =   195
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   405
         Left            =   15
         TabIndex        =   16
         Top             =   15
         Width           =   17460
         _ExtentX        =   30798
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
         Caption         =   " 재세탁 관리 현황 (P_05007)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_05007.frx":0A75
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
End
Attribute VB_Name = "P_05007"
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
'    cmdBtn(5).Enabled = True
'    cmdBtn(6).Enabled = True
'
'    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    If P_05007_Flag = False Then
        Call ComboAdd
        
        dtInput(0).Value = Date
        dtInput(1).Value = Date
        
        ReDim sValue(6)
        
        sValue(0) = "1"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_05007_00", sValue(), Err_Num, Err_Dec)
        
        spdView.MaxCols = RS01.Fields.Count
        spdView.MaxRows = RS01.RecordCount
        
        Call spdDisplay(RS01)
        Call GetColWidth(REG_App, Me.Name, spdView)
        
        P_05007_Flag = True
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
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter
    
    spdView.Col = 3
    spdView.ColWidth(3) = 10
    spdView.CellType = CellTypeDate
    spdView.TypeDateCentury = True
    spdView.TypeDateFormat = TypeDateFormatYYMMDD
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter
    
    spdView.Col = 4
    spdView.ColWidth(4) = 10
    spdView.CellType = CellTypeDate
    spdView.TypeDateCentury = True
    spdView.TypeDateFormat = TypeDateFormatYYMMDD
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter
    
    spdView.Col = 5
    spdView.ColWidth(5) = 20
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft
    
    spdView.Col = 6
    spdView.ColWidth(6) = 8
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter

    spdView.Col = 7
    spdView.ColWidth(7) = 8
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter

    spdView.Col = 8
    spdView.ColWidth(8) = 16
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft

    spdView.Col = 9
    spdView.ColWidth(9) = 16
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_05007_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i As Integer
    Dim iTotal(3) As Long

    ReDim sValue(6)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    sValue(3) = cboInput(1).Text & "%"
    sValue(4) = Mid(cboInput(0).Text, 2, 3) & "%"
    sValue(5) = mskInput.ClipText & "%"
    If optSelect(0).Value = True Then
        sValue(6) = "0"
    ElseIf optSelect(1).Value = True Then
        sValue(6) = "1"
    End If
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_05007_00", sValue(), Err_Num, Err_Dec)
    
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
'    If optSelect(0).Value = True Then
'        P_00000.crPrint.Formulas(3) = "검색구분 = '입고일자'"
'    Else
'        P_00000.crPrint.Formulas(3) = "검색구분 = '출고일자'"
'    End If
'
'    Call ReportPrint(ReportFile, "1")
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

    cboInput(1).Clear
    
    cboInput(1).AddItem ""
    cboInput(1).AddItem "오염"
    cboInput(1).AddItem "이염"
    cboInput(1).AddItem "수축"
    cboInput(1).AddItem "늘어짐"
    cboInput(1).AddItem "경화"
    cboInput(1).AddItem "파손"
    cboInput(1).AddItem "탈색"
    cboInput(1).AddItem "기타"
End Sub
