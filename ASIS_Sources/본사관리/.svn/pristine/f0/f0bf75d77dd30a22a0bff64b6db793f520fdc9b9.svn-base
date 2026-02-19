VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_04018 
   Caption         =   "월별 매출현황 (그래프)"
   ClientHeight    =   9360
   ClientLeft      =   2175
   ClientTop       =   2370
   ClientWidth     =   15945
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_04018.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9360
   ScaleWidth      =   15945
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   9360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15945
      _ExtentX        =   28125
      _ExtentY        =   16510
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_04018.frx":058A
      Begin MSChart20Lib.MSChart chrView 
         Height          =   8010
         Left            =   7005
         OleObjectBlob   =   "P_04018.frx":063C
         TabIndex        =   1
         Top             =   1335
         Width           =   8925
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   8010
         Left            =   15
         TabIndex        =   2
         Top             =   1335
         Width           =   6975
         _Version        =   524288
         _ExtentX        =   12303
         _ExtentY        =   14129
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
         SpreadDesigner  =   "P_04018.frx":25B4
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   3
         Top             =   540
         Width           =   15915
         _ExtentX        =   28072
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Left            =   8520
            Style           =   2  '드롭다운 목록
            TabIndex        =   4
            Top             =   60
            Width           =   2475
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Left            =   1680
            TabIndex        =   5
            Top             =   60
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy"
            Format          =   56164355
            UpDown          =   -1  'True
            CurrentDate     =   37140
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   6
            Top             =   60
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "조 회 년 도"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   2640
            TabIndex        =   7
            Top             =   60
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "그래프형태"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   6900
            TabIndex        =   8
            Top             =   60
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "그래프종류"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   3
            Left            =   4260
            TabIndex        =   9
            Top             =   60
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   262144
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   1
               Left            =   1440
               TabIndex        =   10
               Top             =   30
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "3차원"
            End
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   11
               Top             =   30
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "2차원"
               Value           =   -1
            End
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   12
         Top             =   15
         Width           =   8310
         _ExtentX        =   14658
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
         PictureBackground=   "P_04018.frx":2A0A
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   1
         Left            =   8340
         TabIndex        =   13
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
         PictureBackground=   "P_04018.frx":2C0C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   14
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
            Picture         =   "P_04018.frx":2E0E
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   15
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
            Picture         =   "P_04018.frx":33A8
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   16
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
            Picture         =   "P_04018.frx":3942
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   17
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
            Picture         =   "P_04018.frx":3EDC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   18
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
            Picture         =   "P_04018.frx":4476
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   19
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
            Picture         =   "P_04018.frx":4A10
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   20
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
            Picture         =   "P_04018.frx":4FAA
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   21
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
            Picture         =   "P_04018.frx":5544
         End
      End
   End
End
Attribute VB_Name = "P_04018"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cboInput_Click()
    Select Case cboInput.ListIndex
        Case 0
            If optSelect(0).Value = True Then
                chrView.chartType = VtChChartType2dBar
            ElseIf optSelect(1).Value = True Then
                chrView.chartType = VtChChartType3dBar
            End If
        Case 1
            If optSelect(0).Value = True Then
                chrView.chartType = VtChChartType2dLine
            ElseIf optSelect(1).Value = True Then
                chrView.chartType = VtChChartType3dLine
            End If
        Case 2
            If optSelect(0).Value = True Then
                chrView.chartType = VtChChartType2dArea
            ElseIf optSelect(1).Value = True Then
                chrView.chartType = VtChChartType3dArea
            End If
        Case 3
            If optSelect(0).Value = True Then
                chrView.chartType = VtChChartType2dStep
            ElseIf optSelect(1).Value = True Then
                chrView.chartType = VtChChartType3dStep
            End If
        Case 4
            If optSelect(0).Value = True Then
                chrView.chartType = VtChChartType2dCombination
            ElseIf optSelect(1).Value = True Then
                chrView.chartType = VtChChartType3dCombination
            End If
        Case 5
            chrView.chartType = VtChChartType2dPie
        Case 6
            chrView.chartType = VtChChartType2dXY
    End Select
    
    If cboInput.ListIndex = 5 Then
        
    Else
        chrView.Column = 1
        chrView.ColumnLabel = Val(Format(dtInput.Value, "yyyy")) - 1
        chrView.Column = 2
        chrView.ColumnLabel = Val(Format(dtInput.Value, "yyyy"))
    End If
    
    chrView.Refresh
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: Call Data_Display   ' 조회
        Case 1: 'Call DataAdd        ' 신규
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

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
'    cmdBtn(5).Enabled = True
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    If P_04018_Flag = False Then
        dtInput.Value = Date
        
        Call ChartInit
    
        ReDim sValue(1)
        
        sValue(0) = "1"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_04018_00", sValue(), Err_Num, Err_Dec)
        
        spdView.MaxCols = RS01.Fields.Count
        spdView.MaxRows = RS01.RecordCount
        
        Call spdDisplay(RS01)
        Call GetColWidth(REG_App, Me.Name, spdView)
        
        P_04018_Flag = True
    End If

    Call optSelect_Click(0, True)
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_04018_Flag = False
End Sub

Public Sub Data_Display()
    Dim i, ii As Integer
    Dim lTotal(3) As Long
    
    ReDim sValue(1)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput.Value, "yyyy")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_04018_00", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    Call spdDisplay(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView)

    Do While Not RS01.EOF
        chrView.Row = Val(RS01!월)
        chrView.Column = 1
        chrView.Data = RS01!전년금액 / 1000
        lTotal(0) = lTotal(0) + Val(RS01!전년금액)
        chrView.Column = 2
        chrView.Data = RS01!금년금액 / 1000
        lTotal(1) = lTotal(1) + Val(RS01!금년금액)
        lTotal(2) = lTotal(2) + Val(RS01!전년단가)
        lTotal(3) = lTotal(3) + Val(RS01!금년단가)
        
        If RS01!금년단가 <> 0 Then
            ii = ii + 1
        End If
        
        RS01.MoveNext
    Loop
    
    spdView.MaxRows = spdView.MaxRows + 1
    spdView.Row = spdView.MaxRows
    
    spdView.Col = 1
    spdView.Text = "합  계"
    spdView.Col = 2
    spdView.Text = lTotal(0)
    spdView.Col = 3
    spdView.Text = lTotal(1)
    spdView.Col = 4
    spdView.Text = lTotal(2) / (spdView.MaxRows - 1)
    spdView.Col = 5
    spdView.Text = lTotal(3) / ii
End Sub

Private Sub ChartInit()
    Dim i As Integer
    Dim iDay As Integer
    
    For i = 1 To 12
        chrView.Row = i
        chrView.RowLabel = i
        chrView.Column = 1
        chrView.Data = 0
        chrView.Column = 2
        chrView.Data = 0
    Next i
End Sub

Private Sub spdDisplay(Rs As ADODB.Recordset)
    
    Call fpSpread_Display(spdView, Rs)
    
    spdView.ColsFrozen = 1 '틀고정
    
    spdView.Row = -1
    
    spdView.Col = 1
    spdView.ColWidth(1) = 6
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
    spdView.ColWidth(4) = 8
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 5
    spdView.ColWidth(5) = 8
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
End Sub

Private Sub optSelect_Click(Index As Integer, Value As Integer)
    Select Case Index
        Case 0
            cboInput.Clear
            
            cboInput.AddItem "막대형 / 그림 그래프"
            cboInput.AddItem "꺽은선형"
            cboInput.AddItem "영역형"
            cboInput.AddItem "단계"
            cboInput.AddItem "혼합형"
            cboInput.AddItem "원형"
            cboInput.AddItem "XY(분산형)"
        Case 1
            cboInput.Clear
            
            cboInput.AddItem "막대형(열)"
            cboInput.AddItem "꺽은선형(테입프)"
            cboInput.AddItem "영역형"
            cboInput.AddItem "단계"
            cboInput.AddItem "혼합형"
    End Select
    
    cboInput.ListIndex = 0
End Sub
