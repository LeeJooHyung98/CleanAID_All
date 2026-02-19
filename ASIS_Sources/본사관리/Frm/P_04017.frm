VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_04017 
   Caption         =   "주간 매출현황 (그래프)"
   ClientHeight    =   11715
   ClientLeft      =   585
   ClientTop       =   2070
   ClientWidth     =   16335
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_04017.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11715
   ScaleWidth      =   16335
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   11715
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16335
      _ExtentX        =   28813
      _ExtentY        =   20664
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_04017.frx":058A
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10365
         Left            =   15
         TabIndex        =   1
         Top             =   1335
         Width           =   7740
         _Version        =   524288
         _ExtentX        =   13652
         _ExtentY        =   18283
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
         SpreadDesigner  =   "P_04017.frx":063C
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin MSChart20Lib.MSChart chrView 
         Height          =   10365
         Left            =   7770
         OleObjectBlob   =   "P_04017.frx":0A92
         TabIndex        =   2
         Top             =   1335
         Width           =   8550
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   3
         Top             =   540
         Width           =   16305
         _ExtentX        =   28760
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Left            =   12960
            Style           =   2  '드롭다운 목록
            TabIndex        =   4
            Top             =   60
            Width           =   2295
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
            Format          =   56557571
            UpDown          =   -1  'True
            CurrentDate     =   37140
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   4200
            TabIndex        =   6
            Top             =   60
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            _Version        =   262144
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   0
               Left            =   180
               TabIndex        =   7
               Top             =   30
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "상 반 기"
               Value           =   -1
            End
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   1
               Left            =   1620
               TabIndex        =   8
               Top             =   30
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "하 반 기"
            End
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   9
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
            Left            =   2580
            TabIndex        =   10
            Top             =   60
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "구    분"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   3
            Left            =   7140
            TabIndex        =   11
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
            Index           =   4
            Left            =   11340
            TabIndex        =   12
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
            Index           =   5
            Left            =   8760
            TabIndex        =   13
            Top             =   60
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   262144
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   3
               Left            =   1440
               TabIndex        =   14
               Top             =   30
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "3차원"
            End
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   15
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
         TabIndex        =   16
         Top             =   15
         Width           =   8700
         _ExtentX        =   15346
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
         PictureBackground=   "P_04017.frx":2DAC
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   1
         Left            =   8730
         TabIndex        =   17
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
         PictureBackground=   "P_04017.frx":2FAE
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   18
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
            Picture         =   "P_04017.frx":31B0
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   19
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
            Picture         =   "P_04017.frx":374A
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   20
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
            Picture         =   "P_04017.frx":3CE4
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   21
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
            Picture         =   "P_04017.frx":427E
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   22
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
            Picture         =   "P_04017.frx":4818
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   23
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
            Picture         =   "P_04017.frx":4DB2
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   24
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
            Picture         =   "P_04017.frx":534C
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   25
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
            Picture         =   "P_04017.frx":58E6
         End
      End
   End
End
Attribute VB_Name = "P_04017"
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
    
    If P_04017_Flag = False Then
        dtInput.Value = Date
        
        Call ChartInit
    
        ReDim sValue(2)
        
        sValue(0) = "1"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_04017_00", sValue(), Err_Num, Err_Dec)
        
        spdView.MaxCols = RS01.Fields.Count
        spdView.MaxRows = RS01.RecordCount
        
        Call spdDisplay(RS01)
        Call GetColWidth(REG_App, Me.Name, spdView)
        
        P_04017_Flag = True
    End If

    Call optSelect_Click(2, True)
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_04017_Flag = False
End Sub

Public Sub Data_Display()
    Dim i As Integer
    
    ReDim sValue(2)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput.Value, "yyyy")
    If optSelect(0).Value = True Then
        sValue(2) = "1"
    ElseIf optSelect(1).Value = True Then
        sValue(2) = "2"
    End If
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_04017_00", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    Call spdDisplay(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView)
    
    Do While Not RS01.EOF
        If optSelect(0).Value = True Then
            chrView.Row = Val(RS01!주)
        ElseIf optSelect(1).Value = True Then
            chrView.Row = Val(RS01!주) - 27
        End If
        
        chrView.Column = 1
        chrView.Data = RS01!전년금액 / 1000
        chrView.Column = 2
        chrView.Data = RS01!금년금액 / 1000
        
        RS01.MoveNext
    Loop
End Sub

Private Sub ChartInit()
    Dim i As Integer
    Dim iDay As Integer
    
    If optSelect(0).Value = True Then
        For i = 1 To 27
            chrView.Row = i
            chrView.RowLabel = i
            chrView.Column = 1
            chrView.Data = 0
            chrView.Column = 2
            chrView.Data = 0
        Next i
    ElseIf optSelect(1).Value = True Then
        For i = 28 To 54
            chrView.Row = i
            chrView.RowLabel = i
            chrView.Column = 1
            chrView.Data = 0
            chrView.Column = 2
            chrView.Data = 0
        Next i
    End If
End Sub

Private Sub cboInput_Click()
    Select Case cboInput.ListIndex
        Case 0
            If optSelect(2).Value = True Then
                chrView.chartType = VtChChartType2dBar
            ElseIf optSelect(3).Value = True Then
                chrView.chartType = VtChChartType3dBar
            End If
        Case 1
            If optSelect(2).Value = True Then
                chrView.chartType = VtChChartType2dLine
            ElseIf optSelect(3).Value = True Then
                chrView.chartType = VtChChartType3dLine
            End If
        Case 2
            If optSelect(2).Value = True Then
                chrView.chartType = VtChChartType2dArea
            ElseIf optSelect(3).Value = True Then
                chrView.chartType = VtChChartType3dArea
            End If
        Case 3
            If optSelect(2).Value = True Then
                chrView.chartType = VtChChartType2dStep
            ElseIf optSelect(3).Value = True Then
                chrView.chartType = VtChChartType3dStep
            End If
        Case 4
            If optSelect(2).Value = True Then
                chrView.chartType = VtChChartType2dCombination
            ElseIf optSelect(3).Value = True Then
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

Private Sub optSelect_Click(Index As Integer, Value As Integer)
    Select Case Index
        Case 2
            cboInput.Clear
            
            cboInput.AddItem "막대형 / 그림 그래프"
            cboInput.AddItem "꺽은선형"
            cboInput.AddItem "영역형"
            cboInput.AddItem "단계"
            cboInput.AddItem "혼합형"
            cboInput.AddItem "원형"
            cboInput.AddItem "XY(분산형)"
        
            cboInput.ListIndex = 0
        Case 3
            cboInput.Clear
            
            cboInput.AddItem "막대형(열)"
            cboInput.AddItem "꺽은선형(테입프)"
            cboInput.AddItem "영역형"
            cboInput.AddItem "단계"
            cboInput.AddItem "혼합형"
    
            cboInput.ListIndex = 0
    End Select
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

