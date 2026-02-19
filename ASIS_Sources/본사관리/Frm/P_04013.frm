VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_04013 
   Caption         =   "기간별 매출현황"
   ClientHeight    =   9255
   ClientLeft      =   5445
   ClientTop       =   6600
   ClientWidth     =   16110
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_04013.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   16110
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   9255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16110
      _ExtentX        =   28416
      _ExtentY        =   16325
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_04013.frx":058A
      Begin Threed.SSPanel SSPanel 
         Height          =   1590
         Index           =   0
         Left            =   15
         TabIndex        =   2
         Top             =   7650
         Width           =   16080
         _ExtentX        =   28363
         _ExtentY        =   2805
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   0
            Left            =   1665
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   345
            Width           =   1635
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   1
            Left            =   3285
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   345
            Width           =   1635
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   2
            Left            =   4905
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   345
            Width           =   1635
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   4
            Left            =   1665
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   645
            Width           =   1635
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   5
            Left            =   3285
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   645
            Width           =   1635
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   6
            Left            =   4905
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   645
            Width           =   1635
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   8
            Left            =   1665
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   945
            Width           =   1635
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   9
            Left            =   3285
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   945
            Width           =   1635
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   10
            Left            =   4905
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   945
            Width           =   1635
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   3
            Left            =   6525
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   345
            Width           =   1635
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   7
            Left            =   6525
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   645
            Width           =   1635
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   11
            Left            =   6525
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   945
            Width           =   1635
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   12
            Left            =   1665
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   1245
            Width           =   1635
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   13
            Left            =   3285
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   1245
            Width           =   1635
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   14
            Left            =   4905
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   1245
            Width           =   1635
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   15
            Left            =   6525
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   1245
            Width           =   1635
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   10
            Left            =   45
            TabIndex        =   7
            Top             =   1245
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "총  합  계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   9
            Left            =   6525
            TabIndex        =   13
            Top             =   45
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "점  유  율"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   8
            Left            =   45
            TabIndex        =   21
            Top             =   945
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "할 인 점 계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   7
            Left            =   45
            TabIndex        =   22
            Top             =   645
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "백 화 점 계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   6
            Left            =   45
            TabIndex        =   23
            Top             =   345
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "대 리 점 계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   5
            Left            =   4905
            TabIndex        =   24
            Top             =   45
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "입 고 단 가"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   4
            Left            =   3285
            TabIndex        =   25
            Top             =   45
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "입 고 금 액"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   3
            Left            =   1665
            TabIndex        =   26
            Top             =   45
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "입 고 수 량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   6300
         Left            =   15
         TabIndex        =   1
         Top             =   1335
         Width           =   16080
         _Version        =   524288
         _ExtentX        =   28363
         _ExtentY        =   11112
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
         SpreadDesigner  =   "P_04013.frx":063C
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   27
         Top             =   540
         Width           =   16080
         _ExtentX        =   28363
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   6480
            TabIndex        =   28
            Top             =   60
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            _Version        =   262144
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   29
               Top             =   30
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "일 일"
               Value           =   -1
            End
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   1
               Left            =   1440
               TabIndex        =   30
               Top             =   30
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "월 간"
            End
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   2
               Left            =   2700
               TabIndex        =   31
               Top             =   30
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "연 간"
            End
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Left            =   1680
            TabIndex        =   32
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   64159744
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   33
            Top             =   60
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "기 준 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   4860
            TabIndex        =   34
            Top             =   60
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "구    분"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   35
         Top             =   15
         Width           =   8475
         _ExtentX        =   14949
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
         PictureBackground=   "P_04013.frx":0AB5
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   1
         Left            =   8505
         TabIndex        =   36
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
         PictureBackground=   "P_04013.frx":0CB7
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   37
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
            Picture         =   "P_04013.frx":0EB9
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   38
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
            Picture         =   "P_04013.frx":1453
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   39
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
            Picture         =   "P_04013.frx":19ED
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   40
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
            Picture         =   "P_04013.frx":1F87
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   41
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
            Picture         =   "P_04013.frx":2521
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   42
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
            Picture         =   "P_04013.frx":2ABB
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   43
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
            Picture         =   "P_04013.frx":3055
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   44
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
            Picture         =   "P_04013.frx":35EF
         End
      End
   End
End
Attribute VB_Name = "P_04013"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As New ADODB.Recordset
Dim strSql As String
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

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(5).Enabled = True
    cmdBtn(6).Enabled = True
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    If P_04013_Flag = False Then
        dtInput.Value = Date
        
        ReDim sValue(2)
        
        sValue(0) = "1"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_04013_00", sValue(), Err_Num, Err_Dec)
        
        spdView.MaxCols = RS01.Fields.Count
        spdView.MaxRows = RS01.RecordCount
        
        Call spdDisplay
        Call GetColWidth(REG_App, Me.Name, spdView)
        
        P_04013_Flag = True
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub spdDisplay()
    
    spdView.ColsFrozen = 1 '틀고정
    
    spdView.Row = -1
    
    spdView.Col = 1
    spdView.ColWidth(1) = 8
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft

    spdView.Col = 2
    spdView.ColWidth(2) = 10
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft

    spdView.Col = 3
    spdView.ColWidth(3) = 25
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft

    spdView.Col = 4
    spdView.ColWidth(4) = 10
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
    spdView.ColWidth(6) = 12
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 7
    spdView.ColWidth(7) = 6
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 2
    spdView.TypeVAlign = TypeVAlignCenter
    
    spdView.Row = 0
    spdView.Col = 1:        spdView.Text = "순번"
    spdView.Col = 2:        spdView.Text = "구분"
    spdView.Col = 3:        spdView.Text = "매장명"
    spdView.Col = 4:        spdView.Text = "입고수량"
    spdView.Col = 5:        spdView.Text = "금액"
    spdView.Col = 6:        spdView.Text = "단가"
    spdView.Col = 7:        spdView.Text = "점유율"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_04013_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i As Integer
    Dim lSubTotal1(5) As Long
    Dim lSubTotal2(5) As Long
    Dim lSubTotal3(5) As Long
    Dim lSubTotal4(5) As Long
    
    Dim iMemQty As Long
    
    ReDim sValue(3)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput.Value, "YYYY-MM-DD")
    
    If optSelect(0).Value = True Then
        sValue(1) = Format(dtInput.Value, "YYYY-MM-DD")
        sValue(2) = Format(dtInput.Value, "YYYY-MM-DD")
    ElseIf optSelect(1).Value = True Then
        sValue(1) = Format(dtInput.Value, "YYYY-MM-01")
        sValue(2) = Format(dtInput.Value, "YYYY-MM-31")
    ElseIf optSelect(2).Value = True Then
        sValue(1) = Format(dtInput.Value, "YYYY-01-01")
        sValue(2) = Format(dtInput.Value, "YYYY-12-31")
    End If
    
    strSql = ""
    strSql = strSql + "    SELECT "
    strSql = strSql + "        B.AgencySection ,"
    strSql = strSql + "        '[' + A.AgencyCode + '] ' + B.AgencyName            '매장명', "
    strSql = strSql + "        Sum(A.IpSu)                                         '입고수량', "
    strSql = strSql + "        Sum(A.Amount)                                       '금액',"
    strSql = strSql + "        CASE WHEN Sum(A.IpSu) <= 0 AND Sum(A.Amount) <= 0 THEN '0' ELSE Sum(A.Amount) / Sum(A.IpSu) END '단가', "
    strSql = strSql + "        0                                               '점유율' "

    strSql = strSql + "    FROM    Sugeum      A (NOLOCK) "
    strSql = strSql + "         INNER JOIN AgencyCT AS B(NOLOCK) "
    strSql = strSql + "         ON  A.AgencyCode    =  B.AgencyCode"
    strSql = strSql + "    WHERE     A.SuDate    >=   '" & sValue(1) & "'    AND A.SuDate    <= '" & sValue(2) & "'"
    strSql = strSql + "    AND A.IpSu      <>  0"
    strSql = strSql + "    GROUP BY    A.AgencyCode, B.AgencyName, B.AgencySection"

    Set RS01 = New ADODB.Recordset
    Call SqlDataValue(RS01, strSql)

'    Set RS01 = ExecPro("SP_04013_00", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count + 1
    spdView.MaxRows = RS01.RecordCount
    
    Call spdDisplay
    Call GetColWidth(REG_App, Me.Name, spdView)
    
    If RS01.RecordCount <= 0 Then Exit Sub
    
    spdView.AutoCalc = True
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        spdView.Col = 2
        If spdView.Text = "대리점" Then
            spdView.Col = 4: lSubTotal1(0) = lSubTotal1(0) + spdView.Value
            spdView.Col = 5: lSubTotal1(1) = lSubTotal1(1) + spdView.Value
            spdView.Col = 6: lSubTotal1(2) = lSubTotal1(2) + spdView.Value
            
        ElseIf spdView.Text = "백화점" Then
            spdView.Col = 4: lSubTotal2(0) = lSubTotal2(0) + spdView.Value
            spdView.Col = 5: lSubTotal2(1) = lSubTotal2(1) + spdView.Value
            spdView.Col = 6: lSubTotal2(2) = lSubTotal2(2) + spdView.Value
            
        ElseIf spdView.Text = "할인매장" Then
            spdView.Col = 4: lSubTotal3(0) = lSubTotal3(0) + spdView.Value
            spdView.Col = 5: lSubTotal3(1) = lSubTotal3(1) + spdView.Value
            spdView.Col = 6: lSubTotal3(2) = lSubTotal3(2) + spdView.Value
        End If
        
        spdView.Col = 4: lSubTotal4(0) = lSubTotal4(0) + spdView.Value
        spdView.Col = 5: lSubTotal4(1) = lSubTotal4(1) + spdView.Value
        spdView.Col = 6: lSubTotal4(2) = lSubTotal4(2) + spdView.Value
    Next i
    
    txtInput(0).Text = Format(lSubTotal1(0), "#,##0")
    txtInput(1).Text = Format(lSubTotal1(1), "#,##0")
    If lSubTotal1(0) <> 0 And lSubTotal1(1) <> 0 Then
        txtInput(2).Text = Format(lSubTotal1(1) / lSubTotal1(0), "#,##0")
    Else
        txtInput(2).Text = "0"
    End If

    txtInput(4).Text = Format(lSubTotal2(0), "#,##0")
    txtInput(5).Text = Format(lSubTotal2(1), "#,##0")
    If lSubTotal2(0) <> 0 And lSubTotal2(1) <> 0 Then
        txtInput(6).Text = Format(lSubTotal2(1) / lSubTotal2(0), "#,##0")
    Else
        txtInput(6).Text = "0"
    End If

    txtInput(8).Text = Format(lSubTotal3(0), "#,##0")
    txtInput(9).Text = Format(lSubTotal3(1), "#,##0")
    If lSubTotal3(0) <> 0 And lSubTotal3(1) <> 0 Then
        txtInput(10).Text = Format(lSubTotal3(1) / lSubTotal3(0), "#,##0")
    Else
        txtInput(10).Text = "0"
    End If
    
    txtInput(12).Text = Format(lSubTotal4(0), "#,##0")
    txtInput(13).Text = Format(lSubTotal4(1), "#,##0")
    If lSubTotal4(0) <> 0 And lSubTotal4(1) <> 0 Then
        txtInput(14).Text = Format(lSubTotal4(1) / lSubTotal4(0), "#,##0")
    Else
        txtInput(14).Text = "0"
    End If
    
    If lSubTotal1(0) <> 0 And lSubTotal1(1) <> 0 Then
        txtInput(3).Text = Format((lSubTotal1(0) / lSubTotal4(0)) * 100, "#,##0.00")
    Else
        txtInput(3).Text = "0"
    End If
    
    If lSubTotal2(0) <> 0 And lSubTotal2(1) <> 0 Then
        txtInput(7).Text = Format((lSubTotal2(0) / lSubTotal4(0)) * 100, "#,##0.00")
    Else
        txtInput(7).Text = "0"
    End If
    
    If lSubTotal3(0) <> 0 And lSubTotal3(1) <> 0 Then
        txtInput(11).Text = Format((lSubTotal3(0) / lSubTotal4(0)) * 100, "#,##0.00")
    Else
        txtInput(11).Text = "0"
    End If
    
    txtInput(15).Text = "100.00"
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        
        spdView.Col = 5
        iMemQty = spdView.Value
        
        spdView.Col = 7
        spdView.Text = (iMemQty / lSubTotal4(1)) * 100
    Next i
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub spdView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        'PopupMenu P_00000.PopUp
    End If
End Sub

Public Sub DataPrint()
'    Dim ReportFP As String
'    Dim ReportFile As String
'    Dim sData As String
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
'    If optSelect(0).Value = True Then
'        P_00000.crPrint.Formulas(0) = "기준일자 = '" & Format(dtInput.Value, "YYYY-MM-DD") & "'"
'    ElseIf optSelect(1).Value = True Then
'        P_00000.crPrint.Formulas(0) = "기준일자 = '" & Format(dtInput.Value, "yyyy-mm") & "'"
'    ElseIf optSelect(2).Value = True Then
'        P_00000.crPrint.Formulas(0) = "기준일자 = '" & Format(dtInput.Value, "yyyy") & "'"
'    End If
'
'    sData = Space(11) & RightH(Space(6) & txtInput(0).Text, 6) & Space(1)
'    sData = sData & RightH(Space(12) & txtInput(1).Text, 12) & Space(1)
'    sData = sData & RightH(Space(12) & txtInput(2).Text, 12) & Space(1)
'    sData = sData & RightH(Space(6) & txtInput(3).Text, 6) & Space(1)
'
'    P_00000.crPrint.Formulas(1) = "대리점계 = '" & sData & "'"
'
'    sData = Space(11) & RightH(Space(6) & txtInput(4).Text, 6) & Space(1)
'    sData = sData & RightH(Space(12) & txtInput(5).Text, 12) & Space(1)
'    sData = sData & RightH(Space(12) & txtInput(6).Text, 12) & Space(1)
'    sData = sData & RightH(Space(6) & txtInput(7).Text, 6) & Space(1)
'
'    P_00000.crPrint.Formulas(2) = "백화점계 = '" & sData & "'"
'
'    sData = Space(11) & RightH(Space(6) & txtInput(8).Text, 6) & Space(1)
'    sData = sData & RightH(Space(12) & txtInput(9).Text, 12) & Space(1)
'    sData = sData & RightH(Space(12) & txtInput(10).Text, 12) & Space(1)
'    sData = sData & RightH(Space(6) & txtInput(11).Text, 6) & Space(1)
'
'    P_00000.crPrint.Formulas(3) = "할인매장계 = '" & sData & "'"
'
'    sData = Space(11) & RightH(Space(6) & txtInput(12).Text, 6) & Space(1)
'    sData = sData & RightH(Space(12) & txtInput(13).Text, 12) & Space(1)
'    sData = sData & RightH(Space(12) & txtInput(14).Text, 12) & Space(1)
'    sData = sData & RightH(Space(6) & txtInput(15).Text, 6) & Space(1)
'
'    P_00000.crPrint.Formulas(4) = "총계 = '" & sData & "'"
'
'    Call ReportPrint(ReportFile, "1")
End Sub

Public Sub DataScreen()
'    Dim ReportFP As String
'    Dim ReportFile As String
'    Dim sData As String
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
'    If optSelect(0).Value = True Then
'        P_00000.crPrint.Formulas(0) = "기준일자 = '" & Format(dtInput.Value, "YYYY-MM-DD") & "'"
'    ElseIf optSelect(1).Value = True Then
'        P_00000.crPrint.Formulas(0) = "기준일자 = '" & Format(dtInput.Value, "yyyy-mm") & "'"
'    ElseIf optSelect(2).Value = True Then
'        P_00000.crPrint.Formulas(0) = "기준일자 = '" & Format(dtInput.Value, "yyyy") & "'"
'    End If
'
'    sData = Space(11) & RightH(Space(6) & txtInput(0).Text, 6) & Space(1)
'    sData = sData & RightH(Space(12) & txtInput(1).Text, 12) & Space(1)
'    sData = sData & RightH(Space(12) & txtInput(2).Text, 12) & Space(1)
'    sData = sData & RightH(Space(6) & txtInput(3).Text, 6) & Space(1)
'
'    P_00000.crPrint.Formulas(1) = "대리점계 = '" & sData & "'"
'
'    sData = Space(11) & RightH(Space(6) & txtInput(4).Text, 6) & Space(1)
'    sData = sData & RightH(Space(12) & txtInput(5).Text, 12) & Space(1)
'    sData = sData & RightH(Space(12) & txtInput(6).Text, 12) & Space(1)
'    sData = sData & RightH(Space(6) & txtInput(7).Text, 6) & Space(1)
'
'    P_00000.crPrint.Formulas(2) = "백화점계 = '" & sData & "'"
'
'    sData = Space(11) & RightH(Space(6) & txtInput(8).Text, 6) & Space(1)
'    sData = sData & RightH(Space(12) & txtInput(9).Text, 12) & Space(1)
'    sData = sData & RightH(Space(12) & txtInput(10).Text, 12) & Space(1)
'    sData = sData & RightH(Space(6) & txtInput(11).Text, 6) & Space(1)
'
'    P_00000.crPrint.Formulas(3) = "할인매장계 = '" & sData & "'"
'
'    sData = Space(11) & RightH(Space(6) & txtInput(12).Text, 6) & Space(1)
'    sData = sData & RightH(Space(12) & txtInput(13).Text, 12) & Space(1)
'    sData = sData & RightH(Space(12) & txtInput(14).Text, 12) & Space(1)
'    sData = sData & RightH(Space(6) & txtInput(15).Text, 6) & Space(1)
'
'    P_00000.crPrint.Formulas(4) = "총계 = '" & sData & "'"
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
    
    For i = 1 To spdView.MaxRows - 2
        spdView.Row = i
        
        spdView.Col = 2
        TempText = LeftH(spdView.Text & Space(10), 10)
        spdView.Col = 3
        TempText = TempText & LeftH(spdView.Text & Space(20), 20) & Space(1)
        spdView.Col = 4
        TempText = TempText & RightH(Space(6) & spdView.Text, 6) & Space(1)
        spdView.Col = 5
        TempText = TempText & RightH(Space(12) & spdView.Text, 12) & Space(1)
        spdView.Col = 6
        TempText = TempText & RightH(Space(12) & spdView.Text, 12) & Space(1)
        spdView.Col = 7
        TempText = TempText & RightH(Space(6) & spdView.Text, 6) & Space(1)
        
        Print #1, TempText
    Next i
    
    Close #1
End Sub
