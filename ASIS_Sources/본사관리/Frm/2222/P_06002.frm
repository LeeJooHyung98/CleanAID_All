VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_06002 
   Caption         =   "사고처리 내역"
   ClientHeight    =   9540
   ClientLeft      =   585
   ClientTop       =   2070
   ClientWidth     =   15390
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
   ScaleHeight     =   9540
   ScaleWidth      =   15390
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   16828
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_06002.frx":0000
      Begin Threed.SSPanel SSPanel1 
         Height          =   405
         Left            =   15
         TabIndex        =   18
         Top             =   9120
         Width           =   15360
         _ExtentX        =   27093
         _ExtentY        =   714
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
            TabIndex        =   23
            Top             =   45
            Width           =   2115
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   1
            Left            =   5445
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   45
            Width           =   2115
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   2
            Left            =   9225
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   45
            Width           =   2115
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   3
            Left            =   13005
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   45
            Width           =   2115
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   7
            Left            =   45
            TabIndex        =   24
            Top             =   45
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "총  건  수"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   8
            Left            =   3825
            TabIndex        =   25
            Top             =   45
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "제품금액계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   9
            Left            =   7605
            TabIndex        =   26
            Top             =   45
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "보 상 건 수"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   10
            Left            =   11385
            TabIndex        =   27
            Top             =   45
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "보상금액계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel panInput 
         Height          =   2100
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   15360
         _ExtentX        =   27093
         _ExtentY        =   3704
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            Style           =   2  '드롭다운 목록
            TabIndex        =   42
            Top             =   1350
            Width           =   3735
         End
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   5
            Left            =   1560
            Style           =   2  '드롭다운 목록
            TabIndex        =   29
            Top             =   630
            Width           =   3705
         End
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   4
            Left            =   1560
            Style           =   2  '드롭다운 목록
            TabIndex        =   28
            Top             =   990
            Width           =   3705
         End
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   1
            Left            =   6990
            Style           =   2  '드롭다운 목록
            TabIndex        =   5
            Top             =   1350
            Width           =   3735
         End
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   2
            Left            =   1530
            Style           =   2  '드롭다운 목록
            TabIndex        =   4
            Top             =   1710
            Width           =   3735
         End
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   3
            Left            =   6990
            Style           =   2  '드롭다운 목록
            TabIndex        =   3
            Top             =   1710
            Width           =   3735
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   4
            Left            =   12450
            TabIndex        =   2
            Top             =   1710
            Width           =   1155
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   6990
            TabIndex        =   6
            Top             =   630
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   262144
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   0
               Left            =   420
               TabIndex        =   7
               Top             =   30
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "접수일자"
               Value           =   -1
            End
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   1
               Left            =   2160
               TabIndex        =   8
               Top             =   30
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "보상일자"
            End
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   6990
            TabIndex        =   9
            Top             =   990
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   71892992
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   5520
            TabIndex        =   10
            Top             =   990
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "검 색 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   10290
            TabIndex        =   11
            Top             =   990
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   71892992
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   5520
            TabIndex        =   12
            Top             =   630
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "검 색 기 준"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   4
            Left            =   5520
            TabIndex        =   13
            Top             =   1350
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "크 레 임 명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   5
            Left            =   60
            TabIndex        =   14
            Top             =   1710
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "담 당 자 명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   6
            Left            =   5520
            TabIndex        =   15
            Top             =   1710
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "보 상 구 분"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   11
            Left            =   10980
            TabIndex        =   16
            Top             =   1710
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "접 수 번 호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   12
            Left            =   90
            TabIndex        =   30
            Top             =   630
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "사 업 장"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   25
            Left            =   90
            TabIndex        =   31
            Top             =   990
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "가 맹 점"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel pnlHeader 
            Height          =   555
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Width           =   7620
            _ExtentX        =   13441
            _ExtentY        =   979
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
            Caption         =   "사고 처리 접수 (P_06001)"
            PictureBackgroundStyle=   2
            PictureBackground=   "P_06002.frx":0072
            BorderWidth     =   0
            BevelOuter      =   0
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   555
            Index           =   0
            Left            =   7635
            TabIndex        =   33
            Top             =   0
            Width           =   7605
            _ExtentX        =   13414
            _ExtentY        =   979
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
            PictureBackground=   "P_06002.frx":0274
            BorderWidth     =   0
            BevelOuter      =   0
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin XtremeSuiteControls.PushButton cmdBtn 
               Height          =   450
               Index           =   7
               Left            =   6660
               TabIndex        =   34
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
               Picture         =   "P_06002.frx":0476
            End
            Begin XtremeSuiteControls.PushButton cmdBtn 
               Height          =   450
               Index           =   6
               Left            =   5730
               TabIndex        =   35
               Top             =   30
               Width           =   900
               _Version        =   851970
               _ExtentX        =   1587
               _ExtentY        =   794
               _StockProps     =   79
               Caption         =   "화면"
               ForeColor       =   -2147483640
               BackColor       =   -2147483636
               Appearance      =   6
               Picture         =   "P_06002.frx":0A10
            End
            Begin XtremeSuiteControls.PushButton cmdBtn 
               Height          =   450
               Index           =   5
               Left            =   4800
               TabIndex        =   36
               Top             =   30
               Width           =   900
               _Version        =   851970
               _ExtentX        =   1587
               _ExtentY        =   794
               _StockProps     =   79
               Caption         =   "인쇄"
               ForeColor       =   -2147483640
               BackColor       =   -2147483636
               Appearance      =   6
               Picture         =   "P_06002.frx":0FAA
            End
            Begin XtremeSuiteControls.PushButton cmdBtn 
               Height          =   450
               Index           =   4
               Left            =   3750
               TabIndex        =   37
               Top             =   30
               Width           =   900
               _Version        =   851970
               _ExtentX        =   1587
               _ExtentY        =   794
               _StockProps     =   79
               Caption         =   "취소"
               ForeColor       =   -2147483640
               BackColor       =   -2147483636
               Appearance      =   6
               Picture         =   "P_06002.frx":1544
            End
            Begin XtremeSuiteControls.PushButton cmdBtn 
               Height          =   450
               Index           =   3
               Left            =   2820
               TabIndex        =   38
               Top             =   30
               Width           =   900
               _Version        =   851970
               _ExtentX        =   1587
               _ExtentY        =   794
               _StockProps     =   79
               Caption         =   "삭제"
               ForeColor       =   -2147483640
               BackColor       =   -2147483636
               Appearance      =   6
               Picture         =   "P_06002.frx":1ADE
            End
            Begin XtremeSuiteControls.PushButton cmdBtn 
               Height          =   450
               Index           =   2
               Left            =   1890
               TabIndex        =   39
               Top             =   30
               Width           =   900
               _Version        =   851970
               _ExtentX        =   1587
               _ExtentY        =   794
               _StockProps     =   79
               Caption         =   "저장"
               ForeColor       =   -2147483640
               BackColor       =   -2147483636
               Appearance      =   6
               Picture         =   "P_06002.frx":2078
            End
            Begin XtremeSuiteControls.PushButton cmdBtn 
               Height          =   450
               Index           =   1
               Left            =   960
               TabIndex        =   40
               Top             =   30
               Width           =   900
               _Version        =   851970
               _ExtentX        =   1587
               _ExtentY        =   794
               _StockProps     =   79
               Caption         =   "신규"
               ForeColor       =   -2147483640
               BackColor       =   -2147483636
               Appearance      =   6
               Picture         =   "P_06002.frx":2612
            End
            Begin XtremeSuiteControls.PushButton cmdBtn 
               Height          =   450
               Index           =   0
               Left            =   30
               TabIndex        =   41
               Top             =   30
               Width           =   900
               _Version        =   851970
               _ExtentX        =   1587
               _ExtentY        =   794
               _StockProps     =   79
               Caption         =   "조회"
               ForeColor       =   -2147483640
               BackColor       =   -2147483636
               Appearance      =   6
               Picture         =   "P_06002.frx":2BAC
            End
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   3
            Left            =   60
            TabIndex        =   43
            Top             =   1350
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "품  목   명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin VB.Label Label1 
            Alignment       =   2  '가운데 맞춤
            Caption         =   "~"
            Height          =   195
            Left            =   9990
            TabIndex        =   17
            Top             =   1050
            Width           =   255
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   6975
         Left            =   15
         TabIndex        =   19
         Top             =   2130
         Width           =   15360
         _Version        =   524288
         _ExtentX        =   27093
         _ExtentY        =   12303
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
         SpreadDesigner  =   "P_06002.frx":3146
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_06002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cboInput_Click(Index As Integer)
    Select Case Index
        Case 5
            Call MasterToAgencyComboAdd(cboInput(4), Mid(cboInput(5).Text, 2, 4))

    End Select
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 0: Call Data_Display           ' 조회
        Case 1:                             ' 신규
        Case 2:                   ' 저장
        
        Case 3:            ' 삭제
        Case 4:            ' 취소
        Case 5:            ' 인쇄
        Case 6:            ' 화면
        Case 7: Unload Me           ' 종료
        
        Case Else
            '
    End Select
    
End Sub

Private Sub Form_Activate()

    If Store.Code = MASTER_CODE Then
        Call SubBottonEnable(cmdBtn, "10000111")
    Else
        Call SubBottonEnable(cmdBtn, "10000111")
    
    End If
    
End Sub

Private Sub Form_Load()
    If P_06002_Flag = False Then
        dtInput(0).Value = Date
        dtInput(1).Value = Date
        
        Call ComboAdd
        
        ReDim sValue(8)
        
        sValue(0) = "1"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("PRO_P_06002_00", sValue(), Err_Num, Err_Dec)
        
        spdView.MaxCols = RS01.Fields.Count
        spdView.MaxRows = RS01.RecordCount
        
        Call spdDisplay(RS01)
        Call GetColWidth(REG_App, Me.Name, spdView)
        
        P_06002_Flag = True
    End If
End Sub
Private Sub spdDisplay(Rs As ADODB.Recordset)

    Call fpSpread_Display(spdView, Rs)
    
    spdView.ColsFrozen = 1 '틀고정
    
    spdView.Row = -1
    
    spdView.Col = 1
    spdView.ColWidth(1) = 10
    spdView.CellType = CellTypeDate
    spdView.TypeDateCentury = True
    spdView.TypeDateFormat = TypeDateFormatYYMMDD
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft
    
    spdView.Col = 2
    spdView.ColWidth(2) = 8
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter
    
    spdView.Col = 3
    spdView.ColWidth(3) = 20
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft

    spdView.Col = 4
    spdView.ColWidth(4) = 7
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter

    spdView.Col = 5
    spdView.ColWidth(5) = 20
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft
    
    spdView.Col = 6
    spdView.ColWidth(6) = 10
    spdView.CellType = CellTypeDate
    spdView.TypeDateCentury = True
    spdView.TypeDateFormat = TypeDateFormatYYMMDD
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft
    
    spdView.Col = 7
    spdView.ColWidth(7) = 10
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft
    
    spdView.Col = 8
    spdView.ColWidth(8) = 12
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 9
    spdView.ColWidth(9) = 12
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 10
    spdView.ColWidth(10) = 10
    spdView.CellType = CellTypeDate
    spdView.TypeDateCentury = True
    spdView.TypeDateFormat = TypeDateFormatYYMMDD
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveColWidth(REG_App, Me.Name & "A", spdView)
    P_06002_Flag = False
End Sub

Public Sub Data_Display()
    Dim i As Integer
    
    ReDim sValue(9)
    
    sValue(0) = "0"
    sValue(1) = Trim(Mid(cboInput(4).Text, 2, 6)) & "%"
    sValue(2) = IIf(optSelect(0).Value = True, "1", "2")
    sValue(3) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(4) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    sValue(5) = Trim(cboInput(0).Text) & "%"
    sValue(6) = Trim(cboInput(1).Text) & "%"
    sValue(7) = Trim(cboInput(2).Text) & "%"
    sValue(8) = Trim(cboInput(3).Text) & "%"
    sValue(9) = Trim(txtInput(4).Text) & "%"
    
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_06002_00", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    Call spdDisplay(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView)

    txtInput(0).Text = ""
    txtInput(1).Text = ""
    txtInput(2).Text = ""
    txtInput(3).Text = ""
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        
        spdView.Col = 6
        If IsNumeric(txtInput(1).Text) = True Then
            txtInput(1).Text = Val(txtInput(1).Text) + spdView.Value
        End If
        
        spdView.Col = 8
        If IsNumeric(txtInput(3).Text) = True Then
            txtInput(3).Text = Val(txtInput(3).Text) + spdView.Value
        End If
        
        spdView.Col = 8
        If spdView.Value <> "" And spdView.Value <> 0 Then
            txtInput(2).Text = Val(txtInput(2).Text) + 1
        End If
    Next i
    
    txtInput(0).Text = Format(spdView.MaxRows, "#,##0")
    txtInput(1).Text = Format(txtInput(1).Text, "#,##0")
    txtInput(2).Text = Format(txtInput(2).Text, "#,##0")
    txtInput(3).Text = Format(txtInput(3).Text, "#,##0")
End Sub

Private Sub ComboAdd()

    Call MasterComboAdd(cboInput(5))
    
    
        '-----------------------------------------------------------------
    '


    ReDim sValue(1)
    
    sValue(0) = "0"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_06001_90", sValue(), Err_Num, Err_Dec)

    cboInput(2).AddItem ""

    Do While Not RS01.EOF
        cboInput(2).AddItem "[" & RS01!담당자코드 & "] " & RS01!담당자명
        
        RS01.MoveNext
    Loop

    ReDim sValue(1)
    
    sValue(0) = "0"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_06001_93", sValue(), Err_Num, Err_Dec)

    Do While Not RS01.EOF
        cboInput(0).AddItem "[" & RS01!품목코드 & "] " & RS01!품목명
        
        RS01.MoveNext
    Loop
    
    ' 크래임 구분
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_06001_91", sValue(), Err_Num, Err_Dec)

    Do While Not RS01.EOF
        ' 탈색, 파손, 이염, 분실, 기타
        'cboInput(4).AddItem "[" & RS01!코드 & "] " & RS01!내용
        cboInput(1).AddItem RS01!내용 & ""
        RS01.MoveNext
    Loop
    RS01.Close

    '보상구분
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_06001_92", sValue(), Err_Num, Err_Dec)

    Do While Not RS01.EOF
        ' 수선, 물품이도후 일부보상, 현금, 제품, 복구
        'cboInput(5).AddItem "[" & RS01!코드 & "] " & RS01!내용
        cboInput(3).AddItem RS01!내용 & ""
        RS01.MoveNext
    Loop
    RS01.Close

End Sub

Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If NewRow <> -1 Then
        spdView.Row = Row
        spdView.Col = -1
        spdView.BackColor = vbWhite
        
        spdView.Row = NewRow
        spdView.Col = -1
        spdView.BackColor = glbYellow
    End If
End Sub

Private Sub spdView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu P_00000.PopUp
    End If
End Sub

Public Sub DataPrint()
    Dim ReportFP As String
    Dim ReportFile As String
    
    ReportFP = GetIniStr("REPORT", "FilePath", "", sIniFile)
    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
    
    Call PrintDesc
    
    P_00000.crPrint.WindowTitle = Me.Caption
    
    Dim i As Integer
    For i = 0 To 30
        P_00000.crPrint.Formulas(i) = ""
    Next
    
    If optSelect(0).Value = True Then
        P_00000.crPrint.Formulas(0) = "검색기준 = '접수일자'"
    ElseIf optSelect(1).Value = True Then
        P_00000.crPrint.Formulas(0) = "검색기준 = '입고일자'"
    End If
    P_00000.crPrint.Formulas(1) = "기사명 = '" & Mid(cboInput(3).Text, 7) & "'"
    P_00000.crPrint.Formulas(2) = "담당자명 = '" & Mid(cboInput(2).Text, 7) & "'"
    P_00000.crPrint.Formulas(3) = "대리점명 = '" & Mid(cboInput(0).Text, 7) & "'"
    P_00000.crPrint.Formulas(4) = "보상건수 = '" & txtInput(2).Text & "'"
    P_00000.crPrint.Formulas(5) = "접수일자1 = '" & Format(dtInput(0).Value, "YYYY-MM-DD") & "'"
    P_00000.crPrint.Formulas(6) = "접수일자2 = '" & Format(dtInput(1).Value, "YYYY-MM-DD") & "'"
    P_00000.crPrint.Formulas(7) = "크레임명 = '" & cboInput(1).Text & "'"
    P_00000.crPrint.Formulas(8) = "크레임총건수 = '" & txtInput(0).Text & "'"
    P_00000.crPrint.Formulas(9) = "합계금액 = '" & txtInput(3).Text & "'"
    P_00000.crPrint.Formulas(10) = "제품금액 = '" & txtInput(1).Text & "'"
    
    Call ReportPrint(ReportFile, "1")
End Sub

Public Sub DataScreen()
    Dim ReportFP As String
    Dim ReportFile As String
    
    ReportFP = GetIniStr("REPORT", "FilePath", "", sIniFile)
    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
    
    Call PrintDesc
    
    P_00000.crPrint.WindowTitle = Me.Caption
    
    Dim i As Integer
    For i = 0 To 30
        P_00000.crPrint.Formulas(i) = ""
    Next
    
    If optSelect(0).Value = True Then
        P_00000.crPrint.Formulas(0) = "검색기준 = '접수일자'"
    ElseIf optSelect(1).Value = True Then
        P_00000.crPrint.Formulas(0) = "검색기준 = '입고일자'"
    End If
    P_00000.crPrint.Formulas(1) = "기사명 = '" & Mid(cboInput(3).Text, 7) & "'"
    P_00000.crPrint.Formulas(2) = "담당자명 = '" & Mid(cboInput(2).Text, 7) & "'"
    P_00000.crPrint.Formulas(3) = "대리점명 = '" & Mid(cboInput(0).Text, 7) & "'"
    P_00000.crPrint.Formulas(4) = "보상건수 = '" & txtInput(2).Text & "'"
    P_00000.crPrint.Formulas(5) = "접수일자1 = '" & Format(dtInput(0).Value, "YYYY-MM-DD") & "'"
    P_00000.crPrint.Formulas(6) = "접수일자2 = '" & Format(dtInput(1).Value, "YYYY-MM-DD") & "'"
    P_00000.crPrint.Formulas(7) = "크레임명 = '" & cboInput(1).Text & "'"
    P_00000.crPrint.Formulas(8) = "크레임총건수 = '" & txtInput(0).Text & "'"
    P_00000.crPrint.Formulas(9) = "합계금액 = '" & txtInput(3).Text & "'"
    P_00000.crPrint.Formulas(10) = "제품금액 = '" & txtInput(1).Text & "'"
    
    Call ReportPrint(ReportFile, "2")
End Sub

Private Sub PrintDesc()
    Dim i As Integer
    Dim TempText As String
    Dim TempFP As String
    Dim TempFile As String
    
    TempFP = GetIniStr("REPORT", "TempPath", "", sIniFile)
    If CheckDirectory(TempFP, True) = False Then
        Exit Sub
    End If
    TempFile = TempFP & "\Temp.txt"
    Open TempFile For Output As #1
    
    TempText = ""
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        
        spdView.Col = 1
        TempText = LeftH(spdView.Text & Space(12), 12)
        spdView.Col = 2
        TempText = TempText & LeftH(spdView.Text & Space(6), 6)
        spdView.Col = 3
        TempText = TempText & LeftH(spdView.Text & Space(14), 14)
        spdView.Col = 7
        TempText = TempText & LeftH(spdView.Text & Space(8), 8)
        spdView.Col = 4
        TempText = TempText & LeftH(spdView.Text & Space(6), 6)
        spdView.Col = 5
        TempText = TempText & LeftH(spdView.Text & Space(12), 12)
        spdView.Col = 6
        TempText = TempText & LeftH(spdView.Text & Space(12), 12)
        spdView.Col = 8
        TempText = TempText & RightH(Space(9) & spdView.Text, 9) & Space(1)
        spdView.Col = 9
        TempText = TempText & RightH(Space(9) & spdView.Text, 9) & Space(1)
        spdView.Col = 10
        TempText = TempText & LeftH(spdView.Text & Space(12), 12)
        
        Print #1, TempText
    Next i
    
    Close #1
End Sub
