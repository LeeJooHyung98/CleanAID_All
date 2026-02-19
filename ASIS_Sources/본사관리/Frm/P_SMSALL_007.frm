VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_SMSALL_7 
   Caption         =   "SMS 월별 발송 현황"
   ClientHeight    =   12330
   ClientLeft      =   420
   ClientTop       =   2880
   ClientWidth     =   17580
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_SMSALL_007.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12330
   ScaleWidth      =   17580
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   12330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17580
      _ExtentX        =   31009
      _ExtentY        =   21749
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_SMSALL_007.frx":058A
      Begin Threed.SSPanel panInput 
         Height          =   765
         Left            =   15
         TabIndex        =   1
         Top             =   540
         Width           =   17550
         _ExtentX        =   30956
         _ExtentY        =   1349
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   1305
            TabIndex        =   2
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy-MM"
            Format          =   64552963
            CurrentDate     =   39244
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "검색년월"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   435
         Index           =   1
         Left            =   6840
         TabIndex        =   4
         Top             =   1320
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   767
         _Version        =   262144
         Font3D          =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "지사기준 일자별 내용"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_SMSALL_007.frx":075C
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   5
         Top             =   15
         Width           =   9945
         _ExtentX        =   17542
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
         Caption         =   " SMS 월별 발송 현황 (P_SMSALL_7)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_SMSALL_007.frx":0BBE
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   0
         Left            =   9975
         TabIndex        =   6
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
         PictureBackground=   "P_SMSALL_007.frx":0DC0
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   7
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
            Picture         =   "P_SMSALL_007.frx":0FC2
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   8
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
            Picture         =   "P_SMSALL_007.frx":155C
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   9
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
            Picture         =   "P_SMSALL_007.frx":1AF6
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   10
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
            Picture         =   "P_SMSALL_007.frx":2090
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   11
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
            Picture         =   "P_SMSALL_007.frx":262A
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   12
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
            Picture         =   "P_SMSALL_007.frx":2BC4
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   13
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
            Picture         =   "P_SMSALL_007.frx":315E
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   14
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
            Picture         =   "P_SMSALL_007.frx":36F8
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10545
         Index           =   1
         Left            =   6840
         TabIndex        =   15
         Top             =   1770
         Width           =   3255
         _Version        =   524288
         _ExtentX        =   5741
         _ExtentY        =   18600
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         MaxCols         =   3
         RetainSelBlock  =   0   'False
         ScrollBars      =   0
         SpreadDesigner  =   "P_SMSALL_007.frx":3C92
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10545
         Index           =   3
         Left            =   13455
         TabIndex        =   16
         Top             =   1770
         Width           =   4110
         _Version        =   524288
         _ExtentX        =   7250
         _ExtentY        =   18600
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         MaxCols         =   3
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "P_SMSALL_007.frx":41D4
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10545
         Index           =   2
         Left            =   10200
         TabIndex        =   17
         Top             =   1770
         Width           =   3240
         _Version        =   524288
         _ExtentX        =   5715
         _ExtentY        =   18600
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         MaxCols         =   3
         RetainSelBlock  =   0   'False
         ScrollBars      =   0
         SpreadDesigner  =   "P_SMSALL_007.frx":4721
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   10545
         Left            =   10110
         TabIndex        =   18
         Top             =   1770
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   18600
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
         PictureBackgroundStyle=   2
         PictureBackground=   "P_SMSALL_007.frx":4C63
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   435
         Index           =   2
         Left            =   10110
         TabIndex        =   19
         Top             =   1320
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   767
         _Version        =   262144
         Font3D          =   3
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
         PictureBackground=   "P_SMSALL_007.frx":4E65
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   435
         Index           =   3
         Left            =   10200
         TabIndex        =   20
         Top             =   1320
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   767
         _Version        =   262144
         Font3D          =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "일자기준 지사별 내용"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_SMSALL_007.frx":52C7
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   435
         Index           =   0
         Left            =   2685
         TabIndex        =   21
         Top             =   1320
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   767
         _Version        =   262144
         Font3D          =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "지사별 전송내역"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_SMSALL_007.frx":5729
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10545
         Index           =   0
         Left            =   2685
         TabIndex        =   22
         Top             =   1770
         Width           =   4140
         _Version        =   524288
         _ExtentX        =   7302
         _ExtentY        =   18600
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         MaxCols         =   3
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "P_SMSALL_007.frx":5B8B
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   435
         Index           =   4
         Left            =   15
         TabIndex        =   23
         Top             =   1320
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   767
         _Version        =   262144
         Font3D          =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "월별 내용"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_SMSALL_007.frx":60D8
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10545
         Index           =   4
         Left            =   15
         TabIndex        =   24
         Top             =   1770
         Width           =   2655
         _Version        =   524288
         _ExtentX        =   4683
         _ExtentY        =   18600
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         MaxCols         =   2
         RetainSelBlock  =   0   'False
         ScrollBars      =   0
         SpreadDesigner  =   "P_SMSALL_007.frx":653A
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_SMSALL_7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Dim P_SMS007_Flag As Boolean

Dim sPrintOption As String

Private Sub Data_Display()
    On Error GoTo ErrRtn
    
    
    ReDim sValue(1)

    Screen.MousePointer = vbHourglass
    sValue(0) = "0"
    sValue(1) = Format(DTPicker1.Value, "yyyy")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_SMSALL_007_00", sValue(), Err_Num, Err_Dec)

    spdView(4).MaxRows = RS01.RecordCount
    
    Call fpSpread_Display(spdView(4), RS01)
    
    With spdView(4)
        If .MaxRows > 0 Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = -1: .BackColor = &HD8FCFE
            
            .Col = 1: .Text = "합계"
            .Col = 2: .Formula = "SUM(B1:B" & .MaxRows - 1 & ")"
        End If
    End With


    sValue(0) = "0"
    sValue(1) = Format(DTPicker1.Value, "yyyyMM")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_SMSALL_007_01", sValue(), Err_Num, Err_Dec)

    spdView(0).MaxRows = RS01.RecordCount
    
    Call fpSpread_Display(spdView(0), RS01)
    
    With spdView(0)
        If .MaxRows > 0 Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = -1: .BackColor = &HD8FCFE
            
            .Col = 1: .Text = "합계"
            .Col = 3: .Formula = "SUM(C1:C" & .MaxRows - 1 & ")"
        End If
    End With
        
    sValue(0) = "0"
    sValue(1) = Format(DTPicker1.Value, "yyyyMM")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_SMSALL_007_03", sValue(), Err_Num, Err_Dec)

    spdView(2).MaxRows = RS01.RecordCount
    Call fpSpread_Display(spdView(2), RS01, False)
    With spdView(2)
        If .MaxRows > 0 Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = -1: .BackColor = &HD8FCFE
            
            .Col = 1: .Text = "합계"
            .Col = 3: .Formula = "SUM(C1:C" & .MaxRows - 1 & ")"
        End If
    End With
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrRtn:
    Screen.MousePointer = vbDefault
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub cboInput_Click(Index As Integer)
    Call Data_Display
End Sub

Private Sub cmdPrint_Click()
'    Call DataScreen2
'    panPrint.Visible = False
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 0: Call Data_Display           ' 조회
        Case 1:            ' 신규
        Case 2:            ' 저장
        Case 3:            ' 삭제
        Case 4:            ' 취소
        Case 5:            ' 인쇄
        Case 6:            ' 화면
        Case 7: Unload Me           ' 종료
        
        Case Else
            '
    End Select
End Sub

Private Sub Command1_Click()
    ' 결과 코드 보기
    panCaption(1).ZOrder 0
    panCaption(1).Visible = Not panCaption(1).Visible
End Sub

Private Sub Form_Activate()
    Call SubBottonEnable(cmdBtn, "10000001")
 
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
    
    If P_SMS007_Flag = False Then
        
        DTPicker1.Value = Now
        
        ReDim sValue(2)

        sValue(0) = "1"
        sValue(1) = ""
        sValue(2) = Format(DTPicker1.Value, "yyyyMM")

'        Set RS01 = New ADODB.Recordset
'        Set RS01 = ExecPro("PRO_SMS_002_11", sValue(), Err_Num, Err_Dec)
'
'        spdView(0).MaxRows = RS01.RecordCount
'
'        Call spdDisplay1(RS01)
        
        P_SMS007_Flag = True
    End If
        
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    Dim i As Integer
    
    For i = 0 To 3
        With spdView(i)
            .MaxRows = 0
            .RowHeight(-1) = 14
    
            'Spread 8 - 디자인
            .HighlightHeaders = HighlightHeadersOff
            .AppearanceStyle = AppearanceStyleEnhanced
            .ScrollBarStyle = ScrollBarStyleVisualStyle
    
            '선택된 Row
            .SelBackColor = &HFFFFC0 '황색 ^^
            .SelForeColor = &H0&     '검은글씨
            .OperationMode = OperationModeSingle
            
            'Init the User Sort
            .UserColAction = UserColActionSort
        End With
    Next i

    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_SMS007_Flag = False
End Sub

Private Sub Data_Display2()
    Dim i As Integer
   
    On Error GoTo ErrRtn
    ReDim sValue(2)
    
    Screen.MousePointer = vbHourglass

    sValue(0) = "0"
    spdView(0).Row = spdView(0).ActiveRow
    spdView(0).Col = 1
    sValue(1) = spdView(0).Text
    
    sValue(2) = Format(DTPicker1.Value, "yyyyMM")

    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_SMSALL_007_02", sValue(), Err_Num, Err_Dec)

    With spdView(1)
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = RS01!전송일자 & ""
            .Col = 2: .Text = RS01!요일 & ""
            .Col = 3: .Text = RS01!전송수량 & ""
            
            RS01.MoveNext
        Loop
        If .MaxRows > 0 Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = -1: .BackColor = &HD8FCFE
            
            .Col = 1: .Text = "합계"
            .Col = 3: .Formula = "SUM(C1:C" & .MaxRows - 1 & ")"
        End If
    
        .Redraw = True
        RS01.Close
    
    End With
    Set RS01 = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Data_Display3()
    On Error GoTo ErrRtn
    
    
    ReDim sValue(1)

    Screen.MousePointer = vbHourglass
    sValue(0) = "0"
    spdView(2).Col = 1: spdView(2).Row = spdView(2).ActiveRow
    sValue(1) = Replace(spdView(2).Text, "-", "")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_SMSALL_007_04", sValue(), Err_Num, Err_Dec)

    spdView(3).MaxRows = RS01.RecordCount
    
    Call fpSpread_Display(spdView(3), RS01)
    
    With spdView(3)
        If .MaxRows > 0 Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = -1: .BackColor = &HD8FCFE
            
            .Col = 1: .Text = "합계"
            .Col = 3: .Formula = "SUM(C1:C" & .MaxRows - 1 & ")"
        End If
    End With
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrRtn:
    Screen.MousePointer = vbDefault
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)

End Sub

Public Sub DataAdd()


End Sub

Public Sub DataCancel()
    '
End Sub

Public Sub DataDelete()
    '
End Sub

Public Sub DataSave()

End Sub

Public Sub DataPrint()
    '
End Sub


Public Sub DataScreen()
'    panPrint.Visible = True

    sPrintOption = "2"
End Sub



Private Sub spdView_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
    Select Case Index
        Case 0
            spdView(1).MaxRows = 0
            Call Data_Display2
            
        Case 2
            spdView(3).MaxRows = 0
            Call Data_Display3
    End Select
End Sub

 
