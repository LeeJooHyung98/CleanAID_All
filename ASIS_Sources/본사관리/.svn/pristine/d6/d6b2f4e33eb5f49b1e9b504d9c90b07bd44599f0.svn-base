VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_01011_B 
   Caption         =   "[본사]가맹점 의류분류 등록"
   ClientHeight    =   12165
   ClientLeft      =   3675
   ClientTop       =   2505
   ClientWidth     =   15885
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_01011_B.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12165
   ScaleWidth      =   15885
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   12165
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15885
      _ExtentX        =   28019
      _ExtentY        =   21458
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_01011_B.frx":058A
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   1
         Top             =   540
         Width           =   15855
         _ExtentX        =   27966
         _ExtentY        =   1376
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin Threed.SSOption optGubun 
            Height          =   270
            Index           =   0
            Left            =   180
            TabIndex        =   2
            Top             =   270
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   476
            _Version        =   262144
            BackColor       =   16777215
            Caption         =   "가맹점종류별"
         End
         Begin Threed.SSOption optGubun 
            Height          =   255
            Index           =   1
            Left            =   1950
            TabIndex        =   3
            Top             =   270
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   450
            _Version        =   262144
            BackColor       =   16777215
            Caption         =   "지사별"
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Left            =   4635
            TabIndex        =   16
            Top             =   75
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   69271552
            CurrentDate     =   36686
         End
         Begin XtremeSuiteControls.ProgressBar ProgressBar 
            Height          =   330
            Left            =   4620
            TabIndex        =   18
            Top             =   420
            Visible         =   0   'False
            Width           =   3030
            _Version        =   851970
            _ExtentX        =   5345
            _ExtentY        =   582
            _StockProps     =   93
            Scrolling       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Shape Shape 
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00808080&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  '단색
            Height          =   630
            Left            =   45
            Shape           =   4  '둥근 사각형
            Top             =   75
            Width           =   3000
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "적용일자:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   3315
            TabIndex        =   17
            Top             =   120
            Width           =   1260
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   4
         Top             =   15
         Width           =   8250
         _ExtentX        =   14552
         _ExtentY        =   900
         _Version        =   262144
         Font3D          =   3
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
         Caption         =   " [본사]가맹점 의류분류 등록 (P_01011_B)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_01011_B.frx":065C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Left            =   8280
         TabIndex        =   5
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
         PictureBackground=   "P_01011_B.frx":085E
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   6
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
            Picture         =   "P_01011_B.frx":0A60
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   7
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
            Picture         =   "P_01011_B.frx":0FFA
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   8
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
            Picture         =   "P_01011_B.frx":1594
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   9
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
            Picture         =   "P_01011_B.frx":1B2E
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   10
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
            Picture         =   "P_01011_B.frx":20C8
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   11
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
            Picture         =   "P_01011_B.frx":2662
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   12
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
            Picture         =   "P_01011_B.frx":2BFC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   13
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
            Picture         =   "P_01011_B.frx":3196
         End
      End
      Begin FPSpreadADO.fpSpread sprList 
         Height          =   7455
         Left            =   15
         TabIndex        =   14
         Top             =   4695
         Width           =   4650
         _Version        =   524288
         _ExtentX        =   8202
         _ExtentY        =   13150
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
         MaxCols         =   4
         ScrollBars      =   2
         SpreadDesigner  =   "P_01011_B.frx":3730
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread sprMaster 
         Height          =   3345
         Left            =   15
         TabIndex        =   15
         Top             =   1335
         Width           =   4650
         _Version        =   524288
         _ExtentX        =   8202
         _ExtentY        =   5900
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
         MaxCols         =   2
         ScrollBars      =   2
         SpreadDesigner  =   "P_01011_B.frx":3CE9
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin XtremeSuiteControls.TabControl TabControl1 
         Height          =   10815
         Left            =   4680
         TabIndex        =   19
         Top             =   1335
         Width           =   11190
         _Version        =   851970
         _ExtentX        =   19738
         _ExtentY        =   19076
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   3
         Color           =   64
         PaintManager.BoldSelected=   -1  'True
         PaintManager.HotTracking=   -1  'True
         PaintManager.ButtonMargin=   "2,3,2,3"
         ItemCount       =   3
         SelectedItem    =   1
         Item(0).Caption =   "가맹점 의류분류 현황"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "TabControlPage(0)"
         Item(1).Caption =   "의류분류 등록현황"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "TabControlPage(1)"
         Item(2).Caption =   "마진 변경 내역 조회"
         Item(2).ControlCount=   1
         Item(2).Control(0)=   "TabControlPage1"
         Begin XtremeSuiteControls.TabControlPage TabControlPage1 
            Height          =   10365
            Left            =   -69970
            TabIndex        =   29
            Top             =   420
            Visible         =   0   'False
            Width           =   11130
            _Version        =   851970
            _ExtentX        =   19632
            _ExtentY        =   18283
            _StockProps     =   1
            Page            =   2
            Begin FPSpreadADO.fpSpread spView2 
               Height          =   9735
               Left            =   90
               TabIndex        =   30
               Top             =   510
               Width           =   10440
               _Version        =   524288
               _ExtentX        =   18415
               _ExtentY        =   17171
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
               SpreadDesigner  =   "P_01011_B.frx":420B
               Appearance      =   1
               HighlightHeaders=   1
               HighlightStyle  =   1
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   315
               Index           =   0
               Left            =   1410
               TabIndex        =   31
               Top             =   90
               Width           =   2865
               _ExtentX        =   5054
               _ExtentY        =   556
               _Version        =   393216
               DateIsNull      =   -1  'True
               Format          =   69271552
               CurrentDate     =   36686
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   315
               Index           =   1
               Left            =   4620
               TabIndex        =   33
               Top             =   90
               Width           =   2865
               _ExtentX        =   5054
               _ExtentY        =   556
               _Version        =   393216
               DateIsNull      =   -1  'True
               Format          =   69271552
               CurrentDate     =   36686
            End
            Begin VB.Label Label 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "생성일자:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   1
               Left            =   90
               TabIndex        =   32
               Top             =   135
               Width           =   1260
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage 
            Height          =   10365
            Index           =   1
            Left            =   30
            TabIndex        =   20
            Top             =   420
            Width           =   11130
            _Version        =   851970
            _ExtentX        =   19632
            _ExtentY        =   18283
            _StockProps     =   1
            Page            =   1
            Begin VB.ComboBox cboInput 
               Height          =   315
               Index           =   0
               Left            =   930
               TabIndex        =   22
               Text            =   "cboInput"
               Top             =   90
               Width           =   3015
            End
            Begin VB.ComboBox cboInput 
               Height          =   315
               Index           =   1
               Left            =   930
               Style           =   2  '드롭다운 목록
               TabIndex        =   21
               Top             =   450
               Width           =   3015
            End
            Begin FPSpreadADO.fpSpread sprSchedule 
               Height          =   7245
               Left            =   60
               TabIndex        =   23
               Top             =   825
               Width           =   4785
               _Version        =   524288
               _ExtentX        =   8440
               _ExtentY        =   12779
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
               MaxCols         =   3
               ScrollBars      =   2
               SpreadDesigner  =   "P_01011_B.frx":4912
               Appearance      =   1
               HighlightHeaders=   1
               HighlightStyle  =   1
            End
            Begin FPSpreadADO.fpSpread sprSchedule2 
               Height          =   9225
               Left            =   4890
               TabIndex        =   28
               Top             =   840
               Width           =   7950
               _Version        =   524288
               _ExtentX        =   14023
               _ExtentY        =   16272
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
               SpreadDesigner  =   "P_01011_B.frx":4E7B
               Appearance      =   1
               HighlightHeaders=   1
               HighlightStyle  =   1
            End
            Begin VB.Label lblProgress 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "가맹점명:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   5
               Left            =   0
               TabIndex        =   25
               Top             =   525
               Width           =   900
            End
            Begin VB.Label lblProgress 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "지사명:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   6
               Left            =   0
               TabIndex        =   24
               Top             =   150
               Width           =   900
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage 
            Height          =   10365
            Index           =   0
            Left            =   -69970
            TabIndex        =   26
            Top             =   420
            Visible         =   0   'False
            Width           =   11130
            _Version        =   851970
            _ExtentX        =   19632
            _ExtentY        =   18283
            _StockProps     =   1
            Page            =   0
            Begin FPSpreadADO.fpSpread spdView 
               Height          =   10815
               Left            =   60
               TabIndex        =   27
               Top             =   75
               Width           =   11190
               _Version        =   524288
               _ExtentX        =   19738
               _ExtentY        =   19076
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
               MaxCols         =   6
               SpreadDesigner  =   "P_01011_B.frx":554A
               Appearance      =   1
               HighlightHeaders=   1
               HighlightStyle  =   1
            End
         End
      End
   End
End
Attribute VB_Name = "P_01011_B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim strSql As String
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub SPR_Resize()
    On Error GoTo ErrRtn
    
    sprSchedule.Height = Me.Height - 3140
    
    sprSchedule2.Width = Me.Width - 10000
    sprSchedule2.Height = Me.Height - 3140
    
    spView2.Width = Me.Width - 5000
    spView2.Height = Me.Height - 2800
    
    spdView.Width = Me.Width - 5000
    spdView.Height = Me.Height - 2360

    Exit Sub
    
ErrRtn:

End Sub

Private Sub cboInput_Click(Index As Integer)
    Dim sCode As String

    If Index = 0 Then
        sCode = Trim(Mid(cboInput(0).Text, 2, 4))

        Call Get_가맹점리스트(cboInput(1), sCode)

    ElseIf Index = 1 Then
           Call Data_Display
    End If
End Sub

Private Sub cboInput_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SearchString_한글 KeyAscii
'    Else
       ' SearchString KeyAscii
    End If

End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
'    Me.MousePointer = 11
    
    Select Case Index
        Case 0: Call Data_DisplayHi   ' 조회
        Case 1: 'Call DataAdd        ' 신규
        Case 2: Call DataSave       ' 저장
        Case 3: Call DataDelete     ' 삭제
        Case 4: Call DataCancel     ' 취소
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
    cmdBtn(1).Enabled = True
    cmdBtn(2).Enabled = True
    
    cmdBtn(4).Enabled = True
    cmdBtn(5).Enabled = True
    cmdBtn(6).Enabled = True

    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
End Sub

'Private Sub spdDisplay(RS As ADODB.Recordset)
'
'    Call fpSpread_Display(spdView(0), RS)
'
'    spdView(0).DataSource = Nothing
'End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    With sprMaster
        .MaxRows = 0
        .RowHeight(-1) = 13
        
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
    
    With sprList
        .MaxRows = 0
        .RowHeight(-1) = 13
        
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
    
    With spdView
        .MaxRows = 0
        .RowHeight(-1) = 14
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeNormal
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
    
    With sprSchedule
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
    
    dtInput.Value = Date
    
    
    DTPicker1(0).Value = Format(Date, "yyyy-MM-01")
    DTPicker1(1).Value = Date
    
    optGubun(1).Value = True
    
    TabControl1.SelectedItem = 0
    
    Call Get_지사리스트(cboInput(0), False)
    
    Call SPR_Resize
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Resize()
    Call SPR_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'P_01004_A_Flag = False
End Sub

Private Sub DataAdd()

End Sub

Private Sub DataSave()
    Dim i          As Integer
    Dim iRow       As Long
    Dim iSale(1)   As Integer
    
    Dim 가맹점코드 As String
    Dim 의류코드   As String
    Dim 의류명     As String
    Dim ActionTime  As String
    
    
    
    ProgressBar.Value = 0
    ProgressBar.Min = 0
    ProgressBar.Max = 100
    ProgressBar.Visible = True
    
    If sprList.ActiveRow <= 0 Then Exit Sub
    If spdView.MaxRows <= 0 Then Exit Sub
    
    ActionTime = Format(Now, "yyyy-MM-dd hh:mm:ss")
    
    sprList.Row = sprList.ActiveRow
    sprList.Col = 1: 가맹점코드 = sprList.Text & ""
    
    Set RS01 = New ADODB.Recordset
    
    For i = 1 To spdView.MaxRows
        ProgressBar.Value = (i / spdView.MaxRows) * 100
        DoEvents
    
        '-------------------------------------------------------------------
        ' TB_가맹점의류분류 저장 - SP_01011_B_00
        '-------------------------------------------------------------------
        ReDim sValue(10)
        
        sValue(0) = 가맹점코드                            '1 가맹점코드
        sValue(1) = Format(dtInput.Value, "YYYY-MM-DD")   '2 적용일자
        
        spdView.Row = i
        spdView.Col = 1: sValue(2) = spdView.Text & ""    '3 의류분류코드
        spdView.Col = 2: sValue(3) = spdView.Text & ""    '4 의류분류명
        spdView.Col = 3: sValue(4) = spdView.Value & ""   '5 세탁마진
        spdView.Col = 4: sValue(5) = spdView.Value & ""   '6 외주마진
        spdView.Col = 5: sValue(6) = spdView.Value & ""   '7 수선마진
        spdView.Col = 6: sValue(7) = spdView.Value & ""   '8 순서
        
        sValue(8) = ActionTime
        sValue(9) = UserID
        sValue(10) = USERNAME
        
        Call ExecPro("SP_01011_B_00", sValue(), Err_Num, Err_Dec)
    Next i

    ProgressBar.Visible = False

    If Err_Num = 0 Then
        MsgBox "해당되는 데이터가 정상적으로 저장이 되었습니다.", vbInformation
    End If
End Sub

Public Sub DataDelete()
    Rtn = MsgBox("해당되는 데이터를 삭제하시겠습니까?", vbYesNo + vbInformation + vbDefaultButton2, "데이터 삭제")
    
    If Rtn = vbNo Then Exit Sub
    
    ReDim sValue(1)
    
    sprList.Row = sprList.ActiveRow
    sprList.Col = 1: sValue(0) = sprList.Text & ""                   '1 가맹점코드
                     sValue(1) = Format(dtInput.Value, "YYYY-MM-DD") '2 적용일자
    
    Call ExecPro("SP_01011_B_04", sValue(), Err_Num, Err_Dec)
    
    If Err_Num = 0 Then
        spdView.MaxRows = 0
        
        MsgBox "해당되는 데이터가 정상적으로 삭제가 되었습니다.", vbInformation
    End If
End Sub

Public Sub DataCancel()

End Sub


Private Sub optGubun_Click(Index As Integer, Value As Integer)
    sprList.MaxRows = 0
    DoEvents
    
    If Index = 0 Then
        ReDim sValue(0)
        
        sValue(0) = "0"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_01004_A_06", sValue(), Err_Num, Err_Dec)
    
        With sprMaster
            .MaxRows = 0
            .Redraw = False
            
            Do Until RS01.EOF
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                
                .Col = 1: .Text = RS01!가맹점구분코드 & ""
                .Col = 2: .Text = RS01!가맹점구분명 & ""
                
                RS01.MoveNext
            Loop
            RS01.Close
            Set RS01 = Nothing
            
            .Redraw = True
        End With
        
    Else
        ReDim sValue(0)
        
        sValue(0) = "0"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_A_0001", sValue(), Err_Num, Err_Dec)
    
        With sprMaster
            .EventEnabled(EventButtonClicked) = False '버튼클릭 이벤트 죽임
            
            .MaxRows = 0
            .Redraw = False
            
            Do Until RS01.EOF
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                
                .Col = 1: .Text = RS01!지사코드 & ""
                .Col = 2: .Text = RS01!지사명 & ""
                
                RS01.MoveNext
            Loop
            RS01.Close
            Set RS01 = Nothing
            
            .EventEnabled(EventButtonClicked) = True
            
            .Redraw = True
        End With
    End If
End Sub

Private Sub sprList_Click(ByVal Col As Long, ByVal Row As Long)
    If Row <= 0 Then Exit Sub
    
    ReDim sValue(0)
    
    sprList.Row = Row
    sprList.Col = 1: sValue(0) = sprList.Text & ""     '가맹점코드
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01011_B_03", sValue(), Err_Num, Err_Dec)
    
    With spdView
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = Trim(RS01!의류분류코드) & "" ' 1
            .Col = 2: .Text = Trim(RS01!의류분류명) & ""   ' 2
            .Col = 3: .Text = RS01!세탁마진 & ""           ' 3
            .Col = 4: .Text = RS01!외주마진 & ""           ' 4
            .Col = 5: .Text = RS01!수선마진 & ""           ' 5
            .Col = 6: .Text = RS01!순서 & ""               ' 6
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
    
End Sub

Private Sub sprMaster_Click(ByVal Col As Long, ByVal Row As Long)
    Dim sWork        As String
    Dim 코드         As String
    Dim 가맹정구분명 As String
    
    If Row <= 0 Then Exit Sub
    
    sprMaster.Row = Row
    sprMaster.Col = 1: 코드 = Trim(sprMaster.Text) & ""
    sprMaster.Col = 2: 가맹정구분명 = Trim(sprMaster.Text) & ""
    
    If optGubun(0).Value = True Then
        Call 가맹점2_Display(코드, 가맹정구분명)
    Else
        Call 가맹점_Display(코드)
    End If
End Sub

Private Sub sprMaster_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Call sprMaster_Click(NewCol, NewRow)
End Sub

Private Sub 가맹점_Display(지사코드 As String)
    ReDim sValue(0)
    
    sValue(0) = 지사코드
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01001_00_MASTER", sValue(), Err_Num, Err_Dec)
    
    With sprList
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            If (RS01!종료일자 = "") Or (RS01!종료일자 = "2099-12-31") Then '가맹점이 현 지사에서 관리중...
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                
                .Col = 1: .Text = RS01(0) & ""  '2
                .Col = 2: .Text = RS01(1) & ""  '3
                .Col = 3: .Text = 지사코드 & "" '4
            End If
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
End Sub


Private Sub 가맹점2_Display(가맹점구분 As String, 가맹점구분명 As String)
    ReDim sValue(0)
    
    sValue(0) = 가맹점구분
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01004_A_07", sValue(), Err_Num, Err_Dec)
    
    With sprList
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            If (RS01!종료일자 = "") Or (RS01!종료일자 = "2099-12-31") Then '가맹점이 현 지사에서 관리중...
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                
                .Col = 1: .Text = RS01(0) & ""       '2
                .Col = 2: .Text = RS01(1) & ""       '3
                .Col = 3: .Text = RS01(2) & ""       '4
                .Col = 4: .Text = 가맹점구분명 & ""  '5
            End If
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    If Trim(cboInput(0).Text) = "" Then
        'MsgBox "사업장을 선택 하세요", vbInformation
        cboInput(0).SetFocus
        Exit Sub
    End If
    
    If Trim(cboInput(1).Text) = "" Then
        'MsgBox "가맹점을 선택 하세요", vbInformation
        cboInput(1).SetFocus
        Exit Sub
    End If
    
    '-------------------------------------------------------------------
    ' SP_01004_A_00
    '-------------------------------------------------------------------
    ReDim sValue(0)

    sValue(0) = Mid(cboInput(1).Text, 2, 6)
        
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01011_B_01", sValue(), Err_Num, Err_Dec)
    
    With sprSchedule
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = RS01!가맹점코드 & ""
            .Col = 2: .Text = RS01!가맹점명 & ""
            .Col = 3: .Text = RS01!적용일자 & ""
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub


Private Sub Data_DisplayHi()
    On Error GoTo ErrRtn
 
    ReDim sValue(1)

    sValue(0) = Format(DTPicker1(0).Value, "yyyy-MM-dd")
    sValue(1) = Format(DTPicker1(1).Value, "yyyy-MM-dd")
        
        
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01011_B_07", sValue(), Err_Num, Err_Dec)
    
    With spView2
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
                .Col = 1: .Text = Trim(RS01(0) & "")
                .Col = 2: .Text = Trim(RS01(1) & "")
                .Col = 3: .Text = Trim(RS01(2) & "")
                .Col = 4: .Text = Trim(RS01(3) & "")
                .Col = 5: .Text = Trim(RS01(4) & "")
                .Col = 6: .Text = Trim(RS01(5) & "")
                .Col = 7: .Text = Trim(RS01(6) & "")
                .Col = 8: .Text = Trim(RS01(7) & "")
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub


Private Sub sprSchedule_Click(ByVal Col As Long, ByVal Row As Long)
    
    If Row <= 0 Then Exit Sub
    
    ReDim sValue(1)
    
    sprSchedule.Row = Row
    sprSchedule.Col = 1: sValue(0) = sprSchedule.Text & ""     '가맹점코드
    sprSchedule.Col = 3: sValue(1) = sprSchedule.Text & ""     '적용일자
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01011_B_05", sValue(), Err_Num, Err_Dec)
    
    With sprSchedule2
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = Trim(RS01!의류분류코드) & "" ' 1
            .Col = 2: .Text = Trim(RS01!의류분류명) & ""   ' 2
            .Col = 3: .Text = RS01!세탁마진 & ""           ' 3
            .Col = 4: .Text = RS01!외주마진 & ""           ' 4
            .Col = 5: .Text = RS01!수선마진 & ""           ' 5
            .Col = 6: .Text = RS01!순서 & ""               ' 6
            .Col = 7: .Text = RS01!수신일자 & ""           ' 7
            .Col = 8: .Text = RS01!생성일자 & ""           ' 8
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With

End Sub
