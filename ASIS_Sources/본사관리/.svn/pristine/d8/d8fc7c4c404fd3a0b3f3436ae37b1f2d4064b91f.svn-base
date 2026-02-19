VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_SMSALL_2 
   Caption         =   "SMS 발송 현황"
   ClientHeight    =   12330
   ClientLeft      =   420
   ClientTop       =   2325
   ClientWidth     =   17580
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
   ScaleHeight     =   12330
   ScaleWidth      =   17580
   WindowState     =   2  '최대화
   Begin Threed.SSPanel panCaption 
      Height          =   8595
      Index           =   1
      Left            =   7560
      TabIndex        =   1
      Top             =   2100
      Visible         =   0   'False
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   15161
      _Version        =   262144
      BevelOuter      =   1
      BevelInner      =   1
      RoundedCorners  =   0   'False
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   7605
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   13414
         _Version        =   393217
         TextRTF         =   $"P_SMSALL_002.frx":0000
      End
   End
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
      PaneTree        =   "P_SMSALL_002.frx":090B
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10545
         Index           =   0
         Left            =   15
         TabIndex        =   3
         Top             =   1770
         Width           =   4530
         _Version        =   524288
         _ExtentX        =   7990
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
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "P_SMSALL_002.frx":0A3D
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10545
         Index           =   1
         Left            =   4560
         TabIndex        =   4
         Top             =   1770
         Width           =   2460
         _Version        =   524288
         _ExtentX        =   4339
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
         RetainSelBlock  =   0   'False
         ScrollBars      =   0
         SpreadDesigner  =   "P_SMSALL_002.frx":0ED3
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10545
         Index           =   2
         Left            =   7035
         TabIndex        =   5
         Top             =   1770
         Width           =   10530
         _Version        =   524288
         _ExtentX        =   18574
         _ExtentY        =   18600
         _StockProps     =   64
         BackColorStyle  =   1
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
         SpreadDesigner  =   "P_SMSALL_002.frx":1374
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   795
         Left            =   15
         TabIndex        =   6
         Top             =   540
         Width           =   17550
         _ExtentX        =   30956
         _ExtentY        =   1402
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.CommandButton Command1 
            Caption         =   "결과 코드"
            Height          =   555
            Left            =   7020
            TabIndex        =   8
            Top             =   210
            Width           =   1515
         End
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            Style           =   2  '드롭다운 목록
            TabIndex        =   7
            Top             =   405
            Width           =   2775
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   1530
            TabIndex        =   9
            Top             =   60
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy-MM"
            Format          =   55836675
            CurrentDate     =   39244
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   10
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "검색년월"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   11
            Top             =   405
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "사업장"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   405
         Index           =   2
         Left            =   7035
         TabIndex        =   12
         Top             =   1350
         Width           =   10530
         _ExtentX        =   18574
         _ExtentY        =   714
         _Version        =   262144
         Font3D          =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "전송 메시지 내역"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_SMSALL_002.frx":17FF
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   405
         Index           =   1
         Left            =   4560
         TabIndex        =   13
         Top             =   1350
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   714
         _Version        =   262144
         Font3D          =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "전송일자"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_SMSALL_002.frx":1C61
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   14
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
         Caption         =   " SMS 발송 현황 (P_SMSALL_2)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_SMSALL_002.frx":20C3
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
         TabIndex        =   15
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
         PictureBackground=   "P_SMSALL_002.frx":22C5
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   16
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
            Picture         =   "P_SMSALL_002.frx":24C7
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   17
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
            Picture         =   "P_SMSALL_002.frx":2A61
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   18
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
            Picture         =   "P_SMSALL_002.frx":2FFB
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   19
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
            Picture         =   "P_SMSALL_002.frx":3595
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   20
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
            Picture         =   "P_SMSALL_002.frx":3B2F
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   21
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
            Picture         =   "P_SMSALL_002.frx":40C9
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   22
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
            Picture         =   "P_SMSALL_002.frx":4663
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   23
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
            Picture         =   "P_SMSALL_002.frx":4BFD
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   405
         Index           =   0
         Left            =   15
         TabIndex        =   24
         Top             =   1350
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   714
         _Version        =   262144
         Font3D          =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "가맹점별 전송내역"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_SMSALL_002.frx":5197
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
End
Attribute VB_Name = "P_SMSALL_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Dim P_SMS002_Flag As Boolean

Dim sPrintOption As String

Public Sub Data_Display()

    If Trim(cboInput(0).Text) = "" Then
        MsgBox "사업장을 확인하여 주십시요.", vbInformation, "확인"
        cboInput(0).SetFocus
        Exit Sub
    End If

    ReDim sValue(2)

    sValue(0) = "0"
    sValue(1) = Mid(Trim(cboInput(0).Text), 2, 4)
    sValue(2) = Format(DTPicker1.Value, "yyyyMM")
    ' 대리점 정보
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_SMSALL_002_03", sValue(), Err_Num, Err_Dec)

    spdView(0).MaxCols = RS01.Fields.Count
    spdView(0).MaxRows = RS01.RecordCount

    Call spdDisplay1(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView(0))
    If spdView(0).MaxRows > 0 Then
        Call spdView_Click(0, 1, 1)
    End If
End Sub

Private Sub spdDisplay1(Rs As ADODB.Recordset)
    Call fpSpread_Display(spdView(0), Rs)
End Sub

Private Sub spdDisplay2(Rs As ADODB.Recordset)
    Call fpSpread_Display(spdView(1), Rs)
End Sub

Private Sub spdDisplay3(Rs As ADODB.Recordset)
    Call fpSpread_Display(spdView(2), Rs)
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
End Sub

Private Sub Form_Load()
    With spdView(0)
        .ColsFrozen = 1  '틀고정
        .Row = -1

        .Col = 1
        .ColWidth(1) = 5
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter

        .Col = 2
        .ColWidth(2) = 14
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft

        .Col = 3
        .ColWidth(3) = 8
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 0
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignRight

        .Col = 4
        .ColWidth(4) = 8
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 0
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignRight
    End With

    With spdView(1)
        .ColsFrozen = 1  '틀고정
        .Row = -1

        .Col = 1
        .ColWidth(1) = 10
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter

        .Col = 2
        .ColWidth(2) = 8
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 0
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignRight
    End With
    
    With spdView(2)
        .ColsFrozen = 1  '틀고정
        .Row = -1

        .Col = 1
        .ColWidth(1) = 8
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter

        .Col = 2
        .ColWidth(2) = 10
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
        
        .Col = 3
        .ColWidth(3) = 14
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft

        .Col = 4
        .ColWidth(4) = 10
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
    
        .Col = 5
        .ColWidth(5) = 30
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
    
        .Col = 6
        .ColWidth(6) = 30
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
    End With


    If P_SMS002_Flag = False Then
        ' Combo BOX의 내역을 채운다.
        'Call ComboAdd
'
'        If Store.Code = "1000" Then
'            panCaption(35).Visible = True
'            cboInput(0).Visible = True
'        End If
        DTPicker1.Value = Now
        
        If Store.Code = "1000" Then
            Call Master_tblComboAdd(cboInput(0))
        Else
            cboInput(0).AddItem "[" & Store.Code & "] " & Store.Name
            cboInput(0).ListIndex = 0
            cboInput(0).Enabled = False

        End If

    
        ReDim sValue(2)

        sValue(0) = "1"
        sValue(1) = ""
        sValue(2) = Format(DTPicker1.Value, "yyyyMM")

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("PRO_SMS_002_11", sValue(), Err_Num, Err_Dec)

        spdView(0).MaxCols = RS01.Fields.Count
        spdView(0).MaxRows = RS01.RecordCount

        Call spdDisplay1(RS01)
        Call GetColWidth(REG_App, Me.Name, spdView(0))
        'Call GetColWidth(REG_App, Me.Name, spdView2)

        P_SMS002_Flag = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_SMS002_Flag = False
End Sub



Private Sub Data_Display2()
    Dim i As Integer
   
    ReDim sValue(2)

    sValue(0) = "0"
    spdView(0).Row = spdView(0).ActiveRow
    spdView(0).Col = 1
    sValue(1) = spdView(0).Text
    
    sValue(2) = Format(DTPicker1.Value, "yyyyMM")

    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_SMSALL_002_01", sValue(), Err_Num, Err_Dec)

    spdView(1).MaxCols = RS01.Fields.Count
    spdView(1).MaxRows = RS01.RecordCount

    Call spdDisplay2(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView(1))
    
    If spdView(1).MaxRows > 0 Then
        Call spdView_Click(1, 1, 1)
    End If
    ' 추가 상세 정보를 조회 한다.
    Call Data_Display3
    
    Set RS01 = Nothing

End Sub

Private Sub Data_Display3()
    Dim lRow As Long

    ReDim sValue(2)

    If spdView(1).MaxRows <= 0 Then Exit Sub

    sValue(0) = "0"
    
    
    spdView(0).Row = spdView(0).ActiveRow
    spdView(0).Col = 1
    sValue(1) = spdView(0).Text
    
    spdView(1).Row = spdView(1).ActiveRow
    spdView(1).Col = 1
    'sValue(2) = spdView(1).Text
    sValue(2) = Format(spdView(1).Text, "YYYY-MM-DD")

    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_SMSALL_002_02", sValue(), Err_Num, Err_Dec)

    spdView(2).MaxCols = RS01.Fields.Count
    spdView(2).MaxRows = RS01.RecordCount

    Call spdDisplay3(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView(2))

    ' 비정상적인 내용은 별도로 표시한다.
    For lRow = 0 To spdView(2).MaxRows
        With spdView(2)
            .Row = lRow
            .Col = 6
            If Left(Trim(.Text), 2) <> "06" Then
                .Col = -1: .BackColor = vbRed
            End If
        End With
    
    Next lRow
    Set RS01 = Nothing


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
    panPrint.Visible = True

    sPrintOption = "2"
End Sub

Private Sub spdView_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        'PopupMenu P_00000.PopUp
    End If
End Sub

Private Sub spdView_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
    Select Case Index
        Case 0
            spdView(1).MaxRows = 0
            spdView(2).MaxRows = 0
            Call Data_Display2
            
        Case 1
            spdView(2).MaxRows = 0
            Call Data_Display3
    End Select
End Sub

Private Sub spdView_LeaveCell(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If NewRow <> -1 Then
        With spdView(Index)
            If NewRow <> -1 Then
                .Col = 6:   .Row = Row
                If Index = 2 And Left(Trim(.Text), 2) <> "06" Then

                        .Col = -1:  .BackColor = vbRed

                Else
                    '.Row = Row
                    If (Row Mod 2) = 0 Then
                        .Col = -1: .BackColor = glbGray
                    Else
                        .Col = -1: .BackColor = vbWhite
                    End If
                    
                    .Row = NewRow
                    .Col = -1: .BackColor = glbYellow
                End If
            End If
        End With
    End If
End Sub
