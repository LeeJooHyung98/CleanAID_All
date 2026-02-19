VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_SMSALL_9 
   Caption         =   " SMS 기간별 등록 현황"
   ClientHeight    =   12330
   ClientLeft      =   825
   ClientTop       =   2895
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
   Icon            =   "P_SMSALL_009.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
      PaneTree        =   "P_SMSALL_009.frx":058A
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10980
         Index           =   0
         Left            =   15
         TabIndex        =   1
         Top             =   1335
         Width           =   17550
         _Version        =   524288
         _ExtentX        =   30956
         _ExtentY        =   19368
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
         SpreadDesigner  =   "P_SMSALL_009.frx":061C
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   2
         Top             =   540
         Width           =   17550
         _ExtentX        =   30956
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   0
            Left            =   1245
            Style           =   2  '드롭다운 목록
            TabIndex        =   3
            Top             =   405
            Visible         =   0   'False
            Width           =   2775
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Index           =   0
            Left            =   1245
            TabIndex        =   4
            Top             =   60
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   57016320
            CurrentDate     =   39244
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   5
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "검색기간"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   6
            Top             =   405
            Visible         =   0   'False
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "지 사 명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Index           =   1
            Left            =   4290
            TabIndex        =   7
            Top             =   60
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   57016320
            CurrentDate     =   39244
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "~"
            Height          =   195
            Left            =   4095
            TabIndex        =   8
            Top             =   120
            Width           =   105
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   9
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
         Caption         =   " SMS 기간별 등록 현황 (P_SMSALL_9)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_SMSALL_009.frx":0AA7
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
         TabIndex        =   10
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
         PictureBackground=   "P_SMSALL_009.frx":0CA9
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   11
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
            Picture         =   "P_SMSALL_009.frx":0EAB
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   12
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
            Picture         =   "P_SMSALL_009.frx":1445
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   13
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
            Picture         =   "P_SMSALL_009.frx":19DF
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   14
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
            Picture         =   "P_SMSALL_009.frx":1F79
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   15
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
            Picture         =   "P_SMSALL_009.frx":2513
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   16
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
            Picture         =   "P_SMSALL_009.frx":2AAD
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   17
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
            Picture         =   "P_SMSALL_009.frx":3047
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   18
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
            Picture         =   "P_SMSALL_009.frx":35E1
         End
      End
   End
End
Attribute VB_Name = "P_SMSALL_9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Dim P_SMS009_Flag As Boolean

Dim sPrintOption As String

Private Sub Data_Display()
    On Error GoTo ErrRtn
    
    ReDim sValue(4)

    sValue(0) = "0"
    sValue(1) = "%"
    sValue(2) = Format(DTPicker1(0).Value, "yyyy-MM-dd")
    sValue(3) = Format(DTPicker1(1).Value, "yyyy-MM-dd")
    ' 대리점 정보
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_SMSALL_009_01", sValue(), Err_Num, Err_Dec)

    spdView(0).MaxCols = RS01.Fields.Count
    spdView(0).MaxRows = RS01.RecordCount

    Call spdDisplay1(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView(0))
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub spdDisplay1(Rs As ADODB.Recordset)
    Call fpSpread_Display(spdView(0), Rs)
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
        Case 6: Call Export_Excel(P_00000.cdgExcel, spdView(0))      ' 엑셀
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

    On Error GoTo ErrRtn
    
    Call SubBottonEnable(cmdBtn, "10000011")
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
    If P_SMS009_Flag = False Then
        Screen.MousePointer = vbHourglass
        DTPicker1(0).Value = Format(Date, "yyyy-mm-01")
        DTPicker1(1).Value = Now
        
        If Store.Code = MASTER_OFFICE_CODE Then
            Call Master_tblComboAdd(cboInput(0))
            cboInput(0).AddItem "[1000] 본사", 1

        Else
            cboInput(0).AddItem "[" & Store.Code & "] " & Store.Name
            cboInput(0).ListIndex = 0
            cboInput(0).Enabled = False

        End If
        
        
        
        ReDim sValue(4)
    
        sValue(0) = "1"
        sValue(1) = "%"
        sValue(2) = Format(DTPicker1(0).Value, "yyyy-MM-dd")
        sValue(3) = Format(DTPicker1(1).Value, "yyyy-MM-dd")
        ' 대리점 정보
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_M_SMSALL_009_01", sValue(), Err_Num, Err_Dec)
    
        spdView(0).MaxCols = RS01.Fields.Count
        spdView(0).MaxRows = RS01.RecordCount
    
        Call spdDisplay1(RS01)
        Call GetColWidth(REG_App, Me.Name, spdView(0))
        
        P_SMS009_Flag = True
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)

End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn

    With spdView(0)
        .ColsFrozen = 1  '틀고정
        .Row = -1

        .Col = 1
        .ColWidth(1) = 8
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter

        .Col = 2
        .ColWidth(2) = 15
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
        
        .Col = 3
        .ColWidth(3) = 10
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter

        .Col = 4
        .ColWidth(4) = 5
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 0
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignRight
        
        
        .Col = 5
        .ColWidth(5) = 8
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 0
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignRight

        .Col = 6
        .ColWidth(6) = 10
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 0
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
    
        .Col = 7
        .ColWidth(7) = 25
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
    
        .Col = 8
        .ColWidth(8) = 25
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
    
        .Col = 9
        .ColWidth(9) = 25
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
    
        .Col = 10
        .ColWidth(10) = 20
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
    End With


    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_SMS009_Flag = False
End Sub



Public Sub DataCancel()

End Sub

Public Sub DataDelete()

End Sub

Public Sub DataSave()

End Sub

Public Sub DataPrint()

End Sub


Public Sub DataScreen()
'    panPrint.Visible = True

    sPrintOption = "2"
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
