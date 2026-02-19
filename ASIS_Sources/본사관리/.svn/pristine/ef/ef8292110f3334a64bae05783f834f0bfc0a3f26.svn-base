VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_03014 
   Caption         =   "판매취소 입,출 전환작업"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   1860
   ClientWidth     =   15855
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_03014.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8805
   ScaleWidth      =   15855
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   8805
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15855
      _ExtentX        =   27966
      _ExtentY        =   15531
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_03014.frx":058A
      Begin Threed.SSPanel SSPanel1 
         Height          =   405
         Left            =   15
         TabIndex        =   2
         Top             =   8385
         Width           =   15825
         _ExtentX        =   27914
         _ExtentY        =   714
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            TabIndex        =   4
            Top             =   60
            Width           =   1335
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   1
            Left            =   4530
            TabIndex        =   3
            Top             =   60
            Width           =   1335
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   60
            TabIndex        =   5
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "총  점  수"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   3
            Left            =   3060
            TabIndex        =   6
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "금    액"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   7035
         Left            =   15
         TabIndex        =   1
         Top             =   1335
         Width           =   15825
         _Version        =   524288
         _ExtentX        =   27914
         _ExtentY        =   12409
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
         SpreadDesigner  =   "P_03014.frx":063C
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   7
         Top             =   540
         Width           =   15825
         _ExtentX        =   27914
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            Style           =   2  '드롭다운 목록
            TabIndex        =   8
            Top             =   420
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   4770
            TabIndex        =   9
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   67567616
            CurrentDate     =   36686
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            TabIndex        =   10
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   67567616
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   11
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
            TabIndex        =   12
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "출 고 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin VB.Label Label1 
            Alignment       =   2  '가운데 맞춤
            Caption         =   "~"
            Height          =   195
            Left            =   4530
            TabIndex        =   13
            Top             =   120
            Width           =   255
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   14
         Top             =   15
         Width           =   8220
         _ExtentX        =   14499
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
         PictureBackground=   "P_03014.frx":0AC7
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   0
         Left            =   8250
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
         PictureBackground=   "P_03014.frx":0CC9
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
            Picture         =   "P_03014.frx":0ECB
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
            Picture         =   "P_03014.frx":1465
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
            Picture         =   "P_03014.frx":19FF
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
            Picture         =   "P_03014.frx":1F99
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
            Picture         =   "P_03014.frx":2533
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
            Picture         =   "P_03014.frx":2ACD
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
            Picture         =   "P_03014.frx":3067
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
            Picture         =   "P_03014.frx":3601
         End
      End
   End
End
Attribute VB_Name = "P_03014"
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
    Call Data_Display
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: Call Data_Display   ' 조회
        Case 1: 'Call DataAdd        ' 신규
        Case 2: Call DataSave       ' 저장
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

Private Sub dtInput_Change(Index As Integer)
'    Call Data_Display
End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(2).Enabled = True
    cmdBtn(5).Enabled = True
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    If P_03014_Flag = False Then
        Call AgencyComboAdd(cboInput(0))
        
        dtInput(0).Value = Date
        dtInput(1).Value = Date
        
        ReDim sValue(5)
        
        sValue(0) = "1"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_03014_00", sValue(), Err_Num, Err_Dec)
        
        spdView.MaxCols = RS01.Fields.Count
        spdView.MaxRows = RS01.RecordCount
        
        Call spdDisplay(RS01)
        Call GetColWidth(REG_App, Me.Name, spdView)
        
        P_03014_Flag = True
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
    spdView.ColWidth(1) = 10
    spdView.CellType = CellTypeDate
    spdView.TypeDateCentury = True
    spdView.TypeDateFormat = TypeDateFormatYYMMDD
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter
    
    spdView.Col = 2
    spdView.ColWidth(2) = 10
    spdView.CellType = CellTypeDate
    spdView.TypeDateCentury = True
    spdView.TypeDateFormat = TypeDateFormatYYMMDD
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter
    
    spdView.Col = 3
    spdView.ColWidth(3) = 6
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter
    
    spdView.Col = 4
    spdView.ColWidth(4) = 10
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft

    spdView.Col = 5
    spdView.ColWidth(5) = 20
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft

    spdView.Col = 6
    spdView.ColWidth(6) = 8
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft

    spdView.Col = 7
    spdView.ColWidth(7) = 12
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 8
    spdView.ColWidth(8) = 15
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft

    spdView.Col = 9
    spdView.ColWidth(9) = 15
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft

    spdView.Col = 10
    spdView.ColWidth(10) = 8
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_03014_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    ReDim sValue(3)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    sValue(3) = Mid(cboInput(0).Text, 2, 3)
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_03014_00", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    Call spdDisplay(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView)
    
    spdView.AutoCalc = True
    
    spdView.MaxRows = spdView.MaxRows + 1
    spdView.Row = spdView.MaxRows
    spdView.RowHidden = True
    
    spdView.Col = 8: spdView.Formula = "SUM(H1:H" & spdView.MaxRows - 1 & ")"
    
    txtInput(0).Text = spdView.MaxRows - 1
    txtInput(1).Text = spdView.Text
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Public Sub DataPrint()
    Dim sData As String
    Dim i, ii, iii As Integer
    Dim iRow As Integer
    Dim memRow As Long
    Dim lLineQty As Long
    Dim lLinePri As Double
    Dim lLineAmt As Double
    Dim lTotalQty As Long
    Dim lTotalPri As Double
    Dim lTotalAmt As Double
    Dim lTotalVAT As Double
    Dim sPrintData As String
    Dim Pum_Code As String
    Dim SseekData(2) As Double
    
    Printer.PaperSize = vbPRPSA4
    memRow = 1

PrintHead:

    Printer.Font = "굴림체"                             ' Printer의 사용 글자
    Printer.FontSize = "16"                             ' Print의 글자크기
    Printer.ScaleMode = vbMillimeters                   ' Print의 위치 선정을 밀리미터로 나타낸다.
    iRow = iRow + 2
    Printer.CurrentY = iRow
    Printer.CurrentX = 75
    Printer.Print Me.Caption

    Printer.Font = "굴림체"                             ' Printer의 사용 글자
    Printer.FontSize = "10"                             ' Print의 글자크기
    Printer.ScaleMode = vbMillimeters                   ' Print의 위치 선정을 밀리미터로 나타낸다.
    iRow = iRow + 14
    Printer.CurrentY = iRow
    Printer.CurrentX = 0
    Printer.Print "(주)백상"

    iRow = iRow + 5
    Printer.CurrentY = iRow
    Printer.CurrentX = 0
    sData = ""
    sData = "검색일자 : " & dtInput(0).Value & " ~ " & dtInput(1).Value & Space(20)
    sData = LeftH(RTrim(sData) & Space(65), 65) & "성  명 : " & USERNAME
    Printer.Print sData

    iRow = iRow + 5
    Printer.CurrentY = iRow
    Printer.CurrentX = 0
    sData = ""
    sData = "대리점명 : " & cboInput(0).Text & Space(20)
    sData = LeftH(RTrim(sData) & Space(65), 65) & "출력일자 : " & Format(Now, "YYYY-MM-DD")
    Printer.Print sData

    iRow = iRow + 5
    Printer.CurrentY = iRow
    Printer.CurrentX = 0
    sData = ""
    sData = "품 목 명 : " & cboInput(1).Text & Space(20)
    sData = LeftH(RTrim(sData) & Space(65), 65) & "출력시간 : " & Format(Now, "hh:mm:ss")
    Printer.Print sData
    
    iRow = iRow + 1
    Printer.Line (0, iRow + 3)-(240, iRow + 3)

    iRow = iRow + 4
    Printer.CurrentY = iRow
    Printer.CurrentX = 0
    Printer.Print "       대리점             출고일자    입고일자    택번호 전화번호       품목           금액    색상  내용  상표"
    
    iRow = iRow + 1
    Printer.Line (0, iRow + 3)-(240, iRow + 3)

    For i = memRow To spdView.MaxRows - 1
        spdView.Row = i

        spdView.Col = 1
        sData = LeftH(spdView.Text & Space(15), 15)                                           '대리점
        
        spdView.Col = 2
        sData = sData & LeftH(spdView.Text & Space(6), 6)                                     '대리점

        spdView.Col = 3
        sData = sData & Space(2) & LeftH(spdView.Text, 10)                                    '출고일

        spdView.Col = 4
        sData = sData & Space(2) & LeftH(spdView.Text, 10)                                    '입고일

        spdView.Col = 5
        sData = sData & Space(2) & LeftH(spdView.Text, 8)                                    '택번호

        spdView.Col = 6
        sData = sData & Space(2) & LeftH(spdView.Text, 10)                                    '전화번호
        
        spdView.Col = 7
        sData = sData & Space(2) & LeftH(spdView.Text, 20)                                    '품목
                
        spdView.Col = 8
        'sData = sData & Space(2) & RightH(Space(10) & spdView.Text, 10) & " "                '금액
        sData = sData & Space(2) & RightH(spdView.Text, 10) & " "

        spdView.Col = 9
        sData = sData & Space(2) & LeftH(spdView.Text, 6)                                     '색상
        
        spdView.Col = 10
        sData = sData & Space(2) & LeftH(spdView.Text, 6)                                     '내용
        
        spdView.Col = 11
        sData = sData & Space(2) & LeftH(spdView.Text, 10)                                    '상표
        
        iRow = iRow + 4
        Printer.CurrentY = iRow
        Printer.CurrentX = 0
        Printer.Print sData

        If iRow > 270 Then
            iRow = iRow + 1
            Printer.Line (0, iRow + 3)-(240, iRow + 3)

            memRow = i + 1
            iRow = 0

            Printer.NewPage
            GoTo PrintHead
        End If
    Next i
    
    iRow = iRow + 1
    Printer.Line (0, iRow + 3)-(240, iRow + 3)
    
    sData = ""
    sData = "총 점 수 : " & txtInput(0).Text & Space(5)
    sData = LeftH(RTrim(sData) & Space(65), 65) & "금    액 : " & txtInput(1).Text
    iRow = iRow + 4
    Printer.CurrentY = iRow
    Printer.CurrentX = 0
    Printer.Print sData
    
''    sData = "총   수   량"
''    sData = sData & Space(39) & Right(Space(10) & SseekData(0), 10) & " "
    
''    iRow = iRow + 4
''    Printer.CurrentY = iRow
''    Printer.CurrentX = 0
''    Printer.Print sData
    
''    iRow = iRow + 1
''    Printer.Line (0, iRow + 3)-(240, iRow + 3)

    Printer.EndDoc
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
        'PopupMenu P_00000.PopUp
    End If
End Sub

Public Sub DataSave()
    Dim i As Integer
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        spdView.Col = 8
        If spdView.Value <> "" Then
            If spdView.Value = True Then
                ReDim sValue(4)
                
                spdView.Col = 1: sValue(0) = Format(spdView.Text, "YYYY-MM-DD")
                spdView.Col = 2: sValue(1) = spdView.Text
                spdView.Col = 2: sValue(2) = spdView.Text
                
                Call ExecPro("SP_03014_01", sValue(), Err_Num, Err_Dec)
            End If
        End If
    Next i
    
    If Err_Num = 0 Then
        MsgBox "해당되는 데이터가 정상적으로 저장이 되었습니다.", vbInformation
        Exit Sub
    End If
End Sub
