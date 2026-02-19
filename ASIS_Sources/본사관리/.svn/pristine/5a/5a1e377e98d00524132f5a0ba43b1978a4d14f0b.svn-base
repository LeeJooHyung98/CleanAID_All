VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_02009 
   Caption         =   "검품자료 삭제"
   ClientHeight    =   8205
   ClientLeft      =   1275
   ClientTop       =   1920
   ClientWidth     =   12855
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_02009.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8205
   ScaleWidth      =   12855
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   8205
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   14473
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   0
      PaneTree        =   "P_02009.frx":058A
      Begin Threed.SSPanel SSPanel 
         Height          =   6885
         Index           =   1
         Left            =   0
         TabIndex        =   14
         Top             =   1320
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   12144
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   5250
         _ExtentX        =   9260
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
         PictureBackground=   "P_02009.frx":061C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   0
         Left            =   5265
         TabIndex        =   2
         Top             =   0
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
         PictureBackground=   "P_02009.frx":081E
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   3
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
            Picture         =   "P_02009.frx":0A20
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   4
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
            Picture         =   "P_02009.frx":0FBA
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   5
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
            Picture         =   "P_02009.frx":1554
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   6
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
            Picture         =   "P_02009.frx":1AEE
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   7
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
            Picture         =   "P_02009.frx":2088
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   8
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
            Picture         =   "P_02009.frx":2622
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   9
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
            Picture         =   "P_02009.frx":2BBC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   10
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
            Picture         =   "P_02009.frx":3156
         End
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   0
         TabIndex        =   11
         Top             =   525
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            TabIndex        =   12
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   63307776
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   13
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "입 고 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "P_02009"
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
        Case 0: 'Call Data_Display   ' 조회
        Case 1: 'Call DataAdd        ' 신규
        Case 2: 'Call DataSave       ' 저장
        Case 3: Call DataDelete     ' 삭제
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
    cmdBtn(3).Enabled = True
    cmdBtn(5).Enabled = True
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
    If P_02009_Flag = False Then
        dtInput(0).Value = Date
        
        P_02009_Flag = True
    End If
End Sub

Public Sub DataDelete()
    If MsgBox(Format(dtInput(0).Value, "YYYY-MM-DD") & "까지 이전 데이터는 모두 삭제 됩니다.", vbYesNo + vbQuestion + vbDefaultButton2, "데이터삭제") = vbYes Then
        ReDim sValue(1)
        
        sValue(0) = Format(dtInput(0).Value, "YYYY-MM-DD")
        
        Call ExecPro("SP_02009_00", sValue(), Err_Num, Err_Dec)
        
        If Err_Num = 0 Then
            MsgBox "해당되는 데이터가 정상적으로 삭제되었습니다.", vbInformation, "데이터 삭제"
            Exit Sub
        End If
    End If
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
    sData = "입고일자 : " & dtInput(0).Value & Space(20)
    sData = LeftH(RTrim(sData) & Space(65), 65) & "성  명 : " & USERNAME
    Printer.Print sData

    iRow = iRow + 5
    Printer.CurrentY = iRow
    Printer.CurrentX = 0
    sData = ""
''    sData = "미 전 송 : " & txtInput(1).Text & Space(20)
    sData = LeftH(RTrim(sData) & Space(65), 65) & "출력일자 : " & Format(Now, "YYYY-MM-DD")
    Printer.Print sData

    iRow = iRow + 5
    Printer.CurrentY = iRow
    Printer.CurrentX = 0
    sData = ""
''    sData = "전    송 : " & txtInput(2).Text & Space(20)
    sData = LeftH(RTrim(sData) & Space(65), 65) & "출력시간 : " & Format(Now, "hh:mm:ss")
    Printer.Print sData

    iRow = iRow + 1
    Printer.Line (0, iRow + 3)-(240, iRow + 3)

    iRow = iRow + 4
    Printer.CurrentY = iRow
    Printer.CurrentX = 0
    Printer.Print ""
    
    iRow = iRow + 1
    Printer.Line (0, iRow + 3)-(240, iRow + 3)

''    For i = memRow To spdView.MaxRows - 1
''        spdView.Row = i
''
''        spdView.Col = 1
''        sData = Left(spdView.Text, 10)                                             '대리점
''
''        spdView.Col = 2
''        sData = sData & Left(spdView.Text, 25)                                     '전일
''
''        spdView.Col = 3
''        sData = sData & Left(spdView.Text, 25)                                     '점수
''
''        spdView.Col = 4
''        sData = sData & Left(spdView.Text, 25)                                     '시작
''
''        spdView.Col = 5
''        sData = sData & Left(spdView.Text, 25)                                     '종료
''
''        iRow = iRow + 4
''        Printer.CurrentY = iRow
''        Printer.CurrentX = 0
''        Printer.Print sData
''
''        If iRow > 270 Then
''            iRow = iRow + 1
''            Printer.Line (0, iRow + 3)-(240, iRow + 3)
''
''            memRow = i + 1
''            iRow = 0
''
''            Printer.NewPage
''            GoTo PrintHead
''        End If
''    Next i

    iRow = iRow + 1
    Printer.Line (0, iRow + 3)-(240, iRow + 3)

    iRow = iRow + 1
    Printer.Line (0, iRow + 3)-(240, iRow + 3)

    Printer.EndDoc
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_02009_Flag = False
End Sub
