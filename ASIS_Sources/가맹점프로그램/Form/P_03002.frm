VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form P_03002 
   Caption         =   "일일출고조회"
   ClientHeight    =   7860
   ClientLeft      =   1635
   ClientTop       =   1575
   ClientWidth     =   10575
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
   MDIChild        =   -1  'True
   ScaleHeight     =   7860
   ScaleWidth      =   10575
   WindowState     =   2  '최대화
   Begin Threed.SSPanel panMain 
      Align           =   1  '위 맞춤
      Height          =   9135
      Left            =   0
      TabIndex        =   0
      Top             =   435
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   16113
      _Version        =   262144
      RoundedCorners  =   0   'False
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   11
         Left            =   1800
         TabIndex        =   20
         Top             =   8340
         Width           =   13395
         _ExtentX        =   23627
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "소  품 : [#],     반  품 : [반],     택 분 실 : [T]"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin VB.TextBox txtInput 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Index           =   5
         Left            =   11880
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   8700
         Width           =   795
      End
      Begin VB.TextBox txtInput 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Index           =   4
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   8700
         Width           =   795
      End
      Begin VB.TextBox txtInput 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Index           =   3
         Left            =   14400
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   8700
         Width           =   795
      End
      Begin VB.TextBox txtInput 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Index           =   2
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   8700
         Width           =   795
      End
      Begin VB.TextBox txtInput 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Index           =   1
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   8700
         Width           =   795
      End
      Begin VB.TextBox txtInput 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Index           =   0
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   8700
         Width           =   795
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   8175
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   14865
         _Version        =   524288
         _ExtentX        =   26220
         _ExtentY        =   14420
         _StockProps     =   64
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   15
         MaxRows         =   36
         ScrollBars      =   0
         SpreadDesigner  =   "P_03002.frx":0000
         UserResize      =   1
         AppearanceStyle =   0
         CellNoteIndicatorColor=   11338536
         HighlightAlphaBlendColor=   68726040
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   8700
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "소 품 수 량"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   3
         Left            =   2700
         TabIndex        =   9
         Top             =   8700
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "재  세  탁"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   4
         Left            =   5220
         TabIndex        =   11
         Top             =   8700
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "반    품"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   5
         Left            =   7740
         TabIndex        =   13
         Top             =   8700
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "수    선"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   6
         Left            =   10260
         TabIndex        =   15
         Top             =   8700
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "택  분  실"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   7
         Left            =   12780
         TabIndex        =   17
         Top             =   8700
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "출 고 수 량"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   10
         Left            =   180
         TabIndex        =   19
         Top             =   8340
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "구      분"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
   End
   Begin Threed.SSPanel panInput 
      Align           =   1  '위 맞춤
      Height          =   435
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   767
      _Version        =   262144
      RoundedCorners  =   0   'False
      Begin MSComCtl2.DTPicker dtInput 
         Height          =   315
         Left            =   1680
         TabIndex        =   6
         Top             =   60
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   21430272
         CurrentDate     =   36686
      End
      Begin VB.ComboBox cboInput 
         Height          =   315
         Left            =   6360
         Style           =   2  '드롭다운 목록
         TabIndex        =   4
         Top             =   60
         Width           =   2775
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   0
         Left            =   4740
         TabIndex        =   3
         Top             =   60
         Width           =   1635
         _ExtentX        =   2884
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
         TabIndex        =   5
         Top             =   60
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "출 고 일 자"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
   End
End
Attribute VB_Name = "P_03002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub Form_Activate()
    P_00000.cmdBtn(0).Enabled = True
    P_00000.cmdBtn(5).Enabled = True
    P_00000.cmdBtn(6).Enabled = True
    
    P_00000.panProgramID = Me.Name
    P_00000.panProgramName = Me.Caption
    
    If P_03002_Flag = False Then
        Call AgencyComboAdd(cboInput)
    
        dtInput.Value = Date
        
        P_03002_Flag = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_00000.cmdBtn(0).Enabled = False
    P_00000.cmdBtn(1).Enabled = False
    P_00000.cmdBtn(2).Enabled = False
    P_00000.cmdBtn(3).Enabled = False
    P_00000.cmdBtn(4).Enabled = False
    P_00000.cmdBtn(5).Enabled = False
    P_00000.cmdBtn(6).Enabled = False
    
    P_00000.panProgramID = ""
    P_00000.panProgramName = ""
    
    P_03002_Flag = False
End Sub

Public Sub DataDisplay()
    Dim iCnt(5) As Integer
    Dim iCol As Integer
    Dim iRow As Integer
    
    For iRow = 1 To spdView.MaxRows
        For iCol = 1 To spdView.MaxCols
            spdView.Row = iRow
            spdView.Col = iCol
            spdView.Text = ""
        Next iCol
    Next iRow
    
    ReDim sValue(2)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput.Value, "yyyymmdd")
    sValue(2) = Mid(cboInput.Text, 2, 3)
        
    Set rs01 = New ADODB.Recordset
    Set rs01 = ExecPro("PRO_P_03002_00", sValue(), Err_Num, Err_Dec)
    
    iRow = 1
    iCol = 0
    
    While Not rs01.EOF
        '재세탁 Count
        iCnt(0) = iCnt(0) + rs01!재세탁수량
        
        '수선 Count
        iCnt(1) = iCnt(1) + rs01!수선수량
        
        '품목구분 (소품여부)
        iCnt(2) = iCnt(2) + rs01!수량
        
        iCnt(3) = iCnt(3) + 1
        
        '출고구분(반품)
        iCnt(4) = iCnt(4) + rs01!반품수량
        
        'Lost Tag
        iCnt(5) = iCnt(5) + rs01!LOST수량
        
        iCol = iCol + 1
        
        If iCol > 15 Then
            iCol = 1
            iRow = iRow + 1
        End If
        
        spdView.Row = iRow
        spdView.Col = iCol
        spdView.Text = rs01!택
            
        rs01.MoveNext
    Wend
    
    txtInput(0).Text = Format(iCnt(0), "#,##0")
    txtInput(1).Text = Format(iCnt(1), "#,##0")
    txtInput(2).Text = Format(iCnt(2), "#,##0")
    txtInput(3).Text = Format(iCnt(3), "#,##0")
    txtInput(4).Text = Format(iCnt(4), "#,##0")
    txtInput(5).Text = Format(iCnt(5), "#,##0")
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
    Dim SseekData(2) As Integer
    
    Dim bLIneEnd As String

    Printer.PaperSize = vbPRPSA4
    Printer.Orientation = 1
    memRow = 1

PrintHead:

    Printer.Font = "굴림체"                             ' Printer의 사용 글자
    Printer.FontSize = "16"                             ' Print의 글자크기
    Printer.ScaleMode = vbMillimeters                   ' Print의 위치 선정을 밀리미터로 나타낸다.
    iRow = iRow + 2
    Printer.CurrentY = iRow
    Printer.CurrentX = 70
    Printer.Print Space(10); Me.Caption

    Printer.Font = "굴림체"                             ' Printer의 사용 글자
    Printer.FontSize = "10"                             ' Print의 글자크기
    Printer.ScaleMode = vbMillimeters                   ' Print의 위치 선정을 밀리미터로 나타낸다.
    iRow = iRow + 14
    Printer.CurrentY = iRow
    Printer.CurrentX = 0
    Printer.Print Space(10); "(주)백상"

    iRow = iRow + 5
    Printer.CurrentY = iRow
    Printer.CurrentX = 0
    sData = "출고일자 : " & dtInput.Value & Space(20)
    Printer.Print Space(10); sData

    iRow = iRow + 5
    Printer.CurrentY = iRow
    Printer.CurrentX = 0
    sData = ""
    sData = "대리점명 : " & cboInput.Text & Space(20)
    sData = LeftH(RTrim(sData) & Space(80), 80) & "출력일자 : " & Format(Now, "YYYY-MM-DD")
    Printer.Print Space(10); sData

    iRow = iRow + 5
    Printer.CurrentY = iRow
    Printer.CurrentX = 0
    sData = ""
''    sData = "품 목 명 : " & Trim(cboInput(0).Text) & " ~ " & Trim(cboInput(1).Text) & Space(20)
    sData = LeftH(RTrim(sData) & Space(80), 80) & "출력시간 : " & Format(Now, "hh:mm:ss")
    Printer.Print Space(10); sData
    
    iRow = iRow + 1
    Printer.Line (10, iRow + 3)-(260, iRow + 3)

''    iRow = iRow + 4
''    Printer.CurrentY = iRow
''    Printer.CurrentX = 0
''    Printer.Print "   대리점                     출고수량"
    
''    iRow = iRow + 1
''    Printer.Line (0, iRow + 3)-(240, iRow + 3)
    
    bLIneEnd = "1"
    sData = ""

    For i = memRow To spdView.MaxRows - 1
        spdView.Row = i

        spdView.Col = 1
        sData = sData & Space(1) & Right(Space(7) & spdView.Text, 7) & " "

        spdView.Col = 2
        sData = sData & Space(1) & Right(Space(7) & Trim(spdView.Text), 7) & " "
        
        spdView.Col = 3
        sData = sData & Space(1) & Right(Space(7) & Trim(spdView.Text), 7) & " "

        spdView.Col = 4
        sData = sData & Space(1) & Right(Space(7) & Trim(spdView.Text), 7) & " "
        
        spdView.Col = 5
        sData = sData & Space(1) & Right(Space(7) & Trim(spdView.Text), 7) & " "
        
        If bLIneEnd = "2" Then
            iRow = iRow + 4
            Printer.CurrentY = iRow
            Printer.CurrentX = 0
            Printer.Print Space(10); sData
            
            sData = ""
            
            bLIneEnd = "3"
        End If
    

        spdView.Col = 6
        sData = sData & Space(1) & Right(Space(7) & Trim(spdView.Text), 7) & " "
        
        spdView.Col = 7
        sData = sData & Space(1) & Right(Space(7) & Trim(spdView.Text), 7) & " "

        spdView.Col = 8
        sData = sData & Space(1) & Right(Space(7) & Trim(spdView.Text), 7) & " "
        
        spdView.Col = 9
        sData = sData & Space(1) & Right(Space(7) & Trim(spdView.Text), 7) & " "

        spdView.Col = 10
        sData = sData & Space(1) & Right(Space(7) & Trim(spdView.Text), 7) & " "
        
        If bLIneEnd = "1" Then
            iRow = iRow + 4
            Printer.CurrentY = iRow
            Printer.CurrentX = 0
            Printer.Print Space(10); sData
            
            sData = ""
            
            bLIneEnd = "2"
        End If
        
        spdView.Col = 11
        sData = sData & Space(1) & Right(Space(7) & Trim(spdView.Text), 7) & " "
        
        spdView.Col = 12
        sData = sData & Space(1) & Right(Space(7) & Trim(spdView.Text), 7) & " "
        
        spdView.Col = 13
        sData = sData & Space(1) & Right(Space(7) & Trim(spdView.Text), 7) & " "
        
        spdView.Col = 14
        sData = sData & Space(1) & Right(Space(7) & Trim(spdView.Text), 7) & " "
        
        spdView.Col = 15
        sData = sData & Space(1) & Right(Space(7) & Trim(spdView.Text), 7) & " "
        
        If bLIneEnd = "3" Then
            iRow = iRow + 4
            Printer.CurrentY = iRow
            Printer.CurrentX = 0
            Printer.Print Space(10); sData
            
            sData = ""
            
            bLIneEnd = "1"
        End If

        If iRow > 270 Then
            iRow = iRow + 1
            Printer.Line (10, iRow + 3)-(260, iRow + 3)

            memRow = i + 1
            iRow = 0

            Printer.NewPage
            GoTo PrintHead
        End If
    Next i
    
    iRow = iRow + 1
    Printer.Line (10, iRow + 3)-(260, iRow + 3)
    
''    sData = "총   수   량"
''    sData = sData & Space(12) & Right(Space(10) & SseekData(0), 10) & " "
    
''    iRow = iRow + 4
''    Printer.CurrentY = iRow
''    Printer.CurrentX = 0
''    Printer.Print sData
    
    iRow = iRow + 5
    Printer.CurrentY = iRow
    Printer.CurrentX = 0
    sData = ""
    sData = "구    분 : 소  품 [#], 반  품 [반], 택분실 [T]"
    Printer.Print Space(10); sData
    
    iRow = iRow + 5
    Printer.CurrentY = iRow
    Printer.CurrentX = 0
    sData = ""
    sData = "소품수량 : " & txtInput(2).Text & Space(20)
    sData = LeftH(RTrim(sData) & Space(40), 40) & "재 세 탁 : " & txtInput(0).Text
    sData = LeftH(RTrim(sData) & Space(80), 80) & "반    품 : " & txtInput(4).Text
    Printer.Print Space(10); sData
    
    iRow = iRow + 5
    Printer.CurrentY = iRow
    Printer.CurrentX = 0
    sData = ""
    sData = "수    선 : " & txtInput(1).Text & Space(20)
    sData = LeftH(RTrim(sData) & Space(40), 40) & "택 분 실 : " & txtInput(5).Text
    sData = LeftH(RTrim(sData) & Space(80), 80) & "출고수량 : " & txtInput(3).Text
    Printer.Print Space(10); sData
        
    iRow = iRow + 1
    Printer.Line (10, iRow + 3)-(260, iRow + 3)

    Printer.EndDoc
End Sub

Private Sub spdView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu P_00000.PopUp
    End If
End Sub

Public Sub DataScreen()
    Dim ReportFP As String
    Dim ReportFile As String
    
    ReportFP = GetIniStr("REPORT", "FilePath", "", sIniFile)
    ReportFile = ReportFP & "\" & "P_03001.rpt"
    
    P_00000.crPrint.ReportFileName = ReportFile
    P_00000.crPrint.SelectionFormula = "({ChulgoTag;1.AgencyName} = '" & Mid(cboInput.Text, 2, 3) & "') And {ChulgoTag;1.SendChk} = '0' "
    P_00000.crPrint.StoredProcParam(0) = Format(dtInput.Value, "yyyymmdd")
    P_00000.crPrint.StoredProcParam(1) = "***"
    
    P_SCREEN.Show
End Sub
