VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form P_03010_02 
   Caption         =   "가출고 관리"
   ClientHeight    =   11175
   ClientLeft      =   -600
   ClientTop       =   2340
   ClientWidth     =   16530
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
   ScaleHeight     =   11175
   ScaleWidth      =   16530
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   11175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16530
      _ExtentX        =   29157
      _ExtentY        =   19711
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_03010_02.frx":0000
      Begin Threed.SSPanel SSPanel 
         Height          =   435
         Left            =   4365
         TabIndex        =   6
         Top             =   10725
         Width           =   12150
         _ExtentX        =   21431
         _ExtentY        =   767
         _Version        =   262144
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   0
            Left            =   1515
            TabIndex        =   9
            Top             =   60
            Width           =   1455
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   1
            Left            =   4695
            TabIndex        =   8
            Top             =   60
            Width           =   1455
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   2
            Left            =   7875
            TabIndex        =   7
            Top             =   60
            Width           =   1455
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   45
            TabIndex        =   10
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "검 품 수 량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   3225
            TabIndex        =   11
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "입 고 수 량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   3
            Left            =   6405
            TabIndex        =   12
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "다 른 품 목"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel panInput 
         Height          =   765
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   16500
         _ExtentX        =   29104
         _ExtentY        =   1349
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            TabIndex        =   2
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   54722560
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   3
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "입고일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   4710
            TabIndex        =   4
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   54722560
            CurrentDate     =   36686
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10365
         Index           =   0
         Left            =   15
         TabIndex        =   5
         Top             =   795
         Width           =   4335
         _Version        =   524288
         _ExtentX        =   7646
         _ExtentY        =   18283
         _StockProps     =   64
         BackColorStyle  =   1
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
         SpreadDesigner  =   "P_03010_02.frx":0092
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   9915
         Index           =   1
         Left            =   4365
         TabIndex        =   13
         Top             =   795
         Width           =   12150
         _Version        =   524288
         _ExtentX        =   21431
         _ExtentY        =   17489
         _StockProps     =   64
         BackColorStyle  =   1
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
         SpreadDesigner  =   "P_03010_02.frx":0552
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_03010_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub dtInput_Change(Index As Integer)
    Call Data_Display
End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(5).Enabled = True
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
    
    dtInput(0).Value = Date
    dtInput(1).Value = Date
    
    ReDim sValue(2)
    
    sValue(0) = "1"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_03010_00", sValue(), Err_Num, Err_Dec)
    
    spdView(0).MaxCols = RS01.Fields.Count
    spdView(0).MaxRows = RS01.RecordCount
    
    Call spdDisplay(RS01)
    Call GetColWidth("백상", Me.Name & "A", spdView(0))

    ReDim sValue(3)
    
    sValue(0) = "1"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_03010_01", sValue(), Err_Num, Err_Dec)

    spdView(1).MaxCols = RS01.Fields.Count
    spdView(1).MaxRows = RS01.RecordCount

    Call spdDisplay2(RS01)
    Call GetColWidth("백상", Me.Name & "B", spdView(1))
End Sub

Private Sub spdDisplay(Rs As ADODB.Recordset)
    Set spdView(0).DataSource = Rs
    
    spdView(0).ColsFrozen = 1 '틀고정
    
    spdView(0).Row = -1
    
    spdView(0).Col = 1
    spdView(0).CellType = CellTypeEdit
    spdView(0).TypeVAlign = TypeVAlignCenter
    spdView(0).TypeHAlign = TypeHAlignLeft
    
    spdView(0).Col = 2
    spdView(0).CellType = CellTypeEdit
    spdView(0).TypeVAlign = TypeVAlignCenter
    spdView(0).TypeHAlign = TypeHAlignCenter

    spdView(0).Col = 3
    spdView(0).CellType = CellTypeFloat
    spdView(0).TypeFloatSeparator = True
    spdView(0).TypeFloatDecimalPlaces = 0
    spdView(0).TypeVAlign = TypeVAlignCenter
    
    spdView(0).Col = 4
    spdView(0).CellType = CellTypeEdit
    spdView(0).TypeVAlign = TypeVAlignCenter
    spdView(0).TypeHAlign = TypeHAlignCenter

    spdView(0).Col = 5
    spdView(0).CellType = CellTypeEdit
    spdView(0).TypeVAlign = TypeVAlignCenter
    spdView(0).TypeHAlign = TypeHAlignCenter
End Sub

Private Sub spdDisplay2(Rs As ADODB.Recordset)
    Set spdView(1).DataSource = Rs
    
    spdView(1).ColsFrozen = 1 '틀고정
    
    spdView(1).Row = -1
    
    spdView(1).Col = 1
    spdView(1).CellType = CellTypeEdit
    spdView(1).TypeVAlign = TypeVAlignCenter
    spdView(1).TypeHAlign = TypeHAlignCenter
    
    spdView(1).Col = 2
    spdView(1).CellType = CellTypeEdit
    spdView(1).TypeVAlign = TypeVAlignCenter
    spdView(1).TypeHAlign = TypeHAlignLeft
    
    spdView(1).Col = 3
    spdView(1).CellType = CellTypeEdit
    spdView(1).TypeVAlign = TypeVAlignCenter
    spdView(1).TypeHAlign = TypeHAlignCenter
    
    spdView(1).Col = 4
    spdView(1).CellType = CellTypeEdit
    spdView(1).TypeVAlign = TypeVAlignCenter
    spdView(1).TypeHAlign = TypeHAlignLeft
    
    spdView(1).Col = 5
    spdView(1).CellType = CellTypeFloat
    spdView(1).TypeFloatSeparator = True
    spdView(1).TypeFloatDecimalPlaces = 0
    spdView(1).TypeVAlign = TypeVAlignCenter
    
    spdView(1).Col = 6
    spdView(1).CellType = CellTypeFloat
    spdView(1).TypeFloatSeparator = True
    spdView(1).TypeFloatDecimalPlaces = 0
    spdView(1).TypeVAlign = TypeVAlignCenter

    spdView(1).Col = 7
    spdView(1).CellType = CellTypeFloat
    spdView(1).TypeFloatSeparator = True
    spdView(1).TypeFloatDecimalPlaces = 0
    spdView(1).TypeVAlign = TypeVAlignCenter
    
    spdView(1).Col = 8
    spdView(1).CellType = CellTypeFloat
    spdView(1).TypeFloatSeparator = True
    spdView(1).TypeFloatDecimalPlaces = 0
    spdView(1).TypeVAlign = TypeVAlignCenter

    spdView(1).Col = 9
    spdView(1).CellType = CellTypeEdit
    spdView(1).TypeVAlign = TypeVAlignCenter
    spdView(1).TypeHAlign = TypeHAlignLeft
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveColWidth("백상", Me.Name & "A", spdView(0))
    Call SaveColWidth("백상", Me.Name & "B", spdView(1))
End Sub

Public Sub Data_Display()
    Dim i As Integer
    Dim lAmt As Long
    
    ReDim sValue(2)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
        
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_03010_00", sValue(), Err_Num, Err_Dec)
    
    spdView(0).MaxCols = RS01.Fields.Count
    spdView(0).MaxRows = RS01.RecordCount
    
    Call spdDisplay(RS01)
    Call GetColWidth("백상", Me.Name, spdView(0))
End Sub

Private Sub spdView_DblClick(Index As Integer, ByVal Col As Long, ByVal Row As Long)
    If Index = 0 Then
        ReDim sValue(3)
        
        sValue(0) = "0"
        
        spdView(0).Row = Row
        spdView(0).Col = 1
        sValue(1) = Mid(spdView(0).Text, 2, 3)
        
        sValue(2) = Format(dtInput(0).Value, "YYYY-MM-DD")
        sValue(3) = Format(dtInput(0).Value, "YYYY-MM-DD")
        
        spdView(0).Row = Row
        spdView(0).Col = 1
        
        sValue(2) = Mid(spdView(0).Text, 2, 3)
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_03010_01", sValue(), Err_Num, Err_Dec)
        
        spdView(1).MaxCols = RS01.Fields.Count
        spdView(1).MaxRows = RS01.RecordCount
        
        Call spdDisplay2(RS01)
        Call GetColWidth("백상", Me.Name & "B", spdView(1))
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
    sData = "입고일자 : " & dtInput(0).Value & " ~ " & dtInput(1).Value & Space(20)
    sData = LeftH(RTrim(sData) & Space(65), 65) & "성  명 : " & USERNAME
    Printer.Print sData

    iRow = iRow + 5
    Printer.CurrentY = iRow
    Printer.CurrentX = 0
    spdView(0).Row = spdView(0).ActiveRow
    sData = ""
    spdView(0).Col = 1
    sData = "대리점명 : " & Left(spdView(0).Text & Space(25), 25)
    sData = LeftH(RTrim(sData) & Space(65), 65) & "출력일자 : " & Format(Now, "YYYY-MM-DD")
    Printer.Print sData

    iRow = iRow + 5
    Printer.CurrentY = iRow
    Printer.CurrentX = 0
    spdView(0).Row = spdView(0).ActiveRow
    sData = ""
    spdView(0).Col = 2
    sData = "가 출 고 : " & Left(spdView(0).Text & Space(25), 25)
    sData = LeftH(RTrim(sData) & Space(65), 65) & "출력시간 : " & Format(Now, "hh:mm:ss")
    Printer.Print sData
    
    iRow = iRow + 5
    Printer.CurrentY = iRow
    Printer.CurrentX = 0
    spdView(0).Row = spdView(0).ActiveRow
    sData = ""
    spdView(0).Col = 3
    sData = "입    고 : " & Left(spdView(0).Text & Space(25), 25)
    spdView(0).Col = 4
    sData = LeftH(RTrim(sData) & Space(65), 65) & "미 입 고 : " & spdView(0).Text
    Printer.Print sData
    
    iRow = iRow + 1
    Printer.Line (0, iRow + 3)-(240, iRow + 3)

    iRow = iRow + 4
    Printer.CurrentY = iRow
    Printer.CurrentX = 0
    Printer.Print "택번호   출고일자   출고구분  입고일자    품목      금액   내용    색상       전화번호"
    
    iRow = iRow + 1
    Printer.Line (0, iRow + 3)-(240, iRow + 3)

    For i = memRow To spdView(1).MaxRows
        spdView(1).Row = i

        spdView(1).Col = 1
        sData = Space(1) & Left(spdView(1).Text & Space(6), 6)                                          '택번호

        spdView(1).Col = 2
        sData = sData & Space(2) & Left(spdView(1).Text & Space(10), 10)                                '출고일자

        spdView(1).Col = 3
        sData = sData & Space(3) & Left(spdView(1).Text & Space(5), 5)                                  '출고구분

        spdView(1).Col = 4
        sData = sData & Space(1) & Left(spdView(1).Text & Space(10), 10)                                '입고일자

        spdView(1).Col = 5
        sData = sData & Space(1) & Right(Space(4) & spdView(1).Text, 4) & " "                            '품목

        spdView(1).Col = 6
        sData = sData & Space(1) & Right(Space(9) & spdView(1).Text, 9) & " "                         '금액
                
        spdView(1).Col = 7
        sData = sData & Space(4) & Left(spdView(1).Text & Space(6), 6)                                  '내용

        spdView(1).Col = 8
        sData = sData & Space(2) & Left(spdView(1).Text & Space(6), 6)                                  '색상

        spdView(1).Col = 9
        sData = sData & Space(2) & Left(spdView(1).Text & Space(10), 10)                                '전화번호
        
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
    
''    sData = "총   수   량"
''    sData = sData & Space(39) & Right(Space(10) & SseekData(0), 10) & " "
    
''    iRow = iRow + 4
''    Printer.CurrentY = iRow
''    Printer.CurrentX = 0
''    Printer.Print sData
    
    iRow = iRow + 1
    Printer.Line (0, iRow + 3)-(240, iRow + 3)

    Printer.EndDoc
End Sub

Private Sub spdView_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        'PopupMenu P_00000.PopUp
    End If
End Sub
