VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm본사출고현황2 
   Caption         =   "본사출고현황"
   ClientHeight    =   7875
   ClientLeft      =   300
   ClientTop       =   885
   ClientWidth     =   12600
   ControlBox      =   0   'False
   LinkTopic       =   "Form28"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7875
   ScaleWidth      =   12600
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7875
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12600
      _ExtentX        =   22225
      _ExtentY        =   13891
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm본사출고현황2.frx":0000
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   6150
         Left            =   15
         TabIndex        =   1
         Top             =   1065
         Width           =   12570
         _Version        =   524288
         _ExtentX        =   22172
         _ExtentY        =   10848
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowUserFormulas=   -1  'True
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
         DInformActiveRowChange=   0   'False
         DisplayColHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   10
         MaxRows         =   300
         MoveActiveOnFocus=   0   'False
         OperationMode   =   1
         Protect         =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frm본사출고현황2.frx":0092
         VisibleCols     =   9
         VisibleRows     =   50
         AppearanceStyle =   0
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   630
         Index           =   1
         Left            =   15
         TabIndex        =   2
         Top             =   7230
         Width           =   12570
         _ExtentX        =   22172
         _ExtentY        =   1111
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtInput 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   0
            Left            =   1575
            TabIndex        =   5
            Top             =   105
            Width           =   1155
         End
         Begin VB.TextBox txtInput 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   1
            Left            =   4275
            TabIndex        =   4
            Top             =   105
            Width           =   1155
         End
         Begin VB.TextBox txtInput 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   2
            Left            =   6975
            TabIndex        =   3
            Top             =   105
            Width           =   1155
         End
         Begin Threed.SSPanel ssPanel3 
            Height          =   435
            Index           =   0
            Left            =   75
            TabIndex        =   6
            Top             =   105
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   767
            _Version        =   262144
            Windowless      =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "총 수 량"
            BevelWidth      =   3
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel ssPanel3 
            Height          =   435
            Index           =   1
            Left            =   2775
            TabIndex        =   7
            Top             =   105
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   767
            _Version        =   262144
            Windowless      =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "정  상"
            BevelWidth      =   3
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel ssPanel3 
            Height          =   435
            Index           =   2
            Left            =   5475
            TabIndex        =   8
            Top             =   105
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   767
            _Version        =   262144
            Windowless      =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "반  품"
            BevelWidth      =   3
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   615
         Index           =   0
         Left            =   15
         TabIndex        =   9
         Top             =   435
         Width           =   12570
         _ExtentX        =   22172
         _ExtentY        =   1085
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   495
            Left            =   645
            TabIndex        =   11
            Top             =   60
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   56950787
            UpDown          =   -1  'True
            CurrentDate     =   40279
         End
         Begin XtremeSuiteControls.PushButton SSCommand2 
            Height          =   540
            Left            =   2910
            TabIndex        =   13
            Top             =   45
            Width           =   1395
            _Version        =   851970
            _ExtentX        =   2461
            _ExtentY        =   952
            _StockProps     =   79
            Caption         =   " 조회"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Picture         =   "frm본사출고현황2.frx":081B
         End
         Begin XtremeSuiteControls.PushButton SSCommand1 
            Height          =   540
            Left            =   4335
            TabIndex        =   14
            Top             =   45
            Width           =   1395
            _Version        =   851970
            _ExtentX        =   2461
            _ExtentY        =   952
            _StockProps     =   79
            Caption         =   " 출력"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Picture         =   "frm본사출고현황2.frx":122D
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "기간"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   75
            TabIndex        =   12
            Top             =   120
            Width           =   480
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   405
         Left            =   15
         TabIndex        =   10
         Top             =   15
         Width           =   12570
         _ExtentX        =   22172
         _ExtentY        =   714
         _Version        =   262144
         ForeColor       =   16777215
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frm본사출고현황2.frx":1C3F
         PictureBackgroundStyle=   1
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image1 
            Height          =   240
            Left            =   60
            Picture         =   "frm본사출고현황2.frx":314D
            Top             =   60
            Width           =   240
         End
      End
   End
End
Attribute VB_Name = "frm본사출고현황2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim daoDB As Database        'Access DB

Private Sub Display_View()
    Dim FileName  As String
    Dim sData     As String
    Dim ii        As Integer
    Dim Query     As String
    Dim Query2    As String
    Dim strD      As String
    Dim iTotal(3) As Integer

    strD = Format(dtpDay.Value, "YYYY-MM-DD")
        
    
    i = 1
    ii = 1
        
    '------------------------------------------------------------
    '
    '------------------------------------------------------------
    Query = "SELECT * FROM TB_본사입고"
    Query = Query & " WHERE 본사출고일 = '" & strD & "' "
    Query = Query & " ORDER BY 택번호"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With fpSpread1
        .Col = 1
        .Row = 1
        .Col2 = .MaxCols
        .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClear
        .BlockMode = False
    
        Do Until ADORs.EOF
            If i = .MaxCols + 1 Then
                i = 1
                ii = ii + 1
            End If
            
            .Row = ii
            .Col = i: .Text = Format(ADORs!택번호 & "", "@-@@@")
            
            If ADORs!구분 & "" = "3" Then
                .Text = "반" & .Text
                iTotal(2) = iTotal(2) + 1
            ElseIf ADORs!구분 & "" = "3" = "E" Then
                .Text = "오류" & .Text
                iTotal(3) = iTotal(3) + 1
            Else
                iTotal(1) = iTotal(1) + 1
            End If
            
            iTotal(0) = iTotal(0) + 1
            
            i = i + 1
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
    End With
    
    txtInput(0).Text = iTotal(0)
    txtInput(1).Text = iTotal(1)
    txtInput(2).Text = iTotal(2)
End Sub

Private Sub Form_Activate()
    dtpDay.Value = Format(Date, "YYYY-MM-DD")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyChk (KeyCode)
'    If KeyCode = 13 Then
'       SendKeys "{Tab}"
'       KeyCode = 0
'    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    dtpDay.Value = Format(Date, "YYYY-MM-DD")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    daoDB.Close
End Sub

Private Sub MaskEdBox3_KeyPress(KeyAscii As Integer)
' Dim Query2 As String
' Dim strD As String
'
'   If KeyAscii = 13 Then
'      MsgBox "13"
'      strD = MaskEdBox1.ClipText & MaskEdBox2.ClipText & MaskEdBox3.ClipText
'      If Len(strD) < 6 Then
'         If Len(strD) = 0 Then
''            strD = Mid(Date, 1, 4) + Mid(Date, 6, 2) + Mid(Date, 9, 2)
'         Else
'            MsgBox " 연,월,일을 바르게 입력 하십시요 "
'            Exit Sub
'         End If
''      Else
'         strD = Mid(Date, 1, 2) + strD
'      End If
'      Query = "SELECT  P.접수일자, (P1.전화번호+'-'+ P1.전화2) AS 전화번호 , P1.성명, P.의류명, P.택번호, P.색상, P.내용, P.금액, P.결제여부, P.상표 "
'      Query = Query & " FROM TB_고객정보 AS P1, 입출고 AS P WHERE (P.접수일자 ='" & strD & "') AND P.지사출고상태 <>'출'"
''
'      Data1.RecordSource = Query
'      Data1.Refresh
'      fpSpread1.Refresh
 '  End If

End Sub

Private Sub SSCommand1_Click()
    Dim sData As String
    Dim ii  As Integer
    Dim iii As Integer
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
    sData = "출고일자 : " & Format(dtpDay.Value, "YYYY년MM월DD일") & Space(20)
    Printer.Print Space(10); sData

    iRow = iRow + 5
    Printer.CurrentY = iRow
    Printer.CurrentX = 0
    sData = ""
    sData = ""
    sData = Left(RTrim(sData) & Space(80), 80) & "출력일자 : " & Format(Now, "YYYY-MM-DD")
    Printer.Print Space(10); sData

    iRow = iRow + 5
    Printer.CurrentY = iRow
    Printer.CurrentX = 0
    sData = ""
''    sData = "품 목 명 : " & Trim(cboInput(0).Text) & " ~ " & Trim(cboInput(1).Text) & Space(20)
    sData = Left(RTrim(sData) & Space(80), 80) & "출력시간 : " & Format(Now, "hh:mm:ss")
    Printer.Print Space(10); sData
    
    iRow = iRow + 1
    Printer.Line (10, iRow + 3)-(260, iRow + 3)

''    iRow = iRow + 4
''    Printer.CurrentY = iRow
''    Printer.CurrentX = 0
''    Printer.Print "   가맹점                     출고수량"
    
''    iRow = iRow + 1
''    Printer.Line (0, iRow + 3)-(240, iRow + 3)
    
    sData = ""

    For i = memRow To fpSpread1.MaxRows - 1
        fpSpread1.Row = i

        fpSpread1.Col = 1
        If fpSpread1.Text = "" Then
            Exit For
        End If
        
        sData = Space(1) & Right(Space(7) & fpSpread1.Text, 7) & " "

        fpSpread1.Col = 2: sData = sData & Space(1) & Right(Space(7) & Trim(fpSpread1.Text), 7) & " "
        fpSpread1.Col = 3: sData = sData & Space(1) & Right(Space(7) & Trim(fpSpread1.Text), 7) & " "
        fpSpread1.Col = 4: sData = sData & Space(1) & Right(Space(7) & Trim(fpSpread1.Text), 7) & " "
        fpSpread1.Col = 5: sData = sData & Space(1) & Right(Space(7) & Trim(fpSpread1.Text), 7) & " "
        fpSpread1.Col = 6: sData = sData & Space(1) & Right(Space(7) & Trim(fpSpread1.Text), 7) & " "
        fpSpread1.Col = 7: sData = sData & Space(1) & Right(Space(7) & Trim(fpSpread1.Text), 7) & " "
        fpSpread1.Col = 8: sData = sData & Space(1) & Right(Space(7) & Trim(fpSpread1.Text), 7) & " "
        fpSpread1.Col = 9: sData = sData & Space(1) & Right(Space(7) & Trim(fpSpread1.Text), 7) & " "
        fpSpread1.Col = 10: sData = sData & Space(1) & Right(Space(7) & Trim(fpSpread1.Text), 7) & " "
        
        iRow = iRow + 4
        Printer.CurrentY = iRow
        Printer.CurrentX = 0
        Printer.Print Space(10); sData

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
    sData = "출고수량 : " '& txtInput(0).Text & Space(20)
    sData = Left(RTrim(sData) & Space(40), 40) & "정    상 : " & txtInput(1).Text
    sData = Left(RTrim(sData) & Space(80), 80) & "반    품 : " & txtInput(2).Text
    Printer.Print Space(10); sData
    
    iRow = iRow + 1
    Printer.Line (10, iRow + 3)-(260, iRow + 3)

    Printer.EndDoc
End Sub

Private Sub SSCommand2_Click()
    Call Display_View
End Sub
