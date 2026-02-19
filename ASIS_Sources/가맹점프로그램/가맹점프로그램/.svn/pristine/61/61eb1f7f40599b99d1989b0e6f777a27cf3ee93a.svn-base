VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form frm반품 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   6630
   ClientLeft      =   1140
   ClientTop       =   1425
   ClientWidth     =   9750
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   15.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form27"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   9735
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   4695
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   9510
         _Version        =   524288
         _ExtentX        =   16775
         _ExtentY        =   8281
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         MaxCols         =   7
         MoveActiveOnFocus=   0   'False
         OperationMode   =   1
         Protect         =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frm반품.frx":0000
         VisibleCols     =   500
         VisibleRows     =   500
         HighlightHeaders=   1
         HighlightStyle  =   1
         ScrollBarStyle  =   2
      End
      Begin VB.PictureBox CR1 
         Height          =   480
         Left            =   240
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   5
         Top             =   6120
         Width           =   1200
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   615
         Left            =   2760
         TabIndex        =   4
         Top             =   360
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1085
         _Version        =   262144
         ForeColor       =   16711680
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "환불요청서"
         BevelWidth      =   2
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   615
         Left            =   4200
         TabIndex        =   2
         Top             =   5880
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1085
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "인   쇄"
         ButtonStyle     =   2
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   615
         Left            =   6960
         TabIndex        =   3
         Top             =   5880
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1085
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "종   료"
         ButtonStyle     =   2
      End
   End
End
Attribute VB_Name = "frm반품"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Option Explicit
'
'Dim Myrec As Recordset
'Dim PrintChk As Boolean
'Dim UpdateChk As Boolean
'
'
'Private Sub billPrint_1()
'    Dim chkbill_1 As chkbill
'    Dim lngSumMoney As Long
'    Dim RowCount As Integer
'    Dim strAgentName As String
'    Dim intMaxCnt As Integer
'    Dim lngRatio As Single
'
'
'    intMaxCnt = 0
'    lngSumMoney = 0
'
'    '---------------------------------------------------------
'    '
'    '---------------------------------------------------------
'    Query = "SELECT   가맹점명,"
'    Query = Query & " 비율"
'    Query = Query & " FROM TB_기본정보 "
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
'
'    If ADORs.RecordCount > 1 Then
'        strAgentName = "세탁"
'
'        lngRatio = 60
'    Else
'        strAgentName = ADORs!가맹점명 & ""
'
'        lngRatio = 100 - CLng(ADORs!비율)
'    End If
'    ADORs.Close
'    Set ADORs = Nothing
'
'    lngRatio = lngRatio / 100
'
'    For i = 1 To fpSpread1.MaxRows
'        fpSpread1.Row = i
'        fpSpread1.Col = 7
'        If fpSpread1.Value = -1 Then
'            With fpSpread1
'                .Col = 1: chkbill_1.strchkdate(i) = Trim(fpSpread1.Text)
'                .Col = 2: chkbill_1.strchkTno(i) = Trim(fpSpread1.Text)
'                .Col = 3: chkbill_1.strchkItem(i) = Trim(fpSpread1.Text)
'                .Col = 4: chkbill_1.lngMoney(i) = CLng(Val(fpSpread1.Value))
'
'                '환불금액
'                chkbill_1.lngchkRejectmoney(i) = CLng(Val(fpSpread1.Value)) * lngRatio
'                lngSumMoney = lngSumMoney + (CLng(Val(fpSpread1.Value)) * lngRatio)
'            End With
'
'            intMaxCnt = intMaxCnt + 1
'        Else
'            'Exit For
'        End If
'    Next i
'
'    Printer.Font.Name = "굴림체"
'    Printer.Font.Size = 18
'    Printer.Font.Bold = True
'    Printer.Print
'    Printer.Print
'    Printer.Print
'    Printer.Print Tab(18); "   반 품  환 불   청 구 서 "
'    Printer.Print Tab(18); "  ━━━━━━━━━━━━━"
'    Printer.Print
'    Printer.Font.Bold = False
'    Printer.Font.Size = 12
'    Printer.Print ; Tab(10); "매장명 :"; Spc(1); strAgentName; Tab(66); Date
'    Printer.Print Tab(10); "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
'    Printer.Print Tab(10); " 입고일 "; Tab(24); "TAG-NO"; Tab(38); "품  명"; Tab(52); " 금  액"; Tab(66); "환불금액"
'    Printer.Print Tab(10); "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
'
'    For i = 1 To intMaxCnt
'        With chkbill_1
'           Printer.Print ; Tab(10); .strchkdate(i); Tab(24); .strchkTno(i); Tab(38); .strchkItem(i); Tab(52); .lngMoney(i); Tab(66); .lngchkRejectmoney(i)
'        End With
'    Next i
'
'    Printer.Print Tab(10); "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
'    Printer.Print Tab(10); " 합계금액 "; Tab(66); lngSumMoney
'    Printer.Print Tab(10); "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
'    Printer.Print Tab(10); ""
'    Printer.Print Tab(10); ""
'    Printer.Print Tab(17); "상기 금액을 고객에게 환불하였기에 본사에 청구하나이다 "
'    Printer.Print Tab(10); ""
'    Printer.Print Tab(10); ""
'    Printer.Print Tab(10); "구 TAG"
'    Printer.EndDoc
'
'    DoEvents
'
'    Call UpDateDb
'End Sub
'
''프린터여부확인
'Function PrintCheck() As Boolean
'    If UpdateChk And Not PrintChk Then
'        PrintCheck = False
'    Else
'        PrintCheck = True
'    End If
'End Function
'
'Sub BillPrint()
'    Dim dblRatio As Double
'
'    Query = "SELECT 비율 FROM TB_기본정보 "
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
'
'    If IsNull(ADORs!비율) = True Or Len(ADORs!비율) < 1 Then
'       dblRatio = 60
'    Else
'       dblRatio = 100 - CDbl(ADORs!비율) '본사마진계산 본사에 청구하기위해
'    End If
'    ADORs.Close
'
'    With CR1
'       '.Destination = crptToWindow
'        .DataFiles(0) = App.Path & "\DB\Laundry.mdb"
'        .ReportFileName = App.Path & "\Report\환불요청서.rpt"
'        .Formulas(0) = "ratio=" & (dblRatio) / 100
'       '.PrintDay = Date
'        .Action = 1
'    End With
'
'    DoEvents
'    Call UpDateDb
'
'    PrintChk = True
'End Sub
'
'Sub Spreadfill()
'    Dim iCnt As Integer
'
'    Query = "Select P.입고일, P.택번호,P.의류명,P.금액,P.전화,P1.성명 "
'    Query = Query & " From 환불요청서 AS P LEFT OUTER JOIN 고객정보 AS P1 AND P.고객코드 = P1.고객코드"
'    Query = Query & " WHERE P.구분    = '1'"
'    Set SUBRs = New ADODB.Recordset
'    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
'
'
'    iCnt = 1
'
'    While Not SUBRs.EOF And Not SUBRs.BOF
'        With fpSpread1
'            .Row = iCnt
'
'            .Col = 1: .Value = Mid(SUBRs!입고일, 3, 2) & "/" & Mid(SUBRs!입고일, 5, 2) & "/" & Mid(SUBRs!입고일, 7, 2)
'            .Col = 2: .Value = SUBRs!택번호
'            .Col = 3: .Value = SUBRs!의류명
'            .Col = 4: .Value = SUBRs!금액
'            .Col = 5: .Value = Mid(SUBRs!전화, 1, 4) & "-" & Mid(SUBRs!전화, 5, 4)
'            .Col = 6: .Value = SUBRs!성명
'        End With
'
'        iCnt = iCnt + 1
'
'        SUBRs.MoveNext
'    Wend
'
'    If iCnt > 11 Then
'        fpSpread1.MaxRows = iCnt
'    Else
'        fpSpread1.MaxRows = 10
'    End If
'
'    SUBRs.Close
'End Sub
'
'Sub UpDateDb()
'    Dim iCnt As Integer
'
'    With fpSpread1
'        For iCnt = 1 To .MaxRows
'            .Row = iCnt
'            .Col = 7
'            If Not Val(.Value) = 0 Then
'                .Col = 2
'
'                Query = "UPDATE TB_반품환불 SET 요청일 ='" & Mid(Date, 1, 4) & Mid(Date, 6, 2) & Mid(Date, 9, 2) & "'"
'                Query = Query & " ,구분='3'"
'                Query = Query & " WHERE 택번호 = '" & Trim(.Value) & "'"
'                ADOCon.Execute Query
'
'                UpdateChk = True
'            End If
'        Next iCnt
'    End With
'End Sub
'
'Private Sub cmdExit_Click()
'    PrintChk = True
'    'BillPrint
'    billPrint_1
'    Unload Me
'End Sub
'
'Private Sub cmdPrint_Click()
'    If PrintCheck Then
'        frm반품.Hide
'        Unload frm반품
'
'    ElseIf MsgBox("환불요청서를 인쇄하지 않았습니다." & vbCr & vbLf & "환불요청서를 인쇄하시겠습니까?", vbYesNo) = vbYes Then
'        Call billPrint_1
'        Call UpDateDb
'
'    Else
'        frm반품.Hide
'        Unload frm반품
'    End If
'   ' Query = "UPDATE TB_반품환불 set 구분 = '3' Where 구분 = '1'"
'   ' ADOCon.Execute Query
'End Sub
'
'Private Sub Form_Load()
'    DB_Connect
'
'    UpdateChk = False
'    PrintChk = False
'    Spreadfill
'End Sub
''Private Sub SSCommand1_Click()
''    Dim iCnt As Integer
''
''
''    Query = "Select * FROM TB_입출고"
''    Set Myrec = MyDB.OpenRecordset(Query)
''
''
''    Query = "Delete FROM TB_반품환불"
''    ADOCon.Execute Query
''
''    For iCnt = 1 To 9
''        Query = "INSERT INTO TB_반품환불(입고일,구분,번호) values('" & Myrec!입고일 & "','1','" & Myrec!택번호 & "')"
''        ADOCon.Execute Query
''        Myrec.MoveNext
''    Next iCnt
''    DoEvents
''    Spreadfill
''    Myrec.Close
''End Sub
'
'Private Sub fpSpread1_Click(ByVal Col As Long, ByVal Row As Long)
'    fpSpread1.Row = Row
'    '빈칸인지 확인
'    fpSpread1.Col = 2
'    If Len(Trim(fpSpread1.Value)) = 0 Then
'        Exit Sub
'    End If
'     fpSpread1.Col = 7
'    If fpSpread1.Value Then
'       fpSpread1.Value = False
'    Else
'       fpSpread1.Value = True
'    End If
'
'End Sub
