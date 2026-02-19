VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm일일매출집계 
   Caption         =   "일일 매출집계"
   ClientHeight    =   11970
   ClientLeft      =   1275
   ClientTop       =   3930
   ClientWidth     =   16410
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form20"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11970
   ScaleWidth      =   16410
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   11970
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16410
      _ExtentX        =   28945
      _ExtentY        =   21114
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm달력.frx":0000
      Begin Threed.SSPanel SSPanel1 
         Height          =   870
         Index           =   0
         Left            =   15
         TabIndex        =   1
         Top             =   435
         Width           =   16380
         _ExtentX        =   28893
         _ExtentY        =   1535
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox cboGubun 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   945
            Style           =   2  '드롭다운 목록
            TabIndex        =   10
            Top             =   450
            Width           =   1680
         End
         Begin VB.TextBox txtFind 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2655
            TabIndex        =   2
            Top             =   450
            Width           =   1965
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   360
            Index           =   0
            Left            =   945
            TabIndex        =   3
            Top             =   60
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   635
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM"
            Format          =   54853635
            UpDown          =   -1  'True
            CurrentDate     =   40279
         End
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   750
            Left            =   8340
            TabIndex        =   5
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   79
            Caption         =   " 조회(&F)"
            Appearance      =   6
            Picture         =   "frm달력.frx":0072
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   750
            Index           =   3
            Left            =   9885
            TabIndex        =   6
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   79
            Caption         =   " 엑셀(&E)"
            Appearance      =   6
            Picture         =   "frm달력.frx":076C
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   750
            Index           =   5
            Left            =   13170
            TabIndex        =   7
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frm달력.frx":0EE6
         End
         Begin XtremeSuiteControls.PushButton cmdPrint 
            Height          =   750
            Left            =   11430
            TabIndex        =   8
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   79
            Caption         =   " 출력(&P)"
            Appearance      =   6
            Picture         =   "frm달력.frx":1F78
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "검색조건 :"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   60
            TabIndex        =   9
            Top             =   480
            Width           =   840
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "접수일자 :"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   60
            TabIndex        =   4
            Top             =   90
            Width           =   840
         End
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   10635
         Left            =   15
         TabIndex        =   11
         Top             =   1320
         Width           =   16380
         _Version        =   524288
         _ExtentX        =   28892
         _ExtentY        =   18759
         _StockProps     =   64
         AutoCalc        =   0   'False
         BackColorStyle  =   1
         ColsFrozen      =   4
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
         DInformActiveRowChange=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GrayAreaBackColor=   16777215
         MaxCols         =   14
         MaxRows         =   21
         MoveActiveOnFocus=   0   'False
         OperationMode   =   1
         Protect         =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         ScrollBarMaxAlign=   0   'False
         SpreadDesigner  =   "frm달력.frx":2672
         UserResize      =   1
         VisibleCols     =   7
         VisibleRows     =   21
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   405
         Left            =   15
         TabIndex        =   12
         Top             =   15
         Width           =   16380
         _ExtentX        =   28893
         _ExtentY        =   714
         _Version        =   262144
         Font3D          =   3
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "    일일 매출집계"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm달력.frx":5299
         BorderWidth     =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image1 
            Height          =   240
            Left            =   60
            Picture         =   "frm달력.frx":56FB
            Top             =   60
            Width           =   240
         End
      End
   End
End
Attribute VB_Name = "frm일일매출집계"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strStart As String
Dim strEnd   As String

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        
        Case 3:
        
        Case 5:
            Unload Me
    End Select
End Sub

Private Sub cmdList_Click()
    On Error GoTo ErrRtn
    
    Query = "SELECT    A.성명"
    Query = Query & ", A.휴대폰"
    Query = Query & ", A.전화번호"
    Query = Query & ", B.접수일자"
    Query = Query & ", B.출고일자"
    Query = Query & ", B.본출"
    Query = Query & ", B.품명"
    Query = Query & ", B.택번호"
    Query = Query & ", B.색상"
    Query = Query & ", B.무늬"
    Query = Query & ", B.내용"
    Query = Query & ", B.금액"
    Query = Query & ", B.결제여부"
    Query = Query & ", B.상표"
    Query = Query & ", B.확인"
    Query = Query & " FROM TB_고객정보 AS A LEFT OUTER JOIN TB_입출고 AS B ON (A.고객코드 = B.고객코드) "

    Select Case cboGubun.Text
        Case "성명":     Query = Query & " WHERE (A.성명 LIKE '%" & Trim(txtFind.Text) & "%') "
        Case "전화번호": Query = Query & " WHERE (A.전화번호 LIKE '%" & Trim(txtFind.Text) & "%') "
        Case "고객코드": Query = Query & " WHERE (A.고객코드 LIKE '%" & Trim(txtFind.Text) & "%') "
    End Select
    
    Query = Query & "   AND (B.접수일자 >= '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "'"
    Query = Query & "   AND  B.접수일자 <= '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "')"
    Query = Query & "   AND (B.판매취소 IS NULL OR B.판매취소 <> 'Y')"
    Query = Query & " ORDER BY B.접수일자, B.택번호"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    If ADORs.EOF = True Then
        ADORs.Close
        Set ADORs = Nothing
        
        MsgBox "[" & txtFind.Text & "] 에 해당되는 자료가 없읍니다 !", vbInformation, "접수현황"
        Exit Sub
    End If
    
    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = ADORs!접수일자 & ""
            .Col = 2:  .Text = ADORs!성명 & ""
            .Col = 3:  .Text = ADORs!휴대폰 & ""
            .Col = 4:  .Text = ADORs!전화번호 & ""
            .Col = 5:  .Text = ADORs!출고일자 & ""
            .Col = 6:  .Text = ADORs!본출 & ""
            .Col = 7:  .Text = ADORs!품명 & ""
            .Col = 8:  .Text = ADORs!택번호 & ""
            .Col = 9:  .Text = ADORs!색상 & ""
            .Col = 10: .Text = ADORs!무늬 & ""
            .Col = 11: .Text = ADORs!내용 & ""
            .Col = 12: .Text = ADORs!금액 & ""
            .Col = 13: .Text = ADORs!결제여부 & ""
            .Col = 14: .Text = ADORs!상표 & ""
            .Col = 15: .Text = ADORs!확인 & ""
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Activate()
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    With sprGrid
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeExtended
    End With
    
    dtpDay(0).Value = Format(DateAdd("m", -1, Date), "YYYY-MM-DD")
    
    With cboGubun
        .Clear
        .AddItem "성명"
        .AddItem "전화번호"
        .AddItem "고객코드"
        
        .ListIndex = 0
    End With
End Sub

Private Sub Form_Resize()
    On Error GoTo ErrRtn
    
    cmdBtn(5).Left = Me.Width - cmdBtn(5).Width - 200
    
    Exit Sub
    
ErrRtn:

End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdList_Click
    End If
End Sub

'Private Sub mskTag_Change()
'    mskTag.SelStart = 0
'    mskTag.SelLength = 8
'End Sub
'
'Private Sub mskTag_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        Call Search_All
'
'        'Search_TagNo
'    End If
'End Sub

'Private Sub txtTel_GotFocus(Index As Integer)
'    txtTel(Index).SelStart = 0
'    txtTel(Index).SelLength = 4
'End Sub
'
'Private Sub txtTEL_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        Select Case Index
'            Case 0
'
'            Case 1
'                Query = " SELECT * FROM TB_고객정보 "
'                Query = Query & "WHERE 전화번호 = '" & txtTel(0).Text & "' "
'                Set Rs = New ADODB.Recordset
'                Rs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
'
'                If Not Rs.EOF Or Not Rs.BOF Then
'                    Rs.MoveLast
'                End If
'
'                If Rs.RecordCount = 1 Then
'                    ' 찿기
'                    Search_All
'                    'Search_Tel
'                    Exit Sub
'
'                ElseIf Rs.RecordCount < 1 Then
'                    MsgBox " 등록된 회원이 없습니다.", vbInformation, "확인"
'                    Exit Sub
'
'                ElseIf Rs.RecordCount >= 2 Then
'                    '뿌리고 입력대기상태
'                    frm동명이인.DataDisplay Query
'                    frm동명이인.Show 1
'
'                    If frm동명이인.SELECTCODE = "CANCEL" Then
'                        txtTel(1).SetFocus
'                        Exit Sub
'                    End If
'
'                    If 고객정보.고객코드 <> "Error" Then txtCode = 고객정보.고객코드
'
'                    txtFind.Text = 고객정보.성명
'                    Search_All
'
'                    Exit Sub
'                End If
'        End Select
'    End If
'End Sub

'Private Sub sprGrid_DblClick(ByVal Col As Long, ByVal Row As Long)
'    Dim intMaxRow  As Integer
'    Dim intActrow  As Integer
'    Dim strData(3) As String
'
'    If Col = 6 Then
'        With sprGrid
'            .Row = Row
'            .Col = 6: strData(0) = IIf(.Text = "" Or .Text = " ", "出", "")
'
'            If strData(0) = "" Then strData(0) = " "
'
'            .Col = 6: .Text = strData(0)
'
'            .Col = 1:  strData(1) = .Text
'            .Col = 8:  strData(2) = .Text
'        End With
'
'        Query = "UPDATE TB_입출고 SET 본출         = '" & strData(0) & "'"
'        Query = Query & "           , 본출일자     = '" & IIf(Trim(strData(0)) = "", "", Format(Date, "YYYY-MM-DD")) & "'"
'        Query = Query & "           , 본출입고구분 = '수동'"
'        Query = Query & " WHERE 접수일자 = '" & strData(1) & "'"
'        Query = Query & "  AND 택번호 = '" & strData(2) & "'"
'        ADOCon.Execute Query
'    End If
'End Sub
