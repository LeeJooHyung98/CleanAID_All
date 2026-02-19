VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm일일매출집계 
   Caption         =   "일일매출 집계"
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
      PaneTree        =   "frm일일매출집계.frx":0000
      Begin Threed.SSPanel SSPanel1 
         Height          =   750
         Index           =   0
         Left            =   15
         TabIndex        =   1
         Top             =   450
         Width           =   16380
         _ExtentX        =   28893
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox cboGubun 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   915
            Style           =   2  '드롭다운 목록
            TabIndex        =   10
            Top             =   405
            Width           =   1680
         End
         Begin VB.TextBox txtFind 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2625
            TabIndex        =   2
            Top             =   405
            Width           =   1965
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Left            =   915
            TabIndex        =   3
            Top             =   45
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM"
            Format          =   56426499
            UpDown          =   -1  'True
            CurrentDate     =   40279
         End
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   630
            Left            =   8340
            TabIndex        =   5
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 조회(&F)"
            Appearance      =   6
            Picture         =   "frm일일매출집계.frx":0072
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   3
            Left            =   9885
            TabIndex        =   6
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 엑셀(&E)"
            Appearance      =   6
            Picture         =   "frm일일매출집계.frx":076C
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   13170
            TabIndex        =   7
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frm일일매출집계.frx":0EE6
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   4
            Left            =   11430
            TabIndex        =   8
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 출력(&P)"
            Appearance      =   6
            Picture         =   "frm일일매출집계.frx":1F78
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "검색조건:"
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
            Index           =   3
            Left            =   45
            TabIndex        =   9
            Top             =   465
            Width           =   840
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "조회년월:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   45
            TabIndex        =   4
            Top             =   105
            Width           =   840
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   11
         Top             =   15
         Width           =   16380
         _ExtentX        =   28893
         _ExtentY        =   741
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "      일일매출 집계"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm일일매출집계.frx":2672
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm일일매출집계.frx":2898
            Top             =   -15
            Width           =   765
         End
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   10740
         Left            =   15
         TabIndex        =   12
         Top             =   1215
         Width           =   16380
         _Version        =   524288
         _ExtentX        =   28893
         _ExtentY        =   18944
         _StockProps     =   64
         AutoCalc        =   0   'False
         BackColorStyle  =   1
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
         DInformActiveRowChange=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
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
         MaxRows         =   18
         MoveActiveOnFocus=   0   'False
         OperationMode   =   1
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   0
         SpreadDesigner  =   "frm일일매출집계.frx":3462
         UserResize      =   1
         VisibleCols     =   7
         VisibleRows     =   21
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "frm일일매출집계"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim First_Day As String  '첫 일자
Dim Last_Day  As String  '마지막 일자

Dim iMonth    As Integer '월
Dim iWeek     As Integer '첫날 요일
Dim iLast     As Integer '마지막날짜

Dim tmpDay    As String

Dim strStart As String
Dim strEnd   As String

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 3: Call Export_Excel(frmMain.cdgExcel, sprGrid)
        
        Case 4:
            Rtn = MsgBox("출력 미리보기를 하시겠습니까?", vbQuestion + vbYesNo, "출력")
            
            If Rtn = vbYes Then
                Call Data_Print(True)
            Else
                Call Data_Print(False)
            End If
            
        Case 5:
            Unload Me
    End Select
End Sub

Private Sub Data_Print(Print_PreView As Boolean)
    Dim iRow As String
    
    On Error GoTo ErrRtn
    
    If sprGrid.MaxRows = 0 Then Exit Sub

    If Dir(AppPath & "XML", vbDirectory) = "" Then
        MkDir AppPath & "XML"
    End If

    Open AppPath & "XML\일일매출집계.XML" For Output As #1
    
    Print #1, "<?xml version=""1.0"" encoding=""EUC-KR""?>"
    Print #1, "<root>"
        
          XML = "    <조건>"
    XML = XML & "        <매출월>" & Format(dtpDay.Value, "YYYY년 MM월") & " 일일매출집계</매출월>"
    XML = XML & "        <가맹점>크린에이드 - " & Func_Replace(가맹점정보.가맹점명) & "</가맹점>"
    XML = XML & "   </조건>"
    Print #1, XML
    
    With sprGrid
        For i = 1 To .MaxRows
            .Row = i
            
            If i = 1 Then
                iRow = 1
                
                Print #1, "    <Data>"
            ElseIf i Mod 3 = 1 Then
                Print #1, "    </Data>"
                
                iRow = 1
                
                Print #1, "    <Data>"
            Else
                iRow = iRow + 1
            End If
            
            iRow = Format(iRow, "00")
                    
                             
            .Col = 1:        XML = "        <Cell" & iRow & "01>" & Func_Replace(.Text) & "</Cell" & iRow & "01>"
            .Col = 2:  XML = XML & "        <Cell" & iRow & "02>" & Func_Replace(.Text) & "</Cell" & iRow & "02>"
            .Col = 3:  XML = XML & "        <Cell" & iRow & "03>" & Func_Replace(.Text) & "</Cell" & iRow & "03>"
            .Col = 4:  XML = XML & "        <Cell" & iRow & "04>" & Func_Replace(.Text) & "</Cell" & iRow & "04>"
            .Col = 5:  XML = XML & "        <Cell" & iRow & "05>" & Func_Replace(.Text) & "</Cell" & iRow & "05>"
            .Col = 6:  XML = XML & "        <Cell" & iRow & "06>" & Func_Replace(.Text) & "</Cell" & iRow & "06>"
            .Col = 7:  XML = XML & "        <Cell" & iRow & "07>" & Func_Replace(.Text) & "</Cell" & iRow & "07>"
            .Col = 8:  XML = XML & "        <Cell" & iRow & "08>" & Func_Replace(.Text) & "</Cell" & iRow & "08>"
            .Col = 9:  XML = XML & "        <Cell" & iRow & "09>" & Func_Replace(.Text) & "</Cell" & iRow & "09>"
            .Col = 10: XML = XML & "        <Cell" & iRow & "10>" & Func_Replace(.Text) & "</Cell" & iRow & "10>"
            .Col = 11: XML = XML & "        <Cell" & iRow & "11>" & Func_Replace(.Text) & "</Cell" & iRow & "11>"
            .Col = 12: XML = XML & "        <Cell" & iRow & "12>" & Func_Replace(.Text) & "</Cell" & iRow & "12>"
            .Col = 13: XML = XML & "        <Cell" & iRow & "13>" & Func_Replace(.Text) & "</Cell" & iRow & "13>"
            .Col = 14: XML = XML & "        <Cell" & iRow & "14>" & Func_Replace(.Text) & "</Cell" & iRow & "14>"
                       Print #1, XML
        Next i
        
        Print #1, "    </Data>"
        Print #1, "</root>"
        Close #1
    End With
    
    If Print_PreView = True Then
        With rpt일일매출집계
            .dc.FileURL = AppPath & "XML\일일매출집계.XML"
            .Show 1
        End With
    Else
        With rpt일일매출집계
            .dc.FileURL = AppPath & "XML\일일매출집계.XML"
            .PrintReport False
        End With
    
        Unload rpt일일매출집계
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub cmdList_Click()
    On Error GoTo ErrRtn
    
    Call Calendar_Display
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub dtpDay_Change()
    dtpDay.Enabled = False
    DoEvents
    
    Call Calendar_Display
    
    dtpDay.Enabled = True
End Sub

Private Sub Calendar_Display()
    Dim iRow    As Integer
    Dim iCol    As Integer
    Dim tmpDate As String
    
    Dim j     As Integer
    
    On Error GoTo ErrRtn
    
    iMonth = Right(Format(dtpDay.Value, "YYYY-MM"), 2)        '월
    First_Day = Format(dtpDay.Value, "YYYY-MM") & "-01"       '첫 일자
    
    tmpDay = Format(DateAdd("m", 1, First_Day), "YYYY-MM-DD")
    Last_Day = Format(DateAdd("d", -1, tmpDay), "YYYY-MM-DD") '마지막 일자
    
    '0-일요일 ... 7-토요일
    iWeek = Weekday(First_Day)
    iLast = Right(Last_Day, 2)
    
    With sprGrid
        For i = 1 To .MaxRows
            .Row = i
            
            For j = 1 To .MaxCols
                .Col = j
                .Text = ""
                .CellTag = ""
            Next j
        Next i
        
        i = 1
        iRow = 1
        
        Do Until i > iLast
            .Row = iRow
            
            Select Case iWeek
                Case 1: iCol = 2
                Case 2: iCol = 4
                Case 3: iCol = 6
                Case 4: iCol = 8
                Case 5: iCol = 10
                Case 6: iCol = 12
                Case 7: iCol = 14
            End Select
            
            tmpDate = Format(dtpDay.Value, "YYYY-MM") & "-" & Format(i, "00")
            
            .Col = iCol: .Text = i & "": .CellTag = tmpDate
            
            .FontBold = True
            .RowHeight(iRow) = 20
            
            Select Case iWeek
                Case 1: .ForeColor = vbRed
                Case 7: .ForeColor = vbBlue
                Case Else: .ForeColor = vbBlack
            End Select
                
            '--------------------------------------------------------------------------
            '
            '--------------------------------------------------------------------------
            'Query = "SELECT    COUNT(*)  AS 접수량"
            'Query = Query & ", SUM(금액) AS 접수금액"
            'Query = Query & " FROM TB_입출고"
            'Query = Query & " WHERE 접수일자 = '" & tmpDate & "'"
            'Query = Query & "   AND ((판매취소 <> 'Y')"
            ''Query = Query & "   AND ((판매취소 = '') AND (판매취소일자 IS NULL OR 판매취소일자 = '')"
            'Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
            'Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
            
            Query = "SELECT    ISNULL(COUNT(*),0)  AS 접수량"
            Query = Query & ", ISNULL(SUM(금액),0) AS 접수금액"
            Query = Query & " FROM TB_입출고"
            Query = Query & " WHERE (접수일자 = '" & tmpDate & "')"
            Query = Query & "   AND (판매취소 <> 'Y')"
            Set ADORs = New ADODB.RecordSet
            ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            If ADORs!접수량 = 0 Then
            
            Else
                .Row = iRow + 1
                .Col = iCol - 1: .Text = Format(ADORs!접수량, "#,##0") & "점": .ForeColor = vbBlue
                .Col = iCol:     .Text = Format(ADORs!접수금액, "#,##0") & "원": .ForeColor = vbBlue
            End If
            ADORs.Close
            Set ADORs = Nothing
            
            '
            Query = "SELECT    COUNT(*)  AS 출고수량"
            Query = Query & ", SUM(금액) AS 출고금액"
            Query = Query & " FROM TB_입출고"
            Query = Query & " WHERE 출고일자 = '" & tmpDate & "'"
            'Query = Query & "   AND ((판매취소 <> 'Y')"
            ''Query = Query & "   AND ((판매취소 = '') AND (판매취소일자 IS NULL OR 판매취소일자 = '')"
            'Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
            'Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
            Set ADORs = New ADODB.RecordSet
            ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            If ADORs!출고수량 = 0 Then
            
            Else
                .Row = iRow + 2
                .Col = iCol - 1: .Text = Format(ADORs!출고수량, "#,##0") & "점": .ForeColor = vbRed
                .Col = iCol:     .Text = Format(ADORs!출고금액, "#,##0") & "원": .ForeColor = vbRed
            End If
            ADORs.Close
            Set ADORs = Nothing
            
            '--------------------------------------------------------------------------
            
            i = i + 1
            
            If iWeek = 7 Then
                iRow = iRow + 3
                
                iWeek = 1
            Else
                iWeek = iWeek + 1
            End If
        Loop
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    With sprGrid
        .RowHeight(1) = 20
        .RowHeight(4) = 20
        .RowHeight(7) = 20
        .RowHeight(10) = 20
        .RowHeight(13) = 20
        .RowHeight(16) = 20
        .RowHeight(19) = 20
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeNormal
    End With
    
    dtpDay.Value = Format(DateAdd("m", -1, Date), "YYYY-MM-DD")
    
    With cboGubun
        .Clear
        .AddItem "성명"
        .AddItem "전화번호"
        .AddItem "고객코드"
        
        .ListIndex = 0
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    cmdBtn(5).Left = Me.Width - cmdBtn(5).Width - 200
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
'                Query = Query & " WHERE 전화번호 = '" & txtTel(0).Text & "' "
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
'                    frm고객검색.DataDisplay Query
'                    frm고객검색.Show 1
'
'                    If frm고객검색.SELECTCODE = "CANCEL" Then
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
'        Query = "UPDATE TB_입출고 SET 지사출고상태         = '" & strData(0) & "'"
'        Query = Query & "           , 지사출고일자     = '" & IIf(Trim(strData(0)) = "", "", Format(Date, "YYYY-MM-DD")) & "'"
'        Query = Query & "           , 가맹점입고구분 = '수동'"
'        Query = Query & " WHERE 접수일자 = '" & strData(1) & "'"
'        Query = Query & "  AND 택번호 = '" & strData(2) & "'"
'        ADOCon.Execute Query
'    End If
'End Sub
