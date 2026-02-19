VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm고객미출고현황 
   AutoRedraw      =   -1  'True
   Caption         =   "고객 미출고 현황"
   ClientHeight    =   10110
   ClientLeft      =   4950
   ClientTop       =   3855
   ClientWidth     =   14400
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form16"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10110
   ScaleWidth      =   14400
   ShowInTaskbar   =   0   'False
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   10110
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14400
      _ExtentX        =   25400
      _ExtentY        =   17833
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm고객미출고현황.frx":0000
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   8880
         Left            =   15
         TabIndex        =   1
         Top             =   1215
         Width           =   14370
         _Version        =   524288
         _ExtentX        =   25347
         _ExtentY        =   15663
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BackColorStyle  =   1
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
         DInformActiveRowChange=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         MaxCols         =   14
         MaxRows         =   1000000
         OperationMode   =   1
         Protect         =   0   'False
         SpreadDesigner  =   "frm고객미출고현황.frx":0072
         VisibleCols     =   9
         VisibleRows     =   200
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   750
         Left            =   15
         TabIndex        =   2
         Top             =   450
         Width           =   14370
         _ExtentX        =   25347
         _ExtentY        =   1323
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtData 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   10  '한글 
            Index           =   0
            Left            =   2370
            TabIndex        =   13
            Top             =   405
            Width           =   1875
         End
         Begin VB.ComboBox cboGubun 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   915
            Style           =   2  '드롭다운 목록
            TabIndex        =   12
            Top             =   405
            Width           =   1425
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   0
            Left            =   915
            TabIndex        =   3
            Top             =   45
            Width           =   1425
            _ExtentX        =   2514
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
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   58720259
            CurrentDate     =   40279
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   1
            Left            =   2610
            TabIndex        =   4
            Top             =   45
            Width           =   1425
            _ExtentX        =   2514
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
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   58720259
            CurrentDate     =   40279
         End
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   630
            Left            =   5760
            TabIndex        =   8
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 조회(&F)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm고객미출고현황.frx":0997
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   3
            Left            =   7290
            TabIndex        =   9
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 엑셀(&E)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm고객미출고현황.frx":1091
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   12825
            TabIndex        =   10
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm고객미출고현황.frx":180B
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   4
            Left            =   11085
            TabIndex        =   11
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 출력(&P)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm고객미출고현황.frx":289D
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   0
            Left            =   8820
            TabIndex        =   15
            Top             =   60
            Width           =   2235
            _Version        =   851970
            _ExtentX        =   3942
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 영수증 프린터 출력"
            Appearance      =   6
            Picture         =   "frm고객미출고현황.frx":2F97
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "검색조건:"
            Height          =   210
            Index           =   1
            Left            =   45
            TabIndex        =   14
            Top             =   480
            Width           =   840
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "접수일자:"
            Height          =   225
            Index           =   2
            Left            =   45
            TabIndex        =   6
            Top             =   105
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "~"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   2415
            TabIndex        =   5
            Top             =   105
            Width           =   120
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   7
         Top             =   15
         Width           =   14370
         _ExtentX        =   25347
         _ExtentY        =   741
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   4194304
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
         Caption         =   "      고객 미출고 현황"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm고객미출고현황.frx":3691
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm고객미출고현황.frx":38B7
            Top             =   -15
            Width           =   765
         End
      End
   End
End
Attribute VB_Name = "frm고객미출고현황"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0
            Call DataPrint_미출고내역
        Case 3:
            Call Export_Excel(frmMain.cdgExcel, sprGrid)
        
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
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub Data_Print(Print_PreView As Boolean)
    On Error GoTo ErrRtn
    
    If sprGrid.MaxRows = 0 Then Exit Sub

    If Dir(AppPath & "XML", vbDirectory) = "" Then
        MkDir AppPath & "XML"
    End If

    Open AppPath & "XML\고객미출고현황.XML" For Output As #1
    
    Print #1, "<?xml version=""1.0"" encoding=""EUC-KR""?>"
    Print #1, "<root>"
    
          XML = "    <조건>"
    XML = XML & "        <검색조건>접수일자 : " & dtpDay(0).Value & " ~ " & dtpDay(1).Value & "</검색조건>"
    XML = XML & "        <가맹점>크린에이드 - " & Func_Replace(가맹점정보.가맹점명) & "</가맹점>"
    XML = XML & "   </조건>"
    Print #1, XML
    
    With sprGrid
        For i = 1 To .MaxRows
            .Row = i
            
                             XML = "    <Data>"
            .Col = 1:  XML = XML & "        <접수일자>" & .Text & "</접수일자>"
            .Col = 2:  XML = XML & "        <예정일자>" & Func_Replace(.Text) & "</예정일자>"
            .Col = 3:  XML = XML & "        <지사출고일>" & Func_Replace(.Text) & "</지사출고일>"
            .Col = 4:  XML = XML & "        <성명>" & Func_Replace(.Text) & "</성명>"
            .Col = 5:  XML = XML & "        <전화번호>" & Func_Replace(.Text) & "</전화번호>"
            .Col = 6:  XML = XML & "        <휴대전화>" & Func_Replace(.Text) & "</휴대전화>"
            .Col = 7:  XML = XML & "        <품명>" & Func_Replace(.Text) & "</품명>"
            .Col = 8:  XML = XML & "        <택번호>" & .Text & "</택번호>"
            .Col = 9:  XML = XML & "        <색상>" & Func_Replace(.Text) & "</색상>"
            .Col = 10: XML = XML & "        <무늬>" & Func_Replace(.Text) & "</무늬>"
            .Col = 11: XML = XML & "        <내용>" & Func_Replace(.Text) & "</내용>"
            .Col = 12: XML = XML & "        <금액>" & .Text & "</금액>"
            .Col = 13: XML = XML & "        <결제>" & Func_Replace(.Text) & "</결제>"
            .Col = 14: XML = XML & "        <상표>" & Func_Replace(.Text) & "</상표>"
                       XML = XML & "   </Data>"
                       Print #1, XML
        Next i
        
        Print #1, "</root>"
        Close #1
    End With
    
    If Print_PreView = True Then
        With rpt고객미출고현황
            .dc.FileURL = AppPath & "XML\고객미출고현황.XML"
            .Show 1
        End With
    Else
        With rpt고객미출고현황
            .dc.FileURL = AppPath & "XML\고객미출고현황.XML"
            .PrintReport False
        End With
    
        Unload rpt고객미출고현황
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub cmdList_Click()
    Call Data_Display
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyChk (KeyCode)
    
    If KeyCode = 13 Then
        SendKeys "{Tab}"
    End If
End Sub

Private Sub Form_Load()
    With sprGrid
        .MaxRows = 0
        .RowHeight(-1) = 14
        
        .Col = 1: .ColMerge = MergeRestricted
        .Col = 2: .ColMerge = MergeRestricted
        .Col = 3: .ColMerge = MergeRestricted
        .Col = 4: .ColMerge = MergeRestricted
        .Col = 5: .ColMerge = MergeRestricted
        .Col = 6: .ColMerge = MergeRestricted
        
        .ColsFrozen = 6
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeExtended
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
    
    With cboGubun
        .Clear
        .AddItem "성명"
        .AddItem "전화번호"
        .AddItem "주소"
        
        .ListIndex = 0
    End With
    
    dtpDay(0).Value = Format(Date, "YYYY-MM-DD")
    dtpDay(1).Value = Format(Date, "YYYY-MM-DD")
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn
    
    Query = "SELECT    A.접수일자"
    Query = Query & ", B.휴대전화"
    Query = Query & ", B.전화번호"
    Query = Query & ", B.성명"
    Query = Query & ", A.의류명"
    Query = Query & ", A.택번호"
    Query = Query & ", A.색상"
    Query = Query & ", A.무늬"
    Query = Query & ", A.내용"
    Query = Query & ", A.금액"
    Query = Query & ", A.결제여부"
    Query = Query & ", A.상표"
    Query = Query & ", A.예정일자"
    Query = Query & ", A.지사출고일자"
    Query = Query & " FROM TB_입출고 AS A LEFT OUTER JOIN TB_고객정보 AS B ON A.고객코드 = B.고객코드"
    Query = Query & " WHERE (A.접수일자 >= '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "'"
    Query = Query & "   AND  A.접수일자 <= '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "')"
    Query = Query & "   AND (A.세탁환불일자 IS NULL OR A.세탁환불일자 = '')"
    Query = Query & "   AND (A.반품환불일자 IS NULL OR A.반품환불일자 = '')"
    Query = Query & "   AND RTRIM(출고일자) = '' "

    If txtData(0).Text <> "" Then
        Select Case cboGubun.Text
            Case "성명":
                Query = Query & " AND B.성명 LIKE '%" & txtData(0).Text & "%'"
                    
            Case "전화번호":
                Query = Query & " AND ( B.휴대전화  LIKE '%" & txtData(0).Text & "%'"
                Query = Query & "  OR B.전화번호 LIKE '%" & txtData(0).Text & "%')"
        
            Case "주소":
                Query = Query & " AND B.주소 LIKE '%" & txtData(0).Text & "%'"
        End Select
    End If
    
    Query = Query & "   AND  A.판매취소 <> 'Y'"
    Query = Query & " ORDER BY A.접수일자, B.성명, A.택번호 ASC"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        Do While Not ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = Format(ADORs!접수일자, "YYYY-MM-DD") & ""               ' 1
            .Col = 2:  .Text = Format(ADORs!예정일자, "YYYY-MM-DD") & ""               ' 2
            
            If ADORs!지사출고일자 = "" Then
                .Col = 3: .Text = " "                                                  ' 3
            Else
                .Col = 3: .Text = Left(ADORs!지사출고일자, 10) & ""                    ' 3
            End If
            
            .Col = 4:  .Text = ADORs!성명 & ""                                         ' 2
            .Col = 5:  .Text = ADORs!전화번호 & ""                                     ' 3
            .Col = 6:  .Text = ADORs!휴대전화 & ""                                     ' 4
            .Col = 7:  .Text = ADORs!의류명 & ""                                       ' 5
            
            If Len(Trim(ADORs!택번호)) <= 6 Then
                .Col = 8: .Text = ADORs!택번호                                         ' 6
            Else
                .Col = 8: .Text = Format(ADORs!택번호, "000-00-0000") ' 6
            End If
            
            .Col = 9:  .Text = ADORs!색상 & ""                                         ' 7
            .Col = 10: .Text = ADORs!무늬 & ""                                         ' 8
            .Col = 11: .Text = ADORs!내용 & ""                                         ' 9
            .Col = 12: .Text = ADORs!금액 & ""                                         '10
            .Col = 13: .Text = ADORs!결제여부 & ""                                     '11
            .Col = 14: .Text = ADORs!상표 & ""                                         '12
        
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    cmdBtn(5).Left = (Me.Width - 200) - cmdBtn(5).Width
End Sub



Private Sub DataPrint_미출고내역()
    On Error GoTo ErrRtn
    
'    Dim nRow        As Long
'    Dim 미출고수량  As Long
'    Dim CommPort    As String
'    Dim BaudRate    As String
'
'    Dim tmp         As String
'    Dim PrintStr    As String
'    Dim nGoodsLng   As Integer
'
'
'    nGoodsLng = 15
'    CommPort = GetIniStr("VAN", "KS7500_CommPort", "", iniFile)
'    BaudRate = GetIniStr("VAN", "KS7500_BaudRate", "", iniFile)
'
'    Rtn = KS7500i.CheckPort(CInt(CommPort), CLng(BaudRate))
'    DoEvents
'
'    If Rtn < 0 Then
'        nRow = nRow + 1
'
'        If nRow > 3 Then
'            MsgBox "카드단말기 장치가 연결되어 있지 않습니다", vbCritical, "오류"
'
'            Exit Sub
'        End If
'    End If
'
'    Call KS7500i.SetConfig("", Rtn, CLng(BaudRate))    '첫번째 인자는 "" 로 넣어 준다.
'
'    KS7500i.InitPrint
'    DoEvents
'
'
'    '--------------------------------------------------------------------------------------------------------
'    Query = "SELECT    ISNULL(가맹점명, '')     AS 가맹점명"
'    Query = Query & ", ISNULL(매장전화번호, '') AS 매장전화번호"
'    Query = Query & ", ISNULL(사업장주소, '')   AS 사업장주소"
'    Query = Query & ", ISNULL(사업자번호, '')   AS 사업자번호"
'    Query = Query & ", ISNULL(대표자명, '')     AS 대표자명"
'    Query = Query & " FROM TB_기본정보"
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    If ADORs.EOF Then
'        Call KS7500i.PrintString("상 호 명 :  미출고 내역", 1)
'    Else
'        Call KS7500i.PrintString("상 호 명 : " + ADORs!가맹점명 & " 미출고 내역", 1)
'    End If
'    ADORs.Close
'    Set ADORs = Nothing
'
'    Call KS7500i.PrintString("===============================================", 1)
'    Call KS7500i.PrintString("접수일자 : " + Format(dtpDay(0).Value, "yyyy-MM-dd") + " ~ " + Format(dtpDay(1).Value, "yyyy-MM-dd"), 1)
'    Call KS7500i.PrintString("출력일자 : " + Format(Now(), "yyyy-MM-dd hh:mm:ss"), 1)
'    Call KS7500i.PrintString("-----------------------------------------------", 1)
'    Call KS7500i.PrintString("택번호  의류            택번호  의류           ", 1)
'    Call KS7500i.PrintString("-----------------------------------------------", 1)
'
'    미출고수량 = 0
'
'    With sprGrid
'        For nRow = 1 To .MaxRows
'            .Row = nRow
'
'            ' 택번호 확인
'            .Col = 8
'            If Trim(.Text) = "" Then Exit For
'            미출고수량 = 미출고수량 + 1
'
'
'            '*********************************************************
'            '* 택번호
'            '*********************************************************
'            .Col = 8: PrintStr = Mid(.Text, 5) + " "
'
'            '*********************************************************
'            '* 품명
'            '*********************************************************
'            .Col = 7
'            If LenH(.Text) >= nGoodsLng Then
'                tmp = MidH(.Text, 1, nGoodsLng)
'            Else
'                tmp = Trim(.Text) + String(nGoodsLng - LenH(.Text), " ")
'            End If
'
'            PrintStr = PrintStr + tmp + " "
'
'            '가장 마지막을 찍을때
'            If nRow > .MaxRows Then
'                Call KS7500i.PrintString(PrintStr, 1)
'                Exit For
'            End If
'
'            ' 다음줄( 한라인에 2개를 찍는다.)
'            nRow = nRow + 1
'            .Row = nRow
'
'            ' 택번호 확인
'            .Col = 8
'            If Trim(.Text) = "" Then Exit For
'            미출고수량 = 미출고수량 + 1
'
'
'            '*********************************************************
'            '* 택번호
'            '*********************************************************
'            .Col = 8: PrintStr = PrintStr + Mid(.Text, 5) + " "
'
'            '*********************************************************
'            '* 품명
'            '*********************************************************
'            .Col = 7
'            If LenH(.Text) >= nGoodsLng Then
'                tmp = MidH(.Text, 1, nGoodsLng)
'            Else
'                tmp = Trim(.Text) + String(nGoodsLng - LenH(.Text), " ")
'            End If
'
'            PrintStr = PrintStr + tmp + " "
'
'
'            '*********************************************************
'            '* 출력
'            '*********************************************************
'            Call KS7500i.PrintString(PrintStr, 1)
'        Next nRow
'    End With
'
'    Call KS7500i.PrintString("-----------------------------------------------", 1)
'    Call KS7500i.PrintString("총 미출고 수량 : " + Format(미출고수량, "#,##0"), 1)
'    Call KS7500i.PrintString("===============================================", 1)
'
'    KS7500i.LineFeed (1)
'    KS7500i.CutPaper
'
'    KS7500i.ClosePort
'    DoEvents
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

