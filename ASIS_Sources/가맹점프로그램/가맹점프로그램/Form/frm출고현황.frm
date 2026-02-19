VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm출고현황 
   Caption         =   "출고현황"
   ClientHeight    =   11970
   ClientLeft      =   660
   ClientTop       =   2475
   ClientWidth     =   16410
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form20"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11970
   ScaleWidth      =   16410
   WindowState     =   2  '최대화
   Begin Threed.SSPanel pnlProg 
      Height          =   1215
      Left            =   555
      TabIndex        =   8
      Top             =   1935
      Visible         =   0   'False
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   2143
      _Version        =   262144
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
      Picture         =   "frm출고현황.frx":0000
      Caption         =   " "
      BorderWidth     =   0
      BevelOuter      =   1
      BevelInner      =   2
      Alignment       =   4
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   11970
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   16410
      _ExtentX        =   28945
      _ExtentY        =   21114
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm출고현황.frx":2FCB
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   10395
         Left            =   15
         TabIndex        =   0
         Top             =   1560
         Width           =   16380
         _Version        =   524288
         _ExtentX        =   28893
         _ExtentY        =   18336
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
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         MaxCols         =   15
         MaxRows         =   200
         MoveActiveOnFocus=   0   'False
         OperationMode   =   1
         Protect         =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frm출고현황.frx":303D
         UserResize      =   1
         VisibleCols     =   12
         VisibleRows     =   50
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   1095
         Left            =   15
         TabIndex        =   2
         Top             =   450
         Width           =   16380
         _ExtentX        =   28893
         _ExtentY        =   1931
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin CSTextLibCtl.sitxEdit txtTAGNo 
            Height          =   315
            Left            =   6225
            TabIndex        =   21
            Top             =   750
            Width           =   945
            _Version        =   262145
            _ExtentX        =   1667
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   "__-____"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            EOLTab          =   -1  'True
            FocusSelect     =   -1  'True
            Insert          =   0   'False
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   "__-____"
            StartText.x     =   3
            StartText.y     =   2
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   16
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   "##-####"
            Justification   =   1
            CharacterTable  =   ""
            BorderStyle     =   0
            Characters      =   2
            MaxLength       =   6
         End
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
            Top             =   750
            Width           =   1455
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
            Left            =   2415
            TabIndex        =   9
            Top             =   750
            Width           =   2400
         End
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   630
            Left            =   8340
            TabIndex        =   3
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 조회(&F)"
            Appearance      =   6
            Picture         =   "frm출고현황.frx":3AA8
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   3
            Left            =   9885
            TabIndex        =   4
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 엑셀(&E)"
            Appearance      =   6
            Picture         =   "frm출고현황.frx":41A2
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   13170
            TabIndex        =   5
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frm출고현황.frx":491C
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   4
            Left            =   11430
            TabIndex        =   6
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 출력(&P)"
            Appearance      =   6
            Picture         =   "frm출고현황.frx":59AE
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   0
            Left            =   915
            TabIndex        =   11
            Top             =   405
            Width           =   1455
            _ExtentX        =   2566
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
            Format          =   58916867
            CurrentDate     =   40279
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   1
            Left            =   2640
            TabIndex        =   12
            Top             =   405
            Width           =   1455
            _ExtentX        =   2566
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
            Format          =   58916867
            CurrentDate     =   40279
         End
         Begin XtremeSuiteControls.PushButton btnTerm 
            Height          =   345
            Index           =   0
            Left            =   915
            TabIndex        =   16
            Top             =   30
            Width           =   705
            _Version        =   851970
            _ExtentX        =   1244
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "1개월"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnTerm 
            Height          =   345
            Index           =   1
            Left            =   1650
            TabIndex        =   17
            Top             =   30
            Width           =   705
            _Version        =   851970
            _ExtentX        =   1244
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "1주일"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnTerm 
            Height          =   345
            Index           =   2
            Left            =   2385
            TabIndex        =   18
            Top             =   30
            Width           =   705
            _Version        =   851970
            _ExtentX        =   1244
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "오늘"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "택번호:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   5355
            TabIndex        =   20
            Top             =   810
            Width           =   840
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "조회기간:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   45
            TabIndex        =   19
            Top             =   120
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
            Index           =   0
            Left            =   2430
            TabIndex        =   15
            Top             =   465
            Width           =   120
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
            Height          =   180
            Index           =   3
            Left            =   45
            TabIndex        =   14
            Top             =   810
            Width           =   840
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "출고일자:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   45
            TabIndex        =   13
            Top             =   465
            Width           =   840
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   7
         Top             =   15
         Width           =   16380
         _ExtentX        =   28893
         _ExtentY        =   741
         _Version        =   262144
         Font3D          =   3
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
         Caption         =   "      출고 현황"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm출고현황.frx":60A8
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm출고현황.frx":62CE
            Top             =   -15
            Width           =   765
         End
      End
   End
End
Attribute VB_Name = "frm출고현황"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnTerm_Click(Index As Integer)
    Select Case Index
        Case 0:
            dtpDay(0).Value = Format(DateAdd("m", -1, Date), "YYYY-MM-DD")
            dtpDay(1).Value = Format(Date, "YYYY-MM-DD")
        
        Case 1:
            dtpDay(0).Value = Format(DateAdd("ww", -1, Date), "YYYY-MM-DD")
            dtpDay(1).Value = Format(Date, "YYYY-MM-DD")
        
        Case 2:
            dtpDay(0).Value = Format(Date, "YYYY-MM-DD")
            dtpDay(1).Value = Format(Date, "YYYY-MM-DD")
    End Select
    
    Call Data_Display
End Sub


Private Sub Data_Print(Print_PreView As Boolean)
    On Error GoTo ErrRtn
    
    If sprGrid.MaxRows = 0 Then Exit Sub

    If Dir(AppPath & "XML", vbDirectory) = "" Then
        MkDir AppPath & "XML"
    End If

    Open AppPath & "XML\출고현황.XML" For Output As #1
    
    Print #1, "<?xml version=""1.0"" encoding=""EUC-KR""?>"
    Print #1, "<root>"
    
          XML = "    <조건>"
    XML = XML & "        <출고일자>출고일자 : " & Format(dtpDay(0).Value, "YYYY-MM-DD") & " ~ " & Format(dtpDay(1).Value, "YYYY-MM-DD") & "</출고일자>"
    XML = XML & "        <가맹점>크린에이드 - " & Func_Replace(가맹점정보.가맹점명) & "</가맹점>"
    XML = XML & "   </조건>"
    Print #1, XML
    
    With sprGrid
        For i = 1 To .MaxRows
            .Row = i
            
                             XML = "    <Data>"
            .Col = 1:  XML = XML & "        <출고일자>" & .Text & "</출고일자>"
            .Col = 3:  XML = XML & "        <접수일자>" & .Text & "</접수일자>"
            .Col = 4:  XML = XML & "        <성명>" & Func_Replace(.Text) & "</성명>"
            .Col = 5:  XML = XML & "        <휴대전화>" & Func_Replace(.Text) & "</휴대전화>"
            .Col = 6:  XML = XML & "        <전화번호>" & Func_Replace(.Text) & "</전화번호>"
            .Col = 7:  XML = XML & "        <의류코드>" & Func_Replace(.Text) & "</의류코드>"
            .Col = 8:  XML = XML & "        <품명>" & Func_Replace(.Text) & "</품명>"
            .Col = 9:  XML = XML & "        <택번호>" & .Text & "</택번호>"
            .Col = 10:  XML = XML & "        <색상>" & Func_Replace(.Text) & "</색상>"
            .Col = 11: XML = XML & "        <무늬>" & Func_Replace(.Text) & "</무늬>"
            .Col = 12: XML = XML & "        <내용>" & Func_Replace(.Text) & "</내용>"
            .Col = 13: XML = XML & "        <금액>" & .Text & "</금액>"
            .Col = 14: XML = XML & "        <결제>" & Func_Replace(.Text) & "</결제>"
            .Col = 15: XML = XML & "        <상표>" & Func_Replace(.Text) & "</상표>"
                       XML = XML & "   </Data>"
                       Print #1, XML
        Next i
        
        Print #1, "</root>"
        Close #1
    End With
    
    If Print_PreView = True Then
        With rpt출고현황
            .dc.FileURL = AppPath & "XML\출고현황.XML"
            .Show 1
        End With
    Else
        With rpt출고현황
            .dc.FileURL = AppPath & "XML\출고현황.XML"
            .PrintReport False
        End With
    
        Unload rpt출고현황
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

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
            
        Case 5: Unload Me
    End Select
End Sub

Private Sub cmdList_Click()
    Call Data_Display
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn
        
    '---------------------------------------------------------------------------------------------------
    ' TB_고객정보
    '---------------------------------------------------------------------------------------------------
    Query = "SELECT    B.성명"
    Query = Query & ", B.휴대전화"
    Query = Query & ", B.전화번호"
    Query = Query & ", A.접수일자"
    Query = Query & ", A.출고시간"
    Query = Query & ", A.접수번호"
    Query = Query & ", A.출고일자"
    Query = Query & ", A.지사출고상태"
    Query = Query & ", A.의류코드"
    Query = Query & ", A.의류명"
    Query = Query & ", A.택번호"
    Query = Query & ", A.색상"
    Query = Query & ", A.무늬"
    Query = Query & ", A.내용"
    Query = Query & ", A.금액"
    Query = Query & ", A.결제여부"
    Query = Query & ", A.상표"
    Query = Query & " FROM TB_입출고 AS A LEFT OUTER JOIN TB_고객정보 AS B ON A.고객코드 = B.고객코드"
    Query = Query & " WHERE (A.출고일자 >= '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "'"
    Query = Query & "   AND  A.출고일자 <= '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "')"
    
    If txtTAGNo.RawData <> "" Then
        Query = Query & " AND SUBSTRING(A.택번호,4,6) = '" & txtTAGNo.RawData & "'"
    End If
    
    If Trim(txtFind.Text) <> "" Then
        Select Case cboGubun.Text
            Case "성명":     Query = Query & " AND (B.성명 LIKE '%" & Trim(txtFind.Text) & "%') "
            Case "전화번호": Query = Query & " AND (B.전화번호 LIKE '%" & Trim(txtFind.Text) & "%') "
            Case "고객코드": Query = Query & " AND (B.고객코드 LIKE '%" & Trim(txtFind.Text) & "%') "
        End Select
    End If
    
    Query = Query & "   AND (A.판매취소 IS NULL OR A.판매취소 <> 'Y')"
    Query = Query & " ORDER BY A.출고일자, A.택번호"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = ADORs!출고일자 & ""
            .Col = 2:  .Text = ADORs!출고시간 & ""
            .Col = 3:  .Text = ADORs!접수일자 & ""
            .Col = 4:  .Text = ADORs!성명 & ""
            .Col = 5:  .Text = ADORs!휴대전화 & ""
            .Col = 6:  .Text = ADORs!전화번호 & ""
            .Col = 7:  .Text = ADORs!의류코드 & ""
            .Col = 8:  .Text = ADORs!의류명 & ""
            
            If Len(Trim(ADORs!택번호)) <= 6 Then
                .Col = 9: .Text = ADORs!택번호 & ""
            Else
                .Col = 9: .Text = Format(ADORs!택번호, "000-00-0000")
            End If
                        
            .Col = 10:  .Text = ADORs!색상 & ""
            .Col = 11: .Text = ADORs!무늬 & ""
            .Col = 12: .Text = ADORs!내용 & ""
            .Col = 13: .Text = ADORs!금액 & ""
            .Col = 14: .Text = ADORs!결제여부 & ""
            .Col = 15: .Text = ADORs!상표 & ""
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub Form_Activate()
    dtpDay(1).Value = Format(Date, "YYYY-MM-DD")
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
        KeyAscii = 0
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
    On Error Resume Next
    
    cmdBtn(5).Left = Me.Width - cmdBtn(5).Width - 200
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdList_Click
    End If
End Sub
