VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm고객별매출액 
   Caption         =   "고객별 매출현황"
   ClientHeight    =   8850
   ClientLeft      =   1155
   ClientTop       =   5295
   ClientWidth     =   15360
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   10.5
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8850
   ScaleWidth      =   15360
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8850
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   15610
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm고객별매출액.frx":0000
      Begin Threed.SSPanel SSPanel1 
         Height          =   750
         Left            =   15
         TabIndex        =   4
         Top             =   450
         Width           =   15330
         _ExtentX        =   27040
         _ExtentY        =   1323
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox cboGubun 
            BackColor       =   &H00E0E0E0&
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
            TabIndex        =   12
            Top             =   405
            Width           =   1425
         End
         Begin VB.TextBox txtData 
            BackColor       =   &H00FFFFFF&
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
            IMEMode         =   10  '한글 
            Index           =   0
            Left            =   2370
            TabIndex        =   11
            Top             =   405
            Width           =   1860
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   0
            Left            =   915
            TabIndex        =   0
            Top             =   60
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
            Format          =   58916867
            CurrentDate     =   40279
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   1
            Left            =   2610
            TabIndex        =   1
            Top             =   60
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
            Format          =   58916867
            CurrentDate     =   40279
         End
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   630
            Left            =   8115
            TabIndex        =   2
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
            Picture         =   "frm고객별매출액.frx":00B2
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   3
            Left            =   9660
            TabIndex        =   8
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
            Picture         =   "frm고객별매출액.frx":07AC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   12915
            TabIndex        =   9
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
            Picture         =   "frm고객별매출액.frx":0F26
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   4
            Left            =   11205
            TabIndex        =   10
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
            Picture         =   "frm고객별매출액.frx":1FB8
         End
         Begin VB.Label Label1 
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
            Height          =   210
            Index           =   1
            Left            =   45
            TabIndex        =   13
            Top             =   480
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
            Left            =   2415
            TabIndex        =   7
            Top             =   105
            Width           =   120
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "접수일자:"
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
            Index           =   2
            Left            =   45
            TabIndex        =   6
            Top             =   105
            Width           =   840
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   5
         Top             =   15
         Width           =   15330
         _ExtentX        =   27040
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
         Caption         =   "      고객별 매출현황"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm고객별매출액.frx":26B2
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm고객별매출액.frx":28D8
            Top             =   -15
            Width           =   765
         End
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   4440
         Left            =   15
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1215
         Width           =   15330
         _Version        =   524288
         _ExtentX        =   27040
         _ExtentY        =   7832
         _StockProps     =   64
         AllowMultiBlocks=   -1  'True
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
         FormulaSync     =   0   'False
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         MaxCols         =   6
         Protect         =   0   'False
         RestrictCols    =   -1  'True
         RestrictRows    =   -1  'True
         ScrollBars      =   2
         ShadowText      =   0
         SpreadDesigner  =   "frm고객별매출액.frx":34A2
         VisibleCols     =   5
         VisibleRows     =   500
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread sprList 
         Height          =   2790
         Left            =   15
         TabIndex        =   15
         Top             =   6045
         Width           =   15330
         _Version        =   524288
         _ExtentX        =   27040
         _ExtentY        =   4921
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
         MaxCols         =   11
         MaxRows         =   200
         MoveActiveOnFocus=   0   'False
         OperationMode   =   1
         Protect         =   0   'False
         ScrollBarMaxAlign=   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frm고객별매출액.frx":3BED
         UserResize      =   1
         VisibleCols     =   12
         VisibleRows     =   50
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   360
         Left            =   15
         TabIndex        =   16
         Top             =   5670
         Width           =   15330
         _ExtentX        =   27040
         _ExtentY        =   635
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
         Caption         =   " 고객별 접수현황"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm고객별매출액.frx":44F0
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
End
Attribute VB_Name = "frm고객별매출액"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyChk (KeyCode)
    
    If KeyCode = 13 Then
        SendKeys "{Tab}"
        KeyCode = 0
    End If
End Sub

Private Sub Form_Load()
    With sprGrid
        .MaxRows = 0
        .RowHeight(-1) = 14
        
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
        
        .Row = SpreadHeader
        .Col = SpreadHeader: .Text = "순위"
    End With
    
    With sprList
        .MaxRows = 0
        .RowHeight(-1) = 14
        
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
        
    '---------------------------------------------------------------------------------------------
    '
    '---------------------------------------------------------------------------------------------
    Query = "SELECT    A.고객코드"
    Query = Query & ", B.성명"
    Query = Query & ", B.휴대전화"
    Query = Query & ", B.전화번호"
    Query = Query & ", B.주소"
    Query = Query & ", SUM(A.금액) AS 매출액"
    Query = Query & " FROM TB_입출고 AS A LEFT OUTER JOIN TB_고객정보 AS B ON A.고객코드 = B.고객코드"
    Query = Query & " WHERE (SUBSTRING(A.접수일자,1,10) >= '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "'"
    Query = Query & "   AND  SUBSTRING(A.접수일자,1,10) <= '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "')"
    
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
    
    Query = Query & " GROUP BY A.고객코드, B.성명, B.휴대전화, B.전화번호, B.주소"
    Query = Query & " ORDER BY SUM(A.금액) DESC"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
      
    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        
            .Col = 1: .Text = ADORs!고객코드 & "" ' 1
            .Col = 2: .Text = ADORs!성명 & ""     ' 2
            .Col = 3: .Text = ADORs!전화번호 & "" ' 3
            .Col = 4: .Text = ADORs!휴대전화 & ""   ' 4
            .Col = 5: .Text = ADORs!주소 & ""     ' 5
            .Col = 6: .Text = ADORs!매출액 & ""   ' 6
        
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

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
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

    Open AppPath & "XML\고객별매출액.XML" For Output As #1
    
    Print #1, "<?xml version=""1.0"" encoding=""EUC-KR""?>"
    Print #1, "<root>"
    
          XML = "    <조건>"
    XML = XML & "        <검색조건>접수일자 : " & Format(dtpDay(0).Value, "YYYY-MM") & " ~ " & Format(dtpDay(1).Value, "YYYY-MM") & "</검색조건>"
    XML = XML & "        <가맹점>크린에이드 - " & Func_Replace(가맹점정보.가맹점명) & "</가맹점>"
    XML = XML & "   </조건>"
    Print #1, XML
    
    With sprGrid
        For i = 1 To .MaxRows
            .Row = i
            
                             XML = "    <Data>"
            .Col = 1:  XML = XML & "        <고객코드>" & .Text & "</고객코드>"
            .Col = 2:  XML = XML & "        <성명>" & Func_Replace(.Text) & "</성명>"
            .Col = 3:  XML = XML & "        <전화번호>" & Func_Replace(.Text) & "</전화번호>"
            .Col = 4:  XML = XML & "        <휴대전화>" & Func_Replace(.Text) & "</휴대전화>"
            .Col = 5:  XML = XML & "        <주소>" & Func_Replace(.Text) & "</주소>"
            .Col = 6:  XML = XML & "        <매출액>" & .Text & "</매출액>"
                       XML = XML & "   </Data>"
                       Print #1, XML
        Next i
        
        Print #1, "</root>"
        Close #1
    End With
    
    If Print_PreView = True Then
        With rpt고객별매출액
            .dc.FileURL = AppPath & "XML\고객별매출액.XML"
            .Show 1
        End With
    Else
        With rpt고객별매출액
            .dc.FileURL = AppPath & "XML\고객별매출액.XML"
            .PrintReport False
        End With
        
        Unload rpt고객별매출액
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub cmdList_Click()
    Call Data_Display
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    cmdBtn(5).Left = (Me.Width - 200) - cmdBtn(5).Width
End Sub

Private Sub sprGrid_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 0 Then
        Exit Sub
    End If
    
    sprGrid.Row = Row
    sprGrid.Col = 1
    
    Call 입출고_Display(sprGrid.Text)
End Sub

Private Sub sprGrid_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Call sprGrid_Click(NewCol, NewRow)
End Sub

Private Sub 입출고_Display(고객코드 As String)
    On Error GoTo ErrRtn
    
    Query = "SELECT    A.성명"
    Query = Query & ", A.휴대전화"
    Query = Query & ", A.전화번호"
    Query = Query & ", B.접수일자"
    Query = Query & ", B.출고일자"
    Query = Query & ", B.지사출고상태"
    Query = Query & ", B.의류명"
    Query = Query & ", B.택번호"
    Query = Query & ", B.색상"
    Query = Query & ", B.무늬"
    Query = Query & ", B.내용"
    Query = Query & ", B.금액"
    Query = Query & ", B.결제여부"
    Query = Query & ", B.상표"
    Query = Query & " FROM TB_고객정보 AS A LEFT OUTER JOIN TB_입출고 AS B ON (A.고객코드 = B.고객코드) "
    Query = Query & " WHERE A.고객코드 = '" & 고객코드 & "'"
    Query = Query & "   AND (SUBSTRING(B.접수일자,1,10) >= '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "'"
    Query = Query & "   AND  SUBSTRING(B.접수일자,1,10) <= '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "')"
    Query = Query & "   AND (B.판매취소 IS NULL OR B.판매취소 <> 'Y')"
    Query = Query & " ORDER BY B.접수일자, B.택번호"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    With sprList
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = ADORs!접수일자 & ""
            .Col = 2:  .Text = ADORs!출고일자 & ""
            .Col = 3:  .Text = ADORs!지사출고상태 & ""
            .Col = 4:  .Text = ADORs!의류명 & ""
            
            If Len(Trim(ADORs!택번호)) <= 6 Then
                .Col = 5: .Text = ADORs!택번호 & ""
            Else
                .Col = 5: .Text = Left(ADORs!택번호, 5) & "-" & Right(ADORs!택번호, 4)
            End If
            
            .Col = 6:  .Text = ADORs!색상 & ""
            .Col = 7:  .Text = ADORs!무늬 & ""
            .Col = 8:  .Text = ADORs!내용 & ""
            .Col = 9:  .Text = ADORs!금액 & ""
            .Col = 10: .Text = ADORs!결제여부 & ""
            .Col = 11: .Text = ADORs!상표 & ""
            
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

