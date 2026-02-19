VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm고객조회 
   Caption         =   "고객 조회"
   ClientHeight    =   9270
   ClientLeft      =   1065
   ClientTop       =   1980
   ClientWidth     =   15030
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm고객조회.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   15030
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9270
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   15030
      _ExtentX        =   26511
      _ExtentY        =   16351
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm고객조회.frx":030A
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   7935
         Left            =   15
         TabIndex        =   7
         Top             =   1320
         Width           =   15000
         _Version        =   524288
         _ExtentX        =   26458
         _ExtentY        =   13996
         _StockProps     =   64
         BackColorStyle  =   1
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
         DInformActiveRowChange=   0   'False
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
         GridSolid       =   0   'False
         MaxCols         =   9
         Protect         =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frm고객조회.frx":037C
         VisibleCols     =   3
         VisibleRows     =   30
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   870
         Index           =   1
         Left            =   15
         TabIndex        =   8
         Top             =   435
         Width           =   15000
         _ExtentX        =   26458
         _ExtentY        =   1535
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox cboGubun 
            BackColor       =   &H00E0E0E0&
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
            Left            =   960
            Style           =   2  '드롭다운 목록
            TabIndex        =   1
            Top             =   60
            Width           =   1455
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
            Left            =   2445
            TabIndex        =   0
            Top             =   60
            Width           =   3105
         End
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   750
            Left            =   5865
            TabIndex        =   2
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   79
            Caption         =   " 조회(&F)"
            Appearance      =   6
            Picture         =   "frm고객조회.frx":0B92
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   750
            Index           =   4
            Left            =   11850
            TabIndex        =   4
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   79
            Caption         =   " 출력(&P)"
            Appearance      =   6
            Picture         =   "frm고객조회.frx":128C
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   750
            Index           =   5
            Left            =   13440
            TabIndex        =   5
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frm고객조회.frx":1986
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   750
            Index           =   3
            Left            =   10170
            TabIndex        =   3
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   79
            Caption         =   " 엑셀(&E)"
            Appearance      =   6
            Picture         =   "frm고객조회.frx":2A18
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "검색조건:"
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
            Index           =   4
            Left            =   60
            TabIndex        =   10
            Top             =   90
            Width           =   840
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   405
         Left            =   15
         TabIndex        =   9
         Top             =   15
         Width           =   15000
         _ExtentX        =   26458
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
         Caption         =   "    고객 조회"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm고객조회.frx":3192
         BorderWidth     =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image1 
            Height          =   240
            Left            =   60
            Picture         =   "frm고객조회.frx":35F4
            Top             =   75
            Width           =   240
         End
      End
   End
End
Attribute VB_Name = "frm고객조회"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 3: Call Export_Excel(frmMain.cdgExcel, sprGrid)
        Case 4
                If sprGrid.MaxRows = 0 Then Exit Sub

                If Dir(AppPath & "XML", vbDirectory) = "" Then
                    MkDir AppPath & "XML"
                End If

                Open AppPath & "XML\고객.XML" For Output As #1
                
                Print #1, "<?xml version=""1.0"" encoding=""EUC-KR""?>"
                Print #1, "<root>"
                
                      XML = "    <조건>"
                
                If txtFind.Text = "" Then
                    XML = XML & "        <검색조건>검색조건 : 전체</검색조건>"
                Else
                    XML = XML & "        <검색조건>검색조건 : " & Func_Replace(txtFind.Text) & "</검색조건>"
                End If
                
                XML = XML & "        <가맹점>크린에이드 - " & Func_Replace(가맹점정보.가맹점명) & "</가맹점>"
                XML = XML & "   </조건>"
                Print #1, XML
                
                With sprGrid
                    For i = 1 To .MaxRows
                        .Row = i
                        
                                        XML = "    <Data>"
                        .Col = 1: XML = XML & "        <고객코드>" & .Text & "</고객코드>"
                        .Col = 2: XML = XML & "        <성명>" & Func_Replace(.Text) & "</성명>"
                        .Col = 3: XML = XML & "        <전화번호>" & Func_Replace(.Text) & "</전화번호>"
                        .Col = 4: XML = XML & "        <휴대전화>" & Func_Replace(.Text) & "</휴대전화>"
                        .Col = 5: XML = XML & "        <주소>" & Func_Replace(.Text) & "</주소>"
                        .Col = 6: XML = XML & "        <미수금>" & .Text & "</미수금>"
                        .Col = 7: XML = XML & "        <SMS>" & .Text & "</SMS>"
                        .Col = 8: XML = XML & "        <등록일자>" & .Text & "</등록일자>"
                        .Col = 9: XML = XML & "        <고객등급>" & .Text & "</고객등급>"
                                  XML = XML & "   </Data>"
                                  Print #1, XML
                    Next i
                    
                    Print #1, "</root>"
                    Close #1
                End With
                
                With rpt고객현황
                    .dc.FileURL = AppPath & "XML\고객.XML"
                    '.PrintReport False
                    .Show 1
                End With
                
        Case 5: Unload Me
    End Select
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub cmdList_Click()
    On Error GoTo ErrRtn
    
    Query = "SELECT    고객코드"
    Query = Query & ", 성명"
    Query = Query & ", 전화번호"
    Query = Query & ", 휴대전화"
    Query = Query & ", 주소"
    Query = Query & ", 미수금"
    Query = Query & ", (CASE WHEN SMSSendYN = 'Y' THEN '1' ELSE '0' END) SMS"
    Query = Query & ", 등록일자"
    Query = Query & ", 고객등급코드"
    Query = Query & " FROM TB_고객정보 "
    Query = Query & " WHERE 고객코드 IS NOT NULL"
    
    If txtFind.Text = "" Then
        Query = Query & " ORDER BY 고객코드 ASC"
    Else
        Select Case cboGubun.Text
            Case "성명":     Query = Query & " AND 성명 LIKE '%" & txtFind.Text & "%'"
                             Query = Query & " ORDER BY 성명 ASC"
            
            Case "전화번호": Query = Query & " AND (전화번호 LIKE '%" & txtFind.Text & "%'"
                             Query = Query & "  OR  휴대전화   LIKE '%" & txtFind.Text & "%')"
                             Query = Query & " ORDER BY 전화번호, 휴대전화 ASC"
                             
            Case "주소":     Query = Query & " AND 주소 LIKE '%" & txtFind.Text & "%'"
                             Query = Query & " ORDER BY 주소 ASC"
        End Select
    End If
    
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        
            .Col = 1: .Text = ADORs!고객코드 & ""     ' 1
            .Col = 2: .Text = ADORs!성명 & ""         ' 2
            .Col = 3: .Text = ADORs!전화번호 & ""     ' 3
            .Col = 4: .Text = ADORs!휴대전화 & ""       ' 4
            .Col = 5: .Text = ADORs!주소 & ""         ' 5
            .Col = 6: .Text = ADORs!미수금 & ""       ' 6
            .Col = 7: .Text = ADORs!SMS & ""          ' 7
            .Col = 8: .Text = ADORs!등록일자 & ""     ' 8
            .Col = 9: .Text = ADORs!고객등급코드 & "" ' 9
            
            ADORs.MoveNext
        Loop
        
        .ReDraw = True
    
        ADORs.Close
        Set ADORs = Nothing
    End With

    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyChk (KeyCode)
    
    If KeyCode = 13 Then
        SendKeys "{TAB}"
        KeyCode = 0
    End If
End Sub

Private Sub Form_Load()
    'TitleSet "고객조회/삭제"
    
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
        .OperationMode = OperationModeSingle
        
        '홀수/짝수 Row BankColor
        'Ret = .SetOddEvenRowColor(&HFFFFFF, &H80000008, &H80FFFF, &H80000008)

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
    
End Sub

Private Sub Form_Resize()
    On Error GoTo ErrRtn
    
    cmdBtn(5).Left = Me.Width - cmdBtn(5).Width - 200
    
    Exit Sub
    
ErrRtn:

End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        cmdList_Click
    End If
End Sub
