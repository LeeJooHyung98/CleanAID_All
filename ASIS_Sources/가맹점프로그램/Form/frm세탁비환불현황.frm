VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm세탁비환불현황 
   Caption         =   "세탁비 환불현황"
   ClientHeight    =   9180
   ClientLeft      =   2745
   ClientTop       =   2370
   ClientWidth     =   14295
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form16"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   14295
   ShowInTaskbar   =   0   'False
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9180
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   16193
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm세탁비환불현황.frx":0000
      Begin FPSpreadADO.fpSpread sprGrid 
         Bindings        =   "frm세탁비환불현황.frx":0072
         Height          =   7605
         Left            =   15
         TabIndex        =   1
         Top             =   1560
         Width           =   14265
         _Version        =   524288
         _ExtentX        =   25162
         _ExtentY        =   13414
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
         SpreadDesigner  =   "frm세탁비환불현황.frx":0086
         VisibleCols     =   9
         VisibleRows     =   200
         HighlightHeaders=   1
         HighlightStyle  =   1
         ScrollBarStyle  =   2
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   1095
         Left            =   15
         TabIndex        =   2
         Top             =   450
         Width           =   14265
         _ExtentX        =   25162
         _ExtentY        =   1931
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSOption optGubun 
            Height          =   240
            Index           =   0
            Left            =   915
            TabIndex        =   15
            Top             =   90
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   423
            _Version        =   262144
            Caption         =   "세탁환불"
            Value           =   -1
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
            TabIndex        =   9
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
            TabIndex        =   8
            Top             =   750
            Width           =   2400
         End
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   630
            Left            =   6915
            TabIndex        =   3
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 조회(&F)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm세탁비환불현황.frx":0959
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   3
            Left            =   8460
            TabIndex        =   4
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 엑셀(&E)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm세탁비환불현황.frx":1053
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   11745
            TabIndex        =   5
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm세탁비환불현황.frx":17CD
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   4
            Left            =   10005
            TabIndex        =   6
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 출력(&P)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm세탁비환불현황.frx":285F
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   0
            Left            =   915
            TabIndex        =   10
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
            Format          =   67436547
            CurrentDate     =   40279
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   1
            Left            =   2655
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
            Format          =   67436547
            CurrentDate     =   40279
         End
         Begin Threed.SSOption optGubun 
            Height          =   240
            Index           =   1
            Left            =   2295
            TabIndex        =   16
            Top             =   90
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   423
            _Version        =   262144
            Caption         =   "반품환불"
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "환불구분:"
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
            TabIndex        =   17
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
            Left            =   2445
            TabIndex        =   14
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
            TabIndex        =   13
            Top             =   810
            Width           =   840
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "환불일자:"
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
            TabIndex        =   12
            Top             =   465
            Width           =   840
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   7
         Top             =   15
         Width           =   14265
         _ExtentX        =   25162
         _ExtentY        =   741
         _Version        =   262144
         Font3D          =   1
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "      세탁비 환불현황"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm세탁비환불현황.frx":2F59
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm세탁비환불현황.frx":317F
            Top             =   -15
            Width           =   765
         End
      End
   End
End
Attribute VB_Name = "frm세탁비환불현황"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyChk (KeyCode)
    
    If KeyCode = 13 Then
        SendKeys "{Tab}"
    End If
End Sub

Private Sub Form_Load()
    'TitleSet "세탁비 환불 현황"
    
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
    
    Query = "SELECT    SUBSTRING(A.세탁환불일자,1,10) AS 세탁환불일자"
    Query = Query & ", SUBSTRING(A.반품환불일자,1,10) AS 반품환불일자"
    Query = Query & ", A.접수일자"
    Query = Query & ", B.휴대전화"
    Query = Query & ", B.전화번호"
    Query = Query & ", B.성명"
    Query = Query & ", A.의류명"
    Query = Query & ", A.택번호"
    Query = Query & ", A.색상"
    Query = Query & ", A.무늬"
    Query = Query & ", A.내용"
    Query = Query & ", A.금액"
    Query = Query & ", (ISNULL(A.금액,0) *(100-A.세탁마진))/100 AS 지사금액 "
    Query = Query & ", A.상표"
    Query = Query & ", A.환불사유"
    Query = Query & " FROM TB_입출고 AS A LEFT OUTER JOIN TB_고객정보 AS B ON A.고객코드 = B.고객코드"
    
    If optGubun(0).Value = True Then
        Query = Query & " WHERE (SUBSTRING(A.세탁환불일자,1,10) >= '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "'"
        Query = Query & "   AND  SUBSTRING(A.세탁환불일자,1,10) <= '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "')"
    Else
        Query = Query & " WHERE (SUBSTRING(A.반품환불일자,1,10) >= '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "'"
        Query = Query & "   AND  SUBSTRING(A.반품환불일자,1,10) <= '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "')"
        
    End If
    
    Query = Query & "   AND  판매취소 <> 'Y' "
    
    If txtFind.Text <> "" Then
        Select Case cboGubun.Text
            Case "성명":     Query = Query & " AND B.성명 LIKE '%" & txtFind.Text & "%'"
            Case "전화번호":
                            Query = Query & " AND ( B.휴대전화  LIKE '%" & txtFind.Text & "%'"
                            Query = Query & "  OR B.전화번호 LIKE '%" & txtFind.Text & "%')"
            
            Case "주소":     Query = Query & " AND B.주소 LIKE '%" & txtFind.Text & "%'"
        End Select
    End If
    
    Query = Query & " ORDER BY A.접수일자, B.성명, A.택번호 "
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            If optGubun(0).Value = True Then
                .Col = 1:  .Text = ADORs!세탁환불일자 & ""             '
            Else
                .Col = 1:  .Text = ADORs!반품환불일자 & ""             '
            End If
            
            .Col = 2:  .Text = ADORs!접수일자 & ""                                      '
            .Col = 3:  .Text = ADORs!환불사유 & ""                                      '
            .Col = 4:  .Text = ADORs!성명 & ""                                          '
            .Col = 5:  .Text = ADORs!휴대전화 & ""                                      '
            .Col = 6:  .Text = ADORs!전화번호 & ""                                      '
            .Col = 7:  .Text = ADORs!의류명 & ""                                        '
            
            If Len(Trim(ADORs!택번호)) <= 6 Then
                .Col = 8: .Text = ADORs!택번호 & ""                                     '
            Else
                .Col = 8: .Text = Format(ADORs!택번호, "000-00-0000")                   '
            End If
            
            .Col = 9:  .Text = ADORs!색상 & ""                                          '
            .Col = 10: .Text = ADORs!무늬 & ""                                          '
            .Col = 11: .Text = ADORs!내용 & ""                                          '
            .Col = 12: .Text = ADORs!금액 & ""                                          '
            .Col = 13: .Text = ADORs!지사금액 & ""                                      '
            .Col = 14: .Text = ADORs!상표 & ""                                          '
        
            ADORs.MoveNext
        Loop
        
        .ReDraw = True
    End With
    ADORs.Close
    Set ADORs = Nothing
    
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

    Open AppPath & "XML\환불현황.XML" For Output As #1
    
    Print #1, "<?xml version=""1.0"" encoding=""EUC-KR""?>"
    Print #1, "<root>"
    
          XML = "    <조건>"
    XML = XML & "        <검색조건>환불일자 : " & Format(dtpDay(0).Value, "YYYY-MM-DD") & " ~ " & Format(dtpDay(1).Value, "YYYY-MM-DD") & "</검색조건>"
    XML = XML & "        <가맹점>크린에이드 - " & Func_Replace(가맹점정보.가맹점명) & "</가맹점>"
    XML = XML & "   </조건>"
    Print #1, XML
    
    With sprGrid
        For I = 1 To .MaxRows
            .Row = I
            
                             XML = "    <Data>"
            .Col = 1:  XML = XML & "        <환불일자>" & .Text & "</환불일자>"
            .Col = 2:  XML = XML & "        <접수일자>" & .Text & "</접수일자>"
            .Col = 3:  XML = XML & "        <환불사유>" & Func_Replace(.Text) & "</환불사유>"
            .Col = 4:  XML = XML & "        <성명>" & Func_Replace(.Text) & "</성명>"
            .Col = 5:  XML = XML & "        <전화번호>" & Func_Replace(.Text) & "</전화번호>"
            .Col = 6:  XML = XML & "        <휴대전화>" & Func_Replace(.Text) & "</휴대전화>"
            .Col = 7:  XML = XML & "        <품명>" & Func_Replace(.Text) & "</품명>"
            .Col = 8:  XML = XML & "        <택번호>" & .Text & "</택번호>"
            .Col = 9:  XML = XML & "        <색상>" & Func_Replace(.Text) & "</색상>"
            .Col = 10: XML = XML & "        <무늬>" & Func_Replace(.Text) & "</무늬>"
            .Col = 11: XML = XML & "        <내용>" & Func_Replace(.Text) & "</내용>"
            .Col = 12: XML = XML & "        <금액>" & .Text & "</금액>"
            .Col = 13: XML = XML & "        <상태>" & Func_Replace(.Text) & "</상태>"
            .Col = 14: XML = XML & "        <상표>" & Func_Replace(.Text) & "</상표>"
                       XML = XML & "   </Data>"
                       Print #1, XML
        Next I
        
        Print #1, "</root>"
        Close #1
    End With
    
    If Print_PreView = True Then
        With rpt환불현황
            .dc.FileURL = AppPath & "XML\환불현황.XML"
            .Show 1
        End With
    Else
        With rpt환불현황
            .dc.FileURL = AppPath & "XML\환불현황.XML"
            .PrintReport False
        End With
    
        Unload rpt환불현황
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

Private Sub Form_Resize()
    On Error Resume Next
    
    cmdBtn(5).Left = Me.Width - cmdBtn(5).Width - 200
End Sub
