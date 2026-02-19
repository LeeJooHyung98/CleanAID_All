VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm반품현황 
   Caption         =   "지사 반품 현황"
   ClientHeight    =   9180
   ClientLeft      =   6315
   ClientTop       =   2640
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
      PaneTree        =   "frm반품현황.frx":0000
      Begin FPSpreadADO.fpSpread sprGrid 
         Bindings        =   "frm반품현황.frx":0072
         Height          =   7860
         Left            =   15
         TabIndex        =   1
         Top             =   1305
         Width           =   14265
         _Version        =   524288
         _ExtentX        =   25162
         _ExtentY        =   13864
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
         MaxCols         =   15
         MaxRows         =   1000000
         OperationMode   =   1
         Protect         =   0   'False
         SpreadDesigner  =   "frm반품현황.frx":0086
         VisibleCols     =   9
         VisibleRows     =   200
         HighlightHeaders=   1
         HighlightStyle  =   1
         ScrollBarStyle  =   2
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   840
         Left            =   15
         TabIndex        =   2
         Top             =   450
         Width           =   14265
         _ExtentX        =   25162
         _ExtentY        =   1482
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
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
            Picture         =   "frm반품현황.frx":0A23
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
            Picture         =   "frm반품현황.frx":111D
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
            Picture         =   "frm반품현황.frx":1897
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
            Picture         =   "frm반품현황.frx":2929
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   0
            Left            =   915
            TabIndex        =   8
            Top             =   75
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
            Format          =   64225283
            CurrentDate     =   40279
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   1
            Left            =   2655
            TabIndex        =   9
            Top             =   75
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
            Format          =   64225283
            CurrentDate     =   40279
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
            TabIndex        =   11
            Top             =   135
            Width           =   120
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "반품일자:"
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
            TabIndex        =   10
            Top             =   135
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
         Caption         =   "      지사 반품 현황"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm반품현황.frx":3023
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm반품현황.frx":3249
            Top             =   -15
            Width           =   765
         End
      End
   End
End
Attribute VB_Name = "frm반품현황"
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
    

    
    dtpDay(0).Value = Format(Date, "YYYY-MM-DD")
    dtpDay(1).Value = Format(Date, "YYYY-MM-DD")
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn
    
    
 

    Query = "SELECT    rt.반품일자, rt.차수, IT.접수일자,rt.반품사유, mt.성명, mt.휴대전화, mt.전화번호, IT.의류명, "
    Query = Query & " rt.택번호, IT.색상,IT.무늬, IT.내용, IT.금액,"
    Query = Query & " (ISNULL(it.금액,0) *(100-it.세탁마진))/100 AS 지사금액, IT.상표"
    Query = Query & " from TB_반품현황 RT"
    Query = Query & " left outer join TB_입출고 IT"
    Query = Query & "     on IT.접수일자 = rt.접수일자"
    Query = Query & "     and IT.가맹점코드 =rt.가맹점코드"
    Query = Query & "     and IT.택번호 = rt.택번호"
    Query = Query & " left outer join TB_고객정보 MT"
    Query = Query & "   on IT.가맹점코드 =mt.가맹점코드"
    Query = Query & "   and IT.고객코드 = MT.고객코드"
    Query = Query & " where 반품일자 between '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "' and '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "' "
    
    
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = ADORs!반품일자 & ""             '
            .Col = 2:  .Text = ADORs!차수 & ""                                      '
            .Col = 3:  .Text = ADORs!접수일자 & ""                                      '
            .Col = 4:  .Text = ADORs!반품사유 & ""                                      '
            .Col = 5:  .Text = ADORs!성명 & ""                                          '
            .Col = 6:  .Text = ADORs!휴대전화 & ""                                      '
            .Col = 7:  .Text = ADORs!전화번호 & ""                                      '
            .Col = 8:  .Text = ADORs!의류명 & ""                                        '
            
            If Len(Trim(ADORs!택번호)) <= 6 Then
                .Col = 9: .Text = ADORs!택번호 & ""                                     '
            Else
                .Col = 9: .Text = Format(ADORs!택번호, "000-00-0000")                   '
            End If
            
            .Col = 10:  .Text = ADORs!색상 & ""                                          '
            .Col = 11: .Text = ADORs!무늬 & ""                                          '
            .Col = 12: .Text = ADORs!내용 & ""                                          '
            .Col = 13: .Text = ADORs!금액 & ""                                          '
            .Col = 14: .Text = ADORs!지사금액 & ""                                      '
            .Col = 15: .Text = ADORs!상표 & ""                                          '
        
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
        For i = 1 To .MaxRows
            .Row = i
            
                             XML = "    <Data>"
            .Col = 1:  XML = XML & "        <환불일자>" & .Text & "</환불일자>"
            .Col = 3:  XML = XML & "        <접수일자>" & .Text & "</접수일자>"
            .Col = 4:  XML = XML & "        <환불사유>" & Func_Replace(.Text) & "</환불사유>"
            .Col = 5:  XML = XML & "        <성명>" & Func_Replace(.Text) & "</성명>"
            .Col = 6:  XML = XML & "        <전화번호>" & Func_Replace(.Text) & "</전화번호>"
            .Col = 7:  XML = XML & "        <휴대전화>" & Func_Replace(.Text) & "</휴대전화>"
            .Col = 8:  XML = XML & "        <품명>" & Func_Replace(.Text) & "</품명>"
            .Col = 9:  XML = XML & "        <택번호>" & .Text & "</택번호>"
            .Col = 10:  XML = XML & "        <색상>" & Func_Replace(.Text) & "</색상>"
            .Col = 11: XML = XML & "        <무늬>" & Func_Replace(.Text) & "</무늬>"
            .Col = 12: XML = XML & "        <내용>" & Func_Replace(.Text) & "</내용>"
            .Col = 13: XML = XML & "        <금액>" & .Text & "</금액>"
            .Col = 14: XML = XML & "        <상태>" & Func_Replace(.Text) & "</상태>"
            .Col = 15: XML = XML & "        <상표>" & Func_Replace(.Text) & "</상표>"
                       XML = XML & "   </Data>"
                       Print #1, XML
        Next i
        
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
    Call DataDownload
    Call Data_Display
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    cmdBtn(5).Left = Me.Width - cmdBtn(5).Width - 200
End Sub


Private Sub DataDownload()
    
    
    
    If Server_Connection(HostCon, "LAUNDRY" & 가맹점정보.지사코드) = False Then
        MsgBox "DB 서버와 접속이 되지 않아 본사출고 자료를 조회할수 없습니다.", vbCritical, "확인"
        
        Exit Sub
    End If
    
    
    Query = "SELECT   * "
    Query = Query & " FROM TB_반품현황 "
    Query = Query & " WHERE ISNULL(수신일자,'') = '' "
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, HostCon, adOpenDynamic, adLockOptimistic
        
    i = 0
    
    Do Until ADORs.EOF
        i = i + 1
        
        '--------------------------------------------------------------
        ' TB_의류 - 의류금액 정보 다운로드
        '--------------------------------------------------------------
        Query = "SELECT * FROM TB_반품현황"
        Query = Query & " WHERE 가맹점코드 = '" & ADORs!가맹점코드 & "'"
        Query = Query & " AND 반품일자 = '" & ADORs!반품일자 & "'"
        Query = Query & " AND 차수 = '" & ADORs!차수 & "'"
        Query = Query & " AND 택번호 = '" & ADORs!택번호 & "'"
        Set SUBRs = New ADODB.Recordset
        SUBRs.Open Query, ADOCon, adOpenDynamic, adLockOptimistic
        
        If SUBRs.EOF Then
            SUBRs.AddNew
            SUBRs!가맹점코드 = ADORs!가맹점코드 & ""
            SUBRs!반품일자 = ADORs!반품일자 & ""
            SUBRs!차수 = ADORs!차수 & ""
            SUBRs!택번호 = ADORs!택번호 & ""
            
        End If
        
        SUBRs!접수일자 = ADORs!접수일자 & ""
        SUBRs!반품사유 = Trim(ADORs!반품사유) & ""
        
        SUBRs.Update
        
        SUBRs.Close
        Set SUBRs = Nothing
        
        
        ADORs!수신일자 = Format(Now, "yyyy-MM-dd hh:mm:ss")
        ADORs.Update
        
        
        ADORs.MoveNext
    Loop
    ADORs.Close:    Set ADORs = Nothing
 
    
End Sub

