VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm입고예정 
   Caption         =   "입고 예정현황"
   ClientHeight    =   11805
   ClientLeft      =   3330
   ClientTop       =   3135
   ClientWidth     =   17985
   ControlBox      =   0   'False
   LinkTopic       =   "Form26"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11805
   ScaleWidth      =   17985
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   11805
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17985
      _ExtentX        =   31724
      _ExtentY        =   20823
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm입고예정.frx":0000
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   10575
         Left            =   15
         TabIndex        =   1
         Top             =   1215
         Width           =   17955
         _Version        =   524288
         _ExtentX        =   31671
         _ExtentY        =   18653
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowUserFormulas=   -1  'True
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
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
         MaxCols         =   13
         MaxRows         =   300
         MoveActiveOnFocus=   0   'False
         OperationMode   =   1
         Protect         =   0   'False
         SpreadDesigner  =   "frm입고예정.frx":0072
         VisibleCols     =   9
         VisibleRows     =   50
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   17955
         _ExtentX        =   31671
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
         Caption         =   "      입고 예정현황"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm입고예정.frx":0980
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm입고예정.frx":0BA6
            Top             =   -15
            Width           =   765
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   750
         Index           =   0
         Left            =   15
         TabIndex        =   3
         Top             =   450
         Width           =   17955
         _ExtentX        =   31671
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   330
            Left            =   915
            TabIndex        =   4
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   582
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
            Format          =   59506691
            CurrentDate     =   40279
         End
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   630
            Left            =   5415
            TabIndex        =   5
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 조회(&F)"
            Appearance      =   6
            Picture         =   "frm입고예정.frx":1770
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
            Picture         =   "frm입고예정.frx":1E6A
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
            Picture         =   "frm입고예정.frx":25E4
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
            Picture         =   "frm입고예정.frx":3676
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "예정일자:"
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
            TabIndex        =   9
            Top             =   120
            Width           =   840
         End
      End
   End
End
Attribute VB_Name = "frm입고예정"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Sub Data_Print(Print_PreView As Boolean)
    On Error GoTo ErrRtn
    
    If sprGrid.MaxRows = 0 Then Exit Sub

    If Dir(AppPath & "XML", vbDirectory) = "" Then
        MkDir AppPath & "XML"
    End If

    Open AppPath & "XML\입고예정.XML" For Output As #1
    
    Print #1, "<?xml version=""1.0"" encoding=""EUC-KR""?>"
    Print #1, "<root>"
    
          XML = "    <조건>"
    XML = XML & "        <예정일자>예정일자 : " & Format(dtpDay.Value, "YYYY-MM-DD") & "</예정일자>"
    XML = XML & "        <가맹점>크린에이드 - " & Func_Replace(가맹점정보.가맹점명) & "</가맹점>"
    XML = XML & "   </조건>"
    Print #1, XML
    
    With sprGrid
        For i = 1 To .MaxRows
            .Row = i
            
                             XML = "    <Data>"
            .Col = 1:  XML = XML & "        <접수일자>" & .Text & "</접수일자>"
            .Col = 2:  XML = XML & "        <지사출고일>" & .Text & "</지사출고일>"
            .Col = 3:  XML = XML & "        <가맹점입고일>" & .Text & "</가맹점입고일>"
            .Col = 4:  XML = XML & "        <성명>" & Func_Replace(.Text) & "</성명>"
            .Col = 5:  XML = XML & "        <전화번호>" & Func_Replace(.Text) & "</전화번호>"
            .Col = 6:  XML = XML & "        <휴대전화>" & Func_Replace(.Text) & "</휴대전화>"
            .Col = 7:  XML = XML & "        <품명>" & Func_Replace(.Text) & "</품명>"
            .Col = 8:  XML = XML & "        <택번호>" & .Text & "</택번호>"
            .Col = 9:  XML = XML & "        <색상>" & Func_Replace(.Text) & "</색상>"
            .Col = 10: XML = XML & "        <무늬>" & Func_Replace(.Text) & "</무늬>"
            .Col = 11: XML = XML & "        <내용>" & Func_Replace(.Text) & "</내용>"
            .Col = 12: XML = XML & "        <금액>" & .Text & "</금액>"
            .Col = 13: XML = XML & "        <상표>" & Func_Replace(.Text) & "</상표>"
                       XML = XML & "   </Data>"
                       Print #1, XML
        Next i
        
        Print #1, "</root>"
        Close #1
    End With
    
    If Print_PreView = True Then
        With rpt입고예정
            .dc.FileURL = AppPath & "XML\입고예정.XML"
            .Show 1
        End With
    Else
        With rpt입고예정
            .dc.FileURL = AppPath & "XML\입고예정.XML"
            .PrintReport False
        End With
        
        Unload rpt입고예정
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub cmdList_Click()
    Query = "SELECT    A.접수일자"
    Query = Query & ", SUBSTRING(A.지사출고일자,1,10)   AS 지사출고일자"
    Query = Query & ", SUBSTRING(A.가맹점입고일자,1,10) AS 가맹점입고일자"
    Query = Query & ", B.전화번호"
    Query = Query & ", B.휴대전화"
    Query = Query & ", B.성명"
    Query = Query & ", A.의류명"
    Query = Query & ", A.택번호"
    Query = Query & ", A.색상"
    Query = Query & ", A.무늬"
    Query = Query & ", A.내용"
    Query = Query & ", A.금액"
    Query = Query & ", A.결제여부"
    Query = Query & ", A.상표 "
    Query = Query & " FROM TB_입출고 AS A LEFT OUTER JOIN TB_고객정보 AS B ON A.고객코드 = B.고객코드"
    Query = Query & " WHERE (A.예정일자 = '" & Format(dtpDay.Value, "YYYY-MM-DD") & "') "
    Query = Query & "   AND (A.지사출고상태 <> '出' AND A.지사출고상태 <> '反') "
    Query = Query & " ORDER BY A.지사출고일자 ASC"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
     
    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = ADORs!접수일자 & ""                    '
            .Col = 2:  .Text = ADORs!지사출고일자 & ""                '
            .Col = 3:  .Text = ADORs!가맹점입고일자 & ""              '
            .Col = 4:  .Text = ADORs!성명 & ""                        '
            .Col = 5:  .Text = ADORs!전화번호 & ""                    '
            .Col = 6:  .Text = ADORs!휴대전화 & ""                    '
            .Col = 7:  .Text = ADORs!의류명 & ""                      '
            
            If Len(Trim(ADORs!택번호)) <= 6 Then
                .Col = 8: .Text = ADORs!택번호 & ""                   '
            Else
                .Col = 8: .Text = Format(ADORs!택번호, "000-00-0000") '
            End If
            
            .Col = 9:  .Text = ADORs!색상 & ""                        '
            .Col = 10: .Text = ADORs!무늬 & ""                        '
            .Col = 11: .Text = ADORs!내용 & ""                        '
            .Col = 12: .Text = ADORs!금액 & ""                        '
            .Col = 13: .Text = ADORs!상표 & ""                        '
            
            If Trim(ADORs!가맹점입고일자) <> "" Then
                .Row = .MaxRows: .Row2 = .MaxRows
                .Col = 1: .Col2 = .MaxCols
                .BlockMode = True
                .BackColor = vbYellow
                .BlockMode = False
            End If
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
                
        .ReDraw = True
    End With
End Sub

Private Sub dtpDay_Change()
    cmdList_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyChk (KeyCode)
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
    
    dtpDay.Value = Format(Date, "YYYY-MM-DD")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    cmdBtn(5).Left = Me.Width - cmdBtn(5).Width - 200
End Sub
