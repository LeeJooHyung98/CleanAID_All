VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm품목별집계현황 
   Caption         =   "품목별 집계현황"
   ClientHeight    =   10335
   ClientLeft      =   1035
   ClientTop       =   3360
   ClientWidth     =   15825
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form16"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10335
   ScaleWidth      =   15825
   ShowInTaskbar   =   0   'False
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   10335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15825
      _ExtentX        =   27914
      _ExtentY        =   18230
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm품목별집계현황.frx":0000
      Begin FPSpreadADO.fpSpread sprGrid 
         Bindings        =   "frm품목별집계현황.frx":0092
         Height          =   8625
         Left            =   15
         TabIndex        =   1
         Top             =   1215
         Width           =   15795
         _Version        =   524288
         _ExtentX        =   27861
         _ExtentY        =   15214
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
         MaxCols         =   13
         MaxRows         =   1000000
         OperationMode   =   1
         Protect         =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frm품목별집계현황.frx":00A6
         VisibleCols     =   9
         VisibleRows     =   200
         HighlightHeaders=   1
         HighlightStyle  =   1
         ScrollBarStyle  =   2
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   465
         Index           =   1
         Left            =   15
         TabIndex        =   2
         Top             =   9855
         Width           =   15795
         _ExtentX        =   27861
         _ExtentY        =   820
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin CSTextLibCtl.silgEdit txtSum 
            Height          =   375
            Index           =   0
            Left            =   960
            TabIndex        =   16
            Top             =   45
            Width           =   1080
            _Version        =   262145
            _ExtentX        =   1905
            _ExtentY        =   661
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.74
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "0"
            Text            =   " 0"
            StartText.x     =   2
            StartText.y     =   4
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
            Mask            =   ""
            Justification   =   2
            Undo            =   1
            Data            =   0
         End
         Begin CSTextLibCtl.silgEdit txtSum 
            Height          =   375
            Index           =   1
            Left            =   2745
            TabIndex        =   18
            Top             =   45
            Width           =   1365
            _Version        =   262145
            _ExtentX        =   2408
            _ExtentY        =   661
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.74
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "0"
            Text            =   " 0"
            StartText.x     =   2
            StartText.y     =   4
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
            Mask            =   ""
            Justification   =   2
            Undo            =   1
            Data            =   0
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  '투명
            Caption         =   "원"
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
            Index           =   7
            Left            =   4185
            TabIndex        =   19
            Top             =   150
            Width           =   180
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "합    계:"
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
            Index           =   6
            Left            =   45
            TabIndex        =   17
            Top             =   150
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  '투명
            Caption         =   "점"
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
            Index           =   19
            Left            =   2145
            TabIndex        =   3
            Top             =   150
            Width           =   180
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   750
         Index           =   0
         Left            =   15
         TabIndex        =   4
         Top             =   450
         Width           =   15795
         _ExtentX        =   27861
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox cboGroup 
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
            ItemData        =   "frm품목별집계현황.frx":09A5
            Left            =   915
            List            =   "frm품목별집계현황.frx":09A7
            Style           =   2  '드롭다운 목록
            TabIndex        =   6
            Top             =   420
            Width           =   3090
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   330
            Index           =   0
            Left            =   915
            TabIndex        =   7
            Top             =   60
            Width           =   1425
            _ExtentX        =   2514
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
            Format          =   58916867
            CurrentDate     =   40279
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   330
            Index           =   1
            Left            =   2580
            TabIndex        =   8
            Top             =   60
            Width           =   1425
            _ExtentX        =   2514
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
            Format          =   58916867
            CurrentDate     =   40279
         End
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   645
            Left            =   4500
            TabIndex        =   12
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1138
            _StockProps     =   79
            Caption         =   " 조회(&F)"
            Appearance      =   6
            Picture         =   "frm품목별집계현황.frx":09A9
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   645
            Index           =   3
            Left            =   6795
            TabIndex        =   13
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1138
            _StockProps     =   79
            Caption         =   " 엑셀(&E)"
            Appearance      =   6
            Picture         =   "frm품목별집계현황.frx":10A3
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   645
            Index           =   5
            Left            =   10800
            TabIndex        =   14
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1138
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frm품목별집계현황.frx":181D
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   645
            Index           =   4
            Left            =   8340
            TabIndex        =   15
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1138
            _StockProps     =   79
            Caption         =   " 출력(&P)"
            Appearance      =   6
            Picture         =   "frm품목별집계현황.frx":28AF
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "대 분 류:"
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
            Index           =   5
            Left            =   45
            TabIndex        =   11
            Top             =   480
            Width           =   840
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
            Index           =   4
            Left            =   45
            TabIndex        =   10
            Top             =   120
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "~"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   2385
            TabIndex        =   9
            Top             =   90
            Width           =   135
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   5
         Top             =   15
         Width           =   15795
         _ExtentX        =   27861
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
         Caption         =   "      품목별 집계현황"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm품목별집계현황.frx":2FA9
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm품목별집계현황.frx":31CF
            Top             =   -15
            Width           =   765
         End
      End
   End
End
Attribute VB_Name = "frm품목별집계현황"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboGroup_Click()
    Call Data_Display
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 3: Call Export_Excel(frmMain.cdgExcel, sprGrid)
        
        Case 4
            Rtn = MsgBox("출력 미리보기를 하시겠습니까?", vbQuestion + vbYesNo, "출력")
            
            If Rtn = vbYes Then
                Call Data_Print(True)
            Else
                Call Data_Print(False)
            End If
            
        Case 5
            Unload Me
    End Select
End Sub

Private Sub Data_Print(Print_PreView As Boolean)
    On Error GoTo ErrRtn
    
    If sprGrid.MaxRows = 0 Then Exit Sub

    If Dir(AppPath & "XML", vbDirectory) = "" Then
        MkDir AppPath & "XML"
    End If

    Open AppPath & "XML\품목별집계.XML" For Output As #1
    
    Print #1, "<?xml version=""1.0"" encoding=""EUC-KR""?>"
    Print #1, "<root>"
    
          XML = "    <조건>"
    XML = XML & "        <접수일자>접수일자 : " & Format(dtpDay(0).Value, "YYYY-MM-DD") & " ~ " & Format(dtpDay(1).Value, "YYYY-MM-DD") & "</접수일자>"
    XML = XML & "        <품목분류>품목분류 - " & Func_Replace(cboGroup.Text) & "</품목분류>"
    XML = XML & "   </조건>"
    Print #1, XML
    
    With sprGrid
        For i = 1 To .MaxRows
            .Row = i
            
                             XML = "    <Data>"
            .Col = 1:  XML = XML & "        <접수일자>" & .Text & "</접수일자>"
            .Col = 2:  XML = XML & "        <성명>" & Func_Replace(.Text) & "</성명>"
            .Col = 3:  XML = XML & "        <전화번호>" & Func_Replace(.Text) & "</전화번호>"
            .Col = 4:  XML = XML & "        <휴대전화>" & Func_Replace(.Text) & "</휴대전화>"
            .Col = 5:  XML = XML & "        <의류코드>" & Func_Replace(.Text) & "</의류코드>"
            .Col = 6:  XML = XML & "        <품명>" & Func_Replace(.Text) & "</품명>"
            .Col = 7:  XML = XML & "        <택번호>" & Func_Replace(.Text) & "</택번호>"
            .Col = 8:  XML = XML & "        <색상>" & Func_Replace(.Text) & "</색상>"
            .Col = 9:  XML = XML & "        <무늬>" & Func_Replace(.Text) & "</무늬>"
            .Col = 10: XML = XML & "        <내용>" & Func_Replace(.Text) & "</내용>"
            .Col = 11: XML = XML & "        <금액>" & Func_Replace(.Text) & "</금액>"
            .Col = 12: XML = XML & "        <결제>" & Func_Replace(.Text) & "</결제>"
            .Col = 13: XML = XML & "        <상표>" & Func_Replace(.Text) & "</상표>"
                       XML = XML & "   </Data>"
                       Print #1, XML
        Next i
        
        Print #1, "</root>"
        Close #1
    End With
    
    If Print_PreView = True Then
        With rpt품목별집계
            .dc.FileURL = AppPath & "XML\품목별집계.XML"
            .Show 1
        End With
    Else
        With rpt품목별집계
            .dc.FileURL = AppPath & "XML\품목별집계.XML"
            .PrintReport False
        End With
        
        Unload rpt품목별집계
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
    
    dtpDay(0).Value = Date
    dtpDay(1).Value = Date
    
    '-----------------------------------------------------------
    '
    '-----------------------------------------------------------
    Query = "SELECT * FROM TB_의류분류"
    Query = Query & " ORDER BY 의류분류코드 ASC"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    With cboGroup
        .Clear
        
        Do Until ADORs.EOF
            .AddItem ADORs!의류분류코드 & ":" & ADORs!의류분류명
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        If .ListCount > 1 Then .ListIndex = 0
    End With
End Sub


Private Sub Data_Display()
    Dim dblAmt      As Double
    Dim dblACodeAmt As Double
    
    txtSum(0).Value = 0
    txtSum(1).Value = 0
    
    Query = "SELECT DISTINCT A.의류코드"
    Query = Query & ", A.접수일자"
    Query = Query & ", B.휴대전화"
    Query = Query & ", B.전화번호"
    Query = Query & ", B.성명"
    Query = Query & ", A.의류명"
    Query = Query & ", A.택번호"
    Query = Query & ", A.색상"
    Query = Query & ", A.무늬"
    Query = Query & ", A.내용"
    Query = Query & ", ISNULL(A.금액,0) AS 금액"
    Query = Query & ", A.결제여부"
    Query = Query & ", A.상표"
    Query = Query & " FROM TB_입출고 AS A LEFT OUTER JOIN TB_고객정보 AS B ON A.고객코드 = B.고객코드"
    Query = Query & " WHERE (A.접수일자 >= '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "'"
    Query = Query & "   AND  A.접수일자 <= '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "')"
    Query = Query & "   AND (A.판매취소 <> 'Y') "
    Query = Query & "   AND (SUBSTRING(A.의류코드,1,2) = '" & Left(cboGroup.Text, 2) & "') "
    Query = Query & " ORDER BY A.접수일자, A.택번호 ASC"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = Format(ADORs!접수일자, "YYYY-MM-DD")
            .Col = 2:  .Text = ADORs!성명 & ""
            .Col = 3:  .Text = ADORs!전화번호 & ""
            .Col = 4:  .Text = ADORs!휴대전화 & ""
            .Col = 5:  .Text = ADORs!의류코드 & ""
            .Col = 6:  .Text = ADORs!의류명 & ""
            
            If Len(Trim(ADORs!택번호)) <= 6 Then
                .Col = 7: .Text = ADORs!택번호 & ""
            Else
                .Col = 7: .Text = Format(ADORs!택번호, "000-00-0000")
            End If
            
            .Col = 8:  .Text = ADORs!색상 & ""
            .Col = 9:  .Text = ADORs!무늬 & ""
            .Col = 10: .Text = ADORs!내용 & ""
            .Col = 11: .Text = ADORs!금액 & ""
            .Col = 12: .Text = ADORs!결제여부 & ""
            .Col = 13: .Text = ADORs!상표 & ""
        
            txtSum(1).Value = txtSum(1).Value + ADORs!금액
            
            'If UCase(Left(ADORs!의류코드, 1)) = "A" Then
            '    txtSum(2).Value = txtSum(2).Value + ADORs!금액
            'End If
        
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        .ReDraw = True
        
        txtSum(0).Value = .MaxRows
    End With
    
    'lblBSu(0).Caption = Format(i, "#,##0")
    'lblBSu(1).Caption = Format(dblAmt, "#,##0")
    
    'lblBSu(2).Caption = Format(dblACodeAmt * (1 - (Val(가맹점정보.외주마진) / 100)), "#,##0")
    'lblBSu(3).Caption = Format(dblACodeAmt * (Val(가맹점정보.외주마진) / 100), "#,##0")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    cmdBtn(5).Left = (Me.Width - 200) - cmdBtn(5).Width
End Sub
