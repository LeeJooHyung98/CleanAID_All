VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm마일리지현황 
   Caption         =   "마일리지 현황"
   ClientHeight    =   10305
   ClientLeft      =   2775
   ClientTop       =   4005
   ClientWidth     =   16140
   ClipControls    =   0   'False
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
   ScaleHeight     =   10305
   ScaleWidth      =   16140
   WindowState     =   2  '최대화
   Begin Threed.SSPanel pnlProg 
      Height          =   1215
      Left            =   45
      TabIndex        =   11
      Top             =   1905
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
      Picture         =   "frm마일리지현황.frx":0000
      Caption         =   " "
      BorderWidth     =   0
      BevelOuter      =   1
      BevelInner      =   2
      Alignment       =   4
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   10305
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16140
      _ExtentX        =   28469
      _ExtentY        =   18177
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm마일리지현황.frx":2FCB
      Begin Threed.SSPanel SSPanel1 
         Height          =   750
         Left            =   15
         TabIndex        =   1
         Top             =   450
         Width           =   16110
         _ExtentX        =   28416
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
            TabIndex        =   8
            Top             =   45
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
            TabIndex        =   7
            Top             =   45
            Width           =   2370
         End
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   630
            Left            =   6345
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
            Picture         =   "frm마일리지현황.frx":307D
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   3
            Left            =   7890
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
            Picture         =   "frm마일리지현황.frx":3777
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   11175
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
            Picture         =   "frm마일리지현황.frx":3EF1
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   4
            Left            =   9435
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
            Picture         =   "frm마일리지현황.frx":4F83
         End
         Begin Threed.SSCheck chkMileage 
            Height          =   300
            Left            =   1305
            TabIndex        =   12
            Top             =   390
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
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
            Caption         =   "마일리지:"
         End
         Begin CSTextLibCtl.sidbEdit txtMoney 
            Height          =   315
            Index           =   0
            Left            =   2415
            TabIndex        =   13
            Top             =   390
            Visible         =   0   'False
            Width           =   1050
            _Version        =   262145
            _ExtentX        =   1852
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.01
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Insert          =   0   'False
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   14
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtMoney 
            Height          =   315
            Index           =   1
            Left            =   3735
            TabIndex        =   14
            Top             =   390
            Visible         =   0   'False
            Width           =   1050
            _Version        =   262145
            _ExtentX        =   1852
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.01
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Insert          =   0   'False
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   14
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
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
            Height          =   150
            Index           =   1
            Left            =   3540
            TabIndex        =   15
            Top             =   450
            Visible         =   0   'False
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
            TabIndex        =   9
            Top             =   105
            Width           =   840
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   16110
         _ExtentX        =   28416
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
         Caption         =   "      마일리지 현황"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm마일리지현황.frx":567D
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm마일리지현황.frx":58A3
            Top             =   -15
            Width           =   765
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   360
         Left            =   11130
         TabIndex        =   10
         Top             =   1215
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   635
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " 마일리지 내역"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm마일리지현황.frx":646D
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   9075
         Left            =   15
         TabIndex        =   16
         Top             =   1215
         Width           =   11100
         _Version        =   524288
         _ExtentX        =   19579
         _ExtentY        =   16007
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
         MaxCols         =   9
         MaxRows         =   300
         MoveActiveOnFocus=   0   'False
         OperationMode   =   1
         Protect         =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frm마일리지현황.frx":668F
         VisibleCols     =   9
         VisibleRows     =   50
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread sprList 
         Height          =   8700
         Left            =   11130
         TabIndex        =   17
         Top             =   1590
         Width           =   4995
         _Version        =   524288
         _ExtentX        =   8811
         _ExtentY        =   15346
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   6
         ScrollBars      =   2
         SpreadDesigner  =   "frm마일리지현황.frx":700F
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "frm마일리지현황"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdView_Click()
    Call Data_Display
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyChk (KeyCode)
    
    If KeyCode = 13 Then
        SendKeys "{Tab}"
        KeyCode = 0
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
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
        '.UserColAction = UserColActionSort
    End With
    
    With cboGubun
        .Clear
        .AddItem "성명"
        .AddItem "전화번호"
        .AddItem "고객코드"
        
        .ListIndex = 0
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn
    
    sprList.MaxRows = 0
    
    pnlProg.Visible = True
    DoEvents
    
    Query = "SELECT    전화번호"
    Query = Query & ", 휴대전화"
    Query = Query & ", 성명"
    Query = Query & ", 고객코드"
    Query = Query & ", 사용가능마일리지"
    Query = Query & ", 총접수금액"
    Query = Query & ", 누적마일리지"
    Query = Query & ", 미수금액"
    Query = Query & ", SUBSTRING(최종거래일자,1,10) AS 최종거래일자"
    Query = Query & " FROM TB_고객정보"
    Query = Query & " WHERE 고객코드 <> ''"
    
    If txtFind.Text <> "" Then
        Select Case cboGubun.Text
            Case "성명":     Query = Query & " AND 성명 LIKE '%" & txtFind.Text & "%'"
            Case "전화번호":
                            Query = Query & " AND ( 휴대전화  LIKE '%" & txtFind.Text & "%'"
                            Query = Query & "    OR 전화번호 LIKE '%" & txtFind.Text & "%')"
            
            Case "주소":     Query = Query & " AND 주소 LIKE '%" & txtFind.Text & "%'"
        End Select
    End If
    
    Query = Query & " ORDER BY 성명 ASC"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprGrid
        .MaxRows = 0
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        
            .Col = 1: .Text = ADORs!고객코드 & ""                      ' 1
            .Col = 2: .Text = ADORs!성명 & ""                          ' 2
            .Col = 3: .Text = ADORs!전화번호 & ""                      ' 3
            .Col = 4: .Text = ADORs!휴대전화 & ""                      ' 4
            .Col = 5: .Text = ADORs!총접수금액 & ""                    ' 5
            .Col = 6: .Text = ADORs!누적마일리지 & ""                  ' 6
            .Col = 7: .Text = ADORs!사용가능마일리지 & ""              ' 7
            .Col = 8: .Text = ADORs!미수금액 & ""                      ' 8
            .Col = 9: .Text = Format(ADORs!최종거래일자, "YYYY-MM-DD") ' 9
            
            ADORs.MoveNext
        Loop
    End With
    ADORs.Close
    Set ADORs = Nothing
    
    pnlProg.Visible = False
    
    Exit Sub
    
ErrRtn:
    pnlProg.Visible = False
    
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

    Open AppPath & "XML\마일리지현황.XML" For Output As #1
    
    Print #1, "<?xml version=""1.0"" encoding=""EUC-KR""?>"
    Print #1, "<root>"
    
          XML = "    <조건>"
    XML = XML & "        <미수금액>미수금액 : " & txtMoney(0).Text & " ~ " & txtMoney(0).Text & "</미수금액>"
    XML = XML & "        <가맹점>크린에이드 - " & Func_Replace(가맹점정보.가맹점명) & "</가맹점>"
    XML = XML & "   </조건>"
    XML = XML & "   <합계>"
    'XML = XML & "       <발생마일리지>발생마일리지 : " & txtMoney(2).Text & " 원</발생마일리지>"
    'XML = XML & "       <사용마일리지>사용마일리지 : " & txtMoney(3).Text & " 원</사용마일리지>"
    'XML = XML & "       <사용가능마일리지>사용가능마일리지 : " & txtMoney(4).Text & " 원</사용가능마일리지>"
    XML = XML & "   </합계>"
    Print #1, XML
    
    With sprGrid
        For i = 1 To .MaxRows
            .Row = i
            
                             XML = "    <Data>"
            .Col = 1:  XML = XML & "        <고객코드>" & Func_Replace(.Text) & "</고객코드>"
            .Col = 2:  XML = XML & "        <성명>" & Func_Replace(.Text) & "</성명>"
            .Col = 3:  XML = XML & "        <전화번호>" & Func_Replace(.Text) & "</전화번호>"
            .Col = 4:  XML = XML & "        <휴대전화>" & Func_Replace(.Text) & "</휴대전화>"
            .Col = 5:  XML = XML & "        <매출금액>" & .Text & "</매출금액>"
            .Col = 6:  XML = XML & "        <누계마일리지>" & .Text & "</누계마일리지>"
            .Col = 7:  XML = XML & "        <사용마일리지>" & .Text & "</사용마일리지>"
            .Col = 8:  XML = XML & "        <미수금액>" & .Text & "</미수금액>"
            .Col = 9:  XML = XML & "        <최종거래일자>" & .Text & "</최종거래일자>"
                       XML = XML & "   </Data>"
                       Print #1, XML
        Next i
        
        Print #1, "</root>"
        Close #1
    End With
    
    If Print_PreView = True Then
        With rpt마일리지현황
            .dc.FileURL = AppPath & "XML\마일리지현황.XML"
            .Show 1
        End With
    Else
        With rpt마일리지현황
            .dc.FileURL = AppPath & "XML\마일리지현황.XML"
            .PrintReport False
        End With
    
        Unload rpt마일리지현황
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
    
    Call 마일리지_Display(sprGrid.Text)
End Sub

Private Sub sprGrid_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Call sprGrid_Click(NewCol, NewRow)
End Sub

Private Sub 마일리지_Display(고객코드 As String)
    Dim 발생마일리지 As Long
    Dim 사용마일리지 As Long
    Dim 삭제마일리지 As Long
    Dim 반환마일리지 As Long
    
    Dim 사용가능마일리지 As Long
    
    On Error GoTo ErrRtn

    With sprList
        .MaxRows = 0
        .ReDraw = False
        
        Query = "SELECT * FROM TB_매출"
        Query = Query & " WHERE 고객코드 = '" & 고객코드 & "'"
        Query = Query & "   AND (발생마일리지 <> 0"
        Query = Query & "    OR 사용마일리지 <> 0"
        Query = Query & "    OR 삭제마일리지 <> 0)"
        Query = Query & " ORDER BY 매출일자 DESC, 매출시간 DESC"
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = Format(ADORs!매출일자, "YY-MM-DD") & "" '
            .Col = 2: .Text = ADORs!발생마일리지 & ""                 '
            .Col = 3: .Text = ADORs!사용마일리지 & ""                 '
            .Col = 4: .Text = ADORs!삭제마일리지 & ""                 '
            .Col = 5: .Text = ""                      'ADORs!반환마일리지 & ""
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
                
        .ReDraw = True
        
        If .MaxRows > 0 Then
            사용가능마일리지 = 0
            
            For i = .MaxRows To 1 Step -1
                .Row = i
                .Col = 2: 발생마일리지 = .Value
                .Col = 3: 사용마일리지 = .Value
                .Col = 4: 삭제마일리지 = .Value
                
                사용가능마일리지 = 사용가능마일리지 + 발생마일리지
                사용가능마일리지 = 사용가능마일리지 - 사용마일리지
                사용가능마일리지 = 사용가능마일리지 - 삭제마일리지
                
                .Col = 6: .Value = 사용가능마일리지
            Next i
                    
            .Row = 1
            .Col = 6: .FontBold = True: .RowHeight(-1) = 14
        End If
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub


