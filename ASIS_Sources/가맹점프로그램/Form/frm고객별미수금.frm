VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm고객별미수금 
   Caption         =   "고객별 미수금"
   ClientHeight    =   8850
   ClientLeft      =   2895
   ClientTop       =   4665
   ClientWidth     =   14505
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
   ScaleWidth      =   14505
   WindowState     =   2  '최대화
   Begin Threed.SSPanel pnlProg 
      Height          =   1215
      Left            =   45
      TabIndex        =   18
      Top             =   1620
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
      Picture         =   "frm고객별미수금.frx":0000
      Caption         =   " "
      BorderWidth     =   0
      BevelOuter      =   1
      BevelInner      =   2
      Alignment       =   4
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8850
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   15610
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm고객별미수금.frx":2FCB
      Begin Threed.SSPanel SSPanel1 
         Height          =   750
         Left            =   15
         TabIndex        =   1
         Top             =   450
         Width           =   15360
         _ExtentX        =   27093
         _ExtentY        =   1323
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
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
         Begin Threed.SSCheck chkMisu 
            Height          =   300
            Left            =   1065
            TabIndex        =   16
            Top             =   390
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
            Caption         =   "미수금액:"
         End
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
            TabIndex        =   8
            Top             =   45
            Width           =   1230
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
            Left            =   2175
            TabIndex        =   7
            Top             =   45
            Width           =   2370
         End
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   630
            Left            =   8430
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
            Picture         =   "frm고객별미수금.frx":307D
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   3
            Left            =   10515
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
            Picture         =   "frm고객별미수금.frx":3777
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   13800
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
            Picture         =   "frm고객별미수금.frx":3EF1
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   4
            Left            =   12060
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
            Picture         =   "frm고객별미수금.frx":4F83
         End
         Begin CSTextLibCtl.sidbEdit txtMoney 
            Height          =   315
            Index           =   0
            Left            =   2175
            TabIndex        =   13
            Top             =   390
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
            Left            =   3495
            TabIndex        =   14
            Top             =   390
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
            Left            =   3300
            TabIndex        =   15
            Top             =   450
            Width           =   120
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
         Width           =   15360
         _ExtentX        =   27093
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
         Caption         =   "      고객별 미수금"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm고객별미수금.frx":567D
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm고객별미수금.frx":58A3
            Top             =   -15
            Width           =   765
         End
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   7620
         Left            =   15
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1215
         Width           =   7425
         _Version        =   524288
         _ExtentX        =   13097
         _ExtentY        =   13441
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
         OperationMode   =   1
         Protect         =   0   'False
         RestrictCols    =   -1  'True
         RestrictRows    =   -1  'True
         ScrollBars      =   2
         ShadowText      =   0
         SpreadDesigner  =   "frm고객별미수금.frx":646D
         VisibleCols     =   5
         VisibleRows     =   500
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   360
         Left            =   7455
         TabIndex        =   11
         Top             =   1215
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   635
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " 이용내역"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm고객별미수금.frx":6BD4
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "('초기미수금' 이전 미수금액은 실제 미수금과 다를 수 있습니다)"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   0
            Left            =   1815
            TabIndex        =   17
            Top             =   105
            Width           =   6030
         End
      End
      Begin FPSpreadADO.fpSpread sprList 
         Height          =   7245
         Left            =   7455
         TabIndex        =   12
         Top             =   1590
         Width           =   7920
         _Version        =   524288
         _ExtentX        =   13970
         _ExtentY        =   12779
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
         MaxCols         =   10
         ScrollBars      =   2
         SpreadDesigner  =   "frm고객별미수금.frx":6F16
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "frm고객별미수금"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkMisu_Click(Value As Integer)
    If Value = 0 Then
        txtMoney(0).Value = 0
        txtMoney(1).Value = 0
    End If
End Sub

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
    
        .Row = 0
        .Col = 0: .Text = "순위"
    End With
    
    With sprList
        .MaxRows = 0
        .RowHeight(-1) = 14
        
        .Col = 10: .ColHidden = True
        
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
        .AddItem "주소"
        
        .ListIndex = 0
    End With
    
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn
    
    pnlProg.Visible = True
    DoEvents
    
    sprList.MaxRows = 0
    
    '---------------------------------------------------------------------
    '
    '---------------------------------------------------------------------
    Query = "SELECT     고객코드"
    Query = Query & " , 성명"
    Query = Query & " , 미수금액"
    Query = Query & " , 휴대전화"
    Query = Query & " , 전화번호"
    Query = Query & " , 주소"
    Query = Query & " , 총접수금액"
    Query = Query & " FROM TB_고객정보"
    Query = Query & " WHERE 고객코드 IS NOT NULL"
    
    If chkMisu.Value = -1 Then
        Query = Query & "   AND 미수금액 BETWEEN " & txtMoney(0).Value & " AND " & txtMoney(1).Value
    End If
    
    If txtData(0).Text <> "" Then
        Select Case cboGubun.Text
            Case "성명":     Query = Query & " AND 성명 LIKE '%" & txtData(0).Text & "%'"
            Case "전화번호":
                            Query = Query & " AND ( 휴대전화  LIKE '%" & txtData(0).Text & "%'"
                            Query = Query & "    OR 전화번호 LIKE '%" & txtData(0).Text & "%')"
            
            Case "주소":     Query = Query & " AND 주소 LIKE '%" & txtData(0).Text & "%'"
        End Select
    End If
    
    Query = Query & " ORDER BY 미수금액 DESC "
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        Do While Not ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        
            .Col = 1: .Text = ADORs!고객코드 & ""   ' 1
            .Col = 2: .Text = ADORs!성명 & ""       ' 2
            .Col = 3: .Text = ADORs!전화번호 & ""   ' 3
            .Col = 4: .Text = ADORs!휴대전화 & ""   ' 4
            .Col = 5: .Text = ADORs!총접수금액 & "" ' 5
            .Col = 6: .Text = ADORs!미수금액 & ""   ' 6
        
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
    
        .ReDraw = True
    End With
    
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

    Open AppPath & "XML\고객별미수금.XML" For Output As #1
    
    Print #1, "<?xml version=""1.0"" encoding=""EUC-KR""?>"
    Print #1, "<root>"
    
          XML = "    <조건>"
    XML = XML & "        <미수금액>미수금액 : " & txtMoney(0).Text & " ~ " & txtMoney(0).Text & "</미수금액>"
    XML = XML & "        <가맹점>크린에이드 - " & Func_Replace(가맹점정보.가맹점명) & "</가맹점>"
    XML = XML & "   </조건>"
    Print #1, XML
    
    With sprGrid
        For i = 1 To .MaxRows
            .Row = i
            
                             XML = "    <Data>"
                       XML = XML & "        <순위>" & i & "</순위>"
            .Col = 1:  XML = XML & "        <고객코드>" & .Text & "</고객코드>"
            .Col = 2:  XML = XML & "        <성명>" & Func_Replace(.Text) & "</성명>"
            .Col = 3:  XML = XML & "        <전화번호>" & Func_Replace(.Text) & "</전화번호>"
            .Col = 4:  XML = XML & "        <휴대전화>" & Func_Replace(.Text) & "</휴대전화>"
            .Col = 5:  XML = XML & "        <매출액>" & Func_Replace(.Text) & "</매출액>"
            .Col = 6:  XML = XML & "        <미수금>" & .Text & "</미수금>"
                       XML = XML & "   </Data>"
                       Print #1, XML
        Next i
        
        Print #1, "</root>"
        Close #1
    End With
    
    If Print_PreView = True Then
        With rpt고객별미수금
            .dc.FileURL = AppPath & "XML\고객별미수금.XML"
            .Show 1
        End With
    Else
        With rpt고객별미수금
            .dc.FileURL = AppPath & "XML\고객별미수금.XML"
            .PrintReport False
        End With
        
        Unload rpt고객별미수금
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
    Dim 고객미수금 As Long
    
    If Row <= 0 Then Exit Sub
    
    sprGrid.Row = Row
    sprGrid.Col = 6: 고객미수금 = sprGrid.Value
    
    sprGrid.Col = 1
    
    Call 미수금_Display(sprGrid.Text, 고객미수금)
End Sub

Private Sub sprGrid_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Call sprGrid_Click(NewCol, NewRow)
End Sub

Private Sub 미수금_Display(고객코드 As String, 고객미수금 As Long)
    Dim 초기미수금 As Long
    Dim 이전고객   As String
    Dim 초기시작일  As String
    
    Dim bMisu      As Boolean
    Dim 이전미수   As Long
    
    Dim 미수금     As Long
    
    On Error GoTo ErrRtn
    
    Query = "SELECT    ISNULL(초기미수금,0)"
    Query = Query & ", ISNULL(이전고객, '')"
    Query = Query & " FROM TB_고객정보"
    Query = Query & " WHERE 고객코드 = '" & 고객코드 & "'"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    If ADORs.EOF Then
        초기미수금 = 0
        이전고객 = ""
    Else
        초기미수금 = ADORs(0) & ""
        이전고객 = ADORs(1) & ""
    End If
    ADORs.Close
    Set ADORs = Nothing
    
    '---------------------------------------------------
    ' 초기 미수금 적용일자를 구한다.
    '---------------------------------------------------
    Query = "SELECT TOP 1   수정일자"
    Query = Query & " FROM TB_미수금수정"
    Query = Query & " WHERE 고객코드 = '" & 고객코드 & "'"
    Query = Query & "   AND 내용 = '초기 미수금'"
    
    Query = Query & " ORDER BY 수정일자 DESC"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF Then
        초기시작일 = "1900-01-01"
    Else
        초기시작일 = ADORs!수정일자 & ""
    End If
    ADORs.Close
    Set ADORs = Nothing
    
    bMisu = False

    With sprList
        .MaxRows = 0
        .ReDraw = False
        
        '----------------------------------------------------------
        ' TB_매출
        '----------------------------------------------------------
        Query = "SELECT    매출일자"
        Query = Query & ", 매출시간"
        Query = Query & ", 적요"
        Query = Query & ", 접수금액"
        Query = Query & ", 현금입금"
        Query = Query & ", 카드입금"
        Query = Query & ", 사용마일리지"
        Query = Query & ", 쿠폰입금"
        Query = Query & " FROM TB_매출"
        Query = Query & " WHERE 고객코드 = '" & 고객코드 & "'"
        Query = Query & "   AND 매출일자 >= '" & 초기시작일 & "' "
        Query = Query & " ORDER BY 매출일자 DESC, 매출시간 DESC"
        Set ADORs = New ADODB.RecordSet
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = Format(ADORs!매출일자, "YY-MM-DD") & "" '
            .Col = 2: .Text = ADORs!접수금액 & "" '
            .Col = 3: .Text = ADORs!현금입금 & "" '
            .Col = 4: .Text = ADORs!카드입금 & "" '
            .Col = 5: .Text = ADORs!사용마일리지 & "" '
            .Col = 6: .Text = ADORs!쿠폰입금 & "" '
            .Col = 7: .Text = ADORs!접수금액 - ADORs!현금입금 - ADORs!카드입금 - ADORs!사용마일리지 - ADORs!쿠폰입금 & "" '
                
            .Col = 10: .Text = ADORs!매출시간 & ""
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        '---------------------------------------------------
        ' TB_미수금수정
        '---------------------------------------------------
        Query = "SELECT    수정일자"
        Query = Query & ", 수정미수금, 내용 "
        Query = Query & " FROM TB_미수금수정"
        Query = Query & " WHERE 고객코드 = '" & 고객코드 & "'"
        Query = Query & " ORDER BY 수정일자 DESC"
        Set ADORs = New ADODB.RecordSet
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = Format(Left(ADORs!수정일자, 10), "YY-MM-DD") & ""
            .Col = 2:  .Value = 0
            .Col = 3:  .Value = 0
            .Col = 4:  .Value = 0
            .Col = 5:  .Value = 0
            .Col = 6:  .Value = 0
            .Col = 7:  .Value = ADORs!수정미수금 & "": .FontBold = True: .ForeColor = vbRed: .RowHeight(-1) = 14
            .Col = 8:  .Value = 0
            .Col = 9:  .Text = ADORs!내용 & "": .FontBold = True: .ForeColor = vbRed: .RowHeight(-1) = 14
            .Col = 10: .Text = Right(ADORs!수정일자, 8) & ""
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        .SortKey(1) = 1
        .SortKeyOrder(1) = SortKeyOrderDescending
        
        .SortKey(2) = 10
        .SortKeyOrder(2) = SortKeyOrderDescending
        
        .Sort -1, -1, -1, -1, SortByRow

        
        If .MaxRows > 0 Then
            '---------------------------------------------------
            ' 미수금액 계산
            '---------------------------------------------------
            For i = .MaxRows To 1 Step -1
                .Row = i
                .Col = 9
                If .Text = "초기 미수금" Then
                    '초기미수금 제외
                ElseIf .Text = "조정 - 고객수정" Then
                    '초기미수금 제외
                    .Col = 7: 이전미수 = .Value
                    .Col = 8: .Value = 이전미수 & ""
                    
                Else
                    If i = .MaxRows Then
                        이전미수 = 0
                    Else
                        .Row = i + 1
                        .Col = 8: 이전미수 = .Value
                    End If
                    
                    .Row = i
                    .Col = 7: 이전미수 = 이전미수 + .Value
                    .Col = 8: .Value = 이전미수 & ""
                End If
            Next i
                    
            .Row = 1
            .Col = 8: 미수금 = .Value: .FontBold = True: .RowHeight(-1) = 14
        End If
        
        
        Dim Misu_Row As Long
        
        For i = 1 To .MaxRows
            .Row = i
            
            .Col = 1
            If .Text = "초기 미수금" Then
                Misu_Row = i
            End If
            
            If (Misu_Row > 0) And (i > Misu_Row) Then
                .Col = 8: .Value = 0
            End If
        Next i
        
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

'Private Sub 미수금_Display(고객코드 As String, 고객미수금 As Long)
'    Dim 초기미수금 As Long
'    Dim 이전고객   As String
'
'    Dim bMisu      As Boolean
'    Dim 이전미수   As Long
'
'    Dim 미수금     As Long
'
'    On Error GoTo ErrRtn
'
'    Query = "SELECT    ISNULL(초기미수금,0)"
'    Query = Query & ", ISNULL(이전고객, '')"
'    Query = Query & " FROM TB_고객정보"
'    Query = Query & " WHERE 고객코드 = '" & 고객코드 & "'"
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    If ADORs.EOF Then
'        초기미수금 = 0
'        이전고객 = ""
'    Else
'        초기미수금 = ADORs(0) & ""
'        이전고객 = ADORs(1) & ""
'    End If
'    ADORs.Close
'    Set ADORs = Nothing
'
'    bMisu = False
'
'    Query = "SELECT    매출일자"
'    Query = Query & ", 매출시간"
'    Query = Query & ", 적요"
'    Query = Query & ", 접수금액"
'    Query = Query & ", 현금입금"
'    Query = Query & ", 카드입금"
'    Query = Query & ", 사용마일리지"
'    Query = Query & ", 쿠폰입금"
'    Query = Query & " FROM TB_매출"
'    Query = Query & " WHERE 고객코드 = '" & 고객코드 & "'"
'    Query = Query & " ORDER BY 매출일자 DESC, 매출시간 DESC"
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    With sprList
'        .MaxRows = 0
'        .ReDraw = False
'
'        Do Until ADORs.EOF
'            If (bMisu = False And 이전고객 = "Y") And (ADORs!매출시간 = "00:00:00" Or ADORs!적요 = "[미수금액 입금]") Then
'                bMisu = True
'
'                .MaxRows = .MaxRows + 1
'                .Row = .MaxRows
'
'                .Col = 1: .Text = "초기미수"
'                .Col = 2: .Text = 0 & "" '
'                .Col = 3: .Text = 0 & "" '
'                .Col = 4: .Text = 0 & "" '
'                .Col = 5: .Text = 0 & "" '
'                .Col = 6: .Text = 0 & "" '
'                .Col = 7: .Text = 0 & "" '
'                .Col = 8: .Text = 초기미수금 & "" '
'
'                .Row = .Row
'                .Row2 = .Row
'                .Col = 1
'                .Col2 = .MaxCols
'                .BlockMode = True
'                .BackColor = vbYellow
'                .BlockMode = False
'            End If
'
'            .MaxRows = .MaxRows + 1
'            .Row = .MaxRows
'
'            .Col = 1: .Text = Format(ADORs!매출일자, "YY-MM-DD") & "" '
'            .Col = 2: .Text = ADORs!접수금액 & "" '
'            .Col = 3: .Text = ADORs!현금입금 & "" '
'            .Col = 4: .Text = ADORs!카드입금 & "" '
'            .Col = 5: .Text = ADORs!사용마일리지 & "" '
'            .Col = 6: .Text = ADORs!쿠폰입금 & "" '
'            .Col = 7: .Text = ADORs!접수금액 - ADORs!현금입금 - ADORs!카드입금 - ADORs!사용마일리지 - ADORs!쿠폰입금 & "" '
'
'            .Col = 10: .Text = ADORs!매출시간 & ""
'
'            ADORs.MoveNext
'        Loop
'        ADORs.Close
'        Set ADORs = Nothing
'
'        If (bMisu = False) And 이전고객 = "Y" Then
'            bMisu = True
'
'            .MaxRows = .MaxRows + 1
'            .InsertRows 1, 1
'
'            .Row = 1
'
'            .Col = 1: .Text = "초기미수"      ' 1
'            .Col = 2: .Text = 0 & ""          ' 2
'            .Col = 3: .Text = 0 & ""          ' 3
'            .Col = 4: .Text = 0 & ""          ' 4
'            .Col = 5: .Text = 0 & ""          ' 5
'            .Col = 6: .Text = 0 & ""          ' 6
'            .Col = 7: .Text = 0 & ""          ' 7
'            .Col = 8: .Text = 초기미수금 & "" ' 8
'            .Col = 9: .Text = "초기"          ' 9
'
'            .Row = .Row
'            .Row2 = .Row
'            .Col = 1
'            .Col2 = .MaxCols
'            .BlockMode = True
'            .BackColor = vbYellow
'            .BlockMode = False
'        End If
'
'        If .MaxRows > 0 Then
'            '---------------------------------------------------
'            ' 미수금액 계산
'            '---------------------------------------------------
'            For i = .MaxRows To 1 Step -1
'                .Row = i
'                .Col = 1
'                If .Text = "초기미수금" Then
'                    '초기미수금 제외
'                Else
'                    If i = .MaxRows Then
'                        이전미수 = 0
'                    Else
'                        .Row = i + 1
'                        .Col = 8: 이전미수 = .Value
'                    End If
'
'                    .Row = i
'                    .Col = 7: 이전미수 = 이전미수 + .Value
'                    .Col = 8: .Value = 이전미수 & ""
'                End If
'            Next i
'
'            .Row = 1
'            .Col = 8: 미수금 = .Value: .FontBold = True: .RowHeight(-1) = 14
'
'            If 미수금 <> 고객미수금 Then
'                '---------------------------------------------------
'                ' 미수금액 계산
'                '---------------------------------------------------
'                Query = "SELECT TOP 1 * "
'                Query = Query & " FROM TB_미수금수정"
'                Query = Query & " WHERE 고객코드 = '" & 고객코드 & "'"
'                Query = Query & " ORDER BY 수정일자 DESC"
'                Set ADORs = New ADODB.Recordset
'                ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'                If Not ADORs.EOF Then
'                    .MaxRows = .MaxRows + 1
'                    .InsertRows 1, 1
'
'                    .Row = 1
'                    .Col = 1: .Text = Format(Left(ADORs!수정일자, 10), "YY-MM-DD") & ""
'                    .Col = 2: .Value = 0
'                    .Col = 3: .Value = 0
'                    .Col = 4: .Value = 0
'                    .Col = 5: .Value = 0
'                    .Col = 6: .Value = 0
'                    .Col = 7: .Value = 0
'                    .Col = 8: .Value = ADORs!수정미수금 & "": .FontBold = True: .ForeColor = vbRed: .RowHeight(-1) = 14
'                    .Col = 9: .Text = "조정"
'                End If
'                ADORs.Close
'                Set ADORs = Nothing
'            End If
'        End If
'
'        Dim Misu_Row As Long
'
'        For i = 1 To .MaxRows
'            .Row = i
'
'            .Col = 1
'            If .Text = "초기미수" Then
'                Misu_Row = i
'            End If
'
'            If (Misu_Row > 0) And (i > Misu_Row) Then
'                .Col = 8: .Value = 0
'            End If
'        Next i
'
'        .ReDraw = True
'    End With
'
'    Exit Sub
'
'ErrRtn:
'    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
'
'    Screen.MousePointer = 0
'End Sub

Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        cmdList_Click
    End If
End Sub
