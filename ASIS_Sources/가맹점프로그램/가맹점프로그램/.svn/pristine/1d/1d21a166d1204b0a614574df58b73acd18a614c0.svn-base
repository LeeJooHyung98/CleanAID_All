VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm물건찾기 
   Caption         =   "물건 찾기"
   ClientHeight    =   10080
   ClientLeft      =   6375
   ClientTop       =   1965
   ClientWidth     =   15240
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10080
   ScaleWidth      =   15240
   WindowState     =   2  '최대화
   Begin Threed.SSPanel pnlProg 
      Height          =   1215
      Left            =   600
      TabIndex        =   7
      Top             =   1605
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
      Picture         =   "frm물건찾기.frx":0000
      Caption         =   " "
      BorderWidth     =   0
      BevelOuter      =   1
      BevelInner      =   2
      Alignment       =   4
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   10080
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   17780
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm물건찾기.frx":2FCB
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   8850
         Left            =   15
         TabIndex        =   5
         Top             =   1215
         Width           =   15210
         _Version        =   524288
         _ExtentX        =   26829
         _ExtentY        =   15610
         _StockProps     =   64
         AutoCalc        =   0   'False
         BackColorStyle  =   1
         ColsFrozen      =   7
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         EditModeReplace =   -1  'True
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
         MaxCols         =   10
         MaxRows         =   200
         Protect         =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frm물건찾기.frx":303D
         UserResize      =   1
         VisibleCols     =   7
         VisibleRows     =   30
         HighlightHeaders=   1
         HighlightStyle  =   1
         ScrollBarStyle  =   2
      End
      Begin Threed.SSPanel Panel 
         Height          =   750
         Left            =   15
         TabIndex        =   8
         Top             =   450
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   1323
         _Version        =   262144
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   9765
            TabIndex        =   4
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frm물건찾기.frx":3E1E
         End
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   630
            Left            =   3615
            TabIndex        =   1
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 조회(&F)"
            Appearance      =   6
            Picture         =   "frm물건찾기.frx":4EB0
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   3
            Left            =   5190
            TabIndex        =   2
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 엑셀(&E)"
            Appearance      =   6
            Picture         =   "frm물건찾기.frx":55AA
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   4
            Left            =   6735
            TabIndex        =   3
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 출력(&P)"
            Appearance      =   6
            Picture         =   "frm물건찾기.frx":5D24
         End
         Begin CSTextLibCtl.sitxEdit txtTAGNo 
            Height          =   630
            Left            =   915
            TabIndex        =   0
            Top             =   60
            Width           =   1635
            _Version        =   262145
            _ExtentX        =   2884
            _ExtentY        =   1111
            _StockProps     =   125
            Text            =   "__-____"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
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
            StartText.y     =   6
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   29
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "택번호:"
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
            Index           =   2
            Left            =   60
            TabIndex        =   10
            Top             =   90
            Width           =   795
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   9
         Top             =   15
         Width           =   15210
         _ExtentX        =   26829
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
         Caption         =   "      물건 찾기"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm물건찾기.frx":641E
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm물건찾기.frx":6644
            Top             =   -15
            Width           =   765
         End
      End
   End
End
Attribute VB_Name = "frm물건찾기"
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
        
        Case 5: Unload Me
    End Select
    
    Exit Sub
    
ErrRtn:
    pnlProg.Visible = False
    
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
    
    Screen.MousePointer = 0
End Sub

Private Sub cmdList_Click()
    On Error GoTo ErrRtn
    
    Dim 바코드  As String
    Dim TempMsg As String
    
    If Trim(txtTAGNo.RawData) = "" Then
        MsgBox "택번호를 입력하세요.", vbInformation, "확인"
        
        txtTAGNo.SetFocus
        
        Exit Sub
    End If
    
    바코드 = TAG_Convert(Trim(txtTAGNo.RawData))
    
    '-------------------------------------------------------------
    ' 6 자리 이하의 택번호를 입력한 경우 가맹점 택코드를 넣어준다.
    '-------------------------------------------------------------
    If Len(바코드) = 6 Then
        바코드 = 가맹점정보.택코드 & 바코드
    End If
    
    If Left(바코드, 3) <> 가맹점정보.택코드 Then
        'txtTAGNo(1).Text = "다른 지점의 세탁물입니다."
        'Ret = sndPlaySound(AppPath & "Sound\다른지점.wav", 0)
    
        MsgBox "다른 가맹점의 세탁물입니다.", vbCritical, "확인"
    
        txtTAGNo.Text = ""
        Exit Sub
    End If
    
    '------------------------------------------------------------------------------------
    ' 택번호를 중복해서 읽었는지 체크...
    '------------------------------------------------------------------------------------
    '바코드 = Left(바코드, 9)
    
    With sprGrid
        Rtn = .SearchCol(1, -1, -1, Format(바코드, "000-00-0000"), SearchFlagsValue)
        
        If Rtn > -1 Then
            .SetSelection 1, Rtn, .MaxCols, Rtn
        
            'txtTAGNo(1).Text = "택번호 중복"
            'Ret = sndPlaySound(AppPath & "Sound\택번호중복.wav", 0)

            txtTAGNo.Text = ""
            
            Exit Sub
        End If
    End With
    
    If Server_Connection(HostCon) = True Then
        Query = "SELECT    택번호"
        Query = Query & ", 접수일자"
        Query = Query & ", 의류명"
        Query = Query & ", 색상"
        Query = Query & ", 무늬"
        Query = Query & ", 내용"
        Query = Query & ", 금액"
        Query = Query & ", 상표"
        Query = Query & ", ISNULL(지사출고예정,'')   AS 지사출고예정"
        Query = Query & ", ISNULL(가맹점출고일자,'') AS 가맹점출고일자"
        Query = Query & ", ISNULL(지사입고일자,'')   AS 지사입고일자"
        Query = Query & ", ISNULL(지사출고일자,'')   AS 지사출고일자"
        Query = Query & ", ISNULL(지사출고예정,'')   AS 지사출고예정"
        Query = Query & ", ISNULL(출고일자,'')       AS 출고일자"
        Query = Query & " FROM TB_입출고"
        Query = Query & " WHERE 가맹점코드  = '" & 가맹점정보.가맹점코드 & "'"
        Query = Query & "   AND (택번호     = '" & 바코드 & "'"
        Query = Query & "    OR  택번호 LIKE  '%" & txtTAGNo.RawData & "')"
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, HostCon, adOpenForwardOnly, adLockReadOnly
    
        If ADORs.EOF Then
            With sprGrid
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                
                .Col = 1:  .Text = Format(바코드, "000-00-0000") & ""
                .Col = 2:  .Text = ""
                .Col = 3:  .Text = ""
                .Col = 4:  .Text = ""
                .Col = 5:  .Text = ""
                .Col = 6:  .Text = ""
                .Col = 7:  .Text = ""
                .Col = 8:  .Text = ""
                .Col = 9:  .Text = ""
                .Col = 10: .Text = "지사DB에 데이터가 없음."
            End With
        Else
            With sprGrid
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                
                .Col = 1:  .Text = Format(ADORs!택번호, "000-00-0000") & ""
                .Col = 2:  .Text = ADORs!접수일자 & ""
                .Col = 3:  .Text = ADORs!의류명 & ""
                .Col = 4:  .Text = ADORs!색상 & ""
                .Col = 5:  .Text = ADORs!무늬 & ""
                .Col = 6:  .Text = ADORs!내용 & ""
                .Col = 7:  .Text = ADORs!금액 & ""
                .Col = 8:  .Text = ADORs!상표 & ""
                .Col = 9:  .Text = ADORs!지사출고예정 & ""
                
                If ADORs!가맹점출고일자 = "" Then
                    TempMsg = "가맹점미출고,"
                End If
                
                If ADORs!가맹점출고일자 <> "" And ADORs!지사입고일자 = "" Then
                    TempMsg = TempMsg & "지사미입고,"
                End If
                
                If ADORs!지사입고일자 <> "" And ADORs!지사출고일자 = "" Then
                    TempMsg = TempMsg & "지사미출고,"
                End If
                
                If ADORs!지사출고예정 <> "" And ADORs!지사출고일자 = "" Then
                    TempMsg = TempMsg & "지사출고예정,"
                End If
                
                If ADORs!지사출고일자 <> "" Then
                    TempMsg = TempMsg & "지사출고,"
                End If
                
                If ADORs!출고일자 <> "" Then
                    TempMsg = TempMsg & "고객출고"
                End If
                
                .Col = 10: .Text = TempMsg & ""
            End With
        End If
        ADORs.Close
        Set ADORs = Nothing
    
        With sprGrid
            .Row = .MaxRows
            '----------------------------------------------------------------------------------
            ' TB_물건찾기
            '----------------------------------------------------------------------------------
                               Query = "INSERT INTO TB_물건찾기"
                       Query = Query & "           (지사코드"
                       Query = Query & "           ,가맹점코드"
                       Query = Query & "           ,택번호"
                       Query = Query & "           ,조회일자"
                       Query = Query & "           ,의류코드"
                       Query = Query & "           ,의류명"
                       Query = Query & "           ,색상"
                       Query = Query & "           ,무늬"
                       Query = Query & "           ,내용"
                       Query = Query & "           ,금액"
                       Query = Query & "           ,상표"
                       Query = Query & "           ,메모)"
                       Query = Query & "     VALUES"
                       Query = Query & "           ( '" & 가맹점정보.지사코드 & "'"
                       Query = Query & "           , '" & 가맹점정보.가맹점코드 & "'"
            .Col = 1:  Query = Query & "           , '" & Replace(.Text, "-", "") & "'"
                       Query = Query & "           , '" & Format(Now, "YYYY-MM-DD hh:mm:ss") & "'"
                       Query = Query & "           , ''"
            .Col = 3:  Query = Query & "           , '" & Trim(.Text) & "'"
            .Col = 4:  Query = Query & "           , '" & Trim(.Text) & "'"
            .Col = 5:  Query = Query & "           , '" & Trim(.Text) & "'"
            .Col = 6:  Query = Query & "           , '" & Trim(.Text) & "'"
            .Col = 7:  Query = Query & "           , '" & Trim(.Text) & "'"
            .Col = 8:  Query = Query & "           , '" & Trim(.Text) & "'"
            .Col = 10: Query = Query & "           , '" & Trim(.Text) & "')"
                       HostCon.Execute Query
            '----------------------------------------------------------------------------------
        End With
    End If
    
    txtTAGNo.Text = ""
    txtTAGNo.SetFocus
    
    Exit Sub
    
ErrRtn:

End Sub

Private Sub Form_Load()
    With sprGrid
        .MaxRows = 0
        .RowHeight(-1) = 16
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeSingle
        
        ' the User Sort
        .UserColAction = UserColActionSort
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    pnlHeader.Width = Me.ScaleWidth
    
    cmdBtn(5).Left = Me.Width - cmdBtn(5).Width - 200
End Sub

Public Sub txtTAGNo_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrRtn

    If KeyAscii = 13 Then
        KeyAscii = 0
        
        Call cmdList_Click
    End If
    
Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
    
    Screen.MousePointer = 0
End Sub
