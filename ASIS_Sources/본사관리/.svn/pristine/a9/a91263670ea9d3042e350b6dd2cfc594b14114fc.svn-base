VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMask32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_05018 
   Caption         =   "물세탁 일지"
   ClientHeight    =   9675
   ClientLeft      =   4410
   ClientTop       =   5235
   ClientWidth     =   17220
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_05018.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9675
   ScaleWidth      =   17220
   WindowState     =   2  '최대화
   Begin Threed.SSPanel pnlProg 
      Height          =   1215
      Left            =   6330
      TabIndex        =   0
      Top             =   4380
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
      Picture         =   "P_05018.frx":058A
      Caption         =   " "
      BorderWidth     =   0
      BevelOuter      =   1
      BevelInner      =   2
      Alignment       =   4
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   9675
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   17220
      _ExtentX        =   30374
      _ExtentY        =   17066
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_05018.frx":3555
      Begin Threed.SSPanel panInput 
         Height          =   2265
         Left            =   15
         TabIndex        =   2
         Top             =   540
         Width           =   17190
         _ExtentX        =   30321
         _ExtentY        =   3995
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboGoods2 
            Height          =   315
            Left            =   5940
            Style           =   2  '드롭다운 목록
            TabIndex        =   26
            Top             =   1506
            Width           =   3420
         End
         Begin VB.ComboBox cboGoods 
            Height          =   315
            Left            =   1245
            Style           =   2  '드롭다운 목록
            TabIndex        =   25
            Top             =   1506
            Width           =   3420
         End
         Begin VB.ComboBox cboFablic 
            Height          =   315
            Left            =   5940
            Style           =   2  '드롭다운 목록
            TabIndex        =   24
            Top             =   1860
            Width           =   3420
         End
         Begin VB.ComboBox cboColor 
            Height          =   315
            Left            =   1245
            Style           =   2  '드롭다운 목록
            TabIndex        =   22
            Top             =   1860
            Width           =   3420
         End
         Begin VB.ComboBox cboInput 
            Height          =   315
            Left            =   1245
            Style           =   2  '드롭다운 목록
            TabIndex        =   18
            Top             =   798
            Width           =   3420
         End
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1245
            Locked          =   -1  'True
            TabIndex        =   14
            Text            =   "cboOffice"
            Top             =   444
            Width           =   3420
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   15
            Top             =   798
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "가맹점명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   10
            Left            =   60
            TabIndex        =   16
            Top             =   444
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "지    사"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   60
            TabIndex        =   17
            Top             =   1152
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "택 번 호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSMask.MaskEdBox mskInput 
            Height          =   315
            Left            =   1245
            TabIndex        =   19
            Top             =   1152
            Width           =   3420
            _ExtentX        =   6033
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            MaxLength       =   11
            Mask            =   "###-##-####"
            PromptChar      =   "_"
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   3
            Left            =   60
            TabIndex        =   20
            Top             =   1506
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "품    명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   4
            Left            =   60
            TabIndex        =   21
            Top             =   1860
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "색    상"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   5
            Left            =   4755
            TabIndex        =   23
            Top             =   1860
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "소    재"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   6
            Left            =   4755
            TabIndex        =   27
            Top             =   1506
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "세부품명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1245
            TabIndex        =   28
            Top             =   90
            Width           =   3420
            _ExtentX        =   6033
            _ExtentY        =   556
            _Version        =   393216
            Format          =   63373312
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   29
            Top             =   90
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "조회일자"
            BorderWidth     =   0
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   5010
            TabIndex        =   30
            Top             =   90
            Width           =   3420
            _ExtentX        =   6033
            _ExtentY        =   556
            _Version        =   393216
            Format          =   63373312
            CurrentDate     =   36686
         End
         Begin VB.Label Label 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
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
            Left            =   4785
            TabIndex        =   31
            Top             =   150
            Width           =   120
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   3
         Top             =   15
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   900
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   255
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
         Caption         =   " 물세탁 일지 (P_05018)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_05018.frx":35E7
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   1
         Left            =   9585
         TabIndex        =   4
         Top             =   15
         Width           =   7620
         _ExtentX        =   13441
         _ExtentY        =   900
         _Version        =   262144
         Font3D          =   3
         ForeColor       =   192
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
         PictureBackgroundStyle=   2
         PictureBackground=   "P_05018.frx":37E9
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   5
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "종료"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_05018.frx":39EB
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   6
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "엑셀"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_05018.frx":3F85
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   7
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "인쇄"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_05018.frx":451F
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   8
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "취소"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_05018.frx":4AB9
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   9
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "삭제"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_05018.frx":5053
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   10
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "저장"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_05018.frx":55ED
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   11
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "신규"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_05018.frx":5B87
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   12
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "조회"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_05018.frx":6121
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   6840
         Left            =   15
         TabIndex        =   13
         Top             =   2820
         Width           =   17190
         _Version        =   524288
         _ExtentX        =   30321
         _ExtentY        =   12065
         _StockProps     =   64
         BackColorStyle  =   1
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
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
         MaxCols         =   11
         MaxRows         =   34
         ScrollBars      =   2
         SpreadDesigner  =   "P_05018.frx":66BB
         UserResize      =   1
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_05018"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ADOConEJ As New ADODB.Connection     ' ActiveX Database Object 연결
Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String


Private Sub cboGoods_Click()
    Dim Query As String
    cboGoods2.Clear
    Query = ""
    Query = Query + " SELECT '0000' as [의류코드],'전체' as [의류명]"
    Query = Query + " UNION ALL"
    Query = Query + " SELECT [의류코드], [의류명]"
    Query = Query + " From [TB_의류]"
    Query = Query + " WHERE substring(의류코드,1,2) = '" & Mid(cboGoods.Text, 2, 2) & "'"
    Query = Query + " ORDER BY 1"
    
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecQuery(Query, Err_Num, Err_Dec)

    With cboGoods2
        Do Until RS01.EOF
            .AddItem "[" & RS01!의류코드 & "] " & RS01!의류명
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        If .ListCount > 0 Then .ListIndex = 0
    End With
End Sub

Private Sub cboOffice_Click()
    If cboOffice.ListIndex < 0 Then Exit Sub
    
    '-----------------------------------------------------------------
    '
    '-----------------------------------------------------------------
    cboInput.Clear

    spdView.MaxRows = 0
    Dim Query As String
    ReDim sValue(2)
    
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_01001_00", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_01001_00", sValue(), Err_Num, Err_Dec)
    End If

    With cboInput
        .AddItem "[000000] 전체"
        Do Until RS01.EOF
            .AddItem "[" & RS01!가맹점코드 & "] " & RS01!가맹점명
        
            'If (RS01!종료일자 = "") Or (RS01!종료일자 = "2099-12-31") Then '가맹점이 현 지사에서 관리중...
            '    .AddItem "[" & RS01!가맹점코드 & "] " & RS01!가맹점명
            'End If
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
    
    
    
End Sub

Private Sub cboOffice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SearchString_한글 KeyAscii
    End If
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: Call Data_Display  ' 조회
        Case 7: Unload Me           ' 종료
        Case 1: P_05018_AddNew.Show vbModal, P_00000
        Case 2: 'Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
        Case 4: 'Call DataCancel     ' 취소
        Case 5: 'Call DataPrint        ' 인쇄
        Case 6: Call Export_Excel(P_00000.cdgExcel, spdView(0))      ' 엑셀
    End Select
    
'    Me.MousePointer = 0
    
    Exit Sub
    
ErrRtn:
    Me.MousePointer = 0
    
    If Err.Number = "0" Then
        
    ElseIf Err.Number = "91" Then
        End
    Else
        Resume Next
    End If
End Sub

 

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(1).Enabled = True
    cmdBtn(6).Enabled = True
    
    If HeadOffice = MASTER_OFFICE_CODE Then
        cboOffice.Locked = False
    Else
        cboOffice.Locked = True
    End If
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    Dim i As Integer
    
 
    
    With spdView
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
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
    
    dtInput(0).Value = Date
    dtInput(1).Value = Date

    '
    Call Get_지사리스트(cboOffice)

    With cboOffice
        
        For i = 0 To .ListCount - 1
            If Mid(.List(i), 2, 4) = HeadOffice Then
                .ListIndex = i
                
                Exit For
            End If
        Next i
        .ListIndex = 0
    End With
    
    Call GetGoods
    Call GetColor
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Data_Display()
    Dim Query As String
    
    
    Query = ""
    Query = Query + " SELECT "
    Query = Query + " a.작업일자,"
    Query = Query + " a.지사코드, "
    Query = Query + " b.지사명,"
    Query = Query + " a.가맹점코드,"
    Query = Query + " c.가맹점명,"
    Query = Query + " a.택번호,"
    Query = Query + " a.의류명,"
    Query = Query + " a.색상,"
    Query = Query + " a.무늬,"
    Query = Query + " a.소재,"
    Query = Query + " a.세탁방법"
    Query = Query + " From LAUNDRY1000..TB_물세탁 a INNER JOIN TB_지사 b on a.지사코드 = b.지사코드"
    Query = Query + "                  INNER JOIN TB_가맹점 c ON a.가맹점코드 = c.가맹점코드"
    Query = Query + " WHERE"
    Query = Query + " (A.작업일자 >= '" & dtInput(0).Value & "' AND A.작업일자 <= '" & dtInput(1).Value & "')"
    
    If cboOffice.Text <> "" And Mid(cboOffice.Text, 2, 4) <> "0000" Then Query = Query + " AND A.지사코드 = '" & Mid(cboOffice.Text, 2, 4) & "'"
    If cboInput.Text <> "" And Mid(cboInput.Text, 2, 6) <> "000000" Then Query = Query + " AND A.가맹점코드 = '" & Mid(cboInput.Text, 2, 6) & "'"
    
    If mskInput.Text <> "" Then Query = Query + " AND A.택번호 = '" & mskInput.Text & "'"
    If cboGoods.ListIndex > 0 Then Query = Query + " AND substring(A.의류코드,1,2) = '" & Mid(cboGoods.Text, 2, 2) & "'"
    If cboGoods2.ListIndex > 0 Then Query = Query + " AND A.의류코드 = '" & Mid(cboGoods2.Text, 2, 4) & "'"
    If cboColor.ListIndex > 0 Then Query = Query + " AND A.색상 = '" & cboColor.Text & "'"
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecQuery(Query, Err_Num, Err_Dec)


    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount

    Call fpSpread_Display(spdView, RS01, False)

    If spdView.MaxRows = 0 Then
        MsgBox "조회된 자료가 없습니다."
    End If

End Sub
 

Private Sub DataPrint()

End Sub

Private Sub DataScreen()
    
End Sub

 
Private Sub DataSave()
    On Error GoTo ErrRtn
    
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)

End Sub

Private Sub GetGoods()
    Dim Query As String
    cboGoods.Clear
    
    Query = ""
    Query = Query + " SELECT '00' as [의류분류코드],'전체' as [의류분류명]"
    Query = Query + " UNION ALL"
    Query = Query + " SELECT 의류분류코드, 의류분류명"
    Query = Query + " From TB_의류분류"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecQuery(Query, Err_Num, Err_Dec)

    With cboGoods
        Do Until RS01.EOF
            .AddItem "[" & RS01!의류분류코드 & "] " & RS01!의류분류명
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        If .ListCount > 0 Then .ListIndex = 0
    End With
End Sub


Private Sub GetColor()
    Dim Query As String
    cboColor.Clear
    
    Query = ""
    Query = Query + " SELECT '전체' as [색상명]"
    Query = Query + " UNION ALL"
    Query = Query + " SELECT 색상명"
    Query = Query + " From [TB_색상표]"
    Query = Query + " WHERE 색상명 <> '품목보기'"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecQuery(Query, Err_Num, Err_Dec)

    With cboColor
        Do Until RS01.EOF
            .AddItem "" & RS01!색상명
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        If .ListCount > 0 Then .ListIndex = 0
    End With
End Sub

