VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_05018_AddNew 
   BorderStyle     =   1  '단일 고정
   Caption         =   "물세탁 신규 입력창"
   ClientHeight    =   5265
   ClientLeft      =   12435
   ClientTop       =   4590
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   8760
   Begin VB.TextBox txtTag 
      Appearance      =   0  '평면
      Height          =   315
      Left            =   1245
      TabIndex        =   4
      Top             =   975
      Width           =   3420
   End
   Begin FPSpreadADO.fpSpread spdView 
      Height          =   1365
      Left            =   1245
      TabIndex        =   9
      Top             =   1320
      Visible         =   0   'False
      Width           =   7455
      _Version        =   524288
      _ExtentX        =   13150
      _ExtentY        =   2408
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
      MaxCols         =   9
      MaxRows         =   34
      ScrollBars      =   2
      SpreadDesigner  =   "P_05018_AddNew.frx":0000
      UserResize      =   1
      HighlightHeaders=   1
      HighlightStyle  =   1
   End
   Begin VB.TextBox txtColor 
      Appearance      =   0  '평면
      Height          =   315
      Left            =   1245
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1665
      Width           =   3420
   End
   Begin VB.TextBox txtGoods 
      Appearance      =   0  '평면
      Height          =   315
      Left            =   1245
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1320
      Width           =   3420
   End
   Begin VB.TextBox txtFabric 
      Appearance      =   0  '평면
      Height          =   315
      IMEMode         =   10  '한글 
      Left            =   1245
      TabIndex        =   5
      Top             =   2010
      Width           =   3420
   End
   Begin VB.TextBox txtMethod 
      Appearance      =   0  '평면
      Height          =   2310
      IMEMode         =   10  '한글 
      Left            =   1245
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2355
      Width           =   7455
   End
   Begin VB.ComboBox cboOffice 
      Height          =   300
      Left            =   1245
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "cboOffice"
      Top             =   600
      Width           =   3420
   End
   Begin Threed.SSPanel pnlHeader 
      Height          =   510
      Left            =   -15
      TabIndex        =   0
      Top             =   -15
      Width           =   8760
      _ExtentX        =   15452
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
      Caption         =   " 물세탁 신규입력 (P_05018)"
      PictureBackgroundStyle=   2
      PictureBackground=   "P_05018_AddNew.frx":0687
      BorderWidth     =   0
      BevelOuter      =   0
      Alignment       =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdBtn 
      Height          =   450
      Index           =   7
      Left            =   7800
      TabIndex        =   1
      Top             =   4770
      Width           =   900
      _Version        =   851970
      _ExtentX        =   1587
      _ExtentY        =   794
      _StockProps     =   79
      Caption         =   "닫기"
      ForeColor       =   -2147483640
      BackColor       =   -2147483636
      Appearance      =   6
      Picture         =   "P_05018_AddNew.frx":0889
   End
   Begin XtremeSuiteControls.PushButton cmdBtn 
      Height          =   450
      Index           =   2
      Left            =   6840
      TabIndex        =   2
      Top             =   4770
      Width           =   900
      _Version        =   851970
      _ExtentX        =   1587
      _ExtentY        =   794
      _StockProps     =   79
      Caption         =   "저장"
      ForeColor       =   -2147483640
      BackColor       =   -2147483636
      Appearance      =   6
      Picture         =   "P_05018_AddNew.frx":0E23
   End
   Begin Threed.SSPanel panCaption 
      Height          =   315
      Index           =   10
      Left            =   60
      TabIndex        =   7
      Top             =   600
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
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
      Caption         =   "지    사"
      BevelOuter      =   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel panCaption 
      Height          =   315
      Index           =   1
      Left            =   60
      TabIndex        =   8
      Top             =   975
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
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
      Caption         =   "택 번 호"
      BevelOuter      =   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel panCaption 
      Height          =   315
      Index           =   0
      Left            =   60
      TabIndex        =   10
      Top             =   2355
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
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
      Caption         =   "세탁방법"
      BevelOuter      =   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel panCaption 
      Height          =   315
      Index           =   2
      Left            =   60
      TabIndex        =   11
      Top             =   2010
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
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
      Caption         =   "소    재"
      BevelOuter      =   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel panCaption 
      Height          =   315
      Index           =   3
      Left            =   60
      TabIndex        =   13
      Top             =   1320
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
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
      Caption         =   "품    명"
      BevelOuter      =   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel panCaption 
      Height          =   315
      Index           =   5
      Left            =   60
      TabIndex        =   15
      Top             =   1665
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
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
      Caption         =   "색    상"
      BevelOuter      =   1
      RoundedCorners  =   0   'False
   End
End
Attribute VB_Name = "P_05018_AddNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cboOffice_Click()
On Error Resume Next
    txtTag.SetFocus
End Sub

Private Sub cboOffice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SearchString_한글 KeyAscii
    End If
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
    Case 2
        Call DataSave
    Case 7
        Unload Me
    End Select
End Sub

Private Sub Form_Activate()
    If HeadOffice = MASTER_OFFICE_CODE Then
        cboOffice.Locked = False
    Else
        cboOffice.Locked = True
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = P_00000.Icon
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
    
    Call Get_지사리스트(cboOffice)

    With cboOffice
        
        For i = 0 To .ListCount - 1
            If Mid(.List(i), 2, 4) = HeadOffice Then
                .ListIndex = i
                
                Exit For
            End If
        Next i
    End With
    'cboOffice.SetFocus
End Sub


Private Sub spdView_DblClick(ByVal Col As Long, ByVal Row As Long)
    With spdView
        .Row = .ActiveRow
        .Col = 1: txtTag.Text = .Text       ' 택번호
        .Col = 8: txtTag.Tag = .Text        ' 접수일자
        .Col = 3: txtGoods.Text = .Text     ' 상품명
        .Col = 7: txtGoods.Tag = .Text      ' 상품코드
        .Col = 4: txtColor.Text = .Text     ' 색상
        .Col = 9: txtColor.Tag = .Text      ' 가맹점 코드
        .Visible = False
        txtFabric.SetFocus
    End With
End Sub

Private Sub spdView_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spdView_DblClick 0, 0
    End If
End Sub

Private Sub spdView_LostFocus()
    spdView.Visible = False
End Sub

Private Sub txtTag_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'If Len(txtTag.Text) < 5 Then spdView.Visible = True
        SearchTag
    End If
End Sub

Private Sub SearchTag()
    Dim Query As String
    Query = "SELECT * FROM ("
    Query = Query + " SELECT a.지사코드,a.접수일자,a.가맹점코드,가맹점명,a.택번호,의류코드,의류명,색상,상표,case 출고일자 WHEN '' THEN '9999-99-99' ELSE 출고일자 END as 출고일자"
    Query = Query + " FROM tb_입출고 a left join tb_가맹점 b on a.지사코드 = b.지사코드 and a.가맹점코드 = b.가맹점코드"
    Query = Query + " WHERE a.택번호 like '%" & txtTag.Text & "'"
    Query = Query + " ) a"
    Query = Query + " ORDER BY 출고일자 DESC"

    If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecQueryMaster(Query, Err_Num, Err_Dec)

    With spdView
        .MaxRows = 0
        Do Until RS01.EOF
            Debug.Print RS01!가맹점명
            
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 1: .Text = RS01!택번호
            .Col = 2: .Text = RS01!가맹점명
            .Col = 3: .Text = RS01!의류명
            .Col = 4: .Text = RS01!색상
            .Col = 5: .Text = RS01!상표
            .Col = 6: .Text = RS01!출고일자
            .Col = 7: .Text = RS01!의류코드
            .Col = 8: .Text = RS01!접수일자
            .Col = 9: .Text = RS01!가맹점코드
            RS01.MoveNext
            
            
        Loop
        RS01.Close
        If .MaxRows > 1 Then .Visible = True
        .SetFocus
    End With
    Set RS01 = Nothing
    
    
End Sub

Private Sub DataSave()
Dim Query As String
    Query = "INSERT INTO [LAUNDRY1000]..TB_물세탁"
    Query = Query + " SELECT 지사코드, 가맹점코드, convert(varchar,getdate(),23) as 작업일자,택번호, 의류코드, 의류명, 색상, 무늬,'" & txtFabric.Text & "' as 소재, 상표, '" & txtMethod.Text & "' as 세탁방법"
    Query = Query + " FROM tb_입출고"
    Query = Query + " WHERE 가맹점코드 = '" & txtColor.Tag & "' and 접수일자 = '" & txtTag.Tag & "' and 택번호 = '" & txtTag.Text & "'"

    If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecQueryMaster(Query, Err_Num, Err_Dec)
    
    If Err_Num < 0 Then
        MsgBox "자료저장중 오류가 발생되었습니다."
        Exit Sub
    Else
        txtColor.Text = ""
        txtColor.Tag = ""
        txtTag.Text = ""
        txtTag.Tag = ""
        txtFabric.Text = ""
        txtGoods.Text = ""
        txtGoods.Tag = ""
        txtMethod.Text = ""
        txtTag.SetFocus
    End If

End Sub
