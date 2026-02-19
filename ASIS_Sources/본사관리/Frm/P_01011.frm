VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_01011 
   Caption         =   "가맹점 품목 현황"
   ClientHeight    =   11790
   ClientLeft      =   1665
   ClientTop       =   3375
   ClientWidth     =   16125
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_01011.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11790
   ScaleWidth      =   16125
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   11790
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16125
      _ExtentX        =   28443
      _ExtentY        =   20796
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_01011.frx":058A
      Begin Threed.SSPanel panCaption 
         Height          =   390
         Index           =   2
         Left            =   3555
         TabIndex        =   1
         Top             =   1335
         Width           =   12555
         _ExtentX        =   22146
         _ExtentY        =   688
         _Version        =   262144
         Font3D          =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " 가맹점 품목 현황"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_01011.frx":069C
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   2
         Top             =   540
         Width           =   16095
         _ExtentX        =   28390
         _ExtentY        =   1376
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
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1125
            TabIndex        =   14
            Text            =   "cboOffice"
            Top             =   75
            Width           =   3405
         End
         Begin VB.Label lblTitle 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "지 사 명:"
            Height          =   225
            Index           =   24
            Left            =   0
            TabIndex        =   15
            Top             =   135
            Width           =   1065
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   3
         Top             =   15
         Width           =   8535
         _ExtentX        =   15055
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
         Caption         =   " 가맹점 품목 현황 (P_01011)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_01011.frx":0AFE
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Left            =   8565
         TabIndex        =   4
         Top             =   15
         Width           =   7545
         _ExtentX        =   13309
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
         PictureBackground=   "P_01011.frx":0D00
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
            Picture         =   "P_01011.frx":0F02
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
            Picture         =   "P_01011.frx":149C
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
            Picture         =   "P_01011.frx":1A36
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
            Picture         =   "P_01011.frx":1FD0
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
            Picture         =   "P_01011.frx":256A
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
            Picture         =   "P_01011.frx":2B04
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
            Picture         =   "P_01011.frx":309E
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
            Picture         =   "P_01011.frx":3638
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10440
         Left            =   15
         TabIndex        =   13
         Top             =   1335
         Width           =   3525
         _Version        =   524288
         _ExtentX        =   6218
         _ExtentY        =   18415
         _StockProps     =   64
         BackColorStyle  =   1
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
         MaxCols         =   2
         ScrollBars      =   2
         SpreadDesigner  =   "P_01011.frx":3BD2
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdView2 
         Height          =   10035
         Left            =   9750
         TabIndex        =   16
         Top             =   1740
         Width           =   6360
         _Version        =   524288
         _ExtentX        =   11218
         _ExtentY        =   17701
         _StockProps     =   64
         BackColorStyle  =   1
         EditEnterAction =   5
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
         MaxCols         =   5
         ScrollBars      =   2
         SpreadDesigner  =   "P_01011.frx":40BC
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdView1 
         Height          =   10035
         Left            =   5820
         TabIndex        =   17
         Top             =   1740
         Width           =   3915
         _Version        =   524288
         _ExtentX        =   6906
         _ExtentY        =   17701
         _StockProps     =   64
         BackColorStyle  =   1
         EditEnterAction =   5
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
         MaxCols         =   2
         ScrollBars      =   2
         SpreadDesigner  =   "P_01011.frx":4681
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdView3 
         Height          =   10035
         Left            =   3555
         TabIndex        =   18
         Top             =   1740
         Width           =   2250
         _Version        =   524288
         _ExtentX        =   3969
         _ExtentY        =   17701
         _StockProps     =   64
         BackColorStyle  =   1
         EditEnterAction =   5
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
         MaxCols         =   1
         ScrollBars      =   2
         SpreadDesigner  =   "P_01011.frx":4B74
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_01011"
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
    Call Data_Display(Mid(cboOffice.Text, 2, 4))
End Sub

Private Sub Data_Display(지사코드 As String)
    On Error GoTo ErrRtn

    '-----------------------------------------------------------------
    '
    '-----------------------------------------------------------------
    spdView.MaxRows = 0
    
    ReDim sValue(2)
    
    sValue(0) = 지사코드 'HeadOffice
    sValue(1) = Format(Date, "YYYY-MM-DD")
    sValue(2) = Format(Date, "YYYY-MM-DD")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01001_00", sValue(), Err_Num, Err_Dec)
    
    With spdView
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            If (RS01!종료일자 = "") Or (RS01!종료일자 = "2099-12-31") Then '가맹점이 현 지사에서 관리중...
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                
                .Col = 1: .Text = RS01!가맹점코드 & "" ' 1
                .Col = 2: .Text = RS01!가맹점명 & ""   ' 2
            End If
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Data_Display2()
    '-------------------------------------------------------------
    ' TB_의류분류
    '-------------------------------------------------------------
    ReDim sValue(0)

    sValue(0) = "0"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_00013", sValue(), Err_Num, Err_Dec)

    With spdView1
        .MaxRows = 0
        .Redraw = False
                    
        '-전체-
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        
        .Col = 1: .Text = ""
        .Col = 2: .Text = "-전체-"
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = RS01!의류분류코드 & ""
            .Col = 2: .Text = RS01!의류분류명 & ""
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
End Sub

Private Sub cboOffice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SearchString_한글 KeyAscii
'    Else
       ' SearchString KeyAscii
    End If

End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: 'Call Data_Display   ' 조회
        Case 1: 'Call DataAdd        ' 신규
        Case 2: Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
        Case 4: 'Call DataCancel     ' 취소
        Case 5: 'Call DataPrint      ' 인쇄
        Case 6: Call Export_Excel(P_00000.cdgExcel, spdView2)      ' 엑셀
        Case 7: Unload Me           ' 종료
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
'    cmdBtn(0).Enabled = True
'    cmdBtn(2).Enabled = True
'    cmdBtn(4).Enabled = True
    cmdBtn(6).Enabled = True
    
    '본사 여부
    If HeadOffice = MASTER_OFFICE_CODE Then
        cboOffice.Locked = False
    Else
        cboOffice.Locked = True
    End If
        
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
    If P_01011_Flag = False Then
'        Call AgencyComboAdd(cboInput)
'
'        '----------------------------------------------------------------
'        ' SP_01011_00
'        '----------------------------------------------------------------
'        ReDim sValue(1)
'
'        sValue(0) = "1"
'
'        Set RS01 = New ADODB.Recordset
'        Set RS01 = ExecPro("SP_01011_00", sValue(), Err_Num, Err_Dec)
'
'        spdView(0).MaxCols = RS01.Fields.Count
'        spdView(0).MaxRows = RS01.RecordCount
'
'        'Call spdDisplay(RS01)
'        Call fpSpread_Display(spdView(0), RS01)
'        Call GetColWidth(REG_App, Me.Name & "A", spdView(0))
'
'        '----------------------------------------------------------------
'        ' SP_01011_01
'        '----------------------------------------------------------------
'        ReDim sValue(1)
'
'        sValue(0) = "1"
'
'        Set RS01 = New ADODB.Recordset
'        Set RS01 = ExecPro("SP_01011_01", sValue(), Err_Num, Err_Dec)
'
'        spdView(1).MaxCols = RS01.Fields.Count
'        spdView(1).MaxRows = RS01.RecordCount
'
'        'Call spdDisplay2(RS01)
'        Call fpSpread_Display(spdView(1), RS01)
'        Call GetColWidth(REG_App, Me.Name & "B", spdView(1))
'
'        P_01011_Flag = True
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
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

    With spdView1
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

    With spdView2
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

    With spdView3
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
    
    Dim i As Integer
    
    With cboOffice
        For i = 0 To .ListCount - 1
            If Mid(.List(i), 2, 4) = HeadOffice Then
                .ListIndex = i
                
                Exit For
            End If
        Next i
    End With
    
    If P_01011_Flag = False Then
        'Call AgencyComboAdd(cboInput)
        
'        ReDim sValue(1)
'
'        sValue(0) = "1"
'
'        Set RS01 = New ADODB.Recordset
'        Set RS01 = ExecPro("SP_01011_00", sValue(), Err_Num, Err_Dec)
'
'        spdView(0).MaxCols = RS01.Fields.Count
'        spdView(0).MaxRows = RS01.RecordCount
'
'        'Call spdDisplay(RS01)
'        Call fpSpread_Display(spdView(0), RS01)
'        Call GetColWidth(REG_App, Me.Name & "A", spdView(0))
'
'        ReDim sValue(1)
'
'        sValue(0) = "1"
'
'        Set RS01 = New ADODB.Recordset
'        Set RS01 = ExecPro("SP_01011_01", sValue(), Err_Num, Err_Dec)
'
'        spdView(1).MaxCols = RS01.Fields.Count
'        spdView(1).MaxRows = RS01.RecordCount
'
'        'Call spdDisplay2(RS01)
'        Call fpSpread_Display(spdView(1), RS01)
'        Call GetColWidth(REG_App, Me.Name & "B", spdView(1))
        
        P_01011_Flag = True
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_01011_Flag = False
End Sub

Private Sub Data_Display3(가맹점코드 As String, 적용일자 As String, 의류분류코드 As String)
    ReDim sValue(3)
    
    sValue(0) = "0"
    sValue(1) = 가맹점코드
    sValue(2) = 적용일자
    sValue(3) = 의류분류코드
        
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01011_01", sValue(), Err_Num, Err_Dec)
        
    With spdView2
        .MaxRows = 0
        .Redraw = False
                    
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = RS01!품목코드 & ""
            .Col = 2: .Text = RS01!품목명 & ""
            .Col = 3: .Text = RS01!단가 & ""
            .Col = 4: .Text = RS01!적용일자 & ""
            .Col = 5: .Text = RS01!수신정보 & ""
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
End Sub

Public Sub DataSave()
'    Dim i As Integer
'
'    ReDim sValue(4)
'
'    If Not IsDate(dtInput.Value) Then
'        MsgBox " 적용일자가 선택되지 않았읍니다.", vbInformation, "확인"
'        Exit Sub
'
'    ElseIf dtInput.Value < Date Then
'        MsgBox " 적용일자를 확인하여 주십시요.", vbInformation, "오류"
'        Exit Sub
'    End If
'
'    sValue(0) = Mid(cboInput.Text, 2, 6)
'    sValue(1) = Format(dtInput.Value, "YYYY-MM-DD")
'
'    strSql = "DELETE  AgencyGoodsCT WHERE AgencyCode = '" & sValue(0) & "'"
'    strSql = strSql + "               AND SDate      = '" & sValue(1) & "'"
'    Set RS01 = New ADODB.Recordset
'    Call SqlDataValue(RS01, strSql)
'
'    If spdView(1).MaxRows > 0 Then
'        For i = 1 To spdView(1).MaxRows
'            sValue(0) = Mid(cboInput.Text, 2, 6)
'            sValue(1) = Format(dtInput.Value, "YYYY-MM-DD")
'
'            spdView(1).Row = i
'            spdView(1).Col = 1: sValue(2) = spdView(1).Text & ""
'            spdView(1).Col = 2: sValue(3) = spdView(1).Text & ""
'            spdView(1).Col = 3: sValue(4) = spdView(1).Value & ""
'
'            Call ExecPro("SP_01011_02", sValue(), Err_Num, Err_Dec)
'        Next i
'    End If
End Sub

'-------------------------------------------------------------------
' 가맹점 Spread
'-------------------------------------------------------------------
Private Sub spdView_Click(ByVal Col As Long, ByVal Row As Long)
    If Row <= 0 Then Exit Sub
    
    Call 적용일자_Display

End Sub

Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Call spdView_Click(NewCol, NewRow)
End Sub

'-------------------------------------------------------------------
' 품목분류 Spread
'-------------------------------------------------------------------
Private Sub spdView1_Click(ByVal Col As Long, ByVal Row As Long)
    If Row <= 0 Then Exit Sub
    
    Call 의류_Display
End Sub

Private Sub spdView1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Call spdView1_Click(NewCol, NewRow)
End Sub

Private Sub 의류_Display()
    Dim 가맹점코드   As String
    Dim 의류분류코드 As String
    Dim 적용일자     As String
    
    If spdView.ActiveRow <= 0 Then
        spdView2.MaxRows = 0
        Exit Sub
    End If
    
    If spdView1.ActiveRow <= 0 Then
        spdView2.MaxRows = 0
        Exit Sub
    End If
    
    If spdView3.ActiveRow <= 0 Then
        spdView2.MaxRows = 0
        Exit Sub
    End If
    
    spdView.Row = spdView.ActiveRow
    spdView.Col = 1: 가맹점코드 = spdView.Text & ""
    
    spdView1.Row = spdView1.ActiveRow
    spdView1.Col = 1: 의류분류코드 = spdView1.Text & ""
    
    spdView3.Row = spdView3.ActiveRow
    spdView3.Col = 1: 적용일자 = spdView3.Text & ""
    
    Call Data_Display3(가맹점코드, 적용일자, 의류분류코드)
End Sub


Private Sub 적용일자_Display()
    Dim 가맹점코드   As String
    Dim 의류분류코드 As String
    
    If spdView.ActiveRow <= 0 Then
        spdView1.MaxRows = 0
        spdView2.MaxRows = 0
        spdView3.MaxRows = 0
        Exit Sub
    End If
    
    
    spdView.Row = spdView.ActiveRow
    spdView.Col = 1: 가맹점코드 = spdView.Text & ""
    
    ReDim sValue(1)
    
    sValue(0) = "0"
    sValue(1) = 가맹점코드
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01011_05", sValue(), Err_Num, Err_Dec)
        
    With spdView3
        .MaxRows = 0
        .Redraw = False
                    
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = RS01!적용일자 & ""
            
            RS01.MoveNext
        Loop
        RS01.Close: Set RS01 = Nothing
        .Redraw = True
    End With

End Sub

Private Sub spdView3_Click(ByVal Col As Long, ByVal Row As Long)
                    
    spdView2.MaxRows = 0
    ' 상품 대 분류 정보 출력
    Call Data_Display2

End Sub
