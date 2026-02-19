VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_03003 
   Caption         =   "수기 일일출고"
   ClientHeight    =   11700
   ClientLeft      =   8835
   ClientTop       =   1890
   ClientWidth     =   16035
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_03003.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11700
   ScaleWidth      =   16035
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   11700
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16035
      _ExtentX        =   28284
      _ExtentY        =   20638
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_03003.frx":058A
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10350
         Left            =   15
         TabIndex        =   1
         Top             =   1335
         Width           =   16005
         _Version        =   524288
         _ExtentX        =   28231
         _ExtentY        =   18256
         _StockProps     =   64
         BackColorStyle  =   1
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
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
         SpreadDesigner  =   "P_03003.frx":061C
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   2
         Top             =   540
         Width           =   16005
         _ExtentX        =   28231
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboCount 
            Height          =   315
            Left            =   6960
            Style           =   2  '드롭다운 목록
            TabIndex        =   19
            Top             =   420
            Width           =   1260
         End
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1245
            Locked          =   -1  'True
            Style           =   2  '드롭다운 목록
            TabIndex        =   17
            Top             =   45
            Width           =   3420
         End
         Begin VB.ComboBox cboInput 
            Height          =   315
            Left            =   1245
            Style           =   2  '드롭다운 목록
            TabIndex        =   3
            Top             =   405
            Width           =   3420
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Left            =   6960
            TabIndex        =   4
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Format          =   57540608
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   5775
            TabIndex        =   5
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "출고일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   60
            TabIndex        =   6
            Top             =   405
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "가 맹 점"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   10
            Left            =   60
            TabIndex        =   18
            Top             =   45
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
            Index           =   0
            Left            =   5775
            TabIndex        =   20
            Top             =   420
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "출고차수"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton cmdRefresh 
            Height          =   330
            Left            =   4695
            TabIndex        =   21
            Top             =   390
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   582
            _StockProps     =   79
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_03003.frx":0DFE
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   7
         Top             =   15
         Width           =   8400
         _ExtentX        =   14817
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
         Caption         =   " #"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_03003.frx":1398
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   1
         Left            =   8430
         TabIndex        =   8
         Top             =   15
         Width           =   7590
         _ExtentX        =   13388
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
         PictureBackground=   "P_03003.frx":159A
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   9
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
            Picture         =   "P_03003.frx":179C
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   10
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
            Picture         =   "P_03003.frx":1D36
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   11
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
            Picture         =   "P_03003.frx":22D0
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   12
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
            Picture         =   "P_03003.frx":286A
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   13
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
            Picture         =   "P_03003.frx":2E04
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   14
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
            Picture         =   "P_03003.frx":339E
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   15
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
            Picture         =   "P_03003.frx":3938
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   16
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
            Picture         =   "P_03003.frx":3ED2
         End
      End
   End
End
Attribute VB_Name = "P_03003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim RS02 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cboCount_Click()
    Call Data_Display
End Sub

Private Sub cboInput_Click()
    Call Data_Display
End Sub

Private Sub cboOffice_Click()
    On Error GoTo ErrRtn
    
    If cboOffice.ListIndex < 0 Then Exit Sub
    
    '-----------------------------------------------------------------
    '
    '-----------------------------------------------------------------
    cboInput.Clear
    
    ReDim sValue(2)
    
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    sValue(1) = Format(dtInput.Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput.Value, "YYYY-MM-DD")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01001_00", sValue(), Err_Num, Err_Dec)
    
    Do Until RS01.EOF
        If (RS01!종료일자 = "") Or (RS01!종료일자 = "2099-12-31") Then '가맹점이 현 지사에서 관리중...
            cboInput.AddItem "[" & RS01!가맹점코드 & "] " & RS01!가맹점명
        End If
        
        RS01.MoveNext
    Loop
    RS01.Close
    Set RS01 = Nothing
    
    If cboInput.ListCount > 0 Then cboInput.ListIndex = 0
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: Call Data_Display   ' 조회
        Case 1: Call DataAdd        ' 신규
        Case 2: Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
        Case 4: 'Call DataCancel     ' 취소
        Case 5: 'Call DataPrint      ' 인쇄
        Case 6: Call Export_Excel(P_00000.cdgExcel, spdView)      ' 엑셀
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

Private Sub cmdRefresh_Click()
    cboOffice_Click
End Sub

Private Sub dtInput_Change()
    Call Data_Display
End Sub

Private Sub Form_Activate()
    
    If P_03003_Flag = True Then Exit Sub
    P_03003_Flag = True
    
    cmdBtn(0).Enabled = True
    cmdBtn(1).Enabled = True
    cmdBtn(2).Enabled = True
    cmdBtn(5).Enabled = True
    cmdBtn(6).Enabled = True
    
    '본사 여부
    If HeadOffice = MASTER_OFFICE_CODE Then
        cboOffice.Locked = False
        
        cmdBtn(2).Enabled = False '본사에서는 조회만...
    Else
        cboOffice.Locked = True
    End If
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
        
    With cboCount
        .Clear
        
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        .AddItem "6"
        .AddItem "7"
        .AddItem "8"
        .AddItem "9"
        
        .ListIndex = -1
    End With
    
    Call ComboAdd
    
    dtInput.Value = Date
    
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

'    If P_03003_Flag = False Then
'        P_03003_Flag = True
'    End If
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    P_03003_Flag = False
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
        .OperationMode = OperationModeNormal
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With

    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_03003_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn
    
    If cboInput.ListIndex = -1 Then
        MsgBox "가맹점을 선택하십시오.", vbInformation
        cboInput.SetFocus
        Exit Sub
    End If
    
    If cboCount.ListIndex = -1 Then
        Exit Sub
    End If
    
    
    ReDim sValue(2)
    
    sValue(0) = Format(dtInput.Value, "YYYY-MM-DD") '
    sValue(1) = Mid(cboInput.Text, 2, 6)            '
    sValue(2) = cboCount.Text                       '
    
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_03003_00", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_03003_00", sValue(), Err_Num, Err_Dec)
    End If
        
    With spdView
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = Format(RS01!TAG_NO, "000-00-0000") & ""        '
            .Col = 2:  .Text = "[" & RS01!SCAN_FLAG2 & "] " & Trim(RS01!상태) '
            .Col = 3:  .Text = "[" & RS01!SCAN_FLAG1 & "] " & Trim(RS01!물품) '
            '.Col = 4:  .Text = RS01!OUT_COUNT & ""                            '
            .Col = 4:  .Text = RS01!접수일자 & ""                            '
            .Col = 5:  .Text = "[" & RS01!의류코드 & "] " & Trim(RS01!의류명) '
            .Col = 6:  .Text = RS01!색상 & ""                                 '
            .Col = 7:  .Text = RS01!내용 & ""                                 '
            .Col = 8:  .Text = RS01!금액 & ""                                 '
            .Col = 9:  .Text = RS01!상표 & ""                                 '
            .Col = 10: .Text = RS01!OUT_HAND & ""                             '
            
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

Private Sub ComboAdd()
    Dim sItem As String
    
    '--------------------------------------------------------------------
    ' 출고구분 - 상태
    '--------------------------------------------------------------------
            sItem = "[1] 정상" & Chr(9)
    sItem = sItem & "[2] 반품" & Chr(9)
    sItem = sItem & "[3] 확인" & Chr(9)
    sItem = sItem & "[4] 품명"
    
    spdView.Col = 2: spdView.TypeComboBoxList = sItem
    spdView.Text = "[1] 정상"
    
    '--------------------------------------------------------------------
    ' 소품구분 - 물품
    '--------------------------------------------------------------------
            sItem = "[1] 의류" & Chr(9)
    sItem = sItem & "[2] 소품" & Chr(9)
    sItem = sItem & "[3] 기타"
    
    spdView.Col = 3: spdView.TypeComboBoxList = sItem
    spdView.Text = "[1] 의류"

'    '--------------------------------------------------------------------
'    ' 의류명 SP_00005
'    '--------------------------------------------------------------------
'    ReDim sValue(0)
'
'    sValue(0) = "0"
'
'    Set RS02 = New ADODB.Recordset
'    Set RS02 = ExecPro("SP_00005", sValue(), Err_Num, Err_Dec)
'
'    sItem = ""
'    Do While Not RS02.EOF
'        sItem = sItem & "[" & RS02!의류코드 & "] " & Trim(RS02!의류명) & Chr(9)
'
'        RS02.MoveNext
'    Loop
'
'    spdView.Col = 5
'    spdView.TypeComboBoxList = sItem
    
    '--------------------------------------------------------------------
    ' 색상
    '--------------------------------------------------------------------
    sItem = ""
    sItem = "흰색" & Chr(9)
    sItem = sItem & "상아" & Chr(9)
    sItem = sItem & "회색" & Chr(9)
    sItem = sItem & "쥐색" & Chr(9)
    sItem = sItem & "밤색" & Chr(9)
    sItem = sItem & "검정" & Chr(9)
    sItem = sItem & "분홍" & Chr(9)
    sItem = sItem & "주황" & Chr(9)
    sItem = sItem & "빨강" & Chr(9)
    sItem = sItem & "노랑" & Chr(9)
    sItem = sItem & "베지" & Chr(9)
    sItem = sItem & "황토" & Chr(9)
    sItem = sItem & "연두" & Chr(9)
    sItem = sItem & "초록" & Chr(9)
    sItem = sItem & "카키" & Chr(9)
    sItem = sItem & "쑥색" & Chr(9)
    sItem = sItem & "하늘" & Chr(9)
    sItem = sItem & "파랑" & Chr(9)
    sItem = sItem & "곤색" & Chr(9)
    sItem = sItem & "보라" & Chr(9)
    sItem = sItem & "체크" & Chr(9)
    sItem = sItem & "자주" & Chr(9)
    sItem = sItem & "혼합"
    
    spdView.Col = 6: spdView.TypeComboBoxList = sItem

    '--------------------------------------------------------------------
    ' 출고구분
    '--------------------------------------------------------------------
            sItem = "[1] 수기" & Chr(9)
    sItem = sItem & "[2] 택배" & Chr(9)
    sItem = sItem & "[3] 직배" & Chr(9)
    sItem = sItem & "[4] 점출" & Chr(9)
    sItem = sItem & "[9] 기타"
    
    spdView.Col = 10: spdView.TypeComboBoxList = sItem

    spdView.Text = "[1] 수기"
End Sub

Public Sub DataSave()
    On Error GoTo ErrRtn
        
    Dim i        As Integer
    Dim OUT_DATE As String
    
    If cboCount.ListIndex = -1 Then
        MsgBox "출고차수를 선택하세요.", vbInformation, "확인"
        
        Exit Sub
    End If
    
    With spdView
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            
            If .BackColor = vbYellow Then
                .Col = 1
                If Len(.Value) > 1 And Len(.Value) <> 9 Then
                    MsgBox "택번호를 올바르게 입력하세요.", vbInformation, "확인"
                    
                    Exit Sub
                End If
                
                .Col = 2
                If Trim(.Text) = "" Then
                    MsgBox "출고구분을 올바르게 입력하세요.", vbInformation, "확인"
                    
                    Exit Sub
                End If
                
                .Col = 3
                If Trim(.Text) = "" Then
                    MsgBox "소품구분을 올바르게 입력하세요.", vbInformation, "확인"
                    
                    Exit Sub
                End If
            End If
        Next i
    
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            
            If .BackColor = vbYellow Then
                '-------------------------------------------------------------------------------------
                ' SCANOUTPUT_LOG_TB에 저장하기
                '-------------------------------------------------------------------------------------
                OUT_DATE = Format(dtInput.Value, "YYYY-MM-DD") & " " & Format(Time, "hh:mm:ss")
                
                ReDim sValue(10)
                
                           sValue(0) = OUT_DATE                     ' 0 SCAN_DATE
                           sValue(1) = "00"                         ' 1 PDA_NO
                           sValue(2) = Mid(cboInput.Text, 2, 6)     ' 2 가맹점코드
                .Col = 1:  sValue(3) = Trim(.Value) & ""            ' 3 TAG_NO
                .Col = 3:  sValue(4) = Mid(.Text, 2, 1) & ""        ' 4 SCAN_FLAG1 물품
                .Col = 2:  sValue(5) = Mid(.Text, 2, 1) & ""        ' 5 SCAN_FLAG2 상태
                           sValue(6) = "Y"                          ' 6 OUT_FLAG
                           sValue(7) = OUT_DATE                     ' 7 지사출고일자
                           sValue(8) = cboCount.Text                ' 8 출고차수
                .Col = 10: sValue(9) = Trim(.Text) & ""             ' 9 수기 일일출고에서 사용하는 필드
                .Col = 4:  sValue(10) = Trim(.Text) & ""            ' 10 접수일자
                           
                If sValue(3) <> "" Then
                    Call ExecPro("SP_03003_04", sValue(), Err_Num, Err_Dec)
                End If
            End If
        Next i
    End With
    
    MsgBox "데이터가 정상적으로 저장이 되었습니다.", vbInformation, "확인"
    
    spdView.MaxRows = 0
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Public Sub DataPrint()

End Sub

Private Sub spdView_Change(ByVal Col As Long, ByVal Row As Long)
    spdView.Row = Row
    spdView.Col = -1
    spdView.BackColor = vbYellow
End Sub

Private Sub spdView_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrRtn
    
    Dim sIpDate As String
    
    If KeyCode = vbKeyReturn Then
        Select Case spdView.ActiveCol
            Case 1
                P_03003_01.l_AgencyCode = Mid(cboInput.Text, 2, 6)
                
                spdView.Row = spdView.ActiveRow
                spdView.Col = 1
                P_03003_01.l_TagNo = spdView.Value
                
                P_03003_01.Show vbModal
                
                sIpDate = P_03003_01.l_IpDate
                
                If Len(sIpDate) <> 0 Then
                    Call Data_Display2(sIpDate)
                End If
        End Select
        
        If spdView.ActiveCol = spdView.MaxCols And spdView.ActiveRow = spdView.MaxRows Then
            spdView.MaxRows = spdView.MaxRows + 1
        End If
        
    ElseIf KeyCode = vbKeyDown And spdView.ActiveCol = 2 Then
        spdView.TypeComboBoxIndex = 1
        Debug.Print Now
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Public Sub DataAdd()
    If cboCount.ListIndex = -1 Then
        MsgBox "출고차수를 선택하여 주십시요.", vbInformation, "확인"
        Exit Sub
    End If

    With spdView
        .MaxRows = .MaxRows + 1
'
'        .SetText 2, .MaxRows, "[1] 의류"
'        .SetText 2, .MaxRows, "[1] 의류"
'
'        spdView.Text = "[1] 수기"
    End With
     
End Sub

Private Sub Data_Display2(sIpDate)
    On Error GoTo ErrRtn
    
    ReDim sValue(2)
    
    sValue(0) = Mid(cboInput.Text, 2, 6)       '가맹점코드
    
    spdView.Row = spdView.ActiveRow
    spdView.Col = 1
    sValue(1) = spdView.Value                  '택번호
    sValue(2) = Format(sIpDate, "YYYY-MM-DD")  '접수일자
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_03003_03", sValue(), Err_Num, Err_Dec)
    
    If Not RS01.EOF Then
        spdView.Col = 4: spdView.Text = Format(RS01!접수일자, "YYYY-MM-DD")                                      '입고일자 IpDate
        spdView.Col = 5: spdView.Text = IIf(IsNull(RS01!의류코드), "", "[" & RS01!의류코드 & "] " & Trim(RS01!의류명)) '코드 GoodsCode, GoodsName
        spdView.Col = 6: spdView.Text = IIf(IsNull(RS01!색상), "", RS01!색상)                                    '색상 Color
        spdView.Col = 7: spdView.Value = IIf(IsNull(RS01!내용), "", RS01!내용)                                   '내용 Worked
        spdView.Col = 8: spdView.Value = IIf(IsNull(RS01!금액), 0, RS01!금액)                                    '금액 Amount
        spdView.Col = 9: spdView.Value = IIf(IsNull(RS01!상표), "", RS01!상표)                                   '상표 Label
        
        spdView.Col = -1: spdView.BackColor = vbYellow
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub
