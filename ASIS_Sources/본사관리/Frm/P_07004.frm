VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_07004 
   Caption         =   "외주 입고 현황"
   ClientHeight    =   10185
   ClientLeft      =   630
   ClientTop       =   2460
   ClientWidth     =   17085
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_07004.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10185
   ScaleWidth      =   17085
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   10185
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17085
      _ExtentX        =   30136
      _ExtentY        =   17965
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_07004.frx":058A
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   1
         Top             =   540
         Width           =   17055
         _ExtentX        =   30083
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   0
            Left            =   5040
            TabIndex        =   22
            Top             =   60
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.ComboBox cboNum 
            Height          =   315
            Left            =   8100
            Style           =   2  '드롭다운 목록
            TabIndex        =   21
            Top             =   420
            Width           =   1335
         End
         Begin VB.ComboBox cboPage 
            Height          =   315
            Left            =   10425
            Style           =   2  '드롭다운 목록
            TabIndex        =   20
            Top             =   420
            Width           =   1080
         End
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1245
            Style           =   2  '드롭다운 목록
            TabIndex        =   2
            Top             =   60
            Width           =   2850
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1245
            TabIndex        =   3
            Top             =   405
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            Format          =   68616192
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   4
            Top             =   405
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "입고일자"
            BorderWidth     =   0
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   7
            Left            =   60
            TabIndex        =   5
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "외주업체"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   4140
            TabIndex        =   23
            Top             =   420
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            Format          =   68616192
            CurrentDate     =   36686
         End
         Begin VB.Label lblTitle 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "택번호:"
            Height          =   195
            Index           =   0
            Left            =   4245
            TabIndex        =   26
            Top             =   120
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "입고회차:"
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
            Index           =   4
            Left            =   7170
            TabIndex        =   25
            Top             =   480
            Width           =   885
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "인쇄매수:"
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
            Index           =   5
            Left            =   9300
            TabIndex        =   24
            Top             =   480
            Width           =   1080
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   6
         Top             =   15
         Width           =   8745
         _ExtentX        =   15425
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
         Caption         =   " 외주 입고 현황 (P_07004)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_07004.frx":065C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   8835
         Left            =   15
         TabIndex        =   7
         Top             =   1335
         Width           =   5280
         _Version        =   524288
         _ExtentX        =   9313
         _ExtentY        =   15584
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
         MaxCols         =   4
         MaxRows         =   35
         ScrollBars      =   2
         SpreadDesigner  =   "P_07004.frx":085E
         UserResize      =   1
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   420
         Index           =   2
         Left            =   5310
         TabIndex        =   8
         Top             =   1335
         Width           =   11760
         _ExtentX        =   20743
         _ExtentY        =   741
         _Version        =   262144
         Font3D          =   3
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " 지사 외주 현황"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_07004.frx":0E6D
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.ProgressBar ProgressBar 
            Height          =   270
            Left            =   5910
            TabIndex        =   9
            Top             =   45
            Visible         =   0   'False
            Width           =   3270
            _Version        =   851970
            _ExtentX        =   5768
            _ExtentY        =   476
            _StockProps     =   93
            Scrolling       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
      End
      Begin FPSpreadADO.fpSpread spdViewScan 
         Height          =   8400
         Left            =   5310
         TabIndex        =   10
         Top             =   1770
         Width           =   11760
         _Version        =   524288
         _ExtentX        =   20743
         _ExtentY        =   14817
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
         MaxCols         =   8
         MaxRows         =   35
         ScrollBars      =   0
         SpreadDesigner  =   "P_07004.frx":12CF
         UserResize      =   1
         Appearance      =   1
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   0
         Left            =   8775
         TabIndex        =   11
         Top             =   15
         Width           =   8295
         _ExtentX        =   14631
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
         PictureBackground=   "P_07004.frx":19A8
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   12
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
            Picture         =   "P_07004.frx":1BAA
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   13
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "화면"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_07004.frx":2144
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   14
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
            Picture         =   "P_07004.frx":26DE
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   15
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
            Picture         =   "P_07004.frx":2C78
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   16
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
            Picture         =   "P_07004.frx":3212
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   17
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
            Picture         =   "P_07004.frx":37AC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   18
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
            Picture         =   "P_07004.frx":3D46
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   19
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
            Picture         =   "P_07004.frx":42E0
         End
      End
   End
End
Attribute VB_Name = "P_07004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub SPR_Resize()
    On Error GoTo ErrRtn
    
    spdViewScan.Width = Me.Width - 5610
    spdViewScan.Height = Me.Height - 3900

    Exit Sub
    
ErrRtn:

End Sub

Private Sub cboOffice_Click()
    If cboOffice.ListIndex < 0 Then Exit Sub
    
    Call Data_Display
End Sub

'-----------------------------------------------------------------
'
'-----------------------------------------------------------------
Private Sub Data_Display()
    On Error GoTo ErrRtn

    ReDim sValue(4)
    Dim nCnt    As Long
    
    nCnt = 0
    sValue(0) = Store.Code
    sValue(1) = Trim(Mid(cboOffice.Text, 2, 4)) + "%"
    sValue(2) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(3) = Format(dtInput(1).Value, "YYYY-MM-DD")
    sValue(4) = CStr(cboNum.Text)
    
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("[SP_M_07004_00]", sValue(), Err_Num, Err_Dec)
    
    With spdView
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            
            .Col = 1: .Text = RS01!코드 & ""
            .Col = 2: .Text = RS01!가맹점명 & ""
            .Col = 3: .Text = RS01!스캔수량 & ""
            .Col = 4: .Text = RS01!택번호 & ""
            nCnt = nCnt + Val(RS01!스캔수량 & "")
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
    
    
        If .MaxRows >= 2 Then
            .MaxRows = .MaxRows + 1
            .Row = 1
            .Action = SS_ACTION_INSERT_ROW
            
            .Col = 1: .Text = ""
            .Col = 2: .Text = "전   체"
            .Col = 3: .Text = Format(nCnt, "#,##0")
        End If
    
        .Redraw = True
    End With
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: Call Data_Display    ' 조회
        Case 1: 'Call DataAdd        ' 신규
        Case 2: 'Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
        Case 4: 'Call DataCancel     ' 취소
        Case 5: Call DataPrint      ' 인쇄
        Case 6: 'Call DataScreen     '
        Case 7: Unload Me            ' 종료
        Case 9: 'Call Data_Update
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
 

Private Sub dtInput_Change(Index As Integer)
    dtInput(Index).Enabled = False

    Call Data_Display
    
    dtInput(Index).Enabled = True
End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(5).Enabled = True
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
    If P_07004_Flag = False Then
        Dim i As Integer
        dtInput(0).Value = Date
        dtInput(1).Value = Date

        '
        Call OrderComboAdd(cboOffice)
        
        With cboOffice
            For i = 0 To .ListCount - 1
                If Mid(.List(i), 2, 4) = HeadOffice Then
                    .ListIndex = i
                    
                    Exit For
                End If
            Next i
        End With
        
        P_07004_Flag = True
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
    
    Dim i As Integer
    
    With spdViewScan
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
    
    With cboNum
        .Clear
        .AddItem ""
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        .ListIndex = 0
    End With
    
    With cboPage
        .Clear
        
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        .ListIndex = 0
    End With
    
    Call SPR_Resize

    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Resize()
    Call SPR_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_07004_Flag = False
End Sub

Private Sub Data_Display2()
    On Error GoTo ErrRtn
    
    ReDim sValue(6)
    
    spdView.Row = spdView.ActiveRow
        
    sValue(0) = Store.Code
    spdView.Col = 1:        sValue(1) = spdView.Text + "%"
    spdView.Col = 4:        sValue(2) = spdView.Text + "%"
    sValue(3) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(4) = Format(dtInput(1).Value, "YYYY-MM-DD")
    sValue(5) = "%"
    
'    sValue(3) = Mid(cboOffice.Text, 2, 4)
'    sValue(5) = Format(DateAdd("d", -30, dtInput(0).Value), "YYYY-MM-DD")
'    sValue(6) = Format(dtInput(0).Value, "YYYY-MM-DD")
    
    '------------------------------------------------------------
    ' 외주 출고 등록 - SP_M_07002_01
    '------------------------------------------------------------
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_07004_01", sValue(), Err_Num, Err_Dec)
    
    With spdViewScan
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = Format(RS01!TagNo & "", "@@@-@@-@@@@")   'KEY
            .Col = 2: .Text = RS01!의류코드 & ""  '
            .Col = 3: .Text = RS01!의류명 & ""    '
            .Col = 4: .Text = RS01!입고일자 & ""
            .Col = 5: .Text = RS01!회차 & ""    '
            .Col = 6: .Text = Format(RS01!출고일자 & "", "@@@@-@@-@@")      '
            .Col = 7: .Text = RS01!INSCANDT & ""      'PDA 스켄일자
            .Col = 8: .Text = RS01!INPDANO & ""       'PDA NO
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
        
    Exit Sub
    
ErrRtn:
    MsgBox Err.Description
    Resume
End Sub


Private Sub spdView_Click(ByVal Col As Long, ByVal Row As Long)
    If Row <= 0 Then Exit Sub
    Call Data_Display2
    
End Sub

Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    'Call spdView_Click(NewCol, NewRow)
End Sub



Private Sub DataPrint()
    On Error GoTo ErrRtn
    Dim 지사명      As String
    Dim 지사코드    As String
    
    Dim 택번호      As String
    Dim XML         As String
    Dim i           As Integer
    Dim FileNumber
        
    If spdViewScan.DataRowCnt <= 0 Then Exit Sub
    
    With spdView
        .Row = .ActiveRow
        
        .Col = 1:   지사코드 = .Text
        .Col = 2:   지사명 = .Text
    End With
    
    FileNumber = FreeFile
    
    Open App.Path & "\XML\P_07004.XML" For Output As #FileNumber
    
    Print #FileNumber, "<?xml version=""1.0"" encoding=""EUC-KR""?>"
    Print #FileNumber, "<root>"
    
          XML = "    <조건>" & vbLf
    XML = XML & "        <지사>" & Func_Replace(지사명) & " (" & 지사코드 & ")  입고내역</지사>" & vbLf
    XML = XML & "        <입고수량>입고수량 : " & spdViewScan.MaxRows & " 점</입고수량>" & vbLf
    XML = XML & "        <입고일자>입고일자 : " & Format(dtInput(0).Value, "YYYY년 MM월 DD일") & "~" & Format(dtInput(1).Value, "YYYY년 MM월 DD일") & "</입고일자>" & vbLf
    XML = XML & "   </조건>" & vbLf
    Print #FileNumber, XML
    
    With spdViewScan
        
        For i = 1 To .DataRowCnt
            .Row = i
            XML = "    <Data>" & vbLf
            .Col = 4: XML = XML & "        <입고일자>" & Trim(.Text) & "</입고일자>" & vbLf
            .Col = 1: XML = XML & "        <택번호>" & Trim(.Text) & "</택번호>" & vbLf
            .Col = 2: XML = XML & "        <의류코드>" & Trim(.Text) & "</의류코드>" & vbLf
            .Col = 3: XML = XML & "        <의류명>" & Trim(.Text) & "</의류명>" & vbLf
            .Col = 5: XML = XML & "        <회차>" & Trim(.Text) & "</회차>" & vbLf
            XML = XML & "   </Data>" & vbLf
            Print #FileNumber, XML
        Next i
        
        Print #FileNumber, "</root>" & vbLf
        Close #FileNumber
    End With
    
    With rpt외주입고현황
        .documentName = "외주입고현황"
        .dc.FileURL = App.Path & "\XML\P_07004.XML"
        .Printer.Copies = cboPage.Text '인쇄매수
        .PrintReport False
        
        '.Show 1
    End With

    Unload rpt외주입고현황
    
    Exit Sub

ErrRtn:
    MsgBox Err.Description, vbInformation, "오류"
    Screen.MousePointer = 0
End Sub
