VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_04001 
   Caption         =   "가맹점 매출현황 (지사 기준)"
   ClientHeight    =   9870
   ClientLeft      =   3750
   ClientTop       =   3780
   ClientWidth     =   16485
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_04001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9870
   ScaleWidth      =   16485
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   9870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16485
      _ExtentX        =   29078
      _ExtentY        =   17410
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_04001.frx":058A
      Begin XtremeSuiteControls.TabControl TabControl1 
         Height          =   8520
         Left            =   15
         TabIndex        =   20
         Top             =   1335
         Width           =   16455
         _Version        =   851970
         _ExtentX        =   29025
         _ExtentY        =   15028
         _StockProps     =   68
         Appearance      =   3
         Color           =   64
         PaintManager.ShowIcons=   -1  'True
         PaintManager.ButtonMargin=   "3,2,3,2"
         ItemCount       =   2
         SelectedItem    =   1
         Item(0).Caption =   " 가맹점별 매출현황 "
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "TabControlPage1"
         Item(1).Caption =   "가맹점별 매출집계 "
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "TabControlPage2"
         Begin XtremeSuiteControls.TabControlPage TabControlPage2 
            Height          =   8070
            Left            =   30
            TabIndex        =   22
            Top             =   420
            Width           =   16395
            _Version        =   851970
            _ExtentX        =   28919
            _ExtentY        =   14235
            _StockProps     =   1
            Page            =   1
            Begin FPSpreadADO.fpSpread spdView1 
               Height          =   6405
               Left            =   45
               TabIndex        =   24
               Top             =   45
               Width           =   16455
               _Version        =   524288
               _ExtentX        =   29025
               _ExtentY        =   11298
               _StockProps     =   64
               BackColorStyle  =   1
               ColsFrozen      =   2
               DisplayRowHeaders=   0   'False
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
               MaxCols         =   28
               SpreadDesigner  =   "P_04001.frx":061C
               Appearance      =   1
               HighlightHeaders=   1
               HighlightStyle  =   1
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage1 
            Height          =   8070
            Left            =   -69970
            TabIndex        =   21
            Top             =   420
            Visible         =   0   'False
            Width           =   16395
            _Version        =   851970
            _ExtentX        =   28919
            _ExtentY        =   14235
            _StockProps     =   1
            Page            =   0
            Begin FPSpreadADO.fpSpread spdView 
               Height          =   5505
               Left            =   45
               TabIndex        =   23
               Top             =   45
               Width           =   16455
               _Version        =   524288
               _ExtentX        =   29025
               _ExtentY        =   9710
               _StockProps     =   64
               BackColorStyle  =   1
               ColsFrozen      =   3
               DisplayRowHeaders=   0   'False
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
               MaxCols         =   28
               SpreadDesigner  =   "P_04001.frx":1767
               Appearance      =   1
               HighlightHeaders=   1
               HighlightStyle  =   1
            End
         End
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   1
         Top             =   540
         Width           =   16455
         _ExtentX        =   29025
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1245
            Locked          =   -1  'True
            TabIndex        =   13
            Text            =   "cboOffice"
            Top             =   60
            Width           =   3420
         End
         Begin VB.ComboBox cboInput 
            Height          =   315
            Left            =   1245
            Style           =   2  '드롭다운 목록
            TabIndex        =   12
            Top             =   405
            Width           =   3420
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   6525
            TabIndex        =   14
            Top             =   60
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            Format          =   64356352
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   15
            Top             =   405
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
            Index           =   2
            Left            =   5340
            TabIndex        =   16
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "매출일자"
            BorderWidth     =   0
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   7
            Left            =   60
            TabIndex        =   17
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "지    사"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   9705
            TabIndex        =   18
            Top             =   60
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            Format          =   64356352
            CurrentDate     =   36686
         End
         Begin XtremeSuiteControls.PushButton cmdRefresh 
            Height          =   330
            Left            =   4695
            TabIndex        =   25
            Top             =   390
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   582
            _StockProps     =   79
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_04001.frx":28A2
         End
         Begin VB.Label Label 
            Alignment       =   2  '가운데 맞춤
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
            Height          =   255
            Left            =   9390
            TabIndex        =   19
            Top             =   120
            Width           =   300
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   8850
         _ExtentX        =   15610
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
         Caption         =   " 가맹점 매출현황 (지사 기준) (P_04001)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_04001.frx":2E3C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   1
         Left            =   8880
         TabIndex        =   3
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
         PictureBackground=   "P_04001.frx":303E
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   4
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
            Picture         =   "P_04001.frx":3240
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   5
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
            Picture         =   "P_04001.frx":37DA
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   6
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
            Picture         =   "P_04001.frx":3D74
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   7
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
            Picture         =   "P_04001.frx":430E
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   8
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
            Picture         =   "P_04001.frx":48A8
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   9
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
            Picture         =   "P_04001.frx":4E42
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   10
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
            Picture         =   "P_04001.frx":53DC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   11
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
            Picture         =   "P_04001.frx":5976
         End
      End
   End
End
Attribute VB_Name = "P_04001"
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
    
    spdView.Width = Me.Width - 300
    spdView.Height = Me.Height - 2330

    spdView1.Width = Me.Width - 300
    spdView1.Height = Me.Height - 2330

    Exit Sub
    
ErrRtn:

End Sub

Private Sub cboInput_Click()
    Call Data_Display
End Sub

Private Sub cboOffice_Click()
    If cboOffice.ListIndex < 0 Then Exit Sub
    
    '-----------------------------------------------------------------
    '
    '-----------------------------------------------------------------
    cboInput.Clear
    
    ReDim sValue(2)
    
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01001_00", sValue(), Err_Num, Err_Dec)
    
    cboInput.AddItem "[000000] 전체"
    
    Do Until RS01.EOF
        'If (RS01!종료일자 = "") Or (RS01!종료일자 = "2099-12-31") Then '가맹점이 현 지사에서 관리중...
            cboInput.AddItem "[" & RS01!가맹점코드 & "] " & RS01!가맹점명
        'End If
        
        RS01.MoveNext
    Loop
    RS01.Close
    Set RS01 = Nothing
    
    If cboInput.ListCount > 0 Then cboInput.ListIndex = 0
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
        Case 0: Call Data_Display   ' 조회
        Case 1: 'Call DataAdd        ' 신규
        Case 2: Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
        Case 4: Call DataCancel     ' 취소
        Case 5: 'Call DataPrint      ' 인쇄
        Case 6:
            If TabControl1.SelectedItem = 0 Then
                Call Export_Excel(P_00000.cdgExcel, spdView)      ' 엑셀
            Else
                Call Export_Excel(P_00000.cdgExcel, spdView1)     ' 엑셀
            End If
            
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

'Private Sub cmdSubBtn_Click(Index As Integer)
'    Dim i As Integer
'    Dim iChulTotal As Integer
'
'    DoEvents
'
'    If Index = 0 Then
'        ReDim sValue(0)
'
'        sValue(0) = Format(dtInput.Value, "YYYY-MM-DD")
'
'        Set RS01 = New ADODB.Recordset
'        Set RS01 = ExecPro("SP_04001_04", sValue(), Err_Num, Err_Dec)
'
'        Do While Not RS01.EOF
'            For i = 1 To spdView.MaxRows
'                spdView.Row = i
'                spdView.Col = 1
'
'                If Mid(spdView.Text, 2, 3) = RS01!대리점코드 Then
'                    spdView.Col = 3: spdView.Text = RS01!출고량
'                End If
'            Next i
'
'            iChulTotal = iChulTotal + RS01!출고량
'
'            txtInput(1).Text = Format(iChulTotal, "#,##0")
'
'            RS01.MoveNext
'        Loop
'    Else
'        For i = 1 To spdView.MaxRows
'            spdView.Row = i
'            spdView.Col = 11
'            If spdView.Value = True Then
'                spdView.Value = False
'            Else
'                spdView.Value = True
'            End If
'        Next i
'    End If
'End Sub

Private Sub dtInput_Change(Index As Integer)
'    Call Data_Display
End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    'cmdBtn(2).Enabled = True
    cmdBtn(4).Enabled = True
    'cmdBtn(5).Enabled = True
    cmdBtn(6).Enabled = True
    
    '본사 여부
    If HeadOffice = MASTER_OFFICE_CODE Then
        cboOffice.Locked = False
    Else
        cboOffice.Locked = True
    End If
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
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
'        .OperationMode = OperationModeSingle
        
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
'        .OperationMode = OperationModeSingle
        
        'Init the User Sort
        .UserColAction = UserColActionSort
 
    End With
    
    dtInput(0).Value = Date
    dtInput(1).Value = Date
    
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


    Call SPR_Resize

'    If P_04001_Flag = False Then
'        dtInput.Value = DateAdd("d", -1, Date)
'
'        ReDim sValue(2)
'
'        sValue(0) = "1"
'
'        Set RS01 = New ADODB.Recordset
'        Set RS01 = ExecPro("SP_04001_00", sValue(), Err_Num, Err_Dec)
'
'        spdView.MaxCols = RS01.Fields.Count
'        spdView.MaxRows = RS01.RecordCount
'
'        Call spdDisplay(RS01)
'        Call GetColWidth(REG_App, Me.Name, spdView)
'
'        P_04001_Flag = True
'    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Resize()
    Call SPR_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_04001_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i As Integer
    Dim nCol    As Long
    
    ReDim sValue(3)
    
    Screen.MousePointer = vbHourglass
    
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    
    If Mid(cboInput.Text, 2, 6) = "000000" Then
        sValue(1) = ""
    Else
        sValue(1) = Mid(cboInput.Text, 2, 6)
    End If
    
    sValue(2) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(3) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(HeadOffice) = False Then
            spdView.MaxRows = 0
            
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_04001_A_00", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_04001_00", sValue(), Err_Num, Err_Dec)
    End If
        
    With spdView
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = RS01!가맹점코드 & ""          ' 1
            .Col = 2:  .Text = RS01!가맹점명 & ""            ' 2
            .Col = 3:  .Text = RS01!마감일자 & ""            ' 3
            .Col = 4:  .Text = RS01!접수금액 & ""            ' 4
            .Col = 5:  .Text = RS01!지사금액 & ""            ' 4
            .Col = 6:  .Text = RS01!가맹점금액 & ""          ' 5
            .Col = 7:  .Text = RS01!접수수량 & ""            ' 6
            .Col = 8:  .Text = RS01!출고수량 & ""            ' 7
            
            If Len(RS01!이전종료택번호) = 9 Then
                .Col = 9:  .Text = Format(RS01!이전종료택번호, "000-00-0000") & ""     ' 8
            Else
                .Col = 9:  .Text = RS01!이전종료택번호 & ""  ' 8
            End If
            
            If Len(RS01!시작택번호) = 9 Then
                .Col = 10:  .Text = Format(RS01!시작택번호, "000-00-0000") & ""         ' 9
            Else
                .Col = 10:  .Text = RS01!시작택번호 & ""      ' 9
            End If
            
            If Len(RS01!종료택번호) = 9 Then
                .Col = 11: .Text = Format(RS01!종료택번호, "000-00-0000") & ""         '10
            Else
                .Col = 11: .Text = RS01!종료택번호 & ""      '10
            End If
            
            .Col = 12: .Text = RS01!접수금액 & ""            '11
            .Col = 13: .Text = RS01!현금입금 + RS01!카드금액 & "" '12
            .Col = 14: .Text = RS01!현금입금 & ""            '13
            .Col = 15: .Text = RS01!카드금액 & ""            '14
            .Col = 16: .Text = RS01!카드건수 & ""            '15
            .Col = 17: .Text = RS01!쿠폰금액 & ""            '16
            .Col = 18: .Text = RS01!쿠폰건수 & ""            '17
            .Col = 19: .Text = RS01!발생마일리지 & ""        '18
            .Col = 20: .Text = RS01!사용마일리지 & ""        '19
            .Col = 21: .Text = RS01!삭제마일리지 & ""        '20
            .Col = 22: .Text = RS01!반품환불금액 & ""        '21
            .Col = 23: .Text = RS01!반품환불건수 & ""        '22
            .Col = 24: .Text = RS01!세탁환불금액 & ""        '23
            .Col = 25: .Text = RS01!세탁환불건수 & ""        '24
            .Col = 26: .Text = RS01!재세탁수량 & ""          '25
            .Col = 27: .Text = RS01!수선금액 & ""            '26
            .Col = 28: .Text = RS01!수선수량 & ""            '27
            
            If RS01!접수수량 > 0 Then
                If IsNumeric(RS01!시작택번호) And IsNumeric(RS01!종료택번호) Then
                    If RS01!접수수량 <> (RS01!종료택번호 - RS01!시작택번호) + 1 Then
                        .Col = -1: .BackColor = vbYellow
                    End If
                End If
            End If
            
            If IsNumeric(RS01!이전종료택번호) = True Then
                If (RS01!이전종료택번호 + 1) = RS01!시작택번호 Then
                    .Col = -1: .BackColor = vbBlue
                End If
            End If
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        ' 합계 출력
        If Mid(cboInput.Text, 2, 6) <> "000000" Then
            For nCol = 4 To .MaxCols
                Select Case nCol
                    Case 4: Call SpreadSum(spdView, 2, nCol)
                    Case Else: Call SpreadSum(spdView, -1, nCol)
                End Select
            Next nCol
        End If
         
         .Redraw = True
    End With
    
    Call Data_Display2
        
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrRtn:
    Resume Next
    
    Screen.MousePointer = vbDefault
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Data_Display2()
    Dim i As Integer
    
    ReDim sValue(3)
    
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    
    If Mid(cboInput.Text, 2, 6) = "000000" Then
        sValue(1) = ""
    Else
        sValue(1) = Mid(cboInput.Text, 2, 6)
    End If
    
    sValue(2) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(3) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_04001_A_01", sValue(), Err_Num, Err_Dec)
    
    With spdView1
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = RS01!가맹점코드 & ""               ' 1
            .Col = 2:  .Text = RS01!가맹점명 & ""                 ' 2
            .Col = 3:  .Text = RS01!영업일수 & ""                  ' 3
            .Col = 4:  .Text = RS01!접수금액 & ""                  ' 3
            
            .Col = 5:  .Text = RS01!지사금액 & ""                  ' 3
            .Col = 6:  .Text = RS01!가맹점금액 & ""                ' 4
            .Col = 7:  .Text = RS01!접수수량 & ""                 ' 5
            .Col = 8:  .Text = RS01!출고수량 & ""                 ' 6
            
            If Len(RS01!이전종료택번호) = 9 Then
                .Col = 9:  .Text = Format(RS01!이전종료택번호, "000-00-0000") & ""     ' 8
            Else
                .Col = 9:  .Text = RS01!이전종료택번호 & ""  ' 8
            End If
            
            If Len(RS01!시작택번호) = 9 Then
                .Col = 10:  .Text = Format(RS01!시작택번호, "000-00-0000") & ""         ' 9
            Else
                .Col = 10:  .Text = RS01!시작택번호 & ""      ' 9
            End If
            
            If Len(RS01!종료택번호) = 9 Then
                .Col = 11: .Text = Format(RS01!종료택번호, "000-00-0000") & ""         '10
            Else
                .Col = 11: .Text = RS01!종료택번호 & ""      '10
            End If
                        
            .Col = 12: .Text = RS01!접수금액 & ""                 '10
            .Col = 13: .Text = RS01!현금입금 + RS01!카드금액 & "" '11
            .Col = 14: .Text = RS01!현금입금 & ""                 '12
            .Col = 15: .Text = RS01!카드금액 & ""                 '13
            .Col = 16: .Text = RS01!카드건수 & ""                 '14
            .Col = 17: .Text = RS01!쿠폰금액 & ""                 '15
            .Col = 18: .Text = RS01!쿠폰건수 & ""                 '16
            .Col = 19: .Text = RS01!발생마일리지 & ""             '17
            .Col = 20: .Text = RS01!사용마일리지 & ""             '18
            .Col = 21: .Text = RS01!삭제마일리지 & ""             '19
            .Col = 22: .Text = RS01!반품환불금액 & ""             '20
            .Col = 23: .Text = RS01!반품환불건수 & ""             '21
            .Col = 24: .Text = RS01!세탁환불금액 & ""             '22
            .Col = 25: .Text = RS01!세탁환불건수 & ""             '23
            .Col = 26: .Text = RS01!재세탁수량 & ""               '24
            .Col = 27: .Text = RS01!수선금액 & ""                 '25
            .Col = 28: .Text = RS01!수선수량 & ""                 '26
                        
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        ' 합계 출력
        Dim nCol As Long
        For nCol = 3 To .MaxCols
            Select Case nCol
                Case 3: Call SpreadSum(spdView1, 2, nCol)
                Case Else: Call SpreadSum(spdView1, -1, nCol)
            End Select
        Next nCol

'        If .MaxRows > 0 Then
'            .MaxRows = .MaxRows + 1
'            .Row = .MaxRows
'
'            .Row = .Row
'            .Row2 = .Row
'            .Col = 1
'            .Col2 = .MaxCols
'            .BlockMode = True
'            .BackColor = &HC0FFC0
'            .BlockMode = False
'
'            .Col = 2: .Text = "합계"
'
'            .Col = 3: .Formula = "SUM(C1:C" & .MaxRows - 1 & ")"
'            .Col = 4: .Formula = "SUM(D1:D" & .MaxRows - 1 & ")"
'            .Col = 5: .Formula = "SUM(E1:E" & .MaxRows - 1 & ")"
'            .Col = 6: .Formula = "SUM(F1:F" & .MaxRows - 1 & ")"
'
'            .Col = 10: .Formula = "SUM(J1:J" & .MaxRows - 1 & ")"
'            .Col = 11: .Formula = "SUM(K1:K" & .MaxRows - 1 & ")"
'            .Col = 12: .Formula = "SUM(L1:L" & .MaxRows - 1 & ")"
'            .Col = 13: .Formula = "SUM(M1:M" & .MaxRows - 1 & ")"
'            .Col = 14: .Formula = "SUM(N1:N" & .MaxRows - 1 & ")"
'            .Col = 15: .Formula = "SUM(O1:O" & .MaxRows - 1 & ")"
'            .Col = 16: .Formula = "SUM(P1:P" & .MaxRows - 1 & ")"
'            .Col = 17: .Formula = "SUM(Q1:Q" & .MaxRows - 1 & ")"
'            .Col = 18: .Formula = "SUM(R1:R" & .MaxRows - 1 & ")"
'            .Col = 19: .Formula = "SUM(S1:S" & .MaxRows - 1 & ")"
'            .Col = 20: .Formula = "SUM(T1:T" & .MaxRows - 1 & ")"
'            .Col = 21: .Formula = "SUM(U1:U" & .MaxRows - 1 & ")"
'            .Col = 22: .Formula = "SUM(V1:V" & .MaxRows - 1 & ")"
'            .Col = 23: .Formula = "SUM(W1:W" & .MaxRows - 1 & ")"
'            .Col = 24: .Formula = "SUM(X1:X" & .MaxRows - 1 & ")"
'            .Col = 25: .Formula = "SUM(Y1:Y" & .MaxRows - 1 & ")"
'            .Col = 26: .Formula = "SUM(Z1:Z" & .MaxRows - 1 & ")"
'            .Col = 27: .Formula = "SUM(AA1:AA" & .MaxRows - 1 & ")"
'
'        End If
        
        .Redraw = True
    End With
End Sub

Private Sub DataSave()
'    Dim i As Integer
'
'    ReDim sValue(1)
'
'    sValue(0) = "0"
'    sValue(1) = Format(dtInput.Value, "YYYY-MM-DD")
'
'    Call ExecPro("SP_04001_01", sValue(), Err_Num, Err_Dec)
'
'    ReDim sValue(11)
'
'    For i = 1 To spdView.MaxRows
'        spdView.Row = i
'
'        sValue(0) = "0"
'        sValue(1) = Format(dtInput.Value, "YYYY-MM-DD")   ' 수금일자
'
'        spdView.Col = 1: sValue(2) = Mid(spdView.Text, 2, 3)                                                        ' 매장코드
'        spdView.Col = 2: sValue(3) = spdView.Value                                                                  ' 입고량
'        spdView.Col = 3: sValue(10) = spdView.Value                                                                 ' 출고량
'        spdView.Col = 4: sValue(4) = IIf(spdView.Text = "-", "", Mid(spdView.Text, 1, 1) & Mid(spdView.Text, 3, 3)) ' 시작택
'        spdView.Col = 5: sValue(5) = IIf(spdView.Text = "-", "", Mid(spdView.Text, 1, 1) & Mid(spdView.Text, 3, 3)) ' 종료택
'        spdView.Col = 6: sValue(6) = spdView.Value                                                                  ' 금액
'        spdView.Col = 8: sValue(7) = spdView.Value                                                                  ' 재세탁수량
'        spdView.Col = 9: sValue(8) = spdView.Value                                                                  ' 수선수량
'        spdView.Col = 10: sValue(9) = spdView.Value                                                                 ' 반품수량
'
'        spdView.Col = 12                                 ' UpdateChk
'        If spdView.Text = "수" Then
'            sValue(11) = "U"
'        Else
'            sValue(11) = ""
'        End If
'
'        If Int(sValue(6)) > 0 And Int(sValue(3)) <= 0 Then
'            MsgBox "[오류] " & "수금액이 있을경우 입고수량은 반드시 입력하셔야 합니다.", vbCritical, "확인"
'        Else
'            Call ExecPro("SP_04001_02", sValue(), Err_Num, Err_Dec)
'
'            If Err_Num <> 0 Then
'                MsgBox "[" & Err_Num & "] " & Err_Dec
'            End If
'
'        End If
'
'    Next i
'
'    ReDim sValue(1)
'
'    sValue(0) = Format(dtInput.Value, "YYYY-MM-DD")
'    sValue(1) = txtInput(6).Text
'
'    Call ExecPro("SP_04001_03", sValue(), Err_Num, Err_Dec)
End Sub

Private Sub DataCancel()
    Call Data_Display
End Sub

Private Sub DataPrint()

End Sub

Private Sub DataScreen()
'    Dim i As Integer, ii As Integer
'
'    ii = 0
'    For i = 1 To spdView.MaxRows
'        spdView.Row = i
'        spdView.Col = 11
'        If spdView.Value = True Then
'            ii = ii + 1
'        End If
'    Next i
'
'    If ii = 0 Then Exit Sub
'
'    ' 출력할 자료를 파일로 저장한다.
'    Call PrintDesc
'
'    ReDim PrtParam.Param(11)
'    With PrtParam
'        .Param(0) = "P_04001"
'        .Param(1) = "수금일자 : " & Format(dtInput.Value, "YYYY-MM-DD")
'        .Param(2) = "날씨: " & txtInput(6).Text
'
'        .Param(3) = "수금액 : " & txtInput(2).Text
'        .Param(4) = "단가 : ' " & txtInput(7).Text
'        .Param(5) = "일수금 : " & txtInput(8).Text
'        .Param(6) = "월수금 : " & txtInput(9).Text
'
'        .Param(7) = "입고량 : " & txtInput(0).Text
'        .Param(8) = "출고량 : " & txtInput(1).Text
'        .Param(9) = "재세탁 : " & txtInput(3).Text
'        .Param(10) = "수선 : " & txtInput(4).Text
'        .Param(11) = "반품 : " & txtInput(5).Text
'    End With
'
'    Load P_PRTSCREEN
'    P_PRTSCREEN.Show

End Sub


Private Sub PrintDesc()
    Dim i As Integer
    Dim TempText As String
    Dim TempFP As String
    Dim TempFile As String
    Dim hFile   As Integer
    
    Dim iDanga As Long
    
    On Error GoTo FileError:
    
    TempFP = GetIniStr("REPORT", "TempPath", "", m_iniFile)
    TempFile = TempFP & "\Temp.txt"
    
    hFile = FreeFile
    Open TempFile For Output As #hFile
    
    TempText = ""
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        spdView.Col = 11
        If spdView.Value = True Then
            TempText = Left(i & Space(3), 3) & Space(1) & "|"
            
            spdView.Col = 1
            TempText = TempText & LeftH(spdView.Text & Space(16), 16) & "|"
            
            spdView.Col = 13
            If spdView.Text = "월수금" Then
                TempText = TempText & "M" & "|"
            Else
                TempText = TempText & " " & "|"
            End If
            
            spdView.Col = 6
            TempText = TempText & RightH(Space(12) & spdView.Text, 12) & "|"
            iDanga = Val(spdView.Value)
            spdView.Col = 2
            TempText = TempText & RightH(Space(8) & spdView.Text, 8) & "|"
            If Val(spdView.Value) <> 0 Then iDanga = iDanga / Val(spdView.Value)
            
            spdView.Col = 3
            TempText = TempText & RightH(Space(8) & spdView.Text, 8) & "|"
            
            TempText = TempText & RightH(Space(12) & Format(iDanga, "#,##0"), 12) & "|"
            
            spdView.Col = 8
            TempText = TempText & RightH(Space(8) & spdView.Text, 8) & "|"
            spdView.Col = 9
            TempText = TempText & RightH(Space(8) & spdView.Text, 8) & "|"
            spdView.Col = 10
            TempText = TempText & RightH(Space(8) & spdView.Text, 8) & Space(2) & "|"
            spdView.Col = 4
            TempText = TempText & LeftH(spdView.Text & Space(5), 5) & " ~ "
            spdView.Col = 5
            TempText = TempText & LeftH(spdView.Text & Space(5), 5) & "|"
            
            If spdView.BackColor = &HD8FCFE Then
                TempText = TempText & " *"
            Else
                TempText = TempText & "  "
            End If
                        Print #hFile, TempText
        End If
    Next i
    
    Close #hFile
    Exit Sub
    
FileError:
    MsgBox Err.Description
    If Err.Number = 55 Then
        Resume Next
    End If
    Close #hFile
End Sub

Private Sub spdView_Change(ByVal Col As Long, ByVal Row As Long)
'    Select Case Col
'        Case 2, 6
'            spdView.Row = Row
'            spdView.Col = 12: spdView.Text = "수"
'            spdView.Col = -1: spdView.BackColor = vbYellow
'    End Select
End Sub

Private Sub spdView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        'PopupMenu P_00000.PopUp
    End If
End Sub
