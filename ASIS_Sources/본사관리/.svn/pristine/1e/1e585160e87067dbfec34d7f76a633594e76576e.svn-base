VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_05016 
   Caption         =   "재고 조사 현황"
   ClientHeight    =   10185
   ClientLeft      =   3255
   ClientTop       =   3450
   ClientWidth     =   16380
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_05016.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10185
   ScaleWidth      =   16380
   WindowState     =   2  '최대화
   Begin Threed.SSPanel pnlProg 
      Height          =   1215
      Left            =   4485
      TabIndex        =   18
      Top             =   1710
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
      Picture         =   "P_05016.frx":058A
      Caption         =   " "
      BorderWidth     =   0
      BevelOuter      =   1
      BevelInner      =   2
      Alignment       =   4
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   10185
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16380
      _ExtentX        =   28893
      _ExtentY        =   17965
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_05016.frx":3555
      Begin Threed.SSPanel panInput 
         Height          =   750
         Left            =   15
         TabIndex        =   1
         Top             =   540
         Width           =   16350
         _ExtentX        =   28840
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin Threed.SSPanel SSPanel1 
            Height          =   405
            Left            =   6930
            TabIndex        =   21
            Top             =   360
            Width           =   6045
            _ExtentX        =   10663
            _ExtentY        =   714
            _Version        =   262144
            BackColor       =   12648384
            Caption         =   "재고 조사 이전에 고객 출고 처리가 된것"
            BevelOuter      =   0
            BevelInner      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   915
            Locked          =   -1  'True
            Style           =   2  '드롭다운 목록
            TabIndex        =   12
            Top             =   45
            Width           =   2850
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   915
            TabIndex        =   13
            Top             =   405
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            Format          =   68157440
            CurrentDate     =   36686
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   4095
            TabIndex        =   14
            Top             =   405
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            Format          =   68157440
            CurrentDate     =   36686
         End
         Begin VB.Label lblTitle 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "조사일자:"
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
            Index           =   0
            Left            =   0
            TabIndex        =   17
            Top             =   465
            Width           =   885
         End
         Begin VB.Label lblTitle 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "지 사 명:"
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
            Index           =   24
            Left            =   0
            TabIndex        =   16
            Top             =   105
            Width           =   885
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
            Index           =   0
            Left            =   3780
            TabIndex        =   15
            Top             =   465
            Width           =   300
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   2
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
         Caption         =   " 재고 조사 현황 (P_05016)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_05016.frx":3607
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   0
         Left            =   8775
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
         PictureBackground=   "P_05016.frx":3809
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
            Picture         =   "P_05016.frx":3A0B
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
            Picture         =   "P_05016.frx":3FA5
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
            Picture         =   "P_05016.frx":453F
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
            Picture         =   "P_05016.frx":4AD9
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
            Picture         =   "P_05016.frx":5073
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
            Picture         =   "P_05016.frx":560D
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
            Picture         =   "P_05016.frx":5BA7
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
            Picture         =   "P_05016.frx":6141
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   8865
         Left            =   15
         TabIndex        =   19
         Top             =   1305
         Width           =   4425
         _Version        =   524288
         _ExtentX        =   7805
         _ExtentY        =   15637
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
         MaxCols         =   3
         MaxRows         =   35
         ScrollBars      =   2
         SpreadDesigner  =   "P_05016.frx":66DB
         UserResize      =   1
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdView1 
         Height          =   8865
         Left            =   4455
         TabIndex        =   20
         Top             =   1305
         Width           =   11910
         _Version        =   524288
         _ExtentX        =   21008
         _ExtentY        =   15637
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
         MaxCols         =   7
         MaxRows         =   35
         ScrollBars      =   2
         SpreadDesigner  =   "P_05016.frx":6C62
         UserResize      =   1
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_05016"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String
Dim sSpValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Dim 지사출고번호 As Integer

Private Sub SPR_Resize()

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

    Dim i As Integer
    
    ReDim sValue(2)
    
    spdView1.MaxRows = 0
    


    
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_05016_00", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_05016_00", sValue(), Err_Num, Err_Dec)
    End If
        
    With spdView
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = RS01!가맹점코드 & ""
            .Col = 2: .Text = RS01!가맹점명 & ""
            .Col = 3: .Text = RS01!재고수량 & ""
            
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

Private Sub cmdBtn_Click(Index As Integer)
    Dim iRow        As Integer
    Dim 가맹점코드  As String
    Dim Query       As String

    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0:
            Call Data_Display    ' 조회
        
        Case 1: 'Call DataAdd        ' 신규
        Case 2: 'Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
        Case 4: 'Call DataCancel     ' 취소
        Case 5: 'Call DataPrint      ' 인쇄
        Case 6: Call Export_Excel(P_00000.cdgExcel, spdView1)      ' 엑셀
        Case 7: Unload Me            ' 종료
    
    End Select
    
    Exit Sub
    
ErrRtn:
    Me.MousePointer = 0
    
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub
 

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(5).Enabled = False
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
        .OperationMode = OperationModeSingle
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
    
    Dim i As Integer
    
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
    
    Call SPR_Resize
    
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
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Resize()
    Call SPR_Resize
End Sub

Private Sub Data_Display2(가맹점코드 As String)
    Dim i As Integer
    
    On Error GoTo ErrRtn
    
    pnlProg.Visible = True
    DoEvents
 
    
    ReDim sValue(2)
    
    '------------------------------------------------------------
    ' 지사 입고처리 - SP_03015_01
    '------------------------------------------------------------
    sValue(0) = 가맹점코드 'Mid(cboInput.Text, 2, 6)
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
        
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_05016_01", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_05016_01", sValue(), Err_Num, Err_Dec)
    End If
            
    With spdView1
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = Format(RS01!택번호, "000-00-0000")
            .Col = 2: .Text = RS01!의류코드 & ""  '
            .Col = 3: .Text = RS01!의류명 & ""    '
            
            .Col = 4: .Text = RS01!접수일자 & "" '
            .Col = 5: .Text = RS01!지사출고일자 & "" '
            
            .Col = 6: .Text = RS01!고객출고일자 & "" '
            .Col = 7: .Text = RS01!SCAN_DATE & "" '
            
            .Col = -1
            .BackColor = vbWhite
            If (RS01!SCAN_DATE & "" > RS01!고객출고일자 & "") And Len(RS01!고객출고일자 & "") > 1 Then
                .Col = -1
                .BackColor = &HC0FFC0    'vbGreen
            End If
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
    
    pnlProg.Visible = False
    
    Exit Sub
    
ErrRtn:
    pnlProg.Visible = False
End Sub

Private Sub spdView_Click(ByVal Col As Long, ByVal Row As Long)
    Dim 가맹점코드 As String
    
    If Row <= 0 Then Exit Sub
    
    spdView.Enabled = False
    
    spdView.Row = Row
    spdView.Col = 1: 가맹점코드 = spdView.Text & ""
    
    Call Data_Display2(가맹점코드)
    
    spdView.Enabled = True
End Sub

 
