VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_07003 
   Caption         =   "외주 입고 등록"
   ClientHeight    =   9255
   ClientLeft      =   1620
   ClientTop       =   1920
   ClientWidth     =   15780
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   15780
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15780
      _ExtentX        =   27834
      _ExtentY        =   16325
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_07003.frx":0000
      Begin Threed.SSPanel SSPanel 
         Height          =   570
         Index           =   1
         Left            =   5370
         TabIndex        =   1
         Top             =   1740
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   1005
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   390
            Index           =   1
            Left            =   1800
            TabIndex        =   20
            Top             =   90
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   688
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   16777215
            Format          =   70975488
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   90
            TabIndex        =   21
            Top             =   120
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   556
            _Version        =   262144
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "지사 외주 처리일자:"
            BorderWidth     =   0
            BevelOuter      =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   345
            Index           =   1
            Left            =   4620
            TabIndex        =   25
            Top             =   90
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   262144
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   2
         Top             =   540
         Width           =   15750
         _ExtentX        =   27781
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1245
            Style           =   2  '드롭다운 목록
            TabIndex        =   3
            Top             =   60
            Width           =   2850
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1245
            TabIndex        =   4
            Top             =   405
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            Format          =   70975488
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   5
            Top             =   405
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "스켄일자"
            BorderWidth     =   0
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   7
            Left            =   60
            TabIndex        =   6
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "외주업체"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   7
         Top             =   15
         Width           =   8145
         _ExtentX        =   14367
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
         Caption         =   " 외주 입고 등록 (P_07003)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_07003.frx":00F2
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   0
         Left            =   8175
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
         PictureBackground=   "P_07003.frx":02F4
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
            Picture         =   "P_07003.frx":04F6
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
            Caption         =   "화면"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_07003.frx":0A90
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
            Picture         =   "P_07003.frx":102A
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
            Picture         =   "P_07003.frx":15C4
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
            Picture         =   "P_07003.frx":1B5E
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
            Picture         =   "P_07003.frx":20F8
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
            Picture         =   "P_07003.frx":2692
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
            Picture         =   "P_07003.frx":2C2C
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   7905
         Left            =   15
         TabIndex        =   17
         Top             =   1335
         Width           =   5340
         _Version        =   524288
         _ExtentX        =   9419
         _ExtentY        =   13944
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
         SpreadDesigner  =   "P_07003.frx":31C6
         UserResize      =   1
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   390
         Index           =   2
         Left            =   5370
         TabIndex        =   18
         Top             =   1335
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   688
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
         PictureBackground=   "P_07003.frx":37B2
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.ProgressBar ProgressBar 
            Height          =   270
            Left            =   5910
            TabIndex        =   19
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
      Begin XtremeSuiteControls.TabControl TabControl 
         Height          =   6915
         Left            =   5370
         TabIndex        =   22
         Top             =   2325
         Width           =   10395
         _Version        =   851970
         _ExtentX        =   18336
         _ExtentY        =   12197
         _StockProps     =   68
         Appearance      =   3
         Color           =   64
         PaintManager.BoldSelected=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         PaintManager.ButtonMargin=   "2,3,2,3"
         ItemCount       =   1
         Item(0).Caption =   " PDA 스캔 현황 "
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "TabControlPage(0)"
         Begin XtremeSuiteControls.TabControlPage TabControlPage 
            Height          =   6435
            Index           =   0
            Left            =   30
            TabIndex        =   23
            Top             =   450
            Width           =   10335
            _Version        =   851970
            _ExtentX        =   18230
            _ExtentY        =   11351
            _StockProps     =   1
            Page            =   0
            Begin FPSpreadADO.fpSpread spdViewScan 
               Height          =   8205
               Left            =   30
               TabIndex        =   24
               Top             =   630
               Width           =   10875
               _Version        =   524288
               _ExtentX        =   19182
               _ExtentY        =   14473
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
               MaxCols         =   14
               MaxRows         =   35
               ScrollBars      =   0
               SpreadDesigner  =   "P_07003.frx":3C14
               UserResize      =   1
               Appearance      =   1
               HighlightHeaders=   1
               HighlightStyle  =   1
            End
            Begin XtremeSuiteControls.PushButton cmdBtn 
               Height          =   450
               Index           =   9
               Left            =   120
               TabIndex        =   26
               Top             =   120
               Width           =   3105
               _Version        =   851970
               _ExtentX        =   5477
               _ExtentY        =   794
               _StockProps     =   79
               Caption         =   " PDA 스캔 - 외주 입고 등록"
               ForeColor       =   -2147483640
               BackColor       =   -2147483636
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
               Picture         =   "P_07003.frx":4287
            End
         End
      End
   End
End
Attribute VB_Name = "P_07003"
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
    ReDim sValue(2)
    Dim nCnt    As Long
    
    nCnt = 0
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    sValue(1) = Store.Code
    sValue(2) = Format(dtInput(0).Value, "YYYYMMDD")
    
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("[SP_M_07003_00]", sValue(), Err_Num, Err_Dec)
    
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
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: Call Data_Display    ' 조회
        Case 1: 'Call DataAdd        ' 신규
        Case 2: 'Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
        Case 4: 'Call DataCancel     ' 취소
        Case 5: 'Call DataPrint      ' 인쇄
        Case 6: 'Call DataScreen     '
        Case 7: Unload Me            ' 종료
        Case 9: Call Data_Update
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

Private Sub Data_Update()
    Dim nRow      As Long
    Dim SSQL        As String
    
    On Error GoTo ERR_RTN
    
    If spdView.ActiveRow <= 0 Then Exit Sub
    
    '------------------------------------------------------------
    ' 먼저 스켄 자료를 백업한다. - SP_M_07002_02
    '------------------------------------------------------------
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_07003_03", sValue(), Err_Num, Err_Dec)
        
    MyCon.BeginTrans
    
    ReDim sValue(10)
    Do While Not RS01.EOF
        sValue(0) = RS01.Fields("OUTCD")
        sValue(1) = RS01.Fields("MASTERCD")
        sValue(2) = RS01.Fields("INDATE")
        sValue(3) = RS01.Fields("ICNT")
        sValue(4) = RS01.Fields("TAGNO")
        sValue(5) = RS01.Fields("PDANO")
        sValue(6) = RS01.Fields("IGCODE")
        sValue(7) = RS01.Fields("IGDNM")
        sValue(8) = RS01.Fields("IPRICE")
        sValue(9) = RS01.Fields("ADDTYPE")
        sValue(10) = RS01.Fields("SCANDT")
            
        Call ExecPro("SP_M_07003_02", sValue(), Err_Num, Err_Dec)
        If Err_Num <> 0 Then
            MyCon.RollbackTrans
            MsgBox Err_Dec
            Exit Sub
        End If
    
        RS01.MoveNext
    
    Loop

'    spdView.Row = spdView.ActiveRow
'    spdView.Col = 1: 가맹점코드 = spdView.Text & ""
    
    With spdViewScan
        ReDim sValue(7)
    
        For nRow = 1 To .MaxRows
            .Row = nRow
            
            .Col = 8: sValue(0) = .Text & ""        ' OUTCD
            .Col = 9: sValue(1) = .Text & ""        ' MASTERCD
            .Col = 1: sValue(2) = .Text & ""        ' TagNo
            .Col = 1: sValue(3) = panCaption(1).Tag & ""         ' ICNT
            .Col = 5: sValue(4) = .Text & ""        ' INDATE
            .Col = 6: sValue(5) = .Text & ""        ' SCANDT
            .Col = 7: sValue(6) = .Text & ""        ' PDANO
            .Col = 1: sValue(7) = Format(dtInput(1).Value, "yyyyMMdd") & ""       ' INACTIONDATE
            
            
            '------------------------------------------------------------
            ' 외주 입고 등록 - SP_M_07002_01
            '------------------------------------------------------------
            Set RS01 = New ADODB.Recordset
            Set RS01 = ExecPro("SP_M_07003_04", sValue(), Err_Num, Err_Dec)
            
            If Err_Num <> 0 Then
                MyCon.RollbackTrans
                MsgBox Err_Dec
                Exit Sub
            End If
        
        Next nRow
    End With
    MyCon.CommitTrans
    
    Call Data_Display2
    Call Data_Display
    
    MsgBox "저장 완료", vbInformation, "확인"
    
    Exit Sub

ERR_RTN:
    MyCon.RollbackTrans
    MsgBox Err.Description
    
End Sub

Private Sub dtInput_Change(Index As Integer)
    dtInput(Index).Enabled = False
    
    ReDim sValue(0)
    sValue(0) = Format(dtInput(Index).Value, "yyyyMMdd")
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_07003_05", sValue(), Err_Num, Err_Dec)
    
    If Not RS01.EOF Then
        panCaption(1).Tag = RS01.Fields("CNT") & ""
        panCaption(1).Caption = RS01.Fields("CNT") & " 회차 입고"
    End If
    
    Call Data_Display
    
    dtInput(Index).Enabled = True
End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(5).Enabled = True
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
    If P_07003_Flag = False Then
        Dim I As Integer
        dtInput(0).Value = Date
        dtInput(1).Value = Date

        '
        Call OrderComboAdd(cboOffice)
        
        With cboOffice
            For I = 0 To .ListCount - 1
                If Mid(.List(I), 2, 4) = HeadOffice Then
                    .ListIndex = I
                    
                    Exit For
                End If
            Next I
        End With
        
        P_07003_Flag = True
    End If

End Sub

Private Sub Form_Load()
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
    
    Dim I As Integer
    
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
    
        .Col = 8:   .ColHidden = True
        .Col = 9:   .ColHidden = True
        .Col = 10:   .ColHidden = True
        .Col = 11:   .ColHidden = True
        .Col = 12:   .ColHidden = True
        .Col = 13:   .ColHidden = True
        .Col = 14:   .ColHidden = True
    
    End With
    
    Call SPR_Resize
    
End Sub

Private Sub Form_Resize()
    Call SPR_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_07003_Flag = False
End Sub

Private Sub Data_Display2()
    On Error GoTo ErrRtn
    
    ReDim sValue(6)
    
    spdView.Row = spdView.ActiveRow
        
    spdView.Col = 1:        sValue(0) = spdView.Text
    spdView.Col = 4:        sValue(1) = spdView.Text + "%"
    sValue(2) = Store.Code
    sValue(3) = Mid(cboOffice.Text, 2, 4) + "%"
    sValue(4) = Format(dtInput(0).Value, "YYYYMMDD")
    sValue(5) = Format(DateAdd("d", -30, dtInput(0).Value), "YYYY-MM-DD")
    sValue(6) = Format(dtInput(0).Value, "YYYY-MM-DD")
    
    '------------------------------------------------------------
    ' 외주 출고 등록 - SP_M_07002_01
    '------------------------------------------------------------
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_07003_01", sValue(), Err_Num, Err_Dec)
    
    With spdViewScan
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = RS01!TagNo & ""    'KEY
            .Col = 2: .Text = RS01!의류코드 & ""  '
            .Col = 3: .Text = RS01!의류명 & ""    '
            .Col = 4: .Text = RS01!출고일자 & ""    '
            .Col = 5: .Text = RS01!INDATE & ""      'KEY
            .Col = 6: .Text = RS01!SCANDT & "" '
            .Col = 7: .Text = RS01!PDANO & ""    '
            .Col = 8: .Text = RS01!OUTCD & ""    'KEY
            .Col = 9: .Text = RS01!MASTERCD & ""    'KEY
            .Col = 10: .Text = RS01!iCnt & ""    'KEY
            .Col = 11: .Text = RS01!IGCODE & ""    '
            .Col = 12: .Text = RS01!IPRICE & ""    '
            .Col = 13: .Text = RS01!ADDTYPE & ""    '
            
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

