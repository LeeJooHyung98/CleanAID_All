VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_04001_C 
   Caption         =   "가맹점 매출현황 (가맹점 기준)"
   ClientHeight    =   9870
   ClientLeft      =   645
   ClientTop       =   2790
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
   Icon            =   "P_04001_C.frx":0000
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
      PaneTree        =   "P_04001_C.frx":058A
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
         Item(0).Caption =   " 가맹점별 매출현황 "
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "TabControlPage1"
         Item(1).Caption =   "가맹점별 매출집계 "
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "TabControlPage2"
         Begin XtremeSuiteControls.TabControlPage TabControlPage2 
            Height          =   8070
            Left            =   -69970
            TabIndex        =   22
            Top             =   420
            Visible         =   0   'False
            Width           =   16395
            _Version        =   851970
            _ExtentX        =   28919
            _ExtentY        =   14235
            _StockProps     =   1
            Page            =   1
            Begin FPSpreadADO.fpSpread spdView1 
               Height          =   5055
               Left            =   45
               TabIndex        =   24
               Top             =   45
               Width           =   16455
               _Version        =   524288
               _ExtentX        =   29025
               _ExtentY        =   8916
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
               MaxCols         =   31
               SpreadDesigner  =   "P_04001_C.frx":061C
               Appearance      =   1
               CellNoteIndicatorColor=   16761024
               HighlightHeaders=   1
               HighlightStyle  =   1
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage1 
            Height          =   8070
            Left            =   30
            TabIndex        =   21
            Top             =   420
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
               MaxCols         =   31
               SpreadDesigner  =   "P_04001_C.frx":1896
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
            Format          =   64225280
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
            Format          =   64225280
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
            Picture         =   "P_04001_C.frx":2AF4
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
         Caption         =   " 가맹점 매출현황 (가맹점 기준) (P_04001_C)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_04001_C.frx":308E
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
         PictureBackground=   "P_04001_C.frx":3290
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
            Picture         =   "P_04001_C.frx":3492
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
            Picture         =   "P_04001_C.frx":3A2C
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
            Picture         =   "P_04001_C.frx":3FC6
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
            Picture         =   "P_04001_C.frx":4560
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
            Picture         =   "P_04001_C.frx":4AFA
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
            Picture         =   "P_04001_C.frx":5094
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
            Picture         =   "P_04001_C.frx":562E
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
            Picture         =   "P_04001_C.frx":5BC8
         End
      End
   End
End
Attribute VB_Name = "P_04001_C"
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
        Case 2: 'Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
        Case 4: 'Call DataCancel     ' 취소
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
 

Private Sub dtInput_Change(Index As Integer)
'    Call Data_Display
End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    'cmdBtn(2).Enabled = True
    'cmdBtn(4).Enabled = True
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
        '.OperationMode = OperationModeSingle
        
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
        '.OperationMode = OperationModeSingle
        
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
    Set RS01 = ExecPro("SP_04001_C_00", sValue(), Err_Num, Err_Dec)
    
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
            
            If RS01!접수수량 = 0 Then
                .Col = 9: .Text = 0 & ""   '14
                .Col = 10: .Text = 0 & ""   '15
                .Col = 11: .Text = 0 & ""   '16
            Else
                .Col = 9: .Text = RS01!접수금액 / RS01!접수수량 & ""   '14
                .Col = 10: .Text = RS01!지사금액 / RS01!접수수량 & ""   '15
                .Col = 11: .Text = RS01!가맹점금액 / RS01!접수수량 & "" '16
            End If
            
            
            If Len(RS01!이전종료택번호) = 9 Then
                .Col = 12:  .Text = Format(RS01!이전종료택번호, "000-00-0000") & ""     ' 8
            Else
                .Col = 12:  .Text = RS01!이전종료택번호 & ""  ' 8
            End If
            
            If Len(RS01!시작택번호) = 9 Then
                .Col = 13:  .Text = Format(RS01!시작택번호, "000-00-0000") & ""         ' 9
            Else
                .Col = 13:  .Text = RS01!시작택번호 & ""      ' 9
            End If
            
            If Len(RS01!종료택번호) = 9 Then
                .Col = 14: .Text = Format(RS01!종료택번호, "000-00-0000") & ""         '10
            Else
                .Col = 14: .Text = RS01!종료택번호 & ""      '10
            End If
            
            .Col = 15: .Text = RS01!접수금액 & ""            '11
            .Col = 16: .Text = RS01!현금입금 + RS01!카드금액 & "" '12
            .Col = 17: .Text = RS01!현금입금 & ""            '13
            .Col = 18: .Text = RS01!카드금액 & ""            '14
            .Col = 19: .Text = RS01!카드건수 & ""            '15
            .Col = 20: .Text = RS01!쿠폰금액 & ""            '16
            .Col = 21: .Text = RS01!쿠폰건수 & ""            '17
            .Col = 22: .Text = RS01!발생마일리지 & ""        '18
            .Col = 23: .Text = RS01!사용마일리지 & ""        '19
            .Col = 24: .Text = RS01!삭제마일리지 & ""        '20
            .Col = 25: .Text = RS01!반품환불금액 & ""        '21
            .Col = 26: .Text = RS01!반품환불건수 & ""        '22
            .Col = 27: .Text = RS01!세탁환불금액 & ""        '23
            .Col = 28: .Text = RS01!세탁환불건수 & ""        '24
            .Col = 29: .Text = RS01!재세탁수량 & ""          '25
            .Col = 30: .Text = RS01!수선금액 & ""            '26
            .Col = 31: .Text = RS01!수선수량 & ""            '27
            
            If RS01!접수수량 > 0 Then
                If RS01!접수수량 <> (Replace(RS01!종료택번호, "_", "") - Replace(RS01!시작택번호, "_", "")) + 1 Then
                    .Col = -1: .BackColor = vbYellow
                End If
            End If
            
            If Trim(RS01!이전종료택번호) <> "" And Trim(RS01!시작택번호) <> "" Then
                If (Replace(RS01!이전종료택번호, "_", "") + 1) = Replace(RS01!시작택번호, "_", "") Then
                    .Col = -1: .BackColor = &HFFC0C0
                End If
            End If
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        If .MaxRows > 0 Then
            ' 합계 출력
            Dim nCol    As Long
            If Mid(cboInput.Text, 2, 6) <> "000000" Then
                For nCol = 4 To .MaxCols
                    Select Case nCol
                        Case 4: Call SpreadSum(spdView, 2, nCol)
                        Case Else: Call SpreadSum(spdView, -1, nCol)
                    End Select
                Next nCol
            End If
        End If
'
'        If Mid(cboInput.Text, 2, 6) <> "000000" Then
'            If .MaxRows > 0 Then
'                .MaxRows = .MaxRows + 1
'                .Row = .MaxRows
'
'                .Row = .Row
'                .Row2 = .Row
'                .Col = 1
'                .Col2 = .MaxCols
'                .BlockMode = True
'                .BackColor = &HC0FFC0
'                .BlockMode = False
'
'                .Col = 2: .Text = "합계"
'                .Col = 4: .Formula = "SUM(D1:D" & .MaxRows - 1 & ")"
'                .Col = 5: .Formula = "SUM(E1:E" & .MaxRows - 1 & ")"
'                .Col = 6: .Formula = "SUM(F1:F" & .MaxRows - 1 & ")"
'                .Col = 7: .Formula = "SUM(G1:G" & .MaxRows - 1 & ")"
'
'                .Col = 8: .Formula = "SUM(H1:H" & .MaxRows - 1 & ") / " & .MaxRows - 1
'                .Col = 9: .Formula = "SUM(I1:I" & .MaxRows - 1 & ") / " & .MaxRows - 1
'                .Col = 10: .Formula = "SUM(J1:J" & .MaxRows - 1 & ") / " & .MaxRows - 1
'
'                .Col = 14: .Formula = "SUM(N1:N" & .MaxRows - 1 & ")"
'                .Col = 15: .Formula = "SUM(O1:O" & .MaxRows - 1 & ")"
'                .Col = 16: .Formula = "SUM(P1:P" & .MaxRows - 1 & ")"
'                .Col = 17: .Formula = "SUM(Q1:Q" & .MaxRows - 1 & ")"
'                .Col = 18: .Formula = "SUM(R1:R" & .MaxRows - 1 & ")"
'                .Col = 19: .Formula = "SUM(S1:S" & .MaxRows - 1 & ")"
'                .Col = 20: .Formula = "SUM(T1:T" & .MaxRows - 1 & ")"
'                .Col = 21: .Formula = "SUM(U1:U" & .MaxRows - 1 & ")"
'                .Col = 22: .Formula = "SUM(V1:V" & .MaxRows - 1 & ")"
'                .Col = 23: .Formula = "SUM(W1:W" & .MaxRows - 1 & ")"
'                .Col = 24: .Formula = "SUM(X1:X" & .MaxRows - 1 & ")"
'                .Col = 25: .Formula = "SUM(Y1:Y" & .MaxRows - 1 & ")"
'                .Col = 26: .Formula = "SUM(Z1:Z" & .MaxRows - 1 & ")"
'                .Col = 27: .Formula = "SUM(AA1:AA" & .MaxRows - 1 & ")"
'                .Col = 28: .Formula = "SUM(AB1:AB" & .MaxRows - 1 & ")"
'                .Col = 29: .Formula = "SUM(AC1:AC" & .MaxRows - 1 & ")"
'                .Col = 30: .Formula = "SUM(AD1:AD" & .MaxRows - 1 & ")"
'            End If
'        End If
        
        .Redraw = True
    End With
    
    Call Data_Display2
        
    Exit Sub
    
ErrRtn:
    Resume
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
    Set RS01 = ExecPro("SP_04001_C_01", sValue(), Err_Num, Err_Dec)
    
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
            
            If RS01!접수수량 = 0 Then
                .Col = 9: .Text = 0 & ""   '14
                .Col = 10: .Text = 0 & ""   '15
                .Col = 11: .Text = 0 & ""   '16
            Else
                .Col = 9: .Text = RS01!접수금액 / RS01!접수수량 & ""   '14
                .Col = 10: .Text = RS01!지사금액 / RS01!접수수량 & ""   '15
                .Col = 11: .Text = RS01!가맹점금액 / RS01!접수수량 & "" '16
            End If
            
            If Len(RS01!이전종료택번호) = 9 Then
                .Col = 12:  .Text = Format(RS01!이전종료택번호, "000-00-0000") & ""     ' 8
            Else
                .Col = 12:  .Text = RS01!이전종료택번호 & ""  ' 8
            End If
            
            If Len(RS01!시작택번호) = 9 Then
                .Col = 13:  .Text = Format(RS01!시작택번호, "000-00-0000") & ""         ' 9
            Else
                .Col = 13:  .Text = RS01!시작택번호 & ""      ' 9
            End If
            
            If Len(RS01!종료택번호) = 9 Then
                .Col = 14: .Text = Format(RS01!종료택번호, "000-00-0000") & ""         '10
            Else
                .Col = 14: .Text = RS01!종료택번호 & ""      '10
            End If
            
            .Col = 15: .Text = RS01!접수금액 & ""                 '10
            .Col = 16: .Text = RS01!현금입금 + RS01!카드금액 & "" '11
            .Col = 17: .Text = RS01!현금입금 & ""                 '12
            .Col = 18: .Text = RS01!카드금액 & ""                 '13
            .Col = 19: .Text = RS01!카드건수 & ""                 '14
            .Col = 20: .Text = RS01!쿠폰금액 & ""                 '15
            .Col = 21: .Text = RS01!쿠폰건수 & ""                 '16
            .Col = 22: .Text = RS01!발생마일리지 & ""             '17
            .Col = 23: .Text = RS01!사용마일리지 & ""             '18
            .Col = 24: .Text = RS01!삭제마일리지 & ""             '19
            .Col = 25: .Text = RS01!반품환불금액 & ""             '20
            .Col = 26: .Text = RS01!반품환불건수 & ""             '21
            .Col = 27: .Text = RS01!세탁환불금액 & ""             '22
            .Col = 28: .Text = RS01!세탁환불건수 & ""             '23
            .Col = 29: .Text = RS01!재세탁수량 & ""               '24
            .Col = 30: .Text = RS01!수선금액 & ""                 '25
            .Col = 31: .Text = RS01!수선수량 & ""                 '26
                        
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        If .MaxRows > 0 Then
            Dim nCol    As Long
            For nCol = 3 To .MaxCols
                Select Case nCol
                    Case 3: Call SpreadSum(spdView1, 2, nCol)
                    Case Else: Call SpreadSum(spdView1, -1, nCol)
                End Select
            Next nCol
        End If
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
'            .Col = 7: .Formula = "SUM(G1:G" & .MaxRows - 1 & ")"
'
'            .Col = 8: .Formula = "SUM(H1:H" & .MaxRows - 1 & ") / " & .MaxRows - 1
'            .Col = 9: .Formula = "SUM(I1:I" & .MaxRows - 1 & ") / " & .MaxRows - 1
'            .Col = 10: .Formula = "SUM(J1:J" & .MaxRows - 1 & ") / " & .MaxRows - 1
'
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
'            .Col = 28: .Formula = "SUM(AB1:AB" & .MaxRows - 1 & ")"
'            .Col = 29: .Formula = "SUM(AC1:AC" & .MaxRows - 1 & ")"
'            .Col = 30: .Formula = "SUM(AD1:AD" & .MaxRows - 1 & ")"
'        End If
'
        .Redraw = True
    End With
End Sub

Private Sub DataSave()

End Sub

Private Sub DataCancel()
    Call Data_Display
End Sub

Private Sub DataPrint()

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

