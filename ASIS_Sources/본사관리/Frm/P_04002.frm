VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_04002 
   Caption         =   "지사 수금 입력"
   ClientHeight    =   11730
   ClientLeft      =   4200
   ClientTop       =   3600
   ClientWidth     =   16170
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_04002.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11730
   ScaleWidth      =   16170
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   11730
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16170
      _ExtentX        =   28522
      _ExtentY        =   20690
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_04002.frx":058A
      Begin Threed.SSPanel SSPanel 
         Height          =   2160
         Index           =   0
         Left            =   15
         TabIndex        =   15
         Top             =   1335
         Width           =   16140
         _ExtentX        =   28469
         _ExtentY        =   3810
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   0
            Left            =   1245
            TabIndex        =   34
            Top             =   1455
            Width           =   5145
         End
         Begin VB.ComboBox cboManager 
            Height          =   315
            Left            =   1245
            Style           =   2  '드롭다운 목록
            TabIndex        =   30
            Top             =   765
            Width           =   3420
         End
         Begin VB.ComboBox cboStore 
            Height          =   315
            Left            =   1245
            Style           =   2  '드롭다운 목록
            TabIndex        =   28
            Top             =   420
            Width           =   3420
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   1
            Left            =   1245
            TabIndex        =   16
            Top             =   1800
            Width           =   3420
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   3
            Left            =   1245
            TabIndex        =   17
            Top             =   75
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Format          =   64290816
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   18
            Top             =   75
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "수금일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   60
            TabIndex        =   19
            Top             =   1800
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "경리담당"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   3
            Left            =   60
            TabIndex        =   29
            Top             =   420
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
            Index           =   6
            Left            =   60
            TabIndex        =   31
            Top             =   765
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "배송기사"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin CSTextLibCtl.sidbEdit txtMoney 
            Height          =   315
            Left            =   1245
            TabIndex        =   32
            Top             =   1110
            Width           =   1140
            _Version        =   262145
            _ExtentX        =   2011
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Insert          =   0   'False
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   2
            StartText.y     =   2
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   16
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   7
            Left            =   60
            TabIndex        =   33
            Top             =   1110
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "입 금 액"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   8
            Left            =   60
            TabIndex        =   35
            Top             =   1455
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "비    고"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   8205
         Left            =   15
         TabIndex        =   1
         Top             =   3510
         Width           =   16140
         _Version        =   524288
         _ExtentX        =   28469
         _ExtentY        =   14473
         _StockProps     =   64
         BackColorStyle  =   1
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
         MaxCols         =   13
         MaxRows         =   35
         ScrollBars      =   0
         SpreadDesigner  =   "P_04002.frx":063C
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   2
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
         Caption         =   " 지사 수금 입력 (P_04002)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_04002.frx":0E63
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   1
         Left            =   8565
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
         PictureBackground=   "P_04002.frx":1065
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
            Picture         =   "P_04002.frx":1267
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
            Picture         =   "P_04002.frx":1801
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
            Picture         =   "P_04002.frx":1D9B
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
            Picture         =   "P_04002.frx":2335
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
            Picture         =   "P_04002.frx":28CF
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
            Picture         =   "P_04002.frx":2E69
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
            Picture         =   "P_04002.frx":3403
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
            Picture         =   "P_04002.frx":399D
         End
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   12
         Top             =   540
         Width           =   16140
         _ExtentX        =   28469
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Left            =   1245
            Style           =   2  '드롭다운 목록
            TabIndex        =   21
            Top             =   405
            Width           =   3420
         End
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1245
            Locked          =   -1  'True
            Style           =   2  '드롭다운 목록
            TabIndex        =   20
            Top             =   60
            Width           =   3420
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   2
            Left            =   6465
            TabIndex        =   13
            Top             =   420
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            Format          =   64290816
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   5280
            TabIndex        =   14
            Top             =   420
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "발행일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   6465
            TabIndex        =   22
            Top             =   60
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            Format          =   64290816
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   4
            Left            =   60
            TabIndex        =   23
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
            Index           =   5
            Left            =   5280
            TabIndex        =   24
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "수금일자"
            BorderWidth     =   0
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   10
            Left            =   60
            TabIndex        =   25
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
            Left            =   9645
            TabIndex        =   26
            Top             =   60
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            Format          =   64290816
            CurrentDate     =   36686
         End
         Begin XtremeSuiteControls.PushButton cmdRefresh 
            Height          =   330
            Left            =   4695
            TabIndex        =   36
            Top             =   390
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   582
            _StockProps     =   79
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_04002.frx":3F37
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
            Left            =   9330
            TabIndex        =   27
            Top             =   120
            Width           =   300
         End
      End
   End
End
Attribute VB_Name = "P_04002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub DataSave()
    ReDim sValue(7)
    
    sValue(0) = Mid(cboStore.Text, 2, 6)                          ' 1 가맹점코드
    sValue(1) = Format(dtInput(3).Value, "YYYY-MM-DD")            ' 2 입금일자
    sValue(2) = Mid(cboManager.Text, 2, 3)                        ' 3 배송기사코드
    sValue(3) = Trim(Mid(cboManager.Text, 6, Len(cboManager.Text) - 5)) ' 4 배송기사명
    sValue(4) = txtMoney.Value                                    ' 5 입금액
    sValue(5) = txtInput(0).Text & ""                             ' 6 비고
    sValue(6) = txtInput(1).Text & ""                             ' 7 경리담당자
    sValue(7) = Mid(cboOffice.Text, 2, 4)                         ' 8 지사코드
    
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub
        
        Call ExecProMaster("SP_04002_01", sValue(), Err_Num, Err_Dec)
    Else
        Call ExecPro("SP_04002_01", sValue(), Err_Num, Err_Dec)
    End If
    
    If Err_Num <> 0 Then
        MsgBox "[" & Err_Num & "] " & Err_Dec
    End If
    
    Call Data_Display
End Sub

Private Sub cboOffice_Click()
    If cboOffice.ListIndex < 0 Then Exit Sub
    
    '-----------------------------------------------------------------
    '
    '-----------------------------------------------------------------
    cboInput.Clear
    cboStore.Clear
    
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
    If cboStore.ListCount > 0 Then cboStore.ListIndex = 0
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: Call Data_Display   ' 조회
        Case 1: 'Call DataAdd        ' 신규
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

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(1).Enabled = True
    cmdBtn(2).Enabled = True
    cmdBtn(3).Enabled = True
    cmdBtn(5).Enabled = True
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

    '-------------------------------------------------------------------
    ' 기사
    '-------------------------------------------------------------------
    ReDim sValue(0)
    
    sValue(0) = "0"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_00002", sValue(), Err_Num, Err_Dec)

    cboManager.Clear
    
    Do Until RS01.EOF
        cboManager.AddItem "[" & RS01!기사코드 & "] " & RS01!기사명
        
        RS01.MoveNext
    Loop
    RS01.Close
    Set RS01 = Nothing
    
    If cboManager.ListCount > 0 Then cboManager.ListIndex = 0
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_04002_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i As Integer
    Dim j As Integer
    
    ReDim sValue(3)
    
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    
    If Mid(cboInput.Text, 2, 6) = "000000" Then
        sValue(1) = ""
    Else
        sValue(1) = Mid(cboInput.Text, 2, 6)
    End If
    
    sValue(2) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(3) = Format(dtInput(1).Value, "YYYY-MM-DD")
        
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_04002_00", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_04002_00", sValue(), Err_Num, Err_Dec)
    End If
    
    With spdView
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = RS01!입금일자 & ""
            .Col = 2:  .Text = RS01!가맹점코드 & ""
            .Col = 3:  .Text = RS01!가맹점명 & ""
            .Col = 4:  .Text = RS01!지사코드 & ""
            .Col = 5:  .Text = RS01!지사명 & ""
            .Col = 6:  .Text = RS01!배송기사코드 & ""
            .Col = 7:  .Text = RS01!배송기사명 & ""
            
            .Col = 8:  .Text = RS01!입금액 & ""
            .Col = 9:  .Text = RS01!비고 & ""
            .Col = 10:  .Text = RS01!입금확정 & ""
            .Col = 11:  .Text = RS01!확정일자 & ""
            .Col = 12: .Text = RS01!경리담당자 & ""
            .Col = 13: .Text = "0"
                        
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

Private Sub DataScreen()
'    Dim i As Integer
'    Dim j As Integer
'
'    Dim ReportFP As String
'    Dim ReportFile As String
'    Dim sData As String
'
'    Dim AgencySL As String
'    Dim iCnt As Integer
'
'    AgencySL = "({SP_04002_00;1.대리점명} = '  ' "
'
'    For i = 1 To spdView.MaxRows
'        spdView.Row = i
'        spdView.Col = 4
'
'        If spdView.Value = True Then
'            spdView.Col = 1
'            AgencySL = AgencySL & " Or {SP_04002_00;1.대리점명} = '" & spdView.Value & "' "
'        End If
'    Next i
'
'    AgencySL = AgencySL & ")"
'
'    ReportFP = GetIniStr("REPORT", "FilePath", "", m_iniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    Dim ii As Integer
'    For ii = 0 To 30
'        P_00000.crPrint.Formulas(ii) = ""
'    Next
'
'    P_00000.crPrint.StoredProcParam(0) = "0"
'    P_00000.crPrint.StoredProcParam(1) = Format(dtInput(0).Value, "yyyymmdd")
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'
'    P_00000.crPrint.Formulas(0) = "수금일자 = '" & Format(dtInput(0).Value, "yyyymmdd") & "'"
'    P_00000.crPrint.Formulas(1) = "발행일자 = '" & Format(dtInput(1).Value, "yyyymmdd") & "'"
'    P_00000.crPrint.Formulas(2) = "경리 = '" & txtInput(0).Text & "'"
'    P_00000.crPrint.Formulas(3) = "담당 = '" & txtInput(1).Text & "'"
'
'    P_00000.crPrint.SelectionFormula = AgencySL
'
'    Call ReportPrint(ReportFile, "2")
End Sub

Private Sub DataPrint()
'    Dim i As Integer
'    Dim j As Integer
'
'    Dim ReportFP As String
'    Dim ReportFile As String
'    Dim sData As String
'
'    Dim AgencySL As String
'    Dim iCnt As Integer
'
'    AgencySL = "({SP_04002_00;1.대리점명} = '  ' "
'
'    For i = 1 To spdView.MaxRows
'        spdView.Row = i
'        spdView.Col = 4
'
'        If spdView.Value = True Then
'            spdView.Col = 1
'            AgencySL = AgencySL & " Or {SP_04002_00;1.대리점명} = '" & spdView.Value & "' "
'        End If
'    Next i
'
'    AgencySL = AgencySL & ")"
'
'    ReportFP = GetIniStr("REPORT", "FilePath", "", m_iniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    P_00000.crPrint.StoredProcParam(0) = "0"
'    P_00000.crPrint.StoredProcParam(1) = Format(dtInput(0).Value, "yyyymmdd")
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'
'    Dim ii As Integer
'    For ii = 0 To 30
'        P_00000.crPrint.Formulas(ii) = ""
'    Next
'
'    P_00000.crPrint.Formulas(0) = "수금일자 = '" & Format(dtInput(0).Value, "yyyymmdd") & "'"
'    P_00000.crPrint.Formulas(1) = "발행일자 = '" & Format(dtInput(1).Value, "yyyymmdd") & "'"
'    P_00000.crPrint.Formulas(2) = "경리 = '" & txtInput(0).Text & "'"
'    P_00000.crPrint.Formulas(3) = "담당 = '" & txtInput(1).Text & "'"
'
'    P_00000.crPrint.SelectionFormula = AgencySL
'
'    Call ReportPrint(ReportFile, "1")
End Sub

Private Sub spdView_Click(ByVal Col As Long, ByVal Row As Long)
    Dim i As Integer
    
    If Row <= 0 Then Exit Sub
    
    With spdView
        .Row = Row
        
        .Col = 1: dtInput(3).Value = Format(.Text, "YYYY-MM-DD")
        
        .Col = 2:
        If Trim(.Text) <> "" Then
            For i = 0 To cboStore.ListCount - 1
                If Trim(.Text) = Mid(cboStore.List(i), 2, 6) Then
                    cboStore.ListIndex = i
                    Exit For
                End If
            Next i
        Else
            cboStore.ListIndex = -1
        End If
        
        .Col = 6:
        If Trim(.Text) <> "" Then
            For i = 0 To cboManager.ListCount - 1
                If Trim(.Text) = Mid(cboManager.List(i), 2, 3) Then
                    cboManager.ListIndex = i
                    Exit For
                End If
            Next i
        Else
            cboManager.ListIndex = -1
        End If
        
        .Col = 8: txtMoney.Value = .Value
        .Col = 9: txtInput(0).Text = .Text & ""
        .Col = 12: txtInput(1).Text = .Text & ""
    End With
End Sub
