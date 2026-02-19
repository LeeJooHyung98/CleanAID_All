VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_03006 
   Caption         =   "미출고 현황"
   ClientHeight    =   10095
   ClientLeft      =   615
   ClientTop       =   1575
   ClientWidth     =   15795
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_03006.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10095
   ScaleWidth      =   15795
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   10095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15795
      _ExtentX        =   27861
      _ExtentY        =   17806
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_03006.frx":058A
      Begin Threed.SSPanel SSPanel 
         Height          =   405
         Index           =   0
         Left            =   15
         TabIndex        =   2
         Top             =   9675
         Width           =   15765
         _ExtentX        =   27808
         _ExtentY        =   714
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   45
            TabIndex        =   3
            Top             =   45
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "미출고수량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   330
            Index           =   0
            Left            =   1275
            TabIndex        =   21
            Top             =   45
            Width           =   1080
            _Version        =   262145
            _ExtentX        =   1905
            _ExtentY        =   582
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   3
            StartText.y     =   4
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   13
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   8325
         Left            =   15
         TabIndex        =   1
         Top             =   1335
         Width           =   15765
         _Version        =   524288
         _ExtentX        =   27808
         _ExtentY        =   14684
         _StockProps     =   64
         DisplayColHeaders=   0   'False
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
         MaxCols         =   15
         MaxRows         =   36
         ScrollBars      =   2
         SpreadDesigner  =   "P_03006.frx":063C
         UserResize      =   1
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   4
         Top             =   540
         Width           =   15765
         _ExtentX        =   27808
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin CSTextLibCtl.sitxEdit txtTAG 
            Height          =   315
            Left            =   5805
            TabIndex        =   20
            Top             =   60
            Width           =   1335
            _Version        =   262145
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   "___-__-____"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.76
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   "___-__-____"
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   "###-##-####"
            Justification   =   1
            CharacterTable  =   ""
            BorderStyle     =   0
            MaxLength       =   9
         End
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1245
            Locked          =   -1  'True
            Style           =   2  '드롭다운 목록
            TabIndex        =   18
            Top             =   60
            Width           =   2700
         End
         Begin VB.ComboBox cboInput 
            Height          =   315
            Left            =   1245
            Style           =   2  '드롭다운 목록
            TabIndex        =   5
            Top             =   405
            Width           =   2700
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
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
            Index           =   2
            Left            =   4620
            TabIndex        =   7
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "택 번 호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   10
            Left            =   60
            TabIndex        =   19
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "지    사"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   8
         Top             =   15
         Width           =   8160
         _ExtentX        =   14393
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
         PictureBackground=   "P_03006.frx":0ADB
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   1
         Left            =   8190
         TabIndex        =   9
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
         PictureBackground=   "P_03006.frx":0CDD
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   10
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
            Picture         =   "P_03006.frx":0EDF
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   11
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
            Picture         =   "P_03006.frx":1479
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   12
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
            Picture         =   "P_03006.frx":1A13
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   13
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
            Picture         =   "P_03006.frx":1FAD
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   14
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
            Picture         =   "P_03006.frx":2547
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   15
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
            Picture         =   "P_03006.frx":2AE1
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   16
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
            Picture         =   "P_03006.frx":307B
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   17
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
            Picture         =   "P_03006.frx":3615
         End
      End
   End
End
Attribute VB_Name = "P_03006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cboInput_Click()
    DoEvents
    
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
    sValue(1) = Format(Date, "YYYY-MM-DD")
    sValue(2) = Format(Date, "YYYY-MM-DD")
    
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

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
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
        .RowHeight(-1) = 16
        
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
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_03006_Flag = False
End Sub

Private Sub Data_Display()
    Dim iCnt As Long
    Dim iCol As Long
    Dim iRow As Long
    
    On Error GoTo ErrRtn
    
    txtNum(0).Value = 0
    
    ReDim sValue(1)
    
    sValue(0) = Mid(cboInput.Text, 2, 6)   '
    sValue(1) = txtTAG.RawData & ""        '
        
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_03006_00", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_03006_00", sValue(), Err_Num, Err_Dec)
    End If
            
    iRow = 0
    iCol = 0
    iCnt = 0
    
    With spdView
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            iCnt = iCnt + 1
            
            If iCol = 0 Or iCol > 15 Then
                .MaxRows = .MaxRows + 1
                
                iRow = iRow + 1
                iCol = 0
            End If
            
            iCol = iCol + 1
            
            .Row = iRow
            .Col = iCol: .Value = Format(Right(RS01!택번호, 6), "00-0000") & ""
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
    
    txtNum(0).Value = iCnt
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub DataPrint()

End Sub

Private Sub spdView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        'PopupMenu P_00000.PopUp
    End If
End Sub
