VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_03019 
   Caption         =   "반품요청"
   ClientHeight    =   9435
   ClientLeft      =   4560
   ClientTop       =   3675
   ClientWidth     =   17175
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_03019.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9435
   ScaleWidth      =   17175
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   9435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17175
      _ExtentX        =   30295
      _ExtentY        =   16642
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_03019.frx":058A
      Begin Threed.SSPanel SSPanel 
         Height          =   405
         Index           =   0
         Left            =   15
         TabIndex        =   1
         Top             =   9015
         Width           =   17145
         _ExtentX        =   30242
         _ExtentY        =   714
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   45
            TabIndex        =   2
            Top             =   45
            Visible         =   0   'False
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "총  점  수"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   3
            Left            =   3045
            TabIndex        =   3
            Top             =   45
            Visible         =   0   'False
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "금    액"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   330
            Index           =   0
            Left            =   1230
            TabIndex        =   17
            Top             =   45
            Visible         =   0   'False
            Width           =   945
            _Version        =   262145
            _ExtentX        =   1667
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
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   2
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
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   330
            Index           =   1
            Left            =   4230
            TabIndex        =   18
            Top             =   30
            Visible         =   0   'False
            Width           =   1230
            _Version        =   262145
            _ExtentX        =   2170
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
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   2
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
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   4
         Top             =   540
         Width           =   17145
         _ExtentX        =   30242
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1245
            Locked          =   -1  'True
            TabIndex        =   15
            Text            =   "cboOffice"
            Top             =   60
            Width           =   3015
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   10
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "지    사"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   8
            Left            =   15000
            TabIndex        =   21
            Top             =   90
            Width           =   2115
            _Version        =   851970
            _ExtentX        =   3731
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 반품요청 처리"
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
            Picture         =   "P_03019.frx":065C
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   5
         Top             =   15
         Width           =   9540
         _ExtentX        =   16828
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
         PictureBackground=   "P_03019.frx":16EE
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   1
         Left            =   9570
         TabIndex        =   6
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
         PictureBackground=   "P_03019.frx":18F0
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   7
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
            Picture         =   "P_03019.frx":1AF2
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   8
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
            Picture         =   "P_03019.frx":208C
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   9
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
            Picture         =   "P_03019.frx":2626
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   10
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
            Picture         =   "P_03019.frx":2BC0
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   11
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
            Picture         =   "P_03019.frx":315A
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   12
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
            Picture         =   "P_03019.frx":36F4
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   13
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
            Picture         =   "P_03019.frx":3C8E
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   14
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
            Picture         =   "P_03019.frx":4228
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   7665
         Index           =   1
         Left            =   6000
         TabIndex        =   19
         Top             =   1335
         Width           =   11160
         _Version        =   524288
         _ExtentX        =   19685
         _ExtentY        =   13520
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
         MaxCols         =   13
         ScrollBars      =   2
         SpreadDesigner  =   "P_03019.frx":47C2
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   7665
         Index           =   0
         Left            =   15
         TabIndex        =   20
         Top             =   1335
         Width           =   5970
         _Version        =   524288
         _ExtentX        =   10530
         _ExtentY        =   13520
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
         OperationMode   =   3
         ScrollBars      =   2
         SelectBlockOptions=   2
         SpreadDesigner  =   "P_03019.frx":5015
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_03019"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cboInput_Change()

End Sub

Private Sub cboOffice_Click()
    If cboOffice.ListIndex < 0 Then Exit Sub
    
    
    Call Data_Display
'    '-----------------------------------------------------------------
'    '
'    '-----------------------------------------------------------------
'    cboInput.Clear
'
'    ReDim sValue(2)
'
'    sValue(0) = Mid(cboOffice.Text, 2, 4)
'    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
'    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
'
'    Set RS01 = New ADODB.Recordset
'    Set RS01 = ExecPro("SP_01001_00", sValue(), Err_Num, Err_Dec)
'
'    cboInput.AddItem "[000000] 전체"
'
'    Do Until RS01.EOF
'        'If (RS01!종료일자 = "") Or (RS01!종료일자 = "2099-12-31") Then '가맹점이 현 지사에서 관리중...
'            cboInput.AddItem "[" & RS01!가맹점코드 & "] " & RS01!가맹점명
'        'End If
'
'        RS01.MoveNext
'    Loop
'    RS01.Close
'    Set RS01 = Nothing
'
'    If cboInput.ListCount > 0 Then cboInput.ListIndex = 0
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
        Case 6: 'Call Export_Excel(P_00000.cdgExcel, spdView)      ' 엑셀
        Case 7: Unload Me           ' 종료
        Case 8: Call DataSave       ' 저장
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
    
    With spdView(0)
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
    With spdView(1)
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
    
    '
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
    
''    Call GoodsComboAdd(cboCloth)
    

        
''    If P_03009_Flag = False Then
''        Call AgencyComboAdd(cboInput(0))
''        Call GoodsComboAdd(cboInput(1))
''
''        dtInput(0).Value = Date
''        dtInput(1).Value = Date
''
''        P_03009_Flag = True
''    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Resize()
cmdBtn(8).Left = Me.Width - cmdBtn(8).Width - 200
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_03009_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    txtNum(0).Value = 0
    txtNum(1).Value = 0
    
    
    If Mid(cboOffice.Text, 2, 4) = "" Then Exit Sub
    
    ReDim sValue(3)
    
    sValue(0) = IIf(Mid(cboOffice.Text, 2, 4) = "000000", "%", Mid(cboOffice.Text, 2, 4))

    
    Screen.MousePointer = vbHourglass
    
    If HeadOffice = MASTER_OFFICE_CODE Then
        
        If Mid(cboOffice.Text, 2, 4) = "0000" Then
            If DBOpen_Master(MASTER_OFFICE_CODE) = False Then Exit Sub
        Else
            If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub
        End If

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_03009_03", sValue(), Err_Num, Err_Dec)
    End If
    'Else
        Dim SSQL As String
        SSQL = SSQL & " SELECT "
        SSQL = SSQL & "  가맹점코드, "
        SSQL = SSQL & "  (SELECT 가맹점명 FROM LAUNDRY1000..TB_가맹점 WHERE TB_가맹점.가맹점코드 = tb_입출고.가맹점코드) as 가맹점명, "
        SSQL = SSQL & "  count(*) as 건수, "
        SSQL = SSQL & "  sum(금액) as 금액 "
        SSQL = SSQL & " FROM LAUNDRY" & Mid(cboOffice.Text, 2, 4) & "..tb_입출고 "
        SSQL = SSQL & " WHERE 지사출고상태 = '3' and 접수일자 > convert(varchar(10),DATEADD(mm,-6,getdate()),121) and 접수일자 <= convert(varchar(10),getdate(),121)"
        SSQL = SSQL & " GROUP BY 가맹점코드"
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecQuery(SSQL, Err_Num, Err_Dec)
    'End If
        spdView(1).MaxRows = 0
    With spdView(0)
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = RS01!가맹점코드 & ""
            .Col = 2:  .Text = RS01!가맹점명 & ""
            .Col = 3:  .Text = RS01!건수 & ""
            .Col = 4:  .Text = RS01!금액 & ""
            
    
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        Call SpreadSum(spdView(0), 2, 3)
        Call SpreadSum(spdView(0), -1, 4)
        
        If .MaxRows >= 2 Then Call .SetText(1, .MaxRows, "000000")
        
        
    End With
        
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrRtn:
    
    Screen.MousePointer = vbDefault
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub


Private Sub Data_DisplayStore(sStoreCode As String)
    On Error GoTo ErrRtn

    txtNum(0).Value = 0
    txtNum(1).Value = 0
    
    ReDim sValue(2)
    
    Screen.MousePointer = vbHourglass
    
    sValue(2) = sStoreCode
    
    If Mid(cboOffice.Text, 2, 4) = "" Then Exit Sub
    
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_03009_01", sValue(), Err_Num, Err_Dec)
    '
    'Else
    End If
        Dim Query As String
        Query = "SELECT * from LAUNDRY" & Mid(cboOffice.Text, 2, 4) & "..tb_입출고 where 지사코드 = '" & Mid(cboOffice.Text, 2, 4) & "' and 가맹점코드 = '" & sStoreCode & "'  and 접수일자 > convert(varchar(10),DATEADD(mm,-6,getdate()),121) and 접수일자 <= convert(varchar(10),getdate(),121) and 지사출고상태 = '3'"
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecQuery(Query, Err_Num, Err_Dec)
    'End If
        
    With spdView(1)
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 2:  .Text = Format(RS01!택번호, "000-00-0000") & "" ' 2
            .Col = 3:  .Text = RS01!접수일자 & ""                      ' 1
            .Col = 4:  .Text = RS01!지사출고일자 & ""                  ' 3
            '.Col = 4:  .Text = RS01!출고차수 & ""                      ' 4
            .Col = 6:  .Text = RS01!의류코드 & ""                      ' 5
            .Col = 7:  .Text = RS01!의류명 & ""                        ' 6
            .Col = 8:  .Text = RS01!색상 & ""                          ' 7
            .Col = 9:  .Text = RS01!무늬 & ""                          ' 8
            .Col = 10:  .Text = RS01!내용 & ""                          ' 9
            .Col = 11: .Text = RS01!금액 & ""                          '10
            .Col = 12: .Text = RS01!상표 & ""                          '11
            .Col = 13: .Text = RS01!오점내용 & ""                      '12
            
            txtNum(0).Value = txtNum(0).Value + 1
            txtNum(1).Value = txtNum(1).Value + RS01!금액
    
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
    End With
        
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrRtn:
    
    Screen.MousePointer = vbDefault
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub
Private Sub DataPrint()

End Sub

Private Sub DataScreen()
'    Dim ReportFP As String
'    Dim ReportFile As String
'
'    ReportFP = GetIniStr("REPORT", "FilePath", "", m_iniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    Call PrintDesc
'
'    Dim i As Integer
'    For i = 0 To 30
'        P_00000.crPrint.Formulas(i) = ""
'    Next
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'    P_00000.crPrint.Formulas(0) = "검색일자 = '" & Format(dtInput(0).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(1) = "검색일자2 = '" & Format(dtInput(1).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(2) = "대리점명 = '" & cboInput.Text & "'"
'    P_00000.crPrint.Formulas(3) = "품목명 = '" & cboCloth.Text & "'"
'    P_00000.crPrint.Formulas(4) = "총점수 = '" & txtInput(0).Text & "'"
'    P_00000.crPrint.Formulas(5) = "금액 = '" & txtInput(1).Text & "'"
'
'    Call ReportPrint(ReportFile, "2")
End Sub
'
'Private Sub PrintDesc()
'    Dim i As Integer
'
'    Dim TempText As String
'    Dim TempFP As String
'    Dim TempFile As String
'
'    TempFP = GetIniStr("REPORT", "TempPath", "", m_iniFile)
'    TempFile = TempFP & "\Temp.txt"
'
'    Open TempFile For Output As #1
'
'    TempText = ""
'
'    For i = 1 To spdView(1).MaxRows - 1
'        spdView.Row = i
'
'        spdView.Col = 1:  TempText = LeftH(spdView.Text & Space(20), 20)
'        spdView.Col = 2:  TempText = TempText & LeftH(spdView.Text & Space(11), 11)
'        spdView.Col = 3:  TempText = TempText & LeftH(spdView.Text & Space(11), 11)
'        spdView.Col = 4:  TempText = TempText & LeftH(spdView.Text & Space(6), 6)
'        spdView.Col = 5:  TempText = TempText & LeftH(spdView.Text & Space(9), 9)
'        spdView.Col = 6:  TempText = TempText & LeftH(spdView.Text & Space(16), 16)
'        spdView.Col = 7:  TempText = TempText & LeftH(spdView.Text & Space(6), 6)
'        spdView.Col = 8:  TempText = TempText & RightH(Space(9) & spdView.Text, 9)
'        spdView.Col = 9:  TempText = TempText & LeftH(spdView.Text & Space(8), 8)
'        spdView.Col = 10: TempText = TempText & LeftH(spdView.Text & Space(8), 8)
'
'        Print #1, TempText
'    Next i
'
'    Close #1
'End Sub


Public Sub DataSave()
    Dim i As Integer
    
'    If HeadOffice = MASTER_OFFICE_CODE Then
'        If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub
'    End If

    With spdView(1)
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            If .Value = True Then
                ReDim sValue(5)
                
                .Col = 4:    sValue(0) = .Text   '반품일자
                .Col = 5:    sValue(1) = .Text   '차수
                             sValue(2) = .Tag    '가맹점코드
                .Col = 2:    sValue(3) = Replace(.Text, "-", "") '택번호
                .Col = 3:    sValue(4) = .Text   '접수일자
                .Col = 13:   sValue(5) = .Text   '반품사유
                Dim SSQL As String
                SSQL = ""
                SSQL = SSQL & "INSERT INTO LAUNDRY" & Mid(cboOffice.Text, 2, 4) & "..SCANOUTPUT_LOG_TB (SCAN_DATE, PDA_NO, STORE_CD, TAG_NO, SCAN_FLAG1, SCAN_FLAG2, OUT_DATE, OUT_COUNT) "
                SSQL = SSQL & "SELECT CONVERT(CHAR(19), getdate(), 20),'00',가맹점코드, 택번호, '1', '2', CONVERT(CHAR(19), getdate(), 20), '0' "
                SSQL = SSQL & "FROM LAUNDRY" & Mid(cboOffice.Text, 2, 4) & "..tb_입출고 "
                SSQL = SSQL & "WHERE "
                SSQL = SSQL & "     가맹점코드   = '" & sValue(2) & "'"
                SSQL = SSQL & "AND  택번호       = '" & sValue(3) & "'"
                SSQL = SSQL & "AND  지사출고상태 = '3'  "
                
                If HeadOffice = MASTER_OFFICE_CODE Then
                    Set RS01 = New ADODB.Recordset
                    'Set RS01 = ExecProMaster("SP_03009_02", sValue(), Err_Num, Err_Dec)
                Else
                    Set RS01 = New ADODB.Recordset
                    Set RS01 = ExecQuery(SSQL, Err_Num, Err_Dec)
                End If
                SSQL = ""
                SSQL = SSQL & "UPDATE"
                SSQL = SSQL & "     LAUNDRY" & Mid(cboOffice.Text, 2, 4) & "..tb_입출고 "
                SSQL = SSQL & "SET "
                SSQL = SSQL & " 가맹점출고일자 = '', "
                SSQL = SSQL & " 지사출고상태 = '2'"
                SSQL = SSQL & "WHERE "
                SSQL = SSQL & "     가맹점코드   = '" & sValue(2) & "'"
                SSQL = SSQL & "AND  택번호       = '" & sValue(3) & "'"
                SSQL = SSQL & "AND  지사출고상태 = '3'  "
'                If HeadOffice = MASTER_OFFICE_CODE Then
'                    Set RS01 = New ADODB.Recordset
'                    'Set RS01 = ExecProMaster("SP_03009_02", sValue(), Err_Num, Err_Dec)
'                Else
                    Set RS01 = New ADODB.Recordset
                    Set RS01 = ExecQuery(SSQL, Err_Num, Err_Dec)
'                End If
            End If
        
        Next i
    End With
    
    If Err_Num = 0 Then
        MsgBox "해당되는 데이터가 정상적으로 저장이 되었습니다.", vbInformation
        Exit Sub
    End If
End Sub


Private Sub spdView_Change(Index As Integer, ByVal Col As Long, ByVal Row As Long)
    If Index = 1 Then spdView(Index).SetText 1, Row, "1"
End Sub

Private Sub spdView_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)

    If Index = 0 Then
        Dim vText   As Variant
        
        spdView(Index).GetText 1, Row, vText
        
        If Len(vText) = 6 Then
            spdView(1).Tag = CStr(vText)
            
            Call Data_DisplayStore(spdView(1).Tag)
        End If
        
    End If
End Sub

