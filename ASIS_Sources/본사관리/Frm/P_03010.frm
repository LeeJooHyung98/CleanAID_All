VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_03010 
   Caption         =   "오류 현황"
   ClientHeight    =   10440
   ClientLeft      =   -18345
   ClientTop       =   3075
   ClientWidth     =   16455
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_03010.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10440
   ScaleWidth      =   16455
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   10440
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   18415
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_03010.frx":058A
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   1
         Top             =   540
         Width           =   16425
         _ExtentX        =   28972
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1245
            Locked          =   -1  'True
            Style           =   2  '드롭다운 목록
            TabIndex        =   16
            Top             =   45
            Width           =   2805
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1245
            TabIndex        =   2
            Top             =   405
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Format          =   63045632
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   3
            Top             =   405
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "출고일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   4485
            TabIndex        =   4
            Top             =   405
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Format          =   63045632
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   7
            Left            =   60
            TabIndex        =   17
            Top             =   45
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "지    사"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin VB.Label Label1 
            Alignment       =   2  '가운데 맞춤
            Caption         =   "~"
            Height          =   255
            Left            =   4245
            TabIndex        =   5
            Top             =   465
            Width           =   255
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   6
         Top             =   15
         Width           =   8820
         _ExtentX        =   15558
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
         Caption         =   " 오류현황 (P_03010)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_03010.frx":063C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   1
         Left            =   8850
         TabIndex        =   7
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
         PictureBackground=   "P_03010.frx":083E
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   8
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
            Picture         =   "P_03010.frx":0A40
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   9
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
            Picture         =   "P_03010.frx":0FDA
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   10
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
            Picture         =   "P_03010.frx":1574
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   11
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
            Picture         =   "P_03010.frx":1B0E
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   12
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
            Picture         =   "P_03010.frx":20A8
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   13
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
            Picture         =   "P_03010.frx":2642
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   14
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
            Picture         =   "P_03010.frx":2BDC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   15
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
            Picture         =   "P_03010.frx":3176
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   9090
         Left            =   15
         TabIndex        =   18
         Top             =   1335
         Width           =   5190
         _Version        =   524288
         _ExtentX        =   9155
         _ExtentY        =   16034
         _StockProps     =   64
         BackColorStyle  =   1
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
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
         MaxRows         =   34
         ScrollBars      =   2
         SpreadDesigner  =   "P_03010.frx":3710
         UserResize      =   1
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdView1 
         Height          =   9090
         Left            =   5220
         TabIndex        =   19
         Top             =   1335
         Width           =   11220
         _Version        =   524288
         _ExtentX        =   19791
         _ExtentY        =   16034
         _StockProps     =   64
         BackColorStyle  =   1
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
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
         MaxCols         =   22
         Protect         =   0   'False
         SpreadDesigner  =   "P_03010.frx":3D69
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_03010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: Call Data_Display   ' 조회
        Case 1: 'Call DataAdd        ' 신규
        Case 2: 'Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
        Case 4: 'Call DataCancel     ' 취소
        Case 5:
            
            Dim 가맹점명 As String
            Dim iRow     As Integer
            
            For iRow = 1 To spdView.MaxRows
                spdView.Row = iRow
                spdView.Col = 4
                
                If spdView.Text = "1" Then
                    spdView.Col = 2: 가맹점명 = spdView.Text & ""
                    
                    Call DataPrint(가맹점명)      ' 인쇄
                End If
            Next iRow
            
        Case 6: Call Export_Excel(P_00000.cdgExcel, spdView)      ' 엑셀
        Case 7: Unload Me           ' 종료
    End Select
    
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
'    Call Data_Display
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
    
    Dim i As Integer
    
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
    
    With spdView1
        .MaxRows = 0
        .RowHeight(-1) = 14
                
        .Col = 2: .ColMerge = MergeAlways
        .Col = 3: .ColMerge = MergeRestricted
        
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

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim iTotal As Long

    ReDim sValue(2)
    
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_03010_00", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_03010_00", sValue(), Err_Num, Err_Dec)
    End If
        
    With spdView
        .MaxRows = 0
        
        Do While Not RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = RS01!가맹점코드 & ""
            .Col = 2: .Text = RS01!가맹점명 & ""
            .Col = 3: .Text = RS01!출고수량 & ""
            .Col = 4: .Text = "1"
            
            iTotal = iTotal + RS01!출고수량
            
            RS01.MoveNext
        Loop
    End With
    
    'txtNum(0).Value = iTotal
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Data_Display2(가맹점코드 As String)
    '----------------------------------------------------------------
    ' SP_03001_01
    '----------------------------------------------------------------
    ReDim sValue(2)
    
    Screen.MousePointer = vbHourglass
    
    sValue(0) = 가맹점코드
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_03010_01", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_03010_01", sValue(), Err_Num, Err_Dec)
    End If
    
    
    With spdView1
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = Format(RS01!택번호, "000-00-0000") & ""        ' 1
            .Col = 2:  .Text = RS01!지사출고일자 & ""       ' 2
            .Col = 3:  .Text = RS01!성명 & ""           ' 3
            .Col = 4:  .Text = RS01!전화번호 & ""       ' 4
            .Col = 5:  .Text = RS01!휴대전화 & ""       ' 5
            .Col = 6:  .Text = RS01!의류코드 & ""       ' 6
            .Col = 7:  .Text = RS01!의류명 & ""         ' 7
            .Col = 8:  .Text = RS01!색상 & ""           ' 8
            .Col = 9:  .Text = RS01!무늬 & ""           ' 9
            .Col = 10: .Text = RS01!내용 & ""           '10
            .Col = 11: .Text = RS01!상표 & ""           '11
            .Col = 12: .Text = RS01!금액 & ""           '12
            .Col = 13: .Text = RS01!접수일자 & ""       '13
            .Col = 14: .Text = RS01!가맹점출고일자 & "" '14
            .Col = 15: .Text = RS01!가맹점입고일자 & "" '15
            .Col = 16: .Text = RS01!고객출고일 & "" '15
            .Col = 17: .Text = RS01!부모택번호 & ""     '16
            .Col = 18: .Text = RS01!반품환불일자 & ""   '17
            .Col = 19: .Text = RS01!세탁환불일자 & ""   '18
            .Col = 20: .Text = RS01!판매취소일자 & ""   '19
            .Col = 21: .Text = RS01!환불사유 & ""       '20
            .Col = 22: .Text = RS01!오점내용 & ""       '21
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
            
            ' 색상을 적용 한다.
        Call chkColor_Setting(True)
    
    
        .Redraw = True
    End With
    
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub DataPrint(가맹점명 As String)
    On Error GoTo ErrRtn
    
    Dim XML         As String
    Dim i           As Integer
    Dim idx         As Integer
    Dim FileNumber
        
    If spdView1.MaxRows = 0 Then Exit Sub
    
    FileNumber = FreeFile
    
    Open App.Path & "\XML\지사출고현황.XML" For Output As #FileNumber
    
    Print #FileNumber, "<?xml version=""1.0"" encoding=""EUC-KR""?>"
    Print #FileNumber, "<root>"
    
          XML = "    <조건>"
    XML = XML & "        <가맹점>(" & Func_Replace(가맹점명) & ") 출고내역</가맹점>"
    XML = XML & "        <출고수량>출고수량 : " & spdView1.MaxRows & "</출고수량>"
    
    XML = XML & "        <출고일자>출고일자 : " & Format(dtInput(0).Value, "YYYY-MM-DD") & " ~ " & Format(dtInput(1).Value, "YYYY-MM-DD") & "</출고일자>"
    XML = XML & "   </조건>"
    Print #FileNumber, XML
    
    With spdView1
        idx = 0
        
        For i = 1 To .MaxRows
            .Row = i
            
            If idx = 0 Or idx = 7 Then
                If idx = 0 Then
                    XML = "    <Data>"
                Else
                    XML = XML & "   </Data>"
                    Print #FileNumber, XML
                    
                    XML = "    <Data>"
                End If
                
                idx = 0
            End If
            
            idx = idx + 1
            
            .Col = 1: XML = XML & "        <택번호" & idx & ">" & Right(.Text, 7) & "</택번호" & idx & ">"
        Next i
        
        If idx = 7 Then
            XML = XML & "   </Data>"
            Print #FileNumber, XML
        Else
            For i = idx + 1 To 5
                XML = XML & "        <택번호" & i & "></택번호" & i & ">"
            Next i
            
            XML = XML & "   </Data>"
            Print #FileNumber, XML
        End If
        
        Print #FileNumber, "</root>"
        Close #FileNumber
    End With
    
    With rpt지사출고현황
        .dc.FileURL = App.Path & "\XML\지사출고현황.XML"
        .PrintReport True
        '.Show 1
    End With

    Unload rpt지사출고현황
    
    Exit Sub

ErrRtn:
    MsgBox Err.Description, vbInformation, "오류"
    Screen.MousePointer = 0
End Sub

Private Sub DataScreen()

End Sub

Private Sub spdView_Click(ByVal Col As Long, ByVal Row As Long)
    If Row <= 0 Then Exit Sub
            
    Dim 가맹점코드 As String
    
    spdView.Row = Row
    spdView.Col = 1: 가맹점코드 = spdView.Text & ""
    
    Call Data_Display2(가맹점코드)
End Sub

Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    'Call spdView_Click(NewCol, NewRow)
End Sub




Private Sub chkColor_Setting(bColor As Boolean)
    Dim nRow        As Long
    Dim oldTag    As String
    Dim newTag    As String
    
    On Error GoTo ERR_RTN
    
    With spdView1
    
        If .DataRowCnt <= 1 Then Exit Sub
        
        .Row = 1:   .Col = 1:  oldTag = .Text
        
        For nRow = 2 To .DataRowCnt - 1
            .Row = nRow:    .Col = 1:  newTag = .Text
            
            If Not bColor Then
                .Col = -1
                .BackColor = vbWhite
            
            ElseIf oldTag = newTag Then
                .Col = -1
                .BackColor = &HC0C0FF   'vbRed
            
            Else
                .Col = -1
                .BackColor = vbWhite
            
            End If
        
            oldTag = newTag
        
        Next nRow
    End With
    
    Exit Sub
    
ERR_RTN:
    MsgBox Err.Description

End Sub

