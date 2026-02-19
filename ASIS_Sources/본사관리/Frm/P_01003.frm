VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_01003 
   Caption         =   "대표 품목 현황"
   ClientHeight    =   11205
   ClientLeft      =   1635
   ClientTop       =   2610
   ClientWidth     =   15555
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_01003.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11205
   ScaleWidth      =   15555
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   11205
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15555
      _ExtentX        =   27437
      _ExtentY        =   19764
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_01003.frx":058A
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   1
         Top             =   540
         Width           =   15525
         _ExtentX        =   27384
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   7920
         _ExtentX        =   13970
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
         Caption         =   " 대표 품목 현황 (P_01003)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_01003.frx":063C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Left            =   7950
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
         PictureBackground=   "P_01003.frx":083E
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
            Picture         =   "P_01003.frx":0A40
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
            Picture         =   "P_01003.frx":0FDA
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
            Picture         =   "P_01003.frx":1574
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
            Picture         =   "P_01003.frx":1B0E
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
            Picture         =   "P_01003.frx":20A8
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
            Picture         =   "P_01003.frx":2642
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
            Picture         =   "P_01003.frx":2BDC
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
            Picture         =   "P_01003.frx":3176
         End
      End
      Begin FPSpreadADO.fpSpread spdView1 
         Height          =   9855
         Left            =   4845
         TabIndex        =   12
         Top             =   1335
         Width           =   10695
         _Version        =   524288
         _ExtentX        =   18865
         _ExtentY        =   17383
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
         SpreadDesigner  =   "P_01003.frx":3710
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   9855
         Left            =   15
         TabIndex        =   13
         Top             =   1335
         Width           =   4815
         _Version        =   524288
         _ExtentX        =   8493
         _ExtentY        =   17383
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
         SpreadDesigner  =   "P_01003.frx":3C4E
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_01003"
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
        Case 0: 'Call Data_Display   ' 조회
        Case 1: 'Call DataAdd        ' 신규
        Case 2: 'Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
        Case 4: 'Call DataCancel     ' 취소
        Case 5: Call DataPrint      ' 인쇄
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

Private Sub DataPrint()
    On Error GoTo ErrRtn
    
    Dim XML         As String
    Dim i           As Integer
    Dim FileNumber
        
    Dim 품목분류명  As String
    
    If spdView.ActiveRow <= 0 Then Exit Sub
    
    If spdView1.MaxRows = 0 Then Exit Sub
    
    '
    spdView.Row = spdView.ActiveRow
    spdView.Col = 2: 품목분류명 = spdView.Text & ""
    
    FileNumber = FreeFile
    
    Open App.Path & "\XML\대표품목현황.XML" For Output As #FileNumber
    
    Print #FileNumber, "<?xml version=""1.0"" encoding=""EUC-KR""?>"
    Print #FileNumber, "<root>"
    
          XML = "    <조건>"
    XML = XML & "        <품목분류명>품목분류명 : " & Func_Replace(품목분류명) & "</품목분류명>"
    XML = XML & "   </조건>"
    Print #FileNumber, XML
    
    With spdView
        For i = 1 To .MaxRows
            .Row = i
            
                            XML = "    <Data>"
            .Col = 1: XML = XML & "        <의류코드>" & .Text & "</의류코드>"
            .Col = 2: XML = XML & "        <의류명>" & .Text & "</의류명>"
            .Col = 3: XML = XML & "        <금액>" & .Text & "</금액>"
                      XML = XML & "   </Data>"
            Print #FileNumber, XML
        Next i
    
        Print #FileNumber, "</root>"
        Close #FileNumber
    End With
    
    With rpt대표품목현황
        .dc.FileURL = App.Path & "\XML\대표품목현황.XML"
        .PrintReport False
        
        '.Show 1
    End With

    Unload rpt지사반품현황
    
    Exit Sub

ErrRtn:
    MsgBox Err.Description, vbInformation, "오류"
    Screen.MousePointer = 0
End Sub

Private Sub Form_Activate()
    cmdBtn(5).Enabled = True
    cmdBtn(6).Enabled = True
    
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
        '.UserColAction = UserColActionSort
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
        .OperationMode = OperationModeSingle
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With

    If P_01003_Flag = False Then
        '-------------------------------------------------------------
        ' TB_의류분류
        '-------------------------------------------------------------
        Dim 수량 As Long
        
        ReDim sValue(0)

        sValue(0) = "0"
        
        'LAUNDRY1000 서버 데이터 조회
        If DBOpen_Master("1000") = False Then Exit Sub
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_00013", sValue(), Err_Num, Err_Dec)

'        Set RS01 = New ADODB.Recordset
'        Set RS01 = ExecPro("SP_00013", sValue(), Err_Num, Err_Dec)
    
        With spdView
            .MaxRows = 0
            .Redraw = False
                        
            '-전체-
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = ""
            .Col = 2: .Text = "-전체-"
            .Col = 3: .Text = "0"
                
            수량 = 0
            
            Do Until RS01.EOF
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                
                .Col = 1: .Text = RS01!의류분류코드 & ""
                .Col = 2: .Text = RS01!의류분류명 & ""
                .Col = 3: .Text = RS01!수량 & ""
                
                수량 = 수량 + RS01!수량
                
                RS01.MoveNext
            Loop
            RS01.Close
            Set RS01 = Nothing
            
            .Row = 1
            .Col = 3: .Text = 수량 & ""
            
            .Redraw = True
        End With
                
        P_01003_Flag = True
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_01003_Flag = False
End Sub

Private Sub Data_Display(의류분류코드 As String)
    On Error GoTo ErrRtn

    ReDim sValue(0)
    
    sValue(0) = 의류분류코드
    
    'LAUNDRY1000 서버 데이터 조회
    If DBOpen_Master("1000") = False Then Exit Sub
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecProMaster("SP_01003_00", sValue(), Err_Num, Err_Dec)
    
'    Set RS01 = New ADODB.Recordset
'    Set RS01 = ExecPro("SP_01003_00", sValue(), Err_Num, Err_Dec)

    With spdView1
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = RS01!코드 & ""
            .Col = 2: .Text = RS01!품목명 & ""
            .Col = 3: .Text = RS01!단가 & ""
            
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

Private Sub DataAdd()

End Sub

Private Sub DataSave()

End Sub

Private Sub DataDelete()

End Sub

Private Sub DataCancel()

End Sub

Private Sub spdView_Click(ByVal Col As Long, ByVal Row As Long)
    If Row <= 0 Then Exit Sub
    
    spdView.Row = Row
    spdView.Col = 1
    
    Call Data_Display(spdView.Text)
End Sub

Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Call spdView_Click(NewCol, NewRow)
End Sub
