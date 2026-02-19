VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_01001_M 
   Caption         =   "지사 등록"
   ClientHeight    =   12180
   ClientLeft      =   720
   ClientTop       =   2760
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_01001_M.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12180
   ScaleWidth      =   15240
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   12180
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   21484
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_01001_M.frx":058A
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10830
         Left            =   15
         TabIndex        =   1
         Top             =   1335
         Width           =   5385
         _Version        =   524288
         _ExtentX        =   9499
         _ExtentY        =   19103
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
         MaxCols         =   11
         SpreadDesigner  =   "P_01001_M.frx":067C
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdView1 
         Height          =   7425
         Left            =   5415
         TabIndex        =   2
         Top             =   4740
         Width           =   9810
         _Version        =   524288
         _ExtentX        =   17304
         _ExtentY        =   13097
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
         MaxCols         =   5
         SpreadDesigner  =   "P_01001_M.frx":0DAC
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panCaption 
         Height          =   390
         Index           =   5
         Left            =   5415
         TabIndex        =   3
         Top             =   4335
         Width           =   9810
         _ExtentX        =   17304
         _ExtentY        =   688
         _Version        =   262144
         Font3D          =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " 사업장 택 사용현황"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_01001_M.frx":135F
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   4
         Top             =   540
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel panDetail 
         Height          =   2985
         Left            =   5415
         TabIndex        =   5
         Top             =   1335
         Width           =   9810
         _ExtentX        =   17304
         _ExtentY        =   5265
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.TextBox txtInput 
            Enabled         =   0   'False
            Height          =   315
            Index           =   8
            Left            =   1890
            TabIndex        =   36
            Top             =   2580
            Width           =   6525
         End
         Begin VB.ComboBox cboTeam 
            Height          =   315
            Left            =   5505
            Style           =   2  '드롭다운 목록
            TabIndex        =   34
            Top             =   60
            Width           =   2985
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   7
            Left            =   1170
            MaxLength       =   50
            TabIndex        =   12
            Top             =   1515
            Width           =   7290
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   6
            Left            =   1170
            MaxLength       =   50
            TabIndex        =   11
            Top             =   1155
            Width           =   7290
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   5
            Left            =   5520
            MaxLength       =   50
            TabIndex        =   10
            Top             =   795
            Width           =   2955
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   4
            Left            =   1170
            MaxLength       =   50
            TabIndex        =   9
            Top             =   795
            Width           =   2955
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   3
            Left            =   5520
            MaxLength       =   50
            TabIndex        =   8
            Top             =   435
            Width           =   2955
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   2
            Left            =   1170
            MaxLength       =   50
            TabIndex        =   7
            Top             =   435
            Width           =   2955
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   1
            Left            =   1170
            MaxLength       =   4
            TabIndex        =   6
            Top             =   75
            Width           =   1335
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   4
            Left            =   1170
            TabIndex        =   15
            Top             =   1875
            Width           =   3750
            _ExtentX        =   6615
            _ExtentY        =   582
            _Version        =   262144
            BackColor       =   16777215
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   1
               Left            =   2100
               TabIndex        =   14
               Top             =   30
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   450
               _Version        =   262144
               BackColor       =   16777215
               Caption         =   "폐점"
            End
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   0
               Left            =   360
               TabIndex        =   13
               Top             =   30
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   450
               _Version        =   262144
               BackColor       =   16777215
               Caption         =   "개점"
            End
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1890
            TabIndex        =   37
            Top             =   2250
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   21430272
            CurrentDate     =   36684
         End
         Begin VB.Label lblTitle 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "P/G 사용 종료일:"
            Height          =   225
            Index           =   31
            Left            =   120
            TabIndex        =   39
            ToolTipText     =   "종료일 까지 접속 가능 합니다."
            Top             =   2310
            Width           =   1710
         End
         Begin VB.Label lblTitle 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "최종 접속일자:"
            Height          =   195
            Index           =   32
            Left            =   360
            TabIndex        =   38
            Top             =   2625
            Width           =   1470
         End
         Begin VB.Label lblTitle 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "영업팀 구분:"
            Height          =   225
            Index           =   8
            Left            =   4110
            TabIndex        =   35
            Top             =   150
            Width           =   1335
         End
         Begin VB.Label lblTitle 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "메모:"
            Height          =   225
            Index           =   7
            Left            =   45
            TabIndex        =   33
            Top             =   1575
            Width           =   1065
         End
         Begin VB.Label lblTitle 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "주소:"
            Height          =   225
            Index           =   6
            Left            =   45
            TabIndex        =   32
            Top             =   1215
            Width           =   1065
         End
         Begin VB.Label lblTitle 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "휴대전화:"
            Height          =   225
            Index           =   5
            Left            =   4395
            TabIndex        =   31
            Top             =   855
            Width           =   1065
         End
         Begin VB.Label lblTitle 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "전화번호:"
            Height          =   225
            Index           =   3
            Left            =   45
            TabIndex        =   30
            Top             =   855
            Width           =   1065
         End
         Begin VB.Label lblTitle 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "지사장명:"
            Height          =   225
            Index           =   2
            Left            =   4395
            TabIndex        =   29
            Top             =   495
            Width           =   1065
         End
         Begin VB.Label lblTitle 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "지사상태:"
            Height          =   225
            Index           =   4
            Left            =   45
            TabIndex        =   28
            Top             =   1935
            Width           =   1065
         End
         Begin VB.Label lblTitle 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "지 사 명:"
            Height          =   225
            Index           =   1
            Left            =   45
            TabIndex        =   27
            Top             =   495
            Width           =   1065
         End
         Begin VB.Label lblTitle 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "지사코드:"
            Height          =   225
            Index           =   0
            Left            =   45
            TabIndex        =   26
            Top             =   135
            Width           =   1065
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   16
         Top             =   15
         Width           =   7605
         _ExtentX        =   13414
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
         Caption         =   " 지사 등록 (P_01001_M)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_01001_M.frx":17C1
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Left            =   7635
         TabIndex        =   17
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
         PictureBackground=   "P_01001_M.frx":19C3
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   18
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
            Picture         =   "P_01001_M.frx":1BC5
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   19
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
            Picture         =   "P_01001_M.frx":215F
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   20
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
            Picture         =   "P_01001_M.frx":26F9
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   21
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
            Picture         =   "P_01001_M.frx":2C93
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   22
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
            Picture         =   "P_01001_M.frx":322D
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   23
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
            Picture         =   "P_01001_M.frx":37C7
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   24
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
            Picture         =   "P_01001_M.frx":3D61
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   25
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
            Picture         =   "P_01001_M.frx":42FB
         End
      End
   End
End
Attribute VB_Name = "P_01001_M"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Dim sPrintOption As String

Private Sub Data_Display(Optional 지사코드 As String)
    On Error GoTo ErrRtn

    Dim i As Integer
    
    txtInput(1).Enabled = False
    
    ReDim sValue(0)
    
    sValue(0) = ""
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01001_M0_ALL", sValue(), Err_Num, Err_Dec)
    
    With spdView
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = RS01!지사코드 & ""
            .Col = 2:  .Text = RS01!지사명 & ""
            .Col = 3:  .Text = RS01!지사상태 & ""
            
            If RS01!지사상태 = "개점" Then
                .ForeColor = vbBlack
            Else
                .ForeColor = vbRed
            End If
            
            .Col = 4:  .Text = RS01!지사장명 & ""
            .Col = 5:  .Text = RS01!전화번호 & ""
            .Col = 6:  .Text = RS01!휴대전화 & ""
            .Col = 7:  .Text = RS01!주소 & ""
            .Col = 8:  .Text = RS01!메모 & ""
            .Col = 9:  .Text = RS01!영업팀 & ""
            .Col = 10:  .Text = RS01!pg사용종료일 & ""
            .Col = 11:  .Text = RS01!PG접속일자 & ""
            
            
            RS01.MoveNext
        Loop
        .Redraw = True
        
        RS01.Close
        Set RS01 = Nothing
        
        If 지사코드 <> "" Then
            Rtn = .SearchCol(1, 1, .MaxRows, 지사코드, SearchFlagsValue)
            
            If Rtn > -1 Then
                .SetSelection 1, Rtn, .MaxCols, Rtn
            End If
        End If
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

'Private Sub spdDisplay(RS As ADODB.Recordset)
'    Call fpSpread_Display(spdView, RS)
'End Sub

Private Sub cmdPrint_Click()
    'Call DataScreen2
    'panPrint.Visible = False
End Sub

Private Sub cmdSub_Click()
    'Call DataSubSave
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: Call Data_Display   ' 조회
        Case 1: Call DataAdd        ' 신규
        Case 2: Call DataSave       ' 저장
        Case 3: Call DataDelete     ' 삭제
        Case 4: Call DataCancel     ' 취소
        Case 5: 'Call DataPrint     ' 인쇄
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
    cmdBtn(1).Enabled = True
    cmdBtn(2).Enabled = True
    
    cmdBtn(5).Enabled = True
    cmdBtn(6).Enabled = True
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
'    If P_01001_M_Flag = False Then
'        Call Data_Display
'
'        P_01001_M_Flag = True
'    End If
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
    
    If P_01001_M_Flag = False Then
        Call Data_Display

        P_01001_M_Flag = True
    End If
    
    With cboTeam
        .Clear
        .AddItem "[00] 없음"
        .AddItem "[01] 1팀"
        .AddItem "[02] 2팀"
        
        .ListIndex = 0
    End With
    
    dtInput(0).Value = Date
    dtInput(0).Value = ""
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_01001_M_Flag = False
End Sub

Private Sub spdView_Click(ByVal Col As Long, ByVal Row As Long)
    If Row <= 0 Then Exit Sub
    
    Call Data_Display2(Row)
End Sub

Private Sub Data_Display2(iRow As Long)
    Dim i As Integer

    With spdView
        .Row = iRow
        .Col = 1: txtInput(1).Text = .Text
        .Col = 2: txtInput(2).Text = .Text
        
        .Col = 3
        If .Text = "개점" Then
            optSelect(0).Value = True
        Else
            optSelect(1).Value = True
        End If
        
        .Col = 4: txtInput(3).Text = .Text
        .Col = 5: txtInput(4).Text = .Text
        .Col = 6: txtInput(5).Text = .Text
        .Col = 7: txtInput(6).Text = .Text
        .Col = 8: txtInput(7).Text = .Text
        
        .Col = 9: cboTeam.ListIndex = 0
        
        .Col = 10
        If Trim(.Text) = "" Then
            dtInput(0).Value = ""                                         '20
        Else
            dtInput(0).Value = Format(.Text, "YYYY-MM-DD")        '20
        End If
        .Col = 11: txtInput(8).Text = .Text
        
        If Trim(.Text) <> "" Then
            
            For i = 0 To cboTeam.ListCount - 1
                If Left(.Text, 2) = Mid(cboTeam.List(i), 2, 2) Then '12
                    cboTeam.ListIndex = i
                    Exit For
                End If
            Next i
        End If
        
    End With
    
    Call Data_Display3
End Sub

Private Sub Data_Display3()
    ReDim sValue(0)
    
    sValue(0) = txtInput(1).Text
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01001_M1_ALL", sValue(), Err_Num, Err_Dec)
    
    With spdView1
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = RS01!가맹점코드 & ""
            .Col = 2:  .Text = RS01!가맹점명 & ""
            .Col = 3:  .Text = RS01!택코드 & ""
            .Col = 4:  .Text = RS01!시작일자 & ""
            .Col = 5:  .Text = RS01!종료일자 & ""
            
            RS01.MoveNext
        Loop
        
        RS01.Close
        Set RS01 = Nothing
    End With

End Sub

Private Sub DataAdd()
    Dim i As Integer
      
    txtInput(1).Enabled = True
    
    For i = 1 To txtInput.Count - 1
        txtInput(i).Text = ""
    Next i
    
    spdView1.MaxRows = 0
    txtInput(1).SetFocus
End Sub

Private Sub DataCancel()
    'Call Data_Display2
End Sub

Private Sub DataDelete()
'    If MsgBox("해당되는 대리점코드를 삭제하시겠습니까?", vbYesNo + vbInformation + vbDefaultButton2, "데이터 삭제") = vbYes Then
'
'        ReDim sValue(1)
'
'        sValue(0) = txtInput(1).Text
'        sValue(1) = Mid(cboInput(3).Text, 2, 4)
'
'        Call ExecPro("SP_01001_02_MASTER", sValue(), Err_Num, Err_Dec)
'
'        If Err_Num = 0 Then
'            spdView.Row = spdView.ActiveRow
'            spdView.Action = ActionDeleteRow
'
'            MsgBox "해당되는 데이터가 정상적으로 삭제가 되었습니다.", vbInformation
'        End If
'    End If
End Sub

Private Sub DataSave()
    Rtn = MsgBox("해당되는 내역을 저장하시겠습니까?", vbYesNo + vbInformation, "데이터 저장")
    
    If Rtn = vbNo Then Exit Sub
    
    ReDim sValue(9)
    
    sValue(0) = txtInput(1).Text        ' 사업장코드
    sValue(1) = txtInput(2).Text        ' 사업장명
    
    If optSelect(0).Value = True Then   ' 사업장상태
        sValue(2) = "Y"
    Else
        sValue(2) = "N"
    End If
    
    sValue(3) = txtInput(3).Text        '
    sValue(4) = txtInput(4).Text        '
    sValue(5) = txtInput(5).Text        '
    sValue(6) = txtInput(6).Text        '
    sValue(7) = txtInput(7).Text        '
    sValue(8) = Mid(cboTeam.Text, 2, 2) '
    sValue(9) = Format(dtInput(0).Value, "YYYY-MM-DD")               ' 프로그램 사용 종료일

    
     If Trim(sValue(0)) = "" Or IsNumeric(sValue(0)) = False Or Len(sValue(0)) <> 4 Then
        MsgBox "지사코드를 확인하여 주십시요.", vbInformation
        Exit Sub
     End If
     
    Call ExecPro("SP_01001_M3_ALL_NEW", sValue(), Err_Num, Err_Dec)
    
    If Err_Num = 0 Then
        Call Data_Display(txtInput(1).Text)
            
        MsgBox "해당되는 데이터가 정상적으로 저장이 되었습니다.", vbInformation
    End If
End Sub

Private Sub DataPrint()
    Dim ReportFP As String
    Dim ReportFile As String
    
'    ReportFP = GetIniStr("REPORT", "FilePath", "", m_iniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    P_00000.crPrint.StoredProcParam(0) = "0"
'    P_00000.crPrint.StoredProcParam(1) = txtInput(1).Text
'    P_00000.crPrint.WindowTitle = Me.Caption
'
'    Call ReportPrint(ReportFile, "1")
End Sub

Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Call spdView_Click(NewCol, NewRow)
End Sub

Private Sub DataScreen()

End Sub

Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case Index
            Case 1
                If Len(Trim(txtInput(1).Text)) <> 4 Then
                    MsgBox "지사코드는 4자리로 구성 하여야 합니다", vbInformation
                    txtInput(1).SetFocus
                Else
                    ReDim sValue(0)
                    sValue(0) = txtInput(1).Text
                    
                    Set RS01 = New ADODB.Recordset
                    Set RS01 = ExecPro("SP_A_0002", sValue(), Err_Num, Err_Dec)
                    
                    If RS01.EOF Then
                        RS01.Close
                        Set RS01 = Nothing
                        
                        txtInput(2).Text = ""
                    Else
                        MsgBox "지사코드 [" & txtInput(1).Text & "]는 " & RS01!지사명 & "으로 등록 되어 있습니다." & Chr(13) & "확인후 등록 바랍니다.", vbInformation
                        
                        RS01.Close
                        Set RS01 = Nothing
                        
                        txtInput(1).Text = ""
                        txtInput(1).SetFocus
                    End If
                End If
            Case Else
                SendKeys "{TAB}"
        End Select
    End If
End Sub

