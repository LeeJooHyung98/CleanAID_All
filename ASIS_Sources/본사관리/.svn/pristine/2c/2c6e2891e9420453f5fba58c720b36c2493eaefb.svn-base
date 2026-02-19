VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_02015 
   Caption         =   "가맹점 택번호 누락 현황"
   ClientHeight    =   12300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16035
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_02015.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12300
   ScaleWidth      =   16035
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   12300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16035
      _ExtentX        =   28284
      _ExtentY        =   21696
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_02015.frx":058A
      Begin Threed.SSPanel pnlHeader 
         Height          =   525
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   8400
         _ExtentX        =   14817
         _ExtentY        =   926
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
         Caption         =   " 가맹점 택번호 누락 현황 (P_02015)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_02015.frx":061C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   525
         Left            =   8430
         TabIndex        =   2
         Top             =   15
         Width           =   7590
         _ExtentX        =   13388
         _ExtentY        =   926
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
         PictureBackground=   "P_02015.frx":081E
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   3
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
            Picture         =   "P_02015.frx":0A20
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   4
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
            Picture         =   "P_02015.frx":0FBA
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   5
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
            Picture         =   "P_02015.frx":1554
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   6
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
            Picture         =   "P_02015.frx":1AEE
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   7
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
            Picture         =   "P_02015.frx":2088
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   8
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
            Picture         =   "P_02015.frx":2622
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   9
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
            Picture         =   "P_02015.frx":2BBC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   10
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
            Picture         =   "P_02015.frx":3156
         End
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   11
         Top             =   555
         Width           =   16005
         _ExtentX        =   28231
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Left            =   1245
            Style           =   2  '드롭다운 목록
            TabIndex        =   13
            Top             =   405
            Width           =   2700
         End
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1245
            Locked          =   -1  'True
            Style           =   2  '드롭다운 목록
            TabIndex        =   12
            Top             =   60
            Width           =   2700
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   14
            Top             =   405
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "가맹점명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   5910
            TabIndex        =   15
            Top             =   420
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            Format          =   68550656
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   4725
            TabIndex        =   16
            Top             =   420
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "접수일자"
            BorderWidth     =   0
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   10
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
            Left            =   9090
            TabIndex        =   18
            Top             =   420
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            Format          =   68550656
            CurrentDate     =   36686
         End
         Begin XtremeSuiteControls.PushButton cmdRefresh 
            Height          =   330
            Left            =   3990
            TabIndex        =   21
            Top             =   390
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   582
            _StockProps     =   79
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_02015.frx":36F0
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
            Left            =   8775
            TabIndex        =   19
            Top             =   480
            Width           =   300
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10935
         Left            =   15
         TabIndex        =   20
         Top             =   1350
         Width           =   16005
         _Version        =   524288
         _ExtentX        =   28231
         _ExtentY        =   19288
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
         SpreadDesigner  =   "P_02015.frx":3C8A
         UserResize      =   1
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_02015"
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

Private Sub cmdRefresh_Click()
    cboOffice_Click
End Sub

Private Sub dtInput_Change(Index As Integer)
'    Call Data_Display
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
    
    If P_02015_Flag = False Then
        dtInput(0).Value = Date
        dtInput(1).Value = Date

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
        
        P_02015_Flag = True
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_02015_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i      As String
    Dim 택번호 As String
    
    Dim iRow As Long
    Dim iCol  As Long
    Dim iCnt  As Long
    
    '----------------------------------------------------------------
    ' SP_02001_00
    '----------------------------------------------------------------
    ReDim sValue(2)
    
    sValue(0) = Mid(cboInput.Text, 2, 6)
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
        
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_02015_00", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_02015_00", sValue(), Err_Num, Err_Dec)
    End If
            
    택번호 = ""
    
    With spdView
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            If 택번호 <> "" Then
                If Left(택번호, 3) <> Left(RS01!택번호, 3) Then
                
                Else
                    If Left(택번호, 3) = Left(RS01!택번호, 3) Then
                        If (Format(택번호 + 1, "000000000") <> RS01!택번호) And (Format(택번호 + 1, "000000000") < RS01!택번호) Then
                            Do
                                .MaxRows = .MaxRows + 1
                                .Row = .MaxRows
                                .Col = 1: .Text = Format(택번호 + 1, "000000000")
                                
                                택번호 = 택번호 + 1
                            Loop Until Format(택번호 + 1, "000000000") = RS01!택번호
                        End If
                    End If
                End If
            End If
            
            택번호 = RS01!택번호 & ""
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
    
    
    
'    iRow = 0
'    iCol = 0
'    iCnt = 0
'
'    택번호 = ""
'
'    With spdView
'        .MaxRows = 0
'
'        Do Until RS01.EOF
'            iCnt = iCnt + 1
'
'            If iCol = 0 Or iCol > 15 Then
'                .MaxRows = .MaxRows + 1
'
'                iRow = iRow + 1
'                iCol = 0
'            End If
'
'            iCol = iCol + 1
'
'            .Row = iRow
'            .Col = iCol: .Value = Format(Right(RS01!택번호, 6), "00-0000") & ""
'
'            RS01.MoveNext
'        Loop
'        RS01.Close
'        Set RS01 = Nothing
'    End With
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

