VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_04036 
   Caption         =   "일일 판매 집계 (가맹점)"
   ClientHeight    =   10665
   ClientLeft      =   6300
   ClientTop       =   4245
   ClientWidth     =   14265
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_04036.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10665
   ScaleWidth      =   14265
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   10665
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14265
      _ExtentX        =   25162
      _ExtentY        =   18812
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_04036.frx":058A
      Begin Threed.SSPanel panInput 
         Height          =   825
         Left            =   15
         TabIndex        =   1
         Top             =   540
         Width           =   14235
         _ExtentX        =   25109
         _ExtentY        =   1455
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin XtremeSuiteControls.PushButton cmdRefresh 
            Height          =   330
            Left            =   4710
            TabIndex        =   22
            Top             =   420
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   582
            _StockProps     =   79
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_04036.frx":061C
         End
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1245
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
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   4710
            TabIndex        =   11
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "마감일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   4
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
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   10
            Left            =   60
            TabIndex        =   15
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
            Index           =   0
            Left            =   5880
            TabIndex        =   18
            Top             =   60
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   57606147
            CurrentDate     =   40279
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   7605
            TabIndex        =   19
            Top             =   60
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   57606147
            CurrentDate     =   40279
         End
         Begin VB.Label Label1 
            Caption         =   $"P_04036.frx":0BB6
            ForeColor       =   &H000000FF&
            Height          =   435
            Left            =   5220
            TabIndex        =   21
            Top             =   420
            Width           =   8745
         End
         Begin VB.Label Label 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
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
            Height          =   210
            Left            =   7380
            TabIndex        =   20
            Top             =   120
            Width           =   180
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   900
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   255
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
         Caption         =   " #"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_04036.frx":0C2E
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Left            =   6660
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
         PictureBackground=   "P_04036.frx":0E30
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
            Picture         =   "P_04036.frx":1032
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
            Picture         =   "P_04036.frx":15CC
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
            Picture         =   "P_04036.frx":1B66
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
            Picture         =   "P_04036.frx":2100
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
            Picture         =   "P_04036.frx":269A
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
            Picture         =   "P_04036.frx":2C34
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
            Picture         =   "P_04036.frx":31CE
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   16
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
            Picture         =   "P_04036.frx":3768
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   9270
         Left            =   15
         TabIndex        =   17
         Top             =   1380
         Width           =   14235
         _Version        =   524288
         _ExtentX        =   25109
         _ExtentY        =   16351
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
         MaxCols         =   35
         MaxRows         =   200
         OperationMode   =   1
         Protect         =   0   'False
         SpreadDesigner  =   "P_04036.frx":3D02
         UserResize      =   1
         VisibleCols     =   11
         VisibleRows     =   50
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_04036"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String
 

Private Sub cboOffice_Click()
    If cboOffice.ListIndex < 0 Then Exit Sub
    
    '-----------------------------------------------------------------
    '
    '-----------------------------------------------------------------
    cboInput.Clear
    
    ReDim sValue(2)

    sValue(0) = Mid(cboOffice.Text, 2, 4)
    sValue(1) = Format(dtInput(0).Value, "YYYY-01-01")
    sValue(2) = Format(dtInput(0).Value, "YYYY-12-31")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01001_00", sValue(), Err_Num, Err_Dec)
    
    
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
    Call cboOffice_Click
End Sub

Private Sub dtInput_Change(Index As Integer)
'    Call cboOffice_Click
'    Call Data_Display
End Sub

Private Sub Form_Activate()

    If P_04036_Flag = True Then Exit Sub
    P_04036_Flag = True
    
    cmdBtn(0).Enabled = True
    'cmdBtn(1).Enabled = True
'    cmdBtn(2).Enabled = True
    'cmdBtn(3).Enabled = True
    'cmdBtn(4).Enabled = True
        cmdBtn(6).Enabled = True
        
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
    DoEvents
    
    
End Sub

'Private Sub spdDisplay(RS As ADODB.Recordset)
'    Call fpSpread_Display(spdView, RS)
'End Sub

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
        .OperationMode = OperationModeNormal

        'Init the User Sort
        .UserColAction = UserColActionSort

'        .Col = 1: .ColMerge = MergeRestricted
'        .Col = 2: .ColMerge = MergeRestricted
        
        .ColsFrozen = 5 '틀고정
        .Row = -1

'        .Col = 1
'        .ColWidth(1) = 10
'        .CellType = CellTypeStaticText
'        .TypeVAlign = TypeVAlignCenter
'        .TypeHAlign = TypeHAlignCenter
'
'        .Col = 2
'        .ColWidth(2) = 30
'        .CellType = CellTypeStaticText
'        .TypeVAlign = TypeVAlignCenter
'        .TypeHAlign = TypeHAlignLeft

'        .Col = -1
'        .Row = 3
'        .CellType = CellTypePercent
'        .TypeVAlign = TypeVAlignCenter
'        .TypeHAlign = TypeHAlignLeft

    End With

    dtInput(0).Value = Format(Date, "yyyy-MM") & "-01"
    dtInput(1).Value = Date
    
    
    Call Get_지사리스트(cboOffice)
    Call Set_지사선택고정(cboOffice, HeadOffice)
    
    Call GetColWidth(REG_App, Me.Name, spdView)
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_04036_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn
    

    ReDim sValue(3)
    
    Screen.MousePointer = vbHourglass
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    sValue(1) = Mid(cboInput.Text, 2, 6)
    sValue(2) = Format(dtInput(0).Value, "yyyy-MM-dd")
    sValue(3) = Format(dtInput(1).Value, "yyyy-MM-dd")
    
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(sValue(0)) = False Then Exit Sub
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_04036_01_new", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_04036_01_new", sValue(), Err_Num, Err_Dec)
    End If
    
    spdView.MaxRows = 0
    spdView.MaxRows = RS01.RecordCount
    
    Call fpSpread_Display(spdView, RS01, False)
    
    
    ' 합계 출력
    Dim nCol As Long
    For nCol = 2 To spdView.MaxCols
        Select Case nCol
            Case 2: Call SpreadSum(spdView, 1, nCol)
            Case Else: Call SpreadSum(spdView, -1, nCol)
        End Select
    Next nCol
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrRtn:
    Screen.MousePointer = vbDefault
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub


Public Sub DataSave()
 
End Sub

Public Sub DataDelete()

End Sub


Private Sub spdView_EditChange(ByVal Col As Long, ByVal Row As Long)
'    With spdView
'        .Row = Row
'        .Col = 3: .Formula = "SUM(D" & Row & ":O" & Row & ")"
'    End With

End Sub

Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'    If NewRow <> -1 Then
'        spdView.Row = Row
'        spdView.Col = -1
'        spdView.BackColor = vbWhite
'
'        spdView.Row = NewRow
'        spdView.Col = -1
'        spdView.BackColor = glbYellow
'    End If
End Sub

