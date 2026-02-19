VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_01008 
   Caption         =   "사용자 등록"
   ClientHeight    =   10530
   ClientLeft      =   3495
   ClientTop       =   3285
   ClientWidth     =   14730
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_01008.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10530
   ScaleWidth      =   14730
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   10530
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14730
      _ExtentX        =   25982
      _ExtentY        =   18574
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_01008.frx":058A
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Left            =   7125
         TabIndex        =   1
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
         PictureBackground=   "P_01008.frx":061C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   2
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
            Picture         =   "P_01008.frx":081E
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   3
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
            Picture         =   "P_01008.frx":0DB8
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   4
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
            Picture         =   "P_01008.frx":1352
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   5
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
            Picture         =   "P_01008.frx":18EC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   6
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
            Picture         =   "P_01008.frx":1E86
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   7
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
            Picture         =   "P_01008.frx":2420
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   8
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
            Picture         =   "P_01008.frx":29BA
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   9
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
            Picture         =   "P_01008.frx":2F54
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   10
         Top             =   15
         Width           =   7095
         _ExtentX        =   12515
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
         PictureBackground=   "P_01008.frx":34EE
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   8970
         Left            =   15
         TabIndex        =   11
         Top             =   1545
         Width           =   14700
         _Version        =   524288
         _ExtentX        =   25929
         _ExtentY        =   15822
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
         SpreadDesigner  =   "P_01008.frx":36F0
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   990
         Left            =   15
         TabIndex        =   12
         Top             =   540
         Width           =   14700
         _ExtentX        =   25929
         _ExtentY        =   1746
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1245
            Locked          =   -1  'True
            Style           =   2  '드롭다운 목록
            TabIndex        =   13
            Top             =   60
            Width           =   3420
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   7
            Left            =   60
            TabIndex        =   14
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
   End
End
Attribute VB_Name = "P_01008"
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
        Case 1: Call DataAdd        ' 신규
        Case 2: Call DataSave       ' 저장
        Case 3: Call DataDelete     ' 삭제
        Case 4: Call DataCancel     ' 취소
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
    cmdBtn(1).Enabled = True
    cmdBtn(2).Enabled = True
    cmdBtn(3).Enabled = True
    cmdBtn(4).Enabled = True
    
    '본사 여부
    If HeadOffice = MASTER_OFFICE_CODE Then
        cboOffice.ListIndex = 0
        cboOffice.Locked = False
    Else
        cboOffice.Locked = True
    End If
    
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"

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
            
            
        .ColsFrozen = 1 '틀고정
        .Row = -1
        
        .Col = 1
        .ColWidth(1) = 10
        .CellType = CellTypeEdit
        .TypeEditPassword = False
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 2
        .ColWidth(2) = 10
        .CellType = CellTypeEdit
        .TypeEditPassword = False
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 3
        .ColWidth(3) = 20
        .CellType = CellTypeEdit
        .TypeEditPassword = False
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
    
        .Col = 4
        .ColWidth(4) = 10
        .CellType = CellTypeEdit
        .TypeEditPassword = True
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
    
        .Col = 5
        .ColWidth(4) = 20
        .CellType = CellTypeEdit
        .TypeEditPassword = False
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
    End With
    
    If P_01008_Flag = False Then
        ReDim sValue(1)
        
        sValue(0) = "1"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_01008_00", sValue(), Err_Num, Err_Dec)
        
        spdView.MaxCols = RS01.Fields.Count
        spdView.MaxRows = RS01.RecordCount
        
        'Call spdDisplay(RS01)
        Call fpSpread_Display(spdView, RS01)
        Call GetColWidth(REG_App, Me.Name, spdView)
        
        P_01008_Flag = True
    End If
    
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
    P_01008_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    ReDim sValue(1)
    
    
    sValue(0) = "0"
    sValue(1) = Mid(cboOffice.Text, 2, 4)
    If sValue(1) = "0000" Then sValue(1) = "%"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01008_00", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    'Call spdDisplay(RS01)
    Call fpSpread_Display(spdView, RS01)
    Call GetColWidth(REG_App, Me.Name, spdView)
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Public Sub DataAdd()
    spdView.MaxRows = spdView.MaxRows + 1
    
    spdView.Row = spdView.MaxRows
    spdView.Col = 1
    spdView.Action = ActionActiveCell
    spdView.Lock = False
End Sub

Public Sub DataSave()
    On Error GoTo ErrRtn
    
    Dim i As Integer
    
    ReDim sValue(4)
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        spdView.Col = 1
        
        If Trim(spdView.Text) <> "" Then
            spdView.Col = 1: sValue(0) = spdView.Text & ""
            spdView.Col = 2: sValue(1) = spdView.Text & ""
            spdView.Col = 3: sValue(2) = spdView.Text & ""
            spdView.Col = 4: sValue(3) = spdView.Text & ""
            spdView.Col = 5: sValue(4) = spdView.Text & ""
            
            Call ExecPro("SP_01008_01", sValue(), Err_Num, Err_Dec)
        End If
    Next i

    If Err_Num = 0 Then
        MsgBox "해당되는 데이터가 정상적으로 저장이 되었습니다.", vbInformation
    Else
        MsgBox Err_Dec, vbInformation
    End If
    
    Exit Sub
    
ErrRtn:

End Sub

Public Sub DataDelete()
    If MsgBox("해당되는 데이터를 삭제하시겠습니까?", vbYesNo + vbInformation + vbDefaultButton2, "데이터 삭제") = vbYes Then
    
        ReDim sValue(1)
        
        spdView.Row = spdView.ActiveRow
        
        spdView.Col = 1: sValue(0) = spdView.Text & ""
        spdView.Col = 2: sValue(1) = spdView.Text & ""
        
        Call ExecPro("SP_01008_02", sValue(), Err_Num, Err_Dec)
        
        If Err_Num = 0 Then
            spdView.Row = spdView.ActiveRow
            spdView.Action = ActionDeleteRow
            
            MsgBox "해당되는 데이터가 정상적으로 삭제가 되었습니다.", vbInformation
        End If
    End If
End Sub

Public Sub DataCancel()
    Call Data_Display
End Sub

Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If NewRow <> -1 Then
        spdView.Row = Row
        spdView.Col = -1
        spdView.BackColor = vbWhite
        
        spdView.Row = NewRow
        spdView.Col = -1
        spdView.BackColor = glbYellow
    End If
End Sub

