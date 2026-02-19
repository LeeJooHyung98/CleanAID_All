VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_01010 
   Caption         =   "프로그램 권한 등록"
   ClientHeight    =   10665
   ClientLeft      =   7785
   ClientTop       =   4680
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
   Icon            =   "P_01010.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
      PaneTree        =   "P_01010.frx":058A
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   9420
         Left            =   15
         TabIndex        =   1
         Top             =   1230
         Width           =   14235
         _Version        =   524288
         _ExtentX        =   25109
         _ExtentY        =   16616
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
         SpreadDesigner  =   "P_01010.frx":061C
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   675
         Left            =   15
         TabIndex        =   2
         Top             =   540
         Width           =   14235
         _ExtentX        =   25109
         _ExtentY        =   1191
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Left            =   1530
            Style           =   2  '드롭다운 목록
            TabIndex        =   3
            Top             =   60
            Width           =   2775
         End
         Begin Threed.SSCommand cmdSubBtn 
            Height          =   375
            Index           =   0
            Left            =   11205
            TabIndex        =   4
            Top             =   30
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   262144
            Caption         =   "전 체 선 택"
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   5
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "사용자명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   6
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
         PictureBackground=   "P_01010.frx":0A84
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Left            =   6660
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
         PictureBackground=   "P_01010.frx":0C86
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
            Picture         =   "P_01010.frx":0E88
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
            Caption         =   "화면"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_01010.frx":1422
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
            Picture         =   "P_01010.frx":19BC
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
            Picture         =   "P_01010.frx":1F56
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
            Picture         =   "P_01010.frx":24F0
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
            Picture         =   "P_01010.frx":2A8A
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
            Picture         =   "P_01010.frx":3024
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
            Picture         =   "P_01010.frx":35BE
         End
      End
   End
End
Attribute VB_Name = "P_01010"
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

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: Call Data_Display   ' 조회
        Case 1: Call DataAdd        ' 신규
        Case 2: Call DataSave       ' 저장
        Case 3: Call DataDelete     ' 삭제
        Case 4: Call DataCancel     ' 취소
        Case 5: 'Call DataPrint      ' 인쇄
        Case 6: 'Call DataScreen     '
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

Private Sub cmdSubBtn_Click(Index As Integer)
    Dim i As Integer
    
    Select Case Index
        Case 0
            For i = 1 To spdView.MaxRows
                spdView.Row = i
                spdView.Col = 3
                If spdView.Value = True Then
                    spdView.Value = False
                Else
                    spdView.Value = True
                End If
            Next i
    End Select
End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(1).Enabled = True
    cmdBtn(2).Enabled = True
    cmdBtn(3).Enabled = True
    cmdBtn(4).Enabled = True
    
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
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
        
        .Col = 2
        .ColWidth(2) = 40
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
        
        .Col = 3
        .ColWidth(3) = 10
        .CellType = CellTypeCheckBox
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
    End With

    If P_01010_Flag = False Then
        Call ComboAdd
        
        ReDim sValue(1)
        
        sValue(0) = "1"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_01010_00", sValue(), Err_Num, Err_Dec)
        
        spdView.MaxCols = RS01.Fields.Count
        spdView.MaxRows = RS01.RecordCount
        
        'Call spdDisplay(RS01)
        Call fpSpread_Display(spdView, RS01)
        Call GetColWidth(REG_App, Me.Name, spdView)
        
        P_01010_Flag = True
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_01010_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    ReDim sValue(1)
    
    sValue(0) = "0"
    sValue(1) = Mid(cboInput.Text, 2, 6)
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01010_00", sValue(), Err_Num, Err_Dec)
    
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
    Dim i As Integer
    ReDim sValue(2)
    
    sValue(0) = Mid(cboInput.Text, 2, 6)
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        
        spdView.Col = 1: sValue(1) = spdView.Text
        
        spdView.Col = 3
        If spdView.Value = True Then
            sValue(2) = "Y"
        Else
            sValue(2) = "N"
        End If
        
        If Trim(sValue(0)) = "" Then
            Exit Sub
        End If
        
        Call ExecPro("SP_01010_01", sValue(), Err_Num, Err_Dec)
    Next i

    If Err_Num = 0 Then
        MsgBox "해당되는 데이터가 정상적으로 저장이 되었습니다.", vbInformation
    End If
End Sub

Public Sub DataDelete()
    If MsgBox("해당되는 데이터를 삭제하시겠습니까?", vbYesNo + vbInformation + vbDefaultButton2, "데이터 삭제") = vbYes Then
    
        ReDim sValue(0)
        
        spdView.Row = spdView.ActiveRow
        spdView.Col = 1
        sValue(0) = spdView.Text
        
        Call ExecPro("SP_01010_02", sValue(), Err_Num, Err_Dec)
        
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

Private Sub ComboAdd()
    ReDim sValue(0)
    
    sValue(0) = "0"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_00004", sValue(), Err_Num, Err_Dec)

    Do While Not RS01.EOF
        cboInput.AddItem "[" & RS01!사용자ID & "] " & RS01!사용자명
        
        RS01.MoveNext
    Loop
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

