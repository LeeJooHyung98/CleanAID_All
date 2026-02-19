VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_01005 
   Caption         =   "특정세일 등록"
   ClientHeight    =   11580
   ClientLeft      =   1620
   ClientTop       =   1920
   ClientWidth     =   16305
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_01005.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11580
   ScaleWidth      =   16305
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   11580
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16305
      _ExtentX        =   28760
      _ExtentY        =   20426
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_01005.frx":058A
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10230
         Left            =   4875
         TabIndex        =   1
         Top             =   1335
         Width           =   11415
         _Version        =   524288
         _ExtentX        =   20135
         _ExtentY        =   18045
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
         SpreadDesigner  =   "P_01005.frx":063C
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panSub 
         Height          =   10230
         Left            =   15
         TabIndex        =   2
         Top             =   1335
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   18045
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   1740
            Style           =   2  '드롭다운 목록
            TabIndex        =   7
            Top             =   1200
            Width           =   3075
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   1740
            TabIndex        =   3
            Top             =   120
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   56492032
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   8
            Left            =   1740
            TabIndex        =   4
            Top             =   480
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   262144
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   0
               Left            =   315
               TabIndex        =   5
               Top             =   30
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "전  체"
               Value           =   -1
            End
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   1
               Left            =   1740
               TabIndex        =   6
               Top             =   30
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "부  분"
            End
         End
         Begin MSMask.MaskEdBox mskInput 
            Height          =   315
            Left            =   1740
            TabIndex        =   8
            Top             =   840
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "#,##0.0"
            PromptChar      =   "_"
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   9
            Left            =   90
            TabIndex        =   9
            Top             =   480
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "할 인 구 분"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   4
            Left            =   90
            TabIndex        =   10
            Top             =   840
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "할  인  율"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   5
            Left            =   90
            TabIndex        =   11
            Top             =   1200
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "대 표 품 번"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   3
            Left            =   90
            TabIndex        =   12
            Top             =   120
            Width           =   1635
            _ExtentX        =   2884
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
            Caption         =   "적 용 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   13
         Top             =   540
         Width           =   16275
         _ExtentX        =   28707
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            Style           =   2  '드롭다운 목록
            TabIndex        =   14
            Top             =   60
            Width           =   2775
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   15
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
            Caption         =   "가맹점명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   16
         Top             =   15
         Width           =   8670
         _ExtentX        =   15293
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
         PictureBackground=   "P_01005.frx":0B1C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Left            =   8700
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
         PictureBackground=   "P_01005.frx":0D1E
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
            Picture         =   "P_01005.frx":0F20
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
            Caption         =   "화면"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_01005.frx":14BA
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
            Picture         =   "P_01005.frx":1A54
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
            Picture         =   "P_01005.frx":1FEE
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
            Picture         =   "P_01005.frx":2588
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
            Picture         =   "P_01005.frx":2B22
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
            Picture         =   "P_01005.frx":30BC
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
            Picture         =   "P_01005.frx":3656
         End
      End
   End
End
Attribute VB_Name = "P_01005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cboInput_Click(Index As Integer)
    If Index = 1 Then
        Call Data_Display2
    End If
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: Call Data_Display   ' 조회
        Case 1: Call DataAdd        ' 신규
        Case 2: Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
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

Private Sub dtInput_Change(Index As Integer)
    If Index = 1 Then
        cboInput(1).Clear
        
        ReDim sValue(2)
        
        sValue(0) = "0"
        sValue(1) = Mid(cboInput(0).Text, 2, 3)
        sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_00006", sValue(), Err_Num, Err_Dec)
        
        Do While Not RS01.EOF
            cboInput(1).AddItem "[" + RS01!품목코드 + "] " + RS01!품목명
        
            RS01.MoveNext
        Loop
    End If
End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(2).Enabled = True
    cmdBtn(4).Enabled = True
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
'    If P_01005_Flag = False Then
'        Call AgencyComboAdd(cboInput(0))
'
'        ReDim sValue(2)
'
'        sValue(0) = "1"
'
'        Set RS01 = New ADODB.Recordset
'        Set RS01 = ExecPro("SP_01005_00", sValue(), Err_Num, Err_Dec)
'
'        spdView.MaxCols = RS01.Fields.Count
'        spdView.MaxRows = RS01.RecordCount
'
'        Call spdDisplay(RS01)
'
'        P_01005_Flag = True
'    End If
End Sub

'Private Sub spdDisplay(RS As ADODB.Recordset)
'    Call fpSpread_Display(spdView, RS)
'
'    Set spdView.DataSource = Nothing
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
        .OperationMode = OperationModeSingle
        
        'Init the User Sort
        .UserColAction = UserColActionSort
            
            
        .ColsFrozen = 1 '틀고정
        .Row = -1
        
        .Col = 1
        .ColWidth(1) = 8
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 2
        .ColWidth(2) = 20
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
        
        .Col = 3
        .ColWidth(3) = 12
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 0
        .TypeVAlign = TypeVAlignCenter
    
        .Col = 4
        .ColWidth(4) = 8
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 0
        .TypeVAlign = TypeVAlignCenter
        
        .Col = 5
        .ColWidth(5) = 12
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 0
        .TypeVAlign = TypeVAlignCenter
        
        .Col = 6
        .ColWidth(6) = 6
        .CellType = CellTypeCheckBox
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
    End With
    
    If P_01005_Flag = False Then
        Call AgencyComboAdd(cboInput(0))
        
        ReDim sValue(2)
        
        sValue(0) = "1"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_01005_00", sValue(), Err_Num, Err_Dec)
        
        spdView.MaxCols = RS01.Fields.Count
        spdView.MaxRows = RS01.RecordCount
        
        'Call spdDisplay(RS01)
        Call fpSpread_Display(spdView, RS01)
        
        P_01005_Flag = True
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_01005_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    ReDim sValue(1)
    
    sValue(0) = "0"
    sValue(1) = Mid(cboInput(0).Text, 2, 3)
        
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01005_00", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    'Call spdDisplay(RS01)
    Call fpSpread_Display(spdView, RS01)
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01005_03", sValue(), Err_Num, Err_Dec)
    
    If RS01.EOF Then
    
    Else
        dtInput(1).Value = Format(RS01!적용일자 & "", "####-##-##")
        mskInput.Text = RS01!할인율
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Data_Display2()
    Dim i As Integer
    Dim lMemPrice As Long
    
    ReDim sValue(2)
    
    sValue(0) = "0"
    sValue(1) = Mid(cboInput(0).Text, 2, 3)
    sValue(2) = Mid(cboInput(1).Text, 2, 1) & "%"
        
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01005_05", sValue(), Err_Num, Err_Dec)
    
    If RS01.RecordCount <> 0 Then
        spdView.MaxCols = RS01.Fields.Count
        spdView.MaxRows = RS01.RecordCount
        
        'Call spdDisplay(RS01)
        Call fpSpread_Display(spdView, RS01)
    End If
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        
        spdView.Col = 3: lMemPrice = spdView.Text
        spdView.Col = 4: spdView.Text = mskInput.ClipText
        spdView.Col = 5: spdView.Text = lMemPrice - (lMemPrice * (mskInput.ClipText * 0.01))
    Next i
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
    
    If IsNull(dtInput(1).Value) Then
        MsgBox "적용일자를 선택하여 주시기 바랍니다", vbInformation
        dtInput(1).SetFocus
        Exit Sub
    End If
    
    If optSelect(0).Value = True Then
        ReDim sValue(4)
        
        sValue(0) = Mid(cboInput(0).Text, 2, 3)
        sValue(1) = Format(dtInput(1).Value, "YYYY-MM-DD")
        sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
        sValue(3) = "0"
        sValue(4) = mskInput.ClipText
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_01004_01", sValue(), Err_Num, Err_Dec)
        
        If Err_Num <> 0 Then
            MsgBox "[" & Err_Num & "] " & Err_Dec
            Exit Sub
        End If
        
        ReDim sValue(4)
        
        With spdView
            .MaxRows = 0
        
            Do While Not RS01.EOF
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                
                .Col = 1: .Text = RS01!품번 & ""
                .Col = 2: .Text = RS01!품명 & ""
                .Col = 3: .Text = RS01!금액 & ""
                .Col = 4: .Text = RS01!할인액 & ""
                
                DoEvents
                
                sValue(0) = RS01!매장코드
                sValue(1) = RS01!시작일자
                sValue(2) = RS01!품번
                sValue(3) = RS01!할인액
                sValue(4) = RS01!할인율
                
                Call ExecPro("SP_01005_01", sValue(), Err_Num, Err_Dec)
                
                If Err_Num <> 0 Then
                    MsgBox "[" & Err_Num & "] " & Err_Dec
                    Exit Sub
                End If
                
                RS01.MoveNext
            Loop
        End With
        
    Else
        ReDim sValue(4)
        
        For i = 1 To spdView.MaxRows
            spdView.Row = i
            spdView.Col = 5
            
            If Trim(spdView.Text) = "" Then
                sValue(0) = Mid(cboInput(0).Text, 2, 3)
                sValue(1) = Format(dtInput(1).Value, "YYYY-MM-DD")
                
                spdView.Col = 1: sValue(2) = spdView.Text
                spdView.Col = 4: sValue(3) = spdView.Value
                
                sValue(4) = mskInput.ClipText
                
                Call ExecPro("SP_01005_01", sValue(), Err_Num, Err_Dec)
                
                If Err_Num <> 0 Then
                    MsgBox "[" & Err_Num & "] " & Err_Dec
                    Exit Sub
                End If
            End If
        Next i
    End If

    If Err_Num = 0 Then
        MsgBox "해당되는 데이터가 정상적으로 저장이 되었습니다.", vbInformation
    End If
End Sub

Public Sub DataCancel()
    Call Data_Display
End Sub

Private Sub mskInput_LostFocus()
    On Error GoTo ErrRtn
    
    Dim i As Integer
    Dim lMemPrice As Long
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        
        spdView.Col = 3: lMemPrice = spdView.Value
        spdView.Col = 4: spdView.Value = mskInput.ClipText
        spdView.Col = 5: spdView.Value = lMemPrice * ((100 - mskInput.ClipText) * 0.01)
    Next i
    
    Exit Sub
    
ErrRtn:

End Sub

Private Sub optSelect_Click(Index As Integer, Value As Integer)
    If Index = 0 Then
        cboInput(1).Enabled = False
        
        Call Data_Display3
        
    ElseIf Index = 1 Then
        cboInput(1).Enabled = True
        
        If cboInput(1).ListIndex = -1 Then
            spdView.MaxRows = 0
        Else
            Call Data_Display2
        End If
    End If
End Sub

Private Sub spdView_Click(ByVal Col As Long, ByVal Row As Long)
    spdView.Row = Row
    spdView.Col = -1
    
    If spdView.BackColor = vbWhite Then
        spdView.BackColor = vbYellow
    ElseIf spdView.BackColor = vbYellow Then
        spdView.BackColor = vbWhite
    End If
End Sub

Private Sub Data_Display3()
    Dim i As Integer
    Dim lMemPrice As Long
    
    ReDim sValue(4)
    
    sValue(0) = "0"
    sValue(1) = Mid(cboInput(0).Text, 2, 3)
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    sValue(3) = Mid(cboInput(1).Text, 2, 1) & "%"
    sValue(4) = mskInput.ClipText
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01005_04", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    'Call spdDisplay(RS01)
    Call fpSpread_Display(spdView, RS01)

    If mskInput.ClipText = "" Then
        '
    Else
        For i = 1 To spdView.MaxRows
            spdView.Row = i
            
            spdView.Col = 3: lMemPrice = spdView.Text
            spdView.Col = 4: spdView.Text = mskInput.ClipText
            spdView.Col = 5: spdView.Text = lMemPrice - (lMemPrice * (mskInput.ClipText * 0.01))
        Next i
    End If
End Sub
