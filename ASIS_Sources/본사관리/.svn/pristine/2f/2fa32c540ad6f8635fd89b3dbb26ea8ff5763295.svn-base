VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_02016 
   Caption         =   "품목별 입고현황"
   ClientHeight    =   10230
   ClientLeft      =   3255
   ClientTop       =   5040
   ClientWidth     =   16005
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_02016.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10230
   ScaleWidth      =   16005
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   10230
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16005
      _ExtentX        =   28231
      _ExtentY        =   18045
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_02016.frx":058A
      Begin Threed.SSPanel SSPanel1 
         Height          =   420
         Left            =   15
         TabIndex        =   2
         Top             =   9795
         Width           =   15975
         _ExtentX        =   28178
         _ExtentY        =   741
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   60
            TabIndex        =   3
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "수 량 합 계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   3
            Left            =   3075
            TabIndex        =   4
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "금 액 합 계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   345
            Index           =   0
            Left            =   1530
            TabIndex        =   5
            Top             =   45
            Width           =   1185
            _Version        =   262145
            _ExtentX        =   2090
            _ExtentY        =   609
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
            StartText.y     =   5
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
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   345
            Index           =   1
            Left            =   4545
            TabIndex        =   6
            Top             =   45
            Width           =   1620
            _Version        =   262145
            _ExtentX        =   2857
            _ExtentY        =   609
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
            StartText.y     =   5
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
         Height          =   8445
         Left            =   15
         TabIndex        =   1
         Top             =   1335
         Width           =   15975
         _Version        =   524288
         _ExtentX        =   28178
         _ExtentY        =   14896
         _StockProps     =   64
         BackColorStyle  =   1
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
         GridSolid       =   0   'False
         SpreadDesigner  =   "P_02016.frx":063C
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   7
         Top             =   540
         Width           =   15975
         _ExtentX        =   28178
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   2
            Left            =   9555
            Style           =   2  '드롭다운 목록
            TabIndex        =   10
            Top             =   405
            Width           =   3015
         End
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   1
            Left            =   4815
            Style           =   2  '드롭다운 목록
            TabIndex        =   9
            Top             =   405
            Width           =   2955
         End
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   0
            Left            =   1515
            Style           =   2  '드롭다운 목록
            TabIndex        =   8
            Top             =   405
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   4815
            TabIndex        =   11
            Top             =   60
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   556
            _Version        =   393216
            Format          =   68485120
            CurrentDate     =   36686
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1515
            TabIndex        =   12
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Format          =   68485120
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   8100
            TabIndex        =   13
            Top             =   405
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "대 리 점 명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   14
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "입 고 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   7
            Left            =   60
            TabIndex        =   15
            Top             =   405
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "품  목  명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin VB.Label Label2 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "~"
            Height          =   255
            Left            =   4515
            TabIndex        =   17
            Top             =   465
            Width           =   255
         End
         Begin VB.Label Label1 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "~"
            Height          =   255
            Left            =   4515
            TabIndex        =   16
            Top             =   120
            Width           =   255
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   18
         Top             =   15
         Width           =   8370
         _ExtentX        =   14764
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
         PictureBackground=   "P_02016.frx":0AC7
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Left            =   8400
         TabIndex        =   19
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
         PictureBackground=   "P_02016.frx":0CC9
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   20
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
            Picture         =   "P_02016.frx":0ECB
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   21
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
            Picture         =   "P_02016.frx":1465
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   22
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
            Picture         =   "P_02016.frx":19FF
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   23
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
            Picture         =   "P_02016.frx":1F99
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   24
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
            Picture         =   "P_02016.frx":2533
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   25
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
            Picture         =   "P_02016.frx":2ACD
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   26
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
            Picture         =   "P_02016.frx":3067
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   27
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
            Picture         =   "P_02016.frx":3601
         End
      End
   End
End
Attribute VB_Name = "P_02016"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim RS02 As ADODB.Recordset
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

Private Sub dtInput_Change(Index As Integer)
'    Call Data_Display
End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(5).Enabled = False
    cmdBtn(6).Enabled = False
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
End Sub

'Private Sub spdDisplay(Rs As ADODB.Recordset)
'
'    Call fpSpread_Display(spdView, Rs)
'End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    With spdView
        .ColsFrozen = 1 '틀고정
        .Row = -1
        
        .Col = 1
        .ColWidth(1) = 20
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
        
        .Col = 2
        .ColWidth(2) = 8
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 3
        .ColWidth(3) = 8
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 0
        .TypeVAlign = TypeVAlignCenter
        
        .Col = 4
        .ColWidth(4) = 12
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 0
        .TypeVAlign = TypeVAlignCenter
    
        .Col = 5
        .ColWidth(5) = 14
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 2
        .TypeVAlign = TypeVAlignCenter
    
        .Col = 6
        .ColWidth(6) = 14
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignRight
    
        .Col = 7
        .ColWidth(7) = 14
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignRight
    
        .Col = 8
        .ColWidth(8) = 14
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignRight
    End With
    
    If P_02005_Flag = False Then
        Call GoodsComboAdd(cboInput(0))
        Call GoodsComboAdd(cboInput(1))
        Call AgencyComboAdd(cboInput(2))
        
        dtInput(0).Value = Date
        dtInput(1).Value = Date
        
        ReDim sValue(6)
        
        sValue(0) = "1"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_02016_00", sValue(), Err_Num, Err_Dec)
        
        spdView.MaxCols = RS01.Fields.Count
        spdView.MaxRows = RS01.RecordCount
        
        'Call spdDisplay(RS01)
        Call fpSpread_Display(spdView, RS01)
        Call GetColWidth(REG_App, Me.Name, spdView)
        
        P_02005_Flag = True
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_02005_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i As Integer
    Dim j As Integer
    Dim lTemp(3) As Single
    Dim sTemp    As String
    
    '--------------------------------------------------------------------------------------------
    '
    '--------------------------------------------------------------------------------------------
    ReDim sValue(6)
    
    sTemp = ""
    sValue(0) = "0"
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    If cboInput(0).Text = "" Then sValue(3) = "000" Else sValue(3) = Mid(cboInput(0).Text, 2, 3)
    If cboInput(0).Text = "" Then sValue(4) = "ZZZ" Else sValue(4) = Mid(cboInput(1).Text, 2, 3)
    
    sValue(5) = Mid(cboInput(2).Text, 2, 3) & "%"
    
    If sValue(3) > sValue(4) Then
        MsgBox "품목선택이 조전에 맞지 않습니다.", vbInformation
        Exit Sub
    End If
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_02016_00", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    'Call spdDisplay(RS01)
    Call fpSpread_Display(spdView, RS01)
    
    '--------------------------------------------------------------------------------------------
    '
    '--------------------------------------------------------------------------------------------
    ReDim sValue(5)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    If cboInput(0).Text = "" Then sValue(3) = "000" Else sValue(3) = Mid(cboInput(0).Text, 2, 3)
    If cboInput(0).Text = "" Then sValue(4) = "ZZZ" Else sValue(4) = Mid(cboInput(1).Text, 2, 3)
    
    sValue(5) = Mid(cboInput(2).Text, 2, 3) & "%"

    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_02016_01", sValue(), Err_Num, Err_Dec)
    
    '--------------------------------------------------------------------------------------------
    '
    '--------------------------------------------------------------------------------------------
    ReDim sValue(5)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    If cboInput(0).Text = "" Then sValue(3) = "000" Else sValue(3) = Mid(cboInput(0).Text, 2, 3)
    If cboInput(0).Text = "" Then sValue(4) = "ZZZ" Else sValue(4) = Mid(cboInput(1).Text, 2, 3)
    
    sValue(5) = Mid(cboInput(2).Text, 2, 3) & "%"

    Set RS02 = New ADODB.Recordset
    Set RS02 = ExecPro("SP_02016_02", sValue(), Err_Num, Err_Dec)
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        
        spdView.Col = 4
        lTemp(0) = Val(spdView.Value) / RS02!총금액 * 100
        
        spdView.Col = 5
        spdView.Text = lTemp(0)
    Next i
    
    For i = 1 To RS01.RecordCount
        For j = 1 To spdView.MaxRows
            spdView.Row = j
            spdView.Col = 2
            If Left(spdView.Value, 1) = RS01!상품코드 Then
                If sTemp <> RS01!상품코드 Then
                    sTemp = RS01!상품코드
                    
                    spdView.Col = 6:  spdView.Text = Format(RS01!총수량, "#,##0")
                    spdView.Col = 7:  spdView.Text = Format(RS01!총금액, "#,##0")
                    spdView.Col = 8:  spdView.Text = Format(Val(RS01!총금액) / RS02!총금액 * 100, "#,##0.00")
                    spdView.Col = -1: spdView.BackColor = &HD8FCFE
                End If
            End If
        Next j
        RS01.MoveNext
    Next i
    
    spdView.AutoCalc = True
    
    spdView.MaxRows = spdView.MaxRows + 1
    spdView.Row = spdView.MaxRows
    
    spdView.RowHidden = True
    
    txtNum(0).Value = RS02!총수량
    txtNum(1).Value = RS02!총금액
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Public Sub DataPrint()
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
'    P_00000.crPrint.Formulas(0) = "입고일자1 = '" & Format(dtInput(0).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(1) = "입고일자2 = '" & Format(dtInput(1).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(2) = "순위구분 = '" & IIf(optSelect(0).Value = True, "금액", "수량") & "'"
'    P_00000.crPrint.Formulas(3) = "대리점명 = '" & cboInput(2).Text & "'"
'    P_00000.crPrint.Formulas(4) = "품목명1 = '" & cboInput(0).Text & "'"
'    P_00000.crPrint.Formulas(5) = "품목명2 = '" & cboInput(1).Text & "'"
'
'    P_00000.crPrint.Formulas(6) = "수량합계 = '" & txtInput(0).Text & "'"
'    P_00000.crPrint.Formulas(7) = "금액합계 = '" & txtInput(1).Text & "'"
'    P_00000.crPrint.Formulas(8) = "점유율(단위)수량 = '" & txtInput(2).Text & "'"
'    P_00000.crPrint.Formulas(9) = "점유율(단위)금액 = '" & txtInput(3).Text & "'"
'    P_00000.crPrint.Formulas(10) = "점유율(전체)수량 = '" & txtInput(4).Text & "'"
'    P_00000.crPrint.Formulas(11) = "점유율(전체)금액 = '" & txtInput(5).Text & "'"
'
'    Call ReportPrint(ReportFile, "1")
End Sub


Public Sub DataScreen()
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
'    P_00000.crPrint.Formulas(0) = "입고일자1 = '" & Format(dtInput(0).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(1) = "입고일자2 = '" & Format(dtInput(1).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(2) = "순위구분 = '" & IIf(optSelect(0).Value = True, "금액", "수량") & "'"
'    P_00000.crPrint.Formulas(3) = "대리점명 = '" & cboInput(2).Text & "'"
'    P_00000.crPrint.Formulas(4) = "품목명1 = '" & cboInput(0).Text & "'"
'    P_00000.crPrint.Formulas(5) = "품목명2 = '" & cboInput(1).Text & "'"
'
'    P_00000.crPrint.Formulas(6) = "수량합계 = '" & txtInput(0).Text & "'"
'    P_00000.crPrint.Formulas(7) = "금액합계 = '" & txtInput(1).Text & "'"
'    P_00000.crPrint.Formulas(8) = "점유율(단위)수량 = '" & txtInput(2).Text & "'"
'    P_00000.crPrint.Formulas(9) = "점유율(단위)금액 = '" & txtInput(3).Text & "'"
'    P_00000.crPrint.Formulas(10) = "점유율(전체)수량 = '" & txtInput(4).Text & "'"
'    P_00000.crPrint.Formulas(11) = "점유율(전체)금액 = '" & txtInput(5).Text & "'"
'
'    Call ReportPrint(ReportFile, "2")
End Sub

Private Sub PrintDesc()
'    Dim i As Integer
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
'    For i = 1 To spdView.MaxRows - 1
'        spdView.Row = i
'
'        TempText = Left(i & Space(3), 3)
'
'        spdView.Col = 1
'        TempText = TempText & LeftH(Mid(spdView.Text, 7) & Space(12), 12)
'        spdView.Col = 2
'        TempText = TempText & RightH(Space(6) & spdView.Text, 6) & Space(1)
'        spdView.Col = 3
'        TempText = TempText & RightH(Space(10) & spdView.Text, 10) & Space(4)
'        spdView.Col = 4
'        TempText = TempText & RightH(Space(10) & spdView.Text, 10) & Space(5)
'        spdView.Col = 5
'        TempText = TempText & RightH(Space(10) & spdView.Text, 10) & Space(5)
'        spdView.Col = 6
'        TempText = TempText & RightH(Space(10) & spdView.Text, 10) & Space(5)
'        spdView.Col = 7
'        TempText = TempText & RightH(Space(10) & spdView.Text, 10) & Space(5)
'        spdView.Col = 8
'        TempText = TempText & RightH(Space(10) & spdView.Text, 10)
'
'        Print #1, TempText
'        TempText = ""
'    Next i
'
'    Close #1
End Sub
