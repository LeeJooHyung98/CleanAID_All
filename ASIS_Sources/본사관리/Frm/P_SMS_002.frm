VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form P_SMS_002 
   Caption         =   "거래처별 발송 현황"
   ClientHeight    =   12270
   ClientLeft      =   2070
   ClientTop       =   3120
   ClientWidth     =   17595
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_SMS_002.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12270
   ScaleWidth      =   17595
   WindowState     =   2  '최대화
   Begin Threed.SSPanel panPrint 
      Height          =   3555
      Left            =   600
      TabIndex        =   11
      Top             =   1575
      Visible         =   0   'False
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   6271
      _Version        =   262144
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      FloodShowPct    =   -1  'True
      Begin Threed.SSCommand cmdPrint 
         Height          =   615
         Left            =   480
         TabIndex        =   12
         Top             =   2700
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   1085
         _Version        =   262144
         Caption         =   "확      인"
      End
      Begin Threed.SSFrame SSFrame6 
         Height          =   1695
         Left            =   240
         TabIndex        =   13
         Top             =   780
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   2990
         _Version        =   262144
         Caption         =   "프린트 내용 선택"
         Begin Threed.SSOption optPrint 
            Height          =   375
            Index           =   0
            Left            =   180
            TabIndex        =   14
            Top             =   240
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            _Version        =   262144
            Caption         =   "전    체"
            Value           =   -1
         End
         Begin Threed.SSOption optPrint 
            Height          =   375
            Index           =   1
            Left            =   180
            TabIndex        =   15
            Top             =   660
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            _Version        =   262144
            Caption         =   "구 분 별 (정 상)"
         End
         Begin Threed.SSOption optPrint 
            Height          =   375
            Index           =   2
            Left            =   180
            TabIndex        =   16
            Top             =   1140
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   661
            _Version        =   262144
            Caption         =   "구 분 별 (할 인)"
         End
      End
      Begin Threed.SSPanel panCaption 
         Height          =   435
         Index           =   25
         Left            =   60
         TabIndex        =   17
         Top             =   60
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   767
         _Version        =   262144
         BackColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "대리점 LIST 출력"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
   End
   Begin Threed.SSPanel panCaption 
      Height          =   8595
      Index           =   1
      Left            =   780
      TabIndex        =   18
      Top             =   2280
      Visible         =   0   'False
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   15161
      _Version        =   262144
      BevelOuter      =   1
      BevelInner      =   1
      RoundedCorners  =   0   'False
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   7605
         Left            =   360
         TabIndex        =   19
         Top             =   270
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   13414
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"P_SMS_002.frx":058A
      End
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   12270
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17595
      _ExtentX        =   31036
      _ExtentY        =   21643
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_SMS_002.frx":0E95
      Begin Threed.SSPanel panInput 
         Height          =   765
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   17565
         _ExtentX        =   30983
         _ExtentY        =   1349
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   3
            Left            =   1530
            Style           =   2  '드롭다운 목록
            TabIndex        =   4
            Top             =   420
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.CommandButton Command1 
            Caption         =   "결과 코드"
            Height          =   315
            Left            =   14700
            TabIndex        =   2
            Top             =   60
            Width           =   1275
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   1530
            TabIndex        =   3
            Top             =   60
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Format          =   57278465
            CurrentDate     =   39244
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
            Caption         =   "검색월"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   35
            Left            =   60
            TabIndex        =   6
            Top             =   420
            Visible         =   0   'False
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "지 사"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   11115
         Index           =   0
         Left            =   15
         TabIndex        =   7
         Top             =   1140
         Width           =   4530
         _Version        =   524288
         _ExtentX        =   7990
         _ExtentY        =   19606
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         SpreadDesigner  =   "P_SMS_002.frx":0F87
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   330
         Index           =   0
         Left            =   15
         TabIndex        =   8
         Top             =   795
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   582
         _Version        =   262144
         Font3D          =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "가맹점별 전송내역"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_SMS_002.frx":1440
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   330
         Index           =   1
         Left            =   4560
         TabIndex        =   9
         Top             =   795
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   582
         _Version        =   262144
         Font3D          =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "전송일자"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_SMS_002.frx":18A2
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   330
         Index           =   2
         Left            =   7035
         TabIndex        =   10
         Top             =   795
         Width           =   10545
         _ExtentX        =   18600
         _ExtentY        =   582
         _Version        =   262144
         Font3D          =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "전송 메시지 내역"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_SMS_002.frx":1D04
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   11115
         Index           =   1
         Left            =   4560
         TabIndex        =   20
         Top             =   1140
         Width           =   2460
         _Version        =   524288
         _ExtentX        =   4339
         _ExtentY        =   19606
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         SpreadDesigner  =   "P_SMS_002.frx":2166
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   11115
         Index           =   2
         Left            =   7035
         TabIndex        =   21
         Top             =   1140
         Width           =   10545
         _Version        =   524288
         _ExtentX        =   18600
         _ExtentY        =   19606
         _StockProps     =   64
         BackColorStyle  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         SpreadDesigner  =   "P_SMS_002.frx":261F
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_SMS_002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Dim P_SMS002_Flag As Boolean

Dim sPrintOption As String

Private Sub Data_Display()
    On Error GoTo ErrRtn

    ReDim sValue(2)

    sValue(0) = "0"
    sValue(1) = Mid(Trim(cboInput(3).Text), 2, 4)
    sValue(2) = Format(DTPicker1.Value, "yyyyMM")
    ' 대리점 정보
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("PRO_SMS_002_11", sValue(), Err_Num, Err_Dec)

    spdView(0).MaxCols = RS01.Fields.Count
    spdView(0).MaxRows = RS01.RecordCount

    Call spdDisplay1(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView(0))
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub spdDisplay1(Rs As ADODB.Recordset)

    Call fpSpread_Display(spdView(0), Rs)

End Sub

Private Sub spdDisplay2(Rs As ADODB.Recordset)
    Call fpSpread_Display(spdView(1), Rs)
End Sub

Private Sub spdDisplay3(Rs As ADODB.Recordset)
    Call fpSpread_Display(spdView(2), Rs)
End Sub

Private Sub cboInput_Click(Index As Integer)
    Call Data_Display
End Sub

Private Sub cmdPrint_Click()
'    Call DataScreen2
'    panPrint.Visible = False
End Sub

Private Sub Command1_Click()
    ' 결과 코드 보기
    panCaption(1).ZOrder 0
    panCaption(1).Visible = Not panCaption(1).Visible
End Sub

Private Sub Form_Activate()
'    cmdBtn(0).Enabled = True
'    cmdBtn(1).Enabled = False
'    cmdBtn(2).Enabled = False
'    cmdBtn(3).Enabled = False
'    cmdBtn(4).Enabled = False
'    cmdBtn(5).Enabled = False
'    cmdBtn(6).Enabled = True
'
'    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    

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
        .OperationMode = OperationModeSingle
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    
    
        .ColsFrozen = 1  '틀고정
        .Row = -1

        .Col = 1
        .ColWidth(1) = 5
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter

        .Col = 2
        .ColWidth(2) = 14
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft

        .Col = 3
        .ColWidth(3) = 8
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 0
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignRight

        .Col = 4
        .ColWidth(4) = 8
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 0
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignRight
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
        .OperationMode = OperationModeSingle
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    
    
        .ColsFrozen = 1  '틀고정
        .Row = -1

        .Col = 1
        .ColWidth(1) = 10
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter

        .Col = 2
        .ColWidth(2) = 8
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 0
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignRight
    End With


    With spdView(2)
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
    
        .ColsFrozen = 1  '틀고정
        .Row = -1

        .Col = 1
        .ColWidth(1) = 10
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter

        .Col = 2
        .ColWidth(2) = 10
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignRight
        
        .Col = 3
        .ColWidth(3) = 14
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter

        .Col = 4
        .ColWidth(4) = 14
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
    
        .Col = 5
        .ColWidth(5) = 20
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
    
        .Col = 6
        .ColWidth(6) = 5
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
    End With


    If P_SMS002_Flag = False Then
        ' Combo BOX의 내역을 채운다.
        'Call ComboAdd

        If Store.Code = MASTER_OFFICE_CODE Then
            panCaption(35).Visible = True
            cboInput(3).Visible = True
        End If
        
         Call Get_지사리스트(cboInput(3))

    
        ReDim sValue(2)

        sValue(0) = "0"
        sValue(1) = Store.Code
        sValue(2) = Format(DTPicker1.Value, "yyyyMM")

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("PRO_SMS_002_11", sValue(), Err_Num, Err_Dec)

        spdView(0).MaxCols = RS01.Fields.Count
        spdView(0).MaxRows = RS01.RecordCount

        Call spdDisplay1(RS01)
        Call GetColWidth(REG_App, Me.Name, spdView(0))
        'Call GetColWidth(REG_App, Me.Name, spdView2)

        P_SMS002_Flag = True
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_SMS002_Flag = False
End Sub



Private Sub Data_Display2()
    Dim i As Integer
   
    ReDim sValue(3)

    sValue(0) = "0"
    sValue(1) = Mid(Trim(cboInput(3).Text), 2, 4)
    
    spdView(0).Row = spdView(0).ActiveRow
    spdView(0).Col = 1
    sValue(2) = spdView(0).Text
    
    sValue(3) = Format(DTPicker1.Value, "yyyyMM")

    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("PRO_SMS_002_01", sValue(), Err_Num, Err_Dec)

    spdView(1).MaxCols = RS01.Fields.Count
    spdView(1).MaxRows = RS01.RecordCount

    Call spdDisplay2(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView(1))
    
    ' 추가 상세 정보를 조회 한다.
    Call Data_Display3
    
    Set RS01 = Nothing

End Sub

Private Sub Data_Display3()
    Dim lRow As Long

    ReDim sValue(3)

    If spdView(1).MaxRows <= 0 Then Exit Sub

    sValue(0) = "0"
    sValue(1) = Mid(Trim(cboInput(3).Text), 2, 4)
    
    spdView(0).Row = spdView(0).ActiveRow
    spdView(0).Col = 1
    sValue(2) = spdView(0).Text
    
    spdView(1).Row = spdView(1).ActiveRow
    spdView(1).Col = 1
    sValue(3) = spdView(1).Text
    sValue(3) = Format(sValue(3), "YYYY-MM-DD")

    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("PRO_SMS_002_02", sValue(), Err_Num, Err_Dec)

    spdView(2).MaxCols = RS01.Fields.Count
    spdView(2).MaxRows = RS01.RecordCount

    Call spdDisplay3(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView(2))

    ' 비정상적인 내용은 별도로 표시한다.
    For lRow = 0 To spdView(2).MaxRows
        With spdView(2)
            .Col = 6:   .Row = lRow
            If Trim(.Text) <> "06" Then
                .Col = -1:  .BackColor = vbRed
            End If
        End With
    
    Next lRow
    Set RS01 = Nothing
End Sub

Public Sub DataAdd()


End Sub

Public Sub DataCancel()
    'Call Data_Display2
End Sub

Public Sub DataDelete()
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

Public Sub DataSave()

End Sub

Public Sub DataPrint()
'    Dim ReportFP As String
'    Dim ReportFile As String
'
'    ReportFP = GetIniStr("REPORT", "FilePath", "", m_iniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    P_00000.crPrint.StoredProcParam(0) = "0"
'    P_00000.crPrint.StoredProcParam(1) = txtInput(1).Text
'    P_00000.crPrint.WindowTitle = Me.Caption
'
'    Call ReportPrint(ReportFile, "1")
End Sub


Public Sub DataScreen()
    panPrint.Visible = True

    sPrintOption = "2"
End Sub

Private Sub spdView_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
    Select Case Index
        Case 0
            spdView(1).MaxRows = 0
            spdView(2).MaxRows = 0
            Call Data_Display2
            
        Case 1
            spdView(2).MaxRows = 0
            Call Data_Display3
        
        Case 2
        
        
        Case Else
    End Select
        
End Sub
