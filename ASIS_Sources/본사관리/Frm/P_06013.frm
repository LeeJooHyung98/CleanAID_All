VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_06013 
   Caption         =   " 일자별 사고 접수현황 (CS팀)"
   ClientHeight    =   11415
   ClientLeft      =   270
   ClientTop       =   2595
   ClientWidth     =   22650
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_06013.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11415
   ScaleWidth      =   22650
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   11415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   22650
      _ExtentX        =   39952
      _ExtentY        =   20135
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_06013.frx":058A
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   1
         Top             =   540
         Width           =   22620
         _ExtentX        =   39899
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1800
            TabIndex        =   14
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Format          =   64880640
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   675
            Index           =   0
            Left            =   60
            TabIndex        =   15
            Top             =   60
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   1191
            _Version        =   262144
            Caption         =   "사고접수일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   1800
            TabIndex        =   16
            Top             =   420
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Format          =   64880640
            CurrentDate     =   36686
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   6930
         _ExtentX        =   12224
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
         Caption         =   " 일자별 사고 접수현황 (CS팀)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_06013.frx":063C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Left            =   6960
         TabIndex        =   3
         Top             =   15
         Width           =   15675
         _ExtentX        =   27649
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
         PictureBackground=   "P_06013.frx":083E
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
            Picture         =   "P_06013.frx":0A40
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
            Picture         =   "P_06013.frx":0FDA
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
            Picture         =   "P_06013.frx":1574
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
            Picture         =   "P_06013.frx":1B0E
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
            Picture         =   "P_06013.frx":20A8
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
            Picture         =   "P_06013.frx":2642
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
            Picture         =   "P_06013.frx":2BDC
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
            Picture         =   "P_06013.frx":3176
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10065
         Index           =   0
         Left            =   15
         TabIndex        =   12
         Top             =   1335
         Width           =   3675
         _Version        =   524288
         _ExtentX        =   6482
         _ExtentY        =   17754
         _StockProps     =   64
         BackColorStyle  =   1
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
         MaxCols         =   2
         SpreadDesigner  =   "P_06013.frx":3710
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10065
         Index           =   1
         Left            =   3705
         TabIndex        =   13
         Top             =   1335
         Width           =   18930
         _Version        =   524288
         _ExtentX        =   33390
         _ExtentY        =   17754
         _StockProps     =   64
         BackColorStyle  =   1
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
         MaxCols         =   26
         SpreadDesigner  =   "P_06013.frx":3C2E
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_06013"
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
        Case 0: Call DataDisplay   ' 조회
'        Case 1: Call DataAdd        ' 신규
'        Case 2: Call DataSave       ' 저장
'        Case 3: Call DataDelete     ' 삭제
'        Case 4: Call DataCancel     ' 취소
'        Case 5: 'Call DataPrint      ' 인쇄
        Case 6: Call Export_Excel(P_00000.cdgExcel, spdView(1))      ' 엑셀
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
    cmdBtn(1).Enabled = False
    cmdBtn(2).Enabled = False
    cmdBtn(3).Enabled = False
    cmdBtn(4).Enabled = False
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
    
    If P_06013_Flag = False Then
        P_06013_Flag = True
        Call DataDisplay
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    P_06013_Flag = False
    
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
        
        .ColsFrozen = 1 '틀고정
    
        .Row = -1
    
        .Col = 1
        .ColWidth(1) = 5
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
    
        .Col = 2
        .ColWidth(2) = 8
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
    
        .Col = 3
        .ColWidth(3) = 20
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft

        .Col = 4
        .ColWidth(4) = 10
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter

        .Col = 5 '의류명
        .ColWidth(5) = 20
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
              
        .Col = 6 '상표
        .ColWidth(6) = 20
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
              
        .Col = 7 '색상
        .ColWidth(7) = 20
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
        
        .Col = 8
        .ColWidth(8) = 8 '지사코드
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
    
        .Col = 9
        .ColWidth(9) = 8
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
    
        .Col = 10 '구입가격
        .ColWidth(10) = 8
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 0
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignRight
        
        .Col = 11 '배상금액
        .ColWidth(11) = 8
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 0
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignRight

        .Col = 12 '보상금액
        .ColWidth(12) = 8
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 0
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignRight

        .Col = 13
        .ColWidth(13) = 10
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 14 '성명
        .ColWidth(14) = 12
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
              
        .Col = 15 '전화번호
        .ColWidth(15) = 12
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
                    
        .Col = 16 ' 휴대전화
        .ColWidth(16) = 12
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 17 ' 입고일자
        .ColWidth(17) = 12
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 18 ' 택번호
        .ColWidth(18) = 12
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
              
        .Col = 19
        .ColWidth(19) = 12
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
              
        .Col = 20 '고객출고일자
        .ColWidth(20) = 12
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
              
        .Col = 21 '품목
        .ColWidth(21) = 30
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
              
        .Col = 21 '크레임구분
        .ColWidth(21) = 12
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
              
        .Col = 21 '보상구분
        .ColWidth(21) = 12
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
              
        .Col = 21 '처리구분
        .ColWidth(21) = 12
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
              
        .Col = 21 '처리일자
        .ColWidth(21) = 12
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
    
    End With
    
    dtInput(0).Value = DateAdd("M", -1, Date)
    dtInput(1).Value = Date

    Call GetColWidth(REG_App, Me.Name, spdView)
    
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_06013_Flag = False
End Sub

Public Sub DataAdd()
    With spdView(0)
        .MaxRows = .MaxRows + 1
    
        .Row = .MaxRows
        .Col = 1
        .Action = ActionActiveCell
    End With
End Sub

Public Sub DataSave()
   
End Sub

Public Sub DataDelete()
    
  End Sub

Public Sub DataCancel()
 End Sub

Public Sub DataDisplay()

        ReDim sValue(1)
        
        sValue(0) = Format(dtInput(0).Value, "yyyy-MM-dd")
        sValue(1) = Format(dtInput(1).Value, "yyyy-MM-dd")
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_06013_00", sValue(), Err_Num, Err_Dec)
        
        With spdView(0)
            .MaxRows = 0
            .Redraw = False
                        
            Do Until RS01.EOF
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                
                .Col = 1: .Text = RS01("사고접수일자") & ""
                .Col = 2: .Text = RS01("수량") & ""
                
                RS01.MoveNext
            Loop
            RS01.Close
            Set RS01 = Nothing
            
            .Redraw = True
        End With

End Sub

Public Sub DataDisplayStoreList(sDate As String)

 
    On Error GoTo ErrRtn

    Dim i As Integer
    
    ReDim sValue(10)
    
    sValue(0) = "0"
    sValue(1) = sDate
    sValue(2) = sDate
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_06013_01", sValue(), Err_Num, Err_Dec)
    
    spdView(1).MaxCols = RS01.Fields.Count
    spdView(1).MaxRows = RS01.RecordCount
    
    Call fpSpread_Display(spdView(1), RS01)
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub spdView_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
    Dim vText   As Variant
    
    If Index = 0 Then
        
        spdView(0).GetText 1, Row, vText
        If CStr(vText) <> "" Then Call DataDisplayStoreList(CStr(vText))
        
    ElseIf Index = 1 Then
        Dim sSEQ As String
        Dim sStoreCode As String
        
        spdView(1).GetText 1, Row, vText:   sSEQ = CStr(vText)
        spdView(1).GetText 2, Row, vText:   sStoreCode = CStr(vText)
        
        If sSEQ <> "" And sStoreCode <> "" Then
            Load P_06012
            Call P_06012.Data_Display(sStoreCode, sSEQ)
            P_06012.SetFocus
        
        End If
    
    
    End If
End Sub

Private Sub spdView_LeaveCell(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If NewRow <> -1 Then
        spdView(Index).Row = Row
        spdView(Index).Col = -1
        spdView(Index).BackColor = vbWhite

        spdView(Index).Row = NewRow
        spdView(Index).Col = -1
        spdView(Index).BackColor = glbYellow
    End If
End Sub

 

