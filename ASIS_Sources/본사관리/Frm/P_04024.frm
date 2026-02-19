VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_04024 
   Caption         =   "월간 매출 현황 (일별 합계)"
   ClientHeight    =   9945
   ClientLeft      =   480
   ClientTop       =   2700
   ClientWidth     =   16110
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_04024.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9945
   ScaleWidth      =   16110
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   9945
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16110
      _ExtentX        =   28416
      _ExtentY        =   17542
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   0
      PaneTree        =   "P_04024.frx":058A
      Begin Threed.SSPanel SSPanel 
         Height          =   750
         Index           =   0
         Left            =   0
         TabIndex        =   15
         Top             =   9195
         Width           =   16110
         _ExtentX        =   28416
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSPanel panCaption 
            Height          =   585
            Index           =   0
            Left            =   60
            TabIndex        =   18
            Top             =   60
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   1032
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "2015-06-25일 기준 재고 수량"
            BevelInner      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   585
            Index           =   1
            Left            =   5100
            TabIndex        =   19
            Top             =   60
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   1032
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            RoundedCorners  =   0   'False
         End
         Begin VB.Label Label2 
            Caption         =   "가맹점 입고일자가 있는 경우 제외 처리"
            Height          =   255
            Left            =   8550
            TabIndex        =   21
            Top             =   390
            Width           =   5175
         End
         Begin VB.Label Label1 
            Caption         =   "2015-01-01일 기준 지사 미출고 수량"
            Height          =   195
            Left            =   8550
            TabIndex        =   20
            Top             =   120
            Width           =   6915
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   8505
         _ExtentX        =   15002
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
         PictureBackground=   "P_04024.frx":063C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   1
         Left            =   8520
         TabIndex        =   2
         Top             =   0
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
         PictureBackground=   "P_04024.frx":083E
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
            Picture         =   "P_04024.frx":0A40
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
            Picture         =   "P_04024.frx":0FDA
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
            Picture         =   "P_04024.frx":1574
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
            Picture         =   "P_04024.frx":1B0E
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
            Picture         =   "P_04024.frx":20A8
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
            Picture         =   "P_04024.frx":2642
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
            Picture         =   "P_04024.frx":2BDC
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
            Picture         =   "P_04024.frx":3176
         End
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   0
         TabIndex        =   11
         Top             =   525
         Width           =   16110
         _ExtentX        =   28416
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1245
            Style           =   2  '드롭다운 목록
            TabIndex        =   12
            Top             =   60
            Width           =   3060
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   4605
            TabIndex        =   13
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "검색년월"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   35
            Left            =   60
            TabIndex        =   14
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "사 업 장"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   5790
            TabIndex        =   16
            Top             =   60
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy-MM"
            Format          =   64225283
            CurrentDate     =   37140
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   7860
         Left            =   0
         TabIndex        =   17
         Top             =   1320
         Width           =   16110
         _Version        =   524288
         _ExtentX        =   28416
         _ExtentY        =   13864
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
         MaxCols         =   12
         SpreadDesigner  =   "P_04024.frx":3710
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_04024"
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
    cmdBtn(5).Enabled = True
    cmdBtn(6).Enabled = True
    
    '본사 여부
    If HeadOffice = MASTER_OFFICE_CODE Then
        cboOffice.Locked = False
        cboOffice.Enabled = True
    Else
        cboOffice.Locked = True
        cboOffice.Enabled = False
    End If
    
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
'        .OperationMode = OperationModeSingle
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
    
    dtInput(0).Value = Format(Date, "yyyy-mm")
    
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
    P_04024_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn
    
    Dim i As Integer
    Dim nRow    As Long
    Dim vText   As Variant
    Dim dblTotal1 As Double
    Dim dblTotal2 As Double
    Dim dblOutTotal As Double
         
    ReDim sValue(2)
    
    dblTotal1 = 0: dblTotal2 = 0
    Screen.MousePointer = vbHourglass
    
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    If sValue(0) = "0000" Then sValue(0) = "%"
    sValue(1) = Format(dtInput(0).Value, "yyyy-MM") & "-01"
    sValue(2) = Format(dtInput(0).Value, "yyyy-MM") & "-31"
    
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(sValue(0)) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_04024_00", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_04024_00", sValue(), Err_Num, Err_Dec)
    End If
        
    With spdView
        .MaxRows = 0
        .Redraw = False
        
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            dblTotal1 = dblTotal1 + RS01!매출
            dblTotal2 = dblTotal2 + RS01!입고수량
            
            .Col = 1:  .Text = RS01!매출일자 & ""               '
            .Col = 2:  .Text = RS01!매장수 & ""           '
            .Col = 3:  .Text = RS01!매출 & ""           '
            .Col = 4:  .Text = dblTotal1
            .Col = 5:  .Text = RS01!입고수량 & ""           '
            .Col = 6:  .Text = dblTotal2
            .Col = 7:  .Text = RS01!반품수량 & ""           '
            .Col = 8:  .Text = RS01!재세탁수량 & ""           '
            .Col = 9:  .Text = RS01!단가 & ""           '
            
            If .Row = 1 Then
                panCaption(0).Caption = RS01!재고적용일자 & "일 기준 재고 수량"
                panCaption(1).Caption = Format(RS01!재고수량 & "", "#,##0")
            End If
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        
        
        ReDim sValue(4)
        
        sValue(0) = Mid(cboOffice.Text, 2, 4)
        If sValue(0) = "0000" Then sValue(0) = "%"

        sValue(1) = "%"
        sValue(2) = Format(dtInput(0).Value, "YYYY-MM") & "-01"
        sValue(3) = Format(dtInput(0).Value, "YYYY-MM") & "-31"
        sValue(4) = "MASTER_DAY"
        
        If HeadOffice = MASTER_OFFICE_CODE Then
            If DBOpen_Master(HeadOffice) = False Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            
            Set RS01 = New ADODB.Recordset
            Set RS01 = ExecProMaster("SP_04001_B_01", sValue(), Err_Num, Err_Dec)
        Else
            Set RS01 = New ADODB.Recordset
            Set RS01 = ExecPro("SP_04001_B_01", sValue(), Err_Num, Err_Dec)
        End If
        
        dblOutTotal = 0
        Do While Not RS01.EOF
            ' 지사출고 수량 출력
            For nRow = 1 To .MaxRows
                .GetText 1, nRow, vText
                If CStr(vText) = RS01.Fields(1) Then
                    dblOutTotal = dblOutTotal + Val(RS01.Fields(2))
                    .SetText 10, nRow, CVar(RS01.Fields(2))
                    .SetText 11, nRow, CVar(dblOutTotal)
                    
                    .GetText 6, nRow, vText
                    .SetText 12, nRow, CVar(Val(vText) - dblOutTotal)
                    
                    Exit For
                End If
            Next nRow
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        ' 합계 출력
        Call SpreadSum(spdView, 1, 3)
        spdView.SetText 9, spdView.MaxRows, CVar(dblTotal1 / dblTotal2)
        
         
  
        
        .Redraw = True
    End With
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrRtn:
    Screen.MousePointer = vbDefault
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub DataPrint()

End Sub

Private Sub DataScreen()

End Sub

Private Sub PrintDesc()

End Sub

Private Sub DataSave()

End Sub

Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If NewRow <> -1 Then
        With spdView
            If NewRow <> -1 Then
                .Row = Row
                If (Row Mod 2) = 0 Then
                    .Col = -1
                    .BackColor = vbWhite
                Else
                    .Col = -1
                    .BackColor = vbWhite
                End If
                
                .Row = NewRow
                .Col = -1
                .BackColor = glbYellow
            End If
        End With
    End If

End Sub
