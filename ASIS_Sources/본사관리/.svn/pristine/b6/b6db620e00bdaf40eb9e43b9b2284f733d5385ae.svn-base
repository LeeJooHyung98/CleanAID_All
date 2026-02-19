VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_06014 
   Caption         =   " 사고 유형 분석 (CS팀)"
   ClientHeight    =   11415
   ClientLeft      =   -975
   ClientTop       =   2400
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
   Icon            =   "P_06014.frx":0000
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
      PaneTree        =   "P_06014.frx":058A
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   1
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
         Caption         =   " 사고 유형 분석 (CS팀)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_06014.frx":061C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Left            =   6960
         TabIndex        =   2
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
         PictureBackground=   "P_06014.frx":081E
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
            Picture         =   "P_06014.frx":0A20
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
            Picture         =   "P_06014.frx":0FBA
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
            Picture         =   "P_06014.frx":1554
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
            Picture         =   "P_06014.frx":1AEE
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
            Picture         =   "P_06014.frx":2088
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
            Picture         =   "P_06014.frx":2622
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
            Picture         =   "P_06014.frx":2BBC
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
            Picture         =   "P_06014.frx":3156
         End
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   11
         Top             =   540
         Width           =   22620
         _ExtentX        =   39899
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1560
            TabIndex        =   12
            Text            =   "cboOffice"
            Top             =   30
            Width           =   3735
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   7020
            TabIndex        =   13
            Top             =   30
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   64159744
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   5550
            TabIndex        =   14
            Top             =   30
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "검 색 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   10365
            TabIndex        =   15
            Top             =   30
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   64159744
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   3
            Left            =   90
            TabIndex        =   16
            Top             =   30
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "사업장 명칭"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin VB.Label Label1 
            Alignment       =   2  '가운데 맞춤
            Caption         =   "~"
            Height          =   195
            Left            =   10065
            TabIndex        =   17
            Top             =   90
            Width           =   255
         End
      End
      Begin XtremeSuiteControls.TabControl TabControl1 
         Height          =   10065
         Left            =   15
         TabIndex        =   18
         Top             =   1335
         Width           =   22620
         _Version        =   851970
         _ExtentX        =   39899
         _ExtentY        =   17754
         _StockProps     =   68
         Appearance      =   3
         Color           =   64
         PaintManager.BoldSelected=   -1  'True
         PaintManager.OneNoteColors=   -1  'True
         PaintManager.HotTracking=   -1  'True
         PaintManager.ButtonMargin=   "10,10,10,10"
         ItemCount       =   2
         SelectedItem    =   1
         Item(0).Caption =   "지사별"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "TabControlPage1"
         Item(1).Caption =   "매장별"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "TabControlPage2"
         Begin XtremeSuiteControls.TabControlPage TabControlPage2 
            Height          =   9405
            Left            =   30
            TabIndex        =   19
            Top             =   630
            Width           =   22560
            _Version        =   851970
            _ExtentX        =   39793
            _ExtentY        =   16589
            _StockProps     =   1
            Page            =   1
            Begin SSSplitter.SSSplitter SSSplitter 
               Height          =   9405
               Index           =   1
               Left            =   0
               TabIndex        =   20
               Top             =   0
               Width           =   22560
               _ExtentX        =   39793
               _ExtentY        =   16589
               _Version        =   262144
               AutoSize        =   1
               SplitterBarWidth=   1
               SplitterBarJoinStyle=   0
               SplitterBarAppearance=   0
               PaneTree        =   "P_06014.frx":36F0
               Begin FPSpreadADO.fpSpread spdView 
                  Height          =   9345
                  Index           =   1
                  Left            =   30
                  TabIndex        =   24
                  Top             =   30
                  Width           =   22500
                  _Version        =   524288
                  _ExtentX        =   39688
                  _ExtentY        =   16484
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
                  MaxCols         =   30
                  SpreadDesigner  =   "P_06014.frx":3722
                  HighlightHeaders=   1
                  HighlightStyle  =   1
               End
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage1 
            Height          =   9405
            Left            =   -69970
            TabIndex        =   21
            Top             =   630
            Visible         =   0   'False
            Width           =   22560
            _Version        =   851970
            _ExtentX        =   39793
            _ExtentY        =   16589
            _StockProps     =   1
            Page            =   0
            Begin SSSplitter.SSSplitter SSSplitter 
               Height          =   9405
               Index           =   2
               Left            =   0
               TabIndex        =   22
               Top             =   0
               Width           =   22560
               _ExtentX        =   39793
               _ExtentY        =   16589
               _Version        =   262144
               AutoSize        =   1
               SplitterBarWidth=   1
               SplitterBarJoinStyle=   0
               SplitterBarAppearance=   0
               PaneTree        =   "P_06014.frx":468F
               Begin FPSpreadADO.fpSpread spdView 
                  Height          =   9345
                  Index           =   0
                  Left            =   30
                  TabIndex        =   23
                  Top             =   30
                  Width           =   22500
                  _Version        =   524288
                  _ExtentX        =   39688
                  _ExtentY        =   16484
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
                  MaxCols         =   30
                  SpreadDesigner  =   "P_06014.frx":46C1
                  HighlightHeaders=   1
                  HighlightStyle  =   1
               End
            End
         End
      End
   End
End
Attribute VB_Name = "P_06014"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

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
        Case 0:
            Select Case TabControl1.SelectedItem
                Case 0:     DataDisplay   ' 조회
                Case 1:     DataDisplayStore ("")   ' 조회
            End Select
'        Case 1: Call DataAdd        ' 신규
'        Case 2: Call DataSave       ' 저장
'        Case 3: Call DataDelete     ' 삭제
'        Case 4: Call DataCancel     ' 취소
'        Case 5: 'Call DataPrint      ' 인쇄
        Case 6: Call Export_Excel(P_00000.cdgExcel, IIf(TabControl1.SelectedItem = 0, spdView(0), spdView(1)))      ' 엑셀
        Case 7: Unload Me           ' 종료
'        Case 8: Call DataSMSSave
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
    cmdBtn(6).Enabled = True
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
    '본사 여부
    If HeadOffice = MASTER_OFFICE_CODE Then
        cboOffice.Locked = False
    Else
        cboOffice.Locked = True
    End If
 
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
        .OperationMode = OperationModeRow
        
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
        .OperationMode = OperationModeRow
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With

    dtInput(0).Value = Format(Date, "yyyy-MM") & "-01"
    dtInput(1).Value = Format(Date, "yyyy-MM-dd")

    TabControl1.SelectedItem = 0
    
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
    P_06010_Flag = False
End Sub


Public Sub DataSave()
  
End Sub

Public Sub DataDelete()
   
End Sub

Public Sub DataCancel()
 End Sub

Public Sub DataDisplay()
    On Error GoTo ERR_RTN
    
        ReDim sValue(2)
        
        sValue(0) = Mid(cboOffice.Text, 2, 4)
        sValue(0) = IIf(sValue(0) = "", "0000", sValue(0))
        sValue(0) = IIf(sValue(0) = "0000", "%", sValue(0))

        sValue(1) = Format(dtInput(0).Value, "yyyy-MM-dd")
        sValue(2) = Format(dtInput(1).Value, "yyyy-MM-dd")
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_06014_00", sValue(), Err_Num, Err_Dec)
        
        With spdView(0)
            .MaxRows = 0
            .Redraw = False

            Do Until RS01.EOF
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows

                .Col = 1: .Text = RS01(0) & ""
                .Col = 2: .Text = RS01(1) & ""
                
                .Col = 3: .Text = RS01(2) & ""
                .Col = 4: .Text = RS01(3) & ""
                .Col = 5: .Text = RS01(4) & ""
                .Col = 6: .Text = RS01(5) & ""
                .Col = 7: .Text = RS01(6) & ""
                .Col = 8: .Text = RS01(7) & ""
                .Col = 9: .Text = RS01(8) & ""
                .Col = 10: .Text = RS01(9) & ""
                .Col = 11: .Text = RS01(10) & ""
                .Col = 12: .Text = RS01(11) & ""
                .Col = 13: .Text = RS01(12) & ""
                .Col = 14: .Text = RS01(13) & ""
                .Col = 15: .Text = RS01(14) & ""
                
                .Col = 17: .Text = RS01(15) & ""
                .Col = 18: .Text = RS01(16) & ""
                .Col = 19: .Text = RS01(17) & ""
                .Col = 20: .Text = RS01(18) & ""
                .Col = 21: .Text = RS01(19) & ""
                .Col = 22: .Text = RS01(20) & ""
                .Col = 23: .Text = RS01(21) & ""
                .Col = 24: .Text = RS01(22) & ""
                .Col = 25: .Text = RS01(23) & ""
                .Col = 26: .Text = RS01(24) & ""
                .Col = 27: .Text = RS01(25) & ""
                .Col = 28: .Text = RS01(26) & ""
                .Col = 29: .Text = RS01(27) & ""
                

                RS01.MoveNext
            Loop
            RS01.Close
            Set RS01 = Nothing

            .Redraw = True
        End With
        
    Dim nCol As Long
    
    For nCol = 3 To spdView(0).MaxCols - 1
        
        Select Case nCol
        
            Case 3: Call SpreadSum(spdView(0), 2, 3)
            Case 16
            Case Else
                Call SpreadSum(spdView(0), -1, nCol)
        End Select
        
    Next nCol
    
    Call DataSumPercent
    Exit Sub
    

ERR_RTN:
    PanelsMsg Err.Description
End Sub
 


Public Sub DataDisplayStore(sCode As String)
    Dim vText   As Variant
    
    On Error GoTo ERR_RTN

    ReDim sValue(2)
        
    If sCode <> "" Then
        sValue(0) = CStr(sCode)
    Else
        sValue(0) = CStr(vText)
        sValue(0) = IIf(sValue(0) = "", "0000", sValue(0))
        sValue(0) = IIf(sValue(0) = "0000", "%", sValue(0))
    End If
        sValue(1) = Format(dtInput(0).Value, "yyyy-MM-dd")
        sValue(2) = Format(dtInput(1).Value, "yyyy-MM-dd")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_06014_01", sValue(), Err_Num, Err_Dec)
    
    With spdView(1)
        .MaxRows = 0
        .Redraw = False

        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows

                .Col = 1: .Text = RS01(0) & ""
                .Col = 2: .Text = RS01(1) & ""
                
                .Col = 3: .Text = RS01(2) & ""
                .Col = 4: .Text = RS01(3) & ""
                .Col = 5: .Text = RS01(4) & ""
                .Col = 6: .Text = RS01(5) & ""
                .Col = 7: .Text = RS01(6) & ""
                .Col = 8: .Text = RS01(7) & ""
                .Col = 9: .Text = RS01(8) & ""
                .Col = 10: .Text = RS01(9) & ""
                .Col = 11: .Text = RS01(10) & ""
                .Col = 12: .Text = RS01(11) & ""
                .Col = 13: .Text = RS01(12) & ""
                .Col = 14: .Text = RS01(13) & ""
                .Col = 15: .Text = RS01(14) & ""
                
                .Col = 17: .Text = RS01(15) & ""
                .Col = 18: .Text = RS01(16) & ""
                .Col = 19: .Text = RS01(17) & ""
                .Col = 20: .Text = RS01(18) & ""
                .Col = 21: .Text = RS01(19) & ""
                .Col = 22: .Text = RS01(20) & ""
                .Col = 23: .Text = RS01(21) & ""
                .Col = 24: .Text = RS01(22) & ""
                .Col = 25: .Text = RS01(23) & ""
                .Col = 26: .Text = RS01(24) & ""
                .Col = 27: .Text = RS01(25) & ""
                .Col = 28: .Text = RS01(26) & ""
                .Col = 29: .Text = RS01(27) & ""

            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing

        .Redraw = True
    
    End With
    
    Dim nCol As Long
    
    For nCol = 3 To spdView(1).MaxCols - 1
        
        Select Case nCol
        
            Case 3: Call SpreadSum(spdView(1), 2, 3)
            Case 16
            Case Else
                Call SpreadSum(spdView(1), -1, nCol)
        End Select
        
    Next nCol
        

    Call DataSumPercent
    Exit Sub
    

ERR_RTN:
    PanelsMsg Err.Description

End Sub
 

Private Sub spdView_DblClick(Index As Integer, ByVal Col As Long, ByVal Row As Long)
    Select Case Index
        Case 0
            TabControl1.SelectedItem = 1
            spdView(1).SetFocus
            
            Dim vText As Variant
            
            spdView(0).GetText 1, spdView(0).ActiveRow, vText
            
            Call DataDisplayStore(CStr(vText))
            
        
    End Select
End Sub





Public Sub DataSumPercent()
    Dim vText   As Variant
    Dim nRow    As Long
    
    ReDim sValue(1)
    
    With spdView(TabControl1.SelectedItem)
        
        .GetText 15, .MaxRows, vText:   sValue(0) = CStr(vText)
        .GetText 29, .MaxRows, vText:   sValue(1) = CStr(vText)
        
        
        For nRow = 1 To .MaxRows
            .GetText 15, nRow, vText
            If Val(sValue(0)) > 0 Then .SetText 16, nRow, Val(vText) / Val(sValue(0))
        
        
            .GetText 29, nRow, vText
            If Val(sValue(1)) > 0 Then .SetText 30, nRow, Val(vText) / Val(sValue(1))
        Next
        
    End With

End Sub
 
