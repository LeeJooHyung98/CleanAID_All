VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_03017 
   Caption         =   "[전사업장]  품목별 출고 기간 현황"
   ClientHeight    =   12450
   ClientLeft      =   3540
   ClientTop       =   2700
   ClientWidth     =   16500
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_03017.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12450
   ScaleWidth      =   16500
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   12450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16500
      _ExtentX        =   29104
      _ExtentY        =   21960
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_03017.frx":058A
      Begin Threed.SSPanel panInput 
         Height          =   795
         Left            =   15
         TabIndex        =   1
         Top             =   540
         Width           =   16470
         _ExtentX        =   29051
         _ExtentY        =   1402
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.TextBox txtOutMode 
            Height          =   315
            Left            =   11280
            TabIndex        =   28
            Tag             =   "a0,i0,n0,o0,p0,w0,x0"
            Top             =   420
            Width           =   3675
         End
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1530
            Locked          =   -1  'True
            Style           =   2  '드롭다운 목록
            TabIndex        =   25
            Top             =   60
            Width           =   3015
         End
         Begin VB.CommandButton cmdAllCheck 
            Caption         =   "전체 선택"
            Height          =   315
            Left            =   8280
            TabIndex        =   2
            Top             =   360
            Width           =   1305
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            TabIndex        =   3
            Top             =   405
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   57475072
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   4
            Top             =   405
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "접 수 기 간"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   4815
            TabIndex        =   5
            Top             =   405
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   57475072
            CurrentDate     =   36686
         End
         Begin XtremeSuiteControls.CheckBox chkSelect 
            Height          =   195
            Index           =   1
            Left            =   11190
            TabIndex        =   17
            Tag             =   "유통매장"
            Top             =   150
            Visible         =   0   'False
            Width           =   1185
            _Version        =   851970
            _ExtentX        =   2090
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "유통매장"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox chkSelect 
            Height          =   195
            Index           =   0
            Left            =   9690
            TabIndex        =   18
            Tag             =   "일반매장"
            Top             =   150
            Visible         =   0   'False
            Width           =   1185
            _Version        =   851970
            _ExtentX        =   2090
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "일반매장"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox chkSelect 
            Height          =   195
            Index           =   2
            Left            =   12600
            TabIndex        =   19
            Tag             =   "이마트"
            Top             =   150
            Visible         =   0   'False
            Width           =   945
            _Version        =   851970
            _ExtentX        =   1667
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "이마트"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox chkSelect 
            Height          =   195
            Index           =   3
            Left            =   13740
            TabIndex        =   20
            Tag             =   "크렌즈"
            Top             =   150
            Visible         =   0   'False
            Width           =   1095
            _Version        =   851970
            _ExtentX        =   1931
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "크렌즈"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox chkSelect2 
            Height          =   195
            Index           =   0
            Left            =   8010
            TabIndex        =   21
            Tag             =   "폐점"
            Top             =   120
            Visible         =   0   'False
            Width           =   1515
            _Version        =   851970
            _ExtentX        =   2672
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "폐점 포함"
            ForeColor       =   255
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox chkSelect 
            Height          =   195
            Index           =   4
            Left            =   14880
            TabIndex        =   22
            Tag             =   "유니트샵"
            Top             =   150
            Visible         =   0   'False
            Width           =   1185
            _Version        =   851970
            _ExtentX        =   2090
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "유니트샵"
            UseVisualStyle  =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   60
            TabIndex        =   26
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "지  사  명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   9810
            TabIndex        =   27
            Top             =   420
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "외주 코드"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin VB.Label Label1 
            Caption         =   "~"
            Height          =   225
            Left            =   4620
            TabIndex        =   6
            Top             =   465
            Width           =   225
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   7
         Top             =   15
         Width           =   8865
         _ExtentX        =   15637
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
         Caption         =   " 품목별 출고 기간 현황 (P_03017)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_03017.frx":063C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   0
         Left            =   8895
         TabIndex        =   8
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
         PictureBackground=   "P_03017.frx":083E
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   9
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
            Picture         =   "P_03017.frx":0A40
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
            Picture         =   "P_03017.frx":0FDA
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
            Picture         =   "P_03017.frx":1574
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
            Picture         =   "P_03017.frx":1B0E
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
            Picture         =   "P_03017.frx":20A8
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
            Picture         =   "P_03017.frx":2642
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
            Picture         =   "P_03017.frx":2BDC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   24
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
            Picture         =   "P_03017.frx":3176
         End
      End
      Begin FPSpreadADO.fpSpread spdView2 
         Height          =   11085
         Left            =   4470
         TabIndex        =   16
         Top             =   1350
         Width           =   12015
         _Version        =   524288
         _ExtentX        =   21193
         _ExtentY        =   19553
         _StockProps     =   64
         BackColorStyle  =   1
         ColsFrozen      =   2
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   25
         SpreadDesigner  =   "P_03017.frx":3710
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   11085
         Left            =   15
         TabIndex        =   23
         Top             =   1350
         Width           =   4440
         _Version        =   524288
         _ExtentX        =   7832
         _ExtentY        =   19553
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
         MaxCols         =   3
         ScrollBars      =   2
         SpreadDesigner  =   "P_03017.frx":4646
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_03017"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String
Dim bChkFlag    As Boolean

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cboOffice_Click()
    Dim nRow    As Long
    
    If cboOffice.ListIndex < 0 Then Exit Sub
    
    nRow = 1
    spdView.MaxRows = 0
    
    ReDim sValue(2)
    
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01001_00", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxRows = RS01.RecordCount
    
    Do Until RS01.EOF
                
        spdView.SetText 1, nRow, CVar("1")
        spdView.SetText 2, nRow, CVar(RS01!가맹점코드)
        spdView.SetText 3, nRow, CVar(RS01!가맹점명)
                        
        nRow = nRow + 1
        
        RS01.MoveNext
    Loop
    
    RS01.Close: Set RS01 = Nothing

End Sub

Private Sub chkSelect_Click(Index As Integer)
    Dim nRow    As Long
    Dim vText   As Variant
    
    With spdView
        For nRow = 1 To .MaxRows
            .GetText 4, nRow, vText
            
            ' 현재의 가맹점 종류와 선택 가맹점의 종류가 같을 경우
            If CStr(vText) = chkSelect(Index).Tag Then
                .SetText 2, nRow, CVar(chkSelect(Index).Value)
                
                ' 폐점 여부 다시 확인
                .GetText 5, nRow, vText
                If CStr(vText) = chkSelect2(0).Tag Then
                    If chkSelect2(0).Value = xtpUnchecked Then
                        .SetText 2, nRow, "0"
                    End If
                End If
                
                
            End If
        
        Next nRow
    End With
End Sub

' 제외 취소
Private Sub chkSelect2_Click(Index As Integer)
    Dim Idx     As Long
    Dim nRow    As Long
    Dim vText   As Variant
    
    With spdView
        For nRow = 1 To .MaxRows
            .GetText 5, nRow, vText
            
            ' 현재의 가맹점 종류와 선택 가맹점의 종류가 같을 경우
            If CStr(vText) = chkSelect2(Index).Tag Then
            
                ' 매장 구분을 가저온다.
                .GetText 4, nRow, vText
                For Idx = 0 To 4
                    If chkSelect(Idx).Tag = CStr(vText) Then
                        
                        ' 선택된 구분에 포함일 경우
                        If chkSelect(Idx).Value = xtpChecked Then
                            .SetText 2, nRow, CVar(chkSelect2(Index).Value)
                            
                        ' 선택이 아닐경우 무조건 미처리
                        Else
                            .SetText 2, nRow, "0"
                        End If
                            
                    End If
                Next Idx
            End If
        
        Next nRow
    End With
End Sub

Private Sub cmdAllCheck_Click()
    Dim vText   As Variant
    Dim nRow    As Long
    
    With spdView
        .EventEnabled(EventButtonClicked) = False
        
        For nRow = 1 To .MaxRows
            
            .SetText 1, nRow, CVar(IIf(cmdAllCheck.Caption = "전체 선택", "1", "0"))
        
        Next nRow
        .EventEnabled(EventButtonClicked) = True
    End With
            
    If cmdAllCheck.Caption = "전체 선택" Then
        cmdAllCheck.Caption = "전체 취소"
'        chkSelect(0).Value = xtpChecked
'        chkSelect(1).Value = xtpChecked
'        chkSelect(2).Value = xtpChecked
'        chkSelect(3).Value = xtpChecked
'        chkSelect(4).Value = xtpChecked
'        chkSelect2(0).Value = xtpChecked

    Else
        cmdAllCheck.Caption = "전체 선택"
'        chkSelect(0).Value = xtpUnchecked
'        chkSelect(1).Value = xtpUnchecked
'        chkSelect(2).Value = xtpUnchecked
'        chkSelect(3).Value = xtpUnchecked
'        chkSelect(4).Value = xtpUnchecked
'        chkSelect2(0).Value = xtpUnchecked

    End If

End Sub

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 0: Call Data_Display           ' 조회
        Case 1:                ' 신규
        Case 2:                 ' 저장
        Case 3:            ' 삭제
        Case 4:            ' 취소
        Case 5:            ' 인쇄
        Case 6: Call Export_Excel(P_00000.cdgExcel, spdView2)      ' 엑셀
        Case 7: Unload Me           ' 종료
        
        Case Else
            '
    End Select

End Sub

Private Sub ComboBox1_KeyDown(KeyCode As Integer, Shift As Integer)

End Sub

Private Sub Form_Activate()
    Dim nRow    As Long
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ") - 가맹점 입고일자 기준"

    Call SubBottonEnable(cmdBtn, "10000011")
        
    '본사 여부
    If HeadOffice = MASTER_OFFICE_CODE Then
        cboOffice.Locked = False
    Else
        cboOffice.Locked = True
    End If
    
        
    P_03017_Flag = True
    Screen.MousePointer = vbDefault

End Sub

Private Sub spdDisplay(Rs As ADODB.Recordset)

    Call fpSpread_Display(spdView, Rs)
    
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim strTemp     As String
    
    spdView2.ColsFrozen = 5
    dtInput(0).Value = Date
    dtInput(1).Value = Date
    
    
    '아래 소스 순서를 바꾸지 말것...
    Call Get_지사리스트(cboOffice)
     
    With cboOffice
        For i = 0 To .ListCount - 1
            If Mid(.List(i), 2, 4) = HeadOffice Then
                .ListIndex = i
                
                Exit For
            End If
        Next i
    End With

    
    txtOutMode.Text = GetSetting(REG_App, "P_03017", "OutMode", txtOutMode.Tag)

End Sub

Private Sub spdView_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    
    spdView.EventEnabled(EventButtonClicked) = False
    
    
    If Row = spdView.ActiveRow Then
                    
        Dim nRow    As Long
        ReDim sValue(2)
        
        If Col = 2 Then
            spdView.Row = spdView.ActiveRow
            spdView.Col = Col
            If spdView.Value = False Then
                spdView.Col = 2
                spdView.Text = ""
            
                ' 선택 내용이 지사일 경우 해당 체인점을 모두 선택 시킨다.
                spdView.Col = 1
                sValue(2) = Mid(spdView.Text, 2, 6)
                If Mid(sValue(2), 5, 1) = "]" Then
                    
                    sValue(2) = Left(sValue(2), 4)
                    For nRow = 1 To spdView.MaxRows
                        spdView.Row = nRow
                        spdView.Col = 3
                        If spdView.Text = sValue(2) Then
                            spdView.Col = 2
                            spdView.Value = "0"
                        
                        End If
                    Next nRow
                End If
        
            Else

                
                spdView.Row = Row
                spdView.Col = 2: spdView.Text = "1"
                
                
                ' 선택 내용이 지사일 경우 해당 체인점을 모두 선택 시킨다.
                spdView.Col = 1: sValue(2) = Mid(spdView.Text, 2, 6)
                If Mid(sValue(2), 5, 1) = "]" Then
                    sValue(2) = Left(sValue(2), 4)
                    
                    ' 해당 지사의 매장일 경우
                    For nRow = 1 To spdView.MaxRows
                        spdView.Row = nRow
                        spdView.Col = 3
                        If spdView.Text = sValue(2) Then
                            spdView.Col = 2: spdView.Value = "1"
                        End If
                    Next nRow
                
                End If
            End If
        End If
    End If
    
    spdView.EventEnabled(EventButtonClicked) = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cmdBtn(0).Enabled = False
    cmdBtn(1).Enabled = False
    cmdBtn(2).Enabled = False
    cmdBtn(3).Enabled = False
    cmdBtn(4).Enabled = False
    cmdBtn(5).Enabled = False
    cmdBtn(6).Enabled = False
    
    Call SaveSetting(REG_App, "P_03017", "OutMode", Trim(txtOutMode.Text))
    
    P_03017_Flag = False
End Sub

Public Sub DataSave()

End Sub

Public Sub DataAdd()

End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim nCol        As Long
    Dim nRow        As Long
    Dim SSQL2       As String
    Dim mCodeKey    As String
    Dim vText       As Variant
    Dim vTextKey    As Variant
    Dim nTotal(2)   As Long
    Dim nTotalSum(2)   As Long
    
    ReDim sValue(2)
    mCodeKey = Trim(txtOutMode.Text)
    
    SSQL2 = ""
    '-------------------------------------------------------------------------------
    ' 선택된 체인점의 내용만을 구해서 쿼리한다.
    For nRow = 0 To spdView.MaxRows
        spdView.Col = 1:  spdView.Row = nRow
        ' 매장코드만 적용한다.
        If spdView.Text = "1" Then
            spdView.Col = 2
            SSQL2 = SSQL2 & Mid(spdView.Text, 1, 6) & ","
        End If
    Next nRow
    
    SSQL2 = Trim(SSQL2)
    If Len(SSQL2) > 3 Then
        sValue(0) = Mid(SSQL2, 1, Len(SSQL2) - 1)           ' 마지막 ,을 제거한다.
    End If
        '-------------------------------------------------------------------------------
    
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    If sValue(0) = "" Then Exit Sub
    
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_03017_01_SEL", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_03017_01_SEL", sValue(), Err_Num, Err_Dec)
    End If
    
    ' 내용 출력
    spdView2.MaxRows = RS01.RecordCount
    Call fpSpread_Display(spdView2, RS01, False)
    
    With spdView2
    
        .MaxRows = .MaxRows + 6
        .SetText 2, .MaxRows - 5, CVar("일반 합계")
        .SetText 2, .MaxRows - 4, CVar("일반 비율")
        .SetText 2, .MaxRows - 3, CVar("외주 합계")
        .SetText 2, .MaxRows - 2, CVar("외주 비율")
        .SetText 2, .MaxRows - 1, CVar("전체 합계")
        .SetText 2, .MaxRows - 0, CVar("전체 비율")
        
        .Col = -1
        .Row = .MaxRows - 5: .BackColor = &HC0FFFF
        .Row = .MaxRows - 4: .BackColor = &HC0FFFF
        .Row = .MaxRows - 3: .BackColor = &HC0FFC0
        .Row = .MaxRows - 2: .BackColor = &HC0FFC0
        .Row = .MaxRows - 1: .BackColor = &HC0FFFF
        .Row = .MaxRows - 0: .BackColor = &HC0FFFF
        
        For nCol = 3 To .MaxCols
            nTotal(0) = 0: nTotal(1) = 0: nTotal(2) = 0
            
            For nRow = 1 To .MaxRows
                .GetText 1, nRow, vTextKey
                
                .GetText nCol, nRow, vText
                nTotal(0) = nTotal(0) + Val(CStr(vText))
                
                ' 외주일 경우 합계를 별도로 구한다.
                If InStr(UCase(mCodeKey), UCase(CStr(vTextKey))) > 0 Then
                    nTotal(1) = nTotal(1) + Val(CStr(vText))
                Else
                    nTotal(2) = nTotal(2) + Val(CStr(vText))
                End If
            Next nRow
            
            If nCol = 3 Then
                nTotalSum(0) = nTotal(0): nTotalSum(1) = nTotal(1): nTotalSum(2) = nTotal(2)
            End If
            
            
            .SetText nCol, .MaxRows - 5, CVar(nTotal(2))
            .SetText nCol, .MaxRows - 3, CVar(nTotal(1))
            .SetText nCol, .MaxRows - 1, CVar(nTotal(0))
    
            If nTotalSum(2) > 0 Then .SetText nCol, .MaxRows - 4, CVar(nTotal(2) / nTotalSum(2))
            If nTotalSum(1) > 0 Then .SetText nCol, .MaxRows - 2, CVar(nTotal(1) / nTotalSum(1))
            If nTotalSum(0) > 0 Then .SetText nCol, .MaxRows - 0, CVar(nTotal(0) / nTotalSum(0))
    
            .Row = .MaxRows - 4: .Col = nCol
            .CellType = CellTypePercent:    .TypePercentDecimal = ".":      .TypeVAlign = TypeVAlignCenter:     .TypeHAlign = TypeHAlignRight
    
            .Row = .MaxRows - 2: .Col = nCol
            .CellType = CellTypePercent:    .TypePercentDecimal = ".":      .TypeVAlign = TypeVAlignCenter:     .TypeHAlign = TypeHAlignRight
    
            .Row = .MaxRows - 0: .Col = nCol
            .CellType = CellTypePercent:    .TypePercentDecimal = ".":      .TypeVAlign = TypeVAlignCenter:     .TypeHAlign = TypeHAlignRight
    
        Next nCol
    
    End With
    
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub


