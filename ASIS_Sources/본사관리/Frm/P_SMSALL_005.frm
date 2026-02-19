VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_SMSALL_5 
   Caption         =   "SMS 발송자 등록"
   ClientHeight    =   11415
   ClientLeft      =   1395
   ClientTop       =   3525
   ClientWidth     =   14565
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_SMSALL_005.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11415
   ScaleWidth      =   14565
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   11415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14565
      _ExtentX        =   25691
      _ExtentY        =   20135
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_SMSALL_005.frx":058A
      Begin XtremeSuiteControls.TabControl TabControl1 
         Height          =   10065
         Left            =   15
         TabIndex        =   12
         Top             =   1335
         Width           =   14535
         _Version        =   851970
         _ExtentX        =   25638
         _ExtentY        =   17754
         _StockProps     =   68
         Appearance      =   3
         Color           =   4
         PaintManager.BoldSelected=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         PaintManager.LargeIcons=   -1  'True
         ItemCount       =   2
         SelectedItem    =   1
         Item(0).Caption =   "마트 협력인 "
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "TabControlPage1"
         Item(1).Caption =   "상담자 등록 "
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "TabControlPage2"
         Begin XtremeSuiteControls.TabControlPage TabControlPage2 
            Height          =   9435
            Left            =   30
            TabIndex        =   14
            Top             =   600
            Width           =   14475
            _Version        =   851970
            _ExtentX        =   25532
            _ExtentY        =   16642
            _StockProps     =   1
            Page            =   1
            Begin SSSplitter.SSSplitter SSSplitter2 
               Height          =   9435
               Left            =   0
               TabIndex        =   15
               Top             =   0
               Width           =   14475
               _ExtentX        =   25532
               _ExtentY        =   16642
               _Version        =   262144
               AutoSize        =   1
               PaneTree        =   "P_SMSALL_005.frx":061C
               Begin FPSpreadADO.fpSpread spdView 
                  Height          =   9375
                  Index           =   1
                  Left            =   30
                  TabIndex        =   16
                  Top             =   30
                  Width           =   14415
                  _Version        =   524288
                  _ExtentX        =   25426
                  _ExtentY        =   16536
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
                  MaxCols         =   8
                  SpreadDesigner  =   "P_SMSALL_005.frx":064E
                  HighlightHeaders=   1
                  HighlightStyle  =   1
               End
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage1 
            Height          =   9435
            Left            =   -69970
            TabIndex        =   13
            Top             =   600
            Visible         =   0   'False
            Width           =   14475
            _Version        =   851970
            _ExtentX        =   25532
            _ExtentY        =   16642
            _StockProps     =   1
            Page            =   0
            Begin SSSplitter.SSSplitter SSSplitter3 
               Height          =   9435
               Left            =   0
               TabIndex        =   17
               Top             =   0
               Width           =   14475
               _ExtentX        =   25532
               _ExtentY        =   16642
               _Version        =   262144
               AutoSize        =   1
               PaneTree        =   "P_SMSALL_005.frx":0CF9
               Begin FPSpreadADO.fpSpread spdView 
                  Height          =   9375
                  Index           =   0
                  Left            =   30
                  TabIndex        =   18
                  Top             =   30
                  Width           =   14415
                  _Version        =   524288
                  _ExtentX        =   25426
                  _ExtentY        =   16536
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
                  MaxCols         =   6
                  SpreadDesigner  =   "P_SMSALL_005.frx":0D2B
                  HighlightHeaders=   1
                  HighlightStyle  =   1
               End
            End
         End
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   1
         Top             =   540
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
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
         Caption         =   " SMS 발송자 등록 (P_SMSALL_5)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_SMSALL_005.frx":1342
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
         PictureBackground=   "P_SMSALL_005.frx":1544
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
            Picture         =   "P_SMSALL_005.frx":1746
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
            Picture         =   "P_SMSALL_005.frx":1CE0
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
            Picture         =   "P_SMSALL_005.frx":227A
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
            Picture         =   "P_SMSALL_005.frx":2814
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
            Picture         =   "P_SMSALL_005.frx":2DAE
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
            Picture         =   "P_SMSALL_005.frx":3348
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
            Picture         =   "P_SMSALL_005.frx":38E2
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
            Picture         =   "P_SMSALL_005.frx":3E7C
         End
      End
   End
End
Attribute VB_Name = "P_SMSALL_5"
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
        Case 0
            Select Case TabControl1.SelectedItem
                Case 0: Call DataDisplay1      ' 조회
                Case 1: Call DataDisplay2      ' 조회
            End Select
        Case 1: Call DataAdd        ' 신규
        Case 2
            Select Case TabControl1.SelectedItem
                Case 0: Call DataSave1      ' 저장
                Case 1: Call DataSave2      ' 저장
            End Select
        Case 3: Call DataDelete
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
    cmdBtn(4).Enabled = False
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
    If P_SMSALL_5_Flag = False Then
        Call DataDisplay1
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
        .OperationMode = OperationModeNormal
        
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
        .OperationMode = OperationModeNormal
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With

    TabControl1.SelectedItem = 0
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_SMSALL_5_Flag = False
End Sub

Public Sub DataAdd()
    Dim nidx    As Integer
    
    nidx = TabControl1.SelectedItem
    
    spdView(nidx).MaxRows = spdView(nidx).MaxRows + 1
    
    spdView(nidx).Row = spdView(nidx).MaxRows
    spdView(nidx).Col = 1
    spdView(nidx).Action = ActionActiveCell
End Sub

Public Sub DataDelete()
    Dim Idx As Integer
    
    ReDim sValue(2)
    
    Idx = TabControl1.SelectedItem
    
    sValue(0) = Store.Code
    sValue(1) = IIf(Idx = 0, "01", "02")
    
    spdView(Idx).Row = spdView(Idx).ActiveRow: spdView(Idx).Col = 1
    sValue(2) = spdView(Idx).Text

    If MsgBox("[" & sValue(2) & "]해당되는 데이터를 삭제하시겠습니까?", vbYesNo + vbInformation + vbDefaultButton2, "데이터 삭제") = vbYes Then
        
        Call ExecPro("SP_M_SMSALL_005_02", sValue(), Err_Num, Err_Dec)

        Select Case Idx
            Case 0:   Call DataDisplay1
            Case 1:   Call DataDisplay2
        End Select
    
    End If
End Sub

Public Sub DataCancel()
'    If optSelect(0).Value = True Then
'        Call optSelect_Click(0, True)
'    ElseIf optSelect(1).Value = True Then
'        Call optSelect_Click(1, True)
'    End If
End Sub

Public Sub DataDisplay1()

    ReDim sValue(1)
    
    sValue(0) = Store.Code
    sValue(1) = "01"
    
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_SMSALL_005_00", sValue(), Err_Num, Err_Dec)
    
    With spdView(0)
        .MaxRows = 0
        .Redraw = False
                    
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = RS01("담당자코드") & ""
            .Col = 2: .Text = RS01("구분") & ""
            .Col = 3: .Text = RS01("매장명") & ""
            .Col = 4: .Text = RS01("성명") & ""
            .Col = 5: .Text = RS01("연락처") & ""
            .Col = 6: .Text = RS01("비고") & ""
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
    
End Sub


Public Sub DataDisplay2()

    ReDim sValue(1)
    
    sValue(0) = Store.Code
    sValue(1) = "02"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_SMSALL_005_00", sValue(), Err_Num, Err_Dec)
    
    With spdView(1)
        .MaxRows = 0
        .Redraw = False
                    
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = RS01("코드") & ""
            .Col = 2: .Text = RS01("지역") & ""
            .Col = 3: .Text = RS01("최초상담") & ""
            .Col = 4: .Text = RS01("성명") & ""
            .Col = 5: .Text = RS01("연락처") & ""
            .Col = 6: .Text = RS01("점포상황주소") & ""
            .Col = 7: .Text = RS01("전화상담자") & ""
            .Col = 8: .Text = RS01("비고") & ""
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
    
End Sub

Public Sub DataSave1()
    Dim i As Integer
    ReDim sValue(6)
    Dim sTel(2) As String
    
    For i = 1 To spdView(0).DataRowCnt
        spdView(0).Row = i
        
                            sValue(0) = Store.Code
        spdView(0).Col = 1:    sValue(1) = spdView(0).Text
        spdView(0).Col = 2:    sValue(2) = spdView(0).Text
        spdView(0).Col = 3:    sValue(3) = spdView(0).Text
        spdView(0).Col = 4:    sValue(4) = spdView(0).Text
        spdView(0).Col = 5:    sValue(5) = spdView(0).Text
        spdView(0).Col = 6:    sValue(6) = spdView(0).Text
        
        If Trim(sValue(1)) = "" Then
            MsgBox "코드를 확인해주세요.", vbInformation, "확인"
            Exit Sub
        End If
        
        If Trim(sValue(5)) <> "" Then
            If CheckMobileNumber(sValue(5), sTel) = False Then
                MsgBox "[" & Trim(sValue(5)) & "]의 전화번호를 확인하여 주십시요", vbInformation, "확인"
                Exit Sub
            End If
        End If
    Next i
    
    For i = 1 To spdView(0).DataRowCnt
        spdView(0).Row = i
        
                            sValue(0) = Store.Code
        spdView(0).Col = 1:    sValue(1) = spdView(0).Text
        spdView(0).Col = 2:    sValue(2) = spdView(0).Text
        spdView(0).Col = 3:    sValue(3) = spdView(0).Text
        spdView(0).Col = 4:    sValue(4) = spdView(0).Text
        spdView(0).Col = 5:    sValue(5) = spdView(0).Text
        spdView(0).Col = 6:    sValue(6) = spdView(0).Text
        
        If Trim(sValue(0)) = "" Then
            Exit Sub
        End If
        
        Call ExecPro("SP_M_SMSALL_005_01", sValue(), Err_Num, Err_Dec)
    Next i

    If Err_Num = 0 Then
        MsgBox "해당되는 데이터가 정상적으로 저장이 되었습니다.", vbInformation
    End If
End Sub

Public Sub DataSave2()
    Dim i As Integer
    ReDim sValue(8)
    Dim sTel(2) As String
    
    For i = 1 To spdView(1).DataRowCnt
        spdView(1).Row = i
        
                               sValue(0) = Store.Code
        spdView(1).Col = 1:    sValue(1) = spdView(1).Text
        spdView(1).Col = 2:    sValue(2) = spdView(1).Text
        spdView(1).Col = 3:    sValue(3) = spdView(1).Text
        spdView(1).Col = 4:    sValue(4) = spdView(1).Text
        spdView(1).Col = 5:    sValue(5) = spdView(1).Text
        spdView(1).Col = 6:    sValue(6) = spdView(1).Text
        spdView(1).Col = 7:    sValue(7) = spdView(1).Text
        
        If Trim(sValue(1)) = "" Then
            MsgBox CStr(i) & "번째 코드를 확인해주세요.", vbInformation, "확인"
            Exit Sub
        End If
        
        If Trim(sValue(4)) = "" Then
            MsgBox CStr(i) & "번째 고객명을 확인해주세요.", vbInformation, "확인"
            Exit Sub
        End If
        
'        If Trim(sValue(5)) <> "" Then
'            If CheckMobileNumber(sValue(5), sTel) = False Then
'                MsgBox CStr(i) & "번째 [" & Trim(sValue(5)) & "]의 연락처를 확인하여 주십시요", vbInformation, "확인"
'                Exit Sub
'            End If
'        End If
    Next i
    
    For i = 1 To spdView(1).DataRowCnt
        spdView(1).Row = i
        
                               sValue(0) = Store.Code
        spdView(1).Col = 1:    sValue(1) = spdView(1).Text
        spdView(1).Col = 2:    sValue(2) = spdView(1).Text
        spdView(1).Col = 3:    sValue(3) = spdView(1).Text
        spdView(1).Col = 4:    sValue(4) = spdView(1).Text
        spdView(1).Col = 5:    sValue(5) = spdView(1).Text
        spdView(1).Col = 6:    sValue(6) = spdView(1).Text
        spdView(1).Col = 7:    sValue(7) = spdView(1).Text
        spdView(1).Col = 8:    sValue(8) = spdView(1).Text
        
        If Trim(sValue(0)) = "" Then
            Exit Sub
        End If
        
        Call ExecPro("SP_M_SMSALL_005_03", sValue(), Err_Num, Err_Dec)
    Next i

    If Err_Num = 0 Then
        MsgBox "해당되는 데이터가 정상적으로 저장이 되었습니다.", vbInformation
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
 
