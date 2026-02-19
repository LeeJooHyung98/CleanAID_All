VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_06010 
   Caption         =   " 사고 담당자 등록"
   ClientHeight    =   11415
   ClientLeft      =   630
   ClientTop       =   3765
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
   Icon            =   "P_06010.frx":0000
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
      PaneTree        =   "P_06010.frx":058A
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
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   8
            Left            =   9690
            TabIndex        =   14
            Top             =   240
            Width           =   3045
            _Version        =   851970
            _ExtentX        =   5371
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 담당자 가맹점 저장"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_06010.frx":063C
         End
         Begin XtremeSuiteControls.PushButton cmdBtnChk 
            Height          =   450
            Left            =   8370
            TabIndex        =   15
            Top             =   240
            Width           =   1260
            _Version        =   851970
            _ExtentX        =   2222
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "전체선택"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_06010.frx":0BD6
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
         Caption         =   " 사고 담당자 등록 (P_06010)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_06010.frx":1170
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
         PictureBackground=   "P_06010.frx":1372
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
            Picture         =   "P_06010.frx":1574
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
            Picture         =   "P_06010.frx":1B0E
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
            Picture         =   "P_06010.frx":20A8
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
            Picture         =   "P_06010.frx":2642
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
            Picture         =   "P_06010.frx":2BDC
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
            Picture         =   "P_06010.frx":3176
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
            Picture         =   "P_06010.frx":3710
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
            Picture         =   "P_06010.frx":3CAA
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10065
         Index           =   0
         Left            =   15
         TabIndex        =   12
         Top             =   1335
         Width           =   7830
         _Version        =   524288
         _ExtentX        =   13811
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
         MaxCols         =   3
         SpreadDesigner  =   "P_06010.frx":4244
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10065
         Index           =   1
         Left            =   7860
         TabIndex        =   13
         Top             =   1335
         Width           =   14775
         _Version        =   524288
         _ExtentX        =   26061
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
         MaxCols         =   5
         ScrollBars      =   2
         SpreadDesigner  =   "P_06010.frx":478B
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_06010"
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
        Case 1: Call DataAdd        ' 신규
        Case 2: Call DataSave       ' 저장
        Case 3: Call DataDelete     ' 삭제
        Case 4: Call DataCancel     ' 취소
        Case 5: 'Call DataPrint      ' 인쇄
        Case 6: Call Export_Excel(P_00000.cdgExcel, spdView(0))      ' 엑셀
        Case 7: Unload Me           ' 종료
        Case 8: Call DataSMSSave
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

Private Sub cmdBtnChk_Click()
    Dim nChk    As Integer
    Dim nRow    As Long
    
    nChk = IIf(cmdBtnChk.Caption = "전체선택", 1, 0)
    cmdBtnChk.Caption = IIf(cmdBtnChk.Caption = "전체선택", "선택취소", "전체선택")
    
    With spdView(1)
        For nRow = 1 To .MaxRows
            .SetText 1, nRow, CVar(nChk)
        Next nRow
    End With
End Sub

 
Private Sub cmdBtnStoreSave_Click()

End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(1).Enabled = True
    cmdBtn(2).Enabled = True
    cmdBtn(3).Enabled = True
    cmdBtn(4).Enabled = False
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
    
    If P_06010_Flag = False Then
        Call DataDisplay
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

    
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_06010_Flag = False
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
    Dim i As Integer
    ReDim sValue(3)
    Dim sTel(2) As String
    
    For i = 1 To spdView(0).DataRowCnt
        spdView(0).Row = i
        
                                sValue(0) = Store.Code
        spdView(0).Col = 1:    sValue(1) = spdView(0).Text
        spdView(0).Col = 2:    sValue(2) = spdView(0).Text
        spdView(0).Col = 3:    sValue(3) = spdView(0).Text
        
        If CheckMobileNumber(sValue(3), sTel) = False Then
            MsgBox "[" & Trim(sValue(2)) & "]의 전화번호를 확인하여 주십시요", vbInformation, "확인"
            Exit Sub
        End If
    Next i
    
    For i = 1 To spdView(0).DataRowCnt
        spdView(0).Row = i
        
                                sValue(0) = Store.Code
        spdView(0).Col = 1:    sValue(1) = spdView(0).Text
        spdView(0).Col = 2:    sValue(2) = spdView(0).Text
        spdView(0).Col = 3:    sValue(3) = spdView(0).Text
        
        If Trim(sValue(0)) = "" Then
            Exit Sub
        End If
        
        Call ExecPro("SP_06010_01", sValue(), Err_Num, Err_Dec)
    Next i

    If Err_Num = 0 Then
        MsgBox "해당되는 데이터가 정상적으로 저장이 되었습니다.", vbInformation
    End If
End Sub

Public Sub DataDelete()
    
    If MsgBox("해당되는 데이터를 삭제하시겠습니까?", vbYesNo + vbInformation + vbDefaultButton2, "데이터 삭제") = vbYes Then

        ReDim sValue(1)

        spdView(0).Row = spdView(0).ActiveRow
                                sValue(0) = Store.Code
        spdView(0).Col = 1:     sValue(1) = spdView(0).Text

        Call ExecPro("SP_06010_02", sValue(), Err_Num, Err_Dec)

        MsgBox "해당되는 데이터가 정상적으로 삭제가 되었습니다.", vbInformation
        
        Call DataDisplay
    End If
End Sub

Public Sub DataCancel()
'    If optSelect(0).Value = True Then
'        Call optSelect_Click(0, True)
'    ElseIf optSelect(1).Value = True Then
'        Call optSelect_Click(1, True)
'    End If
End Sub

Public Sub DataDisplay()

        ReDim sValue(0)
        
        sValue(0) = Store.Code
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_06010_00", sValue(), Err_Num, Err_Dec)
        
        With spdView(0)
            .MaxRows = 0
            .Redraw = False
                        
            Do Until RS01.EOF
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                
                .Col = 1: .Text = RS01("담당자코드") & ""
                .Col = 2: .Text = RS01("담당자명") & ""
                .Col = 3: .Text = RS01("휴대폰번호") & ""
                
                RS01.MoveNext
            Loop
            RS01.Close
            Set RS01 = Nothing
            
            .Redraw = True
        End With

End Sub

Public Sub DataDisplayStoreList(sCode As String)

        ReDim sValue(0)
        
        sValue(0) = sCode
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_06010_03", sValue(), Err_Num, Err_Dec)
        
        With spdView(1)
            .MaxRows = 0
            .Redraw = False
            
            .Tag = sCode
                        
            Do Until RS01.EOF
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                
                .Col = 1: .Text = RS01("Chk") & ""
                .Col = 2: .Text = RS01("가맹점코드") & ""
                .Col = 3: .Text = RS01("가맹점명") & ""
                .Col = 4: .Text = RS01("지사코드") & ""
                .Col = 5: .Text = RS01("지사명") & ""
                
                RS01.MoveNext
            Loop
            RS01.Close
            Set RS01 = Nothing
            
            .Redraw = True
        End With

End Sub

Private Sub spdView_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
    If Index = 0 Then
        Dim vText   As Variant
        
        spdView(0).GetText 1, Row, vText
    
        If CStr(vText) <> "" Then Call DataDisplayStoreList(CStr(vText))
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


' SMS 사용자 저장
Private Sub DataSMSSave()
    Dim nRow    As Long
    Dim vText   As Variant
    
    Rtn = MsgBox("저장하시겠습니까?", vbYesNo + vbInformation, "확인")
    If Rtn = vbNo Then Exit Sub
    
    
    '----------------------------------------------------------------------------
    ReDim sValue(4)
    
    With spdView(1)
        For nRow = 1 To .DataRowCnt
            .Row = nRow
            
                                            sValue(0) = Store.Code                          '1 자사코드
            .GetText 2, nRow, vText:        sValue(1) = CStr(vText)                         '2 가맹점코드
                                            sValue(2) = .Tag                                '3 담당자코드
            .GetText 1, nRow, vText:        sValue(3) = IIf(CStr(vText) = "1", "Y", "N")    '4 전송여부
    
            If Trim(sValue(0)) = "" Then
                MsgBox "지사 코드를 선택하여 주십시요", vbInformation, "확인"
                Exit Sub
            End If
            If Trim(sValue(1)) = "" Then
                MsgBox "매장을 선택하여 주십시요", vbInformation, "확인"
                Exit Sub
            End If
            If Trim(sValue(2)) = "" Then
                MsgBox "담당자 코드를 선택하여 주십시요", vbInformation, "확인"
                Exit Sub
            End If
            
            Call ExecPro("SP_01001_SMS_01", sValue(), Err_Num, Err_Dec)
            DoEvents
    
        Next nRow
    
        If Err_Num = 0 Then
            MsgBox "해당되는 데이터가 정상적으로 저장이 되었습니다.", vbInformation
        End If

    End With

End Sub

