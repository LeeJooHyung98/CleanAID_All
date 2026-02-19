VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_03011 
   Caption         =   "출고품목 CHECK"
   ClientHeight    =   11160
   ClientLeft      =   1485
   ClientTop       =   2145
   ClientWidth     =   16260
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_03011.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11160
   ScaleWidth      =   16260
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   11160
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16260
      _ExtentX        =   28681
      _ExtentY        =   19685
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_03011.frx":058A
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   9810
         Left            =   15
         TabIndex        =   1
         Top             =   1335
         Width           =   16230
         _Version        =   524288
         _ExtentX        =   28628
         _ExtentY        =   17304
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
         MaxCols         =   4
         MaxRows         =   35
         ScrollBars      =   0
         SpreadDesigner  =   "P_03011.frx":061C
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   2
         Top             =   540
         Width           =   16230
         _ExtentX        =   28628
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin Threed.SSCommand cmdSubBtn 
            Height          =   315
            Left            =   4800
            TabIndex        =   3
            Top             =   60
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "출 고 검 품 수 신"
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Left            =   1530
            TabIndex        =   4
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   62980096
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   5
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "검 품 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   6
         Top             =   15
         Width           =   8625
         _ExtentX        =   15214
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
         PictureBackground=   "P_03011.frx":0BEB
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   1
         Left            =   8655
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
         PictureBackground=   "P_03011.frx":0DED
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
            Picture         =   "P_03011.frx":0FEF
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
            Caption         =   "엑셀"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_03011.frx":1589
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
            Picture         =   "P_03011.frx":1B23
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
            Picture         =   "P_03011.frx":20BD
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
            Picture         =   "P_03011.frx":2657
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
            Picture         =   "P_03011.frx":2BF1
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
            Picture         =   "P_03011.frx":318B
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
            Picture         =   "P_03011.frx":3725
         End
      End
   End
End
Attribute VB_Name = "P_03011"
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

Private Sub cmdSubBtn_Click()
    Call DataTrans
End Sub

Private Sub dtInput_Change()
    Call Data_Display
End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(5).Enabled = True
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    If P_03011_Flag = False Then
        dtInput.Value = Date
        
        P_03011_Flag = True
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_03011_Flag = True
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i As Long
    
    ReDim sValue(1)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput.Value, "YYYY-MM-DD")
        
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_03011_00", sValue(), Err_Num, Err_Dec)
        
    Do While Not RS01.EOF
        spdView.Row = i
        spdView.Col = 1: spdView.Text = RS01!대리점명 & ""
        spdView.Col = 2: spdView.Text = RS01!검품수량 & ""
        spdView.Col = 3: spdView.Text = RS01!입고수량 & ""
        spdView.Col = 4: spdView.Text = RS01!다른품목 & ""
        
        RS01.MoveNext
    Loop
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub spdView_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim i As Integer
    
    Call AgencyComboAdd(P_03011_01.cboInput)
    
    spdView.Row = Row
    spdView.Col = 1
    
    For i = 1 To P_03011_01.cboInput.ListCount - 1
        If RTrim(P_03011_01.cboInput.List(i)) = RTrim(spdView.Text) Then
            P_03011_01.cboInput.ListIndex = i
        End If
    Next i
    
    P_03011_01.Show
End Sub

Private Sub DataTrans()
    Dim iCnt As Integer
    Dim TmpStr As String
    Dim sInOut As String
    Dim sTagNo As String
    Dim sItem As String
    Dim sDate As String
    Dim sCode As String
    Dim mDate As String
    Dim bDate As String
    
    Dim DupCnt As Integer
    Dim ErrorCnt As Integer
    
    Dim sFilePath As String
    
    P_TRANS.saveYN = False
    P_TRANS.Show 1
    P_TRANS.Hide

    If P_TRANS.saveYN = False Then
        Exit Sub
    End If

    '중복검사 화면
'    spdView(0).MaxRows = 0
'    DupCnt = 0
    
'    spdView(1).MaxRows = 0
    
'    spdView(2).MaxRows = 0
'    ErrorCnt = 0
    
    iCnt = 0
    
    sFilePath = GetIniStr("TERMINAL DATA", "TerminalFilePath", "", m_iniFile)
    
    If Dir(sFilePath & "\Ibchul.dat", vbDirectory) = "" Then
        Exit Sub
    End If
    '핸디로부터 읽은 데이타를 db로
    Open sFilePath & "\Ibchul.dat" For Input As #1
    
    Do While Not EOF(1)
        iCnt = iCnt + 1
        Input #1, TmpStr
        
        sDate = CStr(Year(Date)) & Trim(Mid(TmpStr, 6, 4))  ' dat파일에서 6 자리일자를 8자리로 변환
        mDate = Format(sDate, "####/##/##")
        sCode = Trim(Mid(TmpStr, 10, 3))                    ' dat파일에서 read
        sTagNo = Trim(Mid(TmpStr, 13, 4))                   ' ''
        sItem = Trim(Mid(TmpStr, 28, 3))                    ' 소품구분
        
        '비정상 자료 check
        If Len(Trim(sDate)) <> 8 Or Not IsDate(mDate) Or _
           Len(Trim(sCode)) <> 3 Or Not IsNumeric(sCode) Or _
           Len(Trim(sTagNo)) <> 4 Or Not IsNumeric(sTagNo) Or _
           Len(Trim(sItem)) <> 3 Or sItem = "000" Then
           
'              spdView(0).MaxRows = spdView(0).MaxRows + 1
'              spdView(0).Row = spdView(0).MaxRows
'              spdView(0).Col = 1
'              spdView(0).Text = Mid(TmpStr, 1, 20)
              
              GoTo handy_err
        End If
        
        ReDim sValue(3)
        
        sValue(0) = sDate
        sValue(1) = sCode
        sValue(2) = sTagNo
        sValue(3) = sItem
        
        Call ExecPro("SP_03011_02", sValue(), Err_Num, Err_Dec)
         
        DoEvents
handy_err:
    Loop
    Close #1
    
    If ErrorCnt > 0 Then
       MsgBox CStr(ErrorCnt) & "건이 입고내역이 없습니다."
    End If
End Sub
