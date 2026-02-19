VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_09006 
   Caption         =   "[전사업장]수신 메일 관리"
   ClientHeight    =   11790
   ClientLeft      =   0
   ClientTop       =   1740
   ClientWidth     =   17520
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_09006.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11790
   ScaleWidth      =   17520
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   11790
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17520
      _ExtentX        =   30903
      _ExtentY        =   20796
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_09006.frx":058A
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10395
         Left            =   15
         TabIndex        =   1
         Top             =   1380
         Width           =   6795
         _Version        =   524288
         _ExtentX        =   11986
         _ExtentY        =   18336
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
         SpreadDesigner  =   "P_09006.frx":063C
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin RichTextLib.RichTextBox rtbInput 
         Height          =   10395
         Left            =   6825
         TabIndex        =   2
         Top             =   1380
         Width           =   10680
         _ExtentX        =   18838
         _ExtentY        =   18336
         _Version        =   393217
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"P_09006.frx":0AEA
      End
      Begin Threed.SSPanel panInput 
         Height          =   825
         Left            =   15
         TabIndex        =   3
         Top             =   540
         Width           =   17490
         _ExtentX        =   30850
         _ExtentY        =   1455
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   0
            Left            =   6300
            Style           =   2  '드롭다운 목록
            TabIndex        =   5
            Top             =   420
            Width           =   3075
         End
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   1
            Left            =   1530
            Style           =   2  '드롭다운 목록
            TabIndex        =   4
            Top             =   420
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   4830
            TabIndex        =   6
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   57671680
            CurrentDate     =   36686
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            TabIndex        =   7
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   57671680
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   8
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "송 신 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   4830
            TabIndex        =   9
            Top             =   420
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "가 맹 점"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   35
            Left            =   60
            TabIndex        =   10
            Top             =   420
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "사 업 장"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   11
         Top             =   15
         Width           =   9885
         _ExtentX        =   17436
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
         Caption         =   " 수신 메일 관리 (P_09006)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_09006.frx":0B8F
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   0
         Left            =   9915
         TabIndex        =   12
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
         PictureBackground=   "P_09006.frx":0D91
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   13
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
            Picture         =   "P_09006.frx":0F93
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   14
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
            Picture         =   "P_09006.frx":152D
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   15
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
            Picture         =   "P_09006.frx":1AC7
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   16
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
            Picture         =   "P_09006.frx":2061
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   17
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
            Picture         =   "P_09006.frx":25FB
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   18
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
            Picture         =   "P_09006.frx":2B95
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   19
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
            Picture         =   "P_09006.frx":312F
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   20
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
            Picture         =   "P_09006.frx":36C9
         End
      End
   End
End
Attribute VB_Name = "P_09006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01            As ADODB.Recordset
Dim sValue()        As String
Dim P_09006_Flag    As Boolean

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cboInput_Click(Index As Integer)
    Dim sCode As String

    If Index = 1 Then
        sCode = Trim(Mid(Trim(cboInput(1)) & Space(10), 2, 4))

        Call Get_가맹점리스트(cboInput(0), sCode)
    End If

End Sub

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 0: Call Data_Display           ' 조회
        Case 1:            ' 신규
        Case 2:            ' 저장
        Case 3:            ' 삭제
        Case 4:            ' 취소
        Case 5:            ' 인쇄
        Case 6:            ' 화면
        Case 7: Unload Me           ' 종료
        
        Case Else
            '
    End Select

End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(2).Enabled = True
    cmdBtn(5).Enabled = True
    cmdBtn(6).Enabled = True
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"

    If P_09006_Flag = False Then
        Screen.MousePointer = vbHourglass
        dtInput(0).Value = Date
        dtInput(1).Value = Date
    
        If Store.Code = MASTER_OFFICE_CODE Then
            Call Master_tblComboAdd(cboInput(1))
        Else
            cboInput(1).AddItem "[" & Store.Code & "] " & Store.Name
            cboInput(1).ListIndex = 0
            cboInput(1).Enabled = False

        End If
        
        Call Get_가맹점리스트(cboInput(0), Trim(Mid(Trim(cboInput(1)) & Space(10), 2, 4)))

        ReDim sValue(4)
        
        sValue(0) = "1"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_M_09006_01", sValue(), Err_Num, Err_Dec)
        
        spdView.MaxCols = RS01.Fields.Count
        spdView.MaxRows = RS01.RecordCount
        
        Call spdDisplay(RS01)
        Call GetColWidth(REG_App, Me.Name, spdView)
        
        P_09006_Flag = True
        Screen.MousePointer = vbDefault
    End If
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
        .OperationMode = OperationModeSingle
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    
    
        .ColsFrozen = 1 '틀고정
        .Row = -1
        
        .Col = 1
        .ColWidth(1) = 10
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 2
        .ColWidth(2) = 5
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 3
        .ColWidth(3) = 18
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
        
        .Col = 4
        .ColWidth(4) = 5
        .CellType = CellTypeCheckBox
        .Value = False
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
            
        .Col = 5
        .ColWidth(5) = 5
        .CellType = CellTypeCheckBox
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
            
        .Col = 6
        .ColWidth(6) = 300
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
    End With
        
    dtInput(0).Value = Date
    dtInput(1).Value = Date


    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub spdDisplay(Rs As ADODB.Recordset)

    Call fpSpread_Display(spdView, Rs)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    cmdBtn(0).Enabled = False
    cmdBtn(1).Enabled = False
    cmdBtn(2).Enabled = False
    cmdBtn(3).Enabled = False
    cmdBtn(4).Enabled = False
    cmdBtn(5).Enabled = False
    cmdBtn(6).Enabled = False
    
    P_09006_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    ReDim sValue(4)
     
    sValue(0) = "0"
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    sValue(3) = Mid(cboInput(1).Text, 2, 4)
    sValue(4) = Mid(cboInput(0).Text, 2, 6)
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_09006_01", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    Call spdDisplay(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView)
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

'+------------------------------------------------------
'+
'+ 2003/04/11
'+
'+루틴설명
'+  1. 목록에 선택된 내용을 DB에 적용 시킨다.
'+  2. Mail의 SendChk = "2"로 변경하여 모뎀자료 생성에서
'+      생성 되도록 설정한다.
'+  3. 체크된것만을 작업하기 때문에 이전 작업 내용과 무관하다.
'+------------------------------------------------------
Public Sub DataSave()
    ReDim sValue(4)
    
    Dim i As Long
    Dim strSql As String
    Dim strCnn As String
    Dim rstMail As ADODB.Recordset
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        spdView.Col = 4
        
        spdView.Col = 1: sValue(0) = Format(spdView.Value, "YYYY-MM-DD")
        spdView.Col = 2: sValue(1) = Val(spdView.Text)
        spdView.Col = 3: sValue(2) = Mid(spdView.Text, 2, 6)
        sValue(3) = "2"
        
        strSql = "SELECT * FROM Mail_ALL   WHERE MailDate = '" _
            & sValue(0) & "' AND AgencyCode = '" & sValue(2) & "' AND MailType = '" _
            & sValue(3) & "' AND MailNo = '" & sValue(1) & "'"
        Set rstMail = New ADODB.Recordset
        
        rstMail.CursorType = adOpenKeyset
        rstMail.LockType = adLockOptimistic
        rstMail.Open strSql, m_DBConnect, , , adCmdText
        
        spdView.Col = 4
        rstMail!SendChk = IIf(spdView.Value = True, 0, 1)
        rstMail.Update
        rstMail.Close
    Next i
End Sub

Public Sub DataPrint()

End Sub

Public Sub DataScreen()

End Sub

Private Sub PrintDesc()

End Sub

Private Sub spdView_DblClick(ByVal Col As Long, ByVal Row As Long)
    ReDim sValue(3)
    
   sValue(0) = "0"
    
    spdView.Row = spdView.ActiveRow
    spdView.Col = 1: sValue(1) = Format(spdView.Value, "YYYY-MM-DD")
    spdView.Col = 2: sValue(2) = spdView.Value
    spdView.Col = 3: sValue(3) = Mid(spdView.Text, 2, 6)
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_09006_00", sValue(), Err_Num, Err_Dec)
    
    rtbInput.Text = RS01!메일내역
End Sub



