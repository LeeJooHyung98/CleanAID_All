VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_03013 
   Caption         =   "출고TAG번호 CHECK"
   ClientHeight    =   10500
   ClientLeft      =   1275
   ClientTop       =   1920
   ClientWidth     =   16140
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_03013.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10500
   ScaleWidth      =   16140
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   10500
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16140
      _ExtentX        =   28469
      _ExtentY        =   18521
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_03013.frx":058A
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   9150
         Left            =   15
         TabIndex        =   1
         Top             =   1335
         Width           =   16110
         _Version        =   524288
         _ExtentX        =   28416
         _ExtentY        =   16140
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
         MaxCols         =   7
         MaxRows         =   35
         ScrollBars      =   0
         SpreadDesigner  =   "P_03013.frx":061C
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   2
         Top             =   540
         Width           =   16110
         _ExtentX        =   28416
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            TabIndex        =   3
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
            TabIndex        =   4
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "출 고 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   4770
            TabIndex        =   5
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   62980096
            CurrentDate     =   36686
         End
         Begin VB.Label Label1 
            Alignment       =   2  '가운데 맞춤
            Caption         =   "~"
            Height          =   255
            Left            =   4530
            TabIndex        =   6
            Top             =   120
            Width           =   255
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   7
         Top             =   15
         Width           =   8505
         _ExtentX        =   15002
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
         PictureBackground=   "P_03013.frx":0C1D
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   0
         Left            =   8535
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
         PictureBackground=   "P_03013.frx":0E1F
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
            Picture         =   "P_03013.frx":1021
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   10
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
            Picture         =   "P_03013.frx":15BB
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   11
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
            Picture         =   "P_03013.frx":1B55
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   12
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
            Picture         =   "P_03013.frx":20EF
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   13
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
            Picture         =   "P_03013.frx":2689
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   14
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
            Picture         =   "P_03013.frx":2C23
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   15
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
            Picture         =   "P_03013.frx":31BD
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   16
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
            Picture         =   "P_03013.frx":3757
         End
      End
   End
End
Attribute VB_Name = "P_03013"
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

Private Sub dtInput_Change(Index As Integer)
'    Call Data_Display
End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(5).Enabled = True
    cmdBtn(6).Enabled = True
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    If P_03013_Flag = False Then
        dtInput(0).Value = Date
        dtInput(1).Value = Date
        
        P_03013_Flag = True
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_03013_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i As Integer
    Dim j As Integer
    Dim lAmt As Long
    
    ReDim sValue(2)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
        
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_03013_00", sValue(), Err_Num, Err_Dec)
    
    i = 0
    j = 1
    
    Do Until RS01.EOF
        spdView.Row = j
        
        spdView.Col = 1: spdView.Text = RS01!대리점명 & ""
        spdView.Col = 2: spdView.Text = RS01!출고수량 & ""
        spdView.Col = 3: spdView.Text = RS01!시작택 & ""
        spdView.Col = 4: spdView.Text = RS01!종료택 & ""
        spdView.Col = 5: spdView.Text = RS01!중복수량 & ""
        spdView.Col = 6: spdView.Text = RS01!누락수량 & ""
        
        RS01.MoveNext
    Loop
    
    Call TagCheck
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Public Sub DataPrint()

End Sub

Private Sub spdView_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim i As Integer
    
    Call AgencyComboAdd(P_03013_01.cboInput)
    
    spdView.Row = Row
    spdView.Col = 1
    
    For i = 1 To P_03013_01.cboInput.ListCount - 1
        If P_03013_01.cboInput.List(i) = spdView.Text Then
            P_03013_01.cboInput.ListIndex = i
        End If
    Next i
    
    P_03013_01.Show
End Sub

Private Sub spdView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        'PopupMenu P_00000.PopUp
    End If
End Sub

Private Sub TagCheck()
    Dim i As Integer
    Dim z As Integer
    Dim iCnt As Integer
    Dim rCnt As Integer
    
    Dim cTag As String
    Dim sTag As String
    Dim eTag As String
    
    Dim mSu As Integer
    Dim dSu As Integer
    Dim flag As String
    
    ReDim sValue(0)
    
    sValue(0) = UserID
    
    Call ExecPro("SP_03013_02", sValue(), Err_Num, Err_Dec)
    
    For iCnt = 1 To spdView.MaxRows
        spdView.Row = iCnt
        spdView.Col = 1
        If spdView.Text <> "" Then
            spdView.Col = 3: sTag = Mid(spdView.Text, 1, 1) & Mid(spdView.Text, 3, 3)
            spdView.Col = 4: eTag = Mid(spdView.Text, 1, 1) & Mid(spdView.Text, 3, 3)
            
            rCnt = Val(eTag) - Val(sTag) + 1
            
            ReDim sValue(3)
            
            sValue(0) = "0"
            sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
            sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
            
            spdView.Col = 1: sValue(3) = Mid(spdView.Text, 2, 3)
            
            Set RS01 = New ADODB.Recordset
            Set RS01 = ExecPro("SP_03013_01", sValue(), Err_Num, Err_Dec)
            
            mSu = 0
            dSu = 0
            
            If rCnt < 8000 Then
                cTag = RS01!택번호
                
                Do While Not RS01.EOF
                    dSu = dSu + (RS01!수량 - 1)
                    
                    If cTag <> RS01!택번호 Then
                        mSu = mSu + (Val(RS01!택번호) - Val(cTag) - 1)
                    End If
                    
                    cTag = RS01!택번호
                    
                    RS01.MoveNext
                Loop
                
                If dSu = 0 And mSu = 0 Then
                    flag = "1"
                Else
                    flag = "2"
                End If
            Else
                flag = "3"
                rCnt = 0
                cTag = "0000"
                
                Do While Not RS01.EOF
                    If Val(RS01!택번호) < 5000 Then
                        i = 2
                    Else
                        i = 1
                    End If
                    
                    dSu = dSu + (RS01!수량 - 1)
                    
                    ReDim sValue(4)
                    
                    sValue(0) = UserID
                    
                    spdView.Col = 1
                    sValue(1) = Mid(spdView.Text, 2, 3)
                    
                    sValue(2) = i
                    sValue(3) = RS01!택번호
                    sValue(4) = RS01!수량
                    
                    Call ExecPro("SP_03013_03", sValue(), Err_Num, Err_Dec)
                    
                    Do While rCnt < Val(RS01!택번호)
                        If Val(RS01!택번호) - rCnt < 5000 Then
                            sTag = Trim(Str(rCnt))
                            sTag = Right("0000" & sTag, 4)
                            
                            ReDim sValue(4)
                            
                            sValue(0) = UserID
                            
                            spdView.Col = 1
                            sValue(1) = Mid(spdView.Text, 2, 3)
                            
                            sValue(2) = i + 2
                            sValue(3) = sTag
                            sValue(4) = "1"
                            
                            Call ExecPro("SP_03013_03", sValue(), Err_Num, Err_Dec)
                            
                            flag = "4"
                            rCnt = rCnt + 1
                            mSu = mSu + 1
                        Else
                            rCnt = Val(RS01!택번호) + 1
                        End If
                    Loop
                    
                    rCnt = Val(RS01!택번호) + 1
                    cTag = RS01!택번호
                
                    RS01.MoveNext
                Loop
                
                Do While rCnt < 10000
                    sTag = Trim(Str(rCnt))
                    sTag = Right("0000" & sTag, 4)
                    
                    ReDim sValue(4)
                    
                    sValue(0) = UserID
                    
                    spdView.Col = 1
                    sValue(1) = Mid(spdView.Text, 2, 3)
                    
                    sValue(2) = i + 2
                    sValue(3) = sTag
                    sValue(4) = "1"
                    
                    Call ExecPro("SP_03013_03", sValue(), Err_Num, Err_Dec)
                    
                    flag = "4"
                    rCnt = rCnt + 1
                    mSu = mSu + 1
                Loop
            End If
            
            spdView.Col = 5: spdView.Text = dSu
            spdView.Col = 6: spdView.Text = mSu
            spdView.Col = 7: spdView.Text = flag
            
            If flag = "2" Or flag = "4" Then
                spdView.Col = -1: spdView.BackColor = &HD8FCFE
            End If
        End If
        
        DoEvents
    Next iCnt
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
'    P_00000.crPrint.Formulas(0) = "출고일자1 = '" & Format(dtInput(0).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(1) = "출고일자2 = '" & Format(dtInput(0).Value, "YYYY-MM-DD") & "'"
'
'    Call ReportPrint(ReportFile, "2")
End Sub

Private Sub PrintDesc()
    Dim i As Integer
    Dim TempText As String
    Dim TempFP As String
    Dim TempFile As String
    
    TempFP = GetIniStr("REPORT", "TempPath", "", m_iniFile)
    TempFile = TempFP & "\Temp.txt"
    
    Open TempFile For Output As #1
    
    TempText = ""
    
    For i = 1 To spdView.MaxRows - 1
        spdView.Row = i
        
        spdView.Col = 6
        If spdView.Text = 0 Then
            spdView.Col = 1: TempText = TempText & "    " & LeftH(spdView.Text & Space(20), 20)
        Else
            spdView.Col = 1: TempText = TempText & "   *" & LeftH(spdView.Text & Space(20), 20)
        End If
        
        spdView.Col = 2: TempText = TempText & Right(Space(8) & spdView.Text, 8)
        spdView.Col = 5: TempText = TempText & Right(Space(8) & spdView.Text, 8)
        spdView.Col = 6: TempText = TempText & Right(Space(8) & spdView.Text, 8)
        
        If i Mod 2 = 0 Then
            Print #1, TempText
            TempText = ""
        End If
    Next i
    
    Close #1
End Sub
