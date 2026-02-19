VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_09004 
   Caption         =   "[전사업장] 송신 메일 등록"
   ClientHeight    =   12450
   ClientLeft      =   915
   ClientTop       =   2235
   ClientWidth     =   23535
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_09004.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12450
   ScaleWidth      =   23535
   WindowState     =   2  '최대화
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8010
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   12450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   23535
      _ExtentX        =   41513
      _ExtentY        =   21960
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_09004.frx":058A
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   11085
         Left            =   15
         TabIndex        =   1
         Top             =   1350
         Width           =   6960
         _Version        =   524288
         _ExtentX        =   12277
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
         MaxCols         =   5
         ScrollBars      =   2
         SpreadDesigner  =   "P_09004.frx":063C
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin RichTextLib.RichTextBox rtbInput 
         Height          =   11085
         Left            =   6990
         TabIndex        =   2
         Top             =   1350
         Width           =   16530
         _ExtentX        =   29157
         _ExtentY        =   19553
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"P_09004.frx":0C82
      End
      Begin Threed.SSPanel panInput 
         Height          =   795
         Left            =   15
         TabIndex        =   3
         Top             =   540
         Width           =   23505
         _ExtentX        =   41460
         _ExtentY        =   1402
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.CommandButton Command1 
            Caption         =   "그림 삽입"
            Height          =   315
            Left            =   7980
            TabIndex        =   26
            Top             =   60
            Width           =   1305
         End
         Begin VB.TextBox txtTitle 
            Height          =   315
            Left            =   9420
            TabIndex        =   24
            Top             =   420
            Width           =   5475
         End
         Begin VB.CommandButton cmdAllCheck 
            Caption         =   "전체 선택"
            Height          =   315
            Left            =   9480
            TabIndex        =   5
            Top             =   60
            Width           =   1305
         End
         Begin XtremeSuiteControls.CheckBox chkSelect 
            Height          =   195
            Index           =   1
            Left            =   12420
            TabIndex        =   4
            Tag             =   "2"
            Top             =   120
            Width           =   1185
            _Version        =   851970
            _ExtentX        =   2090
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "유통매장"
            UseVisualStyle  =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            TabIndex        =   6
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   57737216
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   7
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "송 신 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   1530
            TabIndex        =   8
            Top             =   405
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   57737216
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   9
            Top             =   405
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "조 회 기 간"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   2
            Left            =   4815
            TabIndex        =   10
            Top             =   405
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   57737216
            CurrentDate     =   36686
         End
         Begin XtremeSuiteControls.CheckBox chkSelect 
            Height          =   195
            Index           =   0
            Left            =   11010
            TabIndex        =   11
            Tag             =   "1"
            Top             =   120
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
            Left            =   13830
            TabIndex        =   12
            Tag             =   "3"
            Top             =   120
            Width           =   1185
            _Version        =   851970
            _ExtentX        =   2090
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "이마트"
            UseVisualStyle  =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   7950
            TabIndex        =   25
            Top             =   420
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "메일 제목"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            Caption         =   "매장 기본 위치 ↓"
            Height          =   465
            Left            =   16920
            TabIndex        =   27
            Top             =   540
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "~"
            Height          =   225
            Left            =   4620
            TabIndex        =   13
            Top             =   465
            Width           =   225
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   14
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
         Caption         =   " 송신 메일 등록 (P_09004)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_09004.frx":0D27
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
         TabIndex        =   15
         Top             =   15
         Width           =   14625
         _ExtentX        =   25797
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
         PictureBackground=   "P_09004.frx":0F29
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   16
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
            Picture         =   "P_09004.frx":112B
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   17
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
            Picture         =   "P_09004.frx":16C5
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   18
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
            Picture         =   "P_09004.frx":1C5F
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   19
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
            Picture         =   "P_09004.frx":21F9
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   20
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
            Picture         =   "P_09004.frx":2793
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   21
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
            Picture         =   "P_09004.frx":2D2D
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   22
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
            Picture         =   "P_09004.frx":32C7
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   23
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
            Picture         =   "P_09004.frx":3861
         End
      End
   End
End
Attribute VB_Name = "P_09004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String
Dim sCheck  As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub chkSelect_Click(Index As Integer)
    Dim nRow    As Long
    
    With spdView
        For nRow = 1 To .MaxRows
            .Row = nRow
            .Col = 5
            
            ' 현재의 가맹점 종류와 선택 가맹점의 종류가 같을 경우
            If .Text = chkSelect(Index).Caption Then
                sCheck = "N"
                .Row = nRow
                .Col = 3
                .Action = ActionActiveCell
                .Value = chkSelect(Index).Value
            End If
        
        Next nRow
    End With

End Sub

Private Sub cmdAllCheck_Click()
    Dim nRow    As Long
    
    With spdView
        For nRow = 1 To .MaxRows
            .Row = nRow
            .Col = 4
            
            If .Text = "" Then
                sCheck = "N"
                .Row = nRow
                .Col = 3
                .Action = ActionActiveCell
                .Value = IIf(cmdAllCheck.Caption = "전체 선택", "1", "0")
            Else
                Exit For
            End If
        
        Next nRow
    End With
        
    cmdAllCheck.Caption = IIf(cmdAllCheck.Caption = "전체 선택", "전체 취소", "전체 선택")
        
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 0: Call Data_Display           ' 조회
        Case 1: Call DataAdd                ' 신규
        Case 2: Call DataSave               ' 저장
        Case 3:            ' 삭제
        Case 4:            ' 취소
        Case 5:            ' 인쇄
        Case 6:            ' 화면
        Case 7: Unload Me           ' 종료
        
        Case Else
            '
    End Select

End Sub

Sub InsertPictureInRichTextBox(RTB As RichTextBox, Picture As StdPicture)
    Clipboard.Clear
    Clipboard.SetData Picture
    SendMessage RTB.hwnd, WM_PASTE, 0, 0
End Sub

Private Sub Command1_Click()
    On Error GoTo ErrHandler
    CommonDialog1.Filter = "비트맵 파일(*.bmp,*.jpg)|*.bmp;*.jpg|모든 파일(*.*)|*.*|"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.ShowOpen
    InsertPictureInRichTextBox rtbInput, LoadPicture(CommonDialog1.FileName)
    Debug.Print CommonDialog1.FileTitle
    Debug.Print CommonDialog1.FileName
    Exit Sub
ErrHandler:
End Sub
Private Sub Form_Activate()
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"

    Call SubBottonEnable(cmdBtn, "01100001")
    
    If P_09004_Flag = False Then
        Screen.MousePointer = vbHourglass
        dtInput(0).Value = Date
        dtInput(1).Value = Date
        dtInput(2).Value = DateAdd("d", 3, Date)
    
        ReDim sValue(1)
        
        sValue(0) = "0"
        sValue(1) = IIf(Store.Code = "1000", "%", Store.Code & "%")
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_M_09004_00", sValue(), Err_Num, Err_Dec)
        
        'spdView.MaxCols = RS01.Fields.Count
        spdView.MaxRows = RS01.RecordCount
        
        Call spdDisplay(RS01)
        'Call GetColWidth(REG_App, Me.Name, spdView)
        
        P_09004_Flag = True
        sCheck = "N"
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub spdDisplay(Rs As ADODB.Recordset)

    Call fpSpread_Display(spdView, Rs)
    
End Sub

Private Sub spdView_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    If sCheck = "Y" Then Exit Sub
    
    If Row = spdView.ActiveRow Then
                    
        Dim nRow    As Long
        ReDim sValue(2)
        
        If Col = 3 Then
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
                        spdView.Col = 4
                        If spdView.Text = sValue(2) Then
                            sCheck = "Y"
                            spdView.Col = 2
                            spdView.Value = ""
                            spdView.Col = 3
                            spdView.Value = "0"
                        
                        End If
                    Next nRow
                sCheck = "N"
                End If
        
            Else

                
                spdView.Row = Row
                
                sValue(0) = "0"
                sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
                
                Set RS01 = New ADODB.Recordset
                Set RS01 = ExecPro("SP_M_09004_01", sValue(), Err_Num, Err_Dec)
                
                spdView.Col = 2: spdView.Text = "1"
                
                If Not IsNull(RS01!문서번호) Then
                    spdView.Text = RS01!문서번호
                    
                    ' 선택 내용이 지사일 경우 해당 체인점을 모두 선택 시킨다.
                    spdView.Col = 1: sValue(2) = Mid(spdView.Text, 2, 6)
                    
                    If Mid(sValue(2), 5, 1) = "]" Then
                        sValue(2) = Left(sValue(2), 4)
                        
                        For nRow = 1 To spdView.MaxRows
                            spdView.Row = nRow
                            spdView.Col = 4
                            If spdView.Text = sValue(2) Then
                                sCheck = "Y"
                                spdView.Col = 3: spdView.Value = "1"
                                spdView.Col = 2: spdView.Text = RS01!문서번호 & ""
                            End If
                        Next nRow
                        
                        sCheck = "N"
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cmdBtn(0).Enabled = False
    cmdBtn(1).Enabled = False
    cmdBtn(2).Enabled = False
    cmdBtn(3).Enabled = False
    cmdBtn(4).Enabled = False
    cmdBtn(5).Enabled = False
    cmdBtn(6).Enabled = False
    
    P_09004_Flag = False
End Sub

Public Sub DataSave()
    Dim i As Integer
    Dim bFileSave As Boolean
    
    bFileSave = False
    
    If Trim(txtTitle.Text) = "" Then
        MsgBox "메일 제목을 입력 하여 주십시요."
        Exit Sub
    End If
    
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        spdView.Col = 4
        
        If Trim(spdView.Text) <> "" Then
        
            spdView.Row = i
            spdView.Col = 3
            
            If spdView.Value = True Then
                ReDim sValue(8)
                
                sValue(0) = "2"
                sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
                sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
                sValue(3) = Format(dtInput(2).Value, "YYYY-MM-DD")
                
                spdView.Col = 1
                sValue(4) = Mid(spdView.Text, 2, 6)
                spdView.Col = 2
                sValue(5) = spdView.Value
                sValue(6) = txtTitle
                sValue(7) = sValue(1) & Trim(sValue(5))
                sValue(8) = "1"
                
                If HeadOffice = MASTER_OFFICE_CODE Then
                    Call ExecPro("SP_M_09004_12", sValue(), Err_Num, Err_Dec)
                Else
                    
                    If DBOpen_Master(MASTER_OFFICE_CODE) = False Then Exit Sub
                    
                    Set RS01 = New ADODB.Recordset
                    Set RS01 = ExecProMaster("SP_M_09004_12", sValue(), Err_Num, Err_Dec)
                End If
                            
                
                ' 리치박스의 내용을 파일로 한번만 저장한다.
                ' 동일한 내용이기 때문에 작성일자, 문서번호를 기준으로 1번만 저장한다.
                If Not bFileSave Then
                    ' 작성일자, 문서번호
                    Call DataSqlSaveFile(sValue(1), sValue(5))
                    bFileSave = True
                End If
            End If
        End If
    Next i
    
    MsgBox "저장 완료          ", vbInformation, "확인"
    
End Sub

Public Sub DataAdd()
    rtbInput.Text = ""
    
    sValue(0) = "0"
    sValue(1) = IIf(Store.Code = "1000", "%", Store.Code & "%")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_09004_00", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    Call spdDisplay(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView)
    
    sCheck = "N"
End Sub

Private Sub Data_Display()
'
End Sub


Private Sub DataSqlSaveFile(sDate As String, sDocNum As String)
    Dim ADOStream   As ADODB.Stream
    Dim sSql        As String
    Dim Any_RS      As ADODB.Recordset
    Dim sFileFullName As String
    
    On Error GoTo ERR_RTN
    
    
    If Dir(App.Path & "\Temp", vbDirectory) = "" Then MkDir App.Path & "\Temp"
    sFileFullName = App.Path & "\Temp\" & sDate & sDocNum
    rtbInput.SaveFile sFileFullName, rtfRTF
    
    
    Set Any_RS = New ADODB.Recordset
    Set ADOStream = New ADODB.Stream
     
    sSql = "SELECT  문서파일 FROM TB_공지사항파일  "
    sSql = sSql & " WHERE 작성일자 = '" & sDate & "'"
    sSql = sSql & "   AND 문서번호 = " & sDocNum & ""
    
    If HeadOffice = MASTER_OFFICE_CODE Then
        Any_RS.Open sSql, ADOCon, adOpenDynamic, adLockOptimistic
    Else
        
        Any_RS.Open sSql, MSTCon, adOpenDynamic, adLockOptimistic
    End If
    
    
    With ADOStream
        .Type = adTypeBinary
        .Open
        
        If Trim(sFileFullName) <> "" Then
            .LoadFromFile sFileFullName
            Any_RS.Fields(0) = .Read
        Else
            Any_RS!Fields(0) = Null
        End If
        
    End With
    
    Any_RS.Update
    Any_RS.Close
    
    Set Any_RS = Nothing
    Set ADOStream = Nothing
    
    Exit Sub

ERR_RTN:
    
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)

    Set Any_RS = Nothing
    Set ADOStream = Nothing


End Sub


'        Case ID_CHAR_BOLD:        txtClient.SelBold = Not txtClient.SelBold
'        Case ID_CHAR_ITALIC:      txtClient.SelItalic = Not txtClient.SelItalic
'        Case ID_CHAR_UNDERLINE:   txtClient.SelUnderline = Not txtClient.SelUnderline
'
'        Case ID_PARA_LEFT:   txtClient.SelAlignment = rtfLeft
'        Case ID_PARA_CENTER:      txtClient.SelAlignment = rtfCenter
'        Case ID_PARA_RIGHT:  txtClient.SelAlignment = rtfRight
'
'        Case ID_EDIT_CUT:
'            Clipboard.SetText txtClient.SelRTF
'            txtClient.SelText = vbNullString
'        Case ID_EDIT_COPY:
'            Clipboard.SetText txtClient.SelRTF
'        Case ID_EDIT_PASTE:     txtClient.SelRTF = Clipboard.GetText
'        Case ID_EDIT_CUT:
'            Clipboard.SetText txtClient.SelRTF
'            txtClient.SelText = vbNullString
'        Case ID_EDIT_COPY:      Clipboard.SetText txtClient.SelRTF
'        Case ID_EDIT_PASTE:     txtClient.SelRTF = Clipboard.GetText
'
'        Case ID_FONT_SIZE
'            txtClient.SelFontSize = Control.Text
'        Case ID_FONT_FACE
'            txtClient.SelFontName = Control.Text
'
'        Case ID_TEXT_HIGHLIGHTCOLOR:
'        txtClient.SelColor = Control.Color

