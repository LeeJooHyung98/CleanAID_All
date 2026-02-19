VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form P_08003 
   Caption         =   "자료 수신 (MODEM)"
   ClientHeight    =   8250
   ClientLeft      =   1320
   ClientTop       =   3240
   ClientWidth     =   15420
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_08003.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8250
   ScaleWidth      =   15420
   WindowState     =   2  '최대화
   Begin Threed.SSPanel panInput 
      Align           =   1  '위 맞춤
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15420
      _ExtentX        =   27199
      _ExtentY        =   767
      _Version        =   262144
      RoundedCorners  =   0   'False
      Begin VB.TextBox txtInput 
         Height          =   315
         Index           =   0
         Left            =   6630
         TabIndex        =   2
         Top             =   60
         Width           =   3555
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   0
         Left            =   5010
         TabIndex        =   3
         Top             =   60
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "작 업 경 로"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin MSComCtl2.DTPicker dtInput 
         Height          =   315
         Left            =   11910
         TabIndex        =   10
         Top             =   60
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   21430272
         CurrentDate     =   36686
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   3
         Left            =   10290
         TabIndex        =   11
         Top             =   60
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "수선적용일자"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   6
         Left            =   1680
         TabIndex        =   13
         Top             =   60
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   556
         _Version        =   262144
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         Begin Threed.SSOption optSelect 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   30
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   262144
            Caption         =   "모뎀"
            Value           =   -1
         End
         Begin Threed.SSOption optSelect 
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   15
            Top             =   30
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   262144
            Caption         =   "디스켓"
         End
         Begin Threed.SSOption optSelect 
            Height          =   255
            Index           =   2
            Left            =   2010
            TabIndex        =   16
            Top             =   30
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   262144
            Caption         =   "인터넷"
         End
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   5
         Left            =   60
         TabIndex        =   17
         Top             =   60
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "작 업 경 로"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
   End
   Begin Threed.SSPanel panMain 
      Align           =   1  '위 맞춤
      Height          =   9135
      Left            =   0
      TabIndex        =   1
      Top             =   435
      Width           =   15420
      _ExtentX        =   27199
      _ExtentY        =   16113
      _Version        =   262144
      RoundedCorners  =   0   'False
      Begin Threed.SSCommand cmdSubBtn 
         Height          =   555
         Index           =   0
         Left            =   780
         TabIndex        =   4
         Top             =   180
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   979
         _Version        =   262144
         Caption         =   "출고자료 생성"
      End
      Begin Threed.SSCommand cmdSubBtn 
         Height          =   555
         Index           =   1
         Left            =   2760
         TabIndex        =   5
         Top             =   180
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   979
         _Version        =   262144
         Enabled         =   0   'False
         Caption         =   "대리점품목 생성"
      End
      Begin Threed.SSCommand cmdSubBtn 
         Height          =   555
         Index           =   2
         Left            =   4740
         TabIndex        =   6
         Top             =   180
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   979
         _Version        =   262144
         Enabled         =   0   'False
         Caption         =   "할인자료 생성"
      End
      Begin Threed.SSCommand cmdSubBtn 
         Height          =   555
         Index           =   3
         Left            =   6720
         TabIndex        =   7
         Top             =   180
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   979
         _Version        =   262144
         Enabled         =   0   'False
         Caption         =   "목요세일자료 생성"
      End
      Begin Threed.SSCommand cmdSubBtn 
         Height          =   555
         Index           =   4
         Left            =   8700
         TabIndex        =   8
         Top             =   180
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   979
         _Version        =   262144
         Enabled         =   0   'False
         Caption         =   "수선자료 생성"
      End
      Begin Threed.SSCommand cmdSubBtn 
         Height          =   555
         Index           =   5
         Left            =   10680
         TabIndex        =   9
         Top             =   180
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   979
         _Version        =   262144
         Caption         =   "생성자료 지우기"
      End
      Begin Threed.SSCommand cmdSubBtn 
         Height          =   555
         Index           =   6
         Left            =   12660
         TabIndex        =   12
         Top             =   180
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   979
         _Version        =   262144
         Enabled         =   0   'False
         Caption         =   "메일자료 생성"
      End
      Begin Threed.SSCommand cmdSubBtn 
         Height          =   555
         Index           =   7
         Left            =   780
         TabIndex        =   18
         Top             =   900
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   979
         _Version        =   262144
         Enabled         =   0   'False
         Caption         =   "보관 가격 생성"
      End
   End
End
Attribute VB_Name = "P_08003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Err_Num As Long
Dim Err_Dec As String

Dim RS01 As ADODB.Recordset
Dim sValue() As String
Dim ConnectMode As ConnectMode_Type

Private Sub cmdSubBtn_Click(Index As Integer)
    Select Case Index
        Case 0          ' 출고자료생성
        
            If Not Dir(txtInput(0).Text & "\*.*") = "" Then

                    MsgBox "전송 안된 출고 자료가 있으므로 생성자료지우기 작업후 작업 하세요!!!", vbInformation
                    Exit Sub
            End If
        
            P_08003_01.ConnectMode = ConnectMode
            P_08003_01.Show 1
        Case 1          ' 대리점품목생성
            P_08003_02.Show 1
        Case 2          ' 할인자료생성
            P_08003_03.Show 1
        Case 3          ' 목요세일자료생성
            P_08003_04.Show 1
        Case 4          ' 수선자료생성
            Call DataSave1
        Case 5          ' 생성자료지우기
            Call DataSave2
        Case 6          ' 메일자료생성
            Call DataSave3
        Case 7
            P_08003_05.Show 1
        
    End Select
End Sub

Private Sub Form_Activate()
'    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    Dim optSel  As Integer
    optSel = Val(GetIniStr("SERVER DATA", "ConnectMode", "", m_iniFile))
    optSelect(optSel).Value = True
    
    txtInput(0).Text = GetIniStr("SERVER DATA", "SendPath", "", m_iniFile)
    txtInput(0).ToolTipText = txtInput(0).Text
    PanelsMsg ("")
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call INIWrite("SERVER DATA", "ConnectMode", CStr(ConnectMode), m_iniFile)
        
    P_08003_Flag = False
End Sub


Private Sub DataSave1()
    '하드디스크작성
    Dim FileName As String
    Dim rName As String
    Dim rAmt As String
    Dim sDownPath As String
    
    sDownPath = GetIniStr("SERVER DATA", "SendPath", "", m_iniFile)
    
    panCaption(0).Visible = True
    dtInput.Visible = True
    dtInput.Value = Date
    dtInput.Value = ""
    
    If dtInput.Value = "" Then
       Exit Sub
    End If
    
    ReDim sValue(0)
    
    sValue(0) = "0"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_08001_05", sValue(), Err_Num, Err_Dec)
    
    If RS01.RecordCount = 0 Then
        MsgBox "디스켓에 복사할 자료가 없습니다."
        RS01.Close
        Exit Sub
    End If
        
    FileName = "R" & Format(Now, "yyyymmdd") & ".Dat"
        
    Call PanelsMsg("수선자료를 복사중입니다.")
    
    Open sDownPath & "\" & FileName For Output As #1
    
    While Not RS01.EOF
        rAmt = "       "
        LSet rAmt = RS01!금액
        
        Print #1, rAmt;
        Print #1, RS01!명칭
        
        RS01.MoveNext
        
        DoEvents
    Wend
    
    Close #1
    
    RS01.Close
End Sub

Private Sub DataSave2()
    If MsgBox("생성된 자료를 삭제하시겠습니까?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        If Not Dir(txtInput(0).Text & "\*.*") = "" Then
             Kill txtInput(0).Text & "\*.*"
             
             MsgBox "생성자료가 삭제되었습니다.", vbInformation
        End If
    End If
End Sub

Private Sub DataSave3()
    ' 메일자료 생성
    Dim FileName As String
    Dim rName As String
    Dim rAmt As String
    Dim sDownPath As String
    Dim sAgencyCode As String
    
    sDownPath = GetIniStr("SERVER DATA", "SendPath", "", m_iniFile)
    
    ReDim sValue(1)
    
    sValue(0) = "0"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_08001_09", sValue(), Err_Num, Err_Dec)
    
    If RS01.RecordCount = 0 Then
        MsgBox "디스켓에 복사할 자료가 없습니다."
        RS01.Close
        Exit Sub
    End If
    
    Do While Not RS01.EOF
        If sAgencyCode <> RS01!대리점코드 Then
            FileName = "M" & Format(Now, "yyyymmdd") & "." & RS01!대리점코드
            Open sDownPath & "\" & FileName For Output As #1
            sAgencyCode = RS01!대리점코드
        End If
            
        Call PanelsMsg("메일자료를 복사중입니다.")
        
        Print #1, RS01!송신일자;
        Print #1, RS01!메일번호;
        Print #1, RS01!메일내역
        
        RS01.MoveNext
        
        If Not RS01.EOF Then
            If sAgencyCode <> RS01!대리점코드 Then
                Close #1
            End If
        Else
            Close #1
        End If
    Loop
    Call PanelsMsg("메일완료")
    
End Sub

Private Sub optSelect_Click(Index As Integer, Value As Integer)
'    If Index = 1 Then
'        ConnectMode = Floppy
'        txtInput(0).Text = "A:\"
'
'    ElseIf Index = 2 Then
'        ConnectMode = InterNet
'        txtInput(0).Text = "인터넷 연결"
'
'
'    Else
'        ConnectMode = Modem
'        txtInput(0).Text = GetIniStr("SERVER DATA", "ReceivePath", "", m_iniFile)
'
'    End If
End Sub

Private Sub panCaption_Click(Index As Integer)
    If Index = 0 Then cmdSubBtn(1).Enabled = True
End Sub
