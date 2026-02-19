VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_LOGIN 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "사용자 로그인"
   ClientHeight    =   4680
   ClientLeft      =   5610
   ClientTop       =   4200
   ClientWidth     =   4470
   Icon            =   "P_LOGIN.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2765.099
   ScaleMode       =   0  '사용자
   ScaleWidth      =   4197.088
   StartUpPosition =   2  '화면 가운데
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   4680
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   8255
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   0
      PaneTree        =   "P_LOGIN.frx":058A
      Begin Threed.SSPanel SSPanel 
         Height          =   810
         Left            =   0
         TabIndex        =   11
         Top             =   3870
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   1429
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton btnEnd 
            Height          =   645
            Left            =   2865
            TabIndex        =   4
            Top             =   75
            Width           =   1530
            _Version        =   851970
            _ExtentX        =   2699
            _ExtentY        =   1138
            _StockProps     =   79
            Caption         =   " 종료"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "P_LOGIN.frx":05FC
         End
         Begin XtremeSuiteControls.PushButton btnConnect 
            Height          =   645
            Left            =   60
            TabIndex        =   3
            Top             =   75
            Width           =   1530
            _Version        =   851970
            _ExtentX        =   2699
            _ExtentY        =   1138
            _StockProps     =   79
            Caption         =   " 접속"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "P_LOGIN.frx":168E
         End
      End
      Begin Threed.SSPanel panTitle 
         Height          =   2400
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   4233
         _Version        =   262144
         Font3D          =   5
         ForeColor       =   16711680
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "P_LOGIN.frx":1D88
         BevelOuter      =   0
         PictureAlignment=   7
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panMain 
         Height          =   1440
         Left            =   0
         TabIndex        =   7
         Top             =   2415
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   2540
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtInput 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            IMEMode         =   3  '사용 못함
            Index           =   2
            Left            =   1560
            MaxLength       =   4
            PasswordChar    =   "*"
            TabIndex        =   2
            Top             =   975
            Width           =   2535
         End
         Begin VB.TextBox txtInput 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox txtInput 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   1560
            MaxLength       =   6
            TabIndex        =   0
            Top             =   210
            Width           =   2535
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "사용자 ID :"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   10
            Top             =   285
            Width           =   1260
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "사용자명 :"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   225
            TabIndex        =   9
            Top             =   660
            Width           =   1260
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "비밀번호 :"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   225
            TabIndex        =   8
            Top             =   1050
            Width           =   1260
         End
      End
   End
End
Attribute VB_Name = "P_LOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Dim sPassword As String * 4
Dim iCount As Integer

Private Sub btnConnect_Click()
    If sPassword = txtInput(2).Text Then
        P_00000.Show
        
        Unload Me
    Else
        iCount = iCount + 1
        
        If iCount = 3 Then
            MsgBox "비밀번호를 3회이상 틀였으므로 프로그램을 종료합니다.", vbInformation
            ADOCon.Close
            End
        Else
            MsgBox "비밀번호가 틀렸습니다.", vbInformation
        End If
    End If
End Sub

Private Sub btnEnd_Click()
    ADOCon.Close
    End
End Sub

Private Sub Form_Activate()
    ' 자동 로그인 기능을 지원한다.
    Dim Cmd() As String
    
    If Command() <> "" Then
        
        Cmd = Split(Command(), ",")
        
        If UBound(Cmd) >= 1 Then
            txtInput(0).Text = Cmd(0)
            DoEvents
            
            Call txtInput_LostFocus(0)
            
            txtInput(2).Text = Cmd(1)
            DoEvents
            
            Call btnConnect_Click
        End If
    End If
End Sub

Private Sub Form_Initialize()
    '폼이 실행시 처음 처리되는 루틴
    If (App.PrevInstance) Then
        'MsgBox 처리
        MsgBox "프로그램이 이미 실행중 입니다.", vbInformation, "오류 메세지"
        
'        Kill App.Path & "\" & App.EXEName & ".exe"
'        ProgramUpgrade
        End
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    Call SetProgramVersion ' 프로그램 버전및 마지막 수정일자 설정
    Call GetDefaultValues  ' 기본 사항 읽기
    
    P_CONNECT.Show 1
                
    m_Error.VisibleMSG = True
    m_Error.ResumeMode = True
    
    
''    Dim RecentVersion, MyVersion As String
''    Dim RecentMemo As String
'''
''    scUrl = GetSetting(scRegAppname, scRegSection, "Url", scUrl)
''    scName = GetSetting(scRegAppname, scRegSection, "Name", scName)
''    scFold = GetSetting(scRegAppname, scRegSection, "Fold", scFold)
''    scFoldName = GetSetting(scRegAppname, scRegSection, "FoldName", scFoldName)
''
''
''    'url 변경하여 무조건 레지스트리에 저장함.20090115
''
''    If scUrl <> "http://www.clean-aid.co.kr:8090/business/" Then
''        scUrl = "http://www.clean-aid.co.kr:8090/business/"
''        scName = "백상영업.exe"
''        scFold = App.Path & "\"
''        scFoldName = "백상영업UP.exe"
''
''        SaveSetting scRegAppname, scRegSection, "Url", scUrl
''        SaveSetting scRegAppname, scRegSection, "Name", scName
''        SaveSetting scRegAppname, scRegSection, "Fold", scFold
''        SaveSetting scRegAppname, scRegSection, "FoldName", scFoldName
''    End If
''
''    If Trim(scUrl) = "" Then
''        scUrl = "http://www.clean-aid.co.kr:8090/business/"
''        scName = "백상영업.exe"
''        scFold = App.Path & "\"
''        scFoldName = "백상영업UP.exe"
''
''        SaveSetting scRegAppname, scRegSection, "Url", scUrl
''        SaveSetting scRegAppname, scRegSection, "Name", scName
''        SaveSetting scRegAppname, scRegSection, "Fold", scFold
''        SaveSetting scRegAppname, scRegSection, "FoldName", scFoldName
''
''    End If
''    'Url 변경 추가 끝
''
''    'Url 변경 전 내용
''    '    If Trim(scUrl) = "" Then
''    '            scUrl = "http://www.clean-aid.co.kr:8090/business/"
''    '            scName = "백상영업.exe"
''    '            scFold = App.Path & "\"
''    '            scFoldName = "백상영업UP.exe"
''    '    End If
''
''    'RecentVersion = OpenURL("http://www.philipn.com/program/ver.txt", 1000)
''    'RecentMemo = OpenURL("http://www.philipn.com/program/memo.txt", 1000)
''    If Right(scUrl, 1) = "/" Then
''        RecentVersion = OpenURL(scUrl & "Ver.txt", 1000)
''        RecentMemo = OpenURL(scUrl & "memo.txt", 1000)
''    Else
''        RecentVersion = OpenURL(scUrl & "/Ver.txt", 1000)
''        RecentMemo = OpenURL(scUrl & "/memo.txt", 1000)
''    End If
''
''    'UpGrade
''    MyVersion = GetSetting(scRegAppname, scRegSection, "VerSion", MyVersion)
''    'MyVersion = strProgram_Version
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub panLabel_Click(Index As Integer)
'    txtInput(0).Text = "4939"
'    txtInput(2).Text = "4939"
   
End Sub

Private Sub txtInput_GotFocus(Index As Integer)
    txtInput(Index).SelStart = 0
    txtInput(Index).SelLength = Len(txtInput(Index).Text)
End Sub

 
Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtInput_LostFocus(Index As Integer)
    If Index = 0 Then
        ReDim sValue(1)
    
        'sValue(0) = "0"
        sValue(0) = Store.Code
        sValue(1) = txtInput(0).Text
        
        If sValue(1) = "" Then
            Exit Sub
        End If
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_LOGIN", sValue(), Err_Num, Err_Dec)
        
        If RS01.RecordCount = 0 Then
            RS01.Close
            Set RS01 = Nothing
            
            MsgBox "해당되는 사용자ID는 존재하지 않습니다.", vbInformation, "LOGIN"
            txtInput(0).SelStart = 0: txtInput(0).SelLength = Len(txtInput(0).Text)
            txtInput(0).SetFocus
            Exit Sub
        Else
            txtInput(1).Text = RS01!사용자명 & "" '1
            sPassword = RS01!비밀번호 & ""        '2
            UserID = txtInput(0).Text & ""        '3
            USERNAME = RS01!사용자명 & ""         '4
            
            HeadOffice = RS01!지사코드 & ""       '5
            
            RS01.Close
            Set RS01 = Nothing
        End If
    End If
End Sub
