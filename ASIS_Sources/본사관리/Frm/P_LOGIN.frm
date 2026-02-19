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
   KeyPreview      =   -1  'True
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
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Height          =   225
            Left            =   3420
            TabIndex        =   12
            Top             =   1140
            Width           =   555
         End
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

Dim sPASSWORD As String * 4
Dim iCount As Integer

Private Sub btnConnect_Click()
    If sPASSWORD = txtInput(2).Text Then
        P_00000.Show
        
        Unload Me
    Else
        iCount = iCount + 1
        
        If iCount = 3 Then
            MsgBox "비밀번호를 3회오류로 프로그램을 종료합니다.", vbInformation
            ADOCon.Close
            End
        Else
            MsgBox "비밀번호가 틀렸습니다.", vbInformation
            txtInput(2).SelStart = 0
            txtInput(2).SelLength = 10
            txtInput(2).SetFocus
        
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        Label1_Click
    End If

End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    Call SetProgramVersion ' 프로그램 버전및 마지막 수정일자 설정
    Call GetDefaultValues  ' 기본 사항 읽기
    
    P_CONNECT.Show 1
                
    m_Error.VisibleMSG = True
    m_Error.ResumeMode = True
    
    
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Label1_Click()
    Dim sPWD    As String
    
    txtInput(0).Text = ""
    txtInput(1).Text = ""
    
    
    sPWD = InputBox("설정 내용을 변경 하기 위하여 변경 암호를 입력 하여 주십시요", "암호입력")
    
    Select Case UCase(sPWD)
        Case "ISN", "DUDTJSGH", "SHOP500"
            P_SERVER.Show 1
            Exit Sub
            
        Case Else
            Exit Sub
    End Select
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
    
        DoEvents
            
        '--------------------------------------------------------------------------------------------------------------------------------------
        ' 프로그램 사용 종료일 확인 하여 사용하지 못하도록 설정
        ' 폐점 매장에서 계속 사용하는 문제 처리
        '--------------------------------------------------------------------------------------------------------------------------------------
        Call GetProgramCloseDate
        
        sValue(0) = Store.Code
        sValue(1) = "정상 (" & Format(App.Major, "00") & "." & Format(App.Minor, "00") & "." & Format(App.Revision, "00") & ")"
        
        If Store.Stats = "N" And Store.PGCloseDate < Format(Date, "yyyy-MM-dd") And Trim(txtInput(0).Text) <> "" Then
            Dim sMsg    As String
             
             '-------------------------------------------------------------------
             ' @@ 접속 정보를 저장한다.
             '-------------------------------------------------------------------
            sValue(1) = "제한-로그인 실패"
            
            sMsg = "" & vbNewLine
            sMsg = sMsg & "프로그램 사용이 제한된 매장 입니다.          " & vbNewLine & vbNewLine
            sMsg = sMsg & "지사상태 : " & Store.Code & vbNewLine
            sMsg = sMsg & "사용종료일 : " & Store.PGCloseDate & vbNewLine & vbNewLine
            sMsg = sMsg & "연락처(전산실) : 031)522-2025"
        
            ' 접속 정보 저장
            Set RS01 = New ADODB.Recordset
            Set RS01 = ExecPro("SP_PROGRAM_CONNECT_INS", sValue(), Err_Num, Err_Dec)
                
            MsgBox sMsg, vbInformation, "  알 림  "
            
            End
        
        End If
     
        ' 접속 정보 저장
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_PROGRAM_CONNECT_INS", sValue(), Err_Num, Err_Dec)
    
        
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
            sPASSWORD = RS01!비밀번호 & ""        '2
            UserID = txtInput(0).Text & ""        '3
            USERNAME = Trim(RS01!사용자명 & "")   '4
            
            HeadOffice = RS01!지사코드 & ""       '5
            
            RS01.Close
            Set RS01 = Nothing
        End If
    End If
End Sub



Private Function GetProgramCloseDate() As String
    Dim sValue(0) As String
    

        sValue(0) = Store.Code
        
        If sValue(0) = "" Then Exit Function
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_PROGRAM_CLOSE", sValue(), Err_Num, Err_Dec)
        
        If RS01.EOF Then
            '
        Else
            Store.Stats = Trim(RS01!지사상태 & "")
            Store.PGCloseDate = Trim(RS01!pg사용종료일 & "")
        End If
        
        RS01.Close:     Set RS01 = Nothing
        Exit Function
            

End Function

