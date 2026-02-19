VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{739A6F0D-BBFA-4993-B77D-B98A35DD8121}#1.0#0"; "SmartUpdateXX.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmUpdateCheck 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "업그레이드"
   ClientHeight    =   1590
   ClientLeft      =   3900
   ClientTop       =   6105
   ClientWidth     =   5850
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form33"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   1590
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5850
      _ExtentX        =   10319
      _ExtentY        =   2805
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   0
      PaneTree        =   "frmUpdateCheck.frx":0000
      Begin Threed.SSPanel SSPanel 
         Height          =   1590
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   2805
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin SmartUpdateXX.SmartUpdateX SmartUpdateX 
            Left            =   2895
            Top             =   495
            HttpEvent       =   -1  'True
            GetURL          =   ""
            Myinfo          =   ""
            MyVersion       =   "0"
            WorkDir         =   "C:\DOCUME~1\breeze\LOCALS~1\Temp\"
            IniFileName     =   "update.ini"
            Port            =   80
            Enabled         =   -1  'True
            Visible         =   -1  'True
            SourceDelete    =   -1  'True
            UpDateExePathName=   ""
            UpDateRunDelay  =   1000
            UpDateRunParameter=   ""
            UserAgentStr    =   "Mozilla/4.0 (compatible; MSIE 5.5)"
            ProxyPassword   =   ""
            ProxyPort       =   0
            ProxyServer     =   ""
            ProxyUsername   =   ""
         End
         Begin XtremeSuiteControls.ProgressBar ProgressBar 
            Height          =   285
            Left            =   120
            TabIndex        =   2
            Top             =   870
            Width           =   5610
            _Version        =   851970
            _ExtentX        =   9895
            _ExtentY        =   503
            _StockProps     =   93
            BackColor       =   14280169
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Image Image 
            Height          =   765
            Left            =   105
            Picture         =   "frmUpdateCheck.frx":0032
            Top             =   120
            Width           =   765
         End
         Begin VB.Label lblNewVersion 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "Ver 1.0.0"
            ForeColor       =   &H000000C0&
            Height          =   180
            Left            =   150
            TabIndex        =   7
            Top             =   1290
            Width           =   810
         End
         Begin VB.Label lblVersion 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "Ver 1.0.0"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   180
            Left            =   4725
            TabIndex        =   6
            Top             =   150
            Width           =   945
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "(주)크린에이드 관리 시스템"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   960
            TabIndex        =   5
            Top             =   150
            Width           =   2565
         End
         Begin VB.Label lblFileSize 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "#"
            Height          =   180
            Left            =   1785
            TabIndex        =   4
            Top             =   1290
            Width           =   3915
         End
         Begin VB.Label lblFile 
            BackStyle       =   0  '투명
            Caption         =   "#"
            ForeColor       =   &H00E0E0E0&
            Height          =   180
            Left            =   990
            TabIndex        =   3
            Top             =   540
            Width           =   4650
         End
      End
   End
End
Attribute VB_Name = "frmUpdateCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strVersion  As String
Dim TotalSize   As Long
Dim ConnectStop As Boolean  ' 서버에서 화일을 다운로드중에 이값이 true 이면 중지된다..

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    Dim i        As Integer
    Dim temp     As String
    Dim InfoText As String
    
    Dim strURL   As String
   
    strVersion = App.Major & "." & App.Minor & "." & App.Revision
    
    lblFile.Caption = ""
    lblFileSize.Caption = ""
    lblNewVersion.Caption = ""
    
    ProgressBar.Min = 0
    ProgressBar.Max = 100
    ProgressBar.Value = 0
    
    strURL = GetIniStr("UPDATE", "URL", "", iniFile)
    
    SmartUpdateX.MyVersion = strVersion                           ' 현재 프로그램의 버전을 설정합니다.
    
    If strURL = "" Then
        Call SetIniStr("UPDATE", "URL", "tecle.cafe24.com/update", iniFile)
        
        SmartUpdateX.GetURL = "tecle.cafe24.com/update"
    Else
        SmartUpdateX.GetURL = strURL
    End If
    
    SmartUpdateX.iniFileName = "update.ini"                       ' 대소문자를 정확히 해야합니다.. !!
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Public Sub FileDown()
'    Dim FilePath  As String
'    Dim FileName  As String
'    Dim dtemp
'
'    Dim temp       As String
'
'    temp = SmartUpdateX.ReadInfo("DOWN", "MSTFILE")
'
'    If Trim(temp) <> "" Then
'        FilePath = App.Path + "\"
'
'        dtemp = Split(temp, "/")
'
'        FileName = dtemp(UBound(dtemp))
'
'        ConnectStop = False
'
'        i = SmartUpdateX.GetFile(temp, FilePath + FileName)
'
'        If i = 1 Then '성공
'            'tty1 "다운후 저장성공 :" + FilePath + FileName
'        Else
'            'tty1 "접속에러:" + temp
'        End If
'    End If
End Sub

Public Sub SmartUpdate()
    Dim temp       As String
    Dim wwwnewexe  As String
    Dim updatetemp As String

    Dim i          As Long
    
    ConnectStop = False
    
    temp = SmartUpdateX.ReadInfo("DOWN", "MSTFILE")
  
    If Trim(temp) <> "" Then
        lblFile.Caption = temp
        
        updatetemp = "newexetemp.tmp"
       
        wwwnewexe = temp
       
        i = SmartUpdateX.GetFile(wwwnewexe, SmartUpdateX.WorkDir + updatetemp)
       
        If i = 1 Then '' 수신 성공
            'MsgBox wwwnewexe + " 다운로드 성공 임시화일  저장경로명[" + SmartUpdateX.WorkDir + updatetemp + "]"
            
            i = SmartUpdateX.SmartUpdate(SmartUpdateX.WorkDir + updatetemp)
         
            If i = 1 Then ''1 =정상,, -1=  OnProgramClose 이벤트가 없다 .. -2= 바꿔치기할 신규 실행화일 오류. 0=시스템오류
                'MsgBox "업데이트 정상" ''<= 이 메세지는 이미 프로그램이 종료했기때문에 보이지 않는다...
            Else
                'MsgBox "업데이트 에러 에러코드[" + Str(i) + "]"
            End If
        Else
            'MsgBox wwwnewexe + " 다운로드중 에러발생 !! 저장경로명[" + SmartUpdateX.WorkDir + updatetemp + "]"
        End If
    Else
        'MsgBox "업데이트정보가 없습니다 먼저 GetInfo 실행하세요"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ConnectStop = True
End Sub

''======== 스마트업데이트를 하기위해 종료이벤트가 꼭 설정되어야 합니다..
'' ===== 종료이벤트의 마지막에는 꼭 프로그램을 종료하는 코드를 실행해야 합니다.

Private Sub SmartUpdateX_OnProgramClose(ByVal NewProgram As String)
    End ''<== 프로그램 종료
End Sub

Private Sub SmartUpdateX_OnState(ByVal URL As String, ByVal StateText As String)
    'MsgBox "State[" + url + "][" + StateText + "]"
End Sub

''======= StopSW 변수는 화일을 다운받는 중간에 중지시켜야 할경우 true로 설정
Private Sub SmartUpdateX_OnWork(ByVal URL As String, ByVal WorkCount As Long, StopSW As Boolean)
    'MsgBox "Work[" + url + "]크기[" + Str(WorkCount) + "]"
    lblFileSize.Caption = Format(WorkCount, "#,##0") + " bytes / " + Format(TotalSize, "#,##0") & " bytes"
    
    ProgressBar.Value = (WorkCount / TotalSize) * 100
    
    If ConnectStop Then StopSW = True
    DoEvents
End Sub

Private Sub SmartUpdateX_OnWorkBegin(ByVal URL As String, ByVal WorkCountMax As Long, StopSW As Boolean)
    TotalSize = WorkCountMax
    
    'MsgBox "WorkBegin[" + url + "]받을 화일크기[" + Str(WorkCountMax) + "]"
  
    lblFileSize.Caption = "0 bytes / " + Format(TotalSize, "#,##0") & " bytes"
    ProgressBar.Value = 0
    
    If ConnectStop Then StopSW = True
End Sub

Private Sub SmartUpdateX_OnWorkEnd(ByVal URL As String)
    'MsgBox "WorkEnd[" + url + "] 종료"
End Sub
 
