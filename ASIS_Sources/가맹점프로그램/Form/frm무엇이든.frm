VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{739A6F0D-BBFA-4993-B77D-B98A35DD8121}#1.0#0"; "SmartUpdateXX.ocx"
Begin VB.Form frm무엇이든 
   Caption         =   "무엇이든 물어보세요."
   ClientHeight    =   12360
   ClientLeft      =   3780
   ClientTop       =   2955
   ClientWidth     =   15000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12360
   ScaleWidth      =   15000
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   12360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   21802
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      PaneTree        =   "frm무엇이든.frx":0000
      Begin Threed.SSPanel SSPanel3 
         Height          =   8100
         Left            =   30
         TabIndex        =   9
         Top             =   4230
         Width           =   14940
         _ExtentX        =   26353
         _ExtentY        =   14288
         _Version        =   262144
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Label Label1 
            Caption         =   "     1. 다음을 클릭                 2. 설치 클릭                 3 마침 클릭"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   5
            Left            =   420
            TabIndex        =   15
            Top             =   4530
            Width           =   13995
         End
         Begin VB.Image Image1 
            Height          =   5100
            Left            =   570
            Picture         =   "frm무엇이든.frx":0092
            Top             =   240
            Width           =   14025
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   2550
         Left            =   30
         TabIndex        =   4
         Top             =   1665
         Width           =   14940
         _ExtentX        =   26353
         _ExtentY        =   4498
         _Version        =   262144
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin SmartUpdateXX.SmartUpdateX SmartUpdateX 
            Left            =   60
            Top             =   30
            HttpEvent       =   -1  'True
            GetURL          =   ""
            Myinfo          =   ""
            MyVersion       =   "0"
            WorkDir         =   "C:\DOCUME~1\ADMINI~1\LOCALS~1\Temp\"
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
         Begin XtremeSuiteControls.PushButton PushButton2 
            Height          =   855
            Left            =   240
            TabIndex        =   5
            Top             =   630
            Width           =   5745
            _Version        =   851970
            _ExtentX        =   10134
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "다운로드"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ProgressBar ProgressBar 
            Height          =   285
            Left            =   240
            TabIndex        =   6
            Top             =   330
            Width           =   5760
            _Version        =   851970
            _ExtentX        =   10160
            _ExtentY        =   503
            _StockProps     =   93
            BackColor       =   14280169
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton PushButton3 
            Height          =   855
            Left            =   240
            TabIndex        =   7
            Top             =   1500
            Width           =   5745
            _Version        =   851970
            _ExtentX        =   10134
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "설치"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "(어도비 아크로뱃 리더)"" v8.1 한글판  설치"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   4
            Left            =   6240
            TabIndex        =   14
            Top             =   600
            Width           =   6555
         End
         Begin VB.Label Label1 
            Caption         =   "2. 설치 버튼을 클릭 하여 설치를 진행한다."
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   3
            Left            =   6270
            TabIndex        =   13
            Top             =   1620
            Width           =   9405
         End
         Begin VB.Label Label1 
            Caption         =   "1. 다운로드 클릭후 다운로드가 완료 될때 까지 기다린다."
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   2
            Left            =   6270
            TabIndex        =   12
            Top             =   1110
            Width           =   9405
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "PDF 뷰어 ""한글 Adobe Acrobat Reader"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   0
            Left            =   6270
            TabIndex        =   10
            Top             =   270
            Width           =   5715
         End
         Begin VB.Label lblFileSize 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "#"
            Height          =   180
            Left            =   300
            TabIndex        =   8
            Top             =   90
            Width           =   5595
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   1140
         Left            =   30
         TabIndex        =   2
         Top             =   510
         Width           =   14940
         _ExtentX        =   26353
         _ExtentY        =   2011
         _Version        =   262144
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton PushButton1 
            Height          =   855
            Left            =   270
            TabIndex        =   3
            Top             =   150
            Width           =   5745
            _Version        =   851970
            _ExtentX        =   10134
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "무엇이든 물어보세요."
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label Label1 
            Caption         =   "내용이 보이지 않을 경우 아래의 내용의 설치 항목을 설치 하여 주십시요"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Index           =   1
            Left            =   6270
            TabIndex        =   11
            Top             =   360
            Width           =   8595
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   465
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   14940
         _ExtentX        =   26353
         _ExtentY        =   820
         _Version        =   262144
         Font3D          =   1
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "      무엇이든 물어보세요....."
         PictureBackgroundStyle=   2
         PictureBackground=   "frm무엇이든.frx":12C4B
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton btnExit 
            Height          =   405
            Left            =   12420
            TabIndex        =   16
            Top             =   30
            Width           =   2175
            _Version        =   851970
            _ExtentX        =   3836
            _ExtentY        =   714
            _StockProps     =   79
            Caption         =   "닫기"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm무엇이든.frx":12E71
            Top             =   -15
            Width           =   765
         End
      End
   End
End
Attribute VB_Name = "frm무엇이든"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strVersion  As String
Dim TotalSize   As Long
Dim ConnectStop As Boolean  ' 서버에서 화일을 다운로드중에 이값이 true 이면 중지된다..
    
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i        As Integer
    Dim temp     As String
    Dim InfoText As String
   
    'strVersion = App.Major & "." & App.Minor & "." & App.Revision
    
'    lblFile.Caption = ""
'    lblFileSize.Caption = ""
'    lblNewVersion.Caption = ""
    
    ProgressBar.Min = 0
    ProgressBar.MAX = 100
    ProgressBar.Value = 0
    
    lblFileSize.Caption = ""
    
    If Dir(App.Path & "\AdbeRdr810_ko_KR.msi", vbDirectory) = "" Then
        PushButton3.Enabled = False
    Else
        PushButton3.Enabled = True
    End If
    
    'SmartUpdateX.MyVersion = strVersion                           ' 현재 프로그램의 버전을 설정합니다.
    SmartUpdateX.GetURL = GetIniStr("UPDATE", "URL", "", iniFile) ' "tecle.hosting.paran.com/smartupdate"
    SmartUpdateX.IniFileName = "update.ini"                       ' 대소문자를 정확히 해야합니다.. !!

End Sub

Private Sub PushButton1_Click()
 
    Dim strFile As String
    strFile = App.Path & "\CleanAidHelp.pdf"
    
    ShellExecute Me.hWnd, vbNullString, """" & strFile & """", vbNullString, vbNullString, 3
 
End Sub

Private Sub PushButton2_Click()
    Dim i As Integer
    Dim InfoText   As String

    i = SmartUpdateX.GetInfo 'Geturl 에서 버전 정보화일을 읽어옵니다
  
    If i = 1 Then '접속 성공 - 업데이트 정보 가저옴
        InfoText = SmartUpdateX.GetInfoText '서버에서 가저온 IniFileName 모든 내용을 보여줍니다
     
        Open AppPath & "Update.ini" For Output As #1
    
        Print #1, InfoText
        Close #1
    End If
    
    lblFileSize.Caption = ""
    If Dir(App.Path & "\AdbeRdr810_ko_KR.msi", vbDirectory) = "" Then
        PushButton3.Enabled = False
        Call AdbeRdr810_FileDown     '파일 다운로드
        DoEvents
    Else
        PushButton3.Enabled = True
    
    End If
    
End Sub


Public Sub AdbeRdr810_FileDown()

    Dim FilePath  As String
    Dim FileName  As String
    Dim dtemp
   
    Dim temp       As String
   
    temp = SmartUpdateX.ReadInfo("DOWN", "AdobeAcrobatFILE")
    
 
    If Trim(temp) <> "" Then
        FilePath = App.Path + "\"
      
        dtemp = Split(temp, "/")
        
        FileName = dtemp(UBound(dtemp))

'        ConnectStop = False
        
        i = SmartUpdateX.GetFile(temp, FilePath + FileName)
      
        If i = 1 Then '성공
            'tty1 "다운후 저장성공 :" + FilePath + FileName
        Else
            'tty1 "접속에러:" + temp
        End If
    End If
    
    
End Sub


Private Sub PushButton3_Click()
    Dim strFile As String
    strFile = App.Path & "\AdbeRdr810_ko_KR.msi"
    
    ShellExecute Me.hWnd, vbNullString, """" & strFile & """", vbNullString, vbNullString, 3

End Sub

''======== 스마트업데이트를 하기위해 종료이벤트가 꼭 설정되어야 합니다..
'' ===== 종료이벤트의 마지막에는 꼭 프로그램을 종료하는 코드를 실행해야 합니다.

Private Sub SmartUpdateX_OnProgramClose(ByVal NewProgram As String)
    End ''<== 프로그램 종료
End Sub

Private Sub SmartUpdateX_OnState(ByVal Url As String, ByVal StateText As String)
    'MsgBox "State[" + url + "][" + StateText + "]"
End Sub

''======= StopSW 변수는 화일을 다운받는 중간에 중지시켜야 할경우 true로 설정
Private Sub SmartUpdateX_OnWork(ByVal Url As String, ByVal WorkCount As Long, StopSW As Boolean)
    'MsgBox "Work[" + url + "]크기[" + Str(WorkCount) + "]"
    lblFileSize.Caption = Format(WorkCount, "#,##0") + " bytes / " + Format(TotalSize, "#,##0") & " bytes"
    
    ProgressBar.Value = (WorkCount / TotalSize) * 100
    
    If ConnectStop Then StopSW = True
    DoEvents
End Sub

Private Sub SmartUpdateX_OnWorkBegin(ByVal Url As String, ByVal WorkCountMax As Long, StopSW As Boolean)
    TotalSize = WorkCountMax
    
    'MsgBox "WorkBegin[" + url + "]받을 화일크기[" + Str(WorkCountMax) + "]"
  
    lblFileSize.Caption = "0 bytes / " + Format(TotalSize, "#,##0") & " bytes"
    ProgressBar.Value = 0
    
    If ConnectStop Then StopSW = True
End Sub

Private Sub SmartUpdateX_OnWorkEnd(ByVal Url As String)
    'MsgBox "WorkEnd[" + url + "] 종료"
    PushButton3.Enabled = True
End Sub
 
