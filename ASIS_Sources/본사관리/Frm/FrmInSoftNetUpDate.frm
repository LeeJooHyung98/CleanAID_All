VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form FrmInSoftNetUpdate 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Update Server <-> Client"
   ClientHeight    =   4710
   ClientLeft      =   2040
   ClientTop       =   5670
   ClientWidth     =   7470
   Icon            =   "FrmInSoftNetUpDate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7470
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   180
      Picture         =   "FrmInSoftNetUpDate.frx":08CA
      ScaleHeight     =   3195
      ScaleWidth      =   1335
      TabIndex        =   19
      Top             =   210
      Width           =   1395
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Setting"
      Height          =   375
      Left            =   150
      TabIndex        =   6
      Top             =   3570
      Width           =   1425
   End
   Begin VB.CommandButton Command2 
      Caption         =   "종료"
      Height          =   375
      Left            =   6330
      TabIndex        =   3
      Top             =   3570
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "업데이트"
      Height          =   375
      Left            =   5115
      TabIndex        =   2
      Top             =   3570
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   " [ Setting ]"
      Height          =   3375
      Left            =   1680
      TabIndex        =   7
      Top             =   105
      Visible         =   0   'False
      Width           =   5745
      Begin VB.CommandButton Command6 
         Caption         =   "Help"
         Height          =   375
         Index           =   0
         Left            =   75
         TabIndex        =   20
         Top             =   2325
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "저장"
         Height          =   375
         Left            =   2610
         TabIndex        =   17
         Top             =   2310
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "취소"
         Height          =   375
         Left            =   3840
         TabIndex        =   16
         Top             =   2310
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  '평면
         Height          =   315
         Index           =   3
         Left            =   1740
         TabIndex        =   15
         ToolTipText     =   "abc.exe"
         Top             =   1860
         Width           =   3900
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  '평면
         Height          =   315
         Index           =   2
         Left            =   1740
         TabIndex        =   14
         ToolTipText     =   "c:\abc\"
         Top             =   1350
         Width           =   3900
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  '평면
         Height          =   315
         Index           =   1
         Left            =   1740
         TabIndex        =   13
         ToolTipText     =   "abc.exe"
         Top             =   840
         Width           =   3900
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  '평면
         Height          =   315
         Index           =   0
         Left            =   1770
         TabIndex        =   12
         ToolTipText     =   "http://www.abc.com/"
         Top             =   360
         Width           =   3900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "업데이트파일명"
         Height          =   180
         Index           =   3
         Left            =   270
         TabIndex        =   11
         Top             =   1950
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "업데이트폴더"
         Height          =   180
         Index           =   2
         Left            =   270
         TabIndex        =   10
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "파일명   (Server)"
         Height          =   180
         Index           =   1
         Left            =   270
         TabIndex        =   9
         Top             =   930
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "파일위치(Server)"
         Height          =   180
         Index           =   0
         Left            =   270
         TabIndex        =   8
         Top             =   420
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "받는상태"
      Height          =   3375
      Left            =   1680
      TabIndex        =   0
      Top             =   105
      Width           =   5730
      Begin Threed.SSPanel SSPanel_Process 
         Height          =   375
         Left            =   90
         TabIndex        =   18
         Top             =   2940
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   661
         _Version        =   262144
         BackColor       =   12632256
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "SSPanel1"
         BevelOuter      =   0
         BevelInner      =   1
         FloodType       =   1
         RoundedCorners  =   0   'False
      End
      Begin VB.TextBox Text1 
         Height          =   2265
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   5385
      End
      Begin VB.Label SizeInfo 
         AutoSize        =   -1  'True
         Caption         =   "None"
         Height          =   180
         Left            =   1200
         TabIndex        =   5
         Top             =   2700
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "파일크기 : "
         Height          =   180
         Left            =   240
         TabIndex        =   4
         Top             =   2700
         Width           =   900
      End
   End
   Begin 크린에이드.DownLoad Dn 
      Left            =   2010
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label Label5 
      Caption         =   "  - 주소 : 경기도 남양주시 진접읍 내각리 726-11 "
      Height          =   270
      Left            =   75
      TabIndex        =   23
      Top             =   4410
      Width           =   7320
   End
   Begin VB.Label Label4 
      Caption         =   " - 전화 : 031) 522-2000  - Fax : 031)0522-2085"
      Height          =   225
      Left            =   3300
      TabIndex        =   22
      Top             =   4110
      Width           =   4095
   End
   Begin VB.Label Label3 
      Caption         =   "www.clean-aid.co.kr"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   60
      MouseIcon       =   "FrmInSoftNetUpDate.frx":2512
      MousePointer    =   99  '사용자 정의
      TabIndex        =   21
      Top             =   4050
      Width           =   3105
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   15
      X2              =   7455
      Y1              =   3990
      Y2              =   3990
   End
End
Attribute VB_Name = "FrmInSoftNetUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'ftp지원안됩니다....
'박대영
'2001년11월27일 01:00

Option Explicit

Dim MyVersion   As String
Dim TempDbSize  As Double

Private Sub Command1_Click()
    Dim RecentVersion As String
    Dim RecentMemo As String
    '웹상에서 최신버젼이 얼마인지 받아옵니다.
    Frame2.Visible = False
    'RecentVersion = OpenURL("http://www.philipn.com/program/ver.txt", 1000)
    'RecentMemo = OpenURL("http://www.philipn.com/program/memo.txt", 1000)
    If Right(Text2(0).Text, 1) = "/" Then
        RecentVersion = OpenURL(Text2(0).Text & "Ver.txt", 1000)
        RecentMemo = OpenURL(Text2(0).Text & "memo.txt", 1000)
    Else
        RecentVersion = OpenURL(Text2(0).Text & "/Ver.txt", 1000)
        RecentMemo = OpenURL(Text2(0).Text & "/memo.txt", 1000)
    End If
    ShowInfo "최신버젼 : " & RecentVersion
    ShowInfo " "
    'UpGrade
    MyVersion = GetSetting(scRegAppname, scRegSection, "VerSion", MyVersion)
    '현재 버젼이 최신버젼이면 메세지 출력 및 Exit Sub
    If Val(MyVersion) >= Val(RecentVersion) Then
        ShowInfo "최신 버젼 입니다...."
        ShowInfo "업데이트할 필요가 없습니다."
        Text1.Refresh
        'Call Delay(3)
        'Call Command_Shell
        'End
        Exit Sub
    End If
    
    MyVersion = Val(RecentVersion)
    ShowInfo "변경내용 : " & RecentMemo
    ShowInfo " "
    ShowInfo "업데이트된 파일을 받겠습니다."
    
    '받을 파일의 주소
    Dn.Url = "http://www.clean-aid.co.kr:8090/business/백상영업.exe"
    'Dn.Url = "http://www.clean-aid.co.kr:8090/business/a.html"
    
'    If Right(Text2(0).Text, 1) = "/" Then
'        Dn.Url = Text2(0).Text & Text2(1).Text
'    Else
'        Dn.Url = Text2(0).Text & "/" & Text2(1).Text
'    End If
    '파일 정보를 읽어옴
    Dn.GetFileInformation
    '파일이 없으면 메세지 출력 후 Exit Sub
    If Dn.FileSize <= 0 Then
        ShowInfo "업데이트할 파일을 찾을수 없습니다."
        ShowInfo "업데이트 프로그램을 종료 합니다."
        Text1.Refresh
        'Call Delay(3)
        'Call Command_Shell
        'End
        Exit Sub
    Else
        '있으면 프로그레스바의 Max를 파일 사이즈로 설정 후 파일을 받을 폴더 지정
        
        TempDbSize = Dn.FileSize
        SSPanel_Process.FloodPercent = 0
        If Right(Text2(2).Text, 1) = "\" Then
            Dn.SaveLocation = Text2(2).Text & Text2(3).Text
        Else
            Dn.SaveLocation = Text2(2).Text & "\" & Text2(3).Text
        End If
    End If
    
    
    '다운로드 시작!
    ShowInfo "다운로드를 시작했습니다."
    Dn.DownLoad
    DoEvents
End Sub
Private Sub Command2_Click()
    '종료
    Unload Me
End Sub

Private Sub Command3_Click()
    Frame2.Visible = True
    Frame2.ZOrder 0
    Text2(0).SetFocus
End Sub

Private Sub Command4_Click()
    Frame2.Visible = False
End Sub

Private Sub Command5_Click()
    Frame2.Visible = False
    
    SaveSetting scRegAppname, scRegSection, "Url", Text2(0).Text
    SaveSetting scRegAppname, scRegSection, "Name", Text2(1).Text
    SaveSetting scRegAppname, scRegSection, "Fold", Text2(2).Text
    SaveSetting scRegAppname, scRegSection, "FoldName", Text2(3).Text
End Sub

Private Sub Command6_Click(Index As Integer)
    FrmInSoftNetUpdateHelp.Show 1
End Sub

Private Sub Dn_DLComplete()
    '다운로드가 완료되었을때 발생하는 이벤트
    '다운로드가 완료되면 메세지 표시후 실행여부 확인 및 실행
    ShowInfo "완료되었습니다."
    SaveSetting scRegAppname, scRegSection, "VerSion", MyVersion
    
    MsgBox "업데이트 자료를 다운로드 했습니다." & Chr(13) & "업데이트 적용합니다..", vbInformation, "OK"
    
    Call Command_Shell
    End
End Sub
Private Sub Dn_RecievedBytes(lnumBYTES As Long)
    '프로그레스바에 표시
    
    SSPanel_Process.FloodPercent = (lnumBYTES / TempDbSize) * 100
    If SSPanel_Process.FloodPercent > 49 Then
        SSPanel_Process.ForeColor = vbWhite
    Else
        SSPanel_Process.ForeColor = vbBlack
    End If
    '파일 용량 표시
    SizeInfo = CStr(lnumBYTES) & "/" & CStr(Dn.FileSize)
End Sub

Private Sub Form_Activate()


'    Text2(0).Text = GetSetting(scRegAppname, scRegSection, "Url", Text2(0).Text)
'    Text2(1).Text = GetSetting(scRegAppname, scRegSection, "Name", Text2(1).Text)
'    Text2(2).Text = GetSetting(scRegAppname, scRegSection, "Fold", Text2(2).Text)
'    Text2(3).Text = GetSetting(scRegAppname, scRegSection, "FoldName", Text2(3).Text)
'    If Trim(Text2(0).Text) = "" Then
'        Text2(0).Text = scUrl
'        Text2(1).Text = scName
'        Text2(2).Text = scFold
'        Text2(3).Text = scFoldName
'        Call Command5_Click
'    End If

End Sub

Private Sub Form_Load()
Dim MyString As String
Dim DownFile As String
    On Error GoTo ERR_RTN
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    If Dir("Upgrad.txt") <> "" Then
        Text1.Text = ""
        DownUrl_FHandle = FreeFile
        DownFile = Dir("info.txt")
        If DownFile <> "" Then
            Open DownFile For Input As #DownUrl_FHandle     ' Open file for input.
            Do While Not EOF(1)                             ' Loop until end of file.
                Input #DownUrl_FHandle, MyString            ' Read data into two variables.
                ShowInfo MyString
            Loop
            Close #DownUrl_FHandle    ' Close file.
        End If
    Else
        ShowInfo "프로그램이 수정 되었습니다."
        ShowInfo ""
        ShowInfo "최신 프로그램을 자동 업데이트 하는 작업 입니다."
        ShowInfo ""
        ShowInfo "프로그램 문의 : (주)크린에이드본사 전산실 "
        ShowInfo "전화 번호     : 031 - 522 - 2000 "
        ShowInfo "주     소     : 경기도 남양주시 진접읍 내각리 726-11 "
        ShowInfo ""
    End If
    
    Text2(0).Text = GetSetting(scRegAppname, scRegSection, "Url", Text2(0).Text)
    Text2(1).Text = GetSetting(scRegAppname, scRegSection, "Name", Text2(1).Text)
    Text2(2).Text = GetSetting(scRegAppname, scRegSection, "Fold", Text2(2).Text)
    Text2(3).Text = GetSetting(scRegAppname, scRegSection, "FoldName", Text2(3).Text)
    
    
    'Url 변경 추가 시작
    If Trim(Text2(0).Text) = "http://www.clean-aid.co.kr:8090/business/" Then
        Text2(1).Text = "백상영업.exe"
        Text2(2).Text = App.Path & "\"
        Text2(3).Text = "백상영업UP.exe"
    Else

        Text2(0).Text = "http://www.clean-aid.co.kr:8090/business/"
        Text2(1).Text = "백상영업.exe"
        Text2(2).Text = App.Path & "\"
        Text2(3).Text = "백상영업UP.exe"
        
        SaveSetting scRegAppname, scRegSection, "Url", Text2(0).Text
        SaveSetting scRegAppname, scRegSection, "Name", Text2(1).Text
        SaveSetting scRegAppname, scRegSection, "Fold", Text2(2).Text
        SaveSetting scRegAppname, scRegSection, "FoldName", Text2(3).Text

    End If
    'Url 변경 추가 끝
    
    If Trim(Text2(0).Text) = "" Then
        Text2(0).Text = scUrl
        Text2(1).Text = scName
        Text2(2).Text = scFold
        Text2(3).Text = scFoldName
        Call Command5_Click
    End If
    
    Exit Sub
ERR_RTN:
    MsgBox Err.Description, vbInformation, "Error"
    
End Sub
Function ShowInfo(body As String) ' 텍스트박스에 작업내용 표시하는 함수
    Text1 = Text1 & body & vbCrLf
    Text1.SelStart = Len(Text1)
End Function
Private Sub Command_Shell()
    On Error Resume Next
    If Right(Text2(2).Text, 1) = "\" Then
        Shell Text2(2).Text & Text2(3).Text, vbNormalFocus
    Else
        Shell Text2(2).Text & "\" & Text2(3).Text, vbNormalFocus
    End If
End Sub
Private Sub Delay(ByVal utime As Single)
    Dim starttime As Single
    starttime = Timer
    Do
    Loop While Timer < starttime + utime
    
End Sub

Private Sub Label3_Click()
    ShellExecute hwnd, "Open", "http://www.insoftnet.com", "", App.Path, 1
End Sub

