VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Program Update Server <-> Client"
   ClientHeight    =   3990
   ClientLeft      =   3090
   ClientTop       =   2955
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6840
   Begin 크린에이드.DownLoad Dn 
      Left            =   1860
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   180
      Picture         =   "Form1.frx":0000
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
      Left            =   5640
      TabIndex        =   3
      Top             =   3570
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "업데이트"
      Height          =   375
      Left            =   4455
      TabIndex        =   2
      Top             =   3570
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "받는상태"
      Height          =   3375
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin Threed.SSPanel SSPanel_Process 
         Height          =   375
         Left            =   90
         TabIndex        =   18
         Top             =   2940
         Width           =   4905
         _ExtentX        =   8652
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
         Width           =   4575
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
   Begin VB.Frame Frame2 
      Caption         =   " [ Setting ]"
      Height          =   3375
      Left            =   1680
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   5055
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
         Width           =   3225
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  '평면
         Height          =   315
         Index           =   2
         Left            =   1740
         TabIndex        =   14
         ToolTipText     =   "c:\abc\"
         Top             =   1350
         Width           =   3225
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  '평면
         Height          =   315
         Index           =   1
         Left            =   1740
         TabIndex        =   13
         ToolTipText     =   "abc.exe"
         Top             =   840
         Width           =   3225
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  '평면
         Height          =   315
         Index           =   0
         Left            =   1740
         TabIndex        =   12
         ToolTipText     =   "http://www.abc.com/"
         Top             =   360
         Width           =   3225
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
End
Attribute VB_Name = "Form1"
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
    
    'RecentVersion = OpenURL("http://www.insoftnet/upgrade/cleanaid/clever.txt", 1000)
    'RecentMemo = OpenURL("http://www.insoftnet/upgrade/cleanaid/memo.txt", 1000)
    RecentVersion = OpenURL(Text2(0).Text & "clever.txt", 1000)
    RecentMemo = OpenURL(Text2(0).Text & "memo.txt", 1000)
    ShowInfo "최신버젼 : " & RecentVersion
    ShowInfo " "
    
    MyVersion = GetSetting("CleanAid", "FileLoc", "VerSion", MyVersion)
    '현재 버젼이 최신버젼이면 메세지 출력 및 Exit Sub
    If Val(MyVersion) >= Val(RecentVersion) Then
        ShowInfo "업데이트할 필요가 없습니다."
        ShowInfo "업데이트 프로그램을 종료 합니다."
        Text1.Refresh
        Call Delay(3)
        Call Command_Shell
        End
    End If
    
    MyVersion = Val(RecentVersion)
    ShowInfo "변경내용 : " & RecentMemo
    ShowInfo " "
    ShowInfo "업데이트된 파일을 받겠습니다."
    
    '받을 파일의 주소
    'Dn.Url = "http://www.insoftnet/upgrade/cleanaid/sale.exe"
    Dn.Url = Text2(0).Text & Text2(1).Text
    '파일 정보를 읽어옴
    Dn.GetFileInformation
    '파일이 없으면 메세지 출력 후 Exit Sub
    If Dn.FileSize <= 0 Then
        ShowInfo "업데이트할 파일을 찾을수 없습니다."
        ShowInfo "업데이트 프로그램을 종료 합니다."
        Text1.Refresh
        Call Delay(3)
        Call Command_Shell
        End
    Else
        '있으면 프로그레스바의 Max를 파일 사이즈로 설정 후 파일을 받을 폴더 지정
        
        TempDbSize = Dn.FileSize
        SSPanel_Process.FloodPercent = 0
        'Dn.SaveLocation = "c:\windows\바탕 화면\sale.exe"
        Dn.SaveLocation = Text2(2).Text & Text2(3).Text
    End If
    
    
    '다운로드 시작!
    ShowInfo "다운로드를 시작했습니다."
    Dn.DownLoad
    DoEvents
End Sub
Private Sub Command2_Click()
    '종료
    End
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
    SaveSetting "CleanAid", "FileLoc", "Url", Text2(0).Text
    SaveSetting "CleanAid", "FileLoc", "Name", Text2(1).Text
    SaveSetting "CleanAid", "FileLoc", "Fold", Text2(2).Text
    SaveSetting "CleanAid", "FileLoc", "FoldName", Text2(3).Text
End Sub

Private Sub Dn_DLComplete()
    '다운로드가 완료되었을때 발생하는 이벤트
    '다운로드가 완료되면 메세지 표시후 실행여부 확인 및 실행
    ShowInfo "완료되었습니다."
    SaveSetting "Recycle", "FileLoc", "VerSion", MyVersion
    If MsgBox("업데이트가 완료되었습니다. 실행하시겠습니까?", vbYesNo, "알림") = vbYes Then
        Call Command_Shell
    End If
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
    
    Text2(0).Text = GetSetting("CleanAid", "FileLoc", "Url", Text2(0).Text)
    Text2(1).Text = GetSetting("CleanAid", "FileLoc", "Name", Text2(1).Text)
    Text2(2).Text = GetSetting("CleanAid", "FileLoc", "Fold", Text2(2).Text)
    Text2(3).Text = GetSetting("CleanAid", "FileLoc", "FoldName", Text2(3).Text)
    If Trim(Text2(0).Text) = "" Then
        Text2(0).Text = "www.insoftnet.com/upgrade/cleanaid/"
        Text2(1).Text = "백상.exe"
        Text2(2).Text = "C:\백상\"
        Text2(3).Text = "백상.exe"
        Call Command5_Click
    End If

End Sub

Private Sub Form_Load()
Dim MyString As String
Dim DownFile As String
    If Dir("Upgrad.txt") <> "" Then
        Text1.Text = ""
        FHandle = FreeFile
        DownFile = Dir("info.txt")
        Open DownFile For Input As #FHandle   ' Open file for input.
        Do While Not EOF(1)             ' Loop until end of file.
            Input #FHandle, MyString          ' Read data into two variables.
            ShowInfo MyString
        Loop
        Close #FHandle    ' Close file.
    Else
        ShowInfo "upgrad.txt 파일이 없습니다."
        ShowInfo ""
        ShowInfo "최신 프로그램을 자동 업데이트 하는 작업 입니다."
        ShowInfo "프로그램 문의 : 전산실로 문의 바랍니다. "
        ShowInfo ""
        ShowInfo ""
    End If
     
    
End Sub
Function ShowInfo(body As String) ' 텍스트박스에 작업내용 표시하는 함수
    Text1 = Text1 & body & vbCrLf
    Text1.SelStart = Len(Text1)
End Function
Private Sub Command_Shell()
    On Error Resume Next
    
    Shell Text2(2).Text & "\" & Text2(3).Text, vbNormalFocus
End Sub
Private Sub Delay(ByVal utime As Single)
    Dim starttime As Single
    starttime = Timer
    Do
    Loop While Timer < starttime + utime
    
End Sub
