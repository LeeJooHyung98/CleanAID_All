VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form frmBackup 
   ClientHeight    =   7005
   ClientLeft      =   1485
   ClientTop       =   720
   ClientWidth     =   11850
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   11850
   WindowState     =   2  '최대화
   Begin VB.Frame Frame1 
      Height          =   6990
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   11805
      Begin ComCtl2.Animation Ani1 
         Height          =   660
         Left            =   2790
         TabIndex        =   3
         Top             =   2850
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   1164
         _Version        =   327681
         FullWidth       =   431
         FullHeight      =   44
      End
      Begin ComctlLib.ProgressBar ProgBar 
         Height          =   420
         Left            =   2790
         TabIndex        =   2
         Top             =   3540
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   741
         _Version        =   327682
         Appearance      =   1
      End
      Begin Threed.SSCommand CmdBackup 
         Height          =   1020
         Left            =   6120
         TabIndex        =   1
         Top             =   4005
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1799
         _Version        =   262144
         ForeColor       =   12582912
         PictureFrames   =   1
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmBackup.frx":0000
         Caption         =   "저 장"
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   6
         BevelWidth      =   3
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   1020
         Left            =   2805
         TabIndex        =   7
         Top             =   4005
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   1799
         _Version        =   262144
         ForeColor       =   12582912
         PictureFrames   =   1
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmBackup.frx":08DA
         Caption         =   "자료 정리"
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   6
      End
      Begin VB.Label LblMsg 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   525
         Left            =   3120
         TabIndex        =   4
         Top             =   1935
         Width           =   5805
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "궁서체"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   0
         Left            =   2730
         TabIndex        =   5
         Top             =   1710
         Width           =   6390
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "Label1"
         Height          =   915
         Left            =   2865
         TabIndex        =   6
         Top             =   1815
         Width           =   6345
      End
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MyFile As String
Dim cCnt As Integer
Dim dCnt As Integer
Dim D_No As Integer

Private Function DskCHK() As Boolean
    On Error GoTo ErrRtn
    
    Open "A:\CHK" For Random As #1
    
    Close #1
    
    Kill "A:\CHK"
    
    DskCHK = True
    
    Exit Function
        
ErrRtn:
    DskCHK = False
End Function

Private Sub Form_Load()
    'TitleSet "BACKUP"
    LblMsg.Caption = ""
End Sub

Private Sub CmdBackup_Click()
    Dim S_File As String
    Dim D_File As String
    
    If cCnt = dCnt Then Exit Sub
    
    If Not DskCHK Then
        LblMsg = "A:드라이버에 디스켓이 없습니다"
        Exit Sub
    End If
    
    S_File = App.Path & "\DB\BackData.A"
    D_File = "A:\BackData.A"
    
    If cCnt = 0 Then
        S_File = S_File + "RJ"
        D_File = D_File + "RJ"
    Else
        If cCnt < 10 Then
            S_File = S_File & "0" & cCnt
            D_File = D_File & "0" & cCnt
        Else
            S_File = S_File & cCnt
            D_File = D_File & cCnt
        End If
    End If
    
    On Error GoTo DiskError
    
    LblMsg.Caption = "자료 저장중..!"
    
    Ani1.Visible = True
    Ani1.AutoPlay = True
    Ani1.Open App.Path & "\image\filecopy.avi"
    DoEvents
    
    FileCopy S_File, D_File
    DoEvents
    
    Ani1.Visible = False
    cCnt = cCnt + 1
    
    If cCnt = dCnt Then
        LblMsg.Caption = "저장이 완료되었습니다..!"
    Else
        D_No = cCnt + 1
        LblMsg.Caption = D_No & "번 디스켓을 넣고 저장버튼을 누르세요..!"
    End If
    
    Exit Sub
    
DiskError:
    Ani1.Visible = False
    MsgBox "디스켓이 불량합니다..!", vbInformation, "BACKUP"
    LblMsg.Caption = D_No & "번 디스켓을 넣고 저장버튼을 누르세요..!"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not Dir(App.Path & "\BACKUP.OK") = "" Then
        Kill App.Path & "\BACKUP.OK"
    End If
End Sub

Private Sub SSCommand1_Click()
    cCnt = 0
    dCnt = 0

    LblMsg.Caption = "자료를 정리중입니다..!"

    cmdBackup.Enabled = False

    ProgBar.Visible = True
    ProgBar.MAX = 1000
    ProgBar.Min = 0
    ProgBar.Value = 0

    If Dir(App.Path & "\BACKUP.BAT") = "" Then
        MsgBox "BACKUP.BAT 파일이 없어서 이명령을 사용할 수 없습니다.", vbInformation, "오류"
        Exit Sub
    End If
    
    If Not Dir(App.Path & "\BACKUP.OK") = "" Then
        Kill App.Path & "\BACKUP.OK"
    End If
    
    ' 이전 자료가 있을 경우 삭제 여부 확인
    If Not Dir(App.Path & "\DB\BackData.Arj") = "" Then
        If MsgBox("이전 백업 자료가 있습니다. 삭제 하시겠습니까 ?", vbYesNo, "삭제확인") = vbYes Then
            Kill App.Path & "\db\BackData.*"
        Else
            Exit Sub
        End If
    End If
    
    Shell App.Path & "\BACKUP.BAT", 0
    ProgBar.MAX = Int(ShowFolderSize(m_DBPath) / 2048)
    
    Do While Dir(App.Path & "\Backup.OK") = ""
        ProgBar.Value = ProgBar.Value + 1

        If ProgBar.Value = ProgBar.MAX Then
            ProgBar.Value = 0
        End If

        DoEvents
    Loop
    ProgBar.Value = ProgBar.MAX

    If Dir(App.Path & "\DB\BackData.arj") = "" Then
        LblMsg.Caption = "자료가 없습니다..!"
        Exit Sub
    End If

    MyFile = Dir(App.Path & "\DB\BackData.*")
    dCnt = 1

    While Not Dir = ""
        dCnt = dCnt + 1
    Wend

    LblMsg.Caption = ""
    MsgBox "공디스켓 " & dCnt & "장이 필요합니다.", vbInformation, "BACKUP"

    cmdBackup.Enabled = True

    D_No = 1
    If dCnt = 1 Then
        LblMsg.Caption = "디스켓을 넣고 저장버튼을 누르세요..!"
    Else
        LblMsg.Caption = D_No & "번 디스켓을 넣고 저장버튼을 누르세요..!"
    End If
    cCnt = 0

End Sub
