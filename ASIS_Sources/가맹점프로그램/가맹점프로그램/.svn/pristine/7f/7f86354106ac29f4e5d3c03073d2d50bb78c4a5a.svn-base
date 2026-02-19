VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form frmRestore 
   ClientHeight    =   6990
   ClientLeft      =   3735
   ClientTop       =   6510
   ClientWidth     =   11835
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   11835
   WindowState     =   2  '최대화
   Begin VB.Frame Frame1 
      Height          =   6945
      Left            =   90
      TabIndex        =   2
      Top             =   0
      Width           =   11700
      Begin ComCtl2.Animation Ani1 
         Height          =   660
         Left            =   2640
         TabIndex        =   4
         Top             =   4200
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   1164
         _Version        =   327681
         FullWidth       =   421
         FullHeight      =   44
      End
      Begin ComctlLib.ProgressBar ProgBar 
         Height          =   420
         Left            =   2640
         TabIndex        =   3
         Top             =   4935
         Visible         =   0   'False
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   741
         _Version        =   327682
         Appearance      =   1
      End
      Begin Threed.SSCommand CmdCopy 
         Height          =   1020
         Left            =   3690
         TabIndex        =   0
         Top             =   2745
         Width           =   1905
         _ExtentX        =   3360
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
         Picture         =   "frmRestore.frx":0000
         Caption         =   "COPY"
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   6
         BevelWidth      =   3
      End
      Begin Threed.SSCommand CmdRestore 
         Height          =   1020
         Left            =   5625
         TabIndex        =   1
         Top             =   2745
         Width           =   1905
         _ExtentX        =   3360
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
         Picture         =   "frmRestore.frx":1E52
         Caption         =   "RESTORE"
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   6
         BevelWidth      =   3
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
         Left            =   2235
         TabIndex        =   6
         Top             =   1290
         Width           =   6705
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
         Height          =   930
         Index           =   0
         Left            =   2085
         TabIndex        =   7
         Top             =   1110
         Width           =   7005
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Label1"
         Height          =   990
         Left            =   2205
         TabIndex        =   8
         Top             =   1230
         Width           =   7005
      End
      Begin VB.Label Label1 
         Caption         =   "(디스켓이 1장 이상인경우 전부 COPY한후 RESTORE 버튼을 누르세요.)"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   405
         Left            =   1680
         TabIndex        =   5
         Top             =   5655
         Width           =   8400
      End
   End
End
Attribute VB_Name = "frmRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Option Explicit
''
''Dim MyFile As String
''Dim cCnt As Integer
''Dim dCnt As Integer
''
''Private Function DskCHK() As Boolean
''    On Error GoTo ErrRtn
''    Open "A:\CHK" For Random As #1
''    Close #1
''    Kill "A:\CHK"
''
''    DskCHK = True
''    Exit Function
''
''ErrRtn:
''    DskCHK = False
''End Function
''
''Private Sub CmdCopy_Click()
''    Dim Filename As String
''
''    If Not DskCHK Then
''        LblMsg = "A:드라이버에 디스켓이 없습니다"
''        Exit Sub
''    End If
''
''    On Error GoTo DiskError
''
''    If Dir("A:\Backdata.*") = "" Then
''        LblMsg.Caption = "디스켓에 자료가 없습니다..!"
''        Exit Sub
''    End If
''
''        ' 이전 자료가 있을 경우 삭제 여부 확인
''    If Not Dir(App.Path & "\DB\BackData.Arj") = "" Then
''        If MsgBox("이전 백업 자료가 있습니다. 삭제 하시겠습니까 ?", vbYesNo, "삭제확인") = vbYes Then
''            Kill App.Path & "\db\BackData.*"
''        Else
''            Exit Sub
''        End If
''    End If
''
''    Filename = Dir("A:\Backdata.*")
''
''    LblMsg.Caption = "자료 COPY중..!"
''
''    Ani1.Visible = True
''    Ani1.AutoPlay = True
''    Ani1.Open App.Path & "\image\filecopy.avi"
''    DoEvents
''
''    FileCopy "A:\" & Filename, App.Path & "\DB\" & Filename
''    DoEvents
''
''    Ani1.Visible = False
''    LblMsg.Caption = "자료 COPY완료..!"
''
''    Exit Sub
''
''DiskError:
''    Ani1.Visible = False
''    MsgBox "디스켓이 불량합니다..!", vbInformation, "RESTORE"
''    LblMsg.Caption = "디스켓을 넣고 COPY 버튼을 누르세요..!"
''End Sub
''
''
''Private Sub Form_Load()
''    'TitleSet "RESTORE"
''    LblMsg.Caption = ""
''    LblMsg.Caption = "디스켓을 넣고 COPY 버튼을 누르세요..!"
''End Sub
''Private Sub Form_Unload(Cancel As Integer)
''    If Not Dir(App.Path & "\BACKUP.OK") = "" Then
''        Kill App.Path & "\BACKUP.OK"
''    End If
''End Sub
''
''Private Sub CmdRestore_Click()
''
''    If Dir(App.Path & "\DB\BackData.arj") = "" Then
''        LblMsg.Caption = "저장할 자료가 없습니다..!"
''        Exit Sub
''    End If
''
''    ProgBar.Visible = True
''    ProgBar.MAX = 1000
''    ProgBar.Min = 0
''    ProgBar.Value = 0
''
''    If Dir(App.Path & "\RESTORE.BAT") = "" Then
''        MsgBox "RESTORE.BAT 파일이 없어서 이명령을 사용할 수 없습니다.", vbInformation, "오류"
''        Exit Sub
''    End If
''
''    If Not Dir(App.Path & "\Backup.OK") = "" Then
''        Kill App.Path & "\BACKUP.OK"
''    End If
''
''    MyDB.Close
''
''    If Not Dir(App.Path & "\DB\Laundry.mdb") = "" Then
''        Kill App.Path & "\DB\Laundry.mdb"
''    End If
''
''    Shell App.Path & "\Restore.BAT", 0
''
''    Do While Dir(App.Path & "\Backup.OK") = ""
''        ProgBar.Value = ProgBar.Value + 1
''
''        If ProgBar.Value = 1000 Then
''            ProgBar.Value = 0
''        End If
''
''        DoEvents
''    Loop
''
''    ProgBar.Visible = False
''
''    Set MyDB = OpenDatabase(App.Path & "\DB\Laundry.mdb")
''
''    MsgBox "자료 저장이 완료되었습니다..!", vbInformation, "RESOTRE"
''End Sub
''
''
