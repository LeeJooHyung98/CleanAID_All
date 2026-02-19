VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm서버 
   BackColor       =   &H00D9E5E9&
   BorderStyle     =   1  '단일 고정
   Caption         =   "DB 서버 설정"
   ClientHeight    =   5445
   ClientLeft      =   7785
   ClientTop       =   5325
   ClientWidth     =   6540
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm서버.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   6540
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   5445
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   9604
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm서버.frx":08CA
      Begin Threed.SSPanel SSPanel 
         Height          =   4410
         Index           =   1
         Left            =   15
         TabIndex        =   5
         Top             =   435
         Width           =   6510
         _ExtentX        =   11483
         _ExtentY        =   7779
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSFrame SSFrame 
            Height          =   1995
            Index           =   0
            Left            =   105
            TabIndex        =   6
            Top             =   135
            Width           =   6315
            _ExtentX        =   11139
            _ExtentY        =   3519
            _Version        =   262144
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
            Caption         =   "로컬 DB"
            Begin VB.TextBox txtData 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               IMEMode         =   3  '사용 못함
               Index           =   3
               Left            =   1395
               PasswordChar    =   "*"
               TabIndex        =   10
               Top             =   1515
               Width           =   4830
            End
            Begin VB.TextBox txtData 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               IMEMode         =   10  '한글 
               Index           =   2
               Left            =   1395
               TabIndex        =   9
               Top             =   1095
               Width           =   4830
            End
            Begin VB.TextBox txtData 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               IMEMode         =   10  '한글 
               Index           =   1
               Left            =   1395
               TabIndex        =   8
               Top             =   675
               Width           =   4860
            End
            Begin VB.TextBox txtData 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               IMEMode         =   10  '한글 
               Index           =   0
               Left            =   1395
               TabIndex        =   7
               Top             =   255
               Width           =   4860
            End
            Begin VB.Label Label 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "비밀번호:"
               Height          =   195
               Index           =   3
               Left            =   75
               TabIndex        =   14
               Top             =   1605
               Width           =   1275
            End
            Begin VB.Label Label 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "아이디:"
               Height          =   195
               Index           =   2
               Left            =   75
               TabIndex        =   13
               Top             =   1185
               Width           =   1275
            End
            Begin VB.Label Label 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "데이터베이스:"
               Height          =   195
               Index           =   1
               Left            =   75
               TabIndex        =   12
               Top             =   765
               Width           =   1275
            End
            Begin VB.Label Label 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "서버이름:"
               Height          =   195
               Index           =   0
               Left            =   75
               TabIndex        =   11
               Top             =   345
               Width           =   1275
            End
         End
         Begin Threed.SSFrame SSFrame 
            Height          =   1995
            Index           =   1
            Left            =   105
            TabIndex        =   15
            Top             =   2295
            Width           =   6315
            _ExtentX        =   11139
            _ExtentY        =   3519
            _Version        =   262144
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
            Caption         =   "서버 DB"
            Begin VB.TextBox txtData 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               IMEMode         =   10  '한글 
               Index           =   4
               Left            =   1395
               TabIndex        =   19
               Top             =   255
               Width           =   3180
            End
            Begin VB.TextBox txtData 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               IMEMode         =   10  '한글 
               Index           =   5
               Left            =   1395
               TabIndex        =   18
               Top             =   675
               Width           =   3180
            End
            Begin VB.TextBox txtData 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               IMEMode         =   10  '한글 
               Index           =   6
               Left            =   1395
               TabIndex        =   17
               Top             =   1095
               Width           =   4830
            End
            Begin VB.TextBox txtData 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               IMEMode         =   3  '사용 못함
               Index           =   7
               Left            =   1395
               PasswordChar    =   "*"
               TabIndex        =   16
               Top             =   1515
               Width           =   4830
            End
            Begin XtremeSuiteControls.PushButton cmdList 
               Height          =   780
               Left            =   4650
               TabIndex        =   24
               Top             =   270
               Width           =   1560
               _Version        =   851970
               _ExtentX        =   2752
               _ExtentY        =   1376
               _StockProps     =   79
               Caption         =   " 연결 확인"
               BackColor       =   -2147483633
               Appearance      =   6
               Picture         =   "frm서버.frx":093C
            End
            Begin VB.Label Label 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "서버이름:"
               Height          =   195
               Index           =   7
               Left            =   75
               TabIndex        =   23
               Top             =   345
               Width           =   1275
            End
            Begin VB.Label Label 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "데이터베이스:"
               Height          =   195
               Index           =   6
               Left            =   75
               TabIndex        =   22
               Top             =   765
               Width           =   1275
            End
            Begin VB.Label Label 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "아이디:"
               Height          =   195
               Index           =   5
               Left            =   75
               TabIndex        =   21
               Top             =   1185
               Width           =   1275
            End
            Begin VB.Label Label 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "비밀번호:"
               Height          =   195
               Index           =   4
               Left            =   75
               TabIndex        =   20
               Top             =   1605
               Width           =   1275
            End
         End
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   570
         Index           =   0
         Left            =   15
         TabIndex        =   2
         Top             =   4860
         Width           =   6510
         _ExtentX        =   11483
         _ExtentY        =   1005
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   5190
            TabIndex        =   3
            Top             =   60
            Width           =   1275
            _Version        =   851970
            _ExtentX        =   2249
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            BackColor       =   -2147483633
            Appearance      =   6
            Picture         =   "frm서버.frx":1036
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   45
            TabIndex        =   4
            Top             =   60
            Width           =   1260
            _Version        =   851970
            _ExtentX        =   2222
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 저장(&S)"
            BackColor       =   -2147483633
            Appearance      =   6
            Picture         =   "frm서버.frx":1A48
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   405
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   6510
         _ExtentX        =   11483
         _ExtentY        =   714
         _Version        =   262144
         Font3D          =   3
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "   DB 서버 설정"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm서버.frx":245A
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image1 
            Height          =   240
            Left            =   60
            Picture         =   "frm서버.frx":28BC
            Top             =   60
            Width           =   240
         End
      End
   End
End
Attribute VB_Name = "frm서버"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn

    Select Case Index
        Case 0
            Call SetIniStr("DB", "SERVER", Set_Encrypt(txtData(0).Text, ""), iniFile)
            Call SetIniStr("DB", "DATABASE", Set_Encrypt(txtData(1).Text, ""), iniFile)
            Call SetIniStr("DB", "ID", Set_Encrypt(txtData(2).Text, ""), iniFile)
            Call SetIniStr("DB", "PWD", Set_Encrypt(txtData(3).Text, ""), iniFile)
            
            
            Call SetIniStr("SERVER", "SERVER", Set_Encrypt(txtData(4).Text, ""), iniFile)
            Call SetIniStr("SERVER", "DATABASE", Set_Encrypt(txtData(5).Text, ""), iniFile)
            Call SetIniStr("SERVER", "ID", Set_Encrypt(txtData(6).Text, ""), iniFile)
            Call SetIniStr("SERVER", "PWD", Set_Encrypt(txtData(7).Text, ""), iniFile)
            
            Unload Me
            
        Case 1:
            Unload Me
    End Select

    Exit Sub

ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub cmdList_Click()
    Dim THostCon           As ADODB.Connection
    Dim sServer   As String
    Dim sDatabase As String
    Dim sID       As String
    Dim sPWD      As String
    
    Dim Server_Connect As String
    
    On Error GoTo ErrRtn
 
        
   sServer = txtData(4).Text
   sDatabase = txtData(5).Text
   sID = txtData(6).Text
   sPWD = txtData(7).Text
        
    Set THostCon = Nothing
    Set THostCon = New ADODB.Connection
    
    With THostCon
        
        .ConnectionString = "Provider=SQLOLEDB;Persist Security Info=False;User ID=" & sID & ";Password=" & sPWD & ";Initial Catalog=" & sDatabase & ";Data Source=" & sServer
        .ConnectionTimeout = 10
        .CommandTimeout = 30
        .Open
    End With
    
    THostCon.Close
    Set THostCon = Nothing
    
    MsgBox "정상적으로 연결 되었습니다.", vbInformation, "연결확인"
    
    Exit Sub

ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)


End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2

    txtData(0).Text = Get_Decrypt(GetIniStr("DB", "SERVER", "", iniFile), "")
    txtData(1).Text = Get_Decrypt(GetIniStr("DB", "DATABASE", "", iniFile), "")
    txtData(2).Text = Get_Decrypt(GetIniStr("DB", "ID", "", iniFile), "")
    txtData(3).Text = Get_Decrypt(GetIniStr("DB", "PWD", "", iniFile), "")
    
    
    txtData(4).Text = Get_Decrypt(GetIniStr("SERVER", "SERVER", "", iniFile), "")
    txtData(5).Text = Get_Decrypt(GetIniStr("SERVER", "DATABASE", "", iniFile), "")
    txtData(6).Text = Get_Decrypt(GetIniStr("SERVER", "ID", "", iniFile), "")
    txtData(7).Text = Get_Decrypt(GetIniStr("SERVER", "PWD", "", iniFile), "")
    
    Exit Sub

ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub
