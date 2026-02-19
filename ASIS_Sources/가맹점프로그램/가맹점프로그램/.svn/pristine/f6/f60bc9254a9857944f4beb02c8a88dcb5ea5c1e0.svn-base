VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm서비스코드 
   BorderStyle     =   1  '단일 고정
   Caption         =   "서비스코드 확인"
   ClientHeight    =   1845
   ClientLeft      =   3105
   ClientTop       =   10305
   ClientWidth     =   3900
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   11.25
      Charset         =   129
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form32"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   3900
   StartUpPosition =   1  '소유자 가운데
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   1845
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   3254
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   0
      PaneTree        =   "frm서비스코드.frx":0000
      Begin Threed.SSPanel SSPanel 
         Height          =   630
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   1215
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   1111
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdOK 
            Height          =   480
            Left            =   2550
            TabIndex        =   4
            Top             =   60
            Width           =   1260
            _Version        =   851970
            _ExtentX        =   2222
            _ExtentY        =   847
            _StockProps     =   79
            Caption         =   " 확인(&O)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm서비스코드.frx":0052
         End
         Begin XtremeSuiteControls.PushButton cmdCancel 
            Height          =   480
            Left            =   75
            TabIndex        =   5
            Top             =   60
            Width           =   1260
            _Version        =   851970
            _ExtentX        =   2222
            _ExtentY        =   847
            _StockProps     =   79
            Caption         =   " 취소(&X)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm서비스코드.frx":0A64
         End
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   1200
         Index           =   1
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   2117
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSFrame SSFrame1 
            Height          =   915
            Left            =   120
            TabIndex        =   6
            Top             =   150
            Width           =   3690
            _ExtentX        =   6509
            _ExtentY        =   1614
            _Version        =   262144
            Windowless      =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "서비스 코드를 입력하여 주십시요"
            Begin VB.TextBox txtPassWord 
               Height          =   390
               IMEMode         =   8  '영문
               Left            =   165
               TabIndex        =   0
               Top             =   375
               Width           =   3375
            End
         End
      End
   End
End
Attribute VB_Name = "frm서비스코드"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'+------------------------------------------------------
'+ 2003/08/29 수정
'+
'+루틴설명      - 비밀번호확인
'+  1. 암호를 확인하여 암호 규칙에 맞으면 화면을 종료한다.
'+  2. 레지스터리에 저장한다.
'+
'+------------------------------------------------------
Private Sub cmdOK_Click()
    Dim strPass As String
    
    ' 입력 확인
    If Len(txtPassWord.Text) <= 0 Then
        Exit Sub
    End If
    
'   기본 디폴드 암호.. ( 프로그램 셋팅/설치를 위한 암호 )
    If UCase(txtPassWord.Text) = "DUDTJSGH" Then
        chkServicePassWord = frm접수.txtCode.Text
        Unload Me
        Exit Sub
    End If
    
    ' 비밀번호 확인
    strPass = IsServicePassWord(txtPassWord.Text, frm접수.txtCode.Text)
    
    If strPass = "-1" Or strPass = "-3" Then
        txtPassWord.SelStart = 0
        txtPassWord.SelLength = Len(txtPassWord.Text)
        
        If strPass = "-3" Then
            MsgBox "입력한 내용이 정확하지 않습니다.", vbInformation, "입력오류"
        End If
        
        txtPassWord.Text = ""
        txtPassWord.SetFocus
        Exit Sub
    Else
        chkServicePassWord = frm접수.txtCode.Text
        Unload Me
        Exit Sub
    End If
            
End Sub

Private Sub Form_Activate()
    SSFrame1.Caption = "[" & frm접수.txtCode.Text & "]의 서비스 코드를 입력하여 주십시요."
End Sub

Private Sub txtPassWord_KeyDown(KeyCode As Integer, Shift As Integer)

    If (KeyCode < 48 Or KeyCode > 57) And (KeyCode < 96 Or KeyCode > 105) Then
        If txtPassWord.PasswordChar <> "*" Then
            txtPassWord.PasswordChar = "*"
        End If
    Else
        txtPassWord.PasswordChar = ""
    End If
    
    If KeyCode = vbKeyReturn Then
        cmdOK_Click
    End If

End Sub

