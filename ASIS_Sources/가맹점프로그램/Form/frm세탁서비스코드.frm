VERSION 5.00
Begin VB.Form frm세탁서비스코드 
   BorderStyle     =   1  '단일 고정
   Caption         =   "세탁 서비스 확인"
   ClientHeight    =   1785
   ClientLeft      =   720
   ClientTop       =   5400
   ClientWidth     =   5595
   ControlBox      =   0   'False
   LinkTopic       =   "Form33"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   5595
   Begin VB.TextBox txtPassWord 
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   900
      TabIndex        =   3
      Top             =   660
      Width           =   2550
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "취소"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   3810
      TabIndex        =   2
      Top             =   690
      Width           =   1320
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "확인"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   3810
      TabIndex        =   1
      Top             =   105
      Width           =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "코드"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   195
      TabIndex        =   4
      Top             =   690
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "세탁 서비스 코드 입력"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   3330
   End
End
Attribute VB_Name = "frm세탁서비스코드"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    chkServicePassWord = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
'+------------------------------------------------------
'+ 2003/02/11 수정
'+
'+루틴설명      - 비밀번호확인
'+  1. 암호를 확인하여 암호 규칙에 맞으면 화면을 종료한다.
'+  2. 레지스터리에 저장한다.
'+
'+------------------------------------------------------
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
        chkPassWord = False
        txtPassWord.SelStart = 0: txtPassWord.SelLength = Len(txtPassWord.Text)
        If strPass = "-3" Then MsgBox "입력한 내용이 정확하지 않습니다.", vbInformation, "입력오류"
        txtPassWord.Text = ""
        txtPassWord.SetFocus
        Exit Sub
    Else
        chkServicePassWord = frm접수.txtCode.Text
        Unload Me
        Exit Sub
    End If

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
        Call cmdOK_Click
    End If

End Sub
