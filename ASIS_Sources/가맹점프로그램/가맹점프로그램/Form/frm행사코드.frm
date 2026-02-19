VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form frm행사코드 
   BorderStyle     =   1  '단일 고정
   Caption         =   "행사코드 확인"
   ClientHeight    =   1875
   ClientLeft      =   2775
   ClientTop       =   3420
   ClientWidth     =   6930
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
   ScaleHeight     =   1875
   ScaleWidth      =   6930
   Begin Threed.SSFrame SSFrame1 
      Height          =   1380
      Left            =   180
      TabIndex        =   0
      Top             =   225
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   2434
      _Version        =   262144
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "  행사 코드를 입력하여 주십시요.  "
      Begin VB.CommandButton cmdCancel 
         Caption         =   "취소"
         Height          =   525
         Left            =   5220
         TabIndex        =   3
         Top             =   480
         Width           =   1200
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "확인"
         Height          =   525
         Left            =   4020
         TabIndex        =   2
         Top             =   480
         Width           =   1200
      End
      Begin VB.TextBox txtPassWord 
         Height          =   480
         IMEMode         =   8  '영문
         Left            =   225
         TabIndex        =   1
         Top             =   480
         Width           =   3540
      End
   End
End
Attribute VB_Name = "frm행사코드"
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
        chkEventSale = True
        Unload Me
        Exit Sub
    End If
    
    strPass = IsEventPassWord(txtPassWord.Text) ' 비밀번호 확인
    
    If strPass = "-1" Or strPass = "-3" Then
        chkEventSale = False
        
        txtPassWord.SelStart = 0
        txtPassWord.SelLength = Len(txtPassWord.Text)
        
        If strPass = "-3" Then
            MsgBox "입력한 내용이 정확하지 않습니다.", vbInformation, "입력오류"
        End If
        
        txtPassWord.Text = ""
        txtPassWord.SetFocus
        Exit Sub
    Else
        If Not IsEventPassREGSave(txtPassWord.Text) Then
            chkEventSale = False
            
            MsgBox "입력한 내용이 레지스터리에 저장되지 않았습니다.", vbInformation, "저장오류"
            
            Exit Sub
        Else
            chkEventSale = True
            Unload Me
        End If
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
        cmdOK_Click
    End If

End Sub

