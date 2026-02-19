VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "잠시만..."
   ClientHeight    =   990
   ClientLeft      =   11040
   ClientTop       =   9345
   ClientWidth     =   6735
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.ProgressBar ProgressBar1 
      Height          =   210
      Left            =   705
      TabIndex        =   0
      Top             =   615
      Width           =   5910
      _Version        =   851970
      _ExtentX        =   10425
      _ExtentY        =   370
      _StockProps     =   93
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Image Image 
      Height          =   480
      Left            =   120
      Picture         =   "frmSplash.frx":08CA
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblCount 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6465
      TabIndex        =   2
      Top             =   180
      Width           =   120
   End
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "#"
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
      Left            =   750
      TabIndex        =   1
      Top             =   240
      Width           =   105
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Dim WithEvents CMyDB     As CDB_Update
Attribute CMyDB.VB_VarHelpID = -1

Private Sub Form_Load()
    On Error GoTo ERR_RTN
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2

    Call Set_ProgramVersion              ' 프로그램 버전을 설정한다.
    
    lblMsg.Caption = "잠시만 기다려 주세요..." & "1"

    ProgressBar1.Visible = False
    
    
    Me.Caption = Me.Caption & "  " & "Ver " & Program_Version & "  (최종수정일자 : " & Program_LastEdit & ")"
    lblMsg.Caption = "잠시만 기다려 주세요..." & "2"
    
    lblMsg.Caption = "잠시만 기다려 주세요..." & "3"
    
'    ' 신규 SQL 업데이트 확인
'    ' DB 업데이트를 확인하여 실행한다.
    lblMsg.Caption = "잠시만 기다려 주세요..." & "4"
    If DB_Connect = False Then
        lblMsg.Caption = "잠시만 기다려 주세요..." & "5"
        frm서버.Show 1
    
        lblMsg.Caption = "잠시만 기다려 주세요..." & "6"

        End
    End If
    
    lblMsg.Caption = "잠시만 기다려 주세요..." & "7"
    Set CMyDB = New CDB_Update
    
    lblMsg.Caption = "잠시만 기다려 주세요..." & "8"
    CMyDB.SetDef ADOCon
    
    lblMsg.Caption = "잠시만 기다려 주세요..." & "9"
    Call CMyDB.Update_SQL_Check
    
    lblMsg.Caption = "서버 자료 다운로드중 입니다. 잠시만 기다려 주세요..." & "10"
    Exit Sub
    
ERR_RTN:
    Call Error_Msg("", Err.Source, Err.Number, "frmSplash Form_Load " & Err.description & vbNewLine)

    Screen.MousePointer = 0
    
End Sub


Private Sub CMyDB_Error(ByVal Number As Long, description As String)
    MsgBox "DB가 업그레이드 되지 못하였습니다. 프로그램이 정상적으로 동작하지 않을 수 있습니다. " & vbLf & vbLf & description, vbCritical, "경고"
End Sub

