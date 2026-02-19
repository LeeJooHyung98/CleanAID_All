VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_CONNECT 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "서버연결"
   ClientHeight    =   1230
   ClientLeft      =   8895
   ClientTop       =   8265
   ClientWidth     =   7065
   Icon            =   "P_CONNECT.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.PushButton cmdExit 
      Height          =   420
      Left            =   5970
      TabIndex        =   1
      Top             =   720
      Width           =   1020
      _Version        =   851970
      _ExtentX        =   1799
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   " 확인"
      ForeColor       =   -2147483640
      BackColor       =   -2147483636
      Appearance      =   6
      Picture         =   "P_CONNECT.frx":030A
   End
   Begin VB.Image Image 
      Height          =   765
      Left            =   135
      Picture         =   "P_CONNECT.frx":0D1C
      Top             =   15
      Width           =   765
   End
   Begin VB.Label pnl_Stats 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "서버와 연결중 입니다. 잠시만 기다려 주십시요."
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1050
      TabIndex        =   0
      Top             =   270
      Width           =   5790
   End
End
Attribute VB_Name = "P_CONNECT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim P_CONNECT_MODE As Boolean

Private Sub cmdExit_Click()
    End
End Sub

Private Sub Form_Activate()
    cmdExit.Visible = False
    pnl_Stats.Caption = "서버와 연결중 입니다. 잠시만 기다려 주십시요."
    
    'DoEvents
    P_CONNECT_MODE = DBOpen
    
    If P_CONNECT_MODE = False Then
        Beep
        pnl_Stats.Caption = "서버와 연결하지 못하여 프로그램을 종료 합니다."
        cmdExit.Visible = True
        Exit Sub
    End If
    
''    ' DB 업그레이드
''    SQL_DB_Update
    
    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If P_CONNECT_MODE = False Then
        End
    Else
        Set P_CONNECT = Nothing
    End If
End Sub
