VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form P_PDA 
   Caption         =   "PDA - 핸드터미널 전송"
   ClientHeight    =   1485
   ClientLeft      =   5100
   ClientTop       =   3870
   ClientWidth     =   9840
   ControlBox      =   0   'False
   Icon            =   "P_PDA.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1485
   ScaleWidth      =   9840
   StartUpPosition =   1  '소유자 가운데
   Begin Threed.SSPanel panMessage 
      Height          =   1485
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   2619
      _Version        =   262144
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "PDA 수신 작업 중입니다........."
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel panMain 
      Align           =   1  '위 맞춤
      Height          =   1515
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   2672
      _Version        =   262144
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CommandButton cmdSubBtn 
         Caption         =   "복사"
         Height          =   705
         Index           =   0
         Left            =   7770
         TabIndex        =   1
         Top             =   60
         Width           =   960
      End
      Begin VB.CommandButton cmdSubBtn 
         Caption         =   "종료"
         Height          =   705
         Index           =   1
         Left            =   8760
         TabIndex        =   2
         Top             =   60
         Width           =   885
      End
      Begin VB.Timer HanTimer 
         Left            =   7260
         Top             =   30
      End
      Begin VB.TextBox txtInput 
         Height          =   315
         Index           =   1
         Left            =   1650
         TabIndex        =   5
         Top             =   390
         Width           =   6075
      End
      Begin VB.TextBox txtInput 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   3345
      End
      Begin VB.TextBox txtInput 
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1050
         Width           =   3345
      End
      Begin Threed.SSOption optInput 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   6
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         _Version        =   262144
         Caption         =   "입  고"
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   0
         Left            =   30
         TabIndex        =   7
         Top             =   390
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "파일 경로"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   4
         Left            =   30
         TabIndex        =   8
         Top             =   60
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "입출고구분"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSOption optInput 
         Height          =   315
         Index           =   1
         Left            =   3360
         TabIndex        =   9
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "출  고"
         Value           =   -1
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   5
         Left            =   30
         TabIndex        =   10
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "파  일  명"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   6
         Left            =   30
         TabIndex        =   11
         Top             =   1050
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "상       태"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
   End
End
Attribute VB_Name = "P_PDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdSubBtn_Click(Index As Integer)

    Select Case Index
        Case 0  '복사
            
            If txtInput(2).Text = "" Then
                Exit Sub
            End If
            Call PdaIniWrite

            Me.MousePointer = 11
            Call Command_Shell
            panMessage.Visible = True
            HanTimer.Interval = 100
            HanTimer.Enabled = True

        Case 1  '종료
            Unload Me
    End Select

End Sub




Private Sub Form_Activate()

    panMessage.Visible = False
    
    Call INIWrite("COPY", "RESULT", "9:복사준비중", sCopyIniFile)
    
    txtInput(3).Text = Trim(GetIniStr("COPY", "RESULT", "", sCopyIniFile))
    
    'dtInput.Value = Now
    
    'txtInput(1).Text = Trim(GetIniStr("PDA", "Dir", "", m_iniFile))
    
    'Call PdaNoComboAdd
    'If cboPdaNo.ListCount > 0 Then cboPdaNo.ListIndex = 0
    
    'Call PdaCntComboAdd
    'If cboPdaCnt.ListCount > 0 Then cboPdaCnt.ListIndex = 0

End Sub


Private Sub HanTimer_Timer()

    txtInput(3).Text = Trim(GetIniStr("COPY", "RESULT", "", sCopyIniFile))
    DoEvents
    
    DownPathName = ""
    DownFileName = ""
    
    If Left(txtInput(3).Text, 1) <> "9" Then
        Me.MousePointer = 0
        panMessage.Visible = False
        
        Select Case Left(txtInput(3).Text, 1)
        
            Case "1"
                DownPathName = txtInput(1).Text
                DownFileName = txtInput(2).Text
                MsgBox "복사가 완료 되었습니다.", vbOKOnly
                HanTimer.Enabled = False
                HanTimer.Interval = 0
                Exit Sub
                
            Case "0"
                MsgBox txtInput(3).Text & Chr(13) & "복사가 정상적으로 처리 되지 않았습니다..", vbOKOnly
                HanTimer.Enabled = False
                HanTimer.Interval = 0
                Exit Sub
                
        End Select
        
        
        Exit Sub
        
    Else
    
    
    End If
    
End Sub


Sub PdaIniWrite()

    Dim Str As String
    'm_iniFile
    
    Call INIWrite("PDA", "Dir", txtInput(1).Text, m_iniFile)
    'Call INIWrite("PDA", "WorkDate", Format(dtInput.Value, "YYYY-MM-DD"), m_iniFile)
    
    'sCopyIniFile
    Call INIWrite("SAVE", "DIR", txtInput(1).Text, sCopyIniFile)
    Call INIWrite("SAVE", "FILE", txtInput(2).Text, sCopyIniFile)
    Call INIWrite("COPY", "RESULT", "9:복사준비중", sCopyIniFile)
    
    txtInput(3).Text = Trim(GetIniStr("COPY", "RESULT", "", sCopyIniFile))
    
End Sub

Private Sub Command_Shell()
    On Error Resume Next

        Dim RetVal

        RetVal = Shell(App.Path & "\PDA\kiSync.exe", vbNormalFocus)
End Sub


