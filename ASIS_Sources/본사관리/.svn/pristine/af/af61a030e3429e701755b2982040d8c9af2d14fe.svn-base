VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form P_08004 
   Caption         =   "자료 마감 / 복구"
   ClientHeight    =   10020
   ClientLeft      =   1170
   ClientTop       =   1530
   ClientWidth     =   13665
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_08004.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10020
   ScaleWidth      =   13665
   WindowState     =   2  '최대화
   Begin Threed.SSPanel panMain 
      Align           =   1  '위 맞춤
      Height          =   9075
      Left            =   0
      TabIndex        =   0
      Top             =   435
      Width           =   13665
      _ExtentX        =   24104
      _ExtentY        =   16007
      _Version        =   262144
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel panInput 
      Align           =   1  '위 맞춤
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13665
      _ExtentX        =   24104
      _ExtentY        =   767
      _Version        =   262144
      RoundedCorners  =   0   'False
      Begin MSComCtl2.DTPicker dtInput 
         Height          =   315
         Index           =   1
         Left            =   4860
         TabIndex        =   2
         Top             =   60
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   56164352
         CurrentDate     =   36686
      End
      Begin MSComCtl2.DTPicker dtInput 
         Height          =   315
         Index           =   0
         Left            =   1680
         TabIndex        =   3
         Top             =   60
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   56164352
         CurrentDate     =   36686
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   2
         Left            =   60
         TabIndex        =   4
         Top             =   60
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "작 업 일 자"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   8
         Left            =   9660
         TabIndex        =   5
         Top             =   60
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   262144
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         Begin Threed.SSOption optSelect 
            Height          =   255
            Index           =   0
            Left            =   300
            TabIndex        =   6
            Top             =   30
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   262144
            Caption         =   "마  감"
            Value           =   -1
         End
         Begin Threed.SSOption optSelect 
            Height          =   255
            Index           =   1
            Left            =   1740
            TabIndex        =   7
            Top             =   30
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   262144
            Caption         =   "복  구"
         End
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   9
         Left            =   8040
         TabIndex        =   8
         Top             =   60
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "작 업 구 분"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
   End
End
Attribute VB_Name = "P_08004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Err_Num As Long
Dim Err_Dec As String

Dim sValue() As String

Private Sub Form_Activate()
'    cmdBtn(2).Enabled = True
'
'    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    If P_08004_Flag = False Then
        dtInput(0).Value = Date
        dtInput(1).Value = Date
        
        P_08004_Flag = True
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Public Sub DataSave()
    ReDim sValue(2)
    
    sValue(0) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(1) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    If optSelect(0).Value = True Then
        sValue(2) = "0"
    ElseIf optSelect(1).Value = True Then
        sValue(2) = "1"
    End If
    
    If MsgBox("작업을 시작하시 겠습니까?,", vbQuestion + vbYesNo, "마감 / 복구작업") = vbYes Then
        Call ExecPro("SP08004", sValue(), Err_Num, Err_Dec)
    Else
        Exit Sub
    End If
    
    If Err_Num = 0 Then
        MsgBox "작업이 정상적으로 완료되었습니다.", vbInformation
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_08004_Flag = False
End Sub
