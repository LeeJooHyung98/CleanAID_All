VERSION 5.00
Begin VB.Form P_SCREEN 
   Caption         =   "출력물 미리보기"
   ClientHeight    =   7860
   ClientLeft      =   1455
   ClientTop       =   2340
   ClientWidth     =   12135
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7860
   ScaleWidth      =   12135
   WindowState     =   2  '최대화
End
Attribute VB_Name = "P_SCREEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Initialize()
'    Dim sc As Screen
'
'    With P_00000.crPrint
'
'        .WindowParentHandle = Me.hwnd
'        .WindowState = crptNormal
'        .Destination = crptToWindow
'        .WindowControlBox = False
'        .WindowControls = True
'        .WindowLeft = 1
'        .WindowTop = 1
'        .WindowWidth = Me.Width * 0.0665
'        .WindowHeight = Me.Height * 0.0665
'        .Action = 1
'    End With
End Sub

Private Sub Form_Load()
'    cmdBtn(0).Enabled = False
'    cmdBtn(1).Enabled = False
'    cmdBtn(2).Enabled = False
'    cmdBtn(3).Enabled = False
'    cmdBtn(4).Enabled = False
'    cmdBtn(5).Enabled = True
'    cmdBtn(6).Enabled = False
'
'    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
End Sub

Public Sub DataPrint()
'    With P_00000.crPrint
'        .Destination = crptToPrinter
'        .Action = 1
'    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Set P_SCREEN = Nothing
End Sub
