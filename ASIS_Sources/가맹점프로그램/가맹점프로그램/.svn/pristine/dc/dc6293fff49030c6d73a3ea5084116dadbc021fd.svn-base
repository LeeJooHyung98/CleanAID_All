VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmInputEventCode 
   BorderStyle     =   1  '단일 고정
   Caption         =   "할인코드 입력"
   ClientHeight    =   1410
   ClientLeft      =   13470
   ClientTop       =   13695
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4965
   StartUpPosition =   1  '소유자 가운데
   Begin CSTextLibCtl.sitxEdit txtCode 
      Height          =   495
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   4725
      _Version        =   262145
      _ExtentX        =   8334
      _ExtentY        =   873
      _StockProps     =   125
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   "____-____-____"
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   29
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   "####-####-####"
      Justification   =   1
      CharacterTable  =   ""
      BorderStyle     =   0
      Characters      =   12
      MaxLength       =   12
   End
   Begin XtremeSuiteControls.PushButton cmdAction 
      Height          =   660
      Left            =   3255
      TabIndex        =   0
      Top             =   660
      Width           =   1590
      _Version        =   851970
      _ExtentX        =   2805
      _ExtentY        =   1164
      _StockProps     =   79
      Caption         =   "적용"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
End
Attribute VB_Name = "frmInputEventCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Rate As String
Public EventCode As String
Private Sub cmdAction_Click()
    Rate = GetEventRate(txtCode.Text)
    If InStr(Rate, "ERROR") > 0 Then
        MsgBox (Replace(Rate, "ERROR ", ""))
    Else
        EventCode = txtCode.Text
        Unload Me
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    cmdAction_Click
    End If
End Sub
