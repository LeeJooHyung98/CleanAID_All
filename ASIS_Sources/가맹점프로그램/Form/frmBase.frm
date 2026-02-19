VERSION 5.00
Begin VB.Form frmBase 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   8175
   ClientLeft      =   1185
   ClientTop       =   5115
   ClientWidth     =   11880
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   WindowState     =   2  '최대화
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   11430
      Top             =   7740
   End
   Begin VB.Image Image1 
      Height          =   8085
      Left            =   -330
      Picture         =   "frmBase.frx":0000
      Top             =   30
      Width           =   12750
   End
End
Attribute VB_Name = "frmBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call KeyChk(KeyCode)
    
    Select Case KeyCode
        Case vbKeyF12: frm환경설정.Show     ' 대리점 정보
        Case vbKeyF1: frmInSoftNet.Show ' InSoftNet 정보
            
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyChk(KeyAscii)
End Sub

Private Sub Form_Load()
'    Dim strLogo As String
'
'    'Me.Height = 15105
'    'Me.Width = 15225
'
'    frmMain.StatusBar1.Panels(1) = M_CompnyMasterName & "   " & "Ver " & Program_Version
'    frmMain.StatusBar1.Panels(1).ToolTipText = M_CompnyMasterName & "   " & "Ver " & Program_Version & "   최종 수정일:" & Program_LastEdit
'
'    '-------------------------------------------------------
'    '
'    '-------------------------------------------------------
'    Query = "SELECT * FROM TB_기본정보"
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    If ADORs.EOF Then
'        '
'    Else
'        frmMain.StatusBar1.Panels(2) = ADORs!가맹점명
'        frmMain.StatusBar1.Panels(5) = Trim(ADORs!매장전화번호) & ""
'    End If
'    ADORs.Close
'    Set ADORs = Nothing
'
'    strLogo = App.Path & "\image\MainLogo.jpg"
'
'    If Dir(strLogo, vbDirectory) <> "" Then
'        Image1.Picture = LoadPicture(App.Path & "\image\MainLogo.jpg")
'    End If
End Sub

