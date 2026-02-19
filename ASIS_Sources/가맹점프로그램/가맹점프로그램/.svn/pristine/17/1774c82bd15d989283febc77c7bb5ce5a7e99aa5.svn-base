VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frmKeyboard 
   BorderStyle     =   1  '단일 고정
   Caption         =   "가상 키보드"
   ClientHeight    =   4080
   ClientLeft      =   9330
   ClientTop       =   8520
   ClientWidth     =   2745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   4080
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   7197
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   0
      PaneTree        =   "frmKeyboard.frx":0000
      Begin Threed.SSPanel SSPanel1 
         Height          =   555
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   979
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtInput 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   105
            TabIndex        =   3
            Text            =   "#"
            Top             =   75
            Width           =   2490
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   3510
         Left            =   0
         TabIndex        =   1
         Top             =   570
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   6191
         _Version        =   262144
         BackColor       =   16777215
         PictureFrames   =   1
         Picture         =   "frmKeyboard.frx":0052
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image btnNumber 
            Height          =   765
            Index           =   4
            Left            =   105
            Picture         =   "frmKeyboard.frx":46EE
            Tag             =   "N$"
            Top             =   975
            Width           =   795
         End
         Begin VB.Image btnNumber 
            Height          =   765
            Index           =   5
            Left            =   975
            Picture         =   "frmKeyboard.frx":6712
            Tag             =   "N%"
            Top             =   975
            Width           =   795
         End
         Begin VB.Image btnNumber 
            Height          =   765
            Index           =   6
            Left            =   1845
            Picture         =   "frmKeyboard.frx":8736
            Tag             =   "N^"
            Top             =   975
            Width           =   795
         End
         Begin VB.Image btnNumber 
            Height          =   765
            Index           =   7
            Left            =   105
            Picture         =   "frmKeyboard.frx":A75A
            Tag             =   "N&"
            Top             =   150
            Width           =   795
         End
         Begin VB.Image btnNumber 
            Height          =   765
            Index           =   8
            Left            =   975
            Picture         =   "frmKeyboard.frx":C77E
            Tag             =   "N*"
            Top             =   150
            Width           =   795
         End
         Begin VB.Image btnNumber 
            Height          =   765
            Index           =   9
            Left            =   1845
            Picture         =   "frmKeyboard.frx":E7A2
            Tag             =   "N("
            Top             =   150
            Width           =   795
         End
         Begin VB.Image btnNumber 
            Height          =   765
            Index           =   0
            Left            =   105
            Picture         =   "frmKeyboard.frx":107C6
            Tag             =   "N)"
            Top             =   2625
            Width           =   795
         End
         Begin VB.Image btnNumber 
            Height          =   765
            Index           =   1
            Left            =   105
            Picture         =   "frmKeyboard.frx":127EA
            Tag             =   "N!"
            Top             =   1800
            Width           =   795
         End
         Begin VB.Image btnNumber 
            Height          =   765
            Index           =   2
            Left            =   975
            Picture         =   "frmKeyboard.frx":1480E
            Tag             =   "N@"
            Top             =   1800
            Width           =   795
         End
         Begin VB.Image btnNumber 
            Height          =   765
            Index           =   3
            Left            =   1845
            Picture         =   "frmKeyboard.frx":16832
            Tag             =   "N#"
            Top             =   1800
            Width           =   795
         End
         Begin VB.Image btnClear 
            Height          =   765
            Left            =   975
            Picture         =   "frmKeyboard.frx":18856
            Tag             =   "N)"
            Top             =   2625
            Width           =   795
         End
         Begin VB.Image btnEnd 
            Height          =   765
            Left            =   1845
            Picture         =   "frmKeyboard.frx":1A878
            Tag             =   "N)"
            Top             =   2625
            Width           =   795
         End
      End
   End
End
Attribute VB_Name = "frmKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClear_Click()
    txtInput.Text = ""
End Sub

Private Sub btnEnd_Click()
    Dim strInput As String
    
    strInput = txtInput.Text & ""
    
    Unload Me
    
    If strInput = "" Then Exit Sub
    
    If ActiveForm = "접수" Then
        frm접수.txtTel.Text = strInput & ""
    Else
        frm출고.txtTel.Text = strInput & ""
    End If
End Sub

Private Sub btnNumber_Click(Index As Integer)
    txtInput.Text = txtInput.Text & CStr(Index)
    
    txtInput.SelStart = Len(txtInput.Text)
End Sub

Private Sub Form_Load()
    txtInput.Text = ""
End Sub

Private Sub txtInput_Change()
    If Len(txtInput.Text) >= 4 Then
        btnEnd_Click
    End If
End Sub
