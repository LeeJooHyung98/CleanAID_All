VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm환불사유 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "환불사유"
   ClientHeight    =   1620
   ClientLeft      =   7860
   ClientTop       =   6375
   ClientWidth     =   5610
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm환불사유.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   1620
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5610
      _ExtentX        =   9895
      _ExtentY        =   2858
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm환불사유.frx":0A02
      Begin Threed.SSPanel SSPanel 
         Height          =   600
         Left            =   15
         TabIndex        =   5
         Top             =   420
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   1058
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox cboMemo 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   75
            TabIndex        =   0
            Top             =   135
            Width           =   5475
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   390
         Left            =   15
         TabIndex        =   3
         Top             =   15
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   688
         _Version        =   262144
         Font3D          =   1
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "   환불사유"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm환불사유.frx":0A74
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   240
            Left            =   60
            Picture         =   "frm환불사유.frx":0ED6
            Top             =   75
            Width           =   240
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   570
         Left            =   15
         TabIndex        =   4
         Top             =   1035
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   1005
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdClose 
            Height          =   480
            Left            =   4275
            TabIndex        =   1
            Top             =   45
            Width           =   1260
            _Version        =   851970
            _ExtentX        =   2222
            _ExtentY        =   847
            _StockProps     =   79
            Caption         =   " 확인(&O)"
            Appearance      =   6
            Picture         =   "frm환불사유.frx":1460
         End
         Begin XtremeSuiteControls.PushButton cmdCancel 
            Height          =   480
            Left            =   45
            TabIndex        =   6
            Top             =   45
            Width           =   1260
            _Version        =   851970
            _ExtentX        =   2222
            _ExtentY        =   847
            _StockProps     =   79
            Caption         =   " 취소(&C)"
            Appearance      =   6
            Picture         =   "frm환불사유.frx":1E72
         End
      End
   End
End
Attribute VB_Name = "frm환불사유"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboMemo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cboMemo.Text = "" Then
            MsgBox "환불사유를 입력하세요.", vbInformation, "확인"
            
            Exit Sub
        End If
        
        환불사유 = cboMemo.Text & ""
    
        Unload Me
    End If
End Sub

Private Sub cmdCancel_Click()
    Rtn = 0
    
    Unload Me
End Sub

Private Sub cmdClose_Click()
    If cboMemo.Text = "" Then
        MsgBox "환불 사유를 입력하세요.", vbInformation, "확인"
        
        cboMemo.SetFocus
        
        Exit Sub
    End If
    
    환불사유 = cboMemo.Text & ""
    
    Rtn = 1
    
    Unload Me
End Sub

Private Sub Form_Load()
    Query = "SELECT * FROM TB_환불사유"
    Query = Query & " ORDER BY 순서 ASC"
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    With cboMemo
        .Clear
        
        Do Until SUBRs.EOF
            .AddItem SUBRs!환불사유 & ""
            
            SUBRs.MoveNext
        Loop
        SUBRs.Close
        Set SUBRs = Nothing
        
        .ListIndex = -1
    End With
End Sub
