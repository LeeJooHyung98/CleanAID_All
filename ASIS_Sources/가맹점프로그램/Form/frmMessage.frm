VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frmMessage 
   BorderStyle     =   1  '단일 고정
   Caption         =   "메시지"
   ClientHeight    =   7560
   ClientLeft      =   8640
   ClientTop       =   4260
   ClientWidth     =   11520
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   11520
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   7560
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   13335
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frmMessage.frx":0000
      Begin Threed.SSPanel SSPanel 
         Height          =   570
         Index           =   0
         Left            =   15
         TabIndex        =   1
         Top             =   6975
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   1005
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdAction 
            Height          =   480
            Left            =   10290
            TabIndex        =   3
            Top             =   45
            Width           =   1140
            _Version        =   851970
            _ExtentX        =   2011
            _ExtentY        =   847
            _StockProps     =   79
            Caption         =   "확인(&O)"
            Appearance      =   6
         End
         Begin VB.Label pnlMailDay 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "일자"
            Height          =   180
            Left            =   75
            TabIndex        =   5
            Top             =   75
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label pnlMailNo 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "번호"
            Height          =   180
            Left            =   690
            TabIndex        =   4
            Top             =   75
            Visible         =   0   'False
            Width           =   360
         End
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   6075
         Left            =   15
         TabIndex        =   2
         Top             =   885
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   10716
         _Version        =   393217
         BackColor       =   12648447
         Enabled         =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmMessage.frx":0092
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   465
         Index           =   1
         Left            =   15
         TabIndex        =   6
         Top             =   15
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   820
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Label pnlDay 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   1035
            TabIndex        =   8
            Top             =   150
            Width           =   105
         End
         Begin VB.Label Label6 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "조회기간:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   7
            Top             =   150
            Width           =   885
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   375
         Left            =   15
         TabIndex        =   9
         Top             =   495
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   661
         _Version        =   262144
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "메시지 확인 방법 : [조회 -> 메시지내용조회] 에서 추후 확인 가능 합니다."
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAction_Click()
    On Error Resume Next
    
    Query = "UPDATE TB_공지사항 SET "
    Query = Query & "  수신일자     = (CASE WHEN 수신일자 IS NULL THEN '' ELSE 수신일자 END) + '" & Format(Now, "YYYY-MM-DD hh:mm:ss") & "|'"
    Query = Query & ", 수신여부     = 'Y'"
    Query = Query & ", 본사전송여부 = 'N'"
    Query = Query & " WHERE 작성일자 = '" & pnlMailDay.Caption & "'"
    Query = Query & "   AND 문서번호 = " & pnlMailNo.Caption
    ADOCon.Execute Query
    
    Unload Me
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub Form_Load()
    Me.Top = ((frmMain.Height - Me.Height) / 2) + frmMain.Top
    Me.Left = ((frmMain.Width - Me.Width) / 2) + frmMain.Left
End Sub
