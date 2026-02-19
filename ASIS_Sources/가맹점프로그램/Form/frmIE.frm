VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmIE 
   BorderStyle     =   1  '단일 고정
   Caption         =   "원격 A/S"
   ClientHeight    =   8415
   ClientLeft      =   5280
   ClientTop       =   4095
   ClientWidth     =   11250
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIE.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   11250
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   14843
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frmIE.frx":08CA
      Begin Threed.SSPanel SSPanel 
         Height          =   600
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   11220
         _ExtentX        =   19791
         _ExtentY        =   1058
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   480
            Index           =   5
            Left            =   9900
            TabIndex        =   3
            Top             =   60
            Width           =   1260
            _Version        =   851970
            _ExtentX        =   2222
            _ExtentY        =   847
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frmIE.frx":091C
         End
         Begin VB.Image Image 
            Height          =   480
            Left            =   75
            Picture         =   "frmIE.frx":132E
            Top             =   60
            Width           =   480
         End
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   7770
         Left            =   15
         TabIndex        =   1
         Top             =   630
         Width           =   11220
         ExtentX         =   19791
         ExtentY         =   13705
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
End
Attribute VB_Name = "frmIE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
    
        Case 5: Unload Me
    End Select
End Sub
