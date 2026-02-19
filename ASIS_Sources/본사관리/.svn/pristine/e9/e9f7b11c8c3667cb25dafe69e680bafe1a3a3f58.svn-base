VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form P_PRTSCREEN 
   Caption         =   "출력물 미리보기"
   ClientHeight    =   8040
   ClientLeft      =   975
   ClientTop       =   2520
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8040
   ScaleWidth      =   11205
   Begin MSComDlg.CommonDialog cdPrt 
      Left            =   10260
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmd_Action 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   0
      Left            =   30
      Picture         =   "P_PRTSCREEN.frx":0000
      Style           =   1  '그래픽
      TabIndex        =   13
      ToolTipText     =   "다시보기"
      Top             =   30
      Width           =   795
   End
   Begin VB.CommandButton cmd_Action 
      Caption         =   "Export"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   4
      Left            =   3270
      Picture         =   "P_PRTSCREEN.frx":08CA
      Style           =   1  '그래픽
      TabIndex        =   12
      ToolTipText     =   "내보내기"
      Top             =   30
      Width           =   795
   End
   Begin VB.CommandButton cmd_Action 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   2
      Left            =   1650
      Picture         =   "P_PRTSCREEN.frx":1194
      Style           =   1  '그래픽
      TabIndex        =   11
      ToolTipText     =   "인쇄작업"
      Top             =   30
      Width           =   795
   End
   Begin VB.CommandButton cmd_Action 
      Caption         =   "First"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   5
      Left            =   4080
      Picture         =   "P_PRTSCREEN.frx":1E5E
      Style           =   1  '그래픽
      TabIndex        =   10
      ToolTipText     =   "처음페이지"
      Top             =   30
      Width           =   795
   End
   Begin VB.CommandButton cmd_Action 
      Caption         =   "Prev"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   6
      Left            =   4890
      Picture         =   "P_PRTSCREEN.frx":2B28
      Style           =   1  '그래픽
      TabIndex        =   9
      ToolTipText     =   "이전페이지"
      Top             =   30
      Width           =   795
   End
   Begin VB.CommandButton cmd_Action 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   7
      Left            =   5700
      Picture         =   "P_PRTSCREEN.frx":37F2
      Style           =   1  '그래픽
      TabIndex        =   8
      ToolTipText     =   "다음페이지"
      Top             =   30
      Width           =   795
   End
   Begin VB.CommandButton cmd_Action 
      Caption         =   "Last"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   8
      Left            =   6510
      Picture         =   "P_PRTSCREEN.frx":44BC
      Style           =   1  '그래픽
      TabIndex        =   7
      ToolTipText     =   "마지막 페이지"
      Top             =   30
      Width           =   795
   End
   Begin VB.CommandButton cmd_Action 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   1
      Left            =   840
      Picture         =   "P_PRTSCREEN.frx":5186
      Style           =   1  '그래픽
      TabIndex        =   6
      ToolTipText     =   "인쇄취소"
      Top             =   30
      Width           =   795
   End
   Begin VB.CommandButton cmd_Action 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   9
      Left            =   7320
      Picture         =   "P_PRTSCREEN.frx":5A50
      Style           =   1  '그래픽
      TabIndex        =   5
      ToolTipText     =   "종  료"
      Top             =   30
      Width           =   795
   End
   Begin VB.CommandButton cmd_Action 
      Caption         =   "Select"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      HelpContextID   =   3
      Index           =   3
      Left            =   2460
      Picture         =   "P_PRTSCREEN.frx":671A
      Style           =   1  '그래픽
      TabIndex        =   4
      ToolTipText     =   "프린터설정"
      Top             =   30
      Width           =   795
   End
   Begin VB.ComboBox cob_Zoom 
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "P_PRTSCREEN.frx":6FE4
      Left            =   8130
      List            =   "P_PRTSCREEN.frx":7006
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   120
      Width           =   1875
   End
   Begin VB.HScrollBar HScroll 
      Height          =   300
      Left            =   60
      TabIndex        =   2
      Top             =   7530
      Width           =   10125
   End
   Begin VB.VScrollBar VScroll 
      Height          =   6645
      Left            =   10110
      TabIndex        =   1
      Top             =   870
      Width           =   300
   End
   Begin VB.PictureBox PMain 
      BackColor       =   &H8000000C&
      Height          =   6675
      Left            =   57
      ScaleHeight     =   6615
      ScaleWidth      =   10005
      TabIndex        =   0
      Top             =   850
      Width           =   10065
      Begin VB.PictureBox PView 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         Height          =   5925
         Left            =   1080
         ScaleHeight     =   5925
         ScaleWidth      =   7845
         TabIndex        =   17
         Top             =   420
         Width           =   7845
      End
   End
   Begin VB.Label lbl_CurPage 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  '단일 고정
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   8130
      TabIndex        =   16
      Top             =   480
      Width           =   825
   End
   Begin VB.Label lbl_CurPage 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  '단일 고정
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   9180
      TabIndex        =   15
      Top             =   480
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9000
      TabIndex        =   14
      Top             =   510
      Width           =   150
   End
End
Attribute VB_Name = "P_PRTSCREEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const MC As Integer = 57.3  ' 트윕단위를 밀리미터로 변경할 기본 단위
Dim PrintMode   As Integer  ' 1: 미리보기, 2: 인쇄
Dim ViewPage    As Integer  '
Dim FORM_P_PRTSCREEN As Boolean

Private Sub cmd_Action_Click(Index As Integer)

    Select Case Index
    
        Case 0
        ' Refresh
            PrintMode = 1
            Call DataPrintView
            Exit Sub
        
        Case 1
            Call FixedFormReSize
            
        Case 2
            PrintMode = 2
            Call DataPrintView
            Exit Sub
            
        Case 5
            PrintMode = 1
            ViewPage = 1
            Call DataPrintView
            Exit Sub
        
        Case 6
            PrintMode = 1
            ViewPage = ViewPage - 1
            Call DataPrintView
            Exit Sub
        
        Case 7
            PrintMode = 1
            ViewPage = ViewPage + 1
            Call DataPrintView
            Exit Sub
        
        Case 8
            PrintMode = 1
            ViewPage = ViewPage + 10000
            Call DataPrintView
            Exit Sub
            
        Case 9
            Unload Me

        Case Else
        
    End Select
End Sub

Private Sub cob_Zoom_Click()
    If cob_Zoom.Text = "Page Width" Then
        iSL_Type = 1
    ElseIf cob_Zoom.Text = "Whole Page" Then
        iSL_Type = 1
    Else
        iSL_Type = (Val(cob_Zoom.Text) / 100)
    End If

End Sub

Private Sub Form_Activate()
    If FORM_P_PRTSCREEN Then Exit Sub
    FORM_P_PRTSCREEN = True
    PrintMode = 1
    Call DataPrintView
End Sub

Private Sub Form_Initialize()
    Call FixedFormReSize
End Sub

Private Sub Form_Load()

'    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
    cob_Zoom.ListIndex = 4
End Sub

Public Sub DataPrintView()
    Dim CPrt    As New CCAIDPrinter
    
    Select Case PrtParam.Param(0)
    
        ' 물품 명세서
        Case "P_03001"
            HScroll.Max = 100
            VScroll.Max = 100
            With CPrt
                If ViewPage = 0 Then ViewPage = 1
                .PRT_03001_01_WIN IIf(PrintMode = 1, PView, Printer), ViewPage
            End With
            
'            Call PRT_P03001(cdPrt, IIf(PrintMode = 1, PView, Printer), PrtParam)
            
        Case "P_04001"
            HScroll.Max = 100
            VScroll.Max = 100
            With CPrt
                If ViewPage = 0 Then ViewPage = 1
                .PRT_04001_01 IIf(PrintMode = 1, PView, Printer), ViewPage
            End With
        
        Case "P_04001_MASTER"
            HScroll.Max = 100
            VScroll.Max = 100
            With CPrt
                If ViewPage = 0 Then ViewPage = 1
                .PRT_04001_01_MASTER IIf(PrintMode = 1, PView, Printer), ViewPage
            End With
        
        Case Else
        
    End Select
        
        
End Sub

Public Sub DataPrint()
    PrintMode = 2
    Call DataPrintView
End Sub
Private Sub Form_Unload(Cancel As Integer)
    FORM_P_PRTSCREEN = False
    Set P_PRTSCREEN = Nothing
End Sub

Private Sub FixedFormReSize()
    Call cob_Zoom_Click
    
    PMain.Move 57, 850, (Me.Width) - 400, (Me.Height) - 1600
    VScroll.Move PMain.Width + 1, PMain.Top, 300, PMain.Height
    HScroll.Move PMain.Left, PMain.Height + 820, PMain.Width, 300
    
    Call PrnRefresh(PView, iSL_Type)

'    Debug.Print PView.Width
'    Debug.Print PView.Height
'
'    Printer.PaperSize = vbPRPSA4            ' 용지 크기
'    Printer.Orientation = vbPRORPortrait    ' 출력 방향 세로
'    Printer.ScaleMode = vbMillimeters       ' 밀리 단위로
'
'    Debug.Print Printer.PaperSize
'    Debug.Print Printer.ScaleWidth
'    Debug.Print Printer.ScaleHeight
'
'
'    PView.Height = Printer.Height * (Printer.Width / PView.Width)
'    Screen.MousePointer = vbHourglass
'    UseW = Printer.Width / Printer.TwipsPerPixelX * 0.95
'    UseH = Printer.Height / Printer.TwipsPerPixelY * 0.95
'
'    Printer.PaperSize = vbPRPSA4
'
'    PView.ScaleMode = vbMillimeters
    
End Sub

Private Sub VScroll_Change()
    On Error Resume Next
    PView.Top = -VScroll.Value * 100
End Sub
