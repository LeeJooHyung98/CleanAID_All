VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm일일판매리스트1 
   BorderStyle     =   1  '단일 고정
   Caption         =   "일일판매리스트 출력"
   ClientHeight    =   2205
   ClientLeft      =   4275
   ClientTop       =   7110
   ClientWidth     =   4530
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form29"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4530
   StartUpPosition =   1  '소유자 가운데
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   2205
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   3889
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm일일판매리스트1.frx":0000
      Begin Threed.SSPanel SSPanel3 
         Height          =   660
         Left            =   15
         TabIndex        =   5
         Top             =   1530
         Width           =   4500
         _ExtentX        =   7938
         _ExtentY        =   1164
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton sscView 
            Height          =   540
            Left            =   75
            TabIndex        =   6
            Top             =   45
            Width           =   1650
            _Version        =   851970
            _ExtentX        =   2910
            _ExtentY        =   952
            _StockProps     =   79
            Caption         =   " 미리보기"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Picture         =   "frm일일판매리스트1.frx":0052
         End
         Begin XtremeSuiteControls.PushButton cmdPrint 
            Height          =   540
            Left            =   1890
            TabIndex        =   7
            Top             =   45
            Width           =   1245
            _Version        =   851970
            _ExtentX        =   2196
            _ExtentY        =   952
            _StockProps     =   79
            Caption         =   " 출력"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Picture         =   "frm일일판매리스트1.frx":0A64
         End
         Begin XtremeSuiteControls.PushButton cmdExit 
            Height          =   540
            Left            =   3180
            TabIndex        =   8
            Top             =   45
            Width           =   1245
            _Version        =   851970
            _ExtentX        =   2196
            _ExtentY        =   952
            _StockProps     =   79
            Caption         =   " 닫기"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Picture         =   "frm일일판매리스트1.frx":1476
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   0
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1500
         Left            =   15
         TabIndex        =   4
         Top             =   15
         Width           =   4500
         _ExtentX        =   7938
         _ExtentY        =   2646
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtDay 
            Alignment       =   2  '가운데 맞춤
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   18
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   3180
            TabIndex        =   2
            Text            =   "10"
            Top             =   720
            Width           =   705
         End
         Begin VB.TextBox txtMon 
            Alignment       =   2  '가운데 맞춤
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   18
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   1890
            TabIndex        =   1
            Text            =   "10"
            Top             =   720
            Width           =   705
         End
         Begin VB.TextBox txtYr 
            Alignment       =   2  '가운데 맞춤
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   18
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   225
            TabIndex        =   0
            Text            =   "2010"
            Top             =   720
            Width           =   1140
         End
         Begin Threed.SSPanel SSPanel1 
            Height          =   495
            Left            =   60
            TabIndex        =   9
            Top             =   60
            Width           =   4380
            _ExtentX        =   7726
            _ExtentY        =   873
            _Version        =   262144
            ForeColor       =   16711680
            Windowless      =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "일일판매리스트"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm일일판매리스트1.frx":1E88
            BorderWidth     =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "       년         월         일"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   15.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   675
            TabIndex        =   10
            Top             =   750
            Width           =   3720
         End
      End
   End
End
Attribute VB_Name = "frm일일판매리스트1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strDate As String
Dim strDate1 As String
Dim strWeekDay As String

Private Sub cmdExit_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    '**************************************************************************
    '이름 : 일일판매리트 출력
    '기능 : 일자를 받아서 그일자에 있는 자료를 출력
    '**************************************************************************
    Dim strDate As String
    
     '일자체크
    strDate = txtYr.Text & "-" & Format(txtMon.Text, "00") & "-" & Format(txtDay.Text, "00")
    
    If Not IsDate(strDate) Then
        MsgBox " 연도/월/일을 다시 확인하세요..! "
        txtYr.SetFocus
        Exit Sub
    End If
    
    ' 도트 , 잉크젯 공용
    Call subDayListPrint(CommonDialog1, strDate, False)
    Exit Sub
End Sub

Private Sub Form_Activate()
    txtYr.Text = Format(Date, "yyyy")
    txtMon.Text = Format(Date, "mm")
    txtDay.Text = Format(Date, "dd")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{Tab}"
        KeyCode = 0
        DoEvents
    End If
End Sub

Private Sub Form_Load()
    Me.Top = 2900
    Me.Left = 4000
End Sub

Private Sub txtDay_GotFocus()
    txtDay.SelStart = 0
    txtDay.SelLength = 2
End Sub

Private Sub txtMon_GotFocus()
    txtMon.SelStart = 0
    txtMon.SelLength = 2
End Sub

Private Sub txtYr_GotFocus()
    txtYr.SelStart = 0
    txtYr.SelLength = 4
End Sub

Private Sub sscView_Click()
    '**************************************************************************
    '이름 : 일일판매리트 미리보기
    '기능 : 일자를 받아서 그일자에 있는 자료를 출력
    '**************************************************************************
    Dim strDate As String
    
     '일자체크
    strDate = txtYr.Text & "-" & Format(txtMon.Text, "00") & "-" & Format(txtDay.Text, "00")
    
    If Not IsDate(strDate) Then
        MsgBox " 연도/월/일을 다시 확인하세요..! "
        txtYr.SetFocus
        Exit Sub
    End If
    
    ' 도트 , 잉크젯 공용
    Call subDayListPrint(CommonDialog1, strDate, True)
    
    Exit Sub
End Sub
