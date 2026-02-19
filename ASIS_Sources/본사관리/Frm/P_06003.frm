VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMask32.ocx"
Begin VB.Form P_06003 
   Caption         =   "사고접수 보고서"
   ClientHeight    =   11655
   ClientLeft      =   1020
   ClientTop       =   5535
   ClientWidth     =   16710
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_06003.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11655
   ScaleWidth      =   16710
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   11655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16710
      _ExtentX        =   29475
      _ExtentY        =   20558
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_06003.frx":058A
      Begin Threed.SSPanel panInput 
         Height          =   765
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   16680
         _ExtentX        =   29422
         _ExtentY        =   1349
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   6
            Left            =   8085
            Style           =   2  '드롭다운 목록
            TabIndex        =   2
            Top             =   60
            Width           =   5115
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            TabIndex        =   3
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   64552960
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   4
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "접 수 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   19
            Left            =   4800
            TabIndex        =   5
            Top             =   60
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "접수일자 / 접수번호 / 매장명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel panMain 
         Height          =   10845
         Left            =   15
         TabIndex        =   6
         Top             =   795
         Width           =   16680
         _ExtentX        =   29422
         _ExtentY        =   19129
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin Threed.SSFrame SSFrame3 
            Height          =   1455
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   4320
            Width           =   15135
            _ExtentX        =   26696
            _ExtentY        =   2566
            _Version        =   262144
            Caption         =   "대리점 기재"
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   15
               Left            =   1650
               TabIndex        =   10
               Top             =   300
               Width           =   13215
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   16
               Left            =   1650
               TabIndex        =   9
               Top             =   660
               Width           =   13215
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   17
               Left            =   1650
               TabIndex        =   8
               Top             =   1020
               Width           =   13215
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   20
               Left            =   180
               TabIndex        =   11
               Top             =   300
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "사고의 종류"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   21
               Left            =   180
               TabIndex        =   12
               Top             =   660
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "사고의 내용"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   22
               Left            =   180
               TabIndex        =   13
               Top             =   1020
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "소비자 의견"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
         End
         Begin Threed.SSFrame SSFrame2 
            Height          =   2175
            Left            =   120
            TabIndex        =   14
            Top             =   2040
            Width           =   15135
            _ExtentX        =   26696
            _ExtentY        =   3836
            _Version        =   262144
            Caption         =   "피해관련사항"
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   8
               Left            =   3570
               TabIndex        =   21
               Top             =   300
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   9
               Left            =   9150
               TabIndex        =   20
               Top             =   300
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   10
               Left            =   9150
               TabIndex        =   19
               Top             =   660
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   11
               Left            =   3570
               TabIndex        =   18
               Top             =   1020
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   12
               Left            =   9150
               TabIndex        =   17
               Top             =   1020
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   13
               Left            =   3570
               TabIndex        =   16
               Top             =   1380
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   14
               Left            =   3570
               TabIndex        =   15
               Top             =   1740
               Width           =   3735
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   9
               Left            =   2100
               TabIndex        =   22
               Top             =   300
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "품      목"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   10
               Left            =   7680
               TabIndex        =   23
               Top             =   300
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "상      표"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   11
               Left            =   2100
               TabIndex        =   24
               Top             =   660
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "구 입 일 자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtInput 
               Height          =   315
               Index           =   1
               Left            =   3570
               TabIndex        =   25
               Top             =   660
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   64552960
               CurrentDate     =   36686
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   12
               Left            =   7680
               TabIndex        =   26
               Top             =   660
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "색      상"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   13
               Left            =   2100
               TabIndex        =   27
               Top             =   1020
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "구  입  처"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   14
               Left            =   7680
               TabIndex        =   28
               Top             =   1020
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "택  번  호"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   15
               Left            =   2100
               TabIndex        =   29
               Top             =   1380
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "구 입 형 태"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   16
               Left            =   7680
               TabIndex        =   30
               Top             =   1380
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "입 고 일 자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtInput 
               Height          =   315
               Index           =   2
               Left            =   9150
               TabIndex        =   31
               Top             =   1380
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   64552960
               CurrentDate     =   36686
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   17
               Left            =   2100
               TabIndex        =   32
               Top             =   1740
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "구 입 가 격"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   18
               Left            =   7680
               TabIndex        =   33
               Top             =   1740
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "사고 접수일"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtInput 
               Height          =   315
               Index           =   3
               Left            =   9150
               TabIndex        =   34
               Top             =   1740
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   64552960
               CurrentDate     =   36686
            End
         End
         Begin Threed.SSFrame SSFrame1 
            Height          =   1815
            Left            =   120
            TabIndex        =   35
            Top             =   120
            Width           =   15135
            _ExtentX        =   26696
            _ExtentY        =   3201
            _Version        =   262144
            Caption         =   "기본사항"
            Begin VB.TextBox txtInput 
               Height          =   675
               Index           =   0
               Left            =   1650
               TabIndex        =   43
               Top             =   300
               Width           =   2415
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   1
               Left            =   5730
               TabIndex        =   42
               Top             =   300
               Width           =   9135
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   2
               Left            =   5730
               TabIndex        =   41
               Top             =   660
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   3
               Left            =   11130
               TabIndex        =   40
               Top             =   660
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   675
               Index           =   4
               Left            =   1650
               TabIndex        =   39
               Top             =   1020
               Width           =   2415
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   5
               Left            =   5730
               TabIndex        =   38
               Top             =   1020
               Width           =   9135
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   6
               Left            =   5730
               TabIndex        =   37
               Top             =   1380
               Width           =   3735
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   7
               Left            =   11130
               TabIndex        =   36
               Top             =   1380
               Width           =   3735
            End
            Begin Threed.SSPanel panCaption 
               Height          =   675
               Index           =   1
               Left            =   180
               TabIndex        =   44
               Top             =   300
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   1191
               _Version        =   262144
               Caption         =   "대 리 점 명"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   2
               Left            =   4260
               TabIndex        =   45
               Top             =   300
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "주      소"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   3
               Left            =   4260
               TabIndex        =   46
               Top             =   660
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "성      명"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   4
               Left            =   9660
               TabIndex        =   47
               Top             =   660
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "전      화"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   675
               Index           =   5
               Left            =   180
               TabIndex        =   48
               Top             =   1020
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   1191
               _Version        =   262144
               Caption         =   "소 비 자 명"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   6
               Left            =   4260
               TabIndex        =   49
               Top             =   1020
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "주      소"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   7
               Left            =   4260
               TabIndex        =   50
               Top             =   1380
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "전      화"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   8
               Left            =   9660
               TabIndex        =   51
               Top             =   1380
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "핸  드  폰"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
         End
         Begin Threed.SSFrame SSFrame3 
            Height          =   1455
            Index           =   1
            Left            =   120
            TabIndex        =   52
            Top             =   5880
            Width           =   15135
            _ExtentX        =   26696
            _ExtentY        =   2566
            _Version        =   262144
            Caption         =   "본사 기재"
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   18
               Left            =   1650
               TabIndex        =   60
               Top             =   300
               Width           =   3315
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   19
               Left            =   6630
               TabIndex        =   59
               Top             =   300
               Width           =   3315
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   21
               Left            =   6630
               TabIndex        =   58
               Top             =   660
               Width           =   3315
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   23
               Left            =   1650
               TabIndex        =   57
               Top             =   1020
               Width           =   3315
            End
            Begin VB.TextBox txtInput 
               Height          =   315
               Index           =   25
               Left            =   11610
               TabIndex        =   56
               Top             =   1020
               Width           =   3255
            End
            Begin VB.ComboBox cboInput 
               Height          =   315
               Index           =   0
               Left            =   1650
               Style           =   2  '드롭다운 목록
               TabIndex        =   53
               Top             =   660
               Width           =   3315
            End
            Begin MSMask.MaskEdBox mskInput 
               Height          =   315
               Index           =   1
               Left            =   6630
               TabIndex        =   54
               Top             =   1020
               Width           =   3315
               _ExtentX        =   5847
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "#,##0"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskInput 
               Height          =   315
               Index           =   0
               Left            =   11610
               TabIndex        =   55
               Top             =   660
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "#,##0"
               PromptChar      =   "_"
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   23
               Left            =   180
               TabIndex        =   61
               Top             =   300
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "제 조 회 사"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   24
               Left            =   5160
               TabIndex        =   62
               Top             =   300
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "전 화 번 호"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   25
               Left            =   10140
               TabIndex        =   63
               Top             =   300
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "판 매 일 자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtInput 
               Height          =   315
               Index           =   4
               Left            =   11610
               TabIndex        =   64
               Top             =   300
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   64552960
               CurrentDate     =   36686
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   26
               Left            =   180
               TabIndex        =   65
               Top             =   660
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "재 고 현 황"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   27
               Left            =   5160
               TabIndex        =   66
               Top             =   660
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "담      당"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   28
               Left            =   10140
               TabIndex        =   67
               Top             =   660
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "판 매 금 액"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   29
               Left            =   180
               TabIndex        =   68
               Top             =   1020
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "보 상 비 율"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   30
               Left            =   5160
               TabIndex        =   69
               Top             =   1020
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "보상산정금액"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   31
               Left            =   10140
               TabIndex        =   70
               Top             =   1020
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "합 의 내 용"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
         End
      End
   End
End
Attribute VB_Name = "P_06003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cboInput_Click(Index As Integer)
    If Index = 6 Then
        dtInput(0).Value = Format(Mid(cboInput(6).Text, 1, 10), "YYYY-MM-DD")
    End If
End Sub

Private Sub dtInput_Change(Index As Integer)
    If Index = 0 Then
        ReDim sValue(2)
        
        sValue(0) = "0"
        sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_06001_00", sValue(), Err_Num, Err_Dec)
        
        cboInput(6).Clear
        
        Do While Not RS01.EOF
            cboInput(6).AddItem Format(RS01!접수일자, "YYYY-MM-DD") & " / " & RS01!접수번호 & " / " & RS01!매장명
        
            RS01.MoveNext
        Loop
    End If
End Sub

Private Sub Form_Activate()
'    cmdBtn(0).Enabled = True
'    cmdBtn(2).Enabled = True
'    cmdBtn(5).Enabled = True
'    cmdBtn(6).Enabled = True
'
'    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    If P_06003_Flag = False Then
        Call ComboAdd
        
        dtInput(0).Value = Date
        
        ReDim sValue(2)
        
        sValue(0) = "0"
        sValue(1) = Format(dtInput(0).Value, "yyyy")
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_06001_00", sValue(), Err_Num, Err_Dec)    ' 2002.12.07일  SP_06003_02 에서 변경
        
        cboInput(6).Clear
        
        
        Do While Not RS01.EOF
            cboInput(6).AddItem Format(RS01!접수일자, "YYYY-MM-DD") & " / " & RS01!접수번호 & " / " & RS01!매장명
        
            RS01.MoveNext
        Loop
        
        P_06003_Flag = True
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i As Integer
    
    ReDim sValue(2)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Trim(Mid(cboInput(6).Text, 14, 4))
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_06003_00", sValue(), Err_Num, Err_Dec)
    
    If RS01.RecordCount <> 0 Then
        txtInput(0).Text = RS01!대리점명 & ""
        txtInput(1).Text = RS01!대리점주소 & ""
        txtInput(2).Text = RS01!대리점주 & ""
        txtInput(3).Text = RS01!대리점전화번호 & ""
        
        txtInput(4).Text = RS01!고객성명 & ""
        txtInput(5).Text = RS01!고객주소 & ""
        txtInput(6).Text = RS01!고객전화번호 & ""
        txtInput(7).Text = RS01!고객핸드폰 & ""
        
        txtInput(8).Text = RS01!품명 & ""
        txtInput(9).Text = RS01!브랜드 & ""
        
        If Trim(RS01!구입일자) <> "" Then dtInput(1).Value = Format(RS01!구입일자, "####-##-##") Else dtInput(1).Value = ""
        
        If Not IsNull(RS01!색상) Then txtInput(10).Text = RS01!색상 Else txtInput(10).Text = ""
        If Not IsNull(RS01!구입처) Then txtInput(11).Text = RS01!구입처 Else txtInput(11).Text = ""
        If Not IsNull(RS01!택번호) Then txtInput(12).Text = Format(RS01!택번호, "@-@@@")
        If Not IsNull(RS01!구입형태) Then txtInput(13).Text = RS01!구입형태 Else txtInput(13).Text = ""
        If Trim(RS01!입고일자) <> "" Then dtInput(2).Value = Format(RS01!입고일자, "####-##-##") Else dtInput(2).Value = ""
        If Not IsNull(RS01!구입가격) Then txtInput(14).Text = Format(RS01!구입가격, "#,##0") Else txtInput(14).Text = ""
        If Trim(RS01!접수일자) <> "" Then dtInput(3).Value = Format(RS01!접수일자, "####-##-##") Else dtInput(3).Value = ""
        
        If Not IsNull(RS01!크레임구분) Then txtInput(15).Text = RS01!크레임구분 Else txtInput(15).Text = ""
        If Not IsNull(RS01!비고) Then txtInput(16).Text = RS01!브랜드 Else txtInput(16).Text = ""
        If Not IsNull(RS01!대리점의견1) Then txtInput(17).Text = RS01!대리점의견1 Else txtInput(17).Text = ""
        
        If Not IsNull(RS01!제조회사) Then txtInput(18).Text = RS01!제조회사 Else txtInput(18).Text = ""
        If Not IsNull(RS01!전화번호) Then txtInput(19).Text = RS01!전화번호 Else txtInput(19).Text = ""
        
        If Not IsNull(RS01!재고현황) Then
            For i = 0 To cboInput(0).ListCount - 1
                If RS01!재고현황 = Mid(cboInput(0).List(i), 2, 1) Then
                    cboInput(0).ListIndex = i
                    Exit For
                End If
            Next i
        Else
            cboInput(0).ListIndex = -1
        End If
        
        If Not IsNull(RS01!담당) Then txtInput(21).Text = RS01!담당 Else txtInput(21).Text = ""
        If Not IsNull(RS01!판매금액) Then mskInput(0).Text = RS01!판매금액 Else mskInput(0).Text = ""
        If Not IsNull(RS01!보상비율) Then txtInput(23).Text = RS01!보상비율 Else txtInput(23).Text = ""
        If Not IsNull(RS01!보상산정금액) Then mskInput(1).Text = RS01!보상산정금액 Else mskInput(1).Text = ""
        If Not IsNull(RS01!합의내용) Then txtInput(25).Text = RS01!합의내용 Else txtInput(25).Text = ""
    Else
        MsgBox "해당되는 데이타가 존재하지 않습니다.", vbInformation
    End If
    
    RS01.Close
            
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Public Sub DataScreen()
'    Dim ReportFP As String
'    Dim ReportFile As String
'    Dim sData As String
'
'    ReportFP = GetIniStr("REPORT", "FilePath", "", m_iniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'
'    Dim i As Integer
'    For i = 0 To 30
'        P_00000.crPrint.Formulas(i) = ""
'    Next
'
'    P_00000.crPrint.Formulas(0) = "Data01 = '" & txtInput(0).Text & "'"
'    P_00000.crPrint.Formulas(1) = "Data02 = '" & txtInput(1).Text & "'"
'    P_00000.crPrint.Formulas(2) = "Data03 = '" & txtInput(2).Text & "'"
'    P_00000.crPrint.Formulas(3) = "Data04 = '" & txtInput(3).Text & "'"
'    P_00000.crPrint.Formulas(4) = "Data05 = '" & txtInput(4).Text & "'"
'    P_00000.crPrint.Formulas(5) = "Data06 = '" & txtInput(5).Text & "'"
'    P_00000.crPrint.Formulas(6) = "Data07 = '" & txtInput(6).Text & "'"
'    P_00000.crPrint.Formulas(7) = "Data08 = '" & txtInput(7).Text & "'"
'    P_00000.crPrint.Formulas(8) = "Data09 = '" & txtInput(8).Text & "'"
'    P_00000.crPrint.Formulas(9) = "Data10 = '" & txtInput(9).Text & "'"
'    P_00000.crPrint.Formulas(10) = "Data11 = '" & Format(dtInput(1).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(11) = "Data12 = '" & txtInput(10).Text & "'"
'    P_00000.crPrint.Formulas(12) = "Data13 = '" & txtInput(11).Text & "'"
'    P_00000.crPrint.Formulas(13) = "Data14 = '" & txtInput(12).Text & "'"
'    P_00000.crPrint.Formulas(14) = "Data15 = '" & txtInput(13).Text & "'"
'    P_00000.crPrint.Formulas(15) = "Data16 = '" & Format(dtInput(2).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(16) = "Data17 = '" & txtInput(14).Text & "'"
'    P_00000.crPrint.Formulas(17) = "Data18 = '" & Format(dtInput(3).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(18) = "Data19 = '" & txtInput(18).Text & "'"
'    P_00000.crPrint.Formulas(19) = "Data20 = '" & txtInput(19).Text & "'"
'    P_00000.crPrint.Formulas(20) = "Data21 = '" & Format(dtInput(4).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(21) = "Data22 = '" & cboInput(0).Text & "'"
'    P_00000.crPrint.Formulas(22) = "Data23 = '" & txtInput(21).Text & "'"
'    P_00000.crPrint.Formulas(23) = "Data24 = '" & Format(mskInput(0).Text, "#,##0") & "'"
'    P_00000.crPrint.Formulas(24) = "Data25 = '" & txtInput(23).Text & "'"
'    P_00000.crPrint.Formulas(25) = "Data26 = '" & Format(mskInput(1).Text, "#,##0") & "'"
'    P_00000.crPrint.Formulas(26) = "Data27 = '" & txtInput(25).Text & "'"
'
'    Call ReportPrint(ReportFile, "2")
End Sub

Public Sub DataPrint()
'    Dim ReportFP As String
'    Dim ReportFile As String
'    Dim sData As String
'
'    ReportFP = GetIniStr("REPORT", "FilePath", "", m_iniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'
'    Dim i As Integer
'    For i = 0 To 30
'        P_00000.crPrint.Formulas(i) = ""
'    Next
'
'    P_00000.crPrint.Formulas(0) = "Data01 = '" & txtInput(0).Text & "'"
'    P_00000.crPrint.Formulas(1) = "Data02 = '" & txtInput(1).Text & "'"
'    P_00000.crPrint.Formulas(2) = "Data03 = '" & txtInput(2).Text & "'"
'    P_00000.crPrint.Formulas(3) = "Data04 = '" & txtInput(3).Text & "'"
'    P_00000.crPrint.Formulas(4) = "Data05 = '" & txtInput(4).Text & "'"
'    P_00000.crPrint.Formulas(5) = "Data06 = '" & txtInput(5).Text & "'"
'    P_00000.crPrint.Formulas(6) = "Data07 = '" & txtInput(6).Text & "'"
'    P_00000.crPrint.Formulas(7) = "Data08 = '" & txtInput(7).Text & "'"
'    P_00000.crPrint.Formulas(8) = "Data09 = '" & txtInput(8).Text & "'"
'    P_00000.crPrint.Formulas(9) = "Data10 = '" & txtInput(9).Text & "'"
'    P_00000.crPrint.Formulas(10) = "Data11 = '" & Format(dtInput(1).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(11) = "Data12 = '" & txtInput(10).Text & "'"
'    P_00000.crPrint.Formulas(12) = "Data13 = '" & txtInput(11).Text & "'"
'    P_00000.crPrint.Formulas(13) = "Data14 = '" & txtInput(12).Text & "'"
'    P_00000.crPrint.Formulas(14) = "Data15 = '" & txtInput(13).Text & "'"
'    P_00000.crPrint.Formulas(15) = "Data16 = '" & Format(dtInput(2).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(16) = "Data17 = '" & txtInput(14).Text & "'"
'    P_00000.crPrint.Formulas(17) = "Data18 = '" & Format(dtInput(3).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(18) = "Data19 = '" & txtInput(18).Text & "'"
'    P_00000.crPrint.Formulas(19) = "Data20 = '" & txtInput(19).Text & "'"
'    P_00000.crPrint.Formulas(20) = "Data21 = '" & Format(dtInput(4).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(21) = "Data22 = '" & cboInput(0).Text & "'"
'    P_00000.crPrint.Formulas(22) = "Data23 = '" & txtInput(21).Text & "'"
'    P_00000.crPrint.Formulas(23) = "Data24 = '" & Format(mskInput(0).Text, "#,##0") & "'"
'    P_00000.crPrint.Formulas(24) = "Data25 = '" & txtInput(23).Text & "'"
'    P_00000.crPrint.Formulas(25) = "Data26 = '" & Format(mskInput(1).Text, "#,##0") & "'"
'    P_00000.crPrint.Formulas(26) = "Data27 = '" & txtInput(25).Text & "'"
'
'    Call ReportPrint(ReportFile, "1")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_06003_Flag = False
End Sub

Public Sub DataSave()
    If MsgBox("해당되는 내역을 저장하시겠습니까?", vbYesNo + vbInformation, "데이터 저장") = vbYes Then
        ReDim sValue(10)
        
        sValue(0) = Format(dtInput(0).Value, "YYYY-MM-DD")        ' 접수일자
        sValue(1) = Mid(cboInput(6).Text, 1, 4)                 ' 접수번호
        sValue(2) = txtInput(18).Text                           ' 제조회사
        sValue(3) = txtInput(19).Text                           ' 전화번호
        sValue(4) = Format(dtInput(4).Value, "YYYY-MM-DD")        ' 판매일자
        sValue(5) = Mid(cboInput(0).Text, 2, 1)                 ' 재고현황
        sValue(6) = txtInput(21).Text                           ' 담당
        If mskInput(0).ClipText = "" Then
            sValue(7) = 0
        Else
            sValue(7) = mskInput(0).ClipText                    ' 판매금액
        End If
        
        If txtInput(23).Text = "" Then
            sValue(8) = 0
        Else
            sValue(8) = txtInput(23).Text                       ' 보상비율
        End If
        
        If mskInput(1).ClipText = "" Then
            sValue(9) = 0
        Else
            sValue(9) = mskInput(1).ClipText                    ' 판매금액
        End If
        
        sValue(10) = txtInput(25).Text                          ' 합의내용
        
        Call ExecPro("SP_06003_01", sValue(), Err_Num, Err_Dec)
    
        If Err_Num = 0 Then
            MsgBox "해당되는 데이터가 정상적으로 저장이 되었습니다.", vbInformation
        End If
    End If
End Sub

Private Sub ComboAdd()
    cboInput(0).AddItem "[1] 유"
    cboInput(0).AddItem "[2] 무"
End Sub
