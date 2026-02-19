VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm설문조사2 
   BackColor       =   &H00D9E5E9&
   BorderStyle     =   1  '단일 고정
   Caption         =   "설문조사"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
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
   Icon            =   "frm설문조사2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   9450
   StartUpPosition =   1  '소유자 가운데
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8220
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   14499
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm설문조사2.frx":08CA
      Begin Threed.SSPanel SSPanel 
         Height          =   7200
         Index           =   1
         Left            =   15
         TabIndex        =   5
         Top             =   435
         Width           =   9420
         _ExtentX        =   16616
         _ExtentY        =   12700
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtMsg3 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1065
            Left            =   105
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   15
            Top             =   6030
            Width           =   9240
         End
         Begin VB.TextBox txtMsg2 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1065
            Left            =   105
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   14
            Top             =   4260
            Width           =   9240
         End
         Begin VB.TextBox txtMsg1 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1065
            Left            =   105
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   13
            Top             =   2475
            Width           =   9240
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   585
            Index           =   2
            Left            =   105
            TabIndex        =   6
            Top             =   1890
            Width           =   9240
            _ExtentX        =   16298
            _ExtentY        =   1032
            _Version        =   262144
            Font3D          =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   " 1. 요금 인상 요청 품목"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "정장하의 : 2,000 -> 2,300"
               ForeColor       =   &H000000C0&
               Height          =   180
               Index           =   1
               Left            =   6900
               TabIndex        =   17
               Top             =   330
               Width           =   2250
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "예) 정장상의 : 2,500 -> 2,800"
               ForeColor       =   &H000000C0&
               Height          =   180
               Index           =   0
               Left            =   6540
               TabIndex        =   16
               Top             =   90
               Width           =   2610
            End
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   585
            Index           =   3
            Left            =   105
            TabIndex        =   7
            Top             =   3675
            Width           =   9240
            _ExtentX        =   16298
            _ExtentY        =   1032
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   " 2. 주변경쟁사 요금 조사"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "정장하의 : 2,300"
               ForeColor       =   &H000000C0&
               Height          =   180
               Index           =   3
               Left            =   6900
               TabIndex        =   19
               Top             =   330
               Width           =   1440
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "예) 정장상의 : 2,800"
               ForeColor       =   &H000000C0&
               Height          =   180
               Index           =   2
               Left            =   6540
               TabIndex        =   18
               Top             =   90
               Width           =   1800
            End
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   585
            Index           =   4
            Left            =   105
            TabIndex        =   8
            Top             =   5445
            Width           =   9240
            _ExtentX        =   16298
            _ExtentY        =   1032
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   " 3. 종합의견"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   $"frm설문조사2.frx":093C
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   390
            Index           =   6
            Left            =   135
            TabIndex        =   12
            Top             =   1305
            Width           =   9105
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "자재가격의 인상(물가 및 유가상승)으로 인하여 부득이하게 요금을 인상하게 되었음을 알려드립니다."
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   390
            Index           =   5
            Left            =   135
            TabIndex        =   11
            Top             =   870
            Width           =   9105
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "<요금인상에 대한 사유>"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   2
            Left            =   135
            TabIndex        =   10
            Top             =   600
            Width           =   2310
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "크린에이드의 발전을 위하여 정확한 작성 당부 드립니다."
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   1
            Left            =   135
            TabIndex        =   9
            Top             =   135
            Width           =   6810
         End
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   555
         Index           =   0
         Left            =   15
         TabIndex        =   2
         Top             =   7650
         Width           =   9420
         _ExtentX        =   16616
         _ExtentY        =   979
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   45
            TabIndex        =   3
            Top             =   60
            Width           =   1275
            _Version        =   851970
            _ExtentX        =   2249
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            BackColor       =   -2147483633
            Appearance      =   6
            Picture         =   "frm설문조사2.frx":09C6
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   8100
            TabIndex        =   4
            Top             =   60
            Width           =   1260
            _Version        =   851970
            _ExtentX        =   2222
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 저장(&S)"
            BackColor       =   -2147483633
            Appearance      =   6
            Picture         =   "frm설문조사2.frx":13D8
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   405
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   9420
         _ExtentX        =   16616
         _ExtentY        =   714
         _Version        =   262144
         Font3D          =   3
         ForeColor       =   0
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
         Caption         =   "   설문조사 - 요금인상에 대한 사유"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm설문조사2.frx":1DEA
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image1 
            Height          =   240
            Left            =   60
            Picture         =   "frm설문조사2.frx":224C
            Top             =   60
            Width           =   240
         End
      End
   End
End
Attribute VB_Name = "frm설문조사2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn

    Select Case Index
        Case 0
            Query = "SELECT * FROM TB_설문조사01"
            Query = Query & " WHERE KEY_DATE  = '110707'"
            Query = Query & "   AND STORE_CD = '" & 가맹점정보.가맹점코드 & "'"
            Set ADORs = New ADODB.Recordset
            ADORs.Open Query, HostCon, adOpenDynamic, adLockOptimistic

            If ADORs.EOF Then ADORs.AddNew

            ADORs!KEY_DATE = "110707"                   '
            ADORs!STORE_CD = 가맹점정보.가맹점코드 & "" '
            ADORs!MASTER_CD = 가맹점정보.지사코드 & ""  '
            ADORs!msg1 = txtMsg1.Text & ""              '
            ADORs!msg2 = txtMsg2.Text & ""              '
            ADORs!msg3 = txtMsg3.Text & ""              '
            ADORs!ActionDate = Date                     '
            
            ADORs.Update

            ADORs.Close
            Set ADORs = Nothing

            Unload Me

        Case 1:
            Unload Me
    End Select

    Exit Sub

ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)

    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn

    

    If Server_Connection(HostCon, "LAUNDRY1000") = True Then
        Query = "SELECT    Msg1"
        Query = Query & ", Msg2"
        Query = Query & ", Msg3"
        Query = Query & " FROM TB_설문조사01"
        Query = Query & " WHERE KEY_DATE  = '110707'"
        Query = Query & "   AND STORE_CD = '" & 가맹점정보.가맹점코드 & "'"
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, HostCon, adOpenForwardOnly, adLockReadOnly

        If Not ADORs.EOF Then
            txtMsg1.Text = ADORs!msg1 & ""
            txtMsg2.Text = ADORs!msg2 & ""
            txtMsg3.Text = ADORs!msg3 & ""
        End If
        ADORs.Close
        Set ADORs = Nothing
    End If

    Exit Sub

ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)

    Screen.MousePointer = 0
End Sub
