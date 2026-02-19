VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm작업 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "내용"
   ClientHeight    =   3465
   ClientLeft      =   7995
   ClientTop       =   5235
   ClientWidth     =   10590
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawWidth       =   9
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkMode        =   1  '원본
   LinkTopic       =   "Form13"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   3465
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   6112
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm작업.frx":0000
      Begin XtremeSuiteControls.TabControl TabControl 
         Height          =   3435
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   10560
         _Version        =   851970
         _ExtentX        =   18627
         _ExtentY        =   6059
         _StockProps     =   68
         Appearance      =   4
         Color           =   64
         PaintManager.Position=   3
         ItemCount       =   2
         Item(0).Caption =   " 작업 "
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "TabControlPage(0)"
         Item(1).Caption =   ""
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "TabControlPage(1)"
         Begin XtremeSuiteControls.TabControlPage TabControlPage 
            Height          =   5670
            Index           =   1
            Left            =   -69970
            TabIndex        =   2
            Top             =   30
            Visible         =   0   'False
            Width           =   13110
            _Version        =   851970
            _ExtentX        =   23125
            _ExtentY        =   10001
            _StockProps     =   1
            Page            =   1
            Begin Threed.SSPanel SSPanel1 
               Height          =   1080
               Left            =   6075
               TabIndex        =   21
               Top             =   0
               Width           =   3435
               _ExtentX        =   6059
               _ExtentY        =   1905
               _Version        =   262144
               BackColor       =   13160660
               Windowless      =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderWidth     =   0
               BevelOuter      =   1
               BevelInner      =   2
               RoundedCorners  =   0   'False
               Begin VB.ComboBox Combo1 
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   405
                  ItemData        =   "frm작업.frx":0032
                  Left            =   150
                  List            =   "frm작업.frx":0034
                  Style           =   2  '드롭다운 목록
                  TabIndex        =   22
                  Top             =   510
                  Width           =   3150
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "할증 구분"
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   11.25
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   165
                  TabIndex        =   23
                  Top             =   120
                  Width           =   1155
               End
            End
            Begin VB.TextBox txtPassWord 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   9540
               TabIndex        =   3
               Top             =   1245
               Visible         =   0   'False
               Width           =   1635
            End
            Begin Threed.SSPanel pnlCost 
               Height          =   510
               Left            =   75
               TabIndex        =   4
               Top             =   105
               Width           =   1920
               _ExtentX        =   3387
               _ExtentY        =   900
               _Version        =   262144
               ForeColor       =   255
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin XtremeSuiteControls.PushButton Command24 
               Height          =   540
               Left            =   2040
               TabIndex        =   5
               Top             =   90
               Width           =   945
               _Version        =   851970
               _ExtentX        =   1667
               _ExtentY        =   952
               _StockProps     =   79
               BackColor       =   -2147483633
               Appearance      =   6
               Picture         =   "frm작업.frx":0036
            End
            Begin XtremeSuiteControls.PushButton btnNumber 
               Height          =   765
               Index           =   0
               Left            =   3030
               TabIndex        =   6
               Top             =   1560
               Width           =   945
               _Version        =   851970
               _ExtentX        =   1667
               _ExtentY        =   1349
               _StockProps     =   79
               Caption         =   "0"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
            End
            Begin XtremeSuiteControls.PushButton btnNumber 
               Height          =   765
               Index           =   1
               Left            =   60
               TabIndex        =   7
               Top             =   750
               Width           =   945
               _Version        =   851970
               _ExtentX        =   1667
               _ExtentY        =   1349
               _StockProps     =   79
               Caption         =   "1"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
            End
            Begin XtremeSuiteControls.PushButton btnNumber 
               Height          =   765
               Index           =   2
               Left            =   1050
               TabIndex        =   8
               Top             =   750
               Width           =   945
               _Version        =   851970
               _ExtentX        =   1667
               _ExtentY        =   1349
               _StockProps     =   79
               Caption         =   "2"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
            End
            Begin XtremeSuiteControls.PushButton btnNumber 
               Height          =   765
               Index           =   3
               Left            =   2040
               TabIndex        =   9
               Top             =   750
               Width           =   945
               _Version        =   851970
               _ExtentX        =   1667
               _ExtentY        =   1349
               _StockProps     =   79
               Caption         =   "3"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
            End
            Begin XtremeSuiteControls.PushButton btnNumber 
               Height          =   765
               Index           =   4
               Left            =   3030
               TabIndex        =   10
               Top             =   750
               Width           =   945
               _Version        =   851970
               _ExtentX        =   1667
               _ExtentY        =   1349
               _StockProps     =   79
               Caption         =   "4"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
            End
            Begin XtremeSuiteControls.PushButton btnNumber 
               Height          =   765
               Index           =   5
               Left            =   4020
               TabIndex        =   11
               Top             =   750
               Width           =   945
               _Version        =   851970
               _ExtentX        =   1667
               _ExtentY        =   1349
               _StockProps     =   79
               Caption         =   "5"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
            End
            Begin XtremeSuiteControls.PushButton btnNumber 
               Height          =   765
               Index           =   6
               Left            =   5010
               TabIndex        =   12
               Top             =   750
               Width           =   945
               _Version        =   851970
               _ExtentX        =   1667
               _ExtentY        =   1349
               _StockProps     =   79
               Caption         =   "6"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
            End
            Begin XtremeSuiteControls.PushButton btnNumber 
               Height          =   765
               Index           =   7
               Left            =   60
               TabIndex        =   13
               Top             =   1560
               Width           =   945
               _Version        =   851970
               _ExtentX        =   1667
               _ExtentY        =   1349
               _StockProps     =   79
               Caption         =   "7"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
            End
            Begin XtremeSuiteControls.PushButton btnNumber 
               Height          =   765
               Index           =   8
               Left            =   1050
               TabIndex        =   14
               Top             =   1560
               Width           =   945
               _Version        =   851970
               _ExtentX        =   1667
               _ExtentY        =   1349
               _StockProps     =   79
               Caption         =   "8"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
            End
            Begin XtremeSuiteControls.PushButton btnNumber 
               Height          =   765
               Index           =   9
               Left            =   2040
               TabIndex        =   15
               Top             =   1560
               Width           =   945
               _Version        =   851970
               _ExtentX        =   1667
               _ExtentY        =   1349
               _StockProps     =   79
               Caption         =   "9"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
            End
            Begin XtremeSuiteControls.PushButton btnNumber 
               Height          =   765
               Index           =   10
               Left            =   4020
               TabIndex        =   16
               Top             =   1560
               Width           =   945
               _Version        =   851970
               _ExtentX        =   1667
               _ExtentY        =   1349
               _StockProps     =   79
               Caption         =   "00"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
            End
            Begin XtremeSuiteControls.PushButton btnNumber 
               Height          =   765
               Index           =   11
               Left            =   5010
               TabIndex        =   17
               Top             =   1560
               Width           =   945
               _Version        =   851970
               _ExtentX        =   1667
               _ExtentY        =   1349
               _StockProps     =   79
               Caption         =   "000"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
            End
            Begin XtremeSuiteControls.PushButton cmdAgain 
               Height          =   540
               Left            =   3075
               TabIndex        =   18
               Top             =   90
               Width           =   1395
               _Version        =   851970
               _ExtentX        =   2461
               _ExtentY        =   952
               _StockProps     =   79
               Caption         =   "다시"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
            End
            Begin XtremeSuiteControls.PushButton cmdOK 
               Height          =   540
               Left            =   4560
               TabIndex        =   19
               Top             =   90
               Width           =   1395
               _Version        =   851970
               _ExtentX        =   2461
               _ExtentY        =   952
               _StockProps     =   79
               Caption         =   "완료"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
            End
            Begin XtremeSuiteControls.PushButton cmdHeadOffice 
               Height          =   540
               Left            =   6060
               TabIndex        =   20
               Top             =   90
               Visible         =   0   'False
               Width           =   1395
               _Version        =   851970
               _ExtentX        =   2461
               _ExtentY        =   952
               _StockProps     =   79
               Caption         =   "본사확인"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
            End
            Begin Threed.SSPanel pnlAddGubun 
               Height          =   1080
               Left            =   6075
               TabIndex        =   24
               Top             =   750
               Width           =   3435
               _ExtentX        =   6059
               _ExtentY        =   1905
               _Version        =   262144
               BackColor       =   13160660
               Windowless      =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderWidth     =   0
               BevelOuter      =   1
               BevelInner      =   2
               RoundedCorners  =   0   'False
               Begin VB.ComboBox Combo2 
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   405
                  ItemData        =   "frm작업.frx":0389
                  Left            =   150
                  List            =   "frm작업.frx":038B
                  Style           =   2  '드롭다운 목록
                  TabIndex        =   25
                  Top             =   510
                  Width           =   3150
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "할증 구분"
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   11.25
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   165
                  TabIndex        =   26
                  Top             =   120
                  Width           =   1155
               End
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage 
            Height          =   3375
            Index           =   0
            Left            =   30
            TabIndex        =   27
            Top             =   30
            Width           =   10185
            _Version        =   851970
            _ExtentX        =   17965
            _ExtentY        =   5953
            _StockProps     =   1
            Page            =   0
            Begin CleanAID.ctlMenu ctlMenu1 
               Height          =   825
               Index           =   0
               Left            =   75
               TabIndex        =   28
               Top             =   75
               Visible         =   0   'False
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   1455
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   20.25
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
      End
   End
End
Attribute VB_Name = "frm작업"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intMode   As Integer

Private strFirst As String
Private strPrice As String
Private strSu    As String
Private strJa    As String

Public m_sOldPrice As String
Public m_sOldStats As String

Const FORM_WIHTH As Long = 12180 '12000

Private Const MAX = 30


Private Sub ctlMenu1_Click(Index As Integer)
    Dim 작업코드 As String
    Dim 작업명   As String
    
    작업코드 = ctlMenu1(Index).GET_MenuKey
    작업명 = ctlMenu1(Index).GetMenuName
    
    If 작업코드 = "" Then Exit Sub
    
    
    
    ' 변경전 내용을 파악한다.
    frm접수.sprGrid.Row = frm접수.sprGrid.ActiveRow
    
    frm접수.sprGrid.Col = 5: m_sOldStats = Trim(frm접수.sprGrid.Text) '내용
    frm접수.sprGrid.Col = 14: m_sOldPrice = Trim(frm접수.sprGrid.Text) '내용
    
    Select Case 작업코드
        Case "w01": btn오염_Click
        Case "w02": btn하자_Click
        Case "w03": btn고가세탁_Click
        Case "w04": btn급자_Click
        Case "w05": btn행사품목_Click
        Case "w06": btn지정할인_Click
        Case "w07": btn특정할인_Click
        Case "w08": btn재세탁_Click
        Case "w09": btn반품_Click
        Case "w10": btn수선_Click
        Case "w11": btn사고품_Click
        Case "w12": btn세탁서비스_Click
        Case "w13": btn미입고_Click
        Case "w14": btn손세탁_Click
        Case "w15": btn아동복_Click
        Case "w16": btn세탁_Click
        Case "w17": btn세탁2_Click (0)
        Case "w18": btn세탁2_Click (1)
        Case "w19": btn세탁수선_Click
        Case "w20": btn재다림질_Click
        Case "w21": btn할증요금_Click
        Case "w22": btn금액_Click
        Case "w23": btn자연수_Click
        Case "w24": btn직원세탁_Click
    End Select

End Sub


Private Sub btn사고품_Click()
    Dim sTemp       As String
    Dim nActRow     As Long
    Dim iPercentage As Double

    ' 사고품
    strFirst = "사"
    strPrice = "0"
    
    nActRow = frm접수.sprGrid.ActiveRow
    
    frm접수.sprGrid.Row = nActRow
    frm접수.sprGrid.Col = 5: sTemp = Trim(frm접수.sprGrid.Text) '
    
    If InStr(1, sTemp, strFirst, vbTextCompare) <= 0 Then
        ' 내용에 "사"자가 없을 경우 "사"을 추가하여 출력 한다.
        frm접수.sprGrid.Text = Trim(frm접수.sprGrid.Text) + strFirst
    End If
    
    frm접수.sprGrid.Col = 6:  frm접수.sprGrid.Text = strPrice                   '금액
    frm접수.sprGrid.Col = 14: frm접수.sprGrid.Text = Replace(strPrice, ",", "") '원금액을 기록한다.
    
    Unload Me
    Load frm작업구분
    
    frm작업구분.SetFlags "사고품 구분"
    frm작업구분.Show
End Sub

Private Sub btn재다림질_Click()
    Dim sTemp       As String
    Dim nActRow     As Long
    Dim iPercentage As Double

    ' 재다림질
    strFirst = "재다"
    strPrice = "0"
    nActRow = frm접수.sprGrid.ActiveRow
    
    frm접수.sprGrid.Row = nActRow
    frm접수.sprGrid.Col = 5: sTemp = Trim(frm접수.sprGrid.Text)
    
    If InStr(1, sTemp, strFirst, vbTextCompare) <= 0 Then
        ' 내용에 "재다"가 없을 경우 "재다"을 추가하여 출력 한다.
        frm접수.sprGrid.Text = Trim(frm접수.sprGrid.Text) + strFirst
    End If
    
    frm접수.sprGrid.Col = 6:  frm접수.sprGrid.Text = strPrice                   '금액
    frm접수.sprGrid.Col = 14: frm접수.sprGrid.Text = Replace(strPrice, ",", "") '원금액을 기록한다.
    
    Unload Me
    
    Load frm작업구분
    frm작업구분.SetFlags "재다림질 구분"
    frm작업구분.Show
End Sub

Private Sub btn반품_Click()
    Dim sTemp   As String
    Dim nActRow As Long
    
    Dim ClothCode As String
    
    strFirst = "반"
    strPrice = "0"
    
    nActRow = frm접수.sprGrid.ActiveRow
    
    frm접수.sprGrid.Row = nActRow
    frm접수.sprGrid.Col = 5: sTemp = Trim(frm접수.sprGrid.Text) '내용
    
    If InStr(1, sTemp, strFirst, vbTextCompare) <= 0 Then
        frm접수.sprGrid.Text = Trim(frm접수.sprGrid.Text) + strFirst ' 내용에 "반"자가 없을 경우 "반"을 추가하여 출력 한다.
    End If
    
    ClothCode = Get_SpreadText(frm접수.sprGrid, Val(nActRow), 8) '의류코드
    
    ' pds2004 수정 2007-05-01
    ' 드반을 입력시 금액을"0"원으로 등록하고 마일리지는 누적 하도록 처리
    
    frm접수.sprGrid.Col = 6:  frm접수.sprGrid.Text = "0" 'Get_DryPrice(ClothCode)
    frm접수.sprGrid.Col = 14: frm접수.sprGrid.Text = "0" '원금액을 기록한다.
    
    Unload Me
    
    Load frm작업구분
    
    frm작업구분.SetFlags "반품 구분"
    frm작업구분.Show
End Sub

'+------------------------------------------------------
'+
'+ 2003/03/12
'+
'+루틴설명
'+  1. 내용에 "하"자가 없을 경우 "하"자를 추가한다.
'+------------------------------------------------------
Private Sub btn하자_Click()
    Dim sTemp   As String
    Dim nActRow As Long
    
    strFirst = "하"
    nActRow = frm접수.sprGrid.ActiveRow
    
    frm접수.sprGrid.Row = nActRow
    frm접수.sprGrid.Col = 5: sTemp = Trim(frm접수.sprGrid.Text) '내용
    
    If InStr(1, sTemp, strFirst, vbTextCompare) <= 0 Then
        frm접수.sprGrid.Text = Trim(frm접수.sprGrid.Text) + strFirst ' 내용에 "하"자가 없을 경우 "하"을 추가하여 출력 한다.
    End If
    
    ' 자신을 종료 한다.
    Me.Hide
    
    Unload Me
    DoEvents
    
    '------------------------------------------
    ' 내용을 입력 받을 준비를 한다.
    '------------------------------------------
    frm접수.sprGrid.SetActiveCell 7, nActRow
    
    frm접수.sprGrid.EditMode = True
End Sub

'--------------------------------------------------------
'
' 수정시 Form1의 fpSpread1_change 도 수정할것
'
'--------------------------------------------------------
Private Sub cmdOK_Click()
    Dim sDate    As String
    Dim sDivi    As String
    Dim sChkAmt  As String
    Dim sChkAmt2 As String
    Dim sSKU     As String
    Dim sTempTag As String
    Dim nActRow  As Integer
    Dim sTemp    As String
    
    nActRow = iCur
    
    ' 아무것도 입력 안할시 종료 처리
    If Len(Trim(pnlCost.Caption)) <= 0 Then
        Unload Me
        
        '-----------------------------------------
        ' 내용을 입력 받을 준비를 한다.
        '-----------------------------------------
        frm접수.sprGrid.SetActiveCell 7, nActRow
        
        frm접수.sprGrid.EditMode = True
        Exit Sub
    End If
    
    If SSPanel1.Visible = True Then
        If Combo1.ListIndex = -1 Then
            MsgBox "할증구분을 선택하십시오..", vbInformation
            Combo1.SetFocus
            Exit Sub
        End If
    End If
    
    If pnlAddGubun.Visible = True Then
        If Combo2.ListIndex = -1 Then
            MsgBox "할증구분을 선택하십시오..", vbInformation
            Combo1.SetFocus
            Exit Sub
        End If
    End If
    
    ' 현재의 택번호를 구한다.
    frm접수.sprGrid.Row = nActRow
    frm접수.sprGrid.Col = 2: sTempTag = frm접수.sprGrid.Text '택번호
    frm접수.sprGrid.Col = 7
   
    If Trim(frm접수.sprGrid.Text) = "짜집기(cm당)" Then
        frm접수.sprGrid.Row = nActRow
        
        If Trim(pnlCost.Caption) = "" Then
            '
        Else
            frm접수.sprGrid.Col = 6: frm접수.sprGrid.Text = Trim(pnlCost.Caption) '금액
        End If
        
        frm작업.Hide
        Unload frm작업
    Else
'        frm접수.sprGrid.Row = nActRow
'        frm접수.sprGrid.Col = 6
'        '입력 금액이 더클경우 바로 출력한다.
'        If CDbl(Format(frm접수.sprGrid.Text, "###0")) < CDbl(Format(IIf(Trim(pnlCost.Caption) = "", "0", Trim(pnlCost.Caption)), "###0")) Then
'           frm접수.sprGrid.Text = Trim(pnlCost.Caption)
'        End If
    End If
        
    frm접수.sprGrid.Row = nActRow
    frm접수.sprGrid.Col = 1: sSKU = frm접수.sprGrid.Text & "" '의류명
    
    sChkAmt2 = pnlCost.Caption
    
    sDate = Format(Date, "YYYY-MM-DD")
    
    ''----------------------------------------------------------
    '' 현재 세일기간 구분
    ''----------------------------------------------------------
    'Query = "SELECT 할인금액 AS 금액 "
    'Query = Query & "FROM TB_할인정보 "
    'Query = Query & " WHERE 시작일자 <= '" & sDate & "' "
    'Query = Query & "AND   종료일자 >= '" & sDate & "' "
    'Query = Query & "AND   의류명 = '" & sSKU & "' "
    'Set ADORs = New ADODB.Recordset
    'ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    '
    'If ADORs.RecordCount < 1 Then
    '    ADORs.Close
    '    Set ADORs = Nothing
    '
    '    If chkDaySale = True Then  '목요세일
    '        Query = "SELECT 금액 "
    '        Query = Query & "FROM TB_목요세일 "
    '        Query = Query & " WHERE 의류명 = '" & sSKU & "' "
    '    Else                       '정상기간
    '        Query = "SELECT 금액 "
    '        Query = Query & "FROM TB_의류 "
    '        Query = Query & " WHERE 의류명 = '" & sSKU & "' "
    '    End If
    '
    '    Set ADORs = New ADODB.Recordset
    '    ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    'End If
            
    ' 기준금액을 그리드에서 가저오게 수정.
    ' 아동복에서 20% 할인과 고가품에서의 3배중 어떤것을 먼저 입력할지 모르기 때문에
    ' 이곳 수정시 아동복 할인 부분도 수정요망
    ' sChkAmt = ADORs!금액
    
    sChkAmt = CLng(Get_SpreadText(frm접수.sprGrid, Val(nActRow), 6))
            
'            ' 드라이수선인 경우 수선금액을 (+)한다.
'            frm접수.sprGrid.Col = 5
'            If frm접수.sprGrid.Text = "드수" Then
''                frm접수.sprGrid.Col = 7
''
''                Query = "Select 금액 "
''                Query = Query & "FROM TB_수선금액 "
''                Query = Query & " WHERE 수선내용 = '" & frm접수.sprGrid.Text & "' "
''
''                Set ADORs = MyDB.OpenRecordset(Query)
''
''                sChkAmt = Str(Val(sChkAmt) + Val(ADORs!금액))
'            ElseIf InStr(1, frm접수.sprGrid.Text, "고") Then
''                '고가 세탁일 경우 기준값을 3배 업한다.
''                If InStr(1, frm접수.sprGrid.Text, "아") Then
''                    ' 아동복일경우 20% 할인금액을 먼저 적용한후 3배를 업한다.
''                    sChkAmt = (Int((CLng(sChkAmt) * 0.8) / 100) * 100)
''                End If
''                sChkAmt = sChkAmt * 3
'            End If
            
    ' 본사에서 확인 코드를 받은 경우
    If Val(sChkAmt) > Val(sChkAmt2) Then
        If Not IsTagNo(chkPricPassWord) Or sTempTag <> chkPricPassWord Then
            MsgBox "규정금액" & " [" & Format(sChkAmt, "#,##0") & "] 보다 입력금액 [" & Format(sChkAmt2, "#,##0") & "]이 작습니다. ", vbInformation
            Exit Sub
        End If
    End If
    
    frm접수.sprGrid.Col = 6:  frm접수.sprGrid.Text = Trim(pnlCost.Caption) '금액
    frm접수.sprGrid.Col = 14: frm접수.sprGrid.Text = Trim(pnlCost.Caption) '정상금액
    
    ' 수선일 경우 수선 금액을 입력 한다.
    If Combo2.ListIndex = 1 Then
        Dim dblOld As Double
        Dim vTemp  As Variant
                        
        '수선
        frm접수.sprGrid.Row = frm접수.sprGrid.ActiveRow
        frm접수.sprGrid.Col = 9: vTemp = frm접수.sprGrid.Text
        
        dblOld = Val(CStr(vTemp))
    
        vTemp = dblOld + Val(sChkAmt2) - Val(sChkAmt)
        
        '수선금액
        frm접수.sprGrid.Row = frm접수.sprGrid.ActiveRow
        frm접수.sprGrid.Col = 9: frm접수.sprGrid.Text = vTemp
    End If
    
    If SSPanel1.Visible = True Then
        strFirst = "할"
        nActRow = frm접수.sprGrid.ActiveRow
        
        frm접수.sprGrid.Row = nActRow
        frm접수.sprGrid.Col = 5: sTemp = Trim(frm접수.sprGrid.Text) '내용
        
        If InStr(1, sTemp, strFirst, vbTextCompare) <= 0 Then
            frm접수.sprGrid.Text = Trim(frm접수.sprGrid.Text) + strFirst ' 내용에 오자가 없을 경우 "오"을 추가하여 출력 한다.
        End If
        
        frm접수.sprGrid.Row = frm접수.sprGrid.ActiveRow  '현재 선택된 줄에 입력 한다. iCur
        frm접수.sprGrid.Col = 7: frm접수.sprGrid.Text = Combo1.Text & " " '상표
    End If
    
    '-------------------------------------------
    ' 내용을 입력 받을 준비를 한다.
    '-------------------------------------------
    frm접수.sprGrid.SetActiveCell 7, nActRow
    
    frm접수.sprGrid.EditMode = True
    
    DoEvents
    SendKeys "{END}"
   
    Unload Me
End Sub

Private Sub cmdAgain_Click()
    pnlCost.Caption = ""
End Sub

Private Sub Command24_Click()
    Dim intlength As Integer
    
    intlength = Len(pnlCost.Caption)
    If intlength = 0 Then Exit Sub
    
    intlength = intlength - 1
    pnlCost.Caption = Mid(pnlCost.Caption, 1, intlength)
End Sub

Private Sub btn세탁수선_Click()
    Dim nActRow As Long
    Dim sTemp  As String
    
    strFirst = "수"
    nActRow = frm접수.sprGrid.ActiveRow
    
    frm접수.sprGrid.Row = nActRow
    frm접수.sprGrid.Col = 5: sTemp = Trim(frm접수.sprGrid.Text) '내용
    
    If InStr(1, sTemp, strFirst, vbTextCompare) <= 0 Then
        frm접수.sprGrid.Text = Trim(frm접수.sprGrid.Text) + strFirst ' 내용에 오자가 없을 경우 "오"을 추가하여 출력 한다.
    End If
    
    ' 수선 내용을 표시한다.
    Unload Me
    
    Load frm세탁수선
    frm세탁수선.Show
    
    '-----------------------------------------
    ' 내용을 입력 받을수 있게 한다.
    '-----------------------------------------
    frm접수.sprGrid.SetActiveCell 7, nActRow
End Sub

'+------------------------------------------------------
'+
'+ 2003/02/21
'+
'+루틴설명
'+  1. 기존의 대수 버튼을 주석 처리하고  급자버튼으로 변경
'+  2. 내용에 "급"자가 없을 경우 "급"자를 추가한다.
'+------------------------------------------------------
Private Sub btn급자_Click()
    Dim sTemp As String
    Dim nActRow As Long

    strFirst = "급"
    nActRow = frm접수.sprGrid.ActiveRow
    
    frm접수.sprGrid.Row = nActRow
    frm접수.sprGrid.Col = 5: sTemp = Trim(frm접수.sprGrid.Text) '내용
    
    If InStr(1, sTemp, strFirst, vbTextCompare) <= 0 Then
        frm접수.sprGrid.Text = Trim(frm접수.sprGrid.Text) + strFirst ' 내용에 급자가 없을 경우 "급"을 추가하여 출력 한다.
    End If
    
    ' 자신을 종료 한다.
    Me.Hide
    Unload Me
    DoEvents
    
    '------------------------------------------
    ' 내용을 입력 받을 준비를 한다.
    '------------------------------------------
    frm접수.sprGrid.SetActiveCell 7, nActRow
    
    frm접수.sprGrid.EditMode = True
End Sub

'+------------------------------------------------------
'+ 2003/03/01
'+
'+루틴설명      - 비밀번호확인
'+  1. 암호를 확인하여 암호 규칙에 맞으면 화면을 종료한다.
'+  2. 레지스터리에 저장한다.
'+
'+------------------------------------------------------
Private Sub cmdHeadOffice_Click()
    Dim strPass As String
    Dim strTag As String
    
    ' 입력 확인
    If Len(txtPassWord.Text) <= 0 Then
        Exit Sub
    End If
    
    frm접수.sprGrid.Col = 2: strTag = frm접수.sprGrid.Text & "" '택번호

'   기본 디폴드 암호.. ( 프로그램 셋팅/설치를 위한 암호 )
    If UCase(txtPassWord.Text) = "DUDTJSGH" Then
        ' 승인된 택번호를 얻어온다.
        chkPricPassWord = strTag
        txtPassWord.Text = ""
        
        MsgBox "규정 금액보다 적게 입력할 수 있게 확인 되었습니다.", vbInformation, "확인"
        Exit Sub
    End If
    
    strPass = IsPricPassWord(txtPassWord.Text, strTag) ' 비밀번호 확인
    
    If strPass = "-1" Or strPass = "-3" Then
        txtPassWord.SelStart = 0
        txtPassWord.SelLength = Len(txtPassWord.Text)
        DoEvents
        
        chkPricPassWord = ""
        
        If strPass = "-3" Then
            MsgBox "입력값이 올바르지 않습니다.", vbInformation, "입력오류"
        End If
        
        txtPassWord.Text = ""
        txtPassWord.SetFocus
        
        Exit Sub
    Else
        chkPricPassWord = strTag
        txtPassWord.Text = ""
        
        MsgBox "규정 금액보다 적게 입력할 수 있게 확인 되었습니다.", vbInformation, "확인"
        
        Exit Sub
    End If
End Sub

'+------------------------------------------------------
'+
'+ 2003/08/29
'+
'+루틴설명
'+  1. 행사 품목을 추가했다.
'+  2. 행사기간은 비밀번호로 입력하고.
'+  3. 해당 기간의 금액은 0원이다.
'+
'+------------------------------------------------------
Private Sub btn행사품목_Click()
    Dim sTemp   As String
    Dim nActRow As Integer
    
    '행사기간 여부를 확인한다.
    If IsEventPassREGRead > 0 Then
        chkEventSale = True
    ElseIf Not chkEventSale Then
        DoEvents
        frm행사코드.Show 1
    End If
    
    If Not chkEventSale Then
        DoEvents
'        MsgBox "행사 기간에만 사용하실수 있습니다.", vbInformation, "확인"
        Exit Sub
    End If
    
    
    strFirst = "행"
    strPrice = "0"
    
    nActRow = frm접수.sprGrid.ActiveRow
    
    frm접수.sprGrid.Row = nActRow
    frm접수.sprGrid.Col = 5: sTemp = Trim(frm접수.sprGrid.Text)
    
    If InStr(1, sTemp, strFirst, vbTextCompare) <= 0 Then
        ' 내용에 오자가 없을 경우 "재"을 추가하여 출력 한다.
        frm접수.sprGrid.Text = Trim(frm접수.sprGrid.Text) + strFirst
    End If
    
    frm접수.sprGrid.Col = 6:  frm접수.sprGrid.Text = strPrice                   '금액
    frm접수.sprGrid.Col = 14: frm접수.sprGrid.Text = Replace(strPrice, ",", "") '원금액을 기록한다.
    
    Unload Me
    
    '-------------------------------------------
    ' 내용을 입력 받을 준비를 한다.
    '-------------------------------------------
    frm접수.sprGrid.SetActiveCell 7, nActRow
    
    frm접수.sprGrid.EditMode = True
End Sub

Private Sub btn세탁서비스_Click()
    Dim sTemp As String
    Dim nActRow As Integer
    
    '세탁 서비스
    If Trim(chkServicePassWord) <> frm접수.txtCode Then
        chkServicePassWord = ""
        frm서비스코드.Show 1
        
        If Len(Trim(chkServicePassWord)) <= 0 Then
            DoEvents
            Exit Sub
        End If
    End If
    
    strFirst = "서"
    strPrice = "0"
    
    nActRow = frm접수.sprGrid.ActiveRow
    
    frm접수.sprGrid.Row = nActRow
    frm접수.sprGrid.Col = 5: sTemp = Trim(frm접수.sprGrid.Text)
    
    If InStr(1, sTemp, strFirst, vbTextCompare) <= 0 Then
        ' 내용에 "서"자가 없을 경우 "서"을 추가하여 출력 한다.
        frm접수.sprGrid.Text = Trim(frm접수.sprGrid.Text) + strFirst
    End If
    
    frm접수.sprGrid.Col = 6:  frm접수.sprGrid.Text = strPrice                   '금액
    frm접수.sprGrid.Col = 14: frm접수.sprGrid.Text = Replace(strPrice, ",", "") '원금액을 기록한다.
    
    Unload Me
    DoEvents
    
    '------------------------------------------
    ' 내용을 입력 받을 준비를 한다.
    '------------------------------------------
    frm접수.sprGrid.SetActiveCell 7, nActRow
    
    frm접수.sprGrid.EditMode = True
End Sub

'+------------------------------------------------------
'+
'+ 2003/02/21
'+
'+루틴설명
'+  1. 기존의 단체복 버튼을 주석 처리하고  사고품버튼으로 변경
'+
'+------------------------------------------------------
Private Sub btn자연수_Click()
    Dim sTemp       As String
    Dim nActRow     As Long
    Dim iPercentage As Double

    ' 자연(水)
    strFirst = "水"
    
    '2013-09-01일부터 자연수를 130%로 정함
    ' 기타 일반 제품은 가격조정을 하였음.
    If Format(Date, "yyyy-MM-dd") >= "2013-09-01" Then
        iPercentage = 130 / 100  '
    Else
        iPercentage = 160 / 100  ' (할인이 20%일 경우 0.8의 값을 같는다.)
    End If
    
    nActRow = frm접수.sprGrid.ActiveRow
    strPrice = CLng(Get_SpreadText(frm접수.sprGrid, nActRow, 6)) '금액
    
    frm접수.sprGrid.Row = nActRow
    frm접수.sprGrid.Col = 5: sTemp = Trim(frm접수.sprGrid.Text) '내용
    
    If InStr(1, sTemp, strFirst, vbTextCompare) <= 0 Then
        ' 내용에 "水" 없을 경우 "水"을 추가하여 출력 한다.
        frm접수.sprGrid.Text = Trim(frm접수.sprGrid.Text) + strFirst
        
        strPrice = Str(Val(strPrice) * iPercentage)
        
        strPrice = GetNumber500UP(CDbl(strPrice))
    End If
    
    frm접수.sprGrid.Col = 6:  frm접수.sprGrid.Text = strPrice                   '금액
    frm접수.sprGrid.Col = 14: frm접수.sprGrid.Text = Replace(strPrice, ",", "") '원금액을 기록한다.
    
    Unload Me
    
    '-----------------------------------------
    ' 내용을 입력 받을 준비를 한다.
    '-----------------------------------------
    frm접수.sprGrid.SetActiveCell 7, nActRow
    
    frm접수.sprGrid.EditMode = True
End Sub

'+------------------------------------------------------
'+
'+ 2003/02/21
'+
'+루틴설명
'+  1. 기존의 단체복 버튼을 주석 처리하고  사고품버튼으로 변경
'+
'+------------------------------------------------------
Private Sub btn직원세탁_Click()
    Dim nActRow        As Long
    Dim logPrice       As Long
    Dim strChildPrice  As String
    Dim strChildPrice2 As String
    Dim sTemp          As String
    Dim iPercentage    As Double
    Dim nTempMoney     As Double
    
    Dim sGroupActionCode As String
    
    nActRow = frm접수.sprGrid.Row
    strFirst = "직"
    
    iPercentage = (100 - 30) / 100  ' (할인이 30%일 경우 0.7의 값을 갇는다.)
    
    ' 해당 금액을 얻어온다.
    ' 고가품에서 3배 해주는것과 연관해서 어느쪽을 먼저 할지 모르기 때문에
    'logPrice = Get_DryPrice(Get_SpreadText(frm접수.sprGrid, nActRow, 7))
    
    logPrice = CLng(Get_SpreadText(frm접수.sprGrid, nActRow, 6)) '금액
    
    
    frm접수.sprGrid.Row = nActRow
    frm접수.sprGrid.Col = 5: sTemp = Trim(frm접수.sprGrid.Text) '내용
    
    If InStr(1, sTemp, strFirst, vbTextCompare) <= 0 Then
        frm접수.sprGrid.Col = 5: frm접수.sprGrid.Text = Mid(sTemp, 1, 1) & strFirst & Mid(sTemp, 2, Len(sTemp)) ' 내용에 "할"자가 없을 경우 "할"을 추가하여 출력 한다.
        frm접수.sprGrid.Col = 6: sTemp = Trim(frm접수.sprGrid.Text)                                             '금액을 출력한다.
        
        If Val(Format(sTemp, "#")) > Val(logPrice) Then
            strChildPrice2 = Val(Format(sTemp, "#")) - Val(logPrice)                ' 수선이나 기타 "드라이"에서 추가된 부분이 있을 경우 아동복은 드라이에서만 할인한다.
            strChildPrice = CStr(Int((CDbl((logPrice) * iPercentage) / 100)) * 100) ' 10원단위를 절사 한다.
            
            frm접수.sprGrid.Col = 6:  frm접수.sprGrid.Text = Val(strChildPrice) + Val(strChildPrice2)
            frm접수.sprGrid.Col = 14: frm접수.sprGrid.Text = Val(strChildPrice) + Val(strChildPrice2) ' 원금액을 기록한다.
            
        ElseIf sTemp = Format(logPrice, "#,##0") Then
            ' 기본에서 20%로 할인한다.
            strChildPrice = CStr(Int(CDbl((logPrice * iPercentage) / 100)) * 100)            ' 10원단위를 절사 한다.
            
            frm접수.sprGrid.Col = 6:  frm접수.sprGrid.Text = strChildPrice                   '
            frm접수.sprGrid.Col = 14: frm접수.sprGrid.Text = Replace(strChildPrice, ",", "") ' 원금액을 기록한다.
        End If
    End If
            
    DoEvents
    Me.Hide
    Unload Me
    
    ' 내용을 입력 받을 준비를 한다.
    frm접수.sprGrid.SetActiveCell 7, iCur
    
    frm접수.sprGrid.EditMode = True
End Sub

'+------------------------------------------------------
'+
'+ 2003/02/21
'+
'+루틴설명
'+  1. 기존의 단체복 버튼을 주석 처리하고  사고품버튼으로 변경
'+
'+------------------------------------------------------
Private Sub btn고가세탁_Click()
    Dim sTemp       As String
    Dim nActRow     As Long
    Dim iPercentage As Double

    ' 고가세탁
    strFirst = "고"
    iPercentage = Val(가맹점정보.고가세탁비율) / 100  ' (할인이 20%일 경우 0.8의 값을 같는다.)
    
    nActRow = frm접수.sprGrid.ActiveRow
    strPrice = CLng(Get_SpreadText(frm접수.sprGrid, nActRow, 6)) '금액
    
    frm접수.sprGrid.Row = nActRow
    frm접수.sprGrid.Col = 5: sTemp = Trim(frm접수.sprGrid.Text) '내용
    
    If InStr(1, sTemp, strFirst, vbTextCompare) <= 0 Then
        ' 내용에 "고" 없을 경우 "고"을 추가하여 출력 한다.
        frm접수.sprGrid.Text = Trim(frm접수.sprGrid.Text) + strFirst
        
        strPrice = Str(Val(strPrice) * iPercentage)
        
        strPrice = CStr(Int(CDbl(Val(strPrice) / 100)) * 100) ' 10원 단위 절사
    End If
    
    frm접수.sprGrid.Col = 6:  frm접수.sprGrid.Text = strPrice                   '금액
    frm접수.sprGrid.Col = 14: frm접수.sprGrid.Text = Replace(strPrice, ",", "") '원금액을 기록한다.
    
    Unload Me
    
    '-----------------------------------------
    ' 내용을 입력 받을 준비를 한다.
    '-----------------------------------------
    frm접수.sprGrid.SetActiveCell 7, nActRow
    
    frm접수.sprGrid.EditMode = True
End Sub

Private Sub btn할증요금_Click()
    intMode = 1 '할증 구분 콤보를 보이기 위하여
    
    '-----------------------------------------------------------
    ' Active Cell
    '-----------------------------------------------------------
    frm접수.sprGrid.SetActiveCell 7, frm접수.sprGrid.ActiveRow

    TabControl.SelectedItem = 1
    
    SSPanel1.Visible = True
        
    'cmdOK.SetFocus
End Sub

Private Sub btn미입고_Click()
    Dim nActRow As Integer
    Dim sTemp    As String
   
    Dim ClothCode As String
   
    strFirst = "미"
    nActRow = frm접수.sprGrid.ActiveRow
    
    frm접수.sprGrid.Row = nActRow
    frm접수.sprGrid.Col = 5: sTemp = Trim(frm접수.sprGrid.Text) & "" '내용
    
    If InStr(1, sTemp, strFirst, vbTextCompare) <= 0 Then
        ' 내용에 "미"자가 없을 경우 "미"을 추가하여 출력 한다.
        frm접수.sprGrid.Text = Trim(frm접수.sprGrid.Text) + strFirst
    End If
    
    ' 금액 출력
    ClothCode = Get_SpreadText(frm접수.sprGrid, Val(nActRow), 8)                                '의류코드
    
    frm접수.sprGrid.Col = 6:  frm접수.sprGrid.Text = Get_DryPrice(ClothCode)                   '금액
    frm접수.sprGrid.Col = 14: frm접수.sprGrid.Text = Replace(Get_DryPrice(ClothCode), ",", "") '원금액을 기록한다.
    
    ' 자신을 종료 한다.
    Me.Hide
    Unload Me
    DoEvents
    
    '-----------------------------------------
    ' 내용을 입력 받을 준비를 한다.
    '-----------------------------------------
    frm접수.sprGrid.SetActiveCell 7, nActRow
    
    frm접수.sprGrid.EditMode = True
End Sub

'+------------------------------------------------------
'+
'+ 2003/03/12
'+
'+루틴설명
'+  1. "특정 할인"을 클릭하였을겨우.
'+  2. "지"자를 "드"자 다음에 출력한다.
'+  3. "수"(대수) 내용및 기타 추가 금액이 있을 경우 "드"(드라이)에서만 20% 할인한다.
'+------------------------------------------------------
Private Sub btn지정할인_Click()
    Dim nActRow        As Long
    Dim logPrice       As Long
    Dim strChildPrice  As String
    Dim strChildPrice2 As String
    Dim sTemp          As String
    Dim iPercentage    As Double
    Dim nTempMoney     As Double
    Dim varTemp        As Variant
    
    If 가맹점정보.지정할인여부 <> "Y" Then Exit Sub
    
    nActRow = frm접수.sprGrid.Row
    
    If 가맹점정보.지사코드 <> M_COUPON_KLENZ_CODE Then
        ' 이마트 일 경우만 코드 확인을 한다.
        If InStr(가맹점정보.가맹점명, "이마트") > 0 Then
            '+------------------------------------------------------
            ' 2011-10-27 ~ 2011-11-30 일 까지 특정 품목이 입력 되어 있어야 동작 하도록 설정
            If Format(Date, "yyyyMMdd") >= "20111027" And Format(Date, "yyyyMMdd") <= "20111130" Then
                
                Call frm접수.sprGrid.GetText(8, nActRow, varTemp)
                Select Case Action_지정할인_코드확인(CStr(varTemp))
                    Case "Z"
                        sTemp = "행사 코드가 등록 되어 있어야 지정할인을 할 수 있습니다." & vbNewLine & vbNewLine
                        sTemp = sTemp & "지정된 상품 코드를 확인하여 주십시요."
                        MsgBox sTemp, vbCritical, "지정할인 오류"
                        Exit Sub
                    
                    Case "A"
                        sTemp = "할인 대상 품목이 아닙니다." & vbNewLine & vbNewLine
                        sTemp = sTemp & "지정된 할인 품목 코드를 확인하여 주십시요."
                        MsgBox sTemp, vbCritical, "지정할인 오류"
                        Exit Sub
                        
                    Case Else
                
                End Select
                
                ' 11 행사일 경우에 실행 한다.
                frm접수.sprGrid.Col = 7
                frm접수.sprGrid.Text = "E행사 " & Trim(frm접수.sprGrid.Text)
            
            End If
            '+------------------------------------------------------
        End If
    End If
    
    strFirst = "지"
    iPercentage = (100 - 가맹점정보.지정할인비율) / 100  ' (할인이 20%일 경우 0.8의 값을 같는다.)
    
    ' 해당 금액을 얻어온다.
    ' 고가품에서 3배 해주는것과 연관해서 어느쪽을 먼저 할지 모르기 때문에
    'logPrice = Get_DryPrice(Get_SpreadText(frm접수.sprGrid, nActRow, 7))
    logPrice = CLng(Get_SpreadText(frm접수.sprGrid, nActRow, 6))
    
    frm접수.sprGrid.Row = nActRow
    frm접수.sprGrid.Col = 5: sTemp = Trim(frm접수.sprGrid.Text) & "" '내용
    
    If InStr(1, sTemp, strFirst, vbTextCompare) <= 0 Then
        ' 내용에 "할"자가 없을 경우 "할"을 추가하여 출력 한다.
        frm접수.sprGrid.Col = 5: frm접수.sprGrid.Text = Mid(sTemp, 1, 1) & strFirst & Mid(sTemp, 2, Len(sTemp)) '내용
        frm접수.sprGrid.Col = 6: sTemp = Trim(frm접수.sprGrid.Text)                                             '금액을 출력한다.
        
        If Val(Format(sTemp, "#")) > Val(logPrice) Then
            ' 수선이나 기타 "드라이"에서 추가된 부분이 있을 경우 아동복은 드라이에서만 할인한다.
            strChildPrice2 = Val(Format(sTemp, "#")) - Val(logPrice)
            
            strChildPrice = CStr(Int((CDbl((logPrice) * iPercentage) / 100)) * 100)                   '10원단위를 절사 한다.
            
            frm접수.sprGrid.Col = 6:  frm접수.sprGrid.Text = Val(strChildPrice) + Val(strChildPrice2) '금액
            frm접수.sprGrid.Col = 14: frm접수.sprGrid.Text = Replace(strPrice, ",", "")               '원금액을 기록한다.
        
        ElseIf sTemp = Format(logPrice, "#,##0") Then
            ' 기본에서 20%로 할인한다.
            
            ' 10원단위를 절사 한다.
            strChildPrice = CStr(Int(CDbl((logPrice * iPercentage) / 100)) * 100)
            
            frm접수.sprGrid.Col = 6:  frm접수.sprGrid.Text = strChildPrice                            '
            frm접수.sprGrid.Col = 14: frm접수.sprGrid.Text = Val(strChildPrice) + Val(strChildPrice2) ' 원금액을 기록한다.
        End If
    End If
            
    DoEvents
    Me.Hide
    Unload Me
    
    '-----------------------------------------
    ' 내용을 입력 받을 준비를 한다.
    '-----------------------------------------
    frm접수.sprGrid.SetActiveCell 7, iCur
    
    frm접수.sprGrid.EditMode = True
End Sub

'+------------------------------------------------------
'+
'+ 건식, 습식
'+
'+루틴설명
'+  1. "건" 모든 것을 초기화 하고 "드"로 설정한다.
'+------------------------------------------------------
Private Sub btn세탁2_Click(Index As Integer)
    Dim nActRow   As Integer
    Dim ClothCode As String
    
    strFirst = IIf(Index = 0, "건", "습")
    nActRow = frm접수.sprGrid.ActiveRow
    
    ' 내용 출력
    frm접수.sprGrid.Row = nActRow
    frm접수.sprGrid.Col = 5: frm접수.sprGrid.Text = strFirst                      '내용
    
    ClothCode = Get_SpreadText(frm접수.sprGrid, Val(nActRow), 8)                  '금액 출력
    
    frm접수.sprGrid.Col = 6:  frm접수.sprGrid.Text = Get_DryPrice(ClothCode) & "" '금액
    frm접수.sprGrid.Col = 14: frm접수.sprGrid.Text = Get_DryPrice(ClothCode) & "" '원금액을 기록한다.
    frm접수.sprGrid.Col = 9:  frm접수.sprGrid.Text = "0"                          '수선금액을 초기화 한다.
    
    ' 자신을 종료 한다.
    Me.Hide
    Unload Me
    DoEvents
    
    '-------------------------------------
    ' 내용을 입력 받을 준비를 한다.
    '-------------------------------------
    frm접수.sprGrid.SetActiveCell 7, nActRow
    
    frm접수.sprGrid.EditMode = True
End Sub

'+------------------------------------------------------
'+
'+ 2003/03/12
'+
'+루틴설명
'+  1. "특정 할인"을 클릭하였을겨우.
'+  2. "지"자를 "드"자 다음에 출력한다.
'+  3. "수"(대수) 내용및 기타 추가 금액이 있을 경우 "드"(드라이)에서만 20% 할인한다.
'+------------------------------------------------------
Private Sub btn특정할인_Click()
    Dim nActRow        As Long
    Dim logPrice       As Long
    Dim strChildPrice  As String
    Dim strChildPrice2 As String
    Dim sTemp          As String
    Dim iPercentage    As Double
    Dim nTempMoney     As Double
    
    Dim sGroupActionCode As String
       
    If 가맹점정보.특정할인여부 <> "Y" Then Exit Sub
    
    nActRow = frm접수.sprGrid.Row
    strFirst = "특"
    
    iPercentage = (100 - 가맹점정보.특정할인비율) / 100  ' (할인이 20%일 경우 0.8의 값을 갇는다.)
    
    ' 해당 금액을 얻어온다.
    ' 고가품에서 3배 해주는것과 연관해서 어느쪽을 먼저 할지 모르기 때문에
    'logPrice = Get_DryPrice(Get_SpreadText(frm접수.sprGrid, nActRow, 7))
    
    logPrice = CLng(Get_SpreadText(frm접수.sprGrid, nActRow, 6)) '금액
    
    ' 세트상품의 특정 할인을 확인한다.
''    If 가맹점정보.지사코드 <> M_COUPON_KLENZ_CODE Then
''        If Format(Date, "YYYY-MM-DD") >= "20091211" And Format(Date, "YYYY-MM-DD") <= "20100131" Then
''            frm접수.sprGrid.Row = nActRow
''            frm접수.sprGrid.Col = 8: sTemp = Trim(frm접수.sprGrid.Text) '코드
''
''            sGroupActionCode = "m00,m01"
''
''            If InStr(sGroupActionCode, sTemp) <= 0 Then
''                MsgBox "세트 상품 할인은 와이셔츠[" & sGroupActionCode & "]만 가능 합니다.", vbInformation, "확인"
''                Exit Sub
''            End If
''        End If
''    End If
    
    frm접수.sprGrid.Row = nActRow
    frm접수.sprGrid.Col = 5: sTemp = Trim(frm접수.sprGrid.Text) '내용
    
    If InStr(1, sTemp, strFirst, vbTextCompare) <= 0 Then
        frm접수.sprGrid.Col = 5: frm접수.sprGrid.Text = Mid(sTemp, 1, 1) & strFirst & Mid(sTemp, 2, Len(sTemp)) ' 내용에 "할"자가 없을 경우 "할"을 추가하여 출력 한다.
        frm접수.sprGrid.Col = 6: sTemp = Trim(frm접수.sprGrid.Text)                                             '금액을 출력한다.
        
        If Val(Format(sTemp, "#")) > Val(logPrice) Then
            strChildPrice2 = Val(Format(sTemp, "#")) - Val(logPrice)                ' 수선이나 기타 "드라이"에서 추가된 부분이 있을 경우 아동복은 드라이에서만 할인한다.
            strChildPrice = CStr(Int((CDbl((logPrice) * iPercentage) / 100)) * 100) ' 10원단위를 절사 한다.
            
            frm접수.sprGrid.Col = 6:  frm접수.sprGrid.Text = Val(strChildPrice) + Val(strChildPrice2)
            frm접수.sprGrid.Col = 14: frm접수.sprGrid.Text = Val(strChildPrice) + Val(strChildPrice2) ' 원금액을 기록한다.
            
        ElseIf sTemp = Format(logPrice, "#,##0") Then
            ' 기본에서 20%로 할인한다.
            strChildPrice = CStr(Int(CDbl((logPrice * iPercentage) / 100)) * 100)            ' 10원단위를 절사 한다.
            
            frm접수.sprGrid.Col = 6:  frm접수.sprGrid.Text = strChildPrice                   '
            frm접수.sprGrid.Col = 14: frm접수.sprGrid.Text = Replace(strChildPrice, ",", "") ' 원금액을 기록한다.
        End If
    End If
            
    DoEvents
    Me.Hide
    Unload Me
    
    ' 내용을 입력 받을 준비를 한다.
    frm접수.sprGrid.SetActiveCell 7, iCur
    
    frm접수.sprGrid.EditMode = True
End Sub

'+------------------------------------------------------
'+
'+ 2003/03/12
'+
'+루틴설명
'+  1. "손" 모든 것을 초기화 하고 "손"로 설정한다.
'+------------------------------------------------------
Private Sub btn손세탁_Click()
   Dim nActRow As Integer
   Dim strPrice As String

    strFirst = "손"
    nActRow = frm접수.sprGrid.ActiveRow
    
    frm접수.sprGrid.Row = nActRow
    frm접수.sprGrid.Col = 5: frm접수.sprGrid.Text = strFirst   '내용
    
    strPrice = Get_SpreadText(frm접수.sprGrid, Val(nActRow), 8) '금액 출력
    
    frm접수.sprGrid.Col = 6:  frm접수.sprGrid.Text = (Get_DryPrice(strPrice) * 2)                   '금액
    frm접수.sprGrid.Col = 14: frm접수.sprGrid.Text = Replace((Get_DryPrice(strPrice) * 2), ",", "") '원금액을 기록한다.
    frm접수.sprGrid.Col = 9:  frm접수.sprGrid.Text = "0"                                            '수선금액을 초기화 한다.
    
    ' 자신을 종료 한다.
    Me.Hide
    Unload Me
    DoEvents
    
    '-----------------------------------------
    ' 내용을 입력 받을 준비를 한다.
    '-----------------------------------------
    frm접수.sprGrid.SetActiveCell 7, nActRow
    
    frm접수.sprGrid.EditMode = True
End Sub

'------------------------------------------------------
' "수" - 한자로서 수선만을 표현
'------------------------------------------------------
Private Sub btn수선_Click()
    Dim nActRow As Long
    
    nActRow = frm접수.sprGrid.ActiveRow
    
    frm접수.sprGrid.Row = nActRow
    frm접수.sprGrid.Col = 5: frm접수.sprGrid.Value = "수"
    
    Unload Me
    
    If strSu = "2" And strJa = "1" Then
       Load frm수선
       frm수선.Show
    Else
       Load frm세탁수선
       frm세탁수선.Show
    End If
End Sub

Private Sub btn재세탁_Click()
    Dim sTemp   As String
    Dim nActRow As Integer
    
    strFirst = "재"
    strPrice = "0"
    nActRow = frm접수.sprGrid.ActiveRow
    
    frm접수.sprGrid.Row = nActRow
    frm접수.sprGrid.Col = 5: sTemp = Trim(frm접수.sprGrid.Text) '내용
    
    If InStr(1, sTemp, strFirst, vbTextCompare) <= 0 Then
        ' 내용에 오자가 없을 경우 "재"을 추가하여 출력 한다.
        frm접수.sprGrid.Text = Trim(frm접수.sprGrid.Text) + strFirst
    End If
    
    frm접수.sprGrid.Col = 6:  frm접수.sprGrid.Text = strPrice
    frm접수.sprGrid.Col = 14: frm접수.sprGrid.Text = Replace(strPrice, ",", "") ' 원금액을 기록한다.
    
    Unload Me
    
    Load frm작업구분
    frm작업구분.SetFlags "재세탁 구분"
    frm작업구분.Show
    
End Sub

'+------------------------------------------------------
'+
'+ 2003/03/12
'+
'+루틴설명
'+  1. "드" 모든 것을 초기화 하고 "드"로 설정한다.
'+------------------------------------------------------
Private Sub btn세탁_Click()
    Dim nActRow As Integer
    Dim strPrice As String
   
    strFirst = "세"
    nActRow = frm접수.sprGrid.ActiveRow
    
    ' 내용 출력
    frm접수.sprGrid.Row = nActRow
    frm접수.sprGrid.Col = 5: frm접수.sprGrid.Text = strFirst                                  '내용
    
    strPrice = Get_SpreadText(frm접수.sprGrid, Val(nActRow), 8)                                '의류코드
    
    frm접수.sprGrid.Col = 6:  frm접수.sprGrid.Text = Get_DryPrice(strPrice)                   '금액 출력
    frm접수.sprGrid.Col = 14: frm접수.sprGrid.Text = Replace(Get_DryPrice(strPrice), ",", "") '원금액을 기록한다.
    frm접수.sprGrid.Col = 9:  frm접수.sprGrid.Text = "0"                                      '수선금액을 초기화 한다.
    
    ' 자신을 종료 한다.
    Me.Hide
    Unload Me
    DoEvents
    
    '------------------------------------------
    ' 내용을 입력 받을 준비를 한다.
    '------------------------------------------
    frm접수.sprGrid.SetActiveCell 7, nActRow
    
    frm접수.sprGrid.EditMode = True
End Sub

'+------------------------------------------------------
'+
'+ 2003/03/12
'+
'+루틴설명
'+  1. "아동복"을 클릭하였을겨우.
'+  2. "아"자를 "드"자 다음에 출력한다.
'+  3. "수"(대수) 내용및 기타 추가 금액이 있을 경우 "드"(드라이)에서만 20% 할인한다.
'+------------------------------------------------------
Private Sub btn아동복_Click()
    Dim nActRow        As Long
    Dim logPrice       As Long
    Dim strChildPrice  As String
    Dim strChildPrice2 As String
    Dim sTemp          As String
       
    nActRow = frm접수.sprGrid.Row
    strFirst = "아"
    
    ' 해당 금액을 얻어온다. (20% 할인해주기 위하여)
    ' 2003/4/11일 금액을 그리드에서 얻어오게 수정.
    ' 고가품에서 3배 해주는것과 연관해서 어느쪽을 먼저 할지 모르기 때문에
    'logPrice = Get_DryPrice(Get_SpreadText(frm접수.sprGrid, nActRow, 7))
    logPrice = CLng(Get_SpreadText(frm접수.sprGrid, nActRow, 20))
    
    frm접수.sprGrid.Row = nActRow
    frm접수.sprGrid.Col = 5: sTemp = Trim(frm접수.sprGrid.Text)
    
    If InStr(1, sTemp, strFirst, vbTextCompare) <= 0 Then
        Dim TempValue As Variant
        Dim TempValue2 As Variant
        Call frm접수.sprGrid.GetText(20, nActRow, TempValue)
        Call frm접수.sprGrid.GetText(6, nActRow, TempValue2)
        'If TempValue = TempValue2 Then
        
            
            
            ' 내용에 "아"자가 없을 경우 "아"을 추가하여 출력 한다.
            frm접수.sprGrid.Col = 5: frm접수.sprGrid.Text = Mid(sTemp, 1, 1) & strFirst & Mid(sTemp, 2, Len(sTemp))
            frm접수.sprGrid.Col = 6: sTemp = Trim(frm접수.sprGrid.Text) '금액을 출력한다.
            
            If Val(Format(sTemp, "#")) > Val(logPrice) Then
            ' 수선이나 기타 "드라이"에서 추가된 부분이 있을 경우 아동복은 드라이에서만 할인한다.
'                strChildPrice2 = Val(Format(sTemp, "#")) - Val(logPrice)
'
'                ' 10원단위를 절사 한다.
'                strChildPrice = CStr(Int((CLng(logPrice) * 0.8) / 100) * 100)
'
'                frm접수.sprGrid.Col = 6:  frm접수.sprGrid.Text = Val(strChildPrice) + Val(strChildPrice2)
'                frm접수.sprGrid.Col = 14: frm접수.sprGrid.Text = Replace(Val(strChildPrice) + Val(strChildPrice2), ",", "") ' 원금액을 기록한다.
'
                strChildPrice = CStr(Int((CLng(sTemp) * 0.8) / 100) * 100)
                
                frm접수.sprGrid.Col = 6:  frm접수.sprGrid.Text = strChildPrice                   '금액
                frm접수.sprGrid.Col = 14: frm접수.sprGrid.Text = Replace(strChildPrice, ",", "") '원금액을 기록한다.
            Else
            ' 기본에서 20%로 할인한다.
                ' 10원단위를 절사 한다.
                strChildPrice = CStr(Int((CLng(logPrice) * 0.8) / 100) * 100)
                
                frm접수.sprGrid.Col = 6:  frm접수.sprGrid.Text = strChildPrice                   '금액
                frm접수.sprGrid.Col = 14: frm접수.sprGrid.Text = Replace(strChildPrice, ",", "") '원금액을 기록한다.
            End If
'        Else
'            Call MsgBox("할인적용 품목에 누적할인은 불가능합니다.")
'        End If
    End If
            
    DoEvents
    Me.Hide
    Unload Me
    
    ' 내용을 입력 받을 준비를 한다.
    frm접수.sprGrid.SetActiveCell 7, iCur
    
    frm접수.sprGrid.EditMode = True
End Sub

Private Sub btn금액_Click()
    frm접수.sprGrid.SetActiveCell 7, frm접수.sprGrid.ActiveRow
    
    TabControl.SelectedItem = 1
    DoEvents
    
    SSPanel1.Visible = False
    
''    pnlMoney.Visible = True
''    pnlMoney.ZOrder 0
''
''    'Me.Width = pnlMoney.Width
''    Me.Width = 12180
''
''    cmdOK.SetFocus
End Sub

'+------------------------------------------------------
'+
'+ 2003/03/12
'+
'+루틴설명
'+  1. 내용에 "오"자가 없을 경우 "오"자를 추가한다.
'+------------------------------------------------------
Private Sub btn오염_Click()
    Dim sTemp   As String
    Dim nActRow As Long

    strFirst = "오"
    nActRow = frm접수.sprGrid.ActiveRow
    
    frm접수.sprGrid.Row = nActRow
    frm접수.sprGrid.Col = 5: sTemp = Trim(frm접수.sprGrid.Text)
    
    If InStr(1, sTemp, strFirst, vbTextCompare) <= 0 Then
        frm접수.sprGrid.Text = Trim(frm접수.sprGrid.Text) + strFirst ' 내용에 "오"자가 없을 경우 "오"을 추가하여 출력 한다.
    End If

    ' 자신을 종료 한다.
    Me.Hide
    Unload Me
    DoEvents
    
    '--------------------------------------
    ' 내용을 입력 받을 준비를 한다.
    '--------------------------------------
    frm접수.sprGrid.SetActiveCell 7, nActRow
    
    frm접수.sprGrid.EditMode = True
End Sub

'+------------------------------------------------------
'+
'+ 2007/02/28
'+
'+루틴설명
'+  1. E-Mart에서 신규 S.Point 카드를 발급하여 이카드를 이용하여 결재한 경우 10% 할인해주는 내용
'+  2. 할인 품목 정장상의(f00), 정장하의, 스커트
'+  3. 행사 기간 : 2007-03-01 ~ 2007-03-07
'+  4. 적용 이마트
'       일산지사(1004)      043:은평, 005:신월
'       천안유니트(1007)    021:평택, 022:서수원
'       수지유니트(1006)    045:구성
'       춘천지사(1002)      044:원주
'       인천지사(1003)      322:동천
'       안산지사(1005)      011:고잔
'       경산지사(1001)      355:해운대, 245:연재, 015:만촌, 223:월배, 038:칠성, 042:비산, 205:구미, 234:학성, 141:경산
'+------------------------------------------------------
Private Sub Form_Activate()
    On Error GoTo 0
     
    DoEvents
    
''    If Format(Date, "YYYY-MM-DD") >= "20070301" And Format(Date, "YYYY-MM-DD") <= "20070307" Then
''        If CheckSPointCard = True Then
''            btn지정할인.Enabled = True
''            가맹점정보.지정할인비율 = 10
''            가맹점정보.지정할인여부 = "Y"
''
''        ElseIf 가맹점정보.지정할인여부 <> "Y" Then
''            btn지정할인.Enabled = False
''        End If
''
''    ElseIf 가맹점정보.지정할인여부 <> "Y" Then
''        btn지정할인.Enabled = False
''    End If
''
'''+------------------------------------------------------
'''+
'''+ 2007/04/06
'''+
'''+루틴설명
'''+  1. E-Mart에서 지정할인  이용하여 결재한 경우 30% 할인해주는 내용
'''+  2. 할인 품목 정장상의(f00), 정장하의, 스커트
'''+  3. 행사 기간 : 2007-04-06 ~ 2007-04-08
'''+  4. 적용 이마트
'''       경산지사(1001)      355:해운대
'''+------------------------------------------------------
''    If Format(Date, "YYYY-MM-DD") >= "20070406" And Format(Date, "YYYY-MM-DD") <= "20070408" Then
''        If Check_지정할인_20070406 = True Then
''            btn지정할인.Enabled = True
''            가맹점정보.지정할인비율 = 30
''            가맹점정보.지정할인여부 = "Y"
''
''        ElseIf 가맹점정보.지정할인여부 <> "Y" Then
''            btn지정할인.Enabled = False
''        End If
''
''    ElseIf 가맹점정보.지정할인여부 <> "Y" Then
''        btn지정할인.Enabled = False
''    End If
    
'''+------------------------------------------------------
'''+
'''+ 2009/12/11
'''+
'''+루틴설명
'''+  1. E-Mart에서 지정할인  이용하여 결재한 경우 30% 할인해주는 내용
'''+  2. 할인 품목 정장상의(f00), 정장하의, 스커트
'''+  3. 행사 기간 : 2007-04-06 ~ 2007-04-08
'''+  4. 적용 이마트
'''       경산지사(1001)      355:해운대
'''+------------------------------------------------------
''
''    If Format(Date, "YYYY-MM-DD") >= "20091211" And Format(Date, "YYYY-MM-DD") <= "20100131" Then
''        If 가맹점정보.지사코드 <> M_COUPON_KLENZ_CODE Then
''            btn특정할인.Enabled = True
''            btn특정할인.Caption = "Y세츠무료세탁"
''            btn특정할인.BackColor = &HC0C0FF
''
''        ElseIf 가맹점정보.특정할인여부 <> "Y" Then
''            btn특정할인.Enabled = False
''        End If
''
''    ElseIf 가맹점정보.특정할인여부 <> "Y" Then
''        btn특정할인.Enabled = False
''    End If
    
    ' 크랜즈겔러리 쿠폰 할인 30% 적용
''    btn특정할인.Enabled = IIf(가맹점정보.특정할인여부 = "Y", True, False)
    
    With Combo1
        .AddItem "가죽"
        .AddItem "동물털"
        .AddItem "장식"
        .AddItem "피얼룩"
        .AddItem "잉크얼룩"
        .AddItem "페인트"
        .AddItem "기름얼룩"
        .AddItem "보풀제거"
        .AddItem "털제거"
        .AddItem "곰팡이"
        .AddItem "기타"
    End With
    
    Combo2.AddItem "드라이"
    Combo2.AddItem "수선"
    
    If intMode <> 1 Then
        SSPanel1.Visible = False
        pnlAddGubun.Visible = False
    End If
    
    
    If frm접수.sprGrid.ActiveCol = 6 Then
        ' 금액
''        pnlMoney.Visible = True
''        pnlMoney.ZOrder 0
''
''        'Me.Width = pnlMoney.Width
''        'Me.Width = 12180
        
        TabControl.SelectedItem = 1
        DoEvents
        
        If cmdOK.Visible = True Then cmdOK.SetFocus
        
        Dim sTemp As Variant
        
        Call frm접수.sprGrid.GetText(5, frm접수.sprGrid.ActiveRow, sTemp)
        
        If InStr(CStr(sTemp), "수") > 0 Then
            pnlAddGubun.Visible = True
            Combo2.ListIndex = 1
        End If
        
    Else
        TabControl.SelectedItem = 0
        DoEvents
        
''        ' 내용
''        pnlMoney.Visible = False
''        pnlMoney.ZOrder 1
''        'Me.Width = FORM_WIHTH
''
''        DoEvents
''        If btn세탁.Visible = True Then btn세탁.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Me.Hide
        Unload Me
        KeyCode = 0
        frm접수.cmdOK.SetFocus
    End If
End Sub

Private Sub Form_Load()
    frm작업.Top = frmMain.Top   '400
    frm작업.Left = frmMain.Left '10
    
    If ctlMenu1.Count > 1 Then
        For i = 1 To ctlMenu1.Count - 1
            Unload ctlMenu1(i)
        Next i
    End If
    
    i = 1
    
    '----------------------------------------------------------
    ' TB_작업
    '----------------------------------------------------------
    Query = "SELECT    작업코드"
    Query = Query & ", 작업명"
    Query = Query & ", 순서"
    Query = Query & " FROM TB_작업"
    Query = Query & " ORDER BY 순서"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    Do Until ADORs.EOF
        Load ctlMenu1(i)
        ctlMenu1(i).Left = GetLeft(i)
        ctlMenu1(i).Top = GetTop(i)
        
        Call ctlMenu1(i).SET_Item(ADORs!작업명, 0, ADORs!작업코드, "")
        
        If ADORs!작업명 = "지정할인" Then
            ' 일자 계산은 가맹점정보에서 처리한다. (Y일경우 일자에 해당한다.)
             ctlMenu1(i).Enabled = IIf(가맹점정보.지정할인여부 = "Y", True, False)
             
        ' 수선은 사용을 못하도록 처리한다.
        ElseIf ADORs!작업코드 = "w10" Then
             ctlMenu1(i).Enabled = False
             Call ctlMenu1(i).SET_Item("", 0, "", "")
        
        ElseIf ADORs!작업명 = "특정할인" Then
             ctlMenu1(i).Enabled = IIf(가맹점정보.특정할인여부 = "Y", True, False)
        
        ElseIf ADORs!작업명 = "자연수(水)" Then
        
            ' 크렌즈가 아닐 경우는 자연수를 사용하지 못하도록 한다.
            If 가맹점정보.지사코드 <> "1024" Then
                ctlMenu1(i).Enabled = False
                Call ctlMenu1(i).SET_Item("", 0, "", "")
                
                
            ' 현대 본점은 제외 한다.
            ElseIf 가맹점정보.가맹점코드 = "100475" Then
                ctlMenu1(i).Enabled = False
                Call ctlMenu1(i).SET_Item("", 0, "", "")
            End If
        
        ElseIf ADORs!작업명 = "직원세탁" Then
        
            ' 크렌즈가 아닐 경우는 직원세탁을 사용하지 못하도록 한다.
            If 가맹점정보.지사코드 <> "1024" Then
                ctlMenu1(i).Enabled = False
                Call ctlMenu1(i).SET_Item("", 0, "", "")
                
            ' 기간이 2012-10-12 ~ 2012-10-31일 까지임
            ElseIf Format(Date, "yyyy-MM-dd") < "2012-10-12" Or Format(Date, "yyyy-MM-dd") > "2012-10-31" Then
                ctlMenu1(i).Enabled = False
                Call ctlMenu1(i).SET_Item("", 0, "", "")
            End If
        
        Else
            ctlMenu1(i).Enabled = True
        End If
        
        ctlMenu1(i).Visible = True
        i = i + 1
        
        ADORs.MoveNext
    Loop
    ADORs.Close
    Set ADORs = Nothing
    
    If i > 0 Then
        Me.Height = ctlMenu1(i - 1).Top + ctlMenu1(i - 1).Height + 650
        Me.Width = 14000
    End If
    
    
    
    frm접수.sprGrid.Row = frm접수.sprGrid.ActiveRow
    
''    'pnlMoney.Move 0, 0
''    pnlMoney.Move 75, 0
''
''    Me.Width = FORM_WIHTH
''    Me.Height = 3030 '2700
End Sub

Private Sub Form_Unload(Cancel As Integer)
    intMode = 0
End Sub

Private Sub Label1_Change()
   If Len(pnlCost.Caption) > 6 Then
      pnlCost.Caption = Mid(pnlCost.Caption, 1, 6)
   End If
End Sub

Private Sub txtPassWord_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode < 48 Or KeyCode > 57) And (KeyCode < 96 Or KeyCode > 105) Then
        If txtPassWord.PasswordChar <> "*" Then
            txtPassWord.PasswordChar = "*"
        End If
    Else
        txtPassWord.PasswordChar = ""
    End If
    
    If KeyCode = vbKeyReturn Then
        cmdHeadOffice_Click
    End If

End Sub

Private Sub btnNumber_Click(Index As Integer)
    pnlCost.Caption = Trim(pnlCost.Caption) + btnNumber(Index).Caption
End Sub


Private Function GetLeft(ByVal Locate As Integer) As Long
    On Error Resume Next
    
    Select Case Locate
        Case 1, 8, 15, 22, 29, 36, 43, 50, 57, 64
            GetLeft = 45
        Case 2, 9, 16, 23, 30, 37, 44, 51, 58, 65
            GetLeft = 1905
        Case 3, 10, 17, 24, 31, 38, 45, 52, 59, 66
            GetLeft = 3765
        Case 4, 11, 18, 25, 32, 39, 46, 53, 60, 67
            GetLeft = 5625
        Case 5, 12, 19, 26, 33, 40, 47, 54, 61, 68
            GetLeft = 7485
        Case 6, 13, 20, 27, 34, 41, 48, 55, 62, 69
            GetLeft = 9345
        Case 7, 14, 21, 28, 35, 42, 49, 56, 63, 70
            GetLeft = 11205
        Case Else
            GetLeft = 0
    End Select
End Function

Private Function GetTop(ByVal Locate As Integer) As Long
    On Error Resume Next
    
    Select Case Locate
        Case 1, 2, 3, 4, 5, 6, 7
            GetTop = 45
        Case 8, 9, 10, 11, 12, 13, 14
            GetTop = 945 '915
        Case 15, 16, 17, 18, 19, 20, 21
            GetTop = 1845 '1785
        Case 22, 23, 24, 25, 26, 27, 28
            GetTop = 2745 '2655
        Case 29, 30, 31, 32, 33, 34, 35
            GetTop = 3645 '3530
        Case 36, 37, 38, 39, 40, 41, 42
            GetTop = 4545 '4500
        Case 43, 44, 45, 46, 47, 48, 49
            GetTop = 5445 '5370
        Case 50, 51, 52, 53, 54, 55, 56
            GetTop = 6345 '6240
        Case 57, 58, 59, 60, 61, 62, 63
            GetTop = 7245 '7110
        Case 64, 65, 66, 67, 68, 69, 70
            GetTop = 8145 '7980
        Case Else
            GetTop = 0
    End Select
End Function



