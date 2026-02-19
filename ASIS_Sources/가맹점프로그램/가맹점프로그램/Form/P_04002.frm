VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm부자재주문 
   Caption         =   "부자재 주문"
   ClientHeight    =   11730
   ClientLeft      =   585
   ClientTop       =   2070
   ClientWidth     =   14820
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11730
   ScaleWidth      =   14820
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   11730
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14820
      _ExtentX        =   26141
      _ExtentY        =   20690
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_04002.frx":0000
      Begin Threed.SSPanel SSPanel 
         Height          =   540
         Index           =   1
         Left            =   15
         TabIndex        =   27
         Top             =   4290
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   953
         _Version        =   262144
         BevelOuter      =   0
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   45
            TabIndex        =   28
            Top             =   45
            Width           =   1275
            _Version        =   851970
            _ExtentX        =   2249
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 신규(&N)"
            Appearance      =   6
            Picture         =   "P_04002.frx":00B2
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   1395
            TabIndex        =   29
            Top             =   45
            Width           =   1275
            _Version        =   851970
            _ExtentX        =   2249
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 저장(&S)"
            Appearance      =   6
            Picture         =   "P_04002.frx":0AC4
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   2745
            TabIndex        =   30
            Top             =   45
            Width           =   1275
            _Version        =   851970
            _ExtentX        =   2249
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 삭제(&D)"
            Appearance      =   6
            Picture         =   "P_04002.frx":14D6
         End
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   2940
         Index           =   0
         Left            =   15
         TabIndex        =   2
         Top             =   1335
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   5186
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox cboGoods 
            Height          =   315
            Left            =   930
            Style           =   2  '드롭다운 목록
            TabIndex        =   8
            Top             =   420
            Width           =   3015
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   1
            Left            =   930
            TabIndex        =   3
            Top             =   2520
            Width           =   3420
         End
         Begin MSComCtl2.DTPicker dtpOrderDay 
            Height          =   315
            Left            =   930
            TabIndex        =   4
            Top             =   75
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Format          =   56557568
            CurrentDate     =   36686
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   315
            Index           =   0
            Left            =   930
            TabIndex        =   9
            Top             =   780
            Width           =   810
            _Version        =   262145
            _ExtentX        =   1429
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Insert          =   0   'False
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   2
            StartText.y     =   2
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   16
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   315
            Index           =   2
            Left            =   930
            TabIndex        =   19
            Top             =   1470
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Insert          =   0   'False
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   2
            StartText.y     =   2
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   16
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   315
            Index           =   3
            Left            =   930
            TabIndex        =   21
            Top             =   1815
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Insert          =   0   'False
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   2
            StartText.y     =   2
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   16
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   315
            Index           =   4
            Left            =   930
            TabIndex        =   23
            Top             =   2160
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Insert          =   0   'False
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   2
            StartText.y     =   2
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   16
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   315
            Index           =   1
            Left            =   930
            TabIndex        =   31
            Top             =   1125
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Insert          =   0   'False
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   2
            StartText.y     =   2
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   16
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "단가:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   8
            Left            =   60
            TabIndex        =   32
            Top             =   1200
            Width           =   825
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "비고:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   7
            Left            =   60
            TabIndex        =   25
            Top             =   2580
            Width           =   825
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "합계금액:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   6
            Left            =   60
            TabIndex        =   24
            Top             =   2235
            Width           =   825
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "세액:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   5
            Left            =   60
            TabIndex        =   22
            Top             =   1890
            Width           =   825
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "공급가액:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   4
            Left            =   60
            TabIndex        =   20
            Top             =   1545
            Width           =   825
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "주문수량:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   60
            TabIndex        =   18
            Top             =   855
            Width           =   825
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "부자재:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   60
            TabIndex        =   17
            Top             =   480
            Width           =   825
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "주문일자:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   60
            TabIndex        =   16
            Top             =   120
            Width           =   825
         End
      End
      Begin Threed.SSPanel panInput 
         Height          =   885
         Left            =   15
         TabIndex        =   1
         Top             =   435
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   1561
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   0
            Left            =   1035
            TabIndex        =   5
            Top             =   60
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            Format          =   56557568
            CurrentDate     =   36686
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   1
            Left            =   4215
            TabIndex        =   6
            Top             =   60
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            Format          =   56557568
            CurrentDate     =   36686
         End
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   750
            Left            =   7500
            TabIndex        =   11
            Top             =   75
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   79
            Caption         =   " 조회(&F)"
            Appearance      =   6
            Picture         =   "P_04002.frx":1EE8
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   750
            Index           =   3
            Left            =   10035
            TabIndex        =   12
            Top             =   75
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   79
            Caption         =   " 엑셀(&E)"
            Appearance      =   6
            Picture         =   "P_04002.frx":25E2
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   750
            Index           =   5
            Left            =   13125
            TabIndex        =   13
            Top             =   75
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "P_04002.frx":2D5C
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   750
            Index           =   4
            Left            =   11580
            TabIndex        =   14
            Top             =   75
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   79
            Caption         =   " 출력(&P)"
            Appearance      =   6
            Picture         =   "P_04002.frx":3DEE
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "주문일자:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   60
            TabIndex        =   15
            Top             =   135
            Width           =   825
         End
         Begin VB.Label Label 
            Alignment       =   2  '가운데 맞춤
            Caption         =   "~"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3900
            TabIndex        =   7
            Top             =   120
            Width           =   300
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   405
         Left            =   15
         TabIndex        =   10
         Top             =   15
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   714
         _Version        =   262144
         Font3D          =   3
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "      부자재 주문"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_04002.frx":44E8
         BorderWidth     =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image1 
            Height          =   240
            Left            =   60
            Picture         =   "P_04002.frx":494A
            Top             =   75
            Width           =   240
         End
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   6870
         Left            =   15
         TabIndex        =   26
         Top             =   4845
         Width           =   14790
         _Version        =   524288
         _ExtentX        =   26088
         _ExtentY        =   12118
         _StockProps     =   64
         BackColorStyle  =   1
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         MaxCols         =   11
         MaxRows         =   35
         ScrollBars      =   0
         SpreadDesigner  =   "P_04002.frx":4ED4
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "frm부자재주문"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub DataSave()
    ReDim sValue(7)
    
    sValue(0) = Mid(cboStore.Text, 2, 6)                          ' 1 가맹점코드
    sValue(1) = Format(dtpDay(3).Value, "YYYY-MM-DD")            ' 2 입금일자
    sValue(2) = Mid(cboManager.Text, 2, 3)                        ' 3 배송기사코드
    sValue(3) = Trim(Mid(cboManager.Text, 6, Len(cboManager.Text) - 5)) ' 4 배송기사명
    sValue(4) = txtNum.Value                                    ' 5 입금액
    sValue(5) = txtInput(0).Text & ""                             ' 6 비고
    sValue(6) = txtInput(1).Text & ""                             ' 7 경리담당자
    sValue(7) = Mid(cboOffice.Text, 2, 4)                         ' 8 지사코드
    
    Call ExecPro("SP_04002_01", sValue(), Err_Num, Err_Dec)
    
    If Err_Num <> 0 Then
        MsgBox "[" & Err_Num & "] " & Err_Dec
    End If
    
    Call Data_Display
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: Call Data_Display   ' 조회
        Case 1: 'Call DataAdd        ' 신규
        Case 2: Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
        Case 4: 'Call DataCancel     ' 취소
        Case 5: 'Call DataPrint      ' 인쇄
        Case 6: 'Call DataScreen     '
        Case 7: Unload Me           ' 종료
    End Select
    
'    Me.MousePointer = 0
    
    Exit Sub
    
ErrRtn:
    Me.MousePointer = 0
    
    If Err.Number = "0" Then
        
    ElseIf Err.Number = "91" Then
        End
    Else
        Resume Next
    End If
End Sub

Private Sub cmdList_Click()
    If Server_Connection(HostCon) = False Then Exit Sub
    
    ReDim sValue(3)
    
    sValue(0) = 가맹점정보.가맹점코드
    sValue(1) = Format(dtpDay(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtpDay(1).Value, "YYYY-MM-DD")
    sValue(3) = 0
    
    Set ADORs = New ADODB.Recordset
    Set ADORs = SP_Exec("SP_R_부자재주문", sValue(), Err_Num, Err_Dec)
    
    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        
            .Col = 1:  .Text = ADORs!주문일자 & "" ' 1
            .Col = 2:  .Text = ADORs!주문코드 & "" ' 2
            .Col = 3:  .Text = ADORs!부자재명 & "" ' 3
            .Col = 4:  .Text = ADORs!규격 & ""     ' 4
            .Col = 5:  .Text = ADORs!수량 & ""     ' 5
            .Col = 6:  .Text = ADORs!단가 & ""     ' 6
            .Col = 7:  .Text = ADORs!공급가액 & "" ' 7
            .Col = 8:  .Text = ADORs!세액 & ""     ' 8
            .Col = 9:  .Text = ADORs!합계금액 & "" ' 9
            .Col = 10: .Text = ADORs!비고 & ""     '10
            .Col = 11: .Text = ADORs!출고일자 & "" '11
        
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        .ReDraw = True
    End With
    
End Sub

Private Sub Form_Activate()

End Sub

Private Sub Form_Load()
    With sprGrid
        .MaxRows = 0
        .RowHeight(-1) = 14
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeSingle
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
    
    dtpDay(0).Value = Date
    dtpDay(1).Value = Date

End Sub

Private Sub Form_Unload(Cancel As Integer)

End Sub

Private Sub Data_Display()
    Dim i As Integer
    Dim j As Integer
    
    ReDim sValue(3)
    
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    
    If Mid(cboInput.Text, 2, 6) = "000000" Then
        sValue(1) = ""
    Else
        sValue(1) = Mid(cboInput.Text, 2, 6)
    End If
    
    sValue(2) = Format(dtpDay(0).Value, "YYYY-MM-DD")
    sValue(3) = Format(dtpDay(1).Value, "YYYY-MM-DD")
        
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_04002_00", sValue(), Err_Num, Err_Dec)
    
    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = RS01!입금일자 & ""
            .Col = 2:  .Text = RS01!가맹점코드 & ""
            .Col = 3:  .Text = RS01!가맹점명 & ""
            .Col = 4:  .Text = RS01!지사코드 & ""
            .Col = 5:  .Text = RS01!지사명 & ""
            .Col = 6:  .Text = RS01!배송기사코드 & ""
            .Col = 7:  .Text = RS01!배송기사명 & ""
            
            .Col = 8:  .Text = RS01!입금액 & ""
            .Col = 9:  .Text = RS01!비고 & ""
            .Col = 10:  .Text = RS01!입금확정 & ""
            .Col = 11:  .Text = RS01!확정일자 & ""
            .Col = 12: .Text = RS01!경리담당자 & ""
            .Col = 13: .Text = "0"
                        
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .ReDraw = True
    End With
End Sub

Private Sub DataScreen()
    Dim i As Integer
    Dim j As Integer
    
    Dim ReportFP As String
    Dim ReportFile As String
    Dim sData As String
    
    Dim AgencySL As String
    Dim iCnt As Integer
    
    AgencySL = "({PRO_P_04002_00;1.대리점명} = '  ' "
    
    For i = 1 To sprGrid.MaxRows
        sprGrid.Row = i
        sprGrid.Col = 4
        
        If sprGrid.Value = True Then
            sprGrid.Col = 1
            AgencySL = AgencySL & " Or {PRO_P_04002_00;1.대리점명} = '" & spdView.Value & "' "
        End If
    Next i
    
    AgencySL = AgencySL & ")"
    
    ReportFP = GetIniStr("REPORT", "FilePath", "", sIniFile)
    ReportFile = ReportFP & "\" & Me.Name & ".rpt"

    Dim ii As Integer
    For ii = 0 To 30
        P_00000.crPrint.Formulas(ii) = ""
    Next
    
    P_00000.crPrint.StoredProcParam(0) = "0"
    P_00000.crPrint.StoredProcParam(1) = Format(dtpDay(0).Value, "yyyymmdd")
    
    P_00000.crPrint.WindowTitle = Me.Caption
    
    P_00000.crPrint.Formulas(0) = "수금일자 = '" & Format(dtpDay(0).Value, "yyyymmdd") & "'"
    P_00000.crPrint.Formulas(1) = "발행일자 = '" & Format(dtpDay(1).Value, "yyyymmdd") & "'"
    P_00000.crPrint.Formulas(2) = "경리 = '" & txtInput(0).Text & "'"
    P_00000.crPrint.Formulas(3) = "담당 = '" & txtInput(1).Text & "'"
    
    P_00000.crPrint.SelectionFormula = AgencySL
    
    Call ReportPrint(ReportFile, "2")
End Sub

Private Sub DataPrint()
    Dim i As Integer
    Dim j As Integer
    
    Dim ReportFP As String
    Dim ReportFile As String
    Dim sData As String
    
    Dim AgencySL As String
    Dim iCnt As Integer
    
    AgencySL = "({PRO_P_04002_00;1.대리점명} = '  ' "
    
    For i = 1 To sprGrid.MaxRows
        sprGrid.Row = i
        sprGrid.Col = 4
        
        If sprGrid.Value = True Then
            sprGrid.Col = 1
            AgencySL = AgencySL & " Or {PRO_P_04002_00;1.대리점명} = '" & spdView.Value & "' "
        End If
    Next i
    
    AgencySL = AgencySL & ")"
    
    ReportFP = GetIniStr("REPORT", "FilePath", "", sIniFile)
    ReportFile = ReportFP & "\" & Me.Name & ".rpt"

    P_00000.crPrint.StoredProcParam(0) = "0"
    P_00000.crPrint.StoredProcParam(1) = Format(dtpDay(0).Value, "yyyymmdd")
    
    P_00000.crPrint.WindowTitle = Me.Caption
    
    Dim ii As Integer
    For ii = 0 To 30
        P_00000.crPrint.Formulas(ii) = ""
    Next
    
    P_00000.crPrint.Formulas(0) = "수금일자 = '" & Format(dtpDay(0).Value, "yyyymmdd") & "'"
    P_00000.crPrint.Formulas(1) = "발행일자 = '" & Format(dtpDay(1).Value, "yyyymmdd") & "'"
    P_00000.crPrint.Formulas(2) = "경리 = '" & txtInput(0).Text & "'"
    P_00000.crPrint.Formulas(3) = "담당 = '" & txtInput(1).Text & "'"
    
    P_00000.crPrint.SelectionFormula = AgencySL
    
    Call ReportPrint(ReportFile, "1")
End Sub

Private Sub sprGrid_Click(ByVal Col As Long, ByVal Row As Long)
    Dim i As Integer
    
    If Row <= 0 Then Exit Sub
    
    With sprGrid
        .Row = Row
        
        .Col = 1: dtpDay(3).Value = Format(.Text, "YYYY-MM-DD")
        
        .Col = 2:
        If Trim(.Text) <> "" Then
            For i = 0 To cboStore.ListCount - 1
                If Trim(.Text) = Mid(cboStore.List(i), 2, 6) Then
                    cboStore.ListIndex = i
                    Exit For
                End If
            Next i
        Else
            cboStore.ListIndex = -1
        End If
        
        .Col = 6:
        If Trim(.Text) <> "" Then
            For i = 0 To cboManager.ListCount - 1
                If Trim(.Text) = Mid(cboManager.List(i), 2, 3) Then
                    cboManager.ListIndex = i
                    Exit For
                End If
            Next i
        Else
            cboManager.ListIndex = -1
        End If
        
        .Col = 8: txtNum.Value = .Value
        .Col = 9: txtInput(0).Text = .Text & ""
        .Col = 12: txtInput(1).Text = .Text & ""
    End With
End Sub
