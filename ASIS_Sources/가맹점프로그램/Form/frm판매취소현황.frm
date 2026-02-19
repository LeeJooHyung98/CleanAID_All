VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm판매취소현황 
   Caption         =   "판매취소 현황"
   ClientHeight    =   11970
   ClientLeft      =   1995
   ClientTop       =   5025
   ClientWidth     =   16410
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form20"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11970
   ScaleWidth      =   16410
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   11970
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16410
      _ExtentX        =   28945
      _ExtentY        =   21114
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm판매취소현황.frx":0000
      Begin Threed.SSPanel SSPanel2 
         Height          =   735
         Left            =   15
         TabIndex        =   15
         Top             =   11220
         Width           =   16380
         _ExtentX        =   28893
         _ExtentY        =   1296
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   315
            Index           =   0
            Left            =   1425
            TabIndex        =   16
            Top             =   60
            Width           =   1140
            _Version        =   262145
            _ExtentX        =   2011
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   192
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            ReadOnly        =   -1  'True
            Modified        =   0   'False
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Index           =   1
            Left            =   60
            TabIndex        =   17
            Top             =   60
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "판매취소 수량"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Index           =   3
            Left            =   60
            TabIndex        =   18
            Top             =   360
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "판매취소 금액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   315
            Index           =   1
            Left            =   1425
            TabIndex        =   19
            Top             =   360
            Width           =   1140
            _Version        =   262145
            _ExtentX        =   2011
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   192
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            ReadOnly        =   -1  'True
            Modified        =   0   'False
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
            Left            =   3975
            TabIndex        =   20
            Top             =   60
            Visible         =   0   'False
            Width           =   1140
            _Version        =   262145
            _ExtentX        =   2011
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   192
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            ReadOnly        =   -1  'True
            Modified        =   0   'False
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Index           =   0
            Left            =   2610
            TabIndex        =   21
            Top             =   60
            Visible         =   0   'False
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "반품환불 수량"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Index           =   2
            Left            =   2610
            TabIndex        =   22
            Top             =   360
            Visible         =   0   'False
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "반품환불 금액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   315
            Index           =   3
            Left            =   3975
            TabIndex        =   23
            Top             =   360
            Visible         =   0   'False
            Width           =   1140
            _Version        =   262145
            _ExtentX        =   2011
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   192
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            ReadOnly        =   -1  'True
            Modified        =   0   'False
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
            Left            =   6525
            TabIndex        =   24
            Top             =   60
            Visible         =   0   'False
            Width           =   1140
            _Version        =   262145
            _ExtentX        =   2011
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   192
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            ReadOnly        =   -1  'True
            Modified        =   0   'False
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Index           =   4
            Left            =   5160
            TabIndex        =   25
            Top             =   60
            Visible         =   0   'False
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "세탁환불 수량"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Index           =   5
            Left            =   5160
            TabIndex        =   26
            Top             =   360
            Visible         =   0   'False
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "세탁환불 금액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   315
            Index           =   5
            Left            =   6525
            TabIndex        =   27
            Top             =   360
            Visible         =   0   'False
            Width           =   1140
            _Version        =   262145
            _ExtentX        =   2011
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   192
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            ReadOnly        =   -1  'True
            Modified        =   0   'False
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
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   750
         Left            =   15
         TabIndex        =   1
         Top             =   450
         Width           =   16380
         _ExtentX        =   28893
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox cboGubun 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   915
            Style           =   2  '드롭다운 목록
            TabIndex        =   8
            Top             =   405
            Width           =   1455
         End
         Begin VB.TextBox txtFind 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2415
            TabIndex        =   7
            Top             =   405
            Width           =   2400
         End
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   630
            Left            =   5655
            TabIndex        =   2
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 조회(&F)"
            Appearance      =   6
            Picture         =   "frm판매취소현황.frx":0092
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   3
            Left            =   9885
            TabIndex        =   3
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 엑셀(&E)"
            Appearance      =   6
            Picture         =   "frm판매취소현황.frx":078C
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   13170
            TabIndex        =   4
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frm판매취소현황.frx":0F06
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   4
            Left            =   11430
            TabIndex        =   5
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 출력(&P)"
            Appearance      =   6
            Picture         =   "frm판매취소현황.frx":1F98
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   0
            Left            =   915
            TabIndex        =   9
            Top             =   45
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   58916867
            CurrentDate     =   40279
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   1
            Left            =   2655
            TabIndex        =   10
            Top             =   45
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   58916867
            CurrentDate     =   40279
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
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
            Height          =   195
            Index           =   0
            Left            =   2445
            TabIndex        =   13
            Top             =   105
            Width           =   120
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "검색조건:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   3
            Left            =   45
            TabIndex        =   12
            Top             =   465
            Width           =   840
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "접수일자:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   45
            TabIndex        =   11
            Top             =   105
            Width           =   840
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   6
         Top             =   15
         Width           =   16380
         _ExtentX        =   28893
         _ExtentY        =   741
         _Version        =   262144
         Font3D          =   1
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "      판매취소 현황"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm판매취소현황.frx":2692
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm판매취소현황.frx":28B8
            Top             =   -15
            Width           =   765
         End
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   9990
         Left            =   15
         TabIndex        =   14
         Top             =   1215
         Width           =   16380
         _Version        =   524288
         _ExtentX        =   28893
         _ExtentY        =   17621
         _StockProps     =   64
         AutoCalc        =   0   'False
         BackColorStyle  =   1
         ColsFrozen      =   4
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
         DInformActiveRowChange=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         MaxCols         =   15
         MaxRows         =   200
         MoveActiveOnFocus=   0   'False
         OperationMode   =   1
         Protect         =   0   'False
         SpreadDesigner  =   "frm판매취소현황.frx":3482
         UserResize      =   1
         VisibleCols     =   13
         VisibleRows     =   50
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "frm판매취소현황"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strStart As String
Dim strEnd   As String

Private Sub Data_Display()
    On Error GoTo ErrRtn

    For i = 0 To 5
        txtNum(i).Value = 0
    Next i
    
    Query = "SELECT    B.성명"
    Query = Query & ", B.휴대전화"
    Query = Query & ", B.전화번호"
    Query = Query & ", A.접수일자"
    Query = Query & ", SUBSTRING(A.판매취소일자,1,10) AS 판매취소일자"
    Query = Query & ", SUBSTRING(A.반품환불일자,1,10) AS 반품환불일자"
    Query = Query & ", SUBSTRING(A.세탁환불일자,1,10) AS 세탁환불일자"
    Query = Query & ", A.지사출고상태"
    Query = Query & ", A.의류명"
    Query = Query & ", A.택번호"
    Query = Query & ", A.색상"
    Query = Query & ", A.무늬"
    Query = Query & ", A.내용"
    Query = Query & ", A.금액"
    Query = Query & ", A.결제여부"
    Query = Query & ", A.상표"
    Query = Query & " FROM TB_입출고 AS A LEFT OUTER JOIN TB_고객정보 AS B ON A.고객코드 = B.고객코드 "
    Query = Query & " WHERE (A.접수일자 >= '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "'"
    Query = Query & "   AND  A.접수일자 <= '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "')"
    Query = Query & "   AND (A.판매취소 = 'Y')"
'    Query = Query & "   AND (A.반품환불일자 IS NOT NULL OR A.반품환불일자 <> '')"
'    Query = Query & "   AND (A.세탁환불일자 IS NOT NULL OR A.세탁환불일자 <> ''))"
    
    'Query = Query & " WHERE (SUBSTRING(A.판매취소일자,1,10) >= '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "'"
    'Query = Query & "   AND  SUBSTRING(A.판매취소일자,1,10) <= '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "')"
    'Query = Query & "    OR (SUBSTRING(A.반품환불일자,1,10) >= '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "'"
    'Query = Query & "   AND  SUBSTRING(A.반품환불일자,1,10) <= '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "')"
    'Query = Query & "    OR (SUBSTRING(A.세탁환불일자,1,10) >= '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "'"
    'Query = Query & "   AND  SUBSTRING(A.세탁환불일자,1,10) <= '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "')"

    Select Case cboGubun.Text
        Case "성명":     Query = Query & " AND (B.성명 LIKE '%" & Trim(txtFind.Text) & "%') "
        Case "전화번호": Query = Query & " AND (B.전화번호 LIKE '%" & Trim(txtFind.Text) & "%') "
        Case "고객코드": Query = Query & " AND (B.고객코드 LIKE '%" & Trim(txtFind.Text) & "%') "
    End Select
    
    Query = Query & " ORDER BY A.접수일자, A.택번호"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = ADORs!성명 & ""         '
            .Col = 2:  .Text = ADORs!휴대전화 & ""     '
            .Col = 3:  .Text = ADORs!전화번호 & ""     '
            .Col = 4:  .Text = ADORs!접수일자 & ""     '
            .Col = 5:  .Text = ADORs!판매취소일자 & "" '
            .Col = 6:  .Text = ADORs!반품환불일자 & "" '
            .Col = 7:  .Text = ADORs!세탁환불일자 & "" '
            .Col = 8:  .Text = ADORs!의류명 & ""
            
            If Len(Trim(ADORs!택번호)) = 9 Then
                .Col = 9: .Text = Format(ADORs!택번호, "000-00-0000")
            Else
                .Col = 9: .Text = ADORs!택번호 & ""
            End If
            
            .Col = 10: .Text = ADORs!색상 & ""        '
            .Col = 11: .Text = ADORs!무늬 & ""        '
            .Col = 12: .Text = ADORs!내용 & ""        '
            .Col = 13: .Text = ADORs!금액 & ""        '
            .Col = 14: .Text = ADORs!결제여부 & ""    '
            .Col = 15: .Text = ADORs!상표 & ""        '
            
            If ADORs!판매취소일자 <> "" Then
                txtNum(0).Value = txtNum(0).Value + 1
                txtNum(1).Value = txtNum(1).Value + ADORs!금액
            End If
            
            If ADORs!반품환불일자 <> "" Then
                txtNum(2).Value = txtNum(2).Value + 1
                txtNum(3).Value = txtNum(3).Value + ADORs!금액
            End If
            
            If ADORs!세탁환불일자 <> "" Then
                txtNum(4).Value = txtNum(4).Value + 1
                txtNum(5).Value = txtNum(5).Value + ADORs!금액
            End If
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub cmdList_Click()
    Call Data_Display
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim strDate As String
    'TitleSet "찾 아 보 기"

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
        .OperationMode = OperationModeExtended
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
    
    strDate = Format(DateAdd("m", -1, Date), "YYYY-MM-DD")

    dtpDay(0).Value = Format(strDate, "YYYY-MM-DD")
    dtpDay(1).Value = Format(Date, "YYYY-MM-DD")
    
    With cboGubun
        .Clear
        .AddItem "성명"
        .AddItem "전화번호"
        .AddItem "고객코드"
        
        .ListIndex = 0
    End With
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 3: Call Export_Excel(frmMain.cdgExcel, sprGrid)
        
        Case 4:
            Rtn = MsgBox("출력 미리보기를 하시겠습니까?", vbQuestion + vbYesNo, "출력")
            
            If Rtn = vbYes Then
                Call Data_Print(True)
            Else
                Call Data_Print(False)
            End If
            
        Case 5: Unload Me
    End Select
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub Data_Print(Print_PreView As Boolean)
    On Error GoTo ErrRtn
    
    If sprGrid.MaxRows = 0 Then Exit Sub

    If Dir(AppPath & "XML", vbDirectory) = "" Then
        MkDir AppPath & "XML"
    End If

    Open AppPath & "XML\반품현황.XML" For Output As #1
    
    Print #1, "<?xml version=""1.0"" encoding=""EUC-KR""?>"
    Print #1, "<root>"
        
          XML = "    <조건>"
    XML = XML & "        <접수일자>접수일자 : " & Format(dtpDay(0).Value, "YYYY-MM-DD") & " ~ " & Format(dtpDay(1).Value, "YYYY-MM-DD") & "</접수일자>"
    XML = XML & "        <가맹점>크린에이드 - " & Func_Replace(가맹점정보.가맹점명) & "</가맹점>"
    XML = XML & "        <판매취소수량>" & txtNum(0).Text & "</판매취소수량>"
    XML = XML & "        <판매취소금액>" & txtNum(1).Text & "</판매취소금액>"
    XML = XML & "        <반품환불수량>" & txtNum(2).Text & "</반품환불수량>"
    XML = XML & "        <반품환불금액>" & txtNum(3).Text & "</반품환불금액>"
    XML = XML & "        <세탁환불수량>" & txtNum(4).Text & "</세탁환불수량>"
    XML = XML & "        <세탁환불금액>" & txtNum(5).Text & "</세탁환불금액>"
    
    XML = XML & "   </조건>"
    Print #1, XML
    
    With sprGrid
        For i = 1 To .MaxRows
            .Row = i
            
                             XML = "    <Data>"
            .Col = 1:  XML = XML & "        <성명>" & Func_Replace(.Text) & "</성명>"
            .Col = 2:  XML = XML & "        <전화번호>" & Func_Replace(.Text) & "</전화번호>"
            .Col = 3:  XML = XML & "        <휴대전화>" & Func_Replace(.Text) & "</휴대전화>"
            .Col = 4:  XML = XML & "        <접수일자>" & .Text & "</접수일자>"
            .Col = 5:  XML = XML & "        <판매취소>" & Func_Replace(.Text) & "</판매취소>"
            .Col = 6:  XML = XML & "        <반품환불>" & Func_Replace(.Text) & "</반품환불>"
            .Col = 7:  XML = XML & "        <세탁환불>" & Func_Replace(.Text) & "</세탁환불>"
            .Col = 8:  XML = XML & "        <품명>" & Func_Replace(.Text) & "</품명>"
            .Col = 9:  XML = XML & "        <택번호>" & .Text & "</택번호>"
            .Col = 10: XML = XML & "        <색상>" & Func_Replace(.Text) & "</색상>"
            .Col = 11: XML = XML & "        <무늬>" & Func_Replace(.Text) & "</무늬>"
            .Col = 12: XML = XML & "        <내용>" & Func_Replace(.Text) & "</내용>"
            .Col = 13: XML = XML & "        <금액>" & .Text & "</금액>"
            .Col = 14: XML = XML & "        <결제>" & Func_Replace(.Text) & "</결제>"
            .Col = 15: XML = XML & "        <상표>" & Func_Replace(.Text) & "</상표>"
                       XML = XML & "   </Data>"
                       Print #1, XML
        Next i
        
        Print #1, "</root>"
        Close #1
    End With
    
    If Print_PreView = True Then
        With rpt반품현황
            .dc.FileURL = AppPath & "XML\반품현황.XML"
            .Show 1
        End With
    Else
        With rpt반품현황
            .dc.FileURL = AppPath & "XML\반품현황.XML"
            .PrintReport False
        End With
    
        Unload rpt반품현황
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    cmdBtn(5).Left = Me.Width - cmdBtn(5).Width - 200
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdList_Click
    End If
End Sub
