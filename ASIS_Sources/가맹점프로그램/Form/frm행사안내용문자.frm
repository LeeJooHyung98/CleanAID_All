VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm행사안내용문자 
   Caption         =   "행사 안내용 문자"
   ClientHeight    =   11475
   ClientLeft      =   5700
   ClientTop       =   2130
   ClientWidth     =   15045
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
   LinkTopic       =   "Form23"
   MDIChild        =   -1  'True
   ScaleHeight     =   11475
   ScaleWidth      =   15045
   WindowState     =   2  '최대화
   Begin Threed.SSPanel SSPanel_SMS 
      Height          =   1455
      Left            =   2670
      TabIndex        =   97
      Top             =   1230
      Visible         =   0   'False
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   2566
      _Version        =   262144
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "※ 해당 자동충전 이외의 충전은 절대 충전 불가"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   101
         Top             =   870
         Width           =   4725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "   자동 충전 됩니다."
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   100
         Top             =   630
         Width           =   2100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "※ 당일 가맹점 명칭으로 입금 하시면 충전은 오후 5시~6시 사이에"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   99
         Top             =   390
         Width           =   6510
      End
      Begin VB.Label Label1 
         Caption         =   "※ 입금 계좌: 농협 351-1091-8911-93 , 예금주 : (주)크린에이드"
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   0
         Left            =   60
         TabIndex        =   98
         Top             =   90
         Width           =   6615
      End
   End
   Begin Threed.SSPanel pnlProg 
      Height          =   735
      Left            =   45
      TabIndex        =   18
      Top             =   2070
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   1296
      _Version        =   262144
      Font3D          =   3
      CaptionStyle    =   1
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureBackgroundStyle=   2
      PictureBackground=   "frm행사안내용문자.frx":0000
      BorderWidth     =   0
      BevelOuter      =   0
      BevelInner      =   2
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   11475
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   20241
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm행사안내용문자.frx":0226
      Begin Threed.SSPanel SSPanel 
         Height          =   480
         Index           =   0
         Left            =   15
         TabIndex        =   4
         Top             =   1215
         Width           =   9600
         _ExtentX        =   16933
         _ExtentY        =   847
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin CSTextLibCtl.silgEdit txtCount 
            Height          =   375
            Index           =   0
            Left            =   4545
            TabIndex        =   5
            Top             =   60
            Width           =   795
            _Version        =   262145
            _ExtentX        =   1402
            _ExtentY        =   661
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.74
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            ReadOnly        =   -1  'True
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "0"
            Text            =   " 0"
            StartText.x     =   2
            StartText.y     =   4
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
            Undo            =   1
            Data            =   0
         End
         Begin CSTextLibCtl.silgEdit txtCount 
            Height          =   375
            Index           =   1
            Left            =   6480
            TabIndex        =   6
            Top             =   60
            Width           =   795
            _Version        =   262145
            _ExtentX        =   1402
            _ExtentY        =   661
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.74
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            ReadOnly        =   -1  'True
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "0"
            Text            =   " 0"
            StartText.x     =   2
            StartText.y     =   4
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
            Undo            =   1
            Data            =   0
         End
         Begin CSTextLibCtl.silgEdit txtCount 
            Height          =   375
            Index           =   2
            Left            =   8685
            TabIndex        =   7
            Top             =   60
            Width           =   795
            _Version        =   262145
            _ExtentX        =   1402
            _ExtentY        =   661
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.74
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            ReadOnly        =   -1  'True
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "0"
            Text            =   " 0"
            StartText.x     =   2
            StartText.y     =   4
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
            Undo            =   1
            Data            =   0
         End
         Begin XtremeSuiteControls.PushButton btnSelect 
            Height          =   390
            Index           =   0
            Left            =   45
            TabIndex        =   8
            Top             =   45
            Width           =   1005
            _Version        =   851970
            _ExtentX        =   1773
            _ExtentY        =   688
            _StockProps     =   79
            Caption         =   "전체선택"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
         End
         Begin XtremeSuiteControls.PushButton btnSelect 
            Height          =   390
            Index           =   1
            Left            =   1080
            TabIndex        =   9
            Top             =   45
            Width           =   1005
            _Version        =   851970
            _ExtentX        =   1773
            _ExtentY        =   688
            _StockProps     =   79
            Caption         =   "전체취소"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
         End
         Begin XtremeSuiteControls.PushButton btnSelect 
            Height          =   390
            Index           =   2
            Left            =   3150
            TabIndex        =   103
            Top             =   45
            Width           =   1335
            _Version        =   851970
            _ExtentX        =   2355
            _ExtentY        =   688
            _StockProps     =   79
            Caption         =   "전송가능금액"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "전송후 금액:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   7
            Left            =   7545
            TabIndex        =   11
            Top             =   165
            Width           =   1080
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "선택수량:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   5
            Left            =   5610
            TabIndex        =   10
            Top             =   165
            Width           =   810
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   3
         Top             =   15
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   741
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   0
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
         Caption         =   "      행사 안내용 문자"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm행사안내용문자.frx":02D8
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm행사안내용문자.frx":04FE
            Top             =   -15
            Width           =   765
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   750
         Left            =   15
         TabIndex        =   12
         Top             =   450
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton PushButton1 
            Height          =   615
            Left            =   8760
            TabIndex        =   102
            Top             =   60
            Width           =   2085
            _Version        =   851970
            _ExtentX        =   3678
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "문자 충전 방법"
            ForeColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin VB.ComboBox cboView 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frm행사안내용문자.frx":10C8
            Left            =   1050
            List            =   "frm행사안내용문자.frx":10D2
            Style           =   2  '드롭다운 목록
            TabIndex        =   94
            Top             =   30
            Width           =   3165
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   0
            Left            =   1050
            TabIndex        =   0
            Top             =   420
            Width           =   1440
            _ExtentX        =   2540
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
            Format          =   56033283
            CurrentDate     =   39596
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   1
            Left            =   2775
            TabIndex        =   1
            Top             =   420
            Width           =   1440
            _ExtentX        =   2540
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
            Format          =   56033283
            CurrentDate     =   39596
         End
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   630
            Left            =   10890
            TabIndex        =   13
            Top             =   60
            Width           =   1350
            _Version        =   851970
            _ExtentX        =   2381
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 조회(&F)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm행사안내용문자.frx":10F2
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   4
            Left            =   12285
            TabIndex        =   14
            Top             =   60
            Visible         =   0   'False
            Width           =   1320
            _Version        =   851970
            _ExtentX        =   2328
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   "개별발송"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm행사안내용문자.frx":17EC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   13635
            TabIndex        =   15
            Top             =   60
            Width           =   1320
            _Version        =   851970
            _ExtentX        =   2328
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm행사안내용문자.frx":1EE6
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "검색 일자:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   96
            Top             =   480
            Width           =   900
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "조회 구분:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   95
            Top             =   90
            Width           =   900
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
            Left            =   2565
            TabIndex        =   16
            Top             =   480
            Width           =   120
         End
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   9750
         Left            =   15
         TabIndex        =   17
         Top             =   1710
         Width           =   9600
         _Version        =   524288
         _ExtentX        =   16933
         _ExtentY        =   17198
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowUserFormulas=   -1  'True
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
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
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         MaxCols         =   8
         MaxRows         =   300
         MoveActiveOnFocus=   0   'False
         SpreadDesigner  =   "frm행사안내용문자.frx":2F78
         VisibleCols     =   2
         VisibleRows     =   50
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin XtremeSuiteControls.TabControl TabControl 
         Height          =   10245
         Left            =   9630
         TabIndex        =   19
         Top             =   1215
         Width           =   5400
         _Version        =   851970
         _ExtentX        =   9525
         _ExtentY        =   18071
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   3
         Color           =   16
         PaintManager.BoldSelected=   -1  'True
         PaintManager.HotTracking=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         PaintManager.ButtonMargin=   "0,3,0,3"
         ItemCount       =   2
         Item(0).Caption =   "발송정보"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "TabControlPage(0)"
         Item(1).Caption =   "개별 문자발송"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "TabControlPage1"
         Begin XtremeSuiteControls.TabControlPage TabControlPage1 
            Height          =   9765
            Left            =   -69970
            TabIndex        =   20
            Top             =   450
            Visible         =   0   'False
            Width           =   5340
            _Version        =   851970
            _ExtentX        =   9419
            _ExtentY        =   17224
            _StockProps     =   1
            BackColor       =   16777215
            Page            =   1
            Begin VB.TextBox txtRecvTel 
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2565
               Locked          =   -1  'True
               TabIndex        =   25
               Top             =   5880
               Width           =   2265
            End
            Begin VB.TextBox txtServer 
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   1365
               TabIndex        =   24
               Top             =   8685
               Visible         =   0   'False
               Width           =   2895
            End
            Begin VB.TextBox txtServer 
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   1365
               TabIndex        =   23
               Top             =   8340
               Visible         =   0   'False
               Width           =   2895
            End
            Begin VB.TextBox txtServer 
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   1365
               TabIndex        =   22
               Top             =   7995
               Visible         =   0   'False
               Width           =   2895
            End
            Begin VB.TextBox txtServer 
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   1365
               TabIndex        =   21
               Top             =   7650
               Visible         =   0   'False
               Width           =   2895
            End
            Begin XtremeSuiteControls.PushButton cmdSvr 
               Height          =   450
               Index           =   0
               Left            =   885
               TabIndex        =   26
               Top             =   9090
               Visible         =   0   'False
               Width           =   1230
               _Version        =   851970
               _ExtentX        =   2170
               _ExtentY        =   794
               _StockProps     =   79
               Caption         =   "초기화"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.PushButton cmdSvr 
               Height          =   450
               Index           =   1
               Left            =   2145
               TabIndex        =   27
               Top             =   9090
               Visible         =   0   'False
               Width           =   1230
               _Version        =   851970
               _ExtentX        =   2170
               _ExtentY        =   794
               _StockProps     =   79
               Caption         =   "연결 확인"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.PushButton cmdSvr 
               Height          =   450
               Index           =   2
               Left            =   3405
               TabIndex        =   28
               Top             =   9090
               Visible         =   0   'False
               Width           =   1230
               _Version        =   851970
               _ExtentX        =   2170
               _ExtentY        =   794
               _StockProps     =   79
               Caption         =   "저장"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
            End
            Begin Threed.SSPanel SSPanel 
               Height          =   2100
               Index           =   1
               Left            =   75
               TabIndex        =   29
               Top             =   5280
               Width           =   2070
               _ExtentX        =   3651
               _ExtentY        =   3704
               _Version        =   262144
               BackColor       =   16777215
               PictureFrames   =   1
               Picture         =   "frm행사안내용문자.frx":36EE
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin VB.TextBox txtSend2 
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  '없음
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1575
                  Left            =   120
                  MultiLine       =   -1  'True
                  TabIndex        =   30
                  Top             =   360
                  Width           =   1830
               End
               Begin VB.Label lblLan2 
                  Alignment       =   1  '오른쪽 맞춤
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   0  '투명
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000C0&
                  Height          =   180
                  Left            =   1200
                  TabIndex        =   31
                  Top             =   105
                  Width           =   105
               End
            End
            Begin Threed.SSFrame SSFrame 
               Height          =   5040
               Index           =   1
               Left            =   105
               TabIndex        =   32
               Top             =   105
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   8890
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "받는 사람"
               Begin VB.TextBox txtMobile 
                  Height          =   360
                  Index           =   0
                  Left            =   90
                  TabIndex        =   42
                  Top             =   270
                  Width           =   1950
               End
               Begin VB.TextBox txtMobile 
                  Height          =   360
                  Index           =   1
                  Left            =   90
                  TabIndex        =   41
                  Top             =   675
                  Width           =   1950
               End
               Begin VB.TextBox txtMobile 
                  Height          =   360
                  Index           =   2
                  Left            =   90
                  TabIndex        =   40
                  Top             =   1080
                  Width           =   1950
               End
               Begin VB.TextBox txtMobile 
                  Height          =   360
                  Index           =   3
                  Left            =   90
                  TabIndex        =   39
                  Top             =   1485
                  Width           =   1950
               End
               Begin VB.TextBox txtMobile 
                  Height          =   360
                  Index           =   4
                  Left            =   90
                  TabIndex        =   38
                  Top             =   1890
                  Width           =   1950
               End
               Begin VB.TextBox txtMobile 
                  Height          =   360
                  Index           =   5
                  Left            =   90
                  TabIndex        =   37
                  Top             =   2295
                  Width           =   1950
               End
               Begin VB.TextBox txtMobile 
                  Height          =   360
                  Index           =   6
                  Left            =   90
                  TabIndex        =   36
                  Top             =   2700
                  Width           =   1950
               End
               Begin VB.TextBox txtMobile 
                  Height          =   360
                  Index           =   7
                  Left            =   90
                  TabIndex        =   35
                  Top             =   3105
                  Width           =   1950
               End
               Begin VB.TextBox txtMobile 
                  Height          =   360
                  Index           =   8
                  Left            =   90
                  TabIndex        =   34
                  Top             =   3510
                  Width           =   1950
               End
               Begin VB.TextBox txtMobile 
                  Height          =   360
                  Index           =   9
                  Left            =   90
                  TabIndex        =   33
                  Top             =   3915
                  Width           =   1950
               End
            End
            Begin XtremeSuiteControls.PushButton cmdSend 
               Height          =   630
               Left            =   2565
               TabIndex        =   43
               Top             =   6300
               Width           =   2250
               _Version        =   851970
               _ExtentX        =   3969
               _ExtentY        =   1111
               _StockProps     =   79
               Caption         =   " 문자보내기"
               Appearance      =   6
               Picture         =   "frm행사안내용문자.frx":11890
            End
            Begin Threed.SSPanel SSPanel2 
               Height          =   30
               Index           =   1
               Left            =   90
               TabIndex        =   44
               Top             =   5205
               Width           =   5160
               _ExtentX        =   9102
               _ExtentY        =   53
               _Version        =   262144
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   $"frm행사안내용문자.frx":11F8A
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   720
               Index           =   6
               Left            =   2580
               TabIndex        =   104
               Top             =   4050
               Width           =   2340
            End
            Begin VB.Image Image 
               Height          =   240
               Index           =   1
               Left            =   2535
               Picture         =   "frm행사안내용문자.frx":11FF0
               Top             =   5595
               Width           =   240
            End
            Begin VB.Label lblTitle 
               BackColor       =   &H00FFC0C0&
               BackStyle       =   0  '투명
               Caption         =   "보내는 사람"
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
               Left            =   2865
               TabIndex        =   49
               Top             =   5610
               Width           =   1200
            End
            Begin VB.Label lblTitle 
               Alignment       =   1  '오른쪽 맞춤
               BackColor       =   &H00FFC0C0&
               BackStyle       =   0  '투명
               Caption         =   "비밀번호 :"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   4
               Left            =   165
               TabIndex        =   48
               Top             =   8760
               Visible         =   0   'False
               Width           =   1140
            End
            Begin VB.Label lblTitle 
               Alignment       =   1  '오른쪽 맞춤
               BackColor       =   &H00FFC0C0&
               BackStyle       =   0  '투명
               Caption         =   "사용자 이름:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   165
               TabIndex        =   47
               Top             =   8415
               Visible         =   0   'False
               Width           =   1140
            End
            Begin VB.Label lblTitle 
               Alignment       =   1  '오른쪽 맞춤
               BackColor       =   &H00FFC0C0&
               BackStyle       =   0  '투명
               Caption         =   "서버 DB :"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   165
               TabIndex        =   46
               Top             =   8070
               Visible         =   0   'False
               Width           =   1140
            End
            Begin VB.Label lblTitle 
               Alignment       =   1  '오른쪽 맞춤
               BackColor       =   &H00FFC0C0&
               BackStyle       =   0  '투명
               Caption         =   "서버 IP :"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   165
               TabIndex        =   45
               Top             =   7725
               Visible         =   0   'False
               Width           =   1140
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage 
            Height          =   9765
            Index           =   0
            Left            =   30
            TabIndex        =   50
            Top             =   450
            Width           =   5340
            _Version        =   851970
            _ExtentX        =   9419
            _ExtentY        =   17224
            _StockProps     =   1
            BackColor       =   16777215
            Page            =   0
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2565
               Locked          =   -1  'True
               TabIndex        =   52
               Top             =   6750
               Width           =   2265
            End
            Begin Threed.SSPanel SSPanel2 
               Height          =   30
               Index           =   0
               Left            =   90
               TabIndex        =   51
               Top             =   5205
               Width           =   5160
               _ExtentX        =   9102
               _ExtentY        =   53
               _Version        =   262144
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel SSPanel1 
               Height          =   2115
               Left            =   75
               TabIndex        =   53
               Top             =   5280
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   3731
               _Version        =   262144
               PictureFrames   =   1
               Picture         =   "frm행사안내용문자.frx":1237A
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin VB.TextBox txtSMS 
                  Appearance      =   0  '평면
                  BorderStyle     =   0  '없음
                  Height          =   1590
                  Left            =   90
                  MultiLine       =   -1  'True
                  TabIndex        =   54
                  Top             =   360
                  Width           =   1875
               End
               Begin VB.Label lbl_SMS 
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  '투명
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000C0&
                  Height          =   180
                  Left            =   1200
                  TabIndex        =   56
                  Top             =   90
                  Width           =   105
               End
               Begin VB.Label lblNum 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   180
                  Left            =   555
                  TabIndex        =   55
                  Top             =   90
                  Width           =   105
               End
            End
            Begin XtremeSuiteControls.PushButton cmdBtn 
               Height          =   630
               Index           =   1
               Left            =   2565
               TabIndex        =   57
               Top             =   7170
               Width           =   2250
               _Version        =   851970
               _ExtentX        =   3969
               _ExtentY        =   1111
               _StockProps     =   79
               Caption         =   " 문자메시지 보내기"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
               Picture         =   "frm행사안내용문자.frx":2051C
            End
            Begin XtremeSuiteControls.PushButton cmdSendTextSave 
               Height          =   390
               Index           =   2
               Left            =   1065
               TabIndex        =   58
               ToolTipText     =   "메시지 삭제..."
               Top             =   7410
               Width           =   435
               _Version        =   851970
               _ExtentX        =   767
               _ExtentY        =   688
               _StockProps     =   79
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
               Picture         =   "frm행사안내용문자.frx":20C16
            End
            Begin XtremeSuiteControls.PushButton cmdSendTextSave 
               Height          =   390
               Index           =   1
               Left            =   570
               TabIndex        =   59
               ToolTipText     =   "메시지 수정..."
               Top             =   7410
               Width           =   435
               _Version        =   851970
               _ExtentX        =   767
               _ExtentY        =   688
               _StockProps     =   79
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
               Picture         =   "frm행사안내용문자.frx":211B0
            End
            Begin XtremeSuiteControls.PushButton cmdSendTextSave 
               Height          =   390
               Index           =   0
               Left            =   75
               TabIndex        =   60
               ToolTipText     =   "메시지 추가..."
               Top             =   7410
               Width           =   435
               _Version        =   851970
               _ExtentX        =   767
               _ExtentY        =   688
               _StockProps     =   79
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
               Picture         =   "frm행사안내용문자.frx":2174A
            End
            Begin XtremeSuiteControls.PushButton cmdChange 
               Height          =   390
               Left            =   1695
               TabIndex        =   61
               ToolTipText     =   "암호변경..."
               Top             =   7410
               Width           =   435
               _Version        =   851970
               _ExtentX        =   767
               _ExtentY        =   688
               _StockProps     =   79
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
               Picture         =   "frm행사안내용문자.frx":21CE4
            End
            Begin XtremeSuiteControls.PushButton btnMove 
               Height          =   360
               Index           =   0
               Left            =   3570
               TabIndex        =   62
               Top             =   5280
               Width           =   825
               _Version        =   851970
               _ExtentX        =   1455
               _ExtentY        =   635
               _StockProps     =   79
               Caption         =   "< 이전"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
            End
            Begin Threed.SSPanel pnlSMS 
               Height          =   1680
               Index           =   0
               Left            =   75
               TabIndex        =   63
               Top             =   75
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   2963
               _Version        =   262144
               PictureFrames   =   1
               Picture         =   "frm행사안내용문자.frx":2227E
               BorderWidth     =   0
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin VB.Label lblSMSMsg 
                  BackStyle       =   0  '투명
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1125
                  Index           =   0
                  Left            =   105
                  TabIndex        =   65
                  Top             =   360
                  Width           =   1500
               End
               Begin VB.Label lblNo 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   180
                  Index           =   0
                  Left            =   135
                  TabIndex        =   64
                  Top             =   60
                  Width           =   90
               End
            End
            Begin Threed.SSPanel pnlSMS 
               Height          =   1680
               Index           =   1
               Left            =   1800
               TabIndex        =   66
               Top             =   75
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   2963
               _Version        =   262144
               PictureFrames   =   1
               Picture         =   "frm행사안내용문자.frx":26E93
               BorderWidth     =   0
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin VB.Label lblSMSMsg 
                  BackStyle       =   0  '투명
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1125
                  Index           =   1
                  Left            =   105
                  TabIndex        =   68
                  Top             =   360
                  Width           =   1500
               End
               Begin VB.Label lblNo 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   180
                  Index           =   1
                  Left            =   135
                  TabIndex        =   67
                  Top             =   60
                  Width           =   90
               End
            End
            Begin Threed.SSPanel pnlSMS 
               Height          =   1680
               Index           =   2
               Left            =   3525
               TabIndex        =   69
               Top             =   75
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   2963
               _Version        =   262144
               PictureFrames   =   1
               Picture         =   "frm행사안내용문자.frx":2BAA8
               BorderWidth     =   0
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin VB.Label lblSMSMsg 
                  BackStyle       =   0  '투명
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1125
                  Index           =   2
                  Left            =   105
                  TabIndex        =   71
                  Top             =   360
                  Width           =   1500
               End
               Begin VB.Label lblNo 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   180
                  Index           =   2
                  Left            =   135
                  TabIndex        =   70
                  Top             =   60
                  Width           =   90
               End
            End
            Begin Threed.SSPanel pnlSMS 
               Height          =   1680
               Index           =   3
               Left            =   75
               TabIndex        =   72
               Top             =   1770
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   2963
               _Version        =   262144
               PictureFrames   =   1
               Picture         =   "frm행사안내용문자.frx":306BD
               BorderWidth     =   0
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin VB.Label lblSMSMsg 
                  BackStyle       =   0  '투명
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1125
                  Index           =   3
                  Left            =   105
                  TabIndex        =   74
                  Top             =   360
                  Width           =   1500
               End
               Begin VB.Label lblNo 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   180
                  Index           =   3
                  Left            =   135
                  TabIndex        =   73
                  Top             =   60
                  Width           =   90
               End
            End
            Begin Threed.SSPanel pnlSMS 
               Height          =   1680
               Index           =   4
               Left            =   1800
               TabIndex        =   75
               Top             =   1770
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   2963
               _Version        =   262144
               PictureFrames   =   1
               Picture         =   "frm행사안내용문자.frx":352D2
               BorderWidth     =   0
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin VB.Label lblSMSMsg 
                  BackStyle       =   0  '투명
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1125
                  Index           =   4
                  Left            =   105
                  TabIndex        =   77
                  Top             =   360
                  Width           =   1500
               End
               Begin VB.Label lblNo 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   180
                  Index           =   4
                  Left            =   135
                  TabIndex        =   76
                  Top             =   60
                  Width           =   90
               End
            End
            Begin Threed.SSPanel pnlSMS 
               Height          =   1680
               Index           =   5
               Left            =   3525
               TabIndex        =   78
               Top             =   1770
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   2963
               _Version        =   262144
               PictureFrames   =   1
               Picture         =   "frm행사안내용문자.frx":39EE7
               BorderWidth     =   0
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin VB.Label lblSMSMsg 
                  BackStyle       =   0  '투명
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1125
                  Index           =   5
                  Left            =   105
                  TabIndex        =   80
                  Top             =   360
                  Width           =   1500
               End
               Begin VB.Label lblNo 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   180
                  Index           =   5
                  Left            =   135
                  TabIndex        =   79
                  Top             =   60
                  Width           =   90
               End
            End
            Begin Threed.SSPanel pnlSMS 
               Height          =   1680
               Index           =   6
               Left            =   75
               TabIndex        =   81
               Top             =   3465
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   2963
               _Version        =   262144
               PictureFrames   =   1
               Picture         =   "frm행사안내용문자.frx":3EAFC
               BorderWidth     =   0
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin VB.Label lblSMSMsg 
                  BackStyle       =   0  '투명
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1125
                  Index           =   6
                  Left            =   105
                  TabIndex        =   83
                  Top             =   360
                  Width           =   1500
               End
               Begin VB.Label lblNo 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   180
                  Index           =   6
                  Left            =   135
                  TabIndex        =   82
                  Top             =   60
                  Width           =   90
               End
            End
            Begin Threed.SSPanel pnlSMS 
               Height          =   1680
               Index           =   7
               Left            =   1800
               TabIndex        =   84
               Top             =   3465
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   2963
               _Version        =   262144
               PictureFrames   =   1
               Picture         =   "frm행사안내용문자.frx":43711
               BorderWidth     =   0
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin VB.Label lblSMSMsg 
                  BackStyle       =   0  '투명
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1125
                  Index           =   7
                  Left            =   105
                  TabIndex        =   86
                  Top             =   360
                  Width           =   1500
               End
               Begin VB.Label lblNo 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   180
                  Index           =   7
                  Left            =   135
                  TabIndex        =   85
                  Top             =   60
                  Width           =   90
               End
            End
            Begin XtremeSuiteControls.PushButton btnMove 
               Height          =   360
               Index           =   1
               Left            =   4410
               TabIndex        =   87
               Top             =   5280
               Width           =   825
               _Version        =   851970
               _ExtentX        =   1455
               _ExtentY        =   635
               _StockProps     =   79
               Caption         =   "다음 >"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
            End
            Begin Threed.SSPanel pnlSMS 
               Height          =   1680
               Index           =   8
               Left            =   3525
               TabIndex        =   89
               Top             =   3465
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   2963
               _Version        =   262144
               PictureFrames   =   1
               Picture         =   "frm행사안내용문자.frx":48326
               BorderWidth     =   0
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin VB.Label lblNo 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   180
                  Index           =   8
                  Left            =   135
                  TabIndex        =   91
                  Top             =   60
                  Width           =   90
               End
               Begin VB.Label lblSMSMsg 
                  BackStyle       =   0  '투명
                  BeginProperty Font 
                     Name            =   "굴림체"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1125
                  Index           =   8
                  Left            =   105
                  TabIndex        =   90
                  Top             =   360
                  Width           =   1500
               End
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   $"frm행사안내용문자.frx":4CF3B
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   720
               Index           =   8
               Left            =   2190
               TabIndex        =   105
               ToolTipText     =   $"frm행사안내용문자.frx":4CFD6
               Top             =   5730
               Width           =   4680
            End
            Begin VB.Line Line1 
               X1              =   3000
               X2              =   2925
               Y1              =   5295
               Y2              =   5580
            End
            Begin VB.Label lblPageCount 
               Alignment       =   2  '가운데 맞춤
               BackStyle       =   0  '투명
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   3045
               TabIndex        =   93
               Top             =   5325
               Width           =   300
            End
            Begin VB.Label lblPage 
               Alignment       =   2  '가운데 맞춤
               BackStyle       =   0  '투명
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   2565
               TabIndex        =   92
               Top             =   5325
               Width           =   300
            End
            Begin VB.Label lblTitle 
               BackColor       =   &H00FFC0C0&
               BackStyle       =   0  '투명
               Caption         =   "보내는 사람"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   2865
               TabIndex        =   88
               Top             =   6480
               Width           =   1170
            End
            Begin VB.Image Image 
               Height          =   240
               Index           =   2
               Left            =   2535
               Picture         =   "frm행사안내용문자.frx":4D071
               Top             =   6465
               Width           =   240
            End
         End
      End
   End
End
Attribute VB_Name = "frm행사안내용문자"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_Host_DataBase     As ADODB.Connection
Dim m_Connect           As Boolean
Dim FORM_SMS001_ACTIVATE    As Boolean
Dim bCountFlag          As Boolean
Dim bSSChangeFlag       As Boolean
Dim bSSChangeFlag2      As Boolean

Dim sSmsMsg As String
Const SMS_Lng = 55

Private Sub btnSelect_Click(Index As Integer)
    Dim lRow    As Long
    Dim sTel(2) As String
    
    Select Case Index
        Case 0
            For lRow = 1 To sprGrid.MaxRows
                sprGrid.Row = lRow
                
                sprGrid.Col = 4 ' 휴대폰 번호가 있을경우
                If sprGrid.Text <> "" Then
                
                    sprGrid.Col = 8 ' 해당 회원의 정보를 얻어온다.
                    
                    If Get_고객정보(sprGrid.Text) <> "Error" Then
                        If CheckMobileNumber(고객정보.휴대전화, sTel) = True Then
                            If 고객정보.SMS전송여부 = "N" Then
                                sprGrid.Col = -1: sprGrid.BackColor = vbRed
                            Else
                                sprGrid.SetText 1, lRow, "1"
                                
                                ' 선택 수량 누적
                                bCountFlag = False
                                'txtCount(1).value = txtCount(1).value + 1
                            End If
                        End If
                    End If
                End If
            Next lRow
            
            txtCount(1).Value = CStr(GetSelectSpread(sprGrid, 1))
        
        Case 1
            For lRow = 1 To sprGrid.MaxRows
                sprGrid.Row = lRow
            
                sprGrid.Col = 1:
                If sprGrid.Value = 1 Then
                    sprGrid.SetText 1, lRow, "0"
                End If
            Next lRow
            
            txtCount(1).Value = CStr(GetSelectSpread(sprGrid, 1))
            Exit Sub
            
        Case 2
            If CheckConnect = True Then
                Call SetUseSMSCount ' 최종 남은 수량을 설정한다.
            End If
        
        
    End Select
End Sub

'Private Sub cboSendText_Click()
'    If cboSendText.ListIndex >= 0 Then
'        txtSMS.Text = cboSendText.Text
'        txtSend2.Text = cboSendText.Text
'    End If
'End Sub

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 1
            If CheckConnect = False Then Exit Sub
            Call SendSMS ' 발송
        
        Case 4 ' 개별 문자 메시지
            If cmdBtn(Index).Caption = "개별발송" Then
                cmdBtn(Index).Caption = "그룹발송"
                cmdBtn(1).Visible = False
            Else
                cmdBtn(Index).Caption = "개별발송"
                cmdBtn(1).Visible = True
            End If
            
            'pnlSend(0).Visible = Not pnlSend(0).Visible
            
        Case 5: Unload Me
    End Select
End Sub

'+------------------------------------------------------
'+ 2003/02/11 수정
'+
'+루틴설명      - 비밀번호확인
'+  1. 암호를 확인하여 암호 규칙에 맞으면 화면을 종료한다.
'+  2. 레지스터리에 저장한다.
'+
'+------------------------------------------------------
Private Sub cmdChange_Click()
    Dim strPass As String
    Dim bPass   As Boolean
    
    ' 입력 확인
    bPass = False
    
    strPass = InputBox("암호를 입력하여 주십시요", "SMS 암호")
    If Len(strPass) <= 0 Then
        Exit Sub
    End If
    
'   기본 디폴드 암호.. ( 프로그램 셋팅/설치를 위한 암호 )
    If UCase(strPass) = "DUDTJSGH" Or UCase(strPass) = "SHOP500" Then
        bPass = True
    Else
        ' 비밀번호 확인
        strPass = IsPassWord(strPass)
        If strPass = "-1" Or strPass = "-3" Then
            If strPass = "-3" Then MsgBox "입력한 내용이 정확하지 않습니다.", vbCritical, "입력오류"
            Exit Sub
        Else
            bPass = True
        End If
    End If
    
    ' 암호를 입력 하였으면
    If bPass = True Then
        m_SMS_EMART_PASS = True
        cmdBtn(4).Enabled = True
        txtSMS.Enabled = True
        txtSMS.Locked = False
        
    End If
End Sub

Private Sub cmdList_Click()
    Call Data_Display
End Sub

Private Sub cmdSend_Click()
    Dim nIndex      As Integer
    Dim nSendCount  As Integer
    Dim sSendTel(2) As String
    Dim sRecvTel() As String
    Dim sValue(10)  As String
    Dim lRow        As Long
    Dim sMsg        As String
    
    On Error GoTo ErrRtn
    
    If CheckConnect = False Then Exit Sub
    
    If txtCount(0).Value <= 0 Then
        MsgBox "사용 가능 여부및 수량을 확인 하여 주십시요.", vbInformation, "확인"
        Exit Sub
    End If
    
    ' 전송 시간 확인
    If "18:00" < Format(Time, "hh:mm") And 가맹점정보.SMS_EMART = "Y" Then
        MsgBox "이마트에서는 18:00 이후에는 문자 메시지 발송을 할 수 없습니다.", vbInformation, "확인"
        Exit Sub
    End If
    
    If Val(lblLan2.Tag) > m_SMS_Lng Then
        MsgBox "작성된 메시지가 " & CStr(m_SMS_Lng) & "자 이상 입니다. " & CStr(m_SMS_Lng) & "자 이상은 전송할 수 없습니다.", vbCritical, "확인"
        Exit Sub
        
    ElseIf Val(lblLan2.Tag) <= 0 Then
        MsgBox "메시지를 확인하여 주십시요.", vbInformation, "확인"
        Exit Sub
    End If
    
    ' 광고 내용 확인
    If InStr(txtSend2.Text, "세일") > 0 Or InStr(txtSend2.Text, "행사") > 0 Or InStr(txtSend2.Text, "할인") > 0 Or InStr(txtSend2.Text, "광고") > 0 Then
        sMsg = "정보 통신망 이용 촉진 및 정보 보호 등에 관한 법률로 인한 안내" & vbNewLine
        sMsg = sMsg & vbNewLine
        sMsg = sMsg & " - 해당 화면에서는 세일, 광고, 할인, 행사와 관련된 문자를 발송할 수 없습니다." & vbNewLine
        
        Call MsgBox(sMsg, vbInformation, "확인")
        Exit Sub
    End If
    
    ' 발신자 번호 확인
    If GetCheckSMSSendTel(txtRecvTel.Text, sRecvTel, True) = False Then
        sMsg = "전기통신사업법 제84조에 의하여 문자 발신 번호는 반드시 입력 하여야 합니다." & vbNewLine
        sMsg = sMsg & vbNewLine
        sMsg = sMsg & " 주요내용" & vbNewLine
        sMsg = sMsg & " - 발신번호 없이 문자 전송 불가" & vbNewLine
        sMsg = sMsg & " - 발신번호는 수신자가 실제 발신(통화)이 가능한 번호만 허용" & vbNewLine
        sMsg = sMsg & " - 일반번호의 경우 지역번호(02,031등)를 앞자리에 포함한 번호만 허용" & vbNewLine
        sMsg = sMsg & " - 대표번호는 8자리만 입력 허용되며,내선번호 포함불가" & vbNewLine
        sMsg = sMsg & " - 030 번호의 경우 12자리까지 허용" & vbNewLine
        
        Call MsgBox(sMsg, vbInformation, "입력확인")
        Exit Sub
    
    End If
    
    cmdSend.Enabled = False
    DoEvents
    
    ' 입력 전화 번호 확인
    nSendCount = 0
    For nIndex = 0 To txtMobile.Count - 1
        txtMobile(nIndex).Text = Trim(txtMobile(nIndex).Text)
        If txtMobile(nIndex).Text <> "" Then
            If CheckMobileNumber(txtMobile(nIndex).Text, sSendTel) = False Then
                MsgBox txtMobile(nIndex).Text & " 전화 번호로는 문자 메시지를 보낼 수 없습니다.", vbInformation, "확인"
                
                txtMobile(nIndex).SetFocus
                txtMobile(nIndex).SelStart = 0: txtMobile(nIndex).SelLength = Len(txtMobile(nIndex).Text)
                
                cmdSend.Enabled = True
                Exit Sub
            Else
                nSendCount = nSendCount + 1
            End If
        End If
    Next nIndex
    
    
    ' 잔여수량보다 선택 수량이 더 많을 경우
    If txtCount(0).Value < nSendCount Then
        MsgBox "SMS 잔여 수량보다 더 많이 선택되었습니다." & vbLf & vbLf & " 전송 수량을 조절 하여 주십시요.", vbInformation, "확인"
        
        cmdSend.Enabled = True
        Exit Sub
    End If
    
    pnlProg.Visible = True
    pnlProg.Caption = "메시지를 전송 중 입니다. 잠시만 기다려 주십시요."
    Screen.MousePointer = vbHourglass
    DoEvents
    
    nSendCount = 0
    
    For nIndex = 0 To txtMobile.Count - 1
        ' 입력된 전화 번호가 있을 경우
        If txtMobile(nIndex).Text <> "" Then
            ' 전화 번호 검사및 휴대폰 번호 분리
            If CheckMobileNumber(txtMobile(nIndex).Text, sSendTel) = False Then
                MsgBox txtMobile(nIndex).Text & " 전화 번호로는 문자 메시지를 보낼수 없습니다.", vbInformation, "확인"
                
                txtMobile(nIndex).SetFocus
                txtMobile(nIndex).SelStart = 0: txtMobile(nIndex).SelLength = Len(txtMobile(nIndex).Text)
                
                cmdSend.Enabled = True
                Exit Sub
                
            ' 발송 가능한 번호일 경우 발송 한다.
            Else
                ' 전송, 메시지타입, 수신번호, 발신번호, 메시지, 지사코드, 가맹점코드, 고객코드, 고객성명, 참고5, 참고6
                sValue(0) = "1"
                sValue(1) = "0"
                sValue(2) = txtMobile(nIndex).Text
                sValue(3) = txtRecvTel.Text
                sValue(4) = Trim(txtSend2.Text)
                sValue(5) = 가맹점정보.지사코드
                sValue(6) = 가맹점정보.택코드
                sValue(7) = " "
                sValue(8) = " "
                sValue(9) = 가맹점정보.가맹점코드
                sValue(10) = "2"
                
'                Query = "EXEC PRO_SMS_SEND "
'                Query = Query & "'" & sValue(0) & "', "
'                Query = Query & "'" & sValue(1) & "', "
'                Query = Query & "'" & sValue(2) & "', "
'                Query = Query & "'" & sValue(3) & "', "
'                Query = Query & "'" & sValue(4) & "', "
'                Query = Query & "'" & sValue(5) & "', "
'                Query = Query & "'" & sValue(6) & "', "
'                Query = Query & "'" & sValue(7) & "', "
'                Query = Query & "'" & sValue(8) & "', "
'                Query = Query & "'" & sValue(9) & "', "
'                Query = Query & "'" & sValue(10) & "' "
'
'                Debug.Print Query
'                If Dir(App.Path & "\NO_SMS.DAT", vbNormal) = "" Then
'                    m_Host_DataBase.Execute Query
'                    nSendCount = nSendCount + 1
'                End If
                send_Purio_SMS "", txtMobile(nIndex).Text, Trim(txtSend2.Text), "1"
                nSendCount = nSendCount + 1
            End If
        End If
    Next nIndex
    
    ' 최종 남은 수량을 설정한다.
    Call SetUseSMSCount

    pnlProg.Visible = False
    MsgBox CStr(nSendCount) & " 건 발송 되었습니다.   ", vbInformation, "확인"
    Screen.MousePointer = vbDefault
    
    cmdSend.Enabled = True
    
    Exit Sub

ErrRtn:
    cmdSend.Enabled = True
    
    Screen.MousePointer = vbDefault
    
    pnlProg.Visible = False
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub cmdSendTextSave_Click(Index As Integer)
    Dim iCode As String
    
    If 가맹점정보.SMS_EMART = "Y" And m_SMS_EMART_PASS = False Then
        MsgBox "이마트 매장에서는 수정 불가능 합니다.", vbInformation, "확인"
        Exit Sub
    End If
    
    Select Case Index
        Case 0 ' 추가
            Query = "SELECT ISNULL(MAX(순번),0) + 1 FROM TB_문자발송문"
            Set SUBRs = New ADODB.RecordSet
            SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
            iCode = Format(SUBRs(0), "00")
            
            SUBRs.Close
            Set SUBRs = Nothing
            
            '-----------------------------------------------------------
            '
            '-----------------------------------------------------------
            Query = "SELECT * FROM TB_문자발송문"
            Query = Query & " WHERE 순번 = '" & iCode & "'"
            Set SUBRs = New ADODB.RecordSet
            SUBRs.Open Query, ADOCon, adOpenDynamic, adLockOptimistic
            
            If SUBRs.EOF Then SUBRs.AddNew
            
            SUBRs!순번 = iCode & ""
            SUBRs!내용 = Trim(txtSMS.Text) & ""
            
            SUBRs.Update
            
            SUBRs.Close
            Set SUBRs = Nothing
            
        Case 1  ' 수정
            Query = "UPDATE TB_문자발송문 SET "
            Query = Query & " 내용 = '" & Trim(txtSMS.Text) & "'"
            Query = Query & " WHERE 순번 = '" & Format(lblNum.Caption, "00") & "'"
            ADOCon.Execute Query
        
        Case 2  ' 삭제
            'Query = "DELETE FROM TB_문자발송문"
            'Query = Query & " WHERE 순번 = '" & Format(lblNum.Caption, "00") & "'"
            
            Query = "UPDATE TB_문자발송문 SET "
            Query = Query & " 내용 = ''"
            Query = Query & " WHERE 순번 = '" & Format(lblNum.Caption, "00") & "'"
            ADOCon.Execute Query
    End Select
    
    'Call ReadSendTextMessage(cboSendText)
    Call 문자메시지_Display(Int(lblPage.Caption), (lblNo(0).Caption))
End Sub

Private Sub cmdSvr_Click(Index As Integer)
    Select Case Index
        Case 2
            If SaveConnectData = True Then
                MsgBox "저장 완료", vbInformation
            Else
                MsgBox "저장 실패", vbCritical
            End If
            
        Case 1
            Call SaveConnectData ' 저장을 먼저 처리한다.
            
            If CheckConnect = True Then
                MsgBox "연결 완료", vbInformation
            End If
            
        Case 0: Call DefaultServerSetting
    End Select
End Sub

Private Sub Form_Activate()
    On Error GoTo ErrRtn

    If FORM_SMS001_ACTIVATE = True Then Exit Sub
    
    FORM_SMS001_ACTIVATE = True
    
    
        
    sSmsMsg = "정보 통신망 이용 촉진 및 정보 보호 등에 관한 법률로 인하여" & vbNewLine & vbNewLine
    sSmsMsg = sSmsMsg & "(광고)                                <---- 첫번째 줄에 자동 추가" & vbNewLine
    sSmsMsg = sSmsMsg & "무료수신거부 080-863-5771    <---- 마지막 줄에 자동 추가" & vbNewLine & vbNewLine
    sSmsMsg = sSmsMsg & "로 인하여  55자 까지만 발송이 가능 합니다." & vbNewLine
        
    
    
    TabControl.SelectedItem = 0
    cboView.ListIndex = 0
 
    Text1.Text = 가맹점정보.전화SMS & "" '보내는 사람
    
    DoEvents
    Call DefaultServerSetting ' 기본 설정으로 본사에 연결한다.
    
    If CheckConnect = True Then
        Call SetUseSMSCount ' 최종 남은 수량을 설정한다.
    End If

    Exit Sub

ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    With sprGrid
        .MaxRows = 0
        .RowHeight(-1) = 14
        
        .Col = 8: .ColHidden = True
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeNormal
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
    
    dtpDay(0).Value = DateAdd("d", -2, Date)
    dtpDay(1).Value = DateAdd("d", -2, Date)
    
    txtCount(0).Value = 0
    txtCount(1).Value = 0
    txtCount(2).Value = 0
    
    sSmsMsg = "정보 통신망 이용 촉진 및 정보 보호 등에 관한 법률로 인하여" & vbNewLine & vbNewLine
    sSmsMsg = sSmsMsg & "(광고)                                <---- 첫번째 줄에 자동 추가" & vbNewLine
    sSmsMsg = sSmsMsg & "무료수신거부 080-863-5771    <---- 마지막 줄에 자동 추가" & vbNewLine & vbNewLine
    sSmsMsg = sSmsMsg & "로 인하여  55자 까지만 발송이 가능 합니다." & vbNewLine
    
    txtRecvTel.Text = 가맹점정보.전화SMS
    Text1.Text = 가맹점정보.전화SMS
    
    'Call ReadSendTextMessage(cboSendText)
    Call 문자메시지_Display
    
    'TitleSet "행사 안내용 문자"
    
    cmdBtn(4).Enabled = IIf(가맹점정보.SMS_EMART = "Y", False, True)
    cmdBtn(4).Enabled = IIf(가맹점정보.SMS_EMART = "N", True, m_SMS_EMART_PASS)
    
    'txtSMS.Enabled = cmdBtn(4).Enabled
    
    If cmdBtn(4).Enabled = False Then
        txtSMS.Locked = True
    End If
    
    cmdChange.Enabled = Not cmdBtn(4).Enabled
    
    Exit Sub

ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub 문자메시지_Display(Optional iPage As Integer = 1, Optional iNum As Integer = 0)
    On Error GoTo ErrRtn
    
    For i = 0 To 8
        lblNo(i).Caption = ""
        lblSMSMsg(i).Caption = ""
    Next i
    
    If 가맹점정보.SMS_EMART = "Y" Then
        lblPage.Caption = "1"
        lblPageCount.Caption = "1"
        
        btnMove(0).Enabled = False
        btnMove(1).Enabled = False
        
        cmdSendTextSave(0).Enabled = False
        cmdSendTextSave(1).Enabled = False
        cmdSendTextSave(2).Enabled = False
                
        lblNo(0).Caption = "0"
        lblSMSMsg(0).Caption = "고객님 세탁물이 도착했습니다. 인수 부탁 드립니다. 즐겁고 행복한 하루 되세요."
        
        lblNo(1).Caption = "1"
        lblSMSMsg(1).Caption = "고객님 세탁물을 보관 중입니다. 인수 부탁 드립니다. 즐겁고 행복한 하루 되세요."
        
        lblNo(2).Caption = "2"
        lblSMSMsg(2).Caption = "고객님 세탁물을 보관 중입니다. 보관 중 먼지가 쌓일 수 있으니 빠른 인수 바랍니다."
        
        lblNo(3).Caption = "3"
        lblSMSMsg(3).Caption = "고객님 크린에이드를 이용해 주셔서 감사 드립니다. 즐겁고 행복한 하루 되세요."
        
        lblNo(4).Caption = "4"
        lblSMSMsg(4).Caption = "고객님 금일 크린에이드 이용중에 불편 드렸던 점 진심으로 사과 드립니다."
        
        lblNo(5).Caption = "5"
        lblSMSMsg(5).Caption = "고객님 접수하신 세탁물이 다소 지연되고 있습니다. 도착하는 대로 연락 드리겠습니다"
        
        lblNo(6).Caption = "6"
        lblSMSMsg(6).Caption = "고객님 접수하신 세탁물이 반품되어 확인이 필요합니다. 내방 부탁 드립니다."
        
        lblNo(7).Caption = "7"
        lblSMSMsg(7).Caption = "고객님 죄송합니다.^^ 고객님께 메시지전송이 오류로 발송되었습니다. 감사합니다."
        
        lblNo(8).Caption = "8"
        lblSMSMsg(8).Caption = ""
        
    Else
        lblPage.Caption = iPage
        
        If iPage = 1 Then
            Query = "SELECT COUNT(*) FROM TB_문자발송문"
            Set SUBRs = New ADODB.RecordSet
            SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            If SUBRs(0) = 0 Then
                lblPageCount.Caption = "1"
            Else
                If SUBRs(0) Mod 9 = 0 Then
                    lblPageCount.Caption = Int(SUBRs(0) / 9)
                Else
                    lblPageCount.Caption = Int(SUBRs(0) / 9) + 1
                End If
            End If
            SUBRs.Close
            Set SUBRs = Nothing
            
            '
            Query = "SELECT * FROM TB_문자발송문"
            Query = Query & " ORDER BY 순번 ASC"
            Set SUBRs = New ADODB.RecordSet
            SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
            If SUBRs.EOF Then
                SUBRs.Close
                Set SUBRs = Nothing
                
                Query = "INSERT INTO TB_문자발송문 VALUES('01', '전할 메시지를 입력하여 주십시요')"
                ADOCon.Execute Query
            
                '--------------------------------------------
                Query = "SELECT * FROM TB_문자발송문"
                Query = Query & " ORDER BY 순번 ASC"
                Set SUBRs = New ADODB.RecordSet
                SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            End If
        Else
            Query = "SELECT TOP (9) * FROM TB_문자발송문"
            Query = Query & " WHERE 순번 >= '" & Format(iNum, "00") & "'"
            Query = Query & " ORDER BY 순번 ASC"
            Set SUBRs = New ADODB.RecordSet
            SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        End If
        
        i = 0
        
        Do Until SUBRs.EOF
            If i > 8 Then Exit Do
            
            lblSMSMsg(i).Caption = Trim(SUBRs!내용) & ""
            lblNo(i).Caption = Val(SUBRs!순번) & ""
            
            i = i + 1
            
            SUBRs.MoveNext
        Loop
        SUBRs.Close
        Set SUBRs = Nothing
    End If
    
    lblNum.Caption = lblNo(0).Caption
    txtSMS.Text = lblSMSMsg(0).Caption & ""
    
    Exit Sub

ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    cmdBtn(5).Left = (Me.Width - 200) - cmdBtn(5).Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FORM_SMS001_ACTIVATE = False
End Sub

Private Sub lblSMSMsg_Click(Index As Integer)
    lblNum.Caption = lblNo(Index).Caption
    
    txtSMS.Text = lblSMSMsg(Index).Caption & ""
    txtSend2.Text = lblSMSMsg(Index).Caption & ""
End Sub

Private Sub btnMove_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    If 가맹점정보.SMS_EMART = "Y" Then Exit Sub
        
    If Index = 0 Then
        If lblPage.Caption = "1" Then Exit Sub
        
        lblPage.Caption = Int(lblPage.Caption) - 1
        
        Query = "SELECT TOP (9) * FROM TB_문자발송문"
        Query = Query & " WHERE 순번 < '" & Format(lblNo(0).Caption, "00") & "'"
        Query = Query & " ORDER BY 순번 DESC"
        Set SUBRs = New ADODB.RecordSet
        SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
        For i = 0 To 8
            lblNo(i).Caption = ""
            lblSMSMsg(i).Caption = ""
        Next i
        
        i = 8

        Do Until SUBRs.EOF
            If i < 0 Then Exit Do
            
            lblNo(i).Caption = Val(SUBRs!순번) & ""
            lblSMSMsg(i).Caption = Trim(SUBRs!내용) & ""
            
            i = i - 1
            
            SUBRs.MoveNext
        Loop
        SUBRs.Close
        Set SUBRs = Nothing
        
    Else
        If lblPage.Caption = lblPageCount.Caption Then Exit Sub
        
        If lblNo(8).Caption = "" Then Exit Sub
        
        lblPage.Caption = Int(lblPage.Caption) + 1
        
        '
        Query = "SELECT TOP (9) * FROM TB_문자발송문"
        Query = Query & " WHERE 순번 > '" & Format(lblNo(8).Caption, "00") & "'"
        Query = Query & " ORDER BY 순번 ASC"
        Set SUBRs = New ADODB.RecordSet
        SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
        For i = 0 To 8
            lblNo(i).Caption = ""
            lblSMSMsg(i).Caption = ""
        Next i
        
        i = 0

        Do Until SUBRs.EOF
            If i > 8 Then Exit Do
            
            lblNo(i).Caption = Val(SUBRs!순번) & ""
            lblSMSMsg(i).Caption = Trim(SUBRs!내용) & ""
            
            i = i + 1
            
            SUBRs.MoveNext
        Loop
        SUBRs.Close
        Set SUBRs = Nothing
    End If

    lblNum.Caption = lblNo(0).Caption
    txtSMS.Text = lblSMSMsg(0).Caption & ""
    
    Exit Sub

ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub PushButton1_Click()
    SSPanel_SMS.ZOrder 0
    SSPanel_SMS.Visible = Not SSPanel_SMS.Visible
End Sub

Private Sub txtCount_Change(Index As Integer)
    ' 선택 수량이 변경될 경우 전송후 수량을 수정하여 준다.
    
    Debug.Print bCountFlag
    
    If Index = 1 Then
        If bCountFlag = True Then Exit Sub
        
        txtCount(2).Value = txtCount(0).Value - (txtCount(1).Value * SMS_PRICE)
        bCountFlag = True
    End If
End Sub

Private Sub sprGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    If bSSChangeFlag = True Or bSSChangeFlag2 = True Then Exit Sub
    
    If Col = 1 Then
        Dim sTel(2) As String
        
        ' 선택을 해지 하면 선택 수량에서 -1를 해준다.
        If ButtonDown = 0 Then
            bCountFlag = False
            
            txtCount(1).Value = txtCount(1).Value - 1
        Else
            bSSChangeFlag2 = True
            sprGrid.Row = Row
            sprGrid.Col = 4
            
            If CheckMobileNumber(sprGrid.Text, sTel) = True Then
                bCountFlag = False
                txtCount(1).Value = txtCount(1).Value + 1
            Else
                MsgBox "선택된 전화번호로는 문자 메시지를 보낼수 없습니다.", vbInformation, "확인"
                sprGrid.Row = Row
                sprGrid.Col = 1: sprGrid.Value = 0
            End If
            
            bSSChangeFlag2 = False
        End If
    End If
End Sub
 
Private Sub txtMobile_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        
        SendKeys "{Tab}"
    End If
End Sub

Private Sub txtSend2_Change()
    lblLan2.Tag = CStr(LenB(StrConv(txtSend2.Text, vbFromUnicode)))
    lblLan2.Caption = lblLan2.Tag & "자"
    
    If LenB(StrConv(txtSend2.Text, vbFromUnicode)) > m_SMS_Lng Then
        MsgBox "작성된 메시지가 " & CStr(m_SMS_Lng) & "자 이상 입니다. " & CStr(m_SMS_Lng) & "자 이상은 전송할 수 없습니다.", vbCritical, "확인"
        lblLan2.BackColor = vbRed
        Exit Sub
    Else
        lblLan2.BackColor = &HC0C0FF
    End If
    
End Sub

Private Sub txtSend2_KeyPress(KeyAscii As Integer)
    ' '입력이 안되도록 수정
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtSMS_Change()
    lbl_SMS.Tag = CStr(LenB(StrConv(txtSMS.Text, vbFromUnicode)))
    lbl_SMS.Caption = lbl_SMS.Tag & "자"
    Debug.Print lbl_SMS.Tag & "자"

    If LenB(StrConv(txtSMS.Text, vbFromUnicode)) > SMS_Lng Then
        
        MsgBox sSmsMsg, vbInformation, "확인"
        
        lbl_SMS.BackColor = vbRed
        Exit Sub
    Else
        lbl_SMS.BackColor = Me.BackColor
    End If
End Sub

Private Sub DefaultServerSetting()
    ' 기본 설정 정보가 없을 경우
    On Error GoTo ErrRtn
    
    Query = "SELECT    ISNULL(SMS_IP,'115.89.220.5,8657') AS SMS_IP"
    Query = Query & ", ISNULL(SMS_DB,'Laundry1000') AS SMS_DB"
    Query = Query & ", ISNULL(SMS_ID,'sa') AS SMS_ID"
    Query = Query & ", ISNULL(SMS_PWD,'cleanaid1996!@#') AS SMS_PWD"
    Query = Query & ", ISNULL(TIMEOUT,30) AS TIMEOUT"
    Query = Query & " FROM TB_기본정보"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF Then
        txtServer(0).Text = "115.89.220.5,8657" '
        txtServer(1).Text = "Laundry1000"                    '
        txtServer(2).Text = "sa"                         '
        txtServer(3).Text = "cleanaid1996!@#"                           '
        m_CommandTimeOut = 30                            '
    Else
        txtServer(0).Text = Trim(ADORs!SMS_IP) & ""      '
        txtServer(1).Text = Trim(ADORs!SMS_DB) & ""      '
        txtServer(2).Text = Trim(ADORs!SMS_ID) & ""      '
        txtServer(3).Text = Trim(ADORs!SMS_PWD) & ""     '
        m_CommandTimeOut = Val(Trim(ADORs!timeout) & "") '
    End If
    ADORs.Close
    Set ADORs = Nothing
    
    Exit Sub

ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub


'-------------------------------------------------------------------------------------
' 내용 구분의 정보를 저장 한다.
'-------------------------------------------------------------------------------------
Private Function SaveConnectData() As Boolean
    Dim Query    As String
    
    On Error GoTo ErrRtn
    
    SaveConnectData = False
    
    txtServer(0).Text = Trim(txtServer(0).Text)
    txtServer(1).Text = Trim(txtServer(1).Text)
    txtServer(2).Text = Trim(txtServer(2).Text)
    txtServer(3).Text = Trim(txtServer(3).Text)
    
    Query = "UPDATE TB_기본정보 SET "
    Query = Query & " SMS_IP = ' " & txtServer(0).Text & "', "
    Query = Query & " SMS_DB = ' " & txtServer(1).Text & "', "
    Query = Query & " SMS_ID = ' " & txtServer(2).Text & "', "
    Query = Query & " SMS_PWD = ' " & txtServer(3).Text & "' "
    ADOCon.Execute Query
    
    SaveConnectData = True
    
    Exit Function

ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Function


Private Function CheckConnect() As Boolean
    On Error GoTo ErrRtn
    
    Dim HostConn    As String
    
    HostConn = ""
    HostConn = HostConn & "Provider=SQLOLEDB.1;"
    HostConn = HostConn & "Persist Security Info=False;"
    HostConn = HostConn & "User ID=" & txtServer(2).Text & ";"
    HostConn = HostConn & "Password=" & txtServer(3).Text & ";"
    HostConn = HostConn & "Initial Catalog=" & txtServer(1).Text & ";"
    HostConn = HostConn & "Data Source=" & txtServer(0).Text
    m_CommandTimeOut = IIf(m_CommandTimeOut = 0, 30, m_CommandTimeOut)

    Set m_Host_DataBase = Nothing
    Set m_Host_DataBase = New ADODB.Connection
    
    pnlProg.Visible = True
    pnlProg.Caption = "서버에 연결중 입니다. 잠시만 기다려 주십시요..."
    DoEvents
    
    If m_Host_DataBase.State = adStateOpen Then m_Host_DataBase.Close
    
    m_Host_DataBase.ConnectionTimeout = 10
    m_Host_DataBase.CommandTimeout = m_CommandTimeOut
    m_Host_DataBase.Open HostConn
    
    pnlProg.Visible = False
    
    m_Connect = True
    
    CheckConnect = True
    
    Exit Function

ErrRtn:
    pnlProg.Visible = False
    
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Function



'--------------------------------------------------------------------------------------------------------------
' Procedure : SetUseSMSCount
' DateTime  : 2007-06-10 12:26
' Author    : pds2004
' Purpose   : 문자 메시지의 잔여 수량을 가저온다.
'--------------------------------------------------------------------------------------------------------------
Private Sub SetUseSMSCount()
'    Dim bResult     As Boolean
'    Dim ADORset     As New ADODB.RecordSet
'
'    ' 연결되어 있지 않을 경우 다시한번 연결을 시도한다.
'    On Error GoTo ErrRtn
'
'    txtCount(0).Value = "0"
'
'    If m_Connect = False Then
'        If CheckConnect = False Then
'            MsgBox "본사와의 연결 설정을 확인하여 주십시요", vbInformation, "확인"
'            ' 설정 화면을 활성화 한다.
'            Call cmdBtn_Click(2)
'            Exit Sub
'        End If
'    End If
'
'    Query = "EXEC PRO_SMS_STORE_001_01 '0', '" & 가맹점정보.가맹점코드 & "' "
'
'    ADORset.CursorLocation = adUseClient
'    ADORset.Open Query, m_Host_DataBase, adOpenStatic, adLockBatchOptimistic, adCmdText
'
'    If ADORset.EOF = False Then
'        If ADORset.RecordCount > 0 Then
'            txtCount(0).Value = ADORset!잔여수량
'        End If
'    End If
'    ADORset.Close
'    Set ADORset = Nothing
    
    txtCount(0).Value = GetMoney
    txtCount(2).Value = txtCount(0).Value - (txtCount(1).Value * SMS_PRICE)
    
    Exit Sub

ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub Data_Display()
    Dim FindCount As Integer
    Dim sTel(2)   As String
    Dim lRow      As Long
    Dim bResult   As Boolean
    Dim sData(1)  As String

    On Error GoTo ErrRtn
    
    Screen.MousePointer = vbHourglass
        
    pnlProg.Visible = True
    pnlProg.Caption = "고객정보를 조회중입니다..."
    DoEvents
    
    txtCount(1).Value = "0"
    txtCount(2).Value = "0"
    bSSChangeFlag = True
    
    
    '----------------------------------------------------------------------------------------
    ' 해당 일자에 입고된 고객 번호 내역을 구한다.
    '----------------------------------------------------------------------------------------
    If cboView.ListIndex = 0 Then
        Query = " SELECT   고객코드"
        Query = Query & ", MAX(접수일자) AS 접수일자"
        Query = Query & " FROM TB_입출고"
        Query = Query & " WHERE (접수일자 >= '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "'"
        Query = Query & "   AND  접수일자 <= '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "')"
        Query = Query & "   AND ((판매취소 <> 'Y')"
        Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
        Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
        Query = Query & " GROUP BY 고객코드"
        Query = Query & " ORDER BY 고객코드 ASC"
    
        sprGrid.SetText 5, -999, CVar("최종이용일")
    
    Else
        Query = " SELECT   고객코드"
        Query = Query & ", 등록일자 AS 접수일자"
        Query = Query & " FROM TB_고객정보"
        Query = Query & " WHERE (등록일자 >= '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "'"
        Query = Query & "   AND  등록일자 <= '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "')"
        Query = Query & " ORDER BY 고객코드 ASC"
        
        sprGrid.SetText 5, -999, CVar("등록일자")
    End If
    
    
    
    
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        'ButtonClicked 이벤트를 발생시키지 않도록한다. ButtonClicked 발생하면 선택수량이 이중계산됨...
        .EventEnabled(EventButtonClicked) = False
        
        Do Until ADORs.EOF
            If Get_고객정보(ADORs!고객코드) <> "Error" Then ' 해당 회원의 정보를 얻어온다.
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                
                .Col = 2: .Text = 고객정보.성명 & ""                   ' 2
                .Col = 3: .Text = 고객정보.전화번호 & ""               ' 3
                .Col = 4: .Text = 고객정보.휴대전화 & ""               ' 4
                .Col = 5: .Text = Format(ADORs!접수일자, "YYYY-MM-DD") ' 5
                .Col = 6: .Text = 고객정보.미수금액 & ""               ' 6
                .Col = 7: .Text = 고객정보.주소 & ""                   ' 7
                .Col = 8: .Text = 고객정보.고객코드 & ""               ' 8
                
                If CheckMobileNumber(고객정보.휴대전화, sTel) = True Then
                    If 고객정보.SMS전송여부 = "N" Then
                        .Col = 1: .Text = "0"                          ' 1
                        
                        .Col = -1: .BackColor = vbRed
                                   .ForeColor = vbWhite
                    Else
                        .Col = 1: .Text = "1"                          ' 1
                                                
                        bCountFlag = False ' 선택 수량 누적
                        
                        txtCount(1).Value = txtCount(1).Value + 1
                    End If
                Else
                    .Col = 1: .Text = "0"                          ' 1
                    
                    .Col = -1: .BackColor = vbYellow
                               .ForeColor = vbBlack
                End If
            End If
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        .EventEnabled(EventButtonClicked) = True
        .ReDraw = True
    End With
    
    bSSChangeFlag = False
    
    pnlProg.Visible = False
    Screen.MousePointer = vbDefault
    
    Exit Sub

ErrRtn:
    bSSChangeFlag = False
    
    pnlProg.Visible = False
    Screen.MousePointer = vbDefault
    
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

'--------------------------------------------------------------------------------------------------------------
' Procedure : SendSMS
' DateTime  : 2007-05-06 23:16
' Author    : pds2004
' Purpose   : SMS 문자 메시지 전송
'--------------------------------------------------------------------------------------------------------------
Private Function SendSMS() As Boolean
    Dim sFlag      As Boolean
    Dim lRow       As Long
    Dim dLng       As Long
    Dim sValue(10) As String
    Dim vTemp      As Variant
    Dim bResult    As Boolean
    Dim Query      As String
    Dim ADORset    As New ADODB.RecordSet
    Dim sTel() As String
    Dim sMsg        As String
    Dim nSendCnt    As Long
        
    On Error GoTo ErrRtn
    
    nSendCnt = 0
    cmdBtn(1).Enabled = False
    DoEvents
    
    dLng = CheckSendSaleMessageLangth ' 문자 메시지 길이 확인
    
    If dLng <= 0 Or dLng > SMS_Lng Then
        cmdBtn(1).Enabled = True
        
        Exit Function
    End If
    
    ' 전송 시간 확인
    If "18:00" < Format(Time, "hh:mm") And 가맹점정보.SMS_EMART = "Y" Then
        MsgBox "이마트에서는 18:00 이후에는 문자 메시지 발송을 할 수 없습니다.", vbInformation, "확인"
        
        cmdBtn(1).Enabled = True
        Exit Function
    End If
    
    ' 잔여 수량 확인
    If txtCount(2).Value < 0 Then
        MsgBox "잔여 수량보다 선택수량이 더 큼니다. 선택 수량을 조절하여 주십시요.", vbInformation, "확인"
        
        cmdBtn(1).Enabled = True
        Exit Function
    End If
    
    If GetCheckSMSSendTel(Text1.Text, sTel, True) = False Then
        sMsg = "전기통신사업법 제84조에 의하여 문자 발신 번호는 반드시 입력 하여야 합니다." & vbNewLine
        sMsg = sMsg & vbNewLine
        sMsg = sMsg & " 주요내용" & vbNewLine
        sMsg = sMsg & " - 발신번호 없이 문자 전송 불가" & vbNewLine
        sMsg = sMsg & " - 발신번호는 수신자가 실제 발신(통화)이 가능한 번호만 허용" & vbNewLine
        sMsg = sMsg & " - 일반번호의 경우 지역번호(02,031등)를 앞자리에 포함한 번호만 허용" & vbNewLine
        sMsg = sMsg & " - 대표번호는 8자리만 입력 허용되며,내선번호 포함불가" & vbNewLine
        sMsg = sMsg & " - 030 번호의 경우 12자리까지 허용" & vbNewLine
        
        Call MsgBox(sMsg, vbInformation, "입력확인")
        cmdBtn(1).Enabled = True
        Exit Function
    
    End If

    If m_Connect = False Then
        If CheckConnect = False Then
            MsgBox "본사와의 연결 설정을 확인하여 주십시요", vbInformation, "확인"
            ' 설정 화면을 활성화 한다.
            Call cmdBtn_Click(2)
            
            cmdBtn(1).Enabled = True
            Exit Function
        End If
    End If
    
    ' 최종 확인 메시지
    If MsgBox("메시지를 전송 하시겠습니까?" & Space(10), vbCritical + vbYesNo, "확인") = vbNo Then
    
        cmdBtn(1).Enabled = True
        Exit Function
    End If
    
    
    pnlProg.Visible = True
    pnlProg.Caption = "메시지를 전송 중 입니다. 잠시만 기다려 주십시요." & Space(10)
    Screen.MousePointer = vbHourglass
    DoEvents
    
    For lRow = 1 To sprGrid.MaxRows
        sprGrid.Row = lRow
        
        sprGrid.Col = 1: vTemp = sprGrid.Text & ""
        
        ' 전송 구분일 경우 전송 처리한다.
        If vTemp = "1" Then
            ' 전송, 메시지타입, 수신번호, 발신번호, 메시지, 지사코드, 가맹점코드, 고객코드, 고객성명, 참고5, 참고6
                             sValue(0) = "1"
                             sValue(1) = "0"
            sprGrid.Col = 4: sValue(2) = sprGrid.Text & "" '휴대전화
                             sValue(3) = Trim(Text1.Text)
                             sValue(4) = "(광고)" & vbNewLine & Trim(txtSMS.Text) & vbNewLine & "무료수신거부 080-863-5771"
                             sValue(5) = 가맹점정보.지사코드
                             sValue(6) = 가맹점정보.택코드
            
            Call sprGrid.GetText(8, lRow, vTemp):    sValue(7) = CStr(vTemp)
            Call sprGrid.GetText(2, lRow, vTemp):    sValue(8) = CStr(vTemp)
            
            sValue(9) = 가맹점정보.가맹점코드
            sValue(10) = "3"
            Call send_Purio_SMS(sValue(7), sValue(2), sValue(4))
'            Query = "EXEC PRO_SMS_SEND "
'            Query = Query & "'" & sValue(0) & "', "
'            Query = Query & "'" & sValue(1) & "', "
'            Query = Query & "'" & sValue(2) & "', "
'            Query = Query & "'" & sValue(3) & "', "
'            Query = Query & "'" & sValue(4) & "', "
'            Query = Query & "'" & sValue(5) & "', "
'            Query = Query & "'" & sValue(6) & "', "
'            Query = Query & "'" & sValue(7) & "', "
'            Query = Query & "'" & sValue(8) & "', "
'            Query = Query & "'" & sValue(9) & "', "
'            Query = Query & "'" & sValue(10) & "' "
'
'            If Dir(App.Path & "\NO_SMS.DAT", vbNormal) = "" Then
'                nSendCnt = nSendCnt + 1
'                m_Host_DataBase.Execute Query
'
'                vTemp = "0"
'                Call sprGrid.SetText(1, lRow, vTemp)
'            End If
            Call sprGrid.SetText(1, lRow, "0")
            nSendCnt = nSendCnt + 1
        End If
    Next lRow
    
    Call SetUseSMSCount ' 최종 남은 수량을 설정한다.
    Call btnSelect_Click(1)
    Set ADORset = Nothing
    
    cmdBtn(1).Enabled = True
    
    
    pnlProg.Visible = False
    Screen.MousePointer = vbDefault
    DoEvents
    
    MsgBox Format(nSendCnt, "#,##0") & "건 전송 완료" & Space(10)
    
    Exit Function

ErrRtn:
    cmdBtn(1).Enabled = True
    
    Set ADORset = Nothing
    
    ' 최종 남은 수량을 설정한다.
    Call SetUseSMSCount
    
    DoEvents
    
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Function

Private Function CheckSendSaleMessageLangth() As Integer
    If IsNumeric(lbl_SMS.Tag) = False Then
        CheckSendSaleMessageLangth = 0
        txtSMS.SetFocus
        MsgBox "전송할 메시지를 입력 하여 주십시요..  [" & CStr(Val(lbl_SMS.Tag)) & "자]", vbInformation, "확인"
        Exit Function
        
    ElseIf Val(lbl_SMS.Tag) > SMS_Lng Then
        CheckSendSaleMessageLangth = Val(lbl_SMS.Tag)
        txtSMS.SetFocus
        
        MsgBox sSmsMsg, vbInformation, "확인"
        Exit Function
    Else
        CheckSendSaleMessageLangth = Val(lbl_SMS.Tag)
        Exit Function
    End If
End Function


Private Sub txtSMS_KeyPress(KeyAscii As Integer)
    If 가맹점정보.SMS_EMART = "Y" And m_SMS_EMART_PASS = False Then
        KeyAscii = 0
    End If
    
    ' '입력이 안되도록 수정
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
End Sub
