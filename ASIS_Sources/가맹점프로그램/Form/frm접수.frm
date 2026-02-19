VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm접수 
   Caption         =   "세탁물 접수"
   ClientHeight    =   11190
   ClientLeft      =   4965
   ClientTop       =   2970
   ClientWidth     =   15840
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
   Icon            =   "frm접수.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11190
   ScaleWidth      =   15840
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   11190
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   15840
      _ExtentX        =   27940
      _ExtentY        =   19738
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm접수.frx":000C
      Begin Threed.SSPanel pnlPicture 
         Height          =   3465
         Left            =   12030
         TabIndex        =   55
         Top             =   3030
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   6112
         _Version        =   262144
         CaptionStyle    =   1
         ForeColor       =   16711680
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "오점 사진을 찍을려면 마우스로 클릭하세요."
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image imgCapture 
            Height          =   3405
            Left            =   30
            Stretch         =   -1  'True
            Top             =   30
            Width           =   3285
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   390
         Left            =   12030
         TabIndex        =   54
         Top             =   2625
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   688
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   0
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frm접수.frx":00FE
         Caption         =   " 오점 사진"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm접수.frx":0662
         BevelOuter      =   0
         PictureAlignment=   9
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   8550
         Left            =   15
         TabIndex        =   7
         Top             =   2625
         Width           =   12000
         _Version        =   524288
         _ExtentX        =   21167
         _ExtentY        =   15081
         _StockProps     =   64
         BackColorStyle  =   1
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         MaxCols         =   21
         MaxRows         =   200
         ScrollBars      =   2
         SpreadDesigner  =   "frm접수.frx":0AC4
         UserResize      =   1
         VisibleCols     =   7
         VisibleRows     =   30
         HighlightHeaders=   1
         HighlightStyle  =   1
         ScrollBarStyle  =   2
      End
      Begin Threed.SSPanel pnlCustom 
         Height          =   2160
         Left            =   15
         TabIndex        =   8
         Top             =   450
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   3810
         _Version        =   262144
         BackColor       =   16777215
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin XtremeSuiteControls.PushButton btnInternet 
            Height          =   405
            Left            =   7560
            TabIndex        =   67
            Top             =   892
            Width           =   1035
            _Version        =   851970
            _ExtentX        =   1826
            _ExtentY        =   714
            _StockProps     =   79
            Caption         =   "인터넷"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TextAlignment   =   1
            Appearance      =   6
            Picture         =   "frm접수.frx":25EE
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   420
            Left            =   2340
            TabIndex        =   62
            Top             =   60
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   741
            _Version        =   262144
            BackColor       =   16777215
            Enabled         =   0   'False
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
            Begin VB.ComboBox cboClass 
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   30
               Locked          =   -1  'True
               Style           =   2  '드롭다운 목록
               TabIndex        =   63
               Top             =   45
               Width           =   2025
            End
         End
         Begin XtremeSuiteControls.PushButton btnKeyBoard 
            Height          =   405
            Left            =   7560
            TabIndex        =   44
            Top             =   476
            Width           =   1035
            _Version        =   851970
            _ExtentX        =   1826
            _ExtentY        =   714
            _StockProps     =   79
            Caption         =   "자판 "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TextAlignment   =   1
            Appearance      =   6
            Picture         =   "frm접수.frx":29C0
         End
         Begin VB.TextBox txtTel 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   5445
            TabIndex        =   0
            Top             =   60
            Width           =   2055
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1095
            TabIndex        =   2
            Top             =   465
            Width           =   3330
         End
         Begin VB.TextBox txtAddress 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1095
            TabIndex        =   4
            Top             =   870
            Width           =   6405
         End
         Begin VB.TextBox txtCode 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1095
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   60
            Width           =   1260
         End
         Begin VB.TextBox txtHP 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   5445
            TabIndex        =   1
            Top             =   465
            Width           =   2055
         End
         Begin VB.TextBox txtMemo 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   1095
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   5
            Top             =   1275
            Width           =   6405
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   420
            Index           =   0
            Left            =   60
            TabIndex        =   9
            Top             =   60
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   741
            _Version        =   262144
            Font3D          =   1
            BackColor       =   12648384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "고객코드"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm접수.frx":2F5A
            BorderWidth     =   0
            BevelOuter      =   0
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   420
            Index           =   6
            Left            =   4410
            TabIndex        =   10
            Top             =   60
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   741
            _Version        =   262144
            Font3D          =   1
            BackColor       =   12648384
            PictureMaskColorSource=   1
            PictureUseMask  =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "전화번호"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm접수.frx":329C
            BorderWidth     =   0
            BevelOuter      =   0
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   420
            Index           =   9
            Left            =   60
            TabIndex        =   11
            Top             =   870
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   741
            _Version        =   262144
            Font3D          =   1
            BackColor       =   12648384
            PictureMaskColorSource=   1
            PictureUseMask  =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "주    소"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm접수.frx":35DE
            BorderWidth     =   0
            BevelOuter      =   0
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   420
            Index           =   1
            Left            =   4410
            TabIndex        =   12
            Top             =   465
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   741
            _Version        =   262144
            Font3D          =   1
            BackColor       =   12648384
            PictureMaskColorSource=   1
            PictureUseMask  =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "휴대전화"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm접수.frx":3920
            BorderWidth     =   0
            BevelOuter      =   0
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   420
            Index           =   2
            Left            =   60
            TabIndex        =   13
            Top             =   465
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   741
            _Version        =   262144
            Font3D          =   1
            BackColor       =   12648384
            PictureMaskColorSource=   1
            PictureUseMask  =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "고 객 명"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm접수.frx":3C62
            BorderWidth     =   0
            BevelOuter      =   0
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   825
            Index           =   5
            Left            =   60
            TabIndex        =   14
            Top             =   1275
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   1455
            _Version        =   262144
            Font3D          =   1
            BackColor       =   12648384
            PictureMaskColorSource=   1
            PictureUseMask  =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "메    모"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm접수.frx":3FA4
            BorderWidth     =   0
            BevelOuter      =   0
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdEdit 
            Height          =   405
            Left            =   7560
            TabIndex        =   38
            Top             =   1725
            Width           =   1035
            _Version        =   851970
            _ExtentX        =   1826
            _ExtentY        =   714
            _StockProps     =   79
            Caption         =   " 수정"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm접수.frx":42E6
         End
         Begin XtremeSuiteControls.PushButton cmdNew 
            Height          =   405
            Left            =   7560
            TabIndex        =   39
            Top             =   1308
            Width           =   1035
            _Version        =   851970
            _ExtentX        =   1826
            _ExtentY        =   714
            _StockProps     =   79
            Caption         =   " 신규"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm접수.frx":4CF8
         End
         Begin XtremeSuiteControls.PushButton btnClear 
            Height          =   405
            Left            =   7560
            TabIndex        =   61
            ToolTipText     =   "F8..."
            Top             =   60
            Width           =   1035
            _Version        =   851970
            _ExtentX        =   1826
            _ExtentY        =   714
            _StockProps     =   79
            Caption         =   "지우개"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm접수.frx":570A
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   15
         Top             =   15
         Width           =   15810
         _ExtentX        =   27887
         _ExtentY        =   741
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   255
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
         Caption         =   "      세탁물 접수"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm접수.frx":5CA4
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Label lblGoodsPriceStats 
            BackStyle       =   0  '투명
            Caption         =   "Label4"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2430
            TabIndex        =   66
            Top             =   120
            Width           =   3645
         End
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm접수.frx":5ECA
            Top             =   -15
            Width           =   765
         End
      End
      Begin XtremeSuiteControls.TabControl TabControl1 
         Height          =   2160
         Left            =   8685
         TabIndex        =   16
         Top             =   450
         Width           =   7140
         _Version        =   851970
         _ExtentX        =   12594
         _ExtentY        =   3810
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   10
         Color           =   2
         PaintManager.Layout=   5
         PaintManager.Position=   1
         PaintManager.BoldSelected=   -1  'True
         PaintManager.HotTracking=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         PaintManager.LargeIcons=   -1  'True
         ItemCount       =   3
         Item(0).Caption =   "정보"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "TabControlPage1"
         Item(1).Caption =   "실적"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "TabControlPage2"
         Item(2).Caption =   "사고"
         Item(2).ControlCount=   1
         Item(2).Control(0)=   "TabControlPage3"
         Begin XtremeSuiteControls.TabControlPage TabControlPage3 
            Height          =   2100
            Left            =   -69370
            TabIndex        =   42
            Top             =   30
            Visible         =   0   'False
            Width           =   6030
            _Version        =   851970
            _ExtentX        =   10636
            _ExtentY        =   3704
            _StockProps     =   1
            BackColor       =   255
            Page            =   2
            Begin FPSpreadADO.fpSpread sprClaim 
               Bindings        =   "frm접수.frx":6A94
               Height          =   1935
               Left            =   75
               TabIndex        =   43
               Top             =   75
               Width           =   5880
               _Version        =   524288
               _ExtentX        =   10372
               _ExtentY        =   3413
               _StockProps     =   64
               AllowDragDrop   =   -1  'True
               AllowMultiBlocks=   -1  'True
               AllowUserFormulas=   -1  'True
               BackColorStyle  =   1
               DAutoCellTypes  =   0   'False
               DAutoHeadings   =   0   'False
               DAutoSave       =   0   'False
               DAutoSizeCols   =   0
               DInformActiveRowChange=   0   'False
               DisplayRowHeaders=   0   'False
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
               MaxCols         =   18
               MaxRows         =   1000000
               OperationMode   =   1
               Protect         =   0   'False
               ScrollBarExtMode=   -1  'True
               SpreadDesigner  =   "frm접수.frx":6AA8
               VisibleCols     =   9
               VisibleRows     =   200
               Appearance      =   1
               HighlightHeaders=   1
               HighlightStyle  =   1
               ScrollBarStyle  =   2
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage2 
            Height          =   2100
            Left            =   -69370
            TabIndex        =   17
            Top             =   30
            Visible         =   0   'False
            Width           =   6030
            _Version        =   851970
            _ExtentX        =   10636
            _ExtentY        =   3704
            _StockProps     =   1
            Page            =   1
            Begin FPSpreadADO.fpSpread sprHist 
               Height          =   1680
               Left            =   3570
               TabIndex        =   18
               Top             =   330
               Width           =   2370
               _Version        =   524288
               _ExtentX        =   4180
               _ExtentY        =   2963
               _StockProps     =   64
               BackColorStyle  =   1
               DisplayRowHeaders=   0   'False
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
               MaxCols         =   2
               ScrollBars      =   2
               SpreadDesigner  =   "frm접수.frx":7492
               HighlightHeaders=   1
               HighlightStyle  =   1
            End
            Begin FPSpreadADO.fpSpread sprYear 
               Height          =   1680
               Left            =   75
               TabIndex        =   19
               Top             =   330
               Width           =   3465
               _Version        =   524288
               _ExtentX        =   6112
               _ExtentY        =   2963
               _StockProps     =   64
               BackColorStyle  =   1
               DisplayRowHeaders=   0   'False
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
               MaxCols         =   3
               ScrollBars      =   2
               SpreadDesigner  =   "frm접수.frx":79F9
               HighlightHeaders=   1
               HighlightStyle  =   1
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "년도별 이용현황"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   21
               Top             =   105
               Width           =   1470
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "최근이용현황"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   1
               Left            =   3555
               TabIndex        =   20
               Top             =   105
               Width           =   1170
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage1 
            Height          =   2100
            Left            =   630
            TabIndex        =   22
            Top             =   30
            Width           =   6480
            _Version        =   851970
            _ExtentX        =   11430
            _ExtentY        =   3704
            _StockProps     =   1
            Page            =   0
            Begin XtremeSuiteControls.PushButton btnMisu 
               Height          =   375
               Left            =   5535
               TabIndex        =   40
               Top             =   75
               Width           =   390
               _Version        =   851970
               _ExtentX        =   688
               _ExtentY        =   661
               _StockProps     =   79
               Appearance      =   6
               Picture         =   "frm접수.frx":7FCE
            End
            Begin VB.TextBox txtRegistDay 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   945
               TabIndex        =   23
               Text            =   "2010-12-31"
               Top             =   75
               Width           =   1305
            End
            Begin CSTextLibCtl.sidbEdit txtMisu 
               Height          =   375
               Left            =   4200
               TabIndex        =   24
               Top             =   75
               Width           =   1305
               _Version        =   262145
               _ExtentX        =   2302
               _ExtentY        =   661
               _StockProps     =   125
               Text            =   " 0"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9.74
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DataProperty    =   2
               ReadOnly        =   -1  'True
               Modified        =   -1  'True
               HideSelection   =   -1  'True
               RawData         =   ""
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
               FmtControl      =   1
               NumDecDigits    =   0
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txtUseMileage 
               Height          =   375
               Left            =   4200
               TabIndex        =   25
               Top             =   855
               Width           =   1305
               _Version        =   262145
               _ExtentX        =   2302
               _ExtentY        =   661
               _StockProps     =   125
               Text            =   " 0"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9.74
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DataProperty    =   2
               ReadOnly        =   -1  'True
               Modified        =   -1  'True
               HideSelection   =   -1  'True
               RawData         =   ""
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
               FmtControl      =   1
               NumDecDigits    =   0
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txtTotalMileage 
               Height          =   375
               Left            =   4200
               TabIndex        =   26
               Top             =   1245
               Width           =   1305
               _Version        =   262145
               _ExtentX        =   2302
               _ExtentY        =   661
               _StockProps     =   125
               Text            =   " 0"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9.74
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DataProperty    =   2
               ReadOnly        =   -1  'True
               Modified        =   -1  'True
               HideSelection   =   -1  'True
               RawData         =   ""
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
               FmtControl      =   1
               NumDecDigits    =   0
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txtTotalNum 
               Height          =   375
               Index           =   0
               Left            =   945
               TabIndex        =   27
               Top             =   855
               Width           =   1305
               _Version        =   262145
               _ExtentX        =   2302
               _ExtentY        =   661
               _StockProps     =   125
               Text            =   " 0"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9.74
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DataProperty    =   2
               ReadOnly        =   -1  'True
               Modified        =   -1  'True
               HideSelection   =   -1  'True
               RawData         =   ""
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
               FmtControl      =   1
               NumDecDigits    =   0
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txtTotalNum 
               Height          =   375
               Index           =   1
               Left            =   945
               TabIndex        =   28
               Top             =   1245
               Width           =   1305
               _Version        =   262145
               _ExtentX        =   2302
               _ExtentY        =   661
               _StockProps     =   125
               Text            =   " 0"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9.74
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DataProperty    =   2
               ReadOnly        =   -1  'True
               Modified        =   -1  'True
               HideSelection   =   -1  'True
               RawData         =   ""
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
               FmtControl      =   1
               NumDecDigits    =   0
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txtTotalNum 
               Height          =   375
               Index           =   2
               Left            =   945
               TabIndex        =   29
               Top             =   1635
               Width           =   1305
               _Version        =   262145
               _ExtentX        =   2302
               _ExtentY        =   661
               _StockProps     =   125
               Text            =   " 0"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9.74
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DataProperty    =   2
               ReadOnly        =   -1  'True
               Modified        =   -1  'True
               HideSelection   =   -1  'True
               RawData         =   ""
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
               FmtControl      =   1
               NumDecDigits    =   0
               Undo            =   0
               Data            =   0
            End
            Begin XtremeSuiteControls.PushButton btnMileage 
               Height          =   375
               Left            =   5535
               TabIndex        =   41
               Top             =   1245
               Width           =   390
               _Version        =   851970
               _ExtentX        =   688
               _ExtentY        =   661
               _StockProps     =   79
               Appearance      =   6
               Picture         =   "frm접수.frx":89E0
            End
            Begin XtremeSuiteControls.PushButton btnTagCode 
               Height          =   375
               Left            =   4200
               TabIndex        =   56
               Top             =   1650
               Width           =   630
               _Version        =   851970
               _ExtentX        =   1111
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "999"
               ForeColor       =   192
               BackColor       =   12648447
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   2
            End
            Begin XtremeSuiteControls.PushButton cmdTagNo 
               Height          =   375
               Left            =   4860
               TabIndex        =   57
               Top             =   1650
               Width           =   1125
               _Version        =   851970
               _ExtentX        =   1984
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "99-0000"
               ForeColor       =   192
               BackColor       =   12648447
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   2
            End
            Begin CSTextLibCtl.sidbEdit txtVisit 
               Height          =   375
               Left            =   945
               TabIndex        =   59
               Top             =   465
               Width           =   1305
               _Version        =   262145
               _ExtentX        =   2302
               _ExtentY        =   661
               _StockProps     =   125
               Text            =   " 0"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9.74
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DataProperty    =   2
               ReadOnly        =   -1  'True
               Modified        =   -1  'True
               HideSelection   =   -1  'True
               RawData         =   ""
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
               FmtControl      =   1
               NumDecDigits    =   0
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txtNoRepay 
               Height          =   375
               Left            =   4200
               TabIndex        =   64
               Top             =   465
               Visible         =   0   'False
               Width           =   1305
               _Version        =   262145
               _ExtentX        =   2302
               _ExtentY        =   661
               _StockProps     =   125
               Text            =   " 0"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9.74
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DataProperty    =   2
               ReadOnly        =   -1  'True
               Modified        =   -1  'True
               HideSelection   =   -1  'True
               RawData         =   ""
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
               FmtControl      =   1
               NumDecDigits    =   0
               Undo            =   0
               Data            =   0
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "미환불금액:"
               Height          =   180
               Index           =   8
               Left            =   3165
               TabIndex        =   65
               Top             =   570
               Visible         =   0   'False
               Width           =   990
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "이용횟수:"
               Height          =   180
               Index           =   6
               Left            =   90
               TabIndex        =   60
               Top             =   570
               Width           =   810
            End
            Begin VB.Label Label3 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "택 번 호:"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   0
               Left            =   3345
               TabIndex        =   58
               Top             =   1755
               Width           =   810
            End
            Begin VB.Label lblSamSungCardCheck 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "삼성카드할인여부"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   2325
               TabIndex        =   37
               Top             =   1350
               Visible         =   0   'False
               Width           =   735
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "등록일자:"
               Height          =   180
               Index           =   0
               Left            =   90
               TabIndex        =   36
               Top             =   180
               Width           =   810
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "미수금액:"
               Height          =   180
               Index           =   4
               Left            =   3345
               TabIndex        =   35
               Top             =   180
               Width           =   810
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "사용가능 마일리지:"
               Height          =   240
               Index           =   5
               Left            =   2505
               TabIndex        =   34
               Top             =   960
               Width           =   1650
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "누적 마일리지:"
               Height          =   240
               Index           =   7
               Left            =   2505
               TabIndex        =   33
               Top             =   1365
               Width           =   1650
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "총할인액:"
               Height          =   180
               Index           =   3
               Left            =   90
               TabIndex        =   32
               Top             =   1740
               Width           =   810
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "총입금액:"
               Height          =   180
               Index           =   2
               Left            =   90
               TabIndex        =   31
               Top             =   1350
               Width           =   810
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "총매출액:"
               Height          =   180
               Index           =   1
               Left            =   90
               TabIndex        =   30
               Top             =   960
               Width           =   810
            End
         End
      End
      Begin Threed.SSPanel pnlButton 
         Height          =   4665
         Left            =   12030
         TabIndex        =   45
         Top             =   6510
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   8229
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdOK 
            Height          =   690
            Left            =   1695
            TabIndex        =   46
            Top             =   60
            Width           =   1590
            _Version        =   851970
            _ExtentX        =   2805
            _ExtentY        =   1217
            _StockProps     =   79
            Caption         =   " 확인(F2)"
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
            Picture         =   "frm접수.frx":93F2
         End
         Begin XtremeSuiteControls.PushButton cmdSuite 
            Height          =   690
            Left            =   60
            TabIndex        =   47
            Top             =   795
            Width           =   1590
            _Version        =   851970
            _ExtentX        =   2805
            _ExtentY        =   1217
            _StockProps     =   79
            Caption         =   " 한벌(F3)"
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
            Picture         =   "frm접수.frx":998C
         End
         Begin XtremeSuiteControls.PushButton cmdRepeat 
            Height          =   690
            Left            =   1695
            TabIndex        =   48
            Top             =   795
            Width           =   1590
            _Version        =   851970
            _ExtentX        =   2805
            _ExtentY        =   1217
            _StockProps     =   79
            Caption         =   " 반복(F4)"
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
            Picture         =   "frm접수.frx":9F26
         End
         Begin XtremeSuiteControls.PushButton cmdCorrect 
            Height          =   690
            Left            =   60
            TabIndex        =   49
            Top             =   1530
            Width           =   1590
            _Version        =   851970
            _ExtentX        =   2805
            _ExtentY        =   1217
            _StockProps     =   79
            Caption         =   " 정정(F5)"
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
            Picture         =   "frm접수.frx":A4C0
         End
         Begin XtremeSuiteControls.PushButton cmdCancel 
            Height          =   690
            Left            =   1695
            TabIndex        =   50
            Top             =   1530
            Width           =   1590
            _Version        =   851970
            _ExtentX        =   2805
            _ExtentY        =   1217
            _StockProps     =   79
            Caption         =   " 취소(F6)"
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
            Picture         =   "frm접수.frx":AA5A
         End
         Begin XtremeSuiteControls.PushButton cmdCalculate 
            Height          =   690
            Left            =   60
            TabIndex        =   51
            Top             =   2265
            Width           =   1590
            _Version        =   851970
            _ExtentX        =   2805
            _ExtentY        =   1217
            _StockProps     =   79
            Caption         =   " 접수/계산(F7)"
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
            Picture         =   "frm접수.frx":AFF4
         End
         Begin XtremeSuiteControls.PushButton btnExit 
            Height          =   690
            Left            =   1695
            TabIndex        =   52
            Top             =   2265
            Width           =   1590
            _Version        =   851970
            _ExtentX        =   2805
            _ExtentY        =   1217
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
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
            Picture         =   "frm접수.frx":B58E
         End
         Begin Threed.SSCheck chkRepair 
            Height          =   345
            Left            =   225
            TabIndex        =   53
            Top             =   240
            Visible         =   0   'False
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   609
            _Version        =   262144
            Font3D          =   3
            ForeColor       =   192
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "수선접수"
         End
         Begin VB.Shape Shape 
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00C0C000&
            Height          =   690
            Left            =   60
            Shape           =   4  '둥근 사각형
            Top             =   60
            Width           =   1590
         End
      End
   End
End
Attribute VB_Name = "frm접수"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public frm As Form

Dim Tel_Flag    As Boolean ' 전화 번호에서 다시 전화 번호를 변경할때 스택 오류가 나느것을방지
Dim Search_Flag As Boolean '

'*****************************************************************************************
' 제목    : 이용실적 표시
' 기능    : 고객의 이용실적을 전년과 올해로 나누어 표시
' 전해년도: 가맹점이 이전 자료가 있는 경우-conversion시에 write
' 올해년도: 계산시에 write 함
' 처리    : 이용실적 table 읽어 디스플레이
'*****************************************************************************************
Private Sub 이용실적_Display()
    Dim SSQL    As String
    On Error GoTo ErrRtn

    SSQL = "SELECT    연도"
    SSQL = SSQL & ", 이용금액"
    SSQL = SSQL & ", 이용횟수"
    SSQL = SSQL & " FROM TB_이용실적 "
    SSQL = SSQL & " WHERE 고객코드 = '" & Trim(txtCode.Text) & "' "
    SSQL = SSQL & " ORDER BY 연도 DESC"
    Set SUBRs = New ADODB.RecordSet
    SUBRs.Open SSQL, ADOCon, adOpenForwardOnly, adLockReadOnly

    With sprYear
        .MaxRows = 0
        .ReDraw = False
        
        Do Until SUBRs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = SUBRs!연도 & ""
            .Col = 2: .Text = SUBRs!이용횟수 & ""
            .Col = 3: .Text = SUBRs!이용금액 & ""
            
            SUBRs.MoveNext
        Loop
        SUBRs.Close
        Set SUBRs = Nothing
        
        If .MaxRows > 1 Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = "합계"
            .Col = 2: .Formula = "SUM(B1:B" & .MaxRows - 1 & ")"
            .Col = 3: .Formula = "SUM(C1:C" & .MaxRows - 1 & ")"
        End If
        
        .ReDraw = True
    End With

    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub Get_택코드()
    Dim SSQL    As String
    On Error GoTo ErrRtn
    
    SSQL = "SELECT    택코드"
    SSQL = SSQL & ", 세탁소요일"
    SSQL = SSQL & " FROM TB_기본정보"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open SSQL, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF Then
        btnTagCode.Caption = "999"
    Else
        btnTagCode.Caption = Format(ADORs!택코드, "000")
    End If
    ADORs.Close
    Set ADORs = Nothing
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub btnClear_Click()
    Call Text_Clear
    imgCapture.Picture = Nothing
    txtTel.SetFocus
    Search_Flag = False
    Tel_Flag = False
End Sub

Private Sub btnExit_Click()
    Rtn = MsgBox("닫기 작업을 진행 하시겠습니까?", vbInformation + vbYesNo, "닫기")
    
    If Rtn = vbNo Then Exit Sub
    
    Unload Me
End Sub

Private Sub btnInternet_Click()
    frmInternetAccept.GetData
    frmInternetAccept.Show vbModal
    DoEvents
    If btnInternet.tag <> "" Then
        
        txtCode.Text = frmInternetAccept.SELECTCODE
        Call 고객정보_Display(txtCode.Text)
    End If
End Sub

Private Sub btnKeyBoard_Click()
    
    Load frmKeyboard
    frmKeyboard.Left = frmMain.Left + 7000
    frmKeyboard.Top = frmMain.Top + 3000
    
     frmKeyboard.Show 1
End Sub

Private Sub btnMileage_Click()
     On Error GoTo ErrRtn
    
    frm마일리지.lblCode.Caption = txtCode.Text
    frm마일리지.Show 1
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub chkRepair_Click(Value As Integer)
    chkRepair.Enabled = False
End Sub

Private Sub btnMisu_Click()
    frm미수금.lblCode.Caption = txtCode.Text & ""
    frm미수금.lblMisu.Caption = txtMisu.Value
    
    frm미수금.Show 1
End Sub

Private Sub cmdOK_Click()
    Dim strTagNo As String
    Dim strTemp  As String
  
    On Error GoTo ErrRtn
    
    If txtTel.Text = "" Then
        Beep
        Exit Sub
    End If
    
    Call Chk_GroupGoodsReturn '세트 상품의 경우 다시 원상복구를 한다.
    
    For Each frm In Forms
        Select Case frm.Name
            Case "frm의류":     frm의류.Hide   ' Unload frm의류
            Case "frm색상표":   frm색상표.Hide ' Unload frm색상표
            Case "frm무늬":     Unload frm무늬 '
            Case "frm작업":     Unload frm작업 '
        End Select
    Next frm
    
    iCur = GetSpreadLine(sprGrid) - 1 ' 마지막 라인을 구한다.
    
    ' 최대 입력 갯수를 지정한다. 기본( .MaxRows )로 설정
    If iCur = sprGrid.MaxRows Then
        MsgBox "한전표에 [ " & sprGrid.MaxRows & "건 ] 이상 입력 할 수 없습니다. ", vbInformation, "입력 확인"
        Exit Sub
    End If
    
    ' 스프레드에 포커스 - 마지막 라인보다 더 밑에 커서가 있을 경우 종류 화면을 보여준다
    If iCur < sprGrid.ActiveRow Then
        sprGrid.Row = iCur
        sprGrid.Col = 1
        sprGrid.Action = ActionActiveCell
    End If
    
    If sprGrid.ActiveCol = 3 Or sprGrid.ActiveCol = 4 Or sprGrid.ActiveCol = 5 Then
        sprGrid.Row = iCur
        sprGrid.Col = 7
        
        If sprGrid.Text = "짜집기(cm당)" Then
            sprGrid.EditMode = False
        Else
            sprGrid.EditMode = True
        End If
        
        sprGrid.SetActiveCell 6, iCur
        
        Exit Sub
    End If
    
    sprGrid.Row = iCur
    sprGrid.Col = 5: strTemp = Trim(sprGrid.Text) '구택이 필요한경우
    
    If strTemp = "드반" Or strTemp = "드재" Or strTemp = "드사" Or strTemp = "드재다" Then
        sprGrid.Col = 7
        
        If Len(Trim(sprGrid.Text)) < 5 Then
            Load frm작업구분
            
            Select Case strTemp
                Case "드반": frm작업구분.SetFlags "반품 구분"
                Case "드재": frm작업구분.SetFlags "재세탁 구분"
                Case "드사": frm작업구분.SetFlags "사고품 구분"
                Case Else:   frm작업구분.SetFlags "재다림질 구분"
            End Select
                        
            frm작업구분.Show
            
            sprGrid.SetActiveCell 7, iCur ' Active Cell
            
            Exit Sub
        Else
            sprGrid.SetActiveCell 7, iCur ' 있으면 계속진행하고 없으면 다시입력요망
        End If
    End If
   
    For i = 1 To sprGrid.MaxRows
        sprGrid.Row = i
        sprGrid.Col = 1
        
        If Trim(sprGrid.Text) = "" Then
            iCur = i
            
            Exit For
        End If
    Next i
    
    Load frm의류
    frm의류.Show              ' 수정이유 - 자동으로 Load
    
    With sprGrid
        .Row = iCur
        .BackColor = vbGreen
        .SetActiveCell 1, iCur ' Active Cell
    
    End With
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub cmdEdit_Click()
    If Trim(txtCode.Text) <> "" Then
        frm고객수정.txtCode.Text = txtCode.Text
        
        frm고객수정.Show 1
    End If
End Sub

Private Sub cmdRepeat_Click()
    ' 반복
    Dim strTemp(0 To 18) As String ' copy용
    Dim strNum1 As String          ' 택번호1
    Dim strNum2 As String          ' 택번호2
    Dim intNum1 As Integer         ' tagno1
    Dim intNum2 As Integer         ' tagno2
    Dim intCol01 As Integer
 
    On Error GoTo ErrRtn
    
    ' 재세탁및 기타 필수 입력 내용 확인한다.
    If Check_SpreadOrder = False Then Exit Sub
    
    cmdRepeat.Enabled = False
    DoEvents
    
    Call Chk_GroupGoodsReturn '세트 상품의 경우 다시 원상복구를 한다.
    
    For Each frm In Forms
        Select Case frm.Name
            Case "frm의류":     frm의류.Hide   ' Unload frm의류
            Case "frm색상표":   frm색상표.Hide ' Unload frm색상표
            Case "frm무늬":     Unload frm무늬 '
            Case "frm작업":     Unload frm작업 '
        End Select
    Next frm
        
    If (sprGrid.Row = 1 And sprGrid.Col = 1) And (sprGrid.Value) = "" Then
        Beep
        cmdRepeat.Enabled = True
        Exit Sub
    End If
    
    ' 마지막 라인을 구한다.
    iCur = GetSpreadLine(sprGrid) - 1
    
    ' 최대 입력 갯수를 지정한다. 기본( .MaxRows )로 설정
    If iCur = sprGrid.MaxRows Then
        MsgBox "한전표에 [ " & sprGrid.MaxRows & "건 ] 이상 입력 할 수 없습니다. ", vbInformation, "입력 확인"
        Exit Sub
    End If
    
    With sprGrid
        .Row = iCur
        .Col = 1: strTemp(0) = .Text & ""        ' 1 품명
        If 가맹점정보.DualComputer = "Y" Then
            .Col = 2: strTemp(1) = ""
        Else
            .Col = 2: strTemp(1) = cmdTagNo.Caption  ' 2 택번호
        End If
        .Col = 3: strTemp(2) = .Text & ""        ' 3 색상
        .Col = 4: strTemp(3) = .Text & ""        ' 4 무늬
        .Col = 5: strTemp(4) = .Text & ""        ' 5 내용
        .Col = 6: strTemp(5) = .Value            ' 6 금액
        .Col = 7: strTemp(6) = .Text & ""        ' 7 상표
        .Col = 8: strTemp(7) = .Value            ' 8 의류코드
        .Col = 9: strTemp(8) = .Value            ' 9 수선금액
        .Col = 10: strTemp(9) = .Value           '10
        .Col = 11: strTemp(10) = .Value          '11
        .Col = 12: strTemp(11) = .Value          '12
        .Col = 13: strTemp(12) = .Value          '13
        .Col = 14: strTemp(13) = .Value          '14
        
        .Col = 16: strTemp(14) = .Value          '16 세탁마진
        .Col = 17: strTemp(15) = .Value          '17
        .Col = 18: strTemp(16) = .Value          '18
        .Col = 19: strTemp(17) = .Value          '19
        .Col = 20: strTemp(18) = .Value          '20
        
        iCur = iCur + 1                          ' 다시복사
        
        .Row = iCur
        .Col = 1: .Text = strTemp(0)             ' 1 품명
        .Col = 2: .Text = strTemp(1)             ' 2 택번호
        .Col = 3: .Value = strTemp(2)            ' 3 색상
        .Col = 4: .Text = strTemp(3)             ' 4 무늬
        .Col = 5: .Text = strTemp(4)             ' 5 내용
        .Col = 6: .Value = strTemp(5)            ' 6 금액
        .Col = 7: .Text = strTemp(6)             ' 7 상표
        .Col = 8: .Value = strTemp(7)            ' 8 의류코드
        .Col = 9: .Value = strTemp(8)            ' 9 수선금액
        .Col = 10: .Value = strTemp(9)           '10
        .Col = 11: .Value = strTemp(10)          '11
        .Col = 12: .Value = strTemp(11)          '12
        .Col = 13: .Value = strTemp(12)          '13
        .Col = 14: .Value = strTemp(13)          '14
    
        .Col = 16: .Value = strTemp(14)          '16 세탁마진
        .Col = 17: .Value = strTemp(15)          '17
        .Col = 18: .Value = strTemp(16)          '18
        .Col = 19: .Value = strTemp(17)          '19
        .Col = 20: .Value = strTemp(18)          '20
    End With
    
    sprGrid.SetActiveCell 1, iCur ' Active Cell
    
    If 가맹점정보.DualComputer = "Y" Then
        '
    Else
        cmdTagNo.Caption = Get_TagNo(strTemp(1), "+") '택 번호를 증가 한다.
    End If
    
    Erase strTemp
    
    cmdRepeat.Enabled = True
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

'----------------------------------------------------------------
' 선택된 라인을 삭제하고 다음 라인을 당긴다
' 삭제된 라인 다음부터 택번호를 -1씩 감소한다
' 메인의 택번호를 수정한다. ( -1) 삼소
'----------------------------------------------------------------
Private Sub cmdCorrect_Click()
    Dim iRow As Integer
    Dim 택번호 As String
    Dim sFilename   As String
    
    On Error GoTo ErrRtn
    
    Call Chk_GroupGoodsReturn     '세트 상품의 경우 다시 원상복구를 한다.
    
    iCur = sprGrid.ActiveRow '현재 선택된 로우값 저장
    
    For Each frm In Forms
        Select Case frm.Name
            Case "frm의류":     frm의류.Hide   ' Unload frm의류
            Case "frm색상표":   frm색상표.Hide ' Unload frm색상표
            Case "frm무늬":     Unload frm무늬 '
            Case "frm작업":     Unload frm작업 '
        End Select
    Next frm
   
    ' 현재 선택된 곳에 아무 내용이 없을 경우
    sprGrid.Row = iCur
    sprGrid.Col = 1
    If Trim(sprGrid.Text) = "" Then
        Beep
        Exit Sub
    End If
    
    
    sprGrid.Col = 2
    
    sFilename = btnTagCode.Caption & Replace(sprGrid.Text, "-", "")
    
    If Dir(App.Path & "\Capture\" & Format(Date, "YYYYMMDD") & sFilename & ".jpg") <> "" Then
        Kill App.Path & "\Capture\" & Format(Date, "YYYYMMDD") & sFilename & ".jpg"
        imgCapture.Picture = Nothing
    End If
    
    
    
    sprGrid.DeleteRows iCur, 1 ' 선택된 열을 삭제한다.
        
    If 가맹점정보.DualComputer = "Y" Then
    
    Else
        ' TAG번호를 정렬한다.
        For iRow = iCur To sprGrid.MaxRows
            ' 전표 번호가 있을 경우 그 전표번호의 이전 전표 번호를 기록한다.
            sprGrid.Row = iRow
            sprGrid.Col = 2
            If Trim(sprGrid.Text) <> "" Then
                sprGrid.Col = 2: 택번호 = sprGrid.Text & ""
                
                택번호 = Get_TagNo(택번호, "-")
                
                sprGrid.Row = iRow
                sprGrid.Col = 2: sprGrid.Text = 택번호
            Else
                ' 메인의 택 번호를 -1 감소한다.
                Dim sNewTag As String
                
                sNewTag = Get_ChangeTagNo(cmdTagNo.Caption, "-")
                cmdTagNo.Caption = Get_TagNo(sNewTag, "-")
                
                Exit For
            End If
        Next iRow
    End If
    sprGrid.SetActiveCell 1, GetSpreadLine(sprGrid) ' Active Cell
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub Text_Clear()
    btnInternet.tag = ""
    
    pnlCustom.BackColor = vbWhite
    
    txtTel.Text = ""
    txtName.Text = ""
    txtHP.Text = ""
    txtHP.tag = ""
    txtCode.Text = ""
    txtAddress.Text = ""
    txtMemo.Text = ""
    
    txtCode.Locked = False
    txtTel.Locked = False
    txtHP.Locked = False
    txtName.Locked = False
    txtAddress.Locked = False
    txtMemo.Locked = False
    
    
    txtVisit.Value = 0
    txtTotalNum(0).Value = 0
    txtTotalNum(1).Value = 0
    txtTotalNum(2).Value = 0
    
    txtRegistDay.Text = Date
    
    txtMisu.Value = 0
    txtNoRepay.Value = 0
    
    txtUseMileage.Value = 0   '사용가능 마일리지
    txtTotalMileage.Value = 0 '누적 마일리지
    
    '
    sprYear.MaxRows = 0
    sprHist.MaxRows = 0
    sprClaim.MaxRows = 0
    
    sprGrid.MaxRows = 0
    sprGrid.MaxRows = 200 '
    
    Search_Flag = False
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo ErrRtn
    
    For Each frm In Forms
        Select Case frm.Name
            Case "frm작업":     Unload frm작업     '
            Case "frm접수결제": Unload frm접수결제 '
            Case "frm의류":     frm의류.Hide       ' Unload frm의류
            Case "frm색상표":   frm색상표.Hide     ' Unload frm색상표
        End Select
    Next frm
    
    Rtn = MsgBox("취소 하시겠습니까?", vbInformation + vbYesNo, "입고취소")
    
    If Rtn = vbNo Then
        Call Chk_GroupGoodsReturn '세트 상품의 경우 다시 원상복구를 한다.

        Exit Sub
    End If
   
    chkRepair.Value = 0      '
    chkRepair.Enabled = True '2010-05-05
    
    txtTel.SetFocus
   
   
   
    Call 접수_Clear
    
    imgCapture.Picture = Nothing
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Public Sub 접수_Clear()
    Dim sTag  As String
    
    Call Text_Clear
   
    sprGrid.Row = 1
    sprGrid.Col = 1
    If Trim(sprGrid.Text) = "" Then
        Beep
        
        Exit Sub
    End If
            
    sprGrid.MaxRows = 0   '
    sprGrid.MaxRows = 200 '
    
    sprGrid.SetActiveCell 1, 1 ' Active Cell
    
    iCur = 1

    sTag = Get_TagNo("", "R")
    cmdTagNo.Caption = Get_TagNo(sTag, "+")
    
    
    
End Sub

Private Function CheckFile()
    If btnInternet.tag <> "" Then
        With sprGrid
            Dim LoopI As Integer
            For LoopI = 1 To .DataRowCnt
                .Row = LoopI
                .Col = 2
                If .Text <> "" Then
                    Dim imgFileName As String
                    imgFileName = btnTagCode.Caption & Replace(sprGrid.Text, "-", "")
                    
                    If Dir(App.Path & "\Capture\" & Format(Date, "YYYYMMDD") & imgFileName & ".jpg") = "" Then
                        CheckFile = False
                        MsgBox "택번호 " & sprGrid.Text & "의 사진이 첨부되지 않았습니다." & vbCrLf & vbCrLf & "인터넷 접수는 사진을 첨부하여야 합니다.", vbCritical
                        Exit Function
                    End If
                End If

            Next LoopI
        End With

        Shell AppPath & "uploader.exe", vbHide

    End If

    CheckFile = True
End Function

'계산
Private Sub cmdCalculate_Click()
    Dim nRow    As Long
    Dim dPrice            As Double
    Dim dOrgPrice         As Double
    Dim dDiscountTotal    As Double
    Dim dOrgPriceTotal    As Double
    Dim strNum1  As String   ' 택번호1
    Dim strNum2  As String   ' 택번호2
    Dim intNum1  As Integer  ' tagno1
    Dim intNum2  As Integer  ' tagno2
        
    On Error GoTo ErrRtn
    
    If Not CheckFile Then Exit Sub
    '고객이 조회되어 있지 않을경우
    If txtCode.Text = "" Then Exit Sub
    
    ' 재세탁및 기타 필수 입력 내용 확인한다.
    If Check_SpreadOrder = False Then Exit Sub
    
    For Each frm In Forms
        Select Case frm.Name
            Case "frm의류":     frm의류.Hide   ' Unload frm의류
            Case "frm색상표":   frm색상표.Hide ' Unload frm색상표
            Case "frm무늬":     Unload frm무늬 '
            Case "frm작업":     Unload frm작업 '
        End Select
    Next frm
    
    '의류명이 없는 경우 접수가 없으므로 Exit...
    sprGrid.Row = 1
    sprGrid.Col = 1
    If Trim(sprGrid.Text) = "" Then
        Beep
        
        Exit Sub
    End If
    
    '--------------------------------------------------------------
    '세트 상품의 경우 다시 원상복구를 한다.
    '다른 곳을 클릭하였을 경우 원금액을 다시 환원후 다시 계산한다.
    '--------------------------------------------------------------
    Call Chk_GroupGoodsReturn


' 할인 금액을 구한다.
    DoEvents

    dDiscountTotal = 0
    dOrgPriceTotal = 0
    
    With sprGrid
        For nRow = 1 To .MaxRows
            dPrice = 0
            dOrgPrice = 0
            
            ' 현재 수령 금액을 얻어 온다.
            .Row = nRow:    .Col = 6
            If Trim(.Value) <> "" Then dPrice = CDbl(Replace(.Value, ",", ""))
    
            ' 원 금액 금액
            .Row = nRow:    .Col = 20
            If Trim(sprGrid.Value) = "" Then Exit For
            If Trim(.Value) <> "" Then
                dOrgPrice = CDbl(Replace(.Value, ",", ""))
            End If
            
            ' 할증이 있을 경우 할증 금액을 정상금액으로 처리한다.
            If dPrice >= dOrgPrice Then
                dOrgPriceTotal = dOrgPriceTotal + dPrice
                
            ' 할증이 없고 할인이 있을 경우 정상금액을 기준으로 한다.
            Else
                dOrgPriceTotal = dOrgPriceTotal + dOrgPrice
            
            End If
            
            ' 할인된 금액만을 구한다.
            ' 할증처리된 금액이 차감되는 문제 처리
            If dPrice < dOrgPrice Then dDiscountTotal = dDiscountTotal + (dOrgPrice - dPrice)
        Next nRow
    End With


    Call Chk_세트상품확인(sprGrid) ' pds2004 2009-11-14일 추가
    
    chkinputflig = "입고중"
    
    Load frm접수결제
    
    frm접수결제.txtTotalPay.Value = 세트상품정보.d최종수령액  ' Call Spread_SetData(frm접수결제.sprMoney, 3, 1, 세트상품정보.d최종수령액)     '합계 금액
    frm접수결제.txtTotalPay2.Value = Format(dOrgPriceTotal, "#,##0")  ' Call Spread_SetData(frm접수결제.sprMoney, 7, 1, 세트상품정보.d전체금액)       '할인전   금액
    frm접수결제.txtSetDC.Value = Format(dDiscountTotal, "#,##0")  ' Call Spread_SetData(frm접수결제.sprMoney, 8, 1, 세트상품정보.d세트할인금액)   '세트할인 금액
    frm접수결제.txtDC.Value = 0    ' Call Spread_SetData(frm접수결제.sprMoney, 9, 1, 세트상품정보.d에누리할인금액) '에누리   금액
    frm접수결제.txtDCTotal.Value = Format(dDiscountTotal, "#,##0") ' Call Spread_SetData(frm접수결제.sprMoney, 10, 1, 세트상품정보.d전체할인금액)  '할인합계 금액
    
    If btnInternet.tag <> "" Then
        frm접수결제.cmdAction(0).Visible = False
        frm접수결제.cmdAction(1).Visible = False
        frm접수결제.cmdAction(2).Visible = False
    Else
        frm접수결제.cmdAction(3).Visible = False
    End If
    
    frm접수결제.Show 1
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    
    ActiveForm = "접수"
    
        ' 무실동점 전산
    If 가맹점정보.가맹점코드 = "100576" Then '
        Unload Me
        Exit Sub
    End If

    
    Call Resize_Rtn
    
    '-------------------------------------------------------------------
    ' 서버를 이용하여 백업여부를 - DBUpdate.exe 가 죽으면 재실행
    '-------------------------------------------------------------------
    Dim strBACKUP As String
    Dim strFile   As String
    
    strBACKUP = GetIniStr("UPDATE", "BACKUP", "", iniFile) '데이터베이스
    
    If strBACKUP = "Y" Then
        strFile = Dir(AppPath & "DBUpdate.exe")
        
        If strFile <> "" Then
            'Call Excute_Program(Me, "DBUpdate.exe")
            Shell AppPath & "DBUpdate.exe", vbHide
        End If
    End If
    
    txtTel.SetFocus
    
    
    If Len(Trim(btnTagCode.Caption)) <> 3 Then
        MsgBox "대리점 코드가 올바르지 않습니다. 확인 하여 주십시요.", vbInformation, "확인"
        Exit Sub
    End If
    
    If Len(Trim(cmdTagNo.Caption)) <> 7 Then
        MsgBox "택번호가 올바르지 않습니다. 확인 하여 주십시요.", vbInformation, "확인"
        Exit Sub
    End If
    
    
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    With sprGrid
        .RowHeight(-1) = 20
        
        .Row = -1
        
        .Col = 8:  .ColHidden = True '
        .Col = 9:  .ColHidden = True '
        .Col = 10: .ColHidden = True '
        .Col = 11: .ColHidden = True '
        .Col = 12: .ColHidden = True '
        .Col = 13: .ColHidden = True '
        .Col = 14: .ColHidden = True '

        .Col = 16: .ColHidden = True '
        .Col = 17: .ColHidden = True '
        .Col = 18: .ColHidden = True '

        .Col = 20: .ColHidden = True '
        .Col = 21: .ColHidden = True '
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeNormal
        
        .EditModePermanent = False
        .EditModeReplace = True
    End With
        
    With sprYear
        .MaxRows = 0
        .RowHeight(-1) = 14
            
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeNormal
    End With

    With sprHist
        .MaxRows = 0
        .RowHeight(-1) = 14
            
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeNormal
    End With
    
    With sprClaim
        .MaxRows = 0
        .RowHeight(-1) = 14
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeRow
    End With
    
    TabControl1.SelectedItem = 0
    
    Call Resize_Rtn
    
    Call 고객등급_Display(cboClass, False) '고객등급
    cboClass.ListIndex = 2

    chkinputflig = "입고중" '현재 상태..
    
    cmdSuite.Enabled = False
    
    Call Get_택코드 ' 택코드
    
    txtRegistDay.Text = Date
    
    iCur = 1 ' 초기화
    
    sprGrid.SetActiveCell 1, 1 ' Active Cell

    lblGoodsPriceStats.Caption = ""
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ' 입고, 출고, 조회, 종료 체크
    'KeyChk (KeyCode)
    
    Select Case KeyCode
        Case 113: cmdOK_Click        'F2 -
        Case 114: cmdSuite_Click     'F3-
        Case 115: cmdRepeat_Click    'F4 -
        Case 116: cmdCorrect_Click   'F5 -
        Case 117: cmdCancel_Click    'F6 -
        Case 118: cmdCalculate_Click 'F7 -
                
        Case 119: btnClear_Click     'F8 -
        
        Case Else
            If Shift = 4 And KeyCode = 88 Then 'Alt+X
                Unload Me
            End If
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{TAB}"
'        KeyAscii = 0
'    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    'btnExit.Top = pnlButton.Height - btnExit.Height - 70
    
    Call Resize_Rtn
End Sub

Private Sub Form_Unload(Cancel As Integer)
 
    sprGrid.Row = 1
    sprGrid.Col = 1
    ' 입고 완료일때는 폼을 지운다.
    If chkinputflig = "입고완료" Then
        Cancel = 0
        Exit Sub
    End If

End Sub

Private Sub cmdTagNo_Click()
    frmTag.Show 1
End Sub

Private Sub imgCapture_Click()
    Dim sFilename As String
    
    If (Len(txtCode.Text) < 6) Or (txtCode.Text = "") Then Exit Sub
    
    If sprGrid.ActiveRow <= 0 Then Exit Sub
    
    sprGrid.Row = sprGrid.ActiveRow
    sprGrid.Col = 2
    
    ' 선택된 내용이 없을 경우
    If Len(sprGrid.Text) <> 7 Then Exit Sub
    sFilename = btnTagCode.Caption & Replace(sprGrid.Text, "-", "")
    
    frm접수오점표시.lblTag.Caption = sFilename & ""
    frm접수오점표시.lblDate.Caption = Format(Date, "YYYY-MM-DD")
    
    If Dir(App.Path & "\Capture\" & Format(Date, "YYYYMMDD") & sFilename & ".jpg") = "" Then
        frm접수오점표시.picCapture.Picture = LoadPicture()
    Else
        frm접수오점표시.picCapture.Picture = LoadPicture(App.Path & "\Capture\" & Format(Date, "YYYYMMDD") & sFilename & ".jpg")
    End If
    
    frm접수오점표시.Show 1
End Sub

Private Sub Label2_Click(Index As Integer)
    Dim strFile As String
    
    strFile = Dir(AppPath & "DBUpdate.exe")
    
    Select Case Index
        Case 4
            If strFile <> "" Then
                Excute_Program Me, "DBUpdate.exe"
            Else
                MsgBox "DBUpdate.exe 파일이 없습니다."
                Exit Sub
            End If
        
        Case 5
            If strFile <> "" Then
                Shell AppPath & "DBUpdate.exe", vbHide
            Else
                MsgBox "DBUpdate.exe 파일이 없습니다."
                Exit Sub
            End If
        
    End Select
End Sub

Private Sub pnlButton_Resize()
    On Error GoTo ErrRtn
    
    'btnExit.Top = pnlButton.Height - btnExit.Height - 70
    
    Exit Sub
    
ErrRtn:

End Sub

Private Sub txtCode_Change()
    If (txtCode.Text = "") Or (Len(txtCode.Text) <> 6) Then
        Call Text_Clear
    End If
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 40: txtName.SetFocus 'Down Key
        Case 38: txtMemo.SetFocus 'Up Key
    End Select
    
    If KeyCode = vbKeyReturn Then
        If Search_Flag = True Then Exit Sub
        
        Call Get_고객조회("Code", Trim(txtCode.Text)) ' 고객정보를 검색한다.
    End If
End Sub

Private Sub txthp_Change()
    If Trim(txtHP.Text) = "" Then
        Call Text_Clear
    End If
    
    If (Len(txtHP.Text) >= 4) And (Tel_Flag = False) Then
        If txtCode.Text <> "" Then Exit Sub
        
        Tel_Flag = True
        
        Call Get_고객조회("Tel", Trim(txtHP.Text)) ' 고객정보를 검색한다.
        
        Tel_Flag = False
    End If
End Sub

Private Sub txtHP_GotFocus()
    txtHP.SelStart = 0
    txtHP.SelLength = Len(txtHP.Text)
    txtHP.BackColor = vbYellow
End Sub

Private Sub txtHP_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 40: txtAddress.SetFocus  'Down Key
        Case 38: txtTel.SetFocus      'Up Key
    End Select
    
    If KeyCode = vbKeyReturn And txtCode.Text = "" Then
        If Search_Flag = True Then Exit Sub
        
        Call Get_고객조회("Tel", Trim(txtHP.Text)) ' 고객정보를 검색한다.
    End If
End Sub

Private Sub txtHP_LostFocus()
    Dim sSendTel(2)     As String
    
    txtHP.Text = Replace(txtHP.Text, "'", "")
    txtHP.Text = Trim(txtHP.Text)

'    If txtHP.Tag <> txtHP.Text Then
'        If MsgBox("휴대폰 번호를 수정하시겠습니까?", vbYesNo + vbCritical + vbDefaultButton2, "확인") = vbYes Then
'            Query = "UPDATE TB_고객정보 SET 휴대전화 = '" & txtHP.Text & "' "
'            Query = Query & " WHERE 고객코드 = '" & txtCode.Text & "' "
'            ADOCon.Execute Query
'        Else
'            txtHP.Text = txtHP.Tag
'        End If
'    End If
    
    txtHP.tag = txtHP.Text
    txtHP.BackColor = "&H00FFFFFF"
End Sub

''Private Sub txtCardNo_LostFocus()
''    Dim strRuMoney As String
''
''    ' pds2004 수정
''    ' 아무것도 입력 하지 않을경우 출력메시지 무시
''    If txtCardNo.RawData = "" Then Exit Sub
''
''    '-------------------------------------------------------
''    ' 가맹점 코드를 Check한다.
''    '-------------------------------------------------------
''    Query = "SELECT 택코드 FROM TB_기본정보"
''    Set ADORs = New ADODB.Recordset
''    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
''
''    If Left(txtCardNo.Text, 3) <> ADORs!택코드 Then
''        ADORs.Close
''        Set ADORs = Nothing
''
''        MsgBox "본 고객카드는 당가맹점의 고객이 아니므로 사용할 수 없습니다.", vbInformation
''
''        Exit Sub
''    End If
''    ADORs.Close
''    Set ADORs = Nothing
''
''    '------------------------------------------------------------------------------
''    Query = " SELECT * FROM TB_고객정보 "
''    Query = Query & " WHERE 카드번호 = '" & Right(txtCardNo.Text, 6) & "' "
''    Set ADORs = New ADODB.Recordset
''    ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
''
''    If ADORs.EOF Then
''        ADORs.Close
''        Set ADORs = Nothing
''
''        MsgBox "해당되는 카드번호는 존재하지 않습니다..."
''
''        Exit Sub
''    Else
''        '뿌리고 입력대기상태
''        txtTel.Text = Trim(ADORs!전화번호) & ""  '
''        txtCode.Text = Trim(ADORs!고객코드) & "" '
''        txtAddress.Text = Trim(ADORs!주소) & ""  '
''        txtName.Text = Trim(ADORs!성명) & ""     '
''        txtMisu.Value = ADORs!미수금액 & ""        '
''        txtMemo.Text = ADORs!메모 & ""           '
''
''        ADORs.Close
''        Set ADORs = Nothing
''
''        Call 이용실적_Display
''
'''        '---------------------------------------------------------------------------
'''        Query = "SELECT COUNT(택번호) AS Total2 ,SUM(Clng((금액))) AS [Total] "
'''        Query = Query & " FROM TB_입출고 "
'''        Query = Query & " WHERE 고객코드 = '" & Trim(txtCode.Text) & "' "
'''        Set ADORs = New ADODB.Recordset
'''        ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
'''
'''        txtTotal.Text = Format(Trim(ADORs!TOTAL), " #,###,###")
'''        txtCnt.Text = Format(Trim(ADORs!TOTAL2), "###,###")
'''
'''        ADORs.Close
'''        Set ADORs = Nothing
''    End If
''End Sub

Private Sub txtAddress_GotFocus()
    Dim hiMe As Long
    
    txtAddress.BackColor = vbYellow '"&H0080FF80"
    Toggle_Check = True
    
    ' //KEYCODE 123 번은 펑션키12번(F12)
    ' //특정키를 입력하려면 아래 KEYCODE만 바꿔주면됨
    If Toggle_Check = True Then
        ' // 한글로 바꾸기
        hiMe = ImmGetContext(txtAddress.hWnd)
        ImmSetConversionStatus hiMe, IME_HANGUL, IME_NONE
        Toggle_Check = False
    Else
        ' // 영어로 바꾸기
        hiMe = ImmGetContext(txtAddress.hWnd)
        ImmSetConversionStatus hiMe, IME_ENGLISH, IME_NONE
        Toggle_Check = True
    End If
End Sub

Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrRtn
    
    Select Case KeyCode
        Case 40: txtMemo.SetFocus  'Down Key
        Case 38: txtHP.SetFocus 'Up Key
    End Select
       
    If KeyCode = vbKeyReturn Then
        If Search_Flag = True Then Exit Sub
        
        Call Get_고객조회("Addr", Trim(txtAddress.Text))
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub txtAddress_LostFocus()
    txtAddress.BackColor = "&H00FFFFFF"
End Sub

Private Sub txtMemo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 40: txtCode.SetFocus  'Down Key
        Case 38: txtAddress.SetFocus 'Up Key
    End Select
End Sub

Private Sub txtName_Change()
    If Trim(txtName.Text) = "" Then
        Call Text_Clear
    End If
End Sub

Private Sub txtName_GotFocus()
    Dim hiMe As Long
          
    txtName.BackColor = vbYellow ' "&H0080FF80"
          
    Toggle_Check = True
    
    ' //KEYCODE 123 번은 펑션키12번(F12)
    ' //특정키를 입력하려면 아래 KEYCODE만 바꿔주면됨
    If Toggle_Check = True Then
        ' // 한글로 바꾸기
        hiMe = ImmGetContext(txtName.hWnd)
        ImmSetConversionStatus hiMe, IME_HANGUL, IME_NONE
        Toggle_Check = False
    Else
        ' // 영어로 바꾸기
        hiMe = ImmGetContext(txtName.hWnd)
        ImmSetConversionStatus hiMe, IME_ENGLISH, IME_NONE
        Toggle_Check = True
    End If
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrRtn
    
    Select Case KeyCode
        Case 40: txtTel.SetFocus  'Down Key
        Case 38: txtCode.SetFocus 'Up Key
    End Select
    
    If (KeyCode = vbKeyReturn) And (Trim(txtName.Text) <> "") Then
        If Search_Flag = True Then Exit Sub
        
        Call Get_고객조회("Name", Trim(txtName.Text))
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub txtName_LostFocus()
    On Error GoTo ErrRtn

    txtName.BackColor = "&H00FFFFFF"
    Toggle_Check = False
    txtName.Text = Replace(txtName.Text, "'", "")
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub txtTel_Change()
    If Trim(txtTel.Text) = "" Then
        Call Text_Clear
    End If
    
    If Search_Flag = True And Tel_Flag = False Then Exit Sub
    
    If (Len(txtTel.Text) >= 4) And (Tel_Flag = False) Then
        If txtCode.Text <> "" Then Exit Sub
        
        DoEvents
        
        Tel_Flag = True
        Call Get_고객조회("Tel", Trim(txtTel.Text))
        Tel_Flag = False
        
    End If
End Sub

Private Sub txtTel_GotFocus()
    Dim sTag    As String
    
    txtTel.SelStart = 0
    txtTel.SelLength = Len(txtTel.Text)
    
    sTag = Get_TagNo("", "R")
    
    cmdTagNo.Caption = Get_TagNo(sTag, "+")
    
    txtTel.BackColor = vbYellow '"&H0080FF80"
End Sub

Private Sub txtTEL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 40: txtHP.SetFocus   'Down Key
        Case 38: txtName.SetFocus 'Up Key
    End Select
    
    If KeyCode = vbKeyReturn And txtCode.Text = "" Then
        If Search_Flag = True Then Exit Sub
        
        Call Get_고객조회("Tel", Trim(txtTel.Text))
    End If
    
'    If KeyCode = vbKeyReturn Then
'        If Len(Trim(txtTel.Text)) < 4 Then
'            MsgBox "전화번호를 정확히 입력하십시요", vbCritical + vbOKOnly, "고객전화번호입력"
'            txtTel.SetFocus
'            Exit Sub
'        End If
'    End If
End Sub

Private Sub txtTel_LostFocus()
'    txtTel(0).Text = Replace(txtTel(0).Text, "'", "")
'    txtTel(1).Text = Replace(txtTel(1).Text, "'", "")
'
'    Query = "Update TB_고객정보 "
'    Query = Query & "SET 전화번호='" & txtTel(0).Text & "', "
'    Query = Query & "  전화2='" & txtTel(1).Text & "'  "
'    Query = Query & " WHERE 고객코드='" & txtCode.text & "' "
'    ADOCon.Execute Query
    
    txtTel.BackColor = "&H00FFFFFF"
End Sub

' 한벌 클릭
Private Sub cmdSuite_Click()
    On Error GoTo ErrRtn
        
    Dim ADORs       As ADODB.RecordSet
    Dim sGoodsStats As String
    Dim SuiteCode  As String
    Dim chkDate    As String
    Dim 의류명     As String
    Dim iEOF       As Boolean
    
    Dim strTemp(0 To 7) As String
    
    Call Chk_GroupGoodsReturn '세트 상품의 경우 다시 원상복구를 한다.
    
    For Each frm In Forms
        Select Case frm.Name
            Case "frm의류":     frm의류.Hide   ' Unload frm의류
            Case "frm색상표":   frm색상표.Hide ' Unload frm색상표
            Case "frm무늬":     Unload frm무늬 '
            Case "frm작업":     Unload frm작업 '
        End Select
    Next frm
    
    With sprGrid
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1: 의류명 = Trim(.Text) & ""
            
            If 의류명 = "" Then
                iCur = i - 1
                Exit For
            End If
        Next i
        
        .Row = iCur
        .Col = 8: SuiteCode = Trim(.Text) & "" '의류코드
        
    
        If SuiteCode = "f000" Or SuiteCode = "f003" Or SuiteCode = "f004" Or SuiteCode = "f007" Or SuiteCode = "f014" Or Left(SuiteCode, 1) = "a" Then
            '한벌 의류
        Else
            cmdSuite.Enabled = False
            Exit Sub
        End If
    
        '운동화 세탁
'        If Left(SuiteCode, 1) = "a" Then
'            If (Val(Mid(SuiteCode, 3, 2)) Mod 2) Then 'If (Val(Mid(SuiteCode, 2)) Mod 2) Then
'                cmdSuite.Enabled = False
'                Exit Sub
'            End If
'        End If
        
        Select Case SuiteCode
            Case "f000": SuiteCode = "g000" ' 정장상의
            Case "f003": SuiteCode = "g005" ' 실크상의
            Case "f004": SuiteCode = "g004" ' 마상의
            Case "f007": SuiteCode = "g007" ' 예복상의
            Case "f014": SuiteCode = "g015" ' 예복상의
            Case Else
                If Left(SuiteCode, 1) = "a" Then   '운동화 세탁
                    SuiteCode = Left(SuiteCode, 2) & Format(CStr(Val(Mid(SuiteCode, 3, 2) + 1)), "00")
                Else
                    cmdSuite.Enabled = False
                    
                    Exit Sub
                End If
        End Select
    End With
    
    strTemp(0) = ""
    chkDate = Format(Date, "YYYY-MM-DD")
    
    Set ADORs = New ADODB.RecordSet
    Set ADORs = Get_의류정보(SuiteCode, sGoodsStats)

    If ADORs.EOF Then
        ADORs.Close:    Set ADORs = Nothing
        cmdSuite.Enabled = False
        
        MsgBox "해당하는 상품 코드가 없습니다. 본사에 확인 바랍니다. [" & SuiteCode & "]", vbInformation, "확인"
        Exit Sub
    Else
        strTemp(0) = ADORs!의류명 & ""   '
        strTemp(3) = ADORs!금액 & ""     '
        strTemp(6) = ADORs!의류코드 & "" '
        
        frm접수.lblGoodsPriceStats.Caption = sGoodsStats
            
        ADORs.Close:    Set ADORs = Nothing
    End If
    
    With sprGrid
        .Row = iCur
        .Col = 3: strTemp(1) = .Text & ""            '색상
        .Col = 4: strTemp(7) = .Text & ""            '무늬
        .Col = 5: strTemp(2) = .Text & ""            '내용
        .Col = 7: strTemp(4) = .Text & ""            '상표
        If 가맹점정보.DualComputer = "Y" Then
            strTemp(5) = ""
        Else
             strTemp(5) = cmdTagNo.Caption      '택번호
        End If
        
        ' 고가 세탁이 있을 경우 3배를 해준다.
        If InStr(strTemp(2), "고") > 0 Then
            strTemp(3) = CStr(Val(strTemp(3)) * 3)
        End If
        
        ' 손 세탁이 있을 경우 2배를 해준다.
        If InStr(strTemp(2), "손") > 0 Then
            strTemp(3) = CStr(Val(strTemp(3)) * 2)
        End If
        
        ' 특정할인이 있을경우 해당 비율을 적용하여 할인해준다.
        If InStr(strTemp(2), "지") > 0 Then
            Dim iPercentage As Double
            
            iPercentage = (100 - 가맹점정보.특정할인비율) / 100  ' (할인이 20%일 경우 0.8의 값을 같는다.)
            
            strTemp(3) = CStr(Int((CDbl((Val(strTemp(3))) * iPercentage) / 100)) * 100)
            
            
        ' 아동복일 경우
        ElseIf InStr(strTemp(2), "아") > 0 Then
            
            ' 기본에서 20%로 할인한다.
            ' 10원단위를 절사 한다.
            .Col = 6:
            strTemp(3) = .Text
            'strTemp(3) = CStr(Int((CDbl((Val(strTemp(3))) * 0.8) / 100)) * 100)
        
        End If

        
        
        .Row = iCur + 1
        .Col = 1: .Text = strTemp(0) & "" '의류명
        .Col = 2: .Text = strTemp(5) & "" '택번호
        .Col = 3: .Text = strTemp(1) & "" '색상
        .Col = 4: .Text = strTemp(7) & "" '무늬
        .Col = 5: .Text = strTemp(2) & "" '내용
        .Col = 6: .Text = strTemp(3) & "" '금액
        .Col = 7: .Text = strTemp(4) & "" '상표
        .Col = 8: .Text = strTemp(6) & "" '코드
    
        '------------------------------------------------------------
        ' 마진 정보
        '------------------------------------------------------------
        Query = "SELECT * FROM TB_의류분류"
        Query = Query & " WHERE 의류분류코드 = '" & Left(SuiteCode, 2) & "'"
        Set SUBRs = New ADODB.RecordSet
        SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        If SUBRs.EOF Then
            .Col = 16: .Value = 0
            .Col = 17: .Value = 0
            .Col = 18: .Value = 0
        Else
            .Col = 16: .Value = SUBRs!세탁마진 & ""
            .Col = 17: .Value = SUBRs!외주마진 & ""
            .Col = 18: .Value = SUBRs!수선마진 & ""
        End If
        SUBRs.Close
        Set SUBRs = Nothing
        
        .Col = 14: .Text = strTemp(3) & "" '정상금액
        .Col = 20: .Text = Get_세탁정상금액(SuiteCode) & "" '의류금액
    End With
    
    If 가맹점정보.DualComputer = "Y" Then
    
    Else
        cmdTagNo.Caption = Get_TagNo(strTemp(5), "+") ' 택번호 증가
    End If
    
    iCur = iCur + 1
  
    cmdSuite.Enabled = False
  
    sprGrid.SetActiveCell 7, iCur ' Active Cell
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
 End Sub

Private Sub cmdSuite_LostFocus()
    cmdSuite.Enabled = False
End Sub

' 회원 등록
Private Sub cmdNew_Click()
'    frm고객수정.txtName.Text = txtName.Text & "" '
'    frm고객수정.txtTel.Text = txtTel.Text & ""   '
'    frm고객수정.txtHP.Text = txtHP.Text & ""     '
    
    frm고객수정.Show 1
End Sub

Private Sub sprGrid_Change(ByVal Col As Long, ByVal Row As Long)
    Dim 세탁요금  As Long
    Dim 세탁요금2 As Long
    Dim 의류코드  As String
    Dim 택번호    As String
    Dim sGoodsStats As String
    
    On Error GoTo ErrRtn
    
    If Col = 6 Then
        ' 현재의 택번호를 구한다.
        sprGrid.Row = sprGrid.ActiveRow ' iCur
         
        sprGrid.Col = 7                                 '상표
        If Trim(sprGrid.Text) = "짜집기(cm당)" Then
            Exit Sub
        End If
         
        sprGrid.Col = 2: 택번호 = sprGrid.Text & ""     '택번호
        sprGrid.Col = 6: 세탁요금2 = sprGrid.Value      '금액
        sprGrid.Col = 8: 의류코드 = sprGrid.Text & ""   '의류코드
        
        세탁요금 = Get_세탁금액(의류코드, sGoodsStats, btnInternet.tag)
        frm접수.lblGoodsPriceStats.Caption = sGoodsStats
        
        
        ' 드라이수선인 경우 수선금액을 (+)한다.
        sprGrid.Col = 5
        If sprGrid.Text = "드수" Then
            sprGrid.Col = 7
            
            Query = "SELECT ISNULL(금액,0) FROM TB_수선금액"
            Query = Query & " WHERE 수선내용 = '" & sprGrid.Text & "'"
            Set ADORs = New ADODB.RecordSet
            ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            If Not ADORs.EOF Then
                세탁요금 = 세탁요금 + ADORs(0)
            End If
            ADORs.Close
            Set ADORs = Nothing
        End If
        
        ' 본사에서 확인 코드를 받은 경우
        If 세탁요금 > 세탁요금2 Then
            If Not IsTagNo(chkPricPassWord) Or 택번호 <> chkPricPassWord Then
                MsgBox "규정금액" & " [" & Format(세탁요금, "#,##0") & "] 보다 입력금액 [" & Format(세탁요금2, "#,##0") & "]이 작습니다. ", vbInformation
                   
                sprGrid.Col = 6: sprGrid.Value = 세탁요금 & ""
                Exit Sub
            End If
        End If
        
        sprGrid.Col = 6: sprGrid.Value = 세탁요금2 & ""  '금액
        sprGrid.Col = 14: sprGrid.Value = 세탁요금2 & "" '정상가격
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub sprGrid_Click(ByVal Col As Long, ByVal Row As Long)
    Dim frm     As Form
    
    On Error GoTo ErrRtn
    
    If Row < 1 Then Exit Sub
    
    If Col = 15 Then Exit Sub '부속물 여부 체크 박스
        
    If Trim(txtTel.Text) = "" Then
        Beep
        Exit Sub
        
    ElseIf Trim(txtCode.Text) = "" Then
        MsgBox "고객코드 오류 입니다. 고객을 조회 후 작업하여 주세요     ", vbInformation, "확인"
        Beep
        
        txtCode.SetFocus
        
        Exit Sub
    End If
       
        
    iCur = Row '현재 Row
    
    For Each frm In Forms
         ' 무늬
        If frm.Name = Trim("frm무늬") Then frm무늬.Hide            'Unload frm무늬
         
         ' 색상
        If frm.Name = Trim("frm색상표") Then frm색상표.Hide        'Unload frm색상표
       
        ' 내용
        If frm.Name = Trim("frm의류") Then frm의류.Hide            'Unload frm의류
        
        ' 금액
        If frm.Name = Trim("frm작업") Then Unload frm작업
        
        ' 수선내용
        If frm.Name = Trim("frm세탁수선") Then Unload frm세탁수선
        
        ' 택번호 입력 상자
        If frm.Name = Trim("frm작업구분") Then Unload frm작업구분
        
        ' 인쇄
        If frm.Name = Trim("결제") Then
            Unload frm접수결제
            
            Call Chk_GroupGoodsReturn '세트 상품의 경우 다시 원상복구를 한다.
        End If
    Next frm

    sprGrid.Row = Row
    sprGrid.Col = 1
    
    If Trim(sprGrid.Text) = "" Then
        imgCapture.Picture = LoadPicture()
    Else
        Dim imgFileName As String
        
        sprGrid.Col = 2: imgFileName = btnTagCode.Caption & Replace(sprGrid.Text, "-", "")
        
        If Dir(App.Path & "\Capture\" & Format(Date, "YYYYMMDD") & imgFileName & ".jpg") = "" Then
            imgCapture.Picture = LoadPicture()
        Else
            imgCapture.Picture = LoadPicture(App.Path & "\Capture\" & Format(Date, "YYYYMMDD") & imgFileName & ".jpg")
            
            
        End If
    End If
    
    '
    sprGrid.SetActiveCell Col, Row
    DoEvents
    
    If Col = 1 Then
        frm의류.Show
    Else
        sprGrid.Col = 1
        If Trim(sprGrid.Text) = "" Then
            Exit Sub
        End If
        
        Select Case Col
            Case 3
                frm색상표.Show 'Load frm색상표 ' 23
                
            Case 4
                frm무늬.Show   'Load frm무늬 ' 23
                
            Case 5
                frm작업.Show   'Load frm작업 '24
                
            Case 6
                frm작업.Show   'Load frm작업 '24
                frm작업.TabControl.SelectedItem = 1
                'frm작업.pnlMoney.Visible = True
                
            Case 7                                  '2002/11/14 반품및 재세탁시 구 택번호 입력기능 추가
                sprGrid.Col = 5
                
                Select Case sprGrid.Text
                    Case "드재":
                        'Load frm작업구분
                        
                        frm작업구분.SetFlags "재세탁 구분"
                        frm작업구분.Show
                    
                    Case "드반":
                        'Load frm작업구분
                        
                        frm작업구분.SetFlags "반품 구분"
                        frm작업구분.Show
                    
                    Case "드사":
                        'Load frm작업구분
                        
                        frm작업구분.SetFlags "사고품 구분"
                        frm작업구분.Show
                    
                    Case "드재다":
                        'Load frm작업구분
                        
                        frm작업구분.SetFlags "재다림질 구분"
                        frm작업구분.Show
                End Select
        End Select
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub sprGrid_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Col = 7 And sprGrid.Value = "짜집기(cm당)" Then
        cmdOK.SetFocus
    End If
End Sub

Private Sub sprGrid_GotFocus()
    sprGrid.CursorType = 2
End Sub

Private Sub sprGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim varTemp As String
    
    On Error GoTo ErrRtn
    
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        
        ' 원금액을 기록한다.
        If sprGrid.ActiveRow <= 0 Then Exit Sub
        
        Call sprGrid_Click(sprGrid.ActiveCol, sprGrid.ActiveRow)
            
        '-------------------------------------------------------
        '
        '-------------------------------------------------------
        sprGrid.Row = sprGrid.ActiveRow
        sprGrid.Col = 6:  varTemp = sprGrid.Value ' 6-금액
        sprGrid.Col = 14: sprGrid.Value = varTemp '14-정상가격
        
        sprGrid.Col = 2
        If Trim(sprGrid.Text) <> "" Then
            sprGrid.Col = 6: sprGrid.EditMode = False '금액
            
            Call cmdOK_Click  '98/07/22
        End If
    End If

    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub sprGrid_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
     If NewRow <> -1 Then
         If Row <> NewRow Then
            Call Check_SpreadOrder
         End If
    End If
End Sub

Private Sub sprGrid_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    sprGrid.CursorType = 2
End Sub

Public Sub 고객정보_Display(고객코드 As String)
    Dim CustRs  As ADODB.RecordSet
    
    On Error GoTo ErrRtn
        
    TabControl1.SelectedItem = 0
        
    Query = "SELECT    고객코드"
    Query = Query & ", 성명"
    Query = Query & ", 전화번호"
    Query = Query & ", 휴대전화"
    Query = Query & ", 주소"
    Query = Query & ", 메모"
    Query = Query & ", ISNULL(미수금액,0) AS 미수금액"
    Query = Query & ", ISNULL(고객등급코드,'C') AS 고객등급코드"
    Query = Query & ", 등록일자"
    Query = Query & ", 이용횟수"
    Query = Query & ", 총접수금액"
    Query = Query & ", 누적마일리지"
    Query = Query & ", 사용가능마일리지"
    Query = Query & " FROM TB_고객정보"
    Query = Query & " WHERE 고객코드 = '" & 고객코드 & "'"
    Set CustRs = New ADODB.RecordSet
    CustRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If CustRs.EOF Then
        CustRs.Close
        Set CustRs = Nothing
    
        Call Text_Clear
        
        With 마일리지
            .사용가능마일리지 = 0
            .누적마일리지 = 0
        End With
    Else
        txtCode.Text = CustRs!고객코드 & ""                       ' 1
        txtName.Text = Trim(CustRs!성명) & ""                     ' 2
        txtTel.Text = Trim(CustRs!전화번호) & ""                  ' 3
        txtHP.Text = Trim(CustRs!휴대전화) & ""                   ' 4
        txtHP.tag = txtHP.Text & ""                               '
        txtAddress.Text = Trim(CustRs!주소) & ""                  ' 5
                
        '미수금액이 마이너스금액인 경우 미반환금액으로 표기
        If CustRs!미수금액 >= 0 Then
            txtMisu.Value = CustRs!미수금액 & ""                     '
            txtNoRepay.Value = 0                                    '
        Else
            txtMisu.Value = 0                                       '
            txtNoRepay.Value = CustRs!미수금액 & ""                  '
        End If
        
        txtMemo.Text = Trim(CustRs!메모) & ""                     ' 7
        txtRegistDay.Text = Format(CustRs!등록일자, "YYYY-MM-DD") ' 8
        
        Select Case CustRs!고객등급코드                           '
            Case "C":  pnlCustom.BackColor = vbBlue
            Case "D":  pnlCustom.BackColor = vbYellow
            Case "E":  pnlCustom.BackColor = vbRed
            Case Else: pnlCustom.BackColor = vbWhite
        End Select
        
        With cboClass                                            ' 9
            For i = 0 To .ListCount - 1
                If Left(.List(i), 1) = CustRs!고객등급코드 Then
                    .ListIndex = i
                    
                    Exit For
                End If
            Next i
        End With
            
        txtVisit.Value = CustRs!이용횟수 & ""                     '10
        
        With 마일리지
            .사용가능마일리지 = CustRs!사용가능마일리지 & ""      '11
            .누적마일리지 = CustRs!누적마일리지 & ""              '12
        End With

        CustRs.Close
        Set CustRs = Nothing
        
        '-----------------------------------------------------------
        ' TB_매출
        '-----------------------------------------------------------
        Query = "SELECT    ISNULL(SUM(접수금액),0) AS 접수금액"
        Query = Query & ", ISNULL(SUM(입금합계),0) AS 입금액"
        Query = Query & ", ISNULL(SUM(세트할인),0) AS 세트할인"
        Query = Query & ", ISNULL(SUM(에누리),0)   AS 에누리"
        
        'Query = Query & ", ISNULL(SUM(현금입금+카드입금+쿠폰입금),0) AS 입금액"
        'Query = Query & ", ISNULL(SUM(세트할인+에누리),0) AS 할인액"
        
        Query = Query & "  FROM TB_매출"
        Query = Query & "  WHERE 고객코드 = '" & 고객코드 & "'"
        Set CustRs = New ADODB.RecordSet
        CustRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

        If CustRs.EOF Then
            txtTotalNum(0).Value = 0                   '총접수금액
            txtTotalNum(1).Value = 0                   '총입금액
            txtTotalNum(2).Value = 0                   '할인금액
            
            ' << 미수금액은 TB_고객정보의 미수금을 이용 >>
            'txtMisu.Value = 0                          '미수금액
        Else
            txtTotalNum(0).Value = CustRs!접수금액 & ""                                      '총접수금액
            txtTotalNum(1).Value = CustRs!입금액 & ""                                        '총입금액
            If CustRs!세트할인 + CustRs!에누리 & "" > 0 Then
                txtTotalNum(2).Value = CustRs!세트할인 + CustRs!에누리 & ""                       '할인금액
            Else
                txtTotalNum(2).Value = 0                       '할인금액
            End If
            
            ' << 미수금액은 TB_고객정보의 미수금을 이용 >>
            'txtMisu.Value = CustRs!접수금액 - (CustRs!입금액 + CustRs!세트할인 + CustRs!에누리) '미수금액
        End If
        CustRs.Close
        Set CustRs = Nothing
    End If
    
    Debug.Print "이용실적_Display " & Now
    Call 이용실적_Display                            '이용실적
    
    Debug.Print "최근접수_Display " & Now
    Call 최근접수_Display(txtCode.Text)              '최근 접수건수
    
    Debug.Print "사고품_Display in " & Now
    Call 사고품_Display(txtCode.Text)                '
    Debug.Print "사고품_Display out " & Now
    
    txtUseMileage.Value = 마일리지.사용가능마일리지  '
    txtTotalMileage.Value = 마일리지.누적마일리지    '
    
    txtCode.Locked = True
    txtTel.Locked = True
    txtHP.Locked = True
    txtName.Locked = True
    txtAddress.Locked = True
    txtMemo.Locked = True
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

'-------------------------------------------------------------------
' 함수명 : Get_고객조회
'
'-------------------------------------------------------------------
Private Sub Get_고객조회(Gbn As String, strFind As String)
    Dim SSQL    As String
    On Error GoTo ErrRtn
    
    
    ' 마진 정보의 변경 여부를 확인 한다.
    ' If Check_세탁마진 = False Then End
    
    Search_Flag = True
    
    lblSamSungCardCheck.tag = "N"
    
    SSQL = "SELECT * FROM TB_고객정보"
    
    Select Case Gbn
        Case "Code"
            SSQL = SSQL & " WHERE 고객코드 = '" & strFind & "'"
            SSQL = SSQL & " ORDER BY 고객코드 ASC"
        
        Case "Tel"
            SSQL = SSQL & " WHERE (전화번호 LIKE '%" & strFind & "'"
            SSQL = SSQL & "   OR   휴대전화   LIKE '%" & strFind & "')"
            SSQL = SSQL & " ORDER BY 전화번호, 휴대전화 ASC"
            
        Case "Name"
            SSQL = SSQL & " WHERE 성명 LIKE '%" & strFind & "%'"
            SSQL = SSQL & " ORDER BY 성명 ASC"
        
        Case "Addr"
            SSQL = SSQL & " WHERE 주소 LIKE '%" & strFind & "%'"
            SSQL = SSQL & " ORDER BY 주소 ASC"
    End Select
    
    Set ADORs = New ADODB.RecordSet
    ADORs.Open SSQL, ADOCon, adOpenStatic, adLockOptimistic
       
    If ADORs.EOF Then
        ADORs.Close
        Set ADORs = Nothing
        
        If MsgBox("등록되지 않은 고객입니다. 등록 하시겠습니까?", vbInformation + vbYesNo, "회원 입력 확인") = vbYes Then
            Call cmdNew_Click
        Else
            Select Case Gbn
                Case "Code"
                    txtCode.Text = ""
                    txtCode.SetFocus
                    
                Case "Tel"
                    txtTel.Text = ""
                    txtTel.SetFocus
                    
                Case "Name"
                    txtName.Text = ""
                    txtName.SetFocus
                    
                Case "Addr"
                    txtAddress.Text = ""
                    txtAddress.SetFocus
            End Select
        End If

    ElseIf ADORs.RecordCount = 1 Then
        txtCode.Text = ADORs!고객코드 & ""
        
        ADORs.Close
        Set ADORs = Nothing
        
        Search_Flag = False
        
        Call 고객정보_Display(txtCode.Text)
        
        sprGrid.SetFocus
        sprGrid.SetActiveCell 1, 1
        
    ElseIf ADORs.RecordCount >= 2 Then
        ADORs.Close
        Set ADORs = Nothing
        
        DoEvents
        
        frm고객검색.DataDisplay SSQL
        frm고객검색.Show 1
        DoEvents
        
        If frm고객검색.SELECTCODE = "CANCEL" Then
            
            Search_Flag = False
            
            Select Case Gbn
                Case "Code":
                    txtCode.SetFocus
                    Exit Sub
                    
                Case "Tel":
                    txtTel.SetFocus
                    Exit Sub
                    
                Case "Name":
                    txtName.SetFocus
                    Exit Sub
                    
                Case "Addr":
                    txtAddress.SetFocus
                    Exit Sub
            End Select
        End If
        
        txtCode.Text = 고객정보.고객코드 & ""
        
        Call 고객정보_Display(txtCode.Text)
        
        With sprGrid
            .SetFocus
            .SetActiveCell 1, 1
        
            .Row = 1:   .Col = 1
            .BackColor = vbGreen
            
        End With
        
    End If
    
'    '-------------------------------------------------------------
'    ' 요일세일 - 일,월,화,수,목,금,토
'    '-------------------------------------------------------------
    
    chkDaySale = Get_요일행사여부
    
    
    Search_Flag = False
    Exit Sub
    
ErrRtn:
    Search_Flag = False
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

'--------------------------------------------------------------------------------
' 함수명 : Check_SpreadOrder
'
'
'--------------------------------------------------------------------------------
Private Function Check_SpreadOrder() As Boolean
    Dim nRow As Long
    
    On Error GoTo ErrRtn
    
    Check_SpreadOrder = True
    
    For nRow = 1 To sprGrid.DataRowCnt
        sprGrid.Row = nRow
        sprGrid.Col = 5
        
        If sprGrid.Text = "드재" Or sprGrid.Text = "드아재" Then
            sprGrid.Col = 7
            
            If sprGrid.Text = "" Then
                MsgBox "재세탁 택번호를 입력하셔야 합니다.", vbInformation
                sprGrid.Row = nRow
                sprGrid.Col = 7
                sprGrid.Action = ActionActiveCell
                
                Check_SpreadOrder = False
                
                Exit Function
            End If
            
        ElseIf InStr(sprGrid.Text, "반") Then
            sprGrid.Col = 7
            
            If sprGrid.Text = "" Then
                MsgBox "반품 택번호를 입력하셔야 합니다.", vbInformation
                sprGrid.Row = nRow
                sprGrid.Col = 7
                sprGrid.Action = ActionActiveCell
                
                Check_SpreadOrder = False
                
                Exit Function
            End If
            
        ElseIf InStr(sprGrid.Text, "사") > 0 Then
            sprGrid.Col = 7
            
            If sprGrid.Text = "" Then
                MsgBox "사고품의 택번호를 입력하셔야 합니다.", vbInformation
                sprGrid.Row = nRow:   sprGrid.Col = 7
                sprGrid.Action = ActionActiveCell
                
                Check_SpreadOrder = False
                Exit Function
            End If
            
        ElseIf InStr(sprGrid.Text, "재다") > 0 Then
            sprGrid.Col = 7
            
            If sprGrid.Text = "" Then
                MsgBox "재다림질 택번호를 입력하셔야 합니다.", vbInformation
                sprGrid.Row = nRow:   sprGrid.Col = 7
                sprGrid.Action = ActionActiveCell
                
                Check_SpreadOrder = False
                Exit Function
            End If
        End If
        
        ' 입력된 택번호 확인
        If 가맹점정보.DualComputer = "Y" Then
        
        Else
            sprGrid.Col = 1
            If Trim(sprGrid.Text) <> "" Then
                sprGrid.Col = 2
                If Len(sprGrid.Text) <> 7 And txtCode.Text <> "" Then
                    MsgBox "택번호가 올바르지 않습니다.", vbInformation
        
                    Check_SpreadOrder = False
                End If
                Exit Function
            End If
        End If
    Next nRow
    
    Exit Function
    
ErrRtn:
    Check_SpreadOrder = False
    
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Function

'최근 접수건수
Private Sub 최근접수_Display(sCode As String)
    Dim SSQL    As String
    On Error GoTo ErrRtn
    
    SSQL = "SELECT "
    SSQL = SSQL & "  접수일자"
    SSQL = SSQL & ", SUM(금액)"
    SSQL = SSQL & " FROM TB_입출고"
    SSQL = SSQL & " WHERE 고객코드 = '" & sCode & "'"
    SSQL = SSQL & "   AND ((판매취소 <> 'Y')"
    'SSQL = SSQL & "   AND ((판매취소 = '') AND (판매취소일자 IS NULL OR 판매취소일자 = '')"
    SSQL = SSQL & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
    SSQL = SSQL & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
    SSQL = SSQL & " GROUP BY 접수일자"
    SSQL = SSQL & " ORDER BY 접수일자 DESC"
    Set SUBRs = New ADODB.RecordSet
    SUBRs.Open SSQL, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprHist
        .MaxRows = 0
        .ReDraw = False
        
        Do Until SUBRs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = Format(SUBRs(0) & "", "YYYY-MM-DD") '
            .Col = 2: .Text = SUBRs(1) & ""                       '
            
            SUBRs.MoveNext
        Loop
        SUBRs.Close
        Set SUBRs = Nothing
        
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

'-------------------------------------------------------------------------
' 함수명 : Chk_GroupGoodsReturn
'
' 세트 상품의 경우 다시 원상복구를 한다.
'-------------------------------------------------------------------------
Private Sub Chk_GroupGoodsReturn()
    Dim iRow    As Long
    Dim iPrice  As Long

    '세트 상품의 경우 다시 원상복구를 한다.
    For iRow = 1 To sprGrid.MaxRows
        sprGrid.Row = iRow
        
        sprGrid.Col = 1
        If sprGrid.Text = "" Then Exit For
        
        '원금액을 보관한다.
        sprGrid.Col = 14
        If sprGrid.Value <> "" Then
            sprGrid.Col = 14: iPrice = sprGrid.Value & "" '정상금액
            sprGrid.Col = 6:  sprGrid.Value = iPrice & "" '금액
        End If
    Next iRow
End Sub

'사고품
Private Sub 사고품_Display(sCode As String)
    Dim SSQL        As String
    
    On Error GoTo ErrRtn
    
    If Len(Trim(txtHP.Text)) <= 4 Then Exit Sub
    
    SSQL = "SELECT    일련번호"
    SSQL = SSQL & ", 접수일자"
    SSQL = SSQL & ", 성명"
    SSQL = SSQL & ", 전화번호"
    SSQL = SSQL & ", 휴대전화"
    SSQL = SSQL & ", 의류명"
    SSQL = SSQL & ", 색상"
    SSQL = SSQL & ", 상표"
    SSQL = SSQL & ", 구입일자"
    SSQL = SSQL & ", 구입처"
    SSQL = SSQL & ", 구입형태"
    SSQL = SSQL & ", 구입가격"
    SSQL = SSQL & ", 사고접수일자"
    SSQL = SSQL & ", 크레임구분" '사고종류
    SSQL = SSQL & ", 가맹점의견" '사고내용
    SSQL = SSQL & ", 본사의견"   '사고의견
    SSQL = SSQL & ", 보상금액"
    'SSQL = SSQL & ", 합의금액"
    SSQL = SSQL & ", 처리구분"
    SSQL = SSQL & ", 가맹점코드"
    SSQL = SSQL & ", 가맹점명"
    SSQL = SSQL & " FROM TB_사고품내역"
    
    SSQL = SSQL & " WHERE 휴대전화 LIKE '%" & Trim(txtHP.Text) & "%'"
    
    SSQL = SSQL & " ORDER BY 접수일자, 일련번호 ASC"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open SSQL, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprClaim
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = ADORs!일련번호 & ""
            .Col = 2:  .Text = ADORs!접수일자 & ""
            .Col = 3:  .Text = ADORs!보상금액 & ""
            .Col = 4:  .Text = ADORs!의류명 & ""
            .Col = 5:  .Text = ADORs!구입가격 & ""
            .Col = 6:  .Text = ADORs!색상 & ""
            .Col = 7:  .Text = ADORs!상표 & ""
            .Col = 8:  .Text = ADORs!구입일자 & ""
            .Col = 9:  .Text = ADORs!구입처 & ""
            .Col = 10: .Text = ADORs!구입형태 & ""
            .Col = 11: .Text = ADORs!사고접수일자 & ""
            .Col = 12: .Text = ADORs!크레임구분 & ""
            .Col = 13: .Text = ADORs!가맹점의견 & "" '사고내용
            .Col = 14: .Text = ADORs!본사의견 & ""   '사고의견
            .Col = 15: .Text = ADORs!보상금액 & ""
            .Col = 16: .Text = ADORs!처리구분 & ""
            .Col = 17: .Text = ADORs!가맹점코드 & ""
            .Col = 18: .Text = ADORs!가맹점명 & ""
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
            
        .ReDraw = True
        
        If .MaxRows > 0 Then
            TabControl1.SelectedItem = 3
        End If
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub Resize_Rtn()
    sprClaim.Width = TabControl1.Width - 810
End Sub
