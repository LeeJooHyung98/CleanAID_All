VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm판매취소 
   BorderStyle     =   1  '단일 고정
   Caption         =   "판매취소 - 결제"
   ClientHeight    =   8970
   ClientLeft      =   9840
   ClientTop       =   2340
   ClientWidth     =   12450
   ControlBox      =   0   'False
   DrawWidth       =   3
   FillColor       =   &H00C0C0C0&
   Icon            =   "frm판매취소.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form10"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   12450
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8970
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12450
      _ExtentX        =   21960
      _ExtentY        =   15822
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm판매취소.frx":0A02
      Begin Threed.SSPanel SSPanel 
         Height          =   675
         Index           =   0
         Left            =   15
         TabIndex        =   29
         Top             =   8280
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   1191
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton btnAccount 
            Height          =   555
            Index           =   0
            Left            =   4305
            TabIndex        =   30
            Top             =   60
            Width           =   1590
            _Version        =   851970
            _ExtentX        =   2805
            _ExtentY        =   979
            _StockProps     =   79
            Caption         =   " 완불결제"
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
            Picture         =   "frm판매취소.frx":0AF4
         End
         Begin XtremeSuiteControls.PushButton btnAccount 
            Height          =   555
            Index           =   1
            Left            =   6015
            TabIndex        =   35
            Top             =   60
            Width           =   1590
            _Version        =   851970
            _ExtentX        =   2805
            _ExtentY        =   979
            _StockProps     =   79
            Caption         =   "후불결제"
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
         End
      End
      Begin XtremeSuiteControls.TabControl TabControl1 
         Height          =   3375
         Left            =   15
         TabIndex        =   6
         Top             =   15
         Width           =   12420
         _Version        =   851970
         _ExtentX        =   21908
         _ExtentY        =   5953
         _StockProps     =   68
         Appearance      =   3
         Color           =   16
         PaintManager.BoldSelected=   -1  'True
         PaintManager.HotTracking=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         PaintManager.ButtonMargin=   "3,2,3,2"
         ItemCount       =   2
         Item(0).Caption =   " 접수 "
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "TabControlPage1"
         Item(1).Caption =   " 결제 "
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "TabControlPage2"
         Begin XtremeSuiteControls.TabControlPage TabControlPage2 
            Height          =   2925
            Left            =   -69970
            TabIndex        =   8
            Top             =   420
            Visible         =   0   'False
            Width           =   12360
            _Version        =   851970
            _ExtentX        =   21802
            _ExtentY        =   5159
            _StockProps     =   1
            Page            =   1
            Begin FPSpreadADO.fpSpread sprAccount 
               Height          =   2820
               Left            =   45
               TabIndex        =   9
               Top             =   45
               Width           =   12285
               _Version        =   524288
               _ExtentX        =   21669
               _ExtentY        =   4974
               _StockProps     =   64
               BackColorStyle  =   1
               ColsFrozen      =   2
               DAutoCellTypes  =   0   'False
               DAutoHeadings   =   0   'False
               DAutoSave       =   0   'False
               DAutoSizeCols   =   0
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
               MaxCols         =   12
               Protect         =   0   'False
               ScrollBars      =   2
               ShadowColor     =   14737632
               SpreadDesigner  =   "frm판매취소.frx":11EE
               Appearance      =   1
               HighlightHeaders=   1
               HighlightStyle  =   1
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage1 
            Height          =   2925
            Left            =   30
            TabIndex        =   7
            Top             =   420
            Width           =   12360
            _Version        =   851970
            _ExtentX        =   21802
            _ExtentY        =   5159
            _StockProps     =   1
            Page            =   0
            Begin FPSpreadADO.fpSpread sprGrid 
               Height          =   2820
               Left            =   45
               TabIndex        =   36
               Top             =   45
               Width           =   12285
               _Version        =   524288
               _ExtentX        =   21669
               _ExtentY        =   4974
               _StockProps     =   64
               AutoCalc        =   0   'False
               BackColorStyle  =   1
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
               MaxCols         =   18
               MaxRows         =   20
               OperationMode   =   2
               SelectBlockOptions=   0
               SpreadDesigner  =   "frm판매취소.frx":1B6A
               UserResize      =   1
               VisibleCols     =   7
               VisibleRows     =   15
               Appearance      =   1
               HighlightHeaders=   1
               HighlightStyle  =   1
               ScrollBarStyle  =   2
            End
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   405
         Index           =   1
         Left            =   7695
         TabIndex        =   2
         Top             =   3405
         Width           =   4740
         _ExtentX        =   8361
         _ExtentY        =   714
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " 신용카드 승인내역"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm판매취소.frx":46D3
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin FPSpreadADO.fpSpread sprCard 
         Height          =   1305
         Left            =   7695
         TabIndex        =   3
         Top             =   3825
         Width           =   4740
         _Version        =   524288
         _ExtentX        =   8361
         _ExtentY        =   2302
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         EditModePermanent=   -1  'True
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
         MaxCols         =   13
         SpreadDesigner  =   "frm판매취소.frx":48F9
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   420
         Index           =   2
         Left            =   7695
         TabIndex        =   4
         Top             =   5145
         Width           =   4740
         _ExtentX        =   8361
         _ExtentY        =   741
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " 현금영수증 승인내역"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm판매취소.frx":522E
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton btnCashCancel 
            Height          =   390
            Left            =   2745
            TabIndex        =   5
            Top             =   15
            Width           =   1965
            _Version        =   851970
            _ExtentX        =   3466
            _ExtentY        =   688
            _StockProps     =   79
            Caption         =   "현금영수증 승인취소"
            ForeColor       =   0
            UseVisualStyle  =   -1  'True
         End
      End
      Begin FPSpreadADO.fpSpread sprCash 
         Height          =   3375
         Left            =   7695
         TabIndex        =   10
         Top             =   5580
         Width           =   4740
         _Version        =   524288
         _ExtentX        =   8361
         _ExtentY        =   5953
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayColHeaders=   0   'False
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
         MaxCols         =   1
         MaxRows         =   11
         RowHeaderDisplay=   0
         ScrollBars      =   2
         SpreadDesigner  =   "frm판매취소.frx":5454
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   4860
         Left            =   15
         TabIndex        =   11
         Top             =   3405
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   8573
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSPanel SSPanel 
            Height          =   30
            Index           =   1
            Left            =   135
            TabIndex        =   31
            Top             =   1920
            Width           =   7440
            _ExtentX        =   13123
            _ExtentY        =   53
            _Version        =   262144
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnCard 
            Height          =   420
            Left            =   6075
            TabIndex        =   12
            Top             =   3870
            Width           =   1515
            _Version        =   851970
            _ExtentX        =   2672
            _ExtentY        =   741
            _StockProps     =   79
            Caption         =   "신용카드결제"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnCash 
            Height          =   420
            Left            =   6075
            TabIndex        =   13
            Top             =   3420
            Width           =   1515
            _Version        =   851970
            _ExtentX        =   2672
            _ExtentY        =   741
            _StockProps     =   79
            Caption         =   "현금영수증"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin CSTextLibCtl.silgEdit txtCash 
            Height          =   420
            Left            =   4485
            TabIndex        =   14
            Top             =   3420
            Width           =   1575
            _Version        =   262145
            _ExtentX        =   2778
            _ExtentY        =   741
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   11.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            ReadOnly        =   -1  'True
            FocusSelect     =   -1  'True
            Insert          =   0   'False
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "0"
            Text            =   " 0"
            StartText.x     =   3
            StartText.y     =   5
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   18
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            Undo            =   1
            Data            =   0
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   420
            Index           =   12
            Left            =   2925
            TabIndex        =   15
            Top             =   3420
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   741
            _Version        =   262144
            BackColor       =   12648384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "현 금 결 제"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm판매취소.frx":5A50
            BorderWidth     =   0
            BevelOuter      =   1
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   420
            Index           =   13
            Left            =   2925
            TabIndex        =   16
            Top             =   3870
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   741
            _Version        =   262144
            BackColor       =   12648384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "카 드 결 제"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm판매취소.frx":5D92
            BorderWidth     =   0
            BevelOuter      =   1
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.silgEdit txtCard 
            Height          =   420
            Left            =   4485
            TabIndex        =   17
            Top             =   3870
            Width           =   1575
            _Version        =   262145
            _ExtentX        =   2778
            _ExtentY        =   741
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   11.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            ReadOnly        =   -1  'True
            FocusSelect     =   -1  'True
            Insert          =   0   'False
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "0"
            Text            =   " 0"
            StartText.x     =   3
            StartText.y     =   5
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   18
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            Undo            =   1
            Data            =   0
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   420
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   2070
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   741
            _Version        =   262144
            BackColor       =   12648384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "고객코드"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm판매취소.frx":60D4
            BorderWidth     =   0
            BevelOuter      =   1
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   420
            Index           =   1
            Left            =   120
            TabIndex        =   19
            Top             =   2520
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   741
            _Version        =   262144
            BackColor       =   12648384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "접수번호"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm판매취소.frx":6416
            BorderWidth     =   0
            BevelOuter      =   1
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCode 
            Height          =   420
            Left            =   1680
            TabIndex        =   20
            Top             =   2070
            Width           =   1095
            _ExtentX        =   1931
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
            Caption         =   "0"
            PictureBackgroundStyle=   2
            BorderWidth     =   0
            BevelOuter      =   1
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlNum 
            Height          =   420
            Left            =   1680
            TabIndex        =   21
            Top             =   2520
            Width           =   1095
            _ExtentX        =   1931
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
            Caption         =   "0"
            PictureBackgroundStyle=   2
            BorderWidth     =   0
            BevelOuter      =   1
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   420
            Index           =   10
            Left            =   2925
            TabIndex        =   22
            Top             =   2520
            Width           =   1530
            _ExtentX        =   2699
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
            Caption         =   "받 은 금 액"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm판매취소.frx":6758
            BorderWidth     =   0
            BevelOuter      =   1
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   420
            Index           =   11
            Left            =   2925
            TabIndex        =   23
            Top             =   2970
            Width           =   1530
            _ExtentX        =   2699
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
            Caption         =   "거 스 름 돈"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm판매취소.frx":6A9A
            BorderWidth     =   0
            BevelOuter      =   1
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.silgEdit txtIncome 
            Height          =   420
            Left            =   4485
            TabIndex        =   0
            Top             =   2520
            Width           =   1575
            _Version        =   262145
            _ExtentX        =   2778
            _ExtentY        =   741
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   11.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Insert          =   0   'False
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "0"
            Text            =   " 0"
            StartText.x     =   3
            StartText.y     =   5
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   18
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            Undo            =   1
            Data            =   0
         End
         Begin CSTextLibCtl.silgEdit txtChange 
            Height          =   420
            Left            =   4485
            TabIndex        =   24
            Top             =   2970
            Width           =   1575
            _Version        =   262145
            _ExtentX        =   2778
            _ExtentY        =   741
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   11.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            ReadOnly        =   -1  'True
            FocusSelect     =   -1  'True
            Insert          =   0   'False
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "0"
            Text            =   " 0"
            StartText.x     =   3
            StartText.y     =   5
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   18
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            Undo            =   1
            Data            =   0
         End
         Begin CSTextLibCtl.silgEdit txtMisu 
            Height          =   420
            Left            =   4485
            TabIndex        =   25
            Top             =   2070
            Width           =   1575
            _Version        =   262145
            _ExtentX        =   2778
            _ExtentY        =   741
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   11.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            ReadOnly        =   -1  'True
            FocusSelect     =   -1  'True
            Insert          =   0   'False
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "0"
            Text            =   " 0"
            StartText.x     =   3
            StartText.y     =   5
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   18
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            Undo            =   1
            Data            =   0
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   420
            Index           =   2
            Left            =   2925
            TabIndex        =   26
            Top             =   2070
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   741
            _Version        =   262144
            Font3D          =   1
            BackColor       =   12648384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "판매취소후 잔액"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm판매취소.frx":6DDC
            BorderWidth     =   0
            BevelOuter      =   1
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   420
            Index           =   14
            Left            =   2925
            TabIndex        =   27
            Top             =   4320
            Width           =   1530
            _ExtentX        =   2699
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
            Caption         =   "결제후 잔액"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm판매취소.frx":711E
            BorderWidth     =   0
            BevelOuter      =   1
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.silgEdit txtBalance2 
            Height          =   420
            Left            =   4485
            TabIndex        =   28
            Top             =   4320
            Width           =   1575
            _Version        =   262145
            _ExtentX        =   2778
            _ExtentY        =   741
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   255
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   11.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            ReadOnly        =   -1  'True
            FocusSelect     =   -1  'True
            Insert          =   0   'False
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "0"
            Text            =   " 0"
            StartText.x     =   3
            StartText.y     =   5
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   18
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            Undo            =   1
            Data            =   0
         End
         Begin VB.Label Label 
            BackStyle       =   0  '투명
            Caption         =   "판매취소 후 잔액을 결제를 받지 않는 경우에는 그냥 '후불결제' 버튼을 클릭하십시요."
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   390
            Index           =   3
            Left            =   120
            TabIndex        =   34
            Top             =   1365
            Width           =   7410
         End
         Begin VB.Label Label 
            BackStyle       =   0  '투명
            Caption         =   "그리고 판매취소 하지 않는 접수건에 대해서 결제는 아래에서 현금 또는 카드결제를 하신후 '완불결제' 버튼을 클릭하시면 됩니다."
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Index           =   1
            Left            =   120
            TabIndex        =   33
            Top             =   825
            Width           =   7410
         End
         Begin VB.Label Label 
            BackStyle       =   0  '투명
            Caption         =   $"frm판매취소.frx":7460
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   600
            Index           =   0
            Left            =   735
            TabIndex        =   32
            Top             =   135
            Width           =   6810
         End
         Begin VB.Image Image 
            Height          =   480
            Left            =   135
            Picture         =   "frm판매취소.frx":74F8
            Top             =   180
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm판매취소"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Data_Display(접수일자 As String, 접수번호 As Long, 고객코드 As String)
    '----------------------------------------------------------------------------
    ' 입고
    '----------------------------------------------------------------------------
    Query = "SELECT    접수일자"
    Query = Query & ", ISNULL(가맹점입고일자, '') AS 가맹점입고일자"
    Query = Query & ", 의류명"
    Query = Query & ", 택번호"
    Query = Query & ", 색상"
    Query = Query & ", 무늬"
    Query = Query & ", 내용"
    Query = Query & ", 금액"
    Query = Query & ", 결제여부"
    Query = Query & ", ISNULL(지사출고상태, '') AS 지사출고상태"
    Query = Query & ", 상표"
    Query = Query & ", 부모택번호"
    Query = Query & ", 수선금액"
    Query = Query & ", 세트Key"
    Query = Query & ", 세트구분"
    Query = Query & ", 오점내용"
    Query = Query & ", 접수번호"
    Query = Query & ", 판매취소"
    Query = Query & " FROM TB_입출고"
    Query = Query & " WHERE 고객코드 = '" & 고객코드 & "'"
    Query = Query & "   AND 접수일자 = '" & 접수일자 & "'"
    Query = Query & "   AND 접수번호 =  " & 접수번호
    Query = Query & " ORDER BY 택번호 ASC"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
                                            
            .Col = 1:  .Text = Format(ADORs!접수일자, "YY-MM-DD") & ""            ' 1
            .Col = 2:  .Text = IIf(ADORs!가맹점입고일자 = "", "0", "1")           ' 2
            .Col = 3:  .Text = ADORs!의류명 & ""                                  ' 3
            
            If Len(Trim(ADORs!택번호)) <= 6 Then
                .Col = 4:  .Text = ADORs!택번호 & ""                              '
            Else
                .Col = 4:  .Text = Mid(ADORs!택번호, 4, 2) & "-" & Mid(ADORs!택번호, 6, 4) ' 4
            End If
            
            .Col = 5:  .Text = ADORs!색상 & ""                                    ' 5
            .Col = 6:  .Text = ADORs!무늬 & ""                                    ' 6
            .Col = 7:  .Text = ADORs!내용 & ""                                    ' 7
            .Col = 8:  .Text = ADORs!금액 & ""                                    ' 8
            .Col = 9:  .Text = ADORs!결제여부 & ""                                ' 9
            
            If ADORs!결제여부 = "미불" Then
                .ForeColor = vbRed                                                '
            Else
                .ForeColor = vbBlack                                              '
            End If
                            
            .Col = 10:  .Text = ADORs!상표 & ""                                   '10
            
            If (Trim(ADORs!택번호) = Trim(ADORs!부모택번호)) Or (Trim(ADORs!부모택번호) = "") Then
                .Col = 11: .Text = ""
            Else
                .Col = 11: .Text = Mid(ADORs!부모택번호, 4, 2) & "-" & Mid(ADORs!부모택번호, 6, 4)     '11
            End If
            
            .Col = 12: .Text = "0"                                                '12
            .Col = 13: .Text = ADORs!수선금액 & ""                                '13
            .Col = 14: .Text = ADORs!세트Key & ""                                 '14
            .Col = 15: .Text = ADORs!세트구분 & ""                                '15
            .Col = 16: .Text = ADORs!오점내용 & ""                                '16
            .Col = 17: .Text = ADORs!접수번호 & ""                                '17
            .Col = 18: .Text = ADORs!택번호 & ""                                  '18 (전체 택번호 보여주기)
                        
            If ADORs!판매취소 = "Y" Then
                .Row = .Row: .Row2 = .Row
                .Col = 1:    .Col2 = .MaxCols
                .BlockMode = True
                
                '.ForeColor = &HC0E0FF
                .ForeColor = vbRed
                
                .FontStrikethru = True
                .RowHeight(.Row) = 16
                
                .BlockMode = False
            End If
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        Dim 택번호 As String
        
        For i = 1 To frm출고.sprChul.MaxRows
            frm출고.sprChul.Row = i
            
            frm출고.sprChul.Col = 12
            If frm출고.sprChul.Text = "1" Then '확인 체크
                frm출고.sprChul.Col = 18: 택번호 = Trim(frm출고.sprChul.Text) & ""
                        
                Rtn = .SearchCol(18, -1, -1, 택번호, SearchFlagsNone)
                
                If Rtn > -1 Then
                    .Row = Rtn
                    .Col = 12: .Text = "1"
                    
                    .Row = Rtn
                    .Row2 = Rtn
                    .Col = 1
                    .Col2 = .MaxCols
                    .BlockMode = True
                    .BackColor = &HC0FFFF
                    .BlockMode = False
                End If
            End If
        Next i
        
        .ReDraw = True
    End With
End Sub

Private Sub 결제정보_Display()
    On Error GoTo ErrRtn

    '-------------------------------------------------------------------------------------------
    '
    '-------------------------------------------------------------------------------------------
    Query = "SELECT * FROM TB_신용카드승인"
    Query = Query & " WHERE 고객코드 = '" & pnlCode.Caption & "'"
    Query = Query & "   AND 접수번호 =  " & pnlNum.Caption
    Query = Query & "   AND SUBSTRING(메시지2,1,2) <> '취소'"
    Query = Query & " ORDER BY 승인일자, 승인시간 ASC"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprCard
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 2:  .Text = ADORs!승인번호 & ""
            .Col = 3:  .Text = ADORs!승인일자 & ""
            .Col = 4:  .Text = ADORs!승인시간 & ""
            .Col = 5:  .Text = ADORs!할부기간 & ""
            .Col = 6:  .Text = ADORs!결제금액 & ""
            .Col = 7:  .Text = ADORs!발급사코드 & "" '
            .Col = 8:  .Text = ADORs!카드종류명 & "" '
            .Col = 9:  .Text = ADORs!매입사코드 & "" '
            .Col = 10: .Text = ADORs!매입사명 & "" '
            .Col = 11: .Text = Left(ADORs!카드번호, 16) & "" '
            .Col = 12: .Text = ADORs!메시지1 & "" '
            .Col = 13: .Text = ADORs!메시지2 & "" '
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        .ReDraw = True
    End With
        
        
    '-------------------------------------------------------------------------------------------
    '
    '-------------------------------------------------------------------------------------------
    Query = "SELECT * FROM TB_현금영수증"
    Query = Query & " WHERE 고객코드 = '" & pnlCode.Caption & "'"
    Query = Query & "   AND 접수번호 =  " & pnlNum.Caption
    Query = Query & "   AND SUBSTRING(메시지2,1,2) <> '취소'"
    Query = Query & " ORDER BY 승인일자, 승인시간 ASC"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF Then
        With sprCash
            .Col = 1
                        
            For i = 1 To 11
                .Row = i: .Text = ""
            Next i
        End With
    Else
        With sprCash
            .Col = 1
            
            .Row = 1:  .Text = ADORs!승인번호 & ""   '
            .Row = 2:  .Text = ADORs!승인일자 & ""   '
            .Row = 3:  .Text = ADORs!승인시간 & ""   '
            .Row = 4:  .Text = ADORs!거래유형 & ""   ' 입력방법
            .Row = 5:  .Text = ADORs!총금액 & ""     '
            .Row = 6:  .Text = ADORs!사용자정보 & "" '
            .Row = 7:  .Text = ADORs!메시지1 & ""    '
            .Row = 8:  .Text = ADORs!메시지2 & ""    '
            .Row = 9:  .Text = ADORs!소득구분 & ""   '
            .Row = 10: .Text = ADORs!국세청1 & ""    '
            .Row = 11: .Text = ADORs!국세청2 & ""    '
        End With
    End If
    ADORs.Close
    Set ADORs = Nothing
    
    Exit Sub
    
ErrRtn:
    Screen.MousePointer = 0
    
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub 매출정보_Display()
    On Error GoTo ErrRtn

    '-------------------------------------------------------------------------------------------
    '
    '-------------------------------------------------------------------------------------------
    Query = "SELECT    매출일자"
    Query = Query & ", 접수번호"
    Query = Query & ", 일련번호"
    Query = Query & ", 적요"
    Query = Query & ", 접수수량"
    Query = Query & ", 접수금액"
    Query = Query & ", 현금입금"
    Query = Query & ", 카드입금"
    Query = Query & ", 사용마일리지"
    Query = Query & ", 쿠폰입금"
    Query = Query & ", 반품수량"
    Query = Query & ", (접수금액 - 입금합계 - 사용마일리지 - 쿠폰입금) AS 미수금액"
    Query = Query & " FROM TB_매출"
    Query = Query & " WHERE 고객코드 = '" & pnlCode.Caption & "'"
    Query = Query & "   AND 접수번호 =  " & pnlNum.Caption
    Query = Query & " ORDER BY 매출일자, 매출시간, 접수번호, 일련번호, 적요 ASC"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprAccount
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = Format(ADORs!매출일자, "YYYY-MM-DD") & ""           ' 1
            .Col = 2:  .Text = ADORs!접수번호 & ""                                 ' 2
            .Col = 3:  .Text = ADORs!일련번호 & ""                                 ' 3
            .Col = 4:  .Text = ADORs!적요 & ""                                     ' 4
            .Col = 5:  .Text = ADORs!접수수량 & ""                                 ' 5
            .Col = 6:  .Text = ADORs!접수금액 & ""                                 ' 6
            .Col = 7:  .Text = ADORs!현금입금 & "": .ForeColor = vbBlue            ' 7
            .Col = 8:  .Text = ADORs!카드입금 & "": .ForeColor = vbBlue            ' 8
            .Col = 9:  .Text = ADORs!사용마일리지 & ""                             ' 9
            .Col = 10: .Text = ADORs!쿠폰입금 & ""                                 '10
            .Col = 11: .Text = ADORs!미수금액 & ""                                 '11
            .Col = 12: .Text = ADORs!반품수량 & ""                                 '12
            
            If ADORs!반품수량 <> 0 Then
                .Row = .MaxRows: .Row2 = .MaxRows
                .Col = 6: .Col2 = .MaxCols
                .BlockMode = True
                .ForeColor = vbRed
                .BlockMode = False
            End If
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        If .MaxRows >= 1 Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Row = .MaxRows
            .Row2 = .MaxRows
            .Col = 1
            .Col2 = .MaxCols
            .BlockMode = True
            .BackColor = &HC0FFFF
            .BlockMode = False
                    
            .Col = 1:  .Text = "합계"
            .Col = 5:  .Formula = "SUM(E1:E" & .MaxRows - 1 & ")"
            .Col = 6:  .Formula = "SUM(F1:F" & .MaxRows - 1 & ")"
            .Col = 7:  .Formula = "SUM(G1:G" & .MaxRows - 1 & ")"
            .Col = 8:  .Formula = "SUM(H1:H" & .MaxRows - 1 & ")"
            .Col = 9:  .Formula = "SUM(I1:I" & .MaxRows - 1 & ")"
            .Col = 10: .Formula = "SUM(J1:J" & .MaxRows - 1 & ")"
            .Col = 11: .Formula = "SUM(K1:K" & .MaxRows - 1 & ")"
            .Col = 12: .Formula = "SUM(L1:L" & .MaxRows - 1 & ")"
        End If
        
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub btnAccount_Click(Index As Integer)
    Dim 접수일자  As String
    Dim 택번호    As String

    Dim strState  As String
    Dim strTemp   As String
    
    Dim 미수금액  As Long
    
    Dim 현금결제  As Long
    Dim 카드결제  As Long
    
    Dim 고객코드  As String
    Dim 의류명    As String
    
    Dim iPaper    As Integer
    Dim PrintCount As Integer
    
    On Error GoTo ErrRtn

    btnAccount(Index).Enabled = False

    If Index = 0 Then
        '완불
        If txtBalance2.Value = 0 Then
            '이미 완불처리됨
        Else
            txtIncome.Value = txtMisu.Value - txtCard.Value '입금액 = 잔액 - 카드결제금액
        End If
            
    Else
        If txtBalance2.Value = 0 Then
            MsgBox "완불을 하였습니다. 후불처리를 할수 없습니다.", vbInformation, "확인"
            
            btnAccount(Index).Enabled = True
            Exit Sub
        End If
        
        If txtCash.Value > 0 Or txtCard.Value > 0 Then
            MsgBox "부분 결제가 되었습니다. 후불처리를 할수 없습니다.", vbInformation, "확인"
            
            btnAccount(Index).Enabled = True
            Exit Sub
        End If
        
        '-----------------------------------------------------------------
        ' TB_입출고 - 판매취소후 남아 있는 접수내역을 '미불'로 수정
        '-----------------------------------------------------------------
        Query = "UPDATE TB_입출고 SET 결제여부 = '미불'"
        Query = Query & " WHERE 고객코드  = '" & pnlCode.Caption & "'"
        Query = Query & "   AND 접수번호  =  " & pnlNum.Caption
        Query = Query & "   AND 판매취소 <> 'Y'"
        ADOCon.Execute Query
    End If

    고객코드 = pnlCode.Caption & "" '
    미수금액 = txtMisu.Value        '
    현금결제 = txtCash.Value        '
    카드결제 = txtCard.Value        '
    
    If (현금결제 + 카드결제) > 미수금액 Then
       MsgBox "'판매취소후 잔액'보다 '수금액' 처리가 더 많습니다. 수금액을 확인하여 주십시요", vbInformation, "확인"
       
       btnAccount(Index).Enabled = True
       Exit Sub
    End If
            
    Call Set_고객미수금액(고객코드, txtBalance2.Value, "ADD")  ' 미수금액
    
    If (현금결제 > 0) Or (카드결제 > 0) Then
        '-----------------------------------------------------------
        ' TB_매출
        '-----------------------------------------------------------
        Dim iSEQ As Long
        
        Query = "SELECT ISNULL(MAX(일련번호),0) + 1"
        Query = Query & " FROM TB_매출"
        Query = Query & " WHERE 고객코드 = '" & 고객코드 & "'"
        Query = Query & "   AND 접수번호 =  " & pnlNum.Caption
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
                        
        iSEQ = ADORs(0)
        
        ADORs.Close
        Set ADORs = Nothing
        
        '-----------------------------------------------------------
        Query = "SELECT * FROM TB_매출"
        Query = Query & " WHERE 고객코드 = '" & 고객코드 & "'"
        Query = Query & "   AND 접수번호 =  " & pnlNum.Caption
        Query = Query & "   AND 일련번호 =  " & iSEQ
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, ADOCon, adOpenDynamic, adLockOptimistic
        
        If ADORs.EOF Then ADORs.AddNew
        
        ADORs!지사코드 = 가맹점정보.지사코드                    ' 1
        ADORs!가맹점코드 = 가맹점정보.가맹점코드                ' 2
        
        ADORs!고객코드 = 고객코드 & ""                          ' 3
        ADORs!접수번호 = pnlNum.Caption & ""                    ' 4
        ADORs!일련번호 = iSEQ                                   ' 5
        
        ADORs!매출일자 = Format(Date, "YYYY-MM-DD") & ""        ' 6
        ADORs!매출시간 = Format(Now, "hh:mm:ss")                ' 7
        ADORs!적요 = "[판매취소 입금]"                          ' 8
        ADORs!접수금액 = 0                                      ' 9
        ADORs!입금합계 = 현금결제 + 카드결제                    '10
        ADORs!현금입금 = 현금결제                               '11
        ADORs!카드입금 = 카드결제                               '12
        ADORs!쿠폰입금 = 0                                      '13
        ADORs!쿠폰번호 = ""                                     '14
        ADORs!사용마일리지 = 0                                  '15
        ADORs!세트할인 = 0                                      '16
        ADORs!에누리 = 0                                        '17
        ADORs!접수수량 = 0                                      '18
        ADORs!반품수량 = 0                                      '19
        ADORs!발생마일리지 = 0                                  '20
        ADORs!누적마일리지 = 0                                  '21
        ADORs!사용가능마일리지 = 0                              '22
        ADORs!이전미수금 = txtMisu.Value                        '23
        ADORs!본사전송여부 = ""                                 '24
        
        ADORs.Update
        
        ADORs.Close
        Set ADORs = Nothing
    End If
    
    '-------------------------------------------------------------------------------
    ' TB_신용카드승인
    '-------------------------------------------------------------------------------
    Dim 승인번호 As String
    Dim 승인일자 As String
    Dim 승인시간 As String
    
    With sprCard
        For i = 1 To .MaxRows
            .Row = i
            
            .Col = 2: 승인번호 = .Text & ""
            .Col = 3: 승인일자 = .Text & ""
            .Col = 4: 승인시간 = .Text & ""
            
            Query = "UPDATE TB_신용카드승인 SET 접수번호 =  " & pnlNum.Caption
            Query = Query & "                 , 고객코드 = '" & 고객코드 & "'"
            Query = Query & " WHERE 승인번호 = '" & 승인번호 & "'"
            Query = Query & "   AND 승인일자 = '" & 승인일자 & "'"
            Query = Query & "   AND 승인시간 = '" & 승인시간 & "'"
            ADOCon.Execute Query
         Next i
    End With
    
    '-------------------------------------------------------------------------------
    ' TB_현금영수증
    '-------------------------------------------------------------------------------
    With sprCash
        .Col = 1
        
        .Row = 1
        If .Text <> "" Then
            .Row = 1: 승인번호 = .Text & ""
            .Row = 2: 승인일자 = .Text & ""
            .Row = 3: 승인시간 = .Text & ""
            
            Query = "UPDATE TB_현금영수증 SET 접수번호 =  " & pnlNum.Caption
            Query = Query & "               , 고객코드 = '" & 고객코드 & "'"
            Query = Query & " WHERE 승인번호 = '" & 승인번호 & "'"
            Query = Query & "   AND 승인일자 = '" & 승인일자 & "'"
            Query = Query & "   AND 승인시간 = '" & 승인시간 & "'"
            ADOCon.Execute Query
        End If
    End With
    
    '-------------------------------------------------------------------------------
    '
    '-------------------------------------------------------------------------------
    Dim CommPort As String
    Dim BaudRate As String
    
    CommPort = GetIniStr("VAN", "KS7500_CommPort", "", iniFile)
    BaudRate = GetIniStr("VAN", "KS7500_BaudRate", "", iniFile)
    
'    Do
'        Rtn = KS7500i.CheckPort(CInt(CommPort), CLng(BaudRate))
'        DoEvents
'
'        If Rtn < 0 Then
'            i = i + 1
'
'            If i > 3 Then
'                MsgBox "카드단말기 장치가 연결되어 있지 않습니다", vbCritical, "오류"
'
'                Exit Do
'            End If
'        Else
'            Call KS7500i.SetConfig("", Rtn, CLng(BaudRate))    '첫번째 인자는 "" 로 넣어 준다.
'
'            KS7500i.InitPrint
'
'            '"KS4060 보안인증" 는 해당 단말기 에서 바로 출력 처리를 한다.
'            If 가맹점정보.CAT단말기종류 <> "KS4060 보안인증" Then
'
'                For PrintCount = 1 To 2
'                    Call 카드결제_Report(KS7500i, frm판매취소.sprCard, PrintCount)
'                    Call 현금영수증_Report(KS7500i, frm판매취소.sprCash, PrintCount)
'                Next PrintCount
'
'                Call 승인취소_Report(KS7500i, pnlCode.Caption, CLng(pnlNum.Caption))    '신용카드승인 취소, 현금영수증승인 취소 내역 출력
'            End If
'
'            KS7500i.ClosePort
'            DoEvents
'
'            Delay 1 ' * 카드결제 영수증 후 -> frm출고.cmdCancel_Click 의 "승인취소_Report()" 할때 KS7500i 와 충돌 발생을 막기 위해...
'        End If
'    Loop Until Rtn > 0
        
    Unload frm판매취소
    
    Exit Sub

ErrRtn:
    btnAccount(Index).Enabled = True

    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub btnCard_Click()
    If txtMisu.Value = 0 Then
        MsgBox "미수금액이 없습니다.", vbInformation, "확인"
        
        Exit Sub
    End If
    
    If Check_KS7500 = False Then
        MsgBox "환경설정에서 사업자번호, 단말기번호 등이 올바르게 입력되었는지 확인하십시요.", vbInformation, "확인"
        
        Exit Sub
    End If
    
    
    Account_Form = "판매취소2"
    
    frmKSNET2.pnlCustomCode.Caption = pnlCode.Caption  '
    frmKSNET2.pnlNum.Caption = pnlNum.Caption          ' 접수번호
    frmKSNET2.txtMoney.Value = txtBalance2.Value       '
    frmKSNET2.txtMoney.Tag = txtBalance2.Value         '
    
    Call frmKSNET2.신용카드승인요청_Rtn("1")
    
    frmKSNET2.Show 1
End Sub

Private Sub btnCash_Click()
    If Check_KS7500 = False Then
        MsgBox "환경설정에서 사업자번호, 단말기번호 등이 올바르게 입력되었는지 확인하십시요.", vbInformation, "확인"
        
        Exit Sub
    End If
    Unload frmKSNETCash
    
    Account_Form = "판매취소2"
    
    
    frmKSNETCash.pnlCustomCode.Caption = pnlCode.Caption  '고객코드
    frmKSNETCash.pnlNum.Caption = pnlCode.Caption         ' 접수번호
    frmKSNETCash.txtMoney.Value = txtCash.Value           '금액
    
    Call frmKSNETCash.현금영수증승인요청_Rtn("3")
    
    frmKSNETCash.Show 1
End Sub

Private Sub btnCashCancel_Click()
    sprCash.Row = 1
    sprCash.Col = 1
    
    If sprCash.Text = "" Then Exit Sub
    Unload frmKSNETCash
    Account_Form = "판매취소2"
    
    With frmKSNETCash.sprGrid
        .Col = 1
    
        .Row = 1:  .Text = Spread_GetData(sprCash, 1, 1, True)   '승인번호
        .Row = 2:  .Text = Spread_GetData(sprCash, 2, 1, True)   '승인일자
        .Row = 3:  .Text = Spread_GetData(sprCash, 3, 1, True)   '승인시간
        
        .Row = 4:  .Text = Spread_GetData(sprCash, 4, 1, True)   '개인(0), (1)
        
        .Row = 5:  .Text = Spread_GetData(sprCash, 5, 1, True)   '총금액
        .Row = 6:  .Text = Spread_GetData(sprCash, 6, 1, True)   '사용자정보
        .Row = 7:  .Text = Spread_GetData(sprCash, 7, 1, True)   '메시지1
        .Row = 8:  .Text = Spread_GetData(sprCash, 8, 1, True)   '메시지2
        .Row = 9:  .Text = Spread_GetData(sprCash, 9, 1, True)   '소득구분
        .Row = 10: .Text = Spread_GetData(sprCash, 10, 1, True)  '국세청1
        .Row = 11: .Text = Spread_GetData(sprCash, 11, 1, True)  '국세청2
    End With

    frmKSNETCash.pnlCustomCode.Caption = pnlCode.Caption                        '고객코드
    frmKSNETCash.pnlNum.Caption = pnlNum.Caption                                '접수번호
    frmKSNETCash.txtMoney.Value = Spread_GetData(sprCash, 5, 1, True)           '

    frmKSNETCash.pnlApprovalNo.Caption = Spread_GetData(sprCash, 1, 1, True)   '승인번호
    frmKSNETCash.pnlApprovalDay.Caption = Spread_GetData(sprCash, 2, 1, True)  '승인일자
    frmKSNETCash.pnlApprovalTime.Caption = Spread_GetData(sprCash, 3, 1, True) '승인시간

    Call frmKSNETCash.현금영수증승인요청_Rtn("4")
    
    frmKSNETCash.Show 1
End Sub

Private Sub sprCard_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    If Row <= 0 Then Exit Sub
    
    Account_Form = "판매취소2"
    
    With frmKSNET2.sprGrid
        .Col = 1
    
        .Row = 1:  .Text = Spread_GetData(sprCard, Row, 2, True)   '승인번호
        .Row = 2:  .Text = Spread_GetData(sprCard, Row, 3, True)   '승인일자
        .Row = 3:  .Text = Spread_GetData(sprCard, Row, 4, True)   '승인시간
        
        .Row = 4:  .Text = Spread_GetData(sprCard, Row, 5, True)   '할부기간
        .Row = 5:  .Text = Spread_GetData(sprCard, Row, 6, True)   '결제금액
                        
        .Row = 6:  .Text = Spread_GetData(sprCard, Row, 7, True)   '발급사코드
        .Row = 7:  .Text = Spread_GetData(sprCard, Row, 8, True)   '발급사명
        .Row = 8:  .Text = Spread_GetData(sprCard, Row, 9, True)   '매입사코드
        .Row = 9:  .Text = Spread_GetData(sprCard, Row, 10, True)  '매입사명
        .Row = 10: .Text = Spread_GetData(sprCard, Row, 11, True)  '카드번호
        .Row = 11: .Text = Spread_GetData(sprCard, Row, 12, True)  '메시지1
        .Row = 12: .Text = Spread_GetData(sprCard, Row, 13, True)  '메시지2
    End With

    frmKSNET2.pnlCustomCode.Caption = pnlCode.Caption                '고객코드
    frmKSNET2.pnlNum.Caption = pnlNum.Caption                        '접수번호

    frmKSNET2.txtMoney.ReadOnly = True
    frmKSNET2.txtMoney.Value = Spread_GetData(sprCard, Row, 6, True) '

    frmKSNET2.pnlApprovalNo.Caption = Spread_GetData(sprCard, Row, 2, True)   '승인번호
    frmKSNET2.pnlApprovalDay.Caption = Spread_GetData(sprCard, Row, 3, True)  '승인일자
    frmKSNET2.pnlApprovalTime.Caption = Spread_GetData(sprCard, Row, 4, True) '승인시간

    Call frmKSNET2.신용카드승인요청_Rtn("2")
    
    frmKSNET2.Show 1
End Sub

Private Sub Form_Activate()
    If 판매취소_Flag = True Then Exit Sub '2번실행 방지
    
    
    
    Call 매출정보_Display
    Call 결제정보_Display
    
    판매취소_Flag = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    판매취소_Flag = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    
    Me.Top = ((frmMain.Height - Me.Height) / 2) + frmMain.Top
    Me.Left = ((frmMain.Width - Me.Width) / 2) + frmMain.Left

    With sprGrid
        .MaxRows = 0
        .RowHeight(-1) = 16
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeNormal
    End With
    
    With sprCard
        .MaxRows = 0
        .RowHeight(-1) = 18
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeNormal
    End With
    
    With sprCash
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


    

'    Dim iPaper As String
'
'    iPaper = GetIniStr("Printer", "Paper2", "", iniFile)
'
'    If iPaper = "2" Then
'        optReceipt(2).Value = True
'    ElseIf iPaper = "1" Then
'        optReceipt(1).Value = True
'    Else
'        optReceipt(0).Value = True
'    End If

End Sub

Private Sub txtBalance2_Change()
    If txtBalance2.Value = 0 Then
        btnAccount(0).Enabled = True
        btnAccount(1).Enabled = False
    Else
        btnAccount(0).Enabled = False
        btnAccount(1).Enabled = True
    End If
End Sub

Private Sub txtCard_Change()
    txtBalance2.Value = txtMisu.Value - txtCash.Value - txtCard.Value
End Sub

Private Sub txtIncome_Change()
    Dim 미수금액 As Long
    
    미수금액 = txtMisu.Value
    
    If txtIncome.Value >= (미수금액 - txtCard.Value) Then
        txtBalance2.Value = 0                              '결제후 잔액
        txtCash.Value = 미수금액 - txtCard.Value   '현금결제
        txtChange.Value = txtIncome.Value - txtCash.Value  '거스름돈
    Else
        txtChange.Value = 0                                                    '거스름돈
        txtCash.Value = txtIncome.Value                                        '현금결제
        txtBalance2.Value = 미수금액 - (txtCash.Value + txtCard.Value) '결제후 잔액
    End If
End Sub

Private Sub txtMisu_Change()
    txtBalance2.Value = txtMisu.Value - txtCash.Value - txtCard.Value
End Sub
