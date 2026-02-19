VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_03001 
   Caption         =   "일일출고 현황"
   ClientHeight    =   9675
   ClientLeft      =   495
   ClientTop       =   5160
   ClientWidth     =   17190
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_03001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9675
   ScaleWidth      =   17190
   WindowState     =   2  '최대화
   Begin Threed.SSPanel pnlProg 
      Height          =   1215
      Left            =   5670
      TabIndex        =   37
      Top             =   1710
      Visible         =   0   'False
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   2143
      _Version        =   262144
      BackColor       =   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "P_03001.frx":058A
      Caption         =   " "
      BorderWidth     =   0
      BevelOuter      =   1
      BevelInner      =   2
      Alignment       =   4
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   9675
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17190
      _ExtentX        =   30321
      _ExtentY        =   17066
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_03001.frx":3555
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   1
         Top             =   540
         Width           =   17160
         _ExtentX        =   30268
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboPage 
            Height          =   315
            Left            =   12315
            Style           =   2  '드롭다운 목록
            TabIndex        =   39
            Top             =   420
            Width           =   1080
         End
         Begin VB.ComboBox cboNum 
            Height          =   315
            ItemData        =   "P_03001.frx":3647
            Left            =   5610
            List            =   "P_03001.frx":3649
            Style           =   2  '드롭다운 목록
            TabIndex        =   33
            Top             =   420
            Width           =   1335
         End
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   975
            Locked          =   -1  'True
            TabIndex        =   13
            Text            =   "cboOffice"
            Top             =   60
            Width           =   2805
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Left            =   975
            TabIndex        =   2
            Top             =   420
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   556
            _Version        =   393216
            Format          =   59047936
            CurrentDate     =   36686
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "인쇄매수:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   5
            Left            =   11190
            TabIndex        =   38
            Top             =   480
            Width           =   1080
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "출고회차:"
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
            Index           =   4
            Left            =   4680
            TabIndex        =   36
            Top             =   480
            Width           =   885
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "출고일자:"
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
            Index           =   3
            Left            =   45
            TabIndex        =   35
            Top             =   480
            Width           =   885
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "지 사 명:"
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
            TabIndex        =   34
            Top             =   120
            Width           =   885
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   3
         Top             =   15
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   900
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   255
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
         Caption         =   " 일일출고 현황 (P_03001)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_03001.frx":364B
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   1
         Left            =   9585
         TabIndex        =   4
         Top             =   15
         Width           =   7590
         _ExtentX        =   13388
         _ExtentY        =   900
         _Version        =   262144
         Font3D          =   3
         ForeColor       =   192
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureBackgroundStyle=   2
         PictureBackground=   "P_03001.frx":384D
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   5
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "종료"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_03001.frx":3A4F
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   6
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "엑셀"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_03001.frx":3FE9
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   7
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "인쇄"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_03001.frx":4583
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   8
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "취소"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_03001.frx":4B1D
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   9
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "삭제"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_03001.frx":50B7
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   10
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "저장"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_03001.frx":5651
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   11
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "신규"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_03001.frx":5BEB
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   12
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "조회"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_03001.frx":6185
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   7875
         Left            =   15
         TabIndex        =   14
         Top             =   1335
         Width           =   5610
         _Version        =   524288
         _ExtentX        =   9895
         _ExtentY        =   13891
         _StockProps     =   64
         BackColorStyle  =   1
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
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
         MaxCols         =   5
         MaxRows         =   34
         ScrollBars      =   2
         SpreadDesigner  =   "P_03001.frx":671F
         UserResize      =   1
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdView1 
         Height          =   7515
         Left            =   5640
         TabIndex        =   15
         Top             =   1335
         Width           =   11535
         _Version        =   524288
         _ExtentX        =   20346
         _ExtentY        =   13256
         _StockProps     =   64
         BackColorStyle  =   1
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
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
         MaxCols         =   23
         Protect         =   0   'False
         SpreadDesigner  =   "P_03001.frx":6D93
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   795
         Index           =   2
         Left            =   5640
         TabIndex        =   16
         Top             =   8865
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   1402
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSPanel SSPanel 
            Height          =   360
            Index           =   3
            Left            =   45
            TabIndex        =   23
            Top             =   45
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   635
            _Version        =   262144
            BackColor       =   16777215
            Caption         =   "의류"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   360
            Index           =   1
            Left            =   1065
            TabIndex        =   17
            Top             =   45
            Width           =   945
            _Version        =   262145
            _ExtentX        =   1667
            _ExtentY        =   635
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   192
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   2
            StartText.y     =   2
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   20
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   1
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   360
            Index           =   2
            Left            =   3060
            TabIndex        =   18
            Top             =   45
            Width           =   945
            _Version        =   262145
            _ExtentX        =   1667
            _ExtentY        =   635
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   192
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   2
            StartText.y     =   2
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   20
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   1
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   360
            Index           =   3
            Left            =   5055
            TabIndex        =   19
            Top             =   45
            Width           =   945
            _Version        =   262145
            _ExtentX        =   1667
            _ExtentY        =   635
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   192
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   2
            StartText.y     =   2
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   20
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   1
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   360
            Index           =   4
            Left            =   1065
            TabIndex        =   20
            Top             =   390
            Width           =   945
            _Version        =   262145
            _ExtentX        =   1667
            _ExtentY        =   635
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   192
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   2
            StartText.y     =   2
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   20
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   1
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   360
            Index           =   5
            Left            =   3060
            TabIndex        =   21
            Top             =   390
            Width           =   945
            _Version        =   262145
            _ExtentX        =   1667
            _ExtentY        =   635
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   192
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   2
            StartText.y     =   2
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   20
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   1
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   360
            Index           =   6
            Left            =   5055
            TabIndex        =   22
            Top             =   390
            Width           =   945
            _Version        =   262145
            _ExtentX        =   1667
            _ExtentY        =   635
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   192
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   2
            StartText.y     =   2
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   20
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   1
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   360
            Index           =   4
            Left            =   45
            TabIndex        =   24
            Top             =   390
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   635
            _Version        =   262144
            BackColor       =   16777215
            Caption         =   "정상"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   360
            Index           =   5
            Left            =   2040
            TabIndex        =   25
            Top             =   45
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   635
            _Version        =   262144
            BackColor       =   16777215
            Caption         =   "소품"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   360
            Index           =   6
            Left            =   4035
            TabIndex        =   26
            Top             =   45
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   635
            _Version        =   262144
            BackColor       =   16777215
            Caption         =   "기타"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   360
            Index           =   7
            Left            =   2040
            TabIndex        =   27
            Top             =   390
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   635
            _Version        =   262144
            BackColor       =   16777215
            Caption         =   "반품"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   360
            Index           =   8
            Left            =   4035
            TabIndex        =   28
            Top             =   390
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   635
            _Version        =   262144
            BackColor       =   16777215
            Caption         =   "확인"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   360
            Index           =   7
            Left            =   7050
            TabIndex        =   29
            Top             =   390
            Width           =   945
            _Version        =   262145
            _ExtentX        =   1667
            _ExtentY        =   635
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   192
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   2
            StartText.y     =   2
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   20
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   1
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   360
            Index           =   9
            Left            =   6030
            TabIndex        =   30
            Top             =   390
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   635
            _Version        =   262144
            BackColor       =   16777215
            Caption         =   "품명"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   360
            Index           =   8
            Left            =   9045
            TabIndex        =   31
            Top             =   390
            Width           =   945
            _Version        =   262145
            _ExtentX        =   1667
            _ExtentY        =   635
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   192
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   2
            StartText.y     =   2
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   20
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   1
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   360
            Index           =   10
            Left            =   8025
            TabIndex        =   32
            Top             =   390
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   635
            _Version        =   262144
            BackColor       =   16777215
            Caption         =   "재세탁"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   435
         Index           =   0
         Left            =   15
         TabIndex        =   40
         Top             =   9225
         Width           =   5610
         _ExtentX        =   9895
         _ExtentY        =   767
         _Version        =   262144
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   345
            Index           =   0
            Left            =   3900
            TabIndex        =   41
            Top             =   45
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
            _ExtentY        =   609
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   192
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.76
               Charset         =   0
               Weight          =   700
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
            StartText.y     =   3
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
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "출고수량:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   2775
            TabIndex        =   43
            Top             =   135
            Width           =   1080
         End
         Begin VB.Label Label 
            BackStyle       =   0  '투명
            Caption         =   "점"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   4860
            TabIndex        =   42
            Top             =   135
            Width           =   255
         End
      End
   End
End
Attribute VB_Name = "P_03001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cboNum_Click()
    Call Data_Display
End Sub

Private Sub cboOffice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SearchString_한글 KeyAscii
'    Else
       ' SearchString KeyAscii
    End If

End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: Call Data_Display   ' 조회
        Case 1: 'Call DataAdd        ' 신규
        Case 2: 'Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
        Case 4: 'Call DataCancel     ' 취소
        Case 5:

            cmdBtn(Index).Enabled = False

            Dim 가맹점코드 As String
            Dim 가맹점명   As String
            Dim 택코드     As String
            Dim iRow       As Integer
            
            For iRow = 1 To spdView.MaxRows
                spdView.Row = iRow
                spdView.Col = 5
                
                If spdView.Text = "1" Then
                    Call spdView.SetSelection(1, iRow, spdView.MaxCols, iRow)
                    DoEvents
                    
                    spdView.Col = 1: 가맹점코드 = spdView.Text & ""
                    spdView.Col = 2: 가맹점명 = spdView.Text & ""
                    spdView.Col = 3: 택코드 = spdView.Text & ""
                    
                    Call Data_Display2(가맹점코드)
                    DoEvents
                    
                    Debug.Print "DataPrint(가맹점명, 택코드)   " & Now
                    Call DataPrint(가맹점명, 택코드)     ' 인쇄
                    DoEvents
                End If
            Next iRow
            
            cmdBtn(Index).Enabled = True
            
        Case 6: Call Export_Excel(P_00000.cdgExcel, spdView1)      ' 엑셀
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

Private Sub dtInput_Change()
    Call Data_Display
End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(5).Enabled = True
    cmdBtn(6).Enabled = True
    
    '본사 여부
    If HeadOffice = MASTER_OFFICE_CODE Then
        cboOffice.Locked = False
    Else
        cboOffice.Locked = True
    End If
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    cboPage.ListIndex = Val(GetIniStr("DEF SETTING", "P_03001_01", "1", m_iniFile))
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    Dim i As Integer
    
    With spdView
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
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
    
    With spdView1
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
    
    '아래 소스 순서를 바꾸지 말것...
    
    Call Get_지사리스트(cboOffice)
     
    With cboOffice
        For i = 0 To .ListCount - 1
            If Mid(.List(i), 2, 4) = HeadOffice Then
                .ListIndex = i
                
                Exit For
            End If
        Next i
    End With
    
    dtInput.Value = Date
    
    With cboNum
        .Clear
        
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        .AddItem "6"
        .AddItem "7"
        .AddItem "8"
        .AddItem "9"
        .AddItem "10"
        .AddItem "11"
        .AddItem "12"
        .AddItem "13"
        .AddItem "14"
        .AddItem "15"
        .AddItem "16"
        .AddItem "17"
        .AddItem "18"
        .AddItem "19"
        .AddItem "20"
        
        .ListIndex = 0
    End With
    
    With cboPage
        .Clear
        
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        .ListIndex = 1
    End With
    
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i      As Integer
    Dim iTotal As Long

    spdView1.MaxRows = 0
    
    For i = 1 To 8
        txtNum(i).Value = 0
    Next i
    
    ReDim sValue(2)
    
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    sValue(1) = Format(dtInput.Value, "YYYY-MM-DD")
    sValue(2) = cboNum.Text
    
    If sValue(0) = "" Then Exit Sub
        
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_03001_00", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_03001_00", sValue(), Err_Num, Err_Dec)
    End If
            
    With spdView
        .MaxRows = 0
        
        Do While Not RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = RS01!가맹점코드 & ""
            .Col = 2: .Text = RS01!가맹점명 & ""
            .Col = 3: .Text = RS01!택코드 & ""
            .Col = 4: .Text = RS01!출고수량 & ""
            .Col = 5: .Text = "1"
            
            iTotal = iTotal + RS01!출고수량
            
            RS01.MoveNext
        Loop
    End With
    
    txtNum(0).Value = iTotal
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Data_Display2(가맹점코드 As String)
    Dim i    As Integer
    Dim ipos As Integer
    
    On Error GoTo ErrRtn
    
    pnlProg.Visible = True
    DoEvents
    
    For i = 1 To 8
        txtNum(i).Value = 0
    Next i
    
    '----------------------------------------------------------------
    ' SP_03001_01
    '----------------------------------------------------------------
    ReDim sValue(2)
    
    sValue(0) = 가맹점코드
    sValue(1) = Format(dtInput.Value, "YYYY-MM-DD")
    sValue(2) = cboNum.Text
    
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_03001_01", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_03001_01", sValue(), Err_Num, Err_Dec)
    End If
        
    With spdView1
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = Format(RS01!택번호, "000-00-0000") & "" ' 1
            .Col = 2:  .Text = RS01!물품구분 & ""       ' 2
            .Col = 3:  .Text = RS01!물품상태 & ""       ' 2
            .Col = 4:  .Text = RS01!성명 & ""           ' 3
            .Col = 5:  .Text = RS01!전화번호 & ""       ' 4
            .Col = 6:  .Text = RS01!휴대전화 & ""       ' 5
            .Col = 7:  .Text = RS01!의류코드 & ""       ' 6
            .Col = 8:  .Text = RS01!의류명 & ""         ' 7
            .Col = 9:  .Text = RS01!색상 & ""           ' 8
            .Col = 10: .Text = RS01!무늬 & ""           ' 9
            .Col = 11: .Text = RS01!내용 & ""           '10
            .Col = 12: .Text = RS01!상표 & ""           '11
            .Col = 13: .Text = RS01!금액 & ""           '12
            .Col = 14: .Text = RS01!접수일자 & ""       '13
            .Col = 15: .Text = RS01!가맹점출고일자 & "" '14
            .Col = 16: .Text = RS01!가맹점입고일자 & "" '15
            .Col = 17: .Text = RS01!출고일자 & ""     '16
            .Col = 18: .Text = RS01!부모택번호 & ""     '16
            .Col = 19: .Text = RS01!반품환불일자 & ""   '17
            .Col = 20: .Text = RS01!세탁환불일자 & ""   '18
            .Col = 21: .Text = RS01!판매취소일자 & ""   '19
            .Col = 22: .Text = RS01!환불사유 & ""       '20
            .Col = 23: .Text = RS01!오점내용 & ""       '21
            
            If RS01!물품구분 = "의류" Then txtNum(1).Value = txtNum(1).Value + 1
            If RS01!물품구분 = "소품" Then txtNum(2).Value = txtNum(2).Value + 1
            If RS01!물품구분 = "기타" Then txtNum(3).Value = txtNum(3).Value + 1
                        
            
            If RS01!물품상태 = "정상" Then txtNum(4).Value = txtNum(4).Value + 1
            If RS01!물품상태 = "반품" Then txtNum(5).Value = txtNum(5).Value + 1
            If RS01!물품상태 = "확인" Then txtNum(6).Value = txtNum(6).Value + 1
            If RS01!물품상태 = "품명" Then txtNum(7).Value = txtNum(7).Value + 1
            
            ipos = InStr(RS01!내용 & "", "재")
            
            If ipos > 0 Then txtNum(8).Value = txtNum(8).Value + 1
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
    
    pnlProg.Visible = False
    
    Exit Sub
    
ErrRtn:
    pnlProg.Visible = False
    
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub DataPrint(가맹점명 As String, 택코드 As String)
    On Error GoTo ErrRtn
    
    Dim 택번호      As String
    Dim XML         As String
    Dim i           As Integer
    Dim Idx         As Integer
    Dim FileNumber
        
    If spdView1.MaxRows = 0 Then Exit Sub
    
    FileNumber = FreeFile
    
    Open App.Path & "\XML\지사출고현황.XML" For Output As #FileNumber
    
    Print #FileNumber, "<?xml version=""1.0"" encoding=""EUC-KR""?>"
    Print #FileNumber, "<root>"
    
          XML = "    <조건>"
    XML = XML & "        <가맹점>" & Func_Replace(가맹점명) & " (" & 택코드 & ")  출고내역</가맹점>"
    XML = XML & "        <출고수량>출고수량 : " & spdView1.MaxRows & " 점</출고수량>"
    XML = XML & "        <출고일자>출고일자 : " & Format(dtInput.Value, "YYYY년 MM월 DD일") & "  (출고회차 : " & cboNum.Text & ")</출고일자>"
    XML = XML & "        <의류>의류 : " & txtNum(1).Text & " 점</의류>"
    XML = XML & "        <소품>소품 : " & txtNum(2).Text & " 점</소품>"
    XML = XML & "        <기타>기타 : " & txtNum(3).Text & " 점</기타>"
    XML = XML & "        <정상>정상 : " & txtNum(4).Text & " 점</정상>"
    XML = XML & "        <반품>반품 : " & txtNum(5).Text & " 점</반품>"
    XML = XML & "        <확인>확인 : " & txtNum(6).Text & " 점</확인>"
    XML = XML & "        <품명>품명 : " & txtNum(7).Text & " 점</품명>"
    XML = XML & "        <재세탁>재세탁 : " & txtNum(8).Text & " 점</재세탁>"
    XML = XML & "   </조건>"
    Print #FileNumber, XML
    
    With spdView1
        Idx = 0
        
        For i = 1 To .MaxRows
            .Row = i
            
            If Idx = 0 Or Idx = 6 Then
                If Idx = 0 Then
                    XML = "    <Data>"
                Else
                    XML = XML & "   </Data>"
                    Print #FileNumber, XML
                    
                    XML = "    <Data>"
                End If
                
                Idx = 0
            End If
            
            Idx = Idx + 1
            
            .Col = 2
            If Trim(.Text) = "" Or Trim(.Text) = "의류" Then
                XML = XML & "        <구분" & Idx & "></구분" & Idx & ">"
            Else
                XML = XML & "        <구분" & Idx & ">" & Left(.Text, 1) & "</구분" & Idx & ">"
            End If
            
            .Col = 3
            If (Trim(.Text) = "") Or (Trim(.Text) = "정상") Or (Trim(.Text) = "품명") Then
                XML = XML & "        <상태" & Idx & "></상태" & Idx & ">"
            Else
                XML = XML & "        <상태" & Idx & ">" & Left(.Text, 1) & "</상태" & Idx & ">"
            End If
            
            .Col = 1: XML = XML & "        <택번호" & Idx & ">" & Right(.Text, 7) & "</택번호" & Idx & ">"
        Next i
        
        If Idx = 6 Then
            XML = XML & "   </Data>"
            Print #FileNumber, XML
        Else
            For i = Idx + 1 To 6
                XML = XML & "        <구분" & i & "></구분" & i & ">"
                XML = XML & "        <상태" & i & "></상태" & i & ">"
                XML = XML & "        <택번호" & i & "></택번호" & i & ">"
            Next i
            
            XML = XML & "   </Data>"
            Print #FileNumber, XML
        End If
        
        Print #FileNumber, "</root>"
        Close #FileNumber
    End With
    
    With rpt지사출고현황
        .dc.FileURL = App.Path & "\XML\지사출고현황.XML"
        .Printer.Copies = cboPage.Text '인쇄매수
        .PrintReport False
        
        '.Show 1
    End With

    Unload rpt지사출고현황
    
    Exit Sub

ErrRtn:
    MsgBox Err.Description, vbInformation, "오류"
    Screen.MousePointer = 0
End Sub

Private Sub DataScreen()
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SetIniStr("DEF SETTING", "P_03001_01", cboPage.ListIndex, m_iniFile)
End Sub

Private Sub spdView_Click(ByVal Col As Long, ByVal Row As Long)
    If Row <= 0 Then Exit Sub
            
    If Col = 5 Then Exit Sub '출력여부 체크박스인 경우...
    
    Dim 가맹점코드 As String
    
    spdView.Row = Row
    spdView.Col = 1: 가맹점코드 = spdView.Text & ""
    
    Call Data_Display2(가맹점코드)
End Sub

Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    'Call spdView_Click(NewCol, NewRow)
End Sub
