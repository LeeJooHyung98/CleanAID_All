VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_01004_A 
   Caption         =   "가맹점별 품목할인 등록"
   ClientHeight    =   12165
   ClientLeft      =   1365
   ClientTop       =   3225
   ClientWidth     =   15885
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_01004_A.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12165
   ScaleWidth      =   15885
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   12165
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15885
      _ExtentX        =   28019
      _ExtentY        =   21458
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_01004_A.frx":058A
      Begin XtremeSuiteControls.TabControl TabControl1 
         Height          =   10815
         Left            =   9210
         TabIndex        =   8
         Top             =   1335
         Width           =   6660
         _Version        =   851970
         _ExtentX        =   11748
         _ExtentY        =   19076
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
         Color           =   64
         PaintManager.BoldSelected=   -1  'True
         PaintManager.HotTracking=   -1  'True
         PaintManager.ButtonMargin=   "2,3,2,3"
         ItemCount       =   2
         SelectedItem    =   1
         Item(0).Caption =   "할인 품목"
         Item(0).ControlCount=   3
         Item(0).Control(0)=   "TabControlPage(0)"
         Item(0).Control(1)=   "txtFind"
         Item(0).Control(2)=   "cmdBtn(10)"
         Item(1).Caption =   "할인 일정"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "TabControlPage(1)"
         Begin VB.TextBox txtFind 
            Height          =   345
            Left            =   -66580
            TabIndex        =   40
            Top             =   30
            Visible         =   0   'False
            Width           =   3105
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage 
            Height          =   10365
            Index           =   1
            Left            =   30
            TabIndex        =   10
            Top             =   420
            Width           =   6600
            _Version        =   851970
            _ExtentX        =   11642
            _ExtentY        =   18283
            _StockProps     =   1
            Page            =   3
            Begin SSSplitter.SSSplitter SSSplitter2 
               Height          =   10365
               Left            =   0
               TabIndex        =   44
               Top             =   0
               Width           =   6600
               _ExtentX        =   11642
               _ExtentY        =   18283
               _Version        =   262144
               AutoSize        =   1
               SplitterBarWidth=   1
               SplitterBarAppearance=   1
               BorderStyle     =   1
               PaneTree        =   "P_01004_A.frx":069C
               Begin FPSpreadADO.fpSpread sprSchedule 
                  Height          =   9330
                  Left            =   15
                  TabIndex        =   45
                  Top             =   1020
                  Width           =   6390
                  _Version        =   524288
                  _ExtentX        =   11271
                  _ExtentY        =   16457
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
                  MaxCols         =   5
                  ScrollBars      =   2
                  SpreadDesigner  =   "P_01004_A.frx":070E
                  Appearance      =   1
                  HighlightHeaders=   1
                  HighlightStyle  =   1
               End
               Begin FPSpreadADO.fpSpread spdList 
                  Height          =   9330
                  Left            =   6420
                  TabIndex        =   46
                  Top             =   1020
                  Width           =   165
                  _Version        =   524288
                  _ExtentX        =   291
                  _ExtentY        =   16457
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
                  MaxCols         =   3
                  Protect         =   0   'False
                  ScrollBars      =   2
                  SpreadDesigner  =   "P_01004_A.frx":0D4E
                  HighlightHeaders=   1
                  HighlightStyle  =   1
               End
               Begin Threed.SSPanel SSPanel1 
                  Height          =   990
                  Left            =   15
                  TabIndex        =   47
                  Top             =   15
                  Width           =   6570
                  _ExtentX        =   11589
                  _ExtentY        =   1746
                  _Version        =   262144
                  RoundedCorners  =   0   'False
                  FloodShowPct    =   -1  'True
                  Begin VB.ComboBox cboInput 
                     Height          =   315
                     Index           =   1
                     Left            =   1110
                     Style           =   2  '드롭다운 목록
                     TabIndex        =   49
                     Top             =   510
                     Width           =   3015
                  End
                  Begin VB.ComboBox cboInput 
                     Height          =   315
                     Index           =   0
                     Left            =   1110
                     TabIndex        =   48
                     Text            =   "cboInput"
                     Top             =   150
                     Width           =   3015
                  End
                  Begin VB.Label lblProgress 
                     Alignment       =   1  '오른쪽 맞춤
                     BackStyle       =   0  '투명
                     Caption         =   "지사명:"
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
                     Left            =   180
                     TabIndex        =   51
                     Top             =   210
                     Width           =   900
                  End
                  Begin VB.Label lblProgress 
                     Alignment       =   1  '오른쪽 맞춤
                     BackStyle       =   0  '투명
                     Caption         =   "가맹점명:"
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
                     Left            =   180
                     TabIndex        =   50
                     Top             =   585
                     Width           =   900
                  End
               End
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage 
            Height          =   10365
            Index           =   0
            Left            =   -69970
            TabIndex        =   9
            Top             =   420
            Visible         =   0   'False
            Width           =   6600
            _Version        =   851970
            _ExtentX        =   11642
            _ExtentY        =   18283
            _StockProps     =   1
            Page            =   0
            Begin SSSplitter.SSSplitter SSSplitter1 
               Height          =   10365
               Left            =   0
               TabIndex        =   42
               Top             =   0
               Width           =   6600
               _ExtentX        =   11642
               _ExtentY        =   18283
               _Version        =   262144
               AutoSize        =   1
               SplitterBarWidth=   1
               SplitterBarAppearance=   1
               PaneTree        =   "P_01004_A.frx":134D
               Begin FPSpreadADO.fpSpread spdCloth 
                  Height          =   10305
                  Left            =   30
                  TabIndex        =   43
                  Top             =   30
                  Width           =   6540
                  _Version        =   524288
                  _ExtentX        =   11536
                  _ExtentY        =   18177
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
                  MaxCols         =   6
                  SpreadDesigner  =   "P_01004_A.frx":137F
                  Appearance      =   1
                  HighlightHeaders=   1
                  HighlightStyle  =   1
               End
            End
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   330
            Index           =   10
            Left            =   -67480
            TabIndex        =   41
            Top             =   30
            Visible         =   0   'False
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Find"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_01004_A.frx":1AD8
         End
      End
      Begin Threed.SSPanel panSub 
         Height          =   2760
         Left            =   5145
         TabIndex        =   1
         Top             =   1335
         Width           =   4050
         _ExtentX        =   7144
         _ExtentY        =   4868
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin Threed.SSPanel SSPanel 
            Height          =   450
            Index           =   2
            Left            =   75
            TabIndex        =   37
            Top             =   585
            Width           =   3900
            _ExtentX        =   6879
            _ExtentY        =   794
            _Version        =   262144
            BackColor       =   16777215
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin Threed.SSOption optRound 
               Height          =   255
               Index           =   0
               Left            =   105
               TabIndex        =   38
               Top             =   105
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   262144
               BackColor       =   16777215
               Caption         =   "10원 반올림"
            End
            Begin Threed.SSOption optRound 
               Height          =   255
               Index           =   1
               Left            =   2010
               TabIndex        =   39
               Top             =   105
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   262144
               BackColor       =   16777215
               Caption         =   "10원 절삭"
               Value           =   -1
            End
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   450
            Index           =   1
            Left            =   75
            TabIndex        =   34
            Top             =   75
            Width           =   3900
            _ExtentX        =   6879
            _ExtentY        =   794
            _Version        =   262144
            BackColor       =   16777215
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin Threed.SSOption optClass 
               Height          =   255
               Index           =   0
               Left            =   105
               TabIndex        =   35
               Top             =   105
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   262144
               BackColor       =   16777215
               Caption         =   "본사품목단가"
               Value           =   -1
            End
            Begin Threed.SSOption optClass 
               Height          =   255
               Index           =   1
               Left            =   2010
               TabIndex        =   36
               Top             =   105
               Width           =   1770
               _ExtentX        =   3122
               _ExtentY        =   450
               _Version        =   262144
               BackColor       =   16777215
               Caption         =   "가맹점품목단가"
            End
         End
         Begin CSTextLibCtl.silgEdit txtRatio 
            Height          =   360
            Left            =   975
            TabIndex        =   24
            Top             =   1815
            Width           =   720
            _Version        =   262145
            _ExtentX        =   1270
            _ExtentY        =   635
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   255
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
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
            Justification   =   1
            BorderStyle     =   0
            Undo            =   1
            Data            =   0
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   975
            TabIndex        =   2
            Top             =   1095
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Format          =   63700992
            CurrentDate     =   36686
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   975
            TabIndex        =   3
            Top             =   1455
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Format          =   63700992
            CurrentDate     =   36686
         End
         Begin XtremeSuiteControls.PushButton btnApply 
            Height          =   450
            Left            =   975
            TabIndex        =   23
            Top             =   2250
            Width           =   1065
            _Version        =   851970
            _ExtentX        =   1879
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 할인적용"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
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
         Begin XtremeSuiteControls.PushButton btnCalculate 
            Height          =   450
            Left            =   2265
            TabIndex        =   33
            Top             =   2250
            Width           =   1725
            _Version        =   851970
            _ExtentX        =   3043
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 할인계산(&R)"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "P_01004_A.frx":2072
         End
         Begin VB.Label lblProgress 
            BackStyle       =   0  '투명
            Caption         =   "%"
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
            Left            =   1800
            TabIndex        =   32
            Top             =   1920
            Width           =   255
         End
         Begin VB.Label lblProgress 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "할인율:"
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
            Left            =   30
            TabIndex        =   31
            Top             =   1905
            Width           =   900
         End
         Begin VB.Label lblProgress 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "종료일자:"
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
            Left            =   30
            TabIndex        =   30
            Top             =   1515
            Width           =   900
         End
         Begin VB.Label lblProgress 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "시작일자:"
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
            Left            =   30
            TabIndex        =   29
            Top             =   1155
            Width           =   900
         End
      End
      Begin FPSpreadADO.fpSpread spdClass 
         Height          =   8040
         Left            =   5145
         TabIndex        =   4
         Top             =   4110
         Width           =   4050
         _Version        =   524288
         _ExtentX        =   7144
         _ExtentY        =   14182
         _StockProps     =   64
         BackColorStyle  =   1
         EditEnterAction =   2
         EditModeReplace =   -1  'True
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
         SpreadDesigner  =   "P_01004_A.frx":2A84
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   5
         Top             =   540
         Width           =   15855
         _ExtentX        =   27966
         _ExtentY        =   1376
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin Threed.SSOption optGubun 
            Height          =   255
            Index           =   0
            Left            =   195
            TabIndex        =   6
            Top             =   270
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   262144
            BackColor       =   16777215
            Caption         =   "가맹점종류별"
         End
         Begin Threed.SSOption optGubun 
            Height          =   255
            Index           =   1
            Left            =   1980
            TabIndex        =   7
            Top             =   270
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   450
            _Version        =   262144
            BackColor       =   16777215
            Caption         =   "지사별"
         End
         Begin XtremeSuiteControls.ProgressBar ProgressBar 
            Height          =   330
            Left            =   9180
            TabIndex        =   25
            Top             =   45
            Visible         =   0   'False
            Width           =   5325
            _Version        =   851970
            _ExtentX        =   9393
            _ExtentY        =   582
            _StockProps     =   93
            Scrolling       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ProgressBar ProgressBar2 
            Height          =   330
            Left            =   9180
            TabIndex        =   26
            Top             =   405
            Visible         =   0   'False
            Width           =   5325
            _Version        =   851970
            _ExtentX        =   9393
            _ExtentY        =   582
            _StockProps     =   93
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Shape Shape 
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00808080&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  '단색
            Height          =   630
            Index           =   0
            Left            =   60
            Shape           =   4  '둥근 사각형
            Top             =   75
            Width           =   3000
         End
         Begin VB.Label lblProgress 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "할인품목 저장:"
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
            Index           =   1
            Left            =   7890
            TabIndex        =   28
            Top             =   495
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.Label lblProgress 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "지사 저장:"
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
            Left            =   7890
            TabIndex        =   27
            Top             =   135
            Visible         =   0   'False
            Width           =   1260
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   11
         Top             =   15
         Width           =   8250
         _ExtentX        =   14552
         _ExtentY        =   900
         _Version        =   262144
         Font3D          =   3
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
         Caption         =   " 가맹점별 품목할인 등록 (P_01004_A)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_01004_A.frx":3059
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   0
         Left            =   8280
         TabIndex        =   12
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
         PictureBackground=   "P_01004_A.frx":325B
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   13
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
            Picture         =   "P_01004_A.frx":345D
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   14
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
            Picture         =   "P_01004_A.frx":39F7
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   15
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
            Picture         =   "P_01004_A.frx":3F91
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   16
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
            Picture         =   "P_01004_A.frx":452B
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   17
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
            Picture         =   "P_01004_A.frx":4AC5
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   18
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
            Picture         =   "P_01004_A.frx":505F
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   19
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
            Picture         =   "P_01004_A.frx":55F9
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   20
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
            Picture         =   "P_01004_A.frx":5B93
         End
      End
      Begin FPSpreadADO.fpSpread sprList 
         Height          =   7455
         Left            =   15
         TabIndex        =   21
         Top             =   4695
         Width           =   5115
         _Version        =   524288
         _ExtentX        =   9022
         _ExtentY        =   13150
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
         MaxCols         =   5
         ScrollBars      =   2
         SpreadDesigner  =   "P_01004_A.frx":612D
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread sprMaster 
         Height          =   3345
         Left            =   15
         TabIndex        =   22
         Top             =   1335
         Width           =   5115
         _Version        =   524288
         _ExtentX        =   9022
         _ExtentY        =   5900
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
         MaxCols         =   3
         ScrollBars      =   2
         SpreadDesigner  =   "P_01004_A.frx":677E
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_01004_A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim strSql As String
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub SPR_Resize()
    On Error GoTo ErrRtn
    
    sprSchedule.Width = Me.Width - 9500
    sprSchedule.Height = Me.Height - 3140

'    spdCloth.Width = Me.Width - 9500
'    spdCloth.Height = Me.Height - 2360

    Exit Sub
    
ErrRtn:

End Sub

Private Sub btnApply_Click()
    spdCloth.MaxRows = 0
    
    '------------------------------------------------------------------
    ' SP_01004_A_02
    '------------------------------------------------------------------
    ReDim sValue(0)
    
    sValue(0) = txtRatio.Value
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01004_A_02", sValue(), Err_Num, Err_Dec)
    
    With spdClass
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = RS01!의류분류코드 & "" '
            .Col = 2: .Text = RS01!의류분류명 & ""   '
            .Col = 3: .Value = RS01!할인율 & ""      '
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
End Sub

Private Sub btnCalculate_Click()
    Dim i    As Integer
    Dim iCnt As Integer
        
    If optClass(1).Value = True Then
        iCnt = 0
        
        With sprList
            For i = 1 To .MaxRows
                .Row = i
                .Col = 1
                
                If .Text = "1" Then
                    iCnt = iCnt + 1
                End If
            Next i
        End With
        
        If (iCnt = 0) Or (iCnt > 1) Then
            MsgBox "가맹점품목단가를 선택한 경우에는 1개의 가맹점만 선택해야 합니다.", vbInformation, "확인"
            
            Exit Sub
        End If
    End If
    
    If optClass(0).Value = True Then
        ReDim sValue(2)
        
        spdCloth.MaxRows = 0
        
        For i = 1 To spdClass.MaxRows
            spdClass.Row = i
            spdClass.Col = 1: sValue(0) = spdClass.Text  '의류분류코드
            spdClass.Col = 3: sValue(1) = spdClass.Value '할인율
        
            If optRound(0).Value = True Then
                sValue(2) = "0"
            Else
                sValue(2) = "1"
            End If
            
            Set RS01 = New ADODB.Recordset
            Set RS01 = ExecPro("SP_01004_A_03", sValue(), Err_Num, Err_Dec)
            
            With spdCloth
                .Redraw = False
                
                Do Until RS01.EOF
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
        
                    .Col = 1: .Text = RS01!의류코드 & "" ' 1
                    .Col = 2: .Text = RS01!의류명 & ""   ' 2
                    .Col = 3: .Text = RS01!정상가격 & "" ' 3
                    .Col = 4: .Text = RS01!할인가격 & "" ' 4
                    .Col = 5: .Text = RS01!할인률 & ""   ' 5
        
                    RS01.MoveNext
                Loop
                .Redraw = True
                    
                RS01.Close
                Set RS01 = Nothing
            End With
        Next i
        
    Else
        Dim 가맹점코드 As String
        
        With sprList
            For i = 1 To .MaxRows
                .Row = i
                .Col = 1
                
                If .Text = "1" Then
                    .Col = 2: 가맹점코드 = .Text & ""
                    Exit For
                End If
            Next i
        End With
        
        ReDim sValue(3)
        
        spdCloth.MaxRows = 0
        
        For i = 1 To spdClass.MaxRows
            spdClass.Row = i
                              sValue(0) = 가맹점코드     '
            spdClass.Col = 1: sValue(1) = spdClass.Text  '의류분류코드
            spdClass.Col = 3: sValue(2) = spdClass.Value '할인율
        
            If optRound(0).Value = True Then
                sValue(3) = "0"                          '
            Else
                sValue(3) = "1"                          '
            End If
            
            Set RS01 = New ADODB.Recordset
            Set RS01 = ExecPro("SP_01004_A_08", sValue(), Err_Num, Err_Dec)
            
            With spdCloth
                .Redraw = False
                
                Do Until RS01.EOF
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
        
                    .Col = 1: .Text = RS01!의류코드 & "" ' 1
                    .Col = 2: .Text = RS01!의류명 & ""   ' 2
                    .Col = 3: .Text = RS01!정상가격 & "" ' 3
                    .Col = 4: .Text = RS01!할인가격 & "" ' 4
                    .Col = 5: .Text = RS01!할인률 & ""   ' 5
        
                    RS01.MoveNext
                Loop
                .Redraw = True
                    
                RS01.Close
                Set RS01 = Nothing
            End With
        Next i
    End If
End Sub

Private Sub cboInput_Click(Index As Integer)
    Dim sCode As String

    If Index = 0 Then
        sCode = Trim(Mid(cboInput(0).Text, 2, 4))

        Call Get_가맹점리스트(cboInput(1), sCode)

    ElseIf Index = 1 Then
           Call Data_Display
           
           Call Data_Display2(Mid(cboInput(1).Text, 2, 6))
    
    End If
End Sub

Private Sub cboInput_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SearchString_한글 KeyAscii
'    Else
       ' SearchString KeyAscii
    End If

End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
'    Me.MousePointer = 11
    
    Select Case Index
        Case 0: 'Call Data_Display   ' 조회
        Case 1: 'Call DataAdd        ' 신규
        Case 2: Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
        Case 4: Call DataCancel     ' 취소
        Case 5: 'Call DataPrint      ' 인쇄
        Case 6: Call Export_Excel(P_00000.cdgExcel, spdCloth)      ' 엑셀
        Case 7: Unload Me           ' 종료
        Case 10
            txtFind.Visible = Not txtFind.Visible
            txtFind.SelStart = 0:   txtFind.SelLength = Len(txtFind.Text)
            txtFind.SetFocus
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

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(1).Enabled = True
    cmdBtn(2).Enabled = True
    cmdBtn(3).Enabled = True
    cmdBtn(4).Enabled = True
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    With sprMaster
        .MaxRows = 0
        .RowHeight(-1) = 13
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle

        '----------------------------------------------------------------------
        '
        '----------------------------------------------------------------------
        '해줘야 블럭이 잡히지 않는다.
        'EditModePermanent = True
        
        .Col = 1
        .Row = 0
        .CellType = CellTypeCheckBox
        .BackColor = &HE0E0E0
        .TypeCheckCenter = True
        .Value = False

        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeNormal
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
    
    With sprList
        .MaxRows = 0
        .RowHeight(-1) = 13
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle

        '----------------------------------------------------------------------
        '
        '----------------------------------------------------------------------
        '해줘야 블럭이 잡히지 않는다.
        'EditModePermanent = True
        
        .Col = 1
        .Row = 0
        .CellType = CellTypeCheckBox
        .BackColor = &HE0E0E0
        .TypeCheckCenter = True
        .Value = False

        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeNormal
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
    
    With sprSchedule
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
    
    With spdCloth
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
        '.UserColAction = UserColActionSort
    End With
    
    With spdClass
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
        '.UserColAction = UserColActionSort
    End With

    dtInput(0).Value = Date
    dtInput(1).Value = Date
        
    TabControl1.SelectedItem = 0

    Call Get_지사리스트(cboInput(0), False)
            
    '-------------------------------------------------------------
    ' TB_의류분류
    '-------------------------------------------------------------
    ReDim sValue(0)

    sValue(0) = "0"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_00013", sValue(), Err_Num, Err_Dec)

    With spdClass
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = RS01!의류분류코드 & ""
            .Col = 2: .Text = RS01!의류분류명 & ""
            .Col = 3: .Value = 0
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
    
    Call SPR_Resize
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Resize()
'    Call SPR_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'P_01004_A_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    If Trim(cboInput(0).Text) = "" Then
        'MsgBox "사업장을 선택 하세요", vbInformation
        cboInput(0).SetFocus
        Exit Sub
    End If
    
    If Trim(cboInput(1).Text) = "" Then
        'MsgBox "가맹점을 선택 하세요", vbInformation
        cboInput(1).SetFocus
        Exit Sub
    End If
    
    '-------------------------------------------------------------------
    ' SP_01004_A_00
    '-------------------------------------------------------------------
    ReDim sValue(0)

    sValue(0) = Mid(cboInput(1).Text, 2, 6)
        
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01004_A_00", sValue(), Err_Num, Err_Dec)
    
    With sprSchedule
        .MaxRows = 0
        .Redraw = False
        .EventEnabled(EventButtonClicked) = False '버튼클릭 이벤트 죽임

        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 1: .Text = IIf(RS01!삭제 & "" = "Y", True, False)
            .Col = 2: .Text = RS01!가맹점코드 & ""
            .Col = 3: .Text = RS01!가맹점명 & ""
            .Col = 4: .Text = RS01!시작일자 & ""
            .Col = 5: .Text = RS01!종료일자 & ""
            
            ' 적용 대상일 경우
            If RS01!시작일자 & "" <= Format(Date, "yyyy-MM-dd") And RS01!종료일자 & "" >= Format(Date, "yyyy-MM-dd") Then
                .Col = -1
                .BackColor = vbGreen
            End If
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .EventEnabled(EventButtonClicked) = True
        .Redraw = True
    End With
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Public Sub DataAdd()
'    sprSchedule.MaxRows = sprSchedule.MaxRows + 1
'    sprSchedule.Row = sprSchedule.MaxRows
'    sprSchedule.Col = 1
'    sprSchedule.Action = ActionActiveCell
'    sprSchedule.Lock = False
'
''    optSelect(0).Value = True
''    cmdSubBtn(0).Enabled = True
'    txtRatio.Value = 0
'
'    spdCloth.MaxRows = 0
'    spdClass.MaxRows = 0
'
''    If optSelect(1).Value Then
''        cboInput(2).ListIndex = 1
''    End If
End Sub

Public Sub DataSave()
    Dim i          As Integer
    Dim iRow       As Long
    
    Dim 가맹점코드 As String
    Dim 할인율     As Long
    Dim 정상금액   As Long
    Dim 할인금액   As Long
    Dim Ret        As Long
    
    On Error GoTo ErrRtn
    
    lblProgress(0).Visible = True
    lblProgress(1).Visible = True
    
    ProgressBar.Value = 0
    ProgressBar.Min = 0
    ProgressBar.Max = 100
    ProgressBar.Visible = True
    
    ProgressBar2.Value = 0
    ProgressBar2.Min = 0
    ProgressBar2.Max = 100
    ProgressBar2.Visible = True
    
    Set RS01 = New ADODB.Recordset
    
    For iRow = 1 To sprList.MaxRows
        sprList.Row = iRow
        sprList.Col = 1
        
        If sprList.Text = "1" Then
            ProgressBar.Value = (iRow / sprList.MaxRows) * 100
            DoEvents
            
            sprList.Col = 2: 가맹점코드 = sprList.Text
            
            '-------------------------------------------------------------------
            ' TB_가맹점할인 삭제 - SP_01004_A_01
            '-------------------------------------------------------------------
            ReDim sValue(2)
            
            sValue(0) = 가맹점코드                             '
            sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD") '
            sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD") '
            
            Call ExecPro("SP_01004_A_01", sValue(), Err_Num, Err_Dec)
            
            If Err_Num <> 0 Then
                MsgBox "[" & Err_Num & "] " & Err_Dec
                
                lblProgress(0).Visible = False
                lblProgress(1).Visible = False
                
                ProgressBar.Visible = False
                ProgressBar2.Visible = False
                Exit Sub
            End If
        
            i = 0
            
            ReDim sValue(0)
            
            sValue(0) = 가맹점코드
            
            Set RS01 = New ADODB.Recordset
            Set RS01 = ExecPro("SP_01005_B_02", sValue(), Err_Num, Err_Dec)

            Do Until RS01.EOF
                i = i + 1
                
                ProgressBar2.Value = (i / RS01.RecordCount) * 100
                DoEvents
                
                With spdCloth
                    Ret = .SearchCol(1, 0, -1, RS01!의류코드, SearchFlagsValue)
                    
                    If Ret > 0 Then
                        .Row = Ret
                        
                        .Col = 3: 정상금액 = .Value
                        .Col = 4: 할인금액 = .Value
                        
                        If 정상금액 = 할인금액 Then
                            할인율 = 0
                        Else
                            .Col = 6
                            If .Text = "1" Then
                                할인율 = 0
                            Else
                                .Col = 5: 할인율 = .Value '할인율
                            End If
                        End If
                    Else
                        할인율 = txtRatio.Value '할인율
                    End If
                End With
            
               '-------------------------------------------------------------------
                ' TB_가맹점할인 저장 - SP_01004_A_09
                '-------------------------------------------------------------------
                ReDim sValue(7)
                
                sValue(0) = 가맹점코드                             ' 0 가매점코드
                sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD") ' 1 시작일자
                sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD") ' 2 종료일자
                sValue(3) = Trim(RS01!의류코드) & ""               ' 3 의류코드
                sValue(4) = Trim(RS01!의류명) & ""                 ' 4 의류명
                sValue(5) = RS01!금액                              ' 5 정상가격
                sValue(6) = 할인율                                 ' 6 할인률
                                
                If optRound(0).Value = True Then
                    sValue(7) = "0"                                ' 7
                Else
                    sValue(7) = "1"                                ' 7
                End If
            
                Call ExecPro("SP_01004_A_09", sValue(), Err_Num, Err_Dec)
                
                RS01.MoveNext
            Loop
            RS01.Close
            Set RS01 = Nothing
        End If
    Next iRow
        
    lblProgress(0).Visible = False
    lblProgress(1).Visible = False
        
    ProgressBar.Visible = False
    ProgressBar2.Visible = False
    
    ' 지사, 대리점 선택을 먼저 처리해줘야 한다.
    'Call Data_Display
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Public Sub DataDelete()
'    If MsgBox("해당되는 데이터를 삭제하시겠습니까?", vbYesNo + vbInformation + vbDefaultButton2, "데이터 삭제") = vbYes Then
'
'        ReDim sValue(2)
'
'        sValue(0) = Mid(cboInput(1).Text, 2, 6)
'
'        sprSchedule.Row = sprSchedule.ActiveRow
'        sprSchedule.Col = 1: sValue(1) = Format(sprSchedule.Text, "YYYY-MM-DD")
'        sprSchedule.Col = 2: sValue(2) = Format(sprSchedule.Text, "YYYY-MM-DD")
'
'        Call ExecPro("SP_01004_A_01", sValue(), Err_Num, Err_Dec)
'
'        If Err_Num = 0 Then
'            sprSchedule.Row = sprSchedule.ActiveRow
'            sprSchedule.Action = ActionDeleteRow
'            sprSchedule.MaxRows = sprSchedule.MaxRows - 1
'
'            MsgBox "해당되는 데이터가 정상적으로 삭제가 되었습니다.", vbInformation
'        Else
'            MsgBox "오류 " & Err_Num & ":" & Err_Dec & " ", vbInformation
'        End If
'
'        Call Data_Display
'    End If
End Sub

Public Sub DataCancel()
    Call Data_Display
End Sub

Private Sub optGubun_Click(Index As Integer, Value As Integer)
    sprList.MaxRows = 0
    DoEvents
    
    If Index = 0 Then
        ReDim sValue(0)
        
        sValue(0) = "0"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_01004_A_06", sValue(), Err_Num, Err_Dec)
    
        With sprMaster
            .MaxRows = 0
            
            .Row = SpreadHeader
            .Col = 1: .Text = "0"
            
            .Redraw = False
            
            Do Until RS01.EOF
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                
                .Col = 1: .Text = "0"
                .Col = 2: .Text = RS01!가맹점구분코드 & ""
                .Col = 3: .Text = RS01!가맹점구분명 & ""
                
                RS01.MoveNext
            Loop
            RS01.Close
            Set RS01 = Nothing
            
            .Redraw = True
        End With
        
    Else
        ReDim sValue(0)
        
        sValue(0) = "0"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_A_0001", sValue(), Err_Num, Err_Dec)
    
        With sprMaster
            .EventEnabled(EventButtonClicked) = False '버튼클릭 이벤트 죽임
            
            .MaxRows = 0
            
            .Row = SpreadHeader
            .Col = 1: .Text = "0"
            
            .Redraw = False
            
            Do Until RS01.EOF
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                
                .Col = 1: .Text = "0"
                .Col = 2: .Text = RS01!지사코드 & ""
                .Col = 3: .Text = RS01!지사명 & ""
                
                RS01.MoveNext
            Loop
            RS01.Close
            Set RS01 = Nothing
            
            .EventEnabled(EventButtonClicked) = True
            
            .Redraw = True
        End With
    End If
End Sub

Private Sub spdClass_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim vText   As Variant
    
    spdClass.GetText 1, Row, vText
    
    Rtn = spdCloth.SearchCol(1, 1, spdCloth.MaxRows, Trim(CStr(vText)), SearchFlagsPartialMatch)
    
    spdCloth.Redraw = False
    
    spdCloth.Row = spdCloth.MaxRows
    spdCloth.Action = ActionActiveCell
    
    spdCloth.Row = Rtn
    spdCloth.Action = ActionActiveCell
    
    spdCloth.Redraw = True
    DoEvents

End Sub

Private Sub sprList_Click(ByVal Col As Long, ByVal Row As Long)
    Dim varGet As Variant

    If Row = 0 And Col = 1 Then
        With sprList
            .Redraw = False
            .GetText 1, 0, varGet       '헤더의 체크값을 가져옴

            .EventEnabled(EventButtonClicked) = False   '체크값 변경시 ButtonClicked 이벤트를 발생시키지 않도록
            
            '----- 헤더부터 마지막행까지 블럭으로 설정 -----
            .Col = 1
            .Col2 = 1
            .Row = 0            '0부터해야 해더까지 변경됨
            .Row2 = .MaxRows
            .BlockMode = True
            .Value = IIf(varGet = 0, "1", "0")          '헤더의 체크값에 따라서 체크값을 변경함
            .BlockMode = False

            .EventEnabled(EventButtonClicked) = True    '이벤트를 다시 활성시킴
            .Redraw = True
        End With
    End If
End Sub

'-----------------------------------------------------------------
'
' EventButtonClicked 이벤트를 살려놓아야 한다.
'
'-----------------------------------------------------------------
Private Sub sprMaster_Click(ByVal Col As Long, ByVal Row As Long)
    Dim varGet As Variant

    If Row = 0 And Col = 1 Then
        With sprMaster
            .Redraw = False
            .GetText 1, 0, varGet       '헤더의 체크값을 가져옴

            '.EventEnabled(EventButtonClicked) = False   '체크값 변경시 ButtonClicked 이벤트를 발생시키지 않도록
            
            '----- 헤더부터 마지막행까지 블럭으로 설정 -----
            .Col = 1
            .Col2 = 1
            .Row = 0            '0부터해야 해더까지 변경됨
            .Row2 = .MaxRows
            .BlockMode = True
            .Value = IIf(varGet = 0, "1", "0")          '헤더의 체크값에 따라서 체크값을 변경함
            .BlockMode = False

            '.EventEnabled(EventButtonClicked) = True    '이벤트를 다시 활성시킴
            .Redraw = True
        End With
    End If
End Sub

Private Sub sprMaster_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim sWork        As String
    Dim 코드         As String
    Dim 가맹정구분명 As String
    
    If Row <= 0 Then Exit Sub
    
    If Col = 1 Then
        sprMaster.Row = Row
        sprMaster.Col = 1
        
        If sprMaster.Text = "1" Then
            sWork = "ADD"
        Else
            sWork = "DEL"
        End If
        
        sprMaster.Col = 2: 코드 = Trim(sprMaster.Text) & ""
        sprMaster.Col = 3: 가맹정구분명 = Trim(sprMaster.Text) & ""
        
        If optGubun(0).Value = True Then
            Call 가맹점2_Display(sWork, 코드, 가맹정구분명)
        Else
            Call 가맹점_Display(sWork, 코드)
        End If
    End If
End Sub

Private Sub 가맹점_Display(sWork As String, 지사코드 As String)
    Dim Ret As Long
    
    If sWork = "ADD" Then
        ReDim sValue(0)
        
        sValue(0) = 지사코드
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_01001_00_MASTER", sValue(), Err_Num, Err_Dec)
        
        With sprList
            .Redraw = False
            
            Do Until RS01.EOF
                Ret = .SearchCol(2, 0, -1, RS01(0), SearchFlagsValue)
                
                If Ret > 0 Then
                    '이미 존재하는 가맹점...
                Else
                    If (RS01!종료일자 = "") Or (RS01!종료일자 = "2099-12-31") Then  '가맹점이 현 지사에서 관리중...
                        .MaxRows = .MaxRows + 1
                        .Row = .MaxRows
                        
                        .Col = 1: .Text = "1"           '1
                        .Col = 2: .Text = RS01!가맹점코드 & ""  '2
                        .Col = 3: .Text = RS01!가맹점명 & ""  '3
                        .Col = 4: .Text = 지사코드 & "" '4
                    End If
                End If
                
                RS01.MoveNext
            Loop
            RS01.Close
            Set RS01 = Nothing
            
            .Redraw = True
        End With
    Else
        With sprList
            .Redraw = False
                
            Do
                Ret = .SearchCol(4, 0, -1, 지사코드, SearchFlagsValue)
                
                If Ret > 0 Then
                    .DeleteRows Ret, 1
                    
                    .MaxRows = .MaxRows - 1
                End If
            Loop Until Ret = -1
        End With
    End If
End Sub


Private Sub 가맹점2_Display(sWork As String, 가맹점구분 As String, 가맹점구분명 As String)
    Dim Ret As Long
    
    If sWork = "ADD" Then
        ReDim sValue(0)
        
        sValue(0) = 가맹점구분
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_01004_A_07", sValue(), Err_Num, Err_Dec)
        
        With sprList
            .Redraw = False
            
            Do Until RS01.EOF
                Ret = .SearchCol(2, 0, -1, RS01(0), SearchFlagsValue)
                
                If Ret > 0 Then
                    '이미 존재하는 가맹점...
                Else
                    If (RS01!종료일자 = "") Or (RS01!종료일자 = "2099-12-31") Then '가맹점이 현 지사에서 관리중...
                        .MaxRows = .MaxRows + 1
                        .Row = .MaxRows
                        
                        .Col = 1: .Text = "1"                  '1
                        .Col = 2: .Text = RS01!가맹점코드 & "" '2
                        .Col = 3: .Text = RS01!가맹점명 & ""   '3
                        .Col = 4: .Text = RS01!지사코드 & ""   '4
                        .Col = 5: .Text = 가맹점구분명 & ""    '5
                    End If
                End If
                
                RS01.MoveNext
            Loop
            RS01.Close
            Set RS01 = Nothing
            
            .Redraw = True
        End With
    Else
        With sprList
            .Redraw = False
                
            Do
                Ret = .SearchCol(5, 0, -1, 가맹점구분명, SearchFlagsValue)
                
                If Ret > 0 Then
                    .DeleteRows Ret, 1
                    
                    .MaxRows = .MaxRows - 1
                End If
            Loop Until Ret = -1
        End With
    End If
End Sub

Private Sub sprSchedule_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    On Error GoTo ERR_RTN
    
    If Col = 1 Then
        Dim vText   As Variant
        ReDim sValue(3)
                                                
        sprSchedule.GetText 2, Row, vText:      sValue(0) = CStr(vText)
        sprSchedule.GetText 4, Row, vText:      sValue(1) = CStr(vText)
        sprSchedule.GetText 5, Row, vText:      sValue(2) = CStr(vText)
        sprSchedule.GetText 1, Row, vText:      sValue(3) = IIf(CStr(vText) = "1", "Y", "N")
    
        Call ExecPro("SP_01004_A_22", sValue(), Err_Num, Err_Dec)
        
        If Err_Num <> 0 Then GoTo ERR_RTN
    
    End If
    Exit Sub
    
ERR_RTN:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)

End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
    
        ' 조회에서 매장 코드가 설정된 경우 해당 매장이 선택 되도록 한다.
        DoEvents
        If Trim(txtFind.Text) = "" Then Exit Sub
        
        With spdCloth
            .Redraw = False
            Rtn = .SearchCol(2, 1, .MaxRows, Trim(txtFind.Text), SearchFlagsPartialMatch)
            
            .Row = .MaxRows
            .Action = ActionActiveCell
            
            .Row = Rtn
            .Action = ActionActiveCell
            
            .Redraw = True
            DoEvents
            
        End With
    
    End If
End Sub



'---------------------------------------------------------------------
' SP_01004_00 - TB_가맹점할인
'---------------------------------------------------------------------
Private Sub Data_Display2(가맹점코드 As String)
    ReDim sValue(0)
    
    sValue(0) = 가맹점코드
        
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01004_00", sValue(), Err_Num, Err_Dec)
    
    With spdList
        .MaxRows = 0
        .Redraw = False
                    
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = RS01!시작일자 & ""
            .Col = 2: .Text = RS01!종료일자 & ""
            .Col = 3: .Text = RS01!할인율 & ""
            
            ' 적용 대상일 경우
            If RS01!시작일자 & "" <= Format(Date, "yyyy-MM-dd") And RS01!종료일자 & "" >= Format(Date, "yyyy-MM-dd") Then
                .Col = -1
                .BackColor = vbGreen
            End If
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .SortKey(1) = 1
        .SortKeyOrder(1) = SortKeyOrderDescending
        .Sort -1, -1, -1, -1, SortByRow
            
        .Redraw = True
    End With
End Sub

