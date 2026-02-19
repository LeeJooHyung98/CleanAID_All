VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_04027 
   Caption         =   "매장 매출 현황(보고용-가맹점기준)"
   ClientHeight    =   12450
   ClientLeft      =   1380
   ClientTop       =   2250
   ClientWidth     =   16635
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_04027.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12450
   ScaleWidth      =   16635
   WindowState     =   2  '최대화
   Begin VB.CommandButton cmdAllCheck 
      Caption         =   "선택"
      Height          =   315
      Left            =   2580
      TabIndex        =   25
      Top             =   1380
      Width           =   555
   End
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   12450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16635
      _ExtentX        =   29342
      _ExtentY        =   21960
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_04027.frx":058A
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   11085
         Left            =   15
         TabIndex        =   1
         Top             =   1350
         Width           =   5730
         _Version        =   524288
         _ExtentX        =   10107
         _ExtentY        =   19553
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
         SpreadDesigner  =   "P_04027.frx":065C
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   795
         Left            =   15
         TabIndex        =   2
         Top             =   540
         Width           =   16605
         _ExtentX        =   29289
         _ExtentY        =   1402
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.CommandButton Command1 
            Caption         =   "?"
            Height          =   675
            Left            =   15930
            TabIndex        =   37
            Top             =   30
            Width           =   615
         End
         Begin XtremeSuiteControls.GroupBox GroupBox1 
            Height          =   735
            Index           =   0
            Left            =   6300
            TabIndex        =   26
            Top             =   -30
            Width           =   1425
            _Version        =   851970
            _ExtentX        =   2514
            _ExtentY        =   1296
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Begin XtremeSuiteControls.CheckBox chkTeamSelect 
               Height          =   195
               Index           =   2
               Left            =   240
               TabIndex        =   27
               Tag             =   "유통매장"
               Top             =   480
               Width           =   885
               _Version        =   851970
               _ExtentX        =   1561
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   " 2 팀"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox chkTeamSelect 
               Height          =   195
               Index           =   1
               Left            =   240
               TabIndex        =   28
               Tag             =   "일반매장"
               Top             =   180
               Width           =   885
               _Version        =   851970
               _ExtentX        =   1561
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   " 1 팀"
               UseVisualStyle  =   -1  'True
            End
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1500
            TabIndex        =   3
            Top             =   75
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   64618496
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   645
            Index           =   0
            Left            =   60
            TabIndex        =   4
            Top             =   75
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   1138
            _Version        =   262144
            Caption         =   "매 출 기 간"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   1500
            TabIndex        =   5
            Top             =   420
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   64618496
            CurrentDate     =   36686
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   660
            Index           =   8
            Left            =   4530
            TabIndex        =   24
            Top             =   60
            Width           =   1650
            _Version        =   851970
            _ExtentX        =   2910
            _ExtentY        =   1164
            _StockProps     =   79
            Caption         =   "매장 찾기"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_04027.frx":0C3E
         End
         Begin XtremeSuiteControls.GroupBox GroupBox1 
            Height          =   735
            Index           =   1
            Left            =   7830
            TabIndex        =   29
            Top             =   -30
            Width           =   7995
            _Version        =   851970
            _ExtentX        =   14102
            _ExtentY        =   1296
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Begin XtremeSuiteControls.CheckBox chkSelect 
               Height          =   195
               Index           =   1
               Left            =   1740
               TabIndex        =   30
               Tag             =   "유통매장"
               Top             =   150
               Width           =   1185
               _Version        =   851970
               _ExtentX        =   2090
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "유통매장"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox chkSelect 
               Height          =   195
               Index           =   0
               Left            =   240
               TabIndex        =   31
               Tag             =   "일반매장"
               Top             =   150
               Width           =   1185
               _Version        =   851970
               _ExtentX        =   2090
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "일반매장"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox chkSelect 
               Height          =   195
               Index           =   2
               Left            =   3150
               TabIndex        =   32
               Tag             =   "이마트"
               Top             =   150
               Width           =   945
               _Version        =   851970
               _ExtentX        =   1667
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "이마트"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox chkSelect 
               Height          =   195
               Index           =   3
               Left            =   4290
               TabIndex        =   33
               Tag             =   "크렌즈"
               Top             =   150
               Width           =   1095
               _Version        =   851970
               _ExtentX        =   1931
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "크렌즈"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox chkSelect2 
               Height          =   195
               Index           =   0
               Left            =   240
               TabIndex        =   34
               Tag             =   "폐점"
               Top             =   480
               Width           =   1395
               _Version        =   851970
               _ExtentX        =   2461
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "폐점 포함"
               ForeColor       =   255
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox chkSelect 
               Height          =   195
               Index           =   4
               Left            =   5430
               TabIndex        =   35
               Tag             =   "유니트샵"
               Top             =   150
               Width           =   1185
               _Version        =   851970
               _ExtentX        =   2090
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "유니트샵"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox chkSelect3 
               Height          =   195
               Left            =   1740
               TabIndex        =   36
               Tag             =   "폐점"
               Top             =   480
               Width           =   4665
               _Version        =   851970
               _ExtentX        =   8229
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "객수 포함(조회 시간이 길어질 수 있습니다.)"
               ForeColor       =   255
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox chkTotal 
               Height          =   195
               Left            =   6600
               TabIndex        =   47
               Tag             =   "폐점"
               Top             =   480
               Width           =   1395
               _Version        =   851970
               _ExtentX        =   2461
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "합계 포함"
               ForeColor       =   255
               UseVisualStyle  =   -1  'True
               Value           =   1
            End
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   6
         Top             =   15
         Width           =   8865
         _ExtentX        =   15637
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
         Caption         =   "매장 매출 현황(보고용-가맹점기준) (P_04027)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_04027.frx":11D8
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   0
         Left            =   8895
         TabIndex        =   7
         Top             =   15
         Width           =   7725
         _ExtentX        =   13626
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
         PictureBackground=   "P_04027.frx":13DA
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6780
            TabIndex        =   8
            Top             =   0
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "종료"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_04027.frx":15DC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   9
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
            Picture         =   "P_04027.frx":1B76
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   10
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
            Picture         =   "P_04027.frx":2110
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   11
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
            Picture         =   "P_04027.frx":26AA
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   12
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
            Picture         =   "P_04027.frx":2C44
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   13
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
            Picture         =   "P_04027.frx":31DE
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   14
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
            Picture         =   "P_04027.frx":3778
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   21
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
            Picture         =   "P_04027.frx":3D12
         End
      End
      Begin FPSpreadADO.fpSpread spdView2 
         Height          =   10230
         Left            =   5760
         TabIndex        =   15
         Top             =   1350
         Width           =   10860
         _Version        =   524288
         _ExtentX        =   19156
         _ExtentY        =   18045
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
         SpreadDesigner  =   "P_04027.frx":42AC
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   840
         Index           =   1
         Left            =   5760
         TabIndex        =   16
         Top             =   11595
         Width           =   10860
         _ExtentX        =   19156
         _ExtentY        =   1482
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   14
            Left            =   60
            TabIndex        =   17
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   582
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
            Caption         =   "매출액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   15
            Left            =   3150
            TabIndex        =   18
            Top             =   60
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   582
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
            Caption         =   "이전 년도 매출액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   330
            Index           =   0
            Left            =   1200
            TabIndex        =   19
            Top             =   60
            Width           =   1965
            _Version        =   262145
            _ExtentX        =   3466
            _ExtentY        =   582
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
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
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   330
            Index           =   1
            Left            =   4740
            TabIndex        =   20
            Top             =   60
            Width           =   1965
            _Version        =   262145
            _ExtentX        =   3466
            _ExtentY        =   582
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
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
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   1
            Left            =   6690
            TabIndex        =   22
            Top             =   60
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   582
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
            Caption         =   "전년 대비 매출 신장율"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   330
            Index           =   2
            Left            =   8850
            TabIndex        =   23
            Top             =   60
            Width           =   1125
            _Version        =   262145
            _ExtentX        =   1984
            _ExtentY        =   582
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   "0.0"
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
            FmtControl      =   1
            NumDecDigits    =   1
            Undo            =   0
            Data            =   0
         End
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   4635
      Left            =   5790
      TabIndex        =   38
      Top             =   1380
      Visible         =   0   'False
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   8176
      _Version        =   262144
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "※"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   46
         Top             =   3150
         Width           =   210
      End
      Begin VB.Label Label1 
         Caption         =   "팀을 선택후 매장 구분을 선택 한 경우 해당 팀의 매장 구분을 조회할 수 있다."
         Height          =   405
         Index           =   3
         Left            =   510
         TabIndex        =   45
         Top             =   3150
         Width           =   4635
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "3."
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   44
         Top             =   2010
         Width           =   210
      End
      Begin VB.Label Label1 
         Caption         =   "매장의 구분에서 선택한 경우 해당 선택된 매장 중에서 매장 구분에 해당하는 매장만 선택 처리 한다. "
         Height          =   675
         Index           =   2
         Left            =   510
         TabIndex        =   43
         Top             =   2010
         Width           =   4635
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "2."
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   42
         Top             =   1410
         Width           =   210
      End
      Begin VB.Label Label1 
         Caption         =   "팀을 선택할 경우 기존의 모든 선택은 취소가 되며 해당 하는 팀만 다시 선택이 된다."
         Height          =   405
         Index           =   1
         Left            =   510
         TabIndex        =   41
         Top             =   1410
         Width           =   4635
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "1."
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   40
         Top             =   510
         Width           =   210
      End
      Begin VB.Label Label1 
         Caption         =   "가맹점명 옆의 선택은 폐점을 제외한 모든 가맹점을 선택 하는 기능이며 1팀 2팀 선택과는 무관하다."
         Height          =   705
         Index           =   0
         Left            =   510
         TabIndex        =   39
         Top             =   510
         Width           =   4635
      End
   End
End
Attribute VB_Name = "P_04027"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String
Dim m_MasterSelCodeList     As String
Dim m_StoreTypeSelCodeList  As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub chkSelect_Click(Index As Integer)
    Dim nRow        As Long
    Dim vText       As Variant
    
    m_StoreTypeSelCodeList = ""
    
    If chkSelect(0).Value = xtpChecked Then m_StoreTypeSelCodeList = m_StoreTypeSelCodeList & chkSelect(0).Caption & " "
    If chkSelect(1).Value = xtpChecked Then m_StoreTypeSelCodeList = m_StoreTypeSelCodeList & chkSelect(1).Caption & " "
    If chkSelect(2).Value = xtpChecked Then m_StoreTypeSelCodeList = m_StoreTypeSelCodeList & chkSelect(2).Caption & " "
    If chkSelect(3).Value = xtpChecked Then m_StoreTypeSelCodeList = m_StoreTypeSelCodeList & chkSelect(3).Caption & " "
    If chkSelect(4).Value = xtpChecked Then m_StoreTypeSelCodeList = m_StoreTypeSelCodeList & chkSelect(4).Caption & " "
    
    With spdView
        .EventEnabled(EventButtonClicked) = False
        
        For nRow = 1 To .MaxRows
            
            'If nRow = 40 Then MsgBox nRow
            
            ' 선택된 지사의 가맹점 일 경우만 처리한다.
            .GetText 3, nRow, vText
            If InStr(m_MasterSelCodeList, CStr(vText)) > 0 And Trim(CStr(vText)) <> "" Then
            
                ' 해당 매장의 종류일 경우
                .GetText 4, nRow, vText
                If Trim(CStr(vText)) <> "" Then .SetText 2, nRow, IIf(InStr(m_StoreTypeSelCodeList, CStr(vText)) > 0, "1", "0")
                
                ' 해당 지사이면서 폐점일 경우
                If InStr(m_StoreTypeSelCodeList, CStr(vText)) > 0 Then
                    .GetText 5, nRow, vText
                    If Trim(CStr(vText)) = "폐점" Then .SetText 2, nRow, IIf(chkSelect2(0).Value = xtpChecked, "1", "0")
                
                End If
            End If
        Next nRow
    
        .EventEnabled(EventButtonClicked) = True
    End With
End Sub

' 폐점 포함
Private Sub chkSelect2_Click(Index As Integer)
    Dim Idx     As Long
    Dim nRow    As Long
    Dim vText   As Variant
    
    With spdView
        .EventEnabled(EventButtonClicked) = False
        
        For nRow = 1 To .MaxRows
            
            ' 선택된 지사일 경우
            .GetText 3, nRow, vText
            If InStr(m_MasterSelCodeList, CStr(vText)) > 0 And Trim(CStr(vText)) <> "" Then
                
                ' 현재의 가맹점 종류와 선택 가맹점의 종류가 같을 경우
                .GetText 4, nRow, vText
                If InStr(m_StoreTypeSelCodeList, CStr(vText)) > 0 And Trim(CStr(vText)) <> "" Then
                    
                    ' 해당 지사이면서 폐점일 경우
                    .GetText 5, nRow, vText
                    If Trim(CStr(vText)) = "폐점" Then .SetText 2, nRow, IIf(chkSelect2(0).Value = xtpChecked, "1", "0")
                Else
                    .SetText 2, nRow, "0"
                End If
            End If
        Next nRow
        
        .EventEnabled(EventButtonClicked) = True
    End With
End Sub


Private Sub chkTeamSelect_Click(Index As Integer)
    Dim nRow        As Long
    Dim vText       As Variant
    Dim sTeamSel    As String   ' 팀선택
    Dim nMstCnt     As Long
    
    sTeamSel = ""
    m_MasterSelCodeList = ""
    
    chkSelect(0).Value = xtpUnchecked
    chkSelect(1).Value = xtpUnchecked
    chkSelect(2).Value = xtpUnchecked
    chkSelect(3).Value = xtpUnchecked
    chkSelect(4).Value = xtpUnchecked
    chkSelect2(0).Value = xtpUnchecked
    
    
    If chkTeamSelect(1).Value = xtpChecked Then sTeamSel = sTeamSel & chkTeamSelect(1).Caption & " "
    If chkTeamSelect(2).Value = xtpChecked Then sTeamSel = sTeamSel & chkTeamSelect(2).Caption & " "
    
    With spdView
        .EventEnabled(EventButtonClicked) = False
        
        nMstCnt = 0
        For nRow = 1 To .MaxRows
            .GetText 1, nRow, vText
            If Not IsNumeric(Mid(CStr(vText), 6, 1)) Then
                nMstCnt = nMstCnt + 1
            Else
                Exit For
            End If
        Next nRow
         
        
        ' 팀에 해당 하는 지사를 선택하고 해당 지사를 리스트화 한다.
        For nRow = 1 To nMstCnt
            .GetText 3, nRow, vText
            If Trim(CStr(vText)) <> "" Then Exit For
            
            .GetText 4, nRow, vText
            If CStr(vText) <> "" Then
                .SetText 2, nRow, IIf(InStr(sTeamSel, CStr(vText)) > 0, "1", "0")
            
                If InStr(sTeamSel, CStr(vText)) > 0 Then
                    .GetText 1, nRow, vText
                    m_MasterSelCodeList = m_MasterSelCodeList & Mid(CStr(vText), 2, 4) & " "
                End If
            End If
        Next nRow
    
        ' 위에서 선택된 지사의 가맹점을 다시 선택 한다.
        For nRow = nMstCnt + 1 To .MaxRows
            
            .GetText 5, nRow, vText
            ' 해당 지사이면서 폐점일 경우
            If Trim(CStr(vText)) = "폐점" Then
                .GetText 3, nRow, vText
                If InStr(m_MasterSelCodeList, CStr(vText)) > 0 Then
                    .SetText 2, nRow, IIf(chkSelect2(0).Value = xtpChecked, "1", "0")
                Else
                    .SetText 2, nRow, "0"
                End If
            
            ' 지사일 경우
            Else
                .GetText 3, nRow, vText
                If Trim(CStr(vText)) <> "" Then .SetText 2, nRow, IIf(InStr(m_MasterSelCodeList, CStr(vText)) > 0, "1", "0")
            End If
        
        
        Next nRow
    
        .EventEnabled(EventButtonClicked) = True
    End With

End Sub

Private Sub cmdAllCheck_Click()
    Dim vText   As Variant
    Dim nRow    As Long
    
    chkTeamSelect(1).Value = xtpUnchecked
    chkTeamSelect(2).Value = xtpUnchecked
    chkSelect(0).Value = xtpUnchecked
    chkSelect(1).Value = xtpUnchecked
    chkSelect(2).Value = xtpUnchecked
    chkSelect(3).Value = xtpUnchecked
    chkSelect(4).Value = xtpUnchecked
    chkSelect2(0).Value = xtpUnchecked
    
    m_MasterSelCodeList = ""
    
    With spdView
        '.EventEnabled(EventButtonClicked) = False
        
        For nRow = 1 To .MaxRows
            .SetText 2, nRow, CVar(IIf(cmdAllCheck.Caption = "선택", "1", "0"))
            If cmdAllCheck.Caption = "선택" Then
                .GetText 1, nRow, vText
                m_MasterSelCodeList = m_MasterSelCodeList & Mid(CStr(vText), 2, 4) & " "
            End If
        
        Next nRow
        '.EventEnabled(EventButtonClicked) = True
    End With
    
    cmdAllCheck.Caption = IIf(cmdAllCheck.Caption = "선택", "취소", "선택")
            
'    If cmdAllCheck.Caption = "선택" Then
'        cmdAllCheck.Caption = "취소"
'        chkSelect(0).Value = xtpChecked
'        chkSelect(1).Value = xtpChecked
'        chkSelect(2).Value = xtpChecked
'        chkSelect(3).Value = xtpChecked
'        chkSelect(4).Value = xtpChecked
'        chkSelect2(0).Value = xtpChecked
'
'    Else
'        cmdAllCheck.Caption = "선택"
'        chkSelect(0).Value = xtpUnchecked
'        chkSelect(1).Value = xtpUnchecked
'        chkSelect(2).Value = xtpUnchecked
'        chkSelect(3).Value = xtpUnchecked
'        chkSelect(4).Value = xtpUnchecked
'        chkSelect2(0).Value = xtpUnchecked
'
'    End If
        
        
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 0: Call Data_Display           ' 조회
        Case 1:                ' 신규
        Case 2: Call DataSave  ' 저장
        Case 3:            ' 삭제
        Case 4:            ' 취소
        Case 5:            ' 인쇄
        Case 6: Call Export_Excel(P_00000.cdgExcel, spdView2)      ' 엑셀
        Case 7: Unload Me           ' 종료
        Case 8: Call StoreFind      ' 매장찾기
        
        Case Else
            '
    End Select

End Sub

Private Sub Command1_Click()
    SSPanel1.ZOrder 0
    SSPanel1.Visible = Not SSPanel1.Visible
End Sub

Private Sub Form_Activate()
    Dim nRow    As Long
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"

    Call SubBottonEnable(cmdBtn, "10100011")
    
    If P_04027_Flag = False Then
        Screen.MousePointer = vbHourglass
        dtInput(0).Value = Date
        dtInput(1).Value = Date
        
        m_MasterSelCodeList = ""
        
        With spdView2
            
            .Row = -999
            .Col = 5:  .Text = Right(Format(dtInput(0).Value, "yyyy"), 4) & "년"
            .Col = 6:  .Text = Right(Format(DateAdd("yyyy", -1, dtInput(0).Value), "yyyy"), 4) & "년"
            .Col = 9:  .Text = Right(Format(dtInput(0).Value, "yyyy"), 4) & "년"
            .Col = 10:  .Text = Right(Format(DateAdd("yyyy", -1, dtInput(0).Value), "yyyy"), 4) & "년"
            
            
            .Col = 1: .ColMerge = MergeRestricted
            .Col = 2: .ColMerge = MergeRestricted
            
            
        End With
            
        ReDim sValue(1)
        
        sValue(0) = "0"
        sValue(1) = IIf(Store.Code = "1000", "%", Store.Code & "%")
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_04027_00", sValue(), Err_Num, Err_Dec)
        
        spdView.MaxRows = RS01.RecordCount
        Call spdDisplay(RS01)
        
        For nRow = 1 To spdView.DataRowCnt
            spdView.Row = nRow: spdView.Col = 5
            If InStr(spdView.Text, "폐점") > 0 Then
                spdView.Col = -1
                spdView.ForeColor = vbRed
            End If
        Next nRow
        
        
        
        P_04027_Flag = True
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub spdDisplay(Rs As ADODB.Recordset)

    Call fpSpread_Display(spdView, Rs)
    
End Sub

Private Sub Form_Load()
        'Init the User Sort
       spdView2.UserColAction = UserColActionSort

End Sub

Private Sub spdView_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim vText   As Variant
    Dim nRow    As Long
    
    spdView.EventEnabled(EventButtonClicked) = False
    
    
    If Row = spdView.ActiveRow Then
                    
        ReDim sValue(2)
        
        If Col = 2 Then
            spdView.Row = spdView.ActiveRow
            spdView.Col = Col
            If spdView.Value = False Then
                spdView.Col = 2
                spdView.Text = ""
            
                ' 선택 내용이 지사일 경우 해당 체인점을 모두 선택 시킨다.
                spdView.Col = 1
                sValue(2) = Mid(spdView.Text, 2, 6)
                If Mid(sValue(2), 5, 1) = "]" Then
                    
                    sValue(2) = Left(sValue(2), 4)
                    For nRow = 1 To spdView.MaxRows
                        spdView.Row = nRow
                        spdView.Col = 3
                        If spdView.Text = sValue(2) Then
                            spdView.Col = 2
                            spdView.Value = "0"
                        
                        End If
                    Next nRow
                End If
        
            Else

                
                spdView.Row = Row
                spdView.Col = 2: spdView.Text = "1"
                
                
                ' 선택 내용이 지사일 경우 해당 체인점을 모두 선택 시킨다.
                spdView.Col = 1: sValue(2) = Mid(spdView.Text, 2, 6)
                If Mid(sValue(2), 5, 1) = "]" Then
                    sValue(2) = Left(sValue(2), 4)
                    
                    ' 해당 지사의 매장일 경우
                    For nRow = 1 To spdView.MaxRows
                        spdView.Row = nRow
                        spdView.Col = 3
                        If spdView.Text = sValue(2) Then
                            spdView.Col = 2: spdView.Value = "1"
                        End If
                    Next nRow
                
                End If
            End If
        End If
    End If
    
    m_MasterSelCodeList = ""
    
    With spdView
        
        For nRow = 1 To .MaxRows
            .GetText 2, nRow, vText
            
            If CStr(vText) = "1" Then
                    
                ' 선택 내용이 지사일 경우 해당 체인점을 모두 선택 시킨다.
                .GetText 1, nRow, vText
                sValue(2) = Mid(vText, 2, 6)
                If Mid(sValue(2), 5, 1) = "]" Then
                    sValue(2) = Left(sValue(2), 4)
                    
                    m_MasterSelCodeList = m_MasterSelCodeList & Mid(CStr(vText), 2, 4) & " "
                End If
            End If
        
        Next nRow
    End With
    
    
    spdView.EventEnabled(EventButtonClicked) = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cmdBtn(0).Enabled = False
    cmdBtn(1).Enabled = False
    cmdBtn(2).Enabled = False
    cmdBtn(3).Enabled = False
    cmdBtn(4).Enabled = False
    cmdBtn(5).Enabled = False
    cmdBtn(6).Enabled = False
    
    P_04027_Flag = False
End Sub

 
Public Sub DataAdd()

End Sub
Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim nRow            As Long
    Dim SSQL2           As String
    Dim dTemp(2)        As Double
    Dim dTempTotal(2)   As Double
    Dim sKeyCode(1)     As String
    
    With spdView2
        
        .Row = -999
        .Col = 5:   .Text = Right(Format(dtInput(0).Value, "yyyy"), 4) & "년"
        .Col = 6:  .Text = Right(Format(DateAdd("yyyy", -1, dtInput(0).Value), "yyyy"), 4) & "년"
        .Col = 9:  .Text = Right(Format(dtInput(0).Value, "yyyy"), 4) & "년"
        .Col = 10:  .Text = Right(Format(DateAdd("yyyy", -1, dtInput(0).Value), "yyyy"), 4) & "년"
        
    End With
    
    sKeyCode(0) = "":   sKeyCode(1) = ""
    dTemp(0) = 0:       dTemp(1) = 0:       dTemp(2) = 0
    dTempTotal(0) = 0:  dTempTotal(1) = 0:  dTempTotal(2) = 0
    
    ReDim sValue(2)
    
    SSQL2 = ""
    '-------------------------------------------------------------------------------
    ' 선택된 체인점의 내용만을 구해서 쿼리한다.
    For nRow = 0 To spdView.MaxRows
        spdView.Col = 2:  spdView.Row = nRow
        ' 매장코드만 적용한다.
        If spdView.Text = "1" Then
            spdView.Col = 1
            If IsNumeric(Mid(spdView.Text, 6, 1)) Then '<- 매장만, 지사제외
                        
                spdView.Col = 1
                SSQL2 = SSQL2 & Mid(spdView.Text, 2, 6) & ","
            End If
        End If
    Next nRow
    
    SSQL2 = Trim(SSQL2)
    If Len(SSQL2) > 3 Then
        sValue(0) = Mid(SSQL2, 1, Len(SSQL2) - 1)           ' 마지막 ,을 제거한다.
    End If
        '-------------------------------------------------------------------------------
    
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    If sValue(0) = "" Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_04027_01", sValue(), Err_Num, Err_Dec)
    
    With spdView2
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            
            ' 최초 지사 코드 등록
            If .MaxRows = 0 Then
                sKeyCode(0) = Trim(RS01!지사코드 & "")
                sKeyCode(1) = Trim(RS01!지사명 & "")
            End If
            
            ' 지사별 합계 처리
            If sKeyCode(0) <> Trim(RS01!지사코드 & "") And chkTotal.Value = xtpChecked Then
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                
                .Col = 1:  .Text = sKeyCode(0)
                .Col = 2:  .Text = sKeyCode(1)
                .Col = 4:  .Text = "합   계"
                
                .Col = 5:  .Text = dTemp(0)
                .Col = 6:  .Text = dTemp(1)
                If dTemp(1) > 0 Then
                    .Col = 7:  .Text = Round((dTemp(0) - dTemp(1)) / dTemp(1) * 100, 1)
                    
                    .Col = -1
                    .BackColor = &HC0FFC0
                End If
                
                dTempTotal(0) = dTempTotal(0) + dTemp(0)
                dTempTotal(1) = dTempTotal(1) + dTemp(1)
                dTempTotal(2) = dTempTotal(2) + dTemp(2)
                
                sKeyCode(0) = Trim(RS01!지사코드 & "")
                sKeyCode(1) = Trim(RS01!지사명 & "")
                dTemp(0) = 0:   dTemp(1) = 0:   dTemp(2) = 0
            
            End If
            
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = Trim(RS01!지사코드 & "")
            .Col = 2:  .Text = Trim(RS01!지사명 & "")
            .Col = 3:  .Text = RS01!가맹점코드 & ""
            .Col = 4:  .Text = RS01!가맹점명 & ""
            .Col = 5:  .Text = RS01!매출금액 & ""
            .Col = 6:  .Text = RS01!매출금액2 & ""
            .Col = 7:  .Text = RS01!비고메모 & ""
            
            
            dTemp(0) = dTemp(0) + CDbl(RS01!매출금액 & "")
            dTemp(1) = dTemp(1) + CDbl(RS01!매출금액2 & "")
            
            If IsNumeric(RS01!매출금액 & "") = True And IsNumeric(RS01!매출금액2 & "") Then
                If Val(Val(RS01!매출금액2 & "")) > 0 Then
                    .Col = 7:  .Text = Round((Val(RS01!매출금액 & "") - Val(RS01!매출금액2 & "")) / Val(RS01!매출금액2 & "") * 100, 1)
                    
                    .ForeColor = IIf(Val(.Text) < 0, vbRed, vbBlack)
                End If
            End If
            
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing

            
        ' 지사별 합계 처리
        If chkTotal.Value = xtpChecked Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = sKeyCode(0)
            .Col = 2:  .Text = sKeyCode(1)
            .Col = 4:  .Text = "합   계"
            
            .Col = 5:  .Text = dTemp(0)
            .Col = 6:  .Text = dTemp(1)
            If dTemp(1) > 0 Then
                .Col = 7:  .Text = Round((dTemp(0) - dTemp(1)) / dTemp(1) * 100, 1)
                
                .Col = -1
                .BackColor = &HC0FFC0
            End If
        End If
        
        dTempTotal(0) = dTempTotal(0) + dTemp(0)
        dTempTotal(1) = dTempTotal(1) + dTemp(1)
        dTempTotal(2) = dTempTotal(2) + dTemp(2)
        txtNum(0).Value = dTempTotal(0)
        txtNum(1).Value = dTempTotal(1)
        
        ' 합계 전년 대비 매출 신장률
        If IsNumeric(txtNum(0).Value) = True And IsNumeric(txtNum(1).Value) Then
            If Val(txtNum(1).Value) > 0 Then
                txtNum(2).Value = Round((Val(txtNum(0).Value) - Val(txtNum(1).Value)) / Val(txtNum(1).Value) * 100, 1)
            
                txtNum(2).ForeColor = IIf(Val(txtNum(2).Value) < 0, vbRed, vbBlack)
            
            End If
            
        End If
        
        
        '.Row = 1:   .Col = -1:  oldColor = .BackColor
            
        ' 누락된 요일을 설정한다.
'        Call DateCheckAdd(spdView, Format(dtInput(0).Value, "YYYY-MM-DD"), Format(dtInput(1).Value, "YYYY-MM-DD"))
        .Redraw = True
    End With
        
    ' 객수 표시
    Call Data_Display_객수
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrRtn:
    Screen.MousePointer = vbDefault
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Data_Display_객수()
    On Error GoTo ErrRtn

    Dim nRow As Long
    Dim vText   As Variant
    Dim dTemp(1)    As Double
    Dim sKeyCode    As String
    
    If chkSelect3.Value <> xtpChecked Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    ReDim sValue(2)

    ' 당해 년도
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    With spdView2
       For nRow = 1 To .DataRowCnt
       
           .GetText 3, nRow, vText:     sValue(0) = CStr(vText)
           
           If sValue(0) <> "" Then
           
                Set RS01 = New ADODB.Recordset
                Set RS01 = ExecPro("SP_04027_03", sValue(), Err_Num, Err_Dec)
               
                If Not RS01.EOF Then .SetText 9, nRow, CVar(RS01!객수1 & "")
                   
            End If
        Next nRow
    End With
    
    ' 전년도 (DB가 분산되어 있어서 이렇게함.. 한번에 구해올수 있지만 시간이 없어서 ㅠㅠ
    sValue(1) = Format(DateAdd("yyyy", -1, dtInput(0).Value), "YYYY-MM-DD")
    sValue(2) = Format(DateAdd("yyyy", -1, dtInput(1).Value), "YYYY-MM-DD")
    
    With spdView2
       For nRow = 1 To .DataRowCnt
       
           .GetText 3, nRow, vText:     sValue(0) = CStr(vText)
           
           If sValue(0) <> "" Then
           
                Set RS01 = New ADODB.Recordset
                Set RS01 = ExecPro("SP_04027_03", sValue(), Err_Num, Err_Dec)
               
                If Not RS01.EOF Then .SetText 10, nRow, CVar(RS01!객수1 & "")
                   
            End If
        Next nRow
        
        sKeyCode = ""
        dTemp(0) = 0:   dTemp(1) = 0
        
        For nRow = 1 To .DataRowCnt
           .GetText 1, nRow, vText
           If nRow = 1 Then sKeyCode = CStr(vText)
           
           ' 지사가 변경될 경우 합계를 출력한다.
           If sKeyCode <> CStr(vText) And chkTotal.Value = xtpChecked Then
                .SetText 9, nRow - 1, CVar(dTemp(0))
                .SetText 10, nRow - 1, CVar(dTemp(1))
                
                sKeyCode = CStr(vText)
                dTemp(0) = 0:   dTemp(1) = 0
            End If
            
           .GetText 9, nRow, vText:     dTemp(0) = dTemp(0) + CDbl(IIf(IsNumeric(vText) = False, 0, vText))
           .GetText 10, nRow, vText:    dTemp(1) = dTemp(1) + CDbl(IIf(IsNumeric(vText) = False, 0, vText))
             
        Next nRow
        
        .SetText 9, nRow - 1, CVar(dTemp(0))
        .SetText 10, nRow - 1, CVar(dTemp(1))
        
'        .Col = 9:  .Formula = "SUM(I1:I" & .MaxRows - 1 & ")"
'        .Col = 10:  .Formula = "SUM(J1:J" & .MaxRows - 1 & ")"
    End With
    
    

    With spdView2
       For nRow = 1 To .DataRowCnt
       
           .GetText 9, nRow, vText:     sValue(0) = CStr(vText)
           .GetText 10, nRow, vText:     sValue(1) = CStr(vText)
           
            If IsNumeric(sValue(0)) = True And IsNumeric(sValue(1)) Then
                If Val(sValue(1)) > 0 Then
                    .SetText 11, nRow, Round((Val(sValue(0)) - Val(sValue(1))) / Val(sValue(1)) * 100, 1)
                    
                    .ForeColor = IIf(Val(.Text) < 0, vbRed, vbBlack)
                End If
                
            End If
        Next nRow
    End With
                   
 
        
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrRtn:
    Screen.MousePointer = vbDefault
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub




Public Sub DataSave()
    Dim iRow       As Long
    Dim sValue()    As String
    
    On Error GoTo ErrRtn
    
 
    
    For iRow = 1 To spdView2.DataRowCnt
        spdView2.Row = iRow
        spdView2.Col = 1
       
        ReDim sValue(3)
        
        sValue(1) = Format(dtInput(0).Value, "yyyy-MM-dd")  ' 2 시작일자
        sValue(2) = Format(dtInput(1).Value, "yyyy-MM-dd")  ' 3 종료일자
        
        spdView2.Row = iRow
        spdView2.Col = 3:  sValue(0) = Trim(spdView2.Text) & ""    ' 1. 가맹점코드
        spdView2.Col = 8:  sValue(3) = Trim(spdView2.Text) & ""    ' 4. 비고메모
        
        If Trim(sValue(0)) <> "" Then
            Call ExecPro("SP_04027_02_INS", sValue(), Err_Num, Err_Dec)
        
            If Err_Num <> 0 Then
                MsgBox "[" & Err_Num & "] " & Err_Dec
                
                Exit Sub
            End If
        End If
    Next iRow
    
    MsgBox "저장 완료", vbInformation, "확인"
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub StoreFind()
    
    cmdBtn(8).Tag = ""
    
    Set P_01001_A1.m_FormObj = Me
    P_01001_A1.Show vbModal
    
    
    ' 조회에서 매장 코드가 설정된 경우 해당 매장이 선택 되도록 한다.
    DoEvents
    If Trim(cmdBtn(8).Tag) = "" Then Exit Sub
    
    With spdView
        .Redraw = False
        Rtn = .SearchCol(1, 1, .MaxRows, Trim(cmdBtn(8).Tag), SearchFlagsPartialMatch)
        .Row = .MaxRows
        .Action = ActionActiveCell
        .Row = Rtn
        .Action = ActionActiveCell
        .Redraw = True
        DoEvents
        
        .SetText 2, Rtn, CVar("1")
        
        
    End With

End Sub
