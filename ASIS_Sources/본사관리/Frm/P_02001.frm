VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_02001 
   Caption         =   "가맹점 접수현황"
   ClientHeight    =   9750
   ClientLeft      =   -20250
   ClientTop       =   3990
   ClientWidth     =   16860
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_02001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9750
   ScaleWidth      =   16860
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9750
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16860
      _ExtentX        =   29739
      _ExtentY        =   17198
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_02001.frx":058A
      Begin Threed.SSPanel SSPanel1 
         Height          =   810
         Left            =   15
         TabIndex        =   2
         Top             =   8925
         Width           =   16830
         _ExtentX        =   29686
         _ExtentY        =   1429
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   345
            Index           =   0
            Left            =   1245
            TabIndex        =   11
            Top             =   60
            Width           =   1425
            _Version        =   262145
            _ExtentX        =   2514
            _ExtentY        =   609
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
            StartText.y     =   5
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   13
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
            Height          =   345
            Index           =   1
            Left            =   60
            TabIndex        =   3
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   609
            _Version        =   262144
            Caption         =   "총접수량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   345
            Index           =   3
            Left            =   3630
            TabIndex        =   4
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   609
            _Version        =   262144
            Caption         =   "재세탁수량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   345
            Index           =   4
            Left            =   7185
            TabIndex        =   5
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   609
            _Version        =   262144
            Caption         =   "반품수량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   345
            Index           =   5
            Left            =   10650
            TabIndex        =   6
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   609
            _Version        =   262144
            Caption         =   "수선수량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   345
            Index           =   6
            Left            =   60
            TabIndex        =   7
            Top             =   450
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   609
            _Version        =   262144
            Caption         =   "매출금액"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   345
            Index           =   7
            Left            =   3630
            TabIndex        =   8
            Top             =   435
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   609
            _Version        =   262144
            Caption         =   "지    사"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   345
            Index           =   8
            Left            =   7185
            TabIndex        =   9
            Top             =   435
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   609
            _Version        =   262144
            Caption         =   "가 맹 점"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   345
            Index           =   9
            Left            =   10650
            TabIndex        =   10
            Top             =   435
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   609
            _Version        =   262144
            Caption         =   "수선비용"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   345
            Index           =   1
            Left            =   4815
            TabIndex        =   12
            Top             =   60
            Width           =   1425
            _Version        =   262145
            _ExtentX        =   2514
            _ExtentY        =   609
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
            StartText.y     =   5
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   13
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
            Height          =   345
            Index           =   2
            Left            =   8370
            TabIndex        =   13
            Top             =   60
            Width           =   1425
            _Version        =   262145
            _ExtentX        =   2514
            _ExtentY        =   609
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
            StartText.y     =   5
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   13
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
            Height          =   345
            Index           =   3
            Left            =   11835
            TabIndex        =   14
            Top             =   60
            Width           =   1425
            _Version        =   262145
            _ExtentX        =   2514
            _ExtentY        =   609
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
            StartText.y     =   5
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   13
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
            Height          =   345
            Index           =   4
            Left            =   1245
            TabIndex        =   15
            Top             =   435
            Width           =   1425
            _Version        =   262145
            _ExtentX        =   2514
            _ExtentY        =   609
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
            StartText.y     =   5
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   13
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
            Height          =   345
            Index           =   5
            Left            =   4815
            TabIndex        =   16
            Top             =   435
            Width           =   1425
            _Version        =   262145
            _ExtentX        =   2514
            _ExtentY        =   609
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
            StartText.y     =   5
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   13
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
            Height          =   345
            Index           =   6
            Left            =   8370
            TabIndex        =   17
            Top             =   435
            Width           =   1425
            _Version        =   262145
            _ExtentX        =   2514
            _ExtentY        =   609
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
            StartText.y     =   5
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   13
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
            Height          =   345
            Index           =   7
            Left            =   11835
            TabIndex        =   18
            Top             =   435
            Width           =   1425
            _Version        =   262145
            _ExtentX        =   2514
            _ExtentY        =   609
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
            StartText.y     =   5
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   13
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
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   7575
         Left            =   15
         TabIndex        =   1
         Top             =   1335
         Width           =   16830
         _Version        =   524288
         _ExtentX        =   29686
         _ExtentY        =   13361
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
         MaxCols         =   25
         Protect         =   0   'False
         SpreadDesigner  =   "P_02001.frx":063C
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   19
         Top             =   540
         Width           =   16830
         _ExtentX        =   29686
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboGubun 
            Height          =   315
            ItemData        =   "P_02001.frx":1241
            Left            =   6975
            List            =   "P_02001.frx":1243
            Style           =   2  '드롭다운 목록
            TabIndex        =   39
            Top             =   405
            Width           =   1425
         End
         Begin VB.TextBox txtFind 
            Height          =   315
            Left            =   8430
            TabIndex        =   38
            Top             =   405
            Width           =   4575
         End
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1245
            Locked          =   -1  'True
            TabIndex        =   34
            Text            =   "cboOffice"
            Top             =   60
            Width           =   3420
         End
         Begin VB.ComboBox cboInput 
            Height          =   315
            Left            =   1245
            Style           =   2  '드롭다운 목록
            TabIndex        =   20
            Top             =   405
            Width           =   3420
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   6975
            TabIndex        =   21
            Top             =   60
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            Format          =   57212928
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   22
            Top             =   405
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "가맹점명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   5790
            TabIndex        =   23
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "접수일자"
            BorderWidth     =   0
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   10
            Left            =   60
            TabIndex        =   35
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "지    사"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   10155
            TabIndex        =   36
            Top             =   60
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            Format          =   57212928
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   11
            Left            =   5790
            TabIndex        =   40
            Top             =   405
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "검색조건"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton cmdRefresh 
            Height          =   330
            Left            =   4710
            TabIndex        =   41
            Top             =   390
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   582
            _StockProps     =   79
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_02001.frx":1245
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
            Left            =   9840
            TabIndex        =   37
            Top             =   120
            Width           =   300
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   24
         Top             =   15
         Width           =   9225
         _ExtentX        =   16272
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
         Caption         =   " 가맹점 접수현황 (P_02001)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_02001.frx":17DF
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Left            =   9255
         TabIndex        =   25
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
         PictureBackground=   "P_02001.frx":19E1
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   26
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
            Picture         =   "P_02001.frx":1BE3
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   27
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
            Picture         =   "P_02001.frx":217D
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   28
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
            Picture         =   "P_02001.frx":2717
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   29
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
            Picture         =   "P_02001.frx":2CB1
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   30
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
            Picture         =   "P_02001.frx":324B
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   31
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
            Picture         =   "P_02001.frx":37E5
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   32
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
            Picture         =   "P_02001.frx":3D7F
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   33
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
            Picture         =   "P_02001.frx":4319
         End
      End
   End
End
Attribute VB_Name = "P_02001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cboInput_Click()
    Call Data_Display
End Sub

Private Sub cboOffice_Click()
    If cboOffice.ListIndex < 0 Then Exit Sub
    
    '-----------------------------------------------------------------
    '
    '-----------------------------------------------------------------
    cboInput.Clear
    spdView.MaxRows = 0
    
    ReDim sValue(2)
    
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_01001_00", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_01001_00", sValue(), Err_Num, Err_Dec)
    End If

    With cboInput
        Do Until RS01.EOF
            .AddItem "[" & RS01!가맹점코드 & "] " & RS01!가맹점명
        
            'If (RS01!종료일자 = "") Or (RS01!종료일자 = "2099-12-31") Then '가맹점이 현 지사에서 관리중...
            '    .AddItem "[" & RS01!가맹점코드 & "] " & RS01!가맹점명
            'End If
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        If .ListCount > 0 Then .ListIndex = 0
    End With
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
        Case 5: 'Call DataPrint      ' 인쇄
        Case 6: Call Export_Excel(P_00000.cdgExcel, spdView)      ' 엑셀
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

Private Sub cmdRefresh_Click()
    cboOffice_Click
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

'    If P_02001_Flag = False Then
'        dtInput.Value = Date
'
'        Call AgencyComboAdd(cboInput)
'
'        ReDim sValue(2)
'
'        sValue(0) = "1"
'
'        Set RS01 = New ADODB.Recordset
'        Set RS01 = ExecPro("SP_02001_00", sValue(), Err_Num, Err_Dec)
'
'        spdView.MaxCols = RS01.Fields.Count
'        spdView.MaxRows = RS01.RecordCount
'
'        Call spdDisplay(RS01)
'        Call GetColWidth(REG_App, Me.Name, spdView)
'
'        P_02001_Flag = True
'    End If
End Sub

'Private Sub spdDisplay(Rs As ADODB.Recordset)
'    Call fpSpread_Display(spdView, Rs)
'End Sub
'
'Private Sub spdDisplay2(Rs As ADODB.Recordset)
'    Call fpSpread_Display(spdView(1), Rs)
'End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    With spdView
        .MaxRows = 0
        .RowHeight(-1) = 14
        
        .Col = 8:  .ColMerge = MergeAlways
        .Col = 9:  .ColMerge = MergeRestricted
        .Col = 10: .ColMerge = MergeRestricted
        .Col = 11: .ColMerge = MergeRestricted
        
        .ColsFrozen = 5
        
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
    
    dtInput(0).Value = Date
    dtInput(1).Value = Date

    '
    Call Get_지사리스트(cboOffice)
    
    Dim i As Integer
    
    With cboOffice
        
        For i = 0 To .ListCount - 1
            If Mid(.List(i), 2, 4) = HeadOffice Then
                .ListIndex = i
                
                Exit For
            End If
        Next i
    End With

    With cboGubun
        .Clear
        .AddItem "성명"
        .AddItem "전화번호"
        .AddItem "고객코드"
        
        .ListIndex = 0
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_02001_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn
    
    Dim i           As Long
    
    If cboInput.ListIndex < 0 Then
        MsgBox "가맹점을 선택하세요.", vbInformation, "확인"
        Exit Sub
    End If
    
    For i = 0 To 7
        txtNum(i).Value = 0
    Next i
    
    '----------------------------------------------------------------
    ' SP_02001_00
    '----------------------------------------------------------------
    ReDim sValue(5)
    
    sValue(0) = Mid(cboInput.Text, 2, 6)               '
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD") '
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD") '
    sValue(3) = cboGubun.Text                          '
    sValue(4) = Trim(txtFind.Text)                     '
    
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_02001_00", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_02001_00", sValue(), Err_Num, Err_Dec)
    End If
    
    With spdView
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = Format(RS01!택번호, "000-00-0000") & ""        ' 1
            .Col = 2:  .Text = RS01!의류코드 & ""       ' 6
            .Col = 3:  .Text = RS01!의류명 & ""         ' 7
            .Col = 4:  .Text = RS01!색상 & ""           ' 8
            .Col = 5:  .Text = RS01!내용 & ""           ' 9
            .Col = 6:  .Text = RS01!상표 & ""           '10
            .Col = 7:  .Text = RS01!금액 & ""           '11
            
            .Col = 8:  .Text = RS01!접수번호 & ""       ' 2
            .Col = 9:  .Text = RS01!성명 & ""           ' 3
            .Col = 10: .Text = RS01!전화번호 & ""       ' 4
            .Col = 11: .Text = RS01!휴대전화 & ""         ' 5
            
            .Col = 12: .Text = RS01!접수일자 & ""       '12
            .Col = 13: .Text = RS01!접수시간 & ""       '12
            
            .Col = 14: .Text = RS01!가맹점출고일자 & "" '13
            .Col = 15: .Text = RS01!가맹점입고일자 & "" '14
            .Col = 16: .Text = RS01!지사입고일자 & ""   '15
            .Col = 17: .Text = RS01!지사출고일자 & ""   '16
            .Col = 18: .Text = RS01!출고일자 & ""       '17
            .Col = 19: .Text = RS01!출고시간 & ""       '12
            
            .Col = 20: .Text = RS01!부모택번호 & ""     '18
            .Col = 21: .Text = RS01!반품환불일자 & ""   '19
            .Col = 22: .Text = RS01!세탁환불일자 & ""   '20
            .Col = 23: .Text = RS01!판매취소일자 & ""   '21
            .Col = 24: .Text = RS01!환불사유 & ""       '22
            .Col = 25: .Text = RS01!오점내용 & ""       '23
            
            txtNum(0).Value = txtNum(0).Value + 1
            
            '재세탁
            i = InStr(1, RS01!내용, "재", vbTextCompare)
                
            If i > 0 Then
                txtNum(1).Value = txtNum(1).Value + 1
            End If
            
            '반품
            If (RS01!반품환불일자 <> "") Or (RS01!세탁환불일자 <> "") Or (RS01!판매취소일자 <> "") Then
                txtNum(2).Value = txtNum(2).Value + 1
            End If
            
            txtNum(4).Value = txtNum(4).Value + RS01!금액
            txtNum(6).Value = txtNum(6).Value + (RS01!금액 * RS01!세탁마진 / 100)
            txtNum(5).Value = txtNum(4).Value - txtNum(6).Value
            
            '수선
            i = InStr(1, RS01!내용, "수", vbTextCompare)
                
            If i > 0 Then
                txtNum(3).Value = txtNum(3).Value + 1
            End If
            
            txtNum(7).Value = txtNum(7).Value + RS01!수선금액
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub DataPrint()
'    Dim ReportFP As String
'    Dim ReportFile As String
'
'    ReportFP = GetIniStr("REPORT", "FilePath", "", m_iniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    Call PrintDesc
'
'    Dim i As Integer
'    For i = 0 To 30
'        P_00000.crPrint.Formulas(i) = ""
'    Next
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'    P_00000.crPrint.Formulas(0) = "입고일자 = '" & Format(dtInput.Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(1) = "대리점명 = '" & cboInput.Text & "'"
'
'    P_00000.crPrint.Formulas(2) = "총점수 = '" & txtNum(0).Text & "'"
'    P_00000.crPrint.Formulas(3) = "재세탁 = '" & txtNum(1).Text & "'"
'    P_00000.crPrint.Formulas(4) = "반품 = '" & txtNum(2).Text & "'"
'    P_00000.crPrint.Formulas(5) = "수선 = '" & txtNum(3).Text & "'"
'    P_00000.crPrint.Formulas(6) = "매출액 = '" & txtNum(4).Text & "'"
'    P_00000.crPrint.Formulas(7) = "본사 = '" & txtNum(5).Text & "'"
'    P_00000.crPrint.Formulas(8) = "대리점 = '" & txtNum(6).Text & "'"
'    P_00000.crPrint.Formulas(9) = "수선비용 = '" & txtNum(7).Text & "'"
'
'    Call ReportPrint(ReportFile, "1")
End Sub

Private Sub spdView_Change(ByVal Col As Long, ByVal Row As Long)
'    If Index = 0 Then
'        spdView.Row = Row
'        spdView.Col = 10
'        spdView.Text = "U"
'
'        spdView.Col = -1
'        spdView.BackColor = vbYellow
'    End If
End Sub

Private Sub spdView_DblClick(ByVal Col As Long, ByVal Row As Long)
'    If Index = 1 Then
'        spdView.Row = spdView.ActiveRow
'        spdView(1).Row = spdView(1).ActiveRow
'
'        spdView.Col = 4
'        spdView(1).Col = 1
'        spdView.Text = spdView(1).Text
'
'        spdView.Col = 5
'        spdView(1).Col = 2
'        spdView.Text = spdView(1).Text
'
'        spdView.Col = 6
'        spdView(1).Col = 3
'        spdView.Text = spdView(1).Text
'
'        spdView(1).Visible = False
'    End If
End Sub

Private Sub spdView_KeyDown(KeyCode As Integer, Shift As Integer)
'    If Index = 0 Then
'        If KeyCode = vbKeyReturn Then
'            If spdView.ActiveCol = 4 Then
'                '-------------------------------------------------------------------
'                '
'                '-------------------------------------------------------------------
'                ReDim sValue(3)
'
'                sValue(0) = "0"
'                sValue(1) = Mid(cboInput.Text, 2, 6)                ' 대리점
'                sValue(2) = Format(dtInput.Value, "YYYY-MM-DD")       ' 입고일자
'
'                spdView.Row = spdView.ActiveRow
'                spdView.Col = 4
'                sValue(3) = Trim(spdView.Text) & "%"                   ' 품목코드
'
'                Set RS01 = New ADODB.Recordset
'                Set RS01 = ExecPro("SP_02001_04", sValue(), Err_Num, Err_Dec)
'
'                If Not RS01.EOF Then
'                    If RS01.RecordCount = 1 Then
'                        spdView.Col = 5: spdView.Text = RS01!품목명 & ""
'                        spdView.Col = 6: spdView.Value = RS01!금액 & ""
'                    Else
'                        spdView(1).Visible = True
'
'                        spdView(1).MaxCols = RS01.Fields.Count
'                        spdView(1).MaxRows = RS01.RecordCount
'
'                        'Call spdDisplay2(RS01)
'                        Call fpSpread_Display(spdView(1), RS01)
'                        Call GetColWidth(REG_App, Me.Name & "B", spdView(1))
'                    End If
'                End If
'            End If
'        End If
'    End If
End Sub

Private Sub DataScreen()
'    Dim ReportFP As String
'    Dim ReportFile As String
'
'    ReportFP = GetIniStr("REPORT", "FilePath", "", m_iniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    Call PrintDesc
'
'    Dim i As Integer
'    For i = 0 To 30
'        P_00000.crPrint.Formulas(i) = ""
'    Next
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'    P_00000.crPrint.Formulas(0) = "입고일자 = '" & Format(dtInput.Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(1) = "대리점명 = '" & cboInput.Text & "'"
'
'    P_00000.crPrint.Formulas(2) = "총점수 = '" & txtNum(0).Text & "'"
'    P_00000.crPrint.Formulas(3) = "재세탁 = '" & txtNum(1).Text & "'"
'    P_00000.crPrint.Formulas(4) = "반품 = '" & txtNum(2).Text & "'"
'    P_00000.crPrint.Formulas(5) = "수선 = '" & txtNum(3).Text & "'"
'    P_00000.crPrint.Formulas(6) = "매출액 = '" & txtNum(4).Text & "'"
'    P_00000.crPrint.Formulas(7) = "본사 = '" & txtNum(5).Text & "'"
'    P_00000.crPrint.Formulas(8) = "대리점 = '" & txtNum(6).Text & "'"
'    P_00000.crPrint.Formulas(9) = "수선비용 = '" & txtNum(7).Text & "'"
'
'    Call ReportPrint(ReportFile, "2")
End Sub

Private Sub PrintDesc()
    Dim i As Integer
    Dim TempText As String
    Dim TempFP As String
    Dim TempFile As String
    
    TempFP = GetIniStr("REPORT", "TempPath", "", m_iniFile)
    TempFile = TempFP & "\Temp.txt"
    
    Open TempFile For Output As #1
    
    TempText = ""
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        
        spdView.Col = 1
        TempText = LeftH(spdView.Text & Space(7), 7)
        spdView.Col = 2
        TempText = TempText & LeftH(spdView.Text & Space(11), 11)
        spdView.Col = 3
        TempText = TempText & LeftH(spdView.Text & Space(10), 10)
        spdView.Col = 4
        TempText = TempText & LeftH(spdView.Text & Space(5), 5)
        spdView.Col = 5
        TempText = TempText & LeftH(spdView.Text & Space(14), 14)
        spdView.Col = 6
        TempText = TempText & RightH(Space(10) & spdView.Text, 10) & Space(2)
        spdView.Col = 7
        TempText = TempText & LeftH(spdView.Text & Space(8), 8)
        spdView.Col = 8
        TempText = TempText & LeftH(spdView.Text & Space(8), 8)
        spdView.Col = 9
        TempText = TempText & LeftH(spdView.Text & Space(10), 10)
        
        Print #1, TempText
    Next i
    
    Close #1
End Sub

Private Sub DataAdd()
    spdView.MaxRows = spdView.MaxRows + 1
    
    spdView.Row = spdView.MaxRows
    spdView.Col = 1
    spdView.Action = ActionActiveCell
    
    spdView.SetFocus
End Sub

Private Sub DataSave()
'    Dim i As Integer
'
'    For i = 1 To spdView.MaxRows
'        spdView.Row = i
'        spdView.Col = 10
'        If spdView.Text = "U" Then
'            ReDim sValue(11)
'
'            sValue(0) = Format(dtInput.Value, "YYYY-MM-DD")
'            sValue(1) = Mid(cboInput.Text, 2, 6)
'
'            spdView.Col = 1:  sValue(2) = spdView.Value
'            spdView.Col = 2:  sValue(3) = Mid(spdView.Value, 1, 4)
'            spdView.Col = 2:  sValue(4) = Mid(spdView.Value, 5, 4)
'            spdView.Col = 3:  sValue(5) = spdView.Text
'            spdView.Col = 4:  sValue(6) = spdView.Text
'            spdView.Col = 7:  sValue(7) = spdView.Text
'            spdView.Col = 8:  sValue(8) = spdView.Text
'            spdView.Col = 6:  sValue(9) = spdView.Value
'            spdView.Col = 9:  sValue(10) = spdView.Text
'            spdView.Col = 10: sValue(11) = spdView.Text
'
'            Call ExecPro("SP_02001_03", sValue(), Err_Num, Err_Dec)
'
'            If Err_Num <> 0 Then
'                MsgBox "[" & Err_Num & "] " & Err_Dec
'            End If
'        End If
'    Next i
End Sub

Private Sub DataCancel()
    Call Data_Display
End Sub

Private Sub ComboAdd()
    Dim sItem As String
    
    sItem = ""
    sItem = "흰색" & Chr(9)
    sItem = sItem & "상아" & Chr(9)
    sItem = sItem & "회색" & Chr(9)
    sItem = sItem & "쥐색" & Chr(9)
    sItem = sItem & "밤색" & Chr(9)
    sItem = sItem & "검정" & Chr(9)
    sItem = sItem & "분홍" & Chr(9)
    sItem = sItem & "주황" & Chr(9)
    sItem = sItem & "빨강" & Chr(9)
    sItem = sItem & "노랑" & Chr(9)
    sItem = sItem & "베지" & Chr(9)
    sItem = sItem & "황토" & Chr(9)
    sItem = sItem & "연두" & Chr(9)
    sItem = sItem & "초록" & Chr(9)
    sItem = sItem & "카키" & Chr(9)
    sItem = sItem & "쑥색" & Chr(9)
    sItem = sItem & "하늘" & Chr(9)
    sItem = sItem & "파랑" & Chr(9)
    sItem = sItem & "곤색" & Chr(9)
    sItem = sItem & "보라" & Chr(9)
    sItem = sItem & "체크" & Chr(9)
    sItem = sItem & "자주" & Chr(9)
    sItem = sItem & "혼합"
    
    spdView.Col = 7
    spdView.TypeComboBoxList = sItem
End Sub

Private Sub DataDelete()
'    If MsgBox("해당되는 데이터를 삭제하시겠습니까?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
'        ReDim sValue(2)
'
'        sValue(0) = Format(dtInput.Value, "YYYY-MM-DD")
'        sValue(1) = Mid(cboInput.Text, 2, 6)
'
'        spdView.Row = spdView.ActiveRow
'        spdView.Col = 1
'        sValue(2) = spdView.Value
'
'        Call ExecPro("SP_02001_05", sValue(), Err_Num, Err_Dec)
'
'        If Err_Num <> 0 Then
'            MsgBox "[" & Err_Num & "] " & Err_Dec
'            Exit Sub
'        Else
'            spdView.Col = -1
'            spdView.Action = ActionDeleteRow
'
'            spdView.MaxRows = spdView.MaxRows - 1
'
'            MsgBox "정상적으로 데이터가 삭제되었습니다.", vbInformation
'        End If
'    End If
End Sub
