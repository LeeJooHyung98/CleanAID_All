VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{6514F5A0-641C-11D2-9FD0-0020AF131A57}#3.0#0"; "fpFlp30.ocx"
Begin VB.Form frm매출 
   Caption         =   "매출 현황"
   ClientHeight    =   10080
   ClientLeft      =   -25215
   ClientTop       =   2145
   ClientWidth     =   15240
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10080
   ScaleWidth      =   15240
   WindowState     =   2  '최대화
   Begin LpADOLib.fpListAdo fpList1 
      Height          =   2295
      Left            =   5040
      TabIndex        =   1
      Top             =   2325
      Visible         =   0   'False
      Width           =   6360
      _Version        =   196608
      _ExtentX        =   11218
      _ExtentY        =   4048
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BackColor       =   -2147483628
      ForeColor       =   -2147483640
      Columns         =   4
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   0
      WrapList        =   0   'False
      WrapWidth       =   0
      SelMax          =   -1
      AutoSearch      =   1
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483627
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   1
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   1
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ScrollHScale    =   2
      ScrollHInc      =   0
      ColsFrozen      =   0
      ScrollBarV      =   0
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   0
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   0   'False
      DataAutoSizeCols=   0
      SearchIgnoreCase=   -1  'True
      ScrollBarH      =   1
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   16777215
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   -1  'True
      ColumnHeaderHeight=   300
      GrpsFrozen      =   0
      BorderGrayAreaColor=   -2147483637
      ExtendRow       =   0
      DataField       =   ""
      DataMember      =   ""
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      ColDesigner     =   "frm매출.frx":0000
   End
   Begin Threed.SSPanel pnlProg 
      Height          =   1215
      Left            =   60
      TabIndex        =   36
      Top             =   1905
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
      Picture         =   "frm매출.frx":02F6
      Caption         =   " "
      BorderWidth     =   0
      BevelOuter      =   1
      BevelInner      =   2
      Alignment       =   4
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   10080
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   17780
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm매출.frx":32C1
      Begin Threed.SSPanel SSPanel1 
         Height          =   1335
         Left            =   15
         TabIndex        =   2
         Top             =   8730
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   2355
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   315
            Index           =   0
            Left            =   1140
            TabIndex        =   3
            Top             =   30
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
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
            Left            =   330
            TabIndex        =   4
            Top             =   30
            Width           =   825
            _ExtentX        =   1455
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
            Caption         =   "접수수량"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Index           =   2
            Left            =   2220
            TabIndex        =   5
            Top             =   630
            Width           =   825
            _ExtentX        =   1455
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
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   1215
            Index           =   0
            Left            =   15
            TabIndex        =   6
            Top             =   30
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   2143
            _Version        =   262144
            CaptionStyle    =   1
            BackColor       =   12648384
            PictureMaskColorSource=   1
            PictureUseMask  =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "매출"
            BevelOuter      =   0
            PictureAlignment=   11
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Index           =   3
            Left            =   330
            TabIndex        =   7
            Top             =   330
            Width           =   825
            _ExtentX        =   1455
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
            Caption         =   "취소수량"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Index           =   7
            Left            =   330
            TabIndex        =   8
            Top             =   630
            Width           =   825
            _ExtentX        =   1455
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
            Caption         =   "출고수량"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   315
            Index           =   1
            Left            =   1140
            TabIndex        =   9
            Top             =   330
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
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
            Left            =   1140
            TabIndex        =   10
            Top             =   630
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
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
            Index           =   6
            Left            =   3030
            TabIndex        =   11
            Top             =   630
            Width           =   1260
            _Version        =   262145
            _ExtentX        =   2222
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
            Index           =   5
            Left            =   3030
            TabIndex        =   12
            Top             =   330
            Width           =   1260
            _Version        =   262145
            _ExtentX        =   2222
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
            Left            =   3030
            TabIndex        =   13
            ToolTipText     =   "순매출액 = 총매출액 - 반품금액 - 포인트사용"
            Top             =   30
            Width           =   1260
            _Version        =   262145
            _ExtentX        =   2222
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
            Index           =   8
            Left            =   2220
            TabIndex        =   14
            Top             =   330
            Width           =   825
            _ExtentX        =   1455
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
            Caption         =   "취소금액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Index           =   9
            Left            =   2220
            TabIndex        =   15
            Top             =   30
            Width           =   825
            _ExtentX        =   1455
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
            Caption         =   "접수금액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Index           =   10
            Left            =   4590
            TabIndex        =   16
            Top             =   30
            Width           =   825
            _ExtentX        =   1455
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
            Caption         =   "합    계"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   1215
            Index           =   11
            Left            =   4275
            TabIndex        =   17
            Top             =   30
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   2143
            _Version        =   262144
            CaptionStyle    =   1
            BackColor       =   12648384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "선불결제"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Index           =   12
            Left            =   4590
            TabIndex        =   18
            Top             =   330
            Width           =   825
            _ExtentX        =   1455
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
            Caption         =   "현    금"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Index           =   13
            Left            =   4590
            TabIndex        =   19
            Top             =   630
            Width           =   825
            _ExtentX        =   1455
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
            Caption         =   "카    드"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   315
            Index           =   8
            Left            =   5400
            TabIndex        =   20
            Top             =   30
            Width           =   1200
            _Version        =   262145
            _ExtentX        =   2117
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
            Index           =   9
            Left            =   5400
            TabIndex        =   21
            Top             =   330
            Width           =   1200
            _Version        =   262145
            _ExtentX        =   2117
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
            Index           =   10
            Left            =   5400
            TabIndex        =   22
            Top             =   630
            Width           =   1200
            _Version        =   262145
            _ExtentX        =   2117
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
            Left            =   6585
            TabIndex        =   23
            Top             =   30
            Width           =   825
            _ExtentX        =   1455
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
            Caption         =   "쿠    폰"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Index           =   5
            Left            =   6585
            TabIndex        =   24
            Top             =   330
            Width           =   825
            _ExtentX        =   1455
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
            Caption         =   "미    수"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   315
            Index           =   12
            Left            =   7395
            TabIndex        =   25
            Top             =   30
            Width           =   1200
            _Version        =   262145
            _ExtentX        =   2117
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
            Index           =   13
            Left            =   7395
            TabIndex        =   26
            Top             =   330
            Width           =   1200
            _Version        =   262145
            _ExtentX        =   2117
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
            Index           =   6
            Left            =   6585
            TabIndex        =   27
            Top             =   630
            Width           =   825
            _ExtentX        =   1455
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
            Caption         =   "반환현금"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   315
            Index           =   14
            Left            =   7395
            TabIndex        =   28
            Top             =   630
            Width           =   1200
            _Version        =   262145
            _ExtentX        =   2117
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
            Index           =   14
            Left            =   8895
            TabIndex        =   29
            Top             =   30
            Width           =   825
            _ExtentX        =   1455
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
            Caption         =   "합    계"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   1215
            Index           =   15
            Left            =   8580
            TabIndex        =   30
            Top             =   30
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   2143
            _Version        =   262144
            CaptionStyle    =   1
            BackColor       =   12648384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "미수결제"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Index           =   16
            Left            =   8895
            TabIndex        =   31
            Top             =   330
            Width           =   825
            _ExtentX        =   1455
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
            Caption         =   "현    금"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Index           =   17
            Left            =   8895
            TabIndex        =   32
            Top             =   630
            Width           =   825
            _ExtentX        =   1455
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
            Caption         =   "카    드"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   315
            Index           =   16
            Left            =   9705
            TabIndex        =   33
            Top             =   30
            Width           =   1200
            _Version        =   262145
            _ExtentX        =   2117
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
            Index           =   17
            Left            =   9705
            TabIndex        =   34
            Top             =   330
            Width           =   1200
            _Version        =   262145
            _ExtentX        =   2117
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
            Index           =   18
            Left            =   9705
            TabIndex        =   35
            Top             =   630
            Width           =   1200
            _Version        =   262145
            _ExtentX        =   2117
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
            Index           =   24
            Left            =   4590
            TabIndex        =   50
            Top             =   930
            Width           =   825
            _ExtentX        =   1455
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
            Caption         =   "마일리지"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   315
            Index           =   11
            Left            =   5400
            TabIndex        =   51
            Top             =   930
            Width           =   1200
            _Version        =   262145
            _ExtentX        =   2117
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
            Index           =   25
            Left            =   6585
            TabIndex        =   52
            Top             =   930
            Width           =   825
            _ExtentX        =   1455
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
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   315
            Index           =   15
            Left            =   7395
            TabIndex        =   53
            Top             =   930
            Width           =   1200
            _Version        =   262145
            _ExtentX        =   2117
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
            Index           =   18
            Left            =   11205
            TabIndex        =   54
            Top             =   30
            Width           =   825
            _ExtentX        =   1455
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
            Caption         =   "합    계"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   1215
            Index           =   19
            Left            =   10890
            TabIndex        =   55
            Top             =   30
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   2143
            _Version        =   262144
            CaptionStyle    =   1
            BackColor       =   12648384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "결제 합계"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Index           =   20
            Left            =   11205
            TabIndex        =   56
            Top             =   330
            Width           =   825
            _ExtentX        =   1455
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
            Caption         =   "현    금"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Index           =   21
            Left            =   11205
            TabIndex        =   57
            Top             =   630
            Width           =   825
            _ExtentX        =   1455
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
            Caption         =   "카    드"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   315
            Index           =   20
            Left            =   12015
            TabIndex        =   58
            Top             =   30
            Width           =   1200
            _Version        =   262145
            _ExtentX        =   2117
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
            Index           =   22
            Left            =   12015
            TabIndex        =   59
            Top             =   630
            Width           =   1200
            _Version        =   262145
            _ExtentX        =   2117
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
            Index           =   22
            Left            =   13200
            TabIndex        =   60
            Top             =   30
            Width           =   825
            _ExtentX        =   1455
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
            Caption         =   "쿠    폰"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Index           =   23
            Left            =   13200
            TabIndex        =   61
            Top             =   330
            Width           =   825
            _ExtentX        =   1455
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
            Caption         =   "미    수"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   315
            Index           =   24
            Left            =   14010
            TabIndex        =   62
            Top             =   30
            Width           =   1200
            _Version        =   262145
            _ExtentX        =   2117
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
            Index           =   25
            Left            =   14010
            TabIndex        =   63
            Top             =   330
            Width           =   1200
            _Version        =   262145
            _ExtentX        =   2117
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
            Index           =   26
            Left            =   13200
            TabIndex        =   64
            Top             =   630
            Width           =   825
            _ExtentX        =   1455
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
            Caption         =   "반환현금"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   315
            Index           =   26
            Left            =   14010
            TabIndex        =   65
            Top             =   630
            Width           =   1200
            _Version        =   262145
            _ExtentX        =   2117
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
            Index           =   27
            Left            =   11205
            TabIndex        =   66
            Top             =   930
            Width           =   825
            _ExtentX        =   1455
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
            Caption         =   "마일리지"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   315
            Index           =   23
            Left            =   12015
            TabIndex        =   67
            Top             =   930
            Width           =   1200
            _Version        =   262145
            _ExtentX        =   2117
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
            Index           =   28
            Left            =   13200
            TabIndex        =   68
            Top             =   930
            Width           =   825
            _ExtentX        =   1455
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
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   315
            Index           =   27
            Left            =   14010
            TabIndex        =   69
            Top             =   930
            Width           =   1200
            _Version        =   262145
            _ExtentX        =   2117
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
            Index           =   29
            Left            =   2220
            TabIndex        =   70
            Top             =   930
            Width           =   825
            _ExtentX        =   1455
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
            Caption         =   "지사마진"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Index           =   30
            Left            =   330
            TabIndex        =   71
            Top             =   930
            Width           =   825
            _ExtentX        =   1455
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
            Caption         =   "가맹점마진"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   315
            Index           =   3
            Left            =   1140
            TabIndex        =   72
            Top             =   930
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
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
            Index           =   7
            Left            =   3030
            TabIndex        =   73
            Top             =   930
            Width           =   1260
            _Version        =   262145
            _ExtentX        =   2222
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
            Index           =   21
            Left            =   12015
            TabIndex        =   74
            Top             =   330
            Width           =   1200
            _Version        =   262145
            _ExtentX        =   2117
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
            Index           =   31
            Left            =   8895
            TabIndex        =   75
            Top             =   930
            Width           =   825
            _ExtentX        =   1455
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
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   315
            Index           =   19
            Left            =   9705
            TabIndex        =   76
            Top             =   930
            Width           =   1200
            _Version        =   262145
            _ExtentX        =   2117
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
      Begin Threed.SSPanel Panel 
         Height          =   750
         Left            =   15
         TabIndex        =   37
         Top             =   450
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   1323
         _Version        =   262144
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox cboGubun 
            Height          =   300
            Left            =   915
            Style           =   2  '드롭다운 목록
            TabIndex        =   44
            Top             =   405
            Width           =   1455
         End
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   2415
            TabIndex        =   43
            Top             =   405
            Width           =   2400
         End
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   630
            Left            =   8220
            TabIndex        =   38
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 조회(&F)"
            Appearance      =   6
            Picture         =   "frm매출.frx":3353
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   3
            Left            =   10605
            TabIndex        =   39
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 엑셀(&E)"
            Appearance      =   6
            Picture         =   "frm매출.frx":3A4D
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   13695
            TabIndex        =   40
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frm매출.frx":41C7
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   4
            Left            =   12150
            TabIndex        =   41
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 출력(&P)"
            Appearance      =   6
            Picture         =   "frm매출.frx":5259
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   0
            Left            =   915
            TabIndex        =   45
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
            Format          =   57016323
            CurrentDate     =   40279
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   1
            Left            =   2655
            TabIndex        =   46
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
            Format          =   57016323
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
            TabIndex        =   49
            Top             =   105
            Width           =   120
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "검색조건:"
            Height          =   180
            Index           =   3
            Left            =   45
            TabIndex        =   48
            Top             =   465
            Width           =   840
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "매출일자:"
            Height          =   195
            Index           =   2
            Left            =   45
            TabIndex        =   47
            Top             =   105
            Width           =   840
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   42
         Top             =   15
         Width           =   15210
         _ExtentX        =   26829
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
         Caption         =   "      매출 현황"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm매출.frx":5953
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm매출.frx":5B79
            Top             =   -15
            Width           =   765
         End
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   7500
         Left            =   15
         TabIndex        =   77
         Top             =   1215
         Width           =   15210
         _Version        =   524288
         _ExtentX        =   26829
         _ExtentY        =   13229
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
         MaxCols         =   16
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         SpreadDesigner  =   "frm매출.frx":6743
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "frm매출"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub 고객별매출_Display()
    On Error GoTo ErrRtn

    Dim 날짜         As String
    Dim TSUM(0 To 6) As Currency
        
    Dim 잔액         As Currency
    
    pnlProg.Visible = True
    DoEvents
    
    날짜 = ""
    잔액 = 0
        
    For i = 0 To 6
        TSUM(i) = 0
    Next i
    
    '-------------------------------------------------------------------------------------------
    '
    '-------------------------------------------------------------------------------------------
    Query = "SELECT    A.*"
    Query = Query & ", B.성명"
    Query = Query & ", B.전화번호"
    Query = Query & ", B.주소"
    Query = Query & " FROM TB_매출 AS A LEFT JOIN TB_고객정보 AS B ON A.고객코드 = B.고객코드"
    Query = Query & " WHERE (A.매출일자 >= '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "' "
    Query = Query & "   AND  A.매출일자 <= '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "') "
    
    If txtFind.Text <> "" Then
        Select Case cboGubun.Text
            Case "성명":     Query = Query & " AND B.성명 LIKE '%" & txtFind.Text & "%'"
            
            Case "전화번호": Query = Query & " AND (B.전화번호 LIKE '%" & txtFind.Text & "%'"
                             Query = Query & "  OR  B.휴대전화   LIKE '%" & txtFind.Text & "%')"
            
            Case "주소":     Query = Query & " AND B.주소 LIKE '%" & txtFind.Text & "%'"
        End Select
    End If
    
    Query = Query & " ORDER BY A.매출일자, A.매출시간, B.성명, A.접수번호, A.일련번호, A.적요 ASC"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = Format(ADORs!매출일자, "YYYY-MM-DD") & ""                  ' 1
            .Col = 2:  .Text = ADORs!성명 & ""                                          ' 2
            .Col = 3:  .Text = ADORs!전화번호 & ""                                      ' 3
            .Col = 4:  .Text = ADORs!주소 & ""                                          ' 4
            .Col = 5:  .Text = ADORs!접수번호 & ""                                      ' 5
            .Col = 6:  .Text = ADORs!적요 & ""                                          ' 6
            
            If ADORs!반품수량 <> 0 Then
                .Col = 7:  .Text = ""
                .Col = 8:  .Text = ADORs!접수금액
                
                '//
                .Row = .MaxRows: .Row2 = .MaxRows
                .Col = 6: .Col2 = .MaxCols
                .BlockMode = True
                .ForeColor = vbRed
                .BlockMode = False
            Else
                '----------------------------------------------------------------------------
                ' 정상적인 매출만 계산한다...
                '----------------------------------------------------------------------------
                잔액 = 잔액 + ADORs!접수금액 - ADORs!입금합계 '- ADORs!할인액  '
            End If
                        
'            '일련번호가 1 이라는 것은 접수할때 입금처리된 내용...
'            If ADORs!일련번호 = 1 Then
'                .Col = 7:  .Text = IIf(ADORs!접수수량 = 0, "", ADORs!접수수량) & ""     ' 7
'                .Col = 8:  .Text = ADORs!접수금액 & ""                                  ' 8
'            End If
            
            .Col = 7:  .Text = IIf(ADORs!접수수량 = 0, "", ADORs!접수수량) & ""     ' 7
            .Col = 8:  .Text = ADORs!접수금액 & ""                                  ' 8
            .Col = 9:  .Text = ADORs!현금입금 & "": .ForeColor = vbBlue                 ' 9
            .Col = 10: .Text = ADORs!카드입금 & "": .ForeColor = vbBlue                 ' 9
            .Col = 11: .Text = ADORs!사용마일리지 & ""                                        '10
            .Col = 12: .Text = ADORs!쿠폰입금 & ""                                    '11 2008-07-26 포인트사용 추가
            .Col = 13: .Text = 잔액 & ""                                                '12
            .Col = 14: .Text = ADORs!반품수량 & ""                                      '13
            .Col = 15: .Text = ADORs!고객코드 & ""                                      '14
            .Col = 16: .Text = ADORs!일련번호 & ""                                      '15
            
            날짜 = ADORs!매출일자
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        .ReDraw = True
    End With
    
    pnlProg.Visible = False
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub 일자별매출_Display()
    On Error GoTo ErrRtn

    Dim 날짜         As String
    Dim pRow         As Long
    
    pnlProg.Visible = True
    DoEvents
    
    날짜 = ""
        
    '-------------------------------------------------------------------------------------------
    '
    '-------------------------------------------------------------------------------------------
    Query = "SELECT    A.*"
    Query = Query & ", B.성명"
    Query = Query & ", B.전화번호"
    Query = Query & ", B.주소"
    Query = Query & " FROM TB_매출 AS A LEFT JOIN TB_고객정보 AS B ON A.고객코드 = B.고객코드"
    Query = Query & " WHERE (A.매출일자 >= '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "' "
    Query = Query & "   AND  A.매출일자 <= '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "') "
    
    If txtFind.Tag <> "" Then
        Query = Query & " AND A.고객코드 = " & txtFind.Tag
    End If
        
    Query = Query & " ORDER BY A.매출일자, A.매출시간, B.성명, A.접수번호, A.일련번호, A.적요 ASC"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        pRow = 1
        
        Do Until ADORs.EOF
            If (날짜 <> "") And (날짜 <> ADORs!매출일자) Then
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                
                .Row = .Row: .Row2 = .Row
                .Col = 1:    .Col2 = .MaxCols
                .BlockMode = True
                .BackColor = &HF5F5F5 '&HC0E0FF
                '.ForeColor = vbRed
                .BlockMode = False
                
                .Col = 1:  .Text = "소계 :"
                
                .Col = 7:  .Formula = "SUM(G" & pRow & ":G" & .MaxRows - 1 & ")"
                .Col = 8:  .Formula = "SUM(H" & pRow & ":H" & .MaxRows - 1 & ")"
                .Col = 9:  .Formula = "SUM(I" & pRow & ":I" & .MaxRows - 1 & ")"
                .Col = 10: .Formula = "SUM(J" & pRow & ":J" & .MaxRows - 1 & ")"
                .Col = 11: .Formula = "SUM(K" & pRow & ":K" & .MaxRows - 1 & ")"
                .Col = 12: .Formula = "SUM(L" & pRow & ":L" & .MaxRows - 1 & ")"
                .Col = 13: .Formula = "SUM(M" & pRow & ":M" & .MaxRows - 1 & ")"
                .Col = 14: .Formula = "SUM(N" & pRow & ":N" & .MaxRows - 1 & ")"
                            
                pRow = .MaxRows + 1
            End If
            
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = Format(ADORs!매출일자, "YYYY-MM-DD") & ""
            .Col = 2:  .Text = ADORs!성명 & ""
            .Col = 3:  .Text = ADORs!전화번호 & ""
            
            If ADORs!주소 = "" Then
                .Col = 4:  .Text = " "
            Else
                .Col = 4:  .Text = ADORs!주소 & ""
            End If
            
            .Col = 5:  .Text = ADORs!접수번호 & ""
            .Col = 6:  .Text = ADORs!적요 & ""
            .Col = 7:  .Text = ADORs!접수수량 & ""                             '
            .Col = 8:  .Text = ADORs!접수금액 & ""                             '
            .Col = 9:  .Text = ADORs!현금입금 & ""                             '
            .Col = 10: .Text = ADORs!카드입금 & ""                             '
            .Col = 11: .Text = ADORs!사용마일리지 & ""                         '
            .Col = 12: .Text = ADORs!쿠폰입금 & ""                             '
            .Col = 13: .Text = ADORs!접수금액 - ADORs!현금입금 - ADORs!카드입금 - ADORs!사용마일리지 - ADORs!쿠폰입금 & ""
            .Col = 14: .Text = ADORs!반품수량 & ""                             '
            .Col = 15: .Text = ADORs!고객코드 & ""                             '
            .Col = 16: .Text = ADORs!일련번호 & ""                             '
            
            날짜 = ADORs!매출일자
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        If .MaxRows > 0 Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Row = .Row: .Row2 = .Row
            .Col = 1:    .Col2 = .MaxCols
            .BlockMode = True
            .BackColor = &HF5F5F5 '&HC0E0FF
            '.ForeColor = vbRed
            .BlockMode = False
            
            .Col = 1:  .Text = "소계 :"
            
            .Col = 7:  .Formula = "SUM(G" & pRow & ":G" & .MaxRows - 1 & ")"
            .Col = 8:  .Formula = "SUM(H" & pRow & ":H" & .MaxRows - 1 & ")"
            .Col = 9:  .Formula = "SUM(I" & pRow & ":I" & .MaxRows - 1 & ")"
            .Col = 10: .Formula = "SUM(J" & pRow & ":J" & .MaxRows - 1 & ")"
            .Col = 11: .Formula = "SUM(K" & pRow & ":K" & .MaxRows - 1 & ")"
            .Col = 12: .Formula = "SUM(L" & pRow & ":L" & .MaxRows - 1 & ")"
            .Col = 13: .Formula = "SUM(M" & pRow & ":M" & .MaxRows - 1 & ")"
            .Col = 14: .Formula = "SUM(N" & pRow & ":N" & .MaxRows - 1 & ")"
        End If
            
        .ReDraw = True
    End With
    
    pnlProg.Visible = False
    Exit Sub
    
ErrRtn:
    pnlProg.Visible = False
    
    Call Error_Msg("frm매출.일자별매출_Display", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
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
    
    Open AppPath & "XML\매출현황.XML" For Output As #1
    
    Print #1, "<?xml version=""1.0"" encoding=""EUC-KR""?>"
    Print #1, "<root>"
    
          XML = "    <조건>"
    XML = XML & "        <매출일자>매출일자 : " & dtpDay(0).Value & " ~ " & dtpDay(1).Value & "</매출일자>"
    XML = XML & "        <가맹점>크린에이드 - " & Func_Replace(가맹점정보.가맹점명) & "</가맹점>"
    XML = XML & "   </조건>"
    Print #1, XML
    
    With sprGrid
        For i = 1 To .MaxRows
            .Row = i
            
                             XML = "    <Data>"
            .Col = 1:  XML = XML & "        <매출일자>" & .Text & "</매출일자>"
            .Col = 2:  XML = XML & "        <고객명>" & .Text & "</고객명>"
            .Col = 3:  XML = XML & "        <전화번호>" & .Text & "</전화번호>"
            .Col = 4:  XML = XML & "        <주소>" & .Text & "</주소>"
            .Col = 5:  XML = XML & "        <No>" & .Text & "</No>"
            .Col = 6:  XML = XML & "        <적요>" & .Text & "</적요>"
            .Col = 7:  XML = XML & "        <접수량>" & .Text & "</접수량>"
            .Col = 8:  XML = XML & "        <총금액>" & .Text & "</총금액>"
            .Col = 9:  XML = XML & "        <현금입금>" & .Text & "</현금입금>"
            .Col = 10: XML = XML & "        <카드입금>" & .Text & "</카드입금>"
            .Col = 11: XML = XML & "        <마일리지>" & .Text & "</마일리지>"
            .Col = 12: XML = XML & "        <쿠폰금액>" & .Text & "</쿠폰금액>"
            .Col = 13: XML = XML & "        <잔액>" & .Text & "</잔액>"
            .Col = 14: XML = XML & "        <반품수량>" & .Text & "</반품수량>"
                       XML = XML & "   </Data>"
            Print #1, XML
        Next i
        
'              XML = "    <합계>"
'        XML = XML & "        <품목수량>" & txtNum(0).Text & "</품목수량>"
'        XML = XML & "        <접수금액>" & txtNum(5).Text & "</접수금액>"
'        XML = XML & "        <세트할인>" & txtNum(2).Text & "</세트할인>"
'        XML = XML & "        <입금액>" & txtNum(3).Text & "</입금액>"
'        XML = XML & "        <미수금액>" & txtNum(4).Text & "</미수금액>"
'        XML = XML & "   </합계>"
'        Print #1, XML
        
        Print #1, "</root>"
        Close #1
    End With
    
    If Print_PreView = True Then
        With rpt매출현황
            .dc.FileURL = AppPath & "XML\매출현황.XML"
            .Show 1
        End With
    Else
        With rpt매출현황
            .dc.FileURL = AppPath & "XML\매출현황.XML"
            .PrintReport False
        End With
            
        Unload rpt매출현황
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Public Sub cmdList_Click()
    If txtFind.Text = "" Then
        Call 일자별매출_Display
    Else
        Call 고객별매출_Display
    End If
    
    Call Total_Display
End Sub

Private Sub dtpDay_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    With sprGrid
        .MaxRows = 0
        .RowHeight(-1) = 14
        
        '.Col = 5:  .ColHidden = True '접수번호
        .Col = 15: .ColHidden = True '고객코드
        .Col = 16: .ColHidden = True '일련번호
        
        .Col = 1: .ColMerge = MergeRestricted
        .Col = 2: .ColMerge = MergeRestricted
        .Col = 3: .ColMerge = MergeRestricted
        .Col = 4: .ColMerge = MergeRestricted
        .Col = 5: .ColMerge = MergeRestricted
        
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
            
    With fpList1
        .ColumnHeaderHeight = 300
        .RowHeight = 300
    
        .ListApplyTo = ListApplyToColHeaders
        .BackColor = RGB(192, 192, 192)
        .LineStyle = LineStyleRaised
    End With
            
    With cboGubun
        .Clear
        .AddItem "성명"
        .AddItem "전화번호"
        .AddItem "주소"
        
        .ListIndex = 0
    End With
    
    dtpDay(0).Value = Format(Date, "YYYY-MM-DD")
    dtpDay(1).Value = Format(Date, "YYYY-MM-DD")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    pnlHeader.Width = Me.ScaleWidth
    cmdBtn(5).Left = (Me.Width - 200) - cmdBtn(5).Width
End Sub

Private Sub fpList1_DblClick()
    On Error GoTo ErrRtn
    
    With fpList1
        .Col = 0: .Row = .ListIndex: txtFind.Tag = Trim(.ColList)  '코드
        .Col = 1: .Row = .ListIndex: txtFind.Text = Trim(.ColList) '이름
        
        If txtFind.Tag = "" Then
            .Visible = False
            
            txtFind.SetFocus
            Exit Sub
        End If
        
        .Visible = False
        
        Call cmdList_Click
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub fpList1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            KeyAscii = 0
            fpList1_DblClick
        
        Case 27
            KeyAscii = 0
            fpList1.Visible = False
            txtFind.SetFocus
    End Select
End Sub

Private Sub fpList1_LostFocus()
    fpList1.Visible = False
End Sub

Private Sub cmdFind1_Click()
    On Error GoTo ErrRtn
    
    Query = "SELECT    고객코드"
    Query = Query & ", 성명"
    Query = Query & ", 전화번호"
    Query = Query & ", 주소"
    Query = Query & " FROM TB_고객정보"
    'Query = Query & " WHERE 삭제 = False" '2008-07-06
    Query = Query & " WHERE 삭제 = 0" '2008-07-06
    
    If txtFind.Text <> "" Then
        If cboGubun.Text = "전화번호" Then
            Query = Query & " AND (전화번호  LIKE '%" & txtFind.Text & "%'"
            Query = Query & "  OR  전화번호2 LIKE '%" & txtFind.Text & "%'"
            Query = Query & "  OR  전화번호3 LIKE '%" & txtFind.Text & "%'"
            Query = Query & "  OR  전화번호4 LIKE '%" & txtFind.Text & "%')"
            Query = Query & " ORDER BY " & cboGubun.Text & " ASC"
        Else
            Query = Query & " AND " & cboGubun.Text & " LIKE '%" & txtFind.Text & "%' "
            Query = Query & " ORDER BY " & cboGubun.Text & " ASC"
        End If
    Else
        Query = Query & " ORDER BY 성명 ASC"
    End If
    
    'If txtFind.Text <> "" Then
    '    Query = Query & " AND 성명 LIKE '%" & txtFind.Text & "%' "
    'End If
    
    'Query = Query & " ORDER BY 성명 ASC "
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenStatic, adLockReadOnly
    
    With fpList1
        Set .DataSource = ADORs
        
        .Top = 1155
        .Left = 5040
        
        .Visible = True
        .SetFocus
    End With
    
    ADORs.Close
    Set ADORs = Nothing
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub Total_Display()
    On Error GoTo ErrRtn
    
    Dim 조회일자1 As String
    Dim 조회일자2 As String
    
    조회일자1 = Format(dtpDay(0).Value, "YYYY-MM-DD")
    조회일자2 = Format(dtpDay(1).Value, "YYYY-MM-DD")
    
    For i = 0 To 19
        txtNum(i).Value = 0
    Next i
    
    '--------------------------------------------------------------------------------------
    ' 1. 매출 1-1) 접수수량/접수금액
    '--------------------------------------------------------------------------------------
    Query = "SELECT    ISNULL(COUNT(*),0)"
    Query = Query & ", ISNULL(SUM(금액),0)"
    Query = Query & " FROM TB_입출고"
    Query = Query & " WHERE 접수일자 BETWEEN '" & 조회일자1 & "' AND '" & 조회일자2 & "' "
    Query = Query & "  AND (판매취소 <> 'Y')"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    txtNum(0).Value = ADORs(0) & "" '접수수량
    txtNum(4).Value = ADORs(1) & "" '접수금액
    ADORs.Close:    Set ADORs = Nothing
    
    '--------------------------------------------------------------------------------------
    ' 1. 매출 1-2) 판매취소 수량 / 금액
    '--------------------------------------------------------------------------------------
    Query = "SELECT    ISNULL(COUNT(*),0)"
    Query = Query & ", ISNULL(SUM(금액),0)"
    Query = Query & " FROM TB_입출고"
    Query = Query & " WHERE 접수일자 BETWEEN '" & 조회일자1 & "' AND '" & 조회일자2 & "' "
    Query = Query & "  AND (판매취소 = 'Y')"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    txtNum(1).Value = ADORs(0) & "" '판매취소수량
    txtNum(5).Value = ADORs(1) & "" '판매취소금액
    ADORs.Close:    Set ADORs = Nothing
        
    '--------------------------------------------------------------------------------------
    ' 1. 매출 1-3) 출고수량
    '--------------------------------------------------------------------------------------
    Query = "SELECT ISNULL(COUNT(*),0)"
    Query = Query & " FROM TB_입출고"
    Query = Query & " WHERE 출고일자 BETWEEN '" & 조회일자1 & "' AND '" & 조회일자2 & "' "
    Query = Query & "  AND (판매취소 <> 'Y')"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    txtNum(2).Value = ADORs(0) & "" '출고수량
    ADORs.Close:    Set ADORs = Nothing


    
    '----------------------------------------------------------------
    ' 2. 선불결제 2-1) 합계
    '----------------------------------------------------------------
    ' 마지막에 구한다.
    
    '----------------------------------------------------------------
    ' 2. 선불결제 2-2) 현금
    '----------------------------------------------------------------
    Query = "SELECT ISNULL(SUM(현금입금),0)"
    Query = Query & " FROM TB_매출 "
    Query = Query & " WHERE 매출일자 BETWEEN '" & 조회일자1 & "' AND '" & 조회일자2 & "' "
    Query = Query & "   AND 접수금액 <> 0"
    ' Query = Query & "   AND NOT 적요 LIKE '%판매취소%' "
    Query = Query & "   AND NOT 적요 LIKE '%미수금액 입금%'"

    txtNum(9).Value = Recordset_Result(Query) '
    
    '----------------------------------------------------------------
    ' 2. 선불결제 2-3) 카드
    ' 건수= 승인 + 취소 , 금액 = 승인 + 취소
    '----------------------------------------------------------------
    Query = "SELECT    ISNULL(COUNT(카드입금),0)"
    Query = Query & ", ISNULL(SUM(카드입금),0)"
    Query = Query & " FROM TB_매출 "
    Query = Query & " WHERE 매출일자 BETWEEN '" & 조회일자1 & "' AND '" & 조회일자2 & "' "
    Query = Query & "   AND 카드입금 <> 0" ' 카드 금결제가 아닌 경우도 0원이 들어간다.
'   Query = Query & "   AND 접수금액 > 0"
'   Query = Query & "   AND NOT 적요 LIKE '%판매취소%' "
    Query = Query & "   AND NOT 적요 LIKE '%미수금액 입금%'"
    Query = Query & "   AND NOT 적요 LIKE '%반품환불%'"
    Query = Query & "   AND NOT 적요 LIKE '%세탁환불%'"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    txtNum(10).Value = ADORs(1)
    ADORs.Close:    Set ADORs = Nothing
    
    '--------------------------------------------------------------------
    ' 2. 선불결제 2-4) 발생/사용/삭제 마일리지
    '--------------------------------------------------------------------
    Query = "SELECT    ISNULL(SUM(발생마일리지),0)"
    Query = Query & ", ISNULL(SUM(사용마일리지),0)"
    Query = Query & ", ISNULL(SUM(삭제마일리지),0)"
    Query = Query & " FROM TB_매출"
    Query = Query & " WHERE 매출일자 BETWEEN '" & 조회일자1 & "' AND '" & 조회일자2 & "' "
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    txtNum(11).Value = ADORs(1) '사용마일리지
    ADORs.Close:    Set ADORs = Nothing
    
    '--------------------------------------------------------------------
    ' 2. 선불결제 2-5) 쿠폰
    '--------------------------------------------------------------------
    Query = "SELECT    ISNULL(SUM(쿠폰입금),0)"
    Query = Query & ", ISNULL(COUNT(쿠폰번호),0)"
    Query = Query & " FROM TB_매출"
    Query = Query & " WHERE 매출일자 BETWEEN '" & 조회일자1 & "' AND '" & 조회일자2 & "' "
    Query = Query & "   AND 쿠폰입금 > 0"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    txtNum(12).Value = ADORs(0)  ' 금액
    ADORs.Close:    Set ADORs = Nothing
    
    '----------------------------------------------------------------
    ' 2. 선불결제 2-6) 미수금 금액
    ' 마일리지를 사용한 것을 판매취소할 경우 미수금액이 마일리지 사용한 것으로 처리되기 때문에
    ' 별도로 마일리지판매취소 금액을 구해서 -해준다.
    ' 마일리지판매취소 값이 -로 넘어오기 때문에 - 해주면 +로 처리된다.(위에서 마일리지 금액이 처리되어 나요기 때문)
    '----------------------------------------------------------------
    Query = "SELECT ISNULL(SUM(접수금액),0) - ISNULL(SUM(입금합계),0) - ISNULL(SUM(쿠폰입금),0) AS 미수금 "
    Query = Query & " FROM TB_매출"
    Query = Query & " WHERE 매출일자 BETWEEN '" & 조회일자1 & "' AND '" & 조회일자2 & "' "
'   Query = Query & "   AND 접수금액 > 0"
    Query = Query & "   AND NOT 적요  LIKE '%미수금액 입금%'"
    Query = Query & "   AND NOT 적요  LIKE '%반품환불%'"
    Query = Query & "   AND NOT 적요  LIKE '%세탁환불%'"

    txtNum(13).Value = Recordset_Result(Query) '
    
    Query = "SELECT ISNULL(SUM(사용마일리지),0)  AS 판매취소마일리지금액 "
    Query = Query & " FROM TB_매출"
    Query = Query & " WHERE 매출일자 BETWEEN '" & 조회일자1 & "' AND '" & 조회일자2 & "' "
    Query = Query & "   AND 접수금액 < 0"
    Query = Query & "   AND 사용마일리지 < 0"
    Query = Query & "   AND NOT 적요  LIKE '%미수금액 입금%'"
    Query = Query & "   AND NOT 적요  LIKE '%반품환불%'"
    Query = Query & "   AND NOT 적요  LIKE '%세탁환불%'"
    txtNum(13).Value = txtNum(13).Value - Recordset_Result(Query)  '
    
    '----------------------------------------------------------------
    ' 2. 선불결제 2-7) 현금반환/ 현금결제 구하기
    '----------------------------------------------------------------
    Query = "SELECT    ISNULL(SUM(접수금액),0) * -1"
    Query = Query & " FROM TB_매출"
    Query = Query & " WHERE 매출일자 BETWEEN '" & 조회일자1 & "' AND '" & 조회일자2 & "' "
    Query = Query & "   AND 적요 LIKE '%현금반환%' "
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    txtNum(14).Value = ADORs(0)  ' 금액
    ADORs.Close:    Set ADORs = Nothing
    
    '----------------------------------------------------------------
    ' 2. 선불결제 2-1) 합계
    '----------------------------------------------------------------
    ' 합계 = 현금 + 카드 + 마일리지 + 쿠폰 + 미수
    txtNum(8).Value = txtNum(9).Value + txtNum(10).Value + txtNum(11).Value + txtNum(12).Value + txtNum(13).Value
    
    
    '----------------------------------------------------------------
    ' 3. 미수결제 3-1) 미수금 수금 현금결제 구하기
    '----------------------------------------------------------------
    Query = "SELECT ISNULL(SUM(현금입금),0)"
    Query = Query & " FROM TB_매출"
    Query = Query & " WHERE 매출일자 BETWEEN '" & 조회일자1 & "' AND '" & 조회일자2 & "' "
    Query = Query & "   AND 접수금액 = 0"
    Query = Query & "   AND 적요  LIKE '%미수금액 입금%'"

    txtNum(17).Value = Recordset_Result(Query) '
    
    '----------------------------------------------------------------
    ' 3. 미수결제 3-2) 미수금 수금 카드결제 구하기
    '----------------------------------------------------------------
    Query = "SELECT    ISNULL(COUNT(카드입금),0)"
    Query = Query & ", ISNULL(SUM(카드입금),0)"
    Query = Query & " FROM TB_매출"
    Query = Query & " WHERE 매출일자 BETWEEN '" & 조회일자1 & "' AND '" & 조회일자2 & "' "
    Query = Query & "   AND 카드입금 <> 0"
    Query = Query & "   AND 접수금액 = 0" ' 판매취소시 0원으로 들어온다.
    Query = Query & "   AND 적요  LIKE '%미수금액 입금%'"

    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    txtNum(18).Value = ADORs(1)
    ADORs.Close:    Set ADORs = Nothing

    '----------------------------------------------------------------
    ' 3. 미수결제 3-1) 합계
    '----------------------------------------------------------------
    ' 합계 = 현금 + 카드
    txtNum(16).Value = txtNum(17).Value + txtNum(18).Value + txtNum(19).Value
    
    '----------------------------------------------------------------
    ' 4. 결제합계
    '----------------------------------------------------------------
    txtNum(20).Value = txtNum(8).Value + txtNum(16).Value   ' 합계
    txtNum(21).Value = txtNum(9).Value + txtNum(17).Value   ' 현금
    txtNum(22).Value = txtNum(10).Value + txtNum(18).Value  ' 카드
    txtNum(23).Value = txtNum(11).Value + txtNum(19).Value  ' 마일리지
    txtNum(24).Value = txtNum(12).Value                     ' 쿠폰
    txtNum(25).Value = txtNum(13).Value                     ' 미수
    txtNum(26).Value = txtNum(14).Value                     ' 반환현금
    txtNum(27).Value = txtNum(15).Value
    
    '----------------------------------------------------------------
    ' 5. 마진 5-1) 가맹점 마진 / 지사마진
    '----------------------------------------------------------------
    Query = "SELECT ISNULL(SUM(금액 * 세탁마진/100),0)"
    Query = Query & " FROM TB_입출고"
    Query = Query & " WHERE 접수일자 BETWEEN '" & 조회일자1 & "' AND '" & 조회일자2 & "' "
    Query = Query & "   AND 내용 NOT LIKE '%수%'"                            '수선 제외
    Query = Query & "   AND (판매취소 <> 'Y')"

    txtNum(3).Value = Recordset_Result(Query)

    ' 마일리지 사용이 있을 경우
    If txtNum(11).Value > 0 Then
        txtNum(3).Value = txtNum(3).Value - CLng(txtNum(11).Value * 0.4) '가맹점  지사:가맹점(6:4)로 빼준다.
        txtNum(7).Value = txtNum(7).Value - CLng(txtNum(11).Value * 0.6) '지사
    End If

'    '쿠폰 사용이 있는 경우
'    If txtCost21.Value > 0 And 마감일자 <= "2011-12-31" Then
'        txtCost09.Value = txtCost09.Value - CLng(1200 * txtNum12.Value * 0.4) '가맹점
'        txtCost10.Value = txtCost10.Value - CLng(1200 * txtNum12.Value * 0.6) '지사
'    End If

    '----------------------------------------------------------------
    ' 5. 마진 5-2) 지사 마진
    '----------------------------------------------------------------
    txtNum(7).Value = (txtNum(4).Value - txtNum(11).Value) - txtNum(3).Value                    ' 지사 마진 = 접수금액 - 가맹점마진
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

