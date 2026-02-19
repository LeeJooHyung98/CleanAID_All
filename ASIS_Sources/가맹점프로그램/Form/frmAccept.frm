VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{D97EED01-E916-11D3-9B9A-525405E0F0CD}#1.22#0"; "hooncontrol.ocx"
Begin VB.Form frmAccept 
   BackColor       =   &H00FFFFFF&
   Caption         =   "고객입출고"
   ClientHeight    =   11175
   ClientLeft      =   5865
   ClientTop       =   3135
   ClientWidth     =   15390
   LinkTopic       =   "Form1"
   ScaleHeight     =   11175
   ScaleWidth      =   15390
   Begin HoonControl.MyText txtSearch 
      Height          =   420
      Left            =   1125
      TabIndex        =   115
      Top             =   90
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   741
      TextType        =   0
      Text            =   ""
      SelStart        =   0
      SelLength       =   0
      BackColor       =   16777215
      ForeColor       =   0
      EnableColor     =   16777215
      DisableColor    =   12632256
      LockColor       =   14737632
      EditColor       =   12648447
      Maxlength       =   0
      LenDecimal      =   0
      LenInteger      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   11.25
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontUnder       =   0   'False
      Locked          =   0   'False
      PasswordChar    =   ""
      Enabled         =   -1  'True
      FocusSel        =   -1  'True
      DASH            =   1
      AutoColor       =   0
      AutoTrim        =   0
      AutoCalc        =   0
      AutoMove        =   0
      Alignment       =   0
      Appearance      =   0
      Margine         =   0
      IMEMode         =   0
      LenType         =   0
      DateDevider     =   0
      ECase           =   0
      TextType        =   0
   End
   Begin CSTextLibCtl.sitxEdit txtAddress 
      Height          =   420
      Left            =   1125
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   990
      Width           =   7740
      _Version        =   262145
      _ExtentX        =   13652
      _ExtentY        =   741
      _StockProps     =   125
      Text            =   "010-9858-0428"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   11.26
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ReadOnly        =   -1  'True
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   "010-9858-0428"
      Text            =   "010-9858-0428"
      StartText.x     =   2
      StartText.y     =   5
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   17
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      CharacterTable  =   ""
   End
   Begin CSTextLibCtl.sitxEdit txtName 
      Height          =   420
      Left            =   1125
      TabIndex        =   98
      TabStop         =   0   'False
      Top             =   555
      Width           =   1860
      _Version        =   262145
      _ExtentX        =   3281
      _ExtentY        =   741
      _StockProps     =   125
      Text            =   "010-9858-0428"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   11.26
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ReadOnly        =   -1  'True
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   "010-9858-0428"
      Text            =   "010-9858-0428"
      StartText.x     =   2
      StartText.y     =   5
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   17
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      CharacterTable  =   ""
   End
   Begin CSTextLibCtl.sitxEdit txtHP 
      Height          =   420
      Left            =   7005
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   555
      Width           =   1860
      _Version        =   262145
      _ExtentX        =   3281
      _ExtentY        =   741
      _StockProps     =   125
      Text            =   "010-9858-0428"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   11.26
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ReadOnly        =   -1  'True
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   "010-9858-0428"
      Text            =   "010-9858-0428"
      StartText.x     =   2
      StartText.y     =   5
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   17
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      CharacterTable  =   ""
   End
   Begin VB.TextBox txtMemo 
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
      Height          =   825
      Left            =   1125
      MultiLine       =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1425
      Width           =   7740
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   8790
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2325
      Width           =   15270
      _Version        =   851970
      _ExtentX        =   26935
      _ExtentY        =   15505
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
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      ItemCount       =   3
      Item(0).Caption =   "    접  수    "
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "TabControlPage1"
      Item(1).Caption =   "    출  고    "
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "TabControlPage7"
      Item(2).Caption =   "    출고내역  "
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "TabControlPage2"
      Item(2).Control(1)=   "TabControlPage3"
      Begin XtremeSuiteControls.TabControlPage TabControlPage3 
         Height          =   8190
         Left            =   -69970
         TabIndex        =   1
         Top             =   570
         Visible         =   0   'False
         Width           =   15210
         _Version        =   851970
         _ExtentX        =   26829
         _ExtentY        =   14446
         _StockProps     =   1
         BackColor       =   -2147483633
         Page            =   3
         Begin XtremeSuiteControls.GroupBox GroupBox4 
            Height          =   8055
            Left            =   11445
            TabIndex        =   12
            Top             =   60
            Width           =   3735
            _Version        =   851970
            _ExtentX        =   6588
            _ExtentY        =   14208
            _StockProps     =   79
            Caption         =   "2. 결제진행"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   6
            Begin XtremeSuiteControls.FlatEdit FlatEdit5 
               Height          =   315
               Left            =   120
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   3210
               Width           =   3525
               _Version        =   851970
               _ExtentX        =   6218
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   12
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "FlatEdit1"
               Alignment       =   1
               Appearance      =   1
               FlatStyle       =   -1  'True
               UseVisualStyle  =   0   'False
               ShowBorder      =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit6 
               Height          =   315
               Left            =   120
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   2535
               Width           =   3525
               _Version        =   851970
               _ExtentX        =   6218
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   12
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "FlatEdit1"
               Alignment       =   1
               Appearance      =   1
               FlatStyle       =   -1  'True
               UseVisualStyle  =   0   'False
               ShowBorder      =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit7 
               Height          =   315
               Left            =   120
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   1860
               Width           =   3525
               _Version        =   851970
               _ExtentX        =   6218
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   12
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "FlatEdit1"
               Alignment       =   1
               Appearance      =   1
               FlatStyle       =   -1  'True
               UseVisualStyle  =   0   'False
               ShowBorder      =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit8 
               Height          =   315
               Left            =   120
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   1185
               Width           =   3525
               _Version        =   851970
               _ExtentX        =   6218
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   12
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "FlatEdit1"
               Alignment       =   1
               Appearance      =   1
               FlatStyle       =   -1  'True
               UseVisualStyle  =   0   'False
               ShowBorder      =   0   'False
            End
            Begin XtremeSuiteControls.PushButton PushButton2 
               Height          =   825
               Left            =   90
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   3600
               Width           =   3585
               _Version        =   851970
               _ExtentX        =   6324
               _ExtentY        =   1455
               _StockProps     =   79
               Caption         =   "접수결제"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   14.25
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
            End
            Begin Threed.SSPanel SSPanel7 
               Height          =   525
               Index           =   4
               Left            =   105
               TabIndex        =   22
               Top             =   315
               Width           =   3555
               _ExtentX        =   6271
               _ExtentY        =   926
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
               Caption         =   "합계내역"
               PictureBackgroundStyle=   2
               PictureBackground=   "frmAccept.frx":0000
               BorderWidth     =   0
               BevelOuter      =   0
               PictureAlignment=   9
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin XtremeSuiteControls.PushButton PushButton4 
               Height          =   825
               Left            =   90
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   4425
               Width           =   3585
               _Version        =   851970
               _ExtentX        =   6324
               _ExtentY        =   1455
               _StockProps     =   79
               Caption         =   "후불"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   14.25
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
            End
            Begin VB.TextBox Text7 
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
               Height          =   690
               Left            =   105
               TabIndex        =   18
               TabStop         =   0   'False
               Text            =   "건수"
               Top             =   825
               Width           =   3555
            End
            Begin VB.TextBox Text9 
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
               Height          =   690
               Left            =   105
               TabIndex        =   21
               TabStop         =   0   'False
               Text            =   "접수금액"
               Top             =   1500
               Width           =   3555
            End
            Begin VB.TextBox Text8 
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
               Height          =   690
               Left            =   105
               TabIndex        =   20
               TabStop         =   0   'False
               Text            =   "할인금액"
               Top             =   2175
               Width           =   3555
            End
            Begin VB.TextBox Text6 
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
               Height          =   690
               Left            =   105
               TabIndex        =   17
               TabStop         =   0   'False
               Text            =   "합계금액"
               Top             =   2850
               Width           =   3555
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox5 
            Height          =   8070
            Left            =   30
            TabIndex        =   24
            Top             =   75
            Width           =   11370
            _Version        =   851970
            _ExtentX        =   20055
            _ExtentY        =   14235
            _StockProps     =   79
            Caption         =   "1. 주문내역"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   6
            Begin FPSpreadADO.fpSpread fpSpread1 
               Height          =   7695
               Left            =   60
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   315
               Width           =   11250
               _Version        =   524288
               _ExtentX        =   19844
               _ExtentY        =   13573
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
               SpreadDesigner  =   "frmAccept.frx":0342
               UserResize      =   1
               VisibleCols     =   7
               VisibleRows     =   30
               HighlightHeaders=   1
               HighlightStyle  =   1
               ScrollBarStyle  =   2
            End
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage2 
         Height          =   8190
         Left            =   -69970
         TabIndex        =   2
         Top             =   570
         Visible         =   0   'False
         Width           =   15210
         _Version        =   851970
         _ExtentX        =   26829
         _ExtentY        =   14446
         _StockProps     =   1
         BackColor       =   -2147483633
         Page            =   2
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage1 
         Height          =   8190
         Left            =   30
         TabIndex        =   58
         Top             =   570
         Width           =   15210
         _Version        =   851970
         _ExtentX        =   26829
         _ExtentY        =   14446
         _StockProps     =   1
         BackColor       =   -2147483633
         Page            =   0
         Begin XtremeSuiteControls.GroupBox GroupBox3 
            Height          =   5235
            Left            =   11415
            TabIndex        =   59
            Top             =   2895
            Width           =   3735
            _Version        =   851970
            _ExtentX        =   6588
            _ExtentY        =   9234
            _StockProps     =   79
            Caption         =   "3. 결제진행"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   6
            Begin HoonControl.MyLabel lblSalePrice 
               Height          =   375
               Left            =   120
               Top             =   3150
               Width           =   3525
               _ExtentX        =   6218
               _ExtentY        =   661
               Caption         =   "0"
               SelForeColor    =   0
               UnSelForeColor  =   0
               EditForeColor   =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   12
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontSize        =   12
               FontBold        =   -1  'True
               FontItalic      =   0   'False
               FontUnder       =   0   'False
               Alignment       =   1
               BorderStyle     =   0
               Margine         =   0
               BackColor       =   16777215
            End
            Begin HoonControl.MyLabel lblDiscountPrice 
               Height          =   375
               Left            =   120
               Top             =   2475
               Width           =   3525
               _ExtentX        =   6218
               _ExtentY        =   661
               Caption         =   "0"
               ForeColor       =   255
               SelForeColor    =   0
               UnSelForeColor  =   0
               EditForeColor   =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   12
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontSize        =   12
               FontBold        =   -1  'True
               FontItalic      =   0   'False
               FontUnder       =   0   'False
               Alignment       =   1
               BorderStyle     =   0
               Margine         =   0
               BackColor       =   16777215
            End
            Begin HoonControl.MyLabel lblOriginPrice 
               Height          =   375
               Left            =   120
               Top             =   1800
               Width           =   3525
               _ExtentX        =   6218
               _ExtentY        =   661
               Caption         =   "0"
               ForeColor       =   16711680
               SelForeColor    =   0
               UnSelForeColor  =   0
               EditForeColor   =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   12
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontSize        =   12
               FontBold        =   -1  'True
               FontItalic      =   0   'False
               FontUnder       =   0   'False
               Alignment       =   1
               BorderStyle     =   0
               Margine         =   0
               BackColor       =   16777215
            End
            Begin HoonControl.MyLabel lblCount 
               Height          =   375
               Left            =   120
               Top             =   1125
               Width           =   3525
               _ExtentX        =   6218
               _ExtentY        =   661
               Caption         =   "0"
               SelForeColor    =   0
               UnSelForeColor  =   0
               EditForeColor   =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   12
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontSize        =   12
               FontBold        =   -1  'True
               FontItalic      =   0   'False
               FontUnder       =   0   'False
               Alignment       =   1
               BorderStyle     =   0
               Margine         =   0
               BackColor       =   16777215
            End
            Begin XtremeSuiteControls.PushButton PushButton1 
               Height          =   825
               Left            =   90
               TabIndex        =   62
               TabStop         =   0   'False
               Top             =   3600
               Width           =   3585
               _Version        =   851970
               _ExtentX        =   6324
               _ExtentY        =   1455
               _StockProps     =   79
               Caption         =   "접수결제"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   14.25
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
            End
            Begin Threed.SSPanel SSPanel7 
               Height          =   525
               Index           =   3
               Left            =   105
               TabIndex        =   65
               Top             =   315
               Width           =   3555
               _ExtentX        =   6271
               _ExtentY        =   926
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
               Caption         =   "합계내역"
               PictureBackgroundStyle=   2
               PictureBackground=   "frmAccept.frx":1E60
               BorderWidth     =   0
               BevelOuter      =   0
               PictureAlignment=   9
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin XtremeSuiteControls.PushButton PushButton3 
               Height          =   825
               Left            =   90
               TabIndex        =   66
               TabStop         =   0   'False
               Top             =   4425
               Width           =   3585
               _Version        =   851970
               _ExtentX        =   6324
               _ExtentY        =   1455
               _StockProps     =   79
               Caption         =   "후불"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   14.25
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
            End
            Begin VB.TextBox Text2 
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
               ForeColor       =   &H00000000&
               Height          =   690
               Left            =   105
               Locked          =   -1  'True
               TabIndex        =   61
               TabStop         =   0   'False
               Text            =   "건수"
               Top             =   825
               Width           =   3555
            End
            Begin VB.TextBox Text3 
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
               ForeColor       =   &H00FF0000&
               Height          =   690
               Left            =   105
               Locked          =   -1  'True
               TabIndex        =   64
               TabStop         =   0   'False
               Text            =   "정상금액"
               Top             =   1500
               Width           =   3555
            End
            Begin VB.TextBox Text4 
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
               ForeColor       =   &H000000FF&
               Height          =   690
               Left            =   105
               TabIndex        =   63
               TabStop         =   0   'False
               Text            =   "할인금액"
               Top             =   2175
               Width           =   3555
            End
            Begin VB.TextBox Text5 
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
               Height          =   690
               Left            =   105
               Locked          =   -1  'True
               TabIndex        =   60
               TabStop         =   0   'False
               Text            =   "합계금액"
               Top             =   2850
               Width           =   3555
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox1 
            Height          =   2700
            Left            =   30
            TabIndex        =   67
            Top             =   75
            Width           =   15135
            _Version        =   851970
            _ExtentX        =   26696
            _ExtentY        =   4762
            _StockProps     =   79
            Caption         =   "1. 상품선택"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   6
            Begin XtremeSuiteControls.PushButton Goods 
               Height          =   750
               Index           =   0
               Left            =   45
               TabIndex        =   68
               TabStop         =   0   'False
               Top             =   285
               Width           =   1650
               _Version        =   851970
               _ExtentX        =   2910
               _ExtentY        =   1323
               _StockProps     =   79
               Caption         =   "Goods1"
               BackColor       =   -2147483633
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
               PushButtonStyle =   3
            End
            Begin XtremeSuiteControls.PushButton Goods 
               Height          =   750
               Index           =   1
               Left            =   1725
               TabIndex        =   69
               TabStop         =   0   'False
               Top             =   285
               Width           =   1650
               _Version        =   851970
               _ExtentX        =   2910
               _ExtentY        =   1323
               _StockProps     =   79
               Caption         =   "Goods1"
               BackColor       =   -2147483633
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
               PushButtonStyle =   3
            End
            Begin XtremeSuiteControls.PushButton Goods 
               Height          =   750
               Index           =   2
               Left            =   3405
               TabIndex        =   70
               TabStop         =   0   'False
               Top             =   285
               Width           =   1650
               _Version        =   851970
               _ExtentX        =   2910
               _ExtentY        =   1323
               _StockProps     =   79
               Caption         =   "Goods1"
               BackColor       =   -2147483633
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
               PushButtonStyle =   3
            End
            Begin XtremeSuiteControls.PushButton Goods 
               Height          =   750
               Index           =   3
               Left            =   5085
               TabIndex        =   71
               TabStop         =   0   'False
               Top             =   285
               Width           =   1650
               _Version        =   851970
               _ExtentX        =   2910
               _ExtentY        =   1323
               _StockProps     =   79
               Caption         =   "Goods1"
               BackColor       =   -2147483633
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
               PushButtonStyle =   3
            End
            Begin XtremeSuiteControls.PushButton Goods 
               Height          =   750
               Index           =   4
               Left            =   6765
               TabIndex        =   72
               TabStop         =   0   'False
               Top             =   285
               Width           =   1650
               _Version        =   851970
               _ExtentX        =   2910
               _ExtentY        =   1323
               _StockProps     =   79
               Caption         =   "Goods1"
               BackColor       =   -2147483633
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
               PushButtonStyle =   3
            End
            Begin XtremeSuiteControls.PushButton Goods 
               Height          =   750
               Index           =   5
               Left            =   8445
               TabIndex        =   73
               TabStop         =   0   'False
               Top             =   285
               Width           =   1650
               _Version        =   851970
               _ExtentX        =   2910
               _ExtentY        =   1323
               _StockProps     =   79
               Caption         =   "Goods1"
               BackColor       =   -2147483633
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
               PushButtonStyle =   3
            End
            Begin XtremeSuiteControls.PushButton Goods 
               Height          =   750
               Index           =   6
               Left            =   10125
               TabIndex        =   74
               TabStop         =   0   'False
               Top             =   285
               Width           =   1650
               _Version        =   851970
               _ExtentX        =   2910
               _ExtentY        =   1323
               _StockProps     =   79
               Caption         =   "Goods1"
               BackColor       =   -2147483633
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
               PushButtonStyle =   3
            End
            Begin XtremeSuiteControls.PushButton Goods 
               Height          =   750
               Index           =   7
               Left            =   11805
               TabIndex        =   75
               TabStop         =   0   'False
               Top             =   285
               Width           =   1650
               _Version        =   851970
               _ExtentX        =   2910
               _ExtentY        =   1323
               _StockProps     =   79
               Caption         =   "Goods1"
               BackColor       =   -2147483633
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
               PushButtonStyle =   3
            End
            Begin XtremeSuiteControls.PushButton Goods 
               Height          =   750
               Index           =   8
               Left            =   13485
               TabIndex        =   76
               TabStop         =   0   'False
               Top             =   285
               Width           =   1650
               _Version        =   851970
               _ExtentX        =   2910
               _ExtentY        =   1323
               _StockProps     =   79
               Caption         =   "Goods1"
               BackColor       =   -2147483633
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
               PushButtonStyle =   3
            End
            Begin XtremeSuiteControls.PushButton Goods 
               Height          =   750
               Index           =   9
               Left            =   45
               TabIndex        =   77
               TabStop         =   0   'False
               Top             =   1065
               Width           =   1650
               _Version        =   851970
               _ExtentX        =   2910
               _ExtentY        =   1323
               _StockProps     =   79
               Caption         =   "Goods1"
               BackColor       =   -2147483633
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
               PushButtonStyle =   3
            End
            Begin XtremeSuiteControls.PushButton Goods 
               Height          =   750
               Index           =   10
               Left            =   1725
               TabIndex        =   78
               TabStop         =   0   'False
               Top             =   1065
               Width           =   1650
               _Version        =   851970
               _ExtentX        =   2910
               _ExtentY        =   1323
               _StockProps     =   79
               Caption         =   "Goods1"
               BackColor       =   -2147483633
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
               PushButtonStyle =   3
            End
            Begin XtremeSuiteControls.PushButton Goods 
               Height          =   750
               Index           =   11
               Left            =   3405
               TabIndex        =   79
               TabStop         =   0   'False
               Top             =   1065
               Width           =   1650
               _Version        =   851970
               _ExtentX        =   2910
               _ExtentY        =   1323
               _StockProps     =   79
               Caption         =   "Goods1"
               BackColor       =   -2147483633
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
               PushButtonStyle =   3
            End
            Begin XtremeSuiteControls.PushButton Goods 
               Height          =   750
               Index           =   12
               Left            =   5085
               TabIndex        =   80
               TabStop         =   0   'False
               Top             =   1065
               Width           =   1650
               _Version        =   851970
               _ExtentX        =   2910
               _ExtentY        =   1323
               _StockProps     =   79
               Caption         =   "Goods1"
               BackColor       =   -2147483633
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
               PushButtonStyle =   3
            End
            Begin XtremeSuiteControls.PushButton Goods 
               Height          =   750
               Index           =   13
               Left            =   6765
               TabIndex        =   81
               TabStop         =   0   'False
               Top             =   1065
               Width           =   1650
               _Version        =   851970
               _ExtentX        =   2910
               _ExtentY        =   1323
               _StockProps     =   79
               Caption         =   "Goods1"
               BackColor       =   -2147483633
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
               PushButtonStyle =   3
            End
            Begin XtremeSuiteControls.PushButton Goods 
               Height          =   750
               Index           =   14
               Left            =   8445
               TabIndex        =   82
               TabStop         =   0   'False
               Top             =   1065
               Width           =   1650
               _Version        =   851970
               _ExtentX        =   2910
               _ExtentY        =   1323
               _StockProps     =   79
               Caption         =   "Goods1"
               BackColor       =   -2147483633
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
               PushButtonStyle =   3
            End
            Begin XtremeSuiteControls.PushButton Goods 
               Height          =   750
               Index           =   15
               Left            =   10125
               TabIndex        =   83
               TabStop         =   0   'False
               Top             =   1065
               Width           =   1650
               _Version        =   851970
               _ExtentX        =   2910
               _ExtentY        =   1323
               _StockProps     =   79
               Caption         =   "Goods1"
               BackColor       =   -2147483633
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
               PushButtonStyle =   3
            End
            Begin XtremeSuiteControls.PushButton Goods 
               Height          =   750
               Index           =   16
               Left            =   11805
               TabIndex        =   84
               TabStop         =   0   'False
               Top             =   1065
               Width           =   1650
               _Version        =   851970
               _ExtentX        =   2910
               _ExtentY        =   1323
               _StockProps     =   79
               Caption         =   "Goods1"
               BackColor       =   -2147483633
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
               PushButtonStyle =   3
            End
            Begin XtremeSuiteControls.PushButton Goods 
               Height          =   750
               Index           =   17
               Left            =   13485
               TabIndex        =   85
               TabStop         =   0   'False
               Top             =   1065
               Width           =   1650
               _Version        =   851970
               _ExtentX        =   2910
               _ExtentY        =   1323
               _StockProps     =   79
               Caption         =   "Goods1"
               BackColor       =   -2147483633
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
               PushButtonStyle =   3
            End
            Begin XtremeSuiteControls.PushButton Goods 
               Height          =   750
               Index           =   18
               Left            =   45
               TabIndex        =   86
               TabStop         =   0   'False
               Top             =   1845
               Width           =   1650
               _Version        =   851970
               _ExtentX        =   2910
               _ExtentY        =   1323
               _StockProps     =   79
               Caption         =   "Goods1"
               BackColor       =   -2147483633
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
               PushButtonStyle =   3
            End
            Begin XtremeSuiteControls.PushButton Goods 
               Height          =   750
               Index           =   19
               Left            =   1725
               TabIndex        =   87
               TabStop         =   0   'False
               Top             =   1845
               Width           =   1650
               _Version        =   851970
               _ExtentX        =   2910
               _ExtentY        =   1323
               _StockProps     =   79
               Caption         =   "Goods1"
               BackColor       =   -2147483633
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
               PushButtonStyle =   3
            End
            Begin XtremeSuiteControls.PushButton Goods 
               Height          =   750
               Index           =   20
               Left            =   3405
               TabIndex        =   88
               TabStop         =   0   'False
               Top             =   1845
               Width           =   1650
               _Version        =   851970
               _ExtentX        =   2910
               _ExtentY        =   1323
               _StockProps     =   79
               Caption         =   "Goods1"
               BackColor       =   -2147483633
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
               PushButtonStyle =   3
            End
            Begin XtremeSuiteControls.PushButton Goods 
               Height          =   750
               Index           =   21
               Left            =   5085
               TabIndex        =   89
               TabStop         =   0   'False
               Top             =   1845
               Width           =   1650
               _Version        =   851970
               _ExtentX        =   2910
               _ExtentY        =   1323
               _StockProps     =   79
               Caption         =   "Goods1"
               BackColor       =   -2147483633
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
               PushButtonStyle =   3
            End
            Begin XtremeSuiteControls.PushButton Goods 
               Height          =   750
               Index           =   22
               Left            =   6765
               TabIndex        =   90
               TabStop         =   0   'False
               Top             =   1845
               Width           =   1650
               _Version        =   851970
               _ExtentX        =   2910
               _ExtentY        =   1323
               _StockProps     =   79
               Caption         =   "Goods1"
               BackColor       =   -2147483633
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
               PushButtonStyle =   3
            End
            Begin XtremeSuiteControls.PushButton Goods 
               Height          =   750
               Index           =   23
               Left            =   8445
               TabIndex        =   91
               TabStop         =   0   'False
               Top             =   1845
               Width           =   1650
               _Version        =   851970
               _ExtentX        =   2910
               _ExtentY        =   1323
               _StockProps     =   79
               Caption         =   "소품기타(부속품/모자/장갑)"
               BackColor       =   -2147483633
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
               PushButtonStyle =   3
            End
            Begin XtremeSuiteControls.PushButton Goods 
               Height          =   750
               Index           =   24
               Left            =   10125
               TabIndex        =   92
               TabStop         =   0   'False
               Top             =   1845
               Width           =   1650
               _Version        =   851970
               _ExtentX        =   2910
               _ExtentY        =   1323
               _StockProps     =   79
               Caption         =   "Goods1"
               BackColor       =   -2147483633
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
               PushButtonStyle =   3
            End
            Begin XtremeSuiteControls.PushButton Goods 
               Height          =   750
               Index           =   25
               Left            =   11805
               TabIndex        =   93
               TabStop         =   0   'False
               Top             =   1845
               Width           =   1650
               _Version        =   851970
               _ExtentX        =   2910
               _ExtentY        =   1323
               _StockProps     =   79
               Caption         =   "Goods1"
               BackColor       =   -2147483633
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
               PushButtonStyle =   3
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox2 
            Height          =   5235
            Left            =   0
            TabIndex        =   94
            Top             =   2910
            Width           =   11370
            _Version        =   851970
            _ExtentX        =   20055
            _ExtentY        =   9234
            _StockProps     =   79
            Caption         =   "2. 주문내역"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   6
            Begin HoonControl.MyLabel lblFirstTag 
               Height          =   855
               Left            =   4118
               Top             =   2490
               Width           =   3135
               _ExtentX        =   5530
               _ExtentY        =   1508
               Caption         =   "000-00-0000"
               ForeColor       =   192
               SelForeColor    =   0
               UnSelForeColor  =   0
               EditForeColor   =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   18
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontSize        =   18
               FontBold        =   -1  'True
               FontItalic      =   0   'False
               FontUnder       =   0   'False
               Margine         =   0
               BackColor       =   12648447
            End
            Begin FPSpreadADO.fpSpread sprGrid 
               Height          =   4515
               Left            =   60
               TabIndex        =   95
               TabStop         =   0   'False
               Top             =   660
               Width           =   11250
               _Version        =   524288
               _ExtentX        =   19844
               _ExtentY        =   7964
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
               MaxCols         =   22
               MaxRows         =   9
               OperationMode   =   1
               ScrollBars      =   2
               SpreadDesigner  =   "frmAccept.frx":21A2
               UserResize      =   1
               VisibleCols     =   7
               HighlightHeaders=   1
               HighlightStyle  =   1
               ScrollBarStyle  =   2
            End
            Begin XtremeSuiteControls.PushButton btnClear 
               Height          =   420
               Left            =   10275
               TabIndex        =   116
               TabStop         =   0   'False
               Top             =   210
               Width           =   1035
               _Version        =   851970
               _ExtentX        =   1826
               _ExtentY        =   741
               _StockProps     =   79
               Caption         =   " 취소"
               BackColor       =   -2147483633
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
               Picture         =   "frmAccept.frx":3B32
            End
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage7 
         Height          =   8190
         Left            =   -69970
         TabIndex        =   100
         Top             =   570
         Visible         =   0   'False
         Width           =   15210
         _Version        =   851970
         _ExtentX        =   26829
         _ExtentY        =   14446
         _StockProps     =   1
         BackColor       =   -2147483633
         Page            =   1
         Begin XtremeSuiteControls.GroupBox GroupBox6 
            Height          =   8055
            Left            =   11445
            TabIndex        =   101
            Top             =   60
            Width           =   3735
            _Version        =   851970
            _ExtentX        =   6588
            _ExtentY        =   14208
            _StockProps     =   79
            Caption         =   "2. 결제진행"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   6
            Begin XtremeSuiteControls.FlatEdit FlatEdit9 
               Height          =   315
               Left            =   120
               TabIndex        =   106
               TabStop         =   0   'False
               Top             =   3210
               Width           =   3525
               _Version        =   851970
               _ExtentX        =   6218
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   12
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "FlatEdit1"
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   1
               FlatStyle       =   -1  'True
               UseVisualStyle  =   0   'False
               ShowBorder      =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit10 
               Height          =   315
               Left            =   120
               TabIndex        =   107
               TabStop         =   0   'False
               Top             =   2535
               Width           =   3525
               _Version        =   851970
               _ExtentX        =   6218
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   12
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "FlatEdit1"
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   1
               FlatStyle       =   -1  'True
               UseVisualStyle  =   0   'False
               ShowBorder      =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit11 
               Height          =   315
               Left            =   120
               TabIndex        =   108
               TabStop         =   0   'False
               Top             =   1860
               Width           =   3525
               _Version        =   851970
               _ExtentX        =   6218
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   12
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "FlatEdit1"
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   1
               FlatStyle       =   -1  'True
               UseVisualStyle  =   0   'False
               ShowBorder      =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit12 
               Height          =   315
               Left            =   120
               TabIndex        =   109
               TabStop         =   0   'False
               Top             =   1185
               Width           =   3525
               _Version        =   851970
               _ExtentX        =   6218
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   12
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "FlatEdit1"
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   1
               FlatStyle       =   -1  'True
               UseVisualStyle  =   0   'False
               ShowBorder      =   0   'False
            End
            Begin XtremeSuiteControls.PushButton PushButton5 
               Height          =   825
               Left            =   90
               TabIndex        =   110
               TabStop         =   0   'False
               Top             =   3600
               Width           =   3585
               _Version        =   851970
               _ExtentX        =   6324
               _ExtentY        =   1455
               _StockProps     =   79
               Caption         =   "접수결제"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   14.25
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
            End
            Begin Threed.SSPanel SSPanel7 
               Height          =   525
               Index           =   7
               Left            =   105
               TabIndex        =   111
               Top             =   315
               Width           =   3555
               _ExtentX        =   6271
               _ExtentY        =   926
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
               Caption         =   "합계내역"
               PictureBackgroundStyle=   2
               PictureBackground=   "frmAccept.frx":3F08
               BorderWidth     =   0
               BevelOuter      =   0
               PictureAlignment=   9
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin XtremeSuiteControls.PushButton PushButton6 
               Height          =   825
               Left            =   90
               TabIndex        =   112
               TabStop         =   0   'False
               Top             =   4425
               Width           =   3585
               _Version        =   851970
               _ExtentX        =   6324
               _ExtentY        =   1455
               _StockProps     =   79
               Caption         =   "후불"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   14.25
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
            End
            Begin VB.TextBox Text1 
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
               Height          =   690
               Left            =   105
               Locked          =   -1  'True
               TabIndex        =   105
               TabStop         =   0   'False
               Text            =   "건수"
               Top             =   825
               Width           =   3555
            End
            Begin VB.TextBox Text12 
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
               Height          =   690
               Left            =   105
               Locked          =   -1  'True
               TabIndex        =   102
               TabStop         =   0   'False
               Text            =   "합계금액"
               Top             =   2850
               Width           =   3555
            End
            Begin VB.TextBox Text11 
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
               Height          =   690
               Left            =   105
               Locked          =   -1  'True
               TabIndex        =   103
               TabStop         =   0   'False
               Text            =   "할인금액"
               Top             =   2175
               Width           =   3555
            End
            Begin VB.TextBox Text10 
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
               Height          =   690
               Left            =   105
               Locked          =   -1  'True
               TabIndex        =   104
               TabStop         =   0   'False
               Text            =   "접수금액"
               Top             =   1500
               Width           =   3555
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox7 
            Height          =   8070
            Left            =   30
            TabIndex        =   113
            Top             =   75
            Width           =   11370
            _Version        =   851970
            _ExtentX        =   20055
            _ExtentY        =   14235
            _StockProps     =   79
            Caption         =   "1. 주문내역"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   6
            Begin FPSpreadADO.fpSpread fpSpread2 
               Height          =   7695
               Left            =   60
               TabIndex        =   114
               TabStop         =   0   'False
               Top             =   315
               Width           =   11250
               _Version        =   524288
               _ExtentX        =   19844
               _ExtentY        =   13573
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
               SpreadDesigner  =   "frmAccept.frx":424A
               UserResize      =   1
               VisibleCols     =   7
               VisibleRows     =   30
               HighlightHeaders=   1
               HighlightStyle  =   1
               ScrollBarStyle  =   2
            End
         End
      End
   End
   Begin Threed.SSPanel SSPanel7 
      Height          =   420
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   90
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
      Caption         =   "고객조회"
      PictureBackgroundStyle=   2
      PictureBackground=   "frmAccept.frx":5D68
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
      TabIndex        =   4
      Top             =   555
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
      PictureBackground=   "frmAccept.frx":60AA
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
      TabIndex        =   5
      Top             =   990
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
      PictureBackground=   "frmAccept.frx":63EC
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
      Left            =   3000
      TabIndex        =   6
      Top             =   555
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
      PictureBackground=   "frmAccept.frx":672E
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
      Left            =   5940
      TabIndex        =   7
      Top             =   555
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
      PictureBackground=   "frmAccept.frx":6A70
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
      TabIndex        =   9
      Top             =   1425
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
      PictureBackgroundStyle=   3
      PictureBackground=   "frmAccept.frx":6DB2
      BorderWidth     =   0
      BevelOuter      =   0
      PictureAlignment=   9
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      FloodShowPct    =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdEdit 
      Height          =   420
      Left            =   7830
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   90
      Width           =   1035
      _Version        =   851970
      _ExtentX        =   1826
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   " 수정"
      BackColor       =   -2147483633
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
      Picture         =   "frmAccept.frx":70F4
   End
   Begin XtremeSuiteControls.PushButton cmdNew 
      Height          =   420
      Left            =   6780
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   90
      Width           =   1035
      _Version        =   851970
      _ExtentX        =   1826
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   " 신규"
      BackColor       =   -2147483633
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
      Picture         =   "frmAccept.frx":7B06
   End
   Begin XtremeSuiteControls.TabControl TabControl2 
      Height          =   2160
      Left            =   8955
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   90
      Width           =   6375
      _Version        =   851970
      _ExtentX        =   11245
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
      Item(0).Control(0)=   "TabControlPage6"
      Item(1).Caption =   "실적"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "TabControlPage5"
      Item(2).Caption =   "사고"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "TabControlPage4"
      Begin XtremeSuiteControls.TabControlPage TabControlPage4 
         Height          =   2100
         Left            =   -69370
         TabIndex        =   27
         Top             =   30
         Visible         =   0   'False
         Width           =   5715
         _Version        =   851970
         _ExtentX        =   10081
         _ExtentY        =   3704
         _StockProps     =   1
         BackColor       =   255
         Page            =   2
         Begin FPSpreadADO.fpSpread sprClaim 
            Bindings        =   "frmAccept.frx":8518
            Height          =   1935
            Left            =   75
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   75
            Width           =   5565
            _Version        =   524288
            _ExtentX        =   9816
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
            SpreadDesigner  =   "frmAccept.frx":852C
            VisibleCols     =   9
            VisibleRows     =   200
            Appearance      =   1
            HighlightHeaders=   1
            HighlightStyle  =   1
            ScrollBarStyle  =   2
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage5 
         Height          =   2100
         Left            =   -69370
         TabIndex        =   29
         Top             =   30
         Visible         =   0   'False
         Width           =   5715
         _Version        =   851970
         _ExtentX        =   10081
         _ExtentY        =   3704
         _StockProps     =   1
         BackColor       =   -2147483633
         Page            =   1
         Begin FPSpreadADO.fpSpread sprHist 
            Height          =   1680
            Left            =   3270
            TabIndex        =   30
            TabStop         =   0   'False
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
            SpreadDesigner  =   "frmAccept.frx":8F16
            HighlightHeaders=   1
            HighlightStyle  =   1
         End
         Begin FPSpreadADO.fpSpread sprYear 
            Height          =   1680
            Left            =   75
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   330
            Width           =   3150
            _Version        =   524288
            _ExtentX        =   5556
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
            SpreadDesigner  =   "frmAccept.frx":947D
            HighlightHeaders=   1
            HighlightStyle  =   1
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
            Left            =   3255
            TabIndex        =   33
            Top             =   105
            Width           =   1170
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
            TabIndex        =   32
            Top             =   105
            Width           =   1470
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage6 
         Height          =   2100
         Left            =   630
         TabIndex        =   34
         Top             =   30
         Width           =   5715
         _Version        =   851970
         _ExtentX        =   10081
         _ExtentY        =   3704
         _StockProps     =   1
         BackColor       =   -2147483633
         Page            =   0
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
            TabIndex        =   36
            TabStop         =   0   'False
            Text            =   "2010-12-31"
            Top             =   75
            Width           =   1305
         End
         Begin XtremeSuiteControls.PushButton btnMisu 
            Height          =   375
            Left            =   5235
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   75
            Width           =   390
            _Version        =   851970
            _ExtentX        =   688
            _ExtentY        =   661
            _StockProps     =   79
            BackColor       =   -2147483633
            Appearance      =   6
            Picture         =   "frmAccept.frx":9A52
         End
         Begin CSTextLibCtl.sidbEdit txtMisu 
            Height          =   375
            Left            =   3900
            TabIndex        =   37
            TabStop         =   0   'False
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
            Left            =   3900
            TabIndex        =   38
            TabStop         =   0   'False
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
            Left            =   3900
            TabIndex        =   39
            TabStop         =   0   'False
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
            TabIndex        =   40
            TabStop         =   0   'False
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
            TabIndex        =   41
            TabStop         =   0   'False
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
            TabIndex        =   42
            TabStop         =   0   'False
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
            Left            =   5235
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   1245
            Width           =   390
            _Version        =   851970
            _ExtentX        =   688
            _ExtentY        =   661
            _StockProps     =   79
            BackColor       =   -2147483633
            Appearance      =   6
            Picture         =   "frmAccept.frx":A464
         End
         Begin XtremeSuiteControls.PushButton btnTagCode 
            Height          =   375
            Left            =   3900
            TabIndex        =   44
            TabStop         =   0   'False
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
            Left            =   4560
            TabIndex        =   45
            TabStop         =   0   'False
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
            TabIndex        =   46
            TabStop         =   0   'False
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
            Left            =   3900
            TabIndex        =   47
            TabStop         =   0   'False
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
            Caption         =   "총매출액:"
            Height          =   180
            Index           =   1
            Left            =   90
            TabIndex        =   57
            Top             =   960
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
            TabIndex        =   56
            Top             =   1350
            Width           =   810
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "총할인액:"
            Height          =   180
            Index           =   3
            Left            =   90
            TabIndex        =   55
            Top             =   1740
            Width           =   810
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "누적 마일리지:"
            Height          =   240
            Index           =   7
            Left            =   2205
            TabIndex        =   54
            Top             =   1365
            Width           =   1650
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "가능 마일리지:"
            Height          =   240
            Index           =   5
            Left            =   2205
            TabIndex        =   53
            Top             =   960
            Width           =   1650
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "미수금액:"
            Height          =   180
            Index           =   4
            Left            =   3045
            TabIndex        =   52
            Top             =   180
            Width           =   810
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "등록일자:"
            Height          =   180
            Index           =   0
            Left            =   90
            TabIndex        =   51
            Top             =   180
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
            Left            =   3045
            TabIndex        =   50
            Top             =   1755
            Width           =   810
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "이용횟수:"
            Height          =   180
            Index           =   6
            Left            =   90
            TabIndex        =   49
            Top             =   570
            Width           =   810
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "미환불금액:"
            Height          =   180
            Index           =   8
            Left            =   2865
            TabIndex        =   48
            Top             =   570
            Visible         =   0   'False
            Width           =   990
         End
      End
   End
   Begin CSTextLibCtl.sitxEdit txtTel 
      Height          =   420
      Left            =   4065
      TabIndex        =   97
      TabStop         =   0   'False
      Top             =   555
      Width           =   1860
      _Version        =   262145
      _ExtentX        =   3281
      _ExtentY        =   741
      _StockProps     =   125
      Text            =   "010-9858-0428"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   11.26
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ReadOnly        =   -1  'True
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   "010-9858-0428"
      Text            =   "010-9858-0428"
      StartText.x     =   2
      StartText.y     =   5
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   17
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      CharacterTable  =   ""
   End
End
Attribute VB_Name = "frmAccept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClear_Click()
    If sprGrid.MaxRows > 0 Then
        If MsgBox("모든 주문내역이 취소됩니다." & vbCrLf & "진행 하시겠습니까?", vbCritical + vbYesNo, "접수") = vbNo Then Exit Sub
        sprGrid.MaxRows = 0
        Call CheckSummary
        Call ResetTagNo
    End If
End Sub

Private Sub Form_Load()
    GetGoodsTitle
    sprGrid.MaxRows = 0
    lblFirstTag.Caption = btnTagCode.Caption & "-" & cmdTagNo.Caption
    lblFirstTag.Tag = cmdTagNo.Caption
End Sub

Private Sub Goods_Click(Index As Integer)
    Call Sub_의류가격정보(Goods(Index).Tag)
    sprGrid.SetFocus
End Sub

Private Sub Goods_DropDown(Index As Integer)
    Goods(Index).Enabled = False
    Call frmGoods.GetData(Goods(Index).Tag)
    frmGoods.Show vbModal
    Goods(Index).Enabled = True
    sprGrid.SetFocus
End Sub


' 선택한 종류의 의류명 , 택번호, "드", 금액, 코드를 등록 한다.
Public Sub Sub_의류가격정보(의류코드 As String)
    Dim ADORs       As ADODB.RecordSet
    Dim sGoodsStats As String
    Dim iCol       As Integer  '
    Dim strNum1    As String   '택번호1
    Dim intNum1    As Integer  'tagno1
    Dim intNum2    As Integer  'tagno2
    
    Dim intCol01   As Integer  '
    Dim iActrow    As Integer  '
    Dim iPrice     As Long     '
    Dim OriginPrice As Long
    
    Dim 의류명     As String
    
    Dim iEOF       As Boolean
    
    
    Set ADORs = New ADODB.RecordSet
    Set ADORs = Get_의류정보(의류코드, sGoodsStats)
                      
    If ADORs.EOF Then
        ADORs.Close:    Set ADORs = Nothing
        MsgBox "등록되지 않은 품목입니다 ", vbCritical, "확인"
        Exit Sub
    End If
     
    'frmAccept.lblGoodsPriceStats.Caption = sGoodsStats
    의류코드 = ADORs!의류코드 & ""
    의류명 = ADORs!의류명 & ""
    
    ADORs.Close:    Set ADORs = Nothing
     
     ' 새로 입력할 라인을 구한다.
    iActrow = frmAccept.sprGrid.ActiveRow
    sprGrid.MaxRows = sprGrid.MaxRows + 1
    iCur = sprGrid.MaxRows
    
   
    i = 1
    iCol = 1
    
    
    iPrice = Get_Price(의류코드)
    OriginPrice = Get_세탁정상금액(의류코드)
    With frmAccept
        
        .sprGrid.Row = iCur
        .sprGrid.Col = 2:  .sprGrid.Text = Trim(의류명) & ""  ' 1 의류명
        .sprGrid.Col = 4:  .sprGrid.Text = "흰색"             ' 3 색상
        .sprGrid.Col = 5:  .sprGrid.Text = "없음"             ' 4 무늬
        

        .sprGrid.Col = 6: .sprGrid.Text = "세"            ' 5 작업 * 수선접수 여부 *

        
        .sprGrid.Col = 7: .sprGrid.Value = OriginPrice & ""            '20 ** 원래 금액 **
        .sprGrid.Col = 8:  .sprGrid.Value = iPrice & ""       ' 6 금액
        
        If OriginPrice <> iPrice Then
            .sprGrid.ForeColor = vbRed
        Else
            .sprGrid.ForeColor = vbBlack
        End If
        
        .sprGrid.Col = 10:  .sprGrid.Value = 의류코드 & ""     ' 8 의류코드
        .sprGrid.Col = 16: .sprGrid.Value = iPrice & ""       '14 세트 상품의 원 금액을 기록한다.
        
        '------------------------------------------------------------
        ' 마진 정보
        '------------------------------------------------------------
        Query = "SELECT    ISNULL(세탁마진,0) AS 세탁마진"
        Query = Query & ", ISNULL(외주마진,0) AS 외주마진"
        Query = Query & ", ISNULL(수선마진,0) AS 수선마진"
        Query = Query & " FROM TB_의류분류"
        Query = Query & " WHERE 의류분류코드 = '" & Left(의류코드, 2) & "'"
        Set SUBRs = New ADODB.RecordSet
        SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        If SUBRs.EOF Then
            .sprGrid.Col = 18: .sprGrid.Value = 0                   '16
            .sprGrid.Col = 19: .sprGrid.Value = 0                   '17
            .sprGrid.Col = 20: .sprGrid.Value = 0                   '18
        Else
            .sprGrid.Col = 18: .sprGrid.Value = SUBRs!세탁마진 & "" '16
            .sprGrid.Col = 19: .sprGrid.Value = SUBRs!외주마진 & "" '17
            .sprGrid.Col = 20: .sprGrid.Value = SUBRs!수선마진 & "" '18
        End If
        SUBRs.Close
        Set SUBRs = Nothing
        
        
    End With
        
    '---------------------------------------------------------------------
    ' 택 번호 출력
    '---------------------------------------------------------------------

    strNum1 = frmAccept.cmdTagNo.Caption '

    If Len(Trim(Get_SpreadText(frmAccept.sprGrid, CDbl(iCur), 3))) <= 0 Then
        frmAccept.sprGrid.Row = iCur
        frmAccept.sprGrid.Col = 3: frmAccept.sprGrid.Text = strNum1 '택번호
        frmAccept.cmdTagNo.Caption = Get_ChangeTagNo(strNum1, "+")     '
    End If
    
    frmAccept.sprGrid.Row = iCur
    frmAccept.sprGrid.BackColor = vbWhite
    frmAccept.sprGrid.SetActiveCell 3, iCur
    CheckSummary
    DoEvents

End Sub


' 전달된 코드의 금액을 DB에서 읽어온다.
Private Function Get_Price(ClothCode As String) As Long
    Dim iPrice      As Double
    Dim sGoodsStats As String
    
    On Error GoTo ErrRtn
      
    iPrice = Get_세탁금액(ClothCode, sGoodsStats)
    
    If iPrice = -1 Then
        Get_Price = 0
        
        Exit Function
    End If
    
    Get_Price = iPrice
    
    Exit Function
          
ErrRtn:
    Get_Price = 0
End Function

Private Sub CheckSummary()
    Dim Total_Origin_Price As Long
    Dim Total_Sale_Price As Long
    With sprGrid
        Dim LoopI As Integer
        For LoopI = 1 To .MaxRows
        .Row = LoopI
        .Col = 7: Total_Origin_Price = Total_Origin_Price + CLng(.Text)
        .Col = 8: Total_Sale_Price = Total_Sale_Price + CLng(.Text)
        Next LoopI
        lblOriginPrice.Caption = FormatNumber(Total_Origin_Price, 0, vbTrue)
        lblDiscountPrice.Caption = FormatNumber(Total_Origin_Price - Total_Sale_Price, 0, vbTrue)
        lblSalePrice.Caption = FormatNumber(Total_Sale_Price, 0, vbTrue)
        lblCount.Caption = FormatNumber(LoopI - 1, 0, vbTrue)
    End With
    If sprGrid.MaxRows = 0 Then
        lblFirstTag.Visible = True
    Else
        lblFirstTag.Visible = False
    End If
End Sub


Private Sub sprGrid_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    Select Case Col
    Case 1:
        With sprGrid
            .DeleteRows Row, 1
            .MaxRows = .MaxRows - 1
        End With
        Call ResetTagNo
    Case 4, 5:
        Dim Pattern As String
        Dim Color As String
        With sprGrid
            .Row = Row
            .Col = 4: Color = .Text
            .Col = 5: Pattern = .Text
        End With
        
        Call frmGoodsDetail.GetData(Color, Pattern)
        frmGoodsDetail.Show vbModal
    End Select
    Call CheckSummary
End Sub

Public Sub SetDetail(Pattern As String, Color As String, ColorCode As String)
    With sprGrid
        .Row = .ActiveRow
        .Col = 4: .Text = Color: .BackColor = ColorCode
        .Col = 5: .Text = Pattern
    End With
End Sub

Private Sub ResetTagNo()
    Dim LoopI As Long
    Dim TagNo As String
    With sprGrid
        frmAccept.cmdTagNo.Caption = lblFirstTag.Tag
        For LoopI = 1 To .MaxRows
            TagNo = frmAccept.cmdTagNo.Caption '
            
            frmAccept.sprGrid.Row = LoopI
            frmAccept.sprGrid.Col = 3: frmAccept.sprGrid.Text = TagNo '택번호
            frmAccept.cmdTagNo.Caption = Get_ChangeTagNo(TagNo, "+")     '
        
        Next LoopI
    End With
End Sub


