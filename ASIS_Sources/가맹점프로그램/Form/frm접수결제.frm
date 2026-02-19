VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm접수결제 
   BorderStyle     =   1  '단일 고정
   Caption         =   "접수결제"
   ClientHeight    =   8550
   ClientLeft      =   7815
   ClientTop       =   3300
   ClientWidth     =   10965
   ControlBox      =   0   'False
   DrawWidth       =   3
   FillColor       =   &H00C0C0C0&
   Icon            =   "frm접수결제.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   10965
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8550
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   15081
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm접수결제.frx":0A02
      Begin Threed.SSPanel SSPanel1 
         Height          =   780
         Left            =   15
         TabIndex        =   1
         Top             =   7755
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdAction 
            Height          =   660
            Index           =   0
            Left            =   60
            TabIndex        =   2
            Top             =   45
            Width           =   1590
            _Version        =   851970
            _ExtentX        =   2805
            _ExtentY        =   1164
            _StockProps     =   79
            Caption         =   " 완불결제"
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
            Picture         =   "frm접수결제.frx":0B14
         End
         Begin XtremeSuiteControls.PushButton cmdAction 
            Height          =   660
            Index           =   1
            Left            =   1770
            TabIndex        =   3
            Top             =   60
            Width           =   1590
            _Version        =   851970
            _ExtentX        =   2805
            _ExtentY        =   1164
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
         Begin XtremeSuiteControls.PushButton cmdAction 
            Height          =   660
            Index           =   2
            Left            =   3495
            TabIndex        =   4
            Top             =   60
            Width           =   1590
            _Version        =   851970
            _ExtentX        =   2805
            _ExtentY        =   1164
            _StockProps     =   79
            Caption         =   "부분결제"
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
         End
         Begin XtremeSuiteControls.PushButton btnExit 
            Height          =   660
            Left            =   9285
            TabIndex        =   6
            Top             =   60
            Width           =   1590
            _Version        =   851970
            _ExtentX        =   2805
            _ExtentY        =   1164
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
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
            Picture         =   "frm접수결제.frx":120E
         End
         Begin XtremeSuiteControls.PushButton cmdAction 
            Height          =   660
            Index           =   3
            Left            =   60
            TabIndex        =   56
            Top             =   45
            Width           =   1590
            _Version        =   851970
            _ExtentX        =   2805
            _ExtentY        =   1164
            _StockProps     =   79
            Caption         =   " 인터넷접수"
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
            Picture         =   "frm접수결제.frx":17A8
         End
      End
      Begin FPSpreadADO.fpSpread sprCard 
         Height          =   3150
         Left            =   5010
         TabIndex        =   5
         Top             =   465
         Width           =   5940
         _Version        =   524288
         _ExtentX        =   10478
         _ExtentY        =   5556
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
         SpreadDesigner  =   "frm접수결제.frx":1EA2
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   435
         Index           =   1
         Left            =   5010
         TabIndex        =   7
         Top             =   15
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   767
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
         PictureBackground=   "frm접수결제.frx":2821
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   435
         Index           =   0
         Left            =   5010
         TabIndex        =   8
         Top             =   3630
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   767
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
         PictureBackground=   "frm접수결제.frx":2A47
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton btnCashCancel 
            Height          =   405
            Left            =   3960
            TabIndex        =   10
            Top             =   15
            Width           =   1965
            _Version        =   851970
            _ExtentX        =   3466
            _ExtentY        =   714
            _StockProps     =   79
            Caption         =   "현금영수증 승인취소"
            ForeColor       =   192
            UseVisualStyle  =   -1  'True
         End
      End
      Begin FPSpreadADO.fpSpread sprCash 
         Height          =   3660
         Left            =   5010
         TabIndex        =   9
         Top             =   4080
         Width           =   5940
         _Version        =   524288
         _ExtentX        =   10478
         _ExtentY        =   6456
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayColHeaders=   0   'False
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
         MaxCols         =   1
         MaxRows         =   11
         RowHeaderDisplay=   0
         ScrollBars      =   0
         SpreadDesigner  =   "frm접수결제.frx":2C6D
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   4590
         Left            =   15
         TabIndex        =   11
         Top             =   15
         Width           =   4980
         _ExtentX        =   8784
         _ExtentY        =   8096
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin CSTextLibCtl.silgEdit txtTotalPay2 
            Height          =   420
            Left            =   1605
            TabIndex        =   24
            Top             =   1875
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
               Weight          =   400
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
            Index           =   6
            Left            =   45
            TabIndex        =   25
            Top             =   1875
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
            Caption         =   "할인전 금액"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm접수결제.frx":32B3
            BorderWidth     =   0
            BevelOuter      =   1
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnCoupon 
            Height          =   420
            Left            =   3195
            TabIndex        =   12
            Top             =   3405
            Visible         =   0   'False
            Width           =   1740
            _Version        =   851970
            _ExtentX        =   3069
            _ExtentY        =   741
            _StockProps     =   79
            Caption         =   "쿠폰 사용"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnEvent 
            Height          =   420
            Left            =   3210
            TabIndex        =   55
            Top             =   2325
            Width           =   1740
            _Version        =   851970
            _ExtentX        =   3069
            _ExtentY        =   741
            _StockProps     =   79
            Caption         =   "할인행사 입력"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin VB.ComboBox cboDay 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1605
            Style           =   2  '드롭다운 목록
            TabIndex        =   16
            Top             =   60
            Width           =   1590
         End
         Begin VB.TextBox txtCouponNo 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   3210
            TabIndex        =   13
            Top             =   3195
            Visible         =   0   'False
            Width           =   1680
         End
         Begin CSTextLibCtl.silgEdit txtTotalPay 
            Height          =   420
            Left            =   1605
            TabIndex        =   14
            Top             =   960
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
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   390
            Left            =   1605
            TabIndex        =   15
            Top             =   480
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   688
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   56688641
            CurrentDate     =   40499
         End
         Begin XtremeSuiteControls.PushButton cmdSamSungCard 
            Height          =   420
            Left            =   3195
            TabIndex        =   17
            Top             =   4110
            Visible         =   0   'False
            Width           =   1740
            _Version        =   851970
            _ExtentX        =   3069
            _ExtentY        =   741
            _StockProps     =   79
            Caption         =   "삼성카드할인(10%)"
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   390
            Index           =   0
            Left            =   45
            TabIndex        =   18
            Top             =   480
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
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
            Caption         =   "예 정 일 자"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm접수결제.frx":35F5
            BorderWidth     =   0
            BevelOuter      =   1
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   420
            Index           =   2
            Left            =   45
            TabIndex        =   19
            Top             =   960
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
            Caption         =   "접 수 금 액"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm접수결제.frx":3937
            BorderWidth     =   0
            BevelOuter      =   1
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   420
            Index           =   3
            Left            =   45
            TabIndex        =   20
            Top             =   1410
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
            Caption         =   "마 일 리 지"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm접수결제.frx":3C79
            BorderWidth     =   0
            BevelOuter      =   1
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   420
            Index           =   4
            Left            =   45
            TabIndex        =   21
            Top             =   2325
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
            Caption         =   "할인행사 금액"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm접수결제.frx":3FBB
            BorderWidth     =   0
            BevelOuter      =   1
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.silgEdit txtMileage 
            Height          =   420
            Left            =   1605
            TabIndex        =   22
            Top             =   1410
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
               Weight          =   400
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
         Begin CSTextLibCtl.silgEdit txtCoupon 
            Height          =   420
            Left            =   1605
            TabIndex        =   23
            Top             =   2325
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   420
            Index           =   7
            Left            =   45
            TabIndex        =   26
            Top             =   2775
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
            Caption         =   "할인 금액"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm접수결제.frx":42FD
            BorderWidth     =   0
            BevelOuter      =   1
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   420
            Index           =   8
            Left            =   45
            TabIndex        =   27
            Top             =   3225
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
            Caption         =   "에누리 금액"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm접수결제.frx":463F
            BorderWidth     =   0
            BevelOuter      =   1
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.silgEdit txtSetDC 
            Height          =   420
            Left            =   1605
            TabIndex        =   28
            Top             =   2775
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
               Weight          =   400
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
         Begin CSTextLibCtl.silgEdit txtDC 
            Height          =   420
            Left            =   1605
            TabIndex        =   29
            Top             =   3225
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
               Weight          =   400
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
         Begin CSTextLibCtl.silgEdit txtDCTotal 
            Height          =   420
            Left            =   1605
            TabIndex        =   30
            Top             =   3675
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
               Weight          =   400
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
            Index           =   9
            Left            =   45
            TabIndex        =   31
            Top             =   3675
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
            Caption         =   "할인합계 금액"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm접수결제.frx":4981
            BorderWidth     =   0
            BevelOuter      =   1
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.silgEdit txtBalance 
            Height          =   420
            Left            =   1605
            TabIndex        =   46
            Top             =   4125
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
            Index           =   1
            Left            =   45
            TabIndex        =   47
            Top             =   4125
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
            Caption         =   "잔      액"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm접수결제.frx":4CC3
            BorderWidth     =   0
            BevelOuter      =   1
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   390
            Index           =   5
            Left            =   45
            TabIndex        =   48
            Top             =   60
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
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
            Caption         =   "세탁 소요일"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm접수결제.frx":5005
            BorderWidth     =   0
            BevelOuter      =   1
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnEditDay 
            Height          =   390
            Left            =   3195
            TabIndex        =   53
            Top             =   60
            Width           =   1740
            _Version        =   851970
            _ExtentX        =   3069
            _ExtentY        =   688
            _StockProps     =   79
            Caption         =   "소요일 수정"
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
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "당일 발생분 사용불가능 합니다."
            Height          =   345
            Left            =   3300
            TabIndex        =   54
            Top             =   1440
            Width           =   1425
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   2745
         Left            =   15
         TabIndex        =   32
         Top             =   4995
         Width           =   4980
         _ExtentX        =   8784
         _ExtentY        =   4842
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSOption optReceipt 
            Height          =   285
            Index           =   0
            Left            =   1650
            TabIndex        =   50
            Top             =   2370
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   503
            _Version        =   262144
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "미출력"
         End
         Begin XtremeSuiteControls.PushButton btnCard 
            Height          =   420
            Left            =   3195
            TabIndex        =   33
            Top             =   1395
            Width           =   1740
            _Version        =   851970
            _ExtentX        =   3069
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
            Left            =   3195
            TabIndex        =   34
            Top             =   945
            Width           =   1740
            _Version        =   851970
            _ExtentX        =   3069
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   420
            Index           =   10
            Left            =   45
            TabIndex        =   35
            Top             =   45
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
            PictureBackground=   "frm접수결제.frx":5347
            BorderWidth     =   0
            BevelOuter      =   1
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   420
            Index           =   11
            Left            =   45
            TabIndex        =   36
            Top             =   495
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
            PictureBackground=   "frm접수결제.frx":5689
            BorderWidth     =   0
            BevelOuter      =   1
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.silgEdit txtIncome 
            Height          =   420
            Left            =   1605
            TabIndex        =   37
            Top             =   45
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
            Left            =   1605
            TabIndex        =   38
            Top             =   495
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
         Begin CSTextLibCtl.silgEdit txtCash 
            Height          =   420
            Left            =   1605
            TabIndex        =   39
            Top             =   945
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
            Left            =   45
            TabIndex        =   40
            Top             =   945
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
            PictureBackground=   "frm접수결제.frx":59CB
            BorderWidth     =   0
            BevelOuter      =   1
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   420
            Index           =   13
            Left            =   45
            TabIndex        =   41
            Top             =   1395
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
            PictureBackground=   "frm접수결제.frx":5D0D
            BorderWidth     =   0
            BevelOuter      =   1
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   420
            Index           =   14
            Left            =   45
            TabIndex        =   42
            Top             =   1845
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
            PictureBackground=   "frm접수결제.frx":604F
            BorderWidth     =   0
            BevelOuter      =   1
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.silgEdit txtCard 
            Height          =   420
            Left            =   1605
            TabIndex        =   43
            Top             =   1395
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
         Begin CSTextLibCtl.silgEdit txtBalance2 
            Height          =   420
            Left            =   1605
            TabIndex        =   44
            Top             =   1845
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   420
            Index           =   15
            Left            =   45
            TabIndex        =   49
            Top             =   2295
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
            Caption         =   "영수증 출력"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm접수결제.frx":6391
            BorderWidth     =   0
            BevelOuter      =   1
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSOption optReceipt 
            Height          =   285
            Index           =   1
            Left            =   2775
            TabIndex        =   51
            Top             =   2370
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   503
            _Version        =   262144
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "1장"
         End
         Begin Threed.SSOption optReceipt 
            Height          =   285
            Index           =   2
            Left            =   3645
            TabIndex        =   52
            Top             =   2370
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   503
            _Version        =   262144
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "2장"
            Value           =   -1
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   360
         Index           =   2
         Left            =   15
         TabIndex        =   45
         Top             =   4620
         Width           =   4980
         _ExtentX        =   8784
         _ExtentY        =   635
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " 결제금액"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm접수결제.frx":66D3
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
End
Attribute VB_Name = "frm접수결제"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iRowCount   As Integer
Dim 택코드      As String
Dim 세탁소요일  As Integer
Dim 접수번호    As Long      '
Dim CommPort As String
Dim BaudRate As String

Private Function chk_Item(ByVal strDate As String, ByVal CustCode As String) As Boolean
    Query = "SELECT COUNT(고객코드) AS DataCount"
    Query = Query & " FROM TB_입출고 "
    Query = Query & " WHERE 접수일자  >= '" & strDate & "'"
    'Query = Query & "   AND (확인 = '' OR 확인 IS NULL)"
    Query = Query & "   AND 고객코드 = '" & CustCode & "'"
    Set Rs = New ADODB.RecordSet
    Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    chk_Item = IIf(Rs!DataCount > 0, True, False)
    
    Rs.Close
    Set Rs = Nothing
End Function

'----------------------------------------------------------------------------------------
' 함수명 : Receive_Update
'
' 설  명 : 접수저장 & 결제처리 & 영수증 출력
'----------------------------------------------------------------------------------------
Private Function Receive_Update(strDate As String, PayType As Integer) As Boolean
    Dim iRow           As Integer   '
    Dim iCnt           As Integer   '
    
    Dim 고객코드       As String    ' 고객코드
    Dim 의류명         As String    ' 의류명
    Dim 색상           As String    ' 색상
    Dim 무늬           As String    ' 무늬
    
    Dim 세탁내용       As String    ' 세탁내용
    Dim 세탁금액       As Long      ' 금액
    Dim 상표           As String    ' 상표
    Dim 결제여부       As String    ' 결제여부
    Dim 예정일자       As String    ' 예정일자
    Dim 수선금액       As Long      ' 수선금액
        
    Dim 오점내용       As String    ' 오점내용
    Dim 오점이미지파일 As String    ' 오점이미지파일
    
    'Dim strinDate     As String
    'Dim sCupon        As String    ' 쿠폰번호
    
    Dim sGroupGoods(3) As String    '
    
    
    Dim 세탁마진       As Double      '
    Dim 외주마진       As Double      '
    Dim 수선마진       As Double      '
        
    Dim 택번호         As String    ' 택번호
    Dim Tmp_TAG        As String    ' 부모택번호
    Dim Parent_TAG     As String    ' 부모택번호
    
    Dim 수선번호       As Long      '
    
    Dim 기준금액       As Integer   '
    Dim 적립마일리지   As Integer   '
    Dim 접수금액       As Long      '
    Dim 발생마일리지   As Long      '
    
'-----------------------------------------
    
    Dim 현금입금         As Long
    Dim 카드입금         As Long
    Dim 사용마일리지     As Long
    Dim 쿠폰입금         As Long
    
    Dim 완불부분계산잔액     As Long
    Dim 마일리지적용체크잔액  As Long
    
    Dim TmpValue         As Long
    Dim 누적마일리지     As Long
    Dim 사용가능마일리지 As Long
    
    Dim 최소마일리지     As Long
    
    Dim 의류금액         As Long
    
    Dim TempRate As String
    
    On Error GoTo ErrRtn
    
    Receive_Update = False
    
    Erase sGroupGoods
    
    ' 받은 금액과 계산하여 완불, 부분, 후불을 처리한다.
    완불부분계산잔액 = Val(txtIncome.Value) + Val(txtCard.Value) + Val(txtMileage.Value)
    
    '----------------------------------------------------------------------
    ' 2013-07-01일 현금/카드 결제부분만 마일리지 누적 하도록 처리 pds2004
    현금입금 = txtCash.Value
    카드입금 = txtCard.Value
    마일리지적용체크잔액 = 현금입금 + 카드입금
    
    ' 2014-04-22일 100570-청라1동점만 선불 현금 결제시만 적용 되도록 수정
    If 가맹점정보.가맹점코드 = "100570" Then
        마일리지적용체크잔액 = 현금입금
    End If
    '----------------------------------------------------------------------
    
    
    고객코드 = frm접수.txtCode.Text                    '고객코드
    택코드 = Format(frm접수.btnTagCode.Caption, "000") '
    사용마일리지 = txtMileage.Value                    '사용마일리지
    세탁금액 = 0                                       '
    
    If frm접수.chkRepair.Value = -1 Then
        Query = "SELECT ISNULL(MAX(택번호),0) + 1 FROM TB_입출고"
        Query = Query & " WHERE 접수일자 = '" & Format(Date, "YYYY-MM-DD") & "'"
        Query = Query & "   AND 내용 LIKE '%수%'"
        Set ADORs = New ADODB.RecordSet
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        수선번호 = ADORs(0)
        
        ADORs.Close
        Set ADORs = Nothing
    End If
    
'    frm접수.sprGrid.Row = iCnt                                         '
'    frm접수.sprGrid.Col = 1: 의류명 = Trim(frm접수.sprGrid.Text) & ""  '의류명
    
    예정일자 = Format(DateAdd("d", 세탁소요일, strDate), "YYYY-MM-DD") '예정일자
    
    '세트상품정보.d세트Key = Format(Now, "YYYY-MM-DD hh:mm:ss")          '
        
    '-----------------------------------------------------------------
    'Dim 접수수량     As Integer
    
    iCnt = 0
    
    With frm접수.sprGrid
        Do Until iCnt >= .MaxRows
            iCnt = iCnt + 1

            .Row = iCnt
            .Col = 1: 의류명 = Trim(.Text) & ""

            If Trim(의류명) = "" Then
                iCnt = iCnt - 1

                Exit Do
            End If
        Loop
    End With
            
    
    If 가맹점정보.DualComputer = "Y" Then
        Dim Tag_List As String
        Dim StartPos As Integer
        Dim MyPos    As Integer
    
        Dim ADOCmd      As ADODB.Command
    
        Set ADOCmd = New ADODB.Command
    
        With ADOCmd
            .ActiveConnection = ADOCon
            .CommandText = "[SP_TAG]"
            .CommandType = adCmdStoredProc
            
            .Parameters.Append .CreateParameter("@Cnt", adInteger, adParamInput, 4)          ' 1 접수량
            .Parameters.Append .CreateParameter("@ReceiptDate", adVarChar, adParamInput, 10) ' 2 접수일자
            .Parameters.Append .CreateParameter("@ReceiptTime", adVarChar, adParamInput, 10) ' 3 접수시간
            .Parameters.Append .CreateParameter("@TagCode", adVarChar, adParamInput, 3)      ' 4 택코드
            .Parameters.Append .CreateParameter("@BranchCode", adVarChar, adParamInput, 4)   ' 5 지사코드
            .Parameters.Append .CreateParameter("@AgencyCode", adVarChar, adParamInput, 6)   ' 6 가맹점코드
            .Parameters.Append .CreateParameter("@Rtn", adVarChar, adParamOutput, 5000)      ' 7
            
            .Parameters("@Cnt") = iCnt & ""                                                  ' 1 접수량
            .Parameters("@ReceiptDate") = Format(Date, "YYYY-MM-DD") & ""                    ' 2 접수일자
            .Parameters("@ReceiptTime") = Format(Now, "hh:mm:ss") & ""                       ' 3 접수시간
            .Parameters("@TagCode") = 택코드 & ""                                            ' 4 택코드
            .Parameters("@BranchCode") = 가맹점정보.지사코드 & ""                            ' 5 지사코드
            .Parameters("@AgencyCode") = 가맹점정보.가맹점코드 & ""                          ' 6 가맹점코드
            
            .Execute , , adExecuteNoRecords
             
             Tag_List = .Parameters("@Rtn").Value & ""                                       ' 7
        End With
        
        StartPos = 1                                                   '
        
        Tag_List = Mid(Tag_List, 1, Len(Tag_List) - 1)                 ' "접수번호:택번호 리스트"
        
        MyPos = InStr(StartPos, Tag_List, ":", vbTextCompare)          ' 접수번호 찾기
        
        If MyPos = 0 Then
            접수번호 = 1                                               '
        Else
            접수번호 = Mid(Tag_List, 1, MyPos - StartPos)              ' 접수번호
            
            Tag_List = Mid(Tag_List, MyPos + 1, Len(Tag_List) - MyPos) ' 택번호 리스트
        End If
        
        With frm접수.sprGrid
            For i = 1 To iCnt
                .Row = i
                
                MyPos = InStr(StartPos, Tag_List, ",", vbTextCompare)
                
                If MyPos = 0 Then
                    택번호 = Right(Tag_List, 6)
                Else
                    택번호 = Mid(Tag_List, StartPos, MyPos - StartPos)
                    택번호 = Right(택번호, 6)
                End If
                
                .Col = 2: .Text = Format(택번호, "00-0000")
                
                StartPos = MyPos + 1
            Next i
        End With
    Else
        '-----------------------------------------------------------------
        ' TB_기본정보 - 접수번호
        '-----------------------------------------------------------------
        Query = "SELECT ISNULL(MAX(접수번호),0) + 1 FROM TB_기본정보"
        Set ADORs = New ADODB.RecordSet
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        접수번호 = ADORs(0) '접수번호
        
        ADORs.Close
        Set ADORs = Nothing
    End If
    
    'While Len(Trim(의류명)) > 0 And iCnt <= frm접수.sprGrid.MaxRows
    For iRow = 1 To iCnt
        With frm접수.sprGrid
            .Row = iRow
            
            .Col = 1: 의류명 = Trim(.Text) & ""                           ' 1 의류명
            
            If frm접수.chkRepair.Value = -1 Then
                택번호 = ""
            Else
                .Col = 2: 택번호 = 택코드 & Replace(.Text, "-", "") & ""  ' 2 택번호 ('999' + '00-0000')
            End If
            
            .Col = 3:  색상 = Trim(.Text) & ""                            ' 3 색상
            .Col = 4:  무늬 = Trim(.Text) & ""                            ' 4 무늬
            .Col = 5:  세탁내용 = Trim(.Text) & ""                        ' 5 세탁내용
            .Col = 6:  세탁금액 = .Value & ""                             ' 6 금액
            .Col = 7:  상표 = SubSQuotA(Trim(.Text)) & ""                 ' 7 상표
            
            If LenH(상표) > 50 Then
                상표 = MidH(상표, 1, 50)                                  '50자리보다 길면...
            End If
            
            
            .Col = 8:  의류코드 = Trim(.Text) & ""                        ' 8 의류코드
            .Col = 9:  수선금액 = IIf(.Value = "", 0, .Value)             ' 9 수선 금액 별도 입력
            
            .Col = 11: sGroupGoods(0) = Trim(.Text) & ""                  '11 ex. 6-01, 5-01, 5-02
            .Col = 12: sGroupGoods(1) = .Value                            '12 세트 할인률을 기준으로 계산한 금액(10원단위 포함)
            .Col = 13: sGroupGoods(2) = .Value                            '13 원단위 절사후 다시 계산한 금액
            .Col = 14: sGroupGoods(3) = .Value                            '14 세트관련 내용
            
            .Col = 15
            If .Text = "1" Then
                Parent_TAG = Tmp_TAG & ""                                 '15 부모택번호
            Else
                Tmp_TAG = 택번호 & ""                                     '   부모택번호 임시 필드에 저장
                Parent_TAG = ""                                           '15 부모택번호 (빈공간)
            End If
            
            .Col = 16: 세탁마진 = .Value & ""                             '16 세탁마진
            .Col = 17: 외주마진 = .Value & ""                             '17 외주마진
            .Col = 18: 수선마진 = .Value & ""                             '18 수선마진
            .Col = 19: 오점내용 = SubSQuotA(Trim(.Text)) & ""             '19 오점내용
            
            If LenH(오점내용) > 50 Then
                오점내용 = MidH(오점내용, 1, 50)                          '50자리보다 길면...
            End If
            
            .Col = 20: 의류금액 = .Value & ""                             '20 의류금액
            
            .Col = 21: 오점이미지파일 = Trim(.Text) & ""                  '21 오점이미지파일
        End With
        
        상표 = Replace(상표, "'", "")
        
        If frm접수.chkRepair.Value = -1 Then
            택번호 = 수선번호
                        
            수선번호 = 수선번호 + 1
        End If
        
        '-------------------------------------------------------------
        ' 완불,후불,부분 처리를 위한 계산
        
        ' 완불
        If PayType = 0 Then
            결제여부 = "완불"
            
        ' 후불
        ElseIf PayType = 1 Then
            결제여부 = "미불"
        
        ' 부분 결제
        ElseIf PayType = 2 Then
            If 완불부분계산잔액 >= 세탁금액 Then
                결제여부 = "완불"
                완불부분계산잔액 = 완불부분계산잔액 - 세탁금액
                
            ElseIf 완불부분계산잔액 > 0 And 완불부분계산잔액 < 세탁금액 Then
                결제여부 = "부분"
                완불부분계산잔액 = 0
            
            Else
                 결제여부 = "미불"
                완불부분계산잔액 = 0
            End If
        End If
        '-------------------------------------------------------------
        
        
        If 가맹점정보.DualComputer = "Y" Then
            '-------------------------------------------------------------
            ' TB_입출고
            '-------------------------------------------------------------
            Query = "SELECT * FROM TB_입출고"
            Query = Query & " WHERE 접수일자 = '" & strDate & "'"
            Query = Query & "   AND 택번호   = '" & 택번호 & "'"
            Query = Query & "   AND 접수번호 =  " & 접수번호
            Set ADORs = New ADODB.RecordSet
            ADORs.Open Query, ADOCon, adOpenDynamic, adLockOptimistic
        
            If Not ADORs.EOF Then
                
'                '----------------------------------------------------------------------------------
                ' 2013-07-01일 현금/카드 결제부분만 마일리지 누적 하도록 처리 pds2004
                ' 크렌즈 갤러리는 제외
'                '----------------------------------------------------------------------------------
                If 가맹점정보.지사코드 = M_COUPON_KLENZ_CODE Then
                    '----------------------------------------------------------------------------------
                    ' 마일리지로 결제한 금액만큼 마일리지 적립을 해주면 안된다이~~~~
                    '----------------------------------------------------------------------------------
                    If 사용마일리지 = 0 Then
                        TmpValue = 세탁금액
    
                    ElseIf 사용마일리지 > 세탁금액 Then
                        ' TmpValue = 세탁금액 ' 2013-06-24일 pds2004 아래와 같이 수정
                       TmpValue = 0
    
                        사용마일리지 = 사용마일리지 - 세탁금액
    
                    ElseIf 사용마일리지 > 0 Then
                        TmpValue = 세탁금액 - 사용마일리지
    
                        사용마일리지 = 0
                    End If
                    '----------------------------------------------------------------------------------
                
                Else
                    ' 0원일 경우 마일리지 누적 하지 않음
                    If 마일리지적용체크잔액 = 0 Then
                        TmpValue = 0
                    
                    ' 결제 금액이 더 많을 경우 마일리지 누적
                    ElseIf 마일리지적용체크잔액 >= 세탁금액 Then
                        TmpValue = 세탁금액
                        마일리지적용체크잔액 = 마일리지적용체크잔액 - 세탁금액
                    
                    ' 결제 금액이 더 적을 경우 결제 금액만큼만 누적
                    ElseIf 마일리지적용체크잔액 < 세탁금액 Then
                        TmpValue = 마일리지적용체크잔액
                        마일리지적용체크잔액 = 0
                    End If
    '                '----------------------------------------------------------------------------------
                End If
                
                If 가맹점정보.마일리지여부 = "Y" Then
                    If 가맹점정보.기준금액 = 0 Then
                        발생마일리지 = 0
                    Else
                        발생마일리지 = (가맹점정보.적립마일리지 / 가맹점정보.기준금액) * TmpValue
                        '발생마일리지 = (가맹점정보.적립마일리지 / 가맹점정보.기준금액) * 세탁금액
                    End If
                Else
                    발생마일리지 = 0
                End If
                
                '----------------------------------------------------------------------------------
                ADORs!접수일자 = strDate & ""                 ' 1
                ADORs!고객코드 = 고객코드 & ""                ' 2
                ADORs!의류명 = 의류명 & ""                    ' 3
                ADORs!택번호 = 택번호 & ""                    ' 4
                ADORs!색상 = 색상 & ""                        ' 5
                ADORs!무늬 = 무늬 & ""                        ' 6
                ADORs!내용 = 세탁내용 & ""                    ' 7
                
                If frmInputEventCode.Rate <> "" Then
                    TempRate = frmInputEventCode.Rate
                    TempRate = Replace(TempRate, "%", "")
                    세탁금액 = 의류금액 - (의류금액 * (Val(TempRate) / 100))
                    세탁금액 = 의류금액
                End If
                
                ADORs!금액 = 세탁금액 & ""                    ' 8
                ADORs!상표 = 상표 & ""                        ' 9
                ADORs!의류코드 = 의류코드 & ""                '10
                ADORs!결제여부 = 결제여부 & ""                '11
                ADORs!판매취소 = ""                           '12
                ADORs!예정일자 = 예정일자 & ""                '13
                ADORs!수선금액 = 수선금액 & ""                '14
                ADORs!세트구분 = sGroupGoods(0) & ""          '15
                ADORs!세트금액1 = sGroupGoods(1) & ""         '16
                ADORs!세트금액2 = sGroupGoods(2) & ""         '17
                ADORs!정상금액 = sGroupGoods(3) & ""          '18
                ADORs!세트Key = 세트상품정보.d세트Key & ""    '19
                ADORs!부모택번호 = Parent_TAG & ""            '20
                ADORs!근무자명 = strManager & ""              '21
                ADORs!접수번호 = 접수번호 & ""                '22
                ADORs!세탁마진 = 세탁마진 & ""                '23
                ADORs!외주마진 = 외주마진 & ""                '24
                ADORs!수선마진 = 수선마진 & ""                '25
                ADORs!오점내용 = 오점내용 & ""                '26
                ADORs!의류금액 = 의류금액 & ""                '27
                ADORs!마일리지 = 발생마일리지 & ""            '28
                ADORs!접수시간 = Format(Now, "hh:mm:ss") & "" '29
                ADORs!가맹점코드 = 가맹점정보.가맹점코드 & "" '30
                ADORs!지사코드 = 가맹점정보.지사코드 & ""     '31
                ADORs!본사전송여부 = ""                       '32
                
                'If 오점이미지파일 = "" Then
                    ADORs!오점이미지 = ""                     '33
                    
                    ADORs.Update
'                Else
'                    Dim ADOStream As New ADODB.Stream
'
'                    With ADOStream
'                        .Type = adTypeBinary
'                        .Open
'                        .LoadFromFile AppPath & "Capture\" & 오점이미지파일
'
'                        ADORs!오점이미지 = .Read              '33
'                    End With
'
'                    ADORs.Update
'
'                    Set ADOStream = Nothing
'                End If
                
                '부모택번호의 모(母)택번호의 부모택번호 필드에 자기 택번호를 저장
                If Parent_TAG <> "" Then
                    Query = "UPDATE TB_입출고 SET 부모택번호 = '" & Parent_TAG & "'"
                    Query = Query & " WHERE 접수일자 = '" & strDate & "'"
                    Query = Query & "   AND 택번호   = '" & Parent_TAG & "'"
                    ADOCon.Execute Query
                End If
            End If
            ADORs.Close
            Set ADORs = Nothing
            
        Else
            현금입금 = txtCash.Value
            카드입금 = txtCard.Value
                    
            '-------------------------------------------------------------
            ' TB_입출고
            '-------------------------------------------------------------
            Query = "SELECT * FROM TB_입출고"
            Query = Query & " WHERE 접수일자 = '" & strDate & "'"
            Query = Query & "   AND 택번호   = '" & 택번호 & "'"
            Query = Query & "   AND 접수번호 =  " & 접수번호
            Set ADORs = New ADODB.RecordSet
            ADORs.Open Query, ADOCon, adOpenDynamic, adLockOptimistic
            
            If ADORs.EOF Then
'                '----------------------------------------------------------------------------------
                ' 2013-07-01일 현금/카드 결제부분만 마일리지 누적 하도록 처리 pds2004
                ' 크렌즈 갤러리는 제외
'                '----------------------------------------------------------------------------------
                If 가맹점정보.지사코드 = M_COUPON_KLENZ_CODE Then
                    '----------------------------------------------------------------------------------
                    ' 마일리지로 결제한 금액만큼 마일리지 적립을 해주면 안된다이~~~~
                    '----------------------------------------------------------------------------------
                    If 사용마일리지 = 0 Then
                        TmpValue = 세탁금액
    
                    ElseIf 사용마일리지 > 세탁금액 Then
                        ' TmpValue = 세탁금액 ' 2013-06-24일 pds2004 아래와 같이 수정
                       TmpValue = 0
    
                        사용마일리지 = 사용마일리지 - 세탁금액
    
                    ElseIf 사용마일리지 > 0 Then
                        TmpValue = 세탁금액 - 사용마일리지
    
                        사용마일리지 = 0
                    End If
                    '----------------------------------------------------------------------------------

                Else
                    ' 0원일 경우 마일리지 누적 하지 않음
                    If 마일리지적용체크잔액 = 0 Then
                        TmpValue = 0
                    
                    ' 결제 금액이 더 많을 경우 마일리지 누적
                    ElseIf 마일리지적용체크잔액 >= 세탁금액 Then
                        TmpValue = 세탁금액
                        마일리지적용체크잔액 = 마일리지적용체크잔액 - 세탁금액
                    
                    ' 결제 금액이 더 적을 경우 결제 금액만큼만 누적
                    ElseIf 마일리지적용체크잔액 < 세탁금액 Then
                        TmpValue = 마일리지적용체크잔액
                        마일리지적용체크잔액 = 0
                    End If
    '                '----------------------------------------------------------------------------------
                End If
                
                If 가맹점정보.마일리지여부 = "Y" Then
                    If 가맹점정보.기준금액 = 0 Then
                        발생마일리지 = 0
                    Else
                        발생마일리지 = (가맹점정보.적립마일리지 / 가맹점정보.기준금액) * TmpValue
                        '발생마일리지 = (가맹점정보.적립마일리지 / 가맹점정보.기준금액) * 세탁금액
                    End If
                Else
                    발생마일리지 = 0
                End If
                        
                '----------------------------------------------------------------------------------
                ADORs.AddNew
                
                ADORs!접수일자 = strDate & ""                 ' 1
                ADORs!고객코드 = 고객코드 & ""                ' 2
                ADORs!의류명 = 의류명 & ""                    ' 3
                ADORs!택번호 = 택번호 & ""                    ' 4
                ADORs!색상 = 색상 & ""                        ' 5
                ADORs!무늬 = 무늬 & ""                        ' 6
                ADORs!내용 = 세탁내용 & ""                    ' 7
                
                
                
'                If frmInputEventCode.Rate <> "" Then
'                    TempRate = frmInputEventCode.Rate
'                    If InStr(TempRate, "%") > 0 Then
'                        TempRate = Replace(TempRate, "%", "")
'                        If 세탁금액 > 의류금액 Then
'                            세탁금액 = 세탁금액 - (세탁금액 * (Val(TempRate) / 100))
'                        Else
'                            세탁금액 = 의류금액 - (의류금액 * (Val(TempRate) / 100))
'                        End If
'                    Else
'                        If (CDbl(세탁금액) < CDbl(TempRate)) Then
'                            세탁금액 = 0
'                        Else
'                            세탁금액 = 세탁금액 - TempRate
'                        End If
'                    End If
'
'
'                End If
                If frmInputEventCode.Rate <> "" Then
                    If txtCoupon.Tag = "" Then
                    세탁금액 = 의류금액
                    End If
                End If
                
                ADORs!금액 = 세탁금액 & ""                    ' 8
                ADORs!상표 = 상표 & ""                        ' 9
                ADORs!의류코드 = 의류코드 & ""                '10
                ADORs!결제여부 = 결제여부 & ""                '11
                ADORs!판매취소 = ""                           '12
                ADORs!예정일자 = 예정일자 & ""                '13
                ADORs!수선금액 = 수선금액 & ""                '14
                ADORs!세트구분 = sGroupGoods(0) & ""          '15
                ADORs!세트금액1 = sGroupGoods(1) & ""         '16
                ADORs!세트금액2 = sGroupGoods(2) & ""         '17
                ADORs!정상금액 = sGroupGoods(3) & ""          '18
                ADORs!세트Key = 세트상품정보.d세트Key & ""    '19
                ADORs!부모택번호 = Parent_TAG & ""            '20
                ADORs!근무자명 = strManager & ""              '21
                ADORs!접수번호 = 접수번호 & ""                '22
                ADORs!세탁마진 = 세탁마진 & ""                '23
                ADORs!외주마진 = 외주마진 & ""                '24
                ADORs!수선마진 = 수선마진 & ""                '25
                ADORs!오점내용 = 오점내용 & ""                '26
                ADORs!의류금액 = 의류금액 & ""                '27
                ADORs!마일리지 = 발생마일리지 & ""            '28
                ADORs!접수시간 = Format(Now, "hh:mm:ss") & "" '29
                ADORs!가맹점코드 = 가맹점정보.가맹점코드 & "" '30
                ADORs!지사코드 = 가맹점정보.지사코드 & ""     '31
                ADORs!본사전송여부 = ""                       '32
                
                'If 오점이미지파일 = "" Then
                    ADORs!오점이미지 = ""                     '33
                    
                    ADORs.Update
                'Else
'                    Dim ADOStream2 As New ADODB.Stream
'
'                    With ADOStream2
'                        .Type = adTypeBinary
'                        .Open
'                        .LoadFromFile AppPath & "Capture\" & 오점이미지파일
'
'                        ADORs!오점이미지 = .Read              '33
'                    End With
                    
'                    ADORs.Update
'
'                    Set ADOStream2 = Nothing
'                End If
                
                '부모택번호의 모(母)택번호의 부모택번호 필드에 자기 택번호를 저장
                If Parent_TAG <> "" Then
                    Query = "UPDATE TB_입출고 SET 부모택번호 = '" & Parent_TAG & "'"
                    Query = Query & " WHERE 접수일자 = '" & strDate & "'"
                    Query = Query & "   AND 택번호   = '" & Parent_TAG & "'"
                    ADOCon.Execute Query
                End If
                
            ElseIf ADORs!판매취소 = "Y" Then
                Query = "UPDATE TB_입출고 SET"
                Query = Query & "  접수일자     = '" & strDate & "'"               '  1
                Query = Query & ", 고객코드     = '" & 고객코드 & "'"              '  2
                Query = Query & ", 의류명       = '" & 의류명 & "'"                '  3
                Query = Query & ", 택번호       = '" & 택번호 & "'"                '  4
                Query = Query & ", 색상         = '" & 색상 & "'"                  '  5
                Query = Query & ", 무늬         = '" & 무늬 & "'"                  '  6
                Query = Query & ", 내용         = '" & 세탁내용 & "'"              '  7
                Query = Query & ", 금액         =  " & 세탁금액                    '  8
                Query = Query & ", 상표         = '" & 상표 & "'"                  '  9
                Query = Query & ", 의류코드     = '" & 의류코드 & "'"              ' 10
                Query = Query & ", 결제여부     = '" & 결제여부 & "'"              ' 11
                Query = Query & ", 판매취소     = 'R'"                             ' 12
                Query = Query & ", 판매취소일자 = ''"                              ' 13
                Query = Query & ", 예정일자     = '" & 예정일자 & "'"              ' 14
                Query = Query & ", 세트구분     = '" & sGroupGoods(0) & "'"        ' 15
                Query = Query & ", 세트금액1    =  " & sGroupGoods(1)              ' 16
                Query = Query & ", 세트금액2    =  " & sGroupGoods(2)              ' 17
                Query = Query & ", 정상금액     =  " & sGroupGoods(3)              ' 18
                Query = Query & ", 세트Key      = '" & 세트상품정보.d세트Key & "'" ' 19
                Query = Query & ", 수선금액     =  " & 수선금액                    ' 20
                Query = Query & ", 부모택번호   = '" & Parent_TAG & "'"            ' 21
                Query = Query & ", 근무자명     = '" & strManager & "'"            ' 22
                Query = Query & ", 접수번호     =  " & 접수번호                  ' 23
                Query = Query & ", 세탁마진     =  " & 세탁마진                    ' 24
                Query = Query & ", 외주마진     =  " & 외주마진                    ' 25
                Query = Query & ", 수선마진     =  " & 수선마진                    ' 26
                Query = Query & ", 본사전송여부 = ''"                              '
                Query = Query & " WHERE 접수일자 = '" & strDate & "'"
                Query = Query & "   AND 택번호   = '" & 택번호 & "'"
                ADOCon.Execute Query
                
            Else
                ADORs.Close
                Set ADORs = Nothing
                
                ' 택번호를 잘못 수정하였을 경우 오류 메시지.
                MsgBox "[ " & 택번호 & " ]" & "이미 사용한 택번호 입니다. 택번호를 변경하여 주십시요", vbInformation
                
                Receive_Update = False
                
                Exit Function
            End If
            ADORs.Close
            Set ADORs = Nothing
        End If
        
        'iCnt = iCnt + 1
        '
        'frm접수.sprGrid.Row = iCnt
        'frm접수.sprGrid.Col = 1: 의류명 = Trim(frm접수.sprGrid.Text) '의류명
        '
        'If iCnt > frm접수.sprGrid.MaxRows Then
        '    의류명 = ""
        'End If
     Next iRow
    'Wend
    
    If 가맹점정보.DualComputer = "Y" Then
        '
    Else
        '-----------------------------------------------------------
        ' TB_기본정보 - 마지막 접수번호
        '-----------------------------------------------------------
        Query = "UPDATE TB_기본정보 SET 접수번호 = " & 접수번호
        ADOCon.Execute Query
    End If
    
'*****************************************************************************************
'* 마일리지 금액 적립
'*****************************************************************************************
    'txtCoupon.Value = 0
    
    접수금액 = txtTotalPay.Value
    현금입금 = txtCash.Value
    카드입금 = txtCard.Value
    사용마일리지 = txtMileage.Value
    쿠폰입금 = txtCoupon.Value
    
    
    ' 2013-07-01일 현금 카드 결제시에만 3% 적용
    If 가맹점정보.지사코드 = M_COUPON_KLENZ_CODE Then
        TmpValue = 접수금액 - 사용마일리지
    Else
        TmpValue = 현금입금 + 카드입금
    End If
    
    If 가맹점정보.마일리지여부 = "Y" Then
        If 가맹점정보.기준금액 = 0 Then
            발생마일리지 = 0
        Else
            '발생마일리지 = Int(TmpValue / 가맹점정보.기준금액) * 가맹점정보.적립마일리지
            발생마일리지 = (가맹점정보.적립마일리지 / 가맹점정보.기준금액) * TmpValue
        End If
    Else
        발생마일리지 = 0
    End If
    
    Call Get_고객마일리지(고객코드)
    
    '-----------------------------------------------------------
    ' TB_기본정보 - 최소마일리지
    '-----------------------------------------------------------
    Query = "SELECT ISNULL(최소마일리지,0) FROM TB_기본정보"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF Then
        최소마일리지 = 0
    Else
        최소마일리지 = ADORs(0)
    End If
    ADORs.Close
    Set ADORs = Nothing
    
    '-----------------------------------------------------------
    ' TB_매출
    '-----------------------------------------------------------
    Query = "SELECT * FROM TB_매출"
    Query = Query & " WHERE 고객코드 = '" & 고객코드 & "'"
    Query = Query & "   AND 접수번호 =  " & 접수번호
    Query = Query & "   AND 일련번호 = 0"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenDynamic, adLockOptimistic
    
    If ADORs.EOF Then ADORs.AddNew
    
    ADORs!가맹점코드 = 가맹점정보.가맹점코드                       ' 1
    ADORs!지사코드 = 가맹점정보.지사코드                           ' 2
    ADORs!고객코드 = 고객코드 & ""                                 ' 3
    ADORs!접수번호 = 접수번호 & ""                                 ' 4
    ADORs!일련번호 = 0                                             ' 5
    ADORs!매출일자 = strDate & ""                                  ' 6
    ADORs!매출시간 = Format(Now, "hh:mm:ss")                       ' 7
    
    frm접수.sprGrid.Row = 1
    frm접수.sprGrid.Col = 1
    ADORs!적요 = frm접수.sprGrid.Text & " 외"                      ' 8
    
    
    If frmInputEventCode.Rate <> "" Then
        If txtCoupon.Tag = "" Then
            ADORs!접수금액 = CDbl(txtTotalPay2.Value)
        Else
            ADORs!접수금액 = CDbl(txtTotalPay.Value)
        End If
    Else
        ADORs!접수금액 = CDbl(txtTotalPay.Value)                       ' 9
    End If
    Dim Internet_Cost As String
    Internet_Cost = ADORs!접수금액
    ADORs!현금입금 = 현금입금                                      '10
    ADORs!카드입금 = 카드입금                                      '11
    ADORs!쿠폰입금 = txtCoupon.Text   '쿠폰입금                                      '12
    ADORs!쿠폰번호 = txtCouponNo.Text                              '13
    ADORs!세트할인 = txtSetDC.Value                                '14
    ADORs!에누리 = txtDC.Value                                     '15
    ADORs!입금합계 = 현금입금 + 카드입금 + 사용마일리지 ' + 쿠폰입금 '16
    ADORs!접수수량 = iCnt 'iCnt - 1                                '17
    ADORs!반품수량 = 0                                             '18
    ADORs!사용마일리지 = 사용마일리지                              '19
    ADORs!발생마일리지 = 발생마일리지 & ""                         '20
    
    TmpValue = (마일리지.누적마일리지 + 발생마일리지)
    
    If 가맹점정보.마일리지여부 = "Y" Then
        If TmpValue >= 최소마일리지 Then
            If 최소마일리지 > 0 Then
                누적마일리지 = (TmpValue Mod 최소마일리지)
            Else
                ' 최소 마일리지가 0보다 적을 경우는 사용가능 마일리지로 처리되기 때문에 누적 마일리지로 저장하면 안된다.
                '누적마일리지 = TmpValue
            End If
            
            If 최소마일리지 > 0 Then
                '사용가능마일리지 = 마일리지.사용가능마일리지 + (TmpValue - (TmpValue Mod 최소마일리지)) '오류 x
                사용가능마일리지 = (마일리지.사용가능마일리지 - 사용마일리지) + (TmpValue - (TmpValue Mod 최소마일리지))
            Else
                ' 사용가능마일리지 = (마일리지.사용가능마일리지 - 사용마일리지) + TmpValue
                사용가능마일리지 = (마일리지.사용가능마일리지 - 사용마일리지) + TmpValue
            End If
        Else
            누적마일리지 = TmpValue
            사용가능마일리지 = 마일리지.사용가능마일리지 - 사용마일리지
        End If
        
    ' 마일리지를 적용하지 않을 경우
    ' pds2004 수정 2011-12-12일
    Else
        누적마일리지 = 0
        사용가능마일리지 = (마일리지.사용가능마일리지 - 사용마일리지)
    End If
    
    ADORs!누적마일리지 = 누적마일리지 & ""                         '21
    ADORs!사용가능마일리지 = 사용가능마일리지 & ""                 '22
    ADORs!삭제마일리지 = 0                                         '23
    ADORs!이전미수금 = frm접수.txtMisu.Value                       '24 2010-11-18 추가
    ADORs!본사전송여부 = ""                                        '
    
    ADORs.Update
    
    ADORs.Close
    Set ADORs = Nothing
    
    
        '-----------------------------------------------------------
    ' TB_매출_Internet
    '-----------------------------------------------------------
    Query = "SELECT * FROM TB_Internet_접수"
    Query = Query & " WHERE 고객코드 = '" & 고객코드 & "'"
    Query = Query & "   AND 접수번호 =  " & 접수번호
    Query = Query & "   AND 일련번호 = 0"
    Query = Query & "   AND Internet_접수번호 = '" & frm접수.btnInternet.Tag & "'"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenDynamic, adLockOptimistic
    
    If ADORs.EOF Then ADORs.AddNew
    ADORs!가맹점코드 = 가맹점정보.가맹점코드                       ' 1
    ADORs!지사코드 = 가맹점정보.지사코드                           ' 2
    ADORs!고객코드 = 고객코드 & ""                                 ' 3
    ADORs!접수번호 = 접수번호 & ""                                 ' 4
    ADORs!일련번호 = 0                                             ' 5
    ADORs!매출일자 = strDate & ""                                  ' 6
    ADORs!Internet_접수번호 = frm접수.btnInternet.Tag
    ADORs.Update
    
    ADORs.Close
    Set ADORs = Nothing
    
    With frm접수.sprGrid
    
    For i = 1 To iCnt
        Dim goods_name As String
        Dim Tag As String
        .Row = i
        .Col = 6: Internet_Cost = .Value
        .Col = 1: goods_name = .Text
        .Col = 2: Tag = frm접수.btnTagCode.Caption & Replace(.Text, "-", "")
        Call SetInternetAccept(frm접수.btnInternet.Tag, goods_name, Internet_Cost, CStr(i), Tag)
    Next i
    End With
    
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
            
            Query = "UPDATE TB_신용카드승인 SET 접수번호 =  " & 접수번호
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
            
            Query = "UPDATE TB_현금영수증 SET 접수번호 =  " & 접수번호
            Query = Query & "               , 고객코드 = '" & 고객코드 & "'"
            Query = Query & " WHERE 승인번호 = '" & 승인번호 & "'"
            Query = Query & "   AND 승인일자 = '" & 승인일자 & "'"
            Query = Query & "   AND 승인시간 = '" & 승인시간 & "'"
            ADOCon.Execute Query
        End If
    End With
    
    '-------------------------------------------------------------------------------
    ' TB_고객정보
    '-------------------------------------------------------------------------------
    Query = "UPDATE TB_고객정보 SET "
    Query = Query & "  이용횟수         = 이용횟수 + 1"
    Query = Query & ", 총접수금액       = 총접수금액 + " & 접수금액
    Query = Query & ", 누적마일리지     = " & 누적마일리지
    Query = Query & ", 사용가능마일리지 = " & 사용가능마일리지
    Query = Query & ", 최종거래일자     = '" & Format(Now, "YYYY-MM-DD hh:mm:ss") & "'"
    Query = Query & ", 본사전송여부     = ''"
    Query = Query & " WHERE 고객코드 = '" & 고객코드 & "'"
    ADOCon.Execute Query
    
    Receive_Update = True
    
    Exit Function
    
ErrRtn:
    Resume
    Receive_Update = False
    
    Call Error_Msg("Receive_Update", Err.Source, Err.Number, Err.description)
End Function

'*************************************************************************************
' 함수명 : Set_UseAccountUpdate
'
' 제  목 : 이용실적기록
' 기  능 : 연도와 고객코드를 가지고 이용실적Table에 write
' 처  리 : 1.이용횟수=한번계산될때마다 1씩증가
'          2.이용금액=이용금액+합계금액
'*************************************************************************************
Private Sub Set_UseAccountUpdate(고객코드 As String)
    Dim iMoney   As Long
    
    On Error GoTo ErrRtn
    
    iMoney = CCur(txtTotalPay.Value) '합계금액
        
    '------------------------------------------------------------------------
    ' 이용실적
    '------------------------------------------------------------------------
    Query = " SELECT * FROM TB_이용실적"
    Query = Query & " WHERE 고객코드 = '" & 고객코드 & "'"
    Query = Query & "   AND 연도     = '" & Format(Date, "YYYY") & "'"
    Set SUBRs = New ADODB.RecordSet
    SUBRs.Open Query, ADOCon, adOpenDynamic, adLockOptimistic
    
    If SUBRs.EOF Then
        SUBRs.AddNew
        
        SUBRs!가맹점코드 = 가맹점정보.가맹점코드        ' 0
        SUBRs!고객코드 = 고객코드 & ""                  ' 1
        SUBRs!연도 = Format(Date, "YYYY")               ' 2
        SUBRs!이용횟수 = 1 & ""                         ' 3
        SUBRs!이용금액 = iMoney & ""                    ' 4
    Else
        SUBRs!이용횟수 = SUBRs!이용횟수 + 1 & ""        ' 3
        SUBRs!이용금액 = SUBRs!이용금액 + iMoney & ""   ' 4
    End If
        
    SUBRs.Update
    
    SUBRs.Close
    Set SUBRs = Nothing
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("Set_UseAccountUpdate", Err.Source, Err.Number, Err.description)
End Sub

Private Function Get_수선금액() As Long
    Dim dblSuMoney As Long
    
    dblSuMoney = 0
    
    With frm접수.sprGrid
        For i = 1 To .MaxRows
            .Row = i
            
            .Col = 1
            If .Text = "" Then Exit For '의류명
            
            .Col = 9
            If .Text <> "" Then
                dblSuMoney = dblSuMoney + CCur(.Value) '수선금액
            End If
        Next i
    End With
    
    Get_수선금액 = dblSuMoney
End Function

Private Function Get_반품금액() As Long
    Dim dblSuMoney As Long
    Dim sCode      As String
    
    dblSuMoney = 0
    
    With frm접수.sprGrid
        For i = 1 To .MaxRows
            .Row = i
            .Col = 7
            If .Text = "" Then Exit For
            
            .Col = 4
            If InStr(.Text, "반") > 0 Then
                sCode = .Text
                dblSuMoney = dblSuMoney + Get_DryPrice(sCode)
            End If
        Next i
    End With
    
    Get_반품금액 = dblSuMoney
End Function

Private Sub btnCard_Click()
    Dim bBtnEnb(3)  As Boolean
    
    If Check_KS7500 = False Then
        MsgBox "환경설정에서 사업자번호, 단말기번호 등이 올바르게 입력되었는지 확인하십시요.", vbInformation, "확인"
        
        Exit Sub
    End If
    
    If txtBalance2.Value <= 0 Then
        MsgBox "결제 금액이 올바르게 입력되었는지 확인하십시요.", vbInformation, "확인"
        
        Exit Sub
    End If
    
    
    bBtnEnb(0) = cmdAction(0).Enabled
    bBtnEnb(1) = cmdAction(1).Enabled
    bBtnEnb(2) = cmdAction(2).Enabled
    bBtnEnb(3) = btnExit.Enabled
    
    btnCard.Enabled = False
    btnExit.Enabled = False
    cmdAction(0).Enabled = False
    cmdAction(1).Enabled = False
    cmdAction(2).Enabled = False
    
    Unload frmKSNET2
    
    Account_Form = "접수"
        
    frmKSNET2.pnlCustomCode.Caption = frm접수.txtCode.Text & "" '
    frmKSNET2.pnlNum.Caption = 0                                '
    frmKSNET2.txtMoney.Value = txtBalance2.Value                '
    frmKSNET2.txtMoney.Tag = txtBalance2.Value                  '
    
    Call frmKSNET2.신용카드승인요청_Rtn("1")

    frmKSNET2.Show vbModal
    
    cmdAction(0).Enabled = bBtnEnb(0)
    cmdAction(1).Enabled = bBtnEnb(1)
    cmdAction(2).Enabled = bBtnEnb(2)
    btnCard.Enabled = True

    ' 현금영수증이나 신용카드금액이 없을 경우에만 종료가 가능 하도록 변경
    If txtCash.Value = 0 And txtCard.Value = 0 Then
        btnExit.Enabled = True
    Else
        btnExit.Enabled = False
    End If


End Sub

Private Sub btnCash_Click()
    Dim vText   As Variant
    
    If Check_KS7500 = False Then
        MsgBox "환경설정에서 사업자번호, 단말기번호 등이 올바르게 입력되었는지 확인하십시요.", vbInformation, "확인"
        
        Exit Sub
    End If
    
    If txtCash.Value <= 0 Then
        MsgBox "받은 금액이 올바르게 입력되었는지 확인하십시요.", vbInformation, "확인"
        txtIncome.SetFocus
        Exit Sub
    End If
    Unload frmKSNETCash
    Account_Form = "접수"
    
    btnCash.Enabled = False
    frmKSNETCash.pnlCustomCode.Caption = frm접수.txtCode.Text '고객코드
    frmKSNETCash.pnlNum.Caption = 0                           '
    frmKSNETCash.txtMoney.Value = txtCash.Value               '현금결제금액
    
    Call frmKSNETCash.현금영수증승인요청_Rtn("3")
    
    frmKSNETCash.Show 1
    
    ' 현금 영수증을 승인하지 않았을 경우 다시 승인 할 수 있도록 처리
    sprCash.GetText 1, 1, vText
    If CStr(vText) = "" Then btnCash.Enabled = True
    
    ' 현금영수증이나 신용카드금액이 없을 경우에만 종료가 가능 하도록 변경
    If txtCash.Value = 0 And txtCard.Value = 0 Then
        btnExit.Enabled = True
    Else
        btnExit.Enabled = False
    End If
    
    
End Sub

Private Sub btnCashCancel_Click()
    sprCash.Row = 1
    sprCash.Col = 1
    
    If sprCash.Text = "" Then Exit Sub
    Unload frmKSNETCash
    Account_Form = "접수"
    
    With frmKSNETCash.sprGrid
        .Col = 1
    
        .Row = 1:  .Text = Spread_GetData(sprCash, 1, 1, True)   '승인번호
        .Row = 2:  .Text = Spread_GetData(sprCash, 2, 1, True)   '승인일자
        .Row = 3:  .Text = Spread_GetData(sprCash, 3, 1, True)   '승인시간
        .Row = 4:  .Text = Spread_GetData(sprCash, 4, 1, True)   '거래유형 '입력방법
        .Row = 5:  .Text = Spread_GetData(sprCash, 5, 1, True)   '총금액
        .Row = 6:  .Text = Spread_GetData(sprCash, 6, 1, True)   '사용자정보
        .Row = 7:  .Text = Spread_GetData(sprCash, 7, 1, True)   '메시지1
        .Row = 8:  .Text = Spread_GetData(sprCash, 8, 1, True)   '메시지2
        .Row = 9:  .Text = Spread_GetData(sprCash, 9, 1, True)   '소득구분
        .Row = 10: .Text = Spread_GetData(sprCash, 10, 1, True)  '국세청1
        .Row = 11: .Text = Spread_GetData(sprCash, 11, 1, True)  '국세청2
    End With

    frmKSNETCash.pnlCustomCode.Caption = Trim(frm접수.txtCode.Text)             '고객코드
    frmKSNETCash.pnlNum.Caption = 0                                             '접수번호
    frmKSNETCash.txtMoney.Value = Spread_GetData(sprCash, 5, 1, True)           '현금결제금액
    
    frmKSNETCash.pnlApprovalNo.Caption = Spread_GetData(sprCash, 1, 1, True) '승인번호
    frmKSNETCash.pnlApprovalDay.Caption = Spread_GetData(sprCash, 2, 1, True) '승인일자
    frmKSNETCash.pnlApprovalTime.Caption = Spread_GetData(sprCash, 3, 1, True) '승인시간
    
    Call frmKSNETCash.현금영수증승인요청_Rtn("4")
    
    frmKSNETCash.Show 1
    Dim vText   As Variant
    sprCash.GetText 1, 1, vText
    If CStr(vText) = "" Then btnCash.Enabled = True
    ' 현금영수증이나 신용카드금액이 없을 경우에만 종료가 가능 하도록 변경
    If txtCash.Value = 0 And txtCard.Value = 0 Then
        btnExit.Enabled = True
    Else
        btnExit.Enabled = False
    End If
    
End Sub

Private Sub btnEditDay_Click()
    Query = "UPDATE TB_기본정보 SET 세탁소요일 = " & Left(cboDay.Text, 2)
    ADOCon.Execute Query
    
    MsgBox "세탁소요일이 " & cboDay.Text & " 일로 변경되었습니다.", vbInformation, "확인"
End Sub

Private Sub btnEvent_Click()
    If btnEvent.Caption = "할인행사 입력" Then
        Dim TempRate As String
        frmInputEventCode.Rate = ""
        frmInputEventCode.EventCode = ""
        frmInputEventCode.Show 1
        TempRate = frmInputEventCode.Rate
        If InStr(TempRate, "ERROR") > 0 Then
            frmInputEventCode.Rate = ""
            Exit Sub
        End If
        If TempRate <> "" Then
            If InStr(TempRate, "%") > 0 Then
                '%계산식
                
                TempRate = Replace(TempRate, "%", "")
                txtCoupon.Value = (txtTotalPay2.Value * (Val(TempRate) / 100))
                txtSetDC.Text = "0"
            Else
                txtCoupon.Tag = "MONEY"
                If (CDbl(txtTotalPay.Value) < CDbl(TempRate)) Then
                    txtCoupon.Value = txtTotalPay.Value
                Else
                    txtCoupon.Value = TempRate
                End If
                'txtCoupon.Value = TempRate
            End If
            btnEvent.Tag = frmInputEventCode.EventCode
            btnEvent.Caption = "할인행사 취소"
        End If
    Else
        btnEvent.Caption = "할인행사 입력"

        btnEvent.Tag = ""
        txtCoupon.Value = 0
        txtCoupon.Tag = ""
        txtSetDC.Text = txtTotalPay2.Value - txtTotalPay.Value
    End If
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub cboDay_Click()
    Dim 세탁소요일 As Double
    
    세탁소요일 = Left(cboDay.Text, 1)
    
    dtpDay.Value = Format(DateAdd("d", 세탁소요일, Date), "YYYY-MM-DD")
End Sub

'-------------------------------------------------
' PayMode = 0 완불
' PayMode = 1 후불
'-------------------------------------------------
Private Sub cmdAction_Click(Index As Integer)
    Dim 고객코드     As String
    Dim 전화번호     As String
    Dim 접수일자     As String
    Dim PayType      As Integer
    Dim PrintCount   As Integer  ' 프린트 출력 장수
    
    Dim iPaper       As String
    
    Dim iSumMoney    As Long
    Dim iMoney       As Long
    Dim iCardMoney   As Long
    Dim iAccount     As Long
    
    On Error GoTo ErrRtn
        
    cmdAction(Index).Enabled = False
    DoEvents
    
    PayType = Index
    
    iSumMoney = txtTotalPay.Value  ' 합계금액
    iMoney = txtCash.Value         ' 현금입금
    iCardMoney = txtCard.Value     ' 카드입금
    iAccount = txtBalance2.Value   ' 결제후 잔액
    
    Select Case PayType
        Case 0 '완불
            If iAccount = 0 Then
                '이미 완불처리됨
            ElseIf iAccount > 0 And txtIncome.Value > 0 Then
                MsgBox "부분 결제가 되었습니다. 완불처리를 할수 없습니다.", vbInformation, "확인"
                cmdAction(Index).Enabled = True
                Exit Sub
            
            Else
                txtIncome.Value = txtBalance.Value - txtCard.Value '입금액 = 잔액 - 카드결제금액
            End If
            
        Case 1, 3 '후불
            If iAccount = 0 Then
                MsgBox "완불을 하였습니다. 후불처리를 할수 없습니다.", vbInformation, "확인"
                
                cmdAction(Index).Enabled = True
                Exit Sub
            End If
            
            If iMoney > 0 Or iCardMoney > 0 Then
                MsgBox "부분 결제가 되었습니다. 후불처리를 할수 없습니다.", vbInformation, "확인"
                
                cmdAction(Index).Enabled = True
                Exit Sub
            End If
            
        Case 2 '부분결제
            If iAccount = 0 Then
                MsgBox "완불을 하였습니다. 부분결제 처리를 할수 없습니다.", vbInformation, "확인"
                
                cmdAction(Index).Enabled = True
                Exit Sub
            End If
    End Select
        
    고객코드 = frm접수.txtCode.Text & "" '고객코드
    전화번호 = frm접수.txtTel.Text & ""  '전화번호
        
    If Get_일일마감여부(Format(Date, "YYYY-MM-DD")) = True Then
        MsgBox "일마감이 되었으므로 판매내역은 익일로 저장이 됩니다.", vbInformation
        
        접수일자 = Format(DateAdd("d", 1, Date), "YYYY-MM-DD")
    Else
        접수일자 = Format(Date, "YYYY-MM-DD")
    End If
        
    '---------------------------------------------------------------
    ' 수선접수인 경우에는 제외
    '---------------------------------------------------------------
    If frm접수.chkRepair.Value = 0 Then
    
        If 가맹점정보.DualComputer = "Y" Then
            '
        Else
            Call TAG_Update ' 가맹점 정보에 택번호를 저장한다.
        End If
    End If
    
    If "Error" = Get_고객정보(고객코드) Then
        MsgBox "해당 고객이 존재하지 않습니다.", vbCritical, "확인"
        
        cmdAction(Index).Enabled = True
        Exit Sub
    End If
    
    '-------------------------------------------------------------------------------------------
    ' TB_입출고 -
    ' TB_매출   -
    '-------------------------------------------------------------------------------------------
    If Receive_Update(접수일자, PayType) = False Then
        Query = "입출고 자료 저장중 오류가 발생 하였습니다. " & vbNewLine & vbNewLine
        Query = Query & "이미 사용중인 택번호를 다시 사용하였을 경우 발생할 수 있습니다." & vbNewLine
        Query = Query & "오류가 지속될 경우 전산담당자에게 문의하여 주십시요."
        MsgBox Query, vbInformation, "확인"
        
        cmdAction(Index).Enabled = True
        Exit Sub
    End If
    
    '-------------------------------------------------------------------------------------------
    ' TB_입출고 - 저장중 오류 확인...
    '-------------------------------------------------------------------------------------------
    If chk_Item(접수일자, 고객코드) = False Then
        Query = "입출고 자료 저장중 오류가 발생 하였습니다. " & vbNewLine & vbNewLine
        Query = Query & "Null을 저장하려고 시도할 경우 발생할 수 있습니다.." & vbNewLine
        Query = Query & "오류가 지속될 경우 전산담당자에게 문의 하여 주십시요."
        MsgBox Query, vbInformation, "확인"
        
        cmdAction(Index).Enabled = True
        Exit Sub
    End If
    
    Call Set_UseAccountUpdate(고객코드)        ' 이용실적
    
    'Call Set_CouponUpdate(접수일자, 고객코드)  ' 쿠폰 사용 내역 저장   2009-04-20 기능 추가
    
    
    
    
    
    
    '------------------------------------------------------------------------------------
    ' 잔액이 있는 경우만
    '------------------------------------------------------------------------------------
    If txtBalance2.Value > 0 Then
        Call Set_고객미수금액(고객코드, txtBalance2.Value, "ADD")
    End If
    
    Call Get_고객정보(고객코드)               ' 고객정보
    
    '------------------------------------------------------------------------
    ' 보관증출력
    '------------------------------------------------------------------------
    Dim CommPort As String
    Dim BaudRate As String
    Dim CardYN   As String
    Dim CashYN   As String
    Dim sE       As String
        
    CommPort = GetIniStr("VAN", "KS7500_CommPort", "", iniFile)
    BaudRate = GetIniStr("VAN", "KS7500_BaudRate", "", iniFile)
    
    
    ' 출력 여부 확인
    sprCard.Row = 1: sprCard.Col = 2:   CardYN = IIf(Trim(sprCard.Text) = "", "N", "Y")
    sprCash.Row = 1: sprCash.Col = 1:   CashYN = IIf(Trim(sprCash.Text) = "", "N", "Y")
    
    ' 카드전표, 현금영수증, 미출력 모두 아닐경우 처리하지 않는다.
    If CardYN = "N" And CashYN = "N" And optReceipt(0).Value = True Then

    
    ' 출력일 경우
    Else
            Dim TempCardPrint As String

                If optReceipt(0).Value = False Then
                    If optReceipt(0).Value = True Then
                        ' 아래쪽에서 카드 전표가 나오도록 하기 위해서
                        iPaper = 1
                    Else
                        iPaper = IIf(optReceipt(1).Value = True, 1, 2)
                         For PrintCount = 1 To iPaper
                            TempCardPrint = TempCardPrint & 접수영수증_Report(PrintCount, 접수일자)
                            
                        Next PrintCount
                    End If
                End If

            
            Call frmKicc.Card_Print(TempCardPrint)
    End If
    
    If btnEvent.Tag <> "" Then
        UpdateEventCode (btnEvent.Tag)
    End If
    Call frm접수.접수_Clear
        
    If chk_Item(접수일자, 고객코드) = True Then
        bSearch = True
        
        Call frm출고.Get_FindData("Code", 고객코드)
        DoEvents
        
        bSearch = False
        
        chkinputflig = "입고완료"
    
        cmdAction(Index).Enabled = True
        
        Unload frm접수결제
        
        frm출고.SetFocus
    Else
        frm접수.SetFocus
    End If
    
    Exit Sub
    
ErrRtn:
    cmdAction(Index).Enabled = True
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
    Resume
End Sub

Private Sub btnCoupon_Click()
    On Error GoTo ErrRtn
    
    Dim dblTempMoney As Double

    If btnCoupon.Caption = "쿠폰 사용" Then
        txtCouponNo.Visible = False
        
        txtCouponNo.Text = "05" & Right(가맹점정보.가맹점코드, 3) & Format(GetCouponCount(Format(Date, "yyyy-MM-dd"), "05"), "000")
        
        txtCouponNo_KeyPress vbKeyReturn
        DoEvents
        btnCoupon.Caption = "쿠폰 취소"
        
    Else
        txtCouponNo.Visible = False
        txtCoupon.Text = "0"
        txtCouponNo.Text = ""
        
        Sub_마일리지사용
        
        DoEvents
        btnCoupon.Caption = "쿠폰 사용"
    
    End If
    Exit Sub
    
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

'+------------------------------------------------------
'+
'+ 2009/09/09
'+
'+루틴설명
' 삼성 카드 10% 할인
'+------------------------------------------------------
Private Sub cmdSamSungCard_Click()
    Dim nRow        As Long
    Dim logPrice    As Long
    Dim sTemp       As String
    Dim iPercentage As Double
    Dim sumMoney    As Double
    Dim strFirst    As String
       
    If 가맹점정보.삼성카드할인여부 <> "Y" Then Exit Sub
    
    sumMoney = 0
    strFirst = "삼"
    iPercentage = (100 - 가맹점정보.삼성카드할인비율) / 100  ' (할인이 20%일 경우 0.8의 값을 같는다.)
    
    For nRow = 1 To frm접수.sprGrid.MaxRows
    
        If Get_SpreadText(frm접수.sprGrid, nRow, 2) = "" Then Exit For
        
        frm접수.sprGrid.Row = nRow
        frm접수.sprGrid.Col = 5
        sTemp = Trim(frm접수.sprGrid.Text)
        
        If InStr(1, sTemp, strFirst, vbTextCompare) <= 0 Then
            ' 내용에 "삼"자가 없을 경우 "삼"을 추가하여 출력 한다.
            frm접수.sprGrid.Text = Mid(sTemp, 1, 1) & strFirst & Mid(sTemp, 2, Len(sTemp))
            
            ' 해당 금액을 얻어온다.
            logPrice = CLng(Get_SpreadText(frm접수.sprGrid, nRow, 5))
                
            frm접수.sprGrid.Row = nRow
            frm접수.sprGrid.Col = 6
            ' 10원단위까지 수령한다.
            frm접수.sprGrid.Text = CStr(logPrice * iPercentage)
'            frm접수.sprGrid.Text = CStr(Int(CDbl((logPrice * iPercentage) / 100)) * 100) 10원단위 절사
            
            ' 누적 금액을 다시 계산한다.
            sumMoney = sumMoney + CLng(Get_SpreadText(frm접수.sprGrid, nRow, 5))
    
            frm접수.lblSamSungCardCheck.Tag = "Y"
        
        Else
            ' 내용에 "삼"을 제거한다
            frm접수.sprGrid.Text = Replace(frm접수.sprGrid.Text, "삼", "")
            
            ' 해당 금액을 얻어온다.
            logPrice = CLng(Get_SpreadText(frm접수.sprGrid, nRow, 5))
                
            frm접수.sprGrid.Row = nRow
            frm접수.sprGrid.Col = 6
            ' 10원단위까지 수령한다.
            frm접수.sprGrid.Text = CStr(logPrice / iPercentage)
'            frm접수.sprGrid.Text = CStr(Int(CDbl((logPrice * iPercentage) / 100)) * 100) 10원단위 절사
            
            ' 누적 금액을 다시 계산한다.
            sumMoney = sumMoney + CLng(Get_SpreadText(frm접수.sprGrid, nRow, 5))
            
            frm접수.lblSamSungCardCheck.Tag = "N"
        End If
    Next nRow
            
    'If sumMoney > 0 Then txtSum.Text = sumMoney & ""
    If sumMoney > 0 Then txtTotalPay.Value = sumMoney 'Call Spread_SetData(sprMoney, 3, 1, CStr(sumMoney))
    
    cmdSamSungCard.Caption = IIf(frm접수.lblSamSungCardCheck.Tag = "Y", "삼성카드 할인 취소", "삼성카드 할인(10%)")
            
    DoEvents
End Sub

Private Sub Form_Activate()
    Dim 출고예정일   As String
    Dim 접수금액     As Long
    Dim 최소마일리지 As Long
    
    On Error GoTo ErrRtn
    
    Me.Top = ((frmMain.Height - Me.Height) / 2) + frmMain.Top
    Me.Left = ((frmMain.Width - Me.Width) / 2) + frmMain.Left
    
    If 접수결제_Flag = True Then Exit Sub
    
    
    If Format(Date, "yyyy-MM-dd") >= "2011-10-07" And Format(Date, "yyyy-MM-dd") <= "2011-12-31" Then
        btnCoupon.Enabled = True
    Else
        btnCoupon.Enabled = False
    End If
    
    '-----------------------------------------------------------
    ' TB_기본정보
    '-----------------------------------------------------------
    Query = "SELECT    ISNULL(세탁소요일,3)"
    Query = Query & ", ISNULL(최소마일리지,0)"
    Query = Query & " FROM TB_기본정보"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF Then
        세탁소요일 = 3
        최소마일리지 = 0
    Else
        세탁소요일 = Trim(ADORs(0))
        최소마일리지 = ADORs(1)
    End If
    ADORs.Close
    Set ADORs = Nothing
    
    With cboDay
        .Clear
        
        For i = 3 To 15
            .AddItem i & " 일"
        Next i
        
        .Text = 세탁소요일 & " 일"
    End With
                
    '------------------------------------------------------------------------
    ' 당일 마일리지가 누적된 경우에는 사용 불가~~~
    '------------------------------------------------------------------------
    Query = "SELECT TOP 1"
    Query = Query & " 사용가능마일리지 - 발생마일리지"
    Query = Query & " FROM TB_매출"
    Query = Query & " WHERE 고객코드 = '" & frm접수.txtCode.Text & "'"
    Query = Query & "   AND 매출일자 = '" & Format(Date, "YYYY-MM-DD") & "'"
    Query = Query & "   AND 사용가능마일리지 >= " & 최소마일리지
    Query = Query & "   AND 발생마일리지 > 0 "
    Query = Query & " ORDER BY 매출시간 DESC"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF Then
        Call Sub_마일리지사용
    Else
        If ADORs(0) > 최소마일리지 Then
            Call Sub_마일리지사용
        End If
    End If
    ADORs.Close
    Set ADORs = Nothing
    
    frmInputEventCode.Rate = ""
    
    '삼성카드 할인 여부
    cmdSamSungCard.Enabled = IIf(가맹점정보.삼성카드할인여부 = "Y", True, False)
    cmdSamSungCard.Caption = IIf(frm접수.lblSamSungCardCheck.Tag = "Y", "삼성카드 할인 취소", "삼성카드 할인(10%)")
    
    출고예정일 = Format(DateAdd("d", 세탁소요일, Date), "YYYY-MM-DD")
    
    dtpDay.Value = 출고예정일              '출고예정일
    
    접수결제_Flag = True
    
    txtIncome.SetFocus                     '현금결제 Cell Focus
    
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    접수결제_Flag = False
End Sub

Private Sub Sub_마일리지사용()
    Dim 접수금액 As Long
    
    Call Get_고객마일리지(frm접수.txtCode.Text)       '마일리지 금액을 표시한다.
    
    접수금액 = txtTotalPay.Value '접수금액
        
    If 마일리지.사용가능마일리지 > 접수금액 Then
        txtMileage.Value = 접수금액
    Else
        txtMileage.Value = 마일리지.사용가능마일리지
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyEscape Then
'        Unload Me
'    End If
End Sub

Private Sub Form_Load()
    'frm접수결제.Top = 2000
    'frm접수결제.Left = 3000
    
    'If 가맹점정보.가맹점코드 = "999999" And Format(Date, "YYYY-MM-DD") <= "20091211" Then
    '    frm접수결제.Top = 2000
    '    frm접수결제.Left = 8000
    'End If
    
    Me.Top = ((frmMain.Height - Me.Height) / 2) + frmMain.Top
    Me.Left = ((frmMain.Width - Me.Width) / 2) + frmMain.Left
    
    
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
    
    Dim iPaper As String
    
    iPaper = GetIniStr("Printer", "Paper", "", iniFile)
    
    If iPaper = "2" Then
        optReceipt(2).Value = True
    ElseIf iPaper = "1" Then
        optReceipt(1).Value = True
    Else
        optReceipt(0).Value = True
    End If
End Sub


Private Sub sprCard_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    If Row <= 0 Then Exit Sub
    Unload frmKSNET2
    
    Account_Form = "접수"
    
    
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


    frmKSNET2.pnlCustomCode.Caption = frm접수.txtCode.Text
    frmKSNET2.pnlNum.Caption = 0
    
    frmKSNET2.txtMoney.ReadOnly = True
    frmKSNET2.txtMoney.Value = Spread_GetData(sprCard, Row, 6, True)
    
    frmKSNET2.pnlApprovalNo.Caption = Spread_GetData(sprCard, Row, 2, True)   '승인번호
    frmKSNET2.pnlApprovalDay.Caption = Spread_GetData(sprCard, Row, 3, True)  '승인일자
    frmKSNET2.pnlApprovalTime.Caption = Spread_GetData(sprCard, Row, 4, True) '승인시간
    
    Call frmKSNET2.신용카드승인요청_Rtn("2")
    
    frmKSNET2.Show vbModal
End Sub

Private Sub CouponNo_Proc()
    Dim varTemp       As Variant
    Dim nIndex        As Integer
    Dim CouponMoney   As Double
    Dim sCouponNum    As String
    
RE_START:
    txtCoupon.Value = 0
    
    CouponMoney = 0
    
    varTemp = Split(txtCouponNo.Text, vbNewLine)
    
    For nIndex = 0 To UBound(varTemp)
        sCouponNum = Trim(varTemp(nIndex))
        
        If sCouponNum <> "" Then
            '4자리 입력 내용 변환
            If Len(sCouponNum) = 6 And Left(sCouponNum, 2) = "01" Then
                
                txtCouponNo.Text = Replace(txtCouponNo.Text, sCouponNum, Left(sCouponNum, 2) & "00" & Right(sCouponNum, 4))
                
                GoSub RE_START
            End If
            
            Select Case CheckCouponNumber(sCouponNum)
                Case -1
                    MsgBox "쿠폰 번호 오류 [" & sCouponNum & "]", vbInformation, "확인"
                    Exit Sub
                                
                Case -2
                    MsgBox "쿠폰 사용기간 오류 [" & sCouponNum & "]", vbInformation, "확인"
                    Exit Sub
            End Select
            
            CouponMoney = CouponMoney + Get_CouponMoney(sCouponNum) ' 쿠폰 금액을 누적 처리한다.
            
            txtCoupon.Value = CStr(CouponMoney)
        End If
    Next nIndex

    If 마일리지.사용가능마일리지 > 0 Then
        If CouponMoney = 0 Then
            If 마일리지.사용가능마일리지 > txtTotalPay.Value Then
                txtMileage.Value = txtTotalPay.Value
            Else
                txtMileage.Value = 마일리지.사용가능마일리지
            End If
        Else
            If 마일리지.사용가능마일리지 > CouponMoney Then
                ' 마일리지 잔액이 전체 금액보다 적을 경우 마일리지 금액만
                If txtTotalPay.Value > (마일리지.사용가능마일리지 - CouponMoney) Then
                    txtMileage.Value = 마일리지.사용가능마일리지 - CouponMoney
                
                ' 마일리지 금액이 전체 금액 보다 클 경우 전체 금액만 처리한다.
                Else
                    txtMileage.Value = txtTotalPay.Value - CouponMoney
                End If
                
            Else
                If txtTotalPay.Value - CouponMoney Then
                    If 마일리지.사용가능마일리지 < txtTotalPay.Value - CouponMoney Then
                        txtMileage.Value = 마일리지.사용가능마일리지
                    Else
                        txtMileage.Value = txtTotalPay.Value - CouponMoney
                    End If
                Else
                    txtMileage.Value = 0
                End If
            End If
        End If
    End If
End Sub

' 사용한 쿠폰의 자료를 저장한다.
Private Sub Set_CouponUpdate(ByVal strDate As String, 고객코드 As String)
    Dim varTemp     As Variant
    Dim nIndex      As Integer
    Dim CouponMoney As Double
    Dim iSumMoney   As Long
    Dim sCouponNum  As String
    
    On Error GoTo ErrRtn
    
    CouponMoney = 0
    
    iSumMoney = txtTotalPay.Value '합계금액
    
    varTemp = Split(txtCouponNo.Text, vbNewLine)
    
    For nIndex = 0 To UBound(varTemp)
        sCouponNum = CStr(varTemp(nIndex))
        
        If sCouponNum <> "" Then
            Select Case CheckCouponNumber(sCouponNum)
                Case -1
                    MsgBox "쿠폰 번호 오류 [" & sCouponNum & "]", vbInformation, "확인"
                    
                Case -2
                    MsgBox "쿠폰 사용기간 오류 [" & sCouponNum & "]", vbInformation, "확인"
                
                Case Else
                    Query = "SELECT * FROM TB_쿠폰자료"
                    Query = Query & " WHERE 접수일자 = '" & strDate & "'"
                    Query = Query & "   AND 쿠폰번호 = '" & sCouponNum & "'"
                    Query = Query & "   AND 가맹점코드 = '" & 가맹점정보.가맹점코드 & "'"
                    Set SUBRs = New ADODB.RecordSet
                    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
                    
                    If SUBRs.EOF Then
                        Query = "INSERT INTO TB_쿠폰자료("
                        Query = Query & "  접수일자"
                        Query = Query & ", 지사코드"
                        Query = Query & ", 가맹점코드"
                        Query = Query & ", 쿠폰번호"
                        Query = Query & ", 택번호"
                        Query = Query & ", 쿠폰단가"
                        Query = Query & ", 쿠폰금액"
                        Query = Query & ", 고객코드"
                        Query = Query & ", 고객이름"
                        Query = Query & ", 접수금액"
                        Query = Query & ", 본사전송여부"
                        Query = Query & ", 전송일자) VALUES ("
                        Query = Query & "  '" & strDate & "'"
                        Query = Query & ", '" & 가맹점정보.지사코드 & "'"
                        Query = Query & ", '" & 가맹점정보.가맹점코드 & "'"
                        Query = Query & ", '" & sCouponNum & "'"
                        Query = Query & ", '" & 가맹점정보.택코드 & "'"
                        Query = Query & ",  " & Get_CouponCost(sCouponNum)
                        Query = Query & ",  " & Get_CouponMoney(sCouponNum)
                        Query = Query & ", '" & 고객코드 & "'"
                        Query = Query & ", '" & Trim(frm접수.txtName.Text) & "'"
                        Query = Query & ",  " & iSumMoney
                        Query = Query & ", 'N'"
                        Query = Query & ", '')"
                        ADOCon.Execute Query
                    End If
                    SUBRs.Close
                    Set SUBRs = Nothing
            End Select
        End If
    Next nIndex
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

'잔액
Private Sub txtBalance_Change()
    txtBalance2.Value = txtBalance.Value - txtCash.Value - txtCard.Value
End Sub

Private Sub txtBalance2_Change()
    If txtBalance2.Value = 0 Then
        'cmdAction(0).Enabled = True
        cmdAction(1).Enabled = False
        cmdAction(2).Enabled = False
    Else
        'cmdAction(0).Enabled = False
        
        If txtCash.Value > 0 Or txtCard.Value > 0 Then
            cmdAction(1).Enabled = False
        Else
            cmdAction(1).Enabled = True
        End If
        
        If txtCash.Value = 0 And txtCard.Value = 0 Then
            cmdAction(2).Enabled = False
        Else
            cmdAction(2).Enabled = True
        End If
    End If
End Sub

Private Sub txtCard_Change()
    txtBalance2.Value = txtBalance.Value - txtCash.Value - txtCard.Value

    If txtCash.Value = 0 And txtCard.Value = 0 Then
        btnExit.Enabled = True
    Else
        btnExit.Enabled = False
    End If
End Sub

Private Sub txtCash_Change()
    If txtCash.Value = 0 And txtCard.Value = 0 Then
        btnExit.Enabled = True
    Else
        btnExit.Enabled = False
    End If
End Sub

Private Sub txtCouponNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        Call CouponNo_Proc
    End If
End Sub

'접수금액
Private Sub txtTotalPay_Change()
    txtBalance.Value = txtTotalPay.Value - txtMileage.Value - txtCoupon.Value
End Sub

'마일리지
Private Sub txtMileage_Change()
    txtBalance.Value = txtTotalPay.Value - txtMileage.Value - txtCoupon.Value
End Sub

'쿠폰금액
Private Sub txtCoupon_Change()
    If txtCoupon.Value = 0 Then
        txtBalance.Value = txtTotalPay.Value - txtMileage.Value - txtCoupon.Value
    Else
        If txtCoupon.Tag = "MONEY" Then
            txtBalance.Value = txtTotalPay2.Value - txtMileage.Value - txtCoupon.Value - txtSetDC.Value
        Else
            txtBalance.Value = txtTotalPay2.Value - txtMileage.Value - txtCoupon.Value
        End If
    End If
    txtDCTotal.Value = txtDC.Value + txtSetDC.Value + txtCoupon.Value
End Sub

'세트할인 금액
Private Sub txtSetDC_Change()
    txtDCTotal.Value = txtDC.Value + txtSetDC.Value + txtCoupon.Value
End Sub

'에누리 할인
Private Sub txtDC_Change()
    txtDCTotal.Value = txtDC.Value + txtSetDC.Value + txtCoupon.Value
End Sub

'할인합계
Private Sub txtDCTotal_Change()
    txtBalance.Value = txtTotalPay2.Value - txtMileage.Value - txtSetDC.Value - txtCoupon.Value
End Sub

'***********************************************************************************************

'받은금액
Private Sub txtIncome_Change()
    If txtIncome.Value >= (txtBalance.Value - txtCard.Value) Then
        txtBalance2.Value = 0                              '결제후 잔액
        txtCash.Value = txtBalance.Value - txtCard.Value   '현금결제
        txtChange.Value = txtIncome.Value - txtCash.Value  '거스름돈
    Else
        txtChange.Value = 0                                                    '거스름돈
        txtCash.Value = txtIncome.Value                                        '현금결제
        txtBalance2.Value = txtBalance.Value - (txtCash.Value + txtCard.Value) '결제후 잔액
    End If
End Sub

'
Private Sub txtIncome_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        KeyAscii = 0
'
'        If txtCash.Value = 0 And txtCard.Value = 0 Then
'            Rtn = MsgBox("결제 : 후불" & vbNewLine & vbNewLine & "영수증을 출력 하시겠습니까?", vbQuestion + vbYesNo + vbDefaultButton1, "확인")
'
'            If Rtn = vbYes Then
'                Call cmdAction_Click(1)
'            End If
'
'        ElseIf txtBalance.Value > (txtCash.Value + txtCard.Value) Then
'            Rtn = MsgBox("결제 : 부분결제" & vbNewLine & vbNewLine & "영수증을 출력 하시겠습니까?", vbQuestion + vbYesNo + vbDefaultButton1, "확인")
'
'            If Rtn = vbYes Then
'                Call cmdAction_Click(2)
'            End If
'
'        Else
'            Rtn = MsgBox("결제 : 완불" & vbNewLine & vbNewLine & "영수증을 출력 하시겠습니까?", vbQuestion + vbYesNo + vbDefaultButton1, "확인")
'
'            If Rtn = vbYes Then
'                Call cmdAction_Click(0)
'            End If
'        End If
'    End If
End Sub

Private Function 접수영수증_Report(iPaper As Integer, 접수일자 As String) As String
    On Error GoTo ErrRtn
    Dim ESC      As String * 1
    

    
    Dim tmp      As String
    Dim 이전미수 As String
    Dim 접수수량 As Integer
    Dim 접수금액 As String
    
    Dim 현금결제 As String
    Dim 카드결제 As String
    Dim 쿠폰결제 As String
    
    Dim 사용마일 As String
    Dim 누적마일 As String
    Dim 가능마일 As String
    
    Dim 카드번호 As String
    
    Dim 받은금액 As String
    Dim 거스름돈 As String
    
    Dim 당일미수 As String
        
    Dim 전화번호     As String
    Dim 전화번호출력 As String
    Dim 운동화세탁안내 As Boolean
    
    Dim sE       As String
    Dim PrintMsg     As String
    
    운동화세탁안내 = False
    
    
    전화번호출력 = GetIniStr("Printer", "TelPrint", "Y", iniFile)
    

    
    If iPaper = 1 Then
        If 가맹점정보.지사코드 = M_COUPON_KLENZ_CODE Then '크렌즈갤러리
            PrintMsg = PrintMsg & PrintTitle2("크렌즈갤러리 - 세탁물 접수증(고객용)")
        Else
            PrintMsg = PrintMsg & PrintTitle2("크린에이드 - 세탁물 접수증(고객용)")
        End If
    Else
        If 가맹점정보.지사코드 = M_COUPON_KLENZ_CODE Then '크렌즈갤러리
            PrintMsg = PrintMsg & PrintTitle2("크렌즈갤러리 - 세탁물 접수증(보관용)")
        Else
            PrintMsg = PrintMsg & PrintTitle2("크린에이드 - 세탁물 접수증(보관용)")
        End If
    End If
    
    
    '--------------------------------------------------------------------------------------------------------
    Query = "SELECT    ISNULL(가맹점명, '')     AS 가맹점명"
    Query = Query & ", ISNULL(매장전화번호, '') AS 매장전화번호"
    Query = Query & ", ISNULL(사업장주소, '')   AS 사업장주소"
    Query = Query & ", ISNULL(사업자번호, '')   AS 사업자번호"
    Query = Query & ", ISNULL(대표자명, '')     AS 대표자명"
    Query = Query & " FROM TB_기본정보"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF Then
        PrintMsg = PrintMsg & PrintString("상 호 명 : ", 1, True)
        PrintMsg = PrintMsg & PrintString("전화번호 : ", 1, True)
        PrintMsg = PrintMsg & PrintString("주    소 : ", 1, True)
    Else
        PrintMsg = PrintMsg & PrintString("상 호 명 : " + ADORs!가맹점명, 1, True)
        PrintMsg = PrintMsg & PrintString("사업자No : " + ADORs!사업자번호, 1, True)
        PrintMsg = PrintMsg & PrintString("대 표 자 : " + ADORs!대표자명, 1, True)
        PrintMsg = PrintMsg & PrintString("전화번호 : " + ADORs!매장전화번호, 1, True)
        PrintMsg = PrintMsg & PrintString("주    소 : " + ADORs!사업장주소, 1, True)
    End If
    ADORs.Close
    Set ADORs = Nothing
    
    ' 48자 출력               12345678901234567890123456789012345678901234567890
    PrintMsg = PrintMsg & PrintString("==============================================", 1, True)
    PrintMsg = PrintMsg & PrintString("접수일자 : " + Format(Now, "YYYY년 MM월 DD일 AM/PM hh:mm") + " " + IIf(접수일자 > Format(Date, "yyyy-MM-dd"), 접수일자, ""), 1, True)
    PrintMsg = PrintMsg & PrintString("찾을날짜 : " + Format(frm접수결제.dtpDay.Value, "YYYY년 MM월 DD일"), 1, True)
    PrintMsg = PrintMsg & PrintString("고객코드 : " + frm접수.txtCode.Text, 1, True)
    
    PrintMsg = PrintMsg & PrintCustomer(전화번호출력, frm접수.txtName.Text, Trim(frm접수.txtTel.Text), Trim(frm접수.txtHP.Text), frm접수.txtAddress.Text)
    
    PrintMsg = PrintMsg & PrintString("==============================================", 1, True)
    PrintMsg = PrintMsg & PrintString("택번호  의류/상표         작업   색상     금액", 1, True)
    PrintMsg = PrintMsg & PrintString("----------------------------------------------", 1, True)
    
    접수수량 = 0
    
    With frm접수.sprGrid
        For i = 1 To .MaxRows
            Dim TempMoney As String
            .Row = i
            
            .Col = 1
            If Trim(.Text) = "" Then Exit For
            
            접수수량 = 접수수량 + 1
            
            '*********************************************************
            '* 택번호
            '*********************************************************
            .Col = 2
            'PrintMsg = PrintMsg & ESC + "!" + Chr$(8)              'Selects Emphasized mode
            PrintMsg = PrintMsg & PrintString(.Text + " ", 1)
            'PrintMsg = PrintMsg & ESC + "!" + Chr$(0)              'Cancels Emphasized mode
        
            '*********************************************************
            '* 품명
            '*********************************************************
            .Col = 1
            If LenH(.Text) >= 18 Then
                tmp = MidH(.Text, 1, 18)
            Else
                tmp = Trim(.Text) + String(18 - LenH(.Text), " ")
            End If
            
            tmp = Replace(tmp, vbNullChar, " ")
            PrintMsg = PrintMsg & PrintString(tmp + "", 1)
            
            '*********************************************************
            '* 내용
            '*********************************************************
            .Col = 5
            If LenH(.Text) >= 6 Then
                tmp = MidH(.Text, 1, 6)
            Else
                tmp = Trim(.Text) + String(6 - LenH(.Text), " ")
            End If
            
            If InStr(tmp, "水") > 0 Then tmp = "water "
            PrintMsg = PrintMsg & PrintString(tmp + " ", 1)
            
            '*********************************************************
            '* 색상
            '*********************************************************
            .Col = 3
            If LenH(.Text) >= 4 Then
                tmp = MidH(.Text, 1, 4)
            Else
                tmp = Trim(.Text) + String(4 - LenH(.Text), " ")
            End If
            
            PrintMsg = PrintMsg & PrintString(tmp + " ", 1)

            '*********************************************************
            '* 금액
            '*********************************************************
            .Col = 20
            
            If Len(.Text) > 8 Then
                PrintMsg = PrintMsg & PrintString(.Text, 1, True)
            Else
                PrintMsg = PrintMsg & PrintString(String(8 - LenH(.Text), " ") + .Text, 1, True)
            End If
            .Col = 6
            TempMoney = Replace(.Text, ",", "")
            '*********************************************************
            '* 상표
            '*********************************************************
            .Col = 7
            
            If Trim(.Text) <> "" Then
                PrintMsg = PrintMsg & PrintString("        - " + .Text, 1, True)
            End If

            '*********************************************************
            '* 오점
            '*********************************************************
            .Col = 19
            
            If Trim(.Text) <> "" Then
                PrintMsg = PrintMsg & PrintString("        - " + .Text, 1, True)
            End If
            
            '*********************************************************
            '* 운동화세탁안내
            '*********************************************************
            .Col = 8
            
            If Left(Trim(.Text), 2) = "a0" Then 운동화세탁안내 = True
            
            .Col = 20
            If Val(Replace(.Text, ",", "")) > Val(TempMoney) Then
                Dim Calc As String
                
                Calc = "-" + Format(Str(Val(Replace(.Text, ",", "")) - TempMoney), "#,##0")
                If Len(Calc) > 8 Then
                Else
                    Calc = String(8 - LenH(Calc), " ") + Calc
                End If
'                TempMoney = Format(Str(Val(TempMoney)), "#,##0")
'                If Len(CStr(.Text)) > 7 Then
'                Else
'                    TempMoney = String(7 - LenH(CStr(.Text)), " ") + .Text
'                End If
                PrintMsg = PrintMsg & PrintString("        * 할인금액 " + String(19, " ") + Calc, 1, True)
                'Print_Msg = Print_Msg & PrintString("        * 정상금액 :" + CStr(TempMoney) + "/ 할인금액 :" + Calc, 1, True)
            End If
        
        Next i
    End With
    
    접수금액 = frm접수결제.txtTotalPay.Text
    이전미수 = frm접수.txtMisu.Text
    당일미수 = frm접수결제.txtBalance2.Text
    현금결제 = frm접수결제.txtCash.Text
    카드결제 = frm접수결제.txtCard.Text
    쿠폰결제 = frm접수결제.txtCoupon.Text
    
    
    
    받은금액 = frm접수결제.txtIncome.Text
    거스름돈 = frm접수결제.txtChange.Text
    
    사용마일 = frm접수결제.txtMileage.Text
    
    누적마일 = Format(고객정보.누적마일리지, "#,##0")     'frm접수.txtTotalMileage.Text
    가능마일 = Format(고객정보.사용가능마일리지, "#,##0") 'frm접수.txtUseMileage.Text
    
    PrintMsg = PrintMsg & PrintString("----------------------------------------------", 1, True)
    
    If CDbl(frm접수결제.txtSetDC.Text) > 0 Then
        PrintMsg = PrintMsg & PrintString("정상금액 :" + Format(frm접수결제.txtTotalPay2.Text, "@@@@@@@@@@") + "원/ 할인금액 :" + Format(frm접수결제.txtSetDC.Text, "@@@@@@@@@@") + "원", 1, True)
        PrintMsg = PrintMsg & PrintString("----------------------------------------------", 1, True)
    End If
    PrintMsg = PrintMsg & PrintString("이전미수 :" + Format(이전미수, "@@@@@@@@@@") + "원/ 당일미수 :" + Format(당일미수, "@@@@@@@@@@") + "원", 1, True)
    PrintMsg = PrintMsg & PrintString("접수수량 :" + Format(접수수량, "@@@@@@@@@@") + "점/ 접수금액 :" + Format(접수금액, "@@@@@@@@@@") + "원", 1, True)
    PrintMsg = PrintMsg & PrintString(String(24, " ") + "받은금액 :" + Format(받은금액, "@@@@@@@@@@") + "원", 1, True)
    
    PrintMsg = PrintMsg & PrintString("누적마일 :" + Format(누적마일, "@@@@@@@@@@") + "원/ 사용마일 :" + Format(사용마일, "@@@@@@@@@@") + "원", 1, True)
    PrintMsg = PrintMsg & PrintString("가능마일 :" + Format(가능마일, "@@@@@@@@@@") + "원", 1, True)

    
    If Trim(카드결제) <> "0" Then
        PrintMsg = PrintMsg & PrintString(String(24, " ") + "카드결제 :" + Format(카드결제, "@@@@@@@@@@") + "원", 1, True)
    End If
    If Trim(쿠폰결제) <> "0" Then
        PrintMsg = PrintMsg & PrintString(String(24, " ") + "쿠폰결제 :" + Format(쿠폰결제, "@@@@@@@@@@") + "원", 1, True)
    End If
    
    PrintMsg = PrintMsg & PrintString(String(24, " ") + "현금결제 :" + Format(현금결제, "@@@@@@@@@@") + "원", 1, True)
    PrintMsg = PrintMsg & PrintString(String(24, " ") + "거스름돈 :" + Format(거스름돈, "@@@@@@@@@@") + "원", 1, True)
    PrintMsg = PrintMsg & PrintString("===============================================", 1, True)
    PrintMsg = PrintMsg & PrintLineFeed

    
    PrintMsg = PrintMsg & PrintString("※ 인도예정일은 세탁물의 오염정도에 따라 다소", 1, True)
    PrintMsg = PrintMsg & PrintString("   지연될 수 있습니다.", 1, True)
    
    PrintMsg = PrintMsg & PrintLineFeed
    
    ' 운동화 세탁 안내
    If 운동화세탁안내 And iPaper = 2 Then
        PrintMsg = PrintMsg & 운동화세탁안내_Report
    End If
    
    PrintMsg = PrintMsg & PrintLineFeed(4)
    
    PrintMsg = PrintMsg & PrintCut
    
    
    접수영수증_Report = PrintMsg

    Exit Function
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
    Screen.MousePointer = 0
End Function



Private Function 운동화세탁안내_Report() As String
    Dim bReport As Boolean
    Dim vText   As Variant
    Dim nRow    As Long
    Dim msg As String
    Dim ReturnMsg As String
    bReport = False

    On Error GoTo ErrRtn
    
 
    
    With frm접수.sprGrid
        For nRow = 1 To .MaxRows
            .GetText 8, nRow, vText
            
            If Trim(vText) = "" Then Exit For
            
            If Left(CStr(vText), 2) = "a0" Then
                bReport = True
                Exit For
            End If
 
        Next nRow
    End With
 
    If bReport = False Then Exit Function
    
    ReturnMsg = ReturnMsg & PrintString("[ 구두/운동화 세탁 안내 ]", 6, True)
    ReturnMsg = ReturnMsg & PrintLineFeed
    ReturnMsg = ReturnMsg & PrintString("구두와 운동화는 물세탁을 합니다. 세무, 가죽, 면", 1)
    ReturnMsg = ReturnMsg & PrintString("소재는 세탁 후 코팅 탈락 또는 색 벗겨짐 현상이", 1, True)
    ReturnMsg = ReturnMsg & PrintString("일어날 수 있으며 탈변색, 이염, 경화 될 수 있습니다.", 1, True)
    ReturnMsg = ReturnMsg & PrintLineFeed
    ReturnMsg = ReturnMsg & PrintString("위 내용을 숙지하여 세탁에 동의합니다.", 1)
    ReturnMsg = ReturnMsg & PrintLineFeed(2)
    ReturnMsg = ReturnMsg & PrintString("고객 서명 : __________________________________", 1)
    ReturnMsg = ReturnMsg & PrintLineFeed(2)
    운동화세탁안내_Report = ReturnMsg
    
    Screen.MousePointer = vbDefault
    Exit Function
    
ErrRtn:
    Screen.MousePointer = vbDefault
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

End Function

