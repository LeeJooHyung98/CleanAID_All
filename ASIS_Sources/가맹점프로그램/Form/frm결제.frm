VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm결제 
   BorderStyle     =   1  '단일 고정
   Caption         =   "결제"
   ClientHeight    =   4995
   ClientLeft      =   1635
   ClientTop       =   3255
   ClientWidth     =   8970
   DrawWidth       =   3
   FillColor       =   &H00C0C0C0&
   Icon            =   "frm결제.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form10"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   8970
   StartUpPosition =   1  '소유자 가운데
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   4995
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8970
      _ExtentX        =   15822
      _ExtentY        =   8811
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm결제.frx":0A02
      Begin Threed.SSPanel SSPanel1 
         Height          =   1185
         Left            =   15
         TabIndex        =   1
         Top             =   3795
         Width           =   8940
         _ExtentX        =   15769
         _ExtentY        =   2090
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdAction 
            Height          =   780
            Index           =   0
            Left            =   4950
            TabIndex        =   2
            Top             =   240
            Width           =   1830
            _Version        =   851970
            _ExtentX        =   3228
            _ExtentY        =   1376
            _StockProps     =   79
            Caption         =   " 완  불"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Picture         =   "frm결제.frx":0A54
         End
         Begin XtremeSuiteControls.PushButton cmdAction 
            Height          =   780
            Index           =   1
            Left            =   6915
            TabIndex        =   3
            Top             =   240
            Width           =   1830
            _Version        =   851970
            _ExtentX        =   3228
            _ExtentY        =   1376
            _StockProps     =   79
            Caption         =   "후  불"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   0
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            CancelError     =   -1  'True
         End
         Begin Threed.SSOption SSOption1 
            Height          =   315
            Left            =   1110
            TabIndex        =   5
            Top             =   150
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   556
            _Version        =   262144
            Font3D          =   3
            ForeColor       =   0
            Windowless      =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "영수증 출력"
            Value           =   -1
         End
         Begin Threed.SSOption SSOption2 
            Height          =   315
            Left            =   1110
            TabIndex        =   4
            Top             =   615
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   556
            _Version        =   262144
            Font3D          =   3
            Windowless      =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "영수증 미출력"
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   3765
         Left            =   15
         TabIndex        =   6
         Top             =   15
         Width           =   8940
         _ExtentX        =   15769
         _ExtentY        =   6641
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtCouponNo 
            Appearance      =   0  '평면
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1695
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   1935
            Width           =   2640
         End
         Begin MSComCtl2.DTPicker dtpMonth 
            Height          =   420
            Left            =   2790
            TabIndex        =   25
            Top             =   105
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   741
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "MM월 dd일"
            Format          =   54984707
            UpDown          =   -1  'True
            CurrentDate     =   40273
         End
         Begin XtremeSuiteControls.PushButton cmdCoupon 
            Height          =   645
            Left            =   1050
            TabIndex        =   24
            Top             =   1920
            Width           =   630
            _Version        =   851970
            _ExtentX        =   1111
            _ExtentY        =   1138
            _StockProps     =   79
            Caption         =   "쿠폰사용"
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
         Begin Threed.SSPanel SSPanel4 
            Height          =   420
            Index           =   4
            Left            =   1065
            TabIndex        =   13
            Top             =   90
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   741
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "출 고 일 "
            PictureBackgroundStyle=   2
            PictureBackground=   "frm결제.frx":114E
            Alignment       =   4
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   420
            Index           =   5
            Left            =   1065
            TabIndex        =   14
            Top             =   555
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   741
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "합계금액 "
            PictureBackgroundStyle=   2
            PictureBackground=   "frm결제.frx":1490
            Alignment       =   4
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   420
            Index           =   6
            Left            =   1065
            TabIndex        =   15
            Top             =   1020
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   741
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "마일리지 "
            PictureBackgroundStyle=   2
            PictureBackground=   "frm결제.frx":17D2
            Alignment       =   4
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   420
            Index           =   7
            Left            =   1065
            TabIndex        =   16
            Top             =   1485
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   741
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "쿠폰금액 "
            PictureBackgroundStyle=   2
            PictureBackground=   "frm결제.frx":1B14
            Alignment       =   4
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   420
            Index           =   8
            Left            =   1065
            TabIndex        =   17
            Top             =   2580
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   741
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "잔    액 "
            PictureBackgroundStyle=   2
            PictureBackground=   "frm결제.frx":1E56
            Alignment       =   4
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   420
            Index           =   9
            Left            =   1065
            TabIndex        =   18
            Top             =   3045
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   741
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "입 금 액 "
            PictureBackgroundStyle=   2
            PictureBackground=   "frm결제.frx":2198
            Alignment       =   4
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.sidbEdit txtMoney 
            Height          =   420
            Left            =   2790
            TabIndex        =   19
            Top             =   3045
            Width           =   1545
            _Version        =   262145
            _ExtentX        =   2725
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
            DataProperty    =   2
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   2
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
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtMisu 
            Height          =   420
            Left            =   2790
            TabIndex        =   20
            Top             =   2580
            Width           =   1545
            _Version        =   262145
            _ExtentX        =   2725
            _ExtentY        =   741
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   192
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
            DataProperty    =   2
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   2
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
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtCoupon 
            Height          =   420
            Left            =   2790
            TabIndex        =   21
            Top             =   1485
            Width           =   1545
            _Version        =   262145
            _ExtentX        =   2725
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
            DataProperty    =   2
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   2
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
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtMileage 
            Height          =   420
            Left            =   2790
            TabIndex        =   22
            Top             =   1020
            Width           =   1545
            _Version        =   262145
            _ExtentX        =   2725
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
            DataProperty    =   2
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   2
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
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtSum 
            Height          =   420
            Left            =   2790
            TabIndex        =   23
            Top             =   555
            Width           =   1545
            _Version        =   262145
            _ExtentX        =   2725
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
            DataProperty    =   2
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   2
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
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   405
            Index           =   0
            Left            =   5295
            TabIndex        =   26
            Top             =   1755
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   714
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "할인전 금액 "
            PictureBackgroundStyle=   2
            PictureBackground=   "frm결제.frx":24DA
            Alignment       =   4
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.sidbEdit txtSetGoods 
            Height          =   405
            Index           =   0
            Left            =   7020
            TabIndex        =   27
            Top             =   1755
            Width           =   1545
            _Version        =   262145
            _ExtentX        =   2725
            _ExtentY        =   714
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   2
            StartText.y     =   4
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
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtSetGoods 
            Height          =   405
            Index           =   1
            Left            =   7020
            TabIndex        =   28
            Top             =   2190
            Width           =   1545
            _Version        =   262145
            _ExtentX        =   2725
            _ExtentY        =   714
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   2
            StartText.y     =   4
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
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtSetGoods 
            Height          =   405
            Index           =   2
            Left            =   7020
            TabIndex        =   29
            Top             =   2625
            Width           =   1545
            _Version        =   262145
            _ExtentX        =   2725
            _ExtentY        =   714
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   2
            StartText.y     =   4
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
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtSetGoods 
            Height          =   405
            Index           =   3
            Left            =   7020
            TabIndex        =   30
            Top             =   3060
            Width           =   1545
            _Version        =   262145
            _ExtentX        =   2725
            _ExtentY        =   714
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   2
            StartText.y     =   4
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
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   405
            Index           =   1
            Left            =   5295
            TabIndex        =   31
            Top             =   2190
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   714
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "세트할인 금액 "
            PictureBackgroundStyle=   2
            PictureBackground=   "frm결제.frx":281C
            Alignment       =   4
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   405
            Index           =   2
            Left            =   5295
            TabIndex        =   32
            Top             =   2625
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   714
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "에누리 금액 "
            PictureBackgroundStyle=   2
            PictureBackground=   "frm결제.frx":2B5E
            Alignment       =   4
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   405
            Index           =   3
            Left            =   5295
            TabIndex        =   33
            Top             =   3060
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   714
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "할인 합계 금액 "
            PictureBackgroundStyle=   2
            PictureBackground=   "frm결제.frx":2EA0
            Alignment       =   4
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdSamSungCard 
            Height          =   570
            Left            =   5265
            TabIndex        =   34
            Top             =   840
            Width           =   3315
            _Version        =   851970
            _ExtentX        =   5847
            _ExtentY        =   1005
            _StockProps     =   79
            Caption         =   "삼성카드 할인(10%)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Image Image1 
            Height          =   720
            Left            =   135
            Picture         =   "frm결제.frx":31E2
            Top             =   165
            Width           =   720
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "원"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4395
            TabIndex        =   11
            Top             =   1590
            Width           =   255
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "원"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4395
            TabIndex        =   10
            Top             =   1125
            Width           =   255
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "원"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4395
            TabIndex        =   9
            Top             =   3135
            Width           =   255
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "원"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   4395
            TabIndex        =   8
            Top             =   2670
            Width           =   255
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "원"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4395
            TabIndex        =   7
            Top             =   660
            Width           =   255
         End
      End
   End
End
Attribute VB_Name = "frm결제"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iRowCount As Integer
Private bBtnDuClick As Boolean

Dim SUB_TOT     As Double
Dim SUB_S_TOT   As Long
Dim GRD_TOT     As Double
Dim GRD_S_TOT   As Long
Dim L_Page      As Integer
Dim S_Line      As Integer
Dim L_Line      As Integer
Dim Bill_Number As Long
Dim TempBan     As Boolean

Dim sPage_count As Single
'Dim Page_Count As Integer

Dim User        As String
'Const DY       As Integer = 300   ' 줄 간격
'Const YS       As Integer = 1700  ' record 시작점
Dim DY          As Integer
Dim YS          As Integer

Const Title_Start_Y As Integer = 1700  ' record 시작점
Const Title_Start_X As Integer = 1000  ' record 시작점

Dim sSEQ        As String
 
Private Sub subBillPrint() '   printer1.bas에서 정의
'   Public FPArray(1 To 100, 1 To 5) As Variant
'
'   Public FPTop As FPTop          '용지 상단내용
'   Public FPBottom As FPBottom    '용지 하단내용

    Dim iCnt1   As Integer
    Dim iCnt2   As Integer
    Dim strDate As String
    
    On Error GoTo ErrRtn
    
    'array 초기화
    For iCnt1 = 1 To 100
        For iCnt2 = 1 To 5
            FPArray(iCnt1, iCnt2) = ""
        Next iCnt2
    Next iCnt1
    
    '상단
    FPTop.Name = frm접수.txtName.Text
    FPTop.Date = Format(Date, "YYYY-MM-DD")
    FPTop.Tel = frm접수.txtTel.Text
    
    
    '출고예정일
    If DayCloseCheck(Format(Date, "YYYY-MM-DD")) = True Then
        strDate = Format(DateAdd("d", 1, Date), "YYYY-MM-DD")
    Else
        strDate = Format(Date, "YYYY-MM-DD")
    End If
    
    Select Case True
        Case frm접수.Option1.Value:   FPTop.Date2 = Format(DateAdd("d", 3, strDate), "MM-DD")
        Case frm접수.Option2.Value:   FPTop.Date2 = Format(DateAdd("d", 4, strDate), "MM-DD")
        Case frm접수.Option3.Value:   FPTop.Date2 = Format(DateAdd("d", 5, strDate), "MM-DD")
        Case frm접수.SSOption1.Value: FPTop.Date2 = Format(DateAdd("d", Trim(Mid(frm접수.cboWorkDay.Text, 1, 2)), strDate), "MM-DD")
    End Select
        
    '하단
    FPBottom.Addr = frm접수.txtAddress.Text
    
    'FPBottom.Name = frmMain.StatusBar1.Panels(2).Text
    'FPBottom.Tel = frmMain.StatusBar1.Panels(5).Text
    
    FPBottom.Name = 대리점정보.대리점명
    FPBottom.Tel = 대리점정보.전화번호
    
    FPBottom.Account0 = frm결제.txtSum.Text
    FPBottom.Account1 = frm결제.txtMoney.Text
    FPBottom.Account2 = frm결제.txtMisu.Text 'frm결제.label3
    
    '----------------------------------------------------------------
    ' 내역
    '----------------------------------------------------------------
    With frm접수.sprGrid
        For iCnt1 = 1 To 30
            .Row = iCnt1
            
            .Col = 1
            If Len(Trim(.Value)) = 0 Then
                Exit For
            End If
            
            .Col = 1: FPArray(iCnt1, 2) = .Value '품명
            .Col = 2: FPArray(iCnt1, 1) = .Value 'TagNo
            .Col = 3: FPArray(iCnt1, 3) = .Value '색상
            .Col = 5: FPArray(iCnt1, 4) = .Value '금액
            .Col = 4: FPArray(iCnt1, 5) = .Value '내용
        Next iCnt1
    End With
    
    Call FormPrint
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("subBillPrint", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub RowCount()
    i = 1
    iRowCount = 0
    
    frm접수.sprGrid.Row = i
    frm접수.sprGrid.Col = 1
    
    Do
        iRowCount = iRowCount + 1
        
        i = i + 1
        
        If i > frm접수.sprGrid.MaxRows Then
            Exit Do
        End If
        
        frm접수.sprGrid.Row = i
    Loop While Len(Trim(frm접수.sprGrid.Text)) >= 1
End Sub

'----------------------------------------------------------------
' PayMode = 0  완불
' PayMode = 1 후불
' 보관증재출력을 위한 sub routine
'----------------------------------------------------------------
Private Sub Receipt_Insert(PayMode As Integer)
    Dim strTel          As String       ' 전화번호
    Dim strName         As String       ' 고객성명
    Dim strinDate       As String       ' 접수일
    Dim strOutdate      As String       ' 인도예정일
    Dim strTagNo        As String       ' 택번호
    Dim strItemname     As String       ' 품명
    Dim strColor        As String       ' 색상
    Dim strMoney        As String       ' 금액
    Dim strContent      As String       ' 내용
    Dim strItemsum      As String       ' 합계점
    Dim strSumMoney     As String       ' 합계금액
    Dim strReceiveMoney As String       ' 수령액
    Dim strResiMoney    As String       ' 잔액
    Dim strUserAddress  As String       ' 주소
    Dim strUserCode     As String       ' 고객번호
    Dim strLabel        As String       ' 상표
    Dim strGoodsCode    As String       ' 상품코드
    Dim strCardMoney    As String       ' 카드 금액
    Dim strMileage      As String
    Dim strAddMileage   As String
    Dim strTotalMileage As String
    Dim strOldMisu      As String
    Dim strSuGumMoney   As String
    Dim strMiSuTotal    As String
    
    Dim sCoupon(2)      As String
    Dim sGroupGoods(3)  As String

    On Error GoTo ErrRtn
    
    '----------------------------------------------------------
    '
    '----------------------------------------------------------
    Query = "SELECT ISNULL(MAX(일련번호),0) As 번호"
    Query = Query & " FROM TB_보관증"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    sSEQ = Val(ADORs!번호) + 1
    
    ADORs.Close
    Set ADORs = Nothing
    
    'Bill_Number = sSEQ
    
    Call RowCount                                                   ' 카운트
    
    strTel = frm접수.txtTel.Text & ""                               ' 전화번호
    strName = frm접수.txtName.Text & ""                             ' 고객성명
    strinDate = Format(Date, "YYYY-MM-DD")                          ' 접수일
    strOutdate = Format(dtpMonth.Value, "mm:dd")                    ' 인도예정일
    strItemsum = CStr(iRowCount)                                    ' 합계점
    strSumMoney = CStr(txtSum.Value)                                ' 합계금액
    
    strReceiveMoney = txtMoney.Value                                '
    
    ' 입금액이 없을 경우 잔액금액을 처리한다.
    If PayMode = 0 And (strReceiveMoney = "0" Or strReceiveMoney = "") Then
        strReceiveMoney = txtMisu.Value                      '
    End If
    
    strResiMoney = IIf(PayMode = 0, "0", txtMisu.Value)      '(완불)/후불
    strUserAddress = frm접수.txtAddress.Text & ""            '주소
    strUserCode = frm접수.txtCode.Text & ""                  '회원코드
    
    Call Fun_고객정보(strUserCode)
    
    strMiSuTotal = 고객정보.미수금
    strOldMisu = CStr(Val(고객정보.미수금) - Val(Replace(strResiMoney, ",", "")))
    strSuGumMoney = Fun_수금액(strUserCode, Replace(strinDate, "-", ""))
    
    ' pds2004 수정 2007-05-28일 카드 금액을 마감시 한번만 입력 하도록 수정
    ' 카드 금액
    'If IsNumeric(mskCard.Text) = False Then mskCard.Text = "0"
    'strCardMoney = mskCard.Text
    
    strCardMoney = "0"
    
    Call Fun_UserMileage(고객정보.고객번호 & "") ' 사용마일리지, 마일리지 잔액, 누적 마일리지
    
    strMileage = txtMileage.Value
    strTotalMileage = userMileage.잔액
    strAddMileage = userMileage.총사용금액
    
    ' 쿠폰 정보
    ' 쿠폰 수량, 쿠폰 번호, 쿠폰 금액
    Dim varTemp As Variant
    
    sCoupon(0) = "0"
    sCoupon(1) = ""
    sCoupon(2) = "0"
    
    varTemp = Split(txtCouponNo.Text, vbNewLine)
    sCoupon(2) = txtCoupon.Value
    
    For i = 0 To UBound(varTemp)
        If CheckCouponNumber(CStr(varTemp(i))) = 0 Then
            sCoupon(0) = CStr(Val(sCoupon(0)) + 1)
            sCoupon(1) = sCoupon(1) & "," & CStr(varTemp(i))
        End If
    Next
    
    If sCoupon(1) <> "" Then
        sCoupon(1) = Left(sCoupon(1), Len(sCoupon(1)) - 1)
    End If
    
    With frm접수.sprGrid
        For i = 1 To iRowCount
            
            .Row = i
            .Col = 1:         strItemname = .Text    ' 품명
            .Col = 2:         strTagNo = .Text       ' 택번호
            .Col = 3:         strColor = .Text       ' 색상
            .Col = 4:         strContent = .Text     ' 내용
            .Col = 5:         strMoney = .Value      ' 금액
            .Col = 6:         strLabel = .Text       ' 상표
            
            .Col = 7:         strGoodsCode = .Text                         ' 상품코드
            .Col = 11: sGroupGoods(0) = .Value                      ' 세트관련 내용
            .Col = 12: sGroupGoods(1) = IIf(.Value = "", 0, .Value) ' 세트관련 내용
            .Col = 13: sGroupGoods(2) = IIf(.Value = "", 0, .Value) ' 세트관련 내용
            .Col = 14: sGroupGoods(3) = IIf(.Value = "", 0, .Value) ' 세트관련 내용
            
            '-------------------------------------------------------------
            Query = "INSERT INTO TB_보관증("
            Query = Query & "  일련번호"            '1
            Query = Query & ", 고객전화"            '2
            Query = Query & ", 성명"                '3
            Query = Query & ", 접수일"              '4
            Query = Query & ", 인도예정일"          '5
            Query = Query & ", 택번호"              '6
            Query = Query & ", 품명"                '7
            Query = Query & ", 색상"                '8
            Query = Query & ", 금액"                '9
            Query = Query & ", 내용"                '10
            Query = Query & ", 합계"                '11
            Query = Query & ", 합계금액"            '12
            Query = Query & ", 수령액"              '13
            Query = Query & ", 미수합계"            '14
            Query = Query & ", 전일미수"            '15
            Query = Query & ", 수금액"              '16
            Query = Query & ", 마일리지"            '17
            Query = Query & ", 누적마일리지"        '18
            Query = Query & ", 마일리지잔액"        '19
            Query = Query & ", 잔액"                '20
            Query = Query & ", 대리점명"            '21
            Query = Query & ", 대리점전화"          '22
            Query = Query & ", 카드금액"            '23
            Query = Query & ", 상표"                '24
            Query = Query & ", CouponCnt"           '25
            Query = Query & ", CouponNumber"        '26
            Query = Query & ", CouponMoney"         '27
            Query = Query & ", 세트구분"            '28
            Query = Query & ", 세트금액1"           '29
            Query = Query & ", 세트금액2"           '30
            Query = Query & ", 정상가격"            '31
            Query = Query & ", 세트Key"             '32
            Query = Query & ", 상품코드"            '33
            Query = Query & ", 전체정상금액"        '34
            Query = Query & ", 전체세트금액1"       '35
            Query = Query & ", 전체세트할인"        '36
            Query = Query & ", 전체세트에누리할인)" '37
            Query = Query & "VALUES ("
            Query = Query & "  '" & sSEQ & "'"                      '1
            Query = Query & ", '" & strTel & "'"                    '2
            Query = Query & ", '" & Replace(strName, "'", "") & "'" '3
            Query = Query & ", '" & strinDate & "'"                 '4
            Query = Query & ", '" & strOutdate & "'"                '5
            Query = Query & ", '" & strTagNo & "'"                  '6
            Query = Query & ", '" & strItemname & "'"               '7
            Query = Query & ", '" & strColor & "'"                  '8
            Query = Query & ",  " & strMoney                        '9
            Query = Query & ", '" & strContent & "'"                '10
            Query = Query & ",  " & strItemsum                      '11
            Query = Query & ",  " & strSumMoney                     '12
            Query = Query & ",  " & strReceiveMoney                 '13
            
            Query = Query & ",  " & strMiSuTotal                    '14
            Query = Query & ",  " & strOldMisu                      '15
            Query = Query & ",  " & strSuGumMoney                   '16
            
            Query = Query & ",  " & strMileage                      '17
            Query = Query & ",  " & strAddMileage                   '18
            Query = Query & ",  " & strTotalMileage                 '19
            Query = Query & ",  " & strResiMoney                    '20
            Query = Query & ", '" & strUserAddress & "'"            '21
            Query = Query & ", '" & strUserCode & "'"               '22
            Query = Query & ",  " & strCardMoney                    '23
            
            Query = Query & ", '" & Replace(strLabel, "'", "") & "'" '24
            
            Query = Query & ", '" & sCoupon(0) & "'"                 '25
            Query = Query & ", '" & sCoupon(1) & "'"                 '26
            Query = Query & ",  " & sCoupon(2)                       '27
            
            Query = Query & ", '" & sGroupGoods(0) & "'"             '28
            Query = Query & ",  " & sGroupGoods(1)                   '29
            Query = Query & ",  " & sGroupGoods(2)                   '30
            Query = Query & ",  " & sGroupGoods(3)                   '31
            
            Query = Query & ", '" & m_GSGMoney.d세트Key & "'"        '32
            Query = Query & ", '" & strGoodsCode & "'"               '33
            Query = Query & ",  " & m_GSGMoney.d전체금액             '34
            Query = Query & ",  " & m_GSGMoney.d세트금액             '35
            Query = Query & ",  " & m_GSGMoney.d세트할인금액         '36
            Query = Query & ",  " & m_GSGMoney.d에누리할인금액 & ")" '37
            ADOCon.Execute Query
        Next i
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("Receipt_Insert", Err.Source, Err.Number, Err.Description)
End Sub

Private Function chkItem(ByVal strDate As String, ByVal CustCode As String) As Boolean
    Query = "SELECT COUNT(고객번호) AS DataCount "
    Query = Query & " FROM TB_입출고 "
    Query = Query & " WHERE 입고일  >= '" & strDate & "' "
    Query = Query & "   AND (확인 = '' OR 확인 IS NULL)"
    Query = Query & "   AND 고객번호 = '" & CustCode & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    chkItem = IIf(Rs!DataCount > 0, True, False)
    
    Rs.Close
    Set Rs = Nothing
End Function

'----------------------------------------------------------------------------------------
'
'
'----------------------------------------------------------------------------------------
Private Function Receive_Update(strDate As String, PayType As Integer) As Boolean
    Dim iCnt        As Integer
    Dim strCusNo    As String    ' 고객번호
    Dim strName     As String    ' 품명
    Dim strColor    As String    ' 색상
    Dim strTagNo    As String    ' 택번호
    Dim strContent  As String    ' 세탁내용
    Dim douMoney    As Double    ' 가격
    Dim strMark     As String    ' 상표
    Dim strState    As String    ' 상태
    Dim strRDate    As String
    Dim douSuMoney  As Double    ' 수선금액
    
    Dim Parent_TAG  As String    '부모택번호
    
    Dim msg         As String
    Dim strinDate   As String
    Dim strCode     As String
    Dim sCupon      As String    ' 쿠폰번호
    
    Dim sGroupGoods(3) As String
    
    On Error GoTo ErrRtn
    
    Receive_Update = False
    
    Erase sGroupGoods
    
    strState = IIf(PayType = 0, "完", "未") '
    strCusNo = frm접수.txtCode.Text         '
    
    iCnt = 1
    douMoney = 0
    
    frm접수.sprGrid.Row = iCnt
    frm접수.sprGrid.Col = 1: strName = Trim(frm접수.sprGrid.Text) '품명
    
    Select Case True
        Case frm접수.Option1.Value:   strRDate = Format(DateAdd("d", 3, strDate), "YYYY-MM-DD")
        Case frm접수.Option2.Value:   strRDate = Format(DateAdd("d", 4, strDate), "YYYY-MM-DD")
        Case frm접수.Option3.Value:   strRDate = Format(DateAdd("d", 5, strDate), "YYYY-MM-DD")
        Case frm접수.SSOption1.Value: strRDate = Format(DateAdd("d", Trim(Mid(frm접수.cboWorkDay.Text, 1, 2)), strDate), "YYYY-MM-DD")
    End Select
    
    m_GSGMoney.d세트Key = Format(Now, "YYYY-MM-DD hh:mm:ss")
    
    While Len(Trim(strName)) > 0 And iCnt <= frm접수.sprGrid.MaxRows
        With frm접수.sprGrid
            .Row = iCnt
            .Col = 1:    strName = Trim(.Text) & ""               '
            .Col = 2:    strTagNo = .Text & ""                    '
            .Col = 3:    strColor = .Text & ""                    '
            .Col = 4:    strContent = .Text & ""                  '
            .Col = 5:    douMoney = .Value & ""                   '
            .Col = 6:    strMark = .Text & ""                     '
            .Col = 7:    strCode = .Text & ""                     '
            .Col = 9:    douSuMoney = IIf(.Value = "", 0, .Value) ' 수선 금액 별도 입력
            .Col = 15:   Parent_TAG = .Text & ""                  ' 부모택번호
            
            .Col = 11: sGroupGoods(0) = .Text  ' ex. 6-01, 5-01, 5-02
            .Col = 12: sGroupGoods(1) = .Value '  세트 할인률을 기준으로 계산한 금액(10원단위 포함)
            .Col = 13: sGroupGoods(2) = .Value ' 원단위 절사후 다시 계산한 금액
            .Col = 14: sGroupGoods(3) = .Value ' 세트관련 내용
        End With
        
        strMark = Replace(strMark, "'", "")
'       strMark = Replace(strMark, ",", "")
        
        ' 택번호 수정시 키값이 택번호와 일자로 되어 있기 때문에 insert 명령어가 안먹힘
        ' 해결 방법 1. update로 처리한다 (이전 취소 내용 삭제됨)
        
        '---------------------------------------------------------
        '
        '---------------------------------------------------------
        Query = "SELECT * FROM TB_입출고"
        Query = Query & " WHERE 입고일 = '" & strDate & "'"
        Query = Query & "   AND 택번호 = '" & strTagNo & "'"
        Set SUBRs = New ADODB.Recordset
        SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        If SUBRs.EOF Then
            Query = "INSERT INTO TB_입출고(입고일, 고객번호, 품명, 택번호, 색상,  내용, 금액, 상표, 코드, 결제여부,판매취소, 입고예정일, "
            Query = Query & " 수선금액, 외주운동화마진, 세트구분, 세트금액1, 세트금액2, 정상가격, 세트Key, 부모택번호, 근무자명) VALUES ("
            Query = Query & "  '" & strDate & "'"
            Query = Query & ", '" & strCusNo & "'"
            Query = Query & ", '" & strName & "'"
            Query = Query & ", '" & strTagNo & "'"
            Query = Query & ", '" & strColor & "'"
            Query = Query & ", '" & strContent & "'"
            Query = Query & ",  " & douMoney
            Query = Query & ", '" & strMark & "'"
            Query = Query & ", '" & strCode & "'"
            Query = Query & ", '" & strState & "'"
            Query = Query & ", '" & " " & "'"
            Query = Query & ", '" & strRDate & "'"
            Query = Query & ",  " & douSuMoney
            Query = Query & ",  " & Val(대리점정보.외주운동화마진)
            Query = Query & ", '" & sGroupGoods(0) & "'"
            Query = Query & ", '" & sGroupGoods(1) & "'"
            Query = Query & ", '" & sGroupGoods(2) & "'"
            Query = Query & ", '" & sGroupGoods(3) & "'"
            Query = Query & ", '" & m_GSGMoney.d세트Key & "'"
            Query = Query & ", '" & Parent_TAG & "'"
            Query = Query & ", '" & strManager & "')"
            ADOCon.Execute Query
        
        ElseIf SUBRs!판매취소 = "Y" Then
            Query = "UPDATE TB_입출고 SET"
            Query = Query & "  입고일         = '" & strDate & "'"
            Query = Query & ", 고객번호       = '" & strCusNo & "'"
            Query = Query & ", 품명           = '" & strName & "'"
            Query = Query & ", 택번호           = '" & strTagNo & "'"
            Query = Query & ", 색상           = '" & strColor & "'"
            Query = Query & ", 내용           = '" & strContent & "'"
            Query = Query & ", 금액           =  " & douMoney
            Query = Query & ", 상표           = '" & Replace(strMark, "'", "") & "'"
            Query = Query & ", 코드           = '" & strCode & "'"
            Query = Query & ", 결제여부           = '" & strState & "'"
            Query = Query & ", 외주운동화마진        =  " & Val(대리점정보.외주운동화마진)
'            Query = Query & ", 판매취소    = '" & "" & "'"
            Query = Query & ", 판매취소    = 'R'"
            Query = Query & ", 입고예정일  = '" & strRDate & "'"
            Query = Query & ", 세트구분    = '" & sGroupGoods(0) & "'"
            Query = Query & ", 세트금액1   =  " & sGroupGoods(1)
            Query = Query & ", 세트금액2   =  " & sGroupGoods(2)
            Query = Query & ", 정상가격    =  " & sGroupGoods(3)
            Query = Query & ", 세트Key     = '" & m_GSGMoney.d세트Key & "'"
            Query = Query & ", 수선금액    =  " & douSuMoney
            Query = Query & ", 부모택번호  = '" & Parent_TAG & "'"
            Query = Query & " WHERE 입고일 = '" & strDate & "'"
            Query = Query & "   AND 택번호 = '" & strTagNo & "'"
            ADOCon.Execute Query
            
        Else
            SUBRs.Close
            Set SUBRs = Nothing
            
            ' 택번호를 잘못 수정하였을 경우 오류 메시지.
            MsgBox "[ " & strTagNo & " ]" & "이미 사용한 택번호 입니다. 택번호를 변경하여 주십시요", vbInformation
            Receive_Update = False
            Exit Function
        End If
        SUBRs.Close
        Set SUBRs = Nothing
        
        'Call ERR_SAVE("Receive_Update SQL 실행문 : " & Query)
        
        ' 반자가 있을 경우 마일리지 적용을 하지 않기 위하여.
        ' 짜집기가 되어 버렸네 ㅡㅡ
        ' 2007-03-26일 반품도 마일리지 적립하기로 변경
        'If InStr(strContent, "반") > 0 Then TempBan = True
        
        iCnt = iCnt + 1
        
        frm접수.sprGrid.Row = iCnt
        frm접수.sprGrid.Col = 1: strName = Trim(frm접수.sprGrid.Text) '품명
        
        If iCnt > frm접수.sprGrid.MaxRows Then
            strName = ""
        End If
    Wend
    
    Receive_Update = True
    
    Exit Function
    
ErrRtn:
    Receive_Update = False
    
    MsgBox Err.Description, vbInformation, "확인"
    
    'Call ERR_SAVE("Receive_Update ERR_INDEX(" & CStr(ERR_INDEX) & ") " & Err.Description & Query)
End Function

'*************************************************************************************
'제목:이용실적기록
'기능:연도와 고객번호를 가지고 이용실적Table에 write
'처리:1.이용횟수=한번계산될때마다 1씩증가
'     2.이용금액=이용금액+합계금액
'*************************************************************************************
Private Sub UseAccountUpdate()
    Dim useCnt     As Integer
    Dim useMoney   As Long
    
    On Error GoTo ErrRtn
    
    With frm접수
        If .sprYear.MaxRows = 0 Then
            useCnt = 1               '이용횟수
            useMoney = txtSum.Value  '이용금액
        Else
            .sprYear.Row = 1
            .sprYear.Col = 1
            
            If .sprYear.Text = Format(Date, "YYYY") Then
                .sprYear.Col = 2: useCnt = .sprYear.Value + 1              '이용횟수
                .sprYear.Col = 3: useMoney = .sprYear.Value + txtSum.Value '이용금액
            Else
                useCnt = 1               '이용횟수
                useMoney = txtSum.Value  '이용금액
            End If
        End If
    End With
    
    '------------------------------------------------------------------------
    ' 이용실적
    '------------------------------------------------------------------------
    Query = " SELECT * FROM TB_이용실적"
    Query = Query & " WHERE 고객번호 = '" & Trim(frm접수.txtCode.Text) & "'"
    Query = Query & "   AND 연도     = '" & Format(Date, "YYYY") & "'"
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenDynamic, adLockOptimistic
    
    If SUBRs.EOF Then
        SUBRs.AddNew
        
        SUBRs!고객번호 = frm접수.txtCode.Text & "" '
        SUBRs!연도 = Format(Date, "YYYY")          '
        SUBRs!이용횟수 = useCnt & ""               '
        SUBRs!이용금액 = useMoney & ""             '
    Else
        SUBRs!이용횟수 = useCnt & ""               '
        SUBRs!이용금액 = useMoney & ""             '
    End If
        
    SUBRs.Update
    
    SUBRs.Close
    Set SUBRs = Nothing
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("UseAccountUpdate", Err.Source, Err.Number, Err.Description)
    
    Resume Next
End Sub

Private Sub UseMileageUpdate(UserCode As String, strDate As String)
'   1. 마일리지를 저장한다.
'   2. 해당 마일리지 저장후 마일리지가 발생하는지 확인한다.
'   3. 발생할 경우 해당 금액을 생성한다.
'   4. 발생할 경우 히스토리에 저장한다.
'   5. 사용한 마일리지가 있을 경우 해당 내용을 적용한다.

    Dim useMoney            As Long
    Dim dbluserMileage      As Double
    
    Dim dblTmpMileage       As Double
    Dim dblTmpReturnMileage As Double
    Dim dblTmpRtnMileage    As Double
    
    '후불일 경우도 반영이 되어야 한다고해서
    '합계금액에서 마일리지 금액을 뺀것으로함다.
    useMoney = CLng(txtSum.Text) - CLng(txtMileage.Text) - CLng(txtCoupon.Text)  '이용금액
    
    ' 수선 내용을 제외한다.
    ' 다시 반품 환불은 적용하지 않기로해 이전으로 다시 처리한다.
    
    ' 강부장 합의 2007-05-04
    ' 반품 환불 -> 고객에게 출고 처리할 경우    ->  반품 환부로 처리 (마일리지 삭제됨)
    '           -> 본사로 다시 업고치라할 경우  -> 출고로 처리(마일리지 유지) 입고 잡을때 드반으로 입고 처리되며. 마일리지 누적되지 않음
    
    useMoney = useMoney - Get수선금액 ' + Get반품금액
    
    If 대리점정보.마일리지여부 = "Y" Then
        '-------------------------------------------------------------------------------------
        ' 1. 마일리지를 저장한다.
        '-------------------------------------------------------------------------------------
        Query = "SELECT 고객번호 FROM TB_마일리지현황 "
        Query = Query & " WHERE 고객번호 = '" & UserCode & "'"
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        If ADORs.EOF Then
            Query = "INSERT INTO TB_마일리지현황(고객번호"
            Query = Query & ", 총사용금액"
            Query = Query & ", 마일리지"
            Query = Query & ", 최종발생금액"
            Query = Query & ", 최종거래일자"
            Query = Query & ", 발생누계"
            Query = Query & ", 사용마일리지"
            Query = Query & ", 미반환마일리지"
            Query = Query & ", 전송여부)"
            Query = Query & " VALUES ("
            Query = Query & " '" & UserCode & "'" '1
            Query = Query & ", " & useMoney       '2
            Query = Query & ", 0"                 '3
            Query = Query & ", 0"                 '4
            Query = Query & ", '" & strDate & "'" '5
            Query = Query & ", 0"                 '6
            Query = Query & ", 0"                 '7
            Query = Query & ", 0"                 '8
            Query = Query & ", 'N')"              '9
            ADOCon.Execute Query
        Else
            Query = "UPDATE TB_마일리지현황 SET"
            Query = Query & "  총사용금액   = 총사용금액 + " & useMoney
            Query = Query & ", 최종거래일자 = '" & strDate & "'"
            Query = Query & ", 전송여부     = 'N'"
            Query = Query & " WHERE 고객번호 = '" & UserCode & "'"
            ADOCon.Execute Query
        End If
        ADORs.Close
        Set ADORs = Nothing
        
        '-------------------------------------------------------------------------------------
        ' 2. 해당 마일리지 저장후 마일리지가 발생하는지 확인한다.
        '-------------------------------------------------------------------------------------
        Query = "SELECT * FROM TB_마일리지현황 "
        Query = Query & " WHERE 고객번호 = '" & UserCode & "'"
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        If Not ADORs.EOF Then
            If ADORs!총사용금액 >= ADORs!최종발생금액 + NextMileage Then
                dbluserMileage = Fnc_MileagePoint(ADORs!총사용금액, ADORs!최종발생금액, ADORs!고객번호) '3. 발생할 경우 해당 금액을 생성한다.
                
                Call Fun_UserMileage(UserCode) ' 미반환마일리지가 있는지 확인한다.
                
                If userMileage.미반환마일리지 > 0 Then
                    If userMileage.미반환마일리지 <= dbluserMileage Then
                        ' 미반환 마일리지가 있을 경우 반환처리한다.
                        dblTmpMileage = (dbluserMileage - userMileage.미반환마일리지)
                        dblTmpReturnMileage = 0
                        dblTmpRtnMileage = userMileage.미반환마일리지
                    Else
                        '미반환 마일리지가 더 클경우 사용마일리지(가용) 0원으로
                        '처리하고 반환 마일리지에 잔액을 남긴다.
                        dblTmpMileage = 0
                        dblTmpReturnMileage = (userMileage.미반환마일리지 - dbluserMileage)
                        dblTmpRtnMileage = dbluserMileage
                    End If
                Else
                    ' 반환 마일리지가 없을 경우
                    dblTmpMileage = dbluserMileage
                    dblTmpReturnMileage = 0
                    dblTmpRtnMileage = 0
                End If
                    
                '-------------------------------------------------------------------------------------
                ' TB_마일리지현황
                '-------------------------------------------------------------------------------------
                Query = "UPDATE TB_마일리지현황 SET"
                Query = Query & "  마일리지       = 마일리지 + " & dblTmpMileage
                Query = Query & ", 최종발생금액   = " & (ADORs!총사용금액 \ NextMileage) * NextMileage
                Query = Query & ", 미반환마일리지 = " & dblTmpReturnMileage
                Query = Query & ", 발생누계       = 발생누계 + " & dblTmpMileage
                Query = Query & ", 최종거래일자   = '" & strDate & "'"
                Query = Query & ", 전송여부       = 'N'"
                Query = Query & " WHERE 고객번호 = '" & UserCode & "'"
                ADOCon.Execute Query
            
                '-------------------------------------------------------------------------------------
                ' 4. 발생할 경우 히스토리에 저장한다.
                '-------------------------------------------------------------------------------------
                Query = "INSERT INTO TB_마일리지스토리 ("
                Query = Query & "  발생일자"     '1
                Query = Query & ", 고객번호"     '2
                Query = Query & ", 발생마일리지" '3
                Query = Query & ", 사용마일리지" '4
                Query = Query & ", 삭제마일리지" '5
                Query = Query & ", 반환마일리지" '6
                Query = Query & ", 보관증"       '7
                Query = Query & ", 전송여부)"    '8
                Query = Query & " VALUES ("
                Query = Query & "  '" & Format(Now, "YYYY-MM-DD hh:mm:ss") & "'" '1
                Query = Query & ", '" & UserCode & "'"                           '2
                Query = Query & ",  " & dblTmpMileage                            '3
                Query = Query & ", 0"                                            '4
                Query = Query & ", 0"                                            '5
                Query = Query & ", " & dblTmpRtnMileage                          '6
                Query = Query & ", " & Bill_Number                               '7
                Query = Query & ", 'N') "                                        '8
                ADOCon.Execute Query
            End If
        End If
        ADORs.Close
        Set ADORs = Nothing
    End If
    
    ' 마일리지 사용후 다시 사용하지 않는 것으로 수정하였을경우 기존액 누적된 사용자의 마일리지 잔액을
    ' 모두 사용하도록 하기 위하여 아래 노용은 모두 처리 되도록 한다.

'   5. 사용한 마일리지가 있을 경우 해당 내용을 적용한다.
    If Val(txtMileage.Text) > 0 Then
        Query = "UPDATE TB_마일리지현황 SET"
        Query = Query & "  마일리지     = 마일리지 - " & txtMileage.Value
        Query = Query & ", 사용마일리지 = 사용마일리지 + " & txtMileage.Value
        Query = Query & ", 전송여부     = 'N'"
        Query = Query & " WHERE 고객번호 = '" & UserCode & "'"
        ADOCon.Execute Query
        
        '
        Query = "INSERT INTO TB_마일리지스토리 ("
        Query = Query & "  발생일자"
        Query = Query & ", 고객번호"
        Query = Query & ", 발생마일리지"
        Query = Query & ", 사용마일리지"
        Query = Query & ", 삭제마일리지"
        Query = Query & ", 반환마일리지"
        Query = Query & ", 보관증"
        Query = Query & ", 전송여부)"
        Query = Query & " VALUES ('" & Format(Now, "YYYY-MM-DD hh:mm:ss") & "'"
        Query = Query & ", '" & UserCode & "'"
        Query = Query & ", 0, "
        Query = Query & ", " & txtMileage.Value
        Query = Query & ", 0"
        Query = Query & ", 0"
        Query = Query & ", " & Bill_Number & ""
        Query = Query & ", 'N') "
        ADOCon.Execute Query
    End If
End Sub

Private Function Get수선금액()
    Dim dblSuMoney As Double
    
    dblSuMoney = 0
    With frm접수.sprGrid
        For i = 1 To .MaxRows
            .Row = i
            .Col = 2
            
            If .Text = "" Then Exit For
            
            .Col = 9: dblSuMoney = dblSuMoney + Val(Replace(.Text, ",", ""))
        Next i
    End With
    
    Get수선금액 = dblSuMoney
End Function

Private Function Get반품금액()
    Dim dblSuMoney As Double
    Dim Scode       As String
    
    dblSuMoney = 0
    With frm접수.sprGrid
        For i = 1 To .MaxRows
            .Col = 7:   .Row = i
            If .Text = "" Then Exit For
            .Col = 4
            If InStr(.Text, "반") > 0 Then
                Scode = .Text
                dblSuMoney = dblSuMoney + f_dryPrice(Scode)
            End If
        Next i
    End With
    
    Get반품금액 = dblSuMoney
End Function

'-------------------------------------------------
' PayMode = 0 완불
' PayMode = 1 후불
'-------------------------------------------------
Private Sub cmdAction_Click(Index As Integer)
    Dim strCusNo    As String
    Dim strTelNo    As String
    Dim strDate     As String
    Dim PayType     As Integer
    Dim PrintCount  As Integer  ' 프린트 출력 장수
    Dim dblTotal    As Double
    
    PayType = Index
    
    'On Error GoTo ErrRtn
    
    ' 순간적인 더블 클릭 방지
    If bBtnDuClick = True Then Exit Sub
    
    bBtnDuClick = True
    DoEvents
    
    strCusNo = frm접수.txtCode.Text & "" '고객코드
    strTelNo = frm접수.txtTel.Text & ""  '전화번호
    
    '-----------------------------------------------------------------------------------------------------------------------
    ' pds2004 2007-05-08카드 금액을 마감시 한번만 입력 하도록 수정
    ' 미수금 내용을 적용한다.  미수금 = 전체금액-입금액-카드금액-마일리지금액
    'txtmisu.text = CLng(txtSum.text) - CLng(Val(txtMoney.text)) - CLng(Val(mskCard.ClipText)) - CLng(Val(txtMileage.Text))
    
    dblTotal = txtSum.Value - txtMoney.Value - txtMileage.Value - txtCoupon.Value
    
    If dblTotal <= 0 Then
        txtMisu.Value = "0"
    Else
        txtMisu.Value = dblTotal
    End If
    
    If DayCloseCheck(Format(Date, "YYYY-MM-DD")) = True Then
        MsgBox "일마감이 되었으므로 판매내역은 익일로 저장이 됩니다.", vbInformation
        
        strDate = Format(DateAdd("d", 1, Date), "YYYY-MM-DD")
    Else
        strDate = Format(Date, "YYYY-MM-DD")
    End If
        
    Call Tag_Update ' 대리점 정보에 택번호를 저장한다.
    
    If "Error" = Fun_고객정보(strCusNo) Then
        MsgBox "일치하는 전화번호가 없습니다." & Chr(10) & Chr(10) & "다시 입력하세요!"
        bBtnDuClick = False
        Exit Sub
    End If
    
    ' 입출고 테이블에 저장
    If Receive_Update(strDate, PayType) = False Then
        Dim sMSG    As String
        
        sMSG = "입출고 자료 저장중 오류가 발생 하였습니다. " & vbLf & ""
        sMSG = sMSG & "이미 사용중인 택번호를 다시 사 용하였을 경우 발생할 수 있습니다." & vbLf
        sMSG = sMSG & "오류가 지속될경우 프로그램 개발자에게 문의 하여 주십시요." & vbLf
        MsgBox sMSG, vbInformation
        
        'Call ERR_SAVE("cmdAction_Click ERR_INDEX(" & CStr(ERR_INDEX) & ") " & sMSG)
        
        bBtnDuClick = False
        
        Exit Sub
    End If
    
    ' 저정중 오류 확인...
    If chkItem(strDate, strCusNo) = False Then
        sMSG = "입출고 자료 저장중 오류가 발생 하였습니다. " & vbLf & ""
        sMSG = sMSG & "Null을 저장하려고 시도할 경우 발생할 수 있습니다.." & vbLf
        sMSG = sMSG & "오류가 지속될경우 프로그램 개발자에게 문의 하여 주십시요." & vbLf
        MsgBox sMSG, vbInformation
        
        'Call ERR_SAVE("cmdAction_Click ERR_INDEX(" & CStr(ERR_INDEX) & ") " & sMSG)
        
        bBtnDuClick = False
        
        Exit Sub
    End If
    
    Call UseAccountUpdate        ' 이용실적
    Call SaveCouponDate(strDate) '쿠폰 사용 내역 저장   2009-04-20 기능 추가
    
    '------------------------------------------------------
    ' 세트 상품 할인 정보를 저장한다.
    ' 크렌즈 겔러리는 행사에서 제외 처리한다.
    '------------------------------------------------------
    If 대리점정보.MasterCode <> M_COUPON_KLENZ_CODE Then
        If Format(Date, "YYYY-MM-DD") >= "2009-12-11" Then
            Call SaveGroupGoodsINFO(고객정보, m_GSGMoney)
        End If
        
        ' 응모권 번호를 부여한다.
        If Format(Date, "YYYY-MM-DD") >= "2009-12-11" And Format(Date, "YYYY-MM-DD") <= "2009-12-31" Then
            If (m_GSGMoney.d2세트수량 + m_GSGMoney.d3세트수량 + m_GSGMoney.d4세트수량) Then
                Call GetGroupGoodsKeyNumber(고객정보)
            End If
            If (m_GSGMoney.d5세트수량 + m_GSGMoney.d6세트수량) Then
                Call GetGroupGoodsKeyNumber(고객정보)
            End If
        End If
    End If
    
    '------------------------------------------------------
    'Call WriteCardMoney(strDate) ' 카드 결재 금액이 있을 경우 카드 금액을 기록한다.
    
    '--------------------------------------------------------------------------
    ' 마일리지 정보 저장  ( 반품이 아닐 경우만 저장 : 2005-11-30일 변경)
    ' 마일리지 사용을 하던 체인점이 사용 중지를 했을 경우 기존에 있던 사용자는
    ' 잔액을 사용하게 하기 위하여 이쪽에서 확인 하지 않는다.
    '--------------------------------------------------------------------------
    'If 대리점정보.마일리지여부 = "Y" Then
        Call UseMileageUpdate(strCusNo, strDate)
   ' End If
    
    ' 미수일경우만 처리한다.
    If Index = 1 Then
        Call Fnc_MiSuEdit(strCusNo, CDbl(txtMisu.Text), "ADD")
    End If
    
    Call Receipt_Insert(PayType)  ' 보관증 저장
    
    ' 보관증출력
    DoEvents
    If SSOption1.Value = True Then
        If IsNumeric(GetIniStr("Printer", "Count", "", iniFile)) = False Then
            Call SetIniStr("Printer", "Count", "1", iniFile)
        End If

        For PrintCount = 1 To Val(GetIniStr("Printer", "Count", "", iniFile))
            If Printer_Gb = "0" Then
                Call subBillPrint
            ElseIf Printer_Gb = "1" Or Printer_Gb = "2" Then
                Call subinkPrintMM(CommonDialog1, sSEQ, strTelNo)
            End If
        Next PrintCount
    End If
    
    Call frm접수.접수_Clear
        
    If chkItem(strDate, strCusNo) = True Then
        Call 출고_DataDisplay(strCusNo)
        DoEvents
        
        chkinputflig = "입고완료"
    
        bBtnDuClick = False
        
        Unload frm결제
        
        frm출고.SetFocus
    Else
        frm접수.SetFocus
    End If
    
    Exit Sub
    
ErrRtn:
    bBtnDuClick = False
    
    MsgBox Err.Description, vbInformation, "확인"
    
    'Call ERR_SAVE("cmdAction_Click ERR_INDEX(" & CStr(ERR_INDEX) & ") " & Err.Description)
    
    Resume
End Sub

Private Sub cmdCoupon_Click()
    txtCouponNo.Enabled = Not txtCouponNo.Enabled
    
    If txtCouponNo.Enabled = True Then
        txtCouponNo.SetFocus
    End If
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
       
    If 대리점정보.삼성카드할인여부 <> "Y" Then Exit Sub
    
    sumMoney = 0
    strFirst = "삼"
    iPercentage = (100 - 대리점정보.삼성카드할인비율) / 100  ' (할인이 20%일 경우 0.8의 값을 같는다.)
    
    For nRow = 1 To frm접수.sprGrid.MaxRows
    
        If GetSpreadText(frm접수.sprGrid, nRow, 2) = "" Then Exit For
        
        frm접수.sprGrid.Row = nRow
        frm접수.sprGrid.Col = 4
        sTemp = Trim(frm접수.sprGrid.Text)
        
        If InStr(1, sTemp, strFirst, vbTextCompare) <= 0 Then
            ' 내용에 "삼"자가 없을 경우 "삼"을 추가하여 출력 한다.
            frm접수.sprGrid.Text = Mid(sTemp, 1, 1) & strFirst & Mid(sTemp, 2, Len(sTemp))
            
            ' 해당 금액을 얻어온다.
            logPrice = CLng(GetSpreadText(frm접수.sprGrid, nRow, 5))
                
            frm접수.sprGrid.Row = nRow
            frm접수.sprGrid.Col = 5
            ' 10원단위까지 수령한다.
            frm접수.sprGrid.Text = CStr(logPrice * iPercentage)
'            frm접수.sprGrid.Text = CStr(Int(CDbl((logPrice * iPercentage) / 100)) * 100) 10원단위 절사
            
            ' 누적 금액을 다시 계산한다.
            sumMoney = sumMoney + CLng(GetSpreadText(frm접수.sprGrid, nRow, 5))
    
            frm접수.lblSamSungCardCheck.Tag = "Y"
        
        Else
            ' 내용에 "삼"을 제거한다
            frm접수.sprGrid.Text = Replace(frm접수.sprGrid.Text, "삼", "")
            
            ' 해당 금액을 얻어온다.
            logPrice = CLng(GetSpreadText(frm접수.sprGrid, nRow, 5))
                
            frm접수.sprGrid.Row = nRow
            frm접수.sprGrid.Col = 5
            ' 10원단위까지 수령한다.
            frm접수.sprGrid.Text = CStr(logPrice / iPercentage)
'            frm접수.sprGrid.Text = CStr(Int(CDbl((logPrice * iPercentage) / 100)) * 100) 10원단위 절사
            
            ' 누적 금액을 다시 계산한다.
            sumMoney = sumMoney + CLng(GetSpreadText(frm접수.sprGrid, nRow, 5))
            
            frm접수.lblSamSungCardCheck.Tag = "N"
        
        End If
        
    Next nRow
            
    If sumMoney > 0 Then txtSum.Text = Format(sumMoney, "###,##0 ")
    
    cmdSamSungCard.Caption = IIf(frm접수.lblSamSungCardCheck.Tag = "Y", "삼성카드 할인 취소", "삼성카드 할인(10%)")
            
    DoEvents

End Sub

Private Sub Form_Activate()
    Dim myDate As Date
    Dim myDay  As Integer
    
    Printer_Gb = CStr(GetPrtGubun)
    Printer_BO_Gb = CStr(GetPrtBOGubun)

    Select Case True
        Case frm접수.Option1.Value:   myDay = 3
        Case frm접수.Option2.Value:   myDay = 4
        Case frm접수.Option3.Value:   myDay = 5
        Case frm접수.SSOption1.Value: myDay = Left(frm접수.cboWorkDay.Text, Len(frm접수.cboWorkDay.Text) - 2)
    End Select
    
    Call Fun_UserMileage(frm접수.txtCode.Text & "") ' 마일리지 금액을 표시한다.
    
    If userMileage.잔액 > txtSum.Value Then
        txtMileage.Value = txtSum.Value
    Else
        txtMileage.Value = userMileage.잔액 & ""
    End If
    
    '삼성카드 할인 여부
    cmdSamSungCard.Enabled = IIf(대리점정보.삼성카드할인여부 = "Y", True, False)
    cmdSamSungCard.Caption = IIf(frm접수.lblSamSungCardCheck.Tag = "Y", "삼성카드 할인 취소", "삼성카드 할인(10%)")
    
    ' 후불 처리 문제때문에 기록하지 말것
    'txtMoney.Text = Format(CLng(txtSum.text) - CLng(txtMileage.Text), "#,##0")
    myDate = Format(Date + myDay, "YYYY-MM-DD")
    
    dtpMonth.Value = Format(myDate, "MM-DD")
    
    txtMoney.Value = 0
    txtMoney.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    'frm결제.Top = 2000
    'frm결제.Left = 3000
    
    If 대리점정보.StoreCode = "999999" And Format(Date, "YYYY-MM-DD") <= "20091211" Then
        frm결제.Top = 2000
        frm결제.Left = 8000
    End If
    
    TempBan = False
End Sub

Private Sub txtMileage_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strMoney As Long
    
    If KeyCode = vbKeyReturn Then
        strMoney = CLng(txtSum.Text) - CLng(txtMoney.Text) - CLng(txtMileage.Text) - CLng(txtCoupon.Text)
        txtMisu.Text = Format(strMoney, "###,##0")
        cmdAction(1).SetFocus
    End If
    
End Sub

Private Sub txtMoney_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strMoney As Long
    
    Select Case KeyCode
        Case vbKeyReturn
            If Len(txtSum.Text) <= 0 Then txtSum.Text = "0"
            If Len(txtMoney.Text) <= 0 Then txtMoney.Text = "0"
            If Len(txtMileage.Text) <= 0 Then txtMileage.Text = "0"
            If Len(txtCoupon.Text) <= 0 Then txtCoupon.Text = "0"
            
            strMoney = CLng(txtSum.Text) - CLng(txtMoney.Text) - CLng(txtMileage.Text) - CLng(txtCoupon.Text)
            
            If strMoney < 0 Then
                MsgBox "입금액을 확인하여 주십시요.     ", vbCritical, "확인"
                txtMoney.SelStart = 0:  txtMoney.SelLength = Len(txtMoney.Text)
                Exit Sub
            End If
            txtMisu.Text = Format(strMoney, "###,##0")
            cmdAction(1).SetFocus
    
    End Select
    
End Sub

'Private Sub SSCommand1_Click()
'    On Error GoTo 0
'
'    Dim iCnt As Integer
'    Dim strCusNo As String
'    Dim strName As String       ' 품명
'    Dim strColor As String      ' 색상
'    Dim strTagNo As String      ' 택번호
'    Dim strContent As String    ' 세탁내용
'    Dim douMoney As Double      ' 가격
'    Dim strMark As String       ' 상표
'    Dim StrDate As String
'    Dim strState As String      ' 결제여부
'    Dim strRDate As String
'
'    Dim QueryDelete As String
'    Dim msg As String
'    Dim strinDate As String
'    Dim Query3 As String
'    Dim strCode As String
'    Dim Query4 As String
'    Dim rs04 As Recordset
'    Dim Query2 As String       ' 반품
'    Dim RS2 As Recordset        ' 반품환불확인
'    Dim Query3 As String       ' 반품환불
'    Dim sCupon As String        ' 쿠폰번호
'
'
'    Tag_Update                      ' 대리점정보에 기록
'
'
'
'
'    strState = "完"               ' 완불인경우 수령액 =합계금액 ,잔액=0
'    strCusNo = frm접수.txtCode.text
''    strDate = Mid(Date, 1, 4) & Mid(Date, 6, 2) & Mid(Date, 9, 2)
'    If DayCloseCheck(Format(Date, "YYYY-MM-DD")) = True Then
'        MsgBox "일마감이 되었으므로 판매내역은 익일로 저장이 됩니다.", vbInformation
'        StrDate = Format(DateAdd("d", 1, Date), "YYYY-MM-DD")
'    Else
'        StrDate = Format(Date, "YYYY-MM-DD")
'    End If
'
'    iCnt = 1
'    douMoney = 0
'    frm접수.sprGrid.Row = iCnt
'    frm접수.sprGrid.Col = 1
'    strName = frm접수.sprGrid.Value
'
'    If frm접수.Option1.Value = True Then
'        strRDate = Format(DateAdd("d", 3, Format(Format(StrDate, "####-##-##"), "YYYY-MM-DD")), "YYYY-MM-DD")
'    ElseIf frm접수.Option2.Value = True Then
'        strRDate = Format(DateAdd("d", 4, Format(Format(StrDate, "####-##-##"), "YYYY-MM-DD")), "YYYY-MM-DD")
'    ElseIf frm접수.Option3.Value = True Then
'        strRDate = Format(DateAdd("d", 5, Format(Format(StrDate, "####-##-##"), "YYYY-MM-DD")), "YYYY-MM-DD")
'    ElseIf frm접수.SSOption1.Value = True Then
'        strRDate = Format(DateAdd("d", Trim(Mid(frm접수.cboWorkDay.Text, 1, 2)), Format(Format(StrDate, "####-##-##"), "YYYY-MM-DD")), "YYYY-MM-DD")
'    End If
'
'
'    DoEvents
'
'    If Val(txtMoney.Text) = 0 Then
'        txtmisu.text = 0
'        txtMoney.Text = CLng(txtSum.text) - CLng(Val(txtMileage.Text))
'    End If
'
'    ' 이용실적
'    UseAccountUpdate
'
'    ' 마일리지 정보 저장  ( 반품이 아닐 경우만 저장 : 20051130일 변경)
'    If 대리점정보.마일리지여부 = "Y" And InStr(strContent, "반") = 0 Then
'        Call UseMileageUpdate(strCusNo)
'    End If
'
'    ' 보관증 저장
'    Receipt_Insert (True)
'
'
'    DoEvents
'    If SSOption1.Value = True Then  ' 보관증출력
'        If Printer_Gb = "0" Then
'            subBillPrint
'        ElseIf Printer_Gb = "1" Or Printer_Gb = "2" Then
'            Call subinkPrintMM(CommonDialog1, sSEQ, _
'                                frm접수.txtTEL(0).Text & "-" & frm접수.txtTEL(1).Text)
'        End If
'    End If
'
'    If chkItem = True Then
'        frm접수.Hide
'        Load frm출고
'        frm출고.setfocus
'
'        'TitleSet "출고중 ..."
'        frmMain.Command1(0).BackColor = "&H00EBBF76"
'        frmMain.Command1(2).BackColor = "&H00C0C0C0"
'        frmMain.Command1(1).BackColor = "&H00C0C0C0"
'
'        strCode = frm접수.txtCode.Text
'        출고_DataDisplay strCode
'
'        chkinputflig = "입고완료"
'        Unload frm접수
'
'    End If
'
'    Unload Me
'End Sub
'
'Private Sub SSCommand2_Click()
'    Dim strCusNo As String
'    Dim Query2 As String
'    Dim Query3 As String
'    Dim strMisu As String
'    Dim QueryDelete As String      ' 보관증 태이블지움
'    Dim Query4 As String
'    Dim rs04 As Recordset
'    Dim strCode As String
'
'    Tag_Update                ' 대리점정보에 기록
'
'    strCusNo = frm접수.txtCode.text
'    Label3 = CLng(txtSum.text) - CLng(Val(txtMoney.text)) - CLng(Val(txtMileage.Text))
'    Label3 = Format(Label3, "#,#0")
'
'    TempBan = False
'
'    Receive_Update
'
'    ' 미수금 내용을 적용한다.
'    Call Fnc_MiSuEdit(Trim(frm접수.txtCode.text), CDbl(txtmisu.text), "ADD")
'
'    ' 보관증 저장(후불)
'    Receipt_Insert (False)
'
'    If SSOption1.Value = True Then  ' 보관증출력
'        If Printer_Gb = "0" Then
'            subBillPrint
'        Else
'            Call subinkPrintMM(CommonDialog1, sSEQ, _
'                                frm접수.txtTEL(0).Text & "-" & frm접수.txtTEL(1).Text)
'        End If
'    End If
'
'    UseAccountUpdate                 ' 이용실적
'
'    ' 마일리지 정보 저장  ( 반품이 아닐 경우만 저장 : 20051130일 변경)
'    If 대리점정보.마일리지여부 = "Y" And TempBan = False Then
'        Call UseMileageUpdate(strCusNo)
'    End If
'    TempBan = False
'
'    If chkItem = True Then
'        frm접수.Hide
'        Load frm출고
'        frm출고.setfocus
'
'        'TitleSet "출고중 .."
'
'        frmMain.Command1(0).BackColor = "&H0080FF80"
'        frmMain.Command1(2).BackColor = "&H00C0C0C0"
'        frmMain.Command1(1).BackColor = "&H00C0C0C0"
'
'        strCode = frm접수.txtCode.Text
'        출고_DataDisplay strCode
'
'        Call Fun_고객정보(strCode)
'        frm출고.Label1 = Format(고객정보.미수금, "###,##0")
'
'        chkinputflig = "입고완료"
'        Unload frm접수
'    End If
'
'    Unload Me
'End Sub
'
'
''--------------------------------------------------------------------------------------------------------------
'' Procedure : WriteCardMoney
'' DateTime  : 2007-05-06 05:15
'' Author    : pds2004
'' Purpose   : 카드 금액을 저장한다.
''--------------------------------------------------------------------------------------------------------------
'Private Function WriteCardMoney(ByVal sSaleDate As String)
'    Dim Query    As String
'    Dim dCard As Double
'
'
'    On Error GoTo ErrRtn
'
'    dCard = Val(Replace(mskCard.ClipText, ",", ""))
'    If dCard <= 0 Then Exit Function
'
'    Query = "INSERT INTO TB_카드금액(결재일자, 접수시간, 고객번호, 금액)"
'    Query = Query & "VALUES ('" & sSaleDate & "', "
'    Query = Query & "'" & Format(Time, "hhmmss") & "', "
'    Query = Query & "'" & frm접수.txtCode.text & "', "
'    Query = Query & "'" & dCard & "') "
'    ADOCon.Execute Query
'
'    On Error GoTo 0
'    Exit Function
'
'ErrRtn:
'    'sendErrormessage
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteCardMoney of Form frm결제"
'End Function
 
' 사용한 쿠폰의 자료를 저장한다.
Private Sub SaveCouponDate(ByVal strDate As String)
    Dim varTemp As Variant
    Dim nIndex  As Integer
    Dim CouponMoney As Double
    Dim strSumMoney As Double
    Dim sCouponNum  As String
    
    On Error GoTo ErrRtn
    
    CouponMoney = 0
    strSumMoney = txtSum.Value
    
    varTemp = Split(txtCouponNo.Text, vbNewLine)
    
    For nIndex = 0 To UBound(varTemp)
        sCouponNum = CStr(varTemp(nIndex))
        
        If sCouponNum <> "" Then
            Select Case CheckCouponNumber(sCouponNum)
                Case -1
                    MsgBox "쿠폰 번호 오류 [" & sCouponNum & "]     ", vbInformation, "확인"
                    
                Case -2
                    MsgBox "쿠폰 사용 기간 오류 [" & sCouponNum & "]     ", vbInformation, "확인"
                
                Case Else
                    Query = " SELECT * "
                    Query = Query & " FROM TB_쿠폰자료 "
                    Query = Query & " WHERE 접수일자 = '" & Trim(frm접수.txtCode.Text) & "' "
                    Query = Query & " AND 쿠폰번호 = '" & Format(Date, "yyyy") & "'"
                    Set SUBRs = New ADODB.Recordset
                    SUBRs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
                    
                    If SUBRs.RecordCount <= 0 Then
                        Query = "INSERT INTO TB_쿠폰자료(접수일자, 대리점코드, 쿠폰번호, 택번호, "
                        Query = Query & "쿠폰단가, 쿠폰금액, "
                        Query = Query & "고객번호, 고객이름, "
                        Query = Query & "접수금액, 전송여부, 전송일자) "
                        
                        Query = Query & "VALUES ('" & strDate & "', '" & 대리점정보.StoreCode & "','" & sCouponNum & "', '" & 대리점정보.대리점번호 & "', "
                        Query = Query & " " & Get_CouponCost(sCouponNum) & ", " & Get_CouponMoney(sCouponNum) & ", "
                        Query = Query & "'" & Trim(frm접수.txtCode.Text) & "', '" & Trim(frm접수.txtName.Text) & "', "
                        Query = Query & " " & strSumMoney & " , 'N',' ' )"
                        ADOCon.Execute Query
                    End If
                    SUBRs.Close
                    Set SUBRs = Nothing
            End Select
        End If
    Next nIndex
    
    Exit Sub
    
ErrRtn:
    MsgBox Err.Description, vbInformation, "확인"
    'Call ERR_SAVE("SaveCouponDate" & Err.Description & Query)
    Resume Next
End Sub

Private Sub txtMoney_KeyPress(KeyAscii As Integer)
    Dim strMoney As Long
    
    Select Case KeyAscii
        Case vbKey0 To vbKey9
        Case vbKeyNumpad0 To vbKeyNumpad9
        Case vbKeyBack
        Case vbKeyReturn
        Case Else
            KeyAscii = 0
    
    End Select

End Sub

Private Sub txtCouponNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        
        Case vbKeyReturn, vbKeyBack, vbKeyHome, vbKeyEnd
        
        
        Case Else
            KeyAscii = 0
            
    End Select
End Sub

Private Sub txtCouponNo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim varTemp As Variant
        Dim nIndex  As Integer
        Dim CouponMoney As Double
        Dim sCouponNum  As String
        Dim dblTempMoney    As Double
        
RE_START:
        CouponMoney = 0
        varTemp = Split(txtCouponNo.Text, vbNewLine)
        For nIndex = 0 To UBound(varTemp)
            sCouponNum = Trim(varTemp(nIndex))
            If sCouponNum <> "" Then
            
                '4자리 입력 내용 변환
                If Len(sCouponNum) = 6 And Left(sCouponNum, 2) = "01" Then
                    txtCouponNo.Text = Replace(txtCouponNo.Text, sCouponNum, Left(sCouponNum, 2) & "00" & Right(sCouponNum, 4))
                    txtCouponNo.SelStart = Len(txtCouponNo)
                    GoSub RE_START
                End If
            
                
                Select Case CheckCouponNumber(sCouponNum)
                    Case -1
                        MsgBox "쿠폰 번호 오류 [" & sCouponNum & "]     ", vbInformation, "확인"
                        Exit Sub
                                    
                    Case -2
                        MsgBox "쿠폰 사용 기간 오류 [" & sCouponNum & "]     ", vbInformation, "확인"
                        Exit Sub
                    
                End Select
                    
                ' 쿠폰 금액을 누적 처리한다.
                CouponMoney = CouponMoney + Get_CouponMoney(sCouponNum)
                txtCoupon.Text = CStr(CouponMoney)
                
                ' 마일리지 잔액이 있을 경우
                If userMileage.잔액 > 0 Then
                    If userMileage.잔액 > CouponMoney Then
                        ' 마일리지 잔액이 전체 금액보다 적을 경우 마일리지 금액만
                        If CLng(txtSum.Text) > (userMileage.잔액 - CouponMoney) Then
                            txtMileage.Text = userMileage.잔액 - CouponMoney
                            
                        ' 마일리지 금액이 전체 금액 보다 클 경우 전체 금액만 처리한다.
                        Else
                            txtMileage.Text = CLng(txtSum.Text) - CouponMoney
                        
                        End If
                        
                    Else
                        If CLng(txtSum.Text) - CouponMoney Then
                            If userMileage.잔액 < CLng(txtSum.Text) - CouponMoney Then
                                txtMileage.Text = userMileage.잔액
                            
                            Else
                                txtMileage.Text = "0" 'CLng(txtSum.text) - CouponMoney
                            End If
                        Else
                            txtMileage.Text = "0"
                        End If
                    End If
                End If
                
                dblTempMoney = CLng(txtSum.Text) - CLng(txtMoney.Text) - CLng(txtMileage.Text) - CLng(txtCoupon.Text)
                If dblTempMoney <= 0 Then
                    txtMisu.Text = "0"
                Else
                    txtMisu.Text = Format(dblTempMoney, "###,##0")
                End If
            End If
        Next nIndex
    End If
End Sub

Private Sub txtCouponNo_LostFocus()
    Dim varTemp       As Variant
    Dim nIndex        As Integer
    Dim CouponMoney   As Double
    Dim sCouponNum    As String
    Dim dblTempMoney  As Double

RE_START:
    txtCoupon.Text = "0"
    CouponMoney = 0
    varTemp = Split(txtCouponNo.Text, vbNewLine)
    
    For nIndex = 0 To UBound(varTemp)
        sCouponNum = Trim(varTemp(nIndex))
        
        If sCouponNum <> "" Then
            '4자리 입력 내용 변환
            If Len(sCouponNum) = 6 And Left(sCouponNum, 2) = "01" Then
                txtCouponNo.Text = Replace(txtCouponNo.Text, sCouponNum, Left(sCouponNum, 2) & "00" & Right(sCouponNum, 4))
                txtCouponNo.SelStart = Len(txtCouponNo)
                GoSub RE_START
            End If
            
            Select Case CheckCouponNumber(sCouponNum)
                Case -1
                    MsgBox "쿠폰 번호 오류 [" & sCouponNum & "]     ", vbInformation, "확인"
                    Exit Sub
                                
                Case -2
                    MsgBox "쿠폰 사용 기간 오류 [" & sCouponNum & "]     ", vbInformation, "확인"
                    Exit Sub
                
            End Select
                
            ' 쿠폰 금액을 누적 처리한다.
            CouponMoney = CouponMoney + Get_CouponMoney(sCouponNum)
            txtCoupon.Text = CStr(CouponMoney)
        End If
    Next nIndex

    If userMileage.잔액 > 0 Then
        If CouponMoney = 0 Then
            If userMileage.잔액 > Val(Replace(txtSum.Text, ",", "")) Then
                txtMileage.Text = txtSum.Text
            Else
                txtMileage.Text = Format(userMileage.잔액, "#,##0")
            End If
        Else
    
            If userMileage.잔액 > CouponMoney Then
                ' 마일리지 잔액이 전체 금액보다 적을 경우 마일리지 금액만
                If CLng(txtSum.Text) > (userMileage.잔액 - CouponMoney) Then
                    txtMileage.Text = userMileage.잔액 - CouponMoney
                    
                ' 마일리지 금액이 전체 금액 보다 클 경우 전체 금액만 처리한다.
                Else
                    txtMileage.Text = CLng(txtSum.Text) - CouponMoney
                
                End If
                
            Else
                If CLng(txtSum.Text) - CouponMoney Then
                    If userMileage.잔액 < CLng(txtSum.Text) - CouponMoney Then
                        txtMileage.Text = userMileage.잔액
                    
                    Else
                        txtMileage.Text = CLng(txtSum.Text) - CouponMoney
                    End If
                Else
                    txtMileage.Text = "0"
                End If
            End If
        End If
    End If
        
    dblTempMoney = CLng(txtSum.Text) - CLng(txtMoney.Text) - CLng(txtMileage.Text) - CLng(txtCoupon.Text)
    
    If dblTempMoney <= 0 Then
        txtMisu.Text = "0"
    Else
        txtMisu.Text = Format(dblTempMoney, "###,##0")
    End If
End Sub
