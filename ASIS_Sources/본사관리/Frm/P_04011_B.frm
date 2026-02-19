VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_04011_B 
   Caption         =   "가맹점 기간별 매출현황 (일자별)"
   ClientHeight    =   10050
   ClientLeft      =   8355
   ClientTop       =   450
   ClientWidth     =   16110
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_04011_B.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10050
   ScaleWidth      =   16110
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   10050
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   16110
      _ExtentX        =   28416
      _ExtentY        =   17727
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   0
      PaneTree        =   "P_04011_B.frx":058A
      Begin Threed.SSPanel SSPanel 
         Height          =   855
         Index           =   0
         Left            =   0
         TabIndex        =   17
         Top             =   9195
         Width           =   16110
         _ExtentX        =   28416
         _ExtentY        =   1508
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   14
            Left            =   60
            TabIndex        =   19
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
            Caption         =   "전체매출액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   15
            Left            =   2340
            TabIndex        =   20
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
            Caption         =   "지사매출액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   16
            Left            =   4620
            TabIndex        =   21
            Top             =   375
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
            Caption         =   "입고 수량"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   17
            Left            =   4620
            TabIndex        =   22
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
            Caption         =   "가맹점매출액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   18
            Left            =   9180
            TabIndex        =   23
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
            Caption         =   "카드 금액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   19
            Left            =   6900
            TabIndex        =   24
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
            Caption         =   "수선 금액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   20
            Left            =   9180
            TabIndex        =   25
            Top             =   375
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
            Caption         =   "카드 건수"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   21
            Left            =   6900
            TabIndex        =   26
            Top             =   375
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
            Caption         =   "수선 수량"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   22
            Left            =   11460
            TabIndex        =   27
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   582
            _Version        =   262144
            BackColor       =   12632319
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "반품 수량"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   23
            Left            =   11460
            TabIndex        =   28
            Top             =   375
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
            Caption         =   "재세탁수량"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   24
            Left            =   60
            TabIndex        =   29
            Top             =   375
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
            Caption         =   "전체 단가"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   25
            Left            =   2340
            TabIndex        =   30
            Top             =   375
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
            Caption         =   "지사 단가"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   330
            Index           =   0
            Left            =   1200
            TabIndex        =   31
            Top             =   60
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
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
            Index           =   10
            Left            =   1200
            TabIndex        =   32
            Top             =   375
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
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
            Left            =   3480
            TabIndex        =   33
            Top             =   60
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
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
            Index           =   11
            Left            =   3480
            TabIndex        =   34
            Top             =   375
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
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
            Index           =   2
            Left            =   5760
            TabIndex        =   35
            Top             =   60
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
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
            Index           =   7
            Left            =   8040
            TabIndex        =   36
            Top             =   60
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
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
            Index           =   4
            Left            =   10320
            TabIndex        =   37
            Top             =   60
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
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
            Index           =   3
            Left            =   5760
            TabIndex        =   38
            Top             =   375
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
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
            Index           =   6
            Left            =   8040
            TabIndex        =   39
            Top             =   375
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
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
            Index           =   5
            Left            =   10320
            TabIndex        =   40
            Top             =   375
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
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
            Index           =   8
            Left            =   12600
            TabIndex        =   41
            Top             =   60
            Width           =   930
            _Version        =   262145
            _ExtentX        =   1640
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
            Index           =   9
            Left            =   12600
            TabIndex        =   42
            Top             =   375
            Width           =   930
            _Version        =   262145
            _ExtentX        =   1640
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
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   8505
         _ExtentX        =   15002
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
         Caption         =   " #"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_04011_B.frx":063C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   1
         Left            =   8520
         TabIndex        =   3
         Top             =   0
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
         PictureBackground=   "P_04011_B.frx":083E
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   4
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
            Picture         =   "P_04011_B.frx":0A40
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   5
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
            Picture         =   "P_04011_B.frx":0FDA
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   6
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
            Picture         =   "P_04011_B.frx":1574
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   7
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
            Picture         =   "P_04011_B.frx":1B0E
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   8
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
            Picture         =   "P_04011_B.frx":20A8
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   9
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
            Picture         =   "P_04011_B.frx":2642
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   10
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
            Picture         =   "P_04011_B.frx":2BDC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   11
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
            Picture         =   "P_04011_B.frx":3176
         End
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   0
         TabIndex        =   12
         Top             =   525
         Width           =   16110
         _ExtentX        =   28416
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin XtremeSuiteControls.CheckBox chkColor 
            Height          =   255
            Left            =   9120
            TabIndex        =   47
            Top             =   90
            Width           =   1335
            _Version        =   851970
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "색상 적용"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.ComboBox cboInput 
            Height          =   315
            Left            =   1245
            Style           =   2  '드롭다운 목록
            TabIndex        =   0
            Top             =   420
            Width           =   3060
         End
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1245
            TabIndex        =   13
            Text            =   "cboOffice"
            Top             =   60
            Width           =   3060
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   4605
            TabIndex        =   14
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "매출일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   15
            Top             =   420
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "가 맹 점"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   35
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "사 업 장"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin XtremeSuiteControls.DateTimePicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   5790
            TabIndex        =   43
            Top             =   60
            Width           =   1470
            _Version        =   851970
            _ExtentX        =   2593
            _ExtentY        =   556
            _StockProps     =   68
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   3
         End
         Begin XtremeSuiteControls.DateTimePicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   7500
            TabIndex        =   44
            Top             =   60
            Width           =   1470
            _Version        =   851970
            _ExtentX        =   2593
            _ExtentY        =   556
            _StockProps     =   68
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   3
         End
         Begin XtremeSuiteControls.PushButton cmdRefresh 
            Height          =   330
            Left            =   4335
            TabIndex        =   46
            Top             =   405
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   582
            _StockProps     =   79
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_04011_B.frx":3710
         End
         Begin XtremeSuiteControls.CheckBox CheckBox1 
            Height          =   225
            Left            =   9120
            TabIndex        =   48
            Top             =   420
            Width           =   3135
            _Version        =   851970
            _ExtentX        =   5530
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "지사출고 수량 집계"
            UseVisualStyle  =   -1  'True
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
            Height          =   240
            Left            =   7215
            TabIndex        =   45
            Top             =   120
            Width           =   300
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   7860
         Left            =   0
         TabIndex        =   18
         Top             =   1320
         Width           =   16110
         _Version        =   524288
         _ExtentX        =   28416
         _ExtentY        =   13864
         _StockProps     =   64
         BackColorStyle  =   1
         ColsFrozen      =   2
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   31
         SpreadDesigner  =   "P_04011_B.frx":3CAA
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_04011_B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim oldColor    As Long

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cboInput_Click()
    Call Data_Display
End Sub

'Private Sub cboInput_Click(Index As Integer)
'    Dim sCode As String
'
'    If Index = 1 Then
'        sCode = Trim(Mid(Trim(cboInput(1)) & Space(10), 2, 4))
'
'        Call Get_가맹점리스트(cboOffice, sCode)
'    End If
'End Sub

Private Sub cboOffice_Click()
    If cboOffice.ListIndex < 0 Then Exit Sub
    
    '-----------------------------------------------------------------
    '
    '-----------------------------------------------------------------
    cboInput.Clear
    
    ReDim sValue(2)
    
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01001_00", sValue(), Err_Num, Err_Dec)
    
    cboInput.AddItem "[000000] 전체"
    
    Do Until RS01.EOF
        cboInput.AddItem "[" & RS01!가맹점코드 & "] " & RS01!가맹점명
        RS01.MoveNext
    Loop
    RS01.Close
    Set RS01 = Nothing
    
    If cboInput.ListCount > 0 Then cboInput.ListIndex = 0
End Sub

 

Private Sub cboOffice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SearchString_한글 KeyAscii
'    Else
       ' SearchString KeyAscii
    End If
End Sub

Private Sub chkColor_Click()
    
    Call chkColor_Setting(chkColor.Value)

End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: Call Data_Display   ' 조회
        Case 1: 'Call DataAdd        ' 신규
        Case 2: Call DataSave       ' 저장
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

Private Sub dtInput_Change(Index As Integer)
'    Call Data_Display
End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(2).Enabled = True
    cmdBtn(5).Enabled = True
    cmdBtn(6).Enabled = True
    
    '본사 여부
    If HeadOffice = MASTER_OFFICE_CODE Then
        cboOffice.Locked = False
    Else
        cboOffice.Locked = True
    End If
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
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
'        .OperationMode = OperationModeSingle
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
    
    dtInput(0).Value = Format(Date, "YYYY-MM-DD")
    dtInput(1).Value = Format(Date, "YYYY-MM-DD")
    
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
    
     Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_04011_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i As Integer
    Dim nRow    As Long
    Dim vText   As Variant
    Dim sSearch1 As String
    Dim sSearch2 As String
    
    'For i = 0 To 11
    '    txtNum(i).Value = 0
    'Next i
    
    ReDim sValue(3)
    
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    
    If Mid(cboInput.Text, 2, 6) = "000000" Then
        sValue(1) = ""
    Else
        sValue(1) = Mid(cboInput.Text, 2, 6)
    End If
    
    sValue(2) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(3) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(HeadOffice) = False Then Exit Sub
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_04001_00", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_04001_00", sValue(), Err_Num, Err_Dec)
    End If
    
    With spdView
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = RS01!마감일자 & ""              ' 1
            .Col = 2:  .Text = ExecWeekDay(RS01!마감일자) & "" ' 2
            .Col = 3:  .Text = RS01!가맹점코드 & ""            ' 3
            .Col = 4:  .Text = RS01!가맹점명 & ""              ' 4
            .Col = 5:  .Text = RS01!지사금액 & ""              ' 5
            .Col = 6:  .Text = RS01!가맹점금액 & ""            ' 6
            .Col = 7:  .Text = RS01!접수수량 & ""              ' 7
            .Col = 8:  .Text = "0" 'RS01!출고수량 & ""              ' 8
                        
            If Len(RS01!시작택번호) = 9 Then
                .Col = 9:  .Text = Format(RS01!시작택번호, "000-00-0000") & ""         ' 9
            Else
                .Col = 9:  .Text = RS01!시작택번호 & ""        ' 9
            End If
            
            If Len(RS01!종료택번호) = 9 Then
                .Col = 10: .Text = Format(RS01!종료택번호, "000-00-0000") & ""         '10
            Else
                .Col = 10: .Text = RS01!종료택번호 & ""        '10
            End If
            
            .Col = 11
            Select Case RS01!판매구분
                Case "1": .Text = "세일"       '11
                Case "2": .Text = "요일"       '
                Case "3": .Text = "정상"      '
            End Select

            .Col = 12: .Text = RS01!접수금액 & ""                   '12
            .Col = 13: .Text = RS01!현금입금 + RS01!카드금액 & ""   '13
            
            If RS01!접수수량 = 0 Then
                .Col = 14: .Text = 0 & ""   '14
                .Col = 15: .Text = 0 & ""   '15
                .Col = 16: .Text = 0 & ""   '16
            Else
                .Col = 14: .Text = RS01!접수금액 / RS01!접수수량 & ""   '14
                .Col = 15: .Text = RS01!지사금액 / RS01!접수수량 & ""   '15
                .Col = 16: .Text = RS01!가맹점금액 / RS01!접수수량 & "" '16
            End If
            
            .Col = 17: .Text = RS01!현금입금 & ""                   '17
            .Col = 18: .Text = RS01!카드금액 & ""                   '18
            .Col = 19: .Text = RS01!카드건수 & ""                   '19
            .Col = 20: .Text = RS01!쿠폰금액 & ""                   '20
            .Col = 21: .Text = RS01!쿠폰건수 & ""                   '21
            .Col = 22: .Text = RS01!발생마일리지 & ""               '22
            .Col = 23: .Text = RS01!사용마일리지 & ""               '23
            .Col = 24: .Text = RS01!삭제마일리지 & ""               '24
            .Col = 25: .Text = RS01!반품환불금액 & ""               '25
            .Col = 26: .Text = RS01!반품환불건수 & ""               '26
            .Col = 27: .Text = RS01!세탁환불금액 & ""               '27
            .Col = 28: .Text = RS01!세탁환불건수 & ""               '28
            .Col = 29: .Text = RS01!재세탁수량 & ""                 '29
            .Col = 30: .Text = RS01!수선금액 & ""                   '30
            .Col = 31: .Text = RS01!수선수량 & ""                   '31
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        ReDim sValue(4)
        
        sValue(0) = Mid(cboOffice.Text, 2, 4)
        If sValue(0) = "0000" Then sValue(0) = "%"

        sValue(1) = Mid(cboInput.Text, 2, 6)
        If sValue(1) = "000000" Then sValue(1) = "%"
        
        
        sValue(2) = Format(dtInput(0).Value, "YYYY-MM-DD")
        sValue(3) = Format(dtInput(1).Value, "YYYY-MM-DD")
        sValue(4) = "STORE_DAY"
        
        If CheckBox1.Value = xtpChecked Then
            If HeadOffice = MASTER_OFFICE_CODE Then
                If DBOpen_Master(HeadOffice) = False Then Exit Sub
                
                Set RS01 = New ADODB.Recordset
                Set RS01 = ExecProMaster("SP_04001_B_01", sValue(), Err_Num, Err_Dec)
            Else
                Set RS01 = New ADODB.Recordset
                Set RS01 = ExecPro("SP_04001_B_01", sValue(), Err_Num, Err_Dec)
            End If
            
            
            Do While Not RS01.EOF
                ' 지사출고 수량 출력
                sSearch2 = RS01.Fields(1) + " " + RS01.Fields(2)
                
                For nRow = 1 To .MaxRows
                
                    ' 가맹점 코드 + ' ' + 일자
                    .GetText 3, nRow, vText:    sSearch1 = CStr(vText)
                    .GetText 1, nRow, vText:    sSearch1 = sSearch1 + " " + CStr(vText)
                    
                    
                    If sSearch1 = sSearch2 Then
                        .SetText 8, nRow, CVar(RS01.Fields(3))
                        Exit For
                    End If
                Next nRow
                
                RS01.MoveNext
            Loop
            RS01.Close
            Set RS01 = Nothing
        End If
        
        If .MaxRows > 0 Then

            ' 합계 출력
            Dim nCol    As Long
            Dim dblCnt(4)   As Double
            For nCol = 5 To .MaxCols
                Select Case nCol
                    Case 5: dblCnt(2) = SpreadSum(spdView, 2, nCol)
                    Case 6: dblCnt(3) = SpreadSum(spdView, -1, nCol)
                    Case 12: dblCnt(1) = SpreadSum(spdView, -1, nCol)
                    Case 7:  dblCnt(0) = SpreadSum(spdView, -1, nCol)
                    Case 14: .SetText nCol, .MaxRows, CVar(dblCnt(1) / dblCnt(0))
                    Case 15: .SetText nCol, .MaxRows, CVar(dblCnt(2) / dblCnt(0))
                    Case 16: .SetText nCol, .MaxRows, CVar(dblCnt(3) / dblCnt(0))
                    Case Else: Call SpreadSum(spdView, -1, nCol)
                End Select
            Next nCol
    
'            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Row = .Row
            .Row2 = .Row
            .Col = 1
            .Col2 = .MaxCols
            .BlockMode = True
            .BackColor = &HC0FFC0
            .BlockMode = False
'
'            .Col = 4:  .Text = "합계"
'
'            .Col = 5:  .Formula = "SUM(E1:E" & .MaxRows - 1 & ")"
'            .Col = 6:  .Formula = "SUM(F1:F" & .MaxRows - 1 & ")"
'            .Col = 7:  .Formula = "SUM(G1:G" & .MaxRows - 1 & ")"
'            .Col = 8:  .Formula = "SUM(H1:H" & .MaxRows - 1 & ")"
'
'            .Col = 12: .Formula = "SUM(L1:L" & .MaxRows - 1 & ")"
'            .Col = 13: .Formula = "SUM(M1:M" & .MaxRows - 1 & ")"
'
'            .Col = 14: .Formula = "SUM(N1:N" & .MaxRows - 1 & ") / " & .MaxRows - 1 & ""
'            .Col = 15: .Formula = "SUM(O1:O" & .MaxRows - 1 & ") / " & .MaxRows - 1 & ""
'            .Col = 16: .Formula = "SUM(P1:P" & .MaxRows - 1 & ") / " & .MaxRows - 1 & ""
'
'            .Col = 17: .Formula = "SUM(Q1:Q" & .MaxRows - 1 & ")"
'            .Col = 18: .Formula = "SUM(R1:R" & .MaxRows - 1 & ")"
'            .Col = 19: .Formula = "SUM(S1:S" & .MaxRows - 1 & ")"
'            .Col = 20: .Formula = "SUM(T1:T" & .MaxRows - 1 & ")"
'            .Col = 21: .Formula = "SUM(U1:U" & .MaxRows - 1 & ")"
'            .Col = 22: .Formula = "SUM(V1:V" & .MaxRows - 1 & ")"
'            .Col = 23: .Formula = "SUM(W1:W" & .MaxRows - 1 & ")"
'
'            .Col = 24: .Formula = "SUM(X1:X" & .MaxRows - 1 & ")"
'            .Col = 25: .Formula = "SUM(Y1:Y" & .MaxRows - 1 & ")"
'            .Col = 26: .Formula = "SUM(Z1:Z" & .MaxRows - 1 & ")"
'            .Col = 27: .Formula = "SUM(AA1:AA" & .MaxRows - 1 & ")"
'            .Col = 28: .Formula = "SUM(AB1:AB" & .MaxRows - 1 & ")"
'            .Col = 29: .Formula = "SUM(AC1:AC" & .MaxRows - 1 & ")"
'            .Col = 30: .Formula = "SUM(AD1:AD" & .MaxRows - 1 & ")"
'            .Col = 31: .Formula = "SUM(AE1:AE" & .MaxRows - 1 & ")"
            
            '------------------------------------------------------
            
            .Col = 12:  txtNum(0).Value = .Value  '전체매출액
            .Col = 14: txtNum(10).Value = .Value '전체단가
            .Col = 15: txtNum(11).Value = .Value '지사단가

            .Col = 5: txtNum(1).Value = .Value   '지사매출
            .Col = 6: txtNum(2).Value = .Value   '가맹점매출

            .Col = 7: txtNum(3).Value = .Value   '입고수량

            .Col = 30: txtNum(7).Value = .Value   '수선금액
            .Col = 31: txtNum(6).Value = .Value   '수선수량

            .Col = 18: txtNum(4).Value = .Value   '카드금액
            .Col = 19: txtNum(5).Value = .Value   '카드수량

            .Col = 25: txtNum(8).Value = .Value   '반품수량
            .Col = 29: txtNum(9).Value = .Value   '재세탁수량
        End If
        
        ' 색상을 적용 한다.
        Call chkColor_Setting(chkColor.Value)
        
        .Row = 1:   .Col = -1:  oldColor = .BackColor
            
        ' 누락된 요일을 설정한다.
        Call DateCheckAdd(spdView, Format(dtInput(0).Value, "YYYY-MM-DD"), Format(dtInput(1).Value, "YYYY-MM-DD"))
        .Redraw = True
    End With
        
    
    
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

'private Sub Data_Display()
'    Dim sCode   As String
'    Dim i
'    Dim ii As Integer
'
'    sCode = Mid(Trim(cboInput(1).Text) & Space(10), 2, 4)
'
'    If Trim(cboInput(1).Text) = "" Then
'        MsgBox "사업장을 선택하십시오.", vbInformation
'        cboInput(1).SetFocus
'        Exit Sub
'    End If
'
'    If Trim(cboOffice.Text) = "" Then
'        MsgBox "가맹점을 선택하십시오.", vbInformation
'        cboOffice.SetFocus
'        Exit Sub
'    End If
'
'    Set RS01 = New ADODB.Recordset
'
'
'    ReDim sValue(2)
'    sValue(0) = "0"
'    sValue(1) = Format(dtInput.Value, "yyyymm")
'    sValue(2) = Mid(cboOffice.Text, 2, 6)
'    Set RS01 = ExecPro("SP_04011_00_ALL", sValue(), Err_Num, Err_Dec)
'
'
'
'    spdView.MaxCols = RS01.Fields.Count
'    spdView.MaxRows = RS01.RecordCount
'
'    Call spdDisplay(RS01)
'    Call GetColWidth(REG_App, Me.Name, spdView)
'
'    spdView.MaxRows = 0
'
'    For i = 1 To DatePart("d", DateAdd("d", -1, DateAdd("m", 1, dtInput.Value)))
'        spdView.MaxRows = spdView.MaxRows + 1
'        spdView.Row = spdView.MaxRows
'
'        spdView.Col = 1
'        spdView.Text = Format(dtInput.Value, "yyyy-mm" & "-" & Right("0" & i, 2))
'        spdView.Col = 2
'        spdView.Text = ExecWeekDay(Format(dtInput.Value, "yyyy-mm" & "-" & Right("0" & i, 2)))
'
'        If spdView.Text = "일" Then
'            'spdView.Row = NewRow
'            spdView.Col = -1
'            spdView.BackColor = vbYellow
'        End If
'
'        spdView.Col = 7
'        Select Case spdView.Text
'            Case "1"
'                spdView.Text = spdView.Text + ":세일"
'            Case "2"
'                spdView.Text = spdView.Text + ":목요"
'            Case "3"
'                spdView.Text = spdView.Text + ":정상"
'        End Select
'        Dim j As Integer
'
'        For j = 8 To 24
'            spdView.Col = j
'            spdView.Text = "0"
'        Next j
'
''        spdView.Col = 26
''        If spdView.Text = "" Then
''            spdView.Col = 25
''            spdView.Value = False
''        Else
''            spdView.Col = 25
''            spdView.Value = True
''        End If
'
'    Next i
'
'    Do While Not RS01.EOF
'        For i = 1 To spdView.MaxRows
'            spdView.Row = i
'            spdView.Col = 1
'            If Format(spdView.Text, "YYYY-MM-DD") = RS01(0) Then
'
'                For j = 3 To spdView.MaxCols
'                    If j = 7 Then
'                        spdView.Col = j
'                        Select Case RS01(j - 1)
'                            Case "1"
'                                spdView.Text = RS01(j - 1) + ":세일"
'                            Case "2"
'                                spdView.Text = RS01(j - 1) + ":목요"
'                            Case "3"
'                                spdView.Text = RS01(j - 1) + ":정상"
'                        End Select
'                    Else
'                        If j = 25 Then
'                            spdView.Col = j
'                            Select Case RS01(j - 1)
'                                Case "Y"
'                                    spdView.Text = True
'                                Case Else
'                                    spdView.Text = False
'                            End Select
'                        Else
'                            spdView.Col = j
'                            If IsNull(RS01(j - 1)) Then
'                                spdView.Text = ""
'                            Else
'                                spdView.Text = RS01(j - 1)
'                            End If
'                        End If
'                    End If
'                Next j
'            End If
'
'        Next i
'
'        RS01.MoveNext
'    Loop
'
'    spdView.MaxRows = spdView.MaxRows + 1
'    spdView.Row = spdView.MaxRows
'    spdView.RowHidden = True
'
'    spdView.Col = 1
'    spdView.Text = "합계"
'
'    Dim cnt, Tamt, Mamt, Samt As Long
'
'
'    spdView.Col = 8
'    spdView.Formula = "SUM(H1:H" & spdView.MaxRows - 1 & ")"
'    Tamt = spdView.Value
'    txtInput(0).Text = spdView.Text
'    spdView.Col = 9
'    spdView.Formula = "SUM(I1:I" & spdView.MaxRows - 1 & ")"
'    Mamt = spdView.Value
'    txtInput(1).Text = spdView.Text
'
'    spdView.Col = 10
'    spdView.Formula = "SUM(J1:J" & spdView.MaxRows - 1 & ")"
'    Samt = spdView.Value
'    txtInput(2).Text = spdView.Text
'
'    spdView.Col = 11
'    spdView.Formula = "SUM(K1:K" & spdView.MaxRows - 1 & ")"
'    cnt = spdView.Value
'    txtInput(3).Text = spdView.Text
'
'    If cnt = 0 Then
'        spdView.Col = 12
'        spdView.Text = 0
'
'        spdView.Col = 13
'        spdView.Text = 0
'
'        spdView.Col = 14
'        spdView.Text = 0
'    Else
'        spdView.Col = 12
'        spdView.Text = Tamt / cnt
'
'        spdView.Col = 13
'        spdView.Text = Mamt / cnt
'
'        spdView.Col = 14
'        spdView.Text = Samt / cnt
'    End If
'
'    spdView.Col = 15
'    spdView.Formula = "SUM(O1:O" & spdView.MaxRows - 1 & ")"
'    txtInput(4).Text = spdView.Text
'
'    spdView.Col = 16
'    spdView.Formula = "SUM(P1:P" & spdView.MaxRows - 1 & ")"
'    txtInput(5).Text = spdView.Text
'
'    spdView.Col = 17
'    spdView.Formula = "SUM(Q1:Q" & spdView.MaxRows - 1 & ")"
'    txtInput(9).Text = spdView.Text
'
'    spdView.Col = 18
'    spdView.Formula = "SUM(R1:R" & spdView.MaxRows - 1 & ")"
'    txtInput(7).Text = spdView.Text
'
'    spdView.Col = 19
'    spdView.Formula = "SUM(S1:S" & spdView.MaxRows - 1 & ")"
'    txtInput(6).Text = spdView.Text
'
'    spdView.Col = 20
'    spdView.Formula = "SUM(T1:T" & spdView.MaxRows - 1 & ")"
'    txtInput(8).Text = spdView.Text
'
'    spdView.Col = 21
'    spdView.Formula = "SUM(U1:U" & spdView.MaxRows - 1 & ")"
'    spdView.Col = 22
'    spdView.Formula = "SUM(V1:V" & spdView.MaxRows - 1 & ")"
'    spdView.Col = 23
'    spdView.Formula = "SUM(W1:W" & spdView.MaxRows - 1 & ")"
'    spdView.Col = 24
'    spdView.Formula = "SUM(X1:X" & spdView.MaxRows - 1 & ")"
'
'    spdView.MaxRows = spdView.MaxRows - 1
'
'    If txtInput(3).Text = 0 Then
'        txtInput(10).Text = 0
'        txtInput(11).Text = 0
'    Else
'        txtInput(10).Text = Format(txtInput(0).Text / txtInput(3).Text, "#,##0")
'        txtInput(11).Text = Format(txtInput(1).Text / txtInput(3).Text, "#,##0")
'    End If
'End Sub



'Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'    If NewRow <> -1 Then
'        With spdView
'            If NewRow <> -1 Then
'                .Row = Row
'                .Col = 2
'                If spdView.Text = "일" Then
'                    .Col = -1
'                    .BackColor = vbYellow
'                Else
'                    If (Row Mod 2) = 0 Then
'                        .Col = -1
'                        .BackColor = glbGray
'                    Else
'                        .Col = -1
'                        .BackColor = vbWhite
'                    End If
'                End If
'                .Row = NewRow
'                .Col = -1
'                .BackColor = glbYellow
'            End If
'        End With
'    End If
'End Sub

'Private Sub spdView_Change(ByVal Col As Long, ByVal Row As Long)
'    Select Case Col
'        Case 2
'            spdView.Row = Row
'
'            'spdView.Col = 14
'
'            If spdView.Text = "일" Then
'                spdView.Col = -1
'                spdView.BackColor = vbYellow
'            End If
'    End Select
'End Sub

Public Sub DataPrint()
'    Dim ReportFP As String
'    Dim ReportFile As String
'    Dim sData As String
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
'    P_00000.crPrint.Formulas(0) = "월 = '" & Format(dtInput.Value, "yyyy-mm") & "'"
'    P_00000.crPrint.Formulas(1) = "대리점 = '" & Trim(cboOffice.Text) & "'"
'
'
'    sData = Space(2) & LeftH("월  합  계" & Space(12), 12)
'    sData = sData & RightH(Space(11) & txtInput(0).Text, 11) & Space(2)
'    If txtInput(3).Text = 0 Then
'        sData = sData & RightH(Space(11) & Format(0, "#,##0"), 5) & Space(2)
'    Else
'        sData = sData & RightH(Space(11) & Format(txtInput(0).Text / txtInput(3).Text, "#,##0"), 5) & Space(2)
'    End If
'    sData = sData & RightH(Space(11) & txtInput(1).Text, 11) & Space(2)
'    If txtInput(3).Text = 0 Then
'        sData = sData & RightH(Space(11) & Format(0, "#,##0"), 5) & Space(2)
'    Else
'        sData = sData & RightH(Space(11) & Format(txtInput(1).Text / txtInput(3).Text, "#,##0"), 5) & Space(2)
'    End If
'    sData = sData & RightH(Space(11) & txtInput(2).Text, 11) & Space(2)
'    sData = sData & RightH(Space(11) & txtInput(3).Text, 6) & Space(2)
'    sData = sData & RightH(Space(11) & txtInput(7).Text, 5) & Space(2)
'    sData = sData & RightH(Space(11) & txtInput(8).Text, 5) & Space(2)
'    sData = sData & RightH(Space(11) & txtInput(6).Text, 5) & Space(2)
'
'    sData = sData & RightH(Space(11) & txtInput(4).Text, 11) & Space(1)
'    sData = sData & RightH(Space(11) & txtInput(5).Text, 6)
'
'    P_00000.crPrint.Formulas(2) = "합계 = '" & sData & "'"
'    P_00000.crPrint.Formulas(3) = "사업장 = '" & Trim(cboInput(1).Text) & "'"
'
'    Set RS01 = New ADODB.Recordset
'    ReDim sValue(0)
'    sValue(0) = "0"
'    Set RS01 = ExecPro("SP_A_0000", sValue(), Err_Num, Err_Dec)
'    P_00000.crPrint.Formulas(4) = "출력시간 = '" & RS01!DB_DATE & "'"
'
'    Call ReportPrint(ReportFile, "1")
End Sub

Public Sub DataScreen()
'    Dim ReportFP As String
'    Dim ReportFile As String
'    Dim sData As String
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
'    P_00000.crPrint.Formulas(0) = "월 = '" & Format(dtInput.Value, "yyyy-mm") & "'"
'    P_00000.crPrint.Formulas(1) = "대리점 = '" & Trim(cboOffice.Text) & "'"
'
'
'    sData = Space(2) & LeftH("월  합  계" & Space(12), 12)
'    sData = sData & RightH(Space(11) & txtInput(0).Text, 11) & Space(2)
'    If txtInput(3).Text = 0 Then
'        sData = sData & RightH(Space(11) & Format(0, "#,##0"), 5) & Space(2)
'    Else
'        sData = sData & RightH(Space(11) & Format(txtInput(0).Text / txtInput(3).Text, "#,##0"), 5) & Space(2)
'    End If
'    sData = sData & RightH(Space(11) & txtInput(1).Text, 11) & Space(2)
'    If txtInput(3).Text = 0 Then
'        sData = sData & RightH(Space(11) & Format(0, "#,##0"), 5) & Space(2)
'    Else
'        sData = sData & RightH(Space(11) & Format(txtInput(1).Text / txtInput(3).Text, "#,##0"), 5) & Space(2)
'    End If
'    sData = sData & RightH(Space(11) & txtInput(2).Text, 11) & Space(2)
'    sData = sData & RightH(Space(11) & txtInput(3).Text, 6) & Space(2)
'    sData = sData & RightH(Space(11) & txtInput(7).Text, 5) & Space(2)
'    sData = sData & RightH(Space(11) & txtInput(8).Text, 5) & Space(2)
'    sData = sData & RightH(Space(11) & txtInput(6).Text, 5) & Space(2)
'
'    sData = sData & RightH(Space(11) & txtInput(4).Text, 11) & Space(1)
'    sData = sData & RightH(Space(11) & txtInput(5).Text, 6)
'
'    P_00000.crPrint.Formulas(2) = "합계 = '" & sData & "'"
'    P_00000.crPrint.Formulas(3) = "사업장 = '" & Trim(cboInput(1).Text) & "'"
'
'    Set RS01 = New ADODB.Recordset
'    ReDim sValue(0)
'    sValue(0) = "0"
'    Set RS01 = ExecPro("SP_A_0000", sValue(), Err_Num, Err_Dec)
'    P_00000.crPrint.Formulas(4) = "출력시간 = '" & RS01!DB_DATE & "'"
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
        TempText = LeftH(RightH(spdView.Text, 5) & Space(5), 6)
        spdView.Col = 2
        TempText = TempText & LeftH(spdView.Text & Space(4), 4)

        spdView.Col = 26
        If Trim(spdView.Text) = "" Then
            TempText = TempText & LeftH("N" & Space(4), 4)
        Else
            TempText = TempText & LeftH("Y" & Space(4), 4)
        End If
        
        spdView.Col = 8
        TempText = TempText & RightH(Space(11) & spdView.Text, 11) & Space(2)
        spdView.Col = 12
        TempText = TempText & RightH(Space(5) & spdView.Text, 5) & Space(2)
        spdView.Col = 9
        TempText = TempText & RightH(Space(11) & spdView.Text, 11) & Space(2)
        spdView.Col = 13
        TempText = TempText & RightH(Space(5) & spdView.Text, 5) & Space(2)
        spdView.Col = 10
        TempText = TempText & RightH(Space(11) & spdView.Text, 11) & Space(2)
        spdView.Col = 11
        TempText = TempText & RightH(Space(6) & spdView.Text, 6) & Space(2)
        spdView.Col = 19
        TempText = TempText & RightH(Space(5) & spdView.Text, 5) & Space(2)
        spdView.Col = 17
        TempText = TempText & RightH(Space(5) & spdView.Text, 5) & Space(2)
        spdView.Col = 18
        TempText = TempText & RightH(Space(5) & spdView.Text, 5) & Space(2)
        spdView.Col = 15
        TempText = TempText & RightH(Space(11) & spdView.Text, 11) & Space(1)
        spdView.Col = 16
        TempText = TempText & RightH(Space(6) & spdView.Text, 6) & Space(1)
        
        Print #1, TempText
    Next i
    
    Close #1
End Sub

Public Sub DataSave()
    Dim i As Integer
    Dim sCode   As String
    Dim bChk As Boolean
    'ReDim sValue(2)
    
    On Error GoTo ERR_RTN
        
    sCode = Mid(cboOffice.Text, 2, 6)
    If sCode = "" Then
        MsgBox "저장은 특정 가맹점를 선택하여 작업하여 주십시요.", vbInformation, "확인"
        cboOffice.SetFocus
        Exit Sub
    End If
 
    ReDim sValue(2)
    
    For i = 1 To spdView.MaxRows
        
        spdView.Row = i
        spdView.Col = 25
        bChk = spdView.Text
        spdView.Col = 3
        If Trim(spdView.Text) <> "" And bChk = False Then

            sValue(0) = sCode
            spdView.Col = 1
            sValue(1) = Format(spdView.Text, "YYYY-MM-DD")
            sValue(2) = "N"
                        
            Call ExecPro("SP_04011_01_ALL", sValue(), Err_Num, Err_Dec)
            
            If Err_Num <> 0 Then
                MsgBox "[" & Err_Num & "] " & Err_Dec
            End If
        
        End If
        
    Next i
    Call Data_Display
    MsgBox "해당되는 데이터가 정상적으로 저장이 되었습니다.", vbInformation
    
ERR_RTN:
    PanelsMsg Err.Description
    'Resume Next
End Sub

Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If NewRow <> -1 Then
        With spdView
            If NewRow <> -1 Then
                .Row = Row
                If (Row Mod 2) = 0 Then
                    .Col = -1
                    .BackColor = oldColor
                Else
                    .Col = -1
                    .BackColor = oldColor
                End If
                
                .Row = NewRow
                .Col = -1
                oldColor = .BackColor
                
                .Row = NewRow
                .Col = -1
                .BackColor = glbYellow
            End If
        End With
    End If

End Sub

Private Sub chkColor_Setting(bColor As Boolean)
    Dim nRow        As Long
    Dim oldMoney    As Double
    Dim newMoney    As Double
    
    On Error GoTo ERR_RTN
    
    With spdView
    
        If .DataRowCnt <= 1 Then Exit Sub
        
        .Row = 1:   .Col = 12:  oldMoney = Val(Replace(.Text, ",", ""))
        
        For nRow = 2 To .DataRowCnt - 1
            .Row = nRow:    .Col = 12:  newMoney = Val(Replace(.Text, ",", ""))
            
            If Not bColor Then
                .Col = -1
                .BackColor = vbWhite
            
            ' 100% 이상 향상된 가맹점은 녹색 적용
            ElseIf (oldMoney * 2) <= newMoney Then
                .Col = -1
                .BackColor = &HC0FFC0    'vbGreen
            
            ' 50%이하로 하향된 가맹점은 빨강 적용
            ElseIf (oldMoney / 2) >= newMoney Then
                .Col = -1
                .BackColor = &HC0C0FF   'vbRed
            
            Else
                .Col = -1
                .BackColor = vbWhite
            
            End If
        
            oldMoney = newMoney
        
        Next nRow
    End With
    
    Exit Sub
    
ERR_RTN:
    MsgBox Err.Description

End Sub

