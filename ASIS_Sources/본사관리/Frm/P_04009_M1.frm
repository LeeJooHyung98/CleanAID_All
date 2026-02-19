VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_04009_M1 
   BorderStyle     =   1  '단일 고정
   Caption         =   "[전사업장]점별 월간 매출현황"
   ClientHeight    =   10035
   ClientLeft      =   5565
   ClientTop       =   5085
   ClientWidth     =   16080
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_04009_M1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   16080
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   10035
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16080
      _ExtentX        =   28363
      _ExtentY        =   17701
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_04009_M1.frx":058A
      Begin Threed.SSPanel SSPanel 
         Height          =   810
         Index           =   0
         Left            =   15
         TabIndex        =   1
         Top             =   9210
         Width           =   16050
         _ExtentX        =   28310
         _ExtentY        =   1429
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   3
            Left            =   30
            TabIndex        =   20
            Top             =   90
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
            Index           =   4
            Left            =   2310
            TabIndex        =   21
            Top             =   90
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
            Left            =   4590
            TabIndex        =   22
            Top             =   405
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
            Left            =   4590
            TabIndex        =   23
            Top             =   90
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
            Left            =   9150
            TabIndex        =   24
            Top             =   90
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
            Left            =   6870
            TabIndex        =   25
            Top             =   90
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
            Left            =   9150
            TabIndex        =   26
            Top             =   405
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
            Left            =   6870
            TabIndex        =   27
            Top             =   405
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
            Left            =   11430
            TabIndex        =   28
            Top             =   90
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
            Caption         =   "반품 수량"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   23
            Left            =   11430
            TabIndex        =   29
            Top             =   405
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
            Left            =   30
            TabIndex        =   30
            Top             =   405
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
            Left            =   2310
            TabIndex        =   31
            Top             =   405
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
            Left            =   1170
            TabIndex        =   32
            Top             =   90
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
            Left            =   1170
            TabIndex        =   33
            Top             =   405
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
            Left            =   3450
            TabIndex        =   34
            Top             =   90
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
            Left            =   3450
            TabIndex        =   35
            Top             =   405
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
            Left            =   5730
            TabIndex        =   36
            Top             =   90
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
            Left            =   8010
            TabIndex        =   37
            Top             =   90
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
            Left            =   10290
            TabIndex        =   38
            Top             =   90
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
            Left            =   5730
            TabIndex        =   39
            Top             =   405
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
            Left            =   8010
            TabIndex        =   40
            Top             =   405
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
            Left            =   10290
            TabIndex        =   41
            Top             =   405
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
            Left            =   12570
            TabIndex        =   42
            Top             =   90
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
            Left            =   12570
            TabIndex        =   43
            Top             =   405
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
      Begin Threed.SSPanel panInput 
         Height          =   615
         Left            =   15
         TabIndex        =   2
         Top             =   540
         Width           =   16050
         _ExtentX        =   28310
         _ExtentY        =   1085
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   8880
            TabIndex        =   3
            Top             =   60
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   556
            _Version        =   262144
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   1500
            TabIndex        =   4
            Top             =   60
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   262144
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   14
            Left            =   3030
            TabIndex        =   5
            Top             =   60
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "사 업 장"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   4440
            TabIndex        =   6
            Top             =   60
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            _Version        =   262144
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   13
            Left            =   120
            TabIndex        =   7
            Top             =   60
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "해당년월"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   15
            Left            =   7440
            TabIndex        =   8
            Top             =   60
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "가 맹 점"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   9
         Top             =   15
         Width           =   8385
         _ExtentX        =   14790
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
         PictureBackground=   "P_04009_M1.frx":063C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   1
         Left            =   8415
         TabIndex        =   10
         Top             =   15
         Width           =   7650
         _ExtentX        =   13494
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
         PictureBackground=   "P_04009_M1.frx":083E
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   11
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
            Picture         =   "P_04009_M1.frx":0A40
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   12
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
            Picture         =   "P_04009_M1.frx":0FDA
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   13
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "인쇄"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_04009_M1.frx":1574
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   14
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
            Picture         =   "P_04009_M1.frx":1B0E
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   15
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
            Picture         =   "P_04009_M1.frx":20A8
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   16
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
            Picture         =   "P_04009_M1.frx":2642
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   17
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
            Picture         =   "P_04009_M1.frx":2BDC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   18
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
            Picture         =   "P_04009_M1.frx":3176
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   8025
         Left            =   15
         TabIndex        =   19
         Top             =   1170
         Width           =   16050
         _Version        =   524288
         _ExtentX        =   28310
         _ExtentY        =   14155
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
         SpreadDesigner  =   "P_04009_M1.frx":3710
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_04009_M1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

'Private Sub cboInput_Click(Index As Integer)
'    Dim sCode As String
'
'    If Index = 1 Then
'        sCode = Trim(Mid(Trim(cboInput(1)) & Space(10), 2, 4))
'
'        Call Get_가맹점리스트(cboInput(0), sCode)
'    End If
'End Sub



Private Sub spdDisplay(Rs As ADODB.Recordset)
    
    
    Call fpSpread_Display(spdView, Rs)
    
    spdView.ColsFrozen = 9 '틀고정
    
    spdView.Row = -1
    
    spdView.Col = 1
    spdView.ColWidth(1) = 10
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft

    spdView.Col = 2
    spdView.ColWidth(2) = 5
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter

    spdView.Col = 3
    spdView.ColWidth(3) = 18
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft

    spdView.Col = 4
    spdView.ColWidth(4) = 6
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter
    
    spdView.Col = 5
    spdView.ColWidth(5) = 6
    spdView.CellType = CellTypePic
    spdView.TypePicMask = "9-999"
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter

    spdView.Col = 6
    spdView.ColWidth(6) = 6
    spdView.CellType = CellTypePic
    spdView.TypePicMask = "9-999"
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter
    
    spdView.Col = 7
    spdView.ColWidth(7) = 5
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter

    spdView.Col = 8
    spdView.ColWidth(8) = 10
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 9
    spdView.ColWidth(9) = 10
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 10
    spdView.ColWidth(10) = 10
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 11
    spdView.ColWidth(11) = 5
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 12
    spdView.ColWidth(12) = 5
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 13
    spdView.ColWidth(13) = 6
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
    
    spdView.Col = 14
    spdView.ColWidth(14) = 6
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
    
    spdView.Col = 15
    spdView.ColWidth(15) = 10
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
    
    spdView.Col = 16
    spdView.ColWidth(16) = 5
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 17
    spdView.ColWidth(17) = 6
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
    
    spdView.Col = 18
    spdView.ColWidth(18) = 5
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
    
    spdView.Col = 19
    spdView.ColWidth(19) = 10
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
    
    spdView.Col = 20
    spdView.ColWidth(20) = 5
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 21
    spdView.ColWidth(21) = 5
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
    
    spdView.Col = 22
    spdView.ColWidth(22) = 12
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
     
    spdView.Col = 23
    spdView.ColWidth(23) = 12
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
     
    spdView.Col = 24
    spdView.ColWidth(24) = 12
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
    
    spdView.Col = 25
    spdView.ColWidth(25) = 5
    spdView.CellType = CellTypeCheckBox
    spdView.TypeCheckCenter = True
    spdView.Value = False
    spdView.TypeHAlign = TypeHAlignCenter
    spdView.TypeVAlign = TypeVAlignCenter
    
    spdView.Col = 26
    spdView.ColWidth(26) = 8
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter

    
'    spdView.ColHidden = True
    
'    Set spdView.DataSource = Nothing
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


Private Sub Form_Activate()
    Call Data_Display
End Sub

Private Sub Form_Load()

    'Call Data_Display

End Sub


Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i As Integer
    
    'For i = 0 To 11
    '    txtNum(i).Value = 0
    'Next i
    
    ReDim sValue(3)
    
    sValue(0) = Mid(panCaption(1).Caption, 2, 4)
    
    If Mid(panCaption(2).Caption, 2, 6) = "000000" Then
        sValue(1) = ""
    Else
        sValue(1) = Mid(panCaption(2).Caption, 2, 6)
    End If
    
    sValue(2) = panCaption(0).Caption & "-01"
    sValue(3) = panCaption(0).Caption & "-31"
    
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(sValue(0)) = False Then Exit Sub
        
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
            .Col = 8:  .Text = RS01!출고수량 & ""              ' 8
                        
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
        
        If .MaxRows > 0 Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Row = .Row
            .Row2 = .Row
            .Col = 1
            .Col2 = .MaxCols
            .BlockMode = True
            .BackColor = &HC0FFC0
            .BlockMode = False
        
            .Col = 4:  .Text = "합계"
            
            .Col = 5:  .Formula = "SUM(E1:E" & .MaxRows - 1 & ")"
            .Col = 6:  .Formula = "SUM(F1:F" & .MaxRows - 1 & ")"
            .Col = 7:  .Formula = "SUM(G1:G" & .MaxRows - 1 & ")"
            .Col = 8:  .Formula = "SUM(H1:H" & .MaxRows - 1 & ")"
            
            .Col = 12: .Formula = "SUM(L1:L" & .MaxRows - 1 & ")"
            .Col = 13: .Formula = "SUM(M1:M" & .MaxRows - 1 & ")"
            
            .Col = 14: .Formula = "SUM(N1:N" & .MaxRows - 1 & ") / " & .MaxRows - 1 & ""
            .Col = 15: .Formula = "SUM(O1:O" & .MaxRows - 1 & ") / " & .MaxRows - 1 & ""
            .Col = 16: .Formula = "SUM(P1:P" & .MaxRows - 1 & ") / " & .MaxRows - 1 & ""
            
            .Col = 17: .Formula = "SUM(Q1:Q" & .MaxRows - 1 & ")"
            .Col = 18: .Formula = "SUM(R1:R" & .MaxRows - 1 & ")"
            .Col = 19: .Formula = "SUM(S1:S" & .MaxRows - 1 & ")"
            .Col = 20: .Formula = "SUM(T1:T" & .MaxRows - 1 & ")"
            .Col = 21: .Formula = "SUM(U1:U" & .MaxRows - 1 & ")"
            .Col = 22: .Formula = "SUM(V1:V" & .MaxRows - 1 & ")"
            .Col = 23: .Formula = "SUM(W1:W" & .MaxRows - 1 & ")"
            
            .Col = 24: .Formula = "SUM(X1:X" & .MaxRows - 1 & ")"
            .Col = 25: .Formula = "SUM(Y1:Y" & .MaxRows - 1 & ")"
            .Col = 26: .Formula = "SUM(Z1:Z" & .MaxRows - 1 & ")"
            .Col = 27: .Formula = "SUM(AA1:AA" & .MaxRows - 1 & ")"
            .Col = 28: .Formula = "SUM(AB1:AB" & .MaxRows - 1 & ")"
            .Col = 29: .Formula = "SUM(AC1:AC" & .MaxRows - 1 & ")"
            .Col = 30: .Formula = "SUM(AD1:AD" & .MaxRows - 1 & ")"
            .Col = 31: .Formula = "SUM(AE1:AE" & .MaxRows - 1 & ")"
            
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
        
        .Redraw = True
    End With
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If NewRow <> -1 Then
    
           With spdView

                If NewRow <> -1 Then
                    .Row = Row
                    .Col = 2
                    If spdView.Text = "일" Then
                        spdView.Col = -1
                        spdView.BackColor = vbYellow
                    Else
                        If (Row Mod 2) = 0 Then
                            .Col = -1
                            .BackColor = glbGray
                        Else
                            .Col = -1
                            .BackColor = vbWhite
                        End If
                    End If
                    .Row = NewRow
                    .Col = -1
                    .BackColor = glbYellow
                End If
            End With

    End If
End Sub
Private Sub spdView_Change(ByVal Col As Long, ByVal Row As Long)
    Select Case Col
        Case 2
            spdView.Row = Row
            
            'spdView.Col = 14
            
            If spdView.Text = "일" Then
                spdView.Col = -1
                spdView.BackColor = vbYellow
            End If
    End Select
End Sub

Private Sub spdView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        'PopupMenu P_00000.PopUp
    End If
End Sub

Public Sub DataSave()
End Sub
