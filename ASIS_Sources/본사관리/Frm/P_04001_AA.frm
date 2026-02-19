VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{B6C10482-FB89-11D4-93C9-006008A7EED4}#1.0#0"; "TeeChart5.ocx"
Begin VB.Form P_04001_A 
   Caption         =   "가맹점별 매출집계 (특정일)"
   ClientHeight    =   9645
   ClientLeft      =   3270
   ClientTop       =   3420
   ClientWidth     =   16170
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_04001_AA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9645
   ScaleWidth      =   16170
   WindowState     =   2  '최대화
   Begin TeeChart.TChart TChart1 
      Height          =   3750
      Left            =   5085
      TabIndex        =   44
      Top             =   2940
      Width           =   6000
      Base64          =   $"P_04001_AA.frx":058A
   End
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   9645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16170
      _ExtentX        =   28522
      _ExtentY        =   17013
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_04001_AA.frx":0614
      Begin TeeChart.ChartPageNavigator ChartPageNavigator1 
         Height          =   510
         Left            =   15
         Negotiate       =   -1  'True
         OleObjectBlob   =   "P_04001_AA.frx":06E6
         TabIndex        =   30
         Top             =   15
         Width           =   8520
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   750
         Index           =   0
         Left            =   15
         TabIndex        =   1
         Top             =   8880
         Width           =   16140
         _ExtentX        =   28469
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   0
            Left            =   60
            TabIndex        =   2
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
            Index           =   1
            Left            =   2340
            TabIndex        =   3
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
            Index           =   3
            Left            =   4620
            TabIndex        =   4
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
            Index           =   2
            Left            =   4620
            TabIndex        =   5
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
            Index           =   4
            Left            =   9180
            TabIndex        =   6
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
            Index           =   5
            Left            =   6900
            TabIndex        =   7
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
            Index           =   6
            Left            =   9180
            TabIndex        =   8
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
            Index           =   7
            Left            =   6900
            TabIndex        =   9
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
            Index           =   8
            Left            =   11460
            TabIndex        =   10
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
            Caption         =   "반품 수량"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   9
            Left            =   11460
            TabIndex        =   11
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
            Index           =   12
            Left            =   60
            TabIndex        =   12
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
            Index           =   13
            Left            =   2340
            TabIndex        =   13
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
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   14
         Top             =   540
         Width           =   16140
         _ExtentX        =   28469
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1230
            Style           =   2  '드롭다운 목록
            TabIndex        =   28
            Top             =   60
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Left            =   1230
            TabIndex        =   15
            Top             =   420
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Format          =   57671680
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   11
            Left            =   45
            TabIndex        =   16
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "지 사 명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   10
            Left            =   45
            TabIndex        =   17
            Top             =   420
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "매출일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox chkColor 
            Height          =   255
            Left            =   4410
            TabIndex        =   29
            Top             =   120
            Width           =   1335
            _Version        =   851970
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "색상 적용"
            UseVisualStyle  =   -1  'True
            Value           =   1
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   8565
         TabIndex        =   18
         Top             =   15
         Width           =   7590
         _ExtentX        =   13388
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
         Caption         =   " 가맹점별 매출집계 (특정일)(P_04001_A)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_04001_AA.frx":0738
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   4440
         Index           =   1
         Left            =   15
         TabIndex        =   19
         Top             =   1335
         Width           =   16140
         _ExtentX        =   28469
         _ExtentY        =   7832
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
         PictureBackground=   "P_04001_AA.frx":093A
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   20
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
            Picture         =   "P_04001_AA.frx":0B3C
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
            Picture         =   "P_04001_AA.frx":10D6
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   22
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
            Picture         =   "P_04001_AA.frx":1670
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   23
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
            Picture         =   "P_04001_AA.frx":1C0A
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   24
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
            Picture         =   "P_04001_AA.frx":21A4
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   25
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
            Picture         =   "P_04001_AA.frx":273E
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   26
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
            Picture         =   "P_04001_AA.frx":2CD8
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   27
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
            Picture         =   "P_04001_AA.frx":3272
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   3075
         Left            =   15
         TabIndex        =   43
         Top             =   5790
         Width           =   16140
         _Version        =   524288
         _ExtentX        =   28469
         _ExtentY        =   5424
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
         MaxCols         =   29
         SpreadDesigner  =   "P_04001_AA.frx":380C
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_04001_A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents CPrt    As CCAIDPrinter
Attribute CPrt.VB_VarHelpID = -1
Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cboInput_Change(Index As Integer)
'    Select Case Index
'        Case 0
'            Call Data_Display
'    End Select
End Sub

Private Sub cboOffice_Click()
    Call Data_Display
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


Private Sub dtInput_LostFocus()
    'Call Data_Display
End Sub

Private Sub dtInput_Change()
    Call Data_Display
End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    
    'cmdBtn(5).Enabled = True
    cmdBtn(6).Enabled = True
    
    '본사 여부
    If HeadOffice = MASTER_OFFICE_CODE Then
        cboOffice.Locked = False
    Else
        cboOffice.Locked = True
    End If
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
    
'    If P_04001_A_Flag = False Then
'
''        dtInput.Value = Date
''
''        Call Master_tblComboAdd(cboOffice)
''
''        ReDim sValue(2)
''
''        sValue(0) = "1"
''
''        Set RS01 = New ADODB.Recordset
''        Set RS01 = ExecPro("SP_04001_00_ALL", sValue(), Err_Num, Err_Dec)
''
''        spdView.MaxCols = RS01.Fields.Count
''        spdView.MaxRows = RS01.RecordCount
''
''        Call spdDisplay
''        Call GetColWidth(REG_App, Me.Name, spdView)
''
''        Call fpSpread_Display(spdView, RS01)
        
'       P_04001_A_Flag = True
'    End If
End Sub

Private Sub spdDisplay()

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
        .OperationMode = OperationModeSingle
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With

    dtInput.Value = DateAdd("d", -1, Date)
    
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
    
'    Call Master_tblComboAdd(cboOffice)
    
'    ReDim sValue(2)
'
'    sValue(0) = "1"
'
'    Set RS01 = New ADODB.Recordset
'    Set RS01 = ExecPro("SP_04001_00_ALL", sValue(), Err_Num, Err_Dec)
'
'    spdView.MaxCols = RS01.Fields.Count
'    spdView.MaxRows = RS01.RecordCount
'
'    Call spdDisplay
'    Call GetColWidth(REG_App, Me.Name, spdView)
'
'    Call fpSpread_Display(spdView, RS01)
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set CPrt = Nothing
    'P_04001_A_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn
    
    Dim i As Integer
    
    For i = 0 To 11
        txtNum(i).Value = 0
    Next i
    
    TChart1.Series(0).Clear
    TChart1.Series(1).Clear
    
    ReDim sValue(1)
    
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    sValue(1) = Format(dtInput.Value, "YYYY-MM-DD")
    
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_04001_03", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_04001_03", sValue(), Err_Num, Err_Dec)
    End If
        
    With spdView
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = RS01!가맹점코드 & ""                 ' 1
            .Col = 2:  .Text = RS01!가맹점명 & ""                   ' 2
            .Col = 3:  .Text = RS01!지사금액 & ""                   ' 3
            .Col = 4:  .Text = RS01!가맹점금액 & ""                 ' 4
            .Col = 5:  .Text = RS01!접수수량 & ""                   ' 5
            .Col = 6:  .Text = RS01!출고수량 & ""                   ' 6
                        
            If Len(RS01!시작택번호) = 9 Then
                .Col = 7:  .Text = Format(RS01!시작택번호, "000-00-0000") & "" ' 7
            Else
                .Col = 7:  .Text = RS01!시작택번호 & ""             ' 7
            End If
            
            If Len(RS01!종료택번호) = 9 Then
                .Col = 8:  .Text = Format(RS01!종료택번호, "000-00-0000") & "" ' 8
            Else
                .Col = 8:  .Text = RS01!종료택번호 & ""             ' 8
            End If
            
            .Col = 9:  .Text = RS01!전영업일매출액 & ""                   ' 9
            
            
            .Col = 10:  .Text = RS01!접수금액 & ""                   ' 9
            .Col = 11: .Text = RS01!현금입금 + RS01!카드금액 & ""   '10
            
            If RS01!접수수량 = 0 Then
                .Col = 12: .Text = 0 & ""                               '11
                .Col = 13: .Text = 0 & ""                               '12
                .Col = 14: .Text = 0 & ""                               '13
            Else
                .Col = 12: .Text = RS01!접수금액 / RS01!접수수량 & ""   '11
                .Col = 13: .Text = RS01!지사금액 / RS01!접수수량 & ""   '12
                .Col = 14: .Text = RS01!가맹점금액 / RS01!접수수량 & "" '13
            End If
            
            .Col = 15: .Text = RS01!현금입금 & ""                   '14
            .Col = 16: .Text = RS01!카드금액 & ""                   '15
            .Col = 17: .Text = RS01!카드건수 & ""                   '16
            .Col = 18: .Text = RS01!쿠폰금액 & ""                   '17
            .Col = 19: .Text = RS01!쿠폰건수 & ""                   '18
            .Col = 20: .Text = RS01!발생마일리지 & ""               '19
            .Col = 21: .Text = RS01!사용마일리지 & ""               '20
            .Col = 22: .Text = RS01!삭제마일리지 & ""               '21
            .Col = 23: .Text = RS01!반품환불금액 & ""               '22
            .Col = 24: .Text = RS01!반품환불건수 & ""               '23
            .Col = 25: .Text = RS01!세탁환불금액 & ""               '24
            .Col = 26: .Text = RS01!세탁환불건수 & ""               '25
            .Col = 27: .Text = RS01!재세탁수량 & ""                 '26
            .Col = 28: .Text = RS01!수선금액 & ""                   '27
            .Col = 29: .Text = RS01!수선수량 & ""                   '28
                        
            TChart1.Series(0).Add RS01!접수금액, RS01!가맹점코드, vbRed
            TChart1.Series(1).Add RS01!접수수량, RS01!가맹점코드, vbBlue
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        If .MaxRows > 0 Then
            ' 합계 출력
            Dim nCol    As Long
            For nCol = 3 To .MaxCols
                Select Case nCol
                    Case 3: Call SpreadSum(spdView, 2, nCol)
                    Case Else: Call SpreadSum(spdView, -1, nCol)
                End Select
            Next nCol
        End If
        
'        If .MaxRows > 0 Then
'            .MaxRows = .MaxRows + 1
'            .Row = .MaxRows
'
'            .Row = .Row
'            .Row2 = .Row
'            .Col = 1
'            .Col2 = .MaxCols
'            .BlockMode = True
'            .BackColor = &HC0FFC0
'            .BlockMode = False
'
'            .Col = 2:  .Text = "합계"
'            .Col = 3:  .Formula = "SUM(C1:C" & .MaxRows - 1 & ")"
'            .Col = 4:  .Formula = "SUM(D1:D" & .MaxRows - 1 & ")"
'            .Col = 5:  .Formula = "SUM(E1:E" & .MaxRows - 1 & ")"
'            .Col = 6:  .Formula = "SUM(F1:F" & .MaxRows - 1 & ")"
'
'            .Col = 9:  .Formula = "SUM(I1:I" & .MaxRows - 1 & ")"
'
'            .Col = 10: .Formula = "SUM(J1:J" & .MaxRows - 1 & ")"
'            .Col = 11: .Formula = "SUM(K1:K" & .MaxRows - 1 & ")"
'
'            .Col = 12: .Formula = "SUM(J1:J" & .MaxRows - 1 & ") / SUM(E1:E" & .MaxRows - 1 & ")"
'            .Col = 13: .Formula = "SUM(C1:C" & .MaxRows - 1 & ") / SUM(E1:E" & .MaxRows - 1 & ")"
'            .Col = 14: .Formula = "SUM(D1:D" & .MaxRows - 1 & ") / SUM(E1:E" & .MaxRows - 1 & ")"
'
'            '.Col = 11: .Formula = "SUM(K1:K" & .MaxRows - 1 & ") / " & .MaxRows - 1
'            '.Col = 12: .Formula = "SUM(L1:L" & .MaxRows - 1 & ") / " & .MaxRows - 1
'            '.Col = 13: .Formula = "SUM(M1:M" & .MaxRows - 1 & ") / " & .MaxRows - 1
'
'            .Col = 15: .Formula = "SUM(O1:O" & .MaxRows - 1 & ")"
'            .Col = 16: .Formula = "SUM(P1:P" & .MaxRows - 1 & ")"
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
'
'
'            .Col = 10:  txtNum(0).Value = .Value  '전체매출액
'            .Col = 12: txtNum(10).Value = Val(.Value) '전체단가
'            .Col = 13: txtNum(11).Value = Val(.Value) '지사단가
'
'            .Col = 3: txtNum(1).Value = .Value   '지사매출
'            .Col = 4: txtNum(2).Value = .Value   '가맹점매출
'
'            .Col = 5: txtNum(3).Value = .Value   '입고수량
'
'            .Col = 28: txtNum(7).Value = .Value   '수선금액
'            .Col = 29: txtNum(6).Value = .Value   '수선수량
'
'            .Col = 16: txtNum(4).Value = .Value   '카드금액
'            .Col = 17: txtNum(5).Value = .Value   '카드수량
'
'
'            .Col = 24: txtNum(8).Value = .Value   '반품수량
'            .Col = 27: txtNum(9).Value = .Value   '재세탁수량
'        End If
        
        ' 색상을 적용 한다.
        Call chkColor_Setting(chkColor.Value)
        
        .Redraw = True
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

'private Sub Data_Display()
'    Dim i As Integer
'    Dim j As Integer
'    Dim sStartTag As String
'    Dim sEndTag As String
'    Dim lTotal(1) As Long
'
'    ReDim sValue(2)
'
'    sValue(0) = "0"
'    sValue(1) = Mid(cboOffice.Text, 2, 4)
'    sValue(2) = Format(dtInput.Value, "YYYY-MM-DD")
'
'    spdView.MaxRows = 0
'
'    Set RS01 = New ADODB.Recordset
'    Set RS01 = ExecPro("SP_04001_00_ALL", sValue(), Err_Num, Err_Dec)
'
'    spdView.MaxCols = RS01.Fields.Count
'    spdView.MaxRows = RS01.RecordCount
'
'    Call spdDisplay
'    Call GetColWidth(REG_App, Me.Name, spdView)
'
'
'    Call spdDisplay
'    Call GetColWidth(REG_App, Me.Name, spdView)
'
'    For i = 0 To 9
'        txtInput(i).Text = 0
'    Next i
'    With spdView
'        .Redraw = False
'        For i = 1 To RS01.RecordCount
'            .Row = i
'
'            For j = 1 To spdView.MaxCols
'                spdView.Col = j
'
'                If j = 2 Then
'                    If RS01(j - 1) = "Y" Then
'                        spdView.Text = RS01(j - 1) + ":개점"
'                    Else
'                        spdView.Text = RS01(j - 1) + ":폐점"
'                    End If
'                ElseIf j = 7 Then
'
'                        Select Case RS01(j - 1)
'                            Case "1": spdView.Text = RS01(j - 1) + ":세일"
'                            Case "2": spdView.Text = RS01(j - 1) + ":목요"
'                            Case "3": spdView.Text = RS01(j - 1) + ":정상"
'                        End Select
'                ElseIf j = 25 Then
'
'                            Select Case RS01(j - 1)
'                                Case "Y":  spdView.Text = True
'                                Case Else: spdView.Text = False
'                            End Select
'                Else
'                    spdView.Col = j: spdView.Text = RS01(j - 1)
'                End If
'            Next j
'
'            RS01.MoveNext
'        Next i
'        .Redraw = True
'    End With
'
'
'    RS01.Close
'
'    If spdView.MaxRows = 0 Then
'        For i = 0 To 9
'            txtInput(i).Text = 0
'        Next i
'    Else
'        spdView.AutoCalc = True
'
'        spdView.MaxRows = spdView.MaxRows + 1
'        spdView.Row = spdView.MaxRows
'        spdView.RowHidden = True
'
'        spdView.Col = 1
'        spdView.Text = "합계"
'
'        Dim cnt, Tamt, Mamt, Samt As Long
'
'
'        spdView.Col = 8
'        spdView.Formula = "SUM(H1:H" & spdView.MaxRows - 1 & ")"
'        Tamt = spdView.Value
'        txtInput(0).Text = spdView.Text
'        spdView.Col = 9
'        spdView.Formula = "SUM(I1:I" & spdView.MaxRows - 1 & ")"
'        Mamt = spdView.Value
'        txtInput(1).Text = spdView.Text
'
'        spdView.Col = 10
'        spdView.Formula = "SUM(J1:J" & spdView.MaxRows - 1 & ")"
'        Samt = spdView.Value
'        txtInput(2).Text = spdView.Text
'
'        spdView.Col = 11
'        spdView.Formula = "SUM(K1:K" & spdView.MaxRows - 1 & ")"
'        cnt = spdView.Value
'        txtInput(3).Text = spdView.Text
'
'        If cnt = 0 Then
'            spdView.Col = 12
'            spdView.Text = 0
'
'            spdView.Col = 13
'            spdView.Text = 0
'
'            spdView.Col = 14
'            spdView.Text = 0
'        Else
'            spdView.Col = 12
'            spdView.Text = Tamt / cnt
'
'            spdView.Col = 13
'            spdView.Text = Mamt / cnt
'
'            spdView.Col = 14
'            spdView.Text = Samt / cnt
'        End If
'
'        spdView.Col = 15
'        spdView.Formula = "SUM(O1:O" & spdView.MaxRows - 1 & ")"
'        txtInput(4).Text = spdView.Text
'
'        spdView.Col = 16
'        spdView.Formula = "SUM(P1:P" & spdView.MaxRows - 1 & ")"
'        txtInput(5).Text = spdView.Text
'
'        spdView.Col = 17
'        spdView.Formula = "SUM(Q1:Q" & spdView.MaxRows - 1 & ")"
'        txtInput(9).Text = spdView.Text
'
'        spdView.Col = 18
'        spdView.Formula = "SUM(R1:R" & spdView.MaxRows - 1 & ")"
'        txtInput(7).Text = spdView.Text
'
'        spdView.Col = 19
'        spdView.Formula = "SUM(S1:S" & spdView.MaxRows - 1 & ")"
'        txtInput(6).Text = spdView.Text
'
'        spdView.Col = 20
'        spdView.Formula = "SUM(T1:T" & spdView.MaxRows - 1 & ")"
'        txtInput(8).Text = spdView.Text
'
'        If txtInput(3).Text = 0 Then
'            txtInput(10).Text = 0
'            txtInput(11).Text = 0
'        Else
'            txtInput(10).Text = Format(txtInput(0).Text / txtInput(3).Text, "#,##0")
'            txtInput(11).Text = Format(txtInput(1).Text / txtInput(3).Text, "#,##0")
'        End If
'        spdView.Col = 21
'        spdView.Formula = "SUM(U1:U" & spdView.MaxRows - 1 & ")"
'        spdView.Col = 22
'        spdView.Formula = "SUM(V1:V" & spdView.MaxRows - 1 & ")"
'        spdView.Col = 23
'        spdView.Formula = "SUM(W1:W" & spdView.MaxRows - 1 & ")"
'        spdView.Col = 24
'        spdView.Formula = "SUM(X1:X" & spdView.MaxRows - 1 & ")"
'
'        spdView.MaxRows = spdView.MaxRows - 1
'    End If
'End Sub

Public Sub DataSave()
    Dim i As Integer
    Dim sCode   As String
    
    ReDim sValue(2)
        
    sCode = Mid(cboOffice.Text, 2, 4)
    
    If sCode = "" Then
        MsgBox "저장은 특정 지사를 선택하여 작업하여 주십시요.", vbInformation, "확인"
        cboOffice.SetFocus
        Exit Sub
    End If
    
    sValue(0) = "0"
    sValue(1) = sCode
    sValue(2) = Format(dtInput.Value, "YYYY-MM-DD")
    
    ' 이전 자료 삭제
    Call ExecPro("SP_04001_01_Master", sValue(), Err_Num, Err_Dec)
    
    ReDim sValue(14)
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        
        sValue(0) = "0"
        sValue(1) = Format(dtInput.Value, "YYYY-MM-DD")   ' 수금일자
        
        spdView.Col = 1: sValue(2) = Mid(spdView.Text, 2, 3)                      ' 매장코드
        spdView.Col = 2: sValue(3) = spdView.Value                                ' 입고량
        spdView.Col = 3: sValue(12) = spdView.Value                               ' 출고량
        spdView.Col = 4: sValue(4) = IIf(spdView.Text = "-", "", Mid(spdView.Text, 1, 1) & Mid(spdView.Text, 3, 3))  ' 시작택
        spdView.Col = 5: sValue(5) = IIf(spdView.Text = "-", "", Mid(spdView.Text, 1, 1) & Mid(spdView.Text, 3, 3))  ' 종료택
        spdView.Col = 6: sValue(6) = spdView.Value                                ' 금액
        spdView.Col = 8: sValue(7) = spdView.Value                                ' 카드금액
        spdView.Col = 9: sValue(8) = spdView.Value                                ' 카드건수
        spdView.Col = 10: sValue(9) = spdView.Value                               ' 재세탁수량
        spdView.Col = 11: sValue(10) = spdView.Value                              ' 수선수량
        spdView.Col = 12: sValue(11) = spdView.Value                              ' 반품수량
        
        spdView.Col = 14                                 ' UpdateChk
        If spdView.Text = "수" Then
            sValue(13) = "U"
        Else
            sValue(13) = ""
        End If
        
        ' 지사 코드
        sValue(14) = sCode
        
        If Int(sValue(6)) > 0 And Int(sValue(3)) <= 0 Then
            MsgBox "[오류] " & "수금액이 있을경우 입고수량은 반드시 입력하셔야 합니다.", vbCritical, "확인"
        Else
            Call ExecPro("SP_04001_02_Master", sValue(), Err_Num, Err_Dec)
            
            If Err_Num <> 0 Then
                MsgBox "[" & Err_Num & "] " & Err_Dec
            End If
        
        End If
        
    Next i

End Sub

Public Sub DataCancel()
    Call Data_Display
End Sub

Public Sub DataPrint()
'    Dim i, ii As Integer
'    Dim ReportFP As String
'    Dim ReportFile As String
'    Dim sData As String
'
'    On Error GoTo ERR_RTN
'
'    PanelsMsg ""
'
'    ii = 0
'
'    For i = 1 To spdView.MaxRows
'        spdView.Row = i
'        spdView.Col = 13
'        If spdView.Value = True Then
'            ii = ii + 1
'        End If
'    Next i
'
'    If ii = 0 Then Exit Sub
'
'    Call PrintDesc
'
'    ReDim PrtParam.Param(12)
'    With PrtParam
'        .Param(0) = "P_04001_MASTER"
'        .Param(1) = "수금일자 : " & Format(dtInput.Value, "YYYY-MM-DD")
'        .Param(2) = "날씨: " & ""
'
'        .Param(3) = "수금액 : " & txtInput(2).Text
'        .Param(4) = "단가 : ' " & txtInput(7).Text
'        .Param(5) = "월수금 : " & txtInput(9).Text
'        .Param(6) = "일수금 : " & txtInput(8).Text
'
'        .Param(7) = "입고량 : " & txtInput(0).Text
'        .Param(8) = "출고량 : " & txtInput(1).Text
'        .Param(9) = "재세탁 : " & txtInput(3).Text
'        .Param(10) = "수선 : " & txtInput(4).Text
'        .Param(11) = "반품 : " & txtInput(5).Text
'        .Param(12) = "카드금액 : " & txtInput(6).Text
'    End With
'
'    Set CPrt = New CCAIDPrinter
'    CPrt.PRT_04001_01_MASTER Printer, 0
'
'    Exit Sub
'
'ERR_RTN:
'    PanelsMsg Err.Description
    
End Sub

Public Sub DataPrint_OLD()
'    Dim i, ii As Integer
'    Dim ReportFP As String
'    Dim ReportFile As String
'    Dim sData As String
'
'    On Error GoTo ERR_RTN
'
'    PanelsMsg ""
'
'    ReportFP = GetIniStr("REPORT", "FilePath", "", m_iniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    ii = 0
'
'    For i = 1 To spdView.MaxRows
'        spdView.Row = i
'        spdView.Col = 13
'        If spdView.Value = True Then
'            ii = ii + 1
'        End If
'    Next i
'
'    If ii = 0 Then
'        Exit Sub
'    End If
'
'    Call PrintDesc
'
'    For i = 0 To 30
'        P_00000.crPrint.Formulas(i) = ""
'    Next
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'    P_00000.crPrint.Formulas(0) = "수금일자 = '" & Format(dtInput.Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(1) = "입고량 = '" & txtInput(0).Text & "'"
'    P_00000.crPrint.Formulas(2) = "출고량 = '" & txtInput(1).Text & "'"
'    P_00000.crPrint.Formulas(3) = "수금액 = '" & txtInput(2).Text & "'"
'    P_00000.crPrint.Formulas(4) = "재세탁 = '" & txtInput(3).Text & "'"
'    P_00000.crPrint.Formulas(5) = "수선 = '" & txtInput(4).Text & "'"
'    P_00000.crPrint.Formulas(6) = "반품 = '" & txtInput(5).Text & "'"
'    P_00000.crPrint.Formulas(7) = "단가 = '" & txtInput(7).Text & "'"
'    P_00000.crPrint.Formulas(8) = "일수금 = '" & txtInput(8).Text & "'"
'    P_00000.crPrint.Formulas(9) = "월수금 = '" & txtInput(9).Text & "'"
'
'    Call ReportPrint(ReportFile, "1")
'    Exit Sub
'
'ERR_RTN:
'    PanelsMsg Err.Description
    
End Sub

Public Sub DataScreen()


End Sub

Public Sub DataScreen_OLD()
'    Dim i, ii As Integer
'    Dim ReportFP As String
'    Dim ReportFile As String
'    Dim sData As String
'
'    ReportFP = GetIniStr("REPORT", "FilePath", "", m_iniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    ii = 0
'
'    For i = 1 To spdView.MaxRows
'        spdView.Row = i
'        spdView.Col = 13
'        If spdView.Value = True Then
'            ii = ii + 1
'        End If
'    Next i
'
'    If ii = 0 Then
'        Exit Sub
'    End If
'
'    Call PrintDesc
'
'    For i = 0 To 30
'        P_00000.crPrint.Formulas(i) = ""
'    Next
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'    P_00000.crPrint.Formulas(0) = "수금일자 = '" & Format(dtInput.Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(1) = "입고량 = '" & txtInput(0).Text & "'"
'    P_00000.crPrint.Formulas(2) = "출고량 = '" & txtInput(1).Text & "'"
'    P_00000.crPrint.Formulas(3) = "수금액 = '" & txtInput(2).Text & "'"
'    P_00000.crPrint.Formulas(4) = "재세탁 = '" & txtInput(3).Text & "'"
'    P_00000.crPrint.Formulas(5) = "수선 = '" & txtInput(4).Text & "'"
'    P_00000.crPrint.Formulas(6) = "반품 = '" & txtInput(5).Text & "'"
'    P_00000.crPrint.Formulas(7) = "단가 = '" & txtInput(7).Text & "'"
'    P_00000.crPrint.Formulas(8) = "일수금 = '" & txtInput(8).Text & "'"
'    P_00000.crPrint.Formulas(9) = "월수금 = '" & txtInput(9).Text & "'"
'
'    Call ReportPrint(ReportFile, "2")
End Sub

Private Sub PrintDesc()
    Dim i As Integer
    Dim TempText As String
    Dim TempFP As String
    Dim TempFile As String
    Dim hFile   As Integer
    
    Dim iDanga As Long
    
    On Error GoTo FileError:
    
    TempFP = GetIniStr("REPORT", "TempPath", "", m_iniFile)
    TempFile = TempFP & "\Temp.txt"
    
    hFile = FreeFile
    Open TempFile For Output As #hFile
    
    TempText = ""
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        spdView.Col = 13
        If spdView.Value = True Then
            TempText = Left(i & Space(3), 3) & Space(1) & "|"
            
            spdView.Col = 1
            TempText = TempText & LeftH(spdView.Text & Space(16), 16) & "|"
            
            spdView.Col = 13
            If spdView.Text = "월수금" Then
                TempText = TempText & "M" & "|"
            Else
                TempText = TempText & " " & "|"
            End If
            
            spdView.Col = 6
            TempText = TempText & RightH(Space(12) & spdView.Text, 12) & "|"
            iDanga = Val(spdView.Value)
            spdView.Col = 2
            TempText = TempText & RightH(Space(8) & spdView.Text, 8) & "|"
            If Val(spdView.Value) <> 0 Then iDanga = iDanga / Val(spdView.Value)
            
            spdView.Col = 3
            TempText = TempText & RightH(Space(8) & spdView.Text, 8) & "|"
            
            TempText = TempText & RightH(Space(12) & Format(iDanga, "#,##0"), 12) & "|"
            
            spdView.Col = 8
            TempText = TempText & RightH(Space(8) & spdView.Text, 8) & "|"
            spdView.Col = 9
            TempText = TempText & RightH(Space(8) & spdView.Text, 8) & "|"
            spdView.Col = 10
            TempText = TempText & RightH(Space(8) & spdView.Text, 8) & Space(2) & "|"
            spdView.Col = 11
            TempText = TempText & RightH(Space(8) & spdView.Text, 8) & Space(2) & "|"
            spdView.Col = 12
            TempText = TempText & RightH(Space(8) & spdView.Text, 8) & Space(2) & "|"
            spdView.Col = 4
            TempText = TempText & LeftH(spdView.Text & Space(5), 5) & " ~ "
            spdView.Col = 5
            TempText = TempText & LeftH(spdView.Text & Space(5), 5) & "|"
            
            If spdView.BackColor = &HD8FCFE Then
                TempText = TempText & " *"
            Else
                TempText = TempText & "  "
            End If
            
            
            
            Print #hFile, TempText
        End If
    Next i
    
    Close #hFile
    Exit Sub
    
FileError:
    MsgBox Err.Description
    If Err.Number = 55 Then
        Resume Next
    End If
    Close #hFile
End Sub


Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'    If NewRow <> -1 Then
'        With spdView
'            If NewRow <> -1 Then
'                .Row = Row
'                .Col = 14
'                If spdView.Text = "수" Then
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
End Sub

Private Sub chkColor_Setting(bColor As Boolean)
    Dim nRow        As Long
    Dim oldMoney    As Double
    Dim newMoney    As Double
    
    On Error GoTo ERR_RTN
    
    With spdView
    
        If .DataRowCnt <= 1 Then Exit Sub
        
        
        For nRow = 1 To .DataRowCnt - 1
            .Row = nRow
            .Col = 9:  oldMoney = Val(Replace(.Text, ",", ""))
            .Col = 10:  newMoney = Val(Replace(.Text, ",", ""))
            
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
