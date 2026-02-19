VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm일일매출마감 
   BorderStyle     =   1  '단일 고정
   Caption         =   "일일매출 마감"
   ClientHeight    =   9255
   ClientLeft      =   6390
   ClientTop       =   2535
   ClientWidth     =   15045
   ControlBox      =   0   'False
   Icon            =   "frm일일매출마감.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   15045
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   16325
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm일일매출마감.frx":0A02
      Begin XtremeSuiteControls.TabControl TabControl1 
         Height          =   8025
         Left            =   15
         TabIndex        =   6
         Top             =   1215
         Width           =   15015
         _Version        =   851970
         _ExtentX        =   26485
         _ExtentY        =   14155
         _StockProps     =   68
         Appearance      =   3
         Color           =   16
         PaintManager.BoldSelected=   -1  'True
         PaintManager.HotTracking=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         PaintManager.ButtonMargin=   "2,3,2,3"
         ItemCount       =   3
         Item(0).Caption =   " 일일마감 "
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "TabControlPage1"
         Item(1).Caption =   "일일마감 현황"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "TabControlPage2"
         Item(2).Caption =   "접수집계"
         Item(2).ControlCount=   1
         Item(2).Control(0)=   "TabControlPage3"
         Begin XtremeSuiteControls.TabControlPage TabControlPage3 
            Height          =   7545
            Left            =   -69970
            TabIndex        =   133
            Top             =   450
            Visible         =   0   'False
            Width           =   14955
            _Version        =   851970
            _ExtentX        =   26379
            _ExtentY        =   13309
            _StockProps     =   1
            Page            =   2
            Begin XtremeSuiteControls.GroupBox GroupBox1 
               Height          =   4755
               Left            =   9480
               TabIndex        =   139
               Top             =   150
               Width           =   5265
               _Version        =   851970
               _ExtentX        =   9287
               _ExtentY        =   8387
               _StockProps     =   79
               Caption         =   "기타 정보"
               UseVisualStyle  =   -1  'True
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   10
                  Left            =   300
                  TabIndex        =   140
                  Top             =   330
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "가맹점 수선"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":0A74
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   11
                  Left            =   300
                  TabIndex        =   141
                  Top             =   735
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "재세탁 수량"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":0C9A
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   12
                  Left            =   300
                  TabIndex        =   142
                  Top             =   1140
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "운동화 세탁"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":0EC0
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   13
                  Left            =   300
                  TabIndex        =   143
                  Top             =   1545
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "가  죽 세탁"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":10E6
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   14
                  Left            =   300
                  TabIndex        =   144
                  Top             =   1950
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "카페트 세탁"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":130C
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   15
                  Left            =   300
                  TabIndex        =   145
                  Top             =   2355
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "반  품 세탁"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":1532
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   41
                  Left            =   1740
                  TabIndex        =   146
                  Top             =   330
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   42
                     Left            =   975
                     TabIndex        =   147
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "점"
                     PictureBackground=   "frm일일매출마감.frx":1758
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtNum06 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   148
                     Top             =   45
                     Width           =   945
                     _Version        =   262145
                     _ExtentX        =   1667
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit txtCost11 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   149
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   43
                     Left            =   2670
                     TabIndex        =   150
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":1E22
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   44
                  Left            =   1740
                  TabIndex        =   151
                  Top             =   735
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   45
                     Left            =   975
                     TabIndex        =   152
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "점"
                     PictureBackground=   "frm일일매출마감.frx":24EC
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtNum07 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   153
                     Top             =   45
                     Width           =   945
                     _Version        =   262145
                     _ExtentX        =   1667
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit txtCost12 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   154
                     Top             =   45
                     Visible         =   0   'False
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   46
                     Left            =   2670
                     TabIndex        =   155
                     Top             =   60
                     Visible         =   0   'False
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":2BB6
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   47
                  Left            =   1740
                  TabIndex        =   156
                  Top             =   1140
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   48
                     Left            =   975
                     TabIndex        =   157
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "점"
                     PictureBackground=   "frm일일매출마감.frx":3280
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtNum08 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   158
                     Top             =   45
                     Width           =   945
                     _Version        =   262145
                     _ExtentX        =   1667
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit txtCost13 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   159
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   49
                     Left            =   2670
                     TabIndex        =   160
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":394A
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   50
                  Left            =   1740
                  TabIndex        =   161
                  Top             =   1545
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   51
                     Left            =   975
                     TabIndex        =   162
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "점"
                     PictureBackground=   "frm일일매출마감.frx":4014
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtNum09 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   163
                     Top             =   45
                     Width           =   945
                     _Version        =   262145
                     _ExtentX        =   1667
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit txtCost14 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   164
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   52
                     Left            =   2670
                     TabIndex        =   165
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":46DE
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   53
                  Left            =   1740
                  TabIndex        =   166
                  Top             =   1950
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   54
                     Left            =   975
                     TabIndex        =   167
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "점"
                     PictureBackground=   "frm일일매출마감.frx":4DA8
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtNum10 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   168
                     Top             =   45
                     Width           =   945
                     _Version        =   262145
                     _ExtentX        =   1667
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit txtCost15 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   169
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   55
                     Left            =   2670
                     TabIndex        =   170
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":5472
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   56
                  Left            =   1740
                  TabIndex        =   171
                  Top             =   2355
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   57
                     Left            =   975
                     TabIndex        =   172
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "점"
                     PictureBackground=   "frm일일매출마감.frx":5B3C
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtNum11 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   173
                     Top             =   45
                     Width           =   945
                     _Version        =   262145
                     _ExtentX        =   1667
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit txtCost16 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   174
                     Top             =   45
                     Visible         =   0   'False
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   58
                     Left            =   2670
                     TabIndex        =   175
                     Top             =   60
                     Visible         =   0   'False
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":6206
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   20
                  Left            =   300
                  TabIndex        =   176
                  Top             =   2760
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "외주세탁비"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":68D0
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   73
                  Left            =   1740
                  TabIndex        =   177
                  Top             =   2760
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtCost17 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   178
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   74
                     Left            =   2670
                     TabIndex        =   179
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":6AF6
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   18
                  Left            =   300
                  TabIndex        =   180
                  Top             =   3720
                  Visible         =   0   'False
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "삼성카드할인"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":71C0
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   68
                  Left            =   1740
                  TabIndex        =   181
                  Top             =   3720
                  Visible         =   0   'False
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   69
                     Left            =   975
                     TabIndex        =   182
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "건"
                     PictureBackground=   "frm일일매출마감.frx":73E6
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtNum17 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   183
                     Top             =   45
                     Width           =   945
                     _Version        =   262145
                     _ExtentX        =   1667
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit txtCost26 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   184
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   70
                     Left            =   2670
                     TabIndex        =   185
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":7AB0
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   19
                  Left            =   300
                  TabIndex        =   186
                  Top             =   4125
                  Visible         =   0   'False
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "삼성카드고객"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":817A
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   71
                  Left            =   1740
                  TabIndex        =   187
                  Top             =   4125
                  Visible         =   0   'False
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtNum18 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   188
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   72
                     Left            =   2670
                     TabIndex        =   189
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "명"
                     PictureBackground=   "frm일일매출마감.frx":83A0
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
            End
            Begin FPSpreadADO.fpSpread sprCloth 
               Height          =   7380
               Left            =   60
               TabIndex        =   134
               Top             =   60
               Width           =   9105
               _Version        =   524288
               _ExtentX        =   16060
               _ExtentY        =   13018
               _StockProps     =   64
               AllowDragDrop   =   -1  'True
               AllowMultiBlocks=   -1  'True
               AllowUserFormulas=   -1  'True
               BackColorStyle  =   1
               DAutoHeadings   =   0   'False
               DAutoSave       =   0   'False
               DAutoSizeCols   =   0
               DInformActiveRowChange=   0   'False
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
               MaxRows         =   30
               Protect         =   0   'False
               ScrollBars      =   2
               SpreadDesigner  =   "frm일일매출마감.frx":8A6A
               VisibleCols     =   3
               VisibleRows     =   30
               Appearance      =   1
               HighlightHeaders=   1
               HighlightStyle  =   1
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage2 
            Height          =   7545
            Left            =   -69970
            TabIndex        =   8
            Top             =   450
            Visible         =   0   'False
            Width           =   14955
            _Version        =   851970
            _ExtentX        =   26379
            _ExtentY        =   13309
            _StockProps     =   1
            BackColor       =   16777215
            Page            =   1
            Begin FPSpreadADO.fpSpread sprGrid 
               Height          =   6750
               Left            =   60
               TabIndex        =   125
               Top             =   735
               Width           =   14895
               _Version        =   524288
               _ExtentX        =   26273
               _ExtentY        =   11906
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
               MaxCols         =   32
               MaxRows         =   200
               OperationMode   =   1
               Protect         =   0   'False
               SpreadDesigner  =   "frm일일매출마감.frx":924D
               UserResize      =   1
               VisibleCols     =   11
               VisibleRows     =   50
               Appearance      =   1
               HighlightHeaders=   1
               HighlightStyle  =   1
            End
            Begin MSComCtl2.DTPicker dtpDate 
               Height          =   330
               Index           =   0
               Left            =   960
               TabIndex        =   126
               Top             =   90
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   582
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
               Format          =   56819715
               CurrentDate     =   40279
            End
            Begin MSComCtl2.DTPicker dtpDate 
               Height          =   330
               Index           =   1
               Left            =   2685
               TabIndex        =   127
               Top             =   90
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   582
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
               Format          =   56819715
               CurrentDate     =   40279
            End
            Begin XtremeSuiteControls.PushButton cmdList 
               Height          =   570
               Left            =   13395
               TabIndex        =   130
               Top             =   105
               Width           =   1500
               _Version        =   851970
               _ExtentX        =   2646
               _ExtentY        =   1005
               _StockProps     =   79
               Caption         =   " 조회(&F)"
               Appearance      =   6
               Picture         =   "frm일일매출마감.frx":A886
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "마감일자:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   90
               TabIndex        =   129
               Top             =   150
               Width           =   840
            End
            Begin VB.Label Label 
               Alignment       =   2  '가운데 맞춤
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
               Height          =   210
               Left            =   2460
               TabIndex        =   128
               Top             =   150
               Width           =   180
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage1 
            Height          =   7545
            Left            =   30
            TabIndex        =   7
            Top             =   450
            Width           =   14955
            _Version        =   851970
            _ExtentX        =   26379
            _ExtentY        =   13309
            _StockProps     =   1
            BackColor       =   16777215
            Page            =   0
            Begin XtremeSuiteControls.GroupBox GroupBox 
               Height          =   6825
               Index           =   1
               Left            =   10185
               TabIndex        =   9
               Top             =   645
               Width           =   4740
               _Version        =   851970
               _ExtentX        =   8361
               _ExtentY        =   12039
               _StockProps     =   79
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   29
                  Left            =   1560
                  TabIndex        =   10
                  Top             =   225
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   30
                     Left            =   975
                     TabIndex        =   11
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "건"
                     PictureBackground=   "frm일일매출마감.frx":AF80
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtNum13 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   12
                     Top             =   45
                     Width           =   945
                     _Version        =   262145
                     _ExtentX        =   1667
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit txtCost22 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   13
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   31
                     Left            =   2670
                     TabIndex        =   14
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":B64A
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   825
                  Index           =   37
                  Left            =   120
                  TabIndex        =   15
                  Top             =   225
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   1455
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "판매취소 TAG"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":BD14
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   32
                  Left            =   1560
                  TabIndex        =   16
                  Top             =   1065
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   33
                     Left            =   975
                     TabIndex        =   17
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "건"
                     PictureBackground=   "frm일일매출마감.frx":BF3A
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtNum14 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   18
                     Top             =   45
                     Width           =   945
                     _Version        =   262145
                     _ExtentX        =   1667
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit txtCost23 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   19
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   34
                     Left            =   2670
                     TabIndex        =   20
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":C604
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   1215
                  Index           =   39
                  Left            =   120
                  TabIndex        =   21
                  Top             =   1065
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   2143
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "반품환불 TAG"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":CCCE
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   35
                  Left            =   1560
                  TabIndex        =   22
                  Top             =   2295
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   36
                     Left            =   975
                     TabIndex        =   23
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "건"
                     PictureBackground=   "frm일일매출마감.frx":CEF4
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtNum15 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   24
                     Top             =   45
                     Width           =   945
                     _Version        =   262145
                     _ExtentX        =   1667
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit txtCost24 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   25
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   37
                     Left            =   2670
                     TabIndex        =   26
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":D5BE
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   1215
                  Index           =   40
                  Left            =   120
                  TabIndex        =   27
                  Top             =   2295
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   2143
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "세탁환불 TAG"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":DC88
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   38
                  Left            =   1560
                  TabIndex        =   28
                  Top             =   3525
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   39
                     Left            =   975
                     TabIndex        =   29
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "건"
                     PictureBackground=   "frm일일매출마감.frx":DEAE
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtNum16 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   30
                     Top             =   45
                     Width           =   945
                     _Version        =   262145
                     _ExtentX        =   1667
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit txtCost25 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   31
                     Top             =   45
                     Visible         =   0   'False
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   40
                     Left            =   2670
                     TabIndex        =   32
                     Top             =   60
                     Visible         =   0   'False
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":E578
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   825
                  Index           =   9
                  Left            =   120
                  TabIndex        =   33
                  Top             =   3525
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   1455
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "누    락 TAG"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":EC42
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel4 
                  Height          =   420
                  Left            =   1560
                  TabIndex        =   34
                  Top             =   3930
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin VB.ComboBox cboMissTag 
                     BeginProperty Font 
                        Name            =   "굴림체"
                        Size            =   9.75
                        Charset         =   129
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   60
                     Style           =   2  '드롭다운 목록
                     TabIndex        =   35
                     Top             =   60
                     Width           =   2955
                  End
               End
               Begin Threed.SSPanel SSPanel3 
                  Height          =   420
                  Left            =   1560
                  TabIndex        =   36
                  Top             =   3090
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin VB.ComboBox cboRepay 
                     BeginProperty Font 
                        Name            =   "굴림체"
                        Size            =   9.75
                        Charset         =   129
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   60
                     Style           =   2  '드롭다운 목록
                     TabIndex        =   37
                     Top             =   60
                     Width           =   2955
                  End
               End
               Begin Threed.SSPanel SSPanel2 
                  Height          =   420
                  Left            =   1560
                  TabIndex        =   38
                  Top             =   1860
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin VB.ComboBox cboReturn 
                     BeginProperty Font 
                        Name            =   "굴림체"
                        Size            =   9.75
                        Charset         =   129
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   60
                     Style           =   2  '드롭다운 목록
                     TabIndex        =   39
                     Top             =   60
                     Width           =   2955
                  End
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   1
                  Left            =   1560
                  TabIndex        =   40
                  Top             =   630
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin VB.ComboBox cboCancel 
                     BeginProperty Font 
                        Name            =   "굴림체"
                        Size            =   9.75
                        Charset         =   129
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   60
                     Style           =   2  '드롭다운 목록
                     TabIndex        =   41
                     Top             =   60
                     Width           =   2955
                  End
               End
               Begin Threed.SSPanel pnlTAG 
                  Height          =   420
                  Index           =   0
                  Left            =   1560
                  TabIndex        =   42
                  Top             =   4395
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
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
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   28
                  Left            =   120
                  TabIndex        =   43
                  Top             =   4395
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "시작 택번호"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":EE68
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlTAG 
                  Height          =   420
                  Index           =   1
                  Left            =   1560
                  TabIndex        =   44
                  Top             =   4800
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
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
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   27
                  Left            =   120
                  TabIndex        =   45
                  Top             =   4800
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "종료 택번호"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":F08E
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   38
                  Left            =   120
                  TabIndex        =   190
                  Top             =   5250
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "발생마일리지"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":F2B4
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   59
                  Left            =   1560
                  TabIndex        =   191
                  Top             =   5250
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtCost18 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   192
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   60
                     Left            =   2670
                     TabIndex        =   193
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":F4DA
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   42
                  Left            =   120
                  TabIndex        =   194
                  Top             =   5655
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "삭제마일리지"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":FBA4
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   63
                  Left            =   1560
                  TabIndex        =   195
                  Top             =   5655
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtCost20 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   196
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   64
                     Left            =   2670
                     TabIndex        =   197
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":FDCA
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   76
                  Left            =   1560
                  TabIndex        =   198
                  Top             =   1470
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtCost29 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   199
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   78
                     Left            =   2670
                     TabIndex        =   200
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":10494
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   77
                     Left            =   120
                     TabIndex        =   201
                     Top             =   60
                     Width           =   1200
                     _ExtentX        =   2117
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "지사 환급 금액"
                     BevelOuter      =   0
                     Alignment       =   1
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   79
                  Left            =   1560
                  TabIndex        =   202
                  Top             =   2700
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtCost30 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   203
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   80
                     Left            =   2670
                     TabIndex        =   204
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":10B5E
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   81
                     Left            =   120
                     TabIndex        =   205
                     Top             =   60
                     Width           =   1200
                     _ExtentX        =   2117
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "지사 환급 금액"
                     BevelOuter      =   0
                     Alignment       =   1
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
            End
            Begin XtremeSuiteControls.GroupBox GroupBox 
               Height          =   6825
               Index           =   0
               Left            =   60
               TabIndex        =   46
               Top             =   645
               Width           =   5295
               _Version        =   851970
               _ExtentX        =   9340
               _ExtentY        =   12039
               _StockProps     =   79
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   2
                  Left            =   2115
                  TabIndex        =   47
                  Top             =   225
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   14
                     Left            =   975
                     TabIndex        =   48
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "점"
                     PictureBackground=   "frm일일매출마감.frx":11228
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtNum01 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   49
                     Top             =   45
                     Width           =   945
                     _Version        =   262145
                     _ExtentX        =   1667
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit txtCost01 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   50
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   15
                     Left            =   2670
                     TabIndex        =   51
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":118F2
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   0
                  Left            =   675
                  TabIndex        =   52
                  Top             =   225
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "일 일 합 계"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":11FBC
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   1
                  Left            =   675
                  TabIndex        =   53
                  Top             =   630
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "출 고 수 량"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":121E2
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   2
                  Left            =   675
                  TabIndex        =   54
                  Top             =   1515
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "반환/현금결제"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":12408
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   3
                  Left            =   675
                  TabIndex        =   55
                  Top             =   1920
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
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
                  PictureBackground=   "frm일일매출마감.frx":1262E
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   1230
                  Index           =   26
                  Left            =   120
                  TabIndex        =   56
                  Top             =   225
                  Width           =   570
                  _ExtentX        =   1005
                  _ExtentY        =   2170
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "매출"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":12854
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   2040
                  Index           =   31
                  Left            =   120
                  TabIndex        =   57
                  Top             =   1515
                  Width           =   570
                  _ExtentX        =   1005
                  _ExtentY        =   3598
                  _Version        =   262144
                  Font3D          =   1
                  CaptionStyle    =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "선불결제"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":12A7A
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   4
                  Left            =   2115
                  TabIndex        =   58
                  Top             =   630
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtNum02 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   59
                     Top             =   45
                     Width           =   945
                     _Version        =   262145
                     _ExtentX        =   1667
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   16
                     Left            =   975
                     TabIndex        =   60
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "점"
                     PictureBackground=   "frm일일매출마감.frx":12CA0
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   5
                  Left            =   2115
                  TabIndex        =   61
                  Top             =   1515
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtCost02 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   62
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   17
                     Left            =   2670
                     TabIndex        =   63
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":1336A
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   75
                     Left            =   975
                     TabIndex        =   137
                     Top             =   45
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":13A34
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtCost28 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   138
                     Top             =   30
                     Width           =   945
                     _Version        =   262145
                     _ExtentX        =   1667
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   6
                  Left            =   2115
                  TabIndex        =   64
                  Top             =   1920
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtNum03 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   65
                     Top             =   45
                     Width           =   945
                     _Version        =   262145
                     _ExtentX        =   1667
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit txtCost03 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   66
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   18
                     Left            =   975
                     TabIndex        =   67
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "건"
                     PictureBackground=   "frm일일매출마감.frx":140FE
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   19
                     Left            =   2670
                     TabIndex        =   68
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":147C8
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   32
                  Left            =   675
                  TabIndex        =   69
                  Top             =   3570
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
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
                  PictureBackground=   "frm일일매출마감.frx":14E92
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   33
                  Left            =   675
                  TabIndex        =   70
                  Top             =   3975
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
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
                  PictureBackground=   "frm일일매출마감.frx":150B8
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   825
                  Index           =   34
                  Left            =   120
                  TabIndex        =   71
                  Top             =   3570
                  Width           =   570
                  _ExtentX        =   1005
                  _ExtentY        =   1455
                  _Version        =   262144
                  Font3D          =   1
                  CaptionStyle    =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "미수결제"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":152DE
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   7
                  Left            =   2115
                  TabIndex        =   72
                  Top             =   3570
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtCost05 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   73
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   21
                     Left            =   2670
                     TabIndex        =   74
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":15504
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   8
                  Left            =   2115
                  TabIndex        =   75
                  Top             =   3975
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtNum04 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   76
                     Top             =   45
                     Width           =   945
                     _Version        =   262145
                     _ExtentX        =   1667
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit txtCost06 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   77
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   22
                     Left            =   975
                     TabIndex        =   78
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "점"
                     PictureBackground=   "frm일일매출마감.frx":15BCE
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   23
                     Left            =   2670
                     TabIndex        =   79
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":16298
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   35
                  Left            =   675
                  TabIndex        =   80
                  Top             =   6060
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "가맹점 마진"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":16962
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   1230
                  Index           =   36
                  Left            =   120
                  TabIndex        =   81
                  Top             =   5250
                  Width           =   570
                  _ExtentX        =   1005
                  _ExtentY        =   2170
                  _Version        =   262144
                  Font3D          =   1
                  CaptionStyle    =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "마진"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":16B88
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   9
                  Left            =   2115
                  TabIndex        =   82
                  Top             =   6060
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtCost09 
                     Height          =   345
                     Left            =   975
                     TabIndex        =   83
                     Top             =   45
                     Width           =   1725
                     _Version        =   262145
                     _ExtentX        =   3043
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   27
                     Left            =   2670
                     TabIndex        =   84
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":16DAE
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   4
                  Left            =   675
                  TabIndex        =   85
                  Top             =   5655
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "지  사 마진"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":17478
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   10
                  Left            =   2115
                  TabIndex        =   86
                  Top             =   5655
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtCost10 
                     Height          =   345
                     Left            =   975
                     TabIndex        =   87
                     Top             =   45
                     Width           =   1725
                     _Version        =   262145
                     _ExtentX        =   3043
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   28
                     Left            =   2670
                     TabIndex        =   88
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":1769E
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   5
                  Left            =   675
                  TabIndex        =   89
                  Top             =   3135
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "미 수 금 액"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":17D68
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   11
                  Left            =   2115
                  TabIndex        =   90
                  Top             =   3135
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtCost04 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   91
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   20
                     Left            =   2670
                     TabIndex        =   92
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":17F8E
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   6
                  Left            =   675
                  TabIndex        =   93
                  Top             =   4410
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
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
                  PictureBackground=   "frm일일매출마감.frx":18658
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   7
                  Left            =   675
                  TabIndex        =   94
                  Top             =   4815
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
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
                  PictureBackground=   "frm일일매출마감.frx":1887E
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   825
                  Index           =   8
                  Left            =   120
                  TabIndex        =   95
                  Top             =   4410
                  Width           =   570
                  _ExtentX        =   1005
                  _ExtentY        =   1455
                  _Version        =   262144
                  Font3D          =   1
                  CaptionStyle    =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "결제합계"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":18AA4
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   12
                  Left            =   2115
                  TabIndex        =   96
                  Top             =   4410
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtCost07 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   97
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   24
                     Left            =   2670
                     TabIndex        =   98
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":18CCA
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   13
                  Left            =   2115
                  TabIndex        =   99
                  Top             =   4815
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtNum05 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   100
                     Top             =   45
                     Width           =   945
                     _Version        =   262145
                     _ExtentX        =   1667
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit txtCost08 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   101
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   25
                     Left            =   975
                     TabIndex        =   102
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "건"
                     PictureBackground=   "frm일일매출마감.frx":19394
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   26
                     Left            =   2670
                     TabIndex        =   103
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":19A5E
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   41
                  Left            =   675
                  TabIndex        =   104
                  Top             =   2325
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "사용마일리지"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":1A128
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   61
                  Left            =   2115
                  TabIndex        =   105
                  Top             =   2325
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtCost19 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   106
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   62
                     Left            =   2670
                     TabIndex        =   107
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":1A34E
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   43
                  Left            =   675
                  TabIndex        =   108
                  Top             =   2730
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "쿠 폰 사 용"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":1AA18
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   65
                  Left            =   2115
                  TabIndex        =   109
                  Top             =   2730
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   66
                     Left            =   975
                     TabIndex        =   110
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "건"
                     PictureBackground=   "frm일일매출마감.frx":1AC3E
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtNum12 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   111
                     Top             =   45
                     Width           =   945
                     _Version        =   262145
                     _ExtentX        =   1667
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit txtCost21 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   112
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   67
                     Left            =   2670
                     TabIndex        =   113
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":1B308
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   21
                  Left            =   675
                  TabIndex        =   114
                  Top             =   5250
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "사용마일리지"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":1B9D2
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   0
                  Left            =   2115
                  TabIndex        =   115
                  Top             =   5250
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtCost27 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   116
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   3
                     Left            =   2670
                     TabIndex        =   117
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":1BBF8
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   52
                  Left            =   675
                  TabIndex        =   268
                  Top             =   1035
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BackColor       =   8454143
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "인터넷 접수"
                  PictureBackgroundStyle=   2
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   114
                  Left            =   2115
                  TabIndex        =   269
                  Top             =   1035
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   115
                     Left            =   975
                     TabIndex        =   270
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "점"
                     PictureBackground=   "frm일일매출마감.frx":1C2C2
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtNum_Internet 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   271
                     Top             =   45
                     Width           =   945
                     _Version        =   262145
                     _ExtentX        =   1667
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit txtCost_Internet 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   272
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   116
                     Left            =   2670
                     TabIndex        =   273
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":1C98C
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
            End
            Begin XtremeSuiteControls.GroupBox GroupBox 
               Height          =   6795
               Index           =   2
               Left            =   5400
               TabIndex        =   118
               Top             =   660
               Width           =   4740
               _Version        =   851970
               _ExtentX        =   8361
               _ExtentY        =   11986
               _StockProps     =   79
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   22
                  Left            =   75
                  TabIndex        =   206
                  Top             =   2490
                  Width           =   4560
                  _ExtentX        =   8043
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  CaptionStyle    =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "정산 내역"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":1D056
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   23
                  Left            =   75
                  TabIndex        =   207
                  ToolTipText     =   "지사 마진"
                  Top             =   2880
                  Width           =   2085
                  _ExtentX        =   3678
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "지사분매출"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":1D27C
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   82
                  Left            =   2145
                  TabIndex        =   208
                  Top             =   2880
                  Width           =   2490
                  _ExtentX        =   4392
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtMaster 
                     Height          =   345
                     Index           =   0
                     Left            =   210
                     TabIndex        =   209
                     Top             =   45
                     Width           =   2115
                     _Version        =   262145
                     _ExtentX        =   3731
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   83
                     Left            =   2670
                     TabIndex        =   210
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":1D4A2
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   25
                  Left            =   75
                  TabIndex        =   211
                  ToolTipText     =   "가맹점 마진의 % 금액"
                  Top             =   4050
                  Width           =   2085
                  _ExtentX        =   3678
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "+ 유통 로열티"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":1DB6C
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   86
                  Left            =   2145
                  TabIndex        =   212
                  Top             =   4050
                  Width           =   2490
                  _ExtentX        =   4392
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtMaster 
                     Height          =   345
                     Index           =   3
                     Left            =   210
                     TabIndex        =   213
                     Top             =   45
                     Width           =   2115
                     _Version        =   262145
                     _ExtentX        =   3731
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   87
                     Left            =   2670
                     TabIndex        =   214
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":1DD92
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   29
                  Left            =   75
                  TabIndex        =   215
                  ToolTipText     =   "카드 승인 금액의 수수료"
                  Top             =   3270
                  Width           =   2085
                  _ExtentX        =   3678
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "- 카드수수료지원금"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":1E45C
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   88
                  Left            =   2145
                  TabIndex        =   216
                  Top             =   3270
                  Width           =   2490
                  _ExtentX        =   4392
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtMaster 
                     Height          =   345
                     Index           =   1
                     Left            =   210
                     TabIndex        =   217
                     Top             =   45
                     Width           =   2115
                     _Version        =   262145
                     _ExtentX        =   3731
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   89
                     Left            =   2670
                     TabIndex        =   218
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":1E682
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   30
                  Left            =   75
                  TabIndex        =   219
                  ToolTipText     =   "카드승인 취소금액의 수수료"
                  Top             =   3660
                  Width           =   2085
                  _ExtentX        =   3678
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "+ 카드수수료환불금"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":1ED4C
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   90
                  Left            =   2145
                  TabIndex        =   220
                  Top             =   3660
                  Width           =   2490
                  _ExtentX        =   4392
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtMaster 
                     Height          =   345
                     Index           =   2
                     Left            =   210
                     TabIndex        =   221
                     Top             =   45
                     Width           =   2115
                     _Version        =   262145
                     _ExtentX        =   3731
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   91
                     Left            =   2670
                     TabIndex        =   222
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":1EF72
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   44
                  Left            =   75
                  TabIndex        =   223
                  ToolTipText     =   "지사로 수금해줘야 할 금액"
                  Top             =   5640
                  Width           =   2085
                  _ExtentX        =   3678
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  ForeColor       =   255
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   12
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "지사 정산 금액"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":1F63C
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   92
                  Left            =   2145
                  TabIndex        =   224
                  Top             =   5640
                  Width           =   2490
                  _ExtentX        =   4392
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777152
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtMaster 
                     Height          =   345
                     Index           =   5
                     Left            =   210
                     TabIndex        =   225
                     Top             =   45
                     Width           =   2115
                     _Version        =   262145
                     _ExtentX        =   3731
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   16777152
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   93
                     Left            =   2670
                     TabIndex        =   226
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":1F862
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   435
                  Index           =   46
                  Left            =   90
                  TabIndex        =   227
                  Top             =   210
                  Width           =   4530
                  _ExtentX        =   7990
                  _ExtentY        =   767
                  _Version        =   262144
                  Font3D          =   1
                  CaptionStyle    =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "수수료 지원 정보 "
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":1FF2C
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   94
                  Left            =   1560
                  TabIndex        =   228
                  Top             =   630
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   95
                     Left            =   975
                     TabIndex        =   229
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "건"
                     PictureBackground=   "frm일일매출마감.frx":20152
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtCard 
                     Height          =   345
                     Index           =   0
                     Left            =   225
                     TabIndex        =   230
                     Top             =   45
                     Width           =   765
                     _Version        =   262145
                     _ExtentX        =   1349
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit txtCard 
                     Height          =   345
                     Index           =   1
                     Left            =   1380
                     TabIndex        =   231
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   96
                     Left            =   2670
                     TabIndex        =   232
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":2081C
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   825
                  Index           =   45
                  Left            =   90
                  TabIndex        =   233
                  Top             =   630
                  Width           =   1485
                  _ExtentX        =   2619
                  _ExtentY        =   1455
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "카드 승인"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":20EE6
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   97
                  Left            =   1560
                  TabIndex        =   234
                  Top             =   1035
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtCard 
                     Height          =   345
                     Index           =   2
                     Left            =   1380
                     TabIndex        =   235
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   98
                     Left            =   2670
                     TabIndex        =   236
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":2110C
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   99
                     Left            =   120
                     TabIndex        =   237
                     Top             =   60
                     Width           =   1200
                     _ExtentX        =   2117
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "수수료 지원금"
                     BevelOuter      =   0
                     Alignment       =   1
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   100
                  Left            =   1560
                  TabIndex        =   238
                  Top             =   1440
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   101
                     Left            =   975
                     TabIndex        =   239
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "건"
                     PictureBackground=   "frm일일매출마감.frx":217D6
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtCard 
                     Height          =   345
                     Index           =   3
                     Left            =   225
                     TabIndex        =   240
                     Top             =   45
                     Width           =   765
                     _Version        =   262145
                     _ExtentX        =   1349
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit txtCard 
                     Height          =   345
                     Index           =   4
                     Left            =   1380
                     TabIndex        =   241
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   102
                     Left            =   2670
                     TabIndex        =   242
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":21EA0
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   825
                  Index           =   47
                  Left            =   90
                  TabIndex        =   243
                  Top             =   1440
                  Width           =   1485
                  _ExtentX        =   2619
                  _ExtentY        =   1455
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "카드 취소"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":2256A
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   103
                  Left            =   1560
                  TabIndex        =   244
                  Top             =   1845
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtCard 
                     Height          =   345
                     Index           =   5
                     Left            =   1380
                     TabIndex        =   245
                     Top             =   45
                     Width           =   1305
                     _Version        =   262145
                     _ExtentX        =   2302
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   104
                     Left            =   2670
                     TabIndex        =   246
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":22790
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   105
                     Left            =   120
                     TabIndex        =   247
                     Top             =   60
                     Width           =   1200
                     _ExtentX        =   2117
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "수수료 지원금"
                     BevelOuter      =   0
                     Alignment       =   1
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   48
                  Left            =   75
                  TabIndex        =   248
                  ToolTipText     =   "세탁환불 + 반품환불 지사 금액"
                  Top             =   4440
                  Width           =   2085
                  _ExtentX        =   3678
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "- 세탁/반품환불금액"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":22E5A
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   106
                  Left            =   2145
                  TabIndex        =   249
                  Top             =   4440
                  Width           =   2490
                  _ExtentX        =   4392
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtMaster 
                     Height          =   345
                     Index           =   4
                     Left            =   210
                     TabIndex        =   250
                     Top             =   45
                     Width           =   2115
                     _Version        =   262145
                     _ExtentX        =   3731
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   107
                     Left            =   2670
                     TabIndex        =   251
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":23080
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   49
                  Left            =   75
                  TabIndex        =   252
                  ToolTipText     =   "가맹점 마진- (반품/세탁환불 총금액 - 시사환급금액) - 카드 수수료 환불금 + 카드 수수료 지원금 - 유통로열티"
                  Top             =   6030
                  Width           =   2085
                  _ExtentX        =   3678
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  ForeColor       =   255
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   12
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "매장 수익금"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":2374A
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   108
                  Left            =   2145
                  TabIndex        =   253
                  Top             =   6030
                  Width           =   2490
                  _ExtentX        =   4392
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777152
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtMaster 
                     Height          =   345
                     Index           =   6
                     Left            =   210
                     TabIndex        =   254
                     Top             =   45
                     Width           =   2115
                     _Version        =   262145
                     _ExtentX        =   3731
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   16777152
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   109
                     Left            =   2670
                     TabIndex        =   255
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":23970
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   84
                  Left            =   2160
                  TabIndex        =   256
                  Top             =   6180
                  Visible         =   0   'False
                  Width           =   2490
                  _ExtentX        =   4392
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtMaster 
                     Height          =   345
                     Index           =   7
                     Left            =   210
                     TabIndex        =   257
                     Top             =   45
                     Width           =   2115
                     _Version        =   262145
                     _ExtentX        =   3731
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   85
                     Left            =   2670
                     TabIndex        =   258
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":2403A
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   24
                  Left            =   90
                  TabIndex        =   259
                  ToolTipText     =   "가맹점 마진의 % 금액"
                  Top             =   6180
                  Visible         =   0   'False
                  Width           =   2085
                  _ExtentX        =   3678
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "+ 로열티1 사용안함"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":24704
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   50
                  Left            =   75
                  TabIndex        =   260
                  ToolTipText     =   "세탁환불 + 반품환불 지사 금액"
                  Top             =   4830
                  Width           =   2085
                  _ExtentX        =   3678
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "- 쿠폰지사금액(60%)"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":2492A
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   110
                  Left            =   2145
                  TabIndex        =   261
                  Top             =   4830
                  Width           =   2490
                  _ExtentX        =   4392
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtMaster 
                     Height          =   345
                     Index           =   8
                     Left            =   210
                     TabIndex        =   262
                     Top             =   45
                     Width           =   2115
                     _Version        =   262145
                     _ExtentX        =   3731
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   111
                     Left            =   2670
                     TabIndex        =   263
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":24B50
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   51
                  Left            =   75
                  TabIndex        =   264
                  ToolTipText     =   "전산사용료 금액"
                  Top             =   5235
                  Width           =   2085
                  _ExtentX        =   3678
                  _ExtentY        =   741
                  _Version        =   262144
                  Font3D          =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림체"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "+ 전산사용료"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일매출마감.frx":2521A
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   112
                  Left            =   2145
                  TabIndex        =   265
                  Top             =   5235
                  Width           =   2490
                  _ExtentX        =   4392
                  _ExtentY        =   741
                  _Version        =   262144
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit txtMaster 
                     Height          =   345
                     Index           =   9
                     Left            =   210
                     TabIndex        =   266
                     Top             =   45
                     Width           =   2115
                     _Version        =   262145
                     _ExtentX        =   3731
                     _ExtentY        =   609
                     _StockProps     =   125
                     Text            =   " 0"
                     ForeColor       =   255
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   12.01
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DataProperty    =   2
                     FocusSelect     =   -1  'True
                     Insert          =   0   'False
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   ""
                     Text            =   " 0"
                     StartText.x     =   2
                     StartText.y     =   1
                     FirstVisPos     =   0
                     HiAnchor        =   0
                     HiNew           =   0
                     CaretHeight     =   20
                     CurNumDataChars =   0
                     MaxDataChars    =   0
                     FirstDataPos    =   0
                     CurPos          =   0
                     MaxLen          =   0
                     DataReadOnly    =   0   'False
                     Mask            =   ""
                     Justification   =   2
                     BorderStyle     =   0
                     FmtControl      =   1
                     NumDecDigits    =   0
                     Undo            =   0
                     Data            =   0
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   113
                     Left            =   2670
                     TabIndex        =   267
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일매출마감.frx":25440
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
            End
            Begin Threed.SSPanel pnlManager 
               Height          =   570
               Left            =   5820
               TabIndex        =   119
               Top             =   75
               Width           =   2475
               _ExtentX        =   4366
               _ExtentY        =   1005
               _Version        =   262144
               PictureBackgroundStyle=   2
               PictureBackground=   "frm일일매출마감.frx":25B0A
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin VB.ComboBox cboManager 
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   405
                  Left            =   60
                  Locked          =   -1  'True
                  Style           =   2  '드롭다운 목록
                  TabIndex        =   120
                  Top             =   90
                  Width           =   2340
               End
            End
            Begin Threed.SSPanel pnlDate 
               Height          =   555
               Left            =   1155
               TabIndex        =   121
               Top             =   90
               Width           =   3165
               _ExtentX        =   5583
               _ExtentY        =   979
               _Version        =   262144
               Enabled         =   0   'False
               PictureBackgroundStyle=   2
               PictureBackground=   "frm일일매출마감.frx":25D30
               BorderWidth     =   0
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin MSComCtl2.DTPicker dtpDay 
                  Height          =   405
                  Left            =   75
                  TabIndex        =   122
                  Top             =   75
                  Width           =   3015
                  _ExtentX        =   5318
                  _ExtentY        =   714
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
                  CustomFormat    =   "yyyy-MM-dd"
                  Format          =   56819712
                  CurrentDate     =   40279
               End
            End
            Begin XtremeSuiteControls.PushButton cmdFinish 
               Height          =   570
               Left            =   13320
               TabIndex        =   0
               Top             =   75
               Width           =   1590
               _Version        =   851970
               _ExtentX        =   2805
               _ExtentY        =   1005
               _StockProps     =   79
               Caption         =   " 업무마감"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               Appearance      =   6
               Picture         =   "frm일일매출마감.frx":25F56
            End
            Begin Threed.SSPanel pnlData 
               Height          =   555
               Index           =   16
               Left            =   75
               TabIndex        =   123
               Top             =   90
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   979
               _Version        =   262144
               Font3D          =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "마감일자"
               PictureBackgroundStyle=   2
               PictureBackground=   "frm일일매출마감.frx":26830
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlData 
               Height          =   555
               Index           =   17
               Left            =   4725
               TabIndex        =   124
               Top             =   90
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   979
               _Version        =   262144
               Font3D          =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "담 당 자"
               PictureBackgroundStyle=   2
               PictureBackground=   "frm일일매출마감.frx":26A56
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
         End
      End
      Begin Threed.SSPanel pnlMsg 
         Height          =   750
         Left            =   15
         TabIndex        =   2
         Top             =   450
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   1323
         _Version        =   262144
         Font3D          =   3
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   15.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton btnDelete 
            Height          =   360
            Left            =   60
            TabIndex        =   132
            Top             =   345
            Visible         =   0   'False
            Width           =   1440
            _Version        =   851970
            _ExtentX        =   2540
            _ExtentY        =   635
            _StockProps     =   79
            Caption         =   "일일마감 삭제"
            UseVisualStyle  =   -1  'True
         End
         Begin Threed.SSCheck chkDate 
            Height          =   240
            Left            =   75
            TabIndex        =   131
            Top             =   75
            Visible         =   0   'False
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   423
            _Version        =   262144
            ForeColor       =   0
            Caption         =   "세탁환불,반품환불 제외"
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   13515
            TabIndex        =   5
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frm일일매출마감.frx":26C7C
         End
         Begin XtremeSuiteControls.PushButton btnEditDate 
            Height          =   360
            Left            =   1530
            TabIndex        =   135
            Top             =   345
            Visible         =   0   'False
            Width           =   1440
            _Version        =   851970
            _ExtentX        =   2540
            _ExtentY        =   635
            _StockProps     =   79
            Caption         =   "마감일자 수정"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PushButton1 
            Height          =   360
            Left            =   3240
            TabIndex        =   136
            Top             =   360
            Visible         =   0   'False
            Width           =   1440
            _Version        =   851970
            _ExtentX        =   2540
            _ExtentY        =   635
            _StockProps     =   79
            Caption         =   "조회"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   3
         Top             =   15
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   741
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   0
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
         Caption         =   "      일일매출 마감"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm일일매출마감.frx":27D0E
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm일일매출마감.frx":27F34
            Top             =   -15
            Width           =   765
         End
      End
   End
   Begin Threed.SSPanel pnlProg 
      Height          =   1215
      Left            =   6285
      TabIndex        =   4
      Top             =   5655
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
      Picture         =   "frm일일매출마감.frx":28AFE
      Caption         =   " "
      BorderWidth     =   0
      BevelOuter      =   1
      BevelInner      =   2
      Alignment       =   4
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
End
Attribute VB_Name = "frm일일매출마감"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_Activate   As Boolean
Dim 마감일자     As String
Dim chkSale      As String

Dim 시작택번호   As String
Dim 마지막택번호 As String

Private Sub btnDelete_Click()
    On Error GoTo ErrRtn
    
    Rtn = MsgBox("일일마감 자료를 삭제하시겠습니까?", vbQuestion + vbYesNo, "삭제")
    
    If Rtn = vbNo Then Exit Sub
    
    Query = "DELETE FROM TB_일일마감"
    Query = Query & " WHERE 마감일자 = '" & Format(dtpDay.Value, "YYYY-MM-DD") & "'"
    ADOCon.Execute Query
    
    MsgBox "일일마감 자료가 삭제되었습니다.", vbInformation, "확인"
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub btnEditDate_Click()
    pnlDate.Enabled = Not pnlDate.Enabled
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 5:
            '마감일자가 오늘 이전인 경우 마감작업을 안하면 종료가 안된다.
            If Format(Date, "YYYY-MM-DD") > Format(dtpDay.Value, "YYYY-MM-DD") Then
                Query = "SELECT * FROM TB_일일마감"
                Query = Query & " WHERE 마감일자 = '" & Format(dtpDay.Value, "YYYY-MM-DD") & "'"
                Set ADORs = New ADODB.RecordSet
                ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

                If ADORs.EOF Then
                    ADORs.Close
                    Set ADORs = Nothing

                    MsgBox "일일마감을 해주세요.", vbInformation, "확인"
                    Exit Sub
                End If
            End If
            
            Unload Me
    End Select
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub


Private Sub Get_택번호(Query As String, Combo As ComboBox)
    On Error GoTo ErrRtn
    
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    Do Until ADORs.EOF
        Combo.AddItem Format(ADORs!택번호, "000-00-0000")
        
        ADORs.MoveNext
    Loop
    ADORs.Close
    Set ADORs = Nothing
    
    Exit Sub
    
ErrRtn:

End Sub

Private Sub cmdFinish_Click()
    On Error GoTo ErrRtn
    
    마감일자 = Format(dtpDay.Value, "YYYY-MM-DD")
    DoEvents
        
    If Get_일일마감여부(마감일자) = True Then
        MsgBox "일마감이 완료 되었으므로 마감작업을 할 수 없습니다", vbInformation, "확인"
        
        Exit Sub
    End If
    
    If cboManager.ListIndex < 0 Then
        MsgBox "근무자를 선택하세요.", vbInformation, "확인"
        
        cboManager.SetFocus
        Exit Sub
    End If
    
    If Format(Date, "yyyy-MM-dd") = 마감일자 And Format(Time, "hh:mm:ss") <= "11:59:59" Then
        Rtn = MsgBox("[ " & 마감일자 & " ]" & "은 오늘 일자 입니다. " & vbNewLine & _
                     "오늘 일자를 마감하면 오늘 일자로 더이상 접수가 불가능 합니다." & vbNewLine & _
                     "계속 진행 하시겠습니까..?", vbQuestion + vbYesNo, "일일마감")
        If Rtn = vbNo Then Exit Sub
    End If
    
    
    
    Rtn = MsgBox("[ " & 마감일자 & " ]" & " 마감을 하시겠습니까..?", vbQuestion + vbYesNo, "일일마감")
    If Rtn = vbNo Then Exit Sub
    
    pnlDate.Enabled = False
    pnlManager.Enabled = False
    cmdFinish.Enabled = False
    
    pnlProg.Left = 90
    pnlProg.Top = 1305
    pnlProg.Visible = True
    DoEvents
    
    '------------------------------------------------------------------------------------------------------
    ' TB_일일마감
    '------------------------------------------------------------------------------------------------------
    Query = "DELETE FROM TB_일일마감 WHERE 마감일자 = '" & 마감일자 & "'"
    ADOCon.Execute Query
    
    Call Sale_Check
    
    ' 마일리지 마감 ( 3개월동안 이용 실적이 없을 경우 마일리지 삭제 )
    Call Set_마일리지삭제
    
    '--------------------------------------------------------------------
    '
    '--------------------------------------------------------------------
    Query = "INSERT INTO TB_일일마감("
    Query = Query & "  가맹점코드"         ' 1
    Query = Query & ", 마감일자"           ' 2
    Query = Query & ", 접수금액"           ' 3
    Query = Query & ", 접수수량"           ' 4
    Query = Query & ", 출고수량"           ' 5
    Query = Query & ", 반품수량"           ' 6
    Query = Query & ", 재세탁수량"         ' 7
    Query = Query & ", 수선금액"           ' 8
    Query = Query & ", 수선수량"           ' 9
    Query = Query & ", 판매구분"           '10
    Query = Query & ", 시작택번호"         '11
    Query = Query & ", 종료택번호"         '12
    Query = Query & ", 쿠폰금액"           '13
    Query = Query & ", 쿠폰건수"           '14
    Query = Query & ", 발생마일리지"       '15
    Query = Query & ", 사용마일리지"       '16
    Query = Query & ", 삭제마일리지"       '17
    Query = Query & ", 현금입금"           '18
    Query = Query & ", 카드금액"           '19
    Query = Query & ", 카드건수"           '20
    Query = Query & ", 반품환불금액"       '21
    Query = Query & ", 반품환불건수"       '22
    Query = Query & ", 세탁환불금액"       '23
    Query = Query & ", 세탁환불건수"       '24
    Query = Query & ", 삼성카드할인금액"   '25
    Query = Query & ", 삼성카드할인건수"   '26
    Query = Query & ", 삼성카드할인고객수" '27
    Query = Query & ", 근무자명"           '28
    Query = Query & ", 지사금액"           '29
    Query = Query & ", 가맹점금액"         '30
    Query = Query & ", 운동화금액"         '31
    Query = Query & ", 운동화건수"         '32
    Query = Query & ", 운동화비율"         '33
    Query = Query & ", 카페트금액"         '34
    Query = Query & ", 카페트건수"         '35
    Query = Query & ", 명품세탁금액"       '36
    Query = Query & ", 명품세탁건수"       '37
    Query = Query & ", 명품세탁비율"       '38
    Query = Query & ", 명품염색금액"       '39
    Query = Query & ", 명품염색건수"       '40
    Query = Query & ", 명품염색비율"       '41
    Query = Query & ", 마감여부"           '42
    Query = Query & ", 본사전송여부"       '43
    Query = Query & ", 지사코드"           '44
    Query = Query & ", 마감시간"           '45
    
    Query = Query & ", 로열티정보1"        '46
    Query = Query & ", 로열티정보2"        '47
    Query = Query & ", 수수료정보"         '48
    Query = Query & ", 반품환불지사금액"   '49
    Query = Query & ", 세탁환불지사금액"   '50
    
    Query = Query & ", 카드취소금액"       '51
    Query = Query & ", 카드취소건수"       '52
    
    Query = Query & ", 로열티금액1"        '53
    Query = Query & ", 로열티금액2"        '54
    Query = Query & ", 수수료승인금액"     '55
    Query = Query & ", 수수료취소금액"     '56
    
    Query = Query & ", 미수카드건수"       '57
    Query = Query & ", 미수카드금액"       '58
    Query = Query & ", 미수현금수금금액"   '59
    Query = Query & ", 전산사용료"   '59
    
    Query = Query & ") VALUES ("
    Query = Query & "   '" & 가맹점정보.가맹점코드 & "'"               ' 1 가맹점코드
    Query = Query & ",  '" & 마감일자 & "'"                            ' 2 마감일자
    Query = Query & ",  " & txtCost01.Value                            ' 3 접수금액   Spread_GetData(sprGrid(0), 1, 3, False)
    Query = Query & ",  " & txtNum01.Value                             ' 4 접수수량   Spread_GetData(sprGrid(0), 1, 1, False)
    Query = Query & ",  " & txtNum02.Value                             ' 5 출고수량   Spread_GetData(sprGrid(0), 1, 1, False)
    Query = Query & ",  " & txtNum11.Value                             ' 6 반품수량   Spread_GetData(sprGrid(0), 16, 1, False)
    Query = Query & ",  " & txtNum07.Value                             ' 7 재세탁수량 Spread_GetData(sprGrid(0), 12, 1, False)
    Query = Query & ",  " & txtCost11.Value                            ' 8 수선금액   Spread_GetData(sprGrid(0), 11, 3, False)
    Query = Query & ",  " & txtNum06.Value                             ' 9 수선수량
    Query = Query & ", '" & chkSale & "'"                              '10 판매구분
    Query = Query & ", '" & Replace(pnlTAG(0).Caption, "-", "") & "'"  '11 시작택번호 Replace(Spread_GetData(sprGrid(1), 11, 1, True), "-", "")
    Query = Query & ", '" & Replace(pnlTAG(1).Caption, "-", "") & "'"  '12 종료택번호 Replace(Spread_GetData(sprGrid(1), 12, 1, True), "-", "")
    Query = Query & ",  " & txtCost21.Value                            '13 쿠폰금액   Spread_GetData(sprGrid(1), 4, 1, False)
    Query = Query & ",  " & txtNum12.Value                             '14 쿠폰건수   Spread_GetData(sprGrid(1), 5, 1, False)
    Query = Query & ",  " & txtCost18.Value                            '15 발생마일리지 Spread_GetData(sprGrid(1), 1, 1, False)
    Query = Query & ",  " & txtCost19.Value                            '16 사용마일리지 Spread_GetData(sprGrid(1), 2, 1, False)
    Query = Query & ",  " & txtCost20.Value                            '17 삭제마일리지 Spread_GetData(sprGrid(1), 3, 1, False)
    Query = Query & ",  " & txtCost02.Value                            '18 현금입금 Spread_GetData(sprGrid(0), 3, 3, False)
    Query = Query & ",  " & txtCost03.Value                            '19 카드금액 Spread_GetData(sprGrid(0), 4, 3, False)
    Query = Query & ",  " & txtNum03.Value                             '20 카드건수 Spread_GetData(sprGrid(0), 4, 1, False)
    Query = Query & ",  " & txtCost23.Value                            '21 반품환불금액 Spread_GetData(sprGrid(0), 8, 3, False)
    Query = Query & ",  " & txtNum14.Value                             '22 반품환불건수 Spread_GetData(sprGrid(0), 8, 1, False)
    Query = Query & ",  " & txtCost24.Value                            '23 세탁환불금액 Spread_GetData(sprGrid(0), 9, 3, False)
    Query = Query & ",  " & txtNum15.Value                             '24 세탁환불건수 Spread_GetData(sprGrid(0), 9, 1, False)
    Query = Query & ",  " & txtCost26.Value                            '25 삼성카드할인금액   Spread_GetData(sprGrid(1), 14, 1, False)
    Query = Query & ",  " & txtNum17.Value                             '26 삼성카드할인건수   Spread_GetData(sprGrid(1), 15, 1, False)
    Query = Query & ",  " & txtNum18.Value                             '27 삼성카드할인고객수 Spread_GetData(sprGrid(1), 16, 1, False)
    Query = Query & ", '" & cboManager.Text & "'"                      '28 근무자명
    Query = Query & ",  " & txtCost10.Value                            '29 지사금액           Spread_GetData(sprGrid(0), 6, 3, False)
    Query = Query & ",  " & txtCost09.Value                            '30 가맹점금액 Spread_GetData(sprGrid(0), 5, 3, False)
    Query = Query & ",  " & txtCost13.Value                            '31 운동화금액 Spread_GetData(sprGrid(0), 13, 3, False)
    Query = Query & ",  " & txtNum08.Value                             '32 운동화건수 Spread_GetData(sprGrid(0), 13, 1, False)
    Query = Query & ",  0"                                             '33 운동화비율
    Query = Query & ",  " & txtCost15.Value                            '34 카페트금액 Spread_GetData(sprGrid(0), 15, 3, False)
    Query = Query & ",  " & txtNum10.Value                             '35 카페트건수 Spread_GetData(sprGrid(0), 15, 1, False)
    Query = Query & ",  0"                                             '36 명품세탁금액
    Query = Query & ",  0"                                             '37 명품세탁건수
    Query = Query & ",  0"                                             '38 명품세탁비율
    Query = Query & ",  0"                                             '39 명품염색금액
    Query = Query & ",  0"                                             '40 명품염색건수
    Query = Query & ",  0"                                             '41 명품염색비율
    Query = Query & ", 'Y'"                                            '42 마감여부
    Query = Query & ", 'N'"                                            '43 전송여부
    Query = Query & ", '" & 가맹점정보.지사코드 & "'"                  '44 지사코드
    Query = Query & ", '" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "'"   '45 마감시간
    
    Query = Query & ", '" & 가맹점정보.로열티여부1 & 가맹점정보.로열티비율1 & "' "          '46 로열티정보1
    Query = Query & ", '" & 가맹점정보.로열티여부2 & 가맹점정보.로열티비율2 & "' "          '47 로열티정보2
    Query = Query & ", '" & 가맹점정보.수수료지원여부 & 가맹점정보.수수료지원비율 & "' "    '48 수수료정보
    
    Query = Query & ",  " & Replace(txtCost29.Text, ",", "")                            '49 반품환불지사금액
    Query = Query & ",  " & Replace(txtCost30.Text, ",", "")                            '50 세탁환불지사금액
    
    Query = Query & ",  " & Replace(txtCard(4).Text, ",", "")                           '51 카드취소금액
    Query = Query & ",  " & Replace(txtCard(3).Text, ",", "")                           '52 카드취소건수
    
    Query = Query & ",  " & Replace(txtMaster(7).Text, ",", "")                        '53 로열티금액1
    Query = Query & ",  " & Replace(txtMaster(3).Text, ",", "")                        '54 로열티금액2
    Query = Query & ",  " & Replace(txtMaster(1).Text, ",", "")                        '55 수수료승인금액
    Query = Query & ",  " & Replace(txtMaster(2).Text, ",", "")                        '56 수수료취소금액
    
    Query = Query & ",  " & Replace(txtNum04.Text, ",", "")                         '57
    Query = Query & ",  " & Replace(txtCost06.Text, ",", "")                        '58
    Query = Query & ",  " & Replace(txtCost05.Text, ",", "")                        '59
    Query = Query & ",  " & Replace(txtMaster(9), ",", "")                        '60 전산사용료
    Query = Query & ")"

    ADOCon.Execute Query
    
    '---------------------------------------------------------------------------------
    ' TB_근무현황
    '---------------------------------------------------------------------------------
    Query = "UPDATE TB_근무현황 SET 종료일자 = '" & Format(Date, "YYYY-MM-DD") & "'"
    Query = Query & "             , 종료시간 = '" & Format(Now, "hh:mm:ss") & "'"
    Query = Query & "             , 업무마감 = 'Y'"
    Query = Query & " WHERE 근무자명 = '" & cboManager.Text & "'"
    Query = Query & "   AND 시작일자 = '" & Format(Date, "YYYY-MM-DD") & "'"
    Query = Query & "   AND 종료일자 = ''"
    ADOCon.Execute Query
    
    Call 일일마감_Send '마감정보 서버로 전송
    
    MsgBox "> 일일 마감이 완료 되었습니다. <", vbInformation, "일일마감"
    
    pnlDate.Enabled = True
    pnlManager.Enabled = True
    cmdFinish.Enabled = True
'
' 매출 현황을 출력해야 되는데 자꾸 종료가 되어 재실행해야 하는 문제가 있어서..

'    ' 당일 마감일 경우 프로그램을 종료 한다.
'    If 마감일자 = Format(Date, "yyyy-MM-dd") Then
'        MsgBox "> 프로그램을 종료 하겠습니다.. <", vbInformation, "일일마감"
'        End
'    End If
'
    pnlProg.Visible = False
    Unload Me
    
    Exit Sub
    
ErrRtn:
    pnlDate.Enabled = False
    pnlManager.Enabled = True
    cmdFinish.Enabled = True
    
    pnlProg.Visible = False
 
    Screen.MousePointer = 0
    
    Call Error_Msg("", Err.Source, Err.Number, Err.description & Query)
End Sub

Private Sub 일일마감_Send()
    
    Dim SSQL        As String
    Dim sValue(58)  As String
    
    On Error GoTo ERR_RTN
    
    If Server_Connection(HostCon, "LAUNDRY1000") = False Then Exit Sub
    
    SSQL = "SELECT * FROM TB_일일마감"
    SSQL = SSQL & " WHERE 가맹점코드 = '" & 가맹점정보.가맹점코드 & "'"
    SSQL = SSQL & "   AND (본사전송여부 <> 'Y' OR 본사전송여부 IS NULL)"
    SSQL = SSQL & " ORDER BY 마감일자 ASC "
    
    Set SUBRs = New ADODB.RecordSet
    SUBRs.Open SSQL, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    Do Until SUBRs.EOF
        sValue(0) = SUBRs!지사코드 & ""                         '1
        sValue(1) = SUBRs!가맹점코드 & ""                       '2
        sValue(2) = Format(SUBRs!마감일자, "YYYY-MM-DD") & ""   '3
        sValue(3) = SUBRs!접수금액 & ""                         '4
        sValue(4) = SUBRs!접수수량 & ""                         '5
        sValue(5) = SUBRs!출고수량 & ""                         '6
        sValue(6) = SUBRs!반품수량 & ""                         '7
        sValue(7) = SUBRs!재세탁수량 & ""                       '8
        sValue(8) = SUBRs!수선금액 & ""                         '9
        sValue(9) = SUBRs!수선수량 & ""                         '10
        sValue(10) = SUBRs!판매구분 & ""                        '11
        sValue(11) = SUBRs!시작택번호 & ""                      '12
        sValue(12) = SUBRs!종료택번호 & ""                      '13
        sValue(13) = SUBRs!쿠폰금액 & ""                        '14
        sValue(14) = SUBRs!쿠폰건수 & ""                        '15
        sValue(15) = SUBRs!발생마일리지 & ""                    '16
        sValue(16) = SUBRs!사용마일리지 & ""                    '17
        sValue(17) = SUBRs!삭제마일리지 & ""                    '18
        sValue(18) = SUBRs!현금입금 & ""                        '19
        sValue(19) = SUBRs!카드금액 & ""                        '20
        sValue(20) = SUBRs!카드건수 & ""                        '21
        sValue(21) = SUBRs!반품환불금액 & ""                    '22
        sValue(22) = SUBRs!반품환불건수 & ""                    '23
        sValue(23) = SUBRs!세탁환불금액 & ""                    '24
        sValue(24) = SUBRs!세탁환불건수 & ""                    '25
        sValue(25) = SUBRs!삼성카드할인금액 & ""                '26
        sValue(26) = SUBRs!삼성카드할인건수 & ""                '27
        sValue(27) = SUBRs!삼성카드할인고객수 & ""              '28
        sValue(28) = SUBRs!근무자명 & ""                        '29
        sValue(29) = SUBRs!지사금액 & ""                        '30
        sValue(30) = SUBRs!가맹점금액 & ""                      '31
        sValue(31) = SUBRs!운동화금액 & ""                      '32
        sValue(32) = SUBRs!운동화건수 & ""                      '33
        sValue(33) = SUBRs!운동화비율 & ""                      '34
        sValue(34) = SUBRs!카페트금액 & ""                      '35
        sValue(35) = SUBRs!카페트건수 & ""                      '36
        sValue(36) = SUBRs!명품세탁금액 & ""                    '37
        sValue(37) = SUBRs!명품세탁건수 & ""                    '38
        sValue(38) = SUBRs!명품세탁비율 & ""                    '39
        sValue(39) = SUBRs!명품염색금액 & ""                    '40
        sValue(40) = SUBRs!명품염색건수 & ""                    '41
        sValue(41) = SUBRs!명품염색비율 & ""                    '42
        
        sValue(42) = SUBRs!로열티정보1 & ""                    '42
        sValue(43) = SUBRs!로열티정보2 & ""                    '42
        sValue(44) = SUBRs!수수료정보 & ""                    '42
        sValue(45) = SUBRs!반품환불지사금액 & ""                    '42
        sValue(46) = SUBRs!세탁환불지사금액 & ""                    '42
        sValue(47) = SUBRs!카드취소금액 & ""                    '42
        sValue(48) = SUBRs!카드취소건수 & ""                    '42
        sValue(49) = SUBRs!로열티금액1 & ""                    '42
        
        sValue(50) = SUBRs!로열티금액2 & ""                    '42
        sValue(51) = SUBRs!수수료승인금액 & ""                    '42
        sValue(52) = SUBRs!수수료취소금액 & ""                    '42
        
        sValue(53) = SUBRs!미수카드건수 & ""                    '42
        sValue(54) = SUBRs!미수카드금액 & ""                    '42
        sValue(55) = SUBRs!미수현금수금금액 & ""                    '42
        
        sValue(56) = SUBRs!마감여부 & ""                        '43
        sValue(57) = ""                                         '44
        sValue(58) = SUBRs!전산사용료                                         '44
        
        SSQL = "EXEC SP_SE_00003_INS_NEW3"
        SSQL = SSQL & "  '" & sValue(0) & "'"  '
        SSQL = SSQL & ", '" & sValue(1) & "'"  '
        SSQL = SSQL & ", '" & sValue(2) & "'"  '
        SSQL = SSQL & ", '" & sValue(3) & "'"  '
        SSQL = SSQL & ", '" & sValue(4) & "'"  '
        SSQL = SSQL & ", '" & sValue(5) & "'"  '
        SSQL = SSQL & ", '" & sValue(6) & "'"  '
        SSQL = SSQL & ", '" & sValue(7) & "'"  '
        SSQL = SSQL & ", '" & sValue(8) & "'"  '
        SSQL = SSQL & ", '" & sValue(9) & "'"  '
        SSQL = SSQL & ", '" & sValue(10) & "'"  '
        SSQL = SSQL & ", '" & sValue(11) & "'"  '
        SSQL = SSQL & ", '" & sValue(12) & "'"  '
        SSQL = SSQL & ", '" & sValue(13) & "'"  '
        SSQL = SSQL & ", '" & sValue(14) & "'"  '
        SSQL = SSQL & ", '" & sValue(15) & "'"  '
        SSQL = SSQL & ", '" & sValue(16) & "'"  '
        SSQL = SSQL & ", '" & sValue(17) & "'"  '
        SSQL = SSQL & ", '" & sValue(18) & "'"  '
        SSQL = SSQL & ", '" & sValue(19) & "'"  '
        SSQL = SSQL & ", '" & sValue(20) & "'"  '
        SSQL = SSQL & ", '" & sValue(21) & "'"  '
        SSQL = SSQL & ", '" & sValue(22) & "'"  '
        SSQL = SSQL & ", '" & sValue(23) & "'"  '
        SSQL = SSQL & ", '" & sValue(24) & "'"  '
        SSQL = SSQL & ", '" & sValue(25) & "'"  '
        SSQL = SSQL & ", '" & sValue(26) & "'"  '
        SSQL = SSQL & ", '" & sValue(27) & "'"  '
        SSQL = SSQL & ", '" & sValue(28) & "'"  '
        SSQL = SSQL & ", '" & sValue(29) & "'"  '
        SSQL = SSQL & ", '" & sValue(30) & "'"  '
        SSQL = SSQL & ", '" & sValue(31) & "'"  '
        SSQL = SSQL & ", '" & sValue(32) & "'"  '
        SSQL = SSQL & ", '" & sValue(33) & "'"  '
        SSQL = SSQL & ", '" & sValue(34) & "'"  '
        SSQL = SSQL & ", '" & sValue(35) & "'"  '
        SSQL = SSQL & ", '" & sValue(36) & "'"  '
        SSQL = SSQL & ", '" & sValue(37) & "'"  '
        SSQL = SSQL & ", '" & sValue(38) & "'"  '
        SSQL = SSQL & ", '" & sValue(39) & "'"  '
        SSQL = SSQL & ", '" & sValue(40) & "'"  '
        SSQL = SSQL & ", '" & sValue(41) & "'"  '
        SSQL = SSQL & ", '" & sValue(42) & "'"  '
        SSQL = SSQL & ", '" & sValue(43) & "'"  '
        SSQL = SSQL & ", '" & sValue(44) & "'"  '
        SSQL = SSQL & ", '" & sValue(45) & "'"  '
        SSQL = SSQL & ", '" & sValue(46) & "'"  '
        SSQL = SSQL & ", '" & sValue(47) & "'"  '
        SSQL = SSQL & ", '" & sValue(48) & "'"  '
        SSQL = SSQL & ", '" & sValue(49) & "'"  '
        SSQL = SSQL & ", '" & sValue(50) & "'"  '
        SSQL = SSQL & ", '" & sValue(51) & "'"  '
        SSQL = SSQL & ", '" & sValue(52) & "'"  '
        SSQL = SSQL & ", '" & sValue(53) & "'"  '
        SSQL = SSQL & ", '" & sValue(54) & "'"  '
        SSQL = SSQL & ", '" & sValue(55) & "'"  '
        SSQL = SSQL & ", '" & sValue(56) & "'"  '
        SSQL = SSQL & ", '" & sValue(57) & "'"  '
        SSQL = SSQL & ", '" & sValue(58) & "'"  '
        
        HostCon.Execute SSQL
            
        '----------------------------------------------------------
        ' 일일정산 Update
        '----------------------------------------------------------
        SSQL = "UPDATE TB_일일마감 SET 본사전송여부 = 'Y'"
        SSQL = SSQL & " WHERE 마감일자   = '" & Format(SUBRs!마감일자, "YYYY-MM-DD") & "'"
        SSQL = SSQL & "   AND 가맹점코드 = '" & SUBRs!가맹점코드 & "'"
        ADOCon.Execute SSQL
        
        SUBRs.MoveNext
    Loop
    SUBRs.Close:    Set SUBRs = Nothing
    Exit Sub
    
ERR_RTN:
    Call Error_Msg("", Err.Source, Err.Number, Err.description * SSQL)

End Sub

Private Sub Control_Visible()
    cmdFinish.Enabled = Not cmdFinish.Enabled
    
    'lblTitle(0).Visible = Not lblTitle(0).Visible
    'lblTitle(1).Visible = Not lblTitle(1).Visible
    
    dtpDay.Visible = Not dtpDay.Visible
    cboManager.Visible = Not cboManager.Visible
    
    pnlProg.Left = 120
    pnlProg.Top = 135
    pnlProg.Visible = Not pnlProg.Visible
    
    DoEvents
End Sub

Private Sub cmdList_Click()
    On Error GoTo ErrRtn
    
    Query = "SELECT * FROM TB_일일마감"
    Query = Query & " WHERE (마감일자 >= '" & Format(dtpDate(0).Value, "YYYY-MM-DD") & "'"
    Query = Query & "   AND  마감일자 <= '" & Format(dtpDate(1).Value, "YYYY-MM-DD") & "')"
    Query = Query & " ORDER BY 마감일자 ASC"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows

            .Col = 1:  .Text = ADORs!마감일자 & ""
            .Col = 2:  .Text = ADORs!접수수량 & ""
            .Col = 3:  .Text = ADORs!접수금액 & ""
            
            .Col = 4: .Text = ADORs!지사금액 & ""
            .Col = 5: .Text = ADORs!가맹점금액 & ""
            
            .Col = 6:  .Text = ADORs!출고수량 & ""
            .Col = 7:  .Text = ADORs!반품수량 & ""
            .Col = 8:  .Text = ADORs!재세탁수량 & ""
            
            .Col = 9:  .Text = ADORs!수선수량 & ""
            .Col = 10: .Text = ADORs!수선금액 & ""
            
            .Col = 11:  .Text = ADORs!시작택번호 & ""
            .Col = 12: .Text = ADORs!종료택번호 & ""
            .Col = 13: .Text = ADORs!쿠폰건수 & ""
            .Col = 14: .Text = ADORs!쿠폰금액 & ""
            
            .Col = 15: .Text = ADORs!발생마일리지 & ""
            .Col = 16: .Text = ADORs!사용마일리지 & ""
            .Col = 17: .Text = ADORs!삭제마일리지 & ""
            
            .Col = 18: .Text = ADORs!현금입금 & ""
            .Col = 19: .Text = ADORs!카드건수 & ""
            .Col = 20: .Text = ADORs!카드금액 & ""
            
            .Col = 21: .Text = ADORs!반품환불건수 & ""
            .Col = 22: .Text = ADORs!반품환불금액 & ""
            
            .Col = 23: .Text = ADORs!세탁환불건수 & ""
            .Col = 24: .Text = ADORs!세탁환불금액 & ""
            
            .Col = 25: .Text = ADORs!운동화건수 & ""
            .Col = 26: .Text = ADORs!운동화금액 & ""
            
            .Col = 27: .Text = ADORs!카페트건수 & ""
            .Col = 28: .Text = ADORs!카페트금액 & ""
            
            .Col = 29: .Text = ADORs!명품세탁건수 & ""
            .Col = 30: .Text = ADORs!명품세탁금액 & ""
            
            .Col = 31: .Text = ADORs!명품염색건수 & ""
            .Col = 32: .Text = ADORs!명품염색금액 & ""
            
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
            .BackColor = &HC0E0FF
            .ForeColor = vbRed
            .BlockMode = False
            
            .Col = 1:  .Text = "합계"
            .Col = 2:  .Formula = "SUM(B1:B" & .MaxRows - 1 & ")"
            .Col = 3:  .Formula = "SUM(C1:C" & .MaxRows - 1 & ")"
            .Col = 4:  .Formula = "SUM(D1:D" & .MaxRows - 1 & ")"
            .Col = 5:  .Formula = "SUM(E1:E" & .MaxRows - 1 & ")"
            .Col = 6:  .Formula = "SUM(F1:F" & .MaxRows - 1 & ")"
            .Col = 7:  .Formula = "SUM(G1:G" & .MaxRows - 1 & ")"
            .Col = 8:  .Formula = "SUM(H1:H" & .MaxRows - 1 & ")"
            .Col = 9:  .Formula = "SUM(I1:I" & .MaxRows - 1 & ")"
            .Col = 10: .Formula = "SUM(J1:J" & .MaxRows - 1 & ")"
            
            '.Col = 11: .Formula = "SUM(K1:K" & .MaxRows - 1 & ")"
            '.Col = 12: .Formula = "SUM(L1:L" & .MaxRows - 1 & ")"
            
            .Col = 13: .Formula = "SUM(M1:M" & .MaxRows - 1 & ")"
            .Col = 14: .Formula = "SUM(N1:N" & .MaxRows - 1 & ")"
            .Col = 15: .Formula = "SUM(O1:O" & .MaxRows - 1 & ")"
            .Col = 16: .Formula = "SUM(P1:P" & .MaxRows - 1 & ")"
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
            .Col = 32: .Formula = "SUM(AF1:AF" & .MaxRows - 1 & ")"
        End If
        
        .ReDraw = True
    End With
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub dtpDay_Change()
    DoEvents
    
    Call 일일마감_Proc
    
    Call 접수집계_Display
End Sub

Private Sub Form_Activate()

    If m_Activate = True Then Exit Sub
    m_Activate = True
    
    If Not nDayCloseChk Then
        '이전소스 2010-05-04
        'dtpDay.Value = Format(m_strDayClose, "YYYY-MM-DD")
        
        If m_strDayClose = "" Then
            dtpDay.Value = Format(Date, "YYYY-MM-DD")
        Else
            dtpDay.Value = Format(m_strDayClose, "YYYY-MM-DD")
        End If
    Else
        dtpDay.Value = Format(Date, "YYYY-MM-DD")
    End If
                            
    Call Manager_Display(cboManager)
    
    TabControl1.SelectedItem = 0
    If strManager <> "" Then cboManager.Text = strManager & ""
    
    
    
    Call 일일마감_Proc
    Call 접수집계_Display
    
    DoEvents
    cmdFinish.Enabled = True


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyChk (KeyCode)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    m_Activate = False
    Me.Top = ((frmMain.Height - Me.Height) / 2) + frmMain.Top
    Me.Left = ((frmMain.Width - Me.Width) / 2) + frmMain.Left
    
    With sprGrid
        .MaxRows = 0
        .RowHeight(-1) = 14
        
        .ColsFrozen = 1
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeExtended
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
    
    With sprCloth
        .MaxRows = 0
        .RowHeight(-1) = 14
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeExtended
        
        'Init the User Sort
        '.UserColAction = UserColActionSort
    End With
    
    dtpDate(0).Value = Format(Date, "YYYY-MM-DD")
    dtpDate(1).Value = Format(Date, "YYYY-MM-DD")
    
    pnlProg.ZOrder 0

    
End Sub

Private Sub Form_Resize()
    'On Error Resume Next
    
End Sub

'-----------------------------------------------------------
'+  기간할인    1
'+  목요세일    2
'+  정상        3
'-----------------------------------------------------------
Private Sub Sale_Check()
    Dim chkWeekDay As Integer
    
    '-----------------------------------------------------------
    ' TB_할인정보
    '-----------------------------------------------------------
    Query = "SELECT * FROM TB_할인정보"
    Query = Query & " WHERE 시작일자 <= '" & 마감일자 & "' "
    Query = Query & "   AND 종료일자 >= '" & 마감일자 & "' "
    Set SUBRs = New ADODB.RecordSet
    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If Not SUBRs.EOF Then
        SUBRs.Close
        Set SUBRs = Nothing
        
        chkSale = "1"       ' 기간할인
        
        Exit Sub
    End If
    SUBRs.Close
    Set SUBRs = Nothing
    
    '-----------------------------------------------------------
    ' TB_기본정보
    '-----------------------------------------------------------
    Query = "SELECT 요일할인 FROM TB_기본정보"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    If Not ADORs.EOF Then
        i = Weekday(마감일자)
        
        If Mid(ADORs(0), i, 1) = "1" Then
            chkSale = "2"      ' 요일할인
        Else
            chkSale = "3"      ' 정상
        End If
    End If
    ADORs.Close
    Set ADORs = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_Activate = False
End Sub

Private Sub pnlMsg_DblClick()
    Dim sPass    As String
    sPass = InputBox("암호입력", "암호")
    If sPass = "isn" Or sPass = "dudtjsgh" Or sPass = "shop500" Then
        chkDate.Visible = Not chkDate.Visible         '
        btnDelete.Visible = Not btnDelete.Visible     '
        btnEditDate.Visible = Not btnEditDate.Visible '
        PushButton1.Visible = Not PushButton1.Visible '
    End If
End Sub

Private Sub 접수집계_Display()
    On Error GoTo ErrRtn
    
    Query = "SELECT   SUBSTRING(A.의류코드,1,2) AS 의류분류코드"
    Query = Query & ", B.의류분류명"
    Query = Query & ", COUNT(A.택번호) AS 수량"
    Query = Query & ", SUM(A.금액) AS 금액"
    Query = Query & ", B.세탁마진"
    Query = Query & ", ISNULL(SUM(A.금액*B.세탁마진/100),0) as 가맹점마진"
    
    Query = Query & " FROM TB_입출고 AS A LEFT OUTER JOIN TB_의류분류 AS B ON SUBSTRING(A.의류코드,1,2) = B.의류분류코드"
    Query = Query & " WHERE A.접수일자 = '" & Format(dtpDay.Value, "YYYY-MM-DD") & "'"
    Query = Query & "   AND A.판매취소 <> 'Y'"
    Query = Query & " GROUP BY SUBSTRING(A.의류코드,1,2), B.의류분류명, B.세탁마진 "
    Query = Query & " ORDER BY 수량 DESC"
    
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    With sprCloth
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows

            .Col = 1: .Text = ADORs!의류분류코드 & ""
            .Col = 2: .Text = ADORs!의류분류명 & ""
            .Col = 3: .Text = ADORs!수량 & ""
            .Col = 4: .Text = ADORs!금액 & ""
            .Col = 5: .Text = ADORs!세탁마진 & ""
            .Col = 6: .Text = ADORs!가맹점마진 & ""
            
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
            .BackColor = &HC0E0FF
            .ForeColor = vbRed
            .BlockMode = False
            
            .Col = 2:  .Text = "합계"
            .Col = 3:  .Formula = "SUM(C1:C" & .MaxRows - 1 & ")"
            .Col = 4:  .Formula = "SUM(D1:D" & .MaxRows - 1 & ")"
            
            .Col = 6:  .Formula = "SUM(F1:F" & .MaxRows - 1 & ")"
        End If
'
'        .SortKey(1) = 3
'        .SortKeyOrder(1) = SortKeyOrderDescending
'        .Sort -1, -1, -1, -1, SortByRow

        .ReDraw = True
    End With
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub PushButton1_Click()
    dtpDay_Change
End Sub



Private Sub 일일마감_Proc()
    Dim 접수금액   As Long
    Dim 가맹점마진 As Long
    Dim 외주마진   As Long

    Dim 미수금액 As Long
    Dim tmpData  As String

    On Error GoTo ErrRtn

    Screen.MousePointer = 11

    pnlProg.Left = 90
    pnlProg.Top = 1305
    pnlProg.Visible = True
    DoEvents


    '컨트롤 초기화
    Dim ctrl As Control
    Dim txt  As sidbEdit

    For Each ctrl In Me.Controls
        If TypeOf ctrl Is sidbEdit Then
            ctrl.Value = 0
        End If
    Next ctrl
    
    cboCancel.Clear
    cboReturn.Clear
    cboRepay.Clear
    cboMissTag.Clear


    마감일자 = Format(dtpDay.Value, "YYYY-MM-DD")
    
    Debug.Print "1:  " & Now()

    '----------------------------------------------------------------
    ' 1. 매출 -> 1-1) 접수수량, 접수금액 구하기
    '----------------------------------------------------------------
    Query = "SELECT    ISNULL(COUNT(택번호),0)"
    Query = Query & ", ISNULL(SUM(금액),0)"
    Query = Query & " FROM TB_입출고"
    Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
    Query = Query & "   AND (판매취소 <> 'Y') and 고객코드 < 900000"

    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    txtNum01.Value = ADORs(0)  '접수건수
    txtCost01.Value = ADORs(1) '접수금액

    ADORs.Close:    Set ADORs = Nothing


        '----------------------------------------------------------------
    ' 1. 인터넷 매출 -> 1-1) 접수수량, 접수금액 구하기
    '----------------------------------------------------------------
    Query = "SELECT    ISNULL(COUNT(택번호),0)"
    Query = Query & ", ISNULL(SUM(금액),0)"
    Query = Query & " FROM TB_입출고"
    Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
    Query = Query & "   AND (판매취소 <> 'Y') and 고객코드 > 900000"

    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    txtNum_Internet.Value = ADORs(0)  '접수건수
    txtCost_Internet.Value = ADORs(1) '접수금액

    ADORs.Close:    Set ADORs = Nothing
    
    '----------------------------------------------------------------
    ' 1. 매출 -> 1-2) 출고수량 구하기
    '----------------------------------------------------------------
    Query = "SELECT    ISNULL(COUNT(택번호),0)"
    Query = Query & " FROM TB_입출고"
    Query = Query & " WHERE 출고일자 = '" & 마감일자 & "'"
    Query = Query & "   AND (판매취소 <> 'Y')"
    Query = Query & "   AND 고객코드 < 900000"

    txtNum02.Value = Recordset_Result(Query) '출고수량


    '----------------------------------------------------------------
    ' 2. 선불결제 2-1) 현금반환/ 현금결제 구하기
    '----------------------------------------------------------------
    Query = "SELECT    ISNULL(SUM(접수금액),0) * -1"
    Query = Query & " FROM TB_매출"
    Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
    Query = Query & "   AND 적요 LIKE '%현금반환%' "
    Query = Query & "   AND 고객코드 < 900000"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    txtCost28.Value = ADORs(0)  ' 금액
    ADORs.Close:    Set ADORs = Nothing

    '----------------------------------------------------------------
    ' 2. 선불결제 2-2) 현금결제 구하기
    '----------------------------------------------------------------
    '매출중에 접수일자에 발생한 건만 처리
    ' 판매 취소한 내역도 빼주어야 한다. ( 현금입금에 -가 들어 가기 때문에 바로 처리가 가능하다.)
    Query = "SELECT ISNULL(SUM(현금입금),0)"
    Query = Query & " FROM TB_매출 "
    Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
    Query = Query & "   AND 접수금액 <> 0"
    ' Query = Query & "   AND NOT 적요 LIKE '%판매취소%' "
    Query = Query & "   AND NOT 적요 LIKE '%미수금액 입금%'"
    Query = Query & "   AND NOT 적요 LIKE '%반품환불%'"
    Query = Query & "   AND NOT 적요 LIKE '%세탁환불%'"
    Query = Query & "   AND 고객코드 < 900000"

    txtCost02.Value = Recordset_Result(Query) '

    '----------------------------------------------------------------
    ' 2. 선불결제 2-3) 카드결제 건수 / 금액 구하기
    ' 건수= 승인 + 취소 , 금액 = 승인 + 취소
    '----------------------------------------------------------------
    Query = "SELECT    ISNULL(COUNT(카드입금),0)"
    Query = Query & ", ISNULL(SUM(카드입금),0)"
    Query = Query & " FROM TB_매출 "
    Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
    Query = Query & "   AND 카드입금 <> 0" ' 카드 금결제가 아닌 경우도 0원이 들어간다.
'   Query = Query & "   AND 접수금액 > 0"
' 현금 반환 판매취소가 이쪽에 표시가 되기 때문에 표시해주어야 한다. 판매 취소후 카드로 다시 승인한 경우에 [판매취소 입금]으로 표시된다.
    Query = Query & "   AND( NOT 적요 LIKE '%판매취소%'    or 적요 like '%판매취소 입금%') "
    Query = Query & "   AND NOT 적요 LIKE '%미수금액 입금%'"
    Query = Query & "   AND NOT 적요 LIKE '%반품환불%'"
    Query = Query & "   AND NOT 적요 LIKE '%세탁환불%'"
    Query = Query & "   AND 고객코드 < 900000"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    txtNum03.Value = ADORs(0)
    txtCost03.Value = ADORs(1)

    ADORs.Close:    Set ADORs = Nothing
'
'    ' 판매 취소한 내역도 빼주어야 한다. ( 카드입금에 -금액이 들어가지 않고 승인취소 전표에 -금액을 일괄 적용한다.)
'    ' 그렇기 때문에 카드 판매 취소를 구하기가 힘들다 ㅡㅡ
'
'    ' 건수 구하기
'    Query = "SELECT    ISNULL(COUNT(A.접수번호),0) FROM "
'    Query = Query & " ( SELECT 접수번호 "
'    Query = Query & " FROM TB_매출 "
'    Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND 카드입금 = 0"
'    Query = Query & "   AND 접수금액 < 0"
'    Query = Query & "   AND 현금입금 = 0"
'    Query = Query & "   AND 적요 LIKE '%판매취소%' "
'    Query = Query & " GROUP BY 접수번호 ) A "
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    txtNum03.Value = txtNum03.Value - Recordset_Result(Query)
'
'    ' 금액 구하기
'    Query = "SELECT    ISNULL(SUM(접수금액),0)"
'    Query = Query & " FROM TB_매출 "
'    Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND 카드입금 = 0"
'    Query = Query & "   AND 접수금액 < 0"
'    Query = Query & "   AND 현금입금 = 0"
'    Query = Query & "   AND 적요 LIKE '%판매취소%' "
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    txtCost03.Value = txtCost03.Value + Recordset_Result(Query) '- 금액이 날라오기 때문에 더해준다.
    
    '--------------------------------------------------------------------
    ' 2. 선불결제 2-4) 발생/사용/삭제 마일리지
    '--------------------------------------------------------------------
    Query = "SELECT    ISNULL(SUM(발생마일리지),0)"
    Query = Query & ", ISNULL(SUM(사용마일리지),0)"
    Query = Query & ", ISNULL(SUM(삭제마일리지),0)"
    Query = Query & " FROM TB_매출"
    Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
    Query = Query & "   AND 고객코드 < 900000"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    txtCost18.Value = ADORs(0) '
    txtCost19.Value = ADORs(1) '사용마일리지
    txtCost20.Value = ADORs(2) '

    txtCost27.Value = txtCost19.Value '사용마일리지

    ADORs.Close:    Set ADORs = Nothing

    '--------------------------------------------------------------------
    ' 2. 선불결제 2-5) 쿠폰 사용 건수/ 금액
    '--------------------------------------------------------------------
    Query = "SELECT    ISNULL(SUM(쿠폰입금),0)"
    Query = Query & ", ISNULL(COUNT(쿠폰번호),0)"
    Query = Query & " FROM TB_매출"
    Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
    Query = Query & "   AND 쿠폰입금 > 0"
    Query = Query & "   AND 고객코드 < 900000"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    txtNum12.Value = ADORs(1)   ' 수량
    txtCost21.Value = ADORs(0)  ' 금액

    ADORs.Close:    Set ADORs = Nothing

    '----------------------------------------------------------------
    ' 2. 선불결제 2-6) 미수금 금액
    ' 마일리지를 사용한 것을 판매취소할 경우 미수금액이 마일리지 사용한 것으로 처리되기 때문에
    ' 별도로 마일리지판매취소 금액을 구해서 -해준다.
    ' 마일리지판매취소 값이 -로 넘어오기 때문에 - 해주면 +로 처리된다.(위에서 마일리지 금액이 처리되어 나요기 때문)
    '----------------------------------------------------------------
    Query = "SELECT ISNULL(SUM(접수금액),0) - ISNULL(SUM(입금합계),0) - ISNULL(SUM(쿠폰입금),0) AS 미수금 "
    Query = Query & " FROM TB_매출"
    Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
'   Query = Query & "   AND 접수금액 > 0"
    Query = Query & "   AND NOT 적요  LIKE '%미수금액 입금%'"
    Query = Query & "   AND NOT 적요  LIKE '%반품환불%'"
    Query = Query & "   AND NOT 적요  LIKE '%세탁환불%'"
    Query = Query & "   AND NOT 적요  LIKE '%현금반환%'"  ' 현금 반환한 부분이 포함되는 것을 뺀다.
    Query = Query & "   AND 고객코드 < 900000"

    txtCost04.Value = Recordset_Result(Query) '
    
    Query = "SELECT ISNULL(SUM(사용마일리지),0)  AS 판매취소마일리지금액 "
    Query = Query & " FROM TB_매출"
    Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
    Query = Query & "   AND 접수금액 < 0"
    Query = Query & "   AND 사용마일리지 < 0"
    Query = Query & "   AND NOT 적요  LIKE '%미수금액 입금%'"
    Query = Query & "   AND NOT 적요  LIKE '%반품환불%'"
    Query = Query & "   AND NOT 적요  LIKE '%세탁환불%'"
    Query = Query & "   AND 고객코드 < 900000"
    txtCost04.Value = txtCost04.Value - Recordset_Result(Query)  '

    '----------------------------------------------------------------
    ' 3. 미수결제 3-1) 미수금 수금 현금결제 구하기
    '----------------------------------------------------------------
    Query = "SELECT ISNULL(SUM(현금입금),0)"
    Query = Query & " FROM TB_매출"
    Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
    Query = Query & "   AND 접수금액 = 0"
    Query = Query & "   AND 적요  LIKE '%미수금액 입금%'"
    Query = Query & "   AND 고객코드 < 900000"

    txtCost05.Value = Recordset_Result(Query) '

    '----------------------------------------------------------------
    ' 3. 미수결제 3-2) 미수금 수금 카드결제 구하기
    '----------------------------------------------------------------
    Query = "SELECT    ISNULL(COUNT(카드입금),0)"
    Query = Query & ", ISNULL(SUM(카드입금),0)"
    Query = Query & " FROM TB_매출"
    Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
    Query = Query & "   AND 카드입금 <> 0"
    Query = Query & "   AND 접수금액 = 0" ' 판매취소시 0원으로 들어온다.
    Query = Query & "   AND 적요  LIKE '%미수금액 입금%'"
    Query = Query & "   AND 고객코드 < 900000"

    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    txtNum04.Value = ADORs(0)
    txtCost06.Value = ADORs(1)

    ADORs.Close:    Set ADORs = Nothing

    '----------------------------------------------------------------
    ' 4. 결제합계 현금/ 카드 결제
    '----------------------------------------------------------------
    txtCost07.Value = txtCost02.Value + txtCost05.Value '현금결제합계
    txtNum05.Value = txtNum03.Value + txtNum04.Value    '카드결제건수 합계
    txtCost08.Value = txtCost03.Value + txtCost06.Value '카드결제합계

    '----------------------------------------------------------------
    ' 5. 마진 5-1) 사용마일리지  2. 선불결제 2-4)에서 처리
    '----------------------------------------------------------------
    '----------------------------------------------------------------
    ' 5. 마진 5-2) 가맹점 마진
    '----------------------------------------------------------------
    Query = "SELECT ISNULL(ROUND(SUM(금액 * 세탁마진/100.00),0),0)"
    Query = Query & " FROM TB_입출고"
    Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
    Query = Query & "   AND 내용 NOT LIKE '%수%'"                            '수선 제외
    Query = Query & "   AND (판매취소 <> 'Y')"
    Query = Query & "   AND 고객코드 < 900000"

    txtCost09.Value = Recordset_Result(Query)

    ' 마일리지 사용이 있을 경우
    If txtCost19.Value > 0 Then
        txtCost09.Value = txtCost09.Value - CLng(txtCost19.Value * 0.4) '가맹점  지사:가맹점(6:4)로 빼준다.
        'txtCost10.Value = txtCost10.Value - CLng(txtCost19.Value * 0.6) '지사
    End If

    '쿠폰 사용이 있는 경우
    If txtCost21.Value > 0 And 마감일자 <= "2011-12-31" Then
        txtCost09.Value = txtCost09.Value - CLng(1200 * txtNum12.Value * 0.4) '가맹점
        'txtCost10.Value = txtCost10.Value - CLng(1200 * txtNum12.Value * 0.6) '지사
    End If

'    '----------------------------------------------------------------
'    ' 5. 마진 5-2) 지사 마진
'    '----------------------------------------------------------------
    txtCost10.Value = (txtCost01.Value - txtCost19.Value) - txtCost09.Value                     ' 지사 마진 = (접수금액 -마일리지) - 가맹점마진
'


    '--------------------------------------------------------------
    ' 6. 기타자료 6-1) 수선수량 계산
    '--------------------------------------------------------------
    Query = "SELECT    ISNULL(COUNT(택번호),0)"
    Query = Query & ", ISNULL(SUM(수선금액),0)"
    Query = Query & " FROM TB_입출고"
    Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
    Query = Query & "   AND (내용  = '드수' OR 내용 = '수') "
    Query = Query & "   AND (판매취소 <> 'Y')"
    Query = Query & "   AND 고객코드 < 900000"

    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    txtNum06.Value = ADORs(0)
    txtCost11.Value = ADORs(1)

    ADORs.Close:    Set ADORs = Nothing

    '----------------------------------------------------------------
    ' 6. 기타자료 6-2) 재세탁수량 계산
    '----------------------------------------------------------------
    Query = "SELECT ISNULL(COUNT(택번호),0)"
    Query = Query & " FROM TB_입출고"
    Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
    Query = Query & "   AND 내용     = '드재'"
    Query = Query & "   AND (판매취소 <> 'Y')"
    Query = Query & "   AND 고객코드 < 900000"

    txtNum07.Value = Recordset_Result(Query)

    '--------------------------------------------------------------------
    ' 6. 기타자료 6-3) 운동화 매출을 불러온다.
    '--------------------------------------------------------------------
    Query = "SELECT    ISNULL(COUNT(의류코드),0)" '운동화건수
    Query = Query & ", ISNULL(SUM(금액),0)"       '운동화금액
    Query = Query & " FROM TB_입출고 "
    Query = Query & " WHERE SUBSTRING(의류코드,1,2) = 'a0'"
    Query = Query & "   AND 접수일자 = '" & 마감일자 & "'"
    Query = Query & "   AND (판매취소 <> 'Y')"
    Query = Query & "   AND 고객코드 < 900000"

    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    txtNum08.Value = ADORs(0)
    txtCost13.Value = ADORs(1)

    ADORs.Close:        Set ADORs = Nothing

    '--------------------------------------------------------------------
    ' 6. 기타자료 6-4) 가죽/무스탕 매출을 불러온다.
    '--------------------------------------------------------------------
    Query = "SELECT    ISNULL(COUNT(의류코드),0)" '가죽건수
    Query = Query & ", ISNULL(SUM(금액),0)"       '가죽금액
    Query = Query & " FROM TB_입출고 "
    Query = Query & " WHERE SUBSTRING(의류코드,1,2) IN ('b0','n0')"
    Query = Query & "   AND 접수일자 = '" & 마감일자 & "'"
    Query = Query & "   AND (판매취소 <> 'Y')"
    Query = Query & "   AND 고객코드 < 900000"

    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    txtNum09.Value = ADORs(0)
    txtCost14.Value = ADORs(1)

    ADORs.Close:    Set ADORs = Nothing

    '--------------------------------------------------------------------
    ' 6. 기타자료 6-5) 카페트 매출을 불러온다.
    '--------------------------------------------------------------------
    Query = "SELECT    ISNULL(COUNT(의류코드),0)" '카페트건수
    Query = Query & ", ISNULL(SUM(금액),0) "      '카페트금액
    Query = Query & " FROM TB_입출고 "
    Query = Query & " WHERE SUBSTRING(의류코드,1,2) = 'x0'"
    Query = Query & "   AND 접수일자 = '" & 마감일자 & "'"
    Query = Query & "   AND (판매취소 <> 'Y')"
    Query = Query & "   AND 고객코드 < 900000"

    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    txtNum10.Value = ADORs(0)
    txtCost15.Value = ADORs(1)

    ADORs.Close
    Set ADORs = Nothing

    '----------------------------------------------------------------
    ' 6. 기타자료 6-6) 반품 내역
    '----------------------------------------------------------------
    Query = "SELECT ISNULL(COUNT(택번호),0)"
    Query = Query & " FROM TB_입출고"
    Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
    Query = Query & "   AND 내용     = '%반%'"
    Query = Query & "   AND (판매취소 <> 'Y')"
    Query = Query & "   AND 고객코드 < 900000"

    txtNum11.Value = Recordset_Result(Query)

    '----------------------------------------------------------------
    ' 6. 기타자료 6-7) 외주 마진
    '----------------------------------------------------------------
    Query = "SELECT ISNULL(SUM(금액*외주마진/100),0)"
    Query = Query & " FROM TB_입출고"
    Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
    Query = Query & "   AND SUBSTRING(의류코드,1,1) = 'a'"   '운동화
    Query = Query & "   AND 내용 NOT LIKE '%수%'"            '수선 제외
    Query = Query & "   AND (판매취소 <> 'Y')"
    Query = Query & "   AND 고객코드 < 900000"

    txtCost17.Value = Recordset_Result(Query)

    '----------------------------------------------------------------
    ' 6. 기타자료 6-8) 마일리지 2. 선불결제 2-4)에서 처리
    '----------------------------------------------------------------
    '----------------------------------------------------------------
    ' 6. 기타자료 6-9) 마일리지 2. 선불결제 2-4)에서 처리
    '----------------------------------------------------------------

    '----------------------------------------------------------------
    ' 7) 기타자료2 7-1) 판매취소 내역
    '----------------------------------------------------------------
    Query = "SELECT    ISNULL(COUNT(택번호),0)"
    Query = Query & ", ISNULL(SUM(금액),0)"
    Query = Query & " FROM TB_입출고"
    Query = Query & " WHERE SUBSTRING(판매취소일자,1,10) = '" & 마감일자 & "'"
    Query = Query & "   AND 고객코드 < 900000"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    txtNum13.Value = ADORs(0)
    txtCost22.Value = ADORs(1)

    ADORs.Close:    Set ADORs = Nothing

    Query = "SELECT    택번호"
    Query = Query & " FROM TB_입출고"
    Query = Query & " WHERE SUBSTRING(판매취소일자,1,10) = '" & 마감일자 & "'"
    Query = Query & "   AND 고객코드 < 900000"
    Query = Query & " ORDER BY 택번호 ASC"

    Call Get_택번호(Query, cboCancel)
    If cboCancel.ListCount > 0 Then cboCancel.ListIndex = 0

    '----------------------------------------------------------------
    ' 7) 기타자료2 7-2) 반품환불 내역
    '----------------------------------------------------------------
    Query = "SELECT    ISNULL(COUNT(택번호),0)"
    Query = Query & ", ISNULL(SUM(금액),0)"
    Query = Query & ", ISNULL(SUM(금액*(100-세탁마진)/100),0)"
    Query = Query & " FROM TB_입출고"
    Query = Query & " WHERE SUBSTRING(반품환불일자,1,10) = '" & 마감일자 & "'"
    Query = Query & "   AND (판매취소 <> 'Y') "
    Query = Query & "   AND 고객코드 < 900000"
    
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    txtNum14.Value = ADORs(0)
    txtCost23.Value = ADORs(1)
    txtCost29.Value = ADORs(2)

    ADORs.Close:    Set ADORs = Nothing

    Query = "SELECT    택번호"
    Query = Query & " FROM TB_입출고"
    Query = Query & " WHERE SUBSTRING(반품환불일자,1,10) = '" & 마감일자 & "'"
    Query = Query & "   AND (판매취소 <> 'Y') "
    Query = Query & "   AND 고객코드 < 900000"

    Call Get_택번호(Query, cboReturn)
    If cboReturn.ListCount > 0 Then cboReturn.ListIndex = 0

    '----------------------------------------------------------------
    ' 7) 기타자료2 7-3) 세탁환불 내역
    '----------------------------------------------------------------
    Query = "SELECT    ISNULL(COUNT(택번호),0)"
    Query = Query & ", ISNULL(SUM(금액),0)"
    Query = Query & ", ISNULL(SUM(금액*(100-세탁마진)/100),0)"
    Query = Query & " FROM TB_입출고"
    Query = Query & " WHERE SUBSTRING(세탁환불일자,1,10) = '" & 마감일자 & "'"
    Query = Query & "   AND (판매취소 <> 'Y') "
    Query = Query & "   AND 고객코드 < 900000"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    txtNum15.Value = ADORs(0)
    txtCost24.Value = ADORs(1)
    txtCost30.Value = ADORs(2)

    ADORs.Close:    Set ADORs = Nothing

    Query = "SELECT    택번호"
    Query = Query & " FROM TB_입출고"
    Query = Query & " WHERE SUBSTRING(세탁환불일자,1,10) = '" & 마감일자 & "'"
    Query = Query & "   AND (판매취소 <> 'Y') "
    Query = Query & "   AND 고객코드 < 900000"
    
    Call Get_택번호(Query, cboRepay)
    If cboRepay.ListCount > 0 Then cboRepay.ListIndex = 0

    '--------------------------------------------------------------------
    ' 7) 기타자료2 7-4) 누락TAG CHECK 내역
    '--------------------------------------------------------------------
    Dim 시작택번호   As String
    Dim 마지막택번호 As String

    Dim 택번호 As String
    Dim tmpTAG As String

    Query = "SELECT    MIN(택번호)"
    Query = Query & ", MAX(택번호)"
    Query = Query & " FROM TB_입출고 "
    Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
    Query = Query & "   AND (판매취소 <> 'Y')"
    

    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    If Not ADORs.EOF Then
        시작택번호 = ADORs(0) & ""
        마지막택번호 = ADORs(1) & ""
    End If

    ADORs.Close:    Set ADORs = Nothing
    cboMissTag.Clear

    Dim iLoop As Long

    Query = "SELECT 택번호 FROM TB_입출고"
    Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
    Query = Query & "   AND (판매취소 <> 'Y')"
    
    Query = Query & " ORDER BY 택번호 ASC"

    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    iLoop = 0

    택번호 = ""
    tmpTAG = ""

    If Val(마지막택번호) - Val(시작택번호) < 5000 Then
        Do Until ADORs.EOF
            If tmpTAG = "" Then
                tmpTAG = ADORs!택번호
            Else
                Do Until Format(CLng(tmpTAG) + 1, "000000000") >= ADORs!택번호
                    cboMissTag.AddItem Format(CLng(tmpTAG) + 1, "000-00-0000")

                    tmpTAG = Format(CLng(tmpTAG) + 1, "000000000")

                    '100 개가 넘으면 빠져 나옴
                    If iLoop >= 100 Then
                        cboMissTag.AddItem "Err"

                        Exit Do
                    End If

                    iLoop = iLoop + 1
                Loop

                tmpTAG = Format(CLng(tmpTAG) + 1, "000000000")
            End If

            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing

        If cboMissTag.ListCount = 0 Then
            txtNum16.Value = 0
        Else
            txtNum16.Value = cboMissTag.ListCount - 1
        End If
    End If
    If cboMissTag.ListCount > 0 Then cboMissTag.ListIndex = 0

    pnlTAG(0).Caption = Format(시작택번호, "000-00-0000") & ""
    pnlTAG(1).Caption = Format(마지막택번호, "000-00-0000") & ""

    '--------------------------------------------------------------------
    ' 7) 기타자료2 7-4) 삼성 카드 할인 내용 추가
    '--------------------------------------------------------------------
    Dim 삼성카드고객수   As Long
    Dim 삼성카드할인건수 As Long
    Dim 삼성카드할인금액 As Long

    삼성카드고객수 = 0
    삼성카드할인건수 = 0
    삼성카드할인금액 = 0

    Query = "SELECT    고객코드"
    Query = Query & ", ISNULL(COUNT(금액),0)"
    Query = Query & ", ISNULL(SUM(금액),0)"
    Query = Query & " FROM TB_입출고"
    Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
    Query = Query & "   AND 내용  LIKE '%삼%'"
    Query = Query & "   AND (판매취소 <> 'Y')"
    Query = Query & "   AND 고객코드 < 900000"

    Query = Query & " GROUP BY 고객코드"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    Do Until ADORs.EOF
        삼성카드고객수 = 삼성카드고객수 + 1

        삼성카드할인건수 = 삼성카드할인건수 + ADORs(0)
        삼성카드할인금액 = 삼성카드할인금액 + ADORs(1)

        ADORs.MoveNext
    Loop
    ADORs.Close
    Set ADORs = Nothing

    txtCost26.Value = 삼성카드할인금액
    txtNum17.Value = 삼성카드할인건수
    txtNum18.Value = 삼성카드고객수


    '--------------------------------------------------------------------
    ' 8) 반품환불, 세탁환불 확정시 마진 처리
    '--------------------------------------------------------------------
'    If txtCost23.Tag <> "" Then
'        txtCost09.Value = txtCost09.Value - txtCost23.Tag                     '가맹점 반품환불
'
'        txtCost10.Value = txtCost10.Value - (txtCost23.Value - txtCost23.Tag) '지사   반품환불
'    End If
'
'    If txtCost24.Tag <> "" Then
'        txtCost09.Value = txtCost09.Value - txtCost24.Tag                     '가맹점 세탁환불
'
'        txtCost10.Value = txtCost10.Value - (txtCost24.Value - txtCost24.Tag) '지사   세탁환불
'    End If

    '--------------------------------------------------------------------
    ' 8) 지사 정산 참고 사항
    '--------------------------------------------------------------------
    Debug.Print Now & " 8) 지사 정산 참고 사항  1. 로열티 정보"
    
    pnlData(24).Caption = "+ 로열티 1 " & 가맹점정보.로열티여부1 & " " & 가맹점정보.로열티비율1
    If 가맹점정보.로열티여부1 = "Y" And IsNumeric(가맹점정보.로열티비율1) Then txtMaster(7).Value = CDbl(txtCost01.Value) * (CDbl(가맹점정보.로열티비율1) / 100)
        
    pnlData(25).Caption = "+ 유통로열티 " & 가맹점정보.로열티여부2 & " " & 가맹점정보.로열티비율2
    If 가맹점정보.로열티여부2 = "Y" And IsNumeric(가맹점정보.로열티비율2) Then txtMaster(3).Value = CDbl(txtCost09.Value) * (CDbl(가맹점정보.로열티비율2) / 100)

    
    pnlData(46).Caption = "수수료 지원 정보 " & 가맹점정보.수수료지원여부 & " " & 가맹점정보.수수료지원비율
    
    Query = "SELECT    ISNULL(COUNT(결제금액),0)"
    Query = Query & ", ISNULL(SUM(결제금액),0)"
    Query = Query & " FROM TB_신용카드승인 "
    Query = Query & " WHERE 승인일자 = '" & Mid(Replace(마감일자, "-", ""), 3, 6) & "'"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    
    txtCard(0).Value = ADORs.Fields(0)
    txtCard(1).Value = ADORs.Fields(1)
    If 가맹점정보.수수료지원여부 = "Y" And IsNumeric(가맹점정보.수수료지원비율) Then txtCard(2).Value = CDbl(txtCard(1).Value) * (CDbl(가맹점정보.수수료지원비율) / 100)
    ADORs.Close:    Set ADORs = Nothing
    
    '쿠폰 지사 금액
    txtMaster(8).Value = CDbl(txtCost21.Value) * (CDbl(가맹점정보.비율) / 100)
    
    Query = "SELECT    ISNULL(COUNT(결제금액),0)"
    Query = Query & ", ISNULL(SUM(결제금액),0)"
    Query = Query & " FROM TB_신용카드승인 "
    Query = Query & " WHERE substring(취소일자,21,10) = '" & 마감일자 & "'"
    
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    txtCard(3).Value = ADORs.Fields(0)
    txtCard(4).Value = ADORs.Fields(1)
    If 가맹점정보.수수료지원여부 = "Y" And IsNumeric(가맹점정보.수수료지원비율) Then
        txtCard(5).Value = CDbl(txtCard(4).Value) * (CDbl(가맹점정보.수수료지원비율) / 100)
        If txtCard(5).Text = "" Then
            txtCard(5).Value = 0
        End If
    End If
    ADORs.Close:    Set ADORs = Nothing
    
    If CInt(Format(dtpDay.Value, "DD")) > 24 Then
        Query = "select ISNULL(SUM(전산사용료),0) FROM TB_일일마감 WHERE 마감일자 like '" & Format(dtpDay.Value, "YYYY-MM") & "%'"
        Set ADORs = New ADODB.RecordSet
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        If ADORs.Fields(0) = "0" Then
            txtMaster(9).Value = 가맹점정보.전산사용료
        End If
        ADORs.Close:    Set ADORs = Nothing
    End If
        
    Query = "select isnull(전산사용료,0) FROM TB_일일마감 WHERE 마감일자 = '" & 마감일자 & "'"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    If ADORs.RecordCount > 0 Then
        If ADORs.Fields(0) <> "0" Then
            txtMaster(9).Value = ADORs.Fields(0)
        End If
    End If
    ADORs.Close:    Set ADORs = Nothing
    

    txtMaster(0).Value = txtCost10.Value
    txtMaster(1).Value = txtCard(2).Value
    txtMaster(2).Value = txtCard(5).Value
    txtMaster(4).Value = txtCost29.Value + txtCost30.Value
    ' 지사정산금액 = 지사분매출 - (카드수수료지원금+환불금액) + (카드수수료환불금 +로열티2) - 쿠폰금액의 60%
    txtMaster(5).Value = txtMaster(0).Value - (txtMaster(1).Value + txtMaster(4).Value) + (txtMaster(2).Value + txtMaster(3).Value + txtMaster(9).Value) - txtMaster(8).Value
'    txtMaster(6).Value = (txtCost09.Value + txtCost10.Value) - txtMaster(5).Value
    ' 매장 수익금 = 가맹점마진 - ((반품환불금액 + 세탁환불금액) - 세탁/반품환불금액(지산) - 카드수수료환불금 + 카드수수료지원금 - 유통로열티 - 쿠폰금액의 40%
    txtMaster(6).Value = txtCost09.Value - ((txtCost23.Value + txtCost24.Value) - txtMaster(4).Value) - txtMaster(2).Value + txtMaster(1).Value - txtMaster(3).Value - txtMaster(9).Value - (txtCost21.Value - txtMaster(8).Value)
    
    
    

    
    
    Screen.MousePointer = 0
    pnlProg.Visible = False
    DoEvents

    Debug.Print "100:  " & Now()
    Exit Sub

ErrRtn:
    
    Screen.MousePointer = 0
    pnlProg.Visible = False
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub
