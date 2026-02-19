VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm일일판매현황 
   Caption         =   "일일판매 현황"
   ClientHeight    =   11970
   ClientLeft      =   5940
   ClientTop       =   3300
   ClientWidth     =   16410
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form20"
   MDIChild        =   -1  'True
   ScaleHeight     =   11970
   ScaleWidth      =   16410
   WindowState     =   2  '최대화
   Begin Threed.SSPanel pnlProg 
      Height          =   1215
      Left            =   810
      TabIndex        =   9
      Top             =   8775
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
      Picture         =   "frm일일판매현황.frx":0000
      Caption         =   " "
      BorderWidth     =   0
      BevelOuter      =   1
      BevelInner      =   2
      Alignment       =   4
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   11970
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   16410
      _ExtentX        =   28945
      _ExtentY        =   21114
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm일일판매현황.frx":2FCB
      Begin XtremeSuiteControls.TabControl TabControl1 
         Height          =   10755
         Left            =   15
         TabIndex        =   10
         Top             =   1200
         Width           =   16380
         _Version        =   851970
         _ExtentX        =   28893
         _ExtentY        =   18971
         _StockProps     =   68
         Appearance      =   3
         Color           =   64
         PaintManager.ShowIcons=   -1  'True
         PaintManager.ButtonMargin=   "2,3,2,3"
         ItemCount       =   3
         SelectedItem    =   1
         Item(0).Caption =   " 접수현황 "
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "TabControlPage1"
         Item(1).Caption =   " 일일마감 "
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "TabControlPage2"
         Item(2).Caption =   " 접수집계 "
         Item(2).ControlCount=   1
         Item(2).Control(0)=   "TabControlPage"
         Begin XtremeSuiteControls.TabControlPage TabControlPage 
            Height          =   10275
            Left            =   -69970
            TabIndex        =   13
            Top             =   450
            Visible         =   0   'False
            Width           =   16320
            _Version        =   851970
            _ExtentX        =   28787
            _ExtentY        =   18124
            _StockProps     =   1
            Page            =   2
            Begin XtremeSuiteControls.GroupBox GroupBox1 
               Height          =   6015
               Left            =   9270
               TabIndex        =   198
               Top             =   120
               Width           =   5265
               _Version        =   851970
               _ExtentX        =   9287
               _ExtentY        =   10610
               _StockProps     =   79
               Caption         =   "기타 정보"
               UseVisualStyle  =   -1  'True
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   10
                  Left            =   300
                  TabIndex        =   199
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
                  PictureBackground=   "frm일일판매현황.frx":303D
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   11
                  Left            =   300
                  TabIndex        =   200
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
                  PictureBackground=   "frm일일판매현황.frx":3263
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   12
                  Left            =   300
                  TabIndex        =   201
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
                  PictureBackground=   "frm일일판매현황.frx":3489
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   13
                  Left            =   300
                  TabIndex        =   202
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
                  PictureBackground=   "frm일일판매현황.frx":36AF
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   14
                  Left            =   300
                  TabIndex        =   203
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
                  PictureBackground=   "frm일일판매현황.frx":38D5
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   15
                  Left            =   300
                  TabIndex        =   204
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
                  PictureBackground=   "frm일일판매현황.frx":3AFB
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   41
                  Left            =   1740
                  TabIndex        =   205
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
                     TabIndex        =   206
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "점"
                     PictureBackground=   "frm일일판매현황.frx":3D21
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtNum06 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   207
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
                     TabIndex        =   208
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
                     TabIndex        =   209
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":43EB
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   44
                  Left            =   1740
                  TabIndex        =   210
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
                     TabIndex        =   211
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "점"
                     PictureBackground=   "frm일일판매현황.frx":4AB5
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtNum07 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   212
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
                     TabIndex        =   213
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
                     TabIndex        =   214
                     Top             =   60
                     Visible         =   0   'False
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":517F
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   47
                  Left            =   1740
                  TabIndex        =   215
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
                     TabIndex        =   216
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "점"
                     PictureBackground=   "frm일일판매현황.frx":5849
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtNum08 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   217
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
                     TabIndex        =   218
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
                     TabIndex        =   219
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":5F13
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   50
                  Left            =   1740
                  TabIndex        =   220
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
                     TabIndex        =   221
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "점"
                     PictureBackground=   "frm일일판매현황.frx":65DD
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtNum09 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   222
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
                     TabIndex        =   223
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
                     TabIndex        =   224
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":6CA7
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   53
                  Left            =   1740
                  TabIndex        =   225
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
                     TabIndex        =   226
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "점"
                     PictureBackground=   "frm일일판매현황.frx":7371
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtNum10 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   227
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
                     TabIndex        =   228
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
                     TabIndex        =   229
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":7A3B
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   56
                  Left            =   1740
                  TabIndex        =   230
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
                     TabIndex        =   231
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "점"
                     PictureBackground=   "frm일일판매현황.frx":8105
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtNum11 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   232
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
                     TabIndex        =   233
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
                     TabIndex        =   234
                     Top             =   60
                     Visible         =   0   'False
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":87CF
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   20
                  Left            =   300
                  TabIndex        =   235
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
                  PictureBackground=   "frm일일판매현황.frx":8E99
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   73
                  Left            =   1740
                  TabIndex        =   236
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
                     TabIndex        =   237
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
                     TabIndex        =   238
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":90BF
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   18
                  Left            =   300
                  TabIndex        =   239
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
                  PictureBackground=   "frm일일판매현황.frx":9789
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   68
                  Left            =   1740
                  TabIndex        =   240
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
                     TabIndex        =   241
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "건"
                     PictureBackground=   "frm일일판매현황.frx":99AF
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtNum17 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   242
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
                     TabIndex        =   243
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
                     TabIndex        =   244
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":A079
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   19
                  Left            =   300
                  TabIndex        =   245
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
                  PictureBackground=   "frm일일판매현황.frx":A743
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   71
                  Left            =   1740
                  TabIndex        =   246
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
                     TabIndex        =   247
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
                     TabIndex        =   248
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "명"
                     PictureBackground=   "frm일일판매현황.frx":A969
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
            End
            Begin FPSpreadADO.fpSpread sprCloth 
               Height          =   7380
               Left            =   90
               TabIndex        =   254
               Top             =   90
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
               SpreadDesigner  =   "frm일일판매현황.frx":B033
               VisibleCols     =   3
               VisibleRows     =   30
               Appearance      =   1
               HighlightHeaders=   1
               HighlightStyle  =   1
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage2 
            Height          =   10275
            Left            =   30
            TabIndex        =   12
            Top             =   450
            Width           =   16320
            _Version        =   851970
            _ExtentX        =   28787
            _ExtentY        =   18124
            _StockProps     =   1
            BackColor       =   16777215
            Page            =   1
            Begin XtremeSuiteControls.GroupBox GroupBox 
               Height          =   6825
               Index           =   1
               Left            =   10215
               TabIndex        =   16
               Top             =   90
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
                  TabIndex        =   17
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
                     TabIndex        =   18
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "건"
                     PictureBackground=   "frm일일판매현황.frx":B73A
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtNum13 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   19
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
                     TabIndex        =   20
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
                     TabIndex        =   21
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":BE04
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   825
                  Index           =   37
                  Left            =   120
                  TabIndex        =   22
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
                  PictureBackground=   "frm일일판매현황.frx":C4CE
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   32
                  Left            =   1560
                  TabIndex        =   23
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
                     TabIndex        =   24
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "건"
                     PictureBackground=   "frm일일판매현황.frx":C6F4
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtNum14 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   25
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
                     TabIndex        =   26
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
                     TabIndex        =   27
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":CDBE
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   1215
                  Index           =   39
                  Left            =   120
                  TabIndex        =   28
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
                  PictureBackground=   "frm일일판매현황.frx":D488
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   35
                  Left            =   1560
                  TabIndex        =   29
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
                     TabIndex        =   30
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "건"
                     PictureBackground=   "frm일일판매현황.frx":D6AE
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtNum15 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   31
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
                     TabIndex        =   32
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
                     TabIndex        =   33
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":DD78
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   1215
                  Index           =   40
                  Left            =   120
                  TabIndex        =   34
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
                  PictureBackground=   "frm일일판매현황.frx":E442
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   38
                  Left            =   1560
                  TabIndex        =   35
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
                     TabIndex        =   36
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "건"
                     PictureBackground=   "frm일일판매현황.frx":E668
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtNum16 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   37
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
                     TabIndex        =   38
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
                     TabIndex        =   39
                     Top             =   60
                     Visible         =   0   'False
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":ED32
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   825
                  Index           =   9
                  Left            =   120
                  TabIndex        =   40
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
                  PictureBackground=   "frm일일판매현황.frx":F3FC
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel4 
                  Height          =   420
                  Left            =   1560
                  TabIndex        =   41
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
                     TabIndex        =   42
                     Top             =   60
                     Width           =   2955
                  End
               End
               Begin Threed.SSPanel SSPanel3 
                  Height          =   420
                  Left            =   1560
                  TabIndex        =   43
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
                     TabIndex        =   44
                     Top             =   60
                     Width           =   2955
                  End
               End
               Begin Threed.SSPanel SSPanel2 
                  Height          =   420
                  Left            =   1560
                  TabIndex        =   45
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
                     TabIndex        =   46
                     Top             =   60
                     Width           =   2955
                  End
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   1
                  Left            =   1560
                  TabIndex        =   47
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
                     TabIndex        =   48
                     Top             =   60
                     Width           =   2955
                  End
               End
               Begin Threed.SSPanel pnlTAG 
                  Height          =   420
                  Index           =   0
                  Left            =   1560
                  TabIndex        =   49
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
                  TabIndex        =   50
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
                  PictureBackground=   "frm일일판매현황.frx":F622
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlTAG 
                  Height          =   420
                  Index           =   1
                  Left            =   1560
                  TabIndex        =   51
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
                  TabIndex        =   52
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
                  PictureBackground=   "frm일일판매현황.frx":F848
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   38
                  Left            =   120
                  TabIndex        =   53
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
                  PictureBackground=   "frm일일판매현황.frx":FA6E
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   59
                  Left            =   1560
                  TabIndex        =   54
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
                     TabIndex        =   55
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
                     TabIndex        =   56
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":FC94
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   42
                  Left            =   120
                  TabIndex        =   57
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
                  PictureBackground=   "frm일일판매현황.frx":1035E
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   63
                  Left            =   1560
                  TabIndex        =   58
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
                     TabIndex        =   59
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
                     TabIndex        =   60
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":10584
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   76
                  Left            =   1560
                  TabIndex        =   61
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
                     Index           =   78
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
                     PictureBackground=   "frm일일판매현황.frx":10C4E
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   77
                     Left            =   120
                     TabIndex        =   64
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
                  TabIndex        =   65
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
                     Index           =   80
                     Left            =   2670
                     TabIndex        =   67
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":11318
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   81
                     Left            =   120
                     TabIndex        =   68
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
               Left            =   90
               TabIndex        =   69
               Top             =   90
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
                  TabIndex        =   70
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
                     TabIndex        =   71
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "점"
                     PictureBackground=   "frm일일판매현황.frx":119E2
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtNum01 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   72
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
                     Index           =   15
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
                     PictureBackground=   "frm일일판매현황.frx":120AC
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   0
                  Left            =   675
                  TabIndex        =   75
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
                  PictureBackground=   "frm일일판매현황.frx":12776
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   1
                  Left            =   675
                  TabIndex        =   76
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
                  PictureBackground=   "frm일일판매현황.frx":1299C
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   2
                  Left            =   675
                  TabIndex        =   77
                  Top             =   1470
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
                  PictureBackground=   "frm일일판매현황.frx":12BC2
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   3
                  Left            =   675
                  TabIndex        =   78
                  Top             =   1875
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
                  PictureBackground=   "frm일일판매현황.frx":12DE8
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   1230
                  Index           =   26
                  Left            =   120
                  TabIndex        =   79
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
                  PictureBackground=   "frm일일판매현황.frx":1300E
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   2040
                  Index           =   31
                  Left            =   120
                  TabIndex        =   80
                  Top             =   1470
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
                  PictureBackground=   "frm일일판매현황.frx":13234
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   4
                  Left            =   2115
                  TabIndex        =   81
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
                     TabIndex        =   82
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
                     TabIndex        =   83
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "점"
                     PictureBackground=   "frm일일판매현황.frx":1345A
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   5
                  Left            =   2115
                  TabIndex        =   84
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
                  Begin CSTextLibCtl.sidbEdit txtCost02 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   85
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
                     TabIndex        =   86
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":13B24
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   75
                     Left            =   975
                     TabIndex        =   87
                     Top             =   45
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":141EE
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtCost28 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   88
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
                  TabIndex        =   89
                  Top             =   1875
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
                     TabIndex        =   90
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
                     Index           =   18
                     Left            =   975
                     TabIndex        =   92
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "건"
                     PictureBackground=   "frm일일판매현황.frx":148B8
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   19
                     Left            =   2670
                     TabIndex        =   93
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":14F82
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   32
                  Left            =   675
                  TabIndex        =   94
                  Top             =   3525
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
                  PictureBackground=   "frm일일판매현황.frx":1564C
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   33
                  Left            =   675
                  TabIndex        =   95
                  Top             =   3930
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
                  PictureBackground=   "frm일일판매현황.frx":15872
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   825
                  Index           =   34
                  Left            =   120
                  TabIndex        =   96
                  Top             =   3525
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
                  PictureBackground=   "frm일일판매현황.frx":15A98
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   7
                  Left            =   2115
                  TabIndex        =   97
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
                  Begin CSTextLibCtl.sidbEdit txtCost05 
                     Height          =   345
                     Left            =   1380
                     TabIndex        =   98
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
                     TabIndex        =   99
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":15CBE
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   8
                  Left            =   2115
                  TabIndex        =   100
                  Top             =   3930
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
                     TabIndex        =   101
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
                     TabIndex        =   102
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
                     TabIndex        =   103
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "점"
                     PictureBackground=   "frm일일판매현황.frx":16388
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   23
                     Left            =   2670
                     TabIndex        =   104
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":16A52
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   35
                  Left            =   675
                  TabIndex        =   105
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
                  PictureBackground=   "frm일일판매현황.frx":1711C
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   1230
                  Index           =   36
                  Left            =   120
                  TabIndex        =   106
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
                  PictureBackground=   "frm일일판매현황.frx":17342
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   9
                  Left            =   2115
                  TabIndex        =   107
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
                     TabIndex        =   108
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
                     TabIndex        =   109
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":17568
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   4
                  Left            =   675
                  TabIndex        =   110
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
                  PictureBackground=   "frm일일판매현황.frx":17C32
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   10
                  Left            =   2115
                  TabIndex        =   111
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
                     TabIndex        =   112
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
                     TabIndex        =   113
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":17E58
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   5
                  Left            =   675
                  TabIndex        =   114
                  Top             =   3090
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
                  PictureBackground=   "frm일일판매현황.frx":18522
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   11
                  Left            =   2115
                  TabIndex        =   115
                  Top             =   3090
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
                     Index           =   20
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
                     PictureBackground=   "frm일일판매현황.frx":18748
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   6
                  Left            =   675
                  TabIndex        =   118
                  Top             =   4380
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
                  PictureBackground=   "frm일일판매현황.frx":18E12
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   7
                  Left            =   675
                  TabIndex        =   119
                  Top             =   4785
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
                  PictureBackground=   "frm일일판매현황.frx":19038
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   825
                  Index           =   8
                  Left            =   120
                  TabIndex        =   120
                  Top             =   4380
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
                  PictureBackground=   "frm일일판매현황.frx":1925E
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   12
                  Left            =   2115
                  TabIndex        =   121
                  Top             =   4380
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
                     TabIndex        =   122
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
                     TabIndex        =   123
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":19484
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   13
                  Left            =   2115
                  TabIndex        =   124
                  Top             =   4785
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
                     TabIndex        =   125
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
                     TabIndex        =   126
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
                     TabIndex        =   127
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "건"
                     PictureBackground=   "frm일일판매현황.frx":19B4E
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   26
                     Left            =   2670
                     TabIndex        =   128
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":1A218
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   41
                  Left            =   675
                  TabIndex        =   129
                  Top             =   2280
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
                  PictureBackground=   "frm일일판매현황.frx":1A8E2
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   61
                  Left            =   2115
                  TabIndex        =   130
                  Top             =   2280
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
                     TabIndex        =   131
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
                     TabIndex        =   132
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":1AB08
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   43
                  Left            =   675
                  TabIndex        =   133
                  Top             =   2685
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
                  PictureBackground=   "frm일일판매현황.frx":1B1D2
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   65
                  Left            =   2115
                  TabIndex        =   134
                  Top             =   2685
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
                     TabIndex        =   135
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "건"
                     PictureBackground=   "frm일일판매현황.frx":1B3F8
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtNum12 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   136
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
                     TabIndex        =   137
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
                     TabIndex        =   138
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":1BAC2
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   21
                  Left            =   675
                  TabIndex        =   139
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
                  PictureBackground=   "frm일일판매현황.frx":1C18C
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   0
                  Left            =   2115
                  TabIndex        =   140
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
                     TabIndex        =   141
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
                     TabIndex        =   142
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":1C3B2
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   16
                  Left            =   675
                  TabIndex        =   259
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
                  TabIndex        =   260
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
                     TabIndex        =   261
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "점"
                     PictureBackground=   "frm일일판매현황.frx":1CA7C
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtNum_Internet 
                     Height          =   345
                     Left            =   45
                     TabIndex        =   262
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
                     TabIndex        =   263
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
                     TabIndex        =   264
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":1D146
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
            End
            Begin XtremeSuiteControls.GroupBox GroupBox 
               Height          =   6795
               Index           =   2
               Left            =   5430
               TabIndex        =   143
               Top             =   105
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
                  TabIndex        =   144
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
                  PictureBackground=   "frm일일판매현황.frx":1D810
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   23
                  Left            =   75
                  TabIndex        =   145
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
                  PictureBackground=   "frm일일판매현황.frx":1DA36
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   82
                  Left            =   2145
                  TabIndex        =   146
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
                     TabIndex        =   147
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
                     TabIndex        =   148
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":1DC5C
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   24
                  Left            =   75
                  TabIndex        =   149
                  ToolTipText     =   "일일 총 접수금액의 % 금액"
                  Top             =   6615
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
                  Caption         =   "+ 로열티 1 사용안함"
                  PictureBackgroundStyle=   2
                  PictureBackground=   "frm일일판매현황.frx":1E326
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   84
                  Left            =   2145
                  TabIndex        =   150
                  Top             =   6615
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
                     TabIndex        =   151
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
                     TabIndex        =   152
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":1E54C
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   25
                  Left            =   75
                  TabIndex        =   153
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
                  PictureBackground=   "frm일일판매현황.frx":1EC16
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   86
                  Left            =   2145
                  TabIndex        =   154
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
                     TabIndex        =   155
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
                     TabIndex        =   156
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":1EE3C
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   29
                  Left            =   75
                  TabIndex        =   157
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
                  PictureBackground=   "frm일일판매현황.frx":1F506
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   88
                  Left            =   2145
                  TabIndex        =   158
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
                     TabIndex        =   159
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
                     TabIndex        =   160
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":1F72C
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   30
                  Left            =   75
                  TabIndex        =   161
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
                  PictureBackground=   "frm일일판매현황.frx":1FDF6
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   90
                  Left            =   2145
                  TabIndex        =   162
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
                     TabIndex        =   163
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
                     TabIndex        =   164
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":2001C
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   44
                  Left            =   75
                  TabIndex        =   165
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
                  PictureBackground=   "frm일일판매현황.frx":206E6
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   92
                  Left            =   2145
                  TabIndex        =   166
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
                     TabIndex        =   167
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
                     TabIndex        =   168
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":2090C
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   435
                  Index           =   46
                  Left            =   90
                  TabIndex        =   169
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
                  PictureBackground=   "frm일일판매현황.frx":20FD6
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   94
                  Left            =   1560
                  TabIndex        =   170
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
                     TabIndex        =   171
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "건"
                     PictureBackground=   "frm일일판매현황.frx":211FC
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtCard 
                     Height          =   345
                     Index           =   0
                     Left            =   225
                     TabIndex        =   172
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
                     TabIndex        =   173
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
                     TabIndex        =   174
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":218C6
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   825
                  Index           =   45
                  Left            =   90
                  TabIndex        =   175
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
                  PictureBackground=   "frm일일판매현황.frx":21F90
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   97
                  Left            =   1560
                  TabIndex        =   176
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
                     TabIndex        =   177
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
                     TabIndex        =   178
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":221B6
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   99
                     Left            =   120
                     TabIndex        =   179
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
                  TabIndex        =   180
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
                     TabIndex        =   181
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "건"
                     PictureBackground=   "frm일일판매현황.frx":22880
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin CSTextLibCtl.sidbEdit txtCard 
                     Height          =   345
                     Index           =   3
                     Left            =   225
                     TabIndex        =   182
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
                     TabIndex        =   183
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
                     TabIndex        =   184
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":22F4A
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   825
                  Index           =   47
                  Left            =   90
                  TabIndex        =   185
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
                  PictureBackground=   "frm일일판매현황.frx":23614
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   103
                  Left            =   1560
                  TabIndex        =   186
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
                     TabIndex        =   187
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
                     TabIndex        =   188
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":2383A
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   300
                     Index           =   105
                     Left            =   120
                     TabIndex        =   189
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
                  TabIndex        =   190
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
                  PictureBackground=   "frm일일판매현황.frx":23F04
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   106
                  Left            =   2145
                  TabIndex        =   191
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
                     TabIndex        =   192
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
                     TabIndex        =   193
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":2412A
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   49
                  Left            =   75
                  TabIndex        =   194
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
                  PictureBackground=   "frm일일판매현황.frx":247F4
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   108
                  Left            =   2145
                  TabIndex        =   195
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
                     TabIndex        =   196
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
                     TabIndex        =   197
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":24A1A
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   50
                  Left            =   75
                  TabIndex        =   250
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
                  PictureBackground=   "frm일일판매현황.frx":250E4
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   110
                  Left            =   2145
                  TabIndex        =   251
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
                     TabIndex        =   252
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
                     TabIndex        =   253
                     Top             =   60
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   529
                     _Version        =   262144
                     Font3D          =   3
                     BackColor       =   16777215
                     Caption         =   "원"
                     PictureBackground=   "frm일일판매현황.frx":2530A
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
               Begin Threed.SSPanel pnlData 
                  Height          =   420
                  Index           =   51
                  Left            =   75
                  TabIndex        =   255
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
                  PictureBackground=   "frm일일판매현황.frx":259D4
                  BevelOuter      =   0
                  RoundedCorners  =   0   'False
                  Outline         =   -1  'True
                  FloodShowPct    =   -1  'True
               End
               Begin Threed.SSPanel SSPanel 
                  Height          =   420
                  Index           =   112
                  Left            =   2145
                  TabIndex        =   256
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
                     Index           =   113
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
                     PictureBackground=   "frm일일판매현황.frx":25BFA
                     BevelOuter      =   0
                     RoundedCorners  =   0   'False
                     FloodShowPct    =   -1  'True
                  End
               End
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage1 
            Height          =   10275
            Left            =   -69970
            TabIndex        =   11
            Top             =   450
            Visible         =   0   'False
            Width           =   16320
            _Version        =   851970
            _ExtentX        =   28787
            _ExtentY        =   18124
            _StockProps     =   1
            Page            =   0
            Begin FPSpreadADO.fpSpread sprList 
               Height          =   6345
               Left            =   0
               TabIndex        =   15
               Top             =   0
               Width           =   16380
               _Version        =   524288
               _ExtentX        =   28893
               _ExtentY        =   11192
               _StockProps     =   64
               AutoCalc        =   0   'False
               BackColorStyle  =   1
               DAutoCellTypes  =   0   'False
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
               FormulaSync     =   0   'False
               GrayAreaBackColor=   16777215
               GridSolid       =   0   'False
               MaxCols         =   22
               MaxRows         =   200
               MoveActiveOnFocus=   0   'False
               OperationMode   =   1
               Protect         =   0   'False
               SpreadDesigner  =   "frm일일판매현황.frx":262C4
               UserResize      =   1
               VisibleCols     =   13
               VisibleRows     =   50
               Appearance      =   1
               HighlightHeaders=   1
               HighlightStyle  =   1
            End
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   750
         Index           =   0
         Left            =   15
         TabIndex        =   3
         Top             =   435
         Width           =   16380
         _ExtentX        =   28893
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.CheckBox chkPageView 
            Height          =   405
            Left            =   4500
            TabIndex        =   249
            Top             =   150
            Width           =   2145
            _Version        =   851970
            _ExtentX        =   3784
            _ExtentY        =   714
            _StockProps     =   79
            Caption         =   "특정 페이지 출력"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   630
            Left            =   7065
            TabIndex        =   1
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 조회(&F)"
            Appearance      =   6
            Picture         =   "frm일일판매현황.frx":2700D
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   3
            Left            =   8595
            TabIndex        =   5
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 엑셀(&E)"
            Appearance      =   6
            Picture         =   "frm일일판매현황.frx":27707
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   13170
            TabIndex        =   6
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frm일일판매현황.frx":27E81
         End
         Begin XtremeSuiteControls.PushButton cmdPrint 
            Height          =   630
            Left            =   10110
            TabIndex        =   7
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 출력(&P)"
            Appearance      =   6
            Picture         =   "frm일일판매현황.frx":28F13
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Left            =   915
            TabIndex        =   0
            Top             =   60
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
            Format          =   56754179
            CurrentDate     =   40279
         End
         Begin XtremeSuiteControls.PushButton cmdPrintMini 
            Height          =   630
            Left            =   11640
            TabIndex        =   14
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 단말기출력"
            Appearance      =   6
            Picture         =   "frm일일판매현황.frx":2960D
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "접수일자:"
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
            Left            =   45
            TabIndex        =   8
            Top             =   120
            Width           =   840
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   405
         Left            =   15
         TabIndex        =   4
         Top             =   15
         Width           =   16380
         _ExtentX        =   28893
         _ExtentY        =   714
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
         Caption         =   "      일일판매 현황"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm일일판매현황.frx":29D07
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm일일판매현황.frx":29F2D
            Top             =   -15
            Width           =   765
         End
      End
   End
End
Attribute VB_Name = "frm일일판매현황"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim 마감일자 As String

Private Sub Resize_Rtn()
    sprList.Width = TabControl1.Width - 180
    sprList.Height = TabControl1.Height - 580
    
    sprCloth.Height = TabControl1.Height - 580
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 3: Call Export_Excel(frmMain.cdgExcel, sprList)
        
        Case 5:
            Unload Me
    End Select
End Sub

Private Sub cmdList_Click()
    On Error GoTo ErrRtn
    
    cmdBtn(5).Enabled = False
    pnlProg.Left = 540
    pnlProg.Top = 1890
    pnlProg.Visible = True
    DoEvents
    
    Query = "SELECT    A.접수번호"
    Query = Query & ", A.접수일자"
    Query = Query & ", A.접수시간"
    Query = Query & ", A.예정일자"
    Query = Query & ", A.출고일자"
    Query = Query & ", A.출고시간"
    Query = Query & ", A.지사출고상태"
    Query = Query & ", A.의류명"
    Query = Query & ", A.택번호"
    Query = Query & ", A.색상"
    Query = Query & ", A.무늬"
    Query = Query & ", A.내용"
    Query = Query & ", A.금액"
    Query = Query & ", A.결제여부"
    Query = Query & ", A.상표"
    Query = Query & ", A.오점내용"
    Query = Query & ", SUBSTRING(A.가맹점출고일자,1,10) AS 가맹점출고일자"
    Query = Query & ", SUBSTRING(A.가맹점입고일자,1,10) AS 가맹점입고일자"
    Query = Query & ", SUBSTRING(A.지사입고일자,1,10)   AS 지사입고일자"
    Query = Query & ", SUBSTRING(A.지사출고일자,1,10)   AS 지사출고일자"
    Query = Query & ", B.성명"
    Query = Query & ", B.휴대전화"
    Query = Query & ", B.전화번호"
    Query = Query & " FROM TB_입출고 AS A LEFT OUTER JOIN TB_고객정보 AS B ON (A.고객코드 = B.고객코드) "
    Query = Query & " WHERE (A.접수일자 = '" & Format(dtpDay.Value, "YYYY-MM-DD") & "')"
    Query = Query & "   AND (판매취소 <> 'Y')"
    
    'Query = Query & "   AND ((판매취소 <> 'Y')"
    ''Query = Query & "   AND ((A.판매취소 = '') AND (A.판매취소일자 IS NULL OR A.판매취소일자 = '')"
    'Query = Query & "   AND (A.반품환불일자 IS NULL OR A.반품환불일자 = '')"
    'Query = Query & "   AND (A.세탁환불일자 IS NULL OR A.세탁환불일자 = ''))"
    
    Query = Query & " ORDER BY A.접수일자 DESC, A.택번호 ASC, A.접수번호 ASC "
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    With sprList
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = ADORs!접수번호 & ""                   ' 1
            .Col = 2:  .Text = ADORs!접수일자 & ""                   ' 1
            
            If ADORs!접수시간 = "" Then
                .Col = 3: .Text = " "
            Else
                .Col = 3:  .Text = Left(ADORs!접수시간, 5) & ""      ' 2
            End If
            
            If Trim(ADORs!성명) = "" Then
                .Col = 4: .Text = " "
            Else
                .Col = 4:  .Text = ADORs!성명 & ""                       ' 3
            End If
            
            If Trim(ADORs!휴대전화) = "" Then
                .Col = 5: .Text = " "
            Else
                .Col = 5:  .Text = ADORs!휴대전화 & ""                   ' 4
            End If
            
            If Trim(ADORs!전화번호) = "" Then
                .Col = 6: .Text = " "
            Else
                .Col = 6:  .Text = ADORs!전화번호 & ""                   ' 5
            End If
            
            .Col = 7:  .Text = ADORs!예정일자 & ""                   ' 6
            .Col = 8:  .Text = ADORs!출고일자 & ""                   ' 7
            .Col = 9:  .Text = Left(ADORs!출고시간, 5) & ""          ' 8
            .Col = 10: .Text = Format(ADORs!택번호, "000-00-0000")    ' 9
            .Col = 11: .Text = ADORs!의류명 & ""                     '10
            .Col = 12: .Text = ADORs!색상 & ""                       '11
            .Col = 13: .Text = ADORs!무늬 & ""                       '12
            .Col = 14: .Text = ADORs!내용 & ""                       '13
            .Col = 15: .Text = ADORs!금액 & ""                       '14
            .Col = 16: .Text = ADORs!결제여부 & ""                   '15
            .Col = 17: .Text = ADORs!상표 & ""                       '16
            .Col = 18: .Text = ADORs!오점내용 & ""                       '16
            .Col = 19: .Text = ADORs!가맹점출고일자 & ""             '17
            .Col = 20: .Text = ADORs!가맹점입고일자 & ""             '18
            .Col = 21: .Text = ADORs!지사입고일자 & ""               '19
            .Col = 22: .Text = ADORs!지사출고일자 & ""               '20
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        .ReDraw = True
    End With
    
    Debug.Print Now & "일일마감_Proc 시작"
    Call 일일마감_Proc
    
    Debug.Print Now & "접수집계_Display 시작"
    Call 접수집계_Display
    
    Debug.Print Now & "일일마감 종료"
    
    pnlProg.Visible = False
    cmdBtn(5).Enabled = True
    
    Exit Sub
    
ErrRtn:
    pnlProg.Visible = False
    cmdBtn(5).Enabled = True
    
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub cmdPrint_Click()
    Dim vText       As Variant
    Dim sTempKey    As String
    Dim sNameKey    As String
    
    Dim 입금액    As Long
    Dim 요일      As String
    
    Dim ComboList As String
    
    On Error GoTo ErrRtn
    
    Call cmdList_Click
    
    If sprList.MaxRows = 0 Then Exit Sub

    If Dir(AppPath & "XML", vbDirectory) = "" Then
        MkDir AppPath & "XML"
    End If

    If Get_일일마감여부(Format(dtpDay.Value, "YYYY-MM-DD")) = False Then
        MsgBox "일마감이 완료 되지 않아 출력할 수 없습니다.", vbInformation, "확인"
        Exit Sub
    End If
    
    요일 = Fun_Week(dtpDay.Value)
    sNameKey = ""
    Open AppPath & "XML\일일매출현황2.XML" For Output As #1

    Print #1, "<?xml version=""1.0"" encoding=""EUC-KR""?>"
    Print #1, "<root>"

          XML = "    <조건>"
    XML = XML & "        <접수일자>일자 : " & Format(dtpDay.Value, "YYYY년 MM월 DD일") & " (" & 요일 & ")</접수일자>"
    XML = XML & "        <가맹점>(" & Func_Replace(가맹점정보.가맹점명) & ") 일일매출현황</가맹점>"
    XML = XML & "   </조건>"
    Print #1, XML

    With sprList
        For i = 1 To .DataRowCnt
            .Row = i

                             XML = "    <Data>"
            .Col = 10:  XML = XML & "       <택번호>" & Right(.Text, 7) & "</택번호>"
            
            ' 이름이 다를 경우만 출력 한다.
            .GetText 4, i, vText:   sTempKey = CStr(vText)
            .GetText 5, i, vText:   sTempKey = sTempKey & CStr(vText)
            .GetText 6, i, vText:   sTempKey = sTempKey & CStr(vText)
            
            If sNameKey <> sTempKey Then
                sNameKey = sTempKey
                .Col = 4:  XML = XML & "        <성명>" & Func_Replace(.Text) & "</성명>"
                .Col = 5:  XML = XML & "        <휴대전화>" & Func_Replace(.Text) & "</휴대전화>"
                .Col = 6:  XML = XML & "        <전화번호>" & Func_Replace(.Text) & "</전화번호>"
            Else
                .Col = 4:  XML = XML & "        <성명>" & Space(1) & "</성명>"
                .Col = 5:  XML = XML & "        <휴대전화>" & Space(1) & "</휴대전화>"
                .Col = 6:  XML = XML & "        <전화번호>" & Space(1) & "</전화번호>"
            End If
            
            
            ' 사람이 변경되면
            .GetText 4, i + 1, vText: sTempKey = CStr(vText)
            .GetText 5, i + 1, vText: sTempKey = sTempKey & CStr(vText)
            .GetText 6, i + 1, vText: sTempKey = sTempKey & CStr(vText)
             XML = XML & "        <선긋기>" & IIf(sNameKey <> sTempKey, "OK", "NO") & "</선긋기>"
            
            .Col = 11: XML = XML & "        <품명>" & Func_Replace(.Text) & "</품명>"
            .Col = 12: XML = XML & "        <색상>" & Func_Replace(.Text) & "</색상>"
            .Col = 13: XML = XML & "        <무늬>" & Func_Replace(.Text) & "</무늬>"
            .Col = 14: XML = XML & "        <내용>" & Func_Replace(.Text) & "</내용>"
            .Col = 15: XML = XML & "        <금액>" & .Text & "</금액>"
            .Col = 17: XML = XML & "        <상표>" & Func_Replace(.Text) & "</상표>"
            .Col = 16: XML = XML & "        <결제>" & .Text & "</결제>"
                       XML = XML & "   </Data>" & vbNewLine
                       Print #1, XML
        Next i
    End With

          XML = "    <합계>"
    XML = XML & "        <접수수량>" & txtNum01.Text & " 점</접수수량>"
    XML = XML & "        <접수금액>" & txtCost01.Text & " 원</접수금액>"
    
    XML = XML & "        <가맹점마진>" & txtCost09.Text & " 원</가맹점마진>"
    XML = XML & "        <지사마진>" & txtCost10.Text & " 원</지사마진>"
    XML = XML & "        <발생마일리지>" & txtCost18.Text & " 원</발생마일리지>"
    XML = XML & "        <사용마일리지>" & txtCost19.Text & " 원</사용마일리지>"
    XML = XML & "        <삭제마일리지>" & txtCost20.Text & " 원</삭제마일리지>"
    
    
    XML = XML & "        <현금결제>" & txtCost02.Text & " 원</현금결제>"
    XML = XML & "        <미수금액>" & txtCost04.Text & " 원</미수금액>"
    XML = XML & "        <카드결제>" & txtCost08.Text & " 원</카드결제>"
    XML = XML & "        <카드건수>" & txtNum05.Text & " 건</카드건수>"
    XML = XML & "        <쿠폰결제>" & txtCost21.Text & " 원</쿠폰결제>"
    XML = XML & "        <현금반환>" & txtCost28.Text & " 원</현금반환>"
    
    XML = XML & "        <수금합계>" & Format(txtCost05.Value + txtCost06.Value, "#,##0") & " 원</수금합계>"
    XML = XML & "        <수금현금결제>" & txtCost05.Text & " 원</수금현금결제>"
    XML = XML & "        <수금카드결제>" & txtCost06.Text & " 원</수금카드결제>"
    XML = XML & "        <수금카드건수>" & txtNum04.Text & " 원</수금카드건수>"
    '---------------------------------------------------------------------------
    XML = XML & "        <카드수수료지원금>" & txtMaster(1).Text & " 원</카드수수료지원금>"
    XML = XML & "        <카드수수료환불금>" & txtMaster(2).Text & " 원</카드수수료환불금>"
    XML = XML & "        <로열티금액>" & txtMaster(3).Text & " 원</로열티금액>"
    XML = XML & "        <환불금액>" & txtMaster(4).Text & " 원</환불금액>"
    XML = XML & "        <쿠폰지사60>" & txtMaster(8).Text & " 원</쿠폰지사60>"
    XML = XML & "        <전산사용료>" & txtMaster(9).Text & " 원</전산사용료>"
    XML = XML & "        <지사정산금액>" & txtMaster(5).Text & " 원</지사정산금액>"
    
    '---------------------------------------------------------------------------
    XML = XML & "        <판매취소수량>" & txtNum13.Text & " 점</판매취소수량>"
    XML = XML & "        <반품환불수량>" & txtNum14.Text & " 점</반품환불수량>"
    XML = XML & "        <세탁환불수량>" & txtNum15.Text & " 점</세탁환불수량>"
    XML = XML & "        <누락택수량>" & txtNum16.Text & " 점</누락택수량>"
    
    ComboList = ""
    
    For i = 0 To cboCancel.ListCount - 1
        ComboList = ComboList & cboCancel.List(i) & " "
    Next i
    
    XML = XML & "        <판매취소택>" & ComboList & "</판매취소택>"
    '---------------------------------------------------------------------------
    
    ComboList = ""
    
    For i = 0 To cboReturn.ListCount - 1
        ComboList = ComboList & cboReturn.List(i) & " "
    Next i
    
    XML = XML & "        <반품환불택>" & ComboList & "</반품환불택>"
    '---------------------------------------------------------------------------
    
    ComboList = ""
    
    For i = 0 To cboRepay.ListCount - 1
        ComboList = ComboList & cboRepay.List(i) & " "
    Next i
    
    XML = XML & "        <세탁환불택>" & ComboList & "</세탁환불택>"
    '---------------------------------------------------------------------------
    
    ComboList = ""
    
    For i = 0 To cboMissTag.ListCount - 1
        ComboList = ComboList & cboMissTag.List(i) & " "
    Next i
    
    XML = XML & "        <누락택>" & ComboList & "</누락택>"
    
    '---------------------------------------------------------------------------
    
    XML = XML & "   </합계>"
    Print #1, XML

    Print #1, "</root>"
    Close #1

    With rpt일일매출현황2
        .dc.FileURL = AppPath & "XML\일일매출현황2.XML"
        
        If chkPageView.Value = xtpChecked Then
            .Show 1
        Else
            .PrintReport False
        End If
        
    End With
    
    Unload rpt일일매출현황2
    
    Exit Sub
    
ErrRtn:
    
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub cmdPrintMini_Click()
    On Error GoTo ErrRtn
    
    Dim Print_Msg As String

    Dim tmp      As String
    Dim Cnt     As Long

    If Get_일일마감여부(Format(dtpDay.Value, "YYYY-MM-DD")) = False Then
        MsgBox "일마감이 완료 되지 않아 출력할 수 없습니다.", vbInformation, "확인"
        Exit Sub
    End If

    Print_Msg = Print_Msg & PrintTitle(" 매출현황")

    Print_Msg = Print_Msg & PrintLineFeed

    Print_Msg = Print_Msg & PrintString("상 호 명 : " + 가맹점정보.가맹점명, 6, True)
    Print_Msg = Print_Msg & PrintString("전화번호 : " + 가맹점정보.전화매장, 6, True)

    Print_Msg = Print_Msg & PrintString("===============================================", 1)
    Print_Msg = Print_Msg & PrintString("영업일자 : " + Format(dtpDay.Value, "YYYY년 MM월 DD일 ") & WeekdayName(Weekday(dtpDay.Value)), 6, True)
    Print_Msg = Print_Msg & PrintString("택 정 보 : " + pnlTAG(0).Caption & " ~ " & pnlTAG(1).Caption, 6, True)
    Print_Msg = Print_Msg & PrintString("===============================================", 1)

    Print_Msg = Print_Msg & PrintString("접수수량 : " + String(8 - LenH(CStr(txtNum01.Value)), " ") + CStr(txtNum01.Value) + "점 / 접수금액 : " + String(9 - LenH(Format(txtCost01.Value, "#,##0")), " ") + Format(txtCost01.Value, "#,##0") + "원", 6, True)
    Print_Msg = Print_Msg & PrintString("-----------------------------------------------", 1)
    Print_Msg = Print_Msg & PrintString("사용마일리지 : " + String(30 - LenH(Format(txtCost27.Value, "#,##0")), " ") + Format(txtCost27.Value, "#,##0") + "원", 1)
    Print_Msg = Print_Msg & PrintString("가맹점  마진 : " + String(30 - LenH(Format(txtCost09.Value, "#,##0")), " ") + Format(txtCost09.Value, "#,##0") + "원", 1)
    Print_Msg = Print_Msg & PrintString("지  사  마진 : " + String(30 - LenH(Format(txtCost10.Value, "#,##0")), " ") + Format(txtCost10.Value, "#,##0") + "원", 1)
    Print_Msg = Print_Msg & PrintString("===============================================", 1)
    Print_Msg = Print_Msg & PrintString("현금결제 : " + String(9 - LenH(Format(txtCost02.Value, "#,##0")), " ") + Format(txtCost02.Value, "#,##0") + "원 / 카드결제 : " + String(9 - LenH(Format(txtCost03.Value, "#,##0")), " ") + Format(txtCost03.Value, "#,##0") + "원", 1)
    Print_Msg = Print_Msg & PrintString("마일리지 : " + String(9 - LenH(Format(txtCost19.Value, "#,##0")), " ") + Format(txtCost19.Value, "#,##0") + "원 / 미수금액 : " + String(9 - LenH(Format(txtCost04.Value, "#,##0")), " ") + Format(txtCost04.Value, "#,##0") + "원", 1)
    Print_Msg = Print_Msg & PrintString("쿠폰결제 : " + String(9 - LenH(Format(txtCost21.Value, "#,##0")), " ") + Format(txtCost21.Value, "#,##0") + "원 / 현금반환 : " + String(9 - LenH(Format(txtCost28.Value, "#,##0")), " ") + Format(txtCost28.Value, "#,##0") + "원", 1)
    Print_Msg = Print_Msg & PrintString("-----------------------------------------------", 1)
    Print_Msg = Print_Msg & PrintString("발생마일 : " + String(9 - LenH(Format(txtCost18.Value, "#,##0")), " ") + Format(txtCost18.Value, "#,##0") + "원 / 삭제마일 : " + String(9 - LenH(Format(txtCost20.Value, "#,##0")), " ") + Format(txtCost20.Value, "#,##0") + "원", 1)
    Print_Msg = Print_Msg & PrintString("===============================================", 1)
    Print_Msg = Print_Msg & PrintString("**********    미 수 금 수 금 정 보   **********", 1)
    Print_Msg = Print_Msg & PrintString("-----------------------------------------------", 1)
    Print_Msg = Print_Msg & PrintString("수금합계 : " + String(34 - LenH(Format(txtCost05.Value + txtCost06.Value, "#,##0")), " ") + Format(txtCost05.Value + txtCost06.Value, "#,##0") + "원", 1)
    Print_Msg = Print_Msg & PrintString("현금결제 : " + String(9 - LenH(Format(txtCost05.Value, "#,##0")), " ") + Format(txtCost05.Value, "#,##0") + "원 / 카드결제 : " + String(9 - LenH(Format(txtCost06.Value, "#,##0")), " ") + Format(txtCost06.Value, "#,##0") + "원", 1)
    Print_Msg = Print_Msg & PrintString("===============================================", 1)
    Print_Msg = Print_Msg & PrintString("**********    정    산   정    보   **********", 6, True)
    Print_Msg = Print_Msg & PrintString("-----------------------------------------------", 1)
    Print_Msg = Print_Msg & PrintString("지 사 분   매 출 :  " + String(24 - LenH(Format(txtMaster(0).Value, "#,##0")), " ") + Format(txtMaster(0).Value, "#,##0") + "원", 6, True)
    Print_Msg = Print_Msg & PrintString("카드수수료지원금 : -" + String(24 - LenH(Format(txtMaster(1).Value, "#,##0")), " ") + Format(txtMaster(1).Value, "#,##0") + "원", 6, True)
    Print_Msg = Print_Msg & PrintString("카드수수료환불금 : +" + String(24 - LenH(Format(txtMaster(2).Value, "#,##0")), " ") + Format(txtMaster(2).Value, "#,##0") + "원", 6, True)
    Print_Msg = Print_Msg & PrintString("로 열 티   금 액 : +" + String(24 - LenH(Format(txtMaster(3).Value, "#,##0")), " ") + Format(txtMaster(3).Value, "#,##0") + "원", 6, True)
    Print_Msg = Print_Msg & PrintString("환  불   금   액 : -" + String(24 - LenH(Format(txtMaster(4).Value, "#,##0")), " ") + Format(txtMaster(4).Value, "#,##0") + "원", 6, True)
    Print_Msg = Print_Msg & PrintString("쿠폰지사금액(60%): -" + String(24 - LenH(Format(txtMaster(8).Value, "#,##0")), " ") + Format(txtMaster(8).Value, "#,##0") + "원", 6, True)
    Print_Msg = Print_Msg & PrintString("전 산   사 용 료 : -" + String(24 - LenH(Format(txtMaster(9).Value, "#,##0")), " ") + Format(txtMaster(9).Value, "#,##0") + "원", 6, True)
    Print_Msg = Print_Msg & PrintString("-----------------------------------------------", 1)
    Print_Msg = Print_Msg & PrintString("지 사 정 산 금액 : =" + String(24 - LenH(Format(txtMaster(5).Value, "#,##0")), " ") + Format(txtMaster(5).Value, "#,##0") + "원", 6, True)
    Print_Msg = Print_Msg & PrintString("===============================================", 1)

    Print_Msg = Print_Msg & PrintString("판매취소수량 : " + String(30 - LenH(Format(txtNum13.Value, "#,##0")), " ") + Format(txtNum13.Value, "#,##0") + "점", 1)
    tmp = "":   Cnt = 0
    For i = 0 To IIf(cboCancel.ListCount - 1 > 50, 50, cboCancel.ListCount - 1)
        Cnt = Cnt + 1
        tmp = tmp & Mid(cboCancel.List(i), 5, 7) & " "
        If Cnt = 5 Then
            Print_Msg = Print_Msg & PrintString(Space(8) & tmp, 1)
            tmp = "":   Cnt = 0
        End If
    Next i
    tmp = tmp & IIf(cboCancel.ListCount - 1 > 50, "......외", "")
    If tmp <> "" Then Print_Msg = Print_Msg & PrintString(Space(8) & tmp, 1, True)

    Print_Msg = Print_Msg & PrintString("반품환불수량 : " + String(30 - LenH(Format(txtNum14.Value, "#,##0")), " ") + Format(txtNum14.Value, "#,##0") + "점", 1)
    tmp = "":   Cnt = 0
    For i = 0 To IIf(cboReturn.ListCount - 1 > 50, 50, cboReturn.ListCount - 1)
        Cnt = Cnt + 1
        tmp = tmp & Mid(cboReturn.List(i), 5, 7) & " "
        If Cnt = 5 Then
            Print_Msg = Print_Msg & PrintString(Space(8) & tmp, 1)
            tmp = "":   Cnt = 0
        End If
    Next i
    tmp = tmp & IIf(cboReturn.ListCount - 1 > 50, "......외", "")
    If tmp <> "" Then Print_Msg = Print_Msg & PrintString(Space(8) & " " + tmp, 1, True)


    Call PrintString("세탁환불수량 : " + String(30 - LenH(Format(txtNum15.Value, "#,##0")), " ") + Format(txtNum15.Value, "#,##0") + "점", 1)
    tmp = "":   Cnt = 0
    For i = 0 To IIf(cboRepay.ListCount - 1 > 50, 50, cboRepay.ListCount - 1)
        Cnt = Cnt + 1
        tmp = tmp & Mid(cboRepay.List(i), 5, 7) & " "
        If Cnt = 5 Then
            Print_Msg = Print_Msg & PrintString(Space(8) & tmp, 1)
            tmp = "":   Cnt = 0
        End If
    Next i
    tmp = tmp & IIf(cboRepay.ListCount - 1 > 50, "......외", "")
    If tmp <> "" Then Print_Msg = Print_Msg & PrintString(Space(8) & " " + tmp, 1, True)

    Print_Msg = Print_Msg & PrintString("누락택  수량 : " + String(30 - LenH(Format(txtNum16.Value, "#,##0")), " ") + Format(txtNum16.Value, "#,##0") + "점", 1)
    tmp = "":   Cnt = 0
    For i = 0 To IIf(cboMissTag.ListCount - 1 > 50, 50, cboMissTag.ListCount - 1)
        Cnt = Cnt + 1
        tmp = tmp & Mid(cboMissTag.List(i), 5, 7) & " "
        If Cnt = 5 Then
            Print_Msg = Print_Msg & PrintString(Space(8) & tmp, 1)
            tmp = "":   Cnt = 0
        End If
    Next i
    tmp = tmp & IIf(cboMissTag.ListCount - 1 > 50, "......외", "")
    If tmp <> "" Then Print_Msg = Print_Msg & PrintString(Space(8) & " " + tmp, 1, True)

    Print_Msg = Print_Msg & PrintString("-----------------------------------------------", 1)

    Print_Msg = Print_Msg & PrintLineFeed(4)
    Print_Msg = Print_Msg & PrintCut

    Call frmKicc.Card_Print(Print_Msg)
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub dtpDay_Change()
    DoEvents
    Call cmdList_Click
End Sub

Private Sub Form_Activate()
    Call Resize_Rtn
    
    Call cmdList_Click
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    With sprList
        .MaxRows = 0
        .RowHeight(-1) = 14
                
        .Col = 1: .ColHidden = True
        
        .Col = 1: .ColMerge = MergeRestricted
        .Col = 2: .ColMerge = MergeRestricted
        .Col = 3: .ColMerge = MergeRestricted
        .Col = 4: .ColMerge = MergeRestricted
        .Col = 5: .ColMerge = MergeRestricted
        .Col = 6: .ColMerge = MergeRestricted
        
        .ColsFrozen = 6
        
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
        
    dtpDay.Value = Format(Date, "YYYY-MM-DD")
    
    
    TabControl1.SelectedItem = 0
    
    Call Resize_Rtn
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call Resize_Rtn
    
    cmdBtn(5).Left = Me.Width - cmdBtn(5).Width - 200
End Sub

'
'Private Sub 일일마감_Proc()
'    Dim 접수금액   As Long
'    Dim 가맹점마진 As Long
'    Dim 외주마진   As Long
'
'    Dim 미수금액 As Long
'    Dim tmpData  As String
'
'    On Error GoTo ErrRtn
'
'    Screen.MousePointer = 11
'
'    pnlProg.Left = 90
'    pnlProg.Top = 1305
'    pnlProg.Visible = True
'    DoEvents
'
'
'    '컨트롤 초기화
'    Dim ctrl As Control
'    Dim txt  As sidbEdit
'
'    For Each ctrl In Me.Controls
'        If TypeOf ctrl Is sidbEdit Then
'            ctrl.Value = 0
'        End If
'    Next ctrl
'
'    cboCancel.Clear
'    cboReturn.Clear
'    cboRepay.Clear
'    cboMissTag.Clear
'
'    마감일자 = Format(dtpDay.Value, "YYYY-MM-DD")
'
'    '----------------------------------------------------------------
'    ' 1. 매출 -> 1-1) 접수수량, 접수금액 구하기
'    '----------------------------------------------------------------
'    Debug.Print Now & "  1. 매출 -> 1-1) 시작"
'
'    Query = "SELECT    ISNULL(COUNT(택번호),0)"
'    Query = Query & ", ISNULL(SUM(금액),0)"
'    Query = Query & " FROM TB_입출고"
'    Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND (판매취소 <> 'Y') "
'
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    txtNum01.Value = ADORs(0)  '접수건수
'    txtCost01.Value = ADORs(1) '접수금액
'
'    ADORs.Close:    Set ADORs = Nothing
'
'
''''''''''2012-04-25 오후 4:00:53일일마감_Proc 시작
''''''''''2012-04-25 오후 4:00:53  1. 매출 -> 1-1) 시작
''''''''''2012-04-25 오후 4:00:53  1. 매출 -> 1-2) 출고수량 구하기 시작
''''''''''2012-04-25 오후 4:01:24  2. 선불결제 2-1) 현금반환/ 현금결제 구하기 시작
''''''''''2012-04-25 오후 4:01:30  2. 선불결제 2-2) 현금결제 구하기 시작
''''''''''2012-04-25 오후 4:01:30  2. 선불결제 2-3) 카드결제 건수 / 금액 구하기
''''''''''2012-04-25 오후 4:01:30  2. 선불결제 2-4) 발생/사용/삭제 마일리지
''''''''''2012-04-25 오후 4:01:30  2. 선불결제 2-5) 쿠폰 사용 건수/ 금액
''''''''''2012-04-25 오후 4:01:30  2. 선불결제 2-6) 미수금 금액
''''''''''2012-04-25 오후 4:01:30  3. 미수결제 3-1) 미수금 수금 현금결제 구하기
''''''''''2012-04-25 오후 4:01:30  3. 미수결제 3-2) 미수금 수금 카드결제 구하기
''''''''''2012-04-25 오후 4:01:30  4. 결제합계 현금/ 카드 결제
''''''''''2012-04-25 오후 4:01:30  5. 마진 5-2) 가맹점 마진
''''''''''2012-04-25 오후 4:01:30  6. 기타자료 6-1) 수선수량 계산
''''''''''2012-04-25 오후 4:01:30  6. 기타자료 6-2) 재세탁수량 계산
''''''''''2012-04-25 오후 4:01:30  6. 기타자료 6-3) 운동화 매출을 불러온다.
''''''''''2012-04-25 오후 4:01:30  6. 기타자료 6-4) 가죽/무스탕 매출을 불러온다.
''''''''''2012-04-25 오후 4:01:30  6. 기타자료 6-5) 카페트 매출을 불러온다.
''''''''''2012-04-25 오후 4:01:30  6. 기타자료 6-6) 반품 내역
''''''''''2012-04-25 오후 4:01:30  6. 기타자료 6-7) 외주 마진
''''''''''2012-04-25 오후 4:01:30  7) 기타자료2 7-1) 판매취소 내역
''''''''''2012-04-25 오후 4:01:31  7) 기타자료2 7-2) 반품환불 내역
''''''''''2012-04-25 오후 4:01:39  7) 기타자료2 7-3) 세탁환불 내역
''''''''''2012-04-25 오후 4:01:39  7) 기타자료2 7-4) 누락TAG CHECK 내역
''''''''''2012-04-25 오후 4:01:39접수집계_Display 시작
''''''''''2012-04-25 오후 4:01:39일일마감 종료
'
'
'    '----------------------------------------------------------------
'    ' 1. 매출 -> 1-2) 출고수량 구하기
'    '----------------------------------------------------------------
'    Debug.Print Now & "  1. 매출 -> 1-2) 출고수량 구하기 시작"
'
'    Query = "SELECT    ISNULL(COUNT(택번호),0)"
'    Query = Query & " FROM TB_입출고"
'    Query = Query & " WHERE 출고일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND (판매취소 <> 'Y')"
'
'    txtNum02.Value = Recordset_Result(Query) '출고수량
'
'
'    '----------------------------------------------------------------
'    ' 2. 선불결제 2-1) 현금반환/ 현금결제 구하기
'    '----------------------------------------------------------------
'    Debug.Print Now & "  2. 선불결제 2-1) 현금반환/ 현금결제 구하기 시작"
'
'    Query = "SELECT    ISNULL(SUM(접수금액),0) * -1"
'    Query = Query & " FROM TB_매출"
'    Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND 적요 LIKE '%현금반환%' "
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    txtCost28.Value = ADORs(0)  ' 금액
'    ADORs.Close:    Set ADORs = Nothing
'
'    '----------------------------------------------------------------
'    ' 2. 선불결제 2-2) 현금결제 구하기
'    '----------------------------------------------------------------
'    Debug.Print Now & "  2. 선불결제 2-2) 현금결제 구하기 시작"
'    ' 매출중에 접수일자에 발생한 건만 처리
'    ' 판매 취소한 내역도 빼주어야 한다. ( 현금입금에 -가 들어 가기 때문에 바로 처리가 가능하다.)
'    Query = "SELECT ISNULL(SUM(현금입금),0)"
'    Query = Query & " FROM TB_매출 "
'    Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND 접수금액 <> 0"
'    'Query = Query & "   AND NOT 적요 LIKE '%판매취소%' "
'    Query = Query & "   AND NOT 적요 LIKE '%미수금액 입금%'"
'
'    txtCost02.Value = Recordset_Result(Query) '
'
'    '----------------------------------------------------------------
'    ' 2. 선불결제 2-3) 카드결제 건수 / 금액 구하기
'    ' 건수= 승인 + 취소 , 금액 = 승인 + 취소
'    '----------------------------------------------------------------
'    Debug.Print Now & "  2. 선불결제 2-3) 카드결제 건수 / 금액 구하기"
'
'    Query = "SELECT    ISNULL(COUNT(카드입금),0)"
'    Query = Query & ", ISNULL(SUM(카드입금),0)"
'    Query = Query & " FROM TB_매출 "
'    Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND 카드입금 <> 0" ' 카드 금결제가 아닌 경우도 0원이 들어간다.
''   Query = Query & "   AND 접수금액 > 0"
''   Query = Query & "   AND NOT 적요 LIKE '%판매취소%' "
'    Query = Query & "   AND NOT 적요 LIKE '%미수금액 입금%'"
'    Query = Query & "   AND NOT 적요 LIKE '%반품환불%'"
'    Query = Query & "   AND NOT 적요 LIKE '%세탁환불%'"
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    txtNum03.Value = ADORs(0)
'    txtCost03.Value = ADORs(1)
'
'    ADORs.Close:    Set ADORs = Nothing
''    ' 판매 취소한 내역도 빼주어야 한다. ( 카드입금에 -금액이 들어가지 않고 승인취소 전표에 -금액을 일괄 적용한다.)
''    ' 그렇기 때문에 카드 판매 취소를 구하기가 힘들다 ㅡㅡ
''
''
''    ' 건수 구하기
''    Query = "SELECT    ISNULL(COUNT(A.접수번호),0) FROM "
''    Query = Query & " ( SELECT 접수번호 "
''    Query = Query & " FROM TB_매출 "
''    Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
''    Query = Query & "   AND 카드입금 = 0"
''    Query = Query & "   AND 접수금액 < 0"
''    Query = Query & "   AND 현금입금 = 0"
''    Query = Query & "   AND 적요 LIKE '%판매취소%' "
''    Query = Query & " GROUP BY 접수번호 ) A "
''    Set ADORs = New ADODB.Recordset
''    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
''
''    txtNum03.Value = txtNum03.Value - Recordset_Result(Query)
''
''    ' 금액 구하기
''    Query = "SELECT    ISNULL(SUM(접수금액),0)"
''    Query = Query & " FROM TB_매출 "
''    Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
''    Query = Query & "   AND 카드입금 = 0"
''    Query = Query & "   AND 접수금액 < 0"
''    Query = Query & "   AND 현금입금 = 0"
''    Query = Query & "   AND 적요 LIKE '%판매취소%' "
''    Set ADORs = New ADODB.Recordset
''    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
''
''    txtCost03.Value = txtCost03.Value + Recordset_Result(Query) '- 금액이 날라오기 때문에 더해준다.
'
'    '--------------------------------------------------------------------
'    ' 2. 선불결제 2-4) 발생/사용/삭제 마일리지
'    '--------------------------------------------------------------------
'    Debug.Print Now & "  2. 선불결제 2-4) 발생/사용/삭제 마일리지"
'
'    Query = "SELECT    ISNULL(SUM(발생마일리지),0)"
'    Query = Query & ", ISNULL(SUM(사용마일리지),0)"
'    Query = Query & ", ISNULL(SUM(삭제마일리지),0)"
'    Query = Query & " FROM TB_매출"
'    Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    txtCost18.Value = ADORs(0) '
'    txtCost19.Value = ADORs(1) '사용마일리지
'    txtCost20.Value = ADORs(2) '
'
'    txtCost27.Value = txtCost19.Value '사용마일리지
'
'    ADORs.Close:    Set ADORs = Nothing
'
'    '--------------------------------------------------------------------
'    ' 2. 선불결제 2-5) 쿠폰 사용 건수/ 금액
'    '--------------------------------------------------------------------
'    Debug.Print Now & "  2. 선불결제 2-5) 쿠폰 사용 건수/ 금액"
'
'    Query = "SELECT    ISNULL(SUM(쿠폰입금),0)"
'    Query = Query & ", ISNULL(COUNT(쿠폰번호),0)"
'    Query = Query & " FROM TB_매출"
'    Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND 쿠폰입금 > 0"
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    txtNum12.Value = ADORs(1)   ' 수량
'    txtCost21.Value = ADORs(0)  ' 금액
'
'    ADORs.Close:    Set ADORs = Nothing
'
'    '----------------------------------------------------------------
'    ' 2. 선불결제 2-6) 미수금 금액
'    '----------------------------------------------------------------
'    Debug.Print Now & "  2. 선불결제 2-6) 미수금 금액"
'
'    Query = "SELECT ISNULL(SUM(접수금액),0) - ISNULL(SUM(입금합계),0) AS 미수금 "
'    Query = Query & " FROM TB_매출"
'    Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
'    'Query = Query & "   AND 접수금액 > 0"
'    Query = Query & "   AND NOT 적요  LIKE '%미수금액 입금%'"
'    Query = Query & "   AND NOT 적요  LIKE '%반품환불%'"
'    Query = Query & "   AND NOT 적요  LIKE '%세탁환불%'"
'
'    txtCost04.Value = Recordset_Result(Query) '
'
'    '----------------------------------------------------------------
'    ' 3. 미수결제 3-1) 미수금 수금 현금결제 구하기
'    '----------------------------------------------------------------
'    Debug.Print Now & "  3. 미수결제 3-1) 미수금 수금 현금결제 구하기"
'
'    Query = "SELECT ISNULL(SUM(현금입금),0)"
'    Query = Query & " FROM TB_매출"
'    Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND 접수금액 = 0"
'    Query = Query & "   AND 적요  LIKE '%미수금액 입금%'"
'
'    txtCost05.Value = Recordset_Result(Query) '
'
'    '----------------------------------------------------------------
'    ' 3. 미수결제 3-2) 미수금 수금 카드결제 구하기
'    '----------------------------------------------------------------
'    Debug.Print Now & "  3. 미수결제 3-2) 미수금 수금 카드결제 구하기"
'
'    Query = "SELECT    ISNULL(COUNT(카드입금),0)"
'    Query = Query & ", ISNULL(SUM(카드입금),0)"
'    Query = Query & " FROM TB_매출"
'    Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND 카드입금 <> 0"
'    Query = Query & "   AND 접수금액 = 0" ' 판매취소시 0원으로 들어온다.
'    Query = Query & "   AND 적요  LIKE '%미수금액 입금%'"
'
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    txtNum04.Value = ADORs(0)
'    txtCost06.Value = ADORs(1)
'
'    ADORs.Close:    Set ADORs = Nothing
'
'    '----------------------------------------------------------------
'    ' 4. 결제합계 현금/ 카드 결제
'    '----------------------------------------------------------------
'    Debug.Print Now & "  4. 결제합계 현금/ 카드 결제"
'
'    txtCost07.Value = txtCost02.Value + txtCost05.Value '현금결제합계
'    txtNum05.Value = txtNum03.Value + txtNum04.Value    '카드결제건수 합계
'    txtCost08.Value = txtCost03.Value + txtCost06.Value '카드결제합계
'
'    '----------------------------------------------------------------
'    ' 5. 마진 5-1) 사용마일리지  2. 선불결제 2-4)에서 처리
'    '----------------------------------------------------------------
'    '----------------------------------------------------------------
'    ' 5. 마진 5-2) 가맹점 마진
'    '----------------------------------------------------------------
'    Debug.Print Now & "  5. 마진 5-2) 가맹점 마진"
'
'    Query = "SELECT ISNULL(SUM(금액 * 세탁마진/100),0)"
'    Query = Query & " FROM TB_입출고"
'    Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND 내용 NOT LIKE '%수%'"                            '수선 제외
'    Query = Query & "   AND (판매취소 <> 'Y')"
'
'    txtCost09.Value = Recordset_Result(Query)
'
'    ' 마일리지 적용
'    If txtCost19.Value > 0 Then
'        txtCost09.Value = txtCost09.Value - CLng(txtCost19.Value * 0.4) '가맹점  지사:가맹점(6:4)로 빼준다.
'        'txtCost10.Value = txtCost10.Value - CLng(txtCost19.Value * 0.6) '지사
'    End If
'
'    '쿠폰 사용이 있는 경우
'    If txtCost21.Value > 0 And 마감일자 <= "2011-12-31" Then
'        txtCost09.Value = txtCost09.Value - CLng(1200 * txtNum12.Value * 0.4) '가맹점
'        'txtCost10.Value = txtCost10.Value - CLng(1200 * txtNum12.Value * 0.6) '지사
'    End If
'    '----------------------------------------------------------------
'    ' 5. 마진 5-2) 지사 마진
'    '----------------------------------------------------------------
'    txtCost10.Value = (txtCost01.Value - txtCost19.Value) - txtCost09.Value     ' 지사 마진 = (접수금액 -마일리지) - 가맹점마진
'
'
'
'
'    '--------------------------------------------------------------
'    ' 6. 기타자료 6-1) 수선수량 계산
'    '--------------------------------------------------------------
'    Debug.Print Now & "  6. 기타자료 6-1) 수선수량 계산"
'
'    Query = "SELECT    ISNULL(COUNT(택번호),0)"
'    Query = Query & ", ISNULL(SUM(수선금액),0)"
'    Query = Query & " FROM TB_입출고"
'    Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND (내용  = '드수' OR 내용 = '수') "
'    Query = Query & "   AND (판매취소 <> 'Y')"
'
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    txtNum06.Value = ADORs(0)
'    txtCost11.Value = ADORs(1)
'
'    ADORs.Close:    Set ADORs = Nothing
'
'    '----------------------------------------------------------------
'    ' 6. 기타자료 6-2) 재세탁수량 계산
'    '----------------------------------------------------------------
'    Debug.Print Now & "  6. 기타자료 6-2) 재세탁수량 계산"
'
'    Query = "SELECT ISNULL(COUNT(택번호),0)"
'    Query = Query & " FROM TB_입출고"
'    Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND 내용     = '드재'"
'    Query = Query & "   AND (판매취소 <> 'Y')"
'
'    txtNum07.Value = Recordset_Result(Query)
'
'    '--------------------------------------------------------------------
'    ' 6. 기타자료 6-3) 운동화 매출을 불러온다.
'    '--------------------------------------------------------------------
'    Debug.Print Now & "  6. 기타자료 6-3) 운동화 매출을 불러온다."
'
'    Query = "SELECT    ISNULL(COUNT(의류코드),0)" '운동화건수
'    Query = Query & ", ISNULL(SUM(금액),0)"       '운동화금액
'    Query = Query & " FROM TB_입출고 "
'    Query = Query & " WHERE SUBSTRING(의류코드,1,2) = 'a0'"
'    Query = Query & "   AND 접수일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND (판매취소 <> 'Y')"
'
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    txtNum08.Value = ADORs(0)
'    txtCost13.Value = ADORs(1)
'
'    ADORs.Close:        Set ADORs = Nothing
'
'    '--------------------------------------------------------------------
'    ' 6. 기타자료 6-4) 가죽/무스탕 매출을 불러온다.
'    '--------------------------------------------------------------------
'    Debug.Print Now & "  6. 기타자료 6-4) 가죽/무스탕 매출을 불러온다."
'
'    Query = "SELECT    ISNULL(COUNT(의류코드),0)" '가죽건수
'    Query = Query & ", ISNULL(SUM(금액),0)"       '가죽금액
'    Query = Query & " FROM TB_입출고 "
'    Query = Query & " WHERE SUBSTRING(의류코드,1,2) IN ('b0','n0')"
'    Query = Query & "   AND 접수일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND (판매취소 <> 'Y')"
'
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    txtNum09.Value = ADORs(0)
'    txtCost14.Value = ADORs(1)
'
'    ADORs.Close:    Set ADORs = Nothing
'
'    '--------------------------------------------------------------------
'    ' 6. 기타자료 6-5) 카페트 매출을 불러온다.
'    '--------------------------------------------------------------------
'    Debug.Print Now & "  6. 기타자료 6-5) 카페트 매출을 불러온다."
'
'    Query = "SELECT    ISNULL(COUNT(의류코드),0)" '카페트건수
'    Query = Query & ", ISNULL(SUM(금액),0) "      '카페트금액
'    Query = Query & " FROM TB_입출고 "
'    Query = Query & " WHERE SUBSTRING(의류코드,1,2) = 'x0'"
'    Query = Query & "   AND 접수일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND (판매취소 <> 'Y')"
'
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    txtNum10.Value = ADORs(0)
'    txtCost15.Value = ADORs(1)
'
'    ADORs.Close
'    Set ADORs = Nothing
'
'    '----------------------------------------------------------------
'    ' 6. 기타자료 6-6) 반품 내역
'    '----------------------------------------------------------------
'    Debug.Print Now & "  6. 기타자료 6-6) 반품 내역"
'
'    Query = "SELECT ISNULL(COUNT(택번호),0)"
'    Query = Query & " FROM TB_입출고"
'    Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND 내용     = '%반%'"
'    Query = Query & "   AND (판매취소 <> 'Y')"
'
'    txtNum11.Value = Recordset_Result(Query)
'
'    '----------------------------------------------------------------
'    ' 6. 기타자료 6-7) 외주 마진
'    '----------------------------------------------------------------
'    Debug.Print Now & "  6. 기타자료 6-7) 외주 마진"
'
'    Query = "SELECT ISNULL(SUM(금액*외주마진/100),0)"
'    Query = Query & " FROM TB_입출고"
'    Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND SUBSTRING(의류코드,1,1) = 'a'"   '운동화
'    Query = Query & "   AND 내용 NOT LIKE '%수%'"            '수선 제외
'    Query = Query & "   AND (판매취소 <> 'Y')"
'
'    txtCost17.Value = Recordset_Result(Query)
'
'    '----------------------------------------------------------------
'    ' 6. 기타자료 6-8) 마일리지 2. 선불결제 2-4)에서 처리
'    '----------------------------------------------------------------
'    '----------------------------------------------------------------
'    ' 6. 기타자료 6-9) 마일리지 2. 선불결제 2-4)에서 처리
'    '----------------------------------------------------------------
'
'    '----------------------------------------------------------------
'    ' 7) 기타자료2 7-1) 판매취소 내역
'    '----------------------------------------------------------------
'    Debug.Print Now & "  7) 기타자료2 7-1) 판매취소 내역"
'
'    Query = "SELECT    ISNULL(COUNT(택번호),0)"
'    Query = Query & ", ISNULL(SUM(금액),0)"
'    Query = Query & " FROM TB_입출고"
'    Query = Query & " WHERE SUBSTRING(판매취소일자,1,10) = '" & 마감일자 & "'"
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    txtNum13.Value = ADORs(0)
'    txtCost22.Value = ADORs(1)
'
'    ADORs.Close:    Set ADORs = Nothing
'
'    If txtNum13.Value > 0 Then
'        Query = "SELECT    택번호"
'        Query = Query & " FROM TB_입출고"
'        Query = Query & " WHERE SUBSTRING(판매취소일자,1,10) = '" & 마감일자 & "'"
'        Query = Query & " ORDER BY 택번호 ASC"
'
'        Call Get_택번호(Query, cboCancel)
'    End If
'
'    '----------------------------------------------------------------
'    ' 7) 기타자료2 7-2) 반품환불 내역
'    '----------------------------------------------------------------
'    Debug.Print Now & "  7) 기타자료2 7-2) 반품환불 내역"
'
'    Query = "SELECT    ISNULL(COUNT(택번호),0)"
'    Query = Query & ", ISNULL(SUM(금액),0)"
'    Query = Query & ", ISNULL(SUM(금액*(100-세탁마진)/100),0)"
'    Query = Query & " FROM TB_입출고"
'    Query = Query & " WHERE SUBSTRING(반품환불일자,1,10) = '" & 마감일자 & "'"
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    txtNum14.Value = ADORs(0)
'    txtCost23.Value = ADORs(1)
'    txtCost29.Value = ADORs(2)
'
'    ADORs.Close:    Set ADORs = Nothing
'
'    If txtNum14.Value > 0 Then
'        Query = "SELECT    택번호"
'        Query = Query & " FROM TB_입출고"
'        Query = Query & " WHERE SUBSTRING(반품환불일자,1,10) = '" & 마감일자 & "'"
'
'        Call Get_택번호(Query, cboReturn)
'    End If
'
'    '----------------------------------------------------------------
'    ' 7) 기타자료2 7-3) 세탁환불 내역
'    '----------------------------------------------------------------
'    Debug.Print Now & "  7) 기타자료2 7-3) 세탁환불 내역"
'
'    Query = "SELECT    ISNULL(COUNT(택번호),0)"
'    Query = Query & ", ISNULL(SUM(금액),0)"
'    Query = Query & ", ISNULL(SUM(금액*(100-세탁마진)/100),0)"
'    Query = Query & " FROM TB_입출고"
'    Query = Query & " WHERE SUBSTRING(세탁환불일자,1,10) = '" & 마감일자 & "'"
'
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    txtNum15.Value = ADORs(0)
'    txtCost24.Value = ADORs(1)
'    txtCost30.Value = ADORs(2)
'
'    ADORs.Close:    Set ADORs = Nothing
'
'    If txtNum15.Value > 0 Then
'        Query = "SELECT    택번호"
'        Query = Query & " FROM TB_입출고"
'        Query = Query & " WHERE SUBSTRING(세탁환불일자,1,10) = '" & 마감일자 & "'"
'
'        Call Get_택번호(Query, cboRepay)
'    End If
'
'    '--------------------------------------------------------------------
'    ' 7) 기타자료2 7-4) 누락TAG CHECK 내역
'    '--------------------------------------------------------------------
'    Debug.Print Now & "  7) 기타자료2 7-4) 누락TAG CHECK 내역"
'
'    Dim 시작택번호   As String
'    Dim 마지막택번호 As String
'
'    Dim 택번호 As String
'    Dim tmpTAG As String
'
'    Query = "SELECT    MIN(택번호)"
'    Query = Query & ", MAX(택번호)"
'    Query = Query & " FROM TB_입출고 "
'    Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND (판매취소 <> 'Y')"
'
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    If Not ADORs.EOF Then
'        시작택번호 = ADORs(0)
'        마지막택번호 = ADORs(1)
'    End If
'
'    ADORs.Close:    Set ADORs = Nothing
'    cboMissTag.Clear
'
'    Dim iLoop As Long
'
'    Query = "SELECT 택번호 FROM TB_입출고"
'    Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND (판매취소 <> 'Y')"
'
'    Query = Query & " ORDER BY 택번호 ASC"
'
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    iLoop = 0
'
'    택번호 = ""
'    tmpTAG = ""
'
'    If Val(마지막택번호) - Val(시작택번호) < 5000 Then
'        Do Until ADORs.EOF
'            If tmpTAG = "" Then
'                tmpTAG = ADORs!택번호
'            Else
'                Do Until Format(CLng(tmpTAG) + 1, "000000000") >= ADORs!택번호
'                    cboMissTag.AddItem Format(CLng(tmpTAG) + 1, "000-00-0000")
'
'                    tmpTAG = Format(CLng(tmpTAG) + 1, "000000000")
'
'                    '100 개가 넘으면 빠져 나옴
'                    If iLoop >= 100 Then
'                        cboMissTag.AddItem "Err"
'
'                        Exit Do
'                    End If
'
'                    iLoop = iLoop + 1
'                Loop
'
'                tmpTAG = Format(CLng(tmpTAG) + 1, "000000000")
'            End If
'
'            ADORs.MoveNext
'        Loop
'        ADORs.Close
'        Set ADORs = Nothing
'
'        If cboMissTag.ListCount = 0 Then
'            txtNum16.Value = 0
'        Else
'            txtNum16.Value = cboMissTag.ListCount - 1
'        End If
'    End If
'
'    pnlTAG(0).Caption = Format(시작택번호, "000-00-0000") & ""
'    pnlTAG(1).Caption = Format(마지막택번호, "000-00-0000") & ""
'
'    '--------------------------------------------------------------------
'    ' 7) 기타자료2 7-4) 삼성 카드 할인 내용 추가
'    '--------------------------------------------------------------------
'    Debug.Print Now & "  7) 기타자료2 7-4) 삼성 카드 할인 내용 추가"
'
'    Dim 삼성카드고객수   As Long
'    Dim 삼성카드할인건수 As Long
'    Dim 삼성카드할인금액 As Long
'
'    삼성카드고객수 = 0
'    삼성카드할인건수 = 0
'    삼성카드할인금액 = 0
'
'    Query = "SELECT    고객코드"
'    Query = Query & ", ISNULL(COUNT(금액),0)"
'    Query = Query & ", ISNULL(SUM(금액),0)"
'    Query = Query & " FROM TB_입출고"
'    Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND 내용  LIKE '%삼%'"
'    Query = Query & "   AND (판매취소 <> 'Y')"
'
'    Query = Query & " GROUP BY 고객코드"
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    Do Until ADORs.EOF
'        삼성카드고객수 = 삼성카드고객수 + 1
'
'        삼성카드할인건수 = 삼성카드할인건수 + ADORs(0)
'        삼성카드할인금액 = 삼성카드할인금액 + ADORs(1)
'
'        ADORs.MoveNext
'    Loop
'    ADORs.Close
'    Set ADORs = Nothing
'
'    txtCost26.Value = 삼성카드할인금액
'    txtNum17.Value = 삼성카드할인건수
'    txtNum18.Value = 삼성카드고객수
'
'
'    '--------------------------------------------------------------------
'    ' 8) 반품환불, 세탁환불 확정시 마진 처리
'    '--------------------------------------------------------------------
''    If txtCost29.Value <> "" Then
''        txtCost09.Value = txtCost09.Value - txtCost29.Value                     '가맹점 반품환불
''
''        txtCost10.Value = txtCost10.Value - (txtCost23.Value - txtCost29.Value) '지사   반품환불
''    End If
''
''    If txtCost30.Value <> "" Then
''        txtCost09.Value = txtCost09.Value - txtCost30.Value                     '가맹점 세탁환불
''
''        txtCost10.Value = txtCost10.Value - (txtCost24.Value - txtCost30.Value) '지사   세탁환불
''    End If
'
'
''    '--------------------------------------------------------------------
''    ' 8) 지사 정산 참고 사항
''    '--------------------------------------------------------------------
''    Debug.Print Now & " 8) 지사 정산 참고 사항  1. 로열티 정보"
''
''    pnlData(24).Caption = pnlData(24).Caption & " " & 가맹점정보.로열티여부1
''    If 가맹점정보.로열티여부1 = "Y" Then txtMaster(11).Value = CDbl(txtCost01.Value) * CDbl(가맹점정보.로열티비율1)
''
''    pnlData(25).Caption = pnlData(25).Caption & " " & 가맹점정보.로열티여부2
''    If 가맹점정보.로열티여부2 = "Y" Then txtMaster(12).Value = CDbl(txtCost10.Value) * CDbl(가맹점정보.로열티비율2)
''
''
''    pnlData(46).Caption = pnlData(46).Caption & " " & 가맹점정보.수수료지원여부
''    txtCard(0).Value = txtNum05.Value
''    txtCard(1).Value = txtCost08.Value
''    txtCard(2).Value = CDbl(txtCard(1).Value) * CDbl(가맹점정보.수수료지원비율)
''
''    txtCard(3).Value = txtNum05.Value
''    txtCard(4).Value = txtCost08.Value
''    txtCard(5).Value = CDbl(txtCard(4).Value) * CDbl(가맹점정보.수수료지원비율)
''
''
''    txtMaster(0).Value = txtCost09.Value
''    txtMaster(1).Value = txtCard(2).Value - txtCard(5).Value
''    txtMaster(2).Value = txtCost29.Value + txtCost30.Value
''    txtMaster(3).Value = txtMaster(0).Value - (txtMaster(1).Value + txtMaster(2).Value)
'
'    Screen.MousePointer = 0
'    pnlProg.Visible = False
'    DoEvents
'
'    Exit Sub
'
'ErrRtn:
'    Screen.MousePointer = 0
'    pnlProg.Visible = False
'End Sub

'
'Private Sub 일일마감_Proc()
'    Dim 접수금액   As Long
'    Dim 가맹점마진 As Long
'    Dim 외주마진   As Long
'
'    Dim 미수금액 As Long
'    Dim tmpData  As String
'
'    On Error GoTo ErrRtn
'
'    Screen.MousePointer = 11
'
'    pnlProg.Left = 90
'    pnlProg.Top = 1305
'    pnlProg.Visible = True
'    DoEvents
'
'
'    '컨트롤 초기화
'    Dim ctrl As Control
'    Dim txt  As sidbEdit
'
'    For Each ctrl In Me.Controls
'        If TypeOf ctrl Is sidbEdit Then
'            ctrl.Value = 0
'        End If
'    Next ctrl
'
'    마감일자 = Format(dtpDay.Value, "YYYY-MM-DD")
'
'    Debug.Print "마감작업 : " & Now & " 1. 접수수량, 접수금액 구하기"
'    '----------------------------------------------------------------
'    ' 1. 접수수량, 접수금액 구하기
'    '----------------------------------------------------------------
'    Query = "SELECT    ISNULL(COUNT(택번호),0)"
'    Query = Query & ", ISNULL(SUM(금액),0)"
'    Query = Query & " FROM TB_입출고"
'    Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND (판매취소 <> 'Y')"
'
'    'Query = Query & "   AND ((판매취소 <> 'Y')"
'    ''Query = Query & "   AND ((판매취소 = '') AND (판매취소일자 IS NULL OR 판매취소일자 = '')"
'    'Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
'    'Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
'
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    txtNum01.Value = ADORs(0)  '접수건수
'    txtCost01.Value = ADORs(1) '접수금액
'
'    ADORs.Close
'    Set ADORs = Nothing
'
'
'    Debug.Print "마감작업 : " & Now & " 2. 출고수량 구하기"
'    '----------------------------------------------------------------
'    ' 2. 출고수량 구하기
'    '----------------------------------------------------------------
'    Query = "SELECT    ISNULL(COUNT(택번호),0)"
'    Query = Query & " FROM TB_입출고"
'    Query = Query & " WHERE 출고일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND (판매취소 <> 'Y')"
'
'    'Query = Query & "   AND ((판매취소 <> 'Y')"
'    ''Query = Query & "   AND ((판매취소 = '') AND (판매취소일자 IS NULL OR 판매취소일자 = '')"
'    'Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
'    'Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
'
'    txtNum02.Value = Recordset_Result(Query) '출고수량
'
'
'    Debug.Print "마감작업 : " & Now & " 2-1) 현금결제 구하기"
'    '----------------------------------------------------------------
'    ' 2-1) 현금결제 구하기
'    '----------------------------------------------------------------
'    '선불결제
'    'Query = "SELECT ISNULL(SUM(현금입금),0)"
'    'Query = Query & " FROM TB_매출"
'    'Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
'    'Query = Query & "   AND 접수금액 <> 0"
'
'    '매출중에 접수일자에 발생한 건만 처리
'    Query = "SELECT SUM(A.현금입금)"
'    Query = Query & " FROM TB_매출 AS A LEFT OUTER JOIN (SELECT DISTINCT 접수일자"
'    Query = Query & "                                         , 고객코드"
'    Query = Query & "                                         , 접수번호"
'    Query = Query & "                                    FROM TB_입출고"
'    Query = Query & "                  WHERE 접수일자 = '" & 마감일자 & "'"
'    Query = Query & "            ) AS B ON A.고객코드 = B.고객코드"
'    Query = Query & "                               AND A.접수번호 = B.접수번호"
'    Query = Query & " WHERE A.매출일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND B.접수일자 IS NOT NULL"
'    Query = Query & "   AND A.접수번호 <> 0"
'    'Query = Query & "   AND A.접수금액 <> 0"
'
'    txtCost02.Value = Recordset_Result(Query) '
'
'    Debug.Print "마감작업 : " & Now & " 2-2) 미수입금"
'    '----------------------------------------------------------------
'    ' 2-2) 미수입금
'    '----------------------------------------------------------------
'    Query = "SELECT ISNULL(SUM(현금입금),0)"
'    Query = Query & " FROM TB_매출"
'    Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND 접수금액 = 0"
'    Query = Query & "   AND 적요  LIKE '%미수금액 입금%'"
'
'    txtCost05.Value = Recordset_Result(Query) '
'
'
'    Debug.Print "마감작업 : " & Now & " 2-3) 카드결제 구하기"
'    '----------------------------------------------------------------
'    ' 2-3) 카드결제 구하기
'    '----------------------------------------------------------------
'    'Query = "SELECT    ISNULL(COUNT(카드입금),0)"
'    'Query = Query & ", ISNULL(SUM(카드입금),0)"
'    'Query = Query & " FROM TB_매출"
'    'Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
'    'Query = Query & "   AND 카드입금 > 0"
'    'Query = Query & "   AND 접수금액 <> 0"
'
'    Query = "SELECT    ISNULL(COUNT(A.카드입금),0)"
'    Query = Query & ", ISNULL(SUM(A.카드입금),0)"
'    Query = Query & " FROM TB_매출 AS A LEFT OUTER JOIN (SELECT DISTINCT 접수일자"
'    Query = Query & "                                         , 고객코드"
'    Query = Query & "                                         , 접수번호"
'    Query = Query & "                                    FROM TB_입출고"
'    Query = Query & "                  WHERE 접수일자 = '" & 마감일자 & "'"
'    Query = Query & "            ) AS B ON A.고객코드 = B.고객코드"
'    Query = Query & "                  AND A.접수번호 = B.접수번호"
'    Query = Query & " WHERE A.매출일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND B.접수일자 IS NOT NULL"
'    Query = Query & "   AND A.카드입금 > 0"
'    Query = Query & "   AND A.접수번호 <> 0"
'    Query = Query & "   AND NOT A.적요   LIKE  '%미수금액 입금%'"
'    'Query = Query & "   AND A.접수금액 <> 0"
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    txtNum03.Value = ADORs(0)
'    txtCost03.Value = ADORs(1)
'
'    ADORs.Close
'    Set ADORs = Nothing
'
'    Debug.Print "마감작업 : " & Now & " 2-4) 카드결제 취소"
'    ' 2-4) 카드결제 취소
'    'Query = "SELECT    ISNULL(COUNT(카드입금),0)"
'    'Query = Query & ", ISNULL(SUM(카드입금) * -1,0)"
'    'Query = Query & " FROM TB_매출"
'    'Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
'    'Query = Query & "   AND 카드입금 < 0"
'    'Query = Query & "   AND 접수금액 <> 0"
'
'    Query = "SELECT    ISNULL(COUNT(A.카드입금),0)"
'    Query = Query & ", ISNULL(SUM(A.카드입금),0)"
'    Query = Query & " FROM TB_매출 AS A LEFT OUTER JOIN (SELECT DISTINCT 접수일자"
'    Query = Query & "                                         , 고객코드"
'    Query = Query & "                                         , 접수번호"
'    Query = Query & "                                    FROM TB_입출고"
'    Query = Query & "                  WHERE 접수일자 = '" & 마감일자 & "'"
'    Query = Query & "            ) AS B ON A.고객코드 = B.고객코드"
'    Query = Query & "                  AND A.접수번호 = B.접수번호"
'    Query = Query & " WHERE A.매출일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND B.접수일자 IS NOT NULL"
'    Query = Query & "   AND A.카드입금 < 0"
'    Query = Query & "   AND A.접수번호 <> 0"
'    Query = Query & "   AND NOT A.적요   LIKE '%미수금액 입금%'"
'
'    'Query = Query & "   AND A.접수금액 <> 0"
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    txtNum03.Value = txtNum03.Value - ADORs(0)
'    txtCost03.Value = txtCost03.Value - (ADORs(1) * -1)
'
'    ADORs.Close
'    Set ADORs = Nothing
'
'    Debug.Print "마감작업 : " & Now & " 2-5) 카드결제 구하기"
'    '----------------------------------------------------------------
'    ' 2-5) 미수금 수금 카드결제 구하기
'    '----------------------------------------------------------------
'    Query = "SELECT    ISNULL(COUNT(카드입금),0)"
'    Query = Query & ", ISNULL(SUM(카드입금),0)"
'    Query = Query & " FROM TB_매출"
'    Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND 카드입금 <> 0"
'    Query = Query & "   AND 접수금액 = 0" ' 판매취소시 0원으로 들어온다.
'    Query = Query & "   AND 적요  LIKE '%미수금액 입금%'"
'
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    txtNum04.Value = ADORs(0)
'    txtCost06.Value = ADORs(1)
'
'    ADORs.Close
'    Set ADORs = Nothing
'
'
'    txtCost07.Value = txtCost02.Value + txtCost05.Value '현금결제합계
'    txtNum05.Value = txtNum03.Value + txtNum04.Value    '카드결제건수 합계
'    txtCost08.Value = txtCost03.Value + txtCost06.Value '카드결제합계
'
'    Debug.Print "마감작업 : " & Now & " 3-1) 가맹점 마진"
'    '----------------------------------------------------------------
'    ' 3-1) 가맹점 마진
'    '----------------------------------------------------------------
'    Query = "SELECT ISNULL(SUM(금액 * 세탁마진/100),0)"
'    Query = Query & " FROM TB_입출고"
'    Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
'    'Query = Query & "   AND SUBSTRING(의류코드,1,1) <> 'a'"                 '운동화
'    'Query = Query & "   AND (내용 LIKE '%세%' OR 내용 LIKE '%건%'  OR 내용 LIKE '%습%')"
'    Query = Query & "   AND 내용 NOT LIKE '%수%'"                            '수선 제외
'    Query = Query & "   AND (판매취소 <> 'Y')"
'
'    'Query = Query & "   AND ((판매취소 <> 'Y')"
'    ''Query = Query & "   AND ((판매취소 = '') AND (판매취소일자 IS NULL OR 판매취소일자 = '')"
'    'Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
'    'Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
'
'    txtCost09.Value = Recordset_Result(Query)
'
'
'    Debug.Print "마감작업 : " & Now & " 3-2) 외주 마진"
'    '----------------------------------------------------------------
'    ' 3-2) 외주 마진
'    '----------------------------------------------------------------
'    Query = "SELECT ISNULL(SUM(금액*외주마진/100),0)"
'    Query = Query & " FROM TB_입출고"
'    Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND SUBSTRING(의류코드,1,1) = 'a'"                 '운동화
'    'Query = Query & "   AND (내용 LIKE '%세%' OR 내용 LIKE '%건%'  OR 내용 LIKE '%습%')"
'    Query = Query & "   AND 내용 NOT LIKE '%수%'"                            '수선 제외
'    Query = Query & "   AND (판매취소 <> 'Y')"
'
'    'Query = Query & "   AND ((판매취소 <> 'Y')"
'    ''Query = Query & "   AND ((판매취소 = '') AND (판매취소일자 IS NULL OR 판매취소일자 = '')"
'    'Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
'    'Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
'
'    txtCost17.Value = Recordset_Result(Query)
'
'    'txtCost10.Value = txtCost01.Value - txtCost09.Value - txtCost17.Value ' 지사 마진 = 접수금액 - 가맹점마진 - 외주마진 '
'    txtCost10.Value = txtCost01.Value - txtCost09.Value                    ' 지사 마진 = 접수금액 - 가맹점마진
'
'    Debug.Print "마감작업 : " & Now & " 6) 판매취소 계산"
'    '----------------------------------------------------------------
'    ' 6) 판매취소 계산
'    '----------------------------------------------------------------
'    Query = "SELECT    ISNULL(COUNT(택번호),0)"
'    Query = Query & ", ISNULL(SUM(금액),0)"
'    Query = Query & " FROM TB_입출고"
'    Query = Query & " WHERE SUBSTRING(판매취소일자,1,10) = '" & 마감일자 & "'"
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    txtNum13.Value = ADORs(0)
'    txtCost22.Value = ADORs(1)
'
'    ADORs.Close
'    Set ADORs = Nothing
'
'
'    Debug.Print "마감작업 : " & Now & " 7) 반품환불 계산"
'    '----------------------------------------------------------------
'    ' 7) 반품환불 계산
'    '----------------------------------------------------------------
'    Query = "SELECT    ISNULL(COUNT(택번호),0)"
'    Query = Query & ", ISNULL(SUM(금액),0)"
'    Query = Query & " FROM TB_입출고"
'    Query = Query & " WHERE SUBSTRING(반품환불일자,1,10) = '" & 마감일자 & "'"
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    txtNum14.Value = ADORs(0)
'    txtCost23.Value = ADORs(1)
'
'    ADORs.Close
'    Set ADORs = Nothing
'
'    Debug.Print "마감작업 : " & Now & " 8) 세탁환불 계산"
'    '----------------------------------------------------------------
'    ' 8) 세탁환불 계산
'    '----------------------------------------------------------------
'    Query = "SELECT    ISNULL(COUNT(택번호),0)"
'    Query = Query & ", ISNULL(SUM(금액),0)"
'    Query = Query & " FROM TB_입출고"
'    Query = Query & " WHERE SUBSTRING(세탁환불일자,1,10) = '" & 마감일자 & "'"
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    txtNum15.Value = ADORs(0)
'    txtCost24.Value = ADORs(1)
'
'    ADORs.Close
'    Set ADORs = Nothing
'
''=================================================================================
'    Debug.Print "마감작업 : " & Now & " 9-1) 수선수량 계산"
'    '--------------------------------------------------------------
'    ' 9-1) 수선수량 계산
'    '--------------------------------------------------------------
'    Query = "SELECT    ISNULL(COUNT(택번호),0)"
'    Query = Query & ", ISNULL(SUM(수선금액),0)"
'    Query = Query & " FROM TB_입출고"
'    Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND (내용  = '드수' OR 내용 = '수') "
'    Query = Query & "   AND (판매취소 <> 'Y')"
'
'    'Query = Query & "   AND ((판매취소 <> 'Y')"
'    ''Query = Query & "   AND ((판매취소 = '') AND (판매취소일자 IS NULL OR 판매취소일자 = '')"
'    'Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
'    'Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
'
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    txtNum06.Value = ADORs(0)
'    txtCost11.Value = ADORs(1)
'
'    ADORs.Close
'    Set ADORs = Nothing
'
'
'    Debug.Print "마감작업 : " & Now & " 9-2) 재세탁수량 계산"
'    '----------------------------------------------------------------
'    ' 9-2) 재세탁수량 계산
'    '----------------------------------------------------------------
'    Query = "SELECT ISNULL(COUNT(택번호),0)"
'    Query = Query & " FROM TB_입출고"
'    Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND 내용     = '드재'"
'    Query = Query & "   AND (판매취소 <> 'Y')"
'
'    'Query = Query & "   AND ((판매취소 <> 'Y')"
'    ''Query = Query & "   AND ((판매취소 = '') AND (판매취소일자 IS NULL OR 판매취소일자 = '')"
'    'Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
'    'Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
'
'    txtNum07.Value = Recordset_Result(Query)
'
'
'    Debug.Print "마감작업 : " & Now & " 9-3) 운동화 매출을 불러온다."
'    '--------------------------------------------------------------------
'    ' 9-3) 운동화 매출을 불러온다.
'    '--------------------------------------------------------------------
'    Query = "SELECT    ISNULL(COUNT(의류코드),0)" '운동화건수
'    Query = Query & ", ISNULL(SUM(금액),0)"       '운동화금액
'    Query = Query & " FROM TB_입출고 "
'    Query = Query & " WHERE SUBSTRING(의류코드,1,2) = 'a0'"
'    Query = Query & "   AND 접수일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND (판매취소 <> 'Y')"
'
'    'Query = Query & "   AND ((판매취소 <> 'Y')"
'    ''Query = Query & "   AND ((판매취소 = '') AND (판매취소일자 IS NULL OR 판매취소일자 = '')"
'    'Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
'    'Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
'
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    txtNum08.Value = ADORs(0)
'    txtCost13.Value = ADORs(1)
'
'    ADORs.Close
'    Set ADORs = Nothing
'
'
'    Debug.Print "마감작업 : " & Now & " 9-4) 가죽/무스탕 매출을 불러온다."
'    '--------------------------------------------------------------------
'    ' 9-4) 가죽/무스탕 매출을 불러온다.
'    '--------------------------------------------------------------------
'    Query = "SELECT    ISNULL(COUNT(의류코드),0)" '가죽건수
'    Query = Query & ", ISNULL(SUM(금액),0)"       '가죽금액
'    Query = Query & " FROM TB_입출고 "
'    Query = Query & " WHERE SUBSTRING(의류코드,1,2) IN ('b0','n0')"
'    Query = Query & "   AND 접수일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND (판매취소 <> 'Y')"
'
'    'Query = Query & "   AND ((판매취소 <> 'Y')"
'    ''Query = Query & "   AND ((판매취소 = '') AND (판매취소일자 IS NULL OR 판매취소일자 = '')"
'    'Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
'    'Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
'
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    txtNum09.Value = ADORs(0)
'    txtCost14.Value = ADORs(1)
'
'    ADORs.Close
'    Set ADORs = Nothing
'
'    Debug.Print "마감작업 : " & Now & " 10) 가죽/무스탕 매출을 불러온다."
'    '--------------------------------------------------------------------
'    ' 10) 카페트 매출을 불러온다.
'    '--------------------------------------------------------------------
'    Query = "SELECT    ISNULL(COUNT(의류코드),0)" '카페트건수
'    Query = Query & ", ISNULL(SUM(금액),0) "      '카페트금액
'    Query = Query & " FROM TB_입출고 "
'    Query = Query & " WHERE SUBSTRING(의류코드,1,2) = 'x0'"
'    Query = Query & "   AND 접수일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND (판매취소 <> 'Y')"
'
'    'Query = Query & "   AND ((판매취소 <> 'Y')"
'    ''Query = Query & "   AND ((판매취소 = '') AND (판매취소일자 IS NULL OR 판매취소일자 = '')"
'    'Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
'    'Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
'
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    txtNum10.Value = ADORs(0)
'    txtCost15.Value = ADORs(1)
'
'    ADORs.Close
'    Set ADORs = Nothing
'
'
'    Debug.Print "마감작업 : " & Now & " 11) 반품수량"
'    '----------------------------------------------------------------
'    ' 11) 반품수량
'    '----------------------------------------------------------------
'    Query = "SELECT ISNULL(COUNT(택번호),0)"
'    Query = Query & " FROM TB_입출고"
'    Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND 내용     = '%반%'"
'    Query = Query & "   AND (판매취소 <> 'Y')"
'
'    'Query = Query & "   AND ((판매취소 <> 'Y')"
'    ''Query = Query & "   AND ((판매취소 = '') AND (판매취소일자 IS NULL OR 판매취소일자 = '')"
'    'Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
'    'Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
'
'    txtNum11.Value = Recordset_Result(Query)
'
'
'    Debug.Print "마감작업 : " & Now & " 12) 발생/사용/삭제 마일리지"
'    '--------------------------------------------------------------------
'    ' 12) 발생/사용/삭제 마일리지
'    '--------------------------------------------------------------------
'    Query = "SELECT    ISNULL(SUM(발생마일리지),0)"
'    Query = Query & ", ISNULL(SUM(사용마일리지),0)"
'    Query = Query & ", ISNULL(SUM(삭제마일리지),0)"
'    Query = Query & " FROM TB_매출"
'    Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    txtCost18.Value = ADORs(0) '
'    txtCost19.Value = ADORs(1) '사용마일리지
'    txtCost20.Value = ADORs(2) '
'
'    txtCost27.Value = txtCost19.Value '사용마일리지
'
'    ADORs.Close
'    Set ADORs = Nothing
'
'    '사용마일지가 있는 경우 지사:가맹점(6:4)로 빼준다.
'    If txtCost19.Value > 0 Then
'        txtCost09.Value = txtCost09.Value - CLng(txtCost19.Value * 0.4) '가맹점
'        txtCost10.Value = txtCost10.Value - CLng(txtCost19.Value * 0.6) '지사
'    End If
'
'    Debug.Print "마감작업 : " & Now & " 13) 쿠폰"
'    '--------------------------------------------------------------------
'    ' 13) 쿠폰
'    '--------------------------------------------------------------------
'    Query = "SELECT    ISNULL(SUM(쿠폰입금),0)"
'    Query = Query & ", ISNULL(COUNT(쿠폰번호),0)"
'    Query = Query & " FROM TB_매출"
'    Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND 쿠폰입금 > 0"
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    txtNum12.Value = ADORs(1)  ' 건수
'    txtCost21.Value = ADORs(0) ' 금액
'
'    ADORs.Close
'    Set ADORs = Nothing
'
'    '쿠폰 사용이 있는 경우
'    If txtCost21.Value > 0 And 마감일자 <= "2011-12-31" Then
'        txtCost09.Value = txtCost09.Value - CLng(1200 * txtNum12.Value * 0.4) '가맹점
'        txtCost10.Value = txtCost10.Value - CLng(1200 * txtNum12.Value * 0.6) '지사
'    End If
'
'    Debug.Print "마감작업 : " & Now & " 14) 판매취소 계산"
'    '----------------------------------------------------------------
'    ' 14) 판매취소 계산
'    '----------------------------------------------------------------
'    Query = "SELECT    택번호"
'    Query = Query & " FROM TB_입출고"
'    Query = Query & " WHERE SUBSTRING(판매취소일자,1,10) = '" & 마감일자 & "'"
'    Query = Query & " ORDER BY 택번호 ASC"
'
'    Call Get_택번호(Query, cboCancel)
'
'
'    Debug.Print "마감작업 : " & Now & " 15) 반품환불 계산"
'    '----------------------------------------------------------------
'    ' 15) 반품환불 계산
'    '----------------------------------------------------------------
'    Query = "SELECT    택번호"
'    Query = Query & " FROM TB_입출고"
'    Query = Query & " WHERE SUBSTRING(반품환불일자,1,10) = '" & 마감일자 & "'"
'
'    Call Get_택번호(Query, cboReturn)
'
'
'    Debug.Print "마감작업 : " & Now & " 16) 세탁환불 계산"
'    '----------------------------------------------------------------
'    ' 16) 세탁환불 계산
'    '----------------------------------------------------------------
'    Query = "SELECT    택번호"
'    Query = Query & " FROM TB_입출고"
'    Query = Query & " WHERE SUBSTRING(세탁환불일자,1,10) = '" & 마감일자 & "'"
'
'    Call Get_택번호(Query, cboRepay)
'
'
'    Debug.Print "마감작업 : " & Now & " 17) 누락TAG CHECK"
'    '--------------------------------------------------------------------
'    ' 17) 누락TAG CHECK
'    '--------------------------------------------------------------------
'    Dim 시작택번호   As String
'    Dim 마지막택번호 As String
'
'    Dim 택번호 As String
'    Dim tmpTAG As String
'
'    Query = "SELECT    MIN(택번호)"
'    Query = Query & ", MAX(택번호)"
'    Query = Query & " FROM TB_입출고 "
'    Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND (판매취소 <> 'Y')"
'
'    'Query = Query & "   AND ((판매취소 <> 'Y')"
'    ''Query = Query & "   AND ((판매취소 = '') AND (판매취소일자 IS NULL OR 판매취소일자 = '')"
'    'Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
'    'Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    If Not ADORs.EOF Then
'        시작택번호 = ADORs(0)
'        마지막택번호 = ADORs(1)
'    End If
'    ADORs.Close
'    Set ADORs = Nothing
'
'    Debug.Print "마감작업 : " & Now & " 18) 누락택"
'    '--------------------------------------------------------------------
'    ' 18) 누락택
'    '--------------------------------------------------------------------
'    cboMissTag.Clear
'
'    Dim iLoop As Long
'
'    Query = "SELECT 택번호 FROM TB_입출고"
'    Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND (판매취소 <> 'Y')"
'
'    'Query = Query & "   AND ((판매취소 <> 'Y')"
'    ''Query = Query & "   AND ((판매취소 = '') AND (판매취소일자 IS NULL OR 판매취소일자 = '')"
'    'Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
'    'Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
'
'    Query = Query & " ORDER BY 택번호 ASC"
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    iLoop = 0
'
'    택번호 = ""
'    tmpTAG = ""
'
'    If Val(마지막택번호) - Val(시작택번호) < 5000 Then
'        Do Until ADORs.EOF
'            If tmpTAG = "" Then
'                tmpTAG = ADORs!택번호
'            Else
'                Do Until Format(CLng(tmpTAG) + 1, "000000000") >= ADORs!택번호
'                    cboMissTag.AddItem Format(CLng(tmpTAG) + 1, "000-00-0000")
'
'                    tmpTAG = Format(CLng(tmpTAG) + 1, "000000000")
'
'                    '100 개가 넘으면 빠져 나옴
'                    If iLoop >= 100 Then
'                        cboMissTag.AddItem "Err"
'
'                        Exit Do
'                    End If
'
'                    iLoop = iLoop + 1
'                Loop
'
'                tmpTAG = Format(CLng(tmpTAG) + 1, "000000000")
'            End If
'
'            ADORs.MoveNext
'        Loop
'        ADORs.Close
'        Set ADORs = Nothing
'
'        If cboMissTag.ListCount = 0 Then
'            txtNum16.Value = 0
'        Else
'            txtNum16.Value = cboMissTag.ListCount - 1
'        End If
'    End If
'
'    pnlTAG(0).Caption = Format(시작택번호, "000-00-0000") & ""
'    pnlTAG(1).Caption = Format(마지막택번호, "000-00-0000") & ""
'
'    Debug.Print "마감작업 : " & Now & " 19) 삼성 카드 할인 내용 추가"
'    '--------------------------------------------------------------------
'    ' 19) 삼성 카드 할인 내용 추가
'    '--------------------------------------------------------------------
'    Dim 삼성카드고객수   As Long
'    Dim 삼성카드할인건수 As Long
'    Dim 삼성카드할인금액 As Long
'
'    삼성카드고객수 = 0
'    삼성카드할인건수 = 0
'    삼성카드할인금액 = 0
'
'    Query = "SELECT    고객코드"
'    Query = Query & ", ISNULL(COUNT(금액),0)"
'    Query = Query & ", ISNULL(SUM(금액),0)"
'    Query = Query & " FROM TB_입출고"
'    Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
'    Query = Query & "   AND 내용  LIKE '%삼%'"
'    Query = Query & "   AND (판매취소 <> 'Y')"
'
'    'Query = Query & "   AND ((판매취소 <> 'Y')"
'    ''Query = Query & "   AND ((판매취소 = '') AND (판매취소일자 IS NULL OR 판매취소일자 = '')"
'    'Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
'    'Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
'
'    Query = Query & " GROUP BY 고객코드"
'    Set ADORs = New ADODB.Recordset
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    Do Until ADORs.EOF
'        삼성카드고객수 = 삼성카드고객수 + 1
'
'        삼성카드할인건수 = 삼성카드할인건수 + ADORs(0)
'        삼성카드할인금액 = 삼성카드할인금액 + ADORs(1)
'
'        ADORs.MoveNext
'    Loop
'    ADORs.Close
'    Set ADORs = Nothing
'
'    txtCost26.Value = 삼성카드할인금액
'    txtNum17.Value = 삼성카드할인건수
'    txtNum18.Value = 삼성카드고객수
'
'    Debug.Print "마감작업 : " & Now & " 종료"
'
'    '미수금액 = 매출액 - 현금결제 - 카드결제 - 사용마일리지 - 쿠폰
'    txtCost04.Value = txtCost01.Value - txtCost02.Value - txtCost03.Value - txtCost19.Value - txtCost21.Value
'
'    Screen.MousePointer = 0
'    pnlProg.Visible = False
'    DoEvents
'
'    Exit Sub
'
'ErrRtn:
'    Screen.MousePointer = 0
'    pnlProg.Visible = False
'End Sub

Private Sub Get_택번호(Query As String, Combo As ComboBox)
    On Error GoTo ErrRtn
    
    Combo.Clear
    
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
        
'        .SortKey(1) = 3
'        .SortKeyOrder(1) = SortKeyOrderDescending
'        .Sort -1, -1, -1, -1, SortByRow

        .ReDraw = True
    End With
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
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

    txtNum02.Value = Recordset_Result(Query) '출고수량


    '----------------------------------------------------------------
    ' 2. 선불결제 2-1) 현금반환/ 현금결제 구하기
    '----------------------------------------------------------------
    Query = "SELECT    ISNULL(SUM(접수금액),0) * -1"
    Query = Query & " FROM TB_매출"
    Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
    Query = Query & "   AND 적요 LIKE '%현금반환%'"
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
    Query = "SELECT ISNULL(SUM(접수금액),0) - ISNULL(SUM(입금합계),0) AS 미수금 "
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
    Query = Query & "   AND (판매취소 <> 'Y')"
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
    Query = Query & "   AND (판매취소 <> 'Y')"
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
    Query = Query & "   AND (판매취소 <> 'Y') "
    Query = Query & "   AND 고객코드 < 900000"
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
    Query = Query & "   AND (판매취소 <> 'Y') "
    
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
        
    pnlData(25).Caption = "+ 로열티 " & 가맹점정보.로열티여부2 & " " & 가맹점정보.로열티비율2
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
    If 가맹점정보.수수료지원여부 = "Y" And IsNumeric(가맹점정보.수수료지원비율) Then txtCard(5).Value = CDbl(txtCard(4).Value) * (CDbl(가맹점정보.수수료지원비율) / 100)
    ADORs.Close:    Set ADORs = Nothing

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
    ' 지사정산금액 = 지사분매출 - (카드수수료지원금+환불금액) + (카드수수료환불금+로열티1+로열티2)
    txtMaster(5).Value = txtMaster(0).Value - (txtMaster(1).Value + txtMaster(4).Value) + (txtMaster(2).Value + txtMaster(3).Value + txtMaster(9).Value) - txtMaster(8).Value

'    txtMaster(6).Value = (txtCost09.Value + txtCost10.Value) - txtMaster(5).Value
    ' 매장 수익금 = 가맹점마진 - ((반품환불금액 + 세탁환불금액) - 세탁/반품환불금액(지산) - 카드수수료환불금 + 카드수수료지원금 - 유통로열티 - 쿠폰금액의 40%
    txtMaster(6).Value = txtCost09.Value - ((txtCost23.Value + txtCost24.Value) - txtMaster(4).Value) - txtMaster(2).Value + txtMaster(1).Value - txtMaster(3).Value - txtMaster(9).Value - (txtCost21.Value - txtMaster(8).Value)
    
    Screen.MousePointer = 0
    pnlProg.Visible = False
    DoEvents

    Exit Sub

ErrRtn:
    
    Screen.MousePointer = 0
    pnlProg.Visible = False
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub



'Private Sub cmdPrint_Click()
'    Dim vText       As Variant
'    Dim sTempKey    As String
'    Dim sNameKey    As String
'
'    Dim 입금액    As Long
'    Dim 요일      As String
'
'    Dim ComboList As String
'
'    On Error GoTo ErrRtn
'
'    Call cmdList_Click
'
'    If sprList.MaxRows = 0 Then Exit Sub
'
'    If Dir(AppPath & "XML", vbDirectory) = "" Then
'        MkDir AppPath & "XML"
'    End If
'
'    If Get_일일마감여부(Format(dtpDay.Value, "YYYY-MM-DD")) = False Then
'        MsgBox "일마감이 완료 되지 않아 출력할 수 없습니다.", vbInformation, "확인"
'        Exit Sub
'    End If
'
'    요일 = Fun_Week(dtpDay.Value)
'    sNameKey = ""
'    Open AppPath & "XML\일일매출현황.XML" For Output As #1
'
'    Print #1, "<?xml version=""1.0"" encoding=""EUC-KR""?>"
'    Print #1, "<root>"
'
'          XML = "    <조건>"
'    XML = XML & "        <접수일자>일자 : " & Format(dtpDay.Value, "YYYY년 MM월 DD일") & " (" & 요일 & ")</접수일자>"
'    XML = XML & "        <가맹점>(" & Func_Replace(가맹점정보.가맹점명) & ") 일일매출현황</가맹점>"
'    XML = XML & "   </조건>"
'    Print #1, XML
'
'    With sprList
'        For i = 1 To .DataRowCnt
'            .Row = i
'
'                             XML = "    <Data>"
'            .Col = 10:  XML = XML & "       <택번호>" & Right(.Text, 7) & "</택번호>"
'
'            ' 이름이 다를 경우만 출력 한다.
'            .GetText 4, i, vText:   sTempKey = CStr(vText)
'            .GetText 5, i, vText:   sTempKey = sTempKey & CStr(vText)
'            .GetText 6, i, vText:   sTempKey = sTempKey & CStr(vText)
'
'            If sNameKey <> sTempKey Then
'                sNameKey = sTempKey
'                .Col = 4:  XML = XML & "        <성명>" & Func_Replace(.Text) & "</성명>"
'                .Col = 5:  XML = XML & "        <휴대전화>" & Func_Replace(.Text) & "</휴대전화>"
'                .Col = 6:  XML = XML & "        <전화번호>" & Func_Replace(.Text) & "</전화번호>"
'            Else
'                .Col = 4:  XML = XML & "        <성명>" & Space(1) & "</성명>"
'                .Col = 5:  XML = XML & "        <휴대전화>" & Space(1) & "</휴대전화>"
'                .Col = 6:  XML = XML & "        <전화번호>" & Space(1) & "</전화번호>"
'            End If
'
'
'            ' 사람이 변경되면
'            .GetText 4, i + 1, vText: sTempKey = CStr(vText)
'            .GetText 5, i + 1, vText: sTempKey = sTempKey & CStr(vText)
'            .GetText 6, i + 1, vText: sTempKey = sTempKey & CStr(vText)
'             XML = XML & "        <선긋기>" & IIf(sNameKey <> sTempKey, "OK", "NO") & "</선긋기>"
'
'            .Col = 11: XML = XML & "        <품명>" & Func_Replace(.Text) & "</품명>"
'            .Col = 12: XML = XML & "        <색상>" & Func_Replace(.Text) & "</색상>"
'            .Col = 13: XML = XML & "        <무늬>" & Func_Replace(.Text) & "</무늬>"
'            .Col = 14: XML = XML & "        <내용>" & Func_Replace(.Text) & "</내용>"
'            .Col = 15: XML = XML & "        <금액>" & .Text & "</금액>"
'            .Col = 17: XML = XML & "        <상표>" & Func_Replace(.Text) & "</상표>"
'            .Col = 16: XML = XML & "        <결제>" & .Text & "</결제>"
'                       XML = XML & "   </Data>" & vbNewLine
'                       Print #1, XML
'        Next i
'    End With
'
'          XML = "    <합계>"
'    XML = XML & "        <접수수량>" & txtNum01.Text & " 점</접수수량>"
'    XML = XML & "        <접수금액>" & txtCost01.Text & " 원</접수금액>"
'    XML = XML & "        <현금결제>" & txtCost07.Text & " 원</현금결제>"
'    XML = XML & "        <카드건수>" & txtNum05.Text & " 건</카드건수>"
'    XML = XML & "        <카드결제>" & txtCost08.Text & " 원</카드결제>"
'
'    XML = XML & "        <쿠폰건수>" & txtNum12.Text & " 건</쿠폰건수>"
'    XML = XML & "        <쿠폰결제>" & txtCost21.Text & " 원</쿠폰결제>"
'
'    입금액 = txtCost07.Value + txtCost08.Value
'
'    XML = XML & "        <입금액>" & Format(입금액, "#,##0") & " 원</입금액>"
'    XML = XML & "        <미수금액>" & txtCost04.Text & " 원</미수금액>"
'    XML = XML & "        <발생마일리지>" & txtCost18.Text & " 원</발생마일리지>"
'    XML = XML & "        <사용마일리지>" & txtCost19.Text & " 원</사용마일리지>"
'    XML = XML & "        <삭제마일리지>" & txtCost20.Text & " 원</삭제마일리지>"
'    XML = XML & "        <가맹점마진>" & txtCost09.Text & " 원</가맹점마진>"
'    XML = XML & "        <지사마진>" & txtCost10.Text & " 원</지사마진>"
'
''    XML = XML & "        <수선수량>" & txtNum06.Text & " 점</수선수량>"
''    XML = XML & "        <재세탁수량>" & txtNum07.Text & " 점</재세탁수량>"
''    XML = XML & "        <운동화수량>" & txtNum08.Text & " 점</운동화수량>"
''    XML = XML & "        <가죽수량>" & txtNum09.Text & " 점</가죽수량>"
'    XML = XML & "        <카페트수량>" & txtNum10.Text & " 점</카페트수량>"
'    XML = XML & "        <반품수량>" & txtNum11.Text & " 점</반품수량>"
'
'    XML = XML & "        <환불금액>" & Format(CLng(txtCost23.Text) + CLng(txtCost24.Text), "#,##0") & "원중 지사분" _
'                                     & Format(CLng(txtCost29.Value) + CLng(txtCost30.Value), "#,##0") & "원</환불금액>"
'
''    XML = XML & "        <수선금액>" & txtCost11.Text & " 점</수선금액>"
''    XML = XML & "        <재세탁금액>" & txtCost12.Text & " 점</재세탁금액>"
''    XML = XML & "        <운동화금액>" & txtCost13.Text & " 점</운동화금액>"
''    XML = XML & "        <가죽금액>" & txtCost14.Text & " 점</가죽금액>"
''    XML = XML & "        <카페트금액>" & txtCost15.Text & " 점</카페트금액>"
''    XML = XML & "        <반품금액>" & txtCost16.Text & " 점</반품금액>"
'
''    XML = XML & "        <판매취소수량>" & txtNum13.Text & " 점</판매취소수량>"
'    XML = XML & "        <반품환불수량>" & txtNum14.Text & " 점</반품환불수량>"
'    XML = XML & "        <세탁환불수량>" & txtNum15.Text & " 점</세탁환불수량>"
'    XML = XML & "        <누락택수량>" & txtNum16.Text & " 점</누락택수량>"
'
'    '---------------------------------------------------------------------------
'    ComboList = ""
'
'    For i = 0 To cboCancel.ListCount - 1
'        ComboList = ComboList & cboCancel.List(i) & " "
'    Next i
'
'    XML = XML & "        <판매취소택>" & ComboList & "</판매취소택>"
'    '---------------------------------------------------------------------------
'
'    ComboList = ""
'
'    For i = 0 To cboReturn.ListCount - 1
'        ComboList = ComboList & cboReturn.List(i) & " "
'    Next i
'
'    XML = XML & "        <반품환불택>" & ComboList & "</반품환불택>"
'    '---------------------------------------------------------------------------
'
'    ComboList = ""
'
'    For i = 0 To cboRepay.ListCount - 1
'        ComboList = ComboList & cboRepay.List(i) & " "
'    Next i
'
'    XML = XML & "        <세탁환불택>" & ComboList & "</세탁환불택>"
'    '---------------------------------------------------------------------------
'
'    ComboList = ""
'
'    For i = 0 To cboMissTag.ListCount - 1
'        ComboList = ComboList & cboMissTag.List(i) & " "
'    Next i
'
'    XML = XML & "        <누락택>" & ComboList & "</누락택>"
'
'    '---------------------------------------------------------------------------
'
'    XML = XML & "   </합계>"
'    Print #1, XML
'
'    Print #1, "</root>"
'    Close #1
'
'    With rpt일일매출현황2
'        .dc.FileURL = AppPath & "XML\일일매출현황2.XML"
''        .Show 1
'
'        .PrintReport False
'    End With
'
'    Unload rpt일일매출현황2
'
'    Exit Sub
'
'ErrRtn:
'
'    Call Error_Msg("", Err.Source, Err.Number, Err.description)
'End Sub

