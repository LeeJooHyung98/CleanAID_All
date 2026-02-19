VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{83FD3014-2044-4BA5-9B6C-F0A2482D9C0C}#1.0#0"; "KICCPOSIEX.OCX"
Begin VB.Form frm사고품 
   Caption         =   "사고품"
   ClientHeight    =   10185
   ClientLeft      =   7395
   ClientTop       =   2760
   ClientWidth     =   15360
   ControlBox      =   0   'False
   LinkTopic       =   "Form16"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10185
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   WindowState     =   2  '최대화
   Begin KiccPosIE.KiccPosIEX KiccPosOCX 
      Height          =   1920
      Left            =   7200
      TabIndex        =   100
      Top             =   4125
      Visible         =   0   'False
      Width           =   960
      BF0C            =   ""
      Bmp             =   ""
      CardNo          =   ""
      CashNo          =   ""
      CommType        =   1
      Connected       =   0   'False
      Emv             =   ""
      EmvLen          =   0
      MasterClaimerText=   ""
      MasterOfferText =   ""
      PIN             =   ""
      SeqNo           =   ""
      Sign            =   ""
      SignLen         =   0
      TID             =   ""
      RfFlag          =   ""
      VAK             =   ""
      VisaClaimerText =   ""
      VisaOfferText   =   ""
      ErrMsg          =   ""
      ResMsg          =   ""
      RcvData         =   ""
      TRNO            =   ""
      Data            =   ""
      CVER            =   ""
      MVER            =   ""
      PVER            =   ""
      TMTransCount    =   0
      TMOnLineCount   =   0
      EBTransCount    =   0
      Alignment       =   2
      AutoSize        =   0   'False
      BevelInner      =   0
      BevelOuter      =   0
      BorderStyle     =   0
      Caption         =   ""
      Color           =   16777215
      Ctl3D           =   -1  'True
      UseDockManager  =   -1  'True
      DockSite        =   0   'False
      DragCursor      =   -12
      Object.DragMode        =   0
      Enabled         =   -1  'True
      FullRepaint     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   0   'False
      ParentColor     =   0   'False
      ParentCtl3D     =   -1  'True
      Object.Visible         =   -1  'True
      DoubleBuffered  =   -1  'True
      Cursor          =   0
      Protocol        =   0
      JcbClaimerText  =   ""
      JcbOfferText    =   ""
      DccTextVer      =   "00"
      CardHash        =   "$"
      SignAD          =   "0000"
      HandleValue     =   66286
      MemberShip      =   ""
      MemberShipHex   =   ""
      TCPSVCPort      =   0
      TCPSVCActive    =   0   'False
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   10185
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   17965
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm사고품.frx":0000
      Begin Threed.SSPanel SSPanel1 
         Height          =   750
         Left            =   15
         TabIndex        =   1
         Top             =   450
         Width           =   15330
         _ExtentX        =   27040
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   0
            Left            =   945
            TabIndex        =   2
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
            Format          =   56295427
            CurrentDate     =   40279
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   1
            Left            =   2640
            TabIndex        =   3
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
            Format          =   56295427
            CurrentDate     =   40279
         End
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   630
            Left            =   4140
            TabIndex        =   6
            Top             =   60
            Width           =   1320
            _Version        =   851970
            _ExtentX        =   2328
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 조회(&F)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm사고품.frx":00B2
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   3
            Left            =   9375
            TabIndex        =   7
            Top             =   60
            Width           =   1350
            _Version        =   851970
            _ExtentX        =   2381
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 엑셀(&E)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm사고품.frx":07AC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   13650
            TabIndex        =   8
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm사고품.frx":0F26
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   4
            Left            =   10740
            TabIndex        =   9
            Top             =   60
            Width           =   1350
            _Version        =   851970
            _ExtentX        =   2381
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 출력(&P)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm사고품.frx":1FB8
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   5475
            TabIndex        =   12
            Top             =   240
            Width           =   1275
            _Version        =   851970
            _ExtentX        =   2249
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 신규(&N)"
            Appearance      =   6
            Picture         =   "frm사고품.frx":26B2
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   6780
            TabIndex        =   13
            Top             =   240
            Width           =   1275
            _Version        =   851970
            _ExtentX        =   2249
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 저장(&S)"
            Appearance      =   6
            Picture         =   "frm사고품.frx":30C4
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   8085
            TabIndex        =   14
            Top             =   240
            Width           =   1275
            _Version        =   851970
            _ExtentX        =   2249
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 삭제(&D)"
            Appearance      =   6
            Picture         =   "frm사고품.frx":3AD6
         End
         Begin XtremeSuiteControls.PushButton cmdPrintMini 
            Height          =   630
            Left            =   12120
            TabIndex        =   99
            Top             =   60
            Visible         =   0   'False
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   "단말기출력"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm사고품.frx":44E8
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
            Index           =   3
            Left            =   2445
            TabIndex        =   5
            Top             =   120
            Width           =   120
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
            Height          =   225
            Index           =   2
            Left            =   60
            TabIndex        =   4
            Top             =   120
            Width           =   840
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   10
         Top             =   15
         Width           =   15330
         _ExtentX        =   27040
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
         Caption         =   "      사고품"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm사고품.frx":4BE2
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm사고품.frx":4E08
            Top             =   -15
            Width           =   765
         End
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Bindings        =   "frm사고품.frx":59D2
         Height          =   8955
         Left            =   15
         TabIndex        =   11
         Top             =   1215
         Width           =   5505
         _Version        =   524288
         _ExtentX        =   9710
         _ExtentY        =   15796
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
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
         MaxCols         =   5
         MaxRows         =   1000000
         OperationMode   =   1
         Protect         =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frm사고품.frx":59E6
         VisibleCols     =   5
         VisibleRows     =   200
         HighlightHeaders=   1
         HighlightStyle  =   1
         ScrollBarStyle  =   2
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   6945
         Left            =   5535
         TabIndex        =   15
         Top             =   1215
         Width           =   9810
         _ExtentX        =   17304
         _ExtentY        =   12250
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.GroupBox GroupBox 
            Height          =   1770
            Index           =   0
            Left            =   45
            TabIndex        =   16
            Top             =   120
            Width           =   9690
            _Version        =   851970
            _ExtentX        =   17092
            _ExtentY        =   3122
            _StockProps     =   79
            Caption         =   "※ 기본 정보"
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
            Appearance      =   6
            BorderStyle     =   1
            Begin VB.TextBox txtData 
               Height          =   315
               Index           =   2
               Left            =   6210
               Locked          =   -1  'True
               TabIndex        =   22
               Top             =   645
               Width           =   3405
            End
            Begin VB.TextBox txtData 
               Height          =   315
               Index           =   3
               Left            =   6210
               Locked          =   -1  'True
               TabIndex        =   21
               Top             =   1005
               Width           =   3405
            End
            Begin VB.TextBox txtData 
               Height          =   315
               Index           =   4
               Left            =   1260
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   20
               Top             =   1365
               Width           =   8355
            End
            Begin VB.TextBox txtData 
               Height          =   315
               Index           =   1
               Left            =   1260
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   19
               Top             =   1005
               Width           =   3405
            End
            Begin VB.TextBox txtData 
               Height          =   315
               Index           =   5
               Left            =   6210
               Locked          =   -1  'True
               TabIndex        =   18
               Top             =   285
               Width           =   3405
            End
            Begin VB.TextBox txtData 
               Height          =   315
               Index           =   0
               Left            =   1260
               Locked          =   -1  'True
               TabIndex        =   17
               Top             =   645
               Width           =   960
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   1
               Left            =   2790
               TabIndex        =   23
               Top             =   645
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "일련번호"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   2
               Left            =   5055
               TabIndex        =   24
               Top             =   645
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "전화번호"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   17
               Left            =   5055
               TabIndex        =   25
               Top             =   285
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "담당자명"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   34
               Left            =   5055
               TabIndex        =   26
               Top             =   1005
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "휴대전화"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   11
               Left            =   105
               TabIndex        =   27
               Top             =   1005
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "성      명"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   13
               Left            =   105
               TabIndex        =   28
               Top             =   1365
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "주      소"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin CSTextLibCtl.sidbEdit txtCode 
               Height          =   330
               Left            =   3945
               TabIndex        =   29
               Top             =   645
               Width           =   720
               _Version        =   262145
               _ExtentX        =   1270
               _ExtentY        =   582
               _StockProps     =   125
               Text            =   " 999,999,999,999"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
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
               RawData         =   ""
               Text            =   " 0"
               StartText.x     =   3
               StartText.y     =   4
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
               Justification   =   1
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               Undo            =   0
               Data            =   0
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   0
               Left            =   105
               TabIndex        =   30
               Top             =   285
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               BackColor       =   12648384
               Caption         =   "사고접수일"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtpDay 
               Height          =   315
               Index           =   2
               Left            =   1260
               TabIndex        =   31
               Top             =   285
               Width           =   3405
               _ExtentX        =   6006
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CalendarBackColor=   16777215
               Format          =   56295424
               CurrentDate     =   40279
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   19
               Left            =   105
               TabIndex        =   32
               Top             =   645
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "고객코드"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox 
            Height          =   2490
            Index           =   1
            Left            =   45
            TabIndex        =   33
            Top             =   1950
            Width           =   9690
            _Version        =   851970
            _ExtentX        =   17092
            _ExtentY        =   4392
            _StockProps     =   79
            Caption         =   "※ 품목 정보"
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
            Appearance      =   6
            BorderStyle     =   1
            Begin VB.TextBox txtData 
               Height          =   315
               Index           =   11
               Left            =   1260
               MaxLength       =   10
               TabIndex        =   40
               Top             =   2085
               Width           =   3405
            End
            Begin VB.TextBox txtData 
               Height          =   315
               Index           =   10
               Left            =   6210
               MaxLength       =   20
               TabIndex        =   39
               Top             =   1725
               Width           =   3405
            End
            Begin VB.TextBox txtData 
               Height          =   315
               Index           =   9
               Left            =   1260
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   38
               Top             =   1365
               Width           =   3405
            End
            Begin VB.TextBox txtData 
               Height          =   315
               Index           =   8
               Left            =   6210
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   37
               Top             =   1005
               Width           =   3405
            End
            Begin VB.TextBox txtData 
               Height          =   315
               Index           =   7
               Left            =   1260
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   36
               Top             =   1005
               Width           =   3405
            End
            Begin VB.TextBox txtData 
               BackColor       =   &H0080FFFF&
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   6
               Left            =   6210
               MaxLength       =   20
               TabIndex        =   35
               Top             =   285
               Width           =   2715
            End
            Begin VB.CommandButton cmdTag 
               Caption         =   "..."
               Height          =   330
               Left            =   10125
               TabIndex        =   34
               Top             =   270
               Width           =   540
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   5
               Left            =   105
               TabIndex        =   41
               Top             =   285
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "접수일자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   6
               Left            =   105
               TabIndex        =   42
               Top             =   1005
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "품     목"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   7
               Left            =   5055
               TabIndex        =   43
               Top             =   1005
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "상      표"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtpDay 
               Height          =   315
               Index           =   3
               Left            =   1260
               TabIndex        =   44
               Top             =   285
               Width           =   3405
               _ExtentX        =   6006
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   56295424
               CurrentDate     =   36686
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   3
               Left            =   5055
               TabIndex        =   45
               Top             =   285
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               BackColor       =   12648384
               Caption         =   "택 번 호"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   4
               Left            =   105
               TabIndex        =   46
               Top             =   645
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "출고일자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtpDay 
               Height          =   315
               Index           =   4
               Left            =   1260
               TabIndex        =   47
               Top             =   645
               Width           =   3405
               _ExtentX        =   6006
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   56295424
               CurrentDate     =   36686
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   8
               Left            =   5055
               TabIndex        =   48
               Top             =   645
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "인도일자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtpDay 
               Height          =   315
               Index           =   5
               Left            =   6210
               TabIndex        =   49
               Top             =   645
               Width           =   3405
               _ExtentX        =   6006
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   56295424
               CurrentDate     =   36686
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   9
               Left            =   105
               TabIndex        =   50
               Top             =   1365
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "색     상"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   10
               Left            =   105
               TabIndex        =   51
               Top             =   1725
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               BackColor       =   12648384
               Caption         =   "구입일자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtpDay 
               Height          =   315
               Index           =   6
               Left            =   1260
               TabIndex        =   52
               Top             =   1725
               Width           =   3405
               _ExtentX        =   6006
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   56295424
               CurrentDate     =   36686
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   25
               Left            =   5055
               TabIndex        =   53
               Top             =   1725
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               BackColor       =   12648384
               Caption         =   "구 입 처"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   26
               Left            =   105
               TabIndex        =   54
               Top             =   2085
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               BackColor       =   12648384
               Caption         =   "구입형태"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   27
               Left            =   5055
               TabIndex        =   55
               Top             =   2085
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               BackColor       =   12648384
               Caption         =   "구입가격"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin CSTextLibCtl.sidbEdit txtNum 
               Height          =   330
               Index           =   0
               Left            =   6210
               TabIndex        =   56
               Top             =   2070
               Width           =   3405
               _Version        =   262145
               _ExtentX        =   6006
               _ExtentY        =   582
               _StockProps     =   125
               Text            =   " 999,999,999,999"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
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
               RawData         =   ""
               Text            =   " 0"
               StartText.x     =   3
               StartText.y     =   4
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
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               Undo            =   0
               Data            =   0
            End
            Begin XtremeSuiteControls.PushButton btnTAG 
               Height          =   315
               Left            =   8970
               TabIndex        =   57
               Top             =   285
               Width           =   630
               _Version        =   851970
               _ExtentX        =   1111
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "검색"
               Appearance      =   6
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox 
            Height          =   1410
            Index           =   4
            Left            =   45
            TabIndex        =   58
            Top             =   4500
            Width           =   9690
            _Version        =   851970
            _ExtentX        =   17092
            _ExtentY        =   2487
            _StockProps     =   79
            Caption         =   "※ 보상 산정 기준"
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
            Appearance      =   6
            BorderStyle     =   1
            Begin VB.ComboBox cboInput 
               Height          =   300
               Index           =   0
               ItemData        =   "frm사고품.frx":6061
               Left            =   1260
               List            =   "frm사고품.frx":6063
               Locked          =   -1  'True
               TabIndex        =   61
               Text            =   "cboInput"
               Top             =   285
               Width           =   1875
            End
            Begin VB.ComboBox cboInput 
               Height          =   300
               Index           =   1
               Left            =   4500
               Locked          =   -1  'True
               TabIndex        =   60
               Text            =   "cboInput"
               Top             =   285
               Width           =   1875
            End
            Begin VB.ComboBox cboInput 
               Height          =   300
               Index           =   2
               Left            =   7740
               Locked          =   -1  'True
               TabIndex        =   59
               Text            =   "cboInput"
               Top             =   285
               Width           =   1875
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   20
               Left            =   105
               TabIndex        =   62
               Top             =   285
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "품     목"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   23
               Left            =   3345
               TabIndex        =   63
               Top             =   285
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "용      도"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   24
               Left            =   6585
               TabIndex        =   64
               Top             =   285
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "소     재"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   28
               Left            =   105
               TabIndex        =   65
               Top             =   645
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "내용연수"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   29
               Left            =   3345
               TabIndex        =   66
               Top             =   645
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "경과일수"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   30
               Left            =   6585
               TabIndex        =   67
               Top             =   645
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "환산일수"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   31
               Left            =   105
               TabIndex        =   68
               Top             =   1005
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "배상비율"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   32
               Left            =   3345
               TabIndex        =   69
               Top             =   1005
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "배상금액"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin CSTextLibCtl.sidbEdit txtNum 
               Height          =   315
               Index           =   1
               Left            =   1260
               TabIndex        =   70
               Top             =   630
               Width           =   1875
               _Version        =   262145
               _ExtentX        =   3307
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 999,999,999,999"
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
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Insert          =   0   'False
               Modified        =   -1  'True
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   " 0"
               StartText.x     =   3
               StartText.y     =   4
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
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txtNum 
               Height          =   315
               Index           =   2
               Left            =   4500
               TabIndex        =   71
               Top             =   630
               Width           =   1860
               _Version        =   262145
               _ExtentX        =   3281
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 999,999,999,999"
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
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Insert          =   0   'False
               Modified        =   -1  'True
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   " 0"
               StartText.x     =   3
               StartText.y     =   4
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
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txtNum 
               Height          =   315
               Index           =   3
               Left            =   7740
               TabIndex        =   72
               Top             =   630
               Width           =   1860
               _Version        =   262145
               _ExtentX        =   3281
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 999,999,999,999"
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
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Insert          =   0   'False
               Modified        =   -1  'True
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   " 0"
               StartText.x     =   3
               StartText.y     =   4
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
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txtNum 
               Height          =   315
               Index           =   4
               Left            =   1260
               TabIndex        =   73
               Top             =   990
               Width           =   1860
               _Version        =   262145
               _ExtentX        =   3281
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 999,999,999,999"
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
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Insert          =   0   'False
               Modified        =   -1  'True
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   " 0"
               StartText.x     =   3
               StartText.y     =   4
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
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txtNum 
               Height          =   315
               Index           =   5
               Left            =   4500
               TabIndex        =   74
               Top             =   990
               Width           =   1875
               _Version        =   262145
               _ExtentX        =   3307
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 999,999,999,999"
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
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Insert          =   0   'False
               Modified        =   -1  'True
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   " 0"
               StartText.x     =   3
               StartText.y     =   4
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
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               Undo            =   0
               Data            =   0
            End
         End
      End
      Begin XtremeSuiteControls.TabControl TabControl 
         Height          =   1995
         Left            =   5535
         TabIndex        =   75
         Top             =   8175
         Width           =   9810
         _Version        =   851970
         _ExtentX        =   17304
         _ExtentY        =   3519
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   3
         Color           =   16
         PaintManager.BoldSelected=   -1  'True
         PaintManager.HotTracking=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         PaintManager.ButtonMargin=   "2,3,2,3"
         ItemCount       =   4
         Item(0).Caption =   "사고처리 정보"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "TabControlPage(0)"
         Item(1).Caption =   "고객의견"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "TabControlPage(1)"
         Item(2).Caption =   "가맹점 및 지사의견"
         Item(2).ControlCount=   1
         Item(2).Control(0)=   "TabControlPage(2)"
         Item(3).Caption =   "본사의견"
         Item(3).ControlCount=   1
         Item(3).Control(0)=   "TabControlPage1"
         Begin XtremeSuiteControls.TabControlPage TabControlPage1 
            Height          =   1515
            Left            =   -69970
            TabIndex        =   76
            Top             =   450
            Visible         =   0   'False
            Width           =   9750
            _Version        =   851970
            _ExtentX        =   17198
            _ExtentY        =   2672
            _StockProps     =   1
            Page            =   3
            Begin VB.TextBox txtData 
               Height          =   1005
               Index           =   15
               Left            =   60
               MultiLine       =   -1  'True
               TabIndex        =   77
               Top             =   420
               Width           =   9645
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   33
               Left            =   60
               TabIndex        =   78
               Top             =   75
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "본사 승인일자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel pnlDate 
               Height          =   315
               Index           =   1
               Left            =   1440
               TabIndex        =   79
               Top             =   75
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               _Version        =   262144
               BackColor       =   12648447
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage 
            Height          =   1515
            Index           =   2
            Left            =   -69970
            TabIndex        =   80
            Top             =   450
            Visible         =   0   'False
            Width           =   9750
            _Version        =   851970
            _ExtentX        =   17198
            _ExtentY        =   2672
            _StockProps     =   1
            Page            =   2
            Begin VB.TextBox txtData 
               Height          =   1005
               Index           =   14
               Left            =   60
               MultiLine       =   -1  'True
               TabIndex        =   81
               Top             =   420
               Width           =   9645
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   22
               Left            =   60
               TabIndex        =   82
               Top             =   75
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "지사 승인일자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel pnlDate 
               Height          =   315
               Index           =   0
               Left            =   1440
               TabIndex        =   83
               Top             =   75
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               _Version        =   262144
               BackColor       =   12648447
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage 
            Height          =   1515
            Index           =   1
            Left            =   -69970
            TabIndex        =   84
            Top             =   450
            Visible         =   0   'False
            Width           =   9750
            _Version        =   851970
            _ExtentX        =   17198
            _ExtentY        =   2672
            _StockProps     =   1
            Page            =   1
            Begin VB.TextBox txtData 
               Height          =   1380
               Index           =   13
               Left            =   45
               MultiLine       =   -1  'True
               TabIndex        =   85
               Top             =   60
               Width           =   9645
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage 
            Height          =   1515
            Index           =   0
            Left            =   30
            TabIndex        =   86
            Top             =   450
            Width           =   9750
            _Version        =   851970
            _ExtentX        =   17198
            _ExtentY        =   2672
            _StockProps     =   1
            Page            =   0
            Begin VB.TextBox txtData 
               Height          =   315
               Index           =   12
               Left            =   1245
               MaxLength       =   50
               TabIndex        =   90
               Top             =   1125
               Width           =   7290
            End
            Begin VB.ComboBox cboInput 
               Height          =   300
               Index           =   4
               Left            =   1245
               Locked          =   -1  'True
               Style           =   2  '드롭다운 목록
               TabIndex        =   89
               Top             =   435
               Width           =   2580
            End
            Begin VB.ComboBox cboInput 
               Height          =   300
               Index           =   3
               Left            =   1245
               Style           =   2  '드롭다운 목록
               TabIndex        =   88
               Top             =   90
               Width           =   2580
            End
            Begin VB.ComboBox cboInput 
               Height          =   300
               Index           =   5
               Left            =   1245
               Locked          =   -1  'True
               Style           =   2  '드롭다운 목록
               TabIndex        =   87
               Top             =   780
               Width           =   2580
            End
            Begin Threed.SSPanel panCaption 
               Height          =   300
               Index           =   12
               Left            =   90
               TabIndex        =   91
               Top             =   90
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   529
               _Version        =   262144
               Caption         =   "크레임구분"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   300
               Index           =   14
               Left            =   90
               TabIndex        =   92
               Top             =   435
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   529
               _Version        =   262144
               Caption         =   "보상구분"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   15
               Left            =   4800
               TabIndex        =   93
               Top             =   435
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "보상금액"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   16
               Left            =   90
               TabIndex        =   94
               Top             =   1125
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "비      고"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   21
               Left            =   4800
               TabIndex        =   95
               Top             =   795
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "처리일자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtpDay 
               Height          =   315
               Index           =   7
               Left            =   5970
               TabIndex        =   96
               Top             =   795
               Width           =   2580
               _ExtentX        =   4551
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   56295424
               CurrentDate     =   36684
            End
            Begin Threed.SSPanel panCaption 
               Height          =   300
               Index           =   18
               Left            =   90
               TabIndex        =   97
               Top             =   780
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   529
               _Version        =   262144
               Caption         =   "처리구분"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin CSTextLibCtl.sidbEdit txtNum 
               Height          =   315
               Index           =   6
               Left            =   5955
               TabIndex        =   98
               Top             =   435
               Width           =   2580
               _Version        =   262145
               _ExtentX        =   4551
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 999,999,999,999"
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
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Insert          =   0   'False
               Modified        =   -1  'True
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   " 999,999,999,999"
               StartText.x     =   3
               StartText.y     =   4
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
               BorderStyle     =   0
               FmtControl      =   1
               NumDecDigits    =   0
               Undo            =   0
               Data            =   0
            End
         End
      End
   End
End
Attribute VB_Name = "frm사고품"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bSave As Boolean

Private Sub btnTAG_Click()
    On Error GoTo ErrRtn
    
    If txtData(6).Text = "" Then Exit Sub
    
    '-----------------------------------------------------------------------------------
    ' TB_고객
    '-----------------------------------------------------------------------------------
    Query = "SELECT    A.택번호"
    Query = Query & ", B.성명"
    Query = Query & ", B.전화번호"
    Query = Query & ", A.접수일자"
    Query = Query & ", A.의류명"
    Query = Query & " FROM TB_입출고 AS A LEFT OUTER JOIN TB_고객정보 AS B ON A.고객코드 = B.고객코드"
    Query = Query & " WHERE A.택번호 LIKE '%" & Replace(txtData(6).Text, "-", "") & "%'"
    Query = Query & "   AND A.판매취소 <> 'Y'"
    
    Query = Query & " ORDER BY  A.접수일자 DESC, A.택번호 DESC"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenStatic, adLockReadOnly
    
    With fpList1
        Set .DataSource = ADORs
        
        .Top = 3810
        .Left = 7380
        
        .Width = 7845
        .Height = 2175
        
        .Visible = True
        DoEvents
                
        .SetFocus
    End With
    
    ADORs.Close
    Set ADORs = Nothing
    
    '바로출력 할 수 없도록 한다.
    bSave = False
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyChk (KeyCode)
    
    If KeyCode = 13 Then
        SendKeys "{Tab}"
    End If
End Sub

Private Sub Form_Load()
    'Call Error_Msg("", "test", "1", "1")
    With sprGrid
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
    'Call Error_Msg("", "test", "1", "2")
    With fpList1
        .ColumnHeaderHeight = 300
        .RowHeight = 300
    
        .ListApplyTo = ListApplyToColHeaders
        .BackColor = RGB(192, 192, 192)
        .LineStyle = LineStyleRaised
    End With
    
    'Call Error_Msg("", "test", "1", "3")
    For i = 0 To 2
        dtpDay(i).Value = Format(Date, "YYYY-MM-DD")
    Next i
    'Call Error_Msg("", "test", "1", "4")
    For i = 3 To 7
        dtpDay(i).Value = Format(Date, "YYYY-MM-DD")
        dtpDay(i).Value = ""
    Next i
    'Call Error_Msg("", "test", "1", "5")
    For i = 0 To 5
        cboInput(i).Clear
        
    Next i
    'Call Error_Msg("", "test", "1", "6")
    ' 콤보 박스 설정
    Call ComboAdd
    
    TabControl.SelectedItem = 0
    
End Sub

Private Sub Data_Display()
    Query = "SELECT    일련번호"
    Query = Query & ", 사고접수일자"
    Query = Query & ", 성명"
    Query = Query & ", 의류명"
    Query = Query & ", ISNULL(본사전송여부,'N') AS 본사전송여부"
    Query = Query & " FROM TB_사고품내역"
    Query = Query & " WHERE 가맹점코드 = '" & 가맹점정보.가맹점코드 & "' "
    Query = Query & "   AND (사고접수일자 >= '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "'"
    Query = Query & "   AND  사고접수일자 <= '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "')"
    Query = Query & " ORDER BY 사고접수일자, 일련번호 ASC"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprGrid
        .MaxRows = 0
    
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = ADORs!일련번호 & ""                  '1
            .Col = 2: .Text = Format(ADORs!사고접수일자, "YYYY-MM-DD") '2
            .Col = 3: .Text = ADORs!성명 & ""                      '3
            .Col = 4: .Text = ADORs!의류명 & ""                    '4
            
            If ADORs!본사전송여부 = "Y" Then
                .Col = 5: .Text = "1"                              '5
            Else
                .Col = 5: .Text = "0"                              '5
            End If
            
            ADORs.MoveNext
        Loop
    End With
    ADORs.Close
    Set ADORs = Nothing
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    Dim bSMS_YN As Boolean
    
    On Error GoTo ErrRtn
    
    bSMS_YN = False
    Select Case Index
        Case 0
            Call Text_Clear
            bSave = False
        
        Case 1:
            If Trim(txtData(6).Text) = "" Then
                MsgBox "사고품 택번호를 입력해 주세요.", vbInformation, "확인"
                
                txtData(6).SetFocus
                
                Exit Sub
            End If
            
            If Trim(txtData(7).Text) = "" Then
                MsgBox "사고품 택번호를 입력해 주세요.", vbInformation, "확인"
                
                txtData(6).SetFocus
                
                Exit Sub
            End If
            
            If Trim(txtData(0).Text) = "" Then
                MsgBox "사고품 택번호를 입력해 주세요.", vbInformation, "확인"
                
                txtData(6).SetFocus
                
                Exit Sub
            End If
            
            If Trim(cboInput(3).Text) = "" Then
                MsgBox "사고처리 정보 항목중 크레임 구분을 선택하여 주십시요..", vbInformation, "확인"
                TabControl.SelectedItem = 0
                cboInput(3).SetFocus
                
                Exit Sub
            End If
            
            If Format(dtpDay(6).Value, "YYYY-MM-DD") = "" Then
                MsgBox "사고품 구입일자를 입력해 주세요.", vbInformation, "확인"
                
                dtpDay(6).SetFocus
                Exit Sub
            End If
            
            If txtCode.Value = 0 Then
                Query = "SELECT ISNULL(MAX(일련번호),0) + 1 FROM TB_사고품내역 "
                Query = Query & " WHERE 가맹점코드 = '" & 가맹점정보.가맹점코드 & "' "
                Set ADORs = New ADODB.RecordSet
                ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
                                    
                txtCode.Value = ADORs(0) & ""
                
                ADORs.Close
                Set ADORs = Nothing
            End If
            
            '-------------------------------------------------------------
            ' TB_사고품
            '-------------------------------------------------------------
            Query = "SELECT * FROM TB_사고품내역"
            Query = Query & " WHERE 일련번호 = " & txtCode.Value
            Query = Query & "   AND 가맹점코드 = '" & 가맹점정보.가맹점코드 & "' "
            
            Set ADORs = New ADODB.RecordSet
            ADORs.Open Query, ADOCon, adOpenDynamic, adLockOptimistic
        
            If ADORs.EOF Then
                ADORs.AddNew
                bSMS_YN = True
            
                ADORs!처리일자 = ""                                    '34
                ADORs!지사승인일시 = ""                                '39
                ADORs!본사승인일시 = ""                                '40
            End If
            
            ADORs!지사코드 = 가맹점정보.지사코드 & ""                  ' 1
            ADORs!가맹점코드 = 가맹점정보.가맹점코드 & ""              ' 2
            ADORs!일련번호 = txtCode.Value & ""                        ' 3
            ADORs!사고접수일자 = Format(dtpDay(2).Value, "YYYY-MM-DD") ' 4
            
            ADORs!고객코드 = txtData(0).Text & ""                      ' 5
            ADORs!성명 = txtData(1).Text & ""                          ' 6
            ADORs!전화번호 = txtData(2).Text & ""                      ' 7
            ADORs!휴대전화 = txtData(3).Text & ""                      ' 8
            ADORs!주소 = txtData(4).Text & ""                          ' 9
            
            ADORs!담당자명 = txtData(5).Text & ""                      '10
            
            If dtpDay(3).Value = "" Then
                ADORs!접수일자 = ""                                    '11
            Else
                ADORs!접수일자 = Format(dtpDay(3).Value, "YYYY-MM-DD") '11
            End If
            
            ADORs!택번호 = txtData(6).Text & ""                        '12
            
            If dtpDay(4).Value = "" Then
                ADORs!출고일자 = ""                                    '13
            Else
                ADORs!출고일자 = Format(dtpDay(4).Value, "YYYY-MM-DD") '13
            End If
            
            If dtpDay(5).Value = "" Then
                ADORs!인도일자 = ""                                    '14
            Else
                ADORs!인도일자 = Format(dtpDay(5).Value, "YYYY-MM-DD") '14
            End If
            
            ADORs!의류명 = txtData(7).Text & ""                        '15
            ADORs!상표 = txtData(8).Text & ""                          '16
            ADORs!색상 = txtData(9).Text & ""                          '17
            
            If Format(dtpDay(6).Value, "YYYY-MM-DD") = "" Then
                ADORs!구입일자 = ""                                    '18
            Else
                ADORs!구입일자 = Format(dtpDay(6).Value, "YYYY-MM-DD") '18
            End If
            
            ADORs!구입처 = txtData(10).Text & ""                       '19
            ADORs!구입형태 = txtData(11).Text & ""                     '20
            ADORs!구입가격 = txtNum(0).Value & ""                      '21
            
            
            ADORs!품목 = cboInput(0).Text & ""                         '22
            ADORs!용도 = cboInput(1).Text & ""                         '23
            ADORs!소재 = cboInput(2).Text & ""                         '24
            
            ADORs!내용연수 = txtNum(1).Value                           '25
            ADORs!경과일수 = txtNum(2).Value                           '26
            ADORs!환산일수 = txtNum(3).Value                           '27
            ADORs!배상비율 = txtNum(4).Value                           '28
            ADORs!배상금액 = txtNum(5).Value                           '29
                         
            ADORs!크레임구분 = cboInput(3).Text & ""                   '30
            ADORs!보상구분 = cboInput(4).Text & ""                     '31
            ADORs!처리구분 = cboInput(5).Text & ""                     '32
           
            ADORs!보상금액 = txtNum(6).Value                           '33
           
            'If IsNull(dtpDay(7).Value) Then
            '    ADORs!처리일자 = ""                                    '34
            'Else
            '    ADORs!처리일자 = Format(dtpDay(7).Value, "YYYY-MM-DD") '34
            'End If
            
            ADORs!비고 = txtData(12).Text & ""                         '35
            
            ADORs!가맹점의견 = txtData(13).Text & ""                   '36
            ADORs!지사의견 = txtData(14).Text & ""                     '37
            ADORs!본사의견 = txtData(15).Text & ""                     '38
            
            'If IsNull(dtpDay(8).Value) Then
            '    ADORs!지사승인일시 = ""                                    '39
            'Else
            '    ADORs!지사승인일시 = Format(dtpDay(8).Value, "YYYY-MM-DD") '39
            'End If
            
            'If IsNull(dtpDay(9).Value) Then
            '    ADORs!본사승인일자 = ""                                    '40
            'Else
            '    ADORs!본사승인일자 = Format(dtpDay(9).Value, "YYYY-MM-DD") '40
            'End If
            
            ADORs!본사전송여부 = "N"                                       '41
            
            ADORs.Update
            
            ADORs.Close
            Set ADORs = Nothing
            
            '최초 등록일 경우 한번만 보낸다.
            If bSMS_YN = True Then
                ' 해당 내용으로 사고품 정보를 SMS로 발송한다.
                Call Send_SMS_사고품(txtCode.Value)
            End If
            
            Call Text_Clear
            Call Data_Display
            bSave = True
            
            
        Case 2:
            Rtn = MsgBox("삭제하시겠습니까?", vbYesNo + vbDefaultButton2 + vbQuestion, "삭제")
            
            If Rtn = vbYes Then
                Query = "DELETE FROM TB_사고품내역"
                Query = Query & " WHERE 일련번호 = " & txtCode.Value
                Query = Query & "   AND 가맹점코드 = '" & 가맹점정보.가맹점코드 & "' "
                ADOCon.Execute Query
                
                Call Text_Clear
                Call Data_Display
            End If
            
        Case 3: Call Export_Excel(frmMain.cdgExcel, sprGrid)
        
        Case 4:
            If bSave = False Then
                MsgBox "저장 후 출력 하여 주십시요.", vbInformation, "확인"
                Exit Sub
            End If
        
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
    
    Dim XML         As String
    Dim FileNumber
            
    FileNumber = FreeFile
    
    Open App.Path & "\XML\사고접수.XML" For Output As #FileNumber
    
    Print #FileNumber, "<?xml version=""1.0"" encoding=""EUC-KR""?>"
    Print #FileNumber, "<root>"
    
          XML = ""
        
    Query = "SELECT * FROM TB_기본정보"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF Then
        XML = XML & "    <가맹점명></가맹점명>"
        XML = XML & "    <가맹점주소></가맹점주소>"
        XML = XML & "    <가맹점전화번호></가맹점전화번호>"
    Else
        XML = XML & "    <가맹점명>" & Func_Replace(ADORs!가맹점명) & "</가맹점명>"
        XML = XML & "    <가맹점주소>" & Func_Replace(ADORs!사업장주소) & "</가맹점주소>"
        XML = XML & "    <가맹점전화번호>" & Func_Replace(ADORs!매장전화번호) & "</가맹점전화번호>"
    End If
    ADORs.Close
    Set ADORs = Nothing
    
    XML = XML & "    <소비자명>" & Func_Replace(txtData(1).Text) & "</소비자명>"
    XML = XML & "    <소비자주소>" & Func_Replace(txtData(4).Text) & "</소비자주소>"
    XML = XML & "    <소비자전화번호>" & Func_Replace(txtData(2).Text) & "</소비자전화번호>"
    XML = XML & "    <소비자휴대전화>" & Func_Replace(txtData(3).Text) & "</소비자휴대전화>"
    
    XML = XML & "    <품목>" & Func_Replace(txtData(7).Text) & "</품목>"
    XML = XML & "    <상표>" & Func_Replace(txtData(8).Text) & "</상표>"
    XML = XML & "    <구입일자>" & Format(dtpDay(6).Value, "YYYY-MM-DD") & "</구입일자>"
    XML = XML & "    <색상>" & Func_Replace(txtData(9).Text) & "</색상>"
    XML = XML & "    <구입처>" & Func_Replace(txtData(10).Text) & "</구입처>"
    XML = XML & "    <구입형태>" & Func_Replace(txtData(11).Text) & "</구입형태>"
    XML = XML & "    <구입가격>" & txtNum(0).Text & "</구입가격>"
    XML = XML & "    <사고접수일>" & Format(dtpDay(2).Value, "YYYY-MM-DD") & "</사고접수일>"
    
    XML = XML & "    <최초택번호>" & Func_Replace(txtData(6).Text) & "</최초택번호>"
    XML = XML & "    <최초입고일>" & Format(dtpDay(3).Value, "YYYY-MM-DD") & "</최초입고일>"
    
    XML = XML & "    <최종택번호></최종택번호>"
    XML = XML & "    <최종입고일></최종입고일>"
    
    XML = XML & "    <사고종류></사고종류>"
    XML = XML & "    <사고내용></사고내용>"
    XML = XML & "    <요구사항>" & Func_Replace(txtData(13).Text) & "</요구사항>"
    
    XML = XML & "    <제조회사></제조회사>"
    XML = XML & "    <재고현황></재고현황>"
    
    If txtNum(4).Value = 0 Then
        XML = XML & "    <보상비율></보상비율>"
    Else
        XML = XML & "    <보상비율>" & txtNum(4).Text & "</보상비율>"
    End If
    
    If txtNum(5).Value = 0 Then
        XML = XML & "    <보상산정금액></보상산정금액>"
    Else
        XML = XML & "    <보상산정금액>" & txtNum(5).Text & "</보상산정금액>"
    End If
    
    Print #FileNumber, XML
    
    
    Print #FileNumber, "</root>"
    Close #FileNumber
        
    If Print_PreView = True Then
        With rpt사고접수
            .dc.FileURL = AppPath & "XML\사고접수.XML"
            .Show 1
        End With
    Else
        With rpt사고접수
            .dc.FileURL = AppPath & "XML\사고접수.XML"
            .PrintReport False
        End With
    
        Unload rpt사고접수
    End If
    
    Exit Sub

ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
    Screen.MousePointer = 0
End Sub

Private Sub Text_Clear()
    For i = 0 To 15
        txtData(i).Text = ""
        txtData(i).tag = ""
    Next i
    
    For i = 0 To 2
        cboInput(i).Text = ""
    Next i
    
    pnlDate(0).Caption = ""
    pnlDate(1).Caption = ""
    
    '컨트롤 초기화
    Dim ctrl As Control
    Dim txt  As sidbEdit
    
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is sidbEdit Then
            ctrl.Value = 0
        End If
    Next ctrl
    
    For i = 0 To 2
        dtpDay(i).Value = Format(Date, "YYYY-MM-DD")
    Next i
    
    For i = 3 To 7
        dtpDay(i).Value = Format(Date, "YYYY-MM-DD")
        dtpDay(i).Value = ""
    Next i
    
    txtData(6).SetFocus
End Sub

Private Sub cmdList_Click()
    Call Data_Display
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    cmdBtn(5).Left = Me.Width - cmdBtn(5).Width - 200
End Sub

Private Sub sprGrid_Click(ByVal Col As Long, ByVal Row As Long)
    On Error GoTo ErrRtn
    
    If Row <= 0 Then Exit Sub
    
    Dim 일련번호 As Long
    
    sprGrid.Row = Row
    sprGrid.Col = 1: 일련번호 = sprGrid.Text & ""
    
    '-------------------------------------------------------------
    ' TB_사고품
    '-------------------------------------------------------------
    Query = "SELECT * FROM TB_사고품내역"
    Query = Query & " WHERE 일련번호 = " & 일련번호
    Query = Query & "   AND 가맹점코드 = '" & 가맹점정보.가맹점코드 & "' "
    
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    If ADORs.EOF Then ADORs.AddNew
    
    txtCode.Value = ADORs!일련번호 & ""                        ' 3
    dtpDay(2).Value = Format(ADORs!사고접수일자, "YYYY-MM-DD") ' 4
    
    txtData(0).Text = ADORs!고객코드 & ""                      ' 5
    txtData(1).Text = ADORs!성명 & ""                          ' 6
    txtData(2).Text = ADORs!전화번호 & ""                      ' 7
    txtData(3).Text = ADORs!휴대전화 & ""                      ' 8
    txtData(4).Text = ADORs!주소 & ""                          ' 9
    txtData(5).Text = ADORs!담당자명 & ""                      '10
    dtpDay(3).Value = Format(ADORs!접수일자, "YYYY-MM-DD")     '11
    txtData(6).Text = ADORs!택번호 & ""                        '12
    dtpDay(4).Value = Format(ADORs!출고일자, "YYYY-MM-DD")     '13
    dtpDay(5).Value = Format(ADORs!인도일자, "YYYY-MM-DD")     '14
    
    txtData(7).Text = ADORs!의류명 & ""                        '15
    txtData(8).Text = ADORs!상표 & ""                          '16
    txtData(9).Text = ADORs!색상 & ""                          '17
    
    dtpDay(6).Value = Format(ADORs!구입일자, "YYYY-MM-DD")     '18
        
    txtData(10).Text = ADORs!구입처 & ""                       '19
    txtData(11).Text = ADORs!구입형태 & ""                     '20
    txtNum(0).Value = ADORs!구입가격 & ""                      '21
    
    Call cboSelectText(cboInput(0), ADORs!품목 & "", 0)
    Call cboSelectText(cboInput(1), ADORs!용도 & "", 0)
    Call cboSelectText(cboInput(2), ADORs!소재 & "", 0)
    
'    cboInput(0).Text = ADORs!품목 & ""                         '22
'    cboInput(1).Text = ADORs!용도 & ""                         '23
'    cboInput(2).Text = ADORs!소재 & ""                         '24
    
    txtNum(1).Value = ADORs!내용연수 & ""                      '25
    txtNum(2).Value = ADORs!경과일수 & ""                      '26
    txtNum(3).Value = ADORs!환산일수 & ""                      '27
    txtNum(4).Value = ADORs!배상비율 & ""                      '28
    txtNum(5).Value = ADORs!배상금액 & ""                      '29
                 
    Call cboSelectText(cboInput(3), ADORs!크레임구분 & "", 0)
    Call cboSelectText(cboInput(4), ADORs!보상구분 & "", 0)
    Call cboSelectText(cboInput(5), ADORs!처리구분 & "", 0)
'    cboInput(3).Text = ADORs!크레임구분 & ""                   '30
'    cboInput(4).Text = ADORs!보상구분 & ""                     '31
'    cboInput(5).Text = ADORs!처리구분 & ""                     '32
   
    txtNum(6).Value = ADORs!보상금액 & ""                      '33
   
    If IsNull(ADORs!처리일자) Then
        dtpDay(7).Value = ""                                   '34
    Else
        dtpDay(7).Value = Format(ADORs!처리일자, "YYYY-MM-DD") '34
    End If
    
    txtData(12).Text = ADORs!비고 & ""                         '35
    
    txtData(13).Text = ADORs!가맹점의견 & ""                   '36
    txtData(14).Text = ADORs!지사의견 & ""                     '37
    txtData(15).Text = ADORs!본사의견 & ""                     '38
    
    pnlDate(0).Caption = Format(Left(ADORs!지사승인일시, 10), "YYYY-MM-DD")  '39
    pnlDate(1).Caption = Format(Left(ADORs!본사승인일시, 10), "YYYY-MM-DD") '40
    
    'dtpDay(8).Value = Format(Left(ADORs!지사승인일시, 10), "YYYY-MM-DD") '39
    'dtpDay(9).Value = Format(Left(ADORs!본사승인일시, 10), "YYYY-MM-DD") '40
    
    ADORs.Close
    Set ADORs = Nothing
    
    bSave = True
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub txtData_Change(Index As Integer)
    If txtData(Index).Text = "" Then
        txtData(Index).tag = ""
    End If
End Sub

Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Index = 13 Or Index = 14 Or Index = 15 Then
            
        Else
            
            KeyAscii = 0
            
            If Index = 6 Then
                Call btnTAG_Click
            End If
        End If
    End If
End Sub

'
Private Sub fpList1_DblClick()
    On Error GoTo ErrRtn
    
    With fpList1
        .Col = 0: .Row = .ListIndex: txtData(6).Text = Format(.ColList, "000-00-0000") '택번호
                                     txtData(6).tag = txtData(6).Text & ""
                                     
        .Col = 3: .Row = .ListIndex: dtpDay(3).Value = Trim(.ColList)                  '접수일자
        
        If txtData(6).Text = "" Then
            .Visible = False
            
            txtData(6).SetFocus
            Exit Sub
        End If
        
        .Visible = False
        
        '-------------------------------------------------------------
        ' TB_입출고
        '-------------------------------------------------------------
        Query = "SELECT    A.* "
        Query = Query & ", B.성명"
        Query = Query & ", B.전화번호"
        Query = Query & ", B.휴대전화"
        Query = Query & ", B.주소"
        Query = Query & " FROM TB_입출고 AS A LEFT OUTER JOIN TB_고객정보 AS B ON A.고객코드 = B.고객코드"
        Query = Query & " WHERE 택번호   = '" & Replace(txtData(6).Text, "-", "") & "'"
        Query = Query & "   AND 접수일자 = '" & Format(dtpDay(3).Value, "YYYY-MM-DD") & "'"
        Query = Query & "   AND 판매취소 <> 'Y' "
        Set ADORs = New ADODB.RecordSet
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
        If Not ADORs.EOF Then
            txtData(0).Text = ADORs!고객코드 & ""                      ' 5
            txtData(1).Text = ADORs!성명 & ""                          ' 6
            txtData(2).Text = ADORs!전화번호 & ""                      ' 7
            txtData(3).Text = ADORs!휴대전화 & ""                      ' 8
            txtData(4).Text = ADORs!주소 & ""                          ' 9
            
            txtData(6).Text = ADORs!택번호 & ""                        '12
            dtpDay(3).Value = Format(ADORs!접수일자, "YYYY-MM-DD")     '11
            dtpDay(4).Value = Format(ADORs!지사출고일자, "YYYY-MM-DD") '13
            dtpDay(5).Value = Format(ADORs!출고일자, "YYYY-MM-DD")     '13
                   
            txtData(7).Text = ADORs!의류명 & ""                        '15
            txtData(8).Text = ADORs!상표 & ""                          '16
            txtData(9).Text = ADORs!색상 & ""                          '17
        End If
        ADORs.Close
        Set ADORs = Nothing
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
            txtData(6).SetFocus
    End Select
End Sub

Private Sub fpList1_LostFocus()
    fpList1.Visible = False
End Sub



Private Sub cmdPrintMini_Click()
    On Error GoTo ErrRtn
    
    Dim Print_Msg As String
    
    Dim tmp      As String
    Dim Cnt     As Long
   
    
    Print_Msg = Print_Msg & PrintString(가맹점정보.가맹점명 & " 사고 보고서", 4)
    Print_Msg = Print_Msg & PrintLineFeed

    Print_Msg = Print_Msg & PrintString("상 호 명 : " + 가맹점정보.가맹점명, 4)
    Print_Msg = Print_Msg & PrintString("전화번호 : " + 가맹점정보.전화매장, 4)
    
    Print_Msg = Print_Msg & PrintString("===============================================", 1)
    Print_Msg = Print_Msg & PrintString("사고접수일 : " + Format(dtpDay(2).Value, "YYYY년 MM월 DD일 "), 4)
    Print_Msg = Print_Msg & PrintString("-----------------------------------------------", 1)
    
    Print_Msg = Print_Msg & PrintString("고객명 : " + txtData(1).Text, 4)
    Print_Msg = Print_Msg & PrintString("전화번호 : " + txtData(2).Text, 4)
    Print_Msg = Print_Msg & PrintString("휴대전화 : " + txtData(3).Text, 4)
    Print_Msg = Print_Msg & PrintString("주소 : " + txtData(4).Text, 4)
    
    Print_Msg = Print_Msg & PrintString("-----------------------------------------------", 1)
    Print_Msg = Print_Msg & PrintString("-----------------------------------------------", 1)
    
    
    Print_Msg = Print_Msg & PrintLineFeed
    
    Print_Msg = Print_Msg & PrintCut
    
    Call frmKicc.Card_Print(Print_Msg)
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

' 사고품 접수시 SMS 발송 여부를 확인 하여 담당자에게 발송한다.
Private Function Send_SMS_사고품(sCode As String)
    Dim SMS_DataBase     As ADODB.Connection
    Dim sTel(2) As String
    
    On Error GoTo ERR_RTN
    
    If 가맹점정보.SMS_사고품 <> "Y" Then Exit Function
    
    ' SMS 서버에 연결 한다.
    If CheckSMSConnect(SMS_DataBase) = False Then
        MsgBox "SMS 서버와의 연결 설정을 확인하여 주십시요", vbInformation, "확인"
        Exit Function
    End If

    
    ' 서버 접속
    If Server_Connection(HostCon, "LAUNDRY1000") = False Then Exit Function
    
    '
    Query = "EXEC SP_01001_SMS_00 '1000', '" & 가맹점정보.가맹점코드 & "' "
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, HostCon, adOpenForwardOnly, adLockReadOnly

    Do Until ADORs.EOF
        ReDim sValue(10)
        
        ' 전송 여부 (전송하지 않는 사람도 모두 조회 된다.)
        If ADORs.Fields("전송여부") & "" = 1 Then
        
            If CheckMobileNumber(ADORs.Fields("휴대폰번호") & "", sTel) = True Then
            
                                  sValue(0) = "1"                   '전송
                                  sValue(1) = "0"                   '메시지타입
                sprGrid.Col = 4:  sValue(2) = Trim(ADORs.Fields("휴대폰번호") & "") '수신번호
                                  sValue(3) = Trim(가맹점정보.전화SMS)             '발신번호
                                  
                                  sValue(4) = "[" & Trim(가맹점정보.가맹점명 & "]") & vbCrLf            '메시지
                                  sValue(4) = sValue(4) & "" & Trim(cboInput(3).Text) & vbCrLf    '크레임 구분
                                  sValue(4) = sValue(4) & "" & Format(dtpDay(2).Value, "MM-DD") & vbCrLf    '일자
                                  sValue(4) = sValue(4) & "" & Trim(txtData(1).Text) & vbCrLf               '고객명
                                  sValue(4) = sValue(4) & "" & IIf(Trim(txtData(3).Text) <> "", Trim(txtData(3).Text), Trim(txtData(2).Text)) & vbCrLf      '연락처
                                  sValue(4) = sValue(4) & "" & Format(Right(Trim(txtData(6).Text), 6), "@@-@@@@") & vbCrLf      '택번호
                                  sValue(4) = sValue(4) & "" & Trim(txtData(7).Text) & vbCrLf               '품명
                                  
                                  sValue(5) = "1000"                '지사코드(본사로 처리)
                                  sValue(6) = 가맹점정보.택코드     '가맹점코드
                sprGrid.Col = 10: sValue(7) = Trim(txtData(0).Text) '고객코드
                sprGrid.Col = 2:  sValue(8) = Trim(txtData(1).Text) '고객성명
                                  sValue(9) = 가맹점정보.가맹점코드 '참고5
                                  sValue(10) = "1"                  '참고6
                
                Query = "EXEC PRO_SMS_SEND"
                Query = Query & "  '" & sValue(0) & "'"  ' 1 FLAG
                Query = Query & ", '" & sValue(1) & "'"  ' 2 MSGTYPE
                Query = Query & ", '" & sValue(2) & "'"  ' 3 PHONE
                Query = Query & ", '" & sValue(3) & "'"  ' 4 CALLBACK
                Query = Query & ", '" & sValue(4) & "'"  ' 5 MSG
                Query = Query & ", '" & sValue(5) & "'"  ' 6 MASTERCODE
                Query = Query & ", '" & sValue(6) & "'"  ' 7 STORECODE
                Query = Query & ", '" & sValue(7) & "'"  ' 8 CUSTCODE
                Query = Query & ", '" & sValue(8) & "'"  ' 9 ETC4
                Query = Query & ", '" & sValue(9) & "'"  '10 ETC5
                Query = Query & ", '" & sValue(10) & "'" '11 ETC6
                
                If Dir(App.Path & "\NO_SMS.DAT", vbNormal) = "" Then
                    SMS_DataBase.Execute Query
                End If
                
            Else
                MsgBox "[" & ADORs.Fields("담당자명") & "]님의 번호를 확인 하여 주십시요", vbInformation, "확인"
            End If
        End If
        ADORs.MoveNext
    Loop
    ADORs.Close
    Exit Function
    
ERR_RTN:
    MsgBox Err.description

End Function



Private Sub ComboAdd()
    Dim SSQL    As String

    
    ReDim sValue(1)
    
    sValue(0) = "0"

'-----------------------------------------------------------------------
    ' 크래임 구분
    ' 탈색, 파손, 이염, 분실, 기타
    cboInput(3).AddItem "탈색"
    cboInput(3).AddItem "파손"
    cboInput(3).AddItem "이염"
    cboInput(3).AddItem "분실"
    cboInput(3).AddItem "수축"
    cboInput(3).AddItem "변형"
    cboInput(3).AddItem "기타"

    '------------------------------------------------------------------------
    '보상구분
    ' 수선, 물품이도후 일부보상, 현금, 제품, 복구
    cboInput(4).AddItem "수선"
    cboInput(4).AddItem "물품이도후 일부보상"
    cboInput(4).AddItem "현금"
    cboInput(4).AddItem "제품"
    cboInput(4).AddItem "복구"
    
    '------------------------------------------------------------------------
    '처리구분
    cboInput(5).AddItem "[001] 접수"
    cboInput(5).AddItem "[002] 진행중"
    cboInput(5).AddItem "[003] 처리완료"
    
    '------------------------------------------------------------------------
    ' 사고품 품목
    SSQL = "SELECT * FROM TB_사고품품목 "
    Set ADORs = New ADODB.RecordSet
    ADORs.Open SSQL, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    Do Until ADORs.EOF
        cboInput(0).AddItem "[" & ADORs!품목코드 & "] " & ADORs!품목명
        ADORs.MoveNext
    Loop
    
    ADORs.Close:    Set ADORs = Nothing
    
    
    '------------------------------------------------------------------------
    ' 사고품 용도
    SSQL = "SELECT * FROM TB_사고품용도 "
    Set ADORs = New ADODB.RecordSet
    ADORs.Open SSQL, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    Do Until ADORs.EOF
        cboInput(1).AddItem "[" & ADORs!용도코드 & "] " & ADORs!용도내용
        ADORs.MoveNext
    Loop
    
    ADORs.Close:    Set ADORs = Nothing

    
    '------------------------------------------------------------------------
    ' 사고품 소재
    SSQL = "SELECT * FROM TB_사고품소재 "
    Set ADORs = New ADODB.RecordSet
    ADORs.Open SSQL, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    Do Until ADORs.EOF
        cboInput(2).AddItem "[" & ADORs!소재코드 & "] " & ADORs!소재명
        ADORs.MoveNext
    Loop
    
    ADORs.Close:    Set ADORs = Nothing

End Sub
