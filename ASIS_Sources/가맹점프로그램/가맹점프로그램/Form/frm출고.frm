VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm출고 
   Caption         =   "세탁물 출고"
   ClientHeight    =   9825
   ClientLeft      =   4755
   ClientTop       =   2205
   ClientWidth     =   15390
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form43"
   MDIChild        =   -1  'True
   ScaleHeight     =   9825
   ScaleWidth      =   15390
   WindowState     =   2  '최대화
   Begin Threed.SSPanel pnlProg 
      Height          =   1215
      Left            =   135
      TabIndex        =   54
      Top             =   3990
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
      Picture         =   "frm출고.frx":0000
      Caption         =   " "
      BorderWidth     =   0
      BevelOuter      =   1
      BevelInner      =   2
      Alignment       =   4
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9825
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   17330
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm출고.frx":2FCB
      Begin XtremeSuiteControls.TabControl TabControl 
         Height          =   7200
         Left            =   15
         TabIndex        =   55
         Top             =   2610
         Width           =   12000
         _Version        =   851970
         _ExtentX        =   21167
         _ExtentY        =   12700
         _StockProps     =   68
         Appearance      =   3
         Color           =   16
         PaintManager.BoldSelected=   -1  'True
         PaintManager.HotTracking=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         PaintManager.ButtonMargin=   "2,4,2,4"
         ItemCount       =   2
         Item(0).Caption =   "미출고 내역"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "TabControlPage(0)"
         Item(1).Caption =   "출고 내역"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "TabControlPage(1)"
         Begin XtremeSuiteControls.TabControlPage TabControlPage 
            Height          =   6690
            Index           =   1
            Left            =   -69970
            TabIndex        =   58
            Top             =   480
            Visible         =   0   'False
            Width           =   11940
            _Version        =   851970
            _ExtentX        =   21061
            _ExtentY        =   11800
            _StockProps     =   1
            Page            =   1
            Begin FPSpreadADO.fpSpread sprChul2 
               Height          =   6045
               Left            =   45
               TabIndex        =   59
               Top             =   525
               Width           =   11850
               _Version        =   524288
               _ExtentX        =   20902
               _ExtentY        =   10663
               _StockProps     =   64
               AutoCalc        =   0   'False
               BackColorStyle  =   1
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
               MaxCols         =   13
               MaxRows         =   20
               OperationMode   =   2
               ScrollBars      =   2
               SelectBlockOptions=   0
               SpreadDesigner  =   "frm출고.frx":30BD
               UserResize      =   1
               VisibleCols     =   7
               VisibleRows     =   15
               Appearance      =   1
               HighlightHeaders=   1
               HighlightStyle  =   1
               ScrollBarStyle  =   2
            End
            Begin MSComCtl2.DTPicker dtpDay 
               Height          =   375
               Index           =   0
               Left            =   945
               TabIndex        =   60
               Top             =   90
               Width           =   1380
               _ExtentX        =   2434
               _ExtentY        =   661
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
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
               Height          =   360
               Index           =   1
               Left            =   2610
               TabIndex        =   61
               Top             =   90
               Width           =   1380
               _ExtentX        =   2434
               _ExtentY        =   635
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   56295427
               CurrentDate     =   40279
            End
            Begin XtremeSuiteControls.PushButton btnOutCancel 
               Height          =   420
               Left            =   10095
               TabIndex        =   64
               Top             =   60
               Width           =   1770
               _Version        =   851970
               _ExtentX        =   3122
               _ExtentY        =   741
               _StockProps     =   79
               Caption         =   " 고객출고 취소"
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
               Picture         =   "frm출고.frx":4C17
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0FF&
               BackStyle       =   0  '투명
               Caption         =   "반품환불"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   12
               Left            =   5220
               TabIndex        =   74
               Top             =   210
               Width           =   720
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0FF&
               BackStyle       =   0  '투명
               Caption         =   "세탁환불"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   11
               Left            =   4260
               TabIndex        =   73
               Top             =   210
               Width           =   720
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "출고일자:"
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
               Index           =   9
               Left            =   90
               TabIndex        =   63
               Top             =   195
               Width           =   810
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "~"
               Height          =   195
               Index           =   8
               Left            =   2415
               TabIndex        =   62
               Top             =   165
               Width           =   105
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00C0C0FF&
               BackStyle       =   1  '투명하지 않음
               Height          =   345
               Index           =   0
               Left            =   4110
               Top             =   120
               Width           =   1005
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00FFFFC0&
               BackStyle       =   1  '투명하지 않음
               Height          =   345
               Index           =   1
               Left            =   5100
               Top             =   120
               Width           =   1005
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage 
            Height          =   6690
            Index           =   0
            Left            =   30
            TabIndex        =   56
            Top             =   480
            Width           =   11940
            _Version        =   851970
            _ExtentX        =   21061
            _ExtentY        =   11800
            _StockProps     =   1
            Page            =   0
            Begin FPSpreadADO.fpSpread fpSpread1 
               Height          =   6000
               Left            =   45
               TabIndex        =   78
               Top             =   4905
               Visible         =   0   'False
               Width           =   11850
               _Version        =   524288
               _ExtentX        =   20902
               _ExtentY        =   10583
               _StockProps     =   64
               AutoCalc        =   0   'False
               BackColorStyle  =   1
               DAutoHeadings   =   0   'False
               DAutoSave       =   0   'False
               EditEnterAction =   2
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
               MaxCols         =   19
               MaxRows         =   20
               OperationMode   =   2
               SelectBlockOptions=   0
               SpreadDesigner  =   "frm출고.frx":51B1
               UserResize      =   1
               VisibleCols     =   7
               VisibleRows     =   15
               Appearance      =   1
               HighlightHeaders=   1
               HighlightStyle  =   1
               ScrollBarStyle  =   2
            End
            Begin FPSpreadADO.fpSpread sprChul 
               Height          =   6045
               Left            =   45
               TabIndex        =   57
               Top             =   525
               Width           =   11850
               _Version        =   524288
               _ExtentX        =   20902
               _ExtentY        =   10663
               _StockProps     =   64
               AutoCalc        =   0   'False
               BackColorStyle  =   1
               DAutoHeadings   =   0   'False
               DAutoSave       =   0   'False
               EditEnterAction =   2
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
               MaxCols         =   19
               MaxRows         =   20
               OperationMode   =   2
               SelectBlockOptions=   0
               SpreadDesigner  =   "frm출고.frx":7D78
               UserResize      =   1
               VisibleCols     =   7
               VisibleRows     =   15
               Appearance      =   1
               HighlightHeaders=   1
               HighlightStyle  =   1
               ScrollBarStyle  =   2
            End
            Begin XtremeSuiteControls.PushButton btnAllSelect 
               Height          =   420
               Left            =   10605
               TabIndex        =   68
               Top             =   60
               Width           =   1260
               _Version        =   851970
               _ExtentX        =   2222
               _ExtentY        =   741
               _StockProps     =   79
               Caption         =   " 전체선택"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
               Picture         =   "frm출고.frx":9A6D
            End
            Begin XtremeSuiteControls.PushButton cmdBtn 
               Height          =   420
               Left            =   8160
               TabIndex        =   77
               Top             =   60
               Width           =   2430
               _Version        =   851970
               _ExtentX        =   4286
               _ExtentY        =   741
               _StockProps     =   79
               Caption         =   " 입고내역 출력(&P)"
               Appearance      =   6
               Picture         =   "frm출고.frx":A0A7
            End
            Begin VB.Image img_정상 
               Height          =   270
               Left            =   7410
               Picture         =   "frm출고.frx":A7A1
               Top             =   165
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Image img_반품 
               Height          =   270
               Left            =   7050
               Picture         =   "frm출고.frx":AD4B
               Top             =   150
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Image img_요청 
               Height          =   270
               Left            =   6690
               Picture         =   "frm출고.frx":B2F5
               Top             =   150
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label lbl_Brand 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "상표 확인"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   195
               Left            =   5490
               TabIndex        =   75
               Top             =   180
               Visible         =   0   'False
               Width           =   1020
            End
            Begin VB.Image Image 
               Height          =   240
               Index           =   1
               Left            =   120
               Picture         =   "frm출고.frx":B89F
               Top             =   180
               Width           =   240
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "고객에게 출고되지 않은 세탁물입니다."
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   195
               Left            =   480
               TabIndex        =   69
               Top             =   195
               Width           =   4080
            End
         End
      End
      Begin Threed.SSPanel pnlCustom 
         Height          =   2145
         Left            =   15
         TabIndex        =   7
         Top             =   450
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   3784
         _Version        =   262144
         BackColor       =   16777215
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin XtremeSuiteControls.PushButton btnInternet 
            Height          =   405
            Left            =   7560
            TabIndex        =   79
            Top             =   877
            Width           =   1035
            _Version        =   851970
            _ExtentX        =   1826
            _ExtentY        =   714
            _StockProps     =   79
            Caption         =   "인터넷"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TextAlignment   =   1
            Appearance      =   6
            Picture         =   "frm출고.frx":C2A1
         End
         Begin Threed.SSPanel SSPanel1 
            Height          =   420
            Left            =   2340
            TabIndex        =   66
            Top             =   60
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   741
            _Version        =   262144
            BackColor       =   16777215
            Enabled         =   0   'False
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
            Begin VB.ComboBox cboClass 
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   45
               Locked          =   -1  'True
               Style           =   2  '드롭다운 목록
               TabIndex        =   67
               Top             =   45
               Width           =   2010
            End
         End
         Begin VB.TextBox txtMemo 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Left            =   1095
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   5
            Top             =   1275
            Width           =   6405
         End
         Begin VB.TextBox txtHP 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   5445
            TabIndex        =   1
            Top             =   465
            Width           =   2055
         End
         Begin VB.TextBox txtCode 
            Alignment       =   2  '가운데 맞춤
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
            Height          =   420
            Left            =   1095
            TabIndex        =   6
            Top             =   60
            Width           =   1260
         End
         Begin VB.TextBox txtAddress 
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
            Height          =   420
            Left            =   1095
            TabIndex        =   4
            Top             =   870
            Width           =   6405
         End
         Begin VB.TextBox txtName 
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
            Height          =   420
            Left            =   1095
            TabIndex        =   3
            Top             =   465
            Width           =   3330
         End
         Begin VB.TextBox txtTel 
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
            Height          =   420
            Left            =   5445
            TabIndex        =   0
            Top             =   60
            Width           =   2055
         End
         Begin XtremeSuiteControls.PushButton btnKeyBoard 
            Height          =   405
            Left            =   7560
            TabIndex        =   35
            Top             =   461
            Width           =   1035
            _Version        =   851970
            _ExtentX        =   1826
            _ExtentY        =   714
            _StockProps     =   79
            Caption         =   "자판 "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TextAlignment   =   1
            Appearance      =   6
            Picture         =   "frm출고.frx":C673
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   420
            Index           =   0
            Left            =   4410
            TabIndex        =   36
            Top             =   60
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
            Caption         =   "전화번호"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm출고.frx":CC0D
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
            Left            =   60
            TabIndex        =   37
            Top             =   465
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
            PictureBackground=   "frm출고.frx":CF4F
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
            TabIndex        =   38
            Top             =   870
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
            PictureBackground=   "frm출고.frx":D291
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
            TabIndex        =   39
            Top             =   60
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
            Caption         =   "고객코드"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm출고.frx":D5D3
            BorderWidth     =   0
            BevelOuter      =   0
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   810
            Index           =   5
            Left            =   60
            TabIndex        =   40
            Top             =   1275
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   1429
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
            PictureBackgroundStyle=   2
            PictureBackground=   "frm출고.frx":D915
            BorderWidth     =   0
            BevelOuter      =   0
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdEdit 
            Height          =   405
            Left            =   7560
            TabIndex        =   41
            Top             =   1710
            Width           =   1035
            _Version        =   851970
            _ExtentX        =   1826
            _ExtentY        =   714
            _StockProps     =   79
            Caption         =   " 수정"
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
            Picture         =   "frm출고.frx":DC57
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   420
            Index           =   1
            Left            =   4410
            TabIndex        =   42
            Top             =   465
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
            PictureBackground=   "frm출고.frx":E669
            BorderWidth     =   0
            BevelOuter      =   0
            PictureAlignment=   9
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnClear 
            Height          =   420
            Left            =   7560
            TabIndex        =   65
            ToolTipText     =   "F8..."
            Top             =   30
            Width           =   1035
            _Version        =   851970
            _ExtentX        =   1826
            _ExtentY        =   741
            _StockProps     =   79
            Caption         =   "지우개"
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
            Picture         =   "frm출고.frx":E9AB
         End
         Begin XtremeSuiteControls.PushButton btnReceipt 
            Height          =   405
            Left            =   7560
            TabIndex        =   70
            Top             =   1293
            Width           =   1035
            _Version        =   851970
            _ExtentX        =   1826
            _ExtentY        =   714
            _StockProps     =   79
            Caption         =   " 접수"
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
            Picture         =   "frm출고.frx":EF45
         End
      End
      Begin XtremeSuiteControls.TabControl TabControl1 
         Height          =   2145
         Left            =   8685
         TabIndex        =   8
         Top             =   450
         Width           =   6690
         _Version        =   851970
         _ExtentX        =   11800
         _ExtentY        =   3784
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
         Item(0).Control(0)=   "TabControlPage1"
         Item(1).Caption =   "실적"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "TabControlPage2"
         Item(2).Caption =   "사고"
         Item(2).ControlCount=   1
         Item(2).Control(0)=   "TabControlPage3"
         Begin XtremeSuiteControls.TabControlPage TabControlPage3 
            Height          =   2085
            Left            =   -69370
            TabIndex        =   25
            Top             =   30
            Visible         =   0   'False
            Width           =   6030
            _Version        =   851970
            _ExtentX        =   10636
            _ExtentY        =   3678
            _StockProps     =   1
            BackColor       =   255
            Page            =   2
            Begin FPSpreadADO.fpSpread sprClaim 
               Bindings        =   "frm출고.frx":F957
               Height          =   1935
               Left            =   90
               TabIndex        =   76
               Top             =   60
               Width           =   5880
               _Version        =   524288
               _ExtentX        =   10372
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
               SpreadDesigner  =   "frm출고.frx":F96B
               VisibleCols     =   9
               VisibleRows     =   200
               Appearance      =   1
               HighlightHeaders=   1
               HighlightStyle  =   1
               ScrollBarStyle  =   2
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage2 
            Height          =   2100
            Left            =   -69370
            TabIndex        =   9
            Top             =   30
            Visible         =   0   'False
            Width           =   6030
            _Version        =   851970
            _ExtentX        =   10636
            _ExtentY        =   3704
            _StockProps     =   1
            Page            =   1
            Begin FPSpreadADO.fpSpread sprYear 
               Height          =   1680
               Left            =   75
               TabIndex        =   10
               Top             =   330
               Width           =   3465
               _Version        =   524288
               _ExtentX        =   6112
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
               SpreadDesigner  =   "frm출고.frx":1037F
               HighlightHeaders=   1
               HighlightStyle  =   1
            End
            Begin FPSpreadADO.fpSpread sprHist 
               Height          =   1680
               Left            =   3570
               TabIndex        =   44
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
               SpreadDesigner  =   "frm출고.frx":1097E
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
               Left            =   3555
               TabIndex        =   12
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
               TabIndex        =   11
               Top             =   105
               Width           =   1470
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage1 
            Height          =   2085
            Left            =   630
            TabIndex        =   13
            Top             =   30
            Width           =   6030
            _Version        =   851970
            _ExtentX        =   10636
            _ExtentY        =   3678
            _StockProps     =   1
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
               TabIndex        =   14
               Text            =   "2010-12-31"
               Top             =   75
               Width           =   1305
            End
            Begin CSTextLibCtl.sidbEdit txtMisu 
               Height          =   375
               Left            =   4200
               TabIndex        =   15
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
               Left            =   4200
               TabIndex        =   16
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
               Left            =   4200
               TabIndex        =   17
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
            Begin XtremeSuiteControls.PushButton btnMisu 
               Height          =   375
               Left            =   5535
               TabIndex        =   23
               Top             =   75
               Width           =   390
               _Version        =   851970
               _ExtentX        =   688
               _ExtentY        =   661
               _StockProps     =   79
               Appearance      =   6
               Picture         =   "frm출고.frx":10F0F
            End
            Begin XtremeSuiteControls.PushButton btnMileage 
               Height          =   375
               Left            =   5535
               TabIndex        =   24
               Top             =   1245
               Width           =   390
               _Version        =   851970
               _ExtentX        =   688
               _ExtentY        =   661
               _StockProps     =   79
               Appearance      =   6
               Picture         =   "frm출고.frx":11921
            End
            Begin CSTextLibCtl.sidbEdit txtTotalNum 
               Height          =   375
               Index           =   0
               Left            =   945
               TabIndex        =   45
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
               TabIndex        =   46
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
               TabIndex        =   47
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
            Begin CSTextLibCtl.sidbEdit txtVisit 
               Height          =   375
               Left            =   945
               TabIndex        =   48
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
               Left            =   4200
               TabIndex        =   71
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
               Caption         =   "미환불금액:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   10
               Left            =   3165
               TabIndex        =   72
               Top             =   570
               Visible         =   0   'False
               Width           =   990
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "총매출액:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   1
               Left            =   90
               TabIndex        =   52
               Top             =   960
               Width           =   810
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "총입금액:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   2
               Left            =   90
               TabIndex        =   51
               Top             =   1350
               Width           =   810
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "총할인액:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   3
               Left            =   90
               TabIndex        =   50
               Top             =   1740
               Width           =   810
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "이용횟수:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   6
               Left            =   90
               TabIndex        =   49
               Top             =   570
               Width           =   810
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "누적 마일리지:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   7
               Left            =   2505
               TabIndex        =   21
               Top             =   1380
               Width           =   1650
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "사용가능 마일리지:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   5
               Left            =   2505
               TabIndex        =   20
               Top             =   975
               Width           =   1650
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "미수금액:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   4
               Left            =   3345
               TabIndex        =   19
               Top             =   180
               Width           =   810
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "등록일자:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   0
               Left            =   90
               TabIndex        =   18
               Top             =   180
               Width           =   810
            End
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   22
         Top             =   15
         Width           =   15360
         _ExtentX        =   27093
         _ExtentY        =   741
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   16711680
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
         Caption         =   "      세탁물 출고"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm출고.frx":12333
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm출고.frx":12559
            Top             =   -15
            Width           =   765
         End
      End
      Begin Threed.SSPanel pnlButton 
         Height          =   3300
         Left            =   12030
         TabIndex        =   26
         Top             =   6510
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   5821
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdDryRepay 
            Height          =   690
            Left            =   45
            TabIndex        =   27
            Top             =   1530
            Width           =   1590
            _Version        =   851970
            _ExtentX        =   2805
            _ExtentY        =   1217
            _StockProps     =   79
            Caption         =   "지사반품요청"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
         End
         Begin XtremeSuiteControls.PushButton cmdReturnRepay 
            Height          =   690
            Left            =   1695
            TabIndex        =   28
            Top             =   1530
            Width           =   1590
            _Version        =   851970
            _ExtentX        =   2805
            _ExtentY        =   1217
            _StockProps     =   79
            Caption         =   "반품환불"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
         End
         Begin XtremeSuiteControls.PushButton cmdCancel 
            Height          =   690
            Left            =   45
            TabIndex        =   29
            Top             =   2265
            Width           =   1590
            _Version        =   851970
            _ExtentX        =   2805
            _ExtentY        =   1217
            _StockProps     =   79
            Caption         =   " 판매취소(F6)"
            ForeColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm출고.frx":13123
         End
         Begin XtremeSuiteControls.PushButton btnExit 
            Height          =   690
            Left            =   1695
            TabIndex        =   30
            Top             =   2265
            Width           =   1590
            _Version        =   851970
            _ExtentX        =   2805
            _ExtentY        =   1217
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm출고.frx":136BD
         End
         Begin Threed.SSCheck chkRepair 
            Height          =   345
            Left            =   225
            TabIndex        =   33
            Top             =   240
            Visible         =   0   'False
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   609
            _Version        =   262144
            Font3D          =   3
            ForeColor       =   192
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
            Caption         =   "수선출고"
         End
         Begin XtremeSuiteControls.PushButton btnAccount 
            Height          =   690
            Left            =   1695
            TabIndex        =   34
            Top             =   795
            Width           =   1590
            _Version        =   851970
            _ExtentX        =   2805
            _ExtentY        =   1217
            _StockProps     =   79
            Caption         =   " 출고/결제(F7)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm출고.frx":13C57
         End
         Begin XtremeSuiteControls.PushButton btnStock 
            Height          =   690
            Left            =   45
            TabIndex        =   43
            Top             =   795
            Width           =   1590
            _Version        =   851970
            _ExtentX        =   2805
            _ExtentY        =   1217
            _StockProps     =   79
            Caption         =   "가맹점   입고처리"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm출고.frx":141F1
         End
         Begin FPSpreadADO.fpSpread sprTemp 
            Height          =   630
            Left            =   1695
            TabIndex        =   53
            Top             =   90
            Visible         =   0   'False
            Width           =   1575
            _Version        =   524288
            _ExtentX        =   2778
            _ExtentY        =   1111
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   20
            SpreadDesigner  =   "frm출고.frx":14ACB
            HighlightHeaders=   1
            HighlightStyle  =   1
         End
         Begin VB.Shape Shape 
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00C0C000&
            Height          =   690
            Left            =   45
            Shape           =   4  '둥근 사각형
            Top             =   60
            Width           =   1590
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   405
         Index           =   1
         Left            =   12030
         TabIndex        =   31
         Top             =   2610
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   714
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   0
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frm출고.frx":150DB
         Caption         =   " 오점 사진"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm출고.frx":1563F
         BevelOuter      =   0
         PictureAlignment=   9
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlPicture 
         Height          =   3465
         Left            =   12030
         TabIndex        =   32
         Top             =   3030
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   6112
         _Version        =   262144
         CaptionStyle    =   1
         ForeColor       =   16711680
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image imgCapture 
            Height          =   3405
            Left            =   30
            Stretch         =   -1  'True
            Top             =   30
            Width           =   3285
         End
      End
   End
End
Attribute VB_Name = "frm출고"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX = 10

Private strCusNo  As String
Private iMakerRow As Integer

Dim Tel_Flag      As Boolean ' 전화 번호에서 다시 전화 번호를 변경할때 스택 오류가 나느것을방지
Dim Search_Flag   As Boolean '

'Set_입출고수정 에서 설정
Dim s사용마일리지   As Double ' 판매취소할 경우 후불일 경우 미수금액에서 차감할 경우 사용한 마일리지가 있을 경우 반환이 되기 때문에 미수금액에서 반환된 마일리지 만큼 빼주기 위하여

Public Sub Get_FindData(Gbn As String, strFind As String)
    On Error GoTo ErrRtn
    
    sprChul.MaxRows = 0
    sprTemp.MaxRows = 0
    
    Search_Flag = True
    
    Query = "SELECT * FROM TB_고객정보"
    
    Select Case Gbn
        Case "Code"
            Query = Query & " WHERE 고객코드 = '" & strFind & "'"
            Query = Query & " ORDER BY 고객코드 ASC"
        
        Case "Tel"
            Query = Query & " WHERE (전화번호 LIKE '%" & strFind & "'"
            Query = Query & "   OR   휴대전화   LIKE '%" & strFind & "')"
            Query = Query & " ORDER BY 전화번호, 휴대전화 ASC"
            
        Case "Name"
            Query = Query & " WHERE 성명 LIKE '%" & strFind & "%'"
            Query = Query & " ORDER BY 성명 ASC"
        
        Case "Addr"
            Query = Query & " WHERE 주소 LIKE '%" & strFind & "%'"
            Query = Query & " ORDER BY 주소 ASC"
    End Select
            
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If ADORs.EOF Then
        ADORs.Close
        Set ADORs = Nothing
        
        MsgBox " 일치하는 전화번호가 없읍니다.", vbInformation, "확인"
        
        Search_Flag = False
        
        Exit Sub
    
    ElseIf ADORs.RecordCount = 1 Then
        txtCode.Text = ADORs!고객코드 & ""                            ' 1
        
        ADORs.Close
        Set ADORs = Nothing
        
        Call 고객정보_Display(txtCode.Text)
        
        Search_Flag = False
        
    ElseIf ADORs.RecordCount >= 2 Then
        ADORs.Close
        Set ADORs = Nothing
        
        
        frm고객검색.DataDisplay Query
        frm고객검색.Show 1
        
        Set frm고객검색 = Nothing
        
        If 고객정보.전화번호 = "Error" Then
            txtTel.SetFocus
            
            Search_Flag = False
            Exit Sub
        End If
        
        txtCode.Text = 고객정보.고객코드 & ""
            
        Debug.Print "고객정보_Display 1 --> " & Now
        Call 고객정보_Display(txtCode.Text)
        Debug.Print "고객정보_Display 2 --> " & Now
    
    End If
    
    Call 미출고_Display(txtCode.Text, chkRepair.Value, False)
    Call 출고_Display
    
    TabControl.SelectedItem = 0
    
    btnStock.Enabled = True
    btnAccount.Enabled = True
    cmdDryRepay.Enabled = True
    cmdReturnRepay.Enabled = True
    cmdCancel.Enabled = True
    
    
    txtCode.Locked = True
    txtTel.Locked = True
    txtHP.Locked = True
    txtName.Locked = True
    txtAddress.Locked = True
    txtMemo.Locked = True
    
    'sprChul.SetActiveCell 11, 1 ' Active Cell
    Search_Flag = False
    
    Exit Sub
    
ErrRtn:
    Search_Flag = False
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Public Sub 고객정보_Display(고객코드 As String)
    Dim CustRs  As ADODB.RecordSet

    Debug.Print "1 --> " & Now

    On Error GoTo ErrRtn
    
    TabControl1.SelectedItem = 0
    
    Query = "SELECT    고객코드"
    Query = Query & ", 성명"
    Query = Query & ", 전화번호"
    Query = Query & ", 휴대전화"
    Query = Query & ", 주소"
    Query = Query & ", 메모"
    Query = Query & ", ISNULL(미수금액,0) AS 미수금액"
    Query = Query & ", ISNULL(고객등급코드,'C') AS 고객등급코드"
    Query = Query & ", 등록일자"
    Query = Query & ", 이용횟수"
    Query = Query & ", 총접수금액"
    Query = Query & ", 누적마일리지"
    Query = Query & ", 사용가능마일리지"
    Query = Query & " FROM TB_고객정보"
    Query = Query & " WHERE 고객코드 = '" & 고객코드 & "'"
    Set CustRs = New ADODB.RecordSet
    CustRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    Debug.Print "2 --> " & Now

    If CustRs.EOF Then
        CustRs.Close
        Set CustRs = Nothing
        
        Call Text_Clear
        
        With 마일리지
            .사용가능마일리지 = 0
            .누적마일리지 = 0
        End With
    Else
        txtCode.Text = CustRs!고객코드 & ""                       ' 1
        txtName.Text = Trim(CustRs!성명) & ""                     ' 2
        txtTel.Text = Trim(CustRs!전화번호) & ""                  ' 3
        txtHP.Text = Trim(CustRs!휴대전화) & ""                   ' 4
        txtHP.Tag = txtHP.Text & ""                               '
        txtAddress.Text = Trim(CustRs!주소) & ""                  ' 5
        
        If CustRs!미수금액 >= 0 Then
            txtMisu.Value = CustRs!미수금액 & ""                  ' 6
            txtNoRepay.Value = 0                                  '
        Else
            txtMisu.Value = 0                                     ' 6
            txtNoRepay.Value = CustRs!미수금액 & ""               '
        End If
        
        txtMemo.Text = Trim(CustRs!메모) & ""                     ' 7
        txtRegistDay.Text = Format(CustRs!등록일자, "YYYY-MM-DD") ' 8
        
        Select Case CustRs!고객등급코드                           '
            Case "C":  pnlCustom.BackColor = vbBlue
            Case "D":  pnlCustom.BackColor = vbYellow
            Case "E":  pnlCustom.BackColor = vbRed
            Case Else: pnlCustom.BackColor = vbWhite
        End Select
        
        With cboClass                                            ' 9
            For i = 0 To .ListCount - 1
                If Left(.List(i), 1) = CustRs!고객등급코드 Then
                    .ListIndex = i
                    
                    Exit For
                End If
            Next i
        End With
            
        txtVisit.Value = CustRs!이용횟수 & ""                     '10
        
        With 마일리지
            .사용가능마일리지 = CustRs!사용가능마일리지 & ""      '11
            .누적마일리지 = CustRs!누적마일리지 & ""              '12
        End With
        
        'txtTotalNum(0).Value = CustRs!접수금액 & ""              '총접수금액
        'txtTotalNum(1).Value = CustRs!입금액 & ""                '총입금액
        'txtTotalNum(2).Value = CustRs!할인액 & ""                '할인금액
        'txtMisu.Value = CustRs!미수금액                          '미수금액
    
        CustRs.Close
        Set CustRs = Nothing
    
        '---------------------------------------------------------------
        ' TB_매출
        '---------------------------------------------------------------
        Query = "SELECT    ISNULL(SUM(접수금액),0) AS 접수금액"
        Query = Query & ", ISNULL(SUM(입금합계),0) AS 입금액"
        Query = Query & ", ISNULL(SUM(세트할인),0) AS 세트할인"
        Query = Query & ", ISNULL(SUM(에누리),0)   AS 에누리"
        
        'Query = Query & ", ISNULL(SUM(현금입금+카드입금+쿠폰입금),0) AS 입금액"
        'Query = Query & ", ISNULL(SUM(세트할인+에누리),0) AS 할인액"
    Debug.Print "3 --> " & Now
        
        Query = Query & "  FROM TB_매출"
        Query = Query & "  WHERE 고객코드 = '" & 고객코드 & "'"
        Set CustRs = New ADODB.RecordSet
        CustRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    Debug.Print "4 --> " & Now

        If CustRs.EOF Then
            txtTotalNum(0).Value = 0                   '총접수금액
            txtTotalNum(1).Value = 0                   '총입금액
            txtTotalNum(2).Value = 0                   '할인금액
            
            ' << 미수금액은 TB_고객정보의 미수금을 이용 >>
            'txtMisu.Value = 0                          '미수금액
        Else
            txtTotalNum(0).Value = CustRs!접수금액 & ""                                      '총접수금액
            txtTotalNum(1).Value = CustRs!입금액 & ""                                        '총입금액
            If CustRs!세트할인 + CustRs!에누리 & "" > 0 Then
                txtTotalNum(2).Value = CustRs!세트할인 + CustRs!에누리 & ""                       '할인금액
            Else
                txtTotalNum(2).Value = 0                       '할인금액
            End If
            
            ' << 미수금액은 TB_고객정보의 미수금을 이용 >>
            'txtMisu.Value = CustRs!접수금액 - (CustRs!입금액 + CustRs!세트할인 + CustRs!에누리) '미수금액
        End If
        CustRs.Close
        Set CustRs = Nothing
    End If
    Debug.Print "5 --> " & Now
    Call 이용실적_Display                            '
    Debug.Print "6 --> " & Now
    Call 최근접수_Display(txtCode.Text)              '최근 접수건수
    Debug.Print "7 --> " & Now
    Call 사고품_Display(txtCode.Text)                '
    Debug.Print "8 --> " & Now
        
    txtUseMileage.Value = 마일리지.사용가능마일리지  '
    txtTotalMileage.Value = 마일리지.누적마일리지    '
        
    txtCode.Locked = True
    txtTel.Locked = True
    txtHP.Locked = True
    txtName.Locked = True
    txtAddress.Locked = True
    txtMemo.Locked = True
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub s_Change_Color()
    Dim j As Integer

    With sprChul
        For i = 1 To .MaxRows
            For j = 1 To .MaxCols
                .Row = i
                .Col = j: .BackColor = "&H00FFFFFF"
            Next j
        Next i
        
        .Row = .ActiveRow
        For i = 1 To .MaxCols
            .Col = i: .BackColor = "&H00C0FFFF"
        Next i
    End With
End Sub

Private Sub btnAccount_Click()
    Dim iMoney    As Long
    Dim iCheck    As Integer
    
    On Error GoTo ErrRtn
    
    If (frm출고.sprChul.MaxRows = 0) And (txtMisu.Value = 0) Then
        Beep
        Exit Sub
    End If
    
    With sprChul
        '------------------------------------------------------------
        ' 출고가능여부 체크
        '------------------------------------------------------------
        For i = 1 To .MaxRows
            .Row = i
            .Col = 12
            If .Text = "1" Then
                .Col = 2
                If .Text = "0" Then
                    MsgBox "지사로부터 가맹점으로 입고되지 않아 고객에게 출고할 수 없습니다.", vbInformation, "확인"
                    Exit Sub
                End If
            End If
            
            DoEvents
        Next i
        
        '------------------------------------------------------------
        ' '확' 체크한 의류중 미결제한 금액
        '------------------------------------------------------------
        iMoney = 0
        
        For i = 1 To .MaxRows
            .Row = i
            .Col = 9
            If Trim(.Text) = "미불" Then
                .Col = 12
                If Trim(.Text) = "1" Then '"확"
                    .Col = 8: iMoney = iMoney + CCur(.Value)
                End If
            End If
            
            DoEvents
        Next i
        
        '------------------------------------------------------------
        ' 출고체크
        '------------------------------------------------------------
        iCheck = 0
        
        For i = 1 To .MaxRows
            .Row = i
            .Col = 12
            If .Text = "1" Then
                iCheck = iCheck + 1
            End If
            
            DoEvents
        Next i
    End With
        
    If (iCheck = 0) And (txtMisu.Value = 0) Then
        MsgBox "고객출고 선택한 품목도 없고, 미수금액도 없어 출고결제를 할 수 없습니다.", vbInformation, "확인"
        
        Exit Sub
    End If
   
    frm출고결제.txtMisu.Value = txtMisu.Value         ' 미수금액
    
    If txtMisu.Value = 0 Then                         ' 미수금액이 없는 경우
        frm출고결제.btnCard.Enabled = False           ' 카드결제   x
        frm출고결제.btnCash.Enabled = False           ' 현금영수증 x
        
        frm출고결제.txtIncome.Enabled = False         '
    End If
    
    If iMoney > txtMisu.Value Then
        frm출고결제.txtChulPay.Value = txtMisu.Value  ' '확' 체크한 의류중 미결제한 금액이 미수금액보다 큰 경우는 미수금액을 입력한다.
    Else
        frm출고결제.txtChulPay.Value = iMoney         ' '확' 체크한 의류중 미결제한 금액
    End If
    
    frm출고결제.Show 1
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub btnClear_Click()
    Call Text_Clear
    
    Search_Flag = False
    
    chkRepair.Value = ssCBUnchecked
    txtTel.SetFocus
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnInternet_Click()
    frmInternetDelivery.GetData
    frmInternetDelivery.Show vbModal
    DoEvents
    If btnInternet.Tag <> "" Then
        
        txtCode.Text = frmInternetDelivery.SELECTCODE
        Call 고객정보_Display(txtCode.Text)
    End If
End Sub

Private Sub btnKeyBoard_Click()
    Load frmKeyboard
    frmKeyboard.Left = frmMain.Left + 7000
    frmKeyboard.Top = frmMain.Top + 3000
    
    frmKeyboard.Show 1
End Sub

Private Sub btnMileage_Click()
    On Error GoTo ErrRtn
    
    frm마일리지.lblCode.Caption = txtCode.Text
    frm마일리지.Show 1
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub btnMisu_Click()
    frm미수금.lblCode.Caption = txtCode.Text & ""
    frm미수금.lblMisu.Caption = txtMisu.Value
    
    frm미수금.Show 1
End Sub

Private Sub btnOutCancel_Click()
    Dim 택번호   As String
    Dim 접수일자 As String
    Dim 접수번호 As Long
    
    Dim iCnt     As Integer
    
    On Error GoTo ErrRtn
        
    If sprChul2.MaxRows = 0 Then Exit Sub
    
    If Label2(8).Tag = "dudtjsgh" Or Label2(8).Tag = "cleanaid1996" Then
        ' 이전 일자 취소 가능

    Else
        With sprChul2
            For i = 1 To .MaxRows
                .Row = i
                .Col = 11
                
                If .Text = "1" Then '확인 체크
                    
                    .Col = 2
                    If Format(.Text, "YYYY-MM-DD") <> Format(Date, "yyyy-MM-dd") Then
                        MsgBox "당일 출고분만 취소가 가능 합니다.", vbInformation, "확인"
                        Exit Sub
                    End If
                End If
            Next i
        End With
    End If
    
    
    Rtn = MsgBox("선택된 물품을 '출고취소' 하시겠습니까?", vbInformation + vbYesNo, "확인")
   
    If Rtn = vbNo Then Exit Sub
   
    iCnt = 0
   
    With sprChul2
        For i = 1 To .MaxRows
            .Row = i
            .Col = 11
            
            If .Text = "1" Then '확인 체크
                iCnt = iCnt + 1
                
                .Col = 1:  접수일자 = Format(.Text, "YYYY-MM-DD")                     ' 1 일자
                
                .Col = 13:  택번호 = Replace(.Text, "-", "") & ""                     ' 4 택번호
                '.Col = 4:  택번호 = 가맹점정보.택코드 & Replace(.Text, "-", "") & "" ' 4 택번호
                
                .Col = 12: 접수번호 = Trim(.Text) & ""                                '12 접수번호
                
                Query = "UPDATE TB_입출고 SET 출고일자     = ''"
                Query = Query & "           , 출고시간 = ''"
                Query = Query & "           , 본사전송여부 = ''"
                Query = Query & "           , 반품환불일자 = ''"
                Query = Query & "           , 세탁환불일자 = ''"
                Query = Query & "           , 환불사유 = ''"
                Query = Query & " WHERE 접수일자 = '" & 접수일자 & "'"
                Query = Query & "   AND 택번호   = '" & 택번호 & "'"
                Query = Query & "   AND 접수번호 =  " & 접수번호
                ADOCon.Execute Query
                
                ' 반품 환불 인경우 처리하는 경우에 대해 확인이 필요함
                
                
                
                
            End If
        Next i
    End With
    
    If iCnt > 0 Then
        Call Get_FindData("Code", Trim(txtCode.Text)) ' 고객정보를 검색한다.
            
        Query = "출고취소를 정상적으로 처리하였습니다." & vbNewLine & vbNewLine
        Query = Query & "택번호를 확인하여 주십시요."
        
        MsgBox Query, vbInformation, "확인"
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub btnReceipt_Click()
    
    frm접수.고객정보_Display (txtCode.Text)
    DoEvents
    frm접수.SetFocus
End Sub

Private Sub btnStock_Click()
    Dim 접수일자 As String
    Dim 택번호   As String

    On Error GoTo ErrRtn

    If sprChul.MaxRows <= 0 Then Exit Sub

    Rtn = MsgBox("선택된 물품을 '입고처리' 하시겠습니까?", vbInformation + vbYesNo, "확인")
   
    If Rtn = vbNo Then Exit Sub

    With sprChul
        For i = 1 To .MaxRows
            .Row = i
            .Col = 12

            If .Text = "1" Then
                .Col = 2
                
                If .Text = "0" Then '이미 입고처리된것은 제외...
                    .Col = 1:  접수일자 = Format(.Text, "YYYY-MM-DD")
                    .Col = 18: 택번호 = Replace(.Text, "-", "")
                    '.Col = 4: 택번호 = 가맹점정보.택코드 & Replace(.Text, "-", "")
    
                    '-------------------------------------------------------------------------------------------
                    '
                    '-------------------------------------------------------------------------------------------
                    Query = "UPDATE TB_입출고 SET 가맹점입고일자 = '" & Format(Now, "YYYY-MM-DD hh:mm:ss") & "'"
                    Query = Query & "           , 가맹점입고구분   = '수동'"
                    Query = Query & "           , 본사전송여부   = ''"
                    Query = Query & " WHERE 고객코드 = '" & Trim(txtCode.Text) & "'"
                    Query = Query & "   AND 접수일자 = '" & 접수일자 & "'"
                    Query = Query & "   AND 택번호   = '" & 택번호 & "'"
                    ADOCon.Execute Query
                End If
            End If
        Next i
    End With

    Call Get_FindData("Code", Trim(txtCode.Text)) ' 고객정보를 검색한다.

    Exit Sub

ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub chkRepair_Click(Value As Integer)
    'Call Get_FindData("Code", Trim(txtCode.Text)) ' 고객정보를 검색한다.
End Sub

'-----------------------------------------
' 전체선택
'-----------------------------------------
Private Sub btnAllSelect_Click()
    With sprChul
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            
            If .Text = "" Then
                Exit Sub
            Else
                .Col = 12
                
                If Trim(.Text) = "0" Then
                    sprChul.Text = "1" '"확"
                    
                    .Row = i
                    .Row2 = i
                    .Col = 1
                    .Col2 = .MaxCols
                    .BlockMode = True
                    
                    .BackColor = &HC0FFFF
                    
                    .BlockMode = False
                Else
                    .Text = "0"
                    
                    .Row = i
                    .Row2 = i
                    .Col = 1
                    .Col2 = .MaxCols
                    .BlockMode = True
                    
                    .BackColor = vbWhite
                    
                    .BlockMode = False
                End If
            End If
        Next i
    End With
End Sub

 

Private Sub cmdBtn_Click()

    '------------------------------------------------------------------------
    ' 보관증출력
    '------------------------------------------------------------------------
    If txtCode.Text = "" Then Exit Sub
    
    
   
    Call 입고내역출력_Report

    
    Exit Sub

ErrRtn:
    btnAccount.Enabled = True

    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub cmdEdit_Click()
    If Trim(txtCode.Text) <> "" Then
        frm고객수정.txtCode.Text = txtCode.Text
        
        frm고객수정.Show 1
    End If
End Sub

' 환불 - Command5
Private Sub cmdReturnRepay_Click()
    Dim 세트상품키   As String
    Dim 세트상품구분 As String
    
    Dim 접수번호     As String
    
    Dim 접수일자     As String
    Dim 세탁금액     As Long
    Dim 택번호       As String
    Dim 환불일자     As String
    
    On Error GoTo ErrRtn

'-------------------------------------------------------------------------------------------
    i = Get_ConfirmCheck '
    
    If i <= 0 Then
        MsgBox "반품환불 물품을 선택한 후 반품환불를 하세요.", vbInformation, "확인"
        Exit Sub
    End If
'-------------------------------------------------------------------------------------------
    
    With sprChul
        For i = 1 To .MaxRows
            .Row = i
            .Col = 12
            If .Text = "1" Then
                .Col = 1:  접수일자 = Format(.Text, "YYYY-MM-DD")                ' 1
                If 접수일자 = Format(Date, "YYYY-MM-DD") Then
                    Query = "당일 접수분은 반품환불을 처리할수 없습니다." & vbNewLine & vbNewLine
                    Query = Query & " 판매취소 기능을 이용하여 주십시요."
                    
                    MsgBox Query, vbInformation, "확인"
                    Exit Sub
                End If
                .Col = 2
                If .Text <> "2" Then
                    Query = "반품처리는 지사에서 반품 입고된 제품만 처리할수 있습니다." & vbNewLine & vbNewLine
                    Query = Query & "지사반품요청을 진행하여 주십시요."
                    
                    MsgBox Query, vbInformation, "확인"
                    Exit Sub
                End If
            End If
        Next i
    End With
    
    
    frm환불사유.Show 1 '환불사유 입력

    If Rtn = 0 Then Exit Sub 'frm환불사유에서 취소버튼을 클릭한 경우
    
    ' 마감여부 확인
    If Get_일일마감여부(Format(Date, "YYYY-MM-DD")) = True Then
        MsgBox "일마감이 되었으므로 반품환불 정보는 익일로 저장이 됩니다.", vbInformation
        
        환불일자 = Format(DateAdd("d", 1, Date), "YYYY-MM-DD")
    Else
        환불일자 = Format(Date, "YYYY-MM-DD")
    End If
    
    
    For i = 1 To sprChul.MaxRows
        sprChul.Row = i
        
        sprChul.Col = 12
        If sprChul.Text = "1" Then
            'sprChul.Col = 4:  택번호 = 가맹점정보.택코드 & Replace(sprChul.Text, "-", "") ' 4
            
            sprChul.Col = 1:  접수일자 = Format(sprChul.Text, "YYYY-MM-DD")                ' 1
            sprChul.Col = 8:  세탁금액 = sprChul.Value                                     ' 8
            sprChul.Col = 14: 세트상품키 = Trim(sprChul.Text) & ""                         '14 세트키
            sprChul.Col = 15: 세트상품구분 = Trim(sprChul.Text) & ""                       '15 세트구분
            sprChul.Col = 17: 접수번호 = Trim(sprChul.Text) & ""                           '17 접수번호
            sprChul.Col = 18: 택번호 = Replace(sprChul.Text, "-", "")                      '18
            
            '-------------------------------------------------------------------------
            ' 입출고
            '-------------------------------------------------------------------------
            Query = "UPDATE TB_입출고 SET "
            Query = Query & "  반품환불일자 = '" & 환불일자 & Format(Now, " hh:mm:ss") & "'"
            Query = Query & ", 출고일자     = '" & 환불일자 & "'"
            Query = Query & ", 환불사유     = '" & 환불사유 & "'"
            Query = Query & ", 본사전송여부 = 'N' "
            Query = Query & " WHERE 접수일자 = '" & 접수일자 & "'"
            Query = Query & "   AND 택번호   = '" & 택번호 & "'"
            Query = Query & "   AND (판매취소 <> 'Y')"
            ADOCon.Execute Query
            
            Call Set_입출고수정("반품환불", 접수일자, 택번호, 세트상품키, 세트상품구분, CLng(접수번호))     ' 입출고 내용 수정
            
            sprChul.Col = 9
            If sprChul.Text = "미불" Then
                '-----------------------------------------------------------------
                ' TB_고객정보 - 미수금액
                '-----------------------------------------------------------------
                Query = "UPDATE TB_고객정보 SET 미수금액     = 미수금액 - " & 세탁금액
                Query = Query & "             , 본사전송여부 = ''"
                Query = Query & " WHERE 고객코드 = '" & Trim(txtCode.Text) & "'"
                ADOCon.Execute Query
            End If
            
            '---------------------------------------------------------------------
            ' TB_이용실적 - 입고금액
            '---------------------------------------------------------------------
            Query = "UPDATE TB_이용실적 SET 이용금액 = 이용금액 - " & 세탁금액
            Query = Query & " WHERE 고객코드 = '" & Trim(txtCode.Text) & "'"
            Query = Query & "   AND 연도     = '" & Left(접수일자, 4) & "'"
            ADOCon.Execute Query
        End If
    Next i
   
    Call Get_FindData("Code", Trim(txtCode.Text)) ' 고객정보를 검색한다.
        
    Query = "반품환불을 정상적으로 처리하였습니다." & vbNewLine & vbNewLine
    Query = Query & "택번호를 확인하여 주십시요."
    
    MsgBox Query, vbInformation, "확인"
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

'판매취소
Private Sub cmdCancel_Click()
    Dim 세트상품키   As String
    Dim 세트상품구분 As String
    
    Dim 택번호       As String
    Dim 접수일자     As String
    Dim MisuAmt      As Long
    
    Dim 세탁금액     As Long
    Dim 수선금액     As Long
 
 
    Dim 접수번호     As Long
    Dim iNum         As Long
    
    Dim 결제여부     As String
    
    On Error GoTo ErrRtn
        
        
    s사용마일리지 = 0
    '-----------------------------------------------------------------------------
    ' 일일마감 여부체크
    '-----------------------------------------------------------------------------
    i = Get_ConfirmCheck '
    
    If i = 9999 Then
        MsgBox "일일마감이 되어 판매취소를 할수 없습니다.", vbInformation, "확인"
        
        Exit Sub
    ElseIf i <= 0 Then
        MsgBox "판매취소 물품을 선택한 후 판매취소를 하세요.", vbInformation, "확인"
        
        Exit Sub
    End If
    
    Rtn = MsgBox("선택된 물품을 '판매취소' 하시겠습니까?", vbInformation + vbYesNo, "확인")
   
    If Rtn = vbNo Then Exit Sub
    판매취소현금반환 = False
    
    '-----------------------------------------------------------------------------
    ' 접수번호가 다른 접수내역을 동시에 판매취소를 하지 못함.
    '-----------------------------------------------------------------------------
    With sprChul
        iNum = 0
            
        For i = 1 To .MaxRows
            .Row = i
            
            .Col = 12
            If .Text = "1" Then '확인 체크
                .Col = 1:  접수일자 = Format(.Text, "YYYY-MM-DD") '
                .Col = 9:  결제여부 = Trim(.Text) & ""            '
                .Col = 17: 접수번호 = Trim(.Text) & ""            '
                        
                If (iNum <> 0) And (iNum <> 접수번호) Then
                    iNum = -1
                    Exit For
                End If
                
                iNum = 접수번호
            End If
        Next i
    End With
    
    If iNum = -1 Then
        MsgBox "접수번호가 다른 접수세탁물을 동시에 판매취소 할 수 없습니다.", vbInformation, "확인"
        
        Exit Sub
    End If
    
'*********************************************************************************
    
    '-----------------------------------------------------------------------------
    ' 접수결제(신용카드,현금영수증 취소) 처리 루틴
    '-----------------------------------------------------------------------------
    Dim 신용카드   As Boolean
    Dim 현금영수증 As Boolean
    
    신용카드 = False
    현금영수증 = False
    
    Query = "SELECT * FROM TB_신용카드승인"
    Query = Query & " WHERE 고객코드 = '" & txtCode.Text & "'"
    Query = Query & "   AND 접수번호 =  " & 접수번호
    Query = Query & "   AND SUBSTRING(메시지2,1,2)  = 'OK'"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If Not ADORs.EOF Then
        신용카드 = True
    End If
    ADORs.Close
    Set ADORs = Nothing
    
    '-----------------------------------------------------------------------------
    ' TB_현금영수증
    '-----------------------------------------------------------------------------
    Query = "SELECT * FROM TB_현금영수증"
    Query = Query & " WHERE 고객코드 = '" & txtCode.Text & "'"
    Query = Query & "   AND 접수번호 =  " & 접수번호
    Query = Query & "   AND SUBSTRING(메시지2,1,2)  = 'OK'"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If Not ADORs.EOF Then
        현금영수증 = True
    End If
    ADORs.Close
    Set ADORs = Nothing
    
'*********************************************************************************
' 신용카드 취소 창을 띠운다.
    If (신용카드 = True) Or (현금영수증 = True) Then
        결제취소여부 = False
        
        frm판매취소결제.lblCode.Caption = txtCode.Text & ""
        frm판매취소결제.lblNum.Caption = 접수번호 & ""
        frm판매취소결제.Data_Display
        
        frm판매취소결제.Show 1
        
        If 판매취소여부 = True Then
            Exit Sub
        End If
    Else
        결제취소여부 = True '신용카드, 현금영수증이 없다는 것은 현금결제이거나 미결제인 경우...
    End If
'*********************************************************************************
    
    For i = 1 To sprChul.MaxRows
        sprChul.Row = i
        
        sprChul.Col = 12
        If sprChul.Text = "1" Then '확인 체크
            'sprChul.Col = 4:  택번호 = 가맹점정보.택코드 & Replace(sprChul.Text, "-", "") & "" ' 4 택번호
            
            sprChul.Col = 1:  접수일자 = Format(sprChul.Text, "YYYY-MM-DD")                     ' 1 일자
            sprChul.Col = 8:  세탁금액 = sprChul.Value                                          ' 8 세탁비
            sprChul.Col = 14: 세트상품키 = Trim(sprChul.Text) & ""                              '14 세트키
            sprChul.Col = 15: 세트상품구분 = Trim(sprChul.Text) & ""                            '15 세트구분
            sprChul.Col = 17: 접수번호 = Trim(sprChul.Text) & ""                                '17 접수번호
            sprChul.Col = 18: 택번호 = Replace(sprChul.Text, "-", "") & ""                      '18 택번호
            '---------------------------------------------------------------------------------------------------
            
            Call Set_입출고수정("판매취소", 접수일자, 택번호, 세트상품키, 세트상품구분, 접수번호)                  ' 입출고 내용 수정
            Call Get_고객정보(txtCode.Text)                                                    ' 고객정보
            
            
            ' 마일리지 결제가 있을 경우 처리순서
            ' 1.마일리지-> 2. 미수금 -> 현금
            ' Set_입출고수정 에서 이미 마일리지가 환원이 되기 때문에 미수금액을을 마일리지 뺀금액으로 처리한다.
            
            sprChul.Col = 9
            If sprChul.Text <> "완불" And 세탁금액 > 0 Then
                
                '-------------------------------------------------------------------------------------------
                ' TB_고객정보 - 미수금액
                '-------------------------------------------------------------------------------------------
                If Val(txtMisu.Value) >= (세탁금액 - s사용마일리지) Then
                    Query = "UPDATE TB_고객정보 SET 미수금액     =   " & Val(고객정보.미수금액) - (세탁금액 - s사용마일리지)
                    Query = Query & "             , 최종거래일자 = '" & Format(Now, "YYYY-MM-DD hh:mm:ss") & "'"
                    Query = Query & "             , 본사전송여부 = ''"
                    Query = Query & " WHERE 고객코드 = '" & Trim(txtCode.Text) & "'"
                    ADOCon.Execute Query
                
                Else
                    ' 2014-04-03일
                    ' 미수금액이 적을 경우 미수금액을 0원 처리한다.
                    ' 히스토리 표시 부분과 달라 질 수 있음. ㅠㅠ
                    If Val(txtMisu.Value) > 0 Then
                        Query = "UPDATE TB_고객정보 SET 미수금액     =   0 "
                        Query = Query & "             , 최종거래일자 = '" & Format(Now, "YYYY-MM-DD hh:mm:ss") & "'"
                        Query = Query & "             , 본사전송여부 = ''"
                        Query = Query & " WHERE 고객코드 = '" & Trim(txtCode.Text) & "'"
                        ADOCon.Execute Query
                        
                        MsgBox "고객 미수금액이 판매취소 상품의 금액보다 적어 미수금액을 0원으로 처리하였습니다.", vbInformation, "확인"
                    End If
                End If
            
            End If
            
            If 세탁금액 > 0 Then
                '-----------------------------------------------------------------
                ' TB_이용실적 - 입고금액
                '-----------------------------------------------------------------
                Query = "UPDATE TB_이용실적 SET 이용금액 = 이용금액 - " & 세탁금액
                Query = Query & " WHERE 고객코드 = '" & Trim(txtCode.Text) & "'"
                Query = Query & "   AND 연도     = '" & Left(접수일자, 4) & "'"
                ADOCon.Execute Query
            End If
            
        End If
    Next i
'*********************************************************************************
' 카드를 취소하고 다음 잔액이 남아 있는 경우
' 2014-04-03일 현금 부분 결제인 경우도 이창이 뜨는 문제가 있어서 And (신용카드 = True Or 현금영수증 = True) 코드 추가함
    If 결제취소여부 = True And (신용카드 = True Or 현금영수증 = True) Then
        Dim 미수금액 As Long
        
        Query = "SELECT ISNULL(SUM(접수금액) - SUM(입금합계),0) AS 미수금액"
        Query = Query & " FROM TB_매출"
        Query = Query & " WHERE 매출일자 = '" & 접수일자 & "'"
        Query = Query & "   AND 고객코드 = '" & txtCode.Text & "'"
        Query = Query & "   AND 접수번호 =  " & 접수번호
        Set ADORs = New ADODB.RecordSet
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        If ADORs.EOF Then
            미수금액 = 0
        Else
            미수금액 = ADORs!미수금액 & ""
        End If
        ADORs.Close
        Set ADORs = Nothing
        
        ' 2014-03-31
        ' 판매 취소하는 금액이 0원일 경우 처리하지 않는다. And 세탁금액 > 0 추가
        If (결제여부 = "완불") And (미수금액 > 0) And 세탁금액 > 0 Then
            frm판매취소.pnlCode.Caption = txtCode.Text & ""
            frm판매취소.pnlNum.Caption = 접수번호 & ""

            frm판매취소.txtMisu.Value = 미수금액

            Call frm판매취소.Data_Display(접수일자, 접수번호, txtCode.Text)

            frm판매취소.Show 1
            
        ElseIf 가맹점정보.CAT단말기종류 <> "KS4060 보안인증" Then

        End If
    End If
    
'*********************************************************************************
    
'    Dim CommPort As String
'    Dim BaudRate As String
'
'    CommPort = GetIniStr("VAN", "KS7500_CommPort", "", iniFile)
'    BaudRate = GetIniStr("VAN", "KS7500_BaudRate", "", iniFile)
'
'    Do
'        Rtn = KS7500i.CheckPort(CInt(CommPort), CLng(BaudRate))
'        DoEvents
'
'        If Rtn < 0 Then
'            i = i + 1
'
'            If i > 3 Then
'                MsgBox "카드단말기 장치가 연결되어 있지 않습니다", vbCritical, "오류"
'
'                Exit Do
'            End If
'        Else
'            Call KS7500i.SetConfig("", Rtn, CLng(BaudRate))    '첫번째 인자는 "" 로 넣어 준다.
'
'            KS7500i.InitPrint
'
'            Call 승인취소_Report(KS7500i, txtCode.Text, 접수번호)  '신용카드승인 취소, 현금영수증승인 취소 내역 출력
'
'            KS7500i.ClosePort
'            DoEvents
'        End If
'    Loop Until Rtn > 0
        
    Call Get_FindData("Code", Trim(txtCode.Text)) ' 고객정보를 검색한다.
        
    Query = "판매취소를 정상적으로 처리하였습니다." & vbNewLine & vbNewLine
    Query = Query & "택번호를 확인하여 주십시요."
    
    MsgBox Query, vbInformation, "확인"
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

'-----------------------------------------------------------------------------
' 확인 체크 갯수 반환 및 일일마감여부 확인
'
'-----------------------------------------------------------------------------
Private Function Get_ConfirmCheck() As Long
    Dim iCount As Long
    
    iCount = 0
    
    With sprChul
        For i = 1 To .MaxRows
            .Row = i
            
            .Col = 12
            If .Text = "1" Then
                iCount = iCount + 1 '확인 체크
                                
                .Col = 1
                If Get_일일마감여부(Format(.Text, "YYYY-MM-DD")) = True Then
                    Get_ConfirmCheck = 9999
                    
                    Exit Function
                End If
                
            End If
        Next i
    End With
    
    Get_ConfirmCheck = iCount
    
    Exit Function
    
ErrRtn:
    Get_ConfirmCheck = 0
End Function

'-----------------------------------------------------------------------------------------------------------
' 함수명 : Set_입출고수정
' 설  명 :
'-----------------------------------------------------------------------------------------------------------
Private Sub Set_입출고수정(구분 As String, 접수일자 As String, 택번호 As String, sGroupKey As String, sGroupGubun As String, 접수번호 As Long)
    Dim sPrtKey          As String
    Dim sData(5)         As String
    Dim iRow             As Integer
    Dim iRow2            As Integer
    Dim varTemp          As Variant
    
    Dim 일련번호         As Long
    Dim 의류명           As String
    Dim 세탁금액         As Long
    Dim 마일리지금액     As Long
    
    Dim 입금합계         As Long
    Dim 현금입금         As Long
    Dim 카드입금         As Long
    
    Dim 사용마일리지     As Long
    Dim 발생마일리지     As Long
    Dim 누적마일리지     As Long
    Dim 사용가능마일리지 As Long
    
    Dim ComputeValue     As Long
        
    On Error GoTo ErrRtn
    
    If 구분 = "판매취소" Then
        '-----------------------------------------------------------------------------------------
        ' 해당 품목을 판매 취소 처리한다.
        '-----------------------------------------------------------------------------------------
        Query = "UPDATE TB_입출고 SET 판매취소     = 'Y'"
        Query = Query & "           , 판매취소일자 = '" & Format(Now, "YYYY-MM-DD hh:mm:ss") & "'"
        Query = Query & "           , 본사전송여부 = 'N' "
        Query = Query & " WHERE 접수일자 = '" & 접수일자 & "'"
        Query = Query & "   AND 택번호   = '" & 택번호 & "'"
        Query = Query & "   AND 접수번호   = '" & 접수번호 & "'"
        ADOCon.Execute Query
    End If
    
    '------------------------------------------------------------
    ' TB_입출고
    '------------------------------------------------------------
    Query = "SELECT * FROM TB_입출고"
    Query = Query & " WHERE 접수일자 = '" & 접수일자 & "'"
    Query = Query & "   AND 택번호   = '" & 택번호 & "'"
    Query = Query & "   AND 접수번호   = '" & 접수번호 & "'"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF Then
        접수번호 = 0
        의류명 = ""
        세탁금액 = 0
        마일리지금액 = 0
    Else
        접수번호 = ADORs!접수번호        '
        의류명 = Trim(ADORs!의류명) & "" '
        세탁금액 = ADORs!금액 & ""       '세탁금액
        마일리지금액 = ADORs!마일리지 & ""   '발생마일리지
    End If
    ADORs.Close
    Set ADORs = Nothing
    
    '------------------------------------------------------------
    ' TB_고객정보 - 누적된 마일리지를 차감한다.
    '------------------------------------------------------------
    Call Fnc_마일리지삭제(Trim(txtCode.Text), CDbl(마일리지금액))
    Call Get_고객마일리지(Trim(txtCode.Text))
    txtTotalMileage.Value = 마일리지.누적마일리지
    txtUseMileage.Value = 마일리지.사용가능마일리지
    
    
    '------------------------------------------------------------
    ' TB_고객정보 - 총접수금액을 차감한다.
    '------------------------------------------------------------
    Query = "UPDATE TB_고객정보 SET 총접수금액   = 총접수금액   - " & 세탁금액
    Query = Query & " WHERE 고객코드 = '" & Trim(txtCode.Text) & "'"
    ADOCon.Execute Query
        
        
    '------------------------------------------------------------
    ' TB_매출
    '------------------------------------------------------------
    Query = "SELECT    ISNULL(SUM(사용마일리지),0) AS 사용마일리지"
    Query = Query & ", ISNULL(SUM(입금합계),0)     AS 입금합계"
    Query = Query & ", ISNULL(SUM(현금입금),0)     AS 현금입금"
    Query = Query & ", ISNULL(SUM(카드입금),0)     AS 카드입금"
    Query = Query & " FROM TB_매출"
    Query = Query & " WHERE 매출일자 = '" & 접수일자 & "'"
    Query = Query & "   AND 고객코드 = '" & Trim(txtCode.Text) & "'"
    Query = Query & "   AND 접수번호 =  " & 접수번호
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF Then
        사용마일리지 = 0                       '
        
        입금합계 = 0                           '
        현금입금 = 0                           '
        카드입금 = 0                           '
    Else
        사용마일리지 = ADORs!사용마일리지 & "" '
        
        입금합계 = ADORs!입금합계 & ""         '
        현금입금 = ADORs!현금입금 & ""         '
        카드입금 = ADORs!카드입금 & ""         '
    End If
    ADORs.Close
    Set ADORs = Nothing
    
    
    '------------------------------------------------------------
    ' 최근접수번호 - TB_매출
    '------------------------------------------------------------
    Query = "SELECT    ISNULL(발생마일리지,0)"
    Query = Query & " FROM TB_매출"
    Query = Query & " WHERE 매출일자 = '" & 접수일자 & "'"
    Query = Query & "   AND 고객코드 = '" & Trim(txtCode.Text) & "'"
    Query = Query & "   AND 접수번호 =  " & 접수번호
    Query = Query & "   AND 일련번호 = 0"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF Then
        발생마일리지 = 0
    Else
        발생마일리지 = ADORs(0)
    End If
    ADORs.Close
    Set ADORs = Nothing
         
    누적마일리지 = txtTotalMileage.Value   '
    사용가능마일리지 = txtUseMileage.Value '
        
        
    '-----------------------------------------------------------
    ' TB_매출
    '-----------------------------------------------------------
    Query = "SELECT ISNULL(MAX(일련번호),0) + 1 FROM TB_매출"
    Query = Query & " WHERE 고객코드 = '" & txtCode.Text & "'"
    Query = Query & "   AND 접수번호 =  " & 접수번호
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    일련번호 = ADORs(0)
    
    ADORs.Close
    Set ADORs = Nothing
    
    '
    Query = "SELECT * FROM TB_매출"
    Query = Query & " WHERE 고객코드 = '" & txtCode.Text & "'"
    Query = Query & "   AND 접수번호 =  " & 접수번호
    Query = Query & "   AND 일련번호 =  " & 일련번호
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenDynamic, adLockOptimistic
    
    If ADORs.EOF Then ADORs.AddNew
    
    ADORs!지사코드 = 가맹점정보.지사코드                       '
    ADORs!가맹점코드 = 가맹점정보.가맹점코드                   ' 0
    ADORs!고객코드 = txtCode.Text & ""                         ' 1
    ADORs!접수번호 = 접수번호 & ""                             ' 2
    ADORs!일련번호 = 일련번호                                  ' 3
    If 구분 = "판매취소" Then
    ' 마감후 접수한 내역을 판매취소할 경우 매출일자가 오늘날까로 되어서 금액이 맞지 않는 문제 수정
    ' 마감후 접수한 품목을 판매 취소할 경우에는 영업일자가 다음날이기 때문에 date로 처리하면 안된다.
    ' 판매취소일 경우 접수일자가 영업일자이기 때문에 가능함 마감후 오늘날짜는 취소가 안됨.
        ADORs!매출일자 = 접수일자
    Else
        ADORs!매출일자 = Format(Date, "YYYY-MM-DD") & ""           ' 4
    End If
    
    ADORs!매출시간 = Format(Now, "hh:mm:ss")                   ' 5
    
    If 구분 = "판매취소" And 판매취소현금반환 = True Then
        ADORs!적요 = "[" & 구분 & " 현금반환] " & 의류명 & ""               ' 6 "[판매취소] " & 의류명 & ""
    
    Else
        ADORs!적요 = "[" & 구분 & "] " & 의류명 & ""               ' 6 "[판매취소] " & 의류명 & ""
    End If
    ADORs!접수금액 = (세탁금액 * -1)                           ' 7
    
    '-------------------------------------------------------------
    
    ComputeValue = 0
    
    If (세탁금액 = 0) Or (결제취소여부 = False And 현금입금 = 0) Then '세탁금액이 0 원 경우
        ADORs!현금입금 = 0                                             ' 8
        ADORs!카드입금 = 0                                             ' 9
        ADORs!입금합계 = 0                                             '10
        ADORs!사용마일리지 = 0                                         '11
    Else
        If (현금입금 > 0) Or (카드입금 > 0) Then
            If (현금입금 > 0) And (카드입금 = 0) Then
                '*----------------------------------------------------------------
                '* 현금으로만 결제
                '*----------------------------------------------------------------
                
                If 세탁금액 <= 현금입금 Then
                    ADORs!현금입금 = (세탁금액 * -1)                           ' 8
                    ADORs!카드입금 = 0                                         ' 9
                    ADORs!입금합계 = (세탁금액 * -1)                           '10
                    ADORs!사용마일리지 = 0                                     '11
                    
                ElseIf 세탁금액 > 현금입금 Then
                    If (세탁금액 - 현금입금) <= 사용마일리지 Then
                        ComputeValue = (세탁금액 - 현금입금)
                    Else
                        ComputeValue = 사용마일리지
                    End If
                    
                    ADORs!현금입금 = (현금입금 * -1)                           ' 8
                    ADORs!카드입금 = 0                                         ' 9
                    ADORs!입금합계 = (현금입금 * -1)                           '10
                    ADORs!사용마일리지 = (ComputeValue * -1)                   '11
                End If
                
            ElseIf (현금입금 = 0) And (카드입금 > 0) Then
                '*----------------------------------------------------------------
                '* 카드로만 결제
                '*----------------------------------------------------------------
                
                If 세탁금액 <= 카드입금 Then
                    ADORs!현금입금 = 0                                         ' 8
                    ADORs!카드입금 = (카드입금 * -1)                           ' 9 카드인 경우에는 부분취소가 안되고, 전체 카드결제 취소한다.
                    ADORs!입금합계 = (카드입금 * -1)                           '10
                    ADORs!사용마일리지 = 0                                     '11
                                    
                ElseIf 세탁금액 > 카드입금 Then
                    If (세탁금액 - 카드입금) <= 사용마일리지 Then
                        ComputeValue = (세탁금액 - 카드입금)
                    Else
                        ComputeValue = 사용마일리지
                    End If
                    
                    ADORs!현금입금 = 0                                         ' 8
                    ADORs!카드입금 = (카드입금 * -1)                           ' 9
                    ADORs!입금합계 = (카드입금 * -1)                           '10
                    ADORs!사용마일리지 = (ComputeValue * -1)                   '11
                End If
            Else
                '*----------------------------------------------------------------
                '* 현금 + 카드로 결제
                '*----------------------------------------------------------------
                
                If 세탁금액 <= 카드입금 Then
                    ADORs!현금입금 = 0                                         ' 8
                    ADORs!카드입금 = (카드입금 * -1)                           ' 9
                    ADORs!입금합계 = (카드입금 * -1)                           '10
                    ADORs!사용마일리지 = 0                                     '11
                    
                Else
                    If (세탁금액 - 카드입금) < 현금입금 Then
                        ADORs!현금입금 = (세탁금액 - 카드입금) * -1                   ' 8
                        ADORs!카드입금 = (카드입금 * -1)                              ' 9
                        ADORs!입금합계 = (카드입금 * -1) + (세탁금액 - 카드입금) * -1 '10
                        ADORs!사용마일리지 = 0                                        '11
                        
                    Else
                        If (세탁금액 - 카드입금 - 현금입금) <= 사용마일리지 Then
                            ComputeValue = (세탁금액 - 카드입금 - 현금입금)
                        Else
                            ComputeValue = 사용마일리지
                        End If
                    
                        ADORs!현금입금 = (현금입금 * -1)                       ' 8
                        ADORs!카드입금 = (카드입금 * -1)                       ' 9
                        ADORs!입금합계 = (카드입금 * -1) + (현금입금 * -1)     '10
                        ADORs!사용마일리지 = (ComputeValue * -1)               '11
                    End If
                End If
            End If
        Else
            If 사용마일리지 > 0 Then
                If 세탁금액 <= 사용마일리지 Then
                    ComputeValue = 세탁금액
                Else
                    ComputeValue = 사용마일리지
                End If
                
                ADORs!현금입금 = 0                                             ' 8
                ADORs!카드입금 = 0                                             ' 9
                ADORs!입금합계 = 0                                             '10
                ADORs!사용마일리지 = (ComputeValue * -1)                       '11
                    
            Else
                ADORs!현금입금 = 0                                             ' 8
                ADORs!카드입금 = 0                                             ' 9
                ADORs!입금합계 = 0                                             '10
                ADORs!사용마일리지 = 0                                         '11
            End If
        End If
    End If
    
    s사용마일리지 = ComputeValue
    
    ADORs!쿠폰입금 = 0                                         '12
    ADORs!쿠폰번호 = ""                                        '13
    ADORs!세트할인 = 0                                         '14
    ADORs!에누리 = 0                                           '15
    ADORs!접수수량 = 0                                         '16
    ADORs!반품수량 = -1                                        '17
    ADORs!발생마일리지 = (마일리지금액 * -1)                       '18
    ADORs!누적마일리지 = 누적마일리지                           '19 이미 삭제되어서 나온 금액이다.
    ADORs!사용가능마일리지 = (사용가능마일리지 + ComputeValue) '20
    ADORs!본사전송여부 = ""                                    '
    
    ADORs.Update
    
    ADORs.Close
    Set ADORs = Nothing
    
    '---------------------------------------------------------------------------------
    ' 사용마일리지를 TB_고객정보에 원상복귀 시켜준다.
    '---------------------------------------------------------------------------------
    If ComputeValue <> 0 Then
        Query = "UPDATE TB_고객정보 SET"
        Query = Query & " 사용가능마일리지 = " & (사용가능마일리지 + ComputeValue)
        Query = Query & " WHERE 고객코드 = '" & txtCode.Text & "'"
        ADOCon.Execute Query
    End If
              
    If Trim(sGroupGubun) = "" Then Exit Sub ' 세트 상품이 아닐경우 해당 더이상 할것이 없다.
        
    '----------------------------------------------------
    ' 세트 상품이 취소 되었을 경우
    '----------------------------------------------------
    If sGroupKey <> "" And sGroupGubun <> "" Then
        Call DisplaySubSpread(접수일자, sGroupKey) ' 해당일자의 보관증번호(sGroupKey기준)을 임시 스프레드에 출력한다. (세트 상품 구성을 다시 하기 위하여)
        
        Call Chk_세트상품확인(sprTemp)             ' 세트 관련 내용을 다시 정리한다.
        
        세트상품정보.d세트Key = sGroupKey            '
        
        With sprTemp
            For iRow = 1 To .MaxRows
                Erase sData
                
                .Row = iRow
                
                .Col = 8:  If .Value = "" Then Exit For       ' 코드
                
                .Col = 2:  sData(0) = Replace(.Text, "-", "") ' 택번호
                .Col = 6:  sData(1) = Val(.Value)             ' 세트적용후 최종 수령금액(품목별)
                .Col = 11: sData(2) = .Value                  ' ex. 6-01, 5-01, 5-02
                .Col = 12: sData(3) = Val(.Value)             ' 세트 할인률을 기준으로 계산한 금액(10원단위 포함)
                .Col = 13: sData(4) = Val(.Value)             ' 원단위 절사후 다시 계산한 금액
                .Col = 14: sData(5) = Val(.Value)             ' 세트관련 원금액 기록
                
                '------------------------------------------------------------
                ' 입고 테이블에서 해당 제품의 내용을 수정한다.
                '------------------------------------------------------------
                Query = "UPDATE TB_입출고 SET"
                Query = Query & "  금액         = '" & sData(1) & "'"
                Query = Query & ", 세트구분     = '" & sData(2) & "'"
                Query = Query & ", 세트금액1    = '" & sData(3) & "'"
                Query = Query & ", 세트금액2    = '" & sData(4) & "'"
                Query = Query & ", 본사전송여부 = ''"
                
                'Query = Query & ", 정상금액 = '" & sData(5) & "'"
                'Query = Query & ", 세트Key  = '" & 세트상품정보.d세트Key & "'"
                
                Query = Query & " WHERE 접수일자 = '" & 접수일자 & "'"
                Query = Query & "   AND 택번호   = '" & sData(0) & "'"
                ADOCon.Execute Query
                
                For iRow2 = 1 To sprChul.MaxRows
                    sprChul.Row = iRow2
                    
                    sprChul.Col = 8: varTemp = sprChul.Value & "" '금액
                    
                    If CStr(varTemp) = "" Then Exit For
                                                            
                    sprChul.Col = 18: varTemp = sprChul.Text & ""  '택번호
                    'sprChul.Col = 4: varTemp = sprChul.Text & ""  '택번호
                    
                    If CStr(varTemp) = sData(0) Then
                        If sData(4) <> 0 Then
                            sprChul.Col = 8: sprChul.Text = sData(4) & "" '세트금액
                        Else
                            sprChul.Col = 8: sprChul.Text = sData(5) & "" '원래금액
                        End If
                        
                        Exit For
                    End If
                Next iRow2
            Next iRow
        End With
    End If
    
    Exit Sub

ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub DisplaySubSpread(sDate As String, sGroupKey As String)
    On Error GoTo ErrRtn
    
    Query = " SELECT * FROM TB_입출고"
    Query = Query & " WHERE 접수일자  = '" & sDate & "'"
    Query = Query & "   AND 세트Key   = '" & sGroupKey & "'"
    Query = Query & "   AND 판매취소 <> 'Y' "                '판매 취소분 제외
    Query = Query & " ORDER BY 택번호 ASC"
    Set SUBRs = New ADODB.RecordSet
    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprTemp
        .MaxRows = 0
        .ReDraw = False
        
        Do Until SUBRs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = SUBRs!의류명 & ""   '1
            .Col = 2: .Text = SUBRs!택번호 & ""   '2
            .Col = 3: .Text = sGroupKey & ""      '3
            .Col = 6: .Text = SUBRs!정상금액 & "" '6
            .Col = 8: .Text = SUBRs!의류코드 & "" '8
            
            SUBRs.MoveNext
        Loop
        .ReDraw = True
        
        SUBRs.Close
        Set SUBRs = Nothing
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

' 세탁환불 요청건
Private Sub cmdDryRepay_Click()
    Dim 세트상품키   As String
    Dim 세트상품구분 As String

    Dim 접수번호     As String
    Dim 접수일자     As String
    Dim 택번호       As String
    Dim 환불일자     As String
    Dim 세탁금액     As Long

    On Error GoTo ErrRtn

'-------------------------------------------------------------------------------------------
    i = Get_ConfirmCheck '

    If i <= 0 Then
        MsgBox "세탁환불 물품을 선택한 후 세탁환불를 하세요.", vbInformation, "확인"

        Exit Sub
    End If
'-------------------------------------------------------------------------------------------
    With sprChul
        For i = 1 To .MaxRows
            .Row = i
            .Col = 12
            If .Text = "1" Then
                .Col = 1:  접수일자 = Format(.Text, "YYYY-MM-DD")                ' 1

                If 접수일자 = Format(Date, "YYYY-MM-DD") Then
                    Query = "당일 접수분은 반품환불을 처리할수 없습니다." & vbNewLine & vbNewLine
                    Query = Query & " 판매취소 기능을 이용하여 주십시요."

                    MsgBox Query, vbInformation, "확인"
                    Exit Sub
                End If
                .Col = 2
                If .Text <> "1" Then
                    Query = "정상 입고 처리된건만 처리할수 있습니다." & vbNewLine & vbNewLine
                    Query = Query & "확인후 이용하여 주십시요."

                    MsgBox Query, vbInformation, "확인"
                    Exit Sub
                End If
            End If
        Next i
    End With

    frm환불사유.Show 1 '환불사유 입력

    If Rtn = 0 Then Exit Sub 'frm환불사유에서 취소버튼을 클릭한 경우


    For i = 1 To sprChul.MaxRows
        sprChul.Row = i

        sprChul.Col = 12
        If sprChul.Text = "1" Then
            sprChul.Col = 1:  접수일자 = Format(sprChul.Text, "YYYY-MM-DD")                ' 1
            sprChul.Col = 8:  세탁금액 = sprChul.Value                                     ' 8
            sprChul.Col = 14: 세트상품키 = Trim(sprChul.Text) & ""                         '14 세트키
            sprChul.Col = 15: 세트상품구분 = Trim(sprChul.Text) & ""                       '15 세트구분
            sprChul.Col = 17: 접수번호 = Trim(sprChul.Text) & ""                           '17 접수번호
            sprChul.Col = 18: 택번호 = Replace(sprChul.Text, "-", "")                      '18

            Call Set_반품요청(접수일자, 택번호, CLng(접수번호), 환불사유)

        End If
    Next i
'
    Call Get_FindData("Code", Trim(txtCode.Text)) ' 고객정보를 검색한다.
    Query = "반품요청을 정상적으로 처리하였습니다." & vbNewLine & vbNewLine
    Query = Query & "택번호를 확인하여 주십시요."

    MsgBox Query, vbInformation, "확인"
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub dtpDay_Change(Index As Integer)
    Call 출고_Display
End Sub

Private Sub 출고_Display()
    On Error GoTo ErrRtn
    
    '----------------------------------------------------------------------------
    ' 입고
    '----------------------------------------------------------------------------
    Query = "SELECT    접수일자"
    Query = Query & ", 출고일자"
    Query = Query & ", ISNULL(가맹점입고일자, '') AS 가맹점입고일자"
    Query = Query & ", 의류명"
    Query = Query & ", 택번호"
    Query = Query & ", 색상"
    Query = Query & ", 무늬"
    Query = Query & ", 내용"
    Query = Query & ", 금액"
    Query = Query & ", 결제여부"
    Query = Query & ", ISNULL(지사출고상태, '') AS 지사출고상태"
    Query = Query & ", 상표"
    Query = Query & ", 부모택번호"
    'Query = Query & ", ISNULL(확인,'0') AS 확인"
    Query = Query & ", 수선금액"
    Query = Query & ", 세트Key"
    Query = Query & ", 세트구분"
    Query = Query & ", 오점내용"
    Query = Query & ", 접수번호"
    Query = Query & ", 반품환불일자"
    Query = Query & ", 세탁환불일자"
    Query = Query & " FROM TB_입출고"
    Query = Query & " WHERE 고객코드  = '" & txtCode.Text & "'"
    Query = Query & "   AND (출고일자 >= '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "'"
    Query = Query & "   AND  출고일자 <= '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "')"
    Query = Query & "   AND ((판매취소 <> 'Y'))"
    'Query = Query & "   AND ((판매취소 = '') AND (판매취소일자 IS NULL OR 판매취소일자 = '')"
    ' 환불을 잘못 처리한것도 다시 돌리기 위하여 2012-11-28
    'Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')" '
    'Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprChul2
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            If Len(ADORs!세탁환불일자 & "") >= 10 Then
                .Col = -1
                .BackColor = Shape1(0).BackColor
            
            ElseIf Len(ADORs!반품환불일자 & "") >= 10 Then
                .Col = -1
                .BackColor = Shape1(1).BackColor
            End If
            
                                            
            .Col = 1:  .Text = Format(ADORs!접수일자, "YY-MM-DD") & ""            ' 1
            .Col = 2:  .Text = Format(ADORs!출고일자, "YY-MM-DD") & ""            ' 2
            .Col = 3:  .Text = ADORs!의류명 & ""                                  ' 3
            
            If Len(Trim(ADORs!택번호)) <= 6 Then
                .Col = 4: .Text = ADORs!택번호 & ""                               ' 4
            Else
                .Col = 4:  .Text = Mid(ADORs!택번호, 4, 2) & "-" & Mid(ADORs!택번호, 6, 4) ' 4
            End If
            
            .Col = 5:  .Text = ADORs!색상 & ""                                    ' 5
            .Col = 6:  .Text = ADORs!무늬 & ""                                    ' 6
            .Col = 7:  .Text = ADORs!내용 & ""                                    ' 7
            .Col = 8:  .Text = ADORs!금액 & ""                                    ' 8
            .Col = 9:  .Text = ADORs!결제여부 & ""                                ' 9
            
            .ForeColor = IIf(ADORs!결제여부 = "완불", vbBlue, vbRed)                                                 '
            .Col = 10: .Text = ADORs!상표 & ""                                    '10
            .Col = 11: .Text = "0"                                                '11
            .Col = 12: .Text = ADORs!접수번호 & ""                                '12
            .Col = 13: .Text = ADORs!택번호 & ""                                  '18 (전체 택번호 보여주기)
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    
    ActiveForm = "출고"
    
    Call Resize_Rtn
    
    Call Resize_Rtn
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call Resize_Rtn
End Sub

Private Sub Label2_Click(Index As Integer)
    If Index = 8 Then
        Dim sPass   As String
        
        
        Label2(8).Tag = InputBox("암호를 입력 하여 주십시요.", "암호")
        
    End If
End Sub

Private Sub sprChul_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    With sprChul
        .Row = Row
                
        If Col = 12 Then
            .Col = 12
            
            If .Text = "1" Then
                .Row = Row
                .Row2 = Row
                .Col = 1
                .Col2 = .MaxCols
                .BlockMode = True
            
                .BackColor = &HC0FFFF
            
                .BlockMode = False
            Else
                .Row = Row
                .Row2 = Row
                .Col = 1
                .Col2 = .MaxCols
                .BlockMode = True
            
                .BackColor = vbWhite
            
                .BlockMode = False
            End If
        End If
    End With
End Sub

Private Sub sprChul_LostFocus()
    Brand_Edit
End Sub

Private Sub sprChul2_Click(ByVal Col As Long, ByVal Row As Long)
    Dim 접수일자   As String
    Dim 결제확인   As String
    Dim 택번호      As String

    On Error GoTo ErrRtn
    
    With sprChul2
    
        If Col = 9 Then
            .Row = .ActiveRow
            
            .Col = 1
            If .Text = "" Then Exit Sub
            
            .Col = 1:   접수일자 = "20" & .Text
            .Col = 13:  택번호 = Replace(.Text, "-", "") & ""
            
            .Col = 9:   결제확인 = Trim(.Text) & ""
            
            결제확인 = IIf(결제확인 = "완불", "미불", "완불")
            .ForeColor = IIf(결제확인 = "완불", vbBlue, vbRed)
            .Col = 9: .Text = 결제확인
            
            
            '---------------------------------------------------------------------
            '
            '---------------------------------------------------------------------
            Query = " UPDATE TB_입출고 SET 결제여부         = '" & 결제확인 & "'"
            Query = Query & "            , 본사전송여부 = ''"
            Query = Query & " WHERE 접수일자 = '" & 접수일자 & "'"
            Query = Query & "   AND 고객코드 = '" & Trim(txtCode.Text) & "'"
            Query = Query & "   AND 택번호   = '" & 택번호 & "'"
            ADOCon.Execute Query
        End If
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

End Sub

Private Sub TabControl_BeforeItemClick(ByVal Item As XtremeSuiteControls.ITabControlItem, Cancel As Variant)
    If Item.Index = 0 Then
        btnStock.Enabled = True
        btnAccount.Enabled = True
        cmdDryRepay.Enabled = True
        cmdReturnRepay.Enabled = True
        cmdCancel.Enabled = True
    Else
        btnStock.Enabled = False
        btnAccount.Enabled = False
        cmdDryRepay.Enabled = False
        cmdReturnRepay.Enabled = False
        cmdCancel.Enabled = False
    
    
    End If
End Sub

Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo ErrRtn
    
    Select Case KeyCode
        Case 40: txtMemo.SetFocus  'Down Key
        Case 38: txtHP.SetFocus 'Up Key
    End Select
       
    If KeyCode = vbKeyReturn Then
        If Search_Flag = True Then Exit Sub
        
        Call Get_FindData("Addr", Trim(txtAddress.Text))
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
    
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 40: txtName.SetFocus 'Down Key
        Case 38: txtMemo.SetFocus 'Up Key
    End Select
End Sub

Private Sub txtHP_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 40: txtAddress.SetFocus  'Down Key
        Case 38: txtTel.SetFocus 'Up Key
    End Select
End Sub

Private Sub txtMemo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 40: txtCode.SetFocus  'Down Key
        Case 38: txtAddress.SetFocus 'Up Key
    End Select
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 40: txtTel.SetFocus  'Down Key
        Case 38: txtCode.SetFocus 'Up Key
    End Select
        
    If KeyCode = vbKeyReturn Then
        If Search_Flag = True Then Exit Sub
        
        Call Get_FindData("Name", Trim(txtName.Text)) ' 고객정보를 검색한다.
    End If
End Sub

Public Sub Text_Clear()
    pnlCustom.BackColor = vbWhite
    txtTel.Text = ""
    txtName.Text = ""
    txtHP.Text = ""
    txtHP.Tag = ""
    txtCode.Text = ""
    txtAddress.Text = ""
    txtMemo.Text = ""
    
    txtCode.Locked = False
    txtTel.Locked = False
    txtHP.Locked = False
    txtName.Locked = False
    txtAddress.Locked = False
    txtMemo.Locked = False
    
    txtVisit.Value = 0
    txtTotalNum(0).Value = 0
    txtTotalNum(1).Value = 0
    txtTotalNum(2).Value = 0
    
    txtRegistDay.Text = Date
    
    txtMisu.Value = 0
    txtNoRepay.Value = 0
    
    txtUseMileage.Value = 0
    txtTotalMileage.Value = 0
    
    btnInternet.Tag = ""
    '
    sprYear.MaxRows = 0
    sprHist.MaxRows = 0
    sprClaim.MaxRows = 0
    
    sprChul.MaxRows = 0
    sprChul2.MaxRows = 0
End Sub

Private Sub txtTel_Change()
    If Trim(txtTel.Text) = "" Then
        Call Text_Clear
    End If
    
    If Search_Flag = True Then Exit Sub
    
    If Len(txtTel.Text) >= 4 And Tel_Flag = False Then
        Tel_Flag = True
        
        'frm접수결제에서 Get_FindData 실행중에는 이중 실행안되도록...
        If bSearch = False Then
            Call Get_FindData("Tel", Trim(txtTel.Text)) ' 고객정보를 검색한다.
        End If
        
        Tel_Flag = False
    End If
End Sub

Private Sub txtTel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
        KeyAscii = 0
    End If
End Sub

'-----------------------------------------------------------
' Toggle
'-----------------------------------------------------------
Private Sub Toggle(Spread As fpSpread)
    Dim strOut    As String
    Dim strState  As String
    Dim intCurRow As Integer
    Dim intCurCol As Integer
     
    With Spread
        intCurRow = .ActiveRow
        intCurCol = .ActiveCol
        
        .Row = intCurRow
        .Col = intCurCol: strOut = .Text
        
        Select Case intCurCol
            Case 3
            
            Case 9
                If strOut = "완불" Then '完'
                    strOut = "미불"
                    .ForeColor = vbRed
                ElseIf strOut = "미불" Then '未'
                    strOut = "완불"
                    .ForeColor = vbBlue
                End If
                
                .Col = 9: .Text = strOut
                
                Call Edit_결제
                
            Case 12
        End Select
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ' 입고, 출고, 조회, 종료 체크
    'KeyChk (KeyCode)
    
    Select Case KeyCode
        Case 113:                      'F2 -
        Case 114:                      'F3-
        Case 115: cmdDryRepay_Click    'F4 -
        Case 116: cmdReturnRepay_Click 'F5 -
        Case 117: cmdCancel_Click      'F6 -
        Case 118: btnAccount_Click     'F7 -
                
        Case 119: btnClear_Click       'F8 -
        
        Case Else
            If Shift = 4 And KeyCode = 88 Then 'Alt+X
                Unload Me
            End If
    End Select
End Sub

Private Sub Form_Load()
    chkinputflig = "출고중" '현재 상태..
    DoEvents
    
    With sprChul
        .MaxRows = 0
        .RowHeight(-1) = 20
        
        .ColsFrozen = 4
        
        .Col = 13: .ColHidden = True '수선금액
        .Col = 14: .ColHidden = True '수선Key
        .Col = 15: .ColHidden = True '세트구분
            
        .Col = 17: .ColHidden = True '접수번호
        .Col = 18: .ColHidden = True  '택번호 (9자리 전체를 보여주는 Cell)
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle

        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeNormal
    End With
    
    With sprChul2
        .MaxRows = 0
        .RowHeight(-1) = 16
        
        .Col = 12: .ColHidden = True '접수번호
        .Col = 13: .ColHidden = True '택번호 (9자리 전체를 보여주는 Cell)
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle

        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeNormal
    End With
    
    With sprClaim
        .MaxRows = 0
        .RowHeight(-1) = 14
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeRow
    End With
    
    dtpDay(0).Value = Format(Date, "YYYY-MM-DD")
    dtpDay(1).Value = Format(Date, "YYYY-MM-DD")
    
    TabControl.SelectedItem = 0
    TabControl1.SelectedItem = 0
    
    txtRegistDay.Text = Date
    
    Call Resize_Rtn
    
    Call 고객등급_Display(cboClass, False) '고객등급
    
    cmdDryRepay.Enabled = IIf(가맹점정보.세탁환불여부 = "Y", True, False)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    chkinputflig = "메뉴" '현재 상태..
End Sub

Private Sub txtName_GotFocus()
    txtName.BackColor = vbYellow ' "&H0080FF80"
End Sub

Private Sub txtName_LostFocus()
    txtName.BackColor = "&H00FFFFFF"
End Sub

Private Sub txtTel_GotFocus()
    txtTel.SelStart = 0
    txtTel.SelLength = Len(txtTel.Text)
    
    txtTel.BackColor = vbYellow ' "&H0080FF80"
End Sub

Private Sub txtTEL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 40: txtHP.SetFocus   'Down Key
        Case 38: txtName.SetFocus 'Up Key
    End Select
    
    If KeyCode = vbKeyReturn Then
        If Search_Flag = True Then Exit Sub
        
        Call Get_FindData("Tel", Trim(txtTel.Text)) ' 고객정보를 검색한다.
        
        DoEvents
        If TabControl.SelectedItem = 0 Then sprChul.SetFocus
    End If
End Sub

Private Sub txtTel_LostFocus()
    txtTel.BackColor = "&H00FFFFFF"
End Sub

Private Sub txtAddress_GotFocus()
    txtAddress.BackColor = vbYellow '"&H0080FF80"
End Sub

Private Sub txtAddress_LostFocus()
    txtAddress.BackColor = "&H00FFFFFF"
End Sub

Private Sub sprChul_Click(ByVal Col As Long, ByVal Row As Long)
    Dim iMaxRow    As Integer
    Dim strData(3) As String
    
    On Error GoTo ErrRtn
    
    
    sprChul.Row = Row
    sprChul.Col = 10
    lbl_Brand.Caption = sprChul.Text
    
    
    Debug.Print Col
    
    If Row <= 0 Then Exit Sub
    
''    Dim imgFileName As String
''
''    sprChul.Row = Row
''    sprChul.Col = 1: imgFileName = Format(sprChul.Text, "YYYYMMDD")
''    sprChul.Col = 18: imgFileName = imgFileName & Replace(sprChul.Text, "-", "")
''    'sprChul.Col = 4: imgFileName = imgFileName & 가맹점정보.택코드 & Replace(sprChul.Text, "-", "")
''
''    If Dir(App.Path & "\Capture\" & imgFileName & ".jpg") = "" Then
''        imgCapture.Picture = LoadPicture()
''    Else
''        imgCapture.Picture = LoadPicture(App.Path & "\Capture\" & imgFileName & ".jpg")
''    End If
        
    '-------------------------------------------------------------------------------------------------
    ' 오점이미지 보여주기
    '-------------------------------------------------------------------------------------------------
    sprChul.Row = Row
    
'                      Query = "SELECT ISNULL(오점이미지,'') AS 오점이미지 FROM TB_입출고"
'    sprChul.Col = 1:  Query = Query & " WHERE 접수일자 = '" & Format(sprChul.Text, "YYYY-MM-DD") & "'"
'    sprChul.Col = 18: Query = Query & "   AND 택번호 = '" & Replace(sprChul.Text, "-", "") & "'"
'    Set ADORs = New ADODB.RecordSet
'    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
'
'    If ADORs.EOF Then
'        imgCapture.Picture = LoadPicture()
'    Else
'        If Trim(ADORs!오점이미지) = "" Then
'            imgCapture.Picture = LoadPicture()
'        Else
'            Dim ADOStream As New ADODB.Stream
'
'            With ADOStream
'              .Type = adTypeBinary
'              .Open
'              .Write ADORs!오점이미지
'
'              .SaveToFile AppPath & "Temp.JPG", adSaveCreateOverWrite
'            End With
'
'            Set ADOStream = Nothing
'
'            imgCapture.Picture = LoadPicture(AppPath & "Temp.JPG")
'        End If
'    End If
'    ADORs.Close
'    Set ADORs = Nothing
    
    '------------------------------------------------------------------
    ' 상표 수정일 경우 당일 일자를 확인한다.
    '------------------------------------------------------------------
    If Col = 10 Then
        sprChul.Row = Row
        sprChul.Col = 1
        
        If sprChul.Text = "" Then Exit Sub
        
        If sprChul.Text < Format(Date, "YY-MM-DD") Then
            MsgBox "상표 수정은 당일에만 가능 합니다." & Space(10), vbInformation, "확인"
            Exit Sub
        End If
    End If
    
    iMaxRow = 0
    
    For i = 1 To sprChul.MaxRows
        sprChul.Row = i
        sprChul.Col = 1
        
        If sprChul.Text = "" Then
            Exit For
        End If
        
        iMaxRow = iMaxRow + 1
    Next i
    
    If sprChul.ActiveRow > iMaxRow Then
        Exit Sub
    End If
    
    Call Toggle(sprChul)
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub sprChul_KeyDown(KeyCode As Integer, Shift As Integer)
    With sprChul
        '상표에서 엔터를 친경우
        If .ActiveCol = 10 And KeyCode = vbKeyReturn Then
            
            sprChul.EventEnabled(EventLeaveCell) = False
            
            Debug.Print "sprChul_KeyDown"
            Call Brand_Edit ' 브렌드를 저장한다.
            sprChul.EventEnabled(EventLeaveCell) = True
    
        ElseIf .ActiveCol <> 10 Then
            .Row = .ActiveRow
            
            .Col = 1
            If Trim(.Text) = "" Then
                Exit Sub
            End If
            
            '.SetActiveCell 11, .Row ' Active Cell
            
            If KeyCode = vbKeyReturn Then
                .Col = 12
                If .Text = "1" Then
                    .Text = "0"
                Else
                    .Text = "1"
                End If
                
                .Row = .ActiveRow + 1
                
                .SetActiveCell 12, .Row ' Active Cell
            End If
        End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' 해당 브렌드를 수정할 수 있도록 한다.
' 수정 유효일수는 당일에 한정한다.
'---------------------------------------------------------------------------------------
Private Sub Brand_Edit()
    Dim 상표   As String
    Dim 택번호 As String
    Dim sActionDate As String
    Dim vText   As Variant
    
    On Error GoTo ErrRtn
    
    With sprChul
        
        If .ActiveRow = 0 Then Exit Sub
        
        .GetText 1, .ActiveRow, vText
        If CStr(vText) = "" Then Exit Sub
        
        sActionDate = Format(CStr(vText), "YY-MM-DD")
        If sActionDate < Format(Date, "YY-MM-DD") Then
            'MsgBox "상표 수정은 당일에만 가능 합니다." & Space(10), vbInformation, "확인"
            Exit Sub
        End If
        
        .GetText 18, .ActiveRow, vText
        택번호 = Replace(vText, "-", "") & ""                       '택번호
        
        .GetText 10, .ActiveRow, vText
        .Col = 10: 상표 = Trim(Replace(CStr(vText), "'", "")) & ""  '상표
                   
                   
'        ' 상표가 "" 공백으로 수정이 되는경우 확인 메시지를 받는다.
'        Debug.Print "In Trim(lbl_Brand.Caption) => " & Trim(lbl_Brand.Caption)
'        If Trim(상표) = "" And Trim(lbl_Brand.Caption) <> "" Then
'            If MsgBox("상표를 삭제하시겠습니까?", vbInformation + vbYesNo) = vbNo Then
'                Debug.Print "Out Trim(lbl_Brand.Caption) => " & Trim(lbl_Brand.Caption)
'
'                .SetText 10, .ActiveRow, CVar(lbl_Brand.Caption)
'                Exit Sub
'            End If
'        End If
'
        '---------------------------------------------------------------------
        '
        '---------------------------------------------------------------------
        Query = " UPDATE TB_입출고 SET 상표         = '" & 상표 & "'"
        Query = Query & "            , 본사전송여부 = ''"
        Query = Query & " WHERE 접수일자 = '" & "20" & sActionDate & "'"
        Query = Query & "   AND 고객코드 = '" & Trim(txtCode.Text) & "'"
        Query = Query & "   AND 택번호   = '" & 택번호 & "'"
        ADOCon.Execute Query
        
        Debug.Print Query
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub


'---------------------------------------------------------------------------------------
' 해당 브렌드를 수정할 수 있도록 한다.
' 수정 유효일수는 당일에 한정한다.
'---------------------------------------------------------------------------------------
Private Sub Edit_결제()
    Dim 접수일자   As String
    Dim 결제확인   As String
    Dim 택번호 As String

    On Error GoTo ErrRtn
    
    With sprChul
        .Row = .ActiveRow
        
        .Col = 1
        If .Text = "" Then Exit Sub
        
        .Col = 1:   접수일자 = "20" & .Text
        .Col = 18:  택번호 = Replace(.Text, "-", "") & ""
        
        .Col = 9:   결제확인 = Trim(.Text) & ""
        
        '---------------------------------------------------------------------
        '
        '---------------------------------------------------------------------
        Query = " UPDATE TB_입출고 SET 결제여부         = '" & 결제확인 & "'"
        Query = Query & "            , 본사전송여부 = ''"
        Query = Query & " WHERE 접수일자 = '" & 접수일자 & "'"
        Query = Query & "   AND 고객코드 = '" & Trim(txtCode.Text) & "'"
        Query = Query & "   AND 택번호   = '" & 택번호 & "'"
        ADOCon.Execute Query
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub sprChul_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    ' 메이커에서 엔터를 친경우
    Debug.Print "sprChul_LeaveCell Col =  "; Col&; ", Row = " & Row & ", NewCol = " & NewCol & "NewRow=" & NewRow & ", iMakerRow = " & iMakerRow
    
    
    sprChul.EventEnabled(EventLeaveCell) = False
    
    sprChul.Row = Row
    sprChul.Col = 10
    lbl_Brand.Caption = sprChul.Text
    
    If NewRow <> -1 Then
        If Col = 10 Then
            Call Brand_Edit
        End If
    End If
    sprChul.EventEnabled(EventLeaveCell) = True

End Sub

Private Sub sprChul_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
    ' 메이커에서 엔터를 친경우
   ' Debug.Print "sprChul_LeaveCell Col =  "; Col&; ", NewCol = " & NewCol & ", iMakerRow = " & iMakerRow
    sprChul.EventEnabled(EventLeaveRow) = False
    
    
    Debug.Print "sprChul_LeaveRow"
    Call Brand_Edit
    
    Call Edit_결제
    sprChul.EventEnabled(EventLeaveRow) = True
    
End Sub

'*****************************************************************************************
' 제목    : 이용실적 표시
' 기능    : 고객의 이용실적을 전년과 올해로 나누어 표시
' 전해년도: 가맹점이 이전 자료가 있는 경우-conversion시에 write
' 올해년도: 계산시에 write 함
' 처리    : 이용실적 table 읽어 디스플레이
'*****************************************************************************************
Private Sub 이용실적_Display()
    On Error GoTo ErrRtn
    
    Query = "SELECT    연도"
    Query = Query & ", 이용금액"
    Query = Query & ", 이용횟수 "
    Query = Query & " FROM TB_이용실적 "
    Query = Query & " WHERE 고객코드 = '" & Trim(txtCode.Text) & "' "
    Query = Query & " ORDER BY 연도 DESC"
    Set SUBRs = New ADODB.RecordSet
    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    With sprYear
        .MaxRows = 0
        .ReDraw = False
        
        Do Until SUBRs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = SUBRs!연도 & ""
            .Col = 2: .Text = SUBRs!이용횟수 & ""
            .Col = 3: .Text = SUBRs!이용금액 & ""
                        
            SUBRs.MoveNext
        Loop
        SUBRs.Close
        Set SUBRs = Nothing
        
        If .MaxRows > 1 Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = "합계"
            .Col = 2: .Formula = "SUM(B1:B" & .MaxRows - 1 & ")"
            .Col = 3: .Formula = "SUM(C1:C" & .MaxRows - 1 & ")"
        End If
        
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

'최근 접수건수
Private Sub 최근접수_Display(sCode As String)
    On Error GoTo ErrRtn

    Query = "SELECT  "
    Query = Query & "  접수일자"
    Query = Query & ", ISNULL(SUM(금액),0)"
    Query = Query & " FROM TB_입출고"
    Query = Query & " WHERE 고객코드 = '" & sCode & "'"
    Query = Query & "   AND ((판매취소 <> 'Y')"
    'Query = Query & "   AND ((판매취소 = '') AND (판매취소일자 IS NULL OR 판매취소일자 = '')"
    Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
    Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
    Query = Query & " GROUP BY 접수일자"
    Query = Query & " ORDER BY 접수일자 DESC"
    Set SUBRs = New ADODB.RecordSet
    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprHist
        .MaxRows = 0
        .ReDraw = False
        
        Do Until SUBRs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = Format(SUBRs(0) & "", "YYYY-MM-DD") '
            .Col = 2: .Text = SUBRs(1) & ""                       '
            
            SUBRs.MoveNext
        Loop
        SUBRs.Close
        Set SUBRs = Nothing
        
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub


'사고품
Private Sub 사고품_Display(sCode As String)
    On Error GoTo ErrRtn
    
    If Len(Trim(txtHP.Text)) <= 4 Then Exit Sub
    
    Query = "SELECT    일련번호"
    Query = Query & ", 접수일자"
    Query = Query & ", 성명"
    Query = Query & ", 전화번호"
    Query = Query & ", 휴대전화"
    Query = Query & ", 의류명"
    Query = Query & ", 색상"
    Query = Query & ", 상표"
    Query = Query & ", 구입일자"
    Query = Query & ", 구입처"
    Query = Query & ", 구입형태"
    Query = Query & ", 구입가격"
    Query = Query & ", 사고접수일자"
    Query = Query & ", 크레임구분" '사고종류
    Query = Query & ", 가맹점의견" '사고내용
    Query = Query & ", 본사의견"   '사고의견
    Query = Query & ", 보상금액"
    'Query = Query & ", 합의금액"
    Query = Query & ", 처리구분"
    Query = Query & ", 가맹점코드"
    Query = Query & ", 가맹점명"
    Query = Query & " FROM TB_사고품내역"
    
    Query = Query & " WHERE 휴대전화 LIKE '" & Trim(txtHP.Text) & "'"
    
    Query = Query & " ORDER BY 접수일자, 일련번호 ASC"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprClaim
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = ADORs!일련번호 & ""
            .Col = 2:  .Text = ADORs!접수일자 & ""
            .Col = 3:  .Text = ADORs!보상금액 & ""
            .Col = 4:  .Text = ADORs!의류명 & ""
            .Col = 5:  .Text = ADORs!구입가격 & ""
            .Col = 6:  .Text = ADORs!색상 & ""
            .Col = 7:  .Text = ADORs!상표 & ""
            .Col = 8:  .Text = ADORs!구입일자 & ""
            .Col = 9:  .Text = ADORs!구입처 & ""
            .Col = 10: .Text = ADORs!구입형태 & ""
            .Col = 11: .Text = ADORs!사고접수일자 & ""
            .Col = 12: .Text = ADORs!크레임구분 & ""
            .Col = 13: .Text = ADORs!가맹점의견 & "" '사고내용
            .Col = 14: .Text = ADORs!본사의견 & ""   '사고의견
            .Col = 15: .Text = ADORs!보상금액 & ""
            .Col = 16: .Text = ADORs!처리구분 & ""
            .Col = 17: .Text = ADORs!가맹점코드 & ""
            .Col = 18: .Text = ADORs!가맹점명 & ""
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
            
        .ReDraw = True
        
        If .MaxRows > 0 Then
            TabControl1.SelectedItem = 3
        End If
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub 미출고_Display(strCode As String, Optional bRepair As Boolean = False, Optional bCustom As Boolean = True)
    
    On Error GoTo ErrRtn
    
    pnlProg.Visible = True
    DoEvents
    
    If bCustom = True Then
        '----------------------------------------------------------
        ' 고객정보
        '----------------------------------------------------------
        Query = "SELECT    고객코드"
        Query = Query & ", 성명"
        Query = Query & ", 전화번호"
        Query = Query & ", 주소"
        Query = Query & ", ISNULL(미수금액,0) AS 미수금액"
        Query = Query & ", 본사전송여부"
        Query = Query & ", 카드번호"
        Query = Query & ", 휴대전화"
        Query = Query & ", 문자발송여부"
        Query = Query & ", 등록일자"
        Query = Query & ", 메모"
        Query = Query & ", ISNULL(고객등급코드,'C') AS 고객등급코드"
        Query = Query & " FROM TB_고객정보 "
        Query = Query & " WHERE 고객코드 = '" & strCode & "' "
        Set ADORs = New ADODB.RecordSet
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        If ADORs.EOF Then
            ADORs.Close
            Set ADORs = Nothing
            
            MsgBox " 일치하는 회원이 없읍니다  다시입력요망", vbInformation, "확인"
            
            pnlProg.Visible = False
            Exit Sub
        Else
            txtTel.Text = ADORs!전화번호 & ""                           '
            txtCode.Text = ADORs!고객코드 & ""                          '
            txtAddress.Text = ADORs!주소 & ""                           '
            txtName.Text = ADORs!성명 & ""                              '
            txtHP.Text = ADORs!휴대전화 & ""                            '
            txtHP.Tag = ADORs!휴대전화 & ""                             '
            txtMemo.Text = ADORs!메모 & ""                              '
            txtRegistDay.Text = Format(ADORs!등록일자, "YYYY-MM-DD")    '
            
            '미수금액이 마이너스금액인 경우 미반환금액으로 표기
            If ADORs!미수금액 >= 0 Then
                txtMisu.Value = ADORs!미수금액 & ""                     '
                txtNoRepay.Value = 0                                    '
            Else
                txtMisu.Value = 0                                       '
                txtNoRepay.Value = ADORs!미수금액 & ""                  '
            End If
        End If
        ADORs.Close
        Set ADORs = Nothing
    End If
    
    If bRepair = False Then
        '----------------------------------------------------------------------------
        ' 입고
        '----------------------------------------------------------------------------
        Query = "SELECT    접수일자"
        Query = Query & ", ISNULL(가맹점입고일자, '') AS 가맹점입고일자"
        Query = Query & ", 의류명"
        Query = Query & ", 택번호"
        Query = Query & ", 색상"
        Query = Query & ", 무늬"
        Query = Query & ", 내용"
        Query = Query & ", 금액"
        Query = Query & ", 결제여부"
        Query = Query & ", ISNULL(지사출고상태, '') AS 지사출고상태"
        Query = Query & ", 상표"
        Query = Query & ", 부모택번호"
        Query = Query & ", 수선금액"
        Query = Query & ", 세트Key"
        Query = Query & ", 세트구분"
        Query = Query & ", 오점내용"
        Query = Query & ", 접수번호"
        Query = Query & ", 의류금액"
        Query = Query & " FROM TB_입출고"
        Query = Query & " WHERE 고객코드    = '" & strCode & "'"
        Query = Query & "   AND LEN(택번호) = 9"                                        '세탁접수 (수선접수는 제외)
        'Query = Query & "   AND SUBSTRING(택번호,1,3) = '" & 가맹점정보.택코드 & "'"   '세탁접수 (수선접수는 제외)
        Query = Query & "   AND (출고일자 IS NULL OR 출고일자 = '')"
        Query = Query & "   AND ((판매취소 <> 'Y')"
        'Query = Query & "   AND ((판매취소 = '') AND (판매취소일자 IS NULL OR 판매취소일자 = '')"
        Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
        Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
        Query = Query & " ORDER BY 접수일자 ASC, 택번호 ASC"
        Set ADORs = New ADODB.RecordSet
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        With sprChul
            .MaxRows = 0
            .ReDraw = False
            
            Do Until ADORs.EOF
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                
                .Col = 1:  .Text = Format(ADORs!접수일자, "YY-MM-DD") & ""            ' 1
                
                .Col = 3:  .Text = ADORs!의류명 & ""                                  ' 3
                .Col = 3:  .ForeColor = vbBlack
                
                If ADORs!가맹점입고일자 = "" Then
                    .Col = 2: .Text = "0"                                             ' 2
                    If ADORs!지사출고상태 = "3" Then
                        .Col = 2: .Text = "3": .TypePictPicture = img_요청 '.LoadPicture(App.Path & "\icon\반품요청256.bmp", PictureTypeBMP)
                        .Col = 3: .ForeColor = vbBlue
                    End If
                Else
                    Select Case ADORs!지사출고상태
                    Case "1"    '정상상태
                        .Col = 2: .Text = "1": .TypePictPicture = img_정상 '.LoadPicture(App.Path & "\icon\정상입고256.bmp", PictureTypeBMP)
                    Case "2"    '반품상태
                        .Col = 2: .Text = "2": .TypePictPicture = img_반품 '.LoadPicture(App.Path & "\icon\반품입고256.bmp", PictureTypeBMP)
                        .Col = 3: .ForeColor = vbRed
                    Case "3"    '반품요청상태
                        .Col = 2: .Text = "3": .TypePictPicture = img_요청 '.LoadPicture(App.Path & "\icon\반품요청256.bmp", PictureTypeBMP)
                        .Col = 3: .ForeColor = vbBlue
                    Case Else
                        .Col = 2: .Text = "1": .TypePictPicture = img_정상 '.LoadPicture(App.Path & "\icon\정상입고256.bmp", PictureTypeBMP)
                    End Select
                End If
                
                If Len(Trim(ADORs!택번호)) <= 6 Then
                    .Col = 4:  .Text = ADORs!택번호 & ""                              ' 4
                Else
                    .Col = 4:  .Text = Format(Right(ADORs!택번호, 6), "00-0000")      ' 4
                End If
                
                .Col = 5:  .Text = ADORs!색상 & ""                                    ' 5
                .Col = 6:  .Text = ADORs!무늬 & ""                                    ' 6
                .Col = 7:  .Text = ADORs!내용 & ""                                    ' 7
                .Col = 8:  .Text = ADORs!금액 & ""                                    ' 8
                .Col = 9:  .Text = ADORs!결제여부 & ""                                ' 9
                
                .Col = 9:   .ForeColor = IIf(ADORs!결제여부 = "완불", vbBlue, vbRed)
                
                .Col = 10:  .Text = ADORs!상표 & ""                                   '10
                
                If (Trim(ADORs!택번호) = Trim(ADORs!부모택번호)) Or (Trim(ADORs!부모택번호) = "") Then
                    .Col = 11: .Text = ""                                             '11
                Else
                    .Col = 11: .Text = Format(Right(ADORs!부모택번호, 6), "00-0000")  '11
                End If
                
                .Col = 12: .Text = "0"                                                '12
                .Col = 13: .Text = ADORs!수선금액 & ""                                '13
                .Col = 14: .Text = ADORs!세트Key & ""                                 '14
                .Col = 15: .Text = ADORs!세트구분 & ""                                '15
                .Col = 16: .Text = ADORs!오점내용 & ""                                '16
                .Col = 17: .Text = ADORs!접수번호 & ""                                '17
                .Col = 18: .Text = ADORs!택번호 & ""                                  '18 (전체 택번호 보여주기)
                .Col = 19: .Text = ADORs!의류금액 & ""                                  '19 정상금액
                            
                ADORs.MoveNext
            Loop
            ADORs.Close
            Set ADORs = Nothing
            
            .ReDraw = True
        End With
    Else
        '--------------------------------------------------------------------------------------
        ' TB_입출고
        '--------------------------------------------------------------------------------------
        Query = "SELECT    접수일자"
        Query = Query & ", ISNULL(가맹점입고일자, '') AS 가맹점입고일자"
        Query = Query & ", ISNULL(의류명,'') AS 의류명"
        Query = Query & ", 택번호"
        Query = Query & ", 색상"
        Query = Query & ", 무늬"
        Query = Query & ", 내용"
        Query = Query & ", 금액"
        Query = Query & ", 결제여부"
        Query = Query & ", ISNULL(지사출고상태,'') AS 지사출고상태"
        Query = Query & ", 상표"
        Query = Query & ", 부모택번호"
        Query = Query & ", 판매취소"
        Query = Query & ", 수선금액"
        Query = Query & ", 세트Key"
        Query = Query & ", 세트구분"
        Query = Query & ", 오점내용"
        Query = Query & ", 의류금액"
        Query = Query & " FROM TB_입출고"
        Query = Query & " WHERE 고객코드 = '" & strCode & "'"
        Query = Query & "   AND 내용 LIKE '%수%'"    '수선접수
        Query = Query & "   AND (출고일자 IS NULL OR 출고일자 = '')"
        Query = Query & "   AND ((판매취소 <> 'Y')"
        'Query = Query & "   AND ((판매취소 = '') AND (판매취소일자 IS NULL OR 판매취소일자 = '')"
        Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
        Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
        Query = Query & " ORDER BY 접수일자 ASC, 택번호 ASC"
        Set ADORs = New ADODB.RecordSet
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        With sprChul
            .MaxRows = 0
            .ReDraw = False
            
            Do Until ADORs.EOF
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                
                .Col = 1:  .Text = Format(ADORs!접수일자, "YY-MM-DD")       ' 1
                
                .Col = 2:  .Text = IIf(ADORs!가맹점입고일자 = "", "0", "1") ' 2
                
                
                .Col = 3:  .Text = ADORs!의류명 & ""                        ' 3
                .Col = 4:  .Text = ADORs!택번호 & ""                        ' 4
                .Col = 5:  .Text = ADORs!색상 & ""                          ' 5
                .Col = 6:  .Text = ADORs!무늬 & ""                          ' 6
                .Col = 7:  .Text = ADORs!내용 & ""                          ' 7
                .Col = 8:  .Text = ADORs!수선금액 & ""                      ' 8 - ADORs!금액
                .Col = 9:  .Text = ADORs!결제여부 & ""                      ' 9
                            
                            
                .ForeColor = IIf(ADORs!결제여부 = "완불", vbBlack, vbRed)
                .Col = 10:  .Text = ADORs!상표 & ""                         '10
                .Col = 11: .Text = ADORs!부모택번호 & ""                    '11
                .Col = 12: .Text = "0"                                      '12
                .Col = 13: .Text = ADORs!수선금액 & ""                      '13
                .Col = 14: .Text = ADORs!세트Key & ""                       '14
                .Col = 15: .Text = ADORs!세트구분 & ""                      '15
                .Col = 16: .Text = ADORs!오점내용 & ""                      '16
                .Col = 19: .Text = ADORs!의류금액 & ""                                  '19 정상금액
                ADORs.MoveNext
            Loop
            ADORs.Close
            Set ADORs = Nothing
            
            .ReDraw = True
        End With
    End If
    
    pnlProg.Visible = False
        
    Exit Sub
    
ErrRtn:
    pnlProg.Visible = False
End Sub

Private Sub Resize_Rtn()
    sprClaim.Width = TabControl1.Width - 810
    
    sprChul.Width = TabControl.Width - 150
    sprChul.Height = TabControl.Height - 1050
    
    sprChul2.Width = TabControl.Width - 150
    sprChul2.Height = TabControl.Height - 1050
    
    btnAllSelect.Left = TabControl.Width - btnAllSelect.Width - 120
    btnOutCancel.Left = TabControl.Width - btnOutCancel.Width - 120
    cmdBtn.Left = btnAllSelect.Left - cmdBtn.Width - 50
End Sub



Private Sub 입고내역출력_Report()
    On Error GoTo ErrRtn
    
    Dim ESC      As String * 1
    Dim CommPort As String
    Dim BaudRate As String
    
    Dim nRow     As Long
    Dim tmp      As String
    Dim PrintStr As String
    Dim 입고수량 As Integer
    Dim 전화번호출력 As String
    Dim Print_Msg As String
    

    전화번호출력 = GetIniStr("Printer", "TelPrint", "Y", iniFile)
    
    If 가맹점정보.지사코드 = M_COUPON_KLENZ_CODE Then '크렌즈갤러리
        Print_Msg = Print_Msg & PrintTitle2("크렌즈갤러리 - 세탁물 입고 내역")
    Else
        Print_Msg = Print_Msg & PrintTitle2("크린에이드 - 세탁물 입고 내역")
    End If
    
    Print_Msg = Print_Msg & PrintString("==============================================", 1, True)
    
    Print_Msg = Print_Msg & PrintCustomer(전화번호출력, txtName.Text, txtTel.Text, txtHP.Text, frm출고.txtAddress.Text)
    
    Print_Msg = Print_Msg & PrintString("==============================================", 1, True)
    Print_Msg = Print_Msg & PrintString("택번호  의류/상표         작업   색상     금액", 1, True)
    Print_Msg = Print_Msg & PrintString("----------------------------------------------", 1, True)
    
    입고수량 = 0
    
    With sprChul
        For nRow = 1 To .MaxRows
            .Row = nRow
            
            .Col = 2
            If Trim(.Text) = "1" Then
                입고수량 = 입고수량 + 1
                
                '*********************************************************
                '* 택번호
                '*********************************************************
                .Col = 4: PrintStr = .Text + " "
            
                '*********************************************************
                '* 품명
                '*********************************************************
                .Col = 3
                If LenH(.Text) >= 18 Then
                    tmp = MidH(.Text, 1, 18)
                Else
                    tmp = Trim(.Text) + String(18 - LenH(.Text), " ")
                End If
                
                PrintStr = PrintStr & tmp + ""
                
                '*********************************************************
                '* 내용
                '*********************************************************
                .Col = 7
                If LenH(.Text) >= 6 Then
                    tmp = MidH(.Text, 1, 6)
                Else
                    tmp = Trim(.Text) + String(6 - LenH(.Text), " ")
                End If
                
                PrintStr = PrintStr & tmp + " "
                
                '*********************************************************
                '* 색상
                '*********************************************************
                .Col = 5
                If LenH(.Text) >= 4 Then
                    tmp = MidH(.Text, 1, 4)
                Else
                    tmp = Trim(.Text) + String(4 - LenH(.Text), " ")
                End If
                
                PrintStr = PrintStr & tmp + " "
    
                '*********************************************************
                '* 금액
                '*********************************************************
                .Col = 8
                
                If Len(.Text) > 8 Then
                    PrintStr = PrintStr & .Text
                Else
                    PrintStr = PrintStr & String(8 - LenH(.Text), " ") + .Text
                End If
    
                Print_Msg = Print_Msg & PrintString(PrintStr, 1, True)
                
'                '*********************************************************
'                '* 상표
'                '*********************************************************
'                .Col = 10
'
'                If Trim(.Text) <> "" Then
'                    Call KS7500i.PrintString( "        - " + .Text ,1)
'                End If
'
'                '*********************************************************
'                '* 오점
'                '*********************************************************
'                .Col = 16
'
'                If Trim(.Text) <> "" Then
'                    Call KS7500i.PrintString( "        - " + .Text ,1)
'                End If
            End If
        Next nRow
    End With
    
    Print_Msg = Print_Msg & PrintString("-----------------------------------------------", 1, True)
    Print_Msg = Print_Msg & PrintString("출고수량 : " + String(9 - LenH(CStr(입고수량)), " ") + CStr(입고수량), 1, True)
    Print_Msg = Print_Msg & PrintString("===============================================", 1, True)
    

    
    Print_Msg = Print_Msg & PrintLineFeed(4)
    Print_Msg = Print_Msg & PrintCut
    
    Call frmKicc.Card_Print(Print_Msg)
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub Set_반품요청(접수일자 As String, 택번호 As String, 접수번호 As Long, 환불사유 As String)
On Error GoTo ErrRtn
    Query = "UPDATE TB_입출고 SET 지사출고상태 = '3'"
    Query = Query & "           , 가맹점입고구분 = '반품요청' "
    Query = Query & "           , 가맹점입고일자 = '' "
    Query = Query & "           , 오점내용 = '" & 환불사유 & "' "
    Query = Query & "           , 지사코드 = '" & 가맹점정보.지사코드 & "' "
    Query = Query & "           , 본사전송여부 = '' "
    Query = Query & " WHERE "
    Query = Query & "       가맹점코드 = '" & 가맹점정보.가맹점코드 & "'"
    Query = Query & "   AND 접수일자 = '" & 접수일자 & "'"
    Query = Query & "   AND 택번호   = '" & 택번호 & "'"
    Query = Query & "   AND 접수번호 = '" & 접수번호 & "'"
    If Set_반품요청_지사(Query) Then
        ADOCon.Execute Query
    End If
    Exit Sub
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
    Screen.MousePointer = 0
End Sub

Private Function Set_반품요청_지사(Query As String) As Boolean
On Error GoTo ErrRtn
    If Server_Connection(HostCon) = True Then
        HostCon.Execute Query
    End If
    Set_반품요청_지사 = True
    Exit Function
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
    Set_반품요청_지사 = False
    Screen.MousePointer = 0
End Function
