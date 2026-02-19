VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm고객현황 
   Caption         =   "고객 현황"
   ClientHeight    =   9285
   ClientLeft      =   435
   ClientTop       =   3600
   ClientWidth     =   15030
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
   Icon            =   "frm고객현황.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   15030
   WindowState     =   2  '최대화
   Begin Threed.SSPanel pnlProg 
      Height          =   1215
      Left            =   60
      TabIndex        =   34
      Top             =   1890
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
      Picture         =   "frm고객현황.frx":030A
      Caption         =   " "
      BorderWidth     =   0
      BevelOuter      =   1
      BevelInner      =   2
      Alignment       =   4
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9285
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   15030
      _ExtentX        =   26511
      _ExtentY        =   16378
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm고객현황.frx":32D5
      Begin Threed.SSPanel SSPanel2 
         Height          =   750
         Index           =   1
         Left            =   15
         TabIndex        =   16
         Top             =   450
         Width           =   15000
         _ExtentX        =   26458
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox cboGubun 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   915
            Style           =   2  '드롭다운 목록
            TabIndex        =   14
            Top             =   60
            Width           =   1455
         End
         Begin VB.TextBox txtFind 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2400
            TabIndex        =   0
            Top             =   60
            Width           =   3105
         End
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   630
            Left            =   5865
            TabIndex        =   1
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
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
            Picture         =   "frm고객현황.frx":3387
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   4
            Left            =   11850
            TabIndex        =   3
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
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
            Picture         =   "frm고객현황.frx":3A81
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   13440
            TabIndex        =   4
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
            Picture         =   "frm고객현황.frx":417B
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   3
            Left            =   10170
            TabIndex        =   2
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
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
            Picture         =   "frm고객현황.frx":520D
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "검색조건:"
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
            Index           =   4
            Left            =   45
            TabIndex        =   18
            Top             =   120
            Width           =   840
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   17
         Top             =   15
         Width           =   15000
         _ExtentX        =   26458
         _ExtentY        =   741
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   0
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
         Caption         =   "      고객 현황"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm고객현황.frx":5987
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm고객현황.frx":5BAD
            Top             =   15
            Width           =   765
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   390
         Left            =   15
         TabIndex        =   19
         Top             =   7095
         Width           =   15000
         _ExtentX        =   26458
         _ExtentY        =   688
         _Version        =   262144
         Font3D          =   3
         ForeColor       =   0
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
         Caption         =   " 고객 상세정보"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm고객현황.frx":6777
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton btnUpdate 
            Height          =   360
            Left            =   13680
            TabIndex        =   13
            Top             =   15
            Width           =   1290
            _Version        =   851970
            _ExtentX        =   2275
            _ExtentY        =   635
            _StockProps     =   79
            Caption         =   " 고객수정"
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
            Picture         =   "frm고객현황.frx":6999
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   1770
         Left            =   15
         TabIndex        =   20
         Top             =   7500
         Width           =   15000
         _ExtentX        =   26458
         _ExtentY        =   3122
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox cboSMS 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "frm고객현황.frx":73AB
            Left            =   960
            List            =   "frm고객현황.frx":73B5
            Style           =   2  '드롭다운 목록
            TabIndex        =   8
            Top             =   1425
            Width           =   2775
         End
         Begin VB.TextBox txtTel 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   10  '한글 
            Left            =   960
            TabIndex        =   6
            Top             =   735
            Width           =   2775
         End
         Begin VB.TextBox txtCode 
            Alignment       =   2  '가운데 맞춤
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            Text            =   "999999"
            Top             =   45
            Width           =   750
         End
         Begin VB.TextBox txtName 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   10  '한글 
            Left            =   960
            TabIndex        =   5
            Top             =   390
            Width           =   2775
         End
         Begin VB.TextBox txtHP 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   10  '한글 
            Left            =   960
            TabIndex        =   7
            Top             =   1080
            Width           =   2775
         End
         Begin VB.TextBox txtAdd 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   10  '한글 
            Left            =   4770
            TabIndex        =   9
            Top             =   45
            Width           =   3885
         End
         Begin VB.TextBox txtMemo 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            IMEMode         =   10  '한글 
            Left            =   4770
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   12
            Top             =   735
            Width           =   3885
         End
         Begin VB.ComboBox cboClass 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "frm고객현황.frx":73D3
            Left            =   2430
            List            =   "frm고객현황.frx":73D5
            Style           =   2  '드롭다운 목록
            TabIndex        =   21
            Top             =   45
            Width           =   1305
         End
         Begin CSTextLibCtl.sidbEdit txtMoney 
            Height          =   315
            Left            =   4770
            TabIndex        =   10
            Top             =   390
            Width           =   1215
            _Version        =   262145
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.01
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   14
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
         Begin FPSpreadADO.fpSpread sprMisu 
            Height          =   1680
            Left            =   10410
            TabIndex        =   35
            Top             =   45
            Width           =   4545
            _Version        =   524288
            _ExtentX        =   8017
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
            SpreadDesigner  =   "frm고객현황.frx":73D7
            Appearance      =   1
            HighlightHeaders=   1
            HighlightStyle  =   1
         End
         Begin CSTextLibCtl.sitxEdit txtCard 
            Height          =   315
            Left            =   7815
            TabIndex        =   11
            Top             =   390
            Width           =   840
            _Version        =   262145
            _ExtentX        =   1482
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   "______"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.76
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            EOLTab          =   -1  'True
            FocusSelect     =   -1  'True
            Insert          =   0   'False
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   "______"
            StartText.x     =   3
            StartText.y     =   3
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   "######"
            Justification   =   1
            CharacterTable  =   ""
            BorderStyle     =   0
            Characters      =   2
            MaxLength       =   6
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "원"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   180
            Index           =   7
            Left            =   6030
            TabIndex        =   33
            Top             =   450
            Width           =   180
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "고객코드:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   180
            Index           =   0
            Left            =   90
            TabIndex        =   32
            Top             =   105
            Width           =   825
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "성명:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   180
            Index           =   1
            Left            =   90
            TabIndex        =   31
            Top             =   450
            Width           =   825
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "전화번호:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   180
            Index           =   2
            Left            =   90
            TabIndex        =   30
            Top             =   795
            Width           =   825
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "휴대전화:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   180
            Index           =   3
            Left            =   90
            TabIndex        =   29
            Top             =   1140
            Width           =   825
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "문자발송:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   180
            Index           =   4
            Left            =   90
            TabIndex        =   28
            Top             =   1470
            Width           =   825
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "주소:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   180
            Index           =   5
            Left            =   3900
            TabIndex        =   27
            Top             =   105
            Width           =   825
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "미수금:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   180
            Index           =   6
            Left            =   3900
            TabIndex        =   26
            Top             =   450
            Width           =   825
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "카드번호:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   180
            Index           =   8
            Left            =   6945
            TabIndex        =   25
            Top             =   450
            Width           =   825
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "메모:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   180
            Index           =   9
            Left            =   3900
            TabIndex        =   24
            Top             =   750
            Width           =   825
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "등급:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   180
            Index           =   10
            Left            =   1890
            TabIndex        =   23
            Top             =   105
            Width           =   495
         End
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   5865
         Left            =   15
         TabIndex        =   36
         Top             =   1215
         Width           =   15000
         _Version        =   524288
         _ExtentX        =   26458
         _ExtentY        =   10345
         _StockProps     =   64
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
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         MaxCols         =   14
         Protect         =   0   'False
         SpreadDesigner  =   "frm고객현황.frx":7993
         VisibleCols     =   3
         VisibleRows     =   30
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "frm고객현황"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strCode As String

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 3: Call Export_Excel(frmMain.cdgExcel, sprGrid)
        Case 4:
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
    
    If sprGrid.MaxRows = 0 Then Exit Sub

    If Dir(AppPath & "XML", vbDirectory) = "" Then
        MkDir AppPath & "XML"
    End If

    Open AppPath & "XML\고객.XML" For Output As #1
    
    Print #1, "<?xml version=""1.0"" encoding=""EUC-KR""?>"
    Print #1, "<root>"
    
          XML = "    <조건>"
    
    If txtFind.Text = "" Then
        XML = XML & "        <검색조건>검색조건 : 전체</검색조건>"
    Else
        XML = XML & "        <검색조건>검색조건 : " & Func_Replace(txtFind.Text) & "</검색조건>"
    End If
    
    XML = XML & "        <가맹점>크린에이드 - " & Func_Replace(가맹점정보.가맹점명) & "</가맹점>"
    XML = XML & "   </조건>"
    Print #1, XML
    
    With sprGrid
        For i = 1 To .MaxRows
            .Row = i
            
                             XML = "    <Data>"
            .Col = 1:  XML = XML & "        <고객코드>" & .Text & "</고객코드>"
            .Col = 2:  XML = XML & "        <성명>" & Func_Replace(.Text) & "</성명>"
            .Col = 3:  XML = XML & "        <전화번호>" & Func_Replace(.Text) & "</전화번호>"
            .Col = 4:  XML = XML & "        <휴대전화>" & Func_Replace(.Text) & "</휴대전화>"
            .Col = 5:  XML = XML & "        <주소>" & Func_Replace(.Text) & "</주소>"
            .Col = 11: XML = XML & "        <미수금>" & .Text & "</미수금>"
            .Col = 6:  XML = XML & "        <SMS>" & .Text & "</SMS>"
            .Col = 7:  XML = XML & "        <등록일자>" & .Text & "</등록일자>"
            .Col = 8:  XML = XML & "        <고객등급>" & .Text & "</고객등급>"
            .Col = 12: XML = XML & "        <누적마일리지>" & .Text & "</누적마일리지>"
            .Col = 13: XML = XML & "        <사용가능마일리지>" & .Text & "</사용가능마일리지>"
                       XML = XML & "   </Data>"
                       Print #1, XML
        Next i
        
        Print #1, "</root>"
        Close #1
    End With
    
    If Print_PreView = True Then
        With rpt고객현황
            .dc.FileURL = AppPath & "XML\고객.XML"
            .Show 1
        End With
    Else
        With rpt고객현황
            .dc.FileURL = AppPath & "XML\고객.XML"
            .PrintReport False
        End With
        
        Unload rpt고객현황
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
    
    Screen.MousePointer = 0
End Sub

Private Sub cmdList_Click()
    On Error GoTo ErrRtn
    
    txtCode.Text = ""
    txtName.Text = ""
    txtTel.Text = ""
    txtHP.Text = ""
    txtAdd.Text = ""
    txtMoney.Value = 0
    txtCard.Text = ""
    txtMemo.Text = ""
    
    sprGrid.MaxRows = 0
    sprMisu.MaxRows = 0
    
    pnlProg.Visible = True
    DoEvents
    
    Query = "SELECT    A.고객코드"
    Query = Query & ", A.성명"
    Query = Query & ", A.전화번호"
    Query = Query & ", A.휴대전화"
    Query = Query & ", A.주소"
    Query = Query & ", A.미수금액"
    Query = Query & ", (CASE WHEN A.문자발송여부 = 'Y' THEN '1' ELSE '0' END) SMS"
    Query = Query & ", A.등록일자"
    Query = Query & ", A.고객등급코드"
    Query = Query & ", A.이용횟수"
    Query = Query & ", A.총접수금액"
    Query = Query & ", A.누적마일리지"
    Query = Query & ", A.사용가능마일리지"
    Query = Query & ", A.수정일자"
    Query = Query & ", B.고객등급명"
    Query = Query & " FROM TB_고객정보 AS A LEFT OUTER JOIN TB_고객등급 AS B ON A.고객등급코드 = B.고객등급코드"
    Query = Query & " WHERE A.고객코드 IS NOT NULL"
    
    If txtFind.Text = "" Then
        Query = Query & " ORDER BY A.성명 ASC"
    Else
        Select Case cboGubun.Text
            Case "성명":     Query = Query & " AND A.성명 LIKE '%" & txtFind.Text & "%'"
                             Query = Query & " ORDER BY A.성명 ASC"
            
            Case "전화번호": Query = Query & " AND (A.전화번호 LIKE '%" & txtFind.Text & "%'"
                             Query = Query & "  OR  A.휴대전화   LIKE '%" & txtFind.Text & "%')"
                             Query = Query & " ORDER BY A.전화번호, A.휴대전화 ASC"
                             
            Case "주소":     Query = Query & " AND A.주소 LIKE '%" & txtFind.Text & "%'"
                             Query = Query & " ORDER BY A.주소 ASC"
        End Select
    End If
    
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        
            .Col = 1:  .Text = ADORs!고객코드 & ""         ' 1
            .Col = 2:  .Text = ADORs!성명 & ""             ' 2
            .Col = 3:  .Text = ADORs!전화번호 & ""         ' 3
            .Col = 4:  .Text = ADORs!휴대전화 & ""         ' 4
            .Col = 5:  .Text = ADORs!주소 & ""             ' 5
            .Col = 6:  .Text = ADORs!SMS & ""              ' 6
            .Col = 7:  .Text = ADORs!등록일자 & ""         ' 7
            .Col = 8:  .Text = ADORs!고객등급명 & ""       ' 8
            .Col = 9:  .Text = ADORs!이용횟수 & ""         ' 9
            .Col = 10: .Text = ADORs!총접수금액 & ""       '10
            .Col = 11: .Text = ADORs!미수금액 & ""         '12
            .Col = 12: .Text = ADORs!누적마일리지 & ""     '11
            .Col = 13: .Text = ADORs!사용가능마일리지 & "" '13
            .Col = 14: .Text = ADORs!수정일자 & ""         '14
            
            ADORs.MoveNext
        Loop
        
        .ReDraw = True
    
        ADORs.Close
        Set ADORs = Nothing
        
        pnlHeader.Caption = "      고객현황 (" & .MaxRows & " 명)"
        
        If strCode <> "" Then
            Rtn = .SearchCol(1, -1, -1, strCode, SearchFlagsValue) '신규저장하거나 수정한 데이터 위치로 이동
            
            If Rtn > -1 Then
                Call .SetSelection(1, Rtn, .MaxCols, Rtn)
                
                Call sprGrid_Click(1, Rtn)
            End If
            
            strCode = ""
        End If
    End With

    pnlProg.Visible = False

    Exit Sub
    
ErrRtn:
    pnlProg.Visible = False
        
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
    
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyChk (KeyCode)
    
    If KeyCode = 13 Then
        SendKeys "{TAB}"
        KeyCode = 0
    End If
End Sub

Private Sub Form_Load()
    With sprGrid
        .MaxRows = 0
        .RowHeight(-1) = 14
        
        .ColsFrozen = 2
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeSingle
        
        '홀수/짝수 Row BankColor
        'Ret = .SetOddEvenRowColor(&HFFFFFF, &H80000008, &H80FFFF, &H80000008)

        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
    
    With sprMisu
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
        
        '홀수/짝수 Row BankColor
        'Ret = .SetOddEvenRowColor(&HFFFFFF, &H80000008, &H80FFFF, &H80000008)

        'Init the User Sort
        '.UserColAction = UserColActionSort
    End With
    
    With cboGubun
        .Clear
        .AddItem "성명"
        .AddItem "전화번호"
        .AddItem "주소"
        
        .ListIndex = 0
    End With
    
    Call 고객등급_Display(cboClass, False) '고객등급
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    cmdBtn(5).Left = Me.Width - cmdBtn(5).Width - 200
    btnUpdate.Left = Me.Width - btnUpdate.Width - 200
    
    sprMisu.Left = Me.Width - sprMisu.Width - 200
End Sub

Private Sub sprGrid_Click(ByVal Col As Long, ByVal Row As Long)
    On Error GoTo ErrRtn
    
    Dim 고객코드 As String
    
    If Row <= 0 Then Exit Sub
    
    sprGrid.Row = Row
    sprGrid.Col = 1: 고객코드 = sprGrid.Text & ""
    
    Query = "SELECT    A.고객코드"
    Query = Query & ", A.성명"
    Query = Query & ", A.전화번호"
    Query = Query & ", A.휴대전화"
    Query = Query & ", A.주소"
    Query = Query & ", A.미수금액"
    Query = Query & ", (CASE WHEN A.문자발송여부 = 'Y' THEN '1' ELSE '0' END) SMS"
    Query = Query & ", A.등록일자"
    Query = Query & ", A.고객등급코드"
    Query = Query & ", A.이용횟수"
    Query = Query & ", A.총접수금액"
    Query = Query & ", A.누적마일리지"
    Query = Query & ", A.사용가능마일리지"
    Query = Query & ", A.수정일자"
    Query = Query & ", A.카드번호"
    Query = Query & ", A.메모"
    Query = Query & ", B.고객등급명"
    Query = Query & " FROM TB_고객정보 AS A LEFT OUTER JOIN TB_고객등급 AS B ON A.고객등급코드 = B.고객등급코드"
    Query = Query & " WHERE A.고객코드 = '" & 고객코드 & "'"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    If ADORs.EOF Then
        '
    Else
        txtCode.Text = Trim(ADORs!고객코드) & ""      ' 1
        txtName.Text = Trim(ADORs!성명) & ""          ' 2
        txtTel.Text = Trim(ADORs!전화번호) & ""       ' 3
        txtHP.Text = Trim(ADORs!휴대전화) & ""        ' 4
        
        If ADORs!SMS = "1" Then
            cboSMS.ListIndex = 0                      ' 6
        Else
            cboSMS.ListIndex = 1                      ' 6
        End If
        
        txtAdd.Text = Trim(ADORs!주소) & ""           ' 5
        
        txtMoney.Value = ADORs!미수금액 & ""          ' 7
        txtMoney.tag = ADORs!미수금액 & ""            '
        
        txtCard.Text = Trim(ADORs!카드번호) & ""      ' 8
        txtMemo.Text = Trim(ADORs!메모) & ""          ' 9
        
        With cboClass                                 '10
            For i = 0 To .ListCount - 1
                If Left(.List(i), 1) = ADORs!고객등급코드 Then
                    .ListIndex = i
                    
                    Exit For
                End If
            Next i
        End With
    End If
    ADORs.Close
    Set ADORs = Nothing
    
    Call 미수금수정_Display
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
    
    Screen.MousePointer = 0
End Sub

Private Sub 미수금수정_Display()
    Query = "SELECT * FROM TB_미수금수정"
    Query = Query & " WHERE 고객코드 = '" & txtCode.Text & "'"
    Query = Query & " ORDER BY 수정일자 ASC"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprMisu
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        
            .Col = 1:  .Text = ADORs!수정일자 & ""   ' 1
            .Col = 2:  .Text = ADORs!수정미수금 & "" ' 2
            .Col = 3:  .Text = ADORs!이전미수금 & "" ' 3
            
            ADORs.MoveNext
        Loop
        
        .ReDraw = True
    
        ADORs.Close
        Set ADORs = Nothing
    End With

End Sub

Private Sub btnUpdate_Click()
    Dim 수정일자 As String
    
    If Trim(txtCode.Text) = "" Then Exit Sub
    
    Query = "SELECT * FROM TB_고객정보"
    Query = Query & " WHERE 고객코드 = '" & txtCode.Text & "'"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenDynamic, adLockOptimistic

    If Not ADORs.EOF Then
        'ADORs!고객코드 = Trim(txtCode.Text) & ""      ' 1
        ADORs!성명 = Trim(txtName.Text) & ""          ' 2
        ADORs!전화번호 = Trim(txtTel.Text) & ""       ' 3
        ADORs!휴대전화 = Trim(txtHP.Text) & ""        ' 4
        
        If cboSMS.ListIndex = 0 Then
            ADORs!문자발송여부 = "Y"                  ' 6
        Else
            ADORs!문자발송여부 = "N"                  ' 6
        End If
        
        ADORs!주소 = Trim(txtAdd.Text) & ""           ' 5
        ADORs!미수금액 = txtMoney.Value & ""          ' 7
        ADORs!카드번호 = Trim(txtCard.Text) & ""      ' 8
        ADORs!메모 = Trim(txtMemo.Text) & ""          ' 9
        ADORs!고객등급코드 = Left(cboClass.Text, 1)   '10
        
        ADORs!본사전송여부 = ""                       '
        
        ADORs.Update
    End If
    ADORs.Close
    Set ADORs = Nothing
    
    '----------------------------------------------------------------
    ' TB_미수금수정 - 미수금액을 수정한 경우
    '----------------------------------------------------------------
    If txtMoney.tag = "" Then
        '신규 고객
    Else
        If CStr(txtMoney.Value) <> CStr(txtMoney.tag) Then
            수정일자 = Format(Now, "YYYY-MM-DD hh:mm:ss")
            
            Query = "SELECT * FROM TB_미수금수정"
            Query = Query & " WHERE 고객코드 = '" & Trim(txtCode.Text) & "'"
            Query = Query & "   AND 수정일자 = '" & 수정일자 & "'"
            Set ADORs = New ADODB.RecordSet
            ADORs.Open Query, ADOCon, adOpenDynamic, adLockOptimistic
            
            If ADORs.EOF Then ADORs.AddNew
            
            ADORs!지사코드 = 가맹점정보.지사코드 & ""     ' 1
            ADORs!가맹점코드 = 가맹점정보.가맹점코드 & "" ' 2
            ADORs!고객코드 = Trim(txtCode.Text) & ""      ' 3
            ADORs!수정일자 = 수정일자 & ""                ' 4
            ADORs!수정미수금 = txtMoney.Value             ' 5
            ADORs!이전미수금 = txtMoney.tag & ""          ' 6
            ADORs!내용 = "조정 - 고객현황"                ' 6
            ADORs.Update
            
            ADORs.Close
            Set ADORs = Nothing
        End If
    End If
    
    strCode = Trim(txtCode.Text) & ""

    cmdList_Click
End Sub

Private Sub sprGrid_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Call sprGrid_Click(NewCol, NewRow)
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        cmdList_Click
    End If
End Sub
