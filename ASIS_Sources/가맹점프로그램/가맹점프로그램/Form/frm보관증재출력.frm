VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm보관증재출력 
   Caption         =   "보관증 재출력"
   ClientHeight    =   11250
   ClientLeft      =   6810
   ClientTop       =   3315
   ClientWidth     =   15915
   ControlBox      =   0   'False
   LinkTopic       =   "Form25"
   MDIChild        =   -1  'True
   ScaleHeight     =   11250
   ScaleWidth      =   15915
   WindowState     =   2  '최대화
   Begin Threed.SSPanel pnlProg 
      Height          =   1215
      Left            =   45
      TabIndex        =   12
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
      Picture         =   "frm보관증재출력.frx":0000
      Caption         =   " "
      BorderWidth     =   0
      BevelOuter      =   1
      BevelInner      =   2
      Alignment       =   4
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   11250
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15915
      _ExtentX        =   28072
      _ExtentY        =   19844
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm보관증재출력.frx":2FCB
      Begin Threed.SSPanel SSPanel 
         Height          =   3000
         Index           =   0
         Left            =   15
         TabIndex        =   14
         Top             =   8235
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   5292
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin CSTextLibCtl.silgEdit txtMoney 
            Height          =   330
            Index           =   0
            Left            =   900
            TabIndex        =   27
            Top             =   1845
            Width           =   945
            _Version        =   262145
            _ExtentX        =   1667
            _ExtentY        =   582
            _StockProps     =   125
            Text            =   " 0"
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
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "0"
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
            Undo            =   1
            Data            =   0
         End
         Begin Threed.SSPanel pnlData 
            Height          =   300
            Index           =   0
            Left            =   900
            TabIndex        =   15
            Top             =   45
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   529
            _Version        =   262144
            BackColor       =   16777215
            BevelOuter      =   1
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlData 
            Height          =   300
            Index           =   1
            Left            =   900
            TabIndex        =   17
            Top             =   375
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   529
            _Version        =   262144
            BackColor       =   16777215
            BevelOuter      =   1
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlData 
            Height          =   300
            Index           =   2
            Left            =   900
            TabIndex        =   18
            Top             =   705
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   529
            _Version        =   262144
            BackColor       =   16777215
            BevelOuter      =   1
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlData 
            Height          =   300
            Index           =   3
            Left            =   900
            TabIndex        =   19
            Top             =   1035
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   529
            _Version        =   262144
            BackColor       =   16777215
            BevelOuter      =   1
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlData 
            Height          =   300
            Index           =   4
            Left            =   900
            TabIndex        =   23
            Top             =   1380
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   529
            _Version        =   262144
            BackColor       =   16777215
            BevelOuter      =   1
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlData 
            Height          =   300
            Index           =   5
            Left            =   900
            TabIndex        =   25
            Top             =   1500
            Visible         =   0   'False
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   529
            _Version        =   262144
            BackColor       =   16777215
            BevelOuter      =   1
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.silgEdit txtMoney 
            Height          =   330
            Index           =   1
            Left            =   3150
            TabIndex        =   28
            Top             =   1845
            Width           =   945
            _Version        =   262145
            _ExtentX        =   1667
            _ExtentY        =   582
            _StockProps     =   125
            Text            =   " 0"
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
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "0"
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
            Undo            =   1
            Data            =   0
         End
         Begin CSTextLibCtl.silgEdit txtMoney 
            Height          =   330
            Index           =   2
            Left            =   930
            TabIndex        =   29
            Top             =   2205
            Width           =   945
            _Version        =   262145
            _ExtentX        =   1667
            _ExtentY        =   582
            _StockProps     =   125
            Text            =   " 0"
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
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "0"
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
            Undo            =   1
            Data            =   0
         End
         Begin CSTextLibCtl.silgEdit txtMoney 
            Height          =   330
            Index           =   3
            Left            =   3150
            TabIndex        =   30
            Top             =   2205
            Width           =   945
            _Version        =   262145
            _ExtentX        =   1667
            _ExtentY        =   582
            _StockProps     =   125
            Text            =   " 0"
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
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "0"
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
            Undo            =   1
            Data            =   0
         End
         Begin CSTextLibCtl.silgEdit txtMoney 
            Height          =   330
            Index           =   4
            Left            =   930
            TabIndex        =   35
            Top             =   2550
            Width           =   945
            _Version        =   262145
            _ExtentX        =   1667
            _ExtentY        =   582
            _StockProps     =   125
            Text            =   " 0"
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
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "0"
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
            Undo            =   1
            Data            =   0
         End
         Begin CSTextLibCtl.silgEdit txtMoney 
            Height          =   330
            Index           =   5
            Left            =   3150
            TabIndex        =   36
            Top             =   2550
            Width           =   945
            _Version        =   262145
            _ExtentX        =   1667
            _ExtentY        =   582
            _StockProps     =   125
            Text            =   " 0"
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
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "0"
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
            Undo            =   1
            Data            =   0
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "쿠폰결제"
            Height          =   225
            Index           =   11
            Left            =   2340
            TabIndex        =   38
            Top             =   2610
            Width           =   765
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "사용마일"
            Height          =   225
            Index           =   10
            Left            =   120
            TabIndex        =   37
            Top             =   2610
            Width           =   765
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "현금결제"
            Height          =   225
            Index           =   9
            Left            =   120
            TabIndex        =   34
            Top             =   2265
            Width           =   765
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "카드결제"
            Height          =   225
            Index           =   8
            Left            =   2340
            TabIndex        =   33
            Top             =   2265
            Width           =   765
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "접수금액"
            Height          =   225
            Index           =   7
            Left            =   60
            TabIndex        =   32
            Top             =   1905
            Width           =   765
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "이전미수"
            Height          =   225
            Index           =   6
            Left            =   2310
            TabIndex        =   31
            Top             =   1905
            Width           =   765
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "예정일자"
            Height          =   225
            Index           =   5
            Left            =   60
            TabIndex        =   26
            Top             =   1545
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "접수일자"
            Height          =   225
            Index           =   4
            Left            =   60
            TabIndex        =   24
            Top             =   1455
            Width           =   765
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "주소"
            Height          =   225
            Index           =   3
            Left            =   60
            TabIndex        =   22
            Top             =   1110
            Width           =   765
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "휴대전화"
            Height          =   225
            Index           =   2
            Left            =   60
            TabIndex        =   21
            Top             =   780
            Width           =   765
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "전화번호"
            Height          =   225
            Index           =   1
            Left            =   60
            TabIndex        =   20
            Top             =   450
            Width           =   765
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "고객명"
            Height          =   225
            Index           =   0
            Left            =   60
            TabIndex        =   16
            Top             =   120
            Width           =   765
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   750
         Left            =   15
         TabIndex        =   2
         Top             =   450
         Width           =   15885
         _ExtentX        =   28019
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   0
            Left            =   4260
            TabIndex        =   4
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 조회(&F)"
            Appearance      =   6
            Picture         =   "frm보관증재출력.frx":309D
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   4
            Left            =   9330
            TabIndex        =   5
            Top             =   60
            Width           =   1665
            _Version        =   851970
            _ExtentX        =   2937
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 영수증 출력"
            Appearance      =   6
            Picture         =   "frm보관증재출력.frx":3797
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   11085
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
            Picture         =   "frm보관증재출력.frx":3E91
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   0
            Left            =   915
            TabIndex        =   0
            Top             =   60
            Width           =   1425
            _ExtentX        =   2514
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
            Format          =   56557571
            CurrentDate     =   40279
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   1
            Left            =   2610
            TabIndex        =   9
            Top             =   60
            Width           =   1425
            _ExtentX        =   2514
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
            Format          =   56557571
            CurrentDate     =   40279
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
            TabIndex        =   11
            Top             =   120
            Width           =   840
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
            Index           =   0
            Left            =   2400
            TabIndex        =   10
            Top             =   120
            Width           =   120
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   3
         Top             =   15
         Width           =   15885
         _ExtentX        =   28019
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
         Caption         =   "      보관증 재출력"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm보관증재출력.frx":4F23
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm보관증재출력.frx":5149
            Top             =   -15
            Width           =   765
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   390
         Left            =   15
         TabIndex        =   7
         Top             =   7830
         Width           =   15885
         _ExtentX        =   28019
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
         Caption         =   " 접수 상세내역"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm보관증재출력.frx":5D13
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin CSTextLibCtl.silgEdit txtMoney 
            Height          =   330
            Index           =   6
            Left            =   4800
            TabIndex        =   39
            Top             =   0
            Width           =   945
            _Version        =   262145
            _ExtentX        =   1667
            _ExtentY        =   582
            _StockProps     =   125
            Text            =   " 0"
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
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "0"
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
            Undo            =   1
            Data            =   0
         End
         Begin CSTextLibCtl.silgEdit txtMoney 
            Height          =   330
            Index           =   7
            Left            =   7020
            TabIndex        =   40
            Top             =   0
            Width           =   945
            _Version        =   262145
            _ExtentX        =   1667
            _ExtentY        =   582
            _StockProps     =   125
            Text            =   " 0"
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
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "0"
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
            Undo            =   1
            Data            =   0
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "접수정상금액"
            Height          =   225
            Index           =   13
            Left            =   3360
            TabIndex        =   42
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "할인금액"
            Height          =   225
            Index           =   12
            Left            =   6210
            TabIndex        =   41
            Top             =   60
            Width           =   765
         End
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   6600
         Left            =   15
         TabIndex        =   8
         Top             =   1215
         Width           =   15885
         _Version        =   524288
         _ExtentX        =   28019
         _ExtentY        =   11642
         _StockProps     =   64
         BackColorStyle  =   1
         ColsFrozen      =   2
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
         MaxCols         =   17
         Protect         =   0   'False
         ShadowColor     =   14737632
         SpreadDesigner  =   "frm보관증재출력.frx":5F35
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread sprList 
         Height          =   3000
         Left            =   4185
         TabIndex        =   13
         Top             =   8235
         Width           =   11715
         _Version        =   524288
         _ExtentX        =   20664
         _ExtentY        =   5292
         _StockProps     =   64
         BackColorStyle  =   1
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         EditModeReplace =   -1  'True
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
         MaxCols         =   20
         MaxRows         =   200
         ScrollBars      =   2
         SpreadDesigner  =   "frm보관증재출력.frx":69D5
         UserResize      =   1
         VisibleCols     =   7
         VisibleRows     =   30
         HighlightHeaders=   1
         HighlightStyle  =   1
         ScrollBarStyle  =   2
      End
   End
End
Attribute VB_Name = "frm보관증재출력"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 0: Call Data_Display
        Case 4: Call Data_Print
        Case 5
            Unload Me
    End Select
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    pnlData(0).Caption = ""
    pnlData(1).Caption = ""
    pnlData(2).Caption = ""
    pnlData(3).Caption = ""
    pnlData(4).Caption = ""
    pnlData(5).Caption = ""

    txtMoney(0).Value = 0
    txtMoney(1).Value = 0
    txtMoney(2).Value = 0
    txtMoney(3).Value = 0

    sprList.MaxRows = 0
    
    pnlProg.Visible = True
    DoEvents
    
    '-------------------------------------------------------------------------------------------
    '
    '-------------------------------------------------------------------------------------------
    Query = "SELECT    A.*"
    Query = Query & ", B.성명"
    Query = Query & ", B.전화번호"
    Query = Query & ", B.주소"
'    Query = Query & ", c.판매취소금액, c.판매취소수량"
    Query = Query & " FROM TB_매출 AS A "
    Query = Query & "   LEFT OUTER JOIN TB_고객정보 AS B ON A.고객코드 = B.고객코드"
'    Query = Query & "   LEFT OUTER JOIN  (select 고객코드, 접수번호, SUM(접수금액)as 판매취소금액, SUM(반품수량) as 판매취소수량 from TB_매출 "
'    Query = Query & "                       where 매출일자 between '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "' AND '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "'"
'    Query = Query & "                       group by 고객코드, 접수번호 ) C"
'    Query = Query & "   on c.고객코드 = A.고객코드 and c.접수번호 = A.접수번호"
    Query = Query & " WHERE (A.매출일자 >= '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "' "
    Query = Query & "   AND  A.매출일자 <= '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "') "
    Query = Query & "   AND A.접수수량 > 0"
    
    'If txtFind.Tag <> "" Then
    '    Query = Query & " AND A.고객코드 = " & txtFind.Tag
    'End If
    
    Query = Query & " ORDER BY A.매출일자 DESC, A.매출시간 DESC"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = Format(ADORs!매출일자, "YYYY-MM-DD") & ""
            .Col = 2:  .Text = ADORs!성명 & ""
            .Col = 3:  .Text = ADORs!전화번호 & ""
            .Col = 4:  .Text = ADORs!주소 & ""
            .Col = 5:  .Text = ADORs!접수번호 & ""
            .Col = 6:  .Text = ADORs!적요 & ""
            
            If ADORs!반품수량 <> 0 Then
                .Col = 7:  .Text = ""
                .Col = 8:  .Text = ADORs!접수금액
                
                '//
                .Row = .MaxRows: .Row2 = .MaxRows
                .Col = 6: .Col2 = .MaxCols
                .BlockMode = True
                .ForeColor = vbRed
                .BlockMode = False
            End If
                        
            '----------------------------------------------------------------------------------
            ' 반품수량
            '----------------------------------------------------------------------------------
            Query = "SELECT    ISNULL(SUM(반품수량),0) AS 반품수량2"
            Query = Query & ", ISNULL(SUM(접수금액),0) AS 접수금액2"
            Query = Query & ", ISNULL(SUM(현금입금),0) AS 현금입금2"
            Query = Query & ", ISNULL(SUM(카드입금),0) AS 카드입금2"
            Query = Query & " FROM TB_매출"
            Query = Query & " WHERE 고객코드 = '" & ADORs!고객코드 & "'"
            Query = Query & "   AND 접수번호 = '" & ADORs!접수번호 & "'"
            Query = Query & "   AND 반품수량 < 0"
            Set SUBRs = New ADODB.RecordSet
            SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
                        
            If SUBRs.EOF Then
                .Col = 7:  .Text = ADORs!접수수량 & "" '- Val(ADORs!판매취소수량 & "")
                .Col = 8:  .Text = ADORs!접수금액 & "" '- Val(ADORs!판매취소금액 & "")
                .Col = 9:  .Text = ADORs!현금입금 & ""  '
                .Col = 10: .Text = ADORs!카드입금 & ""  '
                .Col = 11: .Text = ADORs!사용마일리지 & "" '
                .Col = 12: .Text = ADORs!쿠폰입금 & "" '
                .Col = 14: .Text = ADORs!반품수량 & "" '
                .Col = 15: .Text = ADORs!고객코드 & "" '
                .Col = 16: .Text = ADORs!일련번호 & "" '
                
                .Col = 17: .Text = ADORs!이전미수금 & "" '
            Else
                .Col = 7:  .Text = ADORs!접수수량 - (SUBRs!반품수량2 * -1) & "" '
                .Col = 8:  .Text = ADORs!접수금액 - (SUBRs!접수금액2 * -1) & "" '
                .Col = 9:  .Text = ADORs!현금입금 - (SUBRs!현금입금2 * -1) & "" '
                .Col = 10: .Text = ADORs!카드입금 - (SUBRs!카드입금2 * -1) & "" '
                .Col = 11: .Text = ADORs!사용마일리지 & "" '
                .Col = 12: .Text = ADORs!쿠폰입금 & "" '
                .Col = 14: .Text = ADORs!반품수량 & "" '
                .Col = 15: .Text = ADORs!고객코드 & "" '
                .Col = 16: .Text = ADORs!일련번호 & "" '
                
                .Col = 17: .Text = ADORs!이전미수금 & "" '
            End If
            SUBRs.Close
            Set SUBRs = Nothing
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        .ReDraw = True
    End With
    
    pnlProg.Visible = False
    
    Exit Sub
    
ErrRtn:
    pnlProg.Visible = False
    
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub Data_Print()
    On Error GoTo ErrRtn
    
    Dim ESC      As String * 1
    Dim CommPort As String
    Dim BaudRate As String
    Dim sE       As String
    Dim Print_Msg As String
    
    Dim tmp      As String
    Dim 이전미수 As String
    Dim 접수수량 As Integer
    Dim 접수금액 As String
    Dim 정상금액 As Double
    
    Dim 현금결제 As String
    Dim 카드결제 As String
    Dim 마일리지결제 As String
    Dim 쿠폰결제 As String
    
    Dim 운동화세탁안내 As Boolean
    
    
    Dim 카드번호 As String
    
    Dim 전화번호     As String
    Dim 전화번호출력 As String
    
    운동화세탁안내 = False

    ESC = Chr(&H1B)
    전화번호출력 = GetIniStr("Printer", "TelPrint", "Y", iniFile)
    
    If 가맹점정보.지사코드 = M_COUPON_KLENZ_CODE Then '크렌즈갤러리
        Print_Msg = Print_Msg & PrintTitle2("크렌즈갤러리 - 세탁물 접수증(재출력)")
    Else
        Print_Msg = Print_Msg & PrintTitle2("크린에이드 - 세탁물 접수증(재출력)")
    End If
    
    
    Print_Msg = Print_Msg & PrintLineFeed
    
    
    '------------------------------------------------------------
    '
    '------------------------------------------------------------
    Query = "SELECT    ISNULL(가맹점명, '')     AS 가맹점명"
    Query = Query & ", ISNULL(매장전화번호, '') AS 매장전화번호"
    Query = Query & ", ISNULL(사업장주소, '')   AS 사업장주소"
    Query = Query & ", ISNULL(사업자번호, '')   AS 사업자번호"
    Query = Query & ", ISNULL(대표자명, '')     AS 대표자명"
    Query = Query & " FROM TB_기본정보"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF Then
        Print_Msg = Print_Msg & PrintString("상 호 명 : ", 1, True)
        Print_Msg = Print_Msg & PrintString("전화번호 : ", 1, True)
        Print_Msg = Print_Msg & PrintString("주    소 : ", 1, True)
    Else
        Print_Msg = Print_Msg & PrintString("상 호 명 : " + ADORs!가맹점명, 1, True)
        Print_Msg = Print_Msg & PrintString("사업자No : " + ADORs!사업자번호, 1, True)
        Print_Msg = Print_Msg & PrintString("대 표 자 : " + ADORs!대표자명, 1, True)
        Print_Msg = Print_Msg & PrintString("전화번호 : " + ADORs!매장전화번호, 1, True)
        Print_Msg = Print_Msg & PrintString("주    소 : " + ADORs!사업장주소, 1, True)
    End If
    ADORs.Close
    Set ADORs = Nothing
    
    Print_Msg = Print_Msg & PrintString("==============================================", 1, True)
    Print_Msg = Print_Msg & PrintString("접수일자 : " + Format(pnlData(4).Caption, "YYYY년 MM월 DD일 AM/PM hh:mm"), 1, True)
    Print_Msg = Print_Msg & PrintString("찾을날짜 : " + Format(pnlData(5).Caption, "YYYY년 MM월 DD일"), 1, True)
    Print_Msg = Print_Msg & PrintString("고객코드 : " + pnlData(0).Tag, 1, True)
    
    Print_Msg = Print_Msg & PrintCustomer(전화번호출력, pnlData(0).Caption, pnlData(1).Caption, pnlData(2).Caption, pnlData(3).Caption)
    
    Print_Msg = Print_Msg & PrintString("==============================================", 1, True)
    Print_Msg = Print_Msg & PrintString("택번호  의류/상표         작업   색상     금액", 1, True)
    Print_Msg = Print_Msg & PrintString("----------------------------------------------", 1, True)
    
    접수수량 = 0
    정상금액 = 0
    
    With sprList
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
            Print_Msg = Print_Msg & ESC + "!" + Chr$(8)              'Selects Emphasized mode
            Print_Msg = Print_Msg & .Text + " "
            Print_Msg = Print_Msg & ESC + "!" + Chr$(0)              'Cancels Emphasized mode
        
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
            Print_Msg = Print_Msg & PrintString(tmp, 1) + ""
            
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
            Print_Msg = Print_Msg & PrintString(tmp, 1) + " "
            
            '*********************************************************
            '* 색상
            '*********************************************************
            .Col = 3
            If LenH(.Text) >= 4 Then
                tmp = MidH(.Text, 1, 4)
            Else
                tmp = Trim(.Text) + String(4 - LenH(.Text), " ")
            End If
            
            Print_Msg = Print_Msg & PrintString(tmp, 1) + " "

            '*********************************************************
            '* 금액
            '*********************************************************
            .Col = 20
            
            If Len(.Text) > 8 Then
                Print_Msg = Print_Msg & PrintString(.Text, 1, True)
            Else
                Print_Msg = Print_Msg & PrintString(String(8 - LenH(.Text), " ") + .Text, 1, True)
            End If
            .Col = 6
            TempMoney = Replace(.Text, ",", "")
            '*********************************************************
            '* 상표
            '*********************************************************
            .Col = 7
            
            If Trim(.Text) <> "" Then
                Print_Msg = Print_Msg & PrintString("        - " + .Text, 1, True)
            End If

            '*********************************************************
            '* 오점
            '*********************************************************
            .Col = 19
            
            If Trim(.Text) <> "" Then
                Print_Msg = Print_Msg & PrintString("        - " + .Text, 1, True)
            End If
            
            ' 정상금액
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
                Print_Msg = Print_Msg & PrintString("        * 할인금액 " + String(19, " ") + Calc, 1, True)
                'Print_Msg = Print_Msg & PrintString("        * 정상금액 :" + CStr(TempMoney) + "/ 할인금액 :" + Calc, 1, True)
            End If
            
            
            '*********************************************************
            '* 운동화세탁안내
            '*********************************************************
            .Col = 8
            
            If Left(Trim(.Text), 2) = "a0" Then 운동화세탁안내 = True
        
        Next i
    End With
    
    접수금액 = Trim(txtMoney(0).Text)
    이전미수 = Trim(txtMoney(1).Text)
    현금결제 = Trim(txtMoney(2).Text)
    카드결제 = Trim(txtMoney(3).Text)
    마일리지결제 = Trim(txtMoney(4).Text)
    쿠폰결제 = Trim(txtMoney(5).Text)
    
    Print_Msg = Print_Msg & PrintString("----------------------------------------------", 1, True)
    If CDbl(Replace(Trim(txtMoney(7).Text), ",", "")) > 0 Then
        Print_Msg = Print_Msg & PrintString("정상금액 :" + Format(Trim(txtMoney(6).Text), "@@@@@@@@@@") + "원", 1)
        Print_Msg = Print_Msg & PrintString("/ 할인금액 :" + Format(Trim(txtMoney(7).Text), "@@@@@@@@@@") + "원", 1, True)
        Print_Msg = Print_Msg & PrintString("----------------------------------------------", 1, True)
    End If
    Print_Msg = Print_Msg & PrintString(String(24, " ") + "이전미수 :" + Format(이전미수, "@@@@@@@@@@") & "원", 1, True)
    Print_Msg = Print_Msg & PrintString("접수수량 :" + Format(접수수량, "@@@@@@@@@@") + "점/ 접수금액 :" + Format(접수금액, "@@@@@@@@@@") + "원", 1, True)
    
    If Trim(쿠폰결제) <> "0" Then Print_Msg = Print_Msg & PrintString(String(24, " ") + "쿠폰결제 : " + Format(쿠폰결제, "@@@@@@@@@@") + "원", 1, True)
    If Trim(마일리지결제) <> "0" Then Print_Msg = Print_Msg & PrintString(String(24, " ") + "마일리지 :" + Format(마일리지결제, "@@@@@@@@@@") + "원", 1, True)
    
    Print_Msg = Print_Msg & PrintString(String(24, " ") + "현금결제 :" + Format(현금결제, "@@@@@@@@@@") + "원", 1, True)
    Print_Msg = Print_Msg & PrintString(String(24, " ") + "카드결제 :" + Format(카드결제, "@@@@@@@@@@") + "원", 1, True)
    
    Print_Msg = Print_Msg & PrintString("==============================================", 1, True)
    Print_Msg = Print_Msg & PrintLineFeed
    
    Print_Msg = Print_Msg & PrintString("※ 인도예정일은 세탁물의 오염정도에 따라 다소", 1, True)
    Print_Msg = Print_Msg & PrintString("   지연될 수 있습니다.", 1, True)
    Print_Msg = Print_Msg & PrintLineFeed
    

    If 운동화세탁안내 Then Print_Msg = Print_Msg & 운동화세탁안내_Report
        
    
    Print_Msg = Print_Msg & PrintLineFeed(4)
    
    Print_Msg = Print_Msg & PrintCut
    
    Call frmKicc.Card_Print(Print_Msg)
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
    
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    With sprGrid
        .MaxRows = 0
        .RowHeight(-1) = 13
        
        .Col = 15: .ColHidden = True '고객코드
        .Col = 16: .ColHidden = True '일련번호
        .Col = 17: .ColHidden = True '이전미수금
                
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

    With sprList
        .MaxRows = 0
        .RowHeight(-1) = 13
        
        .Col = 8:  .ColHidden = True '
        .Col = 9:  .ColHidden = True '
        .Col = 10: .ColHidden = True '
        .Col = 11: .ColHidden = True '
        .Col = 12: .ColHidden = True '
        .Col = 13: .ColHidden = True '
        .Col = 14: .ColHidden = True '

        .Col = 16: .ColHidden = True '
        .Col = 17: .ColHidden = True '
        .Col = 18: .ColHidden = True '
        
        .Col = 20: .ColHidden = True '
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeSingle
    End With
    
    dtpDay(0).Value = Date
    dtpDay(1).Value = Date
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    cmdBtn(5).Left = Me.Width - cmdBtn(5).Width - 200
End Sub


Private Sub sprGrid_Click(ByVal Col As Long, ByVal Row As Long)
    Dim 접수일자 As String
    Dim 접수번호 As String
    Dim 고객코드 As String
    
'    pnlData(0).Caption = ""
'    pnlData(1).Caption = ""
'    pnlData(2).Caption = ""
'    pnlData(3).Caption = ""
'    pnlData(4).Caption = ""
'    pnlData(5).Caption = ""
'
'    txtMoney(0).Value = 0
'    txtMoney(1).Value = 0
'    txtMoney(2).Value = 0
'    txtMoney(3).Value = 0
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    sprGrid.Row = Row
    sprGrid.Col = 1:  접수일자 = sprGrid.Text & ""
    sprGrid.Col = 5:  접수번호 = sprGrid.Text & ""
    sprGrid.Col = 15: 고객코드 = sprGrid.Text & ""
    
    sprGrid.Col = 8:  txtMoney(0).Value = sprGrid.Value & ""
    sprGrid.Col = 17: txtMoney(1).Value = sprGrid.Value & ""
    sprGrid.Col = 9:  txtMoney(2).Value = sprGrid.Value & ""
    sprGrid.Col = 10: txtMoney(3).Value = sprGrid.Value & ""
    
    sprGrid.Col = 11: txtMoney(4).Value = sprGrid.Value & ""
    sprGrid.Col = 12: txtMoney(5).Value = sprGrid.Value & ""
    
    '---------------------------------------------------------
    '
    '---------------------------------------------------------
    Query = "SELECT * FROM TB_고객정보"
    Query = Query & " WHERE 고객코드 = '" & 고객코드 & "'"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF Then
        pnlData(0).Caption = ""
        pnlData(1).Caption = ""
        pnlData(2).Caption = ""
        pnlData(3).Caption = ""
    Else
        pnlData(0).Caption = Trim(ADORs!성명) & "" '
        pnlData(0).Tag = ADORs!고객코드 & ""       '
        pnlData(1).Caption = ADORs!전화번호 & ""   '
        pnlData(2).Caption = ADORs!휴대전화 & ""   '
        pnlData(3).Caption = ADORs!주소 & ""       '
    End If
    ADORs.Close
    Set ADORs = Nothing
    
    '---------------------------------------------------------
    '
    '---------------------------------------------------------
    Query = "SELECT * FROM TB_입출고"
    Query = Query & " WHERE 접수일자 = '" & 접수일자 & "'"
    Query = Query & "   AND 접수번호 =  " & 접수번호
    Query = Query & "   AND 고객코드 = '" & 고객코드 & "'"
    Query = Query & "   AND (판매취소 <> 'Y')"
    'Query = Query & "   AND (판매취소 = '') AND (판매취소일자 = '' OR 판매취소일자 IS NULL)"
    Query = Query & "   AND (반품환불일자 = '' OR 반품환불일자 IS NULL)"
    Query = Query & "   AND (세탁환불일자 = '' OR 세탁환불일자 IS NULL)"
    Query = Query & " ORDER BY 택번호 ASC"
    
    Call 접수현황_Display(Query)
End Sub

Private Sub sprGrid_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Call sprGrid_Click(NewCol, NewRow)
End Sub

Private Sub 접수현황_Display(SQL As String)
    Dim nRow    As Long
    Dim dPrice            As Double
    Dim dOrgPrice         As Double
    Dim dDiscountTotal    As Double
    Dim dOrgPriceTotal    As Double
    
    On Error GoTo ErrRtn
    
    Set ADORs = New ADODB.RecordSet
    ADORs.Open SQL, ADOCon, adOpenForwardOnly, adLockReadOnly

    With sprList
        .MaxRows = 0
        .ReDraw = False
        
        If Not ADORs.EOF Then
            pnlData(4).Caption = ADORs!접수일자 & " " & ADORs!접수시간
            pnlData(5).Caption = ADORs!예정일자 & ""
        End If
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = ADORs!의류명 & ""
            .Col = 2:  .Text = Mid(ADORs!택번호, 4, 2) & "-" & Right(ADORs!택번호, 4)
            .Col = 3:  .Text = ADORs!색상 & ""
            .Col = 4:  .Text = ADORs!무늬 & ""
            .Col = 5:  .Text = ADORs!내용 & ""
            .Col = 6:  .Text = ADORs!금액 & ""
            .Col = 7:  .Text = ADORs!상표 & ""
            
            .Col = 8:  .Text = ADORs!의류코드 & ""
            .Col = 9:  .Text = ADORs!수선금액 & ""
            .Col = 10: .Text = ADORs!세트Key & ""
            .Col = 11: .Text = ADORs!세트구분 & ""
            .Col = 12: .Text = ADORs!세트금액1 & ""
            .Col = 13: .Text = ADORs!세트금액2 & ""
            .Col = 14: .Text = ADORs!정상금액 & ""
            .Col = 15: .Text = Mid(ADORs!부모택번호, 4, 2) & "-" & Right(ADORs!부모택번호, 4) & ""
            .Col = 16: .Text = ADORs!세탁마진 & ""
            .Col = 17: .Text = ADORs!외주마진 & ""
            .Col = 18: .Text = ADORs!수선마진 & ""
            .Col = 19: .Text = ADORs!오점내용 & ""
            .Col = 20: .Text = ADORs!의류금액 & ""
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        .ReDraw = True
    End With
    
    


' 할인 금액을 구한다.
    DoEvents

    dDiscountTotal = 0
    dOrgPriceTotal = 0
    
    With sprList
        For nRow = 1 To .MaxRows
            dPrice = 0
            dOrgPrice = 0
            
            ' 현재 수령 금액을 얻어 온다.
            .Row = nRow:    .Col = 6
            If Trim(.Value) <> "" Then dPrice = CDbl(Replace(.Value, ",", ""))
    
            ' 원 금액 금액
            .Row = nRow:    .Col = 20
            If Trim(sprGrid.Value) = "" Then Exit For
            If Trim(.Value) <> "" Then
                dOrgPrice = CDbl(Replace(.Value, ",", ""))
            End If
            
            ' 할증이 있을 경우 할증 금액을 정상금액으로 처리한다.
            If dPrice >= dOrgPrice Then
                dOrgPriceTotal = dOrgPriceTotal + dPrice
                
            ' 할증이 없고 할인이 있을 경우 정상금액을 기준으로 한다.
            Else
                dOrgPriceTotal = dOrgPriceTotal + dOrgPrice
            
            End If
            
            ' 할인된 금액만을 구한다.
            ' 할증처리된 금액이 차감되는 문제 처리
            If dPrice < dOrgPrice Then dDiscountTotal = dDiscountTotal + (dOrgPrice - dPrice)
        Next nRow
    End With
    
    txtMoney(6).Value = dOrgPriceTotal
    txtMoney(7).Value = dDiscountTotal
     
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub


Private Sub SSPanel1_Click()
'    Dim ESC      As String * 1
'    Dim CommPort As String
'    Dim BaudRate As String
'    Dim GetByte_Buf()   As Byte
'    Dim FileNum     As Integer
'    Dim FileSize    As Double
'    Dim i, j As Integer
'
'    ESC = Chr(&H1B)
'
'    FileNum = FreeFile()
'    Open App.Path & "\Logo.bmp" For Binary Access Read As FileNum
'    FileSize = LOF(FileNum)           ' 파일 사이즈를 구한다.
'
'    CommPort = GetIniStr("VAN", "KS7500_CommPort", "", iniFile)
'    BaudRate = GetIniStr("VAN", "KS7500_BaudRate", "", iniFile)
'
'    ReDim GetByte_Buf(FileSize)
'    Get FileNum, , GetByte_Buf()
'    Close FileNum
'
'    With frmMain.MSComm
'        If .PortOpen = True Then
'            .PortOpen = False
'        End If
'
'        .CommPort = CommPort
'        .InputLen = 0
'        .PortOpen = True
'        .Settings = BaudRate & ",n,8,1"
'
'        .Output = ESC + "!" + Chr$(0)              'Specifies font A (ESC !)
'        .Output = ESC + "a" + Chr$(1)              'Specifies a centered printing position (ESC a)
''        .Output = ESC + "!" + Chr$(16)             'Selects double-height mode
'
'        .Output = "이미지 출력 테스트 시작" + Chr$(&HA)
'        '        .Output = ESC + "*" + Chr$(24) + Chr$(0) + Chr$(3) + GetByte_Buf 'ESC * m nL nH [d1...dk]
'
'        For i = 0 To UBound(GetByte_Buf)
'            .Output = Chr$(GetByte_Buf(i))
'        Next i
'
'        .Output = "이미지 출력 테스트 종료" + Chr$(&HA)
'        .Output = Chr$(&H1D) + "V" + Chr$(66) + Chr$(0) 'Feeds paper & cut
'
'        .PortOpen = False
'    End With
'    Exit Sub

End Sub

Private Function 운동화세탁안내_Report() As String
    Dim bReport As Boolean
    Dim vText   As Variant
    Dim nRow    As Long
    Dim ReturnMsg As String
    
    bReport = False

    On Error GoTo ErrRtn
    
 
    
    With sprList
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

