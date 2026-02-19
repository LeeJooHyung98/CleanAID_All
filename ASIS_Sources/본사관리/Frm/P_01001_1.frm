VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_01001_1 
   Caption         =   "지사별 가맹점 등록"
   ClientHeight    =   11010
   ClientLeft      =   105
   ClientTop       =   3045
   ClientWidth     =   15240
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   11010
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   19420
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_01001_1.frx":0000
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   9660
         Left            =   15
         TabIndex        =   1
         Top             =   1335
         Width           =   3435
         _Version        =   524288
         _ExtentX        =   6059
         _ExtentY        =   17039
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   2
         SpreadDesigner  =   "P_01001_1.frx":00B2
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   2
         Top             =   540
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   3
            Left            =   1530
            Style           =   2  '드롭다운 목록
            TabIndex        =   4
            Top             =   60
            Width           =   2775
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   0
            Left            =   6090
            TabIndex        =   3
            Top             =   60
            Width           =   2400
         End
         Begin Threed.SSOption optGubun 
            Height          =   285
            Index           =   0
            Left            =   6120
            TabIndex        =   5
            Top             =   435
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   503
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
            Caption         =   "전체"
            Value           =   -1
         End
         Begin Threed.SSOption optGubun 
            Height          =   285
            Index           =   1
            Left            =   6960
            TabIndex        =   6
            Top             =   435
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
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
            Caption         =   "정상"
         End
         Begin Threed.SSOption optGubun 
            Height          =   285
            Index           =   2
            Left            =   7860
            TabIndex        =   7
            Top             =   435
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   503
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
            Caption         =   "할인"
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   4620
            TabIndex        =   8
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
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
            Caption         =   "가맹점코드"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   25
            Left            =   4620
            TabIndex        =   9
            Top             =   420
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
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
            Caption         =   "가맹점구분"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   4
            Left            =   60
            TabIndex        =   10
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
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
            Caption         =   "지 사 명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel panDetail 
         Height          =   9660
         Left            =   3465
         TabIndex        =   11
         Top             =   1335
         Width           =   11760
         _ExtentX        =   20743
         _ExtentY        =   17039
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   1
            Left            =   1725
            TabIndex        =   33
            Top             =   75
            Width           =   3735
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   2
            Left            =   7155
            TabIndex        =   32
            Top             =   75
            Width           =   3735
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   3
            Left            =   1725
            TabIndex        =   31
            Top             =   435
            Width           =   3735
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   20
            Left            =   7155
            TabIndex        =   30
            Top             =   795
            Width           =   3735
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   7
            Left            =   7155
            TabIndex        =   29
            Top             =   1155
            Width           =   3735
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   6
            Left            =   1725
            TabIndex        =   28
            Top             =   1155
            Width           =   3735
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   5
            Left            =   1725
            TabIndex        =   27
            Top             =   795
            Width           =   3735
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   8
            Left            =   1725
            TabIndex        =   26
            Top             =   1515
            Width           =   3735
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   9
            Left            =   7155
            TabIndex        =   25
            Top             =   1515
            Width           =   3735
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   10
            Left            =   1725
            TabIndex        =   24
            Top             =   1875
            Width           =   9135
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   14
            Left            =   1725
            TabIndex        =   23
            Top             =   3390
            Width           =   9135
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   16
            Left            =   1725
            TabIndex        =   22
            Top             =   2670
            Width           =   3735
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   17
            Left            =   7155
            TabIndex        =   21
            Top             =   2670
            Width           =   3735
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   18
            Left            =   1725
            TabIndex        =   20
            Top             =   2310
            Width           =   3735
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   19
            Left            =   1725
            TabIndex        =   19
            Top             =   3750
            Width           =   9135
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   15
            Left            =   1725
            TabIndex        =   18
            Top             =   3030
            Width           =   3735
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   11
            Left            =   1725
            TabIndex        =   17
            Top             =   4545
            Width           =   3735
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   12
            Left            =   7155
            TabIndex        =   16
            Top             =   4545
            Width           =   3735
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   13
            Left            =   1725
            TabIndex        =   15
            Top             =   5265
            Width           =   3735
         End
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   0
            Left            =   1725
            Style           =   2  '드롭다운 목록
            TabIndex        =   14
            Top             =   5625
            Width           =   3735
         End
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   1
            Left            =   7155
            Style           =   2  '드롭다운 목록
            TabIndex        =   13
            Top             =   5625
            Width           =   3735
         End
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   2
            Left            =   7155
            Style           =   2  '드롭다운 목록
            TabIndex        =   12
            Top             =   5265
            Width           =   3735
         End
         Begin Threed.SSFrame SSFrame5 
            Height          =   735
            Left            =   -150
            TabIndex        =   34
            Top             =   15360
            Width           =   11235
            _ExtentX        =   19817
            _ExtentY        =   1296
            _Version        =   262144
            Caption         =   "계약일/종료일"
            Begin MSComCtl2.DTPicker dtInput 
               Height          =   315
               Index           =   0
               Left            =   1920
               TabIndex        =   35
               Top             =   300
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   57016320
               CurrentDate     =   36684
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   20
               Left            =   300
               TabIndex        =   36
               Top             =   300
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "계약일자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   21
               Left            =   5700
               TabIndex        =   37
               Top             =   300
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "종료일자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin MSComCtl2.DTPicker dtInput 
               Height          =   315
               Index           =   1
               Left            =   7320
               TabIndex        =   38
               Top             =   300
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   57016320
               CurrentDate     =   36684
            End
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   1
            Left            =   75
            TabIndex        =   39
            Top             =   75
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "대리점코드"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   2
            Left            =   5505
            TabIndex        =   40
            Top             =   75
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "대리점명"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   3
            Left            =   75
            TabIndex        =   41
            Top             =   435
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "대표자명"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   33
            Left            =   7155
            TabIndex        =   42
            Top             =   435
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   262144
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   7
               Left            =   360
               TabIndex        =   43
               Top             =   30
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "일수금매장"
            End
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   8
               Left            =   2100
               TabIndex        =   44
               Top             =   30
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "월수금매장"
            End
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   32
            Left            =   5505
            TabIndex        =   45
            Top             =   435
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "수금형태"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   5
            Left            =   75
            TabIndex        =   46
            Top             =   795
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "사업자번호"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   6
            Left            =   75
            TabIndex        =   47
            Top             =   1155
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "업    태"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   7
            Left            =   5505
            TabIndex        =   48
            Top             =   1155
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "업    종"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   34
            Left            =   5505
            TabIndex        =   49
            Top             =   795
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "장부관리코드"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   8
            Left            =   75
            TabIndex        =   50
            Top             =   1515
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "전화번호"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   9
            Left            =   5505
            TabIndex        =   51
            Top             =   1515
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "우편번호"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   10
            Left            =   75
            TabIndex        =   52
            Top             =   1875
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "주    소"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   26
            Left            =   75
            TabIndex        =   53
            Top             =   2670
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "전화번호"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   27
            Left            =   75
            TabIndex        =   54
            Top             =   3030
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "우편번호"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   28
            Left            =   75
            TabIndex        =   55
            Top             =   3390
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "주    소"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   29
            Left            =   5505
            TabIndex        =   56
            Top             =   2670
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "핸드폰번호"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   30
            Left            =   75
            TabIndex        =   57
            Top             =   2310
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "점주성명"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   31
            Left            =   75
            TabIndex        =   58
            Top             =   3750
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "비    고"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   24
            Left            =   7155
            TabIndex        =   59
            Top             =   4905
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   262144
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   4
               Left            =   360
               TabIndex        =   60
               Top             =   30
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "본  사"
            End
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   5
               Left            =   2280
               TabIndex        =   61
               Top             =   30
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "대리점"
            End
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   23
            Left            =   1725
            TabIndex        =   62
            Top             =   4905
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   262144
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   2
               Left            =   360
               TabIndex        =   63
               Top             =   30
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "적  용"
            End
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   3
               Left            =   2280
               TabIndex        =   64
               Top             =   30
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "비적용"
            End
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   22
            Left            =   1725
            TabIndex        =   65
            Top             =   4185
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   262144
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   0
               Left            =   60
               TabIndex        =   66
               Top             =   30
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "일반매장"
            End
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   1
               Left            =   1365
               TabIndex        =   67
               Top             =   30
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "이마트"
            End
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   6
               Left            =   2460
               TabIndex        =   68
               Top             =   30
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "할인매장"
            End
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   11
            Left            =   75
            TabIndex        =   69
            Top             =   4185
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "대리점구분"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   12
            Left            =   5505
            TabIndex        =   70
            Top             =   4545
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "등    급"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   13
            Left            =   75
            TabIndex        =   71
            Top             =   4545
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "마 진 율"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   14
            Left            =   75
            TabIndex        =   72
            Top             =   4905
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "할인적용"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   15
            Left            =   5505
            TabIndex        =   73
            Top             =   4905
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "수선구분"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   16
            Left            =   75
            TabIndex        =   74
            Top             =   5265
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "TAG색상"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   17
            Left            =   75
            TabIndex        =   75
            Top             =   5625
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "담당자명"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   18
            Left            =   5505
            TabIndex        =   76
            Top             =   5625
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "기 사 명"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   19
            Left            =   5505
            TabIndex        =   77
            Top             =   5265
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "목요세일"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   78
         Top             =   15
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   900
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   255
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
         Caption         =   " 가맹점 등록 (P_01001)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_01001_1.frx":051C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Left            =   7635
         TabIndex        =   79
         Top             =   15
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
         PictureBackground=   "P_01001_1.frx":071E
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   80
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
            Picture         =   "P_01001_1.frx":0920
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   81
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "화면"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_01001_1.frx":0EBA
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   82
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
            Picture         =   "P_01001_1.frx":1454
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   83
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
            Picture         =   "P_01001_1.frx":19EE
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   84
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
            Picture         =   "P_01001_1.frx":1F88
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   85
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
            Picture         =   "P_01001_1.frx":2522
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   86
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
            Picture         =   "P_01001_1.frx":2ABC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   87
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
            Picture         =   "P_01001_1.frx":3056
         End
      End
   End
End
Attribute VB_Name = "P_01001_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Dim sPrintOption As String

Public Sub Data_Display()
    ReDim sValue(3)
    
    sValue(0) = "0"
    sValue(1) = txtInput(0).Text & "%"
    sValue(2) = Mid(cboInput(3).Text, 2, 4)
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("PRO_P_01001_00_MASTER", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    'Call spdDisplay(RS01)
    Call fpSpread_Display(spdView, RS01)
    Call GetColWidth(REG_App, Me.Name, spdView)
End Sub

''Private Sub spdDisplay(RS As ADODB.Recordset)
''    Call fpSpread_Display(spdView, RS)
''
''    With spdView
''        .ColsFrozen = 1  '틀고정
''        .Row = -1
''
''        .Col = 1
''        .ColWidth(1) = 10
''        .CellType = CellTypeStaticText
''        .TypeVAlign = TypeVAlignCenter
''        .TypeHAlign = TypeHAlignCenter
''
''        .Col = 2
''        .ColWidth(2) = 14
''        .CellType = CellTypeStaticText
''        .TypeVAlign = TypeVAlignCenter
''        .TypeHAlign = TypeHAlignLeft
''
''        .Col = 3
''        .ColWidth(3) = 14
''        .CellType = CellTypeStaticText
''        .TypeVAlign = TypeVAlignCenter
''        .TypeHAlign = TypeHAlignLeft
''    End With
''End Sub

Private Sub cboInput_Click(Index As Integer)
    If Index = 3 Then
        Call Data_Display
    End If
End Sub

Private Sub cmdPrint_Click()
    Call DataScreen2
    
    panPrint.Visible = False
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: Call Data_Display   ' 조회
        Case 1: Call DataAdd        ' 신규
        Case 2: Call DataSave       ' 저장
        Case 3: Call DataDelete     ' 삭제
        Case 4: Call DataCancel     ' 취소
        Case 5: 'Call DataPrint      ' 인쇄
        Case 6: 'Call DataScreen     '
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
    cmdBtn(0).Enabled = True
    cmdBtn(1).Enabled = True
    cmdBtn(2).Enabled = True
    cmdBtn(3).Enabled = True
    cmdBtn(4).Enabled = True
    cmdBtn(5).Enabled = True
    cmdBtn(6).Enabled = True
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"

    panInput.Caption = ""
    DoEvents
End Sub

Private Sub Form_Load()
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
            
            
        .ColsFrozen = 1  '틀고정
        .Row = -1
    
        .Col = 1
        .ColWidth(1) = 5
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter

        .Col = 2
        .ColWidth(2) = 17
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
    End With
    
    Call MasterComboAdd(cboInput(3))
    
    
    If P_01001_Flag = False Then
        ' Combo BOX의 내역을 채운다.
        'Call ComboAdd
            
        ReDim sValue(2)
        
        sValue(0) = "1"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("PRO_P_01001_00_MASTER", sValue(), Err_Num, Err_Dec)
        
        spdView.MaxCols = RS01.Fields.Count
        spdView.MaxRows = RS01.RecordCount
        
        'Call spdDisplay(RS01)
        Call fpSpread_Display(spdView, RS01)
        Call GetColWidth(REG_App, Me.Name, spdView)
        
        P_01001_Flag = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_01001_Flag = False
End Sub

Private Sub spdView_Click(ByVal Col As Long, ByVal Row As Long)
    Call Data_Display2(Row)
End Sub

Private Sub Data_Display2(Optional iRow As Long = 0)
    Dim i As Integer
    
    ReDim sValue(2)
    
    sValue(0) = "0"
    
    If iRow = 0 Then
        spdView.Row = spdView.ActiveRow
    Else
        spdView.Row = iRow
    End If
    
    spdView.Col = 1
    
    sValue(1) = spdView.Text
    sValue(2) = Mid(cboInput(3).Text, 2, 4)
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("PRO_P_01001_01_MASTER", sValue(), Err_Num, Err_Dec)
    
    If RS01.RecordCount <> 0 Then
        txtInput(1).Text = Trim(RS01!가맹점코드) & ""
        txtInput(2).Text = Trim(RS01!가맹점명) & ""

        If Not IsNull(RS01!수금형태) Then
            If RS01!수금형태 = "1" Then
                optSelect(7).Value = True
            ElseIf RS01!수금형태 = "2" Then
                optSelect(8).Value = True
            End If
        Else
            optSelect(7).Value = False
            optSelect(8).Value = False
        End If
    End If
End Sub

Public Sub DataAdd()
    Dim i As Integer
    
    dtInput(0).Value = Date
    dtInput(1).Value = Date
    
    dtInput(0).Value = ""
    dtInput(1).Value = ""
'
'    For i = 0 To txtInput.Count - 1
'        txtInput(i).Text = ""
'    Next i
'
'    For i = 0 To cboInput.Count - 1
'        cboInput(i).ListIndex = -1
'    Next i
'
'    For i = 0 To optSelect.Count - 1
'        optSelect(i).Value = False
'    Next i
    
    ReDim sValue(0)
    
    sValue(0) = "0"
    
'    Set RS01 = New ADODB.Recordset
'    Set RS01 = ExecPro("PRO_P_01001_04", sValue(), Err_Num, Err_Dec)
'
'    If Not IsNull(RS01!가맹점코드) Then
'        txtInput(1).Text = RS01!가맹점코드
'    Else
'        txtInput(1).Text = "001"
'    End If
    
    txtInput(1).SetFocus
End Sub

Public Sub DataCancel()
    'Call Data_Display2
End Sub

Public Sub DataDelete()
    If MsgBox("해당되는 가맹점코드를 삭제하시겠습니까?", vbYesNo + vbInformation + vbDefaultButton2, "데이터 삭제") = vbYes Then
    
        ReDim sValue(1)
        
        sValue(0) = txtInput(1).Text
        sValue(1) = Mid(cboInput(3).Text, 2, 4)
        
        Call ExecPro("PRO_P_01001_02_MASTER", sValue(), Err_Num, Err_Dec)
        
        If Err_Num = 0 Then
            spdView.Row = spdView.ActiveRow
            spdView.Action = ActionDeleteRow
            
            MsgBox "해당되는 데이터가 정상적으로 삭제가 되었습니다.", vbInformation
        End If
    End If
End Sub

Private Sub ComboAdd()
    ReDim sValue(1)
    
    sValue(0) = "0"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("PRO_T_00001", sValue(), Err_Num, Err_Dec)

    Do While Not RS01.EOF
        cboInput(0).AddItem "[" & RS01!담당자코드 & "] " & RS01!담당자명
        
        RS01.MoveNext
    Loop

    sValue(0) = "0"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("PRO_T_00002", sValue(), Err_Num, Err_Dec)

    Do While Not RS01.EOF
        cboInput(1).AddItem "[" & RS01!기사코드 & "] " & RS01!기사명
        
        RS01.MoveNext
    Loop
    
    cboInput(2).AddItem "[0] 해당없음"
    cboInput(2).AddItem "[5] 목요일"
    cboInput(2).AddItem "[6] 금요일"
    cboInput(2).AddItem "[7] 토요일"
    cboInput(2).AddItem "[1] 일요일"
    cboInput(2).AddItem "[2] 월요일"
    cboInput(2).AddItem "[3] 화요일"
    cboInput(2).AddItem "[4] 수요일"
End Sub

Public Sub DataSave()
    If MsgBox("해당되는 내역을 저장하시겠습니까?", vbYesNo + vbInformation, "데이터 저장") = vbYes Then
        ReDim sValue(28)
        
        PanelsMsg ""
        
        sValue(0) = txtInput(1).Text                                    ' 가맹점코드
        sValue(1) = Mid(cboInput(3).Text, 2, 4)                         ' 대표자명
        sValue(2) = txtInput(2).Text                                    ' 가맹점명
        
        If optSelect(4).Value = True Then                               ' 수선구분
            sValue(3) = "1"
        ElseIf optSelect(5).Value = True Then
            sValue(3) = "2"
        End If


        Call ExecPro("PRO_P_01001_03_MASTER", sValue(), Err_Num, Err_Dec)
        
        If Err_Num = 0 Then
            MsgBox "해당되는 데이터가 정상적으로 저장이 되었습니다.", vbInformation
        End If
    End If
End Sub

Public Sub DataPrint()
    Dim ReportFP As String
    Dim ReportFile As String
    
    ReportFP = GetIniStr("REPORT", "FilePath", "", sIniFile)
    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
    
    P_00000.crPrint.StoredProcParam(0) = "0"
    P_00000.crPrint.StoredProcParam(1) = txtInput(1).Text
    P_00000.crPrint.WindowTitle = Me.Caption
    
    Call ReportPrint(ReportFile, "1")
End Sub

Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Call spdView_Click(NewCol, NewRow)
End Sub

Private Sub spdView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu P_00000.PopUp
    End If
End Sub

Public Sub DataScreen()
    panPrint.Visible = True
    
    sPrintOption = "2"
End Sub

Private Sub DataScreen2()
    Dim ReportFP As String
    Dim ReportFile As String
    
    ReportFP = GetIniStr("REPORT", "FilePath", "", sIniFile)
    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
    
    Dim i As Integer
    For i = 0 To 30
        P_00000.crPrint.Formulas(i) = ""
    Next
    
    P_00000.crPrint.StoredProcParam(0) = "0"
    
    If optPrint(0).Value = True Then
        P_00000.crPrint.StoredProcParam(1) = "0"
    ElseIf optPrint(1).Value = True Then
        P_00000.crPrint.StoredProcParam(1) = "1"
    ElseIf optPrint(2).Value = True Then
        P_00000.crPrint.StoredProcParam(1) = "2"
    End If
    
    P_00000.crPrint.WindowTitle = Me.Caption
    
    If sPrintOption = "2" Then
        Call ReportPrint(ReportFile, "2")
    ElseIf sPrintOption = "1" Then
        Call ReportPrint(ReportFile, "1")
    End If
End Sub

Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub
