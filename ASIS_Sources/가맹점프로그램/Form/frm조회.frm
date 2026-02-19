VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm조회 
   ClientHeight    =   12990
   ClientLeft      =   1455
   ClientTop       =   5745
   ClientWidth     =   18270
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9776.828
   ScaleMode       =   0  '사용자
   ScaleWidth      =   19328.95
   Visible         =   0   'False
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   12990
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18270
      _ExtentX        =   32226
      _ExtentY        =   22913
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm조회.frx":0000
      Begin Threed.SSPanel SSPanel1 
         Height          =   12960
         Index           =   0
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   18240
         _ExtentX        =   32173
         _ExtentY        =   22860
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdSearch 
            Height          =   855
            Index           =   0
            Left            =   150
            TabIndex        =   2
            Top             =   135
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "고객조회/삭제"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton Command1 
            Height          =   855
            Left            =   150
            TabIndex        =   3
            Top             =   1035
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "대리점 정보"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmd3 
            Height          =   855
            Index           =   3
            Left            =   150
            TabIndex        =   4
            Top             =   1935
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdReprint 
            Height          =   855
            Index           =   1
            Left            =   150
            TabIndex        =   5
            Top             =   2835
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "보관증 재출력"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmd2_1 
            Height          =   855
            Index           =   0
            Left            =   150
            TabIndex        =   6
            Top             =   3735
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "일일매출마감"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmd1 
            Height          =   855
            Index           =   4
            Left            =   150
            TabIndex        =   7
            Top             =   4635
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "판매 리스트"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdSearch 
            Height          =   855
            Index           =   1
            Left            =   150
            TabIndex        =   8
            Top             =   5535
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "메시지내역조회"
            ForeColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdSearch 
            Height          =   855
            Index           =   2
            Left            =   150
            TabIndex        =   9
            Top             =   6435
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "편지작성"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmd2 
            Height          =   855
            Index           =   1
            Left            =   3735
            TabIndex        =   10
            Top             =   135
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "월간 매출 현황"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdAmount 
            Height          =   855
            Index           =   2
            Left            =   3735
            TabIndex        =   11
            Top             =   1035
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "고객별 매출액"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton SSCommand1 
            Height          =   855
            Left            =   3735
            TabIndex        =   12
            Top             =   1935
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "고객 미출고 현황"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton SSCommand2 
            Height          =   855
            Left            =   3735
            TabIndex        =   13
            Top             =   2835
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "본사 미출고 현황"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdSearch 
            Height          =   855
            Index           =   4
            Left            =   3735
            TabIndex        =   14
            Top             =   3735
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "본사 출고 현황"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdDiscount 
            Height          =   855
            Index           =   5
            Left            =   3735
            TabIndex        =   15
            Top             =   4635
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "할인관리"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmd3 
            Height          =   855
            Index           =   2
            Left            =   3735
            TabIndex        =   16
            Top             =   5535
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "찾아보기"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdSearch 
            Height          =   855
            Index           =   5
            Left            =   3735
            TabIndex        =   17
            Top             =   6435
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "사고품관리"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdSearch 
            Height          =   855
            Index           =   7
            Left            =   150
            TabIndex        =   18
            Top             =   7335
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdSearch 
            Height          =   855
            Index           =   8
            Left            =   3735
            TabIndex        =   19
            Top             =   7335
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "세탁비환불조회"
            ForeColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmd3 
            Height          =   855
            Index           =   0
            Left            =   7350
            TabIndex        =   20
            Top             =   135
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "쿠폰 조회"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmd3 
            Height          =   855
            Index           =   1
            Left            =   7350
            TabIndex        =   21
            Top             =   1035
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "품목별 현황"
            ForeColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdPrice 
            Height          =   855
            Left            =   7350
            TabIndex        =   22
            Top             =   1935
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "가격 변경"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdSearch 
            Height          =   855
            Index           =   9
            Left            =   7350
            TabIndex        =   23
            Top             =   2835
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "품목보기 순위조정"
            ForeColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdBackup 
            Height          =   855
            Left            =   7350
            TabIndex        =   24
            Top             =   3735
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "DB 관리"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdRepair 
            Height          =   855
            Left            =   7350
            TabIndex        =   25
            Top             =   4635
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "본사연결"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdSearch 
            Height          =   855
            Index           =   3
            Left            =   7350
            TabIndex        =   26
            Top             =   5535
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "입고예정 LIST"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdRestore 
            Height          =   855
            Left            =   7350
            TabIndex        =   27
            Top             =   6435
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "고객별 미수금 현황"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdBtn4 
            Height          =   855
            Index           =   0
            Left            =   10950
            TabIndex        =   28
            Top             =   135
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "프로그램 업그레이드"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdBtn4 
            Height          =   855
            Index           =   1
            Left            =   10950
            TabIndex        =   29
            Top             =   1035
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "출고일자별"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdBtn4 
            Height          =   855
            Index           =   2
            Left            =   10950
            TabIndex        =   30
            Top             =   1935
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "행사 안내용 문자"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdBtn4 
            Height          =   855
            Index           =   3
            Left            =   10950
            TabIndex        =   31
            Top             =   2835
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "세탁물 인도 문자"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdBtn4 
            Height          =   855
            Index           =   4
            Left            =   10950
            TabIndex        =   32
            Top             =   3735
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "문자 일별 현황"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdBtn4 
            Height          =   855
            Index           =   5
            Left            =   10950
            TabIndex        =   33
            Top             =   4635
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "문자 월별 현황"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdBtn4 
            Height          =   855
            Index           =   6
            Left            =   10950
            TabIndex        =   34
            Top             =   5535
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "입고모드 전환"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdBtn4 
            Height          =   855
            Index           =   7
            Left            =   10950
            TabIndex        =   35
            Top             =   6435
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "프로그램 정보"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdSearch 
            Height          =   855
            Index           =   6
            Left            =   7350
            TabIndex        =   36
            Top             =   7335
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "마일리지 현황"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdBtn4 
            Height          =   855
            Index           =   8
            Left            =   10950
            TabIndex        =   37
            Top             =   7335
            Width           =   3225
            _Version        =   851970
            _ExtentX        =   5689
            _ExtentY        =   1508
            _StockProps     =   79
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frm조회"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Option Explicit
''
''Private Sub Cmd1_Click(Index As Integer)
''    Load frm일일판매리스트
''    frm일일판매리스트.Show 1
''End Sub
''
''Private Sub Cmd2_1_Click(Index As Integer)
''    FormChk
''    Load frm일일매출마감
''End Sub
''
''Private Sub Cmd2_Click(Index As Integer)
''    '월간매출현황
''    Select Case Index
''        Case 1
''            FormChk
''            Load frm월간매출
''    End Select
''End Sub
''
''Private Sub Cmd3_Click(Index As Integer)
''    Select Case Index
''        Case 0
''            FormChk
''            Load frm쿠폰조회
''            frm쿠폰조회.Show
''
''            'Load F_Copy
''            'F_Copy.Show
''
''
''        Case 1
''            FormChk
''            Load frm품목별집계현황
''            Exit Sub
''
''            'FormChk
''            'Load frm자료수신
''        Case 2
''            FormChk
''            Load frm고객별이용실적
''        Case 3
''            FormChk
''            Load frm판매취소
''    End Select
''End Sub
''
''Private Sub CmdAmount_Click(Index As Integer)
''    FormChk
''    Load frm고객별매출액
''    frm고객별매출액.Show
''End Sub
''
''Private Sub CmdBackup_Click()
''    FormChk
''    Load DBControl
''    DBControl.Show
''End Sub
''
''Private Sub cmdBtn4_Click(Index As Integer)
''    If Index = 0 Then
''        Load FrmInSoftNetUpdate
''        FrmInSoftNetUpdate.Show 1
''
''    ' 출고 일자별 조회
''    ElseIf Index = 1 Then
''        Load frm출고일자별
''        frm출고일자별.SetFocus
''
''    ' 전체 문자 메시지
''    ElseIf Index = 2 Then
''        P_SMS004.Show
''
''    ' 문자 메시지
''    ElseIf Index = 3 Then
''        frm세탁물인도문자.Show
''
''    ' 문자 메시지 일자별 발송 현황
''    ElseIf Index = 4 Then
''        P_SMS002.Show
''
''    ' 문자 메시지 월별 발송 현황
''    ElseIf Index = 5 Then
''        P_SMS003.Show
''
''
''    '   입고 작업 처리
''    ElseIf Index = 6 Then
''        Dim msg As String
''        msg = "관리자외 절대 건드리지 마십시요" & vbLf & vbLf
''        msg = msg & cmdBtn4(6).Caption & " 하시겠습니까?" & Space(10)
''        If MsgBox(msg, vbCritical + vbYesNo + vbDefaultButton2, "경고") = vbNo Then Exit Sub
''
''        ' 현재 출고 모드로 사용중일 경우
''        If Trim(chkProgramMode) <> "1" Then
''            Call SetIniStr("RUNMODE", "ProgramMode", "1", iniFile)
''
''
''        ' 현재 입고 모드(1대 사용) 한경우
''        Else
''            Call SetIniStr("RUNMODE", "ProgramMode", "2", iniFile)
''        End If
''
''        msg = "프로그램을 다시 시작해야 적용됩니다. 지금 다시 시작 합니다."
''        Call MsgBox(msg, vbInformation, "확인")
''        End
''
''    ElseIf Index = 7 Then
''        Load frmInSoftNet
''        frmInSoftNet.Show 1
''
''    End If
''End Sub
''
''Private Sub CmdDiscount_Click(Index As Integer)
''   ' FormChk
''    Load frm할인관리
''End Sub
''
''Private Sub cmdRepair_Click()
''    Load frm본사전송
''    frm본사전송.Show 1
''End Sub
''
''Private Sub cmdRePrint_Click(Index As Integer)
''    Load frm보관증재출력
''    frm보관증재출력.Show
''End Sub
''
''
''Private Sub CmdRestore_Click()
''    FormChk
''    Load frm고객별미수금
''    frm고객별미수금.Show
''End Sub
''
''Private Sub cmdSearch_Click(Index As Integer)
''    Select Case Index
''        Case 0
''            FormChk
''            Load frm고객조회
''            frm고객조회.Show
''
''        Case 1
''            FormChk
''            Load frm메일내역
''            frm메일내역.Show
''
''        Case 2
''            FormChk
''            Load frm메일작성
''            frm메일작성.Show
''
''        Case 3
''            FormChk
''            Load frm입고예정
''            frm입고예정.Show
''
''        Case 4
''            FormChk
''            Load frm본사출고현황
''            frm본사출고현황.Show
''
''        Case 5
''            FormChk
''            Load frm사고품
''            frm사고품.Show
''
''        Case 6
''            FormChk
''            Load frm마일리지현황
''            frm마일리지현황.Show
''
''        Case 8
''            FormChk
''            Load frm세탁비환불현황
''            frm세탁비환불현황.Show
''
''        Case 9
''            FormChk
''            Load frm참조코드
''            frm참조코드.Show
''    End Select
''End Sub
''
''Private Sub CmdPrice_Click()
''    FormChk
''    Load frm금액변경
''End Sub
''
''Private Sub Command1_Click()
''    Load frm환경설정
''    frm환경설정.Show 1
''End Sub
''
''Private Sub Form_Activate()
''    'frm접수.cmdTagNo.Visible = False
''
''    ' 클라이언트 모드에서 일부 기능 제거
''    If Trim(chkProgramMode) = "2" Then
''        cmdBackup.Enabled = False
''        cmdBtn4(6).Caption = "입고 모드 전환"
''    Else
''        cmdBtn4(6).Caption = "출고 모드 전환"
''    End If
''
''End Sub
''
''Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
''    KeyChk (KeyCode)
''End Sub
''
''Private Sub Form_KeyPress(KeyAscii As Integer)
''    If KeyAscii = 13 Then
''        SendKeys "{TAB}"
''        KeyAscii = 0
''    End If
''End Sub
''
''Private Sub Form_Load()
''    chkinputflig = "조회중" '현재 상태..
''End Sub
''
''Private Sub Form_Unload(Cancel As Integer)
''    chkinputflig = "출고" '현재 상태..
''End Sub
''
''Private Sub SSCommand1_Click()
''    FormChk
''
''    Load frm고객미출고현황
''End Sub
''
''Private Sub SSCommand2_Click()
''    FormChk
''
''    Load frm본사미출고현황
''End Sub
