VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm고객수정 
   BorderStyle     =   1  '단일 고정
   Caption         =   "고객 등록/수정"
   ClientHeight    =   5295
   ClientLeft      =   15330
   ClientTop       =   5970
   ClientWidth     =   6930
   Icon            =   "frm고객수정.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   6930
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   5295
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6930
      _ExtentX        =   12224
      _ExtentY        =   9340
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm고객수정.frx":0A02
      Begin Threed.SSPanel SSPanel9 
         Height          =   570
         Left            =   15
         TabIndex        =   13
         Top             =   4710
         Width           =   6900
         _ExtentX        =   12171
         _ExtentY        =   1005
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton btnUpdate 
            Height          =   480
            Left            =   5550
            TabIndex        =   9
            Top             =   45
            Width           =   1260
            _Version        =   851970
            _ExtentX        =   2222
            _ExtentY        =   847
            _StockProps     =   79
            Caption         =   " 저장(&S)"
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
            Picture         =   "frm고객수정.frx":0A54
         End
         Begin XtremeSuiteControls.PushButton btnCancel 
            Height          =   480
            Left            =   45
            TabIndex        =   14
            Top             =   45
            Width           =   1260
            _Version        =   851970
            _ExtentX        =   2222
            _ExtentY        =   847
            _StockProps     =   79
            Caption         =   " 취소(&X)"
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
            Picture         =   "frm고객수정.frx":1466
         End
         Begin XtremeSuiteControls.PushButton btn_Kakao_Invite 
            Height          =   480
            Left            =   3300
            TabIndex        =   37
            Top             =   45
            Width           =   2205
            _Version        =   851970
            _ExtentX        =   3889
            _ExtentY        =   847
            _StockProps     =   79
            Caption         =   "카카오 알림톡 전송"
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
            Picture         =   "frm고객수정.frx":1A00
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   4680
         Left            =   15
         TabIndex        =   15
         Top             =   15
         Width           =   6900
         _ExtentX        =   12171
         _ExtentY        =   8255
         _Version        =   262144
         BackColor       =   16777215
         PictureBackgroundStyle=   2
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox cbo080 
            Enabled         =   0   'False
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
            ItemData        =   "frm고객수정.frx":1D52
            Left            =   5205
            List            =   "frm고객수정.frx":1D5C
            Style           =   2  '드롭다운 목록
            TabIndex        =   35
            Top             =   3420
            Width           =   1605
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   30
            Index           =   0
            Left            =   1665
            TabIndex        =   20
            Top             =   3780
            Width           =   5115
            _ExtentX        =   9022
            _ExtentY        =   53
            _Version        =   262144
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin VB.ComboBox cboSMS 
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
            ItemData        =   "frm고객수정.frx":1D7C
            Left            =   1665
            List            =   "frm고객수정.frx":1D86
            Style           =   2  '드롭다운 목록
            TabIndex        =   6
            Top             =   3420
            Width           =   2085
         End
         Begin VB.TextBox txtTel 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   10  '한글 
            Left            =   1665
            TabIndex        =   1
            Top             =   840
            Width           =   2625
         End
         Begin VB.TextBox txtCode 
            Alignment       =   2  '가운데 맞춤
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1665
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   60
            Width           =   1005
         End
         Begin VB.TextBox txtName 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   10  '한글 
            Left            =   1665
            TabIndex        =   0
            Top             =   450
            Width           =   2625
         End
         Begin VB.TextBox txtHP 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   10  '한글 
            Left            =   1665
            TabIndex        =   2
            Top             =   1230
            Width           =   2625
         End
         Begin VB.TextBox txtAdd 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   10  '한글 
            Left            =   1665
            TabIndex        =   3
            Top             =   1620
            Width           =   5130
         End
         Begin VB.TextBox txtMemo 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1035
            IMEMode         =   10  '한글 
            Left            =   1665
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   4
            Top             =   2010
            Width           =   5130
         End
         Begin VB.ComboBox cboClass 
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
            ItemData        =   "frm고객수정.frx":1DA4
            Left            =   1665
            List            =   "frm고객수정.frx":1DA6
            Style           =   2  '드롭다운 목록
            TabIndex        =   5
            Top             =   3075
            Width           =   2085
         End
         Begin CSTextLibCtl.sitxEdit txtCard 
            Height          =   360
            Left            =   1665
            TabIndex        =   7
            Top             =   3870
            Width           =   840
            _Version        =   262145
            _ExtentX        =   1482
            _ExtentY        =   635
            _StockProps     =   125
            Text            =   "______"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
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
            StartText.y     =   4
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
         Begin CSTextLibCtl.sidbEdit txtMoney 
            Height          =   360
            Left            =   5295
            TabIndex        =   8
            Top             =   3870
            Width           =   1215
            _Version        =   262145
            _ExtentX        =   2143
            _ExtentY        =   635
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
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
            CaretHeight     =   16
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
         Begin CSTextLibCtl.sidbEdit txtMileage 
            Height          =   360
            Index           =   0
            Left            =   1665
            TabIndex        =   10
            Top             =   4260
            Width           =   1215
            _Version        =   262145
            _ExtentX        =   2143
            _ExtentY        =   635
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
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
            CaretHeight     =   16
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
         Begin CSTextLibCtl.sidbEdit txtMileage 
            Height          =   360
            Index           =   1
            Left            =   5295
            TabIndex        =   11
            Top             =   4260
            Width           =   1215
            _Version        =   262145
            _ExtentX        =   2143
            _ExtentY        =   635
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
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
            CaretHeight     =   16
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
         Begin XtremeSuiteControls.PushButton btnCopy 
            Height          =   360
            Left            =   4350
            TabIndex        =   33
            Top             =   870
            Width           =   2340
            _Version        =   851970
            _ExtentX        =   4128
            _ExtentY        =   635
            _StockProps     =   79
            Caption         =   " 전화번호 -> 휴대전화"
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
            Picture         =   "frm고객수정.frx":1DA8
         End
         Begin XtremeSuiteControls.PushButton btnCheck 
            Height          =   360
            Left            =   4350
            TabIndex        =   38
            Top             =   1260
            Width           =   2340
            _Version        =   851970
            _ExtentX        =   4128
            _ExtentY        =   635
            _StockProps     =   79
            Caption         =   " 고객확인"
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
            Picture         =   "frm고객수정.frx":27BA
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "080 수신거부:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   13
            Left            =   3810
            TabIndex        =   36
            Top             =   3480
            Width           =   1365
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "작성 후 엔터"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   12
            Left            =   4470
            TabIndex        =   34
            Top             =   570
            Width           =   1260
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "사용가능마일리지:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   11
            Left            =   3465
            TabIndex        =   32
            Top             =   4335
            Width           =   1785
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "미수금:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   3645
            TabIndex        =   31
            Top             =   3945
            Width           =   1605
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "누적마일리지:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   9
            Left            =   0
            TabIndex        =   30
            Top             =   4335
            Width           =   1605
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "카드번호:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   0
            TabIndex        =   29
            Top             =   3945
            Width           =   1605
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "SMS 전송여부:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   0
            TabIndex        =   28
            Top             =   3480
            Width           =   1605
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "고객등급:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   0
            TabIndex        =   27
            Top             =   3135
            Width           =   1605
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "메모:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   0
            TabIndex        =   26
            Top             =   2100
            Width           =   1605
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "주소:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   0
            TabIndex        =   25
            Top             =   1710
            Width           =   1605
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "휴대전화:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   0
            TabIndex        =   24
            Top             =   1320
            Width           =   1605
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "전화번호:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   0
            TabIndex        =   23
            Top             =   930
            Width           =   1605
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "성명:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   22
            Top             =   525
            Width           =   1605
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "고객코드:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   21
            Top             =   135
            Width           =   1605
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
            Index           =   12
            Left            =   6570
            TabIndex        =   19
            Top             =   4365
            Width           =   180
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
            Index           =   11
            Left            =   2940
            TabIndex        =   18
            Top             =   4365
            Width           =   180
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
            Left            =   6570
            TabIndex        =   17
            Top             =   3975
            Width           =   180
         End
      End
   End
End
Attribute VB_Name = "frm고객수정"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btn_Kakao_Invite_Click()
    Call send_Kakao_Invite(txtHP.Text, txtName.Text)
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnCheck_Click()

    Dim Query As String
    
    If txtTel.Text = "" Then
        MsgBox ("전화번호를 입력하여 주십시요")
        Exit Sub
    End If
    
    If txtHP.Text = "" Then
        MsgBox ("휴대폰 번호를 입력하여 주십시요")
        Exit Sub
    End If
    
    
    If Len(txtTel.Text) >= 7 Then txtTel.Text = TelePhone_Number(Trim(txtTel.Text))
    If Len(txtHP.Text) >= 7 Then txtHP.Text = TelePhone_Number(Trim(txtHP.Text))
    
    Query = "SELECT 고객코드 FROM TB_고객정보"
    Query = Query & " WHERE (전화번호 = '" & Trim(txtTel.Text) & "' OR 휴대전화 = '" & Trim(txtTel.Text) & "')"
    Query = Query & " OR (전화번호 = '" & Trim(txtHP.Text) & "' OR 휴대전화 = '" & Trim(txtHP.Text) & "')"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenDynamic, adLockOptimistic
        
    If Not ADORs.EOF Then
        MsgBox ("이미 등록된 고객입니다.")
'        txtCode.Text = ADORs!고객코드
'        고객정보_Display
    Else
        btnUpdate.Enabled = True
    End If
End Sub

Private Sub btnCopy_Click()
    If txtTel.Text <> "" Then
        txtHP.Text = Trim(txtTel.Text)
    End If
    If Len(txtTel.Text) >= 7 Then txtTel.Text = TelePhone_Number(Trim(txtTel.Text))
    If Len(txtHP.Text) >= 7 Then txtHP.Text = TelePhone_Number(Trim(txtHP.Text))
    

End Sub

Private Sub btnUpdate_Click()
    Dim 수정일자 As String
    Dim AddUser As Boolean
    
    AddUser = False
    
    On Error GoTo ErrRtn
        
    If (Trim(txtTel.Text)) = "" Then
        MsgBox "전화번호를 입력해 주십시요.", vbInformation, "확인"
        
        txtTel.SetFocus
        Exit Sub
        
    ElseIf (Trim(txtHP.Text)) = "" Then
        MsgBox "휴대폰번호를 입력해 주십시요.", vbInformation, "확인"
        
        txtHP.SetFocus
        Exit Sub
    End If
        
    If (Trim(txtName.Text)) = "" Then
        MsgBox "성명 입력해 주십시요.", vbInformation, "확인"
        
        txtName.SetFocus
        Exit Sub
    End If
        
    If Trim(txtCode.Text) = "" Then
        txtCode.Text = Get_CustomNo
    End If
    
    If Len(txtTel.Text) >= 7 Then txtTel.Text = TelePhone_Number(Trim(txtTel.Text))
    If Len(txtHP.Text) >= 7 Then txtHP.Text = TelePhone_Number(Trim(txtHP.Text))
    
    '----------------------------------------------------------------
    '
    '----------------------------------------------------------------
    수정일자 = Format(Now, "YYYY-MM-DD hh:mm:ss")
    
    Query = "SELECT * FROM TB_고객정보"
    Query = Query & " WHERE 고객코드 = '" & Trim(txtCode.Text) & "'"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenDynamic, adLockOptimistic
        
    If ADORs.EOF Then
        ADORs.AddNew
    
        ADORs!고객코드 = Trim(txtCode.Text) & ""              ' 1
        ADORs!등록일자 = Format(Date, "YYYY-MM-DD")           ' 2
        ADORs!수정일자 = ""                                   ' 3
        ADORs!이용횟수 = 0                                    ' 4
        ADORs!총접수금액 = 0                                  ' 5
        ADORs!삭제 = 0                                        ' 8
        ADORs!최종거래일자 = ""                               ' 9
        AddUser = True
    Else
        ADORs!수정일자 = 수정일자                             '10
    End If
    
    ADORs!성명 = SubSQuotA(Trim(txtName.Text)) & ""           '12
    ADORs!전화번호 = txtTel.Text & ""                         '13
    ADORs!휴대전화 = txtHP.Text & ""                          '14
    ADORs!주소 = SubSQuotA(Trim(txtAdd.Text)) & ""            '15
    ADORs!미수금액 = txtMoney.Value                           '16
    ADORs!카드번호 = txtCard.Text & ""                        '17
    ADORs!문자발송여부 = Left(cboSMS.Text, 1)                 '18
    ADORs!메모 = SubSQuotA(txtMemo.Text) & ""                 '19
    ADORs!고객등급코드 = Left(cboClass.Text, 1)               '20
    ADORs!본사전송여부 = "N"                                  '21
    ADORs!지사코드 = 가맹점정보.지사코드 & ""                 '22
    ADORs!가맹점코드 = 가맹점정보.가맹점코드 & ""             '23
    
''    ADORs!누적마일리지 = txtMileage(0).Value                  ' 6
''    ADORs!사용가능마일리지 = txtMileage(1).Value              ' 7
    
    ADORs.Update
    
    ADORs.Close
    Set ADORs = Nothing
    
    '----------------------------------------------------------------
    ' TB_미수금수정 - 미수금액을 수정한 경우
    '----------------------------------------------------------------
    If txtMoney.Tag = "" Then
        '신규 고객
    Else
        If CStr(txtMoney.Value) <> CStr(txtMoney.Tag) Then
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
            ADORs!이전미수금 = txtMoney.Tag & ""          ' 6
            ADORs!내용 = "조정 - 고객수정"                ' 7
            
            ADORs.Update
            
            ADORs.Close
            Set ADORs = Nothing
        End If
    End If
    
    
    Call Get_고객정보(txtCode.Text)
    
'    If AddUser Then
'        Call send_Kakao_Invite(txtHP.Text, txtName.Text)
'    End If
    
    Unload Me
    
    If ActiveForm = "접수" Then
        Call frm접수.고객정보_Display(고객정보.고객코드)
    Else
        Call frm출고.고객정보_Display(고객정보.고객코드)
    End If
    
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
    
    Screen.MousePointer = 0
End Sub

Private Sub cboClass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub

Private Sub cboSMS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub

Private Sub Form_Activate()
    Call 고객정보_Display
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        KeyAscii = 0
        
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2

    
    cboSMS.ListIndex = 0
    
    Call 고객등급_Display(cboClass, False)
    
    cboClass.ListIndex = 2
    If txtCode.Text = "" Then
        btnUpdate.Enabled = False
    Else
        btnCheck.Enabled = False
    End If
    
    Exit Sub
    
ErrRtn:

End Sub

Private Sub txtAdd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub

Private Sub txtCard_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub



Private Sub txtHP_GotFocus()
    txtHP.SelStart = 0
    txtHP.SelLength = Len(txtHP.Text)
End Sub

Private Sub 고객정보_Display()
    On Error GoTo ErrRtn
    
    If txtCode.Text = "" Then Exit Sub
    
    '------------------------------------------------------------------------
    ' TB_고객정보
    '------------------------------------------------------------------------
    Query = "SELECT * FROM TB_고객정보"
    Query = Query & " WHERE 고객코드 = '" & Trim(txtCode.Text) & "'"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If Not ADORs.EOF Then
        txtCode.Text = Trim(ADORs!고객코드) & ""          ' 1
        txtName.Text = ADORs!성명 & ""                    ' 2
        txtTel.Text = Trim(ADORs!전화번호) & ""           ' 3
        txtHP.Text = Trim(ADORs!휴대전화) & ""            ' 4
        
        If ADORs!문자발송여부 = "N" Then
            cboSMS.ListIndex = 1                          ' 6
        Else
            cboSMS.ListIndex = 0                          ' 6
        End If
        
        If ADORs!문자080거부여부 = "Y" Then
            cbo080.ListIndex = 0                          ' 6
            cboSMS.Enabled = False
        Else
            cbo080.ListIndex = 1                          ' 6
            cboSMS.Enabled = True
        End If
        
        txtAdd.Text = Trim(ADORs!주소) & ""               ' 5
        
        txtMoney.Value = ADORs!미수금액 & ""              ' 7
        txtMoney.Tag = ADORs!미수금액 & ""                ' 7
        
        txtCard.Text = Trim(ADORs!카드번호) & ""          ' 8
        txtMemo.Text = Trim(ADORs!메모) & ""              ' 9
        
        With cboClass                                     '10
            For i = 0 To .ListCount - 1
                If Left(.List(i), 1) = ADORs!고객등급코드 Then
                    .ListIndex = i
                    
                    Exit For
                End If
            Next i
        End With
        
        txtMileage(0).Value = ADORs!누적마일리지 & ""     '11
        txtMileage(0).Tag = ADORs!누적마일리지 & ""       '11
        
        txtMileage(1).Value = ADORs!사용가능마일리지 & "" '12
        txtMileage(1).Tag = ADORs!사용가능마일리지 & ""   '12
    End If
    ADORs.Close
    Set ADORs = Nothing
    If txtCode.Text <> "" Then btnUpdate.Enabled = True
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
    
    Screen.MousePointer = 0
End Sub

Private Sub txtHP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub

Private Sub txtMileage_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub

Private Sub txtMoney_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub

Private Sub txtTel_Change()
    If txtCode.Text = "" Then btnUpdate.Enabled = False
End Sub

Private Sub txthp_Change()
    If txtCode.Text = "" Then btnUpdate.Enabled = False
End Sub

Private Sub txtTel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
