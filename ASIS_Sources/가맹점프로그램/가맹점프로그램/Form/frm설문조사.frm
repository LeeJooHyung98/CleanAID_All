VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm설문조사 
   BackColor       =   &H00D9E5E9&
   BorderStyle     =   1  '단일 고정
   Caption         =   "설문조사"
   ClientHeight    =   6840
   ClientLeft      =   9675
   ClientTop       =   4620
   ClientWidth     =   9450
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm설문조사.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   9450
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   6840
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   12065
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm설문조사.frx":08CA
      Begin Threed.SSPanel SSPanel 
         Height          =   5805
         Index           =   1
         Left            =   15
         TabIndex        =   5
         Top             =   435
         Width           =   9420
         _ExtentX        =   16616
         _ExtentY        =   10239
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSPanel SSPanel 
            Height          =   540
            Index           =   2
            Left            =   105
            TabIndex        =   6
            Top             =   690
            Width           =   3120
            _ExtentX        =   5503
            _ExtentY        =   953
            _Version        =   262144
            Font3D          =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   " 품질만족도 "
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   540
            Index           =   3
            Left            =   105
            TabIndex        =   7
            Top             =   1260
            Width           =   3120
            _ExtentX        =   5503
            _ExtentY        =   953
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   " 납기만족도 "
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   540
            Index           =   4
            Left            =   105
            TabIndex        =   8
            Top             =   1830
            Width           =   3120
            _ExtentX        =   5503
            _ExtentY        =   953
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   " 컴플레인 처리 만족도 "
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   540
            Index           =   5
            Left            =   105
            TabIndex        =   9
            Top             =   2400
            Width           =   3120
            _ExtentX        =   5503
            _ExtentY        =   953
            _Version        =   262144
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
            Caption         =   " 매장에 대한 지사의 관심도 "
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   540
            Index           =   6
            Left            =   105
            TabIndex        =   10
            Top             =   2970
            Width           =   3120
            _ExtentX        =   5503
            _ExtentY        =   953
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   " 전화 수신 만족도 "
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   540
            Index           =   7
            Left            =   105
            TabIndex        =   11
            Top             =   120
            Width           =   3120
            _ExtentX        =   5503
            _ExtentY        =   953
            _Version        =   262144
            Font3D          =   3
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
            Caption         =   "분류"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   540
            Index           =   8
            Left            =   3255
            TabIndex        =   12
            Top             =   690
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   953
            _Version        =   262144
            Font3D          =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "4 점"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   540
            Index           =   9
            Left            =   3255
            TabIndex        =   13
            Top             =   1260
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   953
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "4 점"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   540
            Index           =   10
            Left            =   3255
            TabIndex        =   14
            Top             =   1830
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   953
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "4 점"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   540
            Index           =   11
            Left            =   3255
            TabIndex        =   15
            Top             =   2400
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   953
            _Version        =   262144
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
            Caption         =   "4 점"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   540
            Index           =   12
            Left            =   3255
            TabIndex        =   16
            Top             =   2970
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   953
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "4 점"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   540
            Index           =   13
            Left            =   3255
            TabIndex        =   17
            Top             =   120
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   953
            _Version        =   262144
            Font3D          =   3
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
            Caption         =   "배점점수"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   540
            Index           =   14
            Left            =   4335
            TabIndex        =   18
            Top             =   690
            Width           =   4995
            _ExtentX        =   8811
            _ExtentY        =   953
            _Version        =   262144
            Font3D          =   3
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin Threed.SSOption optCheck1 
               Height          =   285
               Index           =   0
               Left            =   390
               TabIndex        =   24
               Top             =   135
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSOption optCheck1 
               Height          =   285
               Index           =   1
               Left            =   1395
               TabIndex        =   25
               Top             =   135
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSOption optCheck1 
               Height          =   285
               Index           =   2
               Left            =   2385
               TabIndex        =   26
               Top             =   135
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSOption optCheck1 
               Height          =   285
               Index           =   3
               Left            =   3405
               TabIndex        =   27
               Top             =   135
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSOption optCheck1 
               Height          =   285
               Index           =   4
               Left            =   4410
               TabIndex        =   28
               Top             =   135
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   540
            Index           =   19
            Left            =   4335
            TabIndex        =   19
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   953
            _Version        =   262144
            Font3D          =   3
            CaptionStyle    =   1
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
            Caption         =   "A등급 (4점)"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   540
            Index           =   20
            Left            =   5340
            TabIndex        =   20
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   953
            _Version        =   262144
            Font3D          =   3
            CaptionStyle    =   1
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
            Caption         =   "B등급 (3점)"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   540
            Index           =   21
            Left            =   6345
            TabIndex        =   21
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   953
            _Version        =   262144
            Font3D          =   3
            CaptionStyle    =   1
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
            Caption         =   "C등급 (2점)"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   540
            Index           =   22
            Left            =   7350
            TabIndex        =   22
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   953
            _Version        =   262144
            Font3D          =   3
            CaptionStyle    =   1
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
            Caption         =   "D등급 (1점)"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   540
            Index           =   23
            Left            =   8355
            TabIndex        =   23
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   953
            _Version        =   262144
            Font3D          =   3
            CaptionStyle    =   1
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
            Caption         =   "E등급 (0점)"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   540
            Index           =   15
            Left            =   4335
            TabIndex        =   29
            Top             =   1260
            Width           =   4995
            _ExtentX        =   8811
            _ExtentY        =   953
            _Version        =   262144
            Font3D          =   3
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin Threed.SSOption optCheck2 
               Height          =   285
               Index           =   0
               Left            =   390
               TabIndex        =   30
               Top             =   135
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSOption optCheck2 
               Height          =   285
               Index           =   1
               Left            =   1395
               TabIndex        =   31
               Top             =   135
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSOption optCheck2 
               Height          =   285
               Index           =   2
               Left            =   2385
               TabIndex        =   32
               Top             =   135
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSOption optCheck2 
               Height          =   285
               Index           =   3
               Left            =   3405
               TabIndex        =   33
               Top             =   135
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSOption optCheck2 
               Height          =   285
               Index           =   4
               Left            =   4410
               TabIndex        =   34
               Top             =   135
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   540
            Index           =   16
            Left            =   4335
            TabIndex        =   35
            Top             =   1830
            Width           =   4995
            _ExtentX        =   8811
            _ExtentY        =   953
            _Version        =   262144
            Font3D          =   3
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin Threed.SSOption optCheck3 
               Height          =   285
               Index           =   0
               Left            =   390
               TabIndex        =   36
               Top             =   135
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSOption optCheck3 
               Height          =   285
               Index           =   1
               Left            =   1395
               TabIndex        =   37
               Top             =   135
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSOption optCheck3 
               Height          =   285
               Index           =   2
               Left            =   2385
               TabIndex        =   38
               Top             =   135
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSOption optCheck3 
               Height          =   285
               Index           =   3
               Left            =   3405
               TabIndex        =   39
               Top             =   135
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSOption optCheck3 
               Height          =   285
               Index           =   4
               Left            =   4410
               TabIndex        =   40
               Top             =   135
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   540
            Index           =   17
            Left            =   4335
            TabIndex        =   41
            Top             =   2400
            Width           =   4995
            _ExtentX        =   8811
            _ExtentY        =   953
            _Version        =   262144
            Font3D          =   3
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin Threed.SSOption optCheck4 
               Height          =   285
               Index           =   0
               Left            =   390
               TabIndex        =   42
               Top             =   135
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSOption optCheck4 
               Height          =   285
               Index           =   1
               Left            =   1395
               TabIndex        =   43
               Top             =   135
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSOption optCheck4 
               Height          =   285
               Index           =   2
               Left            =   2385
               TabIndex        =   44
               Top             =   135
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSOption optCheck4 
               Height          =   285
               Index           =   3
               Left            =   3405
               TabIndex        =   45
               Top             =   135
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSOption optCheck4 
               Height          =   285
               Index           =   4
               Left            =   4410
               TabIndex        =   46
               Top             =   135
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   540
            Index           =   18
            Left            =   4335
            TabIndex        =   47
            Top             =   2970
            Width           =   4995
            _ExtentX        =   8811
            _ExtentY        =   953
            _Version        =   262144
            Font3D          =   3
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin Threed.SSOption optCheck5 
               Height          =   285
               Index           =   0
               Left            =   390
               TabIndex        =   48
               Top             =   135
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSOption optCheck5 
               Height          =   285
               Index           =   1
               Left            =   1395
               TabIndex        =   49
               Top             =   135
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSOption optCheck5 
               Height          =   285
               Index           =   2
               Left            =   2385
               TabIndex        =   50
               Top             =   135
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSOption optCheck5 
               Height          =   285
               Index           =   3
               Left            =   3405
               TabIndex        =   51
               Top             =   135
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSOption optCheck5 
               Height          =   285
               Index           =   4
               Left            =   4410
               TabIndex        =   52
               Top             =   135
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   262144
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   540
            Index           =   24
            Left            =   3255
            TabIndex        =   58
            Top             =   3540
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   953
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "20 점"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   540
            Index           =   25
            Left            =   105
            TabIndex        =   59
            Top             =   3540
            Width           =   3120
            _ExtentX        =   5503
            _ExtentY        =   953
            _Version        =   262144
            Font3D          =   3
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
            Caption         =   "합계접수"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   540
            Index           =   26
            Left            =   4335
            TabIndex        =   60
            Top             =   3540
            Width           =   4995
            _ExtentX        =   8811
            _ExtentY        =   953
            _Version        =   262144
            Font3D          =   3
            ForeColor       =   192
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "점 "
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            Alignment       =   4
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin CSTextLibCtl.silgEdit txtScore 
               Height          =   390
               Left            =   3330
               TabIndex        =   61
               Top             =   60
               Width           =   1215
               _Version        =   262145
               _ExtentX        =   2143
               _ExtentY        =   688
               _StockProps     =   125
               Text            =   " 0"
               ForeColor       =   255
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
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
               RawData         =   "0"
               Text            =   " 0"
               StartText.x     =   2
               StartText.y     =   1
               FirstVisPos     =   0
               HiAnchor        =   0
               HiNew           =   0
               CaretHeight     =   24
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
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "    지사가 가맹점의 전화를 안 받는다. -> E등급"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   4
            Left            =   360
            TabIndex        =   57
            Top             =   5490
            Width           =   4830
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "예) 지사가 가맹점의 전화를 잘 받는다. -> A등급"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   56
            Top             =   5220
            Width           =   4830
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "- 전화 수신 만족도는 지사에서의 가맹점 전화 수신율(수신횟수)에 만족도입니다."
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   2
            Left            =   135
            TabIndex        =   55
            Top             =   4920
            Width           =   7980
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "- 컴플레인이 없을 시에는 A등급을 주시면 됩니다."
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   54
            Top             =   4620
            Width           =   4935
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "점수님께서 주시고 싶은 등급에 마우스로 클릭하여 표시하여 주시기 바랍니다."
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
            Index           =   0
            Left            =   135
            TabIndex        =   53
            Top             =   4245
            Width           =   7665
         End
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   570
         Index           =   0
         Left            =   15
         TabIndex        =   2
         Top             =   6255
         Width           =   9420
         _ExtentX        =   16616
         _ExtentY        =   1005
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   45
            TabIndex        =   3
            Top             =   60
            Width           =   1275
            _Version        =   851970
            _ExtentX        =   2249
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            BackColor       =   -2147483633
            Appearance      =   6
            Picture         =   "frm설문조사.frx":093C
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   8100
            TabIndex        =   4
            Top             =   60
            Width           =   1260
            _Version        =   851970
            _ExtentX        =   2222
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 저장(&S)"
            BackColor       =   -2147483633
            Appearance      =   6
            Picture         =   "frm설문조사.frx":134E
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   405
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   9420
         _ExtentX        =   16616
         _ExtentY        =   714
         _Version        =   262144
         Font3D          =   3
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "   설문조사 - 관할지사 평가표 (가맹점의 지사 지지도)"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm설문조사.frx":1D60
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image1 
            Height          =   240
            Left            =   60
            Picture         =   "frm설문조사.frx":21C2
            Top             =   60
            Width           =   240
         End
      End
   End
End
Attribute VB_Name = "frm설문조사"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn

    Select Case Index
        Case 0
            Select Case True
                Case optCheck1(0).Value:
                Case optCheck1(1).Value:
                Case optCheck1(2).Value:
                Case optCheck1(3).Value:
                Case optCheck1(4).Value:
                Case Else:
                    MsgBox "'품질만족도'를 체크해주세요.", vbInformation, "확인"
                    Exit Sub
            End Select
        
            Select Case True
                Case optCheck2(0).Value:
                Case optCheck2(1).Value:
                Case optCheck2(2).Value:
                Case optCheck2(3).Value:
                Case optCheck2(4).Value:
                Case Else:
                    MsgBox "'납기만족도'를  체크해주세요.", vbInformation, "확인"
                    Exit Sub
            End Select
        
            Select Case True
                Case optCheck3(0).Value:
                Case optCheck3(1).Value:
                Case optCheck3(2).Value:
                Case optCheck3(3).Value:
                Case optCheck3(4).Value:
                Case Else:
                    MsgBox "'컴플레인 처리 만족도'를 체크해주세요.", vbInformation, "확인"
                    Exit Sub
            End Select
        
            Select Case True
                Case optCheck4(0).Value:
                Case optCheck4(1).Value:
                Case optCheck4(2).Value:
                Case optCheck4(3).Value:
                Case optCheck4(4).Value:
                Case Else:
                    MsgBox "'매장에 대한 지사의 관심도'를 체크해주세요.", vbInformation, "확인"
                    Exit Sub
            End Select
        
            Select Case True
                Case optCheck5(0).Value:
                Case optCheck5(1).Value:
                Case optCheck5(2).Value:
                Case optCheck5(3).Value:
                Case optCheck5(4).Value:
                Case Else:
                    MsgBox "'전화 수신 만족도'를 체크해주세요.", vbInformation, "확인"
                    Exit Sub
            End Select
        
        
            Query = "SELECT * FROM TB_설문조사"
            Query = Query & " WHERE KEY_DATE   = '110707'"
            Query = Query & "   AND 가맹점코드 = '" & 가맹점정보.가맹점코드 & "'"
            Set ADORs = New ADODB.Recordset
            ADORs.Open Query, HostCon, adOpenDynamic, adLockOptimistic
        
            If ADORs.EOF Then ADORs.AddNew
            
            ADORs!KEY_DATE = "110707"                   '
            ADORs!가맹점코드 = 가맹점정보.가맹점코드 & ""
            ADORs!지사코드 = 가맹점정보.지사코드 & ""
            
            Select Case True
                Case optCheck1(0).Value: ADORs!항목1 = 4
                Case optCheck1(1).Value: ADORs!항목1 = 3
                Case optCheck1(2).Value: ADORs!항목1 = 2
                Case optCheck1(3).Value: ADORs!항목1 = 1
                Case Else: ADORs!항목1 = 0
            End Select
        
            Select Case True
                Case optCheck2(0).Value: ADORs!항목2 = 4
                Case optCheck2(1).Value: ADORs!항목2 = 3
                Case optCheck2(2).Value: ADORs!항목2 = 2
                Case optCheck2(3).Value: ADORs!항목2 = 1
                Case Else: ADORs!항목2 = 0
            End Select
        
            Select Case True
                Case optCheck3(0).Value: ADORs!항목3 = 4
                Case optCheck3(1).Value: ADORs!항목3 = 3
                Case optCheck3(2).Value: ADORs!항목3 = 2
                Case optCheck3(3).Value: ADORs!항목3 = 1
                Case Else: ADORs!항목3 = 0
            End Select
        
            Select Case True
                Case optCheck4(0).Value: ADORs!항목4 = 4
                Case optCheck4(1).Value: ADORs!항목4 = 3
                Case optCheck4(2).Value: ADORs!항목4 = 2
                Case optCheck4(3).Value: ADORs!항목4 = 1
                Case Else: ADORs!항목4 = 0
            End Select
        
            Select Case True
                Case optCheck5(0).Value: ADORs!항목5 = 4
                Case optCheck5(1).Value: ADORs!항목5 = 3
                Case optCheck5(2).Value: ADORs!항목5 = 2
                Case optCheck5(3).Value: ADORs!항목5 = 1
                Case Else: ADORs!항목5 = 0
            End Select
            
            ADORs!설문일자 = Format(Now, "YYYY-MM-DD hh:mm:ss")
            
            ADORs.Update
            
            ADORs.Close
            Set ADORs = Nothing
            
            Unload Me
            
        Case 1:
            Unload Me
    End Select

    Exit Sub

ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    
    
    If Server_Connection(HostCon, "LAUNDRY1000") = True Then
        Query = "SELECT   항목1"
        Query = Query & ",항목2"
        Query = Query & ",항목3"
        Query = Query & ",항목4"
        Query = Query & ",항목5"
        Query = Query & " FROM TB_설문조사"
        Query = Query & " WHERE KEY_DATE   = '110707'"
        Query = Query & "   AND 가맹점코드 = '" & 가맹점정보.가맹점코드 & "'"
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, HostCon, adOpenForwardOnly, adLockReadOnly
                
        If Not ADORs.EOF Then
            '항목1
            If IsNull(ADORs!항목1) Then
                optCheck1(0).Value = False
                optCheck1(1).Value = False
                optCheck1(2).Value = False
                optCheck1(3).Value = False
                optCheck1(4).Value = False
            Else
                Select Case ADORs!항목1
                    Case "4": optCheck1(0).Value = True
                    Case "3": optCheck1(1).Value = True
                    Case "2": optCheck1(2).Value = True
                    Case "1": optCheck1(3).Value = True
                    Case "0": optCheck1(4).Value = True
                End Select
            End If
            
            '항목2
            If IsNull(ADORs!항목2) Then
                optCheck2(0).Value = False
                optCheck2(1).Value = False
                optCheck2(2).Value = False
                optCheck2(3).Value = False
                optCheck2(4).Value = False
            Else
                Select Case ADORs!항목2
                    Case "4": optCheck2(0).Value = True
                    Case "3": optCheck2(1).Value = True
                    Case "2": optCheck2(2).Value = True
                    Case "1": optCheck2(3).Value = True
                    Case "0": optCheck2(4).Value = True
                End Select
            End If
            
            '항목3
            If IsNull(ADORs!항목3) Then
                optCheck3(0).Value = False
                optCheck3(1).Value = False
                optCheck3(2).Value = False
                optCheck3(3).Value = False
                optCheck3(4).Value = False
            Else
                Select Case ADORs!항목3
                    Case "4": optCheck3(0).Value = True
                    Case "3": optCheck3(1).Value = True
                    Case "2": optCheck3(2).Value = True
                    Case "1": optCheck3(3).Value = True
                    Case "0": optCheck3(4).Value = True
                End Select
            End If
            
            '항목4
            If IsNull(ADORs!항목4) Then
                optCheck4(0).Value = False
                optCheck4(1).Value = False
                optCheck4(2).Value = False
                optCheck4(3).Value = False
                optCheck4(4).Value = False
            Else
                Select Case ADORs!항목4
                    Case "4": optCheck4(0).Value = True
                    Case "3": optCheck4(1).Value = True
                    Case "2": optCheck4(2).Value = True
                    Case "1": optCheck4(3).Value = True
                    Case "0": optCheck4(4).Value = True
                End Select
            End If
            
            '항목5
            If IsNull(ADORs!항목5) Then
                optCheck5(0).Value = False
                optCheck5(1).Value = False
                optCheck5(2).Value = False
                optCheck5(3).Value = False
                optCheck5(4).Value = False
            Else
                Select Case ADORs!항목5
                    Case "4": optCheck5(0).Value = True
                    Case "3": optCheck5(1).Value = True
                    Case "2": optCheck5(2).Value = True
                    Case "1": optCheck5(3).Value = True
                    Case "0": optCheck5(4).Value = True
                End Select
            End If
        End If
        ADORs.Close
        Set ADORs = Nothing
    End If
    
    Exit Sub

ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub optCheck1_Click(Index As Integer, Value As Integer)
    Call 점수계산
End Sub

Private Sub optCheck2_Click(Index As Integer, Value As Integer)
    Call 점수계산
End Sub

Private Sub optCheck3_Click(Index As Integer, Value As Integer)
    Call 점수계산
End Sub

Private Sub optCheck4_Click(Index As Integer, Value As Integer)
    Call 점수계산
End Sub

Private Sub optCheck5_Click(Index As Integer, Value As Integer)
    Call 점수계산
End Sub

Private Sub 점수계산()
    txtScore.Value = 0
    
    Select Case True
        Case optCheck1(0).Value: txtScore.Value = txtScore.Value + 4
        Case optCheck1(1).Value: txtScore.Value = txtScore.Value + 3
        Case optCheck1(2).Value: txtScore.Value = txtScore.Value + 2
        Case optCheck1(3).Value: txtScore.Value = txtScore.Value + 1
    End Select

    Select Case True
        Case optCheck2(0).Value: txtScore.Value = txtScore.Value + 4
        Case optCheck2(1).Value: txtScore.Value = txtScore.Value + 3
        Case optCheck2(2).Value: txtScore.Value = txtScore.Value + 2
        Case optCheck2(3).Value: txtScore.Value = txtScore.Value + 1
    End Select

    Select Case True
        Case optCheck3(0).Value: txtScore.Value = txtScore.Value + 4
        Case optCheck3(1).Value: txtScore.Value = txtScore.Value + 3
        Case optCheck3(2).Value: txtScore.Value = txtScore.Value + 2
        Case optCheck3(3).Value: txtScore.Value = txtScore.Value + 1
    End Select

    Select Case True
        Case optCheck4(0).Value: txtScore.Value = txtScore.Value + 4
        Case optCheck4(1).Value: txtScore.Value = txtScore.Value + 3
        Case optCheck4(2).Value: txtScore.Value = txtScore.Value + 2
        Case optCheck4(3).Value: txtScore.Value = txtScore.Value + 1
    End Select

    Select Case True
        Case optCheck5(0).Value: txtScore.Value = txtScore.Value + 4
        Case optCheck5(1).Value: txtScore.Value = txtScore.Value + 3
        Case optCheck5(2).Value: txtScore.Value = txtScore.Value + 2
        Case optCheck5(3).Value: txtScore.Value = txtScore.Value + 1
    End Select
End Sub
