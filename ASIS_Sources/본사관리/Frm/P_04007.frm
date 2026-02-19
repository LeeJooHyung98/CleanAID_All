VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_04007 
   Caption         =   "예상매출관리"
   ClientHeight    =   9300
   ClientLeft      =   1620
   ClientTop       =   1920
   ClientWidth     =   16950
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_04007.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9300
   ScaleWidth      =   16950
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   9300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16950
      _ExtentX        =   29898
      _ExtentY        =   16404
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_04007.frx":058A
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   6930
         Index           =   0
         Left            =   15
         TabIndex        =   1
         Top             =   1335
         Width           =   16920
         _Version        =   524288
         _ExtentX        =   29845
         _ExtentY        =   12224
         _StockProps     =   64
         BackColorStyle  =   1
         ColsFrozen      =   5
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         MaxCols         =   19
         MaxRows         =   502
         RowsFrozen      =   2
         SpreadDesigner  =   "P_04007.frx":063C
         UserResize      =   0
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   1005
         Index           =   1
         Left            =   15
         TabIndex        =   2
         Top             =   8280
         Width           =   16920
         _Version        =   524288
         _ExtentX        =   29845
         _ExtentY        =   1773
         _StockProps     =   64
         BackColorStyle  =   1
         ColsFrozen      =   3
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         EditModePermanent=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         MaxCols         =   17
         MaxRows         =   3
         ScrollBars      =   0
         SpreadDesigner  =   "P_04007.frx":22DEC
         UserResize      =   0
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   3
         Top             =   540
         Width           =   16920
         _ExtentX        =   29845
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Left            =   1530
            TabIndex        =   4
            Top             =   60
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy"
            Format          =   68485123
            UpDown          =   -1  'True
            CurrentDate     =   37140
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   7
            Left            =   4410
            TabIndex        =   5
            Top             =   60
            Width           =   5955
            _ExtentX        =   10504
            _ExtentY        =   556
            _Version        =   262144
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   0
               Left            =   360
               TabIndex        =   6
               Top             =   30
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "전  체"
               Value           =   -1
            End
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   1
               Left            =   1680
               TabIndex        =   7
               Top             =   30
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "대리점"
            End
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   2
               Left            =   3120
               TabIndex        =   8
               Top             =   30
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "백화점"
            End
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   3
               Left            =   4500
               TabIndex        =   9
               Top             =   30
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "할인매장"
            End
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   10
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "연    도"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   6
            Left            =   2940
            TabIndex        =   11
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "구    분"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   8
            Left            =   13320
            TabIndex        =   12
            Top             =   60
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
            _Version        =   262144
            BackColor       =   16761024
            Caption         =   "금액단위 : 천원"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   13
         Top             =   15
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   900
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   255
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
         Caption         =   " #"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_04007.frx":237CF
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   1
         Left            =   9345
         TabIndex        =   14
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
         PictureBackground=   "P_04007.frx":239D1
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   15
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
            Picture         =   "P_04007.frx":23BD3
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   16
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "엑셀"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_04007.frx":2416D
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   17
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
            Picture         =   "P_04007.frx":24707
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   18
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
            Picture         =   "P_04007.frx":24CA1
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   19
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
            Picture         =   "P_04007.frx":2523B
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   20
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
            Picture         =   "P_04007.frx":257D5
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   21
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
            Picture         =   "P_04007.frx":25D6F
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   22
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
            Picture         =   "P_04007.frx":26309
         End
      End
   End
End
Attribute VB_Name = "P_04007"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Dim spRow As Integer

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: Call Data_Display   ' 조회
        Case 1: 'Call DataAdd        ' 신규
        Case 2: Call DataSave       ' 저장
        Case 3: Call DataDelete     ' 삭제
        Case 4: 'Call DataCancel     ' 취소
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
    cmdBtn(2).Enabled = True
    cmdBtn(3).Enabled = True
    cmdBtn(5).Enabled = True
    cmdBtn(6).Enabled = True
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    If P_04007_Flag = False Then
        dtInput.Value = Date
        
        P_04007_Flag = True
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i As Integer
    Dim j As Integer
    Dim qYear As String
    Dim nTotal(1) As Long
    
    ReDim sValue(0)
    
    sValue(0) = "0"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_04007_00", sValue(), Err_Num, Err_Dec)
    
    qYear = Val(Format(dtInput.Value, "yyyy")) - 1
    
    For i = 1 To spdView(0).MaxRows
        For j = 1 To spdView(0).MaxCols
            spdView(0).Col = j
            spdView(0).Row = i
            spdView(0).Text = ""
        Next j
    Next i
    
    i = 2
    j = 0
    spRow = 0
    
    Do While Not RS01.EOF
        If optSelect(0).Value = True Then
            i = i + 1
            spdView(0).Row = i
            spdView(0).Col = 1
            spdView(0).Value = RS01!대리점코드
            spdView(0).Col = 2
            spdView(0).Value = Format(dtInput.Value, "yyyy")
            spdView(0).Col = 3
            spdView(0).Value = "수량"
            
            i = i + 1
            spdView(0).Row = i
            spdView(0).Col = 1
            spdView(0).Value = RS01!대리점명
            spdView(0).Col = 3
            spdView(0).Value = "금액"
            
            i = i + 1
            spdView(0).Row = i
            spdView(0).Col = 2
            spdView(0).Value = qYear
            spdView(0).Col = 3
            spdView(0).Value = "수량"
            
            i = i + 1
            spdView(0).Row = i
            spdView(0).Col = 3
            spdView(0).Value = "금액"
            
            j = j + 1
        ElseIf optSelect(1).Value = True Then
            If RS01!구분 = "1" Then
                i = i + 1
                spdView(0).Row = i
                spdView(0).Col = 1
                spdView(0).Value = RS01!대리점코드
                spdView(0).Col = 2
                spdView(0).Value = Format(dtInput.Value, "yyyy")
                spdView(0).Col = 3
                spdView(0).Value = "수량"
                
                i = i + 1
                spdView(0).Row = i
                spdView(0).Col = 1
                spdView(0).Value = RS01!대리점명
                spdView(0).Col = 3
                spdView(0).Value = "금액"
                
                i = i + 1
                spdView(0).Row = i
                spdView(0).Col = 2
                spdView(0).Value = qYear
                spdView(0).Col = 3
                spdView(0).Value = "수량"
                
                i = i + 1
                spdView(0).Row = i
                spdView(0).Col = 3
                spdView(0).Value = "금액"
                
                j = j + 1
            End If
        ElseIf optSelect(2).Value = True Then
            If RS01!구분 = "2" Then
                i = i + 1
                spdView(0).Row = i
                spdView(0).Col = 1
                spdView(0).Value = RS01!대리점코드
                spdView(0).Col = 2
                spdView(0).Value = Format(dtInput.Value, "yyyy")
                spdView(0).Col = 3
                spdView(0).Value = "수량"
                
                i = i + 1
                spdView(0).Row = i
                spdView(0).Col = 1
                spdView(0).Value = RS01!대리점명
                spdView(0).Col = 3
                spdView(0).Value = "금액"
                
                i = i + 1
                spdView(0).Row = i
                spdView(0).Col = 2
                spdView(0).Value = qYear
                spdView(0).Col = 3
                spdView(0).Value = "수량"
                
                i = i + 1
                spdView(0).Row = i
                spdView(0).Col = 3
                spdView(0).Value = "금액"
                
                j = j + 1
            End If
        ElseIf optSelect(3).Value = True Then
            If RS01!구분 = "3" Then
                i = i + 1
                spdView(0).Row = i
                spdView(0).Col = 1
                spdView(0).Value = RS01!대리점코드
                spdView(0).Col = 2
                spdView(0).Value = Format(dtInput.Value, "yyyy")
                spdView(0).Col = 3
                spdView(0).Value = "수량"
                
                i = i + 1
                spdView(0).Row = i
                spdView(0).Col = 1
                spdView(0).Value = RS01!대리점명
                spdView(0).Col = 3
                spdView(0).Value = "금액"
                
                i = i + 1
                spdView(0).Row = i
                spdView(0).Col = 2
                spdView(0).Value = qYear
                spdView(0).Col = 3
                spdView(0).Value = "수량"
                
                i = i + 1
                spdView(0).Row = i
                spdView(0).Col = 3
                spdView(0).Value = "금액"
                
                j = j + 1
            End If
        End If
        
        RS01.MoveNext
    Loop
    
    spRow = i
    
    ReDim sValue(0)
    
    sValue(0) = Format(dtInput.Value, "yyyy")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_04007_01", sValue(), Err_Num, Err_Dec)
    
    Do While Not RS01.EOF
        If RS01!대리점코드 = "QQQ" Then
            spdView(0).Row = 1
            spdView(0).Col = 6: spdView(0).Value = RS01!수량01
            spdView(0).Col = 7: spdView(0).Value = RS01!수량02
            spdView(0).Col = 8: spdView(0).Value = RS01!수량03
            spdView(0).Col = 9: spdView(0).Value = RS01!수량04
            spdView(0).Col = 10: spdView(0).Value = RS01!수량05
            spdView(0).Col = 11: spdView(0).Value = RS01!수량06
            spdView(0).Col = 13: spdView(0).Value = RS01!수량07
            spdView(0).Col = 14: spdView(0).Value = RS01!수량08
            spdView(0).Col = 15: spdView(0).Value = RS01!수량09
            spdView(0).Col = 16: spdView(0).Value = RS01!수량10
            spdView(0).Col = 17: spdView(0).Value = RS01!수량11
            spdView(0).Col = 18: spdView(0).Value = RS01!수량12
            
            spdView(0).Row = 2
            spdView(0).Col = 6: spdView(0).Value = RS01!금액01
            spdView(0).Col = 7: spdView(0).Value = RS01!금액02
            spdView(0).Col = 8: spdView(0).Value = RS01!금액03
            spdView(0).Col = 9: spdView(0).Value = RS01!금액04
            spdView(0).Col = 10: spdView(0).Value = RS01!금액05
            spdView(0).Col = 11: spdView(0).Value = RS01!금액06
            spdView(0).Col = 13: spdView(0).Value = RS01!금액07
            spdView(0).Col = 14: spdView(0).Value = RS01!금액08
            spdView(0).Col = 15: spdView(0).Value = RS01!금액09
            spdView(0).Col = 16: spdView(0).Value = RS01!금액10
            spdView(0).Col = 17: spdView(0).Value = RS01!금액11
            spdView(0).Col = 18: spdView(0).Value = RS01!금액12
        Else
            i = 3
            
            While i < spRow
                spdView(0).Row = i
                spdView(0).Col = 1
            
                If spdView(0).Value = RS01!대리점코드 Then
                    spdView(0).Row = i
                    spdView(0).Col = 5: spdView(0).Value = RS01!합계수량
                    spdView(0).Col = 6: spdView(0).Value = RS01!수량01
                    spdView(0).Col = 7: spdView(0).Value = RS01!수량02
                    spdView(0).Col = 8: spdView(0).Value = RS01!수량03
                    spdView(0).Col = 9: spdView(0).Value = RS01!수량04
                    spdView(0).Col = 10: spdView(0).Value = RS01!수량05
                    spdView(0).Col = 11: spdView(0).Value = RS01!수량06
                    spdView(0).Col = 13: spdView(0).Value = RS01!수량07
                    spdView(0).Col = 14: spdView(0).Value = RS01!수량08
                    spdView(0).Col = 15: spdView(0).Value = RS01!수량09
                    spdView(0).Col = 16: spdView(0).Value = RS01!수량10
                    spdView(0).Col = 17: spdView(0).Value = RS01!수량11
                    spdView(0).Col = 18: spdView(0).Value = RS01!수량12
                    
                    nTotal(0) = RS01!수량01 + RS01!수량02 + RS01!수량03 + RS01!수량04 + RS01!수량05 + RS01!수량06
                    nTotal(1) = RS01!수량07 + RS01!수량08 + RS01!수량09 + RS01!수량10 + RS01!수량11 + RS01!수량12
                    
                    spdView(0).Col = 4: spdView(0).Value = nTotal(0) + nTotal(1)
                    spdView(0).Col = 12: spdView(0).Value = nTotal(0)
                    spdView(0).Col = 19: spdView(0).Value = nTotal(1)
                                
                    spdView(0).Row = i + 1
                    spdView(0).Col = 5: spdView(0).Value = RS01!합계금액
                    spdView(0).Col = 6: spdView(0).Value = RS01!금액01
                    spdView(0).Col = 7: spdView(0).Value = RS01!금액02
                    spdView(0).Col = 8: spdView(0).Value = RS01!금액03
                    spdView(0).Col = 9: spdView(0).Value = RS01!금액04
                    spdView(0).Col = 10: spdView(0).Value = RS01!금액05
                    spdView(0).Col = 11: spdView(0).Value = RS01!금액06
                    spdView(0).Col = 13: spdView(0).Value = RS01!금액07
                    spdView(0).Col = 14: spdView(0).Value = RS01!금액08
                    spdView(0).Col = 15: spdView(0).Value = RS01!금액09
                    spdView(0).Col = 16: spdView(0).Value = RS01!금액10
                    spdView(0).Col = 17: spdView(0).Value = RS01!금액11
                    spdView(0).Col = 18: spdView(0).Value = RS01!금액12
                    
                    nTotal(0) = RS01!금액01 + RS01!금액02 + RS01!금액03 + RS01!금액04 + RS01!금액05 + RS01!금액06
                    nTotal(1) = RS01!금액07 + RS01!금액08 + RS01!금액09 + RS01!금액10 + RS01!금액11 + RS01!금액12
                    
                    spdView(0).Col = 4: spdView(0).Value = nTotal(0) + nTotal(1)
                    spdView(0).Col = 12: spdView(0).Value = nTotal(0)
                    spdView(0).Col = 19: spdView(0).Value = nTotal(1)
                    i = i + 200
                Else
                    i = i + 4
                End If
            Wend
        End If
        
        RS01.MoveNext
    Loop
        
    Call TotalQty
    Call TotalAmt
    
    ReDim sValue(0)
    
    sValue(0) = qYear
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_04007_02", sValue(), Err_Num, Err_Dec)
    
    Do While Not RS01.EOF
        i = 3
        Do While i < spRow
            spdView(0).Row = i
            spdView(0).Col = 1
            If spdView(0).Value = RS01!대리점코드 Then
                If Val(RS01!월) > 6 Then
                    spdView(0).Col = Val(RS01!월) + 6
                Else
                    spdView(0).Col = Val(RS01!월) + 5
                End If
                
                spdView(0).Row = i + 2
                spdView(0).Value = RS01!입고수량
                spdView(0).Row = i + 3
                spdView(0).Value = RS01!금액 / 1000
                i = i + 200
            Else
                i = i + 4
            End If
        Loop
        RS01.MoveNext
    Loop
    
    spdView(0).Row = 4
    While spdView(0).Row < spRow
        For i = 1 To 2
            spdView(0).Row = spdView(0).Row + 1
            nTotal(0) = 0
            nTotal(1) = 0
            
            For j = 6 To 19
                spdView(0).Col = j
                
                If j = 12 Or j = 19 Then
                    spdView(0).Value = nTotal(0)
                    nTotal(1) = nTotal(1) + nTotal(0)
                    spdView(0).Col = 4
                    spdView(0).Value = nTotal(0)
                    nTotal(1) = 0
                Else
                    nTotal(0) = nTotal(0) + Val(spdView(0).Value)
                End If
            Next j
        Next i
        
        spdView(0).Row = spdView(0).Row + 2
    Wend
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub TotalQty()
    Dim i As Integer
    Dim nTotal(14) As Long
    
    For i = 1 To 14
        nTotal(i) = 0
    Next i
    
    spdView(0).Row = 3
    Do While spdView(0).Row < spRow
        For i = 6 To 19
            spdView(0).Col = i
            nTotal(i - 5) = nTotal(i - 5) + Val(spdView(0).Value)
        Next i
        spdView(0).Row = spdView(0).Row + 4
    Loop
    
    spdView(1).Row = 1
    spdView(1).Col = 2
    spdView(1).Value = nTotal(7) + nTotal(14)
    
    For i = 1 To 14
        spdView(1).Col = i + 3
        spdView(1).Value = nTotal(i)
    Next i
End Sub

Private Sub TotalAmt()
    Dim i As Integer
    Dim nTotal(14) As Long
    
    For i = 1 To 14
        nTotal(i) = 0
    Next i
    
    spdView(0).Row = 4
    Do While spdView(0).Row < spRow
        For i = 6 To 19
            spdView(0).Col = i
            nTotal(i - 5) = nTotal(i - 5) + Val(spdView(0).Value)
        Next i
        spdView(0).Row = spdView(0).Row + 4
    Loop
        
    spdView(1).Row = 2
    spdView(1).Col = 2
    spdView(1).Value = nTotal(7) + nTotal(14)
    
    For i = 1 To 14
        spdView(1).Row = 2
        spdView(1).Col = i + 3
        spdView(1).Value = nTotal(i)
        spdView(1).Row = 3
        
        If nTotal(7) > 0 Or nTotal(14) > 0 Then
            spdView(1).Value = nTotal(i) / (nTotal(7) + nTotal(14)) * 100
        End If
    Next i
End Sub

Public Sub DataSave()
    Dim i As Integer
    Dim j As Integer
    
    ReDim sValue(0)
    
    sValue(0) = Format(dtInput.Value, "yyyy")
    
    Call ExecPro("SP_04007_04", sValue(), Err_Num, Err_Dec)
    
    
    ReDim sValue(27)
    
    sValue(0) = Format(dtInput.Value, "yyyy")
    sValue(1) = "QQQ"
    
    spdView(0).Row = 1
    For j = 6 To 18
        spdView(0).Col = j
        If j < 12 Then
            sValue(j - 4) = Val(spdView(0).Value)
        ElseIf j > 12 Then
            sValue(j - 5) = Val(spdView(0).Value)
        End If
    Next j
    
    spdView(0).Row = 2
    For j = 6 To 18
        spdView(0).Col = j
        If j < 12 Then
            sValue(j + 8) = Val(spdView(0).Value)
        ElseIf j > 12 Then
            sValue(j + 7) = Val(spdView(0).Value)
        End If
    Next j

    sValue(26) = 0
    sValue(27) = 0
    
    Call ExecPro("SP_04007_03", sValue(), Err_Num, Err_Dec)
    
    spdView(0).Row = 3
    sValue(0) = Format(dtInput.Value, "yyyy")
    
    Do While spdView(0).Row < spRow
        spdView(0).Col = 1
        sValue(1) = spdView(0).Value
        For j = 6 To 18
            spdView(0).Col = j
            If j < 12 Then
                sValue(j - 4) = Val(spdView(0).Value)
            ElseIf j > 12 Then
                sValue(j - 5) = Val(spdView(0).Value)
            End If
        Next j
        
        spdView(0).Col = 5
        sValue(26) = Val(spdView(0).Value)
        spdView(0).Row = spdView(0).Row + 1
        For j = 6 To 18
            spdView(0).Col = j
            If j < 12 Then
                sValue(j + 8) = Val(spdView(0).Value)
            ElseIf j > 12 Then
                sValue(j + 7) = Val(spdView(0).Value)
            End If
        Next j
        
        spdView(0).Col = 5
        sValue(27) = Val(spdView(0).Value)
        
        Call ExecPro("SP_04007_03", sValue(), Err_Num, Err_Dec)
        
        spdView(0).Row = spdView(0).Row + 3
    Loop
End Sub

Public Sub DataDelete()
    ReDim sValue(0)
    
    sValue(0) = Format(dtInput.Value, "yyyy")
    
    Call ExecPro("SP_04007_04", sValue(), Err_Num, Err_Dec)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_04007_Flag = False
End Sub

Private Sub spdView_Change(Index As Integer, ByVal Col As Long, ByVal Row As Long)
    If Index = 0 And Col = 5 Then
        Dim i As Integer
        Dim l As Long
        Dim ll As Long
        Dim lTotal(14) As Long
        
        spdView(0).Row = Row
        spdView(0).Col = 5
        ll = spdView(0).Value
        
        For i = 4 To 20
            If i <> 5 Then
                spdView(0).Col = i
                spdView(0).Row = Row + 2
                
                If spdView(0).Value <> "" Then
                    l = spdView(0).Value
                    
                    spdView(0).Row = Row
                    spdView(0).Value = l + l * (ll / 100)
                End If
            End If
        Next i
        
        If Row Mod 2 = 1 Then
            For i = 3 To spdView(0).MaxRows Step 4
                spdView(0).Row = i
                spdView(0).Col = 4
                If spdView(0).Value <> "" Then lTotal(0) = lTotal(0) + spdView(0).Value
                spdView(0).Col = 6
                If spdView(0).Value <> "" Then lTotal(1) = lTotal(1) + spdView(0).Value
                spdView(0).Col = 7
                If spdView(0).Value <> "" Then lTotal(2) = lTotal(2) + spdView(0).Value
                spdView(0).Col = 8
                If spdView(0).Value <> "" Then lTotal(3) = lTotal(3) + spdView(0).Value
                spdView(0).Col = 9
                If spdView(0).Value <> "" Then lTotal(4) = lTotal(4) + spdView(0).Value
                spdView(0).Col = 10
                If spdView(0).Value <> "" Then lTotal(5) = lTotal(5) + spdView(0).Value
                spdView(0).Col = 11
                If spdView(0).Value <> "" Then lTotal(6) = lTotal(6) + spdView(0).Value
                spdView(0).Col = 12
                If spdView(0).Value <> "" Then lTotal(7) = lTotal(7) + spdView(0).Value
                spdView(0).Col = 13
                If spdView(0).Value <> "" Then lTotal(8) = lTotal(8) + spdView(0).Value
                spdView(0).Col = 14
                If spdView(0).Value <> "" Then lTotal(9) = lTotal(9) + spdView(0).Value
                spdView(0).Col = 15
                If spdView(0).Value <> "" Then lTotal(10) = lTotal(10) + spdView(0).Value
                spdView(0).Col = 16
                If spdView(0).Value <> "" Then lTotal(11) = lTotal(11) + spdView(0).Value
                spdView(0).Col = 17
                If spdView(0).Value <> "" Then lTotal(12) = lTotal(12) + spdView(0).Value
                spdView(0).Col = 18
                If spdView(0).Value <> "" Then lTotal(13) = lTotal(13) + spdView(0).Value
                spdView(0).Col = 19
                If spdView(0).Value <> "" Then lTotal(14) = lTotal(14) + spdView(0).Value
            Next i
            
            spdView(1).Row = 1
            spdView(1).Col = 2
            spdView(1).Value = lTotal(0)
            spdView(1).Col = 4
            spdView(1).Value = lTotal(1)
            spdView(1).Col = 5
            spdView(1).Value = lTotal(2)
            spdView(1).Col = 6
            spdView(1).Value = lTotal(3)
            spdView(1).Col = 7
            spdView(1).Value = lTotal(4)
            spdView(1).Col = 8
            spdView(1).Value = lTotal(5)
            spdView(1).Col = 9
            spdView(1).Value = lTotal(6)
            spdView(1).Col = 10
            spdView(1).Value = lTotal(7)
            spdView(1).Col = 11
            spdView(1).Value = lTotal(8)
            spdView(1).Col = 12
            spdView(1).Value = lTotal(9)
            spdView(1).Col = 13
            spdView(1).Value = lTotal(10)
            spdView(1).Col = 14
            spdView(1).Value = lTotal(11)
            spdView(1).Col = 15
            spdView(1).Value = lTotal(12)
            spdView(1).Col = 16
            spdView(1).Value = lTotal(13)
            spdView(1).Col = 17
            spdView(1).Value = lTotal(14)
        Else
            For i = 4 To spdView(0).MaxRows Step 4
                spdView(0).Row = i
                spdView(0).Col = 4
                If spdView(0).Value <> "" Then lTotal(0) = lTotal(0) + spdView(0).Value
                spdView(0).Col = 6
                If spdView(0).Value <> "" Then lTotal(1) = lTotal(1) + spdView(0).Value
                spdView(0).Col = 7
                If spdView(0).Value <> "" Then lTotal(2) = lTotal(2) + spdView(0).Value
                spdView(0).Col = 8
                If spdView(0).Value <> "" Then lTotal(3) = lTotal(3) + spdView(0).Value
                spdView(0).Col = 9
                If spdView(0).Value <> "" Then lTotal(4) = lTotal(4) + spdView(0).Value
                spdView(0).Col = 10
                If spdView(0).Value <> "" Then lTotal(5) = lTotal(5) + spdView(0).Value
                spdView(0).Col = 11
                If spdView(0).Value <> "" Then lTotal(6) = lTotal(6) + spdView(0).Value
                spdView(0).Col = 12
                If spdView(0).Value <> "" Then lTotal(7) = lTotal(7) + spdView(0).Value
                spdView(0).Col = 13
                If spdView(0).Value <> "" Then lTotal(8) = lTotal(8) + spdView(0).Value
                spdView(0).Col = 14
                If spdView(0).Value <> "" Then lTotal(9) = lTotal(9) + spdView(0).Value
                spdView(0).Col = 15
                If spdView(0).Value <> "" Then lTotal(10) = lTotal(10) + spdView(0).Value
                spdView(0).Col = 16
                If spdView(0).Value <> "" Then lTotal(11) = lTotal(11) + spdView(0).Value
                spdView(0).Col = 17
                If spdView(0).Value <> "" Then lTotal(12) = lTotal(12) + spdView(0).Value
                spdView(0).Col = 18
                If spdView(0).Value <> "" Then lTotal(13) = lTotal(13) + spdView(0).Value
                spdView(0).Col = 19
                If spdView(0).Value <> "" Then lTotal(14) = lTotal(14) + spdView(0).Value
            Next i

            spdView(1).Row = 2
            spdView(1).Col = 2
            spdView(1).Value = lTotal(0)
            spdView(1).Col = 4
            spdView(1).Value = lTotal(1)
            spdView(1).Col = 5
            spdView(1).Value = lTotal(2)
            spdView(1).Col = 6
            spdView(1).Value = lTotal(3)
            spdView(1).Col = 7
            spdView(1).Value = lTotal(4)
            spdView(1).Col = 8
            spdView(1).Value = lTotal(5)
            spdView(1).Col = 9
            spdView(1).Value = lTotal(6)
            spdView(1).Col = 10
            spdView(1).Value = lTotal(7)
            spdView(1).Col = 11
            spdView(1).Value = lTotal(8)
            spdView(1).Col = 12
            spdView(1).Value = lTotal(9)
            spdView(1).Col = 13
            spdView(1).Value = lTotal(10)
            spdView(1).Col = 14
            spdView(1).Value = lTotal(11)
            spdView(1).Col = 15
            spdView(1).Value = lTotal(12)
            spdView(1).Col = 16
            spdView(1).Value = lTotal(13)
            spdView(1).Col = 17
            spdView(1).Value = lTotal(14)
        End If
    End If
End Sub

Private Sub spdView_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        'PopupMenu P_00000.PopUp
    End If
End Sub

Private Sub spdView_TopLeftChange(Index As Integer, ByVal OldLeft As Long, ByVal OldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)
    If Index = 0 Then
        spdView(1).LeftCol = NewLeft - 2
    ElseIf Index = 1 Then
        spdView(0).LeftCol = NewLeft + 2
    End If
End Sub

Public Sub DataPrint()
'    Dim ReportFP As String
'    Dim ReportFile As String
'    Dim TempText As String
'
'    ReportFP = GetIniStr("REPORT", "FilePath", "", m_iniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    Call PrintDesc
'
'    Dim i As Integer
'    For i = 0 To 30
'        P_00000.crPrint.Formulas(i) = ""
'    Next
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'
'    spdView(0).Row = 1
'    spdView(0).Col = 6
'    TempText = RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 7
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 8
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 9
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 10
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 11
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 12
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 13
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 14
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 15
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 16
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 17
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 18
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 19
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'
'    P_00000.crPrint.Formulas(0) = "영업일수 = '" & TempText & "'"
'
'    spdView(0).Row = 2
'    spdView(0).Col = 6
'    TempText = RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 7
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 8
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 9
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 10
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 11
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 12
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 13
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 14
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 15
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 16
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 17
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 18
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 19
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'
'    P_00000.crPrint.Formulas(1) = "매장수 = '" & TempText & "'"
'
'    spdView(1).Row = 1
'    spdView(1).Col = 4
'    TempText = RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 5
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 6
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 7
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 8
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 9
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 10
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 11
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 12
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 13
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 14
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 15
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 16
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 17
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'
'    P_00000.crPrint.Formulas(2) = "수량계 = '" & TempText & "'"
'
'    spdView(1).Row = 2
'    spdView(1).Col = 4
'    TempText = RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 5
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 6
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 7
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 8
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 9
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 10
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 11
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 12
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 13
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 14
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 15
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 16
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 17
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'
'    P_00000.crPrint.Formulas(3) = "총계 = '" & TempText & "'"
'
'    spdView(1).Row = 3
'    spdView(1).Col = 4
'    TempText = RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 5
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 6
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 7
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 8
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 9
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 10
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 11
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 12
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 13
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 14
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 15
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 16
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 17
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'
'    P_00000.crPrint.Formulas(4) = "평균 = '" & TempText & "'"
'
'    P_00000.crPrint.Formulas(5) = "연도 = '" & Format(dtInput.Value, "yyyy-mm") & "'"
'
'    If optSelect(0).Value = True Then
'        P_00000.crPrint.Formulas(6) = "구분 = '전  체'"
'    ElseIf optSelect(1).Value = True Then
'        P_00000.crPrint.Formulas(6) = "구분 = '대리점'"
'    ElseIf optSelect(2).Value = True Then
'        P_00000.crPrint.Formulas(6) = "구분 = '백화점'"
'    ElseIf optSelect(3).Value = True Then
'        P_00000.crPrint.Formulas(6) = "구분 = '할인매장'"
'    End If
'
'    Call ReportPrint(ReportFile, "1")
End Sub

Public Sub DataScreen()
'    Dim ReportFP As String
'    Dim ReportFile As String
'    Dim TempText As String
'
'    ReportFP = GetIniStr("REPORT", "FilePath", "", m_iniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    Call PrintDesc
'
'    Dim i As Integer
'    For i = 0 To 30
'        P_00000.crPrint.Formulas(i) = ""
'    Next
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'
'    spdView(0).Row = 1
'    spdView(0).Col = 6
'    TempText = RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 7
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 8
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 9
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 10
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 11
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 12
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 13
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 14
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 15
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 16
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 17
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 18
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 19
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'
'    P_00000.crPrint.Formulas(0) = "영업일수 = '" & TempText & "'"
'
'    spdView(0).Row = 2
'    spdView(0).Col = 6
'    TempText = RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 7
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 8
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 9
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 10
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 11
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 12
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 13
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 14
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 15
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 16
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 17
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 18
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'    spdView(0).Col = 19
'    TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
'
'    P_00000.crPrint.Formulas(1) = "매장수 = '" & TempText & "'"
'
'    spdView(1).Row = 1
'    spdView(1).Col = 4
'    TempText = RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 5
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 6
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 7
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 8
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 9
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 10
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 11
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 12
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 13
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 14
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 15
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 16
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 17
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'
'    P_00000.crPrint.Formulas(2) = "수량계 = '" & TempText & "'"
'
'    spdView(1).Row = 2
'    spdView(1).Col = 4
'    TempText = RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 5
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 6
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 7
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 8
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 9
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 10
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 11
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 12
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 13
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 14
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 15
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 16
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 17
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'
'    P_00000.crPrint.Formulas(3) = "총계 = '" & TempText & "'"
'
'    spdView(1).Row = 3
'    spdView(1).Col = 4
'    TempText = RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 5
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 6
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 7
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 8
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 9
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 10
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 11
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 12
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 13
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 14
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 15
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 16
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'    spdView(1).Col = 17
'    TempText = TempText & RightH(Space(8) & spdView(1).Text, 8)
'
'    P_00000.crPrint.Formulas(4) = "평균 = '" & TempText & "'"
'
'    P_00000.crPrint.Formulas(5) = "연도 = '" & Format(dtInput.Value, "yyyy-mm") & "'"
'
'    If optSelect(0).Value = True Then
'        P_00000.crPrint.Formulas(6) = "구분 = '전  체'"
'    ElseIf optSelect(1).Value = True Then
'        P_00000.crPrint.Formulas(6) = "구분 = '대리점'"
'    ElseIf optSelect(2).Value = True Then
'        P_00000.crPrint.Formulas(6) = "구분 = '백화점'"
'    ElseIf optSelect(3).Value = True Then
'        P_00000.crPrint.Formulas(6) = "구분 = '할인매장'"
'    End If
'
'    Call ReportPrint(ReportFile, "2")
End Sub

Private Sub PrintDesc()
    Dim i As Integer
    Dim TempText As String
    Dim TempFP As String
    Dim TempFile As String
    
    TempFP = GetIniStr("REPORT", "TempPath", "", m_iniFile)
    TempFile = TempFP & "\Temp.txt"
    
    Open TempFile For Output As #1
    
    TempText = ""
    
    For i = 3 To spdView(0).MaxRows - 1
        spdView(0).Row = i
        
        spdView(0).Col = 3
        If spdView(0).Text = "" Then
            Exit For
        End If
        
        spdView(0).Col = 1
        TempText = LeftH(spdView(0).Text & Space(16), 16)
        spdView(0).Col = 2
        TempText = TempText & LeftH(spdView(0).Text & Space(8), 8)
        spdView(0).Col = 3
        TempText = TempText & LeftH(spdView(0).Text & Space(8), 8)
        spdView(0).Col = 4
        TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
        spdView(0).Col = 5
        TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
        spdView(0).Col = 6
        TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
        spdView(0).Col = 7
        TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
        spdView(0).Col = 8
        TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
        spdView(0).Col = 9
        TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
        spdView(0).Col = 10
        TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
        spdView(0).Col = 11
        TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
        spdView(0).Col = 12
        TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
        spdView(0).Col = 13
        TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
        spdView(0).Col = 14
        TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
        spdView(0).Col = 15
        TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
        spdView(0).Col = 16
        TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
        spdView(0).Col = 17
        TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
        spdView(0).Col = 18
        TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
        spdView(0).Col = 19
        TempText = TempText & RightH(Space(8) & spdView(0).Text, 8)
        
        Print #1, TempText
    Next i
    
    Close #1
End Sub

