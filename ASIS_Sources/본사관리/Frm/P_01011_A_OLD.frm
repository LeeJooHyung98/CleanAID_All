VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_01011_A_OLD 
   Caption         =   "[전사업장]가맹점 품목등록"
   ClientHeight    =   12240
   ClientLeft      =   555
   ClientTop       =   3285
   ClientWidth     =   16020
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
   ScaleHeight     =   12240
   ScaleWidth      =   16020
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   12240
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16020
      _ExtentX        =   28258
      _ExtentY        =   21590
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_01011_A_OLD.frx":0000
      Begin Threed.SSPanel SSPanel1 
         Height          =   10890
         Left            =   7470
         TabIndex        =   4
         Top             =   1335
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   19209
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdSubBtn 
            Height          =   630
            Index           =   0
            Left            =   150
            TabIndex        =   6
            Top             =   885
            Width           =   780
            _Version        =   851970
            _ExtentX        =   1376
            _ExtentY        =   1111
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "P_01011_A_OLD.frx":0112
         End
         Begin XtremeSuiteControls.PushButton cmdSubBtn 
            Height          =   630
            Index           =   1
            Left            =   150
            TabIndex        =   7
            Top             =   1680
            Width           =   780
            _Version        =   851970
            _ExtentX        =   1376
            _ExtentY        =   1111
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "P_01011_A_OLD.frx":04AC
         End
         Begin XtremeSuiteControls.PushButton cmdSubBtn 
            Height          =   630
            Index           =   2
            Left            =   150
            TabIndex        =   8
            Top             =   2475
            Width           =   780
            _Version        =   851970
            _ExtentX        =   1376
            _ExtentY        =   1111
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "P_01011_A_OLD.frx":0846
         End
         Begin XtremeSuiteControls.PushButton cmdSubBtn 
            Height          =   630
            Index           =   3
            Left            =   150
            TabIndex        =   9
            Top             =   3270
            Width           =   780
            _Version        =   851970
            _ExtentX        =   1376
            _ExtentY        =   1111
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "P_01011_A_OLD.frx":0BE0
         End
      End
      Begin Threed.SSPanel panCaption 
         Height          =   390
         Index           =   1
         Left            =   15
         TabIndex        =   1
         Top             =   1335
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   688
         _Version        =   262144
         Font3D          =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "본사 품목 코드"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_01011_A_OLD.frx":0F7A
         BevelOuter      =   0
         RoundedCorners  =   0   'False
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10485
         Index           =   0
         Left            =   15
         TabIndex        =   2
         Top             =   1740
         Width           =   7440
         _Version        =   524288
         _ExtentX        =   13123
         _ExtentY        =   18494
         _StockProps     =   64
         BackColorStyle  =   1
         EditEnterAction =   5
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
         SpreadDesigner  =   "P_01011_A_OLD.frx":13DC
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panCaption 
         Height          =   390
         Index           =   2
         Left            =   8565
         TabIndex        =   3
         Top             =   1335
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   688
         _Version        =   262144
         Font3D          =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "  가맹점 품목 코드"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_01011_A_OLD.frx":1844
         BevelOuter      =   0
         RoundedCorners  =   0   'False
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10485
         Index           =   1
         Left            =   8565
         TabIndex        =   5
         Top             =   1740
         Width           =   7440
         _Version        =   524288
         _ExtentX        =   13123
         _ExtentY        =   18494
         _StockProps     =   64
         BackColorStyle  =   1
         EditEnterAction =   5
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
         SpreadDesigner  =   "P_01011_A_OLD.frx":1CA6
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   10
         Top             =   540
         Width           =   15990
         _ExtentX        =   28205
         _ExtentY        =   1376
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
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   1
            Left            =   1530
            Style           =   2  '드롭다운 목록
            TabIndex        =   12
            Top             =   420
            Width           =   3015
         End
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            Style           =   2  '드롭다운 목록
            TabIndex        =   11
            Top             =   60
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Left            =   6720
            TabIndex        =   13
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            DateIsNull      =   -1  'True
            Format          =   21430272
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   14
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
            Caption         =   "사 업 장"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   25
            Left            =   60
            TabIndex        =   15
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
            Caption         =   "가 맹 점"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   3
            Left            =   5250
            TabIndex        =   16
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
            Caption         =   "적용일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   17
         Top             =   15
         Width           =   8385
         _ExtentX        =   14790
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
         Caption         =   " #"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_01011_A_OLD.frx":210E
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Left            =   8415
         TabIndex        =   18
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
         PictureBackground=   "P_01011_A_OLD.frx":2310
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   19
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
            Picture         =   "P_01011_A_OLD.frx":2512
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   20
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
            Picture         =   "P_01011_A_OLD.frx":2AAC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   21
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
            Picture         =   "P_01011_A_OLD.frx":3046
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   22
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
            Picture         =   "P_01011_A_OLD.frx":35E0
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   23
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
            Picture         =   "P_01011_A_OLD.frx":3B7A
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   24
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
            Picture         =   "P_01011_A_OLD.frx":4114
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   25
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
            Picture         =   "P_01011_A_OLD.frx":46AE
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   26
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
            Picture         =   "P_01011_A_OLD.frx":4C48
         End
      End
   End
End
Attribute VB_Name = "P_01011_A_OLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cboInput_Click(Index As Integer)
    Dim sCode As String

    If Index = 0 Then
        sCode = Trim(Mid(Trim(cboInput(0)) & Space(10), 2, 4))

        Call StoreComboAdd(cboInput(1), sCode)
    End If
    
    Call Data_Display
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: Call Data_Display   ' 조회
        Case 1: 'Call DataAdd        ' 신규
        Case 2: Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
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

Private Sub cmdSubBtn_Click(Index As Integer)
    Dim I As Integer
    Dim j As Integer
    Dim bMove As Boolean

    If Index = 0 Then   ' >>
        spdView(1).MaxRows = 0

        If Not IsDate(dtInput.Value) Then
            MsgBox " 적용일자가 선택되지 않았읍니다.", vbInformation, "확인"
            Exit Sub
        ElseIf dtInput.Value < Date Then
            MsgBox " 적용일자를 확인하여 주십시요.", vbInformation, "오류"
            Exit Sub
        End If
        
        If Trim(Mid(cboInput(1).Text, 2, 6)) = "" Then
            MsgBox " 가맹점이 선택되지 않았읍니다.", vbInformation, "확인"
            Exit Sub
        End If

        sValue(0) = Mid(cboInput(1).Text, 2, 6)
        sValue(1) = Format(dtInput.Value, "YYYY-MM-DD")

        Query = "SELECT  isnull(MAX(ISNULL(SDate, CONVERT(CHAR(10),GETDATE(),112))),'') MAXSDATE "
        Query = Query + " From  StoreGoodsCT (nolock)"
        Query = Query + " WHERE STORE_CD = '" & sValue(0) & "'"
        Query = Query + "   AND SDate   >= '" & sValue(1) & "'"

        Set RS01 = New ADODB.Recordset
        Call SqlDataValue(RS01, Query)

        If Not (Trim(RS01!MAXSDATE) = "") Then
            MsgBox "최종일자:" & RS01!MAXSDATE & "일자 이후로 적용일자를 변경후 작업 하세요.."
            Exit Sub
        End If

        For I = 1 To spdView(0).MaxRows
            spdView(0).Row = I

            spdView(1).MaxRows = spdView(1).MaxRows + 1
            spdView(1).Row = spdView(1).MaxRows

            spdView(0).Col = 1
            spdView(1).Col = 1:
            spdView(1).Text = spdView(0).Text

            spdView(0).Col = 2
            spdView(1).Col = 2
            spdView(1).Text = spdView(0).Text

            spdView(0).Col = 4
            If spdView(0).Text = 0 Then
                spdView(0).Col = 3
                spdView(1).Col = 3
                spdView(1).Text = spdView(0).Text
            Else
                spdView(1).Col = 3
                spdView(1).Text = spdView(0).Text
            End If

            spdView(0).Col = 5
            spdView(1).Col = 4
            spdView(1).Text = spdView(0).Text

            spdView(0).Col = 6
            spdView(1).Col = 5
            spdView(1).Text = spdView(0).Text

            DoEvents
        Next I
        
        cmdSubBtn(Index).Enabled = False
        cmdBtn(2).Enabled = True        '저장
    End If
End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True        '조회
    cmdBtn(1).Enabled = False       '신규
    cmdBtn(2).Enabled = False        '저장
    cmdBtn(3).Enabled = False       '삭제
    cmdBtn(4).Enabled = False       '취소
    cmdBtn(5).Enabled = False       '인쇄
    cmdBtn(6).Enabled = False       'Screen
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
End Sub

''Private Sub spdDisplay(RS As ADODB.Recordset)
''
''    Call fpSpread_Display(spdView(0), RS)
''
''
''
'''    Set spdView(0).DataSource = Nothing
''End Sub
''
''Private Sub spdDisplay2(RS As ADODB.Recordset)
''
''    Call fpSpread_Display(spdView(1), RS)
''
''    'spdView(1).ColHidden = True
''
'''    Set spdView(1).DataSource = Nothing
''End Sub

Private Sub Form_Load()
    With spdView(0)
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
            
            
        .ColsFrozen = 1 '틀고정
        .Row = -1
        
        .Col = 1
        .ColWidth(1) = 7
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 2
        .ColWidth(2) = 18
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
        
        .Col = 3
        .ColWidth(3) = 7
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 0
        .TypeVAlign = TypeVAlignCenter
        
        .Col = 4
        .ColWidth(4) = 7
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 0
        .TypeVAlign = TypeVAlignCenter
        
        .Col = 5
        .ColWidth(5) = 7
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 6
        .ColWidth(6) = 7
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
    End With
    
    With spdView(1)
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
            
            
        .ColsFrozen = 1 '틀고정
        .Row = -1
        
        .Col = 1
        .ColWidth(1) = 7
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 2
        .ColWidth(2) = 18
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
        
        .Col = 3
        .ColWidth(3) = 7
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 0
        .TypeVAlign = TypeVAlignCenter
        
        .Col = 4
        .ColWidth(4) = 7
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 5
        .ColWidth(5) = 7
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
    End With

    If P_01011_A_Flag = False Then
        Call Master_tblComboAdd(cboInput(0))

        cboInput(0).ListIndex = 1
        
        ReDim sValue(1)


        sValue(0) = "1"
        sValue(1) = Mid(cboInput(1).Text, 2, 6)

        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_01011_A_00", sValue(), Err_Num, Err_Dec)

        spdView(0).MaxCols = RS01.Fields.Count
        spdView(0).MaxRows = RS01.RecordCount

        'Call spdDisplay(RS01)
        Call fpSpread_Display(spdView(0), RS01)
        Call GetColWidth(REG_App, Me.Name & "A", spdView(0))
    
        ReDim sValue(1)
        
        sValue(0) = "1"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_01011_A_01", sValue(), Err_Num, Err_Dec)
        
        spdView(1).MaxCols = RS01.Fields.Count
        spdView(1).MaxRows = RS01.RecordCount
        
        'Call spdDisplay2(RS01)
        Call fpSpread_Display(spdView(1), RS01)
        Call GetColWidth(REG_App, Me.Name & "B", spdView(1))
        
        dtInput.Value = Now
        
        'DTPicker1.Value = Now
        
        P_01011_A_Flag = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_01011_A_Flag = False
End Sub
'
'Public Sub Data_Display()
'
'    If Not IsDate(dtInput.Value) Then
'        MsgBox " 적용일자가 선택되지 않았읍니다.", vbInformation, "확인"
'        Exit Sub
'
'    ElseIf dtInput.Value < Date Then
'        MsgBox " 적용일자를 확인하여 주십시요.", vbInformation, "오류"
'        Exit Sub
'    End If
'    If Trim(Mid(cboInput(1).Text, 2, 6)) = "" Then
'        MsgBox " 가맹점이 선택되지 않았읍니다.", vbInformation, "확인"
'        Exit Sub
'    End If
'
'
'    ReDim sValue(1)
'
'
'    sValue(0) = "0"
'    sValue(1) = Mid(cboInput(1).Text, 2, 6)
'
'    Set RS01 = New ADODB.Recordset
'    Set RS01 = ExecPro("SP_01011_A_00", sValue(), Err_Num, Err_Dec)
'
'    spdView(0).MaxCols = RS01.Fields.Count
'    spdView(0).MaxRows = RS01.RecordCount
'
'    Call spdDisplay(RS01)
'    Call GetColWidth(REG_App, Me.Name & "A", spdView(0))
'
'    spdView(1).MaxRows = 0
'
''    ReDim sValue(1)
'
''    sValue(0) = "0"
''    sValue(1) = Mid(cboInput(1).Text, 2, 6)
''
''    If Len(sValue(1)) > 0 Then
''        Set RS01 = New ADODB.Recordset
''        Set RS01 = ExecPro("SP_01011_A_01", sValue(), Err_Num, Err_Dec)
''
''        spdView(1).MaxCols = RS01.Fields.Count
''        spdView(1).MaxRows = RS01.RecordCount
''
''        Call spdDisplay2(RS01)
''        Call GetColWidth(REG_App, Me.Name & "A", spdView(1))
''
'''        If Not RS01.EOF Then
'''            dtInput.Value = Format(RS01!적용일자, "####-##-##")
'''        End If
''    Else
''        spdView(1).MaxRows = 0
''    End If
'
'    'cmdSubBtn(2).Enabled = True
'    'optSelect(0).Enabled = True
'    'optSelect(1).Enabled = True
'End Sub

Public Sub Data_Display()
    If Not IsDate(dtInput.Value) Then
        MsgBox " 적용일자가 선택되지 않았읍니다.", vbInformation, "확인"
        Exit Sub

    ElseIf dtInput.Value < Date Then
        MsgBox " 적용일자를 확인하여 주십시요.", vbInformation, "오류"
        Exit Sub
    End If
    
    If Trim(Mid(cboInput(1).Text, 2, 6)) = "" Then
        MsgBox " 가맹점이 선택되지 않았읍니다.", vbInformation, "확인"
        Exit Sub
    End If

    ReDim sValue(1)


    sValue(0) = "0"
    sValue(1) = Mid(cboInput(1).Text, 2, 6)

    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01011_A_00", sValue(), Err_Num, Err_Dec)

    spdView(0).MaxCols = RS01.Fields.Count
    spdView(0).MaxRows = RS01.RecordCount

    'Call spdDisplay(RS01)
    Call fpSpread_Display(spdView(0), RS01)
    Call GetColWidth(REG_App, Me.Name & "A", spdView(0))

    spdView(1).MaxRows = 0

    ReDim sValue(1)

    sValue(0) = "0"
    sValue(1) = Mid(cboInput(1).Text, 2, 6)

    If Len(sValue(1)) > 0 Then
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_01011_A_01", sValue(), Err_Num, Err_Dec)

        spdView(1).MaxCols = RS01.Fields.Count
        spdView(1).MaxRows = RS01.RecordCount

        'Call spdDisplay2(RS01)
        Call fpSpread_Display(spdView(1), RS01)
        Call GetColWidth(REG_App, Me.Name & "A", spdView(1))

'        If Not RS01.EOF Then
'            dtInput.Value = Format(RS01!적용일자, "####-##-##")
'        End If
    Else
        spdView(1).MaxRows = 0
    End If

    cmdBtn(2).Enabled = True        '저장
End Sub

Private Sub optSelect_Click(Index As Integer, Value As Integer)
    Select Case Index
        Case 0
            MsgBox " 대리점 : 적용일자변경시 대리점 품목코드에 등록된 모든 내용이 해당일자로 복사됨" & _
                              CStr(vbLf) & CStr(vbLf) & "(가격 변경이 적용됨).", vbInformation, "확인"
        Case 1
            MsgBox " 본  사 : 적용일자변경시 대리점품목코드에 등록된 품목코드를 기준으로 하여 본사품목이 복사됨" & _
                              CStr(vbLf) & CStr(vbLf) & "(본사 가격이 적용됨).", vbInformation, "확인"
    End Select
End Sub

'Private Sub spdView_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
'    spdView(Index).Row = Row
'    spdView(Index).Col = -1
'
'    If spdView(Index).BackColor = vbWhite Then
'        spdView(Index).BackColor = vbYellow
'    ElseIf spdView(Index).BackColor = vbYellow Then
'        spdView(Index).BackColor = vbWhite
'    End If
'End Sub

'Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'    With spdView
'        If NewRow <> -1 Then
'            .Row = Row
'            If (Row Mod 2) = 0 Then
'                .Col = -1
'                .BackColor = glbGray
'            Else
'                .Col = -1
'                .BackColor = vbWhite
'            End If
'
'            .Row = NewRow
'            .Col = -1
'            .BackColor = glbYellow
'        End If
'    End With
'End Sub

Public Sub DataSave()
    Dim I As Integer
    
    ReDim sValue(6)
    
    If Not IsDate(dtInput.Value) Then
        MsgBox " 적용일자가 선택되지 않았읍니다.", vbInformation, "확인"
        Exit Sub
        
    ElseIf dtInput.Value < Date Then
        MsgBox " 적용일자를 확인하여 주십시요.", vbInformation, "오류"
        Exit Sub
    End If
    If Trim(Mid(cboInput(1).Text, 2, 6)) = "" Then
        MsgBox " 가맹점이 선택되지 않았읍니다.", vbInformation, "확인"
        Exit Sub
    End If
    
    
    sValue(0) = Mid(cboInput(1).Text, 2, 6)
    sValue(1) = Format(dtInput.Value, "YYYY-MM-DD")
    
'    strSql = "          DELETE  AgencyGoodsCT WHERE AgencyCode = '" & sValue(0) & "'"
'    strSql = strSql + "                AND SDate = '" & sValue(1) & "'"
'    Set RS01 = New ADODB.Recordset
'    Call SqlDataValue(RS01, strSql)
    
    If spdView(1).MaxRows > 0 Then
        For I = 1 To spdView(1).MaxRows
            sValue(0) = Mid(cboInput(1).Text, 2, 6)
            sValue(1) = Format(dtInput.Value, "YYYY-MM-DD")
            
            spdView(1).Row = I
            spdView(1).Col = 1: sValue(2) = spdView(1).Text & ""
            spdView(1).Col = 2: sValue(3) = spdView(1).Text & ""
            spdView(1).Col = 3: sValue(4) = spdView(1).Value & ""
            spdView(1).Col = 4: sValue(5) = spdView(1).Value & ""
            spdView(1).Col = 5: sValue(6) = spdView(1).Value & ""
                          
            Call ExecPro("SP_01011_A_02", sValue(), Err_Num, Err_Dec)
        Next I
        
        MsgBox " 저장이 완료 되었습니다.", vbInformation, "확인"
        cmdBtn(2).Enabled = False        '저장
    Else
        MsgBox " 저장 할 Data가 없습니다.", vbInformation, "확인"
        cmdBtn(2).Enabled = False        '저장
        cmdSubBtn(0).Enabled = True
    End If
End Sub

Private Sub spdView_LeaveCell(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    With spdView(Index)
        If NewRow <> -1 Then
            .Row = Row
            If (Row Mod 2) = 0 Then
                .Col = -1
                .BackColor = glbGray
            Else
                .Col = -1
                .BackColor = vbWhite
            End If

            .Row = NewRow
            .Col = -1
            .BackColor = glbYellow
        End If
    End With
End Sub
