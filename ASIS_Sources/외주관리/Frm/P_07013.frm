VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_07013 
   Caption         =   "외주 출고 등록"
   ClientHeight    =   11640
   ClientLeft      =   5160
   ClientTop       =   6120
   ClientWidth     =   18645
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_07013.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11640
   ScaleWidth      =   18645
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   11640
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18645
      _ExtentX        =   32888
      _ExtentY        =   20532
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_07013.frx":058A
      Begin Threed.SSPanel SSPanel 
         Height          =   570
         Index           =   1
         Left            =   5310
         TabIndex        =   1
         Top             =   1740
         Width           =   13320
         _ExtentX        =   23495
         _ExtentY        =   1005
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   390
            Index           =   1
            Left            =   1800
            TabIndex        =   19
            Top             =   90
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   688
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
            CalendarBackColor=   16777215
            Format          =   56557568
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   90
            TabIndex        =   20
            Top             =   120
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   556
            _Version        =   262144
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "지사 출고 처리일자:"
            BorderWidth     =   0
            BevelOuter      =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   345
            Index           =   1
            Left            =   4620
            TabIndex        =   24
            Top             =   90
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   262144
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   2
         Top             =   540
         Width           =   18615
         _ExtentX        =   32835
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.TextBox txtOutClothCode 
            Height          =   315
            Left            =   5490
            TabIndex        =   29
            Text            =   "a0,i0,n0,o0,p0,w0,x0"
            ToolTipText     =   "a0,i0,n0,o0,p0,w0,x0"
            Top             =   420
            Width           =   2565
         End
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1245
            Style           =   2  '드롭다운 목록
            TabIndex        =   3
            Top             =   60
            Width           =   2850
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1245
            TabIndex        =   4
            Top             =   405
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            Format          =   56557568
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   5
            Top             =   405
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "스캔일자"
            BorderWidth     =   0
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   7
            Left            =   60
            TabIndex        =   6
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "지사코드"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   3
            Left            =   4275
            TabIndex        =   30
            Top             =   420
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "의류분류"
            BorderWidth     =   0
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   7
         Top             =   15
         Width           =   8820
         _ExtentX        =   15558
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
         Caption         =   " 외주 출고 등록 (P_07013)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_07013.frx":067C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   0
         Left            =   8850
         TabIndex        =   8
         Top             =   15
         Width           =   9780
         _ExtentX        =   17251
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
         PictureBackground=   "P_07013.frx":087E
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   9
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
            Picture         =   "P_07013.frx":0A80
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   10
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
            Picture         =   "P_07013.frx":101A
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   11
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
            Picture         =   "P_07013.frx":15B4
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   12
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
            Picture         =   "P_07013.frx":1B4E
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   13
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
            Picture         =   "P_07013.frx":20E8
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   14
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
            Picture         =   "P_07013.frx":2682
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   15
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
            Picture         =   "P_07013.frx":2C1C
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   16
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
            Picture         =   "P_07013.frx":31B6
         End
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   390
         Index           =   2
         Left            =   5310
         TabIndex        =   17
         Top             =   1335
         Width           =   13320
         _ExtentX        =   23495
         _ExtentY        =   688
         _Version        =   262144
         Font3D          =   3
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " 지사 외주 현황"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_07013.frx":3750
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.ProgressBar ProgressBar 
            Height          =   270
            Left            =   5910
            TabIndex        =   18
            Top             =   45
            Visible         =   0   'False
            Width           =   3270
            _Version        =   851970
            _ExtentX        =   5768
            _ExtentY        =   476
            _StockProps     =   93
            Scrolling       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
      End
      Begin XtremeSuiteControls.TabControl TabControl 
         Height          =   9300
         Left            =   5310
         TabIndex        =   21
         Top             =   2325
         Width           =   13320
         _Version        =   851970
         _ExtentX        =   23495
         _ExtentY        =   16404
         _StockProps     =   68
         Appearance      =   3
         Color           =   64
         PaintManager.BoldSelected=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         PaintManager.ButtonMargin=   "2,3,2,3"
         ItemCount       =   1
         Item(0).Caption =   " PDA 스캔 현황 "
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "TabControlPage(0)"
         Begin XtremeSuiteControls.TabControlPage TabControlPage 
            Height          =   8820
            Index           =   0
            Left            =   30
            TabIndex        =   22
            Top             =   450
            Width           =   13260
            _Version        =   851970
            _ExtentX        =   23389
            _ExtentY        =   15557
            _StockProps     =   1
            Page            =   0
            Begin FPSpreadADO.fpSpread spdViewScan 
               Height          =   8235
               Left            =   30
               TabIndex        =   23
               Top             =   690
               Width           =   10875
               _Version        =   524288
               _ExtentX        =   19182
               _ExtentY        =   14526
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
               MaxCols         =   19
               MaxRows         =   35
               SpreadDesigner  =   "P_07013.frx":3BB2
               UserResize      =   1
               Appearance      =   1
               HighlightHeaders=   1
               HighlightStyle  =   1
            End
            Begin XtremeSuiteControls.PushButton cmdBtn 
               Height          =   450
               Index           =   9
               Left            =   120
               TabIndex        =   26
               Top             =   90
               Width           =   3105
               _Version        =   851970
               _ExtentX        =   5477
               _ExtentY        =   794
               _StockProps     =   79
               Caption         =   " PDA 스캔 - 외주 출고 등록"
               ForeColor       =   -2147483640
               BackColor       =   -2147483636
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
               Picture         =   "P_07013.frx":453B
            End
            Begin VB.Label Label1 
               BackStyle       =   0  '투명
               Caption         =   "재세탁 건수:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   300
               Left            =   3495
               TabIndex        =   28
               Top             =   180
               Width           =   1695
            End
            Begin VB.Label lblTagReCount 
               BackStyle       =   0  '투명
               Caption         =   "cnt"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   5325
               TabIndex        =   27
               Top             =   180
               Width           =   615
            End
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10290
         Left            =   15
         TabIndex        =   25
         Top             =   1335
         Width           =   5280
         _Version        =   524288
         _ExtentX        =   9313
         _ExtentY        =   18150
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
         MaxCols         =   3
         MaxRows         =   35
         ScrollBars      =   2
         SpreadDesigner  =   "P_07013.frx":4C35
         UserResize      =   1
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_07013"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String
Dim m_SearchType As String


Private Sub SPR_Resize()
    On Error GoTo ErrRtn
    
    spdViewScan.Width = Me.Width - 5610
    spdViewScan.Height = Me.Height - 3900

    Exit Sub
    
ErrRtn:

End Sub

Private Sub cboOffice_Click()
    If cboOffice.ListIndex < 0 Then Exit Sub
    
    Call Data_Display
End Sub

'-----------------------------------------------------------------
'
'-----------------------------------------------------------------
Private Sub Data_Display()
    On Error GoTo ErrRtn

    ReDim sValue(2)
    Dim nCnt    As Long
    
    spdViewScan.MaxRows = 0
    lblTagReCount.Caption = 0
    nCnt = 0
    sValue(0) = Store.Code
    sValue(1) = Trim(Mid(cboOffice.Text, 2, 4)) & "%"
    sValue(2) = Format(dtInput(0).Value, "YYYYMMDD")
    
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("[SP_M_07013_00]", sValue(), Err_Num, Err_Dec)
    
    With spdView
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            
            .Col = 1: .Text = RS01!코드 & ""
            .Col = 2: .Text = RS01!지사명 & ""
            .Col = 3: .Text = RS01!스캔수량 & ""
            nCnt = nCnt + Val(RS01!스캔수량 & "")
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
    
    
        If .MaxRows >= 1 Then
            .MaxRows = .MaxRows + 1
            .Row = 1
            .Action = SS_ACTION_INSERT_ROW
            
            .Col = 1: .Text = ""
            .Col = 2: .Text = "전   체"
            .Col = 3: .Text = Format(nCnt, "#,##0")
        End If
    
        .Redraw = True
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    cmdBtn(Index).Enabled = False
    Select Case Index
        Case 0: Call Data_Display    ' 조회
        Case 1: 'Call DataAdd        ' 신규
        Case 2: 'Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
        Case 4: 'Call DataCancel     ' 취소
        Case 5: 'Call DataPrint      ' 인쇄
        Case 6: 'Call DataScreen     '
        Case 7
            Unload Me            ' 종료
            Exit Sub
        Case 9: Call Data_Update
    End Select
    
'    Me.MousePointer = 0
    cmdBtn(Index).Enabled = True
    
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

Private Sub Data_Update()
    Dim nRow      As Long
    Dim SSQL        As String
    
    On Error GoTo ERR_RTN
    
'    If spdView.ActiveRow <> 1 Or spdViewScan.DataRowCnt <= 0 Then
'        MsgBox "전체를 선택한 후 작업 하여 주십시요", vbInformation, "확인"
'        Exit Sub
'    End If
    
    ADOCon.BeginTrans
    With spdViewScan
        ReDim sValue(9)
    
        For nRow = 1 To .DataRowCnt
            .Row = nRow
            
            
'            .Col = 1: .Text = RS01!TAGNO & ""    'KEY
'            .Col = 2: .Text = RS01!가맹점명 & ""  '
'            .Col = 3: .Text = RS01!의류코드 & ""    '
'            .Col = 4: .Text = RS01!의류명 & ""    '
'            .Col = 5: .Text = RS01!금액 & ""    '
'
'
'            .Col = 6: .Text = RS01!IPRICE & ""    '
'            .Col = 7: .Text = RS01!OUTDATE & ""      'KEY
'            .Col = 8: .Text = RS01!SCANDT & "" '
'            .Col = 9: .Text = RS01!PDANO & ""    '
'            .Col = 10: .Text = RS01!OUTCD & ""    'KEY
'            .Col = 11: .Text = RS01!MASTERCD & ""    'KEY
'            .Col = 12: .Text = RS01!OCNT & ""    'KEY
'            .Col = 13: .Text = RS01!KIND & ""    '
'            .Col = 14: .Text = RS01!flag & ""    '
'            .Col = 15: .Text = RS01!ADDTYPE & ""    '
'
'
'            .Col = 16: .Text = RS01!IGCODE & ""  '
'            .Col = 17: .Text = RS01!IGDNM & ""    '
            
            
            .Col = 19
            
            If .Text <> "1" Then ' 출고제외가 아니면 등록 '2024-05-08 추가
         
            
                sValue(0) = Store.Code & ""             ' OUTCD
                
                
                sValue(3) = panCaption(1).Tag & ""         ' OCNT
                sValue(8) = Format(dtInput(1).Value, "yyyy-MM-dd") & ""       ' OUTACTIONDATE
                
                .Col = 1: sValue(4) = Replace(.Text & "", "-", "")      ' TAGNO
    
                .Col = 7: sValue(2) = Replace(.Text & "", "-", "")      ' OUTDATE
                .Col = 8: sValue(9) = .Text & ""      ' OUTSCANDT
                .Col = 11: sValue(1) = .Text & ""        ' MASTERCD
                
                .Col = 13: sValue(5) = .Text & ""        ' KIND
                .Col = 14: sValue(6) = .Text & ""        ' flag
                .Col = 15: sValue(7) = .Text & ""      ' ADDTYPE
            
                '------------------------------------------------------------
                ' 외주 출고 등록 - SP_M_07013_04
                '------------------------------------------------------------
                Set RS01 = New ADODB.Recordset
                Set RS01 = ExecPro("SP_M_07013_04", sValue(), Err_Num, Err_Dec)
                
                If Err_Num <> 0 Then
                    ADOCon.RollbackTrans
                    MsgBox Err_Dec
                    Exit Sub
                End If
            End If
        
        Next nRow
    End With
    ADOCon.CommitTrans
    
    Call Data_Display2
    Call Data_Display
    
    MsgBox "저장 완료", vbInformation, "확인"
    
    ' 다음 회차를 준비 한다.
    Call dtInput_Change(0)
    
    Exit Sub

ERR_RTN:
    ADOCon.RollbackTrans
    MsgBox Err.Description
    
End Sub

Private Sub dtInput_Change(Index As Integer)
    dtInput(Index).Enabled = False
    
    ReDim sValue(0)
    'sValue(0) = Format(dtInput(Index).Value, "yyyy-MM-dd")
    sValue(0) = Format(dtInput(1).Value, "yyyy-MM-dd")
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_M_07013_05", sValue(), Err_Num, Err_Dec)
    
    If Not RS01.EOF Then
        panCaption(1).Tag = RS01.Fields("CNT") & ""
        panCaption(1).Caption = RS01.Fields("CNT") & " 회차 출고"
    End If
    
    Call Data_Display
    
    dtInput(Index).Enabled = True
End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(5).Enabled = False
    
    lblTagReCount.Caption = 0
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
    If P_07013_Flag = False Then
        Dim i As Integer
        dtInput(0).Value = Date
        dtInput(1).Value = Date

        '
        Call OrderComboAdd(cboOffice)
        
        With cboOffice
            For i = 0 To .ListCount - 1
                If Mid(.List(i), 2, 4) = HeadOffice Then
                    .ListIndex = i
                    
                    Exit For
                End If
            Next i
        End With
        
        ' 회차를 구해온다.
        Call dtInput_Change(0)
        
        P_07013_Flag = True
    End If

End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
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
    End With
    
    Dim i As Integer
    
    With spdViewScan
        .MaxRows = 0
        .RowHeight(-1) = 14
                
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        '.OperationMode = OperationModeSingle
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    
'        .Col = 8:   .ColHidden = True
'        .Col = 9:   .ColHidden = True
'        .Col = 10:   .ColHidden = True
'        .Col = 11:   .ColHidden = True
'        .Col = 12:   .ColHidden = True
'        .Col = 13:   .ColHidden = True
    
    End With
    
    lblTagReCount.Caption = 0
    
    Call SPR_Resize
    
    If Store.OutClothCode <> "" Then
        Me.txtOutClothCode = Store.OutClothCode
        
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Resize()
    Call SPR_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_07013_Flag = False
    
    SetOuterStr (Me.txtOutClothCode)
    
End Sub

Private Sub Data_Display2()
    On Error GoTo ErrRtn
    
    ReDim sValue(3)
    Dim m_outclothcode  As String
    Dim m_outclothcodelen As String
    
    
    spdView.Row = spdView.ActiveRow
    
    If m_SearchType <> "전체" Then
        spdViewScan.MaxRows = 0
    End If
    
    If spdView.ActiveRow < 2 Then Exit Sub
    sValue(0) = Store.Code
    spdView.Col = 1:        sValue(1) = spdView.Text '+ "%"
    spdView.Col = 4:        sValue(2) = spdView.Text + "%"
    sValue(3) = Format(dtInput(0).Value, "YYYYMMDD")
    
    lblTagReCount.Caption = 0
    
    
    m_outclothcode = Replace(LCase(txtOutClothCode), " ", "")
    m_outclothcodelen = InStr(m_outclothcode, ",") - 1
    
    '------------------------------------------------------------
    ' 외주 출고 등록 - SP_M_07012_01
    '------------------------------------------------------------
    Set RS01 = New ADODB.Recordset
    'Set RS01 = ExecPro("SP_M_07013_01", sValue(), Err_Num, Err_Dec)
    Dim Query As String
    Query = ""
    
    Query = Query + " SELECT "
    Query = Query + "   SUBSTRING(a.TAGNO,1,3) + '-' + SUBSTRING(a.TAGNO,4,2) + '-' + SUBSTRING(a.TAGNO,6,4)        'TAGNO' "
    Query = Query + "   , d.가맹점명   , c.의류코드   , c.의류명   , c.금액   , c.내용   , c.상표"
    Query = Query + "   , Z.IGCODE"
    Query = Query + "   , Z.IGDNM"
    Query = Query + "   , Z.IPRICE "
    Query = Query + "   , CONVERT(CHAR(10),CONVERT(DATETIME,a.OUTDATE),120)    'OUTDATE'"
    Query = Query + "   , a.SCANDT"
    Query = Query + "   , a.PDANO"
    Query = Query + "   , a.OUTCD"
    Query = Query + "   , a.MASTERCD"
    Query = Query + "   , a.OCNT"
    Query = Query + "   , a.KIND"
    Query = Query + "   , a.FLAG"
    Query = Query + "   , a.ADDTYPE"
    Query = Query + "   , Z.OUTACTIONDATE"
    'Query = Query + " FROM OUTORDER_TB AS a LEFT OUTER JOIN ORDER_INOUT2_TB AS Z (NOLOCK) ON  Z.OUTCD = a.OUTCD AND a.MASTERCD = Z.MASTERCD AND a.TAGNO = Z.TAGNO AND ( Z.OUTFLAG IS NULL AND Z.OUTFLAG <> 'Y' )"
    Query = Query + " FROM OUTORDER_TB AS a LEFT OUTER JOIN ORDER_INOUT2_TB AS Z (NOLOCK) ON  Z.OUTCD = a.OUTCD AND a.MASTERCD = Z.MASTERCD AND a.TAGNO = Z.TAGNO"
    Query = Query + "                       INNER JOIN master_tb AS b (NOLOCK)  ON a.mastercd = b.mastercd                           "
    Query = Query + "                       LEFT  JOIN LAUNDRY" & sValue(1) & "..tb_입출고 AS c (NOLOCK)  on a.TAGNO = c.택번호 and c.판매취소 <> 'Y' and 반품환불일자=''"
    'Query = Query + "                       INNER  JOIN LAUNDRY" & sValue(1) & "..tb_입출고 AS c (NOLOCK)  on a.TAGNO = c.택번호 and c.판매취소 <> 'Y' and 반품환불일자='' AND 출고일자 = ''"
    Query = Query + "                       LEFT  JOIN LAUNDRY1000..tb_가맹점 AS d (NOLOCK) on c.지사코드 = d.지사코드 and c.가맹점코드 = d.가맹점코드"
    
    If m_outclothcode <> "" Then
        Query = Query + "                       INNER JOIN LAUNDRY1000.dbo.fn_split('" & m_outclothcode & "',',') e on left(c.의류코드," & m_outclothcodelen & ") = e.element"
    End If
    
    Query = Query + " where a.mastercd = '" & sValue(1) & "'"
    Query = Query + " and a.outdate = '" & sValue(3) & "'"
    'Query = Query + " and c.접수일자 = (select MAX(접수일자) FROM  LAUNDRY" & sValue(1) & "..tb_입출고 (NOLOCK) WHERE 택번호 = A.TAGNO)"
    Query = Query + " ORDER BY a.OUTDATE, a.TAGNO ASC"
    
    
    Set RS01 = ExecQuery(Query, Err_Num, Err_Dec)
    
    With spdViewScan
        
        If m_SearchType <> "전체" Then
            .MaxRows = 0
        End If
        
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = RS01!TAGNO & ""    'KEY
            .Col = 2: .Text = RS01!가맹점명 & ""  '
            .Col = 3: .Text = RS01!의류코드 & ""    '
            .Col = 4: .Text = RS01!의류명 & ""    '
            .Col = 5: .Text = RS01!금액 & ""    '
            
            
            .Col = 6: .Text = RS01!IPRICE & ""    '
            .Col = 7: .Text = RS01!OUTDATE & ""      'KEY
            .Col = 8: .Text = RS01!SCANDT & "" '
            .Col = 9: .Text = RS01!PDANO & ""    '
            .Col = 10: .Text = RS01!OUTCD & ""    'KEY
            .Col = 11: .Text = RS01!MASTERCD & ""    'KEY
            .Col = 12: .Text = RS01!OCNT & ""    'KEY
            .Col = 13: .Text = RS01!KIND & ""    '
            .Col = 14: .Text = RS01!flag & ""    '
            .Col = 15: .Text = RS01!ADDTYPE & ""    '
            
            
            .Col = 16: .Text = RS01!IGCODE & ""  '
            .Col = 17: .Text = RS01!IGDNM & ""    '
            .Col = 18: .Text = RS01!OUTACTIONDATE & ""    '
            
            If .Text <> "" Then
                .Col = 19: .Text = "1"
                lblTagReCount.Caption = lblTagReCount.Caption + 1
            End If
'
            RS01.MoveNext
            
            
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
        
    Exit Sub
    
ErrRtn:
    MsgBox Err.Description
    Resume
End Sub


Private Sub spdView_Click(ByVal Col As Long, ByVal Row As Long)
    'If Row <= 0 Then Exit Sub
    'Call Data_Display2
    
    
    Dim nRow    As Long
    Dim vText   As Variant
    
    If spdView.MaxRows <= 0 Then Exit Sub
    
    If Row = 1 Then
        
        m_SearchType = "전체"
        spdViewScan.MaxRows = 0
        
        With spdView
            For nRow = 1 To .MaxRows
                .GetText 1, nRow, vText
                If Len(vText) = 4 And IsNumeric(vText) Then
                    .Col = 1: .Row = nRow
                    .Action = ActionActiveCell
                    
                    DoEvents
                    Call Data_Display2
                    
                End If

            Next nRow
            .Col = 1
            .Row = 1
            .Action = ActionActiveCell
            
            DoEvents
        End With
        
    Else
        m_SearchType = ""
        Call Data_Display2
    End If
    
    
End Sub

Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    'Call spdView_Click(NewCol, NewRow)
End Sub

Private Sub spdViewScan_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    
'    spdViewScan.Row = Row
'    spdViewScan.Col = Col
'    If spdViewScan.Text = "1" Then
'        MsgBox "선택된 품목은 <외주출고등록>을 하더라도 출고처리를 하지 않으며," _
'                & vbCrLf & "<외주출고등록>을 하면 스캔한 데이터는 삭제됩니다." _
'                & vbCrLf & vbCrLf & "재작업이 필요하면 스캔 등록을 다시 하십시오!", vbInformation
'    End If
    
End Sub
