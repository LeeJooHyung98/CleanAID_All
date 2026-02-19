VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_04009_R3 
   Caption         =   "[전사업장]기간별 세트상품 현황"
   ClientHeight    =   10920
   ClientLeft      =   1590
   ClientTop       =   3435
   ClientWidth     =   16305
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_04009_R3.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10920
   ScaleWidth      =   16305
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   10920
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16305
      _ExtentX        =   28760
      _ExtentY        =   19262
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_04009_R3.frx":058A
      Begin Threed.SSPanel SSPanel 
         Height          =   750
         Index           =   0
         Left            =   15
         TabIndex        =   2
         Top             =   10155
         Width           =   16275
         _ExtentX        =   28707
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   11
            Left            =   4425
            TabIndex        =   13
            Top             =   405
            Width           =   1515
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   10
            Left            =   1425
            TabIndex        =   12
            Top             =   405
            Width           =   1515
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   5
            Left            =   8685
            TabIndex        =   11
            Top             =   405
            Width           =   825
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   4
            Left            =   6870
            TabIndex        =   10
            Top             =   405
            Width           =   825
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   0
            Left            =   1425
            TabIndex        =   9
            Top             =   45
            Width           =   1515
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   1
            Left            =   4425
            TabIndex        =   8
            Top             =   45
            Width           =   1515
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   3
            Left            =   10425
            TabIndex        =   7
            Top             =   45
            Width           =   1065
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   2
            Left            =   7425
            TabIndex        =   6
            Top             =   45
            Width           =   1515
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   6
            Left            =   10515
            TabIndex        =   5
            Top             =   405
            Width           =   825
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   7
            Left            =   12330
            TabIndex        =   4
            Top             =   405
            Width           =   825
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   8
            Left            =   14145
            TabIndex        =   3
            Top             =   405
            Width           =   825
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   45
            TabIndex        =   14
            Top             =   45
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   262144
            BackColor       =   12648384
            Caption         =   "전체매출액"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   3045
            TabIndex        =   15
            Top             =   45
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   262144
            BackColor       =   12648384
            Caption         =   "사업장매출액"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   3
            Left            =   9045
            TabIndex        =   16
            Top             =   45
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   262144
            BackColor       =   12648384
            Caption         =   "입고  수량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   4
            Left            =   6045
            TabIndex        =   17
            Top             =   45
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   262144
            BackColor       =   12648384
            Caption         =   "가맹점매출액"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   5
            Left            =   6045
            TabIndex        =   18
            Top             =   405
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   556
            _Version        =   262144
            BackColor       =   16761024
            Caption         =   "W2"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   7
            Left            =   9690
            TabIndex        =   19
            Top             =   405
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   556
            _Version        =   262144
            BackColor       =   16761024
            Caption         =   "W4"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   12
            Left            =   45
            TabIndex        =   20
            Top             =   405
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   262144
            BackColor       =   12648384
            Caption         =   "전체 단가"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   13
            Left            =   3045
            TabIndex        =   21
            Top             =   405
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   262144
            BackColor       =   12648384
            Caption         =   "사업장 단가"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   6
            Left            =   7860
            TabIndex        =   22
            Top             =   405
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   556
            _Version        =   262144
            BackColor       =   16761024
            Caption         =   "W3"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   9
            Left            =   11505
            TabIndex        =   23
            Top             =   405
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   556
            _Version        =   262144
            BackColor       =   16761024
            Caption         =   "W5"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   10
            Left            =   13320
            TabIndex        =   24
            Top             =   405
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   556
            _Version        =   262144
            BackColor       =   16761024
            Caption         =   "빅"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   8805
         Left            =   15
         TabIndex        =   1
         Top             =   1335
         Width           =   16275
         _Version        =   524288
         _ExtentX        =   28707
         _ExtentY        =   15531
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
         MaxCols         =   15
         SpreadDesigner  =   "P_04009_R3.frx":063C
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   25
         Top             =   540
         Width           =   16275
         _ExtentX        =   28707
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1245
            Style           =   2  '드롭다운 목록
            TabIndex        =   26
            Top             =   60
            Width           =   3315
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1245
            TabIndex        =   27
            Top             =   420
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            Format          =   60555265
            CurrentDate     =   37140
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   28
            Top             =   420
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "조회년월"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   8
            Left            =   60
            TabIndex        =   29
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "지    사"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   3045
            TabIndex        =   30
            Top             =   420
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy-MM"
            Format          =   60555265
            CurrentDate     =   37140
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   14
            Left            =   2775
            TabIndex        =   31
            Top             =   420
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
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
            Caption         =   "~"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   32
         Top             =   15
         Width           =   8670
         _ExtentX        =   15293
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
         PictureBackground=   "P_04009_R3.frx":0FA3
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   1
         Left            =   8700
         TabIndex        =   33
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
         PictureBackground=   "P_04009_R3.frx":11A5
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   34
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
            Picture         =   "P_04009_R3.frx":13A7
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   35
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
            Picture         =   "P_04009_R3.frx":1941
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   36
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
            Picture         =   "P_04009_R3.frx":1EDB
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   37
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
            Picture         =   "P_04009_R3.frx":2475
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   38
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
            Picture         =   "P_04009_R3.frx":2A0F
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   39
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
            Picture         =   "P_04009_R3.frx":2FA9
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   40
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
            Picture         =   "P_04009_R3.frx":3543
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   41
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
            Picture         =   "P_04009_R3.frx":3ADD
         End
      End
   End
End
Attribute VB_Name = "P_04009_R3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01, RS02 As ADODB.Recordset
Dim strSql As String
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cboInput_Change(Index As Integer)
'    Select Case Index
'        Case 0
'            Call Data_Display
'    End Select
End Sub

 
Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: Call Data_Display   ' 조회
        Case 1: 'Call DataAdd        ' 신규
        Case 2: 'Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
        Case 4: 'Call DataCancel     ' 취소
        Case 5: 'Call DataPrint      ' 인쇄
        Case 6: Call Export_Excel(P_00000.cdgExcel, spdView)      ' 엑셀
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
    cmdBtn(5).Enabled = True
    cmdBtn(6).Enabled = True
        
    '본사 여부
    If HeadOffice = MASTER_OFFICE_CODE Then
        cboOffice.Locked = False
    Else
        cboOffice.Locked = True
    End If
        
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
        
        
'        dtInput.Value = Format(Date, "yyyy-mm")
'
'        Call Get_지사리스트(cboOffice)
'
'        ReDim sValue(3)
'
'        cboOffice.ListIndex = 1
'        sValue(0) = "1"
'        sValue(1) = ""
'        sValue(2) = ""
'        sValue(3) = ""
'
'
'        Set RS01 = New ADODB.Recordset
'        Set RS01 = ExecPro("SP_04009_00_ALL", sValue(), Err_Num, Err_Dec)
'
'        spdView.MaxCols = RS01.Fields.Count
'        spdView.MaxRows = RS01.RecordCount
'
'        Call spdDisplay
''       Call fpSpread_Display(spdView, RS01)
'        Call GetColWidth(REG_App, Me.Name, spdView)
        
'        P_04009_Flag = True
'    End If
End Sub

Private Sub spdDisplay()
    
    
    spdView.ColsFrozen = 1 '틀고정
    
    spdView.Row = -1
    
    spdView.Col = 1
    spdView.ColWidth(1) = 20
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft

    spdView.Col = 2
    spdView.ColWidth(2) = 6
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter
    
    spdView.Col = 3
    spdView.ColWidth(3) = 6
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter

    spdView.Col = 4
    spdView.ColWidth(4) = 5
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter


    spdView.Col = 5
    spdView.ColWidth(5) = 12
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 6
    spdView.ColWidth(6) = 8
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 7
    spdView.ColWidth(7) = 12
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 8
    spdView.ColWidth(8) = 6
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
    
    spdView.Col = 9
    spdView.ColWidth(9) = 11
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
    
    spdView.Col = 10
    spdView.ColWidth(10) = 5
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 11
    spdView.ColWidth(11) = 5
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
    
    spdView.Col = 12
    spdView.ColWidth(12) = 5
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
    
    spdView.Col = 13
    spdView.ColWidth(13) = 5
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
    
    spdView.Col = 14
    spdView.ColWidth(14) = 5
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
    
    spdView.Col = 15
    spdView.ColWidth(15) = 6
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
    
  
    spdView.MaxRows = 24
    spdView.Row = 0
    spdView.Col = 1:    spdView.Text = "가맹점"
    spdView.Col = 2:    spdView.Text = "상태"
    spdView.Col = 3:    spdView.Text = "택번호"
    spdView.Col = 4:    spdView.Text = "영업일수"
    spdView.Col = 5:    spdView.Text = "전체매출액"
    spdView.Col = 6:    spdView.Text = "매출단가"
    spdView.Col = 7:    spdView.Text = "사업장매출"
    spdView.Col = 8:    spdView.Text = "사업장단가"
    spdView.Col = 9:    spdView.Text = "가맹점매출"
    spdView.Col = 10:   spdView.Text = "입고수량"
    spdView.Col = 11:   spdView.Text = "W2"
    spdView.Col = 12:   spdView.Text = "W3"
    spdView.Col = 13:   spdView.Text = "W4"
    spdView.Col = 14:   spdView.Text = "W5"
    spdView.Col = 15:   spdView.Text = "빅세트"
    
    'spdView.ShadowColor = glbGray
'    spdView.GrayAreaBackColor = glbGray
'    spdView.MaxRows = 2
    
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
        .OperationMode = OperationModeNormal
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
    
    dtInput(0).Value = Format(Date, "YYYY-MM-01")
    dtInput(1).Value = Format(Date, "YYYY-MM-DD")
    
    Call Get_지사리스트(cboOffice)
    
    Dim i As Integer
    
    With cboOffice
        For i = 0 To .ListCount - 1
            If Mid(.List(i), 2, 4) = HeadOffice Then
                .ListIndex = i
                
                Exit For
            End If
        Next i
    End With
    
    'Call Master_tblComboAdd(cboOffice)
    
    'ReDim sValue(3)
    
    'cboOffice.ListIndex = 0
    'sValue(0) = "1"
    'sValue(1) = ""
    'sValue(2) = ""
    'sValue(3) = ""
    
    'Set RS01 = New ADODB.Recordset
    'Set RS01 = ExecPro("SP_04009_R0", sValue(), Err_Num, Err_Dec)
    
    'spdView.MaxCols = RS01.Fields.Count
    'spdView.MaxRows = RS01.RecordCount
    
    'Call spdDisplay
    '       Call fpSpread_Display(spdView, RS01)
    '        Call GetColWidth(REG_App, Me.Name, spdView)
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

'Private Sub Form_Load()
'    dtInput.Value = Format(Date, "yyyy-mm")
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_04009_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    ReDim sValue(3)
    
    Dim i As Integer
    
    For i = 0 To 8
        txtInput(i).Text = 0
    Next i
        
    sValue(0) = "0"
    sValue(1) = Trim(MidH(cboOffice.Text, 2, 4) & "%")
    sValue(2) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(3) = Format(dtInput(1).Value, "YYYY-MM-DD")
        
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_04009_R03_01", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_04009_R03_01", sValue(), Err_Num, Err_Dec)
    End If
            
    With spdView
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:    .Tet = RS01!가맹점
            
            If RS01!상태 = "Y" Then
                .Col = 2:    .Text = RS01!상태 & ":개점"
            Else
                .Col = 2:    .Text = RS01!상태 & ":폐점"
            End If
            
            .Col = 3:    .Text = RS01!택번호
            .Col = 4:    .Text = RS01!영업일수
            .Col = 5:    .Text = RS01!전체매출액
            
            txtInput(0).Text = txtInput(0).Text + RS01!전체매출액
            
            If RS01!입고수량 = 0 Then
                .Col = 6:    .Text = Format(RS01!입고수량, "##,##0")
                .Col = 8:    .Text = Format(RS01!입고수량, "##,##0")
            Else
                .Col = 6:    .Text = Format(RS01!전체매출액 / RS01!입고수량, "##,##0")
                .Col = 8:    .Text = Format(RS01!사업장매출 / RS01!입고수량, "##,##0")
            End If
            
            .Col = 7:    .Text = RS01!사업장매출
            .Col = 9:    .Text = RS01!가맹점매출
            .Col = 10:   .Text = RS01!입고수량
            .Col = 11:   .Text = RS01!W2
            .Col = 12:   .Text = RS01!W3
            .Col = 13:   .Text = RS01!W4
            .Col = 14:   .Text = RS01!W5
            .Col = 15:   .Text = RS01!WB
            
            txtInput(1).Text = txtInput(1).Text + RS01!사업장매출
            txtInput(2).Text = txtInput(2).Text + RS01!가맹점매출
            txtInput(3).Text = txtInput(3).Text + RS01!입고수량
            txtInput(4).Text = txtInput(4).Text + RS01!W2
            txtInput(5).Text = txtInput(5).Text + RS01!W3
            txtInput(6).Text = txtInput(6).Text + RS01!W4
            txtInput(7).Text = txtInput(7).Text + RS01!W5
            txtInput(8).Text = txtInput(8).Text + RS01!WB
            
            RS01.MoveNext
        Loop
        
        .Redraw = True
    End With
    
    If txtInput(3).Text = 0 Then
        txtInput(10).Text = 0
        txtInput(11).Text = 0
    Else
        txtInput(10).Text = Format(txtInput(0).Text / txtInput(3).Text, "#,##0")
        txtInput(11).Text = Format(txtInput(1).Text / txtInput(3).Text, "#,##0")
    End If
    
    
    For i = 0 To 8
        txtInput(i).Text = Format(txtInput(i).Text, "###,###,##0")
    Next i
    
    RS01.Close
             
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'    If NewRow <> -1 Then
'        spdView.Row = Row
'        spdView.Col = -1
'        spdView.BackColor = vbWhite
'
'        spdView.Row = NewRow
'        spdView.Col = -1
'        spdView.BackColor = glbYellow
'    End If

    With spdView
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

Private Sub spdView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        'PopupMenu P_00000.PopUp
    End If
End Sub

Public Sub DataPrint()
'    Dim ReportFP As String
'    Dim ReportFile As String
'    Dim sData As String
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
'    P_00000.crPrint.Formulas(0) = "월 = '" & Format(dtInput(0).Value, "yyyy-mm") & "'"
'    P_00000.crPrint.Formulas(1) = "사업장 = '" & Trim(cboOffice.Text) & "'"
'
'
'    sData = Space(15) & LeftH(" 합         계" & Space(28), 28)
'    sData = sData & RightH(Space(13) & Format(txtInput(0).Text, "#,##0"), 13)
'    sData = sData & RightH(Space(14) & Format(txtInput(1).Text, "#,##0"), 14)
'    sData = sData & RightH(Space(13) & Format(txtInput(2).Text, "#,##0"), 13)
'    sData = sData & RightH(Space(9) & Format(txtInput(3).Text, "#,##0"), 9)
'    sData = sData & RightH(Space(13) & Format(txtInput(4).Text, "#,##0"), 13)
'    sData = sData & RightH(Space(9) & Format(txtInput(5).Text, "#,##0"), 9)
'
'
'    P_00000.crPrint.Formulas(2) = "합계 = '" & sData & "'"
'
'    Set RS01 = New ADODB.Recordset
'    ReDim sValue(0)
'    sValue(0) = "0"
'    Set RS01 = ExecPro("SP_A_0000", sValue(), Err_Num, Err_Dec)
'    P_00000.crPrint.Formulas(3) = "출력시간 = '" & RS01!DB_DATE & "'"
'
'    Call ReportPrint(ReportFile, "1")
End Sub

Public Sub DataScreen()
'    Dim ReportFP As String
'    Dim ReportFile As String
'    Dim sData As String
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
'    P_00000.crPrint.Formulas(0) = "월 = '" & Format(dtInput(0).Value, "yyyy-mm") & "'"
'    P_00000.crPrint.Formulas(1) = "사업장 = '" & Trim(cboOffice.Text) & "'"
'
'
'    sData = Space(15) & LeftH(" 합         계" & Space(28), 28)
'    sData = sData & RightH(Space(13) & Format(txtInput(0).Text, "#,##0"), 13)
'    sData = sData & RightH(Space(14) & Format(txtInput(1).Text, "#,##0"), 14)
'    sData = sData & RightH(Space(13) & Format(txtInput(2).Text, "#,##0"), 13)
'    sData = sData & RightH(Space(9) & Format(txtInput(3).Text, "#,##0"), 9)
'    sData = sData & RightH(Space(13) & Format(txtInput(4).Text, "#,##0"), 13)
'    sData = sData & RightH(Space(9) & Format(txtInput(5).Text, "#,##0"), 9)
'
'
'    P_00000.crPrint.Formulas(2) = "합계 = '" & sData & "'"
'
'    Set RS01 = New ADODB.Recordset
'    ReDim sValue(0)
'    sValue(0) = "0"
'    Set RS01 = ExecPro("SP_A_0000", sValue(), Err_Num, Err_Dec)
'    P_00000.crPrint.Formulas(3) = "출력시간 = '" & RS01!DB_DATE & "'"
'
'    Call ReportPrint(ReportFile, "2")
End Sub

Private Sub PrintDesc()
    Dim i As Integer
    Dim TempText As String
    Dim TempFP As String
    Dim TempFile As String
    Dim FHandel As Integer
    
    FHandle = FreeFile
    
    TempFP = GetIniStr("REPORT", "TempPath", "", m_iniFile)
    TempFile = TempFP & "\Temp.txt"
    
    
    Open TempFile For Output As #FHandle
    
    TempText = ""
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        
        spdView.Col = 1
        TempText = LeftH(spdView.Text & Space(32), 32)
        spdView.Col = 3
        TempText = TempText & LeftH(spdView.Text & Space(3), 3)
        spdView.Col = 4
        TempText = TempText & RightH(Space(8) & spdView.Text, 8)
        spdView.Col = 5
        TempText = TempText & RightH(Space(14) & spdView.Text, 13)
        spdView.Col = 7
        TempText = TempText & RightH(Space(14) & spdView.Text, 14)
        spdView.Col = 9
        TempText = TempText & RightH(Space(13) & spdView.Text, 13)
        spdView.Col = 10
        TempText = TempText & RightH(Space(9) & spdView.Text, 9)
        spdView.Col = 11
        TempText = TempText & RightH(Space(13) & spdView.Text, 13)
        spdView.Col = 12
        TempText = TempText & RightH(Space(9) & spdView.Text, 9)
        
        Print #FHandle, TempText
    Next i
    
    Close #FHandle
End Sub
