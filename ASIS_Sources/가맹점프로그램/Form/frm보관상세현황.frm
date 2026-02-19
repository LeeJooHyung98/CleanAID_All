VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm보관상세현황 
   Caption         =   "보관품 상세현황"
   ClientHeight    =   8115
   ClientLeft      =   1905
   ClientTop       =   2820
   ClientWidth     =   11790
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   10.5
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   11790
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8115
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   14314
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm보관상세현황.frx":0000
      Begin Threed.SSPanel SSPanel1 
         Height          =   1095
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   11760
         _ExtentX        =   20743
         _ExtentY        =   1931
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1800
            TabIndex        =   3
            Top             =   180
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   58327041
            CurrentDate     =   39024
         End
         Begin Threed.SSPanel ssPanel3 
            Height          =   390
            Index           =   7
            Left            =   90
            TabIndex        =   4
            Top             =   180
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   688
            _Version        =   262144
            Windowless      =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "접수 일자"
            BevelWidth      =   2
            RoundedCorners  =   0   'False
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   0
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin Threed.SSCommand cmdView 
            Height          =   930
            Left            =   10485
            TabIndex        =   5
            Top             =   75
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   1640
            _Version        =   262144
            PictureFrames   =   1
            Windowless      =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "frm보관상세현황.frx":0052
            Caption         =   "조회"
            Alignment       =   8
            ButtonStyle     =   2
            PictureAlignment=   6
            BevelWidth      =   3
         End
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   6975
         Left            =   15
         TabIndex        =   1
         Top             =   1125
         Width           =   11760
         _Version        =   524288
         _ExtentX        =   20743
         _ExtentY        =   12303
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   13
         MaxRows         =   10
         SpreadDesigner  =   "frm보관상세현황.frx":04A4
      End
   End
End
Attribute VB_Name = "frm보관상세현황"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdView_Click()
    Call DataRead
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyChk (KeyCode)
    
    If KeyCode = 13 Then
        SendKeys "{Tab}"
        KeyCode = 0
    End If
End Sub

Private Sub Form_Load()
    'TitleSet "보관 상세 접수 현황"
    
    DTPicker1.Value = Date
    
    Bill_Printer = CStr(GetPrtGubun)
    'Printer_BO_Gb = CStr(GetPrtBOGubun)
End Sub


Private Sub DataRead()
    On Error GoTo ErrRtn
    
    Screen.MousePointer = vbHourglass
    
    Query = "SELECT '', SUBSTRING(InputDate,1,10)  AS InDate,"
    Query = Query & "  ItemIndex, Tag, GoodsCode, SizeGubun, SizeCode, Color, BrandName, BuyPrice, BuyDate, "
    Query = Query & "  ASGubun, BleCount "
    Query = Query & " FROM TB_보관상품리스트 AS P  "
    Query = Query & " WHERE SUBSTRING(InputDate,1,10) = '" & Format(DTPicker1.Value, "YYYY-MM-DD") & "' "
    Query = Query & "   AND StatsFlag <> 'C' "
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF Then GoSub END_RTN
    
    With fpSpread1
        .MaxRows = 0
        
        Do While Not ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        
            .Col = 1:  .Text = "0"
            .Col = 2:  .Text = Format(ADORs!InDate & "", "YYYY-MM-DD")
            .Col = 3:  .Text = CStr(Val(ADORs!ItemIndex & ""))
            .Col = 4:  .Text = Format(ADORs!Tag & "", "@-@@@")
            .Col = 5:  .Text = GetGoodsName(ADORs!GoodsCode & "")
            .Col = 6:  .Text = ADORs!SizeGubun & ""
            .Col = 7:  .Text = ADORs!SizeCode & ""
            .Col = 8:  .Text = ADORs!color & ""
            .Col = 9:  .Text = ADORs!BrandName & ""
            .Col = 10: .Text = Format(Val(ADORs!BuyPrice & ""), "#,##0")
            .Col = 11: .Text = Format(ADORs!BuyDate & "", "YYYY-MM-DD")
            .Col = 12: .Text = ADORs!ASGubun & ""
            .Col = 13: .Text = ADORs!BleCount & ""
            
            ADORs.MoveNext
        Loop
    End With
    ADORs.Close
    Set ADORs = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
END_RTN:
    MsgBox "해당자료가 없습니다." & Space(10), vbInformation, "확인"
    Screen.MousePointer = vbDefault
    Exit Sub

ErrRtn:
    MsgBox Err.Description & Space(10), vbInformation, "확인"
    Screen.MousePointer = vbDefault
End Sub

