VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm보관접수현황 
   Caption         =   "보관 접수증"
   ClientHeight    =   11145
   ClientLeft      =   2925
   ClientTop       =   5970
   ClientWidth     =   15120
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
   ScaleHeight     =   11145
   ScaleWidth      =   15120
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   11145
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   19659
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm보관접수현황.frx":0000
      Begin Threed.SSPanel SSPanel1 
         Height          =   1215
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   15090
         _ExtentX        =   26617
         _ExtentY        =   2143
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CommandButton Command1 
            Caption         =   "보관증 재출력"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   915
            Left            =   7740
            TabIndex        =   3
            Top             =   60
            Width           =   2445
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1800
            TabIndex        =   4
            Top             =   150
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
            Format          =   58589185
            CurrentDate     =   39024
         End
         Begin Threed.SSPanel ssPanel3 
            Height          =   390
            Index           =   7
            Left            =   90
            TabIndex        =   5
            Top             =   150
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
            Left            =   10320
            TabIndex        =   6
            Top             =   45
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
            Picture         =   "frm보관접수현황.frx":0052
            Caption         =   "조회"
            Alignment       =   8
            ButtonStyle     =   2
            PictureAlignment=   6
            BevelWidth      =   3
         End
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   9885
         Left            =   15
         TabIndex        =   1
         Top             =   1245
         Width           =   15090
         _Version        =   524288
         _ExtentX        =   26617
         _ExtentY        =   17436
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
         SpreadDesigner  =   "frm보관접수현황.frx":04A4
         ScrollBarStyle  =   2
      End
   End
End
Attribute VB_Name = "frm보관접수현황"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdView_Click()
    Call DataRead
End Sub

Private Sub Command1_Click()
    Dim nRow    As Long
    
    ' 접수증 프린트
    With fpSpread1
        For nRow = 1 To .MaxRows
            .Col = 1:       .Row = nRow
            If .Value = 1 Then
                .Col = 8
                Call Print_QN_MM(.Text)
            End If
        Next nRow
    End With
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyChk (KeyCode)
    
    If KeyCode = 13 Then
        SendKeys "{Tab}"
        KeyCode = 0
    End If
End Sub

Private Sub Form_Load()
    
    'TitleSet "보관 접수 현황"
    DTPicker1.Value = Date
    
    Bill_Printer = CStr(GetPrtGubun)
    'Printer_BO_Gb = CStr(GetPrtBOGubun)
    
    
End Sub


Private Sub DataRead()
    Dim nRow As Long
    
    On Error GoTo ErrRtn
    
    Screen.MousePointer = vbHourglass
    
    
    Query = "SELECT '', SUBSTRING(InputDate,1,10)  AS InDate,"
    Query = Query & "  InputName, SaleGubunCode, "
    Query = Query & "  Left(SaleEndDate,8)   AS EndDate,"
    Query = Query & "  Price, ItemCount, KeyCode "
    Query = Query & " FROM TB_보관리스트 AS P  "
    Query = Query & " WHERE SUBSTRING(InputDate,1,10) = '" & Format(DTPicker1.Value, "YYYY-MM-DD") & "' "
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF Then GoSub END_RTN
    
    With fpSpread1
        .MaxRows = 0
        
        Do While Not ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        
            .Col = 2:   .Text = Format(ADORs!InDate & "", "YYYY-MM-DD")
            .Col = 3:   .Text = ADORs!InputName & ""
            .Col = 4:   .Text = ADORs!SaleGubunCode & "개월"
            .Col = 5:   .Text = Format(ADORs!EndDate & "", "YYYY-MM-DD")
            .Col = 6:   .Text = Format(Val(ADORs!Price & ""), "#,##0")
            .Col = 7:   .Text = ADORs!ItemCount & ""
            .Col = 8:   .Text = ADORs!KeyCode & ""
            
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

