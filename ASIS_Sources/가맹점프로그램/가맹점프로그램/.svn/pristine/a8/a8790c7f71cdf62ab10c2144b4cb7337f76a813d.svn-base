VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm출고일자별 
   Caption         =   "출고 일자별 조회"
   ClientHeight    =   12330
   ClientLeft      =   1290
   ClientTop       =   4620
   ClientWidth     =   16770
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form201"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12330
   ScaleWidth      =   16770
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   12330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16770
      _ExtentX        =   29580
      _ExtentY        =   21749
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm출고일자별.frx":0000
      Begin Threed.SSPanel SSPanel1 
         Height          =   1170
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   16740
         _ExtentX        =   29528
         _ExtentY        =   2064
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtCode 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6900
            TabIndex        =   14
            Top             =   615
            Width           =   1140
         End
         Begin VB.TextBox txtName 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   10800
            TabIndex        =   10
            Top             =   60
            Width           =   1650
         End
         Begin VB.TextBox txtTel 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   8085
            TabIndex        =   8
            Top             =   60
            Width           =   1140
         End
         Begin VB.TextBox txtTel 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   6900
            TabIndex        =   7
            Top             =   60
            Width           =   1140
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   495
            Index           =   0
            Left            =   855
            TabIndex        =   3
            Top             =   60
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   61210627
            CurrentDate     =   40279
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   495
            Index           =   1
            Left            =   3210
            TabIndex        =   4
            Top             =   60
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   61210627
            CurrentDate     =   40279
         End
         Begin MSMask.MaskEdBox mskTag 
            Height          =   495
            Left            =   855
            TabIndex        =   13
            Top             =   615
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   873
            _Version        =   393216
            MaxLength       =   7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "  #-###"
            PromptChar      =   " "
         End
         Begin XtremeSuiteControls.PushButton cmdFind 
            Height          =   540
            Left            =   9915
            TabIndex        =   16
            Top             =   585
            Width           =   1245
            _Version        =   851970
            _ExtentX        =   2196
            _ExtentY        =   952
            _StockProps     =   79
            Caption         =   " 찾기"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Picture         =   "frm출고일자별.frx":0052
         End
         Begin XtremeSuiteControls.PushButton Command1 
            Height          =   540
            Left            =   11205
            TabIndex        =   17
            Top             =   585
            Width           =   1245
            _Version        =   851970
            _ExtentX        =   2196
            _ExtentY        =   952
            _StockProps     =   79
            Caption         =   " Clear"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Picture         =   "frm출고일자별.frx":0A64
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "고객번호"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   5850
            TabIndex        =   15
            Top             =   675
            Width           =   960
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "택번호"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   60
            TabIndex        =   12
            Top             =   660
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "성명"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   10215
            TabIndex        =   11
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "전화번호"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   5850
            TabIndex        =   9
            Top             =   120
            Width           =   960
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "일자"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   285
            TabIndex        =   6
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "~"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   2970
            TabIndex        =   5
            Top             =   120
            Width           =   180
         End
      End
      Begin FPSpreadADO.fpSpread fpSpread2 
         Height          =   11115
         Left            =   15
         TabIndex        =   1
         Top             =   1200
         Width           =   16740
         _Version        =   524288
         _ExtentX        =   29528
         _ExtentY        =   19606
         _StockProps     =   64
         AutoCalc        =   0   'False
         BackColorStyle  =   1
         ColsFrozen      =   4
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
         DInformActiveRowChange=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         MaxCols         =   13
         MaxRows         =   200
         MoveActiveOnFocus=   0   'False
         OperationMode   =   1
         Protect         =   0   'False
         ScrollBarExtMode=   -1  'True
         ScrollBarMaxAlign=   0   'False
         SpreadDesigner  =   "frm출고일자별.frx":1476
         UserResize      =   0
         VisibleCols     =   13
         VisibleRows     =   50
         HighlightHeaders=   1
         HighlightStyle  =   1
         ScrollBarStyle  =   2
      End
   End
End
Attribute VB_Name = "frm출고일자별"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim j        As Integer
Dim strStart As String
Dim strEnd   As String

Private Sub vasp_Clear()
    With fpSpread2
        .Col = 1
        .MaxRows = 0
    End With
End Sub

Private Sub Search_All()
    strStart = Format(dtpDay(0).Value, "YYYY-MM-DD") 'Trim(mskY1.Text) & Trim(mskM1.Text) & Trim(mskD1.Text)
    strEnd = Format(dtpDay(1).Value, "YYYY-MM-DD")   'Trim(mskY2.Text) & Trim(mskM2.Text) & Trim(mskD2.Text)
    
    If IsDate(strStart) Or IsDate(strEnd) Then
        MsgBox " 날짜를 확인 하여주십시요", vbInformation, "확인"
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    Query = "SELECT    P.성명"
    Query = Query & ", P.전화번호"
    Query = Query & ", P1.접수일자"
    Query = Query & ", P1.출고일자"
    Query = Query & ", P1.지사출고상태"
    Query = Query & ", P1.의류명"
    Query = Query & ", P1.택번호"
    Query = Query & ", P1.색상"
    Query = Query & ", P1.내용"
    Query = Query & ", P1.금액"
    Query = Query & ", P1.결제여부"
    Query = Query & ", P1.상표"
    Query = Query & ", P1.확인 "
    Query = Query & " FROM TB_고객정보 AS P LEFT OUTER JOIN TB_입출고 AS P1 ON P.고객코드 = P1.고객코드"
    Query = Query & " WHERE (P1.고객코드 LIKE '%" & Trim(txtCode.Text) & "%') "
    Query = Query & "   AND (P.성명 LIKE '%" & Trim(txtName.Text) & "%' ) "
    Query = Query & "   AND (P.전화번호 LIKE '%" & Trim(txtTel(0).Text) & "%') "
    Query = Query & "   AND ( P1.택번호  LIKE '%" & Trim(mskTag.Text) & "%' )"
    Query = Query & "   AND P1.출고일자 >= '" & strStart & "' "
    Query = Query & "   AND P1.출고일자 <= '" & strEnd & "' "
    Query = Query & "   AND (P1.판매취소 IS NULL OR P1.판매취소 <> 'Y') "
    Query = Query & " GROUP BY P1.고객코드, P.성명, P.전화번호, P1.접수일자, P1.출고일자, "
    Query = Query & " P1.지사출고상태, P1.의류명, P1.택번호, P1.색상, P1.내용, P1.금액, P1.결제여부, P1.상표, P1.확인 "
    Query = Query & " ORDER BY P1.접수일자 DESC, P1.택번호 DESC"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF = True Then
        ADORs.Close
        Set ADORs = Nothing
        
        Screen.MousePointer = vbDefault
        MsgBox "[ " & txtTel(0).Text & "-" & txtTel(1).Text & " ] 에 해당되는 자료가 없읍니다 !", vbInformation, "확인"
        Exit Sub
    End If
    
    i = 1
    j = 1
    Call vasp_Clear
    
    With fpSpread2
        .ReDraw = False
        .MaxRows = 18
        .Row = 0
        
        Do Until ADORs.EOF
            If .Row >= .MaxRows Then
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Action = ActionInsertRow
                .RowHeight(.MaxRows) = .RowHeight(0) ' 마지막 라인의 높이를 맞춘다.
            End If
            
            .Row = .Row + 1
            For j = 1 To 13
                .Col = j: .Text = IIf(IsNull(ADORs(j - 1)), "", ADORs(j - 1))
            Next j
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        .ReDraw = True
    End With
        
    Screen.MousePointer = vbDefault
End Sub

Private Sub Search_Date()
    'Description : 주소로 찾기
    'History : 2002/12/06
    
    Screen.MousePointer = vbHourglass
    
    strStart = Format(dtpDay(0).Value, "YYYY-MM-DD") 'Trim(mskY1.Text) & Trim(mskM1.Text) & Trim(mskD1.Text)
    strEnd = Format(dtpDay(1).Value, "YYYY-MM-DD") 'Trim(mskY2.Text) & Trim(mskM2.Text) & Trim(mskD2.Text)
    
    Query = "SELECT    P.성명"
    Query = Query & ", P.전화번호"
    Query = Query & ", P1.접수일자"
    Query = Query & ", P1.출고일자"
    Query = Query & ", P1.지사출고상태"
    Query = Query & ", P1.의류명"
    Query = Query & ", P1.택번호"
    Query = Query & ", P1.색상"
    Query = Query & ", P1.내용"
    Query = Query & ", P1.금액"
    Query = Query & ", P1.결제여부"
    Query = Query & ", P1.상표"
    Query = Query & ", P1.확인 "
    Query = Query & " FROM TB_고객정보 AS P LEFT OUTER JOIN TB_입출고 AS P1 ON P.고객코드 = P1.고객코드"
    Query = Query & " WHERE (P.전화번호 LIKE '%" & Trim(txtTel(0).Text) & "%') "
    Query = Query & "   AND P1.출고일자 >= '" & strStart & "' "
    Query = Query & "   AND P1.출고일자 <= '" & strEnd & "' "
    Query = Query & "   AND (P1.판매취소 IS NULL OR P1.판매취소 <> 'Y') "
    Query = Query & " GROUP BY P1.고객코드, P.성명, P.전화번호, P1.접수일자, P1.출고일자, "
    Query = Query & " P1.지사출고상태, P1.의류명, P1.택번호, P1.색상, P1.내용, P1.금액, P1.결제여부, P1.상표, P1.확인 "
    Query = Query & " ORDER BY P1.접수일자, P1.택번호 "
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF = True Then
        ADORs.Close
        Screen.MousePointer = vbDefault
        MsgBox "[" & txtTel(0).Text & "-" & txtTel(1).Text & "] 에 해당되는 자료가 없읍니다 !", vbInformation, "고객 찾기"
        Exit Sub
    End If
    
    i = 1
    j = 1
    vasp_Clear
    With fpSpread2
        .ReDraw = False
        .MaxRows = 18
        .Row = 0
        While Not ADORs.EOF = True
            If .Row >= .MaxRows Then
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Action = ActionInsertRow
                .RowHeight(.MaxRows) = .RowHeight(0) ' 마지막 라인의 높이를 맞춘다.
            End If
            .Row = .Row + 1
            For j = 1 To 13
                .Col = j: .Text = IIf(IsNull(ADORs(j - 1)), "", ADORs(j - 1))
            Next j
            ADORs.MoveNext
        Wend
        .ReDraw = True
    End With
    
    ADORs.Close
    Screen.MousePointer = vbDefault
End Sub

Private Sub Search_Tel()
    'Description : 전화번호로 찾기
    'History : 98/04/11
    If Len(Trim(txtTel(1))) < 1 Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    strStart = Format(dtpDay(0).Value, "YYYY-MM-DD") 'Trim(mskY1.Text) & Trim(mskM1.Text) & Trim(mskD1.Text)
    strEnd = Format(dtpDay(1).Value, "YYYY-MM-DD") 'Trim(mskY2.Text) & Trim(mskM2.Text) & Trim(mskD2.Text)
    
    Query = "SELECT P.성명, P.전화번호, P1.접수일자, P1.출고일자, P1.지사출고상태, P1.의류명,"
    Query = Query & "P1.택번호, P1.색상, P1.내용, P1.금액, P1.결제여부, P1.상표, P1.확인 "
    Query = Query & " FROM TB_고객정보 AS P LEFT OUTER JOIN TB_입출고 AS P1 ON P.고객코드 = P1.고객코드"
    Query = Query & " WHERE (P.전화번호 LIKE '%" & Trim(txtTel(0).Text) & "%') "
    Query = Query & "   AND P1.출고일자 >= '" & strStart & "' "
    Query = Query & "   AND P1.출고일자 <= '" & strEnd & "' "
    Query = Query & "   AND (P1.판매취소 IS NULL OR P1.판매취소 <> 'Y') "
    Query = Query & " GROUP BY P1.고객코드, P.성명, P.전화번호, P1.접수일자, P1.출고일자, "
    Query = Query & " P1.지사출고상태, P1.의류명, P1.택번호, P1.색상, P1.내용, P1.금액, P1.결제여부, P1.상표, P1.확인 "
    Query = Query & " ORDER BY P1.접수일자, P1.택번호 "
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF = True Then
        ADORs.Close
        Screen.MousePointer = vbDefault
        MsgBox "[" & txtTel(0).Text & "-" & txtTel(1).Text & "] 에 해당되는 자료가 없읍니다 !", vbInformation, "고객 찾기"
        Exit Sub
    End If
    
    i = 1
    j = 1
    vasp_Clear
    With fpSpread2
        .ReDraw = False
        .MaxRows = 18
        .Row = 0
        While Not ADORs.EOF = True
            If .Row >= .MaxRows Then
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Action = ActionInsertRow
                .RowHeight(.MaxRows) = .RowHeight(0) ' 마지막 라인의 높이를 맞춘다.
            End If
            .Row = .Row + 1
            For j = 1 To 13
                .Col = j: .Text = IIf(IsNull(ADORs(j - 1)), "", ADORs(j - 1))
            Next j
            ADORs.MoveNext
        Wend
        .ReDraw = True
    End With
    
    ADORs.Close
    Screen.MousePointer = vbDefault
End Sub

Private Sub Search_CustomID()
    'Description : 고객코드로 찾기
    'History : 98/04/11
     
    If Len(Trim(txtCode.Text)) < 1 Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    strStart = Format(dtpDay(0).Value, "YYYY-MM-DD") 'Trim(mskY1.Text) & Trim(mskM1.Text) & Trim(mskD1.Text)
    strEnd = Format(dtpDay(1).Value, "YYYY-MM-DD") 'Trim(mskY2.Text) & Trim(mskM2.Text) & Trim(mskD2.Text)
    
    Query = "SELECT P.성명, P.전화번호, P1.접수일자, P1.출고일자, P1.지사출고상태, P1.의류명, "
    Query = Query & "P1.택번호, P1.색상, P1.내용, P1.금액, P1.결제여부, P1.상표, P1.확인 "
    Query = Query & " FROM TB_고객정보 AS P LEFT OUTER JOIN TB_입출고 AS P1 ON P.고객코드 = P1.고객코드"
    Query = Query & " WHERE (P1.고객코드 LIKE '%" & Trim(txtCode.Text) & "%') "
    Query = Query & "   AND  P1.출고일자 >= '" & strStart & "' "
    Query = Query & "   AND  P1.출고일자 <= '" & strEnd & "' "
    Query = Query & "   AND (P1.판매취소 IS NULL OR P1.판매취소 <> 'Y') "
    Query = Query & " GROUP BY P1.고객코드, P.성명, P.전화번호, P1.접수일자, P1.출고일자, "
    Query = Query & " P1.지사출고상태, P1.의류명, P1.택번호, P1.색상, P1.내용, P1.금액, P1.결제여부, P1.상표, P1.확인 "
    Query = Query & " ORDER BY P1.접수일자, P1.택번호 "
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF = True Then
        ADORs.Close
        Screen.MousePointer = vbDefault
        MsgBox "[" & txtCode.Text & "] 에 해당되는 자료가 없읍니다 !", vbInformation, " 고객 찾기"
        Exit Sub
    End If
    
    i = 1
    j = 1
    vasp_Clear
    
    With fpSpread2
        .ReDraw = False
        .MaxRows = 18
        .Row = 0
        While Not ADORs.EOF = True
            If .Row >= .MaxRows Then
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Action = ActionInsertRow
                .RowHeight(.MaxRows) = .RowHeight(0) ' 마지막 라인의 높이를 맞춘다.
            End If
            .Row = .Row + 1
            For j = 1 To 13
                .Col = j: .Text = IIf(IsNull(ADORs(j - 1)), "", ADORs(j - 1))
            Next j
            ADORs.MoveNext
        Wend
        .ReDraw = True
    End With
    Screen.MousePointer = vbDefault
    ADORs.Close
End Sub

Private Sub Search_Name()
    'Description : 고객성명으로 찾기
    'History : 98/04/11
 
    If Len(Trim(txtName.Text)) < 1 Then
       Exit Sub
    End If
     
    Screen.MousePointer = vbHourglass
    strStart = Format(dtpDay(0).Value, "YYYY-MM-DD") 'Trim(mskY1.Text) & Trim(mskM1.Text) & Trim(mskD1.Text)
    strEnd = Format(dtpDay(1).Value, "YYYY-MM-DD") 'Trim(mskY2.Text) & Trim(mskM2.Text) & Trim(mskD2.Text)
    
    Query = "SELECT P.성명, P.전화번호, P1.접수일자, P1.출고일자, P1.지사출고상태, P1.의류명, "
    Query = Query & "P1.택번호, P1.색상, P1.내용, P1.금액, P1.결제여부, P1.상표, P1.확인 "
    Query = Query & " FROM TB_고객정보 AS P LEFT OUTER JOIN TB_입출고 AS P1 ON P.고객코드 = P1.고객코드"
    Query = Query & " WHERE (P.성명 LIKE '%" & Trim(txtName.Text) & "%' ) "
    Query = Query & "   AND P1.출고일자 >= '" & strStart & "' "
    Query = Query & "   AND P1.출고일자 <= '" & strEnd & "' "
    Query = Query & "   AND (P1.판매취소 IS NULL OR P1.판매취소 <> 'Y') "
    Query = Query & " GROUP BY P1.고객코드, P.성명, P.전화번호, P1.접수일자, P1.출고일자, "
    Query = Query & " P1.지사출고상태, P1.의류명, P1.택번호, P1.색상, P1.내용, P1.금액, P1.결제여부, P1.상표, P1.확인 "
    Query = Query & " ORDER BY P1.접수일자, P1.택번호 "
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF = True Then
        ADORs.Close
        Screen.MousePointer = vbDefault
        MsgBox "[" & txtName.Text & "] 에 해당되는 자료가 없읍니다 !", vbInformation, " 고객 찾기"
        Exit Sub
    End If
    
    i = 1
    j = 1
    vasp_Clear
    With fpSpread2
        .ReDraw = False
        .MaxRows = 18
        .Row = 0
        While Not ADORs.EOF = True
            If .Row >= .MaxRows Then
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Action = ActionInsertRow
                .RowHeight(.MaxRows) = .RowHeight(0) ' 마지막 라인의 높이를 맞춘다.
            End If
            .Row = .Row + 1
            For j = 1 To 13
                .Col = j: .Text = IIf(IsNull(ADORs(j - 1)), "", ADORs(j - 1))
            Next j
            ADORs.MoveNext
        Wend
        .ReDraw = True
    End With
    
    Screen.MousePointer = vbDefault
    ADORs.Close
End Sub

Private Sub Search_TagNo()
    'Description : 택번호로 찾기
    'History : 98/04/11
     
    
    If Len(Trim(mskTag.Text)) < 1 Then
        Exit Sub
    End If
     
    Screen.MousePointer = vbHourglass
    
    strStart = Format(dtpDay(0).Value, "YYYY-MM-DD") 'Trim(mskY1.Text) & Trim(mskM1.Text) & Trim(mskD1.Text)
    strEnd = Format(dtpDay(1).Value, "YYYY-MM-DD")   'Trim(mskY2.Text) & Trim(mskM2.Text) & Trim(mskD2.Text)
    
    Query = "SELECT P.성명, P.전화번호, P1.접수일자, P1.출고일자, P1.지사출고상태, P1.의류명, "
    Query = Query & "P1.택번호, P1.색상, P1.내용, P1.금액, P1.결제여부, P1.상표, P1.확인 "
    Query = Query & " FROM TB_고객정보 AS P LEFT OUTER JOIN TB_입출고 AS P1 ON P.고객코드 = P1.고객코드"
    Query = Query & " WHERE ( P1.택번호  LIKE '%" & Trim(mskTag.Text) & "%' )"
    Query = Query & "   AND P1.출고일자 >='" & strStart & "' AND P1.출고일자 <='" & strEnd & "'"
    Query = Query & "   AND (P1.판매취소 IS NULL OR P1.판매취소 <> 'Y') "
    Query = Query & " GROUP BY P1.고객코드, P.성명, P.전화번호, P1.접수일자, P1.출고일자, "
    Query = Query & " P1.지사출고상태, P1.의류명, P1.택번호, P1.색상, P1.내용, P1.금액, P1.결제여부, P1.상표, P1.확인 "
    Query = Query & " ORDER BY P1.접수일자, P1.택번호 "
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    
    If ADORs.EOF = True Then
        Screen.MousePointer = vbDefault
        ADORs.Close
        MsgBox "[" & mskTag.Text & "] 에 해당되는 자료가 없읍니다 !", vbInformation, " 고객 찾기"
        Exit Sub
    End If
    
    i = 1
    j = 1
    vasp_Clear
    With fpSpread2
        .ReDraw = False
        .MaxRows = 18
        .Row = 0
        While Not ADORs.EOF = True
            If .Row >= .MaxRows Then
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Action = ActionInsertRow
                .RowHeight(.MaxRows) = .RowHeight(0) ' 마지막 라인의 높이를 맞춘다.
            End If
            .Row = .Row + 1
            For j = 1 To 13
                .Col = j: .Text = ADORs(j - 1) & ""
            Next j
            ADORs.MoveNext
        Wend
        .ReDraw = True
    End With
    
    Screen.MousePointer = vbDefault
    ADORs.Close
End Sub

Private Sub cmdFind_Click()
    Call Search_All
End Sub

Private Sub Command1_Click()
    Dim strDate As String
    
    'mskY1.Text = ""
    'mskM1.Text = ""
    'mskD1.Text = ""
    
    txtTel(0).Text = ""
    txtTel(1).Text = ""
    txtCode.Text = ""
    mskTag.SelText = ""
    txtName.Text = ""
    
    strDate = Format(DateAdd("m", -1, Date), "YYYY-MM-DD")
    
    dtpDay(0).Value = Format(strDate, "YYYY-MM-DD")
    
    'mskY1.Text = Format(strDate, "yyyy")
    'mskM1.Text = Format(strDate, "mm")
    'mskD1.Text = Format(strDate, "dd")
    
    Call vasp_Clear
    
    txtTel(0).SetFocus
End Sub

Private Sub Form_Activate()
    'mskY2.Text = Format(Date, "yyyy")
    'mskM2.Text = Format(Date, "mm")
    'mskD2.Text = Format(Date, "dd")
    
    dtpDay(1).Value = Format(Date, "YYYY-MM-DD")
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    'TitleSet "출고 일자별 조회"
    
    'mskY1.Text = Format(Date, "yyyy")
    'mskM1.Text = Format(Date, "mm")
    'mskD1.Text = Format(Date, "dd")
    
    dtpDay(0).Value = Format(Date, "YYYY-MM-DD")
    dtpDay(1).Value = Format(Date, "YYYY-MM-DD")
End Sub

Private Sub txtCode_GotFocus()
    txtCode.SelStart = 0
    txtCode.SelLength = 6
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Search_All
        'Search_CustomID
    End If
End Sub

'Private Sub mskD1_GotFocus()
'    mskD1.SelStart = 0
'    mskD1.SelLength = 2
'End Sub
'
'Private Sub mskD2_GotFocus()
'    mskD2.SelStart = 0
'    mskD2.SelLength = 2
'End Sub

'Private Sub mskD2_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        Search_All
'        'Search_Date
'    End If
'End Sub

'Private Sub mskM1_GotFocus()
'    mskM1.SelStart = 0
'    mskM1.SelLength = 2
'End Sub
'
'Private Sub mskM2_GotFocus()
'    mskM2.SelStart = 0
'    mskM2.SelLength = 2
'End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Search_All
        'Search_Name
    End If
End Sub

Private Sub mskTag_Change()
    mskTag.SelStart = 0
    mskTag.SelLength = 8
End Sub

Private Sub mskTag_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Search_All
        'Search_TagNo
    End If
End Sub

Private Sub txtTel_GotFocus(Index As Integer)
    txtTel(Index).SelStart = 0
    txtTel(Index).SelLength = 4
End Sub

Private Sub txtTEL_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Select Case Index
            Case 0
            
            Case 1
                Query = " SELECT * FROM TB_고객정보 "
                Query = Query & " WHERE 전화번호 = '" & txtTel(0).Text & "' "
                Set ADORs = New ADODB.Recordset
                ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
                
                If ADORs.RecordCount = 1 Then
                    ADORs.Close
                    Set ADORs = Nothing
                    
                    Call Search_All ' 찿기
                    'Search_Tel
                    
                    Exit Sub
                ElseIf ADORs.RecordCount < 1 Then
                    ADORs.Close
                    Set ADORs = Nothing
                
                    MsgBox " 등록된 회원이 없습니다.", vbInformation, "확인"
                    Exit Sub
                    
                ElseIf ADORs.RecordCount >= 2 Then
                    ADORs.Close
                    Set ADORs = Nothing
                
                    '------------------------------------------
                    ' 뿌리고 입력대기상태
                    '------------------------------------------
                    frm고객검색.DataDisplay Query
                    frm고객검색.Show 1
                    
                    If frm고객검색.SELECTCODE = "CANCEL" Then
                        txtTel(1).SetFocus
                        Exit Sub
                    End If
                    
                    If 고객정보.고객코드 <> "Error" Then txtCode = 고객정보.고객코드
                    
                    txtName.Text = 고객정보.성명
                    
                    Call Search_All
                    
                    Exit Sub
                End If
        End Select
    End If
End Sub

'Private Sub mskY1_GotFocus()
'    mskY1.SelStart = 0
'    mskY1.SelLength = 4
'End Sub
'
'Private Sub mskY2_GotFocus()
'    mskY2.SelStart = 0
'    mskY2.SelLength = 4
'End Sub

