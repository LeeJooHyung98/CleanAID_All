VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm미수금 
   BorderStyle     =   1  '단일 고정
   Caption         =   "미수금"
   ClientHeight    =   6435
   ClientLeft      =   8205
   ClientTop       =   5595
   ClientWidth     =   10020
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm미수금.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   10020
   StartUpPosition =   1  '소유자 가운데
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   6435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10020
      _ExtentX        =   17674
      _ExtentY        =   11351
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   0
      PaneTree        =   "frm미수금.frx":0A02
      Begin Threed.SSPanel SSPanel 
         Height          =   555
         Left            =   0
         TabIndex        =   1
         Top             =   5880
         Width           =   10020
         _ExtentX        =   17674
         _ExtentY        =   979
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton btnExit 
            Height          =   465
            Left            =   8595
            TabIndex        =   2
            Top             =   45
            Width           =   1305
            _Version        =   851970
            _ExtentX        =   2302
            _ExtentY        =   820
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frm미수금.frx":0A54
         End
         Begin VB.Label lblMisu 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   1110
            TabIndex        =   4
            Top             =   195
            Width           =   120
         End
         Begin VB.Label lblCode 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   105
            TabIndex        =   3
            Top             =   195
            Width           =   120
         End
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   5865
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   10020
         _Version        =   524288
         _ExtentX        =   17674
         _ExtentY        =   10345
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   10
         ScrollBars      =   2
         SpreadDesigner  =   "frm미수금.frx":0FEE
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "frm미수금"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Call 미수금_Display(lblCode.Caption, CLng(lblMisu.Caption))
End Sub

Private Sub Form_Load()

    
    
    With sprGrid
        .MaxRows = 0
        .RowHeight(-1) = 14
        
        .Col = 10: .ColHidden = True
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeExtended
        
        'Init the User Sort
        '.UserColAction = UserColActionSort
    End With
End Sub

Private Sub 미수금_Display(고객코드 As String, 고객미수금 As Long)
    Dim 초기미수금 As Long
    Dim 이전고객   As String
    Dim 초기시작일  As String
    
    Dim bMisu      As Boolean
    Dim 이전미수   As Long
    
    Dim 미수금     As Long
    
    On Error GoTo ErrRtn
    
    Query = "SELECT    ISNULL(초기미수금,0)"
    Query = Query & ", ISNULL(이전고객, '')"
    Query = Query & " FROM TB_고객정보"
    Query = Query & " WHERE 고객코드 = '" & 고객코드 & "'"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF Then
        초기미수금 = 0
        이전고객 = ""
    Else
        초기미수금 = ADORs(0) & ""
        이전고객 = ADORs(1) & ""
    End If
    ADORs.Close
    Set ADORs = Nothing
    
    '---------------------------------------------------
    ' 초기 미수금 적용일자를 구한다.
    '---------------------------------------------------
    Query = "SELECT TOP 1   수정일자"
    Query = Query & " FROM TB_미수금수정"
    Query = Query & " WHERE 고객코드 = '" & 고객코드 & "'"
    Query = Query & "   AND 내용 = '초기 미수금'"
    
    Query = Query & " ORDER BY 수정일자 DESC"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF Then
        초기시작일 = "1900-01-01"
    Else
        초기시작일 = ADORs!수정일자 & ""
    End If
    ADORs.Close
    Set ADORs = Nothing
    
    bMisu = False

    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        '----------------------------------------------------------
        ' TB_매출
        '----------------------------------------------------------
        Query = "SELECT    매출일자"
        Query = Query & ", 매출시간"
        Query = Query & ", 적요"
        Query = Query & ", 접수금액"
        Query = Query & ", 현금입금"
        Query = Query & ", 카드입금"
        Query = Query & ", 사용마일리지"
        Query = Query & ", 쿠폰입금"
        Query = Query & " FROM TB_매출"
        Query = Query & " WHERE 고객코드 = '" & 고객코드 & "'"
        Query = Query & "   AND 매출일자 >= '" & 초기시작일 & "' "

        ' 2013-07-16 일 수정
        'Query = Query & "   AND 카드입금 >=  0 "
        
        Query = Query & " ORDER BY 매출일자 DESC, 매출시간 DESC"
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = Format(ADORs!매출일자, "YY-MM-DD") & "" '
            .Col = 2: .Text = ADORs!접수금액 & "" '
            .Col = 3: .Text = ADORs!현금입금 & "" '
            .Col = 4: .Text = ADORs!카드입금 & "" '
            .Col = 5: .Text = ADORs!사용마일리지 & "" '
            .Col = 6: .Text = ADORs!쿠폰입금 & "" '
            .Col = 7: .Text = ADORs!접수금액 - ADORs!현금입금 - ADORs!카드입금 - ADORs!사용마일리지 - ADORs!쿠폰입금 & "" '
                
            .Col = 10: .Text = ADORs!매출시간 & ""
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        '---------------------------------------------------
        ' TB_미수금수정
        '---------------------------------------------------
        Query = "SELECT    수정일자"
        Query = Query & ", 수정미수금, 내용 "
        Query = Query & " FROM TB_미수금수정"
        Query = Query & " WHERE 고객코드 = '" & 고객코드 & "'"
        Query = Query & " ORDER BY 수정일자 DESC"
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = Format(Left(ADORs!수정일자, 10), "YY-MM-DD") & ""
            .Col = 2:  .Value = 0
            .Col = 3:  .Value = 0
            .Col = 4:  .Value = 0
            .Col = 5:  .Value = 0
            .Col = 6:  .Value = 0
            .Col = 7:  .Value = ADORs!수정미수금 & "": .FontBold = True: .ForeColor = vbRed: .RowHeight(-1) = 14
            .Col = 8:  .Value = 0
            .Col = 9:  .Text = ADORs!내용 & "": .FontBold = True: .ForeColor = vbRed: .RowHeight(-1) = 14
            .Col = 10: .Text = Right(ADORs!수정일자, 8) & ""
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        .SortKey(1) = 1
        .SortKeyOrder(1) = SortKeyOrderDescending
        
        .SortKey(2) = 10
        .SortKeyOrder(2) = SortKeyOrderDescending
        
        .Sort -1, -1, -1, -1, SortByRow
         
        If .MaxRows > 0 Then
            '---------------------------------------------------
            ' 미수금액 계산
            '---------------------------------------------------
            For i = .MaxRows To 1 Step -1
                .Row = i
                .Col = 9
                If .Text = "초기 미수금" Then
                    '초기미수금 제외
                ElseIf .Text = "조정 - 고객수정" Then
                    '초기미수금 제외
                    .Col = 7: 이전미수 = .Value
                    .Col = 8: .Value = 이전미수 & ""
                    
                Else
                    If i = .MaxRows Then
                        이전미수 = 0
                    Else
                        .Row = i + 1
                        .Col = 8: 이전미수 = .Value
                    End If
                    
                    .Row = i
                    .Col = 7: 이전미수 = 이전미수 + .Value
                    .Col = 8: .Value = 이전미수 & ""
                End If
            Next i
                    
            .Row = 1
            .Col = 8: 미수금 = .Value: .FontBold = True: .RowHeight(-1) = 14
        End If
        
        
        Dim Misu_Row As Long
        
        For i = 1 To .MaxRows
            .Row = i
            
            .Col = 1
            If .Text = "초기 미수금" Then
                Misu_Row = i
            End If
            
            If (Misu_Row > 0) And (i > Misu_Row) Then
                .Col = 8: .Value = 0
            End If
        Next i
        
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub
