VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm의류상세 
   Caption         =   "품목"
   ClientHeight    =   8760
   ClientLeft      =   16140
   ClientTop       =   5640
   ClientWidth     =   4380
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
   LinkTopic       =   "Form9"
   LockControls    =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   4380
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8760
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   15452
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm의류상세.frx":0000
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   8730
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   4350
         _Version        =   524288
         _ExtentX        =   7673
         _ExtentY        =   15399
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BackColorStyle  =   1
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
         DInformActiveRowChange=   0   'False
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
         MaxCols         =   4
         MaxRows         =   30
         Protect         =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frm의류상세.frx":0032
         VisibleCols     =   3
         VisibleRows     =   30
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "frm의류상세"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim iEOF As Boolean
    Dim ADORs       As ADODB.RecordSet
    Dim sGoodsStats As String
    
    On Error GoTo ErrRtn
    
    frm의류상세.Top = frmMain.Top   '1000
    frm의류상세.Left = frmMain.Left '6000
    
    With sprGrid
        .MaxRows = 0
        .RowHeight(-1) = 14
    
        .Col = 3: .ColHidden = True '
        .Col = 4: .ColHidden = True '
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle

        '선택된 Row
        .SelBackColor = &H80FFFF '&HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeSingle
                
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
        
    Set ADORs = New ADODB.RecordSet
    Set ADORs = Get_의류정보(Left(의류코드, 2), sGoodsStats, frm접수.btnInternet.tag)
    frm접수.lblGoodsPriceStats.Caption = sGoodsStats
    

    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = ADORs!의류명 & ""   '
            .Col = 2: .Text = ADORs!금액 & ""     '
            .Col = 3: .Text = ADORs!의류코드 & "" '
            .Col = 4: .Text = ADORs!순서 & ""     '
            
            ADORs.MoveNext
        Loop
        ADORs.Close:    Set ADORs = Nothing
        
        .ReDraw = True
    
        .Refresh
        .CursorStyle = 2
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub sprGrid_Click(ByVal Col As Long, ByVal Row As Long)
    Dim 의류명  As String
    Dim strPrice As String
    Dim iPrice   As Long
    
    If Row < 1 Then
        If sprGrid.MaxRows = 0 Then
            Unload Me
        End If
        Exit Sub
    End If
    
    sprGrid.Row = Row
    sprGrid.Col = 1: 의류명 = Trim(sprGrid.Text) & ""
    
    If Trim(의류명) = "" Then
        Unload Me
        
        Exit Sub
    End If
    
    sprGrid.Col = 2: strPrice = sprGrid.Value
    sprGrid.Col = 3: 의류코드 = sprGrid.Value
    
    
    iPrice = Get_DryPrice(의류코드, frm접수.btnInternet.tag)
    
    frm접수.sprGrid.Row = iCur 'frm접수.sprGrid.ActiveRow
    frm접수.sprGrid.Col = 1:  frm접수.sprGrid.Text = 의류명 & ""   ' 1
    frm접수.sprGrid.Col = 6:  frm접수.sprGrid.Value = iPrice       ' 6 strPrice
    frm접수.sprGrid.Col = 8:  frm접수.sprGrid.Text = 의류코드 & "" ' 8
    frm접수.sprGrid.Col = 14: frm접수.sprGrid.Value = iPrice       '14 세트 상품의 원 금액을 기록한다.
    frm접수.sprGrid.Col = 20: frm접수.sprGrid.Value = Get_세탁정상금액(의류코드)     '20 정상 금액을 기록한다.
        
    frm접수.sprGrid.SetActiveCell 3, iCur
    DoEvents
    
    frm의류상세.Hide
    Unload frm의류상세
End Sub

Private Sub sprGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim 의류명   As String
    Dim 의류금액 As String
    Dim iPrice   As Long
   
    If KeyCode = vbKeyReturn Then
        sprGrid.Row = sprGrid.ActiveRow
        
        sprGrid.Col = 1: 의류명 = sprGrid.Text & ""
        
        If Trim(의류명) = "" Then
            Unload Me
            
            Exit Sub
        End If
        
        sprGrid.Col = 2: 의류금액 = sprGrid.Value      '
        sprGrid.Col = 3: 의류코드 = sprGrid.Text & ""  '
        
        iPrice = Get_DryPrice(의류코드)
        
        frm접수.sprGrid.Row = iCur 'frm접수.sprGrid.ActiveRow
        frm접수.sprGrid.Col = 1: frm접수.sprGrid.Value = 의류명 & ""   '
        frm접수.sprGrid.Col = 6: frm접수.sprGrid.Value = iPrice '의류금액 & ""  '
        frm접수.sprGrid.Col = 8: frm접수.sprGrid.Value = 의류코드 & ""  '
        frm접수.sprGrid.Col = 14: frm접수.sprGrid.Value = iPrice       '14 세트 상품의 원 금액을 기록한다.
        frm접수.sprGrid.Col = 20: frm접수.sprGrid.Value = Get_세탁정상금액(의류코드)     '20 정상 금액을 기록한다.
                
        frm접수.sprGrid.SetActiveCell 3, iCur
        
        frm의류상세.Hide
        Unload frm의류상세
    End If
End Sub

Private Sub sprGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        Unload frm의류상세
    End If
End Sub
