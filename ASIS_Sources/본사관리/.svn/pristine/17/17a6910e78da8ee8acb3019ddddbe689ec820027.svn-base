VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form P_03003_01 
   BorderStyle     =   1  '단일 고정
   Caption         =   "입고내역조회"
   ClientHeight    =   3300
   ClientLeft      =   7545
   ClientTop       =   1740
   ClientWidth     =   8205
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_03003_01.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   8205
   StartUpPosition =   1  '소유자 가운데
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   3300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   5821
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_03003_01.frx":058A
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   3270
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   8175
         _Version        =   524288
         _ExtentX        =   14420
         _ExtentY        =   5768
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
         MaxCols         =   5
         ScrollBars      =   2
         SpreadDesigner  =   "P_03003_01.frx":05BC
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_03003_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public l_AgencyCode As String
Public l_TagNo As String
Public l_IpDate As String

Dim sValue() As String
Dim RS01 As ADODB.Recordset

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    With spdView
        .MaxRows = 1
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
    
    
    ReDim sValue(1)
    
    sValue(0) = l_AgencyCode
    sValue(1) = l_TagNo
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_03003_02", sValue(), Err_Num, Err_Dec)
    
    With spdView
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = RS01!일자 & "" '
            .Col = 2: .Text = RS01!품목 & "" '
            .Col = 3: .Text = RS01!색상 & "" '
            .Col = 4: .Text = RS01!내용 & "" '
            .Col = 5: .Text = RS01!상표 & "" '
            
            RS01.MoveNext
        Loop
        
        RS01.Close
        Set RS01 = Nothing
        
        .Redraw = True
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub spdView_DblClick(ByVal Col As Long, ByVal Row As Long)
    If spdView.MaxRows < 1 Then Exit Sub
    
    spdView.Row = spdView.ActiveRow
    spdView.Col = 1
    
    l_IpDate = Format(spdView.Text, "YYYY-MM-DD")
    
    Unload Me
End Sub

Private Sub spdView_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call spdView_DblClick(spdView.ActiveCol, spdView.ActiveRow)
        
    ElseIf KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub
