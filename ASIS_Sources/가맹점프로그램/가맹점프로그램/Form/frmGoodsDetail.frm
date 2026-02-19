VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmGoodsDetail 
   BorderStyle     =   1  '단일 고정
   Caption         =   "상품선택"
   ClientHeight    =   9585
   ClientLeft      =   11370
   ClientTop       =   4110
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   8130
   Begin FPSpreadADO.fpSpread sprGrid_Color 
      Height          =   8940
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   75
      Width           =   3915
      _Version        =   524288
      _ExtentX        =   6906
      _ExtentY        =   15769
      _StockProps     =   64
      BackColorStyle  =   1
      DAutoCellTypes  =   0   'False
      DAutoHeadings   =   0   'False
      DAutoSave       =   0   'False
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
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
      MaxCols         =   3
      MaxRows         =   1
      OperationMode   =   2
      ScrollBars      =   2
      SpreadDesigner  =   "frmGoodsDetail.frx":0000
      UserResize      =   1
      VisibleCols     =   2
      HighlightHeaders=   1
      HighlightStyle  =   1
      ScrollBarStyle  =   2
   End
   Begin FPSpreadADO.fpSpread sprGrid_Pattern 
      Height          =   8940
      Left            =   4110
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   75
      Width           =   3915
      _Version        =   524288
      _ExtentX        =   6906
      _ExtentY        =   15769
      _StockProps     =   64
      BackColorStyle  =   1
      DAutoCellTypes  =   0   'False
      DAutoHeadings   =   0   'False
      DAutoSave       =   0   'False
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
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
      MaxCols         =   2
      MaxRows         =   6
      OperationMode   =   2
      ScrollBars      =   2
      SpreadDesigner  =   "frmGoodsDetail.frx":05BC
      UserResize      =   1
      VisibleCols     =   2
      HighlightHeaders=   1
      HighlightStyle  =   1
      ScrollBarStyle  =   2
   End
   Begin XtremeSuiteControls.PushButton btnAccept 
      Height          =   420
      Left            =   5910
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   9090
      Width           =   1035
      _Version        =   851970
      _ExtentX        =   1826
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   " 선택"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Picture         =   "frmGoodsDetail.frx":4C59
   End
   Begin XtremeSuiteControls.PushButton btnCancel 
      Height          =   420
      Left            =   6990
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   9090
      Width           =   1035
      _Version        =   851970
      _ExtentX        =   1826
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   " 취소"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Picture         =   "frmGoodsDetail.frx":566B
   End
End
Attribute VB_Name = "frmGoodsDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub GetData(Color As String, Pattern As String)
    Dim LoopI As Long
    sprGrid_Color.MaxRows = 0
    
    With sprGrid_Pattern
        
        .Row = 1: .Col = 2: .Text = "없음"
        .Row = 2: .Col = 2: .Text = "가로"
        .Row = 3: .Col = 2: .Text = "세로"
        .Row = 4: .Col = 2: .Text = "체크"
        .Row = 5: .Col = 2: .Text = "혼합"
        .Row = 6: .Col = 2: .Text = "기타"

    End With
    
    
    Query = "SELECT    색상코드"
    Query = Query & ", 색상명"
    Query = Query & ", RGB"
    Query = Query & ", 순서"
    Query = Query & " FROM TB_색상표"
    Query = Query & " WHERE 색상명 <> '품목보기'"
    Query = Query & " ORDER BY 순서"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    Do Until ADORs.EOF
        With sprGrid_Color
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 1
            .BackColor = IIf(Right(ADORs!RGB, 1) = "&", Left(ADORs!RGB, Len(ADORs!RGB) - 1), ADORs!RGB)
            .Col = 2
            .Text = ADORs!색상명
            .Col = 3
            .Text = IIf(Right(ADORs!RGB, 1) = "&", Left(ADORs!RGB, Len(ADORs!RGB) - 1), ADORs!RGB)
        End With
        ADORs.MoveNext
    Loop
    ADORs.Close
    Set ADORs = Nothing
    
    For LoopI = 1 To sprGrid_Color.MaxRows
        With sprGrid_Color
            .Row = LoopI: .Col = 2
            If .Text = Color Then
                Call .SetActiveCell(1, LoopI)
                Exit For
            End If
        End With
    Next LoopI
    
    
    For LoopI = 1 To sprGrid_Pattern.MaxRows
        With sprGrid_Pattern
            .Row = LoopI: .Col = 2
            If .Text = Pattern Then
                Call .SetActiveCell(1, LoopI)
                Exit For
            End If
        End With
    Next LoopI
    
    
End Sub

Private Sub btnAccept_Click()
    Dim Pattern As String
    Dim Color As String
    Dim ColorCode As String
    sprGrid_Pattern.Row = sprGrid_Pattern.ActiveRow
    sprGrid_Pattern.Col = 2
    Pattern = sprGrid_Pattern.Text
    
    sprGrid_Color.Row = sprGrid_Color.ActiveRow
    sprGrid_Color.Col = 2
    Color = sprGrid_Color.Text
    sprGrid_Color.Col = 3
    ColorCode = sprGrid_Color.Text
    Call frmAccept.SetDetail(Pattern, Color, ColorCode)
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

