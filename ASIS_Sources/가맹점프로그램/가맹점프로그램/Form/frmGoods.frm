VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmGoods 
   BorderStyle     =   1  '단일 고정
   Caption         =   "상품선택"
   ClientHeight    =   9585
   ClientLeft      =   9540
   ClientTop       =   5595
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   8130
   Begin XtremeSuiteControls.PushButton btnAccept 
      Height          =   420
      Left            =   5955
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   9075
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
      Picture         =   "frmGoods.frx":0000
   End
   Begin FPSpreadADO.fpSpread sprGrid 
      Height          =   8940
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   8010
      _Version        =   524288
      _ExtentX        =   14129
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
      SpreadDesigner  =   "frmGoods.frx":0A12
      UserResize      =   1
      VisibleCols     =   3
      HighlightHeaders=   1
      HighlightStyle  =   1
      ScrollBarStyle  =   2
   End
   Begin XtremeSuiteControls.PushButton btnCancel 
      Height          =   420
      Left            =   7035
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   9075
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
      Picture         =   "frmGoods.frx":1030
   End
End
Attribute VB_Name = "frmGoods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAccept_Click()
    With sprGrid
    .Row = .ActiveRow
    .Col = 1
    frmAccept.Sub_의류가격정보 (.Text)
    End With
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    sprGrid.MaxRows = 0
End Sub

Public Sub GetData(Search As String)
    Dim ADORs As ADODB.RecordSet
    Set ADORs = GetGoodsSub(Search)
    Do Until ADORs.EOF
        With sprGrid
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 1: .Text = ADORs!의류코드
            .Col = 2: .Text = ADORs!의류명
            .Col = 3: .Text = ADORs!금액
        End With
        
        ADORs.MoveNext
    Loop
    ADORs.Close
End Sub

Private Sub sprGrid_DblClick(ByVal Col As Long, ByVal Row As Long)
    With sprGrid
    .Row = Row
    .Col = 1
    frmAccept.Sub_의류가격정보 (.Text)
    End With
    Unload Me
End Sub
