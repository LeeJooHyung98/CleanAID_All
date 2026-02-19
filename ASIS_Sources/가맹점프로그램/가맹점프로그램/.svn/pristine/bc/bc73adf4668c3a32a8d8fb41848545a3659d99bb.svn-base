VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frm수선 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   ClientHeight    =   960
   ClientLeft      =   5400
   ClientTop       =   2610
   ClientWidth     =   3795
   ControlBox      =   0   'False
   LinkTopic       =   "Form88"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   3795
   StartUpPosition =   1  '소유자 가운데
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   780
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   3645
      _Version        =   524288
      _ExtentX        =   6429
      _ExtentY        =   1376
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "맑은 고딕"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   2
      MaxRows         =   1
      ScrollBars      =   0
      SpreadDesigner  =   "frm수선.frx":0000
      UserResize      =   1
      VisibleCols     =   2
      VisibleRows     =   1
      HighlightHeaders=   1
      HighlightStyle  =   1
   End
End
Attribute VB_Name = "frm수선"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Form_Load()
    fpSpread1.CursorStyle = 2
    
    Query = "SELECT * FROM TB_수선금액 "
    Query = Query & " WHERE 수선내용 =  '짜집기(cm당)'"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF Then
        '
    Else
        fpSpread1.Row = 1
        fpSpread1.Col = 1: fpSpread1.Value = ADORs!수선내용 & ""
        fpSpread1.Col = 2: fpSpread1.Value = ADORs!금액 & ""
    End If
    ADORs.Close
    Set ADORs = Nothing
End Sub

Private Sub fpSpread1_Click(ByVal Col As Long, ByVal Row As Long)
    '기존금액과 상표내용이 변화한다
    Dim strRep   As String
    Dim strPrice As String
    
    If Row < 1 Then
       Exit Sub
    End If
    
    fpSpread1.Row = Row
    fpSpread1.Col = 1: strRep = Trim(fpSpread1.Value)
    
    If Len(strRep) < 1 Then
       Exit Sub
    End If
    
    fpSpread1.Col = 2: strPrice = Trim(fpSpread1.Value)
    
    'frm작업.pnlMoney.Visible = True
    
    Load frm작업
    frm작업.Show
    
    frm접수.sprGrid.Row = frm접수.sprGrid.ActiveRow 'iCur
    frm접수.sprGrid.Col = 7: frm접수.sprGrid.Value = strRep
    frm접수.sprGrid.Col = 6: frm접수.sprGrid.Value = CDbl(Val(strPrice))
    frm접수.cmdOK.SetFocus
    
    Unload Me
End Sub

Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    '기존금액과 상표내용이 변화한다
    Dim strRep As String
    Dim strPrice As String
    
    If KeyCode = 13 Then
        fpSpread1.Row = fpSpread1.ActiveRow
        fpSpread1.Col = 1: strRep = Trim(fpSpread1.Value)
        
        If Len(strRep) < 1 Then
            Exit Sub
        End If
        
        fpSpread1.Col = 2: strPrice = Trim(fpSpread1.Value)
        
        'frm작업.pnlMoney.Visible = True
        
        Load frm작업
        frm작업.Show
        
        frm접수.sprGrid.Row = iCur
        frm접수.sprGrid.Col = 7: frm접수.sprGrid.Value = strRep
        frm접수.sprGrid.Col = 6: frm접수.sprGrid.Value = CDbl(strPrice)
        
        frm접수.cmdOK.SetFocus
        
        Unload Me
    End If
End Sub

Private Sub fpSpread1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
        KeyAscii = 0
    End If
End Sub

Private Sub fpSpread1_LostFocus()
    Unload Me
End Sub
