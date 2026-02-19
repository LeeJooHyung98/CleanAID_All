VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm세탁수선 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '단일 고정
   Caption         =   "세탁수선"
   ClientHeight    =   7650
   ClientLeft      =   3150
   ClientTop       =   1380
   ClientWidth     =   4290
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   4290
   StartUpPosition =   1  '소유자 가운데
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7650
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   13494
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm세탁수선.frx":0000
      Begin FPSpreadADO.fpSpread sprRepair 
         Height          =   7620
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   4260
         _Version        =   524288
         _ExtentX        =   7514
         _ExtentY        =   13441
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowUserFormulas=   -1  'True
         BackColorStyle  =   1
         ColsFrozen      =   2
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
         MaxCols         =   2
         MaxRows         =   30
         Protect         =   0   'False
         RestrictCols    =   -1  'True
         ScrollBars      =   2
         SpreadDesigner  =   "frm세탁수선.frx":0032
         VisibleCols     =   2
         VisibleRows     =   30
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "frm세탁수선"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    Query = "SELECT    수선내용"
    Query = Query & ", 금액 "
    Query = Query & " FROM TB_수선금액"
    Query = Query & " ORDER BY 수선내용 ASC"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprRepair
        .CursorStyle = 2
    
        .MaxRows = 0
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = ADORs!수선내용 & ""
            .Col = 2: .Text = ADORs!금액 & ""
        
            ADORs.MoveNext
        Loop
        
        ADORs.Close
        Set ADORs = Nothing
    End With
    
    Exit Sub
    
ErrRtn:

End Sub

Private Sub sprRepair_Click(ByVal Col As Long, ByVal Row As Long)
    '기존금액과 상표내용이 변화한다
    Dim 수선내용    As String
    Dim 수선금액    As Long
    Dim temp01      As Long
    Dim sGoodsStats As String
    
    Dim ClothCode   As String '의류코드
    
    If Row < 1 Then Exit Sub
    
    sprRepair.Row = Row
    sprRepair.Col = 1: 수선내용 = Trim(sprRepair.Text) & "" '수선내용
    sprRepair.Col = 2: 수선금액 = sprRepair.Value & ""      '수선금액
    
    If Trim(수선내용) = "" Then Exit Sub
    
    If Trim(수선내용) = "짜집기(cm당)" Then
       'frm작업.pnlMoney.Visible = True
       
       frm작업.Show
    End If
    
    With frm접수.sprGrid
        .Row = .ActiveRow 'iCur
        .Col = 7: .Value = 수선내용
    
        .Col = 5
        
        If Trim(.Text) = "수" Then
            .Col = 6: .Value = 수선금액 & ""  '수선금액
            .Col = 9: .Value = 수선금액 & ""  '수선 금액 입력
        
        Else
            .Col = 8: ClothCode = Trim(.Text) & "" '의류코드
            
            temp01 = Get_세탁금액(ClothCode, sGoodsStats)      '세탁금액
            
            .Col = 5
            If Mid(Trim(.Text), 2, 1) = "아" Then
                '아동복일 경우 20% 할인한다.
                
                temp01 = CStr(Int((CLng(temp01) * 0.8) / 100) * 100) ' 10원단위를 절사 한다.
            End If
            
            .Col = 6: .Value = (temp01 + 수선금액) & "" '금액
            .Col = 9: .Value = 수선금액 & ""            '수선 금액 입력
        End If
    End With
    
    Unload Me
    
    frm접수.cmdOK.SetFocus
End Sub

Private Sub sprRepair_KeyDown(KeyCode As Integer, Shift As Integer)
    '기존금액과 상표내용이 변화한다
    Dim strRep As String
    Dim strPrice As String
    
    If KeyCode = 13 Then
        sprRepair.Row = sprRepair.ActiveRow
        sprRepair.Col = 1: strRep = Trim(sprRepair.Value)
        
        If Len(strRep) < 1 Then Exit Sub
            
        sprRepair.Col = 2: strPrice = Val(Trim(sprRepair.Value))
        
        ' ------
        If Trim(strRep) = "짜집기(cm당)" Then
            'frm작업.pnlMoney.Visible = True
            Load frm작업
            frm작업.Show
        End If
        
        frm접수.sprGrid.Row = iCur
        frm접수.sprGrid.Col = 7: frm접수.sprGrid.Value = strRep
        frm접수.sprGrid.Col = 5
        
        If Trim(frm접수.sprGrid.Value) = "수" Then
            frm접수.sprGrid.Col = 6: frm접수.sprGrid.Value = CDbl(strPrice)
        Else
            frm접수.sprGrid.Col = 6: frm접수.sprGrid.Value = CStr(CDbl(frm접수.sprGrid.Value) + CDbl(strPrice))
        End If
        
        Unload Me
        frm접수.cmdOK.SetFocus
    End If
End Sub

Private Sub sprRepair_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
        KeyAscii = 0
    End If
End Sub

Private Sub sprRepair_LostFocus()
    Unload Me
End Sub
