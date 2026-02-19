VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form frm참조코드 
   ClientHeight    =   8175
   ClientLeft      =   3180
   ClientTop       =   4005
   ClientWidth     =   11670
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form11"
   MDIChild        =   -1  'True
   ScaleHeight     =   8175
   ScaleWidth      =   11670
   WindowState     =   2  '최대화
   Begin VB.CommandButton cmdSave 
      Caption         =   "저장"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   20.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   9510
      TabIndex        =   5
      Top             =   2745
      Width           =   1305
   End
   Begin VB.CommandButton Command1 
      Caption         =   "↓"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   48
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Index           =   1
      Left            =   9510
      TabIndex        =   4
      Top             =   4530
      Width           =   1305
   End
   Begin VB.CommandButton Command1 
      Caption         =   "↑"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   48
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Index           =   0
      Left            =   9510
      TabIndex        =   3
      Top             =   960
      Width           =   1305
   End
   Begin VB.ComboBox DBCombo1 
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3135
      Style           =   2  '드롭다운 목록
      TabIndex        =   2
      Top             =   510
      Width           =   3435
   End
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   5445
      Left            =   1035
      TabIndex        =   0
      Top             =   960
      Width           =   8445
      _Version        =   524288
      _ExtentX        =   14896
      _ExtentY        =   9604
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      DAutoCellTypes  =   0   'False
      DAutoHeadings   =   0   'False
      DAutoSave       =   0   'False
      DAutoSizeCols   =   0
      DInformActiveRowChange=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   4
      MaxRows         =   40
      OperationMode   =   3
      Protect         =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "frm참조코드.frx":0000
      VisibleCols     =   4
      VisibleRows     =   30
      AppearanceStyle =   0
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   390
      Index           =   2
      Left            =   1035
      TabIndex        =   1
      Top             =   510
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   688
      _Version        =   262144
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
      Caption         =   "코드품명선택"
      BevelWidth      =   2
      RoundedCorners  =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "품목보기의 순번을 변경 가능 합니다."
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1050
      TabIndex        =   6
      Top             =   120
      Width           =   5865
   End
End
Attribute VB_Name = "frm참조코드"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strdate01 As String
Dim strdate02 As String

Private Sub clear01()
    With fpSpread1
        For i = 1 To .MaxRows
            .Row = i
            .Col = 5
            .Action = 3
        Next i
    End With
End Sub
 
Private Sub GetGoodsTopCode()
    Query = "SELECT 의류코드 +' : '+ 의류명 as aaa  "
    Query = Query & " FROM TB_의류 "
    Query = Query & " WHERE SUBSTRING(의류코드,2,2) = '00'"
    Query = Query & " ORDER BY 의류코드 ASC "
    
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    If ADORs.RecordCount < 1 Then
        Exit Sub
    Else
        ADORs.MoveFirst
        'DBCombo1.Text = ADORs!aaa
        
        While Not ADORs.EOF And Not ADORs.BOF
            DBCombo1.AddItem ADORs!aaa
            ADORs.MoveNext
        Wend
        
        If DBCombo1.ListCount > 0 Then DBCombo1.ListIndex = 0
    
    End If

End Sub

 
Private Sub cmdSave_Click()
    Dim nRow As Long
    Dim varTemp As Variant
    Dim sData(1) As String
    
    With fpSpread1
        For nRow = 1 To .MaxRows
            .GetText 1, nRow, varTemp
            sData(0) = Trim(CStr(varTemp))
            
            .GetText 4, nRow, varTemp
            sData(1) = Trim(CStr(varTemp))
            
            If sData(0) <> "" And sData(1) <> "" Then
                Query = "UPDATE TB_의류 SET "
                Query = Query & " 출력순번  = '" & sData(1) & "' "
                Query = Query & " WHERE 의류코드 = '" & sData(0) & "' "
                ADOCon.Execute Query
                
                Query = "UPDATE TB_할인정보 SET "
                Query = Query & " 출력순번  = '" & sData(1) & "' "
                Query = Query & " WHERE 의류코드 = '" & sData(0) & "' "
                ADOCon.Execute Query
                
                Query = "UPDATE TB_목요세일 SET "
                Query = Query & " 출력순번  = '" & sData(1) & "' "
                Query = Query & " WHERE 의류코드 = '" & sData(0) & "' "
                ADOCon.Execute Query
            Else
                MsgBox "내용을 확인 하여 주십시요" & vbNewLine & "[" & sData(0) & "," & sData(1) & "]", vbInformation, "확인"
                Exit Sub
            End If
        Next nRow
    End With
    
    MsgBox "저장 완료   ", vbInformation, "확인"

End Sub
 

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            Call MoveSpreadDataUpDown(Index)
            Exit Sub
            
        Case 1
            Call MoveSpreadDataUpDown(Index)
            Exit Sub
    End Select
            
End Sub

Private Sub DBCombo1_Click()
    Dim nRow As Long
    
    Call clear01
    
    Query = "SELECT 의류코드, 의류명, 금액, 출력순번 "
    Query = Query & " FROM TB_의류 "
    Query = Query & " WHERE 의류코드 = 의류코드 "
    Query = Query & " AND SUBSTRING(의류코드,1,1) LIKE '" + Mid(DBCombo1.Text, 1, 1) + "%'"
    Query = Query & " ORDER BY 출력순번 ASC, 의류코드 ASC "
    
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    nRow = 0
    Do While Not ADORs.EOF
        nRow = nRow + 1
        fpSpread1.MaxRows = nRow
        fpSpread1.Row = nRow
        
        fpSpread1.Col = 1
        If Not IsNull(ADORs!의류코드) Then fpSpread1.Text = ADORs!의류코드
        fpSpread1.Col = 2
        If Not IsNull(ADORs!의류명) Then fpSpread1.Text = ADORs!의류명
        fpSpread1.Col = 3
        If Not IsNull(ADORs!금액) Then fpSpread1.Text = ADORs!금액
        fpSpread1.Col = 4
        If Not IsNull(ADORs!출력순번) Then
            If Trim(ADORs!출력순번 & "") = "" Then
                fpSpread1.Text = UCase(Mid(DBCombo1.Text, 1, 1)) & Format(nRow, "000")
            Else
                fpSpread1.Text = ADORs!출력순번
            End If
        Else
            fpSpread1.Text = UCase(Mid(DBCombo1.Text, 1, 1)) & Format(nRow, "000")
        End If
        
        ADORs.MoveNext
    Loop
    
 End Sub
 
Private Sub Form_Activate()
    Call GetGoodsTopCode
    
    Call DBCombo1_Click
End Sub

Private Sub Form_Load()
    'TitleSet "품목보기순위조정"
End Sub
 
Private Sub MoveSpreadDataUpDown(Mode As Integer)
    Dim nSelRow     As Long
    Dim sData(3)    As String
    Dim varTemp     As Variant
    Dim nActCnt    As Integer
    
    With fpSpread1
    
        nSelRow = .ActiveRow
        nActCnt = IIf(Mode = 0, 1, -1)
        
 
        ' 업일때는 가장 첫줄이면 아무런 동작을 하지 않는다.
        If Mode = 0 And nSelRow <= 1 Then
            Exit Sub
        ' 다운일경우 가장 마지막일 경우 아무런 동작을 하지 않느다.
        ElseIf Mode = 1 And nSelRow >= .MaxRows Then
            Exit Sub
        End If
        
    
        ' 코드
        .GetText 1, nSelRow, varTemp
        sData(0) = Trim(CStr(varTemp))
        
        ' 의류명
        .GetText 2, nSelRow, varTemp
        sData(1) = Trim(CStr(varTemp))

        ' 금액
        .GetText 3, nSelRow, varTemp
        sData(2) = Trim(CStr(varTemp))
        
        ' 위쪽 내역을 아래로 내린다.
        ' 코드, 의류명, 금액
        .GetText 1, nSelRow - nActCnt, varTemp
        .SetText 1, nSelRow, varTemp
        .GetText 2, nSelRow - nActCnt, varTemp
        .SetText 2, nSelRow, varTemp
        .GetText 3, nSelRow - nActCnt, varTemp
        .SetText 3, nSelRow, varTemp
        
        .SetText 1, nSelRow - nActCnt, CVar(sData(0))
        .SetText 2, nSelRow - nActCnt, CVar(sData(1))
        .SetText 3, nSelRow - nActCnt, CVar(sData(2))
        
        .Row = nSelRow - nActCnt
        .Action = ActionActiveCell
    
    End With

End Sub
