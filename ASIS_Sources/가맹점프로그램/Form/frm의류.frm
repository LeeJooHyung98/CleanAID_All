VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm의류 
   BorderStyle     =   1  '단일 고정
   Caption         =   "의류"
   ClientHeight    =   4755
   ClientLeft      =   7095
   ClientTop       =   8910
   ClientWidth     =   13080
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   20.25
      Charset         =   129
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   13080
   ShowInTaskbar   =   0   'False
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   4755
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13080
      _ExtentX        =   23072
      _ExtentY        =   8387
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   0
      PaneTree        =   "frm의류.frx":0000
      Begin Threed.SSPanel pnlBack 
         Height          =   4755
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   13080
         _ExtentX        =   23072
         _ExtentY        =   8387
         _Version        =   262144
         BackColor       =   16777215
         BorderWidth     =   1
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin CleanAID.ctlMenu ctlMenu1 
            Height          =   825
            Index           =   0
            Left            =   75
            TabIndex        =   2
            Top             =   75
            Visible         =   0   'False
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   1455
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   20.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "frm의류"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX = 30

Private Sub curMove1()
    frm접수.sprGrid.SetActiveCell 3, iCur
End Sub

Private Sub curMove2()
    frm접수.sprGrid.SetActiveCell 4, iCur
End Sub

Private Sub curMove3()
    frm접수.sprGrid.SetActiveCell 3, iCur
End Sub

' 선택한 종류의 의류명 , 택번호, "드", 금액, 코드를 등록 한다.
Private Sub Sub_의류가격정보()
    Dim ADORs       As ADODB.RecordSet
    Dim sGoodsStats As String
    Dim iCol       As Integer  '
    Dim strNum1    As String   '택번호1
    Dim intNum1    As Integer  'tagno1
    Dim intNum2    As Integer  'tagno2
    
    Dim intCol01   As Integer  '
    Dim iActrow    As Integer  '
    Dim iPrice     As Long     '
    
    Dim 의류명     As String
    
    Dim iEOF       As Boolean
    
    
    Set ADORs = New ADODB.RecordSet
    Set ADORs = Get_의류정보(Left(의류코드, 2), sGoodsStats, frm접수.btnInternet.tag)
                      
    If ADORs.EOF Then
        ADORs.Close:    Set ADORs = Nothing
        MsgBox "등록되지 않은 품목입니다 ", vbCritical, "확인"
        Exit Sub
    End If
     
    frm접수.lblGoodsPriceStats.Caption = sGoodsStats
    의류코드 = ADORs!의류코드 & ""
    의류명 = ADORs!의류명 & ""
    
    ADORs.Close:    Set ADORs = Nothing
     
     ' 새로 입력할 라인을 구한다.
    iActrow = frm접수.sprGrid.ActiveRow
    iCur = GetSpreadLine(frm접수.sprGrid)
    
    If iCur > iActrow Then
        iCur = iActrow
    End If
    
    i = 1
    iCol = 1
    
    
    iPrice = Get_DryPrice(의류코드, frm접수.btnInternet.tag)
    
    With frm접수
        .sprGrid.Row = iCur
        .sprGrid.Col = 1:  .sprGrid.Text = Trim(의류명) & ""  ' 1 의류명
        .sprGrid.Col = 3:  .sprGrid.Text = "흰색"             ' 3 색상
        .sprGrid.Col = 4:  .sprGrid.Text = "없음"             ' 4 무늬
        
        If .chkRepair.Value = -1 Then
            .sprGrid.Col = 5: .sprGrid.Text = "수"            ' 5 작업 * 수선접수 여부 *
        Else
            .sprGrid.Col = 5: .sprGrid.Text = "세"            ' 5 작업 * 수선접수 여부 *
        End If
        
        .sprGrid.Col = 6:  .sprGrid.Value = iPrice & ""       ' 6 금액
        .sprGrid.Col = 8:  .sprGrid.Value = 의류코드 & ""     ' 8 의류코드
        .sprGrid.Col = 14: .sprGrid.Value = iPrice & ""       '14 세트 상품의 원 금액을 기록한다.
        
        '------------------------------------------------------------
        ' 마진 정보
        '------------------------------------------------------------
        Query = "SELECT    ISNULL(세탁마진,0) AS 세탁마진"
        Query = Query & ", ISNULL(외주마진,0) AS 외주마진"
        Query = Query & ", ISNULL(수선마진,0) AS 수선마진"
        Query = Query & " FROM TB_의류분류"
        Query = Query & " WHERE 의류분류코드 = '" & Left(의류코드, 2) & "'"
        Set SUBRs = New ADODB.RecordSet
        SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        If SUBRs.EOF Then
            .sprGrid.Col = 16: .sprGrid.Value = 0                   '16
            .sprGrid.Col = 17: .sprGrid.Value = 0                   '17
            .sprGrid.Col = 18: .sprGrid.Value = 0                   '18
        Else
            .sprGrid.Col = 16: .sprGrid.Value = SUBRs!세탁마진 & "" '16
            .sprGrid.Col = 17: .sprGrid.Value = SUBRs!외주마진 & "" '17
            .sprGrid.Col = 18: .sprGrid.Value = SUBRs!수선마진 & "" '18
        End If
        SUBRs.Close
        Set SUBRs = Nothing
        
        .sprGrid.Col = 20: .sprGrid.Value = Get_세탁정상금액(의류코드) & ""           '20 ** 원래 금액 **
    End With
        
    '---------------------------------------------------------------------
    ' 택 번호 출력
    '---------------------------------------------------------------------
    
    If 가맹점정보.DualComputer = "Y" Then  '2대 컴퓨터 이상 접수
        ' frm접수.sprGrid.Row = iCur
        ' frm접수.sprGrid.Col = 2: frm접수.sprGrid.Text = frm접수.cmdTagNo.Caption ' '택번호
    Else
        If frm접수.chkRepair.Value = -1 Then
            '수선접수인경우 택번호를 부여하지 않는다.
        Else
            strNum1 = frm접수.cmdTagNo.Caption '
            
            If Len(Trim(Get_SpreadText(frm접수.sprGrid, CDbl(iCur), 2))) <= 0 Then
                frm접수.sprGrid.Row = iCur
                frm접수.sprGrid.Col = 2: frm접수.sprGrid.Text = strNum1 '택번호
                
                frm접수.cmdTagNo.Caption = Get_TagNo(strNum1, "+")      '
            End If
        End If
    End If
    
    frm의류.Hide
    'Unload Me
    
    '-----------------------------------------
    frm접수.chkRepair.Enabled = False '수선접수 - 수정 못하도록...
    
    frm접수.sprGrid.Row = iCur
    frm접수.sprGrid.BackColor = vbWhite
    frm접수.sprGrid.SetActiveCell 3, iCur
    DoEvents
    
    'frm색상표.Show 1   '1998/01/10 수정
    frm색상표.Show    '2010-11-25
End Sub

Private Sub ctlMenu1_Click(Index As Integer)
    의류코드 = ctlMenu1(Index).GET_MenuKey
    
    If 의류코드 = "" Then Exit Sub
    
    의류코드 = 의류코드 & "00" '** 의류분류코드 이기 때문에 뒤에 '00'을 붙여서 4자리 의류코드로 변환해준다. **
    
    Call Sub_의류가격정보
    Call curMove1
    
    '-----------------------------------------------------------
    '상의, 코트류 - 한벌...
    '-----------------------------------------------------------
    If Left(의류코드, 1) = "f" Then
        frm접수.cmdSuite.Enabled = True
    Else
        frm접수.cmdSuite.Enabled = False
    End If
End Sub

Private Sub ctlMenu1_GotFocus(Index As Integer)
    ctlMenu1(Index).BackColor = &HC0FFFF
End Sub

Private Sub ctlMenu1_LostFocus(Index As Integer)
    ctlMenu1(Index).BackColor = 0
End Sub


Private Sub Form_Activate()
    On Error GoTo ErrRtn
    
    frm의류.Top = frmMain.Top   '390
    frm의류.Left = frmMain.Left '200
    
    If Me.Visible = False Then Exit Sub
        
    If ctlMenu1.Count > 1 Then
        ctlMenu1(1).SetFocus
    End If
'    sprCloth.SetFocus
'
'    sprCloth.SetActiveCell 3, 2
'    sprCloth.TypeButtonColor = "&H00C0FFFF"
    Exit Sub
    
ErrRtn:
    Call Error_Msg("Form_Activate frm의류", Err.Source, Err.Number, Err.description)

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        frm의류.Hide
        'Unload Me
    End If
End Sub

Private Sub Form_Load()

    On Error GoTo ErrRtn
    
    frm의류.Top = frmMain.Top   '390
    frm의류.Left = frmMain.Left '200
        
    If ctlMenu1.Count > 1 Then
        For i = 1 To ctlMenu1.Count - 1
            Unload ctlMenu1(i)
        Next i
    End If
    
    i = 1
    
    '----------------------------------------------------------
    ' TB_의류분류
    '----------------------------------------------------------
    Query = "SELECT    의류분류코드"
    Query = Query & ", 의류분류명"
    Query = Query & ", 순서"
    Query = Query & " FROM TB_의류분류"
    Query = Query & " ORDER BY 순서"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    Do Until ADORs.EOF
        Load ctlMenu1(i)
        ctlMenu1(i).Left = GetLeft(i)
        ctlMenu1(i).Top = GetTop(i)
        
        Call ctlMenu1(i).SET_Item(ADORs!의류분류명, 0, ADORs!의류분류코드, "")
        
        ctlMenu1(i).Enabled = True
        ctlMenu1(i).Visible = True
           
           
        If UCase(Left(ADORs!의류분류코드, 1)) = "W" And 가맹점정보.지사코드 = "1024" Then
        
            ' 2013-10-21일부터 일반가죽을 크렌즈에서 사용하지 못하도록 설정
            If Format(Date, "yyyy-MM-dd") >= "2013-10-21" Then
                ctlMenu1(i).Enabled = False
                Call ctlMenu1(i).SET_Item("", 0, "", "")
            End If
        End If
           
           
        i = i + 1
        
        ADORs.MoveNext
    Loop
    ADORs.Close
    Set ADORs = Nothing
    
    If i > 0 Then
        Me.Height = ctlMenu1(i - 1).Top + ctlMenu1(i - 1).Height + 550
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("Form_Load frm의류", Err.Source, Err.Number, Err.description)
    
End Sub

Private Function GetLeft(ByVal Locate As Integer) As Long
    On Error Resume Next
    
    Select Case Locate
        Case 1, 8, 15, 22, 29, 36, 43, 50, 57, 64
            GetLeft = 45
        Case 2, 9, 16, 23, 30, 37, 44, 51, 58, 65
            GetLeft = 1905
        Case 3, 10, 17, 24, 31, 38, 45, 52, 59, 66
            GetLeft = 3765
        Case 4, 11, 18, 25, 32, 39, 46, 53, 60, 67
            GetLeft = 5625
        Case 5, 12, 19, 26, 33, 40, 47, 54, 61, 68
            GetLeft = 7485
        Case 6, 13, 20, 27, 34, 41, 48, 55, 62, 69
            GetLeft = 9345
        Case 7, 14, 21, 28, 35, 42, 49, 56, 63, 70
            GetLeft = 11205
        Case Else
            GetLeft = 0
    End Select
End Function

Private Function GetTop(ByVal Locate As Integer) As Long
    On Error Resume Next
    
    Select Case Locate
        Case 1, 2, 3, 4, 5, 6, 7
            GetTop = 45
        Case 8, 9, 10, 11, 12, 13, 14
            GetTop = 945 '915
        Case 15, 16, 17, 18, 19, 20, 21
            GetTop = 1845 '1785
        Case 22, 23, 24, 25, 26, 27, 28
            GetTop = 2745 '2655
        Case 29, 30, 31, 32, 33, 34, 35
            GetTop = 3645 '3530
        Case 36, 37, 38, 39, 40, 41, 42
            GetTop = 4545 '4500
        Case 43, 44, 45, 46, 47, 48, 49
            GetTop = 5445 '5370
        Case 50, 51, 52, 53, 54, 55, 56
            GetTop = 6345 '6240
        Case 57, 58, 59, 60, 61, 62, 63
            GetTop = 7245 '7110
        Case 64, 65, 66, 67, 68, 69, 70
            GetTop = 8145 '7980
        Case Else
            GetTop = 0
    End Select
End Function

