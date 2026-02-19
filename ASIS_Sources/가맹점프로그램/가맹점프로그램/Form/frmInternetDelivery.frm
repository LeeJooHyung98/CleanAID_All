VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmInternetDelivery 
   BorderStyle     =   1  '단일 고정
   Caption         =   "인터넷 출고현황"
   ClientHeight    =   4605
   ClientLeft      =   9375
   ClientTop       =   4155
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   8130
   Begin XtremeSuiteControls.PushButton btnAccept 
      Height          =   420
      Left            =   5970
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4140
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
      Picture         =   "frmInternetDelivery.frx":0000
   End
   Begin FPSpreadADO.fpSpread sprGrid 
      Height          =   4005
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   8010
      _Version        =   524288
      _ExtentX        =   14129
      _ExtentY        =   7064
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
      MaxCols         =   6
      MaxRows         =   1
      OperationMode   =   2
      ScrollBars      =   2
      SpreadDesigner  =   "frmInternetDelivery.frx":0A12
      UserResize      =   1
      VisibleCols     =   3
      HighlightHeaders=   1
      HighlightStyle  =   1
      ScrollBarStyle  =   2
   End
   Begin XtremeSuiteControls.PushButton btnCancel 
      Height          =   420
      Left            =   7050
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4140
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
      Picture         =   "frmInternetDelivery.frx":113B
   End
End
Attribute VB_Name = "frmInternetDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SELECTCODE As String

Private Sub btnAccept_Click()
    
    sprGrid.Row = sprGrid.ActiveRow
    sprGrid.Col = 1
    Call CheckMember(sprGrid.Text)
    frm출고.btnInternet.Tag = sprGrid.Text
    
    Unload Me
End Sub

Private Sub btnCancel_Click()
    frm출고.btnInternet.Tag = ""
    
    Unload Me
End Sub

Private Sub Form_Load()
    sprGrid.MaxRows = 0
End Sub

Public Sub GetData()
    Dim ADORs As ADODB.RecordSet
    Set ADORs = GetInternetDelivery()
    Do Until ADORs.EOF
        With sprGrid
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 1: .Text = ADORs!주문번호
            .Col = 2: .Text = ADORs!이름
            .Col = 3: .Text = ADORs!의류
            .Col = 4: .Text = ADORs!신발
            .Col = 5: .Text = ADORs!이불
            .Col = 6: .Text = ADORs!수거일자
        End With
        
        ADORs.MoveNext
    Loop
    ADORs.Close
End Sub

Private Sub sprGrid_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    sprGrid.Row = Row
    sprGrid.Col = 1
    Call CheckMember(sprGrid.Text)
    frm출고.btnInternet.Tag = sprGrid.Text
    
    Unload Me
End Sub


Private Sub CheckMember(Search As String)
    Query = "SELECT * FROM TB_고객정보"
    Query = Query & " WHERE 전화번호 = '" & Search & "'"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenDynamic, adLockOptimistic
        
    If ADORs.EOF Then
        ADORs.AddNew
    
        ADORs!고객코드 = Get_CustomNo & ""              ' 1
        ADORs!등록일자 = Format(Date, "YYYY-MM-DD")           ' 2
        ADORs!수정일자 = ""                                   ' 3
        ADORs!이용횟수 = 0                                    ' 4
        ADORs!총접수금액 = 0                                  ' 5
        ADORs!삭제 = 0                                        ' 8
        ADORs!최종거래일자 = ""                               ' 9
        ADORs!성명 = "인터넷고객" & ""           '12
        ADORs!전화번호 = Search & ""                         '13
        ADORs!휴대전화 = Search & ""                          '14
        ADORs!주소 = "" & ""            '15
        ADORs!미수금액 = 0                           '16
        ADORs!카드번호 = "" & ""                        '17
        ADORs!문자발송여부 = "N"                 '18
        ADORs!메모 = "" & ""                 '19
        ADORs!고객등급코드 = "I"               '20
        ADORs!본사전송여부 = "N"                                  '21
        ADORs!지사코드 = 가맹점정보.지사코드 & ""                 '22
        ADORs!가맹점코드 = 가맹점정보.가맹점코드 & ""             '23
    Else
        
    End If
    SELECTCODE = ADORs!고객코드
    ADORs.Update
    
    ADORs.Close
    Set ADORs = Nothing
End Sub

Public Function Get_CustomNo() As String
    Dim sYEAR As String
    
    sYEAR = "9"
    
    Query = "SELECT ISNULL(MAX(고객코드),'" & sYEAR & "00000') + 1 FROM TB_고객정보"
    Query = Query & " WHERE SUBSTRING(고객코드,1,1) = '" & sYEAR & "'"
    Set SUBRs = New ADODB.RecordSet
    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    Get_CustomNo = Format(SUBRs(0), "000000")

    SUBRs.Close
    Set SUBRs = Nothing
End Function
