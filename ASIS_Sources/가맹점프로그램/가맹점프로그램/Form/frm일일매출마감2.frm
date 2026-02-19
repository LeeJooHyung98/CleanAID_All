VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm일일매출마감2 
   Caption         =   "일일매출마감"
   ClientHeight    =   12510
   ClientLeft      =   1710
   ClientTop       =   3375
   ClientWidth     =   20310
   ControlBox      =   0   'False
   Icon            =   "frm일일매출마감2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   ScaleHeight     =   12510
   ScaleWidth      =   20310
   StartUpPosition =   2  '화면 가운데
   Begin Threed.SSPanel pnlProg 
      Height          =   1215
      Left            =   180
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   2143
      _Version        =   262144
      BackColor       =   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frm일일매출마감2.frx":0A02
      Caption         =   " "
      BorderWidth     =   0
      BevelOuter      =   1
      BevelInner      =   2
      Alignment       =   4
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   12510
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20310
      _ExtentX        =   35825
      _ExtentY        =   22066
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm일일매출마감2.frx":39CD
      Begin Threed.SSPanel SSPanel 
         Height          =   11280
         Index           =   1
         Left            =   15
         TabIndex        =   11
         Top             =   1215
         Width           =   20280
         _ExtentX        =   35772
         _ExtentY        =   19897
         _Version        =   262144
         BackColor       =   16777215
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSFrame SSFrame1 
            Height          =   7080
            Left            =   90
            TabIndex        =   12
            Top             =   105
            Width           =   9900
            _ExtentX        =   17463
            _ExtentY        =   12488
            _Version        =   262144
            BackColor       =   16777215
            Begin FPSpreadADO.fpSpread sprGrid 
               Height          =   6945
               Index           =   0
               Left            =   60
               TabIndex        =   13
               Top             =   60
               Width           =   5280
               _Version        =   524288
               _ExtentX        =   9313
               _ExtentY        =   12250
               _StockProps     =   64
               BackColorStyle  =   1
               BorderStyle     =   0
               DisplayColHeaders=   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   12
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               GrayAreaBackColor=   16777215
               MaxCols         =   4
               MaxRows         =   16
               Protect         =   0   'False
               RowHeaderDisplay=   0
               ScrollBars      =   0
               SpreadDesigner  =   "frm일일매출마감2.frx":3A3F
               UserResize      =   0
               HighlightHeaders=   1
               HighlightStyle  =   1
            End
            Begin FPSpreadADO.fpSpread sprGrid 
               Height          =   6945
               Index           =   1
               Left            =   5340
               TabIndex        =   14
               Top             =   60
               Width           =   4500
               _Version        =   524288
               _ExtentX        =   7938
               _ExtentY        =   12250
               _StockProps     =   64
               BackColorStyle  =   1
               BorderStyle     =   0
               DisplayColHeaders=   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   12
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               GrayAreaBackColor=   16777215
               MaxCols         =   2
               MaxRows         =   16
               RowHeaderDisplay=   0
               ScrollBars      =   0
               SpreadDesigner  =   "frm일일매출마감2.frx":4E0D
               UserResize      =   0
               HighlightHeaders=   1
               HighlightStyle  =   1
            End
         End
         Begin Threed.SSFrame SSFrame2 
            Height          =   7065
            Left            =   10005
            TabIndex        =   15
            Top             =   120
            Width           =   5115
            _ExtentX        =   9022
            _ExtentY        =   12462
            _Version        =   262144
            Font3D          =   3
            BackColor       =   16777215
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "frm일일매출마감2.frx":5C89
            Caption         =   " 의류별 접수내역"
            Begin FPSpreadADO.fpSpread sprCloth 
               Height          =   6585
               Left            =   105
               TabIndex        =   16
               Top             =   375
               Width           =   4905
               _Version        =   524288
               _ExtentX        =   8652
               _ExtentY        =   11615
               _StockProps     =   64
               BorderStyle     =   0
               DisplayRowHeaders=   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               GrayAreaBackColor=   16777215
               GridShowVert    =   0   'False
               MaxCols         =   3
               ScrollBars      =   2
               SpreadDesigner  =   "frm일일매출마감2.frx":669B
               UserResize      =   0
               HighlightHeaders=   1
               HighlightStyle  =   1
            End
         End
      End
      Begin Threed.SSPanel pnlMsg 
         Height          =   750
         Left            =   15
         TabIndex        =   1
         Top             =   450
         Width           =   20280
         _ExtentX        =   35772
         _ExtentY        =   1323
         _Version        =   262144
         Font3D          =   3
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   15.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSPanel SSPanel 
            Height          =   555
            Index           =   0
            Left            =   1110
            TabIndex        =   9
            Top             =   105
            Width           =   3165
            _ExtentX        =   5583
            _ExtentY        =   979
            _Version        =   262144
            BorderWidth     =   0
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin MSComCtl2.DTPicker dtpDay 
               Height          =   405
               Left            =   75
               TabIndex        =   10
               Top             =   75
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   714
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   57016320
               UpDown          =   -1  'True
               CurrentDate     =   40279
            End
         End
         Begin VB.ComboBox cboManager 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5400
            Locked          =   -1  'True
            Style           =   2  '드롭다운 목록
            TabIndex        =   7
            Top             =   180
            Width           =   2340
         End
         Begin XtremeSuiteControls.PushButton cmdFinish 
            Height          =   630
            Index           =   0
            Left            =   8055
            TabIndex        =   2
            Top             =   60
            Width           =   1590
            _Version        =   851970
            _ExtentX        =   2805
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 업무마감"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm일일매출마감2.frx":6C46
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   13710
            TabIndex        =   3
            Top             =   60
            Width           =   1395
            _Version        =   851970
            _ExtentX        =   2461
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frm일일매출마감2.frx":7520
         End
         Begin VB.Label lblTitle 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "담당자:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   4605
            TabIndex        =   8
            Top             =   285
            Width           =   750
         End
         Begin VB.Label lblTitle 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "마감일자:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   15
            TabIndex        =   5
            Top             =   270
            Width           =   1005
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   4
         Top             =   15
         Width           =   20280
         _ExtentX        =   35772
         _ExtentY        =   741
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "   일일매출마감"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm일일매출마감2.frx":85B2
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image1 
            Height          =   240
            Left            =   75
            Picture         =   "frm일일매출마감2.frx":87D8
            Top             =   90
            Width           =   240
         End
      End
   End
End
Attribute VB_Name = "frm일일매출마감2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim 마감일자     As String
Dim 접수수량       As Integer
Dim 반품수량     As Integer
Dim 재세탁수량   As Integer
Dim 수선수량     As Integer
Dim 수선금액     As Long
Dim strDSu       As Integer
Dim dblDSuAmt    As Double
Dim strTAmt      As Long
Dim strAAmt      As Long
Dim strHAmt      As Long
Dim strSAmt      As Long
Dim chkSale      As String

Dim 시작택번호   As String
Dim 마지막택번호 As String

Dim m_FTC_Mode   As Long
Dim SendFlag     As LaundrySendFlag ' 현재 어떤 작업을 하였는지
Dim SendMode     As Integer         ' 1: Modem,  2, Internet
Dim SendFile     As String          ' 서버에 전송할 파일이름
Dim strSendPath  As String          ' 서버로 전송한 자료가 보관될 위치
Dim strRecvPath  As String          ' 서버에 수신할 자료가 보관되어 있는 위치
Dim strPrgPath   As String          ' 서버에 수신할 프로그래이 보관되어 있는 위치
Dim strSendData  As String          ' 서버에 전송할 메시지

Dim g_AgencyCode As String

Dim PauseTime, Start, Finish, TotalTime ' 서버의 전송을 기다리는데 필요

Private Sub 의류접수_Display()
    Query = "SELECT    의류명"
    Query = Query & ", ISNULL(COUNT(택번호),0) AS 수량"
    Query = Query & ", ISNULL(SUM(금액),0)     AS 금액"
    Query = Query & " FROM TB_입출고"
    Query = Query & " WHERE 접수일자 = '" & Format(dtpDay.Value, "YYYY-MM-DD") & "'"
    Query = Query & " GROUP BY 의류명"
    Query = Query & " ORDER BY 의류명 ASC"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    With sprCloth
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = ADORs!의류명 & ""
            .Col = 2: .Text = ADORs!수량 & ""
            .Col = 3: .Text = ADORs!금액 & ""
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        .ReDraw = True
    End With
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
    
        Case 5:
            '마감일자가 오늘 이전인 경우 마감작업을 안하면 종료가 안된다.
            If Format(Date, "YYYY-MM-DD") > Format(dtpDay.Value, "YYYY-MM-DD") Then
                Query = "SELECT * FROM TB_일일마감"
                Query = Query & " WHERE 마감일자 = '" & Format(dtpDay.Value, "YYYY-MM-DD") & "'"
                Set ADORs = New ADODB.Recordset
                ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
                If ADORs.EOF Then
                    ADORs.Close
                    Set ADORs = Nothing
                    
                    MsgBox "일일마감을 해주세요.", vbInformation, "확인"
                    Exit Sub
                End If
            End If
            
            Unload Me
    End Select
    
    Exit Sub
    
ErrRtn:

End Sub

Private Sub 일일마감_Proc()
    Dim 접수금액   As Long
    Dim 가맹점마진 As Long
    Dim 외주마진   As Long
    
    Dim 미수금액 As Long
    Dim tmpData  As String
    
    On Error GoTo ErrRtn
    
    Screen.MousePointer = 11
    pnlProg.Visible = True
    DoEvents
    
    마감일자 = Format(dtpDay.Value, "YYYY-MM-DD")
        
    With sprGrid(0)
        '----------------------------------------------------------------
        ' 1. 접수수량, 접수금액 구하기
        '----------------------------------------------------------------
        Query = "SELECT    ISNULL(COUNT(택번호),0)"
        Query = Query & ", ISNULL(SUM(금액),0)"
        Query = Query & " FROM TB_입출고"
        Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
        Query = Query & "   AND ((판매취소일자 IS NULL OR 판매취소일자 = '')"
        Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
        Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        .Row = 1
        .Col = 1: .Value = ADORs(0)
        .Col = 3: .Value = ADORs(1)
        
        접수금액 = ADORs(1)
        미수금액 = ADORs(1)
        
        ADORs.Close
        Set ADORs = Nothing
            
        '----------------------------------------------------------------
        ' 2. 출고수량 구하기
        '----------------------------------------------------------------
        Query = "SELECT    ISNULL(COUNT(택번호),0)"
        Query = Query & " FROM TB_입출고"
        Query = Query & " WHERE 출고일자 = '" & 마감일자 & "'"
        Query = Query & "   AND ((판매취소일자 IS NULL OR 판매취소일자 = '')"
        Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
        Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
        
        .Row = 2
        .Col = 1: .Text = Recordset_Result(Query)
            
        '----------------------------------------------------------------
        ' 2-1) 현금결제 구하기
        '----------------------------------------------------------------
        Query = "SELECT ISNULL(SUM(현금입금),0)"
        Query = Query & " FROM TB_매출"
        Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
        
        .Row = 3
        .Col = 3: .Value = Recordset_Result(Query)
        .Col = 3: 미수금액 = 미수금액 - .Value
        
        '----------------------------------------------------------------
        ' 2-2) 카드결제 구하기
        '----------------------------------------------------------------
        Query = "SELECT    ISNULL(COUNT(카드입금),0)"
        Query = Query & ", ISNULL(SUM(카드입금),0)"
        Query = Query & " FROM TB_매출"
        Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
        Query = Query & "   AND 카드입금 > 0"
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        .Row = 4
        .Col = 1: .Value = ADORs(0)
        .Col = 3: .Value = ADORs(1)
        
        .Col = 3: 미수금액 = 미수금액 - .Value
        
        ADORs.Close
        Set ADORs = Nothing
    
        '미수금액
        sprGrid(1).Row = 6
        sprGrid(1).Col = 1: sprGrid(1).Value = 미수금액 & ""
        
        '----------------------------------------------------------------
        ' 3-1) 가맹점 마진
        '----------------------------------------------------------------
        Query = "SELECT ISNULL(SUM(금액 * 세탁마진/100),0)"
        Query = Query & " FROM TB_입출고"
        Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
        Query = Query & "   AND 내용 LIKE '%세%'"
        Query = Query & "   AND ((판매취소일자 IS NULL OR 판매취소일자 = '')"
        Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
        Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
        
        .Row = 5
        .Col = 3: .Value = Recordset_Result(Query) '
        .Col = 3: 가맹점마진 = .Value              '
        
        ' 3-2) 외주 마진
        Query = "SELECT ISNULL(SUM(금액*외주마진/100),0)"
        Query = Query & " FROM TB_입출고"
        Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
        Query = Query & "   AND SUBSTRING(의류코드,1,1) = 'a'"                 '운동화
        Query = Query & "   AND 내용  LIKE '%세%'"
        Query = Query & "   AND ((판매취소일자 IS NULL OR 판매취소일자 = '')"
        Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
        Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
        
        sprGrid(1).Row = 13
        sprGrid(1).Col = 1: sprGrid(1).Value = Recordset_Result(Query)
        sprGrid(1).Col = 1: 외주마진 = sprGrid(1).Value
        
        ' 지사 마진
        .Row = 6
        .Col = 3: .Value = 접수금액 - 가맹점마진 - 외주마진 & ""
        
'=================================================================================

        '----------------------------------------------------------------
        ' 6) 판매취소 계산
        '----------------------------------------------------------------
        Query = "SELECT    ISNULL(COUNT(택번호),0)"
        Query = Query & ", ISNULL(SUM(금액),0)"
        Query = Query & " FROM TB_입출고"
        Query = Query & " WHERE SUBSTRING(판매취소일자,1,10) = '" & 마감일자 & "'"
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        .Row = 7
        .Col = 1: .Text = ADORs(0)
        .Col = 3: .Text = ADORs(1)
        
        ADORs.Close
        Set ADORs = Nothing
        
        
        '----------------------------------------------------------------
        ' 7) 반품환불 계산
        '----------------------------------------------------------------
        Query = "SELECT    ISNULL(COUNT(택번호),0)"
        Query = Query & ", ISNULL(SUM(금액),0)"
        Query = Query & " FROM TB_입출고"
        Query = Query & " WHERE SUBSTRING(반품환불일자,1,10) = '" & 마감일자 & "'"
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        .Row = 8
        .Col = 1: .Text = ADORs(0)
        .Col = 3: .Text = ADORs(1)
        
        ADORs.Close
        Set ADORs = Nothing
        
        '----------------------------------------------------------------
        ' 8) 세탁환불 계산
        '----------------------------------------------------------------
        Query = "SELECT    ISNULL(COUNT(택번호),0)"
        Query = Query & ", ISNULL(SUM(금액),0)"
        Query = Query & " FROM TB_입출고"
        Query = Query & " WHERE SUBSTRING(세탁환불일자,1,10) = '" & 마감일자 & "'"
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        .Row = 9
        .Col = 1: .Text = ADORs(0)
        .Col = 3: .Text = ADORs(1)
        
        ADORs.Close
        Set ADORs = Nothing
        
'=================================================================================
        '--------------------------------------------------------------
        ' 9-1) 수선수량 계산
        '--------------------------------------------------------------
        Query = "SELECT    ISNULL(COUNT(택번호),0)"
        Query = Query & ", ISNULL(SUM(수선금액),0)"
        Query = Query & " FROM TB_입출고"
        Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
        Query = Query & "   AND (내용  = '드수' OR 내용 = '수') "
        Query = Query & "   AND ((판매취소일자 IS NULL OR 판매취소일자 = '')"
        Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
        Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
        .Row = 11
        .Col = 1: .Value = ADORs(0)
        .Col = 3: .Value = ADORs(1)
        
        ADORs.Close
        Set ADORs = Nothing
    
    
        '----------------------------------------------------------------
        ' 9-1) 재세탁수량 계산
        '----------------------------------------------------------------
        Query = "SELECT ISNULL(COUNT(택번호),0)"
        Query = Query & " FROM TB_입출고"
        Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
        Query = Query & "   AND 내용     = '드재'"
        Query = Query & "   AND ((판매취소일자 IS NULL OR 판매취소일자 = '')"
        Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
        Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
        
        .Row = 12
        .Col = 1: .Value = Recordset_Result(Query)
        
        '--------------------------------------------------------------------
        '운동화 매출을 불러온다.
        '--------------------------------------------------------------------
        Query = "SELECT    ISNULL(COUNT(의류코드),0)" '운동화건수
        Query = Query & ", ISNULL(SUM(금액),0)"       '운동화금액
        Query = Query & " FROM TB_입출고 "
        Query = Query & " WHERE SUBSTRING(의류코드,1,2) = 'a0'"
        
        'Query = Query & " WHERE (UPPER(의류코드) >= 'A00'"
        'Query = Query & "   AND  UPPER(의류코드) <= 'A99')"
        
        Query = Query & "   AND 접수일자 = '" & 마감일자 & "'"
        Query = Query & "   AND ((판매취소일자 IS NULL OR 판매취소일자 = '')"
        Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
        Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        .Row = 13
        .Col = 1: .Value = ADORs(0) & ""
        .Col = 3: .Value = ADORs(1) & ""
        
        ADORs.Close
        Set ADORs = Nothing

    
        '--------------------------------------------------------------------
        '가죽/무스탕 매출을 불러온다.
        '--------------------------------------------------------------------
        Query = "SELECT    ISNULL(COUNT(의류코드),0)" '가죽건수
        Query = Query & ", ISNULL(SUM(금액),0)"       '가죽금액
        Query = Query & " FROM TB_입출고 "
        Query = Query & " WHERE (SUBSTRING(의류코드,1,2) = 'b0'"
        Query = Query & "    OR SUBSTRING(의류코드,1,2) = 'n0')"
        
        'Query = Query & " WHERE (UPPER(의류코드) >= 'B00' AND  UPPER(의류코드) <= 'B99')"
        'Query = Query & "    OR (UPPER(의류코드) >= 'N00' AND  UPPER(의류코드) <= 'N99')"
        
        Query = Query & "   AND 접수일자 = '" & 마감일자 & "'"
        Query = Query & "   AND ((판매취소일자 IS NULL OR 판매취소일자 = '')"
        Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
        Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        .Row = 14
        .Col = 1: .Value = ADORs(0) & ""
        .Col = 3: .Value = ADORs(1) & ""
        
        ADORs.Close
        Set ADORs = Nothing
        
        '--------------------------------------------------------------------
        '카페트 매출을 불러온다.
        '--------------------------------------------------------------------
        Query = "SELECT    ISNULL(COUNT(의류코드),0)" '카페트건수
        Query = Query & ", ISNULL(SUM(금액),0) "      '카페트금액
        Query = Query & " FROM TB_입출고 "
        Query = Query & " WHERE SUBSTRING(의류코드,1,2) = 'x0'"
        
        'Query = Query & " WHERE (UPPER(의류코드) >= 'X00'"
        'Query = Query & "   AND  UPPER(의류코드) <= 'X99')"
        
        Query = Query & "   AND 접수일자 = '" & 마감일자 & "'"
        Query = Query & "   AND ((판매취소일자 IS NULL OR 판매취소일자 = '')"
        Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
        Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        .Row = 15
        .Col = 1: .Value = ADORs(0) & ""
        .Col = 3: .Value = ADORs(1) & ""
        
        ADORs.Close
        Set ADORs = Nothing
        
        '----------------------------------------------------------------
        ' 반품수량
        '----------------------------------------------------------------
        Query = "SELECT ISNULL(COUNT(택번호),0)"
        Query = Query & " FROM TB_입출고"
        Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
        Query = Query & "   AND 내용     = '%반%'"
        Query = Query & "   AND ((판매취소일자 IS NULL OR 판매취소일자 = '')"
        Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
        Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
        
        .Row = 16
        .Col = 1: .Value = Recordset_Result(Query)
        
    End With
        
    With sprGrid(1)
        '--------------------------------------------------------------------
        ' 발생/사용/삭제 마일리지
        '--------------------------------------------------------------------
        Query = "SELECT    ISNULL(SUM(발생마일리지),0)"
        Query = Query & ", ISNULL(SUM(사용마일리지),0)"
        Query = Query & ", ISNULL(SUM(삭제마일리지),0)"
        Query = Query & " FROM TB_매출"
        Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        .Col = 1
        .Row = 1: .Value = ADORs(0) & ""
        .Row = 2: .Value = ADORs(1) & ""
        .Row = 3: .Value = ADORs(2) & ""
        
        ADORs.Close
        Set ADORs = Nothing
        
        
        '--------------------------------------------------------------------
        ' 쿠폰
        '--------------------------------------------------------------------
        Query = "SELECT    ISNULL(SUM(쿠폰입금),0)"
        Query = Query & ", ISNULL(COUNT(쿠폰번호),0)"
        Query = Query & " FROM TB_매출"
        Query = Query & " WHERE 매출일자 = '" & 마감일자 & "'"
        Query = Query & "   AND 쿠폰입금 > 0"
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        .Col = 1
        .Row = 4: .Value = ADORs(0) & ""
        .Row = 5: .Value = ADORs(1) & ""
        
        ADORs.Close
        Set ADORs = Nothing
        
        '----------------------------------------------------------------
        ' 6) 판매취소 계산
        '----------------------------------------------------------------
        Query = "SELECT    택번호"
        Query = Query & " FROM TB_입출고"
        Query = Query & " WHERE SUBSTRING(판매취소일자,1,10) = '" & 마감일자 & "'"
        Query = Query & " ORDER BY 택번호 ASC"
        
        tmpData = Get_택번호(Query)
        
        .TypeComboBoxClear 1, 7
        
        If tmpData <> "" Then
            .Row = 7
            .Col = 1: .TypeComboBoxList = tmpData & ""
        End If
        
        '----------------------------------------------------------------
        ' 7) 반품환불 계산
        '----------------------------------------------------------------
        Query = "SELECT    택번호"
        Query = Query & " FROM TB_입출고"
        Query = Query & " WHERE SUBSTRING(반품환불일자,1,10) = '" & 마감일자 & "'"
        
        tmpData = Get_택번호(Query)
        
        .TypeComboBoxClear 1, 8
        
        If tmpData <> "" Then
            .Row = 8
            .Col = 1: .TypeComboBoxList = tmpData & ""
        End If
        
        '----------------------------------------------------------------
        ' 8) 세탁환불 계산
        '----------------------------------------------------------------
        Query = "SELECT    택번호"
        Query = Query & " FROM TB_입출고"
        Query = Query & " WHERE SUBSTRING(세탁환불일자,1,10) = '" & 마감일자 & "'"
    
        tmpData = Get_택번호(Query)
        
        .TypeComboBoxClear 1, 9
        
        If tmpData <> "" Then
            .Row = 9
            .Col = 1: .TypeComboBoxList = tmpData & ""
        End If
        
        '--------------------------------------------------------------------
        ' 누락TAG CHECK
        '--------------------------------------------------------------------
        Dim 시작택번호   As String
        Dim 마지막택번호 As String
        
        Dim 택번호 As String
        Dim tmpTAG As String
        
        Query = "SELECT    MIN(택번호)"
        Query = Query & ", MAX(택번호)"
        Query = Query & " FROM TB_입출고 "
        Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
        Query = Query & "   AND ((판매취소일자 IS NULL OR 판매취소일자 = '')"
        Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
        Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        If Not ADORs.EOF Then
            시작택번호 = ADORs(0)
            마지막택번호 = ADORs(1)
        End If
        ADORs.Close
        Set ADORs = Nothing
        
        '--------------------------------------------------------------------
        ' 누락택
        '--------------------------------------------------------------------
        Dim iLoop As Long
        
        Query = "SELECT 택번호 FROM TB_입출고"
        Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
        Query = Query & "   AND ((판매취소일자 IS NULL OR 판매취소일자 = '')"
        Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
        Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
        
        Query = Query & " ORDER BY 택번호 ASC"
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        iLoop = 0
        
        택번호 = ""
        tmpTAG = ""
        
        If Val(마지막택번호) - Val(시작택번호) < 5000 Then
            Do Until ADORs.EOF
                If tmpTAG = "" Then
                    tmpTAG = ADORs!택번호
                Else
                    Do Until Format(CLng(tmpTAG) + 1, "000000000") >= ADORs!택번호
                        택번호 = 택번호 & Chr(9) & Format(CLng(tmpTAG) + 1, "000-00-0000")
                        
                        tmpTAG = Format(CLng(tmpTAG) + 1, "000000000")
                        
                        '100 개가 넘으면 빠져 나옴
                        If iLoop >= 100 Then
                            택번호 = 택번호 & Chr(9) & "Err"
                            Exit Do
                        End If
                        
                        iLoop = iLoop + 1
                    Loop
                    
                    tmpTAG = Format(CLng(tmpTAG) + 1, "000000000")
                End If
                
                ADORs.MoveNext
            Loop
            ADORs.Close
            Set ADORs = Nothing
            
            .TypeComboBoxClear 1, 10
            
            If 택번호 <> "" Then
                .Row = 10
                .Col = 1: .TypeComboBoxList = 택번호 & ""
            End If
            
            sprGrid(0).Row = 10
            sprGrid(0).Col = 1: sprGrid(0).Value = .TypeComboBoxCount '누락택 갯수
        End If
        
        .Row = 11
        .Col = 1: .Text = Format(시작택번호, "000-00-0000") & ""
        
        .Row = 12
        .Col = 1: .Text = Format(마지막택번호, "000-00-0000") & ""
        
        
        '--------------------------------------------------------------------
        ' 삼성 카드 할인 내용 추가
        '--------------------------------------------------------------------
        Dim 삼성카드고객수   As Long
        Dim 삼성카드할인건수 As Long
        Dim 삼성카드할인금액 As Long
    
        삼성카드고객수 = 0
        삼성카드할인건수 = 0
        삼성카드할인금액 = 0
        
        Query = "SELECT    고객코드"
        Query = Query & ", ISNULL(COUNT(금액),0)"
        Query = Query & ", ISNULL(SUM(금액),0)"
        Query = Query & " FROM TB_입출고"
        Query = Query & " WHERE 접수일자 = '" & 마감일자 & "'"
        Query = Query & "   AND 내용  LIKE '%삼%'"
        Query = Query & "   AND ((판매취소일자 IS NULL OR 판매취소일자 = '')"
        Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
        Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
        Query = Query & " GROUP BY 고객코드"
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        Do Until ADORs.EOF
            삼성카드고객수 = 삼성카드고객수 + 1
    
            삼성카드할인건수 = 삼성카드할인건수 + ADORs(0)
            삼성카드할인금액 = 삼성카드할인금액 + ADORs(1)
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        .Col = 1
        .Row = 14: .Text = 삼성카드할인금액 & ""
        .Row = 15: .Text = 삼성카드할인건수 & ""
        .Row = 16: .Text = 삼성카드고객수 & ""
    End With
        
    Screen.MousePointer = 0
    pnlProg.Visible = False
    DoEvents
    
    Exit Sub
    
ErrRtn:
    Screen.MousePointer = 0
    pnlProg.Visible = False
End Sub

Private Function Get_택번호(Query As String) As String
    On Error GoTo ErrRtn
    
    Dim 택번호 As String
    Dim tmpTAG As String
    Dim iBar   As Integer
    
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    iBar = 0
    택번호 = ""
    
    Do Until ADORs.EOF
        'If 택번호 = "" Then
        '    택번호 = Format(ADORs!택번호, "000-00-0000")
        'Else
        '    If Format(CLng(tmpTAG) + 1, "000000000") = ADORs!택번호 Then
        '        iBar = iBar + 1
        '    Else
        '        If iBar > 1 Then
        '            택번호 = 택번호 & " - " & Format(ADORs!택번호, "000-00-0000")
        '        Else
        '            택번호 = 택번호 & Chr(9) & Format(ADORs!택번호, "000-00-0000")
        '        End If
        '
        '        iBar = 0
        '    End If
        'End If
        
        'tmpTAG = ADORs!택번호
        
        If 택번호 = "" Then
            택번호 = Format(ADORs!택번호, "000-00-0000")
        Else
            택번호 = 택번호 & Chr(9) & Format(ADORs!택번호, "000-00-0000")
        End If
        
        ADORs.MoveNext
    Loop
    ADORs.Close
    Set ADORs = Nothing
    
    'If iBar > 1 Then
    '    택번호 = 택번호 & " - " & Format(tmpTAG, "000-00-0000")
    'End If

    Get_택번호 = 택번호
    
    Exit Function
    
ErrRtn:
    Get_택번호 = ""
End Function

Private Sub cmdFinish_Click(Index As Integer)
    'On Error GoTo ErrRtn
    
    pnlMsg.Caption = ""
    
    마감일자 = Format(dtpDay.Value, "YYYY-MM-DD")
    DoEvents
        
    If Get_일일마감여부(마감일자) = True Then
        MsgBox "일마감이 완료 되었으므로 마감작업을 할 수 없습니다", vbInformation, "확인"
        
        dtpDay.SetFocus
        
        Exit Sub
    End If
    
    If cboManager.ListIndex < 0 Then
        MsgBox "근무자를 선택하세요.", vbInformation, "확인"
        
        cboManager.SetFocus
        Exit Sub
    End If
    
    
''        '----------------------------------------------------------------------------------------
''        '
''        '----------------------------------------------------------------------------------------
''                Query = "카드 금액 : " & Format(카드금액, "#,##0") & "원" & vbNewLine
''        Query = Query & "카드 건수 : " & Format(카드건수, "#,##0") & "건" & vbNewLine & vbNewLine
''        Query = Query & "카드 관련 내용은 마감하면 수정이 불가능 합니다." & vbNewLine & vbNewLine
''        Query = Query & "입력한 내용이 맞습니까?"
''
''        If MsgBox(Query, vbCritical + vbYesNo, "확인") = vbNo Then
''            'txtCard(0).SetFocus
''
''            'frmMain.Command1(3).Enabled = True
''            cmdFinish(0).Enabled = True
''            'cmdFinish(1).Enabled = True
''            Exit Sub
''        End If
        
'        txtCoupon.Text = Trim(txtCoupon.Text)
'        txtCoupon.Text = Replace(txtCoupon.Text, ".", ",")

'        If txtCoupon.Text <> "" Then
'            Dim tmpVar As Variant
'            Dim nForIndex As Integer
'
'            tmpVar = Split(txtCoupon.Text, ",")

'            For nForIndex = 0 To UBound(tmpVar)
'                If Len(CStr(tmpVar(nForIndex))) <> M_COUPON_LENGTH Then
'                    MsgBox "입력한 쿠폰 번호가 올바르지 않습니다." & vbNewLine & vbNewLine & "[" & CStr(tmpVar(nForIndex)) & "]", vbInformation, "확인"
'                    cmdFinish(0).Enabled = True
'                    cmdFinish(1).Enabled = True
'                    Exit Sub
'                End If
'            Next nForIndex
'
'            txtCoupon.Tag = UBound(tmpVar) + 1
'        End If
        
    Rtn = MsgBox("[ " & 마감일자 & " ]" & " 마감을 하시겠습니까..?", vbQuestion + vbYesNo, "일일마감")
    
    If Rtn = vbNo Then
        dtpDay.SetFocus
        
        Exit Sub
    End If
    
    '------------------------------------------------------------------------------------------------------
    ' TB_일일마감
    '------------------------------------------------------------------------------------------------------
    Query = "DELETE FROM TB_일일마감 WHERE 마감일자 = '" & 마감일자 & "'"
    ADOCon.Execute Query
    
    Call Sale_Check
    
    ' 마일리지 마감 ( 3개월동안 이용 실적이 없을 경우 마일리지 삭제 )
    If 가맹점정보.마일리지여부 = "Y" Then
        Call Set_마일리지삭제
    End If
    
       
    '--------------------------------------------------------------------
    '
    '--------------------------------------------------------------------
    Query = "INSERT INTO TB_일일마감("
    Query = Query & "  가맹점코드"         ' 1
    Query = Query & ", 마감일자"           ' 2
    Query = Query & ", 접수금액"           ' 3
    Query = Query & ", 접수수량"           ' 4
    Query = Query & ", 출고수량"           ' 5
    Query = Query & ", 반품수량"           ' 6
    Query = Query & ", 재세탁수량"         ' 7
    Query = Query & ", 수선금액"           ' 8
    Query = Query & ", 수선수량"           ' 9
    Query = Query & ", 판매구분"           '10
    Query = Query & ", 시작택번호"         '11
    Query = Query & ", 종료택번호"         '12
    Query = Query & ", 쿠폰금액"           '13
    Query = Query & ", 쿠폰건수"           '14
    Query = Query & ", 발생마일리지"       '15
    Query = Query & ", 사용마일리지"       '16
    Query = Query & ", 삭제마일리지"       '17
    Query = Query & ", 현금입금"           '18
    Query = Query & ", 카드금액"           '19
    Query = Query & ", 카드건수"           '20
    Query = Query & ", 반품환불금액"       '21
    Query = Query & ", 반품환불건수"       '22
    Query = Query & ", 세탁환불금액"       '23
    Query = Query & ", 세탁환불건수"       '24
    Query = Query & ", 삼성카드할인금액"   '25
    Query = Query & ", 삼성카드할인건수"   '26
    Query = Query & ", 삼성카드할인고객수" '27
    Query = Query & ", 근무자명"           '28
    Query = Query & ", 지사금액"           '29
    Query = Query & ", 가맹점금액"         '30
    Query = Query & ", 운동화금액"         '31
    Query = Query & ", 운동화건수"         '32
    Query = Query & ", 운동화비율"         '33
    Query = Query & ", 카페트금액"         '34
    Query = Query & ", 카페트건수"         '35
    Query = Query & ", 명품세탁금액"       '36
    Query = Query & ", 명품세탁건수"       '37
    Query = Query & ", 명품세탁비율"       '38
    Query = Query & ", 명품염색금액"       '39
    Query = Query & ", 명품염색건수"       '40
    Query = Query & ", 명품염색비율"       '41
    Query = Query & ", 마감여부"           '42
    Query = Query & ", 본사전송여부"       '43
    Query = Query & ", 지사코드"           '44
    Query = Query & ") VALUES ("
    Query = Query & "   '" & 가맹점정보.가맹점코드 & "'"                  ' 1 가맹점코드
    Query = Query & ",  '" & 마감일자 & "'"                               ' 2 마감일자
    Query = Query & ",  " & Spread_GetData(sprGrid(0), 1, 3, False)       ' 3 접수금액
    Query = Query & ",  " & Spread_GetData(sprGrid(0), 1, 1, False)       ' 4 접수수량
    Query = Query & ",  " & Spread_GetData(sprGrid(0), 1, 1, False)       ' 5 출고수량
    Query = Query & ",  " & Spread_GetData(sprGrid(0), 16, 1, False)      ' 6 반품수량
    Query = Query & ",  " & Spread_GetData(sprGrid(0), 12, 1, False)      ' 7 재세탁수량
    Query = Query & ",  " & Spread_GetData(sprGrid(0), 11, 3, False)      ' 8 수선금액
    Query = Query & ",  " & Spread_GetData(sprGrid(0), 11, 1, False)      ' 9 수선수량
    Query = Query & ", '" & chkSale & "'"                                 '10 판매구분
    Query = Query & ", '" & Replace(Spread_GetData(sprGrid(1), 11, 1, True), "-", "") & "'" '11 시작택번호
    Query = Query & ", '" & Replace(Spread_GetData(sprGrid(1), 12, 1, True), "-", "") & "'" '12 종료택번호
    Query = Query & ",  " & Spread_GetData(sprGrid(1), 4, 1, False)       '13 쿠폰금액
    Query = Query & ",  " & Spread_GetData(sprGrid(1), 5, 1, False)       '14 쿠폰건수
    Query = Query & ",  " & Spread_GetData(sprGrid(1), 1, 1, False)       '15 발생마일리지
    Query = Query & ",  " & Spread_GetData(sprGrid(1), 2, 1, False)       '16 사용마일리지
    Query = Query & ",  " & Spread_GetData(sprGrid(1), 3, 1, False)       '17 삭제마일리지
    Query = Query & ",  " & Spread_GetData(sprGrid(0), 3, 3, False)       '18 현금입금
    Query = Query & ",  " & Spread_GetData(sprGrid(0), 4, 3, False)       '19 카드금액
    Query = Query & ",  " & Spread_GetData(sprGrid(0), 4, 1, False)       '20 카드건수
    Query = Query & ",  " & Spread_GetData(sprGrid(0), 8, 3, False)       '21 반품환불금액
    Query = Query & ",  " & Spread_GetData(sprGrid(0), 8, 1, False)       '22 반품환불건수
    Query = Query & ",  " & Spread_GetData(sprGrid(0), 9, 3, False)       '23 세탁환불금액
    Query = Query & ",  " & Spread_GetData(sprGrid(0), 9, 1, False)       '24 세탁환불건수
    Query = Query & ",  " & Spread_GetData(sprGrid(1), 14, 1, False)      '25 삼성카드할인금액
    Query = Query & ",  " & Spread_GetData(sprGrid(1), 15, 1, False)      '26 삼성카드할인건수
    Query = Query & ",  " & Spread_GetData(sprGrid(1), 16, 1, False)      '27 삼성카드할인고객수
    Query = Query & ", '" & cboManager.Text & "'"                         '28 근무자명
    Query = Query & ",  " & Spread_GetData(sprGrid(0), 6, 3, False)       '29 지사금액
    Query = Query & ",  " & Spread_GetData(sprGrid(0), 5, 3, False)       '30 가맹점금액
    Query = Query & ",  " & Spread_GetData(sprGrid(0), 13, 3, False)      '31 운동화금액
    Query = Query & ",  " & Spread_GetData(sprGrid(0), 13, 1, False)      '32 운동화건수
    Query = Query & ",  0"                                                '33 운동화비율
    Query = Query & ",  " & Spread_GetData(sprGrid(0), 15, 3, False)      '34 카페트금액
    Query = Query & ",  " & Spread_GetData(sprGrid(0), 15, 1, False)      '35 카페트건수
    Query = Query & ",  0"                                                '36 명품세탁금액
    Query = Query & ",  0"                                                '37 명품세탁건수
    Query = Query & ",  0"                                                '38 명품세탁비율
    Query = Query & ",  0"                                                '39 명품염색금액
    Query = Query & ",  0"                                                '40 명품염색건수
    Query = Query & ",  0"                                                '41 명품염색비율
    Query = Query & ", 'Y'"                                               '42 마감여부
    Query = Query & ", 'N'"                                               '43 전송여부
    Query = Query & ", '" & 가맹점정보.지사코드 & "'"                     '44 지사코드
    Query = Query & ")"
    ADOCon.Execute Query
    
''    '--------------------------------------------------------------------
''    ' 누락TAG CHECK
''    '--------------------------------------------------------------------
''    Dim Tag_No    As String
''    Dim sMemTagNo As String
''    Dim strTag    As String
''
''    Query = "SELECT 택번호 FROM TB_입출고 "
''    Query = Query & " WHERE 접수일자 = '" & Format(dtpDay.Value, "YYYY-MM-DD") & "'"
''    Query = Query & "   AND ((판매취소일자 IS NULL OR 판매취소일자 = '')"
''    Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
''    Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
''    Query = Query & " ORDER BY 택번호 "
''    Set ADORs = New ADODB.Recordset
''    ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
''
''    'cboInput.Clear
''    strTag = ""
''    sprGrid(0).TypeComboBoxClear 1, 6
''
''    If Val(마지막택번호) - Val(시작택번호) < 5000 Then
''        Do Until ADORs.EOF
''            sMemTagNo = Left(ADORs!택번호, 2) & Right(ADORs!택번호, 4)
''            Tag_No = Left(ADORs!택번호, 2) & Right(ADORs!택번호, 4)
''
''            ADORs.MoveNext
''
''            If ADORs.EOF Then
''                Exit Do
''            Else
''                'Do While (Val(sMemTagNo) + 1) <> (Val(Left(ADORs!택번호, 2) & Right(ADORs!택번호, 4)))
''
''                Do Until Str(Val(sMemTagNo) + 1) <> Tag_No
''                    sMemTagNo = Val(sMemTagNo) + 1
''
''                    'cboInput.AddItem Format(sMemTagNo, "00-0000")
''
''                    If strTag = "" Then
''                        strTag = Format(sMemTagNo, "00-0000")
''                    Else
''                        strTag = strTag + Chr(9) + Format(sMemTagNo, "00-0000")
''                    End If
''
''                    DoEvents
''                Loop
''
''                sprGrid(0).Col = 1
''                sprGrid(0).Row = 6: sprGrid(0).TypeComboBoxList = strTag & ""
''            End If
''        Loop
''    End If
''    ADORs.Close
''    Set ADORs = Nothing
''
''
''    sprGrid(0).Col = 1
''    sprGrid(0).Row = 6
''
''    If sprGrid(0).TypeComboBoxCount > 0 Then
''        MsgBox "누락 TAG번호가 " & sprGrid(0).TypeComboBoxCount & "건 발생 하였습니다.", vbInformation
''    End If
    
            
    '---------------------------------------------------------------------------------
    ' TB_근무현황
    '---------------------------------------------------------------------------------
    Query = "UPDATE TB_근무현황 SET 종료일자 = '" & Format(Date, "YYYY-MM-DD") & "'"
    Query = Query & "             , 종료시간 = '" & Format(Now, "hh:mm:ss") & "'"
    Query = Query & "             , 업무마감 = 'Y'"
    Query = Query & " WHERE 근무자명 = '" & cboManager.Text & "'"
    Query = Query & "   AND 시작일자 = '" & Format(Date, "YYYY-MM-DD") & "'"
    Query = Query & "   AND 종료일자 = ''"
    ADOCon.Execute Query
    
    MsgBox "> 일일 마감이 완료 되었습니다. <", vbInformation, "일일마감"
    
    Unload Me
    
    Exit Sub
    
ErrRtn:
    Screen.MousePointer = 0
    
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Control_Visible()
    cmdFinish(0).Enabled = Not cmdFinish(0).Enabled
    
    lblTitle(0).Visible = Not lblTitle(0).Visible
    lblTitle(1).Visible = Not lblTitle(1).Visible
    
    dtpDay.Visible = Not dtpDay.Visible
    cboManager.Visible = Not cboManager.Visible
    
    pnlProg.Left = 120
    pnlProg.Top = 135
    pnlProg.Visible = Not pnlProg.Visible
    
    DoEvents
End Sub

Private Sub dtpDay_Change()
    DoEvents
    
    Call 일일마감_Proc
    Call 의류접수_Display
    
    dtpDay.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyChk (KeyCode)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    If Not nDayCloseChk Then
        '이전소스 2010-05-04
        'dtpDay.Value = Format(strDayClose, "YYYY-MM-DD")
        
        If strDayClose = "" Then
            dtpDay.Value = Format(Date, "YYYY-MM-DD")
        Else
            dtpDay.Value = Format(strDayClose, "YYYY-MM-DD")
        End If
    Else
        dtpDay.Value = Format(Date, "YYYY-MM-DD")
    End If
                            
    Call Manager_Display(cboManager)
    
    cboManager.Text = strManager & ""
    
    'Call Data_Display
    'Call 의류접수_Display

    Call 일일마감_Proc
    Call 의류접수_Display

    g_AgencyCode = 가맹점정보.택코드 '가맹점 코드
End Sub

Private Sub Form_Resize()
    'On Error Resume Next
    
End Sub

'-----------------------------------------------------------
'+  기간할인    1
'+  목요세일    2
'+  정상        3
'-----------------------------------------------------------
Private Sub Sale_Check()
    Dim chkWeekDay As Integer
    
    '-----------------------------------------------------------
    ' TB_할인정보
    '-----------------------------------------------------------
    Query = "SELECT * FROM TB_할인정보"
    Query = Query & " WHERE 시작일자 <= '" & 마감일자 & "' "
    Query = Query & "   AND 종료일자 >= '" & 마감일자 & "' "
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If Not SUBRs.EOF Then
        SUBRs.Close
        Set SUBRs = Nothing
        
        chkSale = "1"       ' 기간할인
        
        Exit Sub
    End If
    SUBRs.Close
    Set SUBRs = Nothing
    
    '-----------------------------------------------------------
    ' TB_기본정보
    '-----------------------------------------------------------
    Query = "SELECT 요일할인 FROM TB_기본정보"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    If Not ADORs.EOF Then
        i = Weekday(Date)
        
        If Mid(ADORs(0), i, 1) = "1" Then
            chkSale = "2"      ' 요일할인
        Else
            chkSale = "3"      ' 정상
        End If
    End If
    ADORs.Close
    Set ADORs = Nothing
End Sub

