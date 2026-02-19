VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm판매취소결제 
   BorderStyle     =   1  '단일 고정
   Caption         =   "결제 취소"
   ClientHeight    =   7140
   ClientLeft      =   1815
   ClientTop       =   3420
   ClientWidth     =   5955
   ControlBox      =   0   'False
   DrawWidth       =   3
   FillColor       =   &H00C0C0C0&
   Icon            =   "frm판매취소결제.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form10"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   5955
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7140
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   12594
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm판매취소결제.frx":0A02
      Begin Threed.SSPanel SSPanel1 
         Height          =   585
         Left            =   15
         TabIndex        =   1
         Top             =   6540
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   1032
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton btnExit 
            Height          =   480
            Left            =   3525
            TabIndex        =   2
            Top             =   60
            Width           =   2355
            _Version        =   851970
            _ExtentX        =   4154
            _ExtentY        =   847
            _StockProps     =   79
            Caption         =   " 현금으로 판매취소(&X)"
            ForeColor       =   0
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
            Picture         =   "frm판매취소결제.frx":0AB4
         End
         Begin XtremeSuiteControls.PushButton btnCancel 
            Height          =   480
            Left            =   45
            TabIndex        =   10
            Top             =   60
            Width           =   1980
            _Version        =   851970
            _ExtentX        =   3492
            _ExtentY        =   847
            _StockProps     =   79
            Caption         =   " 판매취소 취소(&A)"
            ForeColor       =   0
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
            Picture         =   "frm판매취소결제.frx":104E
         End
         Begin VB.Label lblNum 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "0"
            Height          =   180
            Left            =   2100
            TabIndex        =   9
            Top             =   345
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Label lblCode 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "0"
            Height          =   180
            Left            =   2100
            TabIndex        =   8
            Top             =   90
            Visible         =   0   'False
            Width           =   90
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   405
         Index           =   1
         Left            =   15
         TabIndex        =   3
         Top             =   15
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   714
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " 신용카드 승인내역"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm판매취소결제.frx":15E8
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin FPSpreadADO.fpSpread sprCard 
         Height          =   1980
         Left            =   15
         TabIndex        =   4
         Top             =   435
         Width           =   5925
         _Version        =   524288
         _ExtentX        =   10451
         _ExtentY        =   3493
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         EditModePermanent=   -1  'True
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
         MaxCols         =   13
         SpreadDesigner  =   "frm판매취소결제.frx":180E
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   420
         Index           =   2
         Left            =   15
         TabIndex        =   5
         Top             =   2430
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   741
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " 현금영수증 승인내역"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm판매취소결제.frx":208F
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton btnCashCancel 
            Height          =   390
            Left            =   3900
            TabIndex        =   6
            Top             =   15
            Width           =   2010
            _Version        =   851970
            _ExtentX        =   3545
            _ExtentY        =   688
            _StockProps     =   79
            Caption         =   "현금영수증 승인취소"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
      End
      Begin FPSpreadADO.fpSpread sprCash 
         Height          =   3660
         Left            =   15
         TabIndex        =   7
         Top             =   2865
         Width           =   5925
         _Version        =   524288
         _ExtentX        =   10451
         _ExtentY        =   6456
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayColHeaders=   0   'False
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
         MaxCols         =   1
         MaxRows         =   11
         RowHeaderDisplay=   0
         ScrollBars      =   0
         SpreadDesigner  =   "frm판매취소결제.frx":22B5
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "frm판매취소결제"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    결제취소여부 = False
    
    판매취소여부 = True
    
    Unload Me
End Sub

Private Sub btnCashCancel_Click()
    sprCash.Row = 1
    sprCash.Col = 1
    
    If sprCash.Text = "" Then Exit Sub
    Unload frmKSNETCash
    Account_Form = "판매취소"
    
    With frmKSNETCash.sprGrid
        .Col = 1
    
        .Row = 1:  .Text = Spread_GetData(sprCash, 1, 1, True)   '승인번호
        .Row = 2:  .Text = Spread_GetData(sprCash, 2, 1, True)   '승인일자
        .Row = 3:  .Text = Spread_GetData(sprCash, 3, 1, True)   '승인시간
        .Row = 4:  .Text = Spread_GetData(sprCash, 4, 1, True)   '거래유형 '입력방법
        .Row = 5:  .Text = Spread_GetData(sprCash, 5, 1, True)   '총금액
        .Row = 6:  .Text = Spread_GetData(sprCash, 6, 1, True)   '사용자정보
        .Row = 7:  .Text = Spread_GetData(sprCash, 7, 1, True)   '메시지1
        .Row = 8:  .Text = Spread_GetData(sprCash, 8, 1, True)   '메시지2
        .Row = 9:  .Text = Spread_GetData(sprCash, 9, 1, True)   '소득구분
        .Row = 10: .Text = Spread_GetData(sprCash, 10, 1, True)  '국세청1
        .Row = 11: .Text = Spread_GetData(sprCash, 11, 1, True)  '국세청2
    End With

    frmKSNETCash.pnlCustomCode.Caption = lblCode.Caption                        '고객코드
    frmKSNETCash.pnlNum.Caption = lblNum.Caption                                '접수번호
    frmKSNETCash.txtMoney.Value = Spread_GetData(sprCash, 5, 1, True)           '총금액

    frmKSNETCash.pnlApprovalNo.Caption = Spread_GetData(sprCash, 1, 1, True)   '승인번호
    frmKSNETCash.pnlApprovalDay.Caption = Spread_GetData(sprCash, 2, 1, True)  '승인일자
    frmKSNETCash.pnlApprovalTime.Caption = Spread_GetData(sprCash, 3, 1, True) '승인시간

    Call frmKSNETCash.현금영수증승인요청_Rtn("4")
    
    frmKSNETCash.Show 1
End Sub

Private Sub btnExit_Click()
    결제취소여부 = True
    판매취소현금반환 = False
    
    If sprCard.MaxRows > 0 Then
                Query = "카드결제 승인취소를 하지 않았습니다." & vbNewLine & vbNewLine
        Query = Query & "카드결제 승인취소를 안하시겠습니까?" & vbNewLine & vbNewLine
        Query = Query & "※카드결제 승인취소를 안하면 세탁물 판매취소만 되고" & vbNewLine
        Query = Query & "  카드결제 금액은 남게 됩니다." & vbNewLine
        
        Rtn = MsgBox(Query, vbQuestion + vbYesNo, "확인")
        
        If Rtn = vbNo Then Exit Sub
    
        결제취소여부 = False
        판매취소현금반환 = True
    End If
    
    sprCash.Row = 1
    sprCash.Col = 1
    If sprCash.Text <> "" Then
                Query = "현금영수증 승인취소를 하지 않았습니다." & vbNewLine & vbNewLine
        Query = Query & "현금영수증 승인취소를 안하시겠습니까?" & vbNewLine & vbNewLine
        
        Rtn = MsgBox(Query, vbQuestion + vbYesNo, "확인")
        
        If Rtn = vbNo Then Exit Sub
        
        결제취소여부 = False
    End If
    
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim sCash As String
    
    
    
    sprCash.Row = 1
    sprCash.Col = 1
    
    If (sprCard.MaxRows = 0) And (sprCash.Text = "") Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()

    Me.Top = ((frmMain.Height - Me.Height) / 2) + frmMain.Top
    Me.Left = ((frmMain.Width - Me.Width) / 2) + frmMain.Left
    
    With sprCard
        .MaxRows = 0
        .RowHeight(-1) = 18
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeNormal
    End With
    
    With sprCash
        .RowHeight(-1) = 14
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeNormal
    End With

    판매취소여부 = False
End Sub

Private Sub sprCard_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    If Row <= 0 Then Exit Sub
    
    
    Account_Form = "판매취소"
    
    With frmKSNET2.sprGrid
        .Col = 1
    
        .Row = 1:  .Text = Spread_GetData(sprCard, Row, 2, True)  '승인번호
        .Row = 2:  .Text = Spread_GetData(sprCard, Row, 3, True)   '승인일자
        .Row = 3:  .Text = Spread_GetData(sprCard, Row, 4, True)   '승인시간
        
        .Row = 4:  .Text = Spread_GetData(sprCard, Row, 5, True)   '할부기간
        .Row = 5:  .Text = Spread_GetData(sprCard, Row, 6, True)   '결제금액
        
        .Row = 6:  .Text = Spread_GetData(sprCard, Row, 7, True)   '발급사코드
        .Row = 7:  .Text = Spread_GetData(sprCard, Row, 8, True)   '발급사명
        .Row = 8:  .Text = Spread_GetData(sprCard, Row, 9, True)   '매입사코드
        .Row = 9:  .Text = Spread_GetData(sprCard, Row, 10, True)  '매입사명
        .Row = 10: .Text = Spread_GetData(sprCard, Row, 11, True)  '카드번호
        .Row = 11: .Text = Spread_GetData(sprCard, Row, 12, True)  '메시지1
        .Row = 12: .Text = Spread_GetData(sprCard, Row, 13, True)  '메시지2
    End With

    frmKSNET2.pnlCustomCode.Caption = lblCode.Caption                '고객코드
    frmKSNET2.pnlNum.Caption = lblNum.Caption                        '접수번호
    
    frmKSNET2.txtMoney.ReadOnly = True
    frmKSNET2.txtMoney.Value = Spread_GetData(sprCard, Row, 6, True) '결제금액
    
    frmKSNET2.pnlApprovalNo.Caption = Spread_GetData(sprCard, Row, 2, True)   '승인번호
    frmKSNET2.pnlApprovalDay.Caption = Spread_GetData(sprCard, Row, 3, True)  '승인일자
    frmKSNET2.pnlApprovalTime.Caption = Spread_GetData(sprCard, Row, 4, True) '승인시간
    
    Call frmKSNET2.신용카드승인요청_Rtn("2")
    
    frmKSNET2.Show 1
End Sub

Public Sub Data_Display()
    On Error GoTo ErrRtn

    '-------------------------------------------------------------------------------------------
    '
    '-------------------------------------------------------------------------------------------
    Query = "SELECT * FROM TB_신용카드승인"
    Query = Query & " WHERE 고객코드 = '" & lblCode.Caption & "'"
    Query = Query & "   AND 접수번호 =  " & lblNum.Caption
    Query = Query & "   AND SUBSTRING(메시지2,1,2) = 'OK'"
    Query = Query & " ORDER BY 승인일자, 승인시간 ASC"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprCard
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 2:  .Text = ADORs!승인번호 & ""
            .Col = 3:  .Text = ADORs!승인일자 & ""
            .Col = 4:  .Text = ADORs!승인시간 & ""
            .Col = 5:  .Text = ADORs!할부기간 & ""
            .Col = 6:  .Text = ADORs!결제금액 & ""
            .Col = 7:  .Text = ADORs!발급사코드 & "" '
            .Col = 8:  .Text = ADORs!카드종류명 & "" '
            .Col = 9:  .Text = ADORs!매입사코드 & "" '
            .Col = 10: .Text = ADORs!매입사명 & "" '
            .Col = 11: .Text = Left(ADORs!카드번호, 16) & "" '
            .Col = 12: .Text = ADORs!메시지1 & "" '
            .Col = 13: .Text = ADORs!메시지2 & "" '
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        .ReDraw = True
    End With
        
        
    '-------------------------------------------------------------------------------------------
    '
    '-------------------------------------------------------------------------------------------
    Query = "SELECT * FROM TB_현금영수증"
    Query = Query & " WHERE 고객코드 = '" & lblCode.Caption & "'"
    Query = Query & "   AND 접수번호 =  " & lblNum.Caption
    Query = Query & "   AND SUBSTRING(메시지2,1,2) = 'OK'"
    Query = Query & " ORDER BY 승인일자, 승인시간 ASC"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF Then
        With sprCash
            .Col = 1
                        
            For i = 1 To 11
                .Row = i: .Text = ""
            Next i
        End With
    Else
        With sprCash
            .Col = 1
            
            .Row = 1:  .Text = ADORs!승인번호 & ""
            .Row = 2:  .Text = ADORs!승인일자 & ""
            .Row = 3:  .Text = ADORs!승인시간 & ""
            .Row = 4:  .Text = ADORs!거래유형 & "" '입력방법
            .Row = 5:  .Text = ADORs!총금액 & ""
            .Row = 6:  .Text = ADORs!사용자정보 & ""
            .Row = 7:  .Text = ADORs!메시지1 & ""
            .Row = 8:  .Text = ADORs!메시지2 & ""
            .Row = 9:  .Text = ADORs!소득구분 & ""
            .Row = 10: .Text = ADORs!국세청1 & ""
            .Row = 11: .Text = ADORs!국세청2 & ""
        End With
    End If
    ADORs.Close
    Set ADORs = Nothing
    
    Exit Sub
    
ErrRtn:
    Screen.MousePointer = 0
    
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub
