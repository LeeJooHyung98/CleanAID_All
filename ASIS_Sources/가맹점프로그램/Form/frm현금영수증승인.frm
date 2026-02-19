VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm현금영수증승인 
   Caption         =   "현금영수증 승인현황"
   ClientHeight    =   11250
   ClientLeft      =   180
   ClientTop       =   3735
   ClientWidth     =   15915
   ControlBox      =   0   'False
   LinkTopic       =   "Form25"
   MDIChild        =   -1  'True
   ScaleHeight     =   11250
   ScaleWidth      =   15915
   Begin Threed.SSPanel pnlProg 
      Height          =   1215
      Left            =   45
      TabIndex        =   9
      Top             =   1605
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
      Picture         =   "frm현금영수증승인.frx":0000
      Caption         =   " "
      BorderWidth     =   0
      BevelOuter      =   1
      BevelInner      =   2
      Alignment       =   4
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   11250
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15915
      _ExtentX        =   28072
      _ExtentY        =   19844
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm현금영수증승인.frx":2FCB
      Begin Threed.SSPanel SSPanel2 
         Height          =   750
         Left            =   15
         TabIndex        =   2
         Top             =   450
         Width           =   15885
         _ExtentX        =   28019
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   0
            Left            =   4200
            TabIndex        =   4
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 조회(&F)"
            Appearance      =   6
            Picture         =   "frm현금영수증승인.frx":305D
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   12105
            TabIndex        =   5
            Top             =   45
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frm현금영수증승인.frx":3757
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   0
            Left            =   915
            TabIndex        =   0
            Top             =   60
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   56164355
            CurrentDate     =   40279
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   1
            Left            =   2625
            TabIndex        =   6
            Top             =   60
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   56164355
            CurrentDate     =   40279
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   4
            Left            =   8190
            TabIndex        =   17
            Top             =   60
            Width           =   1800
            _Version        =   851970
            _ExtentX        =   3175
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 가맹점용 출력"
            Appearance      =   6
            Picture         =   "frm현금영수증승인.frx":47E9
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   1
            Left            =   10020
            TabIndex        =   18
            Top             =   60
            Width           =   1800
            _Version        =   851970
            _ExtentX        =   3175
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 고객용 출력"
            Appearance      =   6
            Picture         =   "frm현금영수증승인.frx":4EE3
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   3
            Left            =   6660
            TabIndex        =   19
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 엑셀(&E)"
            Appearance      =   6
            Picture         =   "frm현금영수증승인.frx":55DD
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "결제일자:"
            Height          =   195
            Index           =   2
            Left            =   45
            TabIndex        =   8
            Top             =   120
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "~"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   2430
            TabIndex        =   7
            Top             =   120
            Width           =   120
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   3
         Top             =   15
         Width           =   15885
         _ExtentX        =   28019
         _ExtentY        =   741
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   4194304
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
         Caption         =   "      현금영수증 승인/취소 현황"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm현금영수증승인.frx":5D57
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm현금영수증승인.frx":5F7D
            Top             =   -15
            Width           =   765
         End
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   9435
         Left            =   15
         TabIndex        =   10
         Top             =   1215
         Width           =   15885
         _Version        =   524288
         _ExtentX        =   28019
         _ExtentY        =   16642
         _StockProps     =   64
         BackColorStyle  =   1
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
         MaxCols         =   20
         SpreadDesigner  =   "frm현금영수증승인.frx":6B47
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   570
         Left            =   15
         TabIndex        =   11
         Top             =   10665
         Width           =   15885
         _ExtentX        =   28019
         _ExtentY        =   1005
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin CSTextLibCtl.silgEdit txtMoney 
            Height          =   330
            Index           =   0
            Left            =   1050
            TabIndex        =   12
            Top             =   60
            Width           =   945
            _Version        =   262145
            _ExtentX        =   1667
            _ExtentY        =   582
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            ReadOnly        =   -1  'True
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "0"
            Text            =   " 0"
            StartText.x     =   3
            StartText.y     =   4
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   13
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            Undo            =   1
            Data            =   0
         End
         Begin CSTextLibCtl.silgEdit txtMoney 
            Height          =   330
            Index           =   1
            Left            =   3300
            TabIndex        =   13
            Top             =   60
            Width           =   945
            _Version        =   262145
            _ExtentX        =   1667
            _ExtentY        =   582
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            ReadOnly        =   -1  'True
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "0"
            Text            =   " 0"
            StartText.x     =   3
            StartText.y     =   4
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   13
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            Undo            =   1
            Data            =   0
         End
         Begin VB.Label Label 
            BackStyle       =   0  '투명
            Caption         =   "금액"
            Height          =   225
            Index           =   0
            Left            =   4980
            TabIndex        =   16
            Top             =   120
            Width           =   4875
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "건수"
            Height          =   225
            Index           =   7
            Left            =   210
            TabIndex        =   15
            Top             =   120
            Width           =   765
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "금액"
            Height          =   225
            Index           =   6
            Left            =   2460
            TabIndex        =   14
            Top             =   120
            Width           =   765
         End
      End
   End
End
Attribute VB_Name = "frm현금영수증승인"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 0: Call Data_Display
        Case 3: Call Export_Excel(frmMain.cdgExcel, sprGrid)
        Case 4, 1:
            If sprGrid.ActiveRow <= 0 Then Exit Sub
            
            Dim 상태 As String
            
            sprGrid.Row = sprGrid.ActiveRow
            sprGrid.Col = 2: 상태 = sprGrid.Text & ""
            
            '-------------------------------------------------------------------------------
            '
            '-------------------------------------------------------------------------------
            Dim CommPort As String
            Dim BaudRate As String
            
            CommPort = GetIniStr("VAN", "KS7500_CommPort", "", iniFile)
            BaudRate = GetIniStr("VAN", "KS7500_BaudRate", "", iniFile)
            
            Call 현금영수증재발행_Report(sprGrid.ActiveRow, 상태, IIf(Index = 4, 1, 2))
            DoEvents
        
        Case 5: Unload Me
    End Select
End Sub

Public Sub Data_Display()
    Dim nMoney(2)  As Double
    Dim nCount(2)  As Double
    
    On Error GoTo ErrRtn

    pnlProg.Visible = True
    DoEvents
    
    '-------------------------------------------------------------------------------------------
    '
    '-------------------------------------------------------------------------------------------
    Query = "SELECT    A.* "
    Query = Query & ", B.성명"
    Query = Query & ", B.전화번호"
    Query = Query & ", B.휴대전화"
    Query = Query & " FROM TB_현금영수증 AS A LEFT OUTER JOIN TB_고객정보 AS B ON A.고객코드 = B.고객코드"
    Query = Query & " WHERE (A.승인일자 >= '" & Format(dtpDay(0).Value, "YYMMDD") & "' "
    Query = Query & "   AND  A.승인일자 <= '" & Format(dtpDay(1).Value, "YYMMDD") & "') "
    'Query = Query & "   AND SUBSTRING(A.메시지2,1,2) <> '취소'"
    Query = Query & " ORDER BY A.승인일자, A.승인시간 ASC"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            nMoney(0) = nMoney(0) + CDbl(ADORs!총금액 & "")
            If Left(ADORs!메시지2, 2) = "OK" Then
                .Col = 2: .Text = "승인": .ForeColor = vbBlue
                nMoney(1) = nMoney(1) + CDbl(ADORs!총금액 & "")
                nCount(1) = nCount(1) + 1
            Else
                .Col = 2: .Text = "취소": .ForeColor = vbRed
                nMoney(2) = nMoney(2) + CDbl(ADORs!총금액 & "")
                nCount(2) = nCount(2) + 1
            End If
            
            .Col = 3:  .Text = ADORs!성명 & ""               '
            .Col = 4:  .Text = ADORs!전화번호 & ""           '
            .Col = 5:  .Text = ADORs!휴대전화 & ""           '
            .Col = 6:  .Text = ADORs!접수번호 & ""           '
            .Col = 7:  .Text = ADORs!승인번호 & ""           '
            .Col = 8:  .Text = ADORs!승인일자 & ""           '
            .Col = 9:  .Text = ADORs!승인시간 & ""           '
            .Col = 10: .Text = ADORs!입력방법 & ""           '
            .Col = 11: .Text = ADORs!거래유형 & ""           '개인(0), 사업자(1)
            .Col = 12: .Text = Format(ADORs!총금액 & "", "#,##0")            '
            .Col = 13: .Text = ADORs!사용자정보 & ""         '
            .Col = 14: .Text = ADORs!메시지1 & ""            '
            .Col = 15: .Text = ADORs!메시지2 & ""            '
            .Col = 16: .Text = ADORs!소득구분 & ""           '
            .Col = 17: .Text = ADORs!국세청1 & ""            '
            .Col = 18: .Text = ADORs!국세청2 & ""            '
            .Col = 19: .Text = ADORs!고객코드 & ""           '
            .Col = 20
            Select Case Trim(ADORs!취소사유 & "")
                Case "01": .Text = "01.거래취소"
                Case "02": .Text = "02.오류발급취소"
                Case "03": .Text = "03.기타"
                Case Else
            End Select
            
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        .ReDraw = True
    End With
    
    txtMoney(0).Value = Format(nCount(1), "#,##0")
    txtMoney(1).Value = nMoney(1)
    
    Label(0).Caption = "(" & "취소금액:" & Format(nMoney(2), "#,##0") & ", 건수:" & Format(nCount(2), "#,##0") & ") "
    
    pnlProg.Visible = False
    
    Exit Sub
    
ErrRtn:
    pnlProg.Visible = False
    
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub Form_Activate()
    Call Data_Display
End Sub

Private Sub Form_Load()
    With sprGrid
        .MaxRows = 0
        .RowHeight(-1) = 16
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeNormal
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
    
    dtpDay(0).Value = Date
    dtpDay(1).Value = Date
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    cmdBtn(5).Left = Me.Width - cmdBtn(5).Width - 200
End Sub

Private Sub sprGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    If Row <= 0 Then Exit Sub
    
    If Col = 1 Then
        sprGrid.Row = Row
        sprGrid.Col = 2
        If sprGrid.Text = "취소" Then
            MsgBox "이미 신용카드승인 취소된 상태입니다.", vbInformation, "확인"
            
            Exit Sub
        End If
        Unload frmKSNETCash
        Account_Form = "접수2"
        
        With frmKSNETCash.sprGrid
            .Col = 1
        
            .Row = 1:  .Text = Spread_GetData(sprGrid, Row, 7, True)   '승인번호
            .Row = 2:  .Text = Spread_GetData(sprGrid, Row, 8, True)   '승인일자
            .Row = 3:  .Text = Spread_GetData(sprGrid, Row, 9, True)   '승인시간
            .Row = 4:  .Text = Spread_GetData(sprGrid, Row, 11, True)  '거래유형
            .Row = 5:  .Text = Spread_GetData(sprGrid, Row, 12, True)  '총금액
            .Row = 6:  .Text = Spread_GetData(sprGrid, Row, 13, True)  '사용자정보
            .Row = 7:  .Text = Spread_GetData(sprGrid, Row, 14, True)  '메시지1
            .Row = 8:  .Text = Spread_GetData(sprGrid, Row, 15, True)  '메시지2
            .Row = 9:  .Text = Spread_GetData(sprGrid, Row, 16, True)  '소득구분
            .Row = 10: .Text = Spread_GetData(sprGrid, Row, 17, True)  '국세청1
            .Row = 11: .Text = Spread_GetData(sprGrid, Row, 18, True)  '국세청2
        End With
        
        '"KS4060 보안인증" 는 해당 단말기 에서 바로 출력 처리를 한다.
        If 가맹점정보.CAT단말기종류 <> "KS4060 보안인증" Then
            If Spread_GetData(sprGrid, Row, 11, True) = "0" Then
                frmKSNETCash.optGubun(0).Value = True
            Else
                frmKSNETCash.optGubun(1).Value = True
            End If
        Else
            If Spread_GetData(sprGrid, Row, 11, True) = "1" Then
                frmKSNETCash.optGubun(0).Value = True
            Else
                frmKSNETCash.optGubun(1).Value = True
            End If
        End If
        
        frmKSNETCash.pnlCustomCode.Caption = Spread_GetData(sprGrid, Row, 19, True)  '고객코드
        frmKSNETCash.pnlNum.Caption = Spread_GetData(sprGrid, Row, 6, True)          '접수번호
        frmKSNETCash.txtMoney.Value = Spread_GetData(sprGrid, Row, 12, True)         '총금액
        
        frmKSNETCash.pnlApprovalNo.Caption = Spread_GetData(sprGrid, Row, 7, True)   '승인번호
        frmKSNETCash.pnlApprovalDay.Caption = Spread_GetData(sprGrid, Row, 8, True)  '승인일자
        frmKSNETCash.pnlApprovalTime.Caption = Spread_GetData(sprGrid, Row, 9, True) '승인시간
        
        Call frmKSNETCash.현금영수증승인요청_Rtn("4")
        
        frmKSNETCash.Show 1
    End If
End Sub
