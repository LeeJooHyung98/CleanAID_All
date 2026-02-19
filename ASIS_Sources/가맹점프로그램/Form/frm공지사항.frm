VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm공지사항 
   BorderStyle     =   1  '단일 고정
   Caption         =   "확인 사항"
   ClientHeight    =   10035
   ClientLeft      =   7710
   ClientTop       =   2415
   ClientWidth     =   14340
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm공지사항.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   14340
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   10035
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14340
      _ExtentX        =   25294
      _ExtentY        =   17701
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   0
      PaneTree        =   "frm공지사항.frx":08CA
      Begin Threed.SSPanel SSPanel2 
         Height          =   345
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   420
         Width           =   14340
         _ExtentX        =   25294
         _ExtentY        =   609
         _Version        =   262144
         Font3D          =   3
         ForeColor       =   12582912
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "    지사반품"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm공지사항.frx":0A1C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image2 
            Height          =   240
            Index           =   0
            Left            =   75
            Picture         =   "frm공지사항.frx":0E7E
            Top             =   45
            Width           =   240
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   405
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   14340
         _ExtentX        =   25294
         _ExtentY        =   714
         _Version        =   262144
         PictureFrames   =   1
         Picture         =   "frm공지사항.frx":1880
         PictureBackgroundStyle=   2
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Label lblDay 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "2009-12-31"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   435
            TabIndex        =   5
            Top             =   120
            Width           =   1050
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   90
            Picture         =   "frm공지사항.frx":20AEA
            Top             =   90
            Width           =   240
         End
         Begin VB.Label lblWeek 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "(목)"
            DataField       =   "규격"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   1605
            TabIndex        =   4
            Top             =   120
            Width           =   405
         End
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   1260
         Index           =   1
         Left            =   0
         TabIndex        =   2
         Top             =   2700
         Width           =   14340
         _Version        =   524288
         _ExtentX        =   25294
         _ExtentY        =   2222
         _StockProps     =   64
         BackColorStyle  =   1
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
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
         MaxCols         =   6
         ScrollBars      =   2
         ShadowColor     =   14737632
         SpreadDesigner  =   "frm공지사항.frx":21074
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   1545
         Index           =   0
         Left            =   0
         TabIndex        =   7
         Top             =   780
         Width           =   14340
         _Version        =   524288
         _ExtentX        =   25294
         _ExtentY        =   2725
         _StockProps     =   64
         BackColorStyle  =   1
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
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
         MaxCols         =   10
         ScrollBars      =   2
         ShadowColor     =   14737632
         SpreadDesigner  =   "frm공지사항.frx":21730
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   345
         Index           =   1
         Left            =   0
         TabIndex        =   8
         Top             =   2340
         Width           =   14340
         _ExtentX        =   25294
         _ExtentY        =   609
         _Version        =   262144
         Font3D          =   3
         ForeColor       =   12582912
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "    가맹점입금"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm공지사항.frx":21E98
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton btnSelect 
            Height          =   315
            Index           =   1
            Left            =   13305
            TabIndex        =   15
            Top             =   15
            Width           =   1005
            _Version        =   851970
            _ExtentX        =   1773
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "전체선택"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Image Image2 
            Height          =   240
            Index           =   1
            Left            =   75
            Picture         =   "frm공지사항.frx":222FA
            Top             =   45
            Width           =   240
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   345
         Index           =   2
         Left            =   0
         TabIndex        =   9
         Top             =   3975
         Width           =   14340
         _ExtentX        =   25294
         _ExtentY        =   609
         _Version        =   262144
         Font3D          =   3
         ForeColor       =   12582912
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "    부자재"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm공지사항.frx":22CFC
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton btnSelect 
            Height          =   315
            Index           =   2
            Left            =   13305
            TabIndex        =   16
            Top             =   15
            Width           =   1005
            _Version        =   851970
            _ExtentX        =   1773
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "전체선택"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Image Image2 
            Height          =   240
            Index           =   2
            Left            =   75
            Picture         =   "frm공지사항.frx":2315E
            Top             =   45
            Width           =   240
         End
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   1860
         Index           =   2
         Left            =   0
         TabIndex        =   10
         Top             =   4335
         Width           =   14340
         _Version        =   524288
         _ExtentX        =   25294
         _ExtentY        =   3281
         _StockProps     =   64
         BackColorStyle  =   1
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
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
         MaxCols         =   11
         ScrollBars      =   2
         ShadowColor     =   14737632
         SpreadDesigner  =   "frm공지사항.frx":23B60
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   345
         Index           =   3
         Left            =   0
         TabIndex        =   11
         Top             =   6210
         Width           =   14340
         _ExtentX        =   25294
         _ExtentY        =   609
         _Version        =   262144
         Font3D          =   3
         ForeColor       =   192
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "    본사출고확정 (세탁물중 가맹점 미입고 내역)"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm공지사항.frx":243BB
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton btnSelect 
            Height          =   315
            Index           =   3
            Left            =   13305
            TabIndex        =   14
            Top             =   15
            Width           =   1005
            _Version        =   851970
            _ExtentX        =   1773
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "전체선택"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Image Image2 
            Height          =   240
            Index           =   3
            Left            =   75
            Picture         =   "frm공지사항.frx":2481D
            Top             =   45
            Width           =   240
         End
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   2835
         Index           =   3
         Left            =   0
         TabIndex        =   12
         Top             =   6570
         Width           =   14340
         _Version        =   524288
         _ExtentX        =   25294
         _ExtentY        =   5001
         _StockProps     =   64
         BackColorStyle  =   1
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
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
         MaxCols         =   11
         ScrollBars      =   2
         ShadowColor     =   14737632
         SpreadDesigner  =   "frm공지사항.frx":2521F
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   615
         Index           =   4
         Left            =   0
         TabIndex        =   13
         Top             =   9420
         Width           =   14340
         _ExtentX        =   25294
         _ExtentY        =   1085
         _Version        =   262144
         Font3D          =   3
         BackColor       =   16777215
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton btnOK 
            Height          =   495
            Left            =   12795
            TabIndex        =   0
            Top             =   60
            Width           =   1485
            _Version        =   851970
            _ExtentX        =   2619
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   " 확정(&O)"
            Appearance      =   6
            Picture         =   "frm공지사항.frx":25A23
         End
      End
   End
End
Attribute VB_Name = "frm공지사항"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnOK_Click()
    On Error GoTo ErrRtn
    
    '------------------------------------------------------------------
    ' 미입고 사유 입력을 안하면 안넘어감...
    '------------------------------------------------------------------
    With sprGrid(3)
        For i = 1 To .MaxRows
            .Row = i
            .Col = 10
            
            If Trim(.Text) = "" Then
                MsgBox "미입고 사유를 입력하세요.", vbCritical, "확인"
                
                Exit For
                
                'Exit Sub '미입고 내역이 많아 처리를 못하는 경우가 있을것 같아...
            End If
        Next i
    End With
    
    Call Sub_가맹점입금확정 '1
    Call Sub_부자재주문확정 '2
    Call Sub_본사출고확정   '3
        
    Unload Me
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
    
    Screen.MousePointer = 0
End Sub

'--------------------------------------------------------------------------------------
' 함수명 : Sub_가맹점입금확정
' 기  능 : 예정일자가 넘어갔는데도 지사로부터 입고가 되는 않은 세탁물을 조회...
'--------------------------------------------------------------------------------------
Private Sub Sub_가맹점입금확정()
    Dim 가맹점코드  As String
    Dim 입금일자    As String
    
    With sprGrid(1)
        If .MaxRows >= 0 Then
            For i = 1 To .MaxRows
                .Row = i
                .Col = 6
                
                If .Text = "1" Then
                    .Col = 1: 가맹점코드 = .Text & ""     '
                    .Col = 2: 입금일자 = .Text & ""       '
                
                    Query = "UPDATE TB_가맹점입금 SET 입금확정 = 'Y'"
                    Query = Query & "               , 확정일자 = '" & Format(Date, "YYYY-MM-DD hh:mm:ss") & "'"
                    Query = Query & " WHERE 가맹점코드 = '" & 가맹점코드 & "'"
                    Query = Query & "   AND 입금일자   = '" & 입금일자 & "'"
                    ADOCon.Execute Query
                End If
            Next i
        End If
    End With

    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
    
    Screen.MousePointer = 0
End Sub

'--------------------------------------------------------------------------------------
' 함수명 : Sub_부자재주문확정
' 기  능 :
'--------------------------------------------------------------------------------------
Private Sub Sub_부자재주문확정()
    Dim iOrderNo As Long
    Dim iSEQ     As Long
    
    With sprGrid(2)
        If .MaxRows >= 0 Then
            For i = 1 To .MaxRows
                .Row = i
                .Col = 11
                
                If .Text = "1" Then
                    .Col = 1:  iOrderNo = .Text & ""   '주문코드
                
                    '----------------------------------------------------------------------------------
                    ' 서버에 직접 입력한다.
                    '----------------------------------------------------------------------------------
                    Query = "UPDATE TB_부자재주문 SET 입고확정 = 'Y'"
                    Query = Query & "               , 확정일자 = '" & Format(Date, "YYYY-MM-DD") & "'"
                    Query = Query & " WHERE 주문코드 = " & iOrderNo
                    ADOCon.Execute Query
                End If
            Next i
        End If
    End With

    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
    
    Screen.MousePointer = 0
End Sub

Private Sub Sub_본사출고확정()
    Dim strTagNo       As String
    Dim strReceiveDate As String
    Dim strMemo        As String

    On Error GoTo ErrRtn
    
    With sprGrid(3)
        If .MaxRows >= 0 Then
            For i = 1 To .MaxRows
                .Row = i
                .Col = 11
                
                If .Text = "1" Then
                    .Col = 1:  strReceiveDate = .Text & ""     '접수일자
                    .Col = 2:  strTagNo = .Text & ""           '택번호
                    .Col = 10: strMemo = SubSQuotA(.Text) & "" '미입고사유
                
                    Query = "UPDATE TB_입출고 SET 미입고사유   = '" & strMemo & "'"
                    Query = Query & "           , 본사전송여부 = ''"
                    Query = Query & " WHERE 접수일자 = '" & strReceiveDate & "'"
                    Query = Query & "   AND 택번호   = '" & strTagNo & "'"
                    ADOCon.Execute Query
                End If
            Next i
        End If
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("frm공지사항.Sub_본사출고확정", Err.Source, Err.Number, Err.description)
    
    Screen.MousePointer = 0
End Sub

Private Sub btnSelect_Click(Index As Integer)

    Select Case Index
        Case 1
            With sprGrid(Index)
                For i = 1 To .MaxRows
                    .Row = i
                    .Col = 6: .Text = "1"
                Next i
            End With
        
        Case 2
            With sprGrid(Index)
                For i = 1 To .MaxRows
                    .Row = i
                    .Col = 11: .Text = "1"
                Next i
            End With
            
        Case 3
            With sprGrid(Index)
                For i = 1 To .MaxRows
                    .Row = i
                    .Col = 11: .Text = "1"
                Next i
            End With
    End Select
    
End Sub

Private Sub Form_Activate()
    If sprGrid(0).MaxRows = 0 And sprGrid(1).MaxRows = 0 And sprGrid(2).MaxRows = 0 And sprGrid(3).MaxRows = 0 Then
        Unload frm공지사항
    End If

''    If sprGrid(0).MaxRows = 0 And sprGrid(1).MaxRows = 0 And sprGrid(2).MaxRows = 0 And sprGrid(3).MaxRows = 0 Then
''        Unload frm공지사항
''        DoEvents
''
''        frmSplash.Show     '메세지
''        frmSplash.Refresh
''
''        frmSplash.lblMsg.Caption = "잠시만 기다려 주세요..."
''        frmSplash.ProgressBar1.MAX = 100
''        frmSplash.ProgressBar1.Min = 0
''        frmSplash.ProgressBar1.Value = 0
''
''        Load frmMain       '메인화면
''
''        Unload frmSplash
''        DoEvents
''
''        frmMain.Show
''    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
      
    For i = 0 To 3
        With sprGrid(i)
            .MaxRows = 0
            .RowHeight(-1) = 13
            
            'Spread 8 - 디자인
            .HighlightHeaders = HighlightHeadersOff
            .AppearanceStyle = AppearanceStyleEnhanced
            .ScrollBarStyle = ScrollBarStyleVisualStyle
            
            '선택된 Row
            .SelBackColor = &HFFFFC0 '황색 ^^
            .SelForeColor = &H0&     '검은글씨
            .OperationMode = OperationModeNormal
            
            '홀수/짝수 Row BankColor
            'Ret = .SetOddEvenRowColor(&HFFFFFF, &H80000008, &H80FFFF, &H80000008)
    
            'Init the User Sort
            .UserColAction = UserColActionSort
        End With
    Next i
    
    lblDay.Caption = Format(Date, "YYYY-MM-DD")
    lblWeek.Caption = Fun_Week(lblDay.Caption)
    
    '-----------------------------------------------------------
    ' 미입고 사유
    '-----------------------------------------------------------
    Dim tmp As String
    
    Query = "SELECT * FROM TB_미입고사유"
    Query = Query & " ORDER BY 미입고사유 ASC"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    tmp = ""
    
    Do Until ADORs.EOF
        If tmp = "" Then
            tmp = Trim(ADORs!미입고사유)
        Else
            tmp = tmp + Chr$(9) + Trim(ADORs!미입고사유)
        End If
        
        ADORs.MoveNext
    Loop
    ADORs.Close
    Set ADORs = Nothing
    
    With sprGrid(3)
        .Row = -1
        .Col = 10: .TypeComboBoxList = tmp
    End With
    
    '-----------------------------------------------------------
    
    If Server_Connection(HostCon) = False Then Exit Sub

    Call 반품_Display       '
    
    Call 미입고_Display     '
    
    Call 가맹점입금_Display '
    
    Call 부자재주문_Display '
            
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
    
    Screen.MousePointer = 0
End Sub

Private Sub 반품_Display()
    On Error GoTo ErrRtn
    
    Query = "SELECT    A.*"
    Query = Query & ", B.성명"
    Query = Query & ", B.전화번호"
    Query = Query & ", B.휴대전화"
    Query = Query & " FROM TB_입출고 AS A LEFT OUTER JOIN TB_고객정보 AS B ON A.고객코드 = B.고객코드"
    Query = Query & " WHERE A.지사출고상태 = '2'"
    Query = Query & "   AND (A.가맹점입고일자 IS NOT NULL OR A.가맹점입고일자 <> '')"
    Query = Query & "   AND ((A.반품환불일자 IS NULL OR A.반품환불일자 = '')"
    Query = Query & "    OR  (A.세탁환불일자 IS NULL OR A.세탁환불일자 = ''))"
    Query = Query & " ORDER BY A.접수일자, A.택번호 ASC"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    With sprGrid(0)
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = ADORs!접수일자 & ""                      ' 1
            .Col = 2:  .Text = Format(ADORs!택번호, "000-00-0000") & "" ' 2
            .Col = 3:  .Text = ADORs!성명 & ""                          ' 3
            .Col = 4:  .Text = ADORs!전화번호 & ""                      ' 4
            .Col = 5:  .Text = ADORs!휴대전화 & ""                      ' 5
            .Col = 6:  .Text = ADORs!의류명 & ""                        ' 6
            .Col = 7:  .Text = ADORs!색상 & ""                          ' 7
            .Col = 8:  .Text = ADORs!무늬 & ""                          ' 8
            .Col = 9:  .Text = ADORs!내용 & ""                          ' 9
            .Col = 10: .Text = ADORs!상표 & ""                          '10
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("frm공지사항.반품_Display", Err.Source, Err.Number, Err.description)
    
    Screen.MousePointer = 0
End Sub

Private Sub 가맹점입금_Display()
    On Error GoTo ErrRtn
    
    Query = "SELECT * FROM TB_가맹점입금"
    Query = Query & " WHERE 가맹점코드 = '" & 가맹점정보.가맹점코드 & "'"
    Query = Query & "   AND (입금확정 = '' OR 입금확정 IS NULL)"
    Query = Query & " ORDER BY 입금일자 ASC"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    With sprGrid(1)
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = ADORs!가맹점코드 & "" ' 1
            .Col = 2:  .Text = ADORs!입금일자 & ""   ' 2
            .Col = 3:  .Text = ADORs!배송기사명 & "" ' 3
            .Col = 4:  .Text = ADORs!입금액 & ""     ' 4
            .Col = 5:  .Text = ADORs!비고 & ""       ' 5
            .Col = 6:  .Text = "0"                   ' 6
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("frm공지사항.가맹점입금_Display", Err.Source, Err.Number, Err.description)
    
    Screen.MousePointer = 0
End Sub

'
Private Sub 부자재주문_Display()
    On Error GoTo ErrRtn
    
    Query = "SELECT * FROM TB_부자재주문"
    Query = Query & " WHERE 가맹점코드 = '" & 가맹점정보.가맹점코드 & "'"
    Query = Query & "   AND (입고확정 = '' OR 입고확정 IS NULL)"
    Query = Query & " ORDER BY 주문코드 ASC"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    With sprGrid(2)
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = ADORs!주문코드 & ""   ' 1
            .Col = 2:  .Text = ADORs!주문일자 & ""   ' 3
            .Col = 3:  .Text = ADORs!부자재명 & ""   ' 4
            .Col = 4:  .Text = ADORs!규격 & ""       ' 5
            .Col = 5:  .Text = ADORs!수량 & ""       ' 6
            .Col = 6:  .Text = ADORs!단가 & ""       ' 6
            .Col = 7:  .Text = ADORs!공급가액 & ""   ' 7
            .Col = 8:  .Text = ADORs!세액 & ""       ' 8
            .Col = 9:  .Text = ADORs!합계금액 & ""   ' 9
            .Col = 10: .Text = ADORs!출고일자 & ""   '10
            .Col = 11: .Text = "0"                   '11
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("frm공지사항.부자재주문_Display", Err.Source, Err.Number, Err.description)
End Sub

'--------------------------------------------------------------------------------------
' 함수명 : 미입고_Display
' 기  능 : 예정일자가 넘어갔는데도 지사로부터 입고가 되는 않은 세탁물을 조회...
'--------------------------------------------------------------------------------------
Private Sub 미입고_Display()
    On Error GoTo ErrRtn
    
    Query = "SELECT * FROM TB_입출고"
    Query = Query & " WHERE 예정일자 < '" & Format(Date, "YYYY-MM-DD") & "'"
    Query = Query & "   AND (출고일자 = '' OR 출고일자 IS NULL)"
    Query = Query & "   AND ((판매취소 <> 'Y')"
    Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
    Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
    Query = Query & "   AND (가맹점출고일자 IS NOT NULL OR 가맹점출고일자 <> '')"
    Query = Query & "   AND (가맹점입고일자 IS NULL OR 가맹점입고일자 = '')"
    Query = Query & " ORDER BY 접수일자, 택번호 ASC"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    With sprGrid(3)
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = ADORs!접수일자 & ""                     ' 1
            
            If Len(ADORs!택번호) = 9 Then
                .Col = 2:  .Text = Format(ADORs!택번호, "000-00-0000") ' 2
            Else
                .Col = 2:  .Text = ADORs!택번호 & ""                   ' 2
            End If
            
            .Col = 3:  .Text = ADORs!의류명 & ""                       ' 3
            .Col = 4:  .Text = ADORs!색상 & ""                         ' 4
            .Col = 5:  .Text = ADORs!무늬 & ""                         ' 5
            .Col = 6:  .Text = ADORs!내용 & ""                         ' 6
            .Col = 7:  .Text = ADORs!금액 & ""                         ' 7
            .Col = 8:  .Text = ADORs!상표 & ""                         ' 8
            .Col = 9:  .Text = Left(ADORs!지사출고일자, 10) & ""       ' 9
            .Col = 10: .Text = ""                                      '10
            .Col = 11: .Text = "0"                                     '11
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("frm공지사항.미입고_Display", Err.Source, Err.Number, Err.description)
End Sub
