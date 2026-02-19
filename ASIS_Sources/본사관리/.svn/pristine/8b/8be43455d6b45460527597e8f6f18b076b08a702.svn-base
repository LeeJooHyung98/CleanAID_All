VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_04006 
   Caption         =   "수금 월마감"
   ClientHeight    =   11985
   ClientLeft      =   1275
   ClientTop       =   1920
   ClientWidth     =   16095
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_04006.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11985
   ScaleWidth      =   16095
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   11985
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16095
      _ExtentX        =   28390
      _ExtentY        =   21140
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_04006.frx":058A
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   1
         Top             =   540
         Width           =   16065
         _ExtentX        =   28337
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboOffice 
            Height          =   315
            Left            =   1245
            Locked          =   -1  'True
            Style           =   2  '드롭다운 목록
            TabIndex        =   13
            Top             =   60
            Width           =   3420
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   5460
            TabIndex        =   2
            Top             =   405
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "마감년월"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   10
            Left            =   60
            TabIndex        =   14
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "지    사"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   5460
            TabIndex        =   15
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "마감기간"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin XtremeSuiteControls.DateTimePicker dtinput 
            Height          =   330
            Index           =   2
            Left            =   6645
            TabIndex        =   19
            Top             =   405
            Width           =   1140
            _Version        =   851970
            _ExtentX        =   2011
            _ExtentY        =   582
            _StockProps     =   68
            CustomFormat    =   "yyyy-MM"
            Format          =   3
            UpDown          =   -1  'True
            CurrentDate     =   40544
         End
         Begin XtremeSuiteControls.DateTimePicker dtinput 
            Height          =   330
            Index           =   0
            Left            =   6645
            TabIndex        =   20
            Top             =   60
            Width           =   2850
            _Version        =   851970
            _ExtentX        =   5027
            _ExtentY        =   582
            _StockProps     =   68
            CurrentDate     =   40544
         End
         Begin XtremeSuiteControls.DateTimePicker dtinput 
            Height          =   330
            Index           =   1
            Left            =   9765
            TabIndex        =   18
            Top             =   60
            Width           =   2850
            _Version        =   851970
            _ExtentX        =   5027
            _ExtentY        =   582
            _StockProps     =   68
            CurrentDate     =   40544
         End
         Begin VB.Label Label 
            Alignment       =   2  '가운데 맞춤
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
            Height          =   225
            Left            =   9510
            TabIndex        =   16
            Top             =   120
            Width           =   255
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   3
         Top             =   15
         Width           =   8460
         _ExtentX        =   14923
         _ExtentY        =   900
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   255
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
         Caption         =   " 수금 월마감 (P_04006)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_04006.frx":061C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   1
         Left            =   8490
         TabIndex        =   4
         Top             =   15
         Width           =   7590
         _ExtentX        =   13388
         _ExtentY        =   900
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
         PictureBackgroundStyle=   2
         PictureBackground=   "P_04006.frx":081E
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   5
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "종료"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_04006.frx":0A20
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   6
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "엑셀"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_04006.frx":0FBA
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   7
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "인쇄"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_04006.frx":1554
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   8
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "취소"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_04006.frx":1AEE
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   9
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "삭제"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_04006.frx":2088
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   10
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "저장"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_04006.frx":2622
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   11
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "신규"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_04006.frx":2BBC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   12
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "조회"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_04006.frx":3156
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10635
         Left            =   15
         TabIndex        =   17
         Top             =   1335
         Width           =   16065
         _Version        =   524288
         _ExtentX        =   28337
         _ExtentY        =   18759
         _StockProps     =   64
         BackColorStyle  =   1
         ColsFrozen      =   3
         EditEnterAction =   5
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
         MaxCols         =   30
         SpreadDesigner  =   "P_04006.frx":36F0
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_04006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim RS02 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cboInput_Change(Index As Integer)

End Sub

Private Sub cboOffice_Click()
    If cboOffice.ListIndex < 0 Then Exit Sub
    
    '-----------------------------------------------------------------
    '
    '-----------------------------------------------------------------
    ReDim sValue(2)
    
    sValue(0) = Mid(cboOffice.Text, 2, 4)
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01001_00", sValue(), Err_Num, Err_Dec)
    
    With spdView
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            If (RS01!종료일자 = "") Or (RS01!종료일자 = "2099-12-31") Then '가맹점이 현 지사에서 관리중...
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                
                .Col = 1: .Text = RS01!가맹점코드 & ""
                .Col = 2: .Text = RS01!가맹점명 & ""
            End If
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
                
        .Redraw = True
    End With
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: 'Call Data_Display   ' 조회
        Case 1: 'Call DataAdd        ' 신규
        Case 2: Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
        Case 4: 'Call DataCancel     ' 취소
        Case 5: 'Call DataPrint      ' 인쇄
        Case 6: Call Export_Excel(P_00000.cdgExcel, spdView)      ' 엑셀
        Case 7: Unload Me           ' 종료
    End Select
    
'    Me.MousePointer = 0
    
    Exit Sub
    
ErrRtn:
    Me.MousePointer = 0
    
    If Err.Number = "0" Then
        
    ElseIf Err.Number = "91" Then
        End
    Else
        Resume Next
    End If
End Sub

Private Sub Form_Activate()
    cmdBtn(2).Enabled = True
    
    cmdBtn(5).Enabled = True
    cmdBtn(6).Enabled = True
    
    '본사 여부
    If HeadOffice = MASTER_OFFICE_CODE Then
        cboOffice.Locked = False
    Else
        cboOffice.Locked = True
    End If
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    dtInput(0).Value = Format(Date, "YYYY-MM-DD")
    dtInput(1).Value = Format(Date, "YYYY-MM-DD")

    dtInput(2).Value = Format(Date, "YYYY-MM")
    
    Call Get_지사리스트(cboOffice)
    
    Dim i As Integer
    
    With cboOffice
        For i = 0 To .ListCount - 1
            If Mid(.List(i), 2, 4) = HeadOffice Then
                .ListIndex = i
                
                Exit For
            End If
        Next i
    End With
    
''    If P_04006_Flag = False Then
''        If Store.Code = MASTER_OFFICE_CODE Then
''            panCaption(35).Visible = True
''            cboInput(3).Visible = True
''            Call Get_지사리스트(cboInput(3))
''            cboInput(3).AddItem "[0000] 전체 ", 0
''
''            cboInput(3).ListIndex = 2
''        End If
''
''        P_04006_Flag = True
''    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub DataSave()
    Dim i As Integer
    Dim 가맹점코드 As String
    
    With spdView
        For i = 1 To .MaxRows
            .Row = i
            
            .Col = 1: 가맹점코드 = .Text & ""                         '
            .Col = 3:  .Text = Format(dtInput(2).Value, "YYYY-MM")    '
            .Col = 4:  .Text = Format(dtInput(0).Value, "YYYY-MM-DD") '
            .Col = 5:  .Text = Format(dtInput(1).Value, "YYYY-MM-DD") '
            
            '------------------------------------------------------------------------
            '
            '------------------------------------------------------------------------
            ReDim sValue(4)
            
            sValue(0) = Mid(cboOffice.Text, 2, 4)
            sValue(1) = 가맹점코드
            sValue(2) = Format(dtInput(2).Value, "YYYY-MM")
            sValue(3) = Format(dtInput(0).Value, "YYYY-MM-DD")
            sValue(4) = Format(dtInput(1).Value, "YYYY-MM-DD")
            
            If HeadOffice = MASTER_OFFICE_CODE Then
                If DBOpen_Master(Mid(cboOffice.Text, 2, 4)) = False Then Exit Sub
                
                Set RS01 = New ADODB.Recordset
                Set RS01 = ExecProMaster("SP_04006_01", sValue(), Err_Num, Err_Dec)
            Else
                Set RS01 = New ADODB.Recordset
                Set RS01 = ExecPro("SP_04006_01", sValue(), Err_Num, Err_Dec)
            End If
            
            If Not RS01.EOF Then
                .Col = 6:  .Text = RS01!접수수량 & ""            ' 4
                .Col = 7:  .Text = RS01!출고수량 & ""            ' 5
                .Col = 8:  .Text = RS01!이전종료택번호 & ""      ' 7
                .Col = 9:  .Text = RS01!시작택번호 & ""          ' 6
                .Col = 10: .Text = RS01!종료택번호 & ""          ' 7
                .Col = 11: .Text = RS01!접수금액 & ""            ' 8
                .Col = 12: .Text = RS01!현금입금 + RS01!카드금액 & "" ' 9
                .Col = 13: .Text = RS01!현금입금 & ""  ' 9
                .Col = 14: .Text = RS01!카드금액 & ""            '10
                .Col = 15: .Text = RS01!카드건수 & ""            '11
                .Col = 16: .Text = RS01!쿠폰금액 & ""            '12
                .Col = 17: .Text = RS01!쿠폰건수 & ""            '13
                .Col = 18: .Text = RS01!발생마일리지 & ""        '14
                .Col = 19: .Text = RS01!사용마일리지 & ""        '15
                .Col = 20: .Text = RS01!삭제마일리지 & ""        '16
                .Col = 21: .Text = RS01!반품환불금액 & ""        '17
                .Col = 22: .Text = RS01!반품환불건수 & ""        '18
                .Col = 23: .Text = RS01!세탁환불금액 & ""        '19
                .Col = 24: .Text = RS01!세탁환불건수 & ""        '20
                .Col = 25: .Text = RS01!재세탁수량 & ""          '21
                .Col = 26: .Text = RS01!수선금액 & ""            '21
                .Col = 27: .Text = RS01!수선수량 & ""            '21
                .Col = 28: .Text = RS01!지사금액 & ""            '21
                .Col = 29: .Text = RS01!가맹점금액 & ""          '21
                .Col = 30: .Text = RS01!지사입금액 & ""          '21
            End If
            
            RS01.Close
            Set RS01 = Nothing
        Next i
    End With
    
''    ' 본사가 아닌메장은 이전 내용을 그래로 사용한다.
''    If Store.Code <> MASTER_OFFICE_CODE  Then
''        If MsgBox("마감작업을 진행하시겠습니까?", vbQuestion + vbYesNo) = vbYes Then
''            If DataSave_Master = True Then
''                MsgBox "해당되는 데이터가 정상적으로 처리되었습니다.", vbInformation
''            End If
''        End If
''
''    ' 본사일 경우 해당 루틴에서 모든 작업을 처리한다.
''    Else
''        If MsgBox("마감작업을 진행하시겠습니까?", vbQuestion + vbYesNo) = vbYes Then
''            Call DataSave_ALL
''        End If
''    End If
       
End Sub


Public Function DataSave_Master() As Boolean
'    Dim i As Integer
'    Dim Query As String
'
'    Dim sDate As String
'    Dim sDate2 As String
'
'    On Error GoTo SQLERROR
'
'    DataSave_Master = False
'
'    sDate = Format(dtInput(2).Value, "YYYY-MM") & "-01"
'    dtInput(0).Value = Format(sDate, "####-##-##")
'
'    sDate2 = Format(DateAdd("d", -1, DateAdd("m", 1, dtInput(0).Value)), "YYYY-MM-DD")
'    dtInput(1).Value = Format(sDate2, "####-##-##")
'
'    Query = "SELECT  A.AgencyCode        AS 대리점코드, "
'    Query = Query & "      B.AgencyName       AS 대리점명, "
'    Query = Query & "      Sum(A.IpSu)        AS 입고수량, "
'    Query = Query & "      Sum(A.ChulSu)      AS 출고수량, "
'    Query = Query & "      Min(CASE WHEN A.StartTag <> '    ' AND NOT A.StartTag IS NULL THEN A.StartTag END)    AS 시작택, "
'    Query = Query & "      Max(A.EndTag)      AS 종료택, "
'    Query = Query & "      Sum(A.Amount)      AS 금액, "
'    Query = Query & "      Sum(A.JaeSu)       AS 재세탁수량, "
'    Query = Query & "      Sum(A.SuSu)        AS 수선수량, "
'    Query = Query & "      Sum (A.BanSu)      AS 반품수량 "
'    Query = Query & "FROM    Sugeum      A (NOLOCK), "
'    Query = Query & "        AgencyCT    B (NOLOCK) "
'    Query = Query & "WHERE   A.AgencyCode = B.AgencyCode "
'    Query = Query & "AND     A.SuDate BETWEEN '" & sDate & "' AND '" & sDate2 & "' "
'    Query = Query & "GROUP BY    A.AgencyCode, "
'    Query = Query & "            B.AgencyName "
'
'    Set RS01 = New ADODB.Recordset
'    RS01.Open Query, ADOCon, adOpenStatic
'
'    mskInput(1).Text = RS01.RecordCount
'
'    Do While Not RS01.EOF
'        txtInput(0).Text = RS01!대리점코드
'        txtInput(1).Text = RS01!대리점명
'
'        i = i + 1
'        mskInput(2).Text = i
'
'        DoEvents
'
'        Query = "SELECT  Count(*)    AS 레코드건수 "
'        Query = Query & "FROM    SuGeumMST   A (NOLOCK) "
'        Query = Query & "WHERE   A.SYear      = '" & Format(dtInput(2).Value, "yyyy") & "' "
'        Query = Query & "AND     A.SMonth     = '" & Format(dtInput(2).Value, "mm") & "' "
'        Query = Query & "AND     A.AgencyCode = '" & RS01!대리점코드 & "' "
'
'        Set RS02 = New ADODB.Recordset
'        RS02.Open Query, ADOCon, adOpenStatic
'
'        If RS02!레코드건수 = 0 Then
'            Query = "INSERT INTO SuGeumMST "
'            Query = Query & "    (SYear, "
'            Query = Query & "     SMonth, "
'            Query = Query & "     AgencyCode, "
'            Query = Query & "     ISu, "
'            Query = Query & "     CSu, "
'            Query = Query & "     STag, "
'            Query = Query & "     ETag, "
'            Query = Query & "     Amount, "
'            Query = Query & "     JSu, "
'            Query = Query & "     SSu, "
'            Query = Query & "     BSu) "
'            Query = Query & "VALUES ('" & Format(dtInput(2).Value, "yyyy") & "', "
'            Query = Query & "        '" & Format(dtInput(2).Value, "mm") & "', "
'            Query = Query & "        '" & RS01!대리점코드 & "', "
'            Query = Query & "        " & RS01!입고수량 & ", "
'            Query = Query & "        " & RS01!출고수량 & ", "
'            Query = Query & "        '" & RS01!시작택 & "', "
'            Query = Query & "        '" & RS01!종료택 & "', "
'            Query = Query & "        " & RS01!금액 & ", "
'            Query = Query & "        " & RS01!재세탁수량 & ","
'            Query = Query & "        " & RS01!수선수량 & ", "
'            Query = Query & "        " & RS01!반품수량 & ") "
'        Else
'            Query = "UPDATE SuGeumMST "
'            Query = Query & "SET ISu         =   " & RS01!입고수량 & ", "
'            Query = Query & "    CSu         =   " & RS01!출고수량 & ", "
'            Query = Query & "    STag        =   '" & RS01!시작택 & "', "
'            Query = Query & "    ETag        =   '" & RS01!종료택 & "', "
'            Query = Query & "    Amount      =   " & RS01!금액 & ", "
'            Query = Query & "    JSu         =   " & RS01!재세탁수량 & ", "
'            Query = Query & "    SSu         =   " & RS01!수선수량 & ", "
'            Query = Query & "    BSu         =   " & RS01!반품수량 & " "
'            Query = Query & "WHERE   SYear      = '" & Format(dtInput(2).Value, "yyyy") & "' "
'            Query = Query & "AND     SMonth     = '" & Format(dtInput(2).Value, "mm") & "' "
'            Query = Query & "AND     AgencyCode = '" & RS01!대리점코드 & "' "
'        End If
'
'        ADOCon.Execute Query
'
'        RS01.MoveNext
'    Loop
'    DataSave_Master = True
'    Exit Function
'
'SQLERROR:
'    DataSave_Master = False
'    If Err.Number <> 0 Then
'        MsgBox "[" & Err.Number & "] " & Err.Description
'        Exit Function
'    End If
End Function



Private Sub DataSave_ALL()
'    Dim i As Integer
'    Dim sRunCode As String
'
'    On Error GoTo SQLERROR
'
'
'    sRunCode = Mid(Trim(cboInput(3).Text) & Space(10), 2, 4)
'
'    ' 본사일 경우만
'    If sRunCode = MASTER_OFFICE_CODE Then
'        If DataSave_Master = True Then
'            MsgBox "해당되는 데이터가 정상적으로 처리되었습니다.", vbInformation
'        End If
'
'    ' 전체일경우
'    ElseIf sRunCode = "0000" Then
'        For i = 2 To cboInput(3).ListCount
'            sRunCode = Mid(Trim(cboInput(3).List(i)) & Space(10), 2, 4)
'            ' 전체중 본사 이외의 사업장일 경우
'            If Len(sRunCode) = 4 And sRunCode <> MASTER_OFFICE_CODE  Then
'                Call DataSave_MasterCode(sRunCode)
'
'            ' 전체중에 본사일 경우
'            ElseIf Len(sRunCode) = 4 And sRunCode = MASTER_OFFICE_CODE Then
'                Call DataSave_Master
'
'            End If
'
'        Next i
'        MsgBox "해당되는 데이터가 정상적으로 처리되었습니다.", vbInformation
'
'    ' 특정 지사일 경우
'    Else
'        If DataSave_MasterCode(sRunCode) = True Then
'            MsgBox "해당되는 데이터가 정상적으로 처리되었습니다.", vbInformation
'        End If
'    End If
'
'    Exit Sub
'
'SQLERROR:
'    If Err.Number <> 0 Then
'        MsgBox "[" & Err.Number & "] " & Err.Description
'        Exit Sub
'    End If
End Sub

Public Function DataSave_MasterCode(sCode As String) As Boolean
'    Dim i As Integer
'    Dim sRunCode As String
'
'    Dim SSQL As String
'    Dim sDate As String
'    Dim sDate2 As String
'
'    On Error GoTo SQLERROR
'
'    ' 본사는 이쪽에서 작업하지 않는다.
'    DataSave_MasterCode = False
'    If sCode = MASTER_OFFICE_CODE Then Exit Function
'
'    sDate = Format(dtInput(2).Value, "YYYY-MM") & "-01"
'    dtInput(0).Value = Format(sDate, "####-##-##")
'
'    sDate2 = Format(DateAdd("d", -1, DateAdd("m", 1, dtInput(0).Value)), "YYYY-MM-DD")
'    dtInput(1).Value = Format(sDate2, "####-##-##")
'
'    SSQL = "SELECT  A.AgencyCode        AS 대리점코드, "
'    SSQL = SSQL & "      B.AgencyName       AS 대리점명, "
'    SSQL = SSQL & "      B.MasterCode       AS 지사코드, "
'    SSQL = SSQL & "      Sum(A.IpSu)        AS 입고수량, "
'    SSQL = SSQL & "      Sum(A.ChulSu)      AS 출고수량, "
'    SSQL = SSQL & "      Min(CASE WHEN A.StartTag <> '    ' AND NOT A.StartTag IS NULL THEN A.StartTag END)    AS 시작택, "
'    SSQL = SSQL & "      Max(A.EndTag)      AS 종료택, "
'    SSQL = SSQL & "      Sum(A.Amount)      AS 금액, "
'    SSQL = SSQL & "      Sum(A.JaeSu)       AS 재세탁수량, "
'    SSQL = SSQL & "      Sum(A.SuSu)        AS 수선수량, "
'    SSQL = SSQL & "      Sum (A.BanSu)      AS 반품수량 "
'    SSQL = SSQL & "FROM    SugeumTotal      A (NOLOCK), "
'    SSQL = SSQL & "        MasterAgencyCT    B (NOLOCK) "
'    SSQL = SSQL & "WHERE   A.AgencyCode = B.AgencyCode "
'    SSQL = SSQL & "AND     A.MasterCode = '" & sCode & "' "
'    SSQL = SSQL & "AND     B.MasterCode = '" & sCode & "' "
'    SSQL = SSQL & "AND     A.MasterCode = B.MasterCode "
'    SSQL = SSQL & "AND     A.AgencyCode = B.AgencyCode "
'    SSQL = SSQL & "AND     A.SuDate BETWEEN '" & sDate & "' AND '" & sDate2 & "' "
'    SSQL = SSQL & "GROUP BY    A.AgencyCode, "
'    SSQL = SSQL & "            B.AgencyName, "
'    SSQL = SSQL & "            B.MasterCode "
'
'    Set RS01 = New ADODB.Recordset
'    RS01.Open SSQL, ADOCon, adOpenStatic
'
'    mskInput(1).Text = RS01.RecordCount
'
'    Do While Not RS01.EOF
'        txtInput(0).Text = RS01!대리점코드
'        txtInput(1).Text = RS01!대리점명
'
'        i = i + 1
'        mskInput(2).Text = i
'
'        DoEvents
'
'        SSQL = "SELECT  Count(*)    AS 레코드건수 "
'        SSQL = SSQL & "FROM    SuGeumMSTTotal   A (NOLOCK) "
'        SSQL = SSQL & "WHERE   A.SYear      = '" & Format(dtInput(2).Value, "yyyy") & "' "
'        SSQL = SSQL & "AND     A.SMonth     = '" & Format(dtInput(2).Value, "mm") & "' "
'        SSQL = SSQL & "AND     A.AgencyCode = '" & RS01!대리점코드 & "' "
'        SSQL = SSQL & "AND     A.MasterCode = '" & RS01!지사코드 & "' "
'
'        Set RS02 = New ADODB.Recordset
'        RS02.Open SSQL, ADOCon, adOpenStatic
'
'        If RS02!레코드건수 = 0 Then
'            SSQL = "INSERT INTO SuGeumMSTTotal "
'            SSQL = SSQL & "    (SYear, "
'            SSQL = SSQL & "     SMonth, "
'            SSQL = SSQL & "     MasterCode, "
'            SSQL = SSQL & "     AgencyCode, "
'            SSQL = SSQL & "     ISu, "
'            SSQL = SSQL & "     CSu, "
'            SSQL = SSQL & "     STag, "
'            SSQL = SSQL & "     ETag, "
'            SSQL = SSQL & "     Amount, "
'            SSQL = SSQL & "     JSu, "
'            SSQL = SSQL & "     SSu, "
'            SSQL = SSQL & "     BSu) "
'            SSQL = SSQL & "VALUES ('" & Format(dtInput(2).Value, "yyyy") & "', "
'            SSQL = SSQL & "        '" & Format(dtInput(2).Value, "mm") & "', "
'            SSQL = SSQL & "        '" & RS01!지사코드 & "', "
'            SSQL = SSQL & "        '" & RS01!대리점코드 & "', "
'            SSQL = SSQL & "        " & RS01!입고수량 & ", "
'            SSQL = SSQL & "        " & RS01!출고수량 & ", "
'            SSQL = SSQL & "        '" & RS01!시작택 & "', "
'            SSQL = SSQL & "        '" & RS01!종료택 & "', "
'            SSQL = SSQL & "        " & RS01!금액 & ", "
'            SSQL = SSQL & "        " & RS01!재세탁수량 & ","
'            SSQL = SSQL & "        " & RS01!수선수량 & ", "
'            SSQL = SSQL & "        " & RS01!반품수량 & ") "
'        Else
'            SSQL = "UPDATE SuGeumMSTTotal "
'            SSQL = SSQL & "SET ISu         =   " & RS01!입고수량 & ", "
'            SSQL = SSQL & "    CSu         =   " & RS01!출고수량 & ", "
'            SSQL = SSQL & "    STag        =   '" & RS01!시작택 & "', "
'            SSQL = SSQL & "    ETag        =   '" & RS01!종료택 & "', "
'            SSQL = SSQL & "    Amount      =   " & RS01!금액 & ", "
'            SSQL = SSQL & "    JSu         =   " & RS01!재세탁수량 & ", "
'            SSQL = SSQL & "    SSu         =   " & RS01!수선수량 & ", "
'            SSQL = SSQL & "    BSu         =   " & RS01!반품수량 & " "
'            SSQL = SSQL & "WHERE   SYear      = '" & Format(dtInput(2).Value, "yyyy") & "' "
'            SSQL = SSQL & "AND     SMonth     = '" & Format(dtInput(2).Value, "mm") & "' "
'            SSQL = SSQL & "AND     AgencyCode = '" & RS01!대리점코드 & "' "
'            SSQL = SSQL & "AND     MasterCode = '" & RS01!지사코드 & "' "
'        End If
'
'        ADOCon.Execute SSQL
'
'        RS01.MoveNext
'    Loop
'    DataSave_MasterCode = True
'    Exit Function
'
'SQLERROR:
'    DataSave_MasterCode = False
'    If Err.Number <> 0 Then
'        MsgBox "[" & Err.Number & "] " & Err.Description
'        Exit Function
'    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    P_04006_Flag = False
End Sub

