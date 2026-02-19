VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_04009_N1 
   BorderStyle     =   1  '단일 고정
   Caption         =   "[전사업장]점별 월간 매출현황"
   ClientHeight    =   10035
   ClientLeft      =   630
   ClientTop       =   4335
   ClientWidth     =   16080
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_04009_N1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   16080
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   10035
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16080
      _ExtentX        =   28363
      _ExtentY        =   17701
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_04009_N1.frx":058A
      Begin Threed.SSPanel panInput 
         Height          =   615
         Left            =   15
         TabIndex        =   1
         Top             =   540
         Width           =   16050
         _ExtentX        =   28310
         _ExtentY        =   1085
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   8880
            TabIndex        =   2
            Top             =   60
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   556
            _Version        =   262144
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   1500
            TabIndex        =   3
            Top             =   60
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   262144
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   14
            Left            =   3030
            TabIndex        =   4
            Top             =   60
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "사 업 장"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   4440
            TabIndex        =   5
            Top             =   60
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            _Version        =   262144
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   13
            Left            =   120
            TabIndex        =   6
            Top             =   60
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "해당년월"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   15
            Left            =   7440
            TabIndex        =   7
            Top             =   60
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "가 맹 점"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   8
         Top             =   15
         Width           =   8385
         _ExtentX        =   14790
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
         Caption         =   " #"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_04009_N1.frx":061C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   1
         Left            =   8415
         TabIndex        =   9
         Top             =   15
         Width           =   7650
         _ExtentX        =   13494
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
         PictureBackground=   "P_04009_N1.frx":081E
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   10
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
            Picture         =   "P_04009_N1.frx":0A20
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   11
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
            Picture         =   "P_04009_N1.frx":0FBA
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   12
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
            Picture         =   "P_04009_N1.frx":1554
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   13
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
            Picture         =   "P_04009_N1.frx":1AEE
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   14
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
            Picture         =   "P_04009_N1.frx":2088
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   15
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
            Picture         =   "P_04009_N1.frx":2622
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   16
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
            Picture         =   "P_04009_N1.frx":2BBC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   17
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "조회"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_04009_N1.frx":3156
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   8850
         Left            =   15
         TabIndex        =   18
         Top             =   1170
         Width           =   16050
         _Version        =   524288
         _ExtentX        =   28310
         _ExtentY        =   15610
         _StockProps     =   64
         BackColorStyle  =   1
         ColsFrozen      =   2
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   38
         MaxRows         =   501
         SpreadDesigner  =   "P_04009_N1.frx":36F0
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_04009_N1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String



Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: Call Data_Display   ' 조회
        Case 1: 'Call DataAdd        ' 신규
        Case 2: 'Call DataSave       ' 저장
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
    Call Data_Display
End Sub

Private Sub Form_Load()

    'Call Data_Display

End Sub


Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i As Integer
    
    'For i = 0 To 11
    '    txtNum(i).Value = 0
    'Next i
    
    ReDim sValue(3)
    
    sValue(0) = Mid(panCaption(1).Caption, 2, 4)
    sValue(1) = Mid(panCaption(2).Caption, 2, 6)
    
    sValue(2) = panCaption(0).Caption & "-01"
    sValue(3) = DateAdd("d", -1, DateAdd("m", 1, sValue(2)))
    
    If HeadOffice = MASTER_OFFICE_CODE Then
        If DBOpen_Master(sValue(0)) = False Then Exit Sub
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecProMaster("SP_04009_N1_01", sValue(), Err_Num, Err_Dec)
    Else
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_04009_N1_01", sValue(), Err_Num, Err_Dec)
    End If
    
    With spdView
        .MaxRows = 0
        .Redraw = False
        
        Do Until RS01.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            
            .Col = 1:  .Text = RS01!매출일자 & ""               ' 1
            .Col = 2:  .Text = ExecWeekDay(RS01!매출일자) & ""
            .Col = 3:  .Text = RS01!지사금액 & ""                 ' 4
            .Col = 4:  .Text = RS01!가맹점금액 & ""               ' 5
            
            .Col = 5:  .Text = RS01!로열티금액1 & ""                      ' 6
            .Col = 6:  .Text = RS01!로열티금액2 & ""                 ' 7
            .Col = 7:   .Text = RS01!지사차감후금액 & ""                 ' 7
            
            .Col = 8: .Text = RS01!수수료승인금액 & ""                 ' 7
            .Col = 9: .Text = RS01!수수료취소금액 & ""                 ' 7
            .Col = 10: .Text = RS01!수수료지원금액 & ""                 ' 7
            
            .Col = 11:  .Text = RS01!세탁반품환불건수 & ""                 ' 6
            .Col = 12:  .Text = RS01!세탁반품환불지사금액 & ""                 ' 7
            
            .Col = 13:  .Text = RS01!접수수량 & ""                 ' 6
            .Col = 14:  .Text = RS01!출고수량 & ""                 ' 7
            .Col = 15:  .Text = RS01!접수금액 & ""                 ' 8
            .Col = 16:  .Text = RS01!현금입금 + RS01!카드금액 & "" ' 9
            
            If RS01!접수수량 = 0 Then
                .Col = 17: .Text = 0 & ""   '10
                .Col = 18: .Text = 0 & ""   '11
                .Col = 19: .Text = 0 & ""   '12
            Else
                .Col = 17: .Text = RS01!접수금액 / RS01!접수수량 & ""   '10
                .Col = 18: .Text = RS01!지사금액 / RS01!접수수량 & ""   '11
                .Col = 19: .Text = RS01!가맹점금액 / RS01!접수수량 & "" '12
            End If
            
            .Col = 20: .Text = RS01!현금입금 & ""                 '10
            .Col = 21: .Text = RS01!카드금액 & ""                 '11
            .Col = 22: .Text = RS01!카드건수 & ""                 '12
            .Col = 23: .Text = RS01!쿠폰금액 & ""                 '13
            .Col = 24: .Text = RS01!쿠폰건수 & ""                 '14
            .Col = 25: .Text = RS01!발생마일리지 & ""             '15
            .Col = 26: .Text = RS01!사용마일리지 & ""             '16
            .Col = 27: .Text = RS01!삭제마일리지 & ""             '17
            
            .Col = 28: .Text = RS01!미수카드금액 & ""             '19
            .Col = 29: .Text = RS01!미수카드건수 & ""             '19
            .Col = 30: .Text = RS01!카드취소금액 & ""             '19
            .Col = 31: .Text = RS01!카드취소건수 & ""             '19
            
            .Col = 32: .Text = RS01!반품환불금액 & ""             '18
            .Col = 33: .Text = RS01!반품환불건수 & ""             '19
            .Col = 34: .Text = RS01!세탁환불금액 & ""             '20
            .Col = 35: .Text = RS01!세탁환불건수 & ""             '21
            .Col = 36: .Text = RS01!재세탁수량 & ""               '22
            .Col = 37: .Text = RS01!수선금액 & ""                 '23
            .Col = 38: .Text = RS01!수선수량 & ""                 '24
                        
           
            
            RS01.MoveNext
        Loop
        RS01.Close
        Set RS01 = Nothing
        
        If .MaxRows > 0 Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Row = .Row
            .Row2 = .Row
            .Col = 1
            .Col2 = .MaxCols
            .BlockMode = True
            .BackColor = &HC0FFC0
            .BlockMode = False
        
            .Col = 4:  .Text = "합계"
            
            .Col = 3:  .Formula = "SUM(C1:C" & .MaxRows - 1 & ")"
            .Col = 4:  .Formula = "SUM(D1:D" & .MaxRows - 1 & ")"
            .Col = 5:  .Formula = "SUM(E1:E" & .MaxRows - 1 & ")"
            .Col = 6:  .Formula = "SUM(F1:F" & .MaxRows - 1 & ")"
            .Col = 7:  .Formula = "SUM(G1:G" & .MaxRows - 1 & ")"
            .Col = 8:  .Formula = "SUM(H1:H" & .MaxRows - 1 & ")"
            
            .Col = 12: .Formula = "SUM(L1:L" & .MaxRows - 1 & ")"
            .Col = 13: .Formula = "SUM(M1:M" & .MaxRows - 1 & ")"
            .Col = 14: .Formula = "SUM(N1:N" & .MaxRows - 1 & ")"
            .Col = 15: .Formula = "SUM(O1:O" & .MaxRows - 1 & ")"
            .Col = 16: .Formula = "SUM(P1:P" & .MaxRows - 1 & ")"
            
'            .Col = 17: .Formula = "SUM(Q1:Q" & .MaxRows - 1 & ")"
'            .Col = 18: .Formula = "SUM(R1:R" & .MaxRows - 1 & ")"
'            .Col = 19: .Formula = "SUM(S1:S" & .MaxRows - 1 & ")"

            .Col = 20: .Formula = "SUM(T1:T" & .MaxRows - 1 & ")"
            .Col = 21: .Formula = "SUM(U1:U" & .MaxRows - 1 & ")"
            .Col = 22: .Formula = "SUM(V1:V" & .MaxRows - 1 & ")"
            .Col = 23: .Formula = "SUM(W1:W" & .MaxRows - 1 & ")"
            
            .Col = 24: .Formula = "SUM(X1:X" & .MaxRows - 1 & ")"
            .Col = 25: .Formula = "SUM(Y1:Y" & .MaxRows - 1 & ")"
            .Col = 26: .Formula = "SUM(Z1:Z" & .MaxRows - 1 & ")"
            .Col = 27: .Formula = "SUM(AA1:AA" & .MaxRows - 1 & ")"
            .Col = 28: .Formula = "SUM(AB1:AB" & .MaxRows - 1 & ")"
            .Col = 29: .Formula = "SUM(AC1:AC" & .MaxRows - 1 & ")"
            .Col = 30: .Formula = "SUM(AD1:AD" & .MaxRows - 1 & ")"
            .Col = 31: .Formula = "SUM(AE1:AE" & .MaxRows - 1 & ")"
            .Col = 32: .Formula = "SUM(AF1:AF" & .MaxRows - 1 & ")"
            .Col = 33: .Formula = "SUM(AG1:AG" & .MaxRows - 1 & ")"
            .Col = 34: .Formula = "SUM(AH1:AH" & .MaxRows - 1 & ")"
            .Col = 35: .Formula = "SUM(AI1:AI" & .MaxRows - 1 & ")"
            .Col = 36: .Formula = "SUM(AJ1:AJ" & .MaxRows - 1 & ")"
            .Col = 37: .Formula = "SUM(AK1:AK" & .MaxRows - 1 & ")"
            .Col = 38: .Formula = "SUM(AL1:AL" & .MaxRows - 1 & ")"
        
        End If
        
        .Redraw = True
    End With
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If NewRow <> -1 Then
    
           With spdView

                If NewRow <> -1 Then
                    .Row = Row
                    .Col = 2
                    If spdView.Text = "일" Then
                        spdView.Col = -1
                        spdView.BackColor = vbYellow
                    Else
                        If (Row Mod 2) = 0 Then
                            .Col = -1
                            .BackColor = glbGray
                        Else
                            .Col = -1
                            .BackColor = vbWhite
                        End If
                    End If
                    .Row = NewRow
                    .Col = -1
                    .BackColor = glbYellow
                End If
            End With

    End If
End Sub
Private Sub spdView_Change(ByVal Col As Long, ByVal Row As Long)
    Select Case Col
        Case 2
            spdView.Row = Row
            
            'spdView.Col = 14
            
            If spdView.Text = "일" Then
                spdView.Col = -1
                spdView.BackColor = vbYellow
            End If
    End Select
End Sub

Private Sub spdView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        'PopupMenu P_00000.PopUp
    End If
End Sub

Public Sub DataSave()
End Sub
