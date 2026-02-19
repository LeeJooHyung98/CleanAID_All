VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm일일판매집계 
   Caption         =   "일일판매 집계"
   ClientHeight    =   11970
   ClientLeft      =   2565
   ClientTop       =   2880
   ClientWidth     =   16410
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form20"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11970
   ScaleWidth      =   16410
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   11970
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16410
      _ExtentX        =   28945
      _ExtentY        =   21114
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm일일판매집계.frx":0000
      Begin Threed.SSPanel SSPanel1 
         Height          =   750
         Index           =   0
         Left            =   15
         TabIndex        =   1
         Top             =   450
         Width           =   16380
         _ExtentX        =   28893
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   630
            Left            =   7935
            TabIndex        =   3
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 조회(&F)"
            Appearance      =   6
            Picture         =   "frm일일판매집계.frx":0072
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   1
            Left            =   9870
            TabIndex        =   4
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 엑셀(&E)"
            Appearance      =   6
            Picture         =   "frm일일판매집계.frx":076C
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   13170
            TabIndex        =   5
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frm일일판매집계.frx":0EE6
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   3
            Left            =   11430
            TabIndex        =   6
            Top             =   60
            Visible         =   0   'False
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 출력(&P)"
            Appearance      =   6
            Picture         =   "frm일일판매집계.frx":1F78
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   0
            Left            =   915
            TabIndex        =   7
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
            Format          =   56295427
            CurrentDate     =   40279
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   1
            Left            =   2640
            TabIndex        =   9
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
            Format          =   56295427
            CurrentDate     =   40279
         End
         Begin VB.Label Label1 
            Caption         =   $"frm일일판매집계.frx":2672
            ForeColor       =   &H000000FF&
            Height          =   435
            Left            =   30
            TabIndex        =   12
            Top             =   390
            Width           =   8745
         End
         Begin VB.Label Label 
            Alignment       =   2  '가운데 맞춤
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
            Height          =   210
            Left            =   2415
            TabIndex        =   10
            Top             =   120
            Width           =   180
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "마감일자:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   45
            TabIndex        =   8
            Top             =   120
            Width           =   840
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   16380
         _ExtentX        =   28893
         _ExtentY        =   741
         _Version        =   262144
         Font3D          =   1
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "      일일판매 집계"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm일일판매집계.frx":26E8
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm일일판매집계.frx":290E
            Top             =   -15
            Width           =   765
         End
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   10740
         Left            =   15
         TabIndex        =   11
         Top             =   1215
         Width           =   16380
         _Version        =   524288
         _ExtentX        =   28893
         _ExtentY        =   18944
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
         MaxCols         =   35
         MaxRows         =   200
         OperationMode   =   1
         Protect         =   0   'False
         SpreadDesigner  =   "frm일일판매집계.frx":34D8
         UserResize      =   1
         VisibleCols     =   11
         VisibleRows     =   50
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "frm일일판매집계"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strStart As String
Dim strEnd   As String

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 1: Call Export_Excel(frmMain.cdgExcel, sprGrid)
        
        Case 3
''                If sprGrid.MaxRows = 0 Then Exit Sub
''
''                If Dir(AppPath & "XML", vbDirectory) = "" Then
''                    MkDir AppPath & "XML"
''                End If
''
''                Open AppPath & "XML\일일매출집계.XML" For Output As #1
''
''                Print #1, "<?xml version=""1.0"" encoding=""EUC-KR""?>"
''                Print #1, "<root>"
''
''                      XML = "    <조건>"
''                XML = XML & "        <검색조건>일자 : " & Format(dtpDay.Value, "YYYY-MM-DD") & " " & pnlWeek.Caption & "</검색조건>"
''                XML = XML & "        <가맹점>" & Func_Replace(가맹점정보.가맹점명) & " 일일매출현황</가맹점>"
''                XML = XML & "   </조건>"
''                Print #1, XML
''
''                With sprGrid
''                    For i = 1 To .MaxRows
''                        .Row = i
''
''                                         XML = "    <Data>"
''                        .Col = 1:  XML = XML & "        <택번호>" & .Text & "</택번호>"
''                        .Col = 2:  XML = XML & "        <성명>" & Func_Replace(.Text) & "</성명>"
''                        .Col = 3:  XML = XML & "        <휴대전화>" & Func_Replace(.Text) & "</휴대전화>"
''                        .Col = 4:  XML = XML & "        <전화번호>" & Func_Replace(.Text) & "</전화번호>"
''                        .Col = 5:  XML = XML & "        <품명>" & Func_Replace(.Text) & "</품명>"
''                        .Col = 6:  XML = XML & "        <색상>" & .Text & "</색상>"
''                        .Col = 7:  XML = XML & "        <무늬>" & .Text & "</무늬>"
''                        .Col = 8:  XML = XML & "        <내용>" & .Text & "</내용>"
''                        .Col = 9:  XML = XML & "        <금액>" & .Text & "</금액>"
''                        .Col = 10: XML = XML & "        <상표>" & .Text & "</상표>"
''                        .Col = 11: XML = XML & "        <상태>" & .Text & "</상태>"
''                                   XML = XML & "   </Data>"
''                                   Print #1, XML
''                    Next i
''                End With
''
''
''
''                Print #1, "</root>"
''                Close #1
''
''                With rpt일일매출집계
''                    .dc.FileURL = AppPath & "XML\일일매출집계.XML"
''                    .PrintReport False
''                End With
''
''                Unload rpt일일매출집계
                
        Case 5: Unload Me
    End Select
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn
    
    Query = "SELECT * FROM TB_일일마감"
    Query = Query & " WHERE (마감일자 >= '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "'"
    Query = Query & "   AND  마감일자 <= '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "')"
    Query = Query & " ORDER BY 마감일자 ASC"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows

            .Col = 1:  .Text = ADORs!마감일자 & ""
            .Col = 2:  .Text = ADORs!접수수량 & ""
            .Col = 3:  .Text = ADORs!접수금액 & ""
            
            .Col = 4:  .Text = ADORs!지사금액 & ""
            .Col = 5:  .Text = ADORs!가맹점금액 & ""
            
            .Col = 6:  .Text = ADORs!로열티금액2 & ""
            
            .Col = 7:  .Text = Val(ADORs!카드금액 & "") + Val(ADORs!미수카드금액 & "")
            .Col = 8:  .Text = ADORs!수수료승인금액 & ""
            .Col = 9:  .Text = ADORs!카드취소금액 & ""
            .Col = 10:  .Text = ADORs!수수료취소금액 & ""
            
            
            .Col = 11:  .Text = Val(ADORs!전산사용료 & "")
            
            .Col = 12:  .Text = Val(ADORs!세탁환불지사금액 & "") + Val(ADORs!반품환불지사금액 & "")
            
            ' 지사정산금액 = 지사분매출 - (카드수수료지원금+환불금액) + (카드수수료환불금+로열티2)
            .Col = 13:  .Text = Val(ADORs!지사금액 & "") - (Val(ADORs!수수료승인금액 & "") + Val(ADORs!세탁환불지사금액 & "") + Val(ADORs!반품환불지사금액 & "")) _
                            + (Val(ADORs!수수료취소금액 & "") + Val(ADORs!로열티금액2 & "")) + Val(ADORs!전산사용료 & "")
            .Col = 14:  .Text = Val(ADORs!접수금액 & "") - (Val(ADORs!지사금액 & "") - (Val(ADORs!수수료승인금액 & "") + Val(ADORs!세탁환불지사금액 & "") + Val(ADORs!반품환불지사금액 & "")) _
                            + (Val(ADORs!수수료취소금액 & "") + Val(ADORs!로열티금액2 & ""))) - Val(ADORs!전산사용료 & "") - Val(ADORs!사용마일리지 & "")
            
            
            .Col = 15:  .Text = ADORs!현금입금 + ADORs!카드금액 & ""
            .Col = 16:  .Text = ADORs!현금입금 & ""
            .Col = 17:  .Text = ADORs!카드건수 & ""
            .Col = 18:  .Text = ADORs!카드금액 & ""
            
            .Col = 19:  .Text = ADORs!미수현금수금금액 & ""
            .Col = 20:  .Text = ADORs!미수카드건수 & ""
            .Col = 21:  .Text = ADORs!미수카드금액 & ""
            
            
            .Col = 22: .Text = ADORs!출고수량 & ""
            .Col = 23: .Text = ADORs!반품수량 & ""
            .Col = 24: .Text = ADORs!재세탁수량 & ""
            
            If Len(ADORs!시작택번호) = 9 Then
                .Col = 25: .Text = Format(ADORs!시작택번호, "000-00-0000") & ""
            Else
                .Col = 25: .Text = ADORs!시작택번호 & ""
            End If
            
            If Len(ADORs!종료택번호) = 9 Then
                .Col = 26: .Text = Format(ADORs!종료택번호, "000-00-0000") & ""
            Else
                .Col = 26: .Text = ADORs!종료택번호 & ""
            End If
            
            .Col = 27: .Text = ADORs!쿠폰건수 & ""
            .Col = 28: .Text = ADORs!쿠폰금액 & ""
            
            .Col = 29: .Text = ADORs!발생마일리지 & ""
            .Col = 30: .Text = ADORs!사용마일리지 & ""
            .Col = 31: .Text = ADORs!삭제마일리지 & ""
            
            .Col = 32: .Text = ADORs!반품환불건수 & ""
            .Col = 33: .Text = ADORs!반품환불금액 & ""
            
            .Col = 34: .Text = ADORs!세탁환불건수 & ""
            .Col = 35: .Text = ADORs!세탁환불금액 & ""
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        If .MaxRows > 0 Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Row = .Row: .Row2 = .Row
            .Col = 1:    .Col2 = .MaxCols
            .BlockMode = True
            .BackColor = &HC0E0FF
            .ForeColor = vbRed
            .BlockMode = False
            
            .Col = 1:  .Text = "합계"
            .Col = 2:  .Formula = "SUM(B1:B" & .MaxRows - 1 & ")"
            .Col = 3:  .Formula = "SUM(C1:C" & .MaxRows - 1 & ")"
            .Col = 4:  .Formula = "SUM(D1:D" & .MaxRows - 1 & ")"
            .Col = 5:  .Formula = "SUM(E1:E" & .MaxRows - 1 & ")"
            .Col = 6:  .Formula = "SUM(F1:F" & .MaxRows - 1 & ")"
            .Col = 7:  .Formula = "SUM(G1:G" & .MaxRows - 1 & ")"
            .Col = 8:  .Formula = "SUM(H1:H" & .MaxRows - 1 & ")"
            .Col = 9:  .Formula = "SUM(I1:I" & .MaxRows - 1 & ")"
            .Col = 10: .Formula = "SUM(J1:J" & .MaxRows - 1 & ")"
            
            .Col = 11: .Formula = "SUM(K1:K" & .MaxRows - 1 & ")"
            .Col = 12: .Formula = "SUM(L1:L" & .MaxRows - 1 & ")"
            
            .Col = 13: .Formula = "SUM(M1:M" & .MaxRows - 1 & ")"
            .Col = 14: .Formula = "SUM(N1:N" & .MaxRows - 1 & ")"
            .Col = 15: .Formula = "SUM(O1:O" & .MaxRows - 1 & ")"
            .Col = 16: .Formula = "SUM(P1:P" & .MaxRows - 1 & ")"
            .Col = 17: .Formula = "SUM(Q1:Q" & .MaxRows - 1 & ")"
            .Col = 18: .Formula = "SUM(R1:R" & .MaxRows - 1 & ")"
            .Col = 19: .Formula = "SUM(S1:S" & .MaxRows - 1 & ")"
            .Col = 20: .Formula = "SUM(T1:T" & .MaxRows - 1 & ")"
            .Col = 21: .Formula = "SUM(U1:U" & .MaxRows - 1 & ")"
            .Col = 22: .Formula = "SUM(V1:V" & .MaxRows - 1 & ")"
            .Col = 23: .Formula = "SUM(W1:W" & .MaxRows - 1 & ")"
'            .Col = 24: .Formula = "SUM(X1:X" & .MaxRows - 1 & ")"
'            .Col = 25: .Formula = "SUM(Y1:Y" & .MaxRows - 1 & ")"
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
        End If
        
        .ReDraw = True
    End With
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Sub

Private Sub cmdList_Click()
    Call Data_Display
End Sub

Private Sub cmdPrint_Click()

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    With sprGrid
        .MaxRows = 0
        .RowHeight(-1) = 14
        
        .ColsFrozen = 1
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeExtended
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
    
    dtpDay(0).Value = Format(Date, "YYYY-MM-DD")
    dtpDay(1).Value = Format(Date, "YYYY-MM-DD")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    cmdBtn(5).Left = Me.Width - cmdBtn(5).Width - 200
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdList_Click
    End If
End Sub
