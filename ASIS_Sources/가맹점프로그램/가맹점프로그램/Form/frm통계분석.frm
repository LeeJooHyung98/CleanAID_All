VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{B6C10482-FB89-11D4-93C9-006008A7EED4}#1.0#0"; "TeeChart5.ocx"
Begin VB.Form frm통계분석 
   Caption         =   "통계 분석"
   ClientHeight    =   10080
   ClientLeft      =   2955
   ClientTop       =   2715
   ClientWidth     =   15240
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10080
   ScaleWidth      =   15240
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   10080
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   17780
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm통계분석.frx":0000
      Begin TeeChart.TChart TChart1 
         Height          =   8415
         Left            =   15
         TabIndex        =   16
         Top             =   1215
         Width           =   15210
         Base64          =   $"frm통계분석.frx":0092
      End
      Begin Threed.SSPanel Panel 
         Height          =   750
         Left            =   15
         TabIndex        =   1
         Top             =   450
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   1323
         _Version        =   262144
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   330
            Index           =   0
            Left            =   915
            TabIndex        =   2
            Top             =   60
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
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
            Format          =   64290819
            CurrentDate     =   37427
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   330
            Index           =   1
            Left            =   2610
            TabIndex        =   3
            Top             =   60
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
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
            Format          =   64290819
            CurrentDate     =   37427
         End
         Begin XtremeSuiteControls.CheckBox chkDay 
            Height          =   300
            Left            =   4125
            TabIndex        =   4
            Top             =   75
            Width           =   1470
            _Version        =   851970
            _ExtentX        =   2593
            _ExtentY        =   529
            _StockProps     =   79
            Caption         =   "기간전체"
            Transparent     =   -1  'True
            Appearance      =   6
            MultiLine       =   0   'False
         End
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   630
            Left            =   6255
            TabIndex        =   5
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 조회(&F)"
            Appearance      =   6
            Picture         =   "frm통계분석.frx":07CC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   13665
            TabIndex        =   6
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frm통계분석.frx":0EC6
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   4
            Left            =   11205
            TabIndex        =   7
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 출력(&P)"
            Appearance      =   6
            Picture         =   "frm통계분석.frx":1F58
         End
         Begin Threed.SSOption optGubun 
            Height          =   255
            Index           =   0
            Left            =   915
            TabIndex        =   8
            Top             =   450
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   450
            _Version        =   262144
            Caption         =   "일자별"
            Value           =   -1
         End
         Begin Threed.SSOption optGubun 
            Height          =   255
            Index           =   1
            Left            =   2115
            TabIndex        =   9
            Top             =   450
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   450
            _Version        =   262144
            Caption         =   "월별"
         End
         Begin Threed.SSOption optGubun 
            Height          =   255
            Index           =   2
            Left            =   3150
            TabIndex        =   10
            Top             =   450
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   450
            _Version        =   262144
            Caption         =   "품목별"
         End
         Begin Threed.SSOption optGubun 
            Height          =   255
            Index           =   3
            Left            =   4275
            TabIndex        =   11
            Top             =   450
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   262144
            Caption         =   "미수금액별"
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "접수일자:"
            Height          =   165
            Index           =   21
            Left            =   45
            TabIndex        =   14
            Top             =   120
            Width           =   840
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "분석조건:"
            Height          =   165
            Index           =   0
            Left            =   45
            TabIndex        =   13
            Top             =   480
            Width           =   840
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H000000C0&
            Height          =   165
            Index           =   1
            Left            =   2370
            TabIndex        =   12
            Top             =   135
            Width           =   210
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   15
         Top             =   15
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   741
         _Version        =   262144
         Font3D          =   3
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
         Caption         =   "      통계 분석"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm통계분석.frx":2652
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm통계분석.frx":2878
            Top             =   -15
            Width           =   765
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   420
         Left            =   15
         TabIndex        =   17
         Top             =   9645
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   741
         _Version        =   262144
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin TeeChart.ChartPageNavigator ChartPageNavigator1 
            Height          =   345
            Left            =   30
            Negotiate       =   -1  'True
            OleObjectBlob   =   "frm통계분석.frx":3442
            TabIndex        =   18
            Top             =   45
            Width           =   1200
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   360
            Index           =   0
            Left            =   11040
            TabIndex        =   19
            Top             =   30
            Width           =   1065
            _Version        =   262145
            _ExtentX        =   1879
            _ExtentY        =   635
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   192
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            ReadOnly        =   -1  'True
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   3
            StartText.y     =   4
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   16
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   360
            Index           =   1
            Left            =   13575
            TabIndex        =   20
            Top             =   30
            Width           =   1635
            _Version        =   262145
            _ExtentX        =   2884
            _ExtentY        =   635
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   192
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            ReadOnly        =   -1  'True
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   3
            StartText.y     =   4
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   16
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin VB.Label Label3 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "합계수량 :"
            Height          =   210
            Index           =   0
            Left            =   10020
            TabIndex        =   22
            Top             =   120
            Width           =   960
         End
         Begin VB.Label Label3 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "합계금액 :"
            Height          =   210
            Index           =   1
            Left            =   12570
            TabIndex        =   21
            Top             =   120
            Width           =   960
         End
      End
   End
End
Attribute VB_Name = "frm통계분석"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn

    Select Case Index
        Case 4: TChart1.Printer.ShowPreview
        Case 5: Unload Me
    End Select
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Public Sub cmdList_Click()
    On Error GoTo ErrRtn
            
    txtNum(0).Value = 0
    txtNum(1).Value = 0
            
    Select Case True
        Case optGubun(0).Value
            Query = "SELECT   RIGHT(접수일자,5) AS 접수일자, "
            Query = Query & " SUM(금액) AS 접수금액, "
            Query = Query & " COUNT(*)  AS 접수량 "
            Query = Query & " FROM TB_입출고 "
            Query = Query & " WHERE (판매취소 <> 'Y')"
            'Query = Query & " WHERE (판매취소 = '' OR 판매취소 IS NULL)"
        
            If chkDay.Value = xtpUnchecked Then
                Query = Query & "  AND (접수일자 >= '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "' "
                Query = Query & "  AND  접수일자 <= '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "') "
            End If
            Query = Query & " GROUP BY 접수일자 "
            Query = Query & " ORDER BY 접수일자 ASC"
            Set ADORs = New ADODB.Recordset
            ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            With TChart1
                .Series(0).Clear
                .Series(1).Clear
                
                Do Until ADORs.EOF
                    .Series(0).Add ADORs!접수금액, ADORs!접수일자, vbRed
                    .Series(1).Add ADORs!접수량, ADORs!접수일자, vbBlue
                    
                    txtNum(0).Value = txtNum(0).Value + CCur(ADORs!접수량)
                    txtNum(1).Value = txtNum(1).Value + CCur(ADORs!접수금액)
                    
                    ADORs.MoveNext
                Loop
            End With
            ADORs.Close
            Set ADORs = Nothing
            
        Case optGubun(1).Value
            Query = "SELECT   LEFT(접수일자,7) AS 접수년월, "
            Query = Query & " SUM(금액) AS 접수금액, "
            Query = Query & " COUNT(*)  AS 접수량 "
            Query = Query & " FROM TB_입출고 "
            Query = Query & " WHERE (판매취소 <> 'Y')"
            'Query = Query & " WHERE (판매취소 = '' OR 판매취소 IS NULL)"
        
            If chkDay.Value = xtpUnchecked Then
                Query = Query & "  AND (접수일자 >= '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "' "
                Query = Query & "  AND  접수일자 <= '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "') "
            End If
            Query = Query & " GROUP BY LEFT(접수일자,7) "
            Query = Query & " ORDER BY LEFT(접수일자,7) ASC "
            Set ADORs = New ADODB.Recordset
            ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            With TChart1
                .Series(0).Clear
                .Series(1).Clear
                
                Do Until ADORs.EOF
                    .Series(0).Add ADORs!접수금액, ADORs!접수년월, vbRed
                    .Series(1).Add ADORs!접수량, ADORs!접수년월, vbBlue
                    
                    txtNum(0).Value = txtNum(0).Value + CCur(ADORs!접수량)
                    txtNum(1).Value = txtNum(1).Value + CCur(ADORs!접수금액)
                    
                    ADORs.MoveNext
                Loop
            End With
            ADORs.Close
            Set ADORs = Nothing
            
        Case optGubun(2).Value
            Query = "SELECT    의류명"
            Query = Query & ", SUM(금액) AS 접수금액"
            Query = Query & ", COUNT(*)  AS 접수량"
            Query = Query & " FROM TB_입출고"
            Query = Query & " WHERE (판매취소 <> 'Y')"
            'Query = Query & " WHERE (판매취소 = '' OR 판매취소 IS NULL)"
        
            If chkDay.Value = xtpUnchecked Then
                Query = Query & "  AND (접수일자 >= '" & Format(dtpDay(0).Value, "YYYY-MM-DD") & "' "
                Query = Query & "  AND  접수일자 <= '" & Format(dtpDay(1).Value, "YYYY-MM-DD") & "') "
            End If
            
            Query = Query & " GROUP BY 의류명 "
            Query = Query & " ORDER BY COUNT(*) DESC, SUM(금액) DESC, 의류명 ASC "
            Set ADORs = New ADODB.Recordset
            ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            With TChart1
                .Series(0).Clear
                .Series(1).Clear
                
                Do Until ADORs.EOF
                    .Series(0).Add ADORs!접수금액, ADORs!의류명, vbRed
                    .Series(1).Add ADORs!접수량, ADORs!의류명, vbBlue
                    
                    txtNum(0).Value = txtNum(0).Value + CCur(ADORs!접수량)
                    txtNum(1).Value = txtNum(1).Value + CCur(ADORs!접수금액)
                    
                    ADORs.MoveNext
                Loop
            End With
            ADORs.Close
            Set ADORs = Nothing
    
        Case optGubun(3).Value
            Query = "SELECT    A.고객코드"
            Query = Query & ", B.성명"
            Query = Query & ", SUM(A.접수금액) AS 총금액의합계"
            Query = Query & ", SUM(A.입금합계)   AS 수금액의합계"
            Query = Query & ", SUM(A.접수금액) - Sum(A.입금합계) - Sum(A.세트할인) AS 잔액의합계"
            Query = Query & " FROM TB_매출 AS A LEFT JOIN TB_고객정보 AS B ON A.고객코드 = B.고객코드"
            Query = Query & " WHERE A.반품수량 = 0"
            Query = Query & " GROUP BY A.고객코드, B.성명"
            Query = Query & " ORDER BY Sum(A.접수금액) - Sum(A.입금합계) - Sum(A.세트할인) DESC"
            Set ADORs = New ADODB.Recordset
            ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
            
            With TChart1
                .Series(0).Clear
                .Series(1).Clear
                
                Do Until ADORs.EOF
                    .Series(0).Add ADORs!잔액의합계, ADORs!성명, vbRed
                    
                    ADORs.MoveNext
                Loop
            End With
            ADORs.Close
            Set ADORs = Nothing
    End Select
    
    ChartPageNavigator1.ChartLink = TChart1.ChartLink
    ChartPageNavigator1.EnableButtons
    
    txtNum(0).Value = Format(txtNum(0).Value, "#,##0")
    txtNum(1).Value = Format(txtNum(1).Value, "#,##0")
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub dtpDay_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    dtpDay(0).Value = Format(Date, "YYYY-MM-DD")
    dtpDay(1).Value = Format(Date, "YYYY-MM-DD")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    pnlHeader.Width = Me.ScaleWidth
    cmdBtn(5).Left = (Me.Width - 200) - cmdBtn(5).Width
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 37
            KeyCode = 0
            SendKeys "+{TAB}" ' Shift + Tab
        Case 39
            KeyCode = 0
            SendKeys "{TAB}"
    End Select
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdList_Click
    End If
End Sub
