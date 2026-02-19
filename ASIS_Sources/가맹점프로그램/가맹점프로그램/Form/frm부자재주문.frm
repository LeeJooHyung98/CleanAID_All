VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm부자재주문 
   Caption         =   "부자재 주문"
   ClientHeight    =   11730
   ClientLeft      =   2745
   ClientTop       =   2430
   ClientWidth     =   14820
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11730
   ScaleWidth      =   14820
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   11730
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   14820
      _ExtentX        =   26141
      _ExtentY        =   20690
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm부자재주문.frx":0000
      Begin Threed.SSPanel SSPanel 
         Height          =   540
         Index           =   1
         Left            =   15
         TabIndex        =   23
         Top             =   2685
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   953
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   45
            TabIndex        =   4
            Top             =   45
            Width           =   1275
            _Version        =   851970
            _ExtentX        =   2249
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 신규(&N)"
            Appearance      =   6
            Picture         =   "frm부자재주문.frx":00B2
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   1395
            TabIndex        =   3
            Top             =   45
            Width           =   1275
            _Version        =   851970
            _ExtentX        =   2249
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 저장(&S)"
            Appearance      =   6
            Picture         =   "frm부자재주문.frx":0AC4
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   2745
            TabIndex        =   5
            Top             =   45
            Width           =   1275
            _Version        =   851970
            _ExtentX        =   2249
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 삭제(&D)"
            Appearance      =   6
            Picture         =   "frm부자재주문.frx":14D6
         End
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   1455
         Index           =   0
         Left            =   15
         TabIndex        =   8
         Top             =   1215
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   2566
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSPanel pnlDay 
            Height          =   315
            Index           =   0
            Left            =   7125
            TabIndex        =   27
            Top             =   405
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin VB.TextBox txtData 
            Height          =   315
            Index           =   0
            Left            =   930
            ScrollBars      =   2  '수직
            TabIndex        =   2
            Top             =   1095
            Width           =   7365
         End
         Begin VB.ComboBox cboGoods 
            Height          =   315
            Left            =   930
            Style           =   2  '드롭다운 목록
            TabIndex        =   0
            Top             =   405
            Width           =   2850
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   315
            Index           =   0
            Left            =   930
            TabIndex        =   1
            Top             =   750
            Width           =   1155
            _Version        =   262145
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Insert          =   0   'False
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   3
            StartText.y     =   2
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
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin Threed.SSPanel pnlDay 
            Height          =   315
            Index           =   1
            Left            =   7125
            TabIndex        =   28
            Top             =   750
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   262144
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlOrderDate 
            Height          =   315
            Left            =   930
            TabIndex        =   29
            Top             =   45
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   262144
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
            Caption         =   "2010-02-02"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlNo 
            Height          =   315
            Left            =   7125
            TabIndex        =   30
            Top             =   45
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "0"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "주문번호:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   11
            Left            =   6240
            TabIndex        =   26
            Top             =   135
            Width           =   840
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "확정일자:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   10
            Left            =   6255
            TabIndex        =   25
            Top             =   810
            Width           =   825
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "출고일자:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   9
            Left            =   6255
            TabIndex        =   24
            Top             =   465
            Width           =   825
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "주문수량:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   7
            Left            =   45
            TabIndex        =   21
            Top             =   810
            Width           =   825
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "비고:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   45
            TabIndex        =   20
            Top             =   1170
            Width           =   840
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "부자재:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   45
            TabIndex        =   19
            Top             =   465
            Width           =   840
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "주문일자:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   45
            TabIndex        =   18
            Top             =   105
            Width           =   840
         End
      End
      Begin Threed.SSPanel panInput 
         Height          =   750
         Left            =   15
         TabIndex        =   7
         Top             =   450
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   0
            Left            =   915
            TabIndex        =   9
            Top             =   45
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   16646147
            CurrentDate     =   36686
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   1
            Left            =   2640
            TabIndex        =   10
            Top             =   45
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   16646147
            CurrentDate     =   36686
         End
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   630
            Left            =   7500
            TabIndex        =   13
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 조회(&F)"
            Appearance      =   6
            Picture         =   "frm부자재주문.frx":1EE8
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   3
            Left            =   10035
            TabIndex        =   14
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 엑셀(&E)"
            Appearance      =   6
            Picture         =   "frm부자재주문.frx":25E2
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   13125
            TabIndex        =   15
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frm부자재주문.frx":2D5C
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   4
            Left            =   11580
            TabIndex        =   16
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 출력(&P)"
            Appearance      =   6
            Picture         =   "frm부자재주문.frx":3DEE
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "주문일자:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   45
            TabIndex        =   17
            Top             =   105
            Width           =   840
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
            Left            =   2355
            TabIndex        =   11
            Top             =   105
            Width           =   270
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   12
         Top             =   15
         Width           =   14790
         _ExtentX        =   26088
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
         Caption         =   "      부자재 주문"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm부자재주문.frx":44E8
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm부자재주문.frx":470E
            Top             =   -15
            Width           =   765
         End
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   8475
         Left            =   15
         TabIndex        =   22
         Top             =   3240
         Width           =   14790
         _Version        =   524288
         _ExtentX        =   26088
         _ExtentY        =   14949
         _StockProps     =   64
         BackColorStyle  =   1
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
         MaxCols         =   9
         MaxRows         =   35
         ScrollBars      =   2
         SpreadDesigner  =   "frm부자재주문.frx":52D8
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "frm부자재주문"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sValue() As String

Dim Err_Num As Long
Dim Err_Desc As String

Private Sub Text_Clear()
    On Error GoTo ErrRtn
    
    pnlNo.Caption = "0"    '
    pnlOrderDate.Caption = Format(Date, "YYYY-MM-DD")
    
    txtNum(0).Value = 0    '
    
    txtData(0).Text = ""   '
    
    pnlDay(0).Caption = "" '
    pnlDay(1).Caption = "" '
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub Data_Display()
    If Server_Connection(HostCon, "LAUNDRY1000") = False Then Exit Sub
    
    ReDim sValue(3)
    
    sValue(0) = 가맹점정보.가맹점코드                 ' 가맹점코드
    sValue(1) = Format(dtpDay(0).Value, "YYYY-MM-DD") ' 주문일자1
    sValue(2) = Format(dtpDay(1).Value, "YYYY-MM-DD") ' 주문일자2
    sValue(3) = 0                                     ' 주문코드
    
    Set ADORs = New ADODB.Recordset
    Set ADORs = SP_Exec(HostCon, "SP_R_부자재주문", sValue(), Err_Num, Err_Desc)
    
    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        
            .Col = 1: .Text = ADORs!주문코드 & "" ' 1
            .Col = 2: .Text = ADORs!주문일자 & "" ' 2
            .Col = 3: .Text = ADORs!부자재명 & "" ' 3
            .Col = 4: .Text = ADORs!수량 & ""     ' 4
            .Col = 5: .Text = ADORs!단가 & ""     ' 5
            .Col = 6: .Text = ADORs!공급가액 & "" ' 6
            .Col = 7: .Text = ADORs!비고 & ""     ' 7
            .Col = 8: .Text = ADORs!출고일자 & "" ' 8
            .Col = 9: .Text = ADORs!확정일자 & "" ' 9
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
        
        .ReDraw = True
    End With
End Sub

Private Sub cboGoods_Click()
''    If Server_Connection(HostCon,"LAUNDRY1000") = False Then Exit Sub
''
''    ReDim sValue(0)
''
''    sValue(0) = cboGoods.ItemData(cboGoods.ListIndex)
''
''    Set SUBRs = New ADODB.Recordset
''    Set SUBRs = SP_Exec(hostcon,"SP_R_부자재", sValue(), Err_Num, Err_Desc)
''
''    If SUBRs.EOF Then
''        txtData(0).Text = ""
''        txtNum(1).Value = 0
''    Else
''        txtData(0).Text = SUBRs!규격 & ""
''        txtNum(1).Value = SUBRs!판매단가 & ""
''    End If
''    SUBRs.Close
''    Set SUBRs = Nothing
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0:
            Call Text_Clear
            
            cboGoods.SetFocus
            
        Case 1:
            If pnlDay(0).Caption <> "" Then
                MsgBox "출고된 부자재 주문내역은 수정할 수 없습니다.", vbInformation, "확인"
            Else
                If Server_Connection(HostCon, "LAUNDRY1000") = False Then Exit Sub
                
                ReDim sValue(11)
                
                sValue(0) = pnlNo.Caption                              ' 주문번호
                sValue(1) = Format(pnlOrderDate.Caption, "YYYY-MM-DD") ' 주문일자
                sValue(2) = 가맹점정보.가맹점코드                      ' 가맹점코드
                sValue(3) = cboGoods.ItemData(cboGoods.ListIndex)      ' 부자재코드
                sValue(4) = cboGoods.Text & ""                         ' 부자재명
                sValue(5) = ""                                         ' -
                sValue(6) = txtNum(0).Value                            ' 수량
                sValue(7) = 0                                          ' 단가
                sValue(8) = 0                                          ' 공급가액
                sValue(9) = 0                                          ' 세액
                sValue(10) = 0                                         ' 합계금액
                sValue(11) = txtData(0).Text & ""                      ' 비고
                
                Call SP_Exec(HostCon, "SP_CU_부자재주문", sValue(), Err_Num, Err_Desc)
            
                Call Text_Clear
                Call Data_Display
            End If
            
        Case 2:
            If pnlDay(0).Caption <> "" Then
                MsgBox "출고된 부자재 주문내역은 삭제할 수 없습니다.", vbInformation, "확인"
            Else
                If Server_Connection(HostCon, "LAUNDRY1000") = False Then Exit Sub
                
                Rtn = MsgBox("삭제하시겠습니까?", vbQuestion + vbYesNo + vbDefaultButton2, "삭제")
                
                If Rtn = vbYes Then
                    Query = "DELETE FROM TB_부자재주문"
                    Query = Query & " WHERE 주문코드   = " & pnlNo.Caption
                    Query = Query & "   AND 가맹점코드 = '" & 가맹점정보.가맹점코드 & "'"
                    HostCon.Execute Query
                End If
                
                Call Text_Clear
                Call Data_Display
            End If
            
        Case 3: Call Export_Excel(frmMain.cdgExcel, sprGrid)
        Case 4:
            Rtn = MsgBox("출력 미리보기를 하시겠습니까?", vbQuestion + vbYesNo, "출력")
            
            If Rtn = vbYes Then
                Call Data_Print(True)
            Else
                Call Data_Print(False)
            End If
            
        Case 5: Unload Me           ' 종료
    End Select
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub Data_Print(Print_PreView As Boolean)
    On Error GoTo ErrRtn
    
    If sprGrid.MaxRows = 0 Then Exit Sub

    If Dir(AppPath & "XML", vbDirectory) = "" Then
        MkDir AppPath & "XML"
    End If

    Open AppPath & "XML\부자재주문.XML" For Output As #1
    
    Print #1, "<?xml version=""1.0"" encoding=""EUC-KR""?>"
    Print #1, "<root>"
        
          XML = "    <조건>"
    XML = XML & "        <주문일자>주문일자 : " & Format(dtpDay(0).Value, "YYYY-MM-DD") & " ~ " & Format(dtpDay(1).Value, "YYYY-MM-DD") & "</주문일자>"
    XML = XML & "        <가맹점>크린에이드 - " & Func_Replace(가맹점정보.가맹점명) & "</가맹점>"
    XML = XML & "   </조건>"
    Print #1, XML
    
    With sprGrid
        For i = 1 To .MaxRows
            .Row = i
            
                            XML = "    <Data>"
            .Col = 1: XML = XML & "        <주문번호>" & .Text & "</주문번호>"
            .Col = 2: XML = XML & "        <주문일자>" & Func_Replace(.Text) & "</주문일자>"
            .Col = 3: XML = XML & "        <부자재명>" & Func_Replace(.Text) & "</부자재명>"
            .Col = 4: XML = XML & "        <수량>" & Func_Replace(.Text) & "</수량>"
            .Col = 5: XML = XML & "        <단가>" & Func_Replace(.Text) & "</단가>"
            .Col = 6: XML = XML & "        <공급가액>" & Func_Replace(.Text) & "</공급가액>"
            .Col = 7: XML = XML & "        <비고>" & Func_Replace(.Text) & "</비고>"
            .Col = 8: XML = XML & "        <출고일자>" & Func_Replace(.Text) & "</출고일자>"
            .Col = 9: XML = XML & "        <확정일자>" & Func_Replace(.Text) & "</확정일자>"
                      XML = XML & "   </Data>"
                      Print #1, XML
        Next i
        
        Print #1, "</root>"
        Close #1
    End With
    
    If Print_PreView = True Then
        With rpt부자재주문
            .dc.FileURL = AppPath & "XML\부자재주문.XML"
            .Show 1
        End With
    Else
        With rpt부자재주문
            .dc.FileURL = AppPath & "XML\부자재주문.XML"
            .PrintReport False
        End With
    
        Unload rpt부자재주문
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub cmdList_Click()
    Call Data_Display
End Sub

Private Sub Form_Load()
    With sprGrid
        .MaxRows = 0
        .RowHeight(-1) = 14
        
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle
        
        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeSingle
        
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
    
    dtpDay(0).Value = Format(Date, "YYYY-MM-DD")
    dtpDay(1).Value = Format(Date, "YYYY-MM-DD")
    
    Call Text_Clear
    
    '-------------------------------------------------------------------
    ' SP_R_부자재
    '-------------------------------------------------------------------
    If Server_Connection(HostCon, "LAUNDRY1000") = True Then
        ReDim sValue(0)
        
        sValue(0) = "0"
        
        Set ADORs = New ADODB.Recordset
        Set ADORs = SP_Exec(HostCon, "SP_R_부자재", sValue(), Err_Num, Err_Desc)
                
        With cboGoods
            .Clear
        
            Do Until ADORs.EOF
                .AddItem ADORs!부자재명 & "": .ItemData(.NewIndex) = ADORs!부자재코드
            
                ADORs.MoveNext
            Loop
            ADORs.Close
            Set ADORs = Nothing
            
            .ListIndex = -1
        End With
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    cmdBtn(5).Left = Me.Width - cmdBtn(5).Width - 200
End Sub

Private Sub Form_Unload(Cancel As Integer)
    HostCon.Close
    Set HostCon = Nothing
End Sub

Private Sub sprGrid_Click(ByVal Col As Long, ByVal Row As Long)
    Dim 주문코드 As Long
    
    If Row <= 0 Then Exit Sub
        
    sprGrid.Row = Row
    sprGrid.Col = 1: 주문코드 = sprGrid.Text & ""
    
    If Server_Connection(HostCon, "LAUNDRY1000") = False Then Exit Sub
    
    ReDim sValue(3)
    
    sValue(0) = 가맹점정보.가맹점코드
    sValue(1) = ""
    sValue(2) = ""
    sValue(3) = 주문코드
    
    Set ADORs = New ADODB.Recordset
    Set ADORs = SP_Exec(HostCon, "SP_R_부자재주문", sValue(), Err_Num, Err_Desc)
    
    If ADORs.EOF Then
    
    Else
        pnlNo.Caption = ADORs!주문코드 & ""               ' 1
        pnlOrderDate.Caption = Format(ADORs!주문일자, "") ' 2
        
        With cboGoods
            For i = 0 To .ListCount - 1
                If .ItemData(i) = ADORs!부자재코드 Then   ' 3
                    .ListIndex = i
                    Exit For
                End If
            Next i
        End With
        
        txtNum(0).Value = ADORs!수량 & ""                 ' 4
        txtData(0).Text = ADORs!비고 & ""                 ' 5
        
        pnlDay(0).Caption = ADORs!출고일자 & ""           ' 6
        pnlDay(1).Caption = ADORs!확정일자 & ""           ' 7
    End If
    ADORs.Close
    Set ADORs = Nothing
End Sub

Private Sub txtNum_Change(Index As Integer)
'    If Index = 0 Or Index = 1 Then
'        txtNum(2).Value = txtNum(0).Value * txtNum(1).Value '공급가액
'        txtNum(3).Value = txtNum(2).Value * 0.1             '세액
'        txtNum(4).Value = txtNum(2).Value + txtNum(3).Value '합계금액
'    End If
End Sub
