VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm입고작업 
   Caption         =   "가맹점 입고작업"
   ClientHeight    =   10080
   ClientLeft      =   2490
   ClientTop       =   3135
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
   Begin Threed.SSPanel pnlProg 
      Height          =   1215
      Left            =   615
      TabIndex        =   2
      Top             =   2310
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
      Picture         =   "frm입고작업.frx":0000
      Caption         =   " "
      BorderWidth     =   0
      BevelOuter      =   1
      BevelInner      =   2
      Alignment       =   4
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   10080
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   17780
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm입고작업.frx":2FCB
      Begin Threed.SSPanel SSPanel1 
         Height          =   720
         Left            =   15
         TabIndex        =   12
         Top             =   1155
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   1270
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   0
            Left            =   45
            TabIndex        =   13
            Top             =   45
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 입고완료"
            Appearance      =   6
            Picture         =   "frm입고작업.frx":305D
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   2
            Left            =   6195
            TabIndex        =   14
            Top             =   45
            Width           =   1395
            _Version        =   851970
            _ExtentX        =   2461
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frm입고작업.frx":3757
         End
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   8175
         Left            =   15
         TabIndex        =   3
         Top             =   1890
         Width           =   15210
         _Version        =   524288
         _ExtentX        =   26829
         _ExtentY        =   14420
         _StockProps     =   64
         AutoCalc        =   0   'False
         BackColorStyle  =   1
         ColsFrozen      =   7
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         MaxCols         =   10
         MaxRows         =   200
         ScrollBars      =   2
         SpreadDesigner  =   "frm입고작업.frx":47E9
         UserResize      =   1
         VisibleCols     =   7
         VisibleRows     =   30
         HighlightHeaders=   1
         HighlightStyle  =   1
         ScrollBarStyle  =   2
      End
      Begin Threed.SSPanel Panel 
         Height          =   690
         Left            =   15
         TabIndex        =   4
         Top             =   450
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   1217
         _Version        =   262144
         BackColor       =   16777215
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin CSTextLibCtl.sitxEdit txtTag 
            Height          =   600
            Index           =   0
            Left            =   900
            TabIndex        =   0
            Top             =   45
            Width           =   2085
            _Version        =   262145
            _ExtentX        =   3678
            _ExtentY        =   1058
            _StockProps     =   125
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   17.99
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   2
            StartText.y     =   5
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   29
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   1
            CharacterTable  =   ""
         End
         Begin CSTextLibCtl.sitxEdit txtTag 
            Height          =   600
            Index           =   1
            Left            =   4365
            TabIndex        =   5
            Top             =   45
            Width           =   8370
            _Version        =   262145
            _ExtentX        =   14764
            _ExtentY        =   1058
            _StockProps     =   125
            ForeColor       =   -2147483640
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   17.99
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   2
            StartText.y     =   5
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   29
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   1
            CharacterTable  =   ""
         End
         Begin CSTextLibCtl.sitxEdit txtCount 
            Height          =   600
            Left            =   14040
            TabIndex        =   6
            Top             =   45
            Width           =   1050
            _Version        =   262145
            _ExtentX        =   1852
            _ExtentY        =   1058
            _StockProps     =   125
            ForeColor       =   192
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   17.99
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   2
            StartText.y     =   5
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   29
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   1
            CharacterTable  =   ""
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "건수:"
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
            Left            =   13410
            TabIndex        =   11
            Top             =   90
            Width           =   570
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "메시지:"
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
            Left            =   3510
            TabIndex        =   10
            Top             =   90
            Width           =   795
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "택번호:"
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
            Index           =   2
            Left            =   60
            TabIndex        =   9
            Top             =   90
            Width           =   795
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   7
         Top             =   15
         Width           =   15210
         _ExtentX        =   26829
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
         Caption         =   "      가맹점 입고작업"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm입고작업.frx":5817
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm입고작업.frx":5A3D
            Top             =   -15
            Width           =   765
         End
         Begin VB.Label lblTag 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   225
            Left            =   10830
            TabIndex        =   8
            Top             =   90
            Width           =   105
         End
      End
   End
End
Attribute VB_Name = "frm입고작업"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn

    Select Case Index
        Case 0:
            
            With sprGrid
                If .MaxRows = 0 Then
                    MsgBox "지점입고 세탁물이 없습니다.", vbInformation, "확인"
                    
                    Exit Sub
                End If
                                
                pnlProg.Visible = True
                DoEvents
                
                For i = 1 To .MaxRows
                    .Row = i
                    .Col = 1
                    If .Text = "1" Then
                                  Query = "UPDATE TB_입출고 SET 가맹점입고일자 = '" & Format(Now, "YYYY-MM-DD hh:mm:ss") & "'"
                                  Query = Query & "           , 본사전송여부   = ''"
                        .Col = 2: Query = Query & " WHERE 접수일자 = '" & .Text & "'"
                        .Col = 4: Query = Query & "   AND 택번호   = '" & .Text & "'"
                        ADOCon.Execute Query
                    End If
                Next i
                                
                .MaxRows = 0
            End With
            
            pnlProg.Visible = False
                
        Case 2: Unload Me
    End Select
    
    Exit Sub
    
ErrRtn:
    pnlProg.Visible = False
    
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
    
    Screen.MousePointer = 0
End Sub

Private Sub Form_Activate()
    'ActiveTForm = "입고작업"
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
        
        ' the User Sort
        .UserColAction = UserColActionSort
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    pnlHeader.Width = Me.ScaleWidth
    
    cmdBtn(2).Left = Me.Width - cmdBtn(2).Width - 200
    lblTag.Left = Me.Width - lblTag.Width - 250
    
    txtCount.Left = Me.Width - txtCount.Width - 200
    lblTitle.Left = txtCount.Left - lblTitle.Width - 100
End Sub

'frmMain MSComm에서 입력받으므로 Public으로 선언
Public Sub txtTag_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim BarCodeTag As String
    
    On Error GoTo ErrRtn

    txtTag(1).Text = ""
    
    If Index = 0 And KeyAscii = 13 Then
        If (txtTag(0).Text = "") Or Len(txtTag(0).Text) < 9 Then
            Exit Sub
        End If
        
        BarCodeTag = TAG_Convert(txtTag(0).Text)
        
        lblTag.Caption = BarCodeTag
        
        If Left(BarCodeTag, 3) <> 가맹점정보.택코드 Then
            txtTag(1).Text = "다른 지점의 세탁물입니다."
            'Ret = sndPlaySound(AppPath & "Sound\다른지점.wav", 0)
        
            txtTag(0).Text = ""
            Exit Sub
        End If
        
        '------------------------------------------------------------------------------------
        ' 택번호를 중복해서 읽었는지 체크...
        '------------------------------------------------------------------------------------
        BarCodeTag = Left(BarCodeTag, 9)
        
        With sprGrid
            For i = 1 To .MaxRows
                .Row = i
                .Col = 2
                
                If .Text = BarCodeTag Then
                    txtTag(1).Text = "택번호 중복"
                    'Ret = sndPlaySound(AppPath & "Sound\택번호중복.wav", 0)

                    txtTag(0).Text = ""
                    
                    Exit Sub
                End If
            Next i
        End With
        
        '-----------------------------------------------------------------------------------
        ' 해당지점에서 출고된 세탁물(바코드)인지 체크...
        '-----------------------------------------------------------------------------------
        Query = "SELECT    접수일자"
        Query = Query & ", 의류명"
        Query = Query & ", 택번호"
        Query = Query & ", 색상"
        Query = Query & ", 무늬"
        Query = Query & ", 내용"
        Query = Query & ", 금액"
        Query = Query & ", 상표"
        Query = Query & ", 의류코드"
        Query = Query & " FROM TB_입출고"
        Query = Query & " WHERE 택번호 = '" & BarCodeTag & "'"
        'Query = Query & "   AND (가맹점출고일자 <> '' OR 가맹점출고일자 IS NOT NULL)"
        Query = Query & "   AND (가맹점입고일자  = '' OR 가맹점입고일자 IS NULL) "
        Set ADORs = New ADODB.RecordSet
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
        If ADORs.EOF Then
            txtTag(1).Text = "이미 입고처리된 택번호(" & BarCodeTag & ") 입니다."
        Else
            With sprGrid
                Rtn = .SearchCol(4, -1, -1, Format(BarCodeTag, "000-00-0000"), SearchFlagsValue) '신규저장하거나 수정한 데이터 위치로 이동
                
                If Rtn = -1 Then
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    
                    .Col = 1:  .Text = "1"
                    .Col = 2:  .Text = ADORs!접수일자 & ""
                    .Col = 3:  .Text = ADORs!의류명 & ""
                    .Col = 4:  .Text = Format(ADORs!택번호, "000-00-0000") & ""
                    .Col = 5:  .Text = ADORs!색상 & ""
                    .Col = 6:  .Text = ADORs!무늬 & ""
                    .Col = 7:  .Text = ADORs!내용 & ""
                    .Col = 8:  .Text = ADORs!금액 & ""
                    .Col = 9:  .Text = ADORs!상표 & ""
                    .Col = 10: .Text = ADORs!의류코드 & ""
                Else
                    '이미 입고작업을 한 택번호
                    Call .SetSelection(1, Rtn, .MaxCols, Rtn)
                End If
            End With
            
            Beep
            
            Ret = sndPlaySound(AppPath & "Sound\alert.wav", 0)
        End If
        ADORs.Close
        Set ADORs = Nothing
        
        txtCount.Text = sprGrid.MaxRows
        
        txtTag(0).Text = ""
        txtTag(0).SetFocus
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
    
    Screen.MousePointer = 0
End Sub
