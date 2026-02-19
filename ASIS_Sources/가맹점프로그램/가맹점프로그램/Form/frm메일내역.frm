VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RichTx32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm메일내역 
   Caption         =   "공지사항 내역"
   ClientHeight    =   12345
   ClientLeft      =   5865
   ClientTop       =   2610
   ClientWidth     =   15930
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
   LinkTopic       =   "Form23"
   MDIChild        =   -1  'True
   ScaleHeight     =   12345
   ScaleWidth      =   15930
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   12345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15930
      _ExtentX        =   28099
      _ExtentY        =   21775
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm메일내역.frx":0000
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   11115
         Left            =   15
         TabIndex        =   1
         Top             =   1215
         Width           =   3525
         _Version        =   524288
         _ExtentX        =   6218
         _ExtentY        =   19606
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowUserFormulas=   -1  'True
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
         DInformActiveRowChange=   0   'False
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
         MaxCols         =   4
         MaxRows         =   300
         MoveActiveOnFocus=   0   'False
         OperationMode   =   3
         Protect         =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         SpreadDesigner  =   "frm메일내역.frx":00B2
         VisibleCols     =   2
         VisibleRows     =   50
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   750
         Left            =   15
         TabIndex        =   2
         Top             =   450
         Width           =   15900
         _ExtentX        =   28046
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.OptionButton optSelect 
            Caption         =   "송신"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   1770
            TabIndex        =   9
            Top             =   405
            Width           =   825
         End
         Begin VB.OptionButton optSelect 
            Caption         =   "수신"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   960
            TabIndex        =   8
            Top             =   405
            Value           =   -1  'True
            Width           =   825
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   0
            Left            =   915
            TabIndex        =   4
            Top             =   60
            Width           =   1425
            _ExtentX        =   2514
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
            Format          =   64159747
            CurrentDate     =   40279
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   315
            Index           =   1
            Left            =   2610
            TabIndex        =   5
            Top             =   60
            Width           =   1425
            _ExtentX        =   2514
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
            Format          =   64159747
            CurrentDate     =   40279
         End
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   630
            Left            =   4530
            TabIndex        =   11
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 조회(&F)"
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
            Picture         =   "frm메일내역.frx":076C
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   3
            Left            =   9780
            TabIndex        =   12
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 엑셀(&E)"
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
            Picture         =   "frm메일내역.frx":0E66
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   12870
            TabIndex        =   13
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
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
            Picture         =   "frm메일내역.frx":15E0
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   4
            Left            =   11325
            TabIndex        =   14
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 출력(&P)"
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
            Picture         =   "frm메일내역.frx":2672
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   1
            Left            =   8235
            TabIndex        =   15
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 삭제(&D)"
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
            Picture         =   "frm메일내역.frx":2D6C
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   0
            Left            =   6690
            TabIndex        =   16
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 메일작성"
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
            Picture         =   "frm메일내역.frx":3DFE
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "구    분:"
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
            TabIndex        =   10
            Top             =   465
            Width           =   840
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "공지일자:"
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
            TabIndex        =   7
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
            Height          =   180
            Index           =   0
            Left            =   2415
            TabIndex        =   6
            Top             =   120
            Width           =   120
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   3
         Top             =   15
         Width           =   15900
         _ExtentX        =   28046
         _ExtentY        =   741
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   0
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
         Caption         =   "      공지사항 내역"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm메일내역.frx":4E90
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm메일내역.frx":50B6
            Top             =   -15
            Width           =   765
         End
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   10560
         Left            =   3555
         TabIndex        =   17
         Top             =   1770
         Width           =   12360
         _ExtentX        =   21802
         _ExtentY        =   18627
         _Version        =   393217
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frm메일내역.frx":5C80
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   540
         Left            =   3555
         TabIndex        =   18
         Top             =   1215
         Width           =   12360
         _ExtentX        =   21802
         _ExtentY        =   953
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtData 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   12
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            IMEMode         =   10  '한글 
            Index           =   0
            Left            =   1200
            TabIndex        =   21
            Top             =   60
            Width           =   9600
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   480
            Left            =   60
            TabIndex        =   19
            Top             =   60
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   847
            _Version        =   262144
            BackColor       =   16777215
            Enabled         =   0   'False
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "제목 :"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   3
               Left            =   390
               TabIndex        =   20
               Top             =   165
               Width           =   540
            End
         End
      End
   End
End
Attribute VB_Name = "frm메일내역"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: frm메일작성.Show 1
        
        Case 1
            Dim sData As String
            Dim iSEQ  As Integer
            
            With sprGrid
                .Row = .ActiveRow
                .Col = 1: sData = Format(.Text, "YYYY-MM-DD")
                
                If Trim(sData) = "" Then Exit Sub
                
                .Col = 2: iSEQ = .Value
            
                '------------------------------------------------------------
                '
                '------------------------------------------------------------
                Query = "DELETE FROM TB_공지사항"
                Query = Query & " WHERE 작성일자 = '" & sData & "'"
                Query = Query & " AND   문서번호 = " & iSEQ
                
                If optSelect(0).Value = True Then
                    Query = Query & "  AND   공지구분 = '2'"
                Else
                    Query = Query & "  AND   공지구분 = '1'"
                End If
                ADOCon.Execute Query
                
                .Action = ActionDeleteRow
            End With
            
        Case 3: Call Export_Excel(frmMain.cdgExcel, sprGrid)
        Case 4: Call DataPrint
        Case 5: Unload Me
    End Select
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub cmdList_Click()
    Dim sData1 As String
    Dim sData2 As String
    
    On Error GoTo ErrRtn
    
    sData1 = Format(dtpDay(0).Value, "YYYY-MM-DD")
    sData2 = Format(dtpDay(1).Value, "YYYY-MM-DD")

    '------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------
    Query = "SELECT    작성일자"
    Query = Query & ", 문서번호"
    Query = Query & ", 본사전송여부"
    Query = Query & ", 파일명"
    Query = Query & " FROM TB_공지사항"
    Query = Query & " WHERE 작성일자 >= '" & sData1 & "'"
    Query = Query & "   AND 작성일자 <= '" & sData2 & "'"

    If optSelect(0).Value = True Then
        Query = Query & " AND   공지구분 = '2' "
    Else
        Query = Query & " AND   공지구분 = '1' "
    End If
    
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprGrid
        .MaxRows = 0
        .ReDraw = False
    
        Do While Not ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = Format(ADORs!작성일자, "YYYY-MM-DD")
            .Col = 2: .Text = ADORs!문서번호 & ""
            .Col = 3: .Value = IIf(ADORs!본사전송여부 = "Y", True, False)
            .Col = 4: .Text = ADORs!파일명 & ""
            
            ADORs.MoveNext
        Loop
        ADORs.Close
        Set ADORs = Nothing
            
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
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
        
        .Col = 4:   .ColHidden = True
        
    End With
    
    dtpDay(0).Value = Format(Date, "YYYY-MM-01")
    dtpDay(1).Value = Format(Date, "YYYY-MM-DD")
    
    'TitleSet "편지내역조회"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    cmdBtn(5).Left = Me.Width - cmdBtn(5).Width - 200
End Sub

Private Sub sprGrid_Click(ByVal Col As Long, ByVal Row As Long)
    Dim sData As String
    Dim iSEQ As Integer
    Dim sFile   As String
    
    On Error GoTo ErrRtn
    
    
    RichTextBox1.Text = ""
    
    sprGrid.Row = sprGrid.ActiveRow
    
    sprGrid.Col = 1
    If Trim(sprGrid.Text) = "" Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    sData = Format(sprGrid.Text, "YYYY-MM-DD")
    sprGrid.Col = 2: iSEQ = sprGrid.Value
    sprGrid.Col = 4: sFile = sprGrid.Text

    '-------------------------------------------------------
    '
    '-------------------------------------------------------
    Query = "SELECT 공지내용"
    Query = Query & " FROM TB_공지사항 "
    Query = Query & " WHERE 작성일자 = '" & sData & "'"
    Query = Query & "   AND 문서번호 =  " & iSEQ
    
    If optSelect(0).Value = True Then
        Query = Query & "  AND   공지구분 = '2'"
    Else
        Query = Query & "  AND   공지구분 = '1'"
    End If
    
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If Not ADORs.EOF Then
        If sFile = "" Then
            RichTextBox1.Text = GetMailConvert(ADORs!공지내용 & "", "READ")
            txtData(0).Text = ""
            
        Else
            txtData(0).Text = GetMailConvert(ADORs!공지내용 & "", "READ")
        
            Call DataPCSaveFileView(sData, CStr(iSEQ), RichTextBox1)
        
        End If
    End If
    ADORs.Close:    Set ADORs = Nothing
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
ErrRtn:
    Screen.MousePointer = vbDefault
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

End Sub


Private Sub DataPrint()
    With RichTextBox1
        .SelLength = 0
'        .SelStart = 0
'        .SelLength = Len(.TextRTF)
        
'        Printer.Print .Text
        
        .SelPrint Printer.hdc
        Printer.EndDoc
        
    End With
        
End Sub
