VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm의류현황 
   Caption         =   "의류 현황"
   ClientHeight    =   8340
   ClientLeft      =   4050
   ClientTop       =   4665
   ClientWidth     =   10905
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
   LinkTopic       =   "Form9"
   LockControls    =   -1  'True
   ScaleHeight     =   8340
   ScaleWidth      =   10905
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8340
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   14711
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm의류현황.frx":0000
      Begin Threed.SSPanel SSPanel1 
         Height          =   750
         Left            =   15
         TabIndex        =   1
         Top             =   450
         Width           =   10875
         _ExtentX        =   19182
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdDown 
            Height          =   630
            Left            =   6885
            TabIndex        =   3
            Top             =   60
            Width           =   735
            _Version        =   851970
            _ExtentX        =   1296
            _ExtentY        =   1111
            _StockProps     =   79
            Appearance      =   6
            Picture         =   "frm의류현황.frx":0092
         End
         Begin XtremeSuiteControls.PushButton cmdUp 
            Height          =   630
            Left            =   6105
            TabIndex        =   4
            Top             =   60
            Width           =   735
            _Version        =   851970
            _ExtentX        =   1296
            _ExtentY        =   1111
            _StockProps     =   79
            Appearance      =   6
            Picture         =   "frm의류현황.frx":078C
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   9330
            TabIndex        =   5
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm의류현황.frx":0E86
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   1
            Left            =   7755
            TabIndex        =   6
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 저장(&S)"
            Appearance      =   6
            Picture         =   "frm의류현황.frx":1F18
         End
         Begin VB.Label Label1 
            Caption         =   "각 항목의 첫번째 품목이 자동으로 접수됩니다."
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   270
            TabIndex        =   9
            Top             =   300
            Width           =   4995
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   10875
         _ExtentX        =   19182
         _ExtentY        =   741
         _Version        =   262144
         Font3D          =   3
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
         Caption         =   "      의류 현황"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm의류현황.frx":2FAA
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm의류현황.frx":31D0
            Top             =   -15
            Width           =   765
         End
      End
      Begin FPSpreadADO.fpSpread sprClass 
         Height          =   7110
         Left            =   15
         TabIndex        =   7
         Top             =   1215
         Width           =   3765
         _Version        =   524288
         _ExtentX        =   6641
         _ExtentY        =   12541
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BackColorStyle  =   1
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
         MaxCols         =   3
         MaxRows         =   30
         Protect         =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frm의류현황.frx":3D9A
         VisibleCols     =   2
         VisibleRows     =   30
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   7110
         Left            =   3795
         TabIndex        =   8
         Top             =   1215
         Width           =   7095
         _Version        =   524288
         _ExtentX        =   12515
         _ExtentY        =   12541
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BackColorStyle  =   1
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
         MaxCols         =   5
         MaxRows         =   30
         Protect         =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frm의류현황.frx":4380
         VisibleCols     =   3
         VisibleRows     =   30
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "frm의류현황"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tmpData1(1 To 5) As String
Dim tmpData2(1 To 5) As String
Dim iRow             As Integer

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 1:
            With sprGrid
                For i = 1 To .MaxRows
                    .Row = i
                    
                    .Col = 1: Query = "UPDATE TB_의류 SET 순서 = '" & .Text & "'"
                    .Col = 2: Query = Query & " WHERE 의류코드 = '" & .Text & "'"
                    ADOCon.Execute Query
                    
                    .Col = 1: Query = "UPDATE TB_할인정보 SET 순서 = '" & .Text & "'"
                    .Col = 2: Query = Query & " WHERE 의류코드 = '" & .Text & "'"
                    ADOCon.Execute Query
                    
                    .Col = 1: Query = "UPDATE TB_요일할인 SET 순서 = '" & .Text & "'"
                    .Col = 2: Query = Query & " WHERE 의류코드 = '" & .Text & "'"
                    ADOCon.Execute Query
                
                Next i
            End With
            
            MsgBox "올바르게 저장되었습니다.", vbInformation, "확인"
            
        Case 5: Unload Me
    End Select
End Sub

Private Sub cmdDown_Click()
    With sprGrid
        If .ActiveRow = .MaxRows Then Exit Sub
        
        iRow = .ActiveRow
        
        .Row = iRow
        
         For i = 1 To 5
            .Col = i: tmpData1(i) = .Text
         Next i
         
        .Row = iRow + 1
        
         For i = 1 To 5
            .Col = i: tmpData2(i) = .Text
         Next i
                 
        '============================================
        
        .Row = iRow
        
         For i = 2 To 5
            .Col = i: .Text = tmpData2(i)
         Next i
         
        .Row = iRow + 1
        
         For i = 2 To 5
            .Col = i: .Text = tmpData1(i)
         Next i
         
         .SetActiveCell 1, iRow + 1
    End With
End Sub

Private Sub cmdUp_Click()
    With sprGrid
        If .ActiveRow <= 1 Then Exit Sub
        
        iRow = .ActiveRow
        
        .Row = iRow
        
         For i = 1 To 5
            .Col = i: tmpData1(i) = .Text
         Next i
         
        .Row = iRow - 1
        
         For i = 1 To 5
            .Col = i: tmpData2(i) = .Text
         Next i
                 
        '============================================
        
        .Row = iRow
        
         For i = 2 To 5
            .Col = i: .Text = tmpData2(i)
         Next i
         
        .Row = iRow - 1
        
         For i = 2 To 5
            .Col = i: .Text = tmpData1(i)
         Next i
         
         .SetActiveCell 1, iRow - 1
    End With
End Sub

Private Sub Form_Load()
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2

    
    With sprClass
        .MaxRows = 0
        .RowHeight(-1) = 14
    
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle

        '선택된 Row
        .SelBackColor = &H80FFFF '&HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeSingle
                
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
     
    With sprGrid
        .MaxRows = 0
        .RowHeight(-1) = 14
    
        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle

        '선택된 Row
        .SelBackColor = &H80FFFF '&HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeSingle
                
        'Init the User Sort
        .UserColAction = UserColActionSort
    End With
        
    '----------------------------------------------------------
    '
    If 가맹점정보.지사코드 = "1024" And Format(Date, "yyyy-MM-dd") >= "2013-10-21" Then
        '----------------------------------------------------------
        Query = "SELECT    A.의류분류코드"
        Query = Query & ", A.의류분류명"
        Query = Query & ", A.순서"
        Query = Query & ", COUNT(B.의류명) AS 수량"
        Query = Query & " FROM TB_의류분류 AS A LEFT OUTER JOIN TB_의류 AS B ON A.의류분류코드 = SUBSTRING(B.의류코드,1,2)"
        Query = Query & " WHERE substring(A.의류분류코드,1,1) <> 'w'"
        Query = Query & " GROUP BY A.의류분류코드, A.의류분류명, A.순서"
        Query = Query & " ORDER BY A.순서, A.의류분류코드 ASC"
                
                
                
    Else
        '----------------------------------------------------------
        Query = "SELECT    A.의류분류코드"
        Query = Query & ", A.의류분류명"
        Query = Query & ", A.순서"
        Query = Query & ", COUNT(B.의류명) AS 수량"
        Query = Query & " FROM TB_의류분류 AS A LEFT OUTER JOIN TB_의류 AS B ON A.의류분류코드 = SUBSTRING(B.의류코드,1,2)"
        Query = Query & " GROUP BY A.의류분류코드, A.의류분류명, A.순서"
        Query = Query & " ORDER BY A.순서, A.의류분류코드 ASC"
    End If
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    With sprClass
        .MaxRows = 0
        .ReDraw = False
                
        Do Until SUBRs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = SUBRs!의류분류코드 & ""
            .Col = 2: .Text = SUBRs!의류분류명 & ""
            .Col = 3: .Text = SUBRs!수량 & ""
            
            SUBRs.MoveNext
        Loop
        SUBRs.Close
        Set SUBRs = Nothing
        
        .ReDraw = True
    End With
        
    Exit Sub
    
ErrRtn:

End Sub

Private Sub Data_Display(의류분류코드 As String)
    Dim 의류코드 As String
    
    On Error GoTo ErrRtn
    
    Query = "SELECT * FROM TB_의류"
    Query = Query & " WHERE SUBSTRING(의류코드,1,2) = '" & 의류분류코드 & "'"
    Query = Query & " ORDER BY 순서, 의류코드 ASC"
    Set SUBRs = New ADODB.Recordset
    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    i = 0
    
    With sprGrid
        .MaxRows = 0
        .ReDraw = False
                
        Do Until SUBRs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            i = i + 1
            
            .Col = 1: .Text = 의류분류코드 & Format(i, "00") & ""
            .Col = 2: .Text = SUBRs!의류코드 & ""
            .Col = 3: .Text = SUBRs!의류명 & ""
            .Col = 4: .Text = SUBRs!금액 & ""
            .Col = 5: .Text = SUBRs!비고 & ""
            
            SUBRs.MoveNext
        Loop
        SUBRs.Close
        Set SUBRs = Nothing
        
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrRtn:

End Sub

Private Sub sprClass_Click(ByVal Col As Long, ByVal Row As Long)
    Dim 의류분류코드 As String
    
    If Row <= 0 Then Exit Sub

    sprClass.Row = Row
    sprClass.Col = 1: 의류분류코드 = sprClass.Text & ""
    
    Call Data_Display(의류분류코드)
End Sub

Private Sub sprClass_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Call sprClass_Click(NewCol, NewRow)
End Sub
