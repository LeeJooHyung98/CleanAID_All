VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm의류분류현황 
   Caption         =   "의류분류 현황"
   ClientHeight    =   8340
   ClientLeft      =   14730
   ClientTop       =   6225
   ClientWidth     =   7125
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
   ScaleWidth      =   7125
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8340
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   14711
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm의류분류현황.frx":0000
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   7110
         Left            =   15
         TabIndex        =   1
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
         MaxCols         =   6
         MaxRows         =   30
         Protect         =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frm의류분류현황.frx":0072
         VisibleCols     =   3
         VisibleRows     =   30
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   750
         Left            =   15
         TabIndex        =   2
         Top             =   450
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdDown 
            Height          =   630
            Left            =   825
            TabIndex        =   4
            Top             =   60
            Width           =   735
            _Version        =   851970
            _ExtentX        =   1296
            _ExtentY        =   1111
            _StockProps     =   79
            Appearance      =   6
            Picture         =   "frm의류분류현황.frx":07A2
         End
         Begin XtremeSuiteControls.PushButton cmdUp 
            Height          =   630
            Left            =   45
            TabIndex        =   5
            Top             =   60
            Width           =   735
            _Version        =   851970
            _ExtentX        =   1296
            _ExtentY        =   1111
            _StockProps     =   79
            Appearance      =   6
            Picture         =   "frm의류분류현황.frx":0E9C
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   5550
            TabIndex        =   6
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
            Picture         =   "frm의류분류현황.frx":1596
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   1
            Left            =   2775
            TabIndex        =   7
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 저장(&S)"
            Appearance      =   6
            Picture         =   "frm의류분류현황.frx":2628
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   3
         Top             =   15
         Width           =   7095
         _ExtentX        =   12515
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
         Caption         =   "      의류분류 현황"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm의류분류현황.frx":36BA
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm의류분류현황.frx":38E0
            Top             =   -15
            Width           =   765
         End
      End
   End
End
Attribute VB_Name = "frm의류분류현황"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tmpData1(1 To 6) As String
Dim tmpData2(1 To 6) As String
Dim iRow             As Integer

Private Sub cmdBtn_Click(Index As Integer)
    
    Select Case Index
        Case 1:
            With sprGrid
                For i = 1 To .MaxRows
                    .Row = i
                    .Col = 2
                    
                    Query = "UPDATE TB_의류분류 SET 순서 = " & i
                    Query = Query & " WHERE 의류분류코드 = '" & .Text & "'"
                    ADOCon.Execute Query
                Next i
            End With
            
            ' 의류 (의류 순서를 다시 설정 하기 위하여
            Dim frm As Form
            For Each frm In Forms
                If frm.Name = Trim("frm의류") Then Unload frm 'frm의류
            Next frm
            
            MsgBox "저장되었습니다.     ", vbInformation, "확인"
            
            Unload Me
            
        Case 5: Unload Me
    End Select
End Sub

Private Sub cmdDown_Click()
    With sprGrid
        If .ActiveRow = .MaxRows Then Exit Sub
        
        iRow = .ActiveRow
        
        .Row = iRow
        
         For i = 1 To 6
            .Col = i: tmpData1(i) = .Text
         Next i
         
        .Row = iRow + 1
        
         For i = 1 To 6
            .Col = i: tmpData2(i) = .Text
         Next i
                 
        '============================================
        
        .Row = iRow
        
         For i = 2 To 6
            .Col = i: .Text = tmpData2(i)
         Next i
         
        .Row = iRow + 1
        
         For i = 2 To 6
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
        
         For i = 1 To 6
            .Col = i: tmpData1(i) = .Text
         Next i
         
        .Row = iRow - 1
        
         For i = 1 To 6
            .Col = i: tmpData2(i) = .Text
         Next i
                 
        '============================================
        
        .Row = iRow
        
         For i = 2 To 6
            .Col = i: .Text = tmpData2(i)
         Next i
         
        .Row = iRow - 1
        
         For i = 2 To 6
            .Col = i: .Text = tmpData1(i)
         Next i
         
         .SetActiveCell 1, iRow - 1
    End With
End Sub

Private Sub Form_Load()
    
    frm의류분류현황.Top = frmMain.Top   '1000
    frm의류분류현황.Left = frmMain.Left '6000
    
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
     
    Call Data_Display
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn
    
    Query = "SELECT * FROM TB_의류분류"
    Query = Query & " ORDER BY 순서 ASC"
    Set SUBRs = New ADODB.RecordSet
    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    i = 0
    
    With sprGrid
        .MaxRows = 0
        .ReDraw = False
                
        Do Until SUBRs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            i = i + 1
            
            .Col = 1: .Text = i & ""
            .Col = 2: .Text = SUBRs!의류분류코드 & ""
            .Col = 3: .Text = SUBRs!의류분류명 & ""
            .Col = 4: .Text = SUBRs!세탁마진 & ""
            .Col = 5: .Text = SUBRs!외주마진 & ""
            .Col = 6: .Text = SUBRs!수선마진 & ""
            
            SUBRs.MoveNext
        Loop
        SUBRs.Close
        Set SUBRs = Nothing
        
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrRtn:

End Sub
