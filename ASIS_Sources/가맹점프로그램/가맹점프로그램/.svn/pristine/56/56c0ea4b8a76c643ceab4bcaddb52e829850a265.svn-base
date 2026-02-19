VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm근무자 
   BorderStyle     =   1  '단일 고정
   Caption         =   "근무자"
   ClientHeight    =   5655
   ClientLeft      =   3735
   ClientTop       =   4125
   ClientWidth     =   6045
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm근무자.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   6045
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   5655
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   9975
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm근무자.frx":08CA
      Begin Threed.SSPanel SSPanel2 
         Height          =   540
         Left            =   15
         TabIndex        =   13
         Top             =   1245
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   953
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   45
            TabIndex        =   3
            Top             =   45
            Width           =   1275
            _Version        =   851970
            _ExtentX        =   2249
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 신규(&N)"
            Appearance      =   6
            Picture         =   "frm근무자.frx":093C
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   1350
            TabIndex        =   2
            Top             =   45
            Width           =   1275
            _Version        =   851970
            _ExtentX        =   2249
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 저장(&S)"
            Appearance      =   6
            Picture         =   "frm근무자.frx":134E
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   2655
            TabIndex        =   4
            Top             =   45
            Width           =   1275
            _Version        =   851970
            _ExtentX        =   2249
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 삭제(&D)"
            Appearance      =   6
            Picture         =   "frm근무자.frx":1D60
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   4695
            TabIndex        =   5
            Top             =   45
            Width           =   1275
            _Version        =   851970
            _ExtentX        =   2249
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            Appearance      =   6
            Picture         =   "frm근무자.frx":2772
         End
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   3840
         Left            =   15
         TabIndex        =   6
         Top             =   1800
         Width           =   6015
         _Version        =   524288
         _ExtentX        =   10610
         _ExtentY        =   6773
         _StockProps     =   64
         BackColorStyle  =   1
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
         DisplayRowHeaders=   0   'False
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
         ScrollBars      =   2
         ShadowColor     =   14737632
         SpreadDesigner  =   "frm근무자.frx":3184
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   1215
         Left            =   15
         TabIndex        =   9
         Top             =   15
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   2143
         _Version        =   262144
         BackColor       =   16777215
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtData 
            Appearance      =   0  '평면
            Height          =   375
            IMEMode         =   10  '한글 
            Index           =   1
            Left            =   1065
            TabIndex        =   1
            Top             =   780
            Width           =   4890
         End
         Begin VB.TextBox txtData 
            Appearance      =   0  '평면
            Height          =   375
            IMEMode         =   10  '한글 
            Index           =   0
            Left            =   1065
            TabIndex        =   0
            Top             =   420
            Width           =   4890
         End
         Begin CSTextLibCtl.silgEdit txtCode 
            Height          =   375
            Left            =   1065
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   60
            Width           =   1110
            _Version        =   262145
            _ExtentX        =   1958
            _ExtentY        =   661
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.74
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataProperty    =   2
            ReadOnly        =   -1  'True
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "0"
            Text            =   " 0"
            StartText.x     =   2
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
            Justification   =   1
            FmtThousands    =   0
            FmtControl      =   1
            MinValue        =   0
            Undo            =   1
            Data            =   0
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   375
            Index           =   3
            Left            =   60
            TabIndex        =   10
            Top             =   60
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   661
            _Version        =   262144
            BackColor       =   12648384
            Caption         =   "근무코드"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm근무자.frx":37F3
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   375
            Index           =   4
            Left            =   60
            TabIndex        =   11
            Top             =   420
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   661
            _Version        =   262144
            BackColor       =   12648384
            Caption         =   "근무자명"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm근무자.frx":3A15
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   375
            Index           =   17
            Left            =   60
            TabIndex        =   12
            Top             =   780
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   661
            _Version        =   262144
            BackColor       =   12648384
            Caption         =   "비    고"
            PictureBackgroundStyle=   2
            PictureBackground=   "frm근무자.frx":3C37
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frm근무자"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0
            Call Text_Clear
            
            txtData(0).SetFocus
            
        Case 1
            If txtCode.Value = 0 Then
                Query = "SELECT MAX(근무자코드) FROM TB_근무자"
                Set ADORs = New ADODB.Recordset
                ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
                
                If ADORs.EOF Then
                    txtCode.Value = 1
                Else
                    If IsNull(ADORs(0)) Then
                        txtCode.Value = 1
                    Else
                        txtCode.Value = ADORs(0) + 1
                    End If
                End If
                ADORs.Close
                Set ADORs = Nothing
            End If
                
            If Trim(txtData(0).Text) = "" Then
                MsgBox "근무자명을 입력하세요.", vbInformation, "확인"
            
                txtData(0).SetFocus
                Exit Sub
            End If
            
            '---------------------------------------------------------
            '
            '---------------------------------------------------------
            Query = "SELECT * FROM TB_근무자"
            Query = Query & " WHERE 근무자코드 = " & txtCode.Value
            Set ADORs = New ADODB.Recordset
            ADORs.Open Query, ADOCon, adOpenDynamic, adLockOptimistic
            
            If ADORs.EOF Then ADORs.AddNew
            
            ADORs!근무자코드 = txtCode.Value      ' 1
            ADORs!근무자명 = txtData(0).Text & "" ' 2
            ADORs!비고 = txtData(1).Text & ""     ' 3
            
            ADORs.Update
            
            ADORs.Close
            Set ADORs = Nothing
            
            Call Data_Display
            Call Text_Clear
            
            txtData(0).SetFocus
            
        Case 2
            Query = "정말로 삭제하시겠습니까?"
            Rtn = MsgBox(Query, vbQuestion + vbYesNo + vbDefaultButton2, "확인")
            
            If Rtn = vbYes Then
                Query = "DELETE FROM TB_근무자"
                Query = Query & " WHERE 근무자코드 = " & txtCode.Value
                ADOCon.Execute Query
                
                Call Data_Display
                Call Text_Clear
                
                txtData(0).SetFocus
            End If
            
        Case 3
            Unload Me
    End Select
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub Text_Clear()
    On Error GoTo ErrRtn
    
    txtCode.Value = 0
    txtData(0).Text = ""
    txtData(1).Text = ""
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn
    
    Query = "SELECT * FROM TB_근무자"
    Query = Query & " ORDER BY 근무자명 ASC"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    With sprGrid
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = ADORs!근무자코드 & "" '1
            .Col = 2: .Text = ADORs!근무자명 & ""   '2
            .Col = 3: .Text = ADORs!비고 & ""       '3
            
            ADORs.MoveNext
        Loop
        
        ADORs.Close
        Set ADORs = Nothing
    
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    
    
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

    Call Data_Display
    Call Text_Clear
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub sprGrid_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    
    Call Text_Clear
    
    With sprGrid
        .Row = Row
        .Col = 1
        Query = "SELECT * FROM TB_근무자"
        Query = Query & " WHERE 근무자코드 = " & .Text
        Set ADORs = New ADODB.Recordset
        ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
                    
        If Not ADORs.EOF Then
            txtCode.Value = ADORs!근무자코드
            txtData(0).Text = ADORs!근무자명 & ""
            txtData(1).Text = ADORs!비고 & ""
        End If
        ADORs.Close
        Set ADORs = Nothing
    End With
End Sub

Private Sub sprGrid_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Call sprGrid_Click(NewCol, NewRow)
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
