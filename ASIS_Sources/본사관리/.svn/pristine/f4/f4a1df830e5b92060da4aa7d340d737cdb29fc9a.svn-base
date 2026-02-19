VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_01001_A1 
   BorderStyle     =   1  '단일 고정
   Caption         =   "가맹점 찾기"
   ClientHeight    =   5670
   ClientLeft      =   6045
   ClientTop       =   8715
   ClientWidth     =   13650
   Icon            =   "P_01001_A1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   13650
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   5670
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13650
      _ExtentX        =   24077
      _ExtentY        =   10001
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_01001_A1.frx":058A
      Begin Threed.SSPanel SSPanel 
         Height          =   540
         Left            =   15
         TabIndex        =   3
         Top             =   15
         Width           =   13620
         _ExtentX        =   24024
         _ExtentY        =   953
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox cboSearch 
            Height          =   300
            ItemData        =   "P_01001_A1.frx":05FC
            Left            =   930
            List            =   "P_01001_A1.frx":05FE
            Style           =   2  '드롭다운 목록
            TabIndex        =   7
            Top             =   90
            Width           =   2325
         End
         Begin XtremeSuiteControls.PushButton cmdSubBtn 
            Height          =   450
            Index           =   0
            Left            =   7230
            TabIndex        =   5
            Top             =   45
            Width           =   1050
            _Version        =   851970
            _ExtentX        =   1852
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 조회"
            Appearance      =   6
            Picture         =   "P_01001_A1.frx":0600
         End
         Begin VB.TextBox txtInput 
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
            IMEMode         =   10  '한글 
            Index           =   0
            Left            =   3240
            TabIndex        =   4
            Top             =   90
            Width           =   3825
         End
         Begin XtremeSuiteControls.PushButton cmdSubBtn 
            Height          =   450
            Index           =   1
            Left            =   8355
            TabIndex        =   6
            Top             =   45
            Width           =   1050
            _Version        =   851970
            _ExtentX        =   1852
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 취소"
            Appearance      =   6
            Picture         =   "P_01001_A1.frx":0B9A
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "검색 구분"
            Height          =   180
            Index           =   3
            Left            =   90
            TabIndex        =   8
            Top             =   135
            Width           =   780
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   4380
         Left            =   15
         TabIndex        =   1
         Top             =   570
         Width           =   13620
         _Version        =   524288
         _ExtentX        =   24024
         _ExtentY        =   7726
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
         SpreadDesigner  =   "P_01001_A1.frx":1134
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panCaption 
         Height          =   690
         Index           =   1
         Left            =   15
         TabIndex        =   2
         Top             =   4965
         Width           =   13620
         _ExtentX        =   24024
         _ExtentY        =   1217
         _Version        =   262144
         BackColor       =   16777215
         Caption         =   "  검색자료 중 선택할 자료를 더블클릭하세요."
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
      End
   End
End
Attribute VB_Name = "P_01001_A1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public l_AgencyCode As String
Public l_TagNo As String
Public l_IpDate As String '
Public m_FormObj   As Object

Dim sValue() As String
Dim RS01 As ADODB.Recordset

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cmdSubBtn_Click(Index As Integer)
    Select Case Index
        Case 0
            
            ReDim sValue(2)
            
            sValue(0) = 0
            sValue(1) = Mid(cboSearch.Text, 1, 2)
            sValue(2) = Trim(txtInput(0).Text)
            
            
            Set RS01 = New ADODB.Recordset
            Set RS01 = ExecPro("SP_01001_A2_ALL", sValue(), Err_Num, Err_Dec)
            
            spdView.MaxCols = RS01.Fields.Count
            spdView.MaxRows = RS01.RecordCount
            
            Call fpSpread_Display(spdView, RS01)
            Call GetColWidth(REG_App, Me.Name, spdView)
            
            DoEvents
            spdView.SetFocus
            
        Case 1: Unload Me
    End Select
End Sub

Private Sub Form_Activate()
    txtInput(0).SetFocus
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    Me.Left = P_00000.Left + 4000
    Me.Top = P_00000.Top + 3000
    
    
    With spdView
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
    
    
        .ColsFrozen = 2  '틀고정
        .Row = -1
    
        .Col = 1
        .ColWidth(1) = 8
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
    
    
        .Col = 2
        .ColWidth(2) = 20
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft

        .Col = 3
        .ColWidth(3) = 10
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
    
        .Col = 4
        .ColWidth(4) = 10
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
    
        .Col = 5
        .ColWidth(5) = 10
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 6
        .ColWidth(6) = 10
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 7
        .ColWidth(7) = 10
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
         
        .Col = 8
        .ColWidth(8) = 11
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 9
        .ColWidth(9) = 11
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 10
        .ColWidth(10) = 14
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
        
        .Col = 11
        .ColWidth(11) = 8
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 12
        .ColWidth(12) = 40
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
    End With
    
    ReDim sValue(1)
    sValue(0) = 1
    sValue(1) = Trim(txtInput(0).Text)
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01001_A1_ALL", sValue(), Err_Num, Err_Dec)
    
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    Call fpSpread_Display(spdView, RS01)
    Call GetColWidth(REG_App, Me.Name, spdView)
    
    cboSearch.Clear
    cboSearch.AddItem "01: 가맹점 명칭"
    cboSearch.AddItem "02: 대표자명, 점주명"
    cboSearch.AddItem "03: 현재 TAG코드"
    cboSearch.ListIndex = 0
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub spdView_DblClick(ByVal Col As Long, ByVal Row As Long)
    Call spdView_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub spdView_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim nCnt    As Long
    
    If spdView.MaxRows < 1 Then Exit Sub
    
    If KeyCode = vbKeyReturn Then
        
        spdView.Row = spdView.ActiveRow
        spdView.Col = 7
        
        
        Select Case m_FormObj.Name
        
            ' 가맹점 현황
            Case "P_01001"
                For nCnt = 0 To P_01001.cboOffice.ListCount
                    If InStr(P_01001.cboOffice.List(nCnt), Trim(spdView.Text)) > 0 Then
                        
                        spdView.Col = 1
                        P_01001.cboOffice.Tag = spdView.Text
                        
                        P_01001.cboOffice.ListIndex = nCnt
                        
                        Exit For
                    End If
                Next nCnt
        
            Case "P_04027"
                spdView.Col = 1
                P_04027.cmdBtn(8).Tag = spdView.Text
        
            Case "P_06001" ' 사고 처리 접수
                spdView.Col = 1
                P_06001.cmdBtn(8).Tag = spdView.Text
        
        End Select
        
        Unload Me
        
    End If
End Sub

Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    With spdView
        If NewRow <> -1 Then
            .Row = Row
            If (Row Mod 2) = 0 Then
                .Col = -1
                .BackColor = glbGray
            Else
                .Col = -1
                .BackColor = vbWhite
            End If
            
            .Row = NewRow
            .Col = -1
            .BackColor = glbYellow
        End If
    End With
End Sub

Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdSubBtn_Click 0
    End If
End Sub
