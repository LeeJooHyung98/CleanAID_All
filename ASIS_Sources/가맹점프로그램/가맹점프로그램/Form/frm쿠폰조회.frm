VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm쿠폰조회 
   Caption         =   "쿠폰 조회"
   ClientHeight    =   11835
   ClientLeft      =   1650
   ClientTop       =   3255
   ClientWidth     =   16230
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11835
   ScaleWidth      =   16230
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   11835
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16230
      _ExtentX        =   28628
      _ExtentY        =   20876
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm쿠폰조회.frx":0000
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   10605
         Index           =   0
         Left            =   15
         TabIndex        =   1
         Top             =   1215
         Width           =   3540
         _Version        =   524288
         _ExtentX        =   6244
         _ExtentY        =   18706
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
         MaxCols         =   2
         MaxRows         =   300
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frm쿠폰조회.frx":0092
         VisibleCols     =   2
         VisibleRows     =   50
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   10605
         Index           =   1
         Left            =   3570
         TabIndex        =   2
         Top             =   1215
         Width           =   12645
         _Version        =   524288
         _ExtentX        =   22304
         _ExtentY        =   18706
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
         MaxCols         =   5
         MaxRows         =   300
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         SpreadDesigner  =   "frm쿠폰조회.frx":0613
         VisibleCols     =   2
         VisibleRows     =   50
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   750
         Left            =   15
         TabIndex        =   3
         Top             =   450
         Width           =   16200
         _ExtentX        =   28575
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdList 
            Height          =   630
            Left            =   3030
            TabIndex        =   6
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 조회(&F)"
            Appearance      =   6
            Picture         =   "frm쿠폰조회.frx":0C69
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   3
            Left            =   4575
            TabIndex        =   7
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 엑셀(&E)"
            Appearance      =   6
            Picture         =   "frm쿠폰조회.frx":1363
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   7860
            TabIndex        =   8
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frm쿠폰조회.frx":1ADD
         End
         Begin XtremeSuiteControls.PushButton cmdPrint 
            Height          =   630
            Left            =   6120
            TabIndex        =   9
            Top             =   60
            Width           =   1695
            _Version        =   851970
            _ExtentX        =   2990
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 출력(&P)"
            Appearance      =   6
            Picture         =   "frm쿠폰조회.frx":2B6F
         End
         Begin MSComCtl2.DTPicker dtpDay 
            Height          =   360
            Left            =   915
            TabIndex        =   10
            Top             =   45
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   635
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
            CustomFormat    =   "yyyy-MM"
            Format          =   55771139
            UpDown          =   -1  'True
            CurrentDate     =   40279
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "접수일자:"
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
            TabIndex        =   5
            Top             =   120
            Width           =   840
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   4
         Top             =   15
         Width           =   16200
         _ExtentX        =   28575
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
         Caption         =   "      쿠폰 조회"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm쿠폰조회.frx":3269
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm쿠폰조회.frx":348F
            Top             =   -15
            Width           =   765
         End
      End
   End
End
Attribute VB_Name = "frm쿠폰조회"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FORM_COUPON01_ACTIVATE    As Boolean
Dim sMasterCode        As String


Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 3: Call Export_Excel(frmMain.cdgExcel, sprGrid)
        Case 5: Unload Me
    End Select
End Sub
 

Private Sub cmdList_Click()
    Call GetData_View
End Sub

Private Sub Form_Activate()
    On Error GoTo ErrRtn

    If FORM_COUPON01_ACTIVATE = True Then Exit Sub
    FORM_COUPON01_ACTIVATE = True
    
    DoEvents
  
    On Error GoTo 0
    Exit Sub

ErrRtn:
    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure Form_Activate of Form P_SMS001"

End Sub

Private Sub Form_Load()
    
    For i = 0 To 1
        With sprGrid(i)
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
    Next i
    
    dtpDay.Value = Format(Date, "YYYY-MM")
    
    'TitleSet "쿠폰 접수 현황"
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    
    cmdBtn(5).Left = (Me.Width - 200) - cmdBtn(5).Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FORM_COUPON01_ACTIVATE = False
End Sub

Private Sub sprGrid_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
    Dim varTemp As Variant
    
    ' 좌측 그리드를 클릭한 경우 해당 일자의 세부 내역을 조회 한다.
    If Index = 0 Then
        Call sprGrid(0).GetText(1, Row, varTemp)
        
        If IsDate(Format(CStr(varTemp), "YYYY-MM-DD")) = True Then
            Call GetData_ViewDetailed(CStr(varTemp))
        End If
    End If
End Sub



'--------------------------------------------------------------------------------------------------------------
' Procedure : GetData1
' DateTime  : 2007-01-08 22:39
' Author    : pds2004
' Purpose   :
'--------------------------------------------------------------------------------------------------------------
Private Sub GetData_View()
    Dim lRow    As Long
    
    On Error GoTo GetData_View_Error
    
    
    Screen.MousePointer = vbHourglass
    
    Query = "SELECT    접수일자"
    Query = Query & ", COUNT(쿠폰번호) AS CNT"
    Query = Query & " FROM TB_쿠폰자료"
    Query = Query & " WHERE 가맹점코드 = '" & 가맹점정보.가맹점코드 & "' "
    Query = Query & "   AND SUBSTRING(접수일자,1,7) = '" & Format(dtpDay.Value, "YYYY-MM") & "'"
    Query = Query & " GROUP BY 접수일자"
    Query = Query & " ORDER BY 접수일자 ASC"
    Set SUBRs = New ADODB.RecordSet
    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
        
    With sprGrid(0)
        .MaxRows = 0
        .ReDraw = False
        
        Do Until SUBRs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = Format(SUBRs(0), "YYYY-MM-DD") & ""
            .Col = 2: .Text = SUBRs(1) & ""
            
            SUBRs.MoveNext
        Loop
        SUBRs.Close
        Set SUBRs = Nothing
        
        .ReDraw = True
    End With
    
    Screen.MousePointer = vbDefault
    
    Exit Sub

GetData_View_Error:

    Screen.MousePointer = vbDefault
    
    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure GetData_View of Form P_COUPON01"
End Sub

'--------------------------------------------------------------------------------------------------------------
' Procedure : GetData_ViewDetailed
' DateTime  : 2007-01-08 22:39
' Author    : pds2004
' Purpose   :
'--------------------------------------------------------------------------------------------------------------
Private Sub GetData_ViewDetailed(ByVal sDate As String)
    Dim bResult As Boolean
    Dim lRow    As Long
    
    On Error GoTo GetData_View_Error
    
    Screen.MousePointer = vbHourglass
    
    Query = "SELECT 쿠폰번호, 쿠폰금액, 고객코드, 고객이름, 접수금액 "
    Query = Query & " FROM TB_쿠폰자료 "
    Query = Query & " WHERE 가맹점코드 = '" & 가맹점정보.가맹점코드 & "' "
    Query = Query & "   AND 접수일자   = '" & sDate & "'  "
    Query = Query & " ORDER BY 쿠폰번호 "
    Set SUBRs = New ADODB.RecordSet
    SUBRs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    With sprGrid(1)
        .MaxRows = 0
        .ReDraw = False
        
        Do Until SUBRs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Text = SUBRs(0) & ""
            .Col = 2: .Text = SUBRs(1) & ""
            .Col = 3: .Text = SUBRs(2) & ""
            .Col = 4: .Text = SUBRs(3) & ""
            .Col = 5: .Text = SUBRs(4) & ""
            
            SUBRs.MoveNext
        Loop
        SUBRs.Close
        Set SUBRs = Nothing
        
        .ReDraw = True
    End With
    
    Screen.MousePointer = vbDefault
    
    Exit Sub

GetData_View_Error:
    Screen.MousePointer = vbDefault
    Set SUBRs = Nothing
    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure GetData_ViewDetailed of Form P_COUPON01"
End Sub

 
