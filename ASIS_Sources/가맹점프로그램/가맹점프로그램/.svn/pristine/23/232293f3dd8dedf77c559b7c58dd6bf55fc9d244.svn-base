VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form P_COUPON01 
   ClientHeight    =   7965
   ClientLeft      =   1650
   ClientTop       =   2865
   ClientWidth     =   11850
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
   ScaleHeight     =   7965
   ScaleWidth      =   11850
   WindowState     =   2  '최대화
   Begin Threed.SSPanel SSPanel1 
      Height          =   1125
      Left            =   1890
      TabIndex        =   7
      Top             =   3330
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   1984
      _Version        =   262144
      ForeColor       =   16777215
      BackColor       =   16711680
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      FloodColor      =   16777215
      RoundedCorners  =   0   'False
   End
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   6780
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   900
      Width           =   3525
      _Version        =   524288
      _ExtentX        =   6218
      _ExtentY        =   11959
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowUserFormulas=   -1  'True
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      DAutoCellTypes  =   0   'False
      DAutoHeadings   =   0   'False
      DAutoSave       =   0   'False
      DAutoSizeCols   =   0
      DInformActiveRowChange=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   2
      MaxRows         =   300
      MoveActiveOnFocus=   0   'False
      Protect         =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "P_COUPON01.frx":0000
      VisibleCols     =   2
      VisibleRows     =   50
      AppearanceStyle =   0
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   9765
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Index           =   0
         Left            =   1470
         TabIndex        =   1
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   375
         Index           =   0
         Left            =   2550
         TabIndex        =   2
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##"
         PromptChar      =   "_"
      End
      Begin VB.Label lblSMS 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   4275
         TabIndex        =   10
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "합계"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   13
         Left            =   3750
         TabIndex        =   9
         Top             =   300
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검색 일자"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  '단일 고정
         Caption         =   "월"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2910
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  '단일 고정
         Caption         =   "년"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2190
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
   End
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   6840
      Index           =   1
      Left            =   3660
      TabIndex        =   11
      Top             =   900
      Width           =   8025
      _Version        =   524288
      _ExtentX        =   14155
      _ExtentY        =   12065
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowUserFormulas=   -1  'True
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      DAutoCellTypes  =   0   'False
      DAutoHeadings   =   0   'False
      DAutoSave       =   0   'False
      DAutoSizeCols   =   0
      DInformActiveRowChange=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   5
      MaxRows         =   300
      MoveActiveOnFocus=   0   'False
      Protect         =   0   'False
      SpreadDesigner  =   "P_COUPON01.frx":09C1
      VisibleCols     =   2
      VisibleRows     =   50
      AppearanceStyle =   0
   End
   Begin Threed.SSCommand cmdBtn 
      Height          =   615
      Index           =   0
      Left            =   10020
      TabIndex        =   6
      Top             =   210
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1085
      _Version        =   262144
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "조회"
      ButtonStyle     =   2
   End
End
Attribute VB_Name = "P_COUPON01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Query As String
Dim rs01 As DAO.Recordset

Dim FORM_COUPON01_ACTIVATE    As Boolean
Dim sMasterCode        As String


Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        ' 조회
        Case 0
            Call GetData_View
            Exit Sub
     
        Case Else
        
    End Select
End Sub
 

Private Sub Form_Activate()
    On Error GoTo Error_Rtn

    If FORM_COUPON01_ACTIVATE = True Then Exit Sub
    FORM_COUPON01_ACTIVATE = True
    
    DoEvents
  
    On Error GoTo 0
    Exit Sub

Error_Rtn:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Activate of Form P_SMS001"

End Sub

Private Sub Form_Load()
    
    SSPanel1.Visible = False

    MaskEdBox1(0).Text = Format(Date, "yyyy")
    MaskEdBox2(0).Text = Format(Date, "mm")
    
    TitleSet "쿠폰 접수 현황"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    FORM_COUPON01_ACTIVATE = False
End Sub
 
Private Sub MaskEdBox1_GotFocus(Index As Integer)
    MaskEdBox1(Index).SelStart = 0
    MaskEdBox1(Index).SelLength = Len(MaskEdBox1(Index).Text)
End Sub

Private Sub MaskEdBox2_GotFocus(Index As Integer)
    MaskEdBox2(Index).SelStart = 0
    MaskEdBox2(Index).SelLength = Len(MaskEdBox2(Index).Text)
End Sub
 

Private Sub DataTotal()
    Dim lRow    As Long
    Dim varTemp As Variant
    Dim LCount    As Long
    
    LCount = 0
    For lRow = 1 To fpSpread1(0).MaxRows
        Call fpSpread1(0).GetText(2, lRow, varTemp)
        LCount = LCount + Val(Replace(CStr(varTemp), ",", ""))
    Next lRow
    
    lblSMS(0).Caption = Format(LCount, "#,##0")

End Sub

Private Sub fpSpread1_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
    Dim varTemp As Variant
    
    ' 좌측 그리드를 클릭한 경우 해당 일자의 세부 내역을 조회 한다.
    If Index = 0 Then
        Call fpSpread1(0).GetText(1, Row, varTemp)
        If IsDate(Format(CStr(varTemp), "@@@@-@@-@@")) = True Then
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
    Dim rs01 As Recordset
    Dim Query    As String
    Dim lRow    As Long
    
    On Error GoTo GetData_View_Error
    
    
    Screen.MousePointer = vbHourglass
    
    Query = "SELECT 접수일자, Count(쿠폰번호) AS Cnt "
    Query = Query & " FROM 쿠폰자료 "
    Query = Query & " WHERE 대리점코드 = '" & 대리점정보.StoreCode & "' "
    Query = Query & "   AND LEFT(접수일자,6) = '" & MaskEdBox1(0).Text & MaskEdBox2(0).Text & "'  "
    Query = Query & " GROUP BY 접수일자 "
    Query = Query & " ORDER BY 접수일자 "
    Set rs01 = MyDB.OpenRecordset(Query)
    
    If rs01.EOF Then
        Screen.MousePointer = vbDefault
        MsgBox "해당자료가 없습니다."
        Exit Sub
    End If
    
    
    While Not rs01.EOF

        If fpSpread1(0).MaxRows = lRow Then
            fpSpread1(0).MaxRows = fpSpread1(0).MaxRows + 1
            fpSpread1(0).RowHeight(fpSpread1(0).MaxRows) = 20
        End If
                    
        lRow = lRow + 1
        fpSpread1(0).SetText 1, lRow, rs01.Fields(0) & ""
        fpSpread1(0).SetText 2, lRow, rs01.Fields(1) & ""
        
        rs01.MoveNext
    Wend
    
    rs01.Close
    
    ' 합계 출력
    Call DataTotal
    Screen.MousePointer = vbDefault
    
    On Error GoTo 0
    Exit Sub

GetData_View_Error:

    Screen.MousePointer = vbDefault
    Set rs01 = Nothing
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetData_View of Form P_COUPON01"
End Sub

'--------------------------------------------------------------------------------------------------------------
' Procedure : GetData_ViewDetailed
' DateTime  : 2007-01-08 22:39
' Author    : pds2004
' Purpose   :
'--------------------------------------------------------------------------------------------------------------
Private Sub GetData_ViewDetailed(ByVal sDate As String)
    Dim Query    As String
    Dim rs01 As Recordset
    Dim bResult As Boolean
    Dim lRow    As Long
    
    On Error GoTo GetData_View_Error

    
    Screen.MousePointer = vbHourglass
    fpSpread1(1).MaxRows = 0
    
    sDate = Replace(Replace(sDate, "-", ""), "/", "")

    Query = "SELECT 쿠폰번호, 쿠폰금액, 고객번호, 고객이름, 접수금액 "
    Query = Query & " FROM 쿠폰자료 "
    Query = Query & " WHERE 대리점코드 = '" & 대리점정보.StoreCode & "' "
    Query = Query & "   AND 접수일자= '" & sDate & "'  "
    Query = Query & " ORDER BY 쿠폰번호 "
    Set rs01 = MyDB.OpenRecordset(Query)
    
    If rs01.EOF Then
        Screen.MousePointer = vbDefault
        MsgBox "해당자료가 없습니다."
        Exit Sub
    End If
    
    While Not rs01.EOF

        If fpSpread1(1).MaxRows = lRow Then
            fpSpread1(1).MaxRows = fpSpread1(1).MaxRows + 1
            fpSpread1(1).RowHeight(fpSpread1(1).MaxRows) = 20
        End If
                    
        lRow = lRow + 1
        fpSpread1(1).SetText 1, lRow, rs01.Fields(0) & ""
        fpSpread1(1).SetText 2, lRow, Format(rs01.Fields(1) & "", "#,##0")
        fpSpread1(1).SetText 3, lRow, rs01.Fields(2) & ""
        fpSpread1(1).SetText 4, lRow, rs01.Fields(3) & ""
        fpSpread1(1).SetText 5, lRow, Format(rs01.Fields(4) & "", "#,##0")
        
        rs01.MoveNext
    Wend

    
    rs01.Close
    Screen.MousePointer = vbDefault
    
    On Error GoTo 0
    Exit Sub

GetData_View_Error:
    Screen.MousePointer = vbDefault
    Set rs01 = Nothing
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetData_ViewDetailed of Form P_COUPON01"
End Sub

 
