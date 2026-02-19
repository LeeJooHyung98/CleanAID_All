VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form P_04011_A_NEW 
   Caption         =   "[전사업장]점별 기간별 매출현황"
   ClientHeight    =   10215
   ClientLeft      =   960
   ClientTop       =   3570
   ClientWidth     =   15720
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_04011_A_NEW.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15720
   WindowState     =   2  '최대화
   Begin Threed.SSPanel panMain 
      Align           =   1  '위 맞춤
      Height          =   9735
      Left            =   0
      TabIndex        =   0
      Top             =   435
      Width           =   15720
      _ExtentX        =   27728
      _ExtentY        =   17171
      _Version        =   262144
      RoundedCorners  =   0   'False
      Begin VB.TextBox txtInput 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Index           =   11
         Left            =   4440
         TabIndex        =   20
         Top             =   8760
         Width           =   1515
      End
      Begin VB.TextBox txtInput 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Index           =   10
         Left            =   1440
         TabIndex        =   19
         Top             =   8760
         Width           =   1515
      End
      Begin VB.TextBox txtInput 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Index           =   9
         Left            =   12990
         TabIndex        =   18
         Top             =   8760
         Width           =   1065
      End
      Begin VB.TextBox txtInput 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Index           =   8
         Left            =   12990
         TabIndex        =   17
         Top             =   8400
         Width           =   1065
      End
      Begin VB.TextBox txtInput 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Index           =   7
         Left            =   7440
         TabIndex        =   16
         Top             =   8760
         Width           =   1515
      End
      Begin VB.TextBox txtInput 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Index           =   6
         Left            =   10440
         TabIndex        =   15
         Top             =   8760
         Width           =   1065
      End
      Begin VB.TextBox txtInput 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Index           =   5
         Left            =   10440
         TabIndex        =   14
         Top             =   9120
         Width           =   1065
      End
      Begin VB.TextBox txtInput 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Index           =   4
         Left            =   7440
         TabIndex        =   13
         Top             =   9120
         Width           =   1515
      End
      Begin VB.TextBox txtInput 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Index           =   0
         Left            =   1440
         TabIndex        =   12
         Top             =   8400
         Width           =   1515
      End
      Begin VB.TextBox txtInput 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Index           =   1
         Left            =   4440
         TabIndex        =   11
         Top             =   8400
         Width           =   1515
      End
      Begin VB.TextBox txtInput 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Index           =   3
         Left            =   10440
         TabIndex        =   10
         Top             =   8400
         Width           =   1065
      End
      Begin VB.TextBox txtInput 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Index           =   2
         Left            =   7440
         TabIndex        =   9
         Top             =   8400
         Width           =   1515
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   8205
         Left            =   60
         TabIndex        =   1
         Top             =   90
         Width           =   14925
         _Version        =   524288
         _ExtentX        =   26326
         _ExtentY        =   14473
         _StockProps     =   64
         BackColorStyle  =   1
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "P_04011_A_NEW.frx":058A
         AppearanceStyle =   0
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   1
         Left            =   60
         TabIndex        =   21
         Top             =   8400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   262144
         BackColor       =   12648384
         Caption         =   "전체매출액"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   3
         Left            =   3060
         TabIndex        =   22
         Top             =   8400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   262144
         BackColor       =   12648384
         Caption         =   "사업장매출액"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   4
         Left            =   9060
         TabIndex        =   23
         Top             =   8400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   262144
         BackColor       =   16761024
         Caption         =   "입고  수량"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   5
         Left            =   6060
         TabIndex        =   24
         Top             =   8400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   262144
         BackColor       =   12648384
         Caption         =   "가맹점매출액"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   6
         Left            =   6060
         TabIndex        =   25
         Top             =   9120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   262144
         BackColor       =   12648384
         Caption         =   "카드  금액"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   7
         Left            =   6060
         TabIndex        =   26
         Top             =   8760
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   262144
         BackColor       =   12648384
         Caption         =   "수선  금액"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   8
         Left            =   9060
         TabIndex        =   27
         Top             =   9120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   262144
         BackColor       =   16761024
         Caption         =   "카드  건수"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   9
         Left            =   9060
         TabIndex        =   28
         Top             =   8760
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   262144
         BackColor       =   16761024
         Caption         =   "수선  수량"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   10
         Left            =   11610
         TabIndex        =   29
         Top             =   8400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   262144
         BackColor       =   16761024
         Caption         =   "반품  수량"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   11
         Left            =   11610
         TabIndex        =   30
         Top             =   8760
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   262144
         BackColor       =   16761024
         Caption         =   "재세탁수량"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   12
         Left            =   60
         TabIndex        =   31
         Top             =   8760
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   262144
         BackColor       =   12648384
         Caption         =   "전체 단가"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   13
         Left            =   3060
         TabIndex        =   32
         Top             =   8760
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   262144
         BackColor       =   12648384
         Caption         =   "사업장 단가"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
   End
   Begin Threed.SSPanel panInput 
      Align           =   1  '위 맞춤
      Height          =   435
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   15720
      _ExtentX        =   27728
      _ExtentY        =   767
      _Version        =   262144
      RoundedCorners  =   0   'False
      Begin MSComCtl2.DTPicker dtInput 
         Height          =   315
         Index           =   1
         Left            =   3660
         TabIndex        =   33
         Top             =   60
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Format          =   21430275
         UpDown          =   -1  'True
         CurrentDate     =   37140
      End
      Begin VB.ComboBox cboInput 
         Height          =   315
         Index           =   1
         Left            =   6750
         Style           =   2  '드롭다운 목록
         TabIndex        =   4
         Top             =   60
         Visible         =   0   'False
         Width           =   2805
      End
      Begin VB.ComboBox cboInput 
         Height          =   315
         Index           =   0
         Left            =   11070
         Style           =   2  '드롭다운 목록
         TabIndex        =   3
         Top             =   60
         Width           =   3075
      End
      Begin MSComCtl2.DTPicker dtInput 
         Height          =   315
         Index           =   0
         Left            =   1740
         TabIndex        =   5
         Top             =   60
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Format          =   21430275
         UpDown          =   -1  'True
         CurrentDate     =   37140
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   2
         Left            =   60
         TabIndex        =   6
         Top             =   60
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "수 금 년 월"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   0
         Left            =   9600
         TabIndex        =   7
         Top             =   60
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "가 맹 점"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   35
         Left            =   5370
         TabIndex        =   8
         Top             =   60
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "사 업 장"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "~"
         Height          =   195
         Left            =   3420
         TabIndex        =   34
         Top             =   120
         Width           =   105
      End
   End
End
Attribute VB_Name = "P_04011_A_NEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cboInput_Click(Index As Integer)
    Dim sCode As String

    If Index = 1 Then
        sCode = Trim(Mid(Trim(cboInput(1)) & Space(10), 2, 4))

        Call Get_가맹점리스트(cboInput(0), sCode)
    End If
End Sub


Private Sub Form_Activate()
'    cmdBtn(0).Enabled = True
'    cmdBtn(2).Enabled = True
'    cmdBtn(5).Enabled = True
'    cmdBtn(6).Enabled = True
'
'    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    

End Sub

Private Sub spdDisplay(Rs As ADODB.Recordset)
    
    
    Call fpSpread_Display(spdView, Rs)
    
    spdView.ColsFrozen = 2 '틀고정
    
    spdView.Row = -1
    
    spdView.Col = 1
    spdView.ColWidth(1) = 10
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft

    spdView.Col = 2
    spdView.ColWidth(2) = 5
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter

    spdView.Col = 3
    spdView.ColWidth(3) = 18
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft

    spdView.Col = 4
    spdView.ColWidth(4) = 6
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter
    
    spdView.Col = 5
    spdView.ColWidth(5) = 6
    spdView.CellType = CellTypePic
    spdView.TypePicMask = "9-999"
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter

    spdView.Col = 6
    spdView.ColWidth(6) = 6
    spdView.CellType = CellTypePic
    spdView.TypePicMask = "9-999"
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter
    
    spdView.Col = 7
    spdView.ColWidth(7) = 5
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter

    spdView.Col = 8
    spdView.ColWidth(8) = 10
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 9
    spdView.ColWidth(9) = 10
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 10
    spdView.ColWidth(10) = 10
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 11
    spdView.ColWidth(11) = 5
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 12
    spdView.ColWidth(12) = 5
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 13
    spdView.ColWidth(13) = 6
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
    
    spdView.Col = 14
    spdView.ColWidth(14) = 6
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
    
    spdView.Col = 15
    spdView.ColWidth(15) = 10
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
    
    spdView.Col = 16
    spdView.ColWidth(16) = 5
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 17
    spdView.ColWidth(17) = 6
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
    
    spdView.Col = 18
    spdView.ColWidth(18) = 5
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
    
    spdView.Col = 19
    spdView.ColWidth(19) = 10
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
    
    spdView.Col = 20
    spdView.ColWidth(20) = 5
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter

    spdView.Col = 21
    spdView.ColWidth(21) = 5
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
    
    spdView.Col = 22
    spdView.ColWidth(22) = 12
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
     
    spdView.Col = 23
    spdView.ColWidth(23) = 12
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
     
    spdView.Col = 24
    spdView.ColWidth(24) = 12
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
    
    spdView.Col = 25
    spdView.ColWidth(25) = 5
    spdView.CellType = CellTypeCheckBox
    spdView.TypeCheckCenter = True
    spdView.Value = False
    spdView.TypeHAlign = TypeHAlignCenter
    spdView.TypeVAlign = TypeVAlignCenter
    
    spdView.Col = 26
    spdView.ColWidth(26) = 8
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter

    
'    spdView.ColHidden = True
    
'    Set spdView.DataSource = Nothing
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    panCaption(35).Visible = True
    cboInput(1).Visible = True
    Call Master_tblComboAdd(cboInput(1))
    
    Call Get_가맹점리스트(cboInput(0), Trim(Mid(Trim(cboInput(1)) & Space(10), 2, 4)))
    
    
    dtInput(0).Value = Format(Date, "yyyy-mm")
    dtInput(1).Value = Format(Date, "yyyy-mm")
    
    ReDim sValue(3)
    sValue(0) = "1"
    sValue(1) = Format(dtInput(0).Value, "yyyymm")
    sValue(2) = Format(dtInput(1).Value, "yyyymm")
    sValue(3) = Mid(cboInput(0).Text, 2, 6)
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_04011_00_ALL2", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    Call spdDisplay(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView)
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_04011_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim sCode   As String
    Dim i As Integer
    Dim ii As Integer
    
    sCode = Mid(Trim(cboInput(1).Text) & Space(10), 2, 4)
    
    If Trim(cboInput(1).Text) = "" Then
        MsgBox "사업장을 선택하십시오.", vbInformation
        cboInput(1).SetFocus
        Exit Sub
    End If
    
    If Trim(cboInput(0).Text) = "" Then
        MsgBox "가맹점을 선택하십시오.", vbInformation
        cboInput(0).SetFocus
        Exit Sub
    End If
        
    Set RS01 = New ADODB.Recordset
    

    ReDim sValue(3)
    sValue(0) = "0"
    sValue(1) = Format(dtInput(0).Value, "yyyymm")
    sValue(2) = Format(dtInput(1).Value, "yyyymm")
    sValue(3) = Mid(cboInput(0).Text, 2, 6)
    Set RS01 = ExecPro("SP_04011_00_ALL2", sValue(), Err_Num, Err_Dec)

    
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    Call spdDisplay(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView)

    spdView.MaxRows = RS01.RecordCount

    For i = 1 To RS01.RecordCount
        spdView.Row = i
        
        spdView.Col = 1
        spdView.Text = Format(RS01.Fields("매출일자"), "@@@@-@@-@@")
        spdView.Col = 2
        spdView.Text = ExecWeekDay(Format(RS01.Fields("매출일자"), "@@@@-@@-@@"))
        
        If spdView.Text = "일" Then
            'spdView.Row = NewRow
            spdView.Col = -1
            spdView.BackColor = vbYellow
        End If
        
        spdView.Col = 7
        Select Case spdView.Text
            Case "1"
                spdView.Text = spdView.Text + ":세일"
            Case "2"
                spdView.Text = spdView.Text + ":목요"
            Case "3"
                spdView.Text = spdView.Text + ":정상"
        End Select
        Dim j As Integer
        
        For j = 8 To 24
            spdView.Col = j
            spdView.Text = "0"
        Next j
        RS01.MoveNext
    Next i

    RS01.MoveFirst
    Do While Not RS01.EOF
        For i = 1 To spdView.MaxRows
            spdView.Row = i
            spdView.Col = 1
            If Format(spdView.Text, "YYYY-MM-DD") = RS01(0) Then
            
                For j = 3 To spdView.MaxCols
                    If j = 7 Then
                        spdView.Col = j
                        Select Case RS01(j - 1)
                            Case "1"
                                spdView.Text = RS01(j - 1) + ":세일"
                            Case "2"
                                spdView.Text = RS01(j - 1) + ":목요"
                            Case "3"
                                spdView.Text = RS01(j - 1) + ":정상"
                        End Select
                    Else
                        If j = 25 Then
                            spdView.Col = j
                            Select Case RS01(j - 1)
                                Case "Y"
                                    spdView.Text = True
                                Case Else
                                    spdView.Text = False
                            End Select
                        Else
                            spdView.Col = j
                            If IsNull(RS01(j - 1)) Then
                                spdView.Text = ""
                            Else
                                spdView.Text = RS01(j - 1)
                            End If
                        End If
                    End If
                Next j
            End If

        Next i

        RS01.MoveNext
    Loop
        
    spdView.MaxRows = spdView.MaxRows + 1
    spdView.Row = spdView.MaxRows
    spdView.RowHidden = True
    
    spdView.Col = 1
    spdView.Text = "합계"
    
    Dim cnt, Tamt, Mamt, Samt As Long
    
    
    spdView.Col = 8
    spdView.Formula = "SUM(H1:H" & spdView.MaxRows - 1 & ")"
    Tamt = spdView.Value
    txtInput(0).Text = spdView.Text
    spdView.Col = 9
    spdView.Formula = "SUM(I1:I" & spdView.MaxRows - 1 & ")"
    Mamt = spdView.Value
    txtInput(1).Text = spdView.Text
    
    spdView.Col = 10
    spdView.Formula = "SUM(J1:J" & spdView.MaxRows - 1 & ")"
    Samt = spdView.Value
    txtInput(2).Text = spdView.Text
    
    spdView.Col = 11
    spdView.Formula = "SUM(K1:K" & spdView.MaxRows - 1 & ")"
    cnt = spdView.Value
    txtInput(3).Text = spdView.Text
    
    If cnt = 0 Then
        spdView.Col = 12
        spdView.Text = 0
        
        spdView.Col = 13
        spdView.Text = 0
        
        spdView.Col = 14
        spdView.Text = 0
    Else
        spdView.Col = 12
        spdView.Text = Tamt / cnt
        
        spdView.Col = 13
        spdView.Text = Mamt / cnt
        
        spdView.Col = 14
        spdView.Text = Samt / cnt
    End If

    spdView.Col = 15
    spdView.Formula = "SUM(O1:O" & spdView.MaxRows - 1 & ")"
    txtInput(4).Text = spdView.Text
    
    spdView.Col = 16
    spdView.Formula = "SUM(P1:P" & spdView.MaxRows - 1 & ")"
    txtInput(5).Text = spdView.Text
    
    spdView.Col = 17
    spdView.Formula = "SUM(Q1:Q" & spdView.MaxRows - 1 & ")"
    txtInput(9).Text = spdView.Text
    
    spdView.Col = 18
    spdView.Formula = "SUM(R1:R" & spdView.MaxRows - 1 & ")"
    txtInput(7).Text = spdView.Text
    
    spdView.Col = 19
    spdView.Formula = "SUM(S1:S" & spdView.MaxRows - 1 & ")"
    txtInput(6).Text = spdView.Text
    
    spdView.Col = 20
    spdView.Formula = "SUM(T1:T" & spdView.MaxRows - 1 & ")"
    txtInput(8).Text = spdView.Text
    
    spdView.Col = 21
    spdView.Formula = "SUM(U1:U" & spdView.MaxRows - 1 & ")"
    spdView.Col = 22
    spdView.Formula = "SUM(V1:V" & spdView.MaxRows - 1 & ")"
    spdView.Col = 23
    spdView.Formula = "SUM(W1:W" & spdView.MaxRows - 1 & ")"
    spdView.Col = 24
    spdView.Formula = "SUM(X1:X" & spdView.MaxRows - 1 & ")"
    
    spdView.MaxRows = spdView.MaxRows - 1
    
    If txtInput(3).Text = 0 Then
        txtInput(10).Text = 0
        txtInput(11).Text = 0
    Else
        txtInput(10).Text = Format(txtInput(0).Text / txtInput(3).Text, "#,##0")
        txtInput(11).Text = Format(txtInput(1).Text / txtInput(3).Text, "#,##0")
    End If
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub



Private Sub spdView_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If NewRow <> -1 Then
        With spdView
            If NewRow <> -1 Then
                .Row = Row
                .Col = 2
                If spdView.Text = "일" Then
                    .Col = -1
                    .BackColor = vbYellow
                Else
                    If (Row Mod 2) = 0 Then
                        .Col = -1
                        .BackColor = glbGray
                    Else
                        .Col = -1
                        .BackColor = vbWhite
                    End If
                End If
                .Row = NewRow
                .Col = -1
                .BackColor = glbYellow
            End If
        End With
    End If
End Sub

'Private Sub spdView_Change(ByVal Col As Long, ByVal Row As Long)
'    Select Case Col
'        Case 2
'            spdView.Row = Row
'
'            'spdView.Col = 14
'
'            If spdView.Text = "일" Then
'                spdView.Col = -1
'                spdView.BackColor = vbYellow
'            End If
'    End Select
'End Sub


Private Sub spdView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        'PopupMenu P_00000.PopUp
    End If
End Sub

Public Sub DataPrint()
'    Dim ReportFP As String
'    Dim ReportFile As String
'    Dim sData As String
'
'    ReportFP = GetIniStr("REPORT", "FilePath", "", m_iniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    Call PrintDesc
'
'    Dim i As Integer
'    For i = 0 To 30
'        P_00000.crPrint.Formulas(i) = ""
'    Next
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'    P_00000.crPrint.Formulas(0) = "월 = '" & Format(dtInput(0).Value, "yyyy-mm") & "'"
'    P_00000.crPrint.Formulas(1) = "대리점 = '" & Trim(cboInput(0).Text) & "'"
'
'
'    sData = Space(2) & LeftH("월  합  계" & Space(12), 12)
'    sData = sData & RightH(Space(11) & txtInput(0).Text, 11) & Space(2)
'    If txtInput(3).Text = 0 Then
'        sData = sData & RightH(Space(11) & Format(0, "#,##0"), 5) & Space(2)
'    Else
'        sData = sData & RightH(Space(11) & Format(txtInput(0).Text / txtInput(3).Text, "#,##0"), 5) & Space(2)
'    End If
'    sData = sData & RightH(Space(11) & txtInput(1).Text, 11) & Space(2)
'    If txtInput(3).Text = 0 Then
'        sData = sData & RightH(Space(11) & Format(0, "#,##0"), 5) & Space(2)
'    Else
'        sData = sData & RightH(Space(11) & Format(txtInput(1).Text / txtInput(3).Text, "#,##0"), 5) & Space(2)
'    End If
'    sData = sData & RightH(Space(11) & txtInput(2).Text, 11) & Space(2)
'    sData = sData & RightH(Space(11) & txtInput(3).Text, 6) & Space(2)
'    sData = sData & RightH(Space(11) & txtInput(7).Text, 5) & Space(2)
'    sData = sData & RightH(Space(11) & txtInput(8).Text, 5) & Space(2)
'    sData = sData & RightH(Space(11) & txtInput(6).Text, 5) & Space(2)
'
'    sData = sData & RightH(Space(11) & txtInput(4).Text, 11) & Space(1)
'    sData = sData & RightH(Space(11) & txtInput(5).Text, 6)
'
'    P_00000.crPrint.Formulas(2) = "합계 = '" & sData & "'"
'    P_00000.crPrint.Formulas(3) = "사업장 = '" & Trim(cboInput(1).Text) & "'"
'
'    Set RS01 = New ADODB.Recordset
'    ReDim sValue(0)
'    sValue(0) = "0"
'    Set RS01 = ExecPro("SP_A_0000", sValue(), Err_Num, Err_Dec)
'    P_00000.crPrint.Formulas(4) = "출력시간 = '" & RS01!DB_DATE & "'"
'
'    Call ReportPrint(ReportFile, "1")
End Sub

Public Sub DataScreen()
'    Dim ReportFP As String
'    Dim ReportFile As String
'    Dim sData As String
'
'    ReportFP = GetIniStr("REPORT", "FilePath", "", m_iniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    Call PrintDesc
'
'    Dim i As Integer
'    For i = 0 To 30
'        P_00000.crPrint.Formulas(i) = ""
'    Next
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'    P_00000.crPrint.Formulas(0) = "월 = '" & Format(dtInput(0).Value, "yyyy-mm") & "'"
'    P_00000.crPrint.Formulas(1) = "대리점 = '" & Trim(cboInput(0).Text) & "'"
'
'
'    sData = Space(2) & LeftH("월  합  계" & Space(12), 12)
'    sData = sData & RightH(Space(11) & txtInput(0).Text, 11) & Space(2)
'    If txtInput(3).Text = 0 Then
'        sData = sData & RightH(Space(11) & Format(0, "#,##0"), 5) & Space(2)
'    Else
'        sData = sData & RightH(Space(11) & Format(txtInput(0).Text / txtInput(3).Text, "#,##0"), 5) & Space(2)
'    End If
'    sData = sData & RightH(Space(11) & txtInput(1).Text, 11) & Space(2)
'    If txtInput(3).Text = 0 Then
'        sData = sData & RightH(Space(11) & Format(0, "#,##0"), 5) & Space(2)
'    Else
'        sData = sData & RightH(Space(11) & Format(txtInput(1).Text / txtInput(3).Text, "#,##0"), 5) & Space(2)
'    End If
'    sData = sData & RightH(Space(11) & txtInput(2).Text, 11) & Space(2)
'    sData = sData & RightH(Space(11) & txtInput(3).Text, 6) & Space(2)
'    sData = sData & RightH(Space(11) & txtInput(7).Text, 5) & Space(2)
'    sData = sData & RightH(Space(11) & txtInput(8).Text, 5) & Space(2)
'    sData = sData & RightH(Space(11) & txtInput(6).Text, 5) & Space(2)
'
'    sData = sData & RightH(Space(11) & txtInput(4).Text, 11) & Space(1)
'    sData = sData & RightH(Space(11) & txtInput(5).Text, 6)
'
'    P_00000.crPrint.Formulas(2) = "합계 = '" & sData & "'"
'    P_00000.crPrint.Formulas(3) = "사업장 = '" & Trim(cboInput(1).Text) & "'"
'
'    Set RS01 = New ADODB.Recordset
'    ReDim sValue(0)
'    sValue(0) = "0"
'    Set RS01 = ExecPro("SP_A_0000", sValue(), Err_Num, Err_Dec)
'    P_00000.crPrint.Formulas(4) = "출력시간 = '" & RS01!DB_DATE & "'"
'
'    Call ReportPrint(ReportFile, "2")
End Sub

Private Sub PrintDesc()
    Dim i As Integer
    Dim TempText As String
    Dim TempFP As String
    Dim TempFile As String
    
    TempFP = GetIniStr("REPORT", "TempPath", "", m_iniFile)
    TempFile = TempFP & "\Temp.txt"
    
    Open TempFile For Output As #1
    
    TempText = ""
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        
        spdView.Col = 1
        TempText = LeftH(RightH(spdView.Text, 5) & Space(5), 6)
        spdView.Col = 2
        TempText = TempText & LeftH(spdView.Text & Space(4), 4)

        spdView.Col = 26
        If Trim(spdView.Text) = "" Then
            TempText = TempText & LeftH("N" & Space(4), 4)
        Else
            TempText = TempText & LeftH("Y" & Space(4), 4)
        End If
        
        spdView.Col = 8
        TempText = TempText & RightH(Space(11) & spdView.Text, 11) & Space(2)
        spdView.Col = 12
        TempText = TempText & RightH(Space(5) & spdView.Text, 5) & Space(2)
        spdView.Col = 9
        TempText = TempText & RightH(Space(11) & spdView.Text, 11) & Space(2)
        spdView.Col = 13
        TempText = TempText & RightH(Space(5) & spdView.Text, 5) & Space(2)
        spdView.Col = 10
        TempText = TempText & RightH(Space(11) & spdView.Text, 11) & Space(2)
        spdView.Col = 11
        TempText = TempText & RightH(Space(6) & spdView.Text, 6) & Space(2)
        spdView.Col = 19
        TempText = TempText & RightH(Space(5) & spdView.Text, 5) & Space(2)
        spdView.Col = 17
        TempText = TempText & RightH(Space(5) & spdView.Text, 5) & Space(2)
        spdView.Col = 18
        TempText = TempText & RightH(Space(5) & spdView.Text, 5) & Space(2)
        spdView.Col = 15
        TempText = TempText & RightH(Space(11) & spdView.Text, 11) & Space(1)
        spdView.Col = 16
        TempText = TempText & RightH(Space(6) & spdView.Text, 6) & Space(1)
        
        Print #1, TempText
    Next i
    
    Close #1
End Sub
Public Sub DataSave()
    Dim i As Integer
    Dim sCode   As String
    Dim bChk As Boolean
    'ReDim sValue(2)
    
    On Error GoTo ERR_RTN
        
    sCode = Mid(cboInput(0).Text, 2, 6)
    If sCode = "" Then
        MsgBox "저장은 특정 가맹점를 선택하여 작업하여 주십시요.", vbInformation, "확인"
        cboInput(0).SetFocus
        Exit Sub
    End If
 
    ReDim sValue(2)
    
    For i = 1 To spdView.MaxRows
        
        spdView.Row = i
        spdView.Col = 25
        bChk = spdView.Text
        spdView.Col = 3
        If Trim(spdView.Text) <> "" And bChk = False Then

            sValue(0) = sCode
            spdView.Col = 1
            sValue(1) = Format(spdView.Text, "YYYY-MM-DD")
            sValue(2) = "N"
            
            
            Call ExecPro("SP_04011_01_ALL", sValue(), Err_Num, Err_Dec)
            
            If Err_Num <> 0 Then
                MsgBox "[" & Err_Num & "] " & Err_Dec
            End If
        
        End If
        
    Next i
    Call Data_Display
    MsgBox "해당되는 데이터가 정상적으로 저장이 되었습니다.", vbInformation
    
ERR_RTN:
    PanelsMsg Err.Description
    'Resume Next
End Sub
