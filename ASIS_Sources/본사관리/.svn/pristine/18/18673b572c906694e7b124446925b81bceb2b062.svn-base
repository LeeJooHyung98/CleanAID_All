VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form P_05001 
   Caption         =   "TAG분실 관리"
   ClientHeight    =   9375
   ClientLeft      =   2490
   ClientTop       =   2070
   ClientWidth     =   17250
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_05001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9375
   ScaleWidth      =   17250
   WindowState     =   2  '최대화
   Begin Threed.SSPanel panSub 
      Height          =   3915
      Left            =   570
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   6906
      _Version        =   262144
      BorderWidth     =   5
      BevelInner      =   1
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      FloodShowPct    =   -1  'True
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   2895
         Index           =   2
         Left            =   180
         TabIndex        =   4
         Top             =   780
         Width           =   6435
         _Version        =   524288
         _ExtentX        =   11351
         _ExtentY        =   5106
         _StockProps     =   64
         BackColorStyle  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "P_05001.frx":058A
         AppearanceStyle =   0
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   555
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   979
         _Version        =   262144
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   9375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17250
      _ExtentX        =   30427
      _ExtentY        =   16536
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_05001.frx":09CC
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   7395
         Index           =   0
         Left            =   15
         TabIndex        =   1
         Top             =   1230
         Width           =   17220
         _Version        =   524288
         _ExtentX        =   30374
         _ExtentY        =   13044
         _StockProps     =   64
         BackColorStyle  =   1
         EditEnterAction =   5
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         SpreadDesigner  =   "P_05001.frx":0A5E
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   720
         Index           =   1
         Left            =   15
         TabIndex        =   2
         Top             =   8640
         Width           =   17220
         _Version        =   524288
         _ExtentX        =   30374
         _ExtentY        =   1270
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
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         MaxCols         =   9
         MaxRows         =   1
         ScrollBars      =   0
         SpreadDesigner  =   "P_05001.frx":0EE9
         UserResize      =   1
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   6
         Top             =   435
         Width           =   17220
         _ExtentX        =   30374
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Left            =   9870
            Style           =   2  '드롭다운 목록
            TabIndex        =   9
            Top             =   405
            Width           =   2895
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   1
            Left            =   6330
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   405
            Width           =   1035
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   0
            Left            =   6330
            TabIndex        =   7
            Top             =   60
            Width           =   1035
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Left            =   1530
            TabIndex        =   10
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   63963136
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   11
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "분 실 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   4860
            TabIndex        =   12
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "문 서 번 호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   4860
            TabIndex        =   13
            Top             =   405
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "새문서번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   3
            Left            =   7950
            TabIndex        =   14
            Top             =   405
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "분실일자/문서번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   405
         Left            =   15
         TabIndex        =   15
         Top             =   15
         Width           =   17220
         _ExtentX        =   30374
         _ExtentY        =   714
         _Version        =   262144
         Font3D          =   3
         ForeColor       =   192
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
         Caption         =   " TAG분실 관리 (P_05001)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_05001.frx":14F9
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
End
Attribute VB_Name = "P_05001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim RS02 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cboInput_Click()
    dtInput.Value = Mid(cboInput.Text, 1, 10)
    txtInput(0).Text = Mid(cboInput.Text, 13)
End Sub

Private Sub Form_Activate()
'    cmdBtn(0).Enabled = True
'    cmdBtn(1).Enabled = True
'    cmdBtn(2).Enabled = True
'    cmdBtn(3).Enabled = True
'    cmdBtn(5).Enabled = True
'    cmdBtn(6).Enabled = True
'
'    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    If P_05001_Flag = False Then
        dtInput.Value = Date
        
        Call New_No
        
        ReDim sValue(2)
        
        sValue(0) = "1"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_05001_00", sValue(), Err_Num, Err_Dec)
        
        spdView(0).MaxCols = RS01.Fields.Count
        spdView(0).MaxRows = RS01.RecordCount
        
        Call spdDisplay2(RS01)
        Call GetColWidth(REG_App, Me.Name, spdView(0))
        
        Call spdComboAdd
        Call Data_Display2
        
        P_05001_Flag = True
        
        spdView(0).EditModePermanent = True
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub spdDisplay2(Rs As ADODB.Recordset)

    Call fpSpread_Display(spdView(0), Rs)
    
    spdView(0).ColsFrozen = 1 '틀고정
    
    spdView(0).Row = -1
    
    spdView(0).Col = 1
    spdView(0).ColWidth(1) = 18
    spdView(0).CellType = CellTypeComboBox
    spdView(0).TypeVAlign = TypeVAlignCenter
    spdView(0).TypeHAlign = TypeHAlignLeft

    spdView(0).Col = 2
    spdView(0).ColWidth(2) = 15
    spdView(0).CellType = CellTypeEdit
    spdView(0).TypeVAlign = TypeVAlignCenter
    spdView(0).TypeHAlign = TypeHAlignLeft

    spdView(0).Col = 3
    spdView(0).ColWidth(3) = 14
    spdView(0).CellType = CellTypeEdit
    spdView(0).TypeVAlign = TypeVAlignCenter
    spdView(0).TypeHAlign = TypeHAlignLeft

    spdView(0).Col = 4
    spdView(0).ColWidth(4) = 6
    spdView(0).CellType = CellTypeEdit
    spdView(0).TypeVAlign = TypeVAlignCenter
    spdView(0).TypeHAlign = TypeHAlignLeft

    spdView(0).Col = 5
    spdView(0).ColWidth(5) = 14
    spdView(0).CellType = CellTypeEdit
    spdView(0).TypeVAlign = TypeVAlignCenter
    spdView(0).TypeHAlign = TypeHAlignLeft

    spdView(0).Col = 6
    spdView(0).ColWidth(6) = 10
    spdView(0).CellType = CellTypeDate
    spdView(0).TypeDateCentury = True
    spdView(0).TypeDateFormat = TypeDateFormatYYMMDD
    spdView(0).TypeVAlign = TypeVAlignCenter
    spdView(0).TypeHAlign = TypeHAlignCenter

    spdView(0).Col = 7
    spdView(0).ColWidth(7) = 6
    spdView(0).CellType = CellTypePic
    spdView(0).TypePicMask = "9-999"
    spdView(0).TypeVAlign = TypeVAlignCenter
    spdView(0).TypeHAlign = TypeHAlignCenter

    spdView(0).Col = 8
    spdView(0).ColWidth(8) = 10
    spdView(0).CellType = CellTypeDate
    spdView(0).TypeDateCentury = True
    spdView(0).TypeDateFormat = TypeDateFormatYYMMDD
    spdView(0).TypeVAlign = TypeVAlignCenter
    spdView(0).TypeHAlign = TypeHAlignCenter

    spdView(0).Col = 9
    spdView(0).ColWidth(9) = 10
    spdView(0).CellType = CellTypeDate
    spdView(0).TypeDateCentury = True
    spdView(0).TypeDateFormat = TypeDateFormatYYMMDD
    spdView(0).TypeVAlign = TypeVAlignCenter
    spdView(0).TypeHAlign = TypeHAlignCenter

    spdView(0).Col = 10
    spdView(0).ColWidth(10) = 6
    spdView(0).CellType = CellTypeEdit
    spdView(0).TypeVAlign = TypeVAlignCenter
    spdView(0).TypeHAlign = TypeHAlignLeft

    spdView(0).Col = 11
    spdView(0).ColWidth(11) = 6
    spdView(0).CellType = CellTypeEdit
    spdView(0).TypeVAlign = TypeVAlignCenter
    spdView(0).TypeHAlign = TypeHAlignLeft
    
    spdView(0).Col = 12
    spdView(0).ColWidth(12) = 4
    spdView(0).CellType = CellTypeEdit
    spdView(0).TypeVAlign = TypeVAlignCenter
    spdView(0).TypeHAlign = TypeHAlignLeft
    
    spdView(0).EditMode = False
    
    spdView(0).DataSource = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
    P_05001_Flag = False
End Sub

Private Sub Data_Display2()
    ReDim sValue(1)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput.Value, "yyyy")
        
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_05001_02", sValue(), Err_Num, Err_Dec)
    
    Do While Not RS01.EOF
        cboInput.AddItem RS01!분실일자 & Space(2) & RS01!문서번호
        
        RS01.MoveNext
    Loop
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i As Integer
    Dim ii As Integer
    
    ReDim sValue(2)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput.Value, "YYYY-MM-DD")
    sValue(2) = txtInput(0).Text
        
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_05001_00", sValue(), Err_Num, Err_Dec)
    
    spdView(0).MaxCols = RS01.Fields.Count
    spdView(0).MaxRows = RS01.RecordCount
    
    Call spdDisplay2(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView(0))
    Call spdComboAdd
    
    i = 0
    
    Do While Not RS01.EOF
         i = i + 1
        spdView(0).Row = i
        spdView(0).Col = 1
        
        For ii = 0 To spdView(0).TypeComboBoxCount
            spdView(0).TypeComboBoxIndex = ii
            If Trim(RS01!매장코드) = Trim(spdView(0).TypeComboBoxString) Then
                spdView(0).TypeComboBoxCurSel = ii
                Exit For
            End If
        Next ii
       
        RS01.MoveNext
    Loop
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub New_No()
    On Error GoTo ErrRtn
    
    ReDim sValue(2)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput.Value, "YYYY-MM-DD")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_05001_01", sValue(), Err_Num, Err_Dec)
    
    If RS01.RecordCount > 0 Then
        If IsNull(RS01!책번호) Then
            txtInput(1).Text = "1"
        Else
            txtInput(1).Text = RS01!책번호
        End If
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub spdView_Change(Index As Integer, ByVal Col As Long, ByVal Row As Long)
    If Index = 0 Then
        If Col = 7 Then
            Call TagChk
        End If
    End If
End Sub

Private Sub spdView_DblClick(Index As Integer, ByVal Col As Long, ByVal Row As Long)
    If Index = 1 Then
        Dim i As Integer
        
        spdView(1).Row = Row
        spdView(1).Col = 1
        dtInput.Value = Format(spdView(1).Text, "YYYY-MM-DD")
        spdView(1).Col = 2
        txtInput(0).Text = spdView(1).Text
    
        ReDim sValue(2)
        
        sValue(0) = "0"
        sValue(1) = Format(dtInput.Value, "YYYY-MM-DD")
        sValue(2) = txtInput(0).Text
            
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_05001_00", sValue(), Err_Num, Err_Dec)
        
        spdView(0).MaxCols = RS01.Fields.Count
        spdView(0).MaxRows = RS01.RecordCount
        
        Call spdDisplay2(RS01)
        Call GetColWidth(REG_App, Me.Name, spdView(0))
        Call spdComboAdd
    
        Do While Not RS01.EOF
            i = i + 1
            spdView(0).Row = i
            spdView(0).Col = 1
            
            spdView(0).Text = RS01!매장코드
            
            RS01.MoveNext
        Loop
    ElseIf Index = 2 Then
        spdView(1).MaxRows = 0
        spdView(1).MaxRows = 1
        
        spdView(2).Row = Row
        
        spdView(2).Col = 3
        spdView(1).Row = 1
        spdView(1).Col = 1          ' 택번호
        spdView(1).Text = spdView(2).Text
        
        spdView(2).Col = 5
        spdView(1).Row = 1
        spdView(1).Col = 2          ' 전화번호
        spdView(1).Text = spdView(2).Text
        
        spdView(2).Col = 4
        spdView(1).Row = 1
        spdView(1).Col = 3          ' 성명
        spdView(1).Text = spdView(2).Text
    
        spdView(2).Col = 7
        spdView(1).Row = 1
        spdView(1).Col = 4          ' 품명
        spdView(1).Text = spdView(2).Text
        
        spdView(2).Col = 9
        spdView(1).Row = 1
        spdView(1).Col = 5          ' 금액
        spdView(1).Text = spdView(2).Text
    
        spdView(2).Col = 8
        spdView(1).Row = 1
        spdView(1).Col = 6          ' 색상
        spdView(1).Text = spdView(2).Text
        
        spdView(2).Col = 10
        spdView(1).Row = 1
        spdView(1).Col = 7          ' 내용
        spdView(1).Text = spdView(2).Text
        
        spdView(2).Col = 11
        spdView(1).Row = 1
        spdView(1).Col = 8          ' 상표
        spdView(1).Text = spdView(2).Text
    
        spdView(2).Col = 12
        spdView(1).Row = 1
        spdView(1).Col = 9          ' 상태
        spdView(1).Text = spdView(2).Text
        
        spdView(0).Row = spdView(0).ActiveRow
        spdView(0).Col = 6
        spdView(2).Col = 1
        spdView(0).Text = spdView(2).Text
        
        panSub.Visible = False
        spdView(2).MaxRows = 0
    End If
End Sub

Public Sub DataAdd()
    dtInput.Value = Date
    
    spdView(0).MaxRows = 0
    spdView(0).MaxRows = 1
End Sub

Public Sub DataSave()
    Dim i As Integer
    
    ReDim sValue(13)
    
    sValue(0) = Format(dtInput.Value, "YYYY-MM-DD")       ' 분실일자
    
    If txtInput(0).Text = "" Then
        MsgBox "문서번호를 입력하여 주시기 바랍니다..", vbInformation
        txtInput(0).SetFocus
        Exit Sub
    Else
        sValue(1) = Val(txtInput(0).Text)               ' 책번호
    End If
    
    For i = 1 To spdView(0).MaxRows
        spdView(0).Row = i
        spdView(0).Col = 0: sValue(2) = i                                                       ' 순번
        spdView(0).Col = 6: sValue(3) = Format(spdView(0).Text, "YYYY-MM-DD")                   ' 입고일자
        
        spdView(0).Col = 1                                                                      ' 대리점코드
        If Mid(spdView(0).Text, 2, 1) = "#" Then
            sValue(4) = Mid(spdView(0).Text, 2, 1)
        Else
            sValue(4) = Mid(spdView(0).Text, 2, 3)
        End If
        
        spdView(0).Col = 7: sValue(5) = Mid(spdView(0).Text, 1, 1) & Mid(spdView(0).Text, 3, 3) ' 택번호
        spdView(0).Col = 2: sValue(6) = spdView(0).Text                                         ' 아이템
        spdView(0).Col = 3: sValue(7) = spdView(0).Text                                         ' 브랜드
        spdView(0).Col = 4: sValue(8) = spdView(0).Text                                         ' 색상
        spdView(0).Col = 5:  sValue(9) = spdView(0).Text                         ' 특징
        spdView(0).Col = 8:  sValue(10) = Format(spdView(0).Text, "YYYY-MM-DD")  ' 출고일자
        spdView(0).Col = 10: sValue(11) = "0"                                    ' 전송Check
        spdView(0).Col = 9:  sValue(12) = Format(spdView(0).Text, "YYYY-MM-DD")  ' 확인일자
        spdView(0).Col = 11: sValue(13) = spdView(0).Text                        ' 메모
                
        Call ExecPro("SP_05001_04", sValue(), Err_Num, Err_Dec)
        
        If Err_Num <> 0 Then
            MsgBox "[" & Err_Num & "] " & Err_Dec
            Exit Sub
        End If
    Next i
    
    If Err_Num = 0 Then
        MsgBox "해당사항이 정상적으로 저장이 되었습니다.", vbInformation
    End If
End Sub

Private Sub spdComboAdd()
    Dim sItem As String
    
    ReDim sValue(0)
    
    sValue(0) = "0"
    
    Set RS02 = New ADODB.Recordset
    Set RS02 = ExecPro("SP_00003", sValue(), Err_Num, Err_Dec)

    sItem = "[#]" & Chr(9)
    
    Do While Not RS02.EOF
        sItem = sItem & "[" & RS02!가맹점코드 & "] " & RS02!가맹점명 & Chr(9)
        
        RS02.MoveNext
    Loop
    
    spdView(0).Row = -1
    spdView(0).Col = 1
    spdView(0).TypeComboBoxList = sItem
End Sub

Private Sub TagChk()
    ReDim sValue(1)
    
    spdView(0).Row = spdView(0).ActiveRow
    spdView(0).Col = 1
    sValue(0) = Mid(spdView(0).Text, 2, 3)
    spdView(0).Col = 7
    sValue(1) = spdView(0).Value
    'sValue(1) = Mid(spdView(0).Text, 1, 1) & Mid(spdView(0).Text, 3, 3)
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_05001_03", sValue(), Err_Num, Err_Dec)
    
    spdView(2).MaxCols = RS01.Fields.Count
    spdView(2).MaxRows = RS01.RecordCount
    
    If RS01.RecordCount = 0 Then
        MsgBox "입력하신 택번호는 입고일자에 존재하지 않습니다", vbInformation
        Exit Sub
    Else
        panSub.Visible = True
        
        Call spdDisplay3(RS01)
    End If
End Sub

Public Sub DataDelete()
    If MsgBox("해당되는 데이터를 삭제하시겠습니까?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        ReDim sValue(2)
        
        sValue(0) = Format(dtInput.Value, "YYYY-MM-DD")
        sValue(1) = txtInput(0).Text
        spdView(0).Row = spdView(0).ActiveRow
        spdView(0).Col = 12
        sValue(2) = spdView(0).Text
            
        Call ExecPro("SP_05001_05", sValue(), Err_Num, Err_Dec)
        
        spdView(0).Row = spdView(0).ActiveRow
        spdView(0).Action = ActionDeleteRow
        spdView(0).MaxRows = spdView(0).MaxRows - 1
        
        If Err_Num = 0 Then
            MsgBox "해당되는 데이터가 정상적으로 삭제되었습니다.", vbInformation
        Else
            MsgBox "[" & Err_Num & "] " & Err_Dec
        End If
    End If
End Sub

Private Sub spdView_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Or KeyAscii = 13 Then
        If spdView(0).ActiveCol = spdView(0).MaxCols And spdView(0).ActiveRow = spdView(0).MaxRows Then
            spdView(0).MaxRows = spdView(0).MaxRows + 1
            
'            spdView(0).Row = spdView(0).MaxRows
'            spdView(0).Col = 1
'            spdView(0).Action = ActionActiveCell
        End If
    End If
End Sub

Private Sub spdView_LeaveCell(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If NewRow <> -1 Then
        spdView(Index).Row = Row
        spdView(Index).Col = -1
        spdView(Index).BackColor = vbWhite
        
        spdView(Index).Row = NewRow
        spdView(Index).Col = -1
        spdView(Index).BackColor = glbYellow
    End If
End Sub

Private Sub spdView_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        'PopupMenu P_00000.PopUp
    End If
End Sub

Public Sub DataScreen()
'    Dim ReportFP As String
'    Dim ReportFile As String
'
'    ReportFP = GetIniStr("REPORT", "FilePath", "", m_iniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'
'    Dim i As Integer
'    For i = 0 To 30
'        P_00000.crPrint.Formulas(i) = ""
'    Next
'
'    P_00000.crPrint.Formulas(0) = "문서번호 = '" & Format(dtInput.Value, "yyyymmdd") & " - " & txtInput(0).Text & "'"
'
'    P_00000.crPrint.StoredProcParam(0) = "0"
'    P_00000.crPrint.StoredProcParam(1) = Format(dtInput.Value, "yyyymmdd")
'    P_00000.crPrint.StoredProcParam(2) = txtInput(0).Text
'
'    Call ReportPrint(ReportFile, "2")
End Sub

Public Sub DataPrint()
'    Dim ReportFP As String
'    Dim ReportFile As String
'
'    ReportFP = GetIniStr("REPORT", "FilePath", "", m_iniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'
'    Dim i As Integer
'    For i = 0 To 30
'        P_00000.crPrint.Formulas(i) = ""
'    Next
'
'    P_00000.crPrint.Formulas(0) = "문서번호 = '" & Format(dtInput.Value, "yyyymmdd") & " - " & txtInput(0).Text & "'"
'
'    P_00000.crPrint.StoredProcParam(0) = "0"
'    P_00000.crPrint.StoredProcParam(1) = Format(dtInput.Value, "yyyymmdd")
'    P_00000.crPrint.StoredProcParam(2) = txtInput(0).Text
'
'    Call ReportPrint(ReportFile, "1")
End Sub

Private Sub spdDisplay3(Rs As ADODB.Recordset)

    Call fpSpread_Display(spdView(2), Rs)
    
    spdView(2).ColsFrozen = 1 '틀고정
    
    spdView(2).Row = -1
    
    spdView(2).Col = 1
    spdView(2).ColWidth(1) = 10
    spdView(2).CellType = CellTypeDate
    spdView(2).TypeDateCentury = True
    spdView(2).TypeDateFormat = TypeDateFormatYYMMDD
    spdView(2).TypeVAlign = TypeVAlignCenter
    spdView(2).TypeHAlign = TypeHAlignLeft

    spdView(2).Col = 2
    spdView(2).ColWidth(2) = 10
    spdView(2).CellType = CellTypeDate
    spdView(2).TypeDateCentury = True
    spdView(2).TypeDateFormat = TypeDateFormatYYMMDD
    spdView(2).TypeVAlign = TypeVAlignCenter
    spdView(2).TypeHAlign = TypeHAlignLeft

    spdView(2).Col = 3
    spdView(2).ColWidth(3) = 8
    spdView(2).CellType = CellTypePic
    spdView(2).TypePicMask = "9-999"
    spdView(2).TypeVAlign = TypeVAlignCenter
    spdView(2).TypeHAlign = TypeHAlignCenter

    spdView(2).Col = 4
    spdView(2).ColWidth(4) = 8
    spdView(2).CellType = CellTypeEdit
    spdView(2).TypeVAlign = TypeVAlignCenter
    spdView(2).TypeHAlign = TypeHAlignLeft

    spdView(2).Col = 5
    spdView(2).ColWidth(5) = 10
    spdView(2).CellType = CellTypeEdit
    spdView(2).TypeVAlign = TypeVAlignCenter
    spdView(2).TypeHAlign = TypeHAlignLeft

    spdView(2).Col = 6
    spdView(2).ColWidth(6) = 8
    spdView(2).CellType = CellTypeEdit
    spdView(2).TypeVAlign = TypeVAlignCenter
    spdView(2).TypeHAlign = TypeHAlignLeft

    spdView(2).Col = 7
    spdView(2).ColWidth(7) = 15
    spdView(2).CellType = CellTypeEdit
    spdView(2).TypeVAlign = TypeVAlignCenter
    spdView(2).TypeHAlign = TypeHAlignLeft

    spdView(2).Col = 8
    spdView(2).ColWidth(8) = 10
    spdView(2).CellType = CellTypeEdit
    spdView(2).TypeVAlign = TypeVAlignCenter
    spdView(2).TypeHAlign = TypeHAlignLeft

    spdView(2).Col = 9
    spdView(2).ColWidth(9) = 12
    spdView(2).CellType = CellTypeFloat
    spdView(2).TypeFloatSeparator = True
    spdView(2).TypeFloatDecimalPlaces = 0
    spdView(2).TypeVAlign = TypeVAlignCenter
    
    spdView(2).Col = 10
    spdView(2).ColWidth(10) = 6
    spdView(2).CellType = CellTypeEdit
    spdView(2).TypeVAlign = TypeVAlignCenter
    spdView(2).TypeHAlign = TypeHAlignLeft

    spdView(2).Col = 11
    spdView(2).ColWidth(11) = 6
    spdView(2).CellType = CellTypeEdit
    spdView(2).TypeVAlign = TypeVAlignCenter
    spdView(2).TypeHAlign = TypeHAlignLeft

    spdView(2).Col = 12
    spdView(2).ColWidth(12) = 6
    spdView(2).CellType = CellTypeEdit
    spdView(2).TypeVAlign = TypeVAlignCenter
    spdView(2).TypeHAlign = TypeHAlignLeft
End Sub

