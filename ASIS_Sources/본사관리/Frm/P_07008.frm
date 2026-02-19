VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form P_07008 
   Caption         =   "수선 외주현황"
   ClientHeight    =   9780
   ClientLeft      =   1485
   ClientTop       =   2145
   ClientWidth     =   16260
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_07008.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9780
   ScaleWidth      =   16260
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   9780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16260
      _ExtentX        =   28681
      _ExtentY        =   17251
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_07008.frx":058A
      Begin Threed.SSPanel SSPanel 
         Height          =   390
         Left            =   15
         TabIndex        =   16
         Top             =   9375
         Width           =   16230
         _ExtentX        =   28628
         _ExtentY        =   688
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   2
            Left            =   7185
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   45
            Width           =   1335
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   1
            Left            =   4125
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   45
            Width           =   1335
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   0
            Left            =   1665
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   45
            Width           =   735
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   6
            Left            =   10245
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   45
            Width           =   1335
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   7
            Left            =   13305
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   45
            Width           =   1335
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   45
            TabIndex        =   22
            Top             =   45
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "수 량 합 계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   3
            Left            =   2505
            TabIndex        =   23
            Top             =   45
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "금 액 합 계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   4
            Left            =   5565
            TabIndex        =   24
            Top             =   45
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "대리점금액합계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   12
            Left            =   11685
            TabIndex        =   25
            Top             =   45
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "입금액합계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   13
            Left            =   8625
            TabIndex        =   26
            Top             =   45
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "외주금액합계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel panInput 
         Height          =   765
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   16230
         _ExtentX        =   28628
         _ExtentY        =   1349
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            Style           =   2  '드롭다운 목록
            TabIndex        =   3
            Top             =   405
            Width           =   3015
         End
         Begin MSMask.MaskEdBox mskInput 
            Height          =   315
            Left            =   9570
            TabIndex        =   2
            Top             =   405
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "#-###"
            PromptChar      =   "_"
         End
         Begin Threed.SSCommand cmdSubBtn 
            Height          =   315
            Left            =   11040
            TabIndex        =   4
            Top             =   405
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "입 고 수 신"
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            TabIndex        =   5
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   61210624
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   6
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "일      자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   4830
            TabIndex        =   7
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   61210624
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   8
            Left            =   60
            TabIndex        =   8
            Top             =   405
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "대 리 점 명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   9
            Left            =   8100
            TabIndex        =   9
            Top             =   405
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "택  번  호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   11
            Left            =   9570
            TabIndex        =   10
            Top             =   60
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   556
            _Version        =   262144
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   0
               Left            =   1860
               TabIndex        =   11
               Top             =   30
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "외주받은날짜"
            End
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   1
               Left            =   180
               TabIndex        =   12
               Top             =   30
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "외주보낸날짜"
               Value           =   -1
            End
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   10
            Left            =   8100
            TabIndex        =   13
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "일 자 구 분"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin VB.Label Label1 
            Alignment       =   2  '가운데 맞춤
            Caption         =   "~"
            Height          =   195
            Left            =   4530
            TabIndex        =   14
            Top             =   120
            Width           =   255
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   8565
         Index           =   0
         Left            =   15
         TabIndex        =   15
         Top             =   795
         Width           =   16230
         _Version        =   524288
         _ExtentX        =   28628
         _ExtentY        =   15108
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
         SpreadDesigner  =   "P_07008.frx":05FC
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_07008"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cmdSubBtn_Click()
    P_TRANS.Show 1
    
    Call DataTrans
End Sub

Private Sub Form_Activate()
'    cmdBtn(0).Enabled = True
'    cmdBtn(3).Enabled = True
'    cmdBtn(5).Enabled = True
'    cmdBtn(6).Enabled = True
'
'    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    If P_07008_Flag = False Then
        dtInput(0).Value = Date
        dtInput(1).Value = Date
        
        ReDim sValue(5)
        
        Call AgencyComboAdd(cboInput(0))
        
        sValue(0) = "1"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_07008_00", sValue(), Err_Num, Err_Dec)
        
        spdView(0).MaxCols = RS01.Fields.Count
        spdView(0).MaxRows = RS01.RecordCount
        
        Call spdDisplay(RS01)
        Call GetColWidth(REG_App, Me.Name & "A", spdView(0))
        
        P_07008_Flag = True
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub spdDisplay(Rs As ADODB.Recordset)
        
    Call fpSpread_Display(spdView(0), Rs)

    
    spdView(0).ColsFrozen = 1 '틀고정
    
    spdView(0).Row = -1
    
    spdView(0).Col = 1
    spdView(0).ColWidth(1) = 10
    spdView(0).CellType = CellTypeDate
    spdView(0).TypeDateCentury = True
    spdView(0).TypeDateFormat = TypeDateFormatYYMMDD
    spdView(0).TypeVAlign = TypeVAlignCenter
    spdView(0).TypeHAlign = TypeHAlignCenter
    
    spdView(0).Col = 2
    spdView(0).ColWidth(2) = 15
    spdView(0).CellType = CellTypeEdit
    spdView(0).TypeVAlign = TypeVAlignCenter
    spdView(0).TypeHAlign = TypeHAlignLeft
    
    spdView(0).Col = 3
    spdView(0).ColWidth(3) = 8
    spdView(0).CellType = CellTypePic
    spdView(0).TypePicMask = "9-999"
    spdView(0).TypeVAlign = TypeVAlignCenter
    spdView(0).TypeHAlign = TypeHAlignLeft

    spdView(0).Col = 4
    spdView(0).ColWidth(4) = 10
    spdView(0).CellType = CellTypeDate
    spdView(0).TypeDateCentury = True
    spdView(0).TypeDateFormat = TypeDateFormatYYMMDD
    spdView(0).TypeVAlign = TypeVAlignCenter
    spdView(0).TypeHAlign = TypeHAlignCenter

    spdView(0).Col = 5
    spdView(0).ColWidth(5) = 15
    spdView(0).CellType = CellTypeEdit
    spdView(0).TypeVAlign = TypeVAlignCenter
    spdView(0).TypeHAlign = TypeHAlignLeft

    spdView(0).Col = 6
    spdView(0).ColWidth(6) = 15
    spdView(0).CellType = CellTypeEdit
    spdView(0).TypeVAlign = TypeVAlignCenter
    spdView(0).TypeHAlign = TypeHAlignLeft
    
    spdView(0).Col = 7
    spdView(0).ColWidth(7) = 12
    spdView(0).CellType = CellTypeFloat
    spdView(0).TypeFloatSeparator = True
    spdView(0).TypeFloatDecimalPlaces = 0
    spdView(0).TypeVAlign = TypeVAlignCenter

    spdView(0).Col = 8
    spdView(0).ColWidth(8) = 12
    spdView(0).CellType = CellTypeFloat
    spdView(0).TypeFloatSeparator = True
    spdView(0).TypeFloatDecimalPlaces = 0
    spdView(0).TypeVAlign = TypeVAlignCenter

    spdView(0).Col = 9
    spdView(0).ColWidth(9) = 12
    spdView(0).CellType = CellTypeFloat
    spdView(0).TypeFloatSeparator = True
    spdView(0).TypeFloatDecimalPlaces = 0
    spdView(0).TypeVAlign = TypeVAlignCenter

    spdView(0).Col = 10
    spdView(0).ColWidth(10) = 12
    spdView(0).CellType = CellTypeFloat
    spdView(0).TypeFloatSeparator = True
    spdView(0).TypeFloatDecimalPlaces = 0
    spdView(0).TypeVAlign = TypeVAlignCenter

    spdView(0).Col = 11
    spdView(0).ColWidth(11) = 10
    spdView(0).CellType = CellTypeDate
    spdView(0).TypeDateCentury = True
    spdView(0).TypeDateFormat = TypeDateFormatYYMMDD
    spdView(0).TypeVAlign = TypeVAlignCenter
    spdView(0).TypeHAlign = TypeHAlignCenter
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_07008_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    ReDim sValue(5)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    sValue(3) = Mid(cboInput(0).Text, 2, 3) & "%"
    sValue(4) = mskInput.ClipText & "%"
    
    If optSelect(0).Value = True Then
        sValue(5) = "1"
    Else
        sValue(5) = "0"
    End If
        
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_07008_00", sValue(), Err_Num, Err_Dec)
    
    spdView(0).MaxCols = RS01.Fields.Count
    spdView(0).MaxRows = RS01.RecordCount
    
    Call spdDisplay(RS01)
    Call GetColWidth(REG_App, Me.Name & "A", spdView(0))
    
    txtInput(0).Text = spdView(0).MaxRows
    
    spdView(0).MaxRows = spdView(0).MaxRows + 1
    spdView(0).Row = spdView(0).MaxRows
    spdView(0).RowHidden = True
    
    spdView(0).Col = 7
    spdView(0).Formula = "SUM(G1:G" & spdView(0).MaxRows - 1 & ")"
    txtInput(1).Text = spdView(0).Text

    spdView(0).Col = 8
    spdView(0).Formula = "SUM(H1:H" & spdView(0).MaxRows - 1 & ")"
    txtInput(2).Text = spdView(0).Text

    spdView(0).Col = 9
    spdView(0).Formula = "SUM(I1:I" & spdView(0).MaxRows - 1 & ")"
    txtInput(6).Text = spdView(0).Text

    spdView(0).Col = 10
    spdView(0).Formula = "SUM(J1:J" & spdView(0).MaxRows - 1 & ")"
    txtInput(7).Text = spdView(0).Text
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub DataTrans()
    Dim rCnt As Integer
    Dim dCnt As Integer
    Dim wCnt As Integer
    
    Dim sData As String
    Dim mDate As String
    Dim sFilePath As String
    
    sFilePath = GetIniStr("TERMINAL DATA", "TerminalFilePath", "", m_iniFile)
    
    If Dir(sFilePath & "\Ibchul.dat") <> "" Then
        ReDim sValue(3)
    
        '핸디로부터 읽은 데이타를 db로
        rCnt = 0
        dCnt = 0
        wCnt = 0
           
        Open sFilePath & "\Ibchul.dat" For Input As #1
        
        spdView(0).MaxRows = 0
        
        Do While Not EOF(1)         ' Loop until end of file.
            Input #1, sData         ' Read line into variable.
            
            sValue(0) = "0"
            sValue(1) = CStr(Year(Date)) & Trim(Mid(sData, 6, 4)) ' dat파일에서 6 자리일자를 8자리로 변환
            
            mDate = Format(sValue(1), "####/##/##")
            
            sValue(2) = Trim(Mid(sData, 10, 3)) ' dat파일에서 read
            sValue(3) = Trim(Mid(sData, 13, 4))
            'sValue(3) = Trim(Mid(sData, 28, 3))
            
            rCnt = rCnt + 1
            
            If Len(Trim(sValue(1))) <> 8 Or Not IsDate(mDate) Or _
               Len(Trim(sValue(2))) <> 3 Or Not IsNumeric(sValue(2)) Or _
               Len(Trim(sValue(3))) <> 4 Or Not IsNumeric(sValue(3)) Then
'               Len(Trim(sValue(3))) <> 3 Or sValue(3) = "000" Then
                dCnt = dCnt + 1
    '            HandyErr.B.MaxRows = HandyErr.B.MaxRows + 1
    '            HandyErr.B.Row = HandyErr.B.MaxRows
    '            HandyErr.B.Col = 1
    '            HandyErr.B.Text = Mid(sData, 4, 17) & Mid(sData, 28, 3)
                GoTo handy_err
            End If
            
            Set RS01 = New ADODB.Recordset
            Set RS01 = ExecPro("SP_07008_01", sValue(), Err_Num, Err_Dec)
            
            If Not RS01.EOF Then
                spdView(0).MaxRows = spdView(0).MaxRows + 1
                spdView(0).Row = spdView(0).MaxRows
                
                spdView(0).Col = 1:  If Not IsNull(RS01!출고일자) Then spdView(0).Text = RS01!출고일자
                spdView(0).Col = 2:  If Not IsNull(RS01!매장코드) Then spdView(0).Text = "[" & RS01!매장코드 & "] " & RS01!매장명
                spdView(0).Col = 3:  If Not IsNull(RS01!택번호) Then spdView(0).Text = RS01!택번호
                spdView(0).Col = 4:  If Not IsNull(RS01!입고일자) Then spdView(0).Text = RS01!입고일자
                spdView(0).Col = 5:  If Not IsNull(RS01!품목) Then spdView(0).Text = "[" & RS01!품목 & "] " & RS01!품명
                spdView(0).Col = 6:  If Not IsNull(RS01!수선내용) Then spdView(0).Text = RS01!수선내용
                spdView(0).Col = 7:  If Not IsNull(RS01!금액) Then spdView(0).Text = Val(RS01!금액)
                spdView(0).Col = 8:  If Not IsNull(RS01!대리점금액) Then spdView(0).Text = Val(RS01!대리점금액)
                spdView(0).Col = 9:  If Not IsNull(RS01!외주금액) Then spdView(0).Text = Val(RS01!외주금액)
                spdView(0).Col = 10: If Not IsNull(RS01!입금금액) Then spdView(0).Text = Val(RS01!입금금액)
                spdView(0).Col = 11: spdView(0).Text = Format(Now, "YYYY-MM-DD")
            End If
            
            DoEvents
handy_err:
        Loop
        Close #1
    End If
    
'    If spdView(0).MaxRows <> 0 Then
'        cmdBtn(2).Enabled = True
'    End If
        
    If Not Dir(sFilePath & "\Ibchul.dat") = "" Then
        Kill sFilePath & "\Ibchul.dat"
    End If
End Sub



Public Sub DataSave()
    Dim i As Integer
    
    ReDim sValue(3)
    
    For i = 1 To spdView(0).MaxRows
        spdView(0).Row = i
        
        spdView(0).Col = 11: sValue(0) = Format(spdView(0).Text, "YYYY-MM-DD")
        spdView(0).Col = 1:  sValue(1) = Format(spdView(0).Text, "YYYY-MM-DD")
        spdView(0).Col = 2:  sValue(2) = Mid(spdView(0).Text, 2, 3)
        spdView(0).Col = 3:  sValue(3) = spdView(0).Value
        
        Call ExecPro("SP_07008_02", sValue(), Err_Num, Err_Dec)
        
        If Err_Num <> 0 Then
            MsgBox "[" & Err_Num & "] " & Err_Dec
            Exit Sub
        End If
    Next i
    
    MsgBox "정상적으로 저장이 되었습니다", vbInformation
'    cmdBtn(2).Enabled = False
End Sub

Public Sub DataPrint()
'    Dim ReportFP As String
'    Dim ReportFile As String
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
'    P_00000.crPrint.Formulas(0) = "일자1 = '" & Format(dtInput(0).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(1) = "일자2 = '" & Format(dtInput(1).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(2) = "대리점 = '" & cboInput(0).Text & "'"
'
'    Call ReportPrint(ReportFile, "1")
End Sub

Public Sub DataScreen()
'    Dim ReportFP As String
'    Dim ReportFile As String
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
'    P_00000.crPrint.Formulas(0) = "일자1 = '" & Format(dtInput(0).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(1) = "일자2 = '" & Format(dtInput(1).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(2) = "대리점 = '" & cboInput(0).Text & "'"
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
    
    For i = 1 To spdView(0).MaxRows - 1
        spdView(0).Row = i
        
        spdView(0).Col = 1
        TempText = LeftH(spdView(0).Text & Space(11), 11)
        spdView(0).Col = 2
        TempText = TempText & LeftH(spdView(0).Text & Space(20), 20)
        spdView(0).Col = 3
        TempText = TempText & LeftH(spdView(0).Text & Space(6), 6)
        spdView(0).Col = 4
        TempText = TempText & LeftH(spdView(0).Text & Space(11), 11)
        spdView(0).Col = 5
        TempText = TempText & LeftH(spdView(0).Text & Space(20), 20)
        spdView(0).Col = 6
        TempText = TempText & LeftH(spdView(0).Text & Space(20), 20)
        spdView(0).Col = 7
        TempText = TempText & RightH(Space(9) & spdView(0).Text, 9) & Space(1)
        spdView(0).Col = 8
        TempText = TempText & RightH(Space(9) & spdView(0).Text, 9) & Space(1)
        spdView(0).Col = 9
        TempText = TempText & RightH(Space(9) & spdView(0).Text, 9) & Space(1)
        spdView(0).Col = 10
        TempText = TempText & RightH(Space(9) & spdView(0).Text, 9) & Space(1)
        spdView(0).Col = 11
        TempText = TempText & LeftH(spdView(0).Text & Space(11), 11)
        
        Print #1, TempText
    Next i
    
    Close #1
End Sub

