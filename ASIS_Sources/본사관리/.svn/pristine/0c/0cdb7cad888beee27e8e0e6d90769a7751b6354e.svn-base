VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form P_03011_01 
   Caption         =   "출고품목 CHECK"
   ClientHeight    =   8220
   ClientLeft      =   1485
   ClientTop       =   2145
   ClientWidth     =   16590
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_03011_01.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   16590
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   8220
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16590
      _ExtentX        =   29263
      _ExtentY        =   14499
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_03011_01.frx":058A
      Begin Threed.SSPanel SSPanel 
         Height          =   405
         Left            =   15
         TabIndex        =   7
         Top             =   7800
         Width           =   16560
         _ExtentX        =   29210
         _ExtentY        =   714
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   2
            Left            =   7995
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   45
            Width           =   1335
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   1
            Left            =   4755
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   45
            Width           =   1335
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   0
            Left            =   1515
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   45
            Width           =   1335
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   45
            TabIndex        =   11
            Top             =   45
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "검 품 수 량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   3
            Left            =   3285
            TabIndex        =   12
            Top             =   45
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "입 고 수 량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   4
            Left            =   6525
            TabIndex        =   13
            Top             =   45
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "다 른 품 목"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel panInput 
         Height          =   765
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   16560
         _ExtentX        =   29210
         _ExtentY        =   1349
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Left            =   1530
            Style           =   2  '드롭다운 목록
            TabIndex        =   2
            Top             =   405
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Left            =   1530
            TabIndex        =   3
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   62980096
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   4
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "검 품 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   5
            Top             =   405
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "대 리 점 명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   6990
         Index           =   1
         Left            =   15
         TabIndex        =   6
         Top             =   795
         Width           =   16560
         _Version        =   524288
         _ExtentX        =   29210
         _ExtentY        =   12330
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
         SpreadDesigner  =   "P_03011_01.frx":05FC
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_03011_01"
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
    Call DataTrans
End Sub

Private Sub cboInput_Click()
    Call Data_Display
End Sub

Private Sub dtInput_Change()
    Call Data_Display
End Sub

Private Sub Form_Activate()
'    cmdBtn(0).Enabled = True
'    cmdBtn(5).Enabled = True
'
'    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
    
    If P_03011_01_Flag = False Then
        dtInput.Value = P_03011.dtInput.Value
        
        DoEvents
        
        Call Data_Display
    
        P_03011_01_Flag = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_03011_01_Flag = False
End Sub

Sub spdDisplay2(Rs As ADODB.Recordset)
    
    Call fpSpread_Display(spdView(1), Rs)
    
    spdView(1).ColsFrozen = 1 '틀고정
    
    spdView(1).Col = 1
    spdView(1).ColWidth(1) = 10
    spdView(1).CellType = CellTypeEdit
    spdView(1).TypeVAlign = TypeVAlignCenter
    spdView(1).TypeHAlign = TypeHAlignLeft

    spdView(1).Col = 2
    spdView(1).ColWidth(2) = 10
    spdView(1).CellType = CellTypeDate
    spdView(1).TypeDateCentury = True
    spdView(1).TypeDateFormat = TypeDateFormatYYMMDD
    spdView(1).TypeVAlign = TypeVAlignCenter
    spdView(1).TypeHAlign = TypeHAlignLeft

    spdView(1).Col = 3
    spdView(1).ColWidth(3) = 10
    spdView(1).CellType = CellTypeDate
    spdView(1).TypeDateCentury = True
    spdView(1).TypeDateFormat = TypeDateFormatYYMMDD
    spdView(1).TypeVAlign = TypeVAlignCenter
    spdView(1).TypeHAlign = TypeHAlignLeft

    spdView(1).Col = 4
    spdView(1).ColWidth(4) = 10
    spdView(1).CellType = CellTypeEdit
    spdView(1).TypeVAlign = TypeVAlignCenter
    spdView(1).TypeHAlign = TypeHAlignCenter

    spdView(1).Col = 5
    spdView(1).ColWidth(5) = 10
    spdView(1).CellType = CellTypeEdit
    spdView(1).TypeVAlign = TypeVAlignCenter
    spdView(1).TypeHAlign = TypeHAlignLeft

    spdView(1).Col = 6
    spdView(1).ColWidth(6) = 10
    spdView(1).CellType = CellTypeEdit
    spdView(1).TypeVAlign = TypeVAlignCenter
    spdView(1).TypeHAlign = TypeHAlignLeft

    spdView(1).Col = 7
    spdView(1).ColWidth(7) = 10
    spdView(1).CellType = CellTypeFloat
    spdView(1).TypeFloatSeparator = True
    spdView(1).TypeFloatDecimalPlaces = 0
    spdView(1).TypeVAlign = TypeVAlignCenter

    spdView(1).Col = 8
    spdView(1).ColWidth(8) = 10
    spdView(1).CellType = CellTypeEdit
    spdView(1).TypeVAlign = TypeVAlignCenter
    spdView(1).TypeHAlign = TypeHAlignLeft

    spdView(1).Col = 9
    spdView(1).ColWidth(9) = 10
    spdView(1).CellType = CellTypeEdit
    spdView(1).TypeVAlign = TypeVAlignCenter
    spdView(1).TypeHAlign = TypeHAlignLeft
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i As Integer
    
    i = P_03011.ActiveControl.Index
    
    ReDim sValue(2)
    
    sValue(0) = "0"
    
    sValue(1) = Format(dtInput.Value, "YYYY-MM-DD")
    
    P_03011.spdView(i).Row = P_03011.spdView(i).ActiveRow
    P_03011.spdView(i).Col = 1
    
    sValue(2) = Mid(P_03011.spdView(i).Text, 2, 3)
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_03011_01", sValue(), Err_Num, Err_Dec)
    
    spdView(1).MaxCols = RS01.Fields.Count
    spdView(1).MaxRows = RS01.RecordCount
    
    Call spdDisplay2(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView(1))
    
    P_03011.spdView(i).Row = P_03011.spdView(i).ActiveRow
    P_03011.spdView(i).Col = 2: txtInput(0).Text = P_03011.spdView(i).Text
    P_03011.spdView(i).Col = 3: txtInput(1).Text = P_03011.spdView(i).Text
    P_03011.spdView(i).Col = 4: txtInput(2).Text = P_03011.spdView(i).Text
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub DataTrans()
    Dim iCnt As Integer
    Dim TmpStr As String
    Dim sInOut As String
    Dim sTagNo As String
    Dim sItem As String
    Dim sDate As String
    Dim sCode As String
    Dim mDate As String
    Dim bDate As String
    
    Dim DupCnt As Integer
    Dim ErrorCnt As Integer
    
    Dim sFilePath As String
    
    P_TRANS.saveYN = False
    P_TRANS.Show 1
    P_TRANS.Hide

    If P_TRANS.saveYN = False Then
        Exit Sub
    End If

    '중복검사 화면
    spdView(0).MaxRows = 0
    DupCnt = 0
    
    spdView(1).MaxRows = 0
    
    spdView(2).MaxRows = 0
    ErrorCnt = 0
    
    iCnt = 0
    
    sFilePath = GetIniStr("TERMINAL DATA", "TerminalFilePath", "", m_iniFile)
    
    '핸디로부터 읽은 데이타를 db로
    Open sFilePath & "\Ibchul.dat" For Input As #1
    
    Do While Not EOF(1)
        iCnt = iCnt + 1
        Input #1, TmpStr
        
        sDate = CStr(Year(Date)) & Trim(Mid(TmpStr, 6, 4))  ' dat파일에서 6 자리일자를 8자리로 변환
        mDate = Format(sDate, "####/##/##")
        sCode = Trim(Mid(TmpStr, 10, 3))                    ' dat파일에서 read
        sTagNo = Trim(Mid(TmpStr, 13, 4))                   ' ''
        sItem = Trim(Mid(TmpStr, 28, 3))                    ' 소품구분
        
        '비정상 자료 check
        If Len(Trim(sDate)) <> 8 Or Not IsDate(mDate) Or _
           Len(Trim(sCode)) <> 3 Or Not IsNumeric(sCode) Or _
           Len(Trim(sTagNo)) <> 4 Or Not IsNumeric(sTagNo) Or _
           Len(Trim(sItem)) <> 3 Or sItem = "000" Then
           
              spdView(0).MaxRows = spdView(0).MaxRows + 1
              spdView(0).Row = spdView(0).MaxRows
              spdView(0).Col = 1
              spdView(0).Text = Mid(TmpStr, 1, 20)
              
              GoTo handy_err
        End If
        
        ReDim sValue(3)
        
        sValue(0) = sDate
        sValue(1) = sCode
        sValue(2) = sTagNo
        sValue(3) = sItem
        
        Call ExecPro("SP_03011_02", sValue(), Err_Num, Err_Dec)
         
        DoEvents
handy_err:
    Loop
    Close #1
    
    If ErrorCnt > 0 Then
       MsgBox CStr(ErrorCnt) & "건이 입고내역이 없습니다."
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
