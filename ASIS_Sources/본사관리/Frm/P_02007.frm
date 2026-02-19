VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_02007 
   Caption         =   "입고검품 수신"
   ClientHeight    =   10305
   ClientLeft      =   525
   ClientTop       =   2400
   ClientWidth     =   14985
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_02007.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10305
   ScaleWidth      =   14985
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   10305
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   18177
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_02007.frx":058A
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   9315
         Left            =   15
         TabIndex        =   1
         Top             =   540
         Width           =   14955
         _Version        =   524288
         _ExtentX        =   26379
         _ExtentY        =   16431
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
         SpreadDesigner  =   "P_02007.frx":061C
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   420
         Left            =   15
         TabIndex        =   2
         Top             =   9870
         Width           =   14955
         _ExtentX        =   26379
         _ExtentY        =   741
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   2
            Left            =   60
            TabIndex        =   3
            Top             =   45
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "총  건  수"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   0
            Left            =   3240
            TabIndex        =   4
            Top             =   45
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "미  전  송"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   330
            Index           =   1
            Left            =   6375
            TabIndex        =   5
            Top             =   45
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   582
            _Version        =   262144
            Caption         =   "전      송"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   345
            Index           =   0
            Left            =   1530
            TabIndex        =   6
            Top             =   45
            Width           =   1215
            _Version        =   262145
            _ExtentX        =   2143
            _ExtentY        =   609
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   3
            StartText.y     =   5
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   13
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   345
            Index           =   1
            Left            =   4710
            TabIndex        =   7
            Top             =   45
            Width           =   1215
            _Version        =   262145
            _ExtentX        =   2143
            _ExtentY        =   609
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   3
            StartText.y     =   5
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   13
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   345
            Index           =   2
            Left            =   7845
            TabIndex        =   8
            Top             =   45
            Width           =   1215
            _Version        =   262145
            _ExtentX        =   2143
            _ExtentY        =   609
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   3
            StartText.y     =   5
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   13
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   9
         Top             =   15
         Width           =   7350
         _ExtentX        =   12965
         _ExtentY        =   900
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   255
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
         Caption         =   " #"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_02007.frx":0A95
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Left            =   7380
         TabIndex        =   10
         Top             =   15
         Width           =   7590
         _ExtentX        =   13388
         _ExtentY        =   900
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
         PictureBackgroundStyle=   2
         PictureBackground=   "P_02007.frx":0C97
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   11
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "종료"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_02007.frx":0E99
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   12
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "화면"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_02007.frx":1433
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   13
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "인쇄"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_02007.frx":19CD
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   14
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "취소"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_02007.frx":1F67
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   15
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "삭제"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_02007.frx":2501
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   16
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "저장"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_02007.frx":2A9B
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   17
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "신규"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_02007.frx":3035
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   18
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "조회"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_02007.frx":35CF
         End
      End
   End
End
Attribute VB_Name = "P_02007"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Err_Num As Long
Dim Err_Dec As String

Dim sValue() As String

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: 'Call Data_Display   ' 조회
        Case 1: 'Call DataAdd        ' 신규
        Case 2: Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
        Case 4: 'Call DataCancel     ' 취소
        Case 5: 'Call DataPrint      ' 인쇄
        Case 6: 'Call DataScreen     '
        Case 7: Unload Me           ' 종료
    End Select
    
'    Me.MousePointer = 0
    
    Exit Sub
    
ErrRtn:
    Me.MousePointer = 0
    
    If Err.Number = "0" Then
        
    ElseIf Err.Number = "91" Then
        End
    Else
        Resume Next
    End If
End Sub

Private Sub Form_Activate()
    cmdBtn(2).Enabled = True
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
End Sub

Public Sub DataSave()
    Dim rCnt As Integer
    Dim iCnt As Integer
    Dim dCnt As Integer
    Dim wCnt As Integer
    
    Dim TmpStr As String
    Dim sDate As String
    Dim sCode As String
    Dim sTagNo As String
    Dim sItem As String
    Dim sApm As String
    Dim mDate As String
    
    Dim sFilePath As String
  
    P_TRANS.saveYN = False
    P_TRANS.Show 1
    P_TRANS.Hide
    
    If P_TRANS.saveYN = False Then Exit Sub
    
    spdView.MaxRows = 0
       
    '핸디로부터 읽은 데이타를 db로
    rCnt = 0
    dCnt = 0
    wCnt = 0
    
'    sApm = Mid(Date, 20, 2)
    sApm = "AM"
    
    Call PanelsMsg("Data를 DB로 UpDate 중입니다.")
    
    sFilePath = GetIniStr("TERMINAL DATA", "TerminalFilePath", "", m_iniFile)
    
    Open sFilePath & "\Ibchul.dat" For Input As #1
    
    Do While Not EOF(1) ' Loop until end of file.
        ReDim sValue(4)
        
        Input #1, TmpStr  ' Read line into variable.
        
        sValue(0) = CStr(Year(Date)) & Trim(Mid(TmpStr, 6, 4))  ' dat파일에서 6 자리일자를 8자리로 변환
        sValue(1) = "AM"
        sValue(2) = Trim(Mid(TmpStr, 10, 3))                    ' dat파일에서 read
        sValue(3) = Trim(Mid(TmpStr, 13, 4))
        sValue(4) = Trim(Mid(TmpStr, 28, 3))
        
        If Mid(sValue(4), 1, 1) < "A" Or Mid(sValue(4), 1, 1) > "Z" Or Not IsNumeric(Mid(sItem, 2, 2)) Then
            sValue(4) = " "
        End If
        
        rCnt = rCnt + 1
        
        If Len(Trim(sValue(0))) <> 8 Or Not IsDate(Format(sValue(0), "####-##-##")) Or Len(Trim(sValue(2))) <> 3 Or Not IsNumeric(sValue(2)) Or Len(Trim(sValue(3))) <> 4 Or Not IsNumeric(sValue(2)) Then
            dCnt = dCnt + 1
            
            spdView.MaxRows = spdView.MaxRows + 1
            spdView.Row = spdView.MaxRows
            spdView.Col = 1
            spdView.Text = Mid(TmpStr, 4, 17) & Mid(TmpStr, 28, 3)
            
            GoTo handy_err
        End If
      
        Call ExecPro("SP_02007_00", sValue(), Err_Num, Err_Dec)
        
        wCnt = wCnt + 1
        
        txtNum(0).Text = Format(rCnt, "#,##0")
        txtNum(1).Text = Format(dCnt, "#,##0")
        txtNum(2).Text = Format(wCnt, "#,##0")
        
        DoEvents
handy_err:
    Loop
    
    Close #1
        
    If spdView.MaxRows > 0 Then
       txtNum(0).Text = rCnt
       txtNum(1).Text = spdView.MaxRows
    End If
        
    If Not Dir(sFilePath & "\Ibchul.dat") = "" Then
        Kill sFilePath & "\Ibchul.dat"
    End If
    
    Call PanelsMsg("   Data전송 완료..!")
End Sub

Public Sub DataPrint()
    Dim sData As String
    Dim i, ii, iii As Integer
    Dim iRow As Integer
    Dim memRow As Long
    Dim lLineQty As Long
    Dim lLinePri As Double
    Dim lLineAmt As Double
    Dim lTotalQty As Long
    Dim lTotalPri As Double
    Dim lTotalAmt As Double
    Dim lTotalVAT As Double
    Dim sPrintData As String
    Dim Pum_Code As String

    Printer.PaperSize = vbPRPSA4
    memRow = 1

PrintHead:

    Printer.Font = "굴림체"                             ' Printer의 사용 글자
    Printer.FontSize = "16"                             ' Print의 글자크기
    Printer.ScaleMode = vbMillimeters                   ' Print의 위치 선정을 밀리미터로 나타낸다.
    iRow = iRow + 2
    Printer.CurrentY = iRow
    Printer.CurrentX = 75
    Printer.Print Me.Caption

    Printer.Font = "굴림체"                             ' Printer의 사용 글자
    Printer.FontSize = "10"                             ' Print의 글자크기
    Printer.ScaleMode = vbMillimeters                   ' Print의 위치 선정을 밀리미터로 나타낸다.
    iRow = iRow + 14
    Printer.CurrentY = iRow
    Printer.CurrentX = 0
    Printer.Print "(주)백상"

    iRow = iRow + 5
    Printer.CurrentY = iRow
    Printer.CurrentX = 0
    sData = "총 건 수 : " & txtNum(0).Text & Space(20)
    sData = LeftH(RTrim(sData) & Space(65), 65) & "성  명 : " & USERNAME
    Printer.Print sData

    iRow = iRow + 5
    Printer.CurrentY = iRow
    Printer.CurrentX = 0
    sData = ""
    sData = "미 전 송 : " & txtNum(1).Text & Space(20)
    sData = LeftH(RTrim(sData) & Space(65), 65) & "출력일자 : " & Format(Now, "YYYY-MM-DD")
    Printer.Print sData

    iRow = iRow + 5
    Printer.CurrentY = iRow
    Printer.CurrentX = 0
    sData = ""
    sData = "전    송 : " & txtNum(2).Text & Space(20)
    sData = LeftH(RTrim(sData) & Space(65), 65) & "출력시간 : " & Format(Now, "hh:mm:ss")
    Printer.Print sData

    iRow = iRow + 1
    Printer.Line (0, iRow + 3)-(240, iRow + 3)

    iRow = iRow + 4
    Printer.CurrentY = iRow
    Printer.CurrentX = 0
    Printer.Print ""
    
    iRow = iRow + 1
    Printer.Line (0, iRow + 3)-(240, iRow + 3)

''    For i = memRow To spdView.MaxRows - 1
''        spdView.Row = i
''
''        spdView.Col = 1
''        sData = Left(spdView.Text, 10)                                             '순위
''
''        spdView.Col = 2
''        sData = Left(spdView.Text, 25)                                             '품목명
''
''        spdView.Col = 3
''        sData = sData & Space(1) & Right(Space(10) & spdView.Text, 5) & " "        '수량
''
''        spdView.Col = 4
''        sData = sData & Space(1) & Right(Space(10) & spdView.Text, 10) & " "       '금액
''
''        spdView.Col = 5
''        sData = sData & Space(3) & Right(Space(10) & spdView.Text, 10) & " "       '점유율(단위)수량
''
''        spdView.Col = 6
''        sData = sData & Space(3) & Right(Space(10) & spdView.Text, 10) & " "       '점유율(단위)금액
''
''        spdView.Col = 7
''        sData = sData & Space(10) & Right(Space(10) & spdView.Text, 10) & " "      '점유율(전체)수량
''
''        spdView.Col = 8
''        sData = sData & Space(3) & Right(Space(10) & spdView.Text, 10) & " "       '점유율(전체)금액
''
''        iRow = iRow + 4
''        Printer.CurrentY = iRow
''        Printer.CurrentX = 0
''        Printer.Print sData
''
''        If iRow > 270 Then
''            iRow = iRow + 1
''            Printer.Line (0, iRow + 3)-(240, iRow + 3)
''
''            memRow = i + 1
''            iRow = 0
''
''            Printer.NewPage
''            GoTo PrintHead
''        End If
''    Next i

    iRow = iRow + 1
    Printer.Line (0, iRow + 3)-(240, iRow + 3)

    iRow = iRow + 1
    Printer.Line (0, iRow + 3)-(240, iRow + 3)

    Printer.EndDoc
End Sub

