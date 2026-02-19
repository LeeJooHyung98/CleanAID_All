VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form P_05014 
   Caption         =   "쿠폰 사용 현황"
   ClientHeight    =   12180
   ClientLeft      =   750
   ClientTop       =   2715
   ClientWidth     =   16125
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_05014.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12180
   ScaleWidth      =   16125
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   12180
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16125
      _ExtentX        =   28443
      _ExtentY        =   21484
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_05014.frx":058A
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10935
         Index           =   0
         Left            =   15
         TabIndex        =   1
         Top             =   1230
         Width           =   4185
         _Version        =   524288
         _ExtentX        =   7382
         _ExtentY        =   19288
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   3
         MaxRows         =   37
         ScrollBars      =   2
         SpreadDesigner  =   "P_05014.frx":063C
         UserResize      =   1
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10935
         Index           =   1
         Left            =   4215
         TabIndex        =   2
         Top             =   1230
         Width           =   4140
         _Version        =   524288
         _ExtentX        =   7302
         _ExtentY        =   19288
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   3
         MaxRows         =   37
         ScrollBars      =   2
         SpreadDesigner  =   "P_05014.frx":0B82
         UserResize      =   1
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10935
         Index           =   2
         Left            =   8370
         TabIndex        =   3
         Top             =   1230
         Width           =   7740
         _Version        =   524288
         _ExtentX        =   13652
         _ExtentY        =   19288
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   3
         MaxRows         =   37
         ScrollBars      =   0
         SpreadDesigner  =   "P_05014.frx":109F
         UserResize      =   1
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   4
         Top             =   435
         Width           =   16095
         _ExtentX        =   28390
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Index           =   0
            Left            =   6210
            Style           =   2  '드롭다운 목록
            TabIndex        =   5
            Top             =   60
            Width           =   3015
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   6
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "조 회 월"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   4740
            TabIndex        =   7
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "쿠 폰 구 분"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            TabIndex        =   8
            Top             =   45
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy-MM"
            Format          =   63307779
            CurrentDate     =   36686
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   405
         Left            =   15
         TabIndex        =   9
         Top             =   15
         Width           =   16095
         _ExtentX        =   28390
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
         Caption         =   " 쿠폰 사용 현황 (P_05014)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_05014.frx":15F9
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
End
Attribute VB_Name = "P_05014"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub Form_Activate()
'    cmdBtn(0).Enabled = True
'    cmdBtn(5).Enabled = True
'    cmdBtn(6).Enabled = True
'
'    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
    
    If P_02010_Flag = False Then
        dtInput(0).Value = Date

        P_02010_Flag = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_02010_Flag = False
End Sub

Private Sub Data_Display()
    Dim i As Integer
    Dim j As Integer
    Dim lAmt As Long
    
    On Error GoTo ErrRtn
    
    ReDim sValue(1)
    
    For i = 0 To spdView.Count - 1
        For j = 1 To spdView(i).MaxRows
            spdView(i).Row = j
            spdView(i).Col = -1
            spdView(i).BackColor = vbWhite
            
            spdView(i).Col = 1
            spdView(i).Text = ""
            spdView(i).Col = 2
            spdView(i).Text = ""
            spdView(i).Col = 3
            spdView(i).Text = ""
        Next j
    Next i
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput(0).Value, "yyyymm")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_05014_01", sValue(), Err_Num, Err_Dec)
    
    
    j = 1
    If RS01.Fields.Count <= 0 Then Exit Sub
    Do While Not RS01.EOF
        
        spdView(0).Row = j
        spdView(0).Col = 1
        spdView(0).Text = RS01!코드 & ""
        spdView(0).Col = 2
        spdView(0).Text = RS01!대리점명 & ""
        spdView(0).Col = 3
        spdView(0).Text = RS01!수량 & ""
 
        j = j + 1
        RS01.MoveNext
    Loop
    RS01.Close
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub


Public Sub Data_Display2()
    Dim i As Integer
    Dim j As Integer
    Dim lAmt As Long
    
    On Error GoTo ERR_RTN
    
    ReDim sValue(2)
    
    For j = 1 To spdView(1).MaxRows
        spdView(1).Row = j
        spdView(1).Col = -1
        spdView(1).BackColor = vbWhite
        
        spdView(1).Col = 1
        spdView(1).Text = ""
        spdView(1).Col = 2
        spdView(1).Text = ""
        spdView(1).Col = 3
        spdView(1).Text = ""
    Next j

    spdView(0).Row = spdView(0).ActiveRow
    spdView(0).Col = 1
    
    sValue(0) = "0"
    sValue(1) = spdView(0).Text
    sValue(2) = Format(dtInput(0).Value, "yyyymm")
    
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_05014_02", sValue(), Err_Num, Err_Dec)
    
    j = 1
    If RS01.Fields.Count <= 0 Then Exit Sub
    Do While Not RS01.EOF
        spdView(1).Row = j
        
        spdView(1).Col = 1
        spdView(1).Text = RS01!일자 & ""
        spdView(1).Col = 2
        spdView(1).Text = RS01!수량 & ""
        spdView(1).Col = 3
        spdView(1).Text = "" 'RS01!수량 & ""
        
        j = j + 1
        RS01.MoveNext
    Loop
    RS01.Close
    Exit Sub
    
ERR_RTN:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Data_Display2 of Form P_05014"

End Sub


Public Sub Data_Display3()
    Dim i As Integer
    Dim j As Integer
    Dim lAmt As Long
    
    On Error GoTo ERR_RTN
    
    ReDim sValue(2)
    
    For j = 1 To spdView(2).MaxRows
        spdView(2).Row = j
        spdView(2).Col = -1
        spdView(2).BackColor = vbWhite
        
        spdView(2).Col = 1
        spdView(2).Text = ""
        spdView(2).Col = 2
        spdView(2).Text = ""
        spdView(2).Col = 3
        spdView(2).Text = ""
    Next j

    spdView(0).Row = spdView(0).ActiveRow
    spdView(0).Col = 1
    
    sValue(0) = "0"
    sValue(1) = spdView(0).Text
    
    spdView(1).Row = spdView(1).ActiveRow
    spdView(1).Col = 1
    sValue(2) = Format(spdView(1).Text, "YYYY-MM-DD")
    
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_05014_03", sValue(), Err_Num, Err_Dec)
    
    j = 1
    If RS01.Fields.Count <= 0 Then Exit Sub
    Do While Not RS01.EOF
        spdView(2).Row = j
        
        spdView(2).Col = 1: spdView(2).Text = RS01!쿠폰번호 & ""
        spdView(2).Col = 2: spdView(2).Text = RS01!성명 & ""
        spdView(2).Col = 3: spdView(2).Text = RS01!금액 & ""
        
        j = j + 1
        RS01.MoveNext
    Loop
    RS01.Close
    Exit Sub
    
ERR_RTN:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Data_Display3 of Form P_05014"

End Sub

 
Public Sub DataPrint()
'    Dim i As Integer
'    Dim ReportFP As String
'    Dim ReportFile As String
'
'    ReportFP = GetIniStr("REPORT", "FilePath", "", m_iniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    Call PrintDesc
'
'    Dim ii As Integer
'    For ii = 0 To 30
'        P_00000.crPrint.Formulas(ii) = ""
'    Next
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'    P_00000.crPrint.Formulas(0) = "입고일자 = '" & Format(dtInput(0).Value, "YYYY-MM-DD") & "'"
'
'    Call ReportPrint(ReportFile, "1")
End Sub

Public Sub DataScreen()
'    Dim i As Integer
'    Dim ReportFP As String
'    Dim ReportFile As String
'
'    ReportFP = GetIniStr("REPORT", "FilePath", "", m_iniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    Call PrintDesc
'
'    Dim ii As Integer
'    For ii = 0 To 30
'        P_00000.crPrint.Formulas(ii) = ""
'    Next
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'    P_00000.crPrint.Formulas(0) = "입고일자 = '" & Format(dtInput(0).Value, "YYYY-MM-DD") & "'"
'
'    Call ReportPrint(ReportFile, "2")
End Sub

Private Sub PrintDesc()
    Dim i As Integer
    Dim j As Integer
    
    Dim TempTag As String
    Dim TempText As String
    Dim TempFP As String
    Dim TempFile As String
    
    On Error GoTo FileError:
    
    TempFP = GetIniStr("REPORT", "TempPath", "", m_iniFile)
    TempFile = TempFP & "\Temp.txt"
    
    Open TempFile For Output As #1
    
    TempText = ""
    
    For j = 0 To 2
        For i = 1 To spdView(j).MaxRows - 1
            spdView(j).Row = i
            
            spdView(j).Col = 1
            If spdView(j).Text = "" Then
                Close #1
                Exit Sub
            End If
            
            spdView(j).Col = 3
            TempTag = spdView(j).Text
            
            spdView(j).Col = 2
            If spdView(j).BackColor <> &HD8FCFE Then
                spdView(j).Col = 1
                TempText = TempText & "  " & LeftH(spdView(j).Text & Space(20), 20)
            Else
                spdView(j).Col = 1
                TempText = TempText & " *" & LeftH(spdView(j).Text & Space(20), 20)
            End If
            
            spdView(j).Col = 2
            TempText = TempText & "  " & spdView(j).Text
            spdView(j).Col = 3
            TempText = TempText & "  " & spdView(j).Text
            
            If i Mod 2 = 0 Then
                Print #1, TempText
                TempText = ""
            End If
        Next i
    Next j
    
    Close #1
    
FileError:
    If Err.Number = 55 Then
        Resume Next
    End If
End Sub

Private Sub spdView_DblClick(Index As Integer, ByVal Col As Long, ByVal Row As Long)
    
    
    Select Case Index
    
        '-------------------------------------------------------------------------------
        ' 지사 출력 내용을 클릭한경우 지사에 속한 체인점의 내용을 조회할 수 있도록 처리
        Case 0
            Call Data_Display2
            Exit Sub
            
        Case 1
            Call Data_Display3
            Exit Sub
        
        Case Else
            Exit Sub
        
    End Select
End Sub
