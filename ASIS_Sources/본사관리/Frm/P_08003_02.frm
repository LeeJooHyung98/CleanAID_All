VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form P_08003_02 
   BorderStyle     =   1  '단일 고정
   Caption         =   "대리점품목 생성"
   ClientHeight    =   8055
   ClientLeft      =   1620
   ClientTop       =   1920
   ClientWidth     =   8445
   Icon            =   "P_08003_02.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   8445
   StartUpPosition =   2  '화면 가운데
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   14208
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_08003_02.frx":058A
      Begin Threed.SSPanel panTitle 
         Height          =   765
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   1349
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin Threed.SSCommand cmdBtn 
            Height          =   375
            Index           =   2
            Left            =   4920
            TabIndex        =   2
            Top             =   30
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   661
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "전체선택"
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   3
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
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
            CheckBox        =   -1  'True
            Format          =   60227584
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   4
            Top             =   60
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "적 용 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSCommand cmdBtn 
            Height          =   375
            Index           =   0
            Left            =   6060
            TabIndex        =   5
            Top             =   30
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   661
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "작업시작"
         End
         Begin Threed.SSCommand cmdBtn 
            Height          =   375
            Index           =   1
            Left            =   7200
            TabIndex        =   6
            Top             =   30
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "작업종료"
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   7245
         Left            =   15
         TabIndex        =   7
         Top             =   795
         Width           =   8415
         _Version        =   524288
         _ExtentX        =   14843
         _ExtentY        =   12779
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
         SpreadDesigner  =   "P_08003_02.frx":05DC
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_08003_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Dim sDownPath As String

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 0          ' 작업시작
            Call DataSave
        Case 1          ' 작업종료
            Unload Me
        Case 2
            Dim i As Integer
            
            For i = 1 To spdView.MaxRows
                spdView.Row = i
                spdView.Col = 2
                
                If spdView.Value = True Then
                    spdView.Value = False
                Else
                    spdView.Value = True
                End If
            Next i
    End Select
End Sub

Private Sub dtInput_Change(Index As Integer)
    Call Data_Display
End Sub

Private Sub Form_Activate()
    ReDim sValue(1)
    
    sValue(0) = "0"
    sValue(1) = Format(Now, "YYYY-MM-DD")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_08001_24", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    Call spdDisplay(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView)
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    With spdView
        .ColsFrozen = 1 '틀고정
        .Row = -1
        
        .Col = 1
        .ColWidth(1) = 53
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
        
        .Col = 2
        .ColWidth(2) = 8
        .CellType = CellTypeCheckBox
        .TypeCheckCenter = True
        .Value = False
    End With
    
    dtInput(0).Value = Date
    
    sDownPath = GetIniStr("SERVER DATA", "SendPath", "", m_iniFile)
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub DataSave()
    Dim i As Integer
    Dim sPrice As String
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        spdView.Col = 2
        
        If spdView.Value = True Then
            ReDim sValue(2)
            
            sValue(0) = "0"
            sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
            
            spdView.Col = 1
            sValue(2) = Mid(spdView.Text, 2, 3)
            
            Open sDownPath & "\" & sValue(1) & sValue(2) & ".Dat" For Output As #1
            
            Set RS01 = New ADODB.Recordset
            Set RS01 = ExecPro("SP_08001_25", sValue(), Err_Num, Err_Dec)
            
            Do While Not RS01.EOF
                sPrice = "01234567"
                
                RSet sPrice = RS01!단가 & ""
                
                Print #1, " " & RS01!품목코드;
                Print #1, " " & sPrice;
                Print #1, " " & RS01!품명
            
                RS01.MoveNext
            Loop
        End If
    Next i
    
    Close #1
    
    MsgBox "데이터가 정상적으로 처리되었습니다.", vbInformation
End Sub

Private Sub spdDisplay(Rs As ADODB.Recordset)
    Call fpSpread_Display(spdView, Rs)
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    ReDim sValue(1)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_08001_24", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    Call spdDisplay(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView)
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub


