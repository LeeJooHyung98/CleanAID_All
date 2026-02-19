VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form P_01012_01 
   BorderStyle     =   1  '단일 고정
   Caption         =   "대리점품목 생성"
   ClientHeight    =   7320
   ClientLeft      =   1620
   ClientTop       =   1920
   ClientWidth     =   8175
   Icon            =   "P_01012_01.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   8175
   StartUpPosition =   1  '소유자 가운데
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7320
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   12912
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_01012_01.frx":058A
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   6525
         Left            =   15
         TabIndex        =   1
         Top             =   780
         Width           =   8145
         _Version        =   524288
         _ExtentX        =   14367
         _ExtentY        =   11509
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
         MaxCols         =   4
         ScrollBars      =   2
         SpreadDesigner  =   "P_01012_01.frx":05DC
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panTitle 
         Height          =   750
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   1323
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin Threed.SSCommand cmdBtn 
            Height          =   375
            Index           =   2
            Left            =   4590
            TabIndex        =   3
            Top             =   75
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
            Left            =   945
            TabIndex        =   4
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
            Format          =   56229888
            CurrentDate     =   36686
         End
         Begin Threed.SSCommand cmdBtn 
            Height          =   375
            Index           =   0
            Left            =   5790
            TabIndex        =   5
            Top             =   60
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
            Left            =   6990
            TabIndex        =   6
            Top             =   60
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
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "적용일자:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   7
            Top             =   120
            Width           =   810
         End
      End
   End
End
Attribute VB_Name = "P_01012_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01      As ADODB.Recordset
Dim sValue()  As String

Dim Err_Num   As Long
Dim Err_Dec   As String

Dim sDownPath As String

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 0: Call Data_Save     ' 작업시작
        Case 1: Unload Me         ' 작업종료
            
        Case 2
            Dim i As Integer
            
            For i = 1 To spdView.MaxRows
                spdView.Row = i
                spdView.Col = 4
                
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
    Set RS01 = ExecPro("SP_01012_01", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    'Call spdDisplay(RS01)
    Call fpSpread_Display(spdView, RS01)
    Call GetColWidth(REG_App, Me.Name, spdView)
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
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
            
            
        .ColsFrozen = 1 '틀고정
        .Row = -1
        
        .Col = 1
        .ColWidth(1) = 20
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
        
        .Col = 2
        .ColWidth(2) = 25
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
        
        .Col = 3
        .ColWidth(3) = 8
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 4
        .ColWidth(4) = 8
        .CellType = CellTypeCheckBox
        .TypeCheckCenter = True
        .Value = False
    End With
    
    'dtInput(0).Value = Date
    
    'sDownPath = GetIniStr("SERVER DATA", "SendPath", "", sIniFile)
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Data_Save()
    Dim i      As Integer
    Dim sPrice As String
    Dim sDir   As String
    
    With spdView
        For i = 1 To .MaxRows
            .Row = i
            .Col = 4
            
            If .Value = True Then
                ReDim sValue(2)
                
                sValue(0) = "0"
                sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
                
                .Col = 2: sValue(2) = Mid(.Text, 2, 6)
                .Col = 1
                
                sDir = App.Path & "\Data"
                If Dir(sDir, vbDirectory) = "" Then
                    '사업장 dir 생성
                    MkDir sDir
                End If
                
                sDir = sDir & "\" & Mid(.Text, 2, 4)
                If Dir(sDir, vbDirectory) = "" Then
                    '사업장 dir 생성
                    MkDir sDir
                End If
                
                'sDir = sDir & "\SendData"
                If Dir(sDir & "\SendData", vbDirectory) = "" Then
                    '사업장 dir 생성
                    MkDir sDir & "\SendData"
                End If
                
                sDir = sDir & "\RecvData"
                If Dir(sDir, vbDirectory) = "" Then
                    '사업장 dir 생성
                    MkDir sDir
                End If
                
                
                .Col = 3
                Open sDir & "\" & sValue(1) & Trim(.Text) & ".Dat" For Output As #1
                
                Set RS01 = New ADODB.Recordset
                Set RS01 = ExecPro("SP_01012_02", sValue(), Err_Num, Err_Dec)
                
                Do While Not RS01.EOF
                    sPrice = "01234567"
                    
                    RSet sPrice = RS01!단가 & ""
                    
                    Print #1, " " & RS01!품목코드;
                    Print #1, " " & sPrice;
                    Print #1, " " & RS01!품명
                
                    RS01.MoveNext
                Loop
                Close #1
            End If
        Next i
    End With
    
    MsgBox "데이터가 정상적으로 처리되었습니다.", vbInformation
End Sub

'Private Sub spdDisplay(RS As ADODB.Recordset)
'    Call fpSpread_Display(spdView, RS)
'End Sub

Private Sub Data_Display()
    ReDim sValue(1)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01012_01", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    'Call spdDisplay(RS01)
    Call fpSpread_Display(spdView, RS01)
    Call GetColWidth(REG_App, Me.Name, spdView)
End Sub

