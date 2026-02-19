VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form P_01004_A_01 
   BorderStyle     =   1  '단일 고정
   Caption         =   "할인자료 생성"
   ClientHeight    =   11820
   ClientLeft      =   1620
   ClientTop       =   1920
   ClientWidth     =   14940
   Icon            =   "P_01004_A_01.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11820
   ScaleWidth      =   14940
   StartUpPosition =   1  '소유자 가운데
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   11820
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14940
      _ExtentX        =   26353
      _ExtentY        =   20849
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_01004_A_01.frx":058A
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10575
         Left            =   15
         TabIndex        =   1
         Top             =   1230
         Width           =   14910
         _Version        =   524288
         _ExtentX        =   26300
         _ExtentY        =   18653
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
         SpreadDesigner  =   "P_01004_A_01.frx":05FC
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panTitle 
         Height          =   780
         Left            =   15
         TabIndex        =   2
         Top             =   435
         Width           =   14910
         _ExtentX        =   26300
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
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
            Format          =   56557568
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
            Left            =   8040
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
            Left            =   9180
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
         Begin Threed.SSCommand cmdBtn 
            Height          =   375
            Index           =   2
            Left            =   6900
            TabIndex        =   7
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
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   405
         Left            =   15
         TabIndex        =   8
         Top             =   15
         Width           =   14910
         _ExtentX        =   26300
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
         Caption         =   " 할인자료 생성 (P_01004_A_01)"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_01004_A_01.frx":0A64
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
End
Attribute VB_Name = "P_01004_A_01"
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
        Case 0: Call DataSave     ' 작업시작
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
    Set RS01 = ExecPro("SP_01004_A_20", sValue(), Err_Num, Err_Dec)
    
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
        .ColWidth(2) = 20
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
        
        .Col = 3
        .ColWidth(3) = 10
        .CellType = CellTypeStaticText
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 4
        .ColWidth(4) = 10
        .CellType = CellTypeDate
        .TypeDateCentury = True
        .TypeDateFormat = TypeDateFormatYYMMDD
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 5
        .ColWidth(5) = 10
        .CellType = CellTypeDate
        .TypeDateCentury = True
        .TypeDateFormat = TypeDateFormatYYMMDD
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 6
        .ColWidth(6) = 8
        .CellType = CellTypeCheckBox
        .TypeCheckCenter = True
        .Value = False
    End With
    
    dtInput(0).Value = Date
    
    sDownPath = GetIniStr("SERVER DATA", "SendPath", "", sIniFile)
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub DataSave()
    Dim i As Integer
    Dim sPrice As String
    Dim sRatio As String
    Dim sDir As String
    Dim hFile As Integer
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        spdView.Col = 6
        
        If spdView.Value = True Then
            sDir = App.Path & "\Data"
            If Dir(sDir, vbDirectory) = "" Then
                '사업장 dir 생성
                MkDir sDir
            End If
            
            spdView.Col = 1
            sDir = sDir & "\" & Mid(spdView.Text, 2, 4)
            
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
        

            ReDim sValue(2)
            
            hFile = FreeFile
            sValue(0) = "0"
            sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
            
            spdView.Col = 2: sValue(2) = Mid(spdView.Text, 2, 6)
            
            spdView.Col = 3
           
            Open sDir & "\Sale" & Trim(spdView.Text) & ".Dat" For Output As #hFile
            
            Set RS01 = New ADODB.Recordset
            Set RS01 = ExecPro("SP_01004_A_21", sValue(), Err_Num, Err_Dec)
            
            Do While Not RS01.EOF
                sPrice = "01234567"
                sRatio = "01"
                
                RSet sPrice = RS01!단가 & ""
                RSet sRatio = RS01!할인율 & ""
                
                Print #hFile, " " & RS01!할인시작일 & " " & RS01!할인종료일;
                Print #hFile, " " & RS01!품목코드;
                Print #hFile, " " & sPrice;
                Print #hFile, " " & IIf(sRatio = " 0", "  ", sRatio);
                Print #hFile, " " & RS01!품명
            
                RS01.MoveNext
            Loop
        End If
        
        Close #hFile
    Next i
    
    MsgBox "데이터가 정상적으로 처리되었습니다." & Space(10), vbInformation
End Sub

'Private Sub spdDisplay(RS As ADODB.Recordset)
'    Call fpSpread_Display(spdView, RS)
'End Sub

Private Sub Data_Display()
    ReDim sValue(1)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_01004_A_20", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    'Call spdDisplay(RS01)
    Call fpSpread_Display(spdView, RS01)
    Call GetColWidth(REG_App, Me.Name, spdView)
End Sub
