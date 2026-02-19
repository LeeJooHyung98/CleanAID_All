VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form P_08001_01 
   BorderStyle     =   1  '단일 고정
   Caption         =   "할인자료 송신"
   ClientHeight    =   8055
   ClientLeft      =   1620
   ClientTop       =   1920
   ClientWidth     =   8445
   Icon            =   "P_08001_01.frx":0000
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
      PaneTree        =   "P_08001_01.frx":058A
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
            Index           =   0
            Left            =   4980
            TabIndex        =   2
            Top             =   30
            Width           =   1635
            _ExtentX        =   2884
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
            Caption         =   "작 업 시 작"
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
            Format          =   61079552
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
            Caption         =   "출 고 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSCommand cmdBtn 
            Height          =   375
            Index           =   1
            Left            =   6660
            TabIndex        =   5
            Top             =   30
            Width           =   1635
            _ExtentX        =   2884
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
            Caption         =   "작 업 종 료"
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   7245
         Left            =   15
         TabIndex        =   6
         Top             =   795
         Width           =   8415
         _Version        =   524288
         _ExtentX        =   14843
         _ExtentY        =   12779
         _StockProps     =   64
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
         SpreadDesigner  =   "P_08001_01.frx":05DC
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_08001_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cmdBtn_Click(Index As Integer)
    Select Case Index
        Case 0          ' 작업시작
            Call DataSave
        Case 1          ' 작업종료
            Unload Me
    End Select
End Sub

Private Sub dtInput_Change(Index As Integer)
    Call Data_Display
End Sub

Private Sub Form_Activate()
    ReDim sValue(1)
    
    sValue(0) = "1"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_08001_20", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    Call spdDisplay(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView)
    Call Data_Display
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    dtInput(0).Value = Date
    
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub DataSave()
    Dim i As Integer
    Dim sPrice As String
    Dim sRatio As String
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        spdView.Col = 4
        
        If spdView.Value = True Then
            ReDim sValue(2)
            
            sValue(0) = "0"
            sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
            
            spdView.Col = 1
            sValue(2) = Mid(spdView.Text, 2, 3)
            
            Open "A:\Sale" & sValue(2) & ".Dat" For Output As #1
            
            Set RS01 = New ADODB.Recordset
            Set RS01 = ExecPro("SP_08001_21", sValue(), Err_Num, Err_Dec)
            
            Do While Not RS01.EOF
                sPrice = "01234567"
                sRatio = "01"
                
                RSet sPrice = RS01!단가 & ""
                RSet sRatio = RS01!할인율 & ""
                
                Print #1, " " & RS01!할인시작일 & " " & RS01!할인종료일;
                Print #1, " " & RS01!품목코드;
                Print #1, " " & sPrice;
                Print #1, " " & IIf(sRatio = " 0", "  ", sRatio);
                Print #1, " " & RS01!품명
            
                RS01.MoveNext
            Loop
        End If
    Next i
    
    Close #1
End Sub

Private Sub spdDisplay(Rs As ADODB.Recordset)

    
    Call fpSpread_Display(spdView, Rs)
    
    spdView.ColsFrozen = 1 '틀고정
    
    spdView.Row = -1
    
    spdView.Col = 1
    spdView.ColWidth(1) = 20
    spdView.CellType = CellTypeStaticText
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft
    
    spdView.Col = 2
    spdView.ColWidth(2) = 10
    spdView.CellType = CellTypeDate
    spdView.TypeDateCentury = True
    spdView.TypeDateFormat = TypeDateFormatYYMMDD
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft
    
    spdView.Col = 3
    spdView.ColWidth(3) = 10
    spdView.CellType = CellTypeDate
    spdView.TypeDateCentury = True
    spdView.TypeDateFormat = TypeDateFormatYYMMDD
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignLeft
    
    spdView.Col = 4
    spdView.ColWidth(4) = 5
    spdView.CellType = CellTypeCheckBox
    spdView.TypeCheckCenter = True
    spdView.Value = False
    
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    ReDim sValue(1)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_08001_20", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    Call spdDisplay(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView)
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

