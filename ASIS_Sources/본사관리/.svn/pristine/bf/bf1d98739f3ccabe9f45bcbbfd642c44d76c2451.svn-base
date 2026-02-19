VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form P_02012_01 
   Caption         =   "입고품목 CHECK"
   ClientHeight    =   10695
   ClientLeft      =   1275
   ClientTop       =   1920
   ClientWidth     =   14205
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_02012_01.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10695
   ScaleWidth      =   14205
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   10695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14205
      _ExtentX        =   25056
      _ExtentY        =   18865
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_02012_01.frx":058A
      Begin Threed.SSPanel SSPanel1 
         Height          =   465
         Left            =   15
         TabIndex        =   2
         Top             =   10215
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   820
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   2
            Left            =   7890
            TabIndex        =   5
            Top             =   60
            Width           =   1455
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   1
            Left            =   4710
            TabIndex        =   4
            Top             =   60
            Width           =   1455
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            TabIndex        =   3
            Top             =   60
            Width           =   1455
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
            Caption         =   "검 품 수 량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   3240
            TabIndex        =   7
            Top             =   60
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
            Index           =   3
            Left            =   6420
            TabIndex        =   8
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "미입고수량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   9405
         Index           =   1
         Left            =   15
         TabIndex        =   1
         Top             =   795
         Width           =   14175
         _Version        =   524288
         _ExtentX        =   25003
         _ExtentY        =   16589
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
         SpreadDesigner  =   "P_02012_01.frx":05FC
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panInput 
         Height          =   765
         Left            =   15
         TabIndex        =   9
         Top             =   15
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   1349
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Left            =   1530
            Style           =   2  '드롭다운 목록
            TabIndex        =   10
            Top             =   405
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1515
            TabIndex        =   11
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   60948480
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   12
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "입 고 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   7
            Left            =   60
            TabIndex        =   13
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
   End
End
Attribute VB_Name = "P_02012_01"
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
    Dim i As Integer

'    cmdBtn(0).Enabled = True
'    cmdBtn(5).Enabled = True
'
'    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
    dtInput(0).Value = P_02012.dtInput(0).Value
    
    i = P_02012.ActiveControl.Index
    
'    If P_02012_Flag = False Then
        ReDim sValue(2)
        
        sValue(0) = "1"
        sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
        
        P_02012.spdView(i).Row = P_02012.spdView(i).ActiveRow
        P_02012.spdView(i).Col = 1
        
        sValue(2) = Mid(P_02012.spdView(i).Text, 2, 3)
            
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_02012_01", sValue(), Err_Num, Err_Dec)
        
        spdView(1).MaxCols = RS01.Fields.Count
        spdView(1).MaxRows = RS01.RecordCount
        
'        Call spdDisplay2(RS01)
        Call GetColWidth(REG_App, Me.Name, spdView(1))
        
        P_02012_Flag = True
'    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    With spdView(1)
        .ColsFrozen = 1 '틀고정
        .Row = -1
        
        .Col = 1
        .ColWidth(1) = 10
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
        
        .Col = 2
        .ColWidth(2) = 10
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
        
        .Col = 3
        .ColWidth(3) = 10
        .CellType = CellTypeDate
        .TypeDateCentury = True
        .TypeDateFormat = TypeDateFormatYYMMDD
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
    
        .Col = 4
        .ColWidth(4) = 10
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
        
        .Col = 5
        .ColWidth(5) = 10
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
        
        .Col = 6
        .ColWidth(6) = 12
        .CellType = CellTypeFloat
        .TypeFloatSeparator = True
        .TypeFloatDecimalPlaces = 0
        .TypeVAlign = TypeVAlignCenter
        
        .Col = 7
        .ColWidth(7) = 10
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
    
        .Col = 8
        .ColWidth(8) = 20
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_02012_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i As Integer
    Dim lAmt As Long
    
    ReDim sValue(2)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Mid(cboInput.Text, 2, 6)
        
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_02012_01", sValue(), Err_Num, Err_Dec)
    
    spdView(1).MaxCols = RS01.Fields.Count
    spdView(1).MaxRows = RS01.RecordCount
    
    'Call spdDisplay2(RS01)
    Call fpSpread_Display(spdView(1), RS01)
    Call GetColWidth(REG_App, Me.Name, spdView(1))
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

'Private Sub spdDisplay2(Rs As ADODB.Recordset)
'    Call fpSpread_Display(spdView(1), Rs)
'End Sub

Public Sub DataPrint()

End Sub
