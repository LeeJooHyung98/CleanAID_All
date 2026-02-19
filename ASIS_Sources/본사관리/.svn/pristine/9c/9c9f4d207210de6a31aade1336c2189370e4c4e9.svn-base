VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form P_03010_01 
   Caption         =   "가출고 관리"
   ClientHeight    =   10200
   ClientLeft      =   1935
   ClientTop       =   2700
   ClientWidth     =   16650
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10200
   ScaleWidth      =   16650
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   10200
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16650
      _ExtentX        =   29369
      _ExtentY        =   17992
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_03010_01.frx":0000
      Begin Threed.SSPanel SSPanel 
         Height          =   405
         Left            =   15
         TabIndex        =   9
         Top             =   9780
         Width           =   16620
         _ExtentX        =   29316
         _ExtentY        =   714
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   0
            Left            =   1665
            TabIndex        =   12
            Top             =   45
            Width           =   1455
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   1
            Left            =   4845
            TabIndex        =   11
            Top             =   45
            Width           =   1455
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   2
            Left            =   8025
            TabIndex        =   10
            Top             =   45
            Width           =   1455
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   45
            TabIndex        =   13
            Top             =   45
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "검 품 수 량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   3225
            TabIndex        =   14
            Top             =   45
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "입 고 수 량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   3
            Left            =   6405
            TabIndex        =   15
            Top             =   45
            Width           =   1635
            _ExtentX        =   2884
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
         Width           =   16620
         _ExtentX        =   29316
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
            Index           =   0
            Left            =   1530
            TabIndex        =   3
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   21364736
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
            Caption         =   "출 고 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   4830
            TabIndex        =   5
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   21364736
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   4
            Left            =   60
            TabIndex        =   6
            Top             =   405
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "대 리 점 명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin VB.Label Label1 
            Alignment       =   2  '가운데 맞춤
            Caption         =   "~"
            Height          =   255
            Left            =   4575
            TabIndex        =   7
            Top             =   120
            Width           =   195
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   8970
         Index           =   1
         Left            =   15
         TabIndex        =   8
         Top             =   795
         Width           =   16620
         _Version        =   524288
         _ExtentX        =   29316
         _ExtentY        =   15822
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
         SpreadDesigner  =   "P_03010_01.frx":0072
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_03010_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cboInput_Click()
    Call Data_Display
End Sub

Private Sub dtInput_Change(Index As Integer)
    Call Data_Display
End Sub

Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    cmdBtn(5).Enabled = True
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
    Dim i As Integer
    
    dtInput(0).Value = P_03010.dtInput(0).Value
    dtInput(1).Value = P_03010.dtInput(1).Value
    
    i = P_03010.ActiveControl.Index
    
    If P_03010_01_Flag = False Then
        ReDim sValue(3)
        
        sValue(0) = "0"
        
        P_03010.spdView(i).Row = P_03010.spdView(i).ActiveRow
        P_03010.spdView(i).Col = 1
        
        sValue(1) = Mid(P_03010.spdView(i).Text, 2, 3)
        sValue(2) = Format(dtInput(0).Value, "YYYY-MM-DD")
        sValue(3) = Format(dtInput(1).Value, "YYYY-MM-DD")
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_03010_01", sValue(), Err_Num, Err_Dec)
        
        spdView(1).MaxCols = RS01.Fields.Count
        spdView(1).MaxRows = RS01.RecordCount
        
        Call spdDisplay(RS01)
        Call GetColWidth(REG_App, Me.Name, spdView(1))
        
        P_03010_01_Flag = True
    End If
End Sub

Private Sub spdDisplay(Rs As ADODB.Recordset)
    
    Call fpSpread_Display(spdView(1), Rs)
    
    spdView(1).ColsFrozen = 1 '틀고정
    
    spdView(1).Row = -1
    
    spdView(1).Col = 1
    spdView(1).ColWidth(1) = 6
    spdView(1).CellType = CellTypeEdit
    spdView(1).TypeVAlign = TypeVAlignCenter
    spdView(1).TypeHAlign = TypeHAlignCenter
    
    spdView(1).Col = 2
    spdView(1).ColWidth(2) = 10
    spdView(1).CellType = CellTypeEdit
    spdView(1).TypeVAlign = TypeVAlignCenter
    spdView(1).TypeHAlign = TypeHAlignLeft
    
    spdView(1).Col = 3
    spdView(1).ColWidth(3) = 8
    spdView(1).CellType = CellTypeEdit
    spdView(1).TypeVAlign = TypeVAlignCenter
    spdView(1).TypeHAlign = TypeHAlignCenter
    
    spdView(1).Col = 4
    spdView(1).ColWidth(4) = 10
    spdView(1).CellType = CellTypeEdit
    spdView(1).TypeVAlign = TypeVAlignCenter
    spdView(1).TypeHAlign = TypeHAlignLeft
    
    spdView(1).Col = 5
    spdView(1).ColWidth(5) = 25
    spdView(1).CellType = CellTypeEdit
    spdView(1).TypeVAlign = TypeVAlignCenter
    spdView(1).TypeHAlign = TypeHAlignLeft
    
    spdView(1).Col = 6
    spdView(1).ColWidth(6) = 12
    spdView(1).CellType = CellTypeFloat
    spdView(1).TypeFloatSeparator = True
    spdView(1).TypeFloatDecimalPlaces = 0
    spdView(1).TypeVAlign = TypeVAlignCenter

    spdView(1).Col = 7
    spdView(1).ColWidth(7) = 15
    spdView(1).CellType = CellTypeEdit
    spdView(1).TypeVAlign = TypeVAlignCenter
    spdView(1).TypeHAlign = TypeHAlignLeft

    spdView(1).Col = 8
    spdView(1).ColWidth(8) = 6
    spdView(1).CellType = CellTypeEdit
    spdView(1).TypeVAlign = TypeVAlignCenter
    spdView(1).TypeHAlign = TypeHAlignLeft

    spdView(1).Col = 9
    spdView(1).ColWidth(9) = 10
    spdView(1).CellType = CellTypeEdit
    spdView(1).TypeVAlign = TypeVAlignCenter
    spdView(1).TypeHAlign = TypeHAlignLeft
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_03010_01_Flag = False
End Sub

Public Sub Data_Display()
    Dim i As Integer
    Dim lAmt As Long
    
    ReDim sValue(3)
    
    sValue(0) = "0"
    
    sValue(1) = Mid(cboInput.Text, 2, 3)
    sValue(2) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(3) = Format(dtInput(1).Value, "YYYY-MM-DD")
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_03010_01", sValue(), Err_Num, Err_Dec)
    
    spdView(1).MaxCols = RS01.Fields.Count
    spdView(1).MaxRows = RS01.RecordCount
    
    Call spdDisplay(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView(1))
End Sub

Public Sub DataPrint()

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
