VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form P_02011_01 
   Caption         =   "입고자료 CHECK"
   ClientHeight    =   11580
   ClientLeft      =   1275
   ClientTop       =   1920
   ClientWidth     =   14820
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_02011_01.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11580
   ScaleWidth      =   14820
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   11580
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14820
      _ExtentX        =   26141
      _ExtentY        =   20426
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_02011_01.frx":058A
      Begin Threed.SSPanel SSPanel1 
         Height          =   450
         Left            =   15
         TabIndex        =   7
         Top             =   11115
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   794
         _Version        =   262144
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   2
            Left            =   8205
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   60
            Width           =   1455
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   1
            Left            =   4875
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   60
            Width           =   1455
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   0
            Left            =   1545
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   60
            Width           =   1455
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   75
            TabIndex        =   11
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
            Left            =   3405
            TabIndex        =   12
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
            Left            =   6735
            TabIndex        =   13
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
      Begin Threed.SSPanel panInput 
         Height          =   765
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   1349
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.ComboBox cboInput 
            Height          =   315
            Left            =   1515
            Style           =   2  '드롭다운 목록
            TabIndex        =   2
            Top             =   420
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1515
            TabIndex        =   3
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   65077248
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
            BackColor       =   16777215
            Caption         =   "입 고 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   7
            Left            =   60
            TabIndex        =   5
            Top             =   405
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            BackColor       =   16777215
            Caption         =   "대 리 점 명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10305
         Index           =   1
         Left            =   15
         TabIndex        =   6
         Top             =   795
         Width           =   14790
         _Version        =   524288
         _ExtentX        =   26088
         _ExtentY        =   18177
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
         SpreadDesigner  =   "P_02011_01.frx":05FC
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_02011_01"
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
'
'    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
    Dim i As Integer
    Dim iSu As Integer
    
    dtInput(0).Value = P_02011.dtInput(0).Value
    
    i = P_02011.ActiveControl.Index
        
    If P_02011_01_Flag = False Then
        ReDim sValue(2)
        
        sValue(0) = "0"
        sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
        
        P_02011.spdView(i).Row = P_02011.spdView(i).ActiveRow
        P_02011.spdView(i).Col = 1
        
        sValue(2) = Mid(P_02011.spdView(i).Text, 2, 3)
            
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_02011_01", sValue(), Err_Num, Err_Dec)
        
        spdView(1).MaxCols = RS01.Fields.Count
        spdView(1).MaxRows = RS01.RecordCount
        
        Call spdDisplay(RS01)
        Call GetColWidth(REG_App, Me.Name, spdView(1))
        
        For i = 1 To spdView(1).MaxRows
            spdView(1).Row = i
            spdView(1).Col = 3
            If spdView(1).Text <> "" Then
                iSu = iSu + 1
            End If
        Next i
        
        txtInput(0).Text = spdView(1).MaxRows
        txtInput(1).Text = iSu
        txtInput(2).Text = spdView(1).MaxRows - iSu
        
        P_02011_01_Flag = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_02011_01_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i As Integer
    Dim iSu As Integer

    ReDim sValue(2)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    
    sValue(2) = Mid(cboInput.Text, 2, 6)
        
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_02011_01", sValue(), Err_Num, Err_Dec)
    
    spdView(1).MaxCols = RS01.Fields.Count
    spdView(1).MaxRows = RS01.RecordCount
    
    Call spdDisplay(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView(1))

    For i = 1 To spdView(1).MaxRows
        spdView(1).Row = i
        spdView(1).Col = 3
        If spdView(1).Text <> "" Then
            iSu = iSu + 1
        End If
    Next i
    
    txtInput(0).Text = spdView(1).MaxRows
    txtInput(1).Text = iSu
    txtInput(2).Text = spdView(1).MaxRows - iSu
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub spdDisplay(Rs As ADODB.Recordset)
    
    Call fpSpread_Display(spdView, Rs)
    
    spdView(1).ColsFrozen = 1 '틀고정
    
    spdView(1).Row = -1
    
    spdView(1).Col = 1
    spdView(1).ColWidth(1) = 10
    spdView(1).CellType = CellTypeEdit
    spdView(1).TypeVAlign = TypeVAlignCenter
    spdView(1).TypeHAlign = TypeHAlignCenter
    
    spdView(1).Col = 2
    spdView(1).ColWidth(2) = 10
    spdView(1).CellType = CellTypeDate
    spdView(1).TypeDateCentury = True
    spdView(1).TypeDateFormat = TypeDateFormatYYMMDD
    spdView(1).TypeVAlign = TypeVAlignCenter
    spdView(1).TypeHAlign = TypeHAlignCenter

    spdView(1).Col = 3
    spdView(1).ColWidth(3) = 10
    spdView(1).CellType = CellTypeDate
    spdView(1).TypeDateCentury = True
    spdView(1).TypeDateFormat = TypeDateFormatYYMMDD
    spdView(1).TypeVAlign = TypeVAlignCenter
    spdView(1).TypeHAlign = TypeHAlignCenter
    
    spdView(1).Col = 4
    spdView(1).ColWidth(4) = 10
    spdView(1).CellType = CellTypeEdit
    spdView(1).TypeVAlign = TypeVAlignCenter
    spdView(1).TypeHAlign = TypeHAlignLeft
    
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
    spdView(1).ColWidth(8) = 12
    spdView(1).CellType = CellTypeEdit
    spdView(1).TypeVAlign = TypeVAlignCenter
    spdView(1).TypeHAlign = TypeHAlignLeft

    spdView(1).Col = 9
    spdView(1).ColWidth(9) = 10
    spdView(1).CellType = CellTypeEdit
    spdView(1).TypeVAlign = TypeVAlignCenter
    spdView(1).TypeHAlign = TypeHAlignLeft
End Sub

Public Sub DataPrint()

End Sub

