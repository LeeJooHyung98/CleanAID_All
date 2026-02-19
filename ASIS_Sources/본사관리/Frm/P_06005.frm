VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form P_06005 
   Caption         =   "사고 품목별 평균 내용연수 관리"
   ClientHeight    =   11550
   ClientLeft      =   1620
   ClientTop       =   1920
   ClientWidth     =   16185
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_06005.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11550
   ScaleWidth      =   16185
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   11550
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16185
      _ExtentX        =   28549
      _ExtentY        =   20373
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_06005.frx":058A
      Begin Threed.SSPanel panInput 
         Height          =   765
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   16155
         _ExtentX        =   28496
         _ExtentY        =   1349
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10740
         Left            =   15
         TabIndex        =   2
         Top             =   795
         Width           =   16155
         _Version        =   524288
         _ExtentX        =   28496
         _ExtentY        =   18944
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
         SpreadDesigner  =   "P_06005.frx":05DC
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_06005"
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
'    cmdBtn(1).Enabled = True
'    cmdBtn(2).Enabled = True
'    cmdBtn(3).Enabled = True
'    cmdBtn(4).Enabled = True
'
'    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    If P_06005_Flag = False Then
        ReDim sValue(1)
        
        sValue(0) = "1"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_06005_00", sValue(), Err_Num, Err_Dec)
        
        spdView.MaxCols = RS01.Fields.Count
        spdView.MaxRows = RS01.RecordCount
        
        Call spdDisplay(RS01)
        Call GetColWidth(REG_App, Me.Name, spdView)
        
        P_06005_Flag = True
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub spdDisplay(Rs As ADODB.Recordset)

    Call fpSpread_Display(spdView, Rs)
    
    spdView.ColsFrozen = 1 '틀고정
    
    spdView.Row = -1
    
    spdView.Col = 1
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter
    
    spdView.Col = 2
    spdView.CellType = CellTypeEdit
    spdView.TypeVAlign = TypeVAlignCenter
    spdView.TypeHAlign = TypeHAlignCenter
    
    spdView.Col = 3
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
    
    spdView.Col = 4
    spdView.CellType = CellTypeFloat
    spdView.TypeFloatSeparator = True
    spdView.TypeFloatDecimalPlaces = 0
    spdView.TypeVAlign = TypeVAlignCenter
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_06005_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    ReDim sValue(1)
    
    sValue(0) = "0"
        
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_06005_00", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    Call spdDisplay(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView)
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Public Sub DataAdd()
    spdView.MaxRows = spdView.MaxRows + 1
    
    spdView.Row = spdView.MaxRows
    spdView.Col = 1
    spdView.Action = ActionActiveCell
    spdView.Lock = False
    
    spdView.SetFocus
End Sub

Public Sub DataSave()
    Dim i As Integer
    
    ReDim sValue(3)
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        spdView.Col = 1
        sValue(0) = spdView.Text
        
        spdView.Col = 2
        sValue(1) = spdView.Text
        
        spdView.Col = 3
        If IsNull(spdView.Value) = True Then
            sValue(2) = 0
        Else
            sValue(2) = spdView.Value
        End If
        
        spdView.Col = 4
        If spdView.Value = "" Then
            sValue(3) = 0
        Else
            sValue(3) = spdView.Value
        End If
        
        If Trim(sValue(0)) = "" Then
            Exit Sub
        End If
        
        Call ExecPro("SP_06005_01", sValue(), Err_Num, Err_Dec)
        
        If Err_Num <> 0 Then
            MsgBox "[" & Err_Num & "] " & Err_Dec, vbInformation
        End If
    Next i

    If Err_Num = 0 Then
        MsgBox "해당되는 데이터가 정상적으로 저장이 되었습니다.", vbInformation
    End If
End Sub

Public Sub DataDelete()
    If MsgBox("해당되는 데이터를 삭제하시겠습니까?", vbYesNo + vbInformation + vbDefaultButton2, "데이터 삭제") = vbYes Then
    
        ReDim sValue(0)
        
        spdView.Row = spdView.ActiveRow
        spdView.Col = 1
        sValue(0) = spdView.Text
        
        Call ExecPro("SP_06005_02", sValue(), Err_Num, Err_Dec)
        
        If Err_Num = 0 Then
            spdView.Row = spdView.ActiveRow
            spdView.Action = ActionDeleteRow
            DoEvents
            
            MsgBox "해당되는 데이터가 정상적으로 삭제가 되었습니다.", vbInformation
        End If
    End If
End Sub

Public Sub DataCancel()
    Call Data_Display
End Sub



