VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form P_09003 
   Caption         =   "송신 메일등록"
   ClientHeight    =   11745
   ClientLeft      =   1620
   ClientTop       =   1920
   ClientWidth     =   16845
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_09003.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11745
   ScaleWidth      =   16845
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   11745
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16845
      _ExtentX        =   29713
      _ExtentY        =   20717
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_09003.frx":058A
      Begin Threed.SSPanel panInput 
         Height          =   765
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   16815
         _ExtentX        =   29660
         _ExtentY        =   1349
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            TabIndex        =   2
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   57606144
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   3
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "송 신 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin RichTextLib.RichTextBox rtbInput 
         Height          =   10935
         Left            =   15
         TabIndex        =   4
         Top             =   795
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   19288
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"P_09003.frx":05FC
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   10935
         Left            =   7185
         TabIndex        =   5
         Top             =   795
         Width           =   9645
         _Version        =   524288
         _ExtentX        =   17013
         _ExtentY        =   19288
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
         SpreadDesigner  =   "P_09003.frx":06A1
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_09003"
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
'
'    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    With spdView
        .ColsFrozen = 1 '틀고정
        .Row = -1
        
        .Col = 1
        .ColWidth(1) = 18
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
        
        .Col = 2
        .ColWidth(2) = 10
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 3
        .ColWidth(3) = 5
        .CellType = CellTypeCheckBox
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
    End With
    
    If P_09003_Flag = False Then
        dtInput(0).Value = Date
    
        ReDim sValue(0)
        
        sValue(0) = "0"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_09003_00", sValue(), Err_Num, Err_Dec)
        
        spdView.MaxCols = RS01.Fields.Count
        spdView.MaxRows = RS01.RecordCount
        
        Call spdDisplay(RS01)
        Call GetColWidth(REG_App, Me.Name, spdView)
        
        P_09003_Flag = True
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub spdDisplay(Rs As ADODB.Recordset)

    Call fpSpread_Display(spdView, Rs)
    

End Sub

Private Sub spdView_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    If Row = spdView.ActiveRow Then
        If Col = 3 Then
            spdView.Row = spdView.ActiveRow
            spdView.Col = Col
            If spdView.Value = False Then
                spdView.Col = 2
                spdView.Text = ""
            Else
                ReDim sValue(2)
                
                spdView.Row = Row
                
                sValue(0) = "0"
                sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
                spdView.Col = 1
                sValue(2) = Mid(spdView.Text, 2, 3)
                
                Set RS01 = New ADODB.Recordset
                Set RS01 = ExecPro("SP_09003_01", sValue(), Err_Num, Err_Dec)
                
                spdView.Col = 2
                If Not IsNull(RS01!문서번호) Then
                    spdView.Text = RS01!문서번호
                Else
                    spdView.Text = "1"
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_09003_Flag = False
End Sub

Public Sub DataSave()
    Dim i As Integer
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        spdView.Col = 3
        
        If spdView.Value = True Then
            ReDim sValue(5)
            
            sValue(0) = "2"
            sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
            spdView.Col = 1
            sValue(2) = Mid(spdView.Text, 2, 3)
            spdView.Col = 2
            sValue(3) = spdView.Value
            sValue(4) = rtbInput.Text
            sValue(5) = "1"
            
            Call ExecPro("SP_09003_02", sValue(), Err_Num, Err_Dec)
        End If
    Next i
End Sub

Public Sub DataAdd()
    rtbInput.Text = ""
End Sub

Private Sub Data_Display()
'
End Sub


