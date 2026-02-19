VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RichTx32.ocx"
Begin VB.Form P_09002 
   Caption         =   "송신 메일관리"
   ClientHeight    =   12135
   ClientLeft      =   1620
   ClientTop       =   1920
   ClientWidth     =   16875
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_09002.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12135
   ScaleWidth      =   16875
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   12135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16875
      _ExtentX        =   29766
      _ExtentY        =   21405
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_09002.frx":058A
      Begin Threed.SSPanel panInput 
         Height          =   765
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   16845
         _ExtentX        =   29713
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
            Index           =   1
            Left            =   4710
            TabIndex        =   3
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   67043328
            CurrentDate     =   36686
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            TabIndex        =   4
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   67043328
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   5
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "송 신 일 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   9
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
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   11325
         Left            =   15
         TabIndex        =   7
         Top             =   795
         Width           =   6840
         _Version        =   524288
         _ExtentX        =   12065
         _ExtentY        =   19976
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
         SpreadDesigner  =   "P_09002.frx":05FC
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin RichTextLib.RichTextBox rtbInput 
         Height          =   11325
         Left            =   6870
         TabIndex        =   8
         Top             =   795
         Width           =   9990
         _ExtentX        =   17621
         _ExtentY        =   19976
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"P_09002.frx":0A64
      End
   End
End
Attribute VB_Name = "P_09002"
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
'    cmdBtn(2).Enabled = True
'    cmdBtn(5).Enabled = True
'    cmdBtn(6).Enabled = True
'
'    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    With spdView
        .ColsFrozen = 1 '틀고정
        .Row = -1
        
        .Col = 1
        .ColWidth(1) = 10
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 2
        .ColWidth(2) = 10
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 3
        .ColWidth(3) = 18
        .CellType = CellTypeEdit
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignLeft
        
        .Col = 4
        .ColWidth(4) = 5
        .CellType = CellTypeCheckBox
        .Value = False
        .TypeVAlign = TypeVAlignCenter
        .TypeHAlign = TypeHAlignCenter
    End With
    
    If P_09002_Flag = False Then
        dtInput(0).Value = Date
        dtInput(1).Value = Date
    
        Call AgencyComboAdd(cboInput)
        
        ReDim sValue(3)
        
        sValue(0) = "1"
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_09002_00", sValue(), Err_Num, Err_Dec)
        
        spdView.MaxCols = RS01.Fields.Count
        spdView.MaxRows = RS01.RecordCount
        
        Call spdDisplay(RS01)
        Call GetColWidth(REG_App, Me.Name, spdView)
        
        P_09002_Flag = True
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub spdDisplay(Rs As ADODB.Recordset)
        
    Call fpSpread_Display(spdView, Rs)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_09002_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    ReDim sValue(3)
    
    sValue(0) = "0"
    sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
    sValue(2) = Format(dtInput(1).Value, "YYYY-MM-DD")
    sValue(3) = Mid(cboInput.Text, 2, 6) & "%"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_09002_00", sValue(), Err_Num, Err_Dec)
    
    spdView.MaxCols = RS01.Fields.Count
    spdView.MaxRows = RS01.RecordCount
    
    Call spdDisplay(RS01)
    Call GetColWidth(REG_App, Me.Name, spdView)
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Public Sub DataSave()
'+------------------------------------------------------
'+
'+ 2003/04/11
'+
'+루틴설명
'+  1. 목록에 선택된 내용을 DB에 적용 시킨다.
'+  2. Mail의 SendChk = "2"로 변경하여 모뎀자료 생성에서
'+      생성 되도록 설정한다.
'+  3. 체크된것만을 작업하기 때문에 이전 작업 내용과 무관하다.
'+------------------------------------------------------
ReDim sValue(4)
Dim i As Long
Dim strSql As String
Dim strCnn As String
Dim rstMail As ADODB.Recordset
    
    For i = 1 To spdView.MaxRows
        spdView.Row = i
        spdView.Col = 4
        
        If spdView.Value = True Then
            spdView.Col = 1: sValue(0) = Format(spdView.Value, "YYYY-MM-DD")
            spdView.Col = 2: sValue(1) = Val(spdView.Text)
            spdView.Col = 3: sValue(2) = Mid(spdView.Text, 2, 3)
                             sValue(3) = "2"
            
            strSql = "SELECT * FROM Mail (nolock) WHERE MailDate = '" _
                & sValue(0) & "' AND AgencyCode = '" & sValue(2) & "' AND MailType = '" _
                & sValue(3) & "' AND MailNo = '" & sValue(1) & "'"
            Set rstMail = New ADODB.Recordset
            
            rstMail.CursorType = adOpenKeyset
            rstMail.LockType = adLockOptimistic
            rstMail.Open strSql, m_DBConnect, , , adCmdText
            rstMail!SendChk = "2"
            rstMail.Update
            rstMail.Close
        End If
    Next i
    
End Sub

Public Sub DataPrint()

End Sub

Public Sub DataScreen()

End Sub

Private Sub PrintDesc()

End Sub

Private Sub spdView_DblClick(ByVal Col As Long, ByVal Row As Long)
    ReDim sValue(3)
    
    sValue(0) = "0"
    
    spdView.Row = spdView.ActiveRow
    spdView.Col = 1: sValue(1) = Format(spdView.Value, "YYYY-MM-DD")
    spdView.Col = 2: sValue(2) = spdView.Value
    spdView.Col = 3: sValue(3) = Mid(spdView.Text, 2, 3)
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_09002_01", sValue(), Err_Num, Err_Dec)
    
    rtbInput.Text = RS01!메일내역
End Sub


