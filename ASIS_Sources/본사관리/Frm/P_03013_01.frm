VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form P_03013_01 
   Caption         =   "출고TAG번호 CHECK"
   ClientHeight    =   11205
   ClientLeft      =   1275
   ClientTop       =   1920
   ClientWidth     =   16455
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_03013_01.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11205
   ScaleWidth      =   16455
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   11205
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   19764
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_03013_01.frx":058A
      Begin Threed.SSPanel SSPanel 
         Height          =   405
         Left            =   15
         TabIndex        =   14
         Top             =   10785
         Width           =   16425
         _ExtentX        =   28972
         _ExtentY        =   714
         _Version        =   262144
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   2
            Left            =   7875
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   45
            Width           =   1455
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   1
            Left            =   4695
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   45
            Width           =   1455
         End
         Begin VB.TextBox txtInput 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   315
            Index           =   0
            Left            =   1515
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   45
            Width           =   1455
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   45
            TabIndex        =   18
            Top             =   45
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "출 고 수 량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   3225
            TabIndex        =   19
            Top             =   45
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "중 복 수 량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   3
            Left            =   6405
            TabIndex        =   20
            Top             =   45
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "누 락 수 량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel panInput 
         Height          =   765
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   16425
         _ExtentX        =   28972
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
            Format          =   62980096
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
            Left            =   4770
            TabIndex        =   5
            Top             =   60
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   62980096
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   7
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
            Left            =   4530
            TabIndex        =   7
            Top             =   120
            Width           =   255
         End
      End
      Begin Threed.SSPanel panCaption 
         Height          =   345
         Index           =   4
         Left            =   15
         TabIndex        =   8
         Top             =   795
         Width           =   16425
         _ExtentX        =   28972
         _ExtentY        =   609
         _Version        =   262144
         Font3D          =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "출  고  택"
         BevelOuter      =   0
         RoundedCorners  =   0   'False
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   3465
         Index           =   1
         Left            =   15
         TabIndex        =   9
         Top             =   1155
         Width           =   16425
         _Version        =   524288
         _ExtentX        =   28972
         _ExtentY        =   6112
         _StockProps     =   64
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   15
         ScrollBars      =   2
         SpreadDesigner  =   "P_03013_01.frx":069C
         AppearanceStyle =   0
      End
      Begin Threed.SSPanel panCaption 
         Height          =   345
         Index           =   5
         Left            =   15
         TabIndex        =   10
         Top             =   4635
         Width           =   16425
         _ExtentX        =   28972
         _ExtentY        =   609
         _Version        =   262144
         Font3D          =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "중  복  택"
         BevelOuter      =   0
         RoundedCorners  =   0   'False
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   2280
         Index           =   2
         Left            =   15
         TabIndex        =   11
         Top             =   4995
         Width           =   16425
         _Version        =   524288
         _ExtentX        =   28972
         _ExtentY        =   4022
         _StockProps     =   64
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   15
         ScrollBars      =   2
         SpreadDesigner  =   "P_03013_01.frx":0BCA
         AppearanceStyle =   0
      End
      Begin Threed.SSPanel panCaption 
         Height          =   345
         Index           =   6
         Left            =   15
         TabIndex        =   12
         Top             =   7290
         Width           =   16425
         _ExtentX        =   28972
         _ExtentY        =   609
         _Version        =   262144
         Font3D          =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "누  락  택"
         BevelOuter      =   0
         RoundedCorners  =   0   'False
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   3120
         Index           =   3
         Left            =   15
         TabIndex        =   13
         Top             =   7650
         Width           =   16425
         _Version        =   524288
         _ExtentX        =   28972
         _ExtentY        =   5503
         _StockProps     =   64
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   15
         ScrollBars      =   2
         SpreadDesigner  =   "P_03013_01.frx":10E7
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_03013_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String
Dim ii As Integer

Private Sub cboInput_Click()
    Call Data_Display
End Sub

Private Sub dtInput_Change(Index As Integer)
'    Call Data_Display
End Sub

Private Sub Form_Activate()
'    cmdBtn(0).Enabled = True
'    cmdBtn(5).Enabled = True
'    cmdBtn(6).Enabled = True
'
'    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
    
    Dim i As Integer
    
    dtInput(0).Value = P_03013.dtInput(0).Value
    dtInput(1).Value = P_03013.dtInput(1).Value
    
    ii = P_03013.ActiveControl.Index
    
    If P_03013_01_Flag = False Then
        Call Data_Display
        
        P_03013_01_Flag = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_03013_01_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim iCnt As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    Dim m As Integer
    Dim n As Integer
    
    Dim t As Integer
    
    Dim rCnt As Long
    Dim sTag As String
    
    For i = 1 To 3
        For j = 1 To spdView(i).MaxRows
            For k = 1 To spdView(i).MaxCols
                spdView(i).Row = j
                spdView(i).Col = k
                spdView(i).Text = ""
            Next k
        Next j
    Next i
    
    P_03013.spdView(ii).Row = P_03013.spdView(ii).ActiveRow
    P_03013.spdView(ii).Col = 7
    
    i = 1
    j = 0
    k = 1
    l = 0
    m = 1
    n = 0
    
    If P_03013.spdView(ii).Text > "2" Then
        ReDim sValue(2)
        
        sValue(0) = "0"
        sValue(1) = UserID
        P_03013.spdView(ii).Row = P_03013.spdView(ii).ActiveRow
        P_03013.spdView(ii).Col = 1
        sValue(2) = Mid(P_03013.spdView(ii).Text, 2, 3)
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_03013_04", sValue(), Err_Num, Err_Dec)
        
        Do While Not RS01.EOF
            If RS01!구분 < 3 Then
                For t = 1 To RS01!수량
                    j = j + 1
                    If j > 9 Then
                        i = i + 1
                        j = 1
                    End If
                
                    spdView(1).Row = i
                    spdView(1).Col = j
                    
                    spdView(1).Text = RS01!택번호
                Next t
                
                t = 1
                
                Do While RS01!수량 > t
                    n = n + 1
                    If n > 16 Then
                        m = m + 1
                        n = 1
                    End If
                    
                    spdView(2).Row = m
                    spdView(2).Col = n
                    spdView(2).Text = RS01!택번호
                    t = t + 1
                Loop
            Else
                i = i + 1
                If i > 14 Then
                    k = k + 1
                    l = 1
                End If
                
                spdView(3).Row = k
                spdView(3).Col = l
                spdView(3).Text = RS01!택번호
            End If
        
            RS01.MoveNext
        Loop
    Else
        ReDim sValue(3)
        
        sValue(0) = "0"
        sValue(1) = Format(dtInput(0).Value, "YYYY-MM-DD")
        sValue(2) = Format(dtInput(0).Value, "YYYY-MM-DD")
        
        P_03013.spdView(ii).Row = P_03013.spdView(ii).ActiveRow
        P_03013.spdView(ii).Col = 1
        sValue(3) = Mid(P_03013.spdView(ii).Text, 2, 3)
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_03013_05", sValue(), Err_Num, Err_Dec)
        
        If Not RS01.EOF Then
            rCnt = Val(Mid(RS01!택번호, 1, 1) & Mid(RS01!택번호, 3, 3))
        End If
        
        Do While Not RS01.EOF
            For t = 1 To RS01!수량
                j = j + 1
                If j > 16 Then
                    i = i + 1
                    j = 1
                End If
                
                spdView(1).Row = i
                spdView(1).Col = j: spdView(1).Text = RS01!택번호 & ""
            Next t
            
            t = 1
            
            Do While RS01!수량 > t
                n = n + 1
                If n > 16 Then
                    m = m + 1
                    n = 1
                End If
                
                spdView(2).Row = m
                spdView(2).Col = n
                spdView(2).Text = RS01!택번호
            Loop
            
            Do While rCnt < Val(Mid(RS01!택번호, 1, 1) & Mid(RS01!택번호, 3, 3))
                sTag = Trim(Str(rCnt))
                sTag = Right("0000" & sTag, 4)
                
                l = l + 1
                If l > 16 Then
                    k = k + 1
                    l = 1
                End If
                
                spdView(3).Row = k
                spdView(3).Col = l: spdView(3).Text = Mid(sTag, 1, 1) & "-" & Mid(sTag, 2, 3)
                
                rCnt = rCnt + 1
            Loop
            
            rCnt = Val(Mid(RS01!택번호, 1, 1) & Mid(RS01!택번호, 3, 3)) + 1
            
            RS01.MoveNext
        Loop
    End If
    
    P_03013.spdView(ii).Col = 2: txtInput(0).Text = P_03013.spdView(ii).Text
    P_03013.spdView(ii).Col = 5: txtInput(1).Text = P_03013.spdView(ii).Text
    P_03013.spdView(ii).Col = 6: txtInput(2).Text = P_03013.spdView(ii).Text
        
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Public Sub DataPrint()

End Sub

Private Sub spdView_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        'PopupMenu P_00000.PopUp
    End If
End Sub

Public Sub DataScreen()
'    Dim ReportFP As String
'    Dim ReportFile As String
'
'    ReportFP = GetIniStr("REPORT", "FilePath", "", m_iniFile)
'    ReportFile = ReportFP & "\" & Me.Name & ".rpt"
'
'    Call PrintDesc
'
'    Dim i As Integer
'    For i = 0 To 30
'        P_00000.crPrint.Formulas(i) = ""
'    Next
'
'    P_00000.crPrint.WindowTitle = Me.Caption
'    P_00000.crPrint.Formulas(0) = "출고일자1 = '" & Format(dtInput(0).Value, "YYYY-MM-DD") & "'"
'    P_00000.crPrint.Formulas(1) = "출고일자2 = '" & Format(dtInput(0).Value, "YYYY-MM-DD") & "'"
'
'    Call ReportPrint(ReportFile, "2")
End Sub

Private Sub PrintDesc()
    Dim i As Integer
    Dim TempText As String
    Dim TempFP As String
    Dim TempFile As String
    
    TempFP = GetIniStr("REPORT", "TempPath", "", m_iniFile)
    TempFile = TempFP & "\Temp.txt"
    
    Open TempFile For Output As #1
    
    TempText = ""
    
    For i = 1 To spdView(0).MaxRows - 1
        spdView(0).Row = i
        
        spdView(0).Col = 6
        If spdView(0).Text = 0 Then
            spdView(0).Col = 1
            TempText = TempText & "    " & LeftH(spdView(0).Text & Space(20), 20)
        Else
            spdView(0).Col = 1
            TempText = TempText & "   *" & LeftH(spdView(0).Text & Space(20), 20)
        End If
        
        spdView(0).Col = 2
        TempText = TempText & Right(Space(8) & spdView(0).Text, 8)
        spdView(0).Col = 5
        TempText = TempText & Right(Space(8) & spdView(0).Text, 8)
        spdView(0).Col = 6
        TempText = TempText & Right(Space(8) & spdView(0).Text, 8)
        
        If i Mod 2 = 0 Then
            Print #1, TempText
            TempText = ""
        End If
    Next i
    
    Close #1
End Sub
