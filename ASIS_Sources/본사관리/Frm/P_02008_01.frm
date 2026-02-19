VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form P_02008_01 
   Caption         =   "입고검품 현황"
   ClientHeight    =   10725
   ClientLeft      =   2985
   ClientTop       =   4005
   ClientWidth     =   15165
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_02008_01.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10725
   ScaleWidth      =   15165
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   10725
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   18918
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_02008_01.frx":058A
      Begin Threed.SSPanel SSPanel1 
         Height          =   405
         Left            =   15
         TabIndex        =   12
         Top             =   10305
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   714
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtinput 
            Alignment       =   2  '가운데 맞춤
            Height          =   315
            Index           =   0
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   45
            Width           =   1290
         End
         Begin VB.TextBox txtinput 
            Alignment       =   2  '가운데 맞춤
            Height          =   315
            Index           =   1
            Left            =   3450
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   45
            Width           =   1110
         End
         Begin VB.TextBox txtinput 
            Height          =   315
            Index           =   2
            Left            =   5700
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   45
            Width           =   1110
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   345
            Index           =   1
            Left            =   10725
            TabIndex        =   15
            Top             =   45
            Width           =   1215
            _Version        =   262145
            _ExtentX        =   2143
            _ExtentY        =   609
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   3
            StartText.y     =   5
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   13
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txtNum 
            Height          =   345
            Index           =   0
            Left            =   8415
            TabIndex        =   16
            Top             =   45
            Width           =   1215
            _Version        =   262145
            _ExtentX        =   2143
            _ExtentY        =   609
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   " 0"
            StartText.x     =   3
            StartText.y     =   5
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   13
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            Undo            =   0
            Data            =   0
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "누 락 계:"
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
            Index           =   4
            Left            =   9855
            TabIndex        =   22
            Top             =   120
            Width           =   810
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "검 품 계:"
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
            Index           =   3
            Left            =   7560
            TabIndex        =   21
            Top             =   120
            Width           =   810
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "종 료 택:"
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
            Index           =   2
            Left            =   4815
            TabIndex        =   20
            Top             =   120
            Width           =   810
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "시 작 택:"
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
            Index           =   0
            Left            =   2580
            TabIndex        =   18
            Top             =   120
            Width           =   810
         End
         Begin VB.Label Label 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "전일종료:"
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
            Index           =   1
            Left            =   105
            TabIndex        =   17
            Top             =   120
            Width           =   810
         End
      End
      Begin Threed.SSPanel panInput 
         Height          =   765
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   1349
         _Version        =   262144
         BackColor       =   16777215
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
            Format          =   63569920
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
            Format          =   63569920
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   6
            Left            =   60
            TabIndex        =   6
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
         Begin VB.Label Label1 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "~"
            Height          =   255
            Left            =   4515
            TabIndex        =   7
            Top             =   120
            Width           =   315
         End
      End
      Begin Threed.SSPanel panCaption 
         Height          =   390
         Index           =   7
         Left            =   15
         TabIndex        =   8
         Top             =   795
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   688
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
         Caption         =   "검 품 택"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_02008_01.frx":065C
         BevelOuter      =   0
         RoundedCorners  =   0   'False
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   4080
         Index           =   1
         Left            =   15
         TabIndex        =   9
         Top             =   1200
         Width           =   15135
         _Version        =   524288
         _ExtentX        =   26696
         _ExtentY        =   7197
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
         MaxCols         =   15
         ScrollBars      =   2
         SpreadDesigner  =   "P_02008_01.frx":0ABE
         UserResize      =   0
      End
      Begin Threed.SSPanel panCaption 
         Height          =   390
         Index           =   8
         Left            =   15
         TabIndex        =   10
         Top             =   5295
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   688
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
         Caption         =   "누 락 택"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_02008_01.frx":11AE
         BevelOuter      =   0
         RoundedCorners  =   0   'False
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   4590
         Index           =   2
         Left            =   15
         TabIndex        =   11
         Top             =   5700
         Width           =   15135
         _Version        =   524288
         _ExtentX        =   26696
         _ExtentY        =   8096
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
         MaxCols         =   15
         ScrollBars      =   2
         SpreadDesigner  =   "P_02008_01.frx":1610
         UserResize      =   0
         AppearanceStyle =   0
      End
   End
End
Attribute VB_Name = "P_02008_01"
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
    
'    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
    If P_02008_01_Flag = False Then
        dtInput(0).Value = P_02008.dtInput(0).Value
        dtInput(1).Value = P_02008.dtInput(1).Value
        
        DoEvents
        
        Call Data_Display
        
        P_02008_01_Flag = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_02008_01_Flag = True
End Sub

Public Sub DataPrint()

End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    
    Dim z As Integer

    Dim iTotalQty(1) As Integer

    Dim memTag As String
    
    spdView(1).MaxRows = 0
    spdView(1).MaxRows = 100
    
    spdView(2).MaxRows = 0
    spdView(2).MaxRows = 100
    
    z = P_02008.ActiveControl.Index
    
    P_02008.spdView(z).Row = P_02008.spdView(z).ActiveRow
    P_02008.spdView(z).Col = 7
    
    i = 1
    j = 0
    k = 1
    l = 0
    
    If P_02008.spdView(z).Text > "2" Then
        ReDim sValue(1)
        
        sValue(0) = UserID
        sValue(1) = Mid(cboInput.Text, 2, 6)
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_02008_04", sValue(), Err_Num, Err_Dec)
        
        Do While Not RS01.EOF
            '----------------------------------------------------
            ' 구분코드가 3이하 이면 검품택, 그렇치 않으면 누락택
            '----------------------------------------------------
            If RS01!구분 < 3 Then
                j = j + 1
                
                If j > 15 Then
                    i = i + 1
                    j = 1
                End If
                
                spdView(1).Row = i
                spdView(1).Col = j: spdView(1).Text = RS01!택번호 & ""
                
                iTotalQty(0) = iTotalQty(0) + 1
            Else
                l = l + 1
                
                If l > 15 Then
                    k = k + 1
                    l = 1
                End If
                
                spdView(2).Row = k
                spdView(2).Col = l: spdView(2).Text = RS01!택번호
            
                iTotalQty(1) = iTotalQty(1) + 1
            End If
            
            RS01.MoveNext
        Loop
    Else
        ReDim sValue(2)
        
        sValue(0) = Format(dtInput(0).Value, "YYYY-MM-DD")
        sValue(1) = Format(dtInput(1).Value, "YYYY-MM-DD")
        sValue(2) = Mid(cboInput.Text, 2, 6)
        
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_02008_05", sValue(), Err_Num, Err_Dec)
        
        If Not RS01.EOF Then
            memTag = RS01!택번호
        End If
        
        Do While Not RS01.EOF
            j = j + 1
            
            If j > 15 Then
                i = i + 1
                j = 1
            End If
            
            spdView(1).Row = i
            spdView(1).Col = j: spdView(1).Text = RS01!택번호 & ""
            
            iTotalQty(0) = iTotalQty(0) + 1
            
            Do While memTag < RS01!택번호
                l = l + 1
                
                If l > 15 Then
                    k = k + 1
                    l = 1
                End If
                
                spdView(2).Row = k
                spdView(2).Col = l: spdView(2).Text = RS01!택번호 & ""
            
                iTotalQty(1) = iTotalQty(1) + 1
                
                memTag = Right("0000" & Val(Mid(memTag, 1, 1) & Mid(memTag, 3, 3)) + 1, 4)
                memTag = Mid(memTag, 1, 1) & "-" & Mid(memTag, 2, 3)
            Loop
        
            memTag = Right("0000" & Val(Mid(RS01!택번호, 1, 1) & Mid(RS01!택번호, 3, 3)) + 1, 4)
            memTag = Mid(memTag, 1, 1) & "-" & Mid(memTag, 2, 3)
            
            RS01.MoveNext
        Loop
    End If
    
    P_02008.spdView(z).Col = 3: txtInput(0).Text = P_02008.spdView(z).Text
    P_02008.spdView(z).Col = 4: txtInput(1).Text = P_02008.spdView(z).Text
    P_02008.spdView(z).Col = 5: txtInput(2).Text = P_02008.spdView(z).Text
    
    txtNum(0).Value = iTotalQty(0)
    txtNum(1).Value = iTotalQty(1)
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub
