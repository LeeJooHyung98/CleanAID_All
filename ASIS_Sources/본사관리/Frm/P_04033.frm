VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_04033 
   Caption         =   "달성율 순위 조회"
   ClientHeight    =   10665
   ClientLeft      =   2595
   ClientTop       =   3195
   ClientWidth     =   14265
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_04033.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10665
   ScaleWidth      =   14265
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   10665
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14265
      _ExtentX        =   25162
      _ExtentY        =   18812
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_04033.frx":058A
      Begin Threed.SSPanel panInput 
         Height          =   675
         Left            =   15
         TabIndex        =   1
         Top             =   540
         Width           =   14235
         _ExtentX        =   25109
         _ExtentY        =   1191
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   90
            TabIndex        =   12
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "년도"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1275
            TabIndex        =   13
            Top             =   60
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy"
            Format          =   21299203
            CurrentDate     =   37140
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   900
         _Version        =   262144
         Font3D          =   1
         ForeColor       =   255
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
         Caption         =   " #"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_04033.frx":063C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Left            =   6660
         TabIndex        =   3
         Top             =   15
         Width           =   7590
         _ExtentX        =   13388
         _ExtentY        =   900
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
         PictureBackgroundStyle=   2
         PictureBackground=   "P_04033.frx":083E
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   4
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "종료"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Appearance      =   6
            Picture         =   "P_04033.frx":0A40
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   5
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "화면"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_04033.frx":0FDA
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   6
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "인쇄"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_04033.frx":1574
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   7
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "취소"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_04033.frx":1B0E
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   8
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "삭제"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_04033.frx":20A8
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   9
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "저장"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_04033.frx":2642
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   10
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "신규"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_04033.frx":2BDC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   11
            Top             =   30
            Width           =   900
            _Version        =   851970
            _ExtentX        =   1587
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "조회"
            ForeColor       =   -2147483640
            BackColor       =   -2147483636
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_04033.frx":3176
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   9420
         Index           =   0
         Left            =   15
         TabIndex        =   14
         Top             =   1230
         Width           =   6570
         _Version        =   524288
         _ExtentX        =   11589
         _ExtentY        =   16616
         _StockProps     =   64
         BackColorStyle  =   1
         EditEnterAction =   5
         EditModeReplace =   -1  'True
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
         MaxCols         =   6
         MaxRows         =   10
         SpreadDesigner  =   "P_04033.frx":3710
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   9420
         Index           =   1
         Left            =   6600
         TabIndex        =   15
         Top             =   1230
         Width           =   7650
         _Version        =   524288
         _ExtentX        =   13494
         _ExtentY        =   16616
         _StockProps     =   64
         BackColorStyle  =   1
         EditEnterAction =   5
         EditModeReplace =   -1  'True
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
         MaxCols         =   17
         MaxRows         =   10
         SpreadDesigner  =   "P_04033.frx":3E4D
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_04033"
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
    On Error GoTo ErrRtn

    Select Case Index
        Case 0: Call Data_Display   ' 조회
        Case 1: 'Call DataAdd        ' 신규
        Case 2: Call DataSave       ' 저장
        Case 3: 'Call DataDelete     ' 삭제
        Case 4: 'Call DataCancel     ' 취소
        Case 5: 'Call DataPrint      ' 인쇄
        Case 6: 'Call DataScreen     '
        Case 7: Unload Me           ' 종료
    End Select

'    Me.MousePointer = 0

    Exit Sub

ErrRtn:
    Me.MousePointer = 0

    If Err.Number = "0" Then

    ElseIf Err.Number = "91" Then
        End
    Else
        Resume Next
    End If
End Sub


Private Sub Form_Activate()
    cmdBtn(0).Enabled = True
    'cmdBtn(1).Enabled = True
    cmdBtn(2).Enabled = True
    'cmdBtn(3).Enabled = True
    'cmdBtn(4).Enabled = True

    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"


End Sub

'Private Sub spdDisplay(RS As ADODB.Recordset)
'    Call fpSpread_Display(spdView, RS)
'End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn

     With spdView(0)


        .MaxRows = 0
        .RowHeight(-1) = 14

        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle

        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeNormal

        'Init the User Sort
        .UserColAction = UserColActionSort

        .Row = -1

        .Col = 6: .ColHidden = True


    End With

    With spdView(1)


        .MaxRows = 0
        .RowHeight(-1) = 14

        'Spread 8 - 디자인
        .HighlightHeaders = HighlightHeadersOff
        .AppearanceStyle = AppearanceStyleEnhanced
        .ScrollBarStyle = ScrollBarStyleVisualStyle

        '선택된 Row
        .SelBackColor = &HFFFFC0 '황색 ^^
        .SelForeColor = &H0&     '검은글씨
        .OperationMode = OperationModeNormal

        .Col = 1: .ColMerge = MergeRestricted
        .Col = 2: .ColMerge = MergeRestricted
        .Col = 3: .ColMerge = MergeRestricted

        .ColsFrozen = 5 '틀고정
        .Row = -1

    End With

    If P_04033_Flag = False Then

        dtInput(0).Value = Date


        '''''''''''''Call GetColWidth(REG_App, Me.Name, spdView(0))
        '''''''''''''Call GetColWidth(REG_App, Me.Name, spdView(1))

        P_04033_Flag = True
    End If

    Exit Sub

ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_04033_Flag = False
End Sub

Private Sub Data_Display()
    On Error GoTo ErrRtn
    Dim nRow    As Long
    Dim vText   As Variant

    ReDim sValue(1)

    sValue(0) = Format(dtInput(0).Value, "yyyy")

    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_04033_00", sValue(), Err_Num, Err_Dec)

    spdView(0).MaxRows = RS01.RecordCount

    Call fpSpread_Display(spdView(0), RS01)
    Call GetColWidth(REG_App, Me.Name, spdView(0))

    With spdView(0)
        For nRow = 1 To .MaxRows
            .GetText 5, nRow, vText

            .Col = 5
            .Row = nRow
            .ForeColor = IIf(Val(vText) >= 1, vbRed, vbBlack)
        Next nRow
    End With

    Exit Sub

ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub


Public Sub DataSave()

End Sub

Public Sub DataDelete()

End Sub


Private Sub Data_DisplayUser(nRow As Long)
    On Error GoTo ErrRtn

    Dim nCol    As Long
    Dim vText(1)   As Variant

    ReDim sValue(1)

    spdView(0).GetText 6, nRow, vText(0):      sValue(0) = CStr(vText(0))
    sValue(1) = Format(dtInput(0).Value, "yyyy")

    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_04032_00", sValue(), Err_Num, Err_Dec)

    spdView(1).MaxRows = RS01.RecordCount

    Call fpSpread_Display(spdView(1), RS01, False)


    With spdView(1)
        .Redraw = False
        For nRow = 1 To .MaxRows



            If nRow Mod 3 Then
                '.Row = nRow:   .Col = 4: .Formula = "SUM(E" & CStr(nRow) & ":O" & CStr(nRow) & ")"
            Else

                For nCol = 5 To .MaxCols
                    .Row = nRow: .Col = nCol
                    .CellType = CellTypePercent
                    .TypePercentDecimal = "."
                    .TypeVAlign = TypeVAlignCenter
                    .TypeHAlign = TypeHAlignRight


                    .GetText nCol, nRow - 2, vText(0):  .GetText nCol, nRow - 1, vText(1)

                    If Val(vText(0)) > 0 Then
                        .SetText nCol, nRow, CVar(Val(vText(1)) / Val(vText(0)))
                    Else
                        .SetText nCol, nRow, CVar(Val(vText(1)) / 100)
                    End If

                    .Row = nRow: .Col = nCol
                    If Val(vText(0)) > 0 Then
                        .ForeColor = IIf(Val(vText(1)) / Val(vText(0)) >= 1, vbRed, vbBlack)

                    ElseIf Val(vText(1)) > 0 Then
                        .ForeColor = vbRed
                    Else
                        .ForeColor = vbBlack
                    End If

                    .BackColor = &HC0FFFF
                Next nCol

            End If
        Next nRow

        .Redraw = True
    End With

    Exit Sub

ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub spdView_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
    If Index = 0 Then
        Call Data_DisplayUser(Row)
    End If
End Sub
