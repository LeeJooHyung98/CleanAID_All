VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm출고작업 
   Caption         =   "출고작업"
   ClientHeight    =   10080
   ClientLeft      =   8775
   ClientTop       =   2355
   ClientWidth     =   15240
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10080
   ScaleWidth      =   15240
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   10080
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   17780
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm출고작업.frx":0000
      Begin Threed.SSPanel SSPanel3 
         Height          =   360
         Left            =   15
         TabIndex        =   21
         Top             =   5955
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   635
         _Version        =   262144
         BackColor       =   16777215
         PictureBackgroundStyle=   2
         PictureBackground=   "frm출고작업.frx":00F2
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.ProgressBar ProgressBar1 
            Height          =   300
            Left            =   12195
            TabIndex        =   22
            Top             =   30
            Width           =   2985
            _Version        =   851970
            _ExtentX        =   5265
            _ExtentY        =   529
            _StockProps     =   93
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label lblCount 
            BackStyle       =   0  '투명
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   1
            Left            =   1635
            TabIndex        =   24
            Top             =   90
            Width           =   975
         End
         Begin VB.Image Image2 
            Height          =   240
            Index           =   1
            Left            =   105
            MouseIcon       =   "frm출고작업.frx":0314
            MousePointer    =   99  '사용자 정의
            Picture         =   "frm출고작업.frx":0466
            Top             =   75
            Width           =   240
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "출고예정량 :"
            Height          =   180
            Index           =   1
            Left            =   465
            TabIndex        =   23
            Top             =   105
            Width           =   1080
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   360
         Left            =   15
         TabIndex        =   18
         Top             =   1215
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   635
         _Version        =   262144
         BackColor       =   16777215
         PictureBackgroundStyle=   2
         PictureBackground=   "frm출고작업.frx":0E68
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "출고대기량 :"
            Height          =   180
            Index           =   0
            Left            =   450
            TabIndex        =   20
            Top             =   105
            Width           =   1080
         End
         Begin VB.Image Image2 
            Height          =   240
            Index           =   0
            Left            =   105
            MouseIcon       =   "frm출고작업.frx":108A
            MousePointer    =   99  '사용자 정의
            Picture         =   "frm출고작업.frx":11DC
            Top             =   75
            Width           =   240
         End
         Begin VB.Label lblCount 
            BackStyle       =   0  '투명
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   0
            Left            =   1620
            TabIndex        =   19
            Top             =   90
            Width           =   975
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   750
         Index           =   0
         Left            =   15
         TabIndex        =   1
         Top             =   5190
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   1323
         _Version        =   262144
         BackColor       =   14280169
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   5
            Left            =   45
            TabIndex        =   2
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 전체출고"
            Appearance      =   6
            Picture         =   "frm출고작업.frx":1BDE
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   6
            Left            =   1575
            TabIndex        =   3
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 전체취소"
            Appearance      =   6
            Picture         =   "frm출고작업.frx":22D8
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   0
            Left            =   3165
            TabIndex        =   4
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 개별출고"
            Appearance      =   6
            Picture         =   "frm출고작업.frx":29D2
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   1
            Left            =   4695
            TabIndex        =   5
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 개별취소"
            Appearance      =   6
            Picture         =   "frm출고작업.frx":30CC
         End
         Begin XtremeSuiteControls.PushButton cmdOut 
            Height          =   630
            Left            =   8055
            TabIndex        =   6
            Top             =   60
            Width           =   1500
            _Version        =   851970
            _ExtentX        =   2646
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 출고완료"
            Appearance      =   6
            Picture         =   "frm출고작업.frx":37C6
         End
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   3585
         Index           =   0
         Left            =   15
         TabIndex        =   7
         Top             =   1590
         Width           =   15210
         _Version        =   524288
         _ExtentX        =   26829
         _ExtentY        =   6324
         _StockProps     =   64
         AutoCalc        =   0   'False
         BackColorStyle  =   1
         ColsFrozen      =   7
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         MaxCols         =   10
         MaxRows         =   200
         Protect         =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frm출고작업.frx":3EC0
         UserResize      =   1
         VisibleCols     =   7
         VisibleRows     =   30
         HighlightHeaders=   1
         HighlightStyle  =   1
         ScrollBarStyle  =   2
      End
      Begin Threed.SSPanel Panel 
         Height          =   750
         Left            =   15
         TabIndex        =   8
         Top             =   450
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   1323
         _Version        =   262144
         BackColor       =   14280169
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin CSTextLibCtl.sitxEdit txtTag 
            Height          =   630
            Index           =   0
            Left            =   750
            TabIndex        =   9
            Top             =   60
            Width           =   1860
            _Version        =   262145
            _ExtentX        =   3281
            _ExtentY        =   1111
            _StockProps     =   125
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   2
            StartText.y     =   6
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   29
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   1
            CharacterTable  =   ""
         End
         Begin CSTextLibCtl.sitxEdit txtTag 
            Height          =   630
            Index           =   1
            Left            =   5175
            TabIndex        =   10
            Top             =   60
            Visible         =   0   'False
            Width           =   1860
            _Version        =   262145
            _ExtentX        =   3281
            _ExtentY        =   1111
            _StockProps     =   125
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   2
            StartText.y     =   6
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   29
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   1
            CharacterTable  =   ""
         End
         Begin CSTextLibCtl.sitxEdit txtTag 
            Height          =   630
            Index           =   2
            Left            =   8490
            TabIndex        =   11
            Top             =   60
            Visible         =   0   'False
            Width           =   1860
            _Version        =   262145
            _ExtentX        =   3281
            _ExtentY        =   1111
            _StockProps     =   125
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   2
            StartText.y     =   6
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   29
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   1
            CharacterTable  =   ""
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   4
            Left            =   12330
            TabIndex        =   12
            Top             =   60
            Width           =   1395
            _Version        =   851970
            _ExtentX        =   2461
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 조회(&F)"
            Appearance      =   6
            Picture         =   "frm출고작업.frx":4ED9
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   630
            Index           =   2
            Left            =   13770
            TabIndex        =   13
            Top             =   60
            Width           =   1395
            _Version        =   851970
            _ExtentX        =   2461
            _ExtentY        =   1111
            _StockProps     =   79
            Caption         =   " 닫기(&X)"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "frm출고작업.frx":55D3
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "이전 택번호:"
            Height          =   180
            Index           =   1
            Left            =   7335
            TabIndex        =   17
            Top             =   105
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "현재 택번호:"
            Height          =   180
            Index           =   0
            Left            =   4020
            TabIndex        =   16
            Top             =   105
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "택번호:"
            Height          =   180
            Index           =   2
            Left            =   60
            TabIndex        =   15
            Top             =   105
            Width           =   630
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   420
         Left            =   15
         TabIndex        =   14
         Top             =   15
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   741
         _Version        =   262144
         Font3D          =   3
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "      가맹점 출고작업"
         PictureBackgroundStyle=   2
         PictureBackground=   "frm출고작업.frx":6665
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frm출고작업.frx":688B
            Top             =   -15
            Width           =   765
         End
      End
      Begin FPSpreadADO.fpSpread sprGrid 
         Height          =   3735
         Index           =   1
         Left            =   15
         TabIndex        =   25
         Top             =   6330
         Width           =   15210
         _Version        =   524288
         _ExtentX        =   26829
         _ExtentY        =   6588
         _StockProps     =   64
         AutoCalc        =   0   'False
         BackColorStyle  =   1
         ColsFrozen      =   7
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GrayAreaBackColor=   16777215
         GridSolid       =   0   'False
         MaxCols         =   10
         MaxRows         =   200
         Protect         =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frm출고작업.frx":7455
         UserResize      =   1
         VisibleCols     =   7
         VisibleRows     =   30
         HighlightHeaders=   1
         HighlightStyle  =   1
         ScrollBarStyle  =   2
      End
   End
End
Attribute VB_Name = "frm출고작업"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tData(0 To 8) As String

Private Sub 출고_Display()
    On Error GoTo ErrRtn
    
    lblCount(0).Caption = "0"
    
    Query = "SELECT    접수일자"
    Query = Query & ", 의류명"
    Query = Query & ", 택번호"
    Query = Query & ", 색상"
    Query = Query & ", 무늬"
    Query = Query & ", 내용"
    Query = Query & ", 금액"
    Query = Query & ", 상표"
    Query = Query & ", 의류코드"
    Query = Query & " FROM TB_입출고"
    Query = Query & " WHERE (가맹점출고일자 IS NULL OR 가맹점출고일자 = '')"
    Query = Query & "   AND ((판매취소 <> 'Y')"
    'Query = Query & "   AND (판매취소 = '') AND (판매취소일자 IS NULL OR 판매취소일자 = '')"
    Query = Query & "   AND (반품환불일자 IS NULL OR 반품환불일자 = '')"
    Query = Query & "   AND (세탁환불일자 IS NULL OR 세탁환불일자 = ''))"
    Query = Query & " ORDER BY 택번호 ASC"
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly

    With sprGrid(0)
        .MaxRows = 0
        .ReDraw = False
        
        Do Until ADORs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1:  .Text = "0"
            .Col = 2:  .Text = ADORs!접수일자 & ""
            .Col = 3:  .Text = ADORs!의류명 & ""
            
            If Len(Trim(ADORs!택번호)) <= 6 Then
                .Col = 4: .Text = ADORs!택번호 & ""
            Else
                .Col = 4: .Text = Format(ADORs!택번호, "000-00-0000")
            End If
            
            .Col = 5:  .Text = ADORs!색상 & ""
            .Col = 6:  .Text = ADORs!무늬 & ""
            .Col = 7:  .Text = ADORs!내용 & ""
            .Col = 8:  .Text = ADORs!금액 & ""
            .Col = 9:  .Text = ADORs!상표 & ""
            .Col = 10: .Text = ADORs!의류코드 & ""
            
            ADORs.MoveNext
        Loop
        
        ADORs.Close
        Set ADORs = Nothing
        
        lblCount(0).Caption = .MaxRows
        
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    Dim iSource As Integer
    Dim iTarget As Integer
    
    On Error GoTo ErrRtn

    Select Case Index
        Case 0: Call Data_Move(0, 1)
        Case 1: Call Data_Move(1, 0)
        
        Case 2: Unload Me
        Case 3
        Case 4
            Call 출고_Display
            
            txtTag(0).SetFocus
            
        Case 5: Call DataTotal_Move(0, 1)
        Case 6: Call DataTotal_Move(1, 0)
    End Select
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub Data_Move(iSource As Integer, iTarget As Integer)
    On Error GoTo ErrRtn
    
    With sprGrid(iSource)
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            If .Text = "1" Then
                .Col = 2:  tData(0) = .Text
                .Col = 3:  tData(1) = .Text
                .Col = 4:  tData(2) = .Text
                .Col = 5:  tData(3) = .Text
                .Col = 6:  tData(4) = .Text
                .Col = 7:  tData(5) = .Text
                .Col = 8:  tData(6) = .Text
                .Col = 9:  tData(7) = .Text
                .Col = 10: tData(8) = .Text
                
                .DeleteRows i, 1        '
                .MaxRows = .MaxRows - 1 '
                i = i - 1               ' 현재 Row를 삭제하면서 현재  Row가 없어지므로 현재위치를 1 Row 빼준다.
                                            
                '
                sprGrid(iTarget).MaxRows = sprGrid(iTarget).MaxRows + 1
                sprGrid(iTarget).Row = sprGrid(iTarget).MaxRows
                
                sprGrid(iTarget).Col = 1:  sprGrid(iTarget).Text = "0"
                sprGrid(iTarget).Col = 2:  sprGrid(iTarget).Text = tData(0)
                sprGrid(iTarget).Col = 3:  sprGrid(iTarget).Text = tData(iTarget)
                sprGrid(iTarget).Col = 4:  sprGrid(iTarget).Text = tData(2)
                sprGrid(iTarget).Col = 5:  sprGrid(iTarget).Text = tData(3)
                sprGrid(iTarget).Col = 6:  sprGrid(iTarget).Text = tData(4)
                sprGrid(iTarget).Col = 7:  sprGrid(iTarget).Text = tData(5)
                sprGrid(iTarget).Col = 8:  sprGrid(iTarget).Text = tData(6)
                sprGrid(iTarget).Col = 9:  sprGrid(iTarget).Text = tData(7)
                sprGrid(iTarget).Col = 10: sprGrid(iTarget).Text = tData(8)
            End If
        Next i
    End With
        
    sprGrid(iTarget).SortKey(1) = 2
    sprGrid(iTarget).SortKeyOrder(1) = SortKeyOrderAscending
    sprGrid(iTarget).Sort -1, -1, -1, -1, SortByRow
        
    lblCount(iSource).Caption = sprGrid(iSource).MaxRows
    lblCount(iTarget).Caption = sprGrid(iTarget).MaxRows
    
    txtTag(0).SetFocus
    
    Exit Sub
    
ErrRtn:

End Sub

Private Sub DataTotal_Move(iSource As Integer, iTarget As Integer)
    On Error GoTo ErrRtn
    
    With sprGrid(iSource)
        For i = 1 To .MaxRows
            .Row = i
            
            .Col = 2:  tData(0) = .Text
            .Col = 3:  tData(1) = .Text
            .Col = 4:  tData(2) = .Text
            .Col = 5:  tData(3) = .Text
            .Col = 6:  tData(4) = .Text
            .Col = 7:  tData(5) = .Text
            .Col = 8:  tData(6) = .Text
            .Col = 9:  tData(7) = .Text
            .Col = 10: tData(8) = .Text
            
            sprGrid(iTarget).MaxRows = sprGrid(iTarget).MaxRows + 1
            sprGrid(iTarget).Row = sprGrid(iTarget).MaxRows
            
            sprGrid(iTarget).Col = 1:  sprGrid(iTarget).Text = "0"
            sprGrid(iTarget).Col = 2:  sprGrid(iTarget).Text = tData(0)
            sprGrid(iTarget).Col = 3:  sprGrid(iTarget).Text = tData(1)
            sprGrid(iTarget).Col = 4:  sprGrid(iTarget).Text = tData(2)
            sprGrid(iTarget).Col = 5:  sprGrid(iTarget).Text = tData(3)
            sprGrid(iTarget).Col = 6:  sprGrid(iTarget).Text = tData(4)
            sprGrid(iTarget).Col = 7:  sprGrid(iTarget).Text = tData(5)
            sprGrid(iTarget).Col = 8:  sprGrid(iTarget).Text = tData(6)
            sprGrid(iTarget).Col = 9:  sprGrid(iTarget).Text = tData(7)
            sprGrid(iTarget).Col = 10: sprGrid(iTarget).Text = tData(8)
        Next i
        
        .MaxRows = 0
    End With
        
    sprGrid(iTarget).SortKey(1) = 2
    sprGrid(iTarget).SortKeyOrder(1) = SortKeyOrderAscending
    sprGrid(iTarget).Sort -1, -1, -1, -1, SortByRow
        
    lblCount(iSource).Caption = sprGrid(iSource).MaxRows
    lblCount(iTarget).Caption = sprGrid(iTarget).MaxRows
    
    txtTag(0).SetFocus
    
    Exit Sub
    
ErrRtn:

End Sub

Private Sub cmdOut_Click()
    Dim 합계금액 As Currency
    
    Dim 접수일자 As String
    Dim 택번호   As String
    
    On Error GoTo ErrRtn
    
    If sprGrid(1).MaxRows = 0 Then
        MsgBox "지점출고 세탁물이 없습니다.", vbInformation, "확인"
        
        Exit Sub
    End If
    
    ProgressBar1.Visible = True
    ProgressBar1.Min = 0
    ProgressBar1.MAX = 100
    DoEvents
        
    With sprGrid(1)
        For i = 1 To .MaxRows
            .Row = i
            
            .Col = 2: 접수일자 = Format(.Text, "YYYY-MM-DD")
            .Col = 4: 택번호 = Replace(.Text, "-", "")
            
            Query = "UPDATE TB_입출고 SET 가맹점출고일자 = '" & Format(Now, "YYYY-MM-DD hh:mm:ss") & "'"
            Query = Query & "           , 본사전송여부   = ''"
            Query = Query & " WHERE 접수일자 = '" & 접수일자 & "'"
            Query = Query & "   AND 택번호   = '" & 택번호 & "'"
            ADOCon.Execute Query
            
            ProgressBar1.Value = (i / .MaxRows) * 100
            DoEvents
        Next i
    End With
    
    ProgressBar1.Visible = False
    DoEvents
    
    '-----------------------------------------------------------------------------------
    '
    '-----------------------------------------------------------------------------------
    Rtn = MsgBox("출고현황을 출력하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "확인")
    
    If Rtn = vbYes Then
        합계금액 = 0
                    
        '---------------------------------------------------------------------------------------
        Dim FileNum

        FileNum = FreeFile

        Open AppPath & "XML\가맹점출고.XML" For Output As #FileNum

        Print #FileNum, "<?xml version=""1.0"" encoding=""EUC-KR""?>"
        Print #FileNum, "<root>"

              XML = "    <조건>"
        XML = XML & "        <가맹점>(" & 가맹점정보.가맹점명 & ") 가맹점 출고현황</가맹점>"
        XML = XML & "        <출고일자>출고일자 : " & Format(Date, "YYYY년 MM월 DD일") & "</출고일자>"
        XML = XML & "   </조건>"
        Print #FileNum, XML

        With sprGrid(1)
            For i = 1 To .MaxRows
                .Row = i

                                 XML = "    <Data>"
                .Col = 2:  XML = XML & "        <접수일자>" & .Text & "</접수일자>"
                .Col = 3:  XML = XML & "        <품명>" & Func_Replace(.Text) & "</품명>"
                .Col = 4:  XML = XML & "        <택번호>" & .Text & "</택번호>"
                .Col = 5:  XML = XML & "        <색상>" & Func_Replace(.Text) & "</색상>"
                .Col = 6:  XML = XML & "        <무늬>" & Func_Replace(.Text) & "</무늬>"
                .Col = 7:  XML = XML & "        <내용>" & Func_Replace(.Text) & "</내용>"
                .Col = 8:  XML = XML & "        <금액>" & Func_Replace(.Text) & "</금액>"
                .Col = 9:  XML = XML & "        <상표>" & Func_Replace(.Text) & "</상표>"
                           XML = XML & "   </Data>"
                
                .Col = 8: 합계금액 = 합계금액 + CCur(.Value)

                Print #FileNum, XML
            Next i

                  XML = "    <합계>"
            XML = XML & "        <출고량>출 고 량 : " & sprGrid(1).MaxRows & "</출고량>"
            XML = XML & "        <합계금액>합계금액 : " & Format(합계금액, "#,##0") & "</합계금액>"
            XML = XML & "   </합계>"
            Print #FileNum, XML
        End With
            
        Print #FileNum, "</root>"
        Close #FileNum

        With rpt가맹점출고
            .dc.FileURL = AppPath & "XML\가맹점출고.XML"
            .PrintReport True
            '.Show 1
        End With
        
        Unload rpt가맹점출고
    End If
                            
    sprGrid(1).MaxRows = 0
    lblCount(1).Caption = "0"
    txtTag(0).SetFocus
    
    Exit Sub
    
ErrRtn:
'    Close #FileNum

    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    For i = 0 To 1
        With sprGrid(i)
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
        End With
    Next i
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    pnlHeader.Width = Me.ScaleWidth
    
    cmdBtn(2).Left = (Me.Width - 200) - cmdBtn(2).Width
    cmdOut.Left = (Me.Width - 200) - cmdOut.Width
    
    ProgressBar1.Left = (Me.Width - 200) - ProgressBar1.Width
End Sub

Private Sub txtTag_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 And KeyAscii = 13 Then
        With sprGrid(0)
            For i = 1 To .MaxRows
                .Row = i
                .Col = 2
                If .Text = Trim(txtTag(0).Text) Then
                    .Col = 2:  tData(0) = .Text
                    .Col = 3:  tData(1) = .Text
                    .Col = 4:  tData(2) = .Text
                    .Col = 5:  tData(3) = .Text
                    .Col = 6:  tData(4) = .Text
                    .Col = 7:  tData(5) = .Text
                    .Col = 8:  tData(6) = .Text
                    .Col = 9:  tData(7) = .Text
                    .Col = 10: tData(8) = .Text
                    
                    .DeleteRows i, 1
                    .MaxRows = .MaxRows - 1
                    i = i - 1
                    
                    '
                    sprGrid(1).MaxRows = sprGrid(1).MaxRows + 1
                    sprGrid(1).Row = sprGrid(1).MaxRows
                    
                    sprGrid(1).Col = 1:  sprGrid(1).Text = "0"
                    sprGrid(1).Col = 2:  sprGrid(1).Text = tData(0)
                    sprGrid(1).Col = 3:  sprGrid(1).Text = tData(1)
                    sprGrid(1).Col = 4:  sprGrid(1).Text = tData(2)
                    sprGrid(1).Col = 5:  sprGrid(1).Text = tData(3)
                    sprGrid(1).Col = 6:  sprGrid(1).Text = tData(4)
                    sprGrid(1).Col = 7:  sprGrid(1).Text = tData(5)
                    sprGrid(1).Col = 8:  sprGrid(1).Text = tData(6)
                    sprGrid(1).Col = 9:  sprGrid(1).Text = tData(7)
                    sprGrid(1).Col = 10: sprGrid(1).Text = tData(8)
                End If
            Next i
        End With
        
        lblCount(0).Caption = sprGrid(0).MaxRows
        lblCount(1).Caption = sprGrid(1).MaxRows
        
        txtTag(0).Text = ""
        txtTag(0).SetFocus
    End If
End Sub
