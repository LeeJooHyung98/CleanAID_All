VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMask32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_04005 
   Caption         =   "수금 주마감"
   ClientHeight    =   10665
   ClientLeft      =   1395
   ClientTop       =   3270
   ClientWidth     =   16365
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_04005.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10665
   ScaleWidth      =   16365
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   10665
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16365
      _ExtentX        =   28866
      _ExtentY        =   18812
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_04005.frx":058A
      Begin Threed.SSPanel SSPanel 
         Height          =   9315
         Index           =   0
         Left            =   15
         TabIndex        =   1
         Top             =   1335
         Width           =   16335
         _ExtentX        =   28813
         _ExtentY        =   16431
         _Version        =   262144
         BackColor       =   16777215
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSFrame SSFrame 
            Height          =   1485
            Left            =   60
            TabIndex        =   10
            Top             =   1575
            Width           =   5490
            _ExtentX        =   9684
            _ExtentY        =   2619
            _Version        =   262144
            Caption         =   "마감진행정보"
            Begin VB.TextBox txtInput 
               Enabled         =   0   'False
               Height          =   315
               Index           =   3
               Left            =   1590
               TabIndex        =   12
               Top             =   675
               Width           =   3780
            End
            Begin VB.TextBox txtInput 
               Enabled         =   0   'False
               Height          =   315
               Index           =   4
               Left            =   1590
               TabIndex        =   11
               Top             =   315
               Width           =   3780
            End
            Begin MSMask.MaskEdBox mskInput 
               Height          =   315
               Index           =   2
               Left            =   1590
               TabIndex        =   13
               Top             =   1035
               Width           =   3780
               _ExtentX        =   6668
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   "#,##0"
               PromptChar      =   "_"
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   7
               Left            =   120
               TabIndex        =   14
               Top             =   315
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "대리점 코드"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   8
               Left            =   120
               TabIndex        =   15
               Top             =   675
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "대리점 명"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel panCaption 
               Height          =   315
               Index           =   9
               Left            =   120
               TabIndex        =   16
               Top             =   1035
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   262144
               Caption         =   "진 행 율"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
            End
         End
         Begin VB.TextBox txtInput 
            Enabled         =   0   'False
            Height          =   315
            Index           =   5
            Left            =   1530
            TabIndex        =   2
            Top             =   45
            Width           =   3015
         End
         Begin MSMask.MaskEdBox mskInput 
            Height          =   315
            Index           =   1
            Left            =   1530
            TabIndex        =   3
            Top             =   1125
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   "#,##0"
            PromptChar      =   "_"
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   0
            Left            =   1530
            TabIndex        =   4
            Top             =   405
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CheckBox        =   -1  'True
            Format          =   64290816
            CurrentDate     =   36686
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   1
            Left            =   1530
            TabIndex        =   5
            Top             =   765
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CheckBox        =   -1  'True
            Format          =   64290816
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   60
            TabIndex        =   6
            Top             =   405
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "마감 시작일"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   3
            Left            =   60
            TabIndex        =   7
            Top             =   765
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "마감 종료일"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   4
            Left            =   60
            TabIndex        =   8
            Top             =   1125
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "총마감 건수"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   10
            Left            =   60
            TabIndex        =   9
            Top             =   45
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "주"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel panInput 
         Height          =   780
         Left            =   15
         TabIndex        =   17
         Top             =   540
         Width           =   16335
         _ExtentX        =   28813
         _ExtentY        =   1376
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   2
            Left            =   5790
            MaxLength       =   2
            TabIndex        =   19
            Top             =   60
            Width           =   1095
         End
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   1
            Left            =   4410
            MaxLength       =   2
            TabIndex        =   18
            Top             =   60
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Index           =   2
            Left            =   1530
            TabIndex        =   20
            Top             =   60
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy"
            Format          =   64290819
            UpDown          =   -1  'True
            CurrentDate     =   37140
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   21
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "연    도"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   2940
            TabIndex        =   22
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "마  감  주"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   2  '가운데 맞춤
            Caption         =   "~"
            Height          =   195
            Left            =   5550
            TabIndex        =   23
            Top             =   120
            Width           =   195
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   510
         Left            =   15
         TabIndex        =   24
         Top             =   15
         Width           =   8730
         _ExtentX        =   15399
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
         PictureBackground=   "P_04005.frx":061C
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel 
         Height          =   510
         Index           =   1
         Left            =   8760
         TabIndex        =   25
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
         PictureBackground=   "P_04005.frx":081E
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   7
            Left            =   6660
            TabIndex        =   26
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
            Picture         =   "P_04005.frx":0A20
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   6
            Left            =   5730
            TabIndex        =   27
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
            Picture         =   "P_04005.frx":0FBA
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   5
            Left            =   4800
            TabIndex        =   28
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
            Picture         =   "P_04005.frx":1554
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   4
            Left            =   3750
            TabIndex        =   29
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
            Picture         =   "P_04005.frx":1AEE
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   3
            Left            =   2820
            TabIndex        =   30
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
            Picture         =   "P_04005.frx":2088
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   2
            Left            =   1890
            TabIndex        =   31
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
            Picture         =   "P_04005.frx":2622
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   1
            Left            =   960
            TabIndex        =   32
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
            Picture         =   "P_04005.frx":2BBC
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   450
            Index           =   0
            Left            =   30
            TabIndex        =   33
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
            Picture         =   "P_04005.frx":3156
         End
      End
   End
End
Attribute VB_Name = "P_04005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim RS02 As ADODB.Recordset
Dim RS03 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cmdBtn_Click(Index As Integer)
    On Error GoTo ErrRtn
    
    Select Case Index
        Case 0: 'Call Data_Display   ' 조회
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
    cmdBtn(2).Enabled = True
    
    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    If P_04005_Flag = False Then
        dtInput(2).Value = Date
        
        P_04005_Flag = True
    End If
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Public Sub DataSave()
    Dim i As Integer
    
    Dim SSQL As String
    Dim sDate As String
    Dim sDate2 As String
    
    On Error GoTo SQLERROR
    
    If MsgBox("마감작업을 진행하시겠습니까?", vbQuestion + vbYesNo) = vbYes Then
        SSQL = "SELECT A.Week      AS 주, "
        SSQL = SSQL & "      A.sDate        AS 시작일, "
        SSQL = SSQL & "      A.eDate        AS 종료일 "
        SSQL = SSQL & "FROM    WeekMST     A (NOLOCK) "
        SSQL = SSQL & "WHERE   A.WYear     =   '" & Format(dtInput(2).Value, "yyyy") & "' "
        SSQL = SSQL & "AND     A.Week      BETWEEN '" & txtInput(1).Text & "' AND '" & txtInput(2).Text & "' "

        Set RS01 = New ADODB.Recordset
        RS01.Open SSQL, ADOCon, adOpenStatic
    
        Do While Not RS01.EOF
            txtInput(5).Text = RS01!주
            sDate = Format(dtInput(2).Value, "yyyy") & RS01!시작일
            dtInput(0).Value = Format(sDate, "####-##-##")
            sDate2 = Format(dtInput(2).Value, "yyyy") & RS01!종료일
            dtInput(1).Value = Format(sDate2, "####-##-##")
    
            SSQL = "SELECT  A.AgencyCode        AS 대리점코드, "
            SSQL = SSQL & "      B.AgencyName       AS 대리점명, "
            SSQL = SSQL & "      Sum(A.IpSu)        AS 입고수량, "
            SSQL = SSQL & "      Sum(A.ChulSu)      AS 출고수량, "
            SSQL = SSQL & "      Min(A.StartTag)    AS 시작택, "
            SSQL = SSQL & "      Max(A.EndTag)      AS 종료택, "
            SSQL = SSQL & "      Sum(A.Amount)      AS 금액, "
            SSQL = SSQL & "      Sum(A.JaeSu)       AS 재세탁수량, "
            SSQL = SSQL & "      Sum(A.SuSu)        AS 수선수량, "
            SSQL = SSQL & "      Sum (A.BanSu)      AS 반품수량 "
            SSQL = SSQL & "FROM    Sugeum      A (NOLOCK), "
            SSQL = SSQL & "        AgencyCT    B (NOLOCK) "
            SSQL = SSQL & "WHERE   A.AgencyCode = B.AgencyCode "
            SSQL = SSQL & "AND     A.SuDate BETWEEN '" & sDate & "' AND '" & sDate2 & "' "
            SSQL = SSQL & "GROUP BY    A.AgencyCode, "
            SSQL = SSQL & "            B.AgencyName "
            
            Set RS02 = New ADODB.Recordset
            RS02.Open SSQL, ADOCon, adOpenStatic
            
            mskInput(1).Text = RS02.RecordCount
            
            Do While Not RS02.EOF
                txtInput(3).Text = RS02!대리점코드
                txtInput(4).Text = RS02!대리점명
                
                i = i + 1
                mskInput(2).Text = i
                
                DoEvents
            
                SSQL = "SELECT  Count(*)    AS 레코드건수 "
                SSQL = SSQL & "FROM    SuGeumWK   A (NOLOCK) "
                SSQL = SSQL & "WHERE   A.Year       = '" & Format(dtInput(2).Value, "yyyy") & "' "
                SSQL = SSQL & "AND     A.Week       = '" & txtInput(5).Text & "' "
                SSQL = SSQL & "AND     A.AgencyCode = '" & RS02!대리점코드 & "' "
                
                Set RS03 = New ADODB.Recordset
                RS03.Open SSQL, ADOCon, adOpenStatic
                
                If RS03!레코드건수 = 0 Then
                    SSQL = "INSERT INTO SuGeumWK "
                    SSQL = SSQL & "    (Year, "
                    SSQL = SSQL & "     Week, "
                    SSQL = SSQL & "     AgencyCode, "
                    SSQL = SSQL & "     ISu, "
                    SSQL = SSQL & "     ChulSu, "
                    SSQL = SSQL & "     STag, "
                    SSQL = SSQL & "     ETag, "
                    SSQL = SSQL & "     Amount, "
                    SSQL = SSQL & "     JSu, "
                    SSQL = SSQL & "     SSu, "
                    SSQL = SSQL & "     BSu) "
                    SSQL = SSQL & "VALUES ('" & Format(dtInput(2).Value, "yyyy") & "', "
                    SSQL = SSQL & "        '" & txtInput(5).Text & "', "
                    SSQL = SSQL & "        '" & RS02!대리점코드 & "', "
                    SSQL = SSQL & "        " & RS02!입고수량 & ", "
                    SSQL = SSQL & "        " & RS02!출고수량 & ", "
                    SSQL = SSQL & "        '" & RS02!시작택 & "', "
                    SSQL = SSQL & "        '" & RS02!종료택 & "', "
                    SSQL = SSQL & "        " & RS02!금액 & ", "
                    SSQL = SSQL & "        " & RS02!재세탁수량 & ","
                    SSQL = SSQL & "        " & RS02!수선수량 & ", "
                    SSQL = SSQL & "        " & RS02!반품수량 & ") "
                Else
                    SSQL = "UPDATE SuGeumWK "
                    SSQL = SSQL & "SET ISu         =   " & RS02!입고수량 & ", "
                    SSQL = SSQL & "    ChulSu      =   " & RS02!출고수량 & ", "
                    SSQL = SSQL & "    STag        =   '" & RS02!시작택 & "', "
                    SSQL = SSQL & "    ETag        =   '" & RS02!종료택 & "', "
                    SSQL = SSQL & "    Amount      =   " & RS02!금액 & ", "
                    SSQL = SSQL & "    JSu         =   " & RS02!재세탁수량 & ", "
                    SSQL = SSQL & "    SSu         =   " & RS02!수선수량 & ", "
                    SSQL = SSQL & "    BSu         =   " & RS02!반품수량 & " "
                    SSQL = SSQL & "WHERE   Year       = '" & Format(dtInput(2).Value, "yyyy") & "' "
                    SSQL = SSQL & "AND     Week       = '" & txtInput(5).Text & "' "
                    SSQL = SSQL & "AND     AgencyCode = '" & RS02!대리점코드 & "' "
                End If
                
                ADOCon.Execute SSQL
            
                RS02.MoveNext
            Loop
            
            RS01.MoveNext
        Loop
    End If
    
    MsgBox "해당되는 데이터가 정상적으로 처리되었습니다.", vbInformation
    Exit Sub
    
SQLERROR:
    If Err.Number <> 0 Then
        MsgBox "[" & Err.Number & "] " & Err.Description
        Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_04005_Flag = False
End Sub

