VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RichTx32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm메일작성 
   BorderStyle     =   1  '단일 고정
   Caption         =   "메일 작성"
   ClientHeight    =   8070
   ClientLeft      =   3555
   ClientTop       =   5775
   ClientWidth     =   9510
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm메일작성.frx":0000
   LinkTopic       =   "Form24"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   9510
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8070
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   14235
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frm메일작성.frx":0A02
      Begin Threed.SSPanel SSPanel1 
         Height          =   630
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   9480
         _ExtentX        =   16722
         _ExtentY        =   1111
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSPanel SSPanel 
            Height          =   510
            Left            =   60
            TabIndex        =   4
            Top             =   60
            Width           =   2490
            _ExtentX        =   4392
            _ExtentY        =   900
            _Version        =   262144
            BackColor       =   16777215
            Enabled         =   0   'False
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin MSComCtl2.DTPicker dtpDay 
               Height          =   315
               Left            =   960
               TabIndex        =   5
               Top             =   105
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   556
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   59441155
               CurrentDate     =   40279
            End
            Begin VB.Label Label2 
               Alignment       =   1  '오른쪽 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "작성일자:"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   2
               Left            =   120
               TabIndex        =   6
               Top             =   165
               Width           =   810
            End
         End
         Begin XtremeSuiteControls.PushButton cmdBtn 
            Height          =   540
            Left            =   8025
            TabIndex        =   3
            Top             =   45
            Width           =   1395
            _Version        =   851970
            _ExtentX        =   2461
            _ExtentY        =   952
            _StockProps     =   79
            Caption         =   " 저장"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frm메일작성.frx":0A54
         End
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   7395
         Left            =   15
         TabIndex        =   2
         Top             =   660
         Width           =   9480
         _ExtentX        =   16722
         _ExtentY        =   13044
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"frm메일작성.frx":1466
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frm메일작성"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBtn_Click()
    On Error GoTo ErrRtn
    
    Dim iSEQ  As Long
    Dim sDate As String
    
    sDate = Format(Date, "YYYY-MM-DD")
    
    Query = "SELECT ISNULL(MAX(문서번호),0) + 1"
    Query = Query & "FROM TB_공지사항 "
    Query = Query & " WHERE 작성일자 = '" & sDate & "' "
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenStatic, adLockOptimistic
    
    iSEQ = ADORs(0)
    
    ADORs.Close
    Set ADORs = Nothing
    
    '------------------------------------------------------
    '
    '------------------------------------------------------
    Query = "INSERT INTO TB_공지사항 ("
    Query = Query & "  가맹점코드"
    Query = Query & ", 공지구분"
    Query = Query & ", 작성일자"
    Query = Query & ", 문서번호"
    Query = Query & ", 종료일자"
    Query = Query & ", 공지내용"
    Query = Query & ", 수신여부"
    Query = Query & ", 수신일자"
    Query = Query & ", 본사전송여부"
    Query = Query & ") VALUES ("
    Query = Query & "  '" & 가맹점정보.가맹점코드 & "'" '
    Query = Query & ", '1'"                         '
    Query = Query & ", '" & sDate & "'"             '
    Query = Query & ",  " & iSEQ                    '
    Query = Query & ", ''"                          '
    Query = Query & ", '" & RichTextBox1.Text & "'" '
    Query = Query & ", ''"                          '
    Query = Query & ", ''"                          '
    Query = Query & ", 'N')"                        '
    ADOCon.Execute Query
    
    If Server_Connection(HostCon, "LAUNDRY1000") = True Then
        Dim sValue() As String
        
        Dim Err_Num As Long
        Dim Err_Desc As String

        ReDim sValue(7)
        
        sValue(0) = "1"                    ' 공지구분
        sValue(1) = sDate                  ' 작성일자
        sValue(2) = ""                     ' 시작일자
        sValue(3) = ""                     ' 종료일자
        sValue(4) = 가맹점정보.가맹점코드  ' 가맹점코드
        sValue(5) = iSEQ                   ' 문서번호
        sValue(6) = RichTextBox1.Text & "" ' 공지내용
        sValue(7) = "Y"                    ' 전송여부
        
        Call SP_Exec(HostCon, "SP_M_09004_02", sValue(), Err_Num, Err_Desc)
        
        HostCon.Close
        Set HostCon = Nothing
    End If
    
    MsgBox "공지내용이 저장이 되었습니다.", vbInformation
    
    Unload Me
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)

    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    dtpDay.Value = Format(Date, "YYYY-MM-DD")
    
    
    
    'TitleSet "편 지 작 성"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    cmdBtn.Left = Me.ScaleWidth - cmdBtn.Width - 100
End Sub
