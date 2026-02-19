VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form P_08001 
   Caption         =   "자료 송수신 (HANDY/DISK)"
   ClientHeight    =   8295
   ClientLeft      =   390
   ClientTop       =   2460
   ClientWidth     =   16995
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_08001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8295
   ScaleWidth      =   16995
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16995
      _ExtentX        =   29977
      _ExtentY        =   14631
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_08001.frx":058A
      Begin Threed.SSPanel panInput 
         Height          =   1185
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   16965
         _ExtentX        =   29924
         _ExtentY        =   2090
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   11835
            Top             =   90
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin XtremeSuiteControls.PushButton cmdSubBtn 
            Height          =   690
            Index           =   0
            Left            =   60
            TabIndex        =   12
            Top             =   435
            Width           =   1725
            _Version        =   851970
            _ExtentX        =   3043
            _ExtentY        =   1217
            _StockProps     =   79
            Caption         =   "입고자료수신"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            Appearance      =   6
         End
         Begin VB.ComboBox cboHT_Gubun 
            Height          =   315
            ItemData        =   "P_08001.frx":067C
            Left            =   1875
            List            =   "P_08001.frx":068C
            Style           =   2  '드롭다운 목록
            TabIndex        =   2
            Top             =   60
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Left            =   7395
            TabIndex        =   3
            Top             =   60
            Visible         =   0   'False
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   64552960
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   3
            Left            =   5580
            TabIndex        =   4
            Top             =   60
            Visible         =   0   'False
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "수선적용일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   5
            Top             =   60
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "핸드터미널 종류"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton cmdSubBtn 
            Height          =   690
            Index           =   1
            Left            =   1845
            TabIndex        =   13
            Top             =   435
            Width           =   1725
            _Version        =   851970
            _ExtentX        =   3043
            _ExtentY        =   1217
            _StockProps     =   79
            Caption         =   "핸디터미널수신"
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
            Picture         =   "P_08001.frx":06BB
         End
         Begin XtremeSuiteControls.PushButton cmdSubBtn 
            Height          =   690
            Index           =   2
            Left            =   3630
            TabIndex        =   14
            Top             =   435
            Width           =   1725
            _Version        =   851970
            _ExtentX        =   3043
            _ExtentY        =   1217
            _StockProps     =   79
            Caption         =   "출고자료송신"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            Appearance      =   6
            Picture         =   "P_08001.frx":0C55
         End
         Begin XtremeSuiteControls.PushButton cmdSubBtn 
            Height          =   690
            Index           =   3
            Left            =   5415
            TabIndex        =   15
            Top             =   435
            Width           =   1725
            _Version        =   851970
            _ExtentX        =   3043
            _ExtentY        =   1217
            _StockProps     =   79
            Caption         =   "수신자료초기화"
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
            Picture         =   "P_08001.frx":11EF
         End
         Begin XtremeSuiteControls.PushButton cmdSubBtn 
            Height          =   690
            Index           =   4
            Left            =   7200
            TabIndex        =   16
            Top             =   435
            Width           =   1725
            _Version        =   851970
            _ExtentX        =   3043
            _ExtentY        =   1217
            _StockProps     =   79
            Caption         =   "할인자료송신"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            Appearance      =   6
         End
         Begin XtremeSuiteControls.PushButton cmdSubBtn 
            Height          =   690
            Index           =   5
            Left            =   8985
            TabIndex        =   17
            Top             =   435
            Width           =   1725
            _Version        =   851970
            _ExtentX        =   3043
            _ExtentY        =   1217
            _StockProps     =   79
            Caption         =   "목요세일자료송신"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            Appearance      =   6
         End
         Begin XtremeSuiteControls.PushButton cmdSubBtn 
            Height          =   690
            Index           =   6
            Left            =   10770
            TabIndex        =   18
            Top             =   435
            Width           =   1725
            _Version        =   851970
            _ExtentX        =   3043
            _ExtentY        =   1217
            _StockProps     =   79
            Caption         =   "대리점품목송신"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            Appearance      =   6
         End
         Begin XtremeSuiteControls.PushButton cmdSubBtn 
            Height          =   690
            Index           =   7
            Left            =   12555
            TabIndex        =   19
            Top             =   435
            Width           =   1725
            _Version        =   851970
            _ExtentX        =   3043
            _ExtentY        =   1217
            _StockProps     =   79
            Caption         =   "수신자료송신"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            Appearance      =   6
         End
      End
      Begin Threed.SSPanel panTitle 
         Height          =   390
         Index           =   0
         Left            =   15
         TabIndex        =   6
         Top             =   1215
         Width           =   16965
         _ExtentX        =   29924
         _ExtentY        =   688
         _Version        =   262144
         Font3D          =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " 정상 출고 자료"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_08001.frx":1789
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdSubButton 
            Height          =   360
            Index           =   0
            Left            =   13755
            TabIndex        =   20
            Top             =   15
            Width           =   1155
            _Version        =   851970
            _ExtentX        =   2037
            _ExtentY        =   635
            _StockProps     =   79
            Caption         =   "행삭제"
            UseVisualStyle  =   -1  'True
            Picture         =   "P_08001.frx":1BEB
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   1740
         Index           =   0
         Left            =   15
         TabIndex        =   7
         Top             =   1620
         Width           =   16965
         _Version        =   524288
         _ExtentX        =   29924
         _ExtentY        =   3069
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         EditEnterAction =   5
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
         MaxCols         =   10
         SpreadDesigner  =   "P_08001.frx":2185
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panTitle 
         Height          =   390
         Index           =   1
         Left            =   15
         TabIndex        =   8
         Top             =   3375
         Width           =   16965
         _ExtentX        =   29924
         _ExtentY        =   688
         _Version        =   262144
         Font3D          =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " 선택 출고 자료"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_08001.frx":27C3
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdSubButton 
            Height          =   360
            Index           =   1
            Left            =   13755
            TabIndex        =   21
            Top             =   15
            Width           =   1155
            _Version        =   851970
            _ExtentX        =   2037
            _ExtentY        =   635
            _StockProps     =   79
            Caption         =   "행삭제"
            UseVisualStyle  =   -1  'True
            Picture         =   "P_08001.frx":2C25
         End
         Begin XtremeSuiteControls.PushButton cmdSubButton 
            Height          =   360
            Index           =   3
            Left            =   11625
            TabIndex        =   22
            Top             =   15
            Width           =   2115
            _Version        =   851970
            _ExtentX        =   3731
            _ExtentY        =   635
            _StockProps     =   79
            Caption         =   "미선택 삭제(전체)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Picture         =   "P_08001.frx":31BF
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   2400
         Index           =   1
         Left            =   15
         TabIndex        =   9
         Top             =   3780
         Width           =   16965
         _Version        =   524288
         _ExtentX        =   29924
         _ExtentY        =   4233
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         EditEnterAction =   5
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
         MaxCols         =   11
         MaxRows         =   50
         SpreadDesigner  =   "P_08001.frx":3759
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
      Begin Threed.SSPanel panTitle 
         Height          =   390
         Index           =   2
         Left            =   15
         TabIndex        =   10
         Top             =   6195
         Width           =   16965
         _ExtentX        =   29924
         _ExtentY        =   688
         _Version        =   262144
         Font3D          =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " 기존 출고 자료"
         PictureBackgroundStyle=   2
         PictureBackground=   "P_08001.frx":3E1E
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdSubButton 
            Height          =   360
            Index           =   2
            Left            =   13755
            TabIndex        =   23
            Top             =   15
            Width           =   1155
            _Version        =   851970
            _ExtentX        =   2037
            _ExtentY        =   635
            _StockProps     =   79
            Caption         =   "행삭제"
            UseVisualStyle  =   -1  'True
            Picture         =   "P_08001.frx":4280
         End
      End
      Begin FPSpreadADO.fpSpread spdView 
         Height          =   1680
         Index           =   2
         Left            =   15
         TabIndex        =   11
         Top             =   6600
         Width           =   16965
         _Version        =   524288
         _ExtentX        =   29924
         _ExtentY        =   2963
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         EditEnterAction =   5
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
         MaxCols         =   10
         SpreadDesigner  =   "P_08001.frx":481A
         HighlightHeaders=   1
         HighlightStyle  =   1
      End
   End
End
Attribute VB_Name = "P_08001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS01 As ADODB.Recordset
Dim sValue() As String
Public saveYN As Boolean

Dim Err_Num As Long
Dim Err_Dec As String

Private Sub cboHT_Gubun_Click()
    SaveSetting REG_App, Me.Name, "HT_Gubun", CStr(cboHT_Gubun.ListIndex)

End Sub

Private Sub cmdSubBtn_Click(Index As Integer)
    Select Case Index
        Case 0          ' 입고자료수신
            Call DataSave1
        Case 1          ' 핸드터미널수신
            Call DataSave3
        Case 2          ' 출고자료송신
            Call DataSave4
        Case 3          ' 수신자료초기화
            Call DataSave5
        Case 4          ' 할인자료송신
            P_08001_01.Show 1
        Case 5          ' 목요세일자료송신
            P_08001_02.Show 1
        Case 6          ' 대리점품목송신
            P_08001_03.Show 1
        Case 7          ' 수선자료송신
            Call DataSave6
    End Select
End Sub

Private Sub BtnDisable()
    cmdSubBtn(0).Enabled = False
    cmdSubBtn(1).Enabled = False
    cmdSubBtn(2).Enabled = False
    cmdSubBtn(3).Enabled = False
    cmdSubBtn(4).Enabled = False
    cmdSubBtn(5).Enabled = False
    cmdSubBtn(6).Enabled = False
    cmdSubBtn(7).Enabled = False
End Sub

Private Sub BtnEnable()
    cmdSubBtn(0).Enabled = True
    cmdSubBtn(1).Enabled = True
    cmdSubBtn(2).Enabled = True
    cmdSubBtn(3).Enabled = True
    cmdSubBtn(4).Enabled = True
    cmdSubBtn(5).Enabled = True
    cmdSubBtn(6).Enabled = True
    cmdSubBtn(7).Enabled = True
End Sub

Private Sub DataSave1()
    Dim AgencyName As String
    Dim FileName As String
    Dim AgencyCode As String
    Dim DataType As String
    Dim AllFileName() As String
    Dim i As Integer
    Dim j As Integer
    
    Call BtnDisable
    
    MsgBox "A:드라이버에 대리점용 디스켓을 넣으십시요"
    
    While Not DriverChk
        If MsgBox("A:드라이버에 디스켓을 넣으십시요", vbRetryCancel) = vbCancel Then
            Call BtnEnable
            
            Exit Sub
        End If
    Wend
    
    FileName = Dir("A:*.dat")
    
    If FileName = "" Then
        MsgBox "A:드라이버에 데이타화일이 없습니다."
        Exit Sub
    End If
    
    i = 0
    
    ' 여러개의 파일을 배열에 넣어준다.
    Do While Len(FileName) > 0
        ReDim Preserve AllFileName(0 To i)
        
        AllFileName(i) = FileName
        FileName = Dir
        i = i + 1
    Loop
    
    ' 배열을 순서대로 읽어서 입고데이터를 처리한다.
    For j = 0 To i - 1
        FileName = AllFileName(j)
        
        ' 파일명의 확장자는 대리점코드
        AgencyCode = Mid(FileName, 10, 3)
        AgencyName = GetAgencyName(AgencyCode)
        
        If Mid(FileName, 14, 1) = "1" Then
            Call DataSave2(FileName)
        End If
        
    Next j
    
    MsgBox "입고자료 수신이 완료되었습니다."
    
    Call BtnEnable
End Sub

Private Sub DataSave2(FileName As String)
    Dim iCnt As Integer
    Dim uCnt As Integer
    Dim Str As String
    Dim tCnt As Integer
    Dim TempStr As String
    Dim strCode As String
    Dim strDate As String
    
    Dim FilePath As String
    Dim BCPPath As String
    
    Open "A:" & FileName For Input As #1
    
    ' A드라이브로 부터 읽은 내역을 C드라이브로 Data를 이동한다.
    FilePath = GetIniStr("TEXT DATA", "ReceiveJobFilePath", "", m_iniFile)
    BCPPath = GetIniStr("TEXT DATA", "BCPPath", "", m_iniFile)
    
    Open FilePath & "\IpChul.Dat" For Output As #2
    
    iCnt = 0
    
    Line Input #1, TempStr
    
    ' Data가 일일마감이면
    If Not Mid(TempStr, 1, 4) = "일일마감" Then
       TempStr = ""
       
       Close #1
       Open "A:" & FileName For Input As #1
    End If
    
    ' 디스켓 데이터를 일어서 BCP로 사용할 Data File을 만든다.
    While Not EOF(1)
        Line Input #1, Str
        
        ' 읽어온 데이타중 끝의 두자리(CR,LF)를 삭제한다.
        tCnt = Len(Trim(Str)) - 2
        
        Str = Mid(Str, 1, tCnt) & " ||"
        
        ' 일일마감이면 끝에 '|4'를 그렇치 않으면 '|3'을 넣어준다.
        If TempStr = "" Then
           Print #2, Str & "|4"
        Else
           Print #2, Str & "|3"
        End If
        
        iCnt = iCnt + 1
    Wend
    
    Close #1
    Close #2
    
    ' 임시 입출고 테이블의 내역을 삭제한다.
    ReDim sValue(0)
    sValue(0) = "0"
    Call ExecPro("SP_08001_00", sValue(), Err_Num, Err_Dec)
    
    ' BCP를 사용하여서 Text Data File을 DB에 올린다.
    If iCnt > 0 Then
        ' 완료파일이 있으면 삭제한다.
        If Not Dir(BCPPath & "\OK.OK") = "" Then
            Kill BCPPath & "\OK.OK"
        End If
        
        ' BCP를 위한 배치파일을 실행한다.
        Shell BCPPath & "\IpChul.Bat", vbHide
        
        'Bat 화일 완료시 까지 대기
        Do While Dir(BCPPath & "\OK.OK") = ""
            DoEvents
        Loop
        
        Kill BCPPath & "\OK.OK"
        
        ReDim sValue(0)
        
        ' Temp Table -> 입고 Table에 등록한다.
        sValue(0) = "0"
        Call ExecPro("SP_08001_01", sValue(), Err_Num, Err_Dec)
        
        '일일수금 자료 생성
        If Not TempStr = "" Then
            ReDim sValue(9)
            
            sValue(0) = Mid(FileName, 1, 8)         ' 수금일자
            sValue(1) = Mid(FileName, 10, 3)        ' 대리점코드
            sValue(2) = Val(Mid(TempStr, 6, 4))     ' 입고수량
            sValue(3) = Mid(TempStr, 55, 4)         ' 시작택
            sValue(4) = Mid(TempStr, 60, 4)         ' 종료택
            sValue(5) = Val(Mid(TempStr, 35, 8))    ' 금액
            sValue(6) = Val(Mid(TempStr, 16, 4))    ' 재세탁수량
            sValue(7) = Val(Mid(TempStr, 21, 4))    ' 수선수량
            sValue(8) = Val(Mid(TempStr, 11, 4))    ' 반품수량
            sValue(9) = 0                           ' 출고수량
            
            Call ExecPro("SP_08001_02", sValue(), Err_Num, Err_Dec)
        Else
            ReDim sValue(0)
            
            sValue(0) = "1"
            
            Call ExecPro("SP_08001_03", sValue(), Err_Num, Err_Dec)
        End If
    End If
End Sub

Private Sub DataSave3()
    Dim iCnt As Integer
    Dim TmpStr As String
    Dim sInOut As String
    Dim sTagNo As String
    Dim sItem As String
    Dim sDate As String
    Dim sCode As String
    Dim mDate As String
    Dim bDate As String
    Dim sFlag As String
    
    Dim DupCnt As Integer
    Dim ErrorCnt As Integer
    
    Dim TerminalPath As String
    Dim TerminalRecvName As String
    
    ' 업무 특성상 이화면에서는 핸드 터미널을 읽기 때문에 DB 연결을 먼저 해주고 처리한다.]
    On Error Resume Next
    ADOCon.Close
    On Error GoTo 0
    
    If DBOpen_Laundry = False Then
        MsgBox "서버와 연결이 종료 되었습니다. 프로그램을 다시 실행 하여 주십시요.", vbInformation, "확인"
        End
        Exit Sub
    End If
    
    ' 기존 방식의 핸드 터미널일 경우
    If cboHT_Gubun.ListIndex = 0 Then
        P_TRANS.saveYN = False
        P_TRANS.Show 1
        P_TRANS.Hide
    
        If P_TRANS.saveYN = False Then
            Exit Sub
        End If
        
    ' DT-900일 경우
    ElseIf cboHT_Gubun.ListIndex = 1 Then
        saveYN = False
        Call P_DT900.SetMode(출고자료수신)
        P_DT900.Show 1
        
        If saveYN = False Then
            Exit Sub
        End If
            
    'PDA (symbol)
    ElseIf cboHT_Gubun.ListIndex = 2 Then
        If Dir(App.Path & "\PDA", vbDirectory) = "" Then
            MkDir App.Path & "\PDA"
        End If
        
        If Dir(App.Path & "\PDA\PDASEND1.TXT", vbDirectory) = "" Then
            Call PanelsMsg("pdasend1.txt 파일을 찾을 수 없습니다.")
            Exit Sub
        End If
    'PDA(PIDION)
    ElseIf cboHT_Gubun.ListIndex = 3 Then
    
        If Dir(App.Path & "\PDA", vbDirectory) = "" Then
            MkDir App.Path & "\PDA"
        End If
        
        If Dir(App.Path & "\PDA\kiSync.exe", vbDirectory) = "" Then
            Call PanelsMsg("PDA 복사 프로그램 파일을 찾을 수 없습니다.")
            Exit Sub
        End If
        
        If Dir(App.Path & "\PDA\kiSync.ini", vbDirectory) = "" Then
            Call PanelsMsg("PDA INI 파일을 찾을 수 없습니다.")
            Exit Sub
        End If
        '복사 루틴...
        
        P_PDA.optInput(1).Value = True
        P_PDA.txtInput(1).Text = App.Path & "\PDA"
        P_PDA.txtInput(2).Text = "CHULGO.DAT"
        P_PDA.Show 1
        'Text1.Text = DownPathName & "\" & DownFileName
        
        '복사 루틴...끝

        
        If Dir(App.Path & "\PDA\CHULGO.DAT", vbDirectory) = "" Then
            Call PanelsMsg("CHULGO.DAT 파일을 찾을 수 없습니다.")
            Exit Sub
        End If
    End If

    '중복검사 화면
    spdView(0).MaxRows = 0
    spdView(1).MaxRows = 0
    spdView(2).MaxRows = 0
      
    DupCnt = 0
    ErrorCnt = 0
    
    iCnt = 0
    
    '핸디로부터 읽은 데이타를 db로
    TerminalPath = GetIniStr("TERMINAL DATA", "TerminalFilePath", "", m_iniFile)
    TerminalRecvName = GetIniStr("TERMINAL DATA", "TerminalRecvName", "", m_iniFile)
    
    'PDA (symbol)
    If cboHT_Gubun.ListIndex = 2 Then
        FileCopy App.Path & "\PDA\PDASEND1.TXT", TerminalPath & "\" & TerminalRecvName
        Kill App.Path & "\PDA\PDASEND1.TXT"
    End If
    
    'PDA(PIDION)
    If cboHT_Gubun.ListIndex = 3 Then
        FileCopy App.Path & "\PDA\CHULGO.DAT", TerminalPath & "\" & TerminalRecvName
        Kill App.Path & "\PDA\CHULGO.DAT"
    End If
    
    Open TerminalPath & "\" & TerminalRecvName For Input As #1
    
    Do While Not EOF(1)
        DoEvents
        
        iCnt = iCnt + 1
        Input #1, TmpStr
        If Trim(TmpStr) = "" Then
            GoTo handy_err
        End If
        
        'PDA (symbol)
        If cboHT_Gubun.ListIndex = 2 Then
        '2005/10/08111311709720
            sInOut = Trim(Mid(TmpStr, 13, 1))                    ' 입출고구분 첫번째자리
            ' PDA 에서는 1. 정상 2. 반품으로 들어오기 때문에 2.정상,3.반품으로 변경해준다.
            If sInOut = "1" Then
                sInOut = "2"                    ' PDA에서 출고를 1로 처리해놓았다...
            ElseIf sInOut = "2" Then
                sInOut = "3"
            End If
            sDate = CStr(Trim(Replace(Mid(TmpStr, 1, 10), "/", ""))) ' dat파일에서 10 자리일자를 8자리로 변환
            mDate = Format(sDate, "####/##/##")
            sCode = Trim(Mid(TmpStr, 14, 3))                    ' dat파일에서 read
            sTagNo = Trim(Mid(TmpStr, 18, 4))                   ' ''
            sItem = Trim(Mid(TmpStr, 12, 1))                    ' 소품구분
            If sItem = "2" Then                                 ' "0" 정상, "1" 소품 처리됨
                sItem = "1"
            ElseIf sItem = "1" Then
                sItem = "0"
            End If
            
        'PDA(PIDION)
        ElseIf cboHT_Gubun.ListIndex = 3 Then
        '200811191629232110034007
            sInOut = Trim(Mid(TmpStr, 17, 1))                    ' 입출고구분 첫번째자리
            ' PDA 에서는 1. 정상 2. 반품으로 들어오기 때문에 2.정상,3.반품으로 변경해준다.
            If sInOut = "1" Then
                sInOut = "2"                    ' PDA에서 출고를 1로 처리해놓았다...
            ElseIf sInOut = "2" Then
                sInOut = "3"
            End If
            
            sDate = CStr(Trim(Mid(TmpStr, 1, 8)))               ' dat파일에서 8 자리일자
            mDate = Format(sDate, "####/##/##")
            sCode = Trim(Mid(TmpStr, 18, 3))                    ' dat파일에서 read
            sTagNo = Trim(Mid(TmpStr, 21, 4))                   ' ''
            sItem = Trim(Mid(TmpStr, 16, 1))                    ' 소품구분
            If sItem = "2" Then                                 ' "0" 정상, "1" 소품 처리됨
                sItem = "1"
            Else
                sItem = "0"
            End If
        Else
            sInOut = Trim(Mid(TmpStr, 1, 1))                    ' 입출고구분 첫번째자리
            sDate = CStr(Year(Date)) & Trim(Mid(TmpStr, 6, 4))  ' dat파일에서 6 자리일자를 8자리로 변환
            mDate = Format(sDate, "####/##/##")
            sCode = Trim(Mid(TmpStr, 10, 3))                    ' dat파일에서 read
            sTagNo = Trim(Mid(TmpStr, 13, 4))                   ' ''
            sItem = Trim(Mid(TmpStr, 18, 1))                    ' 소품구분
        
        End If
        
        '비정상 자료 check
        If sInOut < "2" Or sInOut > "3" Or _
           Len(Trim(sDate)) <> 8 Or Not IsDate(mDate) Or _
           Len(Trim(sCode)) <> 3 Or Not IsNumeric(sCode) Or _
           Len(Trim(sTagNo)) <> 4 Or Not IsNumeric(sTagNo) Or _
           sItem < "0" Or sItem > "1" Then
           
              spdView(0).MaxRows = spdView(0).MaxRows + 1
              spdView(0).Row = spdView(0).MaxRows
              spdView(0).Col = -1
              spdView(0).BackColor = vbYellow
              spdView(0).Col = 10
              spdView(0).Text = Mid(TmpStr, 1, 20)
              
              GoTo handy_err
        End If
        
        ReDim sValue(2)
        
        sValue(0) = sCode
        sValue(1) = sTagNo
        sValue(2) = sDate
        
        ' 출고일자를 기준으로 30일전의 데이터를 읽어온다.
        Set RS01 = New ADODB.Recordset
        Set RS01 = ExecPro("SP_08001_04", sValue(), Err_Num, Err_Dec)
        
        sFlag = "*"
        
        ' 출고일자를 기준으로 하여 입고일자가 30일 이전인 경우
        If Not RS01.BOF And Not RS01.EOF Then
        
            Do While Not RS01.EOF
                ' 30일이전에 중복TAG이 있는 경우 노란색으로 표시
                If RS01.RecordCount > 1 Then
                    spdView(1).MaxRows = spdView(1).MaxRows + 1
                    spdView(1).Row = spdView(1).MaxRows
                
                    If sFlag = "*" Then
                        spdView(1).Col = 11:    spdView(1).Value = True:    sFlag = ""
                    End If
                    
                    spdView(1).Col = 1:  spdView(1).Text = "[" & RS01!대리점코드 & "] " & RS01!대리점명
                    spdView(1).Col = 2:  spdView(1).Text = RS01!입고일자
                    spdView(1).Col = 3:  spdView(1).Text = RS01!택번호
                    spdView(1).Col = 4:  spdView(1).Text = "[" & RS01!품목코드 & "] " & RS01!품명
                    spdView(1).Col = 5:  spdView(1).Text = RS01!출고일자
                    spdView(1).Col = 6:  spdView(1).Text = IIf(sInOut = "2", "[2] 정상", "[3] 반품")
                    spdView(1).Col = 7:  spdView(1).Text = IIf(sItem = "1", "[1] 소품", "[0] 정상")
                    spdView(1).Col = 8:  spdView(1).Text = IIf(sInOut = "2", "정상", "반품")
                    spdView(1).Col = 9:  spdView(1).Text = IIf(sItem = "1", "소품", "")
                    spdView(1).Col = 10: spdView(1).Text = Mid(TmpStr, 1, 20)
                Else
                
                    If Not (IsNull(RS01!출고일자) Or RS01!출고일자 = "    -  -  ") Then
                        spdView(2).MaxRows = spdView(2).MaxRows + 1
                        spdView(2).Row = spdView(2).MaxRows
                    
                        spdView(2).Col = 1:  spdView(2).Text = "[" & RS01!대리점코드 & "] " & RS01!대리점명
                        spdView(2).Col = 2:  spdView(2).Text = RS01!입고일자
                        spdView(2).Col = 3:  spdView(2).Text = RS01!택번호
                        spdView(2).Col = 4:  spdView(2).Text = "[" & RS01!품목코드 & "] " & RS01!품명
                        spdView(2).Col = 5:  spdView(2).Text = RS01!출고일자
                        spdView(2).Col = 6:  spdView(2).Text = IIf(sInOut = "2", "[2] 정상", "[3] 반품")
                        spdView(2).Col = 7:  spdView(2).Text = IIf(sItem = "1", "[1] 소품", "[0] 정상")
                        spdView(2).Col = 8:  spdView(2).Text = IIf(sInOut = "2", "정상", "반품")
                        spdView(2).Col = 9:  spdView(2).Text = IIf(sItem = "1", "소품", "")
                        spdView(2).Col = 10: spdView(2).Text = Mid(TmpStr, 1, 20)
                        
                    Else
                        spdView(0).MaxRows = spdView(0).MaxRows + 1
                        spdView(0).Row = spdView(0).MaxRows
                        
                        spdView(0).Col = 1:  spdView(0).Text = "[" & RS01!대리점코드 & "] " & RS01!대리점명
                        spdView(0).Col = 2:  spdView(0).Text = RS01!입고일자
                        spdView(0).Col = 3:  spdView(0).Text = RS01!택번호
                        spdView(0).Col = 4:  spdView(0).Text = "[" & RS01!품목코드 & "] " & RS01!품명
                        spdView(0).Col = 5:  spdView(0).Text = RS01!출고일자
                        spdView(0).Col = 6:  spdView(0).Text = IIf(sInOut = "2", "[2] 정상", "[3] 반품")
                        spdView(0).Col = 7:  spdView(0).Text = IIf(sItem = "1", "[1] 소품", "[0] 정상")
                        spdView(0).Col = 8:  spdView(0).Text = IIf(sInOut = "2", "정상", "반품")
                        spdView(0).Col = 9:  spdView(0).Text = IIf(sItem = "1", "소품", "")
                        spdView(0).Col = 10: spdView(0).Text = Mid(TmpStr, 1, 20)
                    End If
                End If
                
                RS01.MoveNext
            Loop
        
            GoTo handy_err
        Else
            Set RS01 = New ADODB.Recordset
            Set RS01 = ExecPro("SP_08001_10", sValue(), Err_Num, Err_Dec)
            
            If RS01.EOF Then
            ' 입고 자료가 없이 출고 시킬경우
                spdView(0).MaxRows = spdView(0).MaxRows + 1
                spdView(0).Row = spdView(0).MaxRows
                
                spdView(0).Col = -1: spdView(0).BackColor = vbGreen
                
                spdView(0).Col = 1:  spdView(0).Text = "[" & sCode & "] "
                spdView(0).Col = 3:  spdView(0).Text = Mid(sTagNo, 1, 1) & "-" & Mid(sTagNo, 2, 3)
            
                spdView(0).Col = 5:  spdView(0).Text = Format(sDate, "0000-00-00")
                spdView(0).Col = 6:  spdView(0).Text = IIf(sInOut = "2", "[2] 정상", "[3] 반품")
                spdView(0).Col = 7:  spdView(0).Text = IIf(sItem = "1", "[1] 소품", "[0] 정상")
                spdView(0).Col = 8:  spdView(0).Text = IIf(sInOut = "2", "정상", "반품")
                spdView(0).Col = 9:  spdView(0).Text = IIf(sItem = "1", "소품", "")
                spdView(0).Col = 10: spdView(0).Text = Mid(TmpStr, 1, 20)
            Else
                If Not IsNull(RS01!출고일자) Then
                    ' 출고 자료가 있는경우 (이전 출고가 있다)
                    spdView(2).MaxRows = spdView(2).MaxRows + 1
                    spdView(2).Row = spdView(2).MaxRows
                    
                    spdView(2).Col = -1:  spdView(2).BackColor = vbGreen
                    spdView(2).Col = 1:  spdView(2).Text = "[" & RS01!대리점코드 & "] " & RS01!대리점명
                    spdView(2).Col = 2:  spdView(2).Text = RS01!입고일자
                    spdView(2).Col = 3:  spdView(2).Text = RS01!택번호
                    spdView(2).Col = 4:  spdView(2).Text = "[" & RS01!품목코드 & "] " & RS01!품명
                    spdView(2).Col = 5:  spdView(2).Text = RS01!출고일자
                    spdView(2).Col = 6:  spdView(2).Text = IIf(sInOut = "2", "[2] 정상", "[3] 반품")
                    spdView(2).Col = 7:  spdView(2).Text = IIf(sItem = "1", "[1] 소품", "[0] 정상")
                    spdView(2).Col = 8:  spdView(2).Text = IIf(sInOut = "2", "정상", "반품")
                    spdView(2).Col = 9:  spdView(2).Text = IIf(sItem = "1", "소품", "")
                    spdView(2).Col = 10: spdView(2).Text = Mid(TmpStr, 1, 20)
                Else
                    
                    spdView(0).MaxRows = spdView(0).MaxRows + 1
                    spdView(0).Row = spdView(0).MaxRows
                    
                    spdView(0).Col = -1: spdView(0).BackColor = vbGreen ' 그린색으로 표시
                    
                    spdView(0).Col = 1:  spdView(0).Text = "[" & RS01!대리점코드 & "] " & RS01!대리점명
                    spdView(0).Col = 2:  spdView(0).Text = RS01!입고일자
                    spdView(0).Col = 3:  spdView(0).Text = RS01!택번호
                    spdView(0).Col = 4:  spdView(0).Text = "[" & RS01!품목코드 & "] " & RS01!품명
                    spdView(0).Col = 5:  spdView(0).Text = RS01!출고일자
                    spdView(0).Col = 6:  spdView(0).Text = IIf(sInOut = "2", "[2] 정상", "[3] 반품")
                    spdView(0).Col = 7:  spdView(0).Text = IIf(sItem = "1", "[1] 소품", "[0] 정상")
                    spdView(0).Col = 8:  spdView(0).Text = IIf(sInOut = "2", "정상", "반품")
                    spdView(0).Col = 9:  spdView(0).Text = IIf(sItem = "1", "소품", "")
                    spdView(0).Col = 10: spdView(0).Text = Mid(TmpStr, 1, 20)
                End If
            
            End If
            
            GoTo handy_err
        End If
        

        
        DoEvents
        
handy_err:
    Loop
    Close #1
    
'    If spdView(0).MaxRows <> 0 Or spdView(1).MaxRows <> 0 Or spdView(2).MaxRows <> 0 Then
'        cmdBtn(2).Enabled = True
'        cmdBtn(3).Enabled = True
'    End If
    
    MsgBox "출고 스캔 " & CStr(iCnt) & "건이 전송 되었습니다."
    
'    If ErrorCnt > 0 Then
'       MsgBox "출고 자료로 " & CStr(iCnt) & "건 중 오류 " CStr(ErrorCnt) & "건이 출고내역이 없습니다."
'    Else
'       MsgBox CStr(iCnt) & "건이 입고내역이 없습니다."
'    End If

End Sub

Private Sub DataSave4()
    Dim FileName As String
    Dim AgencyCode As String
    Dim TagNo As String
    Dim ChulGu As String
    Dim RecCount As Long
    
    ReDim sValue(0)

    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_08001_07", sValue(), Err_Num, Err_Dec)
    
    If RS01.BOF Or RS01.EOF Then
        MsgBox "디스켓에 복사할 자료가 없습니다."
        DoEvents
        Exit Sub
    End If
        
    FileName = "Down" & RS01!대리점코드 & ".Dat"
        
    'a:드라이버에 디스켓이 있는지 확인
    Call PanelsMsg("A:드라이버 확인중")
    
    While Not DriverChk
        If MsgBox("A:드라이버에 디스켓을 넣으십시요", vbRetryCancel) = vbCancel Then
            BtnEnable
            Exit Sub
        End If
    Wend
    
    AgencyCode = RS01!대리점코드
    DoEvents
    Open "A:\" & FileName For Output As #1
    
    While Not RS01.EOF And Not RS01.BOF
        '대리점이 바뀌었으면 다른이름으로 다시 저장한다.
        If AgencyCode <> RS01!대리점코드 Then
            Close #1
            FileName = "Down" & RS01!대리점코드 & ".Dat"
            Open "A:\" & FileName For Output As #1
        End If
        
        AgencyCode = RS01!대리점코드
        TagNo = RS01!택번호
        
        ChulGu = IIf(IsNull(RS01!출고구분), 0, RS01!출고구분)
        
        Print #1, " " & RS01!입고일자 & " " & TagNo & " " & ChulGu
        RS01.MoveNext
        DoEvents
    Wend
    
    Close #1
End Sub

Private Sub DataSave5()
    ReDim sValue(0)
    
    Call ExecPro("SP_08001_08", sValue(), Err_Num, Err_Dec)
End Sub

Private Sub DataSave6()
    '하드디스크작성
    Dim FileName As String
    Dim rName As String
    Dim rAmt As String
    
'    panCaption(0).Visible = True
    dtInput.Visible = True
    dtInput.Value = Date
    dtInput.Value = ""
    
    DoEvents
    
    If dtInput.Value = "" Then
       Exit Sub
    End If
    
    ReDim sValue(0)
    
    sValue(0) = "0"
    
    Set RS01 = New ADODB.Recordset
    Set RS01 = ExecPro("SP_08001_05", sValue(), Err_Num, Err_Dec)
    
    If RS01.RecordCount = 0 Then
        MsgBox "디스켓에 복사할 자료가 없습니다."
        RS01.Close
        Exit Sub
    End If
        
    FileName = "R" & Format(dtInput.Value, "yyyymmdd") & ".Dat"
        
    'a:드라이버에 디스켓이 있는지 확인
    Call PanelsMsg("A:드라이버 확인중")
    
    While Not DriverChk
        If MsgBox("A:드라이버에 디스켓을 넣으십시요", vbRetryCancel) = vbCancel Then
            RS01.Close
            Exit Sub
        End If
    Wend
    
    Call PanelsMsg("수선자료를 복사중입니다.")
    
    Open "A:\" & FileName For Output As #1
    
    While Not RS01.EOF
        rAmt = "       "
        LSet rAmt = RS01!금액
        
        Print #1, rAmt;
        Print #1, RS01!명칭
        
        RS01.MoveNext
        
        DoEvents
    Wend
    
    Close #1
    
    RS01.Close
End Sub

Private Sub cmdSubButton_Click(Index As Integer)
    Select Case Index
        Case 0
            spdView(0).Row = spdView(0).ActiveRow
            spdView(0).Col = -1
            spdView(0).Action = ActionDeleteRow
            
            spdView(0).MaxRows = spdView(0).MaxRows - 1
        Case 1
            spdView(1).Row = spdView(1).ActiveRow
            spdView(1).Col = -1
            spdView(1).Action = ActionDeleteRow
            
            spdView(1).MaxRows = spdView(1).MaxRows - 1
        Case 2
            spdView(2).Row = spdView(2).ActiveRow
            spdView(2).Col = -1
            spdView(2).Action = ActionDeleteRow
            
            spdView(2).MaxRows = spdView(2).MaxRows - 1
        Case 3
            Dim i As Integer
            Dim ii As Integer
            
            Do While i < spdView(1).MaxRows
                i = i + 1
                spdView(1).Row = i
                spdView(1).Col = 11
                If spdView(1).Value <> True Then
                    spdView(1).Action = ActionDeleteRow
                    i = i - 1
                    spdView(1).MaxRows = spdView(1).MaxRows - 1
                Else
                    ii = ii + 1
                End If
                
            Loop
    End Select
End Sub

Private Sub Form_Activate()
'    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
'    cmdBtn(2).Enabled = False
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    cboHT_Gubun.ListIndex = Val(GetIniStr("Order Setting", "P_08001_SCAN_MODE", "", m_iniFile))
        

    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call INIWrite("Order Setting", "P_08001_SCAN_MODE", CStr(cboHT_Gubun.ListIndex), m_iniFile)
    
    P_08001_Flag = False
End Sub



Public Sub DataSave()
    Dim i As Integer
    Dim sMemAgencyCode As String
    Dim sDownPath As String
    Dim sSendData As String
    
    On Error GoTo ERR_RTN
    
    ReDim sValue(5)
    
    For i = 1 To spdView(1).MaxRows
        spdView(1).Row = i
        spdView(1).Col = 11
        If spdView(1).Value <> True Then
            MsgBox "선택출고 자료를 확인 후에 저장하세요.. ", vbInformation
            Exit Sub
        End If
        
    Next i

    
'    spdView(0).Col = 1:  spdView(0).Text = "[" & RS01!대리점코드 & "] " & RS01!대리점명
'    spdView(0).Col = 2:  spdView(0).Text = RS01!입고일자
'    spdView(0).Col = 3:  spdView(0).Text = RS01!택번호
'    spdView(0).Col = 4:  spdView(0).Text = "[" & RS01!품목코드 & "] " & RS01!품명
'    spdView(0).Col = 5:  spdView(0).Text = RS01!출고일자
'    spdView(0).Col = 6:  spdView(0).Text = IIf(sInOut = "2", "[2] 정상", "[3] 반품")
'    spdView(0).Col = 7:  spdView(0).Text = IIf(sItem = "1", "[1] 소품", "[0] 정상")
'
    ' 1: 대리점 코드
    ' 3: 택번호
    ' 2: 입고일자
    ' 5: 출고일자
    ' 6: 2/3    "[2] 정상", "[3] 반품"
    ' 7: 1/0    "[1] 소품", "[0] 정상"
    
    For i = 1 To spdView(0).MaxRows
        spdView(0).Row = i
        
        spdView(0).Col = 1: sValue(0) = Mid(spdView(0).Text, 2, 3)
        spdView(0).Col = 3: sValue(1) = spdView(0).Value
        spdView(0).Col = 2: sValue(2) = Format(spdView(0).Text, "YYYY-MM-DD")
        spdView(0).Col = 5: sValue(3) = Format(Now, "YYYY-MM-DD")
        spdView(0).Col = 6: sValue(4) = Mid(spdView(0).Text, 2, 1)
        spdView(0).Col = 7: sValue(5) = Mid(spdView(0).Text, 2, 1)
        
        Call ExecPro("SP_08001_06", sValue(), Err_Num, Err_Dec)
        
        If Err_Num <> 0 Then
            MsgBox "[" & Err_Num & "] " & Err_Dec
            Exit Sub
        End If
    Next i
    
    For i = 1 To spdView(1).MaxRows
        spdView(1).Row = i
        
        spdView(1).Col = 1: sValue(0) = Mid(spdView(1).Text, 2, 3)
        spdView(1).Col = 3: sValue(1) = spdView(1).Value
        spdView(1).Col = 2: sValue(2) = Format(spdView(1).Text, "YYYY-MM-DD")
        spdView(1).Col = 5: sValue(3) = Format(Now, "YYYY-MM-DD")
        spdView(1).Col = 6: sValue(4) = Mid(spdView(1).Text, 2, 1)
        spdView(1).Col = 7: sValue(5) = Mid(spdView(1).Text, 2, 1)
        
        Call ExecPro("SP_08001_06", sValue(), Err_Num, Err_Dec)
        
        If Err_Num <> 0 Then
            MsgBox "[" & Err_Num & "] " & Err_Dec
            Exit Sub
        End If
    Next i
    
    For i = 1 To spdView(2).MaxRows
        spdView(2).Row = i
        
        spdView(2).Col = 1: sValue(0) = Mid(spdView(2).Text, 2, 3)
        spdView(2).Col = 3: sValue(1) = spdView(2).Value
        spdView(2).Col = 2: sValue(2) = ""
        spdView(2).Col = 5: sValue(3) = Format(Now, "YYYY-MM-DD")
        spdView(2).Col = 6: sValue(4) = Mid(spdView(2).Text, 2, 1)
        spdView(2).Col = 7: sValue(5) = Mid(spdView(2).Text, 2, 1)
        
        Call ExecPro("SP_08001_06", sValue(), Err_Num, Err_Dec)
        
        If Err_Num <> 0 Then
            MsgBox "[" & Err_Num & "] " & Err_Dec
            Exit Sub
        End If
    Next i
    
    sDownPath = GetIniStr("SERVER DATA", "SendPath", "", m_iniFile)
    
    For i = 1 To spdView(0).MaxRows
        spdView(0).Row = i
        spdView(0).Col = 1
            
        If sMemAgencyCode <> Mid(spdView(0).Text, 2, 3) Then
            ReDim sValue(2)
            
            sValue(0) = "0"
            sValue(1) = Format(Now, "YYYY-MM-DD")
            sValue(2) = sMemAgencyCode
            
            Open sDownPath & "\Down" & sValue(2) & sValue(1) & ".Dat" For Output As #1
            
            Set RS01 = New ADODB.Recordset
            Set RS01 = ExecPro("SP_08003_01", sValue(), Err_Num, Err_Dec)
            
            Do While Not RS01.EOF
                Print #1, " " & RS01!입고일자 & " " & RS01!택번호 & " " & RS01!출고구분
            
                RS01.MoveNext
            Loop
        
            Close #1
        End If
        
        sMemAgencyCode = Mid(spdView(0).Text, 2, 3)
    Next i
    
    For i = 1 To spdView(1).MaxRows
        spdView(1).Row = i
        spdView(1).Col = 1
            
        If sMemAgencyCode <> Mid(spdView(1).Text, 2, 3) Then
            ReDim sValue(2)
            
            sValue(0) = "0"
            sValue(1) = Format(Now, "YYYY-MM-DD")
            sValue(2) = sMemAgencyCode
            
            Open sDownPath & "\Down" & sValue(2) & sValue(1) & ".Dat" For Output As #1
            
            Set RS01 = New ADODB.Recordset
            Set RS01 = ExecPro("SP_08003_01", sValue(), Err_Num, Err_Dec)
            
            Do While Not RS01.EOF
                Print #1, " " & RS01!입고일자 & " " & RS01!택번호 & " " & RS01!출고구분
            
                RS01.MoveNext
            Loop
        
            Close #1
        End If
        
        sMemAgencyCode = Mid(spdView(1).Text, 2, 3)
    Next i
    
    For i = 1 To spdView(2).MaxRows
        spdView(2).Row = i
        spdView(2).Col = 1
            
        If sMemAgencyCode <> Mid(spdView(2).Text, 2, 3) Then
            ReDim sValue(2)
            
            sValue(0) = "0"
            sValue(1) = Format(Now, "YYYY-MM-DD")
            sValue(2) = sMemAgencyCode
            
            Open sDownPath & "\Down" & sValue(2) & sValue(1) & ".Dat" For Output As #1
            
            Set RS01 = New ADODB.Recordset
            Set RS01 = ExecPro("SP_08003_01", sValue(), Err_Num, Err_Dec)
            
            Do While Not RS01.EOF
                Print #1, " " & RS01!입고일자 & " " & RS01!택번호 & " " & RS01!출고구분
            
                RS01.MoveNext
            Loop
        
            Close #1
        End If
        
        sMemAgencyCode = Mid(spdView(2).Text, 2, 3)
    Next i
    
    MsgBox "해당되는 데이터가 정상적으로 저장이 되었습니다", vbInformation
    Exit Sub
ERR_RTN:
    MsgBox Err.Description
    Resume Next
End Sub

Public Sub DataDelete()

End Sub
