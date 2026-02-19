VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form frm가격변경 
   Caption         =   "가격 변경"
   ClientHeight    =   7035
   ClientLeft      =   3360
   ClientTop       =   5355
   ClientWidth     =   11850
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frm가격변경.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form17"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   11850
   WindowState     =   2  '최대화
   Begin MSMask.MaskEdBox mskY 
      Height          =   375
      Left            =   3555
      TabIndex        =   2
      Top             =   2040
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox mskM 
      Height          =   375
      Left            =   4755
      TabIndex        =   3
      Top             =   2040
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox mskD 
      Height          =   375
      Left            =   5655
      TabIndex        =   4
      Top             =   2055
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##"
      PromptChar      =   " "
   End
   Begin VB.Frame Frame1 
      Height          =   6840
      Left            =   90
      TabIndex        =   6
      Top             =   45
      Width           =   11685
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   7650
         Top             =   2130
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin ComctlLib.ProgressBar PgBar 
         Height          =   465
         Left            =   1350
         TabIndex        =   13
         Top             =   4380
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   820
         _Version        =   327682
         Appearance      =   1
         Max             =   250
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   420
         Index           =   1
         Left            =   1245
         TabIndex        =   7
         Top             =   780
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   741
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "자료구분"
         BevelWidth      =   2
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1050
         Left            =   3435
         TabIndex        =   8
         Top             =   795
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   1852
         _Version        =   262144
         BackColor       =   12648447
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   6.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Begin VB.OptionButton Opt_Sale 
            BackColor       =   &H00C0FFFF&
            Caption         =   "보관서비스"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   150
            TabIndex        =   17
            Top             =   690
            Width           =   2040
         End
         Begin VB.OptionButton Opt_Sale 
            BackColor       =   &H00C0FFFF&
            Caption         =   "할인자료"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   2820
            TabIndex        =   16
            Top             =   375
            Width           =   1680
         End
         Begin VB.OptionButton Opt_Sale 
            BackColor       =   &H00C0FFFF&
            Caption         =   "수선자료"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   150
            TabIndex        =   15
            Top             =   375
            Width           =   1680
         End
         Begin VB.OptionButton Opt_Sale 
            BackColor       =   &H00C0FFFF&
            Caption         =   "목요세일"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   2805
            TabIndex        =   1
            Top             =   45
            Width           =   1680
         End
         Begin VB.OptionButton Opt_Sale 
            BackColor       =   &H00C0FFFF&
            Caption         =   "품목가격"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   150
            TabIndex        =   0
            Top             =   45
            Value           =   -1  'True
            Width           =   1860
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   420
         Index           =   0
         Left            =   1245
         TabIndex        =   9
         Top             =   1950
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   741
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "적용일자"
         BevelWidth      =   2
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand Command1 
         Height          =   1155
         Left            =   8985
         TabIndex        =   5
         Top             =   780
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   2037
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "확 인"
         ButtonStyle     =   2
         BevelWidth      =   3
      End
      Begin VB.Label LblMsg 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1335
         TabIndex        =   14
         Top             =   2760
         Width           =   9030
      End
      Begin VB.Label Label1 
         Caption         =   "일"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   6090
         TabIndex        =   12
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "월"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   5175
         TabIndex        =   11
         Top             =   2025
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "년"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   4215
         TabIndex        =   10
         Top             =   2025
         Width           =   375
      End
   End
End
Attribute VB_Name = "frm가격변경"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strText     As String

Dim strPath01   As String
Dim strPCode    As String
Dim strName     As String
Dim strPrice    As String
Dim sDate       As String

Private Sub Command1_Click()
    Dim strFilename As String
    Dim Filename As String
    
     On Error GoTo ErrHandler
    
    sDate = Trim(mskY.ClipText) & "-" & Trim(mskM.ClipText) & "-" & Trim(mskD.ClipText)
    LblMsg.Caption = ""
    PgBar.Value = 0
    
    If IsDate(sDate) = False Then
        MsgBox "일자를 잘못입력하셨읍니다", vbCritical
        Exit Sub
    End If
    If Format(sDate, "YYYY-MM-DD") <> Format(Date, "YYYY-MM-DD") Then
        MsgBox "당일 날짜만 변경 가능 합니다.", vbCritical
        Exit Sub
    End If
    
    sDate = Mid(sDate, 1, 4) & Mid(sDate, 6, 2) & Mid(sDate, 9, 2)
    
    If Opt_Sale(0).Value = True Then
        Call dataPrice
    ElseIf Opt_Sale(1).Value = True Then
        Call DaySalePrice
    ElseIf Opt_Sale(2).Value = True Then
        Call RepairPrice
    ElseIf Opt_Sale(3).Value = True Then
        With CommonDialog1
            .CancelError = True
            .Flags = cdlOFNHideReadOnly
            .Filter = "All Files (*.*)|*.*|할인자료(*.dat)|*.dat|"
            .FilterIndex = 2
            .InitDir = App.Path
            .ShowOpen
            
            DoEvents
            If .Filename = "" Then Exit Sub
            strFilename = .Filename
            
            If InStr(UCase(strFilename), "SALE" & 가맹점정보.택코드 & ".DAT") <= 0 Then
                MsgBox "선택한 파일이 해당 대리점에서 사용할 수 없는 파일 입니다." & Space(10), vbInformation, "확인"
                Exit Sub
            End If
            
            ' 작업할 파일이 다른폴더에 있을경우 해당 폴더로 복사한다.
            If UCase(Left(strFilename, 20)) <> UCase(App.Path & "\BackData\") Then
                If Dir(App.Path & "\BackData", vbDirectory) = "" Then
                    MkDir App.Path & "\BackData"
                End If
                
                Filename = Mid(strFilename, InStrRev(UCase(strFilename), "\") + 1, 12)
                FileCopy strFilename, App.Path & "\BackData\" & Filename
                DoEvents
                Delay (2)
            End If
        End With
        Call SaleData
    
    ElseIf Opt_Sale(4).Value = True Then
        Call QN_Price
    Else
    
    End If
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    'TitleSet "금액 변경"
    mskY.SelText = Format(Date, "yyyy")
    mskM.SelText = Format(Date, "mm")
    mskD.SelText = Format(Date, "dd")
End Sub

Private Sub mskD_Change()
    mskD.SelStart = 0
    mskD.SelLength = 2
End Sub

Private Sub mskM_Change()
    mskM.SelStart = 0
    mskM.SelLength = 2
End Sub

Private Sub mskY_GotFocus()
    mskY.SelStart = 0
    mskY.SelLength = 4
End Sub

'**************************************************************************************
' 보관금액 테이블에 update
'**************************************************************************************
Private Sub QN_Price()
    Dim strCode As String
    
    Dim varData   As Variant
    
    Query = "SELECT 택코드 FROM TB_기본정보 "
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF Then
        LblMsg = "대리점코드가 존재하지 않읍니다..!"
        ADORs.Close
        LblMsg = ""
        Exit Sub
    End If
    
    strCode = Format(Trim(ADORs!택코드), "000")
    
    ADORs.Close
    
    strPath01 = App.Path & "\BackData\" & sDate & strCode & ".dat"
    If Dir(strPath01) = "" Then
        MsgBox Mid(sDate, 1, 4) & "년" & Mid(sDate, 5, 2) & "월" & Mid(sDate, 7, 2) & "일 " _
        & " 자료가 없습니다 ", vbCritical
        Exit Sub
    End If
    
    PgBar.Value = 0
    
    On Error GoTo diskError01
    
    Open strPath01 For Input As #1 ' Open file.
    Line Input #1, strText  ' Read line into variable.
    
    If InStr(strText, "보관서비스금액") <= 0 Then
        LblMsg = "보관 서비스 금액 자료가 없습니다."
        Close #1
        Exit Sub
    End If
    
    ADOCon.Execute "DELETE  FROM TB_보관금액"
    
    PgBar.MAX = 1452
    
    Do While Not EOF(1) ' Loop until end of file.
        LblMsg = "보관 서비스 금액자료를 변경하고 있읍니다..!"
        
        Line Input #1, strText  ' Read line into variable.
        
        varData = Split(strText, "|")
        
        If UBound(varData) = 3 Then
            
            Query = "INSERT INTO TB_보관금액 (보관월, 아이템수, 보관개월수, 보관금액) "
            Query = Query & "VALUES ('" & Trim(CStr(varData(0))) & "', "
            Query = Query & "" & Val(Trim(CStr(varData(1)))) & ", "
            Query = Query & "" & Val(Trim(CStr(varData(2)))) & ", "
            Query = Query & "" & Val(Trim(CStr(varData(3)))) & ")"
            ADOCon.Execute Query
        End If
        
        If PgBar.Value = 1452 Then
            PgBar.Value = 0
        End If
        
        PgBar.Value = PgBar.Value + 1
        DoEvents
    Loop
    PgBar.Value = PgBar.MAX
    
    Close #1    ' Close file
    
    'Kill strPath01   '98/06/23   금액자료 삭제취소  _한과장님
    LblMsg = "보관 서비스 금액자료 변경 완료..!"
    Exit Sub
    
diskError01:
    MsgBox " 보관 서비스  금액자료 변경 에러 " & Str(VBA.Err.Number) & "  " & VBA.Err.Description, vbCritical, VBA.Err.Source
End Sub

'**************************************************************************************
'참조코드 테이블에 update
'**************************************************************************************
Private Sub dataPrice()
    Dim strCode As String
    
    Query = "SELECT 택코드 FROM TB_기본정보 "
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF Then
        LblMsg = "대리점코드가 존재하지 않읍니다..!"
        
        ADORs.Close
        Set ADORs = Nothing
        
        LblMsg = ""
        
        Exit Sub
    End If
    
    strCode = Trim(ADORs!택코드)
    
    If Len(strCode) = 1 Then
        strCode = "00" & strCode
    ElseIf Len(strCode) = 2 Then
        strCode = "0" & strCode
    End If
    
    ADORs.Close
    
    strPath01 = App.Path & "\BackData\" & sDate & strCode & ".dat"
    
    If Dir(strPath01) = "" Then
        MsgBox Mid(sDate, 1, 4) & "년" & Mid(sDate, 5, 2) & "월" & Mid(sDate, 7, 2) & "일 " _
        & " 자료가 없습니다 ", vbCritical
        Exit Sub
    End If
    
    PgBar.Value = 0
    
    On Error GoTo diskError01
    
    Open strPath01 For Input As #1 ' Open file.
    Line Input #1, strText  ' Read line into variable.
    
    If InStr(strText, "보관서비스금액") > 0 Then
        LblMsg = "대리점코드가 존재하지 않읍니다..!"
        Close #1
        Exit Sub
    End If
    
    
    Query = "DELETE FROM TB_의류"
    ADOCon.Execute Query
    
    ' pds2004 수정 2007-04-11
    ' 이전에 open 하여 읽어서 파일의 아상유무를 확인하였기 때문에
    ' 파일을 Close하고 다시 오픈한다.
    
    Close #1
    Open strPath01 For Input As #1 ' Open file.
    
    Do While Not EOF(1) ' Loop until end of file.
        LblMsg = "금액자료를 변경하고 있읍니다..!"
        
        Line Input #1, strText  ' Read line into variable.
        
        strPCode = Trim(Mid(strText, 2, 3))
        strPrice = Trim(Mid(strText, 6, 8))
        strName = Trim(Mid(strText, 15, 20))
        
        Query = "INSERT INTO TB_의류(의류코드, 금액, 의류명) "
        Query = Query & "VALUES ('" & Trim(strPCode) & "', "
        Query = Query & "'" & Trim(strPrice) & "', "
        Query = Query & "'" & Trim(strName) & "')"
        ADOCon.Execute Query
        
        If PgBar.Value = 250 Then
            PgBar.Value = 0
        End If
        
        PgBar.Value = PgBar.Value + 1
        DoEvents
    Loop
    PgBar.Value = PgBar.MAX
    
    Close #1    ' Close file
    
    'Kill strPath01   '98/06/23   금액자료 삭제취소  _한과장님
    LblMsg = "금액자료 변경 완료..!"
    Exit Sub
    
diskError01:
    MsgBox " 금액자료 변경 에러 " & Str(VBA.Err.Number) & "  " & VBA.Err.Description, vbCritical, VBA.Err.Source
End Sub

'**************************************************************************************
'목요세일 테이블에 insert
'**************************************************************************************
Private Sub DaySalePrice()
    Dim strCode As String
    
    Query = "SELECT 택코드 FROM TB_기본정보 "
    Set ADORs = New ADODB.Recordset
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.RecordCount < 1 Then
        LblMsg = "대리점코드가 존재하지 않읍니다..!"
        ADORs.Close
        LblMsg = ""
        Exit Sub
    End If
    
    strCode = Trim(ADORs!택코드)
    
    If Len(strCode) = 1 Then
        strCode = "00" & strCode
    ElseIf Len(strCode) = 2 Then
        strCode = "0" & strCode
    End If
    
    ADORs.Close
    
    strPath01 = App.Path & "\BackData\D" & sDate & strCode & ".dat"
    
    If Dir(strPath01) = "" Then
        MsgBox Mid(sDate, 1, 4) & "년" & Mid(sDate, 5, 2) & "월" & Mid(sDate, 7, 2) & "일 " _
        & " 자료가 없습니다 ", vbCritical
        Exit Sub
    End If
    
    PgBar.Value = 0
    
    On Error GoTo diskError02
    
    Query = "DELETE  "
    Query = Query & "FROM TB_목요세일"
    ADOCon.Execute Query
    
    Open strPath01 For Input As #1 ' Open file.
    
    Do While Not EOF(1) ' Loop until end of file.
        LblMsg = "목요세일자료를 변경하고 있읍니다..!"
        Line Input #1, strText  ' Read line into variable.
        
        strPCode = Trim(Mid(strText, 2, 3))
        strPrice = Trim(Mid(strText, 6, 8))
        strName = Trim(Mid(strText, 15, 20))
        
        Query = "INSERT INTO TB_목요세일(의류코드, 금액, 의류명) "
        Query = Query & "VALUES ('" & Trim(strPCode) & "', "
        Query = Query & "'" & Trim(strPrice) & "', "
        Query = Query & "'" & Trim(strName) & "')"
        ADOCon.Execute Query
        
        If PgBar.Value = 250 Then
            PgBar.Value = 0
        End If
        
        PgBar.Value = PgBar.Value + 1
        
        DoEvents
    Loop
    PgBar.Value = PgBar.MAX
    
    Close #1    ' Close file
    'Kill strPath01   '98/06/23   금액자료 삭제취소  _한과장님
    LblMsg = "목요세일자료 변경 완료..!"
    
    Exit Sub
    
diskError02:
    MsgBox " 목요세일자료 변경중 에러 " & Str(VBA.Err.Number) & "  " & VBA.Err.Description, vbCritical, VBA.Err.Source
End Sub

Private Sub RepairPrice()
'**************************************************************************************
'수선금액 테이블에 update
'**************************************************************************************
   
'    strCode = "SELECT 택코드 FROM TB_기본정보 "
'    Set rsCode = MyDB.OpenRecordset(strCode)
'    If rsCode.RecordCount < 1 Then
'       LblMsg = "대리점코드가 존재하지 않읍니다..!"
'       rsCode.Close
'       LblMsg = ""
'       Exit Sub
'    End If
'
'    strCode = Trim(rsCode!택코드)
'    If Len(strCode) = 1 Then
'       strCode = "00" & strCode
'    ElseIf Len(strCode) = 2 Then
'       strCode = "0" & strCode
'    End If
'    rsCode.Close

    strPath01 = App.Path & "\BackData\R" & sDate & ".dat"
    
    If Dir(strPath01) = "" Then
        MsgBox Mid(sDate, 1, 4) & "년" & Mid(sDate, 5, 2) & "월" & Mid(sDate, 7, 2) & "일 " _
        & " 자료가 없습니다 ", vbCritical
        Exit Sub
    End If
    
    PgBar.Value = 0
    
    On Error GoTo diskError01
    
    Query = "DELETE FROM TB_수선금액"
    ADOCon.Execute Query
    
    Open strPath01 For Input As #1 ' Open file.
    
    Do While Not EOF(1) ' Loop until end of file.
        LblMsg = "수선금액을 변경하고 있읍니다..!"
        Line Input #1, strText  ' Read line into variable.
        
        strPrice = Trim(Mid(strText, 1, 7))
        strName = Trim(Mid(strText, 8, 20))
        
        Query = "INSERT INTO TB_수선금액(수선내용, 금액) "
        Query = Query & "VALUES ('" & Trim(strName) & "', "
        Query = Query & "'" & Trim(strPrice) & "') "
        ADOCon.Execute Query
        
        If PgBar.Value = 250 Then
            PgBar.Value = 0
        End If
        PgBar.Value = PgBar.Value + 1
        DoEvents
    Loop
    PgBar.Value = PgBar.MAX
    Close #1    ' Close file
        
    'Kill strPath01   '98/06/23   금액자료 삭제취소  _한과장님
    LblMsg = "수선금액 변경 완료..!"
    Exit Sub
    
diskError01:
    MsgBox " 수선금액 변경 에러 " & Str(VBA.Err.Number) & "  " & VBA.Err.Description, vbCritical, VBA.Err.Source
End Sub

Private Function SaleData() As Boolean
    Dim g_SaleDate As String
    Dim g_Count As Integer
    Dim Filename As String
    Dim St As String
    
'    If Not IsDate(txtDate.Text) Then
'        MsgBox "일자가 잘못 입력되었습니다.", vbExclamation, "입력오류"
'        Exit Function
'    End If
'    g_SaleDate = Format(txtDate.Text, "YYYY-MM-DD")
    
    LblMsg.Caption = "할인자료 확인중..!"
    
    On Error GoTo Err_Loop
    
    strPath01 = App.Path & "\BackData\Sale" & 가맹점정보.택코드 & ".DAT"
    Filename = Dir(strPath01)
    
    If Filename = "" Then
        LblMsg.Caption = ""
        MsgBox "할인자료가 없습니다.", vbInformation, "확인"
        Exit Function
    Else
        ' 모뎀일 경우에만 복사한다
        ' 인터넷일경우 다른곳에서 복사한다.
        FileCopy strPath01, App.Path & "\RecvData\" & Filename
        
        Open App.Path & "\RecvData\" & Filename For Input As #1
        Open App.Path & "\RecvData\Sale.Dat" For Output As #2
        
        Do While Not EOF(1)
            Line Input #1, St
            Print #2, St
            g_Count = g_Count + 1
        Loop
        
        Close
        
        PgBar.MAX = g_Count
        PgBar.Min = 0
        PgBar.Value = 0
    End If
    
    Workspaces(0).BeginTrans
    
    LblMsg.Caption = "기존 자료를 삭제중..!"
    ADOCon.Execute "DELETE FROM TB_할인정보"
    
    LblMsg.Caption = "할인자료 수신중..!" & Chr(13) & "총 " & g_Count & " 건"
    
    Open App.Path & "\RecvData\Sale.Dat" For Input As #1
    
    'Set daoQD = MyDB.CreateQueryDef("", "INSERT INTO TB_할인정보 VALUES (시작일, 종료일, 의류코드, 의류명, 금액, 비율, 출력순번)")
        
    Do While Not EOF(1)
        If PgBar.Value < PgBar.MAX Then
            PgBar.Value = PgBar.Value + 1
        End If
        
        Line Input #1, St
        
        Query = "INSERT INTO TB_할인정보 VALUES (시작일, 종료일, 의류코드, 의류명, 금액, 비율, 출력순번) VALUES "
        Query = Query & "  '" & Mid(St, 2, 8) & "'"
        Query = Query & ", '" & Mid(St, 11, 8) & "'"
        Query = Query & ", '" & Mid(St, 20, 3) & "'"
        Query = Query & ", '" & Mid(St, 36) & "'"
        Query = Query & ", '" & Mid(St, 24, 8) & "'"
        Query = Query & ", '" & Mid(St, 33, 1) & "'"
        Query = Query & ", '" & Mid(St, 33, 1) & "'"
        ADOCon.Execute Query
        
        'daoQD("시작일") = Mid(St, 2, 8)
        'daoQD("종료일") = Mid(St, 11, 8)
        'daoQD("의류코드") = Mid(St, 20, 3)
        'daoQD("의류명") = Mid(St, 36)
        'daoQD("금액") = Mid(St, 24, 8)
        'daoQD("비율") = Mid(St, 33, 1)
        'daoQD("출력순번") = Mid(St, 33, 1)
        'daoQD.Execute
    Loop
        
    LblMsg.Caption = "할인자료 수신이 완료되었습니다. " & g_Count & " 건"
    SaleData = True
    
    'daoQD.Close
    
    Workspaces(0).CommitTrans
    
End_Loop:
    Close
    
    On Error GoTo ERR_FILE
    If Dir(App.Path & "\RecvData\*.*") <> "" Then
        Kill App.Path & "\RecvData\*.*"
    End If
    Exit Function
    
Err_Loop:
    Workspaces(0).ROLLBACK
    LblMsg.Caption = "할인자료 수신이 취소되었습니다."
    
    MsgBox "작업오류 입니다." & Chr(13) & _
           "오류코드 : " & VBA.Err.Number & Chr(13) & _
           "오류설명 : " & VBA.Err.Description, vbExclamation, "오류"
           Resume Next
    Resume End_Loop

ERR_FILE:
    LblMsg.Caption = "파일삭제중 오류가 밸생했습니다."
    MsgBox "작업오류 입니다." & Chr(13) & _
           "오류코드 : " & VBA.Err.Number & Chr(13) & _
           "오류설명 : " & VBA.Err.Description, vbExclamation, "오류"
    Resume Next
End Function
