VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form frm자료수신 
   ClientHeight    =   8715
   ClientLeft      =   1965
   ClientTop       =   2190
   ClientWidth     =   11775
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form18"
   MDIChild        =   -1  'True
   ScaleHeight     =   8715
   ScaleWidth      =   11775
   WindowState     =   2  '최대화
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   615
      TabIndex        =   0
      Top             =   330
      Width           =   10575
      Begin ComCtl2.Animation Ani2 
         Height          =   600
         Left            =   3585
         TabIndex        =   2
         Top             =   4380
         Visible         =   0   'False
         Width           =   3990
         _ExtentX        =   7038
         _ExtentY        =   1058
         _Version        =   327681
         FullWidth       =   266
         FullHeight      =   40
      End
      Begin Threed.SSCommand DnData 
         Height          =   1170
         Left            =   3990
         TabIndex        =   1
         Top             =   1110
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   2064
         _Version        =   262144
         PictureFrames   =   1
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
         Picture         =   "frm자료수신.frx":0000
         Caption         =   "출고자료 수신"
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   6
         BevelWidth      =   3
      End
      Begin VB.Label Label4 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '단일 고정
         Caption         =   "A:드라이버에 디스켓을 넣으세요..!"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   1080
         TabIndex        =   3
         Top             =   3165
         Width           =   8400
      End
   End
End
Attribute VB_Name = "frm자료수신"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Option Explicit
''
''Private strPath1 As String
''
''Private Function DskCHK() As Boolean
''
''    On Error GoTo ErrRtn
''
''    Open "A:\CHK" For Random As #1
''
''    Close #1
''
''    Kill "A:\CHK"
''    DskCHK = True
''    Exit Function
''
''ErrRtn:
''    DskCHK = False
''End Function
''
''Private Function diskErrorHandller(errVal As Integer) As Integer
'''  Const ERR_DEVICEUNAVAILABLE = 68
'''  Const ERR_DISKNOTREADY = 71
'''  Const ERR_DEVICEIO = 57
'''  Const ERR_DISKFULL = 61
'''  Const ERR_BADFILENAME = 64
'''  Const ERR_BADFILENAMEORNUMBER = 52
'''  Const ERR_PATHDOSENOTEXIST = 76
'''  Const ERR_BADFILEMODE = 54
'''  Const ERR_FILEALREADYOPEN = 55
'''  Const ERR_INPUTPASTEENDOFFILE = 62
'''  Const MB_ICONEXCLAMATION = 48
'''  Dim msgtype As Integer, Msg As String, repose As Integer
'''  Dim Response As Long
'''
'''  Select Case errVal
'''      Case ERR_DEVICEUNAVAILABLE
'''         Msg = "장치 사용불가"
'''         msgtype = MB_ICONEXCLAMATION + 5
'''      Case ERR_DISKNOTREADY
'''         Msg = "디스크를 넣어 주십시요"
'''         msgtype = MB_ICONEXCLAMATION + 5
'''
'''      Case ERR_DEVICEIO
'''         Msg = "고장난 디스크 사용불가"
'''         msgtype = MB_ICONEXCLAMATION + 2
'''
'''      Case ERR_DISKFULL
'''         Msg = "여유공간부족, 계속진행하시겠읍니까 ?"
'''         msgtype = MB_ICONEXCLAMATION + 2
'''
'''      Case ERR_BADFILENAME
'''         Msg = "파일이름오류"
'''         msgtype = MB_ICONEXCLAMATION + 5
'''
'''      Case ERR_BADFILENAMEORNUMBER
'''         Msg = "파일이름오류"
'''         msgtype = MB_ICONEXCLAMATION + 5
'''
'''      Case ERR_PATHDOSENOTEXIST
'''         Msg = "경로가 일치하지않음"
'''         msgtype = MB_ICONEXCLAMATION + 5
'''
'''      Case ERR_BADFILEMODE
'''         Msg = "파일모드 불일치"
'''         msgtype = MB_ICONEXCLAMATION + 5
'''
'''      Case ERR_FILEALREADYOPEN
'''         Msg = "파일이 이미열려있음"
'''         msgtype = MB_ICONEXCLAMATION + 5
'''
'''      Case ERR_INPUTPASTEENDOFFILE
'''         Msg = "파일의 마지막에 추가"
'''         msgtype = MB_ICONEXCLAMATION + 5
'''         diskErrorHandller = 3
'''         Exit Function
'''  End Select
'''
'''  Response = MsgBox(Msg, msgtype, "Disk 에러")
'''
'''  Select Case Response
'''     Case 4
'''        diskErrorHandller = 0
'''     Case 5
'''        diskErrorHandller = 1
'''     Case 1, 2, 3
'''        diskErrorHandller = 3
'''  End Select
'''
'''  Select Case intEr1
'''         Case 52 '"잘못된 파일명이나 파일번호
'''            MsgBox "잘못된 파일명이나 파일번호"
'''         Case 53 '"파일이 없읍니다
'''            MsgBox "파일이 없읍니다"
'''         Case 54 '"잘못된 파일모드
'''            MsgBox "잘못된 파일모드"
'''         Case 55 '"파일이 이미열려있음
'''            MsgBox "파일이 이미열려있음"
'''         Case 57 '"기기 I/O 에러
'''            MsgBox "기기 I/O 에러"
'''         Case 58 '"파일이 이미존재함"
'''            MsgBox "파일이 이미존재함"
'''         Case 59 '"잘못된 레코드길이
'''            MsgBox " 잘못된 레코드길이"
'''         Case 61 '"디스크공간부족"
'''            MsgBox "디스크공간부족"
'''         Case 64 '"잘못된 파일명"
'''            MsgBox "잘못된 파일명"
'''         Case 68 '"디바이스 사용불가
'''            MsgBox "디바이스 사용불가"
'''         Case 71 '"디스크가 준비되지 않음"
'''            MsgBox "디스크가 준비되지 않음"
'''         Case 72 '"디스크불량"
'''            MsgBox "디스크불량"
'''         Case 76 '"경로를 찾을수 없음"
'''            MsgBox "경로를 찾을수 없음"
'''  End Select
''
''
''End Function
''
''Private Sub dataReceiv()
''    '**************************************************************************************
''    '출고테이블의 출고란에 표시"出" "反" 표시
''    '**************************************************************************************
''
''    Dim strText As String
''    Dim QueryCode As String
''    Dim rsCode As Recordset
''    Dim douCount As Double
''    Dim strPath   As String
''    Dim strTNo As String
''    Dim strClfy As String
''    Dim strRejec As String
''    Dim QueryreJec As String
''    Dim fileLength1 As Long
''    Dim strdate01 As String
''    Dim r_Value As Integer
''
''    strCode = "SELECT 대리점번호 FROM TB_대리점정보 "
''    Set rsCode = MyDB.OpenRecordset(strCode)
''
''    If rsCode.EOF = True Then
''        Label4 = "대리점코드가 존재하지 않읍니다..!"
''        rsCode.Close
''        Label4 = ""
''        Exit Sub
''    End If
''
''    strCode = Trim(rsCode!대리점번호)
''    If Len(strCode) = 1 Then
''        strCode = "00" & strCode
''    ElseIf Len(strCode) = 2 Then
''        strCode = "0" & strCode
''    End If
''    strPath = "A:\DOWN" & Trim(strCode) & ".dat"
''    rsCode.Close
''
''    Ani2.Visible = True
''    Ani2.AutoPlay = True
''    Ani2.Open (App.Path & "\image\filecopy.avi")
''
''    DoEvents
''    DoEvents
''
''    On Error GoTo diskError01
''
''    fileLength1 = FileLen(strPath)
''
''    Open strPath For Input As #1 ' Open file.
''
''    Do While Not EOF(1) ' Loop until end of file.
''        Label4 = "출고자료를 수신하고 있읍니다..! "
''        Line Input #1, strText  ' Read line into variable.
''
''        strdate01 = Trim(Mid(strText, 2, 8))
''        strTNo = Trim(Mid(strText, 11, 4))
''        strRejec = Trim(Mid(strText, 16, 1))
''        strTNo = Trim(Mid(strTNo, 1, 1)) & "-" & Trim(Mid(strTNo, 2, 3))
''
''        If strRejec = "2" Then
''            Query = "UPDATE TB_입출고 "
''            Query = Query & "SET 본출 = '出' "
''            Query = Query & "WHERE Trim(택번호) = '" & Trim(strTNo) & "' "
''            'Query = Query & "And   접수일자 = '" & Trim(strdate01) & "' "
''        ElseIf strRejec = "3" Then
''            Query = "UPDATE TB_입출고 "
''            Query = Query & "SET 본출 = '反' "
''            Query = Query & "WHERE Trim(택번호) = '" & Trim(strTNo) & "' "
''            'Query = Query & "And   접수일자 = '" & Trim(strdate01) & "'"
''
''        '             QueryreJec = "INSERT INTO TB_반품환불(입고일,택번호,구분) VALUES ('" & Trim(strdate01) & "','" & Trim(strTNo) & "','1')"
''        '             ADOCon.Execute QueryreJec 'INSERT INTO TB_반품환불(입고일,택번호,구분)  SELECT 입고일,택번호,'1' FROM TB_입출고 WHERE 택번호='0-648'
''            DoEvents
''        End If
''
''        ADOCon.Execute Query
''
''        DoEvents
''    Loop
''
''    Label4 = "출고자료 수신을 완료 했읍니다..!"
''
''    Close #1    ' Close file.
''
''    Ani2.Visible = False
''
''    Exit Sub
''
''diskError01:
''    Ani2.Visible = False
''    r_Value = MsgBox(" 출고데이타가 없습니다... " & Chr$(13) & Chr$(13) & _
''                     " 정상.할인자료를 받으려면 확인을 누르십시오." & Chr$(13) & Chr$(13) & _
''                     "[" & Err.Number & "] " & Err.Description, vbInformation, "출고 자료 수신")
''    ' If r_Value = vbRetry Then
''    '    Resume
''    ' ElseIf r_Value = vbCancel Then
''    '    Exit Sub
''    ' End If
''End Sub
''
''Private Sub dataSale()
''    '**************************************************************************************
''    '할인정보 테이블에 insert
''    '**************************************************************************************
''
''    Dim strText As String
''    Dim QueryCode As String
''    Dim rsCode As Recordset
''    Dim douCount As Double
''    Dim strPath   As String
''    Dim strNo As String
''    Dim strPrice As String
''    Dim strName As String
''    Dim strRatio As String
''    Dim strdate01 As String
''    Dim strdate02 As String
''    Dim Query02 As String
''
''    strCode = "SELECT 대리점번호 "
''    strCode = strCode & "FROM TB_대리점정보 "
''
''    Set rsCode = MyDB.OpenRecordset(strCode)
''
''    If rsCode.EOF = True Then
''        Label4 = "대리점코드가 존재하지 않읍니다..!"
''        rsCode.Close
''        Label4 = ""
''        Exit Sub
''    End If
''
''    strCode = Trim(rsCode!대리점번호)
''    If Len(strCode) = 1 Then
''        strCode = "00" & strCode
''    ElseIf Len(strCode) = 2 Then
''        strCode = "0" & strCode
''    End If
''
''    strPath = "A:\SALE" & Trim(strCode) & ".dat"
''    rsCode.Close
''
''    On Error GoTo diskError01
''
''    If Dir(strPath) = "" Then
''        Exit Sub
''    End If
''
''    Ani2.Visible = True
''    Ani2.AutoPlay = True
''    Ani2.Open (App.Path & "\image\filecopy.avi")
''
''    DoEvents
''
''    Open strPath For Input As #1 ' Open file.
''
''    Do While Not EOF(1) ' Loop until end of file.
''        Label4 = "할인자료를 변경하고 있읍니다..!"
''        Line Input #1, strText  ' Read line into variable.
''
''        strdate01 = Trim(Mid(strText, 2, 8))
''        strdate02 = Trim(Mid(strText, 11, 8))
''    Loop
''
''    Close #1
''    Query02 = "DELETE  "
''    Query02 = Query02 & "FROM TB_할인정보 " ' 98/06/23  '할인정보 전부삭제후 Insert
''    ADOCon.Execute Query02
''
''    DoEvents
''
''    Open strPath For Input As #1 ' Open file.
''
''    Do While Not EOF(1) ' Loop until end of file.
''
''        Label4 = "할인자료를 변경하고 있읍니다..!"
''        Line Input #1, strText  ' Read line into variable.
''
''        strdate01 = Trim(Mid(strText, 2, 8))
''        strdate02 = Trim(Mid(strText, 11, 8))
''        strNo = Trim(Mid(strText, 20, 3))
''        strPrice = Trim(Mid(strText, 24, 8))
''        strRatio = Trim(Mid(strText, 33, 2))
''        strName = Trim(Mid(strText, 36, 20))
''
''        If Len(strRatio) < 1 Then
''            strRatio = ""
''        End If
''
''        Query = "INSERT INTO TB_할인정보(시작일, 종료일, 구분코드, 품명, 가격, 비율) "
''        Query = Query & "VALUES ('" & Trim(strdate01) & "', "
''        Query = Query & "'" & Trim(strdate02) & "', "
''        Query = Query & "'" & strNo & "', "
''        Query = Query & "'" & strName & "', "
''        Query = Query & "'" & strPrice & "', "
''        Query = Query & "'" & strRatio & "')"
''
''        ADOCon.Execute Query
''
''        DoEvents
''    Loop
''
''    Close #1    ' Close file.
''    Label4 = "할인자료 변경을 완료했읍니다..!"
''    Ani2.Visible = False
''    Exit Sub
''
''diskError01:
''    Ani2.Visible = False
''    MsgBox " 할인자료를 읽은 중에 디스크에러가 발생하였습니다 " & Str$(VBA.Err.Number) & "  " & VBA.Err.Description, vbCritical, "할인정보"
''End Sub
''
''Private Sub dataPrice()
''    '**************************************************************************************
''    '참조코드 테이블에 update
''    '**************************************************************************************
''
''    Dim strText As String
''    Dim QueryCode As String
''    Dim rsCode As Recordset
''    Dim douCount As Double
''    Dim strPath01   As String
''    Dim strPath02   As String
''    Dim strPath03   As String
''    Dim strTNo As String
''    Dim strClfy As String
''    Dim strRejec As String
''    Dim QueryreJec As String
''    Dim strdate01 As String
''    Dim strFileName As String
''
''    strCode = "SELECT 대리점번호 "
''    strCode = strCode & "FROM TB_대리점정보 "
''
''    Set rsCode = MyDB.OpenRecordset(strCode)
''
''    If rsCode.EOF = True Then
''        Label4 = "대리점코드가 존재하지 않읍니다..!"
''        rsCode.Close
''        Label4 = ""
''        Exit Sub
''    End If
''
''    strCode = Trim(rsCode!대리점번호)
''
''    If Len(strCode) = 1 Then
''        strCode = "00" & strCode
''    ElseIf Len(strCode) = 2 Then
''        strCode = "0" & strCode
''    End If
''
''    rsCode.Close
''
''    On Error GoTo diskError01
''
''    strFileName = Dir("a:\????????" & strCode & ".dat")
''
''    If Not Trim(strFileName) = "" Then
''        Ani2.Visible = True
''        Ani2.AutoPlay = True
''        Ani2.Open (App.Path & "\image\filecopy.avi")
''
''        strPath01 = Trim(strFileName)
''        strPath02 = App.Path & "\BackData\" & Trim(strFileName)
''        DoEvents
''
''        Label4 = "가격자료 복사중 입니다..!"
''        FileCopy "a:\" & strPath01, strPath02
''        DoEvents
''
''        Label4 = "가격자료 수신을 완료했읍니다..!"
''        DoEvents
''    End If
''
''    Ani2.Visible = False
''    Ani2.Enabled = False
''
''    On Error GoTo diskError02
''
''    strFileName = Dir("a:\D????????" & strCode & ".dat")
''
''    If Not Trim(strFileName) = "" Then
''        Ani2.Visible = True
''        Ani2.AutoPlay = True
''        Ani2.Open (App.Path & "\image\filecopy.avi")
''
''        strPath01 = Trim(strFileName)
''        strPath02 = App.Path & "\BackData\" & Trim(strFileName)
''        DoEvents
''
''        Label4 = "목요세일자료 복사중 입니다..!"
''        FileCopy "a:\" & strPath01, strPath02
''        DoEvents
''
''        Label4 = "목요세일자료 수신을 완료 했읍니다..!"
''        DoEvents
''    End If
''
''    Ani2.Visible = False
''    Ani2.Enabled = False
''
''    On Error GoTo diskError03
''
''    strFileName = Dir("a:\R????????" & ".dat")
''
''    If Not Trim(strFileName) = "" Then
''        Ani2.Visible = True
''        Ani2.AutoPlay = True
''        Ani2.Open (App.Path & "\image\filecopy.avi")
''
''        strPath01 = Trim(strFileName)
''        strPath02 = App.Path & "\BackData\" & Trim(strFileName)
''        DoEvents
''
''        Label4 = "수선자료 복사중 입니다..!"
''        FileCopy "a:\" & strPath01, strPath02
''        DoEvents
''
''        Label4 = "수선자료 수신을 완료했읍니다..!"
''        DoEvents
''    End If
''
''    Ani2.Visible = False
''    Ani2.Enabled = False
''    Exit Sub
''
''diskError01:
''    Ani2.Visible = False
''    MsgBox " 가격자료 복사중 디스크에러가 발생하였습니다 " & Str$(VBA.Err.Number) & "  " & VBA.Err.Description, vbRetryCancel + vbCritical, "가격자료복사"
''    Exit Sub
''diskError02:
''    Ani2.Visible = False
''    MsgBox " 목요세일자료 복사중 디스크에러가 발생하였습니다 " & Str$(VBA.Err.Number) & "  " & VBA.Err.Description, vbRetryCancel + vbCritical, "목요세일자료복사"
''diskError03:
''    Ani2.Visible = False
''    MsgBox " 수선자료 복사중 디스크에러가 발생하였습니다 " & Str$(VBA.Err.Number) & "  " & VBA.Err.Description, vbRetryCancel + vbCritical, "수선자료복사"
''End Sub
''
''Private Sub DnData_Click()
''    While Not DskCHK
''        If MsgBox("A:드라이버에 디스켓을 넣으십시요", vbRetryCancel) = vbCancel Then
''            Exit Sub
''        End If
''    Wend
''
''    Call dataReceiv '본사출고 자료수신
''
''    DoEvents
''
''    Call dataSale   '할인자료수신(할인정보 insert)
''
''    DoEvents
''
''    Call dataPrice  '품목가격,목요세일 copy
''
''    DoEvents
''End Sub
''
''Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
''    KeyChk (KeyCode)
''End Sub
''
''Private Sub Form_Load()
''    'TitleSet "자료 받음"
''End Sub
