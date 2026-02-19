VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form P_05009 
   Caption         =   "큐엔솔브 자료 관리"
   ClientHeight    =   11640
   ClientLeft      =   2310
   ClientTop       =   4950
   ClientWidth     =   17070
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_05009.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11640
   ScaleWidth      =   17070
   WindowState     =   2  '최대화
   Begin SSSplitter.SSSplitter SSSplitter 
      Height          =   11640
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17070
      _ExtentX        =   30110
      _ExtentY        =   20532
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "P_05009.frx":058A
      Begin Threed.SSPanel panInput 
         Height          =   765
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   17040
         _ExtentX        =   30057
         _ExtentY        =   1349
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin VB.TextBox txtInput 
            Height          =   315
            Index           =   0
            Left            =   6480
            TabIndex        =   2
            Top             =   60
            Width           =   3555
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   0
            Left            =   5010
            TabIndex        =   3
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "작 업 경 로"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin MSComCtl2.DTPicker dtInput 
            Height          =   315
            Left            =   6480
            TabIndex        =   4
            Top             =   405
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   56360960
            CurrentDate     =   36686
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   3
            Left            =   5010
            TabIndex        =   5
            Top             =   405
            Visible         =   0   'False
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "작업일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   6
            Left            =   1530
            TabIndex        =   6
            Top             =   60
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   556
            _Version        =   262144
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            Begin Threed.SSOption optSelect 
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   7
               Top             =   30
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   450
               _Version        =   262144
               Caption         =   "FTP"
               Value           =   -1
            End
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   5
            Left            =   60
            TabIndex        =   8
            Top             =   60
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "작 업 경 로"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel panMain 
         Height          =   10830
         Left            =   15
         TabIndex        =   9
         Top             =   795
         Width           =   17040
         _ExtentX        =   30057
         _ExtentY        =   19103
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   315
            Left            =   1530
            TabIndex        =   10
            Top             =   120
            Width           =   12645
            _ExtentX        =   22304
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
         End
         Begin Threed.SSCommand cmdSubBtn 
            Height          =   555
            Index           =   0
            Left            =   3150
            TabIndex        =   11
            Top             =   1230
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   979
            _Version        =   262144
            Caption         =   "자료 보내기"
         End
         Begin Threed.SSCommand cmdSubBtn 
            Height          =   555
            Index           =   1
            Left            =   5250
            TabIndex        =   12
            Top             =   1230
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   979
            _Version        =   262144
            Caption         =   "자료 수신"
         End
         Begin Threed.SSPanel panCaption 
            Height          =   315
            Index           =   1
            Left            =   60
            TabIndex        =   13
            Top             =   120
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   262144
            Caption         =   "자 료 수 신"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
         End
      End
   End
End
Attribute VB_Name = "P_05009"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Option Explicit
''
''Dim Err_Num As Long
''Dim Err_Dec As String
''
''Dim sValue() As String
''Dim ConnectMode As ConnectMode_Type
''Dim m_sDateTime     As String
''Dim m_ActionDate    As String
''
''Dim LRecvSize       As Long
''Dim sSendRecvMode   As String
''
''Dim WithEvents myFtp    As clsFTP
''Const m_FTPServerIP = "203.238.178.116" '"211.255.17.10"
''Const m_FTPUserID = "cleanaid"
''Const m_FTPPassword = "!tpxkr?&"
'''Const m_FTPServerIP = "211.105.47.95"
'''Const m_FTPUserID = "abc"
'''Const m_FTPPassword = "abcde"
''
''
''Private Sub cmdSubBtn_Click(Index As Integer)
''    cmdSubBtn(0).Enabled = False
''    cmdSubBtn(1).Enabled = False
''
''    Select Case Store.Code
''
''        Case "1000" ' 본사일 경우
''            Select Case Index
''                Case 0          ' 자료 생성
''                    If DataSave_Offline_Order = True Then
''                        Call FTP_Send
''                    End If
''
''                Case 1          ' 자료 수신
''                    If FTP_Recv = True Then
''                        ' 체인점으로 전송 받은 자료를
''                        Call Move_Store_Input_Data
''                    End If
''
''            End Select
''
''        ' 타 지점및 유니트샵은 본사로 전송한다.
''        Case Else
''            Select Case Index
''                Case 0          ' 자료 생성
''                    Call DataSendToMaster
''
''                Case 1          ' 자료 수신
''
''
''            End Select
''
''    End Select
''
''    cmdSubBtn(0).Enabled = True
''    cmdSubBtn(1).Enabled = True
''
''End Sub
''
''Private Sub Form_Activate()
'''    pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
''
''End Sub
''
''Private Sub Form_Load()
''    Dim optSel  As Integer
''    optSelect(0).Value = True
''
''    dtInput.Value = Date
''
''    txtInput(0).Text = GetIniStr("SERVER DATA", "SendPath", "", sIniFile)
''    txtInput(0).ToolTipText = txtInput(0).Text
''    PanelsMsg ("")
''End Sub
''
''Private Sub Form_Unload(Cancel As Integer)
''    Call INIWrite("SERVER DATA", "ConnectMode", CStr(ConnectMode), sIniFile)
''
''    P_08003_Flag = False
''End Sub
''
''
''Private Function DataSave_Offline_Order() As Boolean
''    '하드디스크작성
''    Dim FileName As String
''    Dim sDownPath As String
''    Dim bResult As Boolean
''
''    On Error GoTo DataSave_Offline_Order_Error
''    DataSave_Offline_Order = False
''
''    sDownPath = GetIniStr("SERVER DATA", "SendPath", "", sIniFile)
''
''    panCaption(0).Visible = True
''
''    If dtInput.Value = "" Then
''       Exit Function
''    End If
''
''    ' 파일 생성 시간을 공통으로 사용하기 위하여
''    m_sDateTime = Format(Now, "yyyymmddhhmmss000")
''    ' 전송 기준 일자
''    m_ActionDate = Format(dtInput.Value, "yyyymmdd")
''
''    If Dir(App.Path & "\QN", vbDirectory) = "" Then
''        MkDir App.Path & "\QN"
''    End If
''
''    bResult = Offline_MetalInfo_File
''    If bResult = True Then
''        bResult = Offline_Document_File
''    End If
''
''    DataSave_Offline_Order = bResult
''    On Error GoTo 0
''    Exit Function
''
''DataSave_Offline_Order_Error:
''    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DataSave_Offline_Order of Form P_05009"
''
''End Function
''
''
''Private Function Offline_MetalInfo_File() As Boolean
''    '하드디스크작성
''    Dim FileName As String
''    Dim FNumber As Integer
''    Dim SCount  As String
''    Dim RS01 As ADODB.Recordset
''
''    On Error GoTo Offline_MetalInfo_File_Error
''
''    Offline_MetalInfo_File = False
''
''    ReDim sValue(0)
''
''    sValue(0) = m_ActionDate
''
''    Set RS01 = New ADODB.Recordset
''    Set RS01 = ExecPro("SP_05009_01", sValue(), Err_Num, Err_Dec)
''
''    If RS01.RecordCount <= 0 Then
''        Call PanelsMsg("생성할 자료가 없습니다.")
''        RS01.Close
''        Exit Function
''    End If
''    SCount = Format(RS01.RecordCount, "000000")
''    RS01.Close
''
''    FileName = "Offline_Order.I01." & Format(Now, "yyyymmdd")
''
''    Call PanelsMsg("Offline Order Metainfo File을 생성중 입니다.")
''
''    FNumber = FreeFile
''    Open App.Path & "\QN\" & FileName For Output As #FNumber
''
''    Print #FNumber, "HM";
''    Print #FNumber, "10";
''    Print #FNumber, "VCO01";
''    Print #FNumber, m_QN_PartnerID;
''    Print #FNumber, m_sDateTime
''
''    Print #FNumber, "DM";
''    Print #FNumber, "00001"
''
''    Print #FNumber, "IF";
''    Print #FNumber, m_sDateTime;
''    Print #FNumber, Left("Offline_Order.D01." & Format(Now, "yyyymmdd") & Space(40), 40);
''    Print #FNumber, SCount
''
''
''    Close #FNumber
''    Call PanelsMsg("Offline Order Metainfo File을 생성 완료")
''    Offline_MetalInfo_File = True
''
''    On Error GoTo 0
''    Exit Function
''
''Offline_MetalInfo_File_Error:
''    Close #FNumber
''    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Offline_MetalInfo_File of Form P_05009"
''
''End Function
''
''
''Private Function Offline_Document_File() As Boolean
''    '하드디스크작성
''    Dim FileName As String
''    Dim FNumber As Integer
''    Dim SCount  As String
''    Dim RS01 As ADODB.Recordset
''
''    On Error GoTo Offline_Document_File_Error
''
''    Offline_Document_File = False
''
''    ReDim sValue(0)
''
''    sValue(0) = m_ActionDate
''
''    Set RS01 = New ADODB.Recordset
''    Set RS01 = ExecPro("SP_05009_01", sValue(), Err_Num, Err_Dec)
''
''    If RS01.RecordCount <= 0 Then
''        Call PanelsMsg("생성할 자료가 없습니다.")
''        RS01.Close
''        Exit Function
''    End If
''    SCount = Format(RS01.RecordCount, "000000")
''
''    FileName = "Offline_Order.D01." & Format(Now, "yyyymmdd")
''
''    Call PanelsMsg("Offline Order Document File을 생성중 입니다.")
''
''    FNumber = FreeFile
''    Open App.Path & "\QN\" & FileName For Output As #FNumber
''
''    Print #FNumber, "HD";
''    Print #FNumber, "10";
''    Print #FNumber, "VCO01";
''    Print #FNumber, m_QN_PartnerID;
''    Print #FNumber, m_sDateTime
''
''    Print #FNumber, "DD";
''    Print #FNumber, SCount
''
''    Do While Not RS01.EOF
''        Print #FNumber, "FO";
''        Print #FNumber, Left(m_QN_PartnerID & Left(RS01.Fields("KeyCode"), 8) & RS01.Fields("InputNumber") & Space(20), 20);
''        Print #FNumber, Left(RS01.Fields("InputDate") & Space(17), 17);
''        Print #FNumber, Left(RS01.Fields("InputID") & Space(20), 20);
''        Print #FNumber, Left(RS01.Fields("InputName") & Space(20), 20);
''        Print #FNumber, Left(RS01.Fields("EMail") & Space(40), 40);
''        Print #FNumber, Left(RS01.Fields("UserCode") & Space(2), 2);
''        Print #FNumber, Left(Replace(RS01.Fields("UserNumber") & Space(13), "-", ""), 13);
''        Print #FNumber, Left(RS01.Fields("StoreCode") & Space(20), 20);
''        Print #FNumber, Left(RS01.Fields("SaleGubunCode") & Space(2), 2);
''        Print #FNumber, Left(RS01.Fields("SaleEndDate") & Space(8), 8);
''        Print #FNumber, Right("00000000" & RS01.Fields("Price"), 8);
''        Print #FNumber, Left(RS01.Fields("DevTimeCode") & Space(2), 2);
''        Print #FNumber, Right("000000" & RS01.Fields("ItemCount"), 6)
''
''        ' 항목별 내용을 추가 기록하낟.
''        Call Offline_Document_Sub1(FNumber, RS01.Fields("KeyCode"))
''        RS01.MoveNext
''    Loop
''
''    Close #FNumber
''    Call PanelsMsg("Offline Order Document File을 생성 완료")
''    Offline_Document_File = True
''
''    On Error GoTo 0
''    Exit Function
''
''
''
''Offline_Document_File_Error:
''    Close #FNumber
''    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Offline_Document_File of Form P_05009"
''End Function
''
''
'''--------------------------------------------------------------------------------------------------------------
''' Procedure : Offline_Document_Sub1
''' DateTime  : 2006-11-05 15:00
''' Author    : pds2004
''' Purpose   : 동일한 KeyCode를 검색하여 각각의 항목을  Document 파일에 추가 기록한다.
'''--------------------------------------------------------------------------------------------------------------
''Private Function Offline_Document_Sub1(ByVal FNumber As Integer, ByVal sKeyCode As String) As Boolean
''    '하드디스크작성
''    Dim SCount  As String
''    Dim RS01 As ADODB.Recordset
''
''    On Error GoTo Offline_Document_File_Error
''
''    Offline_Document_Sub1 = False
''
''    ReDim sValue(0)
''
''    sValue(0) = sKeyCode
''
''    Set RS01 = New ADODB.Recordset
''    Set RS01 = ExecPro("SP_05009_02", sValue(), Err_Num, Err_Dec)
''
''    If RS01.RecordCount <= 0 Then
''        Call PanelsMsg("생성할 자료가 없습니다.")
''        RS01.Close
''        Exit Function
''    End If
''
''    Do While Not RS01.EOF
''        Print #FNumber, "CI";
''        Print #FNumber, Left(RS01.Fields("ItemIndex") & Space(6), 6);
''        Print #FNumber, Left(RS01.Fields("Tag") & Space(20), 20);
''        Print #FNumber, Left(RS01.Fields("GoodsCode") & Space(16), 16);
''        Print #FNumber, Left(RS01.Fields("SizeGubun") & Space(2), 2);
''        Print #FNumber, Left(RS01.Fields("SizeCode") & Space(2), 2);
''        Print #FNumber, Left(RS01.Fields("Color") & Space(10), 10);
''        Print #FNumber, Left(RS01.Fields("BrandName") & Space(20), 20);
''        Print #FNumber, Right("0000000000" & RS01.Fields("BuyPrice"), 10);
''        Print #FNumber, Left(RS01.Fields("BuyDate") & Space(8), 8);
''        Print #FNumber, Left(RS01.Fields("ASGubun") & Space(2), 2);
''        Print #FNumber, Right("000" & RS01.Fields("BleCount"), 3)
''
''        If Val(RS01.Fields("BleCount")) > 0 Then
''            Call Offline_Document_Sub2(FNumber, RS01.Fields("KeyCode"), RS01.Fields("ItemIndex"))
''        End If
''
''
''        RS01.MoveNext
''    Loop
''
''    RS01.Close
''
''    Call PanelsMsg("Offline Order Document File의 항목 생성 완료")
''    Offline_Document_Sub1 = True
''
''    On Error GoTo 0
''    Exit Function
''
''
''
''Offline_Document_File_Error:
''    Resume
''    Close #FNumber
''    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Offline_Document_File of Form P_05009"
''End Function
''
'''--------------------------------------------------------------------------------------------------------------
''' Procedure : Offline_Document_Sub2
''' DateTime  : 2006-11-05 15:00
''' Author    : pds2004
''' Purpose   : 동일한 KeyCode를 검색하여 각각의 항목중 하자 자료를  Document 파일에 추가 기록한다.
'''--------------------------------------------------------------------------------------------------------------
''Private Function Offline_Document_Sub2(ByVal FNumber As Integer, ByVal sKeyCode As String, ByVal sItemIndex As String) As Boolean
''    '하드디스크작성
''    Dim SCount  As String
''    Dim RS01 As ADODB.Recordset
''
''    On Error GoTo Offline_Document_File_Error
''
''    Offline_Document_Sub2 = False
''
''    ReDim sValue(1)
''
''    sValue(0) = sKeyCode
''    sValue(1) = sItemIndex
''
''    Set RS01 = New ADODB.Recordset
''    Set RS01 = ExecPro("SP_05009_03", sValue(), Err_Num, Err_Dec)
''
''    If RS01.RecordCount <= 0 Then
''        Call PanelsMsg("생성할 자료가 없습니다.")
''        RS01.Close
''        Exit Function
''    End If
''
''    Do While Not RS01.EOF
''        Print #FNumber, "DF";
''        Print #FNumber, Left("0000" & RS01.Fields("ItemCount"), 6);
''        Print #FNumber, Left(RS01.Fields("ItemRemark") & Space(50), 50)
''        RS01.MoveNext
''    Loop
''
''    RS01.Close
''
''    Call PanelsMsg("Offline Order Document File을 생성 완료")
''    Offline_Document_Sub2 = True
''
''    On Error GoTo 0
''    Exit Function
''
''
''
''Offline_Document_File_Error:
''    Close #FNumber
''    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Offline_Document_File of Form P_05009"
''End Function
''
''
''Private Sub myFtp_Error(Err_Num As Long, Err_Des As String)
''    If Err_Num <> 12003 Then
''        PanelsMsg Err_Des
''    End If
''
''End Sub
''
''Private Sub myFtp_FileTransferProgress(lCurrentBytes As Long, lTotalBytes As Long)
''    On Error GoTo ERR_RTN
''    If sSendRecvMode = "SEND" Then
''        If ProgressBar1.Max = 1 Then
''            ProgressBar1.Max = lTotalBytes
''        End If
''        ' 프로그래스바를 움직인다.
''        ProgressBar1.Value = lCurrentBytes
''
''    ElseIf sSendRecvMode = "RECV" Then
''        ' 프로그래스바를 움직인다.
''        ProgressBar1.Value = LRecvSize + lCurrentBytes
''
''    End If
''    Exit Sub
''ERR_RTN:
''
''End Sub
''
''Private Sub FTP_Send()
''    Dim sFrom   As String
''    Dim sTo     As String
''    Dim sFilename As String
''
''    Set myFtp = New clsFTP
''
''    sSendRecvMode = "SEND"
''    panCaption(1).Caption = "자 료 전 송"
''    PanelsMsg "파일을 전송중 입니다. 잠시만 기다려 주십시요."
''
''    If myFtp.OpenConnection(m_FTPServerIP, 21, m_FTPUserID, m_FTPPassword) = False Then
''        PanelsMsg "서버와 연결하지 못하였습니다."
''        Exit Sub
''    End If
''
''
''    sFilename = Dir(App.Path & "\QN\*.*")
''    Do While sFilename <> ""   ' 루프(loop)를 시작합니다.
''       ' 현재 디렉토리와 포함하는 디렉토리를 무시합니다.
''        If sFilename <> "." And sFilename <> ".." Then
''            ' sFilename이 디렉토리인지 확인하기 위해서 비트별(bitwise) 비교를 사용합니다.
''            If (GetAttr(App.Path & "\QN\" & sFilename) And vbDirectory) = vbDirectory Then
''                Debug.Print sFilename   ' 항목만 표시합니다
''            Else
''
''                sFrom = App.Path & "\QN\" & sFilename
''                sTo = sFilename
''                If myFtp.FTPUploadFile(sFrom, sTo) = False Then
''                    PanelsMsg "파일을 전송하지 못하였습니다."
''                Else
''                    PanelsMsg "파일을 전송이 완료 되었습니다."
''
''                    Kill sFrom = App.Path & "\QN\" & sFilename
''
''                    ReDim sValue(8)
''                    ' 파일 전송을 완료하면 완료 기록을 한다.
''                    sValue(0) = "2" ' QN 솔브 전송 일자 기록
''                    sValue(1) = "ALL_TB"
''                    sValue(2) = ""
''                    sValue(3) = ""
''                    sValue(4) = ""
''                    sValue(5) = m_ActionDate
''                    sValue(6) = ""
''                    sValue(7) = Format(Now, "yyyymmddhhmmss")
''
''                    ' 본사로 전송한 자료는 전송 일자를 업데이트 한다.
''                    Call ExecPro("SP_05009_10", sValue(), Err_Num, Err_Dec)
''                    If Err_Num > 0 Then PanelsMsg Err_Dec
''
''                End If
''
''            End If   ' 그것은 디렉토리를 표시합니다.
''        End If
''        sFilename = Dir   ' 다음 항목을 읽어들입니다.
''    Loop
''
''    Set myFtp = Nothing
''
''End Sub
''
''
''Private Function FTP_Recv() As Boolean
''    Dim sFrom   As String
''    Dim sTo     As String
''    Dim FtpName As String
''    Dim MyDirList   As cDirList
''    Dim i       As Integer
''    Dim lMaxSize    As Long
''
''    FTP_Recv = False
''    Set myFtp = New clsFTP
''
''    sSendRecvMode = "RECV"
''    panCaption(1).Caption = "자 료 수 신"
''
''    PanelsMsg "[" & m_FTPServerIP & "] 서버와 연결중 입니다. 잠시만 기다려 주십시요."
''
''    If myFtp.OpenConnection(m_FTPServerIP, 21, m_FTPUserID, m_FTPPassword) = False Then
''        PanelsMsg "[" & m_FTPServerIP & "] 서버와 연결하지 못하였습니다."
''        Set myFtp = Nothing:    Exit Function
''    End If
''
''    PanelsMsg "파일을 수신중 입니다. 잠시만 기다려 주십시요."
''    'myFtp.SetFTPDirectory "cleanaid"
''
''    Set MyDirList = New cDirList
''    Set MyDirList = myFtp.GetDirectoryListing("*.*")
''
''    If MyDirList Is Nothing Then
''        PanelsMsg "다운로드할 파일이 없습니다."
''        MsgBox "다운로드할 파일이 없습니다.     ", vbInformation, "확인"
''        Set myFtp = Nothing:    Exit Function
''
''    ElseIf MyDirList.Count > 0 Then
''        sTo = App.Path & "\QN_Downloads\"
''        If Dir(sTo, vbDirectory) = "" Then
''            MkDir sTo
''
''        End If
''
''        lMaxSize = 0
''        For i = 1 To MyDirList.Count
''            If Left(MyDirList.Item(i).FileName, 1) <> "." Then
''                lMaxSize = lMaxSize + Val(MyDirList.Item(i).FileSize)
''            End If
''        Next i
''
''        LRecvSize = 0
''        ProgressBar1.Max = lMaxSize
''
''        For i = 1 To MyDirList.Count
''            If Left(MyDirList.Item(i).FileName, 1) <> "." Then
''                PanelsMsg MyDirList.Item(i).FileName & "다운로드 시작"
''                If myFtp.FTPDownloadFile(sTo & MyDirList.Item(i).FileName, MyDirList.Item(i).FileName) = True Then
''                    'Call myFtp.DeleteFTPFile(MyDirList.Item(i).Filename)
''                End If
''                LRecvSize = LRecvSize + MyDirList.Item(i).FileSize
''                PanelsMsg MyDirList.Item(i).FileName & "다운로드 종료"
''            End If
''        Next i
''        FTP_Recv = True
''        PanelsMsg CStr(MyDirList.Count) & "개의 파일을 다운로드 하였습니다."
''        MsgBox CStr(MyDirList.Count) & "개의 파일을 다운로드 하였습니다.", vbInformation, "확인"
''    End If
''
''    Set myFtp = Nothing
''    Set MyDirList = Nothing
''End Function
''
'''--------------------------------------------------------------------------------------------------------------
''' Procedure : Move_Store_Input_Data
''' DateTime  : 2006-11-09 10:51
''' Author    : pds2004
''' Purpose   : 전달 받은 자료를 \\CleanaAid\CleanData로 이동 시킨다.
'''--------------------------------------------------------------------------------------------------------------
''Public Function Move_Store_Input_Data() As Boolean
''
''    Dim bResult As Boolean
''    Dim sTo     As String
''    Dim sFrom   As String
''    Dim sFilename As String
''
''    On Error GoTo Move_Store_Input_Data_Error
''
''    PanelsMsg "다운받은 체인점 접수 현황 복사중......"
''
''    sFrom = App.Path & "\QN_Downloads\"
''    sTo = GetIniStr("SERVER DATA", "ReceivePath", "", sIniFile)
''    If Right(sTo, 1) <> "\" Then sTo = sTo & "\"
''
''    ' 매출 자료
''    sFilename = Dir(sFrom & "????????-???-?.DAT")
''    ' 여러개의 파일을 배열에 넣어준다.
''    Do While Len(sFilename) > 0
''
''        FileCopy sFrom & sFilename, sTo & sFilename
''        Kill sFrom & sFilename
''
''        sFilename = Dir
''
''    Loop
''
''    ' 고객 자료
''    sFilename = Dir(sFrom & "C????????-???-?.DAT")
''    ' 여러개의 파일을 배열에 넣어준다.
''    Do While Len(sFilename) > 0
''
''        FileCopy sFrom & sFilename, sTo & sFilename
''        Kill sFrom & sFilename
''
''        sFilename = Dir
''
''    Loop
''
''
''
''    PanelsMsg "복사 완료."
''
''    Move_Store_Input_Data = bResult
''
''    On Error GoTo 0
''    Exit Function
''
''Move_Store_Input_Data_Error:
''
''    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Move_Store_Input_Data of Form P_05009"
''
''End Function
''
''
'''--------------------------------------------------------------------------------------------------------------
''' Procedure : DataSendToMaster
''' DateTime  : 2006-11-09 04:00
''' Author    : pds2004
''' Purpose   : 접수된 내용을 본사로 전송 시킨다.
'''--------------------------------------------------------------------------------------------------------------
''Public Function DataSendToMaster() As Boolean
''    Dim SSQL    As String
''    Dim bResult As Boolean
''    Dim RS01    As ADODB.Recordset
''    Dim sUpdateTime As String
''
''
''    On Error GoTo DataSendToMaster_Error
''
''    DataSendToMaster = False
''
''    If DBOpen_Master = False Then
''        PanelsMsg "본사와 연결하지 못하였습니다. 본사 설정 정보를 확인하여 주십시요."
''        Exit Function
''    End If
''
''    ReDim sValue(1)
''
''
''    sValue(0) = "QN_Sale_TB"
''    Set RS01 = New ADODB.Recordset
''    Set RS01 = ExecProMaster("SP_05009_04", sValue(), Err_Num, Err_Dec)
''
''    '모든 업데이트 시간을 동일하게 설정하기 위하여
''    sUpdateTime = Format(Now, "yyyymmddhhmmss")
''    Do While Not RS01.EOF
''        ReDim sValue(16)
''        sValue(0) = CStr(RS01.Fields("CleanStore") & "")
''        sValue(1) = CStr(RS01.Fields("CompanyCode") & "")
''        sValue(2) = CStr(RS01.Fields("KeyCode") & "")
''        sValue(3) = CStr(RS01.Fields("MemRecord") & "")
''        sValue(4) = CStr(RS01.Fields("InputNumber") & "")
''        sValue(5) = CStr(RS01.Fields("InputDate") & "")
''        sValue(6) = CStr(RS01.Fields("InputID") & "")
''        sValue(7) = CStr(RS01.Fields("InputName") & "")
''        sValue(8) = CStr(RS01.Fields("EMail") & "")
''        sValue(9) = CStr(RS01.Fields("UserCode") & "")
''        sValue(10) = CStr(RS01.Fields("UserNumber") & "")
''        sValue(11) = CStr(RS01.Fields("StoreCode") & "")
''        sValue(12) = CStr(RS01.Fields("SaleGubunCode") & "")
''        sValue(13) = CStr(RS01.Fields("SaleEndDate") & "")
''        sValue(14) = CStr(RS01.Fields("Price") & "")
''        sValue(15) = CStr(RS01.Fields("DevTimeCode") & "")
''        sValue(16) = CStr(RS01.Fields("ItemCount") & "")
''
''        Call ExecProMaster("SP_08002_01", sValue(), Err_Num, Err_Dec)
''        If Err_Num > 0 Then
''            PanelsMsg Err_Dec
''        Else
''            sValue(0) = "1" ' 지사/유니트에서 본사 전송일자 세팅
''            sValue(1) = "QN_Sale_TB"
''            sValue(2) = CStr(RS01.Fields("CompanyCode") & "")
''            sValue(3) = CStr(RS01.Fields("CleanStore") & "")
''            sValue(4) = CStr(RS01.Fields("KeyCode") & "")
''            sValue(5) = CStr(RS01.Fields("InputNumber") & "")
''            sValue(6) = sUpdateTime
''            sValue(7) = ""
''
''            ' 본사로 전송한 자료는 전송 일자를 업데이트 한다.
''            Call ExecProMaster("SP_05009_10", sValue(), Err_Num, Err_Dec)
''            If Err_Num > 0 Then PanelsMsg Err_Dec
''
''        End If
''
''        RS01.MoveNext
''    Loop
''    RS01.Close
''
''    sValue(0) = "QN_Sale1_TB"
''    Set RS01 = ExecProMaster("SP_05009_04", sValue(), Err_Num, Err_Dec)
''
''    Do While Not RS01.EOF
''        ReDim sValue(16)
''        sValue(0) = CStr(RS01.Fields("CleanStore") & "")
''        sValue(1) = CStr(RS01.Fields("CompanyCode") & "")
''        sValue(2) = CStr(RS01.Fields("KeyCode") & "")
''        sValue(3) = CStr(RS01.Fields("ItemRecord") & "")
''        sValue(4) = CStr(RS01.Fields("ItemIndex") & "")
''        sValue(5) = CStr(RS01.Fields("InputDate") & "")
''        sValue(6) = CStr(RS01.Fields("Tag") & "")
''        sValue(7) = CStr(RS01.Fields("GoodsCode") & "")
''        sValue(8) = CStr(RS01.Fields("SizeGubun") & "")
''        sValue(9) = CStr(RS01.Fields("SizeCode") & "")
''        sValue(10) = CStr(RS01.Fields("Color") & "")
''        sValue(11) = CStr(RS01.Fields("BrandName") & "")
''        sValue(12) = CStr(RS01.Fields("BuyPrice") & "")
''        sValue(13) = CStr(RS01.Fields("BuyDate") & "")
''        sValue(14) = CStr(RS01.Fields("ASGubun") & "")
''        sValue(15) = CStr(RS01.Fields("BleCount") & "")
''
''        Call ExecProMaster("SP_08002_02", sValue(), Err_Num, Err_Dec)
''        If Err_Num > 0 Then
''            PanelsMsg Err_Dec
''        Else
''            sValue(0) = "1" ' 지사/유니트에서 본사 전송일자 세팅
''            sValue(1) = "QN_Sale1_TB"
''            sValue(2) = CStr(RS01.Fields("CompanyCode") & "")
''            sValue(3) = CStr(RS01.Fields("CleanStore") & "")
''            sValue(4) = CStr(RS01.Fields("KeyCode") & "")
''            sValue(5) = CStr(RS01.Fields("ItemIndex") & "")
''            sValue(6) = sUpdateTime
''            sValue(7) = ""
''
''            ' 본사로 전송한 자료는 전송 일자를 업데이트 한다.
''            Call ExecProMaster("SP_05009_10", sValue(), Err_Num, Err_Dec)
''            If Err_Num > 0 Then PanelsMsg Err_Dec
''
''        End If
''
''        RS01.MoveNext
''    Loop
''    RS01.Close
''
''    sValue(0) = "QN_Sale2_TB"
''    Set RS01 = ExecProMaster("SP_05009_04", sValue(), Err_Num, Err_Dec)
''
''    Do While Not RS01.EOF
''        ReDim sValue(16)
''        sValue(0) = CStr(RS01.Fields("CleanStore") & "")
''        sValue(1) = CStr(RS01.Fields("CompanyCode") & "")
''        sValue(2) = CStr(RS01.Fields("KeyCode") & "")
''        sValue(3) = CStr(RS01.Fields("ItemIndex") & "")
''        sValue(4) = CStr(RS01.Fields("ItemCount") & "")
''        sValue(5) = CStr(RS01.Fields("InputDate") & "")
''        sValue(6) = CStr(RS01.Fields("ItemRemark") & "")
''
''        Call ExecProMaster("SP_08002_03", sValue(), Err_Num, Err_Dec)
''        If Err_Num > 0 Then
''            PanelsMsg Err_Dec
''        Else
''            sValue(0) = "1" ' 지사/유니트에서 본사 전송일자 세팅
''            sValue(1) = "QN_Sale2_TB"
''            sValue(2) = CStr(RS01.Fields("CompanyCode") & "")
''            sValue(3) = CStr(RS01.Fields("CleanStore") & "")
''            sValue(4) = CStr(RS01.Fields("KeyCode") & "")
''            sValue(5) = CStr(RS01.Fields("ItemIndex") & "")
''            sValue(6) = sUpdateTime
''            sValue(7) = ""
''
''            ' 본사로 전송한 자료는 전송 일자를 업데이트 한다.
''            Call ExecProMaster("SP_05009_10", sValue(), Err_Num, Err_Dec)
''            If Err_Num > 0 Then PanelsMsg Err_Dec
''
''        End If
''
''        RS01.MoveNext
''    Loop
''    RS01.Close
''
''    DataSendToMaster = True
''
''    On Error GoTo 0
''    Exit Function
''
''DataSendToMaster_Error:
''
''    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DataSendToMaster of Form P_05009"
''    Resume
''End Function
