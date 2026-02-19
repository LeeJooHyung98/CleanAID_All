Attribute VB_Name = "basMain"
Option Explicit

Public DB_Server As String    '서버
Public DataBase  As String    '데이터베이스
Public DBMS      As String    '데이터베이스 종류

Public m_id As String
Public iRow     As Integer

Public Err_Num As Long
Public Err_Dec As String

Public ADOConCleanAid           As ADODB.Connection
Public ADOOldServer             As ADODB.Connection
Public ADONewServer             As ADODB.Connection

Public ADORs            As ADODB.Recordset
Public ADORs2           As ADODB.Recordset
Public SubRs            As ADODB.Recordset
Public Rs               As ADODB.Recordset

Public Query   As String
Public tMsg    As String
Public AppPath As String

Public DB_Path As String
Public iniFile As String  'ini 파일

Public i       As Long
Public j       As Long

Public Ret     As Long
Public Rtn     As Integer 'MessageBox Return 값...
Public x       As Boolean 'Spread Excel File Save...

Public 구분    As Integer '거래처(0)/거래(1)구분

Public Param   As Integer
Public XML     As String

'------------------------------------------------------------------------------------------
' 자기 IP 알아내기
'------------------------------------------------------------------------------------------
Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128
Public Const ERROR_SUCCESS As Long = 0
Public Const WS_VERSION_REQD As Long = &H101
Public Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD As Long = 1
Public Const SOCKET_ERROR As Long = -1

Public Type HOSTENT
    hName     As Long
    hAliases  As Long
    hAddrType As Integer
    hLen      As Integer
    hAddrList As Long
End Type

Public Type WSADATA
    wVersion     As Integer
    wHighVersion As Integer
    szDescription(0 To MAX_WSADescription) As Byte
    szSystemStatus(0 To MAX_WSASYSStatus) As Byte
    wMaxSockets  As Integer
    wMaxUDPDG    As Integer
    dwVendorInfo As Long
End Type

Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Public Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Public Declare Function gethostname Lib "WSOCK32.DLL" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Public Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal szHost As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

'------------------------------------------------------------------------------------------
' Centerform 함수
'------------------------------------------------------------------------------------------
Public Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Public Const WM_MENUSELECT = &H11F
Public Const HWND_BROADCAST = &HC0E0FF
Public Const GWL_WNDPROC = -4

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'DLL Functions
#If Win32 Then
    Private Declare Function GetClientrect& Lib "user32" Alias "GetClientRect" (ByVal hwnd&, Rct As RECT)
    Private Declare Function GetParent& Lib "user32" (ByVal hwnd&)
#Else
    Private Declare Function GetClientrect% Lib "USER" (ByVal hwnd%, Rct As RECT)
    Private Declare Function GetParent% Lib "USER" (ByVal hwnd%)
#End If
'------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------
' Browser for Folder
'------------------------------------------------------------------------------------------
Public Type BROWSEINFO
   hOwner           As Long
   pidlRoot         As Long
   pszDisplayName   As String
   lpszTitle        As String
   ulFlags          As Long
   lpfn             As Long
   lParam           As Long
   iImage           As Long
End Type

Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const BIF_DONTGOBELOWDOMAIN = &H2
Public Const BIF_STATUSTEXT = &H4
Public Const BIF_RETURNFSANCESTORS = &H8
Public Const BIF_BROWSEFORCOMPUTER = &H1000
Public Const BIF_BROWSEFORPRINTER = &H2000
Public Const MAX_PATH = 260

Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
'------------------------------------------------------------------------------------------

'최상위 폼
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2

Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'------------------------------------------------------------------------
' .INI 핸드링에 관련된 선언...
'------------------------------------------------------------------------
'Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpSectionName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'fn_Readini를 사용할때는 아래 GetPrivateProfileString을 사용한다.
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpSectionName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Declare Function sndPlaySound Lib "mmsystem.dll" (ByVal lpszSoundName As String, ByVal uFlags As Integer)

'====================================================================================================
' Procedure : NewServer_Connection
' DateTime  : 2008-04-15 04:13
' Author    : pds2004
' Last Edit :
' Return Val:
'----------------------------------------------------------------------------------------------------
' 설     명 : 최초 신규 매장코드가 없을 경우 본사에 등록되어있는 내용을 가저온다.
'====================================================================================================
Public Function NewServer_Connection(HostCon As ADODB.Connection, Optional strServerDB As String = "") As Boolean
    Dim sServer   As String
    Dim sDatabase As String
    Dim sID       As String
    Dim sPWD      As String
    
    On Error GoTo ErrRtn

    sServer = Get_Decrypt(GetIniStr("SERVER", "SERVER", "", iniFile), "")    '
    sDatabase = Get_Decrypt(GetIniStr("SERVER", "DATABASE", "", iniFile), "") '
    sID = Get_Decrypt(GetIniStr("SERVER", "ID", "", iniFile), "")            '
    sPWD = Get_Decrypt(GetIniStr("SERVER", "PWD", "", iniFile), "")          '
 
    If strServerDB <> "" Then
        sDatabase = strServerDB
        'sPWD = "cleanaid1996!@#"
    End If
    
'    If strServerDB = "LAUNDRY1000" Then
'        sDatabase = "LAUNDRY1000"
'        sPWD = "cleanaid1996!@#"
'    End If
    
    frmMain.lblServer.Caption = sServer & ""
    
    m_id = sID
    
    Set HostCon = Nothing
    Set HostCon = New ADODB.Connection

    If HostCon.State = adStateOpen Then HostCon.Close

    With HostCon
        .ConnectionString = "Provider=SQLOLEDB;Persist Security Info=False;User ID=" & sID & ";Password=" & sPWD & ";Initial Catalog=" & sDatabase & ";Data Source=" & sServer
        .ConnectionTimeout = 10
        .CommandTimeout = 30
        .Open
    End With
    
    NewServer_Connection = True
    
    On Error GoTo 0
    
    Exit Function

ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
    NewServer_Connection = False
End Function

Public Function ConnectOldServerCheck(HostCon As ADODB.Connection) As Boolean
    On Error GoTo ErrRtn
    
    Query = ""
    Query = Query & "Provider=SQLOLEDB.1;"
    Query = Query & "Persist Security Info=False;"
    Query = Query & "User ID=sa;"
    Query = Query & "Password=;"
    Query = Query & "Initial Catalog=Laundry;"
    Query = Query & "Data Source=store.clean-aid.co.kr,8657"
    
    Set HostCon = Nothing
    Set HostCon = New ADODB.Connection

    If HostCon.State = adStateOpen Then HostCon.Close
    
    HostCon.ConnectionTimeout = 10
    HostCon.CommandTimeout = 30
    HostCon.Open Query

    ConnectOldServerCheck = True
    
    On Error GoTo 0
    
    Exit Function

ErrRtn:
    ConnectOldServerCheck = False
End Function

Public Sub ProfileSaveItem(lpSectionName As String, lpKeyName As String, lpValue As String, iniFile As String)
    Call WritePrivateProfileString(lpSectionName, lpKeyName, lpValue, iniFile)
End Sub

Public Function ProfileGetItem(lpSectionName As String, lpKeyName As String, defaultValue As String, iniFile As String) As String
    Dim success As Long
    Dim nSize   As Long
    Dim Ret     As String
     
    Ret = Space$(2048)
    nSize = Len(Ret)
    success = GetPrivateProfileString(lpSectionName, lpKeyName, defaultValue, Ret, nSize, iniFile)
    
    If success Then
        ProfileGetItem = Left$(Ret, success)
    End If
End Function

Public Sub ProfileDeleteItem(lpSectionName As String, lpKeyName As String, iniFile As String)
    Call WritePrivateProfileString(lpSectionName, lpKeyName, vbNullString, iniFile)
End Sub

Public Sub ProfileDeleteSection(lpSectionName As String, iniFile As String)
    Call WritePrivateProfileString(lpSectionName, vbNullString, vbNullString, iniFile)
End Sub

Public Function fn_Readini(Section As String, Key As String, iniFile As String)
    Dim RetVal  As String
    Dim AppName As String
    Dim Worked  As Integer
    
    'AppName = App.Path & "\info.dat"
    AppName = iniFile
    
    RetVal = String(255, 0)
    Worked = GetPrivateProfileString(UCase(Section), UCase(Key), "", RetVal, Len(RetVal), AppName)
    
    fn_Readini = Replace(Left(RetVal, Worked), Chr(0), "")
End Function

Public Sub Writeini(Section As String, Key As String, iniFile As String, W_KEY As String)
    Dim Worked  As Integer
    Dim AppName As String
    
    'AppName = App.Path & "\setup.ini"
    AppName = iniFile
    
    'Worked = WritePrivateProfileString(UCase(Section), UCase(Key), W_KEY, AppName)
    Worked = WritePrivateProfileString(UCase(Section), Key, W_KEY, AppName)
End Sub


Public Function gf_TrimID(GetStr As String) As String
    Dim TempAsc  As Integer
    Dim Position As Integer
    
    Position = InStr(GetStr, Chr(0))
    If Position <> 0 Then
        gf_TrimID = Left(GetStr, Position - 1)
    Else
        gf_TrimID = GetStr
    End If
    gf_TrimID = Trim(gf_TrimID)
End Function

Public Function GetIPAddress() As String
    Dim sHostName   As String * 256
    Dim lpHost      As Long
    Dim Host        As HOSTENT
    Dim dwIPAddr    As Long
    Dim tmpIPAddr() As Byte
    Dim i           As Integer
    Dim sIPAddr     As String

    If Not SocketsInitialize() Then
        GetIPAddress = ""
        Exit Function
    End If

    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetIPAddress = ""
        SocketsCleanup
        Exit Function
    End If

    sHostName = Trim(sHostName)
    lpHost = gethostbyname(sHostName)

    If lpHost = 0 Then
        GetIPAddress = ""
        SocketsCleanup
        Exit Function
    End If

    CopyMemory Host, lpHost, Len(Host)
    CopyMemory dwIPAddr, Host.hAddrList, 4

    ReDim tmpIPAddr(1 To Host.hLen)
    CopyMemory tmpIPAddr(1), dwIPAddr, Host.hLen

    For i = 1 To Host.hLen
        sIPAddr = sIPAddr & tmpIPAddr(i) & "."
    Next

    GetIPAddress = MidH(sIPAddr, 1, Len(sIPAddr) - 1)

    SocketsCleanup
End Function

Public Function HiByte(ByVal wParam As Integer)
    HiByte = wParam \ &H100 And &HFF&
End Function

Public Function LoByte(ByVal wParam As Integer)
    LoByte = wParam And &HFF&
End Function

Public Sub SocketsCleanup()
    If WSACleanup() <> ERROR_SUCCESS Then
    End If
End Sub

Public Function SocketsInitialize() As Boolean
    Dim WSAD    As WSADATA
    Dim sLoByte As String
    Dim sHiByte As String

    If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
        SocketsInitialize = False
        Exit Function
    End If

    If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        SocketsInitialize = False
        Exit Function
    End If

    If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
        sHiByte = CStr(HiByte(WSAD.wVersion))
        sLoByte = CStr(LoByte(WSAD.wVersion))

        SocketsInitialize = False
        Exit Function
    End If

    SocketsInitialize = True
End Function

'----------------------------------------------------------------
' String에 Single Quote가 존재할 경우 Single Quote를 두개로 만듦.
'----------------------------------------------------------------
Public Function SubSQuotA(parStr As String) As String
    Dim lenStr As Long
    Dim i      As Long
    Dim bufStr As String

    SubSQuotA = parStr
    If InStr(parStr, "'") = 0 Then Exit Function
    lenStr = Len(parStr)
    
    bufStr = ""

    For i = 1 To lenStr
        bufStr = bufStr & IIf(Mid(parStr, i, 1) = "'", "''", Mid(parStr, i, 1))
    Next i
    
    SubSQuotA = bufStr
End Function

'*****************************************************************************
'프로그램명 : MidH
'기 능 : 한글을 2Byte로 처리하여 Mid함수로 처리한다.
'인 수 : sInStr As String Mid처리할 스트링
' iStart As Integer 시작위치
' iCnt As Integer Mid할 스트링 숫
'리 턴 값 : Mid된 결과 스트링
'사 용 예 : strTemp = gfMid("무궁화꽃이피었습니다",3,6)
' 결과 : 궁화꽃
'작 성 자 : 김경학
'작 성 일 : 2001.07.24
'수정 이력 :
'*****************************************************************************
Public Function MidH(sInStr As String, iStart As Integer, iCnt As Integer) As String
    MidH = StrConv(MidB(StrConv(sInStr, vbFromUnicode), iStart, iCnt), vbUnicode)
End Function

'*****************************************************************************
'프로그램명 : gfLen
'기 능 : 한글을 2Byte로 처리하여 Len함수로 처리한다.
'인 수 : strString As String Length를 구할 스트링
'리 턴 값 : 2바이트로 처리된 Length
'사 용 예 : intLen = gfLen("무궁화꽃이피었습니다")
' 결과 : 20
'작 성 자 : 김경학
'작 성 일 : 2001.07.24
'수정 이력 :
'*****************************************************************************
Public Function LenH(strString As String) As Integer
    LenH = LenB(StrConv(strString, vbFromUnicode))
End Function

Public Function FindFile(strSearchStartDir As String, strFileToBeSearchedFor As String) As String
    On Error GoTo FindFile_err: 'If error, just bail (no error handler yet)
    
    Dim lngCurrentDir As Long 'Index To directory in dir array we're examining
    Dim lngNumDirs As Long 'Number of directories to be searched (gets bigger
    Dim strTemp As String
    ReDim strDirs(0) As String 'List of directories to be searched (grows as search goes on)

    If Right(Trim(strSearchStartDir), 1) <> "\" Then
        strDirs(0) = Trim(strSearchStartDir) & "\"
    Else
        strDirs(0) = Trim(strSearchStartDir)
    End If
    
    lngCurrentDir = 0 'Dir list index
    lngNumDirs = 1 'Number of directories In dir list


    Do While lngCurrentDir < lngNumDirs
        If Dir(strDirs(lngCurrentDir) & strFileToBeSearchedFor, vbReadOnly + vbHidden) <> "" Then
            'we're done!
            FindFile = strDirs(lngCurrentDir) & strFileToBeSearchedFor
            Exit Function
        End If
        strTemp = Dir(strDirs(lngCurrentDir), vbDirectory + vbReadOnly + vbHidden)


        Do While strTemp <> ""
            If (GetAttr(strDirs(lngCurrentDir) & strTemp) And vbDirectory) = vbDirectory Then
                If strTemp <> "." And strTemp <> ".." Then
                    ReDim Preserve strDirs(0 To lngNumDirs) As String
                    strDirs(lngNumDirs) = strDirs(lngCurrentDir) & strTemp & "\"
                    lngNumDirs = lngNumDirs + 1
                End If
            End If
            strTemp = Dir 'Get Next matching directory
        Loop
        lngCurrentDir = lngCurrentDir + 1 'bump index To Get Next directory in dir array
    Loop
FindFile_exit:
    'if we're here, search failed
    FindFile = ""
    Exit Function
FindFile_err:
    'error handling would go here - for now,
    '     exit as a failed search
    Resume FindFile_exit
End Function


Public Function GetIniStr(SectionName As String, LineName As String, defValue As String, iniFile As String) As String
    Dim retStr As String * 256
    Dim result As Integer
    
    result = GetPrivateProfileString(SectionName, LineName, defValue, retStr, Len(retStr), iniFile)
    
    GetIniStr = Left(retStr, result)
End Function

Public Function SetIniStr(SectionName As String, LineName As String, defValue As String, iniFile As String) As String
    Dim result As Integer
    
    result = WritePrivateProfileString(SectionName, LineName, defValue, iniFile)
End Function

''---------------------------------------------------------------------------
'' 함수명 : Error_Msg
'' 기  능 :
'' 설  명 :
''---------------------------------------------------------------------------
'Public Sub Error_Msg(strEvent As String, strSource As String, strNumber As String, strDescription As String)
'    Dim Err_Msg As String
'
'              Err_Msg = "발생위치 : " & strEvent & vbNewLine & vbNewLine
'    Err_Msg = Err_Msg & "오류소스 : " & strSource & vbNewLine & vbNewLine
'    Err_Msg = Err_Msg & "오류번호 : " & strNumber & vbNewLine & vbNewLine
'    Err_Msg = Err_Msg & "오류내용 : " & strDescription
'
'    MsgBox Err_Msg, vbCritical, "오류"
'    Screen.MousePointer = 0
'End Sub


'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : ExecPro
' 작  성  자  : IT21
' 작  성  일  : 2000.06.07
' 파 라 미 터 : ProcName - 프로시저명
'               sValue   - 프로시저 파라미터
'               Err_Num  - 에러번호
'               Err_Dec  - 에러명
' 리  턴  값  : Recordset
' 비      고  : Server에 있는 스토어드 프로시저를 실행하기 위한 함수
'-----------------------------------------------------------------------------------------------------------------------------------------
Function ExecPro(ByVal ProcName As String, ByRef sValue() As String, Err_Num As Long, Err_Dec As String) As ADODB.Recordset
    Dim i As Integer
    Dim MyCom As ADODB.Command
    
    On Error GoTo ErrHandle

    Set ExecPro = New ADODB.Recordset
    Set MyCom = New ADODB.Command
    
    MyCom.ActiveConnection = ADONewServer
    MyCom.CommandTimeout = 0
    MyCom.CommandText = ProcName
    MyCom.CommandType = adCmdStoredProc
    
    For i = 1 To MyCom.Parameters.Count - 1
        If IsNull(sValue(i - 1)) Then
            MyCom.Parameters(i).Size = -1
        ElseIf sValue(i - 1) = "" Then
            MyCom.Parameters(i).Size = -1
        Else
            MyCom.Parameters(i).Size = LenH(sValue(i - 1))
        End If
        
        MyCom.Parameters(i) = sValue(i - 1)
    Next i
    
    Set ExecPro = MyCom.Execute
    Set MyCom = Nothing
    
    Err_Num = 0
    Err_Dec = ""
    
    Exit Function
    
ErrHandle:
    
    Err_Num = Err.Number
    Err_Dec = Err.Description
    
    Set MyCom = Nothing
End Function


Public Function ERR_SAVE(ByVal sDescription As String) As Boolean
    Dim FHandle As Integer
    Dim sText   As String

    If Dir(App.Path & "\Logs", vbDirectory) = "" Then MkDir "Logs"

    sText = Now & " : " & sDescription

    FHandle = FreeFile
    Open App.Path & "\Logs\" & Format(Date, "YYYY-MM-DD") & "_SendError.Txt" For Append As FHandle

    Print #FHandle, sText
    Close #FHandle
    Exit Function

End Function


Public Sub Error_Msg(strEvent As String, strSource As String, strNumber As String, strDescription As String)
    Dim FileNum As Integer
    Dim Err_Msg As String
        
'********************************************************************************************
    FileNum = FreeFile
    
    If Dir(AppPath & "ErrFile", vbDirectory) = "" Then
        MkDir AppPath & "ErrFile"
    End If

    Open AppPath & "ErrFile\" & Format(Date, "YYYYMMDD") & ".txt" For Append As #FileNum
        
    Print #FileNum, "발생시간 : " & Format(Now, "YYYY-MM-DD hh:mm:ss")
    Print #FileNum, "발생위치 : DBUpdate.exe " & strEvent
    Print #FileNum, "오류소스 : " & strSource
    Print #FileNum, "오류번호 : " & strNumber
    Print #FileNum, "오류내용 : " & strDescription
    Print #FileNum, "=============================================================="
    Close #FileNum
'********************************************************************************************

              Err_Msg = "발생위치 : " & strEvent & vbNewLine & vbNewLine
    Err_Msg = Err_Msg & "오류소스 : " & strSource & vbNewLine & vbNewLine
    Err_Msg = Err_Msg & "오류번호 : " & strNumber & vbNewLine & vbNewLine
    Err_Msg = Err_Msg & "오류내용 : " & strDescription
    
    'MsgBox Err_Msg, vbCritical, "오류"
    
    Screen.MousePointer = 0
End Sub

