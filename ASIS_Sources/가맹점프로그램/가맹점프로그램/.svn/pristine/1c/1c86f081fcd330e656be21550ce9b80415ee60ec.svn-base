Attribute VB_Name = "basMyIP"
'====================================================================================================
' 자신의 IP를 알아오기
'
Public Const WS_VERSION_REQD As Long = &H101
Public Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&
Public Const SOCKET_ERROR As Long = -1
Public Const WSADESCRIPTION_LEN = 257
Public Const WSASYS_STATUS_LEN = 129
Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128
Public Const MIN_SOCKETS_REQD As Long = 1

Public Type WSAData
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To MAX_WSADescription) As Byte
    szSystemStatus(0 To MAX_WSASYSStatus) As Byte
    wMaxSockets As Integer
    wMaxUDPDG As Integer
    dwVendorInfo As Long
End Type
Public Type WSADataInfo
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSADESCRIPTION_LEN
    szSystemStatus As String * WSASYS_STATUS_LEN
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As String
End Type
Public Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    HLen As Integer
    hAddrList As Long
End Type

Public Declare Function WSAStartupInfo Lib "WSOCK32" Alias "WSAStartup" (ByVal wVersionRequested As Integer, lpWSADATA As WSADataInfo) As Long
Public Declare Function WSACleanup Lib "WSOCK32" () As Long
Public Declare Function WSAGetLastError Lib "WSOCK32" () As Long
Public Declare Function WSAStartup Lib "WSOCK32" (ByVal wVersionRequired As Long, lpWSADATA As WSAData) As Long
Public Declare Function gethostname Lib "WSOCK32" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Public Declare Function gethostbyname Lib "WSOCK32" (ByVal szHost As String) As Long
Public Declare Sub CopyMemoryIP Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

'====================================================================================================
' 작   성   자 : pds2004 박대선
' 작 성  일 자 : 2003.04.26
' 최종 수정 자 :
' 최종수정일자 :
' 사용 API함수 :
'----------------------------------------------------------------------------------------------------
'   자신의 IP를 얻어온다.
'====================================================================================================
Public Function GetIPAddress() As String
    Dim sHostName As String * 256
    Dim lpHost As Long
    Dim HOST As HOSTENT
    Dim dwIPAddr As Long
    Dim tmpIPAddr() As Byte
    Dim sIPAddr As String
    
    If Not SocketsInitialize() Then
        GetIPAddress = ""
        Exit Function
    End If
    
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetIPAddress = ""
        MsgBox "Windows Sockets error " & Str(WSAGetLastError()) & _
                            " has occurred. Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    
    sHostName = Trim(sHostName)
    lpHost = gethostbyname(sHostName)
    
    If lpHost = 0 Then
        GetIPAddress = ""
        MsgBox "Windows Sockets are not responding. " & "Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    
    CopyMemoryIP HOST, lpHost, Len(HOST)
    CopyMemoryIP dwIPAddr, HOST.hAddrList, 4
    ReDim tmpIPAddr(1 To HOST.HLen)
    CopyMemoryIP tmpIPAddr(1), dwIPAddr, HOST.HLen
    For i = 1 To HOST.HLen
        sIPAddr = sIPAddr & tmpIPAddr(i) & "."
    Next
    GetIPAddress = Mid(sIPAddr, 1, Len(sIPAddr) - 1)
    SocketsCleanup
End Function


Public Function GetIPHostName() As String
'====================================================================================================
' 작   성   자 : pds2004 박대선
' 작 성  일 자 : 2003.04.26
' 최종 수정 자 :
' 최종수정일자 :
' 사용 API함수 : SHFileOperation
'----------------------------------------------------------------------------------------------------
'   자신의 IP를 얻어온다.
'====================================================================================================
    Dim sHostName As String * 256
    If Not SocketsInitialize() Then
        GetIPHostName = ""
        Exit Function
    End If
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetIPHostName = ""
        MsgBox "Windows Sockets error " & Str(WSAGetLastError()) & _
                " has occurred. Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    GetIPHostName = Left(sHostName, InStr(sHostName, Chr(0)) - 1)
    SocketsCleanup
End Function

Public Function HiByte(ByVal wParam As Integer)
'====================================================================================================
'   자신의 IP를 얻어오기 위한 함수
'====================================================================================================
    HiByte = wParam \ &H100 And &HFF&
End Function

Public Function LoByte(ByVal wParam As Integer)
'====================================================================================================
'   자신의 IP를 얻어오기 위한 함수
'====================================================================================================
    LoByte = wParam And &HFF&
End Function

Public Sub SocketsCleanup()
'====================================================================================================
'   자신의 IP를 얻어오기 위한 함수
'====================================================================================================
    If WSACleanup() <> 0 Then
        MsgBox "Socket error occurred in Cleanup."
    End If
End Sub

Public Function SocketsInitialize() As Boolean
'====================================================================================================
'   자신의 IP를 얻어오기 위한 함수
'====================================================================================================
    Dim WSAD As WSAData
    Dim sLoByte As String
    Dim sHiByte As String
    If WSAStartup(WS_VERSION_REQD, WSAD) <> 0 Then
        MsgBox "The 32-bit Windows Socket is not responding."
        SocketsInitialize = False
        Exit Function
    End If
    If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        MsgBox "This application requires a minimum of " & _
                CStr(MIN_SOCKETS_REQD) & " supported sockets."
        SocketsInitialize = False
        Exit Function
    End If
    If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or (LoByte(WSAD.wVersion) = _
                            WS_VERSION_MAJOR And HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
        sHiByte = CStr(HiByte(WSAD.wVersion))
        sLoByte = CStr(LoByte(WSAD.wVersion))
        MsgBox "Sockets version " & sLoByte & "." & sHiByte & _
                    " is not supported by 32-bit Windows Sockets."
        SocketsInitialize = False
        Exit Function
    End If
    'must be OK, so lets do it
    SocketsInitialize = True
End Function

'====================================================================================================
' 작   성   자 : pds2004 - 박대선
' 작 성  일 자 : 2003.05.23 -
' 최종 수정 자 :
' 최종수정일자 :
' 사용 API함수 :
' 리   턴   값 : 접속을 기다릴 포트
'----------------------------------------------------------------------------------------------------
'
'====================================================================================================
Public Function Fn_GetRemoteIP() As String
    Dim strValue As String
    
    strValue = GetIniStr("Connect", "RemoteIP", "", iniFile)
    
    If Len(strValue) <= 0 Then
        strValue = "web.clean-aid.co.kr"
        Call SetIniStr("Connect", "RemoteIP", strValue, iniFile)
    End If
    
    Fn_GetRemoteIP = strValue
End Function

'====================================================================================================
' 작   성   자 : pds2004 - 박대선
' 작 성  일 자 : 2003.05.23 -
' 최종 수정 자 :
' 최종수정일자 :
' 사용 API함수 :
' 리   턴   값 : 접속을 기다릴 포트
'----------------------------------------------------------------------------------------------------
'
'====================================================================================================
Public Function Fn_GetMsgRemotePort() As String
    Dim strValue As String
    
    strValue = GetIniStr("Connect", "MsgRemotePort", "", iniFile)
    
    If Len(strValue) <= 0 Then
        strValue = "8607"
        Call SetIniStr("Connect", "MsgRemotePort", strValue, iniFile)
    End If
    
    Fn_GetMsgRemotePort = strValue
End Function

'====================================================================================================
' 작   성   자 : pds2004 - 박대선
' 작 성  일 자 : 2003.05.23 -
' 최종 수정 자 :
' 최종수정일자 :
' 사용 API함수 :
' 리   턴   값 : 접속을 기다릴 포트
'----------------------------------------------------------------------------------------------------
'
'====================================================================================================
Public Function Fn_GetFileLocatPort() As String
    Dim strValue As String
    
    strValue = GetIniStr("Connect", "FileLocalPort", "", iniFile)
    
    If Len(strValue) <= 0 Then
        strValue = "8629"
        
        Call SetIniStr("Connect", "FileLocalPort", strValue, iniFile)
    End If
    
    Fn_GetFileLocatPort = strValue
End Function

'====================================================================================================
' 작   성   자 : pds2004 - 박대선
' 작 성  일 자 : 2003.05.23 -
' 최종 수정 자 :
' 최종수정일자 :
' 사용 API함수 :
' 리   턴   값 : 접속을 기다릴 포트
'----------------------------------------------------------------------------------------------------
'
'====================================================================================================
Public Function Fn_GetFileRemotePort() As String
    Dim strValue As String
    
    strValue = GetIniStr("Connect", "FileRemotePort", "", iniFile)
    
    If Len(strValue) <= 0 Then
        strValue = "8627"
        
        Call SetIniStr("Connect", "FileRemotePort", strValue, iniFile)
    End If
    
    Fn_GetFileRemotePort = strValue
    
End Function

