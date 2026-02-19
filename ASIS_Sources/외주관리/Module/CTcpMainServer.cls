VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTcpMainServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'   S_STA       : 메시지의 시작을 의미한다.
'   CLEANAID    : 프로그램을 사용하고 있는 회사를 의미한다.
'   1001        : 프로그램을 사용하고 있는 회사중 각각의 코드를 의미한다.(지사등등)
'   OK_CREATE_CHULGO_DATA : 출고파일을 올바르게 생성했다.
'   DATA        : 생성한 파일 이름이 전달 된다.
'   S_END       : 메시지의 종료를 의미한다.


'************************************************************************
' Example:
'
'
'   Dim CCAid as CTcpMainServer
'   Set CCAid = New CTcpMainServer
'   With CCAid
'     .WinsockCtrl=Winsock1
'     .WakeUp
'   End With
'   Set CCAid=Nothing
'
'************************************************************************

'Object-level constants and enums...
Private Const ERROR_BASE = vbObjectError + 2000
Private Const ERROR_INVALID_WINSOCK_REFERENCE = ERROR_BASE + 1
Private Const ERROR_NO_WINSOCK_REFERENCE_AVAILABLE = ERROR_BASE + 2
Private Const ERROR_INVALID_IP_ADDRESS = ERROR_BASE + 3
Private Const ERROR_INVALID_PORT = ERROR_BASE + 4

'Object-level variables...
Private m_WinsockControl As Control
Private m_IPAddress As String
Private m_Port As String

Private m_StoreOffice As String
Private m_StoreCode As String
Private m_StoreName As String


'************************************************************************
' Property WinsockControl()
'
' 해당 작업을 수행한후 메시지를 전달한 객체를 설정한다.
'
'************************************************************************
Public Property Let WinsockControl(ByRef RefWinsock As Control)
  
  'Check for uninitialized controls...
  If RefWinsock Is Nothing Then
    Err.Raise ERROR_NO_WINSOCK_REFERENCE_AVAILABLE, _
              "CMagicPacket.WinsockControl(LET)", _
              "An attempt was made to set the WinsockControl property " & _
              "to an uninitialized control"
  End If
  
'  'Check for invalid control types...
'  If TypeOf RefWinsock Is Winsock Then
'    Set m_WinsockControl = RefWinsock
'  Else
'    Err.Raise ERROR_INVALID_WINSOCK_REFERENCE, _
'              "CMagicPacket.WinsockControl(LET)", _
'              "An attempt was made to set the WinsockControl property " & _
'              "to a non-winsock control"
'  End If
  
  'It passed the test, so set our reference to it...
  Set m_WinsockControl = RefWinsock
  
End Property

'************************************************************************
' Property SetIPAddress()
'
' 연결할 아이피를 설정한다.
'
'************************************************************************
Public Property Let SetIPAddress(ByVal NewAddress As String)
  
    'Set the new IP subnet address...
'    If UCase(NewAddress) <> "LOCALHOST" Then
'        m_IPAddress = ValidateIPAddress(NewAddress)
'    Else
'        m_IPAddress = NewAddress
'    End If
  
    m_IPAddress = NewAddress
  
    'And raise an error if it is not valid...
    If m_IPAddress = "" Then
        Err.Raise ERROR_INVALID_IP_ADDRESS, _
              "CTcpMainServer.SetIPAddress(LET)", _
              "Invalid IP Address"
    End If
  
End Property

Public Property Get GetIPAddress() As String
  GetIPAddress = m_IPAddress
End Property


'************************************************************************
' Property SetIPAddress()
'
' 연결할 아이피를 설정한다.
'
'************************************************************************
Public Property Let SetPort(ByVal NewPort As String)
  
  'Set the new IP subnet address...
  m_Port = NewPort
  
  'And raise an error if it is not valid...
  If NewPort = "" Then
    Err.Raise ERROR_INVALID_PORT, _
              "CTcpMainServer.SetPort(LET)", _
              "Invalid Port"
  End If
  
End Property

Public Property Get GetPort() As String
  GetPort = m_Port
End Property


'************************************************************************
' Property SetStoreCode()
'
' 본사/지점의 코드 설정
'
'************************************************************************
Public Property Let SetStoreOffice(ByVal NewName As String)
  
  m_StoreOffice = NewName
  
  '오류 메시지
  If NewName = "" Then
    Err.Raise ERROR_INVALID_PORT, _
              "CTcpMainServer.SetStoreOffice", _
              "Invalid Store Office"
  End If
  
End Property

Public Property Get GetStoreOffice() As String
  GetStoreOffice = m_StoreOffice
End Property


'************************************************************************
' Property SetStoreCode()
'
' 본사/지점의 코드 설정
'
'************************************************************************
Public Property Let SetStoreCode(ByVal NewCode As String)
  
  m_StoreCode = NewCode
  
  '오류 메시지
  If NewCode = "" Then
    Err.Raise ERROR_INVALID_PORT, _
              "CTcpMainServer.SetStoreCode", _
              "Invalid Store Code"
  End If
  
End Property

Public Property Get GetStoreCode() As String
  GetStoreCode = m_StoreCode
End Property


'************************************************************************
' Property SetStoreName()
'
' 본사/지점의 코드 설정
'
'************************************************************************
Public Property Let SetStoreName(ByVal NewName As String)
  
  m_StoreName = NewName
  
  '오류 메시지
  If NewName = "" Then
    
    Err.Raise ERROR_INVALID_PORT, _
              "CTcpMainServer.SetStoreName", _
              "환경설정에서 회사명칭을 확인하여 주십시요."
  End If
  
End Property

Public Property Get GetStoreName() As String
  GetStoreName = m_StoreName
End Property


Public Property Get Stats() As Integer
  Stats = m_WinsockControl.State
End Property


'************************************************************************
' ValidateIPAddress()
'
' This function takes an IP address in dotted decimal format and verifies
' that there are 4 octets of data and that the value of each octet is
' between zero and 255.  This function does NOT determine whether the
' given address is legal for any particular subnetting scheme. If the
' address is valid, it will return the address itself, otherwise it will
' return an empty (zero-length) string.
'
'************************************************************************
Private Function ValidateIPAddress(AddressToValidate As String) As String

'Local Variables...
Dim TempStr As String
Dim ReducedAddress As String
Dim SplitAddress() As String
Dim OctetIndex As Long
Dim OctetValue As Long
  
  'First remove all non-decimal, non-period characters from the address...
  For OctetIndex = 1 To Len(AddressToValidate)
    TempStr = Mid$(AddressToValidate, OctetIndex, 1)
    Select Case TempStr
      Case "0" To "9", ".": ReducedAddress = ReducedAddress & TempStr
    End Select
  Next
  
  'Split the octets into a string array for inspection...
  SplitAddress = Split(ReducedAddress, ".")
  ReducedAddress = ""
  
  'If there are not exactly 4 octets (0->3), return empty...
  If UBound(SplitAddress()) <> 3 Then
    ValidateIPAddress = ""
    Exit Function
  End If

  'Verify that the value of each octet is between 0 and 255...
  For OctetIndex = 0 To 3
    TempStr = SplitAddress(OctetIndex)
    OctetValue = Val(TempStr)
    If Len(TempStr) = 0 Or OctetValue < 0 Or OctetValue > 255 Then
      ValidateIPAddress = ""
      Exit Function
    End If
    ReducedAddress = ReducedAddress & Format$(OctetValue, "##0")
    If OctetIndex < 3 Then ReducedAddress = ReducedAddress & "."
  Next
  
  'Return the validated result...
  ValidateIPAddress = ReducedAddress
  
End Function

'************************************************************************
' Property Send_RecvFileList_ALL()
'
' 해당 파일에 해당 작업을 완료한 경우 본사 관리 프로그램에 완료한 내용을
' 전달 해준다.
'
'************************************************************************
Public Sub Send_RecvFileNameAction(FileName As String, sMode As String)
'   S_STA       : 메시지의 시작을 의미한다.
'   CLEANAID    : 프로그램을 사용하고 있는 회사를 의미한다.
'   1001        : 프로그램을 사용하고 있는 회사중 각각의 코드를 의미한다.(지사등등)
'   FILELISTALL : 해당 프로그래에서 실행할 명령어가 전달된다.
'   자료        : 위 명령에서서 추가 정보를 전달할 경우 사용
'   S_END       : 메시지의 종료를 의미한다.
    Dim sSendData   As String
    Dim sGubun      As String
    
    sGubun = ";"
    sSendData = S_STA & "|" & "CLEANAID" & "|" & "0000" & "|" & _
                "RECEIVE_FILENAME_ACTION" & "|" & _
                FileName & sGubun & sMode & "|" & _
                S_END

    Debug.Print sSendData
    If m_WinsockControl.State = sckConnected Then m_WinsockControl.SendData CStr(sSendData)
    DoEvents

End Sub


Public Function Fnc_TcpCheckDataArrival(work As String) As Boolean
'수신된 데이터를 확인한다.
    Dim varValue    As Variant
    
    Fnc_TcpCheckDataArrival = False
    varValue = Split(work, "|")
    If UBound(varValue) < 2 Then
        ' 최소 3개보다는 많아야 한다.
        ' 잘못 수신된 데이타를 저장한다.
        Debug.Print Now & " ERROR Fnc_TcpCheckDataArrival < 3 => " & work
        Beep
        Exit Function
    End If
    
    '수신된 처음과 마지막을 확인한다.
    If CStr(varValue(0)) <> S_STA Then
        Debug.Print Now & "ERROR Fnc_TcpCheckDataArrival <> S_STA => " & work
        Beep
        Exit Function
        
    ElseIf CStr(varValue(UBound(varValue))) <> S_END Then
        Debug.Print Now & "ERROR Fnc_TcpCheckDataArrival <> S_END => " & work
        Beep
        Exit Function
    End If
    
    Debug.Print Now & " Fnc_TcpCheckDataArrival => " & work
    Fnc_TcpCheckDataArrival = True

End Function

Public Function TcpConnect() As Boolean
    
    TcpConnect = False
    ' 소켓상태가 연결인 경우는 재처리 안함
    If m_WinsockControl.State <> sckConnected Then
    
        Call PanelsMsg("본사와 연결중 입니다. 잠시만 기다려 주십시요.")
        
        ' 클라이언트 소켓을 종료시킴
        m_WinsockControl.Close
    
        ' 클라이언트 소켓이 종료할 때까지 기다림
        Do While m_WinsockControl.State <> sckClosed
            DoEvents
        Loop
    
        ' 클라이언트에서 서버에 연결을 시도함
        m_WinsockControl.RemoteHost = m_IPAddress
        m_WinsockControl.RemotePort = Val(m_Port)
        m_WinsockControl.Connect
    
        ' 클라이언트에서 서버에 연결이 완료할 때까지 기다림
        Do While m_WinsockControl.State <> sckConnected
            DoEvents
            If m_WinsockControl.State = sckError Then
                Call PanelsMsg("본사와 연결되지 않았습니다.")
                Exit Function
            End If
        Loop
    
    End If
    
    Call PanelsMsg("본사와 연결 되었습니다...")
    TcpConnect = True

End Function

' 출고 자료 생성
Public Sub CreateChulGoData(sDate As String, sCode As String)
    Dim sSendData   As String
    
    If sDate = "" Or sCode = "" Then
        Call PanelsMsg("출고 자료 생성에 전달된 값이 올바르지 않습니다.")
        Exit Sub
    End If
        
    sSendData = S_STA & "|" & Store.Office & "|" & _
                       Store.Code & ";" & Store.Name & "|" & _
                       "CREATE_CHLGO_DATA" & "|" & _
                       sDate & ";" & sCode & "|" & _
                       S_END
    
    Debug.Print sSendData
    If m_WinsockControl.State = sckConnected Then m_WinsockControl.SendData CStr(sSendData)
    DoEvents
        
        
End Sub


Public Sub Send_RecvFileListAll()
    Dim sMSG    As String
    
    sMSG = S_STA & "|" & Store.Office & "|" & _
                       Store.Code & ";" & Store.Name & "|" & _
                       "RECEIVE_FILELIST_ALL" & "|" & _
                       "ALL" & "|" & _
                       S_END
    
    Debug.Print sMSG
    If m_WinsockControl.State = sckConnected Then m_WinsockControl.SendData CStr(sMSG)
    DoEvents

End Sub

'  체인점에서 수신된 폴더에 있는 파일을 모두 해당 작업을 시행한다.
Public Sub Send_RecvFileAllAction()
    Dim sMSG    As String
    
    sMSG = S_STA & "|" & Store.Office & "|" & _
                   Store.Code & ";" & Store.Name & "|" & _
                   "RECEIVE_FILE_ALL_ACTION" & "|" & _
                   "ALL" & "|" & _
                   S_END
    
    Debug.Print sMSG
    If m_WinsockControl.State = sckConnected Then m_WinsockControl.SendData CStr(sMSG)
    DoEvents

End Sub
