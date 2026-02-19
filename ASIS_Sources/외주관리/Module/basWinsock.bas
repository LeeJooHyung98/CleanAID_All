Attribute VB_Name = "basWinsock"
Option Explicit

Public CCAid   As New CTcpMainServer


' 관리프로그램에서 서버로 연결하여 자료를 요청할때 주고 받을 메시지를 정의
Public Enum SendMsgMode
    RECEIVE_FILELIST_ALL    ' 체인점에서 수신된 폴더에 있는 모든 파일리스트를 요청한다.
    RECEIVE_FILE_ALL_ACTION ' 체인점에서 수신된 폴더에 있는 파일을 모두 해당 작업을 시행한다.
End Enum


' 서버에서 관리 프로그램으로 전달할 메시지를 정의
Public Enum RecvMsgMode
    RECEIVE_FILELIST_ALL    '체인점에서 수신된 폴더에 있는 모든 파일리스트를 요청한 자료를 보넨다.
    RECEIVE_FILENAME_ACTION ' 작업을 처리한 화일명을 받는다.
End Enum

Public Const S_STA As String = "OK_START"                   ' 전송 시작 시점
Public Const S_END As String = "OK_END"                     ' 전송 종료 시점


Public Function Fnc_TcpSendCreateData(sMode As SendMsgMode) As String
    Dim sMSG    As String

    Select Case UCase(sMode)
        
        ' 체인점에서 서버로 전송한 파일 리스트를 요구한다.
        Case SendMsgMode.RECEIVE_FILELIST_ALL
            sMSG = S_STA & "|" & Store.Office & "|" & _
                   Store.Code & "|" & _
                   Store.Name & "|" & _
                   "RECEIVE_FILELIST_ALL" & "|" & _
                   S_END
        
        '  체인점에서 수신된 폴더에 있는 파일을 모두 해당 작업을 시행한다.
        Case SendMsgMode.RECEIVE_FILE_ALL_ACTION
            sMSG = S_STA & "|" & Store.Office & "|" & _
                   Store.Code & "|" & _
                   Store.Name & "|" & _
                   "RECEIVE_FILE_ALL_ACTION" & "|" & _
                   S_END
        
        Case Else
            sMSG = ""
        
    End Select
    
    Fnc_TcpSendCreateData = sMSG

End Function

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

Public Function Fnc_TcpConnect(tcpHost As Winsock) As Boolean
    Dim strSendData As String
    
    Fnc_TcpConnect = False
    ' 소켓상태가 연결인 경우는 재처리 안함
    If tcpHost.State <> sckConnected Then
    
        Call PanelsMsg("본사와 연결중 입니다. 잠시만 기다려 주십시요.")
        
        ' 클라이언트 소켓을 종료시킴
        tcpHost.Close
    
        ' 클라이언트 소켓이 종료할 때까지 기다림
        Do While tcpHost.State <> sckClosed
            DoEvents
        Loop
    
        ' 클라이언트에서 서버에 연결을 시도함
        tcpHost.RemoteHost = Trim(GetIniStr("Store Server", "ServerNameOrIP", "", sIniFile))
        tcpHost.RemotePort = Val(Trim(GetIniStr("Store Server", "MessagePort", "", sIniFile)))
        tcpHost.Connect
    
        ' 클라이언트에서 서버에 연결이 완료할 때까지 기다림
        Do While tcpHost.State <> sckConnected
            DoEvents
            If tcpHost.State = sckError Then
                Call PanelsMsg("본사와 연결되지 않았습니다.")
                MsgBox "본사와 연결되지 않았습니다." & Space(10), vbInformation, "확인"
                Exit Function
            End If
        Loop
    
    End If
    
    Call PanelsMsg("본사와 연결 되었습니다...")
    Fnc_TcpConnect = True

End Function


Public Function Fnc_TcpSendMessage(tcpHost As Winsock, sSendMessage As String) As Boolean
    
    Fnc_TcpSendMessage = False
    
    ' 소켓상태가 연결인 경우는 재처리 안함
    If tcpHost.State <> sckConnected Then
        Call PanelsMsg("연결 오류 [Fnc_TcpSendMessage]")
        Exit Function
    End If
    
    ' 데이타를 서버에 보냄
    If tcpHost.State = sckConnected Then
        tcpHost.SendData sSendMessage
        DoEvents
        Call PanelsMsg("보넨 메시지 : " & sSendMessage)

    End If
    Fnc_TcpSendMessage = True

End Function

