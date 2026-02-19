VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl ctlFileTransfer 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ctlFileTransfer.ctx":0000
   Begin VB.Timer tmrTimeout 
      Enabled         =   0   'False
      Left            =   1020
      Top             =   0
   End
   Begin MSWinsockLib.Winsock tcpSocket 
      Left            =   540
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "localhost"
      RemotePort      =   3279
      LocalPort       =   3279
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  '투명
      Caption         =   "2.x"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   60
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
   Begin VB.Image imgControl 
      Height          =   480
      Left            =   0
      Picture         =   "ctlFileTransfer.ctx":0312
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "ctlFileTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#################################################################
'##                 File Transfer Control 2.1                   ##
'##                                                             ##
'## 만든이: 이창연(lisyoen@lisyoen.com)                         ##
'## 문  의: http://www.lisyoen.com, http://www.exproject.com    ##
'#################################################################
'
'본 프로그램의 저작권은 이창연 본인에게 있습니다. 이건 말뿐이고...
'소스와 컴파일된 컨트롤은 상업적이든 비상업적이든 자유롭게 사용하십시오.
'수정 및 기능첨삭에 대한 의견은 메일로 보내주시기 바랍니다.
'물론 직접 손봐서 사용하셔도 됩니다.
'
'단, 충분한 설명서와 예제를 함께 배포해 주시기 바랍니다.
'

Option Explicit
'기본 속성 값:
Const m_def_Version = 0
Const m_def_ReceiveFilePath = ""
Const m_def_PayloadSize = 8169
Const m_def_RemotePort = 3279
Const m_def_LocalPort = 3279
Const m_def_RemoteHost = "localhost"
Const m_def_SendFileSize = 0
Const m_def_State = 0
Const m_def_Cps = 0
Const m_def_SendFilePath = ""
Const m_def_ReceiveDirPath = ""
Const m_def_ReceiveFileSize = 0
Const m_def_EnableReceive = True
Const m_def_ReceiveFileName = ""

'속성 변수:
Dim m_Version As Single
Dim m_ReceiveFilePath As String
Dim m_PayloadSize As Long
Dim m_RemotePort As Long
Dim m_LocalPort As Long
Dim m_RemoteHost As String
Dim m_SendFileSize As Long
Dim m_State As TransferState
Dim m_Cps As Long
Dim m_SendFilePath As String
Dim m_ReceiveDirPath As String
Dim m_ReceiveFileSize As Long
Dim m_EnableReceive As Boolean
Dim m_ReceiveFileName As String

'이벤트 선언:
'전송 중지 이벤트, FileSize: 최종 전송 파일 크기
Event TransferCut(ByVal FileSize As Long)
'상태 변경 이벤트, NewState: 변경된 상태
Event ChangeState(ByVal NewState As TransferState)
'수신 시작 이벤트
'Overwrite = True   '파일이 존재하면 덮어쓰기
'Overwrite = False  '파일이 존재하면 이어받기(기본값)
'Cancel = True  '전송 거부
'Cancel = False '전송 수락(기본값)
Event ReceiveStart(ByRef Filename As String, ByVal FileSize As Long, ByRef Overwrite As Boolean, ByRef Cancel As Boolean)
'송신 과정 이벤트
Event SendProgress(ByVal SendSize As Long)
'송신 완료 이벤트
Event SendComplete(ByVal FileSize As Long)
'수신 완료 이벤트
Event ReceiveComplete(ByVal FileSize As Long)
'수신 과정 이벤트
Event ReceiveProgress(ByVal ReceiveSize As Long)
'예외 처리 이벤트
Event Error(ByVal Number As Long, Description As String)

'로컬 변수
Dim p_FileNum As Integer       '전송 파일 번호
Dim p_StartTime As Single   '전송 시작 시간
Dim Buffer() As Byte        '전송 버퍼
Dim p_NowTime As Single     '현재 시간 타이머
Dim p_FileOffset As Long    '전송 파일 오프셋
Dim p_SendComplete As Boolean   '송신 확인
Dim p_ReceiveComplete As Boolean    '수신 확인
Dim p_SendFileName As String    '송신 파일 이름

'참고 사항
'Ambient.UserMode = 런타임
'Err.Raise 387  ' - 디자인타임 오류
'Err.Raise 382  ' - 런타임 오류

'Error 이벤트 번호 및 발생 시점, 설명
'62101 - SendFile 호출 후; SendTimer Timeout. 수신측에서 응답이 없습니다.
'62102 - SendFile 호출 후; 잘못된 파일 오프셋 값이 수신되었습니다.
'62103 - 파일 수신중; 수신된 파일의 크기가 맞지 않습니다.
'62104 - 파일 송신중: 파일 송신중 원격 호스트로부터 전송이 중단되었습니다.
'62105 - 파일 수신중: 파일 수신중 원격 호스트로부터 전송이 중단되었습니다.
'% tcpSocket 의 Error 이벤트가 발생하면 같은 내용으로 컨트롤의 Error 이벤트도 발생

'런타임 오류 번호 및 발생 위치, 설명
'62001 - SendFilePath; 유효한 파일 경로가 아닙니다.
'62002 - ReceiveDirPath; 유효한 경로가 아닙니다.
'62003 - RemotePort; Ready 상태에서만 속성을 변경할 수 있습니다. * 삭제됨
'62004 - LocalPort; Ready 상태에서만 속성을 변경할 수 있습니다.  * 삭제됨
'62005 - RemoteHost; Ready 상태에서만 속성을 변경할 수 있습니다. * 삭제됨
'62006 - RemotePort; 포트 번호의 범위가 잘못 지정되었습니다. (0~65535)
'62007 - LocalPort; 포트 번호의 범위가 잘못 지정되었습니다. (0~65535)
'62008 - RemoteHost; 호스트 값 설정이 잘못되었습니다.
'62009 - SendFile; 파일 전송중에는 SendFile 메소드를 호출할 수 없습니다.
'62010 - SendFile; TimeoutSec 설정 범위가 잘못 지정되었습니다. (0~60초)

'컨트롤 상태 열거형
Public Enum TransferState
    ftcReady = 0    '접속 대기 상태
    ftcSendReady = 1 '송신 대기 상태(상대방으로부터 응답을 기다림)
    ftcSend = 2     '송신중
    ftcReceiveReady = 3 '수신 대기 상태(수신할 것인지를 결정하는 단계)
    ftcReceive = 4  '수신중
End Enum

Private Sub tcpSocket_Close()
    '송신중이라면
    If p_SendComplete = False Then
        RaiseEvent TransferCut(Loc(p_FileNum))
        RaiseEvent Error(62104, "파일 송신중 원격 호스트로부터 전송이 중단되었습니다.")
    End If
    '수신중이라면
    If p_ReceiveComplete = False Then
        RaiseEvent TransferCut(Loc(p_FileNum))
        RaiseEvent Error(62105, "파일 수신중 원격 호스트로부터 전송이 중단되었습니다.")
    End If
    '전송 중지
    Cut
    
    'pds2004가 추가감
    
    '소켓을 닫은 후
    tcpSocket.Close

    '수신 가능이면 대기 상태로 전환
    tcpSocket.LocalPort = m_LocalPort
    If m_EnableReceive Then tcpSocket.Listen
    
End Sub

Private Sub tcpSocket_Connect()

    ' RemodeConnect에서 호출하여 연결된 상태에서는 기다린다.
    If m_State = ftcReady Then Exit Sub
    If Trim(m_SendFilePath) = "" Then Exit Sub
    
    
    '헤더 전송
    '헤더크기는 4 + 128 = 132 bytes
    tcpSocket.SendData m_SendFileSize
    tcpSocket.SendData AscLeft(p_SendFileName + Space(128), 128)
    'Timeout 타이머 온
    tmrTimeout.Enabled = True
End Sub

Private Sub tcpSocket_ConnectionRequest(ByVal requestID As Long)
    '연결을 위해 소켓을 닫고
    tcpSocket.Close
    '소켓 연결
    tcpSocket.Accept requestID
    '상태 변경
    m_State = ftcReceiveReady
    RaiseEvent ChangeState(m_State)
End Sub

Private Sub tcpSocket_DataArrival(ByVal bytesTotal As Long)
    Select Case m_State
        Case TransferState.ftcReceive   '수신중이라면
            '파일에 기록
            tcpSocket.GetData Buffer, vbByte + vbArray
            On Error Resume Next    '파일 오류 처리
            Put p_FileNum, , Buffer
            If Err Then '파일 오류가 발생하면
                '전송 중지
                Cut
                On Error GoTo 0 '오류 처리 중지
                '오류 발생
                Err.Raise Err.Number, , "파일 오류:" & Err.Description
            End If
            On Error GoTo 0 '오류 처리 중지
            '전송 과정 이벤트 발생
            RaiseEvent ReceiveProgress(Loc(p_FileNum))
            '날짜변경을 체크하고 Cps 계산
            p_NowTime = Timer
            If p_NowTime < p_StartTime Then    '전송중에 날짜가 바뀌었다면
                p_StartTime = p_StartTime + 86400
            End If
            m_Cps = Loc(p_FileNum) \ (p_NowTime - p_StartTime + 1)
            '파일 수신이 완료되었다면
            If Loc(p_FileNum) = m_ReceiveFileSize Then
                '수신 완료 이벤트를 발생시키고
                RaiseEvent ReceiveComplete(Loc(p_FileNum))
                '수신 확인
                p_ReceiveComplete = True
                '전송 중지
                Cut
            ElseIf Loc(p_FileNum) > m_ReceiveFileSize Then  '혹시 전송이 초과되면 (그럴리 없다)
                'Error 이벤트 발생
                RaiseEvent Error(62103, "수신된 파일의 크기가 일치하지 않습니다." & vbCrLf & _
                    "원래 파일 크기:" & CStr(m_ReceiveFileSize) & vbCrLf & _
                    "수신된 파일 크기:" & CStr(Loc(p_FileNum)))
                '수신 확인
                p_ReceiveComplete = True
                '전송 중지
                Cut
            End If
            
        Case TransferState.ftcReceiveReady  '수신 대기중이라면
            '헤더가 충분하다면(132 bytes)
            If bytesTotal >= 132 Then
                '헤더를 입력받고(파일 사이즈 4 bytes, 파일 이름 128 bytes)
                tcpSocket.GetData m_ReceiveFileSize, vbLong
                tcpSocket.GetData m_ReceiveFileName, vbString, 128
                '파일 이름과 전체 경로를 구한다.
                m_ReceiveFileName = Trim(m_ReceiveFileName)
                '덮어씌우기, 전송 취소 값
                Dim b_Overwrite As Boolean, b_Cancel As Boolean
                '기본 값은 이어받기
                b_Overwrite = False
                '취소 안함
                b_Cancel = False
                '수신 방식을 결정하기 위해
                '수신 시작 이벤트 발생(이벤트 핸들러 내에서 파일명을 변경할 수 있다)
                RaiseEvent ReceiveStart(m_ReceiveFileName, m_ReceiveFileSize, b_Overwrite, b_Cancel)
                'Overwrite = True 이면 덮어씌우기
                'Overwrite = False 이면 이어받기
                '취소라면
                If b_Cancel = True Then
                    '전송 중지
                    Cut
                    '작업 종료
                    Exit Sub
                End If
                '파일 전체 경로를 구한다.
                m_ReceiveFilePath = m_ReceiveDirPath & "\" & m_ReceiveFileName
                '파일을 연다
                p_FileNum = FreeFile
                '변경된 파일명 때문에 오류가 발생할 수 있으므로 오류 처리
                On Error Resume Next
                '파일 열기
                Open m_ReceiveFilePath For Binary Access Write As p_FileNum
                If b_Overwrite = True Then  '덮어씌우기
                    p_FileOffset = 0    '덮어씌우기는 Offset 이 0 이다.
                    '커서는 파일의 처음 위치에 있다. (Seek p_FileNum, 1)
                Else    '이어받기
                    p_FileOffset = LOF(p_FileNum)
                    Seek p_FileNum, p_FileOffset + 1    '파일 커서를 마지막에 위치하고
                End If
                If Err Then '파일 오류가 발생하면
                    '전송 중지
                    Cut
                    On Error GoTo 0 '오류 처리 중지
                    '오류 발생
                    Err.Raise Err.Number, , "파일 오류:" & Err.Description
                    Exit Sub
                End If
                On Error GoTo 0 '오류 처리 중지
                tcpSocket.SendData p_FileOffset     '오프셋을 전달
                '타이머 온
                p_StartTime = Timer
                '수신 확인
                p_SendComplete = False
                '상태 변경
                m_State = ftcReceive
                RaiseEvent ChangeState(m_State)
            Else    '헤더가 충분하지 않다면 나머지 헤더를 기다린다.
                Exit Sub
            End If
            
        Case TransferState.ftcSendReady '송신 대기중이라면
            'Timeout 타이머 중지
            tmrTimeout.Enabled = False
            '데이터 수신 (송신 파일 Offset)
            Dim FileOffset As Long
            
            tcpSocket.GetData FileOffset, vbInteger
            'Offset 이 음수(-)이면 전송 취소
            If FileOffset < 0 Then
                'Error 이벤트 발생
                RaiseEvent Error(62102, "잘못된 파일 오프셋 값이 수신되었습니다.")
                '전송 중지
                Cut
                Exit Sub
            ElseIf FileOffset > LOF(p_FileNum) Then '파일 크기를 넘어선 오프셋
                '수신측에 같은 이름으로 존재하는 파일의 크기가 더 큰 경우이다.
                FileOffset = LOF(p_FileNum) '오프셋을 파일 끝으로 지정
            ElseIf FileOffset = LOF(p_FileNum) Then '파일 크기와 같은 오프셋
                '수신측에 이미 파일이 존재하는 경우이다.
                '즉시 전송을 중단
                Cut
                Exit Sub
            End If
            '파일 커서를 Offset 으로 이동
            '(Offset은 시작 위치가 0 이고 Seek 은 시작 위치가 1이다)
            Seek p_FileNum, FileOffset + 1
            '타이머 온
            p_StartTime = Timer
            '송신 확인
            p_SendComplete = False
            '상태 변경
            m_State = ftcSend
            RaiseEvent ChangeState(m_State)
            '데이터 송신을 시작하기 위해 소켓의 SendComplete 이벤트 핸들러를 강제 호출
            tcpSocket_SendComplete
        Case TransferState.ftcSend      '송신중이라면... DataArrival 이벤트가 발생할리 없다.
        Case TransferState.ftcReady     '대기 상태라면.. DataArrival 이벤트가 발생할리 없다.
            m_State = ftcReceiveReady
            RaiseEvent ChangeState(m_State)
            '수신일경우 이기때문에 강제로 수신한다.
            Call tcpSocket_DataArrival(bytesTotal)
    End Select
End Sub

Private Sub tcpSocket_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    RaiseEvent Error(Number, "Socket 오류:" & Description)
End Sub

Private Sub tcpSocket_SendComplete()
    '송신중이 아니면 무시
    If m_State <> ftcSend Then Exit Sub
    '파일 송신이 끝났으면
    If EOF(p_FileNum) Then
        '송신 완료 이벤트를 발생시키고
        RaiseEvent SendComplete(LOF(p_FileNum))
        '송신 확인
        p_SendComplete = True
        '연결 해제
        Cut
        '송신 작업 종료
        Exit Sub
    End If
    
    '만약 남은 전송 용량이 PayloadSize 이상이면
    If LOF(p_FileNum) - Loc(p_FileNum) >= m_PayloadSize Then
        '버퍼 사이즈 조절
        ReDim Buffer(m_PayloadSize - 1)
    Else    '남은 전송 용량이 PayloadSize 미만이면
        '송신이 완료 되었다면
        If LOF(p_FileNum) = Loc(p_FileNum) Then
            '송신 완료 이벤트를 발생시키고
            RaiseEvent SendComplete(Loc(p_FileNum))
            '송신 확인
            p_SendComplete = True
            '연결 해제.. 는 하지 않는다. 송신버퍼의 내용이 모두 전달되야 하므로.
            'Cut
            'pds2004 준비 상태로 돌린다.
            Close p_FileNum
            m_State = ftcReady
            RaiseEvent ChangeState(m_State)
            '송신 작업 종료
            Exit Sub
        End If
        '버퍼 사이즈 조절
        ReDim Buffer(LOF(p_FileNum) - Loc(p_FileNum) - 1)
    End If
    '파일로부터 버퍼 사이즈만큼 읽어들인다.
    Get p_FileNum, , Buffer
    On Error Resume Next
    '데이터 송신
    tcpSocket.SendData Buffer
    If Err Then '송신 오류가 발생되면
        '오류 이벤트를 발생시키고
        RaiseEvent Error(Err.Number, Err.Description)
        '전송 중지
        Cut
        Exit Sub
    End If
    On Error GoTo 0
    '전송 과정 이벤트 발생
    RaiseEvent SendProgress(Loc(p_FileNum))

    '날짜변경을 체크하고 Cps 계산
    p_NowTime = Timer
    If p_NowTime < p_StartTime Then    '전송중에 날짜가 바뀌었다면
        p_StartTime = p_StartTime + 86400
    End If
    m_Cps = Loc(p_FileNum) \ (p_NowTime - p_StartTime + 1)
End Sub

Private Sub tmrTimeout_Timer()
    'Error 이벤트 발생
    RaiseEvent Error(62101, "SendTimer Timeout. 수신측에서 응답이 없습니다.")
    '전송 중지
    Cut
    '타이머 중지
    tmrTimeout.Enabled = False
End Sub

Private Sub UserControl_Initialize()
    '송수신 확인값은 기본적으로 True
    p_SendComplete = True
    p_ReceiveComplete = True
    lblVersion.Caption = CStr(App.Major + (App.Minor / 10))
End Sub

'컨트롤의 크기를 이미지 아이콘 크기에 맞춥니다.
Private Sub UserControl_Resize()
    If UserControl.Height <> imgControl.Height Then UserControl.Height = imgControl.Height
    If UserControl.Width <> imgControl.Width Then UserControl.Width = imgControl.Width
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14,1,2,0
Public Property Get State() As TransferState
Attribute State.VB_Description = "컨트롤의 상태를 TransferState 열거형으로 반환합니다."
Attribute State.VB_MemberFlags = "400"
    State = m_State
End Property

Public Property Let State(ByVal New_State As TransferState)
    If Ambient.UserMode Then Err.Raise 382, , "State 속성은 읽기 전용입니다."
    m_State = New_State
    PropertyChanged "State"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=8,1,2,0
Public Property Get Cps() As Long
Attribute Cps.VB_Description = "파일의 전송 속도를 cps 단위로 반환합니다."
Attribute Cps.VB_MemberFlags = "400"
    Cps = m_Cps
End Property

Public Property Let Cps(ByVal New_Cps As Long)
    If Ambient.UserMode Then Err.Raise 382, , "Cps 속성은 읽기 전용입니다."
    m_Cps = New_Cps
    PropertyChanged "Cps"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,2,
Public Property Get SendFilePath() As String
Attribute SendFilePath.VB_Description = "송신중이거나 송신할 파일의 전체 경로를 설정하거나 반환합니다."
Attribute SendFilePath.VB_MemberFlags = "400"
    SendFilePath = m_SendFilePath
End Property

Public Property Let SendFilePath(ByVal New_SendFilePath As String)
    If Ambient.UserMode = False Then Err.Raise 387, , , "SendFilePath 속성은 디자인 모드에서 변경할 수 없습니다."
    If Dir(New_SendFilePath) = "" Then
        Err.Raise 62001, , New_SendFilePath & " 은 유효한 파일 경로가 아닙니다."
        Exit Property
    End If
    m_SendFilePath = New_SendFilePath
    p_SendFileName = Mid(m_SendFilePath, InStrRev(m_SendFilePath, "\") + 1)
    m_SendFileSize = FileLen(m_SendFilePath)
    PropertyChanged "SendFilePath"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,
Public Property Get ReceiveDirPath() As String
Attribute ReceiveDirPath.VB_Description = "파일을 수신할 디렉토리를 설정하거나 반환합니다."
    ReceiveDirPath = m_ReceiveDirPath
End Property

Public Property Let ReceiveDirPath(ByVal New_ReceiveDirPath As String)
    '마지막 \ 를 제거
    If Right(New_ReceiveDirPath, 1) = "\" Then
        New_ReceiveDirPath = Left(New_ReceiveDirPath, Len(New_ReceiveDirPath) - 1)
    End If
    
    '입력이 없을 경우 현재 폴더로 지정
    If Dir(New_ReceiveDirPath, vbDirectory) = "." Then
        New_ReceiveDirPath = "."
    End If
    
    '경로가 존재하지 않을 경우
    If Dir(New_ReceiveDirPath, vbDirectory) = "" Then
        '런타임이라면
        If Ambient.UserMode Then
            Err.Raise 62002, , New_ReceiveDirPath & " 디렉토리는 유효한 경로가 아닙니다."
        Else    '디자인타임 이라면
            MsgBox New_ReceiveDirPath & " 디렉토리는 유효한 경로가 아닙니다.", vbOKOnly, "경로 설정 오류"
        End If
        Exit Property
    End If
    m_ReceiveDirPath = New_ReceiveDirPath
    PropertyChanged "ReceiveDirPath"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=8,1,2,0
Public Property Get ReceiveFileSize() As Long
Attribute ReceiveFileSize.VB_Description = "수신중이거나 수신할 파일의 크기를 반환합니다."
Attribute ReceiveFileSize.VB_MemberFlags = "400"
    ReceiveFileSize = m_ReceiveFileSize
End Property

Public Property Let ReceiveFileSize(ByVal New_ReceiveFileSize As Long)
    If Ambient.UserMode Then Err.Raise 382, , "ReceiveFileSize 속성은 읽기 전용입니다."
    m_ReceiveFileSize = New_ReceiveFileSize
    PropertyChanged "ReceiveFileSize"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=0,0,0,true
Public Property Get EnableReceive() As Boolean
Attribute EnableReceive.VB_Description = "파일 수신 가능 여부를 설정하거나 반환합니다."
    EnableReceive = m_EnableReceive
End Property

Public Property Let EnableReceive(ByVal New_EnableReceive As Boolean)
    '속성이 변경되지 않는다면 무시
    If m_EnableReceive = New_EnableReceive Then Exit Property
    '대기 상태이면 소켓의 상태를 변경
    '(파일 전송중이라면 소켓 상태는 변경하지 않는다; 전송이 끝나면 적용됨)
    If m_State = ftcReady Then
        '수신 가능이 되면 소켓을 대기 상태로
        If New_EnableReceive = True Then
            tcpSocket.Listen
        Else    '수신 불가가 되면 소켓을 닫는다.
            tcpSocket.Close
        End If
    End If
    m_EnableReceive = New_EnableReceive
    PropertyChanged "EnableReceive"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,1,2,
Public Property Get ReceiveFileName() As String
Attribute ReceiveFileName.VB_Description = "수신중이거나 수신할 파일 이름을 반환합니다."
Attribute ReceiveFileName.VB_MemberFlags = "400"
    ReceiveFileName = m_ReceiveFileName
End Property

Public Property Let ReceiveFileName(ByVal New_ReceiveFileName As String)
    If Ambient.UserMode Then Err.Raise 382, , "ReceiveFileName 속성은 읽기 전용입니다." & vbCrLf _
        & "수신되는 파일의 경로를 변경하려면 ReceiveDirPath 속성을 변경하고" & vbCrLf _
        & "수신되는 파일의 이름을 변경하려면 ReceiveStart 이벤트에서 FileName 을 변경하십시오."
    m_ReceiveFileName = New_ReceiveFileName
    PropertyChanged "ReceiveFileName"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=5
Public Sub SendFile(Optional ByVal FilePath As String, Optional ByVal TimeoutSec As Integer = 0)
Attribute SendFile.VB_Description = "파일 송신을 시작합니다. 송신할 파일의 전체 경로를 지정할 수 있습니다."
    If m_State <> ftcReady Then
        'Err.Raise 62009, , "파일 전송중에는 SendFile 메소드를 호출할 수 없습니다."
        RaiseEvent Error(62009, "파일 전송중에는 SendFile 메소드를 호출할 수 없습니다.")
        Exit Sub
    End If
    '타임아웃 타이머 설정 범위 체크
    If TimeoutSec < 0 Or TimeoutSec > 60 Then
        Err.Raise 62010, , "TimeoutSec 설정 범위가 잘못 지정되었습니다. (0~60초)"
        Exit Sub
    End If
    '타임아웃 타이머 설정
    tmrTimeout.Interval = TimeoutSec * 1000
    '파일 경로를 설정
    If FilePath <> "" Then
        SendFilePath = FilePath
    End If
'    '소켓을 닫고 목적 호스트에 접속
'    tcpSocket.Close
'    'Localport 는 임의의 포트가 설정되도록
'    tcpSocket.LocalPort = 0
'    tcpSocket.Connect m_RemoteHost, RemotePort
'    '상태 변경
'    m_State = ftcSendReady
'    RaiseEvent ChangeState(m_State)

    ' 공유기에서 접근한 포트에도 파일을 송신하기 위하여 연결을 유지한다.
    '연결중이지 않을경우 새로 연결한다.
    If tcpSocket.State <> sckConnected Then
        '소켓을 닫고 목적 호스트에 접속
        tcpSocket.Close
        'Localport 는 임의의 포트가 설정되도록
        tcpSocket.LocalPort = 0
        tcpSocket.Connect m_RemoteHost, RemotePort
        '상태 변경
        m_State = ftcSendReady
        RaiseEvent ChangeState(m_State)
    Else
        ' 강제로 연결을 호출한다.
        m_State = ftcSendReady
    
        
        Call tcpSocket_Connect
        
        '파일을 열고
        p_FileNum = FreeFile
        Open m_SendFilePath For Binary Access Read As p_FileNum
        
        '상태 변경
        RaiseEvent ChangeState(m_State)
        
    End If

End Sub

Public Sub RemoteConnect()
    If m_State <> ftcReady Then
        'Err.Raise 62009, , "파일 전송중에는 SendFile 메소드를 호출할 수 없습니다."
        RaiseEvent Error(62009, "파일 전송중에는 SendFile 메소드를 호출할 수 없습니다.")
        Exit Sub
    End If
    
    If m_RemoteHost = "" Or Val(RemotePort) = 0 Then
        MsgBox "연결 대상이 올바르지 않습니다.", vbInformation, "확인"
        Exit Sub
    End If
    
    '소켓을 닫고 목적 호스트에 접속
    tcpSocket.Close
    'Localport 는 임의의 포트가 설정되도록
    tcpSocket.LocalPort = 0
    tcpSocket.Connect m_RemoteHost, RemotePort
    
    '상태 변경
    m_State = ftcReady
    RaiseEvent ChangeState(m_State)
End Sub

Public Sub RemoteClose()
    If m_State <> ftcReady Then
        'Err.Raise 62009, , "파일 전송중에는 SendFile 메소드를 호출할 수 없습니다."
        RaiseEvent Error(62009, "파일 전송중에는 SendFile 메소드를 호출할 수 없습니다.")
        Exit Sub
    End If
    
    '소켓을 닫고 목적 호스트에 접속
    tcpSocket.Close
    
    '상태 변경
    m_State = ftcReady
    RaiseEvent ChangeState(m_State)
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=5
Public Sub Cut()
Attribute Cut.VB_Description = "파일 전송을 중단합니다."
    '컨트롤 상태에 따라
    If m_State = ftcReady Then '대기 상태라면 아무런 작업도 수행하지 않는다. (파일 전송중에만 작동)
        Exit Sub
    End If
    '열려진 파일을 닫고
    Close p_FileNum
    
'    ' pds2004가 주석처리함
'    ' 내부 IP에서도 수신이 가능하도록 수정
'
'    '소켓을 닫은 후
'    tcpSocket.Close
'    '수신 가능이면 대기 상태로 전환
'    tcpSocket.LocalPort = m_LocalPort
'    If m_EnableReceive Then tcpSocket.Listen
    
    '각 변수 값을 초기화
    m_Cps = 0
    p_SendComplete = True
    p_ReceiveComplete = True
    '상태 변경
    m_State = ftcReady
    RaiseEvent ChangeState(m_State)
End Sub

'사용자 정의 컨트롤에 대한 속성을 초기화합니다.
Private Sub UserControl_InitProperties()
    m_State = m_def_State
    m_Cps = m_def_Cps
    m_SendFilePath = m_def_SendFilePath
    m_ReceiveDirPath = CurDir   ' m_def_ReceiveDirPath
    m_ReceiveFileSize = m_def_ReceiveFileSize
    m_EnableReceive = m_def_EnableReceive
    m_ReceiveFileName = m_def_ReceiveFileName
    m_SendFileSize = m_def_SendFileSize
    m_RemotePort = m_def_RemotePort
    m_LocalPort = m_def_LocalPort
    m_RemoteHost = m_def_RemoteHost
    m_PayloadSize = m_def_PayloadSize
    m_ReceiveFilePath = m_def_ReceiveFilePath
    m_Version = App.Major + (App.Minor / 10)
End Sub

'저장소에서 속성값을 로드합니다.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_State = PropBag.ReadProperty("State", m_def_State)
    m_Cps = PropBag.ReadProperty("Cps", m_def_Cps)
    m_SendFilePath = PropBag.ReadProperty("SendFilePath", m_def_SendFilePath)
    m_ReceiveDirPath = PropBag.ReadProperty("ReceiveDirPath", m_def_ReceiveDirPath)
    m_ReceiveFileSize = PropBag.ReadProperty("ReceiveFileSize", m_def_ReceiveFileSize)
    m_EnableReceive = PropBag.ReadProperty("EnableReceive", m_def_EnableReceive)
    m_ReceiveFileName = PropBag.ReadProperty("ReceiveFileName", m_def_ReceiveFileName)
    m_SendFileSize = PropBag.ReadProperty("SendFileSize", m_def_SendFileSize)
    m_RemotePort = PropBag.ReadProperty("RemotePort", m_def_RemotePort)
    m_LocalPort = PropBag.ReadProperty("LocalPort", m_def_LocalPort)
    m_RemoteHost = PropBag.ReadProperty("RemoteHost", m_def_RemoteHost)
    m_PayloadSize = PropBag.ReadProperty("PayloadSize", m_def_PayloadSize)
    m_ReceiveFilePath = PropBag.ReadProperty("ReceiveFilePath", m_def_ReceiveFilePath)
    m_Version = PropBag.ReadProperty("Version", m_Version)
    
    '런타임에 컨트롤이 로딩되면 Ready 상태가 된다.
    '디자인 타임에는 대기 상태가 되지 않는다.
    If Ambient.UserMode Then
        '수신 가능이면 대기 상태로 전환
        tcpSocket.LocalPort = m_LocalPort
        
        If m_EnableReceive Then tcpSocket.Listen
        
        m_State = ftcReady
    End If
End Sub

'속성값을 저장소에 기록합니다.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("State", m_State, m_def_State)
    Call PropBag.WriteProperty("Cps", m_Cps, m_def_Cps)
    Call PropBag.WriteProperty("SendFilePath", m_SendFilePath, m_def_SendFilePath)
    Call PropBag.WriteProperty("ReceiveDirPath", m_ReceiveDirPath, m_def_ReceiveDirPath)
    Call PropBag.WriteProperty("ReceiveFileSize", m_ReceiveFileSize, m_def_ReceiveFileSize)
    Call PropBag.WriteProperty("EnableReceive", m_EnableReceive, m_def_EnableReceive)
    Call PropBag.WriteProperty("ReceiveFileName", m_ReceiveFileName, m_def_ReceiveFileName)
    Call PropBag.WriteProperty("SendFileSize", m_SendFileSize, m_def_SendFileSize)
    Call PropBag.WriteProperty("RemotePort", m_RemotePort, m_def_RemotePort)
    Call PropBag.WriteProperty("LocalPort", m_LocalPort, m_def_LocalPort)
    Call PropBag.WriteProperty("RemoteHost", m_RemoteHost, m_def_RemoteHost)
    Call PropBag.WriteProperty("PayloadSize", m_PayloadSize, m_def_PayloadSize)
    Call PropBag.WriteProperty("ReceiveFilePath", m_ReceiveFilePath, m_def_ReceiveFilePath)
    Call PropBag.WriteProperty("Version", m_Version, m_def_Version)
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=8,1,2,0
Public Property Get SendFileSize() As Long
Attribute SendFileSize.VB_Description = "송신중이거나 송신할 파일의 크기를 반환합니다."
Attribute SendFileSize.VB_MemberFlags = "400"
    SendFileSize = m_SendFileSize
End Property

Public Property Let SendFileSize(ByVal New_SendFileSize As Long)
    If Ambient.UserMode Then Err.Raise 382, , "SendFileSize 는 읽기 전용 속성입니다."
    m_SendFileSize = New_SendFileSize
    PropertyChanged "SendFileSize"
End Property


'사용자 함수 ===============================================================================
Private Function AscLen(DataString As String) As Integer
    AscLen = LenB(StrConv(DataString, vbFromUnicode))
End Function

Private Function AscLeft(ByVal St As String, ByVal Length As Integer)
    For I = 1 To Len(St)
        If MidB(St, I * 2, 1) = ChrB(0) Then    '영문이면
            Length = Length - 1
        Else
            Length = Length - 2
        End If
        If Length < 1 Then '탐색이 끝나면.
            Exit For
        End If
    Next I
    AscLeft = Left(St, I)
End Function

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=8,0,0,3279
Public Property Get RemotePort() As Long
Attribute RemotePort.VB_Description = "원격 컴퓨터에 연결될 포트를 반환하거나 설정합니다."
    RemotePort = m_RemotePort
End Property

Public Property Let RemotePort(ByVal New_RemotePort As Long)
    On Error Resume Next
    '대기 상태일 때만 소켓에 곧장 적용된다.
    If m_State = ftcReady Then
        '상태 전환을 위해 소켓을 닫는다.
        tcpSocket.Close
        tcpSocket.RemotePort = New_RemotePort
    End If
    If Err Then
        If Ambient.UserMode Then
            Err.Raise 62006, , "RemotePort 포트 번호의 범위가 잘못 지정되었습니다. (0~65535)"
        Else
            MsgBox "RemotePort 포트 번호의 범위가 잘못 지정되었습니다. (0~65535)", vbOKOnly, "포트 번호 오류"
        End If
        Exit Property
    End If
    On Error GoTo 0
    m_RemotePort = New_RemotePort
    PropertyChanged "RemotePort"
    '수신 가능이면 소켓을 대기 상태로
    '대기 상태일 때만 소켓에 곧장 적용된다.
    If m_State = ftcReady And m_EnableReceive Then tcpSocket.Listen
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=8,0,0,3279
Public Property Get LocalPort() As Long
Attribute LocalPort.VB_Description = "로컬 컴퓨터에서 사용되는 포트를 반환하거나 설정합니다."
    LocalPort = m_LocalPort
End Property

Public Property Let LocalPort(ByVal New_LocalPort As Long)
    '상태 전환을 위해 소켓을 닫는다.
    On Error Resume Next
    '대기 상태일 때만 소켓에 곧장 적용된다.
    If m_State = ftcReady Then
        tcpSocket.Close
        tcpSocket.LocalPort = New_LocalPort
    End If
    If Err Then
        If Ambient.UserMode Then
            Err.Raise 62007, , "LocalPort 포트 번호의 범위가 잘못 지정되었습니다. (0~65535)"
        Else
            MsgBox "LocalPort 포트 번호의 범위가 잘못 지정되었습니다. (0~65535)", vbOKOnly, "포트 번호 오류"
        End If
        Exit Property
    End If
    On Error GoTo 0
    m_LocalPort = New_LocalPort
    PropertyChanged "LocalPort"
    '수신 가능이면 소켓을 대기 상태로
    '대기 상태일 때만 소켓에 곧장 적용된다.
    If m_State = ftcReady And m_EnableReceive Then tcpSocket.Listen
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,localhost
Public Property Get RemoteHost() As String
Attribute RemoteHost.VB_Description = "원격 컴퓨터를 식별하기 위해 사용된 이름을 반환하거나 설정합니다."
    RemoteHost = m_RemoteHost
End Property

Public Property Let RemoteHost(ByVal New_RemoteHost As String)
    '상태 전환을 위해 소켓을 닫는다.
    On Error Resume Next
    '대기 상태일 때만 소켓에 곧장 적용된다.
    If m_State = ftcReady Then
        tcpSocket.Close
        tcpSocket.RemoteHost = New_RemoteHost
    End If
    If Err Then
        If Ambient.UserMode Then
            Err.Raise 62008, , "RemoteHost 호스트 값 설정이 잘못되었습니다."
        Else
            MsgBox "RemoteHost 호스트 값 설정이 잘못되었습니다.", vbOKOnly, "호스트 주소 오류"
        End If
        Exit Property
    End If
    On Error GoTo 0
    m_RemoteHost = New_RemoteHost
    PropertyChanged "RemoteHost"
    '수신 가능이면 소켓을 대기 상태로
    '대기 상태일 때만 소켓에 곧장 적용된다.
    If m_State = ftcReady And m_EnableReceive Then tcpSocket.Listen
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=8,0,0,8169
Public Property Get PayloadSize() As Long
Attribute PayloadSize.VB_Description = "파일의 송신 단위 크기를 bytes 단위로 반환하거나 설정합니다."
    PayloadSize = m_PayloadSize
End Property

Public Property Let PayloadSize(ByVal New_PayloadSize As Long)
    '범위를 넘어선 입력 처리
    If New_PayloadSize < 1 Or New_PayloadSize > 65535 Then
        If Ambient.UserMode Then
            Err.Raise 62010, , "PayloadSize 입력 범위가 잘못 지정되었습니다. (1~65535)"
        Else
            MsgBox "PayloadSize 입력 범위가 잘못 지정되었습니다. (1~65535)", vbOKOnly, "PayloadSize 입력 오류"
        End If
        Exit Property
    End If
    m_PayloadSize = New_PayloadSize
    PropertyChanged "PayloadSize"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,1,2,
Public Property Get ReceiveFilePath() As String
Attribute ReceiveFilePath.VB_Description = "수신중이거나 수신할 파일의 전체 경로를 반환합니다."
Attribute ReceiveFilePath.VB_MemberFlags = "400"
    ReceiveFilePath = m_ReceiveFilePath
End Property

Public Property Let ReceiveFilePath(ByVal New_ReceiveFilePath As String)
    If Ambient.UserMode Then Err.Raise 382, , "ReceiveFilePath 속성은 읽기 전용입니다." & vbCrLf _
        & "수신되는 파일의 경로를 변경하려면 ReceiveDirPath 속성을 변경하고" & vbCrLf _
        & "수신되는 파일의 이름을 변경하려면 ReceiveStart 이벤트에서 FileName 을 변경하십시오."
    m_ReceiveFilePath = New_ReceiveFilePath
    PropertyChanged "ReceiveFilePath"
End Property
'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=12,1,1,0
Public Property Get Version() As Single
Attribute Version.VB_Description = "컨트롤 버전 빌드번호를 반환합니다."
    Version = m_Version
End Property

Public Property Let Version(ByVal New_Version As Single)
    If Ambient.UserMode = False Then
        MsgBox "Version 속성은 읽기 전용입니다.", vbOKOnly, "접근 권한 오류"
        Exit Property
    End If
    If Ambient.UserMode Then Err.Raise 382, , "Version 속성은 읽기 전용입니다."
    m_Version = New_Version
    PropertyChanged "Version"
End Property

