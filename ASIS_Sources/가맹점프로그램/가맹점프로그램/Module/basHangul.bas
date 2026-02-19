Attribute VB_Name = "basHangul"
Option Explicit

Public Const IME_HANGUL = &H1
Public Const IME_ENGLISH = &H0
Public Const IME_NONE = &H0
Public Toggle_Check As Boolean

Declare Function ImmGetContext Lib "imm32.dll" (ByVal hwnd As Long) As Long

Declare Function ImmSetConversionStatus Lib "imm32.dll" _
        (ByVal hIMC As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
'// 여기까지가 한영전환을 위한 선언

'RAS Connection Status
Global Const RASCS_PAUSED = &H1000
Global Const RASCS_DONE = &H2000

Global Const RASCS_OpenPort = 0
Global Const RASCS_PortOpened = 1
Global Const RASCS_ConnectDevice = 2
Global Const RASCS_DeviceConnected = 3
Global Const RASCS_AllDevicesConnected = 4
Global Const RASCS_Authenticate = 5
Global Const RASCS_AuthNotify = 6
Global Const RASCS_AuthRetry = 7
Global Const RASCS_AuthCallback = 8
Global Const RASCS_AuthChangePassword = 9
Global Const RASCS_AuthProject = 10
Global Const RASCS_AuthLinkSpeed = 11
Global Const RASCS_AuthAck = 12
Global Const RASCS_ReAuthenticate = 13
Global Const RASCS_Authenticated = 14
Global Const RASCS_PrepareForCallback = 15
Global Const RASCS_WaitForModemReset = 16
Global Const RASCS_WaitForCallback = 17
Global Const RASCS_Projected = 18
Global Const RASCS_StartAuthentication = 19
Global Const RASCS_CallbackComplete = 20
Global Const RASCS_LogonNetwork = 21

Global Const RASCS_Interactive = RASCS_PAUSED
Global Const RASCS_RetryAuthentication = RASCS_PAUSED + 1
Global Const RASCS_CallbackSetByCaller = RASCS_PAUSED + 2
Global Const RASCS_PasswordExpired = RASCS_PAUSED + 3

Global Const RASCS_Connected = RASCS_DONE
Global Const RASCS_Disconnected = RASCS_DONE + 1


'RAS Error Status
Global Const RASBASE = 600
Global Const SUCCESS = 0

Global Const PENDING = (RASBASE + 0)
Global Const ERROR_INVALID_PORT_HANDLE = (RASBASE + 1)
Global Const ERROR_PORT_ALREADY_OPEN = (RASBASE + 2)
Global Const ERROR_BUFFER_TOO_SMALL = (RASBASE + 3)
Global Const ERROR_WRONG_INFO_SPECIFIED = (RASBASE + 4)
Global Const ERROR_CANNOT_SET_PORT_INFO = (RASBASE + 5)
Global Const ERROR_PORT_NOT_CONNECTED = (RASBASE + 6)
Global Const ERROR_EVENT_INVALID = (RASBASE + 7)
Global Const ERROR_DEVICE_DOES_NOT_EXIST = (RASBASE + 8)
Global Const ERROR_DEVICETYPE_DOES_NOT_EXIST = (RASBASE + 9)
Global Const ERROR_INVALID_BUFFER = (RASBASE + 10)
Global Const ERROR_ROUTE_NOT_AVAILABLE = (RASBASE + 11)
Global Const ERROR_ROUTE_NOT_ALLOCATED = (RASBASE + 12)
Global Const ERROR_INVALID_COMPRESSION_SPECIFIED = (RASBASE + 13)
Global Const ERROR_OUT_OF_BUFFERS = (RASBASE + 14)
Global Const ERROR_PORT_NOT_FOUND = (RASBASE + 15)
Global Const ERROR_ASYNC_REQUEST_PENDING = (RASBASE + 16)
Global Const ERROR_ALREADY_DISCONNECTING = (RASBASE + 17)
Global Const ERROR_PORT_NOT_OPEN = (RASBASE + 18)
Global Const ERROR_PORT_DISCONNECTED = (RASBASE + 19)
Global Const ERROR_NO_ENDPOINTS = (RASBASE + 20)
Global Const ERROR_CANNOT_OPEN_PHONEBOOK = (RASBASE + 21)
Global Const ERROR_CANNOT_LOAD_PHONEBOOK = (RASBASE + 22)
Global Const ERROR_CANNOT_FIND_PHONEBOOK_ENTRY = (RASBASE + 23)
Global Const ERROR_CANNOT_WRITE_PHONEBOOK = (RASBASE + 24)
Global Const ERROR_CORRUPT_PHONEBOOK = (RASBASE + 25)
Global Const ERROR_CANNOT_LOAD_STRING = (RASBASE + 26)
Global Const ERROR_KEY_NOT_FOUND = (RASBASE + 27)
Global Const ERROR_DISCONNECTION = (RASBASE + 28)
Global Const ERROR_REMOTE_DISCONNECTION = (RASBASE + 29)
Global Const ERROR_HARDWARE_FAILURE = (RASBASE + 30)
Global Const ERROR_USER_DISCONNECTION = (RASBASE + 31)
Global Const ERROR_INVALID_SIZE = (RASBASE + 32)
Global Const ERROR_PORT_NOT_AVAILABLE = (RASBASE + 33)
Global Const ERROR_CANNOT_PROJECT_CLIENT = (RASBASE + 34)
Global Const ERROR_UNKNOWN = (RASBASE + 35)
Global Const ERROR_WRONG_DEVICE_ATTACHED = (RASBASE + 36)
Global Const ERROR_BAD_STRING = (RASBASE + 37)
Global Const ERROR_REQUEST_TIMEOUT = (RASBASE + 38)
Global Const ERROR_CANNOT_GET_LANA = (RASBASE + 39)
Global Const ERROR_NETBIOS_ERROR = (RASBASE + 40)
Global Const ERROR_SERVER_OUT_OF_RESOURCES = (RASBASE + 41)
Global Const ERROR_NAME_EXISTS_ON_NET = (RASBASE + 42)
Global Const ERROR_SERVER_GENERAL_NET_FAILURE = (RASBASE + 43)
Global Const WARNING_MSG_ALIAS_NOT_ADDED = (RASBASE + 44)
Global Const ERROR_AUTH_INTERNAL = (RASBASE + 45)
Global Const ERROR_RESTRICTED_LOGON_HOURS = (RASBASE + 46)
Global Const ERROR_ACCT_DISABLED = (RASBASE + 47)
Global Const ERROR_PASSWD_EXPIRED = (RASBASE + 48)
Global Const ERROR_NO_DIALIN_PERMISSION = (RASBASE + 49)
Global Const ERROR_SERVER_NOT_RESPONDING = (RASBASE + 50)
Global Const ERROR_FROM_DEVICE = (RASBASE + 51)
Global Const ERROR_UNRECOGNIZED_RESPONSE = (RASBASE + 52)
Global Const ERROR_MACRO_NOT_FOUND = (RASBASE + 53)
Global Const ERROR_MACRO_NOT_DEFINED = (RASBASE + 54)
Global Const ERROR_MESSAGE_MACRO_NOT_FOUND = (RASBASE + 55)
Global Const ERROR_DEFAULTOFF_MACRO_NOT_FOUND = (RASBASE + 56)
Global Const ERROR_FILE_COULD_NOT_BE_OPENED = (RASBASE + 57)
Global Const ERROR_DEVICENAME_TOO_LONG = (RASBASE + 58)
Global Const ERROR_DEVICENAME_NOT_FOUND = (RASBASE + 59)
Global Const ERROR_NO_RESPONSES = (RASBASE + 60)
Global Const ERROR_NO_COMMAND_FOUND = (RASBASE + 61)
Global Const ERROR_WRONG_KEY_SPECIFIED = (RASBASE + 62)
Global Const ERROR_UNKNOWN_DEVICE_TYPE = (RASBASE + 63)
Global Const ERROR_ALLOCATING_MEMORY = (RASBASE + 64)
Global Const ERROR_PORT_NOT_CONFIGURED = (RASBASE + 65)
Global Const ERROR_DEVICE_NOT_READY = (RASBASE + 66)
Global Const ERROR_READING_INI_FILE = (RASBASE + 67)
Global Const ERROR_NO_CONNECTION = (RASBASE + 68)
Global Const ERROR_BAD_USAGE_IN_INI_FILE = (RASBASE + 69)
Global Const ERROR_READING_SECTIONNAME = (RASBASE + 70)
Global Const ERROR_READING_DEVICETYPE = (RASBASE + 71)
Global Const ERROR_READING_DEVICENAME = (RASBASE + 72)
Global Const ERROR_READING_USAGE = (RASBASE + 73)
Global Const ERROR_READING_MAXCONNECTBPS = (RASBASE + 74)
Global Const ERROR_READING_MAXCARRIERBPS = (RASBASE + 75)
Global Const ERROR_LINE_BUSY = (RASBASE + 76)
Global Const ERROR_VOICE_ANSWER = (RASBASE + 77)
Global Const ERROR_NO_ANSWER = (RASBASE + 78)
Global Const ERROR_NO_CARRIER = (RASBASE + 79)
Global Const ERROR_NO_DIALTONE = (RASBASE + 80)
Global Const ERROR_IN_COMMAND = (RASBASE + 81)
Global Const ERROR_WRITING_SECTIONNAME = (RASBASE + 82)
Global Const ERROR_WRITING_DEVICETYPE = (RASBASE + 83)
Global Const ERROR_WRITING_DEVICENAME = (RASBASE + 84)
Global Const ERROR_WRITING_MAXCONNECTBPS = (RASBASE + 85)
Global Const ERROR_WRITING_MAXCARRIERBPS = (RASBASE + 86)
Global Const ERROR_WRITING_USAGE = (RASBASE + 87)
Global Const ERROR_WRITING_DEFAULTOFF = (RASBASE + 88)
Global Const ERROR_READING_DEFAULTOFF = (RASBASE + 89)
Global Const ERROR_EMPTY_INI_FILE = (RASBASE + 90)
Global Const ERROR_AUTHENTICATION_FAILURE = (RASBASE + 91)
Global Const ERROR_PORT_OR_DEVICE = (RASBASE + 92)
Global Const ERROR_NOT_BINARY_MACRO = (RASBASE + 93)
Global Const ERROR_DCB_NOT_FOUND = (RASBASE + 94)
Global Const ERROR_STATE_MACHINES_NOT_STARTED = (RASBASE + 95)
Global Const ERROR_STATE_MACHINES_ALREADY_STARTED = (RASBASE + 96)
Global Const ERROR_PARTIAL_RESPONSE_LOOPING = (RASBASE + 97)
Global Const ERROR_UNKNOWN_RESPONSE_KEY = (RASBASE + 98)
Global Const ERROR_RECV_BUF_FULL = (RASBASE + 99)
Global Const ERROR_CMD_TOO_LONG = (RASBASE + 100)
Global Const ERROR_UNSUPPORTED_BPS = (RASBASE + 101)
Global Const ERROR_UNEXPECTED_RESPONSE = (RASBASE + 102)
Global Const ERROR_INTERACTIVE_MODE = (RASBASE + 103)
Global Const ERROR_BAD_CALLBACK_NUMBER = (RASBASE + 104)
Global Const ERROR_INVALID_AUTH_STATE = (RASBASE + 105)
Global Const ERROR_WRITING_INITBPS = (RASBASE + 106)
Global Const ERROR_INVALID_WIN_HANDLE = (RASBASE + 107)
Global Const ERROR_NO_PASSWORD = (RASBASE + 108)
Global Const ERROR_NO_USERNAME = (RASBASE + 109)
Global Const ERROR_CANNOT_START_STATE_MACHINE = (RASBASE + 110)
Global Const ERROR_GETTING_COMMSTATE = (RASBASE + 111)
Global Const ERROR_SETTING_COMMSTATE = (RASBASE + 112)
Global Const ERROR_COMM_FUNCTION = (RASBASE + 113)
Global Const ERROR_CONFIGURATION_PROBLEM = (RASBASE + 114)
Global Const ERROR_X25_DIAGNOSTIC = (RASBASE + 115)
Global Const ERROR_TOO_MANY_LINE_ERRORS = (RASBASE + 116)
Global Const ERROR_OVERRUN = (RASBASE + 117)
Global Const ERROR_ACCT_EXPIRED = (RASBASE + 118)
Global Const ERROR_CHANGING_PASSWORD = (RASBASE + 119)
Global Const ERROR_NO_ACTIVE_ISDN_LINES = (RASBASE + 120)
Global Const ERROR_NO_ISDN_CHANNELS_AVAILABLE = (RASBASE + 121)

'컨트롤 상태 열거형
Public Enum LaundrySendFlag
    lauInput            ' 입고 파일 전송
    lauSendMail         ' 메일 전송
    lauChulGo           ' 출고 파일 수신
    lauRecvMail         ' 메일 수신
    lauSaleData         ' 할인자료 수신
    lauPriceData        ' 금액표 수신
    lauDaySaleData      ' 목요세일 수신
    lauRepairData       ' 수선자료 수신
    lauSendCust         ' 고객자료 전송
    lauSendDB           ' DB 전송
    lauProgram          ' 프로그램
    lauMileage          ' 마일리지
    lauQNData           ' 보관 서비스 자료
    lauSendCoupoon      ' 쿠폰 자료
End Enum

' 통신에 필요한 내역을 선언한다.
Public Const S_STA As String = "OK_START"       ' 전송 시작 시점
Public Const S_END As String = "OK_END"         ' 전송 종료 시점
Public Const S_CUSTCODE As String = "MYCODE"    ' 거래처 코드
Public Const S_CUSTNAME As String = "MYNAME"    ' 거래처 이름
Public Const S_MYIP As String = "MYIP"          ' 자신의 ip를 전달한다
Public Const S_MYFILEPORT As String = "MYFILEPORT"      ' 자신의 접속 포트를 전송한다. ( 파일전송)
Public Const S_CHULGO As String = "CHULGOFILE"  ' 본사가 거래처에게 전송할 출고 파일명
Public Const S_MAIL As String = "MAILFILE"      ' 본사가 거래처에게 전송할 메일 파일명
Public Const S_GETFILE As String = "GETFILE"    ' 거래처에서 송신 파일을 요청한다.
Public Const S_GETFILELIST As String = "GETFILELIST"    ' 파일의 리스트를 요청한다.
Public Const S_GETFILELISTCOUNT As String = "GETFILELISTCOUNT" ' 파일의 리스트를 요청한다.
Public Const S_GETALLFILE As String = "GETALLFILE"      ' 전달한값에 해당하는 모든 파일전송을 요청한다.
Public Const S_GETPROGRAMFILE As String = "GETPROGRAMFILE"  ' 신규 버전의 파일을 요청한다.
Public Const S_FILELISTCOUNT As String = "FILECOUNT"    ' 파일의 수
Public Const S_FILELIST As String = "FILELIST"          ' 파일의 리스트
Public Const S_PROGRAMVERSION As String = "SENDVERSION" ' 프로그램의 버전을 보넨다.
Public Const S_GETPROGRAMVERSION As String = "GETVERSION" ' 프로그램의 버전을 요청한다.
