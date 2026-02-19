Attribute VB_Name = "basCheck"
Option Explicit

Public 판매취소_Flag As Boolean
Public 접수결제_Flag As Boolean

'====================================================================================================
' Procedure : CheckMobileNumber
' DateTime  : 07-01-18 01:50
' Author    : BlueNice
' Last Edit :
' Return Val:
'----------------------------------------------------------------------------------------------------
' 설     명 : 전달된 번호가 휴대전화 인지를 확인한다.
'====================================================================================================
Public Function CheckMobileNumber(ByVal sNumber As String, ByRef sTel() As String) As Boolean
    Dim sTemp   As String
    Dim sLen    As Integer
    On Error GoTo CheckMobileNumber_Error

    CheckMobileNumber = False

    sTemp = Trim(sNumber)
    sTemp = Replace(sTemp, "-", "")
    sTemp = Replace(sTemp, ")", "")
    sTemp = Replace(sTemp, "/", "")
    
    If Left(sTemp, 2) <> "01" Or Len(sTemp) <= 9 Then Exit Function
    
    sLen = Len(sTemp)
    
    ' 0164401234
    If sLen = 10 Then
        sTel(0) = Left(sTemp, 3):   sTel(1) = Mid(sTemp, 4, 3): sTel(2) = Mid(sTemp, 7, 4)
        CheckMobileNumber = True
        
    '01190044523
    ElseIf sLen = 11 Then
        sTel(0) = Left(sTemp, 3):   sTel(1) = Mid(sTemp, 4, 4): sTel(2) = Mid(sTemp, 8, 4)
        CheckMobileNumber = True
        
    Else
        CheckMobileNumber = False
    End If

    On Error GoTo 0
    
    Exit Function

CheckMobileNumber_Error:
    CheckMobileNumber = False
    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure CheckMobileNumber of Form frmSMS"
End Function

'====================================================================================================
' Procedure : CheckTelNumber
' DateTime  : 07-01-18 01:50
' Author    : BlueNice
' Last Edit :
' Return Val:
'----------------------------------------------------------------------------------------------------
' 설     명 : 전달된 번호가 전화번호 인지를 확인한다.
'====================================================================================================
Public Function CheckTelNumber(ByVal sNumber As String, ByRef sTel() As String) As Boolean
    Dim sTemp   As String
    Dim sLen    As Integer
    On Error GoTo CheckTelNumber_Error

    CheckTelNumber = False

    sTemp = Trim(sNumber)
    sTemp = Replace(sTemp, "-", "")
    sTemp = Replace(sTemp, ")", "")
    sTemp = Replace(sTemp, "/", "")
    If Len(sTemp) <= 6 Then Exit Function
    
    sLen = Len(sTemp)
    
    ' 216 1234
    If sLen = 7 Then
        sTel(0) = "":   sTel(1) = Left(sTemp, 3): sTel(2) = Right(sTemp, 4)
        CheckTelNumber = True
        
    '2345 1234
    ElseIf sLen = 8 Then
        sTel(0) = "":   sTel(1) = Left(sTemp, 4): sTel(2) = Right(sTemp, 4)
        CheckTelNumber = True
        
    '2345 1234
    ElseIf sLen >= 9 And sLen <= 12 Then
        sTel(2) = Right(sTemp, 4)
        sTel(1) = Mid(Right(sTemp, 8), 1, 4)
        sTel(0) = Replace(sTemp, sTel(1) & sTel(2), "")
        CheckTelNumber = True
        
    Else
        CheckTelNumber = False
    End If

    On Error GoTo 0
    Exit Function

CheckTelNumber_Error:
    CheckTelNumber = False
    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure CheckTelNumber of Form frmSMS"
End Function


Public Function CheckCouponNumber(sCouponNum As String) As Integer
    ' 정상 0
    CheckCouponNumber = -1

    ' 입력 형태 검사
    If Len(sCouponNum) <> M_COUPON_LENGTH Then
        CheckCouponNumber = -1
        
        Exit Function
        
    ' 크렌즈겔러리 모피 행사 50,000 2009-12-31일까지
    ElseIf Left(sCouponNum, 2) = "02" And 가맹점정보.지사코드 <> M_COUPON_KLENZ_CODE Then
        CheckCouponNumber = -1
        
        Exit Function
        
    ElseIf Left(sCouponNum, 2) = "00" And 가맹점정보.지사코드 <> M_COUPON_KLENZ_CODE Then
        CheckCouponNumber = -1
        
        Exit Function
        
    ElseIf Left(sCouponNum, 2) = "01" And 가맹점정보.지사코드 = M_COUPON_KLENZ_CODE Then
        CheckCouponNumber = -1
        
        Exit Function
    
    ' 유효기간 검사 오류
    ElseIf Left(sCouponNum, 2) = "01" And Format(Date, "YYYY-MM-DD") > "20090831" Then
        CheckCouponNumber = -2
        
        Exit Function
    
    ' 유효기간 검사 오류
    ElseIf Left(sCouponNum, 2) = "05" And Format(Date, "yyyyMMdd") > "20111231" Then
        CheckCouponNumber = -2
        
        Exit Function
    
    End If
    
    CheckCouponNumber = 0
End Function



Public Function CheckSMSConnect(objDBase As ADODB.Connection) As Boolean
    On Error GoTo ErrRtn
    
    Dim HostConn    As String
    
    HostConn = ""
    HostConn = HostConn & "Provider=SQLOLEDB.1;"
    HostConn = HostConn & "Persist Security Info=False;"
    HostConn = HostConn & "User ID=" & m_SMS.UserID & ";"
    HostConn = HostConn & "Password=" & m_SMS.UserPW & ";"
    HostConn = HostConn & "Initial Catalog=" & m_SMS.DBName & ";"
    HostConn = HostConn & "Data Source=" & m_SMS.ServerIP
    
    Set objDBase = Nothing
    Set objDBase = New ADODB.Connection
    
    If objDBase.State = adStateOpen Then objDBase.Close
    
    objDBase.ConnectionTimeout = 10
    objDBase.CommandTimeout = m_SMS.timeout
    objDBase.Open HostConn
    
    CheckSMSConnect = True
    
    Exit Function

ErrRtn:
    
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Function


Public Function SQL_주의대상복구() As Boolean
    Dim FHandle As Integer
    Dim sText   As String
    Dim sFileFullName   As String
    
    On Error GoTo ErrRtn
    sFileFullName = App.Path & "\SQL_Data\SP_주의대상복구.sql"
    
    ' 로그 파일을 생성하지 않는다.
    If Dir(sFileFullName, vbDirectory) <> "" Then Exit Function


    FHandle = FreeFile
    Open sFileFullName For Append As FHandle

    Print #FHandle, "    ----------------------------------------------------------------------"
    Print #FHandle, "    -- 반드시 c:\CleanAid\Sql_Data 폴더의 내용을 백업후 실행 할것"
    Print #FHandle, "    ----------------------------------------------------------------------"

    Print #FHandle, "    EXEC sp_resetstatus 'CleanAid';"
    Print #FHandle, "    ALTER DATABASE CleanAid SET EMERGENCY"
    Print #FHandle, "    DBCC checkdb('CleanAid')"
    Print #FHandle, "    ALTER DATABASE CleanAid SET SINGLE_USER WITH ROLLBACK IMMEDIATE"
    Print #FHandle, "    DBCC CheckDB ('CleanAid', REPAIR_ALLOW_DATA_LOSS)"
    Print #FHandle, "    ALTER DATABASE CleanAid SET MULTI_USER"
    
    Close #FHandle
    Exit Function

ErrRtn:
    
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Function
