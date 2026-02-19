Attribute VB_Name = "basBizPurio"
' Biz Purio 용 문자발신 모듈
Option Explicit
Dim SMSCon As ADODB.Connection
Public Const KAKAO_PRICE = 10
Public Const SMS_PRICE = 15


' 가입안내 문자전송
Public Sub send_Kakao_Invite(str_Customer_Phone As String, str_Customer_Name As String)
    Dim SSQL As String
ReMsg:
    SSQL = ""
    SSQL = SSQL & " INSERT INTO biz_purio..biz_msg (MSG_TYPE, CMID, REQUEST_TIME, SEND_TIME, DEST_PHONE, SEND_PHONE,"
    SSQL = SSQL & " MSG_BODY, TEMPLATE_CODE, SENDER_KEY, NATION_CODE,"
    SSQL = SSQL & " RE_TYPE, RE_BODY, SHOP_CODE,ATTACHED_FILE)"
    SSQL = SSQL & " VALUES ("
    SSQL = SSQL & " 6, "
    SSQL = SSQL & " replace(replace(replace(replace(CONVERT(varchar, getdate(), 121),'-',''),':',''),'.',''),' ',''), "
    SSQL = SSQL & " Getdate(), "
    SSQL = SSQL & " Getdate(), "
    SSQL = SSQL & " '" & str_Customer_Phone & "', "
    SSQL = SSQL & " '" & 가맹점정보.전화SMS & "',"
    SSQL = SSQL & " '" & str_Customer_Name & " 고객님" & vbCrLf
    SSQL = SSQL & "저희 크린에이드 " & 가맹점정보.가맹점명 & "에 고객님의 정보를 등록 하였습니다." & vbCrLf
    SSQL = SSQL & "" & vbCrLf
    SSQL = SSQL & "크린에이드는 월~토 영업하며 일요일에는 영업하지 않습니다. " & vbCrLf
    SSQL = SSQL & "(마트매장 제외) " & vbCrLf
    SSQL = SSQL & "" & vbCrLf
    SSQL = SSQL & "매장별 영업시간은," & vbCrLf
    SSQL = SSQL & "크린에이드 홈페이지를 통해 확인하여 주시기 바랍니다. " & vbCrLf
    SSQL = SSQL & "(링크클릭 시 홈페이지-매장정보 이동)" & vbCrLf
    SSQL = SSQL & "" & vbCrLf
    SSQL = SSQL & "정성을 다하는 크린에이드가 되겠습니다.',"
    SSQL = SSQL & " 'bizp_2019010409335304807674997', "
    SSQL = SSQL & " 'f100882941d6ad62b2b383e6559c61f32208981a', "
    SSQL = SSQL & " '82',"
    SSQL = SSQL & " 'SMS', "
    SSQL = SSQL & " '저희 크린에이드를 사용해 주셔서 감사합니다. " & vbCrLf
    SSQL = SSQL & " 정성을 다하는 크린에이드가 되겠습니다.',"
    SSQL = SSQL & "'" & 가맹점정보.가맹점코드 & "',"
    SSQL = SSQL & "'invite.json'"
    SSQL = SSQL & ")"
    If CheckConnect Then
    On Error Resume Next
        SMSCon.Execute SSQL
        If Err.Number <> 0 Then
            Err.Clear
            GoTo ReMsg
        End If
        Call MsgBox("가입 안내문이 카카오톡으로 전송되었습니다.")
    End If
    On Error GoTo 0
End Sub


' 배송안내 문자전송
Public Sub send_Kakao_Delivery(str_UserKey As String, str_Customer_Phone As String, re_msg As String)
    Dim SSQL As String
ReMsg:
    SSQL = ""
    SSQL = SSQL & " INSERT INTO biz_purio..biz_msg (MSG_TYPE, CMID, REQUEST_TIME, SEND_TIME, DEST_PHONE, SEND_PHONE,"
    SSQL = SSQL & " MSG_BODY, TEMPLATE_CODE, SENDER_KEY, NATION_CODE,"
    SSQL = SSQL & " RE_TYPE, RE_BODY, SHOP_CODE, USER_KEY, AD_FLAG)"
    SSQL = SSQL & " VALUES ("
    SSQL = SSQL & " 6, "
    SSQL = SSQL & " replace(replace(replace(replace(CONVERT(varchar, getdate(), 121),'-',''),':',''),'.',''),' ',''), "
    SSQL = SSQL & " Getdate(), "
    SSQL = SSQL & " Getdate(), "
    SSQL = SSQL & " '" & str_Customer_Phone & "', "
    SSQL = SSQL & " '" & 가맹점정보.전화SMS & "',"
    SSQL = SSQL & " '[세탁 완료 안내]" & vbCrLf
    SSQL = SSQL & vbCrLf
    SSQL = SSQL & "안녕하세요." & vbCrLf
    SSQL = SSQL & "고객님께서 맡기신 소중한" & vbCrLf
    SSQL = SSQL & "세탁물의 세탁이 완료되었습니다." & vbCrLf
    SSQL = SSQL & "수령 부탁드립니다." & vbCrLf
    SSQL = SSQL & vbCrLf
    SSQL = SSQL & "가맹점 : " & 가맹점정보.가맹점명 & "" & vbCrLf
    SSQL = SSQL & "문의 : " & 가맹점정보.전화매장 & "', "
    SSQL = SSQL & " 'bizp_2018120409581704807569091', "
    SSQL = SSQL & " 'f100882941d6ad62b2b383e6559c61f32208981a', "
    SSQL = SSQL & " '82',"
    SSQL = SSQL & " 'SMS', "
    SSQL = SSQL & "'안녕하세요."
    SSQL = SSQL & "고객님께서 맡기신 소중한"
    SSQL = SSQL & "세탁물의 세탁이 완료되었습니다"
    SSQL = SSQL & "감사합니다가맹점입니다',"
    SSQL = SSQL & "'" & 가맹점정보.가맹점코드 & "','" + str_UserKey + "','1'"
    SSQL = SSQL & ")"
    If CheckConnect Then
    On Error Resume Next
        SMSCon.Execute SSQL
        If Err.Number <> 0 Then
            Err.Clear
            GoTo ReMsg
        End If
    End If
    On Error GoTo 0
End Sub


' 배송안내 문자전송
Public Sub send_Purio_SMS(str_UserKey As String, str_Customer_Phone As String, msg As String, Optional AD_FLAG As String = "0")
    Dim SSQL As String
ReMsg:
    SSQL = ""
    SSQL = SSQL & " INSERT INTO biz_purio..biz_msg (MSG_TYPE, CMID, REQUEST_TIME, SEND_TIME, DEST_PHONE, SEND_PHONE, MSG_BODY, SHOP_CODE, USER_KEY, AD_FLAG)"
    SSQL = SSQL & " VALUES ("
    SSQL = SSQL & " 0, "
    SSQL = SSQL & " replace(replace(replace(replace(CONVERT(varchar, getdate(), 121),'-',''),':',''),'.',''),' ',''), "
    SSQL = SSQL & " Getdate(), "
    SSQL = SSQL & " Getdate(), "
    SSQL = SSQL & " '" & str_Customer_Phone & "', "
    SSQL = SSQL & " '" & 가맹점정보.전화SMS & "',"
    SSQL = SSQL & " '" & msg & "', "
    SSQL = SSQL & " '" & 가맹점정보.가맹점코드 & "','" & str_UserKey & "','" & AD_FLAG & "'"
    SSQL = SSQL & ")"
    If CheckConnect Then
    On Error Resume Next

        SMSCon.Execute SSQL
        If Err.Number <> 0 Then
            Err.Clear
            GoTo ReMsg
        End If
    End If
    On Error GoTo 0
End Sub

' 친구톡 문자전송
Public Sub send_Kakao_Friend(str_UserKey As String, str_Customer_Phone As String, str_Message As String, Optional Attached_File As String = "", Optional Event_Type As String = "", Optional re_msg As String = "")
    Dim SSQL As String
    
    If Event_Type <> "" Then
        Dim RecordSet As ADODB.RecordSet
        Dim Event_Code As String
        SSQL = "exec [BIZ_PURIO].[dbo].[GENERATE_EVENT_CODE_PHONE] '" & Event_Type & "','" & str_Customer_Phone & "'"
        Set RecordSet = SMSCon.Execute(SSQL)
        
        If Not RecordSet.EOF Then
            Event_Code = RecordSet!Event_Code
        End If
        RecordSet.Close
        Set RecordSet = Nothing
    End If
ReMsg:
    SSQL = ""
    SSQL = SSQL & " INSERT INTO biz_purio..biz_msg (MSG_TYPE, CMID, REQUEST_TIME, SEND_TIME, DEST_PHONE, SEND_PHONE,"
    SSQL = SSQL & " MSG_BODY, SENDER_KEY, NATION_CODE,SHOP_CODE, USER_KEY, AD_FLAG"
    If Attached_File <> "" Then '첨부파일이 있을경우
        SSQL = SSQL & " ,ATTACHED_FILE"
    End If
    If re_msg <> "" Then '첨부파일이 있을경우
        SSQL = SSQL & " ,RE_TYPE, RE_BODY"
    End If
    SSQL = SSQL & " )"
    SSQL = SSQL & " VALUES ("
    SSQL = SSQL & " 7, "
    SSQL = SSQL & " replace(replace(replace(replace(CONVERT(varchar, getdate(), 121),'-',''),':',''),'.',''),' ',''), "
    SSQL = SSQL & " Getdate(), "
    SSQL = SSQL & " Getdate(), "
    SSQL = SSQL & " '" & str_Customer_Phone & "', "
    SSQL = SSQL & " '" & 가맹점정보.전화SMS & "',"
    SSQL = SSQL & " '" & Replace(str_Message, "{행사코드}", Event_Code) & "', "
    SSQL = SSQL & " 'f100882941d6ad62b2b383e6559c61f32208981a',"
    SSQL = SSQL & " '82',"
    SSQL = SSQL & " '" & 가맹점정보.가맹점코드 & "','" + str_UserKey + "','0'"
    If Attached_File <> "" Then '첨부파일이 있을경우
        SSQL = SSQL & " ,'" & Attached_File & "'"
    End If
    If re_msg <> "" Then '첨부파일이 있을경우
        SSQL = SSQL & " ,'SMS','" & re_msg & "'"
    End If
    SSQL = SSQL & " )"
    If CheckConnect Then
    On Error Resume Next

        SMSCon.Execute SSQL
        If Err.Number <> 0 Then
            Err.Clear
            GoTo ReMsg
        End If
    End If
    On Error GoTo 0
End Sub

Public Function GetMoney() As String
On Error GoTo ErrRtn
    Dim SSQL As String
    Dim TempMoney As Long
    
    If Not CheckConnect Then
        GoTo ErrRtn
    End If
    
    Dim RecordSet As ADODB.RecordSet
    SSQL = "SELECT SUM(MONEY) as MONEY FROM [BIZ_PURIO].[dbo].TBL_SMS_MONEY WHERE SHOP_CODE = '" & 가맹점정보.가맹점코드 & "'"
    Set RecordSet = SMSCon.Execute(SSQL)
    If Not RecordSet.EOF Then
        TempMoney = Val(RecordSet!Money)
    End If
    GetMoney = CStr(TempMoney)
    Exit Function
ErrRtn:
    GetMoney = "0"
'    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure CheckConnect of Module basBizPurio"
End Function

Public Function GetMsgList() As ADODB.RecordSet
On Error GoTo ErrRtn
    Dim SSQL As String
    
    If Not CheckConnect Then
        GoTo ErrRtn
    End If
    
    Dim RecordSet As ADODB.RecordSet
    SSQL = "SELECT * FROM [BIZ_PURIO].[dbo].TBL_SMS_MSG WHERE SHOP_CODE = '" & 가맹점정보.가맹점코드 & "' ORDER BY MSG_ID"
    Set GetMsgList = SMSCon.Execute(SSQL)
    Exit Function
ErrRtn:
    Set GetMsgList = Nothing
End Function

Public Function GetMsg(MsgID As String) As ADODB.RecordSet
On Error GoTo ErrRtn
    Dim SSQL As String
    
    If Not CheckConnect Then
        GoTo ErrRtn
    End If
    
    Dim RecordSet As ADODB.RecordSet
    SSQL = "SELECT * FROM [BIZ_PURIO].[dbo].TBL_SMS_MSG WHERE SHOP_CODE = '" & 가맹점정보.가맹점코드 & "' AND MSG_ID = " & MsgID
    Set GetMsg = SMSCon.Execute(SSQL)
    Exit Function
ErrRtn:
    Set GetMsg = Nothing
    
End Function

Public Function GetEvtList() As ADODB.RecordSet
On Error GoTo ErrRtn
    Dim SSQL As String
    Dim SHOP_TYPE As String
    If Not CheckConnect Then
        GoTo ErrRtn
    End If
    
    Dim RecordSet As ADODB.RecordSet
    If 가맹점정보.SMS_EMART = "Y" Then
        SHOP_TYPE = "EMART"
    Else
        SHOP_TYPE = "SHOP"
    End If
    SSQL = "SELECT * FROM [BIZ_PURIO].[dbo].tbl_event WHERE START_DATE <= getdate() and END_DATE >= getdate() AND SUBSTRING(week,datepart(dw,getdate()),1) = '1' AND shop_type IN ('ALL','" & SHOP_TYPE & "')"
    Set GetEvtList = SMSCon.Execute(SSQL)
    Exit Function
ErrRtn:
    Set GetEvtList = Nothing
End Function


Public Function SaveKaKaoMsg(MsgID As String, Title As String, msg As String, re_msg As String) As Boolean
On Error GoTo ErrRtn
    Dim SSQL As String
    
    If Not CheckConnect Then
        GoTo ErrRtn
    End If
    
    If MsgID <> "" Then
        SSQL = "UPDATE [BIZ_PURIO].[dbo].TBL_SMS_MSG SET title = '" & Title & "', msg = '" & msg & "', re_msg = '" & re_msg & "' WHERE SHOP_CODE = '" & 가맹점정보.가맹점코드 & "' AND MSG_ID = " & MsgID
    Else
        SSQL = "INSERT INTO [BIZ_PURIO].[dbo].TBL_SMS_MSG(SHOP_CODE,TITLE, MSG, RE_MSG, TYPE) VALUES ('" & 가맹점정보.가맹점코드 & "','" & Title & "','" & msg & "','" & re_msg & "','KAKAO')"
    End If
    
    SMSCon.Execute (SSQL)
    SaveKaKaoMsg = True
    Exit Function
ErrRtn:
    SaveKaKaoMsg = False
End Function

Public Function DeleteKaKaoMsg(MsgID As String) As Boolean
On Error GoTo ErrRtn
    Dim SSQL As String
    
    If Not CheckConnect Then
        GoTo ErrRtn
    End If
    
    If MsgID <> "" Then
        SSQL = "DELETE [BIZ_PURIO].[dbo].TBL_SMS_MSG WHERE SHOP_CODE = '" & 가맹점정보.가맹점코드 & "' AND MSG_ID = " & MsgID
    End If
    
    SMSCon.Execute (SSQL)
    DeleteKaKaoMsg = True
    Exit Function
ErrRtn:
    DeleteKaKaoMsg = False
End Function

Public Function GetEventRate(Event_Code As String) As String
On Error GoTo ErrRtn
    Dim SSQL As String
    
    
    If Not CheckConnect Then
        GoTo ErrRtn
    End If
    
    Dim RecordSet As ADODB.RecordSet
    SSQL = "SELECT used, (SELECT rate from [BIZ_PURIO].[dbo].tbl_event_rate WHERE EVENT_TYPE = a.EVENT_TYPE) AS rate FROM [BIZ_PURIO].[dbo].tbl_event_code a WHERE code = '" & Event_Code & "'"
    Set RecordSet = SMSCon.Execute(SSQL)
    If Not RecordSet.EOF Then
        If RecordSet!used = "0" Then
            GetEventRate = RecordSet!Rate
            RecordSet.Close
            Set RecordSet = Nothing
            Exit Function
        Else
            GetEventRate = "ERROR 이미 사용한 행사코드 입니다."
            RecordSet.Close
            Set RecordSet = Nothing
            Exit Function
        End If
    End If
    RecordSet.Close
    Set RecordSet = Nothing
    GetEventRate = "ERROR 행사코드를 확인할수 없습니다."
    Exit Function
ErrRtn:
    GetEventRate = "ERROR 행사코드확인중 오류가 발생되었습니다."
End Function


Public Function UpdateEventCode(Event_Code As String) As Boolean
On Error GoTo ErrRtn
    Dim SSQL As String
    
    
    If Not CheckConnect Then
        GoTo ErrRtn
    End If
    
    Dim RecordSet As ADODB.RecordSet
    SSQL = "UPDATE [BIZ_PURIO].[dbo].tbl_event_code SET used = '1', shop_code = '" & 가맹점정보.가맹점코드 & "' WHERE code = '" & Event_Code & "'"
    SMSCon.Execute (SSQL)

    UpdateEventCode = True
    Exit Function
ErrRtn:
    UpdateEventCode = False
End Function

Private Function CheckConnect() As Boolean
    On Error GoTo ErrRtn
    
    Dim HostConn    As String
    
    HostConn = ""
    HostConn = HostConn & "Provider=SQLOLEDB.1;"
    HostConn = HostConn & "Persist Security Info=False;"
    HostConn = HostConn & "User ID=" & m_SMS.UserID & ";"
    HostConn = HostConn & "Password=" & m_SMS.UserPW & ";"
    HostConn = HostConn & "Initial Catalog=" & m_SMS.DBName & ";"
    HostConn = HostConn & "Data Source=" & m_SMS.ServerIP
    m_CommandTimeOut = IIf(m_SMS.timeout = 0, 30, m_SMS.timeout)

    Set SMSCon = Nothing
    Set SMSCon = New ADODB.Connection
    
    If SMSCon.State = adStateOpen Then SMSCon.Close
    
    SMSCon.ConnectionTimeout = 10
    SMSCon.CommandTimeout = m_CommandTimeOut
    SMSCon.Open HostConn
    
    CheckConnect = True
    
    On Error GoTo 0
    
    Exit Function

ErrRtn:
    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure CheckConnect of Module basBizPurio"
End Function


