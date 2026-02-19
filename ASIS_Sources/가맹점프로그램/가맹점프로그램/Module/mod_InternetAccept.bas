Attribute VB_Name = "mod_InternetAccept"
Option Explicit

Public Function GetInternetAccept() As ADODB.RecordSet
   
    On Error GoTo ErrRtn
     If Lusoft_Connection(HostCon, "Lusoft") = True Then
        '
        Query = ""
        Query = Query & " SELECT"
        Query = Query & "        rq_code as '주문번호',"
        Query = Query & "        (select mb_name from rb_member  where mb_idx = a.mb_idx) as '이름',"
        Query = Query & "        rq_type1 as '의류',"
        Query = Query & "        rq_type2 as '신발',"
        Query = Query & "        rq_type3 as '이불',"
        Query = Query & "        rq_take_date as '수거일자'"
        Query = Query & " FROM rb_request a left join rb_delivery b on a.rq_idx = b.rq_idx"
        Query = Query & " WHERE rq_franchisee_idx in ( select mb_idx from rb_member where mb_level = '9' and mb_grade = '2' and mb_com_code = '" & 가맹점정보.가맹점코드 & "' )"
        Query = Query & "   AND de_status = 3 and (select count(*) from rb_cart where a.rq_idx = rq_idx) = 0"
        Set ADORs = New ADODB.RecordSet
        ADORs.Open Query, HostCon, adOpenForwardOnly, adLockReadOnly
        
        Set GetInternetAccept = ADORs.Clone
        ADORs.Close
        
        Set ADORs = Nothing
    Else
        MsgBox "본사 서버와 연결할 수 없습니다.  인터넷을 확인 하여 주십시요.", vbInformation, "확인"
        Exit Function
    End If
    
    Exit Function
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Function

Public Function GetInternetDelivery() As ADODB.RecordSet
   
    On Error GoTo ErrRtn
     If Lusoft_Connection(HostCon, "Lusoft") = True Then
        '
        Query = ""
        Query = Query & " SELECT"
        Query = Query & "        rq_code as '주문번호',"
        Query = Query & "        (select mb_name from rb_member  where mb_idx = a.mb_idx) as '이름',"
        Query = Query & "        rq_type1 as '의류',"
        Query = Query & "        rq_type2 as '신발',"
        Query = Query & "        rq_type3 as '이불',"
        Query = Query & "        rq_take_date as '수거일자'"
        Query = Query & " FROM rb_request a"
        Query = Query & " WHERE rq_franchisee_idx in ( select mb_idx from rb_member where mb_level = '9' and mb_grade = '2' and mb_com_code = '" & 가맹점정보.가맹점코드 & "' )"
        Query = Query & "   AND a.rq_status = 2 and rq_idx in (select os_rq_idx from rb_order_history where os_status = 5)"
        Set ADORs = New ADODB.RecordSet
        ADORs.Open Query, HostCon, adOpenForwardOnly, adLockReadOnly
        
        Set GetInternetDelivery = ADORs.Clone
        ADORs.Close
        
        Set ADORs = Nothing
    Else
        MsgBox "본사 서버와 연결할 수 없습니다.  인터넷을 확인 하여 주십시요.", vbInformation, "확인"
        Exit Function
    End If
    
    Exit Function
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Function


Public Function SetInternetAccept(Internet_접수번호 As String, 적요 As String, 접수금액 As String, Index As String, tag As String)
   If Internet_접수번호 = "" Then Exit Function
    On Error GoTo ErrRtn
     If Lusoft_Connection(HostCon, "Lusoft") = True Then
        Query = ""
        Query = Query & " insert into rb_cart(rq_idx, rq_code, mb_id, mb_idx, mb_session, pd_kind, pd_kind2, pd_idx, pd_name, ct_cnt, ct_amount, ct_option, cr_idx, ct_wdate)"
        Query = Query & " select rq_idx,"
        Query = Query & "        rq_code,"
        Query = Query & "        b.mb_id,"
        Query = Query & "        a.mb_idx,"
        Query = Query & "        '' as mb_session,"
        Query = Query & "        '1' as pd_kind,"
        Query = Query & "        '1' as pd_kind2,"
        Query = Query & "        " & Index & " as pd_idx,"
        Query = Query & "        '" & 적요 & "' as pd_name,"
        Query = Query & "        1 as ct_cnt,"
        Query = Query & "        " & 접수금액 & " as ct_amount,"
        Query = Query & "        '" & tag & "' as ct_option,"
        Query = Query & "        0 as cr_idx,"
        Query = Query & "         getdate() as ct_wdate"
        Query = Query & " from rb_request a left join rb_member b on a.mb_idx = b.mb_idx"
        Query = Query & " where rq_code = '" & Internet_접수번호 & "'"
        HostCon.Execute Query
        HostCon.Close
    Else
        MsgBox "본사 서버와 연결할 수 없습니다.  인터넷을 확인 하여 주십시요.", vbInformation, "확인"
        Exit Function
    End If
    
    Exit Function
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Function


Public Function SetInternetDelivery(Internet_접수번호 As String)
   If Internet_접수번호 = "" Then Exit Function
    On Error GoTo ErrRtn
     If Lusoft_Connection(HostCon, "Lusoft") = True Then
        Query = ""
        Query = Query & " insert into tb_delivery_request"
        Query = Query & " values ('" & Internet_접수번호 & "')"
        HostCon.Execute Query
        HostCon.Close
    Else
        MsgBox "본사 서버와 연결할 수 없습니다.  인터넷을 확인 하여 주십시요.", vbInformation, "확인"
        Exit Function
    End If
    
    Exit Function
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Function



'====================================================================================================
' Procedure : Server_Connection
' DateTime  : 2008-04-15 04:13
' Author    : pds2004
' Last Edit :
' Return Val:
'----------------------------------------------------------------------------------------------------
' 설     명 : 최초 신규 매장코드가 없을 경우 본사에 등록되어있는 내용을 가저온다.
'====================================================================================================
Public Function Lusoft_Connection(HostCon As ADODB.Connection, Optional sOffice As String) As Boolean
    Dim sServer   As String
    Dim sDatabase As String
    Dim sID       As String
    Dim sPWD      As String
    
    Dim Server_Connect As String
    
    On Error GoTo ErrRtn
 
    '서버에 접속여부
    Server_Connect = GetIniStr("SERVER", "Server_Connect", "", iniFile)                  '
    
    If Server_Connect = "N" Then
        Lusoft_Connection = False
    Else
        sServer = "115.89.220.5,8657"
        sDatabase = "Lusoft"
        sID = Get_Decrypt(GetIniStr("SERVER", "ID", "", iniFile), "")             '
        sPWD = Get_Decrypt(GetIniStr("SERVER", "PWD", "", iniFile), "")           '
        
        Set HostCon = Nothing
        Set HostCon = New ADODB.Connection
    
        If HostCon.State = adStateOpen Then HostCon.Close
    
        With HostCon
            .ConnectionString = "Provider=SQLOLEDB;Persist Security Info=False;User ID=" & sID & ";Password=" & sPWD & ";Initial Catalog=" & sDatabase & ";Data Source=" & sServer
            .CursorLocation = adUseClient
            .ConnectionTimeout = 10
            .CommandTimeout = IIf(m_CommandTimeOut = 0, 30, m_CommandTimeOut)
            .Open
        End With
        
        Lusoft_Connection = True
    End If
    
    Exit Function

ErrRtn:
    Lusoft_Connection = False
    
    '서버접속 오류메시지는 막음
    'Call Error_Msg("Server_Connection", Err.Source, Err.Number, Err.Description)
End Function
