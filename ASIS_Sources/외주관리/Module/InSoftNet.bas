Attribute VB_Name = "insoftnet"
Option Explicit

Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetMenuItemID Lib "User" (ByVal hMenu As Integer, ByVal nPos As Integer) As Integer

Public Const MF_BYPOSITION = &H400&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&

Public HeadOffice       As String '지사코드
Public Rtn              As Integer  'MessageBox Return 값...

Public m_DBConnect          As String       'ADO 정의 프로그램 시작위치에서 읽어옴
Public Laundry_Code         As String
Public userMsg              As String
Public strSql               As String

Public Const REG_App    As String = "크린에이드"
Public gs_ReportConnect As String

Public scUrl As String
Public scName As String
Public scFold As String
Public scFoldName As String
Public scUpgrade As Boolean

Public DBUserID      As String
Public DBUserPwd    As String
Public DBCatalog     As String
Public DBServer      As String

''---------20080421
'Public Const scUrl As String = "http://www.clean-aid.co.kr:8090/business/"
'Public Const scName As String = "백상영업.exe"
'Public Const scFold As String = App.Path & "\"
'Public Const scFoldName As String = "백상영업UP.exe"
''---------20080421

Type ERROR_TYPE
     FileName       As String
     FullPath       As String
     ErrorMsg       As String
     VisibleMSG     As Boolean
     SaveLog        As Boolean
     ResumeMode     As Boolean
End Type
Public m_Error   As ERROR_TYPE

Type STORE_TYPE
    Office  As String
    Code    As String
    Name    As String
    OutClothCode As String
       
End Type

Public Store    As STORE_TYPE

Enum ConnectMode_Type
    Modem = 0
    Floppy = 1
    InterNet = 2
End Enum

Public Function GetDefaultValues() As Boolean
    
    sIniFile = App.Path & "\CleanAid.Ini"
    sCopyIniFile = App.Path & "\PDA\kiSync.INI"
    
    ' 변경 말것. 서버의 연결 프로그램에서 사용됨
    ' 다른 회사는 다른 이름이 있음.
    
    Store.Office = "CLEANAID"
    Store.Code = Trim(GetIniStr("Store DATA", "StoreCode", "", sIniFile))
    Store.Name = Trim(GetIniStr("Store DATA", "StoreName", "", sIniFile))
    Store.OutClothCode = Trim(GetIniStr("Store DATA", "OutClothCode", "", sIniFile))

End Function

Public Function SetOuterStr(defValue As String) As String
    Dim result As Integer
    sIniFile = App.Path & "\CleanAid.Ini"
        
    result = WritePrivateProfileString("Store DATA", "OutClothCode", defValue, sIniFile)
End Function

'Database Open Function
'Return Value ; True  : DB Open
'             ; False : DB Open 실패
Public Function DBOpen_Master(strDatabase As String) As Boolean
    Dim Provider    As String
    Dim Persist     As String
    Dim UserID      As String
    Dim UserPass    As String
    Dim Catalog     As String
    Dim Source      As String
    Dim ReadVal     As String
    Dim iTimeout    As Integer

    Dim DBPort      As String

    On Error GoTo DBOpenError
    
    DBUserID = Trim(GetIniStr("Store Server", "UserID", "", sIniFile))
    DBUserPwd = Trim(GetIniStr("Store Server", "UserPassword", "", sIniFile))
    
    'DBCatalog = Trim(GetIniStr("Store Server", "DatabaseName", "", sIniFile))
    DBCatalog = "LAUNDRY" & strDatabase
    
    DBServer = Trim(GetIniStr("Store Server", "ServerNameOrIP", "", sIniFile))
    iTimeout = Val(Trim(GetIniStr("Store Server", "CommandTimeout", "", sIniFile)))
    DBPort = Trim(GetIniStr("Store Server", "MessagePort", "", sIniFile))
    
    If MSTCon.State = adStateOpen Then
        MSTCon.Close
        Set MSTCon = Nothing
        
        DoEvents
    End If
    
    MSTCon.ConnectionString = "Provider=SQLOLEDB;Persist Security Info=False;User ID=" & DBUserID & ";Password=" & DBUserPwd & ";Initial Catalog=" & DBCatalog & ";Data Source=" & DBServer & "," & DBPort
    MSTCon.CommandTimeout = iTimeout
    MSTCon.CursorLocation = adUseClient    '이문장이 있어야 스프리드에 보여짐... 나머지는 잘모름
    MSTCon.Open
        
    DBOpen_Master = True
    
    Exit Function
    
DBOpenError:
    Err_Num = Err.Number
    Err_Dec = Err.Description
    
    If Err.Number = -2147467259 Then
    
    ElseIf Err.Number <> 0 Then
        Call Error_Msg("", Err.Source, Err.Number, Err.Description)
        
        'userMsg = Err.Description
        'MsgBox userMsg, 16, "Error", Err.HelpFile, Err.HelpContext
    End If
    
    DBOpen_Master = False
End Function

'+------------------------------------------------------
'+
'+ 2003/04/11
'+
'+루틴설명
'+  1. 목록에 선택된 내용을 DB에 적용 시킨다.
'+------------------------------------------------------
Public Sub SqlDataValue(rstSet As ADODB.Recordset, strSql As String)
    rstSet.CursorType = adOpenKeyset
    rstSet.LockType = adLockOptimistic
    rstSet.Open strSql, m_DBConnect, adOpenStatic, adLockBatchOptimistic, adCmdText
End Sub

'--------------------------------------------------------------------
'
'
'--------------------------------------------------------------------
Public Function fpSpread_Display(SS As Object, Rs As ADODB.Recordset)
    Dim i As Long
    Dim j As Long
    
    If Not Rs.BOF Then Rs.MoveFirst
    
    With SS
        .Redraw = False
        
        For i = 1 To Rs.Fields.Count
            .SetText i, 0, Rs.Fields.Item(i - 1).Name & ""
        Next i
        
        For i = 1 To Rs.RecordCount
            For j = 1 To Rs.Fields.Count
                .SetText j, i, Trim(Rs.Fields(j - 1)) & ""
            Next j
            
            .Row = i
            .RowHidden = False
            
            Rs.MoveNext
        Next i
        
        .Redraw = True
    End With
    
    If Not Rs.BOF Then Rs.MoveFirst
End Function

Public Function CheckDirectory(strDir As String, bFlag As Boolean) As Boolean
    Dim MyFile, MyName As String
    Dim bIsDir As Boolean
    
    CheckDirectory = False
    
    If Dir(strDir, vbDirectory) <> "" Then
        CheckDirectory = True
    End If
    
    If bFlag = True And CheckDirectory = False Then
        '디렉토리 생성
        MkDir strDir
        CheckDirectory = True
        Exit Function
    End If
End Function

'====================================================================================================
' Procedure : ProgramErrorLogWrite
' DateTime  : 2004-07-15 11:54
' Author    : pds2004
' Last Edit :
' Return Val:
'----------------------------------------------------------------------------------------------------
' 설     명 : 프로그램도중 오류가 발생하면 오류 내역을 저장한다.
'====================================================================================================
Public Sub ProgramErrorLogWrite(PrgError As ERROR_TYPE)
    Dim fp              As Long
    Dim MSG             As String

    On Error GoTo LogError

    MSG = "[" & Format(Date & Time, "YYYY-MM-DD hh:mm:ss") & "]" & PrgError.ErrorMsg
    
    fp = FreeFile
    Open PrgError.FullPath For Append As fp
    Print #fp, MSG
    Close #fp
    Exit Sub

LogError:
    Resume Next

End Sub


'Database Open Function
'Return Value ; True  : DB Open
'             ; False : DB Open 실패
Public Function DBOpen() As Boolean
    Dim Provider    As String
    Dim Persist     As String
    Dim UserID      As String
    Dim UserPass    As String
    Dim Catalog     As String
    Dim Server      As String
    Dim ReadVal     As String
    Dim iTimeout    As Integer
    
    Dim DBPort      As String

    On Error GoTo DBOpenError
        
    DBUserID = Trim(GetIniStr("Store Server", "UserID", "", sIniFile))
    DBUserPwd = Trim(GetIniStr("Store Server", "UserPassword", "", sIniFile))
    DBCatalog = Trim(GetIniStr("Store Server", "DatabaseName", "", sIniFile))
    DBServer = Trim(GetIniStr("Store Server", "ServerNameOrIP", "", sIniFile))
    iTimeout = Val(Trim(GetIniStr("Store Server", "CommandTimeout", "", sIniFile)))
    DBPort = Trim(GetIniStr("Store Server", "MessagePort", "", sIniFile))

    ADOCon.ConnectionString = "Provider=SQLOLEDB;Persist Security Info=False;User ID=" & DBUserID & ";Password=" & DBUserPwd & ";Initial Catalog=" & DBCatalog & ";Data Source=" & DBServer & "," & DBPort
    ADOCon.CommandTimeout = iTimeout
    ADOCon.CursorLocation = adUseClient    '이문장이 있어야 스프리드에 보여짐... 나머지는 잘모름
    ADOCon.Open
               
    DBOpen = True
    
    Exit Function
    
DBOpenError:
    Err_Num = Err.Number
    Err_Dec = Err.Description
    
    If Err.Number = -2147467259 Then
    
    ElseIf Err.Number <> 0 Then
        Call Error_Msg("", Err.Source, Err.Number, Err.Description)
        
        'userMsg = Err.Description
        'MsgBox userMsg, 16, "Error", Err.HelpFile, Err.HelpContext
    End If
    
    DBOpen = False
End Function
'====================================================================================================
' Procedure : Fnc_FromEnableCheck
' DateTime  : 2005-06-12 22:21
' Author    : pds2004
' Last Edit :
' Return Val:
'----------------------------------------------------------------------------------------------------
' 설     명 : 해당 폼이 활성화 되어 있으면 True를 리턴한다.
'====================================================================================================
Public Function Fnc_FromEnableCheck(FromName As String) As Boolean
    Dim tmpForm As Form
    
    Fnc_FromEnableCheck = False
    
    For Each tmpForm In Forms
        If UCase(tmpForm.Name) = UCase(FromName) Then
            Fnc_FromEnableCheck = True
            Exit Function
        End If
    Next tmpForm

End Function

'--------------------------------------------------------------------------------------------------------------
' Procedure : SQL_DB_Update
' DateTime  : 2007-02-07 13:47
' Author    : pds2004
' Purpose   : SQL DB를 업그레이드 시킨다.
'--------------------------------------------------------------------------------------------------------------
Public Sub SQL_DB_Update()

    On Error GoTo Error_Rtn
    Dim SSQL    As String
    
    
    SSQL = ""
    SSQL = SSQL & "CREATE TABLE [CouponUseData] ("
    SSQL = SSQL & "    [SaleDate] [char] (8) COLLATE Korean_Wansung_CI_AS NOT NULL ,"
    SSQL = SSQL & "    [StoreCode] [char] (6) COLLATE Korean_Wansung_CI_AS NOT NULL ,"
    SSQL = SSQL & "    [Number] [char] (8) COLLATE Korean_Wansung_CI_AS NOT NULL ,"
    SSQL = SSQL & "    [Cost] [int] NULL ,"
    SSQL = SSQL & "    [Money] [int] NULL ,"
    SSQL = SSQL & "    [CustNum] [char] (6) COLLATE Korean_Wansung_CI_AS NULL ,"
    SSQL = SSQL & "    [CustName] [char] (30) COLLATE Korean_Wansung_CI_AS NULL ,"
    SSQL = SSQL & "    [SaleMoney] [int] NULL ,"
    SSQL = SSQL & "    [SendYN] [char] (1) COLLATE Korean_Wansung_CI_AS NULL ,"
    SSQL = SSQL & "    [SendDate] [char] (8) COLLATE Korean_Wansung_CI_AS NULL ,"
    SSQL = SSQL & "    [StoreTag] [char] (3) COLLATE Korean_Wansung_CI_AS NULL ,"
    SSQL = SSQL & "    CONSTRAINT [PK_CouponUseData] PRIMARY KEY  CLUSTERED"
    SSQL = SSQL & "    ("
    SSQL = SSQL & "        [SaleDate],"
    SSQL = SSQL & "        [StoreCode],"
    SSQL = SSQL & "        [Number]"
    SSQL = SSQL & "    )  ON [PRIMARY]"
    SSQL = SSQL & " ) ON [PRIMARY]"
 

    ADOCon.Execute SSQL


    SSQL = " "
    SSQL = SSQL & " CREATE         PROCEDURE [SP_08002_04] " & vbNewLine
    SSQL = SSQL & " (   @IpStoreCode    Nvarchar(3), " & vbNewLine
    SSQL = SSQL & "     @lpData1    Nvarchar(8)," & vbNewLine
    SSQL = SSQL & "     @lpData2    Nvarchar(6)," & vbNewLine
    SSQL = SSQL & "     @lpData3    Nvarchar(8)," & vbNewLine
    SSQL = SSQL & "     @lpData4    INT," & vbNewLine
    SSQL = SSQL & "     @lpData5    INT," & vbNewLine
    SSQL = SSQL & "     @lpData6    Nvarchar(6)," & vbNewLine
    SSQL = SSQL & "     @lpData7    Nvarchar(30)," & vbNewLine
    SSQL = SSQL & "     @lpData8    INT," & vbNewLine
    SSQL = SSQL & "     @lpData9    Nvarchar(3) )" & vbNewLine
    SSQL = SSQL & " AS" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & " BEGIN TRAN" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "     DECLARE @Rec_Count      Smallint" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "     SELECT  @Rec_Count  =   Count(*)" & vbNewLine
    SSQL = SSQL & "     From CouponUseData" & vbNewLine
    SSQL = SSQL & "     WHERE   SaleDate    =   @lpData1" & vbNewLine
    SSQL = SSQL & "          AND  StoreCode =   @lpData2" & vbNewLine
    SSQL = SSQL & "          AND    Number      =   @lpData3" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "     IF  @Rec_Count  =   0" & vbNewLine
    SSQL = SSQL & "     BEGIN" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "         INSERT  INTO    CouponUseData" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "             ( SaleDate, StoreCode, Number, Cost, Money, CustNum, CustName, SaleMoney, StoreTag, SendYN, SendDate  )" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "         VALUES ( @lpData1,@lpData2,@lpData3,@lpData4,@lpData5,@lpData6,@lpData7,@lpData8,@lpData9, 'Y', substring(convert(varchar(8),getdate(),112),1,8) )" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "     End" & vbNewLine
    SSQL = SSQL & "     Else" & vbNewLine
    SSQL = SSQL & "     BEGIN" & vbNewLine
    SSQL = SSQL & "         Update CouponUseData" & vbNewLine
    SSQL = SSQL & "         SET Cost        =   @lpData4," & vbNewLine
    SSQL = SSQL & "             Money       =   @lpData5," & vbNewLine
    SSQL = SSQL & "             CustNum =   @lpData6," & vbNewLine
    SSQL = SSQL & "             CustName    =   @lpData7," & vbNewLine
    SSQL = SSQL & "             SaleMoney   =   @lpData8," & vbNewLine
    SSQL = SSQL & "             StoreTag    =   @lpData9" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "         WHERE   SaleDate    =   @lpData1" & vbNewLine
    SSQL = SSQL & "              AND  StoreCode =   @lpData2" & vbNewLine
    SSQL = SSQL & "              AND    Number      =   @lpData3" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "     End" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "     IF  @@ERROR     <>  0" & vbNewLine
    SSQL = SSQL & "         ROLLBACK TRAN" & vbNewLine
    SSQL = SSQL & "     Else" & vbNewLine
    SSQL = SSQL & "         COMMIT TRAN" & vbNewLine
    SSQL = SSQL & " "
    ADOCon.Execute SSQL
    
    SSQL = " "
    SSQL = SSQL & " CREATE    PROCEDURE SP_05014_01" & vbNewLine
    SSQL = SSQL & "     @INIT_FLAG          Char(1)," & vbNewLine
    SSQL = SSQL & "     @sDate1             Char(6)" & vbNewLine
    SSQL = SSQL & " AS" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & " BEGIN TRAN" & vbNewLine
    SSQL = SSQL & "     IF @INIT_FLAG   =   '1'" & vbNewLine
    SSQL = SSQL & "     BEGIN" & vbNewLine
    SSQL = SSQL & "         SELECT  ''          '코드'," & vbNewLine
    SSQL = SSQL & "             ''          '대리점명'," & vbNewLine
    SSQL = SSQL & "             ''          '수량'" & vbNewLine
    SSQL = SSQL & "     End" & vbNewLine
    SSQL = SSQL & "     Else" & vbNewLine
    SSQL = SSQL & "     BEGIN" & vbNewLine
    SSQL = SSQL & "         SELECT  A.AgencyCode    '코드'," & vbNewLine
    SSQL = SSQL & "             A.AgencyName    '대리점명'," & vbNewLine
    SSQL = SSQL & "             C.cnt '수량'" & vbNewLine
    SSQL = SSQL & "         FROM    AgencyCT    A (NOLOCK)," & vbNewLine
    SSQL = SSQL & "             (   SELECT StoreCode, COUNT(Number) AS Cnt, StoreTag FROM CouponUseData  (NOLOCK)" & vbNewLine
    SSQL = SSQL & "                 WHERE SUBSTRING(SaleDate,1,6) =   @sDate1" & vbNewLine
    SSQL = SSQL & "                 GROUP BY StoreCode, StoreTag    ) C" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "         Where A.AgencyCode = C.StoreTag" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "     End" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "     IF  @@error = 0" & vbNewLine
    SSQL = SSQL & "         COMMIT TRAN" & vbNewLine
    SSQL = SSQL & "     Else" & vbNewLine
    SSQL = SSQL & "         ROLLBACK TRAN" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    ADOCon.Execute SSQL


    SSQL = " " & vbNewLine
    SSQL = SSQL & " CREATE     PROCEDURE SP_05014_02" & vbNewLine
    SSQL = SSQL & "     @INIT_FLAG          Char(1)," & vbNewLine
    SSQL = SSQL & "     @sCode              Char(3)," & vbNewLine
    SSQL = SSQL & "     @sDate1             Char(6)" & vbNewLine
    SSQL = SSQL & " AS" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & " BEGIN TRAN" & vbNewLine
    SSQL = SSQL & "     IF @INIT_FLAG   =   '1'" & vbNewLine
    SSQL = SSQL & "     BEGIN" & vbNewLine
    SSQL = SSQL & "         SELECT  ''          '일자'," & vbNewLine
    SSQL = SSQL & "             ''          '수량'," & vbNewLine
    SSQL = SSQL & "             ''          ' '" & vbNewLine
    SSQL = SSQL & "     End" & vbNewLine
    SSQL = SSQL & "     Else" & vbNewLine
    SSQL = SSQL & "     BEGIN" & vbNewLine
    SSQL = SSQL & "         SELECT  SUBSTRING(C.SaleDate,1,4) + ' - ' + SUBSTRING(C.SaleDate,5,2) + ' - ' + SUBSTRING(C.SaleDate,7,2)   '일자'," & vbNewLine
    SSQL = SSQL & "             C.Cnt   '수량'," & vbNewLine
    SSQL = SSQL & "             ''      ' '" & vbNewLine
    SSQL = SSQL & "         FROM    AgencyCT    A (NOLOCK)," & vbNewLine
    SSQL = SSQL & "             (   SELECT SaleDate, COUNT(Number) AS Cnt, StoreTag FROM CouponUseData   (NOLOCK)" & vbNewLine
    SSQL = SSQL & "                 WHERE SUBSTRING(SaleDate,1,6) =   @sDate1" & vbNewLine
    SSQL = SSQL & "                     AND  StoreTag = @sCode" & vbNewLine
    SSQL = SSQL & "                 GROUP BY SaleDate, StoreTag    ) C" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "         Where A.AgencyCode = C.StoreTag" & vbNewLine
    SSQL = SSQL & "     End" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "     IF  @@error = 0" & vbNewLine
    SSQL = SSQL & "         COMMIT TRAN" & vbNewLine
    SSQL = SSQL & "     Else" & vbNewLine
    SSQL = SSQL & "         ROLLBACK TRAN" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    ADOCon.Execute SSQL

    SSQL = " " & vbNewLine
    SSQL = SSQL & " CREATE     PROCEDURE SP_05014_03" & vbNewLine
    SSQL = SSQL & "     @INIT_FLAG          Char(1)," & vbNewLine
    SSQL = SSQL & "     @sCode              Char(3)," & vbNewLine
    SSQL = SSQL & "     @sDate1             Char(8)" & vbNewLine
    SSQL = SSQL & " AS" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & " BEGIN TRAN" & vbNewLine
    SSQL = SSQL & "     IF @INIT_FLAG   =   '1'" & vbNewLine
    SSQL = SSQL & "     BEGIN" & vbNewLine
    SSQL = SSQL & "         SELECT  ''          '쿠폰번호'," & vbNewLine
    SSQL = SSQL & "             ''          '성명'," & vbNewLine
    SSQL = SSQL & "             ''          ' 금액'" & vbNewLine
    SSQL = SSQL & "     End" & vbNewLine
    SSQL = SSQL & "     Else" & vbNewLine
    SSQL = SSQL & "     BEGIN" & vbNewLine
    SSQL = SSQL & "             SELECT Number       '쿠폰번호'," & vbNewLine
    SSQL = SSQL & "                 CustName    '성명'," & vbNewLine
    SSQL = SSQL & "                 SaleMoney '금액'" & vbNewLine
    SSQL = SSQL & "             From CouponUseData(NOLOCK)" & vbNewLine
    SSQL = SSQL & "             WHERE SUBSTRING(SaleDate,1,8) =   @sDate1" & vbNewLine
    SSQL = SSQL & "                 AND  StoreTag = @sCode" & vbNewLine
    SSQL = SSQL & "     End" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "     IF  @@error = 0" & vbNewLine
    SSQL = SSQL & "         COMMIT TRAN" & vbNewLine
    SSQL = SSQL & "     Else" & vbNewLine
    SSQL = SSQL & "         ROLLBACK TRAN" & vbNewLine
    SSQL = SSQL & " "
    ADOCon.Execute SSQL


    On Error GoTo 0
    Exit Sub

Error_Rtn:
    Debug.Print Err.Description
    Resume Next

End Sub

