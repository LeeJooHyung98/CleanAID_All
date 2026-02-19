VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDB_Update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------
' 업데이트가 있을 경우 Update_Check을 수정한다.
'-------------------------------------------------------------------

Dim MyADOCon        As ADODB.Connection

Event Error(ByVal Number As Long, description As String)

Public Sub SetDef(ByRef objDB As Object)
    Set MyADOCon = objDB
End Sub

Public Function Update_SQL_Check() As Boolean
    Dim sDate   As String
    
    Update_SQL_Check = False
    ' 업데이트 테이블을 만들지 못하였거나 없을 경우
    If DB_Check = False Then Exit Function
    
    ' 해당 일자에 수정한 내용이다
    Call Update_20120217_01
     
    Call Update_20120403_01
    
    Call Update_20120419_01
    
    Call Update_20120425_01
    
    Call Update_20120427_01
    
    Call Update_20120507_01
    
    Call Update_20120511_01
    
    Call Update_20120619_01
    
    Call Update_20120622_01
    
    Call Update_20121130_01
    
    Call Update_20130213_01
    Call Update_20130213_02
    Call Update_20130213_03
    
    
    Call Update_20130430_01
    
    Call Update_20131030_01
    
    
    Call Update_20140128_01
    Call Update_20140128_02
    Call Update_20140128_03
    
    
'    Call Update_20150506_01
    Call Update_20150515_01
    
    Call Update_20150617_01
    
    Call Update_20150714_01
    
    Call Update_20150825_01
    Call Update_20150826_01
    
    Call Update_20160524_01
    
    Call Update_20160617_01
    
    Call Update_20160927_01
    Call Update_20160927_02
    
    Call Update_20160928_01
    
    
    Call Update_20170201_01
    Call Update_20170201_02
    
    Call Update_20170310_01
    
    Call Update_20190123_01
    
    Call Update_20190130_01
    
    Call Update_20190620_01
    
    Call Update_20190624_01
    
    Call Update_20210126_01
    
    Call Update_20210126_02
    
    Update_SQL_Check = True
End Function

Private Function DB_Check() As Boolean
    Dim SSQL    As String
    Dim SUBRs   As ADODB.RecordSet
    Dim nCount  As Integer
    
    nCount = 0
    On Error GoTo DB_Check_Error

SUB_Start:
    DB_Check = False

    nCount = nCount + 1
    SSQL = " SELECT * FROM Update_TB "
    
    Set SUBRs = New ADODB.RecordSet
    SUBRs.Open SSQL, MyADOCon, adOpenStatic, adLockOptimistic
    SUBRs.Close
    
    DB_Check = True
    On Error GoTo 0
    Exit Function

DB_Check_Error:
    ' Update_TB 테이블이 없을 경우
    If Err.Number = -2147217865 Then
        SSQL = "CREATE TABLE Update_TB "
        SSQL = SSQL & "(Gubun Char(11) Not Null, "
        SSQL = SSQL & "ActionDate Char(8)   Null, "
        SSQL = SSQL & "ActionMemo Char(50)  Null) "
        
        MyADOCon.Execute SSQL
        
        If nCount <= 3 Then GoSub SUB_Start
    End If
    
    DB_Check = False
    RaiseEvent Error(Err.Number, "CDB_Update DB_Check " & Err.description)

End Function

'--------------------------------------------------------------------------------------------------------------
' Procedure : GetUpdateInfo
' DateTime  : 2006-11-19 14:48
' Author    : pds2004
' Purpose   : 해당 정보가 있을 경우는 이미 업데이트된 내역이다.
'--------------------------------------------------------------------------------------------------------------
Private Function GetUpdateInfo(ByVal sGubun As String) As Boolean
    Dim SSQL    As String
    Dim SUBRs   As ADODB.RecordSet

    On Error GoTo GetUpdateInfo_Error
    GetUpdateInfo = False
    
    SSQL = " SELECT * FROM Update_TB "
    SSQL = SSQL & " WHERE Gubun = '" & sGubun & "' "
    
    Set SUBRs = New ADODB.RecordSet
    SUBRs.Open SSQL, MyADOCon, adOpenStatic, adLockOptimistic
    If SUBRs.RecordCount <= 0 Then
        GetUpdateInfo = True
    End If
    SUBRs.Close

    On Error GoTo 0
    Exit Function

GetUpdateInfo_Error:
    GetUpdateInfo = False
    RaiseEvent Error(Err.Number, "CDB_Update GetUpdateInfo " & Err.description)
    
End Function


'--------------------------------------------------------------------------------------------------------------
' Procedure : GetUpdateInfo
' DateTime  : 2006-11-19 14:48
' Author    : pds2004
' Purpose   : 해당 정보가 있을 경우는 이미 업데이트된 내역이다.
'--------------------------------------------------------------------------------------------------------------
Private Function GetColumnInfo(ByVal Table As String, ByVal Column As String) As Boolean
    Dim SSQL    As String
    Dim SUBRs   As ADODB.RecordSet

    On Error GoTo GetUpdateInfo_Error
    GetColumnInfo = False
    
    SSQL = " SELECT *"
    SSQL = SSQL & " From INFORMATION_SCHEMA.Columns"
    SSQL = SSQL & " WHERE TABLE_NAME = '" & Table & "' and COLUMN_NAME = '" & Column & "' "
    
    
    Set SUBRs = New ADODB.RecordSet
    SUBRs.Open SSQL, MyADOCon, adOpenStatic, adLockOptimistic
    If SUBRs.RecordCount <= 0 Then
        GetColumnInfo = True
    End If
    SUBRs.Close

    On Error GoTo 0
    Exit Function

GetUpdateInfo_Error:
    GetColumnInfo = False
    RaiseEvent Error(Err.Number, "CDB_Update GetUpdateInfo " & Err.description)
    
End Function

Private Function SaveUpdateInfo(ByVal sGubun As String) As Boolean
    Dim SSQL    As String
    
    ' 업데이트가 완료
    SSQL = "INSERT INTO Update_TB (Gubun, ActionDate, ActionMemo )"
    SSQL = SSQL & " VALUES ('" & sGubun & "', '" & Format(Date, "yyyymmdd") & "','" & Now & "') "
    MyADOCon.Execute SSQL

End Function

Private Function Update_20120217_01() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20120217_01 = False
    sUpdateKey = "20120217_01"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    SSQL = "ALTER TABLE TB_일일마감 ADD 마감시간 datetime "
    MyADOCon.Execute SSQL
    
    ' 업데이트가 완료
    Update_20120217_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20120217_01 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20120217_01 " & Err.description)

End Function
 
Private Function Update_20120403_01() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20120403_01 = False
    sUpdateKey = "20120403_01"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    SSQL = " "
    
    SSQL = SSQL & " IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[SP_TAG]') AND type in (N'P', N'PC'))" & vbNewLine
    SSQL = SSQL & " DROP PROCEDURE [dbo].[SP_TAG]" & vbNewLine
    MyADOCon.Execute SSQL


    SSQL = "  CREATE  PROCEDURE [dbo].[SP_TAG]" & vbNewLine
    SSQL = SSQL & "          @Cnt          INT                 -- 접수수량" & vbNewLine
    SSQL = SSQL & "        , @ReceiptDate  VARCHAR(10)         -- 접수일자" & vbNewLine
    SSQL = SSQL & "        , @ReceiptTime  VARCHAR(10)         -- 접수시간" & vbNewLine
    SSQL = SSQL & "        , @TagCode      VARCHAR(3)          -- 택코드" & vbNewLine
    SSQL = SSQL & "        , @BranchCode   VARCHAR(4)          -- 지사코드" & vbNewLine
    SSQL = SSQL & "        , @AgencyCode   VARCHAR(6)          -- 가맹점코드" & vbNewLine
    SSQL = SSQL & "        , @Rtn          VARCHAR(5000) OUTPUT" & vbNewLine
    SSQL = SSQL & " AS" & vbNewLine
    SSQL = SSQL & " BEGIN" & vbNewLine
    SSQL = SSQL & "    BEGIN TRANSACTION" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "    DECLARE @PadChar        CHAR(1)" & vbNewLine
    SSQL = SSQL & "    DECLARE @FinalLength    INTEGER" & vbNewLine
    SSQL = SSQL & "    DECLARE @TAG_List       VARCHAR(5000)" & vbNewLine
    SSQL = SSQL & "    DECLARE @Receipt_No     INT" & vbNewLine
    SSQL = SSQL & "    DECLARE @I              INT" & vbNewLine
    SSQL = SSQL & "    DECLARE @Tmp_TAG        VARCHAR(9)" & vbNewLine
    SSQL = SSQL & "    DECLARE @Tmp_No         INT" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "    SET @PadChar      = '0'" & vbNewLine
    SSQL = SSQL & "    SET @FinalLength  = 6" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "    -- 택번호, 접수번호" & vbNewLine
    SSQL = SSQL & "    SELECT @Tmp_TAG = ISNULL(택번호, @TagCode + '000001')" & vbNewLine
    SSQL = SSQL & "         , @Receipt_No = ISNULL(접수번호, 0) + 1" & vbNewLine
    SSQL = SSQL & "    FROM TB_기본정보" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "    SET @TAG_List = CONVERT(VARCHAR(10), @Receipt_No) + ':'" & vbNewLine
    SSQL = SSQL & "    SET @I = 1" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "    WHILE @I <= @Cnt" & vbNewLine
    SSQL = SSQL & "        BEGIN" & vbNewLine
    SSQL = SSQL & "            IF SUBSTRING(@Tmp_TAG,4,6) = '999999'" & vbNewLine
    SSQL = SSQL & "                BEGIN" & vbNewLine
    SSQL = SSQL & "                    SET @Tmp_TAG = @TagCode + '000000'" & vbNewLine
    SSQL = SSQL & "                END" & vbNewLine
    SSQL = SSQL & "            Else" & vbNewLine
    SSQL = SSQL & "                BEGIN" & vbNewLine
    SSQL = SSQL & "                    SET @Tmp_No = CONVERT(INT, SUBSTRING(@Tmp_TAG,4,6)) + 1" & vbNewLine
    SSQL = SSQL & "                    SET @Tmp_TAG = @TagCode + RIGHT( REPLICATE(@PadChar, @FinalLength - LEN(@FinalLength)) + CONVERT(VARCHAR(6), @Tmp_No), 6)" & vbNewLine
    SSQL = SSQL & "            END" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "            INSERT INTO TB_입출고(지사코드" & vbNewLine
    SSQL = SSQL & "                                , 가맹점코드" & vbNewLine
    SSQL = SSQL & "                                , 접수번호" & vbNewLine
    SSQL = SSQL & "                                , 택번호" & vbNewLine
    SSQL = SSQL & "                                , 접수일자" & vbNewLine
    SSQL = SSQL & "                                , 접수시간" & vbNewLine
    SSQL = SSQL & "                                ) VALUES (" & vbNewLine
    SSQL = SSQL & "                                  @BranchCode  -- 지사코드" & vbNewLine
    SSQL = SSQL & "                                , @AgencyCode  -- 가맹점코드" & vbNewLine
    SSQL = SSQL & "                                , @Receipt_No    -- 접수번호" & vbNewLine
    SSQL = SSQL & "                                , @Tmp_TAG     -- 택번호" & vbNewLine
    SSQL = SSQL & "                                , @ReceiptDate -- 접수일자" & vbNewLine
    SSQL = SSQL & "                                , @ReceiptTime -- 접수시간" & vbNewLine
    SSQL = SSQL & "                                )" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "            SET @TAG_List = @TAG_List + @Tmp_TAG + ','" & vbNewLine
    SSQL = SSQL & "            SET @I = @I + 1" & vbNewLine
    SSQL = SSQL & "    END" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "    IF @@ERROR > 0" & vbNewLine
    SSQL = SSQL & "        BEGIN" & vbNewLine
    SSQL = SSQL & "            SET @Rtn = ''" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "            Rollback TRANSACTION" & vbNewLine
    SSQL = SSQL & "        END" & vbNewLine
    SSQL = SSQL & "    Else" & vbNewLine
    SSQL = SSQL & "        BEGIN" & vbNewLine
    SSQL = SSQL & "            -- 마지막택번호, 접수번호 저장" & vbNewLine
    SSQL = SSQL & "            UPDATE TB_기본정보 SET 택번호   = @Tmp_TAG" & vbNewLine
    SSQL = SSQL & "                                 , 접수번호 = @Receipt_No" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "            SET @Rtn = @TAG_List" & vbNewLine
    SSQL = SSQL & "" & vbNewLine
    SSQL = SSQL & "            Commit TRANSACTION" & vbNewLine
    SSQL = SSQL & "        END" & vbNewLine
    SSQL = SSQL & "END"
    
    MyADOCon.Execute SSQL
    
    ' 업데이트가 완료
    Update_20120403_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20120403_01 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20120403_01 " & Err.description)

End Function

Private Function Update_20120419_01() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20120419_01 = False
    sUpdateKey = "20120419_01"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    SSQL = "ALTER TABLE TB_사고품내역 ADD 가맹점명 VARCHAR(50) "
    MyADOCon.Execute SSQL
    
    ' 업데이트가 완료
    Update_20120419_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20120419_01 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20120419_01 " & Err.description)

End Function
 

Private Function Update_20120425_01() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20120425_01 = False
    sUpdateKey = "20120425_01"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    SSQL = ""
    SSQL = SSQL & " CREATE NONCLUSTERED INDEX [IDX_일일판매현황] ON [dbo].[TB_매출]" & vbNewLine
    SSQL = SSQL & "(" & vbNewLine
    SSQL = SSQL & "    [매출일자] ASC," & vbNewLine
    SSQL = SSQL & "    [적요] Asc" & vbNewLine
    SSQL = SSQL & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]" & vbNewLine
    MyADOCon.Execute SSQL
    
    SSQL = ""
    SSQL = SSQL & "CREATE NONCLUSTERED INDEX [IDX_일일판매현황] ON [dbo].[TB_입출고]" & vbNewLine
    SSQL = SSQL & "(" & vbNewLine
    SSQL = SSQL & "    [출고일자] ASC," & vbNewLine
    SSQL = SSQL & "    [판매취소] Asc" & vbNewLine
    SSQL = SSQL & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]" & vbNewLine
    MyADOCon.Execute SSQL


    SSQL = ""
    SSQL = SSQL & "CREATE NONCLUSTERED INDEX [IDX_일일판매현황2] ON [dbo].[TB_입출고]" & vbNewLine
    SSQL = SSQL & "(" & vbNewLine
    SSQL = SSQL & "    [판매취소일자] ASC," & vbNewLine
    SSQL = SSQL & "    [반품환불일자] ASC," & vbNewLine
    SSQL = SSQL & "    [세탁환불일자] Asc" & vbNewLine
    SSQL = SSQL & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]" & vbNewLine
    MyADOCon.Execute SSQL
    
    ' 업데이트가 완료
    Update_20120425_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20120425_01 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20120425_01 " & Err.description)

End Function


Private Function Update_20120427_01() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20120427_01 = False
    sUpdateKey = "20120427_01"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    SSQL = ""
    SSQL = SSQL & "CREATE NONCLUSTERED INDEX [IDX_지점입고2] ON [dbo].[TB_입출고]" & vbNewLine
    SSQL = SSQL & "(" & vbNewLine
    SSQL = SSQL & "    [접수일자] ASC," & vbNewLine
    SSQL = SSQL & "    [택번호] ASC," & vbNewLine
    SSQL = SSQL & "    [판매취소] Asc" & vbNewLine
    SSQL = SSQL & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]" & vbNewLine
    MyADOCon.Execute SSQL
   
     
    ' 업데이트가 완료
    Update_20120427_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20120427_01 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20120427_01 " & Err.description)

End Function


Private Function Update_20120507_01() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20120507_01 = False
    sUpdateKey = "20120507_01"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
  
    SSQL = " "
    
    SSQL = SSQL & " IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[SP_DB_복구]') AND type in (N'P', N'PC'))" & vbNewLine
    SSQL = SSQL & " DROP PROCEDURE [dbo].[SP_DB_복구]" & vbNewLine
    MyADOCon.Execute SSQL


    SSQL = "  "
    SSQL = SSQL & " CREATE PROCEDURE [dbo].[SP_DB_복구]" & vbNewLine
    SSQL = SSQL & "           @pw   nvarchar(10) = ''" & vbNewLine
    SSQL = SSQL & " AS" & vbNewLine
    SSQL = SSQL & " Begin" & vbNewLine

    SSQL = SSQL & " /*" & vbNewLine
    SSQL = SSQL & "     exec [SP_DB_복구] 'pds2004'" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & " */" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "     if @pw <> 'pds2004'" & vbNewLine
    SSQL = SSQL & "         Return" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "     select 'ok'" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "  /*" & vbNewLine
    SSQL = SSQL & "    ----------------------------------------------------------------------" & vbNewLine
    SSQL = SSQL & "    -- 반드시 c:\CleanAid\Sql_Data 폴더의 내용을 백업후 실행 할것" & vbNewLine
    SSQL = SSQL & "    ----------------------------------------------------------------------" & vbNewLine
    SSQL = SSQL & "     -- 1. 단일 사용자 처리" & vbNewLine
    SSQL = SSQL & "     -- DB의 상태를 단일 사용자로 변경 한다." & vbNewLine
    SSQL = SSQL & "     -- DB 복구를 위하여 가장 먼저 실행 하여야 한다." & vbNewLine
    SSQL = SSQL & "     -- sp_dboption 'CleanAid', 'Single user', true" & vbNewLine
    SSQL = SSQL & "     alter database CleanAid set single_user" & vbNewLine
    SSQL = SSQL & "     go" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "     -- 2. DB의 상태를 확인한다." & vbNewLine
    SSQL = SSQL & "     -- 불은색으로 표시되는 부분이 있으면 오류가 발생한 부분으로 해당 자료를 복구 하여 주어야 한다." & vbNewLine
    SSQL = SSQL & "     DBCC checkdb('CleanAid')" & vbNewLine
    SSQL = SSQL & "     go" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "     -- 2-1. DB의 모든 상태를 확인 후 복구 처리한다." & vbNewLine
    SSQL = SSQL & "     -- 불은색으로 표시되는 부분이 있으면 오류가 발생한 부분으로 해당 자료를 복구 하여 주어야 한다." & vbNewLine
    SSQL = SSQL & "     -- 자료 손실이 발생할 수 있다." & vbNewLine
    SSQL = SSQL & "     dbcc checkdb(CleanAID, REPAIR_ALLOW_DATA_LOSS)" & vbNewLine
    SSQL = SSQL & "     go" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "     -- 2-2. 특정한 테이블을 복구 한다." & vbNewLine
    SSQL = SSQL & "     -- 자료 손실이 발생할 수 있다." & vbNewLine
    SSQL = SSQL & "     --dbcc checktable (charitem, REPAIR_ALLOW_DATA_LOSS )" & vbNewLine
    SSQL = SSQL & "     --go" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "     -- 1. 멀티 사용자로 복구" & vbNewLine
    SSQL = SSQL & "     -- DB의 상태를 단일 사용자로 변경 한다." & vbNewLine
    SSQL = SSQL & "     -- DB 복구를 위하여 가장 먼저 실행 하여야 한다." & vbNewLine
    SSQL = SSQL & "     -- sp_dboption 'CleanAid', 'Single user', false" & vbNewLine
    SSQL = SSQL & "     alter database CleanAid set multi_user" & vbNewLine
    SSQL = SSQL & "     go" & vbNewLine
    SSQL = SSQL & " */" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & " End" & vbNewLine
    MyADOCon.Execute SSQL
    
    
    SSQL = " "
    SSQL = SSQL & " IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[SP_DB_용량줄이기]') AND type in (N'P', N'PC'))" & vbNewLine
    SSQL = SSQL & " DROP PROCEDURE [dbo].[SP_DB_용량줄이기]" & vbNewLine
    MyADOCon.Execute SSQL


    SSQL = "  " & vbNewLine
    SSQL = SSQL & "CREATE PROCEDURE [dbo].[SP_DB_용량줄이기]" & vbNewLine
    SSQL = SSQL & "          @pw   nvarchar(10) = ''" & vbNewLine
    SSQL = SSQL & "AS" & vbNewLine
    SSQL = SSQL & "Begin" & vbNewLine
    SSQL = SSQL & "/*" & vbNewLine
    SSQL = SSQL & "    EXEC [SP_DB_용량줄이기] 'pds2004'" & vbNewLine
    SSQL = SSQL & "*/" & vbNewLine
    SSQL = SSQL & "" & vbNewLine
    SSQL = SSQL & "    if @pw <> 'pds2004'" & vbNewLine
    SSQL = SSQL & "        Return" & vbNewLine
    SSQL = SSQL & "" & vbNewLine
    SSQL = SSQL & "    select 'ok'" & vbNewLine
    SSQL = SSQL & "    ----------------------------------------------------------------------" & vbNewLine
    SSQL = SSQL & "    -- 반드시 c:\CleanAid\Sql_Data 폴더의 내용을 백업후 실행 할것" & vbNewLine
    SSQL = SSQL & "    ----------------------------------------------------------------------" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "    -- 1. 파일 용량 확인 하기" & vbNewLine
    SSQL = SSQL & "     EXEC sp_helpfile" & vbNewLine
    SSQL = SSQL & "" & vbNewLine
    SSQL = SSQL & "    -- 2. 로그 비우기" & vbNewLine
    SSQL = SSQL & "    BACKUP LOG CleanAid WITH NO_LOG" & vbNewLine
    SSQL = SSQL & "" & vbNewLine
    SSQL = SSQL & "    -- 3. 로그파일 줄이기 (1200은 메가 단위)" & vbNewLine
    SSQL = SSQL & "    DBCC SHRINKFILE(CleanAid_log, 10)" & vbNewLine
    SSQL = SSQL & "" & vbNewLine
    SSQL = SSQL & "End" & vbNewLine
    MyADOCon.Execute SSQL
 
 
    SSQL = " "
    SSQL = SSQL & " IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[SP_DB_주의대상복구]') AND type in (N'P', N'PC'))" & vbNewLine
    SSQL = SSQL & " DROP PROCEDURE [dbo].[SP_DB_주의대상복구]" & vbNewLine
    MyADOCon.Execute SSQL
    
    SSQL = "  " & vbNewLine
    SSQL = SSQL & "CREATE PROCEDURE [dbo].[SP_DB_주의대상복구]" & vbNewLine
    SSQL = SSQL & "          @pw   nvarchar(10) = ''" & vbNewLine
    SSQL = SSQL & "AS" & vbNewLine
    SSQL = SSQL & "Begin" & vbNewLine
    SSQL = SSQL & "/*" & vbNewLine
    SSQL = SSQL & "    EXEC [SP_DB_용량줄이기] 'pds2004'" & vbNewLine
    SSQL = SSQL & "*/" & vbNewLine
    SSQL = SSQL & "" & vbNewLine
    SSQL = SSQL & "    if @pw <> 'pds2004'" & vbNewLine
    SSQL = SSQL & "        Return" & vbNewLine
    SSQL = SSQL & "" & vbNewLine
    SSQL = SSQL & "    select 'ok'" & vbNewLine
    SSQL = SSQL & "    ----------------------------------------------------------------------" & vbNewLine
    SSQL = SSQL & "    -- 반드시 c:\CleanAid\Sql_Data 폴더의 내용을 백업후 실행 할것" & vbNewLine
    SSQL = SSQL & "    ----------------------------------------------------------------------" & vbNewLine
    SSQL = SSQL & "" & vbNewLine
    SSQL = SSQL & "    EXEC sp_resetstatus 'CleanAid';" & vbNewLine
    SSQL = SSQL & "    ALTER DATABASE CleanAid SET EMERGENCY" & vbNewLine
    SSQL = SSQL & "    DBCC checkdb('CleanAid')" & vbNewLine
    SSQL = SSQL & "    ALTER DATABASE CleanAid SET SINGLE_USER WITH ROLLBACK IMMEDIATE" & vbNewLine
    SSQL = SSQL & "    DBCC CheckDB ('CleanAid', REPAIR_ALLOW_DATA_LOSS)" & vbNewLine
    SSQL = SSQL & "    ALTER DATABASE CleanAid SET MULTI_USER" & vbNewLine
    SSQL = SSQL & "" & vbNewLine
    SSQL = SSQL & "End" & vbNewLine
    MyADOCon.Execute SSQL



    ' 업데이트가 완료
    Update_20120507_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20120507_01 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20120507_01 " & Err.description)

End Function

Private Function Update_20120511_01() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20120511_01 = False
    sUpdateKey = "20120511_01"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    SSQL = ""
    SSQL = SSQL & "UPDATE TB_기본정보 SET SMS_IP = '115.89.220.5,8657', SMS_DB = 'LAUNDRY1000', SMS_PWD = 'cleanaid1996!@#'"
    MyADOCon.Execute SSQL
   
     
    ' 업데이트가 완료
    Update_20120511_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20120511_01 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20120511_01 " & Err.description)

End Function

Private Function Update_20120619_01() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20120619_01 = False
    sUpdateKey = "20120619_01"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    SSQL = "ALTER TABLE TB_현금영수증 ADD 취소사유 VARCHAR(02) "
    MyADOCon.Execute SSQL
    
    SSQL = "UPDATE  TB_현금영수증 SET 취소사유 = '' "
    MyADOCon.Execute SSQL
    
   
     
    ' 업데이트가 완료
    Update_20120619_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20120619_01 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20120619_01 " & Err.description)

End Function

Private Function Update_20120622_01() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20120622_01 = False
    sUpdateKey = "20120622_01"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    SSQL = "ALTER TABLE TB_작업 ADD 수정일자 nVARCHAR(10) "
    MyADOCon.Execute SSQL
    
    SSQL = "UPDATE  TB_작업 SET 수정일자 = '2012-01-01' "
    MyADOCon.Execute SSQL
    
    ' 업데이트가 완료
    Update_20120622_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20120622_01 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20120622_01 " & Err.description)

End Function


Private Function Update_20121130_01() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20121130_01 = False
    sUpdateKey = "20121130_01"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    SSQL = "update TB_의류 set 의류명 = '롱코트' where 의류코드 = 'i002'  ":            MyADOCon.Execute SSQL
    SSQL = "update TB_의류 set 의류명 = '알파카코트' where 의류코드 = 'i023'  ":        MyADOCon.Execute SSQL
    SSQL = "update TB_의류 set 의류명 = '알파카반코트' where 의류코드 = 'i024'  ":      MyADOCon.Execute SSQL
    SSQL = "update TB_의류 set 의류명 = '캐시미어코트' where 의류코드 = 'i025'  ":      MyADOCon.Execute SSQL
    SSQL = "update TB_의류 set 의류명 = '캐시미어반코트' where 의류코드 = 'i026'  ":    MyADOCon.Execute SSQL
    SSQL = "update TB_의류 set 의류명 = '벨벳코트' where 의류코드 = 'i027'  ":          MyADOCon.Execute SSQL
    SSQL = "update TB_의류 set 의류명 = '벨벳반코트' where 의류코드 = 'i028'  ":        MyADOCon.Execute SSQL
    
    SSQL = "update TB_요일할인 set 의류명 = '롱코트' where 의류코드 = 'i002'  ":            MyADOCon.Execute SSQL
    SSQL = "update TB_요일할인 set 의류명 = '알파카코트' where 의류코드 = 'i023'  ":        MyADOCon.Execute SSQL
    SSQL = "update TB_요일할인 set 의류명 = '알파카반코트' where 의류코드 = 'i024'  ":      MyADOCon.Execute SSQL
    SSQL = "update TB_요일할인 set 의류명 = '캐시미어코트' where 의류코드 = 'i025'  ":      MyADOCon.Execute SSQL
    SSQL = "update TB_요일할인 set 의류명 = '캐시미어반코트' where 의류코드 = 'i026'  ":    MyADOCon.Execute SSQL
    SSQL = "update TB_요일할인 set 의류명 = '벨벳코트' where 의류코드 = 'i027'  ":          MyADOCon.Execute SSQL
    SSQL = "update TB_요일할인 set 의류명 = '벨벳반코트' where 의류코드 = 'i028'  ":        MyADOCon.Execute SSQL
    
    SSQL = "update TB_할인정보 set 의류명 = '롱코트' where 의류코드 = 'i002'  ":            MyADOCon.Execute SSQL
    SSQL = "update TB_할인정보 set 의류명 = '알파카코트' where 의류코드 = 'i023'  ":        MyADOCon.Execute SSQL
    SSQL = "update TB_할인정보 set 의류명 = '알파카반코트' where 의류코드 = 'i024'  ":      MyADOCon.Execute SSQL
    SSQL = "update TB_할인정보 set 의류명 = '캐시미어코트' where 의류코드 = 'i025'  ":      MyADOCon.Execute SSQL
    SSQL = "update TB_할인정보 set 의류명 = '캐시미어반코트' where 의류코드 = 'i026'  ":    MyADOCon.Execute SSQL
    SSQL = "update TB_할인정보 set 의류명 = '벨벳코트' where 의류코드 = 'i027'  ":          MyADOCon.Execute SSQL
    SSQL = "update TB_할인정보 set 의류명 = '벨벳반코트' where 의류코드 = 'i028'  ":        MyADOCon.Execute SSQL
    
    
    ' 업데이트가 완료
    Update_20121130_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20121130_01 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20121130_01 " & Err.description)

End Function

Private Function Update_20130213_01() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20130213_01 = False
    sUpdateKey = "20130213_01"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    SSQL = ""
    SSQL = SSQL & "CREATE TABLE [dbo].[TB_사고품품목]" & vbNewLine
    SSQL = SSQL & "(" & vbNewLine
    SSQL = SSQL & "    [품목코드] [char](3) NOT NULL," & vbNewLine
    SSQL = SSQL & "    [품목명] [varchar](40) NULL," & vbNewLine
    SSQL = SSQL & "CONSTRAINT [PK_TB_사고품품목] PRIMARY KEY CLUSTERED" & vbNewLine
    SSQL = SSQL & "(" & vbNewLine
    SSQL = SSQL & "    [품목코드] ASC" & vbNewLine
    SSQL = SSQL & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]" & vbNewLine
    SSQL = SSQL & ") ON [PRIMARY] " & vbNewLine
    MyADOCon.Execute SSQL
   
     
    ' 업데이트가 완료
    Update_20130213_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20130213_01 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20130213_01 " & Err.description)

End Function


Private Function Update_20130213_02() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20130213_02 = False
    sUpdateKey = "20130213_02"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    SSQL = ""
    SSQL = SSQL & "CREATE TABLE [dbo].[TB_사고품용도]" & vbNewLine
    SSQL = SSQL & "(" & vbNewLine
    SSQL = SSQL & "    [용도코드] [char](3) NOT NULL," & vbNewLine
    SSQL = SSQL & "    [용도내용] [varchar](40) NULL," & vbNewLine
    SSQL = SSQL & "CONSTRAINT [PK_TB_사고품용도] PRIMARY KEY CLUSTERED" & vbNewLine
    SSQL = SSQL & "(" & vbNewLine
    SSQL = SSQL & "    [용도코드] ASC" & vbNewLine
    SSQL = SSQL & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]" & vbNewLine
    SSQL = SSQL & ") ON [PRIMARY] " & vbNewLine
    MyADOCon.Execute SSQL
   
     
    ' 업데이트가 완료
    Update_20130213_02 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20130213_02 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20130213_02 " & Err.description)

End Function


Private Function Update_20130213_03() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20130213_03 = False
    sUpdateKey = "20130213_03"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    SSQL = ""
    SSQL = SSQL & "CREATE TABLE [dbo].[TB_사고품소재]" & vbNewLine
    SSQL = SSQL & "(" & vbNewLine
    SSQL = SSQL & "    [소재코드] [char](3) NOT NULL," & vbNewLine
    SSQL = SSQL & "    [소재명] [varchar](40) NULL," & vbNewLine
    SSQL = SSQL & "CONSTRAINT [PK_TB_사고품소재] PRIMARY KEY CLUSTERED" & vbNewLine
    SSQL = SSQL & "(" & vbNewLine
    SSQL = SSQL & "    [소재코드] ASC" & vbNewLine
    SSQL = SSQL & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]" & vbNewLine
    SSQL = SSQL & ") ON [PRIMARY] " & vbNewLine
    MyADOCon.Execute SSQL
   
     
    ' 업데이트가 완료
    Update_20130213_03 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20130213_03 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20130213_03 " & Err.description)

End Function

Private Function Update_20130430_01() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20130430_01 = False
    sUpdateKey = "20130430_01"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    SSQL = "ALTER TABLE TB_공지사항 ADD 파일명 nVARCHAR(50) "
    MyADOCon.Execute SSQL
    
    SSQL = "UPDATE  TB_공지사항 SET 파일명 = '' "
    MyADOCon.Execute SSQL
    
    SSQL = ""
    SSQL = SSQL & "CREATE TABLE [dbo].[TB_공지사항파일](" & vbNewLine
    SSQL = SSQL & "    [작성일자] [nvarchar](10) NOT NULL," & vbNewLine
    SSQL = SSQL & "    [문서번호] [int] NOT NULL," & vbNewLine
    SSQL = SSQL & "    [문서파일] [image] NULL," & vbNewLine
    SSQL = SSQL & "    [파일크기] [decimal](10, 0) NULL," & vbNewLine
    SSQL = SSQL & " CONSTRAINT [PK_TB_공지사항파일] PRIMARY KEY CLUSTERED" & vbNewLine
    SSQL = SSQL & "(" & vbNewLine
    SSQL = SSQL & "    [작성일자] ASC," & vbNewLine
    SSQL = SSQL & "    [문서번호] Asc" & vbNewLine
    SSQL = SSQL & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]" & vbNewLine
    SSQL = SSQL & ") ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]" & vbNewLine
    MyADOCon.Execute SSQL
    
    ' 업데이트가 완료
    Update_20130430_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20130430_01 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20130430_01 " & Err.description)

End Function


Private Function Update_20131030_01() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20131030_01 = False
    sUpdateKey = "20131030_01"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
 
    SSQL = ""
    SSQL = SSQL & "CREATE TABLE [dbo].[TB_반품현황](" & vbNewLine
    SSQL = SSQL & "    [반품일자] [nvarchar](10) NOT NULL," & vbNewLine
    SSQL = SSQL & "    [차수] [nvarchar](1) NOT NULL," & vbNewLine
    SSQL = SSQL & "    [가맹점코드] [nvarchar](6) NOT NULL," & vbNewLine
    SSQL = SSQL & "    [택번호] [nvarchar](9) NOT NULL," & vbNewLine
    SSQL = SSQL & "    [접수일자] [nvarchar](10) NULL," & vbNewLine
    SSQL = SSQL & "    [반품사유] [nvarchar](100) NULL," & vbNewLine
    SSQL = SSQL & " CONSTRAINT [PK_TB_반품현황] PRIMARY KEY CLUSTERED" & vbNewLine
    SSQL = SSQL & "(" & vbNewLine
    SSQL = SSQL & "    [반품일자] ASC," & vbNewLine
    SSQL = SSQL & "    [차수] ASC," & vbNewLine
    SSQL = SSQL & "    [가맹점코드] Asc, " & vbNewLine
    SSQL = SSQL & "    [택번호] Asc" & vbNewLine
    SSQL = SSQL & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]" & vbNewLine
    SSQL = SSQL & ") ON [PRIMARY] " & vbNewLine
    MyADOCon.Execute SSQL
    
    ' 업데이트가 완료
    Update_20131030_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20131030_01 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20131030_01 " & Err.description)

End Function

Private Function Update_20140128_01() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20140128_01 = False
    sUpdateKey = "20140128_01"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    SSQL = "ALTER TABLE TB_기본정보 ADD 로열티여부1 nvarchar(1) ":              MyADOCon.Execute SSQL
    SSQL = "ALTER TABLE TB_기본정보 ADD 로열티비율1 nvarchar(10) ":             MyADOCon.Execute SSQL
    SSQL = "ALTER TABLE TB_기본정보 ADD 로열티여부2 nvarchar(1) ":              MyADOCon.Execute SSQL
    SSQL = "ALTER TABLE TB_기본정보 ADD 로열티비율2 nvarchar(10) ":             MyADOCon.Execute SSQL
    SSQL = "ALTER TABLE TB_기본정보 ADD 수수료지원여부 nvarchar(1) ":           MyADOCon.Execute SSQL
    SSQL = "ALTER TABLE TB_기본정보 ADD 수수료지원비율 nvarchar(10) ":          MyADOCon.Execute SSQL
    
    SSQL = "UPDATE  TB_기본정보 SET 로열티여부1 = 'N', 로열티비율1= '0', 로열티여부2 = 'N', 로열티비율2='0', 수수료지원여부 = 'N', 수수료지원비율 = '0' "
    MyADOCon.Execute SSQL
    
    ' 업데이트가 완료
    Update_20140128_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20140128_01 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20140128_01 " & Err.description)

End Function


Private Function Update_20140128_02() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20140128_02 = False
    sUpdateKey = "20140128_02"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    SSQL = "ALTER TABLE TB_신용카드승인 ADD 취소일자 nvarchar(30) ":              MyADOCon.Execute SSQL
    SSQL = "ALTER TABLE TB_신용카드승인 ADD 기타메모 nvarchar(50) ":              MyADOCon.Execute SSQL
    
    ' 업데이트가 완료
    Update_20140128_02 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20140128_02 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20140128_02 " & Err.description)

End Function

Private Function Update_20140128_03() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20140128_03 = False
    sUpdateKey = "20140128_03"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    SSQL = "ALTER TABLE tb_일일마감 ADD 로열티정보1 nvarchar(11) ":     MyADOCon.Execute SSQL
    SSQL = "ALTER TABLE tb_일일마감 ADD 로열티정보2 nvarchar(11) ":     MyADOCon.Execute SSQL
    SSQL = "ALTER TABLE tb_일일마감 ADD 수수료정보  nvarchar(11) ":     MyADOCon.Execute SSQL
    SSQL = "ALTER TABLE tb_일일마감 ADD 반품환불지사금액 int ":         MyADOCon.Execute SSQL
    SSQL = "ALTER TABLE tb_일일마감 ADD 세탁환불지사금액 int ":         MyADOCon.Execute SSQL
    
    SSQL = "ALTER TABLE tb_일일마감 ADD 미수카드건수 int ":             MyADOCon.Execute SSQL
    SSQL = "ALTER TABLE tb_일일마감 ADD 미수카드금액 int ":             MyADOCon.Execute SSQL
    
    SSQL = "ALTER TABLE tb_일일마감 ADD 카드취소건수 int ":             MyADOCon.Execute SSQL
    SSQL = "ALTER TABLE tb_일일마감 ADD 카드취소금액 int ":             MyADOCon.Execute SSQL
    
    SSQL = "ALTER TABLE tb_일일마감 ADD 로열티금액1 int ":              MyADOCon.Execute SSQL
    SSQL = "ALTER TABLE tb_일일마감 ADD 로열티금액2 int ":              MyADOCon.Execute SSQL
    SSQL = "ALTER TABLE tb_일일마감 ADD 수수료승인금액 int ":           MyADOCon.Execute SSQL
    SSQL = "ALTER TABLE tb_일일마감 ADD 수수료취소금액 int ":           MyADOCon.Execute SSQL
    
    ' 업데이트가 완료
    Update_20140128_03 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20140128_03 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20140128_03 " & Err.description)

End Function

'
'Private Function Update_20150506_01() As Boolean
'    Dim SSQL        As String
'    Dim sUpdateKey  As String
'
'    On Error GoTo Update_Error
'
'    Update_20150506_01 = False
'    sUpdateKey = "20150506_01"
'
'    ' 이미 업데이트 되었을 경우 작업하지 않느다.
'    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
'
'    Query = "ALTER TABLE dbo.TB_이용실적 alter CONSTRAINT PK_TB_이용실적 PRIMARY KEY (가맹점코드,고객코드,연도)"
'    ADOCon.Execute Query
'    DoEvents
'
'    ' 업데이트가 완료
'    Update_20150506_01 = SaveUpdateInfo(sUpdateKey)
'    On Error GoTo 0
'    Exit Function
'
'Update_Error:
'    Update_20150506_01 = False
'    RaiseEvent Error(Err.Number, "CDB_Update Update_20150506_01 " & Err.description)
'
'End Function


Private Function Update_20150515_01() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20150515_01 = False
    sUpdateKey = "20150515_01"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    SSQL = "ALTER TABLE TB_기본정보 ADD PG사용종료일 nvarchar(10) ":     MyADOCon.Execute SSQL
    
    ' 업데이트가 완료
    Update_20150515_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20150515_01 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20150515_01 " & Err.description)

End Function


Private Function Update_20150617_01() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20150617_01 = False
    sUpdateKey = "20150617_01"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    SSQL = ""
    SSQL = SSQL & "CREATE TABLE [dbo].[TB_지사]("
    SSQL = SSQL & "    [지사코드] [char](4) NOT NULL,"
    SSQL = SSQL & "        [지사명] [varchar](50) NULL,"
    SSQL = SSQL & "        [지사상태] [char](1) NULL,"
    SSQL = SSQL & "        [지사장명] [varchar](50) NULL,"
    SSQL = SSQL & "        [전화번호] [varchar](50) NULL,"
    SSQL = SSQL & "        [휴대전화] [varchar](50) NULL,"
    SSQL = SSQL & "        [주소] [varchar](100) NULL,"
    SSQL = SSQL & "        [메모] [varchar](100) NULL,"
    SSQL = SSQL & "        [영업팀] [nvarchar](2) NULL,"
    SSQL = SSQL & "     CONSTRAINT [PK_TB_지사] PRIMARY KEY CLUSTERED"
    SSQL = SSQL & "    ("
    SSQL = SSQL & "        [지사코드] Asc"
    SSQL = SSQL & "    )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
    SSQL = SSQL & "    ) ON [PRIMARY]"
        
    MyADOCon.Execute SSQL
    
    ' 업데이트가 완료
    Update_20150617_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20150617_01 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20150617_01 " & Err.description)

End Function


Private Function Update_20150714_01() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20150714_01 = False
    sUpdateKey = "20150714_01"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    
    SSQL = ""
    SSQL = SSQL & " CREATE NONCLUSTERED INDEX [IDX_SMS] ON [dbo].[TB_입출고] "
    SSQL = SSQL & " ("
    SSQL = SSQL & "     [접수일자] ASC,"
    SSQL = SSQL & "     [고객코드] ASC,"
    SSQL = SSQL & "     [출고일자] ASC,"
    SSQL = SSQL & "     [판매취소] ASC,"
    SSQL = SSQL & "     [가맹점입고일자] Asc"
    SSQL = SSQL & " )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
    
    MyADOCon.Execute SSQL
    
    ' 업데이트가 완료
    Update_20150714_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    If -2147217900 = Err.Number Then
        ' 업데이트가 완료
        Update_20150714_01 = SaveUpdateInfo(sUpdateKey)
        On Error GoTo 0
        Exit Function
    
    
    Else
        Update_20150714_01 = False
        RaiseEvent Error(Err.Number, "CDB_Update Update_20150714_01 " & Err.description)
    End If
End Function


Private Function Update_20150825_01() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20150825_01 = False
    sUpdateKey = "20150825_01"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    SSQL = "ALTER TABLE TB_기본정보 ADD CAT단말기종류 nvarchar(20) ":     MyADOCon.Execute SSQL
    
    ' 업데이트가 완료
    Update_20150825_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20150825_01 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20150825_01 " & Err.description)

End Function


Private Function Update_20150826_01() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20150826_01 = False
    sUpdateKey = "20150826_01"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    
    SSQL = ""
    SSQL = SSQL & " CREATE NONCLUSTERED INDEX [idx_cust] ON [dbo].[TB_입출고]"
    SSQL = SSQL & " ("
    SSQL = SSQL & "     [고객코드] Asc"
    SSQL = SSQL & " )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
    
    MyADOCon.Execute SSQL
    
    ' 업데이트가 완료
    Update_20150826_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    If -2147217900 = Err.Number Then
        ' 업데이트가 완료
        Update_20150826_01 = SaveUpdateInfo(sUpdateKey)
        On Error GoTo 0
        Exit Function
    
    
    Else
        Update_20150826_01 = False
        RaiseEvent Error(Err.Number, "CDB_Update Update_20150826_01 " & Err.description)
    End If
End Function



Private Function Update_20160524_01() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20160524_01 = False
    sUpdateKey = "20160524_01"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    SSQL = "ALTER TABLE tb_일일마감 ADD 미수현금수금금액 int ":             MyADOCon.Execute SSQL
    
    ' 업데이트가 완료
    Update_20160524_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20160524_01 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20160524_01 " & Err.description)

End Function



Private Function Update_20160617_01() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20160617_01 = False
    sUpdateKey = "20160617_01"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    SSQL = "/*++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++" & vbNewLine
    SSQL = SSQL & " -- 프로시저명 : [SP_SMS_VIEW1]" & vbNewLine
    SSQL = SSQL & " -- 파라미터명 : @INIT_FLAG  : 구분" & vbNewLine
    SSQL = SSQL & " --              @MASTER_CD  : 지사코드" & vbNewLine
    SSQL = SSQL & " -- 작  성  일 : 2016-06-17" & vbNewLine
    SSQL = SSQL & " -- 작  성  자 : pds2004" & vbNewLine
    SSQL = SSQL & " -- 프로그램명 :  " & vbNewLine
    SSQL = SSQL & " -- 설      명 : ." & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & " SP_SMS_VIEW1 '2016-06-10', '2016-06-16'" & vbNewLine
    SSQL = SSQL & "     select * from tb_입출고 where 접수일자 between '2016-06-01' and '2016-06-13' and 고객코드 = '050562'" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & " --++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++*/" & vbNewLine
    SSQL = SSQL & " CREATE PROCEDURE [dbo].[SP_SMS_VIEW1]" & vbNewLine
    SSQL = SSQL & "     @SDate          nvarchar(10)," & vbNewLine
    SSQL = SSQL & "     @EDate          nvarchar(10)" & vbNewLine
    SSQL = SSQL & " AS" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & " SET NOCOUNT ON" & vbNewLine
    SSQL = SSQL & " SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED " & vbNewLine
    SSQL = SSQL & "  " & vbNewLine
    SSQL = SSQL & "  " & vbNewLine
    SSQL = SSQL & " DECLARE @Temp TABLE  " & vbNewLine
    SSQL = SSQL & " ( " & vbNewLine
    SSQL = SSQL & "       고객코드  varchar(6), " & vbNewLine
    SSQL = SSQL & "       성명      varchar(50), " & vbNewLine
    SSQL = SSQL & "       전화번호  varchar(30), " & vbNewLine
    SSQL = SSQL & "       휴대전화  varchar(30), " & vbNewLine
    SSQL = SSQL & "       문자발송여부  varchar(1) " & vbNewLine
    SSQL = SSQL & " ) " & vbNewLine
    SSQL = SSQL & "  " & vbNewLine
    SSQL = SSQL & "  " & vbNewLine
    SSQL = SSQL & " -- 해당 기간의 접수자 정보를 구한다. " & vbNewLine
    SSQL = SSQL & " INSERT INTO @Temp " & vbNewLine
    SSQL = SSQL & " SELECT aa.고객코드, dd.성명, dd.전화번호, dd.휴대전화, dd.문자발송여부 " & vbNewLine
    SSQL = SSQL & " FROM TB_입출고 AA     " & vbNewLine
    SSQL = SSQL & "  " & vbNewLine
    SSQL = SSQL & "     left outer join TB_고객정보 DD         " & vbNewLine
    SSQL = SSQL & "     on AA.고객코드 = DD.고객코드     " & vbNewLine
    SSQL = SSQL & "  " & vbNewLine
    SSQL = SSQL & "   WHERE aa.접수일자 between @SDate   AND   @EDate    " & vbNewLine
    SSQL = SSQL & "   AND (aa.가맹점입고일자 <> '' OR aa.가맹점입고일자 IS NOT NULL)    " & vbNewLine
    SSQL = SSQL & "   AND (aa.출고일자 = '' OR aa.출고일자 IS NULL)   AND (aa.판매취소 <> 'Y')  " & vbNewLine
    SSQL = SSQL & " group by   aa.고객코드, dd.성명, dd.전화번호, dd.휴대전화, dd.문자발송여부 " & vbNewLine
    SSQL = SSQL & "  " & vbNewLine
    SSQL = SSQL & "  " & vbNewLine
    SSQL = SSQL & " SELECT aa.고객코드, aa.성명, aa.전화번호, aa.휴대전화, aa.문자발송여부,CC.TagS , CC.TagE, CC.TagCnt, EE.ReturnCnt    " & vbNewLine
    SSQL = SSQL & " FROM @Temp AA     " & vbNewLine
    SSQL = SSQL & "   " & vbNewLine
    SSQL = SSQL & "     -- 미출고 고객 코드를 구한다. " & vbNewLine
    SSQL = SSQL & "     left outer join ( " & vbNewLine
    SSQL = SSQL & "         select 고객코드 as 미출고고객코드 from  TB_입출고          " & vbNewLine
    SSQL = SSQL & "          where 접수일자  between @SDate   AND  @EDate         " & vbNewLine
    SSQL = SSQL & "            AND isnull(가맹점입고일자,'') = '' " & vbNewLine
    SSQL = SSQL & "            AND (판매취소 <> 'Y')         " & vbNewLine
    SSQL = SSQL & "            AND isnull(출고일자,'') = '' " & vbNewLine
    SSQL = SSQL & "          group by 고객코드 ) BB " & vbNewLine
    SSQL = SSQL & "      on AA.고객코드 = BB.미출고고객코드 " & vbNewLine
    SSQL = SSQL & "  " & vbNewLine
    SSQL = SSQL & "     -- 접수자의 추가 정보를 구한다. " & vbNewLine
    SSQL = SSQL & "     left outer join (             " & vbNewLine
    SSQL = SSQL & "         select 고객코드, MIN(택번호) as TagS, MAX(택번호) as TagE, COUNT(택번호) as TagCnt               " & vbNewLine
    SSQL = SSQL & "         From TB_입출고             " & vbNewLine
    SSQL = SSQL & "         where 접수일자  between @SDate   AND  @EDate             " & vbNewLine
    SSQL = SSQL & "         AND isnull(가맹점입고일자,'') <> '' " & vbNewLine
    SSQL = SSQL & "         AND 판매취소 <> 'Y'             " & vbNewLine
    SSQL = SSQL & "         AND isnull(출고일자,'') = ''       " & vbNewLine
    SSQL = SSQL & "         group by 고객코드 ) CC         " & vbNewLine
    SSQL = SSQL & "     on AA.고객코드 = CC.고객코드             " & vbNewLine
    SSQL = SSQL & "  " & vbNewLine
    SSQL = SSQL & "     -- 반품환불 수량을 구한다. " & vbNewLine
    SSQL = SSQL & "     left outer join ( " & vbNewLine
    SSQL = SSQL & "         select 고객코드, count(택번호) as ReturnCnt " & vbNewLine
    SSQL = SSQL & "           From TB_입출고 " & vbNewLine
    SSQL = SSQL & "        where 접수일자  between @SDate   AND  @EDate             " & vbNewLine
    SSQL = SSQL & "         AND 지사출고상태 = '2' " & vbNewLine
    SSQL = SSQL & "         AND 판매취소 <> 'Y' " & vbNewLine
    SSQL = SSQL & "         AND isnull(출고일자,'') = '' " & vbNewLine
    SSQL = SSQL & "        group by 고객코드) EE " & vbNewLine
    SSQL = SSQL & "     on AA.고객코드 = EE.고객코드 " & vbNewLine
    SSQL = SSQL & "  where BB.미출고고객코드 is null     " & vbNewLine
    SSQL = SSQL & " order by AA.고객코드  " & vbNewLine
    SSQL = SSQL & "   SET NOCOUNT OFF " & vbNewLine
    
    
    MyADOCon.Execute SSQL
    
    ' 업데이트가 완료
    Update_20160617_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20160617_01 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20160617_01 " & Err.description)

End Function

Private Function Update_20160927_01() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20160927_01 = False
    sUpdateKey = "20160927_01"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    SSQL = ""
    SSQL = SSQL & "CREATE TABLE [dbo].[TB_SMS080]("
    SSQL = SSQL & "    [휴대폰번호] [varchar](19) NOT NULL,"
    SSQL = SSQL & "        [발신전화번호] [varchar](19) NULL,"
    SSQL = SSQL & "        [등록시간] [varchar](14) NULL,"
    SSQL = SSQL & "        [적용여부] [varchar](1) NULL"
    SSQL = SSQL & "     CONSTRAINT [PK_TB_SMS080] PRIMARY KEY CLUSTERED"
    SSQL = SSQL & "    ("
    SSQL = SSQL & "        [휴대폰번호] Asc"
    SSQL = SSQL & "    )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
    SSQL = SSQL & "    ) ON [PRIMARY]"
        
    MyADOCon.Execute SSQL
    
    ' 업데이트가 완료
    Update_20160927_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20160927_01 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20160927_01 " & Err.description)

End Function


Private Function Update_20160927_02() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20160927_02 = False
    sUpdateKey = "20160927_02"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    SSQL = "ALTER TABLE TB_고객정보 ADD 문자080거부여부 char(1) ":             MyADOCon.Execute SSQL
    
    ' 업데이트가 완료
    Update_20160927_02 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20160927_02 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20160927_02 " & Err.description)

End Function


Private Function Update_20160928_01() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20160928_01 = False
    sUpdateKey = "20160928_01"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    SSQL = " "
    
    SSQL = SSQL & " IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[SP_TAG]') AND type in (N'P', N'PC'))" & vbNewLine
    SSQL = SSQL & " DROP PROCEDURE [dbo].[SP_TAG]" & vbNewLine
    MyADOCon.Execute SSQL


    SSQL = "  CREATE  PROCEDURE [dbo].[SP_TAG]" & vbNewLine
    SSQL = SSQL & "          @Cnt          INT                 -- 접수수량" & vbNewLine
    SSQL = SSQL & "        , @ReceiptDate  VARCHAR(10)         -- 접수일자" & vbNewLine
    SSQL = SSQL & "        , @ReceiptTime  VARCHAR(10)         -- 접수시간" & vbNewLine
    SSQL = SSQL & "        , @TagCode      VARCHAR(3)          -- 택코드" & vbNewLine
    SSQL = SSQL & "        , @BranchCode   VARCHAR(4)          -- 지사코드" & vbNewLine
    SSQL = SSQL & "        , @AgencyCode   VARCHAR(6)          -- 가맹점코드" & vbNewLine
    SSQL = SSQL & "        , @Rtn          VARCHAR(5000) OUTPUT" & vbNewLine
    SSQL = SSQL & " AS" & vbNewLine
    SSQL = SSQL & " BEGIN" & vbNewLine
    SSQL = SSQL & "    BEGIN TRANSACTION" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "    DECLARE @PadChar        CHAR(1)" & vbNewLine
    SSQL = SSQL & "    DECLARE @FinalLength    INTEGER" & vbNewLine
    SSQL = SSQL & "    DECLARE @TAG_List       VARCHAR(5000)" & vbNewLine
    SSQL = SSQL & "    DECLARE @Receipt_No     INT" & vbNewLine
    SSQL = SSQL & "    DECLARE @I              INT" & vbNewLine
    SSQL = SSQL & "    DECLARE @Tmp_TAG        VARCHAR(9)" & vbNewLine
    SSQL = SSQL & "    DECLARE @Tmp_No         INT" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "    SET @PadChar      = '0'" & vbNewLine
    SSQL = SSQL & "    SET @FinalLength  = 6" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "    -- 택번호, 접수번호" & vbNewLine
    SSQL = SSQL & "    SELECT @Tmp_TAG = ISNULL(택번호, @TagCode + '000001')" & vbNewLine
    SSQL = SSQL & "         , @Receipt_No = ISNULL(접수번호, 0) + 1" & vbNewLine
    SSQL = SSQL & "    FROM TB_기본정보" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "    SET @TAG_List = CONVERT(VARCHAR(10), @Receipt_No) + ':'" & vbNewLine
    SSQL = SSQL & "    SET @I = 1" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "    WHILE @I <= @Cnt" & vbNewLine
    SSQL = SSQL & "        BEGIN" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    
    SSQL = SSQL & "            -- 해당 택번호가 없을 경우만 처리한다." & vbNewLine
    SSQL = SSQL & "            IF  NOT EXISTS(SELECT 'TRUE' FROM TB_입출고 WHERE  가맹점코드   = @AgencyCode" & vbNewLine
    SSQL = SSQL & "                                                      AND  접수일자     = @ReceiptDate" & vbNewLine
    SSQL = SSQL & "                                                      AND  택번호       = @Tmp_TAG" & vbNewLine
    SSQL = SSQL & "                                                      AND  판매취소     <> 'Y')" & vbNewLine
    SSQL = SSQL & "            BEGIN " & vbNewLine
    
    SSQL = SSQL & "                 INSERT INTO TB_입출고(지사코드" & vbNewLine
    SSQL = SSQL & "                                    , 가맹점코드" & vbNewLine
    SSQL = SSQL & "                                    , 접수번호" & vbNewLine
    SSQL = SSQL & "                                    , 택번호" & vbNewLine
    SSQL = SSQL & "                                    , 접수일자" & vbNewLine
    SSQL = SSQL & "                                    , 접수시간" & vbNewLine
    SSQL = SSQL & "                                   ) VALUES (" & vbNewLine
    SSQL = SSQL & "                                      @BranchCode  -- 지사코드" & vbNewLine
    SSQL = SSQL & "                                    , @AgencyCode  -- 가맹점코드" & vbNewLine
    SSQL = SSQL & "                                    , @Receipt_No    -- 접수번호" & vbNewLine
    SSQL = SSQL & "                                    , @Tmp_TAG     -- 택번호" & vbNewLine
    SSQL = SSQL & "                                    , @ReceiptDate -- 접수일자" & vbNewLine
    SSQL = SSQL & "                                    , @ReceiptTime -- 접수시간" & vbNewLine
    SSQL = SSQL & "                                    )" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "                 SET @TAG_List = @TAG_List + @Tmp_TAG + ','" & vbNewLine
    SSQL = SSQL & "                 SET @I = @I + 1" & vbNewLine
    SSQL = SSQL & "            END" & vbNewLine
    
    SSQL = SSQL & "            IF SUBSTRING(@Tmp_TAG,4,6) = '999999'" & vbNewLine
    SSQL = SSQL & "                BEGIN" & vbNewLine
    SSQL = SSQL & "                    SET @Tmp_TAG = @TagCode + '000000'" & vbNewLine
    SSQL = SSQL & "                END" & vbNewLine
    SSQL = SSQL & "            Else" & vbNewLine
    SSQL = SSQL & "                BEGIN" & vbNewLine
    SSQL = SSQL & "                    SET @Tmp_No = CONVERT(INT, SUBSTRING(@Tmp_TAG,4,6)) + 1" & vbNewLine
    SSQL = SSQL & "                    SET @Tmp_TAG = @TagCode + RIGHT( REPLICATE(@PadChar, @FinalLength - LEN(@FinalLength)) + CONVERT(VARCHAR(6), @Tmp_No), 6)" & vbNewLine
    SSQL = SSQL & "                END" & vbNewLine
    SSQL = SSQL & "    END" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "    IF @@ERROR > 0" & vbNewLine
    SSQL = SSQL & "        BEGIN" & vbNewLine
    SSQL = SSQL & "            SET @Rtn = ''" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "            Rollback TRANSACTION" & vbNewLine
    SSQL = SSQL & "        END" & vbNewLine
    SSQL = SSQL & "    Else" & vbNewLine
    SSQL = SSQL & "        BEGIN" & vbNewLine
    SSQL = SSQL & "            -- 마지막택번호, 접수번호 저장" & vbNewLine
    SSQL = SSQL & "            UPDATE TB_기본정보 SET 택번호   = @Tmp_TAG" & vbNewLine
    SSQL = SSQL & "                                 , 접수번호 = @Receipt_No" & vbNewLine
    SSQL = SSQL & " " & vbNewLine
    SSQL = SSQL & "            SET @Rtn = @TAG_List" & vbNewLine
    SSQL = SSQL & "" & vbNewLine
    SSQL = SSQL & "            Commit TRANSACTION" & vbNewLine
    SSQL = SSQL & "        END" & vbNewLine
    SSQL = SSQL & "END"
    
    MyADOCon.Execute SSQL
    
    ' 업데이트가 완료
    Update_20160928_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20160928_01 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20160928_01 " & Err.description)

End Function



Private Function Update_20170201_01() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20170201_01 = False
    sUpdateKey = "20170201_01"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    SSQL = ""
    SSQL = SSQL & "IF  EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_TB_입출고_세탁마진]') AND type = 'D') " & vbNewLine
    SSQL = SSQL & "Begin " & vbNewLine
    SSQL = SSQL & "ALTER TABLE [dbo].[TB_입출고] DROP CONSTRAINT [DF_TB_입출고_세탁마진] " & vbNewLine
    SSQL = SSQL & "End " & vbNewLine
    MyADOCon.Execute SSQL
    
    SSQL = "ALTER TABLE TB_입출고 alter COLUMN 세탁마진 float ":            MyADOCon.Execute SSQL
    
    SSQL = ""
    SSQL = SSQL & "ALTER TABLE [dbo].[TB_입출고] ADD  CONSTRAINT [DF_TB_입출고_세탁마진]  DEFAULT ((0)) FOR [세탁마진]"
    MyADOCon.Execute SSQL
    
    
    SSQL = ""
    SSQL = SSQL & "IF  EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_입출고_외주운동화마진]') AND type = 'D')" & vbNewLine
    SSQL = SSQL & "Begin" & vbNewLine
    SSQL = SSQL & "ALTER TABLE [dbo].[TB_입출고] DROP CONSTRAINT [DF_입출고_외주운동화마진]" & vbNewLine
    SSQL = SSQL & "End" & vbNewLine
    MyADOCon.Execute SSQL
    
    SSQL = "ALTER TABLE TB_입출고 alter COLUMN 외주마진 float ":            MyADOCon.Execute SSQL
    
    SSQL = ""
    SSQL = SSQL & "ALTER TABLE [dbo].[TB_입출고] ADD  CONSTRAINT [DF_입출고_외주운동화마진]  DEFAULT ((0)) FOR [외주마진]"
    MyADOCon.Execute SSQL


    SSQL = ""
    SSQL = SSQL & "IF  EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_TB_입출고_수선마진]') AND type = 'D')" & vbNewLine
    SSQL = SSQL & "Begin" & vbNewLine
    SSQL = SSQL & "ALTER TABLE [dbo].[TB_입출고] DROP CONSTRAINT [DF_TB_입출고_수선마진]" & vbNewLine
    SSQL = SSQL & "End" & vbNewLine
    MyADOCon.Execute SSQL
    
    SSQL = "ALTER TABLE TB_입출고 alter COLUMN 수선마진 float ":            MyADOCon.Execute SSQL
    
    SSQL = ""
    SSQL = SSQL & "ALTER TABLE [dbo].[TB_입출고] ADD  CONSTRAINT [DF_TB_입출고_수선마진]  DEFAULT ((0)) FOR [수선마진]"
    MyADOCon.Execute SSQL
    
    ' 업데이트가 완료
    Update_20170201_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20170201_01 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20170201_01 " & Err.description)

End Function



Private Function Update_20170201_02() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20170201_02 = False
    sUpdateKey = "20170201_02"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    SSQL = ""
    SSQL = SSQL & "IF  EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_TB_의류분류_세탁마진]') AND type = 'D') " & vbNewLine
    SSQL = SSQL & "Begin " & vbNewLine
    SSQL = SSQL & "ALTER TABLE [dbo].[TB_의류분류] DROP CONSTRAINT [DF_TB_의류분류_세탁마진] " & vbNewLine
    SSQL = SSQL & "End " & vbNewLine
    MyADOCon.Execute SSQL
    
    SSQL = "ALTER TABLE TB_의류분류 alter COLUMN 세탁마진 float ":            MyADOCon.Execute SSQL
    
    SSQL = ""
    SSQL = SSQL & "ALTER TABLE [dbo].[TB_의류분류] ADD  CONSTRAINT [DF_TB_의류분류_세탁마진]  DEFAULT ((0)) FOR [세탁마진]"
    MyADOCon.Execute SSQL
    
    
    SSQL = ""
    SSQL = SSQL & "IF  EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_TB_의류분류_외주마진]') AND type = 'D')" & vbNewLine
    SSQL = SSQL & "Begin" & vbNewLine
    SSQL = SSQL & "ALTER TABLE [dbo].[TB_의류분류] DROP CONSTRAINT [DF_TB_의류분류_외주마진]" & vbNewLine
    SSQL = SSQL & "End" & vbNewLine
    MyADOCon.Execute SSQL
    
    SSQL = "ALTER TABLE TB_의류분류 alter COLUMN 외주마진 float ":            MyADOCon.Execute SSQL
    
    SSQL = ""
    SSQL = SSQL & "ALTER TABLE [dbo].[TB_의류분류] ADD  CONSTRAINT [DF_TB_의류분류_외주마진]  DEFAULT ((0)) FOR [외주마진]"
    MyADOCon.Execute SSQL


    SSQL = ""
    SSQL = SSQL & "IF  EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_TB_의류분류_수선마진]') AND type = 'D')" & vbNewLine
    SSQL = SSQL & "Begin" & vbNewLine
    SSQL = SSQL & "ALTER TABLE [dbo].[TB_의류분류] DROP CONSTRAINT [DF_TB_의류분류_수선마진]" & vbNewLine
    SSQL = SSQL & "End" & vbNewLine
    MyADOCon.Execute SSQL
    
    SSQL = "ALTER TABLE TB_의류분류 alter COLUMN 수선마진 float ":            MyADOCon.Execute SSQL
    
    SSQL = ""
    SSQL = SSQL & "ALTER TABLE [dbo].[TB_의류분류] ADD  CONSTRAINT [DF_TB_의류분류_수선마진]  DEFAULT ((0)) FOR [수선마진]"
    MyADOCon.Execute SSQL
    
    ' 업데이트가 완료
    Update_20170201_02 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20170201_02 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20170201_02 " & Err.description)

End Function



Private Function Update_20170310_01() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20170310_01 = False
    sUpdateKey = "20170310_01"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    

    SSQL = "ALTER TABLE TB_사고품내역 alter COLUMN 비고 text ":            MyADOCon.Execute SSQL
    
 
     
    ' 업데이트가 완료
    Update_20170310_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20170310_01 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20170310_01 " & Err.description)

End Function

Private Function Update_20190123_01() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String

    On Error GoTo Update_Error

    Update_20190123_01 = False
    sUpdateKey = "20190123_01"

    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    'If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    If GetColumnInfo("TB_일일마감", "전산사용료") = False Then Exit Function


    SSQL = "ALTER TABLE TB_일일마감 add 전산사용료 int NOT NULL CONSTRAINT DF_TB_일일마감_전산사용료 DEFAULT 0; ":            MyADOCon.Execute SSQL



    ' 업데이트가 완료
    Update_20190123_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20190123_01 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20190123_01 " & Err.description)

End Function

Private Function Update_20190130_01() As Boolean
    Dim Query        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20190130_01 = False
    sUpdateKey = "20190130_01"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    Query = Query & " CREATE TABLE [dbo].[tbl_SMS_SEND]("
    Query = Query & "     [접수번호] [int] NOT NULL,"
    Query = Query & "     [전송일자] [datetime] NOT NULL,"
    Query = Query & "  CONSTRAINT [PK_tbl_SMS_SEND] PRIMARY KEY CLUSTERED"
    Query = Query & " ("
    Query = Query & "     [접수번호] Asc"
    Query = Query & " )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
    Query = Query & " ) ON [PRIMARY]"
    MyADOCon.Execute Query
    
    
    ' 업데이트가 완료
    Update_20190130_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20190130_01 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20190123_01 " & Err.description)

End Function


Private Function Update_20190620_01() As Boolean
    Dim Query        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20190620_01 = False
    sUpdateKey = "20190620_01"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    
    
    Query = Query & " update tb_매출 SET 본사전송여부 = NULL WHERE 매출일자 > '2019-06-01'"
    MyADOCon.Execute Query
    Query = Query & " update tb_일일마감 SET 본사전송여부 = NULL WHERE 마감일자 > '2019-06-01'"
    MyADOCon.Execute Query
    Query = Query & " update tb_입출고 SET 본사전송여부 = NULL WHERE 접수일자 > '2019-06-01'"
    MyADOCon.Execute Query
    Query = Query & " update tb_신용카드승인 SET 본사전송여부 = NULL WHERE 승인일자 > '2019-06-01'"
    MyADOCon.Execute Query
    

    
    ' 업데이트가 완료
    Update_20190620_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20190620_01 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20190123_01 " & Err.description)

End Function


Private Function Update_20190624_01() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String

    On Error GoTo Update_Error

    Update_20190624_01 = False
    sUpdateKey = "20190624_01"

    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    'If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    If GetColumnInfo("TB_기본정보", "전산사용료") = False Then Exit Function


    SSQL = "ALTER TABLE TB_기본정보 add 전산사용료 int NOT NULL CONSTRAINT DF_TB_기본정보_전산사용료 DEFAULT 0; ":            MyADOCon.Execute SSQL



    ' 업데이트가 완료
    Update_20190624_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20190624_01 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20190624_01 " & Err.description)

End Function

Private Function Update_20210126_01() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20210126_01 = False
    sUpdateKey = "20210126_01"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
    
    SSQL = ""

    SSQL = SSQL & "CREATE TABLE [dbo].[TB_Internet_접수](" & vbNewLine
    SSQL = SSQL & " [지사코드] [varchar](4) NULL," & vbNewLine
    SSQL = SSQL & " [가맹점코드] [varchar](6) NOT NULL," & vbNewLine
    SSQL = SSQL & " [고객코드] [varchar](6) NOT NULL," & vbNewLine
    SSQL = SSQL & " [접수번호] [int] NOT NULL," & vbNewLine
    SSQL = SSQL & " [일련번호] [int] NOT NULL," & vbNewLine
    SSQL = SSQL & " [매출일자] [varchar](50) NOT NULL," & vbNewLine
    SSQL = SSQL & " [Internet_접수번호] [varchar](255) NOT NULL," & vbNewLine
    SSQL = SSQL & " CONSTRAINT [PK_TB_Internet_접수] PRIMARY KEY CLUSTERED " & vbNewLine
    SSQL = SSQL & "(" & vbNewLine
    SSQL = SSQL & " [가맹점코드] ASC," & vbNewLine
    SSQL = SSQL & " [고객코드] ASC," & vbNewLine
    SSQL = SSQL & " [접수번호] ASC" & vbNewLine
    SSQL = SSQL & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]" & vbNewLine
    SSQL = SSQL & ") ON [PRIMARY]" & vbNewLine

    MyADOCon.Execute SSQL
   
    ' 업데이트가 완료
    Update_20210126_01 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20210126_01 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20210126_01 " & Err.description)

End Function


Private Function Update_20210126_02() As Boolean
    Dim SSQL        As String
    Dim sUpdateKey  As String
    
    On Error GoTo Update_Error
    
    Update_20210126_02 = False
    sUpdateKey = "20210126_02"
    
    ' 이미 업데이트 되었을 경우 작업하지 않느다.
    If GetUpdateInfo(sUpdateKey) = False Then Exit Function
  
 
     
     
     
    SSQL = ""
    SSQL = SSQL & "CREATE TABLE [dbo].[TB_의류_인터넷]("
    SSQL = SSQL & " [의류코드] [varchar](4) NOT NULL,"
    SSQL = SSQL & " [의류명] [varchar](50) NOT NULL,"
    SSQL = SSQL & " [금액] [int] NOT NULL,"
    SSQL = SSQL & " [비고] [varchar](50) NULL,"
    SSQL = SSQL & " [순서] [varchar](4) NULL CONSTRAINT [DF_TB_의류_인터넷_순서]  DEFAULT ((0)),"
    SSQL = SSQL & " [적용일자] [varchar](10) NULL,"
    SSQL = SSQL & " [수신일자] [varchar](20) NULL,"
    SSQL = SSQL & " [본사전송여부] [varchar](1) NULL CONSTRAINT [DF_TB_의류_인터넷_본사전송여부]  DEFAULT (''),"
    SSQL = SSQL & " CONSTRAINT [PK_TB__인터넷] PRIMARY KEY CLUSTERED "
    SSQL = SSQL & "("
    SSQL = SSQL & " [의류코드] ASC"
    SSQL = SSQL & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
    SSQL = SSQL & ") ON [PRIMARY]"
    MyADOCon.Execute SSQL


    ' 업데이트가 완료
    Update_20210126_02 = SaveUpdateInfo(sUpdateKey)
    On Error GoTo 0
    Exit Function

Update_Error:
    Update_20210126_02 = False
    RaiseEvent Error(Err.Number, "CDB_Update Update_20210126_02 " & Err.description)

End Function

