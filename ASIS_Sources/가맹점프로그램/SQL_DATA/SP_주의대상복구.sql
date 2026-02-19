    ----------------------------------------------------------------------
    -- 반드시 c:\CleanAid\Sql_Data 폴더의 내용을 백업후 실행 할것
    ----------------------------------------------------------------------
    EXEC sp_resetstatus 'CleanAid';
    ALTER DATABASE CleanAid SET EMERGENCY
    DBCC checkdb('CleanAid')
    ALTER DATABASE CleanAid SET SINGLE_USER WITH ROLLBACK IMMEDIATE
    DBCC CheckDB ('CleanAid', REPAIR_ALLOW_DATA_LOSS)
    ALTER DATABASE CleanAid SET MULTI_USER
