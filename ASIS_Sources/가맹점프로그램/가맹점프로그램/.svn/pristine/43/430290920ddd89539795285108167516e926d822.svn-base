Attribute VB_Name = "baseDBCreate"
Option Explicit

'DBCC SHRINKFILE(CleanAID_log, 2)
'BACKUP LOG CleanAID WITH TRUNCATE_ONLY
'DBCC SHRINKFILE(CleanAID_log, 2)
'
'DBCC SHRINKDATABASE(CleanAID)

Private Function Set_CreateDatabase() As Boolean
    On Error GoTo ErrRtn
    
            Query = " USE MASTER" & vbNewLine
    Query = Query & " CREATE DATABASE CleanAID" & vbNewLine
    Query = Query & " ON" & vbNewLine
    Query = Query & " (NAME = CleanAID_dat" & vbNewLine
    Query = Query & "   FILENAME = 'C:\크린에이드\SQL_DATA\CleanAID.MDF'" & vbNewLine
    Query = Query & " )" & vbNewLine
    Query = Query & " LOG ON" & vbNewLine
    Query = Query & " (NAME = CleanAID_log" & vbNewLine
    Query = Query & "   FILENAME = 'C:\크린에이드\SQL_DATA\CleanAID_log.LDF'" & vbNewLine
    Query = Query & " )" & vbNewLine
    ADOCon.Execute Query
    
    Set_CreateDatabase = True
    Exit Function
    
ErrRtn:
    Set_CreateDatabase = False
End Function

Private Function Set_CreateTable()

End Function
