Attribute VB_Name = "ProgramVersion"
Option Explicit

Public strProgram_Version   As String
Public strProgram_LastEdit  As String


Public Function SetProgramVersion()
    Dim MyVersion As String
    
    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision
    strProgram_LastEdit = "2015.10.28"
    ' 입고및 출고 현황에서 지사별 인쇄 기능 추가
    
'    strProgram_Version = App.Major & "." & App.Minor & "." & App.Revision
'    strProgram_LastEdit = "2011.05.06"
    ' 자료 적용시 이중 동작 되지 않도록 수정
    
    
    

End Function

