Attribute VB_Name = "Mod_Activate"
Option Explicit

Public Const SW_RESTORE = 9
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Sub ActivatePrevInstance(ByRef CurWindow As Form, ByVal sFindTitle As String)
Dim prevHndl As Long, strClass As String, strWindow As String
Dim lHwnd   As Long

    lHwnd = FindWindow(vbNullString, sFindTitle)
    
    '유효한 윈도 핸들일 경우에만 이전 실행 프로세스를 이전 크기로 복구 하기
    If lHwnd > 0 Then
        Call ShowWindow(lHwnd, SW_RESTORE)

        Call SetForegroundWindow(lHwnd)
    Else
        '이전 윈도 핸들 얻기에 실패한 경우 처리
    End If
    End
End Sub


