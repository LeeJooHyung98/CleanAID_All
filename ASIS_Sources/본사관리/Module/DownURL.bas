Attribute VB_Name = "DownURL"
Option Explicit
Public DownUrl_FHandle              As Integer   ' 파일의 핸들

Const scUserAgent As String = "API-Guide test program"
Public Const scRegAppname As String = "CleanAid_Master"
Public Const scRegSection As String = "Maser_UpGrade"

Private Const INTERNET_OPEN_TYPE_PRECONFIG As Long = 0
Private Const INTERNET_FLAG_RELOAD As Long = &H80000000
Private Const INTERNET_FLAG_KEEP_CONNECTION As Long = &H400000
Private Const INTERNET_FLAG_NO_CACHE_WRITE As Long = &H4000000

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                        (ByVal hwnd As Long, ByVal lpOperation As String, _
                         ByVal lpFile As String, ByVal lpParameters As String, _
                         ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Integer
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long

Function OpenURL(sURL As String, Optional bufSize As Long = 1000) As String
       
    Dim hOpen As Long, hFile As Long, sBuffer As String, ret As Long
    
    sBuffer = Space$(bufSize)
    
    hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0): DoEvents
    hFile = InternetOpenUrl(hOpen, sURL, vbNullString, ByVal 0&, INTERNET_FLAG_RELOAD Or INTERNET_FLAG_KEEP_CONNECTION Or INTERNET_FLAG_NO_CACHE_WRITE, ByVal 0&): DoEvents
    
    InternetReadFile hFile, sBuffer, bufSize, ret: DoEvents
    
    InternetCloseHandle hFile
    InternetCloseHandle hOpen
    
    OpenURL = Left$(sBuffer, ret)
    
End Function


Public Function ProgramUpgrade() As Boolean
'+++++++++++++++++++++++++++++++++++++++++++++++++
' 작성일 : 2003/04/15
' 작성자 : pds2004 (박대선)
' 설  명 : 프로그램을 업그레이드 한다.
' 리턴값 : 0 정상실행
'          1 업그레이드 성공
'          -1 업그레이드 실패
'+++++++++++++++++++++++++++++++++++++++++++++++++
Dim PauseTime, Start, Finish, TotalTime
Dim SourceFile As String
Dim DestinationFile As String


SourceFile = App.Path & "\" & App.EXEName & ".exe" ' 판매재고관리UP.EXE
DestinationFile = App.Path & "\" & Mid(App.EXEName, 1, LenH(App.EXEName) - 2) & ".exe" '판매재고관리.EXE

If UCase(Right(App.EXEName, 2)) = "UP" Then
' 업그레이드 파일일경우]
    On Error Resume Next
    PauseTime = 10                  ' 기간을 지정합니다.
    Start = Timer                   ' 시작 시간을 지정합니다.
    FileCopy SourceFile, DestinationFile
        'ProgramUpgrade = True
    Do
        Finish = Timer              ' 종료 시간을 지정합니다.
        TotalTime = Finish - Start  ' 전체 시간을 계산합니다.
        If Dir(DestinationFile, vbDirectory) <> "" Then
            If Timer > Start + PauseTime Then
                MsgBox " 정상적으로 업그레이드 되지 않았습니다." & vbLf & "다시 시도 합니다.", vbCritical, "오류"
                On Error GoTo kill_end
                Kill DestinationFile
            End If
        Else
            Exit Do
        End If
        DoEvents                    ' 다른 프로시저로 넘깁니다.
        Kill DestinationFile        ' 원본 파일을 삭제한다.
    Loop
    
kill_end:
    If Dir(DestinationFile, vbDirectory) = "" Then
        FileCopy SourceFile, DestinationFile
        ProgramUpgrade = True
        Exit Function
    Else
        MsgBox " 정상적으로 업그레이드 되지 않았습니다.", vbCritical, "오류"
        ProgramUpgrade = False
        Exit Function
    End If
    
Else
    ProgramUpgrade = True
End If
    
End Function



Public Function ProgramUpgrade_20090911() As Boolean
'+++++++++++++++++++++++++++++++++++++++++++++++++
' 작성일 : 2003/04/15
' 작성자 : pds2004 (박대선)
' 설  명 : 프로그램을 업그레이드 한다.
' 리턴값 : 0 정상실행
'          1 업그레이드 성공
'          -1 업그레이드 실패
'+++++++++++++++++++++++++++++++++++++++++++++++++
Dim PauseTime, Start, Finish, TotalTime
Dim SourceFile As String
Dim DestinationFile As String
Dim sAppName    As String

sAppName = App.EXEName
SourceFile = App.Path & "\" & sAppName & ".exe" ' 판매재고관리UP.EXE
DestinationFile = App.Path & "\" & Mid(sAppName, 1, Len(sAppName) - 2) & ".exe" '판매재고관리.EXE

If UCase(Right(sAppName, 2)) = "UP" Then
' 업그레이드 파일일경우]
    On Error Resume Next
    PauseTime = 10                  ' 기간을 지정합니다.
    Start = Timer                   ' 시작 시간을 지정합니다.
    FileCopy SourceFile, DestinationFile
        'ProgramUpgrade = True
    Do
        Finish = Timer              ' 종료 시간을 지정합니다.
        TotalTime = Finish - Start  ' 전체 시간을 계산합니다.
        If Dir(DestinationFile, vbDirectory) <> "" Then
            If Timer > Start + PauseTime Then
                MsgBox " 정상적으로 업그레이드 되지 않았습니다." & vbLf & "다시 시도 합니다.     ", vbCritical, "오류"
                On Error GoTo kill_end
                Kill DestinationFile
            End If
        Else
            Exit Do
        End If
        DoEvents                    ' 다른 프로시저로 넘깁니다.
        Kill DestinationFile        ' 원본 파일을 삭제한다.
    Loop
    
kill_end:
    If Dir(DestinationFile, vbDirectory) = "" Then
        FileCopy SourceFile, DestinationFile
        MsgBox " 업그레이드 완료. " & vbLf & vbLf & " 프로그램을 다시 시작 하여 주십시요.     ", vbInformation, "확인"
        End
        
        Exit Function
    Else
        MsgBox " 정상적으로 업그레이드 되지 않았습니다.", vbCritical, "오류"
        ProgramUpgrade_20090911 = False
        Exit Function
    End If
    
Else
    ProgramUpgrade_20090911 = True
End If
    
End Function




