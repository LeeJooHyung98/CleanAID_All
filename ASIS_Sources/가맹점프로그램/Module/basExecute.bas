Attribute VB_Name = "basExecute"
Option Explicit

' 연결정보가 있는 파일을 실행하는 API함수
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' ShellExecute실행시 창모양
Public Const SW_SHOWDEFAULT = 10     ' 기본값
Public Const SW_SHOWMAXIMIZED = 3    ' 최대화면
Public Const SW_SHOWMINIMIZED = 2    ' 아이콘
Public Const SW_SHOWNORMAL = 1       ' 보통창

' ShellExecute 실행후 리턴되는 값중에서 에러상수
Public Const ERROR_FILE_NOT_FOUND = 2&  ' 화일이 없음
Public Const ERROR_PATH_NOT_FOUND = 3&  ' 경로가 없음
Public Const ERROR_BAD_FORMAT = 11&     ' 형식이 맞지않음
Public Const ERROR_GEN_FAILURE = 31&    ' 일반적인 에러

Public Function Excute_Program(ByVal FormObj As Form, ByVal strFilename As String) As String
    '연결된 파일을 실행합니다

    Excute_Program = "Error"

    If Dir(strFilename) = "" Then
        Excute_Program = "Error File Not Found"
    Else
        Excute_Program = OpenShell(FormObj.hWnd, "open", strFilename, SW_SHOWDEFAULT)
    End If
End Function

Public Function OpenShell(lHwnd As Long, sOperation As String, sFile As String, iShowCmd As Integer) As String
    Dim lRet As Long

    lRet = ShellExecute(lHwnd, sOperation, sFile, vbNullString, vbNullString, iShowCmd)

    '에러를 처리할 부분
    Select Case lRet
        Case ERROR_FILE_NOT_FOUND: OpenShell = "Error File Not Found"
        Case ERROR_PATH_NOT_FOUND: OpenShell = "Error Path Not Found"
        Case ERROR_BAD_FORMAT:     OpenShell = "Error File Bad Format"
        Case ERROR_GEN_FAILURE:    OpenShell = "Error General Failure"
        Case Else
            OpenShell = CStr(lRet)
        
    End Select
End Function

'Spread의 내용을 엑셀파일로 Export 시키기...
Public Sub Export_Excel(cdgExcel As XtremeSuiteControls.CommonDialog, Spread As fpSpread)
    Dim j        As Long
    Dim x        As Boolean   'Spread Excel File Save...
    Dim Header() As String

    With cdgExcel
        .CancelError = False
        .InitDir = App.Path
        .Filter = "Excel (*.xls)|*.xls"
        .ShowSave

        If .FileName <> "" Then
            Spread.ReDraw = False
            Spread.Protect = False
            
            For i = 1 To Spread.ColHeaderRows
                '헤더를 배열에 넣기
                ReDim Header(Spread.MaxCols) As String

                Spread.Row = SpreadHeader + (i - 1)

                For j = 1 To Spread.MaxCols
                    Spread.Col = j
                    Header(j) = Spread.Text & ""
                Next j

                '배열값을 Spread에 넣기
                Spread.MaxRows = Spread.MaxRows + 1
                Spread.Row = i
                Spread.Action = ActionInsertRow

                For j = 1 To Spread.MaxCols
                    Spread.Col = j
                    Spread.CellType = CellTypeEdit
                    Spread.TypeHAlign = TypeHAlignCenter
                    Spread.TypeVAlign = TypeVAlignCenter
                    Spread.Text = Header(j) & ""
                Next j
            Next i

            'x = Spread.ExportToExcel(.FileName, "Sheet1", "")
            x = Spread.ExportToTextFile(Replace(.FileName, "xls", "csv"), """", ",", vbCrLf, ExportToTextFileCreateNewFile, "")

            For i = 1 To Spread.ColHeaderRows
                Spread.Row = 1
                Spread.Action = ActionDeleteRow
                Spread.MaxRows = Spread.MaxRows - 1
            Next i

            Spread.Protect = True
            Spread.ReDraw = True

            If x = True Then
                'MsgBox .FileName & vbNewLine & "엑셀파일로 저장되었습니다.", vbInformation, "확인"
            Else
                MsgBox "엑셀파일로 저장하지 못하였습니다.", vbCritical, "확인"
            End If
            
            '-------------------------------------------------------------------------------------
            ' 연결 프로그램
            '-------------------------------------------------------------------------------------
            Dim strReturn As String
        
            strReturn = Excute_Program(frmMain, .FileName)
        
            If strReturn = "" Then
                '성공
            Else
                MsgBox strReturn & Space(4), vbCritical '실패
            End If
            
        End If
    End With
End Sub

'XML 변환...
Public Function Func_Replace(Str As String) As String
    Str = Replace(Str, "&", "&amp;")
    Str = Replace(Str, "<", "&lt;")
    Str = Replace(Str, ">", "&gt;")
    
    Func_Replace = Str
End Function

Public Function Fun_Week(strDay As String) As String
    On Error GoTo ErrRtn
    
    Select Case Weekday(strDay)
        Case 1: Fun_Week = "일요일"
        Case 2: Fun_Week = "월요일"
        Case 3: Fun_Week = "화요일"
        Case 4: Fun_Week = "수요일"
        Case 5: Fun_Week = "목요일"
        Case 6: Fun_Week = "금요일"
        Case Else: Fun_Week = "토요일"
    End Select
    
    Exit Function
    
ErrRtn:
    Fun_Week = ""
End Function
