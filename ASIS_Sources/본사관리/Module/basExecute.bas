Attribute VB_Name = "basExecute"
Option Explicit

' 연결정보가 있는 파일을 실행하는 API함수
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

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
        Excute_Program = OpenShell(FormObj.hwnd, "open", strFilename, SW_SHOWDEFAULT)
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
    End Select
End Function

'Spread의 내용을 엑셀파일로 Export 시키기...
Public Sub Export_Excel(cdgExcel As CommonDialog, Spread As fpSpread)
    Dim i        As Integer
    Dim j        As Integer
    Dim x        As Boolean   'Spread Excel File Save...
    Dim Header() As String

    With cdgExcel
        .CancelError = False
        .InitDir = App.Path
        .Filter = "Excel (*.xls)|*.xls"
        .ShowSave

        If .FileName <> "" Then
            Spread.Redraw = False
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

            x = Spread.ExportToExcel(.FileName, "Sheet1", "")
'            x = Spread.ExportExcelBook(.FileName, .FileName & ".log")
'            x = Spread.ExportExcelBookEx(.FileName, .FileName & ".log", ExcelSaveFlagNoFormulas)

            For i = 1 To Spread.ColHeaderRows
                Spread.Row = 1
                Spread.Action = ActionDeleteRow
                Spread.MaxRows = Spread.MaxRows - 1
            Next i

            Spread.Protect = True
            Spread.Redraw = True

            If x = True Then
                'MsgBox .FileName & vbNewLine & "엑셀파일로 저장되었습니다.", vbInformation, "확인"
            Else
                MsgBox "엑셀파일로 저장하지 못하였습니다.", vbCritical, "확인"
            End If
            
            '-------------------------------------------------------------------------------------
            ' 연결 프로그램
            '-------------------------------------------------------------------------------------
            Dim strreturn As String
        
            strreturn = Excute_Program(P_00000, .FileName)
        
            If strreturn = "" Then
                '성공
            Else
                MsgBox strreturn & Space(4), vbCritical '실패
            End If
            
        End If
    End With
End Sub

