Attribute VB_Name = "basPictureSave"
'Option Explicit
'
'Dim strStream As ADODB.Stream
'
'Public Function LoadPictureFromDB(Rs As ADODB.Recordset)
'    On Error GoTo ErrRtn
'
'    'If Recordset is Empty, Then Exit
'    If Rs Is Nothing Then
'        GoTo ErrRtn
'    End If
'
'    Set strStream = New ADODB.Stream
'
'    strStream.Type = adTypeBinary
'    strStream.Open
'
'    strStream.Write Rs.Fields("오점이미지").Value
'
'    strStream.SaveToFile AppPath & "Temp.bmp", adSaveCreateOverWrite
'
'    Image1.Picture = LoadPicture(AppPath & "Temp.bmp")
'
'    Kill (AppPath & "Temp.bmp")
'
'    LoadPictureFromDB = True
'
'procExitFunction:
'
'    Exit Function
'
'ErrRtn:
'    LoadPictureFromDB = False
'
'    GoTo procExitFunction
'End Function
'
'Public Function SavePictureToDB(Rs As ADODB.Recordset, sFilename As String)
'
'    On Error GoTo ErrRtn
'
'    Dim oPict As StdPicture
'
'    Set oPict = LoadPicture(sFilename)
'
'    'Exit Function if this is NOT a picture file
'    If oPict Is Nothing Then
'        MsgBox "Invalid Picture File!", vbOKOnly, "Oops!"
'        SavePictureToDB = False
'        GoTo procExitSub
'    End If
'
'    Rs.AddNew
'
'    Set strStream = New ADODB.Stream
'    strStream.Type = adTypeBinary
'    strStream.Open
'    strStream.LoadFromFile sFilename
'    Rs.Fields("오점이미지").Value = strStream.Read
'
'    Image1.Picture = LoadPicture(sFilename)
'
'    SavePictureToDB = True
'
'procExitSub:
'    Exit Function
'ErrRtn:
'    SavePictureToDB = False
'    GoTo procExitSub
'End Function
'
'Private Function DownLoad(ByVal FileType As String, ByVal FileName As String) As Boolean
'    On Error GoTo Err:
'
'    Const ChunkSize = 32768 ' 32KByte
'
'    Dim FileExist
'    Dim FilePath As String
'    Dim FileNum, FileSize, FileOffSet As Long
'    Dim Cnt, Remain As Long
'    Dim GetByte_Buf() As Byte
'    Dim i As Long            ' 루프 카운트
'
'    Dim RS_DOWNLOAD As New ADODB.Recordset
'
'    Dim RemoveFilename As String
'
'        Select Case FileType
'            Case "01": FilePath = App.Path & "\"
'            Case "02": FilePath = App.Path & "\RPT"
'            Case "03": FilePath = App.Path & "\RSrc\Photo"
'        End Select
'
'        gsSQL2 = " SELECT * FROM FILEUPLOAD WHERE CODE='" & FileType & "' AND PGMID = '" & FileName & "'"
'        RS_DOWNLOAD.CursorLocation = adUseClient
'        RS_DOWNLOAD.LockType = adLockReadOnly
'        RS_DOWNLOAD.CursorType = adOpenKeyset
'        RS_DOWNLOAD.Open gsSQL2, Conn
'
'        If RS_DOWNLOAD.RecordCount > 0 Then
'            lblMsg = FileName & " DownLoad.... "
'            PB.Min = 0
'            PB.Value = 0
'
'            '=====================================================
'            '==== 파일의 존재를 구별하여 읽기전용속성을 풀어줌 ===
'            '=====================================================
'            FileExist = Dir(FilePath & "\" & FileName)
'
'            If FileExist = "" Then
'            Else
'                SetAttr FilePath & "\" & FileName, vbNormal
'            End If
'            '=====================================================
'
'            FileNum = FreeFile()
'
'            '읽기 전용속성파일인 경우 에러가 발생하므로 지워버리고 받음
'            'RemoveFilename = Dir(FilePath & "\" & FileName)
'            'If Len(RemoveFilename) > 0 Then Kill (FilePath & "\" & FileName)
'
'            Open FilePath & "\" & FileName For Binary Access Write As FileNum
'                FileSize = RS_DOWNLOAD.Fields("IMAGEINF").ActualSize      ' 파일 사이즈를 구한다.
'                Cnt = FileSize \ ChunkSize        ' 졍크사이즈 카운트를 구한다.
'
'                If Cnt > 0 Then PB.MAX = Cnt Else PB.MAX = 100
'
'                Remain = FileSize Mod ChunkSize   ' 졍크사이즈 미만의 나머지 바이트를 구한다.
'                ReDim GetByte_Buf(Remain)
'                GetByte_Buf() = RS_DOWNLOAD.Fields("IMAGEINF").GetChunk(Remain)
'                Put FileNum, , GetByte_Buf()
'                ReDim Byte_Buf(ChunkSize)
'                FileOffSet = Remain
'
'                For i = 1 To Cnt
'                    DoEvents
'                    GetByte_Buf() = RS_DOWNLOAD.Fields("IMAGEINF").GetChunk(ChunkSize)
'                    Put FileNum, , GetByte_Buf()
'                    FileOffSet = FileOffSet + ChunkSize
'                    PB.Value = i
'                Next i
'
'                If Cnt > 0 Then Else PB.Value = 100
'
'            Close FileNum
'            '===============================
'        End If
'        RS_DOWNLOAD.Close
'        Set RS_DOWNLOAD = Nothing
'        Exit Function
'Err:
'MsgBox Err.Number & " : " & Err.Description & vbNewLine & _
'       FileName & "을 다운로드중 에러가 발생했습니다."
'End Function
'
'
'Private Function FileUpLoad(FileType As String, FileName As String, FilePath As String) As Boolean
'
'On Error GoTo Err:
'
'Const ChunkSize = 32768 ' 32KByte
'
'Dim FileNum, FileSize As Double
'Dim Cnt, Remain As Double
'Dim GetByte_Buf() As Byte
'Dim i As Double            ' 루프 카운트
'Dim FileOffSet As Double
'Dim RS_UPLOAD As New ADODB.Recordset
'
'    lblMsg = FileName & " UPLOAD...."
'
'    'INSERT
'    gsSQL2 = " SELECT * FROM FILEUPLOAD WHERE CODE='" & FileType & "' AND PGMID = '" & FileName & "'"
'    RS_UPLOAD.CursorLocation = adUseClient
'    RS_UPLOAD.LockType = adLockOptimistic
'    RS_UPLOAD.CursorType = adOpenKeyset
'    RS_UPLOAD.Open gsSQL2, cDB.Conn
'
'    If RS_UPLOAD.RecordCount = 0 Then
'        RS_UPLOAD.AddNew
'        RS_UPLOAD.Fields("CODE") = FileType
'        RS_UPLOAD.Fields("PGMID") = FileName
'
'        '=============================
'        FileNum = FreeFile()
'        PB.Min = 0
'
'        Open FilePath & "\" & FileName For Binary Access Read As FileNum
'            FileSize = LOF(FileNum)           ' 파일 사이즈를 구한다.
'            Cnt = FileSize \ ChunkSize        ' 졍크사이즈 카운트를 구한다.
'            If Cnt = 0 Then
'                PB.MAX = 100
'            Else
'                PB.MAX = Cnt
'            End If
'            Remain = FileSize Mod ChunkSize   ' 졍크사이즈 미만의 나머지 바이트를 구한다.
'            RS_UPLOAD.Fields("IMAGEINF").AppendChunk ""
'            ReDim GetByte_Buf(Remain)
'            Get FileNum, , GetByte_Buf()
'            RS_UPLOAD.Fields("IMAGEINF").AppendChunk GetByte_Buf()
'            ReDim GetByte_Buf(ChunkSize)
'
'            For i = 1 To Cnt
'                DoEvents
'                Get FileNum, , GetByte_Buf()
'                RS_UPLOAD.Fields("IMAGEINF").AppendChunk GetByte_Buf()
'                PB.Value = i
'            Next i
'
'            If Cnt = 0 Then
'                PB.Value = 100
'            End If
'        Close FileNum
'
'        '===============================
'        RS_UPLOAD.Update
'    End If
'    cDB.Diss_Sel RS_UPLOAD
'    lblMsg = FileName & " UPLOAD....완료"
'    FileUpLoad = True
'
'    cDB.Diss_Sel RS1
'    Exit Function
'
'Err:
'    MsgBox Err.Number & " " & Err.Description
'    cDB.Diss_Sel RS1
'    lblMsg = FileName & " UPLOAD....에러발생"
'    FileUpLoad = False
'End Function
'
'
'
