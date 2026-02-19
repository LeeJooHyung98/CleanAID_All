Attribute VB_Name = "Module1"
Option Explicit

Public MyCon As New ADODB.Connection                ' ActiveX Database Object 연결
Public MyMasterCon As New ADODB.Connection          ' ActiveX Database Object 연결

Public UserID As String * 6                 ' 사용자ID
Public USERNAME As String * 10              ' 사용자명

Public Const m_QN_PartnerID = "B01001"

Public sCopyIniFile As String   'Pda Copy 정보
Public DownPdaOutSide As String   '파일 저장전체 경로
Public DownPathName As String   '파일 저장전체 경로
Public DownFileName As String   '파일이름
Public DownState As String      '전송 상태

Public sIniFile As String

' From에 대한 초기화 값을 설정하여 준다.
Public P_01001_Flag As Boolean
Public P_01001_A_Flag As Boolean
Public P_01001_M_Flag As Boolean
Public P_01002_Flag As Boolean
Public P_01003_Flag As Boolean
Public P_01003_A_Flag As Boolean
Public P_01004_Flag As Boolean
Public P_01004_A_Flag As Boolean
Public P_01005_Flag As Boolean
Public P_01006_Flag As Boolean
Public P_01008_Flag As Boolean
Public P_01009_Flag As Boolean
Public P_01010_Flag As Boolean
Public P_01011_Flag As Boolean
Public P_01011_A_Flag As Boolean
Public P_01012_Flag As Boolean

Public P_02001_Flag As Boolean
Public P_02002_Flag As Boolean
Public P_02002_01_Flag As Boolean
Public P_02002_02_Flag As Boolean
Public P_02004_Flag As Boolean
Public P_02005_Flag As Boolean
Public P_02006_Flag As Boolean
Public P_02007_Flag As Boolean
Public P_02008_Flag As Boolean
Public P_02008_01_Flag As Boolean
Public P_02009_Flag As Boolean
Public P_02010_Flag As Boolean
Public P_02011_Flag As Boolean
Public P_02011_01_Flag As Boolean
Public P_02012_Flag As Boolean
Public P_02012_01_Flag As Boolean
Public P_02013_Flag As Boolean
Public P_02014_Flag As Boolean
Public P_02015_Flag As Boolean

Public P_03001_Flag As Boolean
Public P_03002_Flag As Boolean
Public P_03003_Flag As Boolean
Public P_03003_01_Flag As Boolean
Public P_03005_Flag As Boolean
Public P_03006_Flag As Boolean
Public P_03007_Flag As Boolean
Public P_03008_Flag As Boolean
Public P_03009_Flag As Boolean
Public P_03010_Flag As Boolean
Public P_03010_01_Flag As Boolean
Public P_03011_Flag As Boolean
Public P_03011_01_Flag As Boolean
Public P_03012_Flag As Boolean
Public P_03013_Flag As Boolean
Public P_03013_01_Flag As Boolean
Public P_03014_Flag As Boolean

Public P_04001_Flag As Boolean
Public P_04001_A_Flag As Boolean
Public P_04002_Flag As Boolean
Public P_04003_Flag As Boolean
Public P_04004_Flag As Boolean
Public P_04005_Flag As Boolean
Public P_04006_Flag As Boolean
Public P_04007_Flag As Boolean
Public P_04008_Flag As Boolean
Public P_04009_Flag As Boolean
Public P_04009_A_Flag As Boolean
Public P_04010_Flag As Boolean
Public P_04011_Flag As Boolean
Public P_04011_A_Flag As Boolean
Public P_04012_Flag As Boolean
Public P_04013_Flag As Boolean
Public P_04014_Flag As Boolean
Public P_04016_Flag As Boolean
Public P_04017_Flag As Boolean
Public P_04018_Flag As Boolean
Public P_04019_Flag As Boolean

Public P_05001_Flag As Boolean
Public P_05002_Flag As Boolean
Public P_05004_Flag As Boolean
Public P_05006_Flag As Boolean
Public P_05007_Flag As Boolean
Public P_05010_Flag As Boolean

Public P_06001_Flag As Boolean
Public P_06002_Flag As Boolean
Public P_06003_Flag As Boolean
Public P_06004_Flag As Boolean
Public P_06005_Flag As Boolean
Public P_06006_Flag As Boolean
Public P_06007_Flag As Boolean



Public P_07001_Flag As Boolean
Public P_07002_Flag As Boolean
Public P_07003_Flag As Boolean
Public P_07004_Flag As Boolean
Public P_07005_Flag As Boolean
Public P_07007_Flag As Boolean
Public P_07008_Flag As Boolean
Public P_07010_Flag As Boolean
Public P_07011_Flag As Boolean
Public P_07012_Flag As Boolean
Public P_07013_Flag As Boolean
Public P_07014_Flag As Boolean
Public P_07015_Flag As Boolean

Public P_08001_Flag As Boolean
Public P_08001_01_Flag As Boolean
Public P_08001_02_Flag As Boolean
Public P_08001_03_Flag As Boolean
Public P_08002_Flag As Boolean
Public P_08003_Flag As Boolean
Public P_08003_01_Flag As Boolean
Public P_08003_02_Flag As Boolean
Public P_08003_03_Flag As Boolean
Public P_08004_Flag As Boolean

Public P_09001_Flag As Boolean
Public P_09002_Flag As Boolean
Public P_09003_Flag As Boolean
Public P_09004_Flag As Boolean
Public P_09005_Flag As Boolean
Public P_09006_Flag As Boolean

Public P_10001_Flag As Boolean
Public P_10002_Flag As Boolean
Public P_10003_Flag As Boolean
Public P_10004_Flag As Boolean
Public P_10005_Flag As Boolean

Global Const glbYellow = &HC0FFFF
Global Const glbGreen = &HC0FFC0
Global Const glbGray = &HE0E0E0


' ini File Control에 관한 API라이브러리
Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : ExecPro
' 작  성  자  : IT21
' 작  성  일  : 2000.06.07
' 파 라 미 터 : ProcName - 프로시저명
'               sValue   - 프로시저 파라미터
'               Err_Num  - 에러번호
'               Err_Dec  - 에러명
' 리  턴  값  : Recordset
' 비      고  : Server에 있는 스토어드 프로시저를 실행하기 위한 함수
'-----------------------------------------------------------------------------------------------------------------------------------------
Function ExecProMaster(ByVal ProcName As String, ByRef sValue() As String, Err_Num As Long, Err_Dec As String) As ADODB.Recordset
    Dim i As Integer
    Dim MyCom As ADODB.Command
    
On Error GoTo ErrHandle

    Set ExecProMaster = New ADODB.Recordset
    Set MyCom = New ADODB.Command
    
    With MyCom
        .ActiveConnection = MyMasterCon
        .CommandTimeout = 0
        .CommandText = ProcName
        .CommandType = adCmdStoredProc
        
        For i = 1 To MyCom.Parameters.Count - 1
            If IsNull(sValue(i - 1)) Then
                MyCom.Parameters(i).Size = -1
            ElseIf sValue(i - 1) = "" Then
                MyCom.Parameters(i).Size = -1
            Else
                MyCom.Parameters(i).Size = LenH(sValue(i - 1))
            End If
            
            MyCom.Parameters(i) = sValue(i - 1)
        Next i
        
        Set ExecProMaster = .Execute
        
    End With
    
    Set MyCom = Nothing
    
    Err_Num = 0
    Err_Dec = ""
    
    Exit Function
    
ErrHandle:
    
    Err_Num = Err.Number
    Err_Dec = Err.Description
    
    Set MyCom = Nothing
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : ExecPro
' 작  성  자  : IT21
' 작  성  일  : 2000.06.07
' 파 라 미 터 : ProcName - 프로시저명
'               sValue   - 프로시저 파라미터
'               Err_Num  - 에러번호
'               Err_Dec  - 에러명
' 리  턴  값  : Recordset
' 비      고  : Server에 있는 스토어드 프로시저를 실행하기 위한 함수
'-----------------------------------------------------------------------------------------------------------------------------------------
Function ExecPro(ByVal ProcName As String, ByRef sValue() As String, Err_Num As Long, Err_Dec As String) As ADODB.Recordset
    Dim i As Integer
    Dim MyCom As ADODB.Command
    
    On Error GoTo ErrHandle

    Set ExecPro = New ADODB.Recordset
    Set MyCom = New ADODB.Command
    
    MyCom.ActiveConnection = MyCon
    MyCom.CommandTimeout = 0
    MyCom.CommandText = ProcName
    MyCom.CommandType = adCmdStoredProc
    
    For i = 1 To MyCom.Parameters.Count - 1
        If IsNull(sValue(i - 1)) Then
            MyCom.Parameters(i).Size = -1
        ElseIf sValue(i - 1) = "" Then
            MyCom.Parameters(i).Size = -1
        Else
            MyCom.Parameters(i).Size = LenH(sValue(i - 1))
        End If
        
        MyCom.Parameters(i) = sValue(i - 1)
    Next i
    
    Set ExecPro = MyCom.Execute
    Set MyCom = Nothing
    
    Err_Num = 0
    Err_Dec = ""
    
    Exit Function
    
ErrHandle:
    
    Err_Num = Err.Number
    Err_Dec = Err.Description
    
    Set MyCom = Nothing
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : WhatByte
' 작  성  자  : IT21
' 작  성  일  : 2000.06.07
' 파 라 미 터 : s        - 처리하고자 하는 문자열
'               chk_pos  -
' 리  턴  값  : Integer
' 비      고  : 한글코드를 Check하여서 그 값을 한글이면 '2'를 숫자,영문이면 '1'을 구한다.
'-----------------------------------------------------------------------------------------------------------------------------------------
Private Function WhatByte(ByVal s As String, ByVal chk_pos As Integer) As Integer
    Dim i As Integer

    '******************** 에러 처리 *********************
    If chk_pos > LenH(s) Then WhatByte = 0: Exit Function

    s = StrConv(s, 128)  '한글 코드 페이지

    For i = 1 To chk_pos
        If AscB(MidB(s, i, 1)) >= 128 Then
            WhatByte = 1: i = i + 1
        Else
            WhatByte = 0
        End If
    Next i

    If WhatByte = 1 And (i - 1) = chk_pos Then WhatByte = 2
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : LenH
' 작  성  자  : IT21
' 작  성  일  : 2000.06.07
' 파 라 미 터 : s        - 처리하고자 하는 문자열
' 리  턴  값  : Integer
' 비      고  : 문자열(s)의 길이를 구한다.
'               한글은 2바이트, 영문은 1바이트로 계산하여 전체 문자열의 길이를 구한다.
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Function LenH(ByVal s As String) As Integer
    LenH = LenB(StrConv(s, 128))
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : LeftH
' 작  성  자  : IT21
' 작  성  일  : 2000.06.07
' 파 라 미 터 : s        - 처리하고자 하는 문자열
' 리  턴  값  : string
' 비      고  : 문자열(s)의 왼쪽부터 n바이트 길이만큼 뽑아낸다.
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Function LeftH(ByVal s As String, ByVal n As Integer) As String
    Dim i, flag As Integer

    '***************** 에러 처리 *****************
    If s = "" Or n <= 0 Then Exit Function
    If n >= LenH(s) Then LeftH = s: Exit Function
    If WhatByte(s, n) = 1 Then n = n - 1: flag = 1

    s = StrConv(s, 128) '한글 코드 페이지.

    For i = 1 To n
        LeftH = LeftH & ChrB(AscB(MidB(s, i, 1)))
    Next i

    If flag Then LeftH = LeftH & ChrB(32)

    LeftH = StrConv(LeftH, 64) '유니 코드로 바꾼다.
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : MidH
' 작  성  자  : IT21
' 작  성  일  : 2000.06.07
' 파 라 미 터 : s        - 처리하고자 하는 문자열
' 리  턴  값  : string
' 비      고  : 문자열(s)의 start번째부터 n바이트 길이만큼 뽑아낸다.
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Function MidH(ByVal s As String, ByVal Start As Integer, ByVal n As Integer) As String
    Dim flag, fin, i As Integer

    '******************** 에러 처리 ********************
    If s = "" Or Start <= 0 Or n <= 0 Then Exit Function
    fin = Start + n - 1
    If fin >= LenH(s) Then fin = LenH(s)
    If WhatByte(s, Start) = 2 Then
        MidH = ChrB(32): Start = Start + 1
    End If
    If WhatByte(s, fin) = 1 Then fin = fin - 1: flag = 1

    s = StrConv(s, 128) '한글 코드 페이지.

    For i = Start To fin
        MidH = MidH & ChrB(AscB(MidB(s, i, 1)))
    Next i

    If flag Then MidH = MidH & ChrB(32)

    MidH = StrConv(MidH, 64) '유니 코드로 바꾼다.
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : RightH
' 작  성  자  : IT21
' 작  성  일  : 2000.06.07
' 파 라 미 터 : s        - 처리하고자 하는 문자열
' 리  턴  값  : string
' 비      고  : 문자열(s)의 오른쪽부터 n바이트 길이만큼 뽑아낸다.
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Function RightH(ByVal s As String, ByVal n As Integer)
    Dim Start, fin, i As Integer

    '***************** 에러 처리 *****************
    If s = "" Or n <= 0 Then Exit Function
    If n >= LenH(s) Then RightH = s: Exit Function

    fin = LenH(s)
    Start = fin - n + 1

    If WhatByte(s, Start) = 2 Then
        RightH = ChrB(32): Start = Start + 1
    End If

    s = StrConv(s, 128) '한글 코드 페이지.

    For i = Start To fin
        RightH = RightH & ChrB(AscB(MidB(s, i, 1)))
    Next i

    RightH = StrConv(RightH, 64) '유니 코드로 바꾼다.
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : GetColWidth
' 작  성  자  : IT21
' 작  성  일  : 2000.06.07
' 파 라 미 터 : New_App  -
'               New_Form -
'               New_SS   -
' 리  턴  값  : Boolean
' 비      고  : 스프레드의 Column의 길이를 불러온다.
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Function GetColWidth(ByVal New_App As String, ByVal New_Form As String, ByVal New_SS As Object) As Boolean
    Dim Col As Long

On Error GoTo GetColWidth_Err:
    
    GetColWidth = True
    
    For Col = 1 To New_SS.MaxCols
        New_SS.ColWidth(Col) = GetSetting(New_App, New_Form, New_SS.Name + "_ColWidth_" + CStr(Col), CStr(New_SS.ColWidth(Col)))
    Next Col
    
    Exit Function

GetColWidth_Err:
    GetColWidth = False
    
    Resume Next
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : SaveColWidth
' 작  성  자  : IT21
' 작  성  일  : 2000.06.07
' 파 라 미 터 : New_App  -
'               New_Form -
'               New_SS   -
' 리  턴  값  : Boolean
' 비      고  : 스프레드의 Column의 길이를 저장한다.
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Function SaveColWidth(ByVal New_App As String, ByVal New_Form As String, ByVal New_SS As Object) As Boolean
    Dim Col As Long

On Error GoTo SaveColWidth_Err:
    
    SaveColWidth = True
    For Col = 1 To New_SS.MaxCols
        SaveSetting New_App, New_Form, New_SS.Name + "_ColWidth_" + CStr(Col), CStr(New_SS.ColWidth(Col))
    Next Col

    Exit Function

SaveColWidth_Err:
    SaveColWidth = False
    Resume Next
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : Get_지사리스트
' 작  성  자  : pds2004
' 작  성  일  : 2007.05.04
' 파 라 미 터 : Control  - Combo Box Control Object
' 비      고  : Combo Box에 지사/유니트샆 내역을 Add한다.
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Sub Get_지사리스트(Control As Object, Optional Total As Boolean = True)
    Dim Rs As ADODB.Recordset
    Dim sValue() As String
    
    Dim Err_Num As Long
    Dim Err_Dec As String
    
    Control.Clear
    
    ReDim sValue(0)
    
    sValue(0) = "0"
    
    Set Rs = New ADODB.Recordset
    Set Rs = ExecPro("SP_00012", sValue(), Err_Num, Err_Dec)

    If Total = True Then
        Control.AddItem ""
    End If
    
    Do While Not Rs.EOF
        Control.AddItem "[" & Rs!지사코드 & "] " & Rs!지사명
        
        Rs.MoveNext
    Loop
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : Master_tblComboAdd
' 작  성  자  : pds2004
' 작  성  일  : 2007.05.04
' 파 라 미 터 : Control  - Combo Box Control Object
' 비      고  : Combo Box에 지사/유니트샆 내역을 Add한다.
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Sub Master_tblComboAdd(Control As Object)
    Dim Rs As ADODB.Recordset
    Dim sValue() As String
    
    Dim Err_Num As Long
    Dim Err_Dec As String
    
    ReDim sValue(0)
    
    sValue(0) = "0"
    
    Set Rs = New ADODB.Recordset
    Set Rs = ExecPro("SP_A_0001", sValue(), Err_Num, Err_Dec)

    Control.Clear
    Control.AddItem "[0000] 전체지사"

    Do Until Rs.EOF
        Control.AddItem "[" & Rs!지사코드 & "] " & Rs!지사명 & ""
        
        Rs.MoveNext
    Loop
    Rs.Close
    Set Rs = Nothing
End Sub



'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : AgencyComboAdd
' 작  성  자  : IT21
' 작  성  일  : 2000.06.08
' 파 라 미 터 : Control  - Combo Box Control Object
' 비      고  : Combo Box에 대리점내역을 Add한다.
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Sub MasterToAgencyComboAdd(Control As Object, sMasterCode As String)
    Dim Rs As ADODB.Recordset
    Dim sValue() As String
    
    Dim Err_Num As Long
    Dim Err_Dec As String
    
    Control.Clear
    
    ReDim sValue(0)
    
    sValue(0) = sMasterCode
    
    Set Rs = New ADODB.Recordset
    Set Rs = ExecPro("SP_01001_00", sValue(), Err_Num, Err_Dec) '이전 SP_00003_01

    Control.AddItem ""

    Do While Not Rs.EOF
        Control.AddItem "[" & Rs!가맹점코드 & "] " & Rs!가맹점명
        
        Rs.MoveNext
    Loop
End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : Get_가맹점리스트
' 작  성  자  : IT21
' 작  성  일  : 2000.06.08
' 파 라 미 터 : Control  - Combo Box Control Object
' 비      고  : Combo Box에 대리점내역을 Add한다.
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Sub Get_가맹점리스트(Control As Object, sMasterCode As String)
    Dim Rs As ADODB.Recordset
    Dim sValue() As String
    
    Dim Err_Num As Long
    Dim Err_Dec As String
    
    Control.Clear
    
    ReDim sValue(1)
    
    sValue(0) = "0"
    sValue(1) = sMasterCode
    
    Set Rs = New ADODB.Recordset
    Set Rs = ExecPro("SP_A_0004", sValue(), Err_Num, Err_Dec)

    'Control.AddItem ""

    Do Until Rs.EOF
        Control.AddItem "[" & Rs!가맹점코드 & "] " & Rs!가맹점명
        
        Rs.MoveNext
    Loop
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : AgencyComboAdd
' 작  성  자  : IT21
' 작  성  일  : 2000.06.08
' 파 라 미 터 : Control  - Combo Box Control Object
' 비      고  : Combo Box에 대리점내역을 Add한다.
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Sub AgencyComboAdd(Control As Object)
    Dim Rs As ADODB.Recordset
    Dim sValue() As String
    
    Dim Err_Num As Long
    Dim Err_Dec As String
    
    Control.Clear
    
    ReDim sValue(0)
    
    sValue(0) = "0"
    
    Set Rs = New ADODB.Recordset
    Set Rs = ExecPro("SP_00003", sValue(), Err_Num, Err_Dec)

    Control.AddItem ""

    Do While Not Rs.EOF
        Control.AddItem "[" & Rs!가맹점코드 & "] " & Rs!가맹점명
        
        Rs.MoveNext
    Loop
    
    Rs.Close
    Set Rs = Nothing
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : GoodsComboAdd
' 작  성  자  : IT21
' 작  성  일  : 2000.06.10
' 파 라 미 터 : Control  - Combo Box Control Object
' 비      고  : Combo Box에 품목내역을 Add한다.
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Sub GoodsComboAdd(Control As Object)
    Dim Rs As ADODB.Recordset
    Dim sValue() As String
    
    Dim Err_Num As Long
    Dim Err_Dec As String
    
    Control.Clear
    
    ReDim sValue(0)
    
    sValue(0) = "0"
    
    Set Rs = New ADODB.Recordset
    Set Rs = ExecPro("SP_00005", sValue(), Err_Num, Err_Dec)

    Control.AddItem ""
    
    Do While Not Rs.EOF
        Control.AddItem "[" & Rs!의류코드 & "] " & Rs!의류명
        
        Rs.MoveNext
    Loop
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : GoodsClassComboAdd
' 작  성  자  :
' 작  성  일  :
' 파 라 미 터 : Control  - Combo Box Control Object
' 비      고  : Combo Box에 품목내역을 Add한다.
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Sub GoodsClassComboAdd(Control As Object)
    Dim Rs As ADODB.Recordset
    Dim sValue() As String
    
    Dim Err_Num As Long
    Dim Err_Dec As String
    
    Control.Clear
    
    ReDim sValue(0)
    
    sValue(0) = "0"
    
    Set Rs = New ADODB.Recordset
    Set Rs = ExecPro("SP_00013", sValue(), Err_Num, Err_Dec)

    Control.AddItem ""
    
    Do Until Rs.EOF
        Control.AddItem "[" & Rs!의류분류코드 & "] " & Rs!의류분류명
        
        Rs.MoveNext
    Loop
    Rs.Close
    Set Rs = Nothing
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : PanelsMsg
' 작  성  자  : IT21
' 작  성  일  : 2000.06.16
' 파 라 미 터 : sMsg     - 메세지를 보여준다.
' 비      고  : 상태바의 메세지를 보여준다.
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Sub PanelsMsg(sMSG As String)
    P_00000.stbMsg.Panels(2).Text = sMSG
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : GetIniStr
' 작  성  자  : IT21
' 작  성  일  : 2000.06.21
' 파 라 미 터 : SectionName - Ini 파일의 Section부분 '[]'의 이름을 넣는다.
'               LineName    - Ini 파일의 LINE의 Head부분을 넣는다.
'               defValue    - 디폴트값
'               IniFileName - Ini 파일의 파일명.
' 비      고  : Ini파일에서 Data를 읽어온다.
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Function GetIniStr(SectionName As String, LineName As String, defValue As String, iniFileName As String) As String
    Dim retStr As String * 256
    Dim result As Integer
    
    result = GetPrivateProfileString(SectionName, LineName, defValue, retStr, Len(retStr), iniFileName)
    
    GetIniStr = LeftH(retStr, result)
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : DriverChk
' 작  성  자  : IT21
' 작  성  일  : 2000.06.22
' 비      고  : A:드라이브를 Check한다.
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Function DriverChk() As Boolean
    On Error GoTo DErr
    
    Dir "A:"
    DoEvents
    DriverChk = True
    
    Exit Function
    
DErr:
    DriverChk = False
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : GetAgencyName
' 작  성  자  : IT21
' 작  성  일  : 2000.06.22
' 파 라 미 터 : AgencyCode - 대리점코드
' 비      고  : 대리점명을 리턴한다.
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Function GetAgencyName(AgencyCode As String) As String
    Dim Rs As ADODB.Recordset
    Dim sValue() As String
    
    Dim Err_Num As Long
    Dim Err_Dec As String
    
    ReDim sValue(0)
    
    sValue(0) = AgencyCode
    
    Set Rs = New ADODB.Recordset
    Set Rs = ExecPro("SP_00007", sValue(), Err_Num, Err_Dec)

    If Rs.RecordCount = 0 Then
        GetAgencyName = ""
    Else
        GetAgencyName = Rs!대리점명
    End If
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : Sort_Select
' 작  성  자  : IT21
' 작  성  일  : 2000.06.26
' 파 라 미 터 : MySpread   - Spread Control
'               nSortOrder - Sort 방향
'               lRow       - 시작열
' 비      고  : 대리점명을 리턴한다.
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Sub Sort_Select(MySpread As Object, nSortOrder As Integer, lRow As Long)
    MySpread.Row = lRow
    MySpread.Col = 1
    MySpread.Row2 = MySpread.MaxRows
    MySpread.Col2 = MySpread.MaxCols
    MySpread.SortBy = 0
    MySpread.SortKey(1) = MySpread.ActiveCol
    
    If MySpread.ActiveCol = MySpread.MaxCols Then
        MySpread.SortKey(2) = 1
    Else
        MySpread.SortKey(2) = MySpread.ActiveCol + 1
    End If
    
    If MySpread.ActiveCol + 1 = MySpread.MaxCols Then
        MySpread.SortKey(3) = 1
    Else
        MySpread.SortKey(3) = MySpread.ActiveCol + 2
    End If
    
    MySpread.SortKeyOrder(1) = nSortOrder
    MySpread.SortKeyOrder(2) = nSortOrder
    MySpread.SortKeyOrder(3) = nSortOrder
    MySpread.Action = 25
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : ReportPrint
' 작  성  자  : IT21
' 작  성  일  : 2000.07.18
' 파 라 미 터 : sPrintFileName  - 리포트 파일이름
'               sOptional       - 출력방향
' 비      고  : 출력을 한다.
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Sub ReportPrint(sPrintFileName As String, sOptional As String)
'    P_00000.crPrint.ReportFileName = sPrintFileName
'    'P_00000.crPrint.Connect = gs_ReportConnect '리포트 접속
'
'    If sOptional = "1" Then
'
'
'        P_00000.crPrint.Destination = crptToPrinter
'
'        P_00000.crPrint.Action = 1
'
'    ElseIf sOptional = "2" Then
'        P_SCREEN.Show
'    End If
End Sub

Public Sub ColorComboAdd(Control As Object)
    Control.Clear

    Control.AddItem ""
    Control.AddItem "흰색"
    Control.AddItem "상아"
    Control.AddItem "회색"
    Control.AddItem "쥐색"
    Control.AddItem "밤색"
    Control.AddItem "검정"
    Control.AddItem "분홍"
    Control.AddItem "주황"
    Control.AddItem "빨강"
    Control.AddItem "노랑"
    Control.AddItem "베지"
    Control.AddItem "황토"
    Control.AddItem "연두"
    Control.AddItem "초록"
    Control.AddItem "카키"
    Control.AddItem "쑥색"
    Control.AddItem "하늘"
    Control.AddItem "파랑"
    Control.AddItem "곤색"
    Control.AddItem "보라"
    Control.AddItem "체크"
    Control.AddItem "자주"
    Control.AddItem "혼합"
End Sub

Public Function INIWrite(strSession As String, KeyValue As String, StrData As String, _
                        INIFile As String) As String
'====================================================================================================
' 작   성   자 : pds2004 박대선
' 작 성  일 자 : 2003.04.26
' 최종 수정 자 :
' 최종수정일자 :
' 사용 API함수 : WritepublicProfileString
'----------------------------------------------------------------------------------------------------
'   INI 값 기록
'====================================================================================================
    Dim lngRet As Long
    lngRet = WritePrivateProfileString(strSession, KeyValue, StrData, INIFile)

End Function

Public Function ExecWeekDay(MyDate As Date) As String
    Select Case Weekday(MyDate)
        Case 1: ExecWeekDay = "일"
        Case 2: ExecWeekDay = "월"
        Case 3: ExecWeekDay = "화"
        Case 4: ExecWeekDay = "수"
        Case 5: ExecWeekDay = "목"
        Case 6: ExecWeekDay = "금"
        Case 7: ExecWeekDay = "토"
    End Select
End Function

Public Sub g_GotFocusEvent(pc_RCtl As Control)
    pc_RCtl.SelStart = 0
    pc_RCtl.SelLength = Len(pc_RCtl.Text)
End Sub

'---------------------------------------------------------------------------
' 함수명 : Error_Msg
' 기  능 :
' 설  명 :
'---------------------------------------------------------------------------
Public Sub Error_Msg(strEvent As String, strSource As String, strNumber As String, strDescription As String)
    Dim Err_Msg As String
    
              Err_Msg = "발생위치 : " & strEvent & vbNewLine & vbNewLine
    Err_Msg = Err_Msg & "오류소스 : " & strSource & vbNewLine & vbNewLine
    Err_Msg = Err_Msg & "오류번호 : " & strNumber & vbNewLine & vbNewLine
    Err_Msg = Err_Msg & "오류내용 : " & strDescription
    
    MsgBox Err_Msg, vbCritical, "오류"
    Screen.MousePointer = 0
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : Get_지사리스트
' 작  성  자  : MemberGubunAdd
' 작  성  일  : 2007.05.04
' 파 라 미 터 : Control  - Combo Box Control Object
' 비      고  : Combo Box에 지사/유니트샆 내역을 Add한다.
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Sub MemberGubunAdd(Control As Object)
    Dim Rs As ADODB.Recordset
    Dim sValue() As String
    
    Dim Err_Num As Long
    Dim Err_Dec As String
    
    Control.Clear
    
    ReDim sValue(0)
    
    sValue(0) = "0"
    
    Set Rs = New ADODB.Recordset
    Set Rs = ExecPro("[SP_M_10000_00]", sValue(), Err_Num, Err_Dec)

    Control.AddItem ""

    Do While Not Rs.EOF
        Control.AddItem "[" & Rs!코드 & "] " & Rs!등급명
        
        Rs.MoveNext
    Loop
End Sub

'XML 변환...
Public Function Func_Replace(Str As String) As String
    Str = Replace(Str, "&", "&amp;")
    Str = Replace(Str, "<", "&lt;")
    Str = Replace(Str, ">", "&gt;")
    
    Func_Replace = Str
End Function
