' VBScript로 Get_Decrypt 테스트
' Windows에서 더블클릭으로 실행 가능

Option Explicit

Const sDefaultWHEEL1 = "ABCDEFGHIJKLMNOPQRSTVUWXYZ_1234567890qwertyuiopasd!@#$%^&*(),. ~`-=\?/'""fghjklzxcvbnm"
Const sDefaultWHEEL2 = "IWEHJKTLZVOPFG_1234567890qwerBNMQRYUASDXCfghjklzxc ~`-=\?/'""!@#$%^&*(),.vbnmtyuiopasd"

Function LeftShift(s)
    If Len(s) > 0 Then
        LeftShift = Mid(s, 2, Len(s) - 1) & Mid(s, 1, 1)
    Else
        LeftShift = s
    End If
End Function

Function RightShift(s)
    If Len(s) > 0 Then
        RightShift = Mid(s, Len(s), 1) & Mid(s, 1, Len(s) - 1)
    Else
        RightShift = s
    End If
End Function

Sub ScrambleWheels(ByRef sW1, ByRef sW2, sPASSWORD)
    Dim i, k
    For i = 1 To Len(sPASSWORD)
        For k = 1 To Asc(Mid(sPASSWORD, i, 1)) * i
            sW1 = LeftShift(sW1)
            sW2 = RightShift(sW2)
        Next
    Next
End Sub

Function Get_Decrypt(sInput, sPASSWORD)
    Dim sWHEEL1, sWHEEL2
    Dim k, c, i, sResult

    sWHEEL1 = sDefaultWHEEL1
    sWHEEL2 = sDefaultWHEEL2

    ' 공백 제거
    sInput = Replace(sInput, vbCrLf, "")
    sInput = Replace(sInput, vbCr, "")
    sInput = Replace(sInput, vbLf, "")
    sInput = Replace(sInput, vbTab, "")

    ' 비밀번호로 스크램블
    ScrambleWheels sWHEEL1, sWHEEL2, sPASSWORD

    sResult = ""

    For i = 1 To Len(sInput)
        c = Mid(sInput, i, 1)
        k = InStr(1, sWHEEL1, c, 0)  ' vbBinaryCompare = 0

        If k > 0 Then
            sResult = sResult & Mid(sWHEEL2, k, 1)
        Else
            sResult = sResult & c
        End If

        sWHEEL1 = LeftShift(sWHEEL1)
        sWHEEL2 = RightShift(sWHEEL2)
    Next

    Get_Decrypt = sResult
End Function

' ===== 테스트 실행 =====
Dim result
result = Get_Decrypt("10XX", "")

WScript.Echo "Get_Decrypt(""10XX"", """") = """ & result & """"
WScript.Echo ""
WScript.Echo "테스트 완료!"
