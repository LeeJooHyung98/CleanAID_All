Attribute VB_Name = "basHAN_LIB"

Function Line_Format(Using_Form As String, Print_Form As Variant, Print_Counter As Integer) As String
    Dim ii_loop        As Long
    Dim Old_Using_Form As String
    Dim temp           As String
    
    Old_Using_Form = Using_Form
    
    For ii_loop = 1 To Print_Counter
        For jj_loop = 1 To Len(Using_Form)
            Select Case Mid(Using_Form, jj_loop, 1)
                   Case "#", "@", "!"
                        kk_cnt = 0
                        
                        For kk_loop = jj_loop To Len(Using_Form)
                            If Mid(Using_Form, kk_loop, 1) = " " Then
                               Exit For
                            End If
                            
                            kk_cnt = kk_cnt + 1
                        Next kk_loop
                        
                        temp = Mid(Using_Form, jj_loop, kk_cnt)
                        
                        If Mid(temp, 1, 1) = "#" Then
                           Mid(Using_Form, jj_loop, kk_cnt) = Format_N(temp, Print_Form(ii_loop))
                        ElseIf Mid(temp, 1, 1) = "!" Then
                           Mid(Using_Form, jj_loop, kk_cnt) = Format_D(temp, Print_Form(ii_loop))
                        ElseIf Mid(temp, 1, 1) = "^" Then
                           Mid(Using_Form, jj_loop, kk_cnt) = Format_X(temp, Print_Form(ii_loop))
                        ElseIf Mid(temp, 1, 1) = "@" Then
                           Using_Form = Left(Using_Form, jj_loop - 1) & Format_H(temp, Print_Form(ii_loop)) & Mid(Using_Form, kk_cnt + jj_loop)
                        End If
                        
                        Exit For
                   Case Else
            End Select
        Next jj_loop
    Next ii_loop
    
    Line_Format = Using_Form
    Using_Form = Old_Using_Form
End Function

Function Format_N(Using_Form As String, Print_Form As Variant) As String
    Dim temp As String
    
    Using_Pointer1 = InStr(1, Using_Form, ".") - 1
    Using_Pointer2 = InStr(1, Using_Form, ".")
    
    If Using_Pointer1 = -1 Then
       Using_Pointer1 = Len(Using_Form)
       Using_Pointer2 = 0
    End If
    
    Using_Pointer3 = Using_Pointer1
    
    temp = LTrim(Val(Print_Form))
    Print_Pointer1 = InStr(1, temp, ".") - 1
    Print_Pointer2 = InStr(1, temp, ".")
    
    If Print_Pointer1 = -1 Then
       Print_Pointer1 = Len(temp)
       Print_Pointer2 = 0
    End If
    
    Temp_Pointer1 = 0
    
    For ii_loop = Print_Pointer1 To 1 Step -1
        If Using_Pointer1 < 1 Then
           Format_N = String(Len(Using_Form), "*")
           Exit Function
        End If
        
        Temp_Pointer1 = Temp_Pointer1 + 1
        
        If Mid(Using_Form, Using_Pointer1, 1) = "," Then
           If Mid(temp, ii_loop, 1) <> "-" Then
              Using_Pointer1 = Using_Pointer1 - 1
              Temp_Pointer1 = Temp_Pointer1 + 1
           End If
        End If
        
        Mid(Using_Form, Using_Pointer1, 1) = Mid(temp, ii_loop, 1)
        Using_Pointer1 = Using_Pointer1 - 1
    Next ii_loop
    
    Mid(Using_Form, 1, (Using_Pointer3 - Temp_Pointer1)) = Space(Using_Pointer3 - Temp_Pointer1)
    
    If Using_Pointer2 = 0 Then
       Format_N = IIf(Val(Trim(Using_Form)) = 0, Space(Using_Pointer3 - Temp_Pointer1) & " ", Using_Form)
       
       Exit Function
    End If
    
    For ii_loop = 1 To Len(Mid(Using_Form, Using_Pointer2)) - 1
        If Print_Pointer2 = 0 Or (Print_Pointer2 + ii_loop) > Len(temp) Then
           If Mid(Using_Form, Using_Pointer2 + ii_loop, 1) = "#" Then
              Mid(Using_Form, Using_Pointer2 + ii_loop, 1) = "0"
           Else
              Mid(Using_Form, Using_Pointer2 + ii_loop, 1) = " "
           End If
        Else
           Mid(Using_Form, Using_Pointer2 + ii_loop, 1) = Mid(temp, Print_Pointer2 + ii_loop, 1)
        End If
    Next ii_loop
    
    Format_N = IIf(Val(Trim(Using_Form)) = 0, Space(Using_Pointer3 - Temp_Pointer1) & " ", Using_Form)
End Function

Function Format_X(Using_Form As String, Print_Form As Variant) As String
    For ii_loop = 1 To Len(Using_Form)
        If ii_loop > Len(Using_Form) Then
           Mid(Using_Form, ii_loop, 1) = " "
        Else
           Mid(Using_Form, ii_loop, 1) = Mid(Print_Form, ii_loop, 1)
        End If
    Next ii_loop
    
    Format_X = Using_Form
End Function

Function Format_D(Using_Form As String, Print_Form As Variant) As String
    Select Case Len(Using_Form)
           Case 2, 3
                Mid(Using_Form, 1, 2) = Mid(Print_Form, 1, 2)
           Case 4
                Mid(Using_Form, 1, 4) = Mid(Print_Form, 1, 4)
           Case 5
                If InStr(1, Print_Form, ":") > 0 Or InStr(1, Print_Form, "/") > 0 Or InStr(1, Print_Form, "-") > 0 Or InStr(1, Print_Form, ".") > 0 Then
                   Mid(Using_Form, 1, 2) = Mid(Print_Form, 1, 2)
                   Mid(Using_Form, 4, 2) = Mid(Print_Form, 4, 2)
                Else
                   Mid(Using_Form, 1, 2) = Mid(Print_Form, 1, 2)
                   Mid(Using_Form, 4, 2) = Mid(Print_Form, 3, 2)
                End If
           Case 8
                If InStr(1, Print_Form, ":") > 0 Or InStr(1, Print_Form, "/") > 0 Or InStr(1, Print_Form, "-") > 0 Or InStr(1, Print_Form, ".") > 0 Then
                   Mid(Using_Form, 1, 2) = Mid(Print_Form, 1, 2)
                   Mid(Using_Form, 4, 2) = Mid(Print_Form, 4, 2)
                   Mid(Using_Form, 7, 2) = Mid(Print_Form, 7, 2)
                Else
                   Mid(Using_Form, 1, 2) = Mid(Print_Form, 1, 2)
                   Mid(Using_Form, 4, 2) = Mid(Print_Form, 3, 2)
                   Mid(Using_Form, 7, 2) = Mid(Print_Form, 5, 2)
                End If
           Case 10
                If InStr(1, Print_Form, ":") > 0 Or InStr(1, Print_Form, "/") > 0 Or InStr(1, Print_Form, "-") > 0 Or InStr(1, Print_Form, ".") > 0 Then
                   Mid(Using_Form, 1, 4) = Mid(Print_Form, 1, 4)
                   Mid(Using_Form, 6, 2) = Mid(Print_Form, 6, 2)
                   Mid(Using_Form, 9, 2) = Mid(Print_Form, 9, 2)
                Else
                   Mid(Using_Form, 1, 4) = Mid(Print_Form, 1, 4)
                   Mid(Using_Form, 6, 2) = Mid(Print_Form, 5, 2)
                   Mid(Using_Form, 9, 2) = Mid(Print_Form, 7, 2)
                End If
           Case Else
                Using_Form = String(Len(Using_Form), "*")
    End Select
    Format_D = Using_Form
End Function

Function Format_H(Using_Form As String, Print_Form As Variant) As String
    Format_H = Hangul_Convert(Print_Form, Len(Using_Form))
End Function

Function Hangul_Convert(Convert As Variant, Max_length)
    Dim Pointer_Value        As Long
    Dim Character_Length     As Integer
    Dim Loop_Index           As Integer
    
    Character_Length = 0
    Convert = Convert + Space(Max_length)
    
    For Loop_Index = 1 To Max_length
        Pointer_Value = Asc(Mid(Convert, Loop_Index, 1))
        
        If Pointer_Value < 0 Then
           Character_Length = Character_Length + 1
        End If
    Next Loop_Index
    
    Hangul_Convert = Left(Convert + Space(Max_length), Max_length - Character_Length)
End Function

Function Hangul_Cut(Cutting As String, Length) As String
    Hangul_Cut = ""
    Character_Length = 0
    
    For Loop_Index = 1 To Len(Cutting)
        Pointer_Value = Asc(Mid(Cutting, Loop_Index, 1))
        
        If Pointer_Value < 0 Then
           Character_Length = Character_Length + 2
        Else
           Character_Length = Character_Length + 1
        End If
        
        Hangul_Cut = Hangul_Cut & Mid(Cutting, Loop_Index, 1)
        
        If Character_Length >= Length Then
           Cutting = Mid(Cutting, Loop_Index + 1)
           Exit For
        End If
    Next Loop_Index
End Function

Function Hangul_Length(Length As String) As Integer
    Hangul_Length = 0
    
    For Loop_Index = 1 To Len(Length)
        Pointer_Value = Asc(Mid(Length, Loop_Index, 1))
        
        If Pointer_Value < 0 Then
           Hangul_LengthCharacter_Length = Character_Length + 2
        Else
           Character_Length = Character_Length + 1
        End If
    Next Loop_Index
End Function

Function Hangul_Mid(Cutting As String, Length1, length2) As String
    Hangul_Mid = ""
    Character_Length = 0
    Character_Length1 = 0
    
    For Loop_Index = 1 To Len(Cutting)
        Pointer_Value = Asc(Mid(Cutting, Loop_Index, 1))

        If Pointer_Value < 0 Then
           Character_Length = Character_Length + 2
        Else
           Character_Length = Character_Length + 1
        End If
        
        Hangul_test = Hangul_test & Mid(Cutting, Loop_Index, 1)
        
        If Character_Length >= Length1 Then
            Hangul_Mid = Hangul_Mid & Mid(Cutting, Loop_Index, 1)
            
            If Pointer_Value < 0 Then
                Character_Length1 = Character_Length1 + 2
            Else
                Character_Length1 = Character_Length1 + 1
            End If
        End If
        
        If Character_Length1 >= length2 Then
           Exit For
        End If
    Next Loop_Index

End Function

