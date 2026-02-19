Attribute VB_Name = "basSmatro"
Option Explicit

Public Function GetCheckType(nIDX As Integer)
    If 0 = nIDX Then
        GetCheckType = "00"
    ElseIf 1 = nIDX Then
        GetCheckType = "01"
    ElseIf 2 = nIDX Then
        GetCheckType = "02"
    End If
End Function
            
Public Function GetCheckTypeCode(nIDX As Integer)
    If 0 = nIDX Then
        GetCheckTypeCode = "13"
    ElseIf 1 = nIDX Then
        GetCheckTypeCode = "14"
    ElseIf 2 = nIDX Then
        GetCheckTypeCode = "15"
    ElseIf 3 = nIDX Then
        GetCheckTypeCode = "16"
    End If
End Function

Public Function GetBonusType(nIDX As Integer)
    If 0 = nIDX Then
            GetBonusType = "1"
    ElseIf 1 = nIDX Then
            GetBonusType = "2"
    ElseIf 2 = nIDX Then
            GetBonusType = "3"
    ElseIf 3 = nIDX Then
            GetBonusType = "5"
    ElseIf 4 = nIDX Then
            GetBonusType = "6"
    ElseIf 5 = nIDX Then
            GetBonusType = "7"
    ElseIf 6 = nIDX Then
            GetBonusType = "8"
    ElseIf 7 = nIDX Then
            GetBonusType = "9"
    ElseIf 8 = nIDX Then
            GetBonusType = "A"
    ElseIf 9 = nIDX Then
            GetBonusType = "B"
    ElseIf 10 = nIDX Then
            GetBonusType = "G"
    ElseIf 11 = nIDX Then
            GetBonusType = "K"
    ElseIf 12 = nIDX Then
            GetBonusType = "L"
    ElseIf 13 = nIDX Then
            GetBonusType = "M"
    End If
End Function

Public Function GetUseType(nIDX As Integer)
    If 0 = nIDX Then
        GetUseType = "00"
    ElseIf 1 = nIDX Then
        GetUseType = "01"
    ElseIf 2 = nIDX Then
        GetUseType = "02"
    ElseIf 3 = nIDX Then
        GetUseType = "03"
    ElseIf 4 = nIDX Then
        GetUseType = "04"
    ElseIf 5 = nIDX Then
        GetUseType = "05"
    End If
End Function

    

