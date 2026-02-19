Attribute VB_Name = "pds_Module"
Option Explicit

Public Const MASTER_CODE As String = "1000"

Public Sub SubBottonEnable(cbo As Object, sBtnIndex As String)
    Dim i As Integer
    
    For i = 0 To Len(sBtnIndex) - 1
        cbo(i).Enabled = Val(Mid(sBtnIndex, i + 1, 1))
    Next i
    
End Sub



'-----------------------------------------------------------------------------------------------------------------------------------------
' Function명  : OrderComboAdd
' 작  성  자  : pds2004
' 작  성  일  : 2010.10.20
' 파 라 미 터 : Control  - Combo Box Control Object
' 비      고  : Combo Box에 외주 업체 내역을 Add한다.
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Sub OrderComboAdd(Control As Object)
    Dim Rs As ADODB.Recordset
    Dim sValue() As String
    
    Dim Err_Num As Long
    Dim Err_Dec As String
    
    Control.Clear
    
    ReDim sValue(1)
    
    sValue(0) = "0"
    sValue(1) = Store.Code
    
    Set Rs = New ADODB.Recordset
    Set Rs = ExecPro("SP_M_07000_01", sValue(), Err_Num, Err_Dec)

    Control.AddItem ""

    Do While Not Rs.EOF
        Control.AddItem "[" & Rs!외주코드 & "] " & Rs!외주명
        
        Rs.MoveNext
    Loop
    Rs.Close
    
End Sub
