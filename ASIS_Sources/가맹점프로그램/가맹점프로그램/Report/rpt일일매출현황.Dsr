VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rpt일일매출현황 
   Caption         =   "CleanAID - rpt일일매출현황 (ActiveReport)"
   ClientHeight    =   12015
   ClientLeft      =   7170
   ClientTop       =   2250
   ClientWidth     =   16425
   WindowState     =   2  '최대화
   _ExtentX        =   28972
   _ExtentY        =   21193
   SectionData     =   "rpt일일매출현황.dsx":0000
End
Attribute VB_Name = "rpt일일매출현황"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveReport_ReportStart()
    Me.Printer.RenderMode = 1
    
    lblDate.Caption = "출력일자 : " & Format(Now, "YYYY년 MM월 DD일 hh:mm:ss")
End Sub

Private Sub Detail_BeforePrint()
    
'    If Trim(dc.Field("선긋기")) = "OK" Then
'        Line100.Visible = False
'    Else
'        Line100.Visible = True
'    End If
    
    Line100.Visible = IIf(Trim(dc.Field("선긋기")) = "OK", True, False)

End Sub
