VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rpt환불관리 
   Caption         =   "크린에이드 - rpt환불관리 (ActiveReport)"
   ClientHeight    =   10830
   ClientLeft      =   1410
   ClientTop       =   3765
   ClientWidth     =   18135
   WindowState     =   2  '최대화
   _ExtentX        =   31988
   _ExtentY        =   19103
   SectionData     =   "rpt환불관리.dsx":0000
End
Attribute VB_Name = "rpt환불관리"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveReport_ReportStart()
    Me.Printer.RenderMode = 1
    
    With dc
        .RecordsetPattern = "//HEADERDATA"
            lblTitle.Caption = .Field("타이틀")
            lblStore.Caption = .Field("검색일자")
        .RecordsetPattern = "//DATA"
    End With
    
    lblTime.Caption = "출력일자:" & Format(Date, "YYYY년 MM월 DD일") & "  " & Time
End Sub

