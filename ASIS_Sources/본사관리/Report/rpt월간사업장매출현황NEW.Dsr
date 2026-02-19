VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rpt월간사업장매출현황NEW 
   Caption         =   "크린에이드 - rpt월간사업장매출현황NEW (ActiveReport)"
   ClientHeight    =   10830
   ClientLeft      =   4410
   ClientTop       =   3600
   ClientWidth     =   18135
   WindowState     =   2  '최대화
   _ExtentX        =   31988
   _ExtentY        =   19103
   SectionData     =   "rpt월간사업장매출현황NEW.dsx":0000
End
Attribute VB_Name = "rpt월간사업장매출현황NEW"
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
            lblStore.Caption = .Field("지사")
        .RecordsetPattern = "//DATA"
    End With
    
    lblTime.Caption = "출력일자:" & Format(Date, "YYYY년 MM월 DD일") & "  " & Time
End Sub

