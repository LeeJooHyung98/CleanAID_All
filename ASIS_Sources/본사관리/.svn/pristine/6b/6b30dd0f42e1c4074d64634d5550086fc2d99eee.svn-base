VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rpt_P_04019 
   Caption         =   "크린에이드 - rpt_P_04019 (ActiveReport)"
   ClientHeight    =   8055
   ClientLeft      =   750
   ClientTop       =   6300
   ClientWidth     =   18285
   WindowState     =   2  '최대화
   _ExtentX        =   32253
   _ExtentY        =   14208
   SectionData     =   "rpt_P_04019.dsx":0000
End
Attribute VB_Name = "rpt_P_04019"
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

