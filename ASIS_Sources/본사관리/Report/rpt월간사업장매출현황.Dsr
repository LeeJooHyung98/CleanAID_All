VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rpt월간사업장매출현황 
   Caption         =   "크린에이드 - rpt월간사업장매출현황 (ActiveReport)"
   ClientHeight    =   8055
   ClientLeft      =   210
   ClientTop       =   7290
   ClientWidth     =   21090
   WindowState     =   2  '최대화
   _ExtentX        =   37200
   _ExtentY        =   14208
   SectionData     =   "rpt월간사업장매출현황.dsx":0000
End
Attribute VB_Name = "rpt월간사업장매출현황"
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

