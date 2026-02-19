VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rpt마일리지현황 
   Caption         =   "CleanAID - rpt마일리지현황 (ActiveReport)"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18855
   StartUpPosition =   3  'Windows 기본값
   WindowState     =   2  '최대화
   _ExtentX        =   33258
   _ExtentY        =   14870
   SectionData     =   "rpt마일리지현황.dsx":0000
End
Attribute VB_Name = "rpt마일리지현황"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveReport_ReportStart()
    Me.Printer.RenderMode = 1
    
    lblDate.Caption = Format(Date, "YYYY년 MM월 DD일")
    lblTime.Caption = Time
End Sub

