VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rpt사고접수 
   Caption         =   "크린에이드 - rpt사고접수 (ActiveReport)"
   ClientHeight    =   12015
   ClientLeft      =   3120
   ClientTop       =   2355
   ClientWidth     =   16425
   WindowState     =   2  '최대화
   _ExtentX        =   28972
   _ExtentY        =   21193
   SectionData     =   "rpt사고접수.dsx":0000
End
Attribute VB_Name = "rpt사고접수"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveReport_ReportStart()
    Me.Printer.RenderMode = 1
    
    lblDate.Caption = "출력:" & Format(Now, "YYYY년 MM월 DD일 hh:mm:ss")
End Sub

