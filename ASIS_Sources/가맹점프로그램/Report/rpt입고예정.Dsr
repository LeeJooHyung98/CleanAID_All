VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rpt입고예정 
   Caption         =   "CleanAID - rpt입고예정 (ActiveReport)"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20610
   StartUpPosition =   3  'Windows 기본값
   WindowState     =   2  '최대화
   _ExtentX        =   36354
   _ExtentY        =   13547
   SectionData     =   "rpt입고예정.dsx":0000
End
Attribute VB_Name = "rpt입고예정"
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

