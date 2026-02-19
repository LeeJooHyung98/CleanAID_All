VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rpt일일수금대장 
   Caption         =   "일일수금대장"
   ClientHeight    =   12015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16425
   StartUpPosition =   3  'Windows 기본값
   WindowState     =   2  '최대화
   _ExtentX        =   28972
   _ExtentY        =   21193
   SectionData     =   "rpt일일수금대장.dsx":0000
End
Attribute VB_Name = "rpt일일수금대장"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveReport_ReportStart()
    Me.Printer.RenderMode = 1
    
    lblDate.Caption = Format(Date, "YYYY년 MM월 DD일")
End Sub

