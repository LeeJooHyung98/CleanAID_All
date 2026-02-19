Attribute VB_Name = "PrinterFORM"
Option Explicit
    
Public RS01 As ADODB.Recordset
Public sValue() As String

Public Err_Num As Long
Public Err_Dec As String

' 출력 여백 관련
Public MagRect As RECT_TYPE        ' 출력할 좌표
Public iNextLineSpace  As Integer  ' Details의 다음 라인에 출력될 여백
Public iLeftNextSpace  As Integer  ' 한 라인에 여러개 출력될경우 다음 출력 위치

' 출력 수량 및 위치 관련
Public iLeftLooper As Integer      ' 한 라인에 여러개 출력될경우 현재 출력 갯수
Public iLeftMaxLooper  As Integer  ' 한 라인에 출력할 전체 갯수

Public iTotalPageCnt   As Integer  ' 출력할 전체 페이지 수 (대리점 별)
Public iProcPageCnt    As Integer  ' 현재 출력중인 페이지 수 (대리점 별)

Public iTotalCnt       As Integer  ' 출력할 전체 건수
Public iPageTotProcCnt As Integer  ' 한페이지당 전체 출력수 ( Details 수 )
Public iProcCnt        As Integer  ' 현재출력 중인 전체건수
Public iPageProcCnt    As Integer  ' 한페이지에서 현재 출력중인 수 ( Details 수 )
' 콤보에서 보기 비율 선택
Public iSL_Type     As Single   ' 현재 보여질 뷰의 비율을 선택함 (100% -> 1)

