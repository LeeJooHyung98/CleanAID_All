VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form P_DT900 
   Caption         =   "DT900 - 핸드터미널 전송"
   ClientHeight    =   7680
   ClientLeft      =   5100
   ClientTop       =   3870
   ClientWidth     =   6480
   ControlBox      =   0   'False
   Icon            =   "P_DT900.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   6480
   Begin VB.Timer HanTimer 
      Left            =   5160
      Top             =   1110
   End
   Begin VB.CommandButton cmdSubBtn 
      Caption         =   "자료송신"
      Height          =   705
      Index           =   4
      Left            =   5190
      TabIndex        =   19
      Top             =   5520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5580
      Top             =   1095
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RTSEnable       =   -1  'True
      BaudRate        =   19200
   End
   Begin VB.CommandButton cmdSubBtn 
      Caption         =   "종료"
      Height          =   705
      Index           =   3
      Left            =   5430
      TabIndex        =   10
      Top             =   6915
      Width           =   885
   End
   Begin VB.CommandButton cmdSubBtn 
      Caption         =   "삭제"
      Height          =   705
      Index           =   2
      Left            =   4470
      TabIndex        =   9
      Top             =   6915
      Width           =   960
   End
   Begin VB.CommandButton cmdSubBtn 
      Caption         =   "저장"
      Height          =   705
      Index           =   1
      Left            =   3510
      TabIndex        =   8
      Top             =   6915
      Width           =   960
   End
   Begin VB.CommandButton cmdSubBtn 
      Caption         =   "자료수신"
      Height          =   705
      Index           =   0
      Left            =   2550
      TabIndex        =   7
      Top             =   6915
      Width           =   960
   End
   Begin VB.Frame Frame5 
      Height          =   405
      Left            =   105
      TabIndex        =   6
      Top             =   7215
      Width           =   2355
      Begin VB.TextBox txtInput 
         Height          =   315
         Index           =   2
         Left            =   1140
         TabIndex        =   16
         Top             =   90
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "데이타건수"
         Height          =   180
         Left            =   105
         TabIndex        =   15
         Top             =   150
         Width           =   900
      End
   End
   Begin VB.Frame Frame4 
      Height          =   390
      Left            =   105
      TabIndex        =   5
      Top             =   6825
      Width           =   2355
      Begin VB.TextBox txtInput 
         Height          =   300
         Index           =   1
         Left            =   1140
         TabIndex        =   11
         Top             =   90
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "전송파일명"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   150
         Width           =   900
      End
   End
   Begin VB.Frame Frame3 
      Height          =   495
      Left            =   105
      TabIndex        =   4
      Top             =   6330
      Width           =   6225
      Begin VB.CommandButton cmdSubBtn 
         Caption         =   "전송 취소"
         Height          =   405
         Index           =   5
         Left            =   4380
         TabIndex        =   20
         Top             =   90
         Width           =   1845
      End
      Begin VB.TextBox txtInput 
         Height          =   405
         Index           =   0
         Left            =   1095
         TabIndex        =   17
         Top             =   90
         Width           =   3330
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "환 경 설 정"
         Height          =   180
         Left            =   105
         TabIndex        =   18
         Top             =   210
         Width           =   900
      End
   End
   Begin VB.ListBox lstData 
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5340
      Left            =   105
      TabIndex        =   3
      Top             =   975
      Width           =   6225
   End
   Begin VB.Frame Frame2 
      Height          =   435
      Left            =   105
      TabIndex        =   2
      Top             =   570
      Width           =   6225
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "수  신  자  료"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2070
         TabIndex        =   13
         Top             =   120
         Width           =   1830
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   105
      TabIndex        =   0
      Top             =   15
      Width           =   6225
      Begin VB.ComboBox cboPort 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "P_DT900.frx":08CA
         Left            =   1380
         List            =   "P_DT900.frx":08CC
         Style           =   2  '드롭다운 목록
         TabIndex        =   21
         Top             =   150
         Width           =   2115
      End
      Begin VB.Frame Frame6 
         Height          =   555
         Left            =   1260
         TabIndex        =   12
         Top             =   0
         Width           =   30
      End
      Begin VB.Label Label1 
         Caption         =   "통신포트"
         Height          =   240
         Left            =   300
         TabIndex        =   1
         Top             =   225
         Width           =   750
      End
   End
End
Attribute VB_Name = "P_DT900"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DT900Mode As Integer
Public DownName As String   '파일 저장전체 경로
Public TotalCnt As Double   '전체 카운터수
Dim TmFlag  As Boolean
Dim TMCnt   As Integer
Dim HeadData As String
Dim RecordData As String

Dim STX$, ETX$, ENQ$, ENQ2$, ACK$, NAK$, CR$
Dim rBuf$
Dim ExitSubCheck As Boolean
Dim strErrFile As String
Public rFileName$
Dim cfgFileName$
Dim imsi$
Dim rPort

' 현재 파일에서읽은 자료값
Dim DT_Date As String
Dim DT_Gubun As String
Dim DT_Plu As String
Dim DT_Count As Integer
Dim DT_Price As Long
Dim DT_SaleGubun As Integer

Enum Mode_Type
    출고자료수신
End Enum

Public Sub SetMode(mode As Mode_Type)

    DT900Mode = mode

End Sub


Private Sub Port_Init()
    
    Call Port_Close
    
'       MSComm1.CommPort = 1
'       MSComm1.Settings = "115200,n,8,1"
'       MSComm1.Settings = "57600,n,8,1"
'       MSComm1.Settings = "38400,n,8,1"
'       MSComm1.Settings = "19200,n,8,1"
'       MSComm1.Settings = "9600,n,8,1"
'       MSComm1.Settings = "2400,n,8,1"
    
    MSComm1.CommPort = 1
    
    If cboPort.ListIndex >= 1 And cboPort.ListIndex <= 10 Then
        MSComm1.CommPort = Val(Mid(cboPort.Text, 4))
    End If
    
    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
    End If
End Sub
Private Sub Port_Close()
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If
End Sub

Private Sub cboPort_Click()
    Call INIWrite("TERMINAL DATA", "ComPort", CStr(cboPort.ListIndex), m_iniFile)
End Sub

Private Sub cmdSubBtn_Click(Index As Integer)
    Select Case Index
        Case 0
            '+--------------------------+
            '| 핸드터미널  -> PC        |
            '+--------------------------+
            Debug.Print "Start Time = > " & Time
            ' 리스트 내역을 Clear한다.
            lstData.Clear
            lstData.Enabled = True

            cmdSubBtn(0).Enabled = False
            cmdSubBtn(1).Enabled = False
            cmdSubBtn(2).Enabled = False
            cmdSubBtn(3).Enabled = False

            ' Serial Port를 초기화한다.
            Call Port_Init
            
            If Dir(rFileName$, vbDirectory) <> "" Then
                Kill (rFileName$)
            End If
            If Dir(strErrFile, vbDirectory) <> "" Then
                Kill (strErrFile)
            End If

            TotalCnt = 0

            HanTimer.Interval = 100
            HanTimer.Enabled = True

        Case 4
            '+--------------------------+
            '| PC -> 핸드터미널         |
            '+--------------------------+
            
            If Dir(rFileName$, vbDirectory) = "" Then
                MsgBox "전송할 파일을 찾을수 없습니다. 확인후 다시작업하여 주십시요." & vbLf & _
                "[" & rFileName$ & "]", vbInformation, "파일오류"
                Exit Sub
            End If

            Debug.Print "Start Time = > " & Time
            ' 리스트 내역을 Clear한다.
            lstData.Clear
            lstData.Enabled = True

            cmdSubBtn(1).Enabled = False
            cmdSubBtn(2).Enabled = False
            cmdSubBtn(3).Enabled = False

            ' Serial Port를 초기화한다.
            Call Port_Init
            
            TotalCnt = 0
            HanTimer.Interval = 100
            HanTimer.Enabled = True
'            Call Comm_Msg(HdRet)
        

        Case 1          ' 저장
            Dim i As Integer
            Dim sFilePath As String
            Dim Dfname As String
            Dim FileNumber As Integer
            Dim StrData     As String
            Dim SaveData    As String
            
            On Error GoTo err1
            
            If TotalCnt < 1 Then
                Exit Sub
            End If
            
            If MsgBox("수신한 데이터를 저장 하시겠습니까?", vbYesNo + vbQuestion) = vbYes Then
                sFilePath = GetIniStr("TERMINAL DATA", "TerminalFilePath", "", m_iniFile)
                If Right(sFilePath, 1) <> "\" Then sFilePath = sFilePath & "\"
                Dfname = sFilePath & txtInput(1).Text
                DownName = Dfname
                
                FileNumber = FreeFile
                Open Dfname For Output As #FileNumber
                
                ' 덴소로 저장되던 방식
                ' 123456789+123456789+123456789+
                ' 2  001023045285200      000000
                ' -                              1 : 입출고 구분   ( 2-> 정상, 3-> 반품 )
                '    ------                      6 : 날자  (yyMMdd)
                '          ---                   3 : 체인점 코드
                '             ----               4 : 텍번호
                '                 -              1 : 사용 안함
                '                  -             1 : 소품 구분     ( 1-> 소품, 0-> 정상 )
                
                ' DT-900 저장된 방식
                ' 123456789+123456789+123456789+
                ' 200001012340212119991001
                ' --------                        8 : 날자
                '         ------                  6 : 시간 (시분초)
                '               ---               3 : DT-900 메뉴 ( ex.  2->출고 1-> 의류 -> 정상 )
                '                  -------        7 : 택번호 ( 3: 대리점번호 4: 택번호 )
                
                
                ' 덴소 기준으로 변경하여 저장한다.
                For i = 0 To lstData.ListCount - 2
                    SaveData = ""
                    StrData = lstData.List(i)
                    SaveData = IIf(Mid(StrData, 17, 1) = "1", "2", "3") & Space(2)  ' 2로 변경하여준다
                    SaveData = SaveData & Mid(StrData, 3, 6)            ' 날자
                    SaveData = SaveData & Mid(StrData, 18, 3)           ' 체인점 코드
                    SaveData = SaveData & Mid(StrData, 21, 4)           ' 텍번호
                    SaveData = SaveData & "0"                           ' 사용 안함
                    SaveData = SaveData & IIf(Mid(StrData, 16, 1) = "2", "1", "0")  ' 소품 구분
                    SaveData = SaveData & "      000000"
                    Print #FileNumber, SaveData
                Next i
                
                Close FileNumber
                P_08001.saveYN = True
            End If
                Exit Sub
                
err1:
            Select Case Err
                Case 53
                
                Case Else
                    MsgBox Error, vbCritical, App.EXEName
                End Select
                
                Exit Sub
        Case 2          ' 삭제
            Dim delcnt As Integer
            Dim Response

            If MsgBox("수신한 데이터를 삭제 하시겠습니까?", vbYesNo + vbQuestion) = vbYes Then
                delcnt = lstData.ListCount
                TotalCnt = 0
                txtInput(2).Text = "총 " + Str(TotalCnt) + " 건"
                
                For i = delcnt - 1 To 0 Step -1
                    If lstData.Selected(i) = True Then
                        If Left(lstData.List(i), 4) <> "====" Then
                            lstData.RemoveItem i
                        End If
                    Else
                        TotalCnt = TotalCnt + 1
                    End If
                    
                    txtInput(2).Text = "총 " + Str(TotalCnt) + " 건"
                Next i
            End If
        Case 3          ' 종료
            Call Port_Close
            Unload Me
             
        ' 전송 취소
        Case 5
            ExitSubCheck = True
            cmdSubBtn(0).Enabled = True
            cmdSubBtn(3).Enabled = True
    End Select

End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    Dim i As Integer
    
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

    If GetIniStr("TERMINAL DATA", "ComPort", "ERROR", m_iniFile) = "ERROR" Then
        Call INIWrite("TERMINAL DATA", "TerminalFilePath", App.Path & "\Data\Terminal", m_iniFile)
        Call INIWrite("TERMINAL DATA", "TerminalRecvName", "Ibchul.dat", m_iniFile)
        Call INIWrite("TERMINAL DATA", "TerminalSendFile", "Send.txt", m_iniFile)
        Call INIWrite("TERMINAL DATA", "ComPort", "1", m_iniFile)
    End If
    
    cboPort.Clear
    cboPort.AddItem "포트 선택"
    For i = 1 To 10
        cboPort.AddItem "COM" & CStr(i)
    Next i
    
    rPort = GetIniStr("TERMINAL DATA", "ComPort", "1", m_iniFile)
    If Val(rPort) >= 1 And Val(rPort) <= 10 Then cboPort.ListIndex = Val(rPort)
    
    cmdSubBtn(1).Enabled = False
    cmdSubBtn(2).Enabled = False

    '+---------------------------------------------------------------------------
    ' 출고 자료
    ' 핸드터미널 -> PC
    If DT900Mode = 0 Then
        rFileName$ = GetIniStr("TERMINAL DATA", "TerminalRecvName", "Ibchul.dat", m_iniFile)
        Label2.Caption = "수  신  자  료"
        cmdSubBtn(4).Visible = False
        
    '+---------------------------------------------------------------------------
    ' 판매하기 위한 재고 자료 송신
    ' PC -> 핸드 터미널
    ElseIf DT900Mode = 1 Then
        rFileName$ = GetIniStr("TERMINAL DATA", "TerminalSendFile", "Send.txt", m_iniFile)
        Label2.Caption = "전  송  자  료"
        cmdSubBtn(0).Visible = False
        
    '+---------------------------------------------------------------------------
    ' 재고 조사
    ' 핸드터미널 -> PC
    ElseIf DT900Mode = 2 Then
        rFileName$ = GetIniStr("TERMINAL DATA", "TerminalRecvName", "Ibchul.dat", m_iniFile)
        Label2.Caption = "수  신  자  료"
        cmdSubBtn(4).Visible = False
    
    End If
        
    txtInput(1).Text = rFileName$
    strErrFile = Mid(rFileName$, 1, InStrRev(rFileName$, ".") - 1) & ".ERR"
    
    STX$ = Chr$(2)      '<--- 데이타의 시작부분을 부분을 알리는 표시
    ETX$ = Chr$(3)      '<--- 데이타의 종료시점을 부분을 알리는 표시
    ENQ$ = Chr$(6)      '<--- 다음 자료를 요청한다. DT-900의 카운터가 증가함(연결되었으니 데이타를 보네라는 신호인듯함.)
    ACK$ = Chr$(5)      '<--- 핸드 터미널에서 전송을 룰렀을때 들어옴
    NAK$ = Chr$(21)

    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub HanTimer_Timer()
  Dim dd$
  Dim FHandel As Integer
  
  Static starttime As Variant
  Static TCount As Long
  
    On Error GoTo HanTimer_Timer_Error

    dd$ = MSComm1.Input
    
    '+--------------------------+
    '| 핸드터미널 -> PC         |
    '+--------------------------+
    If dd$ = ACK$ Then
        
        HanTimer.Enabled = False
        
        FHandel = FreeFile
        Open rFileName$ For Output As #FHandel
        Close #FHandel
      
      '  txtMsg.Text = "데이타 수신중...."
        DoEvents
        MSComm1.InputLen = 0
        DelaySub (0.01)
        Output (ENQ$)
        ReceiveSub
        
        DelaySub (1)
        cmdSubBtn(3).Enabled = True
        DoEvents
        Debug.Print "End Time = > " & Time
        
        Exit Sub
    '+--------------------------+
    '| PC -> 핸드터미널         |
    '+--------------------------+
    ElseIf dd$ = ENQ$ Then
    
        HanTimer.Enabled = False
        
        DoEvents
        MSComm1.InputLen = 0
        DelaySub (0.01)
        Output (ACK$)
        Call SendMstJego
        
        DelaySub (1)
        cmdSubBtn(3).Enabled = True
        Debug.Print "End Time = > " & Time
        If txtInput(2).Text > "0" Then
            MsgBox "전송 완료", vbInformation, "확인"
        End If
            
        Exit Sub
        
    End If
    
    ' 시작후 10초동안 한건도 처리하지 못하면 오류로 판단하고 종료한다.
    If starttime = "" Then starttime = Time
    If Val(Format(Time - starttime, "ss")) > 10 Then
        starttime = ""
        HanTimer.Enabled = False
        cmdSubBtn(3).Enabled = True
        cmdSubBtn(0).Enabled = True
        TCount = 0
        Debug.Print "End Time = > " & Time
        Exit Sub
    End If

    On Error GoTo 0
    Exit Sub

HanTimer_Timer_Error:

    m_Error.ErrorMsg = "Error " & Err.Number & " (" & Err.Description & ") in procedure SetFTC of Form P_020400"

    If m_Error.VisibleMSG Then
        MsgBox m_Error.ErrorMsg, vbInformation, "확인 [Error]"
    End If

    If m_Error.SaveLog Then
        Call ProgramErrorLogWrite(m_Error)
    End If

    If m_Error.ResumeMode Then
        Resume
    End If

    
End Sub


Private Sub ReceiveSub()
    '+------------------------------------------+
    '| 핸드터미널 -> PC                         |
    '+------------------------------------------+
    '   2004020541234567890000100001230002
    '   --------                                ( 8자리) 일자
    '           -                               ( 1자리) 작업구분 (1.입고, 2.출고, 3.재고, 4.판매)
    '            -------------                  (13자리) 바코드 번호
    '                         ----              ( 4자리) 판매수량
    '                             -------       ( 7자리) 판매금액
    '                                    -      ( 1자리) 판매구분 (1.현금 2.카드)
    
    Dim iExit, rdd$, i, ipos
    Dim rData$, rType$
    Dim ReturnValue
    iExit = False
    rBuf$ = ""
    MSComm1.InBufferCount = 0
    ExitSubCheck = False
    

    iExit = False
    
    lstData.Clear
    txtInput(2).Text = 0
    While Not iExit
         DoEvents
         If (MSComm1.InBufferCount > 0) Or Len(rBuf$) > 0 Then
               If (MSComm1.InBufferCount > 0) Then
                      rdd$ = MSComm1.Input
                      For i = 1 To Len(rdd$)
                          Debug.Print Mid$(rdd$, i, 1);
                      Next i
                      rBuf$ = rBuf$ + rdd$
                        
               End If
               
               If Mid$(rBuf$, 1, 1) = ACK$ Then
                     DelaySub (0.02)
                     Output (ACK$)
                     rBuf$ = Mid$(rBuf$, 2)
                     Debug.Print "ENQ"
                    
               ElseIf Mid$(rBuf$, 1, 1) = ">" Then
                     Output (ACK$)
                     '"====   수신 끝   ====" 앞에 있는 ====를 다른곳에서 사용한다.
                     lstData.AddItem "====   수신 끝   ===="
                     lstData.ListIndex = lstData.ListCount - 1
                     rBuf$ = Mid$(rBuf$, ipos + 1)
                     iExit = True
                     rBuf$ = ""
                     lstData.Visible = True
                     HanTimer.Enabled = False
                     
                     Debug.Print "End   Time = > " & Time

                     On Error GoTo ERR_RTN
                     
                     If Dir(strErrFile, vbDirectory) <> "" Then
                        MsgBox "바코드 형식에 맞지않은 데이터가 있습니다.   확인 바랍니다.", vbInformation, "확인"
                        If Dir("C:\WINNT\NOTEPAD.EXE", vbDirectory) <> "" Then
                            ReturnValue = Shell("C:\WinNT\NOTEPAD.EXE " & strErrFile, 1)   ' 연결창 실행.
                            AppActivate ReturnValue    ' 메모창 활성.
                        ElseIf Dir("C:\WINDOWS\NOTEPAD.EXE", vbDirectory) <> "" Then
                            ReturnValue = Shell("C:\WINDOWS\NOTEPAD.EXE ", 1)   ' 연결창 실행.
                            AppActivate ReturnValue    ' 메모창 활성.
                        ElseIf Dir("C:\WINDOWS\system\NOTEPAD.EXE", vbDirectory) <> "" Then
                            ReturnValue = Shell("C:\WINDOWS\system\NOTEPAD.EXE " & strErrFile, 1)   ' 연결창 실행.
                            AppActivate ReturnValue    ' 메모창 활성.
                        End If
                     End If
                
                     cmdSubBtn(1).Enabled = True
                     cmdSubBtn(2).Enabled = True
                     cmdSubBtn(3).Enabled = True
                     

               Else
                     ipos = InStr(rBuf$, ETX$)
                     
                     If ipos > 0 Then
                           rData$ = Mid$(rBuf$, 2, ipos - 2)
                                
                            If Trim(rData$) <> "" Then
                               Update (rData$)
                            End If
                                                   
                           rBuf$ = Mid$(rBuf$, ipos + 1)
                           DelaySub (0.02)
                           Output (ENQ$)           '   <--- 다음 자료를 요청한다. DT-900의 카운터가 증가함
                     Else
                          
                     
                     End If
               End If
         
         
         End If
            
         If ExitSubCheck Then
            DelaySub (0.02)
'           Output (ENQ$) <- 수신중 취소 버튼을클릭 하였을 경우
            Exit Sub
         End If
    Wend
    Exit Sub
    
ERR_RTN:
    MsgBox "[" & Err.Number & "] " & Err.Description & vbCrLf
    Resume Next
    
End Sub

Sub DelaySub(iDelay As Double)
   Dim OldTime As Double, NewTime As Double
   OldTime = Timer
   
   Do
       NewTime = Timer
       If OldTime + iDelay < NewTime Then
            Exit Do
       End If
       If Abs(OldTime - NewTime) > 1000 Then
               OldTime = Timer
       End If
   Loop
   
End Sub

Private Sub Update(iData$)
   Dim iDir$, iFile$, iiiFileName$
   Dim strTemp As String
  
    ' 길이 확인
'    If Len(iData$) <> 36 Then GoTo Err_Write
    
'    ' 날짜 확인
'    strTemp = Mid(iData$, 1, 8)
'    If Not IsDate(strTemp) Then GoTo Err_Write
    
    ' 바코드 확인
'    strTemp = Right(iData$, 16)
'    If Not IsNumeric(strTemp) Then GoTo Err_Write
    
   lstData.AddItem iData$
   txtInput(0).Text = iData$
   txtInput(2).Text = txtInput(2).Text + 1
   TotalCnt = TotalCnt + 1
   lstData.ListIndex = lstData.ListCount - 1
   Open rFileName$ For Append As #1
     Print #1, iData$
   Close #1
   Exit Sub
   
Err_Write:
    strErrFile = Mid(rFileName$, 1, InStrRev(rFileName$, ".") - 1) & ".ERR"
   Open strErrFile For Append As #2
     Print #2, iData$
   Close #2
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If MSComm1.PortOpen Then
        MSComm1.PortOpen = False
    End If
   Set P_DT900 = Nothing
End Sub

Private Sub Output(iData)
   MSComm1.Output = iData
   
End Sub

Private Sub SendMstJego()
   ' 전송 데이타 포멧
   ' 바코드(13) + 금액(7) +  품명(21)
   
     Dim iExit, rData$, iCount, rrData$
     Dim FHandel As Integer
     Dim rSendCount As Long
     
    ExitSubCheck = False
    iExit = False
    
    rBuf$ = ""
    iExit = False
    rrData$ = ""
    
    ' 재고 데이타 전송
    FHandel = FreeFile
    Open rFileName$ For Input As #FHandel
    lstData.Clear
    txtInput(2).Text = 0
                                          
    While Not EOF(FHandel)
         DoEvents
         If (MSComm1.InBufferCount > 0) Or Len(rBuf$) > 0 Then
               If (MSComm1.InBufferCount > 0) Then
                      rBuf$ = rBuf$ + MSComm1.Input
               End If
               If Mid$(rBuf$, 1, 1) = ENQ$ Then
                   '  DelaySub (0.0)
                     Line Input #FHandel, rData$
                     rData$ = LeftH(rData$ + Space(51), 41)
                     lstData.AddItem rData$
                     lstData.ListIndex = lstData.ListCount - 1
                     If lstData.ListCount > 30000 Then
                            lstData.Clear
                     End If
                     
                     rSendCount = rSendCount + 1
'                     lblSendCount.Caption = rSendCount
                     
                     Output (STX$ + rData$ + ETX$)
                     rBuf$ = Mid$(rBuf$, 2)
                     txtInput(2).Text = rSendCount
               End If
         End If
         
         If ExitSubCheck Then
            Close #FHandel
            Exit Sub
         End If
         
    Wend
    
    Close #FHandel
    
    
    iExit = False
    ' END 전송
    While Not iExit
         DoEvents
         If (MSComm1.InBufferCount > 0) Or Len(rBuf$) > 0 Then
               If (MSComm1.InBufferCount > 0) Then
                      rBuf$ = rBuf$ + MSComm1.Input
               End If
    
               If Mid$(rBuf$, 1, 1) = ENQ$ Then
                     rData$ = "END"
                     Output (">")
                     iExit = True
               End If
         End If
         If ExitSubCheck Then
               Exit Sub
         End If
         
    Wend
    
End Sub
