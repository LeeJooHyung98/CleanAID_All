VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form P_TRANS 
   BorderStyle     =   1  '단일 고정
   Caption         =   "핸디터미널 전송"
   ClientHeight    =   7695
   ClientLeft      =   1860
   ClientTop       =   3210
   ClientWidth     =   6420
   Icon            =   "P_TRANS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   6420
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5040
      Top             =   960
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5520
      Top             =   900
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin Threed.SSPanel panMain 
      Align           =   1  '위 맞춤
      Height          =   7650
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   13494
      _Version        =   262144
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txtInput 
         Height          =   315
         Index           =   2
         Left            =   1200
         TabIndex        =   14
         Top             =   7200
         Width           =   1695
      End
      Begin VB.TextBox txtInput 
         Height          =   315
         Index           =   1
         Left            =   1200
         TabIndex        =   13
         Top             =   6840
         Width           =   1695
      End
      Begin VB.TextBox txtInput 
         Height          =   315
         Index           =   0
         Left            =   1740
         TabIndex        =   12
         Top             =   6420
         Width           =   4575
      End
      Begin Threed.SSCommand cmdSubBtn 
         Height          =   675
         Index           =   0
         Left            =   2940
         TabIndex        =   7
         Top             =   6840
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1191
         _Version        =   262144
         Caption         =   "자료수신"
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
         Height          =   5580
         Left            =   120
         MultiSelect     =   2  '확장형
         TabIndex        =   6
         Top             =   720
         Width           =   6195
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   0
         Left            =   1740
         TabIndex        =   2
         Top             =   60
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   556
         _Version        =   262144
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSOption optSelect 
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   3
            Top             =   30
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   262144
            Caption         =   "COM1"
            Value           =   -1
         End
         Begin Threed.SSOption optSelect 
            Height          =   255
            Index           =   1
            Left            =   1620
            TabIndex        =   4
            Top             =   30
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   262144
            Caption         =   "COM2"
         End
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   1
         Top             =   60
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "통 신 포 트"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   420
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "수   신   자   료"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdSubBtn 
         Height          =   675
         Index           =   1
         Left            =   3780
         TabIndex        =   8
         Top             =   6840
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1191
         _Version        =   262144
         Caption         =   "저  장"
      End
      Begin Threed.SSCommand cmdSubBtn 
         Height          =   675
         Index           =   2
         Left            =   4620
         TabIndex        =   9
         Top             =   6840
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1191
         _Version        =   262144
         Caption         =   "삭  제"
      End
      Begin Threed.SSCommand cmdSubBtn 
         Height          =   675
         Index           =   3
         Left            =   5460
         TabIndex        =   10
         Top             =   6840
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1191
         _Version        =   262144
         Caption         =   "종  료"
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   6420
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "환 경 설 정"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   6840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "전송파일명"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Top             =   7200
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "데이타건수"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
End
Attribute VB_Name = "P_TRANS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const COM1 = 1
Const COM2 = 2
Const COM3 = 3
Const COM4 = 4

Const SOH = &H1
Const STX = &H2
Const ETX = &H3
Const EOT = &H4
Const ENQ = &H5
Const ACK = &H6
Const NAK = &H15
Const ETB = &H17
Const ESC = &H1B

Const RECORD_MAX = 33
Const TITLE_MAX = 34

Const Err_01 = "[-1] 핸드 터미날 라인점검 (초기에러)"
Const Err_02 = "[-2] 핸드 터미날 수신 시간 초과"
Const Err_03 = "[-3] 핸드터미날 수신 데이터 에러 (SOH)"
Const Err_04 = "[-4] 핸드터미날 수신 데이터 에러 (ETX)"
Const Err_05 = "[-5] 핸드터미날 수신 데이터 에러 (STX)"

Const OK_01 = "[ 1] 핸드터미날 레코드 정상수신"
Const OK_03 = "[ 3] 핸드터미날 수신 성공"

Const Err_Ok = "알수 없는 에러"

Const Ready = "핸드터미날 엔터를 치시요"

Public saveYN As Boolean

Dim TmFlag  As Boolean
Dim TMCnt   As Integer
Dim HeadData As String
Dim RecordData As String
Dim TotalCnt As Integer
Dim DownName As String

Private Sub cmdSubBtn_Click(Index As Integer)
    Select Case Index
        Case 0              ' 자료수신
            Dim HdRet As Integer
            
            MsgBox "핸드터미날 엔터를 치시요", vbInformation, "자료수신"
            
            cmdSubBtn(0).Enabled = False
            cmdSubBtn(1).Enabled = False
            cmdSubBtn(2).Enabled = False
            
            ' 리스트 내역을 Clear한다.
            lstData.Clear
            lstData.Enabled = False
            
            ' Serial Port를 초기화한다.
            Call Port_Init
            
            TotalCnt = 0
            
            ' BHT2000으로 부터 데이터를 받는다.
            HdRet = Recv_Task
            
            ' Serial Port를 닫는다.
            Call Port_Close
            
            cmdSubBtn(0).Enabled = True
            cmdSubBtn(1).Enabled = True
            cmdSubBtn(2).Enabled = True
            
            lstData.Enabled = True
            
            Call Comm_Msg(HdRet)
        Case 1          ' 저장
            Dim i As Integer
            Dim sFilePath As String
            Dim Dfname As String
            
            On Error GoTo Err1
            
            If TotalCnt < 1 Then
                Exit Sub
            End If
            
            If MsgBox("수신한 데이터를 저장 하시겠습니까?", vbYesNo + vbQuestion) = vbYes Then
                sFilePath = GetIniStr("TERMINAL DATA", "TerminalFilePath", "", sIniFile)
                Dfname = sFilePath & "\" & txtInput(1).Text
                
                Open Dfname For Output As #100
                
                For i = 0 To lstData.ListCount - 1
                    Print #100, lstData.List(i)
                Next i
                
                Close 100
                saveYN = True
            End If
                Exit Sub
                
Err1:
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
                        lstData.RemoveItem i
                    Else
                        TotalCnt = TotalCnt + 1
                    End If
                    
                    txtInput(2).Text = "총 " + Str(TotalCnt) + " 건"
                Next i
            End If
        Case 3          ' 종료
            Call Port_Close
            Unload Me
    End Select
End Sub

Private Sub Port_Init()
    On Error GoTo ERR_RTN
    
    If optSelect(0).Value = True Then
        MSComm1.CommPort = 1
        MSComm1.Settings = "9600,N,8,1"
    Else
        MSComm1.CommPort = 2
        MSComm1.Settings = "9600,N,8,1"
     
    End If
    
    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
    End If
    Exit Sub
    
ERR_RTN:
    MsgBox Err.Description, vbInformation, "확인"
End Sub

Private Function Recv_Task() As Integer
    '----------------------------------------------------------------------------------------
    ' 통신 Protocol
    ' 1 - BHT2000으로부터 ENQ를 받는다.
    ' 2 - ACK를 보낸다.
    ' 3 - BHT2000으로부터 Title을 받는다. (데이터의 처음(SOH)과 마지막(ETX)을 Check한다.)
    ' 4 - ACK를 보낸다.
    ' 5 - 실데이터를 받는다. (데이터의 처음(STX)과 마지막(ETX)을 Check한다.)
    ' 6 - ACK를 보낸다.
    ' 7 - BHT2000에서 BOT가 오면 종료한다.
    '------------------------------------------------------------------------------------------
    
    Dim RecvRet As Integer
    
    ' BHT2000에서 오는 값이 ENQ인지를 Check한다.
    On Error GoTo Recv_Task_Error

    RecvRet = Recv_Ch(ENQ)
    
    ' ENQ가 아니면 종료한다.
    If RecvRet < 0 Then
        Recv_Task = RecvRet
        Exit Function
    End If
    
    ' ACK를 BHT2000에 보낸다.
    Call Send_Ch(ACK)
    
    ' Title을 내역을 받는다. - 처음이 SOH이고, 마지막이 ETX인지를 Check한다.
    RecvRet = Recv_Title(SOH, ETX)
    
    If RecvRet < 0 Then
        Recv_Task = RecvRet
        Exit Function
    End If
    
    ' ACK를 BHT2000에 보낸다.
    Call Send_Ch(ACK)
    
    If RecvRet = 3 Then
        Recv_Task = RecvRet
        Exit Function
    End If
    
    ' 실제데이터를 받는다.
    Do
        ' 실데이타의 처음(STX)와 마지막(ETX)를 Check하여서 실데이터만을 받는다.
        RecvRet = Recv_Record(STX, ETX)
        
        If RecvRet < 0 Then
            Recv_Task = RecvRet
            Exit Function
        End If
        
        ' ACK를 BHT2000에 보낸다.
        Call Send_Ch(ACK)
    Loop Until RecvRet <> 1
    
    Recv_Task = RecvRet
    Exit Function
    

    On Error GoTo 0
    Exit Function

Recv_Task_Error:

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

    
End Function

Private Function Recv_Ch(RC As Integer) As Integer
    Dim ChTmp As Byte
    Dim Recv_Tmp As String
    
    Call Back_Time(12)
    
    ' Serial Port로 부터 데이터가 들어올 때까지 Looping을 하면서 대기한다.
    Do
        DoEvents
    
        If TmFlag = False Then
            Recv_Ch = -1  ' -1 recv char timeover
            Exit Function
        End If
    Loop Until MSComm1.InBufferCount = 1
    
    ' Serial Port로 부터 받은 첫번째 값이 파라미터값과 같은지를 Check한다.
    Recv_Tmp = MSComm1.Input
    ChTmp = Asc(Mid(Recv_Tmp, 1, 1))
    
    If ChTmp <> RC Then
        Recv_Ch = -2      ' -2 recv char format data error
        Exit Function
    End If
    
    Recv_Ch = 0
End Function

Private Sub Back_Time(BTime As Integer)
    Timer1.Enabled = False
    
    BTime = BTime * 500
    Timer1.Interval = BTime
    
    Timer1.Enabled = True
    
    TmFlag = True
End Sub


Private Sub Timer1_Timer()
    TmFlag = False ' 시간초과시
End Sub

Private Sub Send_Ch(ch As Integer)
     MSComm1.Output = Chr(ch)
End Sub

Private Function Recv_Title(RC1 As Integer, RC2 As Integer) As Integer
    Dim ChTmp As Byte
    Dim Recv_Tmp As String

    HeadData = ""
    Call Back_Time(12)
    
    ' Serial Port로부터 들어오는 값이 34자리(Title의 길이)일 때 까지 Looping을한다.
    Do
        DoEvents
        If TmFlag = False Then
            Recv_Tmp = MSComm1.Input
            ChTmp = Asc(Mid(Recv_Tmp, 1, 1))
            
            If ChTmp = EOT Then
                Recv_Title = 3      ' eot end
                Exit Function
            End If
            
            Recv_Title = -2         ' -1 recv char timeover
            Exit Function
        End If
    Loop Until MSComm1.InBufferCount >= TITLE_MAX
    
    Recv_Tmp = MSComm1.Input
    
    ' Serial로부터 받은 값의 첫번째를 Check한다.
    ChTmp = Asc(Mid(Recv_Tmp, 1, 1))
    
    If ChTmp <> RC1 Then
        Recv_Title = -3             ' -2 recv first char format data error
        Exit Function
    End If
     
    ' Serial로부터 받은 값의 마지막값을 Check한다.
    ChTmp = Asc(Mid(Recv_Tmp, Len(Recv_Tmp) - 1, 1))
    
    If ChTmp <> RC2 Then
        Recv_Title = -4             ' -2 recv last char format data error
        Exit Function
    End If
    
    ' 처음과 마지막의 값을 뺀 실데이타만 만든다.
    For ChTmp = 2 To Len(Recv_Tmp) - 2 Step 1
        HeadData = HeadData + Mid(Recv_Tmp, ChTmp, 1)
    Next ChTmp
    
    ' 실데이터의 내역을 보여준다.
    txtInput(0).Text = HeadData
    txtInput(1).Text = Mid(HeadData, 1, 12)
    
    Recv_Title = 0
End Function

Private Function Recv_Record(RC1 As Integer, RC2 As Integer) As Integer
    Dim ChTmp As Byte
    Dim Recv_Tmp As String
    
    RecordData = ""
    Call Back_Time(12)
    
    ' Serial Port로부터 들어오는 값이 32자리(실데이터의 길이)일 때 까지 Looping을한다.
    Do
        DoEvents
        
        If TmFlag = False Then
            Recv_Tmp = MSComm1.Input
            ChTmp = Asc(Mid(Recv_Tmp, 1, 1))
            
            ' BHT2000으로 부터 EOT를 받으면 종료한다.
            If ChTmp = EOT Then
                Recv_Record = 3  ' eot end
                Exit Function
            End If
            
            Recv_Record = -2  ' -1 recv char timeover
            Exit Function
        End If
    Loop Until MSComm1.InBufferCount >= RECORD_MAX
   
    Recv_Tmp = MSComm1.Input
    ChTmp = Asc(Mid(Recv_Tmp, 1, 1))
    
    If ChTmp <> RC1 Then
        Recv_Record = -5      ' -2 recv first char format data error
        Exit Function
    End If
    
    ChTmp = Asc(Mid(Recv_Tmp, Len(Recv_Tmp) - 1, 1))
    
    If ChTmp <> RC2 Then
        Recv_Record = -4      ' -2 recv last char format data error
        Exit Function
    End If
    
    For ChTmp = 2 To Len(Recv_Tmp) - 2 Step 1
        RecordData = RecordData + Mid(Recv_Tmp, ChTmp, 1)
    Next ChTmp
    
    ' LIST BOX에 실데이터를 Add한다.
    lstData.AddItem (RecordData)
    
    TotalCnt = TotalCnt + 1
    
    txtInput(2).Text = "총 " + Str(TotalCnt) + " 건"
    
    Recv_Record = 1
End Function

Sub Comm_Msg(mg As Integer)
    Dim DispMsg As String
    Select Case mg
        Case -1
            DispMsg = Err_01
        Case -2
            DispMsg = Err_02
        Case -3
            DispMsg = Err_03
        Case -4
            DispMsg = Err_04
        Case -5
            DispMsg = Err_05
        Case 1
            DispMsg = OK_01
        Case 3
            DispMsg = OK_03
        Case Else
            DispMsg = Err_Ok
    End Select
    
    MsgBox DispMsg, vbInformation, "자료 수신"
End Sub

Private Sub Port_Close()
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If
End Sub
