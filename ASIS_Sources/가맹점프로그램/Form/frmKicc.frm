VERSION 5.00
Object = "{83FD3014-2044-4BA5-9B6C-F0A2482D9C0C}#1.0#0"; "kiccposiex.ocx"
Begin VB.Form frmKicc 
   BorderStyle     =   1  '단일 고정
   Caption         =   "KiccPrint"
   ClientHeight    =   1125
   ClientLeft      =   15390
   ClientTop       =   10275
   ClientWidth     =   1785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   1785
   Begin KiccPosIE.KiccPosIEX KiccPosOCX 
      Height          =   765
      Left            =   195
      TabIndex        =   0
      Top             =   210
      Width           =   750
      BF0C            =   ""
      Bmp             =   ""
      CardNo          =   ""
      CashNo          =   ""
      CommType        =   1
      Connected       =   0   'False
      Emv             =   ""
      EmvLen          =   0
      MasterClaimerText=   ""
      MasterOfferText =   ""
      PIN             =   ""
      SeqNo           =   ""
      Sign            =   ""
      SignLen         =   0
      TID             =   ""
      RfFlag          =   ""
      VAK             =   ""
      VisaClaimerText =   ""
      VisaOfferText   =   ""
      ErrMsg          =   ""
      ResMsg          =   ""
      RcvData         =   ""
      TRNO            =   ""
      Data            =   ""
      CVER            =   ""
      MVER            =   ""
      PVER            =   ""
      TMTransCount    =   0
      TMOnLineCount   =   0
      EBTransCount    =   0
      Alignment       =   2
      AutoSize        =   0   'False
      BevelInner      =   0
      BevelOuter      =   0
      BorderStyle     =   0
      Caption         =   ""
      Color           =   16777215
      Ctl3D           =   -1  'True
      UseDockManager  =   -1  'True
      DockSite        =   0   'False
      DragCursor      =   -12
      Object.DragMode        =   0
      Enabled         =   -1  'True
      FullRepaint     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   0   'False
      ParentColor     =   0   'False
      ParentCtl3D     =   -1  'True
      Object.Visible         =   -1  'True
      DoubleBuffered  =   -1  'True
      Cursor          =   0
      Protocol        =   0
      JcbClaimerText  =   ""
      JcbOfferText    =   ""
      DccTextVer      =   "00"
      CardHash        =   "$"
      SignAD          =   "0000"
      HandleValue     =   1901564
      MemberShip      =   ""
      MemberShipHex   =   ""
      TCPSVCPort      =   0
      TCPSVCActive    =   0   'False
   End
End
Attribute VB_Name = "frmKicc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SendForm As String
Dim Gubun As String
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As String) As Long


Public Function Card_Print(PrintMsg As String) As Boolean
    Dim sE       As String
    Card_Print = False
    On Error GoTo ErrRtn

    Dim TempValue As Variant
    TempValue = Split(PrintMsg, "<C>")
    Dim LoopI As Integer
    For LoopI = LBound(TempValue) To UBound(TempValue)
        If LenB(TempValue(LoopI)) > 4096 Then
            Dim TempPrint As Variant
            TempPrint = Split(TempValue(LoopI), Chr(&HA))
            Dim LoopJ As Integer
            Dim PrintString As String
            For LoopJ = LBound(TempPrint) To UBound(TempPrint)
                PrintString = PrintString & TempPrint(LoopJ) & Chr(&HA)
                If LenB(PrintString) > 4096 Then
                    KiccPosOCX.ReqSendRS232 PrintString, sE
                    PrintString = ""
                End If
                '
            Next LoopJ
            If LenB(CStr(PrintString)) > 0 Then
                KiccPosOCX.ReqSendRS232 PrintString, sE
            End If
        Else
            If TempValue(LoopI) <> "" Then
                KiccPosOCX.ReqSendRS232 TempValue(LoopI), sE
                Call Error_Msg2("인쇄정보", SendForm, "9999", CStr(TempValue(LoopI)))
            End If
        End If
        
    Next LoopI
    
    Card_Print = True
    Exit Function
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.description)
End Function

Public Function Card_Approve(sD As String, SendFormName As String) As Boolean
    Dim sE       As String
    SendForm = SendFormName
    KiccPosOCX.ReqCmd &HFD, 0, 0, sD, sE
End Function

Private Sub Form_Load()
    Dim CommPort As String
    Dim BaudRate As String
    
    Dim sE       As String
    CommPort = GetIniStr("VAN", "KS7500_CommPort", "", iniFile)
    BaudRate = GetIniStr("VAN", "KS7500_BaudRate", "", iniFile)
    Gubun = GetIniStr("VAN", "KSCAT_GUBUN", "", iniFile)
    If Gubun = "KICC" Then
        Rtn = KiccPosOCX.Open(CInt(CommPort), CLng(BaudRate), sE)
    
        If Rtn < 0 Then
            Debug.Print (KiccPosOCX.errMsg)
            KiccPosOCX.Close
            MsgBox "카드단말기 장치가 연결되어 있지 않습니다", vbCritical, "오류"
            Exit Sub
        End If
    End If
End Sub

Public Function Card_Request(sD As String, SendFormName As String) As Boolean
    Dim sE       As String
    SendForm = SendFormName
    KiccPosOCX.ReqCmd &HFD, 0, 0, "F1", sE
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Gubun = "KICC" Then
        KiccPosOCX.Close
    End If
    '
End Sub

Private Sub KiccPosOCX_OnRcvData(ByVal Cmd As Long, ByVal GCD As Long, ByVal JCD As Long, ByVal RCD As Long, ByVal RData As String, ByVal RHexData As String)
    Call Error_Msg2("카드승인정보", SendForm, "9999", RData)
    If SendForm = "frmKSNET2" Then
        Call frmKSNET2.ReceiveMsg(RData)
        SendForm = ""
    ElseIf SendForm = "frmKSNETCash" Then
        Call frmKSNETCash.ReceiveMsg(RData)
        SendForm = ""
    End If
End Sub
