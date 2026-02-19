VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form P_08002 
   Caption         =   "자료 수신 (MODEM)"
   ClientHeight    =   8250
   ClientLeft      =   3315
   ClientTop       =   4080
   ClientWidth     =   12795
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_08002.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8250
   ScaleWidth      =   12795
   WindowState     =   2  '최대화
   Begin Threed.SSPanel panMain 
      Align           =   1  '위 맞춤
      Height          =   9135
      Left            =   0
      TabIndex        =   2
      Top             =   435
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   16113
      _Version        =   262144
      RoundedCorners  =   0   'False
      Begin VB.ListBox lstInput 
         Height          =   8250
         Left            =   60
         TabIndex        =   12
         Top             =   720
         Width           =   15135
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   1
         Left            =   60
         TabIndex        =   5
         Top             =   60
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "작 업 건 수"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin VB.TextBox txtInput 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   8160
         TabIndex        =   10
         Top             =   60
         Width           =   3555
      End
      Begin VB.TextBox txtInput 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   5400
         TabIndex        =   9
         Top             =   60
         Width           =   1095
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   2
         Left            =   3780
         TabIndex        =   7
         Top             =   60
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "대리점 코드"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   3
         Left            =   6540
         TabIndex        =   8
         Top             =   60
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "대 리 점 명"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   4
         Left            =   60
         TabIndex        =   11
         Top             =   420
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "작 업 파 일 LIST"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin VB.Label lblCount 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  '단일 고정
         Height          =   315
         Left            =   1680
         TabIndex        =   6
         Top             =   60
         Width           =   2055
      End
   End
   Begin Threed.SSPanel panInput 
      Align           =   1  '위 맞춤
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   767
      _Version        =   262144
      RoundedCorners  =   0   'False
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   6
         Left            =   1680
         TabIndex        =   14
         Top             =   60
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   556
         _Version        =   262144
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         Begin Threed.SSOption optSelect 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   30
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   262144
            Caption         =   "모뎀"
            Value           =   -1
         End
         Begin Threed.SSOption optSelect 
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   16
            Top             =   30
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   262144
            Caption         =   "디스켓"
         End
         Begin Threed.SSOption optSelect 
            Height          =   255
            Index           =   2
            Left            =   2010
            TabIndex        =   17
            Top             =   30
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   262144
            Caption         =   "인터넷"
         End
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   5
         Left            =   60
         TabIndex        =   13
         Top             =   60
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "작 업 경 로"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
      Begin VB.TextBox txtInput 
         Height          =   315
         Index           =   0
         Left            =   6960
         TabIndex        =   3
         Top             =   60
         Width           =   3555
      End
      Begin Threed.SSCommand cmdBtn 
         Height          =   375
         Left            =   10740
         TabIndex        =   1
         Top             =   30
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   661
         _Version        =   262144
         Caption         =   "작업시작"
      End
      Begin Threed.SSPanel panCaption 
         Height          =   315
         Index           =   0
         Left            =   5340
         TabIndex        =   4
         Top             =   60
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "작 업 경 로"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
      End
   End
End
Attribute VB_Name = "P_08002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Err_Num As Long
Dim Err_Dec As String

Dim sValue() As String

Private Sub cmdBtn_Click()

    ' 모뎀/디스켓
    If optSelect(0).Value = True Or optSelect(1).Value = True Then
        Call DataSave
        
    ' 인터넷
    ElseIf optSelect(2).Value = True Then
        CCAid.Send_RecvFileAllAction
    
    End If
    
End Sub

Public Sub Display_File_List(strFileList As String)
    Dim iCnt        As Integer
    Dim ArrList     As Variant
    Dim FileInfo    As Variant
    Dim FileName    As String
    Dim FileSize    As String
    Dim AgencyName As String
    Dim AgencyCode As String
    
    On Error GoTo ERR_RTN
    
    ArrList = Split(strFileList, ";")
    
    iCnt = 0
    lstInput.Clear
    
    ' 화일명과 정보를 ListBox에 출력한다.
    Do While Len(ArrList(iCnt)) > 0
        
        FileInfo = Split(ArrList(iCnt), "^")
        If UBound(FileInfo) >= 1 Then
            FileName = CStr(FileInfo(0))
            FileSize = CStr(FileInfo(1))
            
        ElseIf UBound(FileInfo) = 0 Then
            FileName = CStr(FileInfo(0))
            FileSize = ""
        End If
        
        '========================   메일 자료  ========================
        If Mid(FileName, 1, 1) = "M" Then
            AgencyCode = Mid(FileName, 11, 3)
            AgencyName = GetAgencyName(AgencyCode)
            
            lstInput.AddItem FileName & Space(9) & _
                             Right(Space(10) & FileSize, 10) & " BYTE" & Space(10) & _
                             Mid(FileName, 2, 4) & "년 " & Mid(FileName, 6, 2) & "월 " & Mid(FileName, 8, 2) & "일" & Space(10) & _
                             "[" & AgencyCode & "] " & _
                             LeftH(AgencyName & Space(25), 25) & Space(10) & "메일"
        
        '========================   고객 자료  ========================
        ElseIf Mid(FileName, 1, 1) = "C" Then
            AgencyCode = Mid(FileName, 11, 3)
            AgencyName = GetAgencyName(AgencyCode)
            
            lstInput.AddItem FileName & Space(9) & _
                             Right(Space(10) & FileSize, 10) & " BYTE" & Space(10) & _
                             Mid(FileName, 2, 4) & "년 " & Mid(FileName, 6, 2) & "월 " & Mid(FileName, 8, 2) & "일" & Space(10) & _
                             "[" & AgencyCode & "] " & _
                             LeftH(AgencyName & Space(25), 25) & Space(10) & "고객"
        
        '========================   매출 자료  ========================
        ElseIf Mid(FileName, 1, 1) = "2" Then
            AgencyCode = Mid(FileName, 10, 3)
            AgencyName = GetAgencyName(AgencyCode)
            
            lstInput.AddItem FileName & Space(10) & _
                             Right(Space(10) & FileSize, 10) & " BYTE" & Space(10) & _
                             Mid(FileName, 1, 4) & "년 " & Mid(FileName, 5, 2) & "월 " & Mid(FileName, 7, 2) & "일" & Space(10) & _
                             "[" & AgencyCode & "] " & _
                             LeftH(AgencyName & Space(25), 25) & Space(10) & "입고"
        
        '========================   마일리지 자료  ========================
        ' pds2004 2005/03/15일 추가
        ElseIf Mid(FileName, 1, 1) = "G" Then
            AgencyCode = Mid(FileName, 11, 3)
            AgencyName = GetAgencyName(AgencyCode)
            
            lstInput.AddItem FileName & Space(10) & _
                             Right(Space(10) & FileSize, 10) & " BYTE" & Space(10) & _
                             Mid(FileName, 2, 4) & "년 " & Mid(FileName, 6, 2) & "월 " & Mid(FileName, 8, 2) & "일" & Space(10) & _
                             "[" & AgencyCode & "] " & _
                             LeftH(AgencyName & Space(25), 25) & Space(10) & "마일리지"
        
        '========================   기타 화일  ========================
        Else
            AgencyName = ""
            lstInput.AddItem Format(FileName, "!@@@@@@@@@@@@@@@@@@@") & Space(10) & _
                             Right(Space(10) & FileSize, 10) & " BYTE" & Space(10) & _
                             Mid(FileName, 1, 4) & "년 " & Mid(FileName, 5, 2) & "월 " & Mid(FileName, 7, 2) & "일" & Space(10) & _
                             "[" & "___" & "] " & _
                             LeftH(AgencyName & Space(25), 25) & Space(10) & "기타화일"
        End If
        
        If UBound(ArrList) = iCnt Then Exit Do
        iCnt = iCnt + 1
    Loop

    DoEvents
    Exit Sub
    
ERR_RTN:
    PanelsMsg Err.Description

End Sub

Public Sub Display_File_Action(strFileMode As String)
    Dim iCnt        As Integer
    Dim ArrList     As Variant
    Dim FileName    As String
    
    ' "파일명;작업내용" 으로 전달됨
    ArrList = Split(strFileMode, ";")
    If UBound(ArrList) <> 1 Then Exit Sub
    
    For iCnt = 1 To lstInput.ListCount
    
        If Trim(Left(lstInput.List(iCnt), 20)) = CStr(ArrList(0)) Then
            lstInput.List(iCnt) = lstInput.List(iCnt) & Space(10) & CStr(ArrList(1))
            Exit For
        End If
    
    Next iCnt

End Sub

Private Sub DataSave()
    Dim AgencyName As String
    Dim FileName As String
    Dim AgencyCode As String
    Dim DataType As String
    Dim AllFileName() As String
    Dim i As Integer
    Dim j As Integer

    cmdBtn.Enabled = False
    
    On Error GoTo Err_Loop
    
    i = 0
    
    lstInput.Clear
    
    FileName = Dir(txtInput(0).Text & "\*.dat")
    
'    If Filename = "" Then
''        GoSub SUB_SUGUM
'
''        MsgBox "수신된 자료가 없습니다.", vbInformation, "확인"
''        cmdBtn.Enabled = True
''
'        Exit Sub
'    End If

    
    ' 화일명과 정보를 ListBox에 출력한다.
    Do While Len(FileName) > 0
        ReDim Preserve AllFileName(0 To i)
        AllFileName(i) = FileName
        
        '========================   메일 자료  ========================
        If Mid(AllFileName(i), 1, 1) = "M" Then
            AgencyCode = Mid(AllFileName(i), 11, 3)
            AgencyName = GetAgencyName(AgencyCode)
            
            lstInput.AddItem AllFileName(i) & Space(9) & _
                             Right(Space(10) & FileLen(txtInput(0).Text & "\" & AllFileName(i)), 10) & " BYTE" & Space(10) & _
                             Mid(AllFileName(i), 2, 4) & "년 " & Mid(AllFileName(i), 6, 2) & "월 " & Mid(AllFileName(i), 8, 2) & "일" & Space(10) & _
                             "[" & AgencyCode & "] " & _
                             LeftH(AgencyName & Space(25), 25) & Space(10) & "메일"
        
        '========================   고객 자료  ========================
        ElseIf Mid(AllFileName(i), 1, 1) = "C" Then
            AgencyCode = Mid(AllFileName(i), 11, 3)
            AgencyName = GetAgencyName(AgencyCode)
            
            lstInput.AddItem AllFileName(i) & Space(9) & _
                             Right(Space(10) & FileLen(txtInput(0).Text & "\" & AllFileName(i)), 10) & " BYTE" & Space(10) & _
                             Mid(AllFileName(i), 2, 4) & "년 " & Mid(AllFileName(i), 6, 2) & "월 " & Mid(AllFileName(i), 8, 2) & "일" & Space(10) & _
                             "[" & AgencyCode & "] " & _
                             LeftH(AgencyName & Space(25), 25) & Space(10) & "고객"
        
        '========================   매출 자료  ========================
        ElseIf Mid(AllFileName(i), 1, 1) = "2" Then
            AgencyCode = Mid(AllFileName(i), 10, 3)
            AgencyName = GetAgencyName(AgencyCode)
            
            lstInput.AddItem AllFileName(i) & Space(10) & _
                             Right(Space(10) & FileLen(txtInput(0).Text & "\" & AllFileName(i)), 10) & " BYTE" & Space(10) & _
                             Mid(AllFileName(i), 1, 4) & "년 " & Mid(AllFileName(i), 5, 2) & "월 " & Mid(AllFileName(i), 7, 2) & "일" & Space(10) & _
                             "[" & AgencyCode & "] " & _
                             LeftH(AgencyName & Space(25), 25) & Space(10) & "입고"
        
        '========================   마일리지 자료  ========================
        ' pds2004 2005/03/15일 추가
        ElseIf Mid(AllFileName(i), 1, 1) = "G" Then
            AgencyCode = Mid(AllFileName(i), 11, 3)
            AgencyName = GetAgencyName(AgencyCode)
            
            lstInput.AddItem AllFileName(i) & Space(10) & _
                             Right(Space(10) & FileLen(txtInput(0).Text & "\" & AllFileName(i)), 10) & " BYTE" & Space(10) & _
                             Mid(AllFileName(i), 2, 4) & "년 " & Mid(AllFileName(i), 6, 2) & "월 " & Mid(AllFileName(i), 8, 2) & "일" & Space(10) & _
                             "[" & AgencyCode & "] " & _
                             LeftH(AgencyName & Space(25), 25) & Space(10) & "마일리지"
        
'        '========================   큐앤솔브 연동(보관 서비스)  ========================
'        ' pds2004 2006/11/05일 추가
'        ElseIf Mid(AllFileName(i), 1, 1) = "Q" Then
'            AgencyCode = Mid(AllFileName(i), 11, 3)
'            AgencyName = GetAgencyName(AgencyCode)
'
'            lstInput.AddItem AllFileName(i) & Space(10) & _
'                             Right(Space(10) & FileLen(txtInput(0).Text & "\" & AllFileName(i)), 10) & " BYTE" & Space(10) & _
'                             Mid(AllFileName(i), 2, 4) & "년 " & Mid(AllFileName(i), 6, 2) & "월 " & Mid(AllFileName(i), 8, 2) & "일" & Space(10) & _
'                             "[" & AgencyCode & "] " & _
'                             LeftH(AgencyName & Space(25), 25) & Space(10) & "보관서비스"
        
        '========================   쿠폰 관련 자료 수신  ========================
        ' pds2004 2009/04/23일 추가
        ElseIf Mid(AllFileName(i), 1, 1) = "P" Then
            AgencyCode = Mid(AllFileName(i), 11, 3)
            AgencyName = GetAgencyName(AgencyCode)
            
            lstInput.AddItem AllFileName(i) & Space(10) & _
                             Right(Space(10) & FileLen(txtInput(0).Text & "\" & AllFileName(i)), 10) & " BYTE" & Space(10) & _
                             Mid(AllFileName(i), 2, 4) & "년 " & Mid(AllFileName(i), 6, 2) & "월 " & Mid(AllFileName(i), 8, 2) & "일" & Space(10) & _
                             "[" & AgencyCode & "] " & _
                             LeftH(AgencyName & Space(25), 25) & Space(10) & "쿠폰자료"
        
        '========================   기타 화일  ========================
        Else
            AgencyName = ""
            lstInput.AddItem Format(AllFileName(i), "!@@@@@@@@@@@@@@@@@@@") & Space(10) & _
                             Right(Space(10) & FileLen(txtInput(0).Text & "\" & AllFileName(i)), 10) & " BYTE" & Space(10) & _
                             Mid(AllFileName(i), 1, 4) & "년 " & Mid(AllFileName(i), 5, 2) & "월 " & Mid(AllFileName(i), 7, 2) & "일" & Space(10) & _
                             "[" & "___" & "] " & _
                             LeftH(AgencyName & Space(25), 25) & Space(10) & "기타화일"
        End If
        
        FileName = Dir
        i = i + 1
    Loop

    DoEvents
    
'    '========================   보관 서비스 자료 적용 ========================
'    For j = 0 To i - 1
'        Filename = AllFileName(j)
'
'        If Mid(Filename, 1, 1) = "Q" And Mid(Filename, 15, 1) = "1" Then
'            If FileLen(txtInput(0).Text & "\" & AllFileName(j)) <> 0 Then
'
'                AgencyCode = Mid(Filename, 11, 3)
'                AgencyName = GetAgencyName(AgencyCode)
'
'                txtInput(1).Text = AgencyCode
'                txtInput(2).Text = AgencyName
'
'                Call PanelsMsg(Trim(AgencyName) & "의 입고데이타를 읽고 있습니다.")
'                DoEvents
'
'                Call DataSave6(Filename, AgencyCode)
'
'                lstInput.ListIndex = j
'                If Trim(AllFileName(j)) = Trim(Mid(lstInput.List(j), 1, 19)) Then
'                    lstInput.List(j) = lstInput.List(j) & Space(10) & "처리완료"
'                End If
'            Else
'                lstInput.ListIndex = j
'                If AllFileName(j) = Trim(Mid(lstInput.List(j), 1, 19)) Then
'                    lstInput.List(j) = lstInput.List(j) & Space(10) & "삭제"
'                    Kill txtInput(0).Text & "\" & AllFileName(j)
'                End If
'            End If
'        End If
'    Next j
    
    '========================   쿠폰 자료 적용 ========================
    For j = 0 To i - 1
        FileName = AllFileName(j)
        
        If Mid(FileName, 1, 1) = "P" And Mid(FileName, 15, 1) = "1" Then
            If FileLen(txtInput(0).Text & "\" & AllFileName(j)) <> 0 Then
                
                AgencyCode = Mid(FileName, 11, 3)
                AgencyName = GetAgencyName(AgencyCode)
                
                txtInput(1).Text = AgencyCode
                txtInput(2).Text = AgencyName
            
                Call PanelsMsg(Trim(AgencyName) & "의 쿠폰데이타를 읽고 있습니다.")
                DoEvents
                
                Call DataSave8(FileName, AgencyCode)
                
                lstInput.ListIndex = j
                If Trim(AllFileName(j)) = Trim(Mid(lstInput.List(j), 1, 19)) Then
                    lstInput.List(j) = lstInput.List(j) & Space(10) & "처리완료"
                End If
            Else
                lstInput.ListIndex = j
                If AllFileName(j) = Trim(Mid(lstInput.List(j), 1, 19)) Then
                    lstInput.List(j) = lstInput.List(j) & Space(10) & "삭제"
                    Kill txtInput(0).Text & "\" & AllFileName(j)
                End If
            End If
        End If
    Next j
    
    '========================   마일리지 자료 적용 ========================
    For j = 0 To i - 1
        FileName = AllFileName(j)
        
        If Mid(FileName, 1, 1) = "G" And Mid(FileName, 15, 1) = "1" Then
            If FileLen(txtInput(0).Text & "\" & AllFileName(j)) <> 0 Then
                AgencyCode = Mid(FileName, 11, 3)
                AgencyName = GetAgencyName(AgencyCode)
                
                If (Mid(FileName, 1, 1) <> "M" And Mid(FileName, 1, 1) <> "C") And Mid(FileName, 15, 1) = "1" Then
                    txtInput(1).Text = AgencyCode
                    txtInput(2).Text = AgencyName
                
                    Call PanelsMsg(Trim(AgencyName) & "의 입고데이타를 읽고 있습니다.")
                    DoEvents
                    
                    Call DataSave5(FileName, AgencyCode)
                End If
                
                lstInput.ListIndex = j
                If Trim(AllFileName(j)) = Trim(Mid(lstInput.List(j), 1, 19)) Then
                    lstInput.List(j) = lstInput.List(j) & Space(10) & "처리완료"
                End If
            Else
                lstInput.ListIndex = j
                If AllFileName(j) = Trim(Mid(lstInput.List(j), 1, 19)) Then
                    lstInput.List(j) = lstInput.List(j) & Space(10) & "삭제"
                    Kill txtInput(0).Text & "\" & AllFileName(j)
                End If
            End If
        End If
    Next j
    
    '========================   매출 자료 적용 ========================
    For j = 0 To i - 1
        FileName = AllFileName(j)
        
        If Mid(FileName, 1, 1) = "2" And Mid(FileName, 14, 1) = "1" Then
            If FileLen(txtInput(0).Text & "\" & AllFileName(j)) <> 0 Then
                AgencyCode = Mid(FileName, 10, 3)
                AgencyName = GetAgencyName(AgencyCode)
                
                If (Mid(FileName, 1, 1) <> "M" And Mid(FileName, 1, 1) <> "C") And Mid(FileName, 14, 1) = "1" Then
                    txtInput(1).Text = AgencyCode
                    txtInput(2).Text = AgencyName
                
                    Call PanelsMsg(Trim(AgencyName) & "의 입고데이타를 읽고 있습니다.")
                    DoEvents
                    
                    Call DataSave2(FileName)
                End If
                
                lstInput.ListIndex = j
                If Trim(AllFileName(j)) = Trim(Mid(lstInput.List(j), 1, 18)) Then
                    lstInput.List(j) = lstInput.List(j) & Space(10) & "처리완료"
                End If
            Else
                lstInput.ListIndex = j
                If AllFileName(j) = Trim(Mid(lstInput.List(j), 1, 19)) Then
                    lstInput.List(j) = lstInput.List(j) & Space(10) & "삭제"
                    Kill txtInput(0).Text & "\" & AllFileName(j)
                End If
            End If
        End If
    Next j

    '========================   매일 자료 적용 ========================
    For j = 0 To i - 1
        FileName = AllFileName(j)
        
        If Mid(FileName, 1, 1) = "M" And Mid(FileName, 15, 1) = "1" Then
            If FileLen(txtInput(0).Text & "\" & AllFileName(j)) <> 0 Then
                AgencyCode = Mid(FileName, 11, 3)
                AgencyName = GetAgencyName(AgencyCode)
                
                Call PanelsMsg(Trim(AgencyName) & "의 메일데이타를 읽고 있습니다.")
                DoEvents
                
                Call DataSave3(FileName)
                
                lstInput.ListIndex = j
                If Trim(AllFileName(j)) = Trim(Mid(lstInput.List(j), 1, 19)) Then
                    lstInput.List(j) = lstInput.List(j) & Space(10) & "처리완료"
                End If
            Else
                lstInput.ListIndex = j
                If AllFileName(j) = Trim(Mid(lstInput.List(j), 1, 19)) Then
                    lstInput.List(j) = lstInput.List(j) & Space(10) & "삭제"
                    Kill txtInput(0).Text & "\" & AllFileName(j)
                End If
            End If
        End If
    Next j

    '========================   회원 자료 적용 ========================
    For j = 0 To i - 1
        FileName = AllFileName(j)
        
        If Mid(FileName, 1, 1) = "C" And Mid(FileName, 15, 1) = "1" Then
            If FileLen(txtInput(0).Text & "\" & AllFileName(j)) <> 0 Then
                AgencyCode = Mid(FileName, 11, 3)
                AgencyName = GetAgencyName(AgencyCode)
                
                Call PanelsMsg(Trim(AgencyName) & "의 고객데이타를 읽고 있습니다.")
                DoEvents
                
                Call DataSave4(FileName)
                
                lstInput.ListIndex = j
                If AllFileName(j) = Trim(Mid(lstInput.List(j), 1, 19)) Then
                    lstInput.List(j) = lstInput.List(j) & Space(10) & "처리완료"
                End If
            Else
                lstInput.ListIndex = j
                If AllFileName(j) = Trim(Mid(lstInput.List(j), 1, 19)) Then
                    lstInput.List(j) = lstInput.List(j) & Space(10) & "삭제"
                    Kill txtInput(0).Text & "\" & AllFileName(j)
                End If
            End If
        End If
    Next j
    
    
'SUB_SUGUM:
'    ' 본사일 경우에만 적용 한다.
'    If Store.Code = MASTER_OFFICE_CODE Then
'        ' 2007-05-04일 적용
'        '==================================================================
'        '==================  지사의 매출 자료 적용 ========================
'        '==================================================================
'        Dim sSuGumPath  As String
''
''
''        sSuGumPath = "G:\VBProgram\크린에이드\백상본사\백상최신\Source\test\SuGum"
''        sSuGumPath = "\\192.168.10.10\SuGum"
'
'
''txtInput(0).Text
'
'        sSuGumPath = Left(txtInput(0).Text, InStr(3, txtInput(0).Text, "\") - 1) & "\SuGum"
'
'
'        'Filename = Dir(sSuGumPath & "\????_200*.dat")
'        '전부 삭제 되도록 수정
'        Filename = Dir(sSuGumPath & "\*.dat")
'        If Filename = "" Then
'
'            MsgBox "수신된 자료가 없습니다.", vbInformation, "확인"
'            cmdBtn.Enabled = True
'
'            Exit Sub
'        End If
'
'
'
'        i = 0
'        Do While Len(Filename) > 0
'            ReDim Preserve AllFileName(0 To i)
'            AllFileName(i) = Filename
'            Filename = Dir
'            i = i + 1
'        Loop
'
'        For j = 0 To i - 1
'            Filename = AllFileName(j)
'
'            If Mid(Filename, 6, 1) = "2" And Mid(Filename, 19, 1) = "1" Then
'                If FileLen(sSuGumPath & "\" & AllFileName(j)) <> 0 Then
'                    AgencyCode = Mid(Filename, 15, 3)
'                    AgencyName = GetAgencyName(AgencyCode)
'
'                    If (Mid(Filename, 6, 1) <> "M" And Mid(Filename, 6, 1) <> "C") And Mid(Filename, 19, 1) = "1" Then
'                        txtInput(1).Text = AgencyCode
'                        txtInput(2).Text = AgencyName
'
'                        Call PanelsMsg(Trim(AgencyName) & "의 수금 데이타를 읽고 있습니다.")
'                        DoEvents
'
'                        Call DataSave7(sSuGumPath, Filename)
'                    End If
'
'                    lstInput.AddItem Filename & " 처리 완료"
'                Else
'                    'Kill Filename
'                    lstInput.AddItem Filename & " 파일 삭제"
'                End If
'            Else
'                Kill sSuGumPath & "\" & Filename
'            End If
'        Next j
'
'        'Kill sSuGumPath & "\*.dat"
'    End If

End_Loop:
    MsgBox "자료수신작업이 완료되었습니다.", vbInformation, "확인"
        
    'cmdBtn.Enabled = True
    Kill txtInput(0).Text & "\" & "*.dat"
    Exit Sub
    
Err_Loop:
    If Err.Number = 70 Then
        Resume Next
    ElseIf Err.Number = 9 Or Err.Number = 55 Then
        MsgBox "수신받은 데이터가 잘못되었으니 삭제후 재전송을 받으시기 바랍니다.", vbInformation
        Close #1
        'Kill sSuGumPath & "\" & Filename
        Resume Next
    End If
    
    'If Err.Number = 55 Then
    
    MsgBox "수신된 자료 작업시 오류가 발생하였습니다." & Chr(13) & _
           "오류코드 : " & VBA.Err.Number & Chr(13) & _
           "오류설명 : " & VBA.Err.Description, vbExclamation, "오류"
           
    Exit Sub
End Sub

Private Sub DataSave2(FileName As String)
    '========================   매출 자료 적용 ========================
    Dim iCnt As Integer
    Dim Str As String
    Dim tCnt As Integer
    Dim TempStr As String
    Dim strCode As String
    Dim strDate As String
    Dim uCnt As Integer
    
    Dim sSuData() As String
    
    Dim JobPath As String
    Dim BCPPath As String
    Dim BackupPath As String
    
    Open txtInput(0).Text & "\" & FileName For Input As #1
    
    JobPath = GetIniStr("TEXT DATA", "ReceiveJobFilePath", "", m_iniFile)
    BCPPath = GetIniStr("TEXT DATA", "BCPPath", "", m_iniFile)
    BackupPath = GetIniStr("TEXT DATA", "BackupJobFilePath", "", m_iniFile)
    
    Open JobPath & "\IpChul.Dat" For Output As #2
    
    iCnt = 0
        
    Line Input #1, TempStr
    If Not Mid(TempStr, 1, 4) = "일일마감" Then
        TempStr = ""
        Close #1
        Open txtInput(0).Text & "\" & FileName For Input As #1
    End If
    
    While Not EOF(1)
        Line Input #1, Str
        
        If TempStr = "" Then
            Print #2, Str & "|4"
        Else
            Print #2, Str & "|3|"
        End If
        
        iCnt = iCnt + 1
        lblCount = iCnt
        DoEvents
    Wend
    
    Close #1
    Close #2
    
    ' 임시 입출고 테이블의 내역을 삭제한다.
    ReDim sValue(0)
    
    sValue(0) = "1"
    Call ExecPro("SP_08001_00", sValue(), Err_Num, Err_Dec)
    
    If iCnt > 0 Then
        ' 임시 입출고 테이블에 데이터를 INSERT한다.
'        If Not Dir(BCPPath & "\OK.OK") = "" Then
'            Kill BCPPath & "\OK.OK"
'        End If
'
'        Shell BCPPath & "\IpChul.Bat", vbHide
'
'        Do While Dir(BCPPath & "\OK.OK") = ""
'            DoEvents
'        Loop
'
'        Kill BCPPath & "\OK.OK"
'
'        Call PanelsMsg("입고자료 UPDATE 중..!")
'
'        DoEvents
'
'        uCnt = 0
        
        Dim sString() As String
        Dim k As Integer
        
        ReDim sValue(18)
        
        Open JobPath & "\IpChul.Dat" For Input As #1
            
        Do While Not EOF(1)
            k = k + 1
            
            lblCount = k & " / " & iCnt
            DoEvents
            Str = ""
            Line Input #1, Str
            Debug.Print Str
            
            sString = Split(Str, "|")
            
            If UBound(sString, 1) = 18 Or UBound(sString, 1) = 17 Then
                If UBound(sString, 1) = 18 Then
                    sValue(0) = sString(0)
                    sValue(1) = sString(1)
                    sValue(2) = sString(2)
                    sValue(3) = sString(3)
                    sValue(4) = sString(4)
                    sValue(5) = sString(5)
                    sValue(6) = sString(6)
                    sValue(7) = sString(7)
                    sValue(8) = sString(8)
                    sValue(9) = sString(9)
                    sValue(10) = sString(10)
                    sValue(11) = sString(11)
                    sValue(12) = sString(12)
                    sValue(13) = sString(13)
                    sValue(14) = sString(14)
                    sValue(15) = sString(15)
    '               If sValue(15) = "Y" Then MsgBox FileName
                    sValue(16) = sString(17)
                    sValue(17) = sString(16)
                    sValue(18) = sString(18)
                Else
                    sValue(0) = sString(0)
                    sValue(1) = sString(1)
                    sValue(2) = sString(2)
                    sValue(3) = sString(3)
                    sValue(4) = sString(4)
                    sValue(5) = sString(5)
                    sValue(6) = sString(6)
                    sValue(7) = sString(7)
                    sValue(8) = sString(8)
                    sValue(9) = sString(9)
                    sValue(10) = sString(10)
                    sValue(11) = sString(11)
                    sValue(12) = sString(12)
                    sValue(13) = sString(13)
                    sValue(14) = sString(14)
                    sValue(15) = sString(15)
                    sValue(16) = sString(16)
                    sValue(17) = sString(17)
                    sValue(18) = ""
                End If
                Call ExecPro("SP_08001_11", sValue(), Err_Num, Err_Dec)
                
                If Err_Num <> 0 Then
                    MsgBox "입고 데이터 INSERT시 Error발생... " & Chr(13) & "[" & Err_Num & "] " & Err_Dec
                End If
            Else
                PanelsMsg "입고 자료 : " & Str
            End If
        Loop
            
        Close #1
        
        ' 입출고 테이블에 임시테이블의 내역을 INSERT한다.
        ReDim sValue(0)
        
        sValue(0) = "1"
        Call ExecPro("SP_08001_01", sValue(), Err_Num, Err_Dec)
        
        If Err_Num <> 0 Then
            MsgBox "입고 데이터 INSERT시 Error발생... " & Chr(13) & "[" & Err_Num & "] " & Err_Dec
            Exit Sub
        End If
        
        '일일수금 자료 생성
        If Not TempStr = "" Then
            sSuData = Split(TempStr, "|")
        
            ReDim sValue(9)
            
            sValue(0) = Mid(FileName, 1, 8)     ' 수금일자
            sValue(1) = Mid(FileName, 10, 3)    ' 대리점코드
            sValue(2) = sSuData(1)              ' 입고수량
            sValue(3) = sSuData(9)              ' 시작TAG
            sValue(4) = sSuData(10)             ' 종료TAG
            sValue(5) = sSuData(6)              ' 금액
            sValue(6) = sSuData(3)              ' 재세탁수량
            sValue(7) = sSuData(4)              ' 수선수량
            sValue(8) = sSuData(2)              ' 반품수량
            sValue(9) = 0                       ' 출고수량
            
            Call ExecPro("SP_08001_02", sValue(), Err_Num, Err_Dec)
        
            If Err_Num <> 0 Then
                MsgBox "수금 데이터 INSERT시 Error발생... " & Chr(13) & "[" & Err_Num & "] " & Err_Dec
                Exit Sub
            End If
        Else
            ReDim sValue(0)

            sValue(0) = "1"

            Call ExecPro("SP_08001_03", sValue(), Err_Num, Err_Dec)

            If Err_Num <> 0 Then
                MsgBox "수금 데이터 INSERT시 Error발생... " & Chr(13) & "[" & Err_Num & "] " & Err_Dec
                Exit Sub
            End If
        End If
    End If
    
    On Error Resume Next
    
    FileCopy txtInput(0).Text & "\" & FileName, BackupPath & "\" & FileName
    Kill txtInput(0).Text & "\" & FileName
    
    On Error GoTo 0
End Sub

Private Sub DataSave3(FileName As String)
    Dim iCnt As Integer
    Dim Str As String
    Dim tCnt As Integer
    Dim TempStr As String
    Dim strCode As String
    Dim strDate As String
    Dim uCnt As Integer
    
    Dim JobPath As String
    Dim BCPPath As String
    Dim BackupPath As String
    
    Open txtInput(0).Text & "\" & FileName For Input As #1
    
    JobPath = GetIniStr("TEXT DATA", "ReceiveJobFilePath", "", m_iniFile)
    BCPPath = GetIniStr("TEXT DATA", "BCPPath", "", m_iniFile)
    BackupPath = GetIniStr("TEXT DATA", "BackupJobFilePath", "", m_iniFile)
    
    Open JobPath & "\Mail.Dat" For Output As #2
    
    iCnt = 0
    
    While Not EOF(1)
        Line Input #1, Str
        
        Print #2, Str
        
        iCnt = iCnt + 1
        lblCount = iCnt
    Wend
    
    Close #1
    Close #2
    
    If iCnt > 0 Then
        If Not Dir(BCPPath & "\OK.OK") = "" Then
            Kill BCPPath & "\OK.OK"
        End If
        
        If Dir(BCPPath & "\Mail.FMT") = "" Then
            Shell BCPPath & "\Mail.Bat", vbHide
        Else
            Dim Para As String
            'c:\백상\BCP\BCP LAUNDRY..CustomCT in C:\백상\Data\Receive\Custom.Dat -fC:\백상\BCP\Custom.FMT -Usa -P -SCleanAid
            'COPY C:\백상\BCP\Custom.Bat C:\백상\BCP\OK.OK
            Para = BCPPath & "\BCP "
            Para = Para & DBCatalog & "..MailT "
            Para = Para & "in " & JobPath & "\Mail.Dat "
            Para = Para & "-f" & BCPPath & "\Mail.FMT "
            Para = Para & "-U" & DBUserID & " "
            Para = Para & "-P" & DBUserPwd & " "
            Para = Para & "-S" & DBServer & " "
            Para = Para & "-o" & BCPPath & "\OK.OK "


            Shell Para, vbHide

        End If
        
        
        Do While Dir(BCPPath & "\OK.OK") = ""
            DoEvents
        Loop
        
        Kill BCPPath & "\OK.OK"
        
        Call PanelsMsg("메일자료 UPDATE 중..!")
        
        DoEvents
    End If
    
    On Error Resume Next
    
    FileCopy txtInput(0).Text & "\" & FileName, BackupPath & "\" & FileName
    Kill txtInput(0).Text & "\" & FileName
    
    On Error GoTo 0
End Sub

Private Sub DataSave4(FileName As String)
    Dim iCnt As Integer
    Dim Str As String
    Dim tCnt As Integer
    Dim TempStr As String
    Dim strCode As String
    Dim strDate As String
    Dim uCnt As Integer
    
    Dim JobPath As String
    Dim BCPPath As String
    Dim BackupPath As String
    
    Open txtInput(0).Text & "\" & FileName For Input As #1
    
    JobPath = GetIniStr("TEXT DATA", "ReceiveJobFilePath", "", m_iniFile)
    BCPPath = GetIniStr("TEXT DATA", "BCPPath", "", m_iniFile)
    BackupPath = GetIniStr("TEXT DATA", "BackupJobFilePath", "", m_iniFile)
    
    Open JobPath & "\Custom.Dat" For Output As #2
    
    iCnt = 0
    
    While Not EOF(1)
        Line Input #1, Str
        
        Print #2, Str & "||"
        
        iCnt = iCnt + 1
        lblCount = iCnt
    Wend
    
    Close #1
    Close #2
    
    If iCnt > 0 Then
        If Not Dir(BCPPath & "\OK.OK") = "" Then
            Kill BCPPath & "\OK.OK"
        End If
        If Dir(BCPPath & "\Custom.FMT") = "" Then
            Shell BCPPath & "\Custom.Bat", vbHide
        Else
            Dim Para As String
            'c:\백상\BCP\BCP LAUNDRY..CustomCT in C:\백상\Data\Receive\Custom.Dat -fC:\백상\BCP\Custom.FMT -Usa -P -SCleanAid
            'COPY C:\백상\BCP\Custom.Bat C:\백상\BCP\OK.OK
            Para = BCPPath & "\BCP "
            Para = Para & DBCatalog & "..CustomCT "
            Para = Para & "in " & JobPath & "\Custom.Dat "
            Para = Para & "-f" & BCPPath & "\Custom.FMT "
            Para = Para & "-U" & DBUserID & " "
            Para = Para & "-P" & DBUserPwd & " "
            Para = Para & "-S" & DBServer & " "
            Para = Para & "-o" & BCPPath & "\OK.OK "

            Shell Para, vbHide

        End If
        Do While Dir(BCPPath & "\OK.OK") = ""
            DoEvents
        Loop

        Kill BCPPath & "\OK.OK"

        Call PanelsMsg("고객자료 UPDATE 중..!")

        DoEvents
    End If
    
    'On Error Resume Next
    
    FileCopy txtInput(0).Text & "\" & FileName, BackupPath & "\" & FileName
    
    Kill txtInput(0).Text & "\" & FileName
    
    'On Error GoTo 0
End Sub


Private Sub DataSave5(FileName As String, AgencyCode As String)
    '========================   매출 자료 적용 ========================
    Dim iCnt As Integer
    Dim Str As String
    Dim tCnt As Integer
    Dim TempStr As String
    Dim strCode As String
    Dim strDate As String
    Dim uCnt As Integer
    Dim iErrCnt As Long
    
    Dim sSuData() As String
    
    Dim JobPath As String
    Dim BCPPath As String
    Dim BackupPath As String
    
    Open txtInput(0).Text & "\" & FileName For Input As #1
    
    JobPath = GetIniStr("TEXT DATA", "ReceiveJobFilePath", "", m_iniFile)
    BCPPath = GetIniStr("TEXT DATA", "BCPPath", "", m_iniFile)
    BackupPath = GetIniStr("TEXT DATA", "BackupJobFilePath", "", m_iniFile)
    
    Open JobPath & "\IpMileage.Dat" For Output As #2
    
    iCnt = 0:   iErrCnt = 0
        
    ' 2줄은 참고 내용이기 때문에 무시한다.
    Line Input #1, TempStr
    Line Input #1, TempStr
    
    While Not EOF(1)
        Line Input #1, Str
        
        If Trim(TempStr) <> "" Then
            Print #2, Str
        End If
        
        iCnt = iCnt + 1
        lblCount = iCnt
        DoEvents
    Wend
    
    Close #1
    Close #2
    
    
    If iCnt > 0 Then
'        Call PanelsMsg("입고자료 UPDATE 중..!")
        
        Dim sString() As String
        Dim k As Integer
        
        ReDim sValue(8)
        
        Open JobPath & "\IpMileage.Dat" For Input As #1
            
        Do While Not EOF(1)
            k = k + 1
            
            lblCount = k & " / " & iCnt
            DoEvents
            
            Line Input #1, Str
            
            sString = Split(Str, "|")
            
            If UBound(sString, 1) = 8 And sString(0) = "마일리지스토리" Then
                sValue(0) = sString(0)
                sValue(1) = AgencyCode
                sValue(2) = sString(1)
                sValue(3) = sString(2)
                sValue(4) = sString(3)
                sValue(5) = sString(4)
                sValue(6) = sString(5)
                sValue(7) = sString(6)
                sValue(8) = " "
                
            ElseIf UBound(sString, 1) = 9 And sString(0) = "마일리지현황" Then
                sValue(0) = sString(0)
                sValue(1) = AgencyCode
                sValue(2) = sString(1)
                sValue(3) = sString(2)
                sValue(4) = sString(3)
                sValue(5) = sString(4)
                sValue(6) = sString(5)
                sValue(7) = sString(6)
                sValue(8) = sString(7)
            End If
            
            Call ExecPro("SP_08001_12", sValue(), Err_Num, Err_Dec)
            
            If Err_Num <> 0 Then iErrCnt = iErrCnt + 1
        Loop
            
        Close #1
    End If
    
    On Error Resume Next
    
    FileCopy txtInput(0).Text & "\" & FileName, BackupPath & "\" & FileName
    Kill txtInput(0).Text & "\" & FileName
    
    If iErrCnt > 0 Then
        Call PanelsMsg(CStr(iErrCnt) & "건이 중복 되었습니다.... " & "[" & Err_Num & "] " & Err_Dec)
    
    End If
    
    On Error GoTo 0
End Sub

Private Sub DataSave6(FileName As String, AgencyCode As String)
    '========================   보관 서비스 자료 적용 ========================
    Dim iCnt As Integer
    Dim Str As String
    Dim tCnt As Integer
    Dim TempStr As String
    Dim strCode As String
    Dim strDate As String
    Dim uCnt As Integer
    Dim iErrCnt As Long
    
    Dim FNumber1    As Integer
    Dim FNumber2    As Integer
    
    Dim sSuData() As String
    
    Dim JobPath As String
    Dim BCPPath As String
    Dim BackupPath As String
    
    
    On Error GoTo DataSave6_Error

    FNumber1 = FreeFile
    Open txtInput(0).Text & "\" & FileName For Input As #FNumber1
    
    JobPath = GetIniStr("TEXT DATA", "ReceiveJobFilePath", "", m_iniFile)
    BCPPath = GetIniStr("TEXT DATA", "BCPPath", "", m_iniFile)
    BackupPath = GetIniStr("TEXT DATA", "BackupJobFilePath", "", m_iniFile)
    
    FNumber2 = FreeFile
    Open JobPath & "\tempQN.Dat" For Output As #FNumber2
    
    iCnt = 0:   iErrCnt = 0
        
    While Not EOF(1)
        Line Input #FNumber1, Str
        
        Print #FNumber2, Str
        
        iCnt = iCnt + 1
        lblCount = iCnt
        DoEvents
    Wend
    
    Close #FNumber1
    Close #FNumber2
    
    
    If iCnt > 0 Then
'        Call PanelsMsg("입고자료 UPDATE 중..!")
        
        Dim sString() As String
        Dim k As Integer
        
        
        FNumber1 = FreeFile
        Open JobPath & "\tempQN.Dat" For Input As #FNumber1
            
        Do While Not EOF(1)
            k = k + 1
            
            lblCount = k & " / " & iCnt
            DoEvents
            
            Line Input #FNumber1, Str
            sString = Split(Str, "|")
            
            ' 보관 리스트일 경우
            If UBound(sString, 1) = 16 And sString(0) = "보관리스트" Then
                ReDim sValue(18)
                
                sValue(0) = AgencyCode
                sValue(1) = Store.Code
                sValue(2) = Trim(sString(1))
                sValue(3) = Trim(sString(2))
                sValue(4) = Trim(sString(3))
                sValue(5) = Trim(sString(4))
                sValue(6) = Trim(sString(5))
                sValue(7) = Trim(sString(6))
                sValue(8) = Trim(sString(7))
                sValue(9) = Trim(sString(8))
                sValue(10) = Trim(sString(9))
                sValue(11) = Trim(sString(10))
                sValue(12) = Trim(sString(11))
                sValue(13) = Trim(sString(12))
                sValue(14) = Trim(sString(13))
                sValue(15) = Trim(sString(14))
                sValue(16) = Trim(sString(15))
                sValue(17) = Trim(sString(16))
                
                Call ExecPro("SP_08002_01", sValue(), Err_Num, Err_Dec)
                If Err_Num <> 0 Then
                    MsgBox Err_Dec, vbCritical
                End If
            
            ' 보관 상품 리스트일 경우
            ElseIf UBound(sString, 1) = 15 And sString(0) = "보관상품리스트" Then
                           
                ReDim sValue(16)
                
                sValue(0) = AgencyCode
                sValue(1) = Store.Code
                sValue(2) = Trim(sString(1))
                sValue(3) = Trim(sString(2))
                sValue(4) = Trim(sString(3))
                sValue(5) = Trim(sString(4))
                sValue(6) = Trim(sString(5))
                sValue(7) = Trim(sString(6))
                sValue(8) = Trim(sString(7))
                sValue(9) = Trim(sString(8))
                sValue(10) = Trim(sString(9))
                sValue(11) = Trim(sString(10))
                sValue(12) = IIf(Trim(Replace(sString(11), ",", "")) = "", "0", Trim(Replace(sString(11), ",", "")))
                sValue(13) = Trim(sString(12))
                sValue(14) = Trim(sString(13))
                sValue(15) = IIf(Trim(Replace(sString(14), ",", "")) = "", "0", Trim(Replace(sString(14), ",", "")))
                    
                Call ExecPro("SP_08002_02", sValue(), Err_Num, Err_Dec)
                If Err_Num <> 0 Then
                    MsgBox Err_Dec, vbCritical
                End If
            
            ' 보관 하자 리스트일 경우
            ElseIf UBound(sString, 1) = 6 And sString(0) = "보관하자리스트" Then
                           
                ReDim sValue(5)
                
                sValue(0) = AgencyCode
                sValue(1) = Store.Code
                sValue(2) = Trim(sString(1))
                sValue(3) = Trim(sString(2))
                sValue(4) = Trim(sString(3))
                sValue(5) = Trim(sString(4))
                sValue(6) = Trim(sString(5))
                
                Call ExecPro("SP_08002_03", sValue(), Err_Num, Err_Dec)
                If Err_Num <> 0 Then
                    MsgBox Err_Dec, vbCritical
                End If
            End If
   
            If Err_Num <> 0 Then iErrCnt = iErrCnt + 1
        Loop
            
        Close #FNumber1
    End If
    
    On Error Resume Next
    
    FileCopy txtInput(0).Text & "\" & FileName, BackupPath & "\" & FileName
    Kill txtInput(0).Text & "\" & FileName
    
    On Error GoTo 0

    On Error GoTo 0
    Exit Sub

DataSave6_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DataSave6 of Form P_08002"
    Resume
End Sub


Private Sub DataSave7(ByVal sPath As String, ByVal FileName As String)
    '========================   매출 자료 적용 ========================
    Dim iCnt As Integer
    Dim Str As String
    Dim tCnt As Integer
    Dim TempStr As String
    Dim strCode As String
    Dim strDate As String
    Dim uCnt As Integer
    
    Dim sSuData() As String
    
    Dim JobPath As String
    Dim BCPPath As String
    Dim BackupPath As String
    
    If Dir(sPath & "\" & FileName, vbDirectory) = "" Then
        MsgBox sPath & "\" & FileName & "파일을 찾을 수 없습니다.", vbInformation, "확인"
        Exit Sub
    End If
    Open sPath & "\" & FileName For Input As #1
    
    JobPath = GetIniStr("TEXT DATA", "ReceiveJobFilePath", "", m_iniFile)
    BCPPath = GetIniStr("TEXT DATA", "BCPPath", "", m_iniFile)
    BackupPath = GetIniStr("TEXT DATA", "BackupJobFilePath", "", m_iniFile)
    
    Line Input #1, TempStr
    Close #1
    If Not Mid(TempStr, 1, 4) = "일일마감" Then
        TempStr = ""
        
        MsgBox "일일 마감 파일이 아님니다. "
        
    Else

        
        '일일수금 자료 생성
        If Not TempStr = "" Then
            sSuData = Split(TempStr, "|")
            
                ReDim Preserve sSuData(12)
                '카드대금의 금액/건수을 확인한다. (이전버전에서는 들어오지 않는다.)
                sSuData(11) = IIf(IsNumeric(sSuData(11)) = False, "0", sSuData(11))
                sSuData(12) = IIf(IsNumeric(sSuData(12)) = False, "0", sSuData(12))
            
            ReDim sValue(12)
            
            sValue(0) = Mid(FileName, 6, 8)     ' 수금일자
            sValue(1) = Mid(FileName, 1, 4)     ' 지사코드
            sValue(2) = Mid(FileName, 15, 3)    ' 대리점코드
            sValue(3) = sSuData(1)              ' 입고수량
            sValue(4) = sSuData(9)              ' 시작TAG
            sValue(5) = sSuData(10)             ' 종료TAG
            sValue(6) = sSuData(6)              ' 금액
            sValue(7) = sSuData(3)              ' 재세탁수량
            sValue(8) = sSuData(4)              ' 수선수량
            sValue(9) = sSuData(2)              ' 반품수량
            sValue(10) = 0                       ' 출고수량
            sValue(11) = sSuData(11)             ' 카드 금액
            sValue(12) = sSuData(12)             ' 카드 건수
            
            
            Call ExecPro("SP_08001_02_SUGUM", sValue(), Err_Num, Err_Dec)
        
            If Err_Num <> 0 Then
                MsgBox "지사 수금 데이터 INSERT시 Error발생... " & Chr(13) & "[" & Err_Num & "] " & Err_Dec
                Exit Sub
            End If
        Else
            MsgBox "지사 수금 데이터 파일이 아님니다. "
            Exit Sub
        End If
    End If
    
    On Error Resume Next
    Kill sPath & "\" & FileName
    
    On Error GoTo 0
End Sub

Private Sub Form_Activate()
    'pnlHeader.Caption = " " & Me.Caption & " (" & Me.Name & ")"
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrRtn
    
    txtInput(0).Text = GetIniStr("SERVER DATA", "ReceivePath", "", m_iniFile)
    txtInput(0).ToolTipText = txtInput(0).Text
    PanelsMsg ("")
    
    Exit Sub
    
ErrRtn:
    Call Error_Msg("", Err.Source, Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_08002_Flag = False
End Sub

Private Sub optSelect_Click(Index As Integer, Value As Integer)
    Select Case Index
        Case 0
            cmdBtn.Enabled = True
            txtInput(0).Text = GetIniStr("SERVER DATA", "ReceivePath", "", m_iniFile)
        Case 1
            cmdBtn.Enabled = True
            txtInput(0).Text = "A:"
            
        Case 2
        
            CCAid.TcpConnect
            If CCAid.Stats = sckConnected Then
                cmdBtn.Enabled = True
                CCAid.Send_RecvFileListAll
            Else
                cmdBtn.Enabled = False
            End If
    End Select
End Sub


Private Sub DataSave8(FileName As String, AgencyCode As String)
    '========================   쿠폰 자료 적용 ========================
    Dim iCnt As Integer
    Dim Str As String
    Dim tCnt As Integer
    Dim TempStr As String
    Dim strCode As String
    Dim strDate As String
    Dim uCnt As Integer
    Dim iErrCnt As Long
    
    Dim FNumber1    As Integer
    Dim FNumber2    As Integer
    
    Dim sSuData() As String
    
    Dim JobPath As String
    Dim BCPPath As String
    Dim BackupPath As String
    
    
    On Error GoTo DataSave6_Error

    FNumber1 = FreeFile
    Open txtInput(0).Text & "\" & FileName For Input As #FNumber1
    
    JobPath = GetIniStr("TEXT DATA", "ReceiveJobFilePath", "", m_iniFile)
    BCPPath = GetIniStr("TEXT DATA", "BCPPath", "", m_iniFile)
    BackupPath = GetIniStr("TEXT DATA", "BackupJobFilePath", "", m_iniFile)
    
    FNumber2 = FreeFile
    Open JobPath & "\tempCP.Dat" For Output As #FNumber2
    
    iCnt = 0:   iErrCnt = 0
        
    While Not EOF(1)
        Line Input #FNumber1, Str
        
        Print #FNumber2, Str
        
        iCnt = iCnt + 1
        lblCount = iCnt
        DoEvents
    Wend
    
    Close #FNumber1
    Close #FNumber2
    
    
    If iCnt > 0 Then
'        Call PanelsMsg("입고자료 UPDATE 중..!")
        
        Dim sString() As String
        Dim k As Integer
        
        
        FNumber1 = FreeFile
        Open JobPath & "\tempCP.Dat" For Input As #FNumber1
            
        Do While Not EOF(1)
            k = k + 1
            
            lblCount = k & " / " & iCnt
            DoEvents
            
            Line Input #FNumber1, Str
            sString = Split(Str, "|")
            
            ' 쿠폰 자료일 경우
            If UBound(sString, 1) = 10 Then
                ReDim sValue(9)
                
'999|20090422|999999|01012456|3000|5000|090165|박대선                        |2000|999|

                sValue(0) = Trim(sString(0))
                sValue(1) = Trim(sString(1))
                sValue(2) = Trim(sString(2))
                sValue(3) = Trim(sString(3))
                sValue(4) = Trim(sString(4))
                sValue(5) = Trim(sString(5))
                sValue(6) = Trim(sString(6))
                sValue(7) = Trim(sString(7))
                sValue(8) = Trim(sString(8))
                sValue(9) = Trim(sString(9))
                
                Call ExecPro("SP_08002_04", sValue(), Err_Num, Err_Dec)
                If Err_Num <> 0 Then
                    MsgBox Err_Dec, vbCritical
                End If
     
            End If
   
            If Err_Num <> 0 Then iErrCnt = iErrCnt + 1
        Loop
            
        Close #FNumber1
    End If
    
    On Error Resume Next
    
    FileCopy txtInput(0).Text & "\" & FileName, BackupPath & "\" & FileName
    'Kill txtInput(0).Text & "\" & Filename
    
    On Error GoTo 0
    Exit Sub

DataSave6_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DataSave8 of Form P_08002"

End Sub

