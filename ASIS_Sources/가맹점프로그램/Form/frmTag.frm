VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frmTag 
   BorderStyle     =   0  '없음
   Caption         =   "택번호 수정"
   ClientHeight    =   2385
   ClientLeft      =   2790
   ClientTop       =   2730
   ClientWidth     =   3705
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   2385
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   4207
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      PaneTree        =   "frmTag.frx":0000
      Begin Threed.SSPanel SSPanel2 
         Height          =   750
         Left            =   15
         TabIndex        =   8
         Top             =   1620
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   1323
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdEdit 
            Height          =   540
            Left            =   75
            TabIndex        =   2
            Top             =   90
            Width           =   1875
            _Version        =   851970
            _ExtentX        =   3307
            _ExtentY        =   952
            _StockProps     =   79
            Caption         =   "택번호 수정(&E)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
         End
         Begin XtremeSuiteControls.PushButton cmdCancel 
            Height          =   555
            Left            =   2475
            TabIndex        =   3
            Top             =   90
            Width           =   1125
            _Version        =   851970
            _ExtentX        =   1984
            _ExtentY        =   979
            _StockProps     =   79
            Caption         =   "취소"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   1185
         Left            =   15
         TabIndex        =   6
         Top             =   420
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   2090
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin CSTextLibCtl.sitxEdit txtTag1 
            Height          =   615
            Left            =   600
            TabIndex        =   0
            Top             =   300
            Width           =   795
            _Version        =   262145
            _ExtentX        =   1402
            _ExtentY        =   1085
            _StockProps     =   125
            Text            =   "__"
            ForeColor       =   -2147483640
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   20.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            EOLTab          =   -1  'True
            FocusSelect     =   -1  'True
            Insert          =   0   'False
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   "__"
            StartText.x     =   3
            StartText.y     =   4
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   33
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   "##"
            Justification   =   1
            CharacterTable  =   ""
            Characters      =   2
            MaxLength       =   2
         End
         Begin CSTextLibCtl.sitxEdit txtTag2 
            Height          =   615
            Left            =   1845
            TabIndex        =   1
            Top             =   300
            Width           =   1305
            _Version        =   262145
            _ExtentX        =   2302
            _ExtentY        =   1085
            _StockProps     =   125
            Text            =   "____"
            ForeColor       =   -2147483640
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   20.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            EOLTab          =   -1  'True
            FocusSelect     =   -1  'True
            Insert          =   0   'False
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   "____"
            StartText.x     =   3
            StartText.y     =   4
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   33
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   "####"
            Justification   =   1
            CharacterTable  =   ""
            Characters      =   2
            MaxLength       =   4
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   36
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1485
            TabIndex        =   7
            Top             =   270
            Width           =   255
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   390
         Left            =   15
         TabIndex        =   5
         Top             =   15
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   688
         _Version        =   262144
         Font3D          =   3
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "    택번호 수정"
         PictureBackgroundStyle=   2
         PictureBackground=   "frmTag.frx":0072
         BorderWidth     =   0
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Image Image1 
            Height          =   240
            Left            =   60
            Picture         =   "frmTag.frx":04D4
            Top             =   75
            Width           =   240
         End
      End
   End
End
Attribute VB_Name = "frmTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{tab}"
        KeyCode = 0
    End If
End Sub

Private Sub Form_Load()
    If Len(frm접수.cmdTagNo.Caption) = 7 Then
        txtTag1.Text = Left(frm접수.cmdTagNo.Caption, 2)
        txtTag2.Text = Right(frm접수.cmdTagNo.Caption, 4)
    End If
End Sub

Private Sub txtTag1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtTag2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cmdEdit_Click()
    Dim strDate   As String
    Dim strTagNo  As String
    Dim tmp       As String
    
    ' 마지막 택번호보다 이전 택번호를 입력시 택번호수정으로 보고 확인한다.
    ' 잘못 수정하여 정상적인 택번호 삭제를 막기위해서..
    
    strDate = Format(DateAdd("d", -6, Date), "YYYY-MM-DD")
    
    
    If Not IsNumeric(txtTag1.Text) Or Len(txtTag1.Text) <> 2 Then
        Query = "택번호의 시작 번호는 반드시 2자리 숫자여야 합니다." & Space(10) & vbLf & "다시 입력하여 주십시요"
        MsgBox Query, vbCritical
        
        Exit Sub
    End If
    If Not IsNumeric(txtTag2.Text) Or Len(txtTag2.Text) <> 4 Then
        Query = "택번호의 종료 번호는 반드시 4자리 숫자여야 합니다." & Space(10) & vbLf & "다시 입력하여 주십시요"
        MsgBox Query, vbCritical
        
        Exit Sub
    End If
    
    
    ' 택번호 입력 내용 확인
    If InStr(txtTag2.Text, "_") > 0 Then
        tmp = ""
        
        For i = 1 To 3
            If Mid(txtTag2.Text, i, 1) <> "_" Then
                tmp = tmp + Mid(txtTag2.Text, i, 1)
            End If
        Next i
        
        strTagNo = Trim(txtTag1.Text) + "-" + Format(tmp, "0000")
    Else
        strTagNo = Trim(txtTag1.Text) + "-" + Format(txtTag2.Text, "0000")
    End If
    
    '------------------------------------------------------------------------
    ' 마지막 택번호를 확인한다. (최초 가맹점 정보가 없을 경우에도 실행된다.)
    '------------------------------------------------------------------------
    Query = "SELECT 택번호 FROM TB_기본정보"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If ADORs.EOF Then
        Query = "UPDATE TB_기본정보 SET 택번호 = '" & 가맹점정보.택코드 & Replace(strTagNo, "-", "") & "'"
        ADOCon.Execute Query
    End If
    ADORs.Close
    Set ADORs = Nothing
    
    '-----------------------------------------------------------------------
    ' 저장 택번호보다 수정택번호가 클경우 무시하고 수정택번호 적용
    ' 작을경우 수정 택번호가 판매취소된 택번호인지 확인후 처리
    '-----------------------------------------------------------------------
    Query = "SELECT * FROM TB_입출고 "
    Query = Query & "  WHERE 접수일자 >= '" & strDate & "'"
    Query = Query & "    AND 택번호    = '" & 가맹점정보.택코드 & Replace(strTagNo, "-", "") & "'"
    Set ADORs = New ADODB.RecordSet
    ADORs.Open Query, ADOCon, adOpenForwardOnly, adLockReadOnly
    
    If Not ADORs.EOF Then
        If ADORs!판매취소 = "Y" Then
            '
        Else
            ' IsNull(ADORs.Fields("판매취소")) 인경우도 정상이기 때문에 이쪽이 실행된다.
            
                    Query = "'" & Format(strDate, "YYYY년 MM월 DD일") & "' 부터 확인한 결과 " & vbNewLine & vbNewLine
            Query = Query & "'" & strTagNo & "' 택번호는 이미 사용 하였습니다.      " & vbNewLine & vbNewLine
            Query = Query & "접수일 : " & Format(ADORs!접수일자, "YYYY년 MM월 DD일") & vbNewLine
            Query = Query & "택번호 : " & strTagNo & vbNewLine
            Query = Query & "의류명 : " & ADORs!의류명 & vbNewLine & vbNewLine
            Query = Query & "택번호를 다시 입력하여 주십시요"
            MsgBox Query, vbCritical, "확인"
            
            ADORs.Close
            Set ADORs = Nothing
            
            txtTag2.SelStart = 0:   txtTag2.SelLength = 5
            txtTag2.SetFocus
            
            Exit Sub
        End If
    End If
    ADORs.Close
    Set ADORs = Nothing
    
    '2006/10/13일 추가
    If TAG_Check(strTagNo) = False Then
        Query = "현재 접수중인 택번호로 이미 사용중 입니다." & Space(10) & vbLf & "다시 입력하여 주십시요"
        MsgBox Query, vbCritical
        
        Exit Sub
    End If
    
    '-----------------------------------------------------------------
    '
    '-----------------------------------------------------------------
    Query = "UPDATE TB_기본정보 SET 택번호 = '" & 가맹점정보.택코드 & Replace(strTagNo, "-", "") & "'"
    ADOCon.Execute Query
    
    frm접수.cmdTagNo.Caption = strTagNo
    
    Unload frmTag
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
