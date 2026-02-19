VERSION 5.00
Begin VB.UserControl ctlMenu 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   5115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   LockControls    =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   5400
   Begin VB.PictureBox imgCloth 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   75
      Picture         =   "ctlMenu.ctx":0000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   180
      Width           =   495
   End
   Begin VB.Label lblClick 
      BackStyle       =   0  '투명
      Height          =   1680
      Left            =   2085
      TabIndex        =   0
      Top             =   330
      Width           =   1905
   End
   Begin VB.Label lblMenuName 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   120
      TabIndex        =   2
      Top             =   105
      Width           =   1620
   End
   Begin VB.Label lblPrice 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   1620
      TabIndex        =   1
      Top             =   570
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image imgIsPrint 
      Height          =   255
      Index           =   1
      Left            =   135
      Picture         =   "ctlMenu.ctx":08CA
      Top             =   2610
      Width           =   360
   End
   Begin VB.Image imgIsPrint 
      Height          =   255
      Index           =   0
      Left            =   150
      Picture         =   "ctlMenu.ctx":0DD4
      Top             =   2205
      Width           =   360
   End
   Begin VB.Image imgBack 
      Height          =   870
      Left            =   0
      Picture         =   "ctlMenu.ctx":12DE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1830
   End
End
Attribute VB_Name = "ctlMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'기본 속성 값:
Const m_def_isPrint = 0
Const m_def_GET_MenuKey = 0
Const m_def_GET_Selected = 0
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0

'속성 변수:
Dim m_isPrint As Boolean
Dim m_GET_MenuKey As Long
Dim m_GET_Selected As Boolean
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer

'이벤트 선언:
Event Click() 'MappingInfo=lblClick,lblClick,-1,Click
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim tMenuKey As String

Private Sub imgCloth_Click()
    'Picture
    lblClick_Click
End Sub

Private Sub lblClick_Click()
    RaiseEvent Click

    'lblMenuName.ForeColor = vbRed
    
    If lblMenuName.Caption = "" Then
        lblMenuName.ForeColor = vbBlack
    End If
End Sub

Private Sub UserControl_Initialize()
    On Error Resume Next
    
    Call SET_Size
    
    lblClick.Top = 0
    lblClick.Left = 0
End Sub

Private Sub SET_Size()
    On Error Resume Next
    
    imgBack.Width = UserControl.Width
    imgBack.Height = UserControl.Height
    
    lblClick.Width = imgBack.Width
    lblClick.Height = imgBack.Height
    
    lblPrice.Top = imgBack.Height - lblPrice.Height - 65
    lblPrice.Left = imgBack.Width - lblPrice.Width - 80
    
    'imgIsPrint(0).Left = 75
    'imgIsPrint(0).Top = UserControl.Height - imgIsPrint(0).Height - 60
    
    'imgIsPrint(1).Left = 75
    'imgIsPrint(1).Top = UserControl.Height - imgIsPrint(1).Height - 60
    
    'imgIsPrint(1).ZOrder 0
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    Call SET_Size
End Sub

'MemberInfo=8,0,0,0
Public Property Get BackColor() As Long
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'MemberInfo=6,0,0,0
Public Property Get Font() As Font
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'MemberInfo=5
Public Sub Refresh()
     
End Sub

'사용자 정의 컨트롤에 대한 속성을 초기화합니다.
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
    m_GET_Selected = m_def_GET_Selected
    m_GET_MenuKey = m_def_GET_MenuKey
    m_isPrint = m_def_isPrint
End Sub

'저장소에서 속성값을 로드합니다.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_GET_Selected = PropBag.ReadProperty("GET_Selected", m_def_GET_Selected)
    m_GET_MenuKey = PropBag.ReadProperty("GET_MenuKey", m_def_GET_MenuKey)
    m_isPrint = PropBag.ReadProperty("isPrint", m_def_isPrint)
End Sub

'속성값을 저장소에 기록합니다.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("GET_Selected", m_GET_Selected, m_def_GET_Selected)
    Call PropBag.WriteProperty("GET_MenuKey", m_GET_MenuKey, m_def_GET_MenuKey)
    Call PropBag.WriteProperty("isPrint", m_isPrint, m_def_isPrint)
End Sub

'MemberInfo=0
Public Function SET_Item(ByVal tMenuName, ByVal tPrice As Long, ByVal MenuKey As String, ClothImagePath As String) As Boolean
    lblMenuName.Caption = tMenuName
    
    If tPrice > 0 Then
        lblPrice.Caption = Format(tPrice, "#,##0")
        lblPrice.Tag = tPrice
        lblPrice.Visible = True
    End If
    
    If Left(ClothImagePath, 1) = "&" Then
        imgCloth.Picture = LoadPicture()
        
        Select Case ClothImagePath
            Case "&HFFFFFF": imgCloth.BackColor = &HFFFFFF '흰색
            Case "&H51E2E6": imgCloth.BackColor = &H51E2E6 '상아
            Case "&HC0C0C0": imgCloth.BackColor = &HC0C0C0 '회색
            
            Case "&H808080": imgCloth.BackColor = &H808080 '쥐색
            Case "&H4080&":  imgCloth.BackColor = &H4080&  '밤색
            Case "&H0&":     imgCloth.BackColor = &H0&     '검정
            
            Case "&H8000FF": imgCloth.BackColor = &H8000FF '
            Case "&H8080FF": imgCloth.BackColor = &H8080FF '
            Case "&HFF&":    imgCloth.BackColor = &HFF&    '
            Case "&HFFFF&":  imgCloth.BackColor = &HFFFF&  '
            Case "&HA2FDF9": imgCloth.BackColor = &HA2FDF9 '
            Case "&H80&":    imgCloth.BackColor = &H80&    '
            Case "&H40FF00": imgCloth.BackColor = &H40FF00 '
            Case "&H16C212": imgCloth.BackColor = &H16C212 '
            Case "&H408000": imgCloth.BackColor = &H408000 '
            Case "&H808040": imgCloth.BackColor = &H808040 '
            Case "&HFFFF00": imgCloth.BackColor = &HFFFF00 '
            Case "&HFF0000": imgCloth.BackColor = &HFF0000 '
            
            Case "&HA00000": imgCloth.BackColor = &HA00000 '
            Case "&HEB14E6": imgCloth.BackColor = &HEB14E6 '
            Case "&H800080": imgCloth.BackColor = &H800080 '
            Case "&HFFFFFF": imgCloth.BackColor = &HFFFFFF '
        End Select
    Else
        If ClothImagePath = "" Then
            imgCloth.Visible = False
        Else
            imgCloth.Picture = LoadPicture(ClothImagePath)
        End If
    End If
    
    tMenuKey = MenuKey
End Function

'MemberInfo=12
Public Function GetMenuName() As String
    GetMenuName = lblMenuName.Caption
End Function

'MemberInfo=8
Public Function GetPrice() As Long
    GetPrice = lblPrice.Tag
End Function

'MemberInfo=0
Public Function Selected(ByVal tMode As Boolean) As Boolean
    If tMode Then
        lblMenuName.ForeColor = vbRed
    Else
        lblMenuName.ForeColor = vbBlack
    End If
    
    Selected = tMode
End Function

'MemberInfo=0,1,2,0
Public Property Get GET_Selected() As Boolean
    If lblMenuName.ForeColor = vbRed Then
        'GET_Selected = m_GET_Selected
        GET_Selected = True
    Else
        GET_Selected = False
    End If
End Property

Public Property Let GET_Selected(ByVal New_GET_Selected As Boolean)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    
    m_GET_Selected = New_GET_Selected
    PropertyChanged "GET_Selected"
End Property

'MemberInfo=8,1,2,0
Public Property Get GET_MenuKey() As String
    'GET_MenuKey = m_GET_MenuKey
    GET_MenuKey = tMenuKey
End Property

Public Property Let GET_MenuKey(ByVal New_GET_MenuKey As String)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    
    m_GET_MenuKey = New_GET_MenuKey
    PropertyChanged "GET_MenuKey"
End Property
