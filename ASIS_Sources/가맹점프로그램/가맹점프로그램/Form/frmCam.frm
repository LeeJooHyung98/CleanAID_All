VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmCam 
   BorderStyle     =   1  '단일 고정
   Caption         =   "사진찍기"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10065
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCam.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   10065
   StartUpPosition =   2  '화면 가운데
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7755
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   13679
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   1
      SplitterBarAppearance=   1
      BorderStyle     =   0
      PaneTree        =   "frmCam.frx":058A
      Begin Threed.SSPanel SSPanel1 
         Height          =   7230
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   12753
         _Version        =   262144
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XtremeSuiteControls.PushButton btnTool 
            Height          =   420
            Index           =   0
            Left            =   45
            TabIndex        =   11
            Top             =   45
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmCam.frx":05FC
         End
         Begin VB.PictureBox picColorL 
            Appearance      =   0  '평면
            BackColor       =   &H80000008&
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   45
            ScaleHeight     =   285
            ScaleWidth      =   300
            TabIndex        =   8
            Top             =   6165
            Width           =   330
         End
         Begin VB.PictureBox picColorR 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   135
            ScaleHeight     =   285
            ScaleWidth      =   300
            TabIndex        =   9
            Top             =   6345
            Width           =   330
         End
         Begin XtremeSuiteControls.PushButton btnTool 
            Height          =   420
            Index           =   1
            Left            =   45
            TabIndex        =   12
            Top             =   465
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmCam.frx":0C26
         End
         Begin XtremeSuiteControls.PushButton btnTool 
            Height          =   420
            Index           =   2
            Left            =   45
            TabIndex        =   13
            Top             =   885
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmCam.frx":1250
         End
         Begin XtremeSuiteControls.PushButton btnTool 
            Height          =   420
            Index           =   3
            Left            =   45
            TabIndex        =   14
            Top             =   1305
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmCam.frx":187A
         End
         Begin XtremeSuiteControls.PushButton btnTool 
            Height          =   420
            Index           =   4
            Left            =   45
            TabIndex        =   15
            Top             =   1725
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmCam.frx":1EA4
         End
         Begin XtremeSuiteControls.PushButton btnTool 
            Height          =   420
            Index           =   5
            Left            =   45
            TabIndex        =   16
            Top             =   2145
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmCam.frx":24CE
         End
         Begin XtremeSuiteControls.PushButton btnTool 
            Height          =   420
            Index           =   6
            Left            =   45
            TabIndex        =   17
            Top             =   2565
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmCam.frx":2AF8
         End
         Begin XtremeSuiteControls.PushButton btnTool 
            Height          =   420
            Index           =   7
            Left            =   45
            TabIndex        =   18
            Top             =   2985
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmCam.frx":3122
         End
         Begin XtremeSuiteControls.PushButton btnTool 
            Height          =   420
            Index           =   8
            Left            =   45
            TabIndex        =   19
            Top             =   3405
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmCam.frx":374C
         End
         Begin XtremeSuiteControls.PushButton btnTool 
            Height          =   420
            Index           =   9
            Left            =   45
            TabIndex        =   20
            Top             =   3825
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmCam.frx":3D76
         End
         Begin XtremeSuiteControls.PushButton btnTool 
            Height          =   420
            Index           =   10
            Left            =   45
            TabIndex        =   21
            Top             =   4245
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmCam.frx":43A0
         End
         Begin XtremeSuiteControls.PushButton btnTool 
            Height          =   420
            Index           =   11
            Left            =   45
            TabIndex        =   22
            Top             =   4665
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmCam.frx":49CA
         End
         Begin XtremeSuiteControls.PushButton btnTool 
            Height          =   420
            Index           =   12
            Left            =   45
            TabIndex        =   23
            Top             =   5085
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmCam.frx":4FF4
         End
         Begin XtremeSuiteControls.PushButton btnTool 
            Height          =   420
            Index           =   15
            Left            =   45
            TabIndex        =   24
            Top             =   5505
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   741
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmCam.frx":561E
         End
      End
      Begin VB.PictureBox picCapture 
         Appearance      =   0  '평면
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   7230
         Left            =   540
         ScaleHeight     =   7230
         ScaleWidth      =   9525
         TabIndex        =   6
         Top             =   0
         Width           =   9525
      End
      Begin FramePlusCtl.FramePlus FramePlus1 
         Height          =   510
         Left            =   0
         TabIndex        =   1
         Top             =   7245
         Width           =   10065
         _ExtentX        =   17754
         _ExtentY        =   900
         Style           =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin XtremeSuiteControls.PushButton btnCapture 
            Height          =   450
            Left            =   6090
            TabIndex        =   10
            Top             =   30
            Width           =   1395
            _Version        =   851970
            _ExtentX        =   2461
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "사진찍기"
            ForeColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Picture         =   "frmCam.frx":5C58
         End
         Begin VB.ListBox lstDevices 
            Height          =   240
            Left            =   3930
            TabIndex        =   4
            Top             =   135
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.CommandButton cmdStart 
            Caption         =   "캡쳐"
            Height          =   375
            Left            =   1860
            TabIndex        =   3
            Top             =   60
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.CommandButton cmdStop 
            Caption         =   "STOP"
            Height          =   375
            Left            =   2820
            TabIndex        =   2
            Top             =   60
            Visible         =   0   'False
            Width           =   960
         End
         Begin XtremeSuiteControls.PushButton btnSave 
            Height          =   450
            Left            =   8865
            TabIndex        =   25
            Top             =   30
            Width           =   1170
            _Version        =   851970
            _ExtentX        =   2064
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   " 저장"
            ForeColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Picture         =   "frmCam.frx":5D2F
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "#"
            Height          =   180
            Left            =   90
            TabIndex        =   27
            Top             =   285
            Width           =   90
         End
         Begin VB.Label lblTag 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "#"
            Height          =   180
            Left            =   90
            TabIndex        =   26
            Top             =   60
            Width           =   90
         End
         Begin VB.Label lblWebCAM_No 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5400
            TabIndex        =   5
            Top             =   105
            Visible         =   0   'False
            Width           =   120
         End
      End
   End
End
Attribute VB_Name = "frmCam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const WM_CAP As Integer = &H400

Const WM_CAP_DRIVER_CONNECT    As Long = WM_CAP + 10
Const WM_CAP_DRIVER_DISCONNECT As Long = WM_CAP + 11
Const WM_CAP_EDIT_COPY         As Long = WM_CAP + 30

Const WM_CAP_SET_PREVIEW     As Long = WM_CAP + 50
Const WM_CAP_SET_PREVIEWRATE As Long = WM_CAP + 52
Const WM_CAP_SET_SCALE       As Long = WM_CAP + 53

Const WS_CHILD     As Long = &H40000000
Const WS_VISIBLE   As Long = &H10000000
Const SWP_NOMOVE   As Long = &H2
Const SWP_NOSIZE   As Integer = 1
Const SWP_NOZORDER As Integer = &H4
Const HWND_BOTTOM  As Integer = 1

Dim iDevice As Long  ' Current device ID
Dim hHwnd   As Long ' Handle to preview window

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hndw As Long) As Boolean
Private Declare Function capCreateCaptureWindowA Lib "avicap32.dll" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Integer, ByVal hwndParent As Long, ByVal nID As Long) As Long
Private Declare Function capGetDriverDescriptionA Lib "avicap32.dll" (ByVal wDriver As Long, ByVal lpszName As String, ByVal cbName As Long, ByVal lpszVer As String, ByVal cbVer As Long) As Boolean

'Paint
'---------------------------------------------------
Private blnSquareFill As Boolean
Private blnCircleFill As Boolean
Private blnElipseFill As Boolean

Private mudtTool  As UDT_Tool
Private mudtPoint As UDT_Point

Private lngLColor As Long
Private lngRColor As Long

Private Sub cmdFormat_Click()

End Sub

Private Sub cmdSource_Click()

End Sub

Private Sub btnCapture_Click()
    On Error Resume Next
    
    Dim DPath As String
    Dim bm    As Image
    
    '
    DPath = AppPath + "Capture"
    
    If Dir(DPath, vbDirectory) = "" Then
        MkDir DPath
    End If
    
    ' Copy image to clipboard
    SendMessage hHwnd, WM_CAP_EDIT_COPY, 0, 0
    ClosePreviewWindow

    picCapture.Picture = Clipboard.GetData
End Sub

Private Sub btnSave_Click()
    On Error Resume Next
    
    Dim DPath As String
    Dim bm    As Image
    
    '
    DPath = AppPath + "Capture"
    
    If Dir(DPath, vbDirectory) = "" Then
        MkDir DPath
    End If
    
''    ' Copy image to clipboard
''    SendMessage hHwnd, WM_CAP_EDIT_COPY, 0, 0
''    ClosePreviewWindow
''
''    picCapture.Picture = Clipboard.GetData
    
    SavePicture picCapture.Image, AppPath & "Capture\" & Format(lblDate.Caption, "YYYYMMDD") & lblTag.Caption & "-" & lblWebCAM_No.Caption & ".JPG"
    
    With frm접수
        .picCAM(0).Picture = LoadPicture(App.Path & "\Capture\" & Format(lblDate.Caption, "YYYYMMDD") & lblTag.Caption & "-0.jpg")
        .picCAM(1).Picture = LoadPicture(App.Path & "\Capture\" & Format(lblDate.Caption, "YYYYMMDD") & lblTag.Caption & "-1.jpg")
        .picCAM(2).Picture = LoadPicture(App.Path & "\Capture\" & Format(lblDate.Caption, "YYYYMMDD") & lblTag.Caption & "-2.jpg")
        .picCAM(3).Picture = LoadPicture(App.Path & "\Capture\" & Format(lblDate.Caption, "YYYYMMDD") & lblTag.Caption & "-3.jpg")
        
        .sprGrid.Row = .sprGrid.ActiveRow
        .sprGrid.Col = 10: .sprGrid.Text = "1"
    End With
    
    Unload Me
End Sub

Private Sub btnTool_Click(Index As Integer)
    'set UDT tool to clear previous tool
    SetAllFalse mudtTool
    
    'set the width in case it chaged with the previous tool
    picCapture.DrawWidth = CurentWidth
    
    'select tool and attributes
    With mudtTool
        Select Case Index
            Case 0: .Point = True     'single point
            Case 1: .Line = True      'straight line
            Case 2: .FreeLine = True  'freehand line
            Case 3      'circle
                picCapture.FillStyle = vbFSTransparent
                blnCircleFill = False
                .Circle = True
            Case 4      'filled Circle
                picCapture.FillStyle = intFillStyle
                blnCircleFill = True
                .Circle = True
            Case 5      'elipse
                picCapture.FillStyle = vbFSTransparent
                blnElipseFill = False
                .Elipse = True
            Case 6      'elipse fill
                picCapture.FillStyle = intFillStyle
                blnElipseFill = True
                .Elipse = True
            Case 7      'empty rectangle
                blnSquareFill = False
                picCapture.FillStyle = vbFSTransparent
                .Square = True
            Case 8      'filled rectangle
                blnSquareFill = True
                picCapture.FillStyle = intFillStyle
                picCapture.FillColor = lngRColor
                .Square = True
            Case 9: .Fan = True     'fan effect
            Case 10: .ColorPicker = True    'colorpicker tool
            Case 11: .Caligraphy = True     'caligraphy tool
            Case 12
                .Filler = True         'filler tool
                picCapture.FillStyle = intFillStyle
            Case 15     'eraser
                picCapture.DrawWidth = CurentWidth + 5
                .Eraser = True
            Case 20: picCapture.BackColor = lngRColor    'fill background with current selected color
            Case 21: RandomizeBackground Me    'fill background with random color dots
        End Select
    End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    picColorL.BackColor = vbRed

    CurentWidth = 5
    
    picCapture.DrawWidth = CurentWidth
    
    lngLColor = vbRed
    lngRColor = vbWhite
    
    picCapture.FillColor = vbRed
    picCapture.FillStyle = vbSolid

    'set UDT tool to clear previous tool
    SetAllFalse mudtTool
        
    'select tool and attributes
    mudtTool.FreeLine = True  'freehand line

    '--------------------------------------------------
    Call LoadDeviceList
    
    If lstDevices.ListCount > 0 Then
        lstDevices.Selected(0) = True
    Else
        cmdStart.Enabled = False
        lstDevices.AddItem ("No Device Available")
    End If
    
    Call cmdStart_Click
End Sub

Private Sub ButtonPlus1_Click(Index As Integer)

End Sub

Private Sub cmdStart_Click()
    iDevice = lstDevices.ListIndex
    OpenPreviewWindow
End Sub

Private Sub cmdStop_Click()
    ClosePreviewWindow
    
    picCapture.Cls
    
    cmdStop.Enabled = False
    btnSave.Enabled = False
    cmdStart.Enabled = True
End Sub

Private Sub cmdTool_Click(Index As Integer)
    'set UDT tool to clear previous tool
    SetAllFalse mudtTool
    
    'set the width in case it chaged with the previous tool
    picCapture.DrawWidth = CurentWidth
    
    'select tool and attributes
    With mudtTool
        Select Case Index
            Case 0: .Point = True     'single point
            Case 1: .Line = True      'straight line
            Case 2: .FreeLine = True  'freehand line
            Case 3      'circle
                picCapture.FillStyle = vbFSTransparent
                blnCircleFill = False
                .Circle = True
            Case 4      'filled Circle
                picCapture.FillStyle = intFillStyle
                blnCircleFill = True
                .Circle = True
            Case 5      'elipse
                picCapture.FillStyle = vbFSTransparent
                blnElipseFill = False
                .Elipse = True
            Case 6      'elipse fill
                picCapture.FillStyle = intFillStyle
                blnElipseFill = True
                .Elipse = True
            Case 7      'empty rectangle
                blnSquareFill = False
                picCapture.FillStyle = vbFSTransparent
                .Square = True
            Case 8      'filled rectangle
                blnSquareFill = True
                picCapture.FillStyle = intFillStyle
                picCapture.FillColor = lngRColor
                .Square = True
            Case 9: .Fan = True     'fan effect
            Case 10: .ColorPicker = True    'colorpicker tool
            Case 11: .Caligraphy = True     'caligraphy tool
            Case 12
                .Filler = True         'filler tool
                picCapture.FillStyle = intFillStyle
            Case 15     'eraser
                picCapture.DrawWidth = CurentWidth + 5
                .Eraser = True
            Case 20: picCapture.BackColor = lngRColor    'fill background with current selected color
            Case 21: RandomizeBackground Me    'fill background with random color dots
        End Select
    End With
End Sub

Private Sub LoadDeviceList()
    Dim strName As String
    Dim strVer  As String
    Dim iReturn As Boolean
    Dim x       As Long
    
    x = 0
    strName = Space(100)
    strVer = Space(100)

    ' Load name of all available devices into lstDevices
    Do
        iReturn = capGetDriverDescriptionA(x, strName, 100, strVer, 100) ' Get Driver name and version
        
        If iReturn Then
            lstDevices.AddItem Trim$(strName)  ' If there was a device add device name to the list
        End If
        
        x = x + 1
    Loop Until iReturn = False
End Sub

Private Sub OpenPreviewWindow()
    ' Open Preview window in picturebox
    hHwnd = capCreateCaptureWindowA(iDevice, WS_VISIBLE Or WS_CHILD, 0, 0, 640, 480, picCapture.hwnd, 0)

    ' Connect to device
    If SendMessage(hHwnd, WM_CAP_DRIVER_CONNECT, iDevice, 0) Then
        SendMessage hHwnd, WM_CAP_SET_SCALE, True, 0 'Set the preview scale
        SendMessage hHwnd, WM_CAP_SET_PREVIEWRATE, 66, 0 'Set the preview rate in milliseconds
        SendMessage hHwnd, WM_CAP_SET_PREVIEW, True, 0 'Start previewing the image from the camera

        ' Resize window to fit in picturebox
        'SetWindowPos hHwnd, HWND_BOTTOM, 0, 0, picCapture.ScaleWidth, picCapture.ScaleHeight, SWP_NOMOVE Or SWP_NOZORDER
        btnSave.Enabled = True
        cmdStop.Enabled = True
        cmdStart.Enabled = False
    Else
        DestroyWindow hHwnd ' Error connecting to device close window

        btnSave.Enabled = False
    End If
 End Sub

Private Sub ClosePreviewWindow()
    SendMessage hHwnd, WM_CAP_DRIVER_DISCONNECT, iDevice, 0 ' Disconnect from device
    DestroyWindow hHwnd ' close window
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdStop.Enabled Then
        Call ClosePreviewWindow
    End If
End Sub

'picCapture
Private Sub picCapture_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'check to see which tool is selected and act accordingly
   If Button = vbLeftButton Then
        lngOutColor = lngLColor
   ElseIf Button = vbRightButton Then
        lngOutColor = lngRColor
   End If
   
   With mudtPoint
        .msglX = x
        .msglY = y
   End With
   
   'start to draw the objects
   With mudtTool
        If .Line Then
            InitiateLine mudtPoint, Me
        ElseIf .Point Then
            blnModified = True
            picCapture.PSet (mudtPoint.msglX, mudtPoint.msglY), lngOutColor
        ElseIf .Circle Then
            InitiateCircle mudtPoint, blnCircleFill, Me
        ElseIf .Elipse Then
            InitiateElipse mudtPoint, blnElipseFill, Me
        ElseIf .Square Then
            InitiateSquare mudtPoint, blnSquareFill, Me
        ElseIf .Eraser Then
            picCapture.PSet (x, y), picCapture.BackColor
        ElseIf .FreeLine Then
            InitiateFreeLine mudtPoint, Me
        ElseIf .Fan Then
            InitiateFan mudtPoint, Me
        ElseIf .ColorPicker Then
            GetColorPicker mudtPoint, Button, Me
        ElseIf .Caligraphy Then
            InitiateCaligraphy mudtPoint, Me
        ElseIf .Filler Then
            picCapture.FillColor = lngOutColor
            DoFiller mudtPoint, picCapture.hdc, Me
        End If
   End With
End Sub

Private Sub picCapture_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Or Button = vbRightButton Then
       With mudtPoint
        .msglX = x
        .msglY = y
       End With
       
        'use live drawing features here
        With mudtTool
            If .FreeLine Then
                DrawFreeLine mudtPoint, Me
            ElseIf .Fan Then
                DrawFan mudtPoint, Me
            ElseIf .Eraser Then
                picCapture.PSet (x, y), picCapture.BackColor
            ElseIf .Line Then
                DrawLine mudtPoint, Me
            ElseIf .Circle Then
                DrawCircle mudtPoint, Me
            ElseIf .Elipse Then
                DrawElipse mudtPoint, Me
            ElseIf .Square Then
                DrawSquare mudtPoint, Me
            ElseIf .Caligraphy Then
                DrawCaligraphy mudtPoint, Me
            End If
        End With
    End If
End Sub

Private Sub picCapture_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    picCapture.DrawMode = 13
    
   With mudtPoint
        .msglX = x
        .msglY = y
   End With
   
    'draw the final objects
    With mudtTool
        If .Line Then
            FinalizeLine mudtPoint, Me
        ElseIf .Circle Then
            FinalizeCircle mudtPoint, Me
        ElseIf .Elipse Then
            FinalizeElipse mudtPoint, Me
        ElseIf .Square Then
            FinalizeSquare mudtPoint, Me
        End If
    End With
End Sub
