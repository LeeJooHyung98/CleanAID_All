VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#13.2#0"; "Codejock.SkinFramework.v13.2.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#13.2#0"; "Codejock.CommandBars.v13.2.1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm P_00000 
   BackColor       =   &H00FFFFFF&
   Caption         =   "(주)크린에이드"
   ClientHeight    =   10710
   ClientLeft      =   1575
   ClientTop       =   3150
   ClientWidth     =   15240
   Icon            =   "P_00000.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  '최대화
   Begin MSWinsockLib.Winsock tcpWinsock 
      Left            =   5070
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdgExcel 
      Left            =   3000
      Top             =   1230
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stbMsg 
      Align           =   2  '아래 맞춤
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   10365
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "설   명"
            TextSave        =   "설   명"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15319
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1940
            MinWidth        =   1940
            TextSave        =   "2023-12-13"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "오후 2:56"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework 
      Left            =   495
      Top             =   3555
      _Version        =   851970
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   210
      Top             =   765
      _Version        =   851970
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      DesignerControls=   -1  'True
      DesignerControlsData=   "P_00000.frx":0A02
   End
End
Attribute VB_Name = "P_00000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long

Dim StatusBar As XtremeCommandBars.IStatusBar

'StatusBar - Place these constants in the General code section
Const ID_INDICATOR_CAPS = 59137
Const ID_INDICATOR_NUM = 59138
Const ID_INDICATOR_SCRL = 59139

Const ID_INDICATOR_DATE = 500

Const ID_VERSION = 100
Const ID_NAME = 200
Const ID_TEL = 300


Dim RS01 As ADODB.Recordset
Dim sValue() As String

Dim Err_Num As Long
Dim Err_Dec As String


Private Sub GetMenuInfo(hMenu As Long, spaces As Integer, Txt As String)
    Dim Num As Integer
    Dim i As Integer
    Dim Length As Long
    Dim sub_hmenu As Long
    Dim sub_name As String

    Num = GetMenuItemCount(hMenu)
    
    For i = 0 To Num - 1
        ' Save this menu's info.
        
        sub_hmenu = GetSubMenu(hMenu, i)
        sub_name = Space$(256)
        Length = GetMenuString(hMenu, i, sub_name, Len(sub_name), MF_BYPOSITION)
        sub_name = Left$(sub_name, Length)

        Txt = Txt & Space$(spaces) & sub_name & vbCrLf
        
        ' Get its child menu's names.
        GetMenuInfo sub_hmenu, spaces + 4, Txt
    Next i
End Sub

Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error Resume Next
    
    Select Case Control.ID
        
        '---------------------------------------------
        Case 1205: P_01090.SetFocus         '"m_01090"  'PDA 사용자 등록
        Case 1206: P_07010.SetFocus         '"m_07010"  '품목 등록
        Case 1207: P_07011.SetFocus         '"m_07011"  '지사 등록
        Case 1208: P_07012.SetFocus         '"m_07012"  '외주 입고 등록
        Case 1209: P_07013.SetFocus         '"m_07013"  '외주 출고 등록
        Case 1210: P_07014.SetFocus         '"m_07014"  '외주 입고 현황
        Case 1211: P_07015.SetFocus         '"m_07015"  '외주 출고 현황
        Case 1222: P_07016.SetFocus         '"m_07016"  '미입고 처리 현황
        Case 1223: P_07017.SetFocus         '"m_07017"  '미출고 처리 현황
        Case 1224: P_07018.SetFocus         '"m_07017"  '미출고 현황
        
        Case 1055:
                Rtn = MsgBox("종료하시겠습니까?", vbQuestion + vbYesNo + vbDefaultButton1, "확인")
                
                If Rtn = vbYes Then
                    Unload P_00000
                    End
                End If
                
        Case 57650: Me.Arrange vbCascade        'ID_WINDOW_CASCADE
        Case 57651: Me.Arrange vbTileHorizontal 'ID_WINDOW_TILE_HORIZANTALLY
        Case 57649: Me.Arrange vbTileVertical   'ID_WINDOW_TILE_VERTICALLY
    End Select
End Sub

Private Sub MDIForm_Initialize()
   InitCommonControls
End Sub

Private Sub MDIForm_Load()
    '-------------------------------------------------------------------------
    '
    '-------------------------------------------------------------------------
    CommandBars.LoadDesignerBars

    '2008-01-10
    'SkinFramework.LoadSkin App.Path + "\Styles\Office2007.cjstyles", ""
    'SkinFramework.LoadSkin App.Path + "\Styles\WinXP.Royale.cjstyles", ""
    SkinFramework.LoadSkin App.Path + "\Styles\Vista.cjstyles", ""
    SkinFramework.ApplyWindow Me.hwnd
    SkinFramework.ApplyOptions = SkinFramework.ApplyOptions Or xtpSkinApplyMetrics


    'Dim StatusBar As XtremeCommandBars.IStatusBar
    Set StatusBar = CommandBars.StatusBar

    StatusBar.Visible = True               'Make the custom status bar visible

    StatusBar.AddPane 0                      'Adds the "special" idle pane to the custom status bar
    StatusBar.SetPaneStyle 0, SBPS_STRETCH 'Set Pane Style
    StatusBar.IdleText = "작업준비"          'Add some Idle Text to be displayed while the application is idle
    StatusBar.SetPaneWidth 0, 75             'Set Pane width

    StatusBar.AddPane ID_VERSION
    StatusBar.SetPaneText ID_VERSION, ""
    StatusBar.SetPaneWidth ID_VERSION, 350

    StatusBar.AddPane ID_NAME
    StatusBar.SetPaneText ID_NAME, ""
    StatusBar.SetPaneWidth ID_NAME, 150

    StatusBar.AddPane ID_TEL
    StatusBar.SetPaneText ID_TEL, "전화 : 000-0000"
    StatusBar.SetPaneWidth ID_TEL, 150

    StatusBar.AddPane ID_INDICATOR_DATE
    StatusBar.SetPaneText ID_INDICATOR_DATE, Format(Date, "YYYY-MM-DD")
    StatusBar.SetPaneWidth ID_INDICATOR_DATE, 80

    StatusBar.AddPane ID_INDICATOR_CAPS 'Adds the special Caps lock indicator pane
    StatusBar.AddPane ID_INDICATOR_NUM  'Adds the special Num lock indicator pane
    StatusBar.AddPane ID_INDICATOR_SCRL 'Adds the special Scroll lock indicator pane
    '-------------------------------------------------------------------------
     

    stbMsg.Panels(3).Text = USERNAME
    stbMsg.Panels(4).Text = "Version " + strProgram_Version
    stbMsg.Panels(4).ToolTipText = "LastEdit " + strProgram_LastEdit
    
    
    Me.Caption = Store.Code & "-" & Store.Name
End Sub
 
