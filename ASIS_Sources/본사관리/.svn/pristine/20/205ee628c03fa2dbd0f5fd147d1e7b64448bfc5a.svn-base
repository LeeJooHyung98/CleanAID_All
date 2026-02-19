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
   ClientLeft      =   5775
   ClientTop       =   2835
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
            TextSave        =   "2019-01-23"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "오전 11:12"
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
Dim WithEvents Workspace        As TabWorkspace
Attribute Workspace.VB_VarHelpID = -1

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
    
    Debug.Print "CommandBars_Execute Control.Id : " & Control.Id
    
    Select Case Control.Id
        Case 1094: P_01001.SetFocus    '"m_01001"
        'Case 1095: P_01001_1.SetFocus '"m_01002_1"
        Case 1218: P_01002.SetFocus    '"m_01002"
        Case 1097: P_01003.SetFocus    '"m_01003"
        Case 1098: P_01011.SetFocus    '"m_01011"
        Case 1099: P_01004.SetFocus    '"m_01004"
        Case 1100: P_01005_A.SetFocus  '"m_01005_A"
        
        Case 1101:
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_01001_M.SetFocus  '"m_01001_M"
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
            
        Case 1102:
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_01001.SetFocus    '"m_01001_A"  P_01001_A.SetFocus
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
            
        Case 1103:
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_01003_A.SetFocus  '"m_01003_A"
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
        
        Case 1104:
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_01011_A.SetFocus  '"m_01003_A"
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
            
        Case 1212:
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_01011_B.SetFocus  '"m_01003_A"
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
        
        Case 1106:
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_01004_A.SetFocus  '"m_01004_A"
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
            
        Case 1227:
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_01004_C.SetFocus  '"m_01004_C"
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
        
        Case 1202:
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_01005_B.SetFocus  '"m_01005_B"
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
        
        Case 1228:
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_01005_C.SetFocus  '"m_01004_C"
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
            
        Case 1159:
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_04022.SetFocus  '"m_01004_C"
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
            
        Case 1105:
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_01012.SetFocus  '"m_01012"
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
            
        
        Case 1107: P_01008.SetFocus   '"m_01008"
        Case 1108: P_01009.SetFocus   '"m_01009"
        Case 1109: P_01010.SetFocus   '"m_01010"
        
        '---------------------------------------------
        
        Case 1110: P_02002.SetFocus   '"m_02002"
        Case 1111: P_02001.SetFocus   '"m_02001"
        
        Case 1219: P_02017.SetFocus   '"m_02017" 오전사진보기
        
        Case 1112: P_02014.SetFocus   '"m_02014"
        Case 1113: P_02015.SetFocus   '"m_02015"
        Case 1114: P_02004.SetFocus   '"m_02004"
        Case 1115: P_02005.SetFocus   '"m_02005"
        Case 1226: P_02005_01.SetFocus   '"m_02005"
        Case 1254: P_02005_02.SetFocus   '"m_02005"
        Case 1116: P_02016.SetFocus   '"m_02016"
        Case 1117: P_02007.SetFocus   '"m_02007"
        Case 1118: P_02008.SetFocus   '"m_02008"
        Case 1119: P_02009.SetFocus   '"m_02009"
        Case 1120: P_02010.SetFocus   '"m_02010"
        Case 1121: P_02011.SetFocus   '"m_02011"
        Case 1122: P_02012.SetFocus   '"m_02012"
        
        '---------------------------------------------
        
        Case 1123: P_03001.SetFocus   '"m_03001"
        Case 1124: P_03002.SetFocus   '"m_03002"
        Case 1125: P_03003.SetFocus   '"m_03003"
        Case 1126: P_03005.SetFocus   '"m_03005"
        Case 1127: P_03006.SetFocus   '"m_03006"
        Case 1128: P_03007.SetFocus   '"m_03007"
        Case 1129: P_03008.SetFocus   '"m_03008"
        Case 1130: P_03009.SetFocus   '"m_03009"
        Case 1131: P_03010.SetFocus   '"m_03010"
        Case 1132: P_03011.SetFocus   '"m_03011"
        Case 1133: P_03012.SetFocus   '"m_03012"
        Case 1134: P_03013.SetFocus   '"m_03013"
        Case 1135: P_03014.SetFocus   '"m_03014"
        Case 1203: P_03015.SetFocus   '"m_03015"
        Case 1236: P_03016.SetFocus   '"m_03015"
        Case 1249: P_03017.SetFocus   '"m_03017"
        Case 1260: P_03018.SetFocus   '"m_03018"
        '---------------------------------------------
        
        Case 1136: P_04001_Master.SetFocus  '"m_04001_1"
        Case 1137: P_04001.SetFocus         '"m_04001"
        Case 1138: P_04002.SetFocus         '"m_04002"
        Case 1139: P_04003.SetFocus         '"m_04003"
        Case 1140: P_04004.SetFocus         '"m_04004"
        Case 1141: P_04005.SetFocus         '"m_04005"
        Case 1142: P_04006.SetFocus         '"m_04006"
        Case 1143: P_04007.SetFocus         '"m_04007"
        
        Case 1144: P_04001_A.SetFocus       '"m_04001_A"
        Case 1222:
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_04001_B.SetFocus  '"m_01005_B"
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If

        Case 1145: P_04011_A.SetFocus       '"m_04011_A"
        Case 1146: P_04011_B.SetFocus       '"m_04011_B"
        Case 1147: P_04009_A.SetFocus       '"m_04009_A"
        Case 1148: P_04009_M.SetFocus       '"m_04009_M"
        
        
        
        Case 1250: P_04009_N.SetFocus              '"m_04009_N"
'
'            If HeadOffice = MASTER_OFFICE_CODE Then
'                P_04009_N.SetFocus               '"m_04009_N"
'            Else
'                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
'            End If
        
        Case 1149: P_04019.SetFocus         '"m_04019"
        Case 1150: P_04009.SetFocus         '"m_04009"
        Case 1151: P_04010.SetFocus         '"m_04010"
        Case 1152: P_04011.SetFocus         '"m_04011"
        Case 1153: P_04012.SetFocus         '"m_04012"
        Case 1154: P_04013.SetFocus         '"m_04013"
        Case 1155: P_04014.SetFocus         '"m_04014"
        Case 1156: P_04009_R.SetFocus       '"m_04009_R"
        Case 1157: P_04009_R1.SetFocus      '"m_04009_R1"
        
        Case 1158
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_04001_C.SetFocus  '"  "
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
        
        ' 매장 매출 현황 보고용-가맹점기준
        Case 1244
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_04027.SetFocus  '"  "
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
        
        ' 매장 매출 현황 보고용-지사기준
        Case 1256
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_04028.SetFocus  '"  "
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
        
        ' 매장 매출 현황 보고용
        Case 1251
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_04034.SetFocus  '"  "
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
        
        Case 1160: P_04016.SetFocus         '"m_04016"
        
        'Case 1161: P_04017.SetFocus         '"m_04017"
        'Case 1162: P_04018.SetFocus         '"m_04018"
        
        Case 1233: P_04020.SetFocus         '"m_04020"
        Case 1232: P_04020_1.SetFocus         '"m_04020_1"
        
        Case 1217: P_04025.SetFocus  '"m_01005_B"
        Case 1243: P_04030.SetFocus
        Case 1252: P_04035.SetFocus
        Case 1258: P_04036.SetFocus  ' 일일판매집계 (가맹점)
        Case 1259: P_04037.SetFocus  ' 매출현황 (가맹점)
        
        
        
        ' 오픈매장 예상 매출 등록
        Case 1246
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_04031.SetFocus  '"  "
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
        ' 오픈 매장 매출 관리
        Case 1247
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_04032.SetFocus  '"  "
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
        ' 오픈매장 달성율 순위 조회
        Case 1248
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_04033.SetFocus  '"  "
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
        
        
        
        
        '---------------------------------------------
        
        Case 1163: P_05001.SetFocus         '"m_05001"
        Case 1164: P_05002.SetFocus         '"m_05002"
        Case 1165: P_05004.SetFocus         '"m_05004"
        Case 1166: P_05006.SetFocus         '"m_05006"
        Case 1167: P_05007.SetFocus         '"m_05007"
        'Case 1168: P_05009.SetFocus         '"m_05009"
        Case 1169: P_05010.SetFocus         '"m_05010"
        Case 1170: P_05011.SetFocus         '"m_05011"
        
        Case 1171: P_05013.SetFocus         '"m_05013"
        Case 1172: P_05014.SetFocus         '"m_05014"
        Case 1240: P_05015.SetFocus         '"m_05015"
        Case 1255: P_05016.SetFocus         '"m_05015"
        ' 장부관리조회
        Case 1264
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_05017.SetFocus  '"  "
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
        
        Case 1265
            ' 물세탁 일지
            P_05018.SetFocus
        '---------------------------------------------
        
        Case 1173: P_06001.SetFocus         '"m_06001"
        Case 1174: P_06002.SetFocus         '"m_06002"
        Case 1175: P_06003.SetFocus         '"m_06003"
        Case 1176: P_06004.SetFocus         '"m_06004"
        Case 1177: P_06005.SetFocus         '"m_06005"
        Case 1178: P_06006.SetFocus         '"m_06006"
        Case 1179: P_06007.SetFocus         '"m_06007"
        Case 1257: P_06011.SetFocus         '"m_06011"
        
        Case 1225
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_06010.SetFocus  '"  "
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
        
        Case 1261
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_06012.SetFocus  '"  "
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
        
        Case 1262
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_06013.SetFocus  '"  "
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
        Case 1263
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_06014.SetFocus  '"  "
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
        
        
        
        
        '---------------------------------------------
        Case 1180: P_07001.SetFocus         '"m_07001"
        Case 1181: P_07002.SetFocus         '"m_07002"
        Case 1204: P_07003.SetFocus         '"m_07003"
        Case 1182: P_07004.SetFocus         '"m_07004"
        Case 1183: P_07005.SetFocus         '"m_07005"
        
        Case 1205: P_01090.SetFocus         '"m_01090"
'        Case 1206: P_07010.SetFocus         '"m_07010"
'        Case 1207: P_07011.SetFocus         '"m_07011"
'        Case 1208: P_07012.SetFocus         '"m_07012"
'        Case 1209: P_07013.SetFocus         '"m_07013"
'        Case 1210: P_07014.SetFocus         '"m_07014"
'        Case 1211: P_07015.SetFocus         '"m_07015"
        
        '---------------------------------------------
        
        Case 1189: P_09004.SetFocus         '"m_09004"
        Case 1190: P_09005.SetFocus         '"m_09005"
        Case 1200: P_09006.SetFocus         '"m_09006"
        
        '---------------------------------------------
        
        Case 1194: P_10001.SetFocus         '"m_10001"
        Case 1191: P_10002.SetFocus         '"m_10002"
        Case 1195: P_10003.SetFocus         '"m_10003"
        Case 1192: P_10004.SetFocus         '"m_10004"
'        Case 1200: P_10003.SetFocus         '"m_10003"
        
        '---------------------------------------------
        
        Case 1198: P_SMSALL_1.SetFocus      '"smsall_001"
        Case 1199: P_SMSALL_2.SetFocus      '"smsall_002"
        Case 1201: P_SMSALL_3.SetFocus      '"smsall_003"
        
        Case 1239: P_SMSALL_7.SetFocus      '"smsall_003"
        
        '---------------------------------------------
        ' 3 큐브
        Case 1235
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_SMSALL_4.SetFocus  '"  "
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
        
        '---------------------------------------------
        ' 마트 인력 협력인 리스트 등록
        Case 1237
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_SMSALL_5.SetFocus  '"  "
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
        
        '---------------------------------------------
        ' 마트 협력인 SMS 전송
        Case 1238
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_SMSALL_6.SetFocus  '"  "
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
        
        '---------------------------------------------
        ' SMS 기간별 등록 현황
        Case 1253
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_SMSALL_9.SetFocus  '"  "
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
        
        '---------------------------------------------
        ' 매출 분석 #1
        Case 1229
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_04023.SetFocus  '"  "
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
        
        ' 특정매장 매출분석
        Case 1234
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_04026.SetFocus  '"  "
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
        
        
        '---------------------------------------------
        ' 월간매출현황(일별합계)
        Case 1230: P_04024.SetFocus
        
        ' 특정매장 매출분석
        Case 1267
            If HeadOffice = MASTER_OFFICE_CODE Then
                P_SMSALL_10.SetFocus  '"  "
            Else
                MsgBox "사용권한이 없습니다.", vbInformation, "확인"
            End If
        '---------------------------------------------
        
        Case 1055:
            If ActiveForm.Name = "P_00000" Then
                Rtn = MsgBox("정말로 종료하시겠습니까?", vbQuestion + vbYesNo + vbDefaultButton1, "확인")
                
                If Rtn = vbYes Then
                    End
                End If
            
            Else
                Unload ActiveForm
            End If
                
        Case 57650: Me.Arrange vbCascade        'ID_WINDOW_CASCADE
        Case 57651: Me.Arrange vbTileHorizontal 'ID_WINDOW_TILE_HORIZANTALLY
        Case 57649: Me.Arrange vbTileVertical   'ID_WINDOW_TILE_VERTICALLY
        
        
        Case 1223
                If Dir(App.Path & "\AidSupport.exe", vbNormal) <> "" Then
                    Shell App.Path & "\AidSupport.exe", vbNormalFocus
                
                Else
'                    frmIE.Caption = "원격 A/S"
'                    frmIE.WebBrowser1.Navigate "http://as82.kr/cleanaid"
'                    frmIE.Show 1
                End If
                
        Case 1266   '반품요청 현황
            P_03019.SetFocus
        Case Else
            Debug.Print "미등록 메뉴 : &  " & CStr(Control.Id)
        
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

    ' 단축기를 비활성화 한다.
    CommandBars.KeyBindings.Enabled = False
    
    '2008-01-10
    'SkinFramework.LoadSkin App.Path + "\Styles\Office2007.cjstyles", ""
    'SkinFramework.LoadSkin App.Path + "\Styles\WinXP.Royale.cjstyles", ""
    SkinFramework.LoadSkin App.Path + "\Styles\Vista.cjstyles", ""
    SkinFramework.ApplyWindow Me.hwnd
    SkinFramework.ApplyOptions = SkinFramework.ApplyOptions Or xtpSkinApplyMetrics

    '-------------------------------------------------------------------------
    Set Workspace = CommandBars.ShowTabWorkspace(True) ' 화면 탭을 나타낸다
    Workspace.PaintManager.ShowIcons = False
    Workspace.PaintManager.OneNoteColors = False


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
    
    With CCAid
        .SetStoreOffice = Store.Office  '"CLEANAID"
        .SetStoreCode = Store.Code
        .SetStoreName = Store.Name
        .SetIPAddress = Trim(GetIniStr("Store Server", "ServerNameOrIP", "", m_iniFile))
        .SetPort = Val(Trim(GetIniStr("Store Server", "MessagePort", "", m_iniFile)))
        .WinsockControl = tcpWinsock
    End With

    stbMsg.Panels(3).Text = USERNAME
    stbMsg.Panels(4).Text = "Version " + strProgram_Version
    stbMsg.Panels(4).ToolTipText = "LastEdit " + strProgram_LastEdit
    
    

End Sub

Private Sub stbMsg_PanelClick(ByVal Panel As MSComctlLib.Panel)
    
'    If HeadOffice <> MASTER_OFFICE_CODE Then
'        Dim nCnt    As Long
'        Dim Controls As CommandBarControls
'
'        Set Controls = P_00000.CommandBars.DesignerControls
'
'        For nCnt = 1 To Controls.Count
'            Debug.Print CStr(Controls(nCnt).Id) & " - > " & CStr(Controls(nCnt).Caption)
'
'            If InStr("1101,1102,1103,1106,1202", CStr(Controls(nCnt).Id)) > 0 Then
'                Controls(nCnt).Enabled = False
'                Controls(nCnt).Visible = False
'
'            End If
'        Next nCnt
'    End If
'
'    P_00000.CommandBars.DesignerControls.Item(30).Enabled = False
'    P_00000.CommandBars.DesignerControls.Item(30).Visible = False
'
''        Dim Popup As CommandBar
''        Set Popup = CommandBars.ContextMenus.Find(400)

End Sub

'Private Sub stbMsg_PanelClick(ByVal Panel As MSComctlLib.Panel)
'    If Panel.Index = 1 Then
'        P_01001_A1.Show vbModal
'    End If
'End Sub

Private Sub tcpWinsock_Close()
    Call PanelsMsg("연결이 종료 되었습니다.....")
End Sub


Private Sub tcpWinsock_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call PanelsMsg(Description)
End Sub


Private Sub tcpWinsock_DataArrival(ByVal bytesTotal As Long)
    
    ' 수신데이타를 텍스트상자에 출력
    Dim work As String
    tcpWinsock.GetData work, vbString
    
    Call DataArrival_Winsock(work)
    
End Sub


Private Sub DataArrival_Winsock(work As String)
' 전달 메시지

'   S_STA       : 메시지의 시작을 의미한다.
'   CLEANAID    : 프로그램을 사용하고 있는 회사를 의미한다.
'   1001        : 프로그램을 사용하고 있는 회사중 각각의 코드를 의미한다.(지사등등)
'   FILELISTALL : 해당 프로그래에서 실행할 명령어가 전달된다.
'   자료        : 요청 자료가 전달된다 (FILELISTALL의 요청한 자료이다.
'   S_END       : 메시지의 종료를 의미한다.


' EX)  S_STA | CLEANAID | 1001 | FILELISTALL| 쟈료 | S_END
' 설명 크린에이드 1001(본사) 프로그램에서 서버에 모든 파일리스트를 요청했다.

    On Error GoTo ERR_RTN

    Dim varValue    As Variant

    If Fnc_TcpCheckDataArrival(work) = False Then Exit Sub
    
    varValue = Split(work, "|")
    If UBound(varValue) <> 5 Then Exit Sub
    
    Select Case UCase(varValue(3))
        
        ' 서버에서 수신파일의 모든 리스트를 받았을 경우
        Case "RECEIVE_FILELIST_ALL"
            If Fnc_FromEnableCheck("P_08002") = True Then
                P_08002.Display_File_List (CStr(varValue(4)))
                Exit Sub
            End If
            
        Case "RECEIVE_FILENAME_ACTION"
            If Fnc_FromEnableCheck("P_08002") = True Then
                P_08002.Display_File_Action (CStr(varValue(4)))
                Exit Sub
            End If
            
        ' 출고 파일을 정상적으로 만들었을 경우
        Case "CREATE_CHULGO_DATA_OK"
            MsgBox CStr(varValue(4)), vbOKOnly, "확인"
            Exit Sub
            
        ' 출고 파일 생성중 오류가 발생하였을 경우
        Case "CREATE_CHULGO_DATA_ERROR"
            MsgBox CStr(varValue(4)), vbOKOnly, "확인"
            Exit Sub
            
            
            
            
        Case Else
    
    End Select
    Exit Sub
    
ERR_RTN:
    PanelsMsg Err.Description

End Sub


