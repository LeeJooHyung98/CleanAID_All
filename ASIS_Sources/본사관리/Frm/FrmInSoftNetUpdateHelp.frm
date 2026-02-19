VERSION 5.00
Begin VB.Form FrmInSoftNetUpdateHelp 
   Caption         =   "Help"
   ClientHeight    =   7875
   ClientLeft      =   5175
   ClientTop       =   1455
   ClientWidth     =   6825
   Icon            =   "FrmInSoftNetUpdateHelp.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7875
   ScaleWidth      =   6825
   Begin VB.CommandButton Command1 
      Caption         =   "닫기"
      Height          =   450
      Left            =   5490
      TabIndex        =   2
      Top             =   225
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   7080
      Left            =   210
      TabIndex        =   1
      Top             =   720
      Width           =   6465
   End
   Begin VB.Label Label1 
      Caption         =   "도움말"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   165
      TabIndex        =   0
      Top             =   225
      Width           =   2415
   End
End
Attribute VB_Name = "FrmInSoftNetUpdateHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim i   As Integer
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    i = 0
    List1.Clear
    List1.AddItem "설정 방법"
    List1.AddItem "  "
    List1.AddItem "   1.파일위치"
    List1.AddItem "  "
    List1.AddItem "     - 서버에 프로그램을 다운 받을수 있는 위치를 적는다.        "
    List1.AddItem "     - 기본적으로 (주)인소프트넷의 서버를 이용하여 프로그래을   "
    List1.AddItem "       업그레이드 받는 것을 원칙으로 한다.                      "
    List1.AddItem "       기본값은 [ http://www.clean-aid.co.kr:8090/business/  ]   "
    List1.AddItem "       * 지정폴더 -> 업그레이드할 프로그램이 존재하는 위치      "
    List1.AddItem "  "
    List1.AddItem "   2. 파일명 "
    List1.AddItem "      - 현재 업그레이드할 프로그램의 명을 입력한다. "
    List1.AddItem "      - [백상영업.exe]"
    List1.AddItem "      - 예) 백상영업.exe"
    List1.AddItem "  "
    List1.AddItem "   3. 업데이트폴더"
    List1.AddItem "      - 매우 중요한 위치입니다."
    List1.AddItem "      - 현재 사용하고 있는 프로그램이 설치되어 있는 폴더를 정확"
    List1.AddItem "        하게 입력해야 합니다. 만일 잘못되면 정상적으로 동작 하지"
    List1.AddItem "        않을수 있습니다. "
    List1.AddItem "  "
    List1.AddItem "   4. 업데이트파일명"
    List1.AddItem "      - 2번 항목의 파일명과 같다 단) UP 를 붙여 주어야 한다."
    List1.AddItem "      - 예) 백상영업UP.exe"
    List1.AddItem "      - 주의 하시기 바랍니다. UP는 정해진 명칭 입니다."
    List1.AddItem "  "
    List1.AddItem "  ---------------------------------------------------------------"
    List1.AddItem "  - (주) 크린에이드 www.clean-aid.co.kr                          "
    List1.AddItem "  ---------------------------------------------------------------"
    List1.AddItem "  - 주소 : 경기도 남양주시 진접읍 내각리 726-11 "
    List1.AddItem "                                                                 "
    List1.AddItem "  - 전화 : 031 - 522 - 2000"
    List1.AddItem "  -        "
    List1.AddItem "  - 팩스 : 031 - 522 - 2085"
    List1.AddItem "  ---------------------------------------------------------------"
    List1.AddItem "  -        "
    List1.AddItem "  ---------------------------------------------------------------"

End Sub
