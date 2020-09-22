VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSchedule 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Settings for Net Hunter"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timCheck 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   2640
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1815
      Left            =   60
      TabIndex        =   8
      Top             =   960
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   3201
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Homepage Url"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Levels"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Save dir"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Add"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   0
      Width           =   3015
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3720
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright VirusTeam 2000 ©"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   60
      TabIndex        =   6
      Top             =   2760
      Width           =   2340
   End
   Begin VB.Shape Shape2 
      Height          =   3135
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   6015
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Schedule"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   3000
      TabIndex        =   2
      Top             =   0
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Net Hunter 2.0 "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.virusteam.cjb.net"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3480
      MouseIcon       =   "frmSchedule.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   240
      Width           =   2400
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   255
      Index           =   1
      Left            =   2880
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Hunter 2.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   435
      Index           =   1
      Left            =   200
      TabIndex        =   3
      Top             =   420
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   0
      Left            =   0
      Top             =   240
      Width           =   6015
   End
   Begin VB.Image Image1 
      Height          =   2280
      Left            =   3720
      Picture         =   "frmSchedule.frx":030A
      Stretch         =   -1  'True
      Top             =   840
      Width           =   2310
   End
End
Attribute VB_Name = "frmSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_SHOWINWINDOW = &H40
Private Const SWP_HIDEWINDOWS = &H80
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const HWND_TOP = 0
Private Const HWND_BOTTOM = 1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInstrtAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Dim DL As Boolean
Dim intMaxsites As Integer


Private Sub Command1_Click()
If ListView1.ListItems.Count > 0 Then
  blnIsdone = True
 intMaxsites = ListView1.ListItems.Count
 'MsgBox intMaxsites
 intSites = 1
 Me.Hide
 timCheck.Enabled = True
Else
 Exit Sub
End If
End Sub

Private Sub Command2_Click()
If ListView1.ListItems.Count > 0 Then
 Me.Hide
Else
 Unload Me
End If
End Sub

Private Sub Command3_Click()
Dim XItem As ListItem
Dim strAsk As String
strAsk = InputBox("Enter the Url for the site : ", "Net Hunter 2.0")
Set XItem = ListView1.ListItems.Add(, strAsk, strAsk)
strAsk = InputBox("Enter the number of levels : ", "Net Hunter 2.0")
XItem.SubItems(1) = strAsk
strAsk = InputBox("Enter a path to save the site : ", "Net Hunter 2.0")
XItem.SubItems(2) = strAsk
End Sub

Private Sub Form_Load()
SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub



Private Sub Label5_Click()
GotoURL Label5.Caption
End Sub

Private Sub Label6_Click()

End Sub

Private Sub timCheck_Timer()
If intSites > intMaxsites Then
 timCheck.Enabled = False
 blnIsdone = False
Else
 If blnIsdone = True Then
  Hunter.Text1.Text = ListView1.ListItems(intSites).Text
  Hunter.Text2.Text = ListView1.ListItems(intSites).SubItems(1)
  Hunter.Text3.Text = ListView1.ListItems(intSites).SubItems(2)
  intSites = intSites + 1
  blnIsdone = False
  Hunter.GoDoIt
 Else
  DoEvents
 End If
End If
End Sub
