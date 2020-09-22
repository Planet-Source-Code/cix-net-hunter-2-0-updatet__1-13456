VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Hunter 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   " VirusTeam 2000 - Net Hunter Version 2.0"
   ClientHeight    =   5670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Hunter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Height          =   255
      Left            =   5520
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   26
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command4 
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "&Browse"
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1935
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Top             =   1920
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   16
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5520
      Top             =   2040
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      Caption         =   "&More"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5160
      TabIndex        =   13
      Top             =   2520
      Width           =   852
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6360
      Top             =   840
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Hunter.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Hunter.frx":0BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Hunter.frx":1036
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Hunter.frx":1488
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Hunter.frx":18DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Hunter.frx":1D2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Hunter.frx":44E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Hunter.frx":6C94
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView Tree1 
      Height          =   2175
      Left            =   0
      TabIndex        =   7
      Top             =   3480
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3836
      _Version        =   393217
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.ListBox lstpicdown2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6120
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.ListBox Download 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6120
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.ListBox lstPicdown 
      Height          =   285
      Left            =   6120
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   135
      Left            =   50
      ScaleHeight     =   105
      ScaleWidth      =   5865
      TabIndex        =   1
      Top             =   3045
      Width           =   5900
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5400
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Sched&ule"
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
      Left            =   3140
      TabIndex        =   25
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Appearance      =   0  'Flat
      Caption         =   "&Settings"
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
      Left            =   2180
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
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
      Left            =   1100
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "&Download"
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
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   180
      Left            =   5640
      Picture         =   "Hunter.frx":9448
      Stretch         =   -1  'True
      ToolTipText     =   "Put Hunter in the Syetem Tray"
      Top             =   20
      Width           =   210
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3120
      TabIndex        =   24
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Output Path :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Url :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   22
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   5400
      Picture         =   "Hunter.frx":9E9C
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Level(s) :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   21
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   1905
      Index           =   1
      Left            =   4200
      Picture         =   "Hunter.frx":A1A6
      Stretch         =   -1  'True
      Top             =   850
      Width           =   1830
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Hunter 2.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   2130
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Hunter 2.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Index           =   3
      Left            =   195
      TabIndex        =   11
      Top             =   420
      Width           =   2130
   End
   Begin VB.Shape Shape4 
      Height          =   5655
      Left            =   0
      Top             =   0
      Width           =   6015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VirusTeam 2000"
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
      Index           =   0
      Left            =   3240
      TabIndex        =   9
      Top             =   0
      Width           =   1500
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   255
      Left            =   3120
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label lblAdress 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2520
      TabIndex        =   8
      Top             =   480
      Width           =   60
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      Height          =   3855
      Left            =   0
      Top             =   3480
      Width           =   6015
   End
   Begin VB.Label lblsites 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   50
      TabIndex        =   6
      Top             =   3200
      Width           =   45
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
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   3480
      MouseIcon       =   "Hunter.frx":4018A
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   240
      Width           =   2400
   End
   Begin VB.Label Statusline 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Downloading :"
      Height          =   225
      Left            =   45
      TabIndex        =   0
      Top             =   2790
      Width           =   1185
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   1
      Left            =   0
      Top             =   2760
      Width           =   6015
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
End
Attribute VB_Name = "Hunter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Xnode As Node
Dim I As Long, J As Long
Dim Levels As Integer
Dim Nrlevel As Integer
Public Usename As Integer
Dim StopAll As Boolean
Public strDir As String
Dim strSaveDir As String
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
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInstrtAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Dim WithEvents NewTray As CTray
Attribute NewTray.VB_VarHelpID = -1

Private Sub Check1_Click()
If Check1.Value = Checked Then
 Usename = 1
Else
 Usename = 2
End If
End Sub

Private Sub Command1_Click()
frmBrowse.Show
End Sub

Private Sub Command2_Click()
On Error Resume Next
blnIsdone = False
StopAll = False
Timer1.Enabled = True
If Right$(Text3.Text, 1) = "\" Then Text3.Text = Mid$(Text3.Text, 1, Len(Text3.Text) - 1)
lstPicdown.Clear
lstpicdown2.Clear
Download.Clear
Nrlevel = 0
'If Val(Text2.Text) = 0 Then
' RipHtmlForMedia Text1.Text, Download, Inet1, 1, 0, Me, Me.Caption
' GoTo OnlyLevel
'End If
lstpicdown2.AddItem Text1.Text
Statusline.Caption = "Makeing DownList.. at level (" & (Nrlevel) + 1 & ")"
lblAdress.Caption = GetAliasAdress(Text1.Text)
NewTray.TipText = "Download from : " & GetAliasAdress(Text1.Text)
strDir = GetAliasAdress(Text1.Text)
If InStr(1, LCase(strDir), "http:", vbTextCompare) Then strDir = Mid$(strDir, 8)
If Right$(strDir, 1) = "/" Then strDir = Mid$(strDir, 1, Len(strDir) - 1)
If InStr(1, strDir, "@", vbTextCompare) Then strDir = Mid$(strDir, InStr(1, strDir, "@", vbTextCompare) + 1)
MkDir Text3.Text & "\" & strDir
strDir = GetAliasAdress(Text1.Text)
Set Xnode = Tree1.Nodes.Add(, , "site" & intSites, GetAliasAdress(Text1.Text), 4, 4)
Set Xnode = Tree1.Nodes.Add("site" & intSites, tvwChild, "urls" & intSites, "HTML", 6, 7)
Set Xnode = Tree1.Nodes.Add("site" & intSites, tvwChild, "images" & intSites, "Media", 6, 7)
Levels = Val(Text2.Text)
RipHtmlForUrLs Text1.Text, lstPicdown, Inet1, 1, 0, Me, Me.Caption
WaitforIT:
Statusline.Caption = "Filtering list.."
RemoveDupes lstPicdown
For I = 0 To lstPicdown.ListCount - 1
 If StopAll = True Then Exit Sub
 lstPicdown.ListIndex = I
 If lstPicdown.Text = "" Then
  lstPicdown.RemoveItem I
 Else
  lstpicdown2.AddItem GetRealAdress(lstPicdown.Text)
 End If
 DoEvents
Next I
 lstPicdown.Clear
If Nrlevel = Levels Then
 Nrlevel = Nrlevel + 1
 RemoveDupes lstpicdown2
 Statusline.Caption = "Finding Media.."
 RippingDone = False
 lstpicdown2.AddItem Text1.Text
 For I = 0 To lstpicdown2.ListCount - 1
  If StopAll = True Then Exit Sub
  lstpicdown2.ListIndex = I
  Statusline.Caption = "Finding Media(files) in (Page " & I & ")"
  RipHtmlForMedia lstpicdown2.Text, Download, Inet1, 1, 0, Me, Me.Caption
  DoEvents
  Do While RippingDone = False
   DoEvents
  Loop
  Next I
  Statusline.Caption = "Filtering list.."
  RemoveDupes Download
  Statusline.Caption = "Download " & GetShortnameEX(Download.Text) & " file number(" & J & ") "
  Dim FilByte() As Byte
  ProcessLineMaxValue = Download.ListCount
  For J = 0 To Download.ListCount - 1
  If StopAll = True Then Exit Sub
  Download.ListIndex = J
  Statusline.Caption = "Download " & GetShortnameEX(Download.Text) & " file number(" & J & ") "
  ProcessLine_ValueChange Hunter.Picture1, (J + 1)
   DoEvents
   If Download.Text = "" Then
   Else
    FilByte() = Inet1.OpenURL(Download.Text, icByteArray)
   Do While Inet1.StillExecuting = True
    DoEvents
   Loop
 
    Close #9
    If frmSettings.Check1.Value = Checked Then
     strSaveDir = CreateDir(Download.Text)
     Open strSaveDir & "\" & GetShortname(Download.Text) For Binary As #9
    Else
     strSaveDir = CreateDir(Download.Text)
     Open strSaveDir & "\Image" & J & Right$(Download.Text, 4) For Binary As #9
    End If
     Put #9, , FilByte()
    Close #9
   End If
   Next J
OnlyLevel:
 Timer1.Enabled = False
 NewTray.TipText = "Net Hunter 2.0"
 Statusline.Caption = "Download done.."
 blnIsdone = True
 Exit Sub
Else
On Error Resume Next
 For I = 0 To lstpicdown2.ListCount - 1
 If StopAll = True Then Exit Sub
  lstpicdown2.ListIndex = I
  If lstpicdown2.Text = "" Then
     lstpicdown2.RemoveItem I
  ElseIf LCase(Right$(lstpicdown2.Text, 4)) = ".jpg" Or LCase(Right$(lstpicdown2.Text, 4)) = ".gif" Or LCase(Right$(lstpicdown2.Text, 4)) = ".bmp" Then
  Download.AddItem lstpicdown2.Text
  lstpicdown2.RemoveItem I
  Else
  lstPicdown.AddItem lstpicdown2.Text
  End If
  DoEvents
 Next I
FormHere:
Statusline.Caption = "Makeing DownList.. at level (" & (Nrlevel) + 1 & ")"
For I = 0 To lstpicdown2.ListCount - 1
   If StopAll = True Then Exit Sub
  Statusline.Caption = "Makeing DownList from (Page " & I & ") at Level (" & (Nrlevel) + 1 & ")"
  lstpicdown2.ListIndex = I
  RipHtmlForUrLs lstpicdown2.Text, lstPicdown, Inet1, 1, 0, Me, Me.Caption
  DoEvents
  Do While RippingDone = False
   DoEvents
  Loop
Next I
lstpicdown2.Clear
 Nrlevel = Nrlevel + 1
End If
If Nrlevel <= Levels Then GoTo WaitforIT

End Sub

Private Sub Command3_Click()
Dim Svar As Integer
If frmSchedule.timCheck.Enabled = True Then
 Svar = MsgBox("Would you like to start the scheduled job ?", vbYesNo + vbQuestion, "Net Hunter 2.0")
 If Svar = vbYes Then
  blnIsdone = True
 Else
  StopAll = True
 End If
Else
  StopAll = True
End If

End Sub

Private Sub Command4_Click()
NewTray.DeleteIcon
SaveSetting "Net Hunter", "Settings", "LastUrl", Text1.Text
SaveSetting "Net Hunter", "Settings", "Levels", Text2.Text
SaveSetting "Net Hunter", "Settings", "SaveDir", Text3.Text
SaveSetting "Net Hunter", "Settings", "Names", Usename
SaveSetting "Net Hunter", "Settings", "Filters", strFilters
Inet1.Cancel
End
End Sub

Private Sub Command5_Click()
If Command5.Caption = "&Less" Then
 Command5.Caption = "&More"
 Me.Height = 3490
Else
 Me.Height = 5650
 Command5.Caption = "&Less"
End If
End Sub

Private Sub Command6_Click()
frmSettings.Show
End Sub

Private Sub Command7_Click()
frmSchedule.Show
End Sub

Private Sub Form_Load()
SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
Set NewTray = New CTray
Picture2.Picture = Image1(0).Picture
NewTray.PicBox = Picture2
NewTray.ShowIcon
NewTray.TipText = "Net Hunter 2.0"
Me.Height = 3490
Text1.Text = GetSetting("Net Hunter", "Settings", "LastUrl", "http://www.site.com/pictures.htm")
Text2.Text = GetSetting("Net Hunter", "Settings", "Levels", 1)
Text3.Text = GetSetting("Net Hunter", "Settings", "SaveDir", "c:\temp")
Usename = GetSetting("Net Hunter", "Settings", "Names", 2)
strFilters = GetSetting("Net Hunter", "Settings", "Filters", ".mov/.asf/.viv/.ra/.rar/.ram/.ace/.zip/.mpg/.mpeg/.avi/.gif/.jpg/.bmp/.ani/.ico/.js/.class/.mdb/.mbr/.jar/.swf/.vbs/.arj/.jpeg/.pdf/.tif/.mp3/.wav/.vid/.wsf/.txt/.doc/.mde/.movie/.tgz")
frmSettings.txtMediafiles.Text = strFilters
If Usename = 1 Then
frmSettings.Check1.Value = Checked
Else
frmSettings.Check1.Value = Unchecked
End If
End Sub

Private Sub Form_Terminate()
NewTray.DeleteIcon
End Sub

Private Sub Form_Unload(Cancel As Integer)
NewTray.DeleteIcon
End Sub

Private Sub Image2_Click()
Me.Hide
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub

Private Sub Label5_Click()
GotoURL Label5.Caption
End Sub

Private Sub lstPicdown_Validate(Cancel As Boolean)
lblsites.Caption = "Pages : " & lstPicdown.ListCount & " images : " & Download.ListCount
End Sub

Private Sub lstpicdown2_Validate(Cancel As Boolean)
lblsites.Caption = "Pages : " & lstpicdown2.ListCount & " images : " & Download.ListCount
End Sub

Private Sub NewTray_LButtonDown()
Me.Show
End Sub

Private Sub NewTray_RButtonDown()
Me.Show
End Sub

Private Sub Timer1_Timer()
DoEvents
If lstpicdown2.ListCount = 0 Then
 lblsites.Caption = "Pages : " & lstPicdown.ListCount & " Media : " & Download.ListCount
Else
 lblsites.Caption = "Pages : " & lstpicdown2.ListCount & " Media : " & Download.ListCount
End If
End Sub

Public Function GoDoIt()
 Command2_Click
End Function

