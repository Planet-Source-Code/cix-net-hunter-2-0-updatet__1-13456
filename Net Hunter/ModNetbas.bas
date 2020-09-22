Attribute VB_Name = "ModNetbas"
Option Explicit
Dim NodeX As Node
Global strFilters As String
Global blnIsdone As Boolean
Public intSites As Integer
Function CreateDir(strUrl As String) As String
On Error Resume Next
 Dim StrPath As String
 Dim StrTemp As String
 Dim StrDirs() As String
 Dim lngStartPoint As Long
 Dim intNrDirs As Integer
 Dim strUrl2 As String
 
  strUrl2 = strUrl
  strUrl2 = GetShortname2(strUrl2)
  strUrl2 = Mid$(strUrl2, 1, Len(strUrl2) - 1)
  StrPath = Hunter.strDir
  StrTemp = StrPath
  If InStr(1, LCase(StrPath), "http:", vbTextCompare) Then StrPath = Mid$(StrPath, 8)
  If InStr(1, LCase(strUrl2), "http:", vbTextCompare) Then strUrl2 = Mid$(strUrl2, 8)
  If Right$(StrPath, 1) = "/" Then StrPath = Mid$(StrPath, 1, Len(StrPath) - 1)
  If InStr(1, strUrl2, "@", vbTextCompare) Then strUrl2 = Mid$(strUrl2, InStr(1, strUrl2, "@", vbTextCompare) + 1)
  If InStr(1, StrPath, "@", vbTextCompare) Then StrPath = Mid$(StrPath, InStr(1, StrPath, "@", vbTextCompare) + 1)
  If InStr(1, LCase(strUrl2), LCase(StrPath), vbTextCompare) Then
   lngStartPoint = InStr(1, LCase(strUrl2), LCase(StrPath), vbTextCompare) + Len(StrPath)
   StrTemp = Mid$(strUrl2, lngStartPoint)
   StrDirs() = Split(StrTemp, "/")
   StrTemp = StrPath
   StrPath = Hunter.Text3.Text & "\" & StrTemp
   For intNrDirs = LBound(StrDirs) To UBound(StrDirs)
    DoEvents
    StrPath = Replace(StrPath, "?", "_", 1, -1)
    StrPath = Replace(StrPath, "~", "_", 1, -1)
    StrPath = Replace(StrPath, "^", "_", 1, -1)
    StrPath = Replace(StrPath, "&", "_", 1, -1)
    StrPath = Replace(StrPath, "#", "_", 1, -1)
On Error Resume Next
    MkDir StrPath & "\" & StrDirs(intNrDirs)
    If StrDirs(intNrDirs) = "" Then
    Else
     StrPath = StrPath & "\" & StrDirs(intNrDirs)
    End If
   Next intNrDirs
  Else
    StrTemp = strUrl
    StrTemp = GetShortname2(StrTemp)
   If InStr(1, LCase(StrTemp), "http:", vbTextCompare) Then StrTemp = Mid$(StrTemp, 8)
   If InStr(1, StrTemp, "@", vbTextCompare) Then StrTemp = Mid$(StrTemp, InStr(1, StrTemp, "@", vbTextCompare) + 1)
    StrPath = Hunter.Text3.Text
    StrDirs() = Split(StrTemp, "/")
   For intNrDirs = LBound(StrDirs) To UBound(StrDirs)
   On Error Resume Next
    DoEvents
    StrPath = Replace(StrPath, "?", "_", 1, -1)
    StrPath = Replace(StrPath, "~", "_", 1, -1)
    StrPath = Replace(StrPath, "^", "_", 1, -1)
    StrPath = Replace(StrPath, "&", "_", 1, -1)
    StrPath = Replace(StrPath, "#", "_", 1, -1)
    MkDir StrPath & "\" & StrDirs(intNrDirs)
    If StrDirs(intNrDirs) = "" Then
    Else
     StrPath = StrPath & "\" & StrDirs(intNrDirs)
    End If
   Next intNrDirs
  End If
    StrPath = Replace(StrPath, "?", "_", 1, -1)
    StrPath = Replace(StrPath, "~", "_", 1, -1)
    StrPath = Replace(StrPath, "^", "_", 1, -1)
    StrPath = Replace(StrPath, "&", "_", 1, -1)
    StrPath = Replace(StrPath, "#", "_", 1, -1)
    
    CreateDir = StrPath
End Function
Function GetShortname2(SName As String) As String
Dim I As Long
I = InStrRev(SName, "/", -1, vbTextCompare)
GetShortname2 = Mid$(SName, 1, I)
End Function
Function GetShortname(SName As String) As String
Dim I As Long
I = InStrRev(SName, "/", -1, vbTextCompare) + 1
GetShortname = Mid$(SName, I)
End Function
'http://3333333:2222222@www.teenfresh.com/members/movies/mpegs/girl1/
Function GetShortnameEX(SName As String) As String
Dim I As Long
I = InStrRev(SName, ".", -1, vbTextCompare) + 1
GetShortnameEX = Mid$(SName, I)
End Function

Function GetRealAdress(Site As String) As String
Dim Tmpsite As String
Dim I As Long
HaltErr:

If Left(Site, 1) = "/" Or Left(Site, 2) = "./" Then
 If Tmpsite = "" Then
  GetRealAdress = GetAliasAdress(Hunter.Text1.Text) & Mid$(Site, 2)
  Tmpsite = ""
 Else
  GetRealAdress = Tmpsite & Mid$(Site, 2)
  Tmpsite = ""
 End If
 
ElseIf Left(Site, 3) = "../" Then
 I = InStrRev(Hunter.Text1.Text, "/", -1, vbTextCompare)
 I = InStrRev(Hunter.Text1.Text, "/", I, vbTextCompare)
 Tmpsite = Mid$(Hunter.Text1.Text, 1, I)
 Site = Mid$(Site, 3)
 GoTo HaltErr
ElseIf LCase(Left$(Site, 5)) = "http:" Then
 GetRealAdress = Site
Else
 
 GetRealAdress = GetShortname2(Hunter.Text1.Text) & Site
End If
End Function

Function GetRealAdress2(Site As String, OldSite As String) As String
Dim Tmpsite As String
Dim I As Long
HaltErr:

If Left(Site, 1) = "/" Or Left(Site, 2) = "./" Then
 If Tmpsite = "" Then
  GetRealAdress2 = GetAliasAdress(OldSite) & Mid$(Site, 2)
  Tmpsite = ""
 Else
  GetRealAdress2 = Tmpsite & Mid$(Site, 2)
  Tmpsite = ""
 End If
 
ElseIf Left(Site, 3) = "../" Then
 If InStr(1, LCase(OldSite), ".htm") Or InStr(1, LCase(OldSite), ".shtm") Or InStr(1, LCase(OldSite), ".asp") Or InStr(1, LCase(OldSite), ".php") Then OldSite = GetShortname2(OldSite)
 If Right$(OldSite, 1) = "/" Then OldSite = Mid$(OldSite, 1, Len(OldSite) - 1)
 I = InStrRev(OldSite, "/", -1, vbTextCompare) - 1
' I = InStrRev(OldSite, "/", I, vbTextCompare)
 Tmpsite = Mid$(OldSite, 1, I)
 'MsgBox Tmpsite
 Site = Mid$(Site, 4)
GetRealAdress2 = GetRealAdress2(Site, Tmpsite)
ElseIf LCase(Left$(Site, 5)) = "http:" Then
 GetRealAdress2 = Site
Else
If InStr(1, LCase(OldSite), ".htm") Or InStr(1, LCase(OldSite), ".shtm") Or InStr(1, LCase(OldSite), ".asp") Or InStr(1, LCase(OldSite), ".php") Then
 GetRealAdress2 = GetShortname2(OldSite) & Site
Else
 If Right$(OldSite, 1) = "/" Then
  GetRealAdress2 = GetShortname2(OldSite) & Site
 Else
  OldSite = OldSite & "/"
  GetRealAdress2 = GetShortname2(OldSite) & Site
 End If
End If
End If
End Function

Function GetAliasAdress(Site As String) As String
Dim I As Long
If Left$(Site, 4) = "http" Then
 I = InStr(8, Site, "/", vbTextCompare)
 GetAliasAdress = Mid$(Site, 1, I)
 If InStr(1, GetAliasAdress, "@", vbTextCompare) Then GetAliasAdress = Mid$(GetAliasAdress, InStr(1, GetAliasAdress, "@", vbTextCompare) + 1)
Else
I = InStr(1, Site, "/", vbTextCompare)
 GetAliasAdress = Mid$(Site, 1, I)
 If InStr(1, GetAliasAdress, "@", vbTextCompare) Then GetAliasAdress = Mid$(GetAliasAdress, InStr(1, GetAliasAdress, "@", vbTextCompare) + 1)
End If
End Function

Function RipHtmlForMedia(Url As String, OutputList As ListBox, Net As Inet, StartPoint As Long, SecondFile As String, Frm As Form, Capline As String) As String
On Error Resume Next
Dim Download As String, Tmp As String
Dim EndPoint As Long
Dim I As Long
If CancelOrder = True Then Exit Function
'ProcesBar.Visible = False
Capline = "Connecting to : " & Url
Capline = "Seaching for URL's in (" & Url & ")"

DoEvents
If StartPoint > 1 Then
'SecondFile = Space$(Len(Download))
Download = ""
Download = Space$(Len(SecondFile))
Download = SecondFile
SecondFile = ""
Else
 Download = Net.OpenURL(Url, icString) ': RipHtmlForUrLs = "Downloading The HTML File"
 
 Do While Net.StillExecuting = True
  DoEvents
 Loop
 
  'If Len(Download) > 6400000 Then Download = ""
 If Download = "" Then
  RippingDone = True
  Net.Cancel
  Download = ""
  Tmp = ""
  StartPoint = 1
  EndPoint = 1
  Exit Function
 End If
  
' Form1.Caption = "URL downloaded"
 'MsgBox Download
End If
ProcessLineMaxValue = Len(Download)
DoEvents
'######################3 RET MIG TIL + "<img " og sæt længde til 10
If InStr(StartPoint, LCase(Download), "src=", vbTextCompare) Then
  StartPoint = InStr(StartPoint, LCase(Download), "src=", vbTextCompare) + 5
If InStr(StartPoint, LCase(Download), """", vbTextCompare) Then
 EndPoint = InStr(StartPoint, LCase(Download), """", vbTextCompare)
Else
 If InStr(StartPoint, LCase(Download), ">", vbTextCompare) Then
  EndPoint = InStr(StartPoint, LCase(Download), ">", vbTextCompare)
 End If
  'EndPoint = InStr(StartPoint, LCase(Download), """", vbTextCompare)
 End If
  ProcessLine_ValueChange Hunter.Picture1, StartPoint
  Tmp = Mid$(Download, StartPoint, (EndPoint - StartPoint))
  Tmp = Trim$(Tmp)
  Tmp = TrimUrl(Tmp, 1)
  
   If InStr(1, LCase(Tmp), "mailto", vbTextCompare) Then
   Else
     Tmp = GetRealAdress2(Tmp, Url)
     OutputList.AddItem Tmp
    
   Set NodeX = Hunter.Tree1.Nodes.Add("images" & intSites, tvwChild, Tmp, Tmp, 3)
   End If
  
  Tmp = ""
  StartPoint = EndPoint + 1
  DoEvents
  RipHtmlForMedia Url, OutputList, Net, StartPoint, Download, Frm, Capline
Else
  Capline = "Done..."
  Download = ""
  Tmp = ""
  StartPoint = 1
  EndPoint = 1
  ProcessLine_ValueChange Hunter.Picture1, ProcessLineMaxValue
  RippingDone = True
  Net.Cancel
End If
End Function

