Attribute VB_Name = "modRipHTML"
Global RippingDone As Boolean
Global TimeOutS As Integer
Global MellemRum As String
Dim NodeX As Node
Dim strSaveHTML As String

Dim K As Long
Function RipHtmlForUrLs(Url As String, OutputList As ListBox, Net As Inet, StartPoint As Long, SecondFile As String, Frm As Form, Capline As String) As String
On Error Resume Next
Dim Download As String, Tmp As String
Dim EndPoint As Long
Dim I As Long
Dim strCheck() As String
Dim blCheck As Boolean
blCheck = False
If CancelOrder = True Then Exit Function
'ProcesBar.Visible = False
Capline = "Connecting to : " & Url
Capline = "Seaching for URL's in (" & Url & ")"

DoEvents
If StartPoint > 1 Then
'SecondFile = Space$(Len(Download))
' Her downloader den fra internettet
Download = ""
Download = Space$(Len(SecondFile))
Download = SecondFile
SecondFile = ""
Else
 Download = Net.OpenURL(Url, icString) ': RipHtmlForUrLs = "Downloading The HTML File"
 Do While Net.StillExecuting = True
  DoEvents
 Loop
 Close #9
 strSaveHTML = CreateDir(Url)
 Open strSaveHTML & "\" & GetShortname(Url) For Output As 9
  Print #9, Download
 Close #9
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
If InStr(1, LCase(Download), "background=", vbTextCompare) Then
 Dim j1 As Long, j2 As Long
  j1 = InStr(1, LCase(Download), "background=", vbTextCompare) + 12
  j2 = InStr(j1 + 1, LCase(Download), """", vbTextCompare)
  Tmp = Mid$(Download, j1, (j2 - j1))
  Tmp = Trim$(Tmp)
  Tmp = TrimUrl(Tmp, 1)
  Tmp = GetRealAdress2(Tmp, Url)
  Hunter.Download.AddItem Tmp
End If
ProcessLineMaxValue = Len(Download)
DoEvents
If InStr(StartPoint, LCase(Download), "href", vbTextCompare) Then
  StartPoint = InStr(StartPoint, LCase(Download), "href", vbTextCompare) + 6
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
  If StopAll = True Then Exit Function
   If InStr(1, LCase(Tmp), "mailto", vbTextCompare) Then
   Else
   Dim intCheck As Long
   strCheck() = Split(strFilters, "/")
   For intCheck = LBound(strCheck()) To UBound(strCheck())
    If Right$(LCase(Tmp), Len(strCheck(intCheck))) = strCheck(intCheck) Then
     blCheck = True
     Exit For
    End If
   Next intCheck
      
   If blCheck = True Then
    Tmp = GetRealAdress2(Tmp, Url)
    Set NodeX = Hunter.Tree1.Nodes.Add("images" & intSites, tvwChild, Tmp, Tmp, 3)
    Hunter.Download.AddItem Tmp
   Else
    Tmp = GetRealAdress2(Tmp, Url)
    Set NodeX = Hunter.Tree1.Nodes.Add("urls" & intSites, tvwChild, Tmp, Tmp, 5)
    OutputList.AddItem Tmp
   End If
   End If
  blCheck = False
  Tmp = ""
  StartPoint = EndPoint + 1
  DoEvents
  RipHtmlForUrLs Url, OutputList, Net, StartPoint, Download, Frm, Capline
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


'#############################################################################################################
' Function TrimUrl
' This Function is to Trim the urls form the function RipHtmlForUrLs.
'
' You have no need for this function, since it only here for use in the function RipHtmlForUrLs
' But if you do, you can call it like this.
' Dim URL as String
' Url = TrimUrl(Url, 1)
'
'#############################################################################################################

Function TrimUrl(Url As String, StartPoint As Long) As String
Dim X As Long, Y As Long
If InStr(StartPoint, Url, ">", vbTextCompare) Then
X = InStr(StartPoint, Url, ">", vbTextCompare) + 1
Url = Mid$(Url, X)
If InStr(StartPoint, Url, "</", vbTextCompare) Then
Y = InStr(StartPoint, Url, "</", vbTextCompare)
Url = Mid$(Url, Y)
End If
TrimUrl Url, X + 1
Else
TrimUrl = Url
End If
End Function

