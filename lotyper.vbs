'©MapMaths Lisence MIT'
Dim Key
Set Key = Wscript.CreateObject("Wscript.Shell")
Mark = InputBox ("How many exclamation marks do you wish?")
Wscript.Sleep 1000
For i = 1 To Mark
  Key.SendKeys "!"
Next
Key.SendKeys "{ENTER}Done. With " + x + " exclamation mark in total. {ENTER}"