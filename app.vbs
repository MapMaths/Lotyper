' ©MapMaths Lotyper
' @file app.vbs
' Lisence MIT

Dim Key
Dim Mark
Dim IfQuit
Dim StepSetup
Dim TitleSetup
Dim StyleSetup
Dim MsgSetup

Set Key = WScript.CreateObject("WScript.Shell")
Set TitleSetup = "Start using Lotyper"
Set StyleSetup = MsgBoxStyle.OKCancel
Set MsgSetup = "Welcome to Lotyper :D" + vbCrLf + "Now follow these steps."

Set StepSetup = MsgBox(MsgSetup, StyleSetup, TitleSetup)
Do
  Mark = InputBox ("Input a sentence")
  Wscript.Sleep 1000
  For i = 1 To Mark
    Key.SendKeys "!"
  Next
  Key.SendKeys "{ENTER}Done. With " + x + " exclamation mark in total. {ENTER}"
Loop