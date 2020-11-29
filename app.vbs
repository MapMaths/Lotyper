' ©MapMaths Lotyper
' app.vbs
' Lisence MIT

Dim Sentence
Dim StepSetup
Dim Times
Dim Reuse

Set Key = WScript.CreateObject("WScript.Shell")

Private Const Message_Setup_1 = "Welcome to Lotyper :D"
Private Const Message_Setup_2 = "Firstly, open yourtext area."
Private Const Message_Setup_3 = "Then press OK to continue and follow the instructions."
Private Const Message_Setup_4 = "Have a good time using Lotyper!"
Private Const Message_1 = "Now input the sentence tou want to Lotype."
Private Const Message_2 = "Input the value of times you want to Lotype."
Private Const Message_3 = "Now press 'OK' and open your text area in one second."
Private Const Message_4 = "Reuse Lotyper?"

Private Const Title_Setup = "Start - Lotyper"
Private Const Title_1 = "Step 1 - Lotyper"
Private Const Title_2 = "Step 2 - Lotyper"
Private Const Title_3 = "Step 3 - Lotyper"
Private Const Title_4 = "Done! - Lotyper"

StepSetup = MsgBox(Message_Setup_1 & vbCrLf & Message_Setup_2 & vbCrLf & Message_Setup_3 & vbCrLf & Message_Setup_4, vbOKCancel, Title_Setup)
If StepSetup = vbCancel Then
  WScript.Quit
Else
  Do
    Sentence = InputBox(Message_1, Title_1, vbOKOnly)
    Times = InputBox(Message_2, Title_2, vbOKOnly)
    MsgBox Message_3, vbOKOnly, Title_3
    Wscript.Sleep 1000
    Dim I
    For I = 1 To Times
      Key.SendKeys Sentense
    Next
    Reuse = MsgBox(Message_4, vbOKCancel, Title_4)
    If Reuse = vbCancel Then
      WScript.Quit
    End If
  Loop
End If