Dim Message, Speak
Message = InputBox("Bir şeyler yazınız...")
Mesajyaz = msgbox(Message)
Set Speak = CreateObject("sapi.spvoice")
Speak.Speak Message