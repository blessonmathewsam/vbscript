subject = "Sending Email by vbs"
body = "Body of message"

Const from = "self@gmail.com"
Const fromname = "name"
Const to = "to@destination.com"
Const SMTPLogon = "user@gmail.com"
Const SMTPPassword = "password"
Const SMTPPort = 465

' MESSAGE

Set objMessage = CreateObject("CDO.Message")
objMessage.Subject = subject
objMessage.From = """" & fromname & """ <" & from & ">"
objMessage.To = to
objMessage.TextBody = body

' SERVER CONFIGURATION

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusername") = SMTPLogon

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendpassword") = SMTPPassword

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTPPort

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

objMessage.Configuration.Fields.Update
'Now send the message!
On Error Resume Next
objMessage.Send

If Err.Number <> 0 Then
    MsgBox Err.Description,16,"Error Sending Mail"
Else 
    MsgBox "Mail was successfully sent !",64,"Information"
End If