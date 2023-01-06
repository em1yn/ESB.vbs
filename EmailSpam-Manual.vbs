Dim User
Dim Pass
Dim Name
Dim Input
Dim Input2
Dim Input3
X=MsgBox("Welcome! To begin, click OK and follow the instructions provided.",0,"Email Spam Bot v1.0")
User = InputBox("Enter your Gmail address:")
Pass = InputBox("Enter your Gmail password:"& vbCrLf & ""& vbCrLf & "Please note passwords are NOT stored in this script and are case sensitive.")
Name = InputBox("Enter your name:")
Input = InputBox("Enter email of recipient:")
Input2 = InputBox("Enter subject:")
Input3 = InputBox("Enter message:")
EmailSubject = (""& Input2)
EmailBody = (""& Input3)

'Const EmailFrom = ""
'Const EmailFromName = ""

Const SMTPServer = "smtp.gmail.com"
'Const SMTPLogon = ""
'Const SMTPPassword = ""
Const SMTPSSL = True
Const SMTPPort = 456

Const cdoSendUsingPickup = 1  'Send message using local SMTP service pickup directory.
Const cdoSendUsingPort = 2  'Send the message using SMTP over TCP/IP networking.

Const cdoAnonymous = 0  ' No authentication
Const cdoBasic = 1  ' BASIC clear text authentication
Const cdoNTLM = 2  ' NTLM, Microsoft proprietary authentication

' First, create the message

Set objMessage = CreateObject("CDO.Message")
objMessage.Subject = EmailSubject
objMessage.From = "<" & User & Name & ">"
objMessage.To = "<" & Input & ">"
objMessage.TextBody = EmailBody

' Second, configure the server

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTPServer

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusername") = User

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Pass

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTPPort

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = SMTPSSL

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

objMessage.Configuration.Fields.Update

do
objMessage.Send
loop
