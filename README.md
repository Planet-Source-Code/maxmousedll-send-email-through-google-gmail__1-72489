<div align="center">

## Send email through Google gmail


</div>

### Description

This code uses CDO to send an e-mail using your Google gmail account
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[MaxMouseDLL](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/maxmousedll.md)
**Level**          |Intermediate
**User Rating**    |4.4 (40 globes from 9 users)
**Compatibility**  |VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/maxmousedll-send-email-through-google-gmail__1-72489/archive/master.zip)

### API Declarations

You must add a reference to Microsoft CDO For Windows 2000 library.


### Source Code

```
Public Function SendMail(msgBody As String)
Dim lobj_cdomsg As CDO.Message
Set lobj_cdomsg = New CDO.Message
lobj_cdomsg.Configuration.Fields(cdoSMTPServer) = "smtp.gmail.com"
lobj_cdomsg.Configuration.Fields(cdoSMTPServerPort) = 465
lobj_cdomsg.Configuration.Fields(cdoSMTPUseSSL) = True
lobj_cdomsg.Configuration.Fields(cdoSMTPAuthenticate) = 1
lobj_cdomsg.Configuration.Fields(cdoSendUsername) = "username@googlemail.com"
lobj_cdomsg.Configuration.Fields(cdoSendPassword) = "password"
lobj_cdomsg.Configuration.Fields(cdoSMTPConnectionTimeout) = 30
lobj_cdomsg.Configuration.Fields(cdoSendUsingMethod) = 2
lobj_cdomsg.Configuration.Fields.Update
lobj_cdomsg.To = "someone@somewhere.com"
lobj_cdomsg.From = "username@googlemail.com"
lobj_cdomsg.Subject = "subject"
lobj_cdomsg.TextBody = "body"
'lobj_cdomsg.AddAttachment ("filepath")
lobj_cdomsg.Send
Set lobj_cdomsg = Nothing
End Function
```

