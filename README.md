<div align="center">

## Email with Attachments

<img src="Mail dll pic.jpg">
</div>

### Description

This code once compiled to an ocx, can be added as a user control to any VB6 form.

It allows you the user to send an mail via an smtp server, but where this code shines is the ability to add an attachment to your Email.
 
### More Info
 
Email1.MailFrom = "Dogsbollox@topbloke.com"

Email1.MailMessage = "This is a test"

Email1.MailSubject = "This is a test"

Email1.MailTo = "andy@andythughes.co.uk"

Email1.PortNumber = 25

Email1.ServerName = "your smtp server"

Email1.Attachment = "C:\Newsletter-v1.jpg"

Email1.SendMail

An event is raised MailFailed

this returns a number but there is a public variable within the routine that returns as text

the actual Text, this can be an error message or a message to say Mail Sent Successfully


<span>             |<span>
---                |---
**Submitted On**   |2006-02-17 12:23:40
**By**             |[Dogs](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dogs.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Email\_with1975082212006\.zip](https://github.com/Planet-Source-Code/dogs-email-with-attachments__1-64399/archive/master.zip)








