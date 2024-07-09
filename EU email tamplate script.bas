Attribute VB_Name = "Module1"

Sub Send_email()


Dim OutApp As Object
Dim OutMail As Object
Dim strbody As String


Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

strbody = "<BODY style = font-size:11pt; font-family:Calibri>" & _
"Hello, <br><br> For test purposes."

On Error Resume Next
With OutMail
.SentOnBehalfOfName = "test@email.com"
.to = "Test <notarealemail@email.com>"
.CC = "notarealemail@email.com"
.Subject = "test"

.display
.HTMLBody = strbody & .HTMLBody

'.send

End With

Set OutApp = Nothing
Set OutMail = Nothing

End Sub
