Imports System.Diagnostics
Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email
	Class SendingBulkEmails

		Public Shared Sub Run()
			' ExStart:SendingBulkEmails
			' Create SmtpClient as client and specify server, port, user name and password
			Dim client As New SmtpClient("mail.server.com", 25, "Username", "Password")

			' Create instances of MailMessage class and Specify To, From, Subject and Message
			Dim message1 As New MailMessage("msg1@from.com", "msg1@to.com", "Subject1", "message1, how are you?")
			Dim message2 As New MailMessage("msg1@from.com", "msg2@to.com", "Subject2", "message2, how are you?")
			Dim message3 As New MailMessage("msg1@from.com", "msg3@to.com", "Subject3", "message3, how are you?")

			' Create an instance of MailMessageCollection class
			Dim manyMsg As New MailMessageCollection()
			manyMsg.Add(message1)
			manyMsg.Add(message2)
			manyMsg.Add(message3)

			' Use client.BulkSend function to complete the bulk send task
			Try
				' Send Message using BulkSend method
				client.Send(manyMsg)
				Console.WriteLine("Message sent")
			Catch ex As Exception
				Trace.WriteLine(ex.ToString())
			End Try
			' ExEnd:SendingBulkEmails
		End Sub

	End Class
End Namespace