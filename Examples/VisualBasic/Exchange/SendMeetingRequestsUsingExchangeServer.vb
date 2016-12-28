Imports System.Net
Imports Aspose.Email.Exchange
Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Exchange
	Class SendMeetingRequestsUsingExchangeServer
		Public Shared Sub Run()
			' ExStart:SendMeetingRequestsUsingExchangeServer
			Try
				Dim mailboxUri As String = "https://ex07sp1/exchange/administrator"
				' WebDAV
				Dim domain As String = "litwareinc.com"
				Dim username As String = "administrator"
				Dim password As String = "Evaluation1"

				Dim credential As New NetworkCredential(username, password, domain)
				Console.WriteLine("Connecting to Exchange server.....")
				' Connect to Exchange Server
				Dim client As New ExchangeClient(mailboxUri, credential)
				' WebDAV
				' Create the meeting request
				Dim app As New Appointment("meeting request", DateTime.Now.AddHours(1), DateTime.Now.AddHours(1.5), Convert.ToString("administrator@") & domain, Convert.ToString("bob@") & domain)
				app.Summary = "meeting request summary"
				app.Description = "description"

				' Create the message and set the meeting request
				Dim msg As New MailMessage()
				msg.From = Convert.ToString("administrator@") & domain
				msg.[To] = Convert.ToString("bob@") & domain
				msg.IsBodyHtml = True
				msg.HtmlBody = "<h3>HTML Heading</h3><p>Email Message detail</p>"
				msg.Subject = "meeting request"
				msg.AddAlternateView(app.RequestApointment(0))

				' Send the appointment
				client.Send(msg)
				Console.WriteLine("Appointment request sent")
			Catch ex As Exception
				Console.WriteLine(ex.Message)
			End Try
			' ExEnd:SendMeetingRequestsUsingExchangeServer
		End Sub
	End Class
End Namespace